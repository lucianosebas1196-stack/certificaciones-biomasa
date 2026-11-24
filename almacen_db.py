"""Usuarios y sesiones del programa guardados en SQLite (trabajo/biomasa.db).

Reemplaza a trabajo/usuarios.json (usuarios) y al diccionario SESIONES en memoria (sesiones abiertas).
Solo usa la librería estándar `sqlite3`.

Cómo se usa desde servidor.py (mismo formato de datos que antes, así el resto no cambia):

    import almacen_db as db
    db.configurar(TRABAJO / "biomasa.db", json_viejo=USUARIOS, inactividad=INACTIVIDAD, duracion_max=DURACION_MAX)
    _leer_usuarios = db.leer_usuarios          # → {"usuario": {"nombre", "rol", "activo", ...}}
    _guardar_usuarios = db.guardar_usuarios    # recibe ese mismo diccionario completo
    SESIONES = db.Sesiones()                   # se usa como el dict de antes: get / pop / [token] = {...} / items

Garantías:
  · Cada escritura es una transacción ACID: se guarda completa o no se guarda nada (nunca queda a medias).
  · Modo WAL: varias PC pueden leer mientras otra escribe, sin bloquearse.
  · Todas las consultas son parametrizadas (con ?): un nombre de usuario raro no puede inyectar SQL.
  · Cada operación abre y cierra su propia conexión (bloques with): sirve igual desde varios hilos.
  · El token de sesión NO se guarda tal cual: se guarda su huella SHA-256. Quien copie la base no puede
    usarla para entrar con la sesión de otro.
  · Las sesiones vencidas (INACTIVIDAD sin uso o DURACION_MAX desde el ingreso) se borran solas.
  · La primera vez, si existe trabajo/usuarios.json, se pasan esos usuarios a la base y el archivo se renombra
    a usuarios.json.migrado (queda como respaldo; no se borra).
"""
from __future__ import annotations

import hashlib
import json
import sqlite3
import threading
import time
from contextlib import closing, contextmanager
from pathlib import Path

__all__ = ["configurar", "leer_usuarios", "guardar_usuarios", "Sesiones", "ErrorBase"]

_RUTA: Path | None = None
_INACTIVIDAD = 60 * 60
_DURACION_MAX = 12 * 60 * 60
_ESPERA_BLOQUEO = 10.0            # segundos que espera si otra escritura tiene la base tomada
_ultima_limpieza = 0.0
_lock_limpieza = threading.Lock()

# Columnas de la tabla usuarios (además de la clave 'usuario'), en el orden del INSERT
_CAMPOS = ("nombre", "rol", "activo", "creado", "ultimo_ingreso", "debe_cambiar", "sal", "hash")

_ESQUEMA = """
CREATE TABLE IF NOT EXISTS usuarios (
    usuario         TEXT PRIMARY KEY,                 -- siempre en minúsculas
    nombre          TEXT NOT NULL,
    rol             TEXT NOT NULL CHECK (rol IN ('admin', 'certificador', 'consulta')),
    activo          INTEGER NOT NULL DEFAULT 1 CHECK (activo IN (0, 1)),
    creado          TEXT NOT NULL DEFAULT '',
    ultimo_ingreso  TEXT NOT NULL DEFAULT '',
    debe_cambiar    INTEGER NOT NULL DEFAULT 0 CHECK (debe_cambiar IN (0, 1)),
    sal             TEXT NOT NULL,
    hash            TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS sesiones (
    token    TEXT PRIMARY KEY,                        -- huella SHA-256 del token de la cookie
    usuario  TEXT NOT NULL REFERENCES usuarios(usuario) ON DELETE CASCADE,
    creada   REAL NOT NULL,                           -- segundos (time.time())
    ultima   REAL NOT NULL,
    ip       TEXT NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS sesiones_por_usuario ON sesiones(usuario);
CREATE INDEX IF NOT EXISTS sesiones_por_ultima ON sesiones(ultima);
"""


class ErrorBase(RuntimeError):
    """Problema con trabajo/biomasa.db (archivo dañado, disco lleno, sin permisos...)."""


# ------------------------------------------------------------------------------------------- conexión
@contextmanager
def _conexion(escritura: bool = False):
    """Abre una conexión, la deja en una transacción y la cierra siempre.
    escritura=True toma el permiso de escritura al empezar (BEGIN IMMEDIATE): dos escrituras a la vez no se
    pisan; la segunda espera hasta _ESPERA_BLOQUEO segundos."""
    if _RUTA is None:
        raise ErrorBase("La base de datos no está configurada: falta llamar a almacen_db.configurar().")
    try:
        with closing(sqlite3.connect(_RUTA, timeout=_ESPERA_BLOQUEO, isolation_level=None)) as con:
            con.row_factory = sqlite3.Row
            con.execute("PRAGMA foreign_keys = ON")
            con.execute("BEGIN IMMEDIATE" if escritura else "BEGIN")
            try:
                yield con
            except BaseException:
                con.execute("ROLLBACK")                # nada a medias: se deshace todo
                raise
            else:
                con.execute("COMMIT")
    except sqlite3.Error as e:
        raise ErrorBase(f"No se pudo usar la base de usuarios ({_RUTA.name}): {e}") from e


def configurar(ruta_db, json_viejo=None, inactividad: int | None = None, duracion_max: int | None = None) -> None:
    """Crea la base y las tablas si no existen, activa WAL y, si hace falta, migra usuarios.json."""
    global _RUTA, _INACTIVIDAD, _DURACION_MAX
    _RUTA = Path(ruta_db)
    _RUTA.parent.mkdir(parents=True, exist_ok=True)
    if inactividad is not None:
        _INACTIVIDAD = int(inactividad)
    if duracion_max is not None:
        _DURACION_MAX = int(duracion_max)
    try:
        with closing(sqlite3.connect(_RUTA, timeout=_ESPERA_BLOQUEO)) as con:
            con.execute("PRAGMA journal_mode = WAL")        # queda guardado en el archivo de la base
            con.execute("PRAGMA synchronous = FULL")         # cada transacción confirmada llega al disco
            con.executescript(_ESQUEMA)
            con.commit()
    except sqlite3.Error as e:
        raise ErrorBase(f"No se pudo crear/abrir {_RUTA}: {e}") from e
    if json_viejo:
        _migrar_json(Path(json_viejo))
    limpiar_sesiones_vencidas()


def _migrar_json(ruta_json: Path) -> None:
    """Pasa los usuarios de usuarios.json a la base (solo si la tabla está vacía) y renombra el json."""
    if not ruta_json.exists():
        return
    with _conexion() as con:
        ya_hay = con.execute("SELECT COUNT(*) FROM usuarios").fetchone()[0]
    if ya_hay:
        return
    try:
        datos = json.loads(ruta_json.read_text(encoding="utf-8"))
    except (OSError, ValueError) as e:
        raise ErrorBase(f"No pude leer {ruta_json.name} para pasarlo a la base: {e}") from e
    if not isinstance(datos, dict):
        raise ErrorBase(f"{ruta_json.name} no tiene el formato esperado.")
    guardar_usuarios(datos)
    try:
        ruta_json.replace(ruta_json.with_name(ruta_json.name + ".migrado"))
    except OSError:
        pass                                            # no es grave: como la tabla ya no está vacía, no se repite


# ------------------------------------------------------------------------------------------- usuarios
def _fila_a_registro(f: sqlite3.Row) -> dict:
    return {"nombre": f["nombre"], "rol": f["rol"], "activo": bool(f["activo"]), "creado": f["creado"],
            "ultimo_ingreso": f["ultimo_ingreso"], "debe_cambiar": bool(f["debe_cambiar"]),
            "sal": f["sal"], "hash": f["hash"]}


def leer_usuarios() -> dict:
    """Todos los usuarios, con el mismo formato que tenía usuarios.json:
    {"usuario": {"nombre", "rol", "activo", "creado", "ultimo_ingreso", "debe_cambiar", "sal", "hash"}}.
    Si la base no se puede leer, devuelve {} (igual que antes con un json ilegible): nadie puede entrar."""
    try:
        with _conexion() as con:
            filas = con.execute(f"SELECT usuario, {', '.join(_CAMPOS)} FROM usuarios ORDER BY usuario").fetchall()
    except ErrorBase:
        return {}
    return {f["usuario"]: _fila_a_registro(f) for f in filas}


def _valores(usuario: str, reg: dict) -> tuple:
    for campo in ("nombre", "rol", "sal", "hash"):
        if not reg.get(campo):
            raise ValueError(f"Al usuario '{usuario}' le falta '{campo}'.")
    return (usuario, str(reg["nombre"]), str(reg["rol"]), int(bool(reg.get("activo", True))),
            str(reg.get("creado") or ""), str(reg.get("ultimo_ingreso") or ""),
            int(bool(reg.get("debe_cambiar", False))), str(reg["sal"]), str(reg["hash"]))


def guardar_usuarios(us: dict) -> None:
    """Deja la tabla igual al diccionario recibido (lo que hacía escribir usuarios.json completo), en UNA
    transacción: agrega los nuevos, actualiza los existentes y borra los que ya no están."""
    if not isinstance(us, dict):
        raise TypeError("guardar_usuarios espera un diccionario {usuario: datos}.")
    filas = [_valores(str(u).strip().casefold(), reg) for u, reg in us.items()]   # valida todo antes de tocar la base
    columnas = ", ".join(("usuario",) + _CAMPOS)
    marcas = ", ".join("?" * (len(_CAMPOS) + 1))
    actualizar = ", ".join(f"{c} = excluded.{c}" for c in _CAMPOS)
    with _conexion(escritura=True) as con:
        con.executemany(f"INSERT INTO usuarios ({columnas}) VALUES ({marcas}) "
                        f"ON CONFLICT(usuario) DO UPDATE SET {actualizar}", filas)
        nombres = [f[0] for f in filas]
        if nombres:
            con.execute(f"DELETE FROM usuarios WHERE usuario NOT IN ({', '.join('?' * len(nombres))})", nombres)
        else:
            con.execute("DELETE FROM usuarios")
    _checkpoint()


def _checkpoint() -> None:
    """Pasa lo escrito al archivo principal biomasa.db: así COPIA_DE_SEGURIDAD.bat copia todo al día."""
    try:
        with closing(sqlite3.connect(_RUTA, timeout=_ESPERA_BLOQUEO)) as con:
            con.execute("PRAGMA wal_checkpoint(PASSIVE)")
    except sqlite3.Error:
        pass                                            # no es crítico: SQLite lo hace solo más adelante


# ------------------------------------------------------------------------------------------- sesiones
def _huella(token: str) -> str:
    return hashlib.sha256(str(token).encode("utf-8")).hexdigest()


def limpiar_sesiones_vencidas(ahora: float | None = None) -> int:
    """Borra las sesiones sin uso hace más de INACTIVIDAD o abiertas hace más de DURACION_MAX.
    Devuelve cuántas borró."""
    global _ultima_limpieza
    ahora = time.time() if ahora is None else ahora
    try:
        with _conexion(escritura=True) as con:
            n = con.execute("DELETE FROM sesiones WHERE ultima < ? OR creada < ?",
                            (ahora - _INACTIVIDAD, ahora - _DURACION_MAX)).rowcount
    except ErrorBase:
        return 0
    _ultima_limpieza = ahora
    return n


def _limpiar_si_toca() -> None:
    """Limpieza automática como mucho una vez por minuto (no en cada pedido)."""
    if time.time() - _ultima_limpieza > 60 and _lock_limpieza.acquire(blocking=False):
        try:
            limpiar_sesiones_vencidas()
        finally:
            _lock_limpieza.release()


class Sesiones:
    """Sesiones abiertas guardadas en la base. Se usa igual que el dict SESIONES de antes:

        SESIONES[token] = {"u": usuario, "creada": t, "ultima": t, "ip": ip}
        ses = SESIONES.get(token)            # → {"u", "creada", "ultima", "ip"} o None (si venció, None)
        SESIONES.pop(token, None)            # cerrar sesión
        for token, ses in SESIONES.items()   # (token = huella; sirve para pop)

    Diferencia: cambiar ses["ultima"] ya no alcanza; para marcar uso se llama a SESIONES.tocar(token).
    Como quedan en la base, las sesiones siguen abiertas aunque se reinicie el servidor."""

    def __setitem__(self, token: str, ses: dict) -> None:
        _limpiar_si_toca()
        with _conexion(escritura=True) as con:
            con.execute("INSERT OR REPLACE INTO sesiones (token, usuario, creada, ultima, ip) VALUES (?, ?, ?, ?, ?)",
                        (_huella(token), str(ses["u"]), float(ses["creada"]), float(ses["ultima"]),
                         str(ses.get("ip", ""))))

    def get(self, token, defecto=None):
        """La sesión de ese token, o `defecto` si no existe o ya venció (en ese caso se borra)."""
        if not token:
            return defecto
        _limpiar_si_toca()
        try:
            with _conexion() as con:
                f = con.execute("SELECT usuario, creada, ultima, ip FROM sesiones WHERE token = ?",
                                (self._clave(token),)).fetchone()
        except ErrorBase:
            return defecto
        if f is None:
            return defecto
        ahora = time.time()
        if ahora - f["ultima"] > _INACTIVIDAD or ahora - f["creada"] > _DURACION_MAX:
            self.pop(token, None)
            return defecto
        return {"u": f["usuario"], "creada": f["creada"], "ultima": f["ultima"], "ip": f["ip"]}

    def __contains__(self, token) -> bool:
        return self.get(token) is not None

    def pop(self, token, defecto=None):
        """Cierra (borra) la sesión y la devuelve; `defecto` si no existía."""
        if not token:
            return defecto
        clave = self._clave(token)
        try:
            with _conexion(escritura=True) as con:
                f = con.execute("SELECT usuario, creada, ultima, ip FROM sesiones WHERE token = ?", (clave,)).fetchone()
                con.execute("DELETE FROM sesiones WHERE token = ?", (clave,))
        except ErrorBase:
            return defecto
        return defecto if f is None else {"u": f["usuario"], "creada": f["creada"], "ultima": f["ultima"], "ip": f["ip"]}

    def tocar(self, token, ahora: float | None = None) -> None:
        """Marca que la sesión se usó ahora (reinicia el contador de inactividad)."""
        try:
            with _conexion(escritura=True) as con:
                con.execute("UPDATE sesiones SET ultima = ? WHERE token = ?",
                            (time.time() if ahora is None else ahora, self._clave(token)))
        except ErrorBase:
            pass                                        # si no se pudo, a lo sumo la sesión vence un poco antes

    def cerrar_de(self, usuario: str) -> int:
        """Cierra todas las sesiones de un usuario (desactivado, cambio de rol o contraseña restablecida)."""
        with _conexion(escritura=True) as con:
            return con.execute("DELETE FROM sesiones WHERE usuario = ?", (str(usuario),)).rowcount

    def items(self):
        """(huella, sesión) de las sesiones vigentes. La huella sirve para pop()."""
        _limpiar_si_toca()
        with _conexion() as con:
            filas = con.execute("SELECT token, usuario, creada, ultima, ip FROM sesiones").fetchall()
        return [(f["token"], {"u": f["usuario"], "creada": f["creada"], "ultima": f["ultima"], "ip": f["ip"]})
                for f in filas]

    def __len__(self) -> int:
        with _conexion() as con:
            return con.execute("SELECT COUNT(*) FROM sesiones").fetchone()[0]

    @staticmethod
    def _clave(token: str) -> str:
        """Acepta el token de la cookie o una huella que ya vino de items()."""
        t = str(token)
        return t if len(t) == 64 and all(c in "0123456789abcdef" for c in t) else _huella(t)
