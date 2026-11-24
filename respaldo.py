"""Copia de seguridad automática de la carpeta trabajo (certificaciones, numeración, config, usuarios).

· Arma un .zip en Sistema_web/respaldos/respaldo_AAAA-MM-DD_HHMMSS.zip y guarda los últimos 30.
· La base de usuarios (biomasa.db) se copia con la función de respaldo de SQLite: queda completa y coherente
  aunque el programa esté funcionando en ese momento.
· No copia trabajo/subidos (maestro y planilla del laboratorio: son archivos de la empresa que ya están en otro
  lado y pesan mucho). Sí copia el config vigente, los estados de cada mes, los PDF y Excel generados.
· Destino extra (recomendado): si trabajo/respaldo_destino.txt tiene una carpeta (un disco externo, una carpeta
  compartida de la red, la carpeta de Google Drive/OneDrive de esa PC), cada copia se guarda también ahí.
  Si se rompe el disco de la PC servidor, la copia de afuera es la que salva todo.
· Antes de dar la copia por buena se verifica que el .zip se pueda leer completo.

Uso:
    python respaldo.py            → hace una copia ahora (lo usa COPIA_DE_SEGURIDAD.bat)
    desde servidor.py             → Respaldo(...).iniciar_automatico(): una copia por día, sola
"""
from __future__ import annotations

import json
import os
import shutil
import sqlite3
import sys
import tempfile
import threading
import time
import zipfile
from contextlib import closing
from datetime import datetime
from pathlib import Path

__all__ = ["Respaldo"]

EXCLUIR_CARPETAS = {"subidos"}                      # dentro de trabajo
EXCLUIR_SUFIJOS = (".tmp", "-wal", "-shm", ".lock")
CADA_HORAS = 24
AVISAR_SI_PASAN_HORAS = 48


class Respaldo:
    def __init__(self, trabajo, carpeta_respaldos, conservar: int = 30):
        self.trabajo = Path(trabajo)
        self.carpeta = Path(carpeta_respaldos)
        self.conservar = int(conservar)
        self.registro = self.trabajo / "ultimo_respaldo.json"
        self.config_destino = self.trabajo / "respaldo_destino.txt"
        self._lock = threading.Lock()
        self._parar = threading.Event()
        self._hilo: threading.Thread | None = None

    # ------------------------------------------------------------------------------------------ copia
    def hacer(self) -> dict:
        """Hace una copia ahora. Devuelve lo mismo que queda en ultimo_respaldo.json."""
        with self._lock:
            self.carpeta.mkdir(parents=True, exist_ok=True)
            nombre = f"respaldo_{datetime.now():%Y-%m-%d_%H%M%S}.zip"
            final = self.carpeta / nombre
            tmp = self.carpeta / (nombre + ".tmp")
            omitidos = []
            try:
                with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as z:
                    for ruta in sorted(self.trabajo.rglob("*")):
                        rel = ruta.relative_to(self.trabajo)
                        if not ruta.is_file() or (rel.parts and rel.parts[0] in EXCLUIR_CARPETAS):
                            continue
                        if ruta.name.endswith(EXCLUIR_SUFIJOS) or ruta.name.startswith("~"):
                            continue
                        if ruta.suffix == ".db":
                            self._agregar_sqlite(z, ruta, rel.as_posix(), omitidos)
                            continue
                        try:
                            z.write(ruta, rel.as_posix())
                        except OSError as e:                     # ej. un archivo tomado por otro programa
                            omitidos.append(f"{rel.as_posix()}: {e}")
                    z.writestr("LEEME_RESPALDO.txt", self._leeme(omitidos))
                with zipfile.ZipFile(tmp) as z:                  # verificación: se puede leer entero
                    malo = z.testzip()
                    if malo:
                        raise OSError(f"el archivo {malo} quedó dañado dentro del zip")
                os.replace(tmp, final)
            except Exception:
                try:
                    tmp.unlink(missing_ok=True)
                except OSError:
                    pass
                raise
            self._podar(self.carpeta)
            info = {"fecha": datetime.now().strftime("%d/%m/%Y %H:%M"), "ts": time.time(), "archivo": final.name,
                    "mb": round(final.stat().st_size / 1_048_576, 1), "omitidos": omitidos}
            info.update(self._copiar_afuera(final))
            self._guardar_registro(info)
            return info

    def _agregar_sqlite(self, z: zipfile.ZipFile, ruta: Path, nombre: str, omitidos: list) -> None:
        """Copia coherente de una base SQLite en uso (API de respaldo de SQLite)."""
        fd, tmp = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        try:
            with closing(sqlite3.connect(ruta, timeout=10)) as origen, closing(sqlite3.connect(tmp)) as copia:
                origen.backup(copia)
            z.write(tmp, nombre)
        except (sqlite3.Error, OSError) as e:
            omitidos.append(f"{nombre}: {e}")
        finally:
            try:
                os.remove(tmp)
            except OSError:
                pass

    def _leeme(self, omitidos: list) -> str:
        return ("Copia de seguridad del programa de certificaciones de biomasa\r\n"
                f"Fecha: {datetime.now():%d/%m/%Y %H:%M}\r\n\r\n"
                "Para restaurar: cerrar la ventana negra del servidor, renombrar la carpeta Sistema_web\\trabajo\r\n"
                "(por ejemplo a trabajo_viejo) y descomprimir este zip en una carpeta nueva Sistema_web\\trabajo.\r\n"
                "Después volver a subir el maestro y la planilla del laboratorio (no se copian: están en otro lado).\r\n"
                + (("\r\nNo se pudieron copiar:\r\n" + "\r\n".join(omitidos) + "\r\n") if omitidos else ""))

    def _copiar_afuera(self, zip_hecho: Path) -> dict:
        """Copia también al destino de respaldo_destino.txt (si está configurado)."""
        destino = self.destino_extra()
        if not destino:
            return {"afuera": None}
        try:
            destino.mkdir(parents=True, exist_ok=True)
            tmp = destino / (zip_hecho.name + ".tmp")
            shutil.copyfile(zip_hecho, tmp)
            os.replace(tmp, destino / zip_hecho.name)
            self._podar(destino)
            return {"afuera": str(destino), "afuera_ok": True}
        except OSError as e:
            return {"afuera": str(destino), "afuera_ok": False, "afuera_error": str(e)}

    def destino_extra(self) -> Path | None:
        try:
            for linea in self.config_destino.read_text(encoding="utf-8-sig").splitlines():
                linea = linea.split("#", 1)[0].strip().strip('"')
                if linea:
                    return Path(linea)
        except OSError:
            pass
        return None

    def _podar(self, carpeta: Path) -> None:
        copias = sorted(carpeta.glob("respaldo_*.zip"))
        for vieja in copias[:-self.conservar] if len(copias) > self.conservar else []:
            try:
                vieja.unlink()
            except OSError:
                pass

    # ------------------------------------------------------------------------------------------ registro
    def _guardar_registro(self, info: dict) -> None:
        try:
            tmp = self.registro.with_suffix(".tmp")
            tmp.write_text(json.dumps(info, ensure_ascii=False, indent=1), encoding="utf-8")
            os.replace(tmp, self.registro)
        except OSError:
            pass

    def ultimo(self) -> dict | None:
        try:
            return json.loads(self.registro.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return None

    def estado(self) -> dict:
        """Para la página: cuándo fue la última copia y si hay que avisar algo."""
        u = self.ultimo()
        if not u:
            return {"ultimo": None, "avisar": True, "mensaje": "Todavía no se hizo ninguna copia de seguridad."}
        horas = (time.time() - float(u.get("ts", 0))) / 3600
        msg = None
        if horas > AVISAR_SI_PASAN_HORAS:
            msg = f"La última copia de seguridad es del {u['fecha']} (hace {int(horas // 24)} días)."
        elif u.get("afuera") and u.get("afuera_ok") is False:
            msg = (f"La copia del {u['fecha']} se hizo en esta PC pero NO en {u['afuera']} "
                   f"({u.get('afuera_error', '')}). Revisá que ese disco o carpeta esté disponible.")
        return {"ultimo": u.get("fecha"), "archivo": u.get("archivo"), "mb": u.get("mb"),
                "afuera": u.get("afuera"), "avisar": bool(msg), "mensaje": msg}

    # ------------------------------------------------------------------------------------------ automático
    def iniciar_automatico(self, primera_espera: float = 120.0) -> None:
        """Hilo: a los 2 minutos de arrancar y después cada hora revisa; si pasaron 24 h, hace una copia."""
        if self._hilo and self._hilo.is_alive():
            return

        def bucle():
            if self._parar.wait(primera_espera):
                return
            while True:
                try:
                    u = self.ultimo()
                    if not u or time.time() - float(u.get("ts", 0)) > CADA_HORAS * 3600:
                        info = self.hacer()
                        print(f"💾 Copia de seguridad hecha: respaldos\\{info['archivo']} ({info['mb']} MB)")
                except Exception as e:                           # el hilo no se muere: reintenta en una hora
                    print(f"⚠ No se pudo hacer la copia de seguridad: {e}")
                if self._parar.wait(3600):
                    return

        self._parar.clear()
        self._hilo = threading.Thread(target=bucle, name="copia-de-seguridad", daemon=True)
        self._hilo.start()

    def detener(self) -> None:
        self._parar.set()


if __name__ == "__main__":
    aqui = Path(__file__).resolve().parent
    try:
        info = Respaldo(aqui / "trabajo", aqui / "respaldos").hacer()
    except Exception as e:
        print(f"ERROR: no se pudo hacer la copia de seguridad: {e}")
        sys.exit(1)
    print(f"Copia hecha: respaldos\\{info['archivo']} ({info['mb']} MB)")
    if info.get("afuera"):
        print(f"También en: {info['afuera']}" if info.get("afuera_ok") else
              f"OJO: no se pudo copiar a {info['afuera']}: {info.get('afuera_error')}")
    else:
        print("Consejo: escribí en trabajo\\respaldo_destino.txt una carpeta FUERA de esta PC (disco externo,")
        print("carpeta compartida o Google Drive) para tener también una copia afuera.")
    for o in info.get("omitidos", []):
        print("  no se copió:", o)
