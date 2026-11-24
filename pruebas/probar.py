"""Pruebas automáticas del programa de certificaciones (PROBAR_PROGRAMA.bat).

Se corren en la PC servidor después de instalar o actualizar. Trabajan sobre una COPIA del programa en una carpeta
temporal: no tocan la carpeta trabajo real, ni la numeración, ni los usuarios.

Qué prueban:
  1. Cálculos: descuento por granulometría, monedas.
  2. Lista blanca de IP.
  3. Base de usuarios y sesiones (SQLite).
  4. Config abierto en Excel (en Windows con un bloqueo REAL, igual al de Excel).
  5. Copia de seguridad.
  6. Seguridad: inicio de sesión, bloqueo por intentos, roles, pedidos de otra página, IP no autorizada.
  7. Un mes completo con TUS archivos (los últimos que se subieron): armar, cuadre, confirmar, PDF, Excel,
     numeración sin repetidos y descargas.

Resultado: en pantalla y en pruebas/resultado_AAAA-MM-DD_HHMM.txt.
"""
from __future__ import annotations

import json
import os
import shutil
import sys
import tempfile
import threading
import time
import traceback
import warnings
from datetime import datetime
from pathlib import Path

warnings.filterwarnings("ignore")
PROG = Path(__file__).resolve().parent.parent              # carpeta Sistema_web
TRABAJO_REAL = PROG / "trabajo"
ARCHIVOS_PROGRAMA = ["servidor.py", "cert_biomasa.py", "editor_colab.py", "almacen_db.py", "lista_blanca.py",
                     "guardado_config.py", "respaldo.py", "index.html", "login.html"]

LINEAS: list[str] = []
FALLAS: list[str] = []


def decir(t=""):
    print(t)
    LINEAS.append(t)


def chk(cond, msg):
    decir(("  OK     " if cond else "  FALLA  ") + msg)
    if not cond:
        FALLAS.append(msg)
    return cond


def seccion(t):
    decir("")
    decir(f"== {t}")


def bloque(nombre):
    """Si una sección entera se cae por un error inesperado, se anota y se sigue con las demás."""
    def deco(f):
        def envuelta(*a, **k):
            seccion(nombre)
            try:
                f(*a, **k)
            except Exception as e:
                chk(False, f"{nombre}: error inesperado ({type(e).__name__}: {e})")
                decir(traceback.format_exc())
        return envuelta
    return deco


# ---------------------------------------------------------------------------------------------- 1
@bloque("1. Cálculos")
def p_calculos(cb):
    casos = [(3.6, 10, 15, 0.0), (10.6, 10, 15, 0.6), (12.6, 10, 15, 2.6), (14.3, 10, 15, 4.3), (16.8, 10, 15, 16.8),
             (9.1, 5, 10, 4.1), (11.6, 5, 10, 11.6), (15.0, 10, 15, 5.0)]
    for finos, tol, lim, esperado in casos:
        chk(cb.descuento_finos(finos, tol, lim) == esperado,
            f"finos {finos} % con {tol}/{lim} → descuento {esperado} %")
    chk(cb.moneda_de("USD") == "USD" and cb.moneda_de("u$s") == "USD" and cb.moneda_de("") == "ARS",
        "moneda del config: USD / U$S / vacío")
    chk(cb.dinero(16089.34, "USD") == "USD 16.089,34" and cb.dinero(1234.5) == "$ 1.234,50", "formato de importes")
    chk(cb.cuit_valido("30-00000001-5") and not cb.cuit_valido("30-00000001-6"), "control del dígito del CUIT")


# ---------------------------------------------------------------------------------------------- 2
@bloque("2. Lista blanca de IP")
def p_lista_blanca(tmp):
    import lista_blanca as lb
    ruta = tmp / "ips.txt"
    ruta.write_text("192.168.1.0/24\n10.0.0.0/8\n", encoding="utf-8")
    L = lb.ListaBlanca(ruta)
    for ip, en_red, esperado in [("192.168.1.50", True, True), ("10.1.2.3", True, True), ("8.8.8.8", True, False),
                                 ("192.168.2.1", True, False), ("127.0.0.1", True, True),
                                 ("192.168.1.50", False, False), ("127.0.0.1", False, True),
                                 ("::ffff:192.168.1.9", True, True), ("no-es-ip", True, False)]:
        chk(L.permitida(ip, en_red) == esperado,
            f"{ip} {'en red' if en_red else 'sin red'} → {'entra' if esperado else 'NO entra'}")
    time.sleep(0.05)
    ruta.write_text("192.168.1.77\nbasura\n", encoding="utf-8")
    os.utime(ruta, (time.time() + 2, time.time() + 2))
    chk(L.permitida("192.168.1.77", True) and not L.permitida("192.168.1.50", True),
        "cambiar la lista vale al instante (y la línea mal escrita se ignora)")


# ---------------------------------------------------------------------------------------------- 3
@bloque("3. Base de usuarios y sesiones (SQLite)")
def p_base(tmp):
    import almacen_db as db
    js = tmp / "usuarios.json"
    js.write_text(json.dumps({"admin": {"nombre": "Prueba", "rol": "admin", "activo": True, "sal": "ab",
                                        "hash": "cd", "creado": "01/01/2026"}}), encoding="utf-8")
    db.configurar(tmp / "base.db", json_viejo=js, inactividad=2, duracion_max=5)
    us = db.leer_usuarios()
    chk(list(us) == ["admin"] and (tmp / "usuarios.json.migrado").exists(), "migración desde usuarios.json")
    malo = "x'); drop table usuarios; --"
    us[malo] = {"nombre": "n", "rol": "consulta", "sal": "s", "hash": "h"}
    db.guardar_usuarios(us)
    chk(malo in db.leer_usuarios(), "un nombre con SQL se guarda como texto (sin inyección)")
    try:
        db.guardar_usuarios({**us, "otro": {"nombre": "n", "rol": "jefe", "sal": "s", "hash": "h"}})
        chk(False, "rol inválido rechazado")
    except db.ErrorBase:
        chk(len(db.leer_usuarios()) == 2, "rol inválido rechazado sin dejar nada a medias")
    S = db.Sesiones()
    t = time.time()
    S["tok"] = {"u": "admin", "creada": t, "ultima": t, "ip": "1.2.3.4"}
    chk(S.get("tok")["u"] == "admin", "crear y leer una sesión")
    time.sleep(2.3)
    chk(S.get("tok") is None, "la sesión vence por inactividad")
    errores = []

    def hilo(i):
        try:
            for k in range(10):
                S[f"t{i}-{k}"] = {"u": "admin", "creada": time.time(), "ultima": time.time(), "ip": ""}
                S.get(f"t{i}-{k}")
        except Exception as e:
            errores.append(e)
    hs = [threading.Thread(target=hilo, args=(i,)) for i in range(8)]
    [h.start() for h in hs]
    [h.join() for h in hs]
    chk(not errores and len(S) == 80, "8 usuarios a la vez sin errores")


# ---------------------------------------------------------------------------------------------- 4
def _bloquear_como_excel(ruta):
    """En Windows abre el archivo como lo hace Excel (sin permitir escribir ni borrar). Devuelve una función para
    soltarlo; fuera de Windows devuelve None (se simula)."""
    if os.name != "nt":
        return None
    import ctypes
    from ctypes import wintypes
    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    k32.CreateFileW.restype = wintypes.HANDLE
    k32.CreateFileW.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD, ctypes.c_void_p, wintypes.DWORD,
                                wintypes.DWORD, wintypes.HANDLE]
    k32.CloseHandle.argtypes = [wintypes.HANDLE]
    h = k32.CreateFileW(str(ruta), 0x80000000, 0x00000001, None, 3, 0x80, None)   # GENERIC_READ, FILE_SHARE_READ
    if h in (None, wintypes.HANDLE(-1).value):
        raise OSError(f"no pude abrir {ruta} ({ctypes.get_last_error()})")
    return lambda: k32.CloseHandle(h)


@bloque("4. Config abierto en Excel")
def p_config_bloqueado(cb, tmp):
    import guardado_config as gc
    conf, precios, par = cb.leer_config(str(tmp / "config_prueba.xlsx"))
    ruta = tmp / "cfg" / "config_proveedores.xlsx"
    ruta.parent.mkdir()
    G = gc.GuardadoConfig(ruta, escribir=cb.guardar_config, intervalo=0.5, registrar=lambda t: None)
    chk(G.guardar(conf, precios, par) and ruta.exists(), "guardado normal")
    soltar = _bloquear_como_excel(ruta)
    real = os.replace
    if soltar is None:                                          # Linux/Mac: se simula el bloqueo de Windows
        bloqueado = {"on": True}

        def falso(a, b):
            if bloqueado["on"] and Path(ruta).resolve() in (Path(a).resolve(), Path(b).resolve()):
                raise PermissionError(13, "en uso por otro proceso (simulado)")
            return real(a, b)
        G._reemplazar = falso
        soltar = lambda: bloqueado.update(on=False)
        decir("  (fuera de Windows: el bloqueo de Excel se simula)")
    else:
        decir("  (bloqueo REAL de Windows, igual al de Excel)")
    try:
        par2 = dict(par, NUMERO_INICIAL=int(par["NUMERO_INICIAL"]) + 7)
        quedo = G.guardar(conf, precios, par2)
        chk(quedo is False and G.pendiente.exists() and G.bandera.exists(),
            "con el archivo abierto no falla: queda en config_proveedores_PENDIENTE.xlsx")
        chk(G.estado()["bloqueado"], "la página recibe el aviso de config bloqueado")
        G.iniciar_sincronizacion()
        time.sleep(1.2)
        chk(G.pendiente.exists(), "mientras siga abierto, no se pisa nada")
    finally:
        soltar()
    for _ in range(20):
        if not G.pendiente.exists():
            break
        time.sleep(0.5)
    G.detener()
    chk(not G.pendiente.exists() and not G.bandera.exists(), "al cerrarlo, los cambios pasan solos al config")
    _, _, p3 = cb.leer_config(str(ruta))
    chk(int(p3["NUMERO_INICIAL"]) == int(par["NUMERO_INICIAL"]) + 7, "el config final tiene el cambio")


# ---------------------------------------------------------------------------------------------- 5
@bloque("5. Copia de seguridad")
def p_respaldo(tmp):
    import respaldo
    import zipfile
    t = tmp / "trab_resp"
    (t / "certificaciones_2026-01").mkdir(parents=True)
    (t / "certificaciones_2026-01" / "estado_2026-01.json").write_text('{"proximo": 1}', encoding="utf-8")
    (t / "subidos").mkdir()
    (t / "subidos" / "maestro.xlsb").write_bytes(b"x" * 100)
    import sqlite3
    with sqlite3.connect(t / "biomasa.db") as con:
        con.execute("create table x(a)")
        con.execute("insert into x values (1)")
    afuera = tmp / "afuera"
    (t / "respaldo_destino.txt").write_text(str(afuera), encoding="utf-8")
    R = respaldo.Respaldo(t, tmp / "respaldos")
    info = R.hacer()
    n = zipfile.ZipFile(tmp / "respaldos" / info["archivo"]).namelist()
    chk("biomasa.db" in n and "certificaciones_2026-01/estado_2026-01.json" in n, "la copia tiene la base y los meses")
    chk(not any(x.startswith("subidos/") for x in n), "no copia los archivos subidos (pesan y están en otro lado)")
    chk(info.get("afuera_ok") and (afuera / info["archivo"]).exists(), "también la guarda en la carpeta de afuera")
    chk(not R.estado()["avisar"], "con una copia reciente no hay aviso")


# ---------------------------------------------------------------------------------------------- 6 y 7
def _archivos_reales():
    """Los últimos maestro / planilla del laboratorio / config que se subieron al programa real."""
    halladas = {}
    try:
        ses = json.loads((TRABAJO_REAL / "sesion.json").read_text(encoding="utf-8"))
        for tipo in ("maestro", "calidad"):
            if ses.get(tipo) and Path(ses[tipo]).exists():
                halladas[tipo] = Path(ses[tipo])
    except (OSError, ValueError):
        pass
    for cand in (TRABAJO_REAL / "config_proveedores_PENDIENTE.xlsx", TRABAJO_REAL / "config_proveedores.xlsx",
                 PROG / "config_proveedores.xlsx"):
        if cand.exists():
            halladas["config"] = cand
            break
    return halladas


@bloque("6. Seguridad (inicio de sesión, roles, IP)")
def p_seguridad(sv):
    from starlette.testclient import TestClient
    sv.EN_RED, sv.LOGIN_ACTIVO = True, True
    reg = {"nombre": "Admin Prueba", "rol": "admin", "activo": True, "creado": ""}
    sv._poner_clave(reg, "Clave1234", False)
    con = {"nombre": "Consulta Prueba", "rol": "consulta", "activo": True, "creado": ""}
    sv._poner_clave(con, "Consulta123", False)
    sv._guardar_usuarios({"admin": reg, "juan": con})
    cli = lambda ip="127.0.0.1": TestClient(sv.app_web, base_url="http://localhost", client=(ip, 5000))
    a = cli()
    r = a.post("/api/login", json={"usuario": "admin", "clave": "Clave1234"})
    chk(r.status_code == 200, "inicio de sesión correcto")
    ck = r.headers.get("set-cookie", "")
    chk("HttpOnly" in ck and "samesite=strict" in ck.lower(), "cookie protegida (HttpOnly, SameSite)")
    chk(a.get("/api/yo").json().get("rol") == "admin", "la sesión reconoce al usuario")
    chk(cli().get("/api/archivos").status_code == 401, "sin iniciar sesión no se puede usar")
    m = cli()
    for _ in range(5):
        m.post("/api/login", json={"usuario": "juan", "clave": "mal"})
    r = m.post("/api/login", json={"usuario": "juan", "clave": "Consulta123"})
    chk(r.status_code == 429, "5 contraseñas mal seguidas → bloqueado")
    sv.FALLOS.clear()
    j = cli()
    j.post("/api/login", json={"usuario": "juan", "clave": "Consulta123"})
    chk(j.post("/api/numeracion", json={"proximo": 1}).status_code == 403, "un usuario de consulta no puede cambiar la numeración")
    chk(j.get("/api/usuarios").status_code == 403, "un usuario de consulta no ve los usuarios")
    r = a.post("/api/numeracion", json={"proximo": 1}, headers={"Origin": "https://otro-sitio.com"})
    chk(r.status_code == 403, "pedido desde otra página rechazado aunque haya sesión")
    sv.LISTA_BLANCA.ruta.write_text("192.168.1.0/24\n", encoding="utf-8")
    os.utime(sv.LISTA_BLANCA.ruta, (time.time() + 3, time.time() + 3))
    chk(cli("8.8.8.8").get("/login").status_code == 403, "una IP no autorizada no ve ni el inicio de sesión")
    chk(cli("192.168.1.30").get("/login").status_code == 200, "una PC de la planta sí")
    token = a.cookies.get(sv.COOKIE)
    a.post("/api/logout")
    viejo = cli()
    viejo.cookies.set(sv.COOKIE, token)
    chk(viejo.get("/api/archivos").status_code == 401, "después de Salir la sesión ya no sirve (aunque se guarde la cookie)")
    sv.EN_RED, sv.LOGIN_ACTIVO = False, False


@bloque("7. Un mes completo con tus archivos")
def p_mes(sv, cb, reales):
    from starlette.testclient import TestClient
    if not {"maestro", "calidad", "config"} <= set(reales):
        decir("  (no hay archivos subidos todavía en el programa: se saltea esta parte)")
        return
    decir(f"  maestro: {reales['maestro'].name} · laboratorio: {reales['calidad'].name} · config: {reales['config'].name}")
    c = TestClient(sv.app_web, base_url="http://localhost", client=("127.0.0.1", 5000))
    fs = [("archivos", (p.name if t != "config" else "config_proveedores.xlsx", open(p, "rb"))) for t, p in reales.items()]
    t0 = time.time()
    r = c.post("/api/archivos", files=fs)
    for _, (_, fh) in fs:
        fh.close()
    chk(r.status_code == 200, f"subir los archivos ({time.time() - t0:.0f} s)")
    info = r.json()
    meses = info["maestro"]["meses"]
    m = meses[-2] if len(meses) > 1 else meses[-1]               # el último mes completo
    t0 = time.time()
    g = c.post("/api/armar", json={"anio": m["anio"], "mes": m["mes"], "corte": info["config"]["corte"]}).json()
    chk(len(g["certs"]) > 0, f"armar {m['nombre']}: {len(g['certs'])} certificaciones ({time.time() - t0:.0f} s)")
    chk(all("✔" in str(x["Resultado"]) for x in g["cuadre"]), "cuadre contra el maestro: las toneladas coinciden")
    sin_precio = [x["proveedor"] for x in g["certs"] if not x["importe"]]
    if sin_precio:
        decir(f"  aviso (no es falla del programa): {len(sin_precio)} sin precio cargado: {', '.join(sin_precio[:8])}"
              + (" …" if len(sin_precio) > 8 else ""))
    q = {"anio": m["anio"], "mes": m["mes"]}
    lista = next((x for x in g["certs"] if x["estado"] == "LISTA" and x["importe"]), None) or g["certs"][0]
    d = c.get("/api/certificacion", params=dict(q, key=lista["key"])).json()
    chk(abs(sum(p["subtotal"] or 0 for p in d["precios"] if p["entregado"]) - (d["importe"] or 0)) < 0.05,
        f"{d['razon']}: el importe es la suma de los productos")
    r = c.post("/api/confirmar", params={"key": lista["key"]}, json={"usuario": "prueba", **q})
    j = r.json()
    chk(r.status_code == 200 and j.get("numero"), f"confirmar {d['razon']} → Nº {j.get('numero')}")
    if r.status_code == 200:
        chk(Path(j["pdf"]).exists() and Path(j["pdf"]).stat().st_size > 3000, "PDF generado")
        chk(Path(j["excel"]).exists(), "Excel generado")
        try:
            import subprocess
            txt = subprocess.run(["pdftotext", j["pdf"], "-"], capture_output=True, text=True, timeout=30).stdout
            if txt:
                chk(cb.decimal(j["importe"]) in txt, "el importe del PDF coincide")
        except (OSError, subprocess.SubprocessError):
            pass
    per = lista["periodo"]
    r = c.post("/api/confirmar_todos", json={"usuario": "prueba", "incluir_revisar": True, "periodos": [per], **q}).json()
    nums = [x["numero"] for x in r["certs"] if x["numero"]]
    chk(len(nums) == len(set(nums)) and len(nums) > 0, f"certificar todos ({per}): {len(nums)} números, ninguno repetido")
    seguidos = sorted(int(n.rsplit("-", 1)[1]) for n in nums)
    chk(seguidos == list(range(seguidos[0], seguidos[0] + len(seguidos))), "numeración correlativa sin saltos")
    for tipo in ("pdf", "excel", "zip"):
        x = c.get("/api/descargar_juntas", params={**q, "tipo": tipo, "periodos": per})
        chk(x.status_code == 200 and len(x.content) > 2000, f"descargar todas juntas ({tipo})")


def main():
    decir(f"PRUEBAS DEL PROGRAMA DE CERTIFICACIONES · {datetime.now():%d/%m/%Y %H:%M} · "
          f"{'Windows' if os.name == 'nt' else sys.platform} · Python {sys.version.split()[0]}")
    decir("Se trabaja sobre una copia temporal: tu carpeta trabajo, la numeración y los usuarios no se tocan.")
    reales = _archivos_reales()
    tmp = Path(tempfile.mkdtemp(prefix="prueba_cert_"))
    try:
        for f in ARCHIVOS_PROGRAMA:
            shutil.copy2(PROG / f, tmp / f)
        for carpeta in ("imagenes", "motor"):
            if (PROG / carpeta).exists():
                shutil.copytree(PROG / carpeta, tmp / carpeta, ignore=shutil.ignore_patterns("__pycache__"))
        cfg = reales.get("config") or (PROG / "config_proveedores.xlsx")
        shutil.copy2(cfg, tmp / "config_prueba.xlsx")
        os.chdir(tmp)
        sys.path.insert(0, str(tmp))
        import cert_biomasa as cb
        p_calculos(cb)
        p_lista_blanca(tmp)
        p_base(tmp)
        p_config_bloqueado(cb, tmp)
        p_respaldo(tmp)
        import servidor as sv                                   # el servidor de la COPIA (su propia carpeta trabajo)
        sv.RESPALDO.detener()
        p_seguridad(sv)
        sv.LISTA_BLANCA.ruta.write_text("127.0.0.1\n", encoding="utf-8")
        p_mes(sv, cb, reales)
    finally:
        os.chdir(PROG)
        decir("")
        decir("=" * 70)
        if FALLAS:
            decir(f"RESULTADO: {len(FALLAS)} PRUEBA(S) FALLARON:")
            for f in FALLAS:
                decir(f"  - {f}")
            decir("Mandá el archivo de resultado a soporte.")
        else:
            decir("RESULTADO: TODO BIEN ✔  (todas las pruebas pasaron)")
        salida = PROG / "pruebas" / f"resultado_{datetime.now():%Y-%m-%d_%H%M}.txt"
        try:
            salida.write_text("\n".join(LINEAS), encoding="utf-8")
            print(f"Quedó guardado en pruebas\\{salida.name}")
        except OSError:
            pass
        shutil.rmtree(tmp, ignore_errors=True)
    return 1 if FALLAS else 0


if __name__ == "__main__":
    sys.exit(main())
