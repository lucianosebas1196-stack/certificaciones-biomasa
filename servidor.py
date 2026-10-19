# =============================================================================
#  Certificaciones de biomasa — versión web (API + página HTML)
#  Usa el mismo motor que el cuaderno de Colab (cert_biomasa.py y editor_colab.py).
#
#  Cómo se usa (en la PC):
#     pip install fastapi uvicorn python-multipart pandas pyxlsb openpyxl reportlab ipywidgets
#     python servidor.py
#  y abrí en el navegador:  http://localhost:8000
# =============================================================================
import io
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")   # avisos de formato de Excel: no afectan
import json
import os
import secrets
import sys
import threading
import webbrowser
from pathlib import Path
from types import SimpleNamespace

import pandas as pd
import contextvars
import hashlib
import hmac
import time
from datetime import datetime
from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse
from pydantic import BaseModel

AQUI = Path(__file__).resolve().parent
sys.path.insert(0, str(AQUI))
sys.path.insert(0, str(AQUI.parent))


def _revisar_archivos():
    """Si falta un archivo del programa (o quedó con otro nombre al descargarlo) avisa cuál y cómo arreglarlo."""
    faltan = [n for n in ("cert_biomasa.py", "editor_colab.py", "index.html", "login.html") if not (AQUI / n).exists()]
    if not faltan:
        return
    print("\n  FALTAN ARCHIVOS DEL PROGRAMA en esta carpeta:\n   ", AQUI)
    for n in faltan:
        base = n.rsplit(".", 1)[0]
        parecidos = [x.name for x in AQUI.iterdir() if x.name.lower().startswith(base.lower()) and x.name != n]
        print(f"   - {n}" + (f"   (está como: {', '.join(parecidos)} → renombralo a {n})" if parecidos else "   (no está: copialo acá)"))
    print("\n  En el Explorador activá Vista → Mostrar → Extensiones de nombre de archivo para ver el nombre completo.")
    sys.exit(1)


_revisar_archivos()
import cert_biomasa as cb          # noqa: E402
import editor_colab as ed          # noqa: E402

TRABAJO = AQUI / "trabajo"
TRABAJO.mkdir(exist_ok=True)

# Seguridad -------------------------------------------------------------------
# · Sin --red el programa escucha solo en esta PC (127.0.0.1) y no pide usuario.
# · Con --red: cada persona entra con SU usuario y contraseña (guardadas cifradas en trabajo/usuarios.json),
#   con roles, sesión que se cierra sola por inactividad, bloqueo tras varios intentos fallidos, registro de
#   ingresos (trabajo/accesos.log) y conexión cifrada HTTPS (certificado propio en trabajo/certificado).
USUARIOS = TRABAJO / "usuarios.json"
ACCESOS = TRABAJO / "accesos.log"
CARPETA_CERT = TRABAJO / "certificado"
ROLES = {"admin": 3, "certificador": 2, "consulta": 1}
NOMBRE_ROL = {"admin": "Administrador", "certificador": "Certificador", "consulta": "Consulta"}
INACTIVIDAD = 60 * 60            # la sesión se cierra tras 1 hora sin usar el programa
DURACION_MAX = 12 * 60 * 60      # y como máximo a las 12 horas de haber entrado
MAX_FALLOS, BLOQUEO = 5, 15 * 60  # 5 intentos fallidos → 15 minutos bloqueado
ITERACIONES = 240_000
COOKIE = "cert_sesion"
LOGIN_ACTIVO = False             # se activa con --red
USAR_HTTPS = False
SESIONES = {}                    # token → {"u", "creada", "ultima", "ip"}
FALLOS = {}                      # (usuario, ip) → [momentos de los intentos fallidos]
_USUARIO_ACTUAL = contextvars.ContextVar("usuario_actual", default=None)
_LOCK_USUARIOS = threading.RLock()
CLAVE_MIN = 8


def _hash(clave, sal):
    return hashlib.pbkdf2_hmac("sha256", clave.encode("utf-8"), bytes.fromhex(sal), ITERACIONES).hex()


def _leer_usuarios():
    try:
        return json.loads(USUARIOS.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _guardar_usuarios(us):
    tmp = USUARIOS.with_suffix(".tmp")
    tmp.write_text(json.dumps(us, ensure_ascii=False, indent=1), encoding="utf-8")
    os.replace(tmp, USUARIOS)


def _validar_clave(clave, usuario=""):
    if len(clave or "") < CLAVE_MIN:
        return f"La contraseña tiene que tener {CLAVE_MIN} caracteres o más."
    if usuario and clave.strip().casefold() == usuario.strip().casefold():
        return "La contraseña no puede ser igual al usuario."
    if not (any(ch.isalpha() for ch in clave) and any(ch.isdigit() for ch in clave)):
        return "La contraseña tiene que tener letras y números."
    return None


def _poner_clave(registro, clave, debe_cambiar):
    registro["sal"] = secrets.token_hex(16)
    registro["hash"] = _hash(clave, registro["sal"])
    registro["debe_cambiar"] = bool(debe_cambiar)


def _registrar_acceso(usuario, ip, resultado):
    try:
        with open(ACCESOS, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now():%d/%m/%Y %H:%M:%S}\t{usuario}\t{ip}\t{resultado}\n")
    except OSError:
        pass


def _usuario_actual():
    """El usuario que inició sesión en este pedido (None si el programa corre sin inicio de sesión)."""
    return _USUARIO_ACTUAL.get()


def _quien(nombre_enviado):
    """Nombre que queda como responsable: el del usuario que inició sesión (no se puede firmar por otro)."""
    u = _usuario_actual()
    return u["nombre"] if u else (nombre_enviado or "web")


def _rol_necesario(metodo, ruta):
    if ruta.startswith("/api/usuarios") or ruta == "/api/numeracion" or (ruta.startswith("/api/imagen/") and metodo != "GET") \
            or (ruta == "/api/tema" and metodo != "GET"):
        return "admin"
    if (metodo in ("GET", "HEAD") and ruta != "/api/descargar_config") or ruta in ("/api/armar", "/api/logout", "/api/clave"):
        return "consulta"           # armar = solo ver (un usuario de consulta no puede cambiar el corte)
    return "certificador"


PUBLICAS = {"/login", "/api/login", "/api/logout", "/api/tema", "/certificado", "/certificado.crt", "/instalar.bat"}
SIN_CAMBIO_CLAVE = {"/", "/api/yo", "/api/clave", "/api/logout", "/api/tema", "/api/imagenes"}


def _sesion_de(request):
    token = request.cookies.get(COOKIE)
    ses = SESIONES.get(token) if token else None
    if not ses:
        return None, None
    ahora = time.time()
    if ahora - ses["ultima"] > INACTIVIDAD or ahora - ses["creada"] > DURACION_MAX:
        SESIONES.pop(token, None)
        return None, None
    us = _leer_usuarios().get(ses["u"])
    if not us or not us.get("activo", True):
        SESIONES.pop(token, None)
        return None, None
    ses["ultima"] = ahora
    return token, dict(us, usuario=ses["u"])


app_web = FastAPI(title="Certificaciones de biomasa", version="1.2",
                  docs_url=None, redoc_url=None, openapi_url=None)
# Un solo cambio a la vez: evita, por ejemplo, que un doble clic en "Confirmar" use dos números correlativos
CANDADO = threading.RLock()
EXTENSIONES = {".xlsb", ".xlsx", ".xlsm", ".json"}
MAX_MB = 200

EN_RED = False          # se activa con --red


async def _silenciar_cortes():
    """En Windows, cuando un navegador corta la conexión (cerró la pestaña, o todavía no confía en el certificado)
    Python muestra 'ConnectionResetError [WinError 10054]'. No es un error del programa: se ignora sin llenar la
    ventana de mensajes."""
    import asyncio
    loop = asyncio.get_running_loop()
    anterior = loop.get_exception_handler()

    def manejador(lp, ctx):
        exc = ctx.get("exception")
        if isinstance(exc, (ConnectionResetError, ConnectionAbortedError, BrokenPipeError)):
            return
        (anterior or (lambda l, c: l.default_exception_handler(c)))(lp, ctx)
    loop.set_exception_handler(manejador)


from contextlib import asynccontextmanager      # noqa: E402


@asynccontextmanager
async def _ciclo_de_vida(app):
    """Lo que se hace al arrancar el servidor (forma actual de FastAPI; reemplaza a on_event("startup"))."""
    await _silenciar_cortes()
    yield


app_web.router.lifespan_context = _ciclo_de_vida


@app_web.middleware("http")
async def _seguridad(request, call_next):
    """1) Sin --red solo atiende a 'localhost'. 2) Los cambios solo se aceptan desde la propia página (otra página
    web abierta en el navegador no puede mandar nada). 3) Con --red: hay que iniciar sesión y cada rol puede
    hacer solo lo suyo. 4) Encabezados de seguridad en todas las respuestas."""
    from urllib.parse import urlparse
    host = request.headers.get("host", "")
    if not EN_RED and host.rsplit(":", 1)[0] not in ("localhost", "127.0.0.1", "[::1]", "testserver"):
        return JSONResponse({"detail": "Acceso no permitido."}, status_code=403)
    if request.method not in ("GET", "HEAD", "OPTIONS"):
        origen = request.headers.get("origin") or request.headers.get("referer")
        if origen and origen != "null" and urlparse(origen).netloc != host:
            return JSONResponse({"detail": "Pedido rechazado: tiene que salir de la página del programa."}, status_code=403)
    ruta = request.url.path
    _USUARIO_ACTUAL.set(None)
    publica = ruta in PUBLICAS and (request.method in ("GET", "HEAD") or ruta in ("/api/login", "/api/logout"))
    if LOGIN_ACTIVO and not publica and not (ruta in ("/api/imagen/logo", "/api/imagen/fondo")
                                             and request.method == "GET"):
        _, us = _sesion_de(request)
        if us is None:
            if ruta.startswith("/api/"):
                return JSONResponse({"detail": "La sesión se cerró: volvé a iniciar sesión.", "login": True},
                                    status_code=401)
            return RedirectResponse("/login", status_code=303)
        if us.get("debe_cambiar") and ruta not in SIN_CAMBIO_CLAVE and not ruta.startswith("/api/imagen"):
            return JSONResponse({"detail": "Antes de seguir tenés que cambiar tu contraseña."}, status_code=403)
        if ROLES.get(us.get("rol"), 0) < ROLES[_rol_necesario(request.method, ruta)]:
            return JSONResponse({"detail": f"Tu usuario ({NOMBRE_ROL.get(us.get('rol'), us.get('rol'))}) "
                                           "no tiene permiso para hacer esto."}, status_code=403)
        _USUARIO_ACTUAL.set(us)
    r = await call_next(request)
    r.headers["X-Content-Type-Options"] = "nosniff"
    r.headers["X-Frame-Options"] = "DENY"
    r.headers["Referrer-Policy"] = "same-origin"
    if ruta.startswith("/api/") or ruta in ("/", "/login"):
        r.headers["Cache-Control"] = "no-store"
    if USAR_HTTPS:
        r.headers["Strict-Transport-Security"] = "max-age=31536000"
    return r


# ----------------------------------------------------------------- inicio de sesión y usuarios
class Login(BaseModel):
    usuario: str
    clave: str


@app_web.get("/login", response_class=HTMLResponse)
def pagina_login():
    return (AQUI / "login.html").read_text(encoding="utf-8")


@app_web.post("/api/login")
async def login(datos: Login, request: Request):
    import anyio
    ip = request.client.host if request.client else "?"
    u = datos.usuario.strip().casefold()
    clave_f = (u, ip)
    ahora = time.time()
    FALLOS[clave_f] = [t for t in FALLOS.get(clave_f, []) if ahora - t < BLOQUEO]
    if len(FALLOS[clave_f]) >= MAX_FALLOS:
        espera = int((BLOQUEO - (ahora - FALLOS[clave_f][0])) // 60) + 1
        _registrar_acceso(u, ip, "BLOQUEADO")
        raise HTTPException(429, f"Demasiados intentos fallidos. Probá de nuevo en {espera} minutos "
                                 "o pedile al administrador que te restablezca la contraseña.")
    us = _leer_usuarios().get(u)
    ok = False
    if us and us.get("activo", True):
        h = await anyio.to_thread.run_sync(_hash, datos.clave, us["sal"])
        ok = hmac.compare_digest(h, us["hash"])
    else:
        await anyio.to_thread.run_sync(_hash, datos.clave, "00" * 16)      # mismo tiempo exista o no el usuario
    if not ok:
        FALLOS[clave_f].append(ahora)
        _registrar_acceso(u, ip, "CONTRASEÑA INCORRECTA" if us else "USUARIO INEXISTENTE")
        await anyio.sleep(0.6)
        quedan = MAX_FALLOS - len(FALLOS[clave_f])
        raise HTTPException(401, "Usuario o contraseña incorrectos." +
                            (f" Te quedan {quedan} intentos." if 0 < quedan <= 2 else ""))
    FALLOS.pop(clave_f, None)
    token = secrets.token_urlsafe(32)
    SESIONES[token] = {"u": u, "creada": ahora, "ultima": ahora, "ip": ip}
    with _LOCK_USUARIOS:
        todos = _leer_usuarios()
        if u in todos:
            todos[u]["ultimo_ingreso"] = f"{datetime.now():%d/%m/%Y %H:%M}"
            _guardar_usuarios(todos)
    _registrar_acceso(u, ip, "INGRESO")
    r = JSONResponse({"ok": True, "debe_cambiar": bool(us.get("debe_cambiar"))})
    r.set_cookie(COOKIE, token, httponly=True, samesite="strict", secure=USAR_HTTPS, path="/", max_age=DURACION_MAX)
    return r


@app_web.post("/api/logout")
def logout(request: Request):
    token = request.cookies.get(COOKIE)
    ses = SESIONES.pop(token, None) if token else None
    if ses:
        _registrar_acceso(ses["u"], request.client.host if request.client else "?", "SALIDA")
    r = JSONResponse({"ok": True})
    r.delete_cookie(COOKIE, path="/")
    return r


@app_web.get("/api/yo")
def yo():
    u = _usuario_actual()
    if not u:
        return {"login": False}
    return {"login": True, "usuario": u["usuario"], "nombre": u["nombre"], "rol": u["rol"],
            "rol_nombre": NOMBRE_ROL.get(u["rol"], u["rol"]), "debe_cambiar": bool(u.get("debe_cambiar")),
            "inactividad_min": INACTIVIDAD // 60}


class CambioClave(BaseModel):
    actual: str
    nueva: str


@app_web.post("/api/clave")
def cambiar_clave(c: CambioClave):
    u = _usuario_actual()
    if not u:
        raise HTTPException(400, "El programa no está usando inicio de sesión.")
    if not hmac.compare_digest(_hash(c.actual, u["sal"]), u["hash"]):
        raise HTTPException(400, "La contraseña actual no es correcta.")
    err = _validar_clave(c.nueva, u["usuario"])
    if err:
        raise HTTPException(400, err)
    if c.nueva == c.actual:
        raise HTTPException(400, "La nueva contraseña tiene que ser distinta de la actual.")
    with _LOCK_USUARIOS:
        todos = _leer_usuarios()
        _poner_clave(todos[u["usuario"]], c.nueva, False)
        _guardar_usuarios(todos)
    _registrar_acceso(u["usuario"], "-", "CAMBIO DE CONTRASEÑA")
    return {"ok": True}


class NuevoUsuario(BaseModel):
    usuario: str
    nombre: str
    rol: str
    clave: str


class CambioUsuario(BaseModel):
    nombre: str | None = None
    rol: str | None = None
    activo: bool | None = None
    clave: str | None = None          # contraseña temporal (la persona la cambia al entrar)


def _lista_usuarios():
    return [{"usuario": k, "nombre": v["nombre"], "rol": v["rol"], "rol_nombre": NOMBRE_ROL.get(v["rol"], v["rol"]),
             "activo": v.get("activo", True), "ultimo_ingreso": v.get("ultimo_ingreso", ""),
             "debe_cambiar": bool(v.get("debe_cambiar"))}
            for k, v in sorted(_leer_usuarios().items())]


@app_web.get("/api/usuarios")
def usuarios():
    return _lista_usuarios()


@app_web.post("/api/usuarios")
def crear_usuario(n: NuevoUsuario):
    import re as _re
    u = n.usuario.strip().casefold()
    if not _re.fullmatch(r"[a-z0-9._-]{3,30}", u):
        raise HTTPException(400, "Usuario: de 3 a 30 letras, números, punto, guion o guion bajo (sin espacios ni acentos).")
    if n.rol not in ROLES:
        raise HTTPException(400, "Rol no válido.")
    if not n.nombre.strip():
        raise HTTPException(400, "Falta el nombre y apellido.")
    err = _validar_clave(n.clave, u)
    if err:
        raise HTTPException(400, err)
    with _LOCK_USUARIOS:
        todos = _leer_usuarios()
        if u in todos:
            raise HTTPException(400, f"El usuario '{u}' ya existe.")
        reg = {"nombre": n.nombre.strip(), "rol": n.rol, "activo": True, "creado": f"{datetime.now():%d/%m/%Y}"}
        _poner_clave(reg, n.clave, True)
        todos[u] = reg
        _guardar_usuarios(todos)
    _registrar_acceso(u, "-", f"USUARIO CREADO por {_usuario_actual()['usuario']}")
    return _lista_usuarios()


@app_web.post("/api/usuarios/{usuario}")
def modificar_usuario(usuario: str, c: CambioUsuario):
    u = usuario.strip().casefold()
    yo_ = _usuario_actual()
    with _LOCK_USUARIOS:
        todos = _leer_usuarios()
        if u not in todos:
            raise HTTPException(404, "No existe ese usuario.")
        reg = dict(todos[u])
        if c.nombre is not None and c.nombre.strip():
            reg["nombre"] = c.nombre.strip()
        if c.rol is not None:
            if c.rol not in ROLES:
                raise HTTPException(400, "Rol no válido.")
            reg["rol"] = c.rol
        if c.activo is not None:
            reg["activo"] = bool(c.activo)
        if c.clave:
            err = _validar_clave(c.clave, u)
            if err:
                raise HTTPException(400, err)
            _poner_clave(reg, c.clave, True)
            for k in [k for k in FALLOS if k[0] == u]:      # contraseña restablecida: se levanta el bloqueo
                FALLOS.pop(k, None)
        admins = [k for k, v in {**todos, u: reg}.items() if v["rol"] == "admin" and v.get("activo", True)]
        if not admins:
            raise HTTPException(400, "Tiene que quedar al menos un Administrador activo.")
        todos[u] = reg
        _guardar_usuarios(todos)
    if (c.activo is False or c.clave or c.rol is not None) and not (yo_ and u == yo_["usuario"]):
        for t in [t for t, s_ in SESIONES.items() if s_["u"] == u]:      # se le cierran las sesiones abiertas
            SESIONES.pop(t, None)
    _registrar_acceso(u, "-", f"USUARIO MODIFICADO por {yo_['usuario'] if yo_ else '?'}: " +
                      ", ".join(k for k, v in c.model_dump().items() if v is not None and k != "clave") +
                      (" + contraseña restablecida" if c.clave else ""))
    return _lista_usuarios()


def _crear_admin_consola(forzar=False):
    """Primera vez en red (o RESTABLECER_ADMIN.bat): crea un Administrador desde esta misma PC."""
    import getpass
    todos = _leer_usuarios()
    if todos and not forzar:
        return
    print("\n=== " + ("CREAR / RESTABLECER UN ADMINISTRADOR" if forzar else
                     "PRIMERA VEZ: creá el usuario Administrador del programa") + " ===")
    print("    El Administrador después crea los usuarios de los demás desde el programa (botón 👥 Usuarios).")
    try:
        while True:
            u = (input("Usuario (ej. luciano): ").strip() or "admin").casefold()
            if __import__("re").fullmatch(r"[a-z0-9._-]{3,30}", u):
                break
            print("  De 3 a 30 letras o números, sin espacios ni acentos.")
        nombre = input("Nombre y apellido (sale como responsable en las certificaciones): ").strip() or u
        while True:
            c1 = getpass.getpass(f"Contraseña ({CLAVE_MIN}+ caracteres, letras y números; no se ve al escribir): ")
            err = _validar_clave(c1, u)
            if err:
                print("  " + err)
                continue
            if getpass.getpass("Repetila: ") != c1:
                print("  No coinciden, de nuevo.")
                continue
            break
    except EOFError:
        sys.exit("Falta crear el Administrador: abrí INICIAR_WEB_RED.bat a mano una vez.")
    with _LOCK_USUARIOS:
        todos = _leer_usuarios()
        reg = todos.get(u, {"creado": f"{datetime.now():%d/%m/%Y}"})
        reg.update({"nombre": nombre, "rol": "admin", "activo": True})
        _poner_clave(reg, c1, False)
        todos[u] = reg
        _guardar_usuarios(todos)
    _registrar_acceso(u, "consola", "ADMINISTRADOR CREADO/RESTABLECIDO EN LA PC SERVIDOR")
    print(f"Listo: '{u}' es Administrador.\n")


# ----------------------------------------------------------------- acceso directo para otras PC
PUERTO = 8000
CARPETA_KIT = AQUI / "PARA_LAS_OTRAS_PC"


def _icono_ico():
    """Ícono del acceso directo: el logo del programa (si hay) sobre fondo blanco cuadrado, o una 'C' del color del tema."""
    from PIL import Image, ImageDraw
    r = _ruta_imagen("logo")
    if r:
        im = Image.open(r).convert("RGBA")
        lado = max(im.size)
        fondo = Image.new("RGBA", (lado, lado), (255, 255, 255, 255))
        fondo.paste(im, ((lado - im.width) // 2, (lado - im.height) // 2), im)
        im = fondo
    else:
        color = tema().get("color", COLOR_BASE)
        im = Image.new("RGBA", (256, 256), (0, 0, 0, 0))
        d = ImageDraw.Draw(im)
        d.rounded_rectangle((8, 8, 248, 248), 48, fill=color)
        d.ellipse((60, 60, 196, 196), outline="white", width=30)
        d.rectangle((150, 100, 210, 156), fill=color)
    b = io.BytesIO()
    im.resize((256, 256)).save(b, format="ICO", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    return b.getvalue()


def _kit_acceso(host_pedido=None):
    """INSTALAR_ACCESO_CERTIFICACIONES.bat: en cada PC, con doble clic crea el acceso directo en el escritorio y
    (con HTTPS) instala el certificado de la empresa para que el navegador muestre el candado. No necesita ser
    administrador ni instalar programas."""
    import base64
    import socket
    esquema = "https" if USAR_HTTPS else "http"
    nombre = socket.gethostname()
    ips = _ips_locales()
    ip = host_pedido if host_pedido and host_pedido.replace(".", "").isdigit() else (ips[0] if ips else nombre)

    def bloque(datos, archivo):
        txt = base64.b64encode(datos).decode()
        lineas = [txt[i:i + 76] for i in range(0, len(txt), 76)]
        if len(lineas) > 1 and lineas[-1].upper() in ("ON", "OFF"):      # "echo ON/OFF" no imprime: se une
            lineas[-2:] = [lineas[-2] + lineas[-1]]
        return "(\r\n" + "".join(f"echo {l}\r\n" for l in lineas) + f')>"%DEST%\\{archivo}.b64"\r\n' \
               f'certutil -f -decode "%DEST%\\{archivo}.b64" "%DEST%\\{archivo}" >nul\r\ndel "%DEST%\\{archivo}.b64"\r\n'
    cert = (CARPETA_CERT / "autoridad.crt").read_bytes() if USAR_HTTPS and (CARPETA_CERT / "autoridad.crt").exists() else None
    b = ["@echo off", "chcp 65001 >nul", "title Acceso directo - Certificaciones de biomasa", "echo.",
         "echo  Acceso directo al programa de Certificaciones de biomasa", "echo  ---------------------------------------------",
         f"set NOMBRE={nombre}", f"set IP={ip}", f"set PUERTO={PUERTO}", f"set ESQUEMA={esquema}",
         'set DEST=%LOCALAPPDATA%\\CertificacionesBiomasa', 'if not exist "%DEST%" mkdir "%DEST%"',
         "rem si esta PC llega al servidor por su nombre se usa el nombre (no cambia aunque cambie la IP); si no, la IP",
         "set HOST=%NOMBRE%",
         "powershell -NoProfile -Command \"try{$c=New-Object Net.Sockets.TcpClient;if($c.ConnectAsync('%NOMBRE%',%PUERTO%).Wait(2500)){exit 0}else{exit 1}}catch{exit 1}\" >nul 2>nul || set HOST=%IP%",
         "echo  Servidor: %ESQUEMA%://%HOST%:%PUERTO%"]
    s_ = "\r\n".join(b) + "\r\n" + bloque(_icono_ico(), "icono.ico")
    if cert:
        s_ += bloque(cert, "certificado.crt")
        s_ += "\r\n".join([
            "echo.", "echo  Instalando el certificado de la empresa (para el candado del navegador).",
            "echo  Si Windows pregunta \"Desea instalar este certificado?\", elegi SI.",
            'certutil -user -addstore Root "%DEST%\\certificado.crt" >nul',
            "if errorlevel 1 (echo  No se instalo el certificado: el navegador va a avisar \"conexion no privada\".) else (echo  Certificado instalado.)",
            ""])
    s_ += "\r\n".join([
        "for /f \"usebackq delims=\" %%d in (`powershell -NoProfile -Command \"[Environment]::GetFolderPath('Desktop')\"`) do set ESC=%%d",
        'if not defined ESC set ESC=%USERPROFILE%\\Desktop',
        '(', "echo [InternetShortcut]", "echo URL=%ESQUEMA%://%HOST%:%PUERTO%/", "echo IconFile=%DEST%\\icono.ico", "echo IconIndex=0",
        ')>"%ESC%\\Certificaciones de biomasa.url"',
        "echo.", "echo  Listo: en el escritorio quedo el acceso \"Certificaciones de biomasa\".",
        "echo  Entra con tu usuario y contrasena (te los da el administrador).",
        'start "" "%ESQUEMA%://%HOST%:%PUERTO%/"', "echo.", "echo  Apreta una tecla para cerrar.", "pause >nul", ""])
    return s_.encode("utf-8")


@app_web.get("/instalar.bat")
def bajar_kit(request: Request):
    host = request.headers.get("host", "").rsplit(":", 1)[0]
    from fastapi.responses import Response
    return Response(_kit_acceso(host), media_type="application/octet-stream",
                        headers={"Content-Disposition": 'attachment; filename="INSTALAR_ACCESO_CERTIFICACIONES.bat"'})


def _guardar_kit():
    try:
        CARPETA_KIT.mkdir(exist_ok=True)
        (CARPETA_KIT / "INSTALAR_ACCESO_CERTIFICACIONES.bat").write_bytes(_kit_acceso())
        (CARPETA_KIT / "LEEME.txt").write_text(
            "Copiá INSTALAR_ACCESO_CERTIFICACIONES.bat a cada PC que vaya a usar el programa (pendrive, mail o carpeta\r\n"
            "compartida) y hacé doble clic. Crea el acceso directo en el escritorio y, si el programa usa HTTPS,\r\n"
            "instala el certificado de la empresa (Windows pregunta: elegí Sí). No hace falta ser administrador.\r\n"
            "También se puede bajar desde la pantalla de inicio de sesión: 'Crear acceso directo en esta PC'.\r\n",
            encoding="utf-8")
    except Exception as e:                      # no es crítico
        print("No pude preparar la carpeta PARA_LAS_OTRAS_PC:", e)


# ----------------------------------------------------------------- HTTPS (certificado propio de la empresa)
def _ips_locales():
    import socket
    try:
        return sorted({ip for ip in socket.gethostbyname_ex(socket.gethostname())[2] if not ip.startswith("127.")})
    except OSError:
        return []


def _preparar_certificado():
    """Crea (una sola vez) una autoridad certificante propia y, con ella, el certificado del servidor con el
    nombre de la PC y sus IP. Si la IP cambia, el certificado del servidor se rehace solo (la autoridad no cambia,
    así las PC que ya la instalaron siguen viendo el candado)."""
    import datetime as dt
    import ipaddress
    import socket
    from cryptography import x509
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import ec
    from cryptography.x509.oid import ExtendedKeyUsageOID, NameOID
    CARPETA_CERT.mkdir(exist_ok=True)
    ca_crt, ca_key = CARPETA_CERT / "autoridad.crt", CARPETA_CERT / "autoridad.key"
    srv_crt, srv_key = CARPETA_CERT / "servidor.crt", CARPETA_CERT / "servidor.key"
    ahora = dt.datetime.now(dt.timezone.utc)
    guardar_clave = lambda k, ruta: ruta.write_bytes(k.private_bytes(serialization.Encoding.PEM,
                                                                     serialization.PrivateFormat.PKCS8,
                                                                     serialization.NoEncryption()))
    if not (ca_crt.exists() and ca_key.exists()):
        k = ec.generate_private_key(ec.SECP256R1())
        nombre = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, f"Certificaciones biomasa - {socket.gethostname()}"),
                            x509.NameAttribute(NameOID.ORGANIZATION_NAME, "Certificaciones de biomasa (red interna)")])
        ca = (x509.CertificateBuilder().subject_name(nombre).issuer_name(nombre).public_key(k.public_key())
              .serial_number(x509.random_serial_number()).not_valid_before(ahora - dt.timedelta(days=1))
              .not_valid_after(ahora + dt.timedelta(days=3650))
              .add_extension(x509.BasicConstraints(ca=True, path_length=0), critical=True)
              .add_extension(x509.SubjectKeyIdentifier.from_public_key(k.public_key()), critical=False)
              .add_extension(x509.AuthorityKeyIdentifier.from_issuer_public_key(k.public_key()), critical=False)
              .add_extension(x509.KeyUsage(digital_signature=True, key_cert_sign=True, crl_sign=True,
                                           content_commitment=False, key_encipherment=False, data_encipherment=False,
                                           key_agreement=False, encipher_only=False, decipher_only=False), critical=True)
              .sign(k, hashes.SHA256()))
        guardar_clave(k, ca_key)
        ca_crt.write_bytes(ca.public_bytes(serialization.Encoding.PEM))
        if srv_crt.exists():
            srv_crt.unlink()
    ca = x509.load_pem_x509_certificate(ca_crt.read_bytes())
    ca_k = serialization.load_pem_private_key(ca_key.read_bytes(), None)
    host = socket.gethostname()
    dns = sorted({host, host.lower(), "localhost"})
    ips = sorted(set(_ips_locales()) | {"127.0.0.1"})
    rehacer = not (srv_crt.exists() and srv_key.exists())
    if not rehacer:
        actual = x509.load_pem_x509_certificate(srv_crt.read_bytes())
        san = actual.extensions.get_extension_for_class(x509.SubjectAlternativeName).value
        tiene_ips = {str(x) for x in san.get_values_for_type(x509.IPAddress)}
        tiene_dns = set(san.get_values_for_type(x509.DNSName))
        vence = actual.not_valid_after_utc
        try:
            actual.extensions.get_extension_for_class(x509.AuthorityKeyIdentifier)
            sin_aki = False
        except x509.ExtensionNotFound:
            sin_aki = True
        rehacer = sin_aki or (not set(ips) <= tiene_ips) or (not set(dns) <= tiene_dns) \
            or vence - ahora < dt.timedelta(days=30)
    if rehacer:
        k = ec.generate_private_key(ec.SECP256R1())
        alt = [x509.DNSName(d) for d in dns] + [x509.IPAddress(ipaddress.ip_address(i)) for i in ips]
        crt = (x509.CertificateBuilder()
               .subject_name(x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, host)]))
               .issuer_name(ca.subject).public_key(k.public_key()).serial_number(x509.random_serial_number())
               .not_valid_before(ahora - dt.timedelta(days=1)).not_valid_after(ahora + dt.timedelta(days=397))
               .add_extension(x509.SubjectAlternativeName(alt), critical=False)
               .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
               .add_extension(x509.ExtendedKeyUsage([ExtendedKeyUsageOID.SERVER_AUTH]), critical=False)
               .add_extension(x509.KeyUsage(digital_signature=True, key_encipherment=False, content_commitment=False,
                                            data_encipherment=False, key_agreement=False, key_cert_sign=False,
                                            crl_sign=False, encipher_only=False, decipher_only=False), critical=True)
               .add_extension(x509.SubjectKeyIdentifier.from_public_key(k.public_key()), critical=False)
               .add_extension(x509.AuthorityKeyIdentifier.from_issuer_public_key(ca_k.public_key()), critical=False)
               .sign(ca_k, hashes.SHA256()))
        guardar_clave(k, srv_key)
        srv_crt.write_bytes(crt.public_bytes(serialization.Encoding.PEM))
    return srv_crt, srv_key


@app_web.get("/certificado.crt")
def bajar_certificado():
    r = CARPETA_CERT / "autoridad.crt"
    if not r.exists():
        raise HTTPException(404, "El programa no está usando HTTPS.")
    return FileResponse(r, filename="Certificaciones-biomasa.crt", media_type="application/x-x509-ca-cert")


@app_web.get("/certificado", response_class=HTMLResponse)
def ayuda_certificado():
    return """<!doctype html><html lang=es><meta charset=utf-8><meta name=viewport content="width=device-width,initial-scale=1">
<title>Instalar certificado</title><style>body{font-family:Segoe UI,system-ui,sans-serif;max-width:720px;margin:40px auto;padding:0 16px;color:#1F2A24;line-height:1.5}
a.b{display:inline-block;background:#3A7D44;color:#fff;padding:10px 16px;border-radius:8px;text-decoration:none;font-weight:600}
code{background:#eef2ef;padding:2px 6px;border-radius:4px}li{margin:6px 0}</style>
<h2>Candado 🔒 en el navegador (una vez por PC)</h2>
<p>El programa usa conexión cifrada con un certificado propio de la empresa. Para que el navegador muestre el candado y no avise
"La conexión no es privada", instalá este certificado en esta PC:</p>
<p><a class=b href="/certificado.crt">⬇ Descargar certificado</a></p>
<ol><li>Abrí el archivo descargado <code>Certificaciones-biomasa.crt</code> → <b>Instalar certificado…</b></li>
<li>Elegí <b>Usuario actual</b> → Siguiente.</li>
<li>Marcá <b>Colocar todos los certificados en el siguiente almacén</b> → Examinar → <b>Entidades de certificación raíz de confianza</b> → Aceptar → Siguiente → Finalizar → <b>Sí</b>.</li>
<li>Cerrá y volvé a abrir Chrome / Edge y entrá de nuevo al programa.</li></ol>
<p><a href="/login">← Volver al inicio de sesión</a></p>"""


class Sesion:
    """Lo que se subió y la certificación en curso (la app es de un solo usuario, corre en la PC)."""
    maestro = None
    humedad = None
    conf = precios = par = None
    logo = None
    archivos = {}
    app = None          # la certificación del mes con el que se está trabajando en este pedido
    aviso = None        # aviso para mostrar en la próxima respuesta (ej. config abierto en Excel)
    apps = {}           # (año, mes) → certificaciones de ese mes (todos los proveedores y períodos)
    version_cfg = 0     # sube cada vez que alguien cambia el config: los otros meses se ponen al día solos


S = Sesion()


# ----------------------------------------------------------------- sesión (sobrevive a cerrar el programa)
SESION = TRABAJO / "sesion.json"
CONFIG_TRABAJO = TRABAJO / "config_proveedores.xlsx"     # el config con todos los cambios hechos acá


def _guardar_sesion():
    SESION.write_text(json.dumps({k: str(v) for k, v in S.archivos.items()}, ensure_ascii=False, indent=1),
                      encoding="utf-8")


def _ruta_config():
    """El config más nuevo: el que se subió o el que el programa fue actualizando (CUIT, precios, numeración)."""
    cands = [r for r in (S.archivos.get("config"), CONFIG_TRABAJO) if r and Path(r).exists()]
    return max(cands, key=lambda r: Path(r).stat().st_mtime) if cands else None


def _leer_config():
    r = _ruta_config()
    if r:
        S.conf, S.precios, S.par = cb.leer_config(str(r))
    elif S.maestro is not None:
        S.conf, S.precios, S.par = cb.plantilla_config(cb.preparar(S.maestro)[0])
    _numeracion_al_dia()


def _numero_mas_alto():
    """El próximo número libre según TODO lo ya confirmado en la carpeta trabajo (todos los meses)."""
    alto = 0
    for r in TRABAJO.glob("certificaciones_*/estado_*.json"):
        try:
            e = json.loads(r.read_text(encoding="utf-8"))
        except Exception:
            continue
        alto = max(alto, int(e.get("proximo") or 0))
        for m in e.get("meta", {}).values():
            n = str(m.get("numero") or "")
            if "-" in n and n.rsplit("-", 1)[1].isdigit():
                alto = max(alto, int(n.rsplit("-", 1)[1]) + 1)
    for a in S.apps.values():
        alto = max(alto, int(a.estado.get("proximo") or 0))
    return alto


def _numeracion_al_dia():
    """Aunque se suba un config viejo (con un NUMERO_INICIAL más bajo), nunca se repite un número ya entregado."""
    if S.par is None:
        return
    try:
        actual = int(float(S.par.get("NUMERO_INICIAL") or 1))
    except (TypeError, ValueError):
        actual = 1
    S.par["NUMERO_INICIAL"] = max(actual, _numero_mas_alto(), 1)


def _sync():
    """Lo que se cambió en la certificación (CUIT, precios, frecuencia, numeración) queda para la próxima vez."""
    if S.app is not None:
        S.conf, S.precios, S.par = S.app.conf, S.app.precios, S.app.par


def restaurar_sesion():
    if not SESION.exists():
        return
    try:
        S.archivos = {k: Path(v) for k, v in json.loads(SESION.read_text(encoding="utf-8")).items() if Path(v).exists()}
        if "maestro" in S.archivos:
            S.maestro = cb.leer_maestro(str(S.archivos["maestro"]))
        if "calidad" in S.archivos:
            S.humedad = cb.leer_humedad(str(S.archivos["calidad"]))
        if "modelo" in S.archivos:
            S.logo = cb.extraer_logo(str(S.archivos["modelo"]))
        _leer_config()
        print("Se retomaron los archivos de la última vez:", ", ".join(p.name for p in S.archivos.values()))
    except Exception as e:                       # si algo no se puede leer, se empieza de cero
        print("No se pudo retomar la sesión anterior:", e)
        S.archivos = {}


# ----------------------------------------------------------------- utilidades
def _num(v, n=2):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    return round(float(v), n)


def _app(anio=None, mes=None):
    """Toma las certificaciones del mes pedido (cada usuario trabaja con su mes; la numeración es una sola)."""
    if anio is None or mes is None:
        raise HTTPException(400, "Falta indicar el mes.")
    S.app = S.apps.get((int(anio), int(mes)))
    if S.app is None and S.maestro is not None:
        _armar(Armar(anio=int(anio), mes=int(mes)))
    if S.app is None:
        raise HTTPException(400, "Primero armá las certificaciones (paso 2).")
    _activar(S.app)
    return S.app


def _activar(a):
    """Deja listo el mes con el que se va a trabajar: su día de corte de quincena (cada mes tiene el suyo) y la
    numeración al día (si otro usuario confirmó en otro mes)."""
    S.app = a
    if S.par is not None and a.par is not S.par:          # numeración y parámetros: uno solo para todos los meses
        S.par["NUMERO_INICIAL"] = max(int(S.par["NUMERO_INICIAL"]), int(a.par["NUMERO_INICIAL"]))
        a.par = S.par
    if getattr(a, "_version_cfg", 0) != S.version_cfg:    # otro usuario cambió CUIT/precios/frecuencia en otro mes
        cb.fijar_corte(int(a.estado.get("corte") or S.par.get("DIA_CORTE_QUINCENA", 12)))
        a.conf, a.precios = S.conf, S.precios
        a.reconstruir()
        a._version_cfg = S.version_cfg
    cb.fijar_corte(int(a.estado.get("corte") or S.par.get("DIA_CORTE_QUINCENA", 12)))
    a.estado["proximo"] = max(int(a.estado["proximo"]), int(S.par["NUMERO_INICIAL"]))


def _invalidar_otros(cambio_config=False):
    """Si cambió el config (CUIT, precios, frecuencia) los otros meses abiertos se recalculan solos la próxima vez
    que alguien los pida (no se tiran: rearmar un mes completo tarda). La numeración se comparte siempre."""
    if cambio_config and S.app is not None:
        S.version_cfg += 1
        S.app._version_cfg = S.version_cfg


def _usuario(nombre):
    S.app.w_usuario = SimpleNamespace(value=(_quien(nombre) or "web").strip() or "web")


def _resumen_cert(c):
    return {"key": c["key"], "periodo": c["periodo"], "proveedor": c["proveedor"], "razon": c["razon"],
            "cuit": c["cuit"], "estado": c["estado"], "icono": cb.ICONOS[c["estado"]], "numero": c["numero"],
            "tn": _num(c["tn"]), "importe": _num(c["importe"]), "formato": c.get("formato") or "GENERAL",
            "camiones": int(len(c["filas"])), "puntos": sorted(set(c["filas"]["Proveedor"])),
            "con_desc": int((c["filas"]["DescKg"] > 0).sum() + (c["filas"]["Pen"] > 0).sum()),
            "pdf": (S.app.estado["meta"].get(c["key"], {}).get("pdf") if c["estado"] == "CONFIRMADA" else None),
            "excel": (S.app.estado["meta"].get(c["key"], {}).get("excel") if c["estado"] == "CONFIRMADA" else None)}


def _detalle(c):
    a = S.app
    fl = c["filas"]
    filas = []
    for _, f in fl.iterrows():
        filas.append({"id": f["_id"], "fecha": f["Fecha"].strftime("%d/%m/%Y"), "rto": f["Rto. Planta"],
                      "punto": f["Proveedor"], "rto_prov": f["Remito Proveedor"], "producto": f["Producto"],
                      "kilos": _num(f["Kilos"], 1), "hum": _num(f.get("Hum"), 4), "finos": _num(f.get("Finos"), 4),
                      "desc_gran": _num(f.get("DescPct"), 2), "desc_hum": _num((f.get("Pen") or 0) * 100, 2),
                      "kilos_neto": _num(f.get("KilosNeto"), 1), "aserrin": _num(f.get("DescKg"), 1),
                      "man_gran": bool(f.get("DescManual")), "man_hum": bool(f.get("PenManual"))})
    productos = list(dict.fromkeys([f["producto"] for f in c.get("tramos", [])] + list(c["precios"])))
    precios = []
    for t in productos:
        entregado = t in c["precios"]
        p = c["precios"][t] if entregado else cb.buscar_precio(a.precios, c["grupo"], t, c["tn_grupo"])
        tn = cb.tn_pagar(c, t) if entregado else 0.0
        precios.append({"producto": t, "etiqueta": (cb.etiqueta_mym(t) if c.get("formato") == "MYM" else cb.nombre_producto(t)),
                        "precio": _num(p), "tn": _num(tn), "subtotal": _num(tn * (p or 0)), "entregado": entregado,
                        "origen": c["origen_precios"].get(t, "config" if p is not None else "sin precio")})
    meta = a.estado["meta"].get(c["key"], {})
    return dict(_resumen_cert(c), filas=filas, precios=precios,
                obs=[{"sev": o["sev"], "msg": o["msg"], "rem": o["rem"]} for o in c["obs"]],
                criticas=cb.criticas_revisables(c), revisada=bool(meta.get("ack")),
                frecuencia=cb.frecuencia(a.conf, c["proveedor"]), sin_gran=bool(c.get("sin_gran")),
                sin_hum=bool(c.get("sin_hum")), desde=c["desde"].strftime("%d/%m/%Y"),
                hasta=c["hasta"].strftime("%d/%m/%Y"), fecha=c["fecha"].strftime("%d/%m/%Y"),
                tol_finos=c.get("tol_finos"), pdf=meta.get("pdf"), excel=meta.get("excel"))


def _corresponde(c, periodos=None, proveedor=None):
    if periodos and c["periodo"] not in periodos:
        return False
    if proveedor:
        return proveedor == c["proveedor"] or proveedor in (c.get("miembros") or []) \
            or cb.unidad(S.app.conf, proveedor) == c["proveedor"]
    return True


def _estado_general(periodos=None, proveedor=None):
    a = S.app
    certs = sorted([c for c in a.certs.values() if _corresponde(c, periodos, proveedor)],
                   key=lambda c: (["1Q", "2Q", "Mensual"].index(c["periodo"]), c["proveedor"].casefold()))
    cuad = a.cuad.fillna(0)
    if periodos and "Período" in cuad.columns:
        cuad = cuad[cuad["Período"].isin(periodos)]
    cuad = cuad.to_dict("records")
    return {"anio": a.estado["anio"], "mes": a.estado["mes"], "mes_nombre": cb.MESES[a.estado["mes"]],
            "corte": cb.DIA_CORTE, "proximo": f"{a.estado['prefijo']}-{a.estado['proximo']:08d}",
            "resumen": a.resumen_estados(), "certs": [_resumen_cert(c) for c in certs], "cuadre": cuad,
            "confirmadas": sum(1 for c in a.certs.values() if c["estado"] == "CONFIRMADA")}


# ----------------------------------------------------------------- imágenes personalizadas
IMAGENES = AQUI / "imagenes"
TIPOS_IMAGEN = {"logo": "Logo del programa (arriba a la izquierda y en la pestaña del navegador)",
                "fondo": "Imagen de fondo del encabezado",
                "fondo_pagina": "Fondo de la página (se ve suave, detrás de todo, como marca de agua)",
                "logo_cert": "Logo de las certificaciones (PDF y Excel)"}
MAX_IMG_MB = 5


def _ruta_imagen(tipo):
    for ext in (".png", ".jpg"):
        r = IMAGENES / f"{tipo}{ext}"
        if r.exists():
            return r
    return None


def _logo_cert():
    """Logo de los PDF/Excel: el personalizado (imagenes/logo_cert) → el de la certificación modelo → el incorporado."""
    r = _ruta_imagen("logo_cert")
    if r:
        return r.read_bytes()
    return S.logo or cb.logo_por_defecto()


def _recortar_bordes(im, margen=6):
    from PIL import ImageChops
    rgba = im.convert("RGBA")
    fondo = rgba.getpixel((0, 0))
    if fondo[3] < 10:                                         # fondo transparente
        caja = rgba.getchannel("A").point(lambda a: 255 if a > 10 else 0).getbbox()
    else:                                                     # fondo liso (blanco u otro color)
        dif = ImageChops.difference(rgba.convert("RGB"), Image_new(rgba.size, fondo[:3])).convert("L")
        caja = dif.point(lambda v: 255 if v > 24 else 0).getbbox()
    if not caja:
        return im
    x0, y0, x1, y1 = caja
    return rgba.crop((max(0, x0 - margen), max(0, y0 - margen), min(rgba.width, x1 + margen), min(rgba.height, y1 + margen)))


def Image_new(tam, color):
    from PIL import Image
    return Image.new("RGB", tam, color)


@app_web.get("/api/imagenes")
def imagenes():
    return {t: {"descripcion": d, "hay": _ruta_imagen(t) is not None,
                "version": int(_ruta_imagen(t).stat().st_mtime) if _ruta_imagen(t) else 0,
                "foto": bool(_ruta_imagen(t) and _ruta_imagen(t).suffix == ".jpg")}
            for t, d in TIPOS_IMAGEN.items()}


@app_web.get("/api/imagen/{tipo}")
def imagen(tipo: str):
    if tipo not in TIPOS_IMAGEN:
        raise HTTPException(404, "No existe esa imagen.")
    r = _ruta_imagen(tipo)
    if r is None:
        if tipo == "logo_cert":
            from fastapi.responses import Response
            return Response(_logo_cert(), media_type="image/png")
        raise HTTPException(404, "Sin imagen personalizada.")
    return FileResponse(r, headers={"Cache-Control": "no-cache"})


@app_web.post("/api/imagen/{tipo}")
async def subir_imagen(tipo: str, archivo: UploadFile = File(...)):
    """Se acepta PNG, JPG o WEBP de hasta 5 MB. Se verifica que sea una imagen de verdad y se guarda como PNG/JPG
    (el logo se achica a 600 px; el fondo a 2000 px) en la carpeta imagenes junto a servidor.py."""
    from PIL import Image
    if tipo not in TIPOS_IMAGEN:
        raise HTTPException(404, "No existe esa imagen.")
    datos = await archivo.read()
    if len(datos) > MAX_IMG_MB * 1024 * 1024:
        raise HTTPException(400, f"La imagen pesa más de {MAX_IMG_MB} MB.")
    try:
        im = Image.open(io.BytesIO(datos))
        im.load()
    except Exception:
        raise HTTPException(400, "El archivo no es una imagen válida (PNG, JPG o WEBP).") from None
    if im.format not in ("PNG", "JPEG", "WEBP"):
        raise HTTPException(400, "Solo se aceptan imágenes PNG, JPG o WEBP.")
    area = im.width * im.height
    if tipo != "fondo":
        im = _recortar_bordes(im)             # saca el margen blanco o transparente alrededor del logo
    # fondo de página: si es una foto (casi sin margen liso) se muestra cubriendo toda la página;
    # si es un logo, centrado y suave como marca de agua
    foto = tipo == "fondo" or (tipo == "fondo_pagina" and im.width * im.height >= 0.8 * area)
    im.thumbnail((2000, 2000) if foto else (1200, 1200) if tipo == "fondo_pagina" else (600, 600))
    IMAGENES.mkdir(exist_ok=True)
    with CANDADO:
        for viejo in IMAGENES.glob(f"{tipo}.*"):
            viejo.unlink()
        if foto:
            im.convert("RGB").save(IMAGENES / f"{tipo}.jpg", "JPEG", quality=85)
        else:
            im.convert("RGBA").save(IMAGENES / f"{tipo}.png", "PNG")
        if tipo == "logo_cert":
            for a in S.apps.values():
                a.logo = _logo_cert()
    return imagenes()


ARCHIVO_TEMA = IMAGENES / "tema.json"
COLOR_BASE = "#3A7D44"


@app_web.get("/api/tema")
def tema():
    try:
        return json.loads(ARCHIVO_TEMA.read_text(encoding="utf-8"))
    except Exception:
        return {"color": COLOR_BASE}


class Tema(BaseModel):
    color: str | None = None          # "#RRGGBB"; None = volver al verde de siempre
    desde_logo: bool = False          # tomar el color principal del logo del programa


@app_web.post("/api/tema")
def fijar_tema(t: Tema):
    import re as _re
    color = t.color or COLOR_BASE
    if t.desde_logo:
        r = _ruta_imagen("logo")
        if r is None:
            raise HTTPException(400, "Primero subí el logo del programa.")
        color = _color_principal(r)
    if not _re.fullmatch(r"#[0-9A-Fa-f]{6}", color):
        raise HTTPException(400, "Color no válido.")
    IMAGENES.mkdir(exist_ok=True)
    ARCHIVO_TEMA.write_text(json.dumps({"color": color.upper()}), encoding="utf-8")
    return tema()


def _color_principal(ruta):
    """El color más presente del logo, sin contar blancos, grises ni transparentes."""
    from PIL import Image
    import colorsys
    im = Image.open(ruta).convert("RGBA").resize((120, 120))
    cuenta = {}
    for r, g, b, a in im.getdata():
        if a < 128:
            continue
        h, l, sat = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
        if sat < 0.25 or l > 0.85 or l < 0.12:
            continue
        k = (r // 16, g // 16, b // 16)
        cuenta.setdefault(k, []).append((r, g, b))
    if not cuenta:
        return COLOR_BASE
    px = max(cuenta.values(), key=len)
    r, g, b = (sum(c[i] for c in px) // len(px) for i in range(3))
    h, l, sat = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
    if l > 0.45:                                    # que el texto blanco del encabezado se lea bien
        r, g, b = (int(x * 255) for x in colorsys.hls_to_rgb(h, 0.42, sat))
    return f"#{r:02X}{g:02X}{b:02X}"


@app_web.delete("/api/imagen/{tipo}")
def quitar_imagen(tipo: str):
    if tipo not in TIPOS_IMAGEN:
        raise HTTPException(404, "No existe esa imagen.")
    with CANDADO:
        for viejo in IMAGENES.glob(f"{tipo}.*"):
            viejo.unlink()
        if tipo == "logo_cert":
            for a in S.apps.values():
                a.logo = _logo_cert()
    return imagenes()


# ----------------------------------------------------------------- página
@app_web.get("/", response_class=HTMLResponse)
def pagina():
    pagina_html = AQUI / "index.html" if (AQUI / "index.html").exists() else AQUI / "static" / "index.html"
    return pagina_html.read_text(encoding="utf-8")


# ----------------------------------------------------------------- API
@app_web.post("/api/archivos")
async def subir(archivos: list[UploadFile] = File(...)):
    """Sube maestro (.xlsb), planilla de calidad, config y estado (opcionales). Se reconocen por su contenido."""
    carpeta = TRABAJO / "subidos"
    carpeta.mkdir(exist_ok=True)
    tipos = {}
    for f in archivos:
        nombre = Path(f.filename or "").name
        if Path(nombre).suffix.lower() not in EXTENSIONES:
            raise HTTPException(400, f"'{nombre}': solo se aceptan .xlsb, .xlsx y .json (maestro, calidad, config, estado).")
        datos = await f.read()
        if len(datos) > MAX_MB * 1024 * 1024:
            raise HTTPException(400, f"'{nombre}' pesa más de {MAX_MB} MB.")
        ruta = carpeta / nombre
        ruta.write_bytes(datos)
        try:
            tipos[ruta.name] = (cb.tipo_archivo(str(ruta)), ruta)
        except Exception:
            tipos[ruta.name] = (None, ruta)
    with CANDADO:
        return _cargar(tipos)


def _cargar(tipos):
    for nombre, (tipo, ruta) in tipos.items():
        if tipo:
            S.archivos[tipo] = ruta
    _guardar_sesion()
    if "maestro" in [t for t, _ in tipos.values()]:
        S.maestro = cb.leer_maestro(str(S.archivos["maestro"]))
    if "calidad" in [t for t, _ in tipos.values()]:
        S.humedad = cb.leer_humedad(str(S.archivos["calidad"]))
    if "config" in [t for t, _ in tipos.values()]:
        S.conf, S.precios, S.par = cb.leer_config(str(S.archivos["config"]))      # el que se acaba de subir
        _numeracion_al_dia()
    elif S.conf is None:
        _leer_config()
    if "modelo" in S.archivos:
        S.logo = cb.extraer_logo(str(S.archivos["modelo"]))
    S.app = None
    S.apps.clear()
    return {"reconocidos": {n: (t or "no reconocido") for n, (t, _) in tipos.items()}, **_info_archivos()}


def _info_archivos():
    info = {"maestro": None, "calidad": None, "config": None}
    if S.maestro is not None:
        d = cb.preparar(S.maestro, S.conf)[0]
        meses = sorted({(int(x.year), int(x.month)) for x in d[cb.C_DIA].dropna()})
        info["maestro"] = {"archivo": S.archivos["maestro"].name, "filas": int(len(S.maestro)),
                           "meses": [{"anio": a, "mes": m, "nombre": f"{cb.MESES[m]} {a}"} for a, m in meses]}
    if S.humedad is not None:
        info["calidad"] = {"archivo": S.archivos["calidad"].name, "muestras": int(len(S.humedad)),
                           "ultima": S.humedad["Fecha"].max().strftime("%d/%m/%Y")}
    if S.conf is not None:
        info["config"] = {"archivo": S.archivos["config"].name if "config" in S.archivos else "(plantilla nueva)",
                          "proveedores": int(len(S.conf)), "con_cuit": int((S.conf["CUIT"].astype(str).str.strip() != "").sum()),
                          "corte": int(S.par.get("DIA_CORTE_QUINCENA", 12)), "proximo": int(S.par["NUMERO_INICIAL"])}
    return info


@app_web.get("/api/archivos")
def info_archivos():
    with CANDADO:                 # si está retomando la sesión anterior, espera a que termine
        return _info_archivos()


@app_web.get("/api/proveedores")
def proveedores(anio: int, mes: int):
    with CANDADO:
        return _proveedores(anio, mes)


def _proveedores(anio, mes):
    if S.maestro is None:
        raise HTTPException(400, "Subí el maestro primero.")
    d = cb.preparar(S.maestro, S.conf)[0]
    x = d[(d[cb.C_DIA].dt.year == anio) & (d[cb.C_DIA].dt.month == mes)]
    g = x.groupby(cb.C_PROV)[cb.C_NETO].agg(["size", "sum"]).sort_index(key=lambda s: s.str.casefold())
    return [{"nombre": p, "remitos": int(r["size"]), "tn": round(r["sum"] / 1000, 2),
             "frecuencia": cb.frecuencia(S.conf, cb.unidad(S.conf, p))} for p, r in g.iterrows()]


class Armar(BaseModel):
    anio: int
    mes: int
    periodos: list[str] = ["1Q", "2Q", "Mensual"]
    corte: int | None = None
    proveedor: str | None = None


@app_web.post("/api/armar")
def armar(p: Armar):
    u = _usuario_actual()
    if u and u.get("rol") == "consulta":
        p.corte = None              # consulta: ve el mes con el corte ya elegido, no lo cambia
    with CANDADO:
        r = _armar(p)
        _sync()
        return r


def _armar(p):
    _sync()
    if S.maestro is None:
        raise HTTPException(400, "Subí el maestro primero.")
    k = (p.anio, p.mes)
    a = S.apps.get(k)
    confirmadas = a is not None and any(m.get("huella") for m in a.estado["meta"].values())
    if a is not None and (p.corte is None or int(a.estado.get("corte") or 0) == int(p.corte) or confirmadas):
        _activar(a)                                 # ya está armado: se usa el mismo (otro usuario puede estar en él)
    else:
        a = _nuevo_app(p)
        S.apps[k] = a
        _activar(a)
    return _estado_general(p.periodos, p.proveedor)


def _nuevo_app(p):
    if S.maestro is None:
        raise HTTPException(400, "Subí el maestro primero.")
    salida = TRABAJO / f"certificaciones_{p.anio}-{p.mes:02d}"
    # estado del mes: el más nuevo entre el que se subió y el que el programa fue guardando
    estado = None
    ruta_est = salida / f"estado_{p.anio}-{p.mes:02d}.json"
    cands = [r for r in (S.archivos.get("estado"), ruta_est) if r and Path(r).exists()]
    for r in sorted(cands, key=lambda r: Path(r).stat().st_mtime, reverse=True):
        e = json.loads(Path(r).read_text(encoding="utf-8"))
        if (int(e["anio"]), int(e["mes"])) == (p.anio, p.mes):
            estado = e
            break
    # se arma el mes completo (todos los proveedores y períodos); cada usuario ve lo que eligió
    a = ed.App(S.maestro, S.conf, S.precios, S.par, p.anio, p.mes, ["1Q", "2Q", "Mensual"], salida=str(salida),
               logo=_logo_cert(), proveedores=None,
               estado=estado, origen=S.archivos["maestro"].name, descargar_pdf=False, humedad=S.humedad,
               corte=p.corte, ruta_config=str(CONFIG_TRABAJO))
    a.w_usuario = SimpleNamespace(value="web")
    a._version_cfg = S.version_cfg
    _guardado_seguro(a)
    a.guardar_estado()
    return a


def _guardado_seguro(a):
    """Si alguien tiene abierto el config en Excel, no se puede escribir: el estado (numeración, confirmadas) ya
    quedó guardado, así que no se corta nada; se avisa y se vuelve a intentar en el próximo guardado."""
    original = a.guardar_estado

    def guardar_estado():
        try:
            original()
        except PermissionError:
            S.aviso = ("Ojo: trabajo/config_proveedores.xlsx está abierto en Excel y no se pudo actualizar. "
                       "Cerralo: se guarda solo en el próximo cambio (las certificaciones no se perdieron).")
    a.guardar_estado = guardar_estado


def _con_aviso(r):
    """Agrega a la respuesta un aviso pendiente (por ejemplo config abierto en Excel)."""
    aviso, S.aviso = getattr(S, "aviso", None), None
    if aviso and isinstance(r, dict):
        r = dict(r, aviso=aviso)
    return r


def _filtros(periodos, proveedor):
    return ([x for x in (periodos or "").split(",") if x] or None), (proveedor or None)


@app_web.get("/api/certificaciones")
def certificaciones(anio: int, mes: int, periodos: str = "", proveedor: str = ""):
    with CANDADO:
        _app(anio, mes)
        return _estado_general(*_filtros(periodos, proveedor))


@app_web.get("/api/certificacion")
def certificacion(key: str, anio: int, mes: int):
    with CANDADO:
        a = _app(anio, mes)
        if key not in a.certs:
            raise HTTPException(404, "No existe esa certificación (puede que otro usuario la haya cambiado: actualizá).")
        return _detalle(a.certs[key])


class Cambios(BaseModel):
    usuario: str
    motivo: str = ""
    razon: str | None = None
    cuit: str | None = None
    frecuencia: str | None = None
    sin_gran: bool | None = None
    sin_hum: bool | None = None
    desc_gran: dict[str, float | None] = {}     # _id → % que pasa a aserrín (None = volver al automático)
    desc_hum: dict[str, float | None] = {}      # _id → % de descuento por humedad
    precios: dict[str, float] = {}               # producto → $ por TN
    precios_solo_esta: bool = False
    revisada: bool = False


@app_web.post("/api/certificacion")
def guardar(key: str, anio: int, mes: int, cb_: Cambios, periodos: str = "", proveedor: str = ""):
    with CANDADO:
        _app(anio, mes)
        r, cambio_config = _guardar(key, cb_, *_filtros(periodos, proveedor))
        _sync()
        _invalidar_otros(cambio_config)
        return _con_aviso(r)


def _guardar(key, cb_, periodos=None, proveedor=None):
    a = S.app
    if key not in a.certs:
        raise HTTPException(404, "No existe esa certificación.")
    cb_.usuario = _quien(cb_.usuario)
    _usuario(cb_.usuario)
    c = a.certs[key]
    meta = a.estado["meta"].setdefault(key, {})
    corr = a.estado["corr"]
    hay_desc = cb_.desc_gran or cb_.desc_hum or (cb_.sin_gran is not None and cb_.sin_gran != bool(meta.get("sin_gran"))) \
        or (cb_.sin_hum is not None and cb_.sin_hum != bool(meta.get("sin_hum")))
    if hay_desc and not cb_.motivo.strip():
        raise HTTPException(400, "Escribí el motivo del cambio en los descuentos.")
    if cb_.cuit and not cb.cuit_valido(cb_.cuit):
        raise HTTPException(400, f"El CUIT {cb_.cuit} no es válido (dígito verificador).")
    # --- primero se controla TODO; recién después se cambia algo (si hay un error no queda nada a medias)
    ids = set(c["filas"]["_id"])
    producto = dict(zip(c["filas"]["_id"], c["filas"]["Producto"]))
    for campo, dic, validos in (("Desc. granulometría %", cb_.desc_gran, ("CHIP",)),
                                ("Desc. humedad %", cb_.desc_hum, ("CHIP", "ASERRIN"))):
        for _id, v in dic.items():
            if _id not in ids:
                raise HTTPException(409, "Un camión de esta certificación cambió (otro usuario la modificó): tocá ↻ y volvé a cargar.")
            if producto[_id] not in validos:
                raise HTTPException(400, f"{campo}: solo se aplica a {' y '.join(validos)} (ese camión es {producto[_id]}).")
            if v is not None and not (0 <= v <= 100):
                raise HTTPException(400, f"{campo}: {v} fuera de rango (0 a 100).")
    for t, v in cb_.precios.items():
        if v is not None and (v < 0 or v > 1e9):
            raise HTTPException(400, f"Precio de {t} fuera de rango.")
    if cb_.frecuencia is not None:
        cb_.frecuencia = cb_.frecuencia.strip().upper()
        if cb_.frecuencia not in cb.FRECUENCIAS:
            raise HTTPException(400, "Frecuencia no válida (QUINCENAL, MENSUAL, AMBAS o vacío).")
        if cb_.frecuencia != cb.frecuencia(a.conf, c["proveedor"]):
            confirmadas = [x["numero"] for k, x in a.certs.items() if x["proveedor"] == c["proveedor"]
                           and a.estado["meta"].get(k, {}).get("huella")]
            if confirmadas:
                raise HTTPException(400, "No se puede cambiar la frecuencia en este mes: ya tiene certificaciones "
                                         f"confirmadas ({', '.join(confirmadas)}). Cambiala desde el mes siguiente.")
    if cb_.razon is not None and not cb_.razon.strip():
        raise HTTPException(400, "La razón social no puede quedar vacía.")
    for campo, dic, destino, escala in (("Desc. granulometría %", cb_.desc_gran, "desc_gran", 1),
                                        ("Desc. humedad %", cb_.desc_hum, "pen_hum", 0.01)):
        for _id, v in dic.items():
            rem = c["filas"].set_index("_id").loc[_id, "Rto. Planta"] if _id in set(c["filas"]["_id"]) else ""
            if v is None:
                corr.setdefault(destino, {}).pop(_id, None)
            else:
                corr.setdefault(destino, {})[_id] = v * escala
            a._log(f"{campo} a mano (web)", key, rem, campo, "", "automático" if v is None else f"{v:g} %", cb_.motivo)
    for k, v in (("sin_gran", cb_.sin_gran), ("sin_hum", cb_.sin_hum)):
        if v is not None and v != bool(meta.get(k)):
            if v:
                meta[k] = True
            else:
                meta.pop(k, None)
            a._log(("Descuento de granulometría" if k == "sin_gran" else "Descuento de humedad")
                   + (" QUITADO" if v else " vuelve a aplicarse"), key, "", "", "", "", cb_.motivo)
    # todos los puntos de entrega de la unidad, aunque en este período no hayan entregado
    miembros = sorted(set(c.get("miembros") or [c["proveedor"]]) | set(cb.miembros(a.conf, c["proveedor"])))
    cambio_config, todo = False, False
    for campo, valor in (("RAZÓN SOCIAL", cb_.razon), ("CUIT", cb_.cuit), ("FRECUENCIA", cb_.frecuencia)):
        if valor is None:
            continue
        if campo not in a.conf.columns:
            a.conf[campo] = ""
        for m in miembros:
            if m not in a.conf.index:
                a.conf.loc[m] = {col: "" for col in a.conf.columns}
                a.conf.loc[m, ["GRUPO", "CERTIFICAR"]] = [c["grupo"], "SI"]
                todo = True                       # proveedor nuevo en el config: se rearma todo (quita el aviso)
            if str(a.conf.at[m, campo]) != valor:
                a._log("Config proveedor (web)", key, "", campo, a.conf.at[m, campo], valor)
                a.conf.at[m, campo] = valor
                cambio_config = True
                todo = todo or campo == "FRECUENCIA"          # cambia qué certificaciones existen
    for t, v in cb_.precios.items():
        if v is None:
            continue
        if cb_.precios_solo_esta:
            meta.setdefault("precios", {})[t] = v
            a._log("Precio solo para esta certificación (web)", key, "", t, "", v)
        else:
            a.precios = cb.fijar_precio(a.precios, c["grupo"], t, c["tn_grupo"], v)
            cambio_config = True
            meta.get("precios", {}).pop(t, None)
            a._log("Precio en config (web)", key, "", f"{c['grupo']} / {t}", "", v)
    if cb_.revisada and cb.criticas_revisables(c):
        a.marcar_revisadas(cb_.usuario, cb_.motivo or "Revisado en la web", [key])
    if todo:
        a.reconstruir()                       # cambio de frecuencia: se rearma el mes completo
    else:
        _rearmar(a, c)                        # solo las certificaciones de ese proveedor / grupo de precios
    a.guardar_estado()
    if key not in a.certs:
        return {"quitada": True, "msg": "Con la frecuencia elegida este proveedor ya no lleva esta certificación.",
                **_estado_general(periodos, proveedor)}, cambio_config
    return _detalle(a.certs[key]), cambio_config


def _rearmar(a, c):
    """Recalcula solo lo que puede cambiar con una edición: las certificaciones del mismo proveedor y las que
    comparten el grupo de precios (1Q, 2Q y Mensual). Es ~20 veces más rápido que rearmar el mes completo y
    da el mismo resultado (los controles de remitos del mes se reutilizan)."""
    a.d = cb.aplicar_correcciones(a.d0, a.estado["corr"])
    g = str(c.get("grupo") or "").casefold()
    keys = {k for k, x in a.certs.items()
            if x["proveedor"] == c["proveedor"] or str(x.get("grupo") or "").casefold() == g}
    nuevos, _, _ = cb.construir(a.d, a.d0, a.conf, a.precios, a.estado, a.tipo, a.raw, a.proveedores,
                                a.humedad, a.metodo_humedad, solo=keys, validacion=a._validacion)
    for k in keys:
        if k in nuevos:
            a.certs[k] = nuevos[k]
        else:
            a.certs.pop(k, None)


class Usuario(BaseModel):
    usuario: str
    incluir_revisar: bool = True
    anio: int                      # obligatorio: cada pedido dice de qué mes es (varios usuarios a la vez)
    mes: int
    periodos: list[str] | None = None
    proveedor: str | None = None
    numero: int | None = None      # Nº de certificación puesto a mano (solo Administrador; opcional)


@app_web.post("/api/confirmar")
def confirmar(key: str, u: Usuario):
    with CANDADO:
        _app(u.anio, u.mes)
        r = _confirmar(key, u)
        _sync()
        _invalidar_otros(False)
        return _con_aviso(r)


def _confirmar(key, u):
    a = S.app
    u.usuario = _quien(u.usuario)
    _usuario(u.usuario)
    if u.incluir_revisar and a.certs.get(key, {}).get("estado") == "REVISAR":
        a.marcar_revisadas(u.usuario, f"Revisado en la web ({u.usuario})", [key])
    deshacer = None
    if u.numero:
        deshacer = _numero_a_mano(a, key, int(u.numero))
    ok, msg, rutas = a.confirmar_cert(key, u.usuario)
    if not ok and deshacer:
        deshacer()
    a.guardar_estado()
    if not ok:
        raise HTTPException(400, msg)
    return _detalle(a.certs[key])


def _numeros_usados():
    usados = set()
    estados = [json.loads(r.read_text(encoding="utf-8")) for r in TRABAJO.glob("certificaciones_*/estado_*.json")]
    for e in estados + [x.estado for x in S.apps.values()]:
        for m in e.get("meta", {}).values():
            n = str(m.get("numero") or "")
            if "-" in n and n.rsplit("-", 1)[1].isdigit():
                usados.add(int(n.rsplit("-", 1)[1]))
    return usados


def _numero_a_mano(a, key, n):
    """Nº elegido a mano para ESTA certificación (ej. para respetar una numeración hecha fuera del programa).
    Solo el Administrador; nunca un número ya usado; si es mayor al próximo, la numeración sigue desde ahí."""
    u = _usuario_actual()
    if LOGIN_ACTIVO and (not u or u.get("rol") != "admin"):
        raise HTTPException(403, "Solo un Administrador puede poner el número de certificación a mano.")
    if key not in a.certs:
        raise HTTPException(404, "No existe esa certificación.")
    meta = a.estado["meta"].setdefault(key, {})
    if meta.get("numero"):
        raise HTTPException(400, f"Esta certificación ya tiene el Nº {meta['numero']}: no se puede cambiar.")
    if not 1 <= n <= 99_999_999:
        raise HTTPException(400, "Número fuera de rango.")
    if n in _numeros_usados():
        raise HTTPException(400, f"El Nº {n} ya se usó en otra certificación.")
    prox_antes, par_antes = int(a.estado["proximo"]), int(S.par["NUMERO_INICIAL"]) if S.par else None
    meta["numero"] = f"{a.estado['prefijo']}-{n:08d}"
    if n >= int(a.estado["proximo"]):
        a.estado["proximo"] = n + 1
        if S.par is not None:
            S.par["NUMERO_INICIAL"] = max(int(S.par["NUMERO_INICIAL"]), n + 1)
    a.recalcular(key)
    uu = u["usuario"] if u else "local"
    _registrar_acceso(uu, "-", f"Nº PUESTO A MANO: {meta['numero']} en {key} ({a.estado['mes']}/{a.estado['anio']})")

    def deshacer():
        meta.pop("numero", None)
        a.estado["proximo"] = prox_antes
        if S.par is not None and par_antes is not None:
            S.par["NUMERO_INICIAL"] = par_antes
        a.recalcular(key)
    return deshacer


@app_web.post("/api/confirmar_todos")
def confirmar_todos(u: Usuario):
    with CANDADO:
        _app(u.anio, u.mes)
        r = _confirmar_todos(u)
        _sync()
        _invalidar_otros(False)
        return _con_aviso(r)


def _confirmar_todos(u):
    a = S.app
    u.usuario = _quien(u.usuario)
    _usuario(u.usuario)
    elegidas = [k for k, c in a.certs.items() if _corresponde(c, u.periodos, u.proveedor)]
    if u.incluir_revisar:
        rev = [k for k in elegidas if a.certs[k]["estado"] == "REVISAR"]
        if rev:
            a.marcar_revisadas(u.usuario, f"Certificación en lote ({u.usuario})", rev)
    hechas, pendientes = [], []
    for k in sorted(elegidas):
        c = a.certs.get(k)
        if c is None or c["estado"] == "CONFIRMADA":
            continue
        ok, msg, rr = a.confirmar_cert(k, u.usuario)
        (hechas if ok else pendientes).append(k)
    a.guardar_cierre()
    return {"hechas": len(hechas), "pendientes": pendientes, **_estado_general(u.periodos, u.proveedor)}


@app_web.get("/api/descargar")
def descargar(ruta: str):
    r = Path(ruta).resolve()
    if TRABAJO.resolve() not in r.parents or not r.exists():
        raise HTTPException(404, "Archivo no encontrado.")
    return FileResponse(r, filename=r.name)


@app_web.get("/api/descargar_juntas")
def descargar_juntas(anio: int, mes: int, tipo: str = "pdf", periodos: str = "", proveedor: str = ""):
    """Todas las certificaciones CONFIRMADAS de lo que está en pantalla, juntas en un solo archivo:
    tipo=pdf → un PDF con todas (en orden de número) · tipo=excel → un Excel con una hoja por certificación
    · tipo=zip → todos los PDF y Excel sueltos en carpetas."""
    import zipfile
    from openpyxl import Workbook
    from fastapi.responses import StreamingResponse
    with CANDADO:
        a = _app(anio, mes)
        per, prov = _filtros(periodos, proveedor)
        certs = sorted([c for c in a.certs.values() if c["estado"] == "CONFIRMADA" and _corresponde(c, per, prov)],
                       key=lambda c: c["numero"])
        if not certs:
            raise HTTPException(404, "No hay certificaciones confirmadas en lo que está en pantalla.")
        meta = a.estado["meta"]
        base = f"Certificaciones_{cb.MESES[mes]}_{anio}" + (f"_{'_'.join(per)}" if per else "")
        buf = io.BytesIO()
        if tipo == "pdf":
            from pypdf import PdfWriter
            w = PdfWriter()
            for c in certs:
                r = meta.get(c["key"], {}).get("pdf")
                if r and Path(r).exists():
                    w.append(str(r), outline_item=f"{c['numero']} · {c['periodo']} · {c['razon']}")
            w.write(buf)
            nombre, mt = base + ".pdf", "application/pdf"
        elif tipo == "excel":
            wb = Workbook()
            ix = wb.active
            ix.title = "Índice"
            ix.append(["N° Certificación", "Período", "Proveedor", "Razón social", "Camiones", "TN", "Importe", "Hoja"])
            usados = {"índice"}
            for c in certs:
                hoja = cb._nombre_hoja(f"{c['periodo']} {c['proveedor']}", usados)
                cb._hoja_certificacion(wb.create_sheet(hoja), c, a.logo)
                ix.append([c["numero"], c["periodo"], c["proveedor"], c["razon"], len(c["filas"]), c["tn"], c["importe"], hoja])
            from openpyxl.styles import Font
            for r in range(2, ix.max_row + 1):                       # la columna Hoja lleva a cada certificación
                ix.cell(r, 8).hyperlink = "#'" + str(ix.cell(r, 8).value).replace("'", "''") + "'!A1"
                ix.cell(r, 8).font = Font(color="0563C1", underline="single")
            ix.append([])
            ix.append(["TOTAL", "", f"{len(certs)} certificaciones", "", sum(len(c["filas"]) for c in certs),
                       round(sum(c["tn"] for c in certs), 2), round(sum(c["importe"] or 0 for c in certs), 2)])
            for celda in list(ix[1]) + list(ix[ix.max_row]):
                celda.font = Font(bold=True)
            for col, ancho in zip("ABCDEFGH", [18, 9, 34, 34, 10, 10, 18, 28]):
                ix.column_dimensions[col].width = ancho
            for fila in ix.iter_rows(min_row=2):
                fila[5].number_format = "#,##0.00"
                fila[6].number_format = cb.MONEDA
            ix.freeze_panes = "A2"
            wb.save(buf)
            nombre, mt = base + ".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        elif tipo == "zip":
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
                for c in certs:
                    m = meta.get(c["key"], {})
                    for clave, carpeta in (("pdf", "PDF"), ("excel", "EXCEL")):
                        if m.get(clave) and Path(m[clave]).exists():
                            z.write(m[clave], f"{carpeta}/{Path(m[clave]).name}")
            nombre, mt = base + ".zip", "application/zip"
        else:
            raise HTTPException(400, "tipo tiene que ser pdf, excel o zip.")
    buf.seek(0)
    return StreamingResponse(buf, media_type=mt, headers={"Content-Disposition": f'attachment; filename="{nombre}"'})


def _ultimo_entregado():
    """El número más alto ya entregado en una certificación confirmada (cualquier mes)."""
    alto = 0
    estados = [json.loads(r.read_text(encoding="utf-8")) for r in TRABAJO.glob("certificaciones_*/estado_*.json")]
    for e in estados + [a.estado for a in S.apps.values()]:
        for m in e.get("meta", {}).values():
            n = str(m.get("numero") or "")
            if "-" in n and n.rsplit("-", 1)[1].isdigit():
                alto = max(alto, int(n.rsplit("-", 1)[1]))
    return alto


class Numeracion(BaseModel):
    proximo: int


@app_web.post("/api/numeracion")
def fijar_numeracion(n: Numeracion):
    """El Administrador fija el próximo número de certificación (ej. para seguir la numeración de las
    certificaciones hechas a mano). Nunca puede quedar por debajo de un número ya entregado."""
    with CANDADO:
        if S.par is None:
            raise HTTPException(400, "Subí primero el config de proveedores.")
        minimo = _ultimo_entregado() + 1
        if n.proximo < minimo:
            raise HTTPException(400, f"No puede ser menor a {minimo}: ese número (o uno anterior) ya se entregó "
                                     "en una certificación confirmada.")
        if n.proximo > 99_999_999:
            raise HTTPException(400, "Número fuera de rango.")
        anterior = int(S.par["NUMERO_INICIAL"])
        S.par["NUMERO_INICIAL"] = int(n.proximo)
        for a in S.apps.values():
            a.par = S.par
            a.estado["proximo"] = int(n.proximo)
            a.guardar_estado()
        if not S.apps:
            cb.guardar_config(CONFIG_TRABAJO, S.conf, S.precios, S.par)
        u = _usuario_actual()
        _registrar_acceso(u["usuario"] if u else "local", "-", f"PRÓXIMO Nº DE CERTIFICACIÓN: {anterior} → {n.proximo}")
        return {"proximo": f"{S.par.get('PREFIJO', '0001')}-{int(n.proximo):08d}", "numero": int(n.proximo)}


@app_web.get("/api/descargar_config")
def descargar_config():
    with CANDADO:
        if S.app is not None:
            S.app.guardar_estado()
        if not CONFIG_TRABAJO.exists():
            raise HTTPException(404, "Todavía no hay config guardado: armá un mes primero.")
        return FileResponse(CONFIG_TRABAJO, filename="config_proveedores.xlsx")


if __name__ == "__main__":
    import uvicorn
    puerto = PUERTO = int(next((a for a in sys.argv[1:] if a.isdigit()), 8000))
    en_red = EN_RED = "--red" in sys.argv
    if "--nuevo-admin" in sys.argv:                  # RESTABLECER_ADMIN.bat
        _crear_admin_consola(forzar=True)
        sys.exit(0)
    if en_red:
        LOGIN_ACTIVO = True
        _crear_admin_consola()
        USAR_HTTPS = "--sin-https" not in sys.argv
    import socket
    _s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    if os.name != "nt":           # Linux/Mac: que una conexión recién cerrada no parezca "puerto en uso"
        _s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    esquema = "https" if USAR_HTTPS else "http"
    try:
        try:                                          # ¿ya hay alguien atendiendo en ese puerto?
            socket.create_connection(("127.0.0.1", puerto), timeout=1).close()
            raise OSError("en uso")
        except (ConnectionRefusedError, socket.timeout):
            pass
        _s.bind(("0.0.0.0" if en_red else "127.0.0.1", puerto))
    except OSError:
        print(f"\n  El puerto {puerto} ya está en uso: el programa YA ESTÁ ABIERTO en otra ventana negra")
        print("  (o arrancó solo con el inicio automático). No hace falta abrirlo de nuevo.")
        print(f"  Entrá con el navegador a {esquema}://localhost:{puerto}")
        print("  Si querés reiniciarlo: cerrá la otra ventana negra (o reiniciá la PC) y volvé a abrir este archivo.")
        if "--sin-navegador" not in sys.argv:
            webbrowser.open(f"{esquema}://localhost:{puerto}")
        sys.exit(1)
    finally:
        _s.close()
    ssl = {}
    if USAR_HTTPS:
        crt, key = _preparar_certificado()
        ssl = {"ssl_certfile": str(crt), "ssl_keyfile": str(key)}
    if en_red:
        _guardar_kit()

    def _retomar():
        with CANDADO:
            restaurar_sesion()
    threading.Thread(target=_retomar, daemon=True).start()
    if "--sin-navegador" not in sys.argv:
        threading.Timer(1.5, lambda: webbrowser.open(f"{esquema}://localhost:{puerto}")).start()
    print(f"Certificaciones de biomasa → {esquema}://localhost:{puerto}   (para cerrar: cerrá esta ventana)")
    if en_red:                                   # direcciones para pasarle a las otras PC (no sale a internet)
        import socket
        n_us = len(_leer_usuarios())
        print(f"🔒 Abierto en la red: cada uno entra con su usuario y contraseña ({n_us} usuario(s) cargado(s)).")
        print("   Conexión cifrada (HTTPS)." if USAR_HTTPS else "   ⚠ Sin cifrar (--sin-https).")
        print("Desde las otras PC entrar con:")
        print(f"   {esquema}://{socket.gethostname()}:{puerto}")
        for ip in _ips_locales():
            print(f"   {esquema}://{ip}:{puerto}")
        print(f"Acceso directo para las otras PC: carpeta PARA_LAS_OTRAS_PC (INSTALAR_ACCESO_CERTIFICACIONES.bat)")
        print(f"   o desde la otra PC: {esquema}://{socket.gethostname()}:{puerto}/instalar.bat")
    else:
        print("🔒 Solo accesible desde esta PC.")
    uvicorn.run(app_web, host="0.0.0.0" if en_red else "127.0.0.1", port=puerto, log_level="warning",
                limit_concurrency=100, timeout_keep_alive=30, **ssl)
