"""Actualiza la instalación que se usa todos los días (Sistema_web) a esta versión, sin tocar los datos.

  ACTUALIZAR_PROGRAMA.bat         → python actualizar.py [carpeta de la instalación actual]
  VOLVER_A_VERSION_ANTERIOR.bat   → python actualizar.py --volver   (se corre desde la instalación ya actualizada)

Qué hace al actualizar:
  1. Controla que el servidor esté cerrado (si está abierto, no hace nada) y compara esta versión con la
     instalada usando tus últimos 6 meses (pruebas/comparar.py). Si da alguna diferencia, no cambia nada.
  2. Copia de seguridad de la carpeta trabajo (respaldos/) y de los archivos del programa viejo
     (versiones_anteriores/programa_vX_FECHA.zip).
  3. Copia SOLO los archivos del programa. NUNCA toca: trabajo (meses, certificaciones, numeración, usuarios,
     config con tus cambios), respaldos, versiones_anteriores, PARA_LAS_OTRAS_PC, ni tus imágenes y config de la raíz.
  4. Corre las pruebas automáticas sobre la instalación actualizada. Si alguna falla, vuelve sola a la versión anterior.
"""
from __future__ import annotations

import re
import shutil
import socket
import subprocess
import sys
import zipfile
from datetime import datetime
from pathlib import Path

NUEVA = Path(__file__).resolve().parent
NO_TOCAR = {"trabajo", "respaldos", "versiones_anteriores", "PARA_LAS_OTRAS_PC", "__pycache__", "python"}
NO_COPIAR = {"actualizar_registro.txt"}
SOLO_SI_FALTA = {"config_proveedores.xlsx"}           # la plantilla de la raíz: si ya hay una, es la tuya
CARPETAS_PROGRAMA = {"motor", "pruebas", "imagenes"}
LINEAS: list[str] = []
REGISTRO_EN = [NUEVA]                                  # dónde se anota lo que pasó


def decir(t=""):
    print(t)
    LINEAS.append(t)


def version_de(carpeta: Path) -> str:
    try:
        m = re.search(r'version="([^"]+)"', (carpeta / "servidor.py").read_text(encoding="utf-8"))
        return m.group(1) if m else "?"
    except OSError:
        return "?"


def servidor_abierto() -> bool:
    for puerto in (8000, 8100):
        try:
            with socket.create_connection(("127.0.0.1", puerto), timeout=1):
                return True
        except OSError:
            pass
    return False


def buscar_actual() -> Path:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if args:
        return Path(args[0].strip('"'))
    for c in (NUEVA.parent / "Sistema_web", Path(r"C:\Certificaciones\Sistema_web"),
              Path.home() / "Desktop" / "Sistema_web", Path.home() / "Escritorio" / "Sistema_web"):
        if (c / "servidor.py").exists() and (c / "trabajo").is_dir() and c.resolve() != NUEVA:
            return c
    r = input("Carpeta de la instalación actual (la que tiene servidor.py y la carpeta trabajo): ").strip().strip('"')
    return Path(r)


def archivos_programa(carpeta: Path):
    """Archivos del programa (relativos), sin datos ni resultados de pruebas."""
    for p in sorted(carpeta.rglob("*")):
        rel = p.relative_to(carpeta)
        if not p.is_file() or rel.parts[0] in NO_TOCAR or "__pycache__" in rel.parts:
            continue
        if len(rel.parts) > 1 and rel.parts[0] not in CARPETAS_PROGRAMA:
            continue                                          # otras carpetas del usuario: no son del programa
        if rel.parts[0] == "pruebas" and p.suffix != ".py":
            continue                                          # resultados de pruebas viejas
        if p.name.startswith("~") or p.suffix in (".tmp", ".log"):
            continue
        yield rel


def guardar_programa_viejo(actual: Path) -> Path:
    carpeta = actual / "versiones_anteriores"
    carpeta.mkdir(exist_ok=True)
    zipf = carpeta / f"programa_v{version_de(actual)}_{datetime.now():%Y-%m-%d_%H%M%S}.zip"
    with zipfile.ZipFile(zipf, "w", zipfile.ZIP_DEFLATED) as z:
        for rel in archivos_programa(actual):
            z.write(actual / rel, rel.as_posix())
    with zipfile.ZipFile(zipf) as z:
        if z.testzip() is not None:
            raise OSError("la copia de la versión anterior quedó dañada")
    return zipf


def restaurar(actual: Path, zipf: Path) -> None:
    with zipfile.ZipFile(zipf) as z:
        nombres = set(z.namelist())
        for rel in list(archivos_programa(actual)):           # lo que la versión nueva agregó y la vieja no tenía
            if rel.as_posix() not in nombres and rel.parts[0] != "imagenes" and rel.name not in SOLO_SI_FALTA:
                (actual / rel).unlink(missing_ok=True)
        z.extractall(actual)
    motor = actual / "motor"
    if motor.is_dir() and not any(motor.glob("*.py")):
        shutil.rmtree(motor, ignore_errors=True)
    limpiar_cache(actual)


def limpiar_cache(carpeta: Path) -> None:
    for c in carpeta.rglob("__pycache__"):
        if c.relative_to(carpeta).parts[0] not in NO_TOCAR:
            shutil.rmtree(c, ignore_errors=True)


def correr_pruebas(actual: Path) -> bool:
    decir("Corriendo las pruebas automáticas sobre la instalación actualizada (1 a 3 minutos)...")
    r = subprocess.run([sys.executable, str(actual / "pruebas" / "probar.py")], cwd=actual)
    return r.returncode == 0


def guardar_registro(carpeta: Path) -> None:
    try:
        with open(carpeta / "actualizar_registro.txt", "a", encoding="utf-8") as f:
            f.write("\n".join(LINEAS) + "\n\n")
    except OSError:
        pass


def actualizar() -> int:
    actual = buscar_actual().resolve()
    if (actual / "servidor.py").exists():
        REGISTRO_EN[0] = actual
    decir(f"ACTUALIZACIÓN · {datetime.now():%d/%m/%Y %H:%M}")
    decir(f"  instalación actual: {actual}  (versión {version_de(actual)})")
    decir(f"  versión nueva:      {NUEVA}  (versión {version_de(NUEVA)})")
    if actual == NUEVA:
        decir("ERROR: esta carpeta ES la versión nueva. Indicá la carpeta de la instalación que usás (Sistema_web).")
        return 2
    if not (actual / "servidor.py").exists() or not (actual / "trabajo").is_dir():
        decir("ERROR: en esa carpeta no está la instalación actual (falta servidor.py o la carpeta trabajo).")
        return 2
    if servidor_abierto():
        decir("ERROR: el programa está abierto. Cerrá la ventana negra del servidor y volvé a ejecutar esto.")
        return 2
    if input("\n¿Actualizar? Tus datos (carpeta trabajo) no se tocan. Escribí SI y Enter: ").strip().upper() != "SI":
        decir("No se cambió nada.")
        return 1

    if "--sin-comparar" not in sys.argv:
        decir("\nPaso 1: comparo esta versión con la tuya, con tus últimos 6 meses (unos minutos; no cambia nada)...")
        r = subprocess.run([sys.executable, str(NUEVA / "pruebas" / "comparar.py"), str(actual)], cwd=NUEVA)
        if r.returncode == 1:
            decir("\nLA COMPARACIÓN DIO DIFERENCIAS: no se actualizó nada. Mandá a soporte el archivo")
            decir(f"pruebas\\comparacion_....txt de {NUEVA.name}.")
            return 1
        if r.returncode != 0:
            decir("  (no se pudo comparar: la instalación todavía no tiene subidos los archivos; sigo con las pruebas)")
        decir("\nPaso 2: actualizo.")

    sys.path.insert(0, str(NUEVA))
    import respaldo
    info = respaldo.Respaldo(actual / "trabajo", actual / "respaldos").hacer()
    decir(f"  ✔ copia de seguridad de los datos: respaldos\\{info['archivo']} ({info['mb']} MB)")
    viejo = guardar_programa_viejo(actual)
    decir(f"  ✔ copia de la versión anterior: versiones_anteriores\\{viejo.name}")

    copiados = 0
    try:
        for rel in archivos_programa(NUEVA):
            if rel.name in NO_COPIAR or rel.name == "LEEME_RESPALDO.txt":
                continue
            destino = actual / rel
            if (rel.name in SOLO_SI_FALTA or rel.parts[0] == "imagenes") and destino.exists():
                continue                                      # tu config de la raíz y tus imágenes quedan como están
            destino.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(NUEVA / rel, destino)
            copiados += 1
        limpiar_cache(actual)
        if (NUEVA / "python" / "python.exe").exists() and not (actual / "python" / "python.exe").exists():
            decir("  copiando el Python portátil (unos 190 MB, 1 a 3 minutos)...")
            shutil.copytree(NUEVA / "python", actual / "python", dirs_exist_ok=True)
            decir("  ✔ Python portátil copiado: el programa ya no depende del Python instalado en la PC")
    except Exception as e:
        decir(f"ERROR copiando archivos ({e}). Vuelvo a la versión anterior...")
        restaurar(actual, viejo)
        decir("  ✔ quedó la versión anterior, como estaba.")
        return 1
    decir(f"  ✔ {copiados} archivos del programa actualizados")

    if not correr_pruebas(actual):
        decir("\nALGUNA PRUEBA FALLÓ. Vuelvo a la versión anterior para que puedas seguir trabajando...")
        restaurar(actual, viejo)
        decir(f"  ✔ quedó la versión {version_de(actual)}, como estaba. Mandá a soporte el último archivo")
        decir("    pruebas\\resultado_....txt de la carpeta del programa.")
        return 1
    decir("")
    decir("=" * 70)
    decir(f"LISTO ✔  La instalación quedó en la versión {version_de(actual)} y todas las pruebas pasaron.")
    decir("Abrí el programa como siempre (INICIAR_WEB_RED.bat o el inicio automático).")
    decir("Si algo no te convence: VOLVER_A_VERSION_ANTERIOR.bat en la carpeta del programa.")
    return 0


def volver() -> int:
    actual = NUEVA                                            # se corre desde la instalación actualizada
    anteriores = sorted((actual / "versiones_anteriores").glob("programa_v*.zip"), key=lambda p: p.stat().st_mtime)
    decir(f"VOLVER A LA VERSIÓN ANTERIOR · {datetime.now():%d/%m/%Y %H:%M}")
    if not anteriores:
        decir("No hay ninguna versión anterior guardada en versiones_anteriores.")
        return 2
    zipf = anteriores[-1]
    decir(f"  versión instalada ahora: {version_de(actual)}")
    decir(f"  se vuelve a: {zipf.name}")
    if servidor_abierto():
        decir("ERROR: el programa está abierto. Cerrá la ventana negra del servidor y volvé a ejecutar esto.")
        return 2
    if input("\n¿Volver? Tus datos (carpeta trabajo) no se tocan. Escribí SI y Enter: ").strip().upper() != "SI":
        decir("No se cambió nada.")
        return 1
    restaurar(actual, zipf)
    decir(f"LISTO ✔  Quedó la versión {version_de(actual)}. Abrí el programa como siempre.")
    return 0


if __name__ == "__main__":
    try:
        codigo = volver() if "--volver" in sys.argv else actualizar()
    except Exception as e:
        decir(f"ERROR inesperado: {e}")
        codigo = 1
    guardar_registro(REGISTRO_EN[0])
    sys.exit(codigo)
