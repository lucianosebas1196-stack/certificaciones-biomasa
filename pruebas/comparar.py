"""COMPARAR_CON_ACTUAL.bat — ¿la versión nueva da EXACTAMENTE lo mismo que la que estás usando?

Arma los mismos meses con las dos versiones del programa, con TUS últimos archivos (los que subiste en la versión
actual), y compara certificación por certificación: importe, toneladas, descuentos, precios, observaciones, estado y
el texto del PDF. No toca nada de ninguna de las dos instalaciones: trabaja sobre copias temporales.

uso: python pruebas/comparar.py [carpeta de la versión actual]     (por defecto: C:\\Certificaciones\\Sistema_web
                                                                   o la carpeta Sistema_web que esté al lado)
"""
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path

NUEVA = Path(__file__).resolve().parent.parent
LINEAS = []


def decir(t=""):
    print(t, flush=True)
    LINEAS.append(t)


# ------------------------------------------------------------------ lo que corre DENTRO de cada versión
FOTO = r'''
import sys, os, json, hashlib, warnings, subprocess
warnings.filterwarnings("ignore")
sys.path.insert(0, os.getcwd())
import servidor as sv, cert_biomasa as cb
from starlette.testclient import TestClient
archivos, meses, salida = json.loads(sys.argv[1]), json.loads(sys.argv[2]), sys.argv[3]
h = lambda x: hashlib.md5(json.dumps(x, default=str, sort_keys=True).encode()).hexdigest()
def redondo(v):
    return round(v, 8) if isinstance(v, float) else v
def firma(c):
    f = c["filas"]
    cols = [x for x in f.columns if not x.startswith("_") or x == "_id"]
    return {"importe": c["importe"], "tn": c["tn"], "estado": c["estado"], "moneda": c.get("moneda", "ARS"),
            "precios": sorted((k, v) for k, v in c["precios"].items()),
            "camiones": h(f[cols].astype(str).values.tolist()),
            "obs": sorted((o["sev"], o["msg"], str(o["rem"])) for o in c["obs"]),
            "tramos": [(t["etiqueta"], t["tn"], t["precio"]) for t in c["tramos"]],
            "humedad": {k: {kk: redondo(vv) for kk, vv in v.items()} for k, v in c.get("humedad", {}).items()},
            "huella": cb.huella(c), "razon": c["razon"], "cuit": c["cuit"]}
cl = TestClient(sv.app_web, base_url="http://localhost", client=("127.0.0.1", 5))
fs = [("archivos", (os.path.basename(p) if t != "config" else "config_proveedores.xlsx", open(p, "rb"))) for t, p in archivos.items()]
r = cl.post("/api/archivos", files=fs)
assert r.status_code == 200, r.text[:300]
out = {}
for anio, mes in meses:
    g = cl.post("/api/armar", json={"anio": anio, "mes": mes})
    if g.status_code != 200:
        out[f"{anio}-{mes:02d}"] = {"error": g.text[:300]}
        continue
    a = sv.S.apps[(anio, mes)]
    certs = {k: firma(c) for k, c in a.certs.items()}
    pdfs = {}
    for k in sorted(a.certs)[:8]:
        c = dict(a.certs[k], numero="0001-00009999")
        ruta = os.path.join(os.getcwd(), "p.pdf")
        cb.pdf_certificacion(ruta, c, sv.S.logo)
        try:
            txt = subprocess.run(["pdftotext", "-layout", ruta, "-"], capture_output=True, text=True).stdout
        except OSError:
            from pypdf import PdfReader
            txt = "\n".join(p.extract_text() or "" for p in PdfReader(ruta).pages)
        pdfs[k] = h(txt)
    out[f"{anio}-{mes:02d}"] = {"certs": certs, "cuadre": g.json()["cuadre"], "pdf": pdfs}
    print(f"  {anio}-{mes:02d}: {len(certs)} certificaciones", flush=True)
json.dump(out, open(salida, "w"), default=str)
'''

PROGRAMA = ["servidor.py", "cert_biomasa.py", "editor_colab.py", "almacen_db.py", "lista_blanca.py",
            "guardado_config.py", "respaldo.py", "index.html", "login.html"]


def copia(origen: Path, destino: Path):
    destino.mkdir(parents=True)
    for f in PROGRAMA:
        if (origen / f).exists():
            shutil.copy2(origen / f, destino / f)
    for carpeta in ("motor", "imagenes"):
        if (origen / carpeta).exists():
            shutil.copytree(origen / carpeta, destino / carpeta, ignore=shutil.ignore_patterns("__pycache__"))


def archivos_de(actual: Path):
    """Los últimos maestro / laboratorio / config de la instalación actual."""
    t = actual / "trabajo"
    res = {}
    try:
        ses = json.loads((t / "sesion.json").read_text(encoding="utf-8"))
        for tipo in ("maestro", "calidad"):
            if ses.get(tipo) and Path(ses[tipo]).exists():
                res[tipo] = str(ses[tipo])
    except (OSError, ValueError):
        pass
    for c in (t / "config_proveedores_PENDIENTE.xlsx", t / "config_proveedores.xlsx", actual / "config_proveedores.xlsx"):
        if c.exists():
            res["config"] = str(c)
            break
    return res


def buscar_actual():
    if len(sys.argv) > 1:
        return Path(sys.argv[1])
    for c in (NUEVA.parent / "Sistema_web", Path(r"C:\Certificaciones\Sistema_web"), Path.home() / "Desktop" / "Sistema_web"):
        if (c / "servidor.py").exists() and c.resolve() != NUEVA.resolve():
            return c
    r = input("Carpeta de la versión actual (la que tiene servidor.py y la carpeta trabajo): ").strip().strip('"')
    return Path(r)


def main():
    actual = buscar_actual()
    decir(f"COMPARACIÓN · {datetime.now():%d/%m/%Y %H:%M}")
    decir(f"  versión actual: {actual}")
    decir(f"  versión nueva:  {NUEVA}")
    if not (actual / "servidor.py").exists():
        decir("ERROR: en esa carpeta no está el programa actual.")
        return 2
    arch = archivos_de(actual)
    if not {"maestro", "calidad", "config"} <= set(arch):
        decir("ERROR: la versión actual todavía no tiene subidos el maestro, la planilla del laboratorio y el config.")
        return 2
    for t, p in arch.items():
        decir(f"  {t}: {Path(p).name}")
    tmp = Path(tempfile.mkdtemp(prefix="comparar_cert_"))
    try:
        # meses a comparar: los últimos 6 del maestro (se leen con la versión actual)
        copia(actual, tmp / "actual")
        copia(NUEVA, tmp / "nueva")
        sondeo = ("import sys,os,json;sys.path.insert(0,os.getcwd());import cert_biomasa as cb;"
                  "d=cb.preparar(cb.leer_maestro(sys.argv[1]))[0];"
                  "m=sorted({(int(x.year),int(x.month)) for x in d[cb.C_DIA].dropna()});print(json.dumps(m[-6:]))")
        r = subprocess.run([sys.executable, "-c", sondeo, arch["maestro"]], cwd=tmp / "actual", capture_output=True, text=True)
        meses = json.loads(r.stdout.strip().splitlines()[-1])
        decir(f"  meses: {', '.join(f'{m:02d}/{a}' for a, m in meses)}")
        fotos = {}
        for nombre in ("actual", "nueva"):
            decir(f"\nArmando con la versión {nombre}…")
            salida = tmp / f"{nombre}.json"
            r = subprocess.run([sys.executable, "-c", FOTO, json.dumps(arch), json.dumps(meses), str(salida)],
                               cwd=tmp / nombre, text=True)
            if r.returncode != 0 or not salida.exists():
                decir(f"ERROR: la versión {nombre} no pudo armar los meses.")
                return 1
            fotos[nombre] = json.loads(salida.read_text())
        decir("")
        difs, total = [], 0
        for mes, a in fotos["actual"].items():
            n = fotos["nueva"].get(mes, {})
            if "error" in a or "error" in n:
                if a.get("error") != n.get("error"):
                    difs.append(f"{mes}: error distinto ({a.get('error')} / {n.get('error')})")
                continue
            if a["cuadre"] != n["cuadre"]:
                difs.append(f"{mes}: el cuadre cambió")
            for k in sorted(set(a["certs"]) | set(n["certs"])):
                total += 1
                ca, cn = a["certs"].get(k), n["certs"].get(k)
                if ca is None or cn is None:
                    difs.append(f"{mes} {k}: está en una sola versión")
                    continue
                for campo in ca:
                    if ca[campo] != cn.get(campo):
                        difs.append(f"{mes} {k}: cambió {campo} ({str(ca[campo])[:60]} → {str(cn.get(campo))[:60]})")
            for k in a["pdf"]:
                if a["pdf"][k] != n["pdf"].get(k):
                    difs.append(f"{mes} {k}: el texto del PDF cambió")
        decir("=" * 70)
        if difs:
            decir(f"RESULTADO: HAY {len(difs)} DIFERENCIA(S). No reemplaces la versión actual; mandá este archivo a soporte.")
            for d in difs[:200]:
                decir("  - " + d)
        else:
            decir(f"RESULTADO: IDÉNTICAS ✔  {total} certificaciones de {len(fotos['actual'])} meses dan exactamente lo mismo")
            decir("(importes, toneladas, descuentos, precios, observaciones, estados, cuadre y texto de los PDF).")
        return 1 if difs else 0
    finally:
        try:
            out = NUEVA / "pruebas" / f"comparacion_{datetime.now():%Y-%m-%d_%H%M}.txt"
            out.write_text("\n".join(LINEAS), encoding="utf-8")
            print(f"Quedó guardado en pruebas\\{out.name}")
        except OSError:
            pass
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
