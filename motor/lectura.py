"""Lectura y preparación del maestro de balanza; reconocimiento de archivos subidos.

Parte del motor de certificaciones (antes en cert_biomasa.py). El código de cada función es el mismo de
la versión 2.1; solo se repartió en módulos."""
import pandas as pd
from openpyxl import load_workbook

from .base import (
    C_BATEA, C_DIA, C_LLENO, C_MOV, C_NETO, C_PATENTE, C_PROV, C_REM_PLANTA, C_REM_PROV, C_TIPO,
    C_VACIO, TIPOLOGIAS)
from .formato import _canon_tipo, _limpio, _patente, _txt


def leer_maestro(ruta, hoja="Datos Entrada"):
    """Lee la hoja de remitos (.xlsb rápido, o .xlsx)."""
    ruta = str(ruta)
    if ruta.lower().endswith(".xlsb"):
        from pyxlsb import open_workbook
        filas, header, vacias = [], None, 0
        with open_workbook(ruta) as wb, wb.get_sheet(hoja) as sh:
            for r in sh.rows():
                v = [c.v for c in r]
                if header is None:
                    if C_DIA in v and C_NETO in v:
                        header = [h if h not in (None, "") else f"_col{j}" for j, h in enumerate(v)]
                    continue
                if v[0] in (None, "") and all(x in (None, "", 0, 0.0) for x in v[3:6]):
                    vacias += 1
                    if vacias > 3000:
                        break
                    continue
                vacias = 0
                filas.append((v + [None] * len(header))[:len(header)])
        df = pd.DataFrame(filas, columns=header)
        df[C_DIA] = pd.to_datetime(pd.to_numeric(df[C_DIA], errors="coerce"),
                                   unit="D", origin="1899-12-30")
    else:
        raw = pd.read_excel(ruta, sheet_name=hoja, header=None, nrows=10)
        fila_h = next(i for i, r in raw.iterrows() if C_DIA in r.values and C_NETO in r.values)
        df = pd.read_excel(ruta, sheet_name=hoja, header=fila_h)
        df[C_DIA] = pd.to_datetime(df[C_DIA], errors="coerce")
    df[C_DIA] = df[C_DIA].dt.normalize()
    return df


def preparar(df, conf=None):
    """Ingresos certificables, normalizados. Cada fila recibe un _id estable
    (N° remito de planta + n° de aparición) para registrar correcciones."""
    d = df[_limpio(df[C_MOV]).str.lower() == "ingreso"].copy()
    d = d.dropna(subset=[C_DIA])
    d[C_PROV] = _limpio(d[C_PROV])
    for c in (C_NETO, C_LLENO, C_VACIO):
        d[c] = pd.to_numeric(d[c], errors="coerce")
    d[C_NETO] = d[C_NETO].fillna(0)
    for c in (C_REM_PROV, C_REM_PLANTA, C_BATEA):
        d[c] = _limpio(d[c])
    tipo_raw = _limpio(d[C_TIPO])
    d["_tipo_valido"] = tipo_raw.str.casefold().isin({t.casefold() for t in TIPOLOGIAS})
    d[C_TIPO] = tipo_raw.map(_canon_tipo)

    clave = d[C_PROV].str.casefold()
    preferido = d.groupby(clave)[C_PROV].agg(lambda s: s.value_counts().index[0])
    unificados = {}
    for k, nombre in preferido.items():
        variantes = sorted(set(d.loc[clave == k, C_PROV]) - {nombre})
        if variantes:
            unificados[nombre] = variantes
    d[C_PROV] = clave.map(preferido).fillna("").astype(str)

    if conf is not None and "UNIFICAR CON" in conf.columns:
        for k, destino in conf["UNIFICAR CON"].items():
            if str(destino).strip():
                m = d[C_PROV].str.casefold() == k.casefold()
                if m.any():
                    d.loc[m, C_PROV] = destino
                    unificados.setdefault(destino, []).append(k)
    d = d[~d[C_PROV].str.casefold().isin(["", "none", "nan"])]
    # nombre con el que el laboratorio y el maestro registraron el camión (para asignar las muestras)
    d["_prov_lab"] = d[C_PROV]
    d = separar_por_patente(d, conf)
    base = d[C_REM_PLANTA].where(d[C_REM_PLANTA] != "", "SINREM")
    d["_id"] = base + "#" + (d.groupby(base).cumcount() + 1).astype(str)
    d["_origen"] = "maestro"
    return d, unificados


def separar_por_patente(d, conf):
    """Regla del config: si un punto de entrega tiene SOLO PATENTES (ej. ROBLESUR (Pérez) = AA123BB), sus camiones
    con otras patentes pasan al proveedor de OTRAS PATENTES A (ej. AGROMAD); si además tiene SI TRANSPORTA (ej.
    AGROMAD), solo los camiones de esa empresa de transporte. El camión sigue mostrando su punto de entrega."""
    if conf is None or "SOLO PATENTES" not in conf.columns or C_PATENTE not in d.columns:
        return d
    emp = d["EMPRESA TRANSPORTE"].map(_txt).str.upper() if "EMPRESA TRANSPORTE" in d.columns else None
    d = d.copy()
    for prov, fila in conf.iterrows():
        solo = {_patente(x) for x in str(fila.get("SOLO PATENTES", "")).replace(";", ",").split(",") if x.strip()}
        destino = str(fila.get("OTRAS PATENTES A", "")).strip()
        if not solo or not destino:
            continue
        m = (d[C_PROV].str.casefold() == str(prov).casefold()) & ~d[C_PATENTE].map(_patente).isin(solo)
        transp = str(fila.get("SI TRANSPORTA", "") or "").strip().upper()
        if transp and emp is not None:
            m &= emp == transp
        d.loc[m, C_PROV] = destino
    return d


def tipo_archivo(ruta):
    """Reconoce un archivo subido por su CONTENIDO (no por el nombre):
    'maestro', 'calidad', 'config', 'modelo', 'estado' o None."""
    r = str(ruta).lower()
    if r.endswith(".json"):
        return "estado"
    if r.endswith(".xlsb"):
        return "maestro"
    if not r.endswith((".xlsx", ".xlsm")):
        return None
    wb = None
    try:
        wb = load_workbook(ruta, read_only=True)       # read_only deja el archivo abierto: se cierra en finally
        hojas = [h.strip().upper() for h in wb.sheetnames]
        if "PROVEEDORES" in hojas and "PRECIOS" in hojas:
            return "config"
        if "DATOS ENTRADA" in hojas:
            return "maestro"
        if "GENERAL" in hojas:
            ws = wb.worksheets[hojas.index("GENERAL")]
            celdas = {str(c).strip() for fila in ws.iter_rows(max_row=15, values_only=True) for c in fila if c}
            if {"PROVEEDOR", "Fecha"} <= celdas and any("Humedad" in c for c in celdas):
                return "calidad"
        texto = " ".join(str(c) for ws in wb.worksheets[:2] for fila in ws.iter_rows(max_row=20, values_only=True)
                         for c in fila if c).upper()
        if "CERTIFICACIÓN ENTREGA BIOMASA" in texto or "CERTIFICACION ENTREGA BIOMASA" in texto:
            return "modelo"
    except Exception:
        return None
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
    return None
