# =============================================================================
#  MOTOR — Certificaciones de entrega de biomasa (quincenal y mensual)
#  Lee el maestro de control de ingreso de biomasa, hoja "Datos Entrada".
#
#  - Las correcciones manuales se guardan POR REMITO (no por certificación), así
#    1ª quincena, 2ª quincena y mensual siempre quedan iguales entre sí.
#  - Cada cambio queda en un registro: quién, cuándo, valor anterior, nuevo, motivo.
#  - Una certificación solo se confirma (y genera PDF) si no tiene errores
#    bloqueantes y las observaciones críticas fueron revisadas con comentario.
# =============================================================================
import io
import re
import json
import zipfile
import hashlib
import calendar
import difflib
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.page import PageMargins

MESES = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
         "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

EMPRESA_LINEAS = [
    "Planta de Biomasa Ejemplo S.A. · Parque Industrial, Ruta 0 km 0",
    "Provincia, Argentina.",
    "C.U.I.T 30-00000001-5",
]
PIE = [
    "*Ante disconformidad tiene 3 días hábiles para realizar el reclamo pertinente. "
    "Pasado este período se tomará por aceptada.",
    "**Remitir la factura junto a la orden de compra a compras@empresa-ejemplo.com "
    "y administracion@empresa-ejemplo.com",
    "Importante: la factura debe indicar el número de orden de compra.",
]

TIPOLOGIAS = ["CHIP", "ASERRIN", "Fino/Raleo", "C LARGO", "C CORTO", "PATAS",
              "CENIZAS", "DESECHOS", "RAMAS"]
MOVIMIENTOS = {"ingreso", "mi", "me"}
PESO_MIN, PESO_MAX = 3000, 45000
DIA_CORTE = 12          # último día de la 1ª quincena (DIA_CORTE_QUINCENA en el config)
# Granulometría (como la planilla de certificación final): finos (< #04) por camión
#   hasta FINOS_TOLERANCIA %           → sin descuento
#   de FINOS_TOLERANCIA a FINOS_TOTALIDAD → se descuenta la diferencia
#   más de FINOS_TOTALIDAD %           → se descuenta la totalidad de los finos
# Lo descontado no se pierde: se paga como ASERRÍN.
FINOS_TOLERANCIA, FINOS_TOTALIDAD = 10.0, 15.0
NOMBRE_PRODUCTO = {"CHIP": "CHIP", "ASERRIN": "ASERRIN", "C CORTO": "C. CORTO", "C LARGO": "C. LARGO",
                   "Fino/Raleo": "FINO/RALEO", "PATAS": "PATAS", "CENIZAS": "CENIZAS", "DESECHOS": "DESECHOS",
                   "RAMAS": "RAMAS"}
PRODUCTOS_CUADRO = ["CHIP", "ASERRIN", "C CORTO", "C LARGO", "Fino/Raleo"]          # rango habitual por camión (kg)
SIN_REMITO = ("", "S/R", "SR", "0", "-")

# Columnas del maestro
C_DIA, C_REM_PLANTA, C_PROV = "DIA", "Nº Remito Planta", "PROVEEDOR NOMBRE"
C_REM_PROV, C_TIPO, C_BATEA = "PROVEEDOR N°REMITO", "TIPOLOGIA", "P BATEA"
C_NETO, C_MOV, C_QUINC = "PESO NETO", "TIPO DE MOVIMINETO", "Quincena"
C_LLENO, C_VACIO = "PESO LLENO", "PESO VACIO"
C_PATENTE = "PATENTE TRANSPORTE"

# Campos editables de un remito → columna del maestro
CAMPOS = {"Batea": C_BATEA, "Fecha": C_DIA, "Rto. Planta": C_REM_PLANTA,
          "Remito Proveedor": C_REM_PROV, "Producto": C_TIPO, "Kilos": C_NETO,
          "Proveedor": C_PROV}
FILA_COLS = ["Batea", "Fecha", "Rto. Planta", "Remito Proveedor", "Producto", "Kilos"]

CRIT, ADV = "CRÍTICO", "ADVERTENCIA"


# -----------------------------------------------------------------------------
# 1. Lectura y preparación
# -----------------------------------------------------------------------------
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


def _txt(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    if isinstance(x, float) and x.is_integer():
        return str(int(x))
    return str(x).strip()


def _limpio(serie):
    return serie.apply(_txt).str.replace(r"\s+", " ", regex=True).str.strip()


def _canon_tipo(t):
    canon = {x.casefold(): x for x in TIPOLOGIAS}
    t = _txt(t)
    return canon.get(t.casefold(), t.upper() or "SIN TIPO")


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


def _patente(s):
    return re.sub(r"[^A-Z0-9]", "", str(s).upper())


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


# -----------------------------------------------------------------------------
# 2. Configuración (proveedores, precios, numeración)
# -----------------------------------------------------------------------------
def plantilla_config(d):
    provs = sorted(d[C_PROV].unique(), key=str.casefold)
    conf = pd.DataFrame({"PROVEEDOR (maestro)": provs, "RAZÓN SOCIAL": provs, "CUIT": "",
                         "GRUPO": provs, "CERTIFICAR": "SI", "UNIFICAR CON": "",
                         "NOMBRE HUMEDAD": "", "DESCUENTA HUMEDAD": "NO", "TOLERANCIA HUMEDAD %": "",
                         "LIMITE HUMEDAD %": "", "DESCUENTA GRANULOMETRIA": "SI", "TOLERANCIA FINOS %": "",
                         "CERTIFICAR CON": "", "FORMATO": "", "SOLO PATENTES": "", "OTRAS PATENTES A": "",
                         "SI TRANSPORTA": "", "FRECUENCIA": ""})
    clar = conf["PROVEEDOR (maestro)"].str.upper().str.startswith("ROBLESUR")
    conf.loc[clar, ["SOLO PATENTES", "OTRAS PATENTES A", "SI TRANSPORTA"]] = ["AA123BB", "AGROMAD", "AGROMAD"]
    agromad = conf["PROVEEDOR (maestro)"].str.upper().str.startswith("AGROMAD")
    conf.loc[agromad, ["RAZÓN SOCIAL", "GRUPO", "CERTIFICAR CON"]] = ["AGROMAD", "AGROMAD", "AGROMAD"]
    conf.loc[conf["PROVEEDOR (maestro)"].str.upper() == "PLANTA PROPIA", "CERTIFICAR"] = "NO"   # movimientos internos
    mym = conf["PROVEEDOR (maestro)"].str.upper().str.replace(" ", "").str.startswith("MYM")
    conf.loc[mym, ["FORMATO", "DESCUENTA HUMEDAD", "TOLERANCIA FINOS %"]] = ["MYM", "SI", "5"]
    tp = conf["PROVEEDOR (maestro)"].str.startswith("Pallets Sur")
    conf.loc[tp, ["RAZÓN SOCIAL", "CUIT", "GRUPO"]] = ["Pallets Sur", "30-00000002-3", "Pallets Sur"]
    pinar = conf["PROVEEDOR (maestro)"].str.upper().str.startswith("PINAR MADERAS")
    conf.loc[pinar, ["RAZÓN SOCIAL", "CUIT", "GRUPO", "CERTIFICAR CON"]] = ["Pinar Maderas SRL", "30-00000003-1",
                                                                           "Pinar Maderas", "Pinar Maderas"]
    precios = pd.DataFrame([["Pallets Sur", "CHIP", 0, 1000, 10000],
                            ["Pallets Sur", "CHIP", 1000, 2000, 10500],
                            ["Pallets Sur", "CHIP", 2000, None, 11000],
                            ["Pallets Sur", "ASERRIN", 0, None, 1000]],
                           columns=["GRUPO", "TIPOLOGIA", "DESDE_TN", "HASTA_TN", "PRECIO"])
    return conf.set_index("PROVEEDOR (maestro)"), precios, {"NUMERO_INICIAL": 1, "PREFIJO": "0001", "METODO_HUMEDAD": "PUNTOS",
                                                          "DIA_CORTE_QUINCENA": 12, "FINOS_TOLERANCIA": 10.0,
                                                          "FINOS_TOTALIDAD": 15.0}


def _porcentajes_excel(ruta, conf):
    """Si en el config una tolerancia se escribió con formato de porcentaje de Excel (se ve 15 %, pero la celda
    guarda 0,15), se pasa a 15 para que se lea igual que si se hubiera escrito 15."""
    try:
        from openpyxl import load_workbook
        ws = load_workbook(ruta, read_only=True)["Proveedores"]
    except Exception:
        return
    filas = ws.iter_rows()
    enc = [str(c.value or "").strip() for c in next(filas)]
    cols = {i: n for i, n in enumerate(enc) if n in ("TOLERANCIA HUMEDAD %", "LIMITE HUMEDAD %", "TOLERANCIA FINOS %",
                                                     "LIMITE FINOS %")}
    for j, fila in enumerate(filas):
        if j >= len(conf):
            break
        for i, n in cols.items():
            if i < len(fila):
                c = fila[i]
                if isinstance(c.value, (int, float)) and "%" in str(c.number_format or ""):
                    conf.iat[j, conf.columns.get_loc(n)] = f"{c.value * 100:g}"


def leer_config(ruta):
    conf = pd.read_excel(ruta, sheet_name="Proveedores", dtype=str).fillna("")
    conf.columns = [c.strip() for c in conf.columns]
    _porcentajes_excel(ruta, conf)
    for col in ("UNIFICAR CON", "NOMBRE HUMEDAD", "DESCUENTA HUMEDAD", "TOLERANCIA HUMEDAD %", "LIMITE HUMEDAD %",
                "DESCUENTA GRANULOMETRIA", "TOLERANCIA FINOS %", "LIMITE FINOS %", "CERTIFICAR CON", "FORMATO",
                "SOLO PATENTES", "OTRAS PATENTES A", "SI TRANSPORTA", "FRECUENCIA"):
        if col not in conf.columns:
            conf[col] = ""
    conf["PROVEEDOR (maestro)"] = conf["PROVEEDOR (maestro)"].str.strip().str.replace(r"\s+", " ", regex=True)
    conf = conf.drop_duplicates("PROVEEDOR (maestro)").set_index("PROVEEDOR (maestro)")
    for c in conf.columns:
        conf[c] = conf[c].str.strip()
    precios = pd.read_excel(ruta, sheet_name="Precios")
    precios["GRUPO"] = precios["GRUPO"].astype(str).str.strip()
    precios["TIPOLOGIA"] = precios["TIPOLOGIA"].map(_canon_tipo)
    for c in ("DESDE_TN", "HASTA_TN", "PRECIO"):
        precios[c] = pd.to_numeric(precios[c], errors="coerce")
    precios = precios.dropna(subset=["GRUPO"]).reset_index(drop=True)
    precios = _columnas_punto(precios)
    try:
        par = pd.read_excel(ruta, sheet_name="Parametros").set_index("PARÁMETRO")["VALOR"].to_dict()
    except ValueError:
        par = {}
    try:                                        # celda vacía o con texto → 1 (el programa la sube sola si hace falta)
        par["NUMERO_INICIAL"] = max(1, int(float(str(par.get("NUMERO_INICIAL")).replace(",", "."))))
    except (TypeError, ValueError):
        par["NUMERO_INICIAL"] = 1
    pref = par.get("PREFIJO")
    pref = "" if pref is None or (isinstance(pref, float) and pd.isna(pref)) else str(pref).strip()
    if pref.endswith(".0"):
        pref = pref[:-2]
    par["PREFIJO"] = pref.zfill(4) if pref.isdigit() else (pref or "0001")
    par["METODO_HUMEDAD"] = str(par.get("METODO_HUMEDAD") or "PUNTOS").strip().upper()
    if par["METODO_HUMEDAD"] not in METODOS_HUMEDAD:
        par["METODO_HUMEDAD"] = "PUNTOS"
    try:
        par["DIA_CORTE_QUINCENA"] = int(float(par.get("DIA_CORTE_QUINCENA") or 12))
    except (TypeError, ValueError):
        par["DIA_CORTE_QUINCENA"] = 12
    if not 10 <= par["DIA_CORTE_QUINCENA"] <= 20:
        par["DIA_CORTE_QUINCENA"] = 12
    fijar_corte(par["DIA_CORTE_QUINCENA"])
    for k, defecto in (("FINOS_TOLERANCIA", 10.0), ("FINOS_TOTALIDAD", 15.0)):
        try:
            par[k] = float(str(par.get(k) if par.get(k) is not None else defecto).replace(",", ".").replace("%", ""))
        except ValueError:
            par[k] = defecto
    fijar_finos(par["FINOS_TOLERANCIA"], par["FINOS_TOTALIDAD"])
    return conf, precios, par


def fijar_finos(tolerancia, totalidad):
    global FINOS_TOLERANCIA, FINOS_TOTALIDAD
    FINOS_TOLERANCIA, FINOS_TOTALIDAD = float(tolerancia), float(totalidad)


def fijar_corte(dia):
    """Define el último día de la 1ª quincena (la 2ª empieza el día siguiente)."""
    global DIA_CORTE
    DIA_CORTE = int(dia)


def guardar_config(ruta, conf, precios, par):
    with pd.ExcelWriter(ruta, engine="openpyxl") as w:
        conf.reset_index().to_excel(w, sheet_name="Proveedores", index=False)
        precios.to_excel(w, sheet_name="Precios", index=False)
        pd.DataFrame({"PARÁMETRO": ["NUMERO_INICIAL", "PREFIJO", "METODO_HUMEDAD", "DIA_CORTE_QUINCENA",
                                    "FINOS_TOLERANCIA", "FINOS_TOTALIDAD"],
                      "VALOR": [int(par["NUMERO_INICIAL"]), str(par["PREFIJO"]).zfill(4),
                                str(par.get("METODO_HUMEDAD", "PUNTOS")).upper(),
                                int(par.get("DIA_CORTE_QUINCENA", DIA_CORTE)),
                                float(par.get("FINOS_TOLERANCIA", FINOS_TOLERANCIA)),
                                float(par.get("FINOS_TOTALIDAD", FINOS_TOTALIDAD))],
                      "AYUDA": ["Próximo número de certificación", "Punto de venta",
                                "PUNTOS = TN × puntos sobre el límite · BASE_SECA = TN × (1−H)/(1−L)",
                                "Último día de la 1ª quincena (ej. 12 = del 1 al 12 y del 13 a fin de mes)",
                                "Granulometría: % de finos sin descuento (general; se puede cambiar por proveedor)",
                                "Granulometría: por encima de este % se descuenta la totalidad de los finos"]}
                     ).to_excel(w, sheet_name="Parametros", index=False)
        pd.DataFrame({"Tipologías válidas": TIPOLOGIAS}).to_excel(w, sheet_name="Listas", index=False)
    wb = load_workbook(ruta)
    for ws in wb:
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = max(
                12, min(45, max(len(str(c.value or "")) for c in col) + 2))
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="3A7D44")
        ws.freeze_panes = "A2"
    ws = wb["Precios"]
    dv = DataValidation(type="list", formula1=f"=Listas!$A$2:$A${len(TIPOLOGIAS) + 1}",
                        showErrorMessage=True, errorTitle="Tipología inválida",
                        error="Elegí una tipología de la lista.")
    ws.add_data_validation(dv)
    dv.add("B2:B1000")
    dvp = DataValidation(type="decimal", operator="greaterThan", formula1="0", showErrorMessage=True,
                         errorTitle="Precio inválido", error="El precio debe ser un número mayor a 0.")
    ws.add_data_validation(dvp)
    dvp.add("E2:E1000")
    ws = wb["Proveedores"]
    dvs = DataValidation(type="list", formula1='"SI,NO"', showErrorMessage=True)
    ws.add_data_validation(dvs)
    cols = [c.value for c in ws[1]]
    for nombre in ("CERTIFICAR",):
        if nombre in cols:
            letra = ws.cell(1, cols.index(nombre) + 1).column_letter
            dvs.add(f"{letra}2:{letra}{len(conf) + 300}")
    if "FRECUENCIA" in cols:
        # QUINCENAL = 1ª y 2ª quincena · MENSUAL = solo el mensual · AMBAS o vacío = lo que se marque en el Paso 4
        dvf = DataValidation(type="list", formula1='"QUINCENAL,MENSUAL,AMBAS"', allow_blank=True, showErrorMessage=True,
                             errorTitle="Frecuencia inválida", error="Elegí QUINCENAL, MENSUAL o AMBAS (o dejalo vacío).",
                             showInputMessage=True, promptTitle="Frecuencia de certificación",
                             prompt="QUINCENAL: 1ª y 2ª quincena · MENSUAL: solo el mensual · AMBAS o vacío: todo lo "
                                    "que se marque en el Paso 4")
        ws.add_data_validation(dvf)
        letra = ws.cell(1, cols.index("FRECUENCIA") + 1).column_letter
        dvf.add(f"{letra}2:{letra}{len(conf) + 300}")
        ws[f"{letra}1"].fill = PatternFill("solid", fgColor="2E75B6")
    # Columnas de descuentos de calidad: encabezado naranja y ayuda al pararse en la celda
    ayudas = {
        "TOLERANCIA HUMEDAD %": ("decimal", "Puntos de humedad que se toleran por encima del límite antes de descontar. "
                                 "Ej.: 0 = se descuenta desde el límite; 2 = recién por encima de límite + 2."),
        "LIMITE HUMEDAD %": ("lab", "Humedad de referencia. LAB = la del laboratorio de ese día (planilla de calidad). "
                             "Un número (ej. 45) = límite fijo para este proveedor."),
        "DESCUENTA HUMEDAD": (None, "SI = se descuenta (humedad medida − límite − tolerancia) % de los kilos de CHIP y "
                              "ASERRÍN. NO = no se descuenta humedad (igual se muestra el % en cada fila)."),
        "DESCUENTA GRANULOMETRIA": (None, "SI = los finos de CHIP por encima de la tolerancia pasan a aserrín. "
                                    "NO = sin descuento por granulometría."),
        "TOLERANCIA FINOS %": ("decimal", "% de finos sin descuento. Entre esto y FINOS_TOTALIDAD (Parametros, 15) se "
                               "descuenta la diferencia; por encima, todos los finos. MyM: es el % 'Alto' de su planilla."),
    }
    naranja = PatternFill("solid", fgColor="C55A11")
    for nombre, (tipo, texto) in ayudas.items():
        if nombre not in cols:
            continue
        letra = ws.cell(1, cols.index(nombre) + 1).column_letter
        ws[f"{letra}1"].fill = naranja
        if tipo == "decimal":
            dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="100", allow_blank=True,
                                showErrorMessage=True, errorTitle="Porcentaje inválido",
                                error="Escribí un número de 0 a 100 (ej. 10 para 10 %).")
        elif tipo == "lab":
            dv = DataValidation(type="custom", formula1=f'=OR({letra}2="LAB",AND(ISNUMBER({letra}2),{letra}2>=0,{letra}2<=100))',
                                allow_blank=True, showErrorMessage=True, errorTitle="Límite inválido",
                                error="Escribí LAB o un número de 0 a 100 (ej. 45).")
        else:
            dv = DataValidation(type="list", formula1='"SI,NO"', allow_blank=True, showErrorMessage=True)
        dv.showInputMessage, dv.promptTitle, dv.prompt = True, nombre[:32], texto[:255]
        ws.add_data_validation(dv)
        dv.add(f"{letra}2:{letra}{len(conf) + 300}")
    wb.save(ruta)


def cuit_valido(cuit):
    n = re.sub(r"\D", "", str(cuit))
    if len(n) != 11:
        return False
    s = sum(int(a) * b for a, b in zip(n[:10], [5, 4, 3, 2, 7, 6, 5, 4, 3, 2]))
    v = 11 - s % 11
    v = 0 if v == 11 else v
    return v != 10 and v == int(n[10])


def puntos(v):
    """Lee un porcentaje cargado en el config: '2', '2%', '2,5' → 0.02 / 0.025. Vacío o LAB → None
    (LAB = se usa el límite de referencia de la planilla del laboratorio de ese día)."""
    t = str(v if v is not None else "").replace("%", "").replace(",", ".").strip()
    if t.upper() in ("", "NAN", "NONE", "LAB", "LABORATORIO"):
        return None
    try:
        return float(t) / 100
    except ValueError:
        return "ERROR"


def unidad(conf, prov):
    """Nombre de la certificación a la que pertenece un punto de entrega (columna CERTIFICAR CON)."""
    if prov in conf.index and "CERTIFICAR CON" in conf.columns:
        v = str(conf.at[prov, "CERTIFICAR CON"]).strip()
        if v and v.lower() not in ("nan", "none"):
            return v
    return prov


def miembros(conf, nombre_unidad, presentes=None):
    """Puntos de entrega (nombres del maestro) que forman una certificación."""
    cand = presentes if presentes is not None else conf.index
    return sorted([p for p in cand if unidad(conf, p) == nombre_unidad], key=str.casefold) or [nombre_unidad]


FRECUENCIAS = {"QUINCENAL": ("1Q", "2Q"), "MENSUAL": ("Mensual",), "AMBAS": ("1Q", "2Q", "Mensual"), "": ("1Q", "2Q", "Mensual")}


def frecuencia(conf, nombre_unidad):
    """FRECUENCIA de una certificación (columna del config; si los puntos de entrega difieren, la del primero cargado).
    QUINCENAL → 1Q y 2Q · MENSUAL → solo Mensual · AMBAS o vacío → todos los períodos elegidos."""
    if conf is None or "FRECUENCIA" not in conf.columns:
        return ""
    for p in miembros(conf, nombre_unidad):
        if p in conf.index:
            v = str(conf.at[p, "FRECUENCIA"] or "").strip().upper()
            v = {"QUINCENA": "QUINCENAL", "QUINCENALES": "QUINCENAL", "MES": "MENSUAL", "AMBOS": "AMBAS"}.get(v, v)
            if v in FRECUENCIAS and v:
                return v
    return ""


def corresponde(conf, nombre_unidad, periodo):
    return periodo in FRECUENCIAS[frecuencia(conf, nombre_unidad)]


def info_proveedor(conf, prov):
    if prov in conf.index:
        f = conf.loc[prov]
        return {"razon": f.get("RAZÓN SOCIAL") or prov, "cuit": f.get("CUIT", ""),
                "grupo": f.get("GRUPO") or prov,
                "certifica": str(f.get("CERTIFICAR", "SI")).upper() != "NO",
                "tolerancia": puntos(f.get("TOLERANCIA HUMEDAD %", "")),
                "limite": puntos(f.get("LIMITE HUMEDAD %", "")),
                "desc_hum": str(f.get("DESCUENTA HUMEDAD", "")).strip().upper() == "SI",
                "desc_gran": str(f.get("DESCUENTA GRANULOMETRIA", "")).strip().upper() != "NO",
                "tol_finos": puntos(f.get("TOLERANCIA FINOS %", "")),
                "lim_finos": puntos(f.get("LIMITE FINOS %", "")),
                "formato": str(f.get("FORMATO", "") or "").strip().upper() or _formato_por_nombre(prov)}
    return {"razon": prov, "cuit": "", "grupo": prov, "certifica": True, "tolerancia": None, "limite": None,
            "desc_hum": False, "desc_gran": True, "tol_finos": None, "lim_finos": None, "formato": _formato_por_nombre(prov)}


def _formato_por_nombre(prov):
    """Todos los proveedores usan el formato de la certificación final (Pinar); solo MyM tiene el suyo."""
    return "MYM" if re.sub(r"\s", "", str(prov)).upper().startswith("MYM") else ""


SEP_VARIANTE = "@"        # clave de precio por punto de entrega: "CHIP@CHIP - 50 KM" (tipología@etiqueta del cuadro)


def _columnas_punto(precios):
    """Columnas opcionales de la hoja Precios: PUNTO (punto de entrega del maestro al que se aplica ese precio,
    dentro de la certificación del GRUPO) y ETIQUETA (cómo se llama esa línea en el cuadro de precios)."""
    for col in ("PUNTO", "ETIQUETA"):
        if col not in precios.columns:
            precios[col] = ""
        precios[col] = precios[col].fillna("").astype(str).str.strip().replace({"nan": "", "None": ""})
    return precios


def tipo_base(t):
    return str(t).split(SEP_VARIANTE, 1)[0]


def _de_grupo(precios, grupo):
    precios = _columnas_punto(precios)
    return precios[precios["GRUPO"].str.casefold() == str(grupo).casefold()]


def variante_punto(precios, grupo, punto, tipologia):
    """Clave de precio de un camión: la tipología, o 'TIPO@ETIQUETA' si ese punto de entrega tiene precio propio."""
    p = _de_grupo(precios, grupo)
    p = p[(p["PUNTO"].str.casefold() == str(punto).casefold()) &
          (p["TIPOLOGIA"].astype(str).str.casefold() == str(tipologia).casefold())]
    if not len(p):
        return tipologia
    et = re.sub(r"[*?~@]", "", p["ETIQUETA"].iloc[0] or f"{nombre_producto(tipologia)} {punto}").strip()
    return f"{tipologia}{SEP_VARIANTE}{et}"


def buscar_precio(precios, grupo, tipologia, tn_grupo_mes):
    precios = _columnas_punto(precios)
    base, _, etiqueta = str(tipologia).partition(SEP_VARIANTE)
    p = precios[(precios["GRUPO"].str.casefold() == str(grupo).casefold()) &
                (precios["TIPOLOGIA"].astype(str).str.casefold() == base.casefold())]
    if etiqueta:          # precio de un punto de entrega: la fila con esa etiqueta (o del punto sin etiqueta)
        p = p[p["PUNTO"] != ""]
        p = p[p.apply(lambda r: re.sub(r"[*?~@]", "", r["ETIQUETA"] or f"{nombre_producto(base)} {r['PUNTO']}").strip()
                      == etiqueta, axis=1)] if len(p) else p
    else:
        p = p[p["PUNTO"] == ""]
    for _, r in p.iterrows():
        desde = r["DESDE_TN"] if pd.notna(r["DESDE_TN"]) else 0
        hasta = r["HASTA_TN"] if pd.notna(r["HASTA_TN"]) else float("inf")
        if desde <= tn_grupo_mes < hasta and pd.notna(r["PRECIO"]) and r["PRECIO"] > 0:
            return float(r["PRECIO"])
    return None


def fijar_precio(precios, grupo, tipologia, tn_grupo_mes, valor):
    """Guarda un precio en la tabla: actualiza el tramo que corresponde o crea uno único."""
    precios = _columnas_punto(precios)
    base, _, etiqueta = str(tipologia).partition(SEP_VARIANTE)
    if etiqueta:          # precio por punto de entrega: todas las filas de esa etiqueta
        m = ((precios["GRUPO"].str.casefold() == str(grupo).casefold()) &
             (precios["TIPOLOGIA"].astype(str).str.casefold() == base.casefold()) & (precios["PUNTO"] != ""))
        m &= precios.apply(lambda r: re.sub(r"[*?~@]", "", r["ETIQUETA"] or f"{nombre_producto(base)} {r['PUNTO']}").strip()
                           == etiqueta, axis=1)
        if m.any():
            precios.loc[m, "PRECIO"] = float(valor)
            return precios
    m = ((precios["GRUPO"].str.casefold() == str(grupo).casefold()) & (precios["PUNTO"] == "") &
         (precios["TIPOLOGIA"].astype(str).str.casefold() == base.casefold()))
    for i in precios.index[m]:
        desde = precios.at[i, "DESDE_TN"] if pd.notna(precios.at[i, "DESDE_TN"]) else 0
        hasta = precios.at[i, "HASTA_TN"] if pd.notna(precios.at[i, "HASTA_TN"]) else float("inf")
        if desde <= tn_grupo_mes < hasta:
            precios.at[i, "PRECIO"] = float(valor)
            return precios
    nueva = pd.DataFrame([{"GRUPO": grupo, "TIPOLOGIA": base, "DESDE_TN": 0, "HASTA_TN": None,
                           "PRECIO": float(valor), "PUNTO": "", "ETIQUETA": ""}])
    return pd.concat([precios, nueva], ignore_index=True)


def validar_config(conf, precios):
    obs = []
    grupos = {g.casefold() for g in conf["GRUPO"] if g} | {p.casefold() for p in conf.index}
    for i, r in precios.iterrows():
        fila = f"Precios fila {i + 2}"
        if r["GRUPO"].casefold() not in grupos:
            parecido = difflib.get_close_matches(r["GRUPO"], list(conf["GRUPO"].unique()), n=1, cutoff=0.6)
            obs.append((CRIT, r["GRUPO"], f"{fila}: el GRUPO '{r['GRUPO']}' no existe en Proveedores"
                        + (f" (¿quisiste decir '{parecido[0]}'?)" if parecido else "")))
        if r["TIPOLOGIA"] not in TIPOLOGIAS:
            obs.append((CRIT, r["GRUPO"], f"{fila}: tipología '{r['TIPOLOGIA']}' no válida"))
        if pd.isna(r["PRECIO"]) or r["PRECIO"] <= 0:
            obs.append((CRIT, r["GRUPO"], f"{fila}: precio vacío, cero o no numérico"))
        if pd.notna(r["HASTA_TN"]) and pd.notna(r["DESDE_TN"]) and r["HASTA_TN"] <= r["DESDE_TN"]:
            obs.append((CRIT, r["GRUPO"], f"{fila}: HASTA_TN menor o igual a DESDE_TN"))
    precios = _columnas_punto(precios)
    for (g, t, _pt), p in precios.dropna(subset=["PRECIO"]).groupby([precios["GRUPO"].str.casefold(), "TIPOLOGIA",
                                                                     precios["PUNTO"].str.casefold()]):
        p = p.sort_values("DESDE_TN")
        desde, hasta = p["DESDE_TN"].fillna(0).tolist(), p["HASTA_TN"].fillna(float("inf")).tolist()
        pr = p["PRECIO"].tolist()
        if desde[0] != 0:
            obs.append((CRIT, p["GRUPO"].iloc[0], f"Tramos de {t}: el primero empieza en {desde[0]} y no en 0"))
        for k in range(1, len(desde)):
            if desde[k] != hasta[k - 1]:
                obs.append((CRIT, p["GRUPO"].iloc[0], f"Tramos de {t}: hueco o superposición entre "
                            f"{hasta[k-1]:g} y {desde[k]:g} TN"))
            if pr[k] > pr[k - 1] * 1.5 or pr[k] < pr[k - 1] / 1.5:
                obs.append((CRIT, p["GRUPO"].iloc[0], f"Tramos de {t}: el precio salta de "
                            f"{pesos(pr[k-1])} a {pesos(pr[k])}. ¿Tipeo?"))
    for t, p in precios.dropna(subset=["PRECIO"]).groupby("TIPOLOGIA"):
        med = p["PRECIO"].median()
        if len(p) >= 3:
            for _, r in p.iterrows():
                if r["PRECIO"] > med * 3 or r["PRECIO"] < med / 3:
                    obs.append((CRIT, r["GRUPO"], f"Precio de {t} {pesos(r['PRECIO'])} muy distinto al "
                                f"habitual ({pesos(med)}). ¿Cero de más o de menos?"))
    return obs


# -----------------------------------------------------------------------------
# 3. Correcciones manuales (capa sobre el maestro)
# -----------------------------------------------------------------------------
def estado_nuevo(anio, mes, proximo_numero, prefijo="0001"):
    return {"anio": int(anio), "mes": int(mes), "proximo": int(proximo_numero), "prefijo": str(prefijo).zfill(4),
            "corr": {"editados": {}, "eliminados": {}, "agregados": {}, "finos": {}, "humedad_camion": {},
                     "desc_gran": {}, "pen_hum": {}},
            "meta": {}, "log": []}


def aplicar_correcciones(d, corr):
    d2 = d.copy()
    for _id, campos in corr["editados"].items():
        m = d2["_id"] == _id
        for campo, v in campos.items():
            col = CAMPOS[campo]
            if campo == "Fecha":
                v = pd.Timestamp(v)
            elif campo == "Kilos":
                v = float(v)
            elif campo == "Producto":
                d2.loc[m, "_tipo_valido"] = v in TIPOLOGIAS
            d2.loc[m, col] = v
        if "Kilos" in campos:          # el neto se fijó a mano: lleno/vacío ya no aplican
            d2.loc[m, [C_LLENO, C_VACIO]] = float("nan")
        if m.any():
            d2.loc[m, "_origen"] = "corregido"
    if corr["eliminados"]:
        d2 = d2[~d2["_id"].isin(list(corr["eliminados"]))]
    if corr["agregados"]:
        filas = []
        for _id, f in corr["agregados"].items():
            filas.append({"_id": _id, C_DIA: pd.Timestamp(f["Fecha"]), C_PROV: f["Proveedor"],
                          C_BATEA: f.get("Batea", ""), C_REM_PLANTA: f.get("Rto. Planta", ""),
                          C_REM_PROV: f.get("Remito Proveedor", ""), C_TIPO: f["Producto"],
                          C_NETO: float(f["Kilos"]), C_MOV: "Ingreso",
                          "_tipo_valido": f["Producto"] in TIPOLOGIAS, "_origen": "agregado"})
        d2 = pd.concat([d2, pd.DataFrame(filas)], ignore_index=True)
    d2[C_DIA] = pd.to_datetime(d2[C_DIA]).dt.normalize()
    return d2


# -----------------------------------------------------------------------------
# 4. Validación de remitos
# -----------------------------------------------------------------------------
def _num_remito(s):
    m = re.fullmatch(r"\d{6}-(\d+)", str(s))
    return int(m.group(1)) if m else None


def validar_maestro(df_raw, d, conf, anio, mes):
    """por_id: {_id: [(sev, msg, bloqueante)]}; globales: observaciones que no son de
    un remito que esté dentro de una certificación (ej. remito que quedó fuera del mes)."""
    ini = pd.Timestamp(anio, mes, 1)
    fin = ini + pd.offsets.MonthEnd(0)
    por_id, globales = {}, []

    def add(_id, sev, msg, bloq=False):
        por_id.setdefault(_id, []).append((sev, msg, bloq))

    mes_df = d[(d[C_DIA] >= ini) & (d[C_DIA] <= fin)]
    for _, f in mes_df.iterrows():
        i = f["_id"]
        if not f[C_REM_PLANTA]:
            add(i, CRIT, "Ingreso sin N° de remito de planta")
        if f[C_NETO] <= 0:
            add(i, CRIT, "Peso neto vacío o cero", True)
        elif f[C_NETO] > PESO_MAX * 1.2:
            add(i, CRIT, f"Peso neto {miles(f[C_NETO])} kg imposible para un camión. ¿Tipeo?", True)
        elif f[C_NETO] > PESO_MAX or f[C_NETO] < PESO_MIN:
            add(i, ADV, f"Peso neto fuera de rango habitual ({miles(f[C_NETO])} kg)")
        if pd.notna(f.get(C_LLENO)) and pd.notna(f.get(C_VACIO)) and abs(f[C_LLENO] - f[C_VACIO] - f[C_NETO]) > 1:
            add(i, CRIT, f"Lleno − vacío ({miles(f[C_LLENO] - f[C_VACIO])}) ≠ neto ({miles(f[C_NETO])})")
        if not f["_tipo_valido"]:
            add(i, CRIT, f"Tipología '{f[C_TIPO]}' no reconocida", True)
        if f[C_REM_PROV].upper() in SIN_REMITO:
            add(i, ADV, "Sin remito de proveedor")
        q = _txt(f.get(C_QUINC, ""))
        if q and f["_origen"] == "maestro" and q[0] != ("1" if f[C_DIA].day <= 15 else "2"):
            add(i, ADV, f"La columna Quincena dice '{q}' y la fecha es {f[C_DIA]:%d/%m}")
        if f["_origen"] in ("corregido", "agregado"):
            add(i, ADV, "Remito corregido o agregado a mano (ver registro de cambios)")

    med = mes_df.groupby(C_PROV)[C_NETO].transform("median")
    for _, f in mes_df[(mes_df[C_NETO] > med * 2.5) | (mes_df[C_NETO] < med * 0.35)].iterrows():
        add(f["_id"], ADV, f"Peso {miles(f[C_NETO])} kg muy distinto a lo habitual del proveedor")

    for rem, g in mes_df[mes_df[C_REM_PLANTA] != ""].groupby(C_REM_PLANTA):
        if len(g) > 1:
            for i in g["_id"]:
                add(i, CRIT, f"N° de remito de planta {rem} repetido {len(g)} veces")

    hist = d[(d[C_DIA] > ini - pd.DateOffset(months=6)) & (d[C_DIA] <= fin)]
    hist = hist[~hist[C_REM_PROV].str.upper().isin(SIN_REMITO)]
    # proveedores que habitualmente ponen el mismo remito a varios camiones (ej. MyM): solo aviso, no crítico
    rep_ = hist.groupby([C_PROV, C_REM_PROV])[C_REM_PLANTA].transform("size") > 1
    habitual = (rep_.groupby(hist[C_PROV]).mean() >= 0.2).to_dict()
    for (prov, rp), g in hist.groupby([C_PROV, C_REM_PROV]):
        if len(g) > 1 and (g[C_DIA] >= ini).any():
            otros = ", ".join(f"{r} ({fe:%d/%m})" for r, fe in zip(g[C_REM_PLANTA], g[C_DIA]))
            for i in g.loc[g[C_DIA] >= ini, "_id"]:
                add(i, ADV if habitual.get(prov) else CRIT, f"Remito de proveedor {rp} usado más de una vez: {otros}"
                    + (" (habitual en este proveedor)" if habitual.get(prov) else ""))

    s = d.loc[d["_origen"] != "agregado", [C_DIA, C_REM_PLANTA, C_PROV, "_id"]].copy()
    s["n"] = s[C_REM_PLANTA].map(_num_remito)
    s = s.dropna(subset=["n"]).sort_values("n")
    s["ant"], s["sig"] = s[C_DIA].shift(1), s[C_DIA].shift(-1)
    vecinos_ok = (s["sig"] - s["ant"]).abs() <= pd.Timedelta(days=3)
    lejos = ((s[C_DIA] - s["ant"]).abs() > pd.Timedelta(days=1)) & ((s[C_DIA] - s["sig"]).abs() > pd.Timedelta(days=1))
    for _, f in s[vecinos_ok & lejos].iterrows():
        en_mes, vecino_en_mes = ini <= f[C_DIA] <= fin, ini <= f["ant"] <= fin
        if not (en_mes or vecino_en_mes):
            continue
        msg = (f"Fecha {f[C_DIA]:%d/%m/%Y} no coincide con los remitos vecinos "
               f"({f['ant']:%d/%m} y {f['sig']:%d/%m}). ¿Error de tipeo en la fecha?")
        if en_mes:
            add(f["_id"], CRIT, msg)
        else:
            globales.append({"sev": CRIT, "prov": f[C_PROV], "fecha": f[C_DIA], "rem": f[C_REM_PLANTA],
                             "msg": msg + f" Por esa fecha NO entra en {MESES[mes]}: si corresponde, "
                                          "traelo con 'Traer remito del maestro' y corregí la fecha."})

    m_raw = df_raw[(df_raw[C_DIA] >= ini) & (df_raw[C_DIA] <= fin)]
    tipo_mov = _limpio(m_raw[C_MOV]).str.lower()
    neto_raw = pd.to_numeric(m_raw[C_NETO], errors="coerce").fillna(0)
    ya = set(d[C_REM_PLANTA])
    for _, f in m_raw[~tipo_mov.isin(MOVIMIENTOS) & (neto_raw != 0)].iterrows():
        if _txt(f[C_REM_PLANTA]) in ya:
            continue
        globales.append({"sev": CRIT, "prov": _txt(f[C_PROV]), "fecha": f[C_DIA], "rem": _txt(f[C_REM_PLANTA]),
                         "msg": f"TIPO DE MOVIMIENTO '{_txt(f[C_MOV])}' no reconocido: no se certifica. "
                                "Si es un ingreso, traelo con 'Traer remito del maestro'."})

    if conf is not None:
        for prov in sorted(mes_df[C_PROV].unique()):
            if prov not in conf.index:
                parecido = difflib.get_close_matches(prov, list(conf.index), n=1, cutoff=0.8)
                globales.append({"sev": CRIT, "prov": prov, "fecha": None, "rem": "",
                                 "msg": "Proveedor nuevo, no está en el config" +
                                        (f". Muy parecido a '{parecido[0]}': ¿error de tipeo? Corregilo en "
                                         "UNIFICAR CON del config." if parecido else
                                         ". Cargá razón social, CUIT y precio en el editor.")})
    return por_id, globales



# -----------------------------------------------------------------------------
# 4b. Humedad (planilla de laboratorio "PROVEEDORES - HUMEDAD Y GRANULOMETRIA")
# -----------------------------------------------------------------------------
# Se usan los valores que calcula la propia planilla en la hoja GENERAL, por muestra:
#   Humedad = máx(horno, termobalanza) · Límite de referencia humedad (según el tipo de
#   proveedor: Normal = referencia chipera + 3 %, Especial 1 = 45 %, etc.)
# Descuento por período = promedio de humedad de las muestras del proveedor y producto
# en ese período contra el promedio de sus límites. Si no supera el límite, no hay descuento.
MATERIAL_HUM = {"CHIP": "CHIP", "ASERRIN": "ASERRIN", "RALEO": "Fino/Raleo", "FINO/RALEO": "Fino/Raleo"}
METODOS_HUMEDAD = {"PUNTOS": "TN × puntos de humedad sobre el límite",
                   "BASE_SECA": "TN equivalentes a la humedad límite: TN × (1 − H) / (1 − L)"}


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
    try:
        wb = load_workbook(ruta, read_only=True)
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
    return None


def leer_humedad(ruta, hoja="GENERAL"):
    raw = pd.read_excel(ruta, sheet_name=hoja, header=None, nrows=15)
    fila_h = next(i for i, r in raw.iterrows() if "PROVEEDOR" in r.values and "Fecha" in r.values)
    g = pd.read_excel(ruta, sheet_name=hoja, header=fila_h)
    g.columns = [str(c).strip() for c in g.columns]
    h = pd.DataFrame({
        "Fecha": pd.to_datetime(g["Fecha"], errors="coerce").dt.normalize(),
        "Proveedor": g["PROVEEDOR"].apply(_txt).str.replace(r"\s+", " ", regex=True),
        "Tipo": g.get("Tipo de proveedor", pd.Series("", index=g.index)).apply(_txt),
        "Material": g["Material"].apply(_txt).str.upper().map(lambda m: MATERIAL_HUM.get(m, m)),
        "Remito": g["REMITO"].apply(_txt),
        "Minuto": g["Hora"].apply(_minutos) if "Hora" in g.columns else float("nan"),
        "Humedad": pd.to_numeric(g["Humedad"], errors="coerce"),
        "Limite": pd.to_numeric(g["Limite de referencia humedad"], errors="coerce"),
        "Finos": pd.to_numeric(g["< #04 (%)"], errors="coerce") if "< #04 (%)" in g.columns else float("nan"),
        "FinosAlto": pd.to_numeric(g["Alto"], errors="coerce") if "Alto" in g.columns else float("nan"),
    })
    h.loc[~h["Humedad"].between(0, 1, inclusive="neither"), "Humedad"] = float("nan")
    h.loc[~h["Finos"].between(0, 1), "Finos"] = float("nan")
    h = h.dropna(subset=["Fecha"])
    h = h[(h["Proveedor"] != "") & (h["Humedad"].notna() | h["Finos"].notna())]
    return h.reset_index(drop=True)


def _minutos(v):
    """Hora de la muestra → minutos desde las 0 h (acepta time, datetime, fracción de día o texto hh:mm)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return float("nan")
    if hasattr(v, "hour"):
        return v.hour * 60 + v.minute
    if isinstance(v, (int, float)):
        return float(v) * 1440 if v < 1 else float("nan")
    m = re.match(r"(\d{1,2})[:.](\d{2})", str(v))
    return int(m.group(1)) * 60 + int(m.group(2)) if m else float("nan")


def descuento_finos(finos_pct, tolerancia=None, totalidad=None):
    """% de finos medido (ej. 11.3) → % que pasa a aserrín, redondeado a 1 decimal como la planilla."""
    tol = FINOS_TOLERANCIA if tolerancia is None else tolerancia
    tot = FINOS_TOTALIDAD if totalidad is None else totalidad
    if finos_pct is not None and not pd.isna(finos_pct):
        finos_pct = round(float(finos_pct), 9)       # como Excel: 14,999999999 cuenta como 15
    if finos_pct is None or pd.isna(finos_pct) or finos_pct <= tol:
        return 0.0
    if finos_pct <= tot:
        return round(finos_pct - tol + 1e-9, 1)
    return round(finos_pct + 1e-9, 1)


def familia_lab(conf, pe, mes_df):
    """Puntos de entrega de la misma empresa (mismo GRUPO en el config) que entregan en el mes. Si son varios
    (ej. Pallets Sur EP y Pallets Sur LT) las muestras del laboratorio se asignan entre todos ellos, porque
    el laboratorio a veces anota uno con el nombre del otro."""
    if conf is None:
        return [pe]
    g = str(info_proveedor(conf, pe)["grupo"]).casefold()
    fam = sorted({p for p in mes_df["_prov_lab"].dropna().unique()
                  if str(info_proveedor(conf, p)["grupo"]).casefold() == g}, key=str.casefold)
    return fam if pe in fam else [pe]


def asignar_finos(h, prov, datos, manual=None, manual_h=None):
    """Asigna a cada camión del proveedor su muestra de granulometría.
    1) mismo N° de remito del proveedor; 2) el resto, por día: si hay tantas muestras como camiones,
    en orden (muestras por hora ↔ camiones por hora de entrada a planta); si no, por cercanía de horario
    (muestra ↔ entrada a planta, hasta 90 min). manual: {_id: fracción} cargado en el editor.
    Devuelve {_id: {"F": fracción, "hora": "hh:mm", "origen": texto}}."""
    out = {}
    if h is not None and len(datos):
        provs = list(prov) if isinstance(prov, (list, tuple, set)) else [prov]
        m = h[(h["ProvM"].isin(provs)) & (h["Finos"].notna() | h["Humedad"].notna())].copy()
        if len(m):
            camiones = datos.assign(_dia=datos[C_DIA].dt.normalize(),
                                    _min=pd.to_numeric(datos.get("ENTRADA"), errors="coerce") * 1440,
                                    _n=datos[C_REM_PLANTA].map(_num_remito))
            usadas = set()
            # 1) por remito del proveedor (solo si ese remito es de un único camión: hay proveedores que
            #    repiten el mismo remito en varios viajes)
            rp_cuenta = camiones[C_REM_PROV].astype(str).map(_rem_clave).value_counts()
            m_clave = m["Remito"].astype(str).map(_rem_clave)
            for i, f in camiones.iterrows():
                rp = _rem_clave(f[C_REM_PROV])
                if not rp or rp_cuenta.get(rp, 0) != 1:
                    continue
                cand = m[(m_clave == rp) & (m["Material"] == f[C_TIPO]) & ~m.index.isin(usadas)]
                if len(cand):
                    j = cand.index[0]
                    usadas.add(j)
                    out[f["_id"]] = dict(_muestra(m, j), origen="muestra por N° de remito")
            # 2) por día y orden / horario
            for (dia, prod), cs in camiones[~camiones["_id"].isin(out)].groupby(["_dia", C_TIPO]):
                ms = m[(m["Fecha"] == dia) & (m["Material"] == prod) & ~m.index.isin(usadas)].sort_values("Minuto")
                if not len(ms):
                    continue
                cs = _orden_entrada(cs)
                if len(ms) == len(cs):
                    pares = list(zip(cs.index, ms.index))
                    como = "muestra del día (en orden)"
                else:
                    pares, libres = [], list(ms.index)
                    for i, f in cs.iterrows():
                        if not libres or pd.isna(f["_min"]):
                            continue
                        j = min(libres, key=lambda k: abs((ms.at[k, "Minuto"] if pd.notna(ms.at[k, "Minuto"]) else 1e9) - f["_min"]))
                        if pd.notna(ms.at[j, "Minuto"]) and abs(ms.at[j, "Minuto"] - f["_min"]) <= 90:
                            pares.append((i, j))
                            libres.remove(j)
                    como = "muestra del día (por horario)"
                for i, j in pares:
                    usadas.add(j)
                    out[camiones.at[i, "_id"]] = dict(_muestra(ms, j), origen=como)
    ids = set(datos["_id"])
    for _id, v in (manual or {}).items():
        if _id in ids:
            previo = out.get(_id, {})
            out[_id] = dict(previo, F=None if v is None else float(v), origen="cargado a mano",
                            hora=previo.get("hora", ""))
    for _id, v in (manual_h or {}).items():
        if _id in ids:
            previo = out.get(_id, {"F": None, "hora": "", "L": None})
            out[_id] = dict(previo, H=None if v is None else float(v), origen_h="cargada a mano")
    return out


def _hermanos(conf, puntos):
    """Otros nombres del maestro de la MISMA empresa (mismo GRUPO o mismo CUIT) que los puntos dados.
    Ej.: el laboratorio anota 'Pallets Sur S' muestras de camiones de Pallets Sur LT / EP."""
    if conf is None:
        return []
    puntos = set(puntos)
    grupos = {str(conf.at[p, "GRUPO"]).strip().casefold() for p in puntos if p in conf.index} - {""}
    cuits = {str(conf.at[p, "CUIT"]).strip() for p in puntos if p in conf.index} - {""}
    return [p for p in conf.index if p not in puntos and
            (str(conf.at[p, "GRUPO"]).strip().casefold() in grupos or str(conf.at[p, "CUIT"]).strip() in cuits)]


def finos_por_remito_hermanos(h, hermanos, faltan, mes_df, dias=3):
    """Camiones que quedaron sin muestra: se busca una muestra anotada con otro nombre de la misma empresa que
    tenga el MISMO N° de remito del proveedor, del mismo producto, con hasta `dias` de diferencia, y que no
    corresponda a un camión propio de ese otro nombre ese día (muestra 'huérfana')."""
    out = {}
    if h is None or not hermanos or not len(faltan):
        return out
    m = h[h["ProvM"].isin(hermanos) & (h["Finos"].notna() | h["Humedad"].notna())].copy()
    if not len(m):
        return out
    dias_propios = set(zip(mes_df["_prov_lab"], mes_df[C_DIA].dt.normalize()))
    m = m[[(p, f) not in dias_propios for p, f in zip(m["ProvM"], m["Fecha"])]]
    m_clave = m["Remito"].astype(str).map(_rem_clave)
    rp_cuenta = faltan[C_REM_PROV].astype(str).map(_rem_clave).value_counts()
    usadas = set()
    for _, f in faltan.iterrows():
        rp = _rem_clave(f[C_REM_PROV])
        if not rp or rp_cuenta.get(rp, 0) != 1:
            continue
        cand = m[(m_clave == rp) & (m["Material"] == f[C_TIPO]) & ~m.index.isin(usadas)
                 & ((m["Fecha"] - f[C_DIA].normalize()).abs() <= pd.Timedelta(days=dias))]
        if len(cand) == 1:
            j = cand.index[0]
            usadas.add(j)
            out[f["_id"]] = dict(_muestra(m, j), origen=f"muestra por N° de remito (anotada como {m.at[j, 'Proveedor']})")
    return out


def _rem_clave(s):
    """Remito del proveedor comparable: '000-465', '0000-465' y '465' son el mismo; S/R → ''."""
    t = str(s).strip().upper()
    if t in SIN_REMITO or t in ("", "NAN", "NONE"):
        return ""
    nums = re.findall(r"\d+", t)
    return str(int(nums[-1])) if nums else t


def _orden_entrada(cs, margen=5):
    """Camiones en orden de entrada a planta; los que entraron con menos de `margen` minutos de diferencia
    se ordenan por N° de remito de planta (la balanza no distingue cuál se muestreó primero)."""
    cs = cs.sort_values(["_min", "_n"])
    grupo, ult, g = [], None, -1
    for v in cs["_min"]:
        if ult is None or pd.isna(v) or pd.isna(ult) or v - ult > margen:
            g += 1
        grupo.append(g)
        ult = v
    return cs.assign(_g=grupo).sort_values(["_g", "_n"])


def _muestra(m, j):
    val = lambda c: None if pd.isna(m.at[j, c]) else float(m.at[j, c])
    return {"F": val("Finos"), "H": val("Humedad"), "L": val("Limite"), "hora": _hhmm(m.at[j, "Minuto"])}


def _hhmm(minuto):
    return "" if minuto is None or pd.isna(minuto) else f"{int(minuto) // 60:02d}:{int(minuto) % 60:02d}"


def _clave(s):
    return re.sub(r"[^a-z0-9]", "", str(s).casefold())


def mapear_humedad(h, d, conf=None):
    """Vincula cada nombre de la planilla de humedad con un proveedor del maestro.
    Orden: columna 'NOMBRE HUMEDAD' del config → nombre igual → nombre contenido →
    remitos de proveedor en común → nombre parecido."""
    # nombres con los que el maestro registró cada camión (antes de separar por patente: ej. LAPACHO (El monte)
    # transportado por AGROMAD sigue siendo LAPACHO para el laboratorio)
    col = "_prov_lab" if "_prov_lab" in d.columns else C_PROV
    provs = sorted(set(d[col].dropna().astype(str)) - {""})
    kp = {_clave(p): p for p in provs}
    recientes = d.loc[d[C_DIA] >= d[C_DIA].max() - pd.Timedelta(days=90), col].value_counts() if len(d) else pd.Series(dtype=int)
    manual = {}
    if conf is not None and "NOMBRE HUMEDAD" in conf.columns:
        for p, v in conf["NOMBRE HUMEDAD"].items():
            for n in str(v).split(","):
                if n.strip():
                    manual[_clave(n)] = p
    rp = d[~d[C_REM_PROV].str.upper().isin(SIN_REMITO)][[C_REM_PROV, col]].rename(columns={col: C_PROV})
    mapa, filas = {}, []
    for n in sorted(h["Proveedor"].unique()):
        k = _clave(n)
        dest, como = None, ""
        if k in manual:
            dest, como = manual[k], "config (NOMBRE HUMEDAD)"
        elif k in kp:
            dest, como = kp[k], "mismo nombre"
        else:
            cont = [p for kk, p in kp.items() if len(k) >= 5 and (k in kk or kk in k)]
            if len(cont) > 1:          # ej. LAPACHO → "LAPACHO (El monte)" antes que "Chipera lapacho"
                pref = [p for p in cont if _clave(p).startswith(k)]
                cont = pref or cont
            if len(cont) > 1 and conf is not None and len({unidad(conf, p) for p in cont}) == 1:
                cont = [max(cont, key=lambda p: int(recientes.get(p, 0)))]      # misma certificación: el que entrega
            if len(cont) == 1:
                dest, como = cont[0], "nombre contenido"
            else:
                rems = set(h.loc[(h["Proveedor"] == n) & ~h["Remito"].str.upper().isin(SIN_REMITO), "Remito"])
                cnt = rp[rp[C_REM_PROV].isin(rems)][C_PROV].value_counts()
                if len(cnt) and cnt.iloc[0] >= 3 and cnt.iloc[0] >= 0.5 * len(rems):
                    dest, como = cnt.index[0], f"remitos en común ({cnt.iloc[0]})"
                else:
                    m = difflib.get_close_matches(k, list(kp), n=1, cutoff=0.82)
                    if m:
                        dest, como = kp[m[0]], "nombre parecido"
        mapa[n] = dest
        filas.append({"Nombre en planilla de humedad": n, "Proveedor en maestro": dest or "— sin vincular —",
                      "Cómo se vinculó": como or "no se encontró: cargalo en NOMBRE HUMEDAD del config"})
    h = h.assign(ProvM=h["Proveedor"].map(mapa))
    return h, pd.DataFrame(filas)


def calcular_humedad(h, prov, producto, desde, hasta, tn, metodo="PUNTOS", manual=None,
                     tolerancia=None, limite_fijo=None, aplica_hum=True, aplica_gran=False, tol_finos=None):
    """Descuentos de un producto en un período (promedio de las muestras del laboratorio).
    HUMEDAD (solo si el proveedor tiene DESCUENTA HUMEDAD = SI):
        límite = límite de la planilla (o LIMITE HUMEDAD % fijo) + TOLERANCIA HUMEDAD %
    GRANULOMETRÍA (si DESCUENTA GRANULOMETRIA = SI): % de finos (< #04) sobre TOLERANCIA FINOS %.
    Descuento total = descuento humedad + descuento granulometría (en puntos de % sobre las TN)."""
    manual = manual or {}
    r = {"muestras": 0, "H": None, "L": None, "L_lab": None, "tol": tolerancia or 0.0, "L_fijo": limite_fijo,
         "exceso": 0.0, "desc_hum": 0.0, "aplica_hum": bool(aplica_hum), "origen": "sin muestras",
         "g_muestras": 0, "F": None, "F_tol": tol_finos, "F_ref": None, "desc_gran": 0.0,
         "aplica_gran": bool(aplica_gran), "g_origen": "sin muestras"}
    if h is not None:
        provs = prov if isinstance(prov, (list, tuple, set)) else [prov]
        m = h[(h["ProvM"].isin(provs)) & (h["Material"] == producto) &
              (h["Fecha"].dt.date >= desde) & (h["Fecha"].dt.date <= hasta)]
        mh = m[m["Humedad"].notna()]
        if len(mh):
            r.update(muestras=len(mh), H=float(mh["Humedad"].mean()), L=float(mh["Limite"].mean()),
                     origen=f"promedio de {len(mh)} muestras del laboratorio")
        mg = m[m["Finos"].notna()]
        if len(mg):
            r.update(g_muestras=len(mg), F=float(mg["Finos"].mean()),
                     F_ref=float(mg["FinosAlto"].mean()) if mg["FinosAlto"].notna().any() else None,
                     g_origen=f"promedio de {len(mg)} muestras del laboratorio")
    if manual.get("H") is not None:
        r["H"], r["origen"] = float(manual["H"]), "humedad cargada a mano"
    if manual.get("L") is not None:
        r["L"] = float(manual["L"])
    if manual.get("F") is not None:
        r["F"], r["g_origen"] = float(manual["F"]), "finos cargados a mano"
    r["L_lab"] = r["L"]
    if limite_fijo is not None and manual.get("L") is None:
        r["L"] = limite_fijo
    if r["L"] is not None:
        r["L"] = r["L"] + r["tol"]
    if r["aplica_hum"] and r["H"] is not None and r["L"] is not None and r["H"] > r["L"]:
        r["exceso"] = r["H"] - r["L"]
        r["desc_hum"] = r["exceso"] if metodo == "PUNTOS" else 1 - (1 - r["H"]) / (1 - r["L"])
    if r["aplica_gran"] and r["F"] is not None and r["F_tol"] is not None and r["F"] > r["F_tol"]:
        r["desc_gran"] = r["F"] - r["F_tol"]
    r["desc_pct"] = r["desc_hum"] + r["desc_gran"]
    r["tn"] = round(tn, 2)
    r["tn_pagar"] = round(r["tn"] * (1 - r["desc_pct"]), 2)       # igual que la fórmula del Excel
    r["tn_desc"] = round(r["tn"] - r["tn_pagar"], 2)
    r["tn_desc_hum"] = round(r["tn"] * r["desc_hum"], 2)
    r["tn_desc_gran"] = round(r["tn_desc"] - r["tn_desc_hum"], 2)
    return r


# -----------------------------------------------------------------------------
# 5. Armado de certificaciones (en memoria)
# -----------------------------------------------------------------------------
def periodos(anio, mes, tipo="ambas"):
    """tipo: lista de períodos (ej. ["1Q", "Mensual"]) o un texto ("ambas", "1Q", "2Q+Mensual"...)."""
    ult = calendar.monthrange(anio, mes)[1]
    if isinstance(tipo, (list, tuple, set)):
        todos = {"1Q": (date(anio, mes, 1), date(anio, mes, DIA_CORTE)),
                 "2Q": (date(anio, mes, DIA_CORTE + 1), date(anio, mes, ult)),
                 "Mensual": (date(anio, mes, 1), date(anio, mes, ult))}
        return [(k, *todos[k]) for k in ("1Q", "2Q", "Mensual") if k in tipo]
    p = []
    if tipo in ("quincenal", "ambas", "1Q"):
        p.append(("1Q", date(anio, mes, 1), date(anio, mes, DIA_CORTE)))
    if tipo in ("quincenal", "ambas", "2Q+Mensual", "2Q"):
        p.append(("2Q", date(anio, mes, DIA_CORTE + 1), date(anio, mes, ult)))
    if tipo in ("mensual", "ambas", "2Q+Mensual"):
        p.append(("Mensual", date(anio, mes, 1), date(anio, mes, ult)))
    return p


# Formato MyM: humedad y granulometría camión por camión (planilla "Certificación MyM")
MYM_COLS = ["Hum", "HumRef", "Pen", "KgChip", "KgAse", "KgCorto", "KgLargo", "KgRaleo"]
MYM_OTROS = {"C CORTO": "KgCorto", "C LARGO": "KgLargo", "Fino/Raleo": "KgRaleo"}


def limites_por_dia(h):
    """Humedad de referencia del día (límite de la planilla del laboratorio, el más repetido de ese día)."""
    if h is None:
        return {}
    x = h.dropna(subset=["Limite"])
    return x.groupby(x["Fecha"].dt.date)["Limite"].agg(lambda s: float(s.round(4).mode().iloc[0])).to_dict()


def calcular_mym(filas, muestras, alto, muy_alto, ref_dia, tol_hum=0.0, sin_gran=False, sin_hum=False,
                 man_desc=None, man_pen=None):
    """Como las fórmulas de la planilla de MyM, por camión:
       penalización = humedad medida − humedad de referencia (si la supera; solo CHIP y ASERRÍN)
       aserrín (CHIP): % < Alto → 0 · % < Muy Alto → (% − Alto) × kg · si no → % × kg ; × (1 − penalización)
       chip = (kg − aserrín) × (1 − penalización) · ASERRÍN = kg × (1 − penalización) · el resto, kg completos.
       sin_gran / sin_hum: descuento quitado a mano en la certificación; man_desc {_id: % aserrín} y
       man_pen {_id: fracción} reemplazan el valor calculado de ese camión."""
    man_desc, man_pen = man_desc or {}, man_pen or {}
    out = {k: [] for k in MYM_COLS + ["Finos", "FinosHora", "FinosOrigen", "DescPct", "DescKg", "KilosNeto",
                                      "DescManual", "PenManual"]}
    for _, f in filas.iterrows():
        mu = muestras.get(f["_id"], {})
        kg, prod = float(f["Kilos"]), f["Producto"]
        H, F = mu.get("H"), mu.get("F")
        ref = mu.get("L") if mu.get("L") is not None else ref_dia.get(f["Fecha"])
        ref = None if ref is None else ref + (tol_hum or 0.0)
        pen = H - ref if (prod in ("CHIP", "ASERRIN") and H is not None and ref is not None
                          and round(H, 12) > round(ref, 12)) else 0.0
        pen_man = prod in ("CHIP", "ASERRIN") and f["_id"] in man_pen
        if sin_hum:
            pen = 0.0
        elif pen_man:
            pen = float(man_pen[f["_id"]] or 0.0)
        desc_man = prod == "CHIP" and f["_id"] in man_desc
        chip = ase = 0.0
        otros = {v: 0.0 for v in MYM_OTROS.values()}
        exceso = 0.0
        if prod == "CHIP":
            Fv = round(F or 0.0, 12)          # como Excel (15 cifras): 0,14999999999999997 cuenta como 15 %
            exceso = 0.0 if Fv < alto else ((Fv - alto) if Fv < muy_alto else Fv)
            if sin_gran:
                exceso = 0.0
            elif desc_man:
                exceso = float(man_desc[f["_id"]] or 0.0) / 100
            ase = exceso * kg * (1 - pen)
            chip = (kg - ase) * (1 - pen)
        elif prod == "ASERRIN":
            ase = kg * (1 - pen)
        elif prod in MYM_OTROS:
            otros[MYM_OTROS[prod]] = kg
        for k, v in (("Hum", H), ("HumRef", ref), ("Pen", pen), ("KgChip", chip), ("KgAse", ase)):
            out[k].append(v)
        for k, v in otros.items():
            out[k].append(v)
        out["Finos"].append(F if prod == "CHIP" else None)
        out["FinosHora"].append(mu.get("hora", ""))
        out["FinosOrigen"].append(mu.get("origen", "") + (" · humedad cargada a mano" if mu.get("origen_h") else ""))
        out["DescPct"].append(exceso * 100)
        out["DescKg"].append(round(ase, 2) if prod == "CHIP" else 0.0)
        out["KilosNeto"].append(chip if prod == "CHIP" else (ase if prod == "ASERRIN" else kg))
        out["DescManual"].append(desc_man)
        out["PenManual"].append(pen_man)
    for k, v in out.items():
        filas[k] = v
    tn = {"CHIP": filas["KgChip"].sum() / 1000, "ASERRIN": filas["KgAse"].sum() / 1000}
    for prod, col in MYM_OTROS.items():
        tn[prod] = filas[col].sum() / 1000
    for prod in set(filas["Producto"]) - set(tn):
        tn[prod] = filas.loc[filas["Producto"] == prod, "Kilos"].sum() / 1000
    return {t: v for t, v in tn.items() if v > 0}


def calcular_estandar(filas, muestras, ip, tol_f, ref_dia, sin_gran=False, sin_hum=False, man_desc=None, man_pen=None):
    """Formato de la certificación final (todos menos MyM), camión por camión:
       Desc. humedad % = humedad medida − referencia del día (solo si DESCUENTA HUMEDAD = SI; CHIP y ASERRÍN)
       Desc. granulometría % = finos por encima de la tolerancia (CHIP; hasta 15 % la diferencia, más, todo)
       Kilos a pagar = KG × (1 − hum.) × (1 − gran.) · Aserrín = KG × (1 − hum.) × gran.
       Ambos % se pueden poner a mano (por camión) o quitar en toda la certificación."""
    man_desc, man_pen = man_desc or {}, man_pen or {}
    tol_h = ip["tolerancia"] if isinstance(ip["tolerancia"], float) else 0.0
    lim_fijo = ip["limite"] if isinstance(ip["limite"], float) else None
    cols = {k: [] for k in ("Finos", "FinosHora", "FinosOrigen", "Hum", "HumRef", "Pen", "PenManual",
                            "DescPct", "DescManual", "DescKg", "KilosNeto")}
    for _, f in filas.iterrows():
        mu = muestras.get(f["_id"], {})
        prod, kg, i = f["Producto"], float(f["Kilos"]), f["_id"]
        F, H = mu.get("F"), mu.get("H")
        ref = lim_fijo if lim_fijo is not None else (mu.get("L") if mu.get("L") is not None else ref_dia.get(f["Fecha"]))
        ref = None if ref is None else ref + tol_h
        pen = 0.0
        if ip["desc_hum"] and prod in ("CHIP", "ASERRIN") and H is not None and ref is not None \
                and round(H, 12) > round(ref, 12):
            pen = H - ref
        pen_man = i in man_pen and prod in ("CHIP", "ASERRIN")
        if sin_hum:
            pen, pen_man = 0.0, False
        elif pen_man:
            pen = float(man_pen[i] or 0.0)
        lim_f = ip["lim_finos"] * 100 if isinstance(ip.get("lim_finos"), float) else None
        auto = (descuento_finos(None if F is None or pd.isna(F) else F * 100, tol_f, lim_f)
                if prod == "CHIP" and ip["desc_gran"] else 0.0)
        desc_man = prod == "CHIP" and i in man_desc and not sin_gran
        dp = 0.0 if sin_gran else (float(man_desc[i] or 0.0) if desc_man else auto)
        ase = round(kg * (1 - pen) * dp / 100, 2)
        for k, v in (("Finos", F), ("FinosHora", mu.get("hora", "")),
                     ("FinosOrigen", mu.get("origen", "") + (" · humedad cargada a mano" if mu.get("origen_h") else "")),
                     ("Hum", H), ("HumRef", ref), ("Pen", pen), ("PenManual", pen_man), ("DescPct", dp),
                     ("DescManual", desc_man), ("DescKg", ase), ("KilosNeto", kg * (1 - pen) - ase)):
            cols[k].append(v)
    for k, v in cols.items():
        filas[k] = v
    tn_prod = (filas.groupby("Producto")["KilosNeto"].sum() / 1000).to_dict()
    if filas["DescKg"].sum() > 0:
        tn_prod["ASERRIN"] = tn_prod.get("ASERRIN", 0) + filas["DescKg"].sum() / 1000
    return tn_prod


def construir(d, d_orig, conf, precios, estado, tipo="ambas", df_raw=None, proveedores=None,
              humedad=None, metodo_humedad="PUNTOS", solo=None, validacion=None):
    """solo: conjunto de claves 'período|proveedor' a armar (None = todas).
    validacion: (por_id, globales) ya calculados, para no repetir los controles del mes."""
    anio, mes = estado["anio"], estado["mes"]
    ini, fin = pd.Timestamp(anio, mes, 1), pd.Timestamp(anio, mes, 1) + pd.offsets.MonthEnd(0)
    mes_df = d[(d[C_DIA] >= ini) & (d[C_DIA] <= fin)].copy()
    if validacion is None:
        validacion = validar_maestro(df_raw if df_raw is not None else d, d, conf, anio, mes)
    por_id, globales = validacion
    obs_cfg = validar_config(conf, precios)
    mes_df["_grupo"] = mes_df[C_PROV].map(lambda p: info_proveedor(conf, p)["grupo"])
    mes_df["_cert"] = mes_df[C_PROV].map(lambda p: unidad(conf, p))
    if "_prov_lab" not in mes_df.columns:
        mes_df["_prov_lab"] = mes_df[C_PROV]
    mes_df["_prov_lab"] = mes_df["_prov_lab"].fillna(mes_df[C_PROV]).where(mes_df["_prov_lab"].astype(str) != "",
                                                                           mes_df[C_PROV])

    certs = {}
    for per, desde, hasta in periodos(anio, mes, tipo):
        sel = mes_df[(mes_df[C_DIA].dt.date >= desde) & (mes_df[C_DIA].dt.date <= hasta)]
        for prov, datos in sorted(sel.groupby("_cert"), key=lambda x: x[0].casefold()):
            puntos_ent = sorted(datos[C_PROV].unique(), key=str.casefold)
            if proveedores and not set(puntos_ent) & set(proveedores) and prov not in proveedores:
                continue
            key = f"{per}|{prov}"
            if solo is not None and key not in solo:
                continue
            if not corresponde(conf, prov, per):          # FRECUENCIA del config (QUINCENAL / MENSUAL)
                continue
            con_cuit = [p for p in puntos_ent if info_proveedor(conf, p)["cuit"]]
            ip = info_proveedor(conf, (con_cuit or puntos_ent)[0])
            if not ip["certifica"]:
                continue
            meta = estado["meta"].get(key, {})
            sin_gran, sin_hum = bool(meta.get("sin_gran")), bool(meta.get("sin_hum"))
            man_desc = estado["corr"].get("desc_gran") or {}
            man_pen = estado["corr"].get("pen_hum") or {}
            datos = datos.sort_values([C_DIA, C_REM_PLANTA])
            filas = pd.DataFrame({k: datos[CAMPOS[k]].values for k in FILA_COLS})
            # punto de entrega tal como figura en el maestro (ej. ROBLESUR (Pérez) dentro de la de AGROMAD)
            filas["Proveedor"] = datos["_prov_lab"].values
            filas["Fecha"] = pd.to_datetime(filas["Fecha"]).dt.date
            filas["_id"] = datos["_id"].values
            # Granulometría por camión: lo descontado pasa a ASERRÍN (como la planilla de certificación final)
            tol_f = ip["tol_finos"] * 100 if isinstance(ip["tol_finos"], float) else None
            es_mym = ip["formato"] == "MYM"
            finos = {}
            if humedad is not None:            # muestras de cada camión: se muestran siempre (humedad y finos %)
                con_lab = set(humedad["ProvM"].dropna())
                for pe in sorted(set(datos["_prov_lab"])):
                    # todos los camiones del mes registrados con ese nombre (aunque se certifiquen en otra
                    # certificación, ej. ROBLESUR / AGROMAD separados por patente): el orden del día no cambia.
                    # Puntos de la misma empresa (mismo GRUPO, ej. Pallets Sur EP / LT): el laboratorio a veces
                    # anota uno por otro, así que sus muestras y camiones se ordenan juntos por día y hora.
                    fam = familia_lab(conf, pe, mes_df)
                    ctx = mes_df[mes_df["_prov_lab"].isin(fam)]
                    res = asignar_finos(humedad, fam, ctx, estado["corr"].get("finos"),
                                        estado["corr"].get("humedad_camion"))
                    ids = set(mes_df.loc[mes_df["_prov_lab"] == pe, "_id"])
                    finos.update({i: v for i, v in res.items() if i in ids})
                # puntos de entrega sin nombre propio en la planilla del laboratorio (ej. AGROMAD (+10 KM)):
                # usan las muestras de los otros puntos de la misma certificación
                sin_lab = [pe for pe in set(datos["_prov_lab"]) if pe not in con_lab]
                con_lab_u = [pe for pe in miembros(conf, prov) if pe in con_lab]      # todos los puntos de la unidad
                if sin_lab and con_lab_u:
                    finos.update(asignar_finos(humedad, con_lab_u,
                                               datos[datos["_prov_lab"].isin(sin_lab)], estado["corr"].get("finos"),
                                               estado["corr"].get("humedad_camion")))
                # los que siguen sin muestra: ¿la anotaron con otro nombre de la misma empresa? (por N° de remito)
                con_dato = {k for k, v in finos.items() if v.get("F") is not None or v.get("H") is not None
                            or v.get("origen") == "cargado a mano"}
                faltan = datos[~datos["_id"].isin(con_dato)]
                if len(faltan):
                    extra = finos_por_remito_hermanos(humedad, _hermanos(conf, set(datos["_prov_lab"]) | set(puntos_ent)),
                                                      faltan, mes_df)
                    for k, v in extra.items():
                        finos[k] = dict(v, **{c: finos.get(k, {}).get(c) for c in ("F", "H") if finos.get(k, {}).get(c) is not None})
            if es_mym:
                alto = ip["tol_finos"] if isinstance(ip["tol_finos"], float) else None
                if alto is None and humedad is not None:
                    fa = humedad.loc[humedad["ProvM"].isin(set(puntos_ent) | set(datos["_prov_lab"])), "FinosAlto"].dropna()
                    alto = float(fa.mode().iloc[0]) if len(fa) else None
                alto = FINOS_TOLERANCIA / 100 if alto is None else alto
                tol_f = alto * 100
                muy_alto = ip["lim_finos"] if isinstance(ip.get("lim_finos"), float) else FINOS_TOTALIDAD / 100
                tn_prod = calcular_mym(filas, finos, alto, muy_alto, limites_por_dia(humedad),
                                       ip["tolerancia"] if isinstance(ip["tolerancia"], float) else 0.0,
                                       sin_gran, sin_hum, man_desc, man_pen)
            else:
                tn_prod = calcular_estandar(filas, finos, ip, tol_f, limites_por_dia(humedad),
                                            sin_gran, sin_hum, man_desc, man_pen)
                # precio según el punto de entrega (hoja Precios, columna PUNTO): cada camión va a su línea del cuadro
                filas["Cuadro"] = [variante_punto(precios, ip["grupo"], pe, pr)
                                   for pe, pr in zip(filas["Proveedor"], filas["Producto"])]
                filas["CuadroAse"] = [variante_punto(precios, ip["grupo"], pe, "ASERRIN") for pe in filas["Proveedor"]]
                if (filas["Cuadro"] != filas["Producto"]).any() or (filas["CuadroAse"] != "ASERRIN").any():
                    tn_prod = (filas.groupby("Cuadro")["KilosNeto"].sum() / 1000).to_dict()
                    for k, v in (filas.groupby("CuadroAse")["DescKg"].sum() / 1000).items():
                        if v > 0:
                            tn_prod[k] = tn_prod.get(k, 0) + v
            orden = {t: i for i, t in enumerate(TIPOLOGIAS)}
            tipos = sorted([t for t, v in tn_prod.items() if v > 0], key=lambda t: (orden.get(tipo_base(t), 99), t))
            tn_auto = mes_df[(mes_df["_grupo"] == ip["grupo"]) & (mes_df[C_DIA].dt.date <= hasta)][C_NETO].sum() / 1000
            tn_grupo = float(meta.get("tn_grupo", tn_auto))
            precios_c, origen_p = {}, {}
            for t in tipos:
                if t in meta.get("precios", {}):
                    precios_c[t], origen_p[t] = float(meta["precios"][t]), "manual (solo esta certificación)"
                else:
                    precios_c[t], origen_p[t] = buscar_precio(precios, ip["grupo"], t, tn_grupo), "config"

            hum = {}        # la humedad ahora se descuenta camión por camión (columna Desc. humedad %)

            obs = []
            if humedad is not None:
                ult_m = humedad["Fecha"].max().date()
                if ult_m < hasta:
                    obs.append({"sev": ADV, "msg": f"La planilla de humedad tiene muestras hasta el {ult_m:%d/%m/%Y} "
                                f"y el período termina el {hasta:%d/%m/%Y}: subí la planilla actualizada",
                                "id": "", "rem": "", "bloq": False})
                for t, hh in hum.items():
                    if hh["aplica_hum"] and hh["muestras"] == 0 and hh["origen"] == "sin muestras" and t in ("CHIP", "ASERRIN"):
                        obs.append({"sev": ADV, "msg": f"Sin muestras de humedad de {t} en el período: no se descuenta",
                                    "id": "", "rem": "", "bloq": False})
                    if hh["origen"] == "humedad cargada a mano":
                        obs.append({"sev": ADV, "msg": f"Humedad de {t} cargada a mano ({decimal(hh['H']*100, 1)} %)",
                                    "id": "", "rem": "", "bloq": False})
            if sin_gran:
                obs.append({"sev": ADV, "msg": "Descuento de granulometría QUITADO a mano en esta certificación",
                            "id": "", "rem": "", "bloq": False})
            if sin_hum:
                obs.append({"sev": ADV, "msg": "Descuento de humedad QUITADO a mano en esta certificación",
                            "id": "", "rem": "", "bloq": False})
            for _, f in filas[filas["DescManual"]].iterrows() if not sin_gran else []:
                obs.append({"sev": ADV, "msg": f"Descuento de granulometría puesto a mano: {decimal(f['DescPct'], 2)} %",
                            "id": f["_id"], "rem": f["Rto. Planta"], "bloq": False})
            for _, f in filas[filas["PenManual"]].iterrows() if not sin_hum else []:
                obs.append({"sev": ADV, "msg": f"Descuento de humedad puesto a mano: {pct(f['Pen'], 2)}",
                            "id": f["_id"], "rem": f["Rto. Planta"], "bloq": False})
            for _, f in filas[filas["Pen"] > 0].iterrows() if not es_mym else []:
                obs.append({"sev": ADV, "msg": f"Humedad {pct(f['Hum'], 1) if pd.notna(f['Hum']) else '—'} sobre la "
                            f"referencia → −{pct(f['Pen'], 2)} de los kilos", "id": f["_id"], "rem": f["Rto. Planta"],
                            "bloq": False})
            if es_mym:
                for _, f in filas[filas["Producto"].isin(["CHIP", "ASERRIN"])].iterrows():
                    if f["Hum"] is None or pd.isna(f["Hum"]):
                        obs.append({"sev": ADV, "msg": "Camión sin muestra de humedad/granulometría asignada: "
                                    "no se le descuenta (se puede cargar a mano)", "id": f["_id"],
                                    "rem": f["Rto. Planta"], "bloq": False})
                for _, f in filas[filas["FinosOrigen"].str.contains("a mano")].iterrows():
                    obs.append({"sev": ADV, "msg": "Valores de laboratorio cargados a mano", "id": f["_id"],
                                "rem": f["Rto. Planta"], "bloq": False})
            elif ip["desc_gran"]:
                chip = filas[filas["Producto"] == "CHIP"]
                if humedad is not None and len(chip) and chip["Finos"].isna().all():
                    obs.append({"sev": ADV, "msg": "Ningún camión de CHIP con muestra de granulometría en el período",
                                "id": "", "rem": "", "bloq": False})
                for _, f in filas[filas["FinosOrigen"] == "cargado a mano"].iterrows():
                    obs.append({"sev": ADV, "msg": f"Finos cargados a mano: {pct(f['Finos'], 1)}", "id": f["_id"],
                                "rem": f["Rto. Planta"], "bloq": False})
                for _, f in filas[filas["DescPct"] > 0].iterrows():
                    finos_txt = (f"finos {decimal(f['Finos'] * 100, 1)} %" if pd.notna(f["Finos"]) else "sin muestra")
                    obs.append({"sev": ADV, "msg": f"Granulometría ({'a mano' if f.get('DescManual') else 'automático'}): {finos_txt} → "
                                f"{decimal(f['DescPct'], 1)} % ({decimal(f['DescKg'], 1)} kg) pasa a aserrín",
                                "id": f["_id"], "rem": f["Rto. Planta"], "bloq": False})
            for campo, val in (("TOLERANCIA HUMEDAD %", ip["tolerancia"]), ("LIMITE HUMEDAD %", ip["limite"]),
                               ("TOLERANCIA FINOS %", ip["tol_finos"])):
                if val == "ERROR" or (isinstance(val, float) and not (0 <= val < (0.2 if "HUMEDAD" in campo and "TOL" in campo else 1))):
                    obs.append({"sev": CRIT, "msg": f"{campo} del proveedor mal cargado en el config", "id": "",
                                "rem": "", "bloq": True})
            if humedad is not None and isinstance(ip["tolerancia"], float) and ip["tolerancia"] > 0 and ip["desc_hum"]:
                obs.append({"sev": ADV, "msg": f"Tolerancia de humedad pactada: +{decimal(ip['tolerancia'] * 100, 1)} puntos "
                            "sobre el límite", "id": "", "rem": "", "bloq": False})
            if humedad is not None and isinstance(ip["limite"], float) and ip["desc_hum"]:
                obs.append({"sev": ADV, "msg": f"Límite de humedad fijo pactado: {decimal(ip['limite'] * 100, 1)} % "
                            "(reemplaza al de la planilla)", "id": "", "rem": "", "bloq": False})
            for i, rem in zip(filas["_id"], filas["Rto. Planta"]):
                for sev, msg, bloq in por_id.get(i, []):
                    obs.append({"sev": sev, "msg": msg, "id": i, "rem": rem, "bloq": bloq})
            if not ip["cuit"]:
                obs.append({"sev": ADV, "msg": "Falta CUIT del proveedor (sale en blanco: cargalo en el editor)",
                            "id": "", "rem": "", "bloq": False})
            elif not cuit_valido(ip["cuit"]):
                obs.append({"sev": CRIT, "msg": f"CUIT '{ip['cuit']}' inválido (dígito verificador)",
                            "id": "", "rem": "", "bloq": False})
            for t in tipos:
                if precios_c[t] is None:
                    obs.append({"sev": ADV, "msg": f"Falta precio de {t} (sale en blanco: cargalo en el editor o en el "
                                f"Excel de la certificación)", "id": "", "rem": "", "bloq": False})
            for sev, g, msg in obs_cfg:
                if str(g).casefold() == str(ip["grupo"]).casefold():
                    obs.append({"sev": sev, "msg": msg, "id": "", "rem": "", "bloq": False})
            for gl in globales:
                if gl["prov"] in puntos_ent or gl["prov"] == prov:
                    obs.append({"sev": gl["sev"], "msg": gl["msg"], "id": "", "rem": gl["rem"], "bloq": False})
            if abs(tn_grupo - tn_auto) > 0.005:
                obs.append({"sev": ADV, "msg": f"TN del grupo cargado a mano ({decimal(tn_grupo)}); "
                            f"según remitos es {decimal(tn_auto)}", "id": "", "rem": "", "bloq": False})
            for o in obs:                     # nada bloquea: los datos que faltan se completan después
                o["bloq"] = False

            c = {"key": key, "periodo": per, "desde": desde, "hasta": hasta, "proveedor": prov, "miembros": puntos_ent,
                 "razon": ip["razon"], "cuit": ip["cuit"], "grupo": ip["grupo"], "filas": filas,
                 "tn": round(filas["Kilos"].sum() / 1000, 2), "tn_grupo": tn_grupo, "tn_grupo_auto": tn_auto,
                 "precios": precios_c, "origen_precios": origen_p, "obs": obs, "humedad": hum, "tn_prod": tn_prod,
                 "tol_finos": tol_f if tol_f is not None else FINOS_TOLERANCIA,
                 "desc_gran": ip["desc_gran"] and not sin_gran, "sin_gran": sin_gran, "sin_hum": sin_hum,
                 "formato": "MYM" if es_mym else "",
                 "muy_alto": ip["lim_finos"] * 100 if isinstance(ip.get("lim_finos"), float) else FINOS_TOTALIDAD,
                 "con_humedad": humedad is not None, "metodo_humedad": metodo_humedad,
                 "numero": meta.get("numero", ""),
                 "fecha": date.fromisoformat(meta["fecha"]) if meta.get("fecha") else hasta + timedelta(days=1)}
            c["tramos"] = armar_tramos(c, precios)
            if es_mym:
                for ft in c["tramos"]:
                    ft["etiqueta"] = ft["etiqueta"].replace(nombre_producto(ft["producto"]),
                                                            etiqueta_mym(ft["producto"]), 1)
            c["importe"] = importe(c)
            c["estado"] = estado_cert(c, meta)
            certs[key] = c
    globales_sueltos = [g for g in globales if not any(g["prov"] in c["miembros"] for c in certs.values())
                        and (not proveedores or g["prov"] in proveedores)]
    if validacion is not None and solo is None:
        pass
    construir.ultima_validacion = (por_id, globales)
    return certs, globales_sueltos, obs_cfg


def pct(x, n=2):
    return "—" if x is None else decimal(x * 100, n) + " %"


def _txt_limite(hh):
    if hh["L"] is None:
        return "—"
    base = hh["L"] - hh["tol"]
    txt = decimal(base * 100, 1) + " %"
    if hh["tol"]:
        txt += f" + {decimal(hh['tol'] * 100, 1)} = {decimal(hh['L'] * 100, 1)} %"
    return txt


def tn_pagar(c, t):
    """TN a pagar de un producto: después de la reclasificación por granulometría y del descuento de humedad."""
    if c.get("formato") == "MYM":
        return c.get("tn_prod", {}).get(t, 0.0)          # sin redondear, como la planilla de MyM
    h = c.get("humedad", {}).get(t)
    if h:
        return h["tn_pagar"]
    return round(c.get("tn_prod", {}).get(t, 0.0), 2)


def orden_productos(tipos):
    """Tipologías en el orden del cuadro de precios (CHIP, ASERRIN, C. CORTO, C. LARGO, FINO/RALEO, otras)."""
    orden = {t: i for i, t in enumerate(PRODUCTOS_CUADRO + [t for t in TIPOLOGIAS if t not in PRODUCTOS_CUADRO])}
    lista = [t for t, v in tipos.items() if v] if isinstance(tipos, dict) else list(tipos)
    return sorted(lista, key=lambda t: (orden.get(t, 99), str(t)))


def nombre_producto(t):
    base, _, etiqueta = str(t).partition(SEP_VARIANTE)
    return etiqueta or NOMBRE_PRODUCTO.get(base, base)


def _etiqueta_tramo(desde, hasta):
    desde = 0 if pd.isna(desde) else desde
    if desde == 0 and pd.notna(hasta):
        return f"-{hasta:g}TN"
    if desde > 0:
        return f"+{desde:g}TN"
    return ""


def armar_tramos(c, precios):
    """Filas del cuadro de precios como la planilla: CHIP, ASERRIN, C. CORTO, C. LARGO, FINO/RALEO (y lo que se
    haya entregado), con todos los tramos pactados del grupo; las TN van solo en el tramo que corresponde.
    Los precios por punto de entrega (columna PUNTO) van cada uno en su línea, con su ETIQUETA."""
    filas = []
    grupo_p = _de_grupo(precios, c["grupo"])
    del_grupo = grupo_p[grupo_p["PUNTO"] == ""]
    variantes = {}
    for _, r in grupo_p[grupo_p["PUNTO"] != ""].iterrows():
        k = variante_punto(precios, c["grupo"], r["PUNTO"], r["TIPOLOGIA"])
        variantes.setdefault(k, None)
    orden = {t: i for i, t in enumerate(PRODUCTOS_CUADRO + [t for t in TIPOLOGIAS if t not in PRODUCTOS_CUADRO])}
    claves = set(PRODUCTOS_CUADRO) | set(c["precios"]) | set(variantes)
    con_variante = {tipo_base(k) for k in claves if SEP_VARIANTE in k}
    claves = {k for k in claves if SEP_VARIANTE in k or k in c["precios"] or k not in con_variante}
    productos = sorted(claves, key=lambda t: (orden.get(tipo_base(t), 99), SEP_VARIANTE in t, t))
    for t in productos:
        entregado = t in c["precios"]
        tn_t = tn_pagar(c, t) if entregado else None
        if SEP_VARIANTE in t:
            precio = c["precios"][t] if entregado else buscar_precio(precios, c["grupo"], t, c["tn_grupo"])
            filas.append({"producto": t, "etiqueta": nombre_producto(t), "precio": precio, "tn": tn_t})
            continue
        p = del_grupo[del_grupo["TIPOLOGIA"].astype(str).str.casefold() == t.casefold()].sort_values("DESDE_TN")
        manual = entregado and c["origen_precios"].get(t, "").startswith("manual")
        if len(p) <= 1 or manual:
            precio = c["precios"][t] if entregado else (float(p["PRECIO"].iloc[0]) if len(p) else None)
            filas.append({"producto": t, "etiqueta": nombre_producto(t), "precio": precio, "tn": tn_t})
            continue
        for _, r in p.iterrows():
            desde = r["DESDE_TN"] if pd.notna(r["DESDE_TN"]) else 0
            hasta = r["HASTA_TN"] if pd.notna(r["HASTA_TN"]) else float("inf")
            aplica = entregado and desde <= c["tn_grupo"] < hasta
            filas.append({"producto": t, "etiqueta": f"{nombre_producto(t)} {_etiqueta_tramo(r['DESDE_TN'], r['HASTA_TN'])}",
                          "precio": float(r["PRECIO"]) if pd.notna(r["PRECIO"]) else None,
                          "tn": tn_t if aplica else None})
    return filas


def con_tramos(c):
    return any("TN" in f["etiqueta"].split()[-1] for f in c.get("tramos", []))


def mostrar_humedad(c):
    """El cuadro de humedad se muestra solo si al proveedor se le descuenta humedad y hay algo para informar."""
    return c.get("con_humedad") and any(hh["aplica_hum"] and (hh["muestras"] or hh["desc_hum"])
                                        for hh in c["humedad"].values())


def importe(c):
    """Importe a facturar; un producto sin precio cuenta 0 (se completa en el editor o en el Excel)."""
    tot = 0.0
    for t, p in c["precios"].items():
        tot += tn_pagar(c, t) * (p or 0.0)
    return round(tot, 2)


def huella(c):
    """Huella del contenido: si algo cambia después de confirmar, se detecta."""
    extra = ["DescKg"] + (["Hum", "Finos", "Pen"] if c.get("formato") == "MYM" else [])
    f = c["filas"][FILA_COLS + extra].astype(str).values.tolist()
    data = [c["razon"], c["cuit"], c["numero"], str(c["fecha"]), round(c["tn_grupo"], 2),
            sorted((k, v) for k, v in c["precios"].items()), f,
            sorted((k, round(v["desc_hum"], 6), round(v["desc_gran"], 6)) for k, v in c.get("humedad", {}).items())]
    return hashlib.md5(json.dumps(data, default=str).encode()).hexdigest()


def criticas_revisables(c):
    return sorted({o["msg"] for o in c["obs"] if o["sev"] == CRIT and not o["bloq"]})


def estado_cert(c, meta):
    if meta.get("huella"):
        return "CONFIRMADA" if meta["huella"] == huella(c) else "MODIFICADA"
    if any(o["bloq"] for o in c["obs"]):
        return "BLOQUEADA"
    pend = set(criticas_revisables(c)) - set(meta.get("ack", {}).get("msgs", []))
    return "REVISAR" if pend else "LISTA"


ICONOS = {"CONFIRMADA": "✅", "MODIFICADA": "🔁", "BLOQUEADA": "⛔", "REVISAR": "⚠️", "LISTA": "🟢"}


def cuadre(d_orig, d, certs, conf, anio, mes, proveedores=None):
    ini = pd.Timestamp(anio, mes, 1)
    fin = ini + pd.offsets.MonthEnd(0)
    filas = []
    for per in sorted({c["periodo"] for c in certs.values()}, key=["1Q", "2Q", "Mensual"].index):
        def base(x):
            x = x[(x[C_DIA] >= ini) & (x[C_DIA] <= fin)]
            if per == "1Q":
                x = x[x[C_DIA].dt.day <= DIA_CORTE]
            elif per == "2Q":
                x = x[x[C_DIA].dt.day > DIA_CORTE]
            return x[x[C_PROV].isin(proveedores)] if proveedores else x
        bo, bc = base(d_orig), base(d)
        provs = bc[C_PROV].unique()            # una vez por proveedor (no por remito): mucho más rápido
        m_cert = {p: info_proveedor(conf, p)["certifica"] for p in provs}
        m_corr = {p: corresponde(conf, unidad(conf, p), per) for p in provs}
        certifica = bc[C_PROV].map(m_cert).astype(bool)
        excl = bc[~certifica]
        otra_f = bc[certifica & ~bc[C_PROV].map(m_corr).astype(bool)]
        tn_o, tn_c = bo[C_NETO].sum() / 1000, bc[C_NETO].sum() / 1000
        tn_cert = sum(c["filas"]["Kilos"].sum() for c in certs.values() if c["periodo"] == per) / 1000
        tn_e = excl[C_NETO].sum() / 1000
        tn_f = otra_f[C_NETO].sum() / 1000
        dif = round(tn_c - tn_cert - tn_e - tn_f, 2)
        filas.append({"Período": per, "TN maestro": round(tn_o, 2),
                      "Ajustes manuales": round(tn_c - tn_o, 2), "TN maestro corregido": round(tn_c, 2),
                      "TN certificadas": round(tn_cert, 2), "TN excluidas (CERTIFICAR=NO)": round(tn_e, 2),
                      "TN de otra frecuencia": round(tn_f, 2),
                      "Diferencia": dif, "Resultado": "✔ CUADRA" if abs(dif) < 0.01 else "✖ NO CUADRA"})
    return pd.DataFrame(filas)


# -----------------------------------------------------------------------------
# 6. Formatos de número (estilo argentino)
# -----------------------------------------------------------------------------
def _ar(s):
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def miles(x):
    return _ar(f"{float(x):,.0f}")


def decimal(x, n=2):
    return _ar(f"{float(x):,.{n}f}")


def pesos(x):
    return "$ " + decimal(x) if x is not None else "—"


def texto_periodo(desde, hasta):
    if desde.day == 1 and hasta.day == calendar.monthrange(hasta.year, hasta.month)[1]:
        return f"Período: {MESES[desde.month]} {desde.year} (del 01 al {hasta.day:02d})"
    return f"Período del {desde.day} al {hasta.day} de {MESES[desde.month]} {desde.year}"


def etiqueta_periodo(per):
    return {"1Q": "Quincenal", "2Q": "Quincenal", "Mensual": "Mensual"}[per]


# -----------------------------------------------------------------------------
# 7. PDF de la certificación
# -----------------------------------------------------------------------------
def kilos_txt(x):
    """16.400 · 11.304,5 (como la planilla)."""
    x = float(x)
    return miles(x) if abs(x - round(x)) < 1e-6 else decimal(x, 1)


def _con_desc(f):
    """El camión tiene descuento (granulometría o, en MyM, humedad)."""
    pen = f.get("Pen", 0) if hasattr(f, "get") else 0
    return f["DescKg"] > 0 or (pen is not None and not pd.isna(pen) and pen > 0)


_ETIQUETA_PDF = re.compile(r"<(?!/?(?:b|i|u|br|font|super|sub|strike)\b)")
_AMP_PDF = re.compile(r"&(?!(?:[a-zA-Z]+|#\d+);)")


def _texto_pdf(t):
    """Texto para el PDF: un '<' o '&' que venga de un nombre o remito (ej. 'A <B>' o 'Pérez & Hijos') se muestra
    tal cual en vez de romper el PDF; las marcas de formato propias (<b>, <br/>…) se respetan."""
    t = "" if t is None else str(t)
    return _ETIQUETA_PDF.sub("&lt;", _AMP_PDF.sub("&amp;", t))


def texto_granulometria(f):
    partes = []
    if f["DescKg"] > 0:
        if f.get("DescManual"):
            partes.append(f"Desc. granulometría a mano {decimal(f['DescPct'], 1)}%")
        elif f["Finos"] is not None and not pd.isna(f["Finos"]):
            finos = f["Finos"] * 100
            partes.append(f"Según granulometría CHIP {decimal(100 - finos, 1)} // ASE {decimal(finos, 1)}% "
                          f"(Desc. {decimal(f['DescPct'], 1)}%)")
        else:
            partes.append(f"Desc. granulometría {decimal(f['DescPct'], 1)}%")
    pen = f.get("Pen", 0)
    if pen is not None and not pd.isna(pen) and pen > 0:
        if f.get("PenManual") or f.get("Hum") is None or pd.isna(f.get("Hum")):
            partes.append(f"Desc. humedad {decimal(pen * 100, 1)}%")
        else:
            partes.append(f"Humedad {decimal(f['Hum'] * 100, 1)}% (ref. {decimal(f['HumRef'] * 100, 1)}%) "
                          f"Desc. {decimal(pen * 100, 1)}%")
    return " · ".join(partes)


def nota_granulometria(c):
    t = (f"Granulometría (finos &lt; #04, por camión): hasta {decimal(c['tol_finos'], 0)} % sin descuento; "
         f"de {decimal(c['tol_finos'], 0)} % a {decimal(c.get('muy_alto', FINOS_TOTALIDAD), 0)} % se descuenta la diferencia; "
         f"desde {decimal(c.get('muy_alto', FINOS_TOTALIDAD), 0)} % la totalidad. Lo descontado se paga como aserrín.")
    if c.get("formato") == "MYM" and not c.get("sin_hum"):
        t += (" Humedad (por camión): si supera la referencia del día se descuenta el porcentaje con el que la "
              "supera, sobre los kilos de chip y aserrín.")
    return t


def etiqueta_producto(c):
    return etiqueta_mym if c.get("formato") == "MYM" else nombre_producto


def tn_detalle(c):
    """TN de la columna Kilos del detalle (después del descuento de humedad por camión, si lo hay)."""
    fl = c["filas"]
    return round((fl["KilosNeto"].sum() + fl["DescKg"].sum()) / 1000, 2)


def pdf_certificacion(ruta, c, logo=None):
    if c.get("formato") == "MYM":          # MyM tiene su propio formato (planilla de MyM)
        return pdf_mym(ruta, c, logo)
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.lib.utils import ImageReader
    from reportlab.platypus import Image, KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    verde, amarillo, gris = colors.HexColor("#92D050"), colors.HexColor("#FFFF00"), colors.HexColor("#808080")
    amarillo_claro, azul, rojo = colors.HexColor("#FFF2A8"), colors.HexColor("#2E75B6"), colors.HexColor("#C00000")

    def P(t, size=9, bold=False, align=0, color=colors.black):
        return Paragraph(_texto_pdf(t), ParagraphStyle("p", fontName="Helvetica-Bold" if bold else "Helvetica",
                                           fontSize=size, leading=size * 1.25, alignment=align, textColor=color))

    ancho = A4[0] - 24 * mm
    el = []
    if logo:
        iw, ih = ImageReader(io.BytesIO(logo)).getSize()
        img = Image(io.BytesIO(logo), width=32 * mm, height=32 * mm * ih / iw)
        img.hAlign = "LEFT"
        el.append(img)
    el += [P(l, 9) for l in EMPRESA_LINEAS]
    num = Table([["", P(f"Nº {c['numero']}", 10, True, 1)]], colWidths=[ancho - 55 * mm, 55 * mm])
    num.setStyle(TableStyle([("BOX", (1, 0), (1, 0), 1.2, colors.black)]))
    el += [Spacer(1, 1 * mm), num, Spacer(1, 2 * mm), P("CERTIFICACIÓN ENTREGA BIOMASA", 12, True, 1, azul),
           Spacer(1, 3 * mm)]
    enc = Table([[P(texto_periodo(c["desde"], c["hasta"]), 10), P(f"Fecha: {c['fecha']:%d/%m/%Y}", 10, True, 1)],
                 [P(f"Proveedor: {c['razon']}", 10), ""],
                 [P(f"CUIT: {c['cuit']}", 10), ""]], colWidths=[ancho - 55 * mm, 55 * mm])
    enc.setStyle(TableStyle([("LINEBELOW", (0, 0), (0, 2), 1, colors.black),
                             ("BOX", (1, 0), (1, 0), 1.2, colors.black),
                             ("RIGHTPADDING", (0, 0), (0, -1), 30 * mm)]))
    el += [enc, Spacer(1, 4 * mm)]

    # Detalle como la planilla, con el % de humedad y de granulometría (finos) de cada camión:
    # Fecha | N° Rto Planta | Proveedor | Rto.Proveedor | Producto | Humedad % | Granulom. % | Kilos | KG TOTAL
    cols = ["Fecha", "N° Rto Planta", "Proveedor", "Rto. Prov.", "Producto", "Humedad %", "Granulom. %",
            "Kilos", "KG TOTAL"]
    data = [[P(x, 7.5, True, 1) for x in cols]]
    estilo = [("GRID", (0, 0), (-1, -2), 0.5, colors.black), ("BOX", (0, 0), (-1, 0), 1.2, colors.black),
              ("ALIGN", (0, 1), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
              ("FONTSIZE", (0, 1), (-1, -1), 8),
              ("TOPPADDING", (0, 1), (-1, -1), 1.2), ("BOTTOMPADDING", (0, 1), (-1, -1), 1.2)]
    etq = etiqueta_producto(c)
    pct_txt = lambda v: "" if v is None or pd.isna(v) else decimal(v * 100, 1) + " %"
    for _, f in c["filas"].iterrows():
        con_desc = _con_desc(f)
        data.append([f["Fecha"].strftime("%d/%m/%Y"), f["Rto. Planta"], P(f["Proveedor"], 7, align=1),
                     P(str(f["Remito Proveedor"]), 6.5 if len(str(f["Remito Proveedor"])) > 11 else 7.5, align=1), etq(f["Producto"]),
                     pct_txt(f.get("Hum")), pct_txt(f.get("Finos")),
                     kilos_txt(f["KilosNeto"]), kilos_txt(f["Kilos"]) if con_desc else ""])
        if con_desc:
            i = len(data) - 1
            data.append([P(texto_granulometria(f), 7.5), "", "", "", "ASERRÍN" if f["DescKg"] > 0 else "",
                         "", "", kilos_txt(f["DescKg"]) if f["DescKg"] > 0 else "", ""])
            estilo += [("SPAN", (0, i + 1), (3, i + 1)), ("BACKGROUND", (0, i), (-1, i + 1), amarillo_claro)]
    data.append(["", "", "", "", "", P("TOTAL TN", 9, True, 2), "", P(decimal(tn_detalle(c)), 9, True, 1), ""])
    estilo += [("BOX", (7, -1), (7, -1), 1.2, colors.black), ("SPAN", (5, -1), (6, -1))]
    w = [18 * mm, 22 * mm, 33 * mm, 22 * mm, 18 * mm, 18 * mm, 19 * mm, 19 * mm]
    w.append(ancho - sum(w))
    det = Table(data, colWidths=w, repeatRows=1)
    det.setStyle(TableStyle(estilo))
    el += [det]
    if c.get("desc_gran"):
        el += [Spacer(1, 1 * mm), P(nota_granulometria(c), 7)]
    el.append(Spacer(1, 4 * mm))

    bloque = []
    if con_tramos(c):
        tg = Table([[P("TOTAL DE TN ENTREGADAS EN EL MES DE TODOS LOS PUNTOS DE ENTREGA  "
                       f"{decimal(c['tn_grupo'])}", 10.5, True, 1)]], colWidths=[ancho])
        tg.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), verde)]))
        bloque += [tg, Spacer(1, 3 * mm)]
    if mostrar_humedad(c):
        hdata = [[P(x, 8, True, 1) for x in ["Producto", "Muestras", "Humedad prom.", "Límite", "Descuento",
                                              "TN", "TN a pagar"]]]
        hst = [("GRID", (0, 0), (-1, -1), 0.5, colors.black), ("ALIGN", (0, 0), (-1, -1), "CENTER"),
               ("FONTSIZE", (0, 1), (-1, -1), 8), ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E2EFDA"))]
        for t, hh in c["humedad"].items():
            if not hh["aplica_hum"]:
                continue
            hdata.append([nombre_producto(t), str(hh["muestras"] or "—"), pct(hh["H"], 1), _txt_limite(hh),
                          pct(hh["desc_hum"]), decimal(hh["tn"]), decimal(hh["tn_pagar"])])
            if hh["desc_hum"] > 0:
                hst.append(("BACKGROUND", (0, len(hdata) - 1), (-1, len(hdata) - 1), amarillo_claro))
        ht = Table(hdata, colWidths=[24 * mm, 18 * mm, 24 * mm, 36 * mm, 20 * mm, 20 * mm, 22 * mm])
        ht.setStyle(TableStyle(hst))
        bloque += [P("DESCUENTO POR HUMEDAD", 9.5, True), Spacer(1, 1 * mm), ht,
                   P("Humedad: promedio de las muestras del laboratorio en el período; límite según tipo de proveedor "
                     "más la tolerancia pactada.", 7), Spacer(1, 3 * mm)]
    pdata = [[P(x, 9, True, 1) for x in ["", "TN", "PRECIO", "TOTAL"]]]
    for f in c["tramos"]:
        pdata.append([f["etiqueta"], decimal(f["tn"]) if f["tn"] is not None else "", pesos(f["precio"]) if f["precio"] else "",
                      pesos(f["tn"] * f["precio"]) if (f["tn"] is not None and f["precio"] is not None) else ""])
    pdata.append(["", P(decimal(sum(tn_pagar(c, t) for t in c["precios"])), 9, True, 1), "",
                  P(pesos(c["importe"]), 11, True, 1)])
    pt = Table(pdata, colWidths=[40 * mm, 26 * mm, 34 * mm, 40 * mm])
    pt.setStyle(TableStyle([("GRID", (0, 0), (-1, -2), 0.5, colors.black),
                            ("BOX", (0, 0), (-1, -2), 1.2, colors.black),
                            ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                            ("FONTSIZE", (0, 1), (-1, -1), 9),
                            ("BOX", (1, -1), (1, -1), 1.2, colors.black),
                            ("BACKGROUND", (3, -1), (3, -1), amarillo), ("BOX", (3, -1), (3, -1), 1.2, colors.black)]))
    fact_txt = Table([[P("FACTURAR", 10, True, 1)]], colWidths=[26 * mm])
    fact_txt.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), amarillo)]))
    fact = Table([[pt, fact_txt]], colWidths=[146 * mm, ancho - 146 * mm])
    fact.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "BOTTOM"), ("LEFTPADDING", (0, 0), (0, 0), 12)]))
    firmas = Table([[[P("Aceptación Proveedor", 9, True)], "", [P("Firma Empresa", 9, True)]]],
                   colWidths=[70 * mm, ancho - 140 * mm, 70 * mm], rowHeights=[16 * mm])
    firmas.setStyle(TableStyle([("BOX", (0, 0), (0, 0), 1, colors.black), ("BOX", (2, 0), (2, 0), 1, colors.black),
                                ("VALIGN", (0, 0), (-1, -1), "TOP")]))
    # Cuadro de precios, total a facturar y firmas siempre juntos (si no entran, pasan a la hoja siguiente)
    el.append(KeepTogether(bloque + [fact, Spacer(1, 6 * mm), firmas, Spacer(1, 3 * mm)]
                           + [P(t, 7, color=rojo) for t in PIE]))

    def pie(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(gris)
        canvas.drawString(12 * mm, 8 * mm, f"Certificación Nº {c['numero']} · {c['razon']} · "
                                           f"{texto_periodo(c['desde'], c['hasta'])}")
        canvas.drawRightString(A4[0] - 12 * mm, 8 * mm, f"Página {doc.page}")
        canvas.restoreState()

    doc = SimpleDocTemplate(str(ruta), pagesize=A4, leftMargin=12 * mm, rightMargin=12 * mm,
                            topMargin=9 * mm, bottomMargin=14 * mm,
                            title=f"Certificación {c['numero']} - {c['razon']}", author=EMPRESA_LINEAS[0])
    doc.build(el, onFirstPage=pie, onLaterPages=pie)


# -----------------------------------------------------------------------------
# 8. Libro Excel único por período (una hoja por proveedor)
# -----------------------------------------------------------------------------
VERDE = PatternFill("solid", fgColor="92D050")
AMARILLO = PatternFill("solid", fgColor="FFFF00")
NARANJA = PatternFill("solid", fgColor="FCE4D6")
ROJO = PatternFill("solid", fgColor="C00000")
ROJO_CLARO = PatternFill("solid", fgColor="F8CBAD")
NARANJA_DESC = PatternFill("solid", fgColor="FFE0B2")     # filas con descuento aplicado
VERDE_CLARO = PatternFill("solid", fgColor="C6EFCE")
ENCAB = PatternFill("solid", fgColor="3A7D44")
FINO, MEDIO = Side(style="thin"), Side(style="medium")
B_FINO = Border(left=FINO, right=FINO, top=FINO, bottom=FINO)
B_MEDIO = Border(left=MEDIO, right=MEDIO, top=MEDIO, bottom=MEDIO)
MONEDA = '"$"\\ #,##0.00'
CENTRO = Alignment(horizontal="center", vertical="center")


def _nombre_hoja(prov, usados):
    base = re.sub(r"[\[\]\*\?/\\:]", "", prov)[:31].strip() or "Proveedor"
    n, k = base, 2
    while n.casefold() in usados:
        n = f"{base[:28]}~{k}"
        k += 1
    usados.add(n.casefold())
    return n


AMARILLO_CLARO = PatternFill("solid", fgColor="FFF2A8")    # camiones con descuento de granulometría


def _hoja_certificacion(ws, c, logo=None):
    """Formato de la certificación final (todos los proveedores menos MyM), editable a mano en el Excel:
    cada camión con su humedad % y granulometría % del laboratorio y dos celdas celestes editables
    (Desc. humedad % y Desc. granulometría %); los kilos, el aserrín, las TN y el importe se recalculan.
    Los precios también son celdas celestes editables."""
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.comments import Comment
    if c.get("formato") == "MYM":          # MyM tiene su propio formato (planilla de MyM)
        return _hoja_mym(ws, c, logo)
    EDIT = PatternFill("solid", fgColor="DDEBF7")       # celdas para completar / corregir a mano
    for col, a in zip("ABCDEFGHIJKLM", [2, 11, 14, 22, 14, 16, 11, 10, 10, 11, 11, 12, 11]):
        ws.column_dimensions[col].width = a
    for col in "NO":                       # línea del cuadro de precios de cada camión (uso interno de las fórmulas)
        ws.column_dimensions[col].hidden = True
    if logo:
        img = XLImage(io.BytesIO(logo))
        img.width, img.height = 170, 90
        ws.add_image(img, "B1")
    if c["estado"] != "CONFIRMADA":
        ws.merge_cells("G1:M3")
        txt = {"MODIFICADA": "MODIFICADA DESPUÉS DE CONFIRMAR — volver a confirmar",
               "BLOQUEADA": "BORRADOR — faltan datos",
               "REVISAR": "BORRADOR — observaciones sin revisar",
               "LISTA": "BORRADOR — lista para confirmar"}[c["estado"]]
        ws["G1"] = txt
        ws["G1"].font = Font(bold=True, color="FFFFFF", size=11)
        ws["G1"].fill = ROJO
        ws["G1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, l in enumerate(EMPRESA_LINEAS, start=6):
        ws.cell(i, 2, l).font = Font(size=10)
    ws.merge_cells("K7:M7")
    ws["K7"] = f"Nº {c['numero'] or '(se asigna al confirmar)'}"
    ws["K7"].font, ws["K7"].alignment = Font(bold=True), CENTRO
    for col in "KLM":
        ws[f"{col}7"].border = B_MEDIO
    ws.merge_cells("B9:M9")
    ws["B9"] = "CERTIFICACIÓN ENTREGA BIOMASA"
    ws["B9"].font, ws["B9"].alignment = Font(bold=True, size=12, color="2E75B6"), CENTRO
    for r, t in [(11, texto_periodo(c["desde"], c["hasta"])), (12, f"Proveedor: {c['razon']}"),
                 (13, f"CUIT: {c['cuit'] or ''}")]:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        ws.cell(r, 2, t).border = Border(bottom=MEDIO)
    if not c["cuit"]:
        ws["B13"].fill = EDIT
        ws["B13"].comment = Comment("Completá el CUIT del proveedor", "Certificaciones")
    ws.merge_cells("K11:M11")
    ws["K11"] = f"Fecha: {c['fecha']:%d/%m/%Y}"
    ws["K11"].font, ws["K11"].alignment = Font(bold=True), CENTRO
    for col in "KLM":
        ws[f"{col}11"].border = B_MEDIO
    cols = ["Fecha", "N° Rto Planta", "Proveedor", "Rto.Proveedor", "Producto", "KG TOTAL", "Humedad %",
            "Granulom. %", "Desc. humedad %", "Desc. granulom. %", "Kilos", "Aserrín (kg)"]
    for j, t in enumerate(cols, 2):
        x = ws.cell(15, j, t)
        x.font = Font(bold=True)
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        x.border = B_MEDIO
        if j in (10, 11):
            x.fill = EDIT
    ws.row_dimensions[15].height = 44
    ws["J15"].comment = Comment("Editable: % de los kilos que se descuenta por humedad (0 = sin descuento)",
                                "Certificaciones")
    ws["K15"].comment = Comment("Editable: % de los kilos que pasa a aserrín por granulometría (0 = sin descuento)",
                                "Certificaciones")
    sev = {}
    for o in c["obs"]:
        if o["id"] and o["sev"] == CRIT:
            sev.setdefault(o["id"], set()).add(o["sev"])
    r = 16
    for _, f in c["filas"].iterrows():
        rp = f["Remito Proveedor"]
        vals = [f["Fecha"], f["Rto. Planta"], f["Proveedor"], int(rp) if re.fullmatch(r"\d{1,9}", rp or "") else rp,
                f["Producto"], float(f["Kilos"]),
                None if pd.isna(f.get("Hum")) else float(f["Hum"]),
                None if pd.isna(f.get("Finos")) else float(f["Finos"]),
                float(f.get("Pen") or 0.0), float(f.get("DescPct") or 0.0) / 100]
        for j, v in enumerate(vals, 2):
            ws.cell(r, j, v)
        # como la planilla: aserrín = KG TOTAL × (1 − desc. humedad) × desc. granulometría; kilos = el resto
        ws.cell(r, 13, f"=ROUND(G{r}*(1-J{r})*K{r},2)")
        ws.cell(r, 12, f"=G{r}*(1-J{r})-M{r}")
        # línea del cuadro de precios de este camión (columnas ocultas): producto o precio de su punto de entrega
        ws.cell(r, 14, f.get("Cuadro") if isinstance(f.get("Cuadro"), str) else f["Producto"])
        ws.cell(r, 15, f.get("CuadroAse") if isinstance(f.get("CuadroAse"), str) else "ASERRIN")
        ws.cell(r, 2).number_format = "dd/mm/yyyy"
        ws.cell(r, 7).number_format = "#,##0"
        for j in (8, 9, 10, 11):
            ws.cell(r, j).number_format = "0.0%"
        for j in (12, 13):
            ws.cell(r, j).number_format = "#,##0.0"
        for j in range(2, 14):
            x = ws.cell(r, j)
            x.border, x.alignment = B_FINO, CENTRO
            if CRIT in sev.get(f["_id"], set()) and j < 10:
                x.fill = ROJO_CLARO
        ws.cell(r, 10).fill = ws.cell(r, 11).fill = EDIT
        r += 1
    ult = r - 1
    if ult >= 16:
        # fila en amarillo si el camión tiene algún descuento (también si lo cargás a mano en el Excel)
        ws.conditional_formatting.add(f"B16:I{ult}", FormulaRule(formula=["OR($J16>0,$K16>0)"], fill=AMARILLO_CLARO))
        ws.conditional_formatting.add(f"L16:M{ult}", FormulaRule(formula=["OR($J16>0,$K16>0)"], fill=AMARILLO_CLARO))
    ws.cell(r, 6, "TOTAL").font = Font(bold=True)
    for j, f_ in ((7, f"=SUM(G16:G{ult})/1000"), (12, f"=SUM(L16:L{ult})/1000"), (13, f"=SUM(M16:M{ult})/1000")):
        x = ws.cell(r, j, f_)
        x.number_format, x.font, x.border, x.alignment = "0.00", Font(bold=True), B_MEDIO, CENTRO
    ws.cell(r, 8, "TN").font = Font(bold=True)
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=13)
    ws.row_dimensions[r].height = 36
    nota = ("Celdas celestes: se pueden corregir a mano (Desc. humedad %, Desc. granulom. % y precios) y todo se "
            "recalcula. Humedad % y Granulom. %: muestra del laboratorio de cada camión. ")
    if c.get("desc_gran"):
        nota += nota_granulometria(c).replace("&lt;", "<")
    ws.cell(r, 2, nota).font = Font(size=8, italic=True)
    ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")
    r += 2
    if con_tramos(c):
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=13)
        x = ws.cell(r, 2, "TOTAL DE TN ENTREGADAS EN EL MES DE TODOS LOS PUNTOS DE ENTREGA  " + decimal(c["tn_grupo"]))
        x.font, x.fill, x.alignment = Font(bold=True, size=12), VERDE, CENTRO
        r += 2
    for j, t in zip(range(4, 7), ["TN", "PRECIO", "TOTAL"]):
        x = ws.cell(r, j, t)
        x.font, x.alignment, x.border = Font(bold=True), CENTRO, B_MEDIO
    ws.cell(r, 3).border = B_MEDIO
    ini = r + 1
    rng = lambda col: f"${col}$16:${col}${ult}"
    for f in c["tramos"]:
        r += 1
        ws.cell(r, 3, f["etiqueta"])
        t = f["producto"]
        tramo_sin_tn = f["tn"] is None and t in c["precios"]          # otro tramo del mismo producto
        if not tramo_sin_tn:
            ws.cell(r, 4, f'=ROUND((SUMIF({rng("N")},"{t}",{rng("L")})+SUMIF({rng("O")},"{t}",{rng("M")}))/1000,2)')
        ws.cell(r, 5, f["precio"])
        ws.cell(r, 6, f'=IF(OR(D{r}="",E{r}=""),"",D{r}*E{r})')
        ws.cell(r, 4).number_format = "0.00"
        ws.cell(r, 5).number_format = ws.cell(r, 6).number_format = MONEDA
        for j in range(3, 7):
            ws.cell(r, j).border, ws.cell(r, j).alignment = B_FINO, CENTRO
        ws.cell(r, 5).fill = ROJO_CLARO if (t in c["precios"] and f["precio"] is None) else EDIT
    r += 1
    x = ws.cell(r, 4, f"=SUM(D{ini}:D{r-1})")
    x.number_format, x.font, x.border, x.alignment = "0.00", Font(bold=True), B_MEDIO, CENTRO
    r += 2
    ws.cell(r, 4, "FACTURAR").font = Font(bold=True)
    ws.cell(r, 4).fill = AMARILLO
    x = ws.cell(r, 6, f"=SUM(F{ini}:F{r-3})")
    x.number_format, x.font, x.fill, x.border = MONEDA, Font(bold=True, size=13), AMARILLO, B_MEDIO
    # Firmas en recuadros
    r += 3
    for col_ini, col_fin, texto in ((2, 5, "Aceptación Proveedor"), (9, 13, "Firma Empresa")):
        ws.merge_cells(start_row=r, start_column=col_ini, end_row=r + 3, end_column=col_fin)
        x = ws.cell(r, col_ini, texto)
        x.font, x.alignment = Font(bold=True), Alignment(vertical="top")
        for rr in range(r, r + 4):
            for cc in range(col_ini, col_fin + 1):
                ws.cell(rr, cc).border = Border(left=MEDIO if cc == col_ini else None,
                                                right=MEDIO if cc == col_fin else None,
                                                top=MEDIO if rr == r else None, bottom=MEDIO if rr == r + 3 else None)
    r += 6
    for t in PIE:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=13)
        x = ws.cell(r, 2, t)
        x.font, x.alignment = Font(size=8, color="C00000"), Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 20
        r += 1
    ws.freeze_panes = "B16"
    ws.print_title_rows = "15:15"
    ws.print_area = f"A1:M{r}"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.4, bottom=0.4)
    ws.sheet_properties.tabColor = {"CONFIRMADA": "3A7D44", "LISTA": "92D050"}.get(c["estado"], "C00000")


# -----------------------------------------------------------------------------
# 7b. Formato MyM (humedad y granulometría por camión)
# -----------------------------------------------------------------------------
MYM_PRODUCTOS = ["CHIP", "ASERRIN", "C CORTO", "C LARGO", "Fino/Raleo"]
MYM_COL_KG = {"CHIP": "KgChip", "ASERRIN": "KgAse", "C CORTO": "KgCorto", "C LARGO": "KgLargo", "Fino/Raleo": "KgRaleo"}


def etiqueta_mym(t):
    return "ROLLO" if t == "Fino/Raleo" else nombre_producto(t)


def criterio_mym(c):
    alto, muy = decimal(c["tol_finos"], 0), decimal(c["muy_alto"], 0)
    return (f'Cuando el contenido de aserrín es superior al valor de "ALTO" ({alto} %) e inferior al valor de '
            f'"MUY ALTO" ({muy} %) se cuentan como aserrín los kilos que superen el valor de "ALTO". En caso de '
            f'igualar o superar el valor de "MUY ALTO" se cuenta como aserrín el total medido.',
            "Cuando la humedad excede el valor de referencia del día (chipeado propio de raleo fresco) se descuenta "
            "el porcentaje con el cual supera ese valor.")


def _pct_txt(v):
    return "" if v is None or pd.isna(v) else f"{v * 100:.2f}%".replace(".", ",")


def _kg2(v):
    return "" if v is None or pd.isna(v) else decimal(v, 2)


def pdf_mym(ruta, c, logo=None):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.lib.utils import ImageReader
    from reportlab.platypus import Image, KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    verde, amarillo, naranja = colors.HexColor("#92D050"), colors.HexColor("#FFFF00"), colors.HexColor("#FFC000")
    azul, rojo, gris, celeste = colors.HexColor("#2E75B6"), colors.HexColor("#C00000"), colors.HexColor("#808080"), colors.HexColor("#DDEBF7")
    pag = landscape(A4)
    ancho = pag[0] - 20 * mm

    def P(t, size=8, bold=False, align=0, color=colors.black):
        return Paragraph(_texto_pdf(t), ParagraphStyle("p", fontName="Helvetica-Bold" if bold else "Helvetica",
                                           fontSize=size, leading=size * 1.2, alignment=align, textColor=color))

    el = []
    img = ""
    if logo:
        iw, ih = ImageReader(io.BytesIO(logo)).getSize()
        img = Image(io.BytesIO(logo), width=28 * mm, height=28 * mm * ih / iw)
    emp = [P(l, 8) for l in EMPRESA_LINEAS]
    caja = Table([["Fecha:", c["fecha"].strftime("%d/%m/%Y")], ["N°:", c["numero"] or "(al confirmar)"]],
                 colWidths=[16 * mm, 38 * mm])
    caja.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black), ("FONTSIZE", (0, 0), (-1, -1), 9),
                              ("FONTNAME", (1, 1), (1, 1), "Helvetica-Bold"), ("ALIGN", (1, 0), (1, -1), "CENTER")]))
    cab = Table([[img, emp, P("CERTIFICACIÓN ENTREGA BIOMASA", 13, True, 1, azul), caja]],
                colWidths=[32 * mm, 72 * mm, ancho - 32 * mm - 72 * mm - 58 * mm, 58 * mm])
    cab.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    el += [cab, Spacer(1, 2 * mm)]

    g_txt, h_txt = criterio_mym(c)
    datos = Table([["Período:", texto_periodo(c["desde"], c["hasta"]).replace("Período ", "")],
                   ["Proveedor:", c["razon"]], ["CUIT:", c["cuit"] or ""]], colWidths=[20 * mm, 72 * mm])
    datos.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black), ("FONTSIZE", (0, 0), (-1, -1), 8.5),
                               ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold")]))
    crit = Table([[P("Criterio de evaluación", 8.5, True, 1), "", "", ""],
                  [P("Granulometría:", 7.5, True), P("Alto", 8, True, 1), P("Muy Alto", 8, True, 1), P(g_txt, 6.5)],
                  ["", P(f"{decimal(c['tol_finos'], 0)}%", 8, align=1), P(f"{decimal(c['muy_alto'], 0)}%", 8, align=1), ""],
                  [P("Humedad:", 8, True), P("Referencias extraídas del chipeado propio de raleo fresco", 7), "", P(h_txt, 6.5)]],
                 colWidths=[25 * mm, 22 * mm, 22 * mm, ancho - 92 * mm - 6 * mm - 69 * mm])
    crit.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black), ("SPAN", (0, 0), (-1, 0)),
                              ("SPAN", (0, 1), (0, 2)), ("SPAN", (3, 1), (3, 2)), ("SPAN", (1, 3), (2, 3)),
                              ("BACKGROUND", (1, 2), (2, 2), amarillo), ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    fila = Table([[datos, "", crit]], colWidths=[92 * mm, 6 * mm, ancho - 98 * mm])
    fila.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    el += [fila, Spacer(1, 3 * mm)]

    cols = ["Fecha", "N° Rto Planta", "Rto.Proveedor", "Producto", "Kilos de carga", "Humedad medida",
            "Humedad de referencia", "Porcentaje de penalización", "Porcentaje medido de aserrín", "Kilos de chip",
            "Kilos de aserrín", "Kilos de C Corto", "Kilos C Largo", "Kilos Rollo"]
    data = [[P(x, 6.8, True, 1) for x in cols]]
    est = [("GRID", (0, 0), (-1, -1), 0.4, colors.black), ("ALIGN", (0, 1), (-1, -1), "CENTER"),
           ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("FONTSIZE", (0, 1), (-1, -1), 7),
           ("TOPPADDING", (0, 1), (-1, -1), 0.6), ("BOTTOMPADDING", (0, 1), (-1, -1), 1.6), ("LEADING", (0, 1), (-1, -1), 8.4),
           ("BACKGROUND", (0, 0), (-1, 0), celeste)]
    alto, muy = c["tol_finos"] / 100, c["muy_alto"] / 100
    for _, f in c["filas"].iterrows():
        prod = f["Producto"]
        cs = prod in ("CHIP", "ASERRIN")
        i = len(data)
        data.append([f["Fecha"].strftime("%d-%m-%y"), f["Rto. Planta"], f["Remito Proveedor"], etiqueta_mym(prod),
                     decimal(f["Kilos"], 2), _pct_txt(f["Hum"]), _pct_txt(f["HumRef"]), _pct_txt(f["Pen"]),
                     _pct_txt(f["Finos"] if prod == "CHIP" else 0.0),
                     _kg2(f["KgChip"]) if prod == "CHIP" else "", _kg2(f["KgAse"]) if cs else "",
                     _kg2(f["KgCorto"]) if f["KgCorto"] else "", _kg2(f["KgLargo"]) if f["KgLargo"] else "",
                     _kg2(f["KgRaleo"]) if f["KgRaleo"] else ""])
        if cs:
            est.append(("BACKGROUND", (7, i), (7, i), naranja if f["Pen"] > 0 else verde))
            fv = f["Finos"] if prod == "CHIP" and pd.notna(f["Finos"]) else 0.0
            est.append(("BACKGROUND", (8, i), (8, i), naranja if fv >= muy else (amarillo if fv >= alto else verde)))
    fl = c["filas"]
    data.append(["", "", "", P("TOTAL", 7, True, 1), P(decimal(fl["Kilos"].sum() / 1000, 2), 7, True, 1), "", "", "", "",
                 P(decimal(fl["KgChip"].sum(), 2), 7, True, 1), P(decimal(fl["KgAse"].sum(), 2), 7, True, 1),
                 P(decimal(fl["KgCorto"].sum(), 2), 7, True, 1), P(decimal(fl["KgLargo"].sum(), 2), 7, True, 1),
                 P(decimal(fl["KgRaleo"].sum(), 2), 7, True, 1)])
    w = [15, 21, 24, 17, 20, 18, 19, 20, 20, 21, 21, 18, 18]
    w = [x * mm for x in w]
    w.append(ancho - sum(w))
    det = Table(data, colWidths=w, repeatRows=1)
    det.setStyle(TableStyle(est))
    el += [det, Spacer(1, 4 * mm)]

    pdata = [["", P("Toneladas", 8.5, True, 1), P("Precio x TN", 8.5, True, 1), P("TOTAL", 8.5, True, 1)]]
    for t in MYM_PRODUCTOS + [t for t in c["precios"] if t not in MYM_PRODUCTOS]:
        tn = c["tn_prod"].get(t, 0.0)
        p = c["precios"].get(t)
        pdata.append([P(etiqueta_mym(t), 8.5, True), decimal(tn, 2), pesos(p) if p else "",
                      pesos(tn * p) if p else pesos(0)])
    pdata.append([P("TN", 9, True), P(decimal(sum(c["tn_prod"].values()), 2), 9, True, 1), "",
                  P(pesos(c["importe"]), 9, True, 1)])
    pt = Table(pdata, colWidths=[24 * mm, 26 * mm, 30 * mm, 36 * mm])
    pt.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black), ("ALIGN", (1, 0), (-1, -1), "CENTER"),
                            ("FONTSIZE", (0, 1), (-1, -1), 8.5), ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    tot = Table([[P("Total a Facturar", 10, True, 1), P(pesos(c["importe"]), 12, True, 1)]],
                colWidths=[42 * mm, 52 * mm])
    tot.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 1.2, colors.black), ("BACKGROUND", (0, 0), (-1, -1), amarillo),
                             ("GRID", (0, 0), (-1, -1), 0.5, colors.black)]))
    firmas = Table([[[P("Aceptación Proveedor", 8.5, True)], "", [P("Firma Empresa", 8.5, True)]]],
                   colWidths=[62 * mm, 10 * mm, 62 * mm], rowHeights=[16 * mm])
    firmas.setStyle(TableStyle([("BOX", (0, 0), (0, 0), 1, colors.black), ("BOX", (2, 0), (2, 0), 1, colors.black),
                                ("VALIGN", (0, 0), (-1, -1), "TOP")]))
    der = [tot, Spacer(1, 4 * mm), firmas, Spacer(1, 3 * mm)] + [P(t, 6.5, color=rojo) for t in PIE]
    abajo = Table([[pt, "", der]], colWidths=[118 * mm, 10 * mm, ancho - 128 * mm])
    abajo.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    el.append(KeepTogether([abajo]))

    def pie(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(gris)
        canvas.drawString(10 * mm, 7 * mm, f"Certificación Nº {c['numero']} · {c['razon']} · "
                                           f"{texto_periodo(c['desde'], c['hasta'])}")
        canvas.drawRightString(pag[0] - 10 * mm, 7 * mm, f"Hoja {doc.page}")
        canvas.restoreState()

    doc = SimpleDocTemplate(str(ruta), pagesize=pag, leftMargin=10 * mm, rightMargin=10 * mm,
                            topMargin=8 * mm, bottomMargin=12 * mm,
                            title=f"Certificación {c['numero']} - {c['razon']}", author=EMPRESA_LINEAS[0])
    doc.build(el, onFirstPage=pie, onLaterPages=pie)


def _hoja_mym(ws, c, logo=None):
    """Formato de la planilla de MyM, con las mismas fórmulas: columnas A a N."""
    from openpyxl.formatting.rule import CellIsRule
    for col, a in zip("ABCDEFGHIJKLMN", [11, 15, 17, 17, 12, 11, 12, 13, 13, 13, 13, 12, 12, 12]):
        ws.column_dimensions[col].width = a
    if logo:
        img = XLImage(io.BytesIO(logo))
        img.width, img.height = 110, 58
        ws.add_image(img, "A1")
    for i, l in enumerate(EMPRESA_LINEAS, start=1):
        ws.cell(i, 3, l).font = Font(size=9)
    ws.merge_cells("E1:H3")
    ws["E1"] = "CERTIFICACIÓN ENTREGA BIOMASA"
    ws["E1"].font, ws["E1"].alignment = Font(bold=True, size=14, color="2E75B6"), CENTRO
    if c["estado"] != "CONFIRMADA":
        ws["E4"] = {"MODIFICADA": "MODIFICADA — volver a confirmar", "BLOQUEADA": "BORRADOR — tiene errores bloqueantes",
                    "REVISAR": "BORRADOR — observaciones sin revisar", "LISTA": "BORRADOR — lista para confirmar"}[c["estado"]]
        ws["E4"].font = Font(bold=True, color="C00000")
    for r, (k, v) in enumerate([("Fecha:", c["fecha"]), ("N° :", c["numero"] or "(al confirmar)")], start=1):
        ws.cell(r, 9, k).font = Font(bold=True)
        ws.merge_cells(start_row=r, start_column=10, end_row=r, end_column=11)
        x = ws.cell(r, 10, v)
        x.alignment, x.border = CENTRO, B_FINO
        if r == 1:
            x.number_format = "dd/mm/yyyy"
    ws["A5"], ws["B5"] = "Período:", texto_periodo(c["desde"], c["hasta"]).replace("Período ", "")
    ws["A6"], ws["B6"] = "Proveedor:", c["razon"]
    ws["A7"], ws["B7"] = "CUIT:", c["cuit"] or ""
    for r in (5, 6, 7):
        ws.cell(r, 1).font = Font(bold=True)
    g_txt, h_txt = criterio_mym(c)
    ws.merge_cells("E5:N5")
    ws["E5"], ws["E5"].alignment = "Criterio de evaluación", CENTRO
    ws["E6"], ws["F6"], ws["G6"] = "Granulometría:", "Alto", "Muy Alto"
    ws["F7"], ws["G7"] = c["tol_finos"] / 100, c["muy_alto"] / 100
    for k in ("F7", "G7"):
        ws[k].number_format, ws[k].fill, ws[k].alignment = "0%", AMARILLO, CENTRO
    ws.merge_cells("H6:N7")
    ws["H6"], ws["H6"].alignment = g_txt, Alignment(wrap_text=True, vertical="top")
    ws["H6"].font = Font(size=8)
    ws["E8"], ws["F8"] = "Humedad:", "Referencias extraídas del chipeado propio de raleo fresco"
    ws.merge_cells("F8:G8")
    ws.merge_cells("H8:N8")
    ws["H8"], ws["H8"].alignment = h_txt, Alignment(wrap_text=True, vertical="top")
    ws["H8"].font = ws["F8"].font = Font(size=8)
    ws.row_dimensions[6].height = ws.row_dimensions[7].height = 22
    ws.row_dimensions[8].height = 24
    cols = ["Fecha", "N° Rto Planta ", "Rto.Proveedor", "Producto", "Kilos de carga", "Humedad medida",
            "Humedad de referencia", "Porcentaje de penalizacion", "Porcentaje medido de aserrin", "Kilos de chip",
            "kilos de aserrin", "Kilos de C Corto", "Kilos C Largo", "Kilos Rollo"]
    for j, t in enumerate(cols, 1):
        x = ws.cell(10, j, t)
        x.font, x.alignment, x.border = Font(bold=True), Alignment(horizontal="center", vertical="center", wrap_text=True), B_FINO
        x.fill = PatternFill("solid", fgColor="DDEBF7")
    ws.row_dimensions[10].height = 42
    n = len(c["filas"])
    ini, ult = 11, 10 + n
    tot = ult + 1
    pr = tot + 2                       # cuadro de precios
    fila_p = {t: pr + 1 + i for i, t in enumerate(MYM_PRODUCTOS)}
    lab = {t: f"$A${fila_p[t]}" for t in MYM_PRODUCTOS}
    for k, (_, f) in enumerate(c["filas"].iterrows()):
        r = ini + k
        rp = f["Remito Proveedor"]
        vals = [f["Fecha"], f["Rto. Planta"], rp, etiqueta_mym(f["Producto"]), float(f["Kilos"]),
                None if pd.isna(f["Hum"]) else float(f["Hum"]), None if pd.isna(f["HumRef"]) else float(f["HumRef"])]
        for j, v in enumerate(vals, 1):
            ws.cell(r, j, v)
        if f.get("PenManual") or c.get("sin_hum"):
            ws.cell(r, 8, float(f["Pen"] or 0.0))           # descuento de humedad puesto / quitado a mano
        else:
            ws.cell(r, 8, f'=IF(F{r}="",0,IF(OR(D{r}={lab["CHIP"]},D{r}={lab["ASERRIN"]}),IF(F{r}>G{r},F{r}-G{r},0),0))')
        fin = f["Finos"] if f["Producto"] == "CHIP" and pd.notna(f["Finos"]) else 0
        ws.cell(r, 9, float(fin))
        ws.cell(r, 10, f'=IF(D{r}={lab["CHIP"]},(E{r}-K{r})*(1-H{r}),"")')
        if f["Producto"] == "CHIP" and (f.get("DescManual") or c.get("sin_gran")):
            # % de aserrín puesto / quitado a mano: reemplaza al criterio Alto / Muy Alto en este camión
            ws.cell(r, 11, f'={float(f["DescPct"] or 0) / 100:.15g}*E{r}*(1-H{r})')
        else:
            ws.cell(r, 11, f'=IF(D{r}={lab["CHIP"]},IF(I{r}<$F$7,0,IF(I{r}<$G$7,(I{r}-$F$7)*E{r},I{r}*E{r}))*(1-H{r}),'
                           f'IF(D{r}={lab["ASERRIN"]},E{r}*(1-H{r}),""))')
        for j, t in ((12, "C CORTO"), (13, "C LARGO"), (14, "Fino/Raleo")):
            ws.cell(r, j, f'=IF(D{r}={lab[t]},E{r},0)')
        ws.cell(r, 1).number_format = "dd-mm-yy"
        for j in (6, 7, 8, 9):
            ws.cell(r, j).number_format = "0.00%"
        for j in (5, 10, 11, 12, 13, 14):
            ws.cell(r, j).number_format = "#,##0.00"
        for j in range(1, 15):
            ws.cell(r, j).border, ws.cell(r, j).alignment = B_FINO, CENTRO
        if f["Producto"] in ("CHIP", "ASERRIN"):
            ws.cell(r, 8).fill = ws.cell(r, 9).fill = VERDE
        for j in (6, 9):                                   # humedad medida y % de aserrín: editables
            ws.cell(r, j).fill = PatternFill("solid", fgColor="DDEBF7")
    ws.conditional_formatting.add(f"H{ini}:H{ult}", CellIsRule(operator="greaterThan", formula=["0"],
                                                               fill=PatternFill("solid", bgColor="FFC000", fgColor="FFC000")))
    ws.conditional_formatting.add(f"I{ini}:I{ult}", CellIsRule(operator="greaterThanOrEqual", formula=["$G$7"],
                                                               fill=PatternFill("solid", bgColor="FFC000", fgColor="FFC000")))
    ws.conditional_formatting.add(f"I{ini}:I{ult}", CellIsRule(operator="between", formula=["$F$7", "$G$7"],
                                                               fill=PatternFill("solid", bgColor="FFFF00", fgColor="FFFF00")))
    ws.cell(tot, 4, "TOTAL").font = Font(bold=True)
    ws.cell(tot, 5, f"=SUM(E{ini}:E{ult})/1000")
    for j in range(10, 15):
        L = ws.cell(ini, j).column_letter
        ws.cell(tot, j, f"=SUM({L}{ini}:{L}{ult})")
    for j in (5, 10, 11, 12, 13, 14):
        x = ws.cell(tot, j)
        x.number_format, x.font, x.border, x.alignment = "#,##0.00", Font(bold=True), B_MEDIO, CENTRO
    for j, t in enumerate(["Toneladas", "Precio x Tn", "TOTAL"], 2):
        x = ws.cell(pr, j, t)
        x.font, x.alignment, x.border = Font(bold=True), CENTRO, B_FINO
    col_kg = {"CHIP": "J", "ASERRIN": "K", "C CORTO": "L", "C LARGO": "M", "Fino/Raleo": "N"}
    for t in MYM_PRODUCTOS:
        r = fila_p[t]
        ws.cell(r, 1, etiqueta_mym(t)).font = Font(bold=True)
        ws.cell(r, 2, f"={col_kg[t]}{tot}/1000").number_format = "0.00"
        p = c["precios"].get(t)
        ws.cell(r, 3).fill = (ROJO_CLARO if (p is None and t in c["tn_prod"])
                              else PatternFill("solid", fgColor="DDEBF7"))     # precio editable
        ws.cell(r, 3, p)
        ws.cell(r, 4, f"=B{r}*C{r}")
        ws.cell(r, 3).number_format = ws.cell(r, 4).number_format = MONEDA
        for j in range(1, 5):
            ws.cell(r, j).border, ws.cell(r, j).alignment = B_FINO, CENTRO
    rt = pr + 1 + len(MYM_PRODUCTOS)
    ws.cell(rt, 1, "TN").font = Font(bold=True, size=12)
    ws.cell(rt, 2, f"=SUM(B{pr + 1}:B{rt - 1})").number_format = "0.00"
    ws.cell(rt, 4, f"=SUM(D{pr + 1}:D{rt - 1})").number_format = MONEDA
    for j in (1, 2, 4):
        ws.cell(rt, j).font, ws.cell(rt, j).border, ws.cell(rt, j).alignment = Font(bold=True), B_MEDIO, CENTRO
    ws.merge_cells(start_row=pr, start_column=7, end_row=pr, end_column=8)
    ws.cell(pr, 7, "Total a Facturar").font = Font(bold=True)
    ws.merge_cells(start_row=pr, start_column=9, end_row=pr, end_column=10)
    x = ws.cell(pr, 9, f"=D{rt}")
    x.number_format, x.font, x.fill, x.border, x.alignment = MONEDA, Font(bold=True, size=13), AMARILLO, B_MEDIO, CENTRO
    ws.cell(pr, 7).fill = AMARILLO
    r = pr + 2
    for col_ini, col_fin, texto in ((7, 9, "Aceptación Proveedor"), (11, 13, "Firma Empresa")):
        ws.merge_cells(start_row=r, start_column=col_ini, end_row=r + 3, end_column=col_fin)
        x = ws.cell(r, col_ini, texto)
        x.font, x.alignment = Font(bold=True), Alignment(vertical="top")
        for rr in range(r, r + 4):
            for cc in range(col_ini, col_fin + 1):
                ws.cell(rr, cc).border = Border(left=MEDIO if cc == col_ini else None, right=MEDIO if cc == col_fin else None,
                                                top=MEDIO if rr == r else None, bottom=MEDIO if rr == r + 3 else None)
    r += 5
    for t in PIE:
        ws.merge_cells(start_row=r, start_column=7, end_row=r, end_column=14)
        x = ws.cell(r, 7, t)
        x.font, x.alignment = Font(size=8, color="C00000"), Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 22
        r += 1
    ws.freeze_panes = "A11"
    ws.print_title_rows = "10:10"
    ws.print_area = f"A1:N{max(r, rt)}"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.4, bottom=0.4)
    ws.sheet_properties.tabColor = {"CONFIRMADA": "3A7D44", "LISTA": "92D050"}.get(c["estado"], "C00000")


def excel_certificacion(ruta, c, logo=None):
    """Excel individual de una certificación, con el mismo formato que el PDF."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Certificación"
    _hoja_certificacion(ws, c, logo)
    wb.properties.title = f"Certificación {c['numero']} - {c['proveedor']}"
    wb.save(ruta)


def escribir_libro(ruta, certs, logo=None):
    """Un solo Excel con todas las certificaciones del período: índice + una hoja por proveedor."""
    wb = Workbook()
    ix = wb.active
    ix.title = "Índice"
    enc = ["Estado", "Proveedor", "N° Certificación", "Remitos", "TN", "Importe", "Hoja"]
    ix.append(enc)
    usados = {"índice"}
    for c in certs:
        hoja = _nombre_hoja(c["proveedor"], usados)
        _hoja_certificacion(wb.create_sheet(hoja), c, logo)
        ix.append([f"{ICONOS[c['estado']]} {c['estado']}", c["proveedor"], c["numero"], len(c["filas"]),
                   c["tn"], c["importe"], hoja])
        x = ix.cell(ix.max_row, 7)
        x.hyperlink, x.style = f"#'{hoja}'!A1", "Hyperlink"
        ix.cell(ix.max_row, 6).number_format = MONEDA
        ix.cell(ix.max_row, 1).fill = VERDE_CLARO if c["estado"] == "CONFIRMADA" else ROJO_CLARO
    ix.append(["", "TOTAL", "", sum(len(c["filas"]) for c in certs), round(sum(c["tn"] for c in certs), 2),
               sum(c["importe"] or 0 for c in certs), ""])
    for x in ix[ix.max_row]:
        x.font = Font(bold=True)
    ix.cell(ix.max_row, 6).number_format = MONEDA
    for c in ix[1]:
        c.font, c.fill = Font(bold=True, color="FFFFFF"), ENCAB
    for col, a in zip("ABCDEFG", [16, 40, 18, 10, 12, 18, 34]):
        ix.column_dimensions[col].width = a
    ix.freeze_panes = "A2"
    ix.page_setup.orientation = "landscape"
    ix.page_setup.paperSize = ix.PAPERSIZE_A4
    ix.page_setup.fitToWidth, ix.page_setup.fitToHeight = 1, 0
    ix.sheet_properties.pageSetUpPr.fitToPage = True
    wb.save(ruta)


def escribir_resumen(ruta, certs, globales, obs_cfg, cuad, log, mapa_humedad=None):
    filas, ctrl, hum, gran, vistos = [], [], [], [], set()
    for c in certs.values():
        for t, h in c.get("humedad", {}).items():
            if not h["aplica_hum"]:
                continue
            hum.append({"Período": c["periodo"], "Certificación": c["proveedor"], "Producto": t,
                        "Muestras": h["muestras"], "Humedad prom.": h["H"], "Límite planilla": h["L_lab"],
                        "Límite fijo pactado": h["L_fijo"], "Tolerancia": h["tol"], "Límite aplicado": h["L"],
                        "Descuento %": h["desc_hum"], "TN": h["tn"], "TN descontadas": h["tn_desc"],
                        "TN a pagar": h["tn_pagar"], "Origen": h["origen"]})
    for c in sorted(certs.values(), key=lambda c: c["periodo"] != "Mensual"):
        for _, f in c["filas"].iterrows():
            if f["_id"] in vistos or f["Producto"] != "CHIP" or not c.get("desc_gran"):
                continue
            vistos.add(f["_id"])
            gran.append({"Certificación": c["proveedor"], "Proveedor (maestro)": f["Proveedor"], "Fecha": f["Fecha"],
                         "Rto. Planta": f["Rto. Planta"], "Rto. Proveedor": f["Remito Proveedor"],
                         "KG TOTAL": f["Kilos"], "Finos %": None if pd.isna(f["Finos"]) else round(f["Finos"] * 100, 2),
                         "Hora muestra": f["FinosHora"], "Cómo se asignó": f["FinosOrigen"] or "sin muestra",
                         "Desc. %": f["DescPct"], "Kg a aserrín": f["DescKg"], "Kg chip": f["KilosNeto"]})
    for c in certs.values():
        filas.append({"Estado": c["estado"], "Período": c["periodo"], "Proveedor": c["proveedor"],
                      "Grupo": c["grupo"], "N° Certificación": c["numero"], "Remitos": len(c["filas"]),
                      "TN entregadas": c["tn"],
                      "TN chip → aserrín (granulometría)": round(c["filas"]["DescKg"].sum() / 1000, 2),
                      "TN desc. humedad": round(sum(h["tn_desc_hum"] for h in c.get("humedad", {}).values()), 2),
                      "TN a pagar": round(sum(tn_pagar(c, t) for t in c["precios"]), 2), "Importe": c["importe"],
                      "Bloqueantes": sum(o["bloq"] for o in c["obs"]),
                      "Críticas": sum(o["sev"] == CRIT for o in c["obs"]),
                      "Advertencias": sum(o["sev"] == ADV for o in c["obs"])})
        for o in c["obs"]:
            ctrl.append({"Severidad": "BLOQUEANTE" if o["bloq"] else o["sev"], "Período": c["periodo"],
                         "Proveedor": c["proveedor"], "Rto. Planta": o["rem"], "Observación": o["msg"]})
    for g in globales:
        ctrl.append({"Severidad": g["sev"], "Período": "Mes", "Proveedor": g["prov"], "Rto. Planta": g["rem"],
                     "Observación": g["msg"]})
    for sev, g, msg in obs_cfg:
        ctrl.append({"Severidad": sev, "Período": "Config", "Proveedor": g, "Rto. Planta": "", "Observación": msg})
    ctrl = pd.DataFrame(ctrl, columns=["Severidad", "Período", "Proveedor", "Rto. Planta", "Observación"])
    orden = {"BLOQUEANTE": 0, CRIT: 1, ADV: 2}
    ctrl = (ctrl.drop_duplicates(["Severidad", "Proveedor", "Rto. Planta", "Observación"])
                .sort_values(["Severidad", "Proveedor"], key=lambda s: s.map(orden) if s.name == "Severidad" else s))
    with pd.ExcelWriter(ruta, engine="openpyxl") as w:
        cuad.to_excel(w, sheet_name="Cuadre", index=False)
        pd.DataFrame(filas).to_excel(w, sheet_name="Certificaciones", index=False)
        if gran:
            pd.DataFrame(gran).sort_values(["Certificación", "Fecha", "Rto. Planta"]).to_excel(
                w, sheet_name="Granulometría por camión", index=False)
        if hum:
            pd.DataFrame(hum).to_excel(w, sheet_name="Humedad", index=False)
        if mapa_humedad is not None:
            mapa_humedad.to_excel(w, sheet_name="Vínculo humedad", index=False)
        ctrl.to_excel(w, sheet_name="Control", index=False)
        pd.DataFrame(log, columns=["Fecha y hora", "Responsable", "Acción", "Certificación", "Rto. Planta",
                                   "Campo", "Antes", "Después", "Motivo"]
                     ).to_excel(w, sheet_name="Registro de cambios", index=False)
    wb = load_workbook(ruta)
    for ws in wb:
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = max(
                10, min(80, max(len(str(c.value or "")) for c in col) + 2))
        for c in ws[1]:
            c.font, c.fill = Font(bold=True, color="FFFFFF"), ENCAB
        ws.freeze_panes = "A2"
    if "Humedad" in wb.sheetnames:
        for row in wb["Humedad"].iter_rows(min_row=2):
            for i in (4, 5, 6, 7, 8, 9):
                row[i].number_format = "0.00%"
            if (row[9].value or 0) > 0:
                for x in row:
                    x.fill = AMARILLO_CLARO
    if "Granulometría por camión" in wb.sheetnames:
        for row in wb["Granulometría por camión"].iter_rows(min_row=2):
            row[2].number_format = "dd/mm/yyyy"
            if (row[10].value or 0) > 0:
                for x in row:
                    x.fill = AMARILLO_CLARO
    for row in wb["Certificaciones"].iter_rows(min_row=2):
        row[11].number_format = MONEDA
        row[0].fill = VERDE_CLARO if row[0].value == "CONFIRMADA" else ROJO_CLARO
    for row in wb["Cuadre"].iter_rows(min_row=2):
        row[-1].fill = VERDE_CLARO if "✔" in str(row[-1].value) else ROJO_CLARO
    wb.save(ruta)


# Logo genérico del encabezado de las certificaciones (se reemplaza desde "Imágenes")
LOGO_B64 = (
    "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAMCAgICAgMCAgIDAwMDBAYEBAQEBAgGBgUGCQgKCgkICQkKDA8MCgsOCwkJDRENDg8Q"
    "EBEQCgwSExIQEw8QEBD/2wBDAQMDAwQDBAgEBAgQCwkLEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQ"
    "EBAQEBAQEBD/wAARCAEYAjADASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUF"
    "BAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVW"
    "V1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi"
    "4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAEC"
    "AxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVm"
    "Z2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq"
    "8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD9U6KKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKK"
    "KACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACi"
    "iigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigA"
    "ooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAoooo"
    "AKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKK"
    "KACiiigAoorL8Q+KfDHhGyTUvFniPS9FtJJRAlxqN5HbRtIQSEDSEAsQrHHXCn0o2FKSirt6GpRXB33x6+CWn2Vxfz/Fvwi8VtE8"
    "zrBrFvNIVUEkJHGxd2wOFUFieACTiuR/4bM/Zs/6KR/5R7//AOMVDqQW7OeeNw1PSdSK+aPaqK8V/wCGzP2bP+ikf+Ue/wD/AIxX"
    "Q+Hv2kfgP4nsnv8ATfir4ehijlMJXUbsWEhYAHIjudjlcMPmA2k5GcggCqQezQo47Czdo1It+q/zPSaKo6Nrmi+I9Nh1nw9q9lqm"
    "n3G7ybuyuEnhk2sVba6EqcMpBweCCO1Xqs6U01dBRRRQMKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACi"
    "iigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKK8v8Aiv8A"
    "tI/Cb4O+ZaeJ/EH2rV0x/wASfTQtxeDOw/OuQsXySK481k3KDt3EYpSkoq7Mq1anQhz1ZJLzPUK57xl8Q/A3w9shf+N/Fml6LE8U"
    "s0S3dyqSTrEAXEUed8rAMvyoGbLKAMkA/C/xe/bt8f8AjD7Tovw3tf8AhE9Ik3xfa9wk1GeM+Yud/wB2DcjIcRgujplZSK+bda1z"
    "WvEepTaz4h1e91TULjb513e3DzzSbVCrudyWOFUAZPAAHauWeLitI6nzmM4mpU7xw0eZ93ov83+B96+J/wDgoR8LNNivovC3hbxD"
    "rV3bymO2aZIrS1uVD4LiQs8iKVyy5i3HgELkkeA+L/25vjtr2tS33hvVrLwzp/zJDY2tjBcYTexVpJJ0dmk2lVJXYp2ghFyc/PdF"
    "cssRUl1PncRnmOxGjnyry0/Hf8TrvE/xe+KfjSK+tfFPxD8Q6jaajKZrmym1GX7K7b94AgBEaqGAKqFCrgYAwK+pv23v+KD+C/wx"
    "+EP/AB/fZtn/ABMf9Vu/s+0S3/1XzY8z7Tu+/wDLsx82cj4/8LeHr3xd4n0jwnpssEd3rV/b6dbvOxWNZJpFjUuQCQoLDJAJx2Nf"
    "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"
    "n4ZXjZ7Upfc1+Z8z0V9c+G/+CdXjS6+0f8Jf8RtE0zbs+z/2bay33mZzu3+Z5GzHy4xuzk9Mc+k+GP8Agnv8LNNisZfFPinxDrV3"
    "byiS5WF4rS1uVD5CGMK8iKVwrYl3HkgrkAWsNUfQ6qWQY+r9iy82v+H/AAPz9rrvDHwh+KfjSKxuvC3w88Q6jaajKIba9h06X7K7"
    "b9hJnIEaqGBDMWCrg5Iwa/Srwh+zJ8CfAutReIfDfw7sotQg2mGa6nnvPJdXV1kjWd3VJAyKQ6gMOcEZOfUK2jg39pnq4fhWT1xF"
    "T7v83/kfC/w0/wCCe+v30sV/8WPFMGm2jRJIdP0dxLdbmRtyPM6+XGyNsztEqt84BHDH6s+F/wAD/hn8H7IQeCfDcEF20XlXGpT/"
    "AL29uBhA2+VuQrGNWMabY93IUV3lFdMKMKeyPocJlWFwWtKOvd6v+vQKKKK1PRCiiigAooooAKKKKACiiigAooooAKKKKACiiigA"
    "ooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAoooo"
    "AKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKK"
    "KACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACi"
    "iigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigA"
    "ooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAoooo"
    "AKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKK"
    "KACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACivzp/aL/4KL/G34RfGzxZ8N/DfhbwPc6ZoV4tvbS31jdvO6mJ"
    "Gy7JcqpOWPRRXnH/AA9e/aJ/6Ez4c/8Aguvv/kyp50cMswoxbi76H6u0V+UsX/BV/wDaEDgzeCfh4ydwthfKfzN2f5V6Z8NP+Cst"
    "ldX8Fh8W/hl9htpGCyajodyZRFk4ybeXkqOpxITgcA0c6HHMKEna9j9DqKxfBvjPwv8AEHwzYeMfBmt22raNqcQmtbu3bKuvQgg4"
    "KsCCCpAKkEEAgivBv24v2lPHX7NHgnw74k8CaVoV/c6vqjWM6avBNLGsYhZ8qIpYyDkDqSMdqd7K50zqxpw9o9j6Tor8sdA/4Kn/"
    "ALQeq67pul3Hg74eLFeXcNu5TT74MFdwpIzdkZwfSv1OoTT2IoYmGIvydAorjfit8XPAPwV8I3Hjb4ia9FpunQfJGv3prmUjKxQx"
    "jl3OOg6AEkgAkfnn8VP+CrXxA1S9uLL4QeCtM0PTslYr3V1N1eOOzbFYRRn/AGT5n1ockhVsVSofG9T9PqK/Glf+CjX7W63RuD8R"
    "LNo/+eB0Kx2fmIt3/j1ezfCT/gq54qs7630/41eB7HUdPOEk1HQlMF1H/ttDI5jk+itH/QypowjmVCTs7o/TCiub+HnxF8F/FXwp"
    "ZeNvAOvW+r6RfLmOaE8ow+9G6n5kdc4KsAR6Vp+ItQm0jw/qeq2yo01lZzXEauCVLIhYA4IOMj1qzuTTV0aNFflF/wAPXv2if+hM"
    "+HP/AILr7/5Mo/4evftE/wDQmfDn/wAF19/8mVHOjh/tKh5/cfq7RXnH7O3xZX44fBjwv8TGit4bvVrPF/DbgiOK8jYxzqoYlgvm"
    "IxUEk7SvJ6n0erO2MlNKS2YUV8M/tlft8ePPgH8W0+Gvw20PwvqK2WmwXGpyatb3Erx3Mu51jXyp4wAIjE3IJ+f2rwj/AIevftE/"
    "9CZ8Of8AwXX3/wAmVLmkcs8fRpycW9Ufq7RXxh+xB+2h8Uf2lviBr3hTx1oPhawtNL0Y6hC+k2txFI0nnxx4YyzyArhz0AOcc19n"
    "007m9KrGtHnhsFFFfmr8Zv8Agp78VvCPxV8U+FPh94Z8E3nh/RtTm0+zub+zu5Z5hCdjuWjuUUhnVyuFHylevUjaW4q+Ihh0nPqf"
    "pVRX5Rf8PXv2if8AoTPhz/4Lr7/5Mr9Fv2dvi1H8cfgz4Y+Jhit4bvVrTF/Bb5EcN5GxjnVQxLBfMRioJJ2leT1KUk9iKOLp4iXL"
    "Dc9HoopGZUUu7BVUZJJwAKo6RaK+Gf2hv+CnfhTwJqt34S+C2g2vizUbRmhm1i7lZdNjkBwRGqYe4AOfmDIvdSwOa+VtY/4KRftZ"
    "anctPZeNtM0lGbIhs9EtGRR6AzJI2PqTUuaRw1Mwo03a9/Q/Y+ivyQ8Ff8FQv2jfD15G3iuLw94qs9w82O5sBazFe4SSDaqn3KMP"
    "av0B/Zp/a0+G37TGkTHw4ZdK8Q6fGsmoaHduDNCpOPMjcYE0WeNwAIJG5VyMikmaUcZSrvli9fM9uoooqjqCivzV+M3/AAU9+K3h"
    "H4q+KfCnw+8M+Cbzw/o2pzafZ3N/Z3cs8whOx3LR3KKQzq5XCj5SvXqeM/4evftE/wDQmfDn/wAF19/8mVPOjhlmNCLtc/V2ivOP"
    "2dvi1H8cfgz4Y+Jhit4bvVrTF/Bb5EcN5GxjnVQxLBfMRioJJ2leT1Po9UdkZKaUlswooooKCiiigAooooAKKKKACiiigAooooAK"
    "KKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAPxD/AG3v+TrPiN/2FE/9ERVY/ZD/AGY7T9qHxfrnhe78Yy+H"
    "V0fTRfiaOxFyZCZVTbtLpj72c5PSq/7b3/J1nxG/7Cif+iIq73/gnj8cPhd8DPiH4o134p+J/wCxLHUdFW0tpfsVxc+ZKJ0bbtgj"
    "cj5QTkgCsdL6nzUVCWJaqbXZ7B4l/wCCR2oQ6XNP4Q+NcF3qKITDbajopghlbsDKkzlB77Gr4D8TeHNZ8H+ItT8KeIrJrPVNHu5b"
    "G8t2IJimjco65HBwQeRwe1frx4o/4KS/sqaHpU99ovjDUvEd0iExWVjo11FJI3YbriONB9S3T16V+UPxB8W658Yfifr3jNtKc6l4"
    "p1We9SxtVaUq0rkrEgA3NtBCjjJxTkl0NcbTw8Lexevk7n3J/wAElvH2sPfeOPhlc3EkumJbwa3aRsSVgm3+VNt9N4MWf+uY967L"
    "/grT/wAkp8E/9jDJ/wCkz11X/BO79l/xP8DvCWseNviHYGw8SeK1hSPT3wZLGyjyyiTH3ZHZssn8IRAcNuA5X/grT/ySnwT/ANjD"
    "J/6TPVfZOvllDAtT3/4J+ang3/kb9D/7CVt/6NWv6Ebu6trC1mvr24jgt7eNpZZZGCrGijLMSeAAASTX893g3/kb9D/7CVt/6NWv"
    "2i/bg8X3Pgn9lf4garZMRPd6emlLg4O27mjt3x9Elc/hShomZZbP2dOpJ9P+Cflb+1h+0Prf7RfxWv8AxJNcyL4e02SSy8P2fIWG"
    "0DcSFf8AnpJgO569F6KuM/8AZ5/Zp+JH7SXiabQ/BFrDb2VgFfUtWvCy2tmjH5QSASztg7UUZOCeACR5PX7W/sIfDiw+HX7Mng8W"
    "9qqXviS1HiC+lCgNM9yA8ZP0h8lR7L71MVzPU5sNSeMrNzfmz5rm/wCCRUY0o/Z/jqx1IJkb/D+IC2Pu8XG4DP8AFz9O1fF3x2+A"
    "PxD/AGePGH/CIeP7CNTOhmsL+2Yva30IOC8TkA5BwGUgMuRkYIJ/eavln/gpD8ONP8bfs0ar4ia2VtT8HXNvqtnKFG8I0iwzpnqF"
    "MchYjuY19BVygraHfisBSVNypqzR8BfsQ/tI6l8Avi1ZWmo6ht8H+J54rHW4ZG/dw7jtjux/daMtknuhcdcEfsR42IPgvXiDx/Zd"
    "1/6Kav57K/cf4OeL7nx5+yX4e8VXzF7q88G7blyc75o7do5G/FkY/jSg+hnllVuMqb6an4cVbvNMvLC2sbu5iKxalbtc25/vxiWS"
    "In/vuJx+FVK908c+Bzc/sj/DP4m20P8Ax5a/rPh68kx2dxcW49sYuP8Avqszyox5k32R9a/8Em/ib9r0Dxj8IL24zJp88evaehOS"
    "YpQIrgD0VWSE/WU19/3d3bWFpNfXs6Q29tG00sjnCoijLMT2AAJr8TP2I/iU/wAL/wBpfwbqss5jsdXu/wCwr4ZwpiusRqWPospi"
    "c/7lfpr+3p8SX+Gv7Mfiq5tZzFfeIUTw9aEHBJucrLg+ogWcj3ArWL0PbwVdLDNv7J+RHxl+INz8Vvit4q+Il0X/AOJ9qk93Cr9Y"
    "4C2IY/8AgMYRf+A1zes6NeaFdx2V+m2WW0trwD/pnPCk0Z/FJFP40mhaNf8AiPW9P8PaVF5t7ql1FZW0f9+WRwiD8Swr179s3QbD"
    "wp+0d4o8KaWMWehW2kaXb8Y/d2+l2sS/ogrPzPFacouo+/53Pd/+CTX/ACWfxh/2LB/9K4K/Uuvy0/4JNf8AJZ/GH/YsH/0rgr9S"
    "61hse9l38BfM85/aJ+JafB/4JeMPiGJVjudL0yT7ET3vJMRW4/7+umfbNfgy7vI7SSOzu5LMzHJJPUk1+mH/AAVh+J/9n+EvCPwi"
    "sbjEur3T61fop5EEIMcKt/stI8h+sIr8+fhN4Lm+I3xP8KeA4UZv7e1i0sH2/wAMckqh2+ioWJ9hUTd3Y8/MZupWVNdPzZW8f+BN"
    "c+G/iM+FvEUXl3y2NjfOuMbVubWK4VfqFlCn3Br7/wD+CTfxR8/TfGPwbvrnL2sieINOQnJ2PthuQPQBhbnHq7GuC/4KtfD1NC+K"
    "HhLx9ZWqxWmvaM2nSbFwvn2j9T6fu54gB6J7V8/fsg/FH/hUP7Q/g7xXcXPk6dNejTNSJOF+y3P7p2b2Qssn1jFJe7Izh/smKt0v"
    "+DP3Kr5U/wCCkXxX1b4a/s9SaR4fupLa/wDGV+uivNGcNHaGN3nwf9pUEZ9pDX1XXxx/wVF+H+q+K/gJp/irSbZ5/wDhEtYS7vFQ"
    "ZKWkqNE0nr8rmLPoCSeBWstj2sW5KjLl3sfkyASQAMk1+jnwx/4JQaRe+F7LUvit8RdVtdZu4FmmsNIgiVLNmUHy2kkD+YynqQqj"
    "OQM9T+cVfo98Dv8AgqhotpoNh4e+OXhPU2v7SJLdtb0gJKtztGPMlgYqUY4yxQsCSSFXpWUbdTw8H7Dmft/keF/te/sMat+zZpFp"
    "448P+J38ReFru6WyleeAQ3NlMykoHCkq6NtbDgDBwCOQT4V8GPifrXwb+J/h74j6HPIk2jXqSzRof+Pi2J2zQn1Dxll/HPUA1+xP"
    "hX9p39lT472ieHbfx34Y1X7W6Y0jXYBA8sgOVUQ3SqJGB6bd3Tiu9/4Ur8G/+iS+DP8AwQ2v/wAbquW+qO14CM5+0oSSX3nYQTRX"
    "MMdxC4eOVQ6MOhUjINeeftE/EtPg/wDBLxh8QxKsdzpemSfYie95JiK3H/f10z7Zr0SOOOGNIYY1SNFCqqjAUDoAOwr4F/4Kw/E/"
    "+z/CXhH4RWNxiXV7p9av0U8iCEGOFW/2WkeQ/WEVcnZHoYmp7KlKZ+Z7u8jtJI7O7kszMckk9STXQ+P/AAJrnw38Rnwt4ii8u+Wx"
    "sb51xjatzaxXCr9QsoU+4NWfhN4Lm+I3xP8ACngOFGb+3tYtLB9v8Mckqh2+ioWJ9hX1z/wVa+HqaF8UPCXj6ytVitNe0ZtOk2Lh"
    "fPtH6n0/dzxAD0T2rG2lz52NFypSqdrf1+R3v/BJv4o+fpvjH4N31zl7WRPEGnITk7H2w3IHoAwtzj1djX6GV+Gv7IPxR/4VD+0P"
    "4O8V3Fz5OnTXo0zUiThfstz+6dm9kLLJ9YxX7lVpB3R7OXVeejyvoFFFFWegFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRR"
    "RQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAfiH+29/ydZ8Rv+won/oiKsD9n/8AZy8e/tI+INT8N+ALzRre60qzF9O2p3Ek"
    "KGMuEwpSNyTlh2Fb/wC29/ydZ8Rv+won/oiKvd/+CS//ACVrxr/2Lqf+lMdYpXkfNRpqrinCWzbOQ1f/AIJe/tN6bZSXdm3hHVZE"
    "UkW1pqrrI+B0BmiRMn3YV4N4Z8b/ABl/Zv8AGuo23h7VNT8H+I7KX7NfwNCm8MhyEkSRSrDnPIIIIPINfvbX4+f8FLNf8Ja9+03d"
    "nwtc29xNp+j2ljq8kDBlN8jSFlJHBZYmhQ9cFSDyCBUopao6sZhIYaKqU20z7K/Yd/bWn/aGS68A/EG2tLPxrpdt9qjntl8uHVLd"
    "SFZ1T+CVSy7lHBB3KAAQvG/8Faf+SU+Cf+xhk/8ASZ6+NP2Fr3UbH9q/4eyaYzCSW/mgkCnrC9tKsgPtsLH8K+y/+CtP/JKfBP8A"
    "2MMn/pM9F7xLjWlWwUnPdafkfmp4N/5G/Q/+wlbf+jVr9eP+CkNrNcfsm+JpomIW2vtNlkA7qbuNMH/gTrX5D+Df+Rv0P/sJW3/o"
    "1a/dT9oP4dP8Wfgn4z+HsESyXWsaTMlkrkBTdoPMt8k9B5qR80o6pmeAi50qkV1X+Z+CdfvR+zfeW1/+z38M7q0ZTG3hHSFwvRSt"
    "pGrL+BBH4V+DdzbXFncS2l3A8M8DtHLHIpVkcHBUg8ggjGK/Uv8A4JpftG+H/E3w3t/gZ4h1WG28SeGml/suOeQA39izmQCMn7zx"
    "szKVHOwIRkBsEHqTllRQquL6n2/Xif7ad5bWH7LHxInuioRtHMI3dN8kiIn47mWvbK/Pv/gp7+0b4fHhaH9nzwrqsN5qd5dxXniL"
    "yHDLawxHfFbuRx5jSBHK9VEYyPmFaSdketiqip0ZN9j81a/Z79k+1ms/2HPDMNwxZm8OahKCf7ryXDqP++WFfjfoeiap4l1rT/Du"
    "h2cl3qOqXUVnaW8Yy0s0jBEQe5YgV+8OleDbb4d/AmDwHZ7TF4f8Lf2cGXo5itdjP9WIJPuazhueXlkXzSl5H4IV+gXwW+Hb/FH/"
    "AIJmeNNDtrfzb7StXvtYsQBlvMtRDMwUerRiRB/v1+ftfrh/wS/jjm/ZhmilRXR/Ed+rKwyGBjhyCO4pQ1ZjgIKpUcX1TPyUtbq4"
    "srmG8tJWingkWWJ1OCjqcgj3BFfe/wDwUs+L0Xjn4cfBqxsJl8nxHpp8WXEanhfMgiWD8vMuB+FfI/7Qvw0k+D/xq8YfDsxMlvpO"
    "pyCy3dTZyYktz9TE8effNYvjj4g6t460/wAJ2OqEn/hFNBj0G3O7IaFLieVDjtgTBf8AgNK9roxjN0oTpPr+jPZv+CfPw5b4hftP"
    "eGpZrfzLHwusviG6OPumAAQn6+e8J/A1i/tyuZP2r/iKx7ajCv5W0Q/pX2X/AMEpPhj/AGN8OfE/xWvrfbceJL9dNsmYc/ZbYEuy"
    "n0aWRlPvCK+Mv24v+TrviL/2Eo//AEniptWib1Kfs8JFvq7/AIHt/wDwSa/5LP4w/wCxYP8A6VwV+pdflp/wSa/5LP4w/wCxYP8A"
    "6VwV+hP7QPxKj+EHwX8X/EQyKk+kaZK1nu6NdviO3U+xleMfSrj8J6OAko4bmfS5+Rn7cXxP/wCFqftK+LdUtrjzdO0WcaDYEHKi"
    "K1yjlT3VpvOcez12X/BNvw5ouoftGQ+LPEOo2NlZeFNLub6OS7nSJGuZAII1G4jJ2yyMPTZn0r5XllkmkeaaRpJJGLOzHJYnkknu"
    "aZWd9bniqtar7WSvrc/Wf/gpDp/g/wCIf7Olxqej+JdGvNU8Kalb6pDHBfRPK8TEwSqoDZI2yhyP+mdfkxRRQ3d3KxNf6xPntY/d"
    "T9lL4o/8Lh+AHg7xrPc+dqD2C2WpMTljeW58qVm9CzJv+jivUtQ0+x1WwudL1Ozhu7O8ieC4t5kDxyxsCrIynhlIJBB6g1+d/wDw"
    "Sa+KOG8Y/Bq+uOuzxFpqE/7sNyB/5LEAf7Rr0z9vH9q34xfs2+J/Ctn4AtdDfTNesLiR31GxeZvPikAYKyyKMbZI+Md/etVLS57d"
    "LExWHVWfz/I85+PX/BLJr3UrrxH8AfEVraQzs0p8P6u7hIiTnbBcAMdvosg47ue3yD48/ZI/aQ+G/myeJ/hFr/2aHJe6sIBfQKo/"
    "iMluXVR/vEV9PfAj/gpf8TvEfxb8N+H/AIuyeFbDwpqV01rf3dvYyQvAXjZYnLtKwVRKY9xIwF3HjrX6YgggEEEHkEVPKpbHNHC4"
    "bGJzpXX9dj+dR0aNijqVZTggjBB9K+uf2Mv24fGHwj8T6X4A+I2u3GreAr6VLTdeSGSTRtx2rLE7HIhBI3xnIC5KgEEN9K/8FPvh"
    "t8Kv+FQL8R7zTdN0/wAaR6lbW1jdxIsdxqCscSQyYwZQsYLgnJXZwQGIP5X1L91nBUjPA1rRZ/RaCGAZSCDyCK/Ev9uL4n/8LU/a"
    "V8W6pbXHm6dos40GwIOVEVrlHKnurTec49nr9PtT+Kt38MP2L7D4p6vPjVbHwJp9wjSHO/UZbSJIgc9czyID9TX4lyyyTSPNNI0k"
    "kjFnZjksTyST3NVNnbmda8YwXXU+qP8Agm34c0XUP2jIfFniHUbGysvCml3N9HJdzpEjXMgEEajcRk7ZZGHpsz6V9ef8FIdP8H/E"
    "P9nS41PR/EujXmqeFNSt9UhjgvonleJiYJVUBskbZQ5H/TOvyYoqVKyscVPFKnRdHl3Cv3U/ZS+KP/C4fgB4O8az3Pnag9gtlqTE"
    "5Y3lufKlZvQsyb/o4r8K6/Rb/gk18UcN4x+DV9cddniLTUJ/3YbkD/yWIA/2jTg7M1y2ryVuV9T9GKKKK1PoAooooAKKKKACiiig"
    "AooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigD8Q/23v+TrPiN/2FE/8AREVeW+DviD47+Hl5"
    "PqPgLxlrfhy6uovJnm0u+ltXljyDtZoyCRkA4PcUUVg9z5Ss2qsmu7Oi1P8AaH+Pms2r2Oq/Gzx3d20gKvDL4hu2RweCCvmYI+te"
    "fck9yTRRSM3Jy3Z+j3/BN39kzxT4a1z/AIX98R9Hn0tltZIPDlhdRlJ281dsl26EZRTGWRAeWDs2ANpbpP8AgrT/AMkp8E/9jDJ/"
    "6TPRRWrVonu1KUaWCaj2PzU8G/8AI36H/wBhK2/9GrX9C1FFKmZZTtP5fqfmr+39+xLrkGv6l8dvhFokl9p+oF7zxFpVqpaa3uCc"
    "yXcSDl43+84GSrbm5UnZ8AW1zc2VxHd2k8kE8DiSOWNirowOQwI5BB7iiilNWZzZjSjSq3j1PRJv2lP2hLjSjok3xu8cPZFPLMR1"
    "255TGNpO/JGOMZxXnRM91OSS8s0r+7M7E/mSTRRUHC5OW7P0k/4J9/sU614W1S1+O3xe0d7K+ij3+HNIuBiWAuCDdTofuNtOEQ8j"
    "JYgELX3Z42/5EzX/APsF3X/opqKK3Ssj6ahSjRpKMT+euv1y/wCCXf8AybLJ/wBjJff+i4aKKzhuePln8f5HhH/BWH4Yf2d4v8Jf"
    "F2xt8Q6zaPot+yjgTwEvCzf7TRu6/SGvgi3t57ueO1toXlmmcRxxoMs7E4AA7kmiilPcjHxUcQ7H74fAz4cwfCT4QeEvhzCiK+ia"
    "XDBclOj3RG+dx/vStI341+P/AO3F/wAnXfEX/sJR/wDpPFRRVT2O/M4qNGMV0f6Ht/8AwSa/5LP4w/7Fg/8ApXBXqf8AwVg+J/8A"
    "Zngvwn8JLG4xNrd4+sX6qeRbwDZErD+60kjMPeGiij7BhGTjgHb+tT80LGyutSvbfTrGB57m6lSCGJBlndiAqj3JIFfv38JvAlp8"
    "MPhl4X+H1nsKaBpVtYu6jiSVEAkk+rPuY+7UUUUy8qivel6HWV+QH/BSv4af8IN+0ddeJLS38uw8aWMOrIVGFFwo8mdfqWjWQ/8A"
    "XWiinPY6cyinQu+jPKP2Vfij/wAKe+P3g7xvPceTYRX62epEnC/Y7gGKZj67VcuPdBX6pfttfs4XH7RnwlOm+HhEPFXh+Y6jovmO"
    "EWdiuJbcseFEi4wTgb0TJAyaKKUNU0cuXRVWnOnLY/GPXND1nwzq93oHiLSrvTdSsZTDc2l1C0UsLjqrKwBBr0vwV+1h+0Z8PNIi"
    "0Dwl8W9dtNOt08uC2ldLmOFP7qCZX2AdguAO1FFZ7Hl80qcnyuxx/wAQPij8Rfipqiaz8RfGeq+ILuFSkLX1wzrCpxlY0+7GDgZC"
    "gAnmu/8A2W/2avFn7R3xCs9FsLG5h8NWU6Sa7qoXEdtbggtGrHgzOOFUZPO4jaCQUU1qzbDQVesoz6n2T/wVU+Ilt4d8AeDfgtoh"
    "S3XU7g6lcwRcCO0tl8uCMj+6zuSPeCvzVsbK61K9t9OsYHnubqVIIYkGWd2ICqPckgUUU5bl45uWIaZ+/fwm8CWnww+GXhf4fWew"
    "poGlW1i7qOJJUQCST6s+5j7tXWUUVsfRpKKsj8gP+Clfw0/4Qb9o668SWlv5dh40sYdWQqMKLhR5M6/UtGsh/wCuteUfsq/FH/hT"
    "3x+8HeN57jybCK/Wz1Ik4X7HcAxTMfXarlx7oKKKxejPm8T+6xLcejufutRRRWx9KFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFF"
    "FAH/2Q=="
)


def logo_por_defecto():
    import base64
    return base64.b64decode(LOGO_B64)


def extraer_logo(ruta_modelo=None):
    """Logo de la certificación modelo subida; si no hay o no se puede leer, el logo incorporado."""
    if ruta_modelo:
        try:
            with zipfile.ZipFile(ruta_modelo) as z:
                # la primera imagen dibujada en la hoja de certificación
                for dib in sorted(n for n in z.namelist() if re.fullmatch(r"xl/drawings/drawing\d+\.xml", n)):
                    rid = re.search(r'r:embed="([^"]+)"', z.read(dib).decode("utf-8", "ignore"))
                    rels = dib.replace("drawings/", "drawings/_rels/") + ".rels"
                    if rid and rels in z.namelist():
                        m = re.search(r'Id="%s"[^>]*Target="\.\./media/([^"]+)"|Target="\.\./media/([^"]+)"[^>]*Id="%s"'
                                      % (rid.group(1), rid.group(1)), z.read(rels).decode("utf-8", "ignore"))
                        if m:
                            return z.read("xl/media/" + (m.group(1) or m.group(2)))
                imgs = sorted(n for n in z.namelist() if n.startswith("xl/media/"))
                if imgs:
                    return z.read(imgs[0])
        except Exception:
            pass
    return logo_por_defecto()
