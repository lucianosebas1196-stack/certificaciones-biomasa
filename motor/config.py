"""Config de proveedores: leer, guardar, validar; datos de cada proveedor; precios y tramos.

Parte del motor de certificaciones (antes en cert_biomasa.py). El código de cada función es el mismo de
la versión 2.1; solo se repartió en módulos."""
import re
import difflib

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

from . import base as _valores   # valores generales que cambian en marcha (corte, finos)
from .base import (
    CRIT, C_PROV, FRECUENCIAS, METODOS_HUMEDAD, SEP_VARIANTE, TIPOLOGIAS, fijar_corte, fijar_finos)
from .formato import _canon_tipo, dinero, moneda_de, nombre_producto, pesos


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
        wb = load_workbook(ruta, read_only=True)
    except Exception:
        return
    try:
        _porcentajes_hoja(wb["Proveedores"], conf)
    except Exception:
        pass
    finally:
        wb.close()            # read_only deja el archivo abierto (en Windows no se podría reemplazar ni borrar)


def _porcentajes_hoja(ws, conf):
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
                "SOLO PATENTES", "OTRAS PATENTES A", "SI TRANSPORTA", "FRECUENCIA", "MONEDA"):
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


def guardar_config(ruta, conf, precios, par):
    with pd.ExcelWriter(ruta, engine="openpyxl") as w:
        conf.reset_index().to_excel(w, sheet_name="Proveedores", index=False)
        precios.to_excel(w, sheet_name="Precios", index=False)
        pd.DataFrame({"PARÁMETRO": ["NUMERO_INICIAL", "PREFIJO", "METODO_HUMEDAD", "DIA_CORTE_QUINCENA",
                                    "FINOS_TOLERANCIA", "FINOS_TOTALIDAD"],
                      "VALOR": [int(par["NUMERO_INICIAL"]), str(par["PREFIJO"]).zfill(4),
                                str(par.get("METODO_HUMEDAD", "PUNTOS")).upper(),
                                int(par.get("DIA_CORTE_QUINCENA", _valores.DIA_CORTE)),
                                float(par.get("FINOS_TOLERANCIA", _valores.FINOS_TOLERANCIA)),
                                float(par.get("FINOS_TOTALIDAD", _valores.FINOS_TOTALIDAD))],
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
        "MONEDA": (None, "Moneda de los precios de este proveedor: vacío = pesos ($); USD = dólares. La certificación "
                   "sale con esa moneda."),
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
                "formato": str(f.get("FORMATO", "") or "").strip().upper() or _formato_por_nombre(prov),
                "moneda": moneda_de(f.get("MONEDA", ""))}
    return {"razon": prov, "cuit": "", "grupo": prov, "certifica": True, "tolerancia": None, "limite": None,
            "desc_hum": False, "desc_gran": True, "tol_finos": None, "lim_finos": None, "formato": _formato_por_nombre(prov),
            "moneda": "ARS"}


def _formato_por_nombre(prov):
    """Todos los proveedores usan el formato de la certificación final (Pinar); solo MyM tiene el suyo."""
    return "MYM" if re.sub(r"\s", "", str(prov)).upper().startswith("MYM") else ""


def _columnas_punto(precios):
    """Columnas opcionales de la hoja Precios: PUNTO (punto de entrega del maestro al que se aplica ese precio,
    dentro de la certificación del GRUPO) y ETIQUETA (cómo se llama esa línea en el cuadro de precios)."""
    for col in ("PUNTO", "ETIQUETA"):
        if col not in precios.columns:
            precios[col] = ""
        precios[col] = precios[col].fillna("").astype(str).str.strip().replace({"nan": "", "None": ""})
    return precios


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
    # precio "muy distinto al habitual": se compara dentro de la misma moneda (USD 13 no es un error frente a $ 12.500)
    moneda_g = {}
    for prov, f in conf.iterrows():
        mon = moneda_de(f.get("MONEDA", "")) if "MONEDA" in conf.columns else "ARS"
        for g in (str(f.get("GRUPO") or prov), str(prov)):
            if mon == "USD" or g.casefold() not in moneda_g:
                moneda_g[g.casefold()] = mon
    con_precio = precios.dropna(subset=["PRECIO"]).copy()
    con_precio["_moneda"] = con_precio["GRUPO"].str.casefold().map(moneda_g).fillna("ARS")
    for (t, _m), p in con_precio.groupby(["TIPOLOGIA", "_moneda"]):
        med = p["PRECIO"].median()
        if len(p) >= 3:
            for _, r in p.iterrows():
                if r["PRECIO"] > med * 3 or r["PRECIO"] < med / 3:
                    obs.append((CRIT, r["GRUPO"], f"Precio de {t} {dinero(r['PRECIO'], _m)} muy distinto al "
                                f"habitual ({dinero(med, _m)}). ¿Cero de más o de menos?"))
    return obs
