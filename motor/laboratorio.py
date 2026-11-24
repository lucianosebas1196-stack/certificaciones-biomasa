"""Planilla del laboratorio: vincular nombres, asignar muestras a cada camión, humedad del período.

Parte del motor de certificaciones (antes en cert_biomasa.py). El código de cada función es el mismo de
la versión 2.1; solo se repartió en módulos."""
import difflib

import pandas as pd

from .base import C_DIA, C_PROV, C_REM_PLANTA, C_REM_PROV, C_TIPO, MATERIAL_HUM, SIN_REMITO
from .formato import _clave, _hhmm, _minutos, _num_remito, _rem_clave, _txt
from .config import info_proveedor, puntos, unidad


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


def limites_por_dia(h):
    """Humedad de referencia del día (límite de la planilla del laboratorio, el más repetido de ese día)."""
    if h is None:
        return {}
    x = h.dropna(subset=["Limite"])
    return x.groupby(x["Fecha"].dt.date)["Limite"].agg(lambda s: float(s.round(4).mode().iloc[0])).to_dict()
