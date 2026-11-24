"""Armado de las certificaciones de un mes: cálculo por camión, precios, importe, estado y cuadre.

Parte del motor de certificaciones (antes en cert_biomasa.py). El código de cada función es el mismo de
la versión 2.1; solo se repartió en módulos."""
import json
import hashlib
import calendar
from datetime import date, timedelta

import pandas as pd

from . import base as _valores   # valores generales que cambian en marcha (corte, finos)
from .base import (
    ADV, CAMPOS, CRIT, C_DIA, C_NETO, C_PROV, C_REM_PLANTA, FILA_COLS, MYM_COLS, MYM_OTROS,
    ORDEN_PRODUCTOS, PRODUCTOS_CUADRO, SEP_VARIANTE, TIPOLOGIAS)
from .formato import decimal, etiqueta_mym, nombre_producto, pct, tipo_base
from .descuentos import descuento_finos
from .config import (
    _de_grupo, buscar_precio, corresponde, cuit_valido, info_proveedor, miembros, unidad,
    validar_config, variante_punto)
from .validacion import validar_maestro
from .laboratorio import _hermanos, asignar_finos, familia_lab, finos_por_remito_hermanos, limites_por_dia


def periodos(anio, mes, tipo="ambas"):
    """tipo: lista de períodos (ej. ["1Q", "Mensual"]) o un texto ("ambas", "1Q", "2Q+Mensual"...)."""
    ult = calendar.monthrange(anio, mes)[1]
    if isinstance(tipo, (list, tuple, set)):
        todos = {"1Q": (date(anio, mes, 1), date(anio, mes, _valores.DIA_CORTE)),
                 "2Q": (date(anio, mes, _valores.DIA_CORTE + 1), date(anio, mes, ult)),
                 "Mensual": (date(anio, mes, 1), date(anio, mes, ult))}
        return [(k, *todos[k]) for k in ("1Q", "2Q", "Mensual") if k in tipo]
    p = []
    if tipo in ("quincenal", "ambas", "1Q"):
        p.append(("1Q", date(anio, mes, 1), date(anio, mes, _valores.DIA_CORTE)))
    if tipo in ("quincenal", "ambas", "2Q+Mensual", "2Q"):
        p.append(("2Q", date(anio, mes, _valores.DIA_CORTE + 1), date(anio, mes, ult)))
    if tipo in ("mensual", "ambas", "2Q+Mensual"):
        p.append(("Mensual", date(anio, mes, 1), date(anio, mes, ult)))
    return p


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
                alto = _valores.FINOS_TOLERANCIA / 100 if alto is None else alto
                tol_f = alto * 100
                muy_alto = ip["lim_finos"] if isinstance(ip.get("lim_finos"), float) else _valores.FINOS_TOTALIDAD / 100
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
                 "tol_finos": tol_f if tol_f is not None else _valores.FINOS_TOLERANCIA,
                 "desc_gran": ip["desc_gran"] and not sin_gran, "sin_gran": sin_gran, "sin_hum": sin_hum,
                 "formato": "MYM" if es_mym else "",
                 "muy_alto": ip["lim_finos"] * 100 if isinstance(ip.get("lim_finos"), float) else _valores.FINOS_TOTALIDAD,
                 "con_humedad": humedad is not None, "metodo_humedad": metodo_humedad,
                 "numero": meta.get("numero", ""), "moneda": ip.get("moneda", "ARS"),
                 "nota": str(meta.get("nota") or "").strip(),
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
    orden = {t: i for i, t in enumerate(ORDEN_PRODUCTOS + [t for t in TIPOLOGIAS if t not in ORDEN_PRODUCTOS])}
    lista = [t for t, v in tipos.items() if v] if isinstance(tipos, dict) else list(tipos)
    return sorted(lista, key=lambda t: (orden.get(t, 99), str(t)))


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
    orden = {t: i for i, t in enumerate(ORDEN_PRODUCTOS + [t for t in TIPOLOGIAS if t not in ORDEN_PRODUCTOS])}
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
    if c.get("moneda", "ARS") != "ARS":
        data.append(["moneda", c["moneda"]])
    if c.get("nota"):
        data.append(["nota", c["nota"]])
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


def cuadre(d_orig, d, certs, conf, anio, mes, proveedores=None):
    ini = pd.Timestamp(anio, mes, 1)
    fin = ini + pd.offsets.MonthEnd(0)
    filas = []
    for per in sorted({c["periodo"] for c in certs.values()}, key=["1Q", "2Q", "Mensual"].index):
        def base(x):
            x = x[(x[C_DIA] >= ini) & (x[C_DIA] <= fin)]
            if per == "1Q":
                x = x[x[C_DIA].dt.day <= _valores.DIA_CORTE]
            elif per == "2Q":
                x = x[x[C_DIA].dt.day > _valores.DIA_CORTE]
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
