"""Controles del maestro: remitos repetidos, fechas, pesos, faltantes.

Parte del motor de certificaciones (antes en cert_biomasa.py). El código de cada función es el mismo de
la versión 2.1; solo se repartió en módulos."""
import difflib

import pandas as pd

from .base import (
    ADV, CRIT, C_DIA, C_LLENO, C_MOV, C_NETO, C_PROV, C_QUINC, C_REM_PLANTA, C_REM_PROV, C_TIPO,
    C_VACIO, MESES, MOVIMIENTOS, PESO_MAX, PESO_MIN, SIN_REMITO)
from .formato import _limpio, _num_remito, _txt, miles


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
