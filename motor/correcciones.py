"""Correcciones a mano sobre el maestro (por camión) y estado vacío de un mes.

Parte del motor de certificaciones (antes en cert_biomasa.py). El código de cada función es el mismo de
la versión 2.1; solo se repartió en módulos."""
import pandas as pd

from .base import (
    CAMPOS, C_BATEA, C_DIA, C_LLENO, C_MOV, C_NETO, C_PROV, C_REM_PLANTA, C_REM_PROV, C_TIPO, C_VACIO,
    TIPOLOGIAS)


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
