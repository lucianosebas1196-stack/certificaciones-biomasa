"""Genera un maestro de ingresos y una planilla de laboratorio FICTICIOS (octubre 2026) para probar el sistema.

Uso:  python generar_datos_ejemplo.py   (crea los dos .xlsx en esta carpeta)
Incluye a propósito algunos errores típicos de carga (remito repetido, peso imposible, tipología mal escrita)
para ver cómo los detecta el sistema."""
import random
from pathlib import Path
from datetime import date, datetime, timedelta, time

import openpyxl

random.seed(7)
ANIO, MES = 2026, 10

# proveedor: (tipologías con peso, camiones por día hábil aprox, patente propia)
PROV = {
    "Pallets Sur EP": ({"CHIP": 8, "ASERRIN": 2}, 2.2),
    "Pallets Sur LT": ({"CHIP": 7, "ASERRIN": 3}, 1.4),
    "Pinar Maderas": ({"CHIP": 7, "ASERRIN": 2, "C LARGO": 1}, 1.8),
    "MYM": ({"CHIP": 6, "ASERRIN": 2, "Fino/Raleo": 2}, 1.6),
    "AGROMAD": ({"CHIP": 7, "ASERRIN": 2, "Fino/Raleo": 1}, 1.5),
    "AGROMAD (+10 KM)": ({"CHIP": 1}, 0.4),
    "ROBLESUR (Pérez)": ({"CHIP": 1}, 0.9),
    "LAPACHO (El monte)": ({"CHIP": 1}, 0.7),
    "Translit Norte": ({"CHIP": 4, "ASERRIN": 1}, 1.0),
    "Translit Sur": ({"CHIP": 4, "ASERRIN": 1}, 0.8),
    "Centro de chipeado Ejemplo": ({"CHIP": 4, "ASERRIN": 1}, 1.2),
    "Aserradero El Ceibo": ({"CHIP": 1}, 0.5),
}
LETRAS = "ABCDEFGHJKLMNPRSTUVWXYZ"


def patente():
    return f"{random.choice('AB')}{random.choice(LETRAS)}{random.randint(100, 999)}{random.choice(LETRAS)}{random.choice(LETRAS)}"


FLOTA = {p: [patente() for _ in range(4)] for p in PROV}
FLOTA["ROBLESUR (Pérez)"] = ["AA123BB"]
FLOTA["LAPACHO (El monte)"] = ["AA123BB"]
FLOTA_AGROMAD = FLOTA["AGROMAD"]

dias = [date(ANIO, MES, d) for d in range(1, 32)]
dias = [d for d in dias if d.weekday() < 6]          # lunes a sábado

filas, muestras = [], []
rem_planta = 45210
rem_prov = {p: random.randint(1200, 8800) for p in PROV}

for d in dias:
    camiones = []
    for p, (tipos, media) in PROV.items():
        n = max(0, round(random.gauss(media * (0.6 if d.weekday() == 5 else 1), 0.7)))
        for _ in range(n):
            camiones.append(p)
    random.shuffle(camiones)
    minuto = 6 * 60 + random.randint(0, 40)
    for p in camiones:
        tipos, _ = PROV[p]
        tipo = random.choices(list(tipos), weights=list(tipos.values()))[0]
        minuto += random.randint(8, 32)
        if minuto > 19 * 60:
            minuto = 19 * 60 - random.randint(0, 50)
        vacio = random.randint(14200, 16800)
        neto = {"CHIP": random.randint(23000, 31500), "ASERRIN": random.randint(18000, 26000),
                "Fino/Raleo": random.randint(20000, 28000), "C LARGO": random.randint(21000, 27000)}[tipo]
        pat, emp = random.choice(FLOTA[p]), p.upper()
        # algunos camiones de ROBLESUR / LAPACHO los trae AGROMAD con su flota (regla de patentes del config)
        if p in ("ROBLESUR (Pérez)", "LAPACHO (El monte)") and random.random() < 0.45:
            pat, emp = random.choice(FLOTA_AGROMAD), "AGROMAD"
        rem_planta += 1
        rem_prov[p] += 1
        filas.append({
            "DIA": datetime.combine(d, time()), "ENTRADA": minuto / 1440,
            "Nº Remito Planta": rem_planta, "PROVEEDOR NOMBRE": p, "PROVEEDOR N°REMITO": rem_prov[p],
            "TIPOLOGIA": tipo, "P BATEA": f"B-{random.randint(1, 24):02d}",
            "PESO LLENO": vacio + neto, "PESO VACIO": vacio, "PESO NETO": neto,
            "TIPO DE MOVIMINETO": "Ingreso", "Quincena": 1 if d.day <= 12 else 2,
            "PATENTE TRANSPORTE": pat, "EMPRESA TRANSPORTE": emp,
        })
        # el laboratorio muestrea ~70 % de los camiones de chip / aserrín / raleo
        if tipo in ("CHIP", "ASERRIN", "Fino/Raleo") and random.random() < 0.7:
            hum = random.gauss(0.47 if tipo != "ASERRIN" else 0.50, 0.035)
            fin = max(0.02, random.gauss(0.085, 0.035)) if tipo == "CHIP" else None
            hora = (datetime.combine(d, time()) + timedelta(minutes=minuto + random.randint(5, 35))).time()
            muestras.append({"Fecha": datetime.combine(d, time()), "Hora": hora, "PROVEEDOR": p,
                             "Tipo de proveedor": "Normal", "Material": "RALEO" if tipo == "Fino/Raleo" else tipo,
                             "REMITO": rem_prov[p], "Humedad": round(min(max(hum, 0.3), 0.62), 4),
                             "Limite de referencia humedad": 0.48 if tipo != "ASERRIN" else 0.52,
                             "< #04 (%)": round(fin, 4) if fin is not None else None,
                             "Alto": round(fin * 1.1, 4) if fin is not None else None})

# errores típicos de carga a mano, para que la demo muestre los controles
filas[40]["PROVEEDOR N°REMITO"] = filas[39]["PROVEEDOR N°REMITO"]
filas[40]["PROVEEDOR NOMBRE"] = filas[39]["PROVEEDOR NOMBRE"]
filas[95]["PESO NETO"] = 78450
filas[95]["PESO LLENO"] = filas[95]["PESO VACIO"] + 78450
filas[160]["TIPOLOGIA"] = "CHIPP"

# ---- maestro
COLS = ["DIA", "ENTRADA", "Nº Remito Planta", "PROVEEDOR NOMBRE", "PROVEEDOR N°REMITO", "TIPOLOGIA", "P BATEA",
        "PESO LLENO", "PESO VACIO", "PESO NETO", "TIPO DE MOVIMINETO", "Quincena", "PATENTE TRANSPORTE",
        "EMPRESA TRANSPORTE"]
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Datos Entrada"
ws.append(["MAESTRO DE INGRESOS – DATOS DE DEMOSTRACIÓN (FICTICIOS)"])
ws.append([])
ws.append(COLS)
for f in filas:
    ws.append([f[c] for c in COLS])
for r in ws.iter_rows(min_row=4, min_col=1, max_col=1):
    r[0].number_format = "dd/mm/yyyy"
for r in ws.iter_rows(min_row=4, min_col=2, max_col=2):
    r[0].number_format = "hh:mm"
AQUI = Path(__file__).resolve().parent
wb.save(AQUI / "maestro_ejemplo_octubre_2026.xlsx")

# ---- laboratorio
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "GENERAL"
CL = ["Fecha", "Hora", "PROVEEDOR", "Tipo de proveedor", "Material", "REMITO", "Humedad",
      "Limite de referencia humedad", "< #04 (%)", "Alto"]
ws.append(["PROVEEDORES - HUMEDAD Y GRANULOMETRIA (DEMO)"])
ws.append([])
ws.append(CL)
for m in muestras:
    ws.append([m[c] for c in CL])
wb.save(AQUI / "laboratorio_ejemplo_octubre_2026.xlsx")
print(len(filas), "camiones,", len(muestras), "muestras,", round(sum(f["PESO NETO"] for f in filas) / 1000), "TN")
