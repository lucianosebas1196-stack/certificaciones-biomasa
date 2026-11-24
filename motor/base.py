"""Constantes del negocio, columnas del maestro y valores generales que cambian en marcha (día de corte de la quincena y tolerancia general de finos).

Parte del motor de certificaciones (antes en cert_biomasa.py). El código de cada función es el mismo de
la versión 2.1; solo se repartió en módulos."""
import re


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


TIPOLOGIAS = ["CHIP", "ASERRIN", "Fino/Raleo", "C LARGO", "C LARGO SECO", "C CORTO", "PATAS",
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
                   "C LARGO SECO": "C. LARGO SECO",
                   "Fino/Raleo": "FINO/RALEO", "PATAS": "PATAS", "CENIZAS": "CENIZAS", "DESECHOS": "DESECHOS",
                   "RAMAS": "RAMAS"}


PRODUCTOS_CUADRO = ["CHIP", "ASERRIN", "C CORTO", "C LARGO", "Fino/Raleo"]          # rango habitual por camión (kg)


# Orden en el cuadro de precios: C. LARGO SECO (tablas) va pegado a C. LARGO, y solo aparece si se entregó
ORDEN_PRODUCTOS = ["CHIP", "ASERRIN", "C CORTO", "C LARGO", "C LARGO SECO", "Fino/Raleo"]


SECO, LARGO = "C LARGO SECO", "C LARGO"          # un camión de C LARGO se puede marcar "seco" (otro precio)


# Moneda de cada proveedor (columna MONEDA del config): vacío o ARS = pesos; USD = dólares
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


def fijar_finos(tolerancia, totalidad):
    global FINOS_TOLERANCIA, FINOS_TOTALIDAD
    FINOS_TOLERANCIA, FINOS_TOTALIDAD = float(tolerancia), float(totalidad)


def fijar_corte(dia):
    """Define el último día de la 1ª quincena (la 2ª empieza el día siguiente)."""
    global DIA_CORTE
    DIA_CORTE = int(dia)


FRECUENCIAS = {"QUINCENAL": ("1Q", "2Q"), "MENSUAL": ("Mensual",), "AMBAS": ("1Q", "2Q", "Mensual"), "": ("1Q", "2Q", "Mensual")}


SEP_VARIANTE = "@"        # clave de precio por punto de entrega: "CHIP@CHIP - 50 KM" (tipología@etiqueta del cuadro)


# Se usan los valores que calcula la propia planilla en la hoja GENERAL, por muestra:
#   Humedad = máx(horno, termobalanza) · Límite de referencia humedad (según el tipo de
#   proveedor: Normal = referencia chipera + 3 %, Especial 1 = 45 %, etc.)
# Descuento por período = promedio de humedad de las muestras del proveedor y producto
# en ese período contra el promedio de sus límites. Si no supera el límite, no hay descuento.
MATERIAL_HUM = {"CHIP": "CHIP", "ASERRIN": "ASERRIN", "RALEO": "Fino/Raleo", "FINO/RALEO": "Fino/Raleo"}


METODOS_HUMEDAD = {"PUNTOS": "TN × puntos de humedad sobre el límite",
                   "BASE_SECA": "TN equivalentes a la humedad límite: TN × (1 − H) / (1 − L)"}


# Formato MyM: humedad y granulometría camión por camión (planilla "Certificación MyM")
MYM_COLS = ["Hum", "HumRef", "Pen", "KgChip", "KgAse", "KgCorto", "KgLargo", "KgRaleo"]


MYM_OTROS = {"C CORTO": "KgCorto", "C LARGO": "KgLargo", "Fino/Raleo": "KgRaleo"}


ICONOS = {"CONFIRMADA": "✅", "MODIFICADA": "🔁", "BLOQUEADA": "⛔", "REVISAR": "⚠️", "LISTA": "🟢"}


_ETIQUETA_PDF = re.compile(r"<(?!/?(?:b|i|u|br|font|super|sub|strike)\b)")


_AMP_PDF = re.compile(r"&(?!(?:[a-zA-Z]+|#\d+);)")


MYM_PRODUCTOS = ["CHIP", "ASERRIN", "C CORTO", "C LARGO", "Fino/Raleo"]


MYM_COL_KG = {"CHIP": "KgChip", "ASERRIN": "KgAse", "C CORTO": "KgCorto", "C LARGO": "KgLargo", "Fino/Raleo": "KgRaleo"}
