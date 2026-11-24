"""Colores, bordes y formatos de celda de los Excel.

Parte del motor de certificaciones (antes en cert_biomasa.py). El código de cada función es el mismo de
la versión 2.1; solo se repartió en módulos."""
from openpyxl.styles import Alignment, Border, PatternFill, Side

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


AMARILLO_CLARO = PatternFill("solid", fgColor="FFF2A8")    # camiones con descuento de granulometría
