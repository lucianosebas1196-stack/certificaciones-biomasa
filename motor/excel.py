"""Excel de la certificación (con fórmulas), libro del período y resumen de control.

Parte del motor de certificaciones (antes en cert_biomasa.py). El código de cada función es el mismo de
la versión 2.1; solo se repartió en módulos."""
import io
import re

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill
from openpyxl.worksheet.page import PageMargins

from .base import ADV, CRIT, EMPRESA_LINEAS, MYM_PRODUCTOS, ICONOS, PIE
from .formato import decimal, etiqueta_mym, formato_moneda, texto_periodo
from .estilos import (
    AMARILLO, AMARILLO_CLARO, B_FINO, B_MEDIO, CENTRO, ENCAB, MEDIO, MONEDA, ROJO, ROJO_CLARO, VERDE,
    VERDE_CLARO)
from .armado import con_tramos, tn_pagar
from .pdf import criterio_mym, nota_granulometria


def _nombre_hoja(prov, usados):
    base = re.sub(r"[\[\]\*\?/\\:]", "", prov)[:31].strip() or "Proveedor"
    n, k = base, 2
    while n.casefold() in usados:
        n = f"{base[:28]}~{k}"
        k += 1
    usados.add(n.casefold())
    return n


def _hoja_certificacion(ws, c, logo=None):
    MONEDA = formato_moneda(c.get("moneda", "ARS"))           # $ o USD según el proveedor
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
    if c.get("nota"):                       # nota al pie de la certificación
        r += 2
        ws.merge_cells(start_row=r, start_column=2, end_row=r + 1, end_column=13)
        x = ws.cell(r, 2, c["nota"])
        x.alignment, x.border = Alignment(wrap_text=True, vertical="top"), B_FINO
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


def _hoja_mym(ws, c, logo=None):
    """Formato de la planilla de MyM, con las mismas fórmulas: columnas A a N."""
    MONEDA = formato_moneda(c.get("moneda", "ARS"))           # $ o USD según el proveedor
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
    enc = ["Estado", "Proveedor", "N° Certificación", "Remitos", "TN", "Importe", "Hoja", "Moneda"]
    ix.append(enc)
    usados = {"índice"}
    for c in certs:
        hoja = _nombre_hoja(c["proveedor"], usados)
        _hoja_certificacion(wb.create_sheet(hoja), c, logo)
        ix.append([f"{ICONOS[c['estado']]} {c['estado']}", c["proveedor"], c["numero"], len(c["filas"]),
                   c["tn"], c["importe"], hoja, c.get("moneda", "ARS")])
        x = ix.cell(ix.max_row, 7)
        x.hyperlink, x.style = f"#'{hoja}'!A1", "Hyperlink"
        ix.cell(ix.max_row, 6).number_format = formato_moneda(c.get("moneda", "ARS"))
        ix.cell(ix.max_row, 1).fill = VERDE_CLARO if c["estado"] == "CONFIRMADA" else ROJO_CLARO
    monedas = sorted({c.get("moneda", "ARS") for c in certs}) or ["ARS"]
    for i, mon in enumerate(monedas):       # un total por moneda (nunca se suman pesos con dólares)
        de_mon = [c for c in certs if c.get("moneda", "ARS") == mon]
        ix.append(["", "TOTAL" + (f" {mon}" if len(monedas) > 1 else ""), "",
                   sum(len(c["filas"]) for c in (certs if i == 0 else [])) or "",
                   round(sum(c["tn"] for c in certs), 2) if i == 0 else "",
                   round(sum(c["importe"] or 0 for c in de_mon), 2), "", mon])
        for x in ix[ix.max_row]:
            x.font = Font(bold=True)
        ix.cell(ix.max_row, 6).number_format = formato_moneda(mon)
    for c in ix[1]:
        c.font, c.fill = Font(bold=True, color="FFFFFF"), ENCAB
    for col, a in zip("ABCDEFGH", [16, 40, 18, 10, 12, 18, 34, 9]):
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
                      "Advertencias": sum(o["sev"] == ADV for o in c["obs"]), "Moneda": c.get("moneda", "ARS")})
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
    enc_c = [x.value for x in wb["Certificaciones"][1]]
    i_imp, i_mon = enc_c.index("Importe"), enc_c.index("Moneda")
    for row in wb["Certificaciones"].iter_rows(min_row=2):
        row[i_imp].number_format = formato_moneda(row[i_mon].value)
        row[0].fill = VERDE_CLARO if row[0].value == "CONFIRMADA" else ROJO_CLARO
    for row in wb["Cuadre"].iter_rows(min_row=2):
        row[-1].fill = VERDE_CLARO if "✔" in str(row[-1].value) else ROJO_CLARO
    wb.save(ruta)
