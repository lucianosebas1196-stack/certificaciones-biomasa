"""PDF de la certificación (formato general y formato MyM).

Parte del motor de certificaciones (antes en cert_biomasa.py). El código de cada función es el mismo de
la versión 2.1; solo se repartió en módulos."""
import io

import pandas as pd

from . import base as _valores   # valores generales que cambian en marcha (corte, finos)
from .base import EMPRESA_LINEAS, MYM_PRODUCTOS, PIE
from .formato import (
    _kg2, _pct_txt, _texto_pdf, decimal, dinero, etiqueta_mym, kilos_txt, nombre_producto, pct, pesos,
    texto_periodo)
from .armado import _txt_limite, con_tramos, mostrar_humedad, tn_pagar


def _con_desc(f):
    """El camión tiene descuento (granulometría o, en MyM, humedad)."""
    pen = f.get("Pen", 0) if hasattr(f, "get") else 0
    return f["DescKg"] > 0 or (pen is not None and not pd.isna(pen) and pen > 0)


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
         f"de {decimal(c['tol_finos'], 0)} % a {decimal(c.get('muy_alto', _valores.FINOS_TOTALIDAD), 0)} % se descuenta la diferencia; "
         f"desde {decimal(c.get('muy_alto', _valores.FINOS_TOTALIDAD), 0)} % la totalidad. Lo descontado se paga como aserrín.")
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
    pesos = lambda x: dinero(x, c.get("moneda", "ARS"))       # $ o USD según el proveedor
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
    nota = []
    if c.get("nota"):                       # nota al pie escrita en la certificación (ej. excedente / bonificación)
        nt = Table([[P(_texto_pdf(c["nota"]).replace("\n", "<br/>"), 9)]], colWidths=[ancho - 4 * mm])
        nt.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.8, colors.black), ("LEFTPADDING", (0, 0), (-1, -1), 6),
                                ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 5)]))
        nota = [Spacer(1, 4 * mm), nt]
    # Cuadro de precios, total a facturar y firmas siempre juntos (si no entran, pasan a la hoja siguiente)
    el.append(KeepTogether(bloque + [fact] + nota + [Spacer(1, 6 * mm), firmas, Spacer(1, 3 * mm)]
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


def criterio_mym(c):
    alto, muy = decimal(c["tol_finos"], 0), decimal(c["muy_alto"], 0)
    return (f'Cuando el contenido de aserrín es superior al valor de "ALTO" ({alto} %) e inferior al valor de '
            f'"MUY ALTO" ({muy} %) se cuentan como aserrín los kilos que superen el valor de "ALTO". En caso de '
            f'igualar o superar el valor de "MUY ALTO" se cuenta como aserrín el total medido.',
            "Cuando la humedad excede el valor de referencia del día (chipeado propio de raleo fresco) se descuenta "
            "el porcentaje con el cual supera ese valor.")


def pdf_mym(ruta, c, logo=None):
    pesos = lambda x: dinero(x, c.get("moneda", "ARS"))       # $ o USD según el proveedor
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
    nota = [P(_texto_pdf(c["nota"]).replace("\n", "<br/>"), 8), Spacer(1, 3 * mm)] if c.get("nota") else []
    der = [tot, Spacer(1, 4 * mm)] + nota + [firmas, Spacer(1, 3 * mm)] + [P(t, 6.5, color=rojo) for t in PIE]
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
