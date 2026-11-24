"""Formatos y textos: números argentinos, monedas, nombres de productos, períodos, limpieza de textos.

Parte del motor de certificaciones (antes en cert_biomasa.py). El código de cada función es el mismo de
la versión 2.1; solo se repartió en módulos."""
import re
import calendar

import pandas as pd

from .base import MESES, NOMBRE_PRODUCTO, SEP_VARIANTE, SIN_REMITO, TIPOLOGIAS, _AMP_PDF, _ETIQUETA_PDF


def _txt(x):
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    if isinstance(x, float) and x.is_integer():
        return str(int(x))
    return str(x).strip()


def _limpio(serie):
    return serie.apply(_txt).str.replace(r"\s+", " ", regex=True).str.strip()


def _canon_tipo(t):
    canon = {x.casefold(): x for x in TIPOLOGIAS}
    t = _txt(t)
    return canon.get(t.casefold(), t.upper() or "SIN TIPO")


def _patente(s):
    return re.sub(r"[^A-Z0-9]", "", str(s).upper())


def moneda_de(texto):
    """'USD', 'U$S', 'dólares'... → 'USD'; vacío o cualquier otra cosa → 'ARS' (pesos)."""
    t = str(texto or "").strip().upper().replace("Ó", "O")
    return "USD" if t and ("US" in t.replace("$", "S") or "DOLAR" in t) else "ARS"


def dinero(x, moneda="ARS"):
    """Importe con su moneda: '$ 1.234,50' o 'USD 1.234,50'."""
    return ("USD " if moneda == "USD" else "$ ") + decimal(x) if x is not None else "—"


def formato_moneda(moneda="ARS"):
    """Formato de celda de Excel para la moneda."""
    return '"USD"\\ #,##0.00' if moneda == "USD" else '"$"\\ #,##0.00'


def tipo_base(t):
    return str(t).split(SEP_VARIANTE, 1)[0]


def _num_remito(s):
    m = re.fullmatch(r"\d{6}-(\d+)", str(s))
    return int(m.group(1)) if m else None


def _minutos(v):
    """Hora de la muestra → minutos desde las 0 h (acepta time, datetime, fracción de día o texto hh:mm)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return float("nan")
    if hasattr(v, "hour"):
        return v.hour * 60 + v.minute
    if isinstance(v, (int, float)):
        return float(v) * 1440 if v < 1 else float("nan")
    m = re.match(r"(\d{1,2})[:.](\d{2})", str(v))
    return int(m.group(1)) * 60 + int(m.group(2)) if m else float("nan")


def _rem_clave(s):
    """Remito del proveedor comparable: '000-465', '0000-465' y '465' son el mismo; S/R → ''."""
    t = str(s).strip().upper()
    if t in SIN_REMITO or t in ("", "NAN", "NONE"):
        return ""
    nums = re.findall(r"\d+", t)
    return str(int(nums[-1])) if nums else t


def _hhmm(minuto):
    return "" if minuto is None or pd.isna(minuto) else f"{int(minuto) // 60:02d}:{int(minuto) % 60:02d}"


def _clave(s):
    return re.sub(r"[^a-z0-9]", "", str(s).casefold())


def pct(x, n=2):
    return "—" if x is None else decimal(x * 100, n) + " %"


def nombre_producto(t):
    base, _, etiqueta = str(t).partition(SEP_VARIANTE)
    return etiqueta or NOMBRE_PRODUCTO.get(base, base)


def _ar(s):
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def miles(x):
    return _ar(f"{float(x):,.0f}")


def decimal(x, n=2):
    return _ar(f"{float(x):,.{n}f}")


def pesos(x):
    return "$ " + decimal(x) if x is not None else "—"


def texto_periodo(desde, hasta):
    if desde.day == 1 and hasta.day == calendar.monthrange(hasta.year, hasta.month)[1]:
        return f"Período: {MESES[desde.month]} {desde.year} (del 01 al {hasta.day:02d})"
    return f"Período del {desde.day} al {hasta.day} de {MESES[desde.month]} {desde.year}"


def etiqueta_periodo(per):
    return {"1Q": "Quincenal", "2Q": "Quincenal", "Mensual": "Mensual"}[per]


def kilos_txt(x):
    """16.400 · 11.304,5 (como la planilla)."""
    x = float(x)
    return miles(x) if abs(x - round(x)) < 1e-6 else decimal(x, 1)


def _texto_pdf(t):
    """Texto para el PDF: un '<' o '&' que venga de un nombre o remito (ej. 'A <B>' o 'Pérez & Hijos') se muestra
    tal cual en vez de romper el PDF; las marcas de formato propias (<b>, <br/>…) se respetan."""
    t = "" if t is None else str(t)
    return _ETIQUETA_PDF.sub("&lt;", _AMP_PDF.sub("&amp;", t))


def etiqueta_mym(t):
    return "ROLLO" if t == "Fino/Raleo" else nombre_producto(t)


def _pct_txt(v):
    return "" if v is None or pd.isna(v) else f"{v * 100:.2f}%".replace(".", ",")


def _kg2(v):
    return "" if v is None or pd.isna(v) else decimal(v, 2)
