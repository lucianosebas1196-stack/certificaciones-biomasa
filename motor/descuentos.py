"""Reglas de descuento por granulometría (finos) de cada proveedor.

Parte del motor de certificaciones (antes en cert_biomasa.py). El código de cada función es el mismo de
la versión 2.1; solo se repartió en módulos."""
import pandas as pd

from . import base as _valores   # valores generales que cambian en marcha (corte, finos)


from dataclasses import dataclass


@dataclass(frozen=True)
class ReglaFinos:
    """Granulometría pactada con un proveedor (en % de finos < #04 por camión):
    hasta `tolerancia` no se descuenta; hasta `limite` se descuenta la diferencia; por encima, todo el fino.
    Lo descontado no se pierde: se paga como aserrín.

        NORMAL.descuento(12.6)      → 2.6    (ejemplo)
        NORMAL.descuento(16.8)      → 16.8   (supera 15 %: todo)
        ESPECIAL_1.descuento(9.1)   → 4.1
    """
    tolerancia: float = 10.0
    limite: float = 15.0

    def descuento(self, finos_pct):
        """% de finos medido (ej. 11.3) → % que pasa a aserrín, redondeado a 1 decimal como la planilla."""
        if finos_pct is not None and not pd.isna(finos_pct):
            finos_pct = round(float(finos_pct), 9)       # como Excel: 14,999999999 cuenta como 15
        if finos_pct is None or pd.isna(finos_pct) or finos_pct <= self.tolerancia:
            return 0.0
        if finos_pct <= self.limite:
            return round(finos_pct - self.tolerancia + 1e-9, 1)
        return round(finos_pct + 1e-9, 1)


# Tipos de proveedor de la planilla del laboratorio (REV 11)
NORMAL = ReglaFinos(10, 15)
ESPECIAL_1 = ReglaFinos(5, 10)          # también "Esp MyM"


def descuento_finos(finos_pct, tolerancia=None, totalidad=None):
    """% de finos medido → % que pasa a aserrín. Sin tolerancia/límite usa los generales (hoja Parametros)."""
    tol = _valores.FINOS_TOLERANCIA if tolerancia is None else tolerancia
    tot = _valores.FINOS_TOTALIDAD if totalidad is None else totalidad
    return ReglaFinos(tol, tot).descuento(finos_pct)
