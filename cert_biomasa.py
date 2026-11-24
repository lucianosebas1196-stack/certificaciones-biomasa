"""PUERTA DE ENTRADA AL MOTOR — Certificaciones de entrega de biomasa.

El código está en la carpeta motor/ (un módulo por tema). Este archivo reexporta todo para que servidor.py,
editor_colab.py, las pruebas y el cuaderno de Colab sigan usando `import cert_biomasa as cb` sin cambios.

Los valores que cambian en marcha (DIA_CORTE, FINOS_TOLERANCIA, FINOS_TOTALIDAD) viven en motor/base.py y se
leen a través de __getattr__: cb.DIA_CORTE devuelve siempre el valor actual."""
from motor.base import (  # noqa: F401
    ADV, CAMPOS, CRIT, C_BATEA, C_DIA, C_LLENO,
    C_MOV, C_NETO, C_PATENTE, C_PROV, C_QUINC, C_REM_PLANTA,
    C_REM_PROV, C_TIPO, C_VACIO, EMPRESA_LINEAS, FILA_COLS, FRECUENCIAS,
    MYM_COLS, MYM_COL_KG, MYM_OTROS, MYM_PRODUCTOS, ICONOS, LARGO,
    MATERIAL_HUM, MESES, METODOS_HUMEDAD, MOVIMIENTOS, NOMBRE_PRODUCTO, ORDEN_PRODUCTOS,
    PESO_MAX, PESO_MIN, PIE, PRODUCTOS_CUADRO, SECO, SEP_VARIANTE,
    SIN_REMITO, TIPOLOGIAS, _AMP_PDF, _ETIQUETA_PDF, fijar_corte, fijar_finos)
from motor.formato import (  # noqa: F401
    _ar, _canon_tipo, _clave, _hhmm, _kg2, _limpio,
    _minutos, _num_remito, _patente, _pct_txt, _rem_clave, _texto_pdf,
    _txt, decimal, dinero, etiqueta_mym, etiqueta_periodo, formato_moneda,
    kilos_txt, miles, moneda_de, nombre_producto, pct, pesos,
    texto_periodo, tipo_base)
from motor.estilos import (  # noqa: F401
    AMARILLO, AMARILLO_CLARO, B_FINO, B_MEDIO, CENTRO, ENCAB,
    FINO, MEDIO, MONEDA, NARANJA, NARANJA_DESC, ROJO,
    ROJO_CLARO, VERDE, VERDE_CLARO)
from motor.descuentos import (  # noqa: F401
    descuento_finos)
from motor.config import (  # noqa: F401
    _columnas_punto, _de_grupo, _formato_por_nombre, _porcentajes_excel, _porcentajes_hoja, buscar_precio,
    corresponde, cuit_valido, fijar_precio, frecuencia, guardar_config, info_proveedor,
    leer_config, miembros, plantilla_config, puntos, unidad, validar_config,
    variante_punto)
from motor.lectura import (  # noqa: F401
    leer_maestro, preparar, separar_por_patente, tipo_archivo)
from motor.correcciones import (  # noqa: F401
    aplicar_correcciones, estado_nuevo)
from motor.validacion import (  # noqa: F401
    validar_maestro)
from motor.laboratorio import (  # noqa: F401
    _hermanos, _muestra, _orden_entrada, asignar_finos, calcular_humedad, familia_lab,
    finos_por_remito_hermanos, leer_humedad, limites_por_dia, mapear_humedad)
from motor.armado import (  # noqa: F401
    _etiqueta_tramo, _txt_limite, armar_tramos, calcular_estandar, calcular_mym, con_tramos,
    construir, criticas_revisables, cuadre, estado_cert, huella, importe,
    mostrar_humedad, orden_productos, periodos, tn_pagar)
from motor.logo import (  # noqa: F401
    LOGO_B64, extraer_logo, logo_por_defecto)
from motor.pdf import (  # noqa: F401
    _con_desc, criterio_mym, etiqueta_producto, nota_granulometria, pdf_certificacion, pdf_mym,
    texto_granulometria, tn_detalle)
from motor.excel import (  # noqa: F401
    _hoja_certificacion, _hoja_mym, _nombre_hoja, escribir_libro, escribir_resumen, excel_certificacion)
from motor import base as _base


def __getattr__(nombre):
    if nombre in ("DIA_CORTE", "FINOS_TOLERANCIA", "FINOS_TOTALIDAD"):
        return getattr(_base, nombre)
    raise AttributeError(f"module 'cert_biomasa' has no attribute '{nombre}'")
