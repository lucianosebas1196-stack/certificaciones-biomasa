"""Motor de certificaciones de biomasa, repartido en módulos:

  base          Constantes del negocio, columnas del maestro y valores generales que cambian en marcha (día de corte de la quincena y tolerancia general de finos).
  formato       Formatos y textos: números argentinos, monedas, nombres de productos, períodos, limpieza de textos.
  estilos       Colores, bordes y formatos de celda de los Excel.
  descuentos    Reglas de descuento por granulometría (finos) de cada proveedor.
  config        Config de proveedores: leer, guardar, validar; datos de cada proveedor; precios y tramos.
  lectura       Lectura y preparación del maestro de balanza; reconocimiento de archivos subidos.
  correcciones  Correcciones a mano sobre el maestro (por camión) y estado vacío de un mes.
  validacion    Controles del maestro: remitos repetidos, fechas, pesos, faltantes.
  laboratorio   Planilla del laboratorio: vincular nombres, asignar muestras a cada camión, humedad del período.
  armado        Armado de las certificaciones de un mes: cálculo por camión, precios, importe, estado y cuadre.
  logo          Logo por defecto incorporado y lectura del logo de una certificación modelo.
  pdf           PDF de la certificación (formato general y formato MyM).
  excel         Excel de la certificación (con fórmulas), libro del período y resumen de control.

Se usa a través de cert_biomasa.py (puerta de entrada), igual que antes."""
