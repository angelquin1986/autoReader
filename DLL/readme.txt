Sii4A Octubre 2005

- Cambio Automatico de transaccion en Facturación dependiendo del tipo de comprobante
- Codigo de Retención segun SRI
- Atributo de Items para venta por fraccion
- Categoriz<ación de Items (Par Venta, Materia Prima, Servicio)
- Control en devoluciones
- Control de formas de credito segun tipo de cliente
- Control de cruce de anticipos
- Reportes con ordenación configurables
- Reportes impresión Grafica
- Reportes con eportación Automático a excel
- Reportes Consolidados (varias empresa)
- Recargo automatico por item



sii4A junio/2005
-------------------------------------------------------------------------------
- kardex de item por transaccion
- igualar hora con el servidor
- cotizacion predefinida
- recargo para el calculo del iva
- item '+' para agregar otra observacion
- posibilidad de modificar la observacion del item solo para esa factura
- verificacion de fecha de comprobante con fechas de cobro/pago
- motivo de devoluciones
- en pantallas de tesoreria agregado cobrador/vendedor


sii4A abril/2005
-------------------------------------------------------------------------------
varios cambios por verificacion del codigo en CAO

sii 15/03/2005
------------------------------------------------------------
-SiiReporte: impresion de etiquetas de clientes
-SiiPrecio: mejorado imprecion de codigo de barras con fuente TrueType
-Sii4A: al momento de grabar en ivkardex se garabara con el redondeo segun cantidad de decimales
   configuradas
- Pantalla de Compras : plantilla de Items por proveedor en IVBQD
- Auditoria: no se elimina sino pasa a otra tabla RegAuditoria la cualpuede ser visualizada desde sireporte


Sii4 02-Feb-05
--------------------------------------------------------------------------------------
-Pantalla de descuentos por grupos de items y grupo de clientes en IVBQD
-Guarda el campo codigo de Usuario que Modifica el comprobante
-SiiReporte IMpresion Grafica
-SiiTols Exportacion e Importacion de Catalogo de Recargos-Descuentos




Otros extras
- Corregido  error en reporte de Pendientes x Familia
- Programando para que exista compatibilidad entre versiones
  de Sitools con respecto a importar exportar informacion
- reportes de  cobros x dias vencios x  grupos de clientes
- reportes de  cobros x Primera cuota
- Reporte de Lista de Items  se aumento la opcion de filtrar por bodega
- Reporte de Lista de Precios se aumento la excsitencia total  sumando de todas las bodegas

Sii4 13 Septiembre 2004
- Mantiene el estado despachado en el caso de regeneracion de asientos y costos siempre y
  cuando no cambie el estado a DESAPROBADO
- En SiiTools en la opción de constatación física agregada la opción de incluir los
  items no contados fisicamente
- Corregido pantalla de facturación para punto de venta.
- Arreglado encabezado para no tener problemas con direcciones ruc de los proveedores,
  en cuanto a la longitud de la cadena.
- En SiiTools arreglado problema que causa al dar dos veces clic en el boton de generar
  asientos pox lote.

Sii4 06 Septiembre 2004
------------------------------------------------------------------------------
Guarda informacion del usuario original que creo la transaccion
Reporte de Ventas Por dias de la Semana
Reportes de Cobros Realizados agregado el campo nombre del deudor
Corregido reporte de comision de venta x item x vendedor x precio 
Nuevos campos en librerias de impresion:
	Visualiza informacion de grupos de Itmes
Opcion: Visualizar documentos pendientes Vencidos hasta la vecha  o todos
Incrementado Recargo x item,   valor  de recargo predefinido para items
Incrementado campos de descuentos independientes de la comision por cada item
Nueva pantalla lista de clientes  con busqueda

              
Sii4 07 de Mayo 2004
-------------------------------------------------------------------------------------
Reportes de Balanace por Mes corregido:
	la columna de saldo Total solo suma las columnas visibles
Reporte de Ventas para Anexos:
	Generalizadas las condiciones de busqueda para  anexos.
Reporte de consumo por Familias: aumentamos informacion de costos
Repotes Kardex de Items separado costos (debe haber)
Corregido rerpote de Cartera x dias Vencidos
filtra transacciones que afectan al saldo de  cliente
Aumentado  campo Observacion en Reporte de Cartera x Cobrar con 
Fecha de Corte.
Incluir bodega en pantalla de busqueda de items IVBQD
Corregido Sibusqueda enfoques en pantalla de busqueda 
Corregido  error que se duplicaba la descripcion de las transacciones
Campos de Nombre Forma Cobro, incluido en reportes de Clientes x Cobrar
 Clientes x cobrar con Fecha de Corte, Preveedores x Pagas  y Proveedores  
 x Pagar con fecha de Corte.




Sii4 05/04/2004
----------------------------------------------------------------------------------
Corregido en pantalla de facturacion ficha F6 para que recupere todos los decimales
Programado proceso  Generar Egresos  en Pantalla de Produccion
Depurado modulo  de Produccion costeo de producto Terminado
Reportes:
 Reporte de Saldos permite filtrar Total/ Anticipados / Por Cobrar 
 Cobro de Cartera Vencida y por Fecha de Prov Cli 
 incluye forma de cobro y opcion de agrupado por
Cambios en SiiRerpote opcion de ocultar costos
Crear nuevos grupos de Inventario desde pantalla de Nuevo Item
Actualizados campos en gnprintg con los de gnprinta,  actualizados documentos
Actualizados manuales de Sii4
Control de limite de descuento en descuento por item

Sii4  19 Marzo 2004
---------------------------------------------------------------------------------------------------
SiiPrecio opcion para modificar los porcentajes de comision de items
Bodega predeterminada para la creacion de Items
Corregido  error en reportes Cartera x cobrar con Fecha de Corte activar solo plazo vencido
Agregado impresion de resumen en reportes: Documentos Bancarios, Balance x Centro de costo y
         Resultado x centro de Costo 
Campo  nuevo en seccion items imprime la existencia
Corregido error interno actualiza bandProrrateado  en IVKardexReccargo
al modificar transacciones. 
Corregido  error  al recuperar IVKardexRecargo  
Incluir nuevamente cambios en dll  para trabajar  en Asisentos por lote 
Incluir nuevamente optimizacion en proceeso de borrado para registros de auditoria
Incluidos Manuales del Ususario en el Instalador
Incluido pkzip  en instalador  de Sii4
Correccion de Reporte Evaluacion de Vendores (consula 3000)
Error  corregido se perdia  el valor en kardex recargo, descuento (consulta 3010)
Siitools nueva version  con  posibilidad de guardar perfiles de ususario
para  importacion  y exportacion
Corregido reporte de Balance por Mes SiiReporte (generado por CT Local consulta 3020)
SiiTools  correcion de procecso  actualizar costos despues de transferir
       la fecha de la trans fuente - se borraba  el recargo de la transaccion
Reporte Compras con Retencion  sale en le reporte asi no haya  enlacee  de transaccion de DocAsignado
Coreccion Sii4 crear proveedor con XML  pantalla IVQD
Corregido Balance consolidado  en SiiConfig (por CT Local)
Copiar CTlocal  en ctas contables para importar datos ojo esto puede producir problemas 
si no actualiza en todos los locales
Impresion de Fecha de Corte en Reporte de CarteraxCDiasVencidos

Sii4 25 Febrero 2004
--------------------------------------------------------------------------------------------------
No permite espacios  al inicio en codigos de clientes proveedores
corregido  mensaje de error Numero de RUC/CI incorrecto
Corregido Reporte de Compras x Mes x Item 
Corregido encabezado busqueda recuperacion de datos de consumidor final
SiiTools proceso de actualizar costos en transacciones importadas
	tambien  actualiza la fecha de la transaccion
Reporte de Movimiento de Items 2 corregido
gnprinta impresion de campo NomTrans (campo nombre en el encabezado de la facura)
Nuevo Reporte de Ventas x Numero Precio
Actualizado SiiEsquema
Corregido error en Balance general 

10 Febrero 2004 Errores detectados  en version anterior
---------------------------------------
Mejorado SiiTools  en control  de errores importacion de datos gnComprobante
Corregido error en Consulta de Asientos 2960 vwConsCTDiario
Corregido Error  en recuperar transacciones  2970 ALTER spConsIVKardexRecargoMod
Corregido problema de versiones en flexgrid ancho de las columnas
Modificacion SiiXml abrir archivos diferentes por transaccion ejemplo: XMLCFC.xml
Corregido error  en pantalla IVBQD busqueda de items seleccion
Configuracion de Ocultar existencia debe servir para pantalla IVBQD de busqueda de items


Sii4 cambios  version 30/01/2004
----------------------------------------
SiiTools  controla cambio de codigos en catalogos dentro de proceso de importacion
Restriccion  para  asignar  Hijos a Familias(no permite duplicar)
Libreria de Impresion corregido campo TotalBanco de transaccion Bancos
Libreria gnprinta campo Observacion de item
Guardar resultados del proceso de importacion Exportacion
En Siitools  actualizacion de Familias.
Nuevos  cambios  en Modulo contabilidad
 Campo  Local  que permite  clasificar a las  cuentas contables 
 por locales  para  filtrar reportes.
 Herramienta  en SiTools  que permite  asignar locales  a las Ctas. Contables
SiiReporte nuevo reporte estado de cuenta por forma de cobro pago 
(me indica cuanto debe y cuento  a pagado  bancos) ?ojo nombre reporte
Reporte comisiones por vendedor por precio de items

Proceso  para Corregir  Reporcesamientos de Costos 
 Arreglar IVA de items despues de reprocesamientos de Costos 
 para transacciones con calculo basado en costos (SiiTools). 
Proceso de Importacion de Transacciones de Inventarios permite
actualizar costos a transacciones relacionadas

----------Correcciones de la Version  anterior---------------------------------
 En Siireporte corregida  la impresion de resumen  para  reportes de Anexos
 SQL  que  permite corregir  las configuraciones de columnas
 en pantalla  de Inventarios  por  Precio + IVA
 Corregir detalles  en  nueva  pantalla  de TSIE
  Antes  tenia problemas al importar transacciones
  Optimizado  el campo descripcion
  Corregido  reporte de Comisiones por vendedor por items (el grupo y el calculo)
  Corregido problema cuando transaccion va directo  al banco.

Sii4 cambios  version 05/01/2004
----------------------------------------
Pantalla de TSIE (tesoreria ingreso egreso mejorada)
permite realizar cobros a clientes mas facilmente
Columna de precio+iva
SiiTools programado nuevos  campos en cierre  de ejercicio
Reporte de Comision de vendedores   / mercaderia pendiente

Sii4  cambios reliazados version 20/12/2003
-------------------------------------------
No se registraron los cambios

Sii4  cambios reliazados version 16/12/2003
----------------------------------------------
Corregir Encabezados (encabezadoBusqueda) seleccion de porveedor
Corregir Recuperacion de recargos  en compras
Mejorar  tiempo de respeusta  en abrir  factura
Corregir  campos  de venta minima en reporte de Evaluacion Vendedores
Reporte de Ventas  x Hora  permite ordenar  por cualquier campo
Nuevo Reporte de Analisis de compra venta de item Resumido 
Impresion de campo Periodo Contable
SiiTools  mejorar mensaje  de error  en proceso de importacion


Sii4  cambios  realizados  version  09/12/2003
----------------------------------------------
Pantalla de facturacion IVBQD mejorada
SiiReporte Modificar reportes de  items filtro  
avanzado  de busqueda de grupos (en niveles)
Reporte de Evaulacion de Vendedores
Reporte de ventas por hora
Reporte de Analisis de Items x Compra y Venta
SiiTools  Actualizacion  de Monto de Ventas
Descuento por  item  predeterminado  en Sii4
Impresion campos  de Items  Observacion  y detalles tecnicos



Sii4  cambios  realizados  version  05/11/2003
-----------------------------------------------
Agregar  en el instalador  Documento de Ayuda de SiiXML
SiiReporte  guarda configuraciones personalizadas de reportes
Corregido  descuento por PcprovCli
SiiTool corregido  problema de exportacion  orden por fecha/hora
SiiTool importacion activado  Anexos  control  dee importacion 
datos de Anexos
Reporte por e-mail y por cumpleaños
Reportes de Anexos  mejorados



Sii4  cambios  realizados  version  22/10/2003
-----------------------------------------------
Llevar  documentos  de ayuda  para librerias de impresion
Corregir  errores  en pantalla de detallesrecargos descuentos
Incluir nuevo  Siibusqueda  y reportes para  SiiAnexos
Corregidos  detalles de importacion



Sii4  cambios  realizados  version  20/10/2003
-----------------------------------------------
Catalago  de Detalle Recargo/Descuento
Descuento  por PCGrupo
Reporte Clientes invluye  nuevos campos (e-mail  fecha cumpleaños)
SiiXML  para  configurar  paantalla  de nuevos  cleintes  usando  archivo de XML
Correccion de error  al presionar  Siguiente   sin permiso para Crear Nuevo
Cambios  en pantalla IVBQD - mostrar  campos de PCGrupo, depurar ingreso
de Nombre,Direccion, telefono,Ruc,   en  el campo descripcion
Corregido  error  en reporte de Madurez de Cartera
Correcion de error en transacciones que tienen importacion requerida
y ponen cancelar.
Reporte de clientes para anexos
Comunicación visual cuando un cliente cumple años 
Corregido error  que no se repita efectivo en Descricpion

Sii4  cambios  realizados  version  16/09/2003
-----------------------------------------------
Enviar  fuente para  pnatalla  de vuelto ASTUTE.TTF  instalar  manualmente  
si  el SO no lo  hace automatico
Corregido problema  que no permitia decimal  en vuelto
corregido  problema de validacion en  combo.




Sii4  cambios  realizados  version  15/09/2003
-----------------------------------------------
Corregido  error,  ya no se activa  el boton de grabar F3deshabilitado
Cambios de Angel  nuevos  campos  en tabla  Clientes
Control de IVVerificaCobroConsfinal para uqe no se pueda dar credito a consumidor final
Pantalla de Vuelto  para pantalla IVBQD



SII4 Errores  corregidos  Version  20/08/2003
---------------------------------------------
Formula IVACTIVO2
Pantalla  de lista  de Items  no  saca    decimales  (Nomeda_PRE)
Compatibilidad  de Librerias  de Impresion  EFSA
Incluir  libreria EfsaPrinta.dll  en instalador
Corregir  RITGH  JOIN  de SIIPrecio

ANgel
-----------------------------------------------------
Cambios  para numero de decimañles:
   El cls modificado fue GNComprobante en los métodos: 
   RedondearAsiento y IVKardexPTotal.

Oliver
-------------------------------------------------------
Acutalizar  Reporte  de Anexos  con  Campo  de RUC






Contactenos.
Ishida & Asociados
Muñoz Vernaza 12-09 Y Tarqui  
07-2826197   07-2833766
Cuenca-Ecuador
email  ishida@etapaonline.net.ec
