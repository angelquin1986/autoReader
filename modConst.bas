Attribute VB_Name = "modConst"
    
Option Explicit

'Constantes
Public Const TIPODB_JET = 0         'Base de datos mdb
Public Const TIPODB_SQL = 1         'Base de datos SQLServer
Public Const APPNAME = "Ishida y Asociados"  'Clave de registro de sistema
Public Const SECTION = "Sii"                 'Clave de registro de sistema
Public Const IVGRUPO_MAX = 6        'Hasta 6 grupos de inventario
Public Const AFGRUPO_MAX = 5
Public Const GNCGRUPO_MAX = 4
Public Const GNVGRUPO_MAX = 4
Public Const PCGRUPO_MAX = 4        'Hasta 4 grupos de Proveedor/Cliente Modificado AUC 01/11/2005
Public Const MONEDA_MAX = 4         'Hasta 4 monedas
Public Const MAXTRANS_DEMO = 200    'Hasta 200 transacciones en versión DEMO
Public Const TECLA_CLICKDERECHO = 93    'Tecla que equivale a Click derecho '*** MAKOTO 30/nov/00

'Clave para cambiar Fecha de caducidad
'Public Const KEY_CADUCIDAD = "1596227#2KLKLÑSA987!%"
'Public Const KEY_CADUCIDAD = "2607336#2KMKMÑSB876!%"
'2014
Public Const KEY_CADUCIDAD = "4829558#4MÑMÑPUD098¡%"


'Constantes que indica tipo de PCKardex
Public Const TIPO_PCK_PORCOBRAR = 1
Public Const TIPO_PCK_PORPAGAR = 2
Public Const TIPO_PCK_COBRO = 3
Public Const TIPO_PCK_PAGO = 4

''Tipo de moneda
Public Const MONEDA_PRE = "SUCRE"       'Moneda predeterminada
Public Const MONEDA_SEC = "USD"       'Moneda predeterminada


'Tipo de costeo de inventario
Public Const COSTO_PROM = 0
Public Const COSTO_FIFO = 1
Public Const COSTO_LIFO = 2
Public Const COSTO_ULTIMO = 3

'Tipo de Depreciacion
Public Const DEP_ACELERADA = 0
Public Const DEP_DESACELERADA = 1
Public Const DEP_LINEAL = 2




'Estado de combprobante     0=No aprobado, 1=Aprobado, 2=Despachado, 3=Anulado
Public Const ESTADO_NOAPROBADO = 0
Public Const ESTADO_APROBADO = 1
Public Const ESTADO_DESPACHADO = 2
Public Const ESTADO_ANULADO = 3
Public Const ESTADO_SEMDESPACHADO = 4        '17/05/2006

'Constantes para el campo BandOrigen de GNTransRecargo, IVKardexRecargo
Public Const REC_SUMA = -1
Public Const REC_TOTAL = 0
Public Const REC_IVAITEM = -2
Public Const REC_SUBTOTAL = -3
Public Const REC_RECITEM = -4   '***Agregado. Angel. 29/jul/2004
Public Const REC_ICEITEM = -5   '***Agregado. JEAA 20/07/2006
Public Const REC_SUMAIVAITEM = -6   '***Agregado. JEAA 25/09-2006
Public Const REC_FINANCIAMIENTO = -7   '***Agregado. JEAA 24/03/2009
Public Const REC_ARANCELITEM = -8   '***Agregado. JEAA 24/03/2009
Public Const REC_FODINITEM = -9   '***Agregado. JEAA 24/03/2009
Public Const REC_ICEITEMIMPORT = -10
Public Const REC_SOBREARANCELITEM = -11
Public Const REC_SEGXTRANS = -12 'por transaccion relacionado con el cat gnseguro

'Constantes que indica contenido de BandIntegridad de CTLibroDetalle
Public Const INTEG_NADA = 0         'No está verificado
Public Const INTEG_AF = 1           'Conflicto con Activos Fijos
Public Const INTEG_TS = 2           'Conflicto con Bancos
Public Const INTEG_IV = 3           'Conflicto con Inventario
Public Const INTEG_PC = 4           'Conflicto con proveedores/clientes
Public Const INTEG_RL = 5           'Conflicto con Roles
Public Const INTEG_AUTO = 100       'Generado automáticamente por GeneraAsiento.
Public Const INTEG_INTEGRADO = 101  'Integrado por usuario manualmente.

'Códigos y mensajes de errores generados en DLL
Public Const ERRNUM = vbObjectError + 531
Public Const ERR_NOREGINFO = ERRNUM + 1
Public Const ERR_INVALIDO = ERRNUM + 2
Public Const MSGERR_INVALIDO = "El valor es inválido."
Public Const ERR_MODIFICADO = ERRNUM + 3
Public Const MSGERR_MODIFICADO = "El registro fue modificado por otro usuario."
Public Const ERR_TEMPDB = ERRNUM + 4
Public Const MSGERR_TEMPDB = "Ha ocurrido un error de la base temporal."
Public Const ERR_NOMODIFICABLE = ERRNUM + 5
Public Const MSGERR_NOMODIFICABLE = "La propiedad no es modificable."
Public Const ERR_NOELIMINABLE = ERRNUM + 6
Public Const MSGERR_NOELIMINABLE = "El registro no es eliminable."
Public Const ERR_NODERECHO = ERRNUM + 7
Public Const MSGERR_NODERECHO = "No tiene derecho para la operación."
Public Const ERR_REPITECODIGO = ERRNUM + 8
Public Const MSGERR_REPITECODIGO = "Ya existe el código. Por favor utilíce otro código."
Public Const ERR_DESCUADRADO = ERRNUM + 9
Public Const MSGERR_DESCUADRADO = "El asiento está descuadrado."
Public Const ERR_INTEGRIDAD = ERRNUM + 10
Public Const MSGERR_INTEGRIDAD = "El asiento no está integrado."
Public Const MSGERR_INTEGRIDAD2 = _
                "Los siguientes detalles de asiento causarán una " & _
                "desintegración con otro modulo del sistema."
Public Const ERR_NOUsuario = ERRNUM + 11
Public Const MSGERR_NOUsuario = "Usuario no está establecido."
Public Const ERR_SOLOVER = ERRNUM + 12
Public Const MSGERR_SOLOVER = "El registro no es modificable."
Public Const ERR_NOHAYCODIGO = ERRNUM + 13
Public Const MSGERR_NOHAYCODIGO = "No se encuentra el código."
Public Const ERR_COTIZACION = ERRNUM + 14
Public Const MSGERR_COTIZACION = "La cotización no puede ser 0 o negativa."
Public Const ERR_NOIMPORTA = ERRNUM + 15
Public Const MSGERR_NOIMPORTA = _
                "No se puede importar debido a que la transacción de orígen no está aprobada."
Public Const ERR_NOGRABADO = ERRNUM + 16
Public Const MSGERR_NOGRABADO = "El cambio que hizo aún no está grabado." & vbCr & _
                                "Primero guárdela e intente de nuevo."
Public Const ERR_NOIMPRIME = ERRNUM + 17
Public Const MSGERR_NOIMPRIME = _
                "La transacción debe estar aprobada para imprimir."
Public Const ERR_NOIMPRIME2 = ERRNUM + 18
Public Const MSGERR_NOIMPRIME2 = _
               "No se puede preparar la impresión." & vbCr & _
               "Revíse la configuración de la transacción si está seleccionada la librería adecuada."
Public Const ERR_NOIMPRIME3 = ERRNUM + 19
Public Const MSGERR_NOIMPRIME3 = _
                "No se puede imrimir debido a que la transacción está anulada."
Public Const ERR_CADUCADO = ERRNUM + 20
Public Const MSGERR_CADUCADO = _
                "No se puede usar el sistema debido a que ya está caducado." & vbCr & _
                "Contácte con su proveedor para continuar trabajándo con el sistema."
Public Const ERR_LIMITEITEM = ERRNUM + 21
Public Const ERR_NODAO = ERRNUM + 22
Public Const MSGERR_NODAO = "Esta función no está soportada en versión DAO." & vbCr & _
                            "Para cambiar de la versión, por favor contácte con su proveedor."
'*** MAKOTO 15/feb/01
Public Const ERR_DESBORDA = 6       'Error de desbordamiento
Public Const MSGERR_DESBORDA As String = _
                "El valor ingresado es demasiado grande, por lo que no se puede completar el cálculo necesario."

'*** ANGEL 19/Mayo/2003
Public Const ERR_SIIPRINT = ERRNUM + 23 'Para poder extraer los mensajes de error de siiprint
'jeaa 06/06/2009
Public Const ERR_REPITECODIGOALT = ERRNUM + 24
Public Const MSGERR_REPITECODIGOALT = "Ya existe el código alterno Por favor utilíce otro código."
Public Const ERR_PRECIOXGRUPO_INCORRECTO = ERRNUM + 25
Public Const MSGERR_PRECIOXGRUPO_INCORRECTO = "El Precio seleccionado es incorrecto segun el tipo de Cliente "
Public Const ERR_TRANS_INVALIDO = ERRNUM + 26
Public Const MSGERR_TRANS_INVALIDO = "La Transacción seleccionada está deshabilitada, actívela en la configuración de la transacción"
Public Const ERR_REPITENNUMESTA = ERRNUM + 27
Public Const MSGERR_REPITENUMESTA = "Ya existe el Número de Establecimiento. Por favor utilíce otro Número."
Public Const ERR_REPITENNUMPUNTO = ERRNUM + 28
Public Const MSGERR_REPITENUMPUNTO = "Ya existe el Número de Punto. Por favor utilíce otro Número."
Public Const ERR_RUCINCORRECTO = ERRNUM + 29
Public Const MSGERR_RUCINCORRECTO = "El número de RUC es incorrecto"
Public Const ERR_REPORTERAGO = ERRNUM + 30
Public Const MSGERR_REPORTERANGO = "Tiene un trámite pendiente de Reporte de Rangos"
Public Const ERR_VALAUTOIMPRESOR = ERRNUM + 31
Public Const MSGERR_VALAUTOIMPRESOR = "Falta validar los documentos con configurados con Autoimpresor"
Public Const ERR_ASIENTOAUTOIMPRESOR = ERRNUM + 32
Public Const MSGERR_ASIENTOAUTOMPRESOR = "Es un asiento generado por un auto impresor, no se lo puede cambiar de estado"

'PARA ROLES
Public Const MSGERR_NOFORMULA = "Fórmula no válida"
Public Const ERR_NOFORMULA = ERRNUM + 11
Public Const ERR_YAEXISTECODIGO = ERRNUM + 10
Public Const MSGERR_YAEXISTECODIGO = "El código ya existe"
Public Const ERR_CANCELMOD = ERRNUM + 13
Public Const MSGERR_CANCELMOD = "El registro está modificado." & vbCr & _
                                    "Desea grabar la modificación?"

Public Const ERR_BASEINCORRECTA = ERRNUM + 23
Public Const MSGERR_BASEINCORRECTA = _
                "No se puede usar el sistema debido a que el servidor no es el que se configuró." & vbCr & _
                "Contácte con su proveedor para continuar trabajándo con el sistema."



'Diego Prod
'------Constante para Tipo de Inventarios
Public Const INV_TIPONORMAL = 0
Public Const INV_TIPOFAMILIA = 1
Public Const INV_TIPORECETA = 2

'----- Constantes para la actualización de ComboBox
Public Const REFRESH_RESPONSABLE = 0
Public Const REFRESH_EMPRESA = 1
Public Const REFRESH_CUENTA = 2
Public Const REFRESH_BANCO = 3
Public Const REFRESH_INVENTARIO = 4
Public Const REFRESH_BODEGA = 5
Public Const REFRESH_GRUPO1 = 6
Public Const REFRESH_GRUPO2 = 7
Public Const REFRESH_GRUPO3 = 8
Public Const REFRESH_GRUPO4 = 9
Public Const REFRESH_GRUPO5 = 10
Public Const REFRESH_VENDEDOR = 11
Public Const REFRESH_PROVCLI = 12
Public Const REFRESH_ZONA = 13
Public Const REFRESH_AFGRUPO = 14
Public Const REFRESH_AF = 15
Public Const REFRESH_TRANS = 16
Public Const REFRESH_TIPODOCBANCO = 17
Public Const REFRESH_FORMACOBROPAGO = 18
Public Const REFRESH_FORMAPAGO = 19
Public Const REFRESH_RECARGO = 20
Public Const REFRESH_OPCION = 21
Public Const REFRESH_CENTROCOSTO = 22
Public Const REFRESH_PCGRUPO1 = 23
Public Const REFRESH_PCGRUPO2 = 24
Public Const REFRESH_PCGRUPO3 = 25
Public Const REFRESH_PCGRUPO4 = 26  'auc
Public Const REFRESH_MONEDA = 27
Public Const REFRESH_RETENCION = 28         '*** MAKOTO 07/feb/01 Agregado
Public Const REFRESH_GNCOMP = 29
Public Const REFRESH_CTLOCAL = 30
Public Const REFRESH_DPCGXIVG = 31
Public Const REFRESH_MOTIVO = 32
Public Const REFRESH_TIPOCOMPRA = 33
Public Const REFRESH_IVUNIDAD = 34
Public Const REFRESH_IVRECARGOICE = 35
Public Const REFRESH_ANEXOCOMPROBANTES = 36 'AUC 27/03/06
Public Const REFRESH_ANEXOSUSTENTOS = 37
Public Const REFRESH_ANEXORETIR = 38
Public Const REFRESH_ANEXOTRANS = 39
Public Const REFRESH_ANEXOTIPODOC = 40 'jeaa 15/05/2007
Public Const REFRESH_COBRADOR = 41          'JEAA 27/05/2007
Public Const REFRESH_GNGASTO = 42          'JEAA 07/01/2008
Public Const REFRESH_ANEXOREGIMEN = 43          'JEAA 26/03/2008
Public Const REFRESH_ANEXODISTRITO = 44          'JEAA 26/03/2008
Public Const REFRESH_ANEXOTIPOEXP = 45          'JEAA 26/03/2008
Public Const REFRESH_DNUMPXIVG = 46          'JEAA 02/09/2008
Public Const REFRESH_AFGRUPO1 = 47          'JEAA 28/10/2008
Public Const REFRESH_AFGRUPO2 = 48          'JEAA 28/10/2008
Public Const REFRESH_AFGRUPO3 = 49          'JEAA 28/10/2008
Public Const REFRESH_AFGRUPO4 = 50          'JEAA 28/10/2008
Public Const REFRESH_AFGRUPO5 = 51          'JEAA 28/10/2008
Public Const REFRESH_AFBODEGA = 52          'JEAA 28/10/2008
Public Const REFRESH_AFINVENTARIO = 53          'JEAA 28/10/2008
Public Const REFRESH_AFTIPOSEGURO = 54          'JEAA 28/10/2008
Public Const REFRESH_GNSUCURSAL = 55          'JEAA 05/02/2009
Public Const REFRESH_GNCGRUPO1 = 56          'JEAA 20/03/2009
Public Const REFRESH_GNCGRUPO2 = 57          'JEAA 20/03/2009
Public Const REFRESH_GNCGRUPO3 = 58          'JEAA 20/03/2009
Public Const REFRESH_GNCGRUPO4 = 59          'JEAA 20/03/2009
Public Const REFRESH_GNVEHICULO = 60          'JEAA 30/03/2009
Public Const REFRESH_GNVGRUPO1 = 56          'JEAA 20/03/2009
Public Const REFRESH_GNVGRUPO2 = 57          'JEAA 20/03/2009
Public Const REFRESH_GNVGRUPO3 = 58          'JEAA 20/03/2009
Public Const REFRESH_GNVGRUPO4 = 59          'JEAA 20/03/2009
Public Const REFRESH_GNDESTINO = 60          'JEAA 12/06/2009
Public Const REFRESH_TIPODOCCOBRO = 61      'JEAA 09/07/2009  'anulado
Public Const REFRESH_IVBANCO = 61      'JEAA 09/07/2009
Public Const REFRESH_IVTARJETA = 62      'JEAA 09/07/2009
Public Const REFRESH_IVESPPRODISO = 63       'JEAA 22/09/2009
Public Const REFRESH_OBRA = 64       'JEAA 01/02/2010
Public Const REFRESH_ZONAS = 65       'JEAA 01/02/2010
Public Const REFRESH_IVDESCUENTO = 66 'JEAA 05/10/2010
Public Const REFRESH_IVPROMOCION = 67 'JEAA 11/10/2010
Public Const REFRESH_PRCUENTA = 68 'jeaa 21/02/2011
Public Const REFRESH_PCPROVINCIA = 69
Public Const REFRESH_PCCANTON = 70
Public Const REFRESH_PCPARROQUIA = 71
Public Const REFRESH_SOLPROVCLI = 72
Public Const REFRESH_PCEEMPLEADO = 73
Public Const REFRESH_PCDIASCREDITO = 74
Public Const REFRESH_ELEMENTOS = 75
Public Const REFRESH_ROLDETALLE = 76
Public Const REFRESH_IR = 77
Public Const REFRESH_IVRECARGOARANCEL = 78
Public Const REFRESH_TIPOROL = 79 '
Public Const REFRESH_IVCOMIXVEN = 80
Public Const REFRESH_GNPROYECTO = 81
Public Const REFRESH_GNCOMPETENCIA = 82
Public Const REFRESH_PCGGASTO = 83
Public Const REFRESH_GNCONTRATO = 84
Public Const REFRESH_DPLAZOPCGXIVG = 85
Public Const REFRESH_IVPROCESO = 86 'AUC 04/04/08
Public Const REFRESH_DETALLEPROCESO = 87 'AUC 04/04/08
Public Const REFRESH_CENTROCOSTOHIJO = 88
Public Const REFRESH_DETALLEPROCESOD = 89 'AUC ecuamueble
Public Const REFRESH_GNESTADOPROD = 90 'AUC ecuamueble
Public Const REFRESH_PROTIEMPOS = 91 'AUC MADERAMICA
Public Const REFRESH_TURNOS = 92 'AUC RELOJ
Public Const REFRESH_ANEXOFORMAPAGO = 93 'jeaa 15/05/2007
Public Const REFRESH_JORNADA = 94
Public Const REFRESH_FICHA = 95
Public Const REFRESH_FICHADETALLE = 96
Public Const REFRESH_ANEXOPAIS = 97
Public Const REFRESH_AFUBICACION = 98
Public Const REFRESH_RMOTIVO = 99  'AUC motivos de permisos para reloj
Public Const REFRESH_RPERMISO = 100  'AUC Permisos reloj
Public Const REFRESH_RFERIADO = 101  'AUC Feriados reloj
Public Const REFRESH_DETALLEPROCESODM = 102
Public Const REFRESH_DETALLEPROCESOORDEN = 103
Public Const REFRESH_IVSERIE = 104
Public Const REFRESH_PCGESTION = 105
Public Const REFRESH_LISTAGESTION = 106
Public Const REFRESH_GNTRANSPORTE = 107
Public Const REFRESH_IVMOTIVO = 108
Public Const REFRESH_GRUPO6 = 109
Public Const REFRESH_PRECIOXCLIXITEM = 110
Public Const REFRESH_FCVTABLACOMISION = 111
Public Const REFRESH_IVPLAN = 112
Public Const REFRESH_IVPLANITEM = 113
Public Const REFRESH_IVPLANCALENDARIO = 114
Public Const REFRESH_ANEXORETIVA = 115 'AUC 10/09/2015
Public Const REFRESH_IVGCOMIXVEN = 116
Public Const REFRESH_IVEQUIPO = 117
Public Const REFRESH_PCPAIS = 118
Public Const REFRESH_GNAGENCIACURIER = 119
Public Const REFRESH_PCCALLE = 120
Public Const REFRESH_ANEXOICE = 121
Public Const REFRESH_IVRECETA = 122
Public Const REFRESH_GNSEGURO = 123
Public Const REFRESH_FCVCALENDARIO = 124
Public Const REFRESH_GNRUTA = 125

Public Const IMPRESION_NOIMPRIME = 0
Public Const IMPRESION_YAIMPRIMIO = 1

Public Const REGISTRO = "HKEY_CLASSES_ROOT\CLSID\"
Public Const ProductCode = "{B5AD5C02-B948-449E-A8AC-3B2945ACF7IA2005}"  'Product Code del Instalador
Public Regedit As Object  ' para los registros especiales de

Public Const ESTADO_ALQUILADO = 0 'AUC 12/05/06
Public Const ESTADO_RESERVADO = 1

'AUC 31/05/07
Public Const ModuloFac = "SiiFactura"
Public Const ModuloSii = "Sii4A"
Public Const ModuloTools = "SiiTools"
Public Const ModuloPrecio = "SiiPrecio"
Public Const ModuloConfig = "SiiConfig"
Public Const ModuloReport = "SiiReporte"
Public Const ModuloPuntoEqui = "PuntoEqui"
Public Const ModuloJefeVentas = "JefeVentas"  'jeaa 21/07/2009
Public Const ModuloPagaRol = "PagaRol"
'AUC24/10/07
Public Const ESTADO_NOFACTURADO = 0
Public Const ESTADO_FACTURADO = 1

Public Const ESTADO_NOCOMPRAS = 0
Public Const ESTADO_COMPRAS = 1

Public Const ERR_NOCLAVE = ERRNUM + 24
Public Const MSGERR_NOCLAVE = "Clave de usuario no válida"

Public Const STROPR = "+,-,/,*,HOY(),(,),[,],>,<,=,#,AÑO,SI,NO,Y,O,;"
'AUC 01/04/2008 variables de estados para procesos
Public Const ESTADO_NOINICIADO = 0
Public Const ESTADO_INICIADO = 1
Public Const ESTADO_DETENIDO = 2
Public Const ESTADO_TERMINADO = 3

'AUC 12/06/06
Global Const for_ini = "Posicionformulario.ini"
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Const HWND_TOPMOST = -1
Public Const SWP_NOACTIVATE = &H10
Public Const HWND_NOTOPMOST = -2
Public Const SWP_SHOWWINDOW = &H40
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

