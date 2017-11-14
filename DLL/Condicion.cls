VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Condicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum ETIPO
    bsqCT = 1           '1
    bsqCTKardex      '2
    bsqCTBalC
    bsqCTBalG
    bsqCTEstado
    bsqCTCentroB
    bsqCTCentroR
    bsqTSKardex
    bsqTSFlujo
    bsqTSSaldo
    bsqIVTrans
    bsqIVKardex
    bsqGNTrans
    bsqCentroCosto
    bsqPVKardex
    bsqPVPago
    bsqPCProyPago
    bsqCLKardex
    bsqCLCobro
    bsqPCProyCobro
    bsqPVSaldo
    bsqCLSALDO
    bsqIVCCosto
    bsqGNLog            'GNLogAccion (Para Auditoria)   *** MAKOTO 12/jul/00
    bsqPVKardexPorTrans         'Kardex de Proveedor por trans  '*** MAKOTO 24/sep/00
    bsqCLKardexPorTrans         'Kardex de Cliente por trans  '*** MAKOTO 24/sep/00
    bsqIVComparativo            'Comparativo Presupuesto - Consumo  '*** MAKOTO 07/nov/00
    bsqPCDocPorCobrarPagar      'Documentos por cobrar/pagar        '*** MAKOTO 10/dic/00
    bsqIVComparativoSuma        'Comparativo Presupuesto - Consumo Totalizado   '*** MAKOTO 05/feb/01
    bsqTSKardexRet              'Kardex de retención
    bsqCTBalPM              '*** MAKOTO 12/mar/01 Agregado
    bsqRutasAuditoria       'Agregado Oliver 6/feb/2002 para consultas de Rutas de Auditoria
    bsqAccesoModulo         'Diego  14-04-2002
    bsqComprasAnexos        ' Esteban compras para anexos 27-8-03
    bsqCliAnexos                ' Esteban listado de clientes para anexos 27-8-03
    bsqProvAnexos                ' Esteban listado de Proveedores para anexos 28-8-03
    bsqRetIrAnexos              ' Esteban listado de Retenciones IR para anexos 28-8-03
    bsqCPAnexos                 ' Esteban listado de Compras para anexos 28-8-03
    bsqNCAnexos                 ' Esteban listado de NOtas de Credito y Notas de DEbito para anexos 28-8-03
    bsqVentasAnexos                 ' 40 Esteban listado de Ventas para anexos 28-8-03
    bsqIVComparativoFecha        '41 Comparativo Presupuesto - Consumo Totalizado   por Fecha'*** JEAA 06/04/04
    bsqIVKardexCC                   '42 Kardex de Centos de Costo '*** jeaa 05/07/05
    bsqCTBalxMesN                   '43 BValance General x mes con niveles  '*** jeaa 26/11/05
    bsqIVMovCC                   '44 Kardex de Centos de Costo '*** jeaa 05/07/05
    bsqIVPendientesCC                   '45 Kardex de Centos de Costo '*** jeaa 05/07/05
    bsqCompAnuladoAnexo              '46
    bsqPVKardexPorTransRes         '47 Kardex de Proveedor por trans  resumido'*** jeaa 29/04/2006
    bsqCLKardexPorTransRes         '48 Kardex de Cliente por trans  resumido '*** jeaa 29/04/2006
    bsqretVentasAnexos                 ' 49 Esteban listado de Retenciones Ventas para anexos 28-8-03
    bsqVentasRetAnexos                 ' 50 Esteban listado de Retenciones Ventas para anexos 28-8-03
    bsqCompraAnexos                    ' 51 Esteban listado de Compras para anexos 03/03/2008
    bsqCompraRetAnexos                 ' 52 Esteban listado de Retenciones Compras para anexos 03/03/2008
    bsqGNTransAnexos                ' 53 Consulta General con Anexos
    bsqLiqCompraAnexos                    ' 54 Esteban listado de Compras para anexos 03/03/2008
    bsqGNTransAereo                ' 54 Consulta General Aereo
    bsqAFKardex                     ' 55 Kardex de Activos Fijos
    bsqCompraACTAnexos                    ' 56 Esteban listado de Compras para anexos 03/03/2008
    bsqCompraACTRetAnexos                 ' 57 Esteban listado de Retenciones Compras para anexos 03/03/2008
    bsqIVKardexItemTipo             '58 AUC listado de kardex item de tipo 4 e item de bandventa = true
    bsqItems             '60 listado de items
    bsqGNTransTiket  '61 consulta tiket ISOLLANTA
    bsqVentaAnexo                 ' 62 Esteban listado de Ventas para anexos anterior
    bsqVentasAnexosN             ' 63 Esteban listado de Ventas para anexos 2010
    bsqCTEstadoComparativo   ' 64 Esteban Perdidas y ganacias Comparativo
    bsqIVKardexCli          '65 AUC encofrados Mostrar el kardex por cliente
    bsqGNTransxTiket  '66 consulta tiket ISOLLANTA
    bsqPR              '67
    bsqPRKardex      '68
    bsqPRBalC          '69
    bsqPRBalG          '70
    bsqprEstado       '71
    bsqAFKardexC    '72                  ' Kardex de Activos Fijos Custodios
    bsqIVKardexP     '73
    bsqEMPKardex        '74
    bsqEMPKardexPorTrans '75
    bsqEmpSALDO '76  AUC saldo empleados
    bsqGNTransIV  '77 jeaa
    bsqAFTrans  '78
    bsqCTFlujoEfectivo '79
    bsqCompraAnexoscg '80
    bsqGNCall '81 AUC este ya  no se utiliza
    bsqGNKardexDoc
    bsqIVKardexSerie
    bsqCLKardexBQD
    bsqPVKardexBQD
    bsqCLKardexPorTransBQD
    bsqPVKardexPorTransBQD
    bsqGNTransBQD
    bsqPCDocPorCobrarPagarBQD
    bsqGARKardexPorTrans
    bsqIVKardexPACA '91 para kardex de guias curier
End Enum

Private Type TPROPS
    Tipo As ETIPO
    CodMoneda As String
    fecha1 As Date
    fecha2 As Date
    CodTrans As String
    NumTrans1 As Long
    NumTrans2 As Long
    Estado As Integer           'Cada Bits desde la DERECHA representa 0(Pendiente),1(Aprobado),2(Despachado),3(Anulado)
    CodCuenta1 As String
    CodCuenta2 As String
    Nivel As Integer
    CodCentro1 As String
    CodCentro2 As String
    CodBanco1 As String
    CodBanco2 As String
    NumDias1 As Integer
    NumDias2 As Integer
    NumDias3 As Integer
    CodItem1 As String
    CodItem2 As String
    CodPC1 As String
    CodPC2 As String
    CodBodega1 As String
    CodBodega2 As String
    FechaCorte As Date
    NumGrupo As Byte        '# de grupo por el que tiene que agrupar  '*** MAKOTO 03/sep/00
    CodGrupo As String      'Cod. grupo                      '*** MAKOTO 10/dic/00
    TopN As Integer         '# de registros a obtener.      '*** MAKOTO 03/oct/00 Agregado
    IncluirCero As Boolean  'True:Incluir saldo cero.         '*** MAKOTO 03/dic/00 Agregado
    CodForma As String      'Forma de cobro/pago                    '*** MAKOTO 10/dic/00
    PorCobrar As Boolean    'True: Por cobrar, False: Por pagar     '*** MAKOTO 10/dic/00
    Cliente As Boolean      'True: Incluir clientes                 '*** MAKOTO 10/dic/00
    Proveedor As Boolean    'True: Incluir proveedores              '*** MAKOTO 10/dic/00
    Empleado As Boolean     'AUC incluye todos los empleados

    CodRetencion1 As String
    CodRetencion2 As String
    NumPCGrupo As Byte      '# de grupo de Prov/Cli por el que tiene que agrupar  '*** MAKOTO 13/mar/01
    CodPCGrupo As String    'Cod. grupo de Prov/Cli                '*** MAKOTO 13/mar/01
    numMoneda As Byte
    Hora1 As Date           'Agregado Oliver para busquedas de Rutas de Auditoria
    Hora2 As Date           '--  06/feb/2002  Agregado Oliver para busquedas de Rutas de Auditoria
    Usuario1 As String      ' busqueda de Rutas de Auditoria
    Usuario2 As String      ' monitorear los transacciones creadas y modificadas por el usuario
    
    Grupo1 As String        ' -jeaA 27-08-03
    Grupo2 As String
    Bienes As String        ' -jeaA 27-08-03
    Servicios As String
    Sucursal As String        ' -jeaA 16-01-04  para filtrar lasctas por sucursal
    TipoSaldo As Integer    '- jeaa 02-04-04 tipo de saldo 0 todo, 1 anticipos, 2 por cobrar
    
    CrediDispo As String  'AUC 11/10/2005
    Crediselec As String 'AUC 11/10/2005
    Estado1 As Integer 'AUC 26/10/07
    Estado2 As Integer 'AUC 26/10/07
    'jeaa 21/07/2008  para busqueda para anexos
    NumEsta As String
    NumPunto As String
    Numsec As String
    NumAuto As String
    CodComp As String
    CodSust As String
    NombreItem As String 'AUC 13/08/06 para filtrar por nombre del item
    Telefono As String 'AUC 13/08/06 para filtrar por telefono
    Direccion As String 'AUC 23/11/06 para filtrar direccion
    nomPcGrupo As String 'AUC 23/11/06 para filtrar grupo cliente
    NomCliente As String 'AUC 23/11/06
    tipoItem1 As Boolean  'AUC 18/12/06
    tipoItem2 As Boolean  'AUC 18/12/06
    TasaMensual As Currency 'jeaa 26/03/2009
    CodVehiculo As String   'jeaa 24/04/2009
    Codelemento As String
    Depart1 As String
    Depart2 As String
    Codciudad1 As String
    Codciudad2 As String
    CodTipoRol As String
    CodForma1 As String
    Precio As Currency
    BandAnticipo As Boolean
    BandTodo As Boolean
    PrecioV As Currency
    CodReceta As String
    Receta As String
End Type

Private mProps As TPROPS

Public Property Let Tipo(ByVal value As ETIPO)
    mProps.Tipo = value
End Property

Public Property Get Tipo() As ETIPO
    Tipo = mProps.Tipo
End Property
    
Public Property Let CodMoneda(ByVal value As String)
    If Len(mProps.CodMoneda) = 0 Then mProps.CodMoneda = MONEDA_PRE
    mProps.CodMoneda = value
End Property

Public Property Get CodMoneda() As String
    If Len(mProps.CodMoneda) = 0 Then mProps.CodMoneda = MONEDA_PRE
    CodMoneda = Trim$(mProps.CodMoneda)
End Property
    
Public Property Get numMoneda() As Integer
    If Len(mProps.CodMoneda) = 0 Then mProps.CodMoneda = MONEDA_PRE
    numMoneda = gobjMain.EmpresaActual.GNOpcion.IndiceMoneda(mProps.CodMoneda)
End Property

Public Property Let fecha1(ByVal value As Date)
    mProps.fecha1 = value
End Property

Public Property Get fecha1() As Date
    fecha1 = mProps.fecha1
End Property
    
Public Property Let fecha2(ByVal value As Date)
    mProps.fecha2 = value
End Property

Public Property Get fecha2() As Date
    fecha2 = mProps.fecha2
End Property


Public Property Let Hora1(ByVal value As Date)
    mProps.Hora1 = value
End Property

Public Property Get Hora1() As Date
    Hora1 = mProps.Hora1
End Property
    
Public Property Let Hora2(ByVal value As Date)
    mProps.Hora2 = value
End Property

Public Property Get Hora2() As Date
    Hora2 = mProps.Hora2
End Property

    
Public Property Let FechaCorte(ByVal value As Date)
    mProps.FechaCorte = value
End Property

Public Property Get FechaCorte() As Date
    FechaCorte = mProps.FechaCorte
End Property
    
    
Public Property Let CodTrans(ByVal value As String)
    mProps.CodTrans = value
End Property

Public Property Get CodTrans() As String
    CodTrans = Trim$(mProps.CodTrans)
End Property


Public Property Let NumTrans1(ByVal value As Long)
    mProps.NumTrans1 = value
End Property

Public Property Get NumTrans1() As Long
    NumTrans1 = mProps.NumTrans1
End Property

Public Property Let NumTrans2(ByVal value As Long)
    mProps.NumTrans2 = value
End Property

Public Property Get NumTrans2() As Long
    NumTrans2 = mProps.NumTrans2
End Property

Public Property Let Estado(ByVal value As Integer)
    mProps.Estado = value
End Property

Public Property Get Estado() As Integer
    Estado = mProps.Estado
End Property

Public Property Let EstadoBool(ByVal index As Integer, ByVal value As Boolean)
    Dim v As Integer
    v = &H2 ^ index
    If value = True Then
        mProps.Estado = mProps.Estado Or v
    Else
        mProps.Estado = mProps.Estado And (Not v)
    End If
End Property

Public Property Get EstadoBool(ByVal index As Integer) As Boolean
    Dim v As Integer
    v = &H2 ^ index
    EstadoBool = CBool(mProps.Estado And v)
End Property

    
Public Property Let CodCuenta1(ByVal value As String)
    mProps.CodCuenta1 = value
End Property

Public Property Get CodCuenta1() As String
    CodCuenta1 = Trim$(mProps.CodCuenta1)
End Property

Public Property Let CodCuenta2(ByVal value As String)
    mProps.CodCuenta2 = value
End Property

Public Property Get CodCuenta2() As String
    CodCuenta2 = Trim$(mProps.CodCuenta2)
End Property
    
Public Property Let Usuario1(ByVal value As String)
    mProps.Usuario1 = value
End Property

Public Property Get Usuario1() As String
    Usuario1 = Trim$(mProps.Usuario1)
End Property
    
Public Property Let Usuario2(ByVal value As String)
    mProps.Usuario2 = value
End Property

Public Property Get Usuario2() As String
    Usuario2 = Trim$(mProps.Usuario2)
End Property
    
Public Property Let Nivel(ByVal value As Integer)
    mProps.Nivel = value
End Property

Public Property Get Nivel() As Integer
    Nivel = mProps.Nivel
End Property
    
Public Property Let CodCentro1(ByVal value As String)
    mProps.CodCentro1 = value
End Property

Public Property Get CodCentro1() As String
    CodCentro1 = Trim$(mProps.CodCentro1)
End Property
    
Public Property Let CodCentro2(ByVal value As String)
    mProps.CodCentro2 = value
End Property

Public Property Get CodCentro2() As String
    CodCentro2 = Trim$(mProps.CodCentro2)
End Property
    
Public Property Let CodBanco1(ByVal value As String)
    mProps.CodBanco1 = value
End Property

Public Property Get CodBanco1() As String
    CodBanco1 = Trim$(mProps.CodBanco1)
End Property
    
Public Property Let CodBanco2(ByVal value As String)
    mProps.CodBanco2 = value
End Property

Public Property Get CodBanco2() As String
    CodBanco2 = Trim$(mProps.CodBanco2)
End Property
    
Public Property Let CodRetencion1(ByVal value As String)
    mProps.CodRetencion1 = value
End Property

Public Property Get CodRetencion1() As String
    CodRetencion1 = Trim$(mProps.CodRetencion1)
End Property
    
Public Property Let CodRetencion2(ByVal value As String)
    mProps.CodRetencion2 = value
End Property

Public Property Get CodRetencion2() As String
    CodRetencion2 = Trim$(mProps.CodRetencion2)
End Property
    
Public Property Let NumDias1(ByVal value As Integer)
    mProps.NumDias1 = value
End Property

Public Property Get NumDias1() As Integer
    NumDias1 = mProps.NumDias1
End Property
    
Public Property Let NumDias2(ByVal value As Integer)
    mProps.NumDias2 = value
End Property

Public Property Get NumDias2() As Integer
    NumDias2 = mProps.NumDias2
End Property
    
Public Property Let NumDias3(ByVal value As Integer)
    mProps.NumDias3 = value
End Property

Public Property Get NumDias3() As Integer
    NumDias3 = mProps.NumDias3
End Property
    
Public Property Let CodItem1(ByVal value As String)
    mProps.CodItem1 = value
End Property

Public Property Get CodItem1() As String
    CodItem1 = Trim$(mProps.CodItem1)
End Property
    
Public Property Let CodItem2(ByVal value As String)
    mProps.CodItem2 = value
End Property

Public Property Get CodItem2() As String
    CodItem2 = Trim$(mProps.CodItem2)
End Property
    
Public Property Let CodPC1(ByVal value As String)
    mProps.CodPC1 = value
End Property

Public Property Get CodPC1() As String
    CodPC1 = Trim$(mProps.CodPC1)
End Property
    
Public Property Let CodPC2(ByVal value As String)
    mProps.CodPC2 = value
End Property

Public Property Get CodPC2() As String
    CodPC2 = Trim$(mProps.CodPC2)
End Property

Public Property Let CodBodega1(ByVal value As String)
    mProps.CodBodega1 = value
End Property

Public Property Get CodBodega1() As String
    CodBodega1 = Trim$(mProps.CodBodega1)
End Property
    
Public Property Let CodBodega2(ByVal value As String)
    mProps.CodBodega2 = value
End Property

Public Property Get CodBodega2() As String
    CodBodega2 = Trim$(mProps.CodBodega2)
End Property

Public Property Let NumGrupo(ByVal value As Byte)
    mProps.NumGrupo = value
End Property

Public Property Get NumGrupo() As Byte
    NumGrupo = mProps.NumGrupo
End Property

'*** MAKOTO 10/dic/00
Public Property Let CodGrupo(ByVal value As String)
    mProps.CodGrupo = value
End Property

Public Property Get CodGrupo() As String
    CodGrupo = mProps.CodGrupo
End Property

'*** MAKOTO 13/mar/01 Agregado
Public Property Let NumPCGrupo(ByVal value As Byte)
    mProps.NumPCGrupo = value
End Property

'*** MAKOTO 13/mar/01 Agregado
Public Property Get NumPCGrupo() As Byte
    NumPCGrupo = mProps.NumPCGrupo
End Property

'*** MAKOTO 13/mar/01 Agregado
Public Property Let CodPCGrupo(ByVal value As String)
    mProps.CodPCGrupo = value
End Property

'*** MAKOTO 13/mar/01 Agregado
Public Property Get CodPCGrupo() As String
    CodPCGrupo = mProps.CodPCGrupo
End Property

'*** MAKOTO 10/dic/00
Public Property Let CodForma(ByVal value As String)
    mProps.CodForma = value
End Property

Public Property Get CodForma() As String
    CodForma = mProps.CodForma
End Property

'*** MAKOTO 10/dic/00
Public Property Let PorCobrar(ByVal value As Boolean)
    mProps.PorCobrar = value
End Property

Public Property Get PorCobrar() As Boolean
    PorCobrar = mProps.PorCobrar
End Property

'*** MAKOTO 10/dic/00
Public Property Let Cliente(ByVal value As Boolean)
    mProps.Cliente = value
End Property

Public Property Get Cliente() As Boolean
    Cliente = mProps.Cliente
End Property

'*** MAKOTO 10/dic/00
Public Property Let Proveedor(ByVal value As Boolean)
    mProps.Proveedor = value
End Property

Public Property Get Proveedor() As Boolean
    Proveedor = mProps.Proveedor
End Property

'*** MAKOTO 03/oct/00 Agregado
Public Property Let TopN(ByVal value As Integer)
    mProps.TopN = value
End Property

Public Property Get TopN() As Integer
    TopN = mProps.TopN
End Property

'*** MAKOTO 03/dic/00 Agregado
Public Property Let IncluirCero(ByVal value As Boolean)
    mProps.IncluirCero = value
End Property

Public Property Get IncluirCero() As Boolean
    IncluirCero = mProps.IncluirCero
End Property

'*** JEAA 27-08-03
Public Property Get Grupo2() As String
    Grupo2 = mProps.Grupo2
End Property

Public Property Let Grupo2(ByVal value As String)
    mProps.Grupo2 = value
End Property

Public Property Get Grupo1() As String
    Grupo1 = mProps.Grupo1
End Property

Public Property Let Grupo1(ByVal value As String)
    mProps.Grupo1 = value
End Property

Public Property Get Bienes() As String
    Bienes = mProps.Bienes
End Property

Public Property Let Bienes(ByVal value As String)
    mProps.Bienes = value
End Property
 
 Public Property Get Servicios() As String
    Servicios = mProps.Servicios
End Property

Public Property Let Servicios(ByVal value As String)
    mProps.Servicios = value
End Property


 Public Property Get Sucursal() As String
    Sucursal = mProps.Sucursal
End Property

Public Property Let Sucursal(ByVal value As String)
    mProps.Sucursal = value
End Property


Public Property Let TipoSaldo(ByVal value As Integer)
    mProps.TipoSaldo = value
End Property

Public Property Get TipoSaldo() As Integer
    TipoSaldo = mProps.TipoSaldo
End Property




Private Sub Class_Initialize()
    With mProps
        '*** MAKOTO 11/jul/00   En vez de un mes antes, cambio al inicio del año
        '.fecha1 = DateSerial(Year(Date), Month(Date) - 1, Day(Date))    'Un mes antes
'        .fecha1 = DateSerial(Year(Date), 1, 1) 'Primer día del año actual
        .fecha1 = Date
        .fecha2 = Date
        .FechaCorte = Date
        .Nivel = 3
'        .CodMoneda = MONEDA_PRE        '*** MAKOTO 08/sep/00 Modificado USD predeterminado
        .CodMoneda = "USD"              '***
        .NumDias1 = 10
        .NumDias2 = 20
        .NumDias3 = 30
        .Estado = &HFFF7        'Incluye todo excepto 3(Anulados)
        .NumGrupo = 1
        .NumPCGrupo = 1                 '*** MAKOTO 13/mar/01
        .TopN = 0
        
        .PorCobrar = False
       ' .Cliente = True                 '*** MAKOTO 10/dic/00
       '.IncluirCero = True      'AUC cambiado par que empiece en todo
       .BandTodo = True 'AUC cambiado da problemas incluircero
        .Proveedor = True               '*** MAKOTO 10/dic/00
    End With
End Sub

'Agregado AUC 11/10/2005 DESDE AQUI
Public Property Let CrediDispo(ByVal value As String)
    mProps.CrediDispo = value
End Property

Public Property Get CrediDispo() As String
    CrediDispo = Trim$(mProps.CrediDispo)
End Property
Public Property Let Crediselec(ByVal value As String)
    mProps.Crediselec = value
End Property

Public Property Get Crediselec() As String
    Crediselec = Trim$(mProps.Crediselec)
End Property
'AUC 26/10/07
Public Property Let Estado1(ByVal value As Integer)
    mProps.Estado1 = value
End Property

Public Property Get Estado1() As Integer
    Estado1 = mProps.Estado1
End Property

Public Property Let Estado2(ByVal value As Integer)
    mProps.Estado2 = value
End Property

Public Property Get Estado2() As Integer
    Estado2 = mProps.Estado2
End Property

Public Property Let Estado1Bool(ByVal index As Integer, ByVal value As Boolean)
    Dim v As Integer
    v = &H2 ^ index
    If value = True Then
        mProps.Estado1 = mProps.Estado1 Or v
    Else
        mProps.Estado1 = mProps.Estado1 And (Not v)
    End If
End Property

Public Property Get Estado1Bool(ByVal index As Integer) As Boolean
    Dim v As Integer
    v = &H2 ^ index
    Estado1Bool = CBool(mProps.Estado1 And v)
End Property

Public Property Let Estado2Bool(ByVal index As Integer, ByVal value As Boolean)
    Dim v As Integer
    v = &H2 ^ index
    If value = True Then
        mProps.Estado2 = mProps.Estado2 Or v
    Else
        mProps.Estado2 = mProps.Estado2 And (Not v)
    End If
End Property

Public Property Get Estado2Bool(ByVal index As Integer) As Boolean
    Dim v As Integer
    v = &H2 ^ index
    Estado2Bool = CBool(mProps.Estado2 And v)
End Property

Public Property Get NumEsta() As String
    NumEsta = mProps.NumEsta
End Property

Public Property Let NumEsta(ByVal value As String)
    mProps.NumEsta = value
End Property

Public Property Get NumPunto() As String
    NumPunto = mProps.NumPunto
End Property

Public Property Let NumPunto(ByVal value As String)
    mProps.NumPunto = value
End Property

Public Property Get Numsec() As String
    Numsec = mProps.Numsec
End Property

Public Property Let Numsec(ByVal value As String)
    mProps.Numsec = value
End Property

Public Property Get NumAuto() As String
    NumAuto = mProps.NumAuto
End Property

Public Property Let NumAuto(ByVal value As String)
    mProps.NumAuto = value
End Property

Public Property Get CodComp() As String
    CodComp = mProps.CodComp
End Property

Public Property Let CodComp(ByVal value As String)
    mProps.CodComp = value
End Property

Public Property Get CodSust() As String
    CodSust = mProps.CodSust
End Property

Public Property Let CodSust(ByVal value As String)
    mProps.CodSust = value
End Property

'AUC 13/09/06
Public Property Let NombreItem(ByVal value As String)
    mProps.NombreItem = value
End Property
Public Property Get NombreItem() As String
    NombreItem = Trim$(mProps.NombreItem)
End Property
Public Property Let Telefono(ByVal value As String)
    mProps.Telefono = value
End Property
Public Property Get Telefono() As String
    Telefono = Trim$(mProps.Telefono)
End Property
'AUC 23/11/06
Public Property Let Direccion(ByVal value As String)
    mProps.Direccion = value
End Property
Public Property Get Direccion() As String
    Direccion = Trim$(mProps.Direccion)
End Property
Public Property Let nomPcGrupo(ByVal value As String)
    mProps.nomPcGrupo = value
End Property
Public Property Get nomPcGrupo() As String
    nomPcGrupo = Trim$(mProps.nomPcGrupo)
End Property
Public Property Let NomCliente(ByVal value As String)
    mProps.NomCliente = value
End Property
Public Property Get NomCliente() As String
    NomCliente = Trim$(mProps.NomCliente)
End Property
'AUC 18/12/06
Public Property Let tipoItem1(ByVal value As Boolean)
    mProps.tipoItem1 = value
End Property
Public Property Get tipoItem1() As Boolean
    tipoItem1 = mProps.tipoItem1
End Property

Public Property Let tipoItem2(ByVal value As Boolean)
    mProps.tipoItem2 = value
End Property
Public Property Get tipoItem2() As Boolean
    tipoItem2 = mProps.tipoItem2
End Property

'jeaa  26/03/2009
Public Property Let TasaMensual(ByVal value As Currency)
    mProps.TasaMensual = value
End Property
Public Property Get TasaMensual() As Currency
    TasaMensual = mProps.TasaMensual
End Property

'jeaa 24/04/2009
Public Property Get CodVehiculo() As String
    CodVehiculo = mProps.CodVehiculo
End Property

Public Property Let CodVehiculo(ByVal value As String)
    mProps.CodVehiculo = value
End Property
Public Property Let Empleado(ByVal value As Boolean)
    mProps.Empleado = value
End Property
Public Property Get Empleado() As Boolean
    Empleado = mProps.Empleado
End Property
Public Property Let Codelemento(ByVal value As String)
    mProps.Codelemento = value
End Property
Public Property Get Codelemento() As String
    Codelemento = mProps.Codelemento
End Property
Public Property Let Depart1(ByVal value As String)
    mProps.Depart1 = value
End Property
Public Property Get Depart1() As String
    Depart1 = mProps.Depart1
End Property
Public Property Let Depart2(ByVal value As String)
    mProps.Depart2 = value
End Property
Public Property Get Depart2() As String
    Depart2 = mProps.Depart2
End Property
Public Property Let Codciudad1(ByVal value As String)
    mProps.Codciudad1 = value
End Property
Public Property Get Codciudad1() As String
    Codciudad1 = mProps.Codciudad1
End Property
Public Property Let Codciudad2(ByVal value As String)
    mProps.Codciudad2 = value
End Property
Public Property Get Codciudad2() As String
    Codciudad2 = mProps.Codciudad2
End Property
'AUC 12/12/2011
Public Property Let CodTipoRol(ByVal value As String)
    mProps.CodTipoRol = value
End Property
Public Property Get CodTipoRol() As String
    CodTipoRol = mProps.CodTipoRol
End Property
'AUC Para cargar dos formas de cobro ROCAR 25/01/2011
Public Property Let CodForma1(ByVal value As String)
    mProps.CodForma1 = value
End Property
Public Property Get CodForma1() As String
    CodForma1 = mProps.CodForma1
End Property
Public Property Let Precio(ByVal value As Currency)
    mProps.Precio = value
End Property
Public Property Get Precio() As Currency
    Precio = mProps.Precio
End Property

Public Property Let BandAnticipo(ByVal value As Boolean)
    mProps.BandAnticipo = value
End Property
Public Property Get BandAnticipo() As Boolean
    BandAnticipo = mProps.BandAnticipo
End Property
Public Property Let BandTodo(ByVal value As Boolean)
    mProps.BandTodo = value
End Property

Public Property Get BandTodo() As Boolean
    BandTodo = mProps.BandTodo
End Property

Public Property Let PrecioV(ByVal value As Currency) 'auc PRECIO VISITA
    mProps.PrecioV = value
End Property

Public Property Get PrecioV() As Currency
    PrecioV = mProps.PrecioV
End Property

Public Property Let CodReceta(ByVal value As String) 'auc PRECIO VISITA
    mProps.CodReceta = value
End Property

Public Property Get CodReceta() As String
    CodReceta = mProps.CodReceta
End Property

Public Property Let Receta(ByVal value As String) 'auc PRECIO VISITA
    mProps.Receta = value
End Property

Public Property Get Receta() As String
    Receta = mProps.Receta
End Property

