VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IRKardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Clase para las Retenciones de Dependencia Directa"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"gnComprobante"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type T_PROPS      'modificado por Oliver  a privado
    id As Long
    IDComp As Long
    TipoSalario     As String
    Ingresos        As Currency
    AportePersonal  As Currency
    BaseImponible   As Currency 'Base imponible total anual
    ValorRetenido   As Currency
    TipoRetencion  As String
    NumRet As Integer
    NumMeses As Currency
    'auc 03/05/07
    OtraRemuneracion As Currency
    XIII As Currency
    XIV As Currency
    Utilidad As Currency
    RebajaEspecial As Currency 'Discapacidad
    IR As Currency 'impuesto a la renta asumido por el empleador
    RebajaEspecial3E As Currency
    ImpRentEmpl As Currency
    
    IdPorcRetInt  As Long           'agregados por Diego para manejar porcentajes
    TipoRetInterno  As String         'oliver modifico por causas de recuperacion y grabacion
    PorcRetInt As Double
    
    IdEmpleado As Long
    CodEmpleado As String
    Orden As Integer
    
    fecha As Date           'Agregado Alex: para recuperar el tipoRet de acuerdo a la fecha
    codTipoComprobante  As String  ' PARA RENDIMIENTOS fINANCIEROS
    Deposito  As Currency
    'AUC
    'formulario 107
    FR As Currency
    Deshaucio As Currency
    GPVivienda As Currency
    GPSalud As Currency
    GPEducacion As Currency
    GPAlimentacion As Currency
    GPVestimenta As Currency
    IngGravOtrosEmp As Currency 'Impuestos gravados generados por otros empleadores
    GPOtrosEmp  As Currency 'Gastos personales otros empleadores
    OtrasReb As Currency 'Otras rebajas consideras por otros empleadores
    IRCausado As Currency
    ValorIR As Currency 'Valor IR empleado actual
    ValorIRAnt  As Currency 'Valor del impuesto a ala renta por empleadores anteriores en el periodo actual
    'CONTRIBUCION SOLIDARIA
    APOSOL1 As Currency
    APOSOL2 As Currency
    APOSOL3 As Currency
    APOSOL4 As Currency
    APOSOL5 As Currency
    APOSOL6 As Currency
    APOSOL7 As Currency
    APOSOL8 As Currency
    APOSOL9 As Currency
    APOSOL10 As Currency
    APOSOL11 As Currency
    APOSOL12 As Currency
    APOSOL13 As Currency
    APOSOL14 As Currency
    APOSOL15 As Currency
    APOSOL16 As Currency
    APOSOL17 As Currency
    APOSOL18 As Currency
    APOSOL19 As Currency
    APOSOL20 As Currency
    APOSOL21 As Currency
    GNComprobante As GNComprobante
End Type
'variables locales para almacenar los valores de las propiedades
Private mProps As T_PROPS

Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

'*********************************
' Metodos
'*********************************
Public Sub Recuperar(rs As Recordset)
    With rs
        If Not .EOF Then
            mProps.id = !id
            If Not IsNull(!TipoSalario) Then mProps.TipoSalario = !TipoSalario
            If Not IsNull(!Ingresos) Then mProps.Ingresos = !Ingresos
            If Not IsNull(!AportePersonal) Then mProps.AportePersonal = !AportePersonal
            If Not IsNull(!BaseImponible) Then mProps.BaseImponible = !BaseImponible
            If Not IsNull(!ValorRetenido) Then mProps.ValorRetenido = !ValorRetenido
            If Not IsNull(!TipoRetencion) Then mProps.TipoRetencion = !TipoRetencion
            If Not IsNull(!IdPorcRetInt) Then mProps.IdPorcRetInt = !IdPorcRetInt
            
            If Not IsNull(!IdPorcRetInt) Then mProps.IdPorcRetInt = !IdPorcRetInt
           ' If Not IsNull(!CodPorcentaje) Then mprops.TipoRetInterno = !CodPorcentaje
        '    If Not IsNull(!Porcentaje) Then mprops.PorcRetInt = !Porcentaje   ' Modificado Oliver no recupera este valor
            
            If Not IsNull(!IdEmpleado) Then mProps.IdEmpleado = !IdEmpleado
            If Not IsNull(!CodEmpleado) Then mProps.CodEmpleado = !CodEmpleado
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            '*Agregado Alex Feb/2003
            If Not IsNull(!NumRet) Then mProps.NumRet = !NumRet
            If Not IsNull(!NumMeses) Then mProps.NumMeses = !NumMeses
            'JEAA 02/06/06
            If Not IsNull(!CodTipoComp) Then mProps.codTipoComprobante = !CodTipoComp
            If Not IsNull(!Deposito) Then mProps.Deposito = !Deposito
            'AUC 04/05/07
            If Not IsNull(!OtraRemuneracion) Then mProps.OtraRemuneracion = !OtraRemuneracion
            If Not IsNull(!XIII) Then mProps.XIII = !XIII
            If Not IsNull(!XIV) Then mProps.XIV = !XIV
            If Not IsNull(!Utilidad) Then mProps.Utilidad = !Utilidad
            If Not IsNull(!RebajaEspecial) Then mProps.RebajaEspecial = !RebajaEspecial
            'If Not IsNull(!IR) Then mprops.IR = !IR
            If Not IsNull(!RebajaEspecial3E) Then mProps.RebajaEspecial3E = !RebajaEspecial3E
            If Not IsNull(!ImpRentEmpl) Then mProps.ImpRentEmpl = !ImpRentEmpl
            If Not IsNull(!FR) Then mProps.FR = !FR
            If Not IsNull(!Deshaucio) Then mProps.Deshaucio = !Deshaucio
            If Not IsNull(!GPVivienda) Then mProps.GPVivienda = !GPVivienda
            If Not IsNull(!GPSalud) Then mProps.GPSalud = !GPSalud
            If Not IsNull(!GPEducacion) Then mProps.GPEducacion = !GPEducacion
            If Not IsNull(!GPAlimentacion) Then mProps.GPAlimentacion = !GPAlimentacion
            If Not IsNull(!GPVestimenta) Then mProps.GPVestimenta = !GPVestimenta
            If Not IsNull(!IngGravOtrosEmp) Then mProps.IngGravOtrosEmp = !IngGravOtrosEmp
            If Not IsNull(!GPOtrosEmp) Then mProps.GPOtrosEmp = !GPOtrosEmp
            If Not IsNull(!OtrasReb) Then mProps.OtrasReb = !OtrasReb
            If Not IsNull(!IRCausado) Then mProps.IRCausado = !IRCausado
            If Not IsNull(!ValorIR) Then mProps.ValorIR = !ValorIR
            If Not IsNull(!ValorIRAnt) Then mProps.ValorIRAnt = !ValorIRAnt
            If Not IsNull(!APOSOL1) Then mProps.APOSOL1 = !APOSOL1
            If Not IsNull(!APOSOL2) Then mProps.APOSOL2 = !APOSOL2
            If Not IsNull(!APOSOL3) Then mProps.APOSOL3 = !APOSOL3
            If Not IsNull(!APOSOL4) Then mProps.APOSOL4 = !APOSOL4
            If Not IsNull(!APOSOL5) Then mProps.APOSOL5 = !APOSOL5
            If Not IsNull(!APOSOL6) Then mProps.APOSOL6 = !APOSOL6
            If Not IsNull(!APOSOL7) Then mProps.APOSOL7 = !APOSOL7
            If Not IsNull(!APOSOL8) Then mProps.APOSOL8 = !APOSOL8
            If Not IsNull(!APOSOL9) Then mProps.APOSOL9 = !APOSOL9
            If Not IsNull(!APOSOL10) Then mProps.APOSOL10 = !APOSOL10
            If Not IsNull(!APOSOL11) Then mProps.APOSOL11 = !APOSOL11
            If Not IsNull(!APOSOL12) Then mProps.APOSOL12 = !APOSOL12
            If Not IsNull(!APOSOL13) Then mProps.APOSOL13 = !APOSOL13
            If Not IsNull(!APOSOL14) Then mProps.APOSOL14 = !APOSOL14
            If Not IsNull(!APOSOL15) Then mProps.APOSOL15 = !APOSOL15
            If Not IsNull(!APOSOL16) Then mProps.APOSOL16 = !APOSOL16
            If Not IsNull(!APOSOL17) Then mProps.APOSOL17 = !APOSOL17
            If Not IsNull(!APOSOL18) Then mProps.APOSOL18 = !APOSOL18
            If Not IsNull(!APOSOL19) Then mProps.APOSOL19 = !APOSOL19
            If Not IsNull(!APOSOL20) Then mProps.APOSOL20 = !APOSOL20
            If Not IsNull(!APOSOL21) Then mProps.APOSOL21 = !APOSOL21
        End If
    End With
End Sub


Public Function CalculaRetencion() As Currency
    'Calcula  el  valor de retencion
    CalculaRetencion = (mProps.BaseImponible * mProps.PorcRetInt) / 100
End Function

Public Sub Grabar()
    Dim sql As String, rs As Recordset
    
    'VerificaDatos
    
    
    sql = "SELECT * FROM IRKardex WHERE 1=0"
    Set rs = mProps.GNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !transid = mProps.GNComprobante.transid
        !TipoSalario = mProps.TipoSalario
        !Ingresos = mProps.Ingresos
        !OtraRemuneracion = mProps.OtraRemuneracion
        !Utilidad = mProps.Utilidad
        !IngGravOtrosEmp = mProps.IngGravOtrosEmp
        !XIII = mProps.XIII
        !XIV = mProps.XIV
        !FR = mProps.FR
        !Deshaucio = mProps.Deshaucio
        !AportePersonal = mProps.AportePersonal
        !GPOtrosEmp = mProps.GPOtrosEmp
        !GPVivienda = mProps.GPVivienda
        !GPSalud = mProps.GPSalud
        !GPEducacion = mProps.GPEducacion
        !GPAlimentacion = mProps.GPAlimentacion
        !GPVestimenta = mProps.GPVestimenta
        !RebajaEspecial = mProps.RebajaEspecial
        !RebajaEspecial3E = mProps.RebajaEspecial3E
        !ValorIRAnt = mProps.ValorIRAnt
        !BaseImponible = mProps.BaseImponible
        !IRCausado = mProps.IRCausado
        !OtrasReb = mProps.OtrasReb
        !ValorIR = mProps.ValorIR
        !ValorRetenido = mProps.ValorRetenido
        !Deposito = mProps.Deposito
        !IdEmpleado = mProps.IdEmpleado              'en caso RTP, porque escoge empleado por cada línea. Para CP, existe relación en cabecera GnComprobante.IdComp
        !APOSOL1 = mProps.APOSOL1
        !APOSOL2 = mProps.APOSOL2
        !APOSOL3 = mProps.APOSOL3
        !APOSOL4 = mProps.APOSOL4
        !APOSOL5 = mProps.APOSOL5
        !APOSOL6 = mProps.APOSOL6
        !APOSOL7 = mProps.APOSOL7
        !APOSOL8 = mProps.APOSOL8
        !APOSOL9 = mProps.APOSOL9
        !APOSOL10 = mProps.APOSOL10
        !APOSOL11 = mProps.APOSOL11
        !APOSOL12 = mProps.APOSOL12
        !APOSOL13 = mProps.APOSOL13
        !APOSOL14 = mProps.APOSOL14
        !APOSOL15 = mProps.APOSOL15
        !APOSOL16 = mProps.APOSOL16
        !APOSOL17 = mProps.APOSOL17
        !APOSOL18 = mProps.APOSOL18
        !APOSOL19 = mProps.APOSOL19
        !APOSOL20 = mProps.APOSOL20
        !APOSOL21 = mProps.APOSOL21
        
        !Orden = mProps.Orden
        .Update
        .Move 0                     'Para actualizar
        mProps.id = !id             'Guarda el Id(Autonumerico)
        .Close
    End With
    Set rs = Nothing
End Sub

'*********************************
' Atributos
'*********************************
Public Property Let CodEmpleado(ByVal value As String)
    Dim sql As String
    Dim rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodEmpleado Then
            If Len(value) > 0 Then
            'Actualiza IdResponsable también
            sql = "SELECT idProvCli FROM Empleado WHERE CodProvCli='" & value & "'"
            Set rs = mProps.GNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNComprobante.CodEmpleado", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdEmpleado = rs!IdProvcli
                mProps.CodEmpleado = value
            End If
            rs.Close
            Set rs = Nothing
        Else
            mProps.IdEmpleado = 0
        End If
    End If
    mProps.GNComprobante.SetModificado
End Property


Public Property Get CodEmpleado() As String
    CodEmpleado = mProps.CodEmpleado
End Property



Friend Property Let IdEmpleado(ByVal value As Long)
    mProps.IdEmpleado = value
    mProps.GNComprobante.SetModificado
End Property


Friend Property Get IdEmpleado() As Long
    IdEmpleado = mProps.IdEmpleado
End Property

Public Property Let FechaTipoRet(ByVal value As Date)
    mProps.fecha = value
End Property

Public Property Let TipoRetencion(ByVal value As String)
    mProps.GNComprobante.SetModificado
End Property


Public Property Get TipoRetencion() As String
    TipoRetencion = mProps.TipoRetencion
End Property

Friend Property Let IdPorcRetInt(ByVal value As Long)
    mProps.IdPorcRetInt = value
    mProps.GNComprobante.SetModificado
End Property


Friend Property Get IdPorcRetInt() As Long
    IdPorcRetInt = mProps.IdPorcRetInt
End Property

Public Property Let TipoRetInterno(ByVal value As String)
'    Dim sql As String, rs As Recordset
'
'    mprops.TipoRetInterno = value
'    'saca el ID y el Porcentaje
'    sql = "Select IdPorcentaje, Porcentaje From Porcentajes Where CodPorcentaje = '" & value & "'"
'    Set rs = mprops.gnComprobante.CoaMain.OpenRecordset(sql)
'    If rs.EOF Then
'        Err.Raise ERR_NOHAYCODIGO, "IRKardex_TipoRetinterno", MSGERR_NOHAYCODIGO
'        Exit Property
'    Else
'        mprops.IdPorcRetInt = rs!IdPorcentaje
'        mprops.PorcRetInt = rs!Porcentaje
'        mprops.gnComprobante.SetModificado
'    End If
'    Set rs = Nothing
'
'
''    Dim rs As Recordset, sql As String
''    If Value <> mProps.TipoRetencion Then
''        If Len(Value) > 0 Then
''            sql = "SELECT TipoRet FROM TipoRetencion WHERE TipoRet='" & Value & "'"
''            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
''            If rs.EOF Then
''                Err.Raise ERR_NOHAYCODIGO, "IRKardex_TipoRetencion", MSGERR_NOHAYCODIGO
''                Exit Property
''            End If
''            rs.Close
''        End If
''    End If
'
End Property

Public Property Get PorcRetInt() As Long
    PorcRetInt = mProps.PorcRetInt
End Property


Public Property Get TipoRetInterno() As String
    TipoRetInterno = mProps.TipoRetInterno
    
    
End Property


Public Property Let ValorRetenido(ByVal value As Currency)
    mProps.ValorRetenido = value
    mProps.GNComprobante.SetModificado
End Property


Public Property Get ValorRetenido() As Currency
    ValorRetenido = mProps.ValorRetenido
End Property



Public Property Let BaseImponible(ByVal value As Currency)
    mProps.BaseImponible = value
    mProps.GNComprobante.SetModificado
End Property


Public Property Get BaseImponible() As Currency
    BaseImponible = mProps.BaseImponible
End Property

Public Property Let AportePersonal(ByVal value As Currency)
    mProps.AportePersonal = value
    mProps.GNComprobante.SetModificado
End Property

Public Property Get AportePersonal() As Currency
    AportePersonal = mProps.AportePersonal
End Property

Public Property Let Ingresos(ByVal value As Currency)
    mProps.Ingresos = value
    mProps.GNComprobante.SetModificado
End Property

Public Property Get Ingresos() As Currency
    Ingresos = mProps.Ingresos
End Property

Public Property Let TipoSalario(ByVal value As String)
    mProps.TipoSalario = value
    mProps.GNComprobante.SetModificado
End Property

Public Property Get TipoSalario() As String
    TipoSalario = mProps.TipoSalario
End Property

Friend Property Get id() As Long
    id = mProps.id
End Property

Friend Property Set GNComprobante(ByVal value As GNComprobante)
    Set mProps.GNComprobante = value
End Property

Public Property Get GNComprobante() As GNComprobante
Attribute GNComprobante.VB_Description = "relacion al Obj Padre"
    Set GNComprobante = mProps.GNComprobante
End Property

Public Property Let NumRet(ByVal value As Integer)
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "IRKardex.NumRetenciones", MSGERR_INVALIDO
    If value < 0 Then Err.Raise ERR_INVALIDO, "IRKardex.NumRetenciones", MSGERR_INVALIDO
    mProps.NumRet = value
    mbooModificado = True
End Property

Public Property Get NumRet() As Integer
    NumRet = mProps.NumRet
End Property

Public Property Let NumMeses(ByVal value As Currency)
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "IRKardex.NumMeses", MSGERR_INVALIDO
    If value < 0 Then Err.Raise ERR_INVALIDO, "IRKardex.NumMeses", MSGERR_INVALIDO
    mProps.NumMeses = value
    mbooModificado = True
End Property

Public Property Get NumMeses() As Currency
    NumMeses = mProps.NumMeses
End Property

Private Sub Class_Terminate()
    Set mProps.GNComprobante = Nothing
End Sub

Friend Sub SetModificado()
    mbooModificado = True
End Sub

Public Property Get Modificado() As Boolean
    Modificado = mbooModificado
End Property

Public Property Let Orden(ByVal value As Integer)
    mProps.Orden = value
    mProps.GNComprobante.SetModificado
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property

Public Property Let codTipoComprobante(ByVal value As String)
    mProps.codTipoComprobante = value
    mProps.GNComprobante.SetModificado
End Property

Public Property Get codTipoComprobante() As String
    codTipoComprobante = mProps.codTipoComprobante
End Property

Public Property Let Deposito(ByVal value As Currency)
    mProps.Deposito = value
    mProps.GNComprobante.SetModificado
End Property

Public Property Get Deposito() As Currency
    Deposito = mProps.Deposito
End Property
'AUC 05/05/07
Public Property Let OtraRemuneracion(ByVal value As Currency)
    mProps.OtraRemuneracion = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get OtraRemuneracion() As Currency
    OtraRemuneracion = mProps.OtraRemuneracion
End Property
Public Property Let XIII(ByVal value As Currency)
    mProps.XIII = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get XIII() As Currency
    XIII = mProps.XIII
End Property
Public Property Let XIV(ByVal value As Currency)
    mProps.XIV = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get XIV() As Currency
    XIV = mProps.XIV
End Property
Public Property Let Utilidad(ByVal value As Currency)
    mProps.Utilidad = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get Utilidad() As Currency
    Utilidad = mProps.Utilidad
End Property
Public Property Let RebajaEspecial(ByVal value As Currency)
    mProps.RebajaEspecial = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get RebajaEspecial() As Currency
    RebajaEspecial = mProps.RebajaEspecial
End Property
Public Property Let IR(ByVal value As Currency)
    mProps.IR = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get IR() As Currency
    IR = mProps.IR
End Property

Public Property Let RebajaEspecial3E(ByVal value As Currency)
    mProps.RebajaEspecial3E = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get RebajaEspecial3E() As Currency
    RebajaEspecial3E = mProps.RebajaEspecial3E
End Property

Public Property Let ImpRentEmpl(ByVal value As Currency)
    mProps.ImpRentEmpl = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get ImpRentEmpl() As Currency
    ImpRentEmpl = mProps.ImpRentEmpl
End Property
    
Public Property Let FR(ByVal value As Currency)
    mProps.FR = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get FR() As Currency
    FR = mProps.FR
End Property
Public Property Let Deshaucio(ByVal value As Currency)
    mProps.Deshaucio = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get Deshaucio() As Currency
    Deshaucio = mProps.Deshaucio
End Property
Public Property Let GPVivienda(ByVal value As Currency)
    mProps.GPVivienda = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get GPVivienda() As Currency
    GPVivienda = mProps.GPVivienda
End Property
Public Property Let GPSalud(ByVal value As Currency)
    mProps.GPSalud = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get GPSalud() As Currency
    GPSalud = mProps.GPSalud
End Property
Public Property Let GPEducacion(ByVal value As Currency)
    mProps.GPEducacion = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get GPEducacion() As Currency
    GPEducacion = mProps.GPEducacion
End Property
Public Property Let GPAlimentacion(ByVal value As Currency)
    mProps.GPAlimentacion = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get GPAlimentacion() As Currency
    GPAlimentacion = mProps.GPAlimentacion
End Property
Public Property Let GPVestimenta(ByVal value As Currency)
    mProps.GPVestimenta = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get GPVestimenta() As Currency
    GPVestimenta = mProps.GPVestimenta
End Property
Public Property Let IngGravOtrosEmp(ByVal value As Currency)
    mProps.IngGravOtrosEmp = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get IngGravOtrosEmp() As Currency
    IngGravOtrosEmp = mProps.IngGravOtrosEmp
End Property
Public Property Let GPOtrosEmp(ByVal value As Currency)
    mProps.GPOtrosEmp = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get GPOtrosEmp() As Currency
    GPOtrosEmp = mProps.GPOtrosEmp
End Property

Public Property Let OtrasReb(ByVal value As Currency)
    mProps.OtrasReb = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get OtrasReb() As Currency
    OtrasReb = mProps.OtrasReb
End Property
Public Property Let IRCausado(ByVal value As Currency)
    mProps.IRCausado = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get IRCausado() As Currency
    IRCausado = mProps.IRCausado
End Property
Public Property Let ValorIR(ByVal value As Currency)
    mProps.ValorIR = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get ValorIR() As Currency
    ValorIR = mProps.ValorIR
End Property
Public Property Let ValorIRAnt(ByVal value As Currency)
    mProps.ValorIRAnt = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get ValorIRAnt() As Currency
    ValorIRAnt = mProps.ValorIRAnt
End Property

Public Property Let APOSOL1(ByVal value As Currency)
    mProps.APOSOL1 = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get APOSOL1() As Currency
    APOSOL1 = mProps.APOSOL1
End Property

Public Property Let APOSOL2(ByVal value As Currency)
    mProps.APOSOL2 = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get APOSOL2() As Currency
    APOSOL2 = mProps.APOSOL2
End Property

Public Property Let APOSOL3(ByVal value As Currency)
    mProps.APOSOL3 = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get APOSOL3() As Currency
    APOSOL3 = mProps.APOSOL3
End Property

Public Property Let APOSOL4(ByVal value As Currency)
    mProps.APOSOL4 = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get APOSOL4() As Currency
    APOSOL4 = mProps.APOSOL4
End Property

Public Property Let APOSOL5(ByVal value As Currency)
    mProps.APOSOL5 = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get APOSOL5() As Currency
    APOSOL5 = mProps.APOSOL5
End Property

Public Property Let APOSOL6(ByVal value As Currency)
    mProps.APOSOL6 = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get APOSOL6() As Currency
    APOSOL6 = mProps.APOSOL6
End Property

Public Property Let APOSOL7(ByVal value As Currency)
    mProps.APOSOL7 = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get APOSOL7() As Currency
    APOSOL7 = mProps.APOSOL7
End Property

Public Property Let APOSOL8(ByVal value As Currency)
    mProps.APOSOL8 = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get APOSOL8() As Currency
    APOSOL8 = mProps.APOSOL8
End Property

Public Property Let APOSOL9(ByVal value As Currency)
    mProps.APOSOL9 = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get APOSOL9() As Currency
    APOSOL9 = mProps.APOSOL9
End Property

Public Property Let APOSOL10(ByVal value As Currency)
    mProps.APOSOL10 = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get APOSOL10() As Currency
    APOSOL10 = mProps.APOSOL10
End Property
Public Property Let APOSOL11(ByVal value As Currency)
    mProps.APOSOL11 = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get APOSOL11() As Currency
    APOSOL11 = mProps.APOSOL11
End Property

Public Property Let APOSOL12(ByVal value As Currency)
    mProps.APOSOL12 = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get APOSOL12() As Currency
    APOSOL12 = mProps.APOSOL12
End Property

Public Property Let APOSOL13(ByVal value As Currency)
    mProps.APOSOL13 = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get APOSOL13() As Currency
    APOSOL13 = mProps.APOSOL13
End Property

Public Property Let APOSOL14(ByVal value As Currency)
    mProps.APOSOL14 = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get APOSOL14() As Currency
    APOSOL14 = mProps.APOSOL14
End Property

Public Property Let APOSOL15(ByVal value As Currency)
    mProps.APOSOL15 = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get APOSOL15() As Currency
    APOSOL15 = mProps.APOSOL15
End Property

Public Property Let APOSOL16(ByVal value As Currency)
    mProps.APOSOL16 = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get APOSOL16() As Currency
    APOSOL16 = mProps.APOSOL16
End Property

Public Property Let APOSOL17(ByVal value As Currency)
    mProps.APOSOL17 = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get APOSOL17() As Currency
    APOSOL17 = mProps.APOSOL17
End Property

Public Property Let APOSOL18(ByVal value As Currency)
    mProps.APOSOL18 = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get APOSOL18() As Currency
    APOSOL18 = mProps.APOSOL18
End Property

Public Property Let APOSOL19(ByVal value As Currency)
    mProps.APOSOL19 = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get APOSOL19() As Currency
    APOSOL19 = mProps.APOSOL19
End Property

Public Property Let APOSOL20(ByVal value As Currency)
    mProps.APOSOL20 = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get APOSOL20() As Currency
    APOSOL20 = mProps.APOSOL20
End Property

Public Property Let APOSOL21(ByVal value As Currency)
    mProps.APOSOL21 = value
    mProps.GNComprobante.SetModificado
End Property
Public Property Get APOSOL21() As Currency
    APOSOL21 = mProps.APOSOL21
End Property


