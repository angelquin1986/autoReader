VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdEmpleado As Long
    Sexo As Byte                    '07/09/2000   Angel --->Antes Long
    EstadoCivil As Byte             '07/09/2000   Angel --->Antes Long
    NumCargas As Byte               '04/09/2000   Angel --->Antes Long
    idCargo As Long
    CodCargoNum As String           '12/11/2000   Angel --->Nuevo Campo
    IdDepartamento As Long
    Salario As Currency
    TipoSalario As Byte             '04/09/2000   Angel --->Antes Long
    NumeroCuenta As String          '04/09/2000   Angel --->Nuevo Campo
    NumeroAfiliacion As String      '04/09/2000   Angel --->Nuevo Campo
    FechaIngreso As Date
    FechaEgreso As Date
    FechaGrabado As Date
    ValorQuincena As Currency 'AUC 08/009/06 AUC agregado  ingresar la quincena
    Antiguedad As Long
    IdAgencia As String
    Varios1 As Currency
    Varios2 As Currency
    Varios3 As Currency
    Varios4 As Currency
    Varios5 As Currency
    Varios6 As Currency
    Varios7 As Currency
    Varios8 As Currency
    PagarSeguro As Byte
    Contador  As Currency
    FechaPagoSeguro As Date
    BandFR As Boolean
    BandPagarHE As Boolean
    mobjEmpresa As Empresa
    BandActivo As Boolean
End Type
Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean
Private mcolHistorial As Collection
Private mcolPrestamo As Collection

Friend Property Set Empresa(ByVal obj As Empresa)
    Set mProps.mobjEmpresa = obj
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.mobjEmpresa
End Property



Public Property Let Sexo(ByVal value As Byte)          '07/09/2000   Angel --->Antes Long
    '0=Femenino, 1=Masculino
    If (Not IsNumeric(value)) Or (Val(value) < 0 Or Val(value) > 1) Then Err.Raise ERR_INVALIDO, "PCEmpleado.Sexo", MSGERR_INVALIDO
    mProps.Sexo = value
    mbooModificado = True
End Property

Public Property Get Sexo() As Byte                     '07/09/2000   Angel --->Antes Long
    Sexo = mProps.Sexo
End Property

Public Property Let EstadoCivil(ByVal value As Byte)   '07/09/2000   Angel --->Antes Long
    '0=Soltero, 1=Casado, 2=Divorciado, 3=Viudo
    If (Not IsNumeric(value)) Or (Val(value) < 0 Or Val(value) > 4) Then Err.Raise ERR_INVALIDO, "PCEmpleado.EstadoCivil", MSGERR_INVALIDO 'AUC Cambiado de 3 a 4
    mProps.EstadoCivil = value
    mbooModificado = True
End Property

Public Property Get EstadoCivil() As Byte              '07/09/2000   Angel --->Antes Long
    EstadoCivil = mProps.EstadoCivil
End Property

Public Property Let NumCargas(ByVal value As Byte)   '04/09/2000  Angel --->Antes Long
    If (Not IsNumeric(value)) Or (value > 250) Then Err.Raise ERR_INVALIDO, "PCEmpleado.NumCargas", MSGERR_INVALIDO
    mProps.NumCargas = value
    mbooModificado = True
End Property

Public Property Get NumCargas() As Byte              '04/09/2000  Angel --->Antes Long
    NumCargas = mProps.NumCargas
End Property

Public Property Let idCargo(ByVal value As Long)
    mProps.idCargo = value
    mbooModificado = True
End Property

Public Property Get idCargo() As Long
    idCargo = mProps.idCargo
End Property

Public Property Let CodCargoNum(ByVal value As String) '12/11/2000 Nuevo Campo
    If Len(value) > 20 Then Err.Raise ERR_INVALIDO, "PCEmpleado.CodCargo", MSGERR_INVALIDO
    mProps.CodCargoNum = value
    mbooModificado = True
End Property

Public Property Get CodCargoNum() As String            '12/11/2000 Nuevo Campo
    CodCargoNum = mProps.CodCargoNum
End Property

Public Property Let IdDepartamento(ByVal value As Long)
    mProps.IdDepartamento = value
    mbooModificado = True
End Property

Public Property Get IdDepartamento() As Long
    IdDepartamento = mProps.IdDepartamento
End Property

 
Public Property Let TipoSalario(ByVal value As Byte)  '04/09/2000  Angel --->Antes Long
    '0=hora, 1=semanal, 2=quincenal, 3=mensual
    If (Not IsNumeric(value)) Or (Val(value) < 0 Or Val(value) > 3) Then Err.Raise ERR_INVALIDO, "PCEmpleado.TipoSalario", MSGERR_INVALIDO
    mProps.TipoSalario = value
    mbooModificado = True
End Property

Public Property Get TipoSalario() As Byte             '04/09/2000  Angel --->Antes Long
    TipoSalario = mProps.TipoSalario
End Property

Public Property Let NumeroCuenta(ByVal value As String) '04/09/2000   Angel --->Propiedad Nueva
    If Len(value) > 30 Then Err.Raise ERR_INVALIDO, "PCEmpleado.NumeroCuenta", MSGERR_INVALIDO
    mProps.NumeroCuenta = value
    mbooModificado = True
End Property

Public Property Get NumeroCuenta() As String            '04/09/2000   Angel --->Propiedad Nueva
    NumeroCuenta = mProps.NumeroCuenta
End Property

Public Property Let NumeroAfiliacion(ByVal value As String) '04/09/2000   Angel --->Propiedad Nueva
    If Len(value) > 30 Then Err.Raise ERR_INVALIDO, "PCEmpleado.NumeroAfiliacion", MSGERR_INVALIDO
    mProps.NumeroAfiliacion = value
    mbooModificado = True
End Property

Public Property Get NumeroAfiliacion() As String            '04/09/2000   Angel --->Propiedad Nueva
    NumeroAfiliacion = mProps.NumeroAfiliacion
End Property

Public Property Let FechaIngreso(ByVal value As Date)
    If Not IsDate(value) Then Err.Raise ERR_INVALIDO, "PCEmpleado.FechaIngreso", MSGERR_INVALIDO
    mProps.FechaIngreso = value
    mbooModificado = True
End Property

Public Property Get FechaIngreso() As Date
    FechaIngreso = mProps.FechaIngreso
End Property

Public Property Let FechaEgreso(ByVal value As Date)
    If Not IsDate(value) Then Err.Raise ERR_INVALIDO, "PCEmpleado.FechaEgreso", MSGERR_INVALIDO
    mProps.FechaEgreso = value
    mbooModificado = True
End Property

Public Property Get FechaEgreso() As Date
    FechaEgreso = mProps.FechaEgreso
End Property

Public Property Let FechaGrabado(ByVal value As Date)
    If Not IsDate(value) Then Err.Raise ERR_INVALIDO, "PCEmpleado.FechaGrabado", MSGERR_INVALIDO
    mProps.FechaGrabado = value
    mbooModificado = True
End Property

Public Property Get FechaGrabado() As Date
    FechaGrabado = mProps.FechaGrabado
End Property

Public Property Get EsNuevo() As Boolean
    EsNuevo = mbooEsNuevo
End Property

Public Property Get Modificado() As Boolean
    Modificado = mbooModificado
End Property

Private Sub Class_Initialize()
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.mobjEmpresa = Nothing
End Sub

Public Sub Recuperar(ByVal Id As Long)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT * FROM PCEmpleado WHERE idEmpleado = " & Id
    
    Set rs = mProps.mobjEmpresa.OpenRecordset(sql)
        
    With rs
        Do While Not .EOF
            mProps.IdEmpleado = !IdEmpleado
            mProps.Sexo = IIf(IsNull(!Sexo), 0, !Sexo)
            mProps.EstadoCivil = IIf(IsNull(!EstadoCivil), 0, !EstadoCivil)
            mProps.NumCargas = IIf(IsNull(!NumCargas), 0, !NumCargas)
            'mProps.CodCargo = IIf(IsNull(!CodCargo), "", !CodCargo)
            mProps.CodCargoNum = IIf(IsNull(!CodCargoNum), "", !CodCargoNum)                '12/11/2000  Angel

            mProps.Salario = IIf(IsNull(!Salario), 0, !Salario)
            mProps.TipoSalario = IIf(IsNull(!TipoSalario), 0, !TipoSalario)
            mProps.NumeroCuenta = IIf(IsNull(!NumeroCuenta), "", !NumeroCuenta)             '04/09/2000  Angel --->Campo Nuevo
            mProps.NumeroAfiliacion = IIf(IsNull(!NumeroAfiliacion), "", !NumeroAfiliacion) '04/09/2000  Angel --->Antes Long
            mProps.FechaIngreso = IIf(IsNull(!FechaIngreso), #12:00:00 AM#, !FechaIngreso)
            mProps.FechaEgreso = IIf(IsNull(!FechaEgreso), #12:00:00 AM#, !FechaEgreso)



            mProps.Varios1 = IIf(IsNull(!Varios1), 0, !Varios1)
            mProps.Varios2 = IIf(IsNull(!Varios2), 0, !Varios2)
            mProps.Varios3 = IIf(IsNull(!Varios3), 0, !Varios3)
            mProps.Varios4 = IIf(IsNull(!Varios4), 0, !Varios4)
            mProps.Varios5 = IIf(IsNull(!Varios5), 0, !Varios5)
            mProps.Varios6 = IIf(IsNull(!Varios6), 0, !Varios6)
            mProps.Varios7 = IIf(IsNull(!Varios7), 0, !Varios7)
            mProps.Varios8 = IIf(IsNull(!Varios8), 0, !Varios8)
            mProps.PagarSeguro = !PagarSeguro
            mProps.Contador = IIf(IsNull(!Contador), 0, !Contador)
            mProps.FechaPagoSeguro = IIf(IsNull(!FechaPagoSeguro), #12:00:00 AM#, !FechaPagoSeguro)
            mProps.BandFR = !BandFR
            mProps.BandPagarHE = !BandPagarHE
            mProps.BandActivo = !BandActivo
            rs.MoveNext
        Loop
            .Close
    End With
'    CargaHistorial
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Grabar(ByVal IdEmpleado As Long)
    Dim sql As String, rs As Recordset
    On Error GoTo Etiq
    
        'sql = "SELECT * FROM PCEmpleado WHERE idempleado = " & IdEmpleado
'        Set rs = mProps.mobjEmpresa.OpenRecordsetParaEdit(sql)
        
            sql = "SELECT * FROM PCEmpleado WHERE idempleado = " & IdEmpleado
            Set rs = mProps.mobjEmpresa.OpenRecordsetParaEdit(sql)
            
        
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            sql = "SELECT * FROM PCEmpleado WHERE 1=0"
            Set rs = mProps.mobjEmpresa.OpenRecordsetParaEdit(sql)
            rs.AddNew
        Else
        End If
    
    With rs
        !IdEmpleado = IdEmpleado
        !Sexo = Me.Sexo
        !EstadoCivil = Me.EstadoCivil
        !NumCargas = Me.NumCargas
        !CodCargoNum = Me.CodCargoNum
        !Salario = Me.Salario
        !TipoSalario = Me.TipoSalario
        !NumeroCuenta = Me.NumeroCuenta
        !NumeroAfiliacion = Me.NumeroAfiliacion
        !FechaIngreso = Me.FechaIngreso
        !FechaEgreso = Me.FechaEgreso
        !Varios1 = Me.Varios1
        !Varios2 = Me.Varios2
        !Varios3 = Me.Varios3
        !Varios4 = Me.Varios4
        !Varios5 = Me.Varios5
        !Varios6 = Me.Varios6
        !Varios7 = Me.Varios7
        !Varios8 = Me.Varios8
        !PagarSeguro = Me.PagarSeguro
        !Contador = Me.Contador
        !FechaPagoSeguro = Me.FechaPagoSeguro
        !BandFR = Me.BandFR
        !BandPagarHE = Me.BandPagarHE
        !BandActivo = Me.BandActivo
        .Update
        .Close
    End With
'    GrabarHistorial (mProps.IdEmpleado)
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    Me.Empresa.RaiseRefreshEvent REFRESH_PCEEMPLEADO
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub


'AUC 08/09/06
Public Property Let ValorQuincena(ByVal value As Currency)
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "PCEmpleado.ValorQuincena", MSGERR_INVALIDO
    mProps.ValorQuincena = value
    mbooModificado = True
End Property

'AUC 08/09/06
Public Property Get ValorQuincena() As Currency   'Antes Long. Mofificado por Angel
    ValorQuincena = mProps.ValorQuincena
End Property

Public Property Let Antiguedad(ByVal value As Long)
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "PCEmpleado.Antiguedad", MSGERR_INVALIDO
    mProps.Antiguedad = value
    mbooModificado = True
End Property

'AUC 08/09/06
Public Property Get Antiguedad() As Long
    Antiguedad = mProps.Antiguedad
End Property

Public Property Let iddAgencia(ByVal value As Long)
    mProps.IdAgencia = value
    mbooModificado = True
End Property

Public Property Get IdAgencia() As String
    IdAgencia = mProps.IdAgencia
End Property

Public Property Let Varios1(ByVal value As Currency) 'Antes Long. Mofificado por Angel
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "PCEmpleado.Varios1", MSGERR_INVALIDO
    mProps.Varios1 = value
    mbooModificado = True
End Property

Public Property Get Varios1() As Currency   'Antes Long. Mofificado por Angel
    Varios1 = mProps.Varios1
End Property

Public Property Let Varios2(ByVal value As Currency) 'Antes Long. Mofificado por Angel
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "PCEmpleado.Varios2", MSGERR_INVALIDO
    mProps.Varios2 = value
    mbooModificado = True
End Property

Public Property Get Varios2() As Currency   'Antes Long. Mofificado por Angel
    Varios2 = mProps.Varios2
End Property

Public Property Let Salario(ByVal value As Currency) 'Antes Long. Mofificado por Angel
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "PCEmpleado.Salario", MSGERR_INVALIDO
    mProps.Salario = value
    mbooModificado = True
End Property

Public Property Get Salario() As Currency   'Antes Long. Mofificado por Angel
    Salario = mProps.Salario
End Property

Public Property Let Varios3(ByVal value As Currency) 'Antes Long. Mofificado por Angel
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "PCEmpleado.Varios3", MSGERR_INVALIDO
    mProps.Varios3 = value
    mbooModificado = True
End Property

Public Property Get Varios3() As Currency   'Antes Long. Mofificado por Angel
    Varios3 = mProps.Varios3
End Property

Public Property Let Varios4(ByVal value As Currency) 'Antes Long. Mofificado por Angel
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "PCEmpleado.Varios4", MSGERR_INVALIDO
    mProps.Varios4 = value
    mbooModificado = True
End Property

Public Property Get Varios4() As Currency   'Antes Long. Mofificado por Angel
    Varios4 = mProps.Varios4
End Property

Public Property Get Varios5() As Currency   'Antes Long. Mofificado por Angel
    Varios5 = mProps.Varios5
End Property

Public Property Let Varios5(ByVal value As Currency) 'Antes Long. Mofificado por Angel
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "PCEmpleado.Varios5", MSGERR_INVALIDO
    mProps.Varios5 = value
    mbooModificado = True
End Property

Public Property Get Varios6() As Currency   'Antes Long. Mofificado por Angel
    Varios6 = mProps.Varios6
End Property

Public Property Let Varios6(ByVal value As Currency) 'Antes Long. Mofificado por Angel
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "PCEmpleado.Varios6", MSGERR_INVALIDO
    mProps.Varios6 = value
    mbooModificado = True
End Property

Public Property Get Varios7() As Currency   'Antes Long. Mofificado por Angel
    Varios7 = mProps.Varios7
End Property

Public Property Let Varios7(ByVal value As Currency) 'Antes Long. Mofificado por Angel
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "PCEmpleado.Varios7", MSGERR_INVALIDO
    mProps.Varios7 = value
    mbooModificado = True
End Property

Public Property Get Varios8() As Currency   'Antes Long. Mofificado por Angel
    Varios8 = mProps.Varios8
End Property

Public Property Let Varios8(ByVal value As Currency) 'Antes Long. Mofificado por Angel
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "PCEmpleado.Varios8", MSGERR_INVALIDO
    mProps.Varios8 = value
    mbooModificado = True
End Property

Public Property Let PagarSeguro(ByVal value As Byte)
    mProps.PagarSeguro = value
    mbooModificado = True
End Property

Public Property Get PagarSeguro() As Byte
    PagarSeguro = mProps.PagarSeguro
End Property
'AUC
Public Property Get Contador() As Currency
    Contador = mProps.Contador
End Property

Public Property Let Contador(ByVal value As Currency)
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "PCEmpleado.Contador", MSGERR_INVALIDO
    mProps.Contador = value
    mbooModificado = True
End Property

Public Property Let FechaPagoSeguro(ByVal value As Date)
    If Not IsDate(value) Then Err.Raise ERR_INVALIDO, "PCEmpleado.FechaPagoSeguro", MSGERR_INVALIDO
    mProps.FechaPagoSeguro = value
    mbooModificado = True
End Property

Public Property Get FechaPagoSeguro() As Date
    FechaPagoSeguro = mProps.FechaPagoSeguro
End Property

Public Property Let BandFR(ByVal value As Boolean)
    mProps.BandFR = value
    mbooModificado = True
End Property

Public Property Get BandFR() As Boolean
    BandFR = mProps.BandFR
End Property

Public Property Let BandPagarHE(ByVal value As Boolean)
    mProps.BandPagarHE = value
    mbooModificado = True
End Property

Public Property Get BandPagarHE() As Boolean
    BandPagarHE = mProps.BandPagarHE
End Property

Public Property Let BandActivo(ByVal value As Boolean)
    mProps.BandActivo = value
    mbooModificado = True
End Property

Public Property Get BandActivo() As Boolean
    BandActivo = mProps.BandActivo
End Property

Public Property Get IdEmpleado() As Long
    IdEmpleado = mProps.IdEmpleado
End Property

