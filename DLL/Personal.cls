VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Personal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdEmpleado As Long
    Sexo As Byte
    EstadoCivil As Byte
    NumCargas As Byte
    idCargo As Long
    CodCargoNum As String
    IdDepartamento As Long
    Salario As Currency
    TipoSalario As Byte             '04/09/2000   Angel --->Antes Long
    NumeroCuenta As String          '04/09/2000   Angel --->Nuevo Campo
    NumeroAfiliacion As String      '04/09/2000   Angel --->Nuevo Campo
    FechaIngreso As Date
    FechaEgreso As Date
    fechagrabado As Date
    ValorQuincena As Currency 'AUC 08/009/06 AUC agregado  ingresar la quincena
    Antiguedad As Long
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
    BandActivo As Boolean
    RutaFoto As String
    CambiaSueldo As Boolean
    SalarioAnt As Currency
    Tipo As Byte
    Discapacidad As String
    IdTipoRol As Long 'agrega tiene que filtrar igual que los elementos
    IdTipoRol1 As Long
    IdTipoRol2 As Long
    IdTipoRol3 As Long
    IdTipoRol4 As Long
    IdTipoRol5 As Long
    IdTipoRol6 As Long
    IdTipoRol7 As Long
    IdTipoRol8 As Long
    IdTipoRol9 As Long
    IdTipoRol10 As Long
    UDecimoTercero As Currency
    idUltElemento As Long
    Codelemento As String
    Password As String
    CodigoReloj As String
    BandEmpPublico As Boolean
    BandAPGerente As Boolean
    BandLosep As Boolean
    IdGrupo1 As Long
    CodGrupo1 As String
    IdGrupo2 As Long
    CodGrupo2 As String
    IdGrupo3 As Long
    CodGrupo3 As String
    IdGrupo4 As Long
    CodGrupo4 As String
    CodigoLosep As String
    CambioCodigoLosep As Boolean
    CodigoLosepAnt As String
    Apellido As String
    Nombre As String
    TipoDiscapacidad As String
    Residencia As String
    Imposicion As String
    CedSustituto As String
    PorcDiscapacidad As Integer
    IDSucursal As Integer
    CodSucursal As String
    IdTipoDocumento As Long
    codtipoDocumento As String
    Ciclo As String
    BandXIII As Boolean
    BandXIV As Boolean
    BandGestion As Boolean
    GP As Currency
    VGP As Currency
    gpMeses As String
    GPAnt As Currency
    CambiaGP As Boolean
    FechaFinalFR As Date
    FechaFinalFRAnt As Date
    CambiaFFR As Boolean
    BandBenGalp As Boolean
    mobjEmpresa As Empresa
End Type
Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean
Private mcolHistorial As Collection
Private mcolCargas As Collection
Private mcolEstudios As Collection
Private mcolHistVarios As Collection
Private mcolHistorialGP As Collection
Private mcolHistorialFechaFR As Collection

Friend Property Set Empresa(ByVal obj As Empresa)
    Set mProps.mobjEmpresa = obj
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.mobjEmpresa
End Property



Public Property Let Sexo(ByVal value As Byte)          '07/09/2000   Angel --->Antes Long
    '0=Femenino, 1=Masculino
    If (Not IsNumeric(value)) Or (Val(value) < 0 Or Val(value) > 1) Then Err.Raise ERR_INVALIDO, "Personal.Sexo", MSGERR_INVALIDO
    mProps.Sexo = value
    mbooModificado = True
End Property

Public Property Get Sexo() As Byte                     '07/09/2000   Angel --->Antes Long
    Sexo = mProps.Sexo
End Property

Public Property Let EstadoCivil(ByVal value As Byte)   '07/09/2000   Angel --->Antes Long
    '0=Soltero, 1=Casado, 2=Divorciado, 3=Viudo
    If (Not IsNumeric(value)) Or (Val(value) < 0 Or Val(value) > 4) Then Err.Raise ERR_INVALIDO, "Personal.EstadoCivil", MSGERR_INVALIDO 'AUC Cambiado de 3 a 4
    mProps.EstadoCivil = value
    mbooModificado = True
End Property

Public Property Get EstadoCivil() As Byte              '07/09/2000   Angel --->Antes Long
    EstadoCivil = mProps.EstadoCivil
End Property

Public Property Let NumCargas(ByVal value As Byte)   '04/09/2000  Angel --->Antes Long
    If (Not IsNumeric(value)) Or (value > 250) Then Err.Raise ERR_INVALIDO, "Personal.NumCargas", MSGERR_INVALIDO
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
    If Len(value) > 20 Then Err.Raise ERR_INVALIDO, "Personal.CodCargo", MSGERR_INVALIDO
    mProps.CodCargoNum = value
    mbooModificado = True
End Property

Public Property Get CodCargoNum() As String            '12/11/2000 Nuevo Campo
    CodCargoNum = mProps.CodCargoNum
End Property

Public Property Let IdEmpleado(ByVal value As Long)
    mProps.IdEmpleado = value
    mbooModificado = True
End Property

Public Property Let TipoSalario(ByVal value As Byte)  '04/09/2000  Angel --->Antes Long
    '0=hora, 1=semanal, 2=quincenal, 3=mensual
    If (Not IsNumeric(value)) Or (Val(value) < 0 Or Val(value) > 3) Then Err.Raise ERR_INVALIDO, "Personal.TipoSalario", MSGERR_INVALIDO
    mProps.TipoSalario = value
    mbooModificado = True
End Property

Public Property Get TipoSalario() As Byte             '04/09/2000  Angel --->Antes Long
    TipoSalario = mProps.TipoSalario
End Property

Public Property Let NumeroCuenta(ByVal value As String) '04/09/2000   Angel --->Propiedad Nueva
    If Len(value) > 30 Then Err.Raise ERR_INVALIDO, "Personal.NumeroCuenta", MSGERR_INVALIDO
    mProps.NumeroCuenta = value
    mbooModificado = True
End Property

Public Property Get NumeroCuenta() As String            '04/09/2000   Angel --->Propiedad Nueva
    NumeroCuenta = mProps.NumeroCuenta
End Property

Public Property Let NumeroAfiliacion(ByVal value As String) '04/09/2000   Angel --->Propiedad Nueva
    If Len(value) > 30 Then Err.Raise ERR_INVALIDO, "Personal.NumeroAfiliacion", MSGERR_INVALIDO
    mProps.NumeroAfiliacion = value
    mbooModificado = True
End Property

Public Property Get NumeroAfiliacion() As String            '04/09/2000   Angel --->Propiedad Nueva
    NumeroAfiliacion = mProps.NumeroAfiliacion
End Property

Public Property Let FechaIngreso(ByVal value As Date)
    If Not IsDate(value) Then Err.Raise ERR_INVALIDO, "Personal.FechaIngreso", MSGERR_INVALIDO
    mProps.FechaIngreso = value
    mbooModificado = True
End Property

Public Property Get FechaIngreso() As Date
    FechaIngreso = mProps.FechaIngreso
End Property

Public Property Let FechaEgreso(ByVal value As Date)
    If Not IsDate(value) Then Err.Raise ERR_INVALIDO, "Personal.FechaEgreso", MSGERR_INVALIDO
    mProps.FechaEgreso = value
    mbooModificado = True
End Property

Public Property Get FechaEgreso() As Date
    FechaEgreso = mProps.FechaEgreso
End Property

Public Property Let fechagrabado(ByVal value As Date)
    If Not IsDate(value) Then Err.Raise ERR_INVALIDO, "Personal.FechaGrabado", MSGERR_INVALIDO
    mProps.fechagrabado = value
    mbooModificado = True
End Property

Public Property Get fechagrabado() As Date
    fechagrabado = mProps.fechagrabado
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

Public Sub Recuperar(ByVal id As Long)
    Dim sql As String, rs As Recordset
    sql = "SELECT personal.*,pcgrupo1.codgrupo1,pcgrupo2.codgrupo2  "
    sql = sql & ",pcgrupo3.codgrupo3,pcgrupo4.codgrupo4,gnsucursal.codSucursal "
    sql = sql & ",ax.codtipodocumento "
    sql = sql & " from Personal "
    sql = sql & "Left Join PCGrupo1 on pcgrupo1.idgrupo1 = personal.idgrupo1 "
    sql = sql & "Left Join PCGrupo2 on pcgrupo2.idgrupo2 = personal.idgrupo2 "
    sql = sql & "Left Join PCGrupo3 on pcgrupo3.idgrupo3 = personal.idgrupo3 "
    sql = sql & "Left Join PCGrupo4 on pcgrupo4.idgrupo4 = personal.idgrupo4 "
    sql = sql & "Left Join GNSucursal on GNSucursal.idsucursal = personal.idsucursal "
    sql = sql & "Left Join Anexo_TipoDocumento ax on ax.id= personal.idtipodocumento "
    sql = sql & "WHERE idEmpleado = " & id
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
            mProps.SalarioAnt = IIf(IsNull(!Salario), 0, !Salario)
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
            If Not IsNull(!RutaFoto) Then mProps.RutaFoto = !RutaFoto
            mProps.Tipo = !Tipo
            If Not IsNull(!Discapacidad) Then mProps.Discapacidad = !Discapacidad
            If Not IsNull(!IdTipoRol) Then mProps.IdTipoRol = !IdTipoRol
            If Not IsNull(!IdTipoRol1) Then mProps.IdTipoRol1 = !IdTipoRol1
            If Not IsNull(!IdTipoRol2) Then mProps.IdTipoRol2 = !IdTipoRol2
            If Not IsNull(!IdTipoRol3) Then mProps.IdTipoRol3 = !IdTipoRol3
            If Not IsNull(!IdTipoRol4) Then mProps.IdTipoRol4 = !IdTipoRol4
            If Not IsNull(!IdTipoRol5) Then mProps.IdTipoRol5 = !IdTipoRol5
            If Not IsNull(!IdTipoRol6) Then mProps.IdTipoRol6 = !IdTipoRol6
            If Not IsNull(!IdTipoRol7) Then mProps.IdTipoRol7 = !IdTipoRol7
            If Not IsNull(!IdTipoRol8) Then mProps.IdTipoRol8 = !IdTipoRol8
            If Not IsNull(!IdTipoRol9) Then mProps.IdTipoRol9 = !IdTipoRol9
            If Not IsNull(!IdTipoRol10) Then mProps.IdTipoRol10 = !IdTipoRol10
            If Not IsNull(!UDecimoTercero) Then mProps.UDecimoTercero = !UDecimoTercero
            If Not IsNull(!idUltElemento) Then mProps.idUltElemento = !idUltElemento
            If Not IsNull(!Password) Then mProps.Password = !Password
            If Not IsNull(!CodigoReloj) Then mProps.CodigoReloj = !CodigoReloj
            If Not IsNull(!BandEmpPublico) Then mProps.BandEmpPublico = !BandEmpPublico
            If Not IsNull(!BandAPGerente) Then mProps.BandAPGerente = !BandAPGerente
            mProps.BandLosep = !BandLosep
            If Not IsNull(!IdGrupo1) Then mProps.IdGrupo1 = !IdGrupo1
            If Not IsNull(!CodGrupo1) Then mProps.CodGrupo1 = !CodGrupo1
            If Not IsNull(!IdGrupo2) Then mProps.IdGrupo2 = !IdGrupo2
            If Not IsNull(!CodGrupo2) Then mProps.CodGrupo2 = !CodGrupo2
            If Not IsNull(!IdGrupo3) Then mProps.IdGrupo3 = !IdGrupo3
            If Not IsNull(!CodGrupo3) Then mProps.CodGrupo3 = !CodGrupo3
            If Not IsNull(!IdGrupo4) Then mProps.IdGrupo4 = !IdGrupo4
            If Not IsNull(!CodGrupo4) Then mProps.CodGrupo4 = !CodGrupo4
            If Not IsNull(!CodigoLosep) Then mProps.CodigoLosep = !CodigoLosep
            If Not IsNull(!CodigoLosep) Then mProps.CodigoLosepAnt = !CodigoLosep
            If Not IsNull(!PApellido) Then mProps.Apellido = rs!PApellido
            If Not IsNull(!Pnombre) Then mProps.Nombre = rs!Pnombre
            If Not IsNull(!TipoDiscapacidad) Then mProps.TipoDiscapacidad = rs!TipoDiscapacidad
            If Not IsNull(!Residencia) Then mProps.Residencia = rs!Residencia
            If Not IsNull(!Imposicion) Then mProps.Imposicion = rs!Imposicion
            If Not IsNull(!CedSustituto) Then mProps.CedSustituto = rs!CedSustituto
            mProps.PorcDiscapacidad = rs!PorcDiscapacidad
            mProps.IDSucursal = rs!IDSucursal
            If Not IsNull(!CodSucursal) Then mProps.CodSucursal = rs!CodSucursal
            If Not IsNull(!IdTipoDocumento) Then mProps.IdTipoDocumento = !IdTipoDocumento 'Agregado auc
            If Not IsNull(!codtipoDocumento) Then mProps.codtipoDocumento = !codtipoDocumento 'Agregado auc
            If Not IsNull(!Ciclo) Then mProps.Ciclo = !Ciclo 'Agregado auc
            If Not IsNull(!BandXIII) Then mProps.BandXIII = !BandXIII 'Agregado auc
            If Not IsNull(!BandXIV) Then mProps.BandXIV = !BandXIV 'Agregado auc
            If Not IsNull(!BandGestion) Then mProps.BandGestion = !BandGestion 'Agregado auc
            If Not IsNull(!GP) Then mProps.GP = !GP 'Agregado auc
            If Not IsNull(!GP) Then mProps.GPAnt = !GP 'Agregado auc
            If Not IsNull(!VGP) Then mProps.VGP = !VGP 'Agregado auc
            If Not IsNull(!gpMeses) Then mProps.gpMeses = !gpMeses 'Agregado auc
            If Not IsNull(!FechaFinal) Then mProps.FechaFinalFR = !FechaFinal   'Agregado auc
            If Not IsNull(!FechaFinal) Then mProps.FechaFinalFRAnt = !FechaFinal   'Agregado auc
            If Not IsNull(!BandBenGalp) Then mProps.BandBenGalp = !BandBenGalp   'Agregado auc
            rs.MoveNext
        Loop
            .Close
    End With
    If mProps.IdEmpleado > 0 Then CargaHistorial
    If mProps.IdEmpleado > 0 Then CargaCargasFam
    If mProps.IdEmpleado > 0 Then CargaEstudios
    If mProps.IdEmpleado > 0 Then CargaHistVarios
    If mProps.IdEmpleado > 0 Then CargaHistorialGP
    'If mProps.IdEmpleado > 0 Then CargaHistorialFechaFR
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Grabar(ByVal IdEmpleado As Long)
    Dim sql As String, rs As Recordset
    On Error GoTo Etiq
    
    If Len(mProps.CodigoReloj) > 0 Then
        sql = "SELECT CodigoReloj FROM Personal " & _
          "WHERE (CodigoReloj='" & mProps.CodigoReloj & "') AND " & _
                "(IdEmpleado<>" & mProps.IdEmpleado & ")"
        Set rs = mProps.mobjEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then      'Si ya existe el codigo
            Err.Raise ERR_REPITECODIGO, "Personal_codigoReloj", MSGERR_REPITECODIGO
        End If
        rs.Close
    End If
            sql = "SELECT * FROM Personal WHERE idempleado = " & IdEmpleado
            Set rs = mProps.mobjEmpresa.OpenRecordsetParaEdit(sql)
        
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            sql = "SELECT * FROM Personal WHERE 1=0"
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
        !RutaFoto = Me.RutaFoto
        !Tipo = Me.Tipo
        !Discapacidad = Me.Discapacidad
        !IdTipoRol = Me.IdTipoRol
        !IdTipoRol1 = Me.IdTipoRol1
        !IdTipoRol2 = Me.IdTipoRol2
        !IdTipoRol3 = Me.IdTipoRol3
        !IdTipoRol4 = Me.IdTipoRol4
        !IdTipoRol5 = Me.IdTipoRol5
        !IdTipoRol6 = Me.IdTipoRol6
        !IdTipoRol7 = Me.IdTipoRol7
        !IdTipoRol8 = Me.IdTipoRol8
        !IdTipoRol9 = Me.IdTipoRol9
        !IdTipoRol10 = Me.IdTipoRol10
        !UDecimoTercero = Me.UDecimoTercero
        !idUltElemento = Me.idUltElemento
        !Password = Me.Password
        !CodigoReloj = Me.CodigoReloj
        !BandEmpPublico = mProps.BandEmpPublico
        !BandAPGerente = mProps.BandAPGerente
        !BandLosep = mProps.BandLosep
        !IdGrupo1 = mProps.IdGrupo1
        !IdGrupo2 = mProps.IdGrupo2
        !IdGrupo3 = mProps.IdGrupo3
        !IdGrupo4 = mProps.IdGrupo4
        !CodigoLosep = mProps.CodigoLosep
        !PApellido = Me.Apellido
        !Pnombre = Me.Nombre
        !TipoDiscapacidad = Me.TipoDiscapacidad
        !Residencia = Me.Residencia
        !Imposicion = Me.Imposicion
        !CedSustituto = Me.CedSustituto
        !PorcDiscapacidad = Me.PorcDiscapacidad
        !IDSucursal = Me.IDSucursal
        !IdTipoDocumento = mProps.IdTipoDocumento   'Agregado AUC 2014
        !Ciclo = mProps.Ciclo   'Agregado AUC 2014
        !BandXIII = mProps.BandXIII   'Agregado AUC 2014
        !BandXIV = mProps.BandXIV 'Agregado AUC 2014
        !BandGestion = mProps.BandGestion
        !GP = mProps.GP
        !VGP = mProps.VGP
        !gpMeses = mProps.gpMeses
        !FechaFinal = mProps.FechaFinalFR
        !BandBenGalp = mProps.BandBenGalp
        
        .Update
        .Close
    End With
    GrabarHistorial (IdEmpleado)
    GrabarCargasFam (IdEmpleado)
    GrabarEstudios (IdEmpleado)
    GrabarHistVarios (IdEmpleado)
    GrabarHistorialGP (IdEmpleado)
'    GrabarHistorialFechaFR (IdEmpleado)
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    Me.Empresa.RaiseRefreshEvent REFRESH_PROVCLI
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    
        sql = "DELETE FROM Personal WHERE idEmpleado = " & Me.IdEmpleado
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        Me.Empresa.Coneccion.Execute sql
    mbooEsNuevo = True
    mbooModificado = True
    
End Sub
'
''Para grabar fila por fila
'Public Function AgregarCuentasPersonal() As CuentasPersonal
'    Dim obj As CuentasPersonal
'    Set obj = New CuentasPersonal
'    Set obj.Personal = Me
'    Set AgregarCuentasPersonal = obj
'    Set obj = Nothing
'End Function
'
''Recupera las cuentas que se deben asignar al hacer el Asiento Contable
'Public Function ListaCuentas() As Variant  'Agregado 07/02/2001    Angel
'    Dim sql As String, rs As Recordset
'    sql = "SELECT CodEmpleado,CodElemento,IdCuenta, space(5) As CodCuenta, " & _
'          "space(5) As Cuenta,IdProveedor, space(5) as CodProveedor, space(5) as Proveedor, " & _
'          "FechaGrabado " & _
'          "FROM CuentaPersonal WHERE CodEmpleado = '" & Me.CodEmpleado & "'"
'    #If DAOLIB Then
'        Set rs = Me.Empresa.Database.OpenRecordset(sql, dbSnapshot, dbReadOnly)
'    #Else
'        Set rs = New ADODB.Recordset
'        rs.CursorLocation = adUseClient
'        rs.Open sql, Me.Empresa.Coneccion, adOpenStatic, adLockReadOnly
'    #End If
'    Set ListaCuentas = rs
'    Set rs = Nothing
'End Function
'
''Eliminar todas las cuentas de un empleado para luego grabar de nuevo
'Public Sub EliminaCuentas()
'    Dim sql As String, rs As Recordset
'    #If DAOLIB Then
'        sql = "DELETE * FROM CuentaPersonal WHERE CodEmpleado = '" & Me.CodEmpleado & "'"
'        Me.Empresa.Database.Execute sql, dbFailOnError
'    #Else
'        sql = "DELETE FROM CuentaPersonal WHERE CodEmpleado = '" & Me.CodEmpleado & "'"
'        Set rs = New ADODB.Recordset
'        rs.CursorLocation = adUseClient
'        Me.Empresa.Coneccion.Execute sql
'    #End If
'End Sub

'AUC 08/09/06
Public Property Let ValorQuincena(ByVal value As Currency)
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "Personal.ValorQuincena", MSGERR_INVALIDO
    mProps.ValorQuincena = value
    mbooModificado = True
End Property
'AUC 08/09/06
Public Property Get ValorQuincena() As Currency   'Antes Long. Mofificado por Angel
    ValorQuincena = mProps.ValorQuincena
End Property
Public Property Let Antiguedad(ByVal value As Long)
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "Personal.Antiguedad", MSGERR_INVALIDO
    mProps.Antiguedad = value
    mbooModificado = True
End Property
'AUC 08/09/06
Public Property Get Antiguedad() As Long
    Antiguedad = mProps.Antiguedad
End Property

Public Property Let Varios1(ByVal value As Currency) 'Antes Long. Mofificado por Angel
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "Personal.Varios1", MSGERR_INVALIDO
    mProps.Varios1 = value
    mbooModificado = True
End Property

Public Property Get Varios1() As Currency   'Antes Long. Mofificado por Angel
    Varios1 = mProps.Varios1
End Property

Public Property Let Varios2(ByVal value As Currency) 'Antes Long. Mofificado por Angel
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "Personal.Varios2", MSGERR_INVALIDO
    mProps.Varios2 = value
    mbooModificado = True
End Property

Public Property Get Varios2() As Currency   'Antes Long. Mofificado por Angel
    Varios2 = mProps.Varios2
End Property

Public Property Let Salario(ByVal value As Currency) 'Antes Long. Mofificado por Angel
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "Personal.Salario", MSGERR_INVALIDO
    mProps.Salario = value
    mbooModificado = True
'    mCambiaSueldo = True
End Property

Public Property Get Salario() As Currency   'Antes Long. Mofificado por Angel
    Salario = mProps.Salario
End Property

Public Property Let Varios3(ByVal value As Currency) 'Antes Long. Mofificado por Angel
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "Personal.Varios3", MSGERR_INVALIDO
    mProps.Varios3 = value
    mbooModificado = True
End Property

Public Property Get Varios3() As Currency   'Antes Long. Mofificado por Angel
    Varios3 = mProps.Varios3
End Property
Public Property Let Varios4(ByVal value As Currency) 'Antes Long. Mofificado por Angel
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "Personal.Varios4", MSGERR_INVALIDO
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
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "Personal.Varios5", MSGERR_INVALIDO
    mProps.Varios5 = value
    mbooModificado = True
End Property

Public Property Get Varios6() As Currency   'Antes Long. Mofificado por Angel
    Varios6 = mProps.Varios6
End Property
Public Property Let Varios6(ByVal value As Currency) 'Antes Long. Mofificado por Angel
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "Personal.Varios6", MSGERR_INVALIDO
    mProps.Varios6 = value
    mbooModificado = True
End Property

Public Property Get Varios7() As Currency   'Antes Long. Mofificado por Angel
    Varios7 = mProps.Varios7
End Property
Public Property Let Varios7(ByVal value As Currency) 'Antes Long. Mofificado por Angel
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "Personal.Varios7", MSGERR_INVALIDO
    mProps.Varios7 = value
    mbooModificado = True
End Property
Public Property Get Varios8() As Currency   'Antes Long. Mofificado por Angel
    Varios8 = mProps.Varios8
End Property
Public Property Let Varios8(ByVal value As Currency) 'Antes Long. Mofificado por Angel
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "Personal.Varios8", MSGERR_INVALIDO
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
    If Not IsNumeric(value) Then Err.Raise ERR_INVALIDO, "Personal.Contador", MSGERR_INVALIDO
    mProps.Contador = value
    mbooModificado = True
End Property
Public Property Let FechaPagoSeguro(ByVal value As Date)
    If Not IsDate(value) Then Err.Raise ERR_INVALIDO, "Personal.FechaPagoSeguro", MSGERR_INVALIDO
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
'AUC 22/11/2005
Public Function RecuperaHistorial(ix As Integer) As EmpHistorial
    Set RecuperaHistorial = mcolHistorial.item(ix)
End Function
Public Property Get NumFilasHistorial() As Long
    If mcolHistorial Is Nothing Then
        NumFilasHistorial = 0
    Else
        NumFilasHistorial = mcolHistorial.Count
    End If
End Property


'AUC agrega detalle historial
Public Function AddDetalleHistorial() As Long
    Dim nuevo As EmpHistorial
    
    If mcolHistorial Is Nothing Then CargaHistorial
    
    Set nuevo = New EmpHistorial
    Set nuevo.Personal = Me
    mcolHistorial.Add nuevo
    AddDetalleHistorial = mcolHistorial.Count      'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function

'AUC 23/11/2005
Private Sub CargaHistorial()
    Dim rs As Recordset, sql As String, obj As EmpHistorial
    Dim i As Long
    Set mcolHistorial = New Collection
    sql = "Select eh.idhistorial, eh.idempleado, eh.descripcion,eh.sueldo,eh.fecha,eh.codigoLosep " & _
            " From emphistorial EH " & _
           " inner join  Personal p on p.idempleado=EH.idEmpleado" & _
           " Where eh.idempleado = '" & Me.IdEmpleado & "'"
    Set rs = Me.Empresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddDetalleHistorial
            Set obj = mcolHistorial.item(i)
            obj.IdHistorial = !IdHistorial
            obj.IdEmpleado = !IdEmpleado
            If Not IsNull(!Descripcion) Then obj.Descripcion = !Descripcion
            If Not IsNull(!Sueldo) Then obj.Sueldo = !Sueldo
            If Not IsNull(!fecha) Then obj.fecha = !fecha
            If Not IsNull(!CodigoLosep) Then obj.CodigoLosep = !CodigoLosep
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
'    mbooEsNuevo = False
    mbooModificado = False
End Sub

Private Sub GrabarHistorial(ByRef IdEmpleado As Long)
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj  As EmpHistorial, sql As String, band As Boolean
    'Elimina  todos  los registros
           sql = "Delete From emphistorial where idempleado ='" & IdEmpleado & "'"
           mProps.mobjEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    band = True
    If Not mcolHistorial Is Nothing Then
        sql = "SELECT * FROM  emphistorial WHERE 1=0"
        Set rs1 = mProps.mobjEmpresa.OpenRecordsetParaEdit(sql)
        For i = 1 To mcolHistorial.Count
            Set obj = mcolHistorial.item(i)
                    rs1.AddNew
                    rs1!IdEmpleado = IdEmpleado
                    rs1!Descripcion = obj.Descripcion
                    rs1!Sueldo = obj.Sueldo
                    rs1!fecha = obj.fecha
                    rs1!CodigoLosep = obj.CodigoLosep
                rs1.Update
            band = True
        Next i
    End If
    Set rs1 = Nothing
    Set obj = Nothing
End Sub
Public Sub RemoveHistorial(ByVal ix As Long, Optional obj As EmpHistorial)
    If mcolHistorial Is Nothing Then CargaHistorial
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolHistorial.Count
            If mcolHistorial.item(ix) Is obj Then
                mcolHistorial.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolHistorial.Remove ix
    End If
    'mcolDetalleFamilia.Remove ix
    mbooModificado = True
End Sub

'AUC para insertar campo RutaFoto
Public Property Let RutaFoto(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "IVInventario_RutaFoto", MSGERR_INVALIDO & "Direccion muy larga"
    End If
    mProps.RutaFoto = value
    mbooModificado = True
End Property

Public Property Get RutaFoto() As String
    RutaFoto = Trim$(mProps.RutaFoto)
End Property

Private Sub GrabarCargasFam(ByRef IdEmpleado As Long)
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj  As CargasFam, sql As String, band As Boolean
    'Elimina  todos  los registros
           sql = "Delete From EmpCargaFam where idempleado ='" & IdEmpleado & "'"
           mProps.mobjEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    band = True
    If Not mcolCargas Is Nothing Then
        sql = "SELECT * FROM  EmpCargaFam WHERE 1=0"
        Set rs1 = mProps.mobjEmpresa.OpenRecordsetParaEdit(sql)
        For i = 1 To mcolCargas.Count
            Set obj = mcolCargas.item(i)
                    rs1.AddNew
                    rs1!IdEmpleado = mProps.IdEmpleado
                    rs1!Nombre = obj.Nombre
                    rs1!Sexo = obj.Sexo
                    rs1!FechaNaci = obj.FechaNaci
                    rs1!Observacion = obj.Observacion
                rs1.Update
            band = True
        Next i
    End If
    Set rs1 = Nothing
    Set obj = Nothing
End Sub
Private Sub CargaCargasFam()
    Dim rs As Recordset, sql As String, obj As CargasFam
    Dim i As Long
    Set mcolCargas = New Collection
    sql = "Select ec.idcarga, ec.idempleado, ec.nombre, ec.sexo,ec.fechanaci,ec.observacion " & _
            " From empcargafam EC " & _
           " inner join  Personal p on p.idempleado=Ec.idEmpleado" & _
           " Where ec.idempleado = '" & Me.IdEmpleado & "'"
    Set rs = Me.Empresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddDetalleCarga
            Set obj = mcolCargas.item(i)
            obj.IdCarga = !IdCarga
            obj.IdEmpleado = !IdEmpleado
            If Not IsNull(!Nombre) Then obj.Nombre = !Nombre
            If Not IsNull(!Sexo) Then obj.Sexo = !Sexo
            If Not IsNull(!FechaNaci) Then obj.FechaNaci = !FechaNaci
            If Not IsNull(!Observacion) Then obj.Observacion = !Observacion
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
    mbooModificado = False
End Sub

Public Function AddDetalleCarga() As Long
    Dim nuevo As CargasFam
    If mcolCargas Is Nothing Then CargaCargasFam
    
    Set nuevo = New CargasFam
    Set nuevo.Personal = Me
    mcolCargas.Add nuevo
    AddDetalleCarga = mcolCargas.Count     'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function

Public Function RecuperaCargas(ix As Integer) As CargasFam
    Set RecuperaCargas = mcolCargas.item(ix)
End Function
Public Property Get NumFilasCargas() As Long
    If mcolCargas Is Nothing Then
        NumFilasCargas = 0
    Else
        NumFilasCargas = mcolCargas.Count
    End If
End Property

Public Sub RemoveCargas(ByVal ix As Long, Optional obj As CargasFam)
    If mcolCargas Is Nothing Then CargaCargasFam
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolCargas.Count
            If mcolCargas.item(ix) Is obj Then
                mcolCargas.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolCargas.Remove ix
    End If
    mbooModificado = True
End Sub


Public Property Get CambiaSueldo() As Boolean
    CambiaSueldo = mProps.CambiaSueldo
End Property

Public Property Let CambiaSueldo(ByVal value As Boolean)
     mProps.CambiaSueldo = value
End Property

Public Property Get SalarioAnt() As Currency
'    Dim sql As String
'    Dim rs As Recordset
'    On Error GoTo CapturaError
'    sql = "Select top 1 sueldo from empHistorial where idempleado = " & mProps.IdEmpleado & " order by  idhistorial desc "
'    Set rs = mProps.mobjEmpresa.OpenRecordset(sql)
'    If rs.RecordCount > 0 Then
        SalarioAnt = mProps.SalarioAnt
'    End If
'    Set rs = Nothing
'    Exit Property
'CapturaError:
'    Set rs = Nothing
'    MsgBox Err.Description
    Exit Property
End Property

Public Property Let SalarioAnt(ByVal value As Currency)
     mProps.SalarioAnt = value
End Property


Public Property Get NumFilasEst() As Long
    If mcolEstudios Is Nothing Then
        NumFilasEst = 0
    Else
        NumFilasEst = mcolEstudios.Count
    End If
End Property

Public Function RecuperaEstudios(ix As Integer) As PEstudio
    Set RecuperaEstudios = mcolEstudios.item(ix)
End Function

Public Function AddDetalleEst() As Long
    Dim nuevo As PEstudio
    If mcolEstudios Is Nothing Then CargaEstudios
    
    Set nuevo = New PEstudio
    Set nuevo.Personal = Me
    mcolEstudios.Add nuevo
    AddDetalleEst = mcolEstudios.Count     'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function

Private Sub CargaEstudios()
    Dim rs As Recordset, sql As String, obj As PEstudio
    Dim i As Long
    Set mcolEstudios = New Collection
    sql = "Select est.idEstudio, est.idempleado, est.Descripcion,est.tipo,est.anogrado,est.titulo,est.observacion " & _
            " From empEstudio est " & _
           " inner join  Personal p on p.idempleado=Est.idEmpleado" & _
           " Where est.idempleado = '" & Me.IdEmpleado & "'"
    Set rs = Me.Empresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddDetalleEst
            Set obj = mcolEstudios.item(i)
            obj.IdEstudio = !IdEstudio
            obj.IdEmpleado = !IdEmpleado
            If Not IsNull(!Descripcion) Then obj.Descripcion = !Descripcion
            If Not IsNull(!Tipo) Then obj.Tipo = !Tipo
            If Not IsNull(!anoGrado) Then obj.anoGrado = !anoGrado
            If Not IsNull(!Titulo) Then obj.Titulo = !Titulo
            If Not IsNull(!Observacion) Then obj.Observacion = !Observacion
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
    mbooModificado = False
End Sub


Public Sub RemoveEstudios(ByVal ix As Long, Optional obj As PEstudio)
    If mcolEstudios Is Nothing Then CargaEstudios
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolEstudios.Count
            If mcolEstudios.item(ix) Is obj Then
                mcolEstudios.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolEstudios.Remove ix
    End If
    'mcolDetalleFamilia.Remove ix
    mbooModificado = True
End Sub

Private Sub GrabarEstudios(ByRef IdEmpleado As Long)
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj  As PEstudio, sql As String, band As Boolean
    'Elimina  todos  los registros
           sql = "Delete From Empestudio where idempleado ='" & IdEmpleado & "'"
           mProps.mobjEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    band = True
    If Not mcolEstudios Is Nothing Then
        sql = "SELECT * FROM  EmpEstudio WHERE 1=0"
        Set rs1 = mProps.mobjEmpresa.OpenRecordsetParaEdit(sql)
        For i = 1 To mcolEstudios.Count
            Set obj = mcolEstudios.item(i)
                    rs1.AddNew
                    rs1!IdEmpleado = mProps.IdEmpleado
                    rs1!Descripcion = obj.Descripcion
                    rs1!Tipo = obj.Tipo
                    rs1!anoGrado = obj.anoGrado
                    rs1!Titulo = obj.Titulo
                    rs1!Observacion = obj.Observacion
                rs1.Update
            band = True
        Next i
    End If
    Set rs1 = Nothing
    Set obj = Nothing
End Sub

Public Property Let Tipo(ByVal value As Byte)   '07/09/2000   Angel --->Antes Long
    '0=Soltero, 1=Casado, 2=Divorciado, 3=Viudo
    If (Not IsNumeric(value)) Or (Val(value) < 0 Or Val(value) > 4) Then Err.Raise ERR_INVALIDO, "Personal.EstadoCivil", MSGERR_INVALIDO 'AUC Cambiado de 3 a 4
    mProps.Tipo = value
    mbooModificado = True
End Property

Public Property Get Tipo() As Byte              '07/09/2000   Angel --->Antes Long
    Tipo = mProps.Tipo
End Property

Public Property Let Discapacidad(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Personal_Discapacidad", MSGERR_INVALIDO & "Longitud de cadena no debe ser mayor a 80c"
    End If
    mProps.Discapacidad = value
    mbooModificado = True
End Property

Public Property Get Discapacidad() As String
    Discapacidad = Trim$(mProps.Discapacidad)
End Property

Public Function ListaCuentas() As Variant  'Agregado 07/02/2001    Angel
    Dim sql As String, rs As Recordset
    On Error GoTo CapturaError
    sql = "SELECT PC.Codprovcli,e.CodElemento,IdCuenta, space(5) As CodCuenta, " & _
        "space(5) As Cuenta " & _
        "" & _
        "FROM CuentaPersonal cp INNER JOIN Empleado PC ON PC.IDPROVCLI = cp.idempleado " & _
        " INNER JOIN ELEMENTO e on e.idelemento = cp.idelemento" & _
        " WHERE cp.idEmpleado = '" & Me.IdEmpleado & "' ORDER BY e.orden"
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)

    Set ListaCuentas = rs
    Set rs = Nothing
    Exit Function
CapturaError:
    Set ListaCuentas = rs
    Set rs = Nothing
    MsgBox Err.Description
    Exit Function
End Function

'Eliminar todas las cuentas de un empleado para luego grabar de nuevo
Public Sub EliminaCuentas()
    Dim sql As String, rs As Recordset
    
        sql = "DELETE FROM CuentaPersonal WHERE IdEmpleado = '" & Me.IdEmpleado & "'"
        Me.Empresa.Coneccion.Execute sql
    
End Sub

Public Function AgregarCuentasPersonal() As CuentasPersonal
    Dim obj As CuentasPersonal
    Set obj = New CuentasPersonal
    Set obj.Personal = Me
    Set AgregarCuentasPersonal = obj
    Set obj = Nothing
End Function

Public Function ListaCuentasPre() As Variant  'Agregado 07/02/2001    Angel
    Dim sql As String, rs As Recordset
    On Error GoTo CapturaError
    sql = "SELECT PC.Codprovcli,e.CodElemento,IdCuenta, space(5) As CodCuenta, " & _
        "space(5) As Cuenta " & _
        "" & _
        "FROM CuentaPersonalPre cp INNER JOIN Empleado PC ON PC.IDPROVCLI = cp.idempleado " & _
        " INNER JOIN ELEMENTO e on e.idelemento = cp.idelemento" & _
        " WHERE cp.idEmpleado = '" & Me.IdEmpleado & "' ORDER BY e.orden"
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)

    Set ListaCuentasPre = rs
    Set rs = Nothing
    Exit Function
CapturaError:
    Set ListaCuentasPre = rs
    Set rs = Nothing
    MsgBox Err.Description
    Exit Function
End Function

Public Sub EliminaCuentasPRE()
    Dim sql As String, rs As Recordset
    
        sql = "DELETE FROM CuentaPersonalpre WHERE IdEmpleado = '" & Me.IdEmpleado & "'"
        Me.Empresa.Coneccion.Execute sql
End Sub

Public Function AgregarCuentasPersonalPre() As CuentasPersonalPre
    Dim obj As CuentasPersonalPre
    Set obj = New CuentasPersonalPre
    Set obj.Personal = Me
    Set AgregarCuentasPersonalPre = obj
    Set obj = Nothing
End Function

Public Sub EliminaCuentasLote(ByVal IdEmp As Long)
    Dim sql As String, rs As Recordset
    
        sql = "DELETE FROM CuentaPersonal WHERE IdEmpleado = '" & IdEmp & "'"
        Me.Empresa.Coneccion.Execute sql
    
End Sub

Public Sub EliminaCuentasLotePre(ByVal IdEmp As Long)
    Dim sql As String, rs As Recordset
        sql = "DELETE FROM CuentaPersonalPre WHERE IdEmpleado = '" & IdEmp & "'"
        Me.Empresa.Coneccion.Execute sql
End Sub

'AUC agregado para filtrar igual que el elemento
Public Property Get IdTipoRol1() As Long
    IdTipoRol1 = mProps.IdTipoRol1
End Property
Public Property Let IdTipoRol1(ByVal value As Long)
    mProps.IdTipoRol1 = value
    mbooModificado = True
End Property
'Public Property Let CodTipoRol1(ByVal value As String)
'    Dim sql As String, rs As Recordset
'    'Cuando cambia el valor
'    If value <> mProps.CodTipoRol1 Then
'        If Len(value) > 0 Then
'            'Actualiza IdCuentaActivo también
'            sql = "SELECT IdTipo FROM tiporol WHERE Codigo='" & value & "'"
'            Set rs = mProps.objEmpresa.OpenRecordset(sql)
'            If rs.EOF Then
'                Err.Raise ERR_NOHAYCODIGO, "Elementos_CodTipoRol1", MSGERR_NOHAYCODIGO
'                Exit Property
'            Else
'                IdTipoRol1 = rs!IdTipo
'            End If
'            rs.Close
'        Else
'            IdTipoRol1 = 0
'        End If
'        mProps.CodTipoRol1 = value
'        mbooModificado = True
'    End If
'    Set rs = Nothing
'End Property
'Public Property Get CodTipoRol1() As String
'    CodTipoRol1 = mProps.CodTipoRol1
'End Property
'

Public Property Get IdTipoRol2() As Long
    IdTipoRol2 = mProps.IdTipoRol2
End Property
Public Property Let IdTipoRol2(ByVal value As Long)
    mProps.IdTipoRol2 = value
    mbooModificado = True
End Property
'Public Property Let CodTipoRol2(ByVal value As String)
'    Dim sql As String, rs As Recordset
'    'Cuando cambia el valor
'    If value <> mProps.CodTipoRol2 Then
'        If Len(value) > 0 Then
'            'Actualiza IdCuentaActivo también
'            sql = "SELECT IdTipo FROM tiporol WHERE Codigo='" & value & "'"
'            Set rs = mProps.objEmpresa.OpenRecordset(sql)
'            If rs.EOF Then
'                Err.Raise ERR_NOHAYCODIGO, "Elementos_CodTipoRol2", MSGERR_NOHAYCODIGO
'                Exit Property
'            Else
'                IdTipoRol2 = rs!IdTipo
'            End If
'            rs.Close
'        Else
'            IdTipoRol2 = 0
'        End If
'        mProps.CodTipoRol2 = value
'        mbooModificado = True
'    End If
'    Set rs = Nothing
'End Property
'Public Property Get CodTipoRol2() As String
'    CodTipoRol2 = mProps.CodTipoRol2
'End Property
'

Public Property Get IdTipoRol3() As Long
    IdTipoRol3 = mProps.IdTipoRol3
End Property
Public Property Let IdTipoRol3(ByVal value As Long)
    mProps.IdTipoRol3 = value
    mbooModificado = True
End Property
Public Property Let Codelemento(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.Codelemento Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdElemento FROM Elemento WHERE codElemento='" & value & "'"
            Set rs = mProps.mobjEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Personal_CodElemento", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                idUltElemento = rs!idElemento
            End If
            rs.Close
        Else
            idUltElemento = 0
        End If
        mProps.idUltElemento = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property
Public Property Get Codelemento() As String
Dim sql As String
Dim rs As Recordset
    If Me.idUltElemento <> 0 Then
       sql = "SELECT codElemento FROM Elemento WHERE idElemento=" & Me.idUltElemento
        Set rs = mProps.mobjEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Personal_CodElemento", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                Codelemento = rs!Codelemento
            End If
            rs.Close
    End If
End Property

Public Property Get IdTipoRol() As Long
    IdTipoRol = mProps.IdTipoRol
End Property
Public Property Let IdTipoRol(ByVal value As Long)
    mProps.IdTipoRol = value
    mbooModificado = True
End Property
'Public Property Let CodTipoRol(ByVal value As String)
'    Dim sql As String, rs As Recordset
'    'Cuando cambia el valor
'    If value <> mProps.CodTipoRol Then
'        If Len(value) > 0 Then
'            'Actualiza IdCuentaActivo también
'            sql = "SELECT IdTipo FROM tiporol WHERE Codigo='" & value & "'"
'            Set rs = mProps.objEmpresa.OpenRecordset(sql)
'            If rs.EOF Then
'                Err.Raise ERR_NOHAYCODIGO, "Elementos_CodTipoRol", MSGERR_NOHAYCODIGO
'                Exit Property
'            Else
'                IdTipoRol = rs!IdTipo
'            End If
'            rs.Close
'        Else
'            IdTipoRol = 0
'        End If
'        mProps.CodTipoRol = value
'        mbooModificado = True
'    End If
'    Set rs = Nothing
'End Property
'Public Property Get CodTipoRol() As String
'    CodTipoRol = mProps.CodTipoRol
'End Property


Public Property Get UDecimoTercero() As Currency
    UDecimoTercero = mProps.UDecimoTercero
End Property
Public Property Let UDecimoTercero(ByVal value As Currency)
    mProps.UDecimoTercero = value
    mbooModificado = True
End Property

Public Property Get idUltElemento() As Long
    idUltElemento = mProps.idUltElemento
End Property
Public Property Let idUltElemento(ByVal value As Long)
    mProps.idUltElemento = value
    mbooModificado = True
End Property


Public Sub EliminaCargasFam(ByVal IdCarga As Long)
    Dim sql As String, rs As Recordset
    sql = "DELETE FROM EmpCargaFam WHERE IdCarga = '" & IdCarga & "'"
    Me.Empresa.Execute sql, 1
End Sub
Public Property Get Password() As String
    Password = mProps.Password
End Property

Public Property Let Password(ByVal value As String)
      If Len(value) > 10 Then Err.Raise 9999, "clspersona", "longitud maxima 10 caracteres"
      mProps.Password = value
      mbooModificado = True
End Property
Public Property Get CodigoReloj() As String
    CodigoReloj = mProps.CodigoReloj
End Property
Public Property Let CodigoReloj(ByVal value As String)
      If Len(value) > 20 Then Err.Raise 9999, "clspersona", "longitud maxima 20 caracteres"
      mProps.CodigoReloj = value
      mbooModificado = True
End Property

Public Property Let BandEmpPublico(ByVal value As Boolean)
    mProps.BandEmpPublico = value
    mbooModificado = True
End Property

Public Property Get BandEmpPublico() As Boolean
    BandEmpPublico = mProps.BandEmpPublico
End Property
Public Property Let BandAPGerente(ByVal value As Boolean)
    mProps.BandAPGerente = value
    mbooModificado = True
End Property

Public Property Get BandAPGerente() As Boolean
    BandAPGerente = mProps.BandAPGerente
End Property

Public Property Let BandLosep(ByVal value As Boolean)
    mProps.BandLosep = value
    mbooModificado = True
End Property

Public Property Get BandLosep() As Boolean
    BandLosep = mProps.BandLosep
End Property

Public Property Let IdGrupo1(ByVal value As Long)
    mProps.IdGrupo1 = value
    mbooModificado = True
End Property

Public Property Get IdGrupo1() As Long
    IdGrupo1 = mProps.IdGrupo1
End Property


Public Property Let CodGrupo1(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodGrupo1 Then
        If Len(value) > 0 Then
            'Actualiza IdGrupo1 también
            sql = "SELECT IdGrupo1 FROM PCGrupo1 WHERE CodGrupo1='" & value & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodGrupo1", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdGrupo1 = rs!IdGrupo1
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdGrupo1 = 0
        End If
    End If

    mProps.CodGrupo1 = value
    mbooModificado = True
End Property

Public Property Get CodGrupo1() As String
    CodGrupo1 = Trim$(mProps.CodGrupo1)
End Property



Public Property Let IdGrupo2(ByVal value As Long)
    mProps.IdGrupo2 = value
    mbooModificado = True
End Property

Public Property Get IdGrupo2() As Long
    IdGrupo2 = mProps.IdGrupo2
End Property


Public Property Let CodGrupo2(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodGrupo2 Then
        If Len(value) > 0 Then
            'Actualiza IdGrupo2 también
            sql = "SELECT IdGrupo2 FROM PCGrupo2 WHERE CodGrupo2='" & value & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodGrupo2", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdGrupo2 = rs!IdGrupo2
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdGrupo2 = 0
        End If
    End If

    mProps.CodGrupo2 = value
    mbooModificado = True
End Property

Public Property Get CodGrupo2() As String
    CodGrupo2 = Trim$(mProps.CodGrupo2)
End Property


Public Property Let IdGrupo3(ByVal value As Long)
    mProps.IdGrupo3 = value
    mbooModificado = True
End Property
Public Property Get IdGrupo3() As Long
    IdGrupo3 = mProps.IdGrupo3
End Property


Public Property Let CodGrupo3(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodGrupo3 Then
        If Len(value) > 0 Then
            'Actualiza IdGrupo3 también
            sql = "SELECT IdGrupo3 FROM PCGrupo3 WHERE CodGrupo3='" & value & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodGrupo3", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdGrupo3 = rs!IdGrupo3
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdGrupo3 = 0
        End If
    End If

    mProps.CodGrupo3 = value
    mbooModificado = True
End Property
Public Property Get CodGrupo3() As String
    CodGrupo3 = Trim$(mProps.CodGrupo3)
End Property

Public Property Get IdGrupo4() As Long
    IdGrupo4 = mProps.IdGrupo4
End Property
Public Property Let IdGrupo4(ByVal value As Long)
    mProps.IdGrupo4 = value
    mbooModificado = True
End Property

Public Property Let CodGrupo4(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodGrupo4 Then
        If Len(value) > 0 Then
            'Actualiza IdGrupo3 también
            sql = "SELECT IdGrupo4 FROM PCGrupo4 WHERE CodGrupo4='" & value & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodGrupo4", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdGrupo4 = rs!IdGrupo4
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdGrupo4 = 0
        End If
    End If
    mProps.CodGrupo4 = value
    mbooModificado = True
End Property

Public Property Get CodGrupo4() As String
    CodGrupo4 = Trim$(mProps.CodGrupo4)
End Property

Public Sub GenerarCodigo()
    Dim sql As String, num As Long, cod As String, rs As Recordset
    Dim pre As String, codg As String, v1 As Variant, v2 As Variant
    Dim Orden As String, size As String, mascara As String, i As Integer
        
    Orden = Me.Empresa.GNOpcion.ObtenerValor("GenerarEmpCodigoOrden")
    size = Me.Empresa.GNOpcion.ObtenerValor("GenerarEmpCodigotamano")
    mascara = Me.Empresa.GNOpcion.ObtenerValor("GenerarEmpCodigoMascara")
    
    codg = ""
    v1 = Split(Orden, ",")
    v2 = Split(size, ",")
    For i = LBound(v1, 1) To UBound(v1, 1)
        Select Case v1(i)
            Case "1"
                codg = codg & Trim$(Mid$(mProps.CodGrupo1, 1, v2(i)))
            Case "2"
                codg = codg & Trim$(Mid$(mProps.CodGrupo2, 1, v2(i)))
            Case "3"
                codg = codg & Trim$(Mid$(mProps.CodGrupo3, 1, v2(i)))
            Case "4"
                codg = codg & Trim$(Mid$(mProps.CodGrupo4, 1, v2(i)))
        End Select
        'codg = codg & Trim$(Mid$(mProps.CodGrupo(v1(i)), 1, v2(i)))
    Next i
    
    pre = Me.Empresa.GNOpcion.PrefijoCodEmpLosep
    pre = Trim$(pre) & Trim$(codg)
    cod = Me.CodigoLosep
#If DAOLIB Then
    sql = "SELECT Max(VAL(CodigoLosep)) AS CodMax FROM Personal " & _
          "WHERE Codigolosep LIKE '" & pre & "*'"
#Else
    'sql = "SELECT (MAX(SUBSTRING(CodigoLosep," & Len(pre) + 1 & ",LEN(codigoLosep)))) As CodMax " & _
          "FROM Personal "
    VerificaExistenciaTabla 0
    sql = "SELECT (MAX(SUBSTRING(CodigoLosep," & Len(pre) + 1 & ",LEN(CodigoLosep)))) As CodMax " & _
          "INTO tmp0  FROM pERSONAL " & _
          "WHERE CodigoLosep LIKE '" & pre & "[0-9][0-9]%'"
    sql = sql & " UNION ALL"
    sql = sql & " SELECT (MAX(SUBSTRING(CodigoLosep," & Len(pre) + 1 & ",LEN(CodigoLosep)))) As CodMax " & _
          "FROM EMPHistorial  " & _
          "WHERE CodigoLosep LIKE '" & pre & "[0-9][0-9]%'"
         gobjMain.EmpresaActual.EjecutarSQL sql, 1
         
    sql = "SELECT max(codmax) as codmax FROM tmp0"
    
#End If

    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not rs.EOF Then
        If Not IsNull(rs!CodMax) Then num = Val(rs!CodMax)
    End If
    rs.Close
    Set rs = Nothing
    num = num + 1
    cod = pre & Format(num, mascara) 'cambiar por plantilla
    'REVISO QUE EL CODIGO NO EXISTA NI EL PERSONAL
    sql = "Select pc.nombre,p.codigolosep from  personal p inner join pcprovcli pc on pc.idprovcli = p.idempleado  "
    sql = sql & "Where pc.bandempleado = 1 And codigolosep = '" & cod & "' "
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        MsgBox "Este codigo ya esta asignado ha " & rs!Nombre
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    'REVISO QUE EL CODIGO NO EXISTA NI EL PERSONAL
    sql = "Select pc.nombre,h.codigolosep from  empHistorial h inner join pcprovcli pc on pc.idprovcli = H.idempleado  "
    sql = sql & "Where pc.bandempleado = 1 And codigolosep = '" & cod & "' "
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        MsgBox "Este codigo ya esta asignado ha " & rs!Nombre
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    mProps.CodigoLosep = cod
'    GenerarCodigo = cod
    Set rs = Nothing
End Sub

Public Property Get CodigoLosep() As String
    CodigoLosep = mProps.CodigoLosep
End Property

Public Property Let CodigoLosep(ByVal value As String)
      If Len(value) > 20 Then Err.Raise 9999, "clspersona", "longitud maxima 20 caracteres"
      mProps.CodigoLosep = value
      mbooModificado = True
End Property

Public Property Get CambiaCodigoLosep() As Boolean
    CambiaCodigoLosep = mProps.CambioCodigoLosep
End Property

Public Property Let CambiaCodigoLosep(ByVal value As Boolean)
     mProps.CambioCodigoLosep = value
End Property

Public Property Get CodigoLosepAnt() As String
        CodigoLosepAnt = mProps.CodigoLosepAnt
    Exit Property
End Property

Public Property Let CodigoLosepAnt(ByVal value As String)
     mProps.CodigoLosepAnt = value
End Property


Public Property Let Apellido(ByVal value As String)
    mProps.Apellido = value
    mbooModificado = True
End Property

Public Property Get Apellido() As String
    Apellido = mProps.Apellido
End Property
Public Property Let Nombre(ByVal value As String)
    mProps.Nombre = value
    mbooModificado = True
End Property

Public Property Get Nombre() As String
    Nombre = mProps.Nombre
End Property

Public Property Let TipoDiscapacidad(ByVal value As String)
    mProps.TipoDiscapacidad = value
    mbooModificado = True
End Property

Public Property Get TipoDiscapacidad() As String
    TipoDiscapacidad = mProps.TipoDiscapacidad
End Property

Public Property Let Residencia(ByVal value As String)
    mProps.Residencia = value
    mbooModificado = True
End Property

Public Property Get Residencia() As String
    Residencia = mProps.Residencia
End Property


Public Property Let Imposicion(ByVal value As String)
    mProps.Imposicion = value
    mbooModificado = True
End Property

Public Property Get Imposicion() As String
    Imposicion = mProps.Imposicion
End Property

Public Property Let CedSustituto(ByVal value As String)
    mProps.CedSustituto = value
    mbooModificado = True
End Property

Public Property Get CedSustituto() As String
    CedSustituto = mProps.CedSustituto
End Property

Public Property Let PorcDiscapacidad(ByVal value As Integer)
    mProps.PorcDiscapacidad = value
    mbooModificado = True
End Property

Public Property Get PorcDiscapacidad() As Integer
    PorcDiscapacidad = mProps.PorcDiscapacidad
End Property


Public Property Let CodSucursal(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodSucursal Then
        If Len(value) > 0 Then
            'Actualiza IdResponsable también
            sql = "SELECT IdSucursal FROM GNSucursal WHERE CodSucursal='" & value & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Empleado_CodSucursal", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IDSucursal = rs!IDSucursal
            End If
            rs.Close
        Else
            IDSucursal = 0
        End If
    End If
    Set rs = Nothing
    mProps.CodSucursal = value
    mbooModificado = True
End Property

Public Property Get CodSucursal() As String
    CodSucursal = mProps.CodSucursal
End Property


Public Property Let IDSucursal(ByVal value As Integer)
    mProps.IDSucursal = value
    mbooModificado = True
End Property

Public Property Get IDSucursal() As Integer
    IDSucursal = mProps.IDSucursal
End Property

Public Property Get IdTipoDocumento() As Long
    IdTipoDocumento = mProps.IdTipoDocumento
End Property

Public Property Let IdTipoDocumento(ByVal value As Long)
    mProps.IdTipoDocumento = value
    mbooModificado = True
End Property


Public Property Let codtipoDocumento(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.codtipoDocumento Then
        If Len(value) > 0 Then
            'Actualiza IdTipoDocumento también
            sql = "SELECT Id FROM Anexo_TipoDocumento WHERE CodTipoDocumento ='" & value & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodTipoDocumento", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdTipoDocumento = rs!id
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdTipoDocumento = 0
        End If
    End If

    mProps.codtipoDocumento = value
    
    
    mbooModificado = True
End Property

Public Property Get codtipoDocumento() As String
    codtipoDocumento = Trim$(mProps.codtipoDocumento)
End Property

Public Property Get IdTipoRol4() As Long
    IdTipoRol4 = mProps.IdTipoRol4
End Property
Public Property Let IdTipoRol4(ByVal value As Long)
    mProps.IdTipoRol4 = value
    mbooModificado = True
End Property

Public Property Get IdTipoRol5() As Long
    IdTipoRol5 = mProps.IdTipoRol5
End Property
Public Property Let IdTipoRol5(ByVal value As Long)
    mProps.IdTipoRol5 = value
    mbooModificado = True
End Property

Public Property Get IdTipoRol6() As Long
    IdTipoRol6 = mProps.IdTipoRol6
End Property
Public Property Let IdTipoRol6(ByVal value As Long)
    mProps.IdTipoRol6 = value
    mbooModificado = True
End Property

Public Property Get IdTipoRol7() As Long
    IdTipoRol7 = mProps.IdTipoRol7
End Property
Public Property Let IdTipoRol7(ByVal value As Long)
    mProps.IdTipoRol7 = value
    mbooModificado = True
End Property

Public Property Get IdTipoRol8() As Long
    IdTipoRol8 = mProps.IdTipoRol8
End Property
Public Property Let IdTipoRol8(ByVal value As Long)
    mProps.IdTipoRol8 = value
    mbooModificado = True
End Property


Public Property Get IdTipoRol9() As Long
    IdTipoRol9 = mProps.IdTipoRol9
End Property
Public Property Let IdTipoRol9(ByVal value As Long)
    mProps.IdTipoRol9 = value
    mbooModificado = True
End Property


Public Property Get IdTipoRol10() As Long
    IdTipoRol10 = mProps.IdTipoRol10
End Property
Public Property Let IdTipoRol10(ByVal value As Long)
    mProps.IdTipoRol10 = value
    mbooModificado = True
End Property

Public Property Let Ciclo(ByVal value As String)
    mProps.Ciclo = value
    mbooModificado = True
End Property

Public Property Get Ciclo() As String
    Ciclo = mProps.Ciclo
End Property

Public Function AddDetalleHistVarios() As Long
    Dim nuevo As EmpHistVarios
    
    If mcolHistVarios Is Nothing Then CargaHistVarios
    
    Set nuevo = New EmpHistVarios
    Set nuevo.Personal = Me
    mcolHistVarios.Add nuevo
    AddDetalleHistVarios = mcolHistVarios.Count      'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function

Private Sub CargaHistVarios()
    Dim rs As Recordset, sql As String, obj As EmpHistVarios
    Dim i As Long
    Set mcolHistVarios = New Collection
    sql = "Select eh.idhistorial,eh.idEmpleado, eh.descripcion,eh.fechaGrabado " & _
           " From emphistVarios EH " & _
           " " & _
           " Where eh.idEmpleado = " & mProps.IdEmpleado
    Set rs = mProps.mobjEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddDetalleHistVarios
            Set obj = mcolHistVarios.item(i)
            obj.IdHistorial = !IdHistorial
            obj.IdEmpleado = !IdEmpleado
            If Not IsNull(!Descripcion) Then obj.Descripcion = !Descripcion
            If Not IsNull(!fechagrabado) Then obj.fechagrabado = !fechagrabado
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
'    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Function RecuperaHistVarios(ix As Integer) As EmpHistVarios
    Set RecuperaHistVarios = mcolHistVarios.item(ix)
End Function
Public Property Get NumFilasHistVarios() As Long
    If mcolHistVarios Is Nothing Then
        NumFilasHistVarios = 0
    Else
        NumFilasHistVarios = mcolHistVarios.Count
    End If
End Property

Public Sub RemoveHistVarios(ByVal ix As Long, Optional obj As EmpHistVarios)
    If mcolHistVarios Is Nothing Then CargaHistVarios
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolHistVarios.Count
            If mcolHistVarios.item(ix) Is obj Then
                mcolHistVarios.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolHistVarios.Remove ix
    End If
    
    mbooModificado = True
End Sub

Private Sub GrabarHistVarios(ByRef IdEmpleado As Long)
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj  As EmpHistVarios, sql As String, band As Boolean
    'Elimina  todos  los registros
           sql = "Delete From emphistvarios where idempleado ='" & IdEmpleado & "'"
           mProps.mobjEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    band = True
    If Not mcolHistVarios Is Nothing Then
        sql = "SELECT * FROM  emphistvarios WHERE 1=0"
        Set rs1 = mProps.mobjEmpresa.OpenRecordsetParaEdit(sql)
        For i = 1 To mcolHistVarios.Count
            Set obj = mcolHistVarios.item(i)
                    rs1.AddNew
                    rs1!IdEmpleado = IdEmpleado
                    rs1!Descripcion = obj.Descripcion
                    rs1!fechagrabado = obj.fechagrabado
                rs1.Update
            band = True
        Next i
    End If
    Set rs1 = Nothing
    Set obj = Nothing
End Sub

Public Property Let BandXIII(ByVal value As Boolean)
    mProps.BandXIII = value
    mbooModificado = True
End Property

Public Property Get BandXIII() As Boolean
    BandXIII = mProps.BandXIII
End Property

Public Property Let BandXIV(ByVal value As Boolean)
    mProps.BandXIV = value
    mbooModificado = True
End Property

Public Property Get BandXIV() As Boolean
    BandXIV = mProps.BandXIV
End Property

Public Property Let BandGestion(ByVal value As Boolean)
    mProps.BandGestion = value
    mbooModificado = True
End Property

Public Property Get BandGestion() As Boolean
    BandGestion = mProps.BandGestion
End Property
Public Property Let GP(ByVal value As Currency)
    mProps.GP = value
    mbooModificado = True
End Property

Public Property Get GP() As Currency
    GP = mProps.GP
End Property
Public Property Let VGP(ByVal value As Currency)
    mProps.VGP = value
    mbooModificado = True
End Property

Public Property Get VGP() As Currency
    VGP = mProps.VGP
End Property

Public Property Let gpMeses(ByVal value As String)
    If Len(value) > 30 Then Err.Raise ERR_INVALIDO, "Personal.Meses", MSGERR_INVALIDO
    mProps.gpMeses = value
    mbooModificado = True
End Property

Public Property Get gpMeses() As String
    gpMeses = mProps.gpMeses
End Property
'----------------

Public Function RecuperaHistorialGP(ix As Integer) As EmpHistorialGP
    Set RecuperaHistorialGP = mcolHistorialGP.item(ix)
End Function
Public Property Get NumFilasHistorialGP() As Long
    If mcolHistorialGP Is Nothing Then
        NumFilasHistorialGP = 0
    Else
        NumFilasHistorialGP = mcolHistorialGP.Count
    End If
End Property


'AUC agrega detalle historial
Public Function AddDetalleHistorialGP() As Long
    Dim nuevo As EmpHistorialGP
    
    If mcolHistorialGP Is Nothing Then CargaHistorialGP
    
    Set nuevo = New EmpHistorialGP
    Set nuevo.Personal = Me
    mcolHistorialGP.Add nuevo
    AddDetalleHistorialGP = mcolHistorialGP.Count      'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function

'AUC 23/11/2005
Private Sub CargaHistorialGP()
    Dim rs As Recordset, sql As String, obj As EmpHistorialGP
    Dim i As Long
    Set mcolHistorialGP = New Collection
    sql = "Select eh.idhistorialgp, eh.idempleado, eh.GP,eh.fecha" & _
            " From emphistorialGP EH " & _
           " inner join  Personal p on p.idempleado=EH.idEmpleado" & _
           " Where eh.idempleado = '" & Me.IdEmpleado & "'"
    Set rs = Me.Empresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddDetalleHistorialGP
            Set obj = mcolHistorialGP.item(i)
            obj.IdHistorialGP = !IdHistorialGP
            obj.IdEmpleado = !IdEmpleado
            If Not IsNull(!GP) Then obj.GP = !GP
            If Not IsNull(!fecha) Then obj.fecha = !fecha
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
'    mbooEsNuevo = False
    mbooModificado = False
End Sub

Private Sub GrabarHistorialGP(ByRef IdEmpleado As Long)
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj  As EmpHistorialGP, sql As String, band As Boolean
    'Elimina  todos  los registros
           sql = "Delete From emphistorialGP where idempleado ='" & IdEmpleado & "'"
           mProps.mobjEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    band = True
    If Not mcolHistorialGP Is Nothing Then
        sql = "SELECT * FROM  emphistorialGP WHERE 1=0"
        Set rs1 = mProps.mobjEmpresa.OpenRecordsetParaEdit(sql)
        For i = 1 To mcolHistorialGP.Count
            Set obj = mcolHistorialGP.item(i)
                    rs1.AddNew
                    rs1!IdEmpleado = IdEmpleado
                    rs1!GP = obj.GP
                    rs1!fecha = obj.fecha
                rs1.Update
            band = True
        Next i
    End If
    Set rs1 = Nothing
    Set obj = Nothing
End Sub
Public Sub RemoveHistorialGP(ByVal ix As Long, Optional obj As EmpHistorialGP)
    If mcolHistorialGP Is Nothing Then CargaHistorialGP
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolHistorialGP.Count
            If mcolHistorialGP.item(ix) Is obj Then
                mcolHistorialGP.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolHistorialGP.Remove ix
    End If
    'mcolDetalleFamilia.Remove ix
    mbooModificado = True
End Sub
Public Property Get GPAnt() As Currency
        GPAnt = mProps.GPAnt
    Exit Property
End Property

Public Property Let GPAnt(ByVal value As Currency)
     mProps.GPAnt = value
End Property

Public Property Get CambiaGP() As Boolean
    CambiaGP = mProps.CambiaGP
End Property

Public Property Let CambiaGP(ByVal value As Boolean)
     mProps.CambiaGP = value
End Property

Public Property Let FechaFinalFR(ByVal value As Date)
    mProps.FechaFinalFR = value
    mbooModificado = True
End Property
Public Property Get FechaFinalFR() As Date
    FechaFinalFR = mProps.FechaFinalFR
End Property
Public Property Get CambiaFFR() As Boolean
    CambiaFFR = mProps.CambiaFFR
End Property
Public Property Let CambiaFFR(ByVal value As Boolean)
     mProps.CambiaFFR = value
End Property
Public Property Let FechaFinalFRAnt(ByVal value As Date)
    mProps.FechaFinalFRAnt = value
    mbooModificado = True
End Property
Public Property Get FechaFinalFRAnt() As Date
    FechaFinalFRAnt = mProps.FechaFinalFRAnt
End Property

Public Function AddDetalleHistorialFFR() As Long
    Dim nuevo As EmpHistorialFechaFR
    If mcolHistorialFechaFR Is Nothing Then CargaHistorialFechaFR
    Set nuevo = New EmpHistorialFechaFR
    Set nuevo.Personal = Me
    mcolHistorialFechaFR.Add nuevo
    AddDetalleHistorialFFR = mcolHistorialFechaFR.Count      'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function

Private Sub CargaHistorialFechaFR()
    Dim rs As Recordset, sql As String, obj As EmpHistorialFechaFR
    Dim i As Long
    Set mcolHistorialFechaFR = New Collection
    sql = "Select eh.idhistfechafr, eh.idempleado, eh.fechaOld,eh.descripcion,eh.fechagrabado" & _
            " From emphistorialFechaFR EH " & _
           " inner join  Personal p on p.idempleado=EH.idEmpleado" & _
           " Where eh.idempleado = '" & Me.IdEmpleado & "'"
    Set rs = Me.Empresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddDetalleHistorialFFR
            Set obj = mcolHistorialFechaFR.item(i)
            obj.IdHistorialFFR = !IdHistfechafr
            obj.IdEmpleado = !IdEmpleado
            If Not IsNull(!fechaOld) Then obj.fechaOld = !fechaOld
            If Not IsNull(!Descripcion) Then obj.Descripcion = !Descripcion
            If Not IsNull(!fechagrabado) Then obj.fechagrabado = !fechagrabado
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
'    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Function RecuperaHistorialFechaFR(ix As Integer) As EmpHistorialFechaFR
    Set RecuperaHistorialFechaFR = mcolHistorialFechaFR.item(ix)
End Function

Public Property Get NumFilasHistorialFechaFR() As Long
    If mcolHistorialFechaFR Is Nothing Then
        NumFilasHistorialFechaFR = 0
    Else
        NumFilasHistorialFechaFR = mcolHistorialFechaFR.Count
    End If
End Property

Private Sub GrabarHistorialFechaFR(ByRef IdEmpleado As Long)
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj  As EmpHistorialFechaFR, sql As String, band As Boolean
    'Elimina  todos  los registros
           sql = "Delete From emphistorialfechafr where idempleado ='" & IdEmpleado & "'"
           mProps.mobjEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    band = True
    If Not mcolHistorialFechaFR Is Nothing Then
        sql = "SELECT * FROM  emphistorialfechafr WHERE 1=0"
        Set rs1 = mProps.mobjEmpresa.OpenRecordsetParaEdit(sql)
        For i = 1 To mcolHistorialFechaFR.Count
            Set obj = mcolHistorialFechaFR.item(i)
                    rs1.AddNew
                    rs1!IdEmpleado = IdEmpleado
                    rs1!fechaOld = obj.fechaOld
                    rs1!Descripcion = obj.Descripcion
                    rs1!fechagrabado = obj.fechagrabado
                rs1.Update
            band = True
        Next i
    End If
    Set rs1 = Nothing
    Set obj = Nothing
End Sub
'beneficio galapagos
Public Property Get BandBenGalp() As Boolean
    BandBenGalp = mProps.BandBenGalp
End Property
Public Property Let BandBenGalp(ByVal value As Boolean)
     mProps.BandBenGalp = value
End Property

