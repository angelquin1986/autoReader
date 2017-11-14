VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "fichaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    Id As Long
    idficha As Long
    Historia As String
    IdCliente As Long
    Sexo As Byte
    Cedula As String
    idCentro As Long
    Piso As Integer
    Nombre As String
    FechaIngreso As Date
    Hora As Date
    Edad As Integer
    Procedencia As String
    IdCiudad As Long
    Direccion As String
    Telefono As String
    Ocupacion As String
    Religion As String
    idmedico As Long
    Sintomas As String
    ExamenFisico As String
    Peso As String
    Medicamentos As String
    CP As String
    Pulso As String
    Temperatura As String
    Respiracion As String
    CodCentro As String
    CodCliente As String
    CodMedico As String
    objFicha As Ficha
    
    CodOrgano As String
    DescOrgano As String
    CPOrgano As Boolean
    SPOrgano As Boolean
    
    CodSigno As String
    DescSigno As String
    
    CodExamen As String
    DescExamen As String
    
    CPExamen As Boolean
    SPExamen As Boolean
    
    CodDiag As String
    DescDiag As String
    DescPlan As String
    BandAlta As Boolean
    Modificar As Boolean
    
    
End Type
Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean
Event Refresh(Tipo As Integer)

Public Property Let idficha(ByVal value As Long)
    mProps.idficha = value
    mbooModificado = True
End Property

Public Property Get idficha() As Long
    idficha = mProps.idficha
End Property

Friend Property Set objFicha(ByVal obj As Ficha)
    Set mProps.objFicha = obj
End Property

Public Property Get objFicha() As Ficha
    Set objFicha = mProps.objFicha
End Property

Public Property Let Sexo(ByVal value As Byte)
    '0=Femenino, 1=Masculino
    If (Not IsNumeric(value)) Or (Val(value) < 0 Or Val(value) > 1) Then Err.Raise ERR_INVALIDO, "PCEmpleado.Sexo", MSGERR_INVALIDO
    mProps.Sexo = value
    mbooModificado = True
End Property

Public Property Get Sexo() As Byte
    Sexo = mProps.Sexo
End Property

Public Property Let IdCliente(ByVal value As Long)
    mProps.IdCliente = value
    mbooModificado = True
End Property

Public Property Get IdCliente() As Long
    IdCliente = mProps.IdCliente
End Property

Public Property Let Cedula(ByVal value As String)
    If Len(value) > 10 Then Err.Raise ERR_INVALIDO, "Ficha.Cedula", MSGERR_INVALIDO
    mProps.Cedula = value
    mbooModificado = True
End Property

Public Property Get Cedula() As String
    Cedula = mProps.Cedula
End Property

Public Property Let idCentro(ByVal value As Long)
    mProps.idCentro = value
    mbooModificado = True
End Property

Public Property Get idCentro() As Long
    idCentro = mProps.idCentro
End Property

Public Property Get Piso() As Integer
    Piso = mProps.Piso
End Property
Public Property Let Piso(ByVal value As Integer)
If value > 5 Then Err.Raise ERR_INVALIDO, "Ficha.Piso", "Rango no permitido. max:5 pisos": value = 5
    mProps.Piso = value
    mbooModificado = True
End Property

Public Property Let Nombre(ByVal value As String)
    mProps.Nombre = value
    mbooModificado = True
End Property

Public Property Get Nombre() As String
    Nombre = mProps.Nombre
End Property

Public Property Let FechaIngreso(ByVal value As Date)
    If Not IsDate(value) Then Err.Raise ERR_INVALIDO, "Ficha.FechaIngreso", MSGERR_INVALIDO
    mProps.FechaIngreso = value
    mbooModificado = True
End Property

Public Property Get FechaIngreso() As Date
    FechaIngreso = mProps.FechaIngreso
End Property

Public Property Let Hora(ByVal value As Date)
    If Not IsDate(value) Then Err.Raise ERR_INVALIDO, "Ficha.hora", MSGERR_INVALIDO
    mProps.Hora = value
    mbooModificado = True
End Property

Public Property Get Hora() As Date
    Hora = mProps.Hora
End Property

Public Property Get Edad() As Integer
    Edad = mProps.Edad
End Property
Public Property Let Edad(ByVal value As Integer)
If value > 130 Then Err.Raise ERR_INVALIDO, "Ficha.Edad", "Rango de Edad no permitido. max:130": value = 130
    mProps.Edad = value
    mbooModificado = True
End Property

Public Property Let Procedencia(ByVal value As String)
    mProps.Procedencia = value
    mbooModificado = True
End Property

Public Property Get Procedencia() As String
    Procedencia = mProps.Procedencia
End Property
Public Property Let IdCiudad(ByVal value As Long)
    mProps.IdCiudad = value
    mbooModificado = True
End Property

Public Property Get IdCiudad() As Long
    IdCiudad = mProps.IdCiudad
End Property

Public Property Let Direccion(ByVal value As String)
    mProps.Direccion = value
    mbooModificado = True
End Property

Public Property Get Direccion() As String
    Direccion = mProps.Direccion
End Property

Public Property Let Telefono(ByVal value As String)
    mProps.Telefono = value
    mbooModificado = True
End Property

Public Property Get Telefono() As String
    Telefono = mProps.Telefono
End Property

Public Property Let Ocupacion(ByVal value As String)
    mProps.Ocupacion = value
    mbooModificado = True
End Property

Public Property Get Ocupacion() As String
    Ocupacion = mProps.Ocupacion
End Property

Public Property Let Religion(ByVal value As String)
    mProps.Religion = value
    mbooModificado = True
End Property

Public Property Get Religion() As String
    Religion = mProps.Religion
End Property

Public Property Let idmedico(ByVal value As Long)
    mProps.idmedico = value
    mbooModificado = True
End Property

Public Property Get idmedico() As Long
    idmedico = mProps.idmedico
End Property

Public Property Let Sintomas(ByVal value As String)
    mProps.Sintomas = value
    mbooModificado = True
End Property

Public Property Get Sintomas() As String
    Sintomas = mProps.Sintomas
End Property

Public Property Let ExamenFisico(ByVal value As String)
    mProps.ExamenFisico = value
    mbooModificado = True
End Property

Public Property Get ExamenFisico() As String
    ExamenFisico = mProps.ExamenFisico
End Property

Public Property Let DX(ByVal value As String)
'    mProps.DX = Value
'    mbooModificado = True
End Property

Public Property Get DX() As String
    'DX = mProps.DX
End Property

Public Property Let Antesedentes(ByVal value As String)
'    mProps.Antesedentes = Value
'    mbooModificado = True
End Property

Public Property Get Antesedentes() As String
    'Antesedentes = mProps.Antesedentes
End Property

Public Property Let Peso(ByVal value As String)
    mProps.Peso = value
    mbooModificado = True
End Property

Public Property Get Peso() As String
    Peso = mProps.Peso
End Property

Public Property Let TipoSangre(ByVal value As String)
    'mProps.TipoSangre = Value
    'mbooModificado = True
End Property

Public Property Get TipoSangre() As String
'    TipoSangre = mProps.TipoSangre
End Property

Public Property Let Infancia(ByVal value As String)
'    mProps.Infancia = Value
'    mbooModificado = True
End Property

Public Property Get Infancia() As String
    'Infancia = mProps.Infancia
End Property

Public Property Let Pubertad(ByVal value As String)
'    mProps.Pubertad = Value
'    mbooModificado = True
End Property

Public Property Get Pubertad() As String
    'Pubertad = mProps.Pubertad
End Property

Public Property Let Adulto(ByVal value As String)
'    mProps.Adulto = Value
'    mbooModificado = True
End Property

Public Property Get Adulto() As String
    'Adulto = mProps.Adulto
End Property

Public Property Let Habito(ByVal value As String)
'    mProps.Habito = Value
'    mbooModificado = True
End Property

Public Property Get Habito() As String
'    Habito = mProps.Habito
End Property
Public Property Let Alergia(ByVal value As String)
'    mProps.Alergia = Value
'    mbooModificado = True
End Property

Public Property Get Alergia() As String
    'Alergia = mProps.Alergia
End Property

Public Property Let Medicamentos(ByVal value As String)
    mProps.Medicamentos = value
    mbooModificado = True
End Property

Public Property Get Medicamentos() As String
    Medicamentos = mProps.Medicamentos
End Property

Public Property Let Corticostiroides(ByVal value As String)
'    mProps.Corticostiroides = Value
'    mbooModificado = True
End Property

Public Property Get Corticostiroides() As String
    'Corticostiroides = mProps.Corticostiroides
End Property

Public Property Let DerivadosFE(ByVal value As String)
'    mProps.DerivadosFE = Value
'    mbooModificado = True
End Property

Public Property Get DerivadosFE() As String
    'DerivadosFE = mProps.DerivadosFE
End Property

Public Property Let DerivadosRA(ByVal value As String)
'    mProps.DerivadosRA = Value
'    mbooModificado = True
End Property

Public Property Get DerivadosRA() As String
    'DerivadosRA = mProps.DerivadosRA
End Property

Public Property Let Antidiabeticos(ByVal value As String)
'    mProps.Antidiabeticos = Value
'    mbooModificado = True
End Property

Public Property Get Antidiabeticos() As String
    'Antidiabeticos = mProps.Antidiabeticos
End Property

Public Property Let CP(ByVal value As String)
    mProps.CP = value
    mbooModificado = True
End Property

Public Property Get CP() As String
    CP = mProps.CP
End Property

Public Property Let Pulso(ByVal value As String)
    mProps.Pulso = value
    mbooModificado = True
End Property

Public Property Get Pulso() As String
    Pulso = mProps.Pulso
End Property



Public Property Let Antibioticos(ByVal value As String)
'    mProps.Antibioticos = Value
'    mbooModificado = True
End Property

Public Property Get Antibioticos() As String
    'Antibioticos = mProps.Antibioticos
End Property


Public Property Let Temperatura(ByVal value As String)
    mProps.Temperatura = value
    mbooModificado = True
End Property

Public Property Get Temperatura() As String
    Temperatura = mProps.Temperatura
End Property

Public Property Let Respiracion(ByVal value As String)
    mProps.Respiracion = value
    mbooModificado = True
End Property

Public Property Get Respiracion() As String
    Respiracion = mProps.Respiracion
End Property

Public Property Let Pupilas(ByVal value As String)
'    mProps.Pupilas = Value
'    mbooModificado = True
End Property

Public Property Get Pupilas() As String
    'Pupilas = mProps.Pupilas
End Property

Public Property Let Otros(ByVal value As String)
'    mProps.Otros = Value
'    mbooModificado = True
End Property

Public Property Get Otros() As String
    'Otros = mProps.Otros
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
    Set mProps.objFicha = Nothing
End Sub

Public Sub Recuperar(ByVal Id As Long)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT * FROM vwFichaDetalle WHERE id= " & Id
    Set rs = mProps.objFicha.objEmpresa.OpenRecordset(sql)
        
    With rs
        Do While Not .EOF
        
            mProps.Id = !Id
            mProps.idficha = !idficha
            If Not IsNull(!idCentro) Then mProps.idCentro = !idCentro
            If Not IsNull(!CodCentro) Then mProps.CodCentro = !CodCentro
            If Not IsNull(!Piso) Then mProps.Piso = !Piso
            If Not IsNull(!FechaIngreso) Then mProps.FechaIngreso = !FechaIngreso
            If Not IsNull(!Hora) Then mProps.Hora = !Hora
            mProps.Edad = !Edad
            If Not IsNull(!CodMedico) Then mProps.CodMedico = !CodMedico
            If Not IsNull(!idmedico) Then mProps.idmedico = !idmedico
            If Not IsNull(!Sintomas) Then mProps.Sintomas = !Sintomas
            If Not IsNull(!ExamenFisico) Then mProps.ExamenFisico = !ExamenFisico
            If Not IsNull(!Peso) Then mProps.Peso = !Peso
            If Not IsNull(!Medicamentos) Then mProps.Medicamentos = !Medicamentos
            If Not IsNull(!CP) Then mProps.CP = !CP
            If Not IsNull(!Pulso) Then mProps.Pulso = !Pulso
            If Not IsNull(!Temperatura) Then mProps.Temperatura = !Temperatura
            If Not IsNull(!Respiracion) Then mProps.Respiracion = !Respiracion
            If Not IsNull(!BandAlta) Then mProps.BandAlta = !BandAlta

               rs.MoveNext
        Loop
            .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Grabar(ByVal Id As Long)
    Dim sql As String, rs As Recordset
    On Error GoTo Etiq
    
            sql = "SELECT * FROM FichaDetalle WHERE id= " & Id
            Set rs = mProps.objFicha.objEmpresa.OpenRecordsetParaEdit(sql)
            mbooEsNuevo = False
        
        If rs.RecordCount = 0 Then
            mbooEsNuevo = True
            Set rs = Nothing
            sql = "SELECT * FROM FichaDetalle WHERE 1=0"
            Set rs = mProps.objFicha.objEmpresa.OpenRecordsetParaEdit(sql)
            rs.AddNew
        Else
        End If
    
    With rs
            !idCentro = mProps.idCentro
            !Piso = mProps.Piso
            !idficha = mProps.objFicha.idficha
            !FechaIngreso = mProps.FechaIngreso
            !Hora = mProps.Hora
            !Edad = mProps.Edad
            !idmedico = mProps.idmedico
            !Sintomas = mProps.Sintomas
            !ExamenFisico = mProps.ExamenFisico
'            !DX = mProps.DX
'            !Antesedentes = mProps.Antesedentes
            !Peso = mProps.Peso
'            !TipoSangre = mProps.TipoSangre
'            !Infancia = mProps.Infancia
'            !Pubertad = mProps.Pubertad
'            !Adulto = mProps.Adulto
'            !Habito = mProps.Habito
'            !Alergia = mProps.Alergia
            !Medicamentos = mProps.Medicamentos
'            !DerivadosFE = mProps.DerivadosFE
'            !DerivadosRA = mProps.DerivadosRA
'            !Antidiabeticos = mProps.Antidiabeticos
'            !Corticostiroides = mProps.Corticostiroides
            !CP = mProps.CP
'            !Antibioticos = mProps.Antibioticos
            !Pulso = mProps.Pulso
            !Temperatura = mProps.Temperatura
            !Respiracion = mProps.Respiracion
            !BandAlta = mProps.BandAlta
            
        .Update
            .Move 0
            mProps.Id = !Id
        .Close
    End With
    
    Set rs = Nothing
    'mbooEsNuevo = False
    mbooModificado = False
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Public Property Let Historia(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "Ficha_Historia", MSGERR_INVALIDO
    End If
    mProps.Historia = value
    mbooModificado = True
End Property

Public Property Get Historia() As String
    Historia = mProps.Historia
End Property
Public Property Let Id(ByVal value As Long)
    mProps.Id = value
    mbooModificado = True
End Property

Public Property Get Id() As Long
    Id = mProps.Id
End Property

Public Property Let CodCentro(ByVal value As String)
    Dim sql As String
    Dim rs As Recordset
    
    'Cuando SoloVer está establecido a True, genera error
    'If mbooSoloVer Then Err.Raise ERR_SOLOVER, "FichaDetalle.IdCentro", MSGERR_SOLOVER
    
    'Cuando cambia el valor
    If value <> mProps.CodCentro Then
        If Len(value) > 0 Then
            'Actualiza IdCentro también
            sql = "SELECT IdCentro FROM GNCentroCosto WHERE CodCentro='" & value & "'"
            Set rs = mProps.objFicha.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "FichaDetalle.CodCentro", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                idCentro = rs!idCentro
            End If
            rs.Close
            Set rs = Nothing
        Else
            idCentro = 0
        End If
        mProps.CodCentro = value
        mbooModificado = True
    End If
End Property

Public Property Get CodCentro() As String
    CodCentro = Trim$(mProps.CodCentro)
End Property

Public Property Let CodMedico(ByVal value As String)
    Dim sql As String
    Dim rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodMedico Then
        If Len(value) > 0 Then
            'Actualiza IdCentro también
            sql = "SELECT IdVendedor FROM FCVendedor WHERE CodVendedor='" & value & "'"
            Set rs = mProps.objFicha.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "FichaMedico.CodMedico", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                idmedico = rs!IdVendedor
            End If
            rs.Close
            Set rs = Nothing
        Else
            idmedico = 0
        End If
        mProps.CodMedico = value
        mbooModificado = True
    End If
End Property

Public Property Get CodMedico() As String
    CodMedico = Trim$(mProps.CodMedico)
End Property

Public Property Let CodCliente(ByVal value As String)
    Dim sql As String
    Dim rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodCliente Then
        If Len(value) > 0 Then
            'Actualiza IdCentro también
            sql = "SELECT IdProvCli FROM PCProvCli WHERE CodProvcli='" & value & "'"
            Set rs = mProps.objFicha.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Ficha.CodCliente", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCliente = rs!IdProvCli
            End If
            rs.Close
            Set rs = Nothing
        Else
            idmedico = 0
        End If
        mProps.CodCliente = value
        mbooModificado = True
    End If
End Property

Public Property Get CodCliente() As String
    CodCliente = Trim$(mProps.CodCliente)
End Property
Public Property Let CodOrgano(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "FichaDetalle_codOrgano", MSGERR_INVALIDO
    End If
    mProps.CodOrgano = value
    mbooModificado = True
End Property

Public Property Get CodOrgano() As String
    CodOrgano = mProps.CodOrgano
End Property

Public Property Let DescOrgano(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "FichaDetalle_DesOrgano", MSGERR_INVALIDO
    End If
    mProps.DescOrgano = value
    mbooModificado = True
End Property

Public Property Get DescOrgano() As String
    DescOrgano = mProps.DescOrgano
End Property

Public Property Let SPOrgano(ByVal value As Boolean)
    mProps.SPOrgano = value
    mbooModificado = True
End Property

Public Property Get SPOrgano() As Boolean
    SPOrgano = mProps.SPOrgano
End Property

Public Property Let CPOrgano(ByVal value As Boolean)
    mProps.CPOrgano = value
    mbooModificado = True
End Property

Public Property Get CPOrgano() As Boolean
    CPOrgano = mProps.CPOrgano
End Property

Public Sub grabarOrgano(ByVal Idorgano As Long)
    Dim sql As String, rs As Recordset
    On Error GoTo Etiq
        
    sql = "SELECT * FROM FOrgano WHERE idorgano= " & Idorgano
   Set rs = mProps.objFicha.objEmpresa.OpenRecordsetParaEdit(sql)
       
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sql = "SELECT * FROM forgano WHERE 1=0"
        Set rs = mProps.objFicha.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    Else
    End If
    With rs
            !IdFichaDetalle = mProps.Id
            !CP = mProps.CPOrgano
            !SP = mProps.SPOrgano
            !Codigo = mProps.CodOrgano
            !Descripcion = mProps.DescOrgano
        .Update
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    'Me.objEmpresa.RaiseRefreshEvent REFRESH_FICHA
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Public Sub grabarExamenFisico(ByVal IdExamenFisico As Long)
    Dim sql As String, rs As Recordset
    On Error GoTo Etiq
        
    sql = "SELECT * FROM FExamenFisico WHERE idExamenFisico= " & IdExamenFisico
   Set rs = mProps.objFicha.objEmpresa.OpenRecordsetParaEdit(sql)
       
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        sql = "SELECT * FROM FExamenFisico WHERE 1=0"
        Set rs = mProps.objFicha.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    Else
    End If
    With rs
            !IdFichaDetalle = mProps.Id
            !CP = mProps.CPExamen
            !SP = mProps.SPExamen
            !Codigo = mProps.CodExamen
            !Descripcion = mProps.DescExamen
        .Update
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    'Me.objEmpresa.RaiseRefreshEvent REFRESH_FICHA
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub


Public Property Let CodExamen(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "FichaDetalle_codExamen", MSGERR_INVALIDO
    End If
    mProps.CodExamen = value
    mbooModificado = True
End Property

Public Property Get CodExamen() As String
    CodExamen = mProps.CodExamen
End Property

Public Property Let DescExamen(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "FichaDetalle_DesExamen", MSGERR_INVALIDO
    End If
    mProps.DescExamen = value
    mbooModificado = True
End Property

Public Property Get DescExamen() As String
    DescExamen = mProps.DescExamen
End Property

Public Property Let SPExamen(ByVal value As Boolean)
    mProps.SPExamen = value
    mbooModificado = True
End Property

Public Property Get SPExamen() As Boolean
    SPExamen = mProps.SPExamen
End Property

Public Property Let CPExamen(ByVal value As Boolean)
    mProps.CPExamen = value
    mbooModificado = True
End Property

Public Property Get CPExamen() As Boolean
    CPExamen = mProps.CPExamen
End Property
'-----------------------------------------
Public Sub grabarDiagnostico(ByVal Id As Long)
    Dim sql As String, rs As Recordset
    On Error GoTo Etiq
    
    sql = "SELECT * FROM FDiagnostico WHERE idDiagnostico= " & Id
    Set rs = mProps.objFicha.objEmpresa.OpenRecordsetParaEdit(sql)
    
    If rs.RecordCount = 0 Then
        sql = "SELECT * FROM FDiagnostico WHERE 1=0"
        Set rs = mProps.objFicha.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    Else
    
    End If

    With rs
        !IdFichaDetalle = mProps.Id
        !Codigo = mProps.CodDiag
        !Descripcion = mProps.DescDiag
        .Update
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub


Public Property Let CodDiag(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "FichaDetalle_codDiag", MSGERR_INVALIDO
    End If
    mProps.CodDiag = value
    mbooModificado = True
End Property

Public Property Get CodDiag() As String
    CodDiag = mProps.CodDiag
End Property

Public Property Let DescDiag(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "FichaDetalle_DesDiag", MSGERR_INVALIDO
    End If
    mProps.DescDiag = value
    mbooModificado = True
End Property

Public Property Get DescDiag() As String
    DescDiag = mProps.DescDiag
End Property

Public Sub EliminarDiag(ByVal Id As Long)
'    Dim sql As String, rs As Recordset
'    sql = "DELETE FROM FDiagnostico WHERE IdDiagnostico=" & id
'    mProps.objFicha.objEmpresa.Execute sql, True
End Sub

Public Sub EliminarPlan(ByVal Id As Long)
'    Dim sql As String, rs As Recordset
'    sql = "DELETE FROM FPlan WHERE IdPlan=" & id
'    mProps.objFicha.objEmpresa.Execute sql, True
End Sub

Public Property Let DescPlan(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "FichaDetalle_DesPlan", MSGERR_INVALIDO
    End If
    mProps.DescPlan = value
    mbooModificado = True
End Property

Public Property Get DescPlan() As String
    DescPlan = mProps.DescPlan
End Property

Public Sub grabarPlan(ByVal Id As Long)
'    Dim sql As String, rs As Recordset
'    On Error GoTo Etiq
'
'    sql = "SELECT * FROM FPlan WHERE idPlan= " & id
'    Set rs = mProps.objFicha.objEmpresa.OpenRecordsetParaEdit(sql)
'
'    If rs.RecordCount = 0 Then
'        sql = "SELECT * FROM FPlan WHERE 1=0"
'        Set rs = mProps.objFicha.objEmpresa.OpenRecordsetParaEdit(sql)
'        rs.AddNew
'    Else
'
'    End If
'
'    With rs
'        !IdFichaDetalle = mProps.id
'        '!Codigo = mProps.CodDiag
'        !Descripcion = mProps.DescPlan
'        .Update
'        .Close
'    End With
'    Set rs = Nothing
'    mbooEsNuevo = False
'    mbooModificado = False
'
'    Exit Sub
'Etiq:
'    Err.Raise Err.Number, Err.Source, Err.Description
'    Exit Sub
End Sub

Public Sub BorradoFisico(cod As Long, idficha As Long)
    Dim sql As String
    Dim rs As Recordset
    On Error GoTo CapturaError
        
    
    sql = "delete from FichaDetalle where IdFicha= " & idficha
    objFicha.objEmpresa.EjecutarSQL sql, 1
        
    sql = "delete from gncentrocostodetalle where idfichaDet = " & cod
    objFicha.objEmpresa.EjecutarSQL sql, 1
    
    Set rs = Nothing
    Exit Sub
CapturaError:
'    MsgBox Err.Description
    Set rs = Nothing
    Exit Sub

        
End Sub

Public Sub recuperarxFicha(ByVal Id As Long)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT * FROM vwFichaDetalle WHERE idFicha= " & Id
    Set rs = mProps.objFicha.objEmpresa.OpenRecordset(sql)
        
    With rs
        Do While Not .EOF
            mProps.Id = !Id
            mProps.idficha = !idficha
            rs.MoveNext
        Loop
            .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub


Public Sub grabarSigno(ByVal Id As Long)
    Dim sql As String, rs As Recordset
    On Error GoTo Etiq
    
    sql = "SELECT * FROM FSigno WHERE idSigno= " & Id
    Set rs = mProps.objFicha.objEmpresa.OpenRecordsetParaEdit(sql)
    
    If rs.RecordCount = 0 Then
        sql = "SELECT * FROM FSigno WHERE 1=0"
        Set rs = mProps.objFicha.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    Else
    
    End If

    With rs
        !IdFichaDetalle = mProps.Id
        !Codigo = mProps.CodSigno
        !Descripcion = mProps.DescSigno
        .Update
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub
Public Property Let CodSigno(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "FichaDetalle_codSigno", MSGERR_INVALIDO
    End If
    mProps.CodSigno = value
    mbooModificado = True
End Property

Public Property Get CodSigno() As String
    CodSigno = mProps.CodSigno
End Property

Public Property Let DescSigno(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "FichaDetalle_DesSigno", MSGERR_INVALIDO
    End If
    mProps.DescSigno = value
    mbooModificado = True
End Property

Public Property Get DescSigno() As String
    DescSigno = mProps.DescSigno
End Property
Public Sub BorradoFisicoDetalle(Id As Long)
   
End Sub

Public Sub CambiaEstadoCC(ByVal idCentro As Long, ByVal Estado As Integer)
    Dim sql As String, rs As Recordset
    On Error GoTo Etiq
        sql = "Update GncentroCosto set bandOcupado = " & Estado
        sql = sql & " Where idCentro = " & idCentro
        mProps.objFicha.objEmpresa.EjecutarSQL sql, 1
        
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Public Property Let BandAlta(ByVal value As Boolean)
    mProps.BandAlta = value
    mbooModificado = True
End Property
    
Public Property Get BandAlta() As Boolean
    BandAlta = mProps.BandAlta
End Property

Public Property Let Modificar(ByVal value As Boolean)
    mProps.Modificar = value
End Property

Public Property Get Modificar() As Boolean
    Modificar = mProps.Modificar
End Property

Public Sub recuperarxCCDetalle(ByVal key As Long)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT fd.* FROM FichaDetalle fd INNER JOIN gncentrocostodetalle ccd on ccd.idfichadet=fd.id  WHERE ccd.idCentrodet =" & key
    Set rs = mProps.objFicha.objEmpresa.OpenRecordset(sql)
        
    With rs
        Do While Not .EOF
        
        
            mProps.Id = !Id
            mProps.idficha = !idficha
            If Not IsNull(!idCentro) Then mProps.idCentro = !idCentro

            If Not IsNull(!Piso) Then mProps.Piso = !Piso
            If Not IsNull(!FechaIngreso) Then mProps.FechaIngreso = !FechaIngreso
            If Not IsNull(!Hora) Then mProps.Hora = !Hora
            mProps.Edad = !Edad

            If Not IsNull(!idmedico) Then mProps.idmedico = !idmedico
            If Not IsNull(!Sintomas) Then mProps.Sintomas = !Sintomas
            If Not IsNull(!ExamenFisico) Then mProps.ExamenFisico = !ExamenFisico
            If Not IsNull(!Peso) Then mProps.Peso = !Peso
            If Not IsNull(!Medicamentos) Then mProps.Medicamentos = !Medicamentos
            If Not IsNull(!CP) Then mProps.CP = !CP
            If Not IsNull(!Pulso) Then mProps.Pulso = !Pulso
            If Not IsNull(!Temperatura) Then mProps.Temperatura = !Temperatura
            If Not IsNull(!Respiracion) Then mProps.Respiracion = !Respiracion
            If Not IsNull(!BandAlta) Then mProps.BandAlta = !BandAlta

               rs.MoveNext
        Loop
            .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Function BorradoFisicoDet(Id As Long) As Boolean
    Dim sql As String
    Dim rs As Recordset
    
    'revisamos que no tenga trans relacionada
    sql = "Select gn.codtrans,gn.numtrans,gn.idcentrohijo from gncomprobante gn "
    sql = sql & " Inner Join gncentrocostodetalle ccd "
    sql = sql & "inner join FichaDetalle fd on fd.id = ccd.idfichadet"
    sql = sql & " on ccd.idcentrodet = gn.idcentrohijo "
    sql = sql & " where idfichaDet = " & Id
    Set rs = objFicha.objEmpresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        MsgBox "No puede borrar la ficha tiene registros relacionados" & Chr(13) & rs!CodTrans & rs!NumTrans, vbCritical
        Set rs = Nothing
        Exit Function
    End If
    
    sql = "delete from FichaDetalle where Id= " & Id
    objFicha.objEmpresa.EjecutarSQL sql, 1
    sql = "delete from GNCentroCostoDetalle where IdFichaDet= " & Id
    objFicha.objEmpresa.EjecutarSQL sql, 1
    Set rs = Nothing
    BorradoFisicoDet = True
    Exit Function
CapturaError:
    MsgBox Err.Description
    BorradoFisicoDet = False
    Set rs = Nothing
    Exit Function
End Function

