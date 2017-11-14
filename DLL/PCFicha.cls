VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCFicha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_Props
    IdFicha As Long
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
    IdMedico As Long
    Sintomas As String
    ExamenFisico As String
    DX As String
    Antesedentes As String
    Peso As String
    TipoSangre As String
    Infancia As String
    Pubertad As String
    Adulto As String
    Habito As String
    Alergia As String
    Medicamentos As String
    Corticostiroides As String
    DerivadosFE As String
    DerivadosRA As String
    Antidiabeticos As String
    Antibioticos As String
    CP As String
    Pulso As String
    Temperatura As String
    Respiracion As String
    Pupilas As String
    Otros As String
    objEmpresa As Empresa
End Type
Private mProps As T_Props
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Public Property Let IdFicha(ByVal Value As Long)
    mProps.IdFicha = Value
    mbooModificado = True
End Property

Public Property Get IdFicha() As Long
    IdFicha = mProps.IdFicha
End Property

Friend Property Set objEmpresa(ByVal obj As Empresa)
    Set mProps.objEmpresa = obj
End Property

Public Property Get objEmpresa() As Empresa
    Set objEmpresa = mProps.objEmpresa
End Property

Public Property Let Sexo(ByVal Value As Byte)
    '0=Femenino, 1=Masculino
    If (Not IsNumeric(Value)) Or (Val(Value) < 0 Or Val(Value) > 1) Then Err.Raise ERR_INVALIDO, "PCEmpleado.Sexo", MSGERR_INVALIDO
    mProps.Sexo = Value
    mbooModificado = True
End Property

Public Property Get Sexo() As Byte
    Sexo = mProps.Sexo
End Property

Public Property Let IdCliente(ByVal Value As Long)
    mProps.IdCliente = Value
    mbooModificado = True
End Property

Public Property Get IdCliente() As Long
    IdCliente = mProps.IdCliente
End Property

Public Property Let Cedula(ByVal Value As String)
    If Len(Value) > 10 Then Err.Raise ERR_INVALIDO, "Ficha.Cedula", MSGERR_INVALIDO
    mProps.Cedula = Value
    mbooModificado = True
End Property

Public Property Get Cedula() As String
    Cedula = mProps.Cedula
End Property

Public Property Let idCentro(ByVal Value As Long)
    mProps.idCentro = Value
    mbooModificado = True
End Property

Public Property Get idCentro() As Long
    idCentro = mProps.idCentro
End Property

Public Property Get Piso() As Integer
    Piso = mProps.Piso
End Property
Public Property Let Piso(ByVal Value As Integer)
    mProps.Piso = Value
    mbooModificado = True
End Property

Public Property Let Nombre(ByVal Value As String)
    mProps.Nombre = Value
    mbooModificado = True
End Property

Public Property Get Nombre() As String
    Nombre = mProps.Nombre
End Property

Public Property Let FechaIngreso(ByVal Value As Date)
    If Not IsDate(Value) Then Err.Raise ERR_INVALIDO, "Ficha.FechaIngreso", MSGERR_INVALIDO
    mProps.FechaIngreso = Value
    mbooModificado = True
End Property

Public Property Get FechaIngreso() As Date
    FechaIngreso = mProps.FechaIngreso
End Property

Public Property Let Hora(ByVal Value As Date)
    If Not IsDate(Value) Then Err.Raise ERR_INVALIDO, "Ficha.hora", MSGERR_INVALIDO
    mProps.Hora = Value
    mbooModificado = True
End Property

Public Property Get Hora() As Date
    Hora = mProps.Hora
End Property

Public Property Get Edad() As Integer
    Edad = mProps.Edad
End Property
Public Property Let Edad(ByVal Value As Integer)
    mProps.Edad = Value
    mbooModificado = True
End Property

Public Property Let Procedencia(ByVal Value As String)
    mProps.Procedencia = Value
    mbooModificado = True
End Property

Public Property Get Procedencia() As String
    Procedencia = mProps.Procedencia
End Property
Public Property Let IdCiudad(ByVal Value As Long)
    mProps.IdCiudad = Value
    mbooModificado = True
End Property

Public Property Get IdCiudad() As Long
    IdCiudad = mProps.IdCiudad
End Property

Public Property Let Direccion(ByVal Value As String)
    mProps.Direccion = Value
    mbooModificado = True
End Property

Public Property Get Direccion() As String
    Direccion = mProps.Direccion
End Property

Public Property Let Telefono(ByVal Value As String)
    mProps.Telefono = Value
    mbooModificado = True
End Property

Public Property Get Telefono() As String
    Telefono = mProps.Telefono
End Property

Public Property Let Ocupacion(ByVal Value As String)
    mProps.Ocupacion = Value
    mbooModificado = True
End Property

Public Property Get Ocupacion() As String
    Ocupacion = mProps.Ocupacion
End Property

Public Property Let Religion(ByVal Value As String)
    mProps.Religion = Value
    mbooModificado = True
End Property

Public Property Get Religion() As String
    Religion = mProps.Religion
End Property

Public Property Let IdMedico(ByVal Value As Long)
    mProps.IdMedico = Value
    mbooModificado = True
End Property

Public Property Get IdMedico() As Long
    IdMedico = mProps.IdMedico
End Property

Public Property Let Sintomas(ByVal Value As String)
    mProps.Sintomas = Value
    mbooModificado = True
End Property

Public Property Get Sintomas() As String
    Sintomas = mProps.Sintomas
End Property

Public Property Let ExamenFisico(ByVal Value As String)
    mProps.ExamenFisico = Value
    mbooModificado = True
End Property

Public Property Get ExamenFisico() As String
    ExamenFisico = mProps.ExamenFisico
End Property

Public Property Let DX(ByVal Value As String)
    mProps.DX = Value
    mbooModificado = True
End Property

Public Property Get DX() As String
    DX = mProps.DX
End Property

Public Property Let Antesedentes(ByVal Value As String)
    mProps.Antesedentes = Value
    mbooModificado = True
End Property

Public Property Get Antesedentes() As String
    Antesedentes = mProps.Antesedentes
End Property

Public Property Let Peso(ByVal Value As String)
    mProps.Peso = Value
    mbooModificado = True
End Property

Public Property Get Peso() As String
    Peso = mProps.Peso
End Property

Public Property Let TipoSangre(ByVal Value As String)
    mProps.TipoSangre = Value
    mbooModificado = True
End Property

Public Property Get TipoSangre() As String
    TipoSangre = mProps.TipoSangre
End Property

Public Property Let Infancia(ByVal Value As String)
    mProps.Infancia = Value
    mbooModificado = True
End Property

Public Property Get Infancia() As String
    Infancia = mProps.Infancia
End Property

Public Property Let Pubertad(ByVal Value As String)
    mProps.Pubertad = Value
    mbooModificado = True
End Property

Public Property Get Pubertad() As String
    Pubertad = mProps.Pubertad
End Property

Public Property Let Adulto(ByVal Value As String)
    mProps.Adulto = Value
    mbooModificado = True
End Property

Public Property Get Adulto() As String
    Adulto = mProps.Adulto
End Property

Public Property Let Habito(ByVal Value As String)
    mProps.Habito = Value
    mbooModificado = True
End Property

Public Property Get Habito() As String
    Habito = mProps.Habito
End Property
Public Property Let Alergia(ByVal Value As String)
    mProps.Alergia = Value
    mbooModificado = True
End Property

Public Property Get Alergia() As String
    Alergia = mProps.Alergia
End Property

Public Property Let Medicamentos(ByVal Value As String)
    mProps.Medicamentos = Value
    mbooModificado = True
End Property

Public Property Get Medicamentos() As String
    Medicamentos = mProps.Medicamentos
End Property

Public Property Let Corticostiroides(ByVal Value As String)
    mProps.Corticostiroides = Value
    mbooModificado = True
End Property

Public Property Get Corticostiroides() As String
    Corticostiroides = mProps.Corticostiroides
End Property

Public Property Let DerivadosFE(ByVal Value As String)
    mProps.DerivadosFE = Value
    mbooModificado = True
End Property

Public Property Get DerivadosFE() As String
    DerivadosFE = mProps.DerivadosFE
End Property

Public Property Let DerivadosRA(ByVal Value As String)
    mProps.DerivadosRA = Value
    mbooModificado = True
End Property

Public Property Get DerivadosRA() As String
    DerivadosRA = mProps.DerivadosRA
End Property

Public Property Let Antidiabeticos(ByVal Value As String)
    mProps.Antidiabeticos = Value
    mbooModificado = True
End Property

Public Property Get Antidiabeticos() As String
    Antidiabeticos = mProps.Antidiabeticos
End Property

Public Property Let CP(ByVal Value As String)
    mProps.CP = Value
    mbooModificado = True
End Property

Public Property Get CP() As String
    CP = mProps.CP
End Property

Public Property Let Pulso(ByVal Value As String)
    mProps.Pulso = Value
    mbooModificado = True
End Property

Public Property Get Pulso() As String
    Pulso = mProps.Pulso
End Property



Public Property Let Antibioticos(ByVal Value As String)
    mProps.Antibioticos = Value
    mbooModificado = True
End Property

Public Property Get Antibioticos() As String
    Antibioticos = mProps.Antibioticos
End Property


Public Property Let Temperatura(ByVal Value As String)
    mProps.Temperatura = Value
    mbooModificado = True
End Property

Public Property Get Temperatura() As String
    Temperatura = mProps.Temperatura
End Property

Public Property Let Respiracion(ByVal Value As String)
    mProps.Respiracion = Value
    mbooModificado = True
End Property

Public Property Get Respiracion() As String
    Respiracion = mProps.Respiracion
End Property

Public Property Let Pupilas(ByVal Value As String)
    mProps.Pupilas = Value
    mbooModificado = True
End Property

Public Property Get Pupilas() As String
    Pupilas = mProps.Pupilas
End Property

Public Property Let Otros(ByVal Value As String)
    mProps.Otros = Value
    mbooModificado = True
End Property

Public Property Get Otros() As String
    Otros = mProps.Otros
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
    Set mProps.objEmpresa = Nothing
End Sub

Public Sub recuperar(ByVal id As Long)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT * FROM Ficha WHERE idFicha= " & id
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
        
    With rs
        Do While Not .EOF
            mProps.IdFicha = !IdFicha
            If Not IsNull(!Historia) Then mProps.Historia = !Historia
            If Not IsNull(!IdCliente) Then mProps.IdCliente = !IdCliente
            If Not IsNull(!Sexo) Then mProps.Sexo = !Sexo
            If Not IsNull(!Cedula) Then mProps.Cedula = !Cedula
'            If Not IsNull(!idCentro) Then mProps.idCentro = !idCentro
            'If Not IsNull(!Piso) Then mProps.Piso = !Piso
            If Not IsNull(!Nombre) Then mProps.Nombre = !Nombre
'            If Not IsNull(!FechaIngreso) Then mProps.FechaIngreso = !FechaIngreso
'            If Not IsNull(!Hora) Then mProps.Hora = !Hora
'            mProps.Edad = !Edad
            If Not IsNull(!Procedencia) Then mProps.Procedencia = !Procedencia
            mProps.IdCiudad = !IdCiudad
            If Not IsNull(!Direccion) Then mProps.Direccion = !Direccion
            If Not IsNull(!Telefono) Then mProps.Telefono = !Telefono
            If Not IsNull(!Ocupacion) Then mProps.Ocupacion = !Ocupacion
            If Not IsNull(!Religion) Then mProps.Religion = !Religion
'            If Not IsNull(!IdMedico) Then mProps.IdMedico = !IdMedico
            'If Not IsNull(!Sintomas) Then mProps.Sintomas = !Sintomas
            'If Not IsNull(!ExamenFisico) Then mProps.ExamenFisico = !ExamenFisico
'            If Not IsNull(!DX) Then mProps.DX = !DX
'            If Not IsNull(!Antesedentes) Then mProps.Antesedentes = !Antesedentes
'            If Not IsNull(!Peso) Then mProps.Peso = !Peso
'            If Not IsNull(!TipoSangre) Then mProps.TipoSangre = !TipoSangre
'            If Not IsNull(!Infancia) Then mProps.Infancia = !Infancia
'            If Not IsNull(!Pubertad) Then mProps.Pubertad = !Pubertad
'            If Not IsNull(!Adulto) Then mProps.Adulto = !Adulto
'            If Not IsNull(!Habito) Then mProps.Habito = !Habito
'            If Not IsNull(!Alergia) Then mProps.Alergia = !Alergia
'            If Not IsNull(!Medicamentos) Then mProps.Medicamentos = !Medicamentos
'            If Not IsNull(!Corticostiroides) Then mProps.Corticostiroides = !Corticostiroides
'            If Not IsNull(!DerivadosFE) Then mProps.DerivadosFE = !DerivadosFE
'            If Not IsNull(!DerivadosRA) Then mProps.DerivadosRA = !DerivadosRA
'            If Not IsNull(!Antidiabeticos) Then mProps.Antidiabeticos = !Antidiabeticos
'            If Not IsNull(!CP) Then mProps.CP = !CP
'            If Not IsNull(!Pulso) Then mProps.Pulso = !Pulso
'            If Not IsNull(!Temperatura) Then mProps.Temperatura = !Temperatura
'            If Not IsNull(!Respiracion) Then mProps.Respiracion = !Respiracion
'            If Not IsNull(!Pupilas) Then mProps.Pupilas = !Pupilas
'            If Not IsNull(!Otros) Then mProps.Otros = !Otros
            rs.MoveNext
        Loop
            .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub grabar(ByVal IdFicha As Long)
    Dim sql As String, rs As Recordset
    On Error GoTo Etiq
            
        
            sql = "SELECT * FROM Ficha WHERE idFicha= " & IdFicha
            Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
            
        
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            sql = "SELECT * FROM Ficha WHERE 1=0"
            Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
            rs.AddNew
        Else
        End If
    
    With rs
             !Historia = mProps.Historia
              !IdCliente = mProps.IdCliente
               !Sexo = mProps.Sexo
             !Cedula = mProps.Cedula
'               !idCentro = mProps.idCentro
'              !Piso = mProps.Piso
             !Nombre = mProps.Nombre
'               !FechaIngreso = mProps.FechaIngreso
'             !Hora = mProps.Hora
              '!Edad = mProps.Edad
             !Procedencia = mProps.Procedencia
             !IdCiudad = mProps.IdCiudad
              !Direccion = mProps.Direccion
              !Telefono = mProps.Telefono
               !Ocupacion = mProps.Ocupacion
              !Religion = mProps.Religion
             '!IdMedico = mProps.IdMedico
        .Update
        .Close
    End With

    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    Me.objEmpresa.RaiseRefreshEvent REFRESH_FICHA
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Public Property Let Historia(ByVal Value As String)
    If Len(Value) > 20 Then
        Err.Raise ERR_INVALIDO, "Ficha_Historia", MSGERR_INVALIDO
    End If
    
    mProps.Historia = Value
    mbooModificado = True
    
End Property

Public Property Get Historia() As String
    Historia = mProps.Historia
End Property
