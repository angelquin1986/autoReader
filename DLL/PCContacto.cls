VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCContacto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROP
    objPC As PCProvCli     'Referencia al objeto padre
    Id As Long             'Id de contacto

    Cargo As String
    Titulo As String
    Nombre As String
    Telefono1 As String
    Telefono2 As String
    EMail As String
    Orden As Integer
End Type

Private mProps As T_PROP



Friend Sub Grabar()
    Dim rs As Recordset, sql As String

    sql = "SELECT * FROM PCContacto WHERE 1=0"
    Set rs = mProps.objPC.Empresa.OpenRecordsetParaEdit(sql)
    rs.AddNew
    
    With rs
        !IdProvcli = mProps.objPC.IdProvcli
        !Cargo = mProps.Cargo
        !Titulo = mProps.Titulo
        !Nombre = mProps.Nombre
        !Telefono1 = mProps.Telefono1
        !Telefono2 = mProps.Telefono2
        !EMail = mProps.EMail
        !Orden = mProps.Orden
        .Update
#If DAOLIB Then
        .Bookmark = .LastModified
#End If
        rs.Move 0
        mProps.Id = !IdContacto
        .Close
    End With
    Set rs = Nothing
End Sub

Friend Sub Recuperar(Id As Long)
    Dim sql As String, rs As Recordset
  
    sql = "SELECT * FROM PCContacto WHERE IdContacto=" & Id
    Set rs = mProps.objPC.Empresa.OpenRecordset(sql)
    With rs
        mProps.Id = Id
        If Not IsNull(!Cargo) Then mProps.Cargo = !Cargo
        If Not IsNull(!Titulo) Then mProps.Titulo = !Titulo
        If Not IsNull(!Nombre) Then mProps.Nombre = !Nombre
        If Not IsNull(!Telefono1) Then mProps.Telefono1 = !Telefono1
        If Not IsNull(!Telefono2) Then mProps.Telefono2 = !Telefono2
        If Not IsNull(!EMail) Then mProps.EMail = !EMail
        If Not IsNull(!Orden) Then mProps.Orden = !Orden
        .Close
    End With
    Set rs = Nothing
End Sub

Friend Function Clone() As PCContacto
    Dim nuevo As PCContacto
    
    Set nuevo = New PCContacto
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROP)          'Para el método Clone
    mProps = value
    mProps.objPC.SetModificado
End Property

Friend Property Set PCProvCli(value As PCProvCli)     'Para el método AddPCContacto de Grupo
    Set mProps.objPC = value
End Property

Public Property Get PCProvCli() As PCProvCli
    Set PCProvCli = mProps.objPC
End Property

Public Property Get IdContacto() As Long
    IdContacto = mProps.Id
End Property

Public Property Let Cargo(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCContacto_Cargo", MSGERR_INVALIDO
    End If
    mProps.Cargo = value
    mProps.objPC.SetModificado
End Property

Public Property Get Cargo() As String
    Cargo = Trim$(mProps.Cargo)
End Property

Public Property Let Titulo(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "PCContacto_Titulo", MSGERR_INVALIDO
    End If
    mProps.Titulo = value
    mProps.objPC.SetModificado
End Property

Public Property Get Titulo() As String
    Titulo = Trim$(mProps.Titulo)
End Property

Public Property Let Nombre(ByVal value As String)
    If Len(value) > 40 Then
        Err.Raise ERR_INVALIDO, "PCContacto_Nombre", MSGERR_INVALIDO
    End If
    mProps.Nombre = value
    mProps.objPC.SetModificado
End Property

Public Property Get Nombre() As String
    Nombre = Trim$(mProps.Nombre)
End Property

Public Property Let Telefono1(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCContacto_Telefono1", MSGERR_INVALIDO
    End If
    mProps.Telefono1 = value
    mProps.objPC.SetModificado
End Property

Public Property Get Telefono1() As String
    Telefono1 = Trim$(mProps.Telefono1)
End Property

Public Property Let Telefono2(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCContacto_Telefono2", MSGERR_INVALIDO
    End If
    mProps.Telefono2 = value
    mProps.objPC.SetModificado
End Property

Public Property Get Telefono2() As String
    Telefono2 = Trim$(mProps.Telefono2)
End Property

Public Property Let EMail(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "PCContacto_EMail", MSGERR_INVALIDO
    End If
    mProps.EMail = value
    mProps.objPC.SetModificado
End Property

Public Property Get EMail() As String
    EMail = Trim$(mProps.EMail)
End Property


Public Property Let Orden(ByVal value As Integer)
    mProps.Orden = value
    mProps.objPC.SetModificado
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property


Private Sub Class_Terminate()
    Set mProps.objPC = Nothing
End Sub




