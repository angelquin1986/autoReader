VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCAsignado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROP
    objPC As PCProvCli     'Referencia al objeto padre
    objEmpresa  As Empresa
    IdProvcli As Long
    IdAsignado As Long
    CodProvcli As String
    CodProvcliAsingado As String
    
End Type
Private mProps As T_PROP

Dim mbooEsNuevo As Boolean
Dim mbooModificado As Boolean

'Public Sub Grabar()
'    Dim rs As Recordset, sql As String
'    If Not mbooModificado Then Exit Sub
'    'If mbooEsNuevo Then
'    If mProps.IdAgencia = 0 Then
'        sql = "SELECT * FROM PCAgencia WHERE 1=0"
'        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
'        rs.AddNew
'    Else
'         sql = "SELECT * FROM PCAgencia WHERE IdAGENCIA=" & mProps.IdAgencia
'        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
'    End If
'    With rs
'        !IdProvCli = mProps.IdProvCli
'        !IdCiudad = mProps.IdCiudad
'        !IdProvincia = mProps.IdProvincia
'        !CodAgencia = mProps.CodAgencia
'        !Descripcion = mProps.Descripcion
'        !Direccion = mProps.Direccion
'        !Telefono = mProps.Telefono
'        !Contacto = mProps.Contacto
'        !Orden = mProps.Orden
'        !BandValida = mProps.BandValida
'        .Update
'#If DAOLIB Then
'        .Bookmark = .LastModified
'#End If
'        rs.Move 0
'        mProps.IdAgencia = !IdAgencia
'        .Close
'    End With
'    mbooEsNuevo = False
'    mbooModificado = False
'    Set rs = Nothing
'End Sub

Friend Sub Recuperar(id As Long)
    Dim sql As String, rs As Recordset
    sql = "SELECT *"
    sql = sql & " FROM PCAsigna "
    sql = sql & " Where IdProvCliAsignado = " & id
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    With rs
        mProps.IdProvcli = !IdProvcli
        mProps.IdAsignado = !IdProvcliAsignado
        .Close
    End With
    Set rs = Nothing
End Sub

'Friend Function Clone() As PCAgencia
'    Dim nuevo As PCAgencia
'
'    Set nuevo = New PCAgencia
'    nuevo.Props = mProps
'    Set Clone = nuevo
'    Set nuevo = Nothing
'End Function

Friend Property Let Props(value As T_PROP)          'Para el método Clone
    mProps = value
    mbooModificado = True
End Property

Friend Property Set PCProvCli(value As PCProvCli)     'Para el método AddPCContacto de Grupo
    Set mProps.objPC = value
End Property

Public Property Get PCProvCli() As PCProvCli
    Set PCProvCli = mProps.objPC
End Property


Friend Property Set Empresa(value As Empresa)
    Set mProps.objEmpresa = value
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Get IdProvcli() As Long
    IdProvcli = mProps.IdProvcli
End Property


Private Sub Class_Initialize()
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objPC = Nothing
    Set mProps.objEmpresa = Nothing
    mbooEsNuevo = False
End Sub


Public Property Get EsNuevo() As Boolean
    EsNuevo = mbooEsNuevo
End Property

Public Property Get Modificado() As Boolean
    Modificado = mbooModificado
End Property

Friend Property Let Modificado(ByVal value As Boolean)
    mbooModificado = value
End Property
Public Sub EliminaFisica()
Dim sql As String
sql = "Delete from pcAsigna where idprovcliAsignado =" & Me.IdAsignado
    mProps.objEmpresa.EjecutarSQL sql, 1
    mbooEsNuevo = False
End Sub

Public Property Get IdAsignado() As Long
    IdAsignado = mProps.IdAsignado
End Property



