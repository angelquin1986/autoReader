VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PermisoMenuReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Private Type T_PROP
    id As Long
    objPermiso As Permiso      'Referencia al objeto padre
    NumMenu As Byte            '# de menu (1-9)
    Orden As Byte
    Caption As String
    Modulo As String
    Param As String
    bandVer As Boolean
    OrdenNew As Integer
    Empresas As String
End Type

Private mProps As T_PROP


Friend Sub grabar()
    Dim sql As String, rs As Recordset
    
'    sql = "DELETE FROM PermisoMenuReporte WHERE idpermiso=" & mProps.objPermiso.id
'    Set rs = New ADODB.Recordset
'    rs.CursorLocation = adUseServer
'    rs.Open sql, gobjMain.Coneccion, adOpenKeyset, adLockPessimistic
''
'

    
        
    sql = "SELECT * FROM PermisoMenuReporte WHERE 1=0"
#If DAOLIB Then
    Set rs = gobjMain.Database.OpenRecordset(sql, , dbAppendOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.Open sql, gobjMain.Coneccion, adOpenKeyset, adLockPessimistic
#End If

    
    With rs
        .AddNew
        !IdPermiso = mProps.objPermiso.id
        !NumMenu = mProps.NumMenu
        !Orden = mProps.OrdenNew
        !Caption = mProps.Caption
        !Modulo = mProps.Modulo
        !Param = mProps.Param
        !bandVer = mProps.bandVer
'        !OrdenNew = mProps.OrdenNew

        .Update
#If DAOLIB Then
        .Bookmark = .LastModified
#End If
        .Move 0
        mProps.id = !id
        .Close
    End With
    Set rs = Nothing
End Sub

Friend Sub recuperar(id As Long)
    Dim sql As String, rs As Recordset
  
    sql = "SELECT * FROM PermisoMenuReporte WHERE Id=" & id
#If DAOLIB Then
    Set rs = gobjMain.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly
#End If
    RecuperarRS rs
    rs.Close
    Set rs = Nothing
End Sub

Friend Sub RecuperarRS(ByVal rs As Recordset)
    With rs
        mProps.id = id
        If Not IsNull(!NumMenu) Then mProps.NumMenu = !NumMenu
        If Not IsNull(!Orden) Then mProps.OrdenNew = !Orden
        If Not IsNull(!Caption) Then mProps.Caption = Trim$(!Caption)
        If Not IsNull(!Modulo) Then mProps.Modulo = Trim$(!Modulo)
        If Not IsNull(!Param) Then mProps.Param = Trim$(!Param)
        mProps.bandVer = !bandVer
'        If Not IsNull(!OrdenNew) Then mProps.OrdenNew = !OrdenNew

    End With
End Sub



Friend Function Clone() As PermisoMenuReporte
    Dim nuevo As PermisoMenuReporte
    
    Set nuevo = New PermisoMenuReporte
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(Value As T_PROP)          'Para el método Clone
    mProps = Value
    mProps.objPermiso.SetModificado
End Property

Friend Property Set Permiso(Value As Permiso)     'Para el método AddPermisoMenuReporte de Grupo
    Set mProps.objPermiso = Value
End Property

Public Property Get Permiso() As Permiso
    Set Permiso = mProps.objPermiso
End Property

Public Property Get id() As Long
    id = mProps.id
End Property

Public Property Let NumMenu(ByVal Value As Byte)
    mProps.NumMenu = Value
    mProps.objPermiso.SetModificado
End Property

Public Property Get NumMenu() As Byte
    NumMenu = mProps.NumMenu
End Property

Public Property Let Orden(ByVal Value As Byte)
    mProps.Orden = Value
    mProps.objPermiso.SetModificado
End Property

Public Property Get Orden() As Byte
    Orden = mProps.Orden
End Property

Public Property Let Modulo(ByVal Value As String)
    If Len(Value) > 10 Then
        Err.Raise ERR_INVALIDO, "PermisoMenuReporte_Modulo", MSGERR_INVALIDO
    End If
    mProps.Modulo = Value
    mProps.objPermiso.SetModificado
End Property

Public Property Get Modulo() As String
    Modulo = Trim$(mProps.Modulo)
End Property

Public Property Let Caption(ByVal Value As String)
    If Len(Value) > 40 Then
        Err.Raise ERR_INVALIDO, "PermisoMenuReporte_Caption", MSGERR_INVALIDO
    End If
    mProps.Caption = Value
    mProps.objPermiso.SetModificado
End Property

Public Property Get Caption() As String
    Caption = Trim$(mProps.Caption)
End Property

Public Property Let Param(ByVal Value As String)
    If Len(Value) > 40 Then
        Err.Raise ERR_INVALIDO, "PermisoMenuReporte_Param", MSGERR_INVALIDO
    End If
    mProps.Param = Value
    mProps.objPermiso.SetModificado
End Property

Public Property Get Param() As String
    Param = Trim$(mProps.Param)
End Property


Private Sub Class_Terminate()
    Set mProps.objPermiso = Nothing
End Sub


Public Property Let bandVer(ByVal Value As Boolean)
    mProps.bandVer = Value
    mProps.objPermiso.SetModificado
End Property

Public Property Get bandVer() As Boolean
    bandVer = mProps.bandVer
End Property


Public Property Let OrdenNew(ByVal Value As Integer)
    mProps.OrdenNew = Value
    mProps.objPermiso.SetModificado
End Property

Public Property Get OrdenNew() As Integer
    OrdenNew = mProps.OrdenNew
End Property

