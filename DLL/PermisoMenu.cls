VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PermisoMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Private Type T_PROP
    Id As Long
    objPermiso As Permiso      'Referencia al objeto padre
    NumMenu As Byte            '# de menu (1-9)
    Orden As Byte
    Caption As String
    CodPantalla As String
    Param As String
End Type

Private mProps As T_PROP



Friend Sub Grabar()
    Dim sql As String, rs As Recordset
    
    sql = "SELECT * FROM PermisoMenu WHERE 1=0"
#If DAOLIB Then
    Set rs = gobjMain.Database.OpenRecordset(sql, , dbAppendOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.Open sql, gobjMain.Coneccion, adOpenKeyset, adLockPessimistic
#End If
    rs.AddNew
    
    With rs
        !IdPermiso = mProps.objPermiso.Id
        !NumMenu = mProps.NumMenu
        !Orden = mProps.Orden
        !Caption = mProps.Caption
        !CodPantalla = mProps.CodPantalla
        !Param = mProps.Param
        .Update
#If DAOLIB Then
        .Bookmark = .LastModified
#End If
        .Move 0
        mProps.Id = !Id
        .Close
    End With
    Set rs = Nothing
End Sub

Friend Sub Recuperar(Id As Long)
    Dim sql As String, rs As Recordset
  
    sql = "SELECT * FROM PermisoMenu WHERE Id=" & Id
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
        mProps.Id = Id
        If Not IsNull(!NumMenu) Then mProps.NumMenu = !NumMenu
        If Not IsNull(!Orden) Then mProps.Orden = !Orden
        If Not IsNull(!Caption) Then mProps.Caption = Trim$(!Caption)
        If Not IsNull(!CodPantalla) Then mProps.CodPantalla = Trim$(!CodPantalla)
        If Not IsNull(!Param) Then mProps.Param = Trim$(!Param)
    End With
End Sub



Friend Function Clone() As PermisoMenu
    Dim nuevo As PermisoMenu
    
    Set nuevo = New PermisoMenu
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROP)          'Para el método Clone
    mProps = value
    mProps.objPermiso.SetModificado
End Property

Friend Property Set Permiso(value As Permiso)     'Para el método AddPermisoMenu de Grupo
    Set mProps.objPermiso = value
End Property

Public Property Get Permiso() As Permiso
    Set Permiso = mProps.objPermiso
End Property

Public Property Get Id() As Long
    Id = mProps.Id
End Property

Public Property Let NumMenu(ByVal value As Byte)
    mProps.NumMenu = value
    mProps.objPermiso.SetModificado
End Property

Public Property Get NumMenu() As Byte
    NumMenu = mProps.NumMenu
End Property

Public Property Let Orden(ByVal value As Byte)
    mProps.Orden = value
    mProps.objPermiso.SetModificado
End Property

Public Property Get Orden() As Byte
    Orden = mProps.Orden
End Property

Public Property Let CodPantalla(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "PermisoMenu_CodPantalla", MSGERR_INVALIDO
    End If
    mProps.CodPantalla = value
    mProps.objPermiso.SetModificado
End Property

Public Property Get CodPantalla() As String
    CodPantalla = Trim$(mProps.CodPantalla)
End Property

Public Property Let Caption(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PermisoMenu_Caption", MSGERR_INVALIDO
    End If
    mProps.Caption = value
    mProps.objPermiso.SetModificado
End Property

Public Property Get Caption() As String
    Caption = Trim$(mProps.Caption)
End Property

Public Property Let Param(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PermisoMenu_Param", MSGERR_INVALIDO
    End If
    mProps.Param = value
    mProps.objPermiso.SetModificado
End Property

Public Property Get Param() As String
    Param = Trim$(mProps.Param)
End Property


Private Sub Class_Terminate()
    Set mProps.objPermiso = Nothing
End Sub




