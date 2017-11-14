VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PermisoTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Type T_PROP
    objPermiso As Permiso           'Referencia al objeto padre
    CodTrans As String            'Codigo de empresa

    Crear As Boolean
    Ver As Boolean
    Modificar As Boolean
    Eliminar As Boolean
    Aprobar As Boolean
    Desaprobar As Boolean
    Despachar As Boolean
    Anular As Boolean
End Type

Private mProps As T_PROP



Friend Sub Grabar()
    Dim sql As String, rs As Recordset
    
    sql = "SELECT * FROM PermisoTrans WHERE 1=0"

#If DAOLIB Then
    Set rs = gobjMain.Database.OpenRecordset(sql)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.Open sql, gobjMain.Coneccion, adOpenKeyset, adLockPessimistic
#End If
    
    With rs
        .AddNew
        !IdPermiso = mProps.objPermiso.Id
        !CodTrans = mProps.CodTrans
        !Crear = mProps.Crear
        !Modificar = mProps.Modificar
        !Ver = mProps.Ver
        !Eliminar = mProps.Eliminar
        !Aprobar = mProps.Aprobar
        !Desaprobar = mProps.Desaprobar
        !Despachar = mProps.Despachar
        !Anular = mProps.Anular
        .Update
        .Close
    End With
    Set rs = Nothing
End Sub



Friend Sub Recuperar()
    Dim sql As String, rs As Recordset
  
    sql = "SELECT * FROM PermisoTrans " & _
          "WHERE IdPermiso=" & mProps.objPermiso.Id & _
             " AND CodTrans='" & mProps.CodTrans & "'"
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
        mProps.Crear = !Crear
        mProps.Ver = !Ver
        mProps.Modificar = !Modificar
        mProps.Eliminar = !Eliminar
        mProps.Aprobar = !Aprobar
        mProps.Desaprobar = !Desaprobar
        mProps.Despachar = !Despachar
        mProps.Anular = !Anular
    End With
End Sub


Friend Function Clone() As PermisoTrans
    Dim nuevo As PermisoTrans
    
    Set nuevo = New PermisoTrans
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROP)          'Para el método Clone
    mProps = value
    mProps.objPermiso.SetModificado
End Property

Friend Property Set Permiso(value As Permiso)     'Para el método AddPermisoTrans de Grupo
    Set mProps.objPermiso = value
End Property

Public Property Get Permiso() As Permiso
    Set Permiso = mProps.objPermiso
End Property

Public Property Let CodTrans(ByVal value As String)
    If Len(value) > 5 Then
        Err.Raise ERR_INVALIDO, "PermisoTrans_CodTrans", MSGERR_INVALIDO
    End If
    
    mProps.CodTrans = value
    mProps.objPermiso.SetModificado
End Property

Public Property Get CodTrans() As String
    CodTrans = Trim$(mProps.CodTrans)
End Property

Public Property Let Crear(ByVal value As Boolean)
    mProps.Crear = value
    mProps.objPermiso.SetModificado
End Property

Public Property Get Crear() As Boolean
    Crear = mProps.Crear
End Property

Public Property Let Ver(ByVal value As Boolean)
    mProps.Ver = value
    mProps.objPermiso.SetModificado
End Property

Public Property Get Ver() As Boolean
    Ver = mProps.Ver
End Property

Public Property Let Modificar(ByVal value As Boolean)
    mProps.Modificar = value
    mProps.objPermiso.SetModificado
End Property

Public Property Get Modificar() As Boolean
    Modificar = mProps.Modificar
End Property

Public Property Let Eliminar(ByVal value As Boolean)
    mProps.Eliminar = value
    mProps.objPermiso.SetModificado
End Property

Public Property Get Eliminar() As Boolean
    Eliminar = mProps.Eliminar
End Property

Public Property Let Aprobar(ByVal value As Boolean)
    mProps.Aprobar = value
    mProps.objPermiso.SetModificado
End Property

Public Property Get Aprobar() As Boolean
    Aprobar = mProps.Aprobar
End Property

Public Property Let Desaprobar(ByVal value As Boolean)
    mProps.Desaprobar = value
    mProps.objPermiso.SetModificado
End Property

Public Property Get Desaprobar() As Boolean
    Desaprobar = mProps.Desaprobar
End Property

Public Property Let Despachar(ByVal value As Boolean)
    mProps.Despachar = value
    mProps.objPermiso.SetModificado
End Property

Public Property Get Despachar() As Boolean
    Despachar = mProps.Despachar
End Property

Public Property Let Anular(ByVal value As Boolean)
    mProps.Anular = value
    mProps.objPermiso.SetModificado
End Property

Public Property Get Anular() As Boolean
    Anular = mProps.Anular
End Property


Private Sub Class_Terminate()
    Set mProps.objPermiso = Nothing
End Sub


