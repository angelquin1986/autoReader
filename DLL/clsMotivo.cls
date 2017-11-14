VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMotivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public mbooEsNuevo As Boolean
Public mbooModificado As Boolean

Private Type T_PROPS
      IdMotivo As Long
      Codigo As String
      Descripcion As String
      Remunerado As Boolean
      objEmpresa As Empresa
End Type
Private mProps As T_PROPS

Friend Property Set objEmpresa(obj As Empresa)
    Set mProps.objEmpresa = obj
End Property
'
Public Property Get objEmpresa() As Empresa
    Set objEmpresa = mProps.objEmpresa
End Property


Public Property Get Codigo() As String
    Codigo = mProps.Codigo
End Property

Public Property Let Codigo(ByVal value As String)
        mProps.Codigo = value
        mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = mProps.Descripcion
End Property

Public Property Let Descripcion(ByVal value As String)
        mProps.Descripcion = value
        mbooModificado = True
End Property

Public Property Get Remunerado() As Boolean
    Remunerado = mProps.Remunerado
End Property

Public Property Let Remumerado(ByVal value As Boolean)
        mProps.Remunerado = value
        mbooModificado = True
End Property


Public Sub Grabar()
    Dim rs As Recordset
    Dim sql, fecha As String
    Dim cond As String
      sql = "Select * from RMotivo"
     If mbooModificado = False Then Exit Sub
      
      Set rs = objEmpresa.OpenRecordsetParaEdit(sql)
      If mbooEsNuevo = True Then rs.AddNew
      With rs
            rs!Codigo = mProps.Codigo
            rs!Descripcion = mProps.Descripcion
            rs!Remunerado = mProps.Remunerado
            rs.Update
            mProps.IdMotivo = rs!IdMotivo
        End With
        Set rs = Nothing
        mProps.objEmpresa.RaiseRefreshEvent REFRESH_RMOTIVO
        mbooEsNuevo = False
        mbooModificado = False
End Sub



Public Sub BorradoFisico(id As Long)
    Dim sql As String
    Dim rs As Recordset
    
    sql = "Select * from rpermiso where idmotivo = " & id
    Set rs = objEmpresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        MsgBox "No puede borrar el registro tiene registros relacionados, Permiso:" & Chr(13) & rs!Observa, vbCritical
        Set rs = Nothing
        Exit Sub
    End If
        
    sql = "delete from Rmotivo where IdMotivo= " & id
    objEmpresa.EjecutarSQL sql, 1
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_RMOTIVO
End Sub

Public Function Recuperar(ByVal key As Long) As Recordset
    Dim rs As Recordset
    Dim sql As String
    Dim i As Integer
    
    sql = "SELECT * " & _
            " FROM RMotivo  " & _
            " WHERE idmotivo=" & key
            
    Set rs = objEmpresa.OpenRecordsetParaEdit(sql)
    If rs.RecordCount > 0 Then
'          Set recuperar = rs
            mbooEsNuevo = False
            mProps.IdMotivo = rs!IdMotivo
            mProps.Codigo = rs!Codigo
            mProps.Descripcion = rs!Descripcion
            mProps.Remunerado = rs!Remunerado
            
            
    Else
'        Err.Raise ERR_NOHAYCODIGO, "clsTurno", MSGERR_NOHAYCODIGO
        mbooEsNuevo = True
    End If
End Function


Public Sub Class_Initialize()
    mbooModificado = False
    mbooEsNuevo = True
End Sub

Public Property Get EsNuevo() As Boolean
    EsNuevo = mbooEsNuevo
End Property

Public Property Get Modificado() As Boolean
    Modificado = mbooModificado
End Property


