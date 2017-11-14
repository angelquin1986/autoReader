VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RMotivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public mbooEsNuevo As Boolean
Public mbooModificado As Boolean

Private Type T_Props
      IdMotivo As Long
      Codigo As String
      Descripcion As String
      Remunerado As Boolean
      objEmpresa As Empresa       'Referencia al objeto padre
End Type
Private mProps As T_Props

Friend Property Set objEmpresa(obj As Empresa)
      Set mProps.objEmpresa = obj
End Property

Public Property Get objEmpresa() As Empresa
      Set objEmpresa = mProps.objEmpresa
End Property

Public Property Get Codigo() As String
      Codigo = mProps.Codigo
End Property

Public Property Let Codigo(ByVal Value As String)
      mProps.Codigo = Value
      mbooModificado = True
End Property


Public Property Get Descripcion() As String
      Descripcion = mProps.Descripcion
End Property

Public Property Let Descripcion(ByVal Value As String)
      mProps.Descripcion = Value
      mbooModificado = True
End Property
Public Property Get Remunerado() As Boolean
      Remunerado = mProps.Remunerado
End Property

Public Property Let Remunerado(ByVal Value As Boolean)
      mProps.Remunerado = Value
      mbooModificado = True
End Property


Public Sub grabar(ByVal cod As Long)
      Dim rs  As Recordset
      Dim sql, Fecha As String
      Dim cond As String
      sql = "Select *  from RMotivo"
      With rs
            If mbooModificado = False Then Exit Sub
            Set rs = objEmpresa.OpenRecordsetParaEdit(sql)
            If mbooEsNuevo = True Then rs.AddNew
            
            rs!Codigo = mProps.Codigo
            rs!Descripcion = mProps.Descripcion
            rs!Remunerado = mProps.Remunerado
            
            rs.Update
      End With
      Set rs = Nothing
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_RMOTIVO
End Sub

Public Sub BorradoFisico(Id As Long)
      Dim sql As String
      sql = "delete from RMotivo where IdMotivo= " & Id
      objEmpresa.EjecutarSQL sql, 1
End Sub

Public Sub recuperar(ByVal key As Long)

    Dim rs As Recordset
    Dim sql As String
       
    sql = "SELECT * From RMotivo"
    sql = sql & "  WHERE t.idMotivo=" & key

      Set rs = objEmpresa.OpenRecordsetParaEdit(sql)
      If rs.RecordCount > 0 Then
            With rs
                  mProps.IdMotivo = !IdMotivo
                  mProps.Codigo = !Codigo
                  mProps.Descripcion = !Descripcion
                  mProps.Remunerado = !Remunerado
                  
            End With
            mbooEsNuevo = False
      Else
            'Err.Raise ERR_NOHAYCODIGO, "clsTurno", MSGERR_NOHAYCODIGO
            mbooEsNuevo = True
      End If
End Sub

Public Sub Class_Initialize()
      mbooModificado = False
      mbooEsNuevo = True
End Sub


