VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFeriado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public mbooEsNuevo As Boolean
Public mbooModificado As Boolean

Private Type T_Props
      IdFeriado As Long
      fecha As String
      Descripcion As String
      objEmpresa As Empresa
End Type
Private mProps As T_Props

Friend Property Set objEmpresa(obj As Empresa)
    Set mProps.objEmpresa = obj
End Property
'
Public Property Get objEmpresa() As Empresa
    Set objEmpresa = mProps.objEmpresa
End Property


Public Property Get fecha() As String
    fecha = mProps.fecha
End Property

Public Property Let fecha(ByVal Value As String)
        mProps.fecha = Value
        mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = mProps.Descripcion
End Property

Public Property Let Descripcion(ByVal Value As String)
        mProps.Descripcion = Value
        mbooModificado = True
End Property

Public Sub grabar()
    Dim rs As Recordset
    Dim sql, fecha As String
    Dim cond As String 'h25,h100,
      sql = "Select * from RFeriado"
     If mbooModificado = False Then Exit Sub
     
      Set rs = objEmpresa.OpenRecordsetParaEdit(sql)
      If mbooEsNuevo = True Then rs.AddNew
      With rs
            rs!fecha = mProps.fecha
            rs!Descripcion = mProps.Descripcion
            rs.Update
        End With
        Set rs = Nothing
        mProps.objEmpresa.RaiseRefreshEvent REFRESH_RFERIADO
        mbooEsNuevo = False
        mbooModificado = False
End Sub

Public Sub BorradoFisico(id As Long)
    Dim sql As String
    sql = "delete from RFeriado where IdFeriado = " & id
    objEmpresa.EjecutarSQL sql, 1
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_RFERIADO
End Sub

Public Function recuperar(ByVal key As Long) As Recordset
    Dim rs As Recordset
    Dim sql As String
    Dim i As Integer
    
    sql = "SELECT * " & _
            " FROM RFeriado  " & _
            " WHERE idFeriado=" & key
            
    Set rs = objEmpresa.OpenRecordsetParaEdit(sql)
    If rs.RecordCount > 0 Then
    
            mProps.IdFeriado = rs!IdFeriado
            mProps.fecha = rs!fecha
            mProps.Descripcion = rs!Descripcion
          Set recuperar = rs
            mbooEsNuevo = False
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

