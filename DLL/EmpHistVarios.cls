VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmpHistVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type T_PROPS
    IdHistorial As Long
    IdEmpleado As Long
    Descripcion As String
    FechaGradado As Date
    objPersonal As Personal
End Type
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean
Private mProps As T_PROPS
Friend Property Set Personal(obj As Personal)
    Set mProps.objPersonal = obj
End Property
Public Property Get Personal() As Personal
    Set Personal = mProps.objPersonal
End Property

Public Property Get IdHistorial() As Long
    IdHistorial = mProps.IdHistorial
End Property
Public Property Let IdHistorial(ByVal value As Long)
    mProps.IdHistorial = value
End Property
Public Property Get IdEmpleado() As Long
    IdEmpleado = mProps.IdEmpleado
End Property
Public Property Let IdEmpleado(ByVal value As Long)
    mProps.IdEmpleado = value
End Property

Private Sub Class_Terminate()
    Set mProps.objPersonal = Nothing
End Sub

Public Property Get Descripcion() As String 'Propiedad  solo de lectura
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 120 Then
        Err.Raise ERR_INVALIDO, "HistVarios_Direccion", MSGERR_INVALIDO
    End If
     mProps.Descripcion = value
    mbooModificado = True
End Property
 'key puede ser IdInventario  o CodInventario
Public Sub Recuperar(key As Long)
    Dim sql As String, i As Integer
    Dim rs As Recordset
        sql = "Select IdHistorial,Descripcion,fechagrabado  FROM EmpHistVarios WHERE idhistorial= " & key
    
    Set rs = mProps.objPersonal.Empresa.OpenRecordset(sql)
    If rs.EOF Then
        Err.Raise Err.Number, "EmpHistorialEmp", "No existe empleado"
    Else
        mProps.IdHistorial = rs!IdHistorial
        mProps.IdEmpleado = rs!IdEmpleado
        mProps.Descripcion = rs!Descripcion
        mProps.FechaGradado = rs!FechaGrabado
    End If
End Sub


Public Property Get FechaGrabado() As String
    FechaGrabado = Trim$(mProps.FechaGradado)
End Property

Public Property Let FechaGrabado(ByVal value As String)
    mProps.FechaGradado = value
    mbooModificado = True
End Property

