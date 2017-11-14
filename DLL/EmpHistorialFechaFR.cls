VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmpHistorialFechaFR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type T_PROPS
    IdHistorialFFR As Long
    IdEmpleado As String
    fechaOld As Date
    Descripcion As String
    fechagrabado As Date
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

Public Property Get IdHistorialFFR() As Long
    IdHistorialFFR = mProps.IdHistorialFFR
End Property
Public Property Let IdHistorialFFR(ByVal value As Long)
    mProps.IdHistorialFFR = value
End Property

Public Property Get IdEmpleado() As String
    IdEmpleado = Trim$(mProps.IdEmpleado)
End Property

Public Property Let IdEmpleado(ByVal value As String)
    mProps.IdEmpleado = value
    mbooModificado = True
End Property

Private Sub Class_Terminate()
    Set mProps.objPersonal = Nothing
End Sub
 
Public Property Get fechaOld() As Date
    fechaOld = Trim$(mProps.fechaOld)
End Property

Public Property Let fechaOld(ByVal value As Date)
    mProps.fechaOld = value
    mbooModificado = True
End Property

Public Property Get fechagrabado() As Date
    fechagrabado = Trim$(mProps.fechagrabado)
End Property

Public Property Let fechagrabado(ByVal value As Date)
    mProps.fechagrabado = value
    mbooModificado = True
End Property
Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let Descripcion(ByVal value As String)
    mProps.Descripcion = value
    mbooModificado = True
End Property


