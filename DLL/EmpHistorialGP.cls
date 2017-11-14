VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmpHistorialGP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type T_PROPS
    IdHistorialGP As Long
    IdEmpleado As String
    GP As Currency
    fecha As Date
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

Public Property Get IdHistorialGP() As Long
    IdHistorialGP = mProps.IdHistorialGP
End Property
Public Property Let IdHistorialGP(ByVal value As Long)
    mProps.IdHistorialGP = value
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

 
Public Property Get fecha() As Date
    fecha = Trim$(mProps.fecha)
End Property

Public Property Let fecha(ByVal value As Date)
    mProps.fecha = value
    mbooModificado = True
End Property
Public Property Get GP() As Currency
    GP = mProps.GP
End Property
Public Property Let GP(ByVal value As Currency)
    mProps.GP = value
End Property

