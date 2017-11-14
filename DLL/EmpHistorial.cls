VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmpHistorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type T_Props
    IdHistorial As Long
    IdEmpleado As String
    Descripcion As String
    Sueldo As Currency
    fecha As Date
    CodigoLosep As String
    objPersonal As Personal
End Type
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean
Private mProps As T_Props
Friend Property Set Personal(obj As Personal)
    Set mProps.objPersonal = obj
End Property
Public Property Get Personal() As Personal
    Set Personal = mProps.objPersonal
End Property

Public Property Get IdHistorial() As Long
    IdHistorial = mProps.IdHistorial
End Property
Public Property Let IdHistorial(ByVal Value As Long)
    mProps.IdHistorial = Value
End Property

Public Property Get IdEmpleado() As String
    IdEmpleado = Trim$(mProps.IdEmpleado)
End Property

Public Property Let IdEmpleado(ByVal Value As String)
    mProps.IdEmpleado = Value
    mbooModificado = True
End Property


Private Sub Class_Terminate()
    Set mProps.objPersonal = Nothing
End Sub

Public Property Get Descripcion() As String 'Propiedad  solo de lectura
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let Descripcion(ByVal Value As String)
    If Len(Value) > 120 Then
        Err.Raise ERR_INVALIDO, "Personal_Direccion", MSGERR_INVALIDO
    End If
     mProps.Descripcion = Value
    mbooModificado = True
End Property
 
Public Property Get fecha() As Date
    fecha = Trim$(mProps.fecha)
End Property

Public Property Let fecha(ByVal Value As Date)
    mProps.fecha = Value
    mbooModificado = True
End Property
Public Property Get Sueldo() As Currency
    Sueldo = mProps.Sueldo
End Property
Public Property Let Sueldo(ByVal Value As Currency)
    mProps.Sueldo = Value
End Property

Public Property Get CodigoLosep() As String 'Propiedad  solo de lectura
    CodigoLosep = Trim$(mProps.CodigoLosep)
End Property

Public Property Let CodigoLosep(ByVal Value As String)
    If Len(Value) > 20 Then
        Err.Raise ERR_INVALIDO, "Personal_CodigoLosep", MSGERR_INVALIDO
    End If
     mProps.CodigoLosep = Value
    mbooModificado = True
End Property

