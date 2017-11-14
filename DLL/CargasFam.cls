VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CargasFam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type T_PROPS
    IdCarga As Long
    IdEmpleado As String
    Nombre As String
    Sexo As Byte
    FechaNaci As Date
    Edad As String
    Observacion As String
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

Public Property Get IdCarga() As Long
    IdCarga = mProps.IdCarga
End Property
Public Property Let IdCarga(ByVal value As Long)
    mProps.IdCarga = value
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

Public Property Get Nombre() As String 'Propiedad  solo de lectura
    Nombre = Trim$(mProps.Nombre)
End Property

Public Property Let Nombre(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Cargo_nombre", MSGERR_INVALIDO
    End If
     mProps.Nombre = value
    mbooModificado = True
End Property

Public Property Get FechaNaci() As String
    FechaNaci = Trim$(mProps.FechaNaci)
End Property

Public Property Let FechaNaci(ByVal value As String)
    mProps.FechaNaci = value
    mbooModificado = True
End Property

Public Property Get Edad() As String
    Edad = mProps.Edad
End Property
Public Property Let Edad(ByVal value As String)
    mProps.Edad = value
End Property

Public Property Get Observacion() As String
    Observacion = Trim$(mProps.Observacion)
End Property

Public Property Let Observacion(ByVal value As String)
    mProps.Observacion = value
    mbooModificado = True
End Property

Public Property Let Sexo(ByVal value As Byte)
    '0=Femenino, 1=Masculino
    If (Not IsNumeric(value)) Or (Val(value) < 0 Or Val(value) > 1) Then Err.Raise ERR_INVALIDO, "Personal.Sexo", MSGERR_INVALIDO
    mProps.Sexo = value
    mbooModificado = True
End Property

Public Property Get Sexo() As Byte
    Sexo = mProps.Sexo
End Property

