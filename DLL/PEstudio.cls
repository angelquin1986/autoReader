VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PEstudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'.FormatString = "<Nombre|<Año Graduacion|<Titulo|<Observacion"
Private Type T_PROPS
    IdEstudio As Long
    IdEmpleado As String
    Descripcion As String
    Tipo As Byte
    anoGrado As Date
    Titulo As String
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

Public Property Get IdEstudio() As Long
    IdEstudio = mProps.IdEstudio
End Property
Public Property Let IdEstudio(ByVal value As Long)
    mProps.IdEstudio = value
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

Public Property Get Descripcion() As String 'Propiedad  solo de lectura
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Estudios.Descripcion", MSGERR_INVALIDO
    End If
     mProps.Descripcion = value
    mbooModificado = True
End Property
Public Property Get anoGrado() As Date
    anoGrado = Trim$(mProps.anoGrado)
End Property

Public Property Let anoGrado(ByVal value As Date)
    mProps.anoGrado = value
    mbooModificado = True
End Property


Public Property Get Observacion() As String
    Observacion = Trim$(mProps.Observacion)
End Property

Public Property Let Observacion(ByVal value As String)
    mProps.Observacion = value
    mbooModificado = True
End Property

Public Property Get Titulo() As String
    Titulo = Trim$(mProps.Titulo)
End Property

Public Property Let Titulo(ByVal value As String)
    mProps.Titulo = value
    mbooModificado = True
End Property

Public Property Let Tipo(ByVal value As Byte)
    
    If (Not IsNumeric(value)) Or (Val(value) < 0 Or Val(value) > 3) Then Err.Raise ERR_INVALIDO, "Estudio.Tipo", MSGERR_INVALIDO
    mProps.Tipo = value
    mbooModificado = True
End Property

Public Property Get Tipo() As Byte
    Tipo = mProps.Tipo
End Property

