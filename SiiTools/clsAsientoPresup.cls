VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAsientoPresup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type T_PROPS
    IdCuenta As Long
    Debe As Currency
    Haber As Currency
    Comparado As Boolean
End Type

Private mProps As T_PROPS

Public Property Let IdCuenta(ByVal value As Long)
    mProps.IdCuenta = value
End Property

Public Property Get IdCuenta() As Long
    IdCuenta = mProps.IdCuenta
End Property

Public Property Let Debe(ByVal value As Currency)
    mProps.Debe = value
End Property

Public Property Get Debe() As Currency
    Debe = mProps.Debe
End Property

Public Property Let Haber(ByVal value As Currency)
    mProps.Haber = value
End Property

Public Property Get Haber() As Currency
    Haber = mProps.Haber
End Property

Public Property Let Comparado(ByVal value As Boolean)
    mProps.Comparado = value
End Property

Public Property Get Comparado() As Boolean
    Comparado = mProps.Comparado
End Property


Private Sub Class_Initialize()
    mProps.Comparado = False
End Sub

