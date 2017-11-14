VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AsientoTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mlngIdCuenta As String
Private mcurValor As Currency
Private mDouValor As Double
Private mlngIdGasto As Long
Private mbandDes As Boolean
Private mbandDebe As Boolean
Private mbandHaber As Boolean

Public Property Let IdCuenta(ByVal value As Long)
    mlngIdCuenta = value
End Property

Public Property Get IdCuenta() As Long
    IdCuenta = mlngIdCuenta
End Property

Public Property Let Valor(ByVal value As Currency)
    mcurValor = value
End Property

Public Property Get Valor() As Currency
    Valor = mcurValor
End Property
Public Property Let idGasto(ByVal value As Long)
    mlngIdGasto = value
End Property

Public Property Get idGasto() As Long
    idGasto = mlngIdGasto
End Property

Public Property Let ValorDou(ByVal value As Double)
    mDouValor = value
End Property

Public Property Get ValorDou() As Double
    ValorDou = mDouValor
End Property
Public Property Let bandDes(ByVal value As Boolean)
    mbandDes = value
End Property

Public Property Get bandDes() As Boolean
    bandDes = mbandDes
End Property

Public Property Let bandDebe(ByVal value As Boolean)
    mbandDebe = value
End Property

Public Property Get bandDebe() As Boolean
    bandDebe = mbandDebe
End Property

Public Property Let bandHaber(ByVal value As Boolean)
    mbandHaber = value
End Property

Public Property Get bandHaber() As Boolean
    bandHaber = mbandHaber
End Property

