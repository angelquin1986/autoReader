VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AFInventarioDetalleC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit




Private Type T_PROPS
    IdProvCli As Long
    CodEmpleado As String
    Cantidad As String
        
    objAFInventario As AFInventario          'Referencia al objeto padre
End Type

Private mProps As T_PROPS


Friend Property Set AFInventario(obj As AFInventario)
    Set mProps.objAFInventario = obj
End Property

Public Property Get AFInventario() As AFInventario
    Set AFInventario = mProps.objAFInventario
End Property



Friend Property Let IdProvCli(ByVal value As Long)
    mProps.IdProvCli = value
End Property

Public Property Get IdProvCli() As Long
    IdProvCli = mProps.IdProvCli
End Property

Friend Property Let CodEmpleado(ByVal value As String)
    mProps.CodEmpleado = value
End Property

Public Property Get CodEmpleado() As String
    CodEmpleado = Trim$(mProps.CodEmpleado)
End Property

Friend Property Let Cantidad(ByVal value As Currency)
    mProps.Cantidad = value
End Property

Public Property Get Cantidad() As Currency
    Cantidad = mProps.Cantidad
End Property

Private Sub Class_Terminate()
    Set mProps.objAFInventario = Nothing
End Sub




