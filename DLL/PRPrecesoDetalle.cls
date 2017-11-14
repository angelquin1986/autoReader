VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PRPrecesoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit




Private Type T_PROPS
    IdBodega As Long
    CodBodega As String
    Cantidad As String
        
    objIVInventario As PRProceso          'Referencia al objeto padre
End Type

Private mProps As T_PROPS


Friend Property Set PRProceso(obj As PRProceso)
    Set mProps.objIVInventario = obj
End Property

Public Property Get PRProceso() As PRProceso
    Set PRProceso = mProps.objIVInventario
End Property



Friend Property Let IdBodega(ByVal value As Long)
    mProps.IdBodega = value
End Property

Public Property Get IdBodega() As Long
    IdBodega = mProps.IdBodega
End Property

Friend Property Let CodBodega(ByVal value As String)
    mProps.CodBodega = value
End Property

Public Property Get CodBodega() As String
    CodBodega = Trim$(mProps.CodBodega)
End Property

Friend Property Let Cantidad(ByVal value As Currency)
    mProps.Cantidad = value
End Property

Public Property Get Cantidad() As Currency
    Cantidad = mProps.Cantidad
End Property

Private Sub Class_Terminate()
    Set mProps.objIVInventario = Nothing
End Sub



