VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colAsientoPresupTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolAsientoPresupTemp As Collection


Public Function Add(IdCuenta As Long, Valor As Currency) As AsientoTemp
    Dim obj As AsientoTemp, band As Boolean
    
    If Valor = 0 Then Exit Function
    
    'Busca dentro de la colección si ya existe la misma cuenta
    For Each obj In mcolAsientoPresupTemp
      'Si ya existe el mismo codigo de cuenta, se suma
      If obj.IdCuenta = IdCuenta Then
          obj.Valor = obj.Valor + Valor
          band = True
          Exit For
      End If
    Next obj
    
    'Si no encuentra, crea nuevo
    If Not band Then
      Set obj = New AsientoTemp
      obj.IdCuenta = IdCuenta
      obj.Valor = Valor
      mcolAsientoPresupTemp.Add item:=obj
    End If
    Set Add = obj
    Set obj = Nothing
End Function

Public Property Get item(index As Variant) As AsientoTemp
  Set item = mcolAsientoPresupTemp(index)
End Property




Public Property Get Count() As Long
  Count = mcolAsientoPresupTemp.Count
End Property


Public Sub Remove(index As Long)
  mcolAsientoPresupTemp.Remove index
End Sub


Public Property Get NewEnum() As IUnknown
    'Esta propiedad permite enumerar la colección con sintaxis For...Each
    Set NewEnum = mcolAsientoPresupTemp.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mcolAsientoPresupTemp = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolAsientoPresupTemp = Nothing
End Sub

