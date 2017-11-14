VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colAsientoTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"AsientoTemp"
Attribute VB_Ext_KEY = "Member0" ,"AsientoTemp"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mcolAsientoTemp As Collection


Public Function Add(IdCuenta As Long, Valor As Currency) As AsientoTemp
    Dim obj As AsientoTemp, band As Boolean
    
    If Valor = 0 Then Exit Function
    
    'Busca dentro de la colección si ya existe la misma cuenta
    For Each obj In mcolAsientoTemp
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
      mcolAsientoTemp.Add item:=obj
    End If
    Set Add = obj
    Set obj = Nothing
End Function

Public Property Get item(index As Variant) As AsientoTemp
Attribute item.VB_UserMemId = 0
  Set item = mcolAsientoTemp(index)
End Property




Public Property Get Count() As Long
  Count = mcolAsientoTemp.Count
End Property


Public Sub Remove(index As Long)
  mcolAsientoTemp.Remove index
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'Esta propiedad permite enumerar la colección con sintaxis For...Each
    Set NewEnum = mcolAsientoTemp.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mcolAsientoTemp = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcolAsientoTemp = Nothing
End Sub


Public Function AddG(IdCuenta As Long, Valor As Currency, idGasto As Long) As AsientoTemp
    Dim obj As AsientoTemp, band As Boolean
    
    If Valor = 0 Then Exit Function
    
    'Busca dentro de la colección si ya existe la misma cuenta
    For Each obj In mcolAsientoTemp
      'Si ya existe el mismo codigo de cuenta, se suma
      If obj.IdCuenta = IdCuenta And obj.idGasto = idGasto Then
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
      obj.idGasto = idGasto
      mcolAsientoTemp.Add item:=obj
    End If
    Set AddG = obj
    Set obj = Nothing
End Function
Public Function AddD(IdCuenta As Long, Valor As Double) As AsientoTemp
    Dim obj As AsientoTemp, band As Boolean
    
    If Valor = 0 Then Exit Function
    
    'Busca dentro de la colección si ya existe la misma cuenta
    For Each obj In mcolAsientoTemp
      'Si ya existe el mismo codigo de cuenta, se suma
      If obj.IdCuenta = IdCuenta Then
          obj.ValorDou = obj.ValorDou + Valor
          band = True
          Exit For
      End If
    Next obj
    
    'Si no encuentra, crea nuevo
    If Not band Then
      Set obj = New AsientoTemp
      obj.IdCuenta = IdCuenta
      obj.ValorDou = Valor
      mcolAsientoTemp.Add item:=obj
    End If
    Set AddD = obj
    Set obj = Nothing
End Function
Public Function AddDes(IdCuenta As Long, Valor As Double, bandDes As Boolean, bandDebe As Boolean, bandHaber As Boolean) As AsientoTemp
    Dim obj As AsientoTemp, band As Boolean
    If Valor = 0 Then Exit Function
    'Busca dentro de la colección si ya existe la misma cuenta
    For Each obj In mcolAsientoTemp
      'Si ya existe el mismo codigo de cuenta, se suma
      If obj.IdCuenta = IdCuenta Then
          obj.ValorDou = obj.ValorDou + Valor
          band = True
          Exit For
      End If
    Next obj
    'Si no encuentra, crea nuevo
    If Not band Then
      Set obj = New AsientoTemp
      obj.IdCuenta = IdCuenta
      obj.ValorDou = Valor
      obj.bandDes = bandDes
      obj.bandDebe = bandDebe
      obj.bandHaber = bandHaber
      mcolAsientoTemp.Add item:=obj
    End If
    Set AddDes = obj
    Set obj = Nothing
End Function

