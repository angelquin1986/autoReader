VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Configuracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private mbooModificado As Boolean
Private ConfigCols(1 To 3) As String * 32 ' Antes 27

'1
Public Property Let Codigo(ix As Long, ByVal value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 1, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get Codigo(ix As Long) As Boolean
    Codigo = (Mid$(ConfigCols(ix), 1, 1) = "S")
End Property

'2
Public Property Let CodAlterno(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 2, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get CodAlterno(ix As Long) As Boolean
    CodAlterno = (Mid$(ConfigCols(ix), 2, 1) = "S")
End Property

'3
Public Property Let Descripcion(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 3, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get Descripcion(ix As Long) As Boolean
    Descripcion = (Mid$(ConfigCols(ix), 3, 1) = "S")
End Property

'4
Public Property Let Existencia(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 4, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get Existencia(ix As Long) As Boolean
    Existencia = (Mid$(ConfigCols(ix), 4, 1) = "S")
End Property

'5
Public Property Let CostoUnitario(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 5, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get CostoUnitario(ix As Long) As Boolean
    CostoUnitario = (Mid$(ConfigCols(ix), 5, 1) = "S")
End Property

'6
Public Property Let Utilidad_1A(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 6, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get Utilidad_1A(ix As Long) As Boolean
    Utilidad_1A = (Mid$(ConfigCols(ix), 6, 1) = "S")
End Property

'7
Public Property Let PrecioU_1A(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 7, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get PrecioU_1A(ix As Long) As Boolean
    PrecioU_1A = (Mid$(ConfigCols(ix), 7, 1) = "S")
End Property

'8
Public Property Let Utilidad_1N(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 8, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get Utilidad_1N(ix As Long) As Boolean
    Utilidad_1N = (Mid$(ConfigCols(ix), 8, 1) = "S")
End Property

'9
Public Property Let PrecioU_1N(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 9, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get PrecioU_1N(ix As Long) As Boolean
    PrecioU_1N = (Mid$(ConfigCols(ix), 9, 1) = "S")
End Property

'10
Public Property Let PrecioU_1N_IVA(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 10, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get PrecioU_1N_IVA(ix As Long) As Boolean
    PrecioU_1N_IVA = (Mid$(ConfigCols(ix), 10, 1) = "S")
End Property

'11
Public Property Let Utilidad_2A(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 11, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get Utilidad_2A(ix As Long) As Boolean
    Utilidad_2A = (Mid$(ConfigCols(ix), 11, 1) = "S")
End Property

'12
Public Property Let PrecioU_2A(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 12, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get PrecioU_2A(ix As Long) As Boolean
    PrecioU_2A = (Mid$(ConfigCols(ix), 12, 1) = "S")
End Property

'13
Public Property Let Utilidad_2N(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 13, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get Utilidad_2N(ix As Long) As Boolean
    Utilidad_2N = (Mid$(ConfigCols(ix), 13, 1) = "S")
End Property

'14
Public Property Let PrecioU_2N(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 14, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get PrecioU_2N(ix As Long) As Boolean
    PrecioU_2N = (Mid$(ConfigCols(ix), 14, 1) = "S")
End Property

'15
Public Property Let PrecioU_2N_IVA(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 15, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get PrecioU_2N_IVA(ix As Long) As Boolean
    PrecioU_2N_IVA = (Mid$(ConfigCols(ix), 15, 1) = "S")
End Property

'16
Public Property Let Utilidad_3A(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 16, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get Utilidad_3A(ix As Long) As Boolean
    Utilidad_3A = (Mid$(ConfigCols(ix), 16, 1) = "S")
End Property

'17
Public Property Let PrecioU_3A(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 17, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get PrecioU_3A(ix As Long) As Boolean
    PrecioU_3A = (Mid$(ConfigCols(ix), 17, 1) = "S")
End Property

'18
Public Property Let Utilidad_3N(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 18, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get Utilidad_3N(ix As Long) As Boolean
    Utilidad_3N = (Mid$(ConfigCols(ix), 18, 1) = "S")
End Property

'19
Public Property Let PrecioU_3N(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 19, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get PrecioU_3N(ix As Long) As Boolean
    PrecioU_3N = (Mid$(ConfigCols(ix), 19, 1) = "S")
End Property

'20
Public Property Let PrecioU_3N_IVA(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 20, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get PrecioU_3N_IVA(ix As Long) As Boolean
    PrecioU_3N_IVA = (Mid$(ConfigCols(ix), 20, 1) = "S")
End Property

'21
Public Property Let Utilidad_4A(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 21, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get Utilidad_4A(ix As Long) As Boolean
    Utilidad_4A = (Mid$(ConfigCols(ix), 21, 1) = "S")
End Property

'22
Public Property Let PrecioU_4A(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 22, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get PrecioU_4A(ix As Long) As Boolean
    PrecioU_4A = (Mid$(ConfigCols(ix), 22, 1) = "S")
End Property

'23
Public Property Let Utilidad_4N(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 23, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get Utilidad_4N(ix As Long) As Boolean
    Utilidad_4N = (Mid$(ConfigCols(ix), 23, 1) = "S")
End Property

'24
Public Property Let PrecioU_4N(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 24, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get PrecioU_4N(ix As Long) As Boolean
    PrecioU_4N = (Mid$(ConfigCols(ix), 24, 1) = "S")
End Property

'25
Public Property Let PrecioU_4N_IVA(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 25, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get PrecioU_4N_IVA(ix As Long) As Boolean
    PrecioU_4N_IVA = (Mid$(ConfigCols(ix), 25, 1) = "S")
End Property

'AUC 29/11/07
Public Property Let Utilidad_5A(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 26, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get Utilidad_5A(ix As Long) As Boolean
    Utilidad_5A = (Mid$(ConfigCols(ix), 26, 1) = "S")
End Property

'27
Public Property Let PrecioU_5A(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 27, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get PrecioU_5A(ix As Long) As Boolean
    PrecioU_5A = (Mid$(ConfigCols(ix), 27, 1) = "S")
End Property

'28
Public Property Let Utilidad_5N(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 28, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get Utilidad_5N(ix As Long) As Boolean
    Utilidad_5N = (Mid$(ConfigCols(ix), 28, 1) = "S")
End Property

'29
Public Property Let PrecioU_5N(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 29, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get PrecioU_5N(ix As Long) As Boolean
    PrecioU_5N = (Mid$(ConfigCols(ix), 29, 1) = "S")
End Property

'30
Public Property Let PrecioU_5N_IVA(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 30, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get PrecioU_5N_IVA(ix As Long) As Boolean
    PrecioU_5N_IVA = (Mid$(ConfigCols(ix), 30, 1) = "S")
End Property


'*** METODOS
Public Sub RecuperarConfigCols()
    ConfigCols(1) = GetSetting(APPNAME, App.Title, "ConfigColsVisible", "SNSSSSSSSSNNNNNNNNNNNNNNNNN")
    ConfigCols(2) = GetSetting(APPNAME, App.Title, "ConfigColsEditable", "NNNNNNNSSSNNNNNNNNNNNNNNNNN")
    'ConfigCols(3) = GetSetting(APPNAME, App.Title, "ConfigColsImprimir", "SNSSSSSSSSNNNNNNNNNNNNNNNNN")
End Sub

Public Sub GrabarConfigCols()
    If mbooModificado Then
        SaveSetting APPNAME, App.Title, "ConfigColsVisible", ConfigCols(1)
        SaveSetting APPNAME, App.Title, "ConfigColsEditable", ConfigCols(2)
        'SaveSetting APPNAME, App.Title, "ConfigColsImprimir", ConfigCols(3)
        RecuperarConfigCols
    End If
End Sub
'AUC precio 6
Public Property Let Utilidad_6A(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 31, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get Utilidad_6A(ix As Long) As Boolean
    Utilidad_6A = (Mid$(ConfigCols(ix), 31, 1) = "S")
End Property

Public Property Let PrecioU_6A(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 32, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get PrecioU_6A(ix As Long) As Boolean
    PrecioU_6A = (Mid$(ConfigCols(ix), 32, 1) = "S")
End Property

'28
Public Property Let Utilidad_6N(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 33, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get Utilidad_6N(ix As Long) As Boolean
    Utilidad_6N = (Mid$(ConfigCols(ix), 33, 1) = "S")
End Property

'29
Public Property Let PrecioU_6N(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 34, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get PrecioU_6N(ix As Long) As Boolean
    PrecioU_6N = (Mid$(ConfigCols(ix), 34, 1) = "S")
End Property

'30
Public Property Let PrecioU_6N_IVA(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 35, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get PrecioU_6N_IVA(ix As Long) As Boolean
    PrecioU_6N_IVA = (Mid$(ConfigCols(ix), 35, 1) = "S")
End Property

'31
Public Property Let Moneda(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 36, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get Moneda(ix As Long) As Boolean
    Moneda = (Mid$(ConfigCols(ix), 36, 1) = "S")
End Property

'32
Public Property Let NEtiqCodBar(ix As Long, value As Boolean)
    If (ix <= 3 And ix >= 1) Then
        Mid$(ConfigCols(ix), 37, 1) = IIf(value, "S", "N")
        mbooModificado = True
    End If
End Property

Public Property Get NEtiqCodBar(ix As Long) As Boolean
    NEtiqCodBar = (Mid$(ConfigCols(ix), 37, 1) = "S")
End Property

