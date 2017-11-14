VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.UserControl IVRPVT 
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5190
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   2205
   ScaleWidth      =   5190
   Begin VSFlex7Ctl.VSFlexGrid grd 
      Align           =   1  'Align Top
      Height          =   1452
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5196
      _cx             =   9165
      _cy             =   2561
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   12
      FixedRows       =   2
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   0
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   1
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   8421504
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "IVRPVT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Ubicación de columnas
Private Const COL_NUMFILA = 0
Private Const COL_CODRECARGO = 1
Private Const COL_SIGNO = 2
Private Const COL_PORCENT = 3
Private Const COL_VALOR = 4
Private Const COL_CALCULO = 5
Private Const COL_DESC = 6
Private Const COL_MODIFICABLE = 7
Private Const COL_ORIGEN = 8
Private Const COL_PRORRAT = 9
Private Const COL_AFECTAIVA = 10
Private Const COL_SELECCIONABLE = 11

'Property Variables:
Private WithEvents mobjGNComp As GNComprobante
Attribute mobjGNComp.VB_VarHelpID = -1
Private mbooVisualizando As Boolean

'Event Declarations:
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)

Event DespuesdeEditarGrd()
Dim ANCHO_COLS(0 To 26) As Long
'Variable Local
Private mIDPcGrupo1 As Long
Private mIDPcGrupo2 As Long
Private mIDPcGrupo3 As Long
Private Total As Currency



Private Sub ConfigCols()
    Dim i As Integer
    With grd
        .FormatString = "^#|<Código|^Signo|^%|>Valor|>Suma|<Descripción" & _
                "|Modific|Orígen|Prorrat|AfectaIvaItem|Selecc"
        GetColsWidth
        .ColWidth(COL_NUMFILA) = 500
        .ColWidth(COL_CODRECARGO) = 800             'Cod.Recargo
        .ColWidth(COL_SIGNO) = 700                  'Signo
        .ColWidth(COL_PORCENT) = 900                'Porcenatje
        .ColWidth(COL_VALOR) = COLANCHO_CUR * 0.9       'Valor antes * 1.2
        .ColWidth(COL_CALCULO) = COLANCHO_CUR * 0.9     'Calculo antes * 1.2
        .ColWidth(COL_DESC) = 1400                  'Descripción antes 2000
        .ColWidth(COL_MODIFICABLE) = 500
        .ColWidth(COL_ORIGEN) = 500
        .ColWidth(COL_PRORRAT) = 500
        .ColWidth(COL_AFECTAIVA) = 500

        'No modificables
        .ColData(COL_CODRECARGO) = -1
        .ColData(COL_SIGNO) = -1
        .ColData(COL_CALCULO) = -1                    'calculo
        .ColData(COL_DESC) = -1                     'Descripcion de item
        .ColData(COL_MODIFICABLE) = -1
        .ColData(COL_ORIGEN) = -1
        .ColData(COL_PRORRAT) = -1
        .ColData(COL_AFECTAIVA) = -1
        .ColData(COL_SELECCIONABLE) = -1
        
        .ColDataType(COL_CODRECARGO) = flexDTString
        .ColDataType(COL_SIGNO) = flexDTString
        .ColDataType(COL_PORCENT) = flexDTSingle
        .ColDataType(COL_VALOR) = flexDTCurrency
        .ColDataType(COL_CALCULO) = flexDTCurrency
        .ColDataType(COL_DESC) = flexDTString
        .ColDataType(COL_MODIFICABLE) = flexDTBoolean
        .ColDataType(COL_SELECCIONABLE) = flexDTBoolean
        .ColDataType(COL_ORIGEN) = flexDTShort
        .ColDataType(COL_PRORRAT) = flexDTBoolean
        .ColDataType(COL_AFECTAIVA) = flexDTBoolean
        
        .ColHidden(COL_MODIFICABLE) = True
        .ColHidden(COL_SELECCIONABLE) = True
        .ColHidden(COL_ORIGEN) = True
        .ColHidden(COL_PRORRAT) = True
        .ColHidden(COL_AFECTAIVA) = True
        For i = 0 To .Cols - 1
            .ColWidth(i) = ANCHO_COLS(i)
        Next i

    End With
    ConfigColsFormato
End Sub

Private Sub ConfigColsFormato()
    With grd
        .ColFormat(COL_PORCENT) = "##.00"
        .ColFormat(COL_VALOR) = mobjGNComp.FormatoMoneda
        .ColFormat(COL_CALCULO) = .ColFormat(COL_VALOR)
'        .TextMatrix(1, 1) = "Subtotal"
    End With
End Sub


Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Function Refresh() As Currency
    If mobjGNComp Is Nothing Then Exit Function
    'Cuando es solo ver, deshabilita grid
    If mobjGNComp.SoloVer Then
        grd.Editable = flexEDNone
        Total = VisualizaTotal()
        Refresh = Total
        Exit Function
    Else
        grd.Editable = flexEDKbdMouse
    End If
     Total = VisualizaTotal()
    ConfigColsFormato      'Llama esta para que actualice el formato de columnas cuando cambia moneda
    Refresh = Total
End Function

Private Sub grd_Click()
    RaiseEvent Click
End Sub

Private Sub grd_DblClick()
    RaiseEvent DblClick
End Sub

'*** MAKOTO 12/dic/00 Agregado
Private Sub grd_GotFocus()
    FlexGridGotFocusColor grd
End Sub

Private Sub grd_LostFocus()
    FlexGridGotFocusColor grd
End Sub

Private Sub grd_KeyDown(KeyCode As Integer, Shift As Integer)
    If mobjGNComp Is Nothing Then Exit Sub
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub


Private Sub grd_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub grd_KeyPressEdit(ByVal Row As Long, ByVal col As Long, KeyAscii As Integer)
    '*** MAKOTO 03/oct/2000
    ValidarTeclaFlexGrid grd, Row, col, KeyAscii
End Sub


Private Sub grd_AfterEdit(ByVal Row As Long, ByVal col As Long)
    On Error GoTo ErrTrap
    VisualizaTotal
    RaiseEvent DespuesdeEditarGrd
'    MueveColumna
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub grd_CellChanged(ByVal Row As Long, ByVal col As Long)
    '*** MAKOTO 29/ene/01 Agregado.
    ConfigColsFormato
    FlexGridRedondear grd, Row, col
End Sub


Private Function VisualizaTotal() As Currency
    Dim t As Currency, tdesc As Currency
    Dim i As Long, v As Currency, ix As Long
    Dim obj As IVKardexRecargo
    Dim p_antes As Currency
    Dim pc As PCProvCli, BandEmpPub As Boolean
    BandEmpPub = False
    If (Not mobjGNComp.SoloVer) And (Not mbooVisualizando) Then
        p_antes = mobjGNComp.IVRecargoTotal(False, True)    'total de recargos prorrateados
    End If
    t = Abs(mobjGNComp.IVKardexTotal(False))    'Total NETO sin recargo prorateado
    tdesc = mobjGNComp.IVKardexDescItemTotal
    t = t - tdesc
    Dim porICE As Currency
    With grd
        'Primera fila es total de items
        .TextMatrix(1, COL_CALCULO) = t
        If mobjGNComp.GNTrans.IVOmitirIVA Then
            Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(mobjGNComp.IdClienteRef)
            If pc Is Nothing Then
                BandEmpPub = False
            Else
                If pc.BandEmpresaPublica Then
                    BandEmpPub = True
                End If
            End If
        End If
        
        For i = .FixedRows To .Rows - 1
            If Not .IsSubtotal(i) Then
                v = 0
                If mobjGNComp.Empresa.GNOpcion.ObtenerValor("OmitirIVA") = "1" And mobjGNComp.GNTrans.IVOmitirIVA And .TextMatrix(i, COL_CODRECARGO) = mobjGNComp.Empresa.GNOpcion.ObtenerValor("RecDescParaIVA") And BandEmpPub Then
                Else
                    Select Case .ValueMatrix(i, COL_ORIGEN)
                    'Si es iva de item
                    Case REC_IVAITEM
                        'Coge valor total de iva de cada item
                        v = mobjGNComp.IVKardexIVAItemTotal
                        .TextMatrix(i, COL_PORCENT) = ""
                        
                    'Si es recargo/descuento a la fila anterior
                    Case REC_SUMA
                        'Si está ingresado el porcentaje, calcula el valor según porcentaje
                        If (Len(.TextMatrix(i, COL_PORCENT)) = 0 Or .ValueMatrix(i, COL_PORCENT) = 0) _
                                                         And CBool(.ValueMatrix(i, COL_SELECCIONABLE)) = False Then        ' esta  en blanco
    '                        v = .ValueMatrix(i, COL_VALOR)
                            v = MiCCur(.Cell(flexcpTextDisplay, i, COL_VALOR))      '*** MAKOTO 29/ene/01 Mod.
                        'Si no, coge el valor fijo
                        Else
                                'Calcula en base a la suma de la fila anterior
                                v = t * .ValueMatrix(i, COL_PORCENT) / 100
                        End If
                        
                    'Si es recargo/descuento al total neto de items
                    Case REC_TOTAL
                        'Si está ingresado el porcentaje, calcula el valor según porcentaje
                         If (Len(.TextMatrix(i, COL_PORCENT)) = 0 Or .ValueMatrix(i, COL_PORCENT) = 0) _
                                                    And CBool(.ValueMatrix(i, COL_SELECCIONABLE)) = False Then     ' esta  en blanco
                            'jeaa 28-07-04 para que se actualice el valor si en 0
                            If .TextMatrix(i, COL_PORCENT) = "0.00" Then
                                v = MiCCur(0)      '*** MAKOTO 29/ene/01 Mod.
                            Else
                                v = MiCCur(.Cell(flexcpTextDisplay, i, COL_VALOR))      '*** MAKOTO 29/ene/01 Mod.
                            End If
                        'Si no, coge el valor fijo
                        Else
                            'Calcula en base a la suma de total REAL de items
                            v = Abs(mobjGNComp.IVKardexTotal(False)) * .ValueMatrix(i, COL_PORCENT) / 100
                        End If
                        
                    'Si es una fila de subtotal
                    Case REC_SUBTOTAL
                        'Visualiza subtotal desde la primera fila hasta la fila anterior
                        .TextMatrix(i, COL_SIGNO) = "-----"
                    
                    'Si es recargo de item
                    Case REC_RECITEM  '***Agregado. Angel. 29/jul/2004
                        'Coge valor total de recargo de cada item
                        v = mobjGNComp.IVKardexRecargoItemTotal
                        .TextMatrix(i, COL_PORCENT) = ""
                    Case REC_ICEITEM  '***Agregado. JEAA 21/07/2006
                        'Coge valor de ice de cada item
                        
                        v = mobjGNComp.IVKardexICEItem(porICE)
                        .TextMatrix(i, COL_PORCENT) = porICE
                        
                    Case REC_SUMAIVAITEM
                        'Coge valor DE LA SUMA total DE los item que tienen IVA
                        v = mobjGNComp.IVKardexSumaIVAItemTotal
                        .TextMatrix(i, COL_PORCENT) = ""
    
                    'Si es recargo/descuento a la fila específica
                    Case Is > 0
                        'Si está ingresado el porcentaje, calcula el valor según porcentaje
                        If (Len(.TextMatrix(i, COL_PORCENT)) = 0 Or .ValueMatrix(i, COL_PORCENT) = 0) _
                                               And CBool(.ValueMatrix(i, COL_SELECCIONABLE)) = False Then
                        
                            v = MiCCur(.Cell(flexcpTextDisplay, i, COL_VALOR))      '*** MAKOTO 29/ene/01 Mod.
                        Else
                                'Calcula con el valor de la fila indicada como origen
                                ix = .ValueMatrix(i, COL_ORIGEN) + 1        '*** MAKOTO 11/nov/00
                                If ix < .Rows Then
    '                                v = .ValueMatrix(ix, COL_CALCULO)           '***
                                    v = MiCCur(.Cell(flexcpTextDisplay, ix, COL_CALCULO))   '*** MAKOTO 29/ene/01 Mod.
                                    v = v * MiCCur(.Cell(flexcpTextDisplay, i, COL_PORCENT)) / 100
                                End If
                        End If
                    End Select
                End If
                'Visualiza el valor (Funciona el redondeo de FlexGrid según ColFormat)
                .TextMatrix(i, COL_VALOR) = v
                
                '*** MAKOTO 29/ene/01 Agregado. Para tomar el valor redondeado
                v = MiCCur(.Cell(flexcpTextDisplay, i, COL_VALOR))  'Obtiene el valor redondeado por FlexGrid
'                .TextMatrix(i, COL_VALOR) = v      'Visualiza de nuevo para que
'                                                    'no haya diferencia entre el valor de TextDisplay y ValueMatrix
                'Suma acumulada
                t = t + v * IIf(.TextMatrix(i, COL_SIGNO) = "-", -1, 1)
                .TextMatrix(i, COL_CALCULO) = t
                
                'Si está en modificación
                If (Not mobjGNComp.SoloVer) And (Not mbooVisualizando) Then
                    'Asigna el valor al objeto IVKardexRecargo
                    Set obj = .RowData(i)
                    '*** MAKOTO 29/ene/01 Mod.
'                    obj.Porcentaje = .ValueMatrix(i, COL_PORCENT) / 100
                    obj.porcentaje = MiCCur(.Cell(flexcpTextDisplay, i, COL_PORCENT)) / 100
                    obj.valor = v * IIf(.TextMatrix(i, COL_SIGNO) = "-", -1, 1)
                End If
            End If
        Next i
    End With
    
    If (Not mobjGNComp.SoloVer) And (Not mbooVisualizando) Then
        'Si cambió total de recargos prorrateados
        If p_antes <> mobjGNComp.IVRecargoTotal(False, True) Then
            'Prorratea los recargos que deben ser prorrateado
            mobjGNComp.ProrratearIVKardexRecargo
            Total = VisualizaTotal
        End If
    End If
    'Visualiza TOTAL GENERAL
    grd.subtotal flexSTNone, -1, COL_CALCULO, , grd.BackColorFixed, vbRed, True, " ", , True
    grd.TextMatrix(grd.Rows - 1, COL_CODRECARGO) = "TOTAL"
    grd.TextMatrix(grd.Rows - 1, COL_CALCULO) = t
    VisualizaTotal = t
End Function

Private Function CalculaSubtotal(r As Long) As Currency
    Dim i As Long, t As Currency
    
    With grd
        For i = .FixedRows To r - 1
            'Si no es fila de subtotal
            If .ValueMatrix(i, COL_ORIGEN) <> REC_SUBTOTAL Then
                t = t + .ValueMatrix(i, COL_VALOR) _
                            * IIf(.TextMatrix(i, COL_SIGNO) = "+", 1, -1)
            End If
        Next i
    End With
    CalculaSubtotal = t
End Function

Private Sub MueveColumna()
    Dim c As Long
    With grd
        If .Rows > .FixedRows Then
            For c = .col + 1 To .Cols - 1
                If .ColData(c) >= 0 And .ColWidth(c) > 0 And (Not .ColHidden(c)) Then
                    .col = c
                    Exit Sub
                End If
            Next c
    
            If .Row < .Rows - 1 Then .Row = .Row + 1
    
            For c = .FixedCols To .Cols - 1
                If .ColData(c) >= 0 And .ColWidth(c) > 0 And (Not .ColHidden(c)) Then
                    .col = c
                    Exit Sub
                End If
            Next c
        End If
    End With
End Sub

Private Sub grd_BeforeEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
    If mobjGNComp.SoloVer Then Exit Sub
    
    'Cuando es una columna no modificable
    If grd.Rows > grd.FixedRows Then
        Cancel = (grd.ColData(col) < 0) Or grd.IsSubtotal(Row) Or grd.ColHidden(col)
    Else
        Cancel = True
    End If
    If Cancel Then Exit Sub

    'Verifica si la fila es modificable o no segun el modelo de recargo
    If grd.Cell(flexcpChecked, Row, COL_MODIFICABLE) = flexUnchecked Then
        Cancel = True
        Exit Sub
    End If
        
    '***Diego   07/10/2003
    'Si la fila  es  Modificable y seleccionable no deja  editar el valor
    'solo  permite seleccionar  el pocentaje
    If grd.Cell(flexcpChecked, Row, COL_SELECCIONABLE) = flexChecked And col = COL_VALOR Then
        Cancel = True
        Exit Sub
    End If
    
    'si no es seleccionable
    grd.ComboList = ""
    If grd.Cell(flexcpChecked, Row, COL_SELECCIONABLE) = flexChecked And col = COL_PORCENT Then
        grd.ComboList = DetalleIVRecargo(grd.TextMatrix(Row, COL_CODRECARGO))
        grd.ComboSearch = flexCmbSearchLists  'OK
        If grd.ComboList = "" Then Cancel = True
    End If

    Select Case col
    Case COL_VALOR
        grd.EditMaxLength = 12
    Case COL_PORCENT
        grd.EditMaxLength = 5
    End Select
End Sub

'***Diego 10/2003
Private Function DetalleIVRecargo(ByVal codRecargo As String) As String
    'devuelve  el  detalle recargo  para  cargar  en  le  combo  de la grilla
    Dim ivr As IvRecargo, i As Long, Condicion As String
    Dim rs As Recordset, sql As String
    'obtener  el valor de numPcgrupo y IDPcGrupo  del cliente seleccionado
    'mobjGNComp.IdClienteRef
    If mIDPcGrupo1 > 0 Then
        Condicion = " (IVRecargoDetalle.IdPCGrupo = " & mIDPcGrupo1 & " AND " & _
                    "NumPCGrupo = 1)"
    End If
    
    If mIDPcGrupo2 > 0 Then
        Condicion = Condicion & " OR (IVRecargoDetalle.IdPCGrupo = " & mIDPcGrupo2 & " AND " & _
                    "NumPCGrupo = 2)"
    End If
    
    If mIDPcGrupo3 > 0 Then
        Condicion = Condicion & " OR (IVRecargoDetalle.IdPCGrupo = " & mIDPcGrupo3 & " AND " & _
                    "NumPCGrupo = 3)"
    End If
    'Sql  aplica formato 10.00
    sql = "Select LTRIM(STR(Valor,10,2)) as Valor From  IVRecargoDetalle Inner Join IvRecargo " & _
          "ON IVRecargo.Idrecargo = IVRecargoDetalle.IDRecargo " & _
          "Where CodRecargo = '" & codRecargo & "'"
     If Len(Condicion) > 0 Then 'no tiene  asignado  grupo PC
        sql = sql & " AND " & Condicion
        'Exit Function
    Else
        Exit Function
    End If
    Set rs = mobjGNComp.Empresa.OpenRecordset(sql)
    If Not rs.EOF Then DetalleIVRecargo = rs.GetString(adClipString, , vbTab, "|")
    Set rs = Nothing
End Function

Private Sub PoneNumFila()
    Dim i As Long
    With grd
        For i = .FixedRows To .Rows - 1
            If Not .IsSubtotal(i) Then
                'Pone numero de fila
                .TextMatrix(i, 0) = i - .FixedRows + 1
            End If
            
            'Si no es modificable, cambia color de fondo
            If .Cell(flexcpChecked, i, COL_MODIFICABLE) = flexUnchecked Then
'                .Select i, COL_CODRECARGO, i, .Cols - 1
'                .CellBackColor = &HC0FFFF
                .Cell(flexcpBackColor, i, COL_CODRECARGO, i, .Cols - 1) = &HC0FFFF
            End If
        Next i
        If .Rows > .FixedRows Then .col = COL_PORCENT
        If .Rows > .FixedRows Then .Row = .FixedRows
    End With
End Sub

Private Sub grd_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
    With grd
        If Len(.EditText) > 0 Then
            If Not IsNumeric(.EditText) Then
                MsgBox "Ingrese un valor numérico.", vbExclamation
                Cancel = True
                Exit Sub
            End If
        End If
        
        If col = COL_VALOR Then
            If .ValueMatrix(Row, col) <> Val(.EditText) Then
                .TextMatrix(Row, COL_PORCENT) = ""
            End If
        End If
    End With
End Sub

Public Property Get GNComprobante() As GNComprobante
    Set GNComprobante = mobjGNComp
End Property

Public Property Set GNComprobante(obj As GNComprobante)
Dim Total As Double
    Set mobjGNComp = obj
    If mobjGNComp.idTransFuente = 0 Or mobjGNComp.EsNuevo = False Then      '***Diego 13/10/2003  optimizar  tiempo
        Visualizar Total
    End If
    'Cuando es solo ver, deshabilita grid
    grd.Editable = IIf(mobjGNComp.SoloVer, flexEDNone, flexEDKbdMouse)
End Property

Public Sub Limpiar()
    Dim i As Long
    With grd
        For i = .FixedRows To .Rows - 1
            .RowData(i) = 0
        Next i
        .Rows = .FixedRows
    End With
    
End Sub

Public Sub Visualizar(ByRef Total As Double)
    Dim i As Long
    
    mbooVisualizando = True
    grd.Redraw = False
    'ConfigColsFormato
    'Visualiza los detalles que está en GNComprobante
    Set grd.DataSource = mobjGNComp.ListaIVKardexRecargo
    ConfigCols
    'Si está en modificación
    If Not mobjGNComp.SoloVer Then
        'Asigna referencia al objeto IVKardexRecargo a cada fila de grid
        With grd
            For i = 1 To mobjGNComp.CountIVKardexRecargo
                'Si existen más de lo que está en grid, no lo asigna
                If i <= .Rows - .FixedRows Then
                    If Not .IsSubtotal(i) Then
                        .RowData(i + 1) = mobjGNComp.IVKardexRecargo(i)
                    End If
                End If
            Next i
        End With
    End If
    PoneNumFila
    Total = VisualizaTotal()
    grd.Redraw = True
    grd.Refresh
    mbooVisualizando = False
End Sub

Public Sub Aceptar()
    Dim i As Long, obj As IVKardexRecargo
    On Error GoTo ErrTrap

    'Pasa los detalles al objeto GNComprobante
    With grd
        For i = .FixedRows To .Rows - 1
            If Not .IsSubtotal(i) Then
                Set obj = .RowData(i)
                obj.Orden = i - .FixedRows + 1
                obj.BandModificable = (.Cell(flexcpChecked, i, COL_MODIFICABLE) = flexChecked)
                obj.BandOrigen = .ValueMatrix(i, COL_ORIGEN)
                obj.BandProrrateado = (.Cell(flexcpChecked, i, COL_PRORRAT) = flexChecked)
                '*** MAKOTO 29/ene/01 Mod.
'                obj.Porcentaje = .ValueMatrix(i, COL_PORCENT) / 100
                obj.porcentaje = MiCCur(.Cell(flexcpTextDisplay, i, COL_PORCENT)) / 100
            End If
        Next i
        Set obj = Nothing
    End With
    Exit Sub
ErrTrap:
    Select Case Err.Number
    Case Else
        DispErr
    End Select
    Exit Sub
End Sub

Private Sub mobjGNComp_ClienteCambiado()
'    'Actualiza IvRecargo cuando  cambia  el cliente
'    Dim pc As PCProvCli
'    Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(mobjGNComp.IdClienteRef)
'    mIDPcGrupo1 = pc.IdGrupo1
'    mIDPcGrupo2 = pc.IdGrupo2
'    mIDPcGrupo3 = pc.IdGrupo3
'    Set pc = Nothing
'    'Selecciona  el descuento  predeterminado  si  esta configurado
'    If mobjGNComp.GNTrans.IVDescXPCGrupo = True Then
'        SelecDescxPCGrupo
'    End If
'    VisualizaTotal
End Sub

Private Sub mobjGNComp_CotizacionCambiado()
    mobjGNComp_MonedaCambiado
End Sub

Private Sub mobjGNComp_MonedaCambiado()
    On Error GoTo ErrTrap
    
    ConfigColsFormato
    VisualizaTotal
'    VisualizaDesdeObjeto
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub



'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_Resize()
    'Ajusta el tamaño del grid
    grd.Height = UserControl.ScaleHeight
End Sub

Private Sub UserControl_Terminate()
    Set mobjGNComp = Nothing
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub

Public Sub VisualizaDesdeObjeto()
    Dim ivkr As IVKardexRecargo, ivr As IvRecargo, i As Long, s As String
    Dim r As Long, Desc As String, rs As Recordset
    With grd
        .Redraw = False
        Limpiar
        ConfigCols
        For i = 1 To mobjGNComp.CountIVKardexRecargo
            Set ivkr = mobjGNComp.IVKardexRecargo(i)
            Set ivr = mobjGNComp.Empresa.RecuperaIVRecargo(ivkr.IdRecargo)
            Desc = ""
            If Not ivr Is Nothing Then Desc = ivr.Descripcion
            
            s = .Rows & vbTab & ivkr.codRecargo & vbTab & _
                    "" & vbTab & NullSiZero(ivkr.porcentaje * 100) & vbTab & _
                    NullSiZero(Abs(ivkr.valor)) & vbTab & 0 & vbTab & Desc & vbTab & "" _
                    & vbTab & "" & vbTab & "" & vbTab & ""
            .AddItem s
            r = .Rows - 1
            
            .TextMatrix(r, COL_SIGNO) = IIf(ivkr.auxSigno < 0, "-", "+")
            .Cell(flexcpChecked, r, COL_MODIFICABLE) = IIf(ivkr.BandModificable, flexChecked, flexUnchecked)
            .TextMatrix(r, COL_ORIGEN) = ivkr.BandOrigen
            .Cell(flexcpChecked, r, COL_PRORRAT) = IIf(ivkr.BandProrrateado, flexChecked, flexUnchecked)
            .Cell(flexcpChecked, r, COL_AFECTAIVA) = IIf(ivkr.AfectaIvaItem, flexChecked, flexUnchecked)
            'Set rs = grd.DataSource
            .Cell(flexcpChecked, r, COL_SELECCIONABLE) = IIf(ivkr.BandSeleccionable, flexChecked, flexUnchecked)
            .RowData(.Rows - 1) = ivkr
        Next i
        '***Diego 04/11/2003
        PoneNumFila
        mobjGNComp_ClienteCambiado
        'VisualizaTotal
        .Redraw = True
    End With
End Sub

'Subrutina para  seleccion automatica  de descuento
Public Sub SelecDescxPCGrupo()
    grd.Redraw = False
    'Estableceer  prioridad  para seleccion
    If mIDPcGrupo1 > 0 Then
        If SeleccionaDescuento(mIDPcGrupo1, 1) Then GoTo salida
    End If
    
    If mIDPcGrupo2 > 0 Then
        If SeleccionaDescuento(mIDPcGrupo2, 2) Then GoTo salida
    End If
    
    If mIDPcGrupo3 > 0 Then
        If SeleccionaDescuento(mIDPcGrupo3, 3) Then GoTo salida
    End If
    'si  no  se ha seleccionad ningun  cliente  borra  el descuento
    'SeleccionaDescuento 0, 0
    Dim j As Long
    For j = grd.FixedRows To grd.Rows - 1
        If grd.IsSubtotal(j) = False And grd.TextMatrix(j, COL_SELECCIONABLE) <> "0" Then
            grd.TextMatrix(j, COL_PORCENT) = Format(0, "##0.00")
            grd.TextMatrix(j, COL_VALOR) = Format(0, "##0.00")
        End If
    Next j
salida:
    grd.Redraw = True
End Sub

Private Function SeleccionaDescuento(idgrupo As Long, numGrupo As Integer) As Boolean
    Dim j As Long, i As Long, obj As IvRecargo
    SeleccionaDescuento = False
    For j = grd.FixedRows To grd.Rows - 1
        If grd.IsSubtotal(j) = False Then
        
            Set obj = gobjMain.EmpresaActual.RecuperaIVRecargo(grd.TextMatrix(j, COL_CODRECARGO))
            For i = 1 To obj.NumRecargoDetalle
                If obj.IVRecargoDetalle(i).NumPCGrupo = numGrupo And _
                                       obj.IVRecargoDetalle(i).IDPCGrupo = idgrupo Then
                    grd.TextMatrix(j, COL_PORCENT) = Format(obj.IVRecargoDetalle(i).valor, "##0.00")
                    'grd.Redraw = True
                    Set obj = Nothing
                    SeleccionaDescuento = True
                    Exit Function
                End If
            Next i
'            If mIDPcGrupo1 = 0 And mIDPcGrupo2 = 0 And mIDPcGrupo3 = 0 Then
'
'            End If
        End If
    Next j
    Set obj = Nothing
End Function


'***Angel. 11/nov/2003.
Public Sub ActualizarRecarSeleccionable()
    Dim Total As Long
    'Actualiza IvRecargo cuando  cambia  el cliente
    Dim pc As PCProvCli
    '***verificar cuando son modo ver y modifcar
    Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(mobjGNComp.IdClienteRef)
    '***Angel. 11/nov/2003. Modificado
    If (pc Is Nothing) Then Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(mobjGNComp.IdProveedorRef)   'Si no encuentra cliente busca por proveedor
    If (pc Is Nothing) Then
        mIDPcGrupo1 = 0
        mIDPcGrupo2 = 0
        mIDPcGrupo3 = 0
        'si no es seleccionable
        grd.ComboList = ""
    Else
        mIDPcGrupo1 = pc.IdGrupo1
        mIDPcGrupo2 = pc.IdGrupo2
        mIDPcGrupo3 = pc.IdGrupo3
        
    End If
    Set pc = Nothing
    'Selecciona  el descuento  predeterminado  si  esta configurado
    If (mobjGNComp.GNTrans.IVDescXPCGrupo = True) And (mobjGNComp.EsNuevo = True) Then
        SelecDescxPCGrupo
    End If
    Total = VisualizaTotal
End Sub

Private Sub grd_AfterUserResize(ByVal Row As Long, ByVal col As Long)
    With grd
        SaveSetting APPNAME, SECTION, "config_col_Rec_" & mobjGNComp.GNTrans.CodTrans & "_" & col, .ColWidth(col)
        ANCHO_COLS(col) = .ColWidth(col)
    End With
End Sub

Private Sub GetColsWidth()
    Dim i As Integer
    With grd
            For i = 0 To .Cols - 1
                ANCHO_COLS(i) = GetSetting(APPNAME, SECTION, "config_col_Rec_" & mobjGNComp.GNTrans.CodTrans & "_" & i, 1200)
            Next i
    End With
End Sub

