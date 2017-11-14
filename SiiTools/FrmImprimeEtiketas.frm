VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "Vsprint7.ocx"
Begin VB.Form FrmImprimeEtiketas 
   Caption         =   "Impresión Etiquetas"
   ClientHeight    =   7920
   ClientLeft      =   2160
   ClientTop       =   1770
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   8160
   WindowState     =   2  'Maximized
   Begin VSPrinter7LibCtl.VSPrinter vp 
      Align           =   2  'Align Bottom
      Height          =   7860
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   8160
      _cx             =   14393
      _cy             =   13864
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _ConvInfo       =   1
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   100
      ZoomMode        =   0
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
   End
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   1395
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   6375
      _cx             =   11245
      _cy             =   2461
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
      Rows            =   1
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   100
      ColWidthMax     =   4000
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   0
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   0
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
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   3
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "FrmImprimeEtiketas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const COL_E_NUMING = 1
Private Const COL_E_ORDEN = 2
Private Const COL_E_TIKET = 3
Private Const COL_E_MARCA = 4
Private Const COL_E_TAMANIO = 5
Private Const COL_E_SERIE = 6
Private Const COL_E_DISENIO = 7
Private Const COL_E_TRABAJO = 8
Private Const COL_E_TRANS = 9
Private Const COL_E_FECHA = 10
Private Const COL_E_CODCLI = 11
Private Const COL_E_NOMCLI = 12
Private Const COL_E_VENDE = 13
Private Const COL_E_GAR = 14
Private Const COL_E_RESULTADO = 15
Private FPAGO As Date

Option Explicit


Private Sub Form_Resize()
    'On Error Resume Next
    vp.Height = ScaleHeight
End Sub

Private Sub btnRender_Click()
    Dim i As Integer
    ' create a long string to add to some table cells
    Dim sLong$
    sLong = "This is the very long string that will cause some rows to break across pages. "
    sLong = sLong & vbCrLf & sLong & vbCrLf & sLong
    
    ' create document
    With vp
        .StartDoc
        .PaperSize = pprA4
        ' show intro
        
        
        .MarginLeft = 300
        .MarginRight = 300
        .MarginTop = 300
        .MarginBottom = 300
        .FontSize = 8
        ' set page and table borders
        .TableBorder = tbBox
        .TableBorder = 4
        .TablePenLR = 10
        .TablePenTB = 10
        
        For i = 1 To 50
        ' build table with 4 rows
        .StartTable
        .AddTable "900|600|600|600|600|900|900|600|600|600|700|880|10", "", "1234", RGB(200, 200, 250)
        .TableCell(tcRows) = 4
        
        ' center align all cells
        .TableCell(tcAlign) = taCenterMiddle
                
        ' add text to all cells
        Dim Row%, col%
'        For row = 1 To 10
'            For col = 1 To 3
'                If (row + col) Mod 7 <> 0 Then
'                    .TableCell(tcText, row, col) = " Row " & row & " Col " & col & " "
'                Else
'                    ' make a few cells have longer text, bold with a background
'                    .TableCell(tcText, row, col) = sLong
'                    .TableCell(tcBackColor, row, col) = RGB(100, 250, 100)
'                    .TableCell(tcFontBold, row, col) = True
'                End If
'            Next
'        Next
        
        .TableCell(tcText, 1, 1) = "Orden"
        .TableCell(tcText, 1, 3) = " INCAR 256987"
        .TableCell(tcText, 1, 6) = "Tiket No."
        .TableCell(tcText, 2, 1) = "256987"

        .TableCell(tcText, 2, 3) = "36987"
        .TableCell(tcText, 2, 5) = Str(1000 + i)
        .TableCell(tcText, 2, 4) = " AQ"
        
        .TableCell(tcText, 3, 1) = " 10.20.256"
        .TableCell(tcText, 3, 2) = " IDY"
        .TableCell(tcText, 3, 3) = " NS"
'        .TableCell(tcText, 3, 3) = " No.SERIE"
        .TableCell(tcText, 3, 5) = " 0102318144"
        .TableCell(tcText, 4, 1) = " Rec."
        .TableCell(tcText, 4, 3) = " Dis."
        .TableCell(tcText, 4, 6) = " 01/10/2009"
        
        
        .TableCell(tcText, 1, 7) = "Orden"
        .TableCell(tcText, 1, 9) = " INCAR"
        .TableCell(tcText, 1, 11) = "Tiket No."
        .TableCell(tcText, 2, 7) = "256987"

        .TableCell(tcText, 2, 9) = "36987"
        .TableCell(tcText, 2, 11) = " 111111"
        .TableCell(tcText, 2, 10) = " AQ"
        
        .TableCell(tcText, 3, 7) = " 10.20.256"
        .TableCell(tcText, 3, 8) = " IDY"
        .TableCell(tcText, 3, 9) = " NS"
'        .TableCell(tcText, 3, 3) = " No.SERIE"
        .TableCell(tcText, 3, 11) = " 0102318144"
        .TableCell(tcText, 4, 7) = " Rec."
        .TableCell(tcText, 4, 9) = " Dis."
        .TableCell(tcText, 4, 12) = " 01/10/2009"
        
        
        
        ' keep rows together
        .TableCell(tcRowKeepTogether) = True
        
                        
        ' apply colspan
'        If chkColSpan.Value Then
            .TableCell(tcColSpan, 1, 1) = 2
            .TableCell(tcColSpan, 1, 3) = 3
            .TableCell(tcColSpan, 2, 1) = 2
'            .TableCell(tcColSpan, 1, 5) = 2
            .TableCell(tcColSpan, 2, 5) = 2
            .TableCell(tcColSpan, 3, 3) = 2
            .TableCell(tcColSpan, 3, 5) = 2
            .TableCell(tcColSpan, 4, 1) = 2
            .TableCell(tcColSpan, 4, 3) = 3
            
            
            .TableCell(tcColSpan, 1, 7) = 2
            .TableCell(tcColSpan, 1, 9) = 2
            .TableCell(tcColSpan, 2, 7) = 2
            .TableCell(tcColSpan, 1, 11) = 2
            .TableCell(tcColSpan, 2, 11) = 2
            .TableCell(tcColSpan, 3, 9) = 2
            .TableCell(tcColSpan, 3, 11) = 2
            .TableCell(tcColSpan, 4, 7) = 2
            .TableCell(tcColSpan, 4, 9) = 3
            
            
            .TableCell(tcAlign, 1, 3) = 0
            .TableCell(tcAlign, 1, 9) = 0
            .TableCell(tcAlign, 3, 1, 3, 12) = 0
            .TableCell(tcAlign, 4, 1, 4, 5) = 0
            .TableCell(tcAlign, 4, 7, 4, 11) = 0
            

            .TableCell(tcFontSize, 1, 1, 1, 12) = 8
            .TableCell(tcFont, 1, 2) = "Arial"
            
            .TableCell(tcFontSize, 2, 1) = 18
            .TableCell(tcFontBold, 2, 1) = True
            .TableCell(tcFontSize, 2, 7) = 18
            .TableCell(tcFontBold, 2, 7) = True
            
            .TableCell(tcFontSize, 1, 3) = 8
            .TableCell(tcFontSize, 1, 9) = 8
            
            .TableCell(tcFontSize, 2, 3) = 8
            .TableCell(tcFontSize, 2, 9) = 8
            
            .TableCell(tcFontSize, 2, 5) = 19
            .TableCell(tcFontBold, 2, 5) = True
            .TableCell(tcFontSize, 4, 6) = 7
            .TableCell(tcColBorderRight, 1, 6, 4, 6) = 10
            .TableCell(tcFontSize, 2, 11) = 19
            .TableCell(tcFontBold, 2, 11) = True
            .TableCell(tcFontSize, 4, 12) = 7
                
        .EndTable
        Select Case i
            Case 15, 30, 45, 60
            .NewPage
        End Select
        Next i
        .EndDoc
    End With
End Sub

Private Sub vp_AfterTableCell(ByVal Row As Long, ByVal col As Long, ByVal Left As Double, ByVal Top As Double, ByVal Right As Double, ByVal Bottom As Double, Text As String, KeepFiring As Boolean)

    ' draw a cross over the cell
    
End Sub

Private Sub CargarDatos()
    Dim i As Integer, s As String, numEtik As Integer
    ' create a long string to add to some table cells
        ' create document
    With vp
    
        .PaperSize = pprA4
        .Refresh
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_MarIzq")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_MarIzq")
            .MarginLeft = s
        Else
            .MarginLeft = 300
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_MarSup")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_MarSup")
            .MarginTop = s
        Else
            .MarginTop = 300
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_NumEtiq")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_NumEtiq")
            numEtik = CInt(s)
        Else
            numEtik = 15
        End If

        
       
        
        ' show intro
        
        .MarginBottom = 300
        .FontSize = 8
        ' set page and table borders

        .TableBorder = 4
        .TablePenLR = 10
        .TablePenTB = 10
        '.TableBorder = tbAll
        .StartDoc
        For i = 1 To grd.Rows - 1
        ' build table with 4 rows
        
        .StartTable
'        .TableBorder = tbAll
        .AddTable "1000|600|400|600|600|1100|900|600|400|600|600|1100|10", "", "1234", RGB(200, 200, 250)
        .TableCell(tcRows) = 4
        
        ' center align all cells
        .TableCell(tcAlign) = taCenterMiddle
        .TablePenTB = 10
        If i > 1 Then
            If grd.TextMatrix(i, COL_E_TRANS) <> grd.TextMatrix(i - 1, COL_E_TRANS) Then
'                .Paragraph = " -----------------------------------------------------------------------------------------------------------"
                .TablePenTB = 30
            End If
        Else
            .TablePenTB = 10
        End If
                        
        .TableCell(tcText, 1, 1) = "Orden"
        .TableCell(tcText, 1, 3) = grd.TextMatrix(i, COL_E_CODCLI)
        .TableCell(tcText, 1, 6) = "Ticket No."
        .TableCell(tcText, 2, 1) = grd.TextMatrix(i, COL_E_NUMING)
        .TableCell(tcText, 2, 3) = grd.TextMatrix(i, COL_E_VENDE)
        .TableCell(tcText, 2, 5) = grd.TextMatrix(i, COL_E_TIKET)
        .TableCell(tcText, 2, 4) = grd.TextMatrix(i, COL_E_ORDEN)
        .TableCell(tcText, 3, 1) = grd.TextMatrix(i, COL_E_MARCA) & "/" & grd.TextMatrix(i, COL_E_TAMANIO)
        .TableCell(tcText, 3, 3) = " NS:" & Mid$(grd.TextMatrix(i, COL_E_SERIE), 1, 6)
        .TableCell(tcText, 3, 5) = Mid$(grd.TextMatrix(i, COL_E_NOMCLI), 1, 11)
        .TableCell(tcText, 4, 1) = " Rec." & grd.TextMatrix(i, COL_E_GAR)
        .TableCell(tcText, 4, 3) = " Dis." & grd.TextMatrix(i, COL_E_TRABAJO)
        .TableCell(tcText, 4, 6) = grd.TextMatrix(i, COL_E_FECHA)
        
        
        .TableCell(tcText, 1, 7) = .TableCell(tcText, 1, 1)
        .TableCell(tcText, 1, 9) = .TableCell(tcText, 1, 3)
        .TableCell(tcText, 1, 12) = .TableCell(tcText, 1, 6)
        .TableCell(tcText, 2, 7) = .TableCell(tcText, 2, 1)
        .TableCell(tcText, 2, 9) = .TableCell(tcText, 2, 3)
        .TableCell(tcText, 2, 10) = .TableCell(tcText, 2, 4)
        .TableCell(tcText, 2, 11) = .TableCell(tcText, 2, 5)
        .TableCell(tcText, 3, 7) = .TableCell(tcText, 3, 1)
        .TableCell(tcText, 3, 8) = .TableCell(tcText, 3, 2)
        .TableCell(tcText, 3, 9) = .TableCell(tcText, 3, 3)
        .TableCell(tcText, 3, 11) = .TableCell(tcText, 3, 5)
        .TableCell(tcText, 4, 7) = .TableCell(tcText, 4, 1)
        .TableCell(tcText, 4, 9) = .TableCell(tcText, 4, 3)
        .TableCell(tcText, 4, 12) = .TableCell(tcText, 4, 6)
        
        .TableCell(tcRowKeepTogether) = True

        .TableCell(tcColSpan, 1, 1) = 2
        .TableCell(tcColSpan, 1, 3) = 3
        .TableCell(tcColSpan, 2, 1) = 2
        .TableCell(tcColSpan, 2, 5) = 2
        .TableCell(tcColSpan, 3, 3) = 2
        
        .TableCell(tcColSpan, 3, 1, 3, 3) = 2
        .TableCell(tcColSpan, 3, 7, 3, 9) = 2
        
        
        .TableCell(tcColSpan, 3, 5) = 2
        .TableCell(tcColSpan, 4, 1) = 2
        .TableCell(tcColSpan, 4, 3) = 3
        .TableCell(tcColSpan, 1, 7) = 2
        .TableCell(tcColSpan, 1, 9) = 3
        .TableCell(tcColSpan, 2, 7) = 2
        .TableCell(tcColSpan, 1, 11) = 2
        .TableCell(tcColSpan, 2, 11) = 2
        .TableCell(tcColSpan, 3, 9) = 2
        .TableCell(tcColSpan, 3, 11) = 2
        .TableCell(tcColSpan, 4, 7) = 2
        .TableCell(tcColSpan, 4, 9) = 3
            
            
        .TableCell(tcAlign, 1, 3) = 0
        .TableCell(tcAlign, 1, 9) = 0
        .TableCell(tcAlign, 3, 1, 3, 12) = 0
        .TableCell(tcAlign, 4, 1, 4, 5) = 0
        .TableCell(tcAlign, 4, 7, 4, 11) = 0
        

        .TableCell(tcFontSize, 1, 1, 1, 12) = 8
        
        .TableCell(tcFont, 1, 2) = "Arial"
        
        
        
        .TableCell(tcFontSize, 2, 1) = 18
        .TableCell(tcFontBold, 2, 1) = True
        
        .TableCell(tcFontSize, 2, 7) = 18
        .TableCell(tcFontBold, 2, 7) = True
        
        .TableCell(tcFontSize, 2, 4) = 16
        .TableCell(tcFontBold, 2, 4) = True
        
        
        
        .TableCell(tcFontSize, 2, 4) = 11
        .TableCell(tcFontBold, 2, 4) = True
        .TableCell(tcFontSize, 2, 10) = 11
        .TableCell(tcFontBold, 2, 10) = True
        
        .TableCell(tcFontSize, 3, 1) = 8
        .TableCell(tcFontBold, 3, 1) = True
        
        
        .TableCell(tcFontSize, 3, 7) = 8
        .TableCell(tcFontBold, 3, 7) = True
        
        
        .TableCell(tcFontSize, 2, 7) = 18
        .TableCell(tcFontBold, 2, 7) = True
        
        .TableCell(tcFontSize, 3, 5) = 11
        .TableCell(tcFontBold, 3, 5) = True
        
        .TableCell(tcFontSize, 3, 11) = 11
        .TableCell(tcFontBold, 3, 11) = True
        
        
        .TableCell(tcFontSize, 1, 3) = 8
        .TableCell(tcFontSize, 1, 9) = 8
        
        .TableCell(tcFontSize, 2, 3) = 8
        .TableCell(tcFontSize, 2, 9) = 8
        
        .TableCell(tcFontSize, 2, 5) = 19
        .TableCell(tcFontBold, 2, 5) = True
        
        .TableCell(tcFontSize, 4, 3) = 9
        .TableCell(tcFontBold, 4, 3) = True
        
        .TableCell(tcFontSize, 4, 9) = 9
        .TableCell(tcFontBold, 4, 9) = True
        
        
        .TableCell(tcFontSize, 4, 6) = 8
        .TableCell(tcFontBold, 4, 6) = True
        
        .TableCell(tcColBorderRight, 1, 6, 4, 6) = 10
        
         .TableCell(tcFontSize, 2, 11) = 19
         .TableCell(tcFontBold, 2, 11) = True
         
         .TableCell(tcFontBold, 4, 12) = True
         .TableCell(tcFontSize, 4, 12) = 8
        
        .EndTable
        
        If numEtik > 0 Then
            If (i Mod numEtik) = 0 Then .NewPage
        End If
        
        Next i
        .EndDoc
    End With
End Sub


Public Sub Inicio(obj As Object)
    Dim v As Variant
    If Not obj.EOF Then
            v = MiGetRows(obj)
            
            grd.Redraw = flexRDNone
            grd.LoadArray v
            grd.Redraw = flexRDDirect
        'CargarDatos
        CargarDatosNew
    End If
    Me.Show

End Sub

Public Sub InicioNew(v As Variant)
'    Dim v As Variant
'    If Not obj.EOF Then
'            v = MiGetRows(obj)
            
            grd.Redraw = flexRDNone
            grd.LoadArray v
            grd.Redraw = flexRDDirect
'        CargarDatos
        CargarDatosNew
'    End If
    Me.Show

End Sub

Private Sub CargarDatosNew()
    Dim i As Integer, s As String, numEtik As Integer
    ' create a long string to add to some table cells
        ' create document
    With vp
    
        .PaperSize = pprA4
        .Refresh
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_MarIzq")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_MarIzq")
            .MarginLeft = s
        Else
            .MarginLeft = 300
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_MarSup")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_MarSup")
            .MarginTop = s
        Else
            .MarginTop = 300
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_NumEtiq")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_NumEtiq")
            numEtik = CInt(s)
        Else
            numEtik = 15
        End If

        
       
        
        ' show intro
        
        .MarginBottom = 300
        .FontSize = 8
        ' set page and table borders

        .TableBorder = 4
        .TablePenLR = 10
        .TablePenTB = 10
        '.TableBorder = tbAll
        .StartDoc
        For i = 1 To grd.Rows - 1
        ' build table with 4 rows
        
        .StartTable
        .FontName = "Arial"
        .FontSize = 8
        '.TableBorder = tbAll
        .AddTable "605|1510|580|800|980|50|605|1510|580|800|980|50|10", "", "", RGB(200, 200, 250)
        .TableCell(tcRows) = 4
        
        ' center align all cells
        .TableCell(tcAlign) = taCenterMiddle
        .TablePenTB = 10
        If i > 1 Then
            If grd.TextMatrix(i, COL_E_TRANS) <> grd.TextMatrix(i - 1, COL_E_TRANS) Then
'                .Paragraph = " -----------------------------------------------------------------------------------------------------------"
                .TablePenTB = 30
            End If
        Else
            .TablePenTB = 10
        End If
                        
        .TableCell(tcText, 1, 1) = "Orden"
        
        .TableCell(tcText, 1, 2) = grd.TextMatrix(i, COL_E_NUMING)
        .TableCell(tcFontSize, 1, 2) = 20
        .TableCell(tcFontBold, 1, 2) = True
        .TableCell(tcAlign, 1, 2) = 0
        .TableCell(tcText, 1, 3) = "Ticket"
        .TableCell(tcText, 1, 4) = grd.TextMatrix(i, COL_E_TIKET)
        .TableCell(tcFontSize, 1, 4) = 20
'        .TableCell(tcFontBold, 1, 4) = True
        .TableCell(tcColSpan, 1, 4) = 2
        .TableCell(tcAlign, 1, 4) = 0
        
        .TableCell(tcText, 2, 1) = Mid$(grd.TextMatrix(i, COL_E_NOMCLI), 1, 15)
        .TableCell(tcFontSize, 2, 1) = 14
        .TableCell(tcFontBold, 2, 1) = True
        .TableCell(tcColSpan, 2, 1) = 3
        .TableCell(tcAlign, 2, 1) = 0
        
        
        .TableCell(tcText, 2, 4) = grd.TextMatrix(i, COL_E_VENDE)
        '.TableCell(tcText, 2, 4) = grd.TextMatrix(i, COL_E_CODCLI)
        
        
        .TableCell(tcText, 2, 5) = grd.TextMatrix(i, COL_E_ORDEN)
        .TableCell(tcFontSize, 2, 5) = 14
        .TableCell(tcFontBold, 2, 5) = True
        
        
        .TableCell(tcText, 3, 1) = "(" & grd.TextMatrix(i, COL_E_MARCA) & ") " & grd.TextMatrix(i, COL_E_TAMANIO)
        .TableCell(tcFontSize, 3, 1) = 10
        .TableCell(tcColSpan, 3, 1) = 3
        .TableCell(tcAlign, 3, 1) = 0
        
        .TableCell(tcText, 3, 4) = " NS:"
        .TableCell(tcAlign, 3, 4) = 2
        .TableCell(tcText, 3, 5) = Mid$(grd.TextMatrix(i, COL_E_SERIE), 1, 6)
        .TableCell(tcFontSize, 3, 5) = 10
        .TableCell(tcAlign, 3, 5) = 0
      
        
        .TableCell(tcText, 4, 1) = " Rec." & grd.TextMatrix(i, COL_E_GAR)
        .TableCell(tcColSpan, 4, 1) = 2
        .TableCell(tcAlign, 4, 1) = 0
        
        
        .TableCell(tcText, 4, 3) = " Dis."
        .TableCell(tcAlign, 4, 3) = 2

        .TableCell(tcText, 4, 4) = grd.TextMatrix(i, COL_E_TRABAJO)
        .TableCell(tcFontSize, 4, 4) = 13
        .TableCell(tcFontBold, 4, 4) = True
        .TableCell(tcAlign, 4, 4) = 0
        .TableCell(tcText, 4, 5) = grd.TextMatrix(i, COL_E_FECHA)
        
        '----------- SEGUNDA ETIQUETA
        
        
        .TableCell(tcText, 1, 7) = "Orden"
        
        .TableCell(tcText, 1, 8) = grd.TextMatrix(i, COL_E_NUMING)
        .TableCell(tcFontSize, 1, 8) = 20
        .TableCell(tcFontBold, 1, 8) = True
        .TableCell(tcAlign, 1, 8) = 0
        
        .TableCell(tcText, 1, 9) = "Ticket"
        .TableCell(tcText, 1, 10) = grd.TextMatrix(i, COL_E_TIKET)
        .TableCell(tcFontSize, 1, 10) = 20
        .TableCell(tcFontBold, 1, 10) = True
        .TableCell(tcColSpan, 1, 10) = 2
        .TableCell(tcAlign, 1, 10) = 0
        
        .TableCell(tcText, 2, 7) = Mid$(grd.TextMatrix(i, COL_E_NOMCLI), 1, 15)
        .TableCell(tcFontSize, 2, 7) = 14
        .TableCell(tcFontBold, 2, 7) = True
        .TableCell(tcColSpan, 2, 7) = 3
        .TableCell(tcAlign, 2, 7) = 0
        
        
        .TableCell(tcText, 2, 10) = grd.TextMatrix(i, COL_E_VENDE)
        '.TableCell(tcText, 2, 10) = grd.TextMatrix(i, COL_E_CODCLI)
        
        
        .TableCell(tcText, 2, 11) = grd.TextMatrix(i, COL_E_ORDEN)
        .TableCell(tcFontSize, 2, 11) = 14
        .TableCell(tcFontBold, 2, 11) = True
        
        
        .TableCell(tcText, 3, 7) = "(" & grd.TextMatrix(i, COL_E_MARCA) & ") " & grd.TextMatrix(i, COL_E_TAMANIO)
        .TableCell(tcFontSize, 3, 7) = 10
        .TableCell(tcColSpan, 3, 7) = 3
        .TableCell(tcAlign, 3, 7) = 0
        
        .TableCell(tcText, 3, 10) = " NS:"
        .TableCell(tcAlign, 3, 10) = 2
        .TableCell(tcText, 3, 11) = Mid$(grd.TextMatrix(i, COL_E_SERIE), 1, 6)
        .TableCell(tcFontSize, 3, 11) = 10
        .TableCell(tcAlign, 3, 11) = 0
      
        
        .TableCell(tcText, 4, 7) = " Rec." & grd.TextMatrix(i, COL_E_GAR)
        .TableCell(tcColSpan, 4, 7) = 2
        .TableCell(tcAlign, 4, 7) = 0
        
        .TableCell(tcText, 4, 9) = " Dis."
        .TableCell(tcAlign, 4, 9) = 2
        .TableCell(tcText, 4, 10) = grd.TextMatrix(i, COL_E_TRABAJO)
        .TableCell(tcFontSize, 4, 10) = 13
        .TableCell(tcAlign, 4, 10) = 0
        .TableCell(tcFontBold, 4, 10) = True
        
        
        .TableCell(tcText, 4, 11) = grd.TextMatrix(i, COL_E_FECHA)
        
        

        .TableCell(tcRowKeepTogether) = True

        .TableCell(tcColBorderRight, 1, 6, 4, 6) = 10


        .EndTable
        
        If numEtik > 0 Then
            If (i Mod numEtik) = 0 Then .NewPage
        End If
        
        Next i
        .EndDoc
    End With
End Sub

Public Sub InicioF101(ByVal g As VSFlexGrid, Tipo As Integer)
    Dim i As Integer, j As Integer, fila As Integer
    grd.Rows = g.Rows
    grd.Cols = g.Cols
    
    If Tipo = 0 Then
        For i = 0 To g.Rows - 1
            For j = 0 To g.Cols - 1
                grd.TextMatrix(i, j) = g.TextMatrix(i, j)
            Next j
        Next i
    Else
        fila = 1
        For i = 0 To g.Rows - 1
            If g.IsSubtotal(i) Then
                For j = 0 To g.Cols - 1
                    grd.TextMatrix(fila, j) = g.TextMatrix(i, j)
                Next j
                fila = fila + 1
            End If
        Next i
        grd.Rows = fila
    End If
    
    
    grd.Redraw = flexRDDirect
    CargarDatosF101
    Me.Show

End Sub


Private Sub CargarDatosF101()
    Dim i As Integer, s As String, numEtik As Integer
    Dim j As Integer, k As Integer
    ' create a long string to add to some table cells
        ' create document
    With vp
    
        .PaperSize = pprA4
        .Refresh
        
       
        
        ' show intro
        
        .MarginBottom = 300
        .FontSize = 8
        ' set page and table borders

        .TableBorder = 4
        .TablePenLR = 10
        .TablePenTB = 10
        '.TableBorder = tbAll
        .StartDoc
'        For i = 1 To grd.Rows - 1
        ' build table with 4 rows
        .FontSize = 14
        .Paragraph = "Plantilla Formulario 101"
        .Paragraph = ""
        .StartTable
        .FontName = "Arial"
        .FontSize = 9
        '.TableBorder = tbAll
        .AddTable "600|1300|1000|1500|4000|1500", "#|Tipo|Campo F101|Codigo Cuenta|Nombre Cuenta|Valor", " ", RGB(200, 200, 250)
        .TableCell(tcRows) = grd.Rows
        
        
        ' center align all cells
        '.TableCell(tcAlign)= = taCenterMiddle
        .TablePenTB = 8
        
        For k = 1 To grd.Rows - 1
            For j = 0 To grd.Cols - 1
                If j = grd.Cols - 1 Then
                    .TableCell(tcText, k, j + 1) = Format(grd.TextMatrix(k, j), "#,#0.00")
                Else
                    .TableCell(tcText, k, j + 1) = grd.TextMatrix(k, j)
                End If
                If Len(grd.TextMatrix(k, 0)) = 0 Then
                    .TableCell(tcBackColor, k, 0, k, 6) = RGB(200, 200, 250)
                End If
            Next j
        Next k
        
        .TableCell(tcAlign, 1, 1, grd.Rows - 1, 1) = 2
        .TableCell(tcAlign, 1, 3, grd.Rows - 1, 3) = 2
        .TableCell(tcAlign, 1, 6, grd.Rows - 1, 6) = 2
        
        
        
        .TableBorder = tbAll
        
        .EndTable
        
        
 '       Next i
        .EndDoc
    End With
End Sub



Public Sub InicioF104(ByVal g As VSFlexGrid)
    Dim i As Integer, j As Integer, fila As Integer
    grd.Rows = 92 'g.Rows
    grd.Cols = g.Cols
    grd.MergeCells = flexMergeSpill
        fila = 1
        For i = 0 To 91 'g.Rows - 1
            For j = 0 To g.Cols - 1
                grd.TextMatrix(i, j) = g.TextMatrix(i, j)
                grd.ColWidth(j) = g.ColWidth(j)
            Next j
        Next i
'        grd.Rows = fila
    
    
    
    grd.Redraw = flexRDDirect
    CargarDatosF104
    Me.Show

End Sub

Private Sub CargarDatosF104()
    Dim i As Integer, s As String, numEtik As Integer
    Dim j As Integer, k As Integer
    ' create a long string to add to some table cells
        ' create document
    With vp
    
        .PaperSize = pprA4
        .Refresh
        
       
        
        ' show intro
        
        .MarginBottom = 300
        .FontSize = 8
        ' set page and table borders
        .MarginLeft = 200
        .MarginTop = 200
        .MarginRight = 200
        .MarginBottom = 500
        .TableBorder = 4
        .TablePenLR = 10
        .TablePenTB = 10
        '.TableBorder = tbAll
        .StartDoc
'        For i = 1 To grd.Rows - 1
        ' build table with 4 rows
        .FontSize = 14
        .Paragraph = "Formulario 104"
        .FontSize = 8
        .Paragraph = " "
        .Paragraph = gobjMain.EmpresaActual.GNOpcion.NombreEmpresa & Space(40) & "Usuario:" & gobjMain.UsuarioActual.NombreUsuario & Space(40) & "Fecha:" & Date & "-" & Time
        'Paragraph = "........."
        'Paragraph = "........."
        .StartTable
        .FontName = "Arial"
        .FontSize = 6
        '.TableBorder = tbAll
        
'        .ColWidth(0) = 350
'        .ColWidth(1) = 800
'        .ColWidth(2) = 1500
'        .ColWidth(3) = 800
'        .ColWidth(4) = 1250
'        .ColWidth(5) = 1500
'        .ColWidth(6) = 1350
'        .ColWidth(7) = 470
'        .ColWidth(8) = 1400
'        .ColWidth(9) = 470
'        .ColWidth(10) = 1400
'        .ColWidth(11) = 450
'        .ColWidth(12) = 1400
'        .ColWidth(13) = 400
'        .ColWidth(14) = 1400
        
        
        .AddTable "300|650|1100|650|700|1430|1000|450|1000|400|1000|400|1000|400|1000", "#|c1|c2|c3|c4|c5|c6|c7|c8|c9|c10|c11|c12|c13|c14", " ", RGB(200, 200, 250)
        
        .TableCell(tcRows) = grd.Rows
        
        
        ' center align all cells
        '.TableCell(tcAlign)= = taCenterMiddle
        .TablePenTB = 8
        
        For k = 1 To 91 'grd.Rows - 1
            For j = 0 To grd.Cols - 1
                If j = grd.Cols - 1 Then
                    .TableCell(tcText, k, j + 1) = Format(grd.TextMatrix(k, j), "#,#0.00")
                Else
                    .TableCell(tcText, k, j + 1) = grd.TextMatrix(k, j)
                End If
                If Len(grd.TextMatrix(k, 0)) = 0 Then
                    .TableCell(tcBackColor, k, 0, k, 6) = RGB(200, 200, 250)
                End If
            Next j
        Next k
        
    
    .TableCell(tcText, 0, 1) = " "
    
    .TableCell(tcText, 56, 13) = Round(grd.ValueMatrix(56, 12), 2)
    .TableCell(tcText, 57, 13) = Round(grd.ValueMatrix(57, 12), 2)
    .TableCell(tcColSpan, 0, 1) = 15
    .TableCell(tcColSpan, 1, 2) = 2
    .TableCell(tcColSpan, 1, 4) = 12
    .TableCell(tcColSpan, 2, 2) = 14
    .TableCell(tcColSpan, 3, 2) = 14
    .TableCell(tcColSpan, 4, 2) = 14
    
    .TableCell(tcColSpan, 5, 2) = 14
    .TableCell(tcText, 5, 2) = " "
    
    .TableCell(tcColSpan, 6, 2) = 14
    .TableCell(tcColSpan, 7, 5) = 4
    .TableCell(tcColSpan, 7, 10) = 6
    .TableCell(tcColSpan, 8, 2) = 8
    .TableCell(tcColSpan, 8, 10) = 6
    .TableCell(tcColSpan, 9, 2) = 14
    .TableCell(tcColSpan, 10, 5) = 11
    .TableCell(tcColSpan, 11, 2) = 14
    .TableCell(tcColSpan, 12, 2) = 6
    .TableCell(tcColSpan, 12, 8) = 2
    .TableCell(tcColSpan, 12, 13) = 2
    .TableCell(tcColSpan, 13, 2) = 6


    
    For k = 14 To 26
        .TableCell(tcColSpan, k, 2) = 6
    Next k
    
    .TableCell(tcColSpan, 27, 2) = 14
    .TableCell(tcColSpan, 28, 2) = 14
    
    .TableCell(tcColSpan, 29, 2) = 2
    .TableCell(tcColSpan, 29, 4) = 2
    .TableCell(tcColSpan, 29, 8) = 2
    .TableCell(tcColSpan, 29, 10) = 2
    .TableCell(tcColSpan, 29, 12) = 2
    .TableCell(tcColSpan, 29, 14) = 2
    
    .TableCell(tcColSpan, 30, 2) = 2
    .TableCell(tcColSpan, 30, 4) = 2
    .TableCell(tcColSpan, 30, 8) = 2
    .TableCell(tcColSpan, 30, 10) = 2
    .TableCell(tcColSpan, 30, 12) = 2
    .TableCell(tcColSpan, 30, 14) = 2
    
    .TableCell(tcColSpan, 31, 2) = 2
    .TableCell(tcColSpan, 31, 4) = 2
    .TableCell(tcColSpan, 31, 8) = 2
    .TableCell(tcColSpan, 31, 10) = 2
    .TableCell(tcColSpan, 31, 12) = 2
    .TableCell(tcColSpan, 31, 14) = 2
    
    .TableCell(tcColSpan, 33, 2) = 14

    .TableCell(tcColSpan, 34, 2) = 6
    .TableCell(tcColSpan, 34, 8) = 2
    .TableCell(tcColSpan, 34, 13) = 2

    .TableCell(tcColSpan, 35, 2) = 6
    
    
    
    For k = 36 To 53
        .TableCell(tcColSpan, k, 2) = 6
    Next k
    
    .TableCell(tcColSpan, 51, 2) = 14
    .TableCell(tcColSpan, 52, 2) = 10
    .TableCell(tcColSpan, 53, 2) = 10
    .TableCell(tcColSpan, 54, 2) = 14
    .TableCell(tcColSpan, 55, 2) = 14
    
    For k = 56 To 99
        .TableCell(tcColSpan, k, 2) = 10
    Next k
    
    .TableCell(tcColSpan, 59, 2) = 14
    .TableCell(tcColSpan, 69, 2) = 14
    .TableCell(tcColSpan, 70, 2) = 14
    .TableCell(tcColSpan, 71, 2) = 14
    .TableCell(tcColSpan, 76, 2) = 14
    .TableCell(tcColSpan, 78, 2) = 14
    .TableCell(tcColSpan, 80, 2) = 14
    .TableCell(tcColSpan, 83, 2) = 14
    .TableCell(tcColSpan, 88, 2) = 14
    .TableCell(tcColSpan, 92, 2) = 14
    
'    For k = 1 To grd.Rows - 1
'        For j = 0 To grd.Cols - 1
'            .TableCell(tcColWidth, k, j, k, j) = grd.ColWidth(j)
'            '.TableCell(tcColWidth, 1, 2, 100, 2) = 3000
'        Next j
'    Next k
        
        
'        .TableCell(tcAlign, 1, 1, grd.Rows - 1, 1) = 2
'        .TableCell(tcAlign, 1, 3, grd.Rows - 1, 3) = 2
'        .TableCell(tcAlign, 1, 6, grd.Rows - 1, 6) = 2

        .TableCell(tcAlign, 1, 9, grd.Rows - 1, 9) = 2
        .TableCell(tcAlign, 1, 11, grd.Rows - 1, 11) = 2
        .TableCell(tcAlign, 1, 13, grd.Rows - 1, 13) = 2
        
        
        .TableCell(tcBackColor, 7, 2, 7, 2) = RGB(200, 200, 250)
        .TableCell(tcBackColor, 7, 4, 7, 4) = RGB(200, 200, 250)
        .TableCell(tcBackColor, 10, 2, 10, 2) = RGB(200, 200, 250)
        .TableCell(tcBackColor, 10, 4, 10, 4) = RGB(200, 200, 250)
        .TableCell(tcBackColor, 12, 8, 13, 14) = RGB(200, 200, 250)
        
        .TableCell(tcBackColor, 32, 2, 32, 2) = RGB(200, 200, 250)
        .TableCell(tcBackColor, 32, 4, 32, 4) = RGB(200, 200, 250)
        .TableCell(tcBackColor, 32, 6, 32, 6) = RGB(200, 200, 250)
        .TableCell(tcBackColor, 34, 8, 35, 14) = RGB(200, 200, 250)
        
        .TableCell(tcBackColor, 12, 8, grd.Rows - 1, 8) = RGB(200, 200, 250)
        .TableCell(tcBackColor, 12, 10, grd.Rows - 1, 10) = RGB(200, 200, 250)
        .TableCell(tcBackColor, 12, 12, grd.Rows - 1, 12) = RGB(200, 200, 250)
        .TableCell(tcBackColor, 12, 14, grd.Rows - 1, 14) = RGB(200, 200, 250)

        .TableCell(tcFontSize, 13, 1, 13, 14) = 5
        .TableCell(tcFontSize, 31, 1, 31, 14) = 5
        .TableCell(tcFontSize, 35, 1, 35, 14) = 5

        .TableCell(tcFontBold, 7, 2, 7, 2) = True
        .TableCell(tcFontBold, 7, 4, 7, 4) = True
        .TableCell(tcFontBold, 10, 2, 10, 2) = True
        .TableCell(tcFontBold, 10, 4, 10, 4) = True
        .TableCell(tcFontBold, 12, 8, 13, 14) = True
        
        .TableCell(tcFontBold, 7, 2, 7, 2) = True
        
        
        
        .TableCell(tcFontBold, 32, 2, 32, 2) = True
        .TableCell(tcFontBold, 32, 4, 32, 4) = True
        .TableCell(tcFontBold, 32, 6, 32, 6) = True
        .TableCell(tcFontBold, 34, 8, 35, 14) = True
        
        .TableCell(tcFontBold, 12, 8, grd.Rows - 1, 8) = True
        .TableCell(tcFontBold, 12, 10, grd.Rows - 1, 10) = True
        .TableCell(tcFontBold, 12, 12, grd.Rows - 1, 12) = True
        .TableCell(tcFontBold, 12, 14, grd.Rows - 1, 14) = True



        .TableBorder = tbAll
        .EndTable
        
        
 '       Next i
        .EndDoc
    End With
End Sub

Public Sub InicioF103(ByVal g As VSFlexGrid)
    Dim i As Integer, j As Integer, fila As Integer
    grd.Rows = 62 'g.Rows
    grd.Cols = g.Cols
    grd.MergeCells = flexMergeSpill
        fila = 1
        For i = 0 To 61 'g.Rows - 1
            For j = 0 To g.Cols - 1
                grd.TextMatrix(i, j) = g.TextMatrix(i, j)
                grd.ColWidth(j) = g.ColWidth(j)
            Next j
        Next i
'        grd.Rows = fila
    
    
    
    grd.Redraw = flexRDDirect
    CargarDatosF103
    Me.Show

End Sub

Private Sub CargarDatosF103()
    Dim i As Integer, s As String, numEtik As Integer
    Dim j As Integer, k As Integer
    ' create a long string to add to some table cells
        ' create document
    With vp
    
        .PaperSize = pprA4
        .Refresh
        
       
        
        ' show intro
        
        .MarginBottom = 300
        .FontSize = 8
        ' set page and table borders
        .MarginLeft = 200
        .MarginTop = 200
        .MarginRight = 200
        .MarginBottom = 500
        .TableBorder = 4
        .TablePenLR = 10
        .TablePenTB = 10
        '.TableBorder = tbAll
        .StartDoc
'        For i = 1 To grd.Rows - 1
        ' build table with 4 rows
        .FontSize = 14
        .Paragraph = "Formulario 103"
        .FontSize = 8
        .Paragraph = " "
        .Paragraph = gobjMain.EmpresaActual.GNOpcion.NombreEmpresa & Space(40) & "Usuario:" & gobjMain.UsuarioActual.NombreUsuario & Space(40) & "Fecha:" & Date & "-" & Time
        'Paragraph = "........."
        'Paragraph = "........."
        .StartTable
        .FontName = "Arial"
        .FontSize = 6
        '.TableBorder = tbAll
        
'        .ColWidth(0) = 350
'        .ColWidth(1) = 800
'        .ColWidth(2) = 1500
'        .ColWidth(3) = 800
'        .ColWidth(4) = 1250
'        .ColWidth(5) = 1500
'        .ColWidth(6) = 1350
'        .ColWidth(7) = 470
'        .ColWidth(8) = 1400
'        .ColWidth(9) = 470
'        .ColWidth(10) = 1400
'        .ColWidth(11) = 450
'        .ColWidth(12) = 1400
'        .ColWidth(13) = 400
'        .ColWidth(14) = 1400
        
        
        .AddTable "300|1450|1300|1300|1000|1430|600|800|400|400|1000|400|1000", "#|c1|c2|c3|c4|c5|c6|c7|c8|c9|c10|c11|c12", " ", RGB(200, 200, 250)
        
        .TableCell(tcRows) = grd.Rows
        
        
        ' center align all cells
        '.TableCell(tcAlign)= = taCenterMiddle
        .TablePenTB = 8
        
        For k = 1 To 61 'grd.Rows - 1
            For j = 0 To grd.Cols - 1
                If j = grd.Cols - 1 Then
                    .TableCell(tcText, k, j + 1) = Format(grd.TextMatrix(k, j), "#,#0.00")
                Else
                    .TableCell(tcText, k, j + 1) = grd.TextMatrix(k, j)
                End If
                If Len(grd.TextMatrix(k, 0)) = 0 Then
                    .TableCell(tcBackColor, k, 0, k, 6) = RGB(200, 200, 250)
                End If
            Next j
        Next k
        
    
    .TableCell(tcText, 0, 1) = " "

    
    .TableCell(tcColSpan, 0, 1) = 13
    .TableCell(tcColSpan, 1, 2) = 2
    .TableCell(tcColSpan, 1, 4) = 10
    .TableCell(tcColSpan, 2, 2) = 12
    .TableCell(tcColSpan, 3, 2) = 12
    .TableCell(tcColSpan, 4, 2) = 12

    .TableCell(tcColSpan, 5, 2) = 12
    .TableCell(tcText, 5, 2) = " "

    .TableCell(tcColSpan, 6, 2) = 12
    .TableCell(tcColSpan, 7, 5) = 3
    .TableCell(tcColSpan, 7, 8) = 6
    .TableCell(tcColSpan, 8, 2) = 6
    .TableCell(tcColSpan, 8, 8) = 6
    .TableCell(tcColSpan, 9, 2) = 12
    .TableCell(tcColSpan, 10, 5) = 11
    .TableCell(tcColSpan, 11, 2) = 12
    .TableCell(tcColSpan, 12, 2) = 2
    .TableCell(tcColSpan, 12, 4) = 10
    .TableCell(tcColSpan, 13, 2) = 3
    .TableCell(tcColSpan, 13, 5) = 9
    .TableCell(tcColSpan, 14, 2) = 8
    .TableCell(tcColSpan, 14, 10) = 2
    .TableCell(tcColSpan, 14, 12) = 2
    .TableCell(tcColSpan, 15, 2) = 8
    For k = 16 To 21
        .TableCell(tcColSpan, k, 3) = 7
    Next k
    .TableCell(tcColSpan, 22, 2) = 8
    .TableCell(tcColSpan, 23, 3) = 7
    .TableCell(tcColSpan, 24, 3) = 7
    For k = 25 To 28
        .TableCell(tcColSpan, k, 2) = 8
    Next k
    .TableCell(tcColSpan, 29, 3) = 7
    .TableCell(tcColSpan, 30, 3) = 7
    
    .TableCell(tcColSpan, 31, 4) = 3
    .TableCell(tcColSpan, 32, 4) = 3
    
    
    .TableCell(tcColSpan, 31, 8) = 2
    .TableCell(tcColSpan, 32, 8) = 2
    .TableCell(tcColSpan, 33, 2) = 8
    For k = 34 To 38
        .TableCell(tcColSpan, k, 3) = 7
    Next k
    .TableCell(tcColSpan, 39, 2) = 8
    .TableCell(tcColSpan, 40, 2) = 4
    .TableCell(tcColSpan, 40, 6) = 8
    .TableCell(tcColSpan, 41, 2) = 8
    For k = 42 To 45
        .TableCell(tcColSpan, k, 3) = 7
    Next k
    .TableCell(tcColSpan, 46, 2) = 8
    .TableCell(tcColSpan, 47, 2) = 8
    .TableCell(tcColSpan, 48, 1) = 13
    .TableCell(tcColSpan, 49, 2) = 10
    .TableCell(tcColSpan, 50, 1) = 13
    .TableCell(tcColSpan, 51, 2) = 10
    .TableCell(tcColSpan, 52, 2) = 12
    
    .TableCell(tcColSpan, 53, 4) = 2
    .TableCell(tcColSpan, 53, 7) = 4
    
    .TableCell(tcColSpan, 54, 3) = 3
    .TableCell(tcColSpan, 54, 6) = 6
    .TableCell(tcColSpan, 55, 1) = 13
    .TableCell(tcColSpan, 56, 2) = 12
    .TableCell(tcColSpan, 57, 2) = 10
    .TableCell(tcColSpan, 58, 2) = 10
    .TableCell(tcColSpan, 59, 2) = 10
    .TableCell(tcColSpan, 60, 2) = 10
  

        .TableCell(tcAlign, 1, 11, grd.Rows - 1, 11) = 2
        .TableCell(tcAlign, 1, 13, grd.Rows - 1, 13) = 2
        
        .TableCell(tcBackColor, 7, 2, 7, 2) = RGB(200, 200, 250)
        .TableCell(tcBackColor, 7, 4, 7, 4) = RGB(200, 200, 250)
        .TableCell(tcBackColor, 10, 2, 10, 2) = RGB(200, 200, 250)
        .TableCell(tcBackColor, 10, 4, 10, 4) = RGB(200, 200, 250)
        .TableCell(tcBackColor, 14, 10, 47, 10) = RGB(200, 200, 250)
        .TableCell(tcBackColor, 14, 12, 60, 12) = RGB(200, 200, 250)
        .TableCell(tcBackColor, 31, 7, 32, 7) = RGB(200, 200, 250)
        .TableCell(tcBackColor, 53, 3, 53, 3) = RGB(200, 200, 250)
        .TableCell(tcBackColor, 53, 6, 53, 6) = RGB(200, 200, 250)
        .TableCell(tcBackColor, 53, 11, 53, 11) = RGB(200, 200, 250)
        .TableCell(tcBackColor, 33, 11, 33, 13) = RGB(200, 200, 250)
        
        .TableCell(tcFontBold, 7, 2, 7, 2) = True
        .TableCell(tcFontBold, 7, 4, 7, 4) = True
        .TableCell(tcFontBold, 10, 2, 10, 2) = True
        .TableCell(tcFontBold, 10, 4, 10, 4) = True
        .TableCell(tcFontBold, 14, 10, 47, 10) = True
        .TableCell(tcFontBold, 14, 12, 60, 12) = True
        .TableCell(tcFontBold, 31, 7, 32, 7) = True
        .TableCell(tcFontBold, 53, 3, 53, 3) = True
        .TableCell(tcFontBold, 53, 6, 53, 6) = True
        .TableCell(tcFontBold, 53, 11, 53, 11) = True
        
        
''''        .TableCell(tcBackColor, 32, 4, 32, 4) = RGB(200, 200, 250)
''''        .TableCell(tcBackColor, 32, 6, 32, 6) = RGB(200, 200, 250)
''''        .TableCell(tcBackColor, 34, 8, 35, 14) = RGB(200, 200, 250)
''''
''''        .TableCell(tcBackColor, 12, 8, grd.Rows - 1, 8) = RGB(200, 200, 250)
''''        .TableCell(tcBackColor, 12, 10, grd.Rows - 1, 10) = RGB(200, 200, 250)
''''        .TableCell(tcBackColor, 12, 12, grd.Rows - 1, 12) = RGB(200, 200, 250)
''''        .TableCell(tcBackColor, 12, 14, grd.Rows - 1, 14) = RGB(200, 200, 250)
''''
''''        .TableCell(tcFontSize, 13, 1, 13, 14) = 5
''''        .TableCell(tcFontSize, 31, 1, 31, 14) = 5
''''        .TableCell(tcFontSize, 35, 1, 35, 14) = 5
''''
''''        .TableCell(tcFontBold, 7, 2, 7, 2) = True
''''        .TableCell(tcFontBold, 7, 4, 7, 4) = True
''''        .TableCell(tcFontBold, 10, 2, 10, 2) = True
''''        .TableCell(tcFontBold, 10, 4, 10, 4) = True
''''        .TableCell(tcFontBold, 12, 8, 13, 14) = True
''''
''''        .TableCell(tcFontBold, 7, 2, 7, 2) = True
''''
''''
''''
''''        .TableCell(tcFontBold, 32, 2, 32, 2) = True
''''        .TableCell(tcFontBold, 32, 4, 32, 4) = True
''''        .TableCell(tcFontBold, 32, 6, 32, 6) = True
''''        .TableCell(tcFontBold, 34, 8, 35, 14) = True
''''
''''        .TableCell(tcFontBold, 12, 8, grd.Rows - 1, 8) = True
''''        .TableCell(tcFontBold, 12, 10, grd.Rows - 1, 10) = True
''''        .TableCell(tcFontBold, 12, 12, grd.Rows - 1, 12) = True
''''        .TableCell(tcFontBold, 12, 14, grd.Rows - 1, 14) = True
''''


        .TableBorder = tbAll
        .EndTable
        
        
        .EndDoc
    End With
End Sub

Public Sub InicioProduccion(v As Variant)
           
    grd.Redraw = flexRDNone
    grd.LoadArray v
    grd.Redraw = flexRDDirect
    CargarDatosProduccion
    Me.Show

End Sub

Private Sub CargarDatosProduccion()
    Dim i As Integer, s As String, numEtik As Integer
    ' create a long string to add to some table cells
        ' create document
    With vp
    
        .PaperSize = pprA4
        .Refresh
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLoteProd_MarIzq")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLoteProd_MarIzq")
            .MarginLeft = s
        Else
            .MarginLeft = 300
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLoteProd_MarSup")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLoteProd_MarSup")
            .MarginTop = s
        Else
            .MarginTop = 300
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLoteProd_NumEtiq")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLoteProd_NumEtiq")
            numEtik = CInt(s)
        Else
            numEtik = 12
        End If

        
       
        
        ' show intro
        
        .MarginBottom = 300
        .FontSize = 8
        ' set page and table borders

        .TableBorder = 4
        .TablePenLR = 10
        .TablePenTB = 10
        '.TableBorder = tbAll
        .StartDoc
        .Paragraph = "."
        For i = 1 To grd.Rows - 1
        ' build table with 4 rows
        
        .StartTable
        .FontName = "Arial"
        .FontSize = 8
        '.TableBorder = tbAll
        .AddTable "605|1510|580|800|980|50|605|1510|580|800|980|50", "", "", RGB(200, 200, 250)
        .TableCell(tcRows) = 4
        
        ' center align all cells
        .TableCell(tcAlign) = taCenterMiddle
        .TablePenTB = 10
        If i > 1 Then
            If grd.TextMatrix(i, COL_E_TRANS) <> grd.TextMatrix(i - 1, COL_E_TRANS) Then
'                .Paragraph = " -----------------------------------------------------------------------------------------------------------"
                .TablePenTB = 30
            End If
        Else
            .TablePenTB = 10
        End If
                        
        .TableCell(tcText, 1, 1) = "Orden"
        
        .TableCell(tcText, 1, 2) = grd.TextMatrix(i, COL_E_NUMING)
        .TableCell(tcFontSize, 1, 2) = 20
        .TableCell(tcFontBold, 1, 2) = True
        .TableCell(tcAlign, 1, 2) = 0
        .TableCell(tcText, 1, 3) = "Ticket"
        .TableCell(tcText, 1, 4) = grd.TextMatrix(i, COL_E_TIKET)
        .TableCell(tcFontSize, 1, 4) = 20
        .TableCell(tcFontBold, 1, 4) = True
        .TableCell(tcColSpan, 1, 4) = 2
        .TableCell(tcAlign, 1, 4) = 0
        
        .TableCell(tcText, 2, 1) = Mid$(grd.TextMatrix(i, COL_E_NOMCLI), 1, 15)
        .TableCell(tcFontSize, 2, 1) = 14
        .TableCell(tcFontBold, 2, 1) = True
        .TableCell(tcColSpan, 2, 1) = 3
        .TableCell(tcAlign, 2, 1) = 0
        
        
        .TableCell(tcText, 2, 4) = grd.TextMatrix(i, COL_E_VENDE)
        '.TableCell(tcText, 2, 4) = grd.TextMatrix(i, COL_E_CODCLI)
        
        
        .TableCell(tcText, 2, 5) = grd.TextMatrix(i, COL_E_ORDEN)
        .TableCell(tcFontSize, 2, 5) = 14
        .TableCell(tcFontBold, 2, 5) = True
        
        
        .TableCell(tcText, 3, 1) = "(" & grd.TextMatrix(i, COL_E_MARCA) & ") " & grd.TextMatrix(i, COL_E_TAMANIO)
        .TableCell(tcFontSize, 3, 1) = 10
        .TableCell(tcColSpan, 3, 1) = 3
        .TableCell(tcAlign, 3, 1) = 0
        
        .TableCell(tcText, 3, 4) = " NS:"
        .TableCell(tcAlign, 3, 4) = 2
        .TableCell(tcText, 3, 5) = Mid$(grd.TextMatrix(i, COL_E_SERIE), 1, 6)
        .TableCell(tcFontSize, 3, 5) = 10
        .TableCell(tcAlign, 3, 5) = 0
      
        
        .TableCell(tcText, 4, 1) = " Rec." & grd.TextMatrix(i, COL_E_GAR)
        .TableCell(tcColSpan, 4, 1) = 2
        .TableCell(tcAlign, 4, 1) = 0
        
        
        .TableCell(tcText, 4, 3) = " Dis."
        .TableCell(tcAlign, 4, 3) = 2

        .TableCell(tcText, 4, 4) = grd.TextMatrix(i, COL_E_TRABAJO)
        .TableCell(tcFontSize, 4, 4) = 13
        .TableCell(tcFontBold, 4, 4) = True
        .TableCell(tcAlign, 4, 4) = 0
        .TableCell(tcText, 4, 5) = grd.TextMatrix(i, COL_E_FECHA)
        
        '----------- SEGUNDA ETIQUETA
        
        
        i = i + 1
        If i <= grd.Rows - 1 Then
            
            .TableCell(tcText, 1, 7) = "Orden"
    
            .TableCell(tcText, 1, 8) = grd.TextMatrix(i, COL_E_NUMING)
            .TableCell(tcFontSize, 1, 8) = 20
            .TableCell(tcFontBold, 1, 8) = True
            .TableCell(tcAlign, 1, 8) = 0
    
            .TableCell(tcText, 1, 9) = "Ticket"
            .TableCell(tcText, 1, 10) = grd.TextMatrix(i, COL_E_TIKET)
            .TableCell(tcFontSize, 1, 10) = 20
            .TableCell(tcFontBold, 1, 10) = True
            .TableCell(tcColSpan, 1, 10) = 2
            .TableCell(tcAlign, 1, 10) = 0
    
            .TableCell(tcText, 2, 7) = Mid$(grd.TextMatrix(i, COL_E_NOMCLI), 1, 15)
            .TableCell(tcFontSize, 2, 7) = 14
            .TableCell(tcFontBold, 2, 7) = True
            .TableCell(tcColSpan, 2, 7) = 3
            .TableCell(tcAlign, 2, 7) = 0
    
    
            .TableCell(tcText, 2, 10) = grd.TextMatrix(i, COL_E_VENDE)
            '.TableCell(tcText, 2, 10) = grd.TextMatrix(i, COL_E_CODCLI)
    
    
            .TableCell(tcText, 2, 11) = grd.TextMatrix(i, COL_E_ORDEN)
            .TableCell(tcFontSize, 2, 11) = 14
            .TableCell(tcFontBold, 2, 11) = True
    
    
            .TableCell(tcText, 3, 7) = "(" & grd.TextMatrix(i, COL_E_MARCA) & ") " & grd.TextMatrix(i, COL_E_TAMANIO)
            .TableCell(tcFontSize, 3, 7) = 10
            .TableCell(tcColSpan, 3, 7) = 3
            .TableCell(tcAlign, 3, 7) = 0
    
            .TableCell(tcText, 3, 10) = " NS:"
            .TableCell(tcAlign, 3, 10) = 2
            .TableCell(tcText, 3, 11) = Mid$(grd.TextMatrix(i, COL_E_SERIE), 1, 6)
            .TableCell(tcFontSize, 3, 11) = 10
            .TableCell(tcAlign, 3, 11) = 0
    
    
            .TableCell(tcText, 4, 7) = " Rec." & grd.TextMatrix(i, COL_E_GAR)
            .TableCell(tcColSpan, 4, 7) = 2
            .TableCell(tcAlign, 4, 7) = 0
    
            .TableCell(tcText, 4, 9) = " Dis."
            .TableCell(tcAlign, 4, 9) = 2
            .TableCell(tcText, 4, 10) = grd.TextMatrix(i, COL_E_TRABAJO)
            .TableCell(tcFontSize, 4, 10) = 13
            .TableCell(tcAlign, 4, 10) = 0
            .TableCell(tcFontBold, 4, 10) = True
    
    
            .TableCell(tcText, 4, 11) = grd.TextMatrix(i, COL_E_FECHA)
        
        End If

        .TableCell(tcRowKeepTogether) = True

        .TableCell(tcColBorderRight, 1, 6, 4, 6) = 10


        .EndTable
        
        If numEtik > 0 Then

                If (i Mod numEtik) = 0 Then
                    If i < grd.Rows - 2 Then
                        .NewPage
                    End If
                End If

        End If
        
        Next i
        .EndDoc
    End With
End Sub

Public Sub InicioNotificaciones(v As Variant, fechaPago As Date)
    grd.Redraw = flexRDNone
    grd.LoadArray v
    grd.Redraw = flexRDDirect
    FPAGO = fechaPago
    CargarDatosNotificaciones

    
    Me.Show
End Sub

Private Sub CargarDatosNotificaciones()
    Dim i As Integer, s As String, numEtik As Integer, fila As Integer, k As Integer
    Dim FilaIni As Integer, FilaFin As Integer, subtotal As Currency, numnoti As Integer
    Dim col As Integer, contpend As Integer
    ' create a long string to add to some table cells
        ' create document
    With vp
    
        .PaperSize = pprA4
        .Refresh
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpNoti_MarIzq")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpNoti_MarIzq")
            .MarginLeft = s
        Else
            .MarginLeft = 300
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpNoti_MarSup")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpNoti_MarSup")
            .MarginTop = s
        Else
            .MarginTop = 300
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpNoti_NumEtiq")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpNoti_NumEtiq")
            numEtik = CInt(s)
        Else
            numEtik = 5
        End If

        
       
        
        ' show intro
        
        .MarginBottom = 300
        .FontSize = 8
        ' set page and table borders

        .TableBorder = 4
        .TablePenLR = 10
        .TablePenTB = 10
        '.TableBorder = tbAll
        .StartDoc
        FilaIni = 1
        numnoti = 1
        For i = 1 To grd.Rows - 1
        ' build table with 4 rows
        If grd.TextMatrix(i, 1) = "Subtotal" Then
            FilaFin = i - 1
            .StartTable
            .FontName = "Arial"
            .FontSize = 10
            .TableBorder = tbBox
            .AddTable "1500|2000|2000|1500|2000|2000", "", "", RGB(200, 200, 250)
            .TableCell(tcRows) = 11
            
            ' center align all cells
            .TableCell(tcAlign) = taCenterMiddle
            .TablePenTB = 10
'            If i > 1 Then
'                If grd.TextMatrix(i, COL_E_TRANS) <> grd.TextMatrix(i - 1, COL_E_TRANS) Then
'    '                .Paragraph = " -----------------------------------------------------------------------------------------------------------"
'                    .TablePenTB = 30
'                End If
'            Else
'                .TablePenTB = 10
'            End If
    
        fila = 1
        'filafin = i

'            filafin = i - 1
        .TableCell(tcText, fila, 1) = gobjMain.EmpresaActual.GNOpcion.NombreEmpresa
        .TableCell(tcFontBold, fila, 1) = True
        .TableCell(tcColSpan, fila, 1) = 6
        .TableCell(tcFontSize, fila, 1) = 12
        
        fila = fila + 1
        .TableCell(tcText, fila, 1) = "VALORES PENDIENTES DE PAGO"
        .TableCell(tcFontBold, fila, 1) = True
        .TableCell(tcColSpan, fila, 1) = 6
        .TableCell(tcFontSize, fila, 1) = 14
        
        
        fila = fila + 1
        .TableCell(tcText, fila, 1) = "Estudiante:"
        .TableCell(tcFontBold, fila, 1) = True
        .TableCell(tcFontSize, fila, 1) = 10
        .TableCell(tcAlign, fila, 1) = 0
        
        .TableCell(tcText, fila, 2) = grd.TextMatrix(FilaFin, 2)
        .TableCell(tcFontBold, fila, 2) = False
        .TableCell(tcFontSize, fila, 2) = 10
        .TableCell(tcAlign, fila, 2) = 0
        .TableCell(tcColSpan, fila, 2) = 5
        
        
'        fila = fila + 1
'        .TableCell(tcText, fila, 1) = "Factura:"
'        .TableCell(tcFontBold, fila, 1) = True
'        .TableCell(tcFontSize, fila, 1) = 10
'        .TableCell(tcAlign, fila, 1) = 0
'
'        .TableCell(tcText, fila, 2) = grd.TextMatrix(filafin, 3)
'        .TableCell(tcFontBold, fila, 2) = False
'        .TableCell(tcFontSize, fila, 2) = 10
'        .TableCell(tcAlign, fila, 2) = 0
'        .TableCell(tcColSpan, fila, 2) = 5
'
        
        fila = fila + 1
        .TableCell(tcText, fila, 1) = "Curso:"
        .TableCell(tcFontBold, fila, 1) = True
        .TableCell(tcFontSize, fila, 1) = 10
        .TableCell(tcAlign, fila, 1) = 0
        
        .TableCell(tcText, fila, 2) = grd.TextMatrix(FilaFin, 19)
        .TableCell(tcFontBold, fila, 2) = False
        .TableCell(tcFontSize, fila, 2) = 10
        .TableCell(tcAlign, fila, 2) = 0
        .TableCell(tcColSpan, fila, 2) = 5
        
        
        fila = fila + 1
        .TableCell(tcText, fila, 1) = "Tran.Mañana:"
        .TableCell(tcFontBold, fila, 1) = True
        .TableCell(tcFontSize, fila, 1) = 10
        .TableCell(tcAlign, fila, 1) = 0
        
        .TableCell(tcText, fila, 2) = Mid$(grd.TextMatrix(FilaFin, 23) & "-" & grd.TextMatrix(FilaFin, 24), 1, 25)
        .TableCell(tcFontBold, fila, 2) = False
        .TableCell(tcFontSize, fila, 2) = 10
        .TableCell(tcAlign, fila, 2) = 0
        .TableCell(tcColSpan, fila, 2) = 2

        
        
        
        .TableCell(tcText, fila, 4) = "Tran.Tarde:"
        .TableCell(tcFontBold, fila, 4) = True
        .TableCell(tcFontSize, fila, 4) = 10
        .TableCell(tcAlign, fila, 4) = 0
        
        
        .TableCell(tcText, fila, 5) = Mid$(grd.TextMatrix(FilaFin, 25) & "-" & grd.TextMatrix(FilaFin, 26), 1, 25)
        .TableCell(tcFontBold, fila, 5) = False
        .TableCell(tcFontSize, fila, 5) = 10
        .TableCell(tcAlign, fila, 5) = 0
        .TableCell(tcColSpan, fila, 5) = 2
        
        fila = fila + 1
        .TableCell(tcText, fila, 1) = "FACTURAS PENDIENTES"
        .TableCell(tcFontBold, fila, 1) = True
        .TableCell(tcFontSize, fila, 1) = 10
        .TableCell(tcAlign, fila, 1) = 1
        .TableCell(tcColSpan, fila, 1) = 3
        
        col = 1
        subtotal = 0
        contpend = 0
        For k = FilaIni To FilaFin
            fila = fila + 1
            .TableCell(tcText, fila, col) = grd.TextMatrix(k, 4)
            .TableCell(tcFontBold, fila, col) = False
            .TableCell(tcFontSize, fila, col) = 10
            .TableCell(tcAlign, fila, col) = 0

            .TableCell(tcText, fila, col + 1) = grd.TextMatrix(k, 6)
            .TableCell(tcFontBold, fila, col + 1) = False
            .TableCell(tcFontSize, fila, col + 1) = 10
            .TableCell(tcAlign, fila, col + 1) = 0

            'valor
            .TableCell(tcText, fila, col + 2) = Format(grd.TextMatrix(k, 10), "#.00")
            .TableCell(tcFontBold, fila, col + 2) = True
            .TableCell(tcFontSize, fila, col + 2) = 10
            .TableCell(tcAlign, fila, col + 2) = 2
            
            subtotal = subtotal + grd.ValueMatrix(k, 10)
            contpend = contpend + 1
            
        Next k
        
        fila = 10
        
        .TableCell(tcText, fila, 1) = "TOTAL PENDIENTE"
        .TableCell(tcFontBold, fila, 1) = True
        .TableCell(tcFontSize, fila, 1) = 10
        .TableCell(tcAlign, fila, 1) = 0
        .TableCell(tcColSpan, fila, 1) = 2
        
        
        .TableCell(tcText, fila, col + 2) = Format(subtotal, "#.00")
        .TableCell(tcFontBold, fila, col + 2) = True
        .TableCell(tcFontSize, fila, col + 2) = 10
        .TableCell(tcAlign, fila, col + 2) = 2

        
        fila = 11
        
        If contpend = 1 Then
            .TableCell(tcText, fila, 1) = "Este valor debera ser cancelado en su totalidad máximo hasta el " & FPAGO & " al transportista o en la oficina de transporte"
        Else
            .TableCell(tcText, fila, 1) = "Este valor debera ser cancelado en su totalidad de inmedito al transportista o en la oficina de transporte"
        End If
        
        .TableCell(tcFontBold, fila, 1) = True
        .TableCell(tcFontSize, fila, 1) = 9
        .TableCell(tcAlign, fila, 1) = 0
        .TableCell(tcColSpan, fila, 1) = 6
       
       

        FilaIni = i + 1
        .EndTable
        .Paragraph = " "
        numnoti = numnoti + 1
        If numnoti > 5 Then
            .NewPage
            numnoti = 1
        End If
        
 End If
        
        Next i
        .EndDoc
    End With
End Sub


Public Sub Inicioats2016(grdcp As Object, grdfc As Object, grdretcp As Object, grdretivacp As Object, Periodo As String, numanualados As Integer)
    Dim v As Variant

            
    CargarDatosATS2016 grdcp, grdfc, grdretcp, grdretivacp, Periodo, numanualados
    Me.Show

End Sub


Private Sub CargarDatosATS2016(grdcp As Object, grdfc As Object, grdretcp As Object, grdretivacp As Object, Periodo As String, numanualados As Integer)
    Dim i As Integer, s As String, numEtik As Integer, ANEXO As Anexos
    Dim j As Integer, k As Integer, fila As Integer, COLUM As Integer
    Dim retiva As Currency, retrenta As Currency
    ' create a long string to add to some table cells
        ' create document
    With vp
    
        .PaperSize = pprA4
        .Refresh
        
       
        
        ' show intro
        .MarginLeft = 300
        .MarginRight = 100
        .MarginTop = 300
        .MarginBottom = 100
        .FontSize = 8
        ' set page and table borders

        .TableBorder = 4
        .TablePenLR = 10
        .TablePenTB = 10
        '.TableBorder = tbAll
        .StartDoc
'        For i = 1 To grd.Rows - 1
        ' build table with 4 rows
        .FontName = "Arial"
        .FontSize = 9
        .FontBold = False
        .StartTable
            .TableBorder = tbNone
            .AddTable "11000", "", " ", RGB(200, 200, 250)
            .TableCell(tcRows) = 7
            .TableCell(tcText, 1, 1) = "TALÓN RESUMEN"
            .TableCell(tcText, 2, 1) = "SERVICIO DE RENTAS INTERNAS"
            .TableCell(tcText, 3, 1) = "ANEXO TRANSACCIONAL"
            .TableCell(tcText, 4, 1) = gobjMain.EmpresaActual.GNOpcion.RazonSocial
            .TableCell(tcText, 5, 1) = "RUC " & gobjMain.EmpresaActual.GNOpcion.ruc
            .TableCell(tcText, 6, 1) = "Periodo: " & Periodo
            .TableCell(tcText, 7, 1) = "Fecha de Generacion: " & Date & " " & Time
            
            .TableCell(tcFontBold, 1, 1, 3, 1) = True
            .TableCell(tcAlign, 1, 1, 7, 1) = 1
            
        .EndTable
        .FontSize = 7
        .Paragraph = ""
        .Paragraph = ""
        .Paragraph = ""
        .Paragraph = "Certifico que la información contenida en el medio magnético del Anexo Transaccional para el período 12-2016, es fiel reflejo del siguiente reporte:"
        .Paragraph = ""
        
        .StartTable
        .FontName = "Arial"
        .FontSize = 8
        .AddTable "900|4400|1000|1200|1200|1200|1200", "", " ", RGB(200, 200, 250)
        .TableCell(tcRows) = grdcp.Rows + 1
        
        fila = 1
        COLUM = 1
        .TableCell(tcText, 1, 1) = "COMPRAS"
        .TableCell(tcColSpan, fila, COLUM) = 7
        
        .FontSize = 7
        fila = 2
        COLUM = 1
        .TableCell(tcText, fila, COLUM) = "Cod"
        
        COLUM = 2
        .TableCell(tcText, fila, COLUM) = "Transacción"
        
        COLUM = 3
        .TableCell(tcText, fila, COLUM) = "No.Registros"
        
        COLUM = 4
        .TableCell(tcText, fila, COLUM) = "BI Tarifa 0%"
        
        COLUM = 5
        .TableCell(tcText, fila, COLUM) = "BI Tarifa 12%"
        
        COLUM = 6
        .TableCell(tcText, fila, COLUM) = "BI No Objeto IVA"
        
        COLUM = 7
        .TableCell(tcText, fila, COLUM) = "Valor IVA"
        
        .TableCell(tcAlign, 1, 1, 2, 7) = 1
        .TableCell(tcFontBold, 1, 1, 2, 7) = True
        
        .FontSize = 8
        For k = 0 To grdcp.Rows - 1
            COLUM = 1
            For j = 2 To grdcp.Cols - 1
                If j = 3 Then COLUM = COLUM + 1
                If j > 3 Then
                    .TableCell(tcText, fila + k, COLUM) = Format(grdcp.TextMatrix(k, j), "#,#0.00")
                Else
                    .TableCell(tcText, fila + k, COLUM) = grdcp.TextMatrix(k, j)
                End If
                COLUM = COLUM + 1
'                If Len(grd.TextMatrix(k, 0)) = 0 Then
'                    .TableCell(tcBackColor, k, 0, k, 6) = RGB(200, 200, 250)
'                End If
            Next j
        Next k
        
        COLUM = 1
        .TableCell(tcText, fila - 1 + k, COLUM) = "TOTAL"
        .TableCell(tcColSpan, fila - 1 + k, COLUM) = 3
        
        For k = 1 To grdcp.Rows - 2
            Set ANEXO = gobjMain.EmpresaActual.RecuperaAnexos(grdcp.ValueMatrix(k, 2))
            If Not ANEXO Is Nothing Then
                .TableCell(tcText, fila + k, 2) = ANEXO.Descripcion
            End If
        Next k
        
        
        .TableCell(tcAlign, 3, 1, grdcp.Rows - 1 + 2, 3) = 1
        .TableCell(tcAlign, 3, 4, grdcp.Rows - 1 + 2, 7) = 2
        
        .TableCell(tcRowHeight, 1, 1, grdcp.Rows - 1 + 2, 7) = 250
        
        
        
        .TableBorder = tbAll
        
        .EndTable
            
        .Paragraph = ""
        .Paragraph = ""
        .StartTable
        .FontName = "Arial"
        .FontSize = 8
        .AddTable "900|4400|1000|1200|1200|1200|1200", "", " ", RGB(200, 200, 250)
        .TableCell(tcRows) = grdfc.Rows + 1
        
        fila = 1
        COLUM = 1
        .TableCell(tcText, 1, 1) = "VENTAS"
        .TableCell(tcColSpan, fila, COLUM) = 7
        
        .FontSize = 7
        fila = 2
        COLUM = 1
        .TableCell(tcText, fila, COLUM) = "Cod"
        
        COLUM = 2
        .TableCell(tcText, fila, COLUM) = "Transacción"
        
        COLUM = 3
        .TableCell(tcText, fila, COLUM) = "No.Registros"
        
        COLUM = 4
        .TableCell(tcText, fila, COLUM) = "BI Tarifa 0%"
        
        COLUM = 5
        .TableCell(tcText, fila, COLUM) = "BI Tarifa 12%"
        
        COLUM = 6
        .TableCell(tcText, fila, COLUM) = "BI No Objeto IVA"
        
        COLUM = 7
        .TableCell(tcText, fila, COLUM) = "Valor IVA"
        
        .TableCell(tcAlign, 1, 1, 2, 7) = 1
        .TableCell(tcFontBold, 1, 1, 2, 7) = True
        
        .FontSize = 8
        For k = 0 To grdfc.Rows - 1
            COLUM = 1
            For j = 2 To grdfc.Cols - 1
                If j = 3 Then COLUM = COLUM + 1
                If j > 3 Then
                    .TableCell(tcText, fila + k, COLUM) = Format(grdfc.TextMatrix(k, j), "#,#0.00")
                    
                Else
                    .TableCell(tcText, fila + k, COLUM) = grdfc.TextMatrix(k, j)
                End If
                COLUM = COLUM + 1
'                If Len(grd.TextMatrix(k, 0)) = 0 Then
'                    .TableCell(tcBackColor, k, 0, k, 6) = RGB(200, 200, 250)
'                End If
            Next j
        Next k
        
    retiva = grdfc.ValueMatrix(grdfc.Rows - 1, 8)
    retrenta = grdfc.ValueMatrix(grdfc.Rows - 1, 9)
        COLUM = 1
        .TableCell(tcText, fila - 1 + k, COLUM) = "TOTAL"
        .TableCell(tcColSpan, fila - 1 + k, COLUM) = 3

        For k = 1 To grdfc.Rows - 2
            Set ANEXO = gobjMain.EmpresaActual.RecuperaAnexos(grdfc.ValueMatrix(k, 2))
            If Not ANEXO Is Nothing Then
                .TableCell(tcText, fila + k, 2) = ANEXO.Descripcion
            End If
        Next k
        
        
        .TableCell(tcAlign, 3, 1, grdfc.Rows - 1 + 2, 3) = 1
        .TableCell(tcAlign, 3, 4, grdfc.Rows - 1 + 2, 7) = 2
        
        .TableCell(tcRowHeight, 1, 1, grdcp.Rows - 1 + 2, 7) = 250
        
'        .TableCell(tcText, FILA - 1 + k, COLUM) = "TOTAL"
'        .TableCell(tcColSpan, FILA - 1 + k, COLUM) = 2
'
        
        .TableBorder = tbAll
        
        .EndTable
        
        .Paragraph = ""
        .Paragraph = ""
        
        .StartTable
        .FontName = "Arial"
        .FontSize = 8
        .AddTable "900|4400|1000|1200|1200|1200|1200", "", " ", RGB(200, 200, 250)
        .TableCell(tcRows) = 2
        
        fila = 1
        COLUM = 1
        .TableCell(tcText, 1, 1) = "COMPROBANTES ANULADOS"
        .TableCell(tcColSpan, fila, COLUM) = 7
        .TableCell(tcAlign, 1, 1, 1, 7) = 1
        .TableCell(tcFontBold, 1, 1, 1, 7) = True

        fila = 2
        COLUM = 1
        .TableCell(tcText, fila, 1) = "Total de Comprobantes Anulados en el período informado (no incluye los dados de baja)"
        .TableCell(tcColSpan, fila, COLUM) = 6
        .TableCell(tcAlign, 2, 1, 2, 7) = 1

        .TableCell(tcText, fila, 7) = numanualados
        
        .EndTable
        
        .Paragraph = ""
        .Paragraph = ""
        
        
        .StartTable
        .FontName = "Arial"
        .FontSize = 8
        .AddTable "900|4400|1000|1200|1200|1200|1200", "", " ", RGB(200, 200, 250)
        .TableCell(tcRows) = grdretcp.Rows + 1
        
        fila = 1
        COLUM = 1
        .TableCell(tcText, 1, 1) = "RETENCION EN LA FUENTE DE IMPUESTO A LA RENTA"
        .TableCell(tcColSpan, fila, COLUM) = 7
        
        .FontSize = 7
        fila = 2
        COLUM = 1
        .TableCell(tcText, fila, COLUM) = "Cod"
        
        COLUM = 2
        .TableCell(tcText, fila, COLUM) = "Concepto de Retención"
        .TableCell(tcColSpan, fila, COLUM) = 3
        
'        COLUM = 3
'        .TableCell(tcText, FILA, COLUM) = "No.Registros"
'
'        COLUM = 4
'        .TableCell(tcText, FILA, COLUM) = "BI Tarifa 0%"
        
        COLUM = 5
        .TableCell(tcText, fila, COLUM) = "No. Registros"
        
        COLUM = 6
        .TableCell(tcText, fila, COLUM) = "Base Imponible"
        
        COLUM = 7
        .TableCell(tcText, fila, COLUM) = "Valor Retenido"
        
        .TableCell(tcAlign, 1, 1, 2, 7) = 1
        .TableCell(tcFontBold, 1, 1, 2, 7) = True
        
        .FontSize = 8
        For k = 1 To grdretcp.Rows - 1
            COLUM = 1
            For j = 2 To grdretcp.Cols - 1
                If j = 3 Then COLUM = COLUM + 3
                If j > 3 Then
                    .TableCell(tcText, fila + k, COLUM) = Format(grdretcp.TextMatrix(k, j), "#,#0.00")
                Else
                    .TableCell(tcText, fila + k, COLUM) = grdretcp.TextMatrix(k, j)
                End If
                COLUM = COLUM + 1
'                If Len(grd.TextMatrix(k, 0)) = 0 Then
'                    .TableCell(tcBackColor, k, 0, k, 6) = RGB(200, 200, 250)
'                End If
            Next j
        Next k
        COLUM = 1
        .TableCell(tcText, fila - 1 + k, COLUM) = "TOTAL"
        .TableCell(tcColSpan, fila - 1 + k, COLUM) = 5

        For k = 1 To grdretcp.Rows - 2
            Set ANEXO = gobjMain.EmpresaActual.RecuperaAnexosRetIR(grdretcp.ValueMatrix(k, 2))
            If Not ANEXO Is Nothing Then
                .TableCell(tcText, fila + k, 2) = ANEXO.DescripcionRetIR
            End If
            .TableCell(tcColSpan, fila + k, 2) = 3
        Next k


        
        .TableCell(tcAlign, 3, 1, grdretcp.Rows - 1 + 2, 3) = 1
        .TableCell(tcAlign, 3, 4, grdretcp.Rows - 1 + 2, 7) = 2
        
        .TableCell(tcRowHeight, 1, 1, grdretcp.Rows - 1 + 2, 7) = 250
'
'        .TableCell(tcText, FILA - 1 + k, COLUM) = "TOTAL"
'        .TableCell(tcColSpan, FILA - 1 + k, COLUM) = 2
        
        
        .TableBorder = tbAll
        
        .EndTable
        
        .Paragraph = ""
        .Paragraph = ""
        .StartTable
        .FontName = "Arial"
        .FontSize = 8
        .AddTable "1900|3400|1000|1200|1200|1200|1200", "", " ", RGB(200, 200, 250)
        .TableCell(tcRows) = grdretivacp.Rows + 1
        
        fila = 1
        COLUM = 1
        .TableCell(tcText, 1, 1) = "RETENCION EN LA FUENTE DE IVA"
        .TableCell(tcColSpan, fila, COLUM) = 7
        
        .FontSize = 7
        fila = 2
        COLUM = 1
        .TableCell(tcText, fila, COLUM) = "Operacion"
        
        COLUM = 2
        .TableCell(tcText, fila, COLUM) = "Concepto de Retención"
        .TableCell(tcColSpan, fila, COLUM) = 3
        
'        COLUM = 3
'        .TableCell(tcText, FILA, COLUM) = "No.Registros"
'
'        COLUM = 4
'        .TableCell(tcText, FILA, COLUM) = "BI Tarifa 0%"
        
'        COLUM = 5
'        .TableCell(tcText, FILA, COLUM) = "No. Registros"
'
'        COLUM = 6
'        .TableCell(tcText, FILA, COLUM) = "Base Imponible"
        
        COLUM = 5
        .TableCell(tcText, fila, COLUM) = "Valor Retenido"
        .TableCell(tcColSpan, fila, COLUM) = 3
        
        .TableCell(tcAlign, 1, 1, 2, 7) = 1
        .TableCell(tcFontBold, 1, 1, 2, 7) = True
        
        .FontSize = 8
        For k = 1 To grdretivacp.Rows - 1
            COLUM = 1
            For j = 2 To grdretivacp.Cols - 1
'                If j = 3 Then COLUM = COLUM + 3
                If j > 3 Then
                    .TableCell(tcText, fila + k, COLUM) = Format(grdretivacp.TextMatrix(k, j), "#,#0.00")
                    .TableCell(tcColSpan, fila + k, COLUM) = 3
                Else
                    If j = 2 Then
                        .TableCell(tcText, fila + k, COLUM) = "COMPRA"
                    Else
                        .TableCell(tcText, fila + k, COLUM) = "RETENCION IVA " & grdretivacp.ValueMatrix(k, j - 1) & " %"
                        .TableCell(tcColSpan, fila + k, COLUM) = 3
                        COLUM = COLUM + 1
                    End If
                End If
                COLUM = COLUM + 1
'                If Len(grd.TextMatrix(k, 0)) = 0 Then
'                    .TableCell(tcBackColor, k, 0, k, 6) = RGB(200, 200, 250)
'                End If
            Next j
        Next k
        COLUM = 1
        .TableCell(tcText, fila - 1 + k, COLUM) = "TOTAL"
        .TableCell(tcColSpan, fila - 1 + k, COLUM) = 4

'        For k = 1 To grdretcp.Rows - 2
'            Set ANEXO = gobjMain.EmpresaActual.RecuperaAnexosRetIR(grdretivacp.ValueMatrix(k, 2))
'            .TableCell(tcText, FILA + k, 2) = ANEXO.DescripcionRetIR
'            .TableCell(tcColSpan, FILA + k, 2) = 3
'        Next k


        
        .TableCell(tcAlign, 3, 1, grdretcp.Rows - 1 + 2, 3) = 1
        .TableCell(tcAlign, 3, 4, grdretcp.Rows - 1 + 2, 7) = 2
        
        .TableCell(tcRowHeight, 1, 1, grdretcp.Rows - 1 + 2, 7) = 250
'
'        .TableCell(tcText, FILA - 1 + k, COLUM) = "TOTAL"
'        .TableCell(tcColSpan, FILA - 1 + k, COLUM) = 2
        
        
        .TableBorder = tbAll
        
        .EndTable
        
        .Paragraph = ""
        .Paragraph = ""
        .StartTable
        .FontName = "Arial"
        .FontSize = 8
        .AddTable "1900|3400|1000|1200|1200|1200|1200", "", " ", RGB(200, 200, 250)
        .TableCell(tcRows) = 5
        
        fila = 1
        COLUM = 1
        .TableCell(tcText, 1, 1) = "RESUMEN DE RETENCIONES QUE LE EFECTUARON EN EL PERIODO"
        .TableCell(tcColSpan, fila, COLUM) = 7
        
        .FontSize = 7
        fila = 2
        COLUM = 1
        .TableCell(tcText, fila, COLUM) = "Operacion"
        
        COLUM = 2
        .TableCell(tcText, fila, COLUM) = "Concepto de Retención"
        .TableCell(tcColSpan, fila, COLUM) = 3
        
        COLUM = 5
        .TableCell(tcText, fila, COLUM) = "Valor Retenido"
        .TableCell(tcColSpan, fila, COLUM) = 3
        
        .TableCell(tcAlign, 1, 1, 2, 7) = 1
        .TableCell(tcFontBold, 1, 1, 2, 7) = True
        
        .FontSize = 8
            k = 1
            COLUM = 1
            .TableCell(tcText, fila + k, COLUM) = "VENTA"
            COLUM = 2
            .TableCell(tcText, fila + k, COLUM) = "Valor de IVA que le han retenido"
            .TableCell(tcColSpan, fila + k, COLUM) = 3
            COLUM = 5
            .TableCell(tcText, fila + k, COLUM) = Format(retiva, "#,#0.00")
            .TableCell(tcColSpan, fila + k, COLUM) = 3
            
            k = 2
            COLUM = 1
            .TableCell(tcText, fila + k, COLUM) = "VENTA"
            COLUM = 2
            .TableCell(tcText, fila + k, COLUM) = "Valor de Renta que le han retenido"
            .TableCell(tcColSpan, fila + k, COLUM) = 3
            COLUM = 5
            .TableCell(tcText, fila + k, COLUM) = Format(retrenta, "#,#0.00")
            .TableCell(tcColSpan, fila + k, COLUM) = 3
            k = 4
        
        
        

        
        
        COLUM = 1
        .TableCell(tcText, fila - 1 + k, COLUM) = "TOTAL"
        .TableCell(tcColSpan, fila - 1 + k, COLUM) = 4

        COLUM = 5
        .TableCell(tcText, fila - 1 + k, COLUM) = Format(retiva + retrenta, "#,#0.00")
        .TableCell(tcColSpan, fila - 1 + k, COLUM) = 4

'        For k = 1 To grdretcp.Rows - 2
'            Set ANEXO = gobjMain.EmpresaActual.RecuperaAnexosRetIR(grdretivacp.ValueMatrix(k, 2))
'            .TableCell(tcText, FILA + k, 2) = ANEXO.DescripcionRetIR
'            .TableCell(tcColSpan, FILA + k, 2) = 3
'        Next k


        
        .TableCell(tcAlign, 3, 1, grdretcp.Rows - 1 + 2, 3) = 1
        .TableCell(tcAlign, 3, 4, grdretcp.Rows - 1 + 2, 7) = 2
        
        .TableCell(tcRowHeight, 1, 1, grdretcp.Rows - 1 + 2, 7) = 250
'
'        .TableCell(tcText, FILA - 1 + k, COLUM) = "TOTAL"
'        .TableCell(tcColSpan, FILA - 1 + k, COLUM) = 2
        
        
        .TableBorder = tbAll
        
        .EndTable
        
        .Paragraph = ""
        .FontSize = 5
        .Paragraph = "Declaro que los datos contenidos en este anexo son verdaderos, por lo que asumo la responsabilidad correspondiente, de acuerdo a lo establecido en el Art. 101 de la Codificación de la Ley de Régimen Tributario Interno"
        .Paragraph = ""
        
        .HdrFontName = "Sansation"
        .HdrFontSize = 8
        .Header = "Generado  por IBZ Ishida Businnes Software"
        
 '       Next i
        .EndDoc
    End With
End Sub

