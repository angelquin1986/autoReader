VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmItemCostosNew 
   BackColor       =   &H8000000A&
   Caption         =   "Costos"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9720
   FillColor       =   &H00C00000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   9720
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   5280
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   3135
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   6135
      _cx             =   10821
      _cy             =   5530
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
      Rows            =   3
      Cols            =   5
      FixedRows       =   1
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
      SubtotalPosition=   1
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
      Editable        =   2
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   12648447
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSComctlLib.ImageList img1 
      Left            =   9120
      Top             =   600
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ItemCostoNew.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ItemCostoNew.frx":0114
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ItemCostoNew.frx":0568
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ItemCostoNew.frx":067C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ItemCostoNew.frx":0790
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ItemCostoNew.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ItemCostoNew.frx":0E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ItemCostoNew.frx":1B20
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ItemCostoNew.frx":1F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ItemCostoNew.frx":2084
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ItemCostoNew.frx":3906
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   1005
      ButtonWidth     =   2540
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "img1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Configura/Busca"
            Key             =   "Configurar"
            Object.ToolTipText     =   "Configurar"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Abrir"
            Key             =   "Abrir"
            Object.ToolTipText     =   "Abir Archivo"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Asignar"
            Key             =   "Asignar"
            Description     =   "Asignar un valor"
            Object.ToolTipText     =   "Asignar un valor "
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Calcular (F5)"
            Key             =   "Calcular"
            Object.ToolTipText     =   "Calcular "
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exp. Excel"
            Key             =   "Excel"
            Object.ToolTipText     =   "A Excel"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Guardar Resul."
            Key             =   "Guardar"
            Object.ToolTipText     =   "Guardar Resultado"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cerrar"
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar (ESC)"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   9720
      TabIndex        =   0
      Top             =   4185
      Width           =   9720
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdCos 
      Height          =   1215
      Left            =   1080
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   3615
      _cx             =   6376
      _cy             =   2143
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
      Rows            =   3
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
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
      SubtotalPosition=   1
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
      Editable        =   2
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   12648447
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "frmItemCostosNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ex As Excel.Application, ws As Worksheet, wkb As Workbook
Private mProcesando As Boolean
Private mCancelado As Boolean
Private mcolItemsSelec As Collection      'Coleccion de items
'jeaa 24/09/04 asignacion de grupo a los items
'IVINVENTARIO
Const IVGRUPO1 = 4
Const IVGRUPO2 = 5
Const IVGRUPO3 = 6
Const IVGRUPO4 = 7
'PC_PROV_CLI
Const PCGRUPO1 = 3
Const PCGRUPO2 = 4
Const PCGRUPO3 = 5
Const PCGRUPO4 = 6 'Agregado AUC 03/10/2005
Dim v() As String
Dim costoFijoMensual As Currency, Precio As Integer



Const COL_BODEGA = 1
Const COL_CODGRUPO = 2
Const COL_DESCGRUPO = 3
Const COL_PVP = 4
Const COL_PESO = 5
Const COL_KILOSPROD = 6
Const COL_KILOSVEN = 7
Const COL_KILOSDEV = 8
Const COL_KILOSMEM = 9
Const COL_COSTOFIJOUNIT = 10
Const COL_COSTOVARIABLE = 11
Const Col_CostoTotal = 12
Const COL_UTINETAUNI = 13
Const COL_UTITOTALNETA = 14
Const COL_MARGEN = 15
Const COL_PORCEM2 = 16
Const COL_PRODM2 = 17
Const COL_VENTAM2 = 18
Const COL_DEVOLM2 = 19
Const COL_MEM_M2 = 20
Const COL_VENTAUSD = 21
Const COL_VENTAUSDDEV = 22
Const COL_VENTAUSDMEM = 23
Const COL_FACT = 24



Public Sub Inicio(ByVal tag As String)
    Dim rutaPlantilla
    Dim i As Integer
    Dim valor As Currency
    On Error GoTo errtrap
    
    Me.tag = tag            'Guarda en la propiedad Tag para distinguir después
    Me.Show
    Me.ZOrder
    Select Case Me.tag
    Case "Costos"
        Me.Caption = "Punto de Equilibrio basado en Producción "
    Case "Produccion"
         Me.Caption = "Punto de Equilibrio basado en Ventas "
    End Select
       
    'Inicializa la grilla
    rutaPlantilla = GetSetting(APPNAME, App.Title, "Ruta Plantilla", "")
    grd.Rows = grd.FixedRows
   ConfigCols
    VisualizarTexto (rutaPlantilla)
    For i = 1 To 8 'Agrega filas adicionales
            grd.AddItem vbTab '& "."
        Next
   ConfigCols
    FijarColor ': CargarConfiguarciones
    CargarBodegas
    CargarConfiguarciones
            
    Exit Sub
errtrap:
    DispErr
    Unload Me
    Exit Sub
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF3
        KeyCode = 0
    Case vbKeyF5
          Calcular
          GuardaConfM2
    Case vbKeyF6
        KeyCode = 0
    Case vbKeyP
        If Shift And vbCtrlMask Then
            KeyCode = 0
        End If
    Case vbKeyEscape
        Cerrar
    Case Else
        MoverCampo Me, KeyCode, Shift, True
    End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    ImpideSonidoEnter Me, KeyAscii
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mProcesando Then
        Cancel = 1      'No permitir cerrar mientras procesa
    Else
        Me.Hide         'Se pone esto para evitar el posible BUG de Windows98
    End If
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    grd.Move 0, tlb1.Height, Me.ScaleWidth, Me.ScaleHeight - tlb1.Height - pic1.Height - 80
    prg1.Width = Me.ScaleWidth - (prg1.Left * 2)
End Sub


Private Sub grd_BeforeEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
    If Row < grd.FixedRows Then Cancel = True
    If grd.IsSubtotal(Row) = True Then Cancel = True
    If grd.ColData(col) < 0 Then Cancel = True
    
    
    If grd.CellBackColor = grd.BackColorFrozen Or grd.CellBackColor = &HC00000 Then
       Cancel = True
    End If
End Sub

Private Sub grd_BeforeSort(ByVal col As Long, Order As Integer)
    'Impide mientras está procesando
    If mProcesando Then Order = flexSortNone
End Sub
Private Sub grd_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
    Case vbKeyDelete
        EliminaFila
    End Select
End Sub
Private Sub grd_KeyPressEdit(ByVal Row As Long, ByVal col As Long, KeyAscii As Integer)
Select Case grd.ColDataType(col)
    Case flexDTCurrency, flexDTSingle, flexDTDouble
        If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
           (KeyAscii <> vbKeyBack) And _
           (KeyAscii <> Asc(".")) And _
           (KeyAscii <> Asc("-")) And _
           (KeyAscii <> vbKeyReturn) Then
            KeyAscii = 0
        End If
    Case flexDTLong
        If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
           (KeyAscii <> vbKeyBack) And _
           (KeyAscii <> vbKeyReturn) Then
            KeyAscii = 0
        End If
    End Select
End Sub

Private Sub tlb1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Configurar":
        Configurar
    Case "Abrir"
        'No hace nada
    Case "Asignar": Asignar
    Case "Calcular":
          Calcular
          GuardaConfM2
    Case "Excel": ExportaExcel ("Calculo de Costos")
    Case "Guardar": GuardarResultado
    Case "Cerrar":      Cerrar
    End Select
End Sub
Private Sub ConfigCols(Optional subt As Integer)
    Dim s As String, i As Long, j As Integer, s1 As String
    Dim fmt As String
    With grd
               s = "^#|<bodega|<Codigo|<Descripción|>P.V.|>Peso"
               s = s & "|>Total Kilos Producidos|>Total Kilos Vendidos|>Total Kilos Devol|>Total Kilos MEM"
               s = s & "|>Costo Fijo Unitario|>Costo Variable|>Costo Total|>Utilidad Neta Unitaria"
               s = s & "|>Utilidad Total NETA|>Margen|>Porcentaje M2|>Produccion M2"
               s = s & "|>Ventas M2|>Devol M2|>MEM M2"
               s = s & "|>Total Ventas $|>Total Devol  $|>Total MEM $|>FactConv"
  
        .FormatString = s
        AjustarAutoSize grd, -1, -1, 4000
        AsignarTituloAColKey grd
        'grilla de resultados
        'Columnas modificables (Longitud maxima)
        For i = 0 To COL_CODGRUPO
            .ColData(i) = -1
        Next i
        .ColData(COL_KILOSPROD) = -1
        .ColData(COL_COSTOFIJOUNIT) = -1
         .ColData(Col_CostoTotal) = -1
         .ColData(COL_UTINETAUNI) = -1
         .ColData(COL_UTITOTALNETA) = -1
         .ColData(COL_MARGEN) = -1
         .ColData(COL_KILOSVEN) = -1
         .ColData(COL_KILOSMEM) = -1
         .ColData(COL_KILOSDEV) = -1
        'Color de fondo
        If .Rows > .FixedRows Then
            .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, COL_DESCGRUPO) = .BackColorFrozen
            .Cell(flexcpBackColor, .FixedRows, COL_KILOSPROD, .Rows - 1, COL_KILOSPROD) = .BackColorFrozen
            .Cell(flexcpBackColor, .FixedRows, COL_KILOSVEN, .Rows - 1, COL_KILOSVEN) = .BackColorFrozen
            .Cell(flexcpBackColor, .FixedRows, COL_KILOSMEM, .Rows - 1, COL_KILOSMEM) = .BackColorFrozen
            .Cell(flexcpBackColor, .FixedRows, COL_KILOSDEV, .Rows - 1, COL_KILOSDEV) = .BackColorFrozen
            .Cell(flexcpBackColor, .FixedRows, COL_COSTOFIJOUNIT, .Rows - 1, COL_COSTOFIJOUNIT) = .BackColorFrozen
            .Cell(flexcpBackColor, .FixedRows, Col_CostoTotal, .Rows - 1, Col_CostoTotal) = .BackColorFrozen
            .Cell(flexcpBackColor, .FixedRows, COL_UTINETAUNI, .Rows - 1, COL_UTINETAUNI) = .BackColorFrozen
            .Cell(flexcpBackColor, .FixedRows, COL_UTITOTALNETA, .Rows - 1, COL_UTITOTALNETA) = .BackColorFrozen
            .Cell(flexcpBackColor, .FixedRows, COL_MARGEN, .Rows - 1, COL_MARGEN) = .BackColorFrozen
            .TextMatrix(.Rows - 8, COL_CODGRUPO) = "Coeficiente"
            .TextMatrix(.Rows - 7, COL_CODGRUPO) = "MargenGlobal"
            .TextMatrix(.Rows - 6, COL_CODGRUPO) = "CostoVariableTotal"
            .TextMatrix(.Rows - 5, COL_CODGRUPO) = "Utilidad Neta"
            .TextMatrix(.Rows - 4, COL_CODGRUPO) = "Anual"
            .TextMatrix(.Rows - 3, COL_CODGRUPO) = "CostoFijo/VentasTotal"
            .TextMatrix(.Rows - 2, COL_CODGRUPO) = "CostoFijoMensual"
            'formatos
            For i = 4 To grd.Cols - 1
              .ColDataType(i) = flexDTCurrency
              .ColFormat(i) = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda(fmt)
            Next
            
            
            
'            .ColDataType(i) = flexDTCurrency
            .ColFormat(COL_MARGEN) = "#.##%"   ' percentage
            .ColFormat(COL_PORCEM2) = "#.##%"   ' percentage

           End If
           If .Rows > .FixedRows Then
             
                          'ultimas filas PONE COLOR AZUL
              grd.Cell(flexcpBackColor, .Rows - 8, 1, .Rows - 1, grd.Cols - 1) = &HC00000 'color de fondo
              grd.Cell(flexcpForeColor, .Rows - 8, 1, .Rows - 1, grd.Cols - 1) = &HFFFF&  'color de letras


                .MergeCells = flexMergeFree
                .MergeCol(0) = True: .MergeCol(1) = True: .MergeCol(2) = True
                .SubtotalPosition = flexSTBelow
                 grd.subtotal flexSTSum, 1, COL_KILOSPROD, , grd.BackColorFixed, , True, " Subtotal", 1, False
                 grd.subtotal flexSTSum, 1, COL_KILOSVEN, , grd.BackColorFixed, , True, " Subtotal", 1, False
                 grd.subtotal flexSTSum, 1, COL_KILOSDEV, , grd.BackColorFixed, , True, " Subtotal", 1, False
                 grd.subtotal flexSTSum, 1, COL_KILOSMEM, , grd.BackColorFixed, , True, " Subtotal", 1, False
                 grd.subtotal flexSTSum, 1, COL_UTITOTALNETA, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, COL_PRODM2, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, COL_VENTAM2, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, COL_DEVOLM2, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, COL_MEM_M2, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, COL_VENTAUSD, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, COL_VENTAUSDDEV, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, COL_VENTAUSDMEM, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, -1, 6, , grd.BackColorFixed, , True, " Total", 1, False
                 
                 grd.FrozenCols = 3
                        For i = 1 To grd.Cols - 1
                            grd.TextMatrix(grd.Rows - 2, i) = 0 'Encera la ultima fila
                            grd.TextMatrix(grd.Rows - 2, i) = "" 'Encera la ultima fila
                        Next
           End If
    End With
End Sub
Private Sub Cerrar()
    If mProcesando Then
        'Si está procesando, pregunta si quere abandonarlo
        If MsgBox("Desea abandonar el proceso?", vbQuestion + vbYesNo) = vbYes Then
            mCancelado = True
        End If
        
        Exit Sub
    Else
        Unload Me
    End If
End Sub
Private Sub Calcular()
Dim i As Integer, j As Integer
Dim TotalKilosPro As Currency, ventasTotal As Currency, UtilidadTotalNeta As Currency
Dim MargenGlobal As Currency, totalPronostico, costoVariableTotal As Currency
Dim TotalKilosVen As Currency, TotalKilosMEM As Currency, TotalKilosDEV As Currency
Dim ValorM2 As Currency
'encera valores iniciales
TotalKilosPro = 0
ventasTotal = 0
UtilidadTotalNeta = 0
MargenGlobal = 0
totalPronostico = 0
costoVariableTotal = 0
TotalKilosVen = 0
costoFijoMensual = grd.ValueMatrix(grd.Rows - 3, COL_DESCGRUPO)
         grd.subtotal flexSTSum, 1, 6, , grd.BackColorFixed, , True, " Subtotal", 1, False
                 grd.subtotal flexSTSum, 1, COL_KILOSPROD, , grd.BackColorFixed, , True, " Subtotal", 1, False
                 grd.subtotal flexSTSum, 1, COL_KILOSVEN, , grd.BackColorFixed, , True, " Subtotal", 1, False
                 grd.subtotal flexSTSum, 1, COL_KILOSDEV, , grd.BackColorFixed, , True, " Subtotal", 1, False
                 grd.subtotal flexSTSum, 1, COL_KILOSMEM, , grd.BackColorFixed, , True, " Subtotal", 1, False
                 grd.subtotal flexSTSum, 1, COL_UTITOTALNETA, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, COL_PRODM2, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, COL_VENTAM2, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, COL_DEVOLM2, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, COL_MEM_M2, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, COL_VENTAUSD, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, COL_VENTAUSDDEV, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, COL_VENTAUSDMEM, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, -1, 6, , grd.BackColorFixed, , True, " Total", 1, False
        
With grd
   For j = 1 To grd.Rows - 10
      If Not .IsSubtotal(j) Then
        If .ValueMatrix(j, COL_PRODM2) <> 0 Then
             .TextMatrix(j, COL_KILOSPROD) = .ValueMatrix(j, COL_PESO) * .ValueMatrix(j, COL_PRODM2)
        Else
            .TextMatrix(j, COL_KILOSPROD) = 0
        End If
         If Len(.TextMatrix(j, COL_VENTAM2)) > 0 Then
            .TextMatrix(j, COL_KILOSVEN) = .ValueMatrix(j, COL_VENTAM2) * .ValueMatrix(j, COL_PESO)
        Else
            .TextMatrix(j, COL_KILOSVEN) = 0
        End If
        If Len(.TextMatrix(j, COL_MEM_M2)) > 0 Then
            .TextMatrix(j, COL_KILOSMEM) = .ValueMatrix(j, COL_MEM_M2) * .ValueMatrix(j, COL_PESO)
        Else
            .TextMatrix(j, COL_KILOSMEM) = 0
        End If

        If Len(.TextMatrix(j, COL_DEVOLM2)) > 0 Then
            .TextMatrix(j, COL_KILOSDEV) = .ValueMatrix(j, COL_DEVOLM2) * .ValueMatrix(j, COL_PESO)
        Else
            .TextMatrix(j, COL_KILOSDEV) = 0
        End If
        
            TotalKilosPro = TotalKilosPro + .ValueMatrix(j, COL_KILOSPROD)
            TotalKilosVen = TotalKilosVen + .ValueMatrix(j, COL_KILOSVEN)
            TotalKilosMEM = TotalKilosMEM + .ValueMatrix(j, COL_KILOSMEM)
            TotalKilosDEV = TotalKilosDEV + .ValueMatrix(j, COL_KILOSDEV)

     End If
   Next
        Select Case Me.tag
         Case "Costos"
            If TotalKilosPro <> 0 Then
                .TextMatrix(.Rows - 10, COL_DESCGRUPO) = Format(.ValueMatrix(.Rows - 4, COL_DESCGRUPO) / TotalKilosPro, "#0.0000")
            Else
                .TextMatrix(.Rows - 10, COL_DESCGRUPO) = 0
            End If
         
         Case "Produccion"
            If TotalKilosVen <> 0 Then
                .TextMatrix(.Rows - 10, COL_DESCGRUPO) = Format(.ValueMatrix(.Rows - 4, COL_DESCGRUPO) / (TotalKilosVen + TotalKilosDEV), "#0.0000")
            Else
                .TextMatrix(.Rows - 10, COL_DESCGRUPO) = 0
            End If
        End Select
   For j = 1 To grd.Rows - 10
   If Not .IsSubtotal(j) Then
      If .ValueMatrix(j, COL_PESO) > 0 Then
        .TextMatrix(j, COL_COSTOFIJOUNIT) = .ValueMatrix(j, COL_PESO) * .ValueMatrix(.Rows - 10, COL_DESCGRUPO)
      End If
      If .ValueMatrix(j, COL_COSTOVARIABLE) <> 0 And .ValueMatrix(j, COL_COSTOFIJOUNIT) <> 0 Then
        .TextMatrix(j, Col_CostoTotal) = .ValueMatrix(j, COL_COSTOVARIABLE) + .ValueMatrix(j, COL_COSTOFIJOUNIT)
      End If
      If .ValueMatrix(j, COL_PVP) <> 0 And .ValueMatrix(j, Col_CostoTotal) <> 0 Then
        .TextMatrix(j, COL_UTINETAUNI) = .ValueMatrix(j, COL_PVP) - .ValueMatrix(j, Col_CostoTotal)
      End If
      
    Select Case Me.tag
         Case "Costos"
            If (.ValueMatrix(j, COL_UTINETAUNI)) <> 0 And (.ValueMatrix(j, COL_PRODM2)) <> 0 Then
              .TextMatrix(j, COL_UTITOTALNETA) = .ValueMatrix(j, COL_UTINETAUNI) * .ValueMatrix(j, COL_PRODM2)
            End If
            If (.ValueMatrix(j, Col_CostoTotal)) <> 0 Then
                .TextMatrix(j, COL_MARGEN) = .ValueMatrix(j, COL_UTINETAUNI) / .ValueMatrix(j, Col_CostoTotal)
            End If
            
        Case "Produccion"
            If (.ValueMatrix(j, COL_UTINETAUNI)) <> 0 And (.ValueMatrix(j, COL_VENTAM2)) <> 0 Then
              .TextMatrix(j, COL_UTITOTALNETA) = (.ValueMatrix(j, COL_UTINETAUNI) * .ValueMatrix(j, COL_VENTAM2)) + (.ValueMatrix(j, COL_UTINETAUNI) * .ValueMatrix(j, COL_DEVOLM2))
            End If
            If (.ValueMatrix(j, COL_VENTAUSD)) <> 0 Then
                If (.ValueMatrix(j, COL_VENTAUSDDEV)) <> 0 Then
                    .TextMatrix(j, COL_MARGEN) = (.ValueMatrix(j, COL_UTITOTALNETA) / .ValueMatrix(j, COL_VENTAUSD)) + (.ValueMatrix(j, COL_UTITOTALNETA) / .ValueMatrix(j, COL_VENTAUSDDEV))
                Else
                    .TextMatrix(j, COL_MARGEN) = (.ValueMatrix(j, COL_UTITOTALNETA) / .ValueMatrix(j, COL_VENTAUSD))
                End If
            End If
            
        End Select
      
        
      'CALCULA Total Ventas $
      ventasTotal = ventasTotal + .ValueMatrix(j, COL_VENTAUSD) + .ValueMatrix(j, COL_VENTAUSDDEV)
      'calculta suma utilidad total neta
      UtilidadTotalNeta = UtilidadTotalNeta + .ValueMatrix(j, COL_UTITOTALNETA)
      'calcula totalpronostico
     totalPronostico = totalPronostico + .ValueMatrix(j, COL_PRODM2)
     'Calcula costovariabletotal
     Select Case Me.tag
         Case "Costos"
            costoVariableTotal = costoVariableTotal + (.ValueMatrix(j, COL_PRODM2) * .ValueMatrix(j, COL_COSTOVARIABLE))
        Case "Produccion"
            costoVariableTotal = costoVariableTotal + (.ValueMatrix(j, COL_VENTAM2) * .ValueMatrix(j, COL_COSTOVARIABLE)) + (.ValueMatrix(j, COL_DEVOLM2) * .ValueMatrix(j, COL_COSTOVARIABLE))
    End Select
   End If
   Next
      'asigna ventastotal
      If ventasTotal <> 0 Then
        MargenGlobal = UtilidadTotalNeta / ventasTotal ' .ValueMatrix(.Rows - 1, COL_VENTAUSD)
      End If
      .TextMatrix(grd.Rows - 9, COL_DESCGRUPO) = Round(MargenGlobal * 100, 2) & " %"
      .TextMatrix(grd.Rows - 8, COL_DESCGRUPO) = Format(costoVariableTotal, "###,#0.00")
      If ventasTotal <> 0 Then
          .TextMatrix(grd.Rows - 4, COL_PVP) = Round((.ValueMatrix(grd.Rows - 4, COL_DESCGRUPO) / ventasTotal) * 100, 2) & " %"
          .TextMatrix(grd.Rows - 8, COL_PVP) = Round((costoVariableTotal / ventasTotal) * 100, 2) & " %"
      End If
      'UTILIDADNETA
      '.TextMatrix(grd.Rows - 7, COL_DESCGRUPO) = ventasTotal - .ValueMatrix(grd.Rows - 4, COL_DESCGRUPO) - costoVariableTotal
      .TextMatrix(grd.Rows - 7, COL_DESCGRUPO) = Format(UtilidadTotalNeta, "###,#0.00")
      ' Porcentaje M2 UTILIDADNETA
      If ventasTotal <> 0 Then
        .TextMatrix(grd.Rows - 7, COL_PVP) = Round((.ValueMatrix(grd.Rows - 7, COL_DESCGRUPO) / ventasTotal) * 100, 2) & " %"
      End If
      'anual
      '.TextMatrix(grd.Rows - 6, COL_DESCGRUPO) = .ValueMatrix(grd.Rows - 7, COL_DESCGRUPO) * 12
End With

                 grd.subtotal flexSTSum, 1, COL_KILOSPROD, , grd.BackColorFixed, , True, " Subtotal", 1, False
                 grd.subtotal flexSTSum, 1, COL_KILOSVEN, , grd.BackColorFixed, , True, " Subtotal", 1, False
                 grd.subtotal flexSTSum, 1, COL_KILOSDEV, , grd.BackColorFixed, , True, " Subtotal", 1, False
                 grd.subtotal flexSTSum, 1, COL_KILOSMEM, , grd.BackColorFixed, , True, " Subtotal", 1, False
                 grd.subtotal flexSTSum, 1, COL_UTITOTALNETA, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, COL_PRODM2, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, COL_VENTAM2, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, COL_DEVOLM2, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, COL_MEM_M2, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, COL_VENTAUSD, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, COL_VENTAUSDDEV, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, COL_VENTAUSDMEM, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, -1, 6, , grd.BackColorFixed, , True, " Total", 1, False



                 grd.subtotal flexSTSum, -1, COL_KILOSPROD, , grd.BackColorFixed, , True, " Total", 1, False
                 grd.subtotal flexSTSum, -1, COL_KILOSVEN, , grd.BackColorFixed, , True, " Total", 1, False
                 grd.subtotal flexSTSum, -1, COL_KILOSDEV, , grd.BackColorFixed, , True, " Total", 1, False
                 grd.subtotal flexSTSum, -1, COL_KILOSMEM, , grd.BackColorFixed, , True, " Total", 1, False
                 grd.subtotal flexSTSum, -1, COL_UTITOTALNETA, , grd.BackColorFixed, , True, " Total", 1, True
                 grd.subtotal flexSTSum, -1, COL_PRODM2, , grd.BackColorFixed, , True, " Total", 1, True
                 grd.subtotal flexSTSum, -1, COL_VENTAM2, , grd.BackColorFixed, , True, " Total", 1, True
                 grd.subtotal flexSTSum, -1, COL_DEVOLM2, , grd.BackColorFixed, , True, " Total", 1, True
                 grd.subtotal flexSTSum, -1, COL_MEM_M2, , grd.BackColorFixed, , True, " Total", 1, True
                 grd.subtotal flexSTSum, -1, COL_VENTAUSD, , grd.BackColorFixed, , True, " Total", 1, True
                 grd.subtotal flexSTSum, -1, COL_VENTAUSDDEV, , grd.BackColorFixed, , True, " Total", 1, True
                 grd.subtotal flexSTSum, -1, COL_VENTAUSDMEM, , grd.BackColorFixed, , True, " Total", 1, True
                 grd.subtotal flexSTSum, -1, 6, , grd.BackColorFixed, , True, " Total", 1, False


For i = 1 To grd.Cols - 1
    grd.TextMatrix(grd.Rows - 2, i) = 0 'Encera la ultima fila
    grd.TextMatrix(grd.Rows - 2, i) = "" 'Encera la ultima fila
Next
grd.Select grd.Rows - 9, COL_DESCGRUPO, grd.Rows - 9, COL_DESCGRUPO
grd.CellAlignment = flexAlignRightBottom
grd.Select grd.Rows - 8, COL_DESCGRUPO, grd.Rows - 8, COL_DESCGRUPO
grd.CellAlignment = flexAlignRightBottom
grd.Select grd.Rows - 7, COL_DESCGRUPO, grd.Rows - 7, COL_DESCGRUPO
grd.CellAlignment = flexAlignRightBottom
grd.Select grd.Rows - 6, COL_DESCGRUPO, grd.Rows - 6, COL_DESCGRUPO
grd.CellAlignment = flexAlignRightBottom
grd.Select grd.Rows - 5, COL_DESCGRUPO, grd.Rows - 5, COL_DESCGRUPO
grd.CellAlignment = flexAlignRightBottom
grd.Select grd.Rows - 4, COL_DESCGRUPO, grd.Rows - 4, COL_DESCGRUPO
grd.CellAlignment = flexAlignRightBottom
grd.Select grd.Rows - 3, COL_DESCGRUPO, grd.Rows - 3, COL_DESCGRUPO
grd.CellAlignment = flexAlignRightBottom
grd.Select grd.Rows - 2, COL_DESCGRUPO, grd.Rows - 2, COL_DESCGRUPO
grd.CellAlignment = flexAlignRightBottom
grd.Select grd.Rows - 10, COL_DESCGRUPO, grd.Rows - 10, COL_DESCGRUPO
grd.CellAlignment = flexAlignRightBottom

With grd
    For j = 1 To grd.Rows - 10
        If .IsSubtotal(j) Then
            If .ValueMatrix(j, COL_VENTAUSD) <> 0 Then
                .TextMatrix(j, COL_MARGEN) = .ValueMatrix(j, COL_UTITOTALNETA) / .ValueMatrix(j, COL_VENTAUSD)
            End If
        End If
    Next j
j = grd.Rows - 1
If .ValueMatrix(j, COL_VENTAUSD) <> 0 Then
    .TextMatrix(j, COL_MARGEN) = .ValueMatrix(j, COL_UTITOTALNETA) / .ValueMatrix(j, COL_VENTAUSD)
End If
.RowHidden(grd.Rows - 5) = True
.RowHidden(grd.Rows - 6) = True
End With
End Sub
Private Sub EliminaFila()
    Dim msg As String, r As Long
    On Error GoTo errtrap
    r = grd.Row
    If (grd.Rows > grd.FixedRows) Then
        If (Not grd.IsSubtotal(r)) Then
            msg = "Desea eliminar la fila #" & r & "?"
            If MsgBox(msg, vbYesNo) <> vbYes Then
                grd.SetFocus
                Exit Sub
            End If
            If grd.TextMatrix(r, COL_BODEGA) = "." Then: Exit Sub
            'Elimina del grid
            grd.RemoveItem r
        End If
    End If
    grd.SetFocus
    Exit Sub
errtrap:
    DispErr
    grd.SetFocus
    Exit Sub
End Sub



Public Sub ExportaExcel(ByVal titulo As String)
    'tipo=0 Roles de Pagos; tipo=1 Reporte Bancos y Provisiones
    Set ex = New Excel.Application  'Crea un instancia nueva de excel
    Set wkb = ex.Workbooks.Add  'Insertar un libro nuevo
    Set ws = ex.Worksheets.Add  'Inserta una nueva hoja
    With ws
        .Name = titulo
        .Range("A1").Font.Name = "Times Roman"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Cells(1) = gobjMain.EmpresaActual.GNOpcion.NombreEmpresa
    End With
        Exportar titulo
    ex.Visible = True
    ws.Activate
    Set ws = Nothing
    Set wkb = Nothing
    Set ex = Nothing
    'ex.Quit
End Sub
Private Sub Exportar(ByVal titulo As String)
        Dim fila As Long, col As Long, i As Long, j As Long
    Dim v() As Long, mayor As Long
    Dim NumCol As Integer
    Dim fmt As String
    prg1.min = 0
    prg1.max = grd.Rows - 1
    MensajeStatus "Está Exportando  a Excel ...", vbHourglass
    With ws
        fila = 2
        .Range("H1").Font.Name = "Arial"
        .Range("H1").Font.Size = 12
        .Range("H1").Font.Bold = True
        .Cells(fila, 1) = titulo
        
        .PageSetup.PaperSize = xlPaperLetter 'Tamaño del papel (carta)
        .PageSetup.BottomMargin = Application.CentimetersToPoints(1.5) 'Margen Superior
        .PageSetup.TopMargin = Application.CentimetersToPoints(1) 'Margen Inferior
        .Range(.Cells(1, 13), .Cells(500, 13)).NumberFormat = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda(fmt)   'Establece el formato para los números
        .Range("A2:AZ1000").Font.Name = "Arial"    'Tipo de letra para toda la hoja
        .Range("A2:AZ1000").Font.Size = 8          'Tamaño de la letra
        
        fila = fila + 2
        NumCol = 0
        For i = 1 To grd.Cols - 1
            NumCol = NumCol + 1
            .Cells(fila, NumCol) = grd.TextMatrix(0, i) 'cabeceras
            ReDim Preserve v(NumCol)
            v(NumCol - 1) = 0
        Next i
                
        .Range(.Cells(fila, 1), .Cells(fila, NumCol)).Font.Bold = True
        .Range(.Cells(fila, 1), .Cells(fila, NumCol)).Borders.LineStyle = 12
        .Range(.Cells(fila, 3), .Cells(fila, grd.Rows - 1)).HorizontalAlignment = xlHAlignLeft
        For i = 2 To grd.Rows - 1
            prg1.value = i
            fila = fila + 1
            If grd.IsSubtotal(i) = True Then
                .Range(.Cells(fila, 1), .Cells(fila, NumCol)).Font.Bold = True
'                .Cells(Fila, 2) = "TOTALES"
                .Cells(fila, NumCol - 11) = grd.TextMatrix(i, COL_KILOSPROD)
                .Cells(fila, NumCol - 10) = grd.TextMatrix(i, COL_KILOSVEN)
                .Cells(fila, NumCol - 9) = grd.TextMatrix(i, COL_KILOSMEM)
                .Cells(fila, NumCol - 8) = grd.TextMatrix(i, COL_COSTOVARIABLE)
                .Cells(fila, NumCol - 5) = grd.TextMatrix(i, COL_UTITOTALNETA)
                .Cells(fila, NumCol - 4) = grd.TextMatrix(i, COL_MARGEN)
                .Cells(fila, NumCol - 3) = grd.TextMatrix(i, COL_PORCEM2)
                .Cells(fila, NumCol - 2) = grd.TextMatrix(i, COL_PRODM2)
                .Cells(fila, NumCol - 1) = grd.TextMatrix(i, COL_VENTAM2)
                .Cells(fila, NumCol - 1) = grd.TextMatrix(i, COL_MEM_M2)
                .Cells(fila, NumCol) = grd.TextMatrix(i, COL_VENTAUSD)
            Else
                j = 1
                mayor = 0
                For col = 1 To grd.Cols - 1
                            .Cells(fila, j) = grd.TextMatrix(i, col)
                        mayor = Len(grd.TextMatrix(i, col)) 'Para ajustar el ancho de columnas
                        If mayor > v(j - 1) Then            'de acuerdo a la celda más grande
                            .Columns(j).ColumnWidth = mayor '13/11/2000 ---> Angel P.
                            v(j - 1) = mayor
                        End If
                        j = j + 1
                Next col
            End If
            .Range(.Cells(fila, 1), .Cells(fila, NumCol)).Borders.LineStyle = 1
        Next i
    End With
     prg1.value = prg1.min
     MensajeStatus "Listo", vbDefault
End Sub

Private Sub Asignar()
    Dim valor As Currency, i As Integer
    Dim colu As Integer
    With grd
    Select Case .col
      Case COL_PVP
        valor = .ValueMatrix(.Row, COL_PVP)
        colu = COL_PVP
    Case COL_PESO
        valor = .ValueMatrix(.Row, COL_PESO)
        colu = COL_PESO
    Case COL_PRODM2
        valor = .ValueMatrix(.Row, COL_PRODM2)
        colu = COL_PRODM2
    Case COL_COSTOVARIABLE
        valor = .ValueMatrix(.Row, COL_COSTOVARIABLE)
        colu = COL_COSTOVARIABLE
    Case COL_PORCEM2
        valor = .ValueMatrix(.Row, COL_PORCEM2)
        colu = COL_PORCEM2
    End Select
     For i = .Row To .Rows - 9
            If .TextMatrix(i, COL_BODEGA) <> "." Then
             .TextMatrix(i, colu) = valor
            End If
        Next i
    End With
End Sub
'Guarda resultado
Private Sub GuardarResultado()
    Dim file As String, NumFile As Integer, Cadena As String
    Dim Filas As Long, Columnas As Long, i As Long, j As Long
    
    If grd.Rows = grd.FixedRows Then Exit Sub
    On Error GoTo errtrap
    
        With dlg1
          .CancelError = True
          '.Filter = "Texto (Separado por coma)|*.txt|Excel 97(XLS)|*.xls"
          .Filter = "Texto (Separado por coma)|*.csv"
          .ShowSave
          
          file = .filename
        End With
    
    If ExisteArchivo(file) Then
        If MsgBox("El nombre del archivo " & file & " ya existe desea sobreescribirlo?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    NumFile = FreeFile
    Open file For Output Access Write As #NumFile
    
    Cadena = ""
    For i = 1 To grd.Rows - 11
    If Not grd.IsSubtotal(i) Then
        For j = 1 To grd.Cols - 1
               Cadena = Cadena & grd.TextMatrix(i, j) & ","
        Next j
        Cadena = Mid(Cadena, 1, Len(Cadena) - 1)
        Print #NumFile, Cadena
        Cadena = ""
     End If
    Next i
    Close NumFile
    MsgBox "El archivo se ha exportado con éxito"
    Exit Sub
errtrap:
    If Err.Number <> 32755 Then
        MsgBox Err.Description
    End If
    Close NumFile
End Sub

Private Sub AbrirArchivo()
    Dim i As Long
    
    On Error GoTo errtrap
    With dlg1
        .CancelError = True
'        .Filter = "Texto (Separado por coma)|*.txt|Excel 97(XLS)|*.xls"
        .Filter = "Texto (Separado por coma *.csv)|*.csv|Texto (Separado por tabuladores *.txt)|*.txt|Todos *.*|*.*"
        .flags = cdlOFNFileMustExist
        If Len(.filename) = 0 Then          'Solo por primera vez, ubica a la carpeta de la aplicación
            .filename = App.Path & "\*.csv"
        End If
        
        .ShowOpen
        
        Select Case UCase$(Right$(dlg1.filename, 4))
        Case ".TXT", ".CSV"
            'ReformartearColumnas
            VisualizarTexto dlg1.filename
            'InsertarColumnas
        Case ".XLS"
       '     VisualizarExcel dlg1.FileName
        Case Else
        End Select
    End With
    Exit Sub
errtrap:
    If Err.Number <> 32755 Then DispErr
    Exit Sub
End Sub
Private Sub VisualizarTexto(ByVal archi As String)
    Dim f As Integer, s As String, i As Integer
    Dim Cadena
    On Error GoTo errtrap
    ReDim rec(0, 1)
    MensajeStatus "Está leyendo el archivo " & archi & " ...", vbHourglass
    grd.Rows = grd.FixedRows    'Limpia la grilla
    grd.Redraw = flexRDNone
    f = FreeFile                'Obtiene número disponible de archivo
    
    'Abre el archivo para lectura
    Open archi For Input As #f
        Do Until EOF(f)
            Line Input #f, s
            s = vbTab & Replace(s, ",", vbTab)      'Convierte ',' a TAB
           grd.AddItem s
        Loop
    Close #f
    RemueveSpace
    grd.Sort = flexSortUseColSort
    grd.Redraw = flexRDDirect
    AjustarAutoSize grd, -1, -1
    grd.ColWidth(grd.Cols - 1) = 4000
    grd.SetFocus
    MensajeStatus
    Exit Sub
errtrap:
    grd.Redraw = flexRDDirect
    MensajeStatus
    DispErr
    Close       'Cierra todo
    grd.SetFocus
    Exit Sub
End Sub



Private Sub RemueveSpace()
    Dim i As Long, j As Long
    With grd
        .Redraw = flexRDNone
        For i = .FixedRows To .Rows - 1
            For j = .FixedCols To .Cols - 1
                .TextMatrix(i, j) = Trim$(.TextMatrix(i, j))
            Next j
        Next i
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub FijarColor()
Dim j As Integer
Dim col As Integer
col = grd.Rows - 9
If grd.Rows = grd.FixedCols Then Exit Sub
    For j = 1 To grd.Rows - 10
         If grd.TextMatrix(j, 1) = "." Then
            grd.Cell(flexcpBackColor, j, 1, j, grd.Cols - 1) = &H8000000C  '&HC00000
            grd.Cell(flexcpForeColor, j, 1, j, grd.Cols - 1) = &HFFFF&
        End If
    Next
              grd.TextMatrix(grd.Rows - 4, COL_DESCGRUPO) = Format(GetSetting(APPNAME, App.Title, "costofijomensual", 0), "###,#0.00")
              grd.Cell(flexcpBackColor, grd.Rows - 4, COL_DESCGRUPO, grd.Rows - 4, COL_DESCGRUPO) = &H80000005  'color blanco
              grd.Cell(flexcpForeColor, grd.Rows - 4, COL_DESCGRUPO, grd.Rows - 4, COL_DESCGRUPO) = &H80000012 'color negro
     
End Sub

Private Sub CargarBodegas()
Dim sql As String, s As String
Dim rs As Recordset
sql = "Select codbodega,descripcion from ivbodega"
Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    With grdCos
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If Not rs.EOF Then .LoadArray MiGetRows(rs)
            s = "<Bodega|<Descripción|>Metros2|>Seleccionar"
            .FormatString = s
            .Redraw = flexRDBuffered
    End With
Set rs = Nothing
End Sub

Private Sub Configurar()
    Static coditem As String, CodAlt As String, _
           Desc As String, _
           codg As String, Numg As Integer, bandIVA As Boolean, bandFraccion As Boolean
    Dim sql As String, cond As String, rs As Recordset, comodin As String
    Dim desde As String, hasta   As String
    Dim valor As Currency, s As String
    Dim costofijoM As Currency, NumReg As Long
    
    Dim i As Integer
    Dim cadbod As String
    Dim bandSi As Boolean, j As Integer
    On Error GoTo errtrap
    Dim strBodegas As String
    Dim strTransVen As String
    Dim strTransProd As String
    Dim strTransVenMEM As String
    Dim strTransNotaCredito As String
    
    
    #If DAOLIB Then
        comodin = "*"
    #Else
        comodin = "%"
    #End If
'    comodin = "%"
    'Abre la pantalla de búsqueda
    If Not frmIVBusquedaPNew.Inicio( _
                desde, _
                hasta, _
                costoFijoMensual, v) Then
        'Si fue cancelada la busqueda, sale no mas
        grd.SetFocus
        Exit Sub
    End If
    MensajeStatus MSG_PREPARA, vbHourglass
    strBodegas = GetSetting(APPNAME, App.Title, "strBodegas", "")
    strTransVen = GetSetting(APPNAME, App.Title, "TransVentas", "")
    strTransProd = GetSetting(APPNAME, App.Title, "TransProd", "")
    strTransVenMEM = GetSetting(APPNAME, App.Title, "TransMEM", "")
    strTransNotaCredito = GetSetting(APPNAME, App.Title, "TransNotaCredito", "")
    'Precio de Venta
     VerificaExistenciaTablaTemp "tmp1"
     sql = "SELECT IVB.CODBODEGa,IVB.DESCRIPCION AS DESCBODEGA, "
     sql = sql & " IV4.CODGRUPO4,IV4.DESCRIPCION, "
     sql = sql & " case when len(max(CodAlterno2))=0 then 1 else  cast(max(CodAlterno2)as money )end as fact,"
     sql = sql & " (SUM(PrecioRealTotal)/SUM(IVK.CANTIDAD)) * (case when len(max(CodAlterno2))=0 then 1 else  cast(max(CodAlterno2)as money )end) as pvp,"
     sql = sql & " IVI.IDGRUPO4,"
     sql = sql & " (SUM(IVK.CANTIDAD)) *-1 / (case when len(max(CodAlterno2))=0 then 1 else  cast(max(CodAlterno2)as money )end) as CantTotal,"
     sql = sql & "SUM(PrecioRealTotal)*-1 as precioTotal "
     sql = sql & " INTO TMP1 "
     sql = sql & " FROM GNCOMPROBANTE GNC "
     sql = sql & " INNER JOIN IVKARDEX IVK INNER JOIN IVBODEGA IVB ON IVK.IDBODEGA=IVB.IDBODEGA "
     sql = sql & " INNER JOIN IVINVENTARIO IVI "
     sql = sql & " INNER JOIN IVGRUPO4 IV4 "
     sql = sql & " ON IVI.IDGRUPO4=IV4.IDGRUPO4 "
     sql = sql & " ON IVK.IDINVENTARIO=IVI.IDINVENTARIO "
     sql = sql & " ON GNC.TRANSID=IVK.TRANSID "
     sql = sql & " WHERE CODTRANS    IN (" & PreparaCadena(strTransVen) & ")"
     sql = sql & " AND GNC.FECHATRANS BETWEEN '" & desde & "' AND '" & hasta & " '"
     sql = sql & " AND IVI.IDGRUPO1=2"  ' SOLO PRODUCTO TERMINADO
     sql = sql & " AND GNC.ESTADO <> 3"
     If Len(strBodegas) > 0 Then
         sql = sql & " and ivb.codbodega   IN (" & PreparaCadena(strBodegas) & ") "
     End If
     sql = sql & " GROUP BY IVB.CODBODEGa,IVB.DESCRIPCION,IVI.IDGRUPO4, IV4.DESCRIPCION,IV4.CODGRUPO4 "
     sql = sql & "  Having Sum(ivk.CANTIDAD) <> 0"
     gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
    
     VerificaExistenciaTablaTemp "tmp2"
    
     sql = "SELECT IVB.CODBODEGa,IVB.DESCRIPCION AS DESCBODEGA, IV4.CODGRUPO4,IV4.DESCRIPCION, IVI.IDGRUPO4, "
     sql = sql & " case when len(max(CodAlterno2))=0 then 1 else  cast(max(CodAlterno2)as money )end as fact,"
     sql = sql & " (SUM(PrecioRealTotal)/SUM(IVK.CANTIDAD)) * (case when len(max(CodAlterno2))=0 then 1 else  cast(max(CodAlterno2)as money )end) as pvp,"
     sql = sql & " Sum (ivk.CANTIDAD)as CantTotal "
     sql = sql & " INTO TMP2 "
     sql = sql & " FROM GNCOMPROBANTE GNC "
     sql = sql & " INNER JOIN IVKARDEX IVK INNER JOIN IVBODEGA IVB ON IVK.IDBODEGA=IVB.IDBODEGA "
     sql = sql & " INNER JOIN IVINVENTARIO IVI "
     sql = sql & " INNER JOIN IVGRUPO4 IV4 "
     sql = sql & " ON IVI.IDGRUPO4=IV4.IDGRUPO4 "
     sql = sql & " ON IVK.IDINVENTARIO=IVI.IDINVENTARIO "
     sql = sql & " ON GNC.TRANSID=IVK.TRANSID "
     sql = sql & " WHERE CODTRANS  IN (" & PreparaCadena(strTransProd) & ")"
     sql = sql & " AND GNC.FECHATRANS BETWEEN '" & desde & "' AND '" & hasta & " '"
     sql = sql & " AND IVI.IDGRUPO1=2"
     sql = sql & " AND GNC.ESTADO <> 3"
     If Len(strBodegas) > 0 Then
             sql = sql & " and ivb.codbodega   IN (" & PreparaCadena(strBodegas) & ") "
     End If
     sql = sql & " GROUP BY IVB.CODBODEGa,IVB.DESCRIPCION,IVI.IDGRUPO4, IV4.DESCRIPCION,IV4.CODGRUPO4 "
      
    gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
     
    'cantidad MEMEOS
    VerificaExistenciaTablaTemp "tmp3"
     sql = "SELECT IVB.CODBODEGa,IVB.DESCRIPCION AS DESCBODEGA, "
     sql = sql & " IV4.CODGRUPO4,IV4.DESCRIPCION, "
    sql = sql & " IVI.IDGRUPO4,"
     sql = sql & " 1  as fact,"
     sql = sql & " (SUM(IVK.CANTIDAD)) *-1 as CantTotal, "
     sql = sql & "SUM(PrecioRealTotal)*-1 as precioTotal "
     sql = sql & " INTO TMP3 "
     sql = sql & " FROM GNCOMPROBANTE GNC inner join gncentrocosto gncc on gnc.idcentro=gncc.idcentro "
     sql = sql & " INNER JOIN IVKARDEX IVK INNER JOIN IVBODEGA IVB ON IVK.IDBODEGA=IVB.IDBODEGA "
     sql = sql & " INNER JOIN IVINVENTARIO IVI "
     sql = sql & " INNER JOIN IVGRUPO4 IV4 "
     sql = sql & " ON IVI.IDGRUPO4=IV4.IDGRUPO4 "
     sql = sql & " ON IVK.IDINVENTARIO=IVI.IDINVENTARIO "
     sql = sql & " ON GNC.TRANSID=IVK.TRANSID "
     sql = sql & " WHERE CODTRANS    IN (" & PreparaCadena(strTransVenMEM) & ")"
     sql = sql & " AND GNC.FECHATRANS BETWEEN '" & desde & "' AND '" & hasta & " '"
     sql = sql & " AND IVI.IDGRUPO1=2"  ' SOLO PRODUCTO TERMINADO
     sql = sql & " AND GNC.ESTADO <> 3"
     sql = sql & " AND GNCC.CODCENTRO NOT LIKE '%TRA%'"
     If Len(strBodegas) > 0 Then
         sql = sql & " and ivb.codbodega   IN (" & PreparaCadena(strBodegas) & ") "
     End If
     sql = sql & " GROUP BY IVB.CODBODEGa,IVB.DESCRIPCION,IVI.IDGRUPO4, IV4.DESCRIPCION,IV4.CODGRUPO4 "
     sql = sql & "  Having Sum(ivk.CANTIDAD) <> 0"
     gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
    
    
    'Precio de Venta
     VerificaExistenciaTablaTemp "tmp4"
sql = "SELECT IVB.CODBODEGa,IVB.DESCRIPCION AS DESCBODEGA, "
     sql = sql & " IV4.CODGRUPO4,IV4.DESCRIPCION, "
     sql = sql & " case when len(max(CodAlterno2))=0 then 1 else  cast(max(CodAlterno2)as money )end as fact,"
     sql = sql & " (SUM(PrecioRealTotal)/SUM(IVK.CANTIDAD)) * (case when len(max(CodAlterno2))=0 then 1 else  cast(max(CodAlterno2)as money )end) as pvp,"
     sql = sql & " IVI.IDGRUPO4,"
     sql = sql & " (SUM(IVK.CANTIDAD)) *-1 / (case when len(max(CodAlterno2))=0 then 1 else  cast(max(CodAlterno2)as money )end) as CantTotal,"
     sql = sql & "SUM(PrecioRealTotal)*-1 as precioTotal "
     sql = sql & " INTO TMP4 "
     sql = sql & " FROM GNCOMPROBANTE GNC "
     sql = sql & " INNER JOIN IVKARDEX IVK INNER JOIN IVBODEGA IVB ON IVK.IDBODEGA=IVB.IDBODEGA "
     sql = sql & " INNER JOIN IVINVENTARIO IVI "
     sql = sql & " INNER JOIN IVGRUPO4 IV4 "
     sql = sql & " ON IVI.IDGRUPO4=IV4.IDGRUPO4 "
     sql = sql & " ON IVK.IDINVENTARIO=IVI.IDINVENTARIO "
     sql = sql & " ON GNC.TRANSID=IVK.TRANSID "
     sql = sql & " WHERE CODTRANS    IN (" & PreparaCadena(strTransNotaCredito) & ")"
     sql = sql & " AND GNC.FECHATRANS BETWEEN '" & desde & "' AND '" & hasta & " '"
     sql = sql & " AND IVI.IDGRUPO1=2"  ' SOLO PRODUCTO TERMINADO
     sql = sql & " AND GNC.ESTADO <> 3"
     If Len(strBodegas) > 0 Then
         sql = sql & " and ivb.codbodega   IN (" & PreparaCadena(strBodegas) & ") "
     End If
     sql = sql & " GROUP BY IVB.CODBODEGa,IVB.DESCRIPCION,IVI.IDGRUPO4, IV4.DESCRIPCION,IV4.CODGRUPO4 "
     sql = sql & "  Having Sum(ivk.CANTIDAD) <> 0"
     gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
    
    
    
     
     
     sql = "SELECT isnull(A.CODBODEGa, ISNULL(b.codbodega,isnull(C.codbodega,d.codbodega))), "
     sql = sql & " isnull(A.DESCBODEGA,ISNULL(b.descbodega,isnull(C.descbodega,d.descbodega))), "
     sql = sql & " isnull(A.DESCRIPCION, ISNULL(b.descripcion,C.descripcion)), "
     sql = sql & " isnull(A.PVP,0),0,0,0,0,0,0,0,0,0,0,0,0,  "
     sql = sql & " B.CantTotal, a.CantTotal,  D.CantTotal , C.CantTotal ,"
     sql = sql & " A.precioTotal, D.precioTotal, C.precioTotal, "
     sql = sql & " isnull(a.fact,ISNULL(b.fact,isnull(D.fact, C.fact)))  "
     sql = sql & " FROM TMP1 A"
     sql = sql & " FULL JOIN TMP2 B "
     sql = sql & " ON A.IDGRUPO4=B.IDGRUPO4 "
     sql = sql & " AND A.CODBODEGa=B.CODBODEGa"
     
     sql = sql & " FULL JOIN TMP3 C "
     sql = sql & " ON (A.IDGRUPO4=C.IDGRUPO4 OR B.IDGRUPO4=C.IDGRUPO4) "
     sql = sql & " AND (A.CODBODEGa=C.CODBODEGa OR B.CODBODEGa=C.CODBODEGa)"
     
     sql = sql & " FULL JOIN TMP4 D "
     sql = sql & " ON (A.IDGRUPO4=D.IDGRUPO4 OR B.IDGRUPO4=D.IDGRUPO4 OR C.IDGRUPO4=D.IDGRUPO4 ) "
     sql = sql & " AND (A.CODBODEGa=D.CODBODEGa OR B.CODBODEGa=D.CODBODEGa OR C.CODBODEGa=D.CODBODEGa)"
     

     sql = sql & " ORDER BY "
     sql = sql & " isnull(A.CODBODEGa,ISNULL(b.codbodega,ISNULL(C.codbodega,D.codbodega))), "
     sql = sql & " isnull(A.DESCBODEGA,ISNULL(b.descbodega,ISNULL(C.descbodega,d.descbodega)))  "
     sql = sql & " , isnull(A.DESCRIPCION,ISNULL(b.descripcion,ISNULL(C.descripcion, d.descripcion))) "
     Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
   
    With grd
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If Not rs.EOF Then .LoadArray MiGetRows(rs)
        For i = 1 To 8 'Agrega filas adicionales
            grd.AddItem vbTab '& "."
        Next
        
''''         s = "^#|<bodega|<Codigo|<Descripción|>P.V.|>Peso|>Total Kilos Producidos|>Total Kilos Vendidos|>Total Kilos MEM"
''''              s = s & "|>Costo Fijo Unitario|>Costo Variable|>Costo Total|>Utilidad Neta Unitaria"
''''              s = s & "|>Utilidad Total NETA|>Margen|>Porcentaje M2|>Produccion M2|>Ventas M2|>Ventas MEM|>Total Ventas $|>Total Ventas MEM $"
              
               s = "^#|<bodega|<Codigo|<Descripción|>P.V.|>Peso"
               s = s & "|>Total Kilos Producidos|>Total Kilos Vendidos|>Total Kilos Devol|>Total Kilos MEM"
               s = s & "|>Costo Fijo Unitario|>Costo Variable|>Costo Total|>Utilidad Neta Unitaria"
               s = s & "|>Utilidad Total NETA|>Margen|>Porcentaje M2|>Produccion M2"
               s = s & "|>Ventas M2|>Devol M2|>MEM M2"
               s = s & "|>Total Ventas $|>Total Devol  $|>Total MEM $|>FactConv"
  
        .FormatString = s
              
              
                .FormatString = s
        .Redraw = flexRDBuffered
        .SetFocus
    End With
    CargarLista
    
   ConfigCols

    AgregarDatosPlantilla
            FijarColor
            costofijoM = GetSetting(APPNAME, App.Title, "costofijomensual", 0)

    MensajeStatus
    Exit Sub
errtrap:
    grd.Redraw = flexRDBuffered
    MensajeStatus
    DispErr
    grd.SetFocus
    Exit Sub
End Sub
Private Sub AgregarDatosPlantilla()
 Dim f As Integer, s As String, i As Integer
    Dim Cadena, s1 As String, s2 As String
    On Error GoTo errtrap
    Dim archi As String
    ReDim rec(0, 1)
    Dim cad As String, bod As String
    archi = GetSetting(APPNAME, App.Title, "Ruta Plantilla", "")
    f = FreeFile                'Obtiene número disponible de archivo
    
    'Abre el archivo para lectura
    Open archi For Input As #f
        Do Until EOF(f)
            Line Input #f, s
            s = vbTab & Replace(s, ",", vbTab)      'Convierte ',' a TAB
            cad = vbTab & Replace(s, ",", vbTab)      'Convierte ',' a TAB
            Select Case Me.tag
            Case "Costos", "Produccion"
                Cadena = Split(s, vbTab, -1, 1)
                s2 = ""
                For i = 1 To UBound(Cadena)
                    Select Case i
                        Case 1
                                  bod = Cadena(i)
                        Case 3
                                s1 = Cadena(i)
                                copiaValor s1, bod, cad
                        Case Else
                                s = s & vbTab & Cadena(i)
                    End Select
                Next i
            End Select
        Loop
    Close #f
    MensajeStatus
    Exit Sub
errtrap:
    grd.Redraw = flexRDDirect
    MensajeStatus
    DispErr
    Close       'Cierra todo
    grd.SetFocus
Exit Sub
End Sub
Private Sub copiaValor(s As String, Bodega As String, Cadena As String)
Dim i As Long
Dim cade, bod As String
Dim peso As String, costoVariable As String, porcentaje As String, pronostico As String
Dim costofijoMen As String
cade = Split(Cadena, vbTab, -1, 1)
For i = 1 To UBound(cade)
    Select Case i
        Case 6
              peso = cade(i)
        Case 10
             costoVariable = cade(i)
        Case 15 'toma porcentaje M2
             porcentaje = cade(i)
        Case 16
             pronostico = cade(i)
        Case 17
             costofijoMen = cade(i)
    End Select
Next
   For i = 1 To grd.Rows - 10 'antes 9
    If Not grd.IsSubtotal(i) Then
        If s = grd.TextMatrix(i, COL_DESCGRUPO) And Bodega = grd.TextMatrix(i, COL_BODEGA) Then
                grd.TextMatrix(i, COL_PESO) = peso
                grd.TextMatrix(i, COL_COSTOVARIABLE) = costoVariable
                grd.TextMatrix(i, COL_PORCEM2) = porcentaje
        End If
    End If
   Next
End Sub
Private Sub CargarConfiguarciones()
Dim i As Integer
Dim valor As Currency
    costoFijoMensual = GetSetting(APPNAME, App.Title, "CostoFijoMensual", 0)
   With grdCos
        For i = 1 To .Rows - 1
             grdCos.TextMatrix(i, 2) = GetSetting(APPNAME, App.Title, "M2" & grdCos.TextMatrix(i, 0), 0)
             grdCos.TextMatrix(i, 3) = GetSetting(APPNAME, App.Title, "band" & grdCos.TextMatrix(i, 0), "")
        Next
        For i = 1 To .Rows - 1
         If Val(.TextMatrix(i, 3)) = -1 Then
           ReDim Preserve v(2, i)
           v(0, i) = .TextMatrix(i, 0) 'CODIGO BODEGA
           v(1, i) = .TextMatrix(i, 2)  'VALOR
           v(2, i) = .TextMatrix(i, 1) ' DESCRIPCION BODEGA
         End If
        Next
        
        
   End With
End Sub
Private Function TomarValor(codbod As String) As Currency
Dim i As Integer
  For i = 1 To grdCos.Rows - 10
      If grdCos.TextMatrix(i, 0) = codbod And Val(grdCos.TextMatrix(i, 3)) = -1 Then
           TomarValor = grdCos.ValueMatrix(i, 2)
           Exit Function
      End If
  Next
End Function
Private Function NoEstaCargado(codbod As String) As Boolean
Dim band As Boolean
Dim i As Integer
For i = 1 To grd.Rows - 10
   If grd.TextMatrix(i, COL_BODEGA) <> "." Then
       If grd.TextMatrix(i, COL_BODEGA) = codbod Then
            If Not Existe(codbod) Then
              NoEstaCargado = True
              Exit Function
            End If
       End If
   End If
Next
'NoEstaCargado = band
End Function
Private Function Existe(codbod As String) As Boolean
   Dim i As Integer
     For i = 1 To grd.Rows - 10
       If grd.TextMatrix(i, COL_BODEGA) = codbod Then
          If Len(grd.TextMatrix(i, grd.ColIndex("Metros2"))) > 0 Then
              Existe = True
              Exit Function
          End If
       End If
     Next
End Function
Private Function MetrosCuadrados(codbod As String) As Currency
Dim i As Integer
      For i = 1 To grd.Rows - 10
          If grd.TextMatrix(i, COL_BODEGA) = codbod Then
          End If
      Next
End Function

Private Sub CargarLista()
Dim i As Integer
Dim valor As Currency
   costoFijoMensual = GetSetting(APPNAME, App.Title, "costofijomensual", 0)
   
End Sub

Private Sub GuardaConfM2()
Dim i As Integer

   With grd
      SaveSetting APPNAME, App.Title, "CostoFijoMensual", grd.ValueMatrix(grd.Rows - 4, COL_DESCGRUPO)
   End With
      
End Sub
