VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFormulario 
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
      Left            =   300
      TabIndex        =   1
      Top             =   540
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
      BackColor       =   -2147483624
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483624
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
      Rows            =   68
      Cols            =   10
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
            Picture         =   "frmFormulario.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario.frx":0114
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario.frx":0568
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario.frx":067C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario.frx":0790
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario.frx":0E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario.frx":1B20
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario.frx":1F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario.frx":2084
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario.frx":3906
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   3
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
         TabIndex        =   2
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
Attribute VB_Name = "frmFormulario"
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
Private Const ColorFondo = &H8000000A
Private Const ColorFondo1 = &H80000013
Private Busqueda As Boolean
Private mobjBusq As Busqueda
Private mCodMoneda  As String
Private objcond As Condicion

Public Sub Inicio(ByVal tag As String)
    Dim rutaPlantilla
    Dim i As Integer
    Dim valor As Currency
    On Error GoTo ErrTrap
    
    Me.tag = tag            'Guarda en la propiedad Tag para distinguir después
    Me.Show
    Me.ZOrder
    Select Case Me.tag
    Case "F104"
        Me.Caption = "Declaración del Impuesto Agrgado "
    Case "Produccion"
         Me.Caption = "Punto de Equilibrio basado en Ventas "
    End Select
    'LlenaFormatoFormulario
    'Inicializa la grilla
    'rutaPlantilla = GetSetting(APPNAME, App.Title, "Ruta Plantilla", "")
    grd.Rows = grd.FixedRows
    ConfigCols
'    VisualizarTexto (rutaPlantilla)
'    For i = 1 To 8 'Agrega filas adicionales
'            grd.AddItem vbTab '& "."
'        Next
'   ConfigCols
'    FijarColor ': CargarConfiguarciones
'    CargarBodegas
'    CargarConfiguarciones
    LlenaFormatoFormulario
    LLENADATOS
    CambiaFondoCeldasEditables
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Initialize()
    Set mobjBusq = New Busqueda
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF3
        KeyCode = 0
    Case vbKeyF5
'          Calcular
'          GuardaConfM2
          BuscarComprasNetas12y0
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
    CalcularPorcentajes
    
    If grd.CellBackColor = grd.BackColorFrozen Or grd.CellBackColor = &HC00000 Then
       Cancel = True
    End If
End Sub

Private Sub grd_BeforeSort(ByVal col As Long, Order As Integer)
    'Impide mientras está procesando
    If mProcesando Then Order = flexSortNone
End Sub

Private Sub grd_ChangeEdit()
    'CalcularPorcentajes
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
        'Buscar
        BuscarComprasNetas12y0
    Case "Abrir"
        'No hace nada
    Case "Asignar": Asignar
'    Case "Calcular":
'          Calcular
'          GuardaConfM2
    Case "Excel": ExportaExcel ("Calculo de Costos")
    Case "Guardar": GuardarResultado
    Case "Cerrar":      Cerrar
    End Select
End Sub
Private Sub ConfigCols(Optional subt As Integer)
    Dim s As String, i As Long, j As Integer, s1 As String
    Dim fmt As String
    With grd
               s = "^#|<c1|<c2|^c3|<c4|<c5|>c6|^c7|>c8|^c9"
        .FormatString = s
        'AjustarAutoSize grd, -1, -1, 4000
        AsignarTituloAColKey grd
        .ColWidth(0) = 350
        .ColWidth(1) = 500
        .ColWidth(2) = 3000
        .ColWidth(3) = 500
        .ColWidth(4) = 3000
        .ColWidth(5) = 500
        .ColWidth(6) = 3000
        .ColWidth(7) = 500
        .ColWidth(8) = 3000
        .ColWidth(9) = 500
        'grilla de resultados
        'Columnas modificables (Longitud maxima)
            .ColFormat(.ColIndex("c4")) = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda(fmt)
            .ColFormat(.ColIndex("c6")) = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda(fmt)
            .ColFormat(.ColIndex("c8")) = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda(fmt)
        .MergeCol(3) = True
        .Refresh
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
Dim TotalKilosVen As Currency
Dim ValorM2 As Currency
'encera valores iniciales
TotalKilosPro = 0
ventasTotal = 0
UtilidadTotalNeta = 0
MargenGlobal = 0
totalPronostico = 0
costoVariableTotal = 0
TotalKilosVen = 0
costoFijoMensual = grd.ValueMatrix(grd.Rows - 3, grd.ColIndex("Descripción"))
         grd.subtotal flexSTSum, 1, 6, , grd.BackColorFixed, , True, " Subtotal", 1, False
                 grd.subtotal flexSTSum, 1, 7, , grd.BackColorFixed, , True, " Subtotal", 1, False
                 grd.subtotal flexSTSum, 1, 9, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, 13, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, 14, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, 15, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, 16, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, 1, 17, , grd.BackColorFixed, , True, " Subtotal", 1, True
                 grd.subtotal flexSTSum, -1, 6, , grd.BackColorFixed, , True, " Total", 1, False
        
With grd
   For j = 1 To grd.Rows - 10
      If Not .IsSubtotal(j) Then
        If Val(.TextMatrix(j, .ColIndex("Produccion M2"))) <> 0 Then
             .TextMatrix(j, .ColIndex("Total Kilos Producidos")) = .ValueMatrix(j, .ColIndex("Peso")) * .ValueMatrix(j, .ColIndex("Produccion M2"))
        Else
            .TextMatrix(j, .ColIndex("Total Kilos Producidos")) = 0
        End If
         If Len(.TextMatrix(j, .ColIndex("Ventas M2"))) > 0 Then
            .TextMatrix(j, .ColIndex("Total Kilos Vendidos")) = .TextMatrix(j, .ColIndex("Ventas M2")) * .ValueMatrix(j, .ColIndex("Peso"))
        Else
            .TextMatrix(j, .ColIndex("Total Kilos Vendidos")) = 0
        End If
            TotalKilosPro = TotalKilosPro + .TextMatrix(j, .ColIndex("Total Kilos Producidos"))
            TotalKilosVen = TotalKilosVen + .TextMatrix(j, .ColIndex("Total Kilos Vendidos"))

     End If
   Next
        Select Case Me.tag
         Case "Costos"
            If TotalKilosPro <> 0 Then
                .TextMatrix(.Rows - 10, .ColIndex("Descripción")) = .ValueMatrix(.Rows - 4, .ColIndex("Descripción")) / TotalKilosPro
            Else
                .TextMatrix(.Rows - 10, .ColIndex("Descripción")) = 0
            End If
         
         Case "Produccion"
            If TotalKilosVen <> 0 Then
                .TextMatrix(.Rows - 10, .ColIndex("Descripción")) = .ValueMatrix(.Rows - 4, .ColIndex("Descripción")) / TotalKilosVen
            Else
                .TextMatrix(.Rows - 10, .ColIndex("Descripción")) = 0
            End If
        End Select
   For j = 1 To grd.Rows - 10
   If Not .IsSubtotal(j) Then
      If Val(.TextMatrix(j, .ColIndex("Peso"))) > 0 Then
        .TextMatrix(j, .ColIndex("Costo Fijo Unitario")) = .ValueMatrix(j, .ColIndex("Peso")) * .ValueMatrix(.Rows - 10, .ColIndex("Descripción"))
      End If
      If Val(.TextMatrix(j, .ColIndex("Costo Variable"))) <> 0 And Val(.TextMatrix(j, .ColIndex("Costo Fijo Unitario"))) <> 0 Then
        .TextMatrix(j, .ColIndex("Costo Total")) = .ValueMatrix(j, .ColIndex("Costo Variable")) + .ValueMatrix(j, .ColIndex("Costo Fijo Unitario"))
      End If
      If Val(.TextMatrix(j, .ColIndex("P.V."))) <> 0 And Val(.TextMatrix(j, .ColIndex("Costo Total"))) <> 0 Then
        .TextMatrix(j, .ColIndex("Utilidad Neta Unitaria")) = .TextMatrix(j, .ColIndex("P.V.")) - .TextMatrix(j, .ColIndex("Costo Total"))
      End If
      
    Select Case Me.tag
         Case "Costos"
            If Val(.TextMatrix(j, .ColIndex("Utilidad Neta Unitaria"))) <> 0 And Val(.TextMatrix(j, .ColIndex("Produccion M2"))) <> 0 Then
              .TextMatrix(j, .ColIndex("Utilidad Total NETA")) = .TextMatrix(j, .ColIndex("Utilidad Neta Unitaria")) * .TextMatrix(j, .ColIndex("Produccion M2"))
            End If
        Case "Produccion"
                If Val(.TextMatrix(j, .ColIndex("Utilidad Neta Unitaria"))) <> 0 And Val(.TextMatrix(j, .ColIndex("Ventas M2"))) <> 0 Then
              .TextMatrix(j, .ColIndex("Utilidad Total NETA")) = .TextMatrix(j, .ColIndex("Utilidad Neta Unitaria")) * .TextMatrix(j, .ColIndex("Ventas M2"))
            End If
        End Select
      If Val(.TextMatrix(j, .ColIndex("Utilidad Neta Unitaria"))) <> 0 And Val(.TextMatrix(j, .ColIndex("P.V."))) <> 0 Then
          .TextMatrix(j, .ColIndex("Margen")) = .TextMatrix(j, .ColIndex("Utilidad Neta Unitaria")) / .TextMatrix(j, .ColIndex("P.V."))
      End If
      
        
      'CALCULA TOTAL VENTAS
      ventasTotal = ventasTotal + .ValueMatrix(j, .ColIndex("Total Ventas"))
      'calculta suma utilidad total neta
      UtilidadTotalNeta = UtilidadTotalNeta + .ValueMatrix(j, .ColIndex("Utilidad Total NETA"))
      'calcula totalpronostico
     totalPronostico = totalPronostico + .ValueMatrix(j, .ColIndex("Produccion M2"))
     'Calcula costovariabletotal
     Select Case Me.tag
         Case "Costos"
            costoVariableTotal = costoVariableTotal + (.ValueMatrix(j, .ColIndex("Produccion M2")) * .ValueMatrix(j, .ColIndex("Costo Variable")))
        Case "Produccion"
            costoVariableTotal = costoVariableTotal + (.ValueMatrix(j, .ColIndex("Ventas M2")) * .ValueMatrix(j, .ColIndex("Costo Variable")))
    End Select
   End If
   Next
      'asigna ventastotal
      If ventasTotal <> 0 Then
        MargenGlobal = UtilidadTotalNeta / ventasTotal ' .ValueMatrix(.Rows - 1, .ColIndex("Total Ventas"))
      End If
      .TextMatrix(grd.Rows - 9, .ColIndex("Descripción")) = Round(MargenGlobal * 100, 2) & " %"
      .TextMatrix(grd.Rows - 8, .ColIndex("Descripción")) = costoVariableTotal
      If ventasTotal <> 0 Then
          .TextMatrix(grd.Rows - 4, .ColIndex("P.V.")) = Round((.ValueMatrix(grd.Rows - 4, .ColIndex("Descripción")) / ventasTotal) * 100, 2) & " %"
          .TextMatrix(grd.Rows - 8, .ColIndex("P.V.")) = Round((costoVariableTotal / ventasTotal) * 100, 2) & " %"
      End If
      'UTILIDADNETA
      .TextMatrix(grd.Rows - 7, .ColIndex("Descripción")) = ventasTotal - .TextMatrix(grd.Rows - 4, .ColIndex("Descripción")) - costoVariableTotal
      ' Porcentaje M2 UTILIDADNETA
      If ventasTotal <> 0 Then
        .TextMatrix(grd.Rows - 7, .ColIndex("P.V.")) = Round((.ValueMatrix(grd.Rows - 7, .ColIndex("Descripción")) / ventasTotal) * 100, 2) & " %"
      End If
      'anual
      .TextMatrix(grd.Rows - 6, .ColIndex("Descripción")) = .ValueMatrix(grd.Rows - 7, .ColIndex("Descripción")) * 12
End With

grd.subtotal flexSTSum, 1, 6, , grd.BackColorFixed, , True, " Subtotal", 1, False
grd.subtotal flexSTSum, 1, 7, , grd.BackColorFixed, , True, " Subtotal", 1, True
grd.subtotal flexSTSum, 1, 12, , grd.BackColorFixed, , True, " Subtotal", 1, True
grd.subtotal flexSTSum, 1, 13, , grd.BackColorFixed, , True, " Subtotal", 1, True
grd.subtotal flexSTSum, -1, 6, , grd.BackColorFixed, , True, " Total", 1, False
grd.subtotal flexSTSum, -1, 7, , grd.BackColorFixed, , True, " Total", 1, False
grd.subtotal flexSTSum, -1, 9, , grd.BackColorFixed, , True, " Total", 1, False
grd.subtotal flexSTSum, -1, 12, , grd.BackColorFixed, , True, " Total", 1, False
grd.subtotal flexSTSum, -1, 13, , grd.BackColorFixed, , True, " Total", 1, False
grd.subtotal flexSTSum, -1, 15, , grd.BackColorFixed, , True, " Total", 1, False
grd.subtotal flexSTSum, -1, 16, , grd.BackColorFixed, , True, " Total", 1, False
grd.subtotal flexSTSum, -1, 17, , grd.BackColorFixed, , True, " Total", 1, False
For i = 1 To grd.Cols - 1
    grd.TextMatrix(grd.Rows - 2, i) = 0 'Encera la ultima fila
    grd.TextMatrix(grd.Rows - 2, i) = "" 'Encera la ultima fila
Next
grd.Select grd.Rows - 9, grd.ColIndex("Descripción"), grd.Rows - 9, grd.ColIndex("Descripción")
grd.CellAlignment = flexAlignRightBottom
grd.Select grd.Rows - 8, grd.ColIndex("Descripción"), grd.Rows - 8, grd.ColIndex("Descripción")
grd.CellAlignment = flexAlignRightBottom
grd.Select grd.Rows - 7, grd.ColIndex("Descripción"), grd.Rows - 7, grd.ColIndex("Descripción")
grd.CellAlignment = flexAlignRightBottom
grd.Select grd.Rows - 6, grd.ColIndex("Descripción"), grd.Rows - 6, grd.ColIndex("Descripción")
grd.CellAlignment = flexAlignRightBottom
grd.Select grd.Rows - 5, grd.ColIndex("Descripción"), grd.Rows - 5, grd.ColIndex("Descripción")
grd.CellAlignment = flexAlignRightBottom
grd.Select grd.Rows - 4, grd.ColIndex("Descripción"), grd.Rows - 4, grd.ColIndex("Descripción")
grd.CellAlignment = flexAlignRightBottom
grd.Select grd.Rows - 3, grd.ColIndex("Descripción"), grd.Rows - 3, grd.ColIndex("Descripción")
grd.CellAlignment = flexAlignRightBottom
grd.Select grd.Rows - 2, grd.ColIndex("Descripción"), grd.Rows - 2, grd.ColIndex("Descripción")
grd.CellAlignment = flexAlignRightBottom
grd.Select grd.Rows - 10, grd.ColIndex("Descripción"), grd.Rows - 10, grd.ColIndex("Descripción")
grd.CellAlignment = flexAlignRightBottom
'grd.Select grd.Rows - 8, grd.ColIndex("P.V."), grd.Rows - 8, grd.ColIndex("P.V.")
'grd.CellAlignment = flexAlignRightBottom
'grd.Cell(flexcpTextDisplay.ColFormat( .CellFloodPercent = grd.ValueMatrix(grd.Rows - 8, grd.ColIndex("P.V."))


'grd.Cell(flexcpTextDisplay, grd.Rows - 9, grd.ColIndex("P.V.")) = "#.##"



End Sub
Private Sub EliminaFila()
    Dim msg As String, r As Long
    On Error GoTo ErrTrap
    r = grd.Row
    If (grd.Rows > grd.FixedRows) Then
        If (Not grd.IsSubtotal(r)) Then
            msg = "Desea eliminar la fila #" & r & "?"
            If MsgBox(msg, vbYesNo) <> vbYes Then
                grd.SetFocus
                Exit Sub
            End If
            If grd.TextMatrix(r, grd.ColIndex("bodega")) = "." Then: Exit Sub
            'Elimina del grid
            grd.RemoveItem r
        End If
    End If
    grd.SetFocus
    Exit Sub
ErrTrap:
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
        Dim Fila As Long, col As Long, i As Long, j As Long
    Dim v() As Long, mayor As Long
    Dim NumCol As Integer
    Dim fmt As String
    prg1.min = 0
    prg1.max = grd.Rows - 1
    MensajeStatus "Está Exportando  a Excel ...", vbHourglass
    With ws
        Fila = 2
        .Range("H1").Font.Name = "Arial"
        .Range("H1").Font.Size = 12
        .Range("H1").Font.Bold = True
        .Cells(Fila, 1) = titulo
        
        .PageSetup.PaperSize = xlPaperLetter 'Tamaño del papel (carta)
        .PageSetup.BottomMargin = Application.CentimetersToPoints(1.5) 'Margen Superior
        .PageSetup.TopMargin = Application.CentimetersToPoints(1) 'Margen Inferior
        .Range(.Cells(1, 13), .Cells(500, 13)).NumberFormat = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda(fmt)   'Establece el formato para los números
        .Range("A2:AZ1000").Font.Name = "Arial"    'Tipo de letra para toda la hoja
        .Range("A2:AZ1000").Font.Size = 8          'Tamaño de la letra
        
        Fila = Fila + 2
        NumCol = 0
        For i = 1 To grd.Cols - 1
            NumCol = NumCol + 1
            .Cells(Fila, NumCol) = grd.TextMatrix(0, i) 'cabeceras
            ReDim Preserve v(NumCol)
            v(NumCol - 1) = 0
        Next i
                
        .Range(.Cells(Fila, 1), .Cells(Fila, NumCol)).Font.Bold = True
        .Range(.Cells(Fila, 1), .Cells(Fila, NumCol)).Borders.LineStyle = 12
        .Range(.Cells(Fila, 3), .Cells(Fila, grd.Rows - 1)).HorizontalAlignment = xlHAlignLeft
        For i = 2 To grd.Rows - 1
            prg1.value = i
            Fila = Fila + 1
            If grd.IsSubtotal(i) = True Then
                .Range(.Cells(Fila, 1), .Cells(Fila, NumCol)).Font.Bold = True
'                .Cells(Fila, 2) = "TOTALES"
                .Cells(Fila, NumCol - 11) = grd.TextMatrix(i, grd.ColIndex("Total Kilos Producidos"))
                .Cells(Fila, NumCol - 10) = grd.TextMatrix(i, grd.ColIndex("Total Kilos Vendidos"))
                .Cells(Fila, NumCol - 8) = grd.TextMatrix(i, grd.ColIndex("Costo Variable"))
                .Cells(Fila, NumCol - 5) = grd.TextMatrix(i, grd.ColIndex("Utilidad Total Neta"))
                .Cells(Fila, NumCol - 4) = grd.TextMatrix(i, grd.ColIndex("Margen"))
                .Cells(Fila, NumCol - 3) = grd.TextMatrix(i, grd.ColIndex("Porcentaje M2"))
                .Cells(Fila, NumCol - 2) = grd.TextMatrix(i, grd.ColIndex("Produccion M2"))
                .Cells(Fila, NumCol - 1) = grd.TextMatrix(i, grd.ColIndex("Ventas M2"))
                .Cells(Fila, NumCol) = grd.TextMatrix(i, grd.ColIndex("Total Ventas"))
            Else
                j = 1
                mayor = 0
                For col = 1 To grd.Cols - 1
                            .Cells(Fila, j) = grd.TextMatrix(i, col)
                        mayor = Len(grd.TextMatrix(i, col)) 'Para ajustar el ancho de columnas
                        If mayor > v(j - 1) Then            'de acuerdo a la celda más grande
                            .Columns(j).ColumnWidth = mayor '13/11/2000 ---> Angel P.
                            v(j - 1) = mayor
                        End If
                        j = j + 1
                Next col
            End If
            .Range(.Cells(Fila, 1), .Cells(Fila, NumCol)).Borders.LineStyle = 1
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
      Case .ColIndex("P.V.")
        valor = .ValueMatrix(.Row, .ColIndex("P.V."))
        colu = .ColIndex("P.V.")
    Case .ColIndex("Peso")
        valor = .ValueMatrix(.Row, .ColIndex("Peso"))
        colu = .ColIndex("Peso")
    Case .ColIndex("Produccion M2")
        valor = .ValueMatrix(.Row, .ColIndex("Produccion M2"))
        colu = .ColIndex("Produccion M2")
    Case .ColIndex("Costo Variable")
        valor = .ValueMatrix(.Row, .ColIndex("Costo Variable"))
        colu = .ColIndex("Costo Variable")
    Case .ColIndex("Porcentaje M2")
        valor = .ValueMatrix(.Row, .ColIndex("Porcentaje M2"))
        colu = .ColIndex("Porcentaje M2")
    End Select
     For i = .Row To .Rows - 9
            If .TextMatrix(i, .ColIndex("bodega")) <> "." Then
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
    On Error GoTo ErrTrap
    
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
ErrTrap:
    If Err.Number <> 32755 Then
        MsgBox Err.Description
    End If
    Close NumFile
End Sub

Private Sub AbrirArchivo()
    Dim i As Long
    
    On Error GoTo ErrTrap
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
ErrTrap:
    If Err.Number <> 32755 Then DispErr
    Exit Sub
End Sub
Private Sub VisualizarTexto(ByVal archi As String)
    Dim f As Integer, s As String, i As Integer
    Dim Cadena
    On Error GoTo ErrTrap
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
ErrTrap:
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
              grd.TextMatrix(grd.Rows - 4, grd.ColIndex("Descripción")) = GetSetting(APPNAME, App.Title, "costofijomensual", 0)
              grd.Cell(flexcpBackColor, grd.Rows - 4, grd.ColIndex("Descripción"), grd.Rows - 4, grd.ColIndex("Descripción")) = &H80000005  'color blanco
              grd.Cell(flexcpForeColor, grd.Rows - 4, grd.ColIndex("Descripción"), grd.Rows - 4, grd.ColIndex("Descripción")) = &H80000012 'color negro
     
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


Private Sub AgregarDatosPlantilla()
 Dim f As Integer, s As String, i As Integer
    Dim Cadena, s1 As String, s2 As String
    On Error GoTo ErrTrap
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
ErrTrap:
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
        If s = grd.TextMatrix(i, grd.ColIndex("Descripción")) And Bodega = grd.TextMatrix(i, grd.ColIndex("bodega")) Then
                grd.TextMatrix(i, grd.ColIndex("Peso")) = peso
                grd.TextMatrix(i, grd.ColIndex("Costo Variable")) = costoVariable
                grd.TextMatrix(i, grd.ColIndex("Porcentaje M2")) = porcentaje
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
   If grd.TextMatrix(i, grd.ColIndex("Bodega")) <> "." Then
       If grd.TextMatrix(i, grd.ColIndex("Bodega")) = codbod Then
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
       If grd.TextMatrix(i, grd.ColIndex("bodega")) = codbod Then
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
          If grd.TextMatrix(i, grd.ColIndex("Bodega")) = codbod Then
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
      SaveSetting APPNAME, App.Title, "CostoFijoMensual", grd.ValueMatrix(grd.Rows - 4, grd.ColIndex("Descripción"))
   End With
      
End Sub


Private Sub LlenaFormatoFormulario()
    With grd
        .MergeCells = flexMergeSpill

        .AddItem "1" & vbTab & vbTab & vbTab & "DECLARACION DEL IMPUESTO AL VALOR AGREGADO" & vbTab & vbTab & vbTab & vbTab & vbTab & "No."
        .Cell(flexcpAlignment, 1, 8, 1, 8) = 1
        .RowHeight(1) = 400
        .Cell(flexcpFontSize, 1, 1, 1, 6) = 14
        .AddItem "2" & vbTab & vbTab & vbTab & "100" & vbTab & "IDENTIFICACION DE LA DECLARACION"
        .Cell(flexcpBackColor, 2, 3, 2, 9) = ColorFondo1
        .Cell(flexcpFontBold, 2, 3, 2, 9) = True
        .Cell(flexcpFontSize, 2, 1, 2, 8) = 10
        .RowHeight(2) = 320
        .AddItem "3" & vbTab & vbTab & vbTab & vbTab & "DECLARACION MENSUAL"
        .AddItem "4" & vbTab & vbTab & "FORMULARIO 104" & vbTab & "101" & vbTab & "MES " & vbTab & vbTab & vbTab & "102" & vbTab & "AÑO"
        .Cell(flexcpBackColor, 4, 3, 4, 3) = ColorFondo
        .Cell(flexcpBackColor, 4, 7, 4, 7) = ColorFondo
        .Cell(flexcpAlignment, 4, 8, 4, 8) = 1
        .Cell(flexcpFontBold, 4, 2, 4, 2) = True
        .Cell(flexcpFontSize, 4, 2, 4, 2) = 11
        .AddItem "5" & vbTab & vbTab & vbTab & vbTab & "DECLARACION SEMESTRAL"
        .AddItem "6" & vbTab & vbTab & vbTab & "103" & vbTab & "SEMESTRE   ENERO - JUNIO" & vbTab & "01" & vbTab & vbTab & "104" & vbTab & "No. FORMULARIO QUE SE RETIFICA"
        .AddItem "7" & vbTab & vbTab & vbTab & vbTab & "                       JULIO DICIEMBRE" & vbTab & "02"
        .Cell(flexcpBackColor, 6, 3, 6, 3) = ColorFondo
        .Cell(flexcpBackColor, 6, 7, 6, 7) = ColorFondo
        .Cell(flexcpAlignment, 6, 8, 6, 8) = 1
        .AddItem "8" & vbTab & "200" & vbTab & "IDENTIFICACION DEL SUJETO PASIVO (AJENTE DE PERCEPCION O RETENCION)"
        .Cell(flexcpBackColor, 8, 1, 8, 9) = ColorFondo1
        .Cell(flexcpFontBold, 8, 1, 8, 9) = True
        .Cell(flexcpFontSize, 8, 1, 8, 8) = 10
        .RowHeight(8) = 320
        .AddItem "9" & vbTab & vbTab & "RUC" & vbTab & vbTab & "RAZON SOCIAL DENOMINACION O APELLIDOS Y NOMBRES COMPLETOS"
        .AddItem "10" & vbTab & "201" & vbTab & vbTab & "202" & vbTab
        .Cell(flexcpBackColor, 9, 1, 10, 1) = ColorFondo
        .Cell(flexcpBackColor, 9, 3, 10, 3) = ColorFondo
        .AddItem "11" & vbTab & "300" & vbTab & "PROPORCION DE CREDITO TRIBUTARIO APLICABLE EN ESTE MES" & vbTab & vbTab & vbTab & "Devoluciones de IVA Solicitada y Recibidas"
        .Cell(flexcpBackColor, 11, 1, 11, 9) = ColorFondo1
        .Cell(flexcpFontBold, 11, 1, 11, 9) = True
        .Cell(flexcpFontSize, 11, 1, 11, 8) = 10
        .RowHeight(11) = 320
        .AddItem "12" & vbTab & "VENTAS CON TARIFA 12+ EXPORTACIONES" & vbTab & vbTab & "301" & vbTab & vbTab & "SALDOS DEL MES ANTERIOR" & vbTab & vbTab & "351"
        .AddItem "13" & vbTab & "SALDO DEL CREDITO TRIBUTARIO MES ANETRIOR" & vbTab & vbTab & "303" & vbTab & vbTab & "(+) SOLICITUD DEVOLUCION IVA" & vbTab & vbTab & "353"
        .AddItem "14" & vbTab & "(-) DEVOLUCIONES DE IVA ESTE MES" & vbTab & vbTab & "305" & vbTab & vbTab & "(-) DEVOLUCION RESIBIDAS EN EL MES" & vbTab & vbTab & "355"
        .AddItem "15" & vbTab & "(=) SALDO CREDITO TRIBUTARIO APLICARSE" & vbTab & vbTab & "399" & vbTab & vbTab & "(=) SALDO FINAL MES 351+353-355" & vbTab & vbTab & "357"
        .Cell(flexcpBackColor, 12, 3, 15, 3) = ColorFondo
        .Cell(flexcpBackColor, 12, 7, 15, 7) = ColorFondo
        .Cell(flexcpAlignment, 12, 4, 15, 4) = 2
'        '  VENTAS Y EXPORTACIONES
        .AddItem "16" & vbTab & "500" & vbTab & "RESUMEN DE VENTAS Y OTRAS OPERACIONES DEL PERIODO QUE DECLARA" & vbTab & vbTab & vbTab & vbTab & "BASE IMPONIBLE" & vbTab & vbTab & "IMPUESTO"
        .Cell(flexcpBackColor, 16, 1, 16, 9) = ColorFondo1
        .Cell(flexcpFontBold, 16, 1, 16, 9) = True
        .Cell(flexcpFontSize, 16, 1, 16, 8) = 10
        .RowHeight(16) = 320
        .AddItem "17" & vbTab & "VENTAS NETAS GRAVADAS CON TARIFA 12% (EXCLUYE ACTIVOS FIJOS Y OTROS)" & vbTab & vbTab & vbTab & vbTab & "501" & vbTab & vbTab & "551"
        .AddItem "18" & vbTab & "VENTAS DE ACTIVOS FIJOS GRAVADAS CON TARIFA 12%" & vbTab & vbTab & vbTab & vbTab & "503" & vbTab & vbTab & "553"
        .AddItem "19" & vbTab & "OTROS CON TARIFA 12% " & vbTab & vbTab & vbTab & vbTab & "505" & vbTab & vbTab & "555"
        .AddItem "20" & vbTab & "(-) DEVOLUCIONES EN VENTAS MEDIANTE NOTA DE CREDITO CON IVA 12%" & vbTab & vbTab & vbTab & vbTab & "507" & vbTab & vbTab & "557"
        .AddItem "21" & vbTab & "VENTAS NETAS GRABADAS CON TARIFA CERO" & vbTab & vbTab & vbTab & vbTab & "509"
        .AddItem "22" & vbTab & "VENTAS DE ACTIVOS FIJO GRABADAS CON TARIFA CERO" & vbTab & vbTab & vbTab & vbTab & "511"
        .AddItem "23" & vbTab & "EXPORTACION DE BIENES" & vbTab & vbTab & vbTab & vbTab & "513"
        .AddItem "24" & vbTab & "EXPORTACION DE SERVICIOS" & vbTab & vbTab & vbTab & vbTab & "515"
        .AddItem "25" & vbTab & "TOTAL VENTAS Y EXPORTACIONES " & vbTab & vbTab & vbTab & "501+505-507+509+513+515" & vbTab & "517"
        .AddItem "26" & vbTab & "TOTAL IMPUESTO" & vbTab & vbTab & vbTab & vbTab & vbTab & "551+553+555-557" & vbTab & "599"
        .Cell(flexcpBackColor, 17, 5, 25, 5) = ColorFondo
        .Cell(flexcpBackColor, 17, 7, 26, 7) = ColorFondo
        .Cell(flexcpBackColor, 21, 8, 25, 8) = ColorFondo
        .Cell(flexcpBackColor, 26, 6, 26, 6) = ColorFondo
'        .Cell(flexcpBackColor, 25, 6, 25, 6) = ColorFondo
'        'COMPRAS E IMPORTACIONES
        .AddItem "27" & vbTab & "600" & vbTab & "RESUMEN DE COMPRAS Y OTRAS OPERACIONES DEL PERIODO QUE DECLARA" & vbTab & vbTab & vbTab & vbTab & "BASE IMPONIBLE" & vbTab & vbTab & "IMPUESTO"
        .Cell(flexcpBackColor, 27, 1, 27, 9) = ColorFondo1
        .Cell(flexcpFontBold, 27, 1, 27, 8) = True
        .Cell(flexcpFontSize, 27, 1, 27, 8) = 10
        .RowHeight(27) = 320
        .AddItem "28" & vbTab & "COMPRAS LOCALES NETAS GRAVADAS CON TARIFA 12% (EXCLUYE ACTIVOS FIJOS)" & vbTab & vbTab & vbTab & vbTab & "601" & vbTab & vbTab & "651"
        .AddItem "29" & vbTab & "COMPRAS LOCALES DE SERVICIOS GRAVADAS CON TARIFA 12%" & vbTab & vbTab & vbTab & vbTab & "603" & vbTab & vbTab & "653"
        .AddItem "30" & vbTab & "COMPRAS LOCALES DE ACTIVOS GRAVADAS CON TARIFA 12%" & vbTab & vbTab & vbTab & vbTab & "605" & vbTab & vbTab & "655"
        .AddItem "31" & vbTab & "IMPORTACIONES GRAVADAS  CON TARIFA 12% (EXCLUYE ACTIVOS FIJOS )" & vbTab & vbTab & vbTab & vbTab & "607" & vbTab & vbTab & "657"
        .AddItem "32" & vbTab & "IMPORTACIONES DE ACTIVOS FIJOS  GRAVADAS  CON TARIFA 12% " & vbTab & vbTab & vbTab & vbTab & "609" & vbTab & vbTab & "659"
        .AddItem "33" & vbTab & "(-) DEVOLUCIONES DE BIENES MEDIANTE NOTA DE CREDITO CON IVA 12%" & vbTab & vbTab & vbTab & vbTab & "611" & vbTab & vbTab & "561"
        .AddItem "34" & vbTab & "IVA SOBRE EL VALOR DE LA DEPRECIACION DE ACTIVOS EN INTERNACION TEMPORAL" & vbTab & vbTab & vbTab & vbTab & "613" & vbTab & vbTab & "563"
        .AddItem "35" & vbTab & "IVA EN EL LEASING INTERNACIONAL" & vbTab & vbTab & vbTab & vbTab & "615" & vbTab & vbTab & "565"
        .AddItem "36" & vbTab & "COMPRAS LOCALES DE BIENES Y SERVICIOS GRABADAS CON TARIFA CERO" & vbTab & vbTab & vbTab & vbTab & "617"
        .AddItem "37" & vbTab & "IMPORTACIONES GRABADAS CON TARIFA CERO" & vbTab & vbTab & vbTab & vbTab & "619"
        .AddItem "38" & vbTab & "SUBTOTAL CREDITO TRIBUTARIO DEL MES" & vbTab & vbTab & vbTab & vbTab & vbTab & "(651+653+655+657+659-661+663+665)x301" & vbTab & "699"
        .Cell(flexcpBackColor, 28, 5, 37, 5) = ColorFondo
        .Cell(flexcpBackColor, 28, 7, 38, 7) = ColorFondo
        .Cell(flexcpBackColor, 36, 8, 37, 8) = ColorFondo
        .Cell(flexcpBackColor, 38, 6, 38, 6) = ColorFondo
       
'        'RESUMEN
        .AddItem "39" & vbTab & "700" & vbTab & "RESUMEN IMPOSITIVO"
        .Cell(flexcpBackColor, 39, 1, 39, 9) = ColorFondo1
        .Cell(flexcpFontBold, 39, 1, 39, 9) = True
        .Cell(flexcpFontSize, 39, 1, 39, 8) = 10
        .RowHeight(39) = 320
        .AddItem "40" & vbTab & "IMPUESTO RESULTANTE DEL MES " & vbTab & vbTab & vbTab & vbTab & vbTab & "599-699" & vbTab & "701"
        .AddItem "41" & vbTab & "(-) SALDO DE CREDITO TRIBUTARIO MES ANTERIOR " & vbTab & vbTab & vbTab & vbTab & vbTab & "Trasladar campo 399" & vbTab & "703"
        .AddItem "42" & vbTab & "(-) RETENCIONES EN LA FUENTE DE IVA QUE LE HAN SIDO EFECTUADAS " & vbTab & vbTab & vbTab & vbTab & vbTab & "" & vbTab & "705"
        .AddItem "43" & vbTab & "(=) SALDO DE CREDITO TRIBUTARIO MES ANTERIOR " & vbTab & vbTab & vbTab & vbTab & vbTab & "701-703-705 <= 0" & vbTab & "798"
        .AddItem "44" & vbTab & "(=) SUBTOTAL A PAGAR " & vbTab & vbTab & vbTab & vbTab & vbTab & "701-703-705 >= 0" & vbTab & "799"
        .Cell(flexcpBackColor, 40, 7, 44, 7) = ColorFondo
'        'DECLARACION DEL SUJETO PASIVO COMO AGENTE RETENCION IVA
        .AddItem "45" & vbTab & "800" & vbTab & "DECLARACION DEL SUJETO PASIVO COMO AGENTE RETENCION IVA" & vbTab & vbTab & vbTab & vbTab & "VALOR DEL IVA" & vbTab & vbTab & "VALOR RETENIDO"
        .Cell(flexcpBackColor, 45, 1, 45, 9) = ColorFondo1
        .Cell(flexcpFontBold, 45, 1, 45, 9) = True
        .Cell(flexcpFontSize, 45, 1, 45, 8) = 10
        .RowHeight(45) = 320
        .AddItem "46" & vbTab & "IVA CAUSADO POR LA COMPRA DE BIENES" & vbTab & vbTab & vbTab & vbTab & "801" & vbTab & vbTab & "851" & vbTab & vbTab & "30%"
        .AddItem "47" & vbTab & "IVA RETENIDO POR EMPRESAS EMISORAS" & vbTab & vbTab & "BIENES" & vbTab & vbTab & "803" & vbTab & vbTab & "853" & vbTab & vbTab & "30%"
        .AddItem "48" & vbTab & "DE TARJETAS DE CREDITO" & vbTab & vbTab & "SERVICIOS" & vbTab & vbTab & "805" & vbTab & vbTab & "855" & vbTab & vbTab & "70%"
        .AddItem "49" & vbTab & "IVA CAUSADO POR LA PRESTACION DE SERVICIOS" & vbTab & vbTab & vbTab & vbTab & "807" & vbTab & vbTab & "857" & vbTab & vbTab & "70%"
        .AddItem "50" & vbTab & "IVA CAUSADO POR LA PRESTACION DE SERVICIOS PROFESIONALES" & vbTab & vbTab & vbTab & vbTab & "809" & vbTab & vbTab & "859" & vbTab & vbTab & "100%"
        .AddItem "51" & vbTab & "IVA CAUSADO POR ELARRENDAMIENTO DE INMUEBLES A PERSONAS NATURALES" & vbTab & vbTab & vbTab & vbTab & "811" & vbTab & vbTab & "861" & vbTab & vbTab & "100%"
        .AddItem "52" & vbTab & "IVA CAUSADO EN LA DISTRIBUCION DE COMBUSTIBLES" & vbTab & vbTab & vbTab & vbTab & "813" & vbTab & vbTab & "863" & vbTab & vbTab & "100%"
        .AddItem "53" & vbTab & "IVA CAUSADO EN OTRAS COMPRA DE BIENES Y SERVICIOS CON EMISION DE LIQUIDACION DE COMPRAS Y PRESTACION SERVICIOS" & vbTab & vbTab & vbTab & vbTab & "815" & vbTab & vbTab & "865" & vbTab & vbTab & "100%"
        .AddItem "54" & vbTab & "IVA CAUSADO EN LA DEPRECIACION DE ACTIVOS EN INTERNACION TEMPORAL" & vbTab & vbTab & vbTab & vbTab & "817" & vbTab & vbTab & "867" & vbTab & vbTab & "100%"
        .AddItem "55" & vbTab & "IVA CAUSADO EN LEASING INTERNACIONAL" & vbTab & vbTab & vbTab & vbTab & "819" & vbTab & vbTab & "869" & vbTab & vbTab & "100%"
        .AddItem "56" & vbTab & "Declaro que los datos contenidos en esta declaración son verdaderos, por lo" & vbTab & vbTab & vbTab & vbTab & "TOTAL RETENIDO            SUMAR 851 AL 869 " & vbTab & vbTab & "898"
        .AddItem "57" & vbTab & "que asumo la responsabilidad correspondiente (Artículo 98 de la L.R.T.I.)" & vbTab & vbTab & vbTab & vbTab & "TOTAL IVA A PAGAR            799+898 " & vbTab & vbTab & "899"
        .Cell(flexcpBackColor, 46, 5, 55, 5) = ColorFondo
        .Cell(flexcpBackColor, 46, 7, 57, 7) = ColorFondo
        .Cell(flexcpBackColor, 47, 1, 48, 2) = ColorFondo
'        'A PAGAR
        .AddItem "58" & vbTab & vbTab & vbTab & vbTab & vbTab & "900" & vbTab & "VALORES A PAGAR Y FORMA DE PAGO"
        .Cell(flexcpBackColor, 58, 5, 58, 9) = ColorFondo1
        .Cell(flexcpFontBold, 58, 5, 58, 9) = True
        .Cell(flexcpFontSize, 58, 5, 58, 8) = 10
        .RowHeight(58) = 320
        .Cell(flexcpAlignment, 58, 6, 58, 6) = 1
        .AddItem "59" & vbTab & vbTab & vbTab & vbTab & vbTab & "PAGO PREVIO" & vbTab & vbTab & "901"
        .AddItem "60" & vbTab & vbTab & "____________________________" & vbTab & vbTab & "____________________________" & vbTab & "TOTAL IMPUESTO A PAGAR    899-901" & vbTab & vbTab & "902"
        .AddItem "61" & vbTab & vbTab & "FIRMA CONTRIBUYENTE(Rep. Legal)" & vbTab & vbTab & "FIRMA CONTADOR" & vbTab & "INTERESES POR MORA" & vbTab & vbTab & "903"
        .AddItem "62" & vbTab & vbTab & "NOMBRE:" & vbTab & vbTab & "NOMBRE:" & vbTab & "MULTAS" & vbTab & vbTab & "904"
        .AddItem "63" & vbTab & "198" & vbTab & "C.I. No." & vbTab & "199" & vbTab & "RUC No." & vbTab & "TOTAL PAGADO     902+903+904" & vbTab & vbTab & "999"
        .Cell(flexcpBackColor, 59, 7, 63, 7) = ColorFondo
        .Cell(flexcpBackColor, 63, 1, 63, 1) = ColorFondo
        .Cell(flexcpBackColor, 63, 3, 63, 3) = ColorFondo
        .AddItem "64" & vbTab
        .Cell(flexcpBackColor, 64, 1, 64, 9) = ColorFondo
        .AddItem "65" & vbTab & "MEDIANTE CHEQUE, DEBITO BANCARIO, EFECTIVO U OTRAS FORMAS DE COBRO" & vbTab & vbTab & vbTab & vbTab & "905" & vbTab & "US $."
        .AddItem "66" & vbTab & "MEDIANTE COMPENSACIONES" & vbTab & vbTab & vbTab & "TOTAL" & vbTab & "906" & vbTab & "US $."
        .AddItem "67" & vbTab & "MEDIANTE NOTAS DE CREDITO" & vbTab & vbTab & vbTab & "TOTAL" & vbTab & "906" & vbTab & "US $."
        .Cell(flexcpBackColor, 65, 5, 67, 5) = ColorFondo
        .Cell(flexcpFontBold, 10, 1, 10, 1) = True
        .Cell(flexcpFontBold, 63, 1, 63, 1) = True
        .Cell(flexcpFontBold, 1, 3, 67, 3) = True
        .Cell(flexcpFontBold, 17, 5, 55, 5) = True
        .Cell(flexcpFontBold, 65, 5, 67, 5) = True
        .Cell(flexcpFontBold, 1, 7, 67, 7) = True
     
'        .RowHeight(0) = 1
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .MergeCol(4) = True
        .MergeCol(5) = True
        .Redraw = flexRDBuffered

 
        .Refresh
    End With
End Sub

Private Sub LLENADATOS()
    With grd
        .Redraw = flexRDBuffered
        .TextMatrix(10, 2) = gobjMain.EmpresaActual.GNOpcion.ruc
        .TextMatrix(10, 4) = gobjMain.EmpresaActual.GNOpcion.NombreEmpresa
         .Refresh
    End With

End Sub

Private Sub CambiaFondoCeldasEditables()
    With grd
        .Redraw = flexRDBuffered
        .Cell(flexcpBackColor, 4, 5, 4, 5) = vbWhite 'mes
        .Cell(flexcpBackColor, 5, 8, 5, 8) = vbWhite 'año
        .Cell(flexcpBackColor, 7, 8, 7, 8) = vbWhite ' formulario modifica
        .Cell(flexcpBackColor, 10, 2, 10, 2) = vbWhite 'ruc
        .Cell(flexcpBackColor, 10, 4, 10, 4) = vbWhite 'razon social
        
        .Cell(flexcpBackColor, 12, 4, 14, 4) = vbWhite
        .Cell(flexcpBackColor, 12, 8, 14, 8) = vbWhite
        .Cell(flexcpBackColor, 17, 6, 24, 6) = vbWhite
        .Cell(flexcpBackColor, 28, 6, 37, 6) = vbWhite
        .Cell(flexcpBackColor, 42, 8, 42, 8) = vbWhite
        .Cell(flexcpBackColor, 46, 6, 55, 6) = vbWhite
        .Cell(flexcpBackColor, 59, 8, 59, 8) = vbWhite
        .Cell(flexcpBackColor, 61, 8, 62, 8) = vbWhite
    End With
End Sub

Private Sub BuscarComprasNetas12y0()
    Dim sql As String, cond As String, CadenaValores As String
    Dim OrdenadoX As String
    Dim CadenaAgrupa  As String, Recargo As String
    Dim v As Variant, max As Integer, i As Integer
    Dim from As String, NumReg As Long, f1 As String
    Dim rs As Recordset
    Dim subtotal As Currency, CompraBienesTarifa_0 As Currency, CompraServiciosTarifa_0 As Currency
    Dim CompraActivosTarifa_0 As Currency
    Dim CP_Ser As String, CP_Act As String, CP_Dev As String
    Dim VT_Bie As String, VT_Ser As String, VT_Dev As String
    Dim Reten As String, ret_real As String, ret_recib As String
    Dim MONEDA As String
    Set objcond = gobjMain.objCondicion
'    If Not frmB_FormSRI.Inicio(objcond, Recargo, CP_Ser, CP_Act, CP_Dev, VT_Dev, VT_Bie, VT_Ser, ret_real, ret_recib, Reten) Then
'        grd.SetFocus
'        Exit Sub
'    End If
    With objcond

'            cond = " AND (GNC.FechaTrans  BETWEEN " & _
'                    FechaYMD(.Fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
'                    FechaYMD(.Fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            
                    'Reporte de un mes a la vez
        f1 = DateSerial(Year(.fecha1), Month(.fecha1), 1)
        cond = " AND GNC.FechaTrans BETWEEN " & FechaYMD(f1, gobjMain.EmpresaActual.TipoDB) & _
               " AND " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB)

            
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(.codforma) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(.CodTrans) & ")" & _
                  "Group by IvkR.TransID"
                  
            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
            
            '--datos de la compra bienes tarifa 12
            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0, "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBaseIVA + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS Valor12 "
            sql = sql & " Into tmp1"
            sql = sql & " from    (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(.CodTrans) & ")"
            sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
            sql = " Select  isnull(sum(Valor0),0) as ValorTotal0, isnull(sum(Valor12),0) as ValorTotal12 from tmp1  "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            CompraBienesTarifa_0 = Round(rs.Fields("ValorTotal0"), 2)
            grd.TextMatrix(28, 6) = Round(rs.Fields("ValorTotal12"), 2)
'            grd.TextMatrix(28, 8) = Round(grd.ValueMatrix(28, 6) * gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA, 2)
            
            '--datos de la compra servicios tarifa 12
            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0, "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBaseIVA + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS Valor12 "
            sql = sql & " Into tmp1"
            sql = sql & " from    (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(CP_Ser) & ")"
            sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
            sql = " Select  isnull(sum(Valor0),0) as ValorTotal0, isnull(sum(Valor12),0) as ValorTotal12 from tmp1  "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            CompraServiciosTarifa_0 = Round(rs.Fields("ValorTotal0"), 2)
            grd.TextMatrix(29, 6) = Round(rs.Fields("ValorTotal12"), 2)
            'grd.TextMatrix(29, 8) = Round(grd.ValueMatrix(29, 6) * gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA, 2)
            
        '--datos de la compra activos tarifa 12
            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0, "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBaseIVA + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS Valor12 "
            sql = sql & " Into tmp1"
            sql = sql & " from    (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(CP_Act) & ")"
            sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1
            
            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
            sql = " Select  isnull(sum(Valor0),0) as ValorTotal0, isnull(sum(Valor12),0) as ValorTotal12 from tmp1  "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            CompraActivosTarifa_0 = rs.Fields("ValorTotal0")
            grd.TextMatrix(30, 6) = Round(rs.Fields("ValorTotal12"), 2)
            
        '--datos de la devoluciones tarifa 12
            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0, "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBaseIVA + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS Valor12 "
            sql = sql & " Into tmp1"
            sql = sql & " from    (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(CP_Dev) & ")"
            sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1
            
            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
            sql = " Select  isnull(sum(Valor0),0)*-1 as ValorTotal0, isnull(sum(Valor12),0)*-1 as ValorTotal12 from tmp1  "
            
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            grd.TextMatrix(33, 6) = Round(rs.Fields("ValorTotal12"), 2) * -1
            
            'compra de bienes y servicios tarifa 0
            grd.TextMatrix(36, 6) = Round(CompraServiciosTarifa_0 + CompraBienesTarifa_0, 2 + CompraActivosTarifa_0)
            
'            For i = 28 To 35
'                subtotal = subtotal + grd.ValueMatrix(i, 8)
'            Next i
'            grd.TextMatrix(38, 8) = Round(subtotal, 2)
                            
    '********** VENTAS BIENES
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 "
            sql = sql & "From IvRecargo ivR inner join "
            sql = sql & " IvKardexRecargo ivkR Inner join "
            sql = sql & " GnComprobante gnc Inner join PcPRovCLi on gnc.IdClienteRef = PCProvCli.IdProvCli "
            sql = sql & " On ivkr.TransID = gnc.TransID "
            sql = sql & " On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gnc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(Recargo) & ") " & cond
            sql = sql & " Group by IvkR.TransID"
            
            VerificaExistenciaTabla 0
            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg

            sql = "SELECT "
            sql = sql & " isnull(SUM(PrecioTotalBase0 + (PrecioTotalBase0 * (cast(isnull(TotalDescuento,0) as float) / cast(PrecioTotal as float)))),0) as Base0, "
            sql = sql & " isnull(SUM(PrecioTotalBaseIVA + (PrecioTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(PrecioTotal as float)))),0) As BaseIVA "
            sql = sql & " FROM tmp0 Right join "
            sql = sql & "vwConsSUMIVKardexIVA inner join "
            sql = sql & "GNComprobante GNC   "
            sql = sql & "ON vwConsSUMIVKardexIVA.TransID = GNC.TransID "
            sql = sql & "ON tmp0.TransID = GNC.TransID"
            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(VT_Bie) & ")  " & cond
            
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            grd.TextMatrix(17, 6) = Round(rs.Fields("BaseIVA"), 2)
            grd.TextMatrix(21, 6) = Round(rs.Fields("Base0"), 2)
            
            
    '********** VENTAS SERVICIOS
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 "
            sql = sql & "From IvRecargo ivR inner join "
            sql = sql & " IvKardexRecargo ivkR Inner join "
            sql = sql & " GnComprobante gnc Inner join PcPRovCLi on gnc.IdClienteRef = PCProvCli.IdProvCli "
            sql = sql & " On ivkr.TransID = gnc.TransID "
            sql = sql & " On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gnc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(Recargo) & ") " & cond
            sql = sql & " Group by IvkR.TransID"
            
            VerificaExistenciaTabla 0
            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg

            sql = "SELECT "
            sql = sql & " isnull(SUM(PrecioTotalBase0 + (PrecioTotalBase0 * (cast(isnull(TotalDescuento,0) as float) / cast(PrecioTotal as float)))),0) as Base0, "
            sql = sql & " isnull(SUM(PrecioTotalBaseIVA + (PrecioTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(PrecioTotal as float)))),0) As BaseIVA "
            sql = sql & " FROM tmp0 Right join "
            sql = sql & "vwConsSUMIVKardexIVA inner join "
            sql = sql & "GNComprobante GNC   "
            sql = sql & "ON vwConsSUMIVKardexIVA.TransID = GNC.TransID "
            sql = sql & "ON tmp0.TransID = GNC.TransID"
            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(VT_Ser) & ")" & cond
            
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            grd.TextMatrix(18, 6) = Round(rs.Fields("BaseIVA"), 2)
            grd.TextMatrix(22, 6) = Round(rs.Fields("Base0"), 2)
            
'********** VENTAS DEVOLUCIONES
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 "
            sql = sql & "From IvRecargo ivR inner join "
            sql = sql & " IvKardexRecargo ivkR Inner join "
            sql = sql & " GnComprobante gnc Inner join PcPRovCLi on gnc.IdClienteRef = PCProvCli.IdProvCli "
            sql = sql & " On ivkr.TransID = gnc.TransID "
            sql = sql & " On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gnc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(Recargo) & ") " & cond
            sql = sql & " Group by IvkR.TransID"
            
            VerificaExistenciaTabla 0
            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg

            sql = "SELECT "
            sql = sql & " isnull(SUM(PrecioTotalBase0 + (PrecioTotalBase0 * (cast(isnull(TotalDescuento,0) as float) / cast(PrecioTotal as float)))),0) as Base0, "
            sql = sql & " isnull(SUM(PrecioTotalBaseIVA + (PrecioTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(PrecioTotal as float)))),0) As BaseIVA "
            sql = sql & " FROM tmp0 Right join "
            sql = sql & "vwConsSUMIVKardexIVA inner join "
            sql = sql & "GNComprobante GNC   "
            sql = sql & "ON vwConsSUMIVKardexIVA.TransID = GNC.TransID "
            sql = sql & "ON tmp0.TransID = GNC.TransID"
            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(VT_Dev) & ")" & cond
            
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            grd.TextMatrix(20, 6) = Round(rs.Fields("BaseIVA"), 2) * -1
            'grd.TextMatrix(24, 6) = Round(rs.Fields("Base0"), 2) * -1
                       
            ' RETENCIONES RECIBIDAS
            .CodMoneda = MONEDA_PRE
            MONEDA = IIf(.NumMoneda > 0, "/Cotizacion" & .NumMoneda + 1, "")
            sql = "SELECT ISNULL(sum(Valor" & MONEDA & "),0) as TotalRetRecibidas "
            sql = sql & " FROM vwConsRetencion "
            sql = sql & " WHERE CodTrans IN (" & PreparaCadena(ret_recib) & ")"
            sql = sql & "  AND FechaTrans BETWEEN " & FechaYMD(f1, gobjMain.EmpresaActual.TipoDB)
            sql = sql & "  AND " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB)
            sql = sql & "  and  DEBE > 0 and bandIVA=1"
        
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                       
            grd.TextMatrix(42, 8) = Round(rs.Fields("TotalRetRecibidas"), 2)
            
            ' RETENCIONES REALIZADAS
            sql = "SELECT CodSRI, ROUND(Porcentaje,0) as PORCENTAJE, "
            sql = sql & " ISNULL(sum(Valor" & MONEDA & "),0) as TotalRetRealizadas "
            sql = sql & " FROM vwConsRetencion "
            sql = sql & " WHERE CodTrans IN (" & PreparaCadena(ret_real) & ")"
            sql = sql & "  AND FechaTrans BETWEEN " & FechaYMD(f1, gobjMain.EmpresaActual.TipoDB)
            sql = sql & "  AND " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB)
            sql = sql & "  and  HABER > 0 and bandIVA=1 "
            sql = sql & "  group by CodSRI, Porcentaje"
            
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            rs.MoveFirst
            While Not rs.EOF
                Select Case rs.Fields("PORCENTAJE")
                    Case 30:
                        grd.TextMatrix(46, 8) = Round(rs.Fields("TotalRetRealizadas"), 2)
                    Case 70:
                        grd.TextMatrix(49, 8) = Round(rs.Fields("TotalRetRealizadas"), 2)
                    Case 100:
                        grd.TextMatrix(50, 8) = Round(rs.Fields("TotalRetRealizadas"), 2)
                End Select
                rs.MoveNext
            Wend
            
            
            CalcularPorcentajes
            grd.Refresh
    End With
End Sub
    

Private Sub CalcularPorcentajes()
    Dim i As Integer, SubtotalCompras As Currency, SubtotalVentas As Currency
    Dim SubtotalRetenido As Currency
    
    'mes anterior credito tributario 399=303-305
    
    grd.TextMatrix(15, 4) = grd.ValueMatrix(13, 4) - grd.ValueMatrix(14, 4)
    
    'mes anterior credito tributario 703=399
    
    grd.TextMatrix(41, 8) = grd.ValueMatrix(15, 4)
    
    
    'ventas 599
    
    For i = 17 To 24
        If Len(grd.TextMatrix(i, 6)) > 0 Then
            SubtotalVentas = SubtotalVentas + grd.ValueMatrix(i, 6)
        End If
    Next i
   grd.TextMatrix(25, 6) = Round(SubtotalVentas, 2)
    
    SubtotalVentas = 0
    
    For i = 17 To 20
        If Len(grd.TextMatrix(i, 6)) > 0 Then
            grd.TextMatrix(i, 8) = Round(grd.ValueMatrix(i, 6) * gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA, 2)
            SubtotalVentas = SubtotalVentas + grd.ValueMatrix(i, 8)
        Else
            grd.TextMatrix(i, 8) = ""
        End If
    Next i
   grd.TextMatrix(26, 8) = Round(SubtotalVentas, 2)
    
    'calculo 301= (501+505-507+513+515)/517
    If grd.ValueMatrix(25, 6) <> 0 Then
        grd.TextMatrix(12, 4) = Round((grd.ValueMatrix(17, 6) + grd.ValueMatrix(19, 6) + grd.ValueMatrix(20, 6) + grd.ValueMatrix(23, 6) + grd.ValueMatrix(24, 6)) / grd.ValueMatrix(25, 6), 0)
    End If
    grd.Cell(flexcpAlignment, 12, 4, 15, 4) = 7
    
    
    'compras 699
    For i = 28 To 35
        If Len(grd.TextMatrix(i, 6)) > 0 Then
            grd.TextMatrix(i, 8) = Round(grd.ValueMatrix(i, 6) * gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA, 2)
            SubtotalCompras = SubtotalCompras + grd.ValueMatrix(i, 8)
        Else
            grd.TextMatrix(i, 8) = ""
        End If
    Next i
   grd.TextMatrix(38, 8) = Round(SubtotalCompras * grd.ValueMatrix(12, 4), 2)

    ' 701=599-699
    grd.TextMatrix(40, 8) = (grd.ValueMatrix(26, 8) - grd.ValueMatrix(38, 8))
    ' 703=599-699
    grd.TextMatrix(41, 8) = grd.ValueMatrix(15, 4)
        
    '798 o 799
    If (grd.ValueMatrix(40, 8) - grd.ValueMatrix(41, 8) - grd.ValueMatrix(42, 8)) < 0 Then
        grd.TextMatrix(43, 8) = Abs(grd.ValueMatrix(40, 8) - grd.ValueMatrix(41, 8) - grd.ValueMatrix(42, 8))
        grd.TextMatrix(44, 8) = "0"
    Else
        grd.TextMatrix(43, 8) = "0"
        grd.TextMatrix(44, 8) = grd.ValueMatrix(40, 8) - grd.ValueMatrix(41, 8) - grd.ValueMatrix(42, 8)
    End If
    '898=851+853+855+857+859+861+863+865+867+869
    For i = 46 To 55
        If Len(grd.TextMatrix(i, 8)) > 0 Then
            SubtotalRetenido = SubtotalRetenido + Round(grd.ValueMatrix(i, 8), 2)
        End If
    Next i
    grd.TextMatrix(56, 8) = SubtotalRetenido
    
    '899=799+898
    grd.TextMatrix(57, 8) = grd.ValueMatrix(56, 8) + grd.ValueMatrix(44, 8)
    
    '902=899-901
    grd.TextMatrix(60, 8) = grd.ValueMatrix(57, 8) - grd.ValueMatrix(59, 8)
    
    '999=902+903+904
    grd.TextMatrix(63, 8) = grd.ValueMatrix(60, 8) + grd.ValueMatrix(61, 8) + grd.ValueMatrix(62, 8)

    
End Sub

