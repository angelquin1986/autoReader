VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFormulario1042014 
   BackColor       =   &H80000013&
   Caption         =   "Costos"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   240
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
      Left            =   60
      TabIndex        =   1
      Top             =   600
      Width           =   6135
      _cx             =   10821
      _cy             =   5530
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
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
      BackColorBkg    =   -2147483633
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
      FormatString    =   $"frmFormulario1042014.frx":0000
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
      BackColorFrozen =   -2147483633
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
            Picture         =   "frmFormulario1042014.frx":00D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario1042014.frx":01E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario1042014.frx":063D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario1042014.frx":0751
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario1042014.frx":0865
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario1042014.frx":0CB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario1042014.frx":0F1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario1042014.frx":1BF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario1042014.frx":2047
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario1042014.frx":2159
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario1042014.frx":39DB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   953
      ButtonWidth     =   2090
      ButtonHeight    =   953
      Style           =   1
      ImageList       =   "img1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Abrir"
            Key             =   "Abrir"
            Object.ToolTipText     =   "Abrir"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo"
            Key             =   "Nuevo"
            Description     =   "Nuevo"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar (F5)"
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Impresion"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exp. Excel"
            Key             =   "Excel"
            Object.ToolTipText     =   "A Excel"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exp. xml"
            Key             =   "XML"
            Object.ToolTipText     =   "A archivo xml"
            ImageIndex      =   2
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
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   9720
      TabIndex        =   0
      Top             =   3705
      Width           =   9720
      Begin VB.CommandButton cmdExplorar 
         Caption         =   "..."
         Height          =   310
         Left            =   7800
         TabIndex        =   6
         Top             =   480
         Width           =   372
      End
      Begin VB.TextBox txtDestino 
         Height          =   320
         Left            =   780
         TabIndex        =   5
         Top             =   480
         Width           =   6972
      End
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Destino  "
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   7
         Top             =   480
         Width           =   630
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
Attribute VB_Name = "frmFormulario1042014"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ex As Excel.Application, ws As Worksheet, wkb As Workbook
Private mProcesando As Boolean
Private mCancelado As Boolean
Private mcolItemsSelec As Collection      'Coleccion de items df
'jeaa 24/09/04 asignacion de grupo a los items
Dim v() As String
Dim costoFijoMensual As Currency, Precio As Integer
Private Const ColorFondo = &H8000000A
Private Const ColorFondo1 = &H80000013
Private Const ColorFondoCalculo = &HFFFF&
Private Busqueda As Boolean
Private mobjBusq As Busqueda
Private mCodMoneda  As String
Private objcond As Condicion
Dim nombre As String
Private mRutaDestino103 As String
Private mRutaDestino104 As String

Public Sub Inicio(ByVal tag As String)
    Dim rutaPlantilla
    Dim i As Integer
    Dim valor As Currency
    On Error GoTo ErrTrap
    Me.tag = tag            'Guarda en la propiedad Tag para distinguir después
    Me.Show
    Me.ZOrder
    Select Case Me.tag
    Case "F104", "F1042010"
        Me.Caption = "Formulario 104 Declaración del Impuesto Valor Agregado"
        grd.Rows = grd.FixedRows
        ConfigCols1042010
        txtDestino.Text = mRutaDestino104
        
    Case "F103", "F1032010"
         Me.Caption = "Formulario 103 Declaración del Impuesto a la Renta"
         grd.Rows = grd.FixedRows
         txtDestino.Text = mRutaDestino103
         ConfigCols103
    End Select
    RecuperarConfig
    tlb1.Buttons(3).Enabled = False
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub cmdExplorar_Click()
    On Error GoTo ErrTrap
    With dlg1
        If Len(.filename) = 0 Then
            .InitDir = App.Path
        Else
            .InitDir = .filename
        End If
        .flags = cdlOFNPathMustExist
        .Filter = "Archivos xml (*.xml)|*.xml|Predefinido " & _
                  "|Todos (*.*)|*.*"
        .filename = nombre
        .ShowSave
        txtDestino.Text = .filename
        nombre = .FileTitle
        End With
    Exit Sub
ErrTrap:
    If Err.Number <> 32755 Then
        DispErr
    End If
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
        Select Case Me.tag
        Case "F104"
            Buscar104
        Case "F104"
            'Buscar
        End Select
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

Private Sub grd_AfterEdit(ByVal Row As Long, ByVal col As Long)
Select Case Me.tag
    Case "F1042010"
            Select Case col
                Case 2
                    grd.TextMatrix(32, 10) = Round(grd.ValueMatrix(32, 2) * gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA, 2)
                   ' grd.TextMatrix(32, 4) = Round(grd.ValueMatrix(22, 10), 2) - Round(grd.ValueMatrix(32, 2), 2) asi estaba
                    grd.TextMatrix(32, 4) = Round(grd.ValueMatrix(14, 10), 2) - Round(grd.ValueMatrix(32, 2), 2)
                    grd.TextMatrix(32, 12) = Round(grd.ValueMatrix(32, 6), 2) - Round(grd.ValueMatrix(32, 10), 2)
                Case 8
                    grd.TextMatrix(32, 14) = Round(grd.ValueMatrix(32, 8), 2) + Round(grd.ValueMatrix(32, 10), 2)
                    If Row = 39 Then
                        'grd.TextMatrix(44, 8) = Round(grd.ValueMatrix(36, 8) + grd.ValueMatrix(37, 8) + grd.ValueMatrix(38, 8) + grd.ValueMatrix(39, 8) + grd.ValueMatrix(40, 8) + grd.ValueMatrix(41, 8) + grd.ValueMatrix(42, 8), 2)
                        'grd.TextMatrix(44, 12) = Round(grd.ValueMatrix(36, 12) + grd.ValueMatrix(37, 12) + grd.ValueMatrix(38, 12) + grd.ValueMatrix(39, 12) + grd.ValueMatrix(40, 12), 12)
                          'grd.TextMatrix(51, 12) = Round((grd.ValueMatrix(36, 12) + grd.ValueMatrix(37, 12) + grd.ValueMatrix(39, 12) + grd.ValueMatrix(40, 12)) * grd.ValueMatrix(50, 12), 2)
                    End If
                Case 10
                        'grd.TextMatrix(39, 12) = Round(grd.ValueMatrix(39, 10) * gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA, 2)
                       ' grd.TextMatrix(44, 8) = Round(grd.ValueMatrix(36, 8) + grd.ValueMatrix(37, 8) + grd.ValueMatrix(38, 8) + grd.ValueMatrix(39, 8) + grd.ValueMatrix(40, 8) + grd.ValueMatrix(41, 8) + grd.ValueMatrix(42, 8), 2)
            End Select
        CalcularPorcentajes1042010
    Case "F1032010"
            CalcularPorcentajes1032010
    End Select
    
End Sub

Private Sub grd_BeforeEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
Dim i As Long, j As Long
    If Row < grd.FixedRows Then Cancel = True
    If grd.IsSubtotal(Row) = True Then Cancel = True
    If grd.ColData(col) < 0 Then Cancel = True
    Select Case Me.tag
    Case "F104"
        CalcularPorcentajes104
    Case "F103"
        CalcularPorcentajes103
    Case "F1042010"
        'If grd.Cell(flexcpBackColor, Row, col, Row, col) = vbRed Then
            grd.ColComboList(14) = "Original|Sustituta"
        'Else
           ' Cancel = True
        'End If
        'CalcularPorcentajes1042010
    End Select
    
    'If grd.CellBackColor = grd.BackColorFrozen Or grd.CellBackColor = &HC00000 Then
    If grd.CellBackColor = grd.BackColorFrozen Or grd.CellBackColor = &HC00000 Then
       Cancel = True
    End If
    If Me.tag = "F1042010" Then
        If grd.Cell(flexcpBackColor, Row, col, Row, col) <> vbWhite Then
            Cancel = True
        End If
    End If
    If Me.tag = "F1032010" Then
        If Row = 7 Then
            grd.ColComboList(12) = "Original|Sustituta"
            If grd.Cell(flexcpBackColor, Row, col, Row, col) <> vbWhite Then
                Cancel = True
            End If
        Else
            grd.ColComboList(12) = ""
        End If
    End If
End Sub

Private Sub grd_BeforeSort(ByVal col As Long, Order As Integer)
    'Impide mientras está procesando
    If mProcesando Then Order = flexSortNone
End Sub

Private Sub grd_ChangeEdit()
    'CalcularPorcentajes
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
    Case "Abrir": AbrirArchivo
    Case "Nuevo":
            nuevo
            tlb1.Buttons(3).Enabled = True
    Case "Buscar": Buscar
    Case "Imprimir": Imprimir
    Case "Excel": ExportaExcel (Me.tag)
    Case "XML": ExportaXML (Me.tag)
    Case "Guardar": GuardarResultado
    Case "Cerrar":      Cerrar
    End Select
End Sub

Private Sub ConfigCols104()
    Dim s As String, i As Long, j As Integer, s1 As String
    Dim fmt As String
    With grd
        s = "^#|<c1|<c2|^c3|^c4|>c5|^c6|>c7|^c8|>c9|^c10|^c11"
        .FormatString = s
        AsignarTituloAColKey grd
        .ColWidth(0) = 350
        .ColWidth(1) = 450
        .ColWidth(2) = 4200
        .ColWidth(3) = 450
        .ColWidth(4) = 450
        .ColWidth(5) = 2400
        .ColWidth(6) = 450
        .ColWidth(7) = 2400
        .ColWidth(8) = 450
        .ColWidth(9) = 2800
        .ColWidth(10) = 0
        .ColWidth(11) = 0
        'Columnas modificables (Longitud maxima)
        .ColFormat(.ColIndex("c5")) = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda(fmt)
        .ColFormat(.ColIndex("c9")) = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda(fmt)
        .ColFormat(.ColIndex("c7")) = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda(fmt)
        .MergeCol(3) = True
        .Refresh
    End With
End Sub
Private Sub ConfigCols1042010()
    Dim s As String, i As Long, j As Integer, s1 As String
    Dim fmt As String
    With grd
        s = "^#|<c1|<c2|<c3|<c4|<c5|<c6|<c7|>c8|<c9|>c10|<c11|>c12|<c13|>c14"
        .FormatString = s
        AsignarTituloAColKey grd
        .ColWidth(0) = 350
        .ColWidth(1) = 800
        .ColWidth(2) = 1500
        .ColWidth(3) = 800
        .ColWidth(4) = 1250
        .ColWidth(5) = 1500
        .ColWidth(6) = 1350
        .ColWidth(7) = 470
        .ColWidth(8) = 1400
        .ColWidth(9) = 470
        .ColWidth(10) = 1400
        .ColWidth(11) = 450
        .ColWidth(12) = 1400
        .ColWidth(13) = 400
        .ColWidth(14) = 1400
        'Columnas modificables (Longitud maxima)
        .ColFormat(2) = gobjMain.EmpresaActual.GNOpcion.FormatoCantidad
        .ColFormat(6) = gobjMain.EmpresaActual.GNOpcion.FormatoCantidad
        .ColFormat(8) = gobjMain.EmpresaActual.GNOpcion.FormatoCantidad
        .ColFormat(10) = gobjMain.EmpresaActual.GNOpcion.FormatoCantidad
        .ColFormat(12) = gobjMain.EmpresaActual.GNOpcion.FormatoCantidad
        .ColFormat(14) = gobjMain.EmpresaActual.GNOpcion.FormatoCantidad
        .ColDataType(8) = flexDTCurrency
        .ColDataType(10) = flexDTCurrency
        .ColDataType(12) = flexDTCurrency
        .ColDataType(7) = flexDTString
        .ColDataType(8) = flexDTString
        .ColDataType(9) = flexDTString
        .ColDataType(10) = flexDTString
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

Public Sub ExportaExcel(ByVal titulo As String)
    'tipo=0 Roles de Pagos; tipo=1 Reporte Bancos y Provisiones
    
    If grd.Rows < 2 Then
        MsgBox "No existe filas para exportar a Excel"
        Exit Sub
    End If
    If objcond Is Nothing Then
        MsgBox "No se realizó la busqueda para exportar a Excel"
        Exit Sub
    End If

    
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
        .Range("H1").Font.Size = 10
        .Range("H1").Font.Bold = True
        .Cells(fila, 1) = titulo
        
        .PageSetup.PaperSize = xlPaperLetter 'Tamaño del papel (carta)
        .PageSetup.BottomMargin = Application.CentimetersToPoints(1.5) 'Margen Superior
        .PageSetup.TopMargin = Application.CentimetersToPoints(1) 'Margen Inferior
'        .Range(.Cells(1, 13), .Cells(500, 23)).NumberFormat = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda(fmt)    'Establece el formato para los números
        .Range("A2:AZ100").Font.Name = "Arial"    'Tipo de letra para toda la hoja
        .Range("A2:AZ100").Font.Size = 7          'Tamaño de la letra
        
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
            Else
                j = 1
                mayor = 0

                For col = 1 To grd.Cols - 1
                
                    Select Case Me.tag
                        Case "F104"
                            Select Case col
                                'Case 1: mayor = 2
                                Case 1, 3, 4, 6, 8, 10, 11: mayor = 4
                                Case 2, 5, 7, 9: mayor = 25
                            End Select
                        Case "F103"
                            Select Case col
                                Case 2: mayor = 50
                                Case 3, 4, 6, 8, 9: mayor = 4
                                Case 5, 7, 10: mayor = 15
                                Case 1: mayor = 10
                            End Select
                       End Select
                
                        .Cells(fila, j) = grd.TextMatrix(i, col)
'                        mayor = Len(grd.TextMatrix(i, Col)) 'Para ajustar el ancho de columnas
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
    For i = 1 To grd.Rows - 1
'    If Not grd.IsSubtotal(i) Then
        For j = 1 To grd.Cols - 1
               Cadena = Cadena & grd.TextMatrix(i, j) & ","
        Next j
        Cadena = Mid(Cadena, 1, Len(Cadena) - 1)
        Print #NumFile, Cadena
        Cadena = ""
'     End If
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
    grd.MergeCells = flexMergeSpill
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
    Select Case Me.tag
    Case "F104"
        ConfigCols104
        CambiaFondoCeldasEditables104
    End Select
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



Private Sub LLENADATOS104()
    With grd
        .Redraw = flexRDBuffered
        .TextMatrix(10, 2) = gobjMain.EmpresaActual.GNOpcion.ruc
        .TextMatrix(10, 4) = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("RazonSocial")
        .TextMatrix(98, 2) = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("IDRepreLegal")
        .TextMatrix(98, 4) = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("RUCContdor")
         .Refresh
    End With

End Sub
Private Sub LLENADATOS1042010()
    With grd
        .Redraw = flexRDBuffered
        .TextMatrix(10, 2) = "'" & gobjMain.EmpresaActual.GNOpcion.ruc
        .TextMatrix(10, 4) = gobjMain.EmpresaActual.GNOpcion.RazonSocial
        .TextMatrix(103, 5) = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("IDRepreLegal")
        .TextMatrix(103, 9) = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("RUCContdor")
        .Refresh
    End With

End Sub


Private Sub CambiaFondoCeldasEditables104()
'    Dim fmt As String
'    With grd
'        .Redraw = flexRDBuffered
'        fondo blanco
'
'        .Cell(flexcpBackColor, 7, 9, 7, 9) = vbWhite ' formulario modifica
'
'        .Cell(flexcpBackColor, 10, 2, 10, 2) = vbWhite 'ruc
'        .Cell(flexcpBackColor, 10, 4, 10, 7) = vbWhite 'razon social
'        .Cell(flexcpBackColor, 12, 9, 15, 9) = vbWhite
'
'        .Cell(flexcpBackColor, 19, 5, 23, 5) = vbWhite
'        .Cell(flexcpBackColor, 25, 5, 31, 5) = vbWhite
'        .Cell(flexcpBackColor, 35, 5, 36, 5) = vbWhite
'        .Cell(flexcpBackColor, 19, 7, 23, 7) = vbWhite
'        .Cell(flexcpBackColor, 25, 7, 25, 7) = vbWhite
'        .Cell(flexcpBackColor, 30, 7, 31, 7) = vbWhite
'        .Cell(flexcpBackColor, 35, 7, 36, 7) = vbWhite
'        .Cell(flexcpBackColor, 37, 9, 37, 9) = vbWhite
'
'        .Cell(flexcpBackColor, 40, 2, 40, 2) = vbWhite
'        .Cell(flexcpBackColor, 40, 5, 40, 5) = vbWhite
'        .Cell(flexcpBackColor, 40, 7, 40, 7) = vbWhite
'        .Cell(flexcpBackColor, 40, 9, 40, 9) = vbWhite
'
'        .Cell(flexcpBackColor, 42, 2, 42, 2) = vbWhite
'        .Cell(flexcpBackColor, 42, 5, 42, 5) = vbWhite
'        .Cell(flexcpBackColor, 42, 9, 42, 9) = vbWhite
'
'        .Cell(flexcpBackColor, 44, 7, 44, 7) = vbWhite
'        .Cell(flexcpBackColor, 45, 5, 45, 5) = vbWhite
'
'        COMPRAS
'        .Cell(flexcpBackColor, 48, 5, 52, 5) = vbWhite
'        .Cell(flexcpBackColor, 55, 5, 56, 5) = vbWhite
'        .Cell(flexcpBackColor, 58, 5, 58, 5) = vbWhite
'        .Cell(flexcpBackColor, 61, 5, 64, 5) = vbWhite
'
'
'        .Cell(flexcpBackColor, 48, 7, 56, 7) = vbWhite
'        .Cell(flexcpBackColor, 59, 7, 59, 7) = vbWhite
'        .Cell(flexcpBackColor, 61, 7, 64, 7) = vbWhite
'
'        .Cell(flexcpBackColor, 48, 9, 56, 9) = vbWhite
'        .Cell(flexcpBackColor, 60, 9, 64, 9) = vbWhite
'
'        .Cell(flexcpBackColor, 66, 7, 66, 7) = vbWhite
'        .Cell(flexcpBackColor, 67, 9, 68, 9) = vbWhite
'
'        .Cell(flexcpBackColor, 70, 9, 75, 9) = vbWhite
'
'        .Cell(flexcpBackColor, 77, 7, 88, 7) = vbWhite
'        .Cell(flexcpBackColor, 77, 9, 88, 9) = vbWhite
'
'        .Cell(flexcpBackColor, 89, 5, 89, 5) = vbWhite
'
'        .Cell(flexcpBackColor, 90, 9, 91, 9) = vbWhite
'        .Cell(flexcpBackColor, 94, 9, 98, 9) = vbWhite
'
'        .Cell(flexcpBackColor, 100, 9, 102, 9) = vbWhite
'
''''        'fondo 1
'        .Cell(flexcpBackColor, 2, 4, 2, 9) = ColorFondo1
'        .Cell(flexcpBackColor, 8, 1, 8, 9) = ColorFondo1
'        .Cell(flexcpBackColor, 11, 1, 11, 9) = ColorFondo1
'        .Cell(flexcpBackColor, 17, 1, 17, 9) = ColorFondo1
'        .Cell(flexcpBackColor, 46, 1, 46, 9) = ColorFondo1
'        .Cell(flexcpBackColor, 69, 1, 69, 9) = ColorFondo1
'        .Cell(flexcpBackColor, 76, 1, 76, 9) = ColorFondo1
'        .Cell(flexcpBackColor, 92, 6, 92, 9) = ColorFondo1
'        .Cell(flexcpBackColor, 99, 6, 99, 9) = ColorFondo1
'        .Cell(flexcpBackColor, 103, 1, 103, 9) = ColorFondo1
'''' campo calculado
'        .Cell(flexcpBackColor, 16, 9, 16, 9) = ColorFondoCalculo
'        .Cell(flexcpBackColor, 19, 9, 32, 9) = ColorFondoCalculo
'        .Cell(flexcpBackColor, 24, 5, 24, 5) = ColorFondoCalculo
'        .Cell(flexcpBackColor, 32, 5, 32, 5) = ColorFondoCalculo
'        .Cell(flexcpBackColor, 33, 7, 33, 7) = ColorFondoCalculo
'        .Cell(flexcpBackColor, 34, 9, 36, 9) = ColorFondoCalculo
'        .Cell(flexcpBackColor, 40, 2, 40, 2) = ColorFondoCalculo
'        .Cell(flexcpBackColor, 40, 7, 40, 7) = ColorFondoCalculo
'        .Cell(flexcpBackColor, 42, 2, 42, 2) = ColorFondoCalculo
'        .Cell(flexcpBackColor, 42, 9, 42, 9) = ColorFondoCalculo
'
'        .Cell(flexcpBackColor, 48, 9, 63, 9) = ColorFondoCalculo
'        .Cell(flexcpBackColor, 58, 5, 58, 5) = ColorFondoCalculo
'
'
'        .Cell(flexcpFontBold, 42, 9, 42, 9) = True '499
'        .Cell(flexcpFontBold, 57, 1, 54, 2) = True '499
'
'
'
'
'
'        .Cell(flexcpFontSize, 16, 9, 16, 9) = 9
''        .Cell(flexcpFontSize, 19, 5, 36, 5) = 9
''        .Cell(flexcpFontSize, 19, 7, 36, 7) = 9
'        .Cell(flexcpFontSize, 19, 9, 36, 9) = 9
''        .Cell(flexcpFontSize, 32, 5, 32, 5) = 9
''        .Cell(flexcpFontSize, 33, 7, 33, 7) = 9
'        .Cell(flexcpFontSize, 34, 9, 36, 9) = 9
''        .Cell(flexcpFontSize, 40, 2, 40, 2) = 9
''        .Cell(flexcpFontSize, 40, 7, 40, 7) = 9
'        .Cell(flexcpFontSize, 40, 9, 40, 9) = 9
''        .Cell(flexcpFontSize, 42, 2, 42, 2) = 9
''        .Cell(flexcpFontSize, 42, 5, 42, 5) = 9
''        .Cell(flexcpFontSize, 42, 7, 42, 7) = 9
'        .Cell(flexcpFontSize, 42, 9, 42, 9) = 9
''        .Cell(flexcpFontSize, 44, 7, 44, 7) = 9
''        .Cell(flexcpFontSize, 45, 5, 45, 5) = 9
''        .Cell(flexcpFontSize, 57, 1, 57, 2) = 9
''        .Cell(flexcpFontSize, 48, 5, 63, 5) = 9
''        .Cell(flexcpFontSize, 48, 7, 63, 7) = 9
'        .Cell(flexcpFontSize, 48, 9, 63, 9) = 9
'
'
'
'
''''        'fondo
'        .Cell(flexcpBackColor, 4, 4, 4, 4) = ColorFondo
'        .Cell(flexcpBackColor, 4, 8, 4, 8) = ColorFondo
'        .Cell(flexcpBackColor, 6, 8, 6, 8) = ColorFondo
'        .Cell(flexcpBackColor, 9, 1, 10, 1) = ColorFondo
'        .Cell(flexcpBackColor, 9, 3, 10, 3) = ColorFondo
'        .Cell(flexcpBackColor, 12, 8, 16, 8) = ColorFondo
'        .Cell(flexcpBackColor, 19, 1, 32, 1) = ColorFondo
'
'        .Cell(flexcpBackColor, 19, 4, 36, 4) = ColorFondo
'        .Cell(flexcpBackColor, 19, 6, 36, 6) = ColorFondo
'        .Cell(flexcpBackColor, 19, 8, 37, 8) = ColorFondo
'        .Cell(flexcpBackColor, 24, 6, 24, 9) = ColorFondo
'        .Cell(flexcpBackColor, 26, 6, 29, 9) = ColorFondo
'        .Cell(flexcpBackColor, 32, 6, 32, 9) = ColorFondo
'        .Cell(flexcpBackColor, 33, 4, 34, 5) = ColorFondo
'        .Cell(flexcpBackColor, 33, 9, 33, 9) = ColorFondo
'        .Cell(flexcpBackColor, 34, 7, 34, 7) = ColorFondo
'        .Cell(flexcpBackColor, 40, 1, 40, 1) = ColorFondo
'        .Cell(flexcpBackColor, 40, 4, 40, 4) = ColorFondo
'        .Cell(flexcpBackColor, 40, 6, 40, 6) = ColorFondo
'        .Cell(flexcpBackColor, 40, 8, 40, 8) = ColorFondo
'        .Cell(flexcpBackColor, 42, 1, 42, 1) = ColorFondo
'        .Cell(flexcpBackColor, 42, 4, 42, 4) = ColorFondo
'        .Cell(flexcpBackColor, 42, 8, 42, 8) = ColorFondo
'
'        .Cell(flexcpBackColor, 44, 6, 44, 6) = ColorFondo
'        .Cell(flexcpBackColor, 45, 4, 45, 4) = ColorFondo
'
'        .Cell(flexcpBackColor, 48, 4, 63, 4) = ColorFondo
'        .Cell(flexcpBackColor, 48, 6, 63, 6) = ColorFondo
'        .Cell(flexcpBackColor, 48, 8, 63, 8) = ColorFondo
'
'        .Cell(flexcpBackColor, 53, 4, 54, 6) = ColorFondo
'        .Cell(flexcpBackColor, 57, 4, 57, 9) = ColorFondo
'        .Cell(flexcpBackColor, 58, 6, 58, 9) = ColorFondo
'        .Cell(flexcpBackColor, 59, 8, 59, 9) = ColorFondo
'        .Cell(flexcpBackColor, 59, 4, 59, 5) = ColorFondo
'        .Cell(flexcpBackColor, 60, 4, 59, 7) = ColorFondo
'        .Cell(flexcpBackColor, 66, 6, 66, 6) = ColorFondo
'        .Cell(flexcpBackColor, 67, 8, 68, 8) = ColorFondo
'
'        .Cell(flexcpBackColor, 70, 8, 75, 8) = ColorFondo
'        .Cell(flexcpBackColor, 77, 6, 88, 6) = ColorFondo
'        .Cell(flexcpBackColor, 77, 8, 88, 8) = ColorFondo
'        .Cell(flexcpBackColor, 89, 4, 89, 4) = ColorFondo
'
'        .Cell(flexcpBackColor, 90, 8, 91, 8) = ColorFondo
'        .Cell(flexcpBackColor, 94, 8, 98, 8) = ColorFondo
'        .Cell(flexcpBackColor, 100, 8, 102, 8) = ColorFondo
'        .Cell(flexcpBackColor, 99, 1, 99, 1) = ColorFondo
'        .Cell(flexcpBackColor, 99, 3, 99, 3) = ColorFondo
'
'        .Cell(flexcpBackColor, 104, 1, 107, 1) = ColorFondo
'        .Cell(flexcpBackColor, 104, 3, 107, 3) = ColorFondo
'        .Cell(flexcpBackColor, 104, 6, 105, 6) = ColorFondo
'        .Cell(flexcpBackColor, 104, 8, 105, 8) = ColorFondo
'        .Cell(flexcpBackColor, 59, 7, 59, 7) = ColorFondoCalculo
''''
''''        'alineacion
'        .Cell(flexcpAlignment, 1, 8, 1, 8) = 1
'        .Cell(flexcpAlignment, 2, 5, 2, 5) = 1
'        .Cell(flexcpAlignment, 4, 5, 7, 5) = 1
'        .Cell(flexcpAlignment, 4, 9, 7, 9) = 1
'
'        .Cell(flexcpAlignment, 39, 1, 39, 9) = 2
'        .Cell(flexcpAlignment, 40, 2, 40, 2) = 6
'        .Cell(flexcpAlignment, 42, 2, 42, 2) = 6
'        .Cell(flexcpAlignment, 41, 1, 41, 9) = 2
'        .Cell(flexcpAlignment, 43, 7, 43, 7) = 2
'        .Cell(flexcpAlignment, 93, 6, 97, 6) = 2
'
'
'
''''        'alto linea
'        .RowHeight(1) = 400
'        .RowHeight(2) = 320
'        .RowHeight(4) = 320
'        .RowHeight(5) = 320
'        .RowHeight(8) = 320
'        .RowHeight(11) = 320
'        .RowHeight(17) = 320
'        .RowHeight(46) = 320
'        .RowHeight(68) = 320
'        .RowHeight(75) = 320
'        .RowHeight(91) = 320
'        .RowHeight(102) = 320
''''
''''        'tamaño letras
'        .Cell(flexcpFontSize, 1, 1, 1, 6) = 14
'        .Cell(flexcpFontSize, 2, 1, 2, 8) = 10
'        .Cell(flexcpFontSize, 4, 2, 5, 3) = 11
'        .Cell(flexcpFontSize, 8, 1, 8, 8) = 10
'        .Cell(flexcpFontSize, 11, 1, 11, 8) = 10
'        .Cell(flexcpFontSize, 17, 1, 17, 8) = 10
'        .Cell(flexcpFontSize, 38, 1, 38, 8) = 9
'        .Cell(flexcpFontSize, 46, 1, 46, 8) = 10
'        .Cell(flexcpFontSize, 68, 1, 68, 8) = 10
'        .Cell(flexcpFontSize, 75, 1, 75, 5) = 10
'        .Cell(flexcpFontSize, 91, 6, 91, 8) = 10
'        .Cell(flexcpFontSize, 102, 1, 102, 8) = 10
''''
''''        'negritas
'        .Cell(flexcpFontBold, 2, 3, 2, 9) = True
'        .Cell(flexcpFontBold, 4, 2, 4, 2) = True
'        .Cell(flexcpFontBold, 8, 1, 8, 9) = True
'        .Cell(flexcpFontBold, 11, 1, 11, 9) = True
'        .Cell(flexcpFontBold, 17, 1, 17, 9) = True
'        .Cell(flexcpFontBold, 38, 1, 38, 9) = True
'        .Cell(flexcpFontBold, 46, 1, 46, 9) = True
'        .Cell(flexcpFontBold, 68, 1, 68, 9) = True
'        .Cell(flexcpFontBold, 75, 1, 75, 9) = True
'        .Cell(flexcpFontBold, 91, 6, 91, 9) = True
'        .Cell(flexcpFontBold, 102, 1, 102, 9) = True
'
'        .Cell(flexcpBackColor, 19, 3, 24, 3) = ColorFondo1
'        .Cell(flexcpBackColor, 28, 3, 29, 3) = ColorFondo1
'
'        .Cell(flexcpTextStyle, 40, 2, 40, 2) = "#,0.00"
'        .Cell(flexcpTextStyle, 42, 2, 42, 2) = "#,0.00"
'
'        .MergeCol(1) = True
'        .MergeCol(2) = True
'        .MergeCol(3) = True
'        .MergeCol(4) = True
'        .MergeCol(5) = True
'    End With
End Sub

Private Sub Buscar104()
    Dim sql As String, cond As String, CadenaValores As String
    Dim OrdenadoX As String
    Dim CadenaAgrupa  As String, Recargo As String
    Dim v As Variant, max As Integer, i As Integer
    Dim from As String, NumReg As Long, f1 As String
    Dim rs As Recordset
    Dim SubTotal As Currency, CompraBienesTarifa_0 As Currency, CompraServiciosTarifa_0 As Currency
    Dim CompraActivosTarifa_0 As Currency
    Dim CP_Ser As String, CP_Act As String, CP_Dev As String
    Dim VT_Bie As String, VT_Ser As String, VT_Dev As String
    Dim VT_ExpBie As String, VT_ExpSer As String, VT_RepGas As String
    Dim NC_Ventas As String, NC_Compras As String, CP_Bie As String
    Dim Reten As String, ret_real As String, ret_recib As String
    Dim Moneda As String
    Set objcond = gobjMain.objCondicion
    If Not frmB_FormSRI.Inicio104(objcond, Recargo, CP_Ser, CP_Act, CP_Dev, VT_Dev, VT_Bie, VT_Ser, _
                                                    VT_ExpBie, VT_ExpSer, VT_RepGas, ret_real, ret_recib, Reten, NC_Ventas, NC_Compras, CP_Bie) Then
        grd.SetFocus
        Exit Sub
    End If
    With objcond
        If Len(Month(.fecha1)) < 2 Then
            grd.TextMatrix(4, 6) = "0" & Month(.fecha1)
        Else
            grd.TextMatrix(4, 6) = Month(.fecha1)
        End If
        
        grd.TextMatrix(4, 9) = " AÑO " & Year(.fecha1)
        
        'Reporte de un mes a la vez
        f1 = DateSerial(Year(.fecha1), Month(.fecha1), 1)
        cond = " AND GNC.FechaTrans BETWEEN " & FechaYMD(f1, gobjMain.EmpresaActual.TipoDB) & _
               " AND " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB)
            
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
            sql = sql & " isnull(SUM(PrecioTotalBase0 *SignoVenta + (PrecioTotalBase0*SignoVenta * (cast(isnull(TotalDescuento,0) as float) *SignoVenta/ cast(PrecioTotal*SignoVenta as float)*SignoVenta))),0) as Base0, "
            sql = sql & " isnull(SUM(PrecioTotalBaseIVA * SignoVenta + (PrecioTotalBaseIVA *SignoVenta* (cast(isnull(TotalDescuento,0) as float) *SignoVenta/ cast(PrecioTotal as float)*SignoVenta))),0) As BaseIVA "
            sql = sql & " FROM tmp0 Right join "
            sql = sql & "vwConsSUMIVKardexIVA inner join "
            sql = sql & "GNComprobante GNC   "
            sql = sql & " INNER JOIN PCPROVCLI PC ON GNC.IDCLIENTEREF=PC.IDPROVCLI"
            sql = sql & " INNER JOIN GNTRANS gnt ON gnc.CODTRANS=GNT.CODTRANS"
            sql = sql & " ON vwConsSUMIVKardexIVA.TransID = GNC.TransID "
            sql = sql & " ON tmp0.TransID = GNC.TransID"
            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(VT_Bie) & ")  "
            sql = sql & " AND pc.BANDEMPRESAPUBLICA=0"
            sql = sql & " AND GNT.AnexoCodTipoTrans=2 " & cond
            
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            grd.TextMatrix(19, 5) = Round(Abs(rs.Fields("Base0")), 2)
            grd.TextMatrix(19, 7) = Round(Abs(rs.Fields("BaseIVA")), 2)
            
    '********** VENTAS BIENES    EMPRESAS PUBLICAS
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
            sql = sql & " isnull(SUM(PrecioTotalBase0 *SignoVenta + (PrecioTotalBase0*SignoVenta * (cast(isnull(TotalDescuento,0) as float) *SignoVenta/ cast(PrecioTotal*SignoVenta as float)*SignoVenta))),0) as Base0, "
            sql = sql & " isnull(SUM(PrecioTotalBaseIVA * SignoVenta + (PrecioTotalBaseIVA *SignoVenta* (cast(isnull(TotalDescuento,0) as float) *SignoVenta/ cast(PrecioTotal as float)*SignoVenta))),0) As BaseIVA "
            sql = sql & " FROM tmp0 Right join "
            sql = sql & "vwConsSUMIVKardexIVA inner join "
            sql = sql & "GNComprobante GNC   "
            sql = sql & " INNER JOIN PCPROVCLI PC ON GNC.IDCLIENTEREF=PC.IDPROVCLI"
            sql = sql & " INNER JOIN GNTRANS gnt ON gnc.CODTRANS=GNT.CODTRANS"
            sql = sql & " ON vwConsSUMIVKardexIVA.TransID = GNC.TransID "
            sql = sql & " ON tmp0.TransID = GNC.TransID"
            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(VT_Bie) & ")  "
            sql = sql & " AND pc.BANDEMPRESAPUBLICA=1"
            sql = sql & " AND GNT.AnexoCodTipoTrans=2 " & cond
            
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            grd.TextMatrix(28, 5) = Round(Abs(rs.Fields("Base0")), 2)
            'grd.TextMatrix(28, 7) = Round(Abs(rs.Fields("BaseIVA")), 2)
            
'********** NOTAS DE CREDITO
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
            sql = sql & " isnull(SUM(PrecioTotalBase0 *SignoVenta + (PrecioTotalBase0*SignoVenta * (cast(isnull(TotalDescuento,0) as float) *SignoVenta/ cast(PrecioTotal*SignoVenta as float)*SignoVenta))),0) as Base0, "
            sql = sql & " isnull(SUM(PrecioTotalBaseIVA * SignoVenta + (PrecioTotalBaseIVA *SignoVenta* (cast(isnull(TotalDescuento,0) as float) *SignoVenta/ cast(PrecioTotal as float)*SignoVenta))),0) As BaseIVA "
            sql = sql & " FROM tmp0 Right join "
            sql = sql & "vwConsSUMIVKardexIVA inner join "
            sql = sql & "GNComprobante GNC   "
            sql = sql & " INNER JOIN PCPROVCLI PC ON GNC.IDCLIENTEREF=PC.IDPROVCLI"
            sql = sql & " INNER JOIN GNTRANS gnt ON gnc.CODTRANS=GNT.CODTRANS"
            sql = sql & " ON vwConsSUMIVKardexIVA.TransID = GNC.TransID "
            sql = sql & " ON tmp0.TransID = GNC.TransID"
            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(NC_Ventas) & ")  "
            sql = sql & " AND pc.BANDEMPRESAPUBLICA=0"
            sql = sql & " AND GNT.AnexoCodTipoTrans=2 " & cond
            
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            grd.TextMatrix(22, 5) = Round(Abs(rs.Fields("Base0")), 2)
            grd.TextMatrix(22, 7) = Round(Abs(rs.Fields("BaseIVA")), 2)
            
            
'********** NOTAS DE CREDITO EMPRESA PUBLICA
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
            sql = sql & " isnull(SUM(PrecioTotalBase0 *SignoVenta + (PrecioTotalBase0*SignoVenta * (cast(isnull(TotalDescuento,0) as float) *SignoVenta/ cast(PrecioTotal*SignoVenta as float)*SignoVenta))),0) as Base0, "
            sql = sql & " isnull(SUM(PrecioTotalBaseIVA * SignoVenta + (PrecioTotalBaseIVA *SignoVenta* (cast(isnull(TotalDescuento,0) as float) *SignoVenta/ cast(PrecioTotal as float)*SignoVenta))),0) As BaseIVA "
            sql = sql & " FROM tmp0 Right join "
            sql = sql & "vwConsSUMIVKardexIVA inner join "
            sql = sql & "GNComprobante GNC   "
            sql = sql & " INNER JOIN PCPROVCLI PC ON GNC.IDCLIENTEREF=PC.IDPROVCLI"
            sql = sql & " INNER JOIN GNTRANS gnt ON gnc.CODTRANS=GNT.CODTRANS"
            sql = sql & " ON vwConsSUMIVKardexIVA.TransID = GNC.TransID "
            sql = sql & " ON tmp0.TransID = GNC.TransID"
            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(NC_Ventas) & ")  "
            sql = sql & " AND pc.BANDEMPRESAPUBLICA=1"
            sql = sql & " AND GNT.AnexoCodTipoTrans=2 " & cond
            
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            grd.TextMatrix(30, 5) = Round(Abs(rs.Fields("Base0")), 2)
            grd.TextMatrix(30, 7) = Round(Abs(rs.Fields("BaseIVA")), 2)
            
            
            
    '********** EXPORTACION BIENES
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
            sql = sql & " isnull(SUM(PrecioTotalBase0 *SignoVenta ),0) as Base0, "
            sql = sql & " isnull(SUM(PrecioTotalBaseIVA * SignoVenta) ,0) As BaseIVA "
            sql = sql & " FROM tmp0 Right join "
            sql = sql & " vwConsSUMIVKardexIVA inner join "
            sql = sql & " GNComprobante GNC   "
            sql = sql & " INNER JOIN GNTRANS gnt ON gnc.CODTRANS=GNT.CODTRANS"
            sql = sql & " ON vwConsSUMIVKardexIVA.TransID = GNC.TransID "
            sql = sql & " ON tmp0.TransID = GNC.TransID"
            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(VT_ExpBie) & ")  "
            sql = sql & " AND GNT.AnexoCodTipoTrans=4 " & cond
            
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
                grd.TextMatrix(26, 5) = Round(Abs(rs.Fields("BaseIVA")), 2) + Round(Abs(rs.Fields("Base0")), 2)
            
'********** EXPORTACION SERVICIOS
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
            sql = sql & " isnull(SUM(PrecioTotalBase0 *SignoVenta ),0) as Base0, "
            sql = sql & " isnull(SUM(PrecioTotalBaseIVA * SignoVenta ),0) As BaseIVA "
            sql = sql & " FROM tmp0 Right join "
            sql = sql & " vwConsSUMIVKardexIVA inner join "
            sql = sql & " GNComprobante GNC   "
            sql = sql & " INNER JOIN GNTRANS gnt ON gnc.CODTRANS=GNT.CODTRANS"
            sql = sql & " ON vwConsSUMIVKardexIVA.TransID = GNC.TransID "
            sql = sql & " ON tmp0.TransID = GNC.TransID"
            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(VT_ExpSer) & ")  "
            sql = sql & " AND GNT.AnexoCodTipoTrans=4 " & cond
            
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            grd.TextMatrix(27, 5) = Round(Abs(rs.Fields("BaseIVA")), 2) + Round(Abs(rs.Fields("Base0")), 2)
            
            
'********** reposicion Gastos
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
            sql = sql & " isnull(SUM(PrecioTotalBase0 *SignoVenta + (PrecioTotalBase0*SignoVenta * (cast(isnull(TotalDescuento,0) as float) *SignoVenta/ cast(PrecioTotal*SignoVenta as float)*SignoVenta))),0) as Base0, "
            sql = sql & " isnull(SUM(PrecioTotalBaseIVA * SignoVenta + (PrecioTotalBaseIVA *SignoVenta* (cast(isnull(TotalDescuento,0) as float) *SignoVenta/ cast(PrecioTotal as float)*SignoVenta))),0) As BaseIVA "
            sql = sql & " FROM tmp0 Right join "
            sql = sql & " vwConsSUMIVKardexIVA inner join "
            sql = sql & " GNComprobante GNC   "
            sql = sql & " INNER JOIN GNTRANS gnt ON gnc.CODTRANS=GNT.CODTRANS"
            sql = sql & " ON vwConsSUMIVKardexIVA.TransID = GNC.TransID "
            sql = sql & " ON tmp0.TransID = GNC.TransID"
            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(VT_RepGas) & ")  "
            sql = sql & " AND GNT.AnexoCodTipoTrans=4 " & cond
            
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            grd.TextMatrix(36, 5) = Round(Abs(rs.Fields("Base0")), 2)
            grd.TextMatrix(36, 7) = Round(Abs(rs.Fields("BaseIVA")), 2)
            
            
            
            
    '********** VENTAS ACtivos
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
            sql = sql & " isnull(SUM(PrecioTotalBase0 *SignoVenta - (PrecioTotalBase0 * (cast(isnull(TotalDescuento,0) as float) / cast(PrecioTotal as float)))),0) as Base0, "
            sql = sql & " isnull(SUM(PrecioTotalBaseIVA *SignoVenta - (PrecioTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(PrecioTotal as float)))),0) As BaseIVA "
            sql = sql & " FROM tmp0 Right join "
            sql = sql & " vwConsSUMIVKardexIVA inner join "
            sql = sql & " GNComprobante GNC   "
            sql = sql & " INNER JOIN PCPROVCLI PC ON GNC.IDCLIENTEREF=PC.IDPROVCLI"
            sql = sql & " INNER JOIN GNTRANS gnt ON gnc.CODTRANS=GNT.CODTRANS"
            sql = sql & " ON vwConsSUMIVKardexIVA.TransID = GNC.TransID "
            sql = sql & " ON tmp0.TransID = GNC.TransID"
            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(VT_Ser) & ")"
            sql = sql & " AND pc.BANDEMPRESAPUBLICA=0"
            sql = sql & " AND GNT.AnexoCodTipoTrans=2 " & cond
            
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            grd.TextMatrix(20, 5) = Round(Abs(rs.Fields("Base0")), 2)
            grd.TextMatrix(20, 7) = Round(Abs(rs.Fields("BaseIVA")), 2)
            
'********** VENTAS ACtivos EMPRESAS PUBLICAS
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
            sql = sql & " isnull(SUM(PrecioTotalBase0 *SignoVenta - (PrecioTotalBase0 * (cast(isnull(TotalDescuento,0) as float) / cast(PrecioTotal as float)))),0) as Base0, "
            sql = sql & " isnull(SUM(PrecioTotalBaseIVA *SignoVenta - (PrecioTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(PrecioTotal as float)))),0) As BaseIVA "
            sql = sql & " FROM tmp0 Right join "
            sql = sql & " vwConsSUMIVKardexIVA inner join "
            sql = sql & " GNComprobante GNC   "
            sql = sql & " INNER JOIN PCPROVCLI PC ON GNC.IDCLIENTEREF=PC.IDPROVCLI"
            sql = sql & " INNER JOIN GNTRANS gnt ON gnc.CODTRANS=GNT.CODTRANS"
            sql = sql & " ON vwConsSUMIVKardexIVA.TransID = GNC.TransID "
            sql = sql & " ON tmp0.TransID = GNC.TransID"
            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(VT_Ser) & ")"
            sql = sql & " AND pc.BANDEMPRESAPUBLICA=1"
            sql = sql & " AND GNT.AnexoCodTipoTrans=2 " & cond
            
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            grd.TextMatrix(29, 5) = Round(Abs(rs.Fields("Base0")), 2) + Round(Abs(rs.Fields("BaseIVA")), 2)
            
            
            ' cantidad de comprobantes
            sql = "SELECT "
            sql = sql & "  count(gnc.codtrans) as NumComp  "
            sql = sql & " FROM  gncomprobante gnc inner join gntrans gnt "
            sql = sql & " on gnc.codtrans=gnt.codtrans"
            sql = sql & " WHERE ( GNC.CodTrans IN (" & PreparaCadena(VT_Bie) & " ) "
            sql = sql & " or GNC.CodTrans IN (" & PreparaCadena(VT_Ser) & "))  " & cond
            'sql = sql & "group by AnexoCodTipoComp"
            
            
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            
            If rs.RecordCount > 0 Then
                grd.TextMatrix(45, 5) = rs.Fields("NumComp")
            Else
                grd.TextMatrix(45, 5) = "0"
            End If
            
'*************************
'***COMPRAS
'*****************************
            
            
    '********** compras BIENES

            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1

            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(Recargo) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(.CodTrans) & ")" & _
                  "Group by IvkR.TransID"

            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg


            '--datos de la compra bienes tarifa 12
            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0, "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBaseIVA  + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS Valor12 "
            sql = sql & " Into tmp1"
            sql = sql & " from    (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(.CodTrans) & ")"
            sql = sql & " and  ane.CodCredTrib not in ('02','07')"
            sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg

            sql = " Select  isnull(sum(Valor0),0) as ValorTotal0, isnull(sum(Valor12),0) as ValorTotal12 from tmp1  "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            grd.TextMatrix(48, 5) = Round(rs.Fields("ValorTotal0"), 2)
            grd.TextMatrix(48, 7) = Round(rs.Fields("ValorTotal12"), 2)



        '--datos de la compra activos tarifa 12
            
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1

            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(Recargo) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(CP_Act) & ")" & _
                  "Group by IvkR.TransID"

            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
            
            
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
            sql = sql & " and  ane.CodCredTrib not in ('02','07')"
            sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1

            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
            sql = " Select  isnull(sum(Valor0),0) as ValorTotal0, isnull(sum(Valor12),0) as ValorTotal12 from tmp1  "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            grd.TextMatrix(49, 5) = Round(rs.Fields("ValorTotal0"), 2)
            grd.TextMatrix(49, 7) = Round(rs.Fields("ValorTotal12"), 2)



            '--datos de la compra servicios tarifa 12
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1

            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(Recargo) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(CP_Ser) & ")" & _
                  "Group by IvkR.TransID"

            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
            
            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0  + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0, "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBaseIVA + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS Valor12 "
            sql = sql & " Into tmp1"
            sql = sql & " from    (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(CP_Ser) & ")"
            sql = sql & " and  ane.CodCredTrib not in ('02','07')"
            sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
            sql = " Select  isnull(sum(Valor0),0) as ValorTotal0, isnull(sum(Valor12),0) as ValorTotal12 from tmp1  "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            grd.TextMatrix(50, 5) = Round(rs.Fields("ValorTotal0"), 2)
            grd.TextMatrix(50, 7) = Round(rs.Fields("ValorTotal12"), 2)


    '********** compras IMprtaciones

            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1

            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(Recargo) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(CP_Bie) & ")" & _
                  "Group by IvkR.TransID"

            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg


            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0, "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBaseIVA  + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS Valor12 "
            sql = sql & " Into tmp1"
            sql = sql & " from    (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " left join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(CP_Bie) & ")"
'            sql = sql & " and  ane.CodCredTrib not in ('02','07')"
            sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg

            sql = " Select  isnull(sum(Valor0),0) as ValorTotal0, isnull(sum(Valor12),0) as ValorTotal12 from tmp1  "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            grd.TextMatrix(51, 5) = Round(rs.Fields("ValorTotal0"), 2)
            grd.TextMatrix(51, 7) = Round(rs.Fields("ValorTotal12"), 2)

    '********** compras notas de credito

            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1

            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(Recargo) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(NC_Compras) & ")" & _
                  "Group by IvkR.TransID"

            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg


            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0, "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBaseIVA  + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS Valor12 "
            sql = sql & " Into tmp1"
            sql = sql & " from    (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(NC_Compras) & ")"
            sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg

            sql = " Select  isnull(sum(Valor0),0) as ValorTotal0, isnull(sum(Valor12),0) as ValorTotal12 from tmp1  "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            grd.TextMatrix(55, 5) = Round(Abs(rs.Fields("ValorTotal0")), 2)
            grd.TextMatrix(55, 7) = Round(Abs(rs.Fields("ValorTotal12")), 2)








        '--datos de la total compra que no sustentan cretito tributario sustento 02y 07
            
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1

            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 and "
            sql = sql & " (GNC.CodTrans IN (" & PreparaCadena(.CodTrans) & ")"
            sql = sql & " or  GNC.CodTrans IN (" & PreparaCadena(CP_Act) & ")"
            sql = sql & " or  GNC.CodTrans IN (" & PreparaCadena(CP_Ser) & "))"
            sql = sql & " AND ivr.CodRecargo IN (" & PreparaCadena(Recargo) & ") " & cond
            sql = sql & "Group by IvkR.TransID"

            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
            
            
            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0, "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBaseIVA + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS Valor12 "
            sql = sql & " Into tmp1"
            sql = sql & " from    (( tmp0 Right join gncomprobante Gnc inner join Anexos  on anexos.transid=gnc.transid "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  (GNC.CodTrans IN (" & PreparaCadena(.CodTrans) & ")"
            sql = sql & " or  GNC.CodTrans IN (" & PreparaCadena(CP_Act) & ")"
            sql = sql & " or  GNC.CodTrans IN (" & PreparaCadena(CP_Ser) & "))"
            sql = sql & " and  anexos.CodCredTrib in ('02','07')"
            sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1

            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
            sql = " Select  isnull(sum(Valor0),0) as ValorTotal0, isnull(sum(Valor12),0) as ValorTotal12 from tmp1  "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            grd.TextMatrix(63, 5) = Round(rs.Fields("ValorTotal0"), 2)
            grd.TextMatrix(63, 7) = Round(rs.Fields("ValorTotal12"), 2)
'''''
'''''' cantidad de comprobantes compras
'''''            sql = "SELECT "
'''''            sql = sql & " CodTipoComp, count(gnc.codtrans) as NumComp  "
'''''            sql = sql & " FROM  gncomprobante gnc inner join anexos ane  on gnc.transid=ane.transid "
'''''            sql = sql & " WHERE estado<> 3and ( GNC.CodTrans IN (" & PreparaCadena(.CodTrans) & " ) "
'''''            sql = sql & " or GNC.CodTrans IN (" & PreparaCadena(CP_Ser) & ")  "
'''''            sql = sql & " or GNC.CodTrans IN (" & PreparaCadena(CP_Act) & "))  " & cond
'''''            sql = sql & " group by CodTipoComp"
'''''
'''''            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
'''''
'''''
'''''            If rs.RecordCount > 0 Then 'rs.MoveFirst
'''''                While Not rs.EOF
'''''                    Select Case rs.Fields("CodTipoComp")
'''''                        Case "1":
'''''                            grd.TextMatrix(46, 4) = rs.Fields("NumComp")
'''''                        Case "01":
'''''                            grd.TextMatrix(46, 4) = grd.ValueMatrix(46, 4) + rs.Fields("NumComp")
'''''                        Case "2":
'''''                            grd.TextMatrix(47, 4) = rs.Fields("NumComp")
'''''                        Case "02"
'''''                            grd.TextMatrix(47, 4) = grd.ValueMatrix(47, 4) + rs.Fields("NumComp")
'''''                        Case "3":
'''''                            grd.TextMatrix(48, 4) = rs.Fields("NumComp")
'''''                        Case "03"
'''''                            grd.TextMatrix(48, 4) = grd.ValueMatrix(48, 4) + rs.Fields("NumComp")
'''''                        Case "4":
'''''                            grd.TextMatrix(47, 11) = rs.Fields("NumComp")
'''''                        Case "04"
'''''                            grd.TextMatrix(47, 11) = grd.ValueMatrix(47, 11) + rs.Fields("NumComp")
'''''                        Case "5":
'''''                            grd.TextMatrix(48, 11) = rs.Fields("NumComp")
'''''                        Case "05"
'''''                            grd.TextMatrix(48, 11) = grd.ValueMatrix(48, 11) + rs.Fields("NumComp")
'''''                        Case Else
'''''                            grd.TextMatrix(47, 7) = grd.ValueMatrix(47, 7) + rs.Fields("NumComp")
'''''                    End Select
'''''                    rs.MoveNext
'''''                Wend
'''''            Else
'''''                grd.TextMatrix(46, 4) = "0"
'''''                grd.TextMatrix(46, 4) = "0"
'''''                grd.TextMatrix(47, 4) = "0"
'''''                grd.TextMatrix(47, 4) = "0"
'''''                grd.TextMatrix(48, 4) = "0"
'''''                grd.TextMatrix(48, 4) = "0"
'''''                grd.TextMatrix(47, 11) = "0"
'''''                grd.TextMatrix(47, 11) = "0"
'''''                grd.TextMatrix(48, 11) = "0"
'''''                grd.TextMatrix(48, 11) = "0"
'''''                grd.TextMatrix(47, 7) = "0"
'''''            End If
'''''
'''''
'''''
'''''
'''''
'''''            ' RETENCIONES RECIBIDAS
'''''            .CodMoneda = MONEDA_PRE
'''''            Moneda = IIf(.NumMoneda > 0, "/Cotizacion" & .NumMoneda + 1, "")
'''''            sql = "SELECT ISNULL(sum(Valor" & Moneda & "),0) as TotalRetRecibidas "
'''''            sql = sql & " FROM vwConsRetencion "
'''''            sql = sql & " WHERE CodTrans IN (" & PreparaCadena(ret_recib) & ")"
'''''            sql = sql & "  AND FechaTrans BETWEEN " & FechaYMD(f1, gobjMain.EmpresaActual.TipoDB)
'''''            sql = sql & "  AND " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB)
'''''            sql = sql & "  and  DEBE > 0 and bandIVA=1"
'''''
'''''            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
'''''
'''''            grd.TextMatrix(53, 9) = Round(rs.Fields("TotalRetRecibidas"), 2)
''''''''
'''''            ' RETENCIONES REALIZADAS
'''''            sql = "SELECT CodF104," '
'''''            sql = sql & " ISNULL(sum(Base),0) as TotalBase "
'''''            sql = sql & " FROM vwConsRetencion "
'''''            If Len(ret_real) > 0 Then
'''''                sql = sql & " WHERE CodTrans IN (" & PreparaCadena(ret_real) & ") AND "
'''''            Else
'''''                sql = sql & " WHERE "
'''''            End If
'''''            sql = sql & "  FechaTrans BETWEEN " & FechaYMD(f1, gobjMain.EmpresaActual.TipoDB)
'''''            sql = sql & "  AND " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB)
'''''            sql = sql & "  and  HABER > 0 and bandIVA=1 "
'''''            sql = sql & "  group by Codf104" ', Porcentaje"
'''''
'''''            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
'''''
'''''            If rs.RecordCount > 0 Then 'rs.MoveFirst
'''''                While Not rs.EOF
'''''                    Select Case rs.Fields("CodF104")
'''''                        Case "801":
'''''                            grd.TextMatrix(57, 7) = Round(rs.Fields("TotalBase"), 2)
'''''                        Case "803":
'''''                            grd.TextMatrix(58, 7) = Round(rs.Fields("TotalBase"), 2)
'''''                        Case "805":
'''''                            grd.TextMatrix(59, 7) = Round(rs.Fields("TotalBase"), 2)
'''''                        Case "807":
'''''                            grd.TextMatrix(60, 7) = Round(rs.Fields("TotalBase"), 2)
'''''                        Case "809":
'''''                            grd.TextMatrix(61, 7) = Round(rs.Fields("TotalBase"), 2)
'''''                        Case "811":
'''''                            grd.TextMatrix(62, 7) = Round(rs.Fields("TotalBase"), 2)
'''''                        Case "813":
'''''                            grd.TextMatrix(63, 7) = Round(rs.Fields("TotalBase"), 2)
'''''                        Case "815":
'''''                            grd.TextMatrix(64, 7) = Round(rs.Fields("TotalBase"), 2)
'''''                        Case "817":
'''''                            grd.TextMatrix(65, 7) = Round(rs.Fields("TotalBase"), 2)
'''''                        Case "819":
'''''                            grd.TextMatrix(66, 7) = Round(rs.Fields("TotalBase"), 2)
'''''                        Case "821":
'''''                            grd.TextMatrix(67, 7) = Round(rs.Fields("TotalBase"), 2)
'''''                    End Select
'''''                    rs.MoveNext
'''''                Wend
'''''            Else
'''''                grd.TextMatrix(57, 7) = "0"
'''''                grd.TextMatrix(58, 7) = "0"
'''''                grd.TextMatrix(59, 7) = "0"
'''''                grd.TextMatrix(60, 7) = "0"
'''''                grd.TextMatrix(61, 7) = "0"
'''''                grd.TextMatrix(62, 7) = "0"
'''''                grd.TextMatrix(63, 7) = "0"
'''''                grd.TextMatrix(64, 7) = "0"
'''''                grd.TextMatrix(65, 7) = "0"
'''''                grd.TextMatrix(66, 7) = "0"
'''''                grd.TextMatrix(67, 7) = "0"
'''''            End If
'''''            'calcula numero de retenciones
'''''            sql = "SELECT count(Trans) as numTrans "
'''''            sql = sql & " FROM vwConsRetencion "
'''''            sql = sql & " WHERE CodTrans IN (" & PreparaCadena(ret_real) & ")"
'''''            sql = sql & "  AND FechaTrans BETWEEN " & FechaYMD(f1, gobjMain.EmpresaActual.TipoDB)
'''''            sql = sql & "  AND " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB)
'''''            sql = sql & "  and  HABER > 0 and bandIVA=1 "
'''''
'''''            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
'''''
'''''            grd.TextMatrix(69, 5) = rs.Fields("numTrans")
            
            CalcularPorcentajes104
            
            Select Case Me.tag
            Case "F103"
                nombre = "103_" & Format(CStr(Month(objcond.fecha1)), "00") & "_" & Year(objcond.fecha1) & ".XML"
                txtDestino.Text = mRutaDestino103 & "103_" & Format(CStr(Month(objcond.fecha1)), "00") & "_" & Year(objcond.fecha1) & ".XML"
            Case "F104"
                nombre = "104_" & Format(CStr(Month(objcond.fecha1)), "00") & "_" & Year(objcond.fecha1) & ".XML"
                txtDestino.Text = mRutaDestino104 & "104_" & Format(CStr(Month(objcond.fecha1)), "00") & "_" & Year(objcond.fecha1) & ".XML"
            End Select
            
            grd.Refresh
    End With
End Sub
    

Private Sub CalcularPorcentajes104()
    Dim i As Integer, SubtotalCompras As Currency, SubtotalVentas As Currency, TotalVentas As Currency
    Dim TotalVentasA As Currency, TotalVentasB As Currency
    Dim SubtotalRetenido As Currency, TotalCompras As Currency
    Dim TotalCompras0 As Currency, TotalImpCompras As Currency

    'mes anterior credito tributario 399=303-305
''''
    grd.TextMatrix(16, 9) = grd.ValueMatrix(13, 9) - grd.ValueMatrix(14, 9) + grd.ValueMatrix(15, 9)
''''
''''    'mes anterior credito tributario 703=399
    grd.TextMatrix(71, 9) = grd.ValueMatrix(13, 5)

    '406=401+402+403-404-405
    grd.TextMatrix(24, 5) = Round(grd.ValueMatrix(19, 5) + grd.ValueMatrix(20, 5) + grd.ValueMatrix(21, 5) - grd.ValueMatrix(22, 5) - grd.ValueMatrix(23, 5), 2)
    
    TotalVentasB = 0
    For i = 25 To 29
            TotalVentasB = TotalVentasB + grd.ValueMatrix(i, 5)
    Next i
    TotalVentasB = TotalVentasB - grd.ValueMatrix(30, 5) - grd.ValueMatrix(31, 5)
    grd.TextMatrix(32, 5) = Round(TotalVentasB, 2)
    '414
    TotalVentasB = 0
    For i = 25 To 29
            TotalVentasB = TotalVentasB + grd.ValueMatrix(i, 5)
    Next i
    TotalVentasB = TotalVentasB - grd.ValueMatrix(30, 5) - grd.ValueMatrix(31, 5)
    
    grd.TextMatrix(32, 5) = Round(TotalVentasB, 2)
    
    
''''    'ventas 549
''''
    TotalVentas = 0
''    For i = 19 To 32
''            TotalVentas = TotalVentas + grd.ValueMatrix(i, 7)
''    Next i

''''    TotalVentas = TotalVentas + grd.ValueMatrix(25, 5) + grd.ValueMatrix(26, 5)

''''
 SubtotalVentas = 0
''''    'ventas 599
    For i = 19 To 31
        If Len(grd.TextMatrix(i, 7)) > 0 Then
            grd.TextMatrix(i, 9) = Round(grd.ValueMatrix(i, 7) * gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA, 2)
            Select Case i
            Case 22, 23, 30, 31
                TotalVentas = TotalVentas - grd.ValueMatrix(i, 7)
                SubtotalVentas = SubtotalVentas - grd.ValueMatrix(i, 9)
            Case Else
                TotalVentas = TotalVentas + grd.ValueMatrix(i, 7)
                SubtotalVentas = SubtotalVentas + grd.ValueMatrix(i, 9)
            End Select
        Else
            grd.TextMatrix(i, 9) = ""
        End If
    Next i
    grd.TextMatrix(33, 7) = Round(TotalVentas, 2)
    grd.TextMatrix(34, 9) = Round(SubtotalVentas, 2)
    For i = 35 To 36
        If Len(grd.TextMatrix(i, 7)) > 0 Then
            grd.TextMatrix(i, 9) = Round(grd.ValueMatrix(i, 7) * gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA, 2)
        Else
            grd.TextMatrix(i, 9) = ""
        End If
    Next i
    
    
    '480
    grd.TextMatrix(40, 2) = Round(TotalVentas, 2)
    '482=454
    grd.TextMatrix(40, 7) = Round(SubtotalVentas, 2)
    '484
    grd.TextMatrix(42, 2) = Round(SubtotalVentas, 2)
    '485=482-484
    grd.TextMatrix(42, 5) = grd.ValueMatrix(40, 7) - grd.ValueMatrix(42, 2)
    '499=457+483+484
    grd.TextMatrix(42, 9) = Round(SubtotalVentas, 2) + grd.ValueMatrix(37, 9) + grd.ValueMatrix(40, 9)
    
    '460=406+414
    grd.TextMatrix(44, 7) = grd.ValueMatrix(24, 5) + grd.ValueMatrix(32, 5)
        
    
    
    
    
''''
    'calculo 301= (434+414+408+409+407)/(414+406+434+408+409)
    If (grd.ValueMatrix(32, 5) + grd.ValueMatrix(24, 5) + grd.ValueMatrix(33, 7) + grd.ValueMatrix(26, 5) + grd.ValueMatrix(27, 5)) <> 0 Then
        grd.TextMatrix(12, 9) = Round((grd.ValueMatrix(33, 7) + grd.ValueMatrix(32, 5) + grd.ValueMatrix(26, 5) + grd.ValueMatrix(27, 5) + grd.ValueMatrix(25, 5)) / (grd.ValueMatrix(32, 5) + grd.ValueMatrix(24, 5) + grd.ValueMatrix(33, 7) + grd.ValueMatrix(26, 5) + grd.ValueMatrix(27, 5)), 4)
    End If
'''''''''
''''

'******** compras
TotalCompras0 = 0
    For i = 48 To 52
            TotalCompras0 = TotalCompras0 + grd.ValueMatrix(i, 5)
    Next i
    grd.TextMatrix(58, 5) = TotalCompras0 - grd.ValueMatrix(55, 5) - grd.ValueMatrix(56, 5)





''''    'compras 699
    For i = 48 To 56
        If Len(grd.TextMatrix(i, 7)) > 0 Then
            Select Case i
            Case 55, 56
                grd.TextMatrix(i, 9) = Round(Abs(grd.ValueMatrix(i, 7)) * gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA, 2)
                SubtotalCompras = SubtotalCompras - grd.ValueMatrix(i, 7)
                TotalImpCompras = TotalImpCompras - grd.ValueMatrix(i, 9)
            Case Else
                grd.TextMatrix(i, 9) = Round(Abs(grd.ValueMatrix(i, 7)) * gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA, 2)
                SubtotalCompras = SubtotalCompras + grd.ValueMatrix(i, 7)
                TotalImpCompras = TotalImpCompras + grd.ValueMatrix(i, 9)
            End Select

        Else
            grd.TextMatrix(i, 9) = ""
        End If
    Next i
    '530
     grd.TextMatrix(59, 7) = Round(SubtotalCompras, 2)
     '550
     grd.TextMatrix(60, 9) = Round(TotalImpCompras, 2)
     '553
     For i = 61 To 63
        If Len(grd.TextMatrix(i, 7)) > 0 Then
            grd.TextMatrix(i, 9) = Round(Abs(grd.ValueMatrix(i, 7)) * gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA, 2)
        End If
    Next i
     
     '560
     grd.TextMatrix(65, 7) = 0
     
     
     If objcond.IncluirCero Then
        '554
        grd.TextMatrix(66, 9) = 0
        If grd.ValueMatrix(12, 9) <> 0 Then
            grd.TextMatrix(67, 9) = (grd.ValueMatrix(60, 9) + grd.ValueMatrix(61, 9) + grd.ValueMatrix(62, 9) + grd.ValueMatrix(63, 9)) / grd.ValueMatrix(12, 9)
        End If
     Else
        grd.TextMatrix(66, 9) = (grd.ValueMatrix(60, 9) + grd.ValueMatrix(61, 9) + grd.ValueMatrix(62, 9) + grd.ValueMatrix(63, 9))
        grd.TextMatrix(67, 9) = 0
     End If
     
''''    ' 601=499-554 ó 555>0
''''
    If (grd.ValueMatrix(42, 9) - grd.ValueMatrix(66, 9)) > 0 Then
        '601
        grd.TextMatrix(69, 9) = (grd.ValueMatrix(42, 9) - grd.ValueMatrix(66, 9))
        '602
        grd.TextMatrix(70, 9) = 0
    Else
        '601
        grd.TextMatrix(69, 9) = 0
        '602
        grd.TextMatrix(70, 9) = Abs((grd.ValueMatrix(42, 9) - grd.ValueMatrix(66, 9)))
    End If
''''    '798= 701-702-703-704<0
''''    If (grd.ValueMatrix(50, 9) - grd.ValueMatrix(51, 9) - grd.ValueMatrix(52, 9) - grd.ValueMatrix(53, 9)) < 0 Then
''''        '798
''''        grd.TextMatrix(54, 9) = Abs(grd.ValueMatrix(50, 9) - grd.ValueMatrix(51, 9) - grd.ValueMatrix(52, 9) - grd.ValueMatrix(53, 9))
''''        '799
''''        grd.TextMatrix(55, 9) = 0
''''    Else
''''        '798
''''        grd.TextMatrix(54, 9) = 0
''''        '799
''''        grd.TextMatrix(55, 9) = Abs(grd.ValueMatrix(50, 9) - grd.ValueMatrix(51, 9) - grd.ValueMatrix(52, 9) - grd.ValueMatrix(53, 9))
''''
''''    End If
''''    'retencion 100% 851-861
''''    For i = 57 To 62
''''        If Len(grd.TextMatrix(i, 7)) > 0 Then
''''            grd.TextMatrix(i, 9) = grd.ValueMatrix(i, 7)
''''        Else
''''            grd.TextMatrix(i, 9) = ""
''''        End If
''''    Next i
''''    '863    =813* 0.7
''''    If Len(grd.TextMatrix(63, 7)) > 0 Then
''''        grd.TextMatrix(63, 9) = Round(grd.ValueMatrix(63, 7) * 0.7, 2)
''''    Else
''''        grd.TextMatrix(63, 9) = ""
''''    End If
''''    '865    =865
''''    If Len(grd.TextMatrix(64, 7)) > 0 Then
''''        grd.TextMatrix(64, 9) = Round(grd.ValueMatrix(64, 7) * 0.7, 2)
''''    Else
''''        grd.TextMatrix(64, 9) = ""
''''    End If
''''
''''    For i = 65 To 67
''''        If Len(grd.TextMatrix(i, 7)) > 0 Then
''''            grd.TextMatrix(i, 9) = Round(grd.ValueMatrix(i, 7) * 0.3, 2)
''''        Else
''''            grd.TextMatrix(i, 9) = ""
''''        End If
''''    Next i
''''
'''''    '898=851+853+855+857+859+861+863+865+867+869
''''    For i = 57 To 67
''''        If Len(grd.TextMatrix(i, 9)) > 0 Then
''''            SubtotalRetenido = SubtotalRetenido + grd.ValueMatrix(i, 9)
''''        End If
''''    Next i
''''    grd.TextMatrix(68, 9) = SubtotalRetenido
''''
''''    '899=799+898
''''    grd.TextMatrix(69, 9) = grd.ValueMatrix(55, 9) + grd.ValueMatrix(68, 9)
''''
''''    '902=899-901
''''    grd.TextMatrix(73, 9) = grd.ValueMatrix(69, 9) - grd.ValueMatrix(72, 9)
''''
''''    '999=902+903+904
''''    grd.TextMatrix(76, 9) = grd.ValueMatrix(73, 9) + grd.ValueMatrix(74, 9) + grd.ValueMatrix(75, 9)
''''
''''
End Sub
Private Sub CalcularPorcentajes1042010()
    Dim i As Integer, SubtotalCompras As Currency, SubtotalVentas As Currency, TotalVentas As Currency
    Dim TotalVentasA As Currency, TotalVentasB As Currency
    Dim SubtotalRetenido As Currency, TotalCompras As Currency
    Dim TotalCompras0 As Currency, TotalImpCompras As Currency
    grd.TextMatrix(32, 6) = grd.TextMatrix(22, 12)
    '409 = 401+402+403+404+405+406+407+408
    grd.TextMatrix(22, 8) = Round(grd.ValueMatrix(14, 8) + grd.ValueMatrix(15, 8) + grd.ValueMatrix(16, 8) + grd.ValueMatrix(17, 8) + grd.ValueMatrix(18, 8) + grd.ValueMatrix(19, 8) + grd.ValueMatrix(20, 8) + grd.ValueMatrix(21, 8), 2)
    grd.TextMatrix(22, 10) = Round(grd.ValueMatrix(14, 10) + grd.ValueMatrix(15, 10) + grd.ValueMatrix(16, 10) + grd.ValueMatrix(17, 10) + grd.ValueMatrix(18, 10) + grd.ValueMatrix(19, 10) + grd.ValueMatrix(20, 10) + grd.ValueMatrix(21, 10), 2)
    grd.TextMatrix(14, 12) = Round(grd.ValueMatrix(14, 10) * gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA, 2)
    grd.TextMatrix(15, 12) = Round(grd.ValueMatrix(15, 10) * gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA, 2)
    grd.TextMatrix(22, 12) = Round(grd.ValueMatrix(14, 12) + grd.ValueMatrix(15, 12), 2)
    TotalVentasB = 0
    For i = 25 To 29
            TotalVentasB = TotalVentasB + grd.ValueMatrix(i, 5)
    Next i
    TotalVentasB = TotalVentasB - grd.ValueMatrix(30, 5) - grd.ValueMatrix(31, 5)
    '414
    TotalVentasB = 0
    For i = 25 To 29
            TotalVentasB = TotalVentasB + grd.ValueMatrix(i, 5)
    Next i
    TotalVentasB = TotalVentasB - grd.ValueMatrix(30, 5) - grd.ValueMatrix(31, 5)
    TotalVentas = 0
    SubtotalVentas = 0
    
    '482=429
    grd.TextMatrix(32, 6) = grd.TextMatrix(22, 12)
    
    If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "WAY") > 0 Then
        grd.TextMatrix(38, 12) = Round(Abs(grd.ValueMatrix(38, 10)) * gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA, 2)
    Else
        grd.TextMatrix(36, 12) = Round(Abs(grd.ValueMatrix(36, 10)) * gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA, 2)
    End If
    grd.TextMatrix(37, 12) = Round(Abs(grd.ValueMatrix(37, 10)) * gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA, 2)
    grd.TextMatrix(38, 12) = Round(Abs(grd.ValueMatrix(38, 10)) * gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA, 2)
    grd.TextMatrix(39, 12) = Round(Abs(grd.ValueMatrix(39, 10)) * gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA, 2)
    grd.TextMatrix(40, 12) = Round(Abs(grd.ValueMatrix(40, 10)) * gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA, 2)
    grd.TextMatrix(41, 12) = Round(Abs(grd.ValueMatrix(41, 10)) * gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA, 2)
    grd.TextMatrix(32, 2) = Round(grd.ValueMatrix(14, 10) + grd.ValueMatrix(15, 10), 1)
    '484
    grd.TextMatrix(32, 10) = Round(grd.ValueMatrix(32, 2) * gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA, 2)
    '499
    grd.TextMatrix(32, 14) = Round(grd.ValueMatrix(32, 8) + grd.ValueMatrix(32, 10), 2)
    grd.TextMatrix(45, 8) = Round(grd.ValueMatrix(36, 8) + grd.ValueMatrix(37, 8) + grd.ValueMatrix(38, 8) + grd.ValueMatrix(39, 8) + grd.ValueMatrix(40, 8) + grd.ValueMatrix(41, 8) + grd.ValueMatrix(42, 8) + grd.ValueMatrix(43, 8) + grd.ValueMatrix(44, 8), 2)
    grd.TextMatrix(45, 10) = Round(grd.ValueMatrix(36, 10) + grd.ValueMatrix(37, 10) + grd.ValueMatrix(38, 10) + grd.ValueMatrix(39, 10) + grd.ValueMatrix(40, 10) + grd.ValueMatrix(41, 10) + grd.ValueMatrix(42, 10) + grd.ValueMatrix(43, 10) + grd.ValueMatrix(44, 10), 2)
    grd.TextMatrix(45, 12) = Round(grd.ValueMatrix(36, 12) + grd.ValueMatrix(37, 12) + grd.ValueMatrix(38, 12) + grd.ValueMatrix(39, 12) + grd.ValueMatrix(40, 12) + grd.ValueMatrix(41, 12), 2)
    
    '******** compras
    TotalCompras0 = 0
    For i = 48 To 53
            TotalCompras0 = TotalCompras0 + grd.ValueMatrix(i, 5)
    Next i
    
    
    If grd.ValueMatrix(22, 10) <> 0 Then
         grd.TextMatrix(52, 12) = Format((grd.ValueMatrix(14, 10) + grd.ValueMatrix(15, 10) + grd.ValueMatrix(18, 10) + grd.ValueMatrix(19, 10) + grd.ValueMatrix(20, 10) + grd.ValueMatrix(21, 10)) / grd.ValueMatrix(22, 10), "#,0.0000")
     Else
         grd.TextMatrix(52, 12) = 0
     End If
    If objcond.IncluirCero Then
        grd.TextMatrix(53, 12) = Round((grd.ValueMatrix(36, 12) + grd.ValueMatrix(37, 12) + grd.ValueMatrix(39, 12) + grd.ValueMatrix(40, 12) + grd.ValueMatrix(41, 12)) * grd.ValueMatrix(52, 12), 2)
    Else
        grd.TextMatrix(53, 12) = Round((grd.ValueMatrix(36, 12) + grd.ValueMatrix(37, 12) + grd.ValueMatrix(39, 12) + grd.ValueMatrix(40, 12) + grd.ValueMatrix(41, 12)) * grd.ValueMatrix(52, 12), 2)
    End If
    
    If (grd.ValueMatrix(32, 14) - grd.ValueMatrix(53, 12)) > 0 Then
        '601
        grd.TextMatrix(56, 12) = (grd.ValueMatrix(32, 14) - grd.ValueMatrix(53, 12))
        '602
        grd.TextMatrix(57, 12) = 0
    Else
        '602
        grd.TextMatrix(56, 12) = 0
        grd.TextMatrix(57, 12) = Abs((grd.ValueMatrix(32, 14) - grd.ValueMatrix(53, 12)))
    End If
    '619
    
    
'''     If objcond.IncluirCero Then
'''        '554
'''        grd.TextMatrix(66, 9) = 0
'''        If grd.ValueMatrix(12, 9) <> 0 Then
'''            grd.TextMatrix(67, 9) = (grd.ValueMatrix(60, 9) + grd.ValueMatrix(61, 9) + grd.ValueMatrix(62, 9) + grd.ValueMatrix(63, 9)) / grd.ValueMatrix(12, 9)
'''        End If
'''     Else
'''        grd.TextMatrix(66, 9) = (grd.ValueMatrix(60, 9) + grd.ValueMatrix(61, 9) + grd.ValueMatrix(62, 9) + grd.ValueMatrix(63, 9))
'''        grd.TextMatrix(67, 9) = 0
'''     End If
    
    
    If objcond.IncluirCero Then
    Else
        If (grd.ValueMatrix(56, 12) - grd.ValueMatrix(57, 12) - grd.ValueMatrix(58, 12) - grd.ValueMatrix(60, 12) - grd.ValueMatrix(60, 12) + grd.ValueMatrix(62, 12)) > 0 Then
            grd.TextMatrix(66, 12) = Round((grd.ValueMatrix(56, 12) - grd.ValueMatrix(57, 12) - grd.ValueMatrix(58, 12) - grd.ValueMatrix(60, 12) - grd.ValueMatrix(61, 12) + grd.ValueMatrix(62, 12)), 2)
        Else
            grd.TextMatrix(66, 12) = 0
        End If
    End If
    grd.TextMatrix(68, 12) = Round(grd.ValueMatrix(66, 12) + grd.ValueMatrix(67, 12), 2)
    grd.TextMatrix(75, 12) = Round(grd.ValueMatrix(72, 12) + grd.ValueMatrix(73, 12) + grd.ValueMatrix(74, 12), 2)
    grd.TextMatrix(77, 12) = Round(grd.ValueMatrix(68, 12) + grd.ValueMatrix(75, 12), 2)
    grd.TextMatrix(84, 12) = Round(grd.ValueMatrix(77, 12) + grd.ValueMatrix(81, 3), 2)
    grd.TextMatrix(87, 12) = Round(grd.ValueMatrix(84, 12) + grd.ValueMatrix(85, 12) + grd.ValueMatrix(86, 12), 2)
    grd.TextMatrix(89, 12) = grd.TextMatrix(87, 12)
End Sub

Public Sub nuevo()
    grd.Rows = grd.FixedRows
    Select Case Me.tag
    Case "F104"
'        ConfigCols104
'        LlenaFormatoFormulario104
'        LLENADATOS104
'        CambiaFondoCeldasEditables104
    Case "F103"
'        ConfigCols103
'        LlenaFormatoFormulario103
'        LLENADATOS103
'        CambiaFondoCeldasEditables103
    Case "F1032010"
        ConfigCols1032010
        LlenaFormatoFormulario103_2010
        LLENADATOS1032010
        CambiaFondoCeldasEditables103_2010
    Case "F1042010"
        ConfigCols1042010
        LlenaFormatoFormulario104_2010
        CambiaFondoCeldasEditables104_2010
    LLENADATOS1042010 'ENCABEZADO
    End Select
End Sub

Public Sub Buscar()
    Select Case Me.tag
    Case "F104"
        Buscar104
    Case "F103"
        Buscar103
    Case "F1032010"
        Buscar1032010
    Case "F1042010"
        Buscar1042010
    End Select
End Sub

Public Sub ExportaXML(ByVal titulo As String)
    Dim v As Variant, file As String
    Dim NumFile As Integer, Cadena As String
    Dim Filas As Long, Columnas As Long, i As Long, j As Long
    On Error GoTo ErrTrap
    
    If grd.Rows < 2 Then
        MsgBox "No existe filas para exportar a formato XML"
        Exit Sub
    End If
    If objcond Is Nothing Then
        MsgBox "No se realizó la busqueda para exportar a formato XML"
        Exit Sub
    End If
    Select Case Me.tag
    Case "F103", "F1032010"
        If grd.TextMatrix(7, 12) = "Original" Then
            nombre = "03ORI_" & MonthName(DatePart("M", objcond.fecha1)) & Year(objcond.fecha1) & ".XML"
        Else
            nombre = "03SUS_" & MonthName(DatePart("M", objcond.fecha1)) & Year(objcond.fecha1) & ".XML"
        End If
    Case "F104", "F1042010"
    
        nombre = "104_" & Format(CStr(Month(objcond.fecha1)), "00") & "_" & Year(objcond.fecha1) & ".XML"
    End Select

    file = txtDestino.Text '& nombre
    If ExisteArchivo(file) Then
        If MsgBox("El nombre del archivo " & nombre & " ya existe desea sobreescribirlo?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    NumFile = FreeFile
    Open file For Output Access Write As #NumFile
    Select Case Me.tag
    Case "F103", "F1032010"
        Cadena = GeneraArchivo103
    Case "F104", "F1042010"
        Cadena = GeneraArchivo104
    End Select
      
   Print #NumFile, Cadena

    Close NumFile
    MsgBox "Exportación Satisfactoria"
    GuardarConfig
    Exit Sub
ErrTrap:
    grd.TextMatrix(grd.Rows - 1, 2) = Err.Description
    Close NumFile
End Sub

Private Function GeneraArchivo104() As String
Dim mes As Integer
    Dim obj As GNOpcion, cad As String
    'Set obj = mobjMain.RecuperarGnOpcion.Recuperar
'    cad = "<?xml version=" & """1.0""" & " encoding=""" & "UTF-8" & stalone = "yes""""?>"
    cad = "<?xml version=" & """1.0""" & " encoding=" & """UTF-8""" & " standalone =" & """yes""" & "?>"
    cad = cad & "<formulario version=" & """0.2""" & ">"
    cad = cad & "<cabecera>"
      cad = cad & "<codigo_version_formulario>04201401</codigo_version_formulario>"
    cad = cad & "<ruc>" & gobjMain.EmpresaActual.GNOpcion.ruc & "</ruc>"
    cad = cad & "<codigo_moneda>1</codigo_moneda>"
    cad = cad & "</cabecera>"
    cad = cad & "<detalle>"
    Select Case grd.TextMatrix(7, 2)
        Case "Enero", "enero": mes = 1
        Case "Febrero", "febrero": mes = 2
        Case "Marzo", "marzo": mes = 3
        Case "Abril", "abril": mes = 4
        Case "Mayo", "mayo": mes = 5
        Case "Junio", "junio": mes = 6
        Case "Julio", "julio": mes = 7
        Case "Agosto", "agosto": mes = 8
        Case "Septiembre", "septiembre": mes = 9
        Case "Octubre", "octubre": mes = 10
        Case "Noviembre", "noviembre": mes = 11
        Case "Diciembre", "diciembre": mes = 12
    End Select
    
    'cad = cad & "<campo numero=""" & "31" & """>" & grd.TextMatrix(7, 14) & "</campo>"
    If grd.TextMatrix(7, 14) = "Original" Then
        cad = cad & "<campo numero=""" & "31" & """>" & "O" & "</campo>"
    ElseIf grd.TextMatrix(7, 14) = "Sustituta" Then
        cad = cad & "<campo numero=""" & "31" & """>" & "S" & "</campo>"
    Else
        MsgBox "Debe escoger el tipo de formulario [Original - Sustituto]"
        Exit Function
    End If
    cad = cad & "<campo numero=""" & "101" & """>" & mes & "</campo>"
    cad = cad & "<campo numero=""" & "102" & """>" & grd.TextMatrix(7, 4) & "</campo>"
    If grd.TextMatrix(7, 14) = "Original" Then
        cad = cad & "<campo numero=""" & "104" & """></campo>"
    ElseIf grd.TextMatrix(7, 14) = "Sustituta" Then
        cad = cad & "<campo numero=""" & "104" & """>" & grd.TextMatrix(8, 13) & "   </campo>"
    End If
    cad = cad & "<campo numero=""" & "198" & """> " & grd.TextMatrix(99, 5) & " </campo>"
    cad = cad & "<campo numero=""" & "199" & """> " & grd.TextMatrix(99, 9) & " </campo>"
    cad = cad & "<campo numero=""" & "201" & """>" & gobjMain.EmpresaActual.GNOpcion.ruc & "</campo>"
    cad = cad & "<campo numero=""" & "202" & """>" & "<![CDATA[" & gobjMain.EmpresaActual.GNOpcion.NombreEmpresa & "]]>" & "</campo>"
    
    cad = cad & "<campo numero=""" & "401" & """>" & Format(grd.ValueMatrix(14, 8), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "402" & """>" & Format(grd.ValueMatrix(15, 8), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "403" & """>" & Format(grd.ValueMatrix(16, 8), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "404" & """>" & Format(grd.ValueMatrix(17, 8), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "405" & """>" & Format(grd.ValueMatrix(18, 8), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "406" & """>" & Format(grd.ValueMatrix(19, 8), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "407" & """>" & Format(grd.ValueMatrix(20, 8), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "408" & """>" & Format(grd.ValueMatrix(21, 8), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "409" & """>" & Format(grd.ValueMatrix(22, 8), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "411" & """>" & Format(grd.ValueMatrix(14, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "412" & """>" & Format(grd.ValueMatrix(15, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "413" & """>" & Format(grd.ValueMatrix(16, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "414" & """>" & Format(grd.ValueMatrix(17, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "415" & """>" & Format(grd.ValueMatrix(18, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "416" & """>" & Format(grd.ValueMatrix(19, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "417" & """>" & Format(grd.ValueMatrix(20, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "418" & """>" & Format(grd.ValueMatrix(21, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "419" & """>" & Format(grd.ValueMatrix(22, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "421" & """>" & Format(grd.ValueMatrix(14, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "422" & """>" & Format(grd.ValueMatrix(15, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "429" & """>" & Format(grd.ValueMatrix(22, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "431" & """>" & Format(grd.ValueMatrix(23, 8), "#.00") & "</campo>"
    
    'cad = cad & "<campo numero=""" & "432" & """>" & Format(grd.ValueMatrix(24, 10), "#.00") & "</campo>"
    'cad = cad & "<campo numero=""" & "433" & """>" & Format(grd.ValueMatrix(25, 10), "#.00") & "</campo>"
    
    cad = cad & "<campo numero=""" & "434" & """>" & Format(grd.ValueMatrix(26, 8), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "441" & """>" & Format(grd.ValueMatrix(23, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "442" & """>" & Format(grd.ValueMatrix(24, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "443" & """>" & Format(grd.ValueMatrix(25, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "444" & """>" & Format(grd.ValueMatrix(26, 10), "#.00") & "</campo>"
    
    cad = cad & "<campo numero=""" & "453" & """>" & Format(grd.ValueMatrix(25, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "454" & """>" & Format(grd.ValueMatrix(26, 12), "#.00") & "</campo>"
    
    cad = cad & "<campo numero=""" & "480" & """>" & Format(grd.ValueMatrix(32, 2), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "481" & """>" & Format(grd.ValueMatrix(32, 4), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "482" & """>" & Format(grd.ValueMatrix(32, 6), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "483" & """>" & Format(grd.ValueMatrix(32, 8), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "484" & """>" & Format(grd.ValueMatrix(32, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "485" & """>" & Format(grd.ValueMatrix(32, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "499" & """>" & Format(grd.ValueMatrix(32, 14), "#.00") & "</campo>"
    'COMPRAS
    cad = cad & "<campo numero=""" & "500" & """>" & Format(grd.ValueMatrix(36, 8), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "501" & """>" & Format(grd.ValueMatrix(37, 8), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "502" & """>" & Format(grd.ValueMatrix(38, 8), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "503" & """>" & Format(grd.ValueMatrix(39, 8), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "504" & """>" & Format(grd.ValueMatrix(40, 8), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "505" & """>" & Format(grd.ValueMatrix(41, 8), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "506" & """>" & Format(grd.ValueMatrix(42, 8), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "507" & """>" & Format(grd.ValueMatrix(43, 8), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "508" & """>" & Format(grd.ValueMatrix(44, 8), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "509" & """>" & Format(grd.ValueMatrix(45, 8), "#.00") & "</campo>"
    'cad = cad & "<campo numero=""" & "510" & """>" & Format(grd.ValueMatrix(36, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "510" & """>" & Format(grd.ValueMatrix(36, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "511" & """>" & Format(grd.ValueMatrix(37, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "512" & """>" & Format(grd.ValueMatrix(38, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "513" & """>" & Format(grd.ValueMatrix(39, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "514" & """>" & Format(grd.ValueMatrix(40, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "515" & """>" & Format(grd.ValueMatrix(41, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "516" & """>" & Format(grd.ValueMatrix(42, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "517" & """>" & Format(grd.ValueMatrix(43, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "518" & """>" & Format(grd.ValueMatrix(44, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "519" & """>" & Format(grd.ValueMatrix(45, 10), "#.00") & "</campo>"
    
    cad = cad & "<campo numero=""" & "520" & """>" & Format(grd.ValueMatrix(36, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "521" & """>" & Format(grd.ValueMatrix(37, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "522" & """>" & Format(grd.ValueMatrix(38, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "523" & """>" & Format(grd.ValueMatrix(39, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "524" & """>" & Format(grd.ValueMatrix(40, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "525" & """>" & Format(grd.ValueMatrix(41, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "529" & """>" & Format(grd.ValueMatrix(45, 12), "#.00") & "</campo>"
    
    cad = cad & "<campo numero=""" & "531" & """>" & Format(grd.ValueMatrix(46, 8), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "532" & """>" & Format(grd.ValueMatrix(47, 8), "#.00") & "</campo>"
    'cad = cad & "<campo numero=""" & "533" & """>" & Format(grd.ValueMatrix(47, 10), "#.00") & "</campo>"
    'cad = cad & "<campo numero=""" & "534" & """>" & Format(grd.ValueMatrix(48, 10), "#.00") & "</campo>"
    
    cad = cad & "<campo numero=""" & "535" & """>" & Format(grd.ValueMatrix(50, 8), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "541" & """>" & Format(grd.ValueMatrix(46, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "542" & """>" & Format(grd.ValueMatrix(47, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "543" & """>" & Format(grd.ValueMatrix(48, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "544" & """>" & Format(grd.ValueMatrix(49, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "545" & """>" & Format(grd.ValueMatrix(50, 10), "#.00") & "</campo>"
    
    cad = cad & "<campo numero=""" & "554" & """>" & Format(grd.ValueMatrix(49, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "555" & """>" & Format(grd.ValueMatrix(50, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "563" & """>" & Format(grd.ValueMatrix(52, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "564" & """>" & Format(grd.ValueMatrix(53, 12), "#.00") & "</campo>"
    'RESUMEN
    cad = cad & "<campo numero=""" & "601" & """>" & Format(grd.ValueMatrix(56, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "602" & """>" & Format(grd.ValueMatrix(57, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "605" & """>" & Format(grd.ValueMatrix(58, 12), "#.00") & "</campo>"
    
    cad = cad & "<campo numero=""" & "607" & """>" & Format(grd.ValueMatrix(60, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "609" & """>" & Format(grd.ValueMatrix(61, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "611" & """>" & Format(grd.ValueMatrix(62, 12), "#.00") & "</campo>"
    
    cad = cad & "<campo numero=""" & "613   " & """>" & Format(grd.ValueMatrix(63, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "615" & """>" & Format(grd.ValueMatrix(64, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "617" & """>" & Format(grd.ValueMatrix(65, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "619" & """>" & Format(grd.ValueMatrix(66, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "621" & """>" & Format(grd.ValueMatrix(67, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "699" & """>" & Format(grd.ValueMatrix(68, 12), "#.00") & "</campo>"
    'AGENTE
    cad = cad & "<campo numero=""" & "721" & """>" & Format(grd.ValueMatrix(72, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "723" & """>" & Format(grd.ValueMatrix(73, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "725" & """>" & Format(grd.ValueMatrix(74, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "799" & """>" & Format(grd.ValueMatrix(75, 12), "#.00") & "</campo>"
    
    cad = cad & "<campo numero=""" & "859" & """>" & Format(grd.ValueMatrix(77, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "890" & """>" & Format(grd.ValueMatrix(79, 12), "#.00") & "</campo>"
    
    cad = cad & "<campo numero=""" & "897" & """>" & Format(grd.ValueMatrix(81, 3), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "898" & """>" & Format(grd.ValueMatrix(81, 6), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "899" & """>" & Format(grd.ValueMatrix(81, 12), "#.00") & "</campo>"
    
    'VALORES A PAGAR
    cad = cad & "<campo numero=""" & "880" & """>" & Format(grd.ValueMatrix(82, 12), "#.00") & "</campo>"
    
    cad = cad & "<campo numero=""" & "902" & """>" & Format(grd.ValueMatrix(84, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "903" & """>" & Format(grd.ValueMatrix(85, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "904" & """>" & Format(grd.ValueMatrix(86, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "905" & """>" & Format(grd.ValueMatrix(89, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "906" & """>" & Format(grd.ValueMatrix(90, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "907" & """>" & Format(grd.ValueMatrix(91, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "908" & """>" & Format(grd.TextMatrix(94, 3), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "909" & """>" & Format(grd.ValueMatrix(95, 3), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "910" & """>" & Format(grd.TextMatrix(94, 5), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "911" & """>" & Format(grd.ValueMatrix(95, 5), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "912" & """>" & Format(grd.TextMatrix(94, 8), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "913" & """>" & Format(grd.ValueMatrix(95, 8), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "915" & """>" & Format(grd.ValueMatrix(95, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "916" & """>" & Format(grd.TextMatrix(97, 6), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "917" & """>" & Format(grd.ValueMatrix(98, 6), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "918" & """>" & Format(grd.TextMatrix(97, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "919" & """>" & Format(grd.ValueMatrix(98, 10), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "922" & """>" & "" & "</campo>"
    cad = cad & "<campo numero=""" & "999" & """>" & Format(grd.ValueMatrix(87, 12), "#.00") & "</campo>"
    cad = cad & "</detalle>"
    cad = cad & "</formulario>"
    GeneraArchivo104 = cad
End Function

Private Sub GuardarConfig()
    Dim pos As Integer, s As String
    
    pos = InStrRev(txtDestino.Text, "\")
    If pos > 0 Then
        s = Mid$(txtDestino.Text, 1, pos)
    Else
        s = ""
    End If
    Select Case Me.tag
    Case "F104", "F1042010"
        SaveSetting APPNAME, App.Title, Me.Name & ".RutaDestino104", s
    Case "F103", "F1032010"
        SaveSetting APPNAME, App.Title, Me.Name & ".RutaDestino103", s
    End Select
    
End Sub

Private Sub RecuperarConfig()
    Select Case Me.tag
    Case "F104"
        mRutaDestino104 = GetSetting(APPNAME, App.Title, Me.Name & ".RutaDestino104", App.Path)
    Case "F103", "F1032010"
         mRutaDestino103 = GetSetting(APPNAME, App.Title, Me.Name & ".RutaDestino103", App.Path)
    Case "F1042010"
        mRutaDestino104 = GetSetting(APPNAME, App.Title, Me.Name & ".RutaDestino104", App.Path)
    End Select
End Sub

Private Sub ConfigCols103()
    Dim s As String, i As Long, j As Integer, s1 As String
    Dim fmt As String
    With grd
        s = "^#|<c1|<c2|^c3|^c4|>c5|^c6|>c7|^c8|^c9|>c10"
        .FormatString = s
        'AjustarAutoSize grd, -1, -1, 4000
        AsignarTituloAColKey grd
        .ColWidth(0) = 350
        .ColWidth(1) = 1000
        .ColWidth(2) = 5000
        .ColWidth(3) = 450
        .ColWidth(4) = 450
        .ColWidth(5) = 2300
        .ColWidth(6) = 450
        .ColWidth(7) = 2100
        .ColWidth(8) = 450
        .ColWidth(9) = 450
        .ColWidth(10) = 2000
        '.ColWidth(11) = 650
        
        'grilla de resultados
        'Columnas modificables (Longitud maxima)
        .ColFormat(.ColIndex("c5")) = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda(fmt)
        .ColFormat(.ColIndex("c7")) = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda(fmt)
        .ColFormat(.ColIndex("c10")) = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda(fmt)
        .MergeCol(3) = True
        .Refresh
    End With
End Sub

Private Sub CambiaFondoCeldasEditables103()
    With grd
        .Redraw = flexRDBuffered
        'fondo blanco
        .Cell(flexcpBackColor, 4, 6, 4, 6) = vbWhite ' formulario modifica
        .Cell(flexcpBackColor, 4, 10, 4, 10) = vbWhite ' formulario modifica
        .Cell(flexcpBackColor, 5, 10, 5, 10) = vbWhite ' formulario modifica
        .Cell(flexcpBackColor, 8, 2, 8, 2) = vbWhite 'ruc
        .Cell(flexcpBackColor, 8, 4, 8, 9) = vbWhite 'razon social
        .Cell(flexcpBackColor, 10, 7, 41, 7) = vbWhite
        .Cell(flexcpBackColor, 43, 4, 45, 4) = vbWhite
        .Cell(flexcpBackColor, 45, 7, 45, 7) = vbWhite
        .Cell(flexcpBackColor, 47, 7, 58, 7) = vbWhite
        .Cell(flexcpBackColor, 62, 10, 62, 10) = vbWhite
        .Cell(flexcpBackColor, 64, 10, 65, 10) = vbWhite
        .Cell(flexcpBackColor, 67, 10, 68, 10) = vbWhite
        .Cell(flexcpBackColor, 65, 2, 66, 2) = vbWhite
        .Cell(flexcpBackColor, 65, 4, 66, 5) = vbWhite
        .Cell(flexcpBackColor, 70, 2, 73, 2) = vbWhite
        .Cell(flexcpBackColor, 70, 5, 73, 5) = vbWhite
        
        'fondo 1
        .Cell(flexcpBackColor, 2, 4, 2, 9) = ColorFondo1
        .Cell(flexcpBackColor, 6, 1, 6, 9) = ColorFondo1
        .Cell(flexcpBackColor, 9, 1, 9, 9) = ColorFondo1
        .Cell(flexcpBackColor, 46, 1, 46, 9) = ColorFondo1
        .Cell(flexcpBackColor, 61, 6, 61, 9) = ColorFondo1
        'fondo
        .Cell(flexcpBackColor, 4, 4, 4, 4) = ColorFondo
        .Cell(flexcpBackColor, 5, 6, 5, 6) = ColorFondo
        .Cell(flexcpBackColor, 4, 8, 4, 8) = ColorFondo
        .Cell(flexcpBackColor, 6, 4, 6, 4) = ColorFondo
        .Cell(flexcpBackColor, 7, 1, 8, 1) = ColorFondo
        .Cell(flexcpBackColor, 7, 3, 8, 3) = ColorFondo
        .Cell(flexcpBackColor, 10, 1, 14, 1) = ColorFondo
        .Cell(flexcpBackColor, 10, 8, 10, 10) = ColorFondo
        .Cell(flexcpBackColor, 11, 8, 11, 8) = ColorFondo
        .Cell(flexcpBackColor, 10, 6, 41, 6) = ColorFondo
        .Cell(flexcpBackColor, 10, 9, 42, 9) = ColorFondo
        .Cell(flexcpBackColor, 17, 8, 17, 10) = ColorFondo
        .Cell(flexcpBackColor, 36, 1, 37, 1) = ColorFondo
        .Cell(flexcpBackColor, 41, 9, 41, 10) = ColorFondo
        .Cell(flexcpBackColor, 43, 3, 45, 3) = ColorFondo
        .Cell(flexcpBackColor, 45, 6, 45, 6) = ColorFondo
        .Cell(flexcpBackColor, 48, 1, 58, 1) = ColorFondo
        .Cell(flexcpBackColor, 47, 6, 58, 6) = ColorFondo
        .Cell(flexcpBackColor, 47, 9, 60, 9) = ColorFondo
        .Cell(flexcpBackColor, 47, 1, 47, 6) = ColorFondo
        .Cell(flexcpBackColor, 47, 8, 47, 8) = ColorFondo
        .Cell(flexcpBackColor, 62, 9, 66, 9) = ColorFondo
        .Cell(flexcpBackColor, 66, 1, 66, 1) = ColorFondo
        .Cell(flexcpBackColor, 66, 3, 66, 3) = ColorFondo
        .Cell(flexcpBackColor, 67, 9, 68, 9) = ColorFondo
        .Cell(flexcpBackColor, 70, 1, 73, 1) = ColorFondo
        .Cell(flexcpBackColor, 70, 3, 73, 3) = ColorFondo
        'alineacion
        .Cell(flexcpAlignment, 7, 4, 7, 4) = 1
        .Cell(flexcpAlignment, 10, 1, 14, 1) = 3
        .Cell(flexcpAlignment, 59, 7, 66, 7) = 1
        .Cell(flexcpAlignment, 66, 4, 66, 4) = 1
        .Cell(flexcpAlignment, 45, 5, 45, 5) = 1
        .Cell(flexcpAlignment, 61, 6, 61, 6) = 1
        
        'alto linea
        .RowHeight(1) = 400
        .RowHeight(2) = 320
        .RowHeight(4) = 320
        .RowHeight(6) = 320
        .RowHeight(9) = 320
        .RowHeight(46) = 320
        .RowHeight(61) = 400
       
        'tamaño letras
        .Cell(flexcpFontSize, 1, 1, 1, 6) = 14
        .Cell(flexcpFontSize, 2, 1, 2, 8) = 10
        .Cell(flexcpFontSize, 4, 2, 4, 2) = 11
        .Cell(flexcpFontSize, 6, 1, 6, 8) = 10
        .Cell(flexcpFontSize, 9, 1, 9, 10) = 10
        .Cell(flexcpFontSize, 46, 1, 46, 10) = 10
        .Cell(flexcpFontSize, 61, 6, 61, 8) = 10
        
        'negritas
        .Cell(flexcpFontBold, 2, 3, 2, 9) = True
        .Cell(flexcpFontBold, 4, 2, 4, 2) = True
        .Cell(flexcpFontBold, 6, 1, 6, 9) = True
        .Cell(flexcpFontBold, 9, 1, 9, 10) = True
        .Cell(flexcpFontBold, 46, 1, 46, 10) = True
        .Cell(flexcpFontBold, 61, 6, 61, 9) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .MergeCol(4) = True
        .MergeCol(5) = True
    End With
End Sub

Private Sub LLENADATOS103()
    With grd
        .Redraw = flexRDBuffered
        .TextMatrix(8, 2) = gobjMain.EmpresaActual.GNOpcion.ruc
        .TextMatrix(8, 4) = gobjMain.EmpresaActual.GNOpcion.NombreEmpresa
         .Refresh
    End With
End Sub

Private Sub LlenaFormatoFormulario103()
    With grd
        .MergeCells = flexMergeSpill

        .AddItem "1" & vbTab & vbTab & "DECLARACION DE RETENCIONES EN LA FUENTE DEL IMPUESTO A LA RENTA" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "No."
        .AddItem "2" & vbTab & vbTab & vbTab & vbTab & "100" & vbTab & "IDENTIFICACION DE LA DECLARACION"
        .AddItem "3" & vbTab & vbTab & vbTab & vbTab & "DECLARACION MENSUAL"
        .AddItem "4" & vbTab & vbTab & "FORMULARIO 103" & vbTab & vbTab & "101" & vbTab & "MES " & vbTab & vbTab & vbTab & "102" & vbTab & "AÑO" & vbTab
        .AddItem "5" & vbTab & "" & vbTab & "" & vbTab & vbTab & vbTab & vbTab & "104" & vbTab & "No. FORMULARIO QUE SE RETIFICA"
        .AddItem "6" & vbTab & "200" & vbTab & "IDENTIFICACION DEL SUJETO PASIVO (AJENTE DE PERCEPCION O RETENCION)"
        .AddItem "7" & vbTab & vbTab & "RUC" & vbTab & vbTab & "RAZON SOCIAL DENOMINACION O APELLIDOS Y NOMBRES COMPLETOS"
        .AddItem "8" & vbTab & "201" & vbTab & vbTab & "202" & vbTab
        .AddItem "9" & vbTab & "300" & vbTab & "POR PAGOS EN EL PAIS" & vbTab & vbTab & vbTab & vbTab & "BASE IMPONIBLE " & vbTab & vbTab & "% RET" & vbTab & "IMPUESTO RETENIDO"
        .AddItem "10" & vbTab & "ORIGINADOS" & vbTab & "EN RELACIÓN DE DEPENDENCIA QUE NO SUPERA LA BASE DESGRAVADA" & vbTab & vbTab & vbTab & vbTab & "301"
        .AddItem "11" & vbTab & "" & vbTab & "EN RELACIÓN DE DEPENDENCIA QUE SUPERA LA BASE DESGRAVADA" & vbTab & vbTab & vbTab & vbTab & "302" & vbTab & vbTab & vbTab & "352"
        .AddItem "12" & vbTab & "EN EL" & vbTab & "HONORARIOS, COMISIONES Y DIETAS A PERSONAS NATURALES" & vbTab & vbTab & vbTab & vbTab & "303" & vbTab & vbTab & "8%" & vbTab & "353"
        .AddItem "13" & vbTab & "" & vbTab & "REMUNERACION A OTROS TRABAJADORES AUTONOMOS" & vbTab & vbTab & vbTab & vbTab & "304" & vbTab & vbTab & "1%" & vbTab & "354"
        .AddItem "14" & vbTab & "TRABAJO" & vbTab & "HONORARIOS A EXTRANJEROS POR SERVICIOS OCASIONALES" & vbTab & vbTab & vbTab & vbTab & "305" & vbTab & vbTab & "25%" & vbTab & "355"
        .AddItem "15" & vbTab & "POR COMPRAS LOCALES DE MATERIA PRIMA" & vbTab & vbTab & vbTab & vbTab & vbTab & "306" & vbTab & vbTab & "2%" & vbTab & "356"
        .AddItem "16" & vbTab & "POR COMPRAS LOCALES DE BIENES NO PRODUCIDOS POR LA SOCIEDAD" & vbTab & vbTab & vbTab & vbTab & vbTab & "307" & vbTab & vbTab & "2%" & vbTab & "357"
        .AddItem "17" & vbTab & "POR COMPRAS LOCALES DE MATERIA PRIMA NO SUJETA A RETENCIÓN" & vbTab & vbTab & vbTab & vbTab & vbTab & "308"
        .AddItem "18" & vbTab & "POR SUMINISTROS Y MATERIALES" & vbTab & vbTab & vbTab & vbTab & vbTab & "309" & vbTab & vbTab & "2%" & vbTab & "359"
        .AddItem "19" & vbTab & "POR REPUESTOS Y HERRAMIENTAS" & vbTab & vbTab & vbTab & vbTab & vbTab & "310" & vbTab & vbTab & "2%" & vbTab & "360"
        .AddItem "20" & vbTab & "POR LUBRICANTES" & vbTab & vbTab & vbTab & vbTab & vbTab & "311" & vbTab & vbTab & "1%" & vbTab & "361"
        .AddItem "21" & vbTab & "POR ACTIVOS FIJOS" & vbTab & vbTab & vbTab & vbTab & vbTab & "312" & vbTab & vbTab & "2%" & vbTab & "362"
        .AddItem "22" & vbTab & "POR CONCEPTO DE SERVICIO DE TRANSPORTE PRIVADO DE PASAJEROS O SERVICIO PUBLICO O PRIVADO DE CARGA" & vbTab & vbTab & vbTab & vbTab & vbTab & "313" & vbTab & vbTab & "2%" & vbTab & "363"
        .AddItem "23" & vbTab & "POR REGALIAS, DERECHOS DE AUTOR, MARCAS, PATENTES Y SIMILARES" & vbTab & vbTab & vbTab & vbTab & vbTab & "314" & vbTab & vbTab & "8%" & vbTab & "364"
        .AddItem "24" & vbTab & "POR REMUNERACIONES A DEPORTISTAS, ENTRENADORES, CUERPO TECNICO, ARBITROS Y ARTISTAS RESIDENTES" & vbTab & vbTab & vbTab & vbTab & vbTab & "315" & vbTab & vbTab & "5%" & vbTab & "365"
        .AddItem "25" & vbTab & "POR PAGOS REALIZADOS A NOTARIOS Y REGISTRADORES DE LA PROPIEDAD O MERCANTILES" & vbTab & vbTab & vbTab & vbTab & vbTab & "316" & vbTab & vbTab & "8%" & vbTab & "366"
        .AddItem "26" & vbTab & "POR COMISIONES PAGADAS A SOCIEDADES " & vbTab & vbTab & vbTab & vbTab & vbTab & "317" & vbTab & vbTab & "5%" & vbTab & "367"
        .AddItem "27" & vbTab & "POR PROMOCION Y PUBLICIDAD" & vbTab & vbTab & vbTab & vbTab & vbTab & "318" & vbTab & vbTab & "8%" & vbTab & "368"
        .AddItem "28" & vbTab & "POR ARRENDAMIENTO MERCANTIL LOCAL" & vbTab & vbTab & vbTab & vbTab & vbTab & "319" & vbTab & vbTab & "1%" & vbTab & "369"
        .AddItem "29" & vbTab & "POR ARRENDAMIENTO DE BIENES INMUEBLES DE PROPIEDAD DE PERSONAS NATURALES" & vbTab & vbTab & vbTab & vbTab & vbTab & "320" & vbTab & vbTab & "8%" & vbTab & "370"
        .AddItem "30" & vbTab & "POR ARRENDAMIENTO DE BIENES INMUEBLES A SOCIEDADES" & vbTab & vbTab & vbTab & vbTab & vbTab & "321" & vbTab & vbTab & "5%" & vbTab & "371"
        .AddItem "31" & vbTab & "POR SEGUROS Y REASEGUROS (10% del valor de las primas facturadas)" & vbTab & vbTab & vbTab & vbTab & vbTab & "322" & vbTab & vbTab & "1%" & vbTab & "372"
        .AddItem "32" & vbTab & "POR RENDIMIENTOS FINANCIEROS" & vbTab & vbTab & vbTab & vbTab & vbTab & "323" & vbTab & vbTab & "5%" & vbTab & "373"
        .AddItem "33" & vbTab & "POR PAGOS O CREDITOS EN CUENTA REALIZADOS POR EMPRESAS EMISORAS DE TARJETAS DE CREDITO" & vbTab & vbTab & vbTab & vbTab & vbTab & "324" & vbTab & vbTab & "1%" & vbTab & "374"
        .AddItem "34" & vbTab & "POR LOTERIAS, RIFAS, APUESTAS Y SIMILARES" & vbTab & vbTab & vbTab & vbTab & vbTab & "325" & vbTab & vbTab & "15%" & vbTab & "375"
        .AddItem "35" & vbTab & "POR INTERESES Y COMISIONES EN OPERACIONES DE CREDITO ENTRE LAS INST. DEL SISTEMA FINANCIERO" & vbTab & vbTab & vbTab & vbTab & vbTab & "326" & vbTab & vbTab & "1%" & vbTab & "376"
        .AddItem "36" & vbTab & "POR VENTA DE" & vbTab & "A COMERCIALIZADORAS " & vbTab & vbTab & vbTab & vbTab & "327" & vbTab & vbTab & "2/mil" & vbTab & "377"
        .AddItem "37" & vbTab & "COMBUSTIBLES" & vbTab & "A DISTRIBUIDORES" & vbTab & vbTab & vbTab & vbTab & "328" & vbTab & vbTab & "3/mil" & vbTab & "378"
        .AddItem "38" & vbTab & "POR OTROS SERVICIOS" & vbTab & vbTab & vbTab & vbTab & vbTab & "329" & vbTab & vbTab & "2%" & vbTab & "379"
        .AddItem "39" & vbTab & "POR PAGOS DE DIVIDENDOS ANTICIPADOS" & vbTab & vbTab & vbTab & vbTab & vbTab & "330" & vbTab & vbTab & "25%" & vbTab & "380"
        .AddItem "40" & vbTab & "POR AGUA, ENERGÍA, LUZ Y TELECOMUNICACIONES" & vbTab & vbTab & vbTab & vbTab & vbTab & "331" & vbTab & vbTab & "1%" & vbTab & "381"
        .AddItem "41" & vbTab & "OTRAS COMPRAS DE BIENES Y SERVICIOS NO SUJETAS A RETENCIÓN" & vbTab & vbTab & vbTab & vbTab & vbTab & "332"
        .AddItem "42" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "SUBTOTAL SUMAR 352 AL 381" & vbTab & vbTab & "399"
        .AddItem "43" & vbTab & "No. de empleados en nómina" & vbTab & vbTab & "105"
        .AddItem "44" & vbTab & "No. de empleados tercerizados" & vbTab & vbTab & "106"
        .AddItem "45" & vbTab & "No. de empleados bajo contrato" & vbTab & vbTab & "107" & vbTab & vbTab & "No. de comprobantes de retención emitidos" & vbTab & "108"
        .AddItem "46" & vbTab & "300" & vbTab & " 400 POR PAGOS AL EXTERIOR" & vbTab & vbTab & vbTab & vbTab & "BASE IMPONIBLE " & vbTab & vbTab & "% RET" & vbTab & "IMPUESTO RETENIDO"
        .AddItem "47" & vbTab & "CON CONVENIO DE DOBLE TRIBUTACION" & vbTab & vbTab & vbTab & vbTab & vbTab & "401" & vbTab & vbTab & vbTab & "451"
        .AddItem "48" & vbTab & "" & vbTab & "INTERESES Y COSTOS FINANCIDEROS POR FINANCIAMIENTO DE PROVEEDORES EXTERNOS (EN LA CUANTIA QUE EXCEDE A LA TASA MÁXIMA)" & vbTab & vbTab & vbTab & vbTab & "403" & vbTab & vbTab & "25%" & vbTab & "453"
        .AddItem "49" & vbTab & "SIN" & vbTab & "INTERESES DE CRÉDITOS EXTERNOS REGISTRADOS EN EL BCE (EN LA CUANTIA QUE EXCEDE A LA TASA MÁXIMA)" & vbTab & vbTab & vbTab & vbTab & "405" & vbTab & vbTab & "25%" & vbTab & "455"
        .AddItem "50" & vbTab & "" & vbTab & "INTERESES DE CRÉDITOS EXTERNOS NO REGISTRADOS EN EL BCE" & vbTab & vbTab & vbTab & vbTab & "407" & vbTab & vbTab & "25%" & vbTab & "457"
        .AddItem "51" & vbTab & "CONVENIO" & vbTab & "COMISIONES POR EXPORTACIONES" & vbTab & vbTab & vbTab & vbTab & "409" & vbTab & vbTab & "25%" & vbTab & "459"
        .AddItem "52" & vbTab & "" & vbTab & "COMISIONES PAGADAS PARA LA PROMOCION DEL TURISMO RECEPTIVO" & vbTab & vbTab & vbTab & vbTab & "411" & vbTab & vbTab & "25%" & vbTab & "461"
        .AddItem "53" & vbTab & "DOBLE" & vbTab & "EL 4% DE LAS PRIMAS DE CESIÓN O REASEGUROS CONTRATADOS CON EMPRESAS QUE NO TENGAN ESTABLECIMIENTO O REPRESENTACIÓN PERMANENTE EN EL ECUADOR" & vbTab & vbTab & vbTab & vbTab & "413" & vbTab & vbTab & "25%" & vbTab & "463"
        .AddItem "54" & vbTab & "" & vbTab & "EL 10% DE LOS PAGOS EFECTUADOS POR LAS AGENCIAS INTERNACIONALES DE PRENSA REGISTRADAS EN LA SECRETARÍA DE COMUNICACIÓN DEL ESTADO" & vbTab & vbTab & vbTab & vbTab & "415" & vbTab & vbTab & "25%" & vbTab & "465"
        .AddItem "55" & vbTab & "TRIBUTACION" & vbTab & "EL 10% DEL VALOR DE LOS CONTRATOS DE FLETAMENTO DE NAVES PARA EMPRESAS DE TRANSPORTE AÉREO O MARÍTIMO INTERNACIONAL" & vbTab & vbTab & vbTab & vbTab & "417" & vbTab & vbTab & "25%" & vbTab & "467"
        .AddItem "56" & vbTab & "" & vbTab & "EL 15% DE LOS PAGOS EFECTUADOS POR PRODUCTORAS Y DISTRIBUIDORAS DE CINTAS CINEMATOGRÁFICAS Y DE TELEVISIÓN POR CONCEPTO DE ARRENDAMIENTO DE CINTAS Y VIDEOCINTAS" & vbTab & vbTab & vbTab & vbTab & "419" & vbTab & vbTab & "25%" & vbTab & "469"
        .AddItem "57" & vbTab & "ARRENDAMIENTO" & vbTab & "POR PAGO DE INTERESES (cuando supera la tasa autorizada por el BCE)" & vbTab & vbTab & vbTab & vbTab & "421" & vbTab & vbTab & "25%" & vbTab & "471"
        .AddItem "58" & vbTab & "MERCANTTIL INT." & vbTab & "CUANDO NO SE EJERCE LA OPCION DE COMPRA (sobre la depreciación acumulada)" & vbTab & vbTab & vbTab & vbTab & "423" & vbTab & vbTab & "25%" & vbTab & "473"
        .AddItem "59" & vbTab & "Declaro que los datos contenidos en esta declaración son verdaderos por lo que asumo la responsabilidad " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "SUBTOTAL SUMAR 451 al 473" & vbTab & vbTab & "498"
        .AddItem "60" & vbTab & "correspondiente (Artículo 101 de la de la Codificación 2004-026 de la L.R.T.I.) " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "TOTAL RETENCIONES 399 + 498" & vbTab & vbTab & "499"
        .AddItem "61" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "900 VALORES A PAGAR Y FORMA DE PAGO"
        .AddItem "62" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "PAGO PREVIO" & vbTab & vbTab & "901"
        .AddItem "63" & vbTab & vbTab & "____________________________" & vbTab & vbTab & vbTab & "____________________________" & vbTab & vbTab & "TOTAL IMPUESTO A PAGAR 499 - 901" & vbTab & vbTab & "902"
        .AddItem "64" & vbTab & vbTab & "FIRMA SUJETO PASIVO" & vbTab & vbTab & vbTab & "FIRMA CONTADOR" & vbTab & vbTab & "INTERESES POR MORA" & vbTab & vbTab & "903"
        .AddItem "65" & vbTab & "NOMBRE:" & vbTab & vbTab & "NOMBRE:" & vbTab & vbTab & vbTab & vbTab & "MULTAS" & vbTab & vbTab & "904"
        .AddItem "66" & vbTab & "198" & vbTab & "C.I. No." & vbTab & "199" & vbTab & "RUC No." & vbTab & vbTab & vbTab & "TOTAL PAGADO     902+903+904" & vbTab & vbTab & "999"
        .AddItem "67" & vbTab & "MEDIANTE CHEQUE DEBITO BANCARIO EFECTIVO U OTRAS FORMAS DE COBRO" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "905"
        .AddItem "68" & vbTab & "MEDIANTE NOTAS DE CREDITO" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "907"
        .AddItem "69" & vbTab & vbTab & "DETALLE DE NOTAS DE CREDITO"
        .AddItem "70" & vbTab & "908" & vbTab & "N/C No." & vbTab & "909" & vbTab & "USD"
        .AddItem "71" & vbTab & "910" & vbTab & "N/C No." & vbTab & "911" & vbTab & "USD"
        .AddItem "72" & vbTab & "912" & vbTab & "N/C No." & vbTab & "913" & vbTab & "USD"
        .AddItem "73" & vbTab & "914" & vbTab & "N/C No." & vbTab & "915" & vbTab & "USD"

        .Redraw = flexRDBuffered
        .Refresh
    End With
    CambiaFondoCeldasEditables103
End Sub

Private Sub Buscar103()
    Dim sql As String, cond As String, Reten As String, CPNotRet As String
    Dim OrdenadoX As String, f1 As String
    Dim rs As Recordset
    Dim Moneda As String
    Set objcond = gobjMain.objCondicion
    If Not frmB_FormSRI103.Inicio103(objcond, Reten) Then
        grd.SetFocus
        Exit Sub
    End If
    With objcond
        If Len(Month(.fecha1)) < 2 Then
            grd.TextMatrix(4, 6) = "0" & Month(.fecha1)
        Else
            grd.TextMatrix(4, 6) = Month(.fecha1)
        End If
        
        
        grd.TextMatrix(4, 9) = " AÑO " & Year(.fecha1)
           
        'Reporte de un mes a la vez
        f1 = DateSerial(Year(.fecha1), Month(.fecha1), 1)
        cond = " AND GNC.FechaTrans BETWEEN " & FechaYMD(f1, gobjMain.EmpresaActual.TipoDB) & _
               " AND " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB)

        sql = " SELECT"
        sql = sql & " CodF104,  sum(tskardexret.BAse) AS TBase"
        sql = sql & " from GNComprobante gnc "
        sql = sql & " inner join gntrans gnt on gnc.codtrans=gnt.codtrans"
        sql = sql & " INNER JOIN tskardexret "
        sql = sql & " INNER JOIN tsretencion  ON tskardexret.idretencion = tsretencion.idretencion  "
        sql = sql & " ON GNC.TransID = tskardexret.transid"
        sql = sql & " Where GNC.Estado <> 3 " & cond
        sql = sql & " and tsretencion.BandValida=1"
        sql = sql & " and len(CodF104)>0"
        sql = sql & " and bandiva=0"
        sql = sql & " and AnexoCodTipoComp=7"
        If Len(Reten) > 0 Then
            sql = sql & " and tsretencion.CodRetencion in(" & PreparaCadena(Reten) & ") "
        End If
        sql = sql & " group BY CodF104"
            
            
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            
            If rs.RecordCount > 0 Then 'rs.MoveFirst
                While Not rs.EOF
                        If Len(rs.Fields("TBase")) > 0 Then
                            grd.TextMatrix(rs.Fields("CodF104") - 291, 7) = rs.Fields("TBase")
                        End If
                    rs.MoveNext
                Wend
            End If

            sql = " SELECT"
            sql = sql & " count(gnc.numtrans) As numtrans"
            sql = sql & " from GNComprobante gnc"
            sql = sql & " inner join gntrans gnt on gnc.codtrans=gnt.codtrans"
            sql = sql & " Where gnc.Estado <> 3" & cond
            sql = sql & " and AnexoCodTipoComp=7"
            
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            
            If rs.RecordCount > 0 Then
                grd.TextMatrix(45, 7) = rs.Fields("numtrans")
            End If
            
            
            CalcularPorcentajes103
            nombre = "103_" & Format((Month(objcond.fecha1)), "00") & "_" & Year(objcond.fecha1) & ".XML"
            txtDestino.Text = mRutaDestino104 & "103_" & Format(CStr(Month(objcond.fecha1)), "00") & "_" & Year(objcond.fecha1) & ".XML"
            grd.Refresh
    End With
End Sub


Private Sub CalcularPorcentajes103()
    Dim i As Integer, TotalRet As Currency
    Dim SubTotal As Currency
    Dim subtotal1 As Currency
    Dim SubtotalRetenido As Currency, TotalCompras As Currency
    TotalRet = 0
    For i = 16 To 34
        'If i = 17 Then
        '    i = i + 1
       ' End If
        SubTotal = SubTotal + Round(grd.ValueMatrix(i, 10), 2)
    '    subtotal1 = subtotal1 + Round(grd.ValueMatrix(i, 12), 2)
        
    Next i
    grd.TextMatrix(35, 10) = SubTotal
'    grd.TextMatrix(35, 12) = subtotal1
'    For i = 12 To 40
'        If i = 17 Then
'            i = i + 1
'        End If
'        grd.TextMatrix(i, 10) = grd.ValueMatrix(i, 7) * grd.ValueMatrix(i, 8)
'        TotalRet = TotalRet + grd.ValueMatrix(i, 10)
'    Next i
    
    grd.TextMatrix(42, 10) = TotalRet
    
    TotalRet = 0
    
    For i = 48 To 58
        grd.TextMatrix(i, 10) = grd.ValueMatrix(i, 7) * grd.ValueMatrix(i, 8)
        TotalRet = TotalRet + grd.ValueMatrix(i, 10)
    Next i
    grd.TextMatrix(59, 10) = TotalRet
    
    grd.TextMatrix(60, 10) = grd.ValueMatrix(42, 10) + grd.ValueMatrix(59, 10)
    grd.TextMatrix(63, 10) = grd.ValueMatrix(60, 10) + grd.ValueMatrix(62, 10)
    grd.TextMatrix(66, 10) = grd.ValueMatrix(63, 10) + grd.ValueMatrix(64, 10) + grd.ValueMatrix(65, 10)
End Sub

Private Function GeneraArchivo103() As String
    Dim obj As GNOpcion, cad As String, i As Integer
    Dim mes As Integer
    
    cad = "<?xml version=" & """1.0""" & " encoding=" & """UTF-8""" & " standalone =" & """yes""" & "?>"
    cad = cad & "<formulario version=" & """0.2""" & ">"
    cad = cad & "<cabecera>"
    cad = cad & "<codigo_version_formulario>03201202</codigo_version_formulario>"
    cad = cad & "<ruc>" & gobjMain.EmpresaActual.GNOpcion.ruc & "</ruc>"
    cad = cad & "<codigo_moneda>1</codigo_moneda>"
    cad = cad & "</cabecera>"
    cad = cad & "<detalle>"
    Select Case grd.TextMatrix(7, 2)
        Case "Enero", "enero": mes = 1
        Case "Febrero", "febrero": mes = 2
        Case "Marzo", "marzo": mes = 3
        Case "Abril", "abril": mes = 4
        Case "Mayo", "mayo": mes = 5
        Case "Junio", "junio": mes = 6
        Case "Julio", "julio": mes = 7
        Case "Agosto", "agosto": mes = 8
        Case "Septiembre", "septiembre": mes = 9
        Case "Octubre", "octubre": mes = 10
        Case "Noviembre", "noviembre": mes = 11
        Case "Diciembre", "diciembre": mes = 12
    End Select
     'cad = cad & "<campo numero=""" & "31" & """>" & "O" & "</campo>"
     If grd.TextMatrix(7, 12) = "Original" Then
        cad = cad & "<campo numero=""" & "31" & """>" & "O" & "</campo>"
    ElseIf grd.TextMatrix(7, 12) = "Sustituta" Then
        cad = cad & "<campo numero=""" & "31" & """>" & "S" & "</campo>"
    Else
        MsgBox "Debe escoger el tipo de formulario [Original - Sustituto]"
        Exit Function
    End If
    cad = cad & "<campo numero=""" & "101" & """>" & mes & "</campo>"
    cad = cad & "<campo numero=""" & "102" & """>" & grd.TextMatrix(7, 4) & "</campo>"
    
    'cad = cad & "<campo numero=""" & "104" & """> </campo>"
    If grd.TextMatrix(7, 12) = "Original" Then
        cad = cad & "<campo numero=""" & "104" & """>   </campo>"
    ElseIf grd.TextMatrix(7, 12) = "Sustituta" Then
        cad = cad & "<campo numero=""" & "104" & """>" & grd.TextMatrix(8, 11) & "   </campo>"
    End If
    
    cad = cad & "<campo numero=""" & "198" & """>" & grd.TextMatrix(64, 4) & " </campo>"
    cad = cad & "<campo numero=""" & "199" & """>" & grd.TextMatrix(64, 9) & " </campo>"
    cad = cad & "<campo numero=""" & "201" & """>" & grd.TextMatrix(10, 2) & "</campo>"
    cad = cad & "<campo numero=""" & "202" & """>" & grd.TextMatrix(10, 4) & "</campo>"
    For i = 15 To 39
                cad = cad & "<campo numero=""" & grd.TextMatrix(i, 9) & """>" & Format(grd.ValueMatrix(i, 10), "#.00") & "</campo>"
    Next i
    For i = 15 To 32
                cad = cad & "<campo numero=""" & grd.TextMatrix(i, 11) & """>" & Format(grd.ValueMatrix(i, 12), "#.00") & "</campo>"
    Next i
    
    For i = 34 To 39
                cad = cad & "<campo numero=""" & grd.TextMatrix(i, 11) & """>" & Format(grd.ValueMatrix(i, 12), "#.00") & "</campo>"
    Next i
    
    For i = 41 To 47
                cad = cad & "<campo numero=""" & grd.TextMatrix(i, 9) & """>" & Format(grd.ValueMatrix(i, 10), "#.00") & "</campo>"
    Next i
    
    For i = 41 To 45
                cad = cad & "<campo numero=""" & grd.TextMatrix(i, 11) & """>" & Format(grd.ValueMatrix(i, 12), "#.00") & "</campo>"
    Next i
    
    
    
    cad = cad & "<campo numero=""" & grd.TextMatrix(47, 11) & """>" & Format(grd.ValueMatrix(47, 12), "#.00") & "</campo>" '498
    cad = cad & "<campo numero=""" & grd.TextMatrix(49, 11) & """>" & Format(grd.ValueMatrix(49, 12), "#.00") & "</campo>" '499
    
    cad = cad & "<campo numero=""" & "510" & """></campo><campo numero=""" & "520" & """></campo>"
    cad = cad & "<campo numero=""" & "897" & """>" & Format(grd.ValueMatrix(53, 3), "#.00") & "</campo>"
    
    cad = cad & "<campo numero=""" & grd.TextMatrix(31, 6) & """>" & Format(grd.ValueMatrix(31, 8), "#.00") & "</campo>" '510
    cad = cad & "<campo numero=""" & grd.TextMatrix(32, 6) & """>" & Format(grd.ValueMatrix(32, 8), "#.00") & "</campo>" '520
    
    
    cad = cad & "<campo numero=""" & "880" & """>" & Format(grd.ValueMatrix(53, 12), "#.00") & "</campo>" '880
    
    cad = cad & "<campo numero=""" & grd.TextMatrix(51, 11) & """>" & Format(grd.ValueMatrix(51, 12), "#.00") & "</campo>"
    
    cad = cad & "<campo numero=""" & "897" & """>" & Format(grd.ValueMatrix(53, 3), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "898" & """>" & Format(grd.ValueMatrix(53, 6), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "899" & """>" & Format(grd.ValueMatrix(53, 12), "#.00") & "</campo>"
    
     For i = 57 To 59
                cad = cad & "<campo numero=""" & grd.TextMatrix(i, 11) & """>" & Format(grd.ValueMatrix(i, 12), "#.00") & "</campo>"
    Next i
    For i = 62 To 63
                cad = cad & "<campo numero=""" & grd.TextMatrix(i, 11) & """>" & Format(grd.ValueMatrix(i, 12), "#.00") & "</campo>"
    Next i
'If grd.TextMatrix(59, 2) = "N/C No." Then
 '       cad = cad & "<campo numero=""" & "908" & """>" & "</campo>"
  '  Else
    cad = cad & "<campo numero=""" & "908" & """>" & Format(grd.TextMatrix(65, 3), "#.00") & "</campo>"
   ' End If
    cad = cad & "<campo numero=""" & "909" & """>" & Format(grd.ValueMatrix(66, 3), "#.00") & "</campo>"
    'If grd.TextMatrix(59, 4) = "N/C No." Then
     '   cad = cad & "<campo numero=""" & "910" & """>" & "</campo>"
    'Else
    cad = cad & "<campo numero=""" & "910" & """>" & Format(grd.TextMatrix(65, 5), "#.00") & "</campo>"
    'End If
    cad = cad & "<campo numero=""" & "911" & """>" & Format(grd.ValueMatrix(66, 5), "#.00") & "</campo>"
    'If grd.TextMatrix(59, 6) = "N/C No." Then
    '    cad = cad & "<campo numero=""" & "912" & """>" & "</campo>"
    'Else
    cad = cad & "<campo numero=""" & "912" & """>" & Format(grd.TextMatrix(65, 8), "#.00") & "</campo>"
   ' End If
    cad = cad & "<campo numero=""" & "913" & """>" & Format(grd.ValueMatrix(65, 8), "#.00") & "</campo>"
    'If grd.TextMatrix(59, 10) = "N/C No." Then
     '   cad = cad & "<campo numero=""" & "914" & """>" & "</campo>"
    'Else
        'cad = cad & "<campo numero=""" & "914" & """>" & Format(grd.TextMatrix(59, 12), "#.00") & "</campo>"
    'End If
    cad = cad & "<campo numero=""" & "915" & """>" & Format(grd.ValueMatrix(66, 12), "#.00") & "</campo>"
    cad = cad & "<campo numero=""" & "922" & """>" & "16" & "</campo>"
    cad = cad & "<campo numero=""" & "999" & """>" & Format(grd.ValueMatrix(60, 12), "#.00") & "</campo>"
    cad = cad & "</detalle>"
    cad = cad & "</formulario>"
    GeneraArchivo103 = cad
End Function

'*******************2008
Private Sub LlenaFormatoFormulario104()
    With grd
        .MergeCells = flexMergeSpill

        .AddItem "1" & vbTab & vbTab & "DECLARACION DEL IMPUESTO AL VALOR AGREGADO" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "No."
        .AddItem "2" & vbTab & vbTab & vbTab & vbTab & "100" & vbTab & "IDENTIFICACION DE LA DECLARACION"
        .AddItem "3" & vbTab & vbTab & vbTab & vbTab & "DECLARACION MENSUAL"
        .AddItem "4" & vbTab & vbTab & "FORMULARIO 104" & vbTab & vbTab & "101" & vbTab & "MES " & vbTab & vbTab & vbTab & "102" & vbTab & "AÑO"
        .AddItem "5" & vbTab & vbTab & "2008"
        .AddItem "6" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "104" & vbTab & "No. FORMULARIO QUE SE RETIFICA"
        .AddItem "7" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab
        .AddItem "8" & vbTab & "200" & vbTab & "IDENTIFICACION DEL SUJETO PASIVO (AJENTE DE PERCEPCION O RETENCION)"
        .AddItem "9" & vbTab & vbTab & "RUC" & vbTab & vbTab & "RAZON SOCIAL DENOMINACION O APELLIDOS Y NOMBRES COMPLETOS"
        .AddItem "10" & vbTab & "201" & vbTab & vbTab & "202" & vbTab
        .AddItem "11" & vbTab & "300" & vbTab & "PROPORCION DE CREDITO TRIBUTARIO APLICABLE EN ESTE MES"
        .AddItem "12" & vbTab & "VENTAS CON TARIFA 12+ VENTAS NETAS A INSTITUCIONES YEMPRESAS PUBLICAS + EXPORTACIONES NETAS+VENTAS NETAS  ..." & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "301"
        .AddItem "13" & vbTab & "SALDO DEL CREDITO TRIBUTARIO MES ANETRIOR" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "302"
        .AddItem "14" & vbTab & "(-) DEVOLUCIONES DE IVA SOLICITADAS EN ESTE MES" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "303"
        .AddItem "15" & vbTab & "(+) DEVOLUCIONES RECHAZADAS IMPUTABLES A CRED. TRIBUTARIO" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "304"
        .AddItem "16" & vbTab & "(=) SALDO CREDITO TRIBUTARIO APLICARSE ENESTE MES" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "SUMA (302-303+304)" & vbTab & "399"
'        '  VENTAS Y EXPORTACIONES
        .AddItem "17" & vbTab & "400" & vbTab & "RESUMEN DE TRANSFERENCIAS Y OTRAS OPERACIONES DEL PERIODO QUE DECLARA"
        .AddItem "18" & vbTab & "TRANSFERENCIAS OBJETO DEL VALOR AGRGADO" & vbTab & vbTab & vbTab & vbTab & "BASE IMPONIBLE 0%" & vbTab & vbTab & "BASE IMPONIBLE 12%" & vbTab & "IMPUESTOGENERADO"
        .AddItem "19" & vbTab & "G" & vbTab & "VENTAS LOCALES  (EXCLUYE ACT. FIJOS Y OTROS)" & vbTab & vbTab & "401" & vbTab & vbTab & "421" & vbTab & vbTab & "441"
        .AddItem "20" & vbTab & "R" & vbTab & "VENTAS DE ACTIVOS FIJOS " & vbTab & vbTab & "402" & vbTab & vbTab & "422" & vbTab & vbTab & "442"
        .AddItem "21" & vbTab & "U" & vbTab & "OTROS (Donaciones promociones autoconsumos etc) " & vbTab & vbTab & "403" & vbTab & vbTab & "423" & vbTab & vbTab & "443"
        .AddItem "22" & vbTab & "P" & vbTab & "NOT. CREDITO POR TRANSF GRUPO A MES ACTUAL" & vbTab & vbTab & "404" & vbTab & vbTab & "424" & vbTab & vbTab & "444"
        .AddItem "23" & vbTab & "O" & vbTab & "NOT. CREDITO POR TRANSF GRUPO A NO COMP MES ANT" & vbTab & vbTab & "405" & vbTab & vbTab & "425" & vbTab & vbTab & "445"
        .AddItem "24" & vbTab & "A" & vbTab & "SUBTOTAL GRUPO A BASE IMP 0% 401+402+403-404-405" & vbTab & vbTab & "406"
        .AddItem "25" & vbTab & " " & vbTab & "VENTAS DIRECTAS A EXPORTADORES" & vbTab & vbTab & "407" & vbTab & vbTab & "427" & vbTab & vbTab & "447"
        .AddItem "26" & vbTab & "G" & vbTab & "EXPORTACION DE BIENES" & vbTab & vbTab & "408"
        .AddItem "27" & vbTab & "R" & vbTab & "EXPORTACION DE SERVICIOS" & vbTab & vbTab & "409"
        .AddItem "28" & vbTab & "U" & vbTab & "VENTAS LOCALES (excluyen activos fijos)" & vbTab & vbTab & "410"
        .AddItem "29" & vbTab & "P" & vbTab & "VENTAS DE ACTIVOS FIJOS" & vbTab & vbTab & "411"
        .AddItem "30" & vbTab & "O" & vbTab & "NOT. CREDITO POR TRANSF GRUPO B MES ACTUAL" & vbTab & vbTab & "412" & vbTab & vbTab & "432" & vbTab & vbTab & "452"
        .AddItem "31" & vbTab & "" & vbTab & "NOT. CREDITO POR TRANSF GRUPO A NO COMPENSADAS MES ANT" & vbTab & vbTab & "413" & vbTab & vbTab & "433" & vbTab & vbTab & "453"
        .AddItem "32" & vbTab & "B" & vbTab & "SUBTOTAL GRUPO B BASE IMPUESTO 0%" & vbTab & vbTab & "414"
        .AddItem "33" & vbTab & "TOTAL BASE IMPUESTO 12%              421+422+423-424-425+427-432-453" & vbTab & vbTab & "" & vbTab & vbTab & vbTab & "434"
        .AddItem "34" & vbTab & "TOTAL IMPUESTO GENERADO            441+442+443-444-445+447-452-453" & vbTab & vbTab & "" & vbTab & vbTab & vbTab & vbTab & vbTab & "454"
        .AddItem "35" & vbTab & "NOTAS DE CREDITO POR TRANSFERENCIAS NETAS OBJETO DE IVA" & vbTab & vbTab & vbTab & "415" & vbTab & vbTab & "435" & vbTab & vbTab & "455"
        .AddItem "36" & vbTab & "INGRESO NETO POR CONCEPTO DE REEMBOLOSO GASTOS" & vbTab & vbTab & vbTab & "416" & vbTab & vbTab & "436" & vbTab & vbTab & "456"
        .AddItem "37" & vbTab & "IVA PRESUNTIVO DE SALAS DE JUEGO (BINGO-MECANICOS) Y OTROS JUEGOS DE AZAR" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "457"
        .AddItem "38" & vbTab & "LIQUIDACION DEL IMPUESTO AL VALOR AGRGADO (SEGUN ART.67 DE L.R.T.I)"
        .AddItem "39" & vbTab & vbTab & "TOTAL BASE IMP. 12% VTAS. CONTADO" & vbTab & vbTab & vbTab & "TOTAL BASE IMP. 12% VTAS. CREDITO" & vbTab & vbTab & "TOTAL IMPUESTO GENERADO 454" & vbTab & vbTab & "IMPUESTO A LIQUIDAR MES ANTEIOR"
        .AddItem "40" & vbTab & "480" & vbTab & vbTab & vbTab & "481" & vbTab & vbTab & "482" & vbTab & vbTab & "483"
        .AddItem "41" & vbTab & vbTab & "IMP. A LIQUIDAR ESTE MES" & vbTab & vbTab & vbTab & "IMP. LIQUIDAR PROX. MES" & vbTab & vbTab & vbTab & vbTab & "TOTAL IMP. LIQUIDAR ESTE MES"
        .AddItem "42" & vbTab & "484" & vbTab & vbTab & vbTab & "485" & vbTab & vbTab & vbTab & vbTab & "499"
        .AddItem "43" & vbTab & "TRANSFERENCIAS NO OBJETO DEL IMPUESTO AL VALOR AGREGADO" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "BASE NO OBJETO IVA"
        .AddItem "44" & vbTab & "TRANSFERENCIAS NETAS NO OBJETO DE IVA" & vbTab & vbTab & vbTab & vbTab & vbTab & "460"
        .AddItem "45" & vbTab & "TOTAL COMPROBANTES DE VENTA EMITIDAS" & vbTab & vbTab & vbTab & "107"
        'COMPRAS
        .AddItem "46" & vbTab & "500" & vbTab & "RESUMEN DE ADQUICICIONES DEL PERIODO QUE DECLARA"
        .AddItem "47" & vbTab & "ADQUICICIONES OBJETO DEL IMPUESTO AL VALOR AGRGADO" & vbTab & vbTab & vbTab & "BASE IMPONIBLE 0%" & vbTab & vbTab & vbTab & "BASE IMPONIBLE 12%" & vbTab & "IMPUESTO"
        .AddItem "48" & vbTab & "COMPRAS LOCALES NETAS DE BIENES ( EXCLUYE ACT. FIJOS )" & vbTab & vbTab & vbTab & "501" & vbTab & vbTab & "521" & vbTab & vbTab & "541"
        .AddItem "49" & vbTab & "COMPRAS LOCALES DE ACTIVOS FIJOS" & vbTab & vbTab & vbTab & "502" & vbTab & vbTab & "522" & vbTab & vbTab & "542"
        .AddItem "50" & vbTab & "PAGOS DE SERVICIOS LOCALES Y DEL EXTERIOR" & vbTab & vbTab & vbTab & "502" & vbTab & vbTab & "522" & vbTab & vbTab & "542"
        .AddItem "51" & vbTab & "IMPORTACIONES DE BIENES (EXCLUYE ACT. FIJO)" & vbTab & vbTab & vbTab & "504" & vbTab & vbTab & "524" & vbTab & vbTab & "544"
        .AddItem "52" & vbTab & "IMPORTACIONES DE ACTIVOS FIJOS" & vbTab & vbTab & vbTab & "505" & vbTab & vbTab & "525" & vbTab & vbTab & "545"
        .AddItem "53" & vbTab & "IVA SOBRE VALOR DE LA DEPRECIACION DE ACTIVOS TEMPORAL" & vbTab & vbTab & vbTab & vbTab & vbTab & "526" & vbTab & vbTab & "546"
        .AddItem "54" & vbTab & "IVA EN ARRENDAMIENTO MERCANTIL INTERNACIONAL" & vbTab & vbTab & vbTab & vbTab & vbTab & "527" & vbTab & vbTab & "547"
        .AddItem "55" & vbTab & "NOTAS DE CREDITO POR ADQUISICIONES OBJETO IVA MES ACTUAL" & vbTab & vbTab & vbTab & "508" & vbTab & vbTab & "528" & vbTab & vbTab & "548"
        .AddItem "56" & vbTab & "NOTAS DE CREDITO POR ADQUISICIONES OBJETO IVA MES ANTERIOR" & vbTab & vbTab & vbTab & "509" & vbTab & vbTab & "529" & vbTab & vbTab & "549"
        .AddItem "57" & vbTab & "TOTAL ADQ.  NETAS OBJETO DE IVA E IMPORT NETAS"
        .AddItem "58" & vbTab & "BASE IMPONIBLE 0% (SUMAR DE 501 AL 505) - 508 - 509" & vbTab & vbTab & vbTab & "510"
        .AddItem "59" & vbTab & "BASE IMPONIBLE 12% (SUMAR DE 521 AL 527) - 528 - 529" & vbTab & vbTab & vbTab & vbTab & vbTab & "530"
        .AddItem "60" & vbTab & "IMPUESTO (SUMAR DEL 541 AL 547) - 548 - 549 " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "550"
        .AddItem "61" & vbTab & "NOTAS DE CREDITO POR ADQUISICIONES OBJETO IVA PROX. MES" & vbTab & vbTab & vbTab & "511" & vbTab & vbTab & "531" & vbTab & vbTab & "551"
        .AddItem "62" & vbTab & "PAGO NETO POR CONCEPTO DE REEMBOLSO DE GASTOS INTERMEDIADO" & vbTab & vbTab & vbTab & "512" & vbTab & vbTab & "532" & vbTab & vbTab & "552"
        .AddItem "63" & vbTab & "COMPRAS NETAS O BIENES O SERVICIOS QUE NO DAN DERECHO A CREDITO TRIBUTARUO" & vbTab & vbTab & vbTab & "513" & vbTab & vbTab & "533" & vbTab & vbTab & "553"
        .AddItem "64" & vbTab & "ADQUISICIONES NO OBJETO DEL IMPUESTO AL VALOR AGREGADO" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "BASE NO OBJETO IVA"
        .AddItem "65" & vbTab & "ADQUISICIONES NETAS NO OBJETO DE IVA" & vbTab & vbTab & vbTab & vbTab & vbTab & "560"
        .AddItem "66" & vbTab & "CREDITO TRIBUTARIO DE ACUERDO A CONTABILIDAD O A REGISTRO DE INGRESOS Y GASTO" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "554"
        .AddItem "67" & vbTab & "CREDITO TRIBUTARIO DE ACUERDO AL FACTOR DE PROPORCIONALIDAD" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "555"

'        'RESUMEN
        .AddItem "68" & vbTab & "600" & vbTab & "RESUMEN IMPOSITIVO"
        .AddItem "69" & vbTab & "IMPUESTO CAUSADO" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "499-554 ó 555>0" & vbTab & "601"
        .AddItem "70" & vbTab & "(-) CREDITO TRIBUTARIO DEL MES " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "499-554 ó 499<0" & vbTab & "602"
        .AddItem "71" & vbTab & "(-) SALDO DE CREDITO TRIBUTARIO APLICARSE EN ESTE MES " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "Trasladar campo 399" & vbTab & "603"
        .AddItem "72" & vbTab & "(-) RETENCIONES EN LA FUENTE DE IVA QUE LE HAN SIDO EFECTUADAS " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "604"
        .AddItem "73" & vbTab & "(=) SALDO DE CREDITO TRIBUTARIO PARA EL PROXIMO MES " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "601-602-603-604 < 0" & vbTab & "698"
        .AddItem "74" & vbTab & "(=) SUBTOTAL A PAGAR " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "601-602-603-604 > 0" & vbTab & "699"
 '        'DECLARACION DEL SUJETO PASIVO COMO AGENTE RETENCION IVA
        .AddItem "75" & vbTab & "700" & vbTab & "DECLARACION DEL SUJETO PASIVO COMO AGENTE RETENCION IVA" & vbTab & vbTab & vbTab & vbTab & vbTab & "VALOR DEL IVA" & vbTab & vbTab & "VALOR RETENIDO"
        .AddItem "76" & vbTab & "IVA POR LA PRESTACION DE SERVICIOS PROFESIONALES" & vbTab & vbTab & vbTab & vbTab & vbTab & "701" & vbTab & vbTab & "721"
        .AddItem "77" & vbTab & "IVA POR EL ARRENDAMIENTO DE INMUEBLES A PERSONAS NATURALES" & vbTab & vbTab & vbTab & vbTab & vbTab & "702" & vbTab & vbTab & "722"
        .AddItem "78" & vbTab & "IVA EN OTRAS COMPRA DE BIENES Y SERVICIOS CON EMISION DE LIQUIDACION DE COMPRAS Y PRESTACION SERVICIOS" & vbTab & vbTab & vbTab & vbTab & vbTab & "703" & vbTab & vbTab & "723"
        .AddItem "79" & vbTab & "IVA EN LA DEPRECIACION DE ACTIVOS EN INTERNACION TEMPORAL" & vbTab & vbTab & vbTab & vbTab & vbTab & "704" & vbTab & vbTab & "724"
        .AddItem "80" & vbTab & "IVA EN LA DISTRIBUCION DE COMBUSTIBLES" & vbTab & vbTab & vbTab & vbTab & vbTab & "705" & vbTab & vbTab & "725"
        .AddItem "81" & vbTab & "IVA EN LEASING INTERNACIONAL" & vbTab & vbTab & vbTab & vbTab & vbTab & "706" & vbTab & vbTab & "726"
        .AddItem "82" & vbTab & "IVA EN OPERACIONES REALIZADA POR EXPORTADORES" & vbTab & vbTab & vbTab & vbTab & vbTab & "707" & vbTab & vbTab & "727"
        .AddItem "83" & vbTab & "IVA POR LA PRESTACION DE SERVICIOS" & vbTab & vbTab & vbTab & vbTab & vbTab & "708" & vbTab & vbTab & "728"
        .AddItem "84" & vbTab & "IVA RETENIDO POR EMISORAS DE TARJETAS DE CREDITO SERVICIOS" & vbTab & vbTab & vbTab & vbTab & vbTab & "709" & vbTab & vbTab & "729"
        .AddItem "85" & vbTab & "IVA RETENIDO POR EMISORAS DE TARJETAS DE CREDITO BIENES" & vbTab & vbTab & vbTab & vbTab & vbTab & "710" & vbTab & vbTab & "730"
        .AddItem "86" & vbTab & "IVA POR LA COMPRA DE BIENES" & vbTab & vbTab & vbTab & vbTab & vbTab & "711" & vbTab & vbTab & "731"
        .AddItem "87" & vbTab & "IVA EN CONTRATOS DE CONSTRUCCION" & vbTab & vbTab & vbTab & vbTab & vbTab & "712" & vbTab & vbTab & "732"
        .AddItem "88" & vbTab & "COMPROBANTES DE RETENCION EMITIDOS" & vbTab & vbTab & vbTab & "118"
        .AddItem "89" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "TOTAL IVA RETENIDO " & vbTab & vbTab & "798"
        .AddItem "90" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "TOTAL IVA A PAGAR " & vbTab & vbTab & "799"
'        'A PAGAR
        .AddItem "91" & vbTab & "Declaro que los datos contenidos en esta declaración son verdaderos por lo que asumo la responsabilidad correspondiente (Artículo 98 de la L.R.T.I.) " & vbTab & vbTab & vbTab & vbTab & vbTab & "900" & vbTab & "VALORES A PAGAR Y FORMA DE PAGO"
        .AddItem "92"
        .AddItem "93" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "PAGO PREVIO" & vbTab & vbTab & "901"
        .AddItem "94" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "TOTAL IMPUESTO A PAGAR 899-901" & vbTab & vbTab & "902"
        .AddItem "95" & vbTab & vbTab & "____________________________" & vbTab & vbTab & vbTab & "____________________________" & vbTab & "INTERESES POR MORA" & vbTab & vbTab & "903"
        .AddItem "96" & vbTab & vbTab & "FIRMA SUJETO PASIVO" & vbTab & vbTab & vbTab & "FIRMA CONTADOR" & vbTab & "MULTAS" & vbTab & vbTab & "904"
        .AddItem "97" & vbTab & "NOMBRE:" & vbTab & vbTab & "NOMBRE:" & vbTab & "" & vbTab & vbTab & "TOTAL PAGADO     902+903+904" & vbTab & vbTab & "999"
        .AddItem "98" & vbTab & "198" & vbTab & "C.I. No." & vbTab & "199" & vbTab & "RUC No." & vbTab & ""
        .AddItem "99" & vbTab & "MEDIANTE CHEQUE DEBITO BANCARIO EFECTIVO U OTRAS FORMAS DE COBRO" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "905"
        .AddItem "100" & vbTab & "MEDIANTE COMPENSACIONES" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "906"
        .AddItem "101" & vbTab & "MEDIANTE NOTAS DE CREDITO" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "907"
        .AddItem "102" & vbTab & vbTab & "DETALLE DE NOTAS DE CREDITO" & vbTab & vbTab & vbTab & vbTab & "DETALLE DE NOTAS DE COMPENSACIONES"
        .AddItem "103" & vbTab & "908" & vbTab & "N/C No." & vbTab & "909" & vbTab & "USD" & vbTab & vbTab & "916" & vbTab & "Resol No." & vbTab & "917" & vbTab & "USD"
        .AddItem "104" & vbTab & "910" & vbTab & "N/C No." & vbTab & "911" & vbTab & "USD" & vbTab & vbTab & "918" & vbTab & "Resol No." & vbTab & "919" & vbTab & "USD"
        .AddItem "105" & vbTab & "912" & vbTab & "N/C No." & vbTab & "913" & vbTab & "USD"
        .AddItem "106" & vbTab & "914" & vbTab & "N/C No." & vbTab & "915" & vbTab & "USD"
        .Redraw = flexRDBuffered
        .Refresh
    End With
    CambiaFondoCeldasEditables104
End Sub

Private Sub LlenaFormatoFormulario104_2010()
    With grd
        
        .MergeCells = flexMergeSpill

        .AddItem "1" & vbTab & "Formulario" & vbTab & vbTab & " DECLARACION DEL IMPUESTO AL VALOR AGREGADO"
        .AddItem "2" & vbTab & "    104"
        .AddItem "3" & vbTab & "Resolucion No "
        .AddItem "4" & vbTab & "NAG-DGER2008-1520 "
        .AddItem "5" & vbTab & "________________________________________________________________________________________________________________________________________________________________________________________________"
        .AddItem "6" & vbTab & "100 IDENTIFICACION DE LA DECLARACION"
        .AddItem "7" & vbTab & "MES 101" & vbTab & vbTab & "AÑO 102" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "(O) ORIGINAL - (S) SUSTITUTIVA   031"
        .AddItem "8" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "No FORMULARIO QUE SUSTITUYE 104"
        .AddItem "9" & vbTab & "200 IDENTIFICACION DEL SUJETO PASIVO"
        .AddItem "10" & vbTab & "RUC 201" & vbTab & vbTab & "202"
        .AddItem "11"
        .AddItem "12" & vbTab & "RESUMEN DE VENTAS Y OTRAS OPERACIONES DEL PERIODO QUE DECLARA" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "Valor Bruto" & vbTab & vbTab & vbTab & "Valor Neto" & vbTab & vbTab & "Impuesto Generado"
        .AddItem "13" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "Valor Bruto-N/C"
        .AddItem "14" & vbTab & "Ventas locales  (excluye activos fijos) gravadas tarifa 12%" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 401" & vbTab & vbTab & "411" & vbTab & vbTab & "421"
        .AddItem "15" & vbTab & "Ventas de activos fijos grabados  Tarifa 12%" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 402" & vbTab & vbTab & "412" & vbTab & vbTab & "422"
        .AddItem "16" & vbTab & "Ventas locales (excluye activos fijos) gravadas tarifa 0% que no dan derecho a credito tributario" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 403" & vbTab & vbTab & "413"
        .AddItem "17" & vbTab & "Ventas de activos fijos gravados tarifa 0% que no dan derecho a credito tributario" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 404" & vbTab & vbTab & "414"
        .AddItem "18" & vbTab & "Ventas locales (excluye activos fijos) gravadas tarifa 0% que  dan derecho a credito tributario" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 405" & vbTab & vbTab & "415"
        .AddItem "19" & vbTab & "Ventas de activos fijos gravadas tarifa 0% que dan derecho a credito tributario" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 406" & vbTab & vbTab & "416"
        .AddItem "20" & vbTab & "Exportaciones de bienes" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 407" & vbTab & vbTab & "417"
        .AddItem "21" & vbTab & "Exportaciones de Servicios" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 408" & vbTab & vbTab & "418"
        .AddItem "22" & vbTab & "TOTAL VENTAS Y OTRAS OPERACIONES" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 409" & vbTab & vbTab & "419" & vbTab & vbTab & "429"
        .AddItem "23" & vbTab & "Transferencias no objeto o excento de IVA" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 431" & vbTab & vbTab & " 441"
        .AddItem "24" & vbTab & "Notas de crédito tarifa 0%  por compensar proximo mes (informativo)" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 442"
        .AddItem "25" & vbTab & "Notas de crédito tarifa 12%  por compensar proximo mes (informativo)" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 443" & vbTab & vbTab & " 453"
        .AddItem "26" & vbTab & "Ingresos por reembolso como intermediario (informativo)" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 434" & vbTab & vbTab & " 444" & vbTab & vbTab & " 454"
        .AddItem "27"
        .AddItem "28" & vbTab & "LIQUIDACION DEL IVA EN EL MES"
        .AddItem "29" & vbTab & "Total Transferencias" & vbTab & vbTab & "Total Transferencias" & vbTab & vbTab & "Total Impuesto" & vbTab & vbTab & "Impuesto a liquidar" & vbTab & vbTab & "Impuesto a liquidar" & vbTab & vbTab & "Impuesto a liquidar" & vbTab & vbTab & "Total Impuesto a"
        .AddItem "30" & vbTab & "Gravadas 12% a" & vbTab & vbTab & "Gravadas 12% a" & vbTab & vbTab & "Generado" & vbTab & vbTab & "del mes anterior" & vbTab & vbTab & "en este mes" & vbTab & vbTab & "en el proximo mes " & vbTab & vbTab & "Liquidar en este mes"
        .AddItem "31" & vbTab & "contado este mes" & vbTab & vbTab & "credito este mes% a" & vbTab & vbTab & "(trasladese campo  429)" & vbTab & vbTab & "(campo 485 periodo ant)" & vbTab & vbTab & "(Min. 12%  campo 480)" & vbTab & vbTab & "(482-484) " & vbTab & vbTab & "(483+484)"
        .AddItem "32" & vbTab & "480" & vbTab & vbTab & "481" & vbTab & vbTab & "482" & vbTab & vbTab & "483" & vbTab & vbTab & "484" & vbTab & vbTab & "485 " & vbTab & vbTab & "499"
        .AddItem "33"
        .AddItem "34" & vbTab & "RESUMEN DE ADQUISICIONES Y PAGOS DEL PERIODO QUE DECLARA" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "Valor Bruto" & vbTab & vbTab & vbTab & "Valor Neto" & vbTab & vbTab & "Impuesto Generado"
        .AddItem "35" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "Valor Bruto-N/C"
        
        .AddItem "36" & vbTab & "Adquisiciones y pagos (excluye activos fijos) gravados tarifa 12% (con derecho a crédito tributario)" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 500" & vbTab & vbTab & "510" & vbTab & vbTab & "520"
        .AddItem "37" & vbTab & "Adquisiciones locales de activos fijos gravados tarifa 12% (con derecho a crédito tributario)" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 501" & vbTab & vbTab & "511" & vbTab & vbTab & "521"
        .AddItem "38" & vbTab & "Otras adquisiciones y pagos gravados tarifa 12% (sin derecho a crédito tributario)" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 502" & vbTab & vbTab & "512" & vbTab & vbTab & "522"
        .AddItem "39" & vbTab & "Importaciones de servicios gravados tarifa 12% " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 503" & vbTab & vbTab & "513" & vbTab & vbTab & "523"
        .AddItem "40" & vbTab & "Importaciones de bienes (excluye activos fijos) gravados tarifa 12%" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 504" & vbTab & vbTab & "514" & vbTab & vbTab & "524"
        .AddItem "41" & vbTab & "Importaciones de activos fijos gravados tarifa 12%" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 505" & vbTab & vbTab & "515" & vbTab & vbTab & "525"
        .AddItem "42" & vbTab & "Importaciones de bienes(incluye activos fijos) gravados tarifa 0%" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 506" & vbTab & vbTab & "516"
        .AddItem "43" & vbTab & "Adquisiciones y pagos (incluye activos fijos) gravados tarifa 0% " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 507" & vbTab & vbTab & "517"
        .AddItem "44" & vbTab & "Adquisiciones realizadas a contribuyentes RISE " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 508" & vbTab & vbTab & "518"
        .AddItem "45" & vbTab & "TOTAL ADQUISIONES Y PAGOS" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 509" & vbTab & vbTab & "519" & vbTab & vbTab & "529"
        .AddItem "46" & vbTab & "Adquisiciones no objeto de IVA " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 531" & vbTab & vbTab & "541"
        .AddItem "47" & vbTab & "Adquisiciones excentas pago de IVA " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 532" & vbTab & vbTab & "542"
        .AddItem "48" & vbTab & "Notas de crédito tarifa 0% por compensar proximo mes (informativo) " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "543"
        .AddItem "49" & vbTab & "Notas de crédito tarifa 12% por compensar proximo mes (informativo) " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "544" & vbTab & vbTab & "554"
        .AddItem "50" & vbTab & "Pagos netos por reembolso con intermediario(informativo) " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & " 535" & vbTab & vbTab & "545" & vbTab & vbTab & "555"
        .AddItem "51"
        .AddItem "52" & vbTab & "Factor de proporcionalidad para credito tributario (411+412+415+416+417+418)/419" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "563"
        .AddItem "53" & vbTab & "Crédito tributario aplicable a este periodo (De acuerdo a factor de proporcionalidad a su contabilidad) (520+521+523+524+525)x563" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "564"
        .AddItem "54"
        .AddItem "55" & vbTab & "RESUMEN IMPOSITIVO: AGENTE DE PERCEPCION DEL IMPUESTO AL VALOR AGREGADO"
        .AddItem "56" & vbTab & "Impuesto causado (Si 499-564 es mayor a cero)" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "601"
        .AddItem "57" & vbTab & "Crédito tributario aplicable a este periodo (Si 499-564 es menor a cero)" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "602"
        .AddItem "58" & vbTab & "(-) Saldo credito por adquisiones e importaciones (traslade al campo 615 de la declaracion del periodo anterior)" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "605"
        .AddItem "59" & vbTab & "Tributario del  por retenciones en la fuente de IVA que le han sido"
        .AddItem "60" & vbTab & "mes anterior efecutados(Traslade al campo 617  de la declaracion del periodo anterior)" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "607"
        .AddItem "61" & vbTab & "(-) Retenciones en la fuente de IVA que han sido efectuadas en este periodo" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "609"
        .AddItem "62" & vbTab & "Ajuste por IVA devuelto e IVA rechazado imputable al credito tributario en el mes" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "611"
        .AddItem "63" & vbTab & "Ajuste por IVA devuelto por otras instituciones del sector publico imputable al credito tributario en el mes" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "613"
        .AddItem "64" & vbTab & "Saldo crédito tributario para el proximo mes/Por adquisiones e importaciones" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "615"
        .AddItem "65" & vbTab & "Saldo crédito tributario para el proximo mes/Por retenciones en la fuente de IVA que han sido efectuadas" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "617"
        .AddItem "66" & vbTab & "SUBTOTAL A PAGAR (Si 601-602-605-607-609+611 es mayor que 0)" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "619"
        .AddItem "67" & vbTab & "IVA presuntivo del salas de juego (bingo mecanismos) y otros juegos al azar" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "621"
        .AddItem "68" & vbTab & "TOTAL IMPUESTO A PAGAR POR RECEPCION (619+621)" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "699"
        .AddItem "69"
        .AddItem "70"
        .AddItem "71" & vbTab & "AGENTE DE RETENCION DEL IMPUESTO AL VALOR AGREGADO"
        .AddItem "72" & vbTab & "Retencion de 30%" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "721"
        .AddItem "73" & vbTab & "Retencion de 70%" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "723"
        .AddItem "74" & vbTab & "Retencion de 100%" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "725"
        .AddItem "75" & vbTab & "TOTAL IMPUESTO A PAGAR POR RETENCION (721+723+725 )" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "799"
        .AddItem "76"
        .AddItem "77" & vbTab & "TOTAL CONSOLIDADO  DEL IMPUESTO AL VALOR AGREGADO(699+799 )" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "859"
        .AddItem "78"
        .AddItem "79" & vbTab & "Pago previo (informativo)" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "890"
        .AddItem "80" & vbTab & "DETALLE DE IMPUTACION AL PAGO(para declaraciones sustitutivas)"
        .AddItem "81" & vbTab & vbTab & "Interes 897" & vbTab & vbTab & "Impuesto 898 " & vbTab & vbTab & vbTab & vbTab & "Multa 899" & vbTab & vbTab & vbTab & "899"
        .AddItem "82" & vbTab & "PAGO DIRECTO EN CUENTA UNICA DEL TESORO NACIONAL (Uso exclusivo para instituciones y Empresas del sector publico autorizadas" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "880"
        
        .AddItem "83" & vbTab & "VALORES A PAGAR Y FORMA DE PAGO (Luego de imputacion  al pago en declaraciones sustitutivas)"
        .AddItem "84" & vbTab & "Total Impuesto a pagar (859+897)" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "902"
        .AddItem "85" & vbTab & "Interes por mora" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "903"
        .AddItem "86" & vbTab & "Multas" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "904"
        .AddItem "87" & vbTab & "TOTAL PAGADO (902+903+904)" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "999"
        .AddItem "88"
        .AddItem "89" & vbTab & "Mediante cheque, débito bancario,efectivo u otras formas de pago" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "905"
        .AddItem "90" & vbTab & "Mediante compensaciones" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "906"
        .AddItem "91" & vbTab & "Mediante notas de crédito" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "907"
        .AddItem "92"
        .AddItem "93" & vbTab & "DETALLE NOTAS DE CREDITO"
        .AddItem "94" & vbTab & vbTab & "N/C No.908 " & vbTab & vbTab & "N/C No.910 " & vbTab & vbTab & "N/C No.912 " & vbTab & vbTab & vbTab & vbTab & "N/C No.914 "
        .AddItem "95" & vbTab & vbTab & "Valor USD 909 " & vbTab & vbTab & "Valor  USD 911 " & vbTab & vbTab & "Valor USD 913 " & vbTab & vbTab & vbTab & vbTab & "Valor USD 915 "
        .AddItem "96"
        .AddItem "97" & vbTab & "DETALLE DE COMPENSACIONES" & vbTab & vbTab & vbTab & vbTab & "Resolucion 916" & vbTab & vbTab & "Resolucion 918"
        .AddItem "98" & vbTab & vbTab & vbTab & vbTab & vbTab & "Valor USD 917" & vbTab & vbTab & "Valor  USD 919"
        .AddItem "99"
        .AddItem "100" & vbTab & "Declaro que los datos proporcionados en este documento son exactos y verdaderos, por lo que asumo la responsabilidad legal que de ella se deriven (Art. 101 de la L.O.R.T.I.)"
        .AddItem "101"
        .AddItem "102"
        .AddItem "103" & vbTab & vbTab & "No. ID SUJETO PASIVO/REP. LEGAL 198" & vbTab & vbTab & vbTab & vbTab & vbTab & "RUC CONTADOR 199 "
        .AddItem "104"
        .AddItem "105" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "FORMA DE PAGO 921 "
        .AddItem "106" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "BANCO 922 "
        .AddItem "107"
        .AddItem "108"
        .Redraw = flexRDBuffered
        .Refresh
    End With
    CambiaFondoCeldasEditables104_2010
End Sub

Private Sub CambiaFondoCeldasEditables104_2010()
    Dim fmt As String
    With grd
        .Redraw = flexRDBuffered
        'fondo blanco
        .BackColor = &HE0E0E0      '&H80000013
        .Cell(flexcpBackColor, 7, 2, 7, 2) = &H800000
        .Cell(flexcpBackColor, 7, 4, 7, 4) = &H800000
        .Cell(flexcpBackColor, 7, 13, 7, 13) = &H80000014 'tambien es blanco
        .Cell(flexcpBackColor, 7, 14, 7, 14) = vbWhite
        .Cell(flexcpBackColor, 8, 13, 8, 13) = vbWhite
        .Cell(flexcpBackColor, 8, 14, 8, 14) = &H80000014 'tambien es blanco
        .Cell(flexcpBackColor, 10, 2, 10, 2) = &H800000
        .Cell(flexcpBackColor, 10, 4, 10, 6) = &H800000
        
        .Cell(flexcpBackColor, 14, 8, 23, 8) = vbWhite
        .Cell(flexcpBackColor, 14, 10, 26, 10) = vbWhite
        .Cell(flexcpBackColor, 14, 12, 15, 12) = vbWhite
        .Cell(flexcpBackColor, 22, 12, 22, 12) = vbWhite
        .Cell(flexcpBackColor, 25, 12, 26, 12) = vbWhite
        .Cell(flexcpBackColor, 26, 8, 26, 8) = vbWhite
        '.Cell(flexcpBackColor, 22, 10, 22, 10) = vbWhite
        .Cell(flexcpBackColor, 8, 10, 8, 10) = vbWhite
        .Cell(flexcpBackColor, 23, 10, 26, 10) = vbWhite
        .Cell(flexcpBackColor, 32, 2, 32, 2) = vbWhite
        .Cell(flexcpBackColor, 32, 4, 32, 4) = vbWhite
        .Cell(flexcpBackColor, 32, 6, 32, 6) = vbWhite
        .Cell(flexcpBackColor, 32, 8, 32, 8) = vbWhite
        .Cell(flexcpBackColor, 32, 10, 32, 10) = vbWhite
        .Cell(flexcpBackColor, 32, 12, 32, 12) = vbWhite
        .Cell(flexcpBackColor, 32, 14, 32, 14) = vbWhite
        .Cell(flexcpBackColor, 36, 8, 47, 8) = vbWhite
        .Cell(flexcpBackColor, 36, 10, 43, 10) = vbWhite
        .Cell(flexcpBackColor, 36, 12, 41, 12) = vbWhite
        .Cell(flexcpBackColor, 44, 8, 44, 8) = vbWhite
        .Cell(flexcpBackColor, 44, 10, 44, 10) = vbWhite
        .Cell(flexcpBackColor, 45, 12, 45, 12) = vbWhite
        .Cell(flexcpBackColor, 45, 10, 49, 10) = vbWhite
        .Cell(flexcpBackColor, 50, 8, 50, 8) = vbWhite
        
        .Cell(flexcpBackColor, 50, 10, 50, 10) = vbWhite
        .Cell(flexcpBackColor, 49, 12, 50, 12) = vbWhite
        
        .Cell(flexcpBackColor, 52, 12, 53, 12) = vbWhite
        .Cell(flexcpBackColor, 56, 12, 58, 12) = vbWhite
        .Cell(flexcpBackColor, 57, 12, 57, 12) = vbWhite
        .Cell(flexcpBackColor, 60, 12, 68, 12) = vbWhite
       ' .Cell(flexcpBackColor, 65, 12, 65, 12) = vbWhite
        '.Cell(flexcpBackColor, 67, 12, 67, 12) = vbWhite
        
        .Cell(flexcpBackColor, 66, 12, 66, 12) = vbWhite
        .Cell(flexcpBackColor, 72, 12, 75, 12) = vbWhite
       ' .Cell(flexcpBackColor, 74, 12, 74, 12) = vbWhite
       ' .Cell(flexcpBackColor, 76, 12, 76, 12) = vbWhite '&H800000
        .Cell(flexcpBackColor, 77, 12, 77, 12) = vbWhite
        .Cell(flexcpBackColor, 79, 12, 79, 12) = vbWhite
        .Cell(flexcpBackColor, 81, 3, 81, 3) = vbWhite
        .Cell(flexcpBackColor, 81, 6, 81, 6) = vbWhite
        .Cell(flexcpBackColor, 81, 10, 81, 10) = vbWhite
        .Cell(flexcpBackColor, 81, 12, 82, 12) = vbWhite
        .Cell(flexcpBackColor, 81, 12, 81, 12) = vbWhite
        
       ' .Cell(flexcpBackColor, 83, 12, 83, 12) = vbWhite
        
        .Cell(flexcpBackColor, 84, 12, 87, 12) = vbWhite
        
        '.Cell(flexcpBackColor, 88, 12, 88, 12) = vbWhite
        .Cell(flexcpBackColor, 89, 12, 91, 12) = vbWhite
        .Cell(flexcpBackColor, 94, 3, 95, 3) = vbWhite
        .Cell(flexcpBackColor, 94, 5, 95, 5) = vbWhite
        .Cell(flexcpBackColor, 94, 8, 95, 8) = vbWhite
        .Cell(flexcpBackColor, 94, 12, 95, 12) = vbWhite
        .Cell(flexcpBackColor, 97, 6, 97, 6) = vbWhite
        .Cell(flexcpBackColor, 98, 6, 98, 6) = vbWhite
        .Cell(flexcpBackColor, 97, 10, 98, 10) = vbWhite
        .Cell(flexcpBackColor, 103, 5, 103, 5) = vbWhite
        .Cell(flexcpBackColor, 103, 9, 104, 10) = vbWhite
        .Cell(flexcpBackColor, 105, 8, 105, 8) = vbWhite
        
        

        .Cell(flexcpForeColor, 10, 2, 10, 2) = vbYellow
        .Cell(flexcpForeColor, 10, 4, 10, 4) = vbYellow
        .Cell(flexcpForeColor, 7, 2, 7, 2) = vbYellow
        .Cell(flexcpForeColor, 7, 4, 7, 4) = vbYellow
        
        .Cell(flexcpForeColor, 29, 5, 32, 5) = vbRed
        .Cell(flexcpForeColor, 29, 11, 32, 11) = vbRed
        .Cell(flexcpForeColor, 29, 13, 32, 13) = vbRed
        
        .Cell(flexcpFontSize, 1, 1, 1, 1) = 11
        .Cell(flexcpFontSize, 1, 3, 1, 3) = 14
        .Cell(flexcpFontSize, 2, 1, 2, 1) = 13
        .Cell(flexcpFontSize, 29, 1, 31, 14) = 8
        'COLOR AMARILLO PARA LOS NO EDITABLES
'       .Cell(flexcpForeColor, 14, 12, 15, 12) = &H000000C0&
        
        .Cell(flexcpFontBold, 2, 1, 2, 1) = True
        .RowHeight(1) = 400
        .RowHeight(2) = 400
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .MergeCol(4) = True
        .MergeCol(5) = True
    End With
End Sub

Private Sub Buscar1042010()
    Dim sql As String, cond As String, CadenaValores As String
    Dim OrdenadoX As String
    Dim CadenaAgrupa  As String, Recargo As String
    Dim v As Variant, max As Integer, i As Integer
    Dim from As String, NumReg As Long, f1 As String
    Dim rs As Recordset
    Dim SubTotal As Currency, CompraBienesTarifa_0 As Currency, CompraServiciosTarifa_0 As Currency
    Dim CompraActivosTarifa_0 As Currency
    Dim CP_Ser As String, CP_Act As String, CP_Dev As String
    Dim VT_Bie As String, VT_Ser As String, VT_Dev As String
    Dim VT_ExpBie As String, VT_ExpSer As String, VT_RepGas As String
    Dim NC_Ventas As String, NC_Compras As String, CP_Bie As String
    Dim Reten As String, ret_real As String, ret_recib As String
    Dim CP_Rise As String, DescNC As Currency
    Dim Moneda As String
    Dim VentaNeta  As Currency
    Dim NCredito As Currency, NCTrans As Currency, rsTrans As Recordset, trans As Currency
    Dim ventasSinIva As Currency
    Dim CompraActivos As Currency
    Dim rsd As Recordset
    Dim VentaNetaExp As Currency
    Dim ventasSinIvaExp As Currency
    Dim NCreditoExp As Currency
    Dim NCreditoExpB As Currency
    Dim NCreditoExpS As Currency
    Dim NC_ExpB As String, NC_ExpS As String
    
    Dim Compras12 As Currency
    Dim Compras0 As Currency
    Dim ComprasNoIva As Currency
    
    Dim NCCompras0 As Currency
    Dim NCCompras12 As Currency
    Dim NCComprasNoIVA As Currency
    
    Set objcond = gobjMain.objCondicion
    If Not frmB_FormSRI.Inicio104_2013(objcond, Recargo, CP_Ser, CP_Act, CP_Dev, VT_Dev, VT_Bie, VT_Ser, _
                                                    VT_ExpBie, VT_ExpSer, VT_RepGas, ret_real, ret_recib, Reten, NC_ExpB, NC_ExpS, NC_Ventas, NC_Compras, CP_Bie, CP_Rise) Then
        grd.SetFocus
        Exit Sub
    End If
    With objcond
        If Len(Month(.fecha1)) < 2 Then
            grd.TextMatrix(7, 2) = Format(.fecha1, "MMMM")
        Else
            grd.TextMatrix(7, 2) = Format(.fecha1, "MMMM")
        End If
        grd.TextMatrix(7, 4) = Year(.fecha1)
        'Reporte de un mes a la vez
        f1 = DateSerial(Year(.fecha1), Month(.fecha1), 1)
        cond = " AND GNC.FechaTrans BETWEEN " & FechaYMD(f1, gobjMain.EmpresaActual.TipoDB) & _
               " AND " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB)
            'VENTAS BIENES
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 "
            sql = sql & "From IvRecargo ivR inner join "
            sql = sql & " IvKardexRecargo ivkR Inner join "
            sql = sql & " GnComprobante gnc Inner join PcPRovCLi on gnc.IdClienteRef = PCProvCli.IdProvCli "
            sql = sql & " On ivkr.TransID = gnc.TransID "
            sql = sql & " On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(VT_Bie) & ")  "
            sql = sql & " and bandCierre=0 "
            sql = sql & " AND ivr.CodRecargo IN (" & PreparaCadena(Recargo) & ") " & cond
            sql = sql & " Group by IvkR.TransID"
            VerificaExistenciaTabla 0
            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
            sql = "SELECT "
            sql = sql & " isnull(SUM(PrecioTotalBase0 *SignoVenta + (PrecioTotalBase0*SignoVenta * (cast(isnull(TotalDescuento,0) as float) *SignoVenta/ cast(PrecioTotal*SignoVenta as float)*SignoVenta))),0) as Base0, "
            sql = sql & " isnull(SUM(PrecioTotalBaseIVA * SignoVenta + (PrecioTotalBaseIVA *SignoVenta* (cast(isnull(TotalDescuento,0) as float) *SignoVenta/ cast(PrecioTotal as float)*SignoVenta))),0) As BaseIVA "
            sql = sql & " FROM tmp0 Right join "
            sql = sql & "vwConsSUMIVKardexIVA inner join "
            sql = sql & "GNComprobante GNC   "
            sql = sql & " INNER JOIN PCPROVCLI PC ON GNC.IDCLIENTEREF=PC.IDPROVCLI"
            sql = sql & " INNER JOIN GNTRANS gnt ON gnc.CODTRANS=GNT.CODTRANS"
            sql = sql & " ON vwConsSUMIVKardexIVA.TransID = GNC.TransID "
            sql = sql & " ON tmp0.TransID = GNC.TransID"
            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(VT_Bie) & ")  "
            sql = sql & " and bandCierre=0 "
            sql = sql & " AND GNT.AnexoCodTipoTrans=2 " & cond
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            '401
            grd.TextMatrix(14, 8) = Round(Abs(rs.Fields("BaseIVA")), 2) '+ trans 'carga 404
            '405
            grd.TextMatrix(18, 8) = Round(Abs(rs.Fields("Base0")), 2) '+ trans 'carga 404
            Set rs = Nothing
            
            
            Dim DESIVA As Currency
            
            'NOTAS DE CREDITO DE VENTAS NETAS
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp100 "
            sql = sql & "From IvRecargo ivR inner join "
            sql = sql & " IvKardexRecargo ivkR Inner join "
            sql = sql & " GnComprobante gnc Inner join PcPRovCLi on gnc.IdClienteRef = PCProvCli.IdProvCli "
            sql = sql & " On ivkr.TransID = gnc.TransID "
            sql = sql & " On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(NC_Ventas) & ")  "
            sql = sql & " and bandCierre=0 "
            sql = sql & " AND ivr.CodRecargo IN (" & PreparaCadena(Recargo) & ") " & cond
            sql = sql & " Group by IvkR.TransID"
            
            VerificaExistenciaTabla 100
            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
            
            
            
            
            sql = "SELECT "
            sql = sql & " isnull(SUM(PrecioTotalBase0 *SignoVenta + (PrecioTotalBase0*SignoVenta * (cast(isnull(TotalDescuento,0) as float) *SignoVenta/ cast(PrecioTotal*SignoVenta as float)*SignoVenta))),0) as NCBase0, "
            sql = sql & " isnull(SUM(PrecioTotalBaseIVA * SignoVenta + (PrecioTotalBaseIVA *SignoVenta* (cast(isnull(TotalDescuento,0) as float) *SignoVenta/ cast(PrecioTotal as float)*SignoVenta))),0) As NCBaseIVA "
            sql = sql & " FROM tmp100 Right join "
            sql = sql & "vwConsSUMIVKardexIVA inner join "
            sql = sql & "GNComprobante GNC   "
            sql = sql & " INNER JOIN PCPROVCLI PC ON GNC.IDCLIENTEREF=PC.IDPROVCLI"
            sql = sql & " INNER JOIN GNTRANS gnt ON gnc.CODTRANS=GNT.CODTRANS"
            sql = sql & " ON vwConsSUMIVKardexIVA.TransID = GNC.TransID "
            sql = sql & " ON tmp100.TransID = GNC.TransID"
            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(NC_Ventas) & ")  "
            sql = sql & " and bandCierre=0 "
            sql = sql & " AND GNT.AnexoCodTipoTrans=2 " & cond
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            grd.TextMatrix(14, 10) = grd.ValueMatrix(14, 8) - Round(Abs(rs.Fields("NCBaseIVA")), 2)
            '405
            grd.TextMatrix(18, 10) = grd.ValueMatrix(18, 8) - Round(Abs(rs.Fields("NCBase0")), 2)
            '------------EXPORTACION bienes
'            If Len(VT_ExpBie) > 0 Then
'            End If
'                sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 "
'                sql = sql & "From IvRecargo ivR inner join "
'                sql = sql & " IvKardexRecargo ivkR Inner join "
'                sql = sql & " GnComprobante gnc Inner join PcPRovCLi on gnc.IdClienteRef = PCProvCli.IdProvCli "
'                sql = sql & " On ivkr.TransID = gnc.TransID "
'                sql = sql & " On Ivr.IdRecargo = IvkR.IdRecargo "
'                sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(VT_ExpBie) & ")  "
'                sql = sql & " and bandCierre=9 "
'                sql = sql & " AND ivr.CodRecargo IN (" & PreparaCadena(Recargo) & ") " & Cond
'                sql = sql & " Group by IvkR.TransID"
'                VerificaExistenciaTabla 0
'                gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
'                sql = "Select  SUM(IvKr.Valor) as TotalDescuento "
'                sql = sql & "From IvRecargo ivR inner join "
'                sql = sql & " IvKardexRecargo ivkR Inner join "
'                sql = sql & " GnComprobante gnc Inner join PcPRovCLi on gnc.IdClienteRef = PCProvCli.IdProvCli "
'                sql = sql & " On ivkr.TransID = gnc.TransID "
'                sql = sql & " On Ivr.IdRecargo = IvkR.IdRecargo "
'                sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(VT_ExpBie) & ")  "
'                sql = sql & " and bandCierre=9 "
'                sql = sql & " AND ivr.CodRecargo IN ('TRANS') " & Cond
'                Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
'                If Not IsNull(rs!TotalDescuento) Then trans = Round(Abs(rs!TotalDescuento), 2)    '403
'                sql = "SELECT "
'                sql = sql & " isnull(SUM(PrecioTotalBase0 ),0) as Base0, "
'                sql = sql & " isnull(SUM(PrecioTotalBaseIVA ),0) As BaseIVA "
'                sql = sql & " FROM tmp0 Right join "
'                sql = sql & "vwConsSUMIVKardexIVA inner join "
'                sql = sql & "GNComprobante GNC   "
'                sql = sql & " INNER JOIN PCPROVCLI PC ON GNC.IDCLIENTEREF=PC.IDPROVCLI"
'                sql = sql & " INNER JOIN GNTRANS gnt ON gnc.CODTRANS=GNT.CODTRANS"
'                sql = sql & " ON vwConsSUMIVKardexIVA.TransID = GNC.TransID "
'                sql = sql & " ON tmp0.TransID = GNC.TransID"
'                sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(VT_ExpBie) & ")  "
'                sql = sql & " and bandCierre=9 "
'                sql = sql & " AND GNT.AnexoCodTipoTrans=2 " & Cond
'                Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
'                'VentaNeta = Round(Abs(rs.Fields("BaseIVA")), 2)
'                'ventasSinIva = Round(Abs(rs.Fields("Base0")), 2) 'carga 404
'                VentaNetaExp = Round(Abs(rs.Fields("BaseIVA")), 2)
'                ventasSinIvaExp = Round(Abs(rs.Fields("Base0")), 2) 'carga 404
'                'grd.TextMatrix(18, 8) = Round(Abs(rs.Fields("Base0")), 2) + trans 'carga 404
'            'descuentos  12 % ventas
'            sql = "SELECT "
'            sql = sql & " SUM(preciototal-PrecioRealTotal) As DesNC  "
'            sql = sql & " FROM ivkardex  inner join "
'            sql = sql & "GNComprobante GNC   "
'            sql = sql & " INNER JOIN PCPROVCLI PC ON GNC.IDCLIENTEREF=PC.IDPROVCLI"
'            sql = sql & " INNER JOIN GNTRANS gnt ON gnc.CODTRANS=GNT.CODTRANS"
'            sql = sql & " ON ivkardex.TransID = GNC.TransID "
'            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(VT_ExpBie) & ")  "
'            sql = sql & " AND ivkardex.iva <> 0 and  GNT.AnexoCodTipoTrans=2 " & Cond
'            'Dim DESIVA As Currency
'            Set rsd = gobjMain.EmpresaActual.OpenRecordset(sql)
'            If Not IsNull(rsd!desnc) Then DESIVA = Abs(rsd!desnc)
'            grd.TextMatrix(20, 8) = Round(Abs(rs.Fields("Baseiva")), 2) '- DESIVA - DescNC  'carga 401
'            'notas de credito EXPORTACIONES
'            sql = "SELECT "
'            sql = sql & " ABS(Sum(CASE IVA   WHEN 0 THEN PrecioRealTotal   ELSE 0 END)) AS NCTotalBase0  "
'            'sql = sql & " ABS(Sum(CASE IVA   WHEN 0 THEN 0   ELSE PrecioRealTotal  END)) AS NCTotalBaseIVA" 'para exporaciones solo base 0
'            sql = sql & " FROM ivkardex  inner join "
'            sql = sql & " GNComprobante GNC "
'            sql = sql & " INNER JOIN PCPROVCLI PC ON GNC.IDCLIENTEREF=PC.IDPROVCLI"
'            sql = sql & " INNER JOIN GNTRANS gnt ON gnc.CODTRANS=GNT.CODTRANS"
'            sql = sql & " ON ivkardex.TransID = GNC.TransID "
'            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(NC_Ventas) & ")  "
'            sql = sql & " AND GNT.AnexoCodTipoTrans=4 " & Cond
'            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
'            NCreditoExp = IIf(Not IsNull(rs.Fields("NCTotalBase0")), Round(Abs(rs.Fields("NCTotalBase0")), 2), 0)
'            'grd.TextMatrix(20, 10) = VentaNetaExp - NCreditoExp
'            '-------------------
'            'notas de credito EXPORTACIONES BIENES
'            sql = "SELECT "
'            sql = sql & " ABS(Sum(CASE IVA   WHEN 0 THEN PrecioRealTotal   ELSE 0 END)) AS NCTotalBase0  "
'            'sql = sql & " ABS(Sum(CASE IVA   WHEN 0 THEN 0   ELSE PrecioRealTotal  END)) AS NCTotalBaseIVA" 'para exporaciones solo base 0
'            sql = sql & " FROM ivkardex  inner join "
'            sql = sql & " GNComprobante GNC "
'            sql = sql & " INNER JOIN PCPROVCLI PC ON GNC.IDCLIENTEREF=PC.IDPROVCLI"
'            sql = sql & " INNER JOIN GNTRANS gnt ON gnc.CODTRANS=GNT.CODTRANS"
'            sql = sql & " ON ivkardex.TransID = GNC.TransID "
'            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(NC_ExpB) & ")  "
'            sql = sql & " and bandCierre=9 "
'            sql = sql & " AND GNT.AnexoCodTipoTrans=2 " & Cond
'            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
'            NCreditoExpB = IIf(Not IsNull(rs.Fields("NCTotalBase0")), Round(Abs(rs.Fields("NCTotalBase0")), 2), 0)
'            grd.TextMatrix(20, 10) = VentaNetaExp - NCreditoExpB
            Cargadatos407_417 Recargo, VT_ExpBie, NC_ExpB, cond
            '--------------HASTA AQUI EXPORTACION
            sql = "SELECT "
            sql = sql & " isnull(SUM(PrecioTotalBase0 *SignoVenta + (PrecioTotalBase0*SignoVenta * (cast(isnull(TotalDescuento,0) as float) *SignoVenta/ cast(PrecioTotal*SignoVenta as float)*SignoVenta))),0) as Base0, "
            sql = sql & " isnull(SUM(PrecioTotalBaseIVA * SignoVenta + (PrecioTotalBaseIVA *SignoVenta* (cast(isnull(TotalDescuento,0) as float) *SignoVenta/ cast(PrecioTotal as float)*SignoVenta))),0) As BaseIVA "
            sql = sql & " FROM tmp0 Right join "
            sql = sql & "vwConsSUMIVKardexIVA inner join "
            sql = sql & "GNComprobante GNC   "
            sql = sql & " INNER JOIN PCPROVCLI PC ON GNC.IDCLIENTEREF=PC.IDPROVCLI"
            sql = sql & " INNER JOIN GNTRANS gnt ON gnc.CODTRANS=GNT.CODTRANS"
            sql = sql & " ON vwConsSUMIVKardexIVA.TransID = GNC.TransID "
            sql = sql & " ON tmp0.TransID = GNC.TransID"
            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(NC_Ventas) & ")  "
            sql = sql & " AND pc.BANDEMPRESAPUBLICA=1"
            sql = sql & " AND GNT.AnexoCodTipoTrans=2 " & cond
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            grd.TextMatrix(15, 10) = Round(Abs(rs.Fields("Base0")), 2)
            grd.TextMatrix(24, 10) = Round(Abs(rs.Fields("BaseIVA")), 2)
'********** EXPORTACION SERVICIOS
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
            sql = sql & " isnull(SUM(PrecioTotalBase0 *SignoVenta ),0) as Base0, "
            sql = sql & " isnull(SUM(PrecioTotalBaseIVA * SignoVenta ),0) As BaseIVA "
            sql = sql & " FROM tmp0 Right join "
            sql = sql & " vwConsSUMIVKardexIVA inner join "
            sql = sql & " GNComprobante GNC   "
            sql = sql & " INNER JOIN GNTRANS gnt ON gnc.CODTRANS=GNT.CODTRANS"
            sql = sql & " ON vwConsSUMIVKardexIVA.TransID = GNC.TransID "
            sql = sql & " ON tmp0.TransID = GNC.TransID"
            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(VT_ExpSer) & ")  "
            sql = sql & " AND GNT.AnexoCodTipoTrans=4 " & cond
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            grd.TextMatrix(22, 8) = Round(Abs(rs.Fields("BaseIVA")), 2) + Round(Abs(rs.Fields("Base0")), 2)
'********** reposicion Gastos
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
            sql = sql & " isnull(SUM(PrecioTotalBase0 *SignoVenta + (PrecioTotalBase0*SignoVenta * (cast(isnull(TotalDescuento,0) as float) *SignoVenta/ cast(PrecioTotal*SignoVenta as float)*SignoVenta))),0) as Base0, "
            sql = sql & " isnull(SUM(PrecioTotalBaseIVA * SignoVenta + (PrecioTotalBaseIVA *SignoVenta* (cast(isnull(TotalDescuento,0) as float) *SignoVenta/ cast(PrecioTotal as float)*SignoVenta))),0) As BaseIVA "
            sql = sql & " FROM tmp0 Right join "
            sql = sql & " vwConsSUMIVKardexIVA inner join "
            sql = sql & " GNComprobante GNC   "
            sql = sql & " INNER JOIN GNTRANS gnt ON gnc.CODTRANS=GNT.CODTRANS"
            sql = sql & " ON vwConsSUMIVKardexIVA.TransID = GNC.TransID "
            sql = sql & " ON tmp0.TransID = GNC.TransID"
            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(VT_RepGas) & ")  "
            sql = sql & " AND GNT.AnexoCodTipoTrans=4 " & cond
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            grd.TextMatrix(19, 10) = Round(Abs(rs.Fields("Base0")), 2)
       '     grd.TextMatrix(36, 7) = Round(Abs(rs.Fields("BaseIVA")), 2)
    '********** VENTAS ACtivos
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
            sql = sql & " isnull(SUM(PrecioTotalBase0 *SignoVenta - (PrecioTotalBase0 * (cast(isnull(TotalDescuento,0) as float) / cast(PrecioTotal as float)))),0) as Base0, "
            sql = sql & " isnull(SUM(PrecioTotalBaseIVA *SignoVenta - (PrecioTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(PrecioTotal as float)))),0) As BaseIVA "
            sql = sql & " FROM tmp0 Right join "
            sql = sql & " vwConsSUMIVKardexIVA inner join "
            sql = sql & " GNComprobante GNC   "
            sql = sql & " INNER JOIN PCPROVCLI PC ON GNC.IDCLIENTEREF=PC.IDPROVCLI"
            sql = sql & " INNER JOIN GNTRANS gnt ON gnc.CODTRANS=GNT.CODTRANS"
            sql = sql & " ON vwConsSUMIVKardexIVA.TransID = GNC.TransID "
            sql = sql & " ON tmp0.TransID = GNC.TransID"
            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(VT_Ser) & ")"
            sql = sql & " AND pc.BANDEMPRESAPUBLICA=0"
            sql = sql & " AND GNT.AnexoCodTipoTrans=2 " & cond
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            grd.TextMatrix(15, 8) = Round(Abs(rs.Fields("Base0")), 2)
            grd.TextMatrix(15, 12) = Round(Abs(rs.Fields("BaseIVA")), 2)
            sql = "SELECT "
            sql = sql & " isnull(SUM(PrecioTotalBase0 *SignoVenta - (PrecioTotalBase0 * (cast(isnull(TotalDescuento,0) as float) / cast(PrecioTotal as float)))),0) as Base0, "
            sql = sql & " isnull(SUM(PrecioTotalBaseIVA *SignoVenta - (PrecioTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(PrecioTotal as float)))),0) As BaseIVA "
            sql = sql & " FROM tmp0 Right join "
            sql = sql & " vwConsSUMIVKardexIVA inner join "
            sql = sql & " GNComprobante GNC   "
            sql = sql & " INNER JOIN PCPROVCLI PC ON GNC.IDCLIENTEREF=PC.IDPROVCLI"
            sql = sql & " INNER JOIN GNTRANS gnt ON gnc.CODTRANS=GNT.CODTRANS"
            sql = sql & " ON vwConsSUMIVKardexIVA.TransID = GNC.TransID "
            sql = sql & " ON tmp0.TransID = GNC.TransID"
            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(VT_Ser) & ")"
            sql = sql & " AND pc.BANDEMPRESAPUBLICA=1"
            sql = sql & " AND GNT.AnexoCodTipoTrans=2 " & cond
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            'grd.TextMatrix(14, 10) = Round(Abs(rs.Fields("Base0")), 2) + Round(Abs(rs.Fields("BaseIVA")), 2)
            ' cantidad de comprobantes
            sql = "SELECT "
            sql = sql & "  count(gnc.codtrans) as NumComp  "
            sql = sql & " FROM  gncomprobante gnc inner join gntrans gnt "
            sql = sql & " on gnc.codtrans=gnt.codtrans"
            sql = sql & " WHERE ( GNC.CodTrans IN (" & PreparaCadena(VT_Bie) & " ) "
            sql = sql & " or GNC.CodTrans IN (" & PreparaCadena(VT_Ser) & "))  " & cond
            'sql = sql & "group by AnexoCodTipoComp"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If rs.RecordCount > 0 Then
            '    grd.TextMatrix(45, 5) = rs.Fields("NumComp")
            Else
             '   grd.TextMatrix(45, 5) = "0"
            End If
            '*************************
            '***COMPRAS
            '*****************************
    
            Cargadatos500_510 Recargo, objcond.CodTrans, NC_Compras, cond
            Cargadatos501_511 Recargo, CP_Act, "", cond
            Cargadatos502_512 Recargo, objcond.CodTrans, NC_Compras, cond
            Cargadatos503_513 Recargo, objcond.CodTrans, NC_Compras, cond
            Dim pos As Integer
            pos = InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "HORMI")
            If pos > 0 Then
                Cargadatos504_514Hormi Recargo, CP_Bie, "", cond
                Cargadatos505_515Hormi Recargo, CP_Rise, "", cond
            Else
                Cargadatos504_514 Recargo, CP_Bie, "", cond
                Cargadatos505_515 Recargo, CP_Rise, "", cond
            End If
            
            Cargadatos506_516 Recargo, CP_Bie, CP_Rise, cond
            Cargadatos507_517 Recargo, objcond.CodTrans, CP_Act, NC_Compras, cond
            Cargadatos508_518 Recargo, objcond.CodTrans, NC_Compras, cond
            Cargadatos531_541 Recargo, objcond.CodTrans, NC_Compras, cond
                       
                     
        
           ' RETENCIONES REALIZADAS
            sql = "SELECT CodF104,"
            sql = sql & " ISNULL(sum(Haber),0) as TotalHaber "
            sql = sql & " FROM vwConsRetencion "
            If Len(ret_real) > 0 Then
                sql = sql & " WHERE CodTrans IN (" & PreparaCadena(ret_real) & ") AND "
            Else
                sql = sql & " WHERE "
            End If
            sql = sql & "  FechaTrans BETWEEN " & FechaYMD(f1, gobjMain.EmpresaActual.TipoDB)
            sql = sql & "  AND " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB)
            sql = sql & "  and  HABER > 0 and bandIVA=1 "
            sql = sql & "  group by Codf104" ', Porcentaje"
           Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            Do While Not rs.EOF
                 Select Case rs.Fields("CodF104")
                         Case 721
                             grd.TextMatrix(72, 12) = Round(rs.Fields("TotalHaber"), 2)
                         Case 723
                             grd.TextMatrix(73, 12) = Round(rs.Fields("TotalHaber"), 2)
                         Case 725
                             grd.TextMatrix(74, 12) = Round(rs.Fields("TotalHaber"), 2)
                End Select
                rs.MoveNext
            Loop
            
' RETENCIONES RECIBIDAS
            sql = "SELECT "
            sql = sql & " ISNULL(sum(DEBE),0) as TotalDEBE "
            sql = sql & " FROM vwConsRetencion "
            If Len(ret_real) > 0 Then
                sql = sql & " WHERE CodTrans IN (" & PreparaCadena(ret_recib) & ") AND "
            Else
                sql = sql & " WHERE "
            End If
            sql = sql & "  FechaTrans BETWEEN " & FechaYMD(f1, gobjMain.EmpresaActual.TipoDB)
            sql = sql & "  AND " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB)
            sql = sql & "  and  DEBE > 0 and bandIVA=1 "
'            sql = sql & "  group by Codf104" ', Porcentaje"
           Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
           grd.TextMatrix(61, 12) = Round(rs.Fields("TotalDEBE"), 2)
'            Do While Not rs.EOF
'                 Select Case rs.Fields("CodF104")
'                         Case 721
'                             grd.TextMatrix(72, 12) = Round(rs.Fields("TotalHaber"), 2)
'                         Case 723
'                             grd.TextMatrix(73, 12) = Round(rs.Fields("TotalHaber"), 2)
'                         Case 725
'                             grd.TextMatrix(74, 12) = Round(rs.Fields("TotalHaber"), 2)
'                End Select
'                rs.MoveNext
'            Loop
            
            CalcularPorcentajes1042010
            Select Case Me.tag
            Case "F103"
                nombre = "103_" & Format(CStr(Month(objcond.fecha1)), "00") & "_" & Year(objcond.fecha1) & ".XML"
                txtDestino.Text = mRutaDestino103 & "103_" & Format(CStr(Month(objcond.fecha1)), "00") & "_" & Year(objcond.fecha1) & ".XML"
            Case "F104", "F1042010"
                nombre = "104_" & Format(CStr(Month(objcond.fecha1)), "00") & "_" & Year(objcond.fecha1) & ".XML"
                txtDestino.Text = mRutaDestino104 & "\104_" & Format(CStr(Month(objcond.fecha1)), "00") & "_" & Year(objcond.fecha1) & ".XML"
            End Select
            grd.Refresh
    End With
End Sub

Private Sub LlenaFormatoFormulario103_2010()
    With grd
        .MergeCells = flexMergeSpill
        
        .AddItem "1" & vbTab & "Formulario" & vbTab & vbTab & " DECLARACION DE RETENCION EN LA FUENTE"
        .AddItem "2" & vbTab & "    103" & vbTab & vbTab & vbTab & "DEL IMPUESTO A LA RENTA"
        .AddItem "3" & vbTab & "Resolucion No "
        .AddItem "4" & vbTab & "NAG-DGER2008-1520 "
        .AddItem "5" & vbTab & "________________________________________________________________________________________________________________________________________________________________________________________________"
        .AddItem "6" & vbTab & "100 IDENTIFICACION DE LA DECLARACION"
        .AddItem "7" & vbTab & "MES 101" & vbTab & vbTab & "AÑO 102" & vbTab & vbTab & vbTab & vbTab & "(O) ORIGINAL - (S) SUSTITUTIVA   031"
        .AddItem "8" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "No FORMULARIO QUE SUSTITUYE 104"
        .AddItem "9" & vbTab & "200 IDENTIFICACION DEL SUJETO PASIVO (AGENTE DE RETENCION)"
        .AddItem "10" & vbTab & "RUC 201" & vbTab & vbTab & "202"
        .AddItem "11"
        .AddItem "12" & vbTab & vbTab & vbTab & "DETALLE DE PAGOS Y RETENCIÒN POR IMPUESTO A LA RENTA"
        .AddItem "13" & vbTab & vbTab & vbTab & vbTab & "POR PAGOS  EFECTUADOS EN EL PAIS "
        .AddItem "14" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "Base Imponible" & vbTab & vbTab & "Valor Retenido"
        .AddItem "15" & vbTab & "En relacion de dependencia que supera o no a ala base gravada" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "302" & vbTab & vbTab & "352"
        .AddItem "16" & vbTab & "Servicios" & vbTab & "Honorarios profesionales y dietas" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "303" & vbTab & vbTab & "353"
        .AddItem "17" & vbTab & vbTab & "Predomina el Intelecto" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "304" & vbTab & vbTab & "354"
        .AddItem "18" & vbTab & vbTab & "Predomina mano de obra" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "307" & vbTab & vbTab & "357"
        .AddItem "19" & vbTab & vbTab & "Entre Sociedades" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "308" & vbTab & vbTab & "358"
        .AddItem "20" & vbTab & vbTab & "Publicidad y comunicacion" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "309" & vbTab & vbTab & "359"
        .AddItem "21" & vbTab & vbTab & "Transporte Privado de pasajeros o servicio publico o privado de  carga" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "310" & vbTab & vbTab & "360"
        .AddItem "22" & vbTab & "Transferencia de bienes muebles de naturaleza corporal" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "312" & vbTab & vbTab & "362"
        .AddItem "23" & vbTab & "Arrendamiento" & vbTab & "Mercantil" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "319" & vbTab & vbTab & "369"
        .AddItem "24" & vbTab & vbTab & "Bienes Inmuebles" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "320" & vbTab & vbTab & "370"
        .AddItem "25" & vbTab & "Seguros y reaseguros (primas y cesiones)" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "322" & vbTab & vbTab & "372"
        .AddItem "26" & vbTab & "Rendimientos Financieros" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "323" & vbTab & vbTab & "373"
        .AddItem "27" & vbTab & "Dividendos" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "324" & vbTab & vbTab & "374"
        .AddItem "28" & vbTab & "Loterias, rifas, apuestas y similares" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "325" & vbTab & vbTab & "375"
        .AddItem "29" & vbTab & "Ventas de" & vbTab & "A comercializadoras" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "327" & vbTab & vbTab & "377"
        .AddItem "30" & vbTab & "Combustibles" & vbTab & "A Distribuidores" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "328" & vbTab & vbTab & "378"
        .AddItem "31" & vbTab & "Compra local de banano" & vbTab & "a productor" & vbTab & "No Cajas transferidas" & vbTab & vbTab & vbTab & 510 & vbTab & vbTab & vbTab & "329" & vbTab & vbTab & "379"
        .AddItem "32" & vbTab & "Impuesto a la actividad bananera" & vbTab & "Productor-Exportador" & vbTab & "No Cajas transferidas" & vbTab & vbTab & vbTab & 520 & vbTab & vbTab & vbTab & "330" & vbTab & vbTab & "380"
        .AddItem "33" & vbTab & "Pagos de bienes o servicios no sujetos a retencion" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "332"
        .AddItem "34" & vbTab & "Otras retenciones " & vbTab & "Aplicables al 1%" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "340" & vbTab & vbTab & "390"
        .AddItem "35" & vbTab & vbTab & "Aplicables al 2%" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "341" & vbTab & vbTab & "391"
        .AddItem "36" & vbTab & vbTab & "Aplicables al 8%" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "342" & vbTab & vbTab & "392"
        .AddItem "37" & vbTab & vbTab & "Aplicables al 25%" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "343" & vbTab & vbTab & "393"
        .AddItem "38" & vbTab & vbTab & "Aplicables a otros porcentajes" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "344" & vbTab & vbTab & "394"
        .AddItem "39" & vbTab & "SUBTOTALES  OPERACIONES EFECTUADAS EN EL PAIS" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "349" & vbTab & vbTab & "399"
        .AddItem "40" & vbTab & vbTab & vbTab & vbTab & vbTab & "POR PAGOS AL EXTERIOR"
        .AddItem "41" & vbTab & "Con convenio de doble tributacion" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "401" & vbTab & vbTab & "451"
        .AddItem "42" & vbTab & "Sin convenio " & vbTab & "Interes por financiamiento  de proveedores externos" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "403" & vbTab & vbTab & "453"
        .AddItem "43" & vbTab & "De doble " & vbTab & "Interes de creditos externos" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "405" & vbTab & vbTab & "455"
        .AddItem "44" & vbTab & "De doble " & vbTab & "Dividendos" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "407" & vbTab & vbTab & "457"
        
        
        .AddItem "45" & vbTab & "Tributacion " & vbTab & "Otros conceptos" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "421" & vbTab & vbTab & "471"
        .AddItem "46" & vbTab & "Pagos al exterior no sujetos a retencion " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "427"
        .AddItem "47" & vbTab & "SUBTOTAL OPERACIONES EFECTUADAS EN EL EXTERIOR " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "429" & vbTab & vbTab & "498"
        .AddItem "48"
        .AddItem "49" & vbTab & "TOTAL DE RETENCION DE IMPUESTO A LA RENTA (399+498) " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "499"
        .AddItem "50"
        .AddItem "51" & vbTab & "Pago previo (informativo) " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "890"
        .AddItem "52" & vbTab & "DETALLE DE IMPUTACION AL PAGO (para declaraciones sustitutivas)"
        .AddItem "53" & vbTab & vbTab & "Interes 897" & vbTab & vbTab & vbTab & "Impuesto  898 " & vbTab & vbTab & vbTab & vbTab & vbTab & "Multa 899 "
        
        .AddItem "54" & vbTab & vbTab & "PAGO DIRECTO EN CUENTA UNICA DE TESORO NACIONAL" & vbTab & vbTab & vbTab & "(Uso exclusivo para instituciones y empresas del sector publico autorizadas) " & vbTab & vbTab & vbTab & vbTab & vbTab
        
        
        .AddItem "55"
        .AddItem "56" & vbTab & "VALORES A PAGAR Y FORMAS DE PAGO (luego de imputacion al pago en declaraciones sustitutivas)"
        .AddItem "57" & vbTab & "Total Impuesto a Pagar (499-897) " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "902"
        .AddItem "58" & vbTab & "Interes por mora " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "903"
        .AddItem "59" & vbTab & "Multas  " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "904"
        .AddItem "60" & vbTab & "Total Pagado (902+903+904)  " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "999"
        .AddItem "61"
        .AddItem "62" & vbTab & "Mediante Cheque,debito bancario,efectivo u otras formas de pago  " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "905"
        .AddItem "63" & vbTab & "Mediante Notas de credito  " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "907"
        .AddItem "64" & vbTab & "DETALLE NOTAS DE CREDITO"
        .AddItem "65" & vbTab & vbTab & "N/C No.908 " & vbTab & vbTab & "N/C No.910 " & vbTab & vbTab & "N/C No.912 " & vbTab & vbTab & vbTab & vbTab & "DETALLE DE NOTAS DE CREDITO DESMATERIALIZADAS "
        .AddItem "66" & vbTab & vbTab & "Valor USD 909 " & vbTab & vbTab & "Valor  USD 911 " & vbTab & vbTab & "Valor USD 913 " & vbTab & vbTab & vbTab & vbTab & "Valor USD 915 "
        .AddItem "67"
        .AddItem "68" & vbTab & "Declaro que todos los datos proporcionados en este documento son exactos y verdaderos,por lo que asumo la responsabilidad legal que de ella se deriven  (Art,101 de la  L.O.R.T.I.)"
        .AddItem "69"
        .AddItem "70" & vbTab & "No. ID SUJETO PASIVO/REP. LEGAL 198" & vbTab & vbTab & vbTab & vbTab & vbTab & "RUC CONTADOR 199 "
        .AddItem "71"
         .AddItem "72" & vbTab & vbTab & vbTab & vbTab & vbTab & "FORMA DE PAGO 921 "
        .AddItem "73" & vbTab & vbTab & vbTab & vbTab & vbTab & "BANCO 922 "
        .Redraw = flexRDBuffered
        .Refresh
    End With
    CambiaFondoCeldasEditables103_2010
End Sub

Private Sub ConfigCols1032010()
    Dim s As String, i As Long, j As Integer, s1 As String
    Dim fmt As String
    With grd
        s = "^#|<c1|<c2|<c3|<c4|<c5|<c6|<c7|>c8|<c9|>c10|<c11|>c12"
        .FormatString = s
        AsignarTituloAColKey grd
        .ColWidth(0) = 350
        .ColWidth(1) = 850
        .ColWidth(2) = 1600
        .ColWidth(3) = 1500
        .ColWidth(4) = 1450
        .ColWidth(5) = 1550
        .ColWidth(6) = 1350
        .ColWidth(7) = 0
        .ColWidth(8) = 1600
        .ColWidth(9) = 400
        .ColWidth(10) = 1400
        .ColWidth(11) = 450
        .ColWidth(12) = 1400
        .ColDataType(8) = flexDTCurrency
        .ColDataType(10) = flexDTCurrency
        .ColDataType(12) = flexDTCurrency
        .ColDataType(7) = flexDTString
        .ColDataType(8) = flexDTString
        .ColDataType(9) = flexDTString
        .ColDataType(10) = flexDTString
        
        'Columnas modificables (Longitud maxima)
        '.ColFormat(2) = gobjMain.EmpresaActual.GNOpcion.FormatoCantidad
       ' .ColFormat(6) = gobjMain.EmpresaActual.GNOpcion.FormatoCantidad
        .ColFormat(8) = gobjMain.EmpresaActual.GNOpcion.FormatoCantidad
        .ColFormat(10) = gobjMain.EmpresaActual.GNOpcion.FormatoCantidad
        .ColFormat(12) = gobjMain.EmpresaActual.GNOpcion.FormatoCantidad
        
        
        .MergeCol(3) = True
        .Refresh
    End With
End Sub

Private Sub CambiaFondoCeldasEditables103_2010()
    Dim fmt As String
    With grd
        .Redraw = flexRDBuffered
        'fondo blanco
        .BackColor = &HE0E0E0   '&HE0E0E0
        .Cell(flexcpBackColor, 7, 2, 7, 2) = &H800000
        .Cell(flexcpBackColor, 7, 4, 7, 4) = &H800000
        .Cell(flexcpBackColor, 7, 11, 7, 11) = &H80000014 'tambien es blanco
        .Cell(flexcpBackColor, 7, 12, 7, 12) = vbWhite
        
        .Cell(flexcpBackColor, 8, 11, 8, 11) = vbWhite
        .Cell(flexcpBackColor, 8, 12, 8, 12) = &H80000014 'tambien es blanco

        .Cell(flexcpBackColor, 10, 2, 10, 2) = &H800000
        .Cell(flexcpBackColor, 10, 4, 10, 6) = &H800000
        
        .Cell(flexcpBackColor, 31, 8, 32, 8) = vbWhite
        
        .Cell(flexcpBackColor, 15, 10, 38, 10) = vbWhite
        .Cell(flexcpBackColor, 15, 12, 32, 12) = vbWhite
        
        .Cell(flexcpBackColor, 8, 10, 8, 10) = vbWhite
        .Cell(flexcpBackColor, 23, 10, 26, 10) = vbWhite
        .Cell(flexcpBackColor, 35, 10, 39, 10) = vbWhite
        .Cell(flexcpBackColor, 34, 12, 39, 12) = vbWhite
        
'        .Cell(flexcpBackColor, 37, 10, 41, 10) = vbWhite
'        .Cell(flexcpBackColor, 37, 12, 40, 12) = vbWhite
        
        .Cell(flexcpBackColor, 41, 10, 47, 10) = vbWhite
        .Cell(flexcpBackColor, 41, 12, 45, 12) = vbWhite
        
        '.Cell(flexcpBackColor, 44, 12, 44, 12) = vbWhite
        .Cell(flexcpBackColor, 47, 12, 47, 12) = vbWhite
'        .Cell(flexcpBackColor, 48, 3, 48, 3) = vbWhite
        '.Cell(flexcpBackColor, 48, 6, 48, 6) = vbWhite
        .Cell(flexcpBackColor, 49, 12, 49, 12) = vbWhite
        .Cell(flexcpBackColor, 51, 12, 51, 12) = vbWhite
        .Cell(flexcpBackColor, 53, 3, 53, 3) = vbWhite
        .Cell(flexcpBackColor, 53, 6, 53, 6) = vbWhite
        .Cell(flexcpBackColor, 53, 12, 53, 12) = vbWhite
        

        .Cell(flexcpBackColor, 54, 12, 54, 12) = vbWhite
        .Cell(flexcpBackColor, 57, 12, 60, 12) = vbWhite
        .Cell(flexcpBackColor, 62, 12, 63, 12) = vbWhite
        
        .Cell(flexcpBackColor, 65, 3, 66, 3) = vbWhite
        .Cell(flexcpBackColor, 65, 5, 66, 5) = vbWhite
        .Cell(flexcpBackColor, 65, 8, 66, 8) = vbWhite
        .Cell(flexcpBackColor, 65, 12, 66, 12) = vbWhite
        
        .Cell(flexcpBackColor, 70, 4, 70, 4) = vbWhite
        .Cell(flexcpBackColor, 70, 9, 70, 10) = vbWhite
        .Cell(flexcpBackColor, 72, 8, 73, 8) = vbWhite
        
        
        .Cell(flexcpForeColor, 1, 1, 5, 8) = vbBlue
        .Cell(flexcpForeColor, 14, 1, 21, 1) = vbBlue
        .Cell(flexcpForeColor, 14, 9, 21, 9) = vbBlue
        .Cell(flexcpForeColor, 14, 7, 21, 7) = vbBlue
        .Cell(flexcpForeColor, 14, 11, 15, 11) = vbBlue
        .Cell(flexcpForeColor, 36, 1, 43, 1) = vbBlue

        .Cell(flexcpForeColor, 7, 1, 60, 12) = vbBlue
        .Cell(flexcpForeColor, 9, 1, 9, 1) = vbBlack
        .Cell(flexcpForeColor, 12, 1, 13, 12) = vbBlack
        .Cell(flexcpForeColor, 36, 5, 36, 5) = vbBlack
        



        .Cell(flexcpForeColor, 63, 1, 63, 1) = &HC0&
        .Cell(flexcpForeColor, 65, 1, 65, 1) = &HC0&
        
        
        .Cell(flexcpForeColor, 7, 2, 7, 2) = vbYellow
'
        .Cell(flexcpForeColor, 7, 4, 7, 4) = vbYellow
        .Cell(flexcpForeColor, 10, 2, 10, 2) = vbYellow
'
        .Cell(flexcpForeColor, 10, 4, 10, 6) = vbYellow
        
'        .Cell(flexcpForeColor, 7, 12, 8, 12) = vbYellow
'        .Cell(flexcpForeColor, 35, 10, 35, 10) = vbYellow
'        .Cell(flexcpForeColor, 35, 12, 35, 12) = vbYellow
'        .Cell(flexcpForeColor, 42, 10, 42, 10) = vbYellow
'        .Cell(flexcpForeColor, 42, 12, 42, 12) = vbYellow
'        .Cell(flexcpForeColor, 44, 12, 44, 12) = vbYellow
'        .Cell(flexcpForeColor, 51, 12, 51, 12) = vbYellow
'        .Cell(flexcpForeColor, 54, 12, 54, 12) = vbYellow
        
        .Cell(flexcpFontSize, 1, 1, 1, 1) = 11
        .Cell(flexcpFontSize, 1, 3, 1, 3) = 14
        .Cell(flexcpFontSize, 2, 4, 2, 4) = 14
        .Cell(flexcpFontSize, 2, 1, 2, 1) = 13
        .Cell(flexcpFontSize, 62, 1, 62, 1) = 7
'        .Cell(flexcpForeColor, 35, 1, 35, 12) = &HC0& 'rojo
'        .Cell(flexcpForeColor, 42, 1, 42, 12) = &HC0& 'rojo
'        .Cell(flexcpForeColor, 44, 1, 44, 12) = &HC0& 'rojo
'        .Cell(flexcpForeColor, 51, 1, 51, 12) = &HC0& 'rojo
'        .Cell(flexcpForeColor, 54, 1, 54, 12) = &HC0& 'rojo
'        .Cell(flexcpForeColor, 35, 10, 35, 10) = vbYellow
'        .Cell(flexcpForeColor, 35, 12, 35, 12) = vbYellow
'        .Cell(flexcpForeColor, 42, 10, 42, 10) = vbYellow
'        .Cell(flexcpForeColor, 42, 12, 42, 12) = vbYellow
'        .Cell(flexcpForeColor, 44, 12, 44, 12) = vbYellow
'        .Cell(flexcpForeColor, 51, 12, 51, 12) = vbYellow
'        .Cell(flexcpForeColor, 54, 12, 54, 12) = vbYellow
        .Cell(flexcpFontBold, 2, 1, 2, 1) = True
        .Cell(flexcpForeColor, 7, 12, 7, 12) = vbBlack
        .Cell(flexcpForeColor, 8, 11, 8, 11) = vbBlack
        .RowHeight(1) = 400
        .RowHeight(2) = 400
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .MergeCol(4) = True
        .MergeCol(5) = True
    End With
End Sub




Private Sub Buscar1032010()
    Dim sql As String, cond As String, Reten As String, CPNotRet As String, Recargo As String
    Dim OrdenadoX As String, f1 As String, baseret As Currency
    Dim rs As Recordset
    Dim Moneda As String
    Set objcond = gobjMain.objCondicion
    If Not frmB_FormSRI103.Inicio103(objcond, Reten) Then
        grd.SetFocus
        Exit Sub
    End If
    With objcond
        grd.TextMatrix(7, 2) = Format(.fecha1, "mmmm")
        grd.TextMatrix(7, 4) = Format(.fecha1, "yyyy")
        baseret = 0
        'Reporte de un mes a la vez
        f1 = DateSerial(Year(.fecha1), Month(.fecha1), 1)
        cond = " AND GNC.FechaTrans BETWEEN " & FechaYMD(f1, gobjMain.EmpresaActual.TipoDB) & _
               " AND " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB)
               
        sql = " SELECT"
        sql = sql & " CodF104,  sum(tskardexret.BAse) AS TBase, tsretencion.porcentaje as Porcentaje"
        sql = sql & " from GNComprobante gnc "
        sql = sql & " inner join gntrans gnt on gnc.codtrans=gnt.codtrans"
        sql = sql & " INNER JOIN tskardexret "
        sql = sql & " INNER JOIN tsretencion  ON tskardexret.idretencion = tsretencion.idretencion  "
        sql = sql & " ON GNC.TransID = tskardexret.transid"
        sql = sql & " Where GNC.Estado <> 3 " & cond
        sql = sql & " and tsretencion.BandValida=1"
        sql = sql & " and len(CodF104)>0"
        sql = sql & " and bandiva=0"
        sql = sql & " and AnexoCodTipoComp=7"
        If Len(Reten) > 0 Then
            sql = sql & " and tsretencion.CodRetencion in(" & PreparaCadena(Reten) & ") "
        End If
        sql = sql & " group BY CodF104,tsretencion.porcentaje"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                Do While Not rs.EOF
                    Select Case rs.Fields("CodF104")
                        Case "309", "359"
                           grd.TextMatrix(20, 10) = Round(rs.Fields("TBase"), 2)
                            grd.TextMatrix(20, 12) = Round(rs.Fields("TBase") * rs.Fields("porcentaje"), 2)
                        Case "310", "360"
                            grd.TextMatrix(21, 10) = Round(rs.Fields("TBase"), 2)
                            grd.TextMatrix(21, 12) = Round(rs.Fields("TBase") * rs.Fields("porcentaje"), 2)
                        Case "312"
                            grd.TextMatrix(22, 10) = Round(rs.Fields("TBase"), 2)
                            grd.TextMatrix(22, 12) = Round(rs.Fields("TBase") * rs.Fields("porcentaje"), 2)
                        Case "320", "370"
                            grd.TextMatrix(24, 10) = Round(rs.Fields("TBase"), 2)
                            grd.TextMatrix(24, 12) = Round(rs.Fields("TBase") * rs.Fields("porcentaje"), 2)
                        Case "322"
                            grd.TextMatrix(25, 10) = Round(rs.Fields("TBase"), 2)
                            grd.TextMatrix(25, 12) = Round(rs.Fields("TBase") * rs.Fields("porcentaje"), 2)
                        Case "341", "392"
                            grd.TextMatrix(35, 10) = Round(rs.Fields("TBase"), 2)
                            grd.TextMatrix(35, 12) = Round(rs.Fields("TBase") * rs.Fields("porcentaje"), 2)
                        Case "303"
                            grd.TextMatrix(16, 10) = Round(rs.Fields("TBase"), 2)
                            grd.TextMatrix(16, 12) = Round(rs.Fields("TBase") * rs.Fields("porcentaje"), 2)
                        Case "307"
                            grd.TextMatrix(18, 10) = Round(rs.Fields("TBase"), 2)
                            grd.TextMatrix(18, 12) = Round(rs.Fields("TBase") * rs.Fields("porcentaje"), 2)
                         Case "340"
                            grd.TextMatrix(34, 10) = Round(rs.Fields("TBase"), 2)
                            grd.TextMatrix(34, 12) = Round(rs.Fields("TBase") * rs.Fields("porcentaje"), 2)
                         Case "341"
                            grd.TextMatrix(35, 10) = Round(rs.Fields("TBase"), 2)
                            grd.TextMatrix(35, 12) = Round(rs.Fields("TBase") * rs.Fields("porcentaje"), 2)
                         Case "342"
                            grd.TextMatrix(36, 10) = Round(rs.Fields("TBase"), 2)
                            grd.TextMatrix(36, 12) = Round(rs.Fields("TBase") * rs.Fields("porcentaje"), 2)
                         Case "343"
                            grd.TextMatrix(37, 10) = Round(rs.Fields("TBase"), 2)
                            grd.TextMatrix(37, 12) = Round(rs.Fields("TBase") * rs.Fields("porcentaje"), 2)
                        Case "308"
                            grd.TextMatrix(19, 10) = Round(rs.Fields("TBase"), 2)
                            grd.TextMatrix(19, 12) = Round(rs.Fields("TBase") * rs.Fields("porcentaje"), 2)
                        Case "304"
                            grd.TextMatrix(17, 10) = Round(rs.Fields("TBase"), 2)
                            grd.TextMatrix(17, 12) = Round(rs.Fields("TBase") * rs.Fields("porcentaje"), 2)
                        Case "421"
                            grd.TextMatrix(45, 10) = Round(rs.Fields("TBase"), 2)
                            grd.TextMatrix(45, 12) = Round(rs.Fields("TBase") * rs.Fields("porcentaje"), 2)
                    
                        Case 332, 333, 334
                            baseret = baseret + Round(rs.Fields("TBase"), 2)
                        
                    End Select
                    rs.MoveNext
                    Loop
                    'CASO ROLES 302
                    
                   ' VerificaExistenciaTabla (1)
                    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("BASEIR")) > 0 Then
                        Dim cad As String
                        cad = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("BASEIR")
                        sql = "SELECT"
                        sql = sql & " sum(rd.valor) as valor "
                        sql = sql & "  from RolDetalle rd INNER JOIN ELEMENTO e on e.idelemento = rd.idelemento "
                        sql = sql & " INNER JOIN GnComprobante gnc  "
                        sql = sql & " inner join gntrans gnt on gnc.codtrans=gnt.codtrans"
                        sql = sql & " ON GNC.TransID = rd.transid"
                        sql = sql & " Where gnt.modulo = 'RL' and GNC.Estado <> 3 "
                        sql = sql & " AND e.codelemento = '" & cad & "'"
                        sql = sql & cond
                     '   gobjMain.EmpresaActual.Execute sql
                        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                        Do While Not rs.EOF
                            If Not IsNull(rs.Fields("valor")) Then grd.TextMatrix(15, 10) = Round(rs.Fields("valor"), 2)
                            rs.MoveNext
                        Loop
                    End If
                    
                    VerificaExistenciaTabla (2)
                        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("IR")) > 0 Then
                        cad = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("IR")
                        sql = "SELECT"
                        sql = sql & " sum(rd.valor) as valor "
                        sql = sql & "  from RolDetalle rd INNER JOIN ELEMENTO e on e.idelemento = rd.idelemento "
                        sql = sql & " INNER JOIN GnComprobante gnc  "
                        sql = sql & " inner join gntrans gnt on gnc.codtrans=gnt.codtrans"
                        sql = sql & " ON GNC.TransID = rd.transid"
                        sql = sql & " Where gnt.modulo = 'RL' and GNC.Estado <> 3 "
                        sql = sql & " AND e.codelemento = '" & cad & "'"
                        sql = sql & cond
                      '  gobjMain.EmpresaActual.Execute sql
                       ' sql = "Select  T1.codelemento,sum(t1.valor) as BASE, SUM(T2.valor) as IR  from TMP1 T1 inner join TMP2 T2 on T1.idempleado = T2.idempleado"
                      '  sql = sql & " Where t2.valor <> 0 GROUP BY T1.CODELEMENTO"
                      Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                      
                      Do While Not rs.EOF
                        
                        If Not IsNull(rs.Fields("valor")) Then grd.TextMatrix(15, 12) = Round(rs.Fields("valor"), 2)
                        rs.MoveNext
                      Loop
                        'calula fondo
                    End If
                    
                
            ''--------COMPRAS SIN RETENCION 332,333,334,337
            sql = " Select  "
            sql = sql & " SUM(vw.CostoTotal) As CostoTotal "
            sql = sql & " from    (  gncomprobante Gnc inner join gntrans gnt on gnc.codtrans=gnt.codtrans"
            sql = sql & " inner join vwConsSUMIVKardexIVAreal vw ON Gnc.TransID = vw.transid "
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            'sql = sql & " WHERE GNC.Estado<>3 and anexocodtipotrans='1' and (aNE.BandCompraSinRetencion = 1  or CodTipoComp='41')" & Cond
            sql = sql & " WHERE GNC.Estado<>3 and anexocodtipotrans='1' and (aNE.BandCompraSinRetencion = 1  )" & cond
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If Not IsNull(rs.Fields("CostoTotal")) Then grd.TextMatrix(33, 10) = Round(rs.Fields("CostoTotal") + baseret, 2) '332
            
            CalcularPorcentajes1032010
            nombre = "03ORI_" & MonthName(DatePart("M", objcond.fecha1)) & Year(objcond.fecha1) & ".XML"
            txtDestino.Text = mRutaDestino103 & nombre
            grd.Refresh
    End With
End Sub

Private Sub CalcularPorcentajes1032010()
    Dim i As Integer, TotalRet As Currency
    Dim SubTotal As Currency
    Dim subtotal1 As Currency
    Dim SubtotalRetenido As Currency, TotalCompras As Currency
    SubTotal = 0
    subtotal1 = 0
    For i = 15 To 38
        SubTotal = SubTotal + Round(grd.ValueMatrix(i, 10), 2)
        subtotal1 = subtotal1 + Round(grd.ValueMatrix(i, 12), 2)
    Next i
    grd.TextMatrix(39, 10) = SubTotal
    grd.TextMatrix(39, 12) = subtotal1
    SubTotal = 0
    subtotal1 = 0
    For i = 41 To 46
        SubTotal = SubTotal + Round(grd.ValueMatrix(i, 10), 2)
        subtotal1 = subtotal1 + Round(grd.ValueMatrix(i, 12), 2)
    Next
    grd.TextMatrix(47, 10) = SubTotal
    grd.TextMatrix(47, 12) = subtotal1
    
    grd.TextMatrix(49, 12) = Round(grd.ValueMatrix(39, 12) + grd.ValueMatrix(47, 12), 2) '499
    
    grd.TextMatrix(57, 12) = Round(grd.ValueMatrix(49, 12) - grd.ValueMatrix(53, 6), 2) '902
    
    grd.TextMatrix(60, 12) = Round(grd.ValueMatrix(57, 12) + grd.ValueMatrix(58, 12) + grd.ValueMatrix(59, 12), 2) '999
    
    grd.TextMatrix(62, 12) = grd.TextMatrix(60, 12) '905
    
End Sub
Private Sub LLENADATOS1032010()
    With grd
        .Redraw = flexRDBuffered
        .TextMatrix(10, 2) = gobjMain.EmpresaActual.GNOpcion.ruc
        .TextMatrix(10, 4) = gobjMain.EmpresaActual.GNOpcion.NombreEmpresa
        .TextMatrix(70, 4) = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("IDRepreLegal")
        .TextMatrix(70, 9) = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("RUCContdor")
         .Refresh
    End With
End Sub

Private Sub Cargadatos500_510(ByVal param1 As String, ByVal param2 As String, ByVal param3 As String, ByVal cond As String)
Dim sql As String
Dim rs As Recordset
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(param1) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(param2) & ")" & _
                  "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            Set rs = Nothing
            '--datos de la compras brutas
            sql = "Select "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else vw.SignoCompra * (vw.CostoTotalBaseIVA  + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS Valor12 "
            sql = sql & " Into tmp1"
            sql = sql & " from (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(param2) & ")"
            sql = sql & " and GNC.Estado<>3 " & cond
            sql = sql & " and pcprovcli.tipoprovcli<>'RISE'"
            sql = sql & " and ane.codcredtrib<>'02'"
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
           sql = " Select  isnull(sum(Valor12),0) as ValorTotal12 from tmp1   "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "WAY") > 0 Then
                grd.TextMatrix(38, 10) = Round(rs.Fields("ValorTotal12"), 2) '503
            Else
                'grd.TextMatrix(36, 10) = Round(rs.Fields("ValorTotal12"), 2)
                grd.TextMatrix(36, 8) = Round(rs.Fields("ValorTotal12"), 2) 'CON IVA
            End If
            
            'NOTA DE CREDITO
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(param1) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(param3) & ")" & _
                  "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else vw.SignoCompra * (vw.CostoTotalBaseIVA  + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS Valor12 "
            sql = sql & " Into tmp1"
            sql = sql & " from (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(param3) & ")"
            sql = sql & " and pcprovcli.tipoprovcli<>'RISE'"
            sql = sql & " and ane.codcredtrib='01'"
            sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select isnull(sum(Valor12),0) as ValorTotal12 from tmp1  "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            grd.TextMatrix(36, 10) = grd.ValueMatrix(36, 8) - Abs(rs!valortotal12)
            Set rs = Nothing
End Sub

Private Sub Cargadatos501_511(ByVal param1 As String, ByVal param2 As String, ByVal param3 As String, ByVal cond As String)
Dim sql As String
Dim rs As Recordset
'--datos de la compra activos tarifa 12
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(param1) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(param2) & ")" & _
                  "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            'COMPRAS DE ACTIVOS
            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As AFValor0, "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBaseIVA  + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS AFValor12 "
            sql = sql & " Into tmp1"
            sql = sql & " from (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(param2) & ")"
            sql = sql & " and  ane.CodCredTrib not in ('02','07')"
            sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  isnull(sum(AFValor0),0) as AFValorTotal0, isnull(sum(AFValor12),0) as AFValorTotal12 from tmp1  "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            '517
            grd.TextMatrix(37, 8) = Round(rs.Fields("AFValorTotal12"), 2)
            grd.TextMatrix(37, 10) = Round(rs.Fields("AFValorTotal12"), 2)  'PONGO LO MISMO PORQUE NO TENGO NC DE ACTIVOS FIJOS
            Set rs = Nothing
End Sub

Private Sub Cargadatos502_512(ByVal param1 As String, ByVal param2 As String, ByVal param3 As String, ByVal cond As String)
Dim sql As String
Dim rs As Recordset
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(param1) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(param2) & ")" & _
                  "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            Set rs = Nothing
            sql = "Select "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0SCT, "
            sql = sql & " Case vw.CostoTotalBaseNoIVA When 0 then 0 else  vw.SignoCompra * (vw.CostoTotalBaseNoIVA + (vw.CostoTotalBaseNoIVA * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))))  end As ValorNoIvaSCT, "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else vw.SignoCompra * (vw.CostoTotalBaseIVA  + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS Valor12SCT"
            sql = sql & " Into tmp1"
            sql = sql & " from (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(param2) & ")"
            sql = sql & " and GNC.Estado<>3 " & cond
            sql = sql & " and pcprovcli.tipoprovcli<>'RISE'"
            sql = sql & " and ane.codcredtrib in ('02','07')"
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            
            sql = " Select  isnull(sum(Valor0SCT),0) as ValorTotal0SCT,isnull(sum(ValorNoIvaSCT),0) as ValorTotalNoIvaSCT,isnull(sum(Valor12SCT),0) as ValorTotal12SCT from tmp1   "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            '503
            grd.TextMatrix(38, 8) = Round(rs!ValorTotal12SCT, 2)
            
            '********** compras notas de credito SCT=02
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(param1) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(param3) & ")" & _
                  "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0SCT, "
            sql = sql & " Case vw.CostoTotalBaseNOIVA When 0 then 0 else vw.SignoCompra * (vw.CostoTotalBaseNOIVA + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As ValorNoIvaSCT, "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else vw.SignoCompra * (vw.CostoTotalBaseIVA  + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS Valor12SCT "
            sql = sql & " Into tmp1"
            sql = sql & " from (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(param3) & ")"
            sql = sql & " and pcprovcli.tipoprovcli<>'RISE'"
            sql = sql & " and ane.codcredtrib in ('02','07')"
            sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, 1 'AQUI ESTOY
            sql = " Select  isnull(sum(Valor0SCT),0) as ValorTotal10SCT,isnull(sum(ValorNoIVASCT),0) as ValorNOIVASCT, "
            sql = sql & " isnull(sum(Valor12SCT),0) as ValorTotal12SCT from tmp1"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            Dim NCValor0SCT As Currency
            NCValor0SCT = Abs(rs!ValorTotal10SCT) 'para 507
            'grd.TextMatrix(36, 10) = grd.ValueMatrix(36, 8) - Abs(rs!valortotal12) - DescFact
            '(513)
            grd.TextMatrix(38, 10) = grd.ValueMatrix(38, 8) - Abs(rs!ValorTotal12SCT)
End Sub

Private Sub Cargadatos503_513(ByVal param1 As String, ByVal param2 As String, ByVal param3 As String, ByVal cond As String)
'PONER AQUI CODIGO
End Sub

Private Sub Cargadatos504_514(ByVal param1 As String, ByVal param2 As String, ByVal param3 As String, ByVal cond As String)
Dim sql As String
Dim rs As Recordset
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(param1) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(param2) & ")" & _
                  "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0, "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBaseIVA  + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS Valor12 "
            sql = sql & " Into tmp1"
            sql = sql & " from    (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " left join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(param2) & ")"
                     
           
'           sql = sql & " and  ane.CodCredTrib not in ('02','07')"
            Dim cpImpBien0 As Currency

            sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  isnull(sum(Valor0),0) as ValorTotal0, isnull(sum(Valor12),0) as ValorTotal12 from tmp1  "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            '504
            grd.TextMatrix(40, 8) = Round(rs.Fields("ValorTotal12"), 2)
            grd.TextMatrix(40, 10) = Round(rs.Fields("ValorTotal12"), 2) 'pongo lomismo no tengo notas de credito
            Set rs = Nothing
            
'            cpImpBien0 = rs.Fields("ValorTotal0")
End Sub

Private Sub Cargadatos505_515(ByVal param1 As String, ByVal param2 As String, ByVal param3 As String, ByVal cond As String)
Dim sql As String
Dim rs As Recordset
'--IMPORTACIONES ACTIVOS FIJOS GRABADOS 12%
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 and "
            sql = sql & " GNC.CodTrans IN (" & PreparaCadena(param2) & ")"
            'sql = sql & " or  GNC.CodTrans IN (" & PreparaCadena(CP_t) & ")"
            'sql = sql & " or  GNC.CodTrans IN (" & PreparaCadena(CP_Ser) & "))"
            sql = sql & " AND ivr.CodRecargo IN (" & PreparaCadena(param1) & ") " & cond
            sql = sql & " Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0AFImp, "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBaseIVA + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS Valor12AFImp "
            sql = sql & " Into tmp1"
            sql = sql & " from    ( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            'sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(param2) & ")"
            'sql = sql & " or  GNC.CodTrans IN (" & PreparaCadena(CP_Rise) & ")"
            'sql = sql & " or  GNC.CodTrans IN (" & PreparaCadena(CP_Ser) & "))"
            'sql = sql & " and  anexos.CodCredTrib in ('02','07')"
            sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  isnull(sum(Valor0AFImp),0) as ValorTotal0AFImp, isnull(sum(Valor12AFImp),0) as ValorTotal12AFIMp from tmp1  "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            '505
            grd.TextMatrix(41, 8) = Round(Abs(rs.Fields("ValorTotal12AFimp")), 2)
            '515
            grd.TextMatrix(41, 10) = Round(Abs(rs.Fields("ValorTotal12AFImp")), 2) 'pongo lomismo hasta ver si hay nc
            '506 aqui mismo sumo con Impbienes 0%
            'grd.TextMatrix(41, 8) = Round(Abs(rs.Fields("ValorTotal0AFimp")) + cpImpBien0, 2)
            '515
            'grd.TextMatrix(41, 10) = Round(Abs(rs.Fields("ValorTotal0AFImp")) + cpImpBien0, 2) 'pongo lomismo hasta ver si hay nc
'            grd.TextMatrix(45, 10) = ComprasNoIva - NCComprasNoIVA
End Sub

Private Sub Cargadatos506_516(ByVal param1 As String, ByVal param2 As String, ByVal param3 As String, ByVal cond As String)
Dim sql As String
Dim rs As Recordset, rsAF As Recordset
            'aqui para bienes 0%
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(param1) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(param2) & ")" & _
                  "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0 "
            sql = sql & " Into tmp1"
            sql = sql & " from (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " left join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(param2) & ")"
            sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  isnull(sum(Valor0),0) as ValorTotal0 from tmp1  "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                        
            'aqui para activos fijos 0%
            VerificaExistenciaTabla 2
            VerificaExistenciaTabla 3
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp2 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 and "
            sql = sql & " GNC.CodTrans IN (" & PreparaCadena(param3) & ")"
            sql = sql & " AND ivr.CodRecargo IN (" & PreparaCadena(param1) & ") " & cond
            sql = sql & " Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0AFImp "
            sql = sql & " Into tmp3"
            sql = sql & " from    ( tmp2 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp2.TransID = Gnc.TransID)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(param3) & ")"
            sql = sql & " and GNC.Estado<>3 " & cond

            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            
            sql = " Select  isnull(sum(Valor0AFImp),0) as ValorTotal0AFImp from tmp3  "
            Set rsAF = gobjMain.EmpresaActual.OpenRecordset(sql)
            grd.TextMatrix(42, 8) = Round(Abs(rs.Fields("ValorTotal0") + rsAF.Fields("ValorTotal0AFImp")), 2)
            grd.TextMatrix(42, 10) = Round(Abs(rs.Fields("ValorTotal0") + rsAF.Fields("ValorTotal0AFImp")), 2) 'pongo lomismo hasta ver si hay nc
            
            Set rs = Nothing
            Set rsAF = Nothing
    End Sub

Private Sub Cargadatos507_517(ByVal param1 As String, ByVal param2 As String, ByVal param3 As String, ByVal Param4 As String, ByVal cond As String)
Dim sql As String
Dim rs As Recordset
Dim Compras0 As Currency
Dim ComprasAF0 As Currency
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(param1) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(param2) & ")" & _
                  "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            Set rs = Nothing
            '--datos de la compras brutas
            sql = "Select "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0 "
            sql = sql & " Into tmp1"
            sql = sql & " from (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(param2) & ")"
            sql = sql & " and GNC.Estado<>3 " & cond
            sql = sql & " and pcprovcli.tipoprovcli<>'RISE'"
'            sql = sql & " and ane.codcredtrib<>'02'"
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  isnull(sum(Valor0),0) as ValorTotal0 from tmp1   "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            Compras0 = rs.Fields("ValorTotal0")
            
            'compras activos 0%
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(param1) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(param3) & ")" & _
                  "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As AFValor0 "
            sql = sql & " Into tmp1"
            sql = sql & " from (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(param3) & ")"
            sql = sql & " and  ane.CodCredTrib not in ('02','07')"
            sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  isnull(sum(AFValor0),0) as AFValorTotal0 from Tmp1 "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            '517
            ComprasAF0 = rs.Fields("AFValorTotal0")
            grd.TextMatrix(43, 8) = Round(Compras0 + ComprasAF0, 2)
            'grd.TextMatrix(37, 10) = Round(rs.Fields("AFValorTotal12"), 2)  'PONGO LO MISMO PORQUE NO TENGO NC DE ACTIVOS FIJOS
                        
            
            Set rs = Nothing
            
            'NOTA DE CREDITO SOLO DE BIENES TODAVIA NO TENGO DE AF
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(param1) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(Param4) & ")" & _
                  "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0 "
            sql = sql & " Into tmp1"
            sql = sql & " from (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(Param4) & ")"
            sql = sql & " and pcprovcli.tipoprovcli<>'RISE'"
            'sql = sql & " and ane.codcredtrib<>'02'"
           sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  isnull(sum(Valor0),0) as ValorTotal0 from tmp1  "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            grd.TextMatrix(43, 10) = Round(grd.TextMatrix(43, 8) - Abs(rs!valortotal0), 2)
            Set rs = Nothing
End Sub
Private Sub Cargadatos508_518(ByVal param1 As String, ByVal param2 As String, ByVal param3 As String, ByVal cond As String)
Dim rs As Recordset
Dim sql As String
    VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(param1) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(param2) & ")" & _
                  "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            Set rs = Nothing
            '--datos de la compras brutas
            sql = "Select "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0 "
            sql = sql & " Into tmp1"
            sql = sql & " from (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(param2) & ")"
            sql = sql & " and GNC.Estado<>3 " & cond
            sql = sql & " and pcprovcli.tipoprovcli='RISE'"
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  isnull(sum(Valor0),0) as ValorTotal0 from tmp1   "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                grd.TextMatrix(44, 8) = Round(rs.Fields("ValorTotal0"), 2) 'CON IVA
            Set rs = Nothing
            'NOTA DE CREDITO
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(param1) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(param3) & ")" & _
                  "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0 "
            sql = sql & " Into tmp1"
            sql = sql & " from (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(param3) & ")"
            sql = sql & " and pcprovcli.tipoprovcli='RISE'"
           sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  isnull(sum(Valor0),0) as ValorTotal0 from tmp1  "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            grd.TextMatrix(44, 10) = Round(grd.ValueMatrix(44, 8) - Abs(rs!valortotal0), 2)
            Set rs = Nothing
End Sub

Private Sub Cargadatos531_541(ByVal param1 As String, ByVal param2 As String, ByVal param3 As String, ByVal cond As String)
Dim rs As Recordset
Dim sql As String
    VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(param1) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(param2) & ")" & _
                  "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            Set rs = Nothing
            '--datos de la compras brutas
            sql = "Select "
            sql = sql & " Case vw.CostoTotalBaseNoIva When 0 then 0 else vw.SignoCompra * (vw.CostoTotalBaseNoIva + (vw.CostoTotalBaseNoIva * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As ValorNoIva "
            sql = sql & " Into tmp1"
            sql = sql & " from (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(param2) & ")"
            sql = sql & " and GNC.Estado<>3 " & cond
            'sql = sql & " and pcprovcli.tipoprovcli='RISE'"
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  isnull(sum(ValorNoIva),0) as ValorTotalNoIva from tmp1   "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            grd.TextMatrix(46, 8) = Round(rs.Fields("ValorTotalNoiva"), 2)
            Set rs = Nothing
            
            'NOTA DE CREDITO
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(param1) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(param3) & ")" & _
                  "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBaseNoiva When 0 then 0 else vw.SignoCompra * (vw.CostoTotalBaseNoIva + (vw.CostoTotalBaseNoIva * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As ValorNoIva "
            sql = sql & " Into tmp1"
            sql = sql & " from (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(param3) & ")"
            'sql = sql & " and pcprovcli.tipoprovcli='RISE'"
            sql = sql & " and GNC.Estado<>3 " & cond
            
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  isnull(sum(ValorNoIVA),0) as ValorTotalNoIva from tmp1  "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            grd.TextMatrix(46, 10) = Round(grd.ValueMatrix(46, 8) - Abs(rs!ValorTotalnoiva), 2)
            Set rs = Nothing
End Sub

Private Sub Cargadatos504_514Hormi(ByVal param1 As String, ByVal param2 As String, ByVal param3 As String, ByVal cond As String)
Dim sql As String
Dim rs As Recordset
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo ='IVA' " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(param2) & ")" & _
                  "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0, "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " tmp0.totaldescuento end AS ValorIVA "
            sql = sql & " Into tmp1"
            sql = sql & " from    (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " left join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(param2) & ")"

            sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  isnull(sum(Valor0),0) as ValorTotal0, isnull((sum(ValorIVA)*100)/12,0) as ValorTotal12 from tmp1  "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            '504
                grd.TextMatrix(40, 8) = Round(rs.Fields("ValorTotal12"), 2)
               grd.TextMatrix(40, 10) = Round(rs.Fields("ValorTotal12"), 2) 'pongo lomismo no tengo notas de credito
            Set rs = Nothing
End Sub

Private Sub Cargadatos505_515Hormi(ByVal param1 As String, ByVal param2 As String, ByVal param3 As String, ByVal cond As String)
Dim sql As String
Dim rs As Recordset
'--IMPORTACIONES ACTIVOS FIJOS GRABADOS 12%
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND "
            sql = sql & " GNC.CodTrans IN (" & PreparaCadena(param2) & ")"
            sql = sql & " AND ivr.CodRecargo ='IVA' " & cond
            sql = sql & " Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0AFImp, "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " tmp0.totaldescuento end AS ValorIVAAFImp "
            sql = sql & " Into tmp1"
            sql = sql & " from    ( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(param2) & ")"
            sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  isnull(sum(Valor0AFImp),0) as ValorTotal0AFImp, isnull((sum(ValorIVAAFImp)*100)/12,0) as ValorTotal12AFIMp from tmp1  "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            grd.TextMatrix(41, 8) = Round(Abs(rs.Fields("ValorTotal12AFimp")), 2)
            grd.TextMatrix(41, 10) = Round(Abs(rs.Fields("ValorTotal12AFImp")), 2) 'pongo lomismo hasta ver si hay nc
End Sub

Private Sub Cargadatos407_417(ByVal param1 As String, ByVal param2 As String, ByVal param3 As String, ByVal cond As String)
Dim sql As String
Dim rs As Recordset
Dim Exportaciones As Currency
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(param1) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(param2) & ")" & _
                  "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            Set rs = Nothing
            'Explicacion solo para waynoro
            'se va toma el preciobaseiva porque el  item esta configurado con IVA
            sql = "Select "
            sql = sql & " Case vw.PrecioTotalBaseIVA When 0 then 0 else vw.SignoVenta * (vw.PrecioTotalBaseIVA + (vw.PrecioTotalBaseIVA * (cast( isnull(TotalDescuento,0) as float) / cast(vw.PrecioTotal as float))) ) end As ValorIVA "
            sql = sql & " Into tmp1"
            sql = sql & " from (tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(param2) & ")"
            sql = sql & " and GNC.Estado<>3 " & cond

            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  isnull(sum(ValorIVA),0) as ValorTotalIVA from tmp1   "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            Exportaciones = Abs(rs.Fields("ValorTotalIVA"))
            '407
            grd.TextMatrix(20, 8) = Round(Exportaciones, 2)
            Set rs = Nothing
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(param1) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(param3) & ")" & _
                  "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  "
            sql = sql & " Case vw.PrecioTotalBase0 When 0 then 0 else vw.SignoVenta * (vw.PrecioTotalBase0 + (vw.PrecioTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.PrecioTotal as float))) ) end As ValorIVA "
            sql = sql & " Into tmp1"
            sql = sql & " from ( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(param3) & ")"
            sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            sql = " Select  isnull(sum(ValorIVA),0) as ValorTotalIVA from tmp1  "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            grd.TextMatrix(20, 10) = Round(Exportaciones - Abs(rs!valortotalIVA), 2)
            Set rs = Nothing
End Sub

Public Sub Imprimir()
'    grd.Rows = grd.FixedRows
    Select Case Me.tag
    Case "F1032010"
    Case "F1042010"
        FrmImprimeEtiketas.InicioF104 grd
    End Select
End Sub

