VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAjustesItem 
   Caption         =   "Ajustes / Bajas de Items con Existencias Negativas"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   6240
   WindowState     =   2  'Maximized
   Begin SiiToolsA.IVAjuste IVAjuste 
      Height          =   2775
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   4895
   End
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   2055
      Left            =   900
      TabIndex        =   3
      Top             =   1860
      Width           =   5055
      _cx             =   8911
      _cy             =   3619
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
      Left            =   5520
      Top             =   120
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjustesItem.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjustesItem.frx":0114
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjustesItem.frx":0568
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjustesItem.frx":067C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjustesItem.frx":0790
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   953
      ButtonWidth     =   1191
      ButtonHeight    =   953
      Style           =   1
      ImageList       =   "img1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar"
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar (F5)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ajustar"
            Key             =   "Ajustar"
            Description     =   "Asignar un valor"
            Object.ToolTipText     =   "Asignar un valor (F6)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar"
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar (F3)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Imprimir"
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir (Ctrl+P)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
      ScaleWidth      =   6240
      TabIndex        =   0
      Top             =   4185
      Width           =   6240
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
   Begin VSFlex7LCtl.VSFlexGrid grdNegativos 
      Height          =   2055
      Left            =   780
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   5055
      _cx             =   8911
      _cy             =   3619
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
   Begin VSFlex7LCtl.VSFlexGrid grdKardex 
      Height          =   2055
      Left            =   60
      TabIndex        =   5
      Top             =   600
      Width           =   5055
      _cx             =   8911
      _cy             =   3619
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
   Begin SiiToolsA.IVAjuste IVBaja 
      Height          =   2775
      Left            =   120
      TabIndex        =   7
      Top             =   5280
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   4895
   End
   Begin SiiToolsA.IVFISICO IVFisico 
      Height          =   1695
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   2990
   End
End
Attribute VB_Name = "frmAjustesItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mProcesando As Boolean
Private mCancelado As Boolean
Private mcolItemsSelec As Collection      'Coleccion de items
Private mObjCond As RepCondicion
Private mobjBusq As Busqueda

'Private mobjItem As IVinventario

Public Sub Inicio(ByVal tag As String)
    Dim i As Integer
    On Error GoTo Errtrap
    
    Me.tag = tag            'Guarda en la propiedad Tag para distinguir después
    Me.Show
    Me.ZOrder
    
    Select Case Me.tag
    Case "AjustesInventario"
        Me.Caption = "Ajustes / Bajas de Items con Existencias Negativas"
    End Select
       
    'Inicializa la grilla
    grdKardex.Rows = grdKardex.FixedRows
    grd.Rows = grd.FixedRows
    grdNegativos.Rows = grdNegativos.FixedRows
    ConfigCols
    
    Exit Sub
Errtrap:
    DispErr
    Unload Me
    Exit Sub
End Sub



Private Sub Form_Initialize()
    Set mObjCond = New RepCondicion
    Set mobjBusq = New Busqueda
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF3
        Grabar
        KeyCode = 0
    Case vbKeyF5
        Select Case Me.tag
            Case "AjustesInventario": Buscar
        End Select
        KeyCode = 0
    Case vbKeyF6
        Ajustar
        KeyCode = 0
    Case vbKeyP
        If Shift And vbCtrlMask Then
            Imprimir
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
    grd.Move 0, tlb1.Height, Me.ScaleWidth, (Me.ScaleHeight - tlb1.Height - pic1.Height - 80)
'    grdNegativos.Move 0, grdKardex.Top + grdKardex.Height, Me.ScaleWidth, (Me.ScaleHeight - tlb1.Height - pic1.Height - 80) / 3
'    grd.Move 0, grdNegativos.Top + grdKardex.Height, Me.ScaleWidth, (Me.ScaleHeight - tlb1.Height - pic1.Height - 80) / 3
    grdNegativos.Visible = False
    grdKardex.Visible = False
    prg1.Width = Me.ScaleWidth - (prg1.Left * 2)
End Sub



Private Sub grdkardex_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row < grdKardex.FixedRows Then Cancel = True
    If grdKardex.IsSubtotal(Row) = True Then Cancel = True
    If grdKardex.ColData(Col) < 0 Then Cancel = True
    
    If Not Cancel Then
        'Longitud maxima para editar
        grdKardex.EditMaxLength = grdKardex.ColData(Col)
    End If
End Sub

Private Sub grdkardex_BeforeSort(ByVal Col As Long, Order As Integer)
    'Impide mientras está procesando
    If mProcesando Then Order = flexSortNone
End Sub

Private Sub tlb1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Buscar":
        Select Case Me.tag
            Case "AjustesInventario": Buscar
        End Select
    Case "Ajustar":     Ajustar
    Case "Grabar":
        Select Case Me.tag
            Case "AjustesInventario": Grabar
        End Select
    Case "Imprimir":    Imprimir
    Case "Cerrar":      Cerrar
    End Select
End Sub

Private Sub Buscar()
    Dim i As Long
    With grd
        For i = .FixedRows To .Rows - 1
            .RowData(i) = 0
        Next i
        .Rows = .FixedRows
    End With

    With grdKardex
        For i = .FixedRows To .Rows - 1
            .RowData(i) = 0
        Next i
        .Rows = .FixedRows
    End With

    With grdNegativos
        For i = .FixedRows To .Rows - 1
            .RowData(i) = 0
        Next i
        .Rows = .FixedRows
    End With


            gobjMain.objCondicion.Tipo = bsqIVKardex
            If Not mobjBusq.Show(gobjMain) Then
                grdKardex.SetFocus
                Exit Sub
            End If
        MensajeStatus MSG_BUSCANDO, vbHourglass

    Cargar_IVListadoAjuste
    LlenaNegativos
    MensajeStatus
End Sub


Private Sub ConfigCols()
    Dim s As String, i As Long, j As Integer
    With grdKardex
    Select Case Me.tag
        Case "AjustesInventario"
            s = "^#|<Modulo|<Codigo|<Item|<Fecha|<(CodT)|<Trans|<#Ref|<Nombre|<Descripcion"
            s = s & "|<Bodega|>Ingreso|>Egreso|>Saldo|>Costo Unit.|>Costo Total I |>Costo Total E"
            s = s & "|>Costo Total|CostoPromedio|>Cotizacion|^Estado|Orden|HoraTrans"
        End Select
        .FormatString = s
        
        GNPoneNumFila grdKardex, False
'        AjustarAutoSize grdkardex, -1, -1, 4000
        AsignarTituloAColKey grdKardex
    
        .ColHidden(.ColIndex("Modulo")) = True
        .ColHidden(.ColIndex("(CodT)")) = True
        .ColHidden(.ColIndex("#Ref")) = True
        .ColHidden(.ColIndex("Nombre")) = True
        .ColHidden(.ColIndex("Costo Unit.")) = True
        .ColHidden(.ColIndex("Costo Total I")) = True
        .ColHidden(.ColIndex("Costo Total E")) = True
        .ColHidden(.ColIndex("Costo Total")) = True
        .ColHidden(.ColIndex("CostoPromedio")) = True
        .ColHidden(.ColIndex("Cotizacion")) = True
        .ColHidden(.ColIndex("Orden")) = True
        
        .ColFormat(.ColIndex("Saldo")) = "#0.00"
        .ColFormat(.ColIndex("Ingreso")) = "#0.00"
        .ColFormat(.ColIndex("Egreso")) = "#0.00"
        
        .ColHidden(.ColIndex("HoraTrans")) = True
        For i = 1 To .ColIndex("HoraTrans")
            .ColWidth(i) = 1000
        Next i
        .ColWidth(.ColIndex("#")) = 500
        .ColWidth(.ColIndex("Codigo")) = 1500
        .ColWidth(.ColIndex("Item")) = 2500
        .ColWidth(.ColIndex("Descripcion")) = 3500
        .ColWidth(.ColIndex("Estado")) = 500
        'Columnas modificables (Longitud maxima)
        Select Case Me.tag
            Case "AjustesInventario"
                .ColData(.ColIndex("IVA")) = 5
        End Select
        'Columnas No modificables
        For i = 0 To .ColIndex("HoraTrans")
            .ColData(i) = -1
        Next i
        
        
        .ColFormat(.ColIndex("Saldo")) = "#0.00"
        
        'Color de fondo
        If .Rows > .FixedRows Then
            .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("HoraTrans")) = .BackColorFrozen
        End If
        
        .Subtotal flexSTSum, 2, 1, , grdKardex.GridColor, vbBlack, , "Subtotal", 1, True
    
    End With
        
    With grdNegativos
        .Cols = 4
        s = "^#|<Codigo|<Item|>Saldo"
        .FormatString = s
        AsignarTituloAColKey grdNegativos
        .ColWidth(.ColIndex("#")) = 500
        .ColWidth(.ColIndex("Codigo")) = 1500
        .ColWidth(.ColIndex("Item")) = 2500
        .ColWidth(.ColIndex("Saldo")) = 1500
        .Subtotal flexSTSum, 2, 1, , grdKardex.GridColor, vbBlack, , "Subtotal", 1, True
        .ColFormat(.ColIndex("Saldo")) = "#0.00"
        .Refresh
        
    End With
    With grd
        .Cols = 5
         s = "^#|<Codigo|<Item|>Saldo|>Resultado"
        .FormatString = s
        AsignarTituloAColKey grd
        .ColWidth(.ColIndex("#")) = 500
        .ColWidth(.ColIndex("Codigo")) = 1500
        .ColWidth(.ColIndex("Item")) = 2500
        .ColWidth(.ColIndex("Saldo")) = 1500
        .ColWidth(.ColIndex("Resultado")) = 2500
        .ColFormat(.ColIndex("Saldo")) = "#0.00"
        For i = 0 To .ColIndex("Saldo")
            .ColData(i) = -1
        Next i
        If .Rows > .FixedRows Then
            .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("Saldo")) = .BackColorFrozen
        End If
        .Refresh
    End With

End Sub

Private Sub Ajustar()
    Dim i As Long, saldo As Currency
    Dim COL_HABER As Integer, COL_SALDO As Integer
    Dim v As Currency, saldo_t As Currency
    Dim Fila As Long
    COL_HABER = 12
    COL_SALDO = 13
Fila = 1
    With grd
'        .Clear 0
        '.Rows = 1
        For i = grdNegativos.FixedRows To grdNegativos.Rows - 1
            If grdNegativos.IsSubtotal(i) Then
                    .AddItem Fila & vbTab & grdNegativos.TextMatrix(i - 1, grdNegativos.ColIndex("Codigo")) & vbTab & grdNegativos.TextMatrix(i - 1, grdNegativos.ColIndex("Item")) & vbTab & grdNegativos.TextMatrix(i - 1, grdNegativos.ColIndex("Saldo")), Fila
                    Fila = Fila + 1
            End If
        Next i
'        .SubtotalPosition = flexSTBelow '= flexSTAbove
'        .Subtotal flexSTMax, 1, 3, , grdNegativos.GridColor, vbBlack, , "Cant", 1, True
        If .Rows > .FixedRows Then
            .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("Saldo")) = .BackColorFrozen
        End If

        .Refresh
    
    End With

   grdNegativos.Visible = True
'   grdkardex.Visible = False
End Sub

Private Sub AsignarIVA()
    Dim s As String, v As Single
    Dim i As Long
    
    s = InputBox("Ingrese el valor de IVA (%)", "Asignar un valor", "15")
    If IsNumeric(s) Then
        v = CSng(s)
    Else
        MsgBox "Debe ingresar un valor numérico. (ejm. 15 para 15%)", vbInformation
        grdKardex.SetFocus
        Exit Sub
    End If
    
    With grdKardex
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("IVA")) = v
        Next i
    End With
End Sub

Private Sub Grabar()
    Dim i As Long, iv As IVinventario, cod As String
    On Error GoTo Errtrap
    
    'Confirmación
    If MsgBox("Está seguro que desea grabar?", vbQuestion + vbYesNo) <> vbYes Then
        grdKardex.SetFocus
        Exit Sub
    End If
    
    'Deshabilita los botónes y menus
    Habilitar False
    mCancelado = False
    
    With grdKardex
        prg1.Min = 0
        prg1.max = 1
        If .Rows > .FixedRows Then prg1.max = .Rows - 1
        For i = .FixedRows To .Rows - 1
            'Si es que se canceló el proceso
            If mCancelado Then GoTo salida
        
            prg1.value = i
            cod = .TextMatrix(i, .ColIndex("Código"))
            MensajeStatus i & " de " & .Rows - .FixedRows, vbHourglass
            DoEvents
            
            'Recupera el objeto de Inventario
            Set iv = gobjMain.EmpresaActual.RecuperaIVInventario(cod)
            
            Select Case Me.tag
            Case "AjustesInventario"
                If iv.PorcentajeIVA <> .ValueMatrix(i, .ColIndex("IVA")) / 100 Then
                    iv.PorcentajeIVA = .ValueMatrix(i, .ColIndex("IVA")) / 100
                End If
            End Select
            iv.Grabar
        Next i
    End With
    
salida:
    MensajeStatus
    Set iv = Nothing
    Habilitar True
    Exit Sub
Errtrap:
    MensajeStatus
    DispErr
    GoTo salida
    Exit Sub
End Sub


Private Sub Habilitar(ByVal v As Boolean)
    mProcesando = Not v
    
    tlb1.Buttons("Buscar").Enabled = v
    tlb1.Buttons("Asignar").Enabled = v
    tlb1.Buttons("Grabar").Enabled = v
'    tlb1.Buttons("Imprimir").Enabled = v
    tlb1.Buttons("Imprimir").Enabled = False        '*** MAKOTO PENDIENTE Por ahora
    
    If v Then
        tlb1.Buttons("Cerrar").Caption = "Cerrar"
    Else
        tlb1.Buttons("Cerrar").Caption = "Cancelar"
    End If
    
    frmMain.mnuFile.Enabled = v
    frmMain.mnuHerramienta.Enabled = v
    frmMain.mnuTransferir.Enabled = v
    frmMain.mnuCerrarTodas.Enabled = v
    
    prg1.value = prg1.Min
End Sub

Private Sub Imprimir()

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

Private Sub Cargar_IVListadoAjuste()
    Dim i As Long
    Dim mCodMoneda  As String
    On Error GoTo Errtrap
    With grdKardex
        mCodMoneda = MONEDA_SEC
        .Redraw = False
        .Rows = .FixedRows
        gobjMain.objCondicion.CodMoneda = mCodMoneda
        'grdkardex.LoadArray MiGetRows(gobjMain.EmpresaActual.ConsIVKardex())
        MiGetRowsRep gobjMain.EmpresaActual.ConsIVKardex2Col, grdKardex
        
        'Título de columnas

        Const COL_IVK_CODITEM = 2
        Const COL_IVK_ITEM = 3
        Const COL_IVK_FECHA = 4
        Const COL_IVK_TRANS = 6
        Const COL_IVK_DEBE = 11
        Const COL_IVK_SALDOCANTIDAD = 13
        Const Col_IVK_DEBECOSTO = 15
        Const COL_IVK_SALDOCOSTO = 17
        Const COL_IVK_COSTOPROMEDIO = 18
        'ConfigColumns grdkardex, mobjReporte

        'Combina celdas del mismo valor
        .MergeCol(COL_IVK_CODITEM) = True
        .MergeCol(COL_IVK_ITEM) = True
        .MergeCol(COL_IVK_FECHA) = True
        .MergeCol(COL_IVK_TRANS) = True
        VisualizaSaldoKardex (COL_IVK_DEBE)
        VisualizaSaldoKardex (Col_IVK_DEBECOSTO)
        VisualizaCostoPromedio COL_IVK_SALDOCANTIDAD, COL_IVK_SALDOCOSTO, COL_IVK_COSTOPROMEDIO

        'fMain.MarcaVerMoneda gobjMain.objCondicion.CodMoneda
        GNPoneNumFila grdKardex, False
        mObjCond.Moneda = gobjMain.objCondicion.CodMoneda
        mObjCond.Fecha1 = gobjMain.objCondicion.Fecha1
        mObjCond.Fecha2 = gobjMain.objCondicion.Fecha2
        mObjCond.Bodega = gobjMain.objCondicion.CodBodega1
    End With
'    grdNegativos.Clear
    Exit Sub
Errtrap:
    grdKardex.Redraw = True
    DispErr
End Sub


Public Sub MiGetRowsRep(ByVal rs As Recordset, grdKardex As VSFlexGrid)
    grdKardex.LoadArray MiGetRows(rs)
    'ConfigTipoDatoCol grdkardex, rs
    grdKardex.MergeCol(1) = True
    grdKardex.SubtotalPosition = flexSTBelow '= flexSTAbove
    grdKardex.Subtotal flexSTSum, 2, 2, , grdKardex.GridColor, vbBlack, , "Subtotal", 2, True
    grdKardex.Redraw = True
            grdNegativos.SubtotalPosition = flexSTBelow '= flexSTAbove
            grdNegativos.Subtotal flexSTMax, 1, 3, , grdKardex.GridColor, vbBlack, , "Cant", 1, True
            grdNegativos.Refresh
            grdNegativos.Redraw = True
    
End Sub

Private Sub VisualizaSaldoKardex(COL_DEBE As Integer)
    'Subrutina general que visualiza el saldo
    Dim i As Long, saldo As Currency
    Dim COL_HABER As Integer, COL_SALDO As Integer
    Dim v As Currency, saldo_t As Currency
    COL_HABER = COL_DEBE + 1
    COL_SALDO = COL_DEBE + 2
    With grdKardex
        For i = .FixedRows To .Rows - 1
            If Not .IsSubtotal(i) Then
                v = .ValueMatrix(i, COL_DEBE) - .ValueMatrix(i, COL_HABER)
                saldo = saldo + v
                .TextMatrix(i, COL_SALDO) = saldo
                If saldo < 0 Then
                    .Select i, COL_SALDO
                    .CellBackColor = &HFF
               End If
            Else
                'If mobjReporte.Detalle(COL_SALDO).Subtotal Then
                    .TextMatrix(i, COL_SALDO) = saldo   'Para qu visualize el saldo
                    saldo_t = saldo_t + saldo
                'End If
                saldo = 0
                'Cargar  el saldo total
                If i = .Rows - 1 Then
                 '   If mobjReporte.Detalle(COL_SALDO).Subtotal Then
                         .TextMatrix(i, COL_SALDO) = saldo_t
                  '  End If
                End If
                If Me.tag = "ConsIVKardex" Then
                    'Columna de  costo Unitario
                    'col_saldo +1 = col_CU    col_saldo+2 = col_CT
                    'If mobjReporte.Detalle(COL_SALDO + 1).Subtotal Then
                        If Abs(.ValueMatrix(i, COL_SALDO)) > 0 Then
                            .TextMatrix(i, COL_SALDO + 1) = .ValueMatrix(i, COL_SALDO + 2) / _
                                                     .ValueMatrix(i, COL_SALDO)
                        Else
                            .TextMatrix(i, COL_SALDO + 1) = "0.00"
                        End If
                    'End If
                End If
                
            End If
        Next i
    End With
End Sub


'Sub para calcular el CostoUnitario en una columna especial, dandole como parametros de donde calcular y donde poner
Private Sub VisualizaCostoPromedio(Col_Cantidad As Integer, Col_CostoTotal As Integer, Col_CostoPromedio)
Dim i As Long

    With grdKardex
        For i = .FixedRows To .Rows - 1
            If .ValueMatrix(i, Col_Cantidad) <> 0 Then
                .TextMatrix(i, Col_CostoPromedio) = .ValueMatrix(i, Col_CostoTotal) / .ValueMatrix(i, Col_Cantidad)
            Else
                .TextMatrix(i, Col_CostoPromedio) = "0.00"
            End If
        Next i
    End With
End Sub


Private Sub LlenaNegativos()
    Dim COL_HABER As Integer, COL_SALDO As Integer
    Dim v As Currency, saldo_t As Currency, i As Long
    Dim Fila As Long
    COL_HABER = 12
    COL_SALDO = 13
    
    With grdNegativos
            .Clear 0
            For i = grdKardex.FixedRows To grdKardex.Rows - 1
                If Not grdKardex.IsSubtotal(i) Then
                    If grdKardex.TextMatrix(i, COL_SALDO) < 0 Then
                        .AddItem Fila & vbTab & grdKardex.TextMatrix(i, grdKardex.ColIndex("Codigo")) & vbTab & grdKardex.TextMatrix(i, grdKardex.ColIndex("Item")) & vbTab & grdKardex.TextMatrix(i, grdKardex.ColIndex("Saldo")) * -1
                        Fila = Fila + 1
                    End If
                End If
            Next i
            .SubtotalPosition = flexSTBelow '= flexSTAbove
            .Subtotal flexSTMax, 1, 3, , grdKardex.GridColor, vbBlack, , "Cant", 1, True
        If .Rows > .FixedRows Then
            .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("Saldo")) = .BackColorFrozen
        End If
            
            .Refresh
        
        End With
End Sub

Private Sub Procesar()
    Dim rt As Integer
    
    MensajeStatus MSG_PREPARA, vbArrowHourglass
    
    'Limpia los objetos que van a guardar el resultado del proceso
    IVAjuste.GNComprobante.BorrarIVKardex
    IVBaja.GNComprobante.BorrarIVKardex
    IVAjuste.VisualizaDesdeObjeto
    IVBaja.VisualizaDesdeObjeto

    IVFisico.EliminaFilasIncompletas
    If IVFisico.GNComprobante.CountIVKardex = 0 Then
        MsgBox "No hay filas para procesar", vbOKOnly + vbInformation
        Exit Sub
    End If
    
    If gConfigIVFisico.BandLineaAuto = False Then
        rt = MsgBox("Desea totalizar filas repetidas", vbYesNo + vbQuestion)
        If rt = vbYes Then IVFisico.TotalizarItem
    End If
    
'    If Not (mBandRevisarNoContados) Then
'        If MsgBox("Desea agregar los items con existencia y que no han sido contados fisicamente", vbYesNo) = vbYes Then
'            IVFisico.GNComprobante.FechaTrans = dtpFecha.value
'            IVFisico.CargarItemsNoContados
'            IVFisico.VisualizaDesdeObjeto
'            IVFisico.Refresh_Items
'            mBandRevisarNoContados = True
'        End If
'    End If
        ProcesarAjuste
    MensajeStatus "", vbNormal
End Sub




Private Sub ProcesarAjuste()
    Dim ix As Long, ivk As IVKardex, dif As Currency
    Dim i As Long, signo As Integer, cant As String
    Dim iv As IVinventario, c As Currency
    
    IVFisico.CargaItemsOrdenado
        
    For i = 1 To IVFisico.GNComprobante.CountIVKardex
        c = 0
        dif = 0
        cant = 0
       
        dif = IVFisico.DiferenciaExistencia(i)
        'If dif > 0 Then
            With IVAjuste
                ix = .GNComprobante.AddIVKardex
                Set ivk = IVFisico.GNComprobante.IVKardex(i)
                cant = dif
                .GNComprobante.IVKardex(ix).Cantidad = cant
                .GNComprobante.IVKardex(ix).CodBodega = ivk.CodBodega
                .GNComprobante.IVKardex(ix).CodInventario = ivk.CodInventario
                
                'Calcula el costo
                Set iv = .GNComprobante.Empresa.RecuperaIVInventario(ivk.CodInventario)
                c = iv.CostoDouble2(.GNComprobante.FechaTrans, _
                                     cant, _
                                     .GNComprobante.TransID, _
                                     .GNComprobante.HoraTrans)
            
                'Si el costo calculado está en otra moneda, convierte en moneda de trans.
                If .GNComprobante.CodMoneda <> iv.CodMoneda Then
                    c = c * .GNComprobante.Cotizacion(iv.CodMoneda) / .GNComprobante.Cotizacion(" ")
                End If
                
                .GNComprobante.IVKardex(ix).CostoTotal = c * cant
            End With
        'ElseIf dif < 0 Then
            With IVBaja
                ix = .GNComprobante.AddIVKardex
                Set ivk = IVFisico.GNComprobante.IVKardex(i)
                cant = dif
                .GNComprobante.IVKardex(ix).Cantidad = cant
                .GNComprobante.IVKardex(ix).CodBodega = ivk.CodBodega
                .GNComprobante.IVKardex(ix).CodInventario = ivk.CodInventario
                
                'Calcula el costo
                Set iv = .GNComprobante.Empresa.RecuperaIVInventario(ivk.CodInventario)
                c = iv.CostoDouble2(.GNComprobante.FechaTrans, _
                                     cant, _
                                     .GNComprobante.TransID, _
                                     .GNComprobante.HoraTrans)
            
                'Si el costo calculado está en otra moneda, convierte en moneda de trans.
                If .GNComprobante.CodMoneda <> iv.CodMoneda Then
                    c = c * .GNComprobante.Cotizacion(iv.CodMoneda) / .GNComprobante.Cotizacion(" ")
                End If
                
                .GNComprobante.IVKardex(ix).CostoTotal = c * cant
            End With
        'Else
            'Si es cero no hace nada
        'End If
    Next i
    
    IVAjuste.VisualizaDesdeObjeto
    IVBaja.VisualizaDesdeObjeto
End Sub




