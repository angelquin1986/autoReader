VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Begin VB.Form frmB_Desintegridad 
   Caption         =   "Buscar Desintegridad"
   ClientHeight    =   5970
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10185
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   375
      Left            =   8760
      TabIndex        =   7
      Top             =   5520
      Width           =   1200
   End
   Begin VB.Frame Frame3 
      Caption         =   "Registros"
      Height          =   4095
      Left            =   4920
      TabIndex        =   12
      Top             =   1200
      Width           =   5175
      Begin VSFlex7Ctl.VSFlexGrid grdRegistros 
         Height          =   3615
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   4935
         _cx             =   8705
         _cy             =   6376
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
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   1
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
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
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tablas"
      Height          =   4695
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   4695
      Begin VSFlex7Ctl.VSFlexGrid grdCat_Trans 
         Height          =   3615
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   4455
         _cx             =   7858
         _cy             =   6376
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
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   1
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
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
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.CommandButton cmdVerReg 
         Caption         =   "&Ver Registros"
         Height          =   495
         Left            =   1440
         TabIndex        =   5
         Top             =   4080
         Width           =   1500
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   10185
      TabIndex        =   8
      Top             =   0
      Width           =   10185
      Begin VB.Frame Frame1 
         Caption         =   "Seleccione la Cuenta Contable"
         Height          =   975
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   9975
         Begin VB.OptionButton optTransaccion 
            Caption         =   "Por Transacción"
            Height          =   255
            Left            =   6600
            TabIndex        =   2
            Top             =   600
            Width           =   1575
         End
         Begin VB.OptionButton optCatalogo 
            Caption         =   "Por &Catálogo"
            Height          =   255
            Left            =   6600
            TabIndex        =   1
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "&Buscar - F5"
            Height          =   375
            Left            =   8640
            TabIndex        =   3
            Top             =   240
            Width           =   1200
         End
         Begin FlexComboProy.FlexCombo fcbCuentaContable 
            Height          =   375
            Left            =   240
            TabIndex        =   0
            Top             =   480
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
            ColWidth1       =   2400
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre"
            Height          =   195
            Left            =   2760
            TabIndex        =   14
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cod. Cuenta"
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Width           =   885
         End
         Begin VB.Label lblCuenta 
            BackColor       =   &H00C0FFFF&
            Height          =   375
            Left            =   2760
            TabIndex        =   10
            Top             =   480
            Width           =   3615
         End
      End
   End
End
Attribute VB_Name = "frmB_Desintegridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Inicio()
    fcbCuentaContable.SetData gobjMain.EmpresaActual.ListaCTCuentaParaCombo(False, 0)
    Me.Show vbModal
End Sub

Private Sub Limpiar()
    lblCuenta.Caption = ""
    grdCat_Trans.Rows = grdCat_Trans.FixedRows
    grdRegistros.Rows = grdRegistros.FixedRows
End Sub

Private Sub Buscar()
    Dim ct As CtCuenta, cod_ct As String
    On Error GoTo ErrTrap
    
    Limpiar
    Set ct = gobjMain.EmpresaActual.RecuperaCTCuenta(fcbCuentaContable.Text)
    If Not (ct Is Nothing) Then
        lblCuenta.Caption = ct.NombreCuenta
        cod_ct = ct.codcuenta
    Else
        lblCuenta.Caption = "No existe la cuenta contable"
        GoTo Salir
    End If
    
    If optCatalogo.value Then
        Set grdCat_Trans.DataSource = gobjMain.EmpresaActual.ListaDesintegridad(cod_ct, True, True)
    Else
        Set grdCat_Trans.DataSource = gobjMain.EmpresaActual.ListaDesintegridad(cod_ct, False, True)
    End If
    
Salir:
    Set ct = Nothing
    Exit Sub
    
ErrTrap:
    MsgBox Err.Description, vbExclamation + vbOKOnly, "Buscar"
    GoTo Salir
End Sub

Private Sub cmdBuscar_Click()
    MensajeStatus "Buscando Información", vbHourglass
    Buscar
    ConfigColsCatTrans
    GNPoneNumFila grdCat_Trans, False
    MensajeStatus "", vbNormal
End Sub

Private Sub ConfigColsCatTrans()
    With grdCat_Trans
        If optCatalogo.value Then
            .Cols = 3
            .FormatString = ">#|<Catálogo|>Num_Reg"
            .ColWidth(0) = 500
            .ColWidth(1) = 2000
            .ColWidth(2) = 1000
        Else
            .Cols = 4
            .FormatString = ">#|<CodTrans|<Transacción|>Num_Reg"
            .ColWidth(0) = 500
            .ColWidth(1) = 900
            .ColWidth(2) = 2000
            .ColWidth(3) = 800
        End If
    End With
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdVerReg_Click()
    grdRegistros.Rows = grdRegistros.FixedRows
    If grdCat_Trans.Rows = grdCat_Trans.Row Then
        MsgBox "Seleccione una fila de la grilla"
        Exit Sub
    End If
    MensajeStatus "Buscando Información...", vbHourglass
    VisualizarRegistros
    ConfigColsRegistros
    GNPoneNumFila grdRegistros, False
    MensajeStatus "", vbNormal
End Sub

Private Sub VisualizarRegistros()
    Dim codcuenta As String, codcat_trans As String
    On Error GoTo ErrTrap
    
    codcuenta = fcbCuentaContable.Text
    If Len(codcuenta) = 0 Then Exit Sub
    If optCatalogo.value Then
        codcat_trans = UCase$(Mid$(grdCat_Trans.TextMatrix(grdCat_Trans.Row, 1), 1, 2))
        Set grdRegistros.DataSource = gobjMain.EmpresaActual.ListaDetalleDesintegridad(codcuenta, codcat_trans, True, True)
    Else
        codcat_trans = UCase$(grdCat_Trans.TextMatrix(grdCat_Trans.Row, 1))
        Set grdRegistros.DataSource = gobjMain.EmpresaActual.ListaDetalleDesintegridad(codcuenta, codcat_trans, False, True)
    End If
    Exit Sub
    
ErrTrap:
    MsgBox Err.Description, vbExclamation + vbExclamation, "VisualizarRegistros"
    Exit Sub
End Sub

Private Sub ConfigColsRegistros()
    With grdRegistros
        .Cols = 4
        If optCatalogo.value Then
            cod_cat = UCase(Mid$(grdCat_Trans.TextMatrix(grdCat_Trans.Row, 1), 1, 2))
            Select Case cod_cat
            Case "IV"
                .FormatString = ">#|<Código|<Descripción|>IdCuentaActivo|<Cuenta Activo|>IdCuentaCosto|<Cuenta Costo|>IdCuentaVenta|<Cuenta Venta"
                .ColWidth(0) = 600
                .ColWidth(1) = 1200
                .ColWidth(2) = 2000
                .ColWidth(4) = 1200
                .ColWidth(6) = 1200
                .ColWidth(8) = 1200
                .ColHidden(3) = True
                .ColHidden(5) = True
                .ColHidden(7) = True
                CargarCuentasIV
            Case "PC"
                .FormatString = ">#|<Código|<Nombre|IdCuentaCuenta1|<Cuenta Contable1|IdCuentaCuenta2|<CuentaContable2|<Tipo"
                .ColWidth(0) = 600
                .ColWidth(1) = 1200
                .ColWidth(2) = 2000
                .ColWidth(4) = 1200
                .ColWidth(6) = 1200
                .ColWidth(7) = 1200
                .ColHidden(3) = True
                .ColHidden(5) = True
                .ColHidden(7) = True
                CargarCuentasPC
            Case "TS"
                .FormatString = ">#|<Código|<Nombre|>Cuenta Contable"
                .ColWidth(0) = 600
                .ColWidth(1) = 1200
                .ColWidth(2) = 2000
                .ColWidth(3) = 1500
            End Select
        Else
            .FormatString = ">#|<Trans.|<Fecha Trans.|>Num. Asiento"
            .ColWidth(0) = 600
            .ColWidth(1) = 1200
            .ColWidth(2) = 1200
            .ColWidth(3) = 1200
        End If
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF5
        cmdBuscar_Click
        KeyCode = 0
    Case vbKeyEscape
        Unload Me
    Case Else
        MoverCampo Me, KeyCode, Shift, True
    End Select
End Sub

Private Sub CargarCuentasIV()
    Dim i As Long, iv As IVinventario, cod_inv As String
    
    grdRegistros.ColHidden(4) = True
    grdRegistros.ColHidden(6) = True
    grdRegistros.ColHidden(8) = True
'***Pensar en un código mucho más rapido, se demora demasiado
'    With grdRegistros
'        If .Rows = .FixedRows Then Exit Sub
'        For i = .FixedRows To .Rows - 1
'            cod_inv = Trim$(.TextMatrix(i, 1))
'            Set iv = gobjMain.EmpresaActual.RecuperaIVInventario(cod_inv)
'            If Not (iv Is Nothing) Then
'                .TextMatrix(i, 3) = iv.CodCuentaActivo
'                .TextMatrix(i, 5) = iv.CodCuentaCosto
'                .TextMatrix(i, 7) = iv.CodCuentaVenta
'            End If
'            Set iv = Nothing
'        Next i
'    End With
End Sub

Private Sub CargarCuentasPC()
    Dim i As Long, pc As PCProvCli, cod_pc As String
    grdRegistros.ColHidden(4) = True
    grdRegistros.ColHidden(6) = True
'    With grdRegistros
'        If .Rows = .FixedRows Then Exit Sub
'        For i = .FixedRows To .Rows - 1
'            cod_pc = Trim$(.TextMatrix(i, 1))
'            Set pc = gobjMain.EmpresaActual.RecuperaPCProvCli(cod_pc)
'            If Not (pc Is Nothing) Then
'                .TextMatrix(i, 3) = pc.CodCuentaContable
'                .TextMatrix(i, 5) = pc.CodCuentaContable2
'            End If
'            Set pc = Nothing
'        Next i
'    End With
End Sub
