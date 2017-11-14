VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInventarioFisicoNew 
   Caption         =   "Constatación Física de Inventario"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10290
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   10290
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar tlb1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "img1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Abrir"
            Object.ToolTipText     =   "Abrir - F2"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Importar"
            Object.ToolTipText     =   "Importar Transacción"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar Items - F5"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar - F3"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NoContados"
            Object.ToolTipText     =   "Revisar Items no Contados"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   612
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   10290
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   7560
      Width           =   10290
      Begin VB.PictureBox pic2 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   2640
         ScaleHeight     =   495
         ScaleWidth      =   5895
         TabIndex        =   17
         Top             =   120
         Width           =   5895
         Begin VB.CommandButton cmdSiguiente 
            Caption         =   "&Siguiente - F9"
            Height          =   375
            Left            =   3000
            TabIndex        =   21
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "&Cancelar"
            Height          =   372
            Left            =   4440
            TabIndex        =   18
            Top             =   0
            Width           =   1212
         End
         Begin VB.CommandButton cmdProcesar 
            Caption         =   "&Procesar - F10"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdGrabar 
            Caption         =   "&Grabar - F3"
            Height          =   375
            Left            =   1560
            TabIndex        =   6
            Top             =   0
            Width           =   1215
         End
      End
   End
   Begin VB.PictureBox picEncabezado 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   1530
      Left            =   0
      ScaleHeight     =   1500
      ScaleWidth      =   10260
      TabIndex        =   8
      Top             =   420
      Visible         =   0   'False
      Width           =   10290
      Begin VB.TextBox txtDescripcion 
         Height          =   450
         Left            =   2400
         MaxLength       =   120
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         ToolTipText     =   "Descripción de la transacción"
         Top             =   960
         Width           =   6705
      End
      Begin VB.TextBox txtCotizacion 
         Height          =   336
         Left            =   960
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   330
         Left            =   960
         TabIndex        =   0
         ToolTipText     =   "Fecha de la transacción"
         Top             =   360
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy/MM/dd"
         Format          =   105578497
         CurrentDate     =   37078
         MaxDate         =   73415
         MinDate         =   29221
      End
      Begin FlexComboProy.FlexCombo fcbResp 
         Height          =   330
         Left            =   2400
         TabIndex        =   1
         ToolTipText     =   "Responsable de la transacción"
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         ColWidth2       =   1200
         ColWidth3       =   1200
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
      Begin FlexComboProy.FlexCombo fcbMoneda 
         Height          =   330
         Left            =   960
         TabIndex        =   2
         ToolTipText     =   "Responsable de la transacción"
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         ColWidth2       =   1200
         ColWidth3       =   1200
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "&Moneda  "
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Fecha Transaccion  "
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "&Descripción  "
         Height          =   195
         Left            =   2400
         TabIndex        =   11
         Top             =   720
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "C&otización  "
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   810
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "&Responsable  "
         Height          =   195
         Left            =   2400
         TabIndex        =   9
         Top             =   120
         Width           =   1050
      End
   End
   Begin TabDlg.SSTab sst1 
      Height          =   4095
      Left            =   0
      TabIndex        =   7
      Top             =   2160
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7223
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Inventario Fìsico (F6)"
      TabPicture(0)   =   "frmInventarioFisicoNew.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grd"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "IVFisico"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Ajuste de Bodega (F7)"
      TabPicture(1)   =   "frmInventarioFisicoNew.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "IVAjuste"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Baja de Bodega (F8)"
      TabPicture(2)   =   "frmInventarioFisicoNew.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "IVBaja"
      Tab(2).ControlCount=   1
      Begin SiiToolsA.IVFISICO IVFisico 
         Height          =   1695
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   2990
      End
      Begin SiiToolsA.IVAjuste IVBaja 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   20
         Top             =   480
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   1931
      End
      Begin SiiToolsA.IVAjuste IVAjuste 
         Height          =   1335
         Left            =   -74880
         TabIndex        =   19
         Top             =   480
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   2355
      End
      Begin VSFlex7LCtl.VSFlexGrid grd 
         Height          =   2055
         Left            =   4860
         TabIndex        =   23
         Top             =   1800
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
   End
   Begin MSComctlLib.ImageList img1 
      Left            =   9240
      Top             =   6960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventarioFisicoNew.frx":0054
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventarioFisicoNew.frx":0166
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventarioFisicoNew.frx":05B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventarioFisicoNew.frx":06CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInventarioFisicoNew.frx":07DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex7LCtl.VSFlexGrid grdMsg 
      Align           =   2  'Align Bottom
      Height          =   1095
      Left            =   0
      TabIndex        =   16
      Top             =   6465
      Width           =   10290
      _cx             =   18150
      _cy             =   1931
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
      FocusRect       =   2
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
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmInventarioFisicoNew.frx":10B6
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
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "frmInventarioFisicoNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mobjEmpresa As Empresa
Attribute mobjEmpresa.VB_VarHelpID = -1
Private mobjGNComp_CF As GNComprobante 'Transacciòn para el Conteo Físico
Private mobjGNComp_AJ As GNComprobante 'Transacción para el Ajuste de Bodega
Private mobjGNComp_BJ As GNComprobante 'Transacción para la Baja de Bodega
Private mBandSST As Boolean            '*** MAKOTO 07/mar/01 Para evitar que entraba en ciclo infinito de SetFocus
Private mObjCond As RepCondicion
Private mbooGrabado As Boolean
Private mBandRevisarNoContados As Boolean
Private dias As Integer
Const TRANS_CF = 1
Const TRANS_AJ = 2
Const TRANS_BJ = 3

Public Sub Inicio(tag As String)
    Me.tag = tag
    RecuperarConfigIVFisico 'Recupera información en registros de windows
    CargarDatos
    'Visualiza la pantalla
    Me.Show
    Me.ZOrder
    Me.WindowState = vbMaximized
End Sub

Private Sub CargarDatos()
    'Llena los datos de cabecera
    CargarEncabezado
    
    'Crea transacciiones
    If CrearTransacciones Then
        Habilitar True
        mbooGrabado = False
    Else
        Habilitar False
        mbooGrabado = True
        cmdSiguiente.Enabled = False
    End If
End Sub

Private Sub CargarEncabezado()
    picEncabezado.Visible = True
    dtpFecha.value = Date
    fcbResp.SetData gobjMain.EmpresaActual.ListaGNResponsable(False)
    fcbMoneda.SetData gobjMain.EmpresaActual.ListaGNMoneda
    fcbMoneda.KeyText = "USD"
    txtCotizacion.Text = "1"
    txtDescripcion.Text = "Constatación Física de Inventario"
End Sub

Private Function CrearTransacciones() As Boolean
    On Error GoTo mensaje
    CrearTransacciones = True
    'Transaccion para conteo fisico
    Set mobjGNComp_CF = mobjEmpresa.CreaGNComprobante(gConfigIVFisico.CodTrans_CF)
    Set IVFisico.GNComprobante = mobjGNComp_CF
    'Transaccion para Ajuste de Bodega
    Set mobjGNComp_AJ = mobjEmpresa.CreaGNComprobante(gConfigIVFisico.CodTrans_AJ)
    Set IVAjuste.GNComprobante = mobjGNComp_AJ
    'Transaccion para Baja de Bodega
    Set mobjGNComp_BJ = mobjEmpresa.CreaGNComprobante(gConfigIVFisico.CodTrans_BJ)
    Set IVBaja.GNComprobante = mobjGNComp_BJ
    Exit Function

mensaje:
    DispErr
    CrearTransacciones = False
End Function

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGrabar_Click()
    GrabarTransacciones
End Sub

Private Sub cmdProcesar_Click()
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
    
'    If gConfigIVFisico.BandLineaAuto = False Then
'        rt = MsgBox("Desea totalizar filas repetidas", vbYesNo + vbQuestion)
'        If rt = vbYes Then IVFisico.TotalizarItem
'    End If
    
'    If Not (mBandRevisarNoContados) Then
'        If MsgBox("Desea agregar los items con existencia y que no han sido contados fisicamente", vbYesNo) = vbYes Then
'            IVFisico.GNComprobante.FechaTrans = dtpFecha.value
'            IVFisico.CargarItemsNoContados
'            IVFisico.VisualizaDesdeObjeto
'            IVFisico.Refresh_Items
'            mBandRevisarNoContados = True
'        End If
'    End If
    If Me.tag <> "AjustesInventario" Then
        Procesar
    Else
        ProcesarAjuste
    End If
    MensajeStatus "", vbNormal
End Sub

Private Sub cmdSiguiente_Click()
    Dim rt As Integer
       
    If Not (mbooGrabado) Then
        rt = MsgBox(MSG_CANCELMOD, vbYesNo + vbQuestion)
        Select Case rt
        Case vbYes           'Graba y cierra
            If Grabar Then
                siguiente
            Else
                Exit Sub
            End If
        Case vbNo          'Cierra sin grabar
            siguiente
        End Select
    Else
        siguiente
    End If
End Sub

Private Sub siguiente()
    IVFisico.Limpiar
    IVAjuste.Limpiar
    IVBaja.Limpiar
    grdMsg.Rows = grdMsg.FixedRows
    If Not (mobjGNComp_CF Is Nothing) Then Set mobjGNComp_CF = Nothing
    If Not (mobjGNComp_AJ Is Nothing) Then Set mobjGNComp_AJ = Nothing
    If Not (mobjGNComp_BJ Is Nothing) Then Set mobjGNComp_BJ = Nothing
    CargarDatos
    dtpFecha.SetFocus
End Sub

Private Sub Form_Initialize()
    Set mobjEmpresa = gobjMain.EmpresaActual
    Set mObjCond = New RepCondicion
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ctl As Control
    Select Case KeyCode
    Case vbKeyF2
        AbrirArchivo
        KeyCode = 0
    Case vbKeyF3
        GrabarTransacciones
        KeyCode = 0
    Case vbKeyF5
        BuscarItems
        KeyCode = 0
    Case vbKeyF6
        If sst1.TabVisible(0) Then sst1.Tab = 0
        KeyCode = 0
    Case vbKeyF7
        If sst1.TabVisible(1) Then sst1.Tab = 1
        KeyCode = 0
    Case vbKeyF8
        If sst1.TabVisible(2) Then sst1.Tab = 2
        KeyCode = 0
    Case vbKeyF9
        cmdSiguiente_Click
        KeyCode = 0
    Case vbKeyF10
        cmdProcesar_Click
        KeyCode = 0
    Case vbKeyReturn
        'No tiene que hacer nada para que funcione Enter en grid.
    
        Set ctl = Me.ActiveControl
        If Not (ctl Is Nothing) Then
            'Si el enfoque está en fcbBanco, mueve a la siguiente
            If TypeName(ctl) = "FlexCombo" Or _
               TypeName(ctl) = "TextBox" Or _
               TypeName(ctl) = "CommandButton" Then
                ctl.Enabled = False
                ctl.Enabled = True
            End If
        End If
    Case vbKeyEscape
        Unload Me
        KeyCode = 0
    Case Else
        MoverCampo Me, KeyCode, Shift, True
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ImpideSonidoEnter Me, KeyAscii
End Sub

Private Sub Form_Load()
    mBandRevisarNoContados = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim rt As Integer
    
    If Not (mbooGrabado) Then
        Me.ZOrder
        rt = MsgBox(MSG_CANCELMOD, vbYesNoCancel + vbQuestion)
        Select Case rt
        Case vbYes           'Graba y cierra
            If Grabar Then
                Me.Hide
            Else
                Cancel = -1    'Si ocurre error al grabar,no cierra
            End If
        Case vbNo          'Cierra sin grabar
            Me.Hide
        Case vbCancel
            Cancel = -1      'No se cierra la ventana
        End Select
        If Cancel Then Me.Show              '*** MAKOTO 11/nov/00 Para que no pierda enfoque
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    sst1.Width = Me.ScaleWidth
    sst1.Height = Me.ScaleHeight - sst1.Top - (pic1.Height + grdMsg.Height)

    With IVFisico
        .Width = Me.ScaleWidth - 240
        .Height = (Me.ScaleHeight - sst1.Top - 600 - (pic1.Height + grdMsg.Height))
'        grd.Top = .Top + .Height
'        grd.Height = .Height
'        grd.Width = .Width
'        grd.Left = .Left
    End With

    With IVAjuste
        .Width = Me.ScaleWidth - 240
        .Height = Me.ScaleHeight - sst1.Top - 600 - (pic1.Height + grdMsg.Height)
    End With

    With IVBaja
        .Width = Me.ScaleWidth - 240
        .Height = Me.ScaleHeight - sst1.Top - 600 - (pic1.Height + grdMsg.Height)
    End With

    
    'Centra los botones
    pic2.Left = (Me.ScaleWidth - pic2.Width) / 2
End Sub

Private Sub Procesar()
    Dim ix As Long, ivk As IVKardex, dif As Currency
    Dim i As Long, signo As Integer, cant As String
    Dim iv As IVinventario, c As Currency
    Dim sql As String, rsc As Recordset
    IVFisico.CargaItemsOrdenado
        
    For i = 1 To IVFisico.GNComprobante.CountIVKardex
        c = 0
        dif = 0
        cant = 0
       
        dif = IVFisico.AjusteExistencia(i)
        If dif > 0 Then
            With IVAjuste
                ix = .GNComprobante.AddIVKardex
                Set ivk = IVFisico.GNComprobante.IVKardex(i)
                cant = dif
                .GNComprobante.IVKardex(ix).cantidad = cant
                .GNComprobante.IVKardex(ix).CodBodega = ivk.CodBodega
                .GNComprobante.IVKardex(ix).CodInventario = ivk.CodInventario
                
                'Calcula el costo
                Set iv = .GNComprobante.Empresa.RecuperaIVInventarioQuick(ivk.CodInventario)
                c = iv.CostoDouble2(.GNComprobante.FechaTrans, _
                                     cant, _
                                     .GNComprobante.TransID, _
                                     .GNComprobante.HoraTrans)
            
                'Si el costo calculado está en otra moneda, convierte en moneda de trans.
                If .GNComprobante.CodMoneda <> iv.CodMoneda Then
                    c = c * .GNComprobante.Cotizacion(iv.CodMoneda) / .GNComprobante.Cotizacion(" ")
                End If
                
                
                
                'Calcula el costo
'                Set iv = .GNComprobante.Empresa.RecuperaIVInventario(.GNComprobante.IVKardex(ix).CodInventario)
                
                sql = " select  top 1 g.fechatrans,g.horatrans "
                sql = sql & " from ivinventario ivi inner join  ivkardex ivk"
                sql = sql & " inner join gncomprobante g on g.transid=ivk.transid"
                sql = sql & " on ivk.idinventario=ivi.idinventario"
                sql = sql & " where ivi.codinventario='" & .GNComprobante.IVKardex(ix).CodInventario & "' and Cantidad>0"
                sql = sql & " and g.estado<>3 and CostoTotal<>0 "
                sql = sql & " order by g.fechatrans,g.horatrans"
                Set rsc = gobjMain.EmpresaActual.OpenRecordset(sql)
                If rsc.RecordCount = 1 Then
                    'obtiene el costo del primer ingreso
                    c = iv.CostoDouble2(rsc.Fields("fechatrans"), _
                                         .GNComprobante.IVKardex(ix).cantidad, _
                                         .GNComprobante.TransID, _
                                         DateAdd("s", 1, rsc.Fields("horatrans")))
                Else
                    c = 0
                End If
                If .GNComprobante.CodMoneda <> iv.CodMoneda Then
                    c = c * .GNComprobante.Cotizacion(iv.CodMoneda) / .GNComprobante.Cotizacion(" ")
                End If
                
                .GNComprobante.IVKardex(ix).CostoTotal = c * cant
                
                
                
                '.GNComprobante.IVKardex(ix).CostoTotal = c * cant
            End With
'        ElseIf dif < 0 Then
            With IVBaja
                ix = .GNComprobante.AddIVKardex
                Set ivk = IVFisico.GNComprobante.IVKardex(i)
                cant = dif
                .GNComprobante.IVKardex(ix).cantidad = cant * -1
                .GNComprobante.IVKardex(ix).CodBodega = ivk.CodBodega
                .GNComprobante.IVKardex(ix).CodInventario = ivk.CodInventario
                
                'Calcula el costo
'                Set iv = .GNComprobante.Empresa.RecuperaIVInventarioquick(ivk.CodInventario)
                c = iv.CostoDouble2(.GNComprobante.FechaTrans, _
                                     cant * -1, _
                                     .GNComprobante.TransID, _
                                     .GNComprobante.HoraTrans)

            
                'Si el costo calculado está en otra moneda, convierte en moneda de trans.
                If .GNComprobante.CodMoneda <> iv.CodMoneda Then
                    c = c * .GNComprobante.Cotizacion(iv.CodMoneda) / .GNComprobante.Cotizacion(" ")
                End If
                
                .GNComprobante.IVKardex(ix).CostoTotal = c * cant * -1
            End With
        Else
            'Si es cero no hace nada
        End If
        Debug.Print i
    Next i
    
    IVAjuste.VisualizaDesdeObjeto
    IVBaja.VisualizaDesdeObjeto
End Sub

Private Sub Form_Terminate()
    Set mObjCond = Nothing
End Sub

Private Sub IVAjuste_GotFocus()
    If sst1.Tab <> 1 Then
        mBandSST = True
        If sst1.TabVisible(1) Then sst1.Tab = 1
        mBandSST = False
    End If
End Sub

Private Sub IVBaja_GotFocus()
    If sst1.Tab <> 2 Then
        mBandSST = True
        If sst1.TabVisible(2) Then sst1.Tab = 2
        mBandSST = False
    End If
End Sub

Private Sub IVFisico_GotFocus()
    If sst1.Tab <> 0 Then
        mBandSST = True
        If sst1.TabVisible(0) Then sst1.Tab = 0
        mBandSST = False
    End If
End Sub

Private Sub sst1_Click(PreviousTab As Integer)
    '*** Para evitar error de ciclo infinito  'MAKOTO 06/mar/01
    If mBandSST Then Exit Sub
    
    On Error GoTo ErrTrap
    Select Case sst1.Tab
    Case 0          'Conteo Fisico
        IVFisico.Refresh
        If IVFisico.Enabled Then IVFisico.SetFocus
    Case 1          'Ajuste de Inventario
        IVAjuste.Refresh
        If IVAjuste.Enabled Then IVAjuste.SetFocus
    Case 2          'Baja de Bodega
        IVBaja.Refresh
        If IVBaja.Enabled Then IVBaja.SetFocus
    End Select
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub tlb1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Abrir"
        AbrirArchivo
    Case "Importar"
        ImportacionDatos
    Case "Buscar"
        BuscarItems
    Case "Grabar"
        Grabar
    Case "NoContados"
        RevisarItemsNoContados
    End Select
End Sub

Private Function Grabar() As Boolean
    GrabarTransacciones
    Grabar = mbooGrabado
End Function

Private Sub GrabarTransacciones()
    Dim trans_conteo As String, proceso As Integer, msg As String
    Dim archi As String
    
    On Error GoTo ErrTrap
    mbooGrabado = False
    ' verificar si estan todos los datos
    If Len(fcbMoneda.Text) = 0 Then
        MsgBox "Debe selecciona una tipo de Modena", vbInformation
        fcbMoneda.SetFocus
        Exit Sub
    End If
    If Val(txtCotizacion.Text) = 0 Then
        MsgBox "Escriba una cotizacion valida", vbInformation
        txtCotizacion.SetFocus
        Exit Sub
    End If
    
    If Len(txtDescripcion.Text) = 0 Then
        MsgBox "Debe escribir una Descripcion para estas transaciones", vbInformation
        txtDescripcion.SetFocus
        Exit Sub
    End If
    
    If Len(fcbResp.Text) = 0 Then
        MsgBox "Debe seleccionar un responsable", vbInformation
        fcbResp.SetFocus
        Exit Sub
    End If
    
    IVFisico.EliminaFilasIncompletas
    If IVFisico.GNComprobante.CountIVKardex = 0 Then
        MsgBox "No hay ningúna fila para grabar.", vbInformation
        Exit Sub
    End If
        
'    'Confirmación
'    If MsgBox("Está seguro que desea comenzar el proceso?", _
'                vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    
    grdMsg.Rows = grdMsg.FixedRows
    
    MensajeStatus "Grabando Ajuste de Inventario", vbHourglass
    'Graba los ajustes de inventario
    IVAjuste.Aceptar
    If IVAjuste.GNComprobante.CountIVKardex > 0 Then
        proceso = 2
        With IVAjuste.GNComprobante
            If Me.tag <> "AjustesInventario" Then
                If Me.tag = "AjusteAutomatico" Then
                    .HoraTrans = "00:00:01"
                    .FechaTrans = IVAjuste.GNComprobante.Empresa.GNOpcion.FechaLimiteDesde
                Else
                    .HoraTrans = Timer
                    .FechaTrans = dtpFecha.value
                End If
            Else
                .HoraTrans = "00:00:01"
                .FechaTrans = mObjCond.fecha1
            End If
            .CodResponsable = fcbResp.KeyText
            .CodMoneda = fcbMoneda.KeyText
            .Cotizacion(fcbMoneda.KeyText) = Val(txtCotizacion.Text)
            If Me.tag <> "AjustesInventario" Then
                If Me.tag = "AjusteAutomatico" Then
                    .Descripcion = .GNTrans.NombreTrans & " x " & Trim$(txtDescripcion.Text) & " Fecha desde: " & IVAjuste.GNComprobante.Empresa.GNOpcion.FechaLimiteDesde & " hasta: " & DateAdd("d", -1, DateAdd("m", 1, IVAjuste.GNComprobante.Empresa.GNOpcion.FechaLimiteDesde))
                Else
                    .Descripcion = .GNTrans.NombreTrans & " x " & Trim$(txtDescripcion.Text)
                End If
            Else
                .Descripcion = .GNTrans.NombreTrans & " x " & Trim$(txtDescripcion.Text) & " Fecha desde: " & mObjCond.fecha1 & " hasta: " & mObjCond.fecha2
            End If
            
            .GeneraAsiento
            'Verificación de datos
            .VerificaDatos
            'Verifica si está cuadrado el asiento
            'If Not VerificaAsiento(mobjGNComp) Then Exit Function
            .Grabar False, False
            grdMsg.AddItem "Grabando Ajuste de Inventario" & vbTab & "OK" & vbTab & .CodTrans & Str$(.numtrans)
        End With
    End If
    
    MensajeStatus "Grabando Baja de Inventario", vbHourglass
    'Graba las bajs de bodega
    IVBaja.Aceptar
    If IVBaja.GNComprobante.CountIVKardex > 0 Then
        proceso = 3
        With IVBaja.GNComprobante
            If Me.tag <> "AjustesInventario" Then
                If Me.tag = "AjusteAutomatico" Then
                    .HoraTrans = "23:59:59"
                    .FechaTrans = DateAdd("d", -1, DateAdd("m", 1, IVAjuste.GNComprobante.Empresa.GNOpcion.FechaLimiteDesde))
                Else
                    .HoraTrans = Timer
                    .FechaTrans = dtpFecha.value
                End If
            Else
                .HoraTrans = "23:59:59"
                .FechaTrans = mObjCond.fecha2
            End If
            '.FechaTrans = dtpFecha.value
            .CodResponsable = fcbResp.KeyText
            .CodMoneda = fcbMoneda.KeyText
            .Cotizacion(fcbMoneda.KeyText) = Val(txtCotizacion.Text)
            '.Descripcion = .GNTrans.NombreTrans & " x " & Trim$(txtDescripcion.Text)
            If Me.tag <> "AjustesInventario" Then
                If Me.tag = "AjusteAutomatico" Then
                    .Descripcion = .GNTrans.NombreTrans & " x " & Trim$(txtDescripcion.Text) & " Fecha desde: " & IVAjuste.GNComprobante.Empresa.GNOpcion.FechaLimiteDesde & " hasta: " & DateAdd("d", -1, DateAdd("m", 1, IVAjuste.GNComprobante.Empresa.GNOpcion.FechaLimiteDesde))
                Else
                    .Descripcion = .GNTrans.NombreTrans & " x " & Trim$(txtDescripcion.Text)
                End If
            Else
                .Descripcion = .GNTrans.NombreTrans & " x " & Trim$(txtDescripcion.Text) & " Fecha desde: " & mObjCond.fecha1 & " hasta: " & mObjCond.fecha2
            End If
                    
            .GeneraAsiento
            'Verificación de datos
            .VerificaDatos
            'Verifica si está cuadrado el asiento
            'If Not VerificaAsiento(mobjGNComp) Then Exit Function
            .Grabar False, False
            grdMsg.AddItem "Grabando Baja de Inventario" & vbTab & "OK" & vbTab & .CodTrans & Str$(.numtrans)
        End With
    End If
    
    MensajeStatus "Grabando Constatación Física", vbHourglass
    'Graba la transacción usada para el conteo físico
    IVFisico.Aceptar
    proceso = 1
    archi = IVFisico.Grabar
    If Len(archi) > 0 Then
        grdMsg.AddItem "Grabando Constatación Inventario" & vbTab & "OK" & vbTab & archi
    Else
        grdMsg.AddItem "Grabando Constatación Inventario" & vbTab & "CANCEL" & vbTab & "No guardo archivo de texto"
    End If
        
    Habilitar False
    
    MensajeStatus
    MsgBox "Proceso terminado con éxito", vbOKOnly + vbInformation
    mbooGrabado = True
    Exit Sub
    
ErrTrap:
    MensajeStatus
    Select Case proceso
    Case 1
        msg = "Grabando Constatación Inventario" & vbTab & "ERROR" & vbTab
    Case 2
        msg = "Grabando Ajuste de Inventario" & vbTab & "ERROR" & vbTab
    Case 3
        msg = "Grabando Baja de Inventario" & vbTab & "ERROR" & vbTab
    End Select
    grdMsg.AddItem msg & "No se pudo completar el proceso"
    DispErr
    Exit Sub
End Sub

Private Sub BuscarItems()
    If Me.tag = "InventarioFisico" Then
        Cargar_IVListado
    ElseIf Me.tag = "AjustesInventario" Then
        Cargar_IVListadoAjuste
    ElseIf Me.tag = "AjusteAutomatico" Then
        BuscarIvExistNegativa
        Cargar_IVListado
    
    End If
End Sub

Private Sub Cargar_IVListado()
   Dim sql As String, cond As String, rs As Recordset, signo As Integer, i As Integer, rsc As Recordset
   Dim c As Currency, cant As Integer
   Dim iv As IVinventario
   On Error GoTo ErrTrap
       'dias = DateDiff("d", "01/01/2010", "10/01/2010")
'   Cond = CondicionBusquedaItem
   With mObjCond
'        If .numGrupo = 0 Then Exit Sub  'Cuando presiona cancelar
         'Genera SQL
'         sql = "SELECT IVGrupo1.CodGrupo1 , IVGrupo1.Descripcion," & _
'              "IVGrupo2.CodGrupo2 , IVGrupo2.Descripcion," & _
'              "IVGrupo3.CodGrupo3 , IVGrupo3.Descripcion," & _
'              "IVGrupo4.CodGrupo4 , IVGrupo4.Descripcion," & _
'              "IVInventario.CodInventario, IVInventario.CodAlterno1, " & _
'              "IVInventario.Descripcion, IVBodega.CodBodega, IVInventario.Unidad, " & _
'              "CASE WHEN SUM(IVKardex.Cantidad)<0 THEN 0 ELSE SUM(IVKardex.Cantidad) END AS Existencia, " & _
'              "0 As CU, 0 As CT, "
'
'            sql = sql & "IVInventario.CodMoneda, IVInventario.Precio1, 0 As Util1, IVInventario.Precio2, 0 As Util2, " & _
'             "IVInventario.Precio3, 0 As Util3, IVInventario.Precio4, 0 As Util4 " & _
'          "From IVGrupo1 RIGHT JOIN " & _
'           "(IVGrupo2 RIGHT JOIN " & _
'           "(IVGrupo3 RIGHT JOIN " & _
'           "(IVGrupo4 RIGHT JOIN " & _
'               "(IVInventario INNER JOIN " & _
'                 "(IVBodega INNER JOIN " & _
'                   "(IVKardex INNER JOIN " & _
'                     "(GNtrans INNER JOIN GNComprobante " & _
'                      "ON GNtrans.Codtrans = GNCOmprobante.Codtrans) " & _
'                   "ON IVKardex.transID = GNComprobante.transID) " & _
'                "ON IVBodega.IdBodega = IVKArdex.IdBodega)" & _
'              "ON IVInventario.IdInventario = IVKardex.IdInventario) " & _
'           "ON  IVGrupo4.IdGrupo4 = IVInventario.IdGrupo4)" & _
'           " ON IVGrupo3.Idgrupo3 = IvInventario.Idgrupo3) " & _
'           " ON IVGrupo2.Idgrupo2 = IvInventario.Idgrupo2) " & _
'           " ON IVGrupo1.Idgrupo1 = IvInventario.Idgrupo1 "
'        If Len(Cond) > 0 Then Cond = Cond & " AND "
'        Cond = Cond & " (IVInventario.BandValida=" & CadenaBool(True, gobjMain.EmpresaActual.TipoDB) & ")" & _
'               " AND ((GNtrans.AfectaCantidad) = " & CadenaBool(True, gobjMain.EmpresaActual.TipoDB) & ") " & _
'                      " AND GNComprobante.Estado <> 3 AND BandServicio = " & CadenaBool(False, gobjMain.EmpresaActual.TipoDB)
'                              'Diego 08/09/2002 Condicion de items de Servicio
'         If InStr(Cond, "WHERE") = 0 Then sql = sql & " WHERE "
'         sql = sql & Cond
'         sql = sql & " GROUP BY IVGrupo1.CodGrupo1 , IVGrupo1.Descripcion," & _
'              "IVGrupo2.CodGrupo2 , IVGrupo2.Descripcion," & _
'              "IVGrupo3.CodGrupo3 , IVGrupo3.Descripcion," & _
'              "IVGrupo4.CodGrupo4 , IVGrupo4.Descripcion," & _
'              "IVInventario.CodInventario, IVInventario.CodAlterno1, IVInventario.Descripcion, " & _
'              "IVBodega.CodBodega, IVInventario.Unidad, IVInventario.CodMoneda, IVInventario.Precio1, " & _
'              "IVInventario.Precio2, IVInventario.Precio3, IVInventario.Precio4 "
'         If .Bandera = False Then   ' Bandera incluye existencia cero
'            sql = sql & " HAVING SUM(IVKardex.Cantidad) <> 0 "
'         End If
'         sql = sql & " ORDER BY IVGrupo" & .numGrupo & ".Descripcion, IVInventario.CodInventario"

    End With
    
        sql = " SELECT"
        sql = sql & " idInventario, CodInventario, CodAlterno1,"
        sql = sql & " Descripcion, idbodega, CodBodega,  "
        sql = sql & "0 as Mayor, "
        For i = 0 To dias
            sql = sql & "exist" & i & ", "
        Next i
        
'        For i = 0 To dias
'            sql = sql & "exist" & i & "+ "
'        Next i

        sql = Mid(sql, 1, Len(sql) - 2)
  '      sql = sql & " as existt "
        sql = sql & " from t1"
        sql = sql & " where ("
        For i = 0 To dias
            sql = sql & "exist" & i & "+ "
        Next i
        sql = Mid(sql, 1, Len(sql) - 2)
        sql = sql & " ) <>0"
    
    
    MensajeStatus MSG_PREPARA, vbArrowHourglass
    Me.Refresh
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    IVFisico.GNComprobante.BorrarIVKardex
    IVFisico.VisualizaDesdeObjeto
    
    If Not (rs.BOF And rs.EOF) Then
        'Pasa al objeto los items seleccionados
        Dim ix As Long
        With IVFisico.GNComprobante
            signo = IIf(.GNTrans.IVTipoTrans = "E", -1, 1) '-1 si es egreso
            rs.MoveFirst
            For ix = 1 To rs.RecordCount
            'Do Until rs.EOF
                ix = .AddIVKardex
                .IVKardex(ix).CodBodega = grd.TextMatrix(ix, 6)
                .IVKardex(ix).CodInventario = grd.TextMatrix(ix, 2)
                .IVKardex(ix).cantidad = grd.ValueMatrix(ix, 7) * signo
                
                'Calcula el costo
                Set iv = .Empresa.RecuperaIVInventarioQuick(.IVKardex(ix).CodInventario)
                
                sql = " select  top 1 g.fechatrans,g.horatrans "
                sql = sql & " from ivinventario ivi inner join  ivkardex ivk"
                sql = sql & " inner join gncomprobante g on g.transid=ivk.transid"
                sql = sql & " on ivk.idinventario=ivi.idinventario"
                sql = sql & " where ivi.codinventario='" & .IVKardex(ix).CodInventario & "' and Cantidad>0"
                sql = sql & " and g.estado<>3 and CostoTotal<>0 "
                sql = sql & " order by g.fechatrans,g.horatrans"
                Set rsc = gobjMain.EmpresaActual.OpenRecordset(sql)
                If rsc.RecordCount = 1 Then
                    'obtiene el costo del primer ingreso
                    c = iv.CostoDouble2(rsc.Fields("fechatrans"), _
                                         .IVKardex(ix).cantidad, _
                                         .TransID, _
                                         DateAdd("s", 1, rsc.Fields("horatrans")))
                Else
                    c = 0
                End If
                If .CodMoneda <> iv.CodMoneda Then
                    c = c * .Cotizacion(iv.CodMoneda) / .Cotizacion(" ")
                End If
                
                .IVKardex(ix).CostoTotal = c * cant
                
                
                rs.MoveNext
            'Loop
            Next ix
        End With
    Else
        IVFisico.GNComprobante.BorrarIVKardex
    End If
    IVFisico.VisualizaDesdeObjeto
    IVFisico.Refresh_Items
    MensajeStatus
    Set rs = Nothing
    Exit Sub
    
ErrTrap:
    MensajeStatus
    DispErr
    Exit Sub
End Sub

Private Function CondicionBusquedaItem() As String
    
    Static CodAlt As String, CodBodega As String, Desc As String
    Dim cond As String, Bandfirst As Boolean, comodin As String
    
#If DAOLIB Then
    comodin = "*"       'DAO
#Else
    comodin = "%"       'ADO
#End If
    

    If Not frmB_IV.Inicio(CodAlt, Desc, CodBodega, Me.tag, mObjCond) Then
        CondicionBusquedaItem = ""
        Exit Function
    End If
    Bandfirst = True
    With mObjCond
        If Len(.Item1) > 0 Then
            cond = cond & " (codInventario LIKE '" & .Item1 & comodin & "')"
            Bandfirst = False
        End If
        If Len(CodAlt) > 0 Then
            If Bandfirst = False Then cond = cond & " AND "
            cond = cond & " (codAlterno1 LIKE '" & CodAlt & comodin & "')"
            Bandfirst = False
        End If
        If Len(Desc) > 0 Then
            If Bandfirst = False Then cond = cond & " AND "
            cond = cond & " (IVInventario.Descripcion LIKE '" & Desc & comodin & "')"
            Bandfirst = False
        End If
        If Len(CodBodega) > 0 Then
            If Bandfirst = False Then cond = cond & " AND "
            cond = cond & " (IVBodega.CodBodega ='" & CodBodega & "')"
            Bandfirst = False
        End If
        
        If Not .Bandera2 Then   'esta activado el filtro avanzaado de grupos
           If (Len(.Grupo1) > 0) Or (Len(.Grupo2) > 0) Then
                If Bandfirst = False Then cond = cond & " AND "
                cond = cond & " (IVGrupo" & .numGrupo & ".CodGrupo" & _
                       CStr(.numGrupo) & " BETWEEN '" & .Grupo1 & "' AND '" & .Grupo2 & "')"
                Bandfirst = False
            End If
        Else
            'Condiciones de busqueda de grupos segun filtro avanzado
            If Len(.CodGrupo1) > 0 Then
                If Bandfirst = False Then cond = cond & " AND "
                cond = cond & " (IVGrupo1.CodGrupo1 = '" & .CodGrupo1 & "')"
                Bandfirst = False
            End If
            
            If Len(.CodGrupo2) > 0 Then
                If Bandfirst = False Then cond = cond & " AND "
                cond = cond & " (IVGrupo2.CodGrupo2 = '" & .CodGrupo2 & "')"
                Bandfirst = False
            End If
            
            If Len(.CodGrupo3) > 0 Then
                If Bandfirst = False Then cond = cond & " AND "
                cond = cond & " (IVGrupo3.CodGrupo3 = '" & .CodGrupo3 & "')"
                Bandfirst = False
            End If
            
            If Len(.CodGrupo4) > 0 Then
                If Bandfirst = False Then cond = cond & " AND "
                cond = cond & " (IVGrupo4.CodGrupo4 = '" & .CodGrupo4 & "')"
                Bandfirst = False
            End If
            
            If Len(.CodGrupo5) > 0 Then
                If Bandfirst = False Then cond = cond & " AND "
                cond = cond & " (IVGrupo5.CodGrupo5 = '" & .CodGrupo5 & "')"
                Bandfirst = False
            End If
        End If
        
       If Me.tag = "ExisMin" Then
            If .Bandera = False Then
                If Not Bandfirst Then cond = cond & " AND "
                cond = cond & " (IVExist.Exist>0) "
            End If
        End If
        If Me.tag = "Exis" Then
            'If .Bandera = False Then
            If Not Bandfirst Then cond = cond & " AND "
            cond = cond & " GNComprobante.FechaTrans <= " & FechaYMD(mObjCond.Fcorte, gobjMain.EmpresaActual.TipoDB) & " "
            'End If
        End If
    End With
    If Bandfirst = False Then cond = " WHERE " & cond
    CondicionBusquedaItem = cond
End Function

Private Sub ImportacionDatos()
    Dim Incremental As Boolean, TransIDs As String
    On Error GoTo ErrTrap
    
    If frmImportacionDatos.Inicio(mobjGNComp_CF, Incremental, TransIDs) Then
        MensajeStatus MSG_PREPARA, vbHourglass
        'Importa y visualiza los datos
        mobjGNComp_CF.ImportaDatos2 TransIDs, Incremental         '*** MAKOTO 15/dic/00
        IVFisico.VisualizaDesdeObjeto
        IVFisico.Refresh_Items
    End If
    MensajeStatus
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    Exit Sub
End Sub

Private Sub AbrirArchivo()
    Dim i As Long, filtro As String
    On Error GoTo ErrTrap
    filtro = Trim$(IVFisico.GNComprobante.CodTrans) & "*.txt"
    With frmMain.dlg1
        .CancelError = True
        .Filter = "Predeterminados (" & filtro & ")|" & filtro & "|Texto (Separado por coma)|*.txt|Todos los archivos|*.*"
        .flags = cdlOFNFileMustExist
        If Len(.filename) = 0 Then          'Solo por primera vez, ubica a la carpeta de la aplicación
            .filename = App.Path & "\" & filtro
        End If
        
        .ShowOpen
        
        Select Case UCase$(Right$(frmMain.dlg1.filename, 4))
        Case ".TXT"
'            ReformartearColumnas
            VisualizarTexto frmMain.dlg1.filename
        Case ".XLS"
            MsgBox "No disponible"
'            VisualizarExcel dlg1.FileName
        Case Else
        End Select
    End With
    Exit Sub
ErrTrap:
    If Err.Number <> 32755 Then DispErr
    Exit Sub
End Sub

Private Sub VisualizarTexto(ByVal archi As String)
    Dim f As Integer, s As String, v As Variant, cont As Integer, i As Integer
    Dim COL_COD As Long, COL_BD As Long, COL_CANT As Long, COL_NOTA As Long, ix As Long
    On Error GoTo ErrTrap
    
    MensajeStatus "Está leyendo el archivo " & archi & " ...", vbHourglass
    f = FreeFile                'Obtiene número disponible de archivo
    cont = 0
    COL_BD = -1
    COL_COD = -1
    COL_CANT = -1
    COL_NOTA = -1
    'Abre el archivo para lectura
    Open archi For Input As #f
        IVFisico.GNComprobante.BorrarIVKardex
        IVFisico.VisualizaDesdeObjeto
        Do Until EOF(f)
            Line Input #f, s
            v = Split(s, vbTab)
            If cont = 0 Then
                For i = LBound(v, 1) To UBound(v, 1)
                    If InStr(1, UCase(v(i)), "BODEGA") Then COL_BD = i
                    If InStr(1, UCase(v(i)), "CODIGO") Then COL_COD = i
                    If InStr(1, UCase(v(i)), "CANT") Then COL_CANT = i
                    If InStr(1, UCase(v(i)), "NOTA") Then COL_NOTA = i
                Next i
                If (COL_BD = -1) Or (COL_COD = -1) Or (COL_CANT = -1) Or (COL_NOTA = -1) Then
                    MsgBox "No se puede leer el archivo seleccionado." & vbCrLf & _
                           "Nombres de Columnas no reconocidos", vbOKOnly + vbInformation
                    GoTo seguir
                End If
            Else
                ix = IVFisico.GNComprobante.AddIVKardex
                IVFisico.GNComprobante.IVKardex(ix).CodBodega = QuitaComillas(v(COL_BD))
                IVFisico.GNComprobante.IVKardex(ix).CodInventario = QuitaComillas(v(COL_COD))
                IVFisico.GNComprobante.IVKardex(ix).cantidad = CCur("0" & QuitaComillas(v(COL_CANT))) * -1
                IVFisico.GNComprobante.IVKardex(ix).Nota = QuitaComillas(v(COL_NOTA))
            End If
            cont = 1
        Loop
        IVFisico.VisualizaDesdeObjeto
        IVFisico.Refresh_Items
seguir:
    Close #f
    
    MensajeStatus
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    Close       'Cierra todo
    Exit Sub
End Sub

Private Function QuitaComillas(ByVal Cadena As String) As String
    Dim s As String
    s = Cadena
    If Mid$(s, 1, 1) = """" Then s = Mid$(s, 2)
    If Right$(s, 1) = """" Then s = Mid$(s, 1, Len(s) - 1)
    QuitaComillas = s
End Function

Private Sub Habilitar(ByVal band As Boolean)
    cmdProcesar.Enabled = band
    cmdGrabar.Enabled = band
    tlb1.Buttons(1).Enabled = band
    tlb1.Buttons(2).Enabled = band
    tlb1.Buttons(3).Enabled = band
    tlb1.Buttons(4).Enabled = band
    
    dtpFecha.Enabled = band
    fcbResp.Enabled = band
    fcbMoneda.Enabled = band
    txtCotizacion.Enabled = band
    txtDescripcion.Enabled = band
    
    IVFisico.Enabled = band
    IVAjuste.Enabled = band
    IVBaja.Enabled = band
End Sub

Private Sub RevisarItemsNoContados()
    If IVFisico.GNComprobante.CountIVKardex = 0 Then
        MsgBox "No hay filas para comparar", vbOKOnly + vbInformation
        Exit Sub
    End If
    
    'Para controlar que el usuario se equivoque y agrege nuevamente los items no contados
    If mBandRevisarNoContados Then
        If MsgBox("Este proceso ya fue realizado. Desea ejecutarlo nuevamente", vbYesNo) = vbNo Then Exit Sub
    End If
    MensajeStatus MSG_PREPARA, vbArrowHourglass
    IVFisico.GNComprobante.FechaTrans = dtpFecha.value
    IVFisico.CargarItemsNoContados
    IVFisico.VisualizaDesdeObjeto
    IVFisico.Refresh_Items
    mBandRevisarNoContados = True
    MensajeStatus "", vbNormal
End Sub

Private Sub Cargar_IVListadoAjuste()
   Dim sql As String, cond As String, rs As Recordset, signo As Integer
   Dim BandVolumen As Boolean
   BandVolumen = False
   On Error GoTo ErrTrap
   
   cond = CondicionBusquedaItemAjuste
With mObjCond
        If .numGrupo = 0 Then Exit Sub  'Cuando presiona cancelar
         'Genera SQL
         sql = "SELECT IVGrupo1.CodGrupo1 , IVGrupo1.Descripcion," & _
              "IVGrupo2.CodGrupo2 , IVGrupo2.Descripcion," & _
              "IVGrupo3.CodGrupo3 , IVGrupo3.Descripcion," & _
              "IVGrupo4.CodGrupo4 , IVGrupo4.Descripcion," & _
              "IVInventario.CodInventario, IVInventario.CodAlterno1, " & _
              "IVInventario.Descripcion, IVBodega.CodBodega, IVInventario.Unidad, SUM(IVKardex.Cantidad) AS Existencia, " & _
              "0 As CU, 0 As CT, "
              If BandVolumen Then sql = sql & " 0  as Volumen, "
            sql = sql & "IVInventario.CodMoneda, IVInventario.Precio1, 0 As Util1, IVInventario.Precio2, 0 As Util2, " & _
             "IVInventario.Precio3, 0 As Util3, IVInventario.Precio4, 0 As Util4, IVInventario.Observacion " & _
         "From IVGrupo1 RIGHT JOIN " & _
           "(IVGrupo2 RIGHT JOIN " & _
           "(IVGrupo3 RIGHT JOIN " & _
           "(IVGrupo4 RIGHT JOIN " & _
               "(IVInventario INNER JOIN " & _
                 "(IVBodega INNER JOIN " & _
                   "(IVKardex INNER JOIN " & _
                     "(GNtrans INNER JOIN GNComprobante " & _
                      "ON GNtrans.Codtrans = GNCOmprobante.Codtrans) " & _
                   "ON IVKardex.transID = GNComprobante.transID) " & _
                "ON IVBodega.IdBodega = IVKArdex.IdBodega)" & _
              "ON IVInventario.IdInventario = IVKardex.IdInventario) " & _
           "ON  IVGrupo4.IdGrupo4 = IVInventario.IdGrupo4)" & _
           " ON IVGrupo3.Idgrupo3 = IvInventario.Idgrupo3) " & _
           " ON IVGrupo2.Idgrupo2 = IvInventario.Idgrupo2) " & _
           " ON IVGrupo1.Idgrupo1 = IvInventario.Idgrupo1 "
        If Len(cond) > 0 Then cond = cond & " AND "
        cond = cond & " (IVInventario.BandValida=" & CadenaBool(True, gobjMain.EmpresaActual.TipoDB) & ")" & _
               " AND ((GNtrans.AfectaCantidad) = " & CadenaBool(True, gobjMain.EmpresaActual.TipoDB) & ") " & _
                      " AND GNComprobante.Estado <> 3 AND BandServicio = " & CadenaBool(False, gobjMain.EmpresaActual.TipoDB)
                              'Diego 08/09/2002 Condicion de items de Servicio
         If InStr(cond, "WHERE") = 0 Then sql = sql & " WHERE "
         sql = sql & cond
         sql = sql & " GROUP BY IVGrupo1.CodGrupo1 , IVGrupo1.Descripcion," & _
              "IVGrupo2.CodGrupo2 , IVGrupo2.Descripcion," & _
              "IVGrupo3.CodGrupo3 , IVGrupo3.Descripcion," & _
              "IVGrupo4.CodGrupo4 , IVGrupo4.Descripcion," & _
              "IVInventario.CodInventario, IVInventario.CodAlterno1, IVInventario.Descripcion, " & _
              "IVBodega.CodBodega, IVInventario.Unidad, IVInventario.CodMoneda, IVInventario.Precio1, " & _
              "IVInventario.Precio2, IVInventario.Precio3, IVInventario.Precio4, IvInventario.Observacion "
        'If .Bandera = False Then   ' Bandera incluye existencia cero
            sql = sql & " HAVING SUM(IVKardex.Cantidad) < 0 "
         'End If
         sql = sql & " ORDER BY IVGrupo" & .numGrupo & ".Descripcion, IVInventario.CodInventario"
    End With
    MensajeStatus MSG_PREPARA, vbArrowHourglass
    Me.Refresh
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    IVFisico.GNComprobante.BorrarIVKardex
    IVFisico.VisualizaDesdeObjeto
    
    If Not (rs.BOF And rs.EOF) Then
        'Pasa al objeto los items seleccionados
        Dim ix As Long
        With IVFisico.GNComprobante
            signo = IIf(.GNTrans.IVTipoTrans = "E", -1, 1) '-1 si es egreso
            Do Until rs.EOF
                ix = .AddIVKardex
                .IVKardex(ix).CodBodega = rs!CodBodega
                .IVKardex(ix).CodInventario = rs!CodInventario
                .IVKardex(ix).cantidad = rs!Existencia * signo
                rs.MoveNext
            Loop
        End With
    Else
        IVFisico.GNComprobante.BorrarIVKardex
    End If
    If Me.tag <> "AjustesInventario" Then
        IVFisico.VisualizaDesdeObjeto
    Else
        IVFisico.VisualizaDesdeObjetoAjustes
    End If
    IVFisico.Refresh_Items
    MensajeStatus
    Set rs = Nothing
    Exit Sub
    
ErrTrap:
    MensajeStatus
    DispErr
    Exit Sub
End Sub

Private Function CondicionBusquedaItemAjuste() As String
    
    Static CodAlt As String, CodBodega As String, Desc As String
    Dim cond As String, Bandfirst As Boolean, comodin As String
    
#If DAOLIB Then
    comodin = "*"       'DAO
#Else
    comodin = "%"       'ADO
#End If
    

    If Not frmB_IVAjustes.Inicio(CodAlt, Desc, CodBodega, Me.tag, mObjCond) Then
        CondicionBusquedaItemAjuste = ""
        Exit Function
    End If
    Bandfirst = True
    With mObjCond
        If Len(.Item1) > 0 Then
            cond = cond & " (codInventario LIKE '" & .Item1 & comodin & "')"
            Bandfirst = False
        End If
        If Len(CodAlt) > 0 Then
            If Bandfirst = False Then cond = cond & " AND "
            cond = cond & " (codAlterno1 LIKE '" & CodAlt & comodin & "')"
            Bandfirst = False
        End If
        If Len(Desc) > 0 Then
            If Bandfirst = False Then cond = cond & " AND "
            cond = cond & " (IVInventario.Descripcion LIKE '" & Desc & comodin & "')"
            Bandfirst = False
        End If
        If Len(CodBodega) > 0 Then
            If Bandfirst = False Then cond = cond & " AND "
            cond = cond & " (IVBodega.CodBodega ='" & CodBodega & "')"
            Bandfirst = False
        End If
        
        If Not .Bandera2 Then   'esta activado el filtro avanzaado de grupos
           If (Len(.Grupo1) > 0) Or (Len(.Grupo2) > 0) Then
                If Bandfirst = False Then cond = cond & " AND "
                cond = cond & " (IVGrupo" & .numGrupo & ".CodGrupo" & _
                       CStr(.numGrupo) & " BETWEEN '" & .Grupo1 & "' AND '" & .Grupo2 & "')"
                Bandfirst = False
            End If
        Else
            'Condiciones de busqueda de grupos segun filtro avanzado
            If Len(.CodGrupo1) > 0 Then
                If Bandfirst = False Then cond = cond & " AND "
                cond = cond & " (IVGrupo1.CodGrupo1 = '" & .CodGrupo1 & "')"
                Bandfirst = False
            End If
            
            If Len(.CodGrupo2) > 0 Then
                If Bandfirst = False Then cond = cond & " AND "
                cond = cond & " (IVGrupo2.CodGrupo2 = '" & .CodGrupo2 & "')"
                Bandfirst = False
            End If
            
            If Len(.CodGrupo3) > 0 Then
                If Bandfirst = False Then cond = cond & " AND "
                cond = cond & " (IVGrupo3.CodGrupo3 = '" & .CodGrupo3 & "')"
                Bandfirst = False
            End If
            
            If Len(.CodGrupo4) > 0 Then
                If Bandfirst = False Then cond = cond & " AND "
                cond = cond & " (IVGrupo4.CodGrupo4 = '" & .CodGrupo4 & "')"
                Bandfirst = False
            End If
            
            If Len(.CodGrupo5) > 0 Then
                If Bandfirst = False Then cond = cond & " AND "
                cond = cond & " (IVGrupo5.CodGrupo5 = '" & .CodGrupo5 & "')"
                Bandfirst = False
            End If
        End If
        
       If Me.tag = "ExisMin" Then
            If .Bandera = False Then
                If Not Bandfirst Then cond = cond & " AND "
                cond = cond & " (IVExist.Exist>0) "
            End If
        End If
        If Me.tag = "Exis" Then
            'If .Bandera = False Then
            If Not Bandfirst Then cond = cond & " AND "
            cond = cond & " GNComprobante.FechaTrans <= " & FechaYMD(mObjCond.Fcorte, gobjMain.EmpresaActual.TipoDB) & " "
            'End If
        End If
        If Me.tag = "AjustesInventario" Then
            'If .Bandera = False Then
            If Not Bandfirst Then cond = cond & " AND "
            'Cond = Cond & " GNComprobante.FechaTrans between " & FechaYMD(mObjCond.Fecha1, gobjMain.EmpresaActual.TipoDB) & " and " & FechaYMD(mObjCond.Fecha2, gobjMain.EmpresaActual.TipoDB) & " "
            cond = cond & " GNComprobante.FechaTrans <= " & FechaYMD(mObjCond.fecha2, gobjMain.EmpresaActual.TipoDB) & " "
        End If
    End With
    If Bandfirst = False Then cond = " WHERE " & cond
    CondicionBusquedaItemAjuste = cond
End Function



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
                .GNComprobante.IVKardex(ix).cantidad = cant
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
                .GNComprobante.IVKardex(ix).cantidad = cant
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


Private Sub BuscarIvExistNegativa()
Static coditem As String, CodAlt As String, _
           Desc As String, _
           codg As String, Numg As Integer, bandIVA As Boolean, bandFraccion As Boolean
    Dim codg1 As String, codg2 As String, codg3 As String, codg4 As String, codg5 As String
    Dim CodBodega As String, sum As String
    Static CodTrans As String, desde As Long, hasta As Long
    Dim sql As String, cond As String, rs As Recordset, comodin As String
    Dim fechaIni As Date, FechaHasta As Date
    Dim i As Integer, NumReg As Long, j As Integer
    Dim diainicio  As Date
    On Error GoTo ErrTrap
    
    #If DAOLIB Then
        comodin = "*"
    #Else
        comodin = "%"
    #End If
'    comodin = "%"
    'Abre la pantalla de búsqueda
'    If Not frmIVBusqueda.Inicio( _
'                coditem, _
'                CodAlt, _
'                Desc, _
'                codg1, codg2, codg3, codg4, codg5, _
'                Numg, _
'                bandIVA, _
'                Me.tag, CodBodega) Then
'      'if not frmivbusqueda.InicioTrans (
'        'Si fue cancelada la busqueda, sale no mas
'        grd.SetFocus
'        Exit Sub
'    End If
    
    'Cambia la forma de cursor
    MensajeStatus MSG_PREPARA, vbHourglass
    
    diainicio = gobjMain.EmpresaActual.GNOpcion.FechaLimiteDesde
    dias = DateDiff("d", diainicio, DateAdd("d", -1, DateAdd("m", 1, gobjMain.EmpresaActual.GNOpcion.FechaLimiteDesde))) + 1
    dias = 3
    For i = 0 To dias
        sql = " SELECT"
        sql = sql & " IVInventario.idInventario,"
        sql = sql & " IVBodega.idBodega, SUM(IVKardex.Cantidad)*-1 AS Existencia"
        sql = sql & " into tmp" & i
        sql = sql & " From (IVInventario"
        sql = sql & " INNER JOIN (IVBodega"
        sql = sql & " INNER JOIN (IVKardex"
        sql = sql & " INNER JOIN (GNtrans"
        sql = sql & " INNER JOIN GNComprobante"
        sql = sql & " ON GNtrans.Codtrans = GNCOmprobante.Codtrans)"
        sql = sql & " ON IVKardex.transID = GNComprobante.transID)"
        sql = sql & " ON IVBodega.IdBodega = IVKArdex.IdBodega)"
        sql = sql & " ON IVInventario.IdInventario = IVKardex.IdInventario)"
        sql = sql & " WHERE  GNComprobante.FechaTrans <= '" & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaLimiteDesde) & "'"
        sql = sql & " AND  ((GNtrans.AfectaCantidad) = 1)"
        sql = sql & " AND GNComprobante.Estado <> 3"
        'sql = sql & " AND BandServicio = 0"
        sql = sql & " GROUP BY IVInventario.IdInventario,  IVBodega.idBodega"
        sql = sql & " Having Sum(IVKardex.cantidad) < 0"
  
         VerificaExistenciaTablaTemp "tmp" & i
         gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
         Debug.Print i
         MensajeStatus "Está procesando día " & DateAdd("d", i, diainicio)
    Next i
        sum = ""
        sql = " SELECT"
        sql = sql & " IVInventario.idInventario, IVInventario.CodInventario, IVInventario.CodAlterno1,"
        sql = sql & " IVInventario.Descripcion, IVBodega.idBodega, IVBodega.CodBodega, 0 AS Existencia, "
        For i = 0 To dias
            sql = sql & " isnull(tmp" & i & ". existencia,0) as exist" & i & ", "
        Next i
       
       sql = Mid(sql, 1, Len(sql) - 2)
        sql = sql & " into t" & 1
        sql = sql & " From IVexist"
        sql = sql & " INNER JOIN IVBodega"
        sql = sql & " ON IVBodega.IdBodega = IVexist.IdBodega"
        sql = sql & " INNER JOIN IVInventario"
        sql = sql & " ON IVInventario.IdInventario = IVexist.IdInventario"
      For i = 0 To dias
            sql = sql & " left JOIN tmp" & i
            sql = sql & " ON IVInventario.IdInventario = tmp" & i & ".idInventario "
            sql = sql & " and IVexist.idbodega = tmp" & i & ".idbodega "
    Next i
     sql = sql & " where BandServicio = 0"
    
    
    VerificaExistenciaTablaTemporal 1
    gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
    
    
        sql = " SELECT"
        sql = sql & " idInventario, CodInventario, CodAlterno1,"
        sql = sql & " Descripcion, idbodega, CodBodega,  "
        sql = sql & "0 as Mayor, "
        For i = 0 To dias
            sql = sql & "isnull(exist" & i & ",0),"
        Next i
        
'        For i = 0 To dias
'            sql = sql & "exist" & i & "+ "
'        Next i

        sql = Mid(sql, 1, Len(sql) - 1)
  '      sql = sql & " as existt "
        sql = sql & " from t1"
        sql = sql & " where ("
        For i = 0 To dias
            sql = sql & "isnull(exist" & i & ",0) + "
        Next i
        sql = Mid(sql, 1, Len(sql) - 2)
        sql = sql & " ) <>0"
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    With grd
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If Not rs.EOF Then .LoadArray MiGetRows(rs)
        
        .Redraw = flexRDBuffered
'        .SetFocus
        For i = 1 To grd.Rows - 1
            For j = 8 To grd.Cols - 1
                If j = 8 Then
                    grd.TextMatrix(i, 7) = grd.TextMatrix(i, j)
                Else
                    If grd.ValueMatrix(i, j) >= grd.ValueMatrix(i, 7) Then
                        grd.TextMatrix(i, 7) = grd.TextMatrix(i, j)
                    End If
                End If
            Next j
            .TextMatrix(i, 0) = i
        Next i
        
    End With
    
    
        For i = 0 To dias
             VerificaExistenciaTablaTemp "tmp" & i
        Next i

    
    For j = 8 To grd.Cols - 1
'        grd.ColHidden(j) = True
    Next j
    MensajeStatus
    Exit Sub
ErrTrap:
    grd.Redraw = flexRDBuffered
    MensajeStatus
    DispErr
'    grd.SetFocus
    Exit Sub
End Sub




