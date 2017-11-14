VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{50067EB3-D6AF-11D3-8297-000021C5085D}#1.0#0"; "NTextBox.ocx"
Begin VB.Form frmGenerarCompAutomatico 
   Caption         =   "Generación de Pagos/Ingreso"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8520
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   8520
   WindowState     =   2  'Maximized
   Begin VB.Frame FraConFigEgreso 
      Caption         =   "Datos para Egresos"
      Height          =   855
      Left            =   60
      TabIndex        =   32
      Top             =   1620
      Visible         =   0   'False
      Width           =   8115
      Begin VB.CommandButton cmdExplorarCH 
         Caption         =   "..."
         Height          =   310
         Left            =   7620
         TabIndex        =   11
         Top             =   480
         Width           =   372
      End
      Begin VB.CommandButton cmdExplorar 
         Caption         =   "..."
         Height          =   310
         Left            =   7620
         TabIndex        =   9
         Top             =   180
         Width           =   372
      End
      Begin VB.TextBox txtCheque 
         Height          =   315
         Left            =   2220
         TabIndex        =   10
         Top             =   480
         Width           =   5415
      End
      Begin VB.TextBox txtEgreso 
         Height          =   315
         Left            =   2220
         TabIndex        =   8
         Top             =   180
         Width           =   5415
      End
      Begin MSComDlg.CommonDialog dlg1 
         Left            =   6180
         Top             =   120
         _ExtentX        =   688
         _ExtentY        =   688
         _Version        =   393216
         CancelError     =   -1  'True
         DefaultExt      =   "mdb"
         DialogTitle     =   "Destino de exportación"
      End
      Begin VB.Label Label7 
         Caption         =   "Lib. Impresion Cheque"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   540
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Lib. Impresion Egreso"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1875
      End
   End
   Begin VB.Frame fraCodTrans 
      Caption         =   "Cod.&Trans.Ingreso"
      Height          =   1515
      Left            =   8220
      TabIndex        =   25
      Top             =   60
      Width           =   3255
      Begin VB.PictureBox PicForma 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   30
         ScaleHeight     =   375
         ScaleWidth      =   2415
         TabIndex        =   35
         Top             =   600
         Visible         =   0   'False
         Width           =   2415
         Begin FlexComboProy.FlexCombo fcbFormaCobro 
            Height          =   345
            Left            =   570
            TabIndex        =   36
            Top             =   0
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   609
            ColWidth0       =   500
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
         Begin VB.Label Label3 
            Caption         =   "F.Cobro"
            Height          =   255
            Left            =   60
            TabIndex        =   37
            Top             =   60
            Width           =   435
         End
      End
      Begin FlexComboProy.FlexCombo fcbTrans 
         Height          =   345
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   609
         ColWidth0       =   500
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
      Begin FlexComboProy.FlexCombo fcbBanco 
         Height          =   345
         Left            =   600
         TabIndex        =   5
         Top             =   660
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   609
         ColWidth0       =   300
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
      Begin VB.Label lblnumche 
         Caption         =   "#. Cheque"
         Height          =   195
         Left            =   2400
         TabIndex        =   31
         Top             =   780
         Width           =   795
      End
      Begin VB.Label lblNumCheque 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   2400
         TabIndex        =   30
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblsaldoBanco 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   600
         TabIndex        =   29
         Top             =   1080
         Width           =   1755
      End
      Begin VB.Label lblSal 
         Caption         =   "Saldo"
         Height          =   255
         Left            =   60
         TabIndex        =   28
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label lblBanco 
         Caption         =   "Banco"
         Height          =   255
         Left            =   60
         TabIndex        =   27
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Trans."
         Height          =   255
         Left            =   60
         TabIndex        =   26
         Top             =   300
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar - F5"
      Height          =   372
      Left            =   780
      TabIndex        =   7
      Top             =   1140
      Width           =   1212
   End
   Begin VB.Frame fraCodTransVenta 
      Caption         =   "Cod.&Trans. Venta"
      Height          =   1515
      Left            =   3540
      TabIndex        =   22
      Top             =   60
      Width           =   2295
      Begin VB.ListBox lstTrans 
         Columns         =   2
         Height          =   1275
         IntegralHeight  =   0   'False
         Left            =   120
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   180
         Width           =   2055
      End
   End
   Begin VB.Frame FrameFC 
      Caption         =   "Cod.&Forma Cobro"
      Height          =   1515
      Left            =   5880
      TabIndex        =   21
      Top             =   60
      Width           =   2295
      Begin VB.ListBox lstForma 
         Columns         =   2
         Height          =   1275
         IntegralHeight  =   0   'False
         Left            =   120
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   180
         Width           =   2055
      End
   End
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   852
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   8520
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5565
      Width           =   8520
      Begin VB.CommandButton cmdImprimiCH 
         Caption         =   "&Imprimir Cheques"
         Enabled         =   0   'False
         Height          =   372
         Left            =   5520
         TabIndex        =   16
         Top             =   360
         Width           =   1452
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   372
         Left            =   4080
         TabIndex        =   15
         Top             =   360
         Width           =   1452
      End
      Begin VB.CommandButton cmdAsiento 
         Caption         =   "&Asiento"
         Enabled         =   0   'False
         Height          =   372
         Left            =   2880
         TabIndex        =   14
         Top             =   360
         Width           =   1212
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Proceder - F8"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   360
         Width           =   1212
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar -F3"
         Height          =   372
         Left            =   8340
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   9660
         TabIndex        =   17
         Top             =   360
         Width           =   1212
      End
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   20
         Top             =   60
         Width           =   8280
         _ExtentX        =   14605
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   2775
      Left            =   0
      TabIndex        =   12
      Top             =   2580
      Width           =   8175
      _cx             =   14420
      _cy             =   4895
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
      AllowSelection  =   -1  'True
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
   Begin VB.Frame FraNumero 
      Caption         =   "&Número (desde - hasta)"
      Height          =   1035
      Left            =   60
      TabIndex        =   38
      Top             =   60
      Visible         =   0   'False
      Width           =   3435
      Begin NTextBoxProy.NTextBox ntxDesde 
         Height          =   315
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Text            =   "0"
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
      Begin NTextBoxProy.NTextBox ntxHasta 
         Height          =   315
         Left            =   1860
         TabIndex        =   41
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Text            =   "0"
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
      Begin FlexComboProy.FlexCombo fcbTransAnulada 
         Height          =   345
         Left            =   1440
         TabIndex        =   42
         Top             =   600
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   609
         ColWidth0       =   500
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
      Begin VB.Label Label4 
         Caption         =   "Transacciones"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "~  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   1560
         TabIndex        =   39
         Top             =   300
         Width           =   315
      End
   End
   Begin VB.Frame fraFecha 
      Caption         =   "&Fecha (desde - hasta)"
      Height          =   675
      Left            =   60
      TabIndex        =   23
      Top             =   60
      Width           =   3435
      Begin MSComCtl2.DTPicker dtpFecha2 
         Height          =   330
         Left            =   1800
         TabIndex        =   1
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   106692609
         CurrentDate     =   36902
      End
      Begin MSComCtl2.DTPicker dtpFecha1 
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   106692609
         CurrentDate     =   36902
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "~  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1560
         TabIndex        =   24
         Top             =   300
         Width           =   315
      End
   End
   Begin VB.CheckBox chkAgrupaProv 
      Caption         =   "Agrupar por Proveedor"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "frmGenerarCompAutomatico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Constantes para las columnas
Private Const COL_NUMFILA = 0
Private Const COL_TID = 1
Private Const COL_FECHA = 2
Private Const COL_ID = 3
Private Const COL_CODTRANS = 4
Private Const COL_NUMTRANS = 5
Private Const COL_NUMDOCREF = 6
Private Const COL_NOMBRE = 7
Private Const COL_DESC = 8
Private Const COL_FORMA = 9
Private Const COL_VALOR = 10
Private Const COL_CODFORMA = 10
Private Const COL_AUTO = 11
Private Const COL_VALORPAGO = 12
Private Const COL_RESULTADO = 13
Private Const COL_TIDIN = 14
Private Const COL_NUMTRANSIN = 15

Private Const COL_CODTARJETA = 8
Private Const COL_CODPROVCLI = 9
Private Const COL_NOMBANCO = 10
Private Const COL_IDPAGO = 11

Private mProcesando As Boolean
Private mCancelado As Boolean
Private mVerificado As Boolean

Private WithEvents mobjGNComp As GNComprobante
Attribute mobjGNComp.VB_VarHelpID = -1
Private mobjGNCompOrigen As GNComprobante
Private mobjGNCompAux As GNComprobante
Private mColItems As Collection
Private Const MSG_NG = "Asiento incorrecto."
Private mCodMoneda As String
Private mobjSiiMain As SiiMain

Public Sub Inicio()
    Dim i As Integer
    On Error GoTo ErrTrap
    lblBanco.Visible = False
    fcbBanco.Visible = False
    fcbBanco.TabStop = False
    lblnumche.Visible = False
    lblSal.Visible = False
    lblsaldoBanco.Visible = False
    lblNumCheque.Visible = False
    cmdImprimiCH.Visible = False
    cmdImprimiCH.TabStop = False
    chkAgrupaProv.Caption = "Agrupar por F.Cobro"
    ConfigCols
    Me.Show
    Me.ZOrder
    dtpFecha1.value = Date
    dtpFecha2.value = Date
    CargaTrans
    
    CargaFormas
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub CargaTrans()
    Dim i As Long, v As Variant
    Dim s As String
    
    fcbTrans.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "IV")
    fcbTransAnulada.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "")
    'fcbFormaCobro.SetData gobjMain.EmpresaActual.ListaTSFormaCobroPago(True, True, False)
    
    lstTrans.Clear
    'v = gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "IV")
    v = gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "")
    
    For i = LBound(v, 2) To UBound(v, 2)
        lstTrans.AddItem v(0, i)        '& " " & v(1, i)
    Next i
    
    'jeaa 25/09/206
    If Me.tag = "CruceTarjetas" Then
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransParaCruceTarjetas")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransParaCruceTarjetas")
            RecuperaTrans "KeyT", lstTrans, s
        End If
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransIngresoParaCruceTarjetas")) > 0 Then
            fcbTrans.KeyText = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransIngresoParaCruceTarjetas")
        End If
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransFormaCobroParaCruceTarjetas")) > 0 Then
            fcbFormaCobro.KeyText = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransFormaCobroParaCruceTarjetas")
        End If
    ElseIf Me.tag = "CruceIVTarjetas" Then
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransParaCruceIVTarjetas")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransParaCruceIVTarjetas")
            RecuperaTrans "KeyT", lstTrans, s
        End If
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransIngresoParaCruceIVTarjetas")) > 0 Then
            fcbTrans.KeyText = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransIngresoParaCruceIVTarjetas")
        End If
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransFormaCobroParaCruceIVTarjetas")) > 0 Then
            fcbFormaCobro.KeyText = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransFormaCobroParaCruceIVTarjetas")
        End If
        
    Else
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransParaIngresoAutomatico")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransParaIngresoAutomatico")
            RecuperaTrans "KeyT", lstTrans, s
        End If
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransIngresoParaIngresoAutomatico")) > 0 Then
            fcbTrans.KeyText = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransIngresoParaIngresoAutomatico")
        End If
    End If
        
End Sub



Private Sub chkAgrupaProv_Click()
'    grd.ColSort(COL_NOMBRE) = flexSortGenericAscending
'    grd.Refresh
'    grd.subtotal flexSTSum, COL_NOMBRE - 1, 10, , vbBlue, vbWhite
End Sub

Private Sub cmdAceptar_Click()
    If Not mProcesando Then
        'Si no hay transacciones
        If grd.Rows <= grd.FixedRows Then
            MsgBox "No hay ningúna transacción para procesar."
            Exit Sub
        End If
    Select Case Me.tag
    Case "InicioAnulados"
        cmdAsiento.Visible = False
        cmdImprimir.Visible = False
        cmdImprimiCH.Visible = False
        If TransAnuladaAuto(True, False) Then
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
            mVerificado = True
            cmdAsiento.Enabled = False
            cmdImprimir.Enabled = False
        Else
            cmdAceptar.Enabled = False
            cmdAsiento.Enabled = True
        End If
    
    Case Else
        If Len(fcbTrans.KeyText) = 0 Then
            MsgBox "No hay ningúna transacción de Ingreso no se podrá procesar."
            fcbTrans.SetFocus
            Exit Sub
        End If

        If IngresoAuto(True, False) Then
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
            mVerificado = True
            cmdAsiento.Enabled = False
            cmdImprimir.Enabled = False
        Else
            cmdAceptar.Enabled = False
            cmdAsiento.Enabled = True
        End If
    End Select
End If
    
End Sub





Private Sub cmdAsiento_Click()
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna transacción para procesar."
        Exit Sub
    End If
    
    
    If RegenerarAsiento(False, True) Then
        cmdCancelar.SetFocus
        cmdAsiento.Enabled = False
        cmdImprimir.Enabled = True
        cmdImprimiCH.Enabled = True
    Else
        cmdImprimir.Enabled = False
        cmdImprimiCH.Enabled = False
    End If

End Sub

Private Sub cmdBuscar_Click()
    Dim v As Variant, obj As Object, s As String
    On Error GoTo ErrTrap
    fcbBanco.KeyText = ""
    lblsaldoBanco.Caption = ""
    lblNumCheque.Caption = ""
    With gobjMain.objCondicion
        mCodMoneda = GetSetting(APPNAME, SECTION, Me.Name & "_" & Me.tag & "_Moneda", "USD")    '*** MAKOTO 08/sep/00
        .fecha1 = dtpFecha1.value
        .fecha2 = dtpFecha2.value
        .CodTrans = PreparaCodTrans
        .codforma = PreparaCodForma
        gobjMain.objCondicion.CodMoneda = mCodMoneda
        'Estados no incluye anulados
        .EstadoBool(ESTADO_NOAPROBADO) = True
        .EstadoBool(ESTADO_APROBADO) = True
        .EstadoBool(ESTADO_DESPACHADO) = True
        .EstadoBool(ESTADO_ANULADO) = False
        
        Select Case Me.tag
        Case "Pago"
            s = PreparaTransParaGnopcion(.codforma)
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "FormaParaEgresoAutomatico", s
            s = PreparaTransParaGnopcion(.CodTrans)
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransParaEgresoAutomatico", s
            s = fcbTrans.KeyText
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransEgresoParaEgresoAutomatico", s
            s = txtEgreso.Text
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "EgresosAutoLibImpPago", s
            s = txtCheque.Text
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "EgresosAutoLibImpCheque", s
        Case "CruceTarjetas"
            s = PreparaTransParaGnopcion(.codforma)
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "FormaParaCruceTarjetas", s
            s = PreparaTransParaGnopcion(.CodTrans)
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransParaCruceTarjetas", s
            s = fcbTrans.KeyText
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransIngresoParaCruceTarjetas", s
            s = fcbFormaCobro.KeyText
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransFormaCobroParaCruceTarjetas", s
        Case "CruceIVTarjetas"
            s = PreparaTransParaGnopcion(.codforma)
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "FormaParaCruceIVTarjetas", s
            s = PreparaTransParaGnopcion(.CodTrans)
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransParaCruceIVTarjetas", s
            s = fcbTrans.KeyText
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransIngresoParaCruceIVTarjetas", s
            s = fcbFormaCobro.KeyText
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransFormaCobroParaCruceIVTarjetas", s
        Case "InicioAnulados"
            CreaTransAnulas
        Case Else
'            s = PreparaTransParaGnopcion(.codforma)
'            gobjMain.EmpresaActual.GNOpcion.AsignarValor "FormaParaIngresoAutomatico", s
            s = PreparaTransParaGnopcion(.CodTrans)
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransParaIngresoAutomatico", s
            s = fcbTrans.KeyText
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransIngresoParaIngresoAutomatico", s
        End Select
        
''''        If Me.tag <> "Pago" Then
''''            s = PreparaTransParaGnopcion(.codforma)
''''            gobjMain.EmpresaActual.GNOpcion.AsignarValor "FormaParaIngresoAutomatico", s
''''            s = PreparaTransParaGnopcion(.CodTrans)
''''            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransParaIngresoAutomatico", s
''''            s = fcbTrans.KeyText
''''            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransIngresoParaIngresoAutomatico", s
''''        Else
''''            s = PreparaTransParaGnopcion(.codforma)
''''            gobjMain.EmpresaActual.GNOpcion.AsignarValor "FormaParaEgresoAutomatico", s
''''            s = PreparaTransParaGnopcion(.CodTrans)
''''            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransParaEgresoAutomatico", s
''''            s = fcbTrans.KeyText
''''            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransEgresoParaEgresoAutomatico", s
''''
''''            s = txtEgreso.Text
''''            gobjMain.EmpresaActual.GNOpcion.AsignarValor "EgresosAutoLibImpPago", s
''''
''''            s = txtCheque.Text
''''            gobjMain.EmpresaActual.GNOpcion.AsignarValor "EgresosAutoLibImpCheque", s
''''
''''        End If
    End With
    Select Case Me.tag
    Case "InicioAnulados"
        cmdAceptar.Enabled = True
    Case Else
        'Graba en la base
        gobjMain.EmpresaActual.GNOpcion.Grabar
            
            
        
        grd.Rows = 1
            Set obj = gobjMain.EmpresaActual.ConsGNTransComprometidoAuto(chkAgrupaProv.value = vbChecked, True) 'Ascendente
    
        
        If Not obj.EOF Then
            v = MiGetRows(obj)
            grd.Redraw = flexRDNone
            grd.LoadArray v
            ConfigCols
            grd.Redraw = flexRDDirect
        Else
            grd.Rows = grd.FixedRows
            ConfigCols
        End If
        cmdAceptar.Enabled = True
        cmdAsiento.Enabled = False
        cmdImprimir.Enabled = False
        cmdImprimiCH.Enabled = False
        mVerificado = True
    End Select
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub ConfigCols()
    Dim i As Integer
    With grd
    
    
                .FormatString = "^#|tid|<Fecha Trans|<id|<Trans|<#|<#Ref.|<Nombre|<Descripción|<Cuenta|<Nombre|<Debe|>Haber|<Resultado|tidIn"
        
'            .ColHidden(COL_TID) = True
'            .ColHidden(COL_TIDIN) = True
'            .ColHidden(COL_ID) = True
            
            .ColWidth(COL_VALOR) = 1000
            
'''            Select Case Me.tag
'''                Case "Pago"
'''                Case "CruceTarjetas"
'''                    .ColHidden(COL_NUMDOCREF) = True
'''                    .ColHidden(COL_CODPROVCLI) = True
'''                    .ColHidden(COL_AUTO) = True
'''                    .ColHidden(COL_VALORPAGO) = True
'''                    .ColWidth(COL_VALOR) = 3000
'''                Case "CruceIVTarjetas"
'''                    .ColHidden(COL_NUMDOCREF) = True
'''                    .ColHidden(COL_CODPROVCLI) = True
'''                    .ColHidden(COL_AUTO) = True
'''                    .ColHidden(COL_VALORPAGO) = True
'''                    .ColWidth(COL_VALOR) = 3000
'''
'''                 Case Else
'''                    .ColHidden(COL_NUMDOCREF) = True
''''                    .ColHidden(COL_VALOR) = True
'''                    .ColHidden(COL_AUTO) = True
'''                   .ColHidden(COL_VALORPAGO) = True
'''            End Select
                
'            If Me.tag <> "Pago" Then
'                .ColHidden(COL_NUMDOCREF) = True
'                .ColHidden(COL_VALOR) = True
'                .ColHidden(COL_AUTO) = True
'                .ColHidden(COL_VALORPAGO) = True
'            End If
            
            .ColDataType(COL_FECHA) = flexDTDate   '*** MAKOTO 14/ago/2000 para que ordene bien por fecha
            
            If Me.tag = "Pago" Then
                .ColFormat(COL_VALOR) = "#.00"
                .ColFormat(COL_VALORPAGO) = "#.00"
            End If
        For i = 1 To COL_VALOR
            .ColData(i) = -1
        Next i
        .ColWidth(COL_NUMTRANS) = 800
            .SubtotalPosition = flexSTBelow
            grd.SubTotal flexSTClear
        
        
        If Me.tag <> "Pago" Then
            .ColFormat(COL_FECHA) = "dd/mm/yyyy"
            .ColWidth(COL_NOMBRE) = 2800
            .ColWidth(COL_FORMA) = 3200
            .ColWidth(COL_DESC) = 2400
            If chkAgrupaProv.value = vbChecked Then
                grd.SubTotal flexSTSum, COL_FORMA, 1, , grd.GridColor, vbBlack, , "Subtotal", COL_FORMA, True
            Else
                grd.SubTotal flexSTSum, COL_TID, 1, , grd.GridColor, vbBlack, , "Subtotal", COL_TID, True
            End If
        Else
            .ColWidth(COL_NOMBRE) = 2200
            .ColWidth(COL_DESC) = 1000
            If chkAgrupaProv.value = vbChecked Then
                grd.SubTotal flexSTSum, COL_NOMBRE, 10, , grd.GridColor, vbBlack, , "Subtotal", COL_NOMBRE, True
                grd.SubTotal flexSTSum, COL_NOMBRE, 12, , grd.GridColor, vbBlack, , "Subtotal", COL_NOMBRE, True
            End If
            
            
            .SubTotal flexSTSum, -1, 10, , vbBlue, vbWhite
            .SubTotal flexSTSum, -1, 12, , vbBlue, vbWhite
            .TextMatrix(.Rows - 1, 2) = "TOTAL"
            .ColDataType(COL_AUTO) = flexDTBoolean
            .ColWidth(COL_FORMA) = 1000
        End If
        GNPoneNumFila grd, False
        .AutoSize 0, grd.Cols - 1


'        .ColWidth(COL_VALOR) = 1000
        .ColWidth(COL_RESULTADO) = 2000
        
    End With
    PoneColorFilas
End Sub

Private Sub cmdCancelar_Click()
    If mProcesando Then
        mCancelado = True
    Else
        Unload Me
    End If
End Sub




Private Sub cmdExplorar_Click()
    On Error GoTo ErrTrap
    
    With dlg1
        If Len(.filename) = 0 Then
            .InitDir = txtEgreso.Text
            '.FileName = mPlantilla.BDDestino
        Else
            .InitDir = .filename
            '.FileName = mPlantilla.BDDestino
        End If
        .flags = cdlOFNPathMustExist
        .Filter = "Base de datos Jet (*.txt)|*.txt|Predefinido *.txt |Todos (*.*)|*.*"
        .ShowSave
        txtEgreso.Text = .filename
    End With
    
    Exit Sub
ErrTrap:
    If Err.Number <> 32755 Then
        DispErr
    End If
    Exit Sub
End Sub

Private Sub cmdExplorarCH_Click()
On Error GoTo ErrTrap
    
    With dlg1
        If Len(.filename) = 0 Then
            .InitDir = txtCheque.Text
            '.FileName = mPlantilla.BDDestino
        Else
            .InitDir = .filename
            '.FileName = mPlantilla.BDDestino
        End If
        .flags = cdlOFNPathMustExist
        .Filter = "Base de datos Jet (*.txt)|*.txt|Predefinido *.txt |Todos (*.*)|*.*"
        .ShowSave
        txtCheque.Text = .filename
    End With
    
    Exit Sub
ErrTrap:
    If Err.Number <> 32755 Then
        DispErr
    End If
    Exit Sub
End Sub

Private Sub cmdImprimiCH_Click()
    'Si no hay transacciones
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna transacción para imprimir."
        Exit Sub
    End If
    
    If ImprimirCheque Then
        cmdCancelar.SetFocus
    End If
End Sub

Private Sub cmdImprimir_Click()
    'Si no hay transacciones
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna transacción para imprimir."
        Exit Sub
    End If
    
    If Imprimir Then
        cmdCancelar.SetFocus
    End If
End Sub



Private Sub fcbBanco_Selected(ByVal Text As String, ByVal KeyText As String)
    Dim saldo As Currency, tsb As TSBanco
    saldo = Format(gobjMain.EmpresaActual.ConsTSSaldoActualBanco(Text, Date), "#0.00")
    lblsaldoBanco.Caption = saldo
    Set tsb = gobjMain.EmpresaActual.RecuperaTSBanco(fcbBanco.KeyText)
    lblNumCheque.Caption = tsb.NumChequeSiguiente
    Set tsb = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF3
        KeyCode = 0
    Case vbKeyF5
        KeyCode = 0
    Case vbKeyF6
        dtpFecha1.SetFocus
        KeyCode = 0
    Case vbKeyF7
        KeyCode = 0
    Case vbKeyF8
        KeyCode = 0
    Case vbKeyEscape
        Unload Me
    Case Else
        MoverCampo Me, KeyCode, Shift, True
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ImpideSonidoEnter Me, KeyAscii
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mProcesando Then
        Cancel = True
        Exit Sub
    End If
    Me.Hide         'Se pone esto para evitar el posible BUG de Windows98
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    With grd
        If Me.tag <> "Pago" Then
                .Top = FraConFigEgreso.Top
                .Width = Me.ScaleWidth - 200
                .Height = Me.ScaleHeight - .Top - pic1.Height - 380
        Else
                .Top = FraConFigEgreso.Top + FraConFigEgreso.Height
                .Width = Me.ScaleWidth - 200
                .Height = Me.ScaleHeight - .Top - pic1.Height '- 380
        End If
    End With
    prg1.Width = Me.ScaleWidth - (prg1.Left * 2)
        
End Sub


Private Sub grd_CellChanged(ByVal Row As Long, ByVal col As Long)
''    grd.subtotal flexSTSum, -1, 10, , vbBlue, vbWhite
''    grd.subtotal flexSTSum, -1, 12, , vbBlue, vbWhite
    grd.Refresh
End Sub

Private Sub grd_KeyUp(KeyCode As Integer, Shift As Integer)
    If Me.tag = "Pago" Then
        If Not grd.IsSubtotal(grd.Row) And grd.Row <> 0 Then
            If grd.ValueMatrix(grd.Row, COL_VALORPAGO) = 0 And CBool(grd.TextMatrix(grd.Row, COL_AUTO)) Then
                grd.TextMatrix(grd.Row, COL_VALORPAGO) = grd.ValueMatrix(grd.Row, COL_VALOR)
            End If
            If Not CBool(grd.TextMatrix(grd.Row, COL_AUTO)) Then
                grd.TextMatrix(grd.Row, COL_VALORPAGO) = 0
            End If
        End If
    End If
End Sub

Private Sub grd_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Me.tag = "Pago" Then
        If Not grd.IsSubtotal(grd.Row) And grd.Row <> 0 Then
            If grd.ValueMatrix(grd.Row, COL_VALORPAGO) = 0 And CBool(grd.TextMatrix(grd.Row, COL_AUTO)) Then
                grd.TextMatrix(grd.Row, COL_VALORPAGO) = grd.ValueMatrix(grd.Row, COL_VALOR)
            End If
            If Not CBool(grd.TextMatrix(grd.Row, COL_AUTO)) Then
                grd.TextMatrix(grd.Row, COL_VALORPAGO) = 0
            End If
        End If
    End If
End Sub



Private Sub grd_SelChange()
'    grd.subtotal flexSTClear
    If Me.tag <> "Pago" Then
        If chkAgrupaProv.value = vbChecked Then
            grd.SubTotal flexSTSum, COL_FORMA, 1, , grd.GridColor, vbBlack, , "Subtotal", COL_FORMA, True
        End If
    Else
        If chkAgrupaProv.value = vbChecked Then
            grd.SubTotal flexSTSum, COL_NOMBRE, 10, , grd.GridColor, vbBlack, , "Subtotal", COL_NOMBRE, True
            grd.SubTotal flexSTSum, COL_NOMBRE, 12, , grd.GridColor, vbBlack, , "Subtotal", COL_NOMBRE, True
        End If
        grd.SubTotal flexSTSum, -1, 10, , vbBlue, vbWhite
        grd.SubTotal flexSTSum, -1, 12, , vbBlue, vbWhite
    End If
    grd.Refresh
End Sub


Private Sub mobjGNComp_EstadoGeneracion1AsientoxLote(ByVal ix As Long, ByVal Estado As String, Cancel As Boolean)
    prg1.value = ix
    grd.TextMatrix(ix, COL_RESULTADO) = Estado
    Cancel = mCancelado
End Sub

Private Sub CargaFormas()
    Dim i As Long, v As Variant
    Dim s As String

    lstForma.Clear
    v = gobjMain.EmpresaActual.ListaTSFormaCobroPagoIngresoAuto(True, True, False)
    For i = LBound(v, 2) To UBound(v, 2)
        lstForma.AddItem v(0, i)        '& " " & v(1, i)
    Next i
    
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("FormaParaIngresoAutomatico")) > 0 Then
        s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("FormaParaIngresoAutomatico")
        RecuperaTrans "KeyT", lstForma, s
    End If
End Sub


Private Function PreparaCodForma() As String
    Dim i As Long, s As String
    
    With lstForma
        'Si está seleccionado solo una
        If lstForma.SelCount = 1 Then
            For i = 0 To .ListCount - 1
                If .Selected(i) Then
                    s = s & "'" & .List(i) & "'"
                    Exit For
                End If
            Next i
        'Si está TODO o NINGUNO, no hay condición
        ElseIf (.SelCount > 0) Then
            For i = 0 To .ListCount - 1
                If .Selected(i) Then
                    s = s & "'" & .List(i) & "', "
                End If
            Next i
            If Len(s) > 0 Then s = Left$(s, Len(s) - 2)    'Quita la ultima ", "
        End If
    End With
    PreparaCodForma = s
End Function

'jeaa 25/09/2006 elimina los apostrofes
Private Function PreparaTransParaGnopcion(cad As String) As String
    Dim v As Variant, i As Integer, s As String, pos As Integer
    s = ""
    v = Split(cad, ",")
    For i = 0 To UBound(v)
        v(i) = Trim(v(i))
        pos = InStr(1, v(i), "'")
        If pos <> 0 Then
            s = s & Mid$(v(i), 2, Len(v(i)) - 2) & ","
        Else
            s = s & v(i) & ","
        End If
    Next i
    'quita ultima coma
    PreparaTransParaGnopcion = Mid$(s, 1, Len(s) - 1)
End Function


Public Sub RecuperaTrans(ByVal Key As String, lst As ListBox, Optional s As String)
Dim Vector As Variant
Dim i As Integer, j As Integer, Selec As Integer
'Dim S As String
    If s <> "_VACIO_" Then
        Vector = Split(s, ",")
         Selec = UBound(Vector, 1)
         For i = 0 To Selec
            For j = 0 To lst.ListCount - 1
'                If Vector(i) = Left(lst.List(j), lst.ItemData(j)) Then
                If Trim(Vector(i)) = lst.List(j) Then
                    lst.Selected(j) = True
                End If
            Next j
         Next i
    End If
End Sub

Private Function PreparaCodTrans() As String
    Dim i As Long, s As String
    
    With lstTrans
        'Si está seleccionado solo una
        If lstTrans.SelCount = 1 Then
            For i = 0 To .ListCount - 1
                If .Selected(i) Then
                    s = s & "'" & .List(i) & "'"
                    Exit For
                End If
            Next i
        'Si está TODO o NINGUNO, no hay condición
        ElseIf (.SelCount > 0) Then
            For i = 0 To .ListCount - 1
                If .Selected(i) Then
                    s = s & "'" & .List(i) & "', "
                End If
            Next i
            If Len(s) > 0 Then s = Left$(s, Len(s) - 2)    'Quita la ultima ", "
        End If
    End With
    PreparaCodTrans = s
End Function


Private Function IngresoAuto(ByVal bandVerificar As Boolean, BandTodo As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, x As Single, j As Integer, filaSubTotal As Long
    Dim gnc As GNComprobante, cambiado As Boolean, TransGen As String
    
    On Error GoTo ErrTrap
    
    'Si no es solo verificacion, confirma
    If Not bandVerificar Then
        'Confirma la actualización
        s = "Este proceso creará Ingresos Automáticos  de la transacción seleccionada." & vbCr & vbCr
        s = s & "Está seguro que desea proceder?"
        If MsgBox(s, vbYesNo + vbQuestion) <> vbYes Then Exit Function
    End If
    
    'Verifica si está seleccionado una trans. de ingreso
    s = VerificaIngresoAutomatico
    If Len(s) > 0 Then
        'Si está seleccinada, confirma si está seguro
        s = "Está seleccionada una o más transacciones de ingreso. " & vbCr & _
            "(" & s & ")" & vbCr & _
            "Generalmente no se hace Ingresos Automáticos con transacciones de ingreso." & vbCr & vbCr
        s = s & "Confirma que desea proceder?" & vbCr & _
            "Aplaste 'Sí' unicamente cuando está seguro de lo que está haciendo."
        If MsgBox(s, vbYesNo + vbQuestion + vbDefaultButton2) <> vbYes Then Exit Function
    End If
    s = ""
    
    Set mColItems = Nothing     'Limpia lo anterior
    Set mColItems = New Collection
    
    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdAceptar.Enabled = False
    cmdBuscar.Enabled = False
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    
    For i = grd.FixedRows To grd.Rows - 1
        DoEvents
        If mCancelado Then
            MsgBox "El proceso fue cancelado.", vbInformation
            Exit For
        End If
        
        prg1.value = i
        grd.Row = i
        x = grd.CellTop                 'Para visualizar la celda actual
        
        If Not grd.IsSubtotal(i) Then
        'Si es verificación procesa todas las filas sino solo las que tengan "Costo Incorrecto"
        
            tid = grd.ValueMatrix(i, COL_TID)
            grd.TextMatrix(i, COL_RESULTADO) = "Procesando  ..."
            grd.Refresh
            
            'Recupera la transaccion
            
            Set mobjGNCompOrigen = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
'            If gnc.numtrans = 50 Then MsgBox "HOLA"
            If Not (mobjGNCompOrigen Is Nothing) Then
                'Si la transacción es de Inventario y es Egreso/Transferencia
                ' Y no está anulado
                If ((mobjGNCompOrigen.GNTrans.Modulo = "TS") Or (mobjGNCompOrigen.GNTrans.Modulo = "CT")) And _
                   (mobjGNCompOrigen.Estado <> ESTADO_ANULADO) Then
'                   (gnc.GNTrans.IVTipoTrans = "E" Or gnc.GNTrans.IVTipoTrans = "T") And _      '*** MAKOTO 06/sep/00 Eliminado

                    'Forzar recuperar todos los datos de transacción para que no se pierdan al grabar de nuveo
                    mobjGNCompOrigen.RecuperaDetalleTodo
                    'Recalcula costo de los items
                    If chkAgrupaProv.value <> vbChecked Then
'                        If GrabarTransAuto(TransGen) Then
'                                'Graba la transacción
'                                grd.TextMatrix(i, COL_RESULTADO) = "OK.. Trans " & TransGen
'                                grd.TextMatrix(i, COL_TIDIN) = mobjGNCompAux.TransID
'
'                        Else
'                                'Si no está cambiado no graba
'                                grd.TextMatrix(i, COL_RESULTADO) = "Falló Proceso"
'                        End If
                            For j = i To grd.Rows - 1
                                If grd.IsSubtotal(j) Then
                                    filaSubTotal = j
                                    j = grd.Rows - 1
                                End If
                            Next j
                            TransGen = fcbTrans.KeyText
                        If GrabarTransAutoxTrans(TransGen, i, filaSubTotal) Then
                                    'Graba la transacción
                                    For j = i To filaSubTotal - 1
                                            If j = filaSubTotal - 1 Then
                                                grd.TextMatrix(j, COL_RESULTADO) = "OK.. Trans " & TransGen
                                            Else
                                                grd.TextMatrix(j, COL_RESULTADO) = "Trans " & TransGen
                                            End If
                                            grd.TextMatrix(j, COL_TIDIN) = mobjGNCompAux.TransID
                                    Next j
                                    i = filaSubTotal
                            Else
                                    'Si no está cambiado no graba
                                    grd.TextMatrix(i, COL_RESULTADO) = "Falló Proceso"
                            End If
                    Else
                            For j = i To grd.Rows - 1
                                If grd.IsSubtotal(j) Then
                                    filaSubTotal = j
                                    j = grd.Rows - 1
                                End If
                            Next j
                        If GrabarTransAutoxForma(TransGen, i, filaSubTotal) Then
                                    'Graba la transacción
                                    For j = i To filaSubTotal - 1
                                            If j = filaSubTotal - 1 Then
                                                grd.TextMatrix(j, COL_RESULTADO) = "OK.. Trans " & TransGen
                                            Else
                                                If grd.TextMatrix(j, COL_RESULTADO) = "Procesando" Then
                                                    grd.TextMatrix(j, COL_RESULTADO) = "Trans " & TransGen
                                                End If
                                            End If
                                            grd.TextMatrix(j, COL_TIDIN) = mobjGNCompAux.TransID
                                    Next j
                                    i = filaSubTotal
                            Else
                                    'Si no está cambiado no graba
                                    grd.TextMatrix(i, COL_RESULTADO) = "Falló Proceso"
                                    i = filaSubTotal
                            End If
                    End If
                Else
                    'Si está anulado
                    If gnc.Estado = ESTADO_ANULADO Then
                        grd.TextMatrix(i, COL_RESULTADO) = "Anulado"
                    'Si no tiene nada que ver con recalculo de costo
                    Else
                        grd.TextMatrix(i, COL_RESULTADO) = "---"
                    End If
                End If
            Else
                grd.TextMatrix(i, COL_RESULTADO) = "No pudo recuperar la transación."
            End If
        End If
    Next i
    
    Screen.MousePointer = 0
'''    ReprocCosto = Not mCancelado
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    If i < grd.Rows And i >= grd.FixedRows Then
        grd.TextMatrix(i, COL_RESULTADO) = Err.Description
    End If
    DispErr
    prg1.value = prg1.min
salida:
    Set mColItems = Nothing         'Libera el objeto de coleccion
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdBuscar.Enabled = True
    cmdAceptar.Enabled = True
    prg1.value = prg1.min
    Exit Function
End Function


Private Function VerificaIngresoAutomatico() As String
    Dim i As Long, cod As String, gnt As GNTrans
    Dim s As String
    
    For i = 0 To lstTrans.ListCount - 1
        'Si está seleccionado
        If lstTrans.Selected(i) Then
            'Recupera el objeto GNTrans
            cod = lstTrans.List(i)
            Set gnt = gobjMain.EmpresaActual.RecuperaGNTrans(cod)
            'Si la transaccion es de ingreso, devuelve el codigo
            If gnt.IVTipoTrans = "I" Then s = s & cod & ", "
        End If
    Next i
    Set gnt = Nothing
    If Len(s) > 2 Then s = Left$(s, Len(s) - 2)     'Quita la ultima ", "
    VerificaIngresoAutomatico = s
End Function

Private Function GrabarTransAuto(ByRef trans As String) As Boolean
    Dim Imprime As Boolean, i As Long, ix As Long, orden1 As Integer, orden2 As Integer
    Dim pc As PCProvCli, Cadena As String, obser As String
    Dim tsf As TSFormaCobroPago
    On Error GoTo ErrTrap
    GrabarTransAuto = True
    orden1 = 1
    orden2 = 1
    If CreaComprobanteIngresoAuto(i) Then
        'Si es solo lectura, no hace nada
        If mobjGNCompAux.SoloVer Then
            MsgBox MSG_NODISPONE, vbInformation
            Exit Function
        End If
        'carga la nueva deuda a los bancos de las tarjetas
        For i = 1 To mobjGNCompOrigen.CountPCKardex
            Set tsf = mobjGNCompOrigen.Empresa.RecuperaTSFormaCobroPago(mobjGNCompOrigen.PCKardex(i).codforma)
            If tsf.IngresoAutomatico Then
                If tsf.DeudaMismoCliente Then
                    Set pc = mobjGNCompOrigen.Empresa.RecuperaPCProvCli(mobjGNCompOrigen.CodClienteRef)
                Else
                    Set pc = mobjGNCompOrigen.Empresa.RecuperaPCProvCli(tsf.CodProvCli)
                End If
    
                
                If tsf.IngresoAutomatico Then
                    ix = mobjGNCompAux.AddPCKardex
                    mobjGNCompAux.PCKardex(ix).Debe = mobjGNCompOrigen.PCKardex(i).Debe
                    If tsf.DeudaMismoCliente Then
                        mobjGNCompAux.PCKardex(ix).CodProvCli = mobjGNCompOrigen.CodClienteRef
                    Else
                        mobjGNCompAux.PCKardex(ix).CodProvCli = tsf.CodProvCli
                    End If
                    mobjGNCompAux.PCKardex(ix).codforma = tsf.CodFormaTC
                    mobjGNCompAux.PCKardex(ix).NumLetra = mobjGNCompOrigen.CodTrans & " " & mobjGNCompOrigen.GNTrans.NumTransSiguiente
                    mobjGNCompAux.PCKardex(ix).FechaEmision = mobjGNCompOrigen.FechaTrans
                    mobjGNCompAux.PCKardex(ix).FechaVenci = mobjGNCompOrigen.FechaTrans
                    obser = "Por pago con: " & tsf.codforma & " de " & mobjGNCompOrigen.CodTrans & "-" & mobjGNCompOrigen.numtrans & " Cliente: " & mobjGNCompOrigen.CodClienteRef & " - " & mobjGNCompOrigen.nombre
                    mobjGNCompAux.PCKardex(ix).Observacion = IIf(Len(obser) > 80, Left(obser, 80), obser)
                    mobjGNCompAux.PCKardex(ix).CodVendedor = mobjGNCompOrigen.CodVendedor
                    mobjGNCompAux.PCKardex(ix).orden = orden1
                    orden1 = orden1 + 1
                    'PAGO
                    ix = mobjGNCompAux.AddPCKardex
                    mobjGNCompAux.PCKardex(ix).Haber = mobjGNCompOrigen.PCKardex(i).Debe
                    mobjGNCompAux.PCKardex(ix).CodProvCli = mobjGNCompOrigen.CodClienteRef
                    mobjGNCompAux.PCKardex(ix).codforma = tsf.CodFormaTC
                    mobjGNCompAux.PCKardex(ix).idAsignado = mobjGNCompOrigen.PCKardex(i).id
                    mobjGNCompAux.PCKardex(ix).NumLetra = mobjGNCompOrigen.CodTrans & " " & mobjGNCompOrigen.GNTrans.NumTransSiguiente
                    mobjGNCompAux.PCKardex(ix).FechaEmision = mobjGNCompOrigen.FechaTrans
                    mobjGNCompAux.PCKardex(ix).CodVendedor = mobjGNCompOrigen.CodVendedor
                    mobjGNCompAux.PCKardex(ix).orden = orden1
                    orden1 = orden1 + 1
                Else
    '                MsgBox "La forma de cobro " & tsf.NombreForma & " No esta Configurado para Ingreso Automático"
    '                    GrabarTransAuto = False
    '                Exit Function
                End If
            End If
        Next i

        mobjGNCompAux.FechaTrans = mobjGNCompOrigen.FechaTrans
        mobjGNCompAux.HoraTrans = mobjGNCompOrigen.HoraTrans
        If mobjGNCompOrigen.CountPCKardex > 1 Then
            Cadena = "Por pago con varias formas de " & mobjGNCompOrigen.CodTrans & "-" & mobjGNCompOrigen.numtrans & " Cliente: " & mobjGNCompOrigen.CodClienteRef & " - " & mobjGNCompOrigen.nombre & " / Banco: " & pc.nombre
        Else
             Cadena = "Por pago con :" & mobjGNCompOrigen.PCKardex(1).codforma & " de " & mobjGNCompOrigen.CodTrans & "-" & mobjGNCompOrigen.numtrans & " Cliente: " & mobjGNCompOrigen.CodClienteRef & " - " & mobjGNCompOrigen.nombre & " / Banco: " & pc.nombre
        End If
        If Len(Cadena) > 120 Then
            mobjGNCompAux.Descripcion = Mid$(Cadena, 1, 120)
        Else
            mobjGNCompAux.Descripcion = Cadena
        End If
            
        mobjGNCompAux.codUsuario = mobjGNCompOrigen.codUsuario
        mobjGNCompAux.IdResponsable = mobjGNCompOrigen.IdResponsable
        mobjGNCompAux.numDocRef = mobjGNCompOrigen.CodTrans & " " & mobjGNCompOrigen.numtrans
        mobjGNCompAux.idCentro = mobjGNCompOrigen.idCentro
        mobjGNCompAux.IdTransFuente = mobjGNCompOrigen.TransID
        mobjGNCompAux.CodMoneda = mobjGNCompOrigen.CodMoneda
        If mobjGNCompOrigen.CountPCKardex > 1 Then
'            mobjGNCompAux.nombre = ""
'            mobjGNCompAux.CodClienteRef = ""
            mobjGNCompAux.nombre = mobjGNCompOrigen.nombre
            mobjGNCompAux.CodClienteRef = mobjGNCompOrigen.CodClienteRef
        Else
            mobjGNCompAux.nombre = mobjGNCompOrigen.nombre
            mobjGNCompAux.CodClienteRef = mobjGNCompAux.PCKardex(i).CodProvCli
        End If
        mobjGNCompAux.CodVendedor = mobjGNCompOrigen.CodVendedor
        mobjGNCompAux.IdTransFuente = mobjGNCompAux.TransID
    
        'Si es que algo está modificado
        If mobjGNCompAux.Modificado Then
            MensajeStatus MSG_GENERANDOASIENTO, vbHourglass
            MensajeStatus
        End If
        If mobjGNCompAux.GNTrans.AfectaSaldoPC And _
           mobjGNCompAux.GNTrans.TSVerificaTotalCuadrado Then
            'Verifica si está cuadrado el total de transacción y total de PCKardex.
            If Not TotalCuadrado Then Exit Function
        End If
        'Verificación de datos
        mobjGNCompAux.VerificaDatos
    
        'Verifica si está cuadrado el asiento
        If Not VerificaAsiento(mobjGNCompAux) Then Exit Function
    
        'Verifica si tiene detalle de banco
        If (mobjGNCompAux.CountTSKardex = 0) And _
            (mobjGNCompAux.CountTSKardexRet = 0) And _
            (mobjGNCompAux.CountPCKardex = 0) Then
            MsgBox "No existe ningún detalle.", vbInformation
            Exit Function
        End If
    
        MensajeStatus MSG_GRABANDO, vbHourglass
    
        'Manda a grabar
        '       Aquí ya no hacemos verificación de asiento por que ya está hecho en Control Asiento
        mobjGNCompAux.Grabar False, False
        
        
'        grd.TextMatrix(grd.Row, COL_TIDIN) = mobjGNCompAux.TransID
'        grd.TextMatrix(grd.Row, COL_NUMTRANSIN) = mobjGNCompAux.numtrans
        
        '***  Oliver 26/12/2002
        'Agregado para el control ded Impresion Configurado en la Transaccion
        
        
        MensajeStatus
    '    Me.caption = "Transacción " & mobjGNCompAux.codTrans & " " & mobjGNCompAux.NumTrans
        Me.Caption = mobjGNCompAux.CodTrans & " " & mobjGNCompAux.numtrans
        trans = mobjGNCompAux.CodTrans & " " & mobjGNCompAux.numtrans
        GrabarTransAuto = True
    Else
        GrabarTransAuto = False
    End If
    Exit Function
ErrTrap:
    MensajeStatus
    Select Case Err.Number
    Case ERR_DESCUADRADO, ERR_INTEGRIDAD
        'Si es que el usuario seleccionó 'No' en el cuadro de dialogo,
        'No hace nada
    Case Else
        DispErr
    End Select
    GrabarTransAuto = False
    Exit Function
    
End Function



Private Function CreaComprobanteIngresoAuto(ByRef Num As Long) As Boolean
    Dim v As Currency, tsf As TSFormaCobroPago
    Dim i As Long
    CreaComprobanteIngresoAuto = False
    For i = 1 To mobjGNCompOrigen.CountPRLibroDetalle
'        Set tsf = mobjGNCompOrigen.Empresa.RecuperaTSFormaCobroPago(mobjGNCompOrigen.PCKardex(i).codforma)
 '       If Not tsf Is Nothing Then
'                If tsf.IngresoAutomatico Then
'                    If Len(mobjGNCompOrigen.GNTrans.IVTransAuto) > 0 Then
                        If Len(fcbTrans.KeyText) > 0 Then
                            Set mobjGNCompAux = gobjMain.EmpresaActual.CreaGNComprobante(fcbTrans.KeyText)
                        Else
                            MsgBox "Falta seleccionar transaccion de Ingreso"
                            CreaComprobanteIngresoAuto = False
                            Exit Function
                        End If
'                    End If
                    CreaComprobanteIngresoAuto = True
'                End If
        'End If
        Set tsf = Nothing
    Next i
  

End Function

Private Function DatosIngresoAuto(ByRef Num As Long, ByRef TRansIT As String, ByRef FormaCobro As String, ByRef CodProvCli As String, ByRef valor As Currency) As Boolean
    Dim v As Currency, tsf As TSFormaCobroPago
    Dim i As Long
    DatosIngresoAuto = False
    For i = 1 To mobjGNCompOrigen.CountPCKardex
        Set tsf = mobjGNCompOrigen.Empresa.RecuperaTSFormaCobroPago(mobjGNCompOrigen.PCKardex(i).codforma)
        If Not tsf Is Nothing Then
                If tsf.IngresoAutomatico Then
                    If Len(mobjGNCompOrigen.GNTrans.IVTransAuto) > 0 Then
                        TRansIT = mobjGNCompOrigen.GNTrans.IVTransAuto
                    End If
                    If Len(tsf.CodFormaTC) > 0 Then
                        FormaCobro = tsf.CodFormaTC
                    End If
                    If tsf.DeudaMismoCliente Then
                        CodProvCli = mobjGNCompOrigen.CodClienteRef
                    Else
                        If Len(tsf.CodProvCli) > 0 Then
                            CodProvCli = tsf.CodProvCli
                        End If
                    End If
                    valor = mobjGNCompOrigen.PCKardex(i).Debe
                    Num = i
                    i = mobjGNCompOrigen.CountPCKardex
                    DatosIngresoAuto = True
                End If
        End If
        Set tsf = Nothing
    Next i
End Function





Private Function TotalCuadrado() As Boolean
    Dim t As Currency, p As Currency
    With mobjGNCompOrigen
        t = .IVKardexTotal(True)
        t = MiCCur(Format$(t, .FormatoMoneda))  'Redondea al formato de moneda
        t = t + .IVRecargoTotal(True, False) * Sgn(t)
        p = .PCKardexHaberTotal - .PCKardexDebeTotal + .TSKardexHaberTotal - .TSKardexDebeTotal
        
        If t <> p Then
            MsgBox "El valor total de transacción (" & Format(t, "#,0.0000") & _
                   ") y forma de pago/cobro (" & Format(p, "#,0.0000") & _
                   ") no están cuadrados por la diferencia de " & _
                        Format(t - p, "#,0.0000") & " " & _
                        mobjGNCompOrigen.CodMoneda & "." & vbCr & vbCr & _
                   "Para grabar la transacción tiene que estar cuadrado.", vbInformation
            TotalCuadrado = False
        Else
            TotalCuadrado = True
        End If
    End With
End Function


Private Function RegenerarAsiento(bandVerificar As Boolean, BandTodo As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, x As Single, pos As Integer
    Dim gnc As GNComprobante, cambiado As Boolean
    
    On Error GoTo ErrTrap

    'Si no es solo verificacion, confirma
    If Not bandVerificar Then
        s = "Este proceso modificará los asientos de la transacción seleccionada." & vbCr & vbCr
        s = s & "Está seguro que desea proceder?"
        If MsgBox(s, vbYesNo + vbQuestion) <> vbYes Then Exit Function
    End If
    
    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
'    cmdVerificar.Enabled = False
    cmdBuscar.Enabled = False
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    
    For i = grd.FixedRows To grd.Rows - 1
        DoEvents
        If mCancelado Then
            MsgBox "El proceso fue cancelado.", vbInformation
            Exit For
        End If
        
        prg1.value = i
        grd.Row = i
        x = grd.CellTop                 'Para visualizar la celda actual
        
        pos = InStr(1, grd.TextMatrix(i, COL_RESULTADO), "OK")
        'Si es verificación, procesa todas las filas sino solo las que tengan "Asiento incorrecto."
        If pos <> 0 Then
        
            tid = grd.ValueMatrix(i, COL_TIDIN)
            grd.TextMatrix(i, COL_RESULTADO) = "Verificando..."
            grd.Refresh
            
            'Recupera la transaccion
            Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
            If Not (gnc Is Nothing) Then
                'Si la transacción no está anulada
                If gnc.Estado <> ESTADO_ANULADO Then
                
                    'Forzar recuperar todos los datos de transacción para que no se pierdan al grabar de nuveo
                    gnc.RecuperaDetalleTodo
                
                    'Recalcula costo de los items
                    If RegenerarAsientoSub(gnc, cambiado) Then
                        'Si está cambiado algo o está forzado regenerar todo
                        If cambiado Or BandTodo Then
                            'Si no es solo verificacion
                            If (Not bandVerificar) Or BandTodo Then
                                grd.TextMatrix(i, COL_RESULTADO) = "Grabando..."
                                grd.Refresh
                                
                                'Graba la transacción
                                gnc.Grabar False, False
                                grd.TextMatrix(i, COL_RESULTADO) = "OK..Actualizado."
                                
                            'Si es solo verificacion
                            Else
                                grd.TextMatrix(i, COL_RESULTADO) = MSG_NG
                            End If
                        Else
                            'Si no está cambiado no graba
                            grd.TextMatrix(i, COL_RESULTADO) = "OK."
                        End If
                    Else
                        grd.TextMatrix(i, COL_RESULTADO) = "Falló al regenerar."
                    End If
                Else
                    'Si está anulada
                    grd.TextMatrix(i, COL_RESULTADO) = "Anulado."
                End If
            Else
                grd.TextMatrix(i, COL_RESULTADO) = "No pudo recuperar la transación."
            End If
        Else
            If Len(grd.TextMatrix(i, COL_RESULTADO)) > 0 Then
                grd.TextMatrix(i, COL_RESULTADO) = "Procesada"
            End If
        End If
    Next i
    
    Screen.MousePointer = 0
    RegenerarAsiento = Not mCancelado
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    DispErr
salida:
    mProcesando = False
    frmMain.mnuFile.Enabled = True
'    cmdVerificar.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    Exit Function
End Function


Private Function RegenerarAsientoSub(ByVal gnc As GNComprobante, _
                                     ByRef cambiado As Boolean) As Boolean
    Dim i As Long, cta As CtCuenta, ctd As CTLibroDetalle
    Dim colCtd As Collection, a As clsAsiento
    On Error GoTo ErrTrap
    
    cambiado = False
    Set colCtd = New Collection
    
    'Guarda todos los detalles de asiento en la colección para después comparar
    With gnc
        For i = 1 To .CountCTLibroDetalle
            Set ctd = .CTLibroDetalle(i)
            Set a = New clsAsiento
            a.IdCuenta = ctd.IdCuenta
            a.Debe = ctd.Debe
            a.Haber = ctd.Haber
            colCtd.Add item:=a
        Next i
    End With
    
    'Regenera el asiento
    gnc.GeneraAsiento
    
    'Compara el asiento para saber si ha cambiado o no
    cambiado = Not CompararAsiento(gnc, colCtd)
    
    RegenerarAsientoSub = True
    GoTo salida
    Exit Function
ErrTrap:
    cambiado = False
    DispErr
    RegenerarAsientoSub = False
salida:
    Set a = Nothing
    Set colCtd = Nothing
    Set gnc = Nothing
    Exit Function
End Function


'Devuelve True si los asientos son iguales, False si no lo son
Private Function CompararAsiento(ByVal gnc As GNComprobante, ByVal col As Collection) As Boolean
    Dim a As clsAsiento, i As Long, ctd As CTLibroDetalle
    Dim encontrado As Boolean
    
    'Si número de detalles son diferentes ya no son iguales
    If col.Count <> gnc.CountCTLibroDetalle Then Exit Function
    
    For i = 1 To gnc.CountCTLibroDetalle
        Set ctd = gnc.CTLibroDetalle(i)
        encontrado = False
        For Each a In col
            If (ctd.IdCuenta = a.IdCuenta) And _
               (ctd.Debe = a.Debe) And _
               (ctd.Haber = a.Haber) And _
               (a.Comparado = False) Then
                a.Comparado = True
                encontrado = True
                Exit For
            End If
        Next a
        'Si no se encuentra uno igual
        If Not encontrado Then
            CompararAsiento = False
            Exit Function
        End If
    Next i
    CompararAsiento = True
End Function


Private Function Imprimir() As Boolean
    Dim s As String, tid As Long, i As Long, x As Single, res As String, pos As Integer
    Dim gnc As GNComprobante, cambiado As Boolean, cntError As Long
    
    On Error GoTo ErrTrap

    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdBuscar.Enabled = False
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    
    For i = grd.FixedRows To grd.Rows - 1
        DoEvents
        If mCancelado Then
            MsgBox "El proceso fue cancelado."
            Exit For
        End If
        
        prg1.value = i
        grd.Row = i
        x = grd.CellTop                 'Para visualizar la celda actual
        pos = InStr(1, grd.TextMatrix(i, COL_RESULTADO), "OK")
        'Si es verificación, procesa todas las filas sino solo las que tengan "Asiento incorrecto."
        If pos <> 0 Then
        
            tid = grd.ValueMatrix(i, COL_TIDIN)
            grd.TextMatrix(i, COL_RESULTADO) = "Procesando ..."
            grd.Refresh
            
            'Recupera la transaccion
            Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
            If Not (gnc Is Nothing) Then
                'Si la transacción no está anulado
                If gnc.Estado <> ESTADO_ANULADO Then
    '                'Forzar recuperar todos los datos de transacción
    '                ' para que no se pierdan al grabar de nuveo
    '                gnc.RecuperaDetalleTodo
                
                    'Imprime la transaccion o asiento contable
                    res = ImprimeTrans(gnc, False)
                    If Len(res) = 0 Then
                        grd.TextMatrix(i, COL_RESULTADO) = "OK..Enviado."
                    Else
                        grd.TextMatrix(i, COL_RESULTADO) = res
                        cntError = cntError + 1
                    End If
                                
                'Si la transaccion está anulado
                Else
                    grd.TextMatrix(i, COL_RESULTADO) = "Anulado."
                    cntError = cntError + 1
                End If
            Else
                grd.TextMatrix(i, COL_RESULTADO) = "No pudo recuperar la transación."
                cntError = cntError + 1
            End If
        End If
    Next i
    
    Screen.MousePointer = 0
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdImprimir.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    
    'Si algúna transaccion no se imprimió, avisa
    If cntError Then
        MsgBox "No se pudo imprimir " & cntError & " transacciones.", vbInformation
    End If
    
    Imprimir = True
    Exit Function
ErrTrap:
    Screen.MousePointer = 0
    DispErr
    prg1.value = prg1.min
    Exit Function
End Function

Public Function ImprimeTrans(ByVal gc As GNComprobante, ByVal bandAsiento As Boolean) As String
    Dim crear As Boolean
    Static objImp As Object
    On Error GoTo ErrTrap

    'Si no tiene TransID quiere decir que no está grabada
    If (gc.TransID = 0) Or gc.Modificado Then
        MsgBox MSGERR_NOGRABADO
        ImprimeTrans = False
        Exit Function
    End If
    
    'Solo por primera vez o cuando cambia la librería de impresión
    '  crea una instancia del objeto para la impresión
    crear = (objImp Is Nothing)
    If Not crear Then crear = (objImp.NombreDLL <> gc.GNTrans.ArchivoReporte)
    If crear Then
        Set objImp = Nothing
        Set objImp = CreateObject(gc.GNTrans.ArchivoReporte & ".PrintTrans")
    End If
    
    MensajeStatus "Está imprimiéndo ...", vbHourglass
    If Me.tag <> "Pago" Then
        objImp.PrintTrans gobjMain.EmpresaActual, True, 1, 0, "", 0, gc
    Else
        objImp.PrintTransLoteRuta gobjMain.EmpresaActual, True, 1, 0, txtEgreso.Text, "", 0, gc
    End If
    
    MensajeStatus "", 0
    ImprimeTrans = ""       'Sin problema
    Exit Function
ErrTrap:
    MensajeStatus "", 0
    Select Case Err.Number
    Case ERR_NOIMPRIME, ERR_NOIMPRIME2, ERR_NOIMPRIME3, ERR_NOHAYCODIGO
        ImprimeTrans = Err.Description
    Case Else
        ImprimeTrans = MSGERR_NOIMPRIME2
    End Select
    Exit Function
End Function

Public Sub InicioPago(Name As String)
    Dim i As Integer, s As String
    On Error GoTo ErrTrap
    Me.tag = Name
    Form_Resize
    Me.Show
    Me.ZOrder
    dtpFecha2.Visible = False
    dtpFecha2.TabStop = False
    Label2(1).Visible = False
    dtpFecha1.value = Date
    dtpFecha2.value = Date
    FraConFigEgreso.Visible = True
    ConfigCols
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("EgresosAutoLibImpPago")) > 0 Then
        s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("EgresosAutoLibImpPago")
        txtEgreso.Text = s
    End If

    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("EgresosAutoLibImpCheque")) > 0 Then
        s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("EgresosAutoLibImpCheque")
        txtCheque.Text = s
    End If
    
    fraCodTrans.Caption = "Trans. para Pago"
    fraCodTransVenta = "Trans. Compra "
    fraFecha.Caption = "Fecha Vencimiento "
    CargaTransPago
    CargaFormasPago
    'Carga la lista de Bancos
    fcbBanco.SetData gobjMain.EmpresaActual.ListaTSBanco(True, False)
    grd.Editable = flexEDKbdMouse
    grd.Refresh
    
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub CargaTransPago()
    Dim i As Long, v As Variant
    Dim s As String
    
    fcbTrans.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False)
    lstTrans.Clear
    v = gobjMain.GrupoActual.PermisoActual.ListaTransxTipoTransNew(False, "IV", "1")
    For i = LBound(v, 2) To UBound(v, 2)
        lstTrans.AddItem v(0, i)        '& " " & v(1, i)
    Next i
    
    'jeaa 25/09/206
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransParaEgresoAutomatico")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransParaEgresoAutomatico")
            RecuperaTrans "KeyT", lstTrans, s
        End If
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransEgresoParaEgresoAutomatico")) > 0 Then
            fcbTrans.KeyText = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransEgresoParaIngresoAutomatico")
        End If
        
        
End Sub

Private Sub CargaFormasPago()
    Dim i As Long, v As Variant
    Dim s As String
    
    lstForma.Clear
    v = gobjMain.EmpresaActual.ListaTSFormaCobroPago(False, True, False)
    For i = LBound(v, 2) To UBound(v, 2)
        lstForma.AddItem v(0, i)        '& " " & v(1, i)
    Next i
    
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("FormaParaEgresoAutomatico")) > 0 Then
        s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("FormaParaEgresoAutomatico")
        RecuperaTrans "KeyT", lstForma, s
    End If


   
End Sub

Private Sub PoneColorFilas()
    Dim i As Long, j As Long, Elemento As String, k As Long, l As Long
    With grd
        If .Rows <= .FixedRows Then Exit Sub
        For i = 1 To .Rows - 1
            For j = 1 To COL_RESULTADO
                  If Not (grd.IsSubtotal(i)) Then
                    If .ColData(j) = -1 Then
                      .Cell(flexcpBackColor, i, j, i, j) = &H80000018 'vbYellow
                    End If
                  End If
            Next j
        Next i
    End With
End Sub


Private Function EgresoAuto(ByVal bandVerificar As Boolean, BandTodo As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, x As Single, j As Integer, filaSubTotal As Long
    Dim gnc As GNComprobante, cambiado As Boolean, TransGen As String
    
    On Error GoTo ErrTrap
    
    
    'Verifica si está seleccionado una trans. de ingreso
    s = VerificaEgresoAutomatico
    If Len(s) > 0 Then
        'Si está seleccinada, confirma si está seguro
        s = "Está seleccionada una o más transacciones de egreso. " & vbCr & _
            "(" & s & ")" & vbCr & _
            "Generalmente no se hace Egresos Automáticos con transacciones de Egreso." & vbCr & vbCr
        s = s & "Confirma que desea proceder?" & vbCr & _
            "Aplaste 'Sí' unicamente cuando está seguro de lo que está haciendo."
        If MsgBox(s, vbYesNo + vbQuestion + vbDefaultButton2) <> vbYes Then Exit Function
    End If
    s = ""
    
    Set mColItems = Nothing     'Limpia lo anterior
    Set mColItems = New Collection
    
    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdAceptar.Enabled = False
    cmdBuscar.Enabled = False
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    
    For i = grd.FixedRows To grd.Rows - 1
        DoEvents
        If mCancelado Then
            MsgBox "El proceso fue cancelado.", vbInformation
            Exit For
        End If
        
        prg1.value = i
        grd.Row = i
        x = grd.CellTop                 'Para visualizar la celda actual
        
        If Not grd.IsSubtotal(i) Then
            If grd.ValueMatrix(i, COL_AUTO) = -1 Then
                tid = grd.ValueMatrix(i, COL_TID)
                grd.TextMatrix(i, COL_RESULTADO) = "Procesando  ..."
                grd.Refresh
                
                'Recupera la transaccion
                
                Set mobjGNCompOrigen = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
                If Not (mobjGNCompOrigen Is Nothing) Then
                    'Si la transacción es de Inventario y es Egreso/Transferencia
                    ' Y no está anulado
                    If (mobjGNCompOrigen.GNTrans.Modulo = "IV") And _
                       (mobjGNCompOrigen.Estado <> ESTADO_ANULADO) Then
    '                   (gnc.GNTrans.IVTipoTrans = "E" Or gnc.GNTrans.IVTipoTrans = "T") And _      '*** MAKOTO 06/sep/00 Eliminado
    
                        'Forzar recuperar todos los datos de transacción para que no se pierdan al grabar de nuveo
                        mobjGNCompOrigen.RecuperaDetalleTodo
                        'Recalcula costo de los items
                        If chkAgrupaProv.value <> vbChecked Then
                            If GrabarTransAutoEgreso(TransGen, grd.ValueMatrix(i, COL_VALORPAGO), grd.ValueMatrix(i, COL_ID)) Then
                                    'Graba la transacción
                                    grd.TextMatrix(i, COL_RESULTADO) = "OK.. Trans " & TransGen
                                    grd.TextMatrix(i, COL_TIDIN) = mobjGNCompAux.TransID
        
                            Else
                                    'Si no está cambiado no graba
                                    grd.TextMatrix(i, COL_RESULTADO) = "Falló Proceso"
                            End If
                        Else
                            For j = i To grd.Rows - 1
                                If grd.IsSubtotal(j) Then
                                    filaSubTotal = j
                                    j = grd.Rows - 1
                                End If
                            Next j
                            If GrabarTransAutoEgresoxProv(TransGen, i, filaSubTotal) Then
                                    'Graba la transacción
                                    For j = i To filaSubTotal - 1
                                        If grd.ValueMatrix(j, COL_AUTO) = -1 Then
                                            If j = filaSubTotal - 1 Then
                                                grd.TextMatrix(j, COL_RESULTADO) = "OK.. Trans " & TransGen
                                            Else
                                                grd.TextMatrix(j, COL_RESULTADO) = "Trans " & TransGen
                                            End If
                                            grd.TextMatrix(j, COL_TIDIN) = mobjGNCompAux.TransID
                                        End If
                                    Next j
                                    i = filaSubTotal
                            Else
                                    'Si no está cambiado no graba
                                    grd.TextMatrix(i, COL_RESULTADO) = "Falló Proceso"
                            End If
                        
                        End If
                    Else
                        'Si está anulado
                        If gnc.Estado = ESTADO_ANULADO Then
                            grd.TextMatrix(i, COL_RESULTADO) = "Anulado"
                        'Si no tiene nada que ver con recalculo de costo
                        Else
                            grd.TextMatrix(i, COL_RESULTADO) = "---"
                        End If
                    End If
                Else
                    grd.TextMatrix(i, COL_RESULTADO) = "No pudo recuperar la transación."
                End If
            End If
        End If
    Next i
    
    Screen.MousePointer = 0
'''    ReprocCosto = Not mCancelado
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    If i < grd.Rows And i >= grd.FixedRows Then
        grd.TextMatrix(i, COL_RESULTADO) = Err.Description
    End If
    DispErr
    prg1.value = prg1.min
salida:
    Set mColItems = Nothing         'Libera el objeto de coleccion
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdBuscar.Enabled = True
    cmdAceptar.Enabled = True
    prg1.value = prg1.min
    Exit Function
End Function


Private Function VerificaEgresoAutomatico() As String
    Dim i As Long, cod As String, gnt As GNTrans
    Dim s As String
    
    For i = 0 To lstTrans.ListCount - 1
        'Si está seleccionado
        If lstTrans.Selected(i) Then
            'Recupera el objeto GNTrans
            cod = lstTrans.List(i)
            Set gnt = gobjMain.EmpresaActual.RecuperaGNTrans(cod)
            'Si la transaccion es de ingreso, devuelve el codigo
            If gnt.IVTipoTrans = "E" Then s = s & cod & ", "
        End If
    Next i
    Set gnt = Nothing
    If Len(s) > 2 Then s = Left$(s, Len(s) - 2)     'Quita la ultima ", "
    VerificaEgresoAutomatico = s
End Function

Private Function GrabarTransAutoEgreso(ByRef trans As String, ValorPago As Currency, IdPago As Long) As Boolean
    Dim Imprime As Boolean, i As Long, ix As Long, orden1 As Integer, orden2 As Integer
    Dim pc As PCProvCli, Cadena As String, obser As String, Numcheque As Long
    Dim tsf As TSFormaCobroPago
    On Error GoTo ErrTrap
    GrabarTransAutoEgreso = True
    orden1 = 1
    orden2 = 1
    Numcheque = 0
    Set pc = mobjGNCompOrigen.Empresa.RecuperaPCProvCli(mobjGNCompOrigen.CodProveedorRef)
    If CreaComprobanteEgresoAuto(i) Then
        'Si es solo lectura, no hace nada
        If mobjGNCompAux.SoloVer Then
            MsgBox MSG_NODISPONE, vbInformation
            Exit Function
        End If
            'carga el pago al proveedor
            ix = mobjGNCompAux.AddPCKardex
            'PAGO
            mobjGNCompAux.PCKardex(ix).Debe = ValorPago
            mobjGNCompAux.PCKardex(ix).CodProvCli = mobjGNCompOrigen.CodProveedorRef
            mobjGNCompAux.PCKardex(ix).codforma = mobjGNCompOrigen.PCKardex(1).codforma
            mobjGNCompAux.PCKardex(ix).idAsignado = IdPago
            mobjGNCompAux.PCKardex(ix).NumLetra = mobjGNCompOrigen.CodTrans & " " & mobjGNCompOrigen.GNTrans.NumTransSiguiente
            mobjGNCompAux.PCKardex(ix).FechaEmision = mobjGNCompOrigen.FechaTrans
            mobjGNCompAux.PCKardex(ix).orden = orden1
            orden1 = orden1 + 1
            'carga el banco
            ix = mobjGNCompAux.AddTSKardex
            'PAGO
            mobjGNCompAux.TSKardex(ix).CodBanco = fcbBanco.KeyText
            mobjGNCompAux.TSKardex(ix).CodTipoDoc = "CH-E"
            mobjGNCompAux.TSKardex(ix).Haber = ValorPago
            mobjGNCompAux.TSKardex(ix).FechaEmision = Date
            mobjGNCompAux.TSKardex(ix).nombre = pc.nombre
            Numcheque = mobjGNCompAux.TSKardex(1).AsignaNumCheque(fcbBanco.KeyText)
            mobjGNCompAux.TSKardex(ix).numdoc = Numcheque

            mobjGNCompAux.FechaTrans = Date
            mobjGNCompAux.HoraTrans = Time()
             Cadena = "Cancelacion " & mobjGNCompOrigen.numDocRef & " de " & mobjGNCompOrigen.CodTrans & "-" & mobjGNCompOrigen.numtrans & " Proveedor: " & mobjGNCompOrigen.CodProveedorRef & " - " & pc.nombre & " / Banco: " & fcbBanco.KeyText
        If Len(Cadena) > 120 Then
            mobjGNCompAux.Descripcion = Mid$(Cadena, 1, 120)
        Else
            mobjGNCompAux.Descripcion = Cadena
        End If
            
        mobjGNCompAux.codUsuario = mobjGNCompOrigen.codUsuario
        mobjGNCompAux.IdResponsable = mobjGNCompOrigen.IdResponsable
        mobjGNCompAux.numDocRef = mobjGNCompOrigen.CodTrans & " " & mobjGNCompOrigen.numtrans
        mobjGNCompAux.CodMoneda = mobjGNCompOrigen.CodMoneda
        mobjGNCompAux.nombre = pc.nombre
        mobjGNCompAux.CodProveedorRef = mobjGNCompOrigen.PCKardex(1).CodProvCli
        mobjGNCompAux.idCentro = mobjGNCompOrigen.idCentro
        
        'Si es que algo está modificado
        If mobjGNCompAux.Modificado Then
            MensajeStatus MSG_GENERANDOASIENTO, vbHourglass
            MensajeStatus
        End If
        If mobjGNCompAux.GNTrans.AfectaSaldoPC And _
           mobjGNCompAux.GNTrans.TSVerificaTotalCuadrado Then
            'Verifica si está cuadrado el total de transacción y total de PCKardex.
            If Not TotalCuadrado Then Exit Function
        End If
        'Verificación de datos
        mobjGNCompAux.VerificaDatos
    
        'Verifica si está cuadrado el asiento
        If Not VerificaAsiento(mobjGNCompAux) Then Exit Function
    
        'Verifica si tiene detalle de banco
        If (mobjGNCompAux.CountTSKardex = 0) And _
            (mobjGNCompAux.CountTSKardexRet = 0) And _
            (mobjGNCompAux.CountPCKardex = 0) Then
            MsgBox "No existe ningún detalle.", vbInformation
            Exit Function
        End If
    
        MensajeStatus MSG_GRABANDO, vbHourglass
    
        'Manda a grabar
        '       Aquí ya no hacemos verificación de asiento por que ya está hecho en Control Asiento
        mobjGNCompAux.Grabar False, False
        
        
        MensajeStatus
        Me.Caption = mobjGNCompAux.CodTrans & " " & mobjGNCompAux.numtrans
        trans = mobjGNCompAux.CodTrans & " " & mobjGNCompAux.numtrans & " CH-E: " & Numcheque
        GrabarTransAutoEgreso = True
    End If
    Exit Function
ErrTrap:
    MensajeStatus
    Select Case Err.Number
    Case ERR_DESCUADRADO, ERR_INTEGRIDAD
        'Si es que el usuario seleccionó 'No' en el cuadro de dialogo,
        'No hace nada
    Case Else
        DispErr
    End Select
    GrabarTransAutoEgreso = False
    Exit Function
    
End Function


Private Function CreaComprobanteEgresoAuto(ByRef Num As Long) As Boolean
    Dim v As Currency, tsf As TSFormaCobroPago
    Dim i As Long
    CreaComprobanteEgresoAuto = False
    If Len(fcbTrans.KeyText) > 0 Then
        Set mobjGNCompAux = gobjMain.EmpresaActual.CreaGNComprobante(fcbTrans.KeyText)
    Else
        MsgBox "Falta seleccionar transaccion de Egreso"
        CreaComprobanteEgresoAuto = False
        Exit Function
    End If
    CreaComprobanteEgresoAuto = True

End Function

Private Function GrabarTransAutoEgresoxProv(ByRef trans As String, FilaIni As Long, FilaFin As Long) As Boolean
    Dim Imprime As Boolean, i As Long, ix As Long, orden1 As Integer, orden2 As Integer
    Dim pc As PCProvCli, Cadena As String, obser As String, Numcheque As Long
    Dim tsf As TSFormaCobroPago, ValorTotal  As Currency, x As Single
    On Error GoTo ErrTrap
    GrabarTransAutoEgresoxProv = True
    orden1 = 1
    orden2 = 1
    Numcheque = 0
    ValorTotal = 0
    Set pc = mobjGNCompOrigen.Empresa.RecuperaPCProvCli(mobjGNCompOrigen.CodProveedorRef)
    If CreaComprobanteEgresoAuto(i) Then
        'Si es solo lectura, no hace nada
        If mobjGNCompAux.SoloVer Then
            MsgBox MSG_NODISPONE, vbInformation
            Exit Function
        End If
            'carga el pago al proveedor
            For i = FilaIni To FilaFin - 1
                grd.Row = i
                x = grd.CellTop
                If grd.ValueMatrix(i, COL_AUTO) = -1 Then
                    ix = mobjGNCompAux.AddPCKardex
                    'PAGO
                    ValorTotal = ValorTotal + grd.ValueMatrix(i, COL_VALORPAGO)
                    mobjGNCompAux.PCKardex(ix).Debe = grd.ValueMatrix(i, COL_VALORPAGO)
                    mobjGNCompAux.PCKardex(ix).CodProvCli = mobjGNCompOrigen.CodProveedorRef
                    mobjGNCompAux.PCKardex(ix).codforma = mobjGNCompOrigen.PCKardex(1).codforma
                    mobjGNCompAux.PCKardex(ix).idAsignado = grd.ValueMatrix(i, COL_ID)
                    mobjGNCompAux.PCKardex(ix).NumLetra = mobjGNCompOrigen.CodTrans & " " & mobjGNCompOrigen.GNTrans.NumTransSiguiente
                    mobjGNCompAux.PCKardex(ix).FechaEmision = mobjGNCompOrigen.FechaTrans
                    mobjGNCompAux.PCKardex(ix).orden = orden1
                    orden1 = orden1 + 1
                End If
            Next i
            'carga el banco
            ix = mobjGNCompAux.AddTSKardex
            'PAGO
            mobjGNCompAux.TSKardex(ix).CodBanco = fcbBanco.KeyText
            mobjGNCompAux.TSKardex(ix).CodTipoDoc = "CH-E"
            mobjGNCompAux.TSKardex(ix).Haber = ValorTotal
            mobjGNCompAux.TSKardex(ix).FechaEmision = Date
            mobjGNCompAux.TSKardex(ix).nombre = pc.nombre
            Numcheque = mobjGNCompAux.TSKardex(1).AsignaNumCheque(fcbBanco.KeyText)
            mobjGNCompAux.TSKardex(ix).numdoc = Numcheque

            mobjGNCompAux.FechaTrans = Date
            mobjGNCompAux.HoraTrans = Time()
             Cadena = "Cancelacion " & mobjGNCompOrigen.numDocRef & " de " & mobjGNCompOrigen.CodTrans & "-" & mobjGNCompOrigen.numtrans & " Proveedor: " & mobjGNCompOrigen.CodProveedorRef & " - " & pc.nombre & " / Banco: " & fcbBanco.KeyText
        If Len(Cadena) > 120 Then
            mobjGNCompAux.Descripcion = Mid$(Cadena, 1, 120)
        Else
            mobjGNCompAux.Descripcion = Cadena
        End If
            
        mobjGNCompAux.codUsuario = mobjGNCompOrigen.codUsuario
        mobjGNCompAux.IdResponsable = mobjGNCompOrigen.IdResponsable
        mobjGNCompAux.numDocRef = mobjGNCompOrigen.CodTrans & " " & mobjGNCompOrigen.numtrans
        mobjGNCompAux.CodMoneda = mobjGNCompOrigen.CodMoneda
        mobjGNCompAux.nombre = pc.nombre
        mobjGNCompAux.CodProveedorRef = mobjGNCompOrigen.PCKardex(1).CodProvCli
        mobjGNCompAux.idCentro = mobjGNCompOrigen.idCentro
        
        'Si es que algo está modificado
        If mobjGNCompAux.Modificado Then
            MensajeStatus MSG_GENERANDOASIENTO, vbHourglass
            MensajeStatus
        End If
        If mobjGNCompAux.GNTrans.AfectaSaldoPC And _
           mobjGNCompAux.GNTrans.TSVerificaTotalCuadrado Then
            'Verifica si está cuadrado el total de transacción y total de PCKardex.
            If Not TotalCuadrado Then Exit Function
        End If
        'Verificación de datos
        mobjGNCompAux.VerificaDatos
    
        'Verifica si está cuadrado el asiento
        If Not VerificaAsiento(mobjGNCompAux) Then Exit Function
    
        'Verifica si tiene detalle de banco
        If (mobjGNCompAux.CountTSKardex = 0) And _
            (mobjGNCompAux.CountTSKardexRet = 0) And _
            (mobjGNCompAux.CountPCKardex = 0) Then
            MsgBox "No existe ningún detalle.", vbInformation
            Exit Function
        End If
    
        MensajeStatus MSG_GRABANDO, vbHourglass
    
        'Manda a grabar
        '       Aquí ya no hacemos verificación de asiento por que ya está hecho en Control Asiento
        mobjGNCompAux.Grabar False, False
        
        
        MensajeStatus
        Me.Caption = mobjGNCompAux.CodTrans & " " & mobjGNCompAux.numtrans
        trans = mobjGNCompAux.CodTrans & " " & mobjGNCompAux.numtrans & " CH-E: " & Numcheque
        GrabarTransAutoEgresoxProv = True
    End If
    Exit Function
ErrTrap:
    MensajeStatus
    Select Case Err.Number
    Case ERR_DESCUADRADO, ERR_INTEGRIDAD
        'Si es que el usuario seleccionó 'No' en el cuadro de dialogo,
        'No hace nada
    Case Else
        DispErr
    End Select
    GrabarTransAutoEgresoxProv = False
    Exit Function
    
End Function


Private Function ImprimirCheque() As Boolean
    Dim s As String, tid As Long, i As Long, x As Single, res As String, pos As Integer
    Dim gnc As GNComprobante, cambiado As Boolean, cntError As Long
    
    On Error GoTo ErrTrap

    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdBuscar.Enabled = False
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    
    For i = grd.FixedRows To grd.Rows - 1
        DoEvents
        If mCancelado Then
            MsgBox "El proceso fue cancelado."
            Exit For
        End If
        
        prg1.value = i
        grd.Row = i
        x = grd.CellTop                 'Para visualizar la celda actual
        pos = InStr(1, grd.TextMatrix(i, COL_RESULTADO), "OK")
        'Si es verificación, procesa todas las filas sino solo las que tengan "Asiento incorrecto."
        If pos <> 0 Then
        
            tid = grd.ValueMatrix(i, COL_TIDIN)
            grd.TextMatrix(i, COL_RESULTADO) = "Procesando ..."
            grd.Refresh
            
            'Recupera la transaccion
            Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
            If Not (gnc Is Nothing) Then
                'Si la transacción no está anulado
                If gnc.Estado <> ESTADO_ANULADO Then
    '                'Forzar recuperar todos los datos de transacción
    '                ' para que no se pierdan al grabar de nuveo
    '                gnc.RecuperaDetalleTodo
                
                    'Imprime la transaccion o asiento contable
                    res = ImprimeCheque(gnc, False)
                    If Len(res) = 0 Then
                        grd.TextMatrix(i, COL_RESULTADO) = "Enviado."
                    Else
                        grd.TextMatrix(i, COL_RESULTADO) = res
                        cntError = cntError + 1
                    End If
                                
                'Si la transaccion está anulado
                Else
                    grd.TextMatrix(i, COL_RESULTADO) = "Anulado."
                    cntError = cntError + 1
                End If
            Else
                grd.TextMatrix(i, COL_RESULTADO) = "No pudo recuperar la transación."
                cntError = cntError + 1
            End If
        End If
    Next i
    
    Screen.MousePointer = 0
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdImprimir.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    
    'Si algúna transaccion no se imprimió, avisa
    If cntError Then
        MsgBox "No se pudo imprimir " & cntError & " transacciones.", vbInformation
    End If
    
    ImprimirCheque = True
    Exit Function
ErrTrap:
    Screen.MousePointer = 0
    DispErr
    prg1.value = prg1.min
    Exit Function
End Function

Public Function ImprimeCheque(ByVal gc As GNComprobante, ByVal bandAsiento As Boolean) As String
    Dim crear As Boolean
    Static objImp As Object
    On Error GoTo ErrTrap

    'Si no tiene TransID quiere decir que no está grabada
    If (gc.TransID = 0) Or gc.Modificado Then
        MsgBox MSGERR_NOGRABADO
        ImprimeCheque = False
        Exit Function
    End If
    
    'Solo por primera vez o cuando cambia la librería de impresión
    '  crea una instancia del objeto para la impresión
    crear = (objImp Is Nothing)
    If Not crear Then crear = (objImp.NombreDLL <> gc.GNTrans.ArchivoReporte)
    If crear Then
        Set objImp = Nothing
        Set objImp = CreateObject(gc.GNTrans.ArchivoReporte & ".PrintTrans")
    End If
    
    MensajeStatus "Está imprimiéndo ...", vbHourglass
    objImp.PrintTransLoteRuta gobjMain.EmpresaActual, True, 1, 0, txtCheque.Text, "", 0, gc
        
    MensajeStatus "", 0
    ImprimeCheque = ""       'Sin problema
    Exit Function
ErrTrap:
    MensajeStatus "", 0
    Select Case Err.Number
    Case ERR_NOIMPRIME, ERR_NOIMPRIME2, ERR_NOIMPRIME3, ERR_NOHAYCODIGO
        ImprimeCheque = Err.Description
    Case Else
        ImprimeCheque = MSGERR_NOIMPRIME2
    End Select
    Exit Function
End Function


Private Function GrabarTransAutoxForma(ByRef trans As String, FilaIni As Long, FilaFin As Long) As Boolean
    Dim Imprime As Boolean, i As Long, ix As Long, orden1 As Integer, orden2 As Integer, j As Integer
    Dim pc As PCProvCli, Cadena As String, obser As String, codforma As String, saldo As Currency
    Dim tsf As TSFormaCobroPago
    Dim tid As Long, NumReg As Integer, x As Single
    
    On Error GoTo ErrTrap
    GrabarTransAutoxForma = True
    orden1 = 1
    orden2 = 1
    If CreaComprobanteIngresoAuto(i) Then
        'Si es solo lectura, no hace nada
        If mobjGNCompAux.SoloVer Then
            MsgBox MSG_NODISPONE, vbInformation
            Exit Function
        End If
        'carga la nueva deuda a los bancos de las tarjetas
        
        For j = FilaIni To FilaFin - 1
                grd.Row = j
                x = grd.CellTop                 'Para visualizar la celda actual

'            If grd.ValueMatrix(j, COL_NUMTRANS) = 1053 Then MsgBox "hola"
            
            tid = grd.ValueMatrix(j, COL_TID)
            'Recupera la transaccion
            Set mobjGNCompOrigen = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
            If j = FilaIni Then
                codforma = grd.TextMatrix(j, COL_CODFORMA)
            End If
            NumReg = 0
            For i = 1 To mobjGNCompOrigen.CountPCKardex
                If mobjGNCompOrigen.PCKardex(i).codforma = codforma Then
                    saldo = mobjGNCompOrigen.PCKardex(i).ObtieneSaldodePago(mobjGNCompOrigen.PCKardex(i).id)
                    If mobjGNCompOrigen.PCKardex(i).Debe = saldo Then
                        grd.TextMatrix(j, COL_RESULTADO) = "Procesando"
                        NumReg = NumReg + 1
                        Set tsf = mobjGNCompOrigen.Empresa.RecuperaTSFormaCobroPago(mobjGNCompOrigen.PCKardex(i).codforma)
                        If tsf.IngresoAutomatico Then
                            If tsf.DeudaMismoCliente Then
                                Set pc = mobjGNCompOrigen.Empresa.RecuperaPCProvCli(mobjGNCompOrigen.CodClienteRef)
                            Else
                                Set pc = mobjGNCompOrigen.Empresa.RecuperaPCProvCli(tsf.CodProvCli)
                            End If
                
                            
                            If tsf.IngresoAutomatico Then
                                ix = mobjGNCompAux.AddPCKardex
                                
                                mobjGNCompAux.PCKardex(ix).Debe = mobjGNCompOrigen.PCKardex(i).Debe
                                If tsf.DeudaMismoCliente Then
                                    mobjGNCompAux.PCKardex(ix).CodProvCli = mobjGNCompOrigen.CodClienteRef
                                Else
                                    mobjGNCompAux.PCKardex(ix).CodProvCli = tsf.CodProvCli
                                End If
                                mobjGNCompAux.PCKardex(ix).codforma = tsf.CodFormaTC
                                mobjGNCompAux.PCKardex(ix).NumLetra = mobjGNCompOrigen.CodTrans & " " & mobjGNCompOrigen.GNTrans.NumTransSiguiente & "-" & i
                                mobjGNCompAux.PCKardex(ix).FechaEmision = mobjGNCompOrigen.FechaTrans
                                mobjGNCompAux.PCKardex(ix).FechaVenci = mobjGNCompOrigen.FechaTrans
                                obser = "Por pago con: " & tsf.codforma & " de " & mobjGNCompOrigen.CodTrans & "-" & mobjGNCompOrigen.numtrans & " Cliente: " & mobjGNCompOrigen.CodClienteRef & " - " & mobjGNCompOrigen.nombre
                                mobjGNCompAux.PCKardex(ix).Observacion = IIf(Len(obser) > 80, Left(obser, 80), obser)
                                mobjGNCompAux.PCKardex(ix).CodVendedor = mobjGNCompOrigen.CodVendedor
                                mobjGNCompAux.PCKardex(ix).orden = orden1
                                orden1 = orden1 + 1
                                'PAGO
                                ix = mobjGNCompAux.AddPCKardex
                                
                        
                                
                                mobjGNCompAux.PCKardex(ix).Haber = mobjGNCompOrigen.PCKardex(i).Debe
                                mobjGNCompAux.PCKardex(ix).CodProvCli = mobjGNCompOrigen.CodClienteRef
                                mobjGNCompAux.PCKardex(ix).codforma = tsf.CodFormaTC
                                mobjGNCompAux.PCKardex(ix).idAsignado = mobjGNCompOrigen.PCKardex(i).id
                                mobjGNCompAux.PCKardex(ix).NumLetra = mobjGNCompOrigen.CodTrans & " " & mobjGNCompOrigen.GNTrans.NumTransSiguiente & "-" & i
                                mobjGNCompAux.PCKardex(ix).FechaEmision = mobjGNCompOrigen.FechaTrans
                                mobjGNCompAux.PCKardex(ix).CodVendedor = mobjGNCompOrigen.CodVendedor
                                mobjGNCompAux.PCKardex(ix).orden = orden1
                                orden1 = orden1 + 1
                            Else
                '                MsgBox "La forma de cobro " & tsf.NombreForma & " No esta Configurado para Ingreso Automático"
                '                    GrabarTransAuto = False
                '                Exit Function
                            End If
                        End If
                    Else
'                        MsgBox " valor que paga es mayor: " & saldo
                        grd.TextMatrix(j, COL_RESULTADO) = " Valor que paga es mayor: " & saldo & "- " & mobjGNCompOrigen.PCKardex(i).Debe
                    End If
                End If
            Next i
            If NumReg > 1 Then
                j = j + (NumReg - 1)
                grd.TextMatrix(j, COL_RESULTADO) = "Procesando"
            End If
            'i = i + 1
        Next j
        mobjGNCompAux.FechaTrans = mobjGNCompOrigen.FechaTrans
        mobjGNCompAux.HoraTrans = mobjGNCompOrigen.HoraTrans
'        If mobjGNCompOrigen.CountPCKardex > 1 Then
            Cadena = "Por pago con varias Clientes con " & codforma
 '       Else
  '           Cadena = "Por pago con :" & mobjGNCompOrigen.PCKardex(1).codforma & " de " & mobjGNCompOrigen.CodTrans & "-" & mobjGNCompOrigen.numtrans & " Cliente: " & mobjGNCompOrigen.CodClienteRef & " - " & mobjGNCompOrigen.nombre & " / Banco: " & pc.nombre
   '     End If
        If Len(Cadena) > 120 Then
            mobjGNCompAux.Descripcion = Mid$(Cadena, 1, 120)
        Else
            mobjGNCompAux.Descripcion = Cadena
        End If
            
        mobjGNCompAux.codUsuario = mobjGNCompOrigen.codUsuario
        mobjGNCompAux.IdResponsable = mobjGNCompOrigen.IdResponsable
        mobjGNCompAux.numDocRef = mobjGNCompOrigen.CodTrans & " " & mobjGNCompOrigen.numtrans
        mobjGNCompAux.idCentro = mobjGNCompOrigen.idCentro
        mobjGNCompAux.IdTransFuente = mobjGNCompOrigen.TransID
        mobjGNCompAux.CodMoneda = mobjGNCompOrigen.CodMoneda
'        If mobjGNCompOrigen.CountPCKardex > 1 Then
''            mobjGNCompAux.nombre = ""
''            mobjGNCompAux.CodClienteRef = ""
'            mobjGNCompAux.nombre = mobjGNCompOrigen.nombre
'            mobjGNCompAux.CodClienteRef = mobjGNCompOrigen.CodClienteRef
'        Else
'            mobjGNCompAux.nombre = mobjGNCompOrigen.nombre
'            mobjGNCompAux.CodClienteRef = mobjGNCompAux.PCKardex(i).codProvCli
'        End If
        mobjGNCompAux.CodVendedor = mobjGNCompOrigen.CodVendedor
        mobjGNCompAux.IdTransFuente = mobjGNCompAux.TransID
    
        'Si es que algo está modificado
        If mobjGNCompAux.Modificado Then
            MensajeStatus MSG_GENERANDOASIENTO, vbHourglass
            MensajeStatus
        End If
        If mobjGNCompAux.GNTrans.AfectaSaldoPC And _
           mobjGNCompAux.GNTrans.TSVerificaTotalCuadrado Then
            'Verifica si está cuadrado el total de transacción y total de PCKardex.
            If Not TotalCuadrado Then Exit Function
        End If
        Dim s As String
        For i = 1 To mobjGNCompAux.CountPCKardex
            s = mobjGNCompAux.PCKardex(i).idAsignado & " - " & mobjGNCompAux.PCKardex(i).Haber & " - " & mobjGNCompAux.PCKardex(i).Debe
            Debug.Print s
        Next i
        
        'Verificación de datos
        mobjGNCompAux.VerificaDatos
    
        'Verifica si está cuadrado el asiento
        If Not VerificaAsiento(mobjGNCompAux) Then Exit Function
    
        'Verifica si tiene detalle de banco
        If (mobjGNCompAux.CountTSKardex = 0) And _
            (mobjGNCompAux.CountTSKardexRet = 0) And _
            (mobjGNCompAux.CountPCKardex = 0) Then
            MsgBox "No existe ningún detalle.", vbInformation
            Exit Function
        End If
    
        MensajeStatus MSG_GRABANDO, vbHourglass
    
        'Manda a grabar
        '       Aquí ya no hacemos verificación de asiento por que ya está hecho en Control Asiento
        mobjGNCompAux.Grabar False, False
        
        
'        grd.TextMatrix(grd.Row, COL_TIDIN) = mobjGNCompAux.TransID
'        grd.TextMatrix(grd.Row, COL_NUMTRANSIN) = mobjGNCompAux.numtrans
        
        '***  Oliver 26/12/2002
        'Agregado para el control ded Impresion Configurado en la Transaccion
        
        
        MensajeStatus
    '    Me.caption = "Transacción " & mobjGNCompAux.codTrans & " " & mobjGNCompAux.NumTrans
        Me.Caption = mobjGNCompAux.CodTrans & " " & mobjGNCompAux.numtrans
        trans = mobjGNCompAux.CodTrans & " " & mobjGNCompAux.numtrans
        GrabarTransAutoxForma = True
    Else
        GrabarTransAutoxForma = False
    End If
    Exit Function
ErrTrap:
    MensajeStatus
    Select Case Err.Number
    Case ERR_DESCUADRADO, ERR_INTEGRIDAD
        'Si es que el usuario seleccionó 'No' en el cuadro de dialogo,
        'No hace nada
    Case Else
        DispErr
    End Select
    GrabarTransAutoxForma = False
    Exit Function
    
End Function

Private Function GrabarTransAutoxTrans(ByRef trans As String, FilaIni As Long, FilaFin As Long) As Boolean
    Dim Imprime As Boolean, i As Long, ix As Long, orden1 As Integer, orden2 As Integer, j As Integer
    Dim pc As PCProvCli, Cadena As String, obser As String, codforma As String, Num As Long
    Dim tsf As TSFormaCobroPago, x As Single, k As Long, sUMATotal  As Currency
    Dim tid As Long
    On Error GoTo ErrTrap
    GrabarTransAutoxTrans = True
    orden1 = 1
    orden2 = 1
    sUMATotal = 0
    If CreaComprobanteIngresoAuto(i) Then
        'Si es solo lectura, no hace nada
        If mobjGNCompAux.SoloVer Then
            MsgBox MSG_NODISPONE, vbInformation
            Exit Function
        End If
        'carga la nueva deuda a los bancos de las tarjetas
        i = 1
        For j = FilaIni To FilaFin - 1
            grd.Row = j
            x = grd.CellTop
            tid = grd.ValueMatrix(j, COL_TID)
            'Recupera la transaccion
            Set mobjGNCompOrigen = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
            If j = FilaIni Then
                codforma = grd.TextMatrix(j, COL_TID)
            End If
            For k = 1 To mobjGNCompOrigen.CountPRLibroDetalle
'                If mobjGNCompOrigen.PCKardex(k).codforma = grd.TextMatrix(j, COL_CODFORMA) Then
 '                   i = k
                    k = mobjGNCompOrigen.CountPRLibroDetalle
                'End If
             Next k
            
'                If mobjGNCompOrigen.TransID = codforma Then
 '                   Set tsf = mobjGNCompOrigen.Empresa.RecuperaTSFormaCobroPago(grd.TextMatrix(j, COL_CODFORMA))
  '                  If tsf.IngresoAutomatico Then
   '                     If tsf.DeudaMismoCliente Then
    '                        Set pc = mobjGNCompOrigen.Empresa.RecuperaPCProvCli(mobjGNCompOrigen.CodClienteRef)
     '                   Else
      '                      Set pc = mobjGNCompOrigen.Empresa.RecuperaPCProvCli(tsf.CodProvCli)
       '                 End If
            
                        
        '                If tsf.IngresoAutomatico Then
                            ix = mobjGNCompAux.AddPRLibroDetalle
                            mobjGNCompAux.PRLibroDetalle(ix).Debe = mobjGNCompOrigen.PRLibroDetalle(ix).Debe
                            mobjGNCompAux.PRLibroDetalle(ix).Haber = mobjGNCompOrigen.PRLibroDetalle(ix).Haber
                            mobjGNCompAux.PRLibroDetalle(ix).codcuenta = mobjGNCompOrigen.PRLibroDetalle(ix).codcuenta
                            mobjGNCompAux.PRLibroDetalle(ix).orden = mobjGNCompOrigen.PRLibroDetalle(ix).orden
                            mobjGNCompAux.PRLibroDetalle(ix).BandIntegridad = mobjGNCompOrigen.PRLibroDetalle(ix).BandIntegridad
                            mobjGNCompAux.PRLibroDetalle(ix).Descripcion = mobjGNCompOrigen.PRLibroDetalle(ix).Descripcion
                            If mobjGNCompAux.PRLibroDetalle(ix).codcuenta = "9999.01" Then
                                sUMATotal = sUMATotal + mobjGNCompAux.PRLibroDetalle(ix).Haber
                            End If
                i = i + 1
            'Next i
        Next j
        
                                ix = mobjGNCompAux.AddIVKardex
                        mobjGNCompAux.IVKardex(ix).CodInventario = "CONTA12"
                        mobjGNCompAux.IVKardex(ix).cantidad = 1
                        mobjGNCompAux.IVKardex(ix).orden = 1
                        mobjGNCompAux.IVKardex(ix).CostoTotal = sUMATotal

        mobjGNCompAux.FechaTrans = DateAdd("n", -1, mobjGNCompOrigen.FechaTrans)
        mobjGNCompAux.HoraTrans = mobjGNCompOrigen.HoraTrans
        'If Len(cadena) > 120 Then
            mobjGNCompAux.Descripcion = Mid$(mobjGNCompOrigen.Descripcion, 1, 120)
        'Else
         '   mobjGNCompAux.Descripcion = cadena
        'End If
            
        mobjGNCompAux.codUsuario = mobjGNCompOrigen.codUsuario
        mobjGNCompAux.IdResponsable = mobjGNCompOrigen.IdResponsable
        mobjGNCompAux.numDocRef = mobjGNCompOrigen.CodTrans & " " & mobjGNCompOrigen.numtrans
        mobjGNCompAux.idCentro = mobjGNCompOrigen.idCentro
        mobjGNCompAux.IdTransFuente = mobjGNCompOrigen.TransID
        mobjGNCompAux.CodMoneda = mobjGNCompOrigen.CodMoneda
        mobjGNCompAux.CodVendedor = mobjGNCompOrigen.CodVendedor
        mobjGNCompAux.IdTransFuente = mobjGNCompAux.TransID
    
        'Si es que algo está modificado
        If mobjGNCompAux.Modificado Then
            MensajeStatus MSG_GENERANDOASIENTO, vbHourglass
            MensajeStatus
        End If
        If mobjGNCompAux.GNTrans.AfectaSaldoPC And _
           mobjGNCompAux.GNTrans.TSVerificaTotalCuadrado Then
            'Verifica si está cuadrado el total de transacción y total de PCKardex.
            If Not TotalCuadrado Then Exit Function
        End If
        'Verificación de datos
        mobjGNCompAux.VerificaDatos
    
        'Verifica si está cuadrado el asiento
        If Not VerificaAsiento(mobjGNCompAux) Then Exit Function
        'If Not VerificaAsientoP(mobjGNCompAux) Then Exit Function
    
        'Verifica si tiene detalle de banco
'        If (mobjGNCompAux.CountTSKardex = 0) And _
'            (mobjGNCompAux.CountTSKardexRet = 0) And _
'            (mobjGNCompAux.CountPCKardex = 0) Then
'            MsgBox "No existe ningún detalle.", vbInformation
'            Exit Function
'        End If
    
        MensajeStatus MSG_GRABANDO, vbHourglass
    
        'Manda a grabar
        '       Aquí ya no hacemos verificación de asiento por que ya está hecho en Control Asiento
        mobjGNCompAux.Grabar False, False
        
        
'        grd.TextMatrix(grd.Row, COL_TIDIN) = mobjGNCompAux.TransID
'        grd.TextMatrix(grd.Row, COL_NUMTRANSIN) = mobjGNCompAux.numtrans
        
        '***  Oliver 26/12/2002
        'Agregado para el control ded Impresion Configurado en la Transaccion
        
        
        MensajeStatus
    '    Me.caption = "Transacción " & mobjGNCompAux.codTrans & " " & mobjGNCompAux.NumTrans
        Me.Caption = mobjGNCompAux.CodTrans & " " & mobjGNCompAux.numtrans
        trans = mobjGNCompAux.CodTrans & " " & mobjGNCompAux.numtrans
        GrabarTransAutoxTrans = True
    Else
        GrabarTransAutoxTrans = False
    End If
    Exit Function
ErrTrap:
    MensajeStatus
    Select Case Err.Number
    Case ERR_DESCUADRADO, ERR_INTEGRIDAD
        'Si es que el usuario seleccionó 'No' en el cuadro de dialogo,
        'No hace nada
    Case Else
        DispErr
    End Select
    GrabarTransAutoxTrans = False
    Exit Function
    
End Function


Private Function VerificaSaldoxPagar(TransID As Long) As Currency
Dim sql As String, rs As Recordset
End Function

Public Sub InicioCruceTarjetas(Name As String)
    Dim i As Integer
    On Error GoTo ErrTrap
    Me.tag = Name
    PicForma.Visible = True
    lblBanco.Visible = False
    fcbBanco.Visible = False
    fcbBanco.TabStop = False
    lblnumche.Visible = False
    lblSal.Visible = False
    lblsaldoBanco.Visible = False
    lblNumCheque.Visible = False
    cmdImprimiCH.Visible = False
    cmdImprimiCH.TabStop = False
    chkAgrupaProv.Caption = "Agrupar por Banco"
    ConfigCols
    Me.Show
    Me.ZOrder
    dtpFecha1.value = Date
    dtpFecha2.value = Date
    CargaTrans
    CargaFormasTarjetas
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub


Private Sub CargaFormasTarjetas()
    Dim i As Long, v As Variant
    Dim s As String

    lstForma.Clear
    v = gobjMain.EmpresaActual.ListaTSFormaCobroPago(True, True, False)
    For i = LBound(v, 2) To UBound(v, 2)
        lstForma.AddItem v(0, i)        '& " " & v(1, i)
    Next i
    
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("FormaParaCruceTarjetas")) > 0 Then
        s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("FormaParaCruceTarjetas")
        RecuperaTrans "KeyT", lstForma, s
    End If

End Sub

Private Function CruceTarjetas(ByVal bandVerificar As Boolean, BandTodo As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, x As Single, j As Integer, filaSubTotal As Long
    Dim gnc As GNComprobante, cambiado As Boolean, TransGen As String
    
    On Error GoTo ErrTrap
    
    'Si no es solo verificacion, confirma
    If Not bandVerificar Then
        'Confirma la actualización
        s = "Este proceso creará un Cruce de Tarjetas  de la transacción seleccionada." & vbCr & vbCr
        s = s & "Está seguro que desea proceder?"
        If MsgBox(s, vbYesNo + vbQuestion) <> vbYes Then Exit Function
    End If
    
    'Verifica si está seleccionado una trans. de ingreso
    s = VerificaIngresoAutomatico
    If Len(s) > 0 Then
        'Si está seleccinada, confirma si está seguro
        s = "Está seleccionada una o más transacciones de ingreso. " & vbCr & _
            "(" & s & ")" & vbCr & _
            "Generalmente no se hace Cruce de Tarjeta con transacciones de ingreso." & vbCr & vbCr
        s = s & "Confirma que desea proceder?" & vbCr & _
            "Aplaste 'Sí' unicamente cuando está seguro de lo que está haciendo."
        If MsgBox(s, vbYesNo + vbQuestion + vbDefaultButton2) <> vbYes Then Exit Function
    End If
    s = ""
    
    Set mColItems = Nothing     'Limpia lo anterior
    Set mColItems = New Collection
    
    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdAceptar.Enabled = False
    cmdBuscar.Enabled = False
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    
    For i = grd.FixedRows To grd.Rows - 1
        DoEvents
        If mCancelado Then
            MsgBox "El proceso fue cancelado.", vbInformation
            Exit For
        End If
        
        prg1.value = i
        grd.Row = i
        x = grd.CellTop                 'Para visualizar la celda actual
        
        If Not grd.IsSubtotal(i) Then
        'Si es verificación procesa todas las filas sino solo las que tengan "Costo Incorrecto"
        
            tid = grd.ValueMatrix(i, COL_TID)
            grd.TextMatrix(i, COL_RESULTADO) = "Procesando  ..."
            grd.Refresh
            
            'Recupera la transaccion
            
            Set mobjGNCompOrigen = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
            If Not (mobjGNCompOrigen Is Nothing) Then
                'Si la transacción es de Inventario y es Egreso/Transferencia
                ' Y no está anulado
                If (mobjGNCompOrigen.GNTrans.Modulo = "IV") And _
                   (mobjGNCompOrigen.Estado <> ESTADO_ANULADO) Then
                    'Forzar recuperar todos los datos de transacción para que no se pierdan al grabar de nuveo
                    mobjGNCompOrigen.RecuperaDetalleTodo
                    'Recalcula costo de los items
                    If chkAgrupaProv.value <> vbChecked Then
                        For j = i To grd.Rows - 1
                            If grd.IsSubtotal(j) Then
                                filaSubTotal = j
                                j = grd.Rows - 1
                            End If
                        Next j
                        If GrabarTransCrucexTrans(TransGen, i, filaSubTotal) Then
                            'Graba la transacción
                            For j = i To filaSubTotal - 1
                                    If j = filaSubTotal - 1 Then
                                        grd.TextMatrix(j, COL_RESULTADO) = "OK.. Trans " & TransGen
                                    Else
                                        grd.TextMatrix(j, COL_RESULTADO) = "Trans " & TransGen
                                    End If
                                    grd.TextMatrix(j, COL_TIDIN) = mobjGNCompAux.TransID
                            Next j
                            i = filaSubTotal
                        Else
                            'Si no está cambiado no graba
                            grd.TextMatrix(i, COL_RESULTADO) = "Falló Proceso"
                        End If
                    Else
                        For j = i To grd.Rows - 1
                            If grd.IsSubtotal(j) Then
                                filaSubTotal = j
                                j = grd.Rows - 1
                            End If
                        Next j
                        If GrabarTransCruceAutoxForma(TransGen, i, filaSubTotal) Then
                            'Graba la transacción
                            For j = i To filaSubTotal - 1
                                    If j = filaSubTotal - 1 Then
                                        grd.TextMatrix(j, COL_RESULTADO) = "OK.. Trans " & TransGen
                                    Else
                                        If grd.TextMatrix(j, COL_RESULTADO) = "Procesando" Then
                                            grd.TextMatrix(j, COL_RESULTADO) = "Trans " & TransGen
                                        End If
                                    End If
                                    grd.TextMatrix(j, COL_TIDIN) = mobjGNCompAux.TransID
                            Next j
                            i = filaSubTotal
                        Else
                            'Si no está cambiado no graba
                            grd.TextMatrix(i, COL_RESULTADO) = "Falló Proceso"
                            i = filaSubTotal
                        End If
                    End If
                Else
                    'Si está anulado
                    If gnc.Estado = ESTADO_ANULADO Then
                        grd.TextMatrix(i, COL_RESULTADO) = "Anulado"
                    'Si no tiene nada que ver con recalculo de costo
                    Else
                        grd.TextMatrix(i, COL_RESULTADO) = "---"
                    End If
                End If
            Else
                grd.TextMatrix(i, COL_RESULTADO) = "No pudo recuperar la transación."
            End If
        End If
    Next i
    Screen.MousePointer = 0
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    If i < grd.Rows And i >= grd.FixedRows Then
        grd.TextMatrix(i, COL_RESULTADO) = Err.Description
    End If
    DispErr
    prg1.value = prg1.min
salida:
    Set mColItems = Nothing         'Libera el objeto de coleccion
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdBuscar.Enabled = True
    cmdAceptar.Enabled = True
    prg1.value = prg1.min
    Exit Function
End Function


Private Function GrabarTransCrucexTrans(ByRef trans As String, FilaIni As Long, FilaFin As Long) As Boolean
    Dim Imprime As Boolean, i As Long, ix As Long, orden1 As Integer, orden2 As Integer, j As Integer
    Dim pc As PCProvCli, Cadena As String, obser As String, codforma As String, Num As Long
    Dim tsf As TSFormaCobroPago, x As Single, k As Long
    Dim tid As Long, s As String
    On Error GoTo ErrTrap
    GrabarTransCrucexTrans = True
    orden1 = 1
    orden2 = 1
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("FormaParaCruceTarjetas")) > 0 Then
        s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("FormaParaCruceTarjetas")
    End If

    If CreaComprobanteIngresoAuto(i) Then
        'Si es solo lectura, no hace nada
        If mobjGNCompAux.SoloVer Then
            MsgBox MSG_NODISPONE, vbInformation
            Exit Function
        End If
        'carga la nueva deuda a los bancos de las tarjetas
        i = 1
        For j = FilaIni To FilaFin - 1
        For i = 1 To mobjGNCompOrigen.CountPCKardex
            grd.Row = j
            x = grd.CellTop
            tid = grd.ValueMatrix(j, COL_TID)
            'Recupera la transaccion
            Set mobjGNCompOrigen = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
            If j = FilaIni Then
                codforma = grd.TextMatrix(j, COL_TID)
            End If
            If InStr(1, s, mobjGNCompOrigen.PCKardex(i).codforma) > 0 Then
            If mobjGNCompOrigen.PCKardex(i).id = grd.TextMatrix(j, COL_ID) Then
                Set pc = mobjGNCompOrigen.Empresa.RecuperaPCProvCli(grd.TextMatrix(j, COL_CODPROVCLI))
                ix = mobjGNCompAux.AddPCKardex
                mobjGNCompAux.PCKardex(ix).Debe = mobjGNCompOrigen.PCKardex(i).Debe
                mobjGNCompAux.PCKardex(ix).CodProvCli = grd.TextMatrix(j, COL_CODPROVCLI)
                mobjGNCompAux.PCKardex(ix).codforma = fcbFormaCobro.KeyText
                mobjGNCompAux.PCKardex(ix).NumLetra = mobjGNCompOrigen.CodTrans & " " & mobjGNCompOrigen.numtrans
                mobjGNCompAux.PCKardex(ix).FechaEmision = mobjGNCompOrigen.FechaTrans
                mobjGNCompAux.PCKardex(ix).FechaVenci = mobjGNCompOrigen.FechaTrans
                obser = "Por pago con: " & grd.TextMatrix(j, COL_CODTARJETA) & " de " & mobjGNCompOrigen.CodTrans & "-" & mobjGNCompOrigen.numtrans & " Cliente: " & mobjGNCompOrigen.CodClienteRef & " - " & mobjGNCompOrigen.nombre
                mobjGNCompAux.PCKardex(ix).Observacion = IIf(Len(obser) > 80, Left(obser, 80), obser)
                mobjGNCompAux.PCKardex(ix).CodVendedor = mobjGNCompOrigen.CodVendedor
                mobjGNCompAux.PCKardex(ix).orden = orden1
                orden1 = orden1 + 1
                'PAGO
                ix = mobjGNCompAux.AddPCKardex
                mobjGNCompAux.PCKardex(ix).Haber = mobjGNCompOrigen.PCKardex(i).Debe
                mobjGNCompAux.PCKardex(ix).CodProvCli = mobjGNCompOrigen.CodClienteRef
                mobjGNCompAux.PCKardex(ix).codforma = fcbFormaCobro.KeyText
                mobjGNCompAux.PCKardex(ix).idAsignado = mobjGNCompOrigen.PCKardex(i).id
                mobjGNCompAux.PCKardex(ix).NumLetra = mobjGNCompOrigen.CodTrans & " " & mobjGNCompOrigen.GNTrans.NumTransSiguiente & "-" & i
                mobjGNCompAux.PCKardex(ix).FechaEmision = mobjGNCompOrigen.FechaTrans
                mobjGNCompAux.PCKardex(ix).CodVendedor = mobjGNCompOrigen.CodVendedor
                mobjGNCompAux.PCKardex(ix).orden = orden1
                orden1 = orden1 + 1
            End If
            End If
            Next i
            'i = i + 1
        Next j
        mobjGNCompAux.FechaTrans = mobjGNCompOrigen.FechaTrans
        mobjGNCompAux.HoraTrans = mobjGNCompOrigen.HoraTrans
        If mobjGNCompOrigen.CountPCKardex > 1 Then
            Cadena = "Por pago con varias Clientes con " & codforma
        Else
             Cadena = "Por pago con :" & mobjGNCompOrigen.PCKardex(1).codforma & " de " & mobjGNCompOrigen.CodTrans & "-" & mobjGNCompOrigen.numtrans & " Cliente: " & mobjGNCompOrigen.CodClienteRef & " - " & mobjGNCompOrigen.nombre & " / Banco: " & pc.nombre
        End If
        If Len(Cadena) > 120 Then
            mobjGNCompAux.Descripcion = Mid$(Cadena, 1, 120)
        Else
            mobjGNCompAux.Descripcion = Cadena
        End If
            
        mobjGNCompAux.codUsuario = mobjGNCompOrigen.codUsuario
        mobjGNCompAux.IdResponsable = mobjGNCompOrigen.IdResponsable
        mobjGNCompAux.numDocRef = mobjGNCompOrigen.CodTrans & " " & mobjGNCompOrigen.numtrans
        mobjGNCompAux.idCentro = mobjGNCompOrigen.idCentro
        mobjGNCompAux.IdTransFuente = mobjGNCompOrigen.TransID
        mobjGNCompAux.CodMoneda = mobjGNCompOrigen.CodMoneda
        If mobjGNCompOrigen.CountPCKardex > 1 Then
            mobjGNCompAux.nombre = mobjGNCompOrigen.nombre
            mobjGNCompAux.CodClienteRef = mobjGNCompOrigen.CodClienteRef
        Else
            mobjGNCompAux.nombre = mobjGNCompOrigen.nombre
            mobjGNCompAux.CodClienteRef = mobjGNCompAux.PCKardex(i).CodProvCli
        End If
        mobjGNCompAux.CodVendedor = mobjGNCompOrigen.CodVendedor
        mobjGNCompAux.IdTransFuente = mobjGNCompAux.TransID
    
        'Si es que algo está modificado
        If mobjGNCompAux.Modificado Then
            MensajeStatus MSG_GENERANDOASIENTO, vbHourglass
            MensajeStatus
        End If
        If mobjGNCompAux.GNTrans.AfectaSaldoPC And _
           mobjGNCompAux.GNTrans.TSVerificaTotalCuadrado Then
            'Verifica si está cuadrado el total de transacción y total de PCKardex.
            If Not TotalCuadrado Then Exit Function
        End If
        'Verificación de datos
        mobjGNCompAux.VerificaDatos
    
        'Verifica si está cuadrado el asiento
        If Not VerificaAsiento(mobjGNCompAux) Then Exit Function
    
        'Verifica si tiene detalle de banco
        If (mobjGNCompAux.CountTSKardex = 0) And _
            (mobjGNCompAux.CountTSKardexRet = 0) And _
            (mobjGNCompAux.CountPCKardex = 0) Then
            MsgBox "No existe ningún detalle.", vbInformation
            Exit Function
        End If
    
        MensajeStatus MSG_GRABANDO, vbHourglass
    
        'Manda a grabar
        '       Aquí ya no hacemos verificación de asiento por que ya está hecho en Control Asiento
        mobjGNCompAux.Grabar False, False
        
        
        '***  Oliver 26/12/2002
        'Agregado para el control ded Impresion Configurado en la Transaccion
        
        
        MensajeStatus
    '    Me.caption = "Transacción " & mobjGNCompAux.codTrans & " " & mobjGNCompAux.NumTrans
        Me.Caption = mobjGNCompAux.CodTrans & " " & mobjGNCompAux.numtrans
        trans = mobjGNCompAux.CodTrans & " " & mobjGNCompAux.numtrans
        GrabarTransCrucexTrans = True
    Else
        GrabarTransCrucexTrans = False
    End If
    Exit Function
ErrTrap:
    MensajeStatus
    Select Case Err.Number
    Case ERR_DESCUADRADO, ERR_INTEGRIDAD
        'Si es que el usuario seleccionó 'No' en el cuadro de dialogo,
        'No hace nada
    Case Else
        DispErr
    End Select
    GrabarTransCrucexTrans = False
    Exit Function
    
End Function


Private Function GrabarTransCruceAutoxForma(ByRef trans As String, FilaIni As Long, FilaFin As Long) As Boolean
    Dim Imprime As Boolean, i As Long, ix As Long, orden1 As Integer, orden2 As Integer, j As Integer
    Dim pc As PCProvCli, Cadena As String, obser As String, codforma As String, saldo As Currency
    Dim tsf As TSFormaCobroPago
    Dim tid As Long, NumReg As Integer, x As Single
    
    On Error GoTo ErrTrap
    GrabarTransCruceAutoxForma = True
    orden1 = 1
    orden2 = 1
    If CreaComprobanteIngresoAuto(i) Then
        'Si es solo lectura, no hace nada
        If mobjGNCompAux.SoloVer Then
            MsgBox MSG_NODISPONE, vbInformation
            Exit Function
        End If
        'carga la nueva deuda a los bancos de las tarjetas
        
        For j = FilaIni To FilaFin - 1
            grd.Row = j
            x = grd.CellTop                 'Para visualizar la celda actual
            tid = grd.ValueMatrix(j, COL_TID)
            'Recupera la transaccion
            Set mobjGNCompOrigen = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
            If j = FilaIni Then
                codforma = grd.TextMatrix(j, COL_NOMBANCO)
            End If
            NumReg = 0
            For i = 1 To mobjGNCompOrigen.CountPCKardex
                If mobjGNCompOrigen.PCKardex(i).id = grd.TextMatrix(j, COL_ID) Then
                    grd.TextMatrix(j, COL_RESULTADO) = "Procesando"
                    NumReg = NumReg + 1
                    Set pc = mobjGNCompOrigen.Empresa.RecuperaPCProvCli(grd.TextMatrix(j, COL_CODPROVCLI))
                    ix = mobjGNCompAux.AddPCKardex
                    mobjGNCompAux.PCKardex(ix).Debe = mobjGNCompOrigen.PCKardex(i).Debe
                    mobjGNCompAux.PCKardex(ix).CodProvCli = grd.TextMatrix(j, COL_CODPROVCLI)
                    mobjGNCompAux.PCKardex(ix).codforma = fcbFormaCobro.KeyText
                    mobjGNCompAux.PCKardex(ix).NumLetra = mobjGNCompOrigen.CodTrans & " " & mobjGNCompOrigen.numtrans
                    mobjGNCompAux.PCKardex(ix).FechaEmision = mobjGNCompOrigen.FechaTrans
                    mobjGNCompAux.PCKardex(ix).FechaVenci = mobjGNCompOrigen.FechaTrans
                    obser = "Por pago con: " & grd.TextMatrix(j, COL_CODTARJETA) & " de " & mobjGNCompOrigen.CodTrans & "-" & mobjGNCompOrigen.numtrans & " Cliente: " & mobjGNCompOrigen.CodClienteRef & " - " & mobjGNCompOrigen.nombre
                    mobjGNCompAux.PCKardex(ix).Observacion = IIf(Len(obser) > 80, Left(obser, 80), obser)
                    mobjGNCompAux.PCKardex(ix).CodVendedor = mobjGNCompOrigen.CodVendedor
                    mobjGNCompAux.PCKardex(ix).orden = orden1
                    orden1 = orden1 + 1
                    'PAGO
                    ix = mobjGNCompAux.AddPCKardex
                    mobjGNCompAux.PCKardex(ix).Haber = mobjGNCompOrigen.PCKardex(i).Debe
                    mobjGNCompAux.PCKardex(ix).CodProvCli = mobjGNCompOrigen.CodClienteRef
                    mobjGNCompAux.PCKardex(ix).codforma = fcbFormaCobro.KeyText
                    mobjGNCompAux.PCKardex(ix).idAsignado = mobjGNCompOrigen.PCKardex(i).id
                    mobjGNCompAux.PCKardex(ix).NumLetra = mobjGNCompOrigen.CodTrans & " " & mobjGNCompOrigen.GNTrans.NumTransSiguiente & "-" & i
                    mobjGNCompAux.PCKardex(ix).FechaEmision = mobjGNCompOrigen.FechaTrans
                    mobjGNCompAux.PCKardex(ix).CodVendedor = mobjGNCompOrigen.CodVendedor
                    mobjGNCompAux.PCKardex(ix).orden = orden1
                    orden1 = orden1 + 1
                End If
            Next i
            If NumReg > 1 Then
                j = j + (NumReg - 1)
                grd.TextMatrix(j, COL_RESULTADO) = "Procesando"
            End If
        Next j
        mobjGNCompAux.FechaTrans = mobjGNCompOrigen.FechaTrans
        mobjGNCompAux.HoraTrans = mobjGNCompOrigen.HoraTrans
            Cadena = "Por pago con varias Clientes con " & codforma
        If Len(Cadena) > 120 Then
            mobjGNCompAux.Descripcion = Mid$(Cadena, 1, 120)
        Else
            mobjGNCompAux.Descripcion = Cadena
        End If
            
        mobjGNCompAux.codUsuario = mobjGNCompOrigen.codUsuario
        mobjGNCompAux.IdResponsable = mobjGNCompOrigen.IdResponsable
        mobjGNCompAux.numDocRef = mobjGNCompOrigen.CodTrans & " " & mobjGNCompOrigen.numtrans
        mobjGNCompAux.idCentro = mobjGNCompOrigen.idCentro
        mobjGNCompAux.IdTransFuente = mobjGNCompOrigen.TransID
        mobjGNCompAux.CodMoneda = mobjGNCompOrigen.CodMoneda
        mobjGNCompAux.CodVendedor = mobjGNCompOrigen.CodVendedor
        mobjGNCompAux.IdTransFuente = mobjGNCompAux.TransID
    
        'Si es que algo está modificado
        If mobjGNCompAux.Modificado Then
            MensajeStatus MSG_GENERANDOASIENTO, vbHourglass
            MensajeStatus
        End If
        If mobjGNCompAux.GNTrans.AfectaSaldoPC And _
           mobjGNCompAux.GNTrans.TSVerificaTotalCuadrado Then
            'Verifica si está cuadrado el total de transacción y total de PCKardex.
            If Not TotalCuadrado Then Exit Function
        End If
        Dim s As String
        For i = 1 To mobjGNCompAux.CountPCKardex
            s = mobjGNCompAux.PCKardex(i).idAsignado & " - " & mobjGNCompAux.PCKardex(i).Haber & " - " & mobjGNCompAux.PCKardex(i).Debe
            Debug.Print s
        Next i
        
        'Verificación de datos
        mobjGNCompAux.VerificaDatos
    
        'Verifica si está cuadrado el asiento
        If Not VerificaAsiento(mobjGNCompAux) Then Exit Function
    
        'Verifica si tiene detalle de banco
        If (mobjGNCompAux.CountTSKardex = 0) And _
            (mobjGNCompAux.CountTSKardexRet = 0) And _
            (mobjGNCompAux.CountPCKardex = 0) Then
            MsgBox "No existe ningún detalle.", vbInformation
            Exit Function
        End If
    
        MensajeStatus MSG_GRABANDO, vbHourglass
    
        'Manda a grabar
        '       Aquí ya no hacemos verificación de asiento por que ya está hecho en Control Asiento
        mobjGNCompAux.Grabar False, False
       
        MensajeStatus
        Me.Caption = mobjGNCompAux.CodTrans & " " & mobjGNCompAux.numtrans
        trans = mobjGNCompAux.CodTrans & " " & mobjGNCompAux.numtrans
        GrabarTransCruceAutoxForma = True
    Else
        GrabarTransCruceAutoxForma = False
    End If
    Exit Function
ErrTrap:
    MensajeStatus
    Select Case Err.Number
    Case ERR_DESCUADRADO, ERR_INTEGRIDAD
        'Si es que el usuario seleccionó 'No' en el cuadro de dialogo,
        'No hace nada
    Case Else
        DispErr
    End Select
    GrabarTransCruceAutoxForma = False
    Exit Function
    
End Function


Public Sub InicioCruceIVTarjetas(Name As String)
    Dim i As Integer
    On Error GoTo ErrTrap
    Me.tag = Name
    PicForma.Visible = True
    lblBanco.Visible = False
    fcbBanco.Visible = False
    fcbBanco.TabStop = False
    lblnumche.Visible = False
    lblSal.Visible = False
    lblsaldoBanco.Visible = False
    lblNumCheque.Visible = False
    cmdImprimiCH.Visible = False
    cmdImprimiCH.TabStop = False
    chkAgrupaProv.Caption = "Agrupar por Tarjeta"
    ConfigCols
    Me.Show
    Me.ZOrder
    dtpFecha1.value = Date
    dtpFecha2.value = Date
    CargaTrans
    CargaFormasTarjetas
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Public Sub InicioAnulados(Name As String)
    Dim i As Integer, s As String
    On Error GoTo ErrTrap
    Me.tag = Name
    Form_Resize
    Me.Show
    Me.ZOrder
    dtpFecha2.Visible = False
    dtpFecha2.TabStop = False
    Label2(1).Visible = False
    dtpFecha1.value = Date
    dtpFecha2.value = Date
    FraConFigEgreso.Visible = False
    ConfigCols
'    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("EgresosAutoLibImpPago")) > 0 Then
'        s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("EgresosAutoLibImpPago")
'        txtEgreso.Text = s
'    End If
'
'    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("EgresosAutoLibImpCheque")) > 0 Then
'        s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("EgresosAutoLibImpCheque")
'        txtCheque.Text = s
'    End If
    fraCodTrans.Visible = False
    fraFecha.Visible = False
    FrameFC.Visible = False
    chkAgrupaProv.Visible = False
    
    'fraCodTrans.Caption = "Transaccion"
    fraCodTransVenta.Visible = False
'    fraFecha.Caption = "Fecha Vencimiento "
    CargaTrans
    FraNumero.Visible = True
    cmdBuscar.Caption = "Crear"
    'CargaFormasPago
    'Carga la lista de Bancos
    'fcbBanco.SetData gobjMain.EmpresaActual.ListaTSBanco(True, False)
'    grd.Editable = flexEDKbdMouse
    grd.Refresh
    
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub CreaTransAnulas()
    Dim i As Integer
     With grd
        For i = 1 To .Rows - 1
            .RemoveItem 1
        Next i
        For i = ntxDesde.value To ntxHasta.value
            .AddItem i & vbTab & " " & vbTab & Date & vbTab & vbTab & fcbTransAnulada.KeyText & vbTab & i
        
        Next i
        .Redraw = flexRDDirect
        .Refresh
        
    End With
End Sub

Private Function TransAnuladaAuto(ByVal bandVerificar As Boolean, BandTodo As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, x As Single, j As Integer, filaSubTotal As Long
    Dim gnc As GNComprobante, cambiado As Boolean, TransGen As String
    
    On Error GoTo ErrTrap
    
    'Si no es solo verificacion, confirma
    If Not bandVerificar Then
        'Confirma la actualización
        s = "Este proceso creará Transacciones Anuladas Automáticos  de la transacción seleccionada." & vbCr & vbCr
        s = s & "Está seguro que desea proceder?"
        If MsgBox(s, vbYesNo + vbQuestion) <> vbYes Then Exit Function
    End If
    
    
    Set mColItems = Nothing     'Limpia lo anterior
    Set mColItems = New Collection
    
    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdAceptar.Enabled = False
    cmdBuscar.Enabled = False
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    
    For i = grd.FixedRows To grd.Rows - 1
        DoEvents
        If mCancelado Then
            MsgBox "El proceso fue cancelado.", vbInformation
            Exit For
        End If
        
        prg1.value = i
        grd.Row = i
        x = grd.CellTop                 'Para visualizar la celda actual
        
        If Not grd.IsSubtotal(i) Then
            grd.TextMatrix(i, COL_RESULTADO) = "Procesando  ..."
            grd.Refresh
            TransGen = fcbTransAnulada.KeyText
            If GrabarTransAnuladaAuto(TransGen, grd.TextMatrix(i, COL_NUMTRANS), i) Then
            Else
                    'Si no está cambiado no graba
                grd.TextMatrix(i, COL_RESULTADO) = "Falló Proceso"
            End If
        End If
    Next i
    
    Screen.MousePointer = 0
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    If i < grd.Rows And i >= grd.FixedRows Then
        grd.TextMatrix(i, COL_RESULTADO) = Err.Description
    End If
    DispErr
    prg1.value = prg1.min
salida:
    Set mColItems = Nothing         'Libera el objeto de coleccion
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdBuscar.Enabled = True
    cmdAceptar.Enabled = True
    prg1.value = prg1.min
    Exit Function
End Function


Private Function GrabarTransAnuladaAuto(ByRef trans As String, Num As Long, fila As Long) As Boolean
    Dim Imprime As Boolean, i As Long, ix As Long, orden1 As Integer, orden2 As Integer, j As Integer
    Dim pc As PCProvCli, Cadena As String, obser As String, codforma As String
    Dim tsf As TSFormaCobroPago, x As Single, k As Long, sUMATotal  As Currency
    Dim tid As Long
    On Error GoTo ErrTrap
    GrabarTransAnuladaAuto = True
    If CreaComprobanteAnuladoAuto(Num) Then
        'Si es solo lectura, no hace nada
        If mobjGNCompAux.SoloVer Then
            MsgBox MSG_NODISPONE, vbInformation
            Exit Function
        End If
        
        mobjGNCompAux.numtrans = Num
        mobjGNCompAux.FechaTrans = CDate(grd.TextMatrix(fila, COL_FECHA))
        mobjGNCompAux.HoraTrans = "01:00:00"
        mobjGNCompAux.Descripcion = "Comprobante Anulado Automaticamente"
        mobjGNCompAux.Descripcion = "Comprobante Anulado Automaticamente"
        'mobjGNCompAux.codUsuario = mobjSiiMain.UsuarioActual
        'mobjGNCompAux.Estado = 3
        mobjGNCompAux.NumTransCierrePOS = 3
    
        MensajeStatus MSG_GRABANDO, vbHourglass
    
        mobjGNCompAux.Grabar False, False
        
        mobjGNCompAux.Empresa.CambiaEstadoGNComp mobjGNCompAux.TransID, 0
        mobjGNCompAux.Empresa.CambiaEstadoGNComp mobjGNCompAux.TransID, 3
        
        MensajeStatus
        Me.Caption = mobjGNCompAux.CodTrans & " " & mobjGNCompAux.numtrans
        trans = mobjGNCompAux.CodTrans & " " & mobjGNCompAux.numtrans
        GrabarTransAnuladaAuto = True
    Else
        GrabarTransAnuladaAuto = False
    End If
    Exit Function
ErrTrap:
    MensajeStatus
    Select Case Err.Number
    Case ERR_DESCUADRADO, ERR_INTEGRIDAD
        'Si es que el usuario seleccionó 'No' en el cuadro de dialogo,
        'No hace nada
    Case Else
        DispErr
    End Select
    GrabarTransAnuladaAuto = False
    Exit Function
    
End Function


Private Function CreaComprobanteAnuladoAuto(ByRef Num As Long) As Boolean
    Dim i As Long, gc As GNComprobante
    CreaComprobanteAnuladoAuto = False
    Set gc = gobjMain.EmpresaActual.RecuperaGNComprobante(0, fcbTransAnulada.KeyText, Num)
    If Not gc Is Nothing Then
        MsgBox " Ya existe la transaccion " & fcbTransAnulada.KeyText & " " & Num
        CreaComprobanteAnuladoAuto = False
    Else
        If Len(fcbTransAnulada.KeyText) > 0 Then
            Set mobjGNCompAux = gobjMain.EmpresaActual.CreaGNComprobante(fcbTransAnulada.KeyText)
            CreaComprobanteAnuladoAuto = True
        Else
            MsgBox "Falta seleccionar transaccion "
            CreaComprobanteAnuladoAuto = False
            Exit Function
        End If
    End If
End Function

