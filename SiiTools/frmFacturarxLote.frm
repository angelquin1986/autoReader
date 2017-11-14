VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl1.ocx"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "flexcombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{50067EB3-D6AF-11D3-8297-000021C5085D}#1.0#0"; "ntextbox.ocx"
Begin VB.Form frmFacturarxLote 
   Caption         =   "Facturacion x Lote"
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
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   852
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   8520
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5565
      Width           =   8520
      Begin VB.CommandButton cmdGuardarRes 
         Caption         =   "&Guardar Res."
         Enabled         =   0   'False
         Height          =   372
         Left            =   4560
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar -F3"
         Enabled         =   0   'False
         Height          =   372
         Left            =   2880
         TabIndex        =   0
         Top             =   360
         Width           =   1332
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   6240
         TabIndex        =   1
         Top             =   360
         Width           =   1212
      End
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   3
         Top             =   60
         Width           =   8280
         _ExtentX        =   14605
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin TabDlg.SSTab sst1 
      Height          =   8235
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   19515
      _ExtentX        =   34422
      _ExtentY        =   14526
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Parametros de Busqueda - F6"
      TabPicture(0)   =   "frmFacturarxLote.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraFecha"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "grd"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdBuscar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraEnc"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkImprimir"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdarchivo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "FraDevol"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraCobro"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Factura - F7"
      TabPicture(1)   =   "frmFacturarxLote.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DocsCHP"
      Tab(1).Control(1)=   "Recargos"
      Tab(1).Control(2)=   "ITEMS"
      Tab(1).Control(3)=   "Docs"
      Tab(1).ControlCount=   4
      Begin VB.Frame fraCobro 
         Caption         =   "Formas de cobro"
         Height          =   1155
         Left            =   15420
         TabIndex        =   48
         Top             =   360
         Width           =   2475
         Begin VB.ListBox lstCobro 
            Columns         =   3
            Height          =   852
            IntegralHeight  =   0   'False
            Left            =   0
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   49
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame FraDevol 
         Caption         =   "Parametros"
         Height          =   1215
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Visible         =   0   'False
         Width           =   7695
         Begin VB.PictureBox picCliente 
            BackColor       =   &H00C0C0FF&
            Height          =   555
            Left            =   4020
            ScaleHeight     =   495
            ScaleWidth      =   3555
            TabIndex        =   45
            Top             =   600
            Visible         =   0   'False
            Width           =   3615
            Begin FlexComboProy.FlexCombo fcbCliente 
               Height          =   330
               Left            =   0
               TabIndex        =   46
               Top             =   180
               Width           =   3495
               _ExtentX        =   6165
               _ExtentY        =   582
               DispCol         =   1
               ColWidth1       =   2400
               ColWidth2       =   2400
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
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Cliente"
               Height          =   195
               Left            =   0
               TabIndex        =   47
               Top             =   0
               Width           =   480
            End
         End
         Begin FlexComboProy.FlexCombo fcbTransOri 
            Height          =   330
            Left            =   960
            TabIndex        =   32
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
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
         Begin FlexComboProy.FlexCombo FcbTransDevol 
            Height          =   330
            Left            =   960
            TabIndex        =   33
            Top             =   780
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
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
         Begin MSComCtl2.DTPicker dtpFechaDesde 
            Height          =   360
            Left            =   2520
            TabIndex        =   34
            ToolTipText     =   "Fecha de la transacción"
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
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
            Format          =   99155969
            CurrentDate     =   37078
            MaxDate         =   73415
            MinDate         =   29221
         End
         Begin MSComCtl2.DTPicker dtpFechaHasta 
            Height          =   360
            Left            =   4020
            TabIndex        =   35
            ToolTipText     =   "Fecha de la transacción"
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   635
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
            Format          =   99155969
            CurrentDate     =   37078
            MaxDate         =   73415
            MinDate         =   29221
         End
         Begin FlexComboProy.FlexCombo fcbFormaCobro 
            Height          =   330
            Left            =   2520
            TabIndex        =   36
            Top             =   780
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            ColWidth1       =   2400
            ColWidth2       =   2400
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
         Begin FlexComboProy.FlexCombo fcbComprobante 
            Height          =   330
            Left            =   4020
            TabIndex        =   37
            Top             =   780
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            ColWidth1       =   2400
            ColWidth2       =   2400
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
         Begin FlexComboProy.FlexCombo fcbGrupo2 
            Height          =   330
            Left            =   6180
            TabIndex        =   38
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            ColWidth1       =   2400
            ColWidth2       =   2400
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
         Begin FlexComboProy.FlexCombo fcbBodega 
            Height          =   330
            Left            =   6180
            TabIndex        =   39
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            ColWidth1       =   2400
            ColWidth2       =   2400
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
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Forma Cobro/pago"
            Height          =   195
            Left            =   2520
            TabIndex        =   44
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Trans Origen"
            Height          =   255
            Left            =   50
            TabIndex        =   43
            Top             =   300
            Width           =   915
         End
         Begin VB.Label Label4 
            Caption         =   "Trans Devol"
            Height          =   255
            Left            =   60
            TabIndex        =   42
            Top             =   840
            Width           =   915
         End
         Begin VB.Label Label7 
            Caption         =   "Grupo IV"
            Height          =   255
            Left            =   5520
            TabIndex        =   41
            Top             =   780
            Width           =   915
         End
         Begin VB.Label Label10 
            Caption         =   "Bodega"
            Height          =   255
            Left            =   5520
            TabIndex        =   40
            Top             =   360
            Width           =   915
         End
      End
      Begin VB.CommandButton cmdarchivo 
         Caption         =   "Desde Archivo"
         Height          =   372
         Left            =   1560
         TabIndex        =   30
         Top             =   1620
         Width           =   1212
      End
      Begin SiiToolsA.PCDocCHP DocsCHP 
         Height          =   2175
         Left            =   -66840
         TabIndex        =   29
         Top             =   5640
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   3836
         ProvCliVisible  =   -1  'True
         PorCobrar       =   -1  'True
         FontSize        =   9.75
      End
      Begin VB.CheckBox chkImprimir 
         Caption         =   "Imprimir Despues de Grabar"
         Height          =   255
         Left            =   3120
         TabIndex        =   28
         Top             =   1740
         Width           =   2415
      End
      Begin VB.Frame fraEnc 
         Height          =   1215
         Left            =   7860
         TabIndex        =   17
         Top             =   360
         Width           =   7455
         Begin VB.TextBox txtDescripcion 
            Height          =   510
            Left            =   3600
            MaxLength       =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            ToolTipText     =   "Descripción de la transacción"
            Top             =   600
            Width           =   3780
         End
         Begin NTextBoxProy.NTextBox ntxCotizacion 
            Height          =   330
            Left            =   3600
            TabIndex        =   19
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
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
         Begin MSComCtl2.DTPicker dtpFecha 
            Height          =   360
            Left            =   960
            TabIndex        =   20
            ToolTipText     =   "Fecha de la transacción"
            Top             =   492
            Width           =   1452
            _ExtentX        =   2566
            _ExtentY        =   635
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
            Format          =   99155969
            CurrentDate     =   37078
            MaxDate         =   73415
            MinDate         =   29221
         End
         Begin FlexComboProy.FlexCombo fcbResp 
            Height          =   330
            Left            =   5880
            TabIndex        =   21
            ToolTipText     =   "Responsable de la transacción"
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
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
            Height          =   324
            Left            =   960
            TabIndex        =   22
            ToolTipText     =   "Responsable de la transacción"
            Top             =   840
            Width           =   1452
            _ExtentX        =   2566
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
            Left            =   270
            TabIndex        =   27
            Top             =   840
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&Fecha Transaccion  "
            Height          =   195
            Left            =   1020
            TabIndex        =   26
            Top             =   240
            Width           =   1470
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "&Descripción  "
            Height          =   195
            Left            =   2670
            TabIndex        =   25
            Top             =   600
            Width           =   930
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "C&otización  "
            Height          =   195
            Left            =   2640
            TabIndex        =   24
            Top             =   240
            Width           =   810
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "&Responsable  "
            Height          =   195
            Left            =   4920
            TabIndex        =   23
            Top             =   240
            Width           =   1050
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar - F5"
         Height          =   372
         Left            =   180
         TabIndex        =   6
         Top             =   1620
         Width           =   1212
      End
      Begin SiiToolsA.IVRPVT Recargos 
         Height          =   2235
         Left            =   -74820
         TabIndex        =   13
         Top             =   5580
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   3942
      End
      Begin SiiToolsA.IVGNPVT ITEMS 
         Height          =   2595
         Left            =   -74880
         TabIndex        =   14
         Top             =   480
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   4577
      End
      Begin VSFlex7LCtl.VSFlexGrid grd 
         Height          =   2175
         Left            =   120
         TabIndex        =   15
         Top             =   2100
         Width           =   8175
         _cx             =   14420
         _cy             =   3836
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   4
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
         TabBehavior     =   1
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
         AllowUserFreezing=   2
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin SiiToolsA.PCDoc Docs 
         Height          =   2235
         Left            =   -69180
         TabIndex        =   16
         Top             =   5580
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   3942
         ProvCliVisible  =   0   'False
         PorCobrar       =   0   'False
      End
      Begin VB.Frame fraFecha 
         Caption         =   "Escoja el Grupo a Facturar"
         Height          =   1215
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   3015
         Begin FlexComboProy.FlexCombo fcbGrupo 
            Height          =   330
            Left            =   120
            TabIndex        =   8
            Top             =   480
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   582
            DispCol         =   1
            ColWidth1       =   2400
            ColWidth2       =   2400
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
            Caption         =   "Grupo"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Escoja la transaccion a generar"
         Height          =   1215
         Left            =   3240
         TabIndex        =   10
         Top             =   360
         Width           =   4515
         Begin FlexComboProy.FlexCombo fcbTrans 
            Height          =   345
            Left            =   1200
            TabIndex        =   11
            Top             =   360
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   609
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
         Begin VB.Label Label11 
            Caption         =   "Facturacion"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   1035
         End
      End
   End
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   8640
      Top             =   8760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmFacturarxLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mProcesando As Boolean
Private mCancelado As Boolean
Private mVerificado As Boolean
Private WithEvents mobjGNComp As GNComprobante
Attribute mobjGNComp.VB_VarHelpID = -1
Private gnc As GNComprobante
Private mbooGrabado As Boolean
Private YaImprimio As Boolean
Private mobjImp As Object
Private RecargosItem(10, 20) As Variant
Dim BandCargado As Boolean
Const COL_IDPROVCLI = 1
Const COL_REF = 4
Const COL_IDINVENTARIO = 5
Const COL_IDPROVCLI_GAR = 10
Const COL_IDINVENTARIO_GAR = 5
Const COL_REF_GAR = 3
Const COL_CODITEM_GAR = 6
Const COL_PU_GAR = 8
Const COL_TRANSID = 1

'Conf columnas para fargentex
Const COL_NC_TRANSID = 1
Const COL_NC_CODTRANS = 2
Const COL_NC_NUMTRANS = 3
Const COL_NC_CODCLI = 4
Const COL_NC_NOMCLI = 5
Const COL_NC_CODITEM = 6
Const COL_NC_DESCITEM = 7
Const COL_NC_CANT = 8
Const COL_NC_PRT = 9
Const COL_NC_DSCTO = 10
Const COL_NC_FVENCI = 11
Const COL_NC_FUP = 12
Const COL_NC_DIAS = 13
Const COL_NC_APLINC = 14
Const COL_NC_RES = 15

Private mobjxml As Object

Public Sub Inicio()
    Dim i As Integer
    On Error GoTo ErrTrap
    sst1.Tab = 0
    cmdGrabar.Enabled = False
    Me.Show
    Me.ZOrder
    CargarEncabezado
    CargaCliente
    CargaTrans
    CargarDatos
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Public Sub InicioxCliente() 'facturacion x lote x cliente
  Dim i As Integer
    On Error GoTo ErrTrap
    sst1.Tab = 0
    sst1.TabVisible(1) = False
    cmdGrabar.Caption = "Generar Trans"
    Me.tag = "Xcliente"
    If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "CUENCA") > 0 Then
        ConfigColsDC
    Else
        ConfigCols
    End If
   Me.Show
    Me.ZOrder
    CargarEncabezado
    If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "CUENCA") = 0 Then
        CargaCliente
    End If
    CargaTrans
    CargarDatos
    grd.Editable = False
    'grd.Enabled = False
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub cmdArchivo_Click()
 grd.ColSort(2) = flexSortGenericAscending
    grd.Sort = flexSortUseColSort
    cmdGrabar.Enabled = True
    If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "CUENCA") > 0 Then
        grd.subtotal flexSTSum, 2, 7, , grd.BackColorSel, vbYellow, , "Total", 1, True
        'grd.subtotal flexSTSum, -1, 10, , grd.BackColorSel, vbYellow, , , , True
    Else
        grd.subtotal flexSTSum, 2, 10, , grd.BackColorSel, vbYellow, , "Total", 1, True
        grd.subtotal flexSTSum, -1, 10, , grd.BackColorSel, vbYellow, , , , True
    End If
End Sub

Private Sub cmdBuscar_Click()
    Dim v As Variant, obj As Recordset, s As String
    Dim numGrupo As Integer, NumGrupoDesde  As String
    On Error GoTo ErrTrap
    If Me.tag = "x Garante" Then
        If Len(fcbGrupo.KeyText) = 0 Then
            MsgBox "Seleccione un grupo de cliente.", vbInformation
            fcbGrupo.SetFocus
            Exit Sub
        End If
    ElseIf Me.tag = "Xcliente" Then
        If Len(fcbGrupo.KeyText) = 0 Then
            MsgBox "Seleccione un grupo de cliente.", vbInformation
            fcbGrupo.SetFocus
            Exit Sub
        End If
    End If
    If Me.tag <> "XDevol" Or Me.tag <> "XDevolDscto" Then
        If Len(fcbTrans.Text) = 0 Then
            MsgBox "Seleccione solo un tipo de transacción", vbInformation
                fcbTrans.SetFocus
            Exit Sub
        End If
    End If
    With gobjMain.objCondicion
        .CodTrans = fcbTrans.KeyText  '"'" & fcbTransArribo.Text & "','" & fcbTransSalida.Text & "'"
        .CodPCGrupo = fcbGrupo.KeyText
        If Me.tag = "Xcliente" Then
            Set obj = gobjMain.EmpresaActual.ListaPCProvCliXCliente(numGrupo, .CodPCGrupo, dtpFecha.value) 'VER DESPUES SI SE PUEDE NECESITAR UN FILTRO POR GRUPO
        ElseIf Me.tag = "x Garante" Then
            Set obj = gobjMain.EmpresaActual.ListaPCProvCliXGarante(4, .CodPCGrupo, dtpFecha.value) 'VER DESPUES SI SE PUEDE NECESITAR UN FILTRO POR GRUPO
        ElseIf Me.tag = "XCC" Then
            Set obj = ListaPCProvCliXCCNew(numGrupo, dtpFecha.value, .CodPCGrupo) 'VER DESPUES SI SE PUEDE NECESITAR UN FILTRO POR GRUPO
        ElseIf Me.tag = "XDevol" Then
            Set obj = gobjMain.EmpresaActual.Empresa2.ListaTransxDevolucion(fcbTransOri.KeyText, fcbFormaCobro.KeyText, fcbComprobante.KeyText, dtpFechaDesde.value, dtpFechaHasta.value, fcbGrupo2.KeyText)
        ElseIf Me.tag = "XDevolDscto" Then
            Set obj = gobjMain.EmpresaActual.Empresa2.ListaTransxDevolucionDsctoNew(fcbTransOri.KeyText, fcbCliente.KeyText, FcbTransDevol.KeyText, dtpFechaDesde.value, dtpFechaHasta.value, PreparaCodForma)
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "XDevolDscto", fcbTransOri.KeyText & ";" & FcbTransDevol.KeyText & ";" & fcbFormaCobro.KeyText & ";" & PreparaCodForma
            gobjMain.EmpresaActual.GNOpcion.GrabarSoloGnOpcion2
        Else
            numGrupo = CInt(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("cboFacturaxGrupo"))
            Set obj = gobjMain.EmpresaActual.ListaPCProvCliXGrupo(numGrupo, .CodPCGrupo)
        End If
    End With
    grd.Redraw = flexRDNone
    grd.Rows = 1
    ITEMS.Limpiar
    BandCargado = False
    If Not obj.EOF Then
        v = MiGetRows(obj)
        grd.Redraw = flexRDNone
        grd.LoadArray v
        If Me.tag = "x Garante" Then
            ConfigColsGarantes
        ElseIf Me.tag = "XCC" Then
            ConfigColsXCC
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "GeneraXCC", fcbTrans.KeyText
            gobjMain.EmpresaActual.GNOpcion.GrabarSoloGnOpcion2
        ElseIf Me.tag = "XDevol" Then
            ConfigColsXDevol
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "GeneraXDevol", fcbTransOri.KeyText & ";" & FcbTransDevol.KeyText & ";" & fcbFormaCobro.KeyText & ";" & fcbComprobante.KeyText & ";" & fcbGrupo2.KeyText & ";" & fcbBodega.KeyText
            gobjMain.EmpresaActual.GNOpcion.GrabarSoloGnOpcion2
        ElseIf Me.tag = "XDevolDscto" Then
            ConfigColsXDevolDscto
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "GeneraXDevol", fcbTransOri.KeyText & ";" & FcbTransDevol.KeyText & ";" & fcbFormaCobro.KeyText & ";" & fcbComprobante.KeyText & ";" & fcbGrupo2.KeyText & ";" & fcbBodega.KeyText
            gobjMain.EmpresaActual.GNOpcion.GrabarSoloGnOpcion2
        Else
            ConfigCols
        End If
        grd.Redraw = flexRDDirect
    Else
        grd.Rows = grd.FixedRows
        If Me.tag = "XDevol" Then
            ConfigColsXDevol
        'ElseIf Me.tag = "XDevolCobro" Then
        '    ConfigColsXDevolDscto
         ElseIf Me.tag = "XDevolDscto" Then
            ConfigColsXDevolDscto
        Else
            ConfigCols
        End If
    End If
'    grd.AutoSize 0, grd.Cols - 1
    grd.subtotal flexSTClear
    If Me.tag = "Xcliente" Then
        grd.subtotal flexSTSum, 1, 8, , grd.BackColorSel, vbYellow, , "Total", 1, True
        grd.subtotal flexSTSum, -1, 8, , grd.BackColorSel, vbYellow, , , , True
    ElseIf Me.tag = "XCC" Then
'        grd.subtotal flexSTSum, 2, 10, , grd.BackColorSel, vbYellow, , "Total", 1, True
'        grd.subtotal flexSTSum, -1, 10, , grd.BackColorSel, vbYellow, , , , True
    ElseIf Me.tag = "x Garante" Then
        grd.subtotal flexSTSum, 1, 8, , grd.BackColorSel, vbYellow, , "Total", 1, True
        grd.subtotal flexSTSum, -1, 8, , grd.BackColorSel, vbYellow, , , , True
    ElseIf Me.tag = "XDevol" Then
        grd.subtotal flexSTSum, 1, 8, , grd.BackColorSel, vbYellow, , "Total", 1, True
        grd.subtotal flexSTSum, 1, 9, , grd.BackColorSel, vbYellow, , "Total", 1, True
        grd.subtotal flexSTSum, 1, 10, , grd.BackColorSel, vbYellow, , "Total", 1, True
    ElseIf Me.tag = "XDevolDscto" Then
        grd.subtotal flexSTSum, 3, 9, , grd.BackColorSel, vbYellow, , "Total", 3, True
'        grd.subtotal flexSTSum, 1, 8, , grd.BackColorSel, vbYellow, , "Total", 1, True
'        grd.subtotal flexSTSum, 1, 9, , grd.BackColorSel, vbYellow, , "Total", 1, True
'        grd.subtotal flexSTSum, 1, 10, , grd.BackColorSel, vbYellow, , "Total", 1, True
    End If
    
        grd.MergeCol(1) = True
    grd.MergeCol(2) = True
    grd.MergeCol(3) = True
    
    grd.Redraw = True
    grd.Refresh
    SaveSetting APPNAME, App.Title, "KeyTFacLote", fcbTrans.KeyText
    mVerificado = True
    cmdGrabar.Enabled = True
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub ConfigCols()
    Dim s As String
    With grd
            s = "^#|<IdProvCli|<Codigo|<Nombre|<Referencia|>idinventario|<Cod. Item|<Descripcion|>PU|<TransID|>PLAZO|<Resultado"
            .FormatString = s
            .ColHidden(1) = True
            .ColHidden(5) = True
            .ColHidden(9) = True
            .ColWidth(0) = 500
            .ColWidth(2) = 1500
            .ColWidth(3) = 5500
            .ColWidth(4) = 700
            .ColWidth(5) = 1000
            .ColWidth(6) = 1000
            .ColWidth(7) = 3000
            .ColWidth(8) = 1500
         
            .ColFormat(7) = "#,0.00"
            .ColFormat(10) = "#,0"
        
            GNPoneNumFila grd, False
            
            If Me.tag = "Xcliente" Or Me.tag = "XCC" Then
            Else
                .ColHidden(4) = True
                .ColHidden(6) = True
                .ColHidden(7) = True
                .ColHidden(8) = True
            End If
    End With
End Sub

Private Sub cmdCancelar_Click()
    If mProcesando Then
        mCancelado = True
    Else
        Unload Me
    End If
End Sub


Private Sub cmdGrabar_Click()
Dim gnt As GNTrans

If Len(fcbTrans.KeyText) = 0 Then
    MsgBox "Debe seleccionar Transaccion de Factura"
    fcbTrans.SetFocus
    Exit Sub
End If

If Me.tag = "Xcliente" Then
    Set gnt = gobjMain.EmpresaActual.RecuperaGNTrans(fcbTrans.KeyText)
    If gnt.PcKardex2 Then
        If GrabarxLoteCliente2PCK Then
            cmdGrabar.Enabled = False
            cmdGuardarRes.Enabled = True
        End If
    Else
        If GrabarxLoteCliente Then
            cmdGrabar.Enabled = False
            cmdGuardarRes.Enabled = True
        End If
    End If
ElseIf Me.tag = "XCC" Then
    Set gnt = gobjMain.EmpresaActual.RecuperaGNTrans(fcbTrans.KeyText)
    If gnt.PcKardex2 Then
        If GrabarxLoteCC2PCK() Then
            cmdGrabar.Enabled = False
            cmdGuardarRes.Enabled = True
        End If
    Else
        If GrabarxLoteCliente Then
            cmdGrabar.Enabled = False
            cmdGuardarRes.Enabled = True
        End If
    End If
ElseIf Me.tag = "x Garante" Then
    If GrabarxLoteGarante Then
        cmdGrabar.Enabled = False
        cmdGuardarRes.Enabled = True
    End If
ElseIf Me.tag = "XDevol" Then
    Set gnt = gobjMain.EmpresaActual.RecuperaGNTrans(fcbTrans.KeyText)
    If GrabarDevolucionxTrans Then
        cmdGrabar.Enabled = False
        cmdGuardarRes.Enabled = True
    End If
ElseIf Me.tag = "XDevolDscto" Then
    Set gnt = gobjMain.EmpresaActual.RecuperaGNTrans(fcbTrans.KeyText)
    If GrabarDevolucionxTransDscto Then
        cmdGrabar.Enabled = False
        cmdGuardarRes.Enabled = True
    End If

Else
    If Grabar Then
        cmdGrabar.Enabled = False
        cmdGuardarRes.Enabled = True
    End If
End If
Set gnt = Nothing
End Sub


Private Sub cmdImprimir_Click()
    If Imprimir Then
        cmdCancelar.SetFocus
    End If
End Sub

Private Sub CmdGuardarRes_Click()
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
    For i = 0 To grd.Rows - 1
    If Not grd.IsSubtotal(i) Then
        For j = 1 To grd.Cols - 1
            If Not grd.ColHidden(j) Then
                Cadena = Cadena & grd.TextMatrix(i, j) & ","
            End If
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

Private Sub FcbTransDevol_Selected(ByVal Text As String, ByVal KeyText As String)
    fcbTrans.KeyText = KeyText
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF3
        cmdGrabar_Click
        KeyCode = 0
    Case vbKeyF5
        If sst1.Tab = 0 Then cmdBuscar_Click
        KeyCode = 0
'    Case vbKeyF6
'        sst1.Tab = 0
'
'        KeyCode = 0
'    Case vbKeyF7
'        sst1.Tab = 1

'
'        KeyCode = 0
    
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
    sst1.Move 0, sst1.Top, Me.ScaleWidth, Me.ScaleHeight - pic1.Height - 300
    With grd
        .Width = Me.ScaleWidth - 200
        .Height = Me.ScaleHeight - .Top - pic1.Height - 380
    End With
    With ITEMS
        .Width = Me.ScaleWidth - 200
        .Height = Me.ScaleHeight - (Recargos.Height + pic1.Height + 1000)
    End With
    With Recargos
        .Left = ITEMS.Left
        .Top = ITEMS.Height + 500
        .Width = Me.ScaleWidth / 2
        .Height = 2000
    End With
    With Docs
        .Left = Recargos.Width
        .Top = Recargos.Top
        .Width = (Me.ScaleWidth / 2) - 100
        .Height = 2000
    End With
    prg1.Width = Me.ScaleWidth - (prg1.Left * 2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjImp = Nothing
    Set mobjGNComp = Nothing
End Sub

Private Sub grd_BeforeEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
If Row = 0 Then
    Cancel = True
End If
End Sub

Private Sub grd_BeforeSort(ByVal col As Long, Order As Integer)
Select Case col
Case 0, 1, 3, 4, 5, 6, 7, 8, 9, 10
    Order = 0
End Select
End Sub

Private Sub grd_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyDelete
        EliminarFila
End Select
End Sub
Private Sub EliminarFila()
Dim i As Long
Dim resp
    resp = "¿Realmente desea Borrar"
    If MsgBox(resp, vbYesNo) = vbYes Then
        grd.RemoveItem grd.Row
    End If
End Sub
Private Sub grd_LostFocus()
    If sst1.Tab = 0 Then
'        sst1.Tab = 1
    End If
End Sub

Private Sub sst1_Click(PreviousTab As Integer)
Dim numGrupo As Integer
    On Error GoTo ErrTrap
    Select Case sst1.Tab
    Case 0          'Parametros de Busqueda
        'cmdGrabar.Enabled = False
    Case 1
        If grd.Rows = 1 Then MsgBox "Datos incompletos": sst1.Tab = 0: Exit Sub
        If Not BandCargado Then
            'If lblCodTrans.Caption <> fcbTrans.KeyText Then
             '   lblCodTrans.Caption = fcbTrans.KeyText
                SacaDatosGnTrans (fcbTrans.KeyText)
                CrearGnComprobante
                Docs.PorCobrar = Not mobjGNComp.GNTrans.IVPorPagar
            'End If
            numGrupo = CInt(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("cboFacturaxGrupo"))
            If Me.tag = "Xcliente" Or Me.tag = "XCC" Then
            Else
                ITEMS.MostrarSubItems fcbGrupo.KeyText, numGrupo
            End If
            Recargos.Refresh
            Docs.ActualizarFormaCobroPago
            Docs.VisualizaDesdeObjeto
            Docs.Refresh
            BandCargado = True
        End If
        cmdGrabar.Enabled = True
    End Select
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub CargarEncabezado()
    dtpFecha.value = Date
    fcbResp.SetData gobjMain.EmpresaActual.ListaGNResponsable(False)
    fcbMoneda.SetData gobjMain.EmpresaActual.ListaGNMoneda
    fcbMoneda.KeyText = "USD"
    ntxCotizacion.Text = " 1"
    txtDescripcion.Text = "Factura x Lote..."
End Sub


Private Sub SacaDatosGnTrans(ByVal CodTrans As String)
    Dim gnt As GNTrans
    Set gnt = gobjMain.EmpresaActual.RecuperaGNTrans(CodTrans)
    If Not gnt Is Nothing Then
        fcbResp.KeyText = gnt.CodResponsablePre
    End If
End Sub


Private Sub CrearGnComprobante()
    'Eliminar el que haya tenido
    Set mobjGNComp = Nothing
    'crear el comprobante para luego grabar
    If Len(fcbTrans.KeyText) = 0 Then
        MsgBox "No hay un tipo de transacción para crear"
    Else
        Set mobjGNComp = gobjMain.EmpresaActual.CreaGNComprobante(fcbTrans.KeyText)
        Set ITEMS.GNComprobante = mobjGNComp
        Set Recargos.GNComprobante = mobjGNComp
        Set Docs.GNComprobante = mobjGNComp
    End If
    
End Sub

Private Sub Enc_Aceptar(ByVal i As Long)
    
    If fcbResp.Vacio Then
        fcbResp.SetFocus
        MsgBox "Seleccione el responsable..", vbInformation
        Exit Sub
    End If
    
    'Cotización no puede ser 0
    If ntxCotizacion.value <= 0 Then
        ntxCotizacion.SetFocus
        MsgBox "La cotización no puede ser 0.", vbInformation
        Exit Sub
    End If
    
    
    If Not (mobjGNComp Is Nothing) Then
    
        If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "CATA") <> 0 Then
            mobjGNComp.PCKardex(1).CodProvCli = grd.TextMatrix(i, COL_IDPROVCLI_GAR + 1)
            mobjGNComp.CodClienteRef = grd.TextMatrix(i, COL_IDPROVCLI_GAR + 1)
            mobjGNComp.Nombre = grd.TextMatrix(i, COL_IDPROVCLI_GAR + 2)
            mobjGNComp.CodGaranteRef = grd.TextMatrix(i, COL_REF_GAR)
        Else
            If Me.tag = "XDevol" Then
'                mobjGNComp.PCKardex(1).CodProvCli = grd.TextMatrix(i, 1)
'                mobjGNComp.CodClienteRef = grd.TextMatrix(i, 2)
'                mobjGNComp.nombre = grd.TextMatrix(i, 3)

            Else
                mobjGNComp.PCKardex(1).CodProvCli = grd.TextMatrix(i, 2)
                mobjGNComp.CodClienteRef = grd.TextMatrix(i, 2)
                mobjGNComp.Nombre = grd.TextMatrix(i, 3)
            End If
        
        End If
        mobjGNComp.FechaTrans = dtpFecha.value
        
        mobjGNComp.CodResponsable = fcbResp.KeyText
        mobjGNComp.CodMoneda = fcbMoneda.Text
        mobjGNComp.Cotizacion("") = ntxCotizacion.value
'        mobjGNComp.Descripcion = Trim$(txtDescripcion.Text)
        If mobjGNComp.GNTrans.HoraAuto And mobjGNComp.EsNuevo = True Then
            mobjGNComp.HoraTrans = Time
        End If
    End If
End Sub

    
Private Function Grabar() As Boolean
Dim i As Long
Dim bandFac As Boolean
Dim gc As GNComprobante
    On Error GoTo ErrTrap
    If mobjGNComp Is Nothing Then Exit Function
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    
    For i = grd.FixedRows To grd.Rows - 1
        Enc_Aceptar (i)
        If bandFac = False Then
            GrabarTransacciones (i)
            bandFac = True
        Else
            Comprobante_NuevoFacMulta (i)
        End If
        prg1.value = i
    Next
    prg1.value = prg1.min
    Screen.MousePointer = vbNormal
    Grabar = mbooGrabado
    Exit Function
ErrTrap:
    MensajeStatus
    prg1.value = prg1.min
    Screen.MousePointer = vbNormal
    Select Case Err.Number
    Case ERR_DESCUADRADO, ERR_INTEGRIDAD
        'Si es que el usuario seleccionó 'No' en el cuadro de dialogo,
        'No hace nada
    Case Else
        DispErr
    End Select
    ITEMS.SetFocus  'Para que no se pierda el enfoque
    Exit Function
End Function
Public Sub Limpiar()
    Dim i As Long
    With grd
        For i = .FixedRows To .Rows - 1
            .RowData(i) = 0
        Next i
        .Rows = .FixedRows
        .Rows = .FixedRows
    End With
    
End Sub

Private Sub CargaCliente()
Dim numGrupo As Integer
    numGrupo = CInt(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("cboFacturaxGrupo"))
    
    If Me.tag <> "XDevol" Then
        If numGrupo = 0 Then: MsgBox "No tiene Configurado el sistema para esta opción, ingreses a la informaciones de la empresa y configure": Exit Sub
    End If
        
    
        If Me.tag = "XCC" Then
            fcbGrupo.SetData ListaGNCentroCostoTV(True, True, False)
            fcbGrupo.ColWidth0 = 1000
            fcbGrupo.ColWidth1 = 4000
            fcbGrupo.ColWidth2 = 2000
        ElseIf Me.tag = "XDevol" Then
            FraDevol.Visible = True
            fcbTransOri.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False)
            FcbTransDevol.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False)
            fcbFormaCobro.SetData gobjMain.EmpresaActual.ListaTSFormaCobroPago(True, False, False)
            fcbComprobante.SetData gobjMain.EmpresaActual.ListaAnexoTipoDocumento(False, False)
            fcbGrupo2.SetData gobjMain.EmpresaActual.ListaIVGrupo(2, False, False)
            fcbBodega.SetData gobjMain.EmpresaActual.ListaIVBodega(True, False)
              fcbGrupo.ColWidth0 = 1000
            fcbGrupo.ColWidth1 = 4000
            fcbGrupo.ColWidth2 = 2000
            Label3.Caption = "Transaccion Fuente"
            dtpFechaDesde.value = Date
            dtpFechaHasta.value = Date
        ElseIf Me.tag = "XDevolDscto" Then
            FraDevol.Visible = True
            fcbComprobante.Visible = False
            Label7.Visible = False
            fcbGrupo2.Visible = False
            Label10.Visible = False
            fcbBodega.Visible = False
            fcbTransOri.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False)
            FcbTransDevol.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False)
            fcbFormaCobro.SetData gobjMain.EmpresaActual.ListaTSFormaCobroPago(False, True, False)
            'fcbComprobante.SetData gobjMain.EmpresaActual.ListaAnexoTipoDocumento(False, False)
            'fcbGrupo2.SetData gobjMain.EmpresaActual.ListaIVGrupo(2, False, False)
            fcbBodega.SetData gobjMain.EmpresaActual.ListaIVBodega(True, False)
              fcbGrupo.ColWidth0 = 1000
            fcbGrupo.ColWidth1 = 4000
            fcbGrupo.ColWidth2 = 2000
            Label3.Caption = "Transaccion Fuente"
            dtpFechaDesde.value = Date
            dtpFechaHasta.value = Date
        Else
            fcbGrupo.SetData gobjMain.EmpresaActual.ListaPCGrupo(numGrupo, True, False)
        End If
End Sub

Private Sub CargarDatos()
    'Llena los datos de cabecera
    CargarEncabezado
End Sub

Private Function CrearTransacciones() As Boolean
    On Error GoTo mensaje
    CrearTransacciones = True
    'Transaccion para conteo fisico
    If Len(fcbTrans.KeyText) > 0 Then
            Set mobjGNComp = gobjMain.EmpresaActual.CreaGNComprobante(fcbTrans.KeyText)
            Set ITEMS.GNComprobante = mobjGNComp
            Set Recargos.GNComprobante = mobjGNComp
            Set Docs.GNComprobante = mobjGNComp
            Exit Function
    End If
mensaje:
    DispErr
    CrearTransacciones = False
End Function

Private Sub GrabarTransacciones(ByVal i As Long)
    Dim trans_conteo As String, proceso As Integer, msg As String, pc As PCProvCli, cad As String
    Dim archi As String
    Dim Imprime As Boolean
    On Error GoTo ErrTrap
    YaImprimio = False
    mbooGrabado = False
    ' verificar si estan todos los datos
    If Len(fcbMoneda.Text) = 0 Then
        MsgBox "Debe selecciona una tipo de Modena", vbInformation
        fcbMoneda.SetFocus
        Exit Sub
    End If
    If Val(ntxCotizacion.Text) = 0 Then
        MsgBox "Escriba una cotizacion valida", vbInformation
        ntxCotizacion.SetFocus
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

    If mobjGNComp.CountIVKardex = 0 Then
        MsgBox "No hay ningúna fila para grabar.", vbInformation
        Exit Sub
    End If
    
    MensajeStatus "Grabando Facturacion x lote", vbHourglass
    'Graba los ajustes de inventario
    
    
    
    ITEMS.Aceptar
    If mobjGNComp.CountIVKardex > 0 Then
        proceso = 2
        With mobjGNComp
            If grd.ValueMatrix(i, grd.Cols - 2) <> 0 Then
                .PCKardex(1).FechaVenci = DateAdd("D", grd.ValueMatrix(i, grd.Cols - 2), .PCKardex(1).FechaEmision)
            End If
            .CodResponsable = fcbResp.KeyText
            .CodMoneda = fcbMoneda.KeyText
            .GeneraAsiento
            .GeneraAsientoPresupuesto
            'Verificación de datos
            .VerificaDatos
            .Grabar False, False
            
            Set pc = Nothing
            If .GNTrans.IVComprobanteElectronico Then
                If GeneraComprobanteElectronico(mobjGNComp, mobjxml) Then
                End If
            End If
            
            
            If chkImprimir.value = vbChecked Then
                If ImprimirxCliente(mobjGNComp.TransID) Then

                End If
            End If
        End With
    End If
    If Me.tag = "XDevol" Then
    Else
        grd.TextMatrix(i, grd.Cols - 3) = mobjGNComp.TransID
    End If
        
    grd.TextMatrix(i, grd.Cols - 1) = "Grabado como " & mobjGNComp.CodTrans & mobjGNComp.numtrans
    MensajeStatus "Grabando Factura ", vbHourglass
    mbooGrabado = True
    Exit Sub
ErrTrap:
    grd.TextMatrix(i, grd.Cols - 1) = Err.Description
    MensajeStatus
    DispErr
    Exit Sub
End Sub


Private Sub ActualizaTotalCobrar(ByVal refresh_item As Boolean)
    Dim t As Currency, anticipos As Currency
    t = Recargos.Refresh
End Sub

Private Sub Recargos_DespuesdeEditarGrd()
    ActualizaTotalCobrar True
End Sub

Private Sub Recargos_GotFocus()
    ActualizaTotalCobrar True
End Sub

Private Sub Items_DespuesdeEditarGrd()
    ActualizaTotalCobrar False
End Sub

Private Sub items_TotalizadoItem()
    'Actualiza Recargos para que recalcule y prorratée de nuevo
    ActualizaTotalCobrar True
End Sub

Public Sub CargaPorCobrar()
    On Error GoTo ErrTrap
    Docs.ProvCliVisible = mobjGNComp.GNTrans.IVProvCliPorFila '*** MAKOTO 12/oct/00 Modificado
'    Set Docs.GNComprobante = mobjGNComp
    With Docs
        If mobjGNComp.GNTrans.IVPorPagar Then
            .CodProvCli = mobjGNComp.CodClienteRef
        End If
        .ProvCliVisible = mobjGNComp.GNTrans.IVProvCliPorFila   '*** MAKOTO 12/oct/00 Modificado
        .PorCobrar = Not mobjGNComp.GNTrans.IVPorPagar
        .ModoProveedor = Not .PorCobrar
    End With
    
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub Docs_AgregarFilaAuto(Cancel As Boolean)
    Dim v As Currency
    
    Docs_PorAgregarFila v           'Calcula valor pendiente
    If Docs.PorCobrar Then
        If v <= 0 Then Cancel = True    'Para que no inserte primera fila automáticamente
    Else
        If v >= 0 Then Cancel = True    'Para que no inserte primera fila automáticamente
    End If
End Sub

Private Sub Docs_PorAgregarFila(valorPre As Currency)
    Dim costo As Currency, anticipos As Currency
    With mobjGNComp
'        anticipos
        costo = .IVKardexTotal(True)
        costo = costo + (.IVRecargoTotal(True, False)) * Sgn(costo)
        valorPre = .PCKardexHaberTotal _
                    - .PCKardexDebeTotal _
                    + .TSKardexHaberTotal _
                    - .TSKardexDebeTotal _
                    - costo - anticipos
    End With
    Docs.ActualizarFormaCobroPago
End Sub
'jeaa 07/09/04 para eliminar las filas con valor cero de forma de pago
Private Function VerificaValores() As Boolean
    Dim ix As Long, Num As Long
    Dim i As Long
    Num = Docs.GNComprobante.CountPCKardex
    For ix = Num To 1 Step -1
        If Docs.GNComprobante.PCKardex(ix).Debe = 0 Then
            Docs.EliminaFilaDocs ix
        End If
    Next ix
    VerificaValores = True
    'Si la transacción afecta a PCKardex
    If mobjGNComp.GNTrans.AfectaSaldoPC And _
       mobjGNComp.GNTrans.IVNoVerificaTotalCuadrado = False Then
        '***jeaa 07/09/04 verifica que a consumidor  final  no se de credito
        If mobjGNComp.GNTrans.IVVerificaCobroConsFinal And mobjGNComp.CodClienteRef = "C0001" Then
            'No se puede dar Credito Consumidor Final
            For i = 1 To mobjGNComp.CountPCKardex
                If mobjGNComp.Empresa.RecuperaTSFormaCobroPago( _
                   mobjGNComp.PCKardex(i).codforma).ConsiderarComoEfectivo = False Then
                    MsgBox "No se puede dar Credito Consumidor Final", vbInformation
                    VerificaValores = False
                    Exit Function
                End If
            Next i
        End If
    End If
    If Not TotalCuadrado Then
        VerificaValores = False
        Exit Function
    End If
    Exit Function
End Function


Private Function TotalCuadrado() As Boolean
    Dim t As Currency, p As Currency
    With mobjGNComp
        t = .IVKardexTotal(True)
        t = MiCCur(Format$(t, .FormatoMoneda))  'Redondea al formato de moneda
        t = t + .IVRecargoTotal(True, False) * Sgn(t)
        p = .PCKardexHaberTotal - .PCKardexDebeTotal + .TSKardexHaberTotal - .TSKardexDebeTotal
        
        If t <> p Then
            MsgBox "El valor total de transacción (" & Format(t, "#,0.0000") & _
                   ") y forma de pago/cobro (" & Format(p, "#,0.0000") & _
                   ") no están cuadrados por la diferencia de " & _
                        Format(t - p, "#,0.0000") & " " & _
                        mobjGNComp.CodMoneda & "." & vbCr & vbCr & _
                   "Para grabar la transacción tiene que estar cuadrado.", vbInformation
            TotalCuadrado = False
        Else
            TotalCuadrado = True
        End If
    End With
End Function

Private Sub CargaTrans()
Dim s As String
    fcbTrans.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False)
    s = GetSetting(APPNAME, App.Title, "KeyTFacLote", "")
    fcbTrans.KeyText = s

End Sub

Public Sub Comprobante_NuevoFacMulta(ByVal fil As Long)
    Dim Incremental As Boolean, TransIDs As String
    Dim ix As Long, i As Long, v As Currency, ixr As Long
    Dim item As IVinventario
    Dim CodTrans As String
    Dim Ndias As String
    Dim X As Long
    On Error GoTo ErrTrap
     
    MensajeStatus MSG_PREPARA, vbHourglass
    'Crea el objeto GNComprobante

    Set ITEMS.GNComprobante = mobjGNComp.Empresa.CreaGNComprobante(fcbTrans.KeyText)
    ITEMS.GNComprobante.CodClienteRef = grd.TextMatrix(fil, 2)
    ITEMS.cargarValores mobjGNComp
    RecargosMulta 'calcula recargos
    InicioFacLote fil
    
    MensajeStatus
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    Unload Me
    Exit Sub
End Sub

Public Sub RecargosMulta()
    Dim i As Long, j As Long
    Dim rsRec As Recordset
    Dim totalFilas As Integer
    'Visualiza los detalles que está en GNComprobante
    Set rsRec = ITEMS.GNComprobante.ListaIVKardexRecargo
    totalFilas = rsRec.RecordCount
        Do While Not rsRec.EOF
'          ReDim Preserve RecargosItem(11, i) 'Carga los valores que se facturan despues
          RecargosItem(i, 1) = rsRec!codRecargo
          RecargosItem(i, 2) = rsRec!sign
          RecargosItem(i, 3) = rsRec!porcent
          RecargosItem(i, 4) = rsRec!valor1
          RecargosItem(i, 5) = rsRec!calculo
          RecargosItem(i, 6) = rsRec!Descripcion
          RecargosItem(i, 7) = rsRec!BandModificable
          RecargosItem(i, 8) = rsRec!BandOrigen
          RecargosItem(i, 9) = rsRec!BandProrrateado
          RecargosItem(i, 10) = rsRec!AfectaIvaItem
          RecargosItem(i, 11) = rsRec!BandSeleccionable
          rsRec.MoveNext
          i = i + 1
        Loop
    VisualizaTotalLote (totalFilas)
Set rsRec = Nothing
End Sub
Private Sub VisualizaTotalLote(Filas As Integer)
    Dim t As Currency, tdesc As Currency
    Dim i As Long, v As Currency, ix As Long
    Dim obj As IVKardexRecargo
    Dim p_antes As Currency
'    If (Not mobjGNComp.SoloVer) And (Not mbooVisualizando) Then
'        p_antes = mobjGNComp.IVRecargoTotal(False, True)    'total de recargos prorrateados
'    End If
    t = Abs(ITEMS.GNComprobante.IVKardexTotal(False))    'Total NETO sin recargo prorateado
    tdesc = ITEMS.GNComprobante.IVKardexDescItemTotal
    t = t - tdesc
         For i = 0 To Filas - 1
            'If Not .IsSubtotal(i) Then
                v = 0
                Select Case RecargosItem(i, 8)
                'Si es iva de item
                Case REC_IVAITEM
                    'Coge valor total de iva de cada item
                    v = ITEMS.GNComprobante.IVKardexIVAItemTotal
                    '.TextMatrix(i, COL_PORCENT) = ""
                    RecargosItem(i, 3) = ""
                    
                'Si es recargo/descuento a la fila anterior
                Case REC_SUMA
                    'Si está ingresado el porcentaje, calcula el valor según porcentaje
                    If (Len(RecargosItem(i, 3)) = 0 Or RecargosItem(i, 3) = 0) _
                                                     And CBool(RecargosItem(i, 11) = False) Then        ' esta  en blanco
'                        v = .ValueMatrix(i, COL_VALOR)
                        v = MiCCur(RecargosItem(i, 4))       '*** MAKOTO 29/ene/01 Mod.
                    'Si no, coge el valor fijo
                    Else
                            'Calcula en base a la suma de la fila anterior
                            v = t * RecargosItem(i, 3) / 100
                    End If
                    
                'Si es recargo/descuento al total neto de items
                Case REC_TOTAL
                    'Si está ingresado el porcentaje, calcula el valor según porcentaje
                     If (Len(RecargosItem(i, 3)) = 0 Or RecargosItem(i, 3) = 0) _
                                                And CBool(RecargosItem(i, 11)) = False Then     ' esta  en blanco
                        v = RecargosItem(i, 4)
                    'Si no, coge el valor fijo
                    Else
                        'Calcula en base a la suma de total REAL de items
                        v = Abs(ITEMS.GNComprobante.IVKardexTotal(False)) * RecargosItem(i, 3) / 100
                    End If
                'Si es recargo de item
                Case REC_RECITEM  '***Agregado. Angel. 29/jul/2004
                    'Coge valor total de recargo de cada item
                    v = ITEMS.GNComprobante.IVKardexRecargoItemTotal
                    RecargosItem(i, 3) = ""
                    
                'Si es recargo/descuento a la fila específica
                Case Is > 0
                    'Si está ingresado el porcentaje, calcula el valor según porcentaje
                    If (Len(RecargosItem(i, 3)) = 0 Or RecargosItem(i, 3) = 0) _
                                           And CBool(RecargosItem(i, 11)) = False Then
                    
                        v = MiCCur(RecargosItem(i, 4))      '*** MAKOTO 29/ene/01 Mod.
                    Else
                            'Calcula con el valor de la fila indicada como origen
                            ix = RecargosItem(i, 8) + 1        '*** MAKOTO 11/nov/00
                            If ix < Filas - 1 Then
'                                v = .ValueMatrix(ix, COL_CALCULO)           '***
                                v = MiCCur(RecargosItem(ix, 5))   '*** MAKOTO 29/ene/01 Mod.
                                v = v * MiCCur(RecargosItem(i, 3)) / 100
                            End If
                    End If
                End Select
                
                'Visualiza el valor (Funciona el redondeo de FlexGrid según ColFormat)
                
                RecargosItem(i, 4) = v
                
                '*** MAKOTO 29/ene/01 Agregado. Para tomar el valor redondeado
                v = RecargosItem(i, 4)  'Obtiene el valor redondeado por FlexGrid
'                .TextMatrix(i, COL_VALOR) = v      'Visualiza de nuevo para que
'                                                    'no haya diferencia entre el valor de TextDisplay y ValueMatrix
                'Suma acumulada
                t = t + v * IIf(RecargosItem(i, 2) = "-", -1, 1)
                
                RecargosItem(i, 5) = t
                
                'Si está en modificación
                If (Not ITEMS.GNComprobante.SoloVer) Then
                    'Asigna el valor al objeto IVKardexRecargo
                    Set obj = ITEMS.GNComprobante.IVKardexRecargo(i + 1)
                    obj.porcentaje = MiCCur(RecargosItem(i, 3)) / 100
                    obj.Valor = v * IIf(RecargosItem(i, 2) = "-", -1, 1)
                End If
            'End If
        Next i
    
    
    If (Not ITEMS.GNComprobante.SoloVer) Then
        'Si cambió total de recargos prorrateados
        If p_antes <> ITEMS.GNComprobante.IVRecargoTotal(False, True) Then
            'Prorratea los recargos que deben ser prorrateado
            ITEMS.GNComprobante.ProrratearIVKardexRecargo
            VisualizaTotalLote (Filas)
        End If
    End If
End Sub


Private Sub InicioFacLote(ByVal fil As Long)
    
    mobjGNComp.EsNuevo = True
    'cmdGrabar.Enabled = True
    Form_Resize         'Forzar que ajuste el tamaño de sst1
If GrabarFacLote(fil) Then
    grd.TextMatrix(fil, grd.Cols - 1) = "OK"
Else
    grd.TextMatrix(fil, grd.Cols - 1) = "Error"
End If
End Sub

Private Function GrabarFacLote(ByVal fil As Long) As Boolean
'   Dim obj As GNComprobante        'Para verificar Trans.Afectada
    Dim Imprime As Boolean
    Dim pc As PCProvCli
    Dim NumGrupoControl As String

    Const CONS_NOCONTROLA = 0
    Const CONS_MENSAJE = 1
    Const CONS_NOGRABAR = 2
    Dim i As Long
    Dim pt As PermisoTrans
   On Error GoTo ErrTrap
   'Si la transacción afecta a PCKardex
    If mobjGNComp.GNTrans.AfectaSaldoPC And _
       mobjGNComp.GNTrans.IVNoVerificaTotalCuadrado = False Then
        '***Diego 12/09/2003 verifica que a consumidor  final  no se de credito
        If mobjGNComp.GNTrans.IVVerificaCobroConsFinal And mobjGNComp.CodClienteRef = "C0001" Then
            'No se puede dar Credito Consumidor Final
            For i = 1 To mobjGNComp.CountPCKardex
                If mobjGNComp.Empresa.RecuperaTSFormaCobroPago( _
                   mobjGNComp.PCKardex(i).codforma).ConsiderarComoEfectivo = False Then
                    MsgBox "No se puede dar Credito Consumidor Final", vbInformation
                    Exit Function
                End If
            Next i
        End If
            'AUC verifica que las formas de cobro
        If mobjGNComp.GNTrans.IVControlaCreditos Then
            For i = 1 To mobjGNComp.CountPCKardex
                Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(mobjGNComp.IdClienteRef)
                If mobjGNComp.GNTrans.IVControlaCreditos And mobjGNComp.GNTrans.ClienteVisible Then
                    Select Case mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroCli")
                        Case 0
                            NumGrupoControl = pc.CodGrupo1
                        Case 1
                            NumGrupoControl = pc.CodGrupo2
                        Case 2
                            NumGrupoControl = pc.CodGrupo3
                        Case 3
                            NumGrupoControl = pc.CodGrupo4
                    End Select
                End If
                If Not mobjGNComp.Empresa.Verifica(mobjGNComp.PCKardex(i).codforma, NumGrupoControl) Then
                     MsgBox " Error en las formas de Cobro/Pago ", vbInformation
                    Exit Function
                End If
                Set pc = Nothing
            Next i
         End If

        'Verifica si está cuadrado el total de transacción y total de PCKardex.
        If Not TotalCuadrado Then
            GrabarFacLote = False
          Exit Function
        End If
    End If
    'Verificación de datos
   mobjGNComp.VerificaDatos
    'Verifica si está cuadrado el asiento
    'Manda a grabar   ' Agregado en control de que si esta configurado para mostrar los vueltos  y esta modificada la transaccion
    Dim CContado As Currency, efec As String, mEfec As Currency
    CContado = mobjGNComp.CalculaCobroContado(mobjGNComp.GNTrans.IVCobroContado)
    'Si el usuario puede aprobar la transaccion
    Set pt = gobjMain.GrupoActual.PermisoActual.trans(fcbTrans.KeyText)
    MensajeStatus MSG_GRABANDO, vbHourglass
    
    mobjGNComp.Nombre = grd.TextMatrix(fil, 3)  '**** jeaa 28-11-03 para que cuando se haga una modificacion no se pierda el nombre del consumidor final
   'Si es que algo está modificado debe estar siempre antes de grabar
   
    If mobjGNComp.Modificado Then
        'Regenera el asiento contable
        MensajeStatus MSG_GENERANDOASIENTO, vbHourglass
        mobjGNComp.CodAsiento = 0
        PreparaAsiento True
     '   mobjGNComp.GeneraAsiento
        MensajeStatus
    End If
    
    mobjGNComp.Grabar False, False
    MensajeStatus
    grd.TextMatrix(fil, grd.Cols - 2) = mobjGNComp.TransID
    GrabarFacLote = True
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
   ' sst1_Click sst1.Tab     'Para que no se pierda el enfoque
    Exit Function
End Function


Private Function Imprimir() As Boolean
    Dim s As String, tid As Long, i As Long, X As Single, res As String
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
        If grd.ValueMatrix(i, grd.Cols - 2) <> 0 Then
            DoEvents
            If mCancelado Then
                MsgBox "El proceso fue cancelado."
                Exit For
            End If
            
            prg1.value = i
            grd.Row = i
            X = grd.CellTop                 'Para visualizar la celda actual
            
            tid = grd.ValueMatrix(i, grd.Cols - 2)
            grd.TextMatrix(i, grd.Cols - 1) = "Procesando ..."
            grd.Refresh
            
            'Recupera la transaccion
            Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
            If Not (gnc Is Nothing) Then
                'Si la transacción no está anulado
                If gnc.Estado <> ESTADO_ANULADO Then
                        res = ImprimeTrans(gnc)
                    If Len(res) = 0 Then
                        grd.TextMatrix(i, grd.Cols - 1) = "Enviado."
                    Else
                        grd.TextMatrix(i, grd.Cols - 1) = res
                        cntError = cntError + 1
                    End If
                'Si la transaccion está anulado
                Else
                    grd.TextMatrix(i, grd.Cols - 1) = "Anulado."
                    cntError = cntError + 1
                End If
            Else
                grd.TextMatrix(i, grd.Cols - 1) = "No pudo recuperar la transación."
                cntError = cntError + 1
            End If
        End If
    Next i
    
    Screen.MousePointer = 0
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdGuardarRes.Enabled = True
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

Public Function ImprimeTrans(ByVal gc As GNComprobante) As String
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
    objImp.PrintTrans gobjMain.EmpresaActual, True, 1, 0, "", 0, gc
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

Private Sub PreparaAsiento(Aceptar As Boolean)
    If mobjGNComp.SoloVer Then Exit Sub
    If Aceptar Then
        ITEMS.Aceptar
    End If
    mobjGNComp.GeneraAsiento
End Sub

Private Function GrabarxLoteCliente() As Boolean
Dim i As Long
Dim CodCli As String, ref As String
Dim gc As GNComprobante
Dim bandsaldoFactura As Boolean
    On Error GoTo ErrTrap
    Screen.MousePointer = vbHourglass
    bandsaldoFactura = False
    prg1.min = 0
    prg1.max = grd.Rows - 1
    CodCli = ""
    ref = ""
    For i = grd.FixedRows To grd.Rows - 1
        If Mid(grd.TextMatrix(i, grd.Cols - 1), 1, 7) <> "Grabado" Then
            If Not grd.IsSubtotal(i) Then
'                    If gobjMain.EmpresaActual.VerificaTipoPublicidad(grd.ValueMatrix(i, COL_IDPROVCLI), grd.ValueMatrix(i, COL_IDINVENTARIO)) Then
'                        CrearComprobante grd.ValueMatrix(i, COL_IDPROVCLI), True
'                        Docs.AgregaFila
'                        Enc_Aceptar (i)
'                        GrabarTransacciones (i)
'                        prg1.value = i
'                        ITEMS.Limpiar
'                        Recargos.Limpiar
'                        Docs.Limpiar
'                    Else
                        If CodCli <> grd.TextMatrix(i, 1) Or ref <> grd.TextMatrix(i, COL_REF) Then
                            CrearComprobantexReferencia grd.ValueMatrix(i, COL_IDPROVCLI), True, grd.ValueMatrix(i, COL_REF)
                            Docs.AgregaFila
                            Enc_Aceptar (i)
                            GrabarTransacciones (i)
                            prg1.value = i
                            ITEMS.Limpiar
                            Recargos.Limpiar
                            Docs.Limpiar
                            CodCli = grd.ValueMatrix(i, 1)
                            ref = grd.ValueMatrix(i, COL_REF)
'                        Else
                        End If
'                    End If
            Else
                CodCli = ""
            End If
        End If
    Next
    GrabarxLoteCliente = True
    prg1.value = prg1.min
    Screen.MousePointer = vbNormal
    GrabarxLoteCliente = mbooGrabado
    MensajeStatus "Listo ", vbNormal
    Exit Function
ErrTrap:
    MensajeStatus
    prg1.value = prg1.min
    Screen.MousePointer = vbNormal
    Select Case Err.Number
    Case ERR_DESCUADRADO, ERR_INTEGRIDAD
        'Si es que el usuario seleccionó 'No' en el cuadro de dialogo,
        'No hace nada
    Case Else
        DispErr
    End Select
    ITEMS.SetFocus  'Para que no se pierda el enfoque
    Exit Function
End Function

Private Sub CrearComprobante(ByVal IdProvCli As Long, ByVal bandPublicidad As Boolean)
    SacaDatosGnTrans (fcbTrans.KeyText)
    CrearGnComprobante
    ITEMS.MostrarSubItemsXCliente IdProvCli, bandPublicidad
    Docs.PorCobrar = Not mobjGNComp.GNTrans.IVPorPagar
    Recargos.Refresh
    Docs.ActualizarFormaCobroPago
    Docs.VisualizaDesdeObjeto
    Docs.Refresh
    BandCargado = True
End Sub

Private Function ImprimirxCliente(ByVal tid As Long) As Boolean
    Dim s As String, i As Long, X As Single, res As String
    Dim gnc As GNComprobante, cambiado As Boolean, cntError As Long
    On Error GoTo ErrTrap
            'Recupera la transaccion
            Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
            If Not (gnc Is Nothing) Then
                'Si la transacción no está anulado
                If gnc.Estado <> ESTADO_ANULADO Then
                        res = ImprimeTransxCliente(gnc)
                    If Len(res) = 0 Then
                        grd.TextMatrix(i, grd.Cols - 1) = "Enviado."
                    Else
                        grd.TextMatrix(i, grd.Cols - 1) = res
                        cntError = cntError + 1
                    End If
                'Si la transaccion está anulado
                Else
                    grd.TextMatrix(i, grd.Cols - 1) = "Anulado."
                    cntError = cntError + 1
                End If
            Else
                grd.TextMatrix(i, grd.Cols - 1) = "No pudo recuperar la transación."
                cntError = cntError + 1
            End If
        'End If
    'Next i
    Screen.MousePointer = 0
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdGuardarRes.Enabled = True
    cmdBuscar.Enabled = True
    'Si algúna transaccion no se imprimió, avisa
    If cntError Then
        MsgBox "No se pudo imprimir " & cntError & " transacciones.", vbInformation
    End If
    
    ImprimirxCliente = True
    Exit Function
ErrTrap:
    Screen.MousePointer = 0
    DispErr
    prg1.value = prg1.min
    Exit Function
End Function

Public Function ImprimeTransxCliente(ByVal gc As GNComprobante) As String
    Dim crear As Boolean
    Static objImp As Object
    On Error GoTo ErrTrap

    'Si no tiene TransID quiere decir que no está grabada
    If (gc.TransID = 0) Or gc.Modificado Then
        MsgBox MSGERR_NOGRABADO
        ImprimeTransxCliente = False
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
    objImp.PrintTrans gobjMain.EmpresaActual, True, 1, 0, "", 0, gc
    MensajeStatus "", 0
    ImprimeTransxCliente = ""       'Sin problema
    Exit Function
ErrTrap:
    MensajeStatus "", 0
    Select Case Err.Number
    Case ERR_NOIMPRIME, ERR_NOIMPRIME2, ERR_NOIMPRIME3, ERR_NOHAYCODIGO
        ImprimeTransxCliente = Err.Description
    Case Else
        ImprimeTransxCliente = MSGERR_NOIMPRIME2
    End Select
    Exit Function
End Function

Private Sub CrearComprobantexReferencia(ByVal IdProvCli As Long, ByVal bandPublicidad As Boolean, ByVal ref As String)
    SacaDatosGnTrans (fcbTrans.KeyText)
    CrearGnComprobante
    ITEMS.MostrarSubItemsXClientexReferencia IdProvCli, bandPublicidad, ref
    Docs.PorCobrar = Not mobjGNComp.GNTrans.IVPorPagar
    Recargos.Refresh
    Docs.ActualizarFormaCobroPago
    Docs.VisualizaDesdeObjeto
    Docs.Refresh
    BandCargado = True
End Sub

Public Function GeneraComprobanteElectronico(ByVal gc As GNComprobante, ByRef objImp As Object) As Boolean
    Dim crear As Boolean
    Dim crearRIDE As Boolean
    On Error GoTo ErrTrap

    'Si no tiene TransID quere decir que no está grabada
    If (gc.TransID = 0) Or gc.Modificado Then
        MsgBox MSGERR_NOGRABADO, vbInformation
        GeneraComprobanteElectronico = False
        Exit Function
    End If
    
    
    If gc.CodigoMensaje = "60" Then
        MsgBox "El Documento Electrónico ya fue Autorizado por el SRI "
        Exit Function
    End If
    
    'Solo por primera vez o cuando cambia la librería de impresión
    '  crea una instancia del objeto para la impresión
    crear = (objImp Is Nothing)
    If Not crear Then crear = (objImp.NombreDLL <> gc.GNTrans.ArchivoReporte)
    If crear Then
        Set objImp = Nothing
        'Set objImp = CreateObject(gc.GNTrans.ArchivoReporteRIDE & ".PrintTrans")
        Set objImp = CreateObject("gnxmla.PrintTrans")
    End If
    
   
    MensajeStatus MSG_PREPARA, vbHourglass
    'jeaa 23/11/2006
    objImp.PrintTrans gobjMain.EmpresaActual, True, 1, 0, "", 0, gc
    MensajeStatus
    'jeaa 30/09/04
'    gc.CambiaEstadoImpresion
    GeneraComprobanteElectronico = True
    
    
    
    Exit Function
ErrTrap:
    MensajeStatus
    Select Case Err.Number
    Case ERR_NOIMPRIME, ERR_NOIMPRIME2, ERR_NOIMPRIME3, ERR_NOHAYCODIGO
        DispErr
    Case Else
        
        MsgBox MSGERR_NOIMPRIME2, vbInformation
        
    End Select
    GeneraComprobanteElectronico = False
    Exit Function
End Function

Public Sub InicioxGarante() 'facturacion x lote x cliente
    
    Dim i As Integer
    On Error GoTo ErrTrap
    sst1.Tab = 0
    sst1.TabVisible(1) = False
    cmdGrabar.Caption = "Generar Trans"
    Me.tag = "x Garante"
    ConfigColsGarantes
    Me.Show
    Me.ZOrder
    CargarEncabezado
    CargaGarante
    CargaTrans
    CargarDatos
    grd.Editable = False
    'grd.Enabled = False
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub CargaGarante()
Dim numGrupo As Integer
    numGrupo = 4 'CInt(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("cboFacturaxGrupoGar"))
    
    If numGrupo = 0 Then: MsgBox "No tiene Configurado el sistema para esta opción, ingreses a la informaciones de la empresa y configure": Exit Sub
        
    
    fcbGrupo.SetData gobjMain.EmpresaActual.ListaPCGrupoOrigen(numGrupo, True, False, 3)
End Sub

Private Sub ConfigColsGarantes()
    Dim s As String
    With grd
            s = "^#|<Paralelo|<IdProvCli|<Codigo|<Nombre|>idinventario|<Cod. Item|<Descripcion|>PU|<TransID|>idFactura|<CodCli Factura|<NomCliFactura|<Resultado............................................."
            .FormatString = s
            .ColHidden(2) = True
            .ColHidden(5) = True
            .ColHidden(6) = True
            .ColHidden(9) = True
            .ColHidden(10) = True
            .ColHidden(11) = True
            .ColHidden(12) = True
            
            .ColWidth(0) = 500
            .ColWidth(1) = 1500
            .ColWidth(3) = 1500
            .ColWidth(4) = 5500
            .ColWidth(5) = 1500
            .ColWidth(6) = 1000
            .ColWidth(7) = 4000
            .ColWidth(8) = 1000
            .ColWidth(9) = 1500
         
            .ColFormat(8) = "#,0.00"
            .ColFormat(10) = "#,0"
        
            GNPoneNumFila grd, False
            
   
    End With
End Sub

Private Function GrabarxLoteGarante() As Boolean
Dim i As Long
Dim CodCli As String, ref As String
Dim gc As GNComprobante
Dim bandsaldoFactura As Boolean
    On Error GoTo ErrTrap
    Screen.MousePointer = vbHourglass
    bandsaldoFactura = False
    prg1.min = 0
    prg1.max = grd.Rows - 1
    CodCli = ""
    ref = ""
    For i = grd.FixedRows To grd.Rows - 1
        DoEvents
        If Mid(grd.TextMatrix(i, grd.Cols - 1), 1, 7) <> "Grabado" Then
            If Not grd.IsSubtotal(i) Then
                        'If CodCli <> grd.TextMatrix(i, 1) Or ref <> grd.TextMatrix(i, COL_REF) Then
                            CrearComprobantexGarante grd.TextMatrix(i, COL_CODITEM_GAR), grd.ValueMatrix(i, COL_PU_GAR), "B01"
                            Docs.AgregaFila
                            Enc_Aceptar (i)
                            GrabarTransaccionesGarante (i)
                            prg1.value = i
                            ITEMS.Limpiar
                            Recargos.Limpiar
                            Docs.Limpiar
                            CodCli = grd.ValueMatrix(i, 1)
                            ref = grd.ValueMatrix(i, COL_REF)
'                        Else
'                        End If
'                    End If
            Else
                CodCli = ""
            End If
        End If
    Next
    GrabarxLoteGarante = True
    prg1.value = prg1.min
    Screen.MousePointer = vbNormal
    GrabarxLoteGarante = mbooGrabado
    MensajeStatus "Listo ", vbNormal
    Exit Function
ErrTrap:
    MensajeStatus
    prg1.value = prg1.min
    Screen.MousePointer = vbNormal
    Select Case Err.Number
    Case ERR_DESCUADRADO, ERR_INTEGRIDAD
        'Si es que el usuario seleccionó 'No' en el cuadro de dialogo,
        'No hace nada
    Case Else
        DispErr
    End Select
    ITEMS.SetFocus  'Para que no se pierda el enfoque
    Exit Function
End Function


Private Sub CrearComprobantexGarante(ByVal coditem As String, PU As Currency, bod As String)
    SacaDatosGnTrans (fcbTrans.KeyText)
    CrearGnComprobante
    ITEMS.MostrarItemsXGarante coditem, PU, bod
    Docs.PorCobrar = Not mobjGNComp.GNTrans.IVPorPagar
    Recargos.Refresh
    Docs.ActualizarFormaCobroPago
    Docs.VisualizaDesdeObjeto
    Docs.Refresh
    BandCargado = True
End Sub

Private Sub GrabarTransaccionesGarante(ByVal i As Long)
    Dim trans_conteo As String, proceso As Integer, msg As String, pc As PCProvCli, cad As String
    Dim archi As String
    Dim Imprime As Boolean
    On Error GoTo ErrTrap
    YaImprimio = False
    mbooGrabado = False
    ' verificar si estan todos los datos
    If Len(fcbMoneda.Text) = 0 Then
        MsgBox "Debe selecciona una tipo de Modena", vbInformation
        fcbMoneda.SetFocus
        Exit Sub
    End If
    If Val(ntxCotizacion.Text) = 0 Then
        MsgBox "Escriba una cotizacion valida", vbInformation
        ntxCotizacion.SetFocus
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

    If mobjGNComp.CountIVKardex = 0 Then
        MsgBox "No hay ningúna fila para grabar.", vbInformation
        Exit Sub
    End If
    
    MensajeStatus "Grabando Facturacion x lote", vbHourglass
    'Graba los ajustes de inventario
    
    
    
    ITEMS.Aceptar
    If mobjGNComp.CountIVKardex > 0 Then
        proceso = 2
        With mobjGNComp
            If grd.ValueMatrix(i, grd.Cols - 2) <> 0 Then
                .PCKardex(1).FechaVenci = DateAdd("D", grd.ValueMatrix(i, grd.Cols - 2), .PCKardex(1).FechaEmision)
            End If
            cad = .Descripcion
            If Len(.CodGaranteRef) > 0 Then
                Set pc = .Empresa.RecuperaPCProvCli(.CodGaranteRef)
                cad = "Por Pago Servicio Transporte " & .CodGaranteRef & "-" & pc.Nombre
            End If
            .PCKardex(1).Observacion = Mid$(UCase(MonthName(Month(.FechaTrans))), 1, 3) & "/" & Year(.FechaTrans)
            If Len(pc.CodVehiculoAM) > 0 Then
                .PCKardex(1).CodVendedor = pc.CodVehiculoAM
            Else
                .PCKardex(1).CodVendedor = pc.CodVehiculoPM
            End If
            .IVKardex(1).Nota = Mid$(UCase(MonthName(Month(.FechaTrans))), 1, 3) & "/" & Year(.FechaTrans)
            .Descripcion = Mid$(cad, 1, 119)
            If Len(pc.CodVehiculoAM) > 0 Then
                .CodVendedor = pc.CodVehiculoAM
            Else
                .CodVendedor = pc.CodVehiculoPM
            End If
            
            .CodResponsable = fcbResp.KeyText
            .CodMoneda = fcbMoneda.KeyText
            .GeneraAsiento
            .GeneraAsientoPresupuesto
            'Verificación de datos
            .VerificaDatos
            .Grabar False, False
            
            Set pc = Nothing
            If .GNTrans.IVComprobanteElectronico Then
                If GeneraComprobanteElectronico(mobjGNComp, mobjxml) Then
                End If
            End If
            
            
            If chkImprimir.value = vbChecked Then
                If ImprimirxCliente(mobjGNComp.TransID) Then

                End If
            End If
        End With
    End If
    grd.TextMatrix(i, 9) = mobjGNComp.TransID
    grd.TextMatrix(i, 13) = "Grabado como " & mobjGNComp.CodTrans & mobjGNComp.numtrans
    MensajeStatus "Grabando Factura ", vbHourglass
    mbooGrabado = True
    Exit Sub
ErrTrap:
    grd.TextMatrix(i, grd.Cols - 1) = Err.Description
    MensajeStatus
    DispErr
    Exit Sub
End Sub


Private Function GrabarxLoteCliente2PCK() As Boolean
Dim i As Long
Dim CodCli As String, ref As String
Dim gc As GNComprobante
Dim bandsaldoFactura As Boolean
    On Error GoTo ErrTrap
    Screen.MousePointer = vbHourglass
    bandsaldoFactura = False
    prg1.min = 0
    prg1.max = grd.Rows - 1
    CodCli = ""
    ref = ""
    For i = grd.FixedRows To grd.Rows - 1
        If Mid(grd.TextMatrix(i, grd.Cols - 1), 1, 7) <> "Grabado" Then
            If Not grd.IsSubtotal(i) Then
                        If CodCli <> grd.TextMatrix(i, 1) Or ref <> grd.TextMatrix(i, COL_REF) Then
                            CrearComprobantexReferencia2PCK grd.ValueMatrix(i, COL_IDPROVCLI), True, grd.ValueMatrix(i, COL_REF)
                            DocsCHP.AgregaFila
                            Enc_Aceptar2PCK (i)
                            GrabarTransacciones2PCK (i)
                            prg1.value = i
                            ITEMS.Limpiar
                            Recargos.Limpiar
                            DocsCHP.Limpiar
                            CodCli = grd.ValueMatrix(i, 1)
                            ref = grd.ValueMatrix(i, COL_REF)
'                        Else
                        End If
'                    End If
            Else
                CodCli = ""
            End If
        End If
    Next
    GrabarxLoteCliente2PCK = True
    prg1.value = prg1.min
    Screen.MousePointer = vbNormal
    GrabarxLoteCliente2PCK = mbooGrabado
    MensajeStatus "Listo ", vbNormal
    Exit Function
ErrTrap:
    MensajeStatus
    prg1.value = prg1.min
    Screen.MousePointer = vbNormal
    Select Case Err.Number
    Case ERR_DESCUADRADO, ERR_INTEGRIDAD
        'Si es que el usuario seleccionó 'No' en el cuadro de dialogo,
        'No hace nada
    Case Else
        DispErr
    End Select
    ITEMS.SetFocus  'Para que no se pierda el enfoque
    Exit Function
End Function


Private Sub GrabarTransacciones2PCK(ByVal i As Long)
    Dim trans_conteo As String, proceso As Integer, msg As String, pc As PCProvCli, cad As String
    Dim archi As String
    Dim Imprime As Boolean
    On Error GoTo ErrTrap
    YaImprimio = False
    mbooGrabado = False
    ' verificar si estan todos los datos
    If Len(fcbMoneda.Text) = 0 Then
        MsgBox "Debe selecciona una tipo de Modena", vbInformation
        fcbMoneda.SetFocus
        Exit Sub
    End If
    If Val(ntxCotizacion.Text) = 0 Then
        MsgBox "Escriba una cotizacion valida", vbInformation
        ntxCotizacion.SetFocus
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

    If mobjGNComp.CountIVKardex = 0 Then
'        MsgBox "No hay ningúna fila para grabar.", vbInformation
        Exit Sub
    End If
    
    MensajeStatus "Grabando Facturacion x lote", vbHourglass
    'Graba los ajustes de inventario
    
    
    
    ITEMS.Aceptar
    If mobjGNComp.CountIVKardex > 0 Then
        proceso = 2
        With mobjGNComp
            If grd.ValueMatrix(i, grd.Cols - 2) <> 0 Then
                .PCKardexCHP(1).FechaVenci = DateAdd("D", grd.ValueMatrix(i, grd.Cols - 2), .PCKardexCHP(1).FechaEmision)
            End If
            .CodResponsable = fcbResp.KeyText
            .CodMoneda = fcbMoneda.KeyText
            .GeneraAsiento
            .GeneraAsientoPresupuesto
            'Verificación de datos
            .VerificaDatos
            .Grabar False, False
            
            Set pc = Nothing
            If .GNTrans.IVComprobanteElectronico Then
                If GeneraComprobanteElectronico(mobjGNComp, mobjxml) Then
                End If
            End If
            
            
            If chkImprimir.value = vbChecked Then
                If ImprimirxCliente(mobjGNComp.TransID) Then

                End If
            End If
        End With
    End If
    grd.TextMatrix(i, grd.Cols - 3) = mobjGNComp.TransID
    grd.TextMatrix(i, grd.Cols - 1) = "Grabado como " & mobjGNComp.CodTrans & mobjGNComp.numtrans
    MensajeStatus "Grabando Factura ", vbHourglass
    mbooGrabado = True
    Exit Sub
ErrTrap:
    grd.TextMatrix(i, grd.Cols - 1) = Err.Description
    MensajeStatus
    DispErr
    Exit Sub
End Sub

Private Sub CrearComprobante2PCK(ByVal IdProvCli As Long, ByVal bandPublicidad As Boolean)
    SacaDatosGnTrans (fcbTrans.KeyText)
    CrearGnComprobante2PCK
    ITEMS.MostrarSubItemsXCliente IdProvCli, bandPublicidad
    DocsCHP.PorCobrar = Not mobjGNComp.GNTrans.IVPorPagar
    Recargos.Refresh
    DocsCHP.ActualizarFormaCobroPago
    DocsCHP.VisualizaDesdeObjeto
    DocsCHP.Refresh
    BandCargado = True
End Sub

Private Sub CrearGnComprobante2PCK()
    'Eliminar el que haya tenido
    Set mobjGNComp = Nothing
    'crear el comprobante para luego grabar
    If Len(fcbTrans.KeyText) = 0 Then
        MsgBox "No hay un tipo de transacción para crear"
    Else
        Set mobjGNComp = gobjMain.EmpresaActual.CreaGNComprobante(fcbTrans.KeyText)
        Set ITEMS.GNComprobante = mobjGNComp
        Set Recargos.GNComprobante = mobjGNComp
        Set DocsCHP.GNComprobante = mobjGNComp
    End If
    
End Sub

Private Sub CrearComprobantexReferencia2PCK(ByVal IdProvCli As Long, ByVal bandPublicidad As Boolean, ByVal ref As String)
    SacaDatosGnTrans (fcbTrans.KeyText)
    CrearGnComprobante2PCK
    ITEMS.MostrarSubItemsXClientexReferencia IdProvCli, bandPublicidad, ref
    DocsCHP.PorCobrar = Not mobjGNComp.GNTrans.IVPorPagar
    Recargos.Refresh
    DocsCHP.ActualizarFormaCobroPago
    DocsCHP.VisualizaDesdeObjeto
    DocsCHP.Refresh
    BandCargado = True
End Sub


Private Sub Enc_Aceptar2PCK(ByVal i As Long)
    
    If fcbResp.Vacio Then
        fcbResp.SetFocus
        MsgBox "Seleccione el responsable..", vbInformation
        Exit Sub
    End If
    
    'Cotización no puede ser 0
    If ntxCotizacion.value <= 0 Then
        ntxCotizacion.SetFocus
        MsgBox "La cotización no puede ser 0.", vbInformation
        Exit Sub
    End If
    
    
    If Not (mobjGNComp Is Nothing) Then
    
        mobjGNComp.PCKardexCHP(1).CodProvCli = grd.TextMatrix(i, 2)
        mobjGNComp.CodClienteRef = grd.TextMatrix(i, 2)
        mobjGNComp.Nombre = grd.TextMatrix(i, 3)
        mobjGNComp.FechaTrans = dtpFecha.value
        mobjGNComp.CodResponsable = fcbResp.KeyText
        mobjGNComp.CodMoneda = fcbMoneda.Text
        mobjGNComp.Cotizacion("") = ntxCotizacion.value
        If mobjGNComp.GNTrans.HoraAuto And mobjGNComp.EsNuevo = True Then
            mobjGNComp.HoraTrans = Time
        End If
    End If
End Sub

Private Sub DocsCHP_AgregarFilaAuto(Cancel As Boolean)
    Dim v As Currency
    
    DocsCHP_PorAgregarFila v           'Calcula valor pendiente
    If DocsCHP.PorCobrar Then
        If v <= 0 Then Cancel = True    'Para que no inserte primera fila automáticamente
    Else
        If v >= 0 Then Cancel = True    'Para que no inserte primera fila automáticamente
    End If
End Sub

Private Sub DocsCHP_PorAgregarFila(valorPre As Currency)
    Dim costo As Currency, anticipos As Currency
    With mobjGNComp
'        anticipos
        costo = .IVKardexTotal(True)
        costo = costo + (.IVRecargoTotal(True, False)) * Sgn(costo)
        valorPre = .PCKardexHaberTotal _
                    - .PCKardexDebeTotal _
                    + .TSKardexHaberTotal _
                    - .TSKardexDebeTotal _
                    - costo - anticipos
    End With
    DocsCHP.ActualizarFormaCobroPago
End Sub



Public Sub InicioxCC() 'facturacion x lote x cliente
    Dim trans As String
    Dim i As Integer
    On Error GoTo ErrTrap
    sst1.Tab = 0
    sst1.TabVisible(1) = False
    cmdGrabar.Caption = "Generar Trans"
    Me.tag = "XCC"
    ConfigColsXCC
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("GeneraXCC")) > 0 Then
        trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("GeneraXCC")
        fcbTransOri.KeyText = trans
    End If
    
    
    Me.Show
    Me.ZOrder
    CargarEncabezado
    CargaCliente
    CargaTrans
    CargarDatos
    grd.Editable = False
    'grd.Enabled = False
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Function GrabarxLoteCC2PCK() As Boolean
Dim i As Long
Dim codcontrato As String, ref As String, transant As String
Dim gc As GNComprobante, fecha As String, X As Single
Dim bandsaldoFactura As Boolean
    On Error GoTo ErrTrap
    Screen.MousePointer = vbHourglass
    bandsaldoFactura = False
    prg1.min = 0
    prg1.max = grd.Rows - 1
    codcontrato = ""
    ref = ""
    'fecha = Right("00" & DatePart("M", Date), 2) & "/" & DatePart("YYYY", Date)
    fecha = Right("00" & DatePart("M", dtpFecha.value), 2) & "/" & DatePart("YYYY", dtpFecha.value)
    
    For i = grd.FixedRows To grd.Rows - 1
        DoEvents
        If mCancelado Then
            MsgBox "El proceso fue cancelado.", vbInformation
            Exit For
        End If
        
        grd.Row = i
        X = grd.CellTop
        If Mid(grd.TextMatrix(i, grd.Cols - 1), 1, 7) <> "Grabado" Then
            If Not grd.IsSubtotal(i) Then
                        If codcontrato <> grd.TextMatrix(i, 2) Then
                            If Not VerificaYaFacturado(fcbTrans.Text, grd.TextMatrix(i, 2), fecha, transant, grd.TextMatrix(i, 6)) Then
                                CrearComprobantexCC2PCK grd.ValueMatrix(i, COL_IDPROVCLI), True, grd.TextMatrix(i, 2)
                                DocsCHP.AgregaFila
                                Enc_AceptarCC2PCK (i)
                                GrabarTransacciones2PCK (i)
                                prg1.value = i
                                ITEMS.Limpiar
                                Recargos.Limpiar
                                DocsCHP.Limpiar
                                codcontrato = grd.ValueMatrix(i, 2)
                                ref = grd.ValueMatrix(i, COL_REF)
                           Else
                                grd.TextMatrix(i, grd.Cols - 1) = "Ya Facturado con anterioridad, con " & transant
                                
                            End If
'                        Else
                        End If
'                    End If
            Else
                codcontrato = ""
            End If
        End If
    Next
    GrabarxLoteCC2PCK = True
    prg1.value = prg1.min
    Screen.MousePointer = vbNormal
    GrabarxLoteCC2PCK = mbooGrabado
    MensajeStatus "Listo ", vbNormal
    Exit Function
ErrTrap:

    Screen.MousePointer = 0
    If i < grd.Rows And i >= grd.FixedRows Then
'        grd.TextMatrix(i, COL_RESULTADO) = Err.Description
    End If
    DispErr
    prg1.value = prg1.min

    MensajeStatus
    prg1.value = prg1.min
    Screen.MousePointer = vbNormal
    Select Case Err.Number
    Case ERR_DESCUADRADO, ERR_INTEGRIDAD
        'Si es que el usuario seleccionó 'No' en el cuadro de dialogo,
        'No hace nada
    Case Else
        DispErr
    End Select
    ITEMS.SetFocus  'Para que no se pierda el enfoque
    Exit Function
End Function

Private Sub Enc_AceptarCC2PCK(ByVal i As Long)
    Dim pc As PCProvCli, idpc As Long, gcc As GNCentroCosto
    If fcbResp.Vacio Then
        fcbResp.SetFocus
        MsgBox "Seleccione el responsable..", vbInformation
        Exit Sub
    End If
    
    'Cotización no puede ser 0
    If ntxCotizacion.value <= 0 Then
        ntxCotizacion.SetFocus
        MsgBox "La cotización no puede ser 0.", vbInformation
        Exit Sub
    End If
    
    
    
    If Not (mobjGNComp Is Nothing) Then
        idpc = grd.TextMatrix(i, 1)
        Set pc = mobjGNComp.Empresa.RecuperaPCProvCliQuick(idpc)
        mobjGNComp.PCKardexCHP(1).CodProvCli = pc.CodProvCli
        mobjGNComp.CodClienteRef = pc.CodProvCli
        mobjGNComp.Nombre = pc.Nombre
        mobjGNComp.FechaTrans = dtpFecha.value
        mobjGNComp.CodResponsable = fcbResp.KeyText
        mobjGNComp.CodMoneda = fcbMoneda.Text
        mobjGNComp.CodCentro = grd.TextMatrix(i, 2)
        mobjGNComp.Cotizacion("") = ntxCotizacion.value
        'mobjGNComp.numDocRef = Right("00" & DatePart("M", Date), 2) & "/" & DatePart("YYYY", Date)
        mobjGNComp.numDocRef = UCase(Format(mobjGNComp.FechaTrans, "MMM/yyyy"))
        mobjGNComp.Atencion = UCase(Format(mobjGNComp.FechaTrans, "MMM/yyyy"))
        
        
        
        If mobjGNComp.GNTrans.HoraAuto And mobjGNComp.EsNuevo = True Then
            mobjGNComp.HoraTrans = Time
        End If
        Set pc = Nothing
    End If
End Sub

Private Sub CrearComprobantexCC2PCK(ByVal IdProvCli As Long, ByVal bandPublicidad As Boolean, ByVal ref As String)
    Dim gnc As GNCentroCosto
    SacaDatosGnTrans (fcbTrans.KeyText)
    CrearGnComprobante2PCK
    Set gnc = mobjGNComp.Empresa.RecuperaGNCentroCosto(ref)
    mobjGNComp.CodCentro = ref
    mobjGNComp.CodClienteRef = gnc.codcliente
    mobjGNComp.FechaTrans = dtpFecha.value
    mobjGNComp.CodVendedor = gnc.CodVendedor
    Set gnc = Nothing
    ITEMS.MostrarSubItemsXCC IdProvCli, bandPublicidad, ref
    DocsCHP.PorCobrar = Not mobjGNComp.GNTrans.IVPorPagar
    Recargos.Refresh
    DocsCHP.ActualizarFormaCobroPago
    DocsCHP.VisualizaDesdeObjeto
    DocsCHP.Refresh
    BandCargado = True
End Sub

Private Function VerificaYaFacturado(ByVal CodTrans As String, ByVal codcontrato As String, ByVal mes As String, ByRef trans As String, coditem As String) As Boolean
    Dim sql As String, rs As Recordset
    VerificaYaFacturado = False

    sql = "SELECT top 1 codtrans , numtrans "
    sql = sql & " FROM GNComprobante g inner join gncentrocosto gc on g.idcentro=gc.idcentro "
    sql = sql & " inner join ivkardex i  "
    sql = sql & " inner join ivInventario iv on i.idinventario = iv.idinventario "
    sql = sql & " on g.transid=i.transid "
    sql = sql & "  Where g.estado <>3"
    sql = sql & "  and codtrans='" & CodTrans & "'"
    sql = sql & "  and gc.codcentro='" & codcontrato & "'"
    sql = sql & "  and codinventario='" & coditem & "'"
    sql = sql & "  and i.nota='" & UCase(Format(mes, "MMM/yyyy")) & "'"

    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not rs.EOF Then
        trans = rs!CodTrans & "-" & Str(rs!numtrans)
        VerificaYaFacturado = True
    End If

End Function

Private Sub ConfigColsXCC()
    Dim s As String
    With grd
            s = "^#|<Id|<No.Contrato|<Nombre Factura|<Nombre|>idinventario|<Cod. Item|<Descripcion Plan|>PU|<TransID|>idFactura|<CodCli Factura|<NomCliFactura|<Resultado..................................................................................................."
            .FormatString = s
            .ColHidden(1) = True
            .ColHidden(4) = True
            .ColHidden(5) = True
            .ColHidden(6) = True
            .ColHidden(9) = True
            .ColHidden(10) = True
            .ColHidden(11) = True
            .ColHidden(12) = True
            
            .ColWidth(0) = 500
            .ColWidth(1) = 1000
            .ColWidth(2) = 1000
            .ColWidth(3) = 6500
            .ColWidth(4) = 5500
            .ColWidth(5) = 1500
            .ColWidth(6) = 1000
            .ColWidth(7) = 2000
            .ColWidth(8) = 1000
            .ColWidth(9) = 1500
            .ColWidth(10) = 1500
         
            .ColFormat(8) = "#,0.00"
            .ColFormat(10) = "#,0"
        
            GNPoneNumFila grd, False
            
   
    End With
End Sub

Private Sub AbrirArchivo()
    Dim i As Long
    On Error GoTo ErrTrap
    With dlg1
        .CancelError = True
'        .Filter = "Texto (Separado por coma)|*.txt|Excel 97(XLS)|*.xls"
        .Filter = "Texto (Separado por coma)|*.txt"
        .flags = cdlOFNFileMustExist
        'If Me.tag = "VENTASLOCUTORIOS" Then .filename = lblArchivoLocutorio.Caption
        If Len(.filename) = 0 Then          'Solo por primera vez, ubica a la carpeta de la aplicación
            .filename = App.Path & "\*.txt"
        End If
        
        .ShowOpen
        'If Me.tag = "VENTASLOCUTORIOS" Then lblArchivoLocutorio.Caption = dlg1.filename
        LeerArchivo (dlg1.filename)
    End With
    Exit Sub
ErrTrap:
    If Err.Number <> 32755 Then DispErr
    Exit Sub
End Sub

Private Sub LeerArchivo(ByVal archi As String)
    Select Case UCase$(Right$(archi, 4))
        Case ".TXT"
            ReformartearColumnas
            VisualizarTexto archi
            InsertarColumnas
        Case ".XLS"
            'VisualizarExcel archi
        Case Else
        End Select
End Sub

Private Sub ReformartearColumnas()
' SOLO EN ESTOS CASOS
Select Case UCase(Me.tag)
    Case "DIARIO", "INVENTARIO", "AFINVENTARIO", "AFINVENTARIOC", "PRDIARIO"
        ConfigCols
End Select

End Sub
Private Sub VisualizarTexto(ByVal archi As String)
    Dim f As Integer, s As String, Separador As String, i As Integer
    Dim v As Variant
    ' dim   encontro As Boolean  no  esta el archivo ordenado
    On Error GoTo ErrTrap
    
    MensajeStatus "Está leyendo el archivo " & archi & " ...", vbHourglass
    grd.Rows = grd.FixedRows    'Limpia la grilla
    grd.Redraw = flexRDNone
    f = FreeFile                'Obtiene número disponible de archivo
    
    'Abre el archivo para lectura
    '*** Agregado Oliver 26/03/2004   agrege una opcion especial porque
    Select Case Me.tag                  ' para importar el archivo de ventas de los locutorios
        Case "VENTASLOCUTORIOS"         ' tienen el separador como ;
            Separador = ";"
        Case Else
            Separador = ","
    End Select
    
    'encontro = False
    
    Open archi For Input As #f
        Do Until EOF(f)
            Line Input #f, s
            s = vbTab & Replace(s, Separador, vbTab)      'Convierte ',' a TAB
            
'            If Me.tag = "VENTASLOCUTORIOS" And gConfig.AbrirArchivoenFormaDiferencial Then
'                v = Split(s, vbTab)
'                'Debug.Print s
'                If Not IsEmpty(v) Then
'                       If UBound(v) >= 19 Then
'                            If Val(v(19)) > Val(UltimoNumTransImportado) Then
'                                'encontro = True   no esta ordenado el archivo
'                                grd.AddItem s
'                            End If
'                       End If
'                End If
'            Else
                grd.AddItem s
'            End If
        Loop
    Close #f
    RemueveSpace
    ' ordenar
    If grd.Rows > 1 Then
    Select Case Me.tag
        Case "PLANCUENTA"
            grd.Select 1, 1, 1, 1
        Case "PLANPRCUENTA"
            grd.Select 1, 1, 1, 1
        Case "PLANCUENTASC"
            grd.Select 1, 1, 1, 1
        Case "PLANCUENTAFE"
            grd.Select 1, 1, 1, 1
        Case "ITEM"
            grd.Select 1, 1, 1, 1
        Case "PCPROV"
            grd.Select 1, 1, 1, 1
        Case "PCCLI"
            grd.Select 1, 1, 1, 1
        Case "PORPAGAR"
            grd.Select 1, 1, 1, 4
        Case "PORCOBRAR"
            grd.Select 1, 1, 1, 4
        Case "DIARIO"
            grd.Select 1, 1, 1, 1
        Case "INVENTARIO"
            grd.Select 1, 1, 1, 2
        Case "VENTASLOCUTORIOS"
            'DEB0 ORDENAR SOLO POR LA COLUMNA DE NUMERO DE NOTA DE VENTA
            For i = 0 To 18
                grd.ColSort(i) = flexSortNone
            Next i
            grd.ColSort(19) = flexSortGenericAscending
            grd.Select 1, 1, 1, 19
        Case "AFITEM"
            grd.Select 1, 1, 1, 1
        Case "AFINVENTARIO"
            grd.Select 1, 1, 1, 2
        Case "AFINVENTARIOC"
            grd.Select 1, 1, 1, 2
        Case "PRDIARIO"
            grd.Select 1, 1, 1, 1
        Case "INVENTARIOSERIES"
            grd.Select 1, 1, 1, 2
    End Select
    End If
    grd.Sort = flexSortUseColSort

' poner numero
    GNPoneNumFila grd, False
    
'    If Me.tag = "VENTASLOCUTORIOS" Then
'        SubTotalizar (19)
'        'Totalizar
'        'Almacena la ruta del archivo de importaciòn
'        SaveSetting APPNAME, App.Title, "ArchivoLocutorio", lblArchivoLocutorio.Caption
'    End If
    
    grd.Redraw = flexRDDirect
    AjustarAutoSize grd, -1, -1
    
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

Private Sub InsertarColumnas()
'Dim i As Integer
'Select Case UCase(Me.tag)
'    Case "DIARIO"
'        'sumar
'        With grd
'
'        InsertarColumnaCuenta
'        For i = .FixedRows To .Rows - 1 ' poner nombre de cuentas en columna cuenta
'            DoEvents
'            .TextMatrix(i, .ColIndex("Cuenta")) = ponerCuentaFila(.TextMatrix(i, 1))
'        Next i
'        Sumar
'        End With
'    Case "INVENTARIO"
'        InsertarColumnaDesc_y_Cost
'        For i = grd.FixedRows To grd.Rows - 1 ' poner nombre de cuentas en columna cuenta
'            DoEvents
'            grd.TextMatrix(i, grd.ColIndex("Descripción")) = ponerDescripcionFila(grd.TextMatrix(i, 1))
'            grd.TextMatrix(i, grd.ColIndex("Costo Unitario")) = ponerCostoUnitarioFila(i)
'        Next i
'    Case "AFINVENTARIO"
'        InsertarColumnaDesc_y_Cost
'        For i = grd.FixedRows To grd.Rows - 1 ' poner nombre de cuentas en columna cuenta
'            DoEvents
'            grd.TextMatrix(i, grd.ColIndex("Descripción")) = ponerDescripcionFilaAF(grd.TextMatrix(i, 1))
''            grd.TextMatrix(i, grd.ColIndex("Costo Unitario")) = ponerCostoUnitarioFila(i)
'        Next i
'    Case "AFINVENTARIOC"
'        InsertarColumnaDesc_y_Cost
'        For i = grd.FixedRows To grd.Rows - 1 ' poner nombre de cuentas en columna cuenta
'            DoEvents
'            grd.TextMatrix(i, grd.ColIndex("Descripción")) = ponerDescripcionFilaAF(grd.TextMatrix(i, 1))
''            grd.TextMatrix(i, grd.ColIndex("Costo Unitario")) = ponerCostoUnitarioFila(i)
'        Next i
'    Case "PRDIARIO"
'        'sumar
'        With grd
'
'        InsertarColumnaCuenta
'        For i = .FixedRows To .Rows - 1 ' poner nombre de cuentas en columna cuenta
'            DoEvents
'            .TextMatrix(i, .ColIndex("Cuenta")) = ponerPRCuentaFila(.TextMatrix(i, 1))
'        Next i
'        Sumar
'        End With
'    Case "INVENTARIOSERIES"
'        InsertarColumnaDescSeries
'        For i = grd.FixedRows To grd.Rows - 1 ' poner nombre de cuentas en columna cuenta
'            DoEvents
'            grd.TextMatrix(i, grd.ColIndex("Descripción")) = ponerDescripcionFila(grd.TextMatrix(i, 1))
'        Next i
'
'End Select
'AjustarAutoSize grd, -1, -1
End Sub

Public Function ListaPCProvCliXCCNew(numGrupo As Integer, f As Date, codcontrato As String) As Variant
    Dim sql As String, rs As Recordset, NumReg As Long
    On Error GoTo CapturaError
     
    sql = "SELECT IdProvCli,COUNT(idprovcli) AS NumReg,FechaIni, FechaFin, "
    sql = sql & " i.codinventario, i.descripcion as descItem, (precio1*cantidad) as pu , "
    sql = sql & " i.idInventario, Referencia,bandPublicidad, plazO, CONTRATO "
    sql = sql & " Into TMP1299 "
    sql = sql & " FROM PCProvCliInv inner join ivinventario i on  PCProvCliInv.idinventario= i.idinventario"
    sql = sql & " WHERE CHARINDEX('" & DatePart("m", f) & "', frecuencia)> 0 "
    sql = sql & " GROUP BY IDPROVCLI, FechaIni, FechaFin,i.codinventario, i.descripcion, precio1, "
    sql = sql & " cantidad, i.idInventario, Referencia,bandPublicidad, plazo, CONTRATO  "
    
    VerificaExistenciaTabla 1299 'tmp99 sera utilizada para guardar la primer parte del sql
'    Me.EjecutarSQL sql, numReg
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    sql = "SELECT   "
    sql = sql & " pc.IdProvCli , gcc.CodCentro, pc.Nombre, T.Referencia, t.IdInventario, "
    sql = sql & " codinventario, descItem, pu, 0 as transid, idcentro "
    sql = sql & " FROM gncentrocosto gcc  "
    sql = sql & " inner join PCProvCli pc  on gcc.idcliente =pc.idprovcli  "
    sql = sql & " INNER JOIN TMP1299 T ON T.idprovcli = PC.idprovcli AND GCC.CODCENTRO= T.CONTRATO"
    sql = sql & " WHERE BandCliente=1"
    sql = sql & " AND BandLote=1 "
    sql = sql & " AND Bandocupado=0 "
    sql = sql & " and '" & f & "' between t.FechaIni AND t.FechaFin "
    If Len(codcontrato) > 0 Then
        sql = sql & " and gcc.codcentro='" & codcontrato & "'"
    End If
    sql = sql & " GROUP BY pc.idprovcli, pc.CodProvCli, pc.Nombre,pu, t.idinventario,codinventario, "
    sql = sql & " descItem , T.Referencia,bandPublicidad "
    sql = sql & " ,idcentro, gcc.codcentro"
    sql = sql & " ORDER BY pc.nombre, idcentro, T.Referencia ,bandPublicidad desc "
     
     Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
     Set ListaPCProvCliXCCNew = rs
      Set rs = Nothing
    Exit Function
CapturaError:
    MsgBox Err.Description
    Set rs = Nothing
    Exit Function
End Function

Public Function ListaGNCentroCostoTV(BandSoloValida As Boolean, BandDetallado As Boolean, BandRS As Boolean) As Variant
    Dim sql As String, rs As Recordset
        sql = "SELECT CodCentro, Nombre,  "
        sql = sql & " pccp.descripcion + ' ' + GNCentroCosto.numcasa + ' ' + pccs.descripcion"
        sql = sql & " IdCentro, BANDOCUPADO "
        sql = sql & " FROM GNCentroCosto  "
        sql = sql & " left join pccalle as pccp on GNCentroCosto.idcallepri = pccp.idcalle "
        sql = sql & " left join pccalle as pccs on GNCentroCosto.idcallepri = pccs.idcalle "
    If BandSoloValida Then sql = sql & "WHERE FechaFinal Is Null "
    sql = sql & " ORDER BY Nombre"

   Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If BandRS Then
        Set ListaGNCentroCostoTV = rs
    Else
        ListaGNCentroCostoTV = MiGetRows(rs)
        rs.Close
    End If
    Set rs = Nothing
End Function

Public Sub InicioDevolucion()
    Dim i As Integer
    Dim trans As String, v As Variant
    On Error GoTo ErrTrap
    sst1.Tab = 0
    sst1.TabVisible(1) = False
    cmdGrabar.Caption = "Generar Trans"
    Me.tag = "XDevol"
    ConfigColsXDevol
    Me.Show
    Me.ZOrder
    CargarEncabezado
    CargaCliente
    CargaTrans
    CargarDatos
    grd.Editable = False
    fraFecha.Visible = False
    Frame5.Visible = False
    FraDevol.Visible = True
    
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("GeneraXDevol")) > 0 Then
        trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("GeneraXDevol")
        v = Split(trans, ";")
        fcbTransOri.KeyText = v(0)
        FcbTransDevol.KeyText = v(1)
        fcbFormaCobro.KeyText = v(2)
        fcbComprobante.KeyText = v(3)
        fcbGrupo2.KeyText = v(4)
        fcbBodega.KeyText = v(5)
    End If

    
    'grd.Enabled = False
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub ConfigColsXDevol()
    Dim s As String
    With grd
            s = "^#|<TransId|<Fecha|<N.Factura|<Nombre Cliente|<Grupo|<Codigo Item|<Descripcion|>Cantidad|>P.Unitario|>P.Total|>Orden|<FormaPago|<Tipo documento|<Resultado"
            .FormatString = s
            .ColHidden(1) = True
            .ColWidth(0) = 500
            .ColWidth(1) = 1000
            .ColWidth(2) = 1200
            .ColWidth(3) = 1000
            .ColWidth(4) = 3500
            .ColWidth(5) = 1500
            
            .ColWidth(6) = 1500
            .ColWidth(7) = 3500
            
            .ColWidth(8) = 800
            .ColWidth(9) = 1000
            .ColWidth(9) = 1000
            .ColWidth(10) = 1000
            .ColWidth(11) = 1000
            .ColWidth(12) = 2500
            .ColWidth(14) = 3500
         
         .ColHidden(6) = True
         .ColHidden(13) = True
        
            GNPoneNumFila grd, False
            
  
    End With
End Sub


Private Function GrabarDevolucionxTrans() As Boolean
Dim i As Long, FilaIni As Integer, FilaFin As Integer
Dim CodCli As String, ref As String
Dim gc As GNComprobante
Dim bandsaldoFactura As Boolean
    On Error GoTo ErrTrap
    Screen.MousePointer = vbHourglass
    bandsaldoFactura = False
    prg1.min = 0
    prg1.max = grd.Rows - 1
    CodCli = ""
    ref = ""
    FilaIni = 1
    For i = grd.FixedRows To grd.Rows - 1
        If Mid(grd.TextMatrix(i, grd.Cols - 1), 1, 7) <> "Grabado" Then
            If grd.IsSubtotal(i) Then
                FilaFin = i - 1
                CrearComprobantexTrans FilaIni, FilaFin
                FilaIni = i + 1
                Docs.AgregaFila
                Enc_Aceptar (i)
                GrabarTransacciones (i)
                prg1.value = i
                ITEMS.Limpiar
                Recargos.Limpiar
                Docs.Limpiar
                CodCli = grd.ValueMatrix(i, 1)
                ref = grd.ValueMatrix(i, COL_REF)
            Else
                CodCli = ""
            End If
        End If
    Next
    GrabarDevolucionxTrans = True
    prg1.value = prg1.min
    Screen.MousePointer = vbNormal
'    GrabarxLoteCliente = mbooGrabado
    MensajeStatus "Listo ", vbNormal
    Exit Function
ErrTrap:
    MensajeStatus
    prg1.value = prg1.min
    Screen.MousePointer = vbNormal
    Select Case Err.Number
    Case ERR_DESCUADRADO, ERR_INTEGRIDAD
        'Si es que el usuario seleccionó 'No' en el cuadro de dialogo,
        'No hace nada
    Case Else
        DispErr
    End Select
    ITEMS.SetFocus  'Para que no se pierda el enfoque
    Exit Function
End Function


Private Sub CrearComprobantexTrans(fila As Integer, FilaFin As Integer)
    Dim id As Long, i As Long, j As Integer, numitem As Integer, ix As Long, ct As Currency
    Dim item As IVinventario
    SacaDatosGnTrans (fcbTrans.KeyText)
    CrearGnComprobante
    id = grd.ValueMatrix(fila, COL_TRANSID)
'    Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(id)
    mobjGNComp.ImportaDatos2 id, False
    mobjGNComp.idTransFuente = id
    
    For i = 1 To mobjGNComp.CountIVKardex
        mobjGNComp.RemoveIVKardex 1
    Next i
        
    
        For j = fila To FilaFin
            Set item = mobjGNComp.Empresa.RecuperaIVInventarioQuick(grd.TextMatrix(j, 6))
            ix = mobjGNComp.AddIVKardex
            mobjGNComp.IVKardex(ix).CodBodega = fcbBodega.KeyText
            mobjGNComp.IVKardex(ix).CodInventario = item.CodInventario
            mobjGNComp.IVKardex(ix).cantidad = grd.ValueMatrix(j, 8)
            mobjGNComp.IVKardex(ix).PrecioTotal = grd.ValueMatrix(j, 10)
            mobjGNComp.IVKardex(ix).PrecioRealTotal = grd.ValueMatrix(j, 10)
            mobjGNComp.IVKardex(ix).IVA = item.PorcentajeIVA
            mobjGNComp.IVKardex(ix).Orden = ix
                            ct = item.CostoDouble2(mobjGNComp.FechaTrans, _
                    Abs(grd.ValueMatrix(j, 8)), _
                    mobjGNComp.TransID, _
                    mobjGNComp.HoraTrans)
            mobjGNComp.IVKardex(ix).CostoTotal = ct
            mobjGNComp.IVKardex(ix).CostoRealTotal = ct
        Next j
    If mobjGNComp.GNTrans.IVComprobanteElectronico Then
        mobjGNComp.codMotivoDev = "CLI-INCONF"
    End If
    'importacionItems
    Docs.PorCobrar = Not mobjGNComp.GNTrans.IVPorPagar
    Recargos.Refresh
    Docs.ActualizarFormaCobroPago
    Docs.VisualizaDesdeObjeto
    Docs.Refresh
    BandCargado = True
End Sub

Private Sub ConfigColsDC()
    Dim s As String
    With grd
            s = "^#|<Fecha|<NumTrans|<CodCliente|<CotItem|>Cant|>PU|>PT|<Observacion|<Resultado"
            .FormatString = s
            .ColWidth(0) = 500
            .ColWidth(2) = 1500
            .ColWidth(3) = 5500
            .ColWidth(4) = 700
            .ColWidth(5) = 1000
            .ColWidth(6) = 1000
            .ColWidth(7) = 3000
            .ColWidth(8) = 1500
            .ColFormat(6) = "#,0.00"
            .ColFormat(7) = "#,0.00"
            GNPoneNumFila grd, False
    End With
End Sub
Private Function GrabarxLoteClienteDC() As Boolean
Dim i As Long
Dim CodCli As String, ref As String
Dim gc As GNComprobante
Dim bandsaldoFactura As Boolean
Dim ix As Long
Dim iy As Long
Dim Valor As Currency
    On Error GoTo ErrTrap
    Screen.MousePointer = vbHourglass
    bandsaldoFactura = False
    prg1.min = 0
    prg1.max = grd.Rows - 1
    CodCli = ""
    ref = ""
    For i = grd.FixedRows To grd.Rows - 1
        If Mid(grd.TextMatrix(i, grd.Cols - 1), 1, 7) <> "Grabado" Then
            If Not grd.IsSubtotal(i) Then
                If CodCli <> grd.TextMatrix(i, 3) Or ref <> grd.TextMatrix(i, 2) Then
                    CrearComprobantexReferenciaDC
'                    DOCS.AgregaFila
                    Enc_Aceptar (i)
                    Valor = cargaPCKardex(i)
                        iy = mobjGNComp.AddPCKardex
                    mobjGNComp.PCKardex(iy).codforma = mobjGNComp.GNTrans.CodFormaPre
                    mobjGNComp.PCKardex(iy).CodProvCli = grd.TextMatrix(i, 3)
                    mobjGNComp.PCKardex(iy).NumLetra = grd.TextMatrix(i, 2)
                        mobjGNComp.PCKardex(iy).FechaEmision = mobjGNComp.FechaTrans
                    mobjGNComp.PCKardex(iy).FechaVenci = mobjGNComp.FechaTrans
                    mobjGNComp.PCKardex(iy).Debe = Valor
                     i = CargaIvKardex(i)
                    GrabarTransacciones (i - 1)
                    prg1.value = i
                    CodCli = grd.TextMatrix(i, 3)
                    ref = grd.TextMatrix(i, 2)
    '                        Else
                End If
'                    End If
            Else
                CodCli = ""
            End If
        End If
    Next
    GrabarxLoteClienteDC = True
    prg1.value = prg1.min
    Screen.MousePointer = vbNormal
    GrabarxLoteClienteDC = mbooGrabado
    MensajeStatus "Listo ", vbNormal
    Exit Function
ErrTrap:
    MensajeStatus
    prg1.value = prg1.min
    Screen.MousePointer = vbNormal
    Select Case Err.Number
    Case ERR_DESCUADRADO, ERR_INTEGRIDAD
        'Si es que el usuario seleccionó 'No' en el cuadro de dialogo,
        'No hace nada
    Case Else
        DispErr
    End Select
    ITEMS.SetFocus  'Para que no se pierda el enfoque
    Exit Function
End Function
Private Sub CrearComprobantexReferenciaDC()
    SacaDatosGnTrans (fcbTrans.KeyText)
    CrearGnComprobante
    Docs.PorCobrar = Not mobjGNComp.GNTrans.IVPorPagar
    Recargos.Refresh
    Docs.ActualizarFormaCobroPago
    Docs.VisualizaDesdeObjeto
    Docs.Refresh
    BandCargado = True
End Sub
Private Sub GrabarTransaccionesDC(ByVal i As Long)
    Dim trans_conteo As String, proceso As Integer, msg As String, pc As PCProvCli, cad As String
    Dim archi As String
    Dim Imprime As Boolean
    On Error GoTo ErrTrap
    YaImprimio = False
    mbooGrabado = False
    ' verificar si estan todos los datos
    If Len(fcbMoneda.Text) = 0 Then
        MsgBox "Debe selecciona una tipo de Modena", vbInformation
        fcbMoneda.SetFocus
        Exit Sub
    End If
    If Val(ntxCotizacion.Text) = 0 Then
        MsgBox "Escriba una cotizacion valida", vbInformation
        ntxCotizacion.SetFocus
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
    If mobjGNComp.CountIVKardex = 0 Then
        MsgBox "No hay ningúna fila para grabar.", vbInformation
        Exit Sub
    End If
    MensajeStatus "Grabando Facturacion x lote", vbHourglass
    'Graba los ajustes de inventario
    'ITEMS.Aceptar
    If mobjGNComp.CountIVKardex > 0 Then
        proceso = 2
        With mobjGNComp
            If grd.ValueMatrix(i, grd.Cols - 2) <> 0 Then
                .PCKardex(1).FechaVenci = DateAdd("D", grd.ValueMatrix(i, grd.Cols - 2), .PCKardex(1).FechaEmision)
            End If
            .CodResponsable = fcbResp.KeyText
            .CodMoneda = fcbMoneda.KeyText
            .GeneraAsiento
            .GeneraAsientoPresupuesto
            'Verificación de datos
            .VerificaDatos
            .Grabar False, False
            Set pc = Nothing
            If .GNTrans.IVComprobanteElectronico Then
                If GeneraComprobanteElectronico(mobjGNComp, mobjxml) Then
                End If
            End If
            If chkImprimir.value = vbChecked Then
                If ImprimirxCliente(mobjGNComp.TransID) Then
                End If
            End If
        End With
    End If
    If Me.tag = "XDevol" Then
    Else
        grd.TextMatrix(i, grd.Cols - 3) = mobjGNComp.TransID
    End If
    grd.TextMatrix(i, grd.Cols - 1) = "Grabado como " & mobjGNComp.CodTrans & mobjGNComp.numtrans
    MensajeStatus "Grabando Factura ", vbHourglass
    mbooGrabado = True
    Exit Sub
ErrTrap:
    grd.TextMatrix(i, grd.Cols - 1) = Err.Description
    MensajeStatus
    DispErr
    Exit Sub
End Sub
Private Function CargaIvKardex(ByVal i As Long) As Long
Dim ix As Long
Dim j As Long
    For j = i To grd.Rows - 1
        If grd.IsSubtotal(j) Then
            Exit For
        Else
            ix = mobjGNComp.AddIVKardex
            mobjGNComp.IVKardex(ix).CodInventario = grd.TextMatrix(j, 4)
            mobjGNComp.IVKardex(ix).CodBodega = mobjGNComp.GNTrans.CodBodegaPre
            mobjGNComp.IVKardex(ix).cantidad = grd.ValueMatrix(j, 5) * -1
            mobjGNComp.IVKardex(ix).PrecioTotal = grd.ValueMatrix(j, 7) * -1
            mobjGNComp.IVKardex(ix).PrecioRealTotal = grd.ValueMatrix(j, 7) * -1
        End If
    Next
    CargaIvKardex = j
End Function
Private Function cargaPCKardex(ByVal i As Long) As Currency
Dim ix As Currency
Dim j As Long
    For j = i To grd.Rows - 1
        If grd.IsSubtotal(j) Then
            Exit For
        Else
            ix = ix + grd.ValueMatrix(j, 7)
        End If
    Next
    cargaPCKardex = ix
End Function
Public Sub InicioDevolucionxDscto()
    Dim i As Integer
    Dim trans As String, v As Variant
    On Error GoTo ErrTrap
    
    sst1.Tab = 0
    sst1.TabVisible(1) = False
    cmdGrabar.Caption = "Generar Trans"
    Me.tag = "XDevolDscto"
    ConfigColsXDevolDscto
    Me.Show
    Me.ZOrder
    CargarEncabezadoxDscto
    CargaCliente
    CargaTrans
'    CargarDatos
    grd.Editable = True
    fraFecha.Visible = False
    Frame5.Visible = False
    FraDevol.Visible = True
    picCliente.Visible = True
    fraCobro.Visible = True
    CargaFCobro
    
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("XDevolDscto")) > 0 Then
        trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("XDevolDscto")
        v = Split(trans, ";")
        fcbTransOri.KeyText = v(0)
        FcbTransDevol.KeyText = v(1)
        fcbFormaCobro.KeyText = v(2)
        RecuperaCodForma v(3), lstCobro
    End If
    'grd.Enabled = False
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub CargarEncabezadoxDscto()
    dtpFecha.value = Date
    fcbResp.SetData gobjMain.EmpresaActual.ListaGNResponsable(False)
    fcbMoneda.SetData gobjMain.EmpresaActual.ListaGNMoneda
    fcbCliente.SetData gobjMain.EmpresaActual.ListaPCProvCli(False, True, False)
    fcbMoneda.KeyText = "USD"
    ntxCotizacion.Text = " 1"
    txtDescripcion.Text = "Notas de Credito..."
End Sub

Private Sub ConfigColsXDevolDscto()
    Dim s As String
    With grd
            s = "^#|<Transid|<CodTrans|<NumTrans|<Codigo|<Nombre Cliente|<Codigo Item|<Descripcion|>Cantidad|>P.Real.Total|>Descuento|<FechaVenci|<FechaPago|>DiasPago|^AplicaNC|<Resultado"
            .FormatString = s
            .ColHidden(COL_NC_TRANSID) = True
            .ColWidth(COL_NC_CODTRANS) = 800
            .ColWidth(COL_NC_NUMTRANS) = 800
            .ColWidth(COL_NC_CODCLI) = 1000
            .ColWidth(COL_NC_NOMCLI) = 3500
            .ColWidth(COL_NC_CODITEM) = 2200
            .ColWidth(COL_NC_DESCITEM) = 3500
            .ColWidth(COL_NC_CANT) = 800
            .ColWidth(COL_NC_PRT) = 1000
            .ColWidth(COL_NC_DSCTO) = 1000
            .ColWidth(COL_NC_FVENCI) = 1000
            .ColWidth(COL_NC_FUP) = 1000
            .ColWidth(COL_NC_DIAS) = 800
            .ColWidth(COL_NC_APLINC) = 1000
            .ColWidth(COL_NC_RES) = 1200
         '  .ColWidth(13) = 1000
'         .ColHidden(6) = True
'         .ColHidden(13) = True
        .ColDataType(COL_NC_APLINC) = flexDTBoolean
        .ColDataType(COL_NC_CANT) = flexDTCurrency
        .ColDataType(COL_NC_PRT) = flexDTCurrency
        .ColDataType(COL_NC_DSCTO) = flexDTCurrency
        .ColDataType(COL_NC_FVENCI) = flexDTDate
        .ColDataType(COL_NC_FUP) = flexDTDate
        
        .ColFormat(COL_NC_CANT) = "#,0.00"
        .ColFormat(COL_NC_PRT) = "#,0.00"
        .ColFormat(COL_NC_DSCTO) = "#,0.00"
        
        GNPoneNumFila grd, False
    End With
End Sub

Private Function GrabarDevolucionxTransDscto() As Boolean
Dim i As Long, FilaIni As Integer, FilaFin As Integer
Dim CodCli As String, ref As String
Dim gc As GNComprobante
Dim bandsaldoFactura As Boolean
    On Error GoTo ErrTrap
    
    MensajeStatus "Generando notas de credito  x Dscto", vbHourglass
    bandsaldoFactura = False
    prg1.min = 0
    prg1.max = grd.Rows - 1
    CodCli = ""
    ref = ""
    FilaIni = 1
    For i = grd.FixedRows To grd.Rows - 1
        If Mid(grd.TextMatrix(i, COL_NC_RES), 1, 7) <> "Grabado" Then
            If grd.IsSubtotal(i) Then
                FilaFin = i - 1
                CrearComprobantexTransDscto FilaIni, FilaFin
                FilaIni = i + 1
                'Docs.AgregaFila
     '           Enc_Aceptar (i)
      '          GrabarTransacciones (i)
                mobjGNComp.CodClienteRef = grd.TextMatrix(i - 1, COL_NC_CODCLI)
                mobjGNComp.Nombre = grd.TextMatrix(i - 1, COL_NC_NOMCLI)
                mobjGNComp.numDocRef = grd.TextMatrix(i - 1, COL_NC_NUMTRANS)
                mobjGNComp.idTransFuente = grd.TextMatrix(i - 1, COL_NC_TRANSID)
                mobjGNComp.VerificaAsiento True, True
                
                mobjGNComp.Grabar True, True
                prg1.value = i
                    grd.TextMatrix(i - 1, COL_NC_RES) = "Grabado como " & mobjGNComp.CodTrans & mobjGNComp.numtrans
                Set mobjGNComp = Nothing
                'ITEMS.Limpiar
               ' Recargos.Limpiar
               ' Docs.Limpiar
                'CodCli = grd.ValueMatrix(i, 1)
               ' ref = grd.ValueMatrix(i, COL_REF)
            Else
                CodCli = ""
            End If
        End If
    Next
    GrabarDevolucionxTransDscto = True
    prg1.value = prg1.min
    Screen.MousePointer = vbNormal
'    GrabarxLoteCliente = mbooGrabado
    MensajeStatus "Listo ", vbNormal
    Exit Function
ErrTrap:
    MensajeStatus
    prg1.value = prg1.min
    Screen.MousePointer = vbNormal
    Select Case Err.Number
    Case ERR_DESCUADRADO, ERR_INTEGRIDAD
        'Si es que el usuario seleccionó 'No' en el cuadro de dialogo,
        'No hace nada
    Case Else
        DispErr
    End Select
    ITEMS.SetFocus  'Para que no se pierda el enfoque
    Exit Function
End Function

Private Sub CrearComprobantexTransDscto(fila As Integer, FilaFin As Integer)
    Dim id As Long, i As Long, j As Integer, numitem As Integer, ix As Long, ct As Currency
    Dim item As IVinventario
    Dim Valor As Currency
    Dim BandGraba As Boolean
    SacaDatosGnTrans (fcbTrans.KeyText)
'    CrearGnComprobante
   
            For j = fila To FilaFin
                If grd.ValueMatrix(j, COL_NC_APLINC) = -1 Then
                    Valor = Valor + grd.ValueMatrix(j, COL_NC_DSCTO) * grd.ValueMatrix(j, COL_NC_PRT)
                    BandGraba = True
                End If
            Next j
    If BandGraba Then
        If Len(fcbTrans.KeyText) = 0 Then
            MsgBox "No hay un tipo de transacción para crear"
        Else
            Set mobjGNComp = gobjMain.EmpresaActual.CreaGNComprobante(fcbTrans.KeyText)
        End If
            Set item = mobjGNComp.Empresa.RecuperaIVInventarioQuick("DESC")
            ix = mobjGNComp.AddIVKardex
            mobjGNComp.IVKardex(ix).CodBodega = mobjGNComp.GNTrans.CodBodegaPre
            mobjGNComp.IVKardex(ix).CodInventario = item.CodInventario
            mobjGNComp.IVKardex(ix).cantidad = 1
            mobjGNComp.IVKardex(ix).PrecioTotal = Valor
            mobjGNComp.IVKardex(ix).PrecioRealTotal = Valor
            mobjGNComp.IVKardex(ix).IVA = item.PorcentajeIVA
            mobjGNComp.IVKardex(ix).Orden = ix
                            ct = item.CostoDouble2(mobjGNComp.FechaTrans, _
                    Abs(Valor), _
                    mobjGNComp.TransID, _
                    mobjGNComp.HoraTrans)
            mobjGNComp.IVKardex(ix).CostoTotal = ct
            mobjGNComp.IVKardex(ix).CostoRealTotal = ct
            ix = mobjGNComp.AddPCKardex
            mobjGNComp.PCKardex(ix).codforma = fcbFormaCobro.KeyText
            mobjGNComp.PCKardex(ix).CodProvCli = grd.TextMatrix(fila, COL_NC_CODCLI)
            mobjGNComp.PCKardex(ix).Haber = Valor
    BandCargado = True
    End If
    BandGraba = False
End Sub

Private Sub CargaFCobro()
    Dim i As Long, v As Variant
    Dim s As String, cod As String, aux  As Integer, gt As GNTrans
    lstCobro.Clear
    v = gobjMain.EmpresaActual.ListaTSFormaCobroPago(True, True, False)
    For i = LBound(v, 2) To UBound(v, 2)
        lstCobro.AddItem v(0, i)        '& " " & v(1, i)
    Next i
End Sub

Private Function PreparaCodForma() As String
    Dim i As Long, s As String
    With lstCobro
        'Si está seleccionado solo una
        If lstCobro.SelCount = 1 Then
            For i = 0 To .ListCount - 1
                If .Selected(i) Then
                    s = .List(i)
                    Exit For
                End If
            Next i
        'Si está TODO o NINGUNO, no hay condición
        ElseIf (.SelCount < .ListCount) And (.SelCount > 0) Then
            For i = 0 To .ListCount - 1
                If .Selected(i) Then
                    s = s & .List(i) & ","
                End If
            Next i
            If Len(s) > 0 Then s = Left$(s, Len(s) - 1)    'Quita la ultima ", "
        End If
    End With
    PreparaCodForma = s
End Function

Public Sub RecuperaCodForma(ByVal s As String, lst As ListBox)
Dim Vector As Variant
Dim i As Integer, j As Integer, Selec As Integer
    If s <> "_VACIO_" Then
        Vector = Split(s, ",")
         Selec = UBound(Vector, 1)
         For i = 0 To Selec
            For j = 0 To lst.ListCount - 1
                If Trim(Vector(i)) = lst.List(j) Then
                    lst.Selected(j) = True
                End If
            Next j
         Next i
    End If
End Sub

