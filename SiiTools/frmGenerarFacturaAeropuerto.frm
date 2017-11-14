VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{50067EB3-D6AF-11D3-8297-000021C5085D}#1.0#0"; "NTextBox.ocx"
Begin VB.Form frmGenerarFacturaAeropuerto 
   Caption         =   "Generación de una factura para Aerolineas"
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
   Begin TabDlg.SSTab sst1 
      Height          =   8235
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   14835
      _ExtentX        =   26167
      _ExtentY        =   14526
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Parametros de Busqueda - F6"
      TabPicture(0)   =   "frmGenerarFacturaAeropuerto.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grd"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdBuscar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraFecha"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraGrupos"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Factura - F7"
      TabPicture(1)   =   "frmGenerarFacturaAeropuerto.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Recargos"
      Tab(1).Control(1)=   "ITEMS"
      Tab(1).Control(2)=   "fraEnc"
      Tab(1).Control(3)=   "grdItems"
      Tab(1).Control(4)=   "Docs"
      Tab(1).Control(5)=   "grdItems2"
      Tab(1).ControlCount=   6
      Begin VB.Frame Frame1 
         Caption         =   "Rango de Grupos Items 2"
         Height          =   1095
         Left            =   9480
         TabIndex        =   43
         Top             =   360
         Width           =   2355
         Begin VB.ComboBox cboGrupo2 
            Height          =   315
            Left            =   780
            Style           =   2  'Dropdown List
            TabIndex        =   44
            Top             =   300
            Width           =   1452
         End
         Begin FlexComboProy.FlexCombo fcbGrupoDesde2 
            Height          =   300
            Left            =   780
            TabIndex        =   45
            Top             =   660
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
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
         Begin VB.Label lblGrupo2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "D&esde"
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   720
            Width           =   465
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Grupo"
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   450
         End
      End
      Begin SiiToolsA.IVRPVT Recargos 
         Height          =   2235
         Left            =   -74820
         TabIndex        =   40
         Top             =   5580
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   3942
      End
      Begin VB.Frame fraGrupos 
         Caption         =   "Rango de Grupos Items 1"
         Height          =   1095
         Left            =   7080
         TabIndex        =   35
         Top             =   360
         Width           =   2355
         Begin VB.ComboBox cboGrupo 
            Height          =   315
            Left            =   780
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   300
            Width           =   1452
         End
         Begin FlexComboProy.FlexCombo fcbGrupoDesde 
            Height          =   300
            Left            =   780
            TabIndex        =   37
            Top             =   660
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Grupo"
            Height          =   195
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   450
         End
         Begin VB.Label lblGrupo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Desde"
            Height          =   195
            Left            =   120
            TabIndex        =   38
            Top             =   720
            Width           =   495
         End
      End
      Begin SiiToolsA.IVGNPVT ITEMS 
         Height          =   2595
         Left            =   -74940
         TabIndex        =   27
         Top             =   2100
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   4577
      End
      Begin VB.Frame fraEnc 
         Height          =   1575
         Left            =   -74940
         TabIndex        =   19
         Top             =   360
         Width           =   8175
         Begin VB.CommandButton cmdAceptar2 
            Caption         =   "Facturar"
            Height          =   375
            Left            =   5880
            TabIndex        =   49
            Top             =   1140
            Width           =   2115
         End
         Begin NTextBoxProy.NTextBox ntxCotizacion 
            Height          =   324
            Left            =   960
            TabIndex        =   8
            Top             =   1164
            Width           =   1452
            _ExtentX        =   2566
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
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "&Facturar"
            Height          =   375
            Left            =   3660
            TabIndex        =   12
            Top             =   1155
            Width           =   2175
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   510
            Left            =   3660
            MaxLength       =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            ToolTipText     =   "Descripción de la transacción"
            Top             =   600
            Width           =   4380
         End
         Begin MSComCtl2.DTPicker dtpFecha 
            Height          =   360
            Left            =   960
            TabIndex        =   6
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
            Format          =   105578497
            CurrentDate     =   37078
            MaxDate         =   73415
            MinDate         =   29221
         End
         Begin FlexComboProy.FlexCombo fcbResp 
            Height          =   336
            Left            =   6600
            TabIndex        =   10
            ToolTipText     =   "Responsable de la transacción"
            Top             =   240
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
         Begin NTextBoxProy.NTextBox ntxNumTrans 
            Height          =   360
            Left            =   4515
            TabIndex        =   9
            ToolTipText     =   "Número de la transacción"
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
            Text            =   "0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SeparadoPorComa =   0   'False
         End
         Begin FlexComboProy.FlexCombo fcbMoneda 
            Height          =   324
            Left            =   960
            TabIndex        =   7
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cod.Trans  "
            Height          =   195
            Left            =   2775
            TabIndex        =   26
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "&Responsable  "
            Height          =   195
            Left            =   5580
            TabIndex        =   25
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "C&otización  "
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   1230
            Width           =   810
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "&Descripción  "
            Height          =   195
            Left            =   2670
            TabIndex        =   23
            Top             =   600
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&Fecha Transaccion  "
            Height          =   195
            Left            =   1020
            TabIndex        =   22
            Top             =   240
            Width           =   1470
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "&Moneda  "
            Height          =   195
            Left            =   270
            TabIndex        =   21
            Top             =   840
            Width           =   675
         End
         Begin VB.Label lblCodTrans 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   360
            Left            =   3660
            TabIndex        =   20
            ToolTipText     =   "Código de la transacción"
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame fraFecha 
         Caption         =   "&Fecha (desde - hasta)"
         Height          =   1575
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   3375
         Begin MSComCtl2.DTPicker dtpFecha2 
            Height          =   330
            Left            =   1800
            TabIndex        =   2
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            _Version        =   393216
            Format          =   105578497
            CurrentDate     =   36902
         End
         Begin MSComCtl2.DTPicker dtpFecha1 
            Height          =   330
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            _Version        =   393216
            Format          =   105578497
            CurrentDate     =   36902
         End
         Begin FlexComboProy.FlexCombo fcbCliente 
            Height          =   330
            Left            =   120
            TabIndex        =   28
            Top             =   1020
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   582
            DispCol         =   1
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
         Begin VB.Label Label3 
            Caption         =   "Aerolinea:"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   780
            Width           =   915
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar - F5"
         Height          =   372
         Left            =   240
         TabIndex        =   4
         Top             =   1980
         Width           =   1212
      End
      Begin VB.Frame Frame5 
         Caption         =   "Configuracion &Transacciones"
         Height          =   1575
         Left            =   3540
         TabIndex        =   17
         Top             =   360
         Width           =   3495
         Begin FlexComboProy.FlexCombo fcbTransArribo 
            Height          =   345
            Left            =   1320
            TabIndex        =   3
            Top             =   360
            Width           =   2115
            _ExtentX        =   3731
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
         Begin FlexComboProy.FlexCombo fcbTransSalida 
            Height          =   345
            Left            =   1320
            TabIndex        =   31
            Top             =   720
            Width           =   2115
            _ExtentX        =   3731
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
         Begin FlexComboProy.FlexCombo fcbTrans 
            Height          =   345
            Left            =   1320
            TabIndex        =   32
            Top             =   1080
            Width           =   2115
            _ExtentX        =   3731
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
            Left            =   60
            TabIndex        =   34
            Top             =   1140
            Width           =   1035
         End
         Begin VB.Label Label10 
            Caption         =   "Salida de Vuelo"
            Height          =   255
            Left            =   60
            TabIndex        =   33
            Top             =   780
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Arribo de Vuelo"
            Height          =   255
            Left            =   60
            TabIndex        =   30
            Top             =   420
            Width           =   1215
         End
      End
      Begin VSFlex7LCtl.VSFlexGrid grd 
         Height          =   2175
         Left            =   120
         TabIndex        =   5
         Top             =   2460
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
      Begin VSFlex7LCtl.VSFlexGrid grdItems 
         Height          =   1455
         Left            =   -66660
         TabIndex        =   41
         Top             =   480
         Width           =   3915
         _cx             =   6906
         _cy             =   2566
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
      Begin SiiToolsA.PCDoc Docs 
         Height          =   2055
         Left            =   -69180
         TabIndex        =   42
         Top             =   5400
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   3625
         ProvCliVisible  =   0   'False
         PorCobrar       =   0   'False
      End
      Begin VSFlex7LCtl.VSFlexGrid grdItems2 
         Height          =   1455
         Left            =   -62640
         TabIndex        =   48
         Top             =   480
         Width           =   2355
         _cx             =   4154
         _cy             =   2566
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
   End
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   852
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   8520
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5565
      Width           =   8520
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar -F3"
         Height          =   372
         Left            =   2880
         TabIndex        =   13
         Top             =   360
         Width           =   1332
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   4320
         TabIndex        =   14
         Top             =   360
         Width           =   1212
      End
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   60
         Width           =   8280
         _ExtentX        =   14605
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "frmGenerarFacturaAeropuerto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Constantes para las columnas
Private Const COL_NUMFILA = 0
Private Const COL_TID = 6
Private Const COL_FECHA = 2
Private Const COL_CODASIENTO = 3
Private Const COL_CODTRANS = 4
Private Const COL_NUMTRANS = 5
Private Const COL_NUMDOCREF = 6
Private Const COL_NOMBRE = 7
Private Const COL_DESC = 8
Private Const COL_CENTROCOSTO = 9
Private Const COL_ESTADO = 10
Private Const COL_RESULTADO = 11

Private Const COL_GNTAR_HORA = 3
Private Const COL_GNTAR_CODTRANS = 4
Private Const COL_GNTAR_NUMTRANS = 5
Private Const COL_GNTAR_NOMBRE = 1
Private Const COL_GNTAR_DESC = 7
Private Const COL_GNTAR_CENTROCOSTO = 8
Private Const COL_GNTAR_VILU = 9
Private Const COL_GNTAR_VATE = 10
Private Const COL_GNTAR_VESTA = 11
Private Const COL_GNTAR_FACTURADO = 12
Private Const COL_GNTAR_ADULTO = 13
Private Const COL_GNTAR_MEDIO = 14
Private Const COL_GNTAR_INFAN = 15
Private Const COL_GNTAR_ESTA = 16
Private Const COL_GNTAR_ILU = 17

Private Const COL_FC_NUM = 0
Private Const COL_FC_CODG1 = 1
Private Const COL_FC_CODITEM = 2
Private Const COL_FC_CODALT = 3    'Diego 06/12/2000
Private Const COL_FC_DESC = 4
Private Const COL_FC_EXIST = 5
Private Const COL_FC_CANT = 6
Private Const COL_FC_CU = 7
Private Const COL_FC_CT = 8
Private Const COL_FC_PU = 9
Private Const COL_FC_PUIVA = 10
Private Const COL_FC_PT = 11
Private Const COL_FC_PTIVA = 12
Private Const COL_FC_VALIVA = 13
Private Const COL_FC_PORIVA = 14
Private Const COL_FC_DSCTO = 15  '*** ANGEL 01/Nov/2001 Columna para Descuento por Item
Private Const MAXLEN_NOTA As Integer = 80


Private Const ROW_FC_SEGURIDAD = 1
Private Const ROW_FC_USOAEROPUERTO = 2

Private Const ROW_FC2_ATERRIZAJE = 1
Private Const ROW_FC2_ESTACIONAMIENTO = 2
Private Const ROW_FC2_ILUMINACION = 3

Private Const COL_AUX_CODIGO = 1
Private Const COL_AUX_CANTIDAD = 2
Private Const COL_AUX_PRECIO = 3
Private Const COL_AUX_IVA = 4


Private mProcesando As Boolean
Private mCancelado As Boolean
Private mVerificado As Boolean
Private WithEvents mobjGNComp As GNComprobante
Attribute mobjGNComp.VB_VarHelpID = -1
Private mbooGrabado As Boolean
Private YaImprimio As Boolean
Private mobjImp As Object
Private mFactura1 As Boolean

Public Sub Inicio()
    Dim i As Integer
    On Error GoTo ErrTrap
    sst1.Tab = 0
    mFactura1 = True
    
    For i = 1 To IVGRUPO_MAX
        cboGrupo.AddItem gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(i), i - 1
    Next i
    
    For i = 1 To IVGRUPO_MAX
        cboGrupo2.AddItem gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(i), i - 1
    Next i
    
    
    cboGrupo.ListIndex = gobjMain.objCondicion.numGrupo - 1
    cboGrupo2.ListIndex = gobjMain.objCondicion.nivel - 1
    
    Me.Show
    Me.ZOrder
    dtpFecha1.value = gobjMain.EmpresaActual.GNOpcion.FechaInicio
    dtpFecha2.value = Date
    CargarEncabezado
    CargaTrans
    CargaCliente
    CargarDatos
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub CargaTrans()
    'Carga la lista de transacción
    fcbTrans.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False)
    fcbTransArribo.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "IV")
    fcbTransSalida.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "IV")
    
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransFacturacionAerolineas")) > 0 Then
        fcbTrans.KeyText = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransFacturacionAerolineas")
    End If
    
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ArriboFacturacionAerolineas")) > 0 Then
        fcbTransArribo.KeyText = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ArriboFacturacionAerolineas")
    End If
    
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SalidaFacturacionAerolineas")) > 0 Then
        fcbTransSalida.KeyText = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SalidaFacturacionAerolineas")
    End If
    
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("GrupoFacturacionAerolineas")) > 0 Then
        cboGrupo.ListIndex = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("GrupoFacturacionAerolineas")
    End If
    
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("GrupoDesdeFacturacionAerolineas")) > 0 Then
        fcbGrupoDesde.KeyText = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("GrupoDesdeFacturacionAerolineas")
    End If
    
    
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("GrupoFacturacionAerolineas2")) > 0 Then
        cboGrupo2.ListIndex = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("GrupoFacturacionAerolineas2")
    End If
    
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("GrupoDesdeFacturacionAerolineas2")) > 0 Then
        fcbGrupoDesde2.KeyText = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("GrupoDesdeFacturacionAerolineas2")
    End If
    
    
End Sub



Private Sub cmdAceptar_Click()
    mFactura1 = True
    If Not mProcesando Then
        'Si no hay transacciones
        If grd.Rows <= grd.FixedRows Then
            MsgBox "No hay ningúna transacción para procesar."
            Exit Sub
        End If
    'Crea transacciiones
    If CrearTransacciones Then
'        Habilitar True
        mbooGrabado = False
    Else
 '       Habilitar False
        mbooGrabado = True
    End If
        
        Enc_Aceptar
        
        If GenerarFactura() Then
            cmdCancelar.SetFocus
        End If
    End If
End Sub

Private Function GenerarAsiento() As Boolean
    Dim s As String, tid As Long, i As Long, x As Single
    Dim gnc As GNComprobante, cambiado As Boolean
    
    On Error GoTo ErrTrap

    
'    s = "Este proceso modificará los asientos de la transacción seleccionada." & vbCr & vbCr
'    s = s & "Está seguro que desea proceder?"
'    If MsgBox(s, vbYesNo + vbQuestion) <> vbYes Then Exit Function
    
    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdBuscar.Enabled = False
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    
    ' Proceeder a Generar un solo asiento por todas las transacciones que Estan.

    If Not mobjGNComp Is Nothing Then
        mobjGNComp.Generar1AsientoxLote
        'Asiento.VisualizaDesdeObjeto
    End If
    
    Screen.MousePointer = 0
    GenerarAsiento = Not mCancelado
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    DispErr
salida:
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    Exit Function
End Function




Private Sub cmdAceptar2_Click()
    mFactura1 = False
    If Not mProcesando Then
        'Si no hay transacciones
        If grd.Rows <= grd.FixedRows Then
            MsgBox "No hay ningúna transacción para procesar."
            Exit Sub
        End If
    'Crea transacciiones
    If CrearTransacciones Then
'        Habilitar True
        mbooGrabado = False
    Else
 '       Habilitar False
        mbooGrabado = True
    End If
        
        Enc_Aceptar
        
        If GenerarFactura() Then
            cmdCancelar.SetFocus
        End If
    End If
End Sub

Private Sub cmdBuscar_Click()
    Dim v As Variant, obj As Object, s As String
    Dim numGrupo As Integer, NumGrupoDesde  As String
    On Error GoTo ErrTrap
    
    
    If Len(fcbCliente.KeyText) = 0 Then
        MsgBox "Seleccione cliente.", vbInformation
        fcbCliente.SetFocus
        Exit Sub
    End If
    
    
    If Len(fcbTrans.Text) = 0 Then
        MsgBox "Seleccione solo un tipo de transacción", vbInformation
        fcbTrans.SetFocus
        Exit Sub
    End If
    
    With gobjMain.objCondicion
        .fecha1 = dtpFecha1.value
        .fecha2 = dtpFecha2.value
        .CodTrans = "'" & fcbTransArribo.Text & "','" & fcbTransSalida.Text & "'"
        .CodPC1 = fcbCliente.KeyText
        .CodPC2 = .CodPC1
        'Estados no incluye anulados
        .EstadoBool(ESTADO_NOAPROBADO) = True
        .EstadoBool(ESTADO_APROBADO) = True
        .EstadoBool(ESTADO_DESPACHADO) = True
        .EstadoBool(ESTADO_ANULADO) = False
        .numGrupo = cboGrupo.ListIndex + 1
        .nivel = cboGrupo2.ListIndex + 1
    End With
    'Set obj = gobjMain.EmpresaActual.ConsGNTransAerolineas(True) 'Ascendente
    numGrupo = IIf(Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiAereo_num_grupo")) = 0, 1, gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiAereo_num_grupo"))
    NumGrupoDesde = IIf(Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiAereo_num_grupo")) = 0, 1, gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiAereo_num_grupoDesde"))

    Set obj = gobjMain.EmpresaActual.ConsGNTransAerolineas(numGrupo, NumGrupoDesde)
    grd.Redraw = flexRDNone
    grd.Rows = 1
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
    VisualizaTotalAereo
    grd.AutoSize 0, grd.Cols - 1
    grd.Redraw = True
    cmdAceptar.Enabled = True
    mVerificado = True
    
            s = fcbTrans.KeyText
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransFacturacionAerolineas", s
            s = fcbTransArribo.KeyText
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "ArriboFacturacionAerolineas", s
            s = fcbTransSalida.KeyText
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "SalidaFacturacionAerolineas", s
    
            s = cboGrupo.ListIndex
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "GrupoFacturacionAerolineas", s
    
            s = fcbGrupoDesde.KeyText
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "GrupoDesdeFacturacionAerolineas", s
    
            s = cboGrupo2.ListIndex
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "GrupoFacturacionAerolineas2", s
    
            s = fcbGrupoDesde2.KeyText
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "GrupoDesdeFacturacionAerolineas2", s
    
    
    
        'Graba en la base
    gobjMain.EmpresaActual.GNOpcion.Grabar

    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub ConfigCols()
    Dim rsItem As Recordset, sql As String, s As String, CantItem As Integer, i As Integer
    Dim numGrupo As Integer, NumGrupoDesde  As String
    With grd
        '.FormatString = "^#|tid|<Fecha|<Asiento|<Trans|<#|<#Ref.|<Nombre|<Descripción|<C.Costo|<Estado|<Resultado"
                s = "^#|<Aereolínea|<Fecha|<Hora|<Trans|>#Trans.|<TransID|<Descripción|" & _
                        "<Matricula|>V. Ilum|>V. Aterr|>V. Esta|^Facturado"
                        
            numGrupo = IIf(Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiAereo_num_grupo")) = 0, 1, gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiAereo_num_grupo"))
            NumGrupoDesde = IIf(Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiAereo_num_grupo")) = 0, 1, gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiAereo_num_grupoDesde"))
                        
                        
        sql = " select Codinventario from ivinventario ivi"
        sql = sql & " inner join ivgrupo" & numGrupo & " ivg on ivg.idgrupo" & numGrupo & "=ivi.idgrupo" & numGrupo
        sql = sql & " where codgrupo" & numGrupo & "='" & NumGrupoDesde & "'"
        Set rsItem = gobjMain.EmpresaActual.OpenRecordset(sql)
        
        CantItem = rsItem.RecordCount
        rsItem.MoveFirst
        
        For i = 1 To CantItem
            s = s & "|>" & rsItem.Fields("CodInventario")
            rsItem.MoveNext
        Next i


        
        .FormatString = s

        .ColHidden(COL_TID) = True
        .ColHidden(COL_GNTAR_DESC) = True
        .ColHidden(COL_GNTAR_FACTURADO) = True
''''        .ColHidden(COL_GNTAR_PESO) = True
        
        
        .ColFormat(COL_GNTAR_HORA) = "HH:mm"
        .ColDataType(COL_FECHA) = flexDTDate    '*** MAKOTO 14/ago/2000 para que ordene bien por fecha

        
        VisualizaTotalAereo
        
         GNPoneNumFila grd, False

    End With
    
    With grdItems
        s = "^#|<Codigo del Item|<Cantidad|Precio|IVA|ccccccc"
        .FormatString = s
        GNPoneNumFila grdItems, False
    End With
    With grdItems2
        s = "^#|<Codigo del Item|<Cantidad|Precio|IVA|ccccccc"
        .FormatString = s
        GNPoneNumFila grdItems2, False
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
    If Grabar Then
        'cmdGrabar.Enabled = False
        If Not mFactura1 Then
            Limpiar
'            cmdAceptar.Enabled = False
'            cmdAceptar2.Enabled = False
            cmdGrabar.Enabled = False
        Else
'            cmdAceptar2.Enabled = True
'            cmdAceptar.Enabled = False
        End If
    Else
        'cmdGrabar.Enabled = True
    End If
End Sub




Private Sub fcbTrans_BeforeSelect(ByVal Row As Long, Cancel As Boolean)
'    SacaTransAsientoGnTrans fcbTrans.Text
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF3
        cmdGrabar_Click
        KeyCode = 0
    Case vbKeyF5
        If sst1.Tab = 0 Then cmdBuscar_Click
        KeyCode = 0
    Case vbKeyF6
        sst1.Tab = 0
        dtpFecha1.SetFocus
        KeyCode = 0
    Case vbKeyF7
        sst1.Tab = 1
        dtpFecha.SetFocus
        KeyCode = 0
    Case vbKeyF8
        If sst1.Tab = 1 Then cmdAceptar_Click
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
    
    sst1.Move 0, sst1.Top, Me.ScaleWidth, Me.ScaleHeight - pic1.Height - 300
    
    With grd
        .Width = Me.ScaleWidth - 200
        .Height = Me.ScaleHeight - .Top - pic1.Height - 380
    End With
    
    With ITEMS
        .Width = Me.ScaleWidth - 200
        .Height = 4000
    End With
    
    With Recargos
        .Left = ITEMS.Left
        .Top = ITEMS.Top + 4100
        .Width = 6000
        .Height = 2500
    End With

    With Docs
        .Left = 6200
        .Top = ITEMS.Top + 4100
        .Width = 8300
        .Height = 2500
    End With

    
    prg1.Width = Me.ScaleWidth - (prg1.Left * 2)
        
End Sub




Private Sub Form_Unload(Cancel As Integer)
    Set mobjImp = Nothing
    Set mobjGNComp = Nothing
End Sub

Private Sub grd_LostFocus()
    If sst1.Tab = 0 Then
        sst1.Tab = 1
    End If
End Sub



Private Sub sst1_Click(PreviousTab As Integer)
    '*** Para evitar error de ciclo infinito
 
    On Error GoTo ErrTrap
    Select Case sst1.Tab
    Case 0          'Parametros de Busqueda
    
    Case 1
        'Transaccion de Asiento
        cmdAceptar.Caption = "Facturar - " & fcbGrupoDesde.KeyText
        cmdAceptar2.Caption = "Facturar - " & fcbGrupoDesde2.KeyText
        CalculaCantidad
        If lblCodTrans.Caption <> fcbTrans.KeyText Then
            lblCodTrans.Caption = fcbTrans.KeyText
            SacaDatosGnTrans (lblCodTrans.Caption)
            CrearGnComprobante
        End If
        PoneDescripcion
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
    txtDescripcion.Text = "Factura de Tasas Aeroportuarias..."
End Sub


Private Sub SacaDatosGnTrans(ByVal CodTrans As String)
    Dim gnt As GNTrans
    Set gnt = gobjMain.EmpresaActual.RecuperaGNTrans(CodTrans)
    If Not gnt Is Nothing Then
        fcbResp.KeyText = gnt.CodResponsablePre
    End If
End Sub

Private Sub SacaTransAsientoGnTrans(ByVal CodTrans As String)
    Dim gnt As GNTrans
    Set gnt = gobjMain.EmpresaActual.RecuperaGNTrans(CodTrans)
    If Not gnt Is Nothing Then
        fcbTransArribo.KeyText = gnt.TransAsiento
    End If
End Sub

Private Sub CrearGnComprobante()
    'Eliminar el que haya tenido
    Set mobjGNComp = Nothing
    'crear el comprobante para luego grabar
    If Len(lblCodTrans.Caption) = 0 Then
        MsgBox "No hay un tipo de transacción para crear"
    Else
        Set mobjGNComp = gobjMain.EmpresaActual.CreaGNComprobante(lblCodTrans.Caption)
        Set ITEMS.GNComprobante = mobjGNComp
    End If
    
End Sub

Private Sub Enc_Aceptar()
    
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

    'Si es hora automático, graba con la hora del momento de grabación
    'Solo  cuando no es modificacion
    
    
    If Not (mobjGNComp Is Nothing) Then
        mobjGNComp.CodClienteRef = fcbCliente.KeyText
        mobjGNComp.nombre = fcbCliente.Text
        mobjGNComp.FechaTrans = dtpFecha.value
        mobjGNComp.numtrans = ntxNumTrans.value
        mobjGNComp.CodResponsable = fcbResp.KeyText
        mobjGNComp.CodMoneda = fcbMoneda.Text
        mobjGNComp.Cotizacion("") = ntxCotizacion.value
        mobjGNComp.Descripcion = Trim$(txtDescripcion.Text)
        If mobjGNComp.GNTrans.HoraAuto And mobjGNComp.EsNuevo = True Then
            mobjGNComp.HoraTrans = Time
        End If
    End If
End Sub


Private Function Grabar() As Boolean
    On Error GoTo ErrTrap
    
    If mobjGNComp Is Nothing Then Exit Function
        
    Enc_Aceptar
    GrabarTransacciones
    
    Grabar = mbooGrabado
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
    ITEMS.SetFocus  'Para que no se pierda el enfoque
    Exit Function
End Function


Private Sub PoneDescripcion()
    Dim gnt As GNTrans, fdesde As Date, fhasta As Date
    'pone descripcion en la transaccion
    Set gnt = gobjMain.EmpresaActual.RecuperaGNTrans(fcbTrans.Text)
    If Not gnt Is Nothing Then
        fdesde = (dtpFecha1.value)
        fhasta = (dtpFecha2.value)
        If mFactura1 Then
            txtDescripcion.Text = "Facturacion de " & fcbGrupoDesde.KeyText & " de los Vuelos  desde " & fdesde & " hasta " & fhasta
        Else
            txtDescripcion.Text = "Facturacion de " & fcbGrupoDesde2.KeyText & " de los Vuelos  desde " & fdesde & " hasta " & fhasta
        End If
    Else
        txtDescripcion.Text = "Error inesperado en la transaccion seleccionada"
    End If
    ntxNumTrans = 0
    Set gnt = Nothing
End Sub

Public Sub Limpiar()
    Dim i As Long
    
    With grd
        For i = .FixedRows To .Rows - 1
            .RowData(i) = 0
        Next i
        .Rows = .FixedRows
        .Rows = .FixedRows
    
    End With
'    ITEMS.Limpiar
    Enc_Limpiar
End Sub

Private Sub Enc_Limpiar()
    lblCodTrans.Caption = fcbTrans.KeyText
    SacaDatosGnTrans (lblCodTrans.Caption)
    CrearGnComprobante
    PoneDescripcion
End Sub

Private Sub CargaCliente()
    fcbCliente.SetData gobjMain.EmpresaActual.ListaPCProvCli(False, True, False)
End Sub

Private Sub VisualizaTotalAereo()
    Dim i As Long
    Dim bc1 As Long, bc2 As Long, bc3 As Long, fc As Long
    Dim mUltimaColumna As Long
    PrepararColor bc1, bc2, bc3, fc
    
    With grd
        'Subtotal por tipo de cuenta
        mUltimaColumna = COL_GNTAR_CODTRANS
        .SubTotal flexSTSum, mUltimaColumna, COL_GNTAR_VILU, , bc2, , , " ", mUltimaColumna, True
        .SubTotal flexSTSum, mUltimaColumna, COL_GNTAR_VATE, , bc2, , , " ", mUltimaColumna, True
        .SubTotal flexSTSum, mUltimaColumna, COL_GNTAR_VESTA, , bc2, , , " ", mUltimaColumna, True
        .SubTotal flexSTSum, mUltimaColumna, COL_GNTAR_ADULTO, , bc2, , , " ", mUltimaColumna, True
        .SubTotal flexSTSum, mUltimaColumna, COL_GNTAR_MEDIO, , bc2, , , " ", mUltimaColumna, True
        .SubTotal flexSTSum, mUltimaColumna, COL_GNTAR_INFAN, , bc2, , , " ", mUltimaColumna, True
        .SubTotal flexSTSum, mUltimaColumna, COL_GNTAR_ESTA, , bc2, , , " ", mUltimaColumna, True
        .SubTotal flexSTSum, mUltimaColumna, COL_GNTAR_ILU, , bc2, , , " ", mUltimaColumna, True
        

        
        'Total general
        .SubTotal flexSTSum, -1, COL_GNTAR_VILU, , bc3, fc, , " ", -1, True
        .SubTotal flexSTSum, -1, COL_GNTAR_VATE, , bc3, fc, , " ", -1, True
        .SubTotal flexSTSum, -1, COL_GNTAR_VESTA, , bc3, fc, , " ", -1, True
        .SubTotal flexSTSum, -1, COL_GNTAR_ADULTO, , bc3, fc, , " ", -1, True
        .SubTotal flexSTSum, -1, COL_GNTAR_MEDIO, , bc3, fc, , " ", -1, True
        .SubTotal flexSTSum, -1, COL_GNTAR_INFAN, , bc3, fc, , " ", -1, True
        .SubTotal flexSTSum, -1, COL_GNTAR_ESTA, , bc3, fc, , " ", -1, True
        .SubTotal flexSTSum, -1, COL_GNTAR_ILU, , bc3, fc, , " ", -1, True
    End With
End Sub

Private Sub PrepararColor( _
                ByRef bc1 As Long, _
                ByRef bc2 As Long, _
                ByRef bc3 As Long, _
                ByRef fc As Long)
                
    'Color de visualización normal
        bc1 = RGB(220, 220, 255)        'Subtotal-2
        bc2 = grd.BackColorFixed        'Subtotal
        bc3 = grd.BackColorSel          'Total general
        fc = vbYellow                   'Letra en Total general
        
End Sub


Private Function GenerarFactura() As Boolean
    Dim s As String, tid As Long, i As Long, x As Single
    Dim gnc As GNComprobante, cambiado As Boolean
    
    On Error GoTo ErrTrap
    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdBuscar.Enabled = False
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    ProcesarDatos
    ActualizaTotalCobrar False
    CargaPorCobrar
    Docs.SetFocus
    Docs.Aceptar
    Screen.MousePointer = 0
    GenerarFactura = Not mCancelado
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    DispErr
salida:
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    Exit Function
End Function



Private Sub cboGrupo_Change()
 Dim Numg As Integer
    On Error GoTo ErrTrap
    If cboGrupo.ListIndex < 0 Then Exit Sub
    Numg = cboGrupo.ListIndex + 1
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub cboGrupo_Click()
    Dim Numg As Integer
    On Error GoTo ErrTrap
    If cboGrupo.ListIndex < 0 Then Exit Sub

    'MensajeStatus MSG_PREPARA, vbHourglass
    lblGrupo.Caption = cboGrupo.Text
    Numg = cboGrupo.ListIndex + 1
    fcbGrupoDesde.SetData gobjMain.EmpresaActual.ListaIVGrupo(Numg, False, False)
    fcbGrupoDesde.KeyText = ""
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    Exit Sub
End Sub


Private Sub cboGrupo2_Change()
 Dim Numg As Integer
    On Error GoTo ErrTrap
    If cboGrupo2.ListIndex < 0 Then Exit Sub
    Numg = cboGrupo2.ListIndex + 1
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub cboGrupo2_Click()
    Dim Numg As Integer
    On Error GoTo ErrTrap
    If cboGrupo.ListIndex < 0 Then Exit Sub

    'MensajeStatus MSG_PREPARA, vbHourglass
    lblGrupo2.Caption = cboGrupo2.Text
    Numg = cboGrupo2.ListIndex + 1
    fcbGrupoDesde2.SetData gobjMain.EmpresaActual.ListaIVGrupo(Numg, False, False)
    fcbGrupoDesde2.KeyText = ""
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    Exit Sub
End Sub


Private Sub CalculaCantidad()
    Dim i As Integer, tasaAterizaje As Currency
    Dim tasaSeguridad As Currency, tasaInstalaciones As Currency
    Dim tasaEstacionamiento As Currency, tasaIluminacion As Currency
    tasaAterizaje = 0
    tasaEstacionamiento = 0
    tasaIluminacion = 0
    For i = 1 To grd.Rows - 1
        If Not grd.IsSubtotal(i) Then
            If grd.TextMatrix(i, COL_GNTAR_CODTRANS) = fcbTransArribo.KeyText Then
                'tasaAterizaje = tasaAterizaje + 1
                tasaAterizaje = tasaAterizaje + grd.ValueMatrix(i, COL_GNTAR_VATE)
                'tasaSeguridad = tasaSeguridad + grd.ValueMatrix(i, COL_GNTAR_ADULTO) + grd.ValueMatrix(i, COL_GNTAR_MEDIO)
            Else
                tasaSeguridad = tasaSeguridad + grd.ValueMatrix(i, COL_GNTAR_ADULTO) + grd.ValueMatrix(i, COL_GNTAR_MEDIO)
                tasaEstacionamiento = tasaEstacionamiento + (grd.ValueMatrix(i, COL_GNTAR_ESTA) * grd.ValueMatrix(i, COL_GNTAR_VESTA))
            End If
            tasaIluminacion = tasaIluminacion + (grd.ValueMatrix(i, COL_GNTAR_ILU) * grd.ValueMatrix(i, COL_GNTAR_VILU))
            
'''        Else
'''            If i <> grd.Rows - 1 Then
'''                tasaAterizaje = grd.ValueMatrix(i, COL_GNTAR_VATE)
'''            End If
''''         MsgBox "hola"
        End If
    Next i
    'SEGURIDAD
    grdItems.TextMatrix(ROW_FC_SEGURIDAD, COL_AUX_CANTIDAD) = tasaSeguridad
    'TASA USO AEROPUERTO
    grdItems.TextMatrix(ROW_FC_USOAEROPUERTO, COL_AUX_CANTIDAD) = tasaSeguridad
    
    'aterrizaje
    grdItems2.TextMatrix(ROW_FC2_ATERRIZAJE, COL_AUX_CANTIDAD) = 1
    grdItems2.TextMatrix(ROW_FC2_ATERRIZAJE, COL_AUX_PRECIO) = tasaAterizaje
    'estacionamiento
    grdItems2.TextMatrix(ROW_FC2_ESTACIONAMIENTO, COL_AUX_CANTIDAD) = 1
    grdItems2.TextMatrix(ROW_FC2_ESTACIONAMIENTO, COL_AUX_PRECIO) = tasaEstacionamiento
    'iluminacion
    grdItems2.TextMatrix(ROW_FC2_ILUMINACION, COL_AUX_CANTIDAD) = 1
    grdItems2.TextMatrix(ROW_FC2_ILUMINACION, COL_AUX_PRECIO) = tasaIluminacion

End Sub

Private Sub CargarDatos()
    'Llena los datos de cabecera
    CargarEncabezado
    CargaItems
    
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


Private Sub ProcesarDatos()
    Dim ix As Long, ivk As IVKardex, dif As Currency
    Dim i As Long, signo As Integer, cant As String
    Dim iv As IVinventario, c As Currency
    Dim sql As String, rs As Recordset
    
    If mFactura1 Then
        For i = 1 To grdItems.Rows - 1
            c = 0
            dif = 0
            cant = 0
           
                With mobjGNComp
                    ix = .AddIVKardex
                    Set ivk = .IVKardex(i)
                    .IVKardex(ix).cantidad = grdItems.TextMatrix(i, COL_AUX_CANTIDAD) * -1
                    
                    .IVKardex(ix).CodBodega = ivk.CodBodega
                    .IVKardex(ix).CodInventario = grdItems.TextMatrix(i, COL_AUX_CODIGO) ' ivk.CodInventario
                    .IVKardex(ix).IVA = grdItems.ValueMatrix(i, COL_AUX_IVA) ' ivk.CodInventario
                    
                    
                    'Calcula el costo
                    Set iv = .Empresa.RecuperaIVInventario(ivk.CodInventario)
                    
                    sql = " select  top 1 g.fechatrans,g.horatrans "
                    sql = sql & " from ivinventario ivi inner join  ivkardex ivk"
                    sql = sql & " inner join gncomprobante g on g.transid=ivk.transid"
                    sql = sql & " on ivk.idinventario=ivi.idinventario"
                    sql = sql & " where ivi.codinventario='" & ivk.CodInventario & "' and Cantidad>0"
                    sql = sql & " and g.estado<>3 and CostoTotal<>0 "
                    sql = sql & " order by g.fechatrans,g.horatrans"
                    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                    If rs.RecordCount = 1 Then
                        'obtiene el costo del primer ingreso
                        c = iv.CostoDouble2(rs.Fields("fechatrans"), _
                                             cant, _
                                             .TransID, _
                                             DateAdd("s", 1, rs.Fields("horatrans")))
                    Else
                        c = 0
                    End If
                    
                
                    'Si el costo calculado está en otra moneda, convierte en moneda de trans.
                    If .CodMoneda <> iv.CodMoneda Then
                        c = c * .Cotizacion(iv.CodMoneda) / .Cotizacion(" ")
                    End If
                    
                    .IVKardex(ix).CostoTotal = c * grdItems.TextMatrix(i, COL_AUX_CANTIDAD) * -1
                    .IVKardex(ix).PrecioTotal = grdItems.TextMatrix(i, COL_AUX_PRECIO) * grdItems.TextMatrix(i, COL_AUX_CANTIDAD) * -1
                End With
           
           
            
        Next i
    Else
        For i = 1 To grdItems2.Rows - 1
            c = 0
            dif = 0
            cant = 0
            With mobjGNComp
                ix = .AddIVKardex
                Set ivk = .IVKardex(i)
                .IVKardex(ix).cantidad = grdItems2.TextMatrix(i, COL_AUX_CANTIDAD) * -1
                
                .IVKardex(ix).CodBodega = ivk.CodBodega
                .IVKardex(ix).CodInventario = grdItems2.TextMatrix(i, COL_AUX_CODIGO) ' ivk.CodInventario
                .IVKardex(ix).IVA = grdItems2.ValueMatrix(i, COL_AUX_IVA) ' ivk.CodInventario
                
                
                'Calcula el costo
                Set iv = .Empresa.RecuperaIVInventario(ivk.CodInventario)
                
                sql = " select  top 1 g.fechatrans,g.horatrans "
                sql = sql & " from ivinventario ivi inner join  ivkardex ivk"
                sql = sql & " inner join gncomprobante g on g.transid=ivk.transid"
                sql = sql & " on ivk.idinventario=ivi.idinventario"
                sql = sql & " where ivi.codinventario='" & ivk.CodInventario & "' and Cantidad>0"
                sql = sql & " and g.estado<>3 and CostoTotal<>0 "
                sql = sql & " order by g.fechatrans,g.horatrans"
                Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                If rs.RecordCount = 1 Then
                    'obtiene el costo del primer ingreso
                    c = iv.CostoDouble2(rs.Fields("fechatrans"), _
                                         cant, _
                                         .TransID, _
                                         DateAdd("s", 1, rs.Fields("horatrans")))
                Else
                    c = 0
                End If
                
            
                'Si el costo calculado está en otra moneda, convierte en moneda de trans.
                If .CodMoneda <> iv.CodMoneda Then
                    c = c * .Cotizacion(iv.CodMoneda) / .Cotizacion(" ")
                End If
                
                .IVKardex(ix).CostoTotal = c * grdItems2.TextMatrix(i, COL_AUX_CANTIDAD) * -1
                .IVKardex(ix).PrecioTotal = grdItems2.TextMatrix(i, COL_AUX_PRECIO) * grdItems2.TextMatrix(i, COL_AUX_CANTIDAD) * -1
            End With
        Next i
    End If
    ITEMS.VisualizaDesdeObjeto
    ActualizaTotalCobrar False
    Recargos.Aceptar
    Recargos.Refresh
    Recargos.Refresh
End Sub


Private Sub CargaItems()
    Dim rsItem As Recordset, sql As String, s As String, CantItem As Integer, i As Integer
    Dim numGrupo As Integer, NumGrupoDesde  As String
    Dim GrupoFacturacion As Integer, NumGrupoDesdeFacturacion  As String
    Dim GrupoFacturacion2 As Integer, NumGrupoDesdeFacturacion2  As String
    With grdItems
        .Clear
        numGrupo = IIf(Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiAereo_num_grupo")) = 0, 1, gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiAereo_num_grupo"))
        NumGrupoDesde = IIf(Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiAereo_num_grupo")) = 0, 1, gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SiiAereo_num_grupoDesde"))
                        
        GrupoFacturacion = IIf(Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("GrupoFacturacionAerolineas")) = 0, 0, gobjMain.EmpresaActual.GNOpcion.ObtenerValor("GrupoFacturacionAerolineas")) + 1
        NumGrupoDesdeFacturacion = IIf(Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("GrupoDesdeFacturacionAerolineas")) = 0, 1, gobjMain.EmpresaActual.GNOpcion.ObtenerValor("GrupoDesdeFacturacionAerolineas"))
                        
                        
        
        sql = " SELECT CodInventario,0, Precio1,porcentajeIVA FROM IVInventario "
        sql = sql & " inner join ivgrupo" & GrupoFacturacion
        
        sql = sql & " on IVInventario.idgrupo" & GrupoFacturacion
        sql = sql & "  = ivgrupo" & GrupoFacturacion & ".idgrupo" & GrupoFacturacion
        sql = sql & " where bandventa=1 "
        sql = sql & " and Codgrupo" & GrupoFacturacion
        sql = sql & " ='" & NumGrupoDesdeFacturacion & "'"
         sql = sql & "  ORDER BY CodInventario"
       
        
        
        Set rsItem = gobjMain.EmpresaActual.OpenRecordset(sql)
        
        CantItem = rsItem.RecordCount
        If CantItem > 0 Then
            rsItem.MoveFirst
        End If

        grdItems.Refresh
        grdItems.LoadArray MiGetRows(rsItem)
        
        .FormatString = s

         GNPoneNumFila grdItems, False
    End With
    
    With grdItems2
        .Clear
        GrupoFacturacion2 = IIf(Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("GrupoFacturacionAerolineas2")) = 0, 0, gobjMain.EmpresaActual.GNOpcion.ObtenerValor("GrupoFacturacionAerolineas2")) + 1
        NumGrupoDesdeFacturacion2 = IIf(Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("GrupoDesdeFacturacionAerolineas2")) = 0, 1, gobjMain.EmpresaActual.GNOpcion.ObtenerValor("GrupoDesdeFacturacionAerolineas2"))
                        
                        
        
        sql = " SELECT CodInventario,0, Precio1,porcentajeIVA FROM IVInventario "
        sql = sql & " inner join ivgrupo" & GrupoFacturacion2
        
        sql = sql & " on IVInventario.idgrupo" & GrupoFacturacion2
        sql = sql & "  = ivgrupo" & GrupoFacturacion & ".idgrupo" & GrupoFacturacion2
        sql = sql & " where bandventa=1 "
        sql = sql & " and Codgrupo" & GrupoFacturacion2
        sql = sql & " ='" & NumGrupoDesdeFacturacion2 & "'"
         sql = sql & "  ORDER BY CodInventario"
       
        
        
        Set rsItem = gobjMain.EmpresaActual.OpenRecordset(sql)
        
        CantItem = rsItem.RecordCount
        If CantItem > 0 Then
            rsItem.MoveFirst
        End If

        grdItems2.Refresh
        grdItems2.LoadArray MiGetRows(rsItem)
        
        .FormatString = s

         GNPoneNumFila grdItems2, False


    End With
End Sub


Private Sub GrabarTransacciones()
    Dim trans_conteo As String, proceso As Integer, msg As String
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
    
    ITEMS.EliminaFilasIncompletas
    If ITEMS.GNComprobante.CountIVKardex = 0 Then
        MsgBox "No hay ningúna fila para grabar.", vbInformation
        Exit Sub
    End If
        
'    'Confirmación
'    If MsgBox("Está seguro que desea comenzar el proceso?", _
'                vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    
    
    MensajeStatus "Grabando Factura Aerolineas", vbHourglass
    'Graba los ajustes de inventario
    ITEMS.Aceptar
    If ITEMS.GNComprobante.CountIVKardex > 0 Then
        proceso = 2
        With ITEMS.GNComprobante
            .CodResponsable = fcbResp.KeyText
            .CodMoneda = fcbMoneda.KeyText
            
            .GeneraAsiento
            .GeneraAsientoPresupuesto
            'Verificación de datos
            .VerificaDatos
            .Grabar False, False
        End With
    End If
    
    
    
    MensajeStatus "Grabando Fctura Aerolinea", vbHourglass
    'Graba la transacción usada para el conteo físico
    If Not mFactura1 Then
        CambiaEstadoVuelos
    End If
    If Not YaImprimio Then
        Select Case mobjGNComp.GNTrans.ImprimeComprobante
            Case "S", "P"                            ' si esta seleccionado que imprime Siempre, o preguntando pero la libreria saca el mensaje de aprobacion de impresion
                Imprime = True
            Case "N"   ' no imprime automatico
                'Imprime = False
        End Select
        If Imprime Then
            Me.ZOrder
            Imprimir True
        End If
    End If
        
    
    MensajeStatus
'    MsgBox "Proceso terminado con éxito", vbOKOnly + vbInformation
ITEMS.Limpiar
Recargos.Limpiar
Docs.Limpiar
    mbooGrabado = True
    Exit Sub
    
ErrTrap:
    MensajeStatus
    DispErr
    Exit Sub
End Sub


Private Sub ActualizaTotalCobrar(ByVal refresh_item As Boolean)
    Dim t As Currency, anticipos As Currency
    
'    If refresh_item Then    ITEMS.Refresh
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


Private Sub Imprimir(Directo)
    On Error GoTo ErrTrap
    'If Not Directo Then
    
    If Not ImprimeTrans(mobjGNComp, mobjImp) Then       '*** MAKOTO 11/nov/00
        Me.Show             '*** MAKOTO 11/nov/00 Para que no se pierda el enfoque
    End If
    YaImprimio = True  'Es un nuevo para imprimir
    Exit Sub
ErrTrap:
    DispErr
    Me.Show             '*** MAKOTO 11/nov/00 Para que no se pierda el enfoque
    Exit Sub
End Sub


Public Function ImprimeTrans(ByVal gc As GNComprobante, ByRef objImp As Object) As Boolean
    Dim crear As Boolean
    On Error GoTo ErrTrap

    'Si no tiene TransID quere decir que no está grabada
    If (gc.TransID = 0) Or gc.Modificado Then
        MsgBox MSGERR_NOGRABADO, vbInformation
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
    
    MensajeStatus MSG_PREPARA, vbHourglass
    'objImp.PrintTrans gobjMain.EmpresaActual, true, 1, 0, "", 0, gc
    'jeaa 23/11/2006
    objImp.PrintTrans gobjMain.EmpresaActual, IIf(gc.GNTrans.ImprimeComprobante = "S", True, False), 1, 0, "", 0, gc
    MensajeStatus
    'jeaa 30/09/04
    gc.CambiaEstadoImpresion
    ImprimeTrans = True
    Exit Function
ErrTrap:
    MensajeStatus
    Select Case Err.Number
    Case ERR_NOIMPRIME, ERR_NOIMPRIME2, ERR_NOIMPRIME3, ERR_NOHAYCODIGO
        DispErr
    Case Else
        
        MsgBox MSGERR_NOIMPRIME2, vbInformation
        
    End Select
    ImprimeTrans = False
    Exit Function
End Function


Private Sub CambiaEstadoVuelos()
    Dim i As Integer, sql As String, rs As Recordset
    For i = 1 To grd.Rows - 1
        If Not grd.IsSubtotal(i) Then
            If grd.TextMatrix(i, COL_GNTAR_FACTURADO) = "0" Then
                sql = " UPDATE Gncomprobante set Estado1=1 Where transid=" & grd.ValueMatrix(i, COL_TID)
                Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            End If
        End If
    Next i
End Sub
