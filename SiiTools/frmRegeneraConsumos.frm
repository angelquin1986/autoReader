VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRegeneraConsumos 
   Caption         =   "Regeneración de Recetas"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6585
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   6585
   WindowState     =   2  'Maximized
   Begin VB.CheckBox CHKBAJA 
      Caption         =   "Regenerar Baja y Transferencia"
      Height          =   255
      Left            =   6180
      TabIndex        =   31
      Top             =   840
      Width           =   4095
   End
   Begin VSFlex7LCtl.VSFlexGrid grdVenta 
      Height          =   2715
      Left            =   120
      TabIndex        =   19
      Top             =   5340
      Visible         =   0   'False
      Width           =   6375
      _cx             =   11245
      _cy             =   4789
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
   Begin VB.Frame fraitem 
      Caption         =   "    Items"
      Height          =   675
      Left            =   6128
      TabIndex        =   26
      Top             =   120
      Width           =   5052
      Begin FlexComboProy.FlexCombo fcbDesde2 
         Height          =   315
         Left            =   840
         TabIndex        =   27
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
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
      Begin FlexComboProy.FlexCombo fcbHasta2 
         Height          =   315
         Left            =   3225
         TabIndex        =   28
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
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
         Caption         =   "Desde"
         Height          =   252
         Left            =   240
         TabIndex        =   30
         Top             =   240
         Width           =   612
      End
      Begin VB.Label Label7 
         Caption         =   "Hasta"
         Height          =   252
         Left            =   2760
         TabIndex        =   29
         Top             =   240
         Width           =   612
      End
   End
   Begin VB.CheckBox chkTodo 
      Caption         =   "&Regenerar todo sin verificar"
      Enabled         =   0   'False
      Height          =   192
      Left            =   3960
      TabIndex        =   15
      Top             =   1440
      Width           =   3252
   End
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   852
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   6585
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3855
      Width           =   6585
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Proceder"
         Enabled         =   0   'False
         Height          =   372
         Left            =   1728
         TabIndex        =   13
         Top             =   0
         Width           =   1212
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   4968
         TabIndex        =   12
         Top             =   0
         Width           =   1212
      End
      Begin VB.CommandButton cmdVerificar 
         Caption         =   "&Verificar"
         Enabled         =   0   'False
         Height          =   372
         Left            =   288
         TabIndex        =   11
         Top             =   0
         Width           =   1212
      End
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   540
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   1932
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   6372
      _cx             =   11239
      _cy             =   3408
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
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   372
      Left            =   1704
      TabIndex        =   8
      Top             =   1320
      Width           =   1212
   End
   Begin VB.Frame fraFecha 
      Caption         =   "&Fecha (desde - hasta)"
      Height          =   1092
      Left            =   402
      TabIndex        =   0
      Top             =   120
      Width           =   1932
      Begin MSComCtl2.DTPicker dtpFecha1 
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1692
         _ExtentX        =   2990
         _ExtentY        =   529
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
         Format          =   106692609
         CurrentDate     =   36348
      End
      Begin MSComCtl2.DTPicker dtpFecha2 
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1692
         _ExtentX        =   2990
         _ExtentY        =   529
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
         Format          =   106692609
         CurrentDate     =   36348
      End
   End
   Begin VB.Frame fraCodTrans 
      Caption         =   "Cod.&Trans."
      Height          =   1092
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   1932
      Begin VB.CheckBox chkNoAprobadas 
         Caption         =   "Solo no aprobados"
         Height          =   255
         Left            =   150
         TabIndex        =   16
         Top             =   780
         Width           =   1665
      End
      Begin FlexComboProy.FlexCombo fcbTrans 
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1635
         _ExtentX        =   2884
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
   End
   Begin VB.Frame fraNumTrans 
      Caption         =   "# T&rans. (desde - hasta)"
      Height          =   1092
      Left            =   4200
      TabIndex        =   5
      Top             =   120
      Width           =   1932
      Begin VB.TextBox txtNumTrans1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   360
         TabIndex        =   6
         Top             =   280
         Width           =   1212
      End
      Begin VB.TextBox txtNumTrans2 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   360
         TabIndex        =   7
         Top             =   640
         Width           =   1212
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdReceta 
      Height          =   1935
      Left            =   60
      TabIndex        =   17
      Top             =   5700
      Visible         =   0   'False
      Width           =   6375
      _cx             =   11239
      _cy             =   3408
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
      Begin VB.Label Label2 
         Caption         =   "grdreceta"
         Height          =   255
         Left            =   840
         TabIndex        =   22
         Top             =   0
         Width           =   975
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdItems 
      Height          =   1935
      Left            =   1080
      TabIndex        =   18
      Top             =   6120
      Visible         =   0   'False
      Width           =   6375
      _cx             =   11239
      _cy             =   3408
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
      Begin VB.Label Label3 
         Caption         =   "grdItems"
         Height          =   255
         Left            =   960
         TabIndex        =   23
         Top             =   0
         Width           =   975
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdIvk 
      Height          =   5235
      Left            =   6600
      TabIndex        =   20
      Top             =   1800
      Width           =   6375
      _cx             =   11245
      _cy             =   9234
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
   Begin VSFlex7LCtl.VSFlexGrid grdReceta1 
      Height          =   1935
      Left            =   120
      TabIndex        =   21
      Top             =   5340
      Visible         =   0   'False
      Width           =   6375
      _cx             =   11239
      _cy             =   3408
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
      Begin VB.Label Label4 
         Caption         =   "grdRecetas1"
         Height          =   255
         Left            =   840
         TabIndex        =   24
         Top             =   0
         Width           =   975
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdIvkn 
      Height          =   1995
      Left            =   6900
      TabIndex        =   25
      Top             =   7920
      Width           =   6375
      _cx             =   11245
      _cy             =   3519
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
Attribute VB_Name = "frmRegeneraConsumos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Constantes para las columnas
Private Const COL_NUMFILA = 0
Private Const COL_TID = 1
Private Const COL_FECHA = 2
Private Const COL_CODASIENTO = 3
Private Const COL_CODTRANS = 4
Private Const COL_NUMTRANS = 5
Private Const COL_NUMDOCREF = 6     '*** MAKOTO 07/feb/01 Agregado
Private Const COL_NOMBRE = 7        '*** MAKOTO 07/feb/01 Agregado
Private Const COL_DESC = 8
Private Const COL_CENTROCOSTO = 9
Private Const COL_ESTADO = 10
Private Const COL_RESULTADO = 11

Private Const TIPORECETA = 4
Private Const COL_VENTA_ID = 1
Private Const COL_VENTA_IDINV = 2
Private Const COL_VENTA_CANT = 5
Private Const COL_VENTA_TIPO = 6
Private Const COL_VENTA_IDIDPADRE = 7

Private Const COL_ITEM_ID = 1
Private Const COL_ITEM_IDINV = 2
Private Const COL_ITEM_CANT = 5
Private Const COL_ITEM_TIPO = 6
Private Const COL_ITEM_IDIDPADRE = 7



Private Const COL_RECETA_IDINV = 1
Private Const COL_RECETA_CANT = 4



Private Const MSG_NG = "Receta incorrecta."
Private mProcesando As Boolean
Private mCancelado As Boolean
Private mVerificado As Boolean
Private num_fila_trans As Long
Private num_fila_itemVenta As Long
Private num_fila_Receta As Long
Private num_fila_IVkitem As Long
Private mobjGNCompAux As GNComprobante
Private mobjGNCompTransf As GNComprobante


Public Sub Inicio()
    Dim i As Integer
    On Error GoTo ErrTrap
    
    Me.Show
    Me.ZOrder
    dtpFecha1.value = gobjMain.EmpresaActual.GNOpcion.FechaInicio
    dtpFecha2.value = Date
    CargaTrans
    CargaItemsCombo 'AUC  17/05/2010
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub CargaTrans()
    'Carga la lista de transacción
    fcbTrans.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False)
End Sub



Private Sub cmdAceptar_Click()
    'Si no hay transacciones
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna transacción para procesar."
        Exit Sub
    End If
    
    If dtpFecha1 < gobjMain.EmpresaActual.GNOpcion.FechaLimiteDesde Then
        MsgBox "La Rango de Fecha de regeneración es menor a la Fecha Limite Aceptable  ", vbExclamation
        Exit Sub
    End If

    
    If RegenerarReceta(False, (chkTodo.value = vbChecked)) Then
        cmdAceptar.Enabled = True
        cmdAceptar.SetFocus
        mVerificado = True
    End If
    
    
'    If RegenerarAsiento(False, ) Then
'        cmdCancelar.SetFocus
'    End If
End Sub

'Private Function RegenerarAsiento(bandVerificar As Boolean, bandTodo As Boolean) As Boolean
'    Dim s As String, tid As Long, i As Long, x As Single
'    Dim gnc As GNComprobante, cambiado As Boolean
'
'    On Error GoTo ErrTrap
'
'    'Si no es solo verificacion, confirma
'    If Not bandVerificar Then
'        s = "Este proceso modificará los asientos de la transacción seleccionada." & vbCr & vbCr
'        s = s & "Está seguro que desea proceder?"
'        If MsgBox(s, vbYesNo + vbQuestion) <> vbYes Then Exit Function
'    End If
'
'    mProcesando = True
'    mCancelado = False
'    frmMain.mnuFile.Enabled = False
'    cmdVerificar.Enabled = False
'    cmdBuscar.Enabled = False
'    Screen.MousePointer = vbHourglass
'    prg1.Min = 0
'    prg1.max = grd.Rows - 1
'
'    For i = grd.FixedRows To grd.Rows - 1
'        DoEvents
'        If mCancelado Then
'            MsgBox "El proceso fue cancelado.", vbInformation
'            Exit For
'        End If
'
'        prg1.value = i
'        grd.Row = i
'        x = grd.CellTop                 'Para visualizar la celda actual
'
'        'Si es verificación, procesa todas las filas sino solo las que tengan "Asiento incorrecto."
'        If (grd.TextMatrix(i, COL_RESULTADO) = MSG_NG) Or bandVerificar Or bandTodo Then
'
'            tid = grd.ValueMatrix(i, COL_TID)
'            grd.TextMatrix(i, COL_RESULTADO) = "Verificando..."
'            grd.Refresh
'
'            'Recupera la transaccion
'            Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
'            If Not (gnc Is Nothing) Then
'                'Si la transacción no está anulada
'                If gnc.Estado <> ESTADO_ANULADO Then
'
'                    'Forzar recuperar todos los datos de transacción para que no se pierdan al grabar de nuveo
'                    gnc.RecuperaDetalleTodo
'
'                    'Recalcula costo de los items
'                    If RegenerarAsientoSub(gnc, cambiado) Then
'                        'Si está cambiado algo o está forzado regenerar todo
'                        If cambiado Or bandTodo Then
'                            'Si no es solo verificacion
'                            If (Not bandVerificar) Or bandTodo Then
'                                grd.TextMatrix(i, COL_RESULTADO) = "Grabando..."
'                                grd.Refresh
'
'                                'Graba la transacción
'                                gnc.Grabar False, False
'                                grd.TextMatrix(i, COL_RESULTADO) = "Actualizado."
'
'                            'Si es solo verificacion
'                            Else
'                                grd.TextMatrix(i, COL_RESULTADO) = MSG_NG
'                            End If
'                        Else
'                            'Si no está cambiado no graba
'                            grd.TextMatrix(i, COL_RESULTADO) = "OK."
'                        End If
'                    Else
'                        grd.TextMatrix(i, COL_RESULTADO) = "Falló al regenerar."
'                    End If
'                Else
'                    'Si está anulada
'                    grd.TextMatrix(i, COL_RESULTADO) = "Anulado."
'                End If
'            Else
'                grd.TextMatrix(i, COL_RESULTADO) = "No pudo recuperar la transación."
'            End If
'        End If
'    Next i
'
'    Screen.MousePointer = 0
'    RegenerarAsiento = Not mCancelado
'    GoTo salida
'ErrTrap:
'    Screen.MousePointer = 0
'    DispErr
'salida:
'    mProcesando = False
'    frmMain.mnuFile.Enabled = True
'    cmdVerificar.Enabled = True
'    cmdBuscar.Enabled = True
'    prg1.value = prg1.Min
'    Exit Function
'End Function


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

Private Sub cmdBuscar_Click()
    Dim v As Variant, obj As Object
    On Error GoTo ErrTrap
    
    grdItems.Rows = 1
    grdReceta.Rows = 1
    grdReceta1.Rows = 1
    grdIvk.Rows = 1
    grdIvkn.Rows = 1
    grdVenta.Rows = 1
    grdItems.Refresh
    grdReceta.Refresh
    grdReceta1.Refresh
    
    grdIvk.Refresh
    grdIvkn.Refresh
    grdVenta.Refresh
    
    
    With gobjMain.objCondicion
        .fecha1 = dtpFecha1.value
        .fecha2 = dtpFecha2.value
        .CodTrans = fcbTrans.Text
        .NumTrans1 = Val(txtNumTrans1.Text)
        .NumTrans2 = Val(txtNumTrans2.Text)
        
        'Estados no incluye anulados
        If chkNoAprobadas.value = vbChecked Then
            .EstadoBool(ESTADO_NOAPROBADO) = True
            .EstadoBool(ESTADO_APROBADO) = False
            .EstadoBool(ESTADO_DESPACHADO) = False
            .EstadoBool(ESTADO_ANULADO) = False
        Else
            .EstadoBool(ESTADO_NOAPROBADO) = True
            .EstadoBool(ESTADO_APROBADO) = True
            .EstadoBool(ESTADO_DESPACHADO) = True
            .EstadoBool(ESTADO_ANULADO) = False
        End If
        'AUC para filtrar por item
            .CodItem1 = Trim$(fcbDesde2.Text)
            .CodItem2 = Trim$(fcbHasta2.Text)
    End With
    Set obj = gobjMain.EmpresaActual.ConsGNTrans2(True) 'Ascendente     '*** MAKOTO 20/oct/00
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
    cmdVerificar.Enabled = True
    cmdVerificar.SetFocus
    cmdAceptar.Enabled = False
    chkTodo.Enabled = True
    mVerificado = False
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub ConfigCols()
    With grd
        .FormatString = "^#|tid|<Fecha|<Asiento|<Trans|<#|<#Ref.|<Nombre|<Descripción|<C.Costo|<Estado|<Resultado"
        .ColHidden(COL_NUMFILA) = False
        .ColHidden(COL_TID) = True
        .ColHidden(COL_FECHA) = False
        .ColHidden(COL_CODASIENTO) = True
        .ColHidden(COL_CODTRANS) = False
        .ColHidden(COL_NUMTRANS) = False
        .ColHidden(COL_NUMDOCREF) = True
        .ColHidden(COL_NOMBRE) = False      'True
        .ColHidden(COL_DESC) = False
        .ColHidden(COL_CENTROCOSTO) = True
        .ColHidden(COL_ESTADO) = True
        
        .ColDataType(COL_FECHA) = flexDTDate    '*** MAKOTO 14/ago/2000 para que ordene bien por fecha
        
        GNPoneNumFila grd, False
        .AutoSize 0, grd.Cols - 1
        
        .ColWidth(COL_NUMTRANS) = 500
        .ColWidth(COL_NOMBRE) = 1400
        .ColWidth(COL_DESC) = 2400
        .ColWidth(COL_RESULTADO) = 2000
    End With
    
    With grdVenta
        .FormatString = "^#|ID|idInventario|<Cod.Inventario|<Descripcion|>Cantidad|>Tipo|>IdPadre|<Resultado"
        GNPoneNumFila grd, False
        .AutoSize 0, .Cols - 1
        .ColWidth(1) = 700
        .ColWidth(2) = 700
        .ColWidth(3) = 800
        .ColWidth(4) = 2500
        .ColWidth(5) = 500
        .ColWidth(6) = 500
        .ColWidth(7) = 500
        .ColWidth(8) = 800
    End With
    
    With grdIvk
        .FormatString = "^#|ID|>Tamanio|<Cod. Tamaño|>Procesado|<Cod. Procesado|>Cant.|>Ticket|>IdPadre|<Resultado"
        GNPoneNumFila grd, False
        .AutoSize 0, .Cols - 1
        .ColWidth(1) = 700
        .ColWidth(2) = 850
        .ColWidth(3) = 1500
        .ColWidth(4) = 850
        .ColWidth(5) = 1500
        .ColWidth(6) = 700
        .ColWidth(7) = 700
        .ColWidth(8) = 800
        .ColWidth(9) = 800
    End With
    
    
    
    With grdItems
        .FormatString = "^#|ID|idInventario|<Cod.Inventario|<Descripcion|>Cantidad|>Tipo|>IdPadre|<Resultado"
        GNPoneNumFila grd, False
        .AutoSize 0, .Cols - 1
        .ColWidth(1) = 700
        .ColWidth(2) = 700
        .ColWidth(3) = 800
        .ColWidth(4) = 2500
        .ColWidth(5) = 500
        .ColWidth(6) = 500
        .ColWidth(7) = 500
        .ColWidth(8) = 800
    End With
    
    With grdReceta
        .FormatString = "^#|ID|idInventario|<Cod.Inventario|<Descripcion|>Cantidad|>Tipo|>IdPadre|<Resultado"
        GNPoneNumFila grd, False
        .AutoSize 0, .Cols - 1
        .ColWidth(1) = 700
        .ColWidth(2) = 700
        .ColWidth(3) = 800
        .ColWidth(4) = 2500
        .ColWidth(5) = 500
        .ColWidth(6) = 500
        .ColWidth(7) = 500
        .ColWidth(8) = 800
    End With
    
    With grdReceta1
        .FormatString = "^#|ID|idInventario|<Cod.Inventario|<Descripcion|>Cantidad|>Tipo|>IdPadre|<Resultado"
        GNPoneNumFila grd, False
        .AutoSize 0, .Cols - 1
        .ColWidth(1) = 700
        .ColWidth(2) = 700
        .ColWidth(3) = 800
        .ColWidth(4) = 2500
        .ColWidth(5) = 500
        .ColWidth(6) = 500
        .ColWidth(7) = 500
        .ColWidth(8) = 800
    End With
    
    
    With grdIvkn
        .FormatString = "^#|ID|idInventario|<Cod.Inventario|<Descripcion|>Cantidad|>Tipo|>IdPadre|<Resultado"
        GNPoneNumFila grd, False
        .AutoSize 0, .Cols - 1
        .ColWidth(1) = 700
        .ColWidth(2) = 700
        .ColWidth(3) = 800
        .ColWidth(4) = 2500
        .ColWidth(5) = 500
        .ColWidth(6) = 500
        .ColWidth(7) = 500
        .ColWidth(8) = 800
    End With
    
End Sub

Private Sub cmdCancelar_Click()
    If mProcesando Then
        mCancelado = True
    Else
        Unload Me
    End If
End Sub


Private Sub cmdVerificar_Click()
    'Si no hay transacciones
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna transacción para verificar."
        Exit Sub
    End If
    
    If dtpFecha1 < gobjMain.EmpresaActual.GNOpcion.FechaLimiteDesde Then
        MsgBox "La Rango de Fecha de regeneración es menor a la Fecha Limite Aceptable  ", vbExclamation
        Exit Sub
    End If

'    If RegenerarAsiento(True, False) Then
'        cmdAceptar.Enabled = True
'        cmdAceptar.SetFocus
'        mVerificado = True
'    End If
    
        If RegenerarConsumo(True, False) Then
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
            mVerificado = True
        End If
    
    
End Sub


Private Sub chkTodo_Click()
    If chkTodo.value = vbChecked Then
        cmdVerificar.Enabled = False
        cmdAceptar.Enabled = (grd.Rows > grd.FixedRows)
    Else
        cmdVerificar.Enabled = Not mVerificado
        cmdAceptar.Enabled = mVerificado
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF9
        cmdAceptar_Click
        KeyCode = 0
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
'    grd.Move 0, grd.Top, Me.ScaleWidth, (Me.ScaleHeight - grd.Top - pic1.Height - 80)
    grd.Move 0, grd.Top, Me.ScaleWidth / 2, (Me.ScaleHeight - grd.Top - pic1.Height - 80) / 2
    grdIvk.Move grd.Left + grd.Width, grd.Top, Me.ScaleWidth / 2, (Me.ScaleHeight - grd.Top - pic1.Height - 80) / 2
    grdIvkn.Move grd.Left + grd.Width, grdIvk.Top + grdIvk.Height, Me.ScaleWidth / 2, (Me.ScaleHeight - grd.Top - pic1.Height - 80) / 2
    grdItems.Visible = True
    grdItems.Move 0, grd.Top + grd.Height, Me.ScaleWidth / 2, (Me.ScaleHeight - grd.Top - pic1.Height - 80) / 2
'
'    grdIvkn.Move 0, grd.Top + grd.Height + grdIvk.Height, Me.ScaleWidth / 2, (Me.ScaleHeight - grd.Top - pic1.Height - 80) / 4
'
'
'    grdItems.Move grd.Left + grd.Width, grdVenta.Top + grdVenta.Height, Me.ScaleWidth / 2, (Me.ScaleHeight - grd.Top - pic1.Height - 80) / 4
'    grdReceta.Move grd.Left + grd.Width, grdItems.Top + grdItems.Height, Me.ScaleWidth / 2, (Me.ScaleHeight - grd.Top - pic1.Height - 80) / 4
'    grdReceta1.Move grd.Left + grd.Width, grdReceta.Top + grdReceta.Height, Me.ScaleWidth / 2, (Me.ScaleHeight - grd.Top - pic1.Height - 80) / 4
    prg1.Width = Me.ScaleWidth - (prg1.Left * 2)
End Sub


Private Sub txtNumTrans1_KeyPress(KeyAscii As Integer)
    'Acepta solo numericos
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNumTrans2_KeyPress(KeyAscii As Integer)
    'Acepta solo numericos
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Function RegenerarReceta(bandVerificar As Boolean, BandTodo As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, x As Single
    Dim gnc As GNComprobante, cambiado As Boolean
    On Error GoTo ErrTrap

    'Si no es solo verificacion, confirma
    If Not bandVerificar Then
        s = "Este proceso modificará las cantidades y costos de las recetas de la transacción seleccionada." & vbCr & vbCr
        s = s & "Está seguro que desea proceder?"
        If MsgBox(s, vbYesNo + vbQuestion) <> vbYes Then Exit Function
    End If
    
    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdVerificar.Enabled = False
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
        
        'Si es verificación, procesa todas las filas sino solo las que tengan "Asiento incorrecto."
        If (grd.TextMatrix(i, COL_RESULTADO) = MSG_NG) Or bandVerificar Or BandTodo Then
        
            tid = grd.ValueMatrix(i, COL_TID)
            grd.TextMatrix(i, COL_RESULTADO) = "Verificando..."
            grd.Refresh
            
            'Recupera la transaccion
            Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
            If Not (gnc Is Nothing) Then
                'Si la transacción no está anulada
                If gnc.Estado <> ESTADO_ANULADO Then
                    
                    'Forzar recuperar todos los datos de transacción para que no se pierdan al grabar de nuveo
                    gnc.RecuperaDetalleTodo
                    
                    cargaItemsKardex gnc, i
                    num_fila_trans = i
                    cargaItemsVenta gnc, i
                   
                    'Recalcula costo de los items
'                    If RegenerarAsientoSub(gnc, cambiado) Then
'                        'Si está cambiado algo o está forzado regenerar todo
'                        If cambiado Or bandTodo Then
'                            'Si no es solo verificacion
'                            If (Not bandVerificar) Or bandTodo Then
'                                grd.TextMatrix(i, COL_RESULTADO) = "Grabando..."
'                                grd.Refresh
'
'                                'Graba la transacción
'                                gnc.Grabar False, False
'                                grd.TextMatrix(i, COL_RESULTADO) = "Actualizado."
'
'                            'Si es solo verificacion
'                            Else
'                                grd.TextMatrix(i, COL_RESULTADO) = MSG_NG
'                            End If
'                        Else
'                            'Si no está cambiado no graba
'                            grd.TextMatrix(i, COL_RESULTADO) = "OK."
'                        End If
'                    Else
'                        grd.TextMatrix(i, COL_RESULTADO) = "Falló al regenerar."
'                    End If
                Else
                    'Si está anulada
                    grd.TextMatrix(i, COL_RESULTADO) = "Anulado."
                End If
            Else
                grd.TextMatrix(i, COL_RESULTADO) = "No pudo recuperar la transación."
            End If
        End If
    Next i
    
    Screen.MousePointer = 0
    RegenerarReceta = Not mCancelado
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    DispErr
salida:
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdVerificar.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    Exit Function
End Function



Private Sub cargaItemsVenta(ByRef gnc As GNComprobante, ByVal i As Long)
    Dim j As Long, ivk As IVKardex, item As IVinventario
    grdVenta.Rows = 1
    'carga la el detalle transaccion
    For j = 1 To gnc.CountIVKardex
        With grdVenta
            Set ivk = gnc.IVKardex(j)
            Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
            If ivk.PrecioReal <> 0 Or ivk.PrecioRealTotal <> 0 Then
                .AddItem i & vbTab & ivk.id & vbTab & ivk.idinventario & vbTab & ivk.CodInventario & vbTab & item.Descripcion & vbTab & ivk.cantidad * -1 & vbTab & item.Tipo & vbTab & ivk.IdPadre
                If "4" Then
                    grdIvkn.AddItem i & vbTab & "0" & vbTab & ivk.idinventario & vbTab & ivk.CodInventario & vbTab & item.Descripcion & vbTab & ivk.cantidad & vbTab & item.Tipo & vbTab & ivk.IdPadre
'                    cargaItemsKardexNuevo grdIvk, gnc
                End If
                
            Set item = Nothing
                
                
            End If
            Set item = Nothing
        End With
    Next j
    grdItems.col = COL_ITEM_TIPO
    grdItems.Sort = flexSortGenericDescending
    grdVenta.Refresh
    CargaItems gnc, i
 '   VerificaItemPadre gnc, i
End Sub

Private Sub CargaItems(ByRef gnc As GNComprobante, ByVal i As Long)
    Dim j As Long, k As Long, item As IVinventario
    Dim obj As Object, v As Variant
    On Error GoTo ErrTrap
    grdItems.Rows = 1
    grdReceta.Rows = 1
    grdReceta1.Rows = 1
    grdItems.Refresh
    grdReceta.Refresh
    grdReceta1.Refresh

    'carga la el detalle transaccion
    
        For j = 1 To grdVenta.Rows - 1
        With grdItems
            If grdVenta.TextMatrix(j, COL_VENTA_TIPO) = "0" Then
                ActualizaIdPadre_EnVenta grdVenta.TextMatrix(j, COL_ITEM_ID), j, gnc, i
            ElseIf grdVenta.TextMatrix(j, COL_VENTA_TIPO) = "4" Then
                If grdVenta.ValueMatrix(j, COL_VENTA_IDIDPADRE) = "0" Then
                    Set obj = gobjMain.EmpresaActual.RecuperaRecetaReproceso(grdVenta.TextMatrix(j, COL_VENTA_IDINV), gnc.TransID, True)
                Else
                    'Es preparacion recupara los subitems
                    If grdVenta.TextMatrix(j, COL_VENTA_IDIDPADRE) <> 0 Then
                        'recupera los que ya tiene idpadre
                        Set obj = gobjMain.EmpresaActual.RecuperaRecetaReproceso1(grdVenta.TextMatrix(j, COL_VENTA_IDINV), gnc.TransID, grdVenta.TextMatrix(j, COL_VENTA_IDIDPADRE))
                    Else
                        Set obj = gobjMain.EmpresaActual.RecuperaRecetaReproceso(grdVenta.TextMatrix(j, COL_VENTA_IDINV), gnc.TransID, False)
                    End If
                End If
                v = MiGetRows(obj)
                .Redraw = flexRDNone
                If IsEmpty(v) Then
                'AUC 09/03/07 si esta vacio quiere decir que no tiene subitems entoces se regenera
                   If Not CorrigeItemsFaltantes(grdVenta.ValueMatrix(j, COL_ITEM_IDINV), grdVenta.ValueMatrix(j, COL_ITEM_CANT), grdVenta.ValueMatrix(j, COL_VENTA_IDIDPADRE), gnc) Then
'                        MsgBox MSG_NG & " " & gnc.CodTrans & " " & gnc.numtrans & " " & grdVenta.TextMatrix(j, 4) & " " & Chr(13) & "Revise la Configuraciòn del Item ", vbInformation
                   End If
                Else
                    .LoadArray v
                End If
                .Redraw = flexRDDirect
                grdItems.Refresh
                grdItems.col = COL_ITEM_TIPO
                grdItems.Sort = flexSortGenericDescending

                If grdItems.Rows > 1 Then
                
                For k = 1 To grdItems.Rows - 1
                    With grdItems
                            Set item = gnc.Empresa.RecuperaIVInventario(grdItems.ValueMatrix(k, COL_ITEM_IDINV))
                            If grdItems.ValueMatrix(k, COL_ITEM_TIPO) <> TIPORECETA And (grdItems.ValueMatrix(k, COL_ITEM_TIPO) <> TIPORECETA - 1) Then
                                grdIvkn.AddItem i & vbTab & "0" & vbTab & grdItems.ValueMatrix(k, COL_ITEM_IDINV) & vbTab & item.CodInventario & vbTab & item.Descripcion & vbTab & grdItems.ValueMatrix(k, COL_ITEM_CANT) * -1 & vbTab & item.Tipo & vbTab & grdItems.ValueMatrix(k, COL_ITEM_IDIDPADRE)
                            ElseIf grdItems.ValueMatrix(k, COL_ITEM_TIPO) = 3 Then
                                grdIvkn.AddItem i & vbTab & "0" & vbTab & grdItems.ValueMatrix(k, COL_ITEM_IDINV) & vbTab & item.Descripcion & vbTab & item.Descripcion & vbTab & grdItems.ValueMatrix(k, COL_ITEM_CANT) * -1 & vbTab & item.Tipo & vbTab & grdItems.ValueMatrix(k, COL_ITEM_IDIDPADRE)
                                ActualizaCantidad_EnDetalleItemsTransformacion grdVenta.TextMatrix(k, COL_ITEM_ID), k, gnc, grdVenta.TextMatrix(j, COL_VENTA_CANT)
                            Else
                                grdIvkn.AddItem i & vbTab & "0" & vbTab & grdItems.ValueMatrix(k, COL_ITEM_IDINV) & vbTab & item.Descripcion & vbTab & item.Descripcion & vbTab & grdItems.ValueMatrix(k, COL_ITEM_CANT) * -1 & vbTab & item.Tipo & vbTab & grdItems.ValueMatrix(k, COL_ITEM_IDIDPADRE)
                            End If
                        Set item = Nothing
                    End With
                Next k
                
                
                
                    VerificaItemPadre gnc, j, grdVenta.TextMatrix(j, COL_VENTA_CANT)
                End If
                If grdVenta.TextMatrix(j, COL_ITEM_IDIDPADRE) = 0 Then
                    ActualizaIdPadre_EnVenta grdVenta.TextMatrix(j, COL_ITEM_ID), j, gnc, i
                End If
                
                
                
            End If
        End With
    Next j
    grdItems.Refresh
    grdItems.col = COL_ITEM_TIPO
    grdItems.Sort = flexSortGenericDescending
    'VerificaItemPadre gnc, i
    
    
    
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    DispErr
salida:
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdVerificar.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    Exit Sub

End Sub


Private Sub VerificaItemPadre(ByRef gnc As GNComprobante, k As Long, cant As Currency)
    Dim i As Long, obj As Object, v As Variant, m As Long, cantidad As Currency
    Dim n As Integer, item As IVinventario
    Dim rs As Recordset
    On Error GoTo ErrTrap
    For i = 1 To grdItems.Rows - 1
        If (grdItems.ValueMatrix(i, COL_ITEM_TIPO) = TIPORECETA Or (grdItems.ValueMatrix(i, COL_ITEM_TIPO) = TIPORECETA - 1)) Then
            If grdItems.ValueMatrix(i, COL_VENTA_IDIDPADRE) = "0" Then
                grdReceta.Rows = 1
                grdReceta1.Rows = 1
                Set obj = gobjMain.EmpresaActual.RecuperaRecetaReproceso(grdItems.TextMatrix(i, COL_ITEM_IDINV), gnc.TransID, True)
            Else
                Set obj = gobjMain.EmpresaActual.RecuperaRecetaReproceso(grdItems.TextMatrix(i, COL_ITEM_IDINV), gnc.TransID, False)
            End If
            'If obj.RecordCount = 0 Then 'AUC cambiado
             '   i = i + 1
            'Else
                v = MiGetRows(obj)
                grdReceta.Redraw = flexRDNone
                If IsEmpty(v) Then
                   If Not CorrigeItemsFaltantes(grdItems.ValueMatrix(i, COL_ITEM_IDINV), grdItems.ValueMatrix(i, COL_ITEM_CANT), grdItems.ValueMatrix(i, COL_VENTA_IDIDPADRE), gnc) Then
                        MsgBox grdItems.TextMatrix(i, 4) & " No es una preparacion, " & Chr(13) & "Revise la Configuraciòn del Item ", vbInformation
                        Exit For
                   End If
                   'Despues de corregir vuelve a cargar
                        Set rs = gobjMain.EmpresaActual.RecuperaRecetaReproceso(grdItems.TextMatrix(i, COL_ITEM_IDINV), gnc.TransID, True)
                        If rs.RecordCount > 0 Then
                            v = MiGetRows(rs)
                            grdReceta.Redraw = flexRDNone
                            grdReceta.LoadArray v
                        End If
                Else
                       grdReceta.LoadArray v
                End If
                grdReceta.col = COL_VENTA_TIPO
                grdReceta.Sort = flexSortGenericDescending
                grdReceta.Refresh
                grdReceta.Redraw = flexRDDirect
                    If grdReceta.Rows > 1 Then
                    
                    
                For n = 1 To grdReceta.Rows - 1
                    With grdReceta
                            Set item = gnc.Empresa.RecuperaIVInventario(grdReceta.ValueMatrix(n, COL_ITEM_IDINV))
                            If grdReceta.ValueMatrix(n, COL_ITEM_TIPO) <> TIPORECETA Then
                                grdIvkn.AddItem i & vbTab & "0" & vbTab & grdReceta.ValueMatrix(n, COL_ITEM_IDINV) & vbTab & item.CodInventario & vbTab & item.Descripcion & vbTab & grdReceta.ValueMatrix(n, COL_ITEM_CANT) * -1 & vbTab & item.Tipo & vbTab & grdReceta.ValueMatrix(n, COL_ITEM_IDIDPADRE)
                            End If
                        Set item = Nothing
                    End With
                Next n
                    
                    'TIENE QUE REGRESAR AQUI
                    For m = 1 To grdReceta.Rows - 1
                        If grdReceta.TextMatrix(m, COL_ITEM_TIPO) = 0 Or grdReceta.TextMatrix(m, COL_ITEM_TIPO) = 3 Then
                            If grdReceta.TextMatrix(m, COL_ITEM_IDIDPADRE) = 0 Then
                                ActualizaIdPadre_EnDetalleItems grdVenta.TextMatrix(k, COL_ITEM_ID), k, gnc, cant
                            Else
                                ActualizaCantidad_EnDetalleItems grdVenta.TextMatrix(k, COL_ITEM_ID), k, gnc, cant
                            End If
                        Else
                            VerificaItemPadre1 gnc, k, cant
                            'TIENE QUE REGRESAR AQUI PARA  COLOCAR EL IDPADRE EN GRDRECETA
                            If grdReceta.TextMatrix(m, COL_ITEM_IDIDPADRE) = 0 Then
                                ActualizaIdPadre_EnReceta grdVenta.TextMatrix(k, COL_ITEM_ID), k, gnc, m
                            End If
                        End If
                    Next m
                    If grdItems.TextMatrix(i, COL_ITEM_IDIDPADRE) = 0 Then
                        ActualizaIdPadre_EnItem grdVenta.TextMatrix(k, COL_ITEM_ID), k, gnc, i
                    End If
                End If
            'End If
        Else
            'si es cambio de preparacion jeaa 05/05/2010
            If grdItems.TextMatrix(i, COL_ITEM_TIPO) = 0 Or grdItems.TextMatrix(i, COL_ITEM_TIPO) = 3 Then
                If grdItems.TextMatrix(i, COL_ITEM_IDIDPADRE) = 0 Then
                    ActualizaIdPadre_EnItem grdVenta.TextMatrix(k, COL_ITEM_ID), k, gnc, i
                 Else 'AUC
                    'verifica si la preparacion fue cambiada
                    cantidad = cant * grdItems.TextMatrix(i, COL_ITEM_CANT) 'saca la cantidad de vendidad
                    If preparacionCambiada(grdItems.TextMatrix(i, COL_ITEM_IDINV), grdItems.TextMatrix(i, COL_ITEM_IDIDPADRE), cantidad) Then
                        ActualizaCantidad_EnDetalleItems2 grdItems.TextMatrix(i, COL_VENTA_IDIDPADRE), i, gnc, cantidad
                    End If
                End If
            End If
        End If
    Next i
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    DispErr
salida:
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdVerificar.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    Exit Sub
    
End Sub
Private Function preparacionCambiada(idinventario As Long, IdPadre As Long, cantidad As Currency) As Boolean
Dim i As Long
For i = 1 To grdIvk.Rows - 1
    If grdIvk.TextMatrix(i, COL_ITEM_IDINV) = idinventario And grdIvk.TextMatrix(i, COL_ITEM_IDIDPADRE) = IdPadre Then
        If grdIvk.ValueMatrix(i, COL_ITEM_CANT) <> cantidad Then
            preparacionCambiada = True
            Exit Function
        End If
    End If
Next
End Function
    Private Sub ActualizaIdPadre_EnDetalleItems(ByVal idPadres As Long, k As Long, gnc As GNComprobante, cant As Currency)
        Dim l As Integer, sql As String, rs As Recordset, fila As Integer
        On Error GoTo ErrTrap
        For l = 1 To grdReceta.Rows - 1
            fila = 1
            While fila < grdIvk.Rows
                If grdReceta.TextMatrix(l, COL_ITEM_TIPO) = 0 Then
                    If grdReceta.TextMatrix(l, COL_ITEM_IDINV) = grdIvk.TextMatrix(fila, COL_ITEM_IDINV) _
                            And grdIvk.TextMatrix(fila, COL_ITEM_IDIDPADRE) = 0 Then
                                sql = " update ivkardex set idpadre=" & idPadres
                                sql = sql & " , cantidad =" & grdReceta.ValueMatrix(l, COL_ITEM_CANT) * -1
                                sql = sql & " , costorealtotal = (costorealtotal / cantidad) * " & grdReceta.ValueMatrix(l, COL_ITEM_CANT) * -1
                                sql = sql & " , costototal = (costototal / cantidad) * " & grdReceta.ValueMatrix(l, COL_ITEM_CANT) * -1
                                sql = sql & " Where id=" & grdIvk.TextMatrix(fila, COL_ITEM_ID)
                                sql = sql & " and (idPadre=0 or idpadre is null) "
                                gobjMain.EmpresaActual.OpenRecordset (sql)
                                grdIvk.TextMatrix(fila, COL_ITEM_IDIDPADRE) = idPadres
                                grdReceta.TextMatrix(l, COL_ITEM_IDIDPADRE) = idPadres
                                grdItems.Refresh
                                grdReceta.Refresh
                                fila = grdIvk.Rows + 1
                    Else
                        fila = fila + 1
                    End If
                Else
                    If grdReceta.TextMatrix(fila, COL_ITEM_IDIDPADRE) = 0 Then
                        VerificaItemPadre1 gnc, k, cant
                        fila = grdIvk.Rows + 1
                    Else
                        fila = grdIvk.Rows + 1
                    End If
                End If
            Wend
        Next l
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    DispErr
salida:
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdVerificar.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    Exit Sub
    End Sub
Private Sub VerificaItemPadre1(ByRef gnc As GNComprobante, k As Long, cant As Currency)
    Dim i As Long, obj As Object, v As Variant, rs As Recordset
    On Error GoTo ErrTrap
    For i = 1 To grdReceta.Rows - 1
        If grdReceta.ValueMatrix(i, COL_ITEM_TIPO) = TIPORECETA Then
            If grdReceta.ValueMatrix(i, COL_VENTA_IDIDPADRE) = "0" Then
                Set obj = gobjMain.EmpresaActual.RecuperaRecetaReproceso(grdReceta.TextMatrix(i, COL_ITEM_IDINV), gnc.TransID, True)
            Else
                Set obj = gobjMain.EmpresaActual.RecuperaRecetaReproceso(grdReceta.TextMatrix(i, COL_ITEM_IDINV), gnc.TransID, False)
            End If
            v = MiGetRows(obj)
            grdReceta1.Redraw = flexRDNone
             grdReceta1.Redraw = flexRDDirect
            'VeficaIgualda para saber que esten todos los items de la receta en la grilla cargada
            If IsEmpty(v) Then
                If Not CorrigeItemsFaltantes(grdReceta.ValueMatrix(i, COL_ITEM_IDINV), grdReceta.ValueMatrix(i, COL_ITEM_CANT), grdReceta.ValueMatrix(i, COL_VENTA_IDIDPADRE), gnc) Then
                   MsgBox MSG_NG & " " & grdVenta.TextMatrix(i, 4) & " " & Chr(13) & "Revise la Configuraciòn del Item ", vbInformation
                   Exit For
                End If
                        'Despues de corregir vuelve a cargar
                        Set rs = gobjMain.EmpresaActual.RecuperaRecetaReproceso(grdReceta.TextMatrix(i, COL_ITEM_IDINV), gnc.TransID, True)
                        If rs.RecordCount > 0 Then
                            v = MiGetRows(obj)
                            grdReceta1.Redraw = flexRDNone
                            grdReceta1.LoadArray v
                        End If
            Else
'                If VerificaIgualdad(grdReceta.TextMatrix(i, COL_ITEM_IDINV), gnc, grdReceta1) Then
                    grdReceta1.LoadArray v
'               End If
            End If
            If grdReceta1.Rows > 1 Then
                If grdReceta1.TextMatrix(i, COL_ITEM_TIPO) = 0 Then
                    If grdReceta1.TextMatrix(i, COL_ITEM_IDIDPADRE) = 0 Then
                        ActualizaIdPadre_EnDetalleItems1 grdVenta.TextMatrix(k, COL_ITEM_ID), k, gnc
                    Else
                        ActualizaCantidad_EnDetalleItems1 grdVenta.TextMatrix(k, COL_ITEM_ID), k, gnc, cant
                    End If
                Else
                    MsgBox "hola otro item"
                End If
            End If
        End If
    Next i
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    DispErr
salida:
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdVerificar.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    Exit Sub

End Sub

    Private Sub ActualizaIdPadre_EnDetalleItems1(ByVal idPadres As Long, k As Long, gnc As GNComprobante)
        Dim l As Integer, sql As String, rs As Recordset, fila As Integer
        On Error GoTo ErrTrap
        For l = 1 To grdReceta1.Rows - 1
            fila = 1
            While fila < grdIvk.Rows
                    If grdReceta1.TextMatrix(l, COL_ITEM_IDINV) = grdIvk.TextMatrix(fila, COL_ITEM_IDINV) _
                            And grdIvk.TextMatrix(fila, COL_ITEM_IDIDPADRE) = 0 Then
                                    sql = " update ivkardex set idpadre=" & idPadres
                                    sql = sql & " , cantidad =" & grdReceta1.ValueMatrix(l, COL_ITEM_CANT) * -1
                                    sql = sql & " , costorealtotal = (costorealtotal / cantidad) * " & grdReceta1.ValueMatrix(l, COL_ITEM_CANT) * -1
                                    sql = sql & " , costototal = (costototal / cantidad) * " & grdReceta1.ValueMatrix(l, COL_ITEM_CANT) * -1
                                    sql = sql & " Where id=" & grdIvk.TextMatrix(fila, COL_ITEM_ID)
                                    sql = sql & " and (idPadre=0 or idpadre is null) "
                                    gobjMain.EmpresaActual.OpenRecordset (sql)
                                    grdIvk.TextMatrix(fila, COL_ITEM_IDIDPADRE) = idPadres
                                    grdReceta1.TextMatrix(l, COL_ITEM_IDIDPADRE) = idPadres
                                    grdIvk.Refresh
                                    grdReceta1.Refresh
                                    fila = grdIvk.Rows + 1
                    Else
                        fila = fila + 1
                    End If
            Wend
        Next l
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    DispErr
salida:
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdVerificar.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    Exit Sub

    End Sub

    Private Sub ActualizaIdPadre_EnReceta(ByVal idPadres As Long, k As Long, gnc As GNComprobante, l As Long)
        Dim sql As String, rs As Recordset, fila As Integer
        On Error GoTo ErrTrap
        fila = 1
        While fila < grdIvk.Rows
            If grdReceta.TextMatrix(l, COL_ITEM_IDINV) = grdIvk.TextMatrix(fila, COL_ITEM_IDINV) Then
                If grdIvk.TextMatrix(fila, COL_ITEM_IDIDPADRE) = 0 Then
                    sql = " update ivkardex set idpadre=" & idPadres
                    sql = sql & " Where id=" & grdIvk.TextMatrix(fila, COL_ITEM_ID)
                    sql = sql & " and (idPadre=0 or idpadre is null) "
                    gobjMain.EmpresaActual.OpenRecordset (sql)
                    grdIvk.TextMatrix(fila, COL_ITEM_IDIDPADRE) = idPadres
                    grdReceta.TextMatrix(l, COL_ITEM_IDIDPADRE) = idPadres
                    grdIvk.Refresh
                    grdReceta.Refresh
                    fila = grdIvk.Rows + 1
                Else
                    fila = fila + 1
                   ' Fila = grdIvk.Rows + 1
                End If
            Else
                fila = fila + 1
            End If
        Wend
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    DispErr
salida:
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdVerificar.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    Exit Sub
    
    End Sub



Private Sub ActualizaIdPadre_EnItem(ByVal idPadres As Long, k As Long, gnc As GNComprobante, l As Long)
        Dim sql As String, rs As Recordset, fila As Integer
        On Error GoTo ErrTrap
        fila = 1
        While fila < grdIvk.Rows
            If grdItems.TextMatrix(l, COL_ITEM_IDINV) = grdIvk.TextMatrix(fila, COL_ITEM_IDINV) Then
                If grdIvk.TextMatrix(fila, COL_ITEM_IDIDPADRE) = 0 Then
                    sql = " update ivkardex set idpadre=" & idPadres
                    sql = sql & " Where id=" & grdIvk.TextMatrix(fila, COL_ITEM_ID)
                    sql = sql & " and (idPadre=0 or idpadre is null) "
                    gobjMain.EmpresaActual.OpenRecordset (sql)
                    grdIvk.TextMatrix(fila, COL_ITEM_IDIDPADRE) = idPadres
                    grdItems.TextMatrix(l, COL_ITEM_IDIDPADRE) = idPadres
                    grdIvk.Refresh
                    grdItems.Refresh
                    fila = grdIvk.Rows + 1
                Else
                    fila = fila + 1
                End If
            Else
                fila = fila + 1
            End If
        Wend
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    DispErr
salida:
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdVerificar.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    Exit Sub
    
    End Sub

Private Sub ActualizaIdPadre_EnVenta(ByVal idPadres As Long, k As Long, gnc As GNComprobante, l As Long)
        Dim sql As String, rs As Recordset, fila As Integer
        On Error GoTo ErrTrap
        fila = 1
        While fila < grdIvk.Rows
            If grdVenta.TextMatrix(k, COL_ITEM_IDINV) = grdIvk.TextMatrix(fila, COL_ITEM_IDINV) Then
                If grdIvk.TextMatrix(fila, COL_ITEM_IDIDPADRE) = 0 Then
                    sql = " update ivkardex set idpadre=" & idPadres
                    sql = sql & " Where id=" & grdIvk.TextMatrix(fila, COL_ITEM_ID)
                    sql = sql & " and (idPadre=0 or idpadre is null) "
                    gobjMain.EmpresaActual.OpenRecordset (sql)
                    grdIvk.TextMatrix(fila, COL_ITEM_IDIDPADRE) = idPadres
                    grdVenta.TextMatrix(k, COL_ITEM_IDIDPADRE) = idPadres
                    grdIvk.Refresh
                    grdVenta.Refresh
                    fila = grdIvk.Rows + 1
                Else
                    fila = fila + 1
                End If
            Else
                fila = fila + 1
            End If
        Wend
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    DispErr
salida:
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdVerificar.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    Exit Sub
        
End Sub


    Private Sub ActualizaCantidad_EnDetalleItems(ByVal idPadres As Long, k As Long, gnc As GNComprobante, cant As Currency)
        Dim l As Integer, sql As String, rs As Recordset, fila As Integer
        On Error GoTo ErrTrap
        For l = 1 To grdReceta.Rows - 1
            fila = 1
            While fila < grdIvk.Rows
                If grdReceta.TextMatrix(l, COL_ITEM_TIPO) = 0 Or grdReceta.TextMatrix(l, COL_ITEM_TIPO) = 3 Then
                
                    If grdReceta.TextMatrix(l, COL_ITEM_IDINV) = grdIvk.TextMatrix(fila, COL_ITEM_IDINV) Then
                                sql = " update ivkardex "
                                sql = sql & " set cantidad =" & grdReceta.ValueMatrix(l, COL_ITEM_CANT) * -1 * cant
                                sql = sql & " , costorealtotal = (costorealtotal / cantidad) * " & grdReceta.ValueMatrix(l, COL_ITEM_CANT) * -1 * cant
                                sql = sql & " , costototal = (costototal / cantidad) * " & grdReceta.ValueMatrix(l, COL_ITEM_CANT) * -1 * cant
                                sql = sql & " Where id=" & grdIvk.TextMatrix(fila, COL_ITEM_ID)
'                                sql = sql & " and idPadre=" & idPadres
                                gobjMain.EmpresaActual.OpenRecordset (sql)
                                grdIvk.TextMatrix(fila, COL_ITEM_IDIDPADRE) = idPadres
                                grdReceta.TextMatrix(l, COL_ITEM_IDIDPADRE) = idPadres
                                grdItems.Refresh
                                grdReceta.Refresh
                                fila = grdIvk.Rows + 1
                    Else
                        fila = fila + 1
                    End If
                 'auc creo que esto no va
                Else
                    If grdReceta.TextMatrix(fila, COL_ITEM_IDIDPADRE) = 0 Then
                      '  VerificaItemPadre1 gnc, k
                        fila = grdIvk.Rows + 1
                    Else
                        fila = grdIvk.Rows + 1
                    End If
                End If
            Wend
        Next l
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    DispErr
salida:
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdVerificar.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    Exit Sub

    End Sub


    Private Sub ActualizaCantidad_EnDetalleItems1(ByVal idPadres As Long, k As Long, gnc As GNComprobante, cant As Currency)
        Dim l As Integer, sql As String, rs As Recordset, fila As Integer
        On Error GoTo ErrTrap
        For l = 1 To grdReceta1.Rows - 1
            fila = 1
            While fila < grdIvk.Rows
                    If grdReceta1.TextMatrix(l, COL_ITEM_IDINV) = grdIvk.TextMatrix(fila, COL_ITEM_IDINV) Then
                                    sql = " update ivkardex "
                                    sql = sql & " set cantidad =" & grdReceta1.ValueMatrix(l, COL_ITEM_CANT) * -1 * cant
                                    sql = sql & " , costorealtotal = (costorealtotal / cantidad) * " & grdReceta1.ValueMatrix(l, COL_ITEM_CANT) * -1 * cant
                                    sql = sql & " , costototal = (costototal / cantidad) * " & grdReceta1.ValueMatrix(l, COL_ITEM_CANT) * -1 * cant
                                    sql = sql & " Where id=" & grdIvk.TextMatrix(fila, COL_ITEM_ID)
'                                    sql = sql & " and idPadre=" & idPadres
                                    gobjMain.EmpresaActual.OpenRecordset (sql)
                                    grdIvk.TextMatrix(fila, COL_ITEM_IDIDPADRE) = idPadres
                                    grdReceta1.TextMatrix(l, COL_ITEM_IDIDPADRE) = idPadres
                                    grdIvk.Refresh
                                    grdReceta1.Refresh
                                    fila = grdIvk.Rows + 1
                    Else
                        fila = fila + 1
                    End If
            Wend
        Next l
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    DispErr
salida:
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdVerificar.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    Exit Sub

    End Sub
    
 'AUC corrige cantidad en segundo nivel
Private Sub ActualizaCantidad_EnDetalleItems2(ByVal idPadres As Long, k As Long, gnc As GNComprobante, cantidad As Currency)
        Dim l As Integer, sql As String, rs As Recordset, fila As Integer
        On Error GoTo ErrTrap
        For l = k To grdItems.Rows - 1
            fila = 1
            While fila < grdIvk.Rows
                    If grdItems.TextMatrix(l, COL_ITEM_IDINV) = grdIvk.TextMatrix(fila, COL_ITEM_IDINV) _
                        And grdIvk.TextMatrix(fila, COL_VENTA_IDIDPADRE) = idPadres Then
                                    sql = " update ivkardex "
                                    sql = sql & " set cantidad =" & cantidad * -1
                                    sql = sql & " , costorealtotal = (costorealtotal / cantidad) * " & cantidad * -1
                                    sql = sql & " , costototal = (costototal / cantidad) * " & cantidad * -1
                                    sql = sql & " Where id=" & grdIvk.TextMatrix(fila, COL_ITEM_ID)
'                                    sql = sql & " and idPadre=" & idPadres
                                    gobjMain.EmpresaActual.OpenRecordset (sql)
                                    grdIvk.TextMatrix(fila, COL_ITEM_IDIDPADRE) = idPadres
                                    grdItems.TextMatrix(l, COL_ITEM_IDIDPADRE) = idPadres
                                    grdIvk.TextMatrix(fila, COL_ITEM_CANT) = cantidad
                                    grdIvk.Refresh
                                    grdItems.Refresh
                                    fila = grdIvk.Rows + 1
                                    Exit For
                    Else
                        fila = fila + 1
                    End If
            Wend
        Next l
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    DispErr
salida:
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdVerificar.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    Exit Sub

    End Sub



'AUC 09/03/07
Private Function CorrigeItemsFaltantes(ByVal idInven As Long, ByVal cant As Currency, ByVal idPadres As Long, gnc As GNComprobante) As Boolean
        Dim l As Integer, sql As String, rs As Recordset, fila As Integer
        Dim j As Long, idsubItem As Long, CantSubItem As Currency
        Dim Item2 As Recordset, Item3 As IVinventario
        Dim sql1 As String
        Dim item As IVinventario, CostoTotal
        Dim c As Currency
        Dim IdBod As Long, BODAUX As String
        
        On Error GoTo ErrTrap
        Set item = gnc.Empresa.RecuperaIVInventario(idInven)
        If item.NumFamiliaDetalle = 0 Then CorrigeItemsFaltantes = False: Exit Function
'        MsgBox "Se procedera a crear Items faltantes para  " & Chr(13) & gnc.CodTrans & " " & gnc.numtrans & " " & item.Descripcion, vbInformation
        
        If item.NumFamiliaDetalle > 0 And item.Tipo = Preparacion Then
                For j = 1 To item.NumFamiliaDetalle 'ciclo para detalles de items
                    idsubItem = item.RecuperaId(item.RecuperaDetalleFamilia(j).CodInventario)
                    CantSubItem = item.RecuperaDetalleFamilia(j).cantidad
                    'RECUPERO ESE ITEM
                    Set Item3 = gnc.Empresa.RecuperaIVInventario(idsubItem)
                    'abre el ivkardex
                    sql = " Select * from ivkardex where transid = " & gnc.TransID
                    Set Item2 = gobjMain.EmpresaActual.OpenRecordset(sql)
                    Item2.MoveLast
                    
                    c = Item3.CostoDouble2(Date, 0, 0, Time)
                    'inserta el item faltante
                       CostoTotal = -(c * CantSubItem * cant)
                        sql = "INSERT INTO IVkardex (TransID,IdInventario,IdBodega,Cantidad,CostoTotal,CostoRealTotal,PrecioTotal,PrecioRealTotal,Descuento,IVA,Orden,Nota,NumeroPrecio,ValorRecargoItem,TiempoEntrega,bandImprimir,ValorICEItem,IdICE,idPadre,bandver ) "
                        If gnc.GNTrans.CodPantalla = "IVPVTS" Then
                            BODAUX = Mid$(gnc.GNTrans.CodBodegaPre, 1, Len(gnc.GNTrans.CodBodegaPre) - 1) & "2"
                            
                            IdBod = gnc.Empresa.RecuperaIdBodega(BODAUX)
                            sql = sql & " Values(" & gnc.TransID & ", " & idsubItem & ", " & IdBod & ", " & -(CantSubItem * cant) & ", " & CostoTotal & ", " & CostoTotal & " "
                        Else
                            sql = sql & " Values(" & gnc.TransID & ", " & idsubItem & ", " & Item2!IdBodega & ", " & -(CantSubItem * cant) & ", " & CostoTotal & ", " & CostoTotal & " "
                        End If
                        sql = sql & ", 0 ,0,0, " & Item3.PorcentajeIVA & ", " & Item2!orden + 1 & ", 0, 0, 0, 0, 0, 0, 0," & idPadres & ", 0)"
                        gobjMain.EmpresaActual.EjecutarSQL sql, 0
                        'colocar aqui proceso recursivo
                Next
        End If
    CorrigeItemsFaltantes = True
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    DispErr
salida:
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdVerificar.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    Exit Function

    End Function

'AUC verfica que lo mostrado sea igual que la receta, ver si se necesita despues
'Private Function VerificaIgualdad(idItem As Long, gnc As GNComprobante, grd As VSFlexGrid) As Boolean
'Dim item As IVinventario
'        Set item = gnc.Empresa.RecuperaIVInventario(idItem)
'        If item.NumFamiliaDetalle > 0 Then
'            idsubItem = item.RecuperaID(item.RecuperaDetalleFamilia(j).CodInventario, grd)
'            If idsubItem = verificalista(idsubItem) Then
'                VerificaIgualdad = True
'            Else
'                VerificaIgualdad = False
'            End If
'        End If
'End Function
''Vefica que el item mostrado corresponda a la receta devuelta
'Private Function verificalista(idsubItem, grd As VSFlexGrid) As Boolean
'Dim i As Long
'    For i = 1 To grd.Rows - 1
'        If idsubItem = grd.TextMatrix(i, 1) Then
'            verificalista = True
'            Exit Sub
'        End If
'    Next
'End Function

Private Sub cargaItemsKardexNuevo(ByVal gr As Object, ByRef gnc As GNComprobante)
    Dim j As Long, ivk As IVKardex, item As IVinventario, i As Integer
    i = 1
'    grdIvk.Rows = 1
    'carga la el detalle transaccion
    For j = 2 To gr.Rows
        With grdIvk
'            Set ivk = gnc.IVKardex(j)
            Set item = gnc.Empresa.RecuperaIVInventario(gr.TextMatrix(2, 1))
                .AddItem i & vbTab & ivk.id & vbTab & ivk.idinventario & vbTab & ivk.CodInventario & vbTab & item.Descripcion & vbTab & ivk.cantidad * -1 & vbTab & item.Tipo & vbTab & ivk.IdPadre
                If item.Tipo = "4" Then
                    grdIvkn.AddItem i & vbTab & ivk.id & vbTab & ivk.idinventario & vbTab & ivk.CodInventario & vbTab & item.Descripcion & vbTab & ivk.cantidad * -1 & vbTab & item.Tipo & vbTab & ivk.IdPadre
                End If
                
            Set item = Nothing
        End With
    Next j

    grdIvkn.col = COL_VENTA_TIPO
    grdIvkn.Sort = flexSortGenericDescending
    grdIvkn.Refresh


End Sub

Private Sub ActualizaCantidad_EnDetalleItemsTransformacion(ByVal idPadres As Long, k As Long, gnc As GNComprobante, cant As Currency)
        Dim l As Integer, sql As String, rs As Recordset, fila As Integer
        On Error GoTo ErrTrap
        For l = 1 To grdIvk.Rows - 1
            fila = 1
            While fila <= grdIvkn.Rows - 1
                If grdIvk.TextMatrix(l, COL_ITEM_TIPO) = 3 Then
                
                    If grdIvk.TextMatrix(l, COL_ITEM_IDINV) = grdIvkn.TextMatrix(fila, COL_ITEM_IDINV) Then
                                sql = " update ivkardex "
                                sql = sql & " set cantidad =" & Abs(grdIvkn.ValueMatrix(fila, COL_ITEM_CANT)) * -1 * cant
                                sql = sql & " , costorealtotal = (costorealtotal / cantidad) * " & grdIvkn.ValueMatrix(fila, COL_ITEM_CANT) * -1 * cant
                                sql = sql & " , costototal = (costototal / cantidad) * " & grdIvkn.ValueMatrix(fila, COL_ITEM_CANT) * -1 * cant
                                sql = sql & " Where id=" & grdIvk.TextMatrix(l, COL_ITEM_ID)
'                                sql = sql & " and idPadre=" & idPadres
                                gobjMain.EmpresaActual.OpenRecordset (sql)
'                                grdIvk.TextMatrix(Fila, COL_ITEM_IDIDPADRE) = idPadres
'                                grdIvkn.TextMatrix(Fila, COL_ITEM_IDIDPADRE) = idPadres
                                grdItems.Refresh
                                grdReceta.Refresh
'                                Fila = grdIvk.Rows + 1
                                fila = fila + 1
                    Else
                        fila = fila + 1
                    End If
                 'auc creo que esto no va
                Else
                    fila = fila + 1
''                    If grdReceta.TextMatrix(Fila, COL_ITEM_IDIDPADRE) = 0 Then
''                      '  VerificaItemPadre1 gnc, k
''                        Fila = grdIvk.Rows + 1
''                    Else
''                        Fila = grdIvk.Rows + 1
''                    End If
                End If
            Wend
        Next l
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    DispErr
salida:
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdVerificar.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    Exit Sub

    End Sub

Private Sub fcbDesde2_Selected(ByVal Text As String, ByVal KeyText As String)
    fcbHasta2.KeyText = fcbDesde2.KeyText   '*** MAKOTO 27/jun/2000
End Sub

Private Sub CargaItemsCombo()
    Dim numGrupo As Integer, v() As Variant
    Dim sql  As String, rs As Recordset, cond As String
    fcbDesde2.Clear
    fcbHasta2.Clear
    sql = "SELECT CodInventario, IVInventario.Descripcion FROM IVInventario ORDER BY Descripcion "
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not rs.EOF Then
        v = MiGetRows(rs)
        fcbDesde2.SetData v
        fcbHasta2.SetData v
    End If
    fcbDesde2.Text = ""
    fcbHasta2.Text = ""
End Sub

Private Function RegenerarConsumo(bandVerificar As Boolean, BandTodo As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, x As Single
    Dim gnc As GNComprobante, cambiado As Boolean
    On Error GoTo ErrTrap

    'Si no es solo verificacion, confirma
    If Not bandVerificar Then
        s = "Este proceso modificará las cantidades y costos de las recetas de la transacción seleccionada." & vbCr & vbCr
        s = s & "Está seguro que desea proceder?"
        If MsgBox(s, vbYesNo + vbQuestion) <> vbYes Then Exit Function
    End If
    
    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdVerificar.Enabled = False
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
        
        'Si es verificación, procesa todas las filas sino solo las que tengan "Asiento incorrecto."
        If (grd.TextMatrix(i, COL_RESULTADO) = MSG_NG) Or bandVerificar Or BandTodo Then
        
            tid = grd.ValueMatrix(i, COL_TID)
            grd.TextMatrix(i, COL_RESULTADO) = "Verificando..."
            grd.Refresh
            
            'Recupera la transaccion
            Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
            If Not (gnc Is Nothing) Then
                'Si la transacción no está anulada
                If gnc.Estado <> ESTADO_ANULADO Then
                    
                    'Forzar recuperar todos los datos de transacción para que no se pierdan al grabar de nuveo
                    gnc.RecuperaDetalleTodo
                    
                    cargaItemsKardex gnc, i
                    num_fila_trans = i
                    cargaItemsConsumo gnc, i
                Else
                    'Si está anulada
                    grd.TextMatrix(i, COL_RESULTADO) = "Anulado."
                End If
            Else
                grd.TextMatrix(i, COL_RESULTADO) = "No pudo recuperar la transación."
            End If
        End If
    Next i
    
    Screen.MousePointer = 0
    RegenerarConsumo = Not mCancelado
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    DispErr
salida:
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdVerificar.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    Exit Function
End Function

Private Sub cargaItemsKardex(ByRef gnc As GNComprobante, ByVal i As Long)
    Dim j As Long, ivk As IVKardex, item As IVinventario, itemT As IVinventario, gncing As GNComprobante
    Dim ivking As IVKardex, idTamanio As Long
    grdIvk.Rows = 1
    grdIvkn.Rows = 1
        
    For j = 1 To gnc.CountIVKardex
        With grdIvk
            Set ivk = gnc.IVKardex(j)
            Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
            idTamanio = gnc.Empresa.BuscaDatosInventarioISO(ivk.TiempoEntrega, "idTamanio")
            Set itemT = gnc.Empresa.RecuperaIVInventario(idTamanio)
            
                .AddItem ivk.orden & vbTab & ivk.id & vbTab & idTamanio & vbTab & itemT.CodInventario & vbTab & ivk.idinventario & vbTab & ivk.CodInventario & vbTab & ivk.cantidad & vbTab & ivk.TiempoEntrega & vbTab & IIf(ivk.bandImprimir, "S", "N") & vbTab & ivk.CodMotivoPro
            
            Set item = Nothing
        End With
    Next j
    grdIvk.col = COL_VENTA_TIPO
    grdIvk.Sort = flexSortGenericDescending
    grdIvk.Refresh

    grdIvkn.col = COL_VENTA_TIPO
    grdIvkn.Sort = flexSortGenericDescending
    grdIvkn.Refresh


End Sub


Private Sub cargaItemsConsumo(ByRef gnc As GNComprobante, ByVal i As Long)
    Dim j As Long, ivk As IVKardex, item As IVinventario, ITEMCONS As IVinventario, Banda As IVinventario
    Dim codRelleno As String, codCemento As String, codCojin As String ', codParche As String

    Dim porRelleno As Currency, porCemento As Currency, porCojin As Currency
    Dim costoRelleno As Currency, costoCemento As Currency, costoCojin As Currency, codbanda As String
    Dim costobanda As Currency
    Dim v As Variant, k As Integer, ix As Long
    Dim CONSUMO As IVConsumoDetalle, sql As String
    
    'carga la el detalle transaccion
    For j = 1 To grdIvk.Rows - 1
        
        With grdIvkn
            grdItems.Clear
            grdItems.Rows = 1
            Set ivk = gnc.IVKardex(j)
            Set item = gnc.Empresa.RecuperaIVInventario(grdIvk.TextMatrix(j, 3))
            If item.CodGrupo(2) = "NOR" Then
                
                If Len(gnc.Empresa.GNOpcion.ObtenerValor("Porcentaje_RELLENO")) > 0 Then
                    v = Split(gnc.Empresa.GNOpcion.ObtenerValor("Porcentaje_RELLENO"), ",")
                    codRelleno = v(0)
                    porRelleno = Round((v(1) / 100), 4)
                    Set ITEMCONS = gnc.Empresa.RecuperaIVInventarioQuick(codRelleno)
                    costoRelleno = 0
                    If Not item Is Nothing Then
                        costoRelleno = ITEMCONS.CostoDouble2(gnc.FechaTrans, _
                                    1, _
                                    gnc.TransID, _
                                    gnc.HoraTrans)  '*** MAKOTO 08/dic/00 Agregado Hora
                    End If
                End If
                
                If Len(gnc.Empresa.GNOpcion.ObtenerValor("Porcentaje_CEMENTO")) > 0 Then
                    v = Split(gnc.Empresa.GNOpcion.ObtenerValor("Porcentaje_CEMENTO"), ",")
                    codCemento = v(0)
                    porCemento = Round((v(1) / 100), 4)
                    Set ITEMCONS = gnc.Empresa.RecuperaIVInventarioQuick(codCemento)
                    costoCemento = 0
                    If Not item Is Nothing Then
                        costoCemento = ITEMCONS.CostoDouble2(gnc.FechaTrans, _
                                    1, _
                                    gnc.TransID, _
                                    gnc.HoraTrans)  '*** MAKOTO 08/dic/00 Agregado Hora
                    End If
                End If
                
                If Len(gnc.Empresa.GNOpcion.ObtenerValor("Porcentaje_COJIN")) > 0 Then
                    v = Split(gnc.Empresa.GNOpcion.ObtenerValor("Porcentaje_COJIN"), ",")
                    codCojin = v(0)
                    porCojin = Round((v(1) / 100), 4)
                    Set ITEMCONS = gnc.Empresa.RecuperaIVInventarioQuick(codCojin)

                    costoCojin = 0
                    If Not item Is Nothing Then
                        costoCojin = ITEMCONS.CostoDouble2(gnc.FechaTrans, _
                                    1, _
                                    gnc.TransID, _
                                    gnc.HoraTrans)  '*** MAKOTO 08/dic/00 Agregado Hora
                    End If
                End If
                
                Set Banda = gnc.Empresa.RecuperaIVInventario(grdIvk.TextMatrix(j, 5))
                costobanda = Banda.CostoDouble2(gnc.FechaTrans, 1, gnc.TransID, gnc.HoraTrans)
                If grdIvk.ValueMatrix(j, 6) > 0 Then
                    If gnc.IVKardex(j).Motivo <> 3 Then
                            .AddItem i & vbTab & "BODMATPRI" & vbTab & Banda.CodInventario & vbTab & grdIvk.ValueMatrix(j, 6) & vbTab & grdIvk.ValueMatrix(j, 6) * costobanda & vbTab & grdIvk.ValueMatrix(j, 6) * costobanda & vbTab & grdIvk.TextMatrix(j, 7)
                            .AddItem i & vbTab & "BODMATPRI" & vbTab & codRelleno & vbTab & grdIvk.ValueMatrix(j, 6) * porRelleno & vbTab & grdIvk.ValueMatrix(j, 6) * porRelleno * costoRelleno & vbTab & grdIvk.ValueMatrix(j, 6) * porRelleno * costoRelleno & vbTab & grdIvk.TextMatrix(j, 7)
                            .AddItem i & vbTab & "BODMATPRI" & vbTab & codCemento & vbTab & grdIvk.ValueMatrix(j, 6) * porCemento & vbTab & grdIvk.ValueMatrix(j, 6) * porCemento * costoCemento & vbTab & grdIvk.ValueMatrix(j, 6) * porCemento * costoCemento & vbTab & grdIvk.TextMatrix(j, 7)
                            .AddItem i & vbTab & "BODMATPRI" & vbTab & codCojin & vbTab & grdIvk.ValueMatrix(j, 6) * porCojin & vbTab & grdIvk.ValueMatrix(j, 6) * porCojin * costoCojin & vbTab & grdIvk.ValueMatrix(j, 6) * porCojin * costoCojin & vbTab & grdIvk.TextMatrix(j, 7)
                            .Refresh
                            
                            
                            
    '                    ivk.cantidad * -1
                    End If
                    
    
                    If gnc.IVKardex(j).Motivo <> 3 Then
                            grdItems.AddItem "1" & vbTab & "BODMATPRI" & vbTab & Banda.CodInventario & vbTab & grdIvk.ValueMatrix(j, 6) & vbTab & grdIvk.ValueMatrix(j, 6) * costobanda & vbTab & grdIvk.ValueMatrix(j, 6) * costobanda & vbTab & grdIvk.TextMatrix(j, 7)
                            grdItems.AddItem "2" & vbTab & "BODMATPRI" & vbTab & codRelleno & vbTab & grdIvk.ValueMatrix(j, 6) * porRelleno & vbTab & grdIvk.ValueMatrix(j, 6) * porRelleno * costoRelleno & vbTab & grdIvk.ValueMatrix(j, 6) * porRelleno * costoRelleno & vbTab & grdIvk.TextMatrix(j, 7)
                            grdItems.AddItem "3" & vbTab & "BODMATPRI" & vbTab & codCemento & vbTab & grdIvk.ValueMatrix(j, 6) * porCemento & vbTab & grdIvk.ValueMatrix(j, 6) * porCemento * costoCemento & vbTab & grdIvk.ValueMatrix(j, 6) * porCemento * costoCemento & vbTab & grdIvk.TextMatrix(j, 7)
                            grdItems.AddItem "4" & vbTab & "BODMATPRI" & vbTab & codCojin & vbTab & grdIvk.ValueMatrix(j, 6) * porCojin & vbTab & grdIvk.ValueMatrix(j, 6) * porCojin * costoCojin & vbTab & grdIvk.ValueMatrix(j, 6) * porCojin * costoCojin & vbTab & grdIvk.TextMatrix(j, 7)
                            grdItems.Refresh
                            
                            
                                sql = "delete IVConsumoDetalle where IdKardexRef = " & gnc.IVKardex(j).id
                                gobjMain.EmpresaActual.OpenRecordset (sql)
                            
                            For k = 1 To 4
                            
                            
                            
                                sql = "Insert  IVConsumoDetalle (TransID,IdKardexRef,Cant,Precio,Costo,Orden,FechaGrabado,ticket,IdInventario)"
                                sql = sql & "  (select " & gnc.TransID & "," & gnc.IVKardex(j).id & "," & grdItems.ValueMatrix(k, 3) * -1 & ",0," & grdItems.ValueMatrix(k, 4) * -1 & "," & k & ",'" & gnc.FechaTrans & "', " & grdItems.TextMatrix(k, 6) & " , idinventario from ivinventario where codinventario='" & grdItems.TextMatrix(k, 2) & "')"
                                gobjMain.EmpresaActual.OpenRecordset (sql)
                            Next k
                            
                            
                            For k = 1 To grdItems.Rows - 1
                                ix = gnc.IVKardex(j).AddConsumoDetalle 'Aumenta  item  a la coleccion
                                Set CONSUMO = gnc.IVKardex(j).RecuperaConsumoDetalle(ix)
                                CONSUMO.cantidad = grdItems.ValueMatrix(k, 3) * -1
                                CONSUMO.CodInventario = grdItems.TextMatrix(k, 2)
                                CONSUMO.orden = grdItems.ValueMatrix(k, 0)
                                CONSUMO.costo = grdItems.ValueMatrix(k, 5) * -1
                            Next k
    '                    ivk.cantidad * -1
                            If gnc.IVKardex(j).bandImprimir Then
                                gnc.IVKardex(j).BandProceso = True
                            End If
                            
    
                            
                            
                    End If
                End If
                
                
                
                Set item = Nothing
                Set Banda = Nothing
                Set ITEMCONS = Nothing
                
            Else

'                        MsgBox "hola"
            
            End If
            Set item = Nothing
        End With
        
    Next j
'    gnc.Grabar False, False
If CHKBAJA.value = vbChecked Then
    GrabarTransAuto gnc
    GrabarTransFerenciaAuto gnc
End If
       grdItems.col = COL_ITEM_TIPO
    grdItems.Sort = flexSortGenericDescending
    grdVenta.Refresh
    CargaItems gnc, i
End Sub


Private Function GrabarTransAuto(gnc As GNComprobante) As Boolean
    Dim Imprime As Boolean, i As Long, ix As Long, j As Integer
    Dim item As IVinventario, rsReceta As Recordset
    Dim Cadena As String, peso As Currency, c As Currency, v As Variant
    Dim codRelleno As String, codCemento As String, codCojin As String ', codParche As String
    Dim porRelleno As Currency, porCemento As Currency, porCojin As Currency
    Dim costoRelleno As Currency, costoCemento As Currency, costoCojin As Currency
    On Error GoTo ErrTrap
    
    Set mobjGNCompAux = gnc.Empresa.CreaGNComprobante("BD")
    
    If Not mobjGNCompAux Is Nothing Then
    
        If mobjGNCompAux.SoloVer Then
            MsgBox MSG_NODISPONE, vbInformation
            Exit Function
        End If
        
        'carga los hijos de los items seleccionados
        For i = 1 To grdIvkn.Rows - 1
                    ' banda
                    ix = mobjGNCompAux.AddIVKardex
                    mobjGNCompAux.IVKardex(ix).CodBodega = "BODMATPRI"
                    mobjGNCompAux.IVKardex(ix).CodInventario = grdIvkn.TextMatrix(i, 2)
                    mobjGNCompAux.IVKardex(ix).cantidad = grdIvkn.ValueMatrix(i, 3) * -1
                    mobjGNCompAux.IVKardex(ix).CostoTotal = grdIvkn.ValueMatrix(i, 4) * -1
        Next i
        
        If mobjGNCompAux.CountIVKardex = 0 Then
            Exit Function
        End If
        
        mobjGNCompAux.TotalizaItemRepetido
        
        mobjGNCompAux.FechaTrans = gnc.FechaTrans
        mobjGNCompAux.HoraTrans = gnc.HoraTrans
        Cadena = "Por produccion de " & gnc.CodTrans & "-" & gnc.numtrans & " Maquina: " & gnc.CodCentro
        If Len(Cadena) > 120 Then
            mobjGNCompAux.Descripcion = Mid$(Cadena, 1, 120)
        Else
            mobjGNCompAux.Descripcion = Cadena
        End If
            
        mobjGNCompAux.codUsuario = gnc.codUsuario
        mobjGNCompAux.IdResponsable = gnc.IdResponsable
        mobjGNCompAux.numDocRef = gnc.CodTrans & " " & gnc.numtrans
        mobjGNCompAux.idCentro = gnc.idCentro
        mobjGNCompAux.IdTransFuente = gnc.Empresa.RecuperarTransIDGncomprobante(gnc.CodTrans, gnc.numtrans)
        mobjGNCompAux.CodMoneda = gnc.CodMoneda
'        mobjGNCompAux.CodVendedor = fcbVendedor.KeyText
    
        'Si es que algo está modificado
        If mobjGNCompAux.Modificado Then
            MensajeStatus MSG_GENERANDOASIENTO, vbHourglass
            MensajeStatus
        End If
        'Verificación de datos
        mobjGNCompAux.VerificaDatos
    
        PreparaAsientoAuto True
        'Verifica si está cuadrado el asiento
        If Not VerificaAsiento(mobjGNCompAux) Then Exit Function
    
        'Verifica si tiene detalle de banco
        If (mobjGNCompAux.CountIVKardex = 0) Then
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

    Exit Function

End Function


Private Sub PreparaAsientoAuto(Aceptar As Boolean)
    If mobjGNCompAux.SoloVer Then Exit Sub
    
    
'    'Genera el registro correspondiente al pago/cobro al contado
    mobjGNCompAux.GeneraAsiento
End Sub


Private Function GrabarTransFerenciaAuto(gnc As GNComprobante) As Boolean
    Dim Imprime As Boolean, i As Long, ix As Long, j As Integer
    Dim ItemReceta As IVinventario, rsReceta As Recordset
    Dim Cadena As String, BandCarcasaISO As Boolean
    Dim BandFactCarcasa As Boolean
    On Error GoTo ErrTrap
    Set mobjGNCompTransf = gnc.Empresa.CreaGNComprobante("TBPPT")
    
    If Not mobjGNCompTransf Is Nothing Then
    
        If mobjGNCompTransf.SoloVer Then
            MsgBox MSG_NODISPONE, vbInformation
            Exit Function
        End If
        'carga los hijos de los items seleccionados
        For i = 1 To grdIvk.Rows - 1
            
            ix = mobjGNCompTransf.AddIVKardex
            'If mobjGNComp.IVKardex(i).bandImprimir Then
            If gnc.IVKardex(i).BandProceso Then
            
                BandCarcasaISO = gnc.Empresa.RecuperaBandISO(gnc.IVKardex(i).TiempoEntrega, BandFactCarcasa)
                If BandCarcasaISO Then
                    mobjGNCompTransf.IVKardex(ix).CodBodega = "BODPTISO"
                Else
                    mobjGNCompTransf.IVKardex(ix).CodBodega = "BODPT"
                End If
            Else
                mobjGNCompTransf.IVKardex(ix).CodBodega = "BODNOCONF"
            End If
            mobjGNCompTransf.IVKardex(ix).CodInventario = grdIvk.TextMatrix(i, 3)
            mobjGNCompTransf.IVKardex(ix).TiempoEntrega = grdIvk.TextMatrix(i, 7)
            mobjGNCompTransf.IVKardex(ix).bandVer = True
            mobjGNCompTransf.IVKardex(ix).cantidad = 1
            
            ix = mobjGNCompTransf.AddIVKardex
            
            mobjGNCompTransf.IVKardex(ix).CodBodega = gnc.IVKardex(i).CodBodega
            mobjGNCompTransf.IVKardex(ix).CodInventario = grdIvk.TextMatrix(i, 3)
            mobjGNCompTransf.IVKardex(ix).TiempoEntrega = grdIvk.TextMatrix(i, 7)
            mobjGNCompTransf.IVKardex(ix).bandVer = True
            mobjGNCompTransf.IVKardex(ix).cantidad = -1
            
            
        Next i

        mobjGNCompTransf.FechaTrans = gnc.FechaTrans
        mobjGNCompTransf.HoraTrans = gnc.HoraTrans
        Cadena = "Por produccion de " & gnc.CodTrans & "-" & gnc.numtrans & " Maquina: " & gnc.CodCentro
        If Len(Cadena) > 120 Then
            mobjGNCompTransf.Descripcion = Mid$(Cadena, 1, 120)
        Else
            mobjGNCompTransf.Descripcion = Cadena
        End If
            
        mobjGNCompTransf.codUsuario = gnc.codUsuario
        mobjGNCompTransf.IdResponsable = gnc.IdResponsable
        mobjGNCompTransf.numDocRef = gnc.CodTrans & " " & gnc.numtrans
        mobjGNCompTransf.idCentro = gnc.idCentro
        mobjGNCompTransf.IdTransFuente = gnc.Empresa.RecuperarTransIDGncomprobante(gnc.CodTrans, gnc.numtrans)
        mobjGNCompTransf.CodMoneda = gnc.CodMoneda
'        mobjGNCompTransf.CodVendedor = fcbVendedor.KeyText
    
        'Si es que algo está modificado
        If mobjGNCompTransf.Modificado Then
            MensajeStatus MSG_GENERANDOASIENTO, vbHourglass
            MensajeStatus
        End If
        If mobjGNCompTransf.GNTrans.AfectaSaldoPC And _
           mobjGNCompTransf.GNTrans.TSVerificaTotalCuadrado Then
            'Verifica si está cuadrado el total de transacción y total de PCKardex.
'            If Not TotalCuadrado Then Exit Function
        End If
        'Verificación de datos
        mobjGNCompTransf.VerificaDatos
    
        PreparaAsientoAuto True
        'Verifica si está cuadrado el asiento
        If Not VerificaAsiento(mobjGNCompTransf) Then Exit Function
    
        'Verifica si tiene detalle de banco
        If (mobjGNCompTransf.CountIVKardex = 0) Then
            MsgBox "No existe ningún detalle.", vbInformation
            'sst1.Tab = 0
            Exit Function
        End If

        MensajeStatus MSG_GRABANDO, vbHourglass
    
        'Manda a grabar
        '       Aquí ya no hacemos verificación de asiento por que ya está hecho en Control Asiento
        mobjGNCompTransf.Grabar False, False

        '***  Oliver 26/12/2002
        'Agregado para el control ded Impresion Configurado en la Transaccion
        

        MensajeStatus
        GrabarTransFerenciaAuto = True
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
    Exit Function

End Function



