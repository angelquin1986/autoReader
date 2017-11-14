VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{ED5A9B02-5BDB-48C7-BAB1-642DCC8C9E4D}#2.0#0"; "SelFold.ocx"
Begin VB.Form frmAnexoTransaccional2016 
   Caption         =   "Anexo Transaccional"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9825
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   9825
   WindowState     =   2  'Maximized
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   3870
      Left            =   -300
      TabIndex        =   9
      Top             =   5460
      Width           =   7395
      _cx             =   13044
      _cy             =   6826
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
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAnexoTransaccional2016.frx":0000
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
   Begin VB.Frame frmfecha 
      Height          =   1755
      Left            =   60
      TabIndex        =   17
      Top             =   60
      Width           =   5595
      Begin VB.ComboBox cboTipo 
         Height          =   315
         ItemData        =   "frmAnexoTransaccional2016.frx":0063
         Left            =   840
         List            =   "frmAnexoTransaccional2016.frx":006D
         TabIndex        =   1
         Top             =   240
         Width           =   2235
      End
      Begin VB.TextBox txtCarpeta 
         Height          =   320
         Left            =   840
         TabIndex        =   3
         Text            =   "c:\"
         Top             =   840
         Width           =   4170
      End
      Begin VB.CommandButton cmdExaminarCarpeta 
         Caption         =   "..."
         Height          =   320
         Index           =   0
         Left            =   4980
         TabIndex        =   4
         Top             =   840
         Width           =   372
      End
      Begin SelFold.SelFolder slf 
         Left            =   4200
         Top             =   660
         _ExtentX        =   1349
         _ExtentY        =   265
         Title           =   "Seleccione una carpeta"
         Caption         =   "Selección de carpeta"
         RootFolder      =   "\"
         Path            =   "C:\VBPROG_ESP\SII\SELFOLD"
      End
      Begin MSComCtl2.DTPicker dtpPeriodo 
         Height          =   315
         Left            =   840
         TabIndex        =   2
         Top             =   540
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
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
         CustomFormat    =   "MMMM/yyyy"
         Format          =   113180675
         CurrentDate     =   37356
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   60
         TabIndex        =   20
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label2 
         Caption         =   "Mes:"
         Height          =   255
         Left            =   60
         TabIndex        =   19
         Top             =   600
         Width           =   570
      End
      Begin VB.Label Label1 
         Caption         =   "Ubicacion:"
         Height          =   255
         Left            =   60
         TabIndex        =   18
         Top             =   900
         Width           =   870
      End
   End
   Begin VB.Frame fraPasos 
      Height          =   2475
      Left            =   5760
      TabIndex        =   12
      Top             =   60
      Width           =   14355
      Begin VB.CommandButton cmdExcel 
         Caption         =   "Exportar"
         Height          =   540
         Left            =   5340
         Picture         =   "frmAnexoTransaccional2016.frx":008C
         Style           =   1  'Graphical
         TabIndex        =   47
         TabStop         =   0   'False
         ToolTipText     =   "Exporta a Excel"
         Top             =   1740
         Width           =   900
      End
      Begin VB.CheckBox chkConsFinal 
         Caption         =   "Cargar Consumidor Final"
         Height          =   195
         Left            =   5340
         TabIndex        =   40
         Top             =   660
         Visible         =   0   'False
         Width           =   2050
      End
      Begin VB.CheckBox chkSoloError 
         Caption         =   "Solo con error"
         Height          =   195
         Left            =   5340
         TabIndex        =   35
         Top             =   300
         Width           =   1455
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "Buscar"
         Height          =   330
         Index           =   8
         Left            =   2940
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1680
         Width           =   675
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "Generar Archivo"
         Height          =   330
         Index           =   10
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "Generar"
         Height          =   330
         Index           =   9
         Left            =   3660
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "Generar"
         Height          =   330
         Index           =   5
         Left            =   3660
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "Generar"
         Height          =   330
         Index           =   7
         Left            =   3660
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1320
         Width           =   735
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "Generar"
         Height          =   330
         Index           =   3
         Left            =   3660
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "Generar"
         Height          =   330
         Index           =   1
         Left            =   3660
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "Buscar"
         Height          =   330
         Index           =   4
         Left            =   2940
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   960
         Width           =   675
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "Buscar"
         Height          =   330
         Index           =   0
         Left            =   2940
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   675
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "Buscar"
         Height          =   330
         Index           =   2
         Left            =   2940
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         Width           =   675
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "Buscar"
         Height          =   330
         Index           =   6
         Left            =   2940
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1320
         Width           =   675
      End
      Begin MSComDlg.CommonDialog dlg1 
         Left            =   3120
         Top             =   1860
         _ExtentX        =   688
         _ExtentY        =   688
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VSFlex7LCtl.VSFlexGrid grdTalonCP 
         Height          =   1995
         Left            =   6960
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   6735
         _cx             =   11880
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAnexoTransaccional2016.frx":044E
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
      Begin VSFlex7LCtl.VSFlexGrid grdTalonRTPCP 
         Height          =   915
         Left            =   8520
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   3015
         _cx             =   5318
         _cy             =   1614
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAnexoTransaccional2016.frx":058F
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
      Begin VSFlex7LCtl.VSFlexGrid grdTalonRTPIVACP 
         Height          =   855
         Left            =   8940
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   3255
         _cx             =   5741
         _cy             =   1508
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAnexoTransaccional2016.frx":06D0
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
      Begin VSFlex7LCtl.VSFlexGrid grdTalonFC 
         Height          =   1995
         Left            =   6960
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   6675
         _cx             =   11774
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmAnexoTransaccional2016.frx":0811
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
      Begin VB.Label lblnumAnulados 
         Caption         =   "0"
         Height          =   195
         Left            =   12600
         TabIndex        =   46
         Top             =   780
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblResp 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   2
         Left            =   4440
         TabIndex        =   39
         Top             =   960
         Width           =   825
      End
      Begin VB.Label lblPasos 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3. Pasar Ventas x Establecimiento"
         Height          =   330
         Index           =   2
         Left            =   120
         TabIndex        =   38
         Top             =   960
         Width           =   2805
      End
      Begin VB.Label lblPasos 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5. Generar Archivo ATS"
         Height          =   330
         Index           =   5
         Left            =   120
         TabIndex        =   34
         Top             =   2040
         Width           =   2805
      End
      Begin VB.Label lblResp 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   5
         Left            =   4440
         TabIndex        =   33
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label lblResp 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   3
         Left            =   4440
         TabIndex        =   25
         Top             =   1320
         Width           =   825
      End
      Begin VB.Label lblResp 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   4
         Left            =   4440
         TabIndex        =   24
         Top             =   1680
         Width           =   825
      End
      Begin VB.Label lblResp 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   1
         Left            =   4440
         TabIndex        =   23
         Top             =   600
         Width           =   825
      End
      Begin VB.Label lblResp 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   0
         Left            =   4440
         TabIndex        =   22
         Top             =   240
         Width           =   825
      End
      Begin VB.Label lblPasos 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3. Pasar Exportaciones"
         Height          =   330
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   2805
      End
      Begin VB.Label lblPasos 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1. Pasar Compras"
         Height          =   330
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   2800
      End
      Begin VB.Label lblPasos 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2. Pasar Ventas"
         Height          =   330
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   2800
      End
      Begin VB.Label lblPasos 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4. Pasar Comprobantes Anulados"
         Height          =   330
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   2805
      End
   End
   Begin VB.PictureBox picBoton 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   9825
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7365
      Width           =   9825
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   288
         Left            =   10020
         TabIndex        =   10
         Top             =   60
         Width           =   1212
      End
      Begin MSComctlLib.ProgressBar prg 
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   120
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdCF 
      Height          =   4050
      Left            =   15120
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2880
      Visible         =   0   'False
      Width           =   3435
      _cx             =   6059
      _cy             =   7144
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
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAnexoTransaccional2016.frx":0952
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
   Begin VSFlex7LCtl.VSFlexGrid grdRet 
      Height          =   3870
      Left            =   -180
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2580
      Width           =   6615
      _cx             =   11668
      _cy             =   6826
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
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAnexoTransaccional2016.frx":09B5
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
   Begin VSFlex7LCtl.VSFlexGrid GrdRetVentas 
      Height          =   3870
      Left            =   7440
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   2700
      Width           =   6555
      _cx             =   11562
      _cy             =   6826
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
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAnexoTransaccional2016.frx":0A18
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
   Begin VB.Label lblPasos 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5. Generar Archivo ATS"
      Height          =   330
      Index           =   6
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   2805
   End
End
Attribute VB_Name = "frmAnexoTransaccional2016"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbooProcesando As Boolean
Private mbooCancelado As Boolean
Private mEmpOrigen As Empresa
Private Const MSG_OK As String = "OK"
Private mObjCond As RepCondicion
Private mobjBusq As Busqueda

Private WithEvents mGrupo As grupo
Attribute mGrupo.VB_VarHelpID = -1
Const COL_C_TRANSID = 1
Const COL_C_FECHATRANS = 2
Const COL_C_FECHAREGISTRO = 3
Const COL_C_TRANS = 4
Const COL_C_NUMTRANS = 5
Const COL_C_TIPODOC = 6

Const COL_C_IDPROV = 7
Const COL_C_RUC = 8
Const COL_C_NOMBRE = 9
Const COL_C_REL = 10
Const COL_C_NUMSERESTA = 11
Const COL_C_NUMSERPUNTO = 12
Const COL_C_NUMSECUENCIAL = 13
Const COL_C_NUMAUTOSRI = 14
Const COL_C_FECHACADUCIDAD = 15
Const COL_C_CODSUSTENTO = 16
Const COL_C_CODTIPOCOMP = 17
Const COL_C_TIPOPAGO = 18
Const COL_C_PAGOEXTERIOR = 19
Const COL_C_CODPAIS = 20
Const COL_C_DOBLETRIB = 21
Const COL_C_PAGOSUJRET = 22
Const COL_C_BASE0 = 23
Const COL_C_BASE12 = 24
Const COL_C_BASENO12 = 25
Const COL_C_CODICE = 26
Const COL_C_MONTOICE = 27
'Const COL_C_MONTOIVA = 27
Const COL_C_ENOTRARET = 28
Const COL_C_IVA = 29
Const COL_C_OTROVALOR = 30
Const COL_C_RESP = 31

Const COL_R_TIPO = 1
Const COL_R_CODIGORET = 2
Const COL_R_CODIGOSRI = 3
Const COL_R_PORCEN = 4
Const COL_R_TRANS = 5
Const COL_R_NUMTRANS = 6
Const COL_R_RUC = 7
Const COL_R_RETTRANS = 8
Const COL_R_RETNUMTRANS = 9
Const COL_R_FECHARET = 10
Const COL_R_NUMEST = 11
Const COL_R_NUMPTO = 12
Const COL_R_NUMRET = 13
Const COL_R_NUMAUTO = 14
Const COL_R_BASE = 15
Const COL_R_VALOR = 16
Const COL_R_OTROVALR = 17

Const COL_V_FECHA = 1
Const COL_V_TIPODOC = 2
Const COL_V_IDPROVCLI = 3
Const COL_V_RUC = 4
Const COL_V_CLIENTE = 5
Const COL_V_TIPOCOMP = 6
Const COL_V_CANTRANS = 7
Const COL_V_BASE0 = 8
Const COL_V_BASEIVA = 9
Const COL_V_BASENOIVA = 10
Const COL_V_VALORIVA = 11
Const COL_V_IVA = 12
Const COL_V_ICE = 13
Const COL_V_TIPOEMI = 14
Const COL_V_FORMAPAGO = 15

Const COL_V_BASERETIVA = 25
Const COL_V_RETIVA = 26
Const COL_V_BASERETRENTA = 27
Const COL_V_RETRENTA = 28

Const COL_V_RESP = 29

Const COL_VE_SUC = 1
Const COL_VE_TIPOCOMP = 2
Const COL_VE_CANTRANS = 3
Const COL_VE_BASE0 = 4
Const COL_VE_BASEIVA = 5
Const COL_VE_BASENOIVA = 6
Const COL_VE_ICE = 7
Const COL_VE_TOTAL = 8
Const COL_VE_RESP = 9


Const COL_RF_TIPO = 1
Const COL_RF_RUC = 2
Const COL_RF_BASE = 3
Const COL_RF_VALOR = 4


Const COL_A_FECHA = 1
Const COL_A_TCODTRAN = 2
Const COL_A_TIPODOC = 3
Const COL_A_NUMESTA = 4
Const COL_A_NUMPUNTO = 5
Const COL_A_NUMSECUE = 6
Const COL_A_NUMAUTO = 7
Const COL_A_FECHAANULA = 8
Const COL_A_RESP = 9


Const COL_E_RUC = 1
Const COL_E_NOMBRE = 2
Const COL_E_REFERENDO = 3
Const COL_E_TIPOCOMPROBANTE = 4
Const COL_E_DISTRITO = 5
Const COL_E_ANIO = 6
Const COL_E_REGIMEN = 7
Const COL_E_CORRELATIVO = 8
Const COL_E_DOCTRANSPORTE = 9
Const COL_E_FECHAEMBARQUE = 10
Const COL_E_VALORFOB = 11
Const COL_E_VALORFOBLOCAL = 12
Const COL_E_NUMSERESTA = 13
Const COL_E_NUMSERPUNTO = 14
Const COL_E_NUMSECUENCIAL = 15
Const COL_E_NUMAUTOSRI = 16
Const COL_E_FECHATRANS = 17
Const COL_E_PRECIOREALTOTAL = 18
Const COL_E_DESTINO = 19
Const COL_E_RESP = 20

'              FECHATRANS              PRECIOREALTOTAL

Private Cadena As String
Private cadEncabezado As String
Private cadCompras As String
Private cadVentas As String
Private cadVentaEsta As String
Private cadAnulados As String
Private cadExportacion As String

Private NumFile As Integer
Private NumProc As Integer
Private TotalVentas As Currency

Private FilaTalon As Integer
Private FilaIni As Integer

Public Sub Inicio(ByVal tag As String)
    On Error GoTo ErrTrap
    Set mObjCond = New RepCondicion
    Select Case tag
        Case "FAT"
            Me.Caption = "Anexo Transaccional"
    End Select

    FilaTalon = 1
    TotalVentas = 0
    dtpPeriodo.value = CDate("01/" & IIf(Month(Date) - 1 <> 0, Month(Date) - 1, 12) & "/" & Year(Date))
    mObjCond.fecha1 = dtpPeriodo.value
    cboTipo.ListIndex = 0
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("RutaATS-REOC")) > 0 Then
        txtCarpeta.Text = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("RutaATS-REOC")
    End If
    Me.tag = tag
    Me.Show
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub cboTipo_Click()
    If cboTipo.ListIndex = 1 Then
        lblPasos(1).Caption = ""
        lblPasos(2).Caption = ""
        lblPasos(3).Caption = ""
        lblPasos(4).Caption = ""
        lblPasos(5).Caption = "2. Generar Archivo REOC"
        cmdPasos(2).Enabled = False
        cmdPasos(3).Enabled = False
        cmdPasos(4).Enabled = False
        cmdPasos(5).Enabled = False
        cmdPasos(6).Enabled = False
        cmdPasos(7).Enabled = False
        cmdPasos(8).Enabled = False
        cmdPasos(9).Enabled = False
        cmdPasos(10).Enabled = False

    Else
        lblPasos(1).Caption = "2. Pasar Ventas"
        lblPasos(2).Caption = "3. Pasar Ventas x Establecimiento"
        lblPasos(3).Caption = "4. Pasar Exportaciones"
        lblPasos(4).Caption = "5. Pasar Comprobantes Anulados"
        lblPasos(5).Caption = "6. Generar Archivo ATS "
        cmdPasos(2).Enabled = True
        cmdPasos(3).Enabled = True
        cmdPasos(4).Enabled = True
        cmdPasos(5).Enabled = True
        cmdPasos(6).Enabled = True
        cmdPasos(7).Enabled = True
        cmdPasos(8).Enabled = True
        cmdPasos(9).Enabled = True
        cmdPasos(10).Enabled = True

    
    End If
End Sub



Private Sub chkSoloError_Click()
Dim i As Integer
    For i = 1 To grd.Rows - 1
        If grd.TextMatrix(i, grd.ColIndex("Resultado")) = " OK " Then
            If chkSoloError.value = vbChecked Then
                grd.RowHidden(i) = True
            Else
                grd.RowHidden(i) = False
            End If
        End If
    Next i
End Sub

Private Sub cmdCancelar_Click()
    mbooCancelado = True
End Sub


Private Sub cmdPasos_Click(Index As Integer)
    Dim r As Boolean, cad As String, nombre As String, file As String
    Dim resp As Integer
    NumProc = Index + 1
    
    If Index Mod 2 = 0 Then
        cmdPasos(0).BackColor = vbButtonFace
        cmdPasos(2).BackColor = vbButtonFace
        cmdPasos(4).BackColor = vbButtonFace
        cmdPasos(6).BackColor = vbButtonFace
        cmdPasos(8).BackColor = vbButtonFace
    End If
    
    
    Select Case Index + 1
    Case 1      '1. Busca Compras
        BuscarComprasATS
        cadCompras = ""
        cmdPasos(0).BackColor = &HFFFF00
    Case 2      '1. Genera Compras
        lblResp(0).Caption = ""
        cadCompras = ""
        If cboTipo.ListIndex = 0 Then
            r = GenerarComprasATS(cadCompras)
        Else
            r = GenerarComprasREOC(cadCompras)
        End If
    Case 3      '2. Busca Ventas
            BuscarVentasATS
            cadVentas = ""
            cmdPasos(2).BackColor = &HFFFF00
    Case 4      '2. Generar ventas
        lblResp(1).Caption = ""
        cadVentas = ""
        If cboTipo.ListIndex = 0 Then
            r = GenerarVentasATS(cadVentas)
        End If
    Case 5      '2. Busca Ventas x Establecimiento
            BuscarVentasEstablecimientoATS
            cadVentaEsta = ""
            cmdPasos(4).BackColor = &HFFFF00
    Case 6      '2. Generar ventas Establecimiento
        lblResp(2).Caption = ""
        cadVentaEsta = ""
        If cboTipo.ListIndex = 0 Then
            r = GenerarVentasEstablecimientoATS(cadVentaEsta)
        End If
    
    Case 7      '3. Busca Exportaciones
            BuscarExportacionesATS
            cadExportacion = ""
            cmdPasos(6).BackColor = &HFFFF00
    Case 8      '3. Generar Exportaciones
        lblResp(3).Caption = ""
        If cboTipo.ListIndex = 0 Then
            r = GenerarExportacionATS(cadExportacion)
        End If
    
    Case 9      '3. Busca Anulados
            BuscarANuladosATS
            cadAnulados = ""
            cmdPasos(8).BackColor = &HFFFF00
    Case 10      '7. Generar Anulados
        lblResp(4).Caption = ""
        cadAnulados = ""
        r = GenerarANuladosATS(cadAnulados)
    
    Case 11      '8. Generar Archivo
        If cboTipo.ListIndex = 0 Then
            nombre = "AT" & Format(CStr(Month(dtpPeriodo.value)), "00") & Year(dtpPeriodo.value) & ".XML"
            file = txtCarpeta.Text & nombre
            If ExisteArchivo(file) Then
                If MsgBox("El nombre del archivo " & nombre & " ya existe desea sobreescribirlo?", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
            NumFile = FreeFile
            Open file For Output Access Write As #NumFile
            cadEncabezado = GeneraArchivoEncabezadoATSXML
            Cadena = cadEncabezado & cadCompras & cadVentas & cadVentaEsta & cadExportacion & cadAnulados & "</iva>"
            Print #NumFile, Cadena
            Close NumFile
            
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "RutaATS-REOC", txtCarpeta.Text
            gobjMain.EmpresaActual.GNOpcion.GrabarSoloGnOpcion2
            
            resp = MsgBox("Desea imprimir Talón Resumen", vbYesNo)
            If resp = vbYes Then
                FrmImprimeEtiketas.Inicioats2016 grdTalonCP, grdTalonFC, grdTalonRTPCP, grdTalonRTPIVACP, Format(CStr(Month(dtpPeriodo.value)), "00") & "-" & Year(dtpPeriodo.value), lblnumAnulados.Caption
            End If
            
            r = True
            
            
        Else
            nombre = "REOC" & Format(CStr(Month(dtpPeriodo.value)), "00") & Year(dtpPeriodo.value) & ".XML"
            file = txtCarpeta.Text & nombre
            If ExisteArchivo(file) Then
                If MsgBox("El nombre del archivo " & nombre & " ya existe desea sobreescribirlo?", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
            NumFile = FreeFile
            Open file For Output Access Write As #NumFile
        
            cadEncabezado = GeneraArchivoEncabezadoREOCXML
               
            Cadena = cadEncabezado & cadCompras & "</reoc>"
            Print #NumFile, Cadena
            Close NumFile
            r = True
        End If
        lblResp(5).Caption = "OK."
    
    End Select
    
    
    
    If r Then
        If Index < cmdPasos.count - 1 Then
            If cboTipo.ListIndex = 0 Then
            If Index <> 10 Then
            Select Case Index
                Case 1
                    If lblResp(0).Caption <> "Error" Then
                        lblResp(0).BackColor = vbBlue
                        lblResp(0).ForeColor = vbYellow
                    End If
                Case 3
                    If lblResp(1).Caption <> "Error" Then
                        lblResp(1).BackColor = vbBlue
                        lblResp(1).ForeColor = vbYellow
                    End If
                Case 5
                    If lblResp(2).Caption <> "Error" Then
                        lblResp(2).BackColor = vbBlue
                        lblResp(2).ForeColor = vbYellow
                    End If
                Case 7
                    If lblResp(3).Caption <> "Error" Then
                        lblResp(3).BackColor = vbBlue
                        lblResp(3).ForeColor = vbYellow
                    End If
                Case 9
                    If lblResp(4).Caption <> "Error" Then
                        lblResp(4).BackColor = vbBlue
                        lblResp(4).ForeColor = vbYellow
                    End If
                
            End Select
            
            
            Else
                lblPasos(Index + 6).BackColor = vbBlue
                lblPasos(Index + 6).ForeColor = vbYellow
            
            End If
            Else
                If cmdPasos(10).Enabled Then
                    cmdPasos(10).SetFocus
                Else
'                    cmdPasos(9).SetFocus
                End If
            End If
        End If
        If Index <> 10 Then
            Select Case Index
                Case 1
                    If lblResp(0).Caption <> "Error" Then
                        lblResp(0).BackColor = vbBlue
                        lblResp(0).ForeColor = vbYellow
                    End If
                Case 3
                    If lblResp(1).Caption <> "Error" Then
                        lblResp(1).BackColor = vbBlue
                        lblResp(1).ForeColor = vbYellow
                    End If
                Case 5
                    If lblResp(2).Caption <> "Error" Then
                        lblResp(2).BackColor = vbBlue
                        lblResp(2).ForeColor = vbYellow
                    End If
                Case 7
                    If lblResp(3).Caption <> "Error" Then
                        lblResp(3).BackColor = vbBlue
                        lblResp(3).ForeColor = vbYellow
                    End If
                Case 9
                    If lblResp(4).Caption <> "Error" Then
                        lblResp(4).BackColor = vbBlue
                        lblResp(4).ForeColor = vbYellow
                    End If
                
            End Select
        Else
                lblResp(5).BackColor = vbBlue
                lblResp(5).ForeColor = vbYellow
        End If
        If Index <> 10 Then
        
            Select Case Index
                Case 1
                    If lblResp(0).Caption <> "Error" Then
                        lblResp(0).BackColor = vbBlue
                        lblResp(0).ForeColor = vbYellow
                    End If
                Case 3
                    If lblResp(1).Caption <> "Error" Then
                        lblResp(1).BackColor = vbBlue
                        lblResp(1).ForeColor = vbYellow
                    End If
                Case 5
                    If lblResp(2).Caption <> "Error" Then
                        lblResp(2).BackColor = vbBlue
                        lblResp(2).ForeColor = vbYellow
                    End If
                Case 7
                    If lblResp(3).Caption <> "Error" Then
                        lblResp(3).BackColor = vbBlue
                        lblResp(3).ForeColor = vbYellow
                    End If
                Case 9
                    If lblResp(4).Caption <> "Error" Then
                        lblResp(4).BackColor = vbBlue
                        lblResp(4).ForeColor = vbYellow
                    End If
                
            End Select
        
        
        End If
    End If

End Sub

Private Sub dtpPeriodo_Change()
 Dim i As Integer
    For i = 0 To 10
        cmdPasos(i).Enabled = True
    Next i
    
    For i = 0 To 5
        lblResp(i).BackColor = &HC0FFFF
        lblResp(i).Caption = ""
    Next i
End Sub

Private Sub Form_Initialize()
'    Set mobjBusq = New Busqueda
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        Unload Me
    Case Else
        MoverCampo Me, KeyCode, Shift, True
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ImpideSonidoEnter Me, KeyAscii
End Sub

Private Sub Form_Load()
    'Guarda referencia a la empresa de origen
    Set mEmpOrigen = gobjMain.EmpresaActual

    'Fecha de corte asignamos predeterminadamente FechaFinal
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = mbooProcesando
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
'    grd.Move 0, fraPasos.Height + 100, Me.ScaleWidth - 10000, (Me.ScaleHeight - (fraPasos.Height + picBoton.Height) - 105) * 0.75
    grd.Move 0, fraPasos.Height + 100, Me.ScaleWidth, (Me.ScaleHeight - (fraPasos.Height + picBoton.Height) - 105) * 1



    GrdRetVentas.Visible = False
    grdRet.Visible = False
    grdRet.Move 0, grd.Top + grd.Height + 100, Me.ScaleWidth, (Me.ScaleHeight - (fraPasos.Height + picBoton.Height) - 200) * 0.25
    GrdRetVentas.Move grd.Left + grd.Width, fraPasos.Height + 100, Me.ScaleWidth / 2, (Me.ScaleHeight - (fraPasos.Height + picBoton.Height) - 105) * 0.75
    grdCF.Height = 4000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    
    MensajeStatus

    'Cierra y abre de nuevo para que quede como EmpresaActual
    mEmpOrigen.Cerrar
    mEmpOrigen.Abrir
    
    'Libera la referencia
    Set mEmpOrigen = Nothing
    Exit Sub
ErrTrap:
    Set mEmpOrigen = Nothing
    DispErr
    Exit Sub
End Sub


Public Sub MiGetRowsRep(ByVal rs As Recordset, grd As VSFlexGrid)
    grd.LoadArray MiGetRows(rs)
End Sub

Private Sub BuscarComprasATS()
    
    On Error GoTo ErrTrap
        With grd
        .Redraw = False
        .Rows = .FixedRows
        If Not frmB_Trans.Inicio(gobjMain, "IMPCPI2013", dtpPeriodo.value) Then
            grd.SetFocus
        End If
        mObjCond.fecha1 = gobjMain.objCondicion.fecha1
        mObjCond.fecha2 = gobjMain.objCondicion.fecha2
'        MiGetRowsRep gobjMain.EmpresaActual.ConsANCompras2015ParaXML(), grd
        MiGetRowsRep gobjMain.EmpresaActual.Empresa2.ConsANCompras2016ParaXML(), grd
        MiGetRowsRep gobjMain.EmpresaActual.Empresa2.ConsANCompras2016ParaTalon(), grdTalonCP
        MiGetRowsRep gobjMain.EmpresaActual.Empresa2.ConsANRetencionCompras2016paraTALON(False), grdTalonRTPCP
        MiGetRowsRep gobjMain.EmpresaActual.Empresa2.ConsANRetencionCompras2016paraTALON(True), grdTalonRTPIVACP
        
        

        
        'GeneraArchivo
        
        ConfigCols "IMPCPI"
        ConfigCols "IMPCPIR"
        AjustarAutoSize grd, -1, -1
        AjustarAutoSize grdRet, -1, -1
        AjustarAutoSize grdTalonCP, -1, -1
        AjustarAutoSize grdTalonRTPCP, -1, -1
        AjustarAutoSize grdTalonRTPIVACP, -1, -1
        grd.ColWidth(0) = "800"
        grd.ColWidth(COL_C_NOMBRE) = "1500"
        grd.ColWidth(COL_C_NUMAUTOSRI) = "7500"
        grdRet.ColFormat(COL_R_BASE) = "#,#0.00"
        grdRet.ColFormat(COL_R_VALOR) = "#,#0.00"
        
        SubTotalizar (COL_C_CODTIPOCOMP)
        'SubTotalizar (COL_C_TRANS)
        
        Totalizar
        GNPoneNumFila grd, False
        GNPoneNumFila grdRet, False
        GNPoneNumFila grdTalonCP, False
        GNPoneNumFila grdTalonRTPCP, False
        GNPoneNumFila grdTalonRTPIVACP, False
        .Redraw = True
    End With
    Exit Sub
ErrTrap:
    grd.Redraw = True
    DispErr
    Exit Sub
End Sub

Private Function GenerarComprasATS(ByRef cad As String) As Boolean
    On Error GoTo ErrTrap
        GenerarComprasATS = False
        GenerarComprasATS = GeneraArchivoATSComprasXML(cad)
    Exit Function
ErrTrap:
    grd.Redraw = True
    DispErr
    Exit Function
End Function

Private Sub ConfigCols(cad As String)
    Dim s As String, i As Integer
    Dim sql As String, rs As Recordset
    
    
    Select Case cad
    Case "IMPCPI"           'Compras
        s = "^#|TransId|^Fecha|^FechaReg|<CodTrans|<Num|^Doc.|idproveedorref|<RUC|<Proveedor|<Rel.|^Estab|^Punto|^Secuencial|>AutSRI|^Caducidad|^Sustento|^TipoComp|^Tipo Pago|^Pago Exterior|>CodPaisSRI  |^BandDobleTributa |^BandPagoSujRet|>Base Cero|>Base IVA|>Base NO IVA|<Cod ICE|>Val ICE|^BANDREToTRO|>IVA|>Otro Valor"
        '|>Base Ser|<Cod Ser|>Val Ser|>Base Bien|<Cod Bien|>Val Bien|>Base IR|<Cod IR|>Val IR|<NumDocRef|^NumSerieEstabRet|^NumSeriePuntoRet|^NumSecuencialRet|^NumAutSRIRet|^FechaEmisionRet|>Base ICE|<Cod ICE|>Val ICE"
        grd.FormatString = s & "|<         Resultado           "
        AsignarTituloAColKey grd
        
        
        s = "^#|<TipoTrans|^TipoComp|>Cant.Trans|>Base Cero|>Base IVA|>Base NO IVA|>IVA"
        '|>Base Ser|<Cod Ser|>Val Ser|>Base Bien|<Cod Bien|>Val Bien|>Base IR|<Cod IR|>Val IR|<NumDocRef|^NumSerieEstabRet|^NumSeriePuntoRet|^NumSecuencialRet|^NumAutSRIRet|^FechaEmisionRet|>Base ICE|<Cod ICE|>Val ICE"
        grdTalonCP.FormatString = s
        AsignarTituloAColKey grdTalonCP
        
        
        s = "^#|<TipoTrans|^Cod|>Cant.Trans|>Base Imp|>Valor Rete"
        grdTalonRTPCP.FormatString = s
        AsignarTituloAColKey grdTalonRTPCP
        
        s = "^#|<TipoTrans|^Cod|>Cant.Trans|>Base Imp|>Valor Rete"
        grdTalonRTPIVACP.FormatString = s
        AsignarTituloAColKey grdTalonRTPIVACP
        
        
        
        
    
    Case "IMPCPIR"           ' Retencion Compras
        s = "^#|Tipo Ret|<Codigo|<Codigo SRI|>Porcen|<CodTrans|<NumTrans|<RUC|<CodTrans Ret|<NumTrans Ret|^FechaEmisionRet|^NumSerieEstabRet|^NumSeriePuntoRet|^NumSecuencialRet|^NumAutSRIRet....................................................|>Base Ret|>Valor Ret"
        grdRet.FormatString = s
        AsignarTituloAColKey grdRet
    Case "IMPFC"

        s = "^#|<Fecha|^Doc|<IdProvcli|<RUC|<Rel.|^Tipo Comp|^Cant Trans |>Base 0|>Base IVA|>Base NO IVA|>V IVA|>VAlor de IVA|>Valor de ICE|^Tipo Emision"
        For i = 1 To 10
            s = s & "|<Forma" & i
        Next i
        s = s & "|>BaseIVA|>ValorRetIVA|>BaseRenta|>ValorRetRenta"
        grd.FormatString = s & "|<         Resultado           "
        AsignarTituloAColKey grd
    
        grdCF.FormatString = s & "|<         Resultado           "
        AsignarTituloAColKey grdCF
        
        s = "^#|<Fecha|^Doc|<IdProvcli|<RUC|<Cliente|^Tipo Comp|^Cant Trans |>Base 0|>Base IVA|>Base NO IVA|>IVA "
        GrdRetVentas.FormatString = s & "|<         Resultado           "
        
        s = "^#|<TipoTrans|^TipoComp|>Cant.Trans|>Base Cero|>Base IVA|>Base NO IVA|>IVA|>RETNCION IVA|>RETENCION RENTA"
        '|>Base Ser|<Cod Ser|>Val Ser|>Base Bien|<Cod Bien|>Val Bien|>Base IR|<Cod IR|>Val IR|<NumDocRef|^NumSerieEstabRet|^NumSeriePuntoRet|^NumSecuencialRet|^NumAutSRIRet|^FechaEmisionRet|>Base ICE|<Cod ICE|>Val ICE"
        grdTalonFC.FormatString = s
        AsignarTituloAColKey grdTalonFC
        
        
        
        
    Case "IMPFCxE"

        s = "^#|>Establecimiento|^Tipo Comp|>Cant. Documentos|>Base 0|>Base IVA|>Base NO IVA|>ICE|>Total|>Total"
        grd.FormatString = s & "|<         Resultado           "
        AsignarTituloAColKey grd
    
    
    Case "IMPEX"

        s = "^#|<RUC|<Cliente|^Referendo|^Tipo Comp|<Distrito|<Anio|<Regimen|<Correlativo|<Doc. Transporte|<Fecha Embarque|>Valor FOB|>Valor FOB Local|>Establecimiento|>Punto|>Secuencial|>Autorizacion|Fecha Trans|>Valor Factura|<CodPaisDestino "
        grd.FormatString = s & "|<         Resultado           "
        AsignarTituloAColKey grd
    
    
    Case "IMPFCIR"           ' Retencion Ventas
        s = "^#|Tipo Ret|<RUC|>Base Ret|>Valor Ret"
        grdRet.FormatString = s
        AsignarTituloAColKey grdRet
    
    Case "IMPCA"
        s = "^#|<Fecha|<Tipo Comp|^Doc|^Num Serie Estab|^Num Serie Punto|>Num Secuencial|>Num Aut SRI|^Fecha Anulacion "       ' jeaa 27/03/2006
        grd.FormatString = s & "|<         Resultado           "
        AsignarTituloAColKey grd
    
    End Select
   
    
    Select Case cad
    Case "IMPCPI"
            For i = 1 To COL_C_ENOTRARET
                grd.ColHidden(i) = False
'                If i = COL_C_RUC Then i = i + 1
'                grd.ColFormat(i) = flexDTString
            Next i
'            grd.ColFormat(COL_C_RUC) =
            grd.ColHidden(COL_C_TRANSID) = True
            grd.ColHidden(COL_C_IDPROV) = True
            
           
           grd.ColFormat(COL_C_NUMSERESTA) = "000"
           grd.ColFormat(COL_C_NUMSERPUNTO) = "000"
           grd.ColFormat(COL_C_NUMSECUENCIAL) = "000000000"
'           grd.ColFormat(COL_C_NUMAUTOSRI) = "0000000000"
           grd.ColFormat(COL_C_CODSUSTENTO) = "00"
           
            grd.ColFormat(grd.ColIndex("Base Cero")) = "#,#0.00"
            grd.ColFormat(grd.ColIndex("Base IVA")) = "#,#0.00"
            grd.ColFormat(grd.ColIndex("Base NO IVA")) = "#,#0.00"
            grd.ColFormat(COL_C_MONTOICE) = "#,#0.00"
            grd.ColFormat(COL_C_OTROVALOR) = "#,#0.00"
            
            
            grd.ColDataType(grd.ColIndex("BANDREToTRO")) = flexDTString
            'grd.ColFormat(grd.ColIndex("BANDREToTRO")) = "0"
            
            grd.ColDataType(COL_C_PAGOEXTERIOR) = flexDTBoolean
            grd.ColDataType(COL_C_DOBLETRIB) = flexDTBoolean
            grd.ColDataType(COL_C_PAGOSUJRET) = flexDTBoolean
            grd.ColDataType(COL_C_ENOTRARET) = flexDTBoolean
                    
    
            grd.ColData(COL_C_CODTIPOCOMP) = "SubTotal"
            grd.ColData(COL_C_BASE0) = "SubTotal"
            grd.ColData(COL_C_BASE12) = "SubTotal"
            grd.ColData(COL_C_BASENO12) = "SubTotal"
            grd.ColData(COL_C_MONTOICE) = "SubTotal"
            grd.ColData(COL_C_OTROVALOR) = "SubTotal"
            
            
            grdTalonCP.ColFormat(grdTalonCP.ColIndex("Base Cero")) = "#,#0.00"
            grdTalonCP.ColFormat(grdTalonCP.ColIndex("Base IVA")) = "#,#0.00"
            grdTalonCP.ColFormat(grdTalonCP.ColIndex("Base NO IVA")) = "#,#0.00"
            grdTalonCP.ColFormat(grdTalonCP.ColIndex("IVA")) = "#,#0.00"
            grdTalonCP.ColFormat(grdTalonCP.ColIndex("Cant.Trans")) = "#,#0"
            
            grdTalonCP.ColData(grdTalonCP.ColIndex("Base Cero")) = "SubTotal"
            grdTalonCP.ColData(grdTalonCP.ColIndex("Base IVA")) = "SubTotal"
            grdTalonCP.ColData(grdTalonCP.ColIndex("Base NO IVA")) = "SubTotal"
            grdTalonCP.ColData(grdTalonCP.ColIndex("IVA")) = "SubTotal"
            grdTalonCP.ColData(grdTalonCP.ColIndex("Cant.Trans")) = "SubTotal"
            
            
            grdTalonRTPCP.ColFormat(grdTalonRTPCP.ColIndex("Base Imp")) = "#,#0.00"
            grdTalonRTPCP.ColFormat(grdTalonRTPCP.ColIndex("Valor Rete")) = "#,#0.00"
            grdTalonRTPCP.ColFormat(grdTalonRTPCP.ColIndex("Cant.Trans")) = "#,#0"
            
            grdTalonRTPCP.ColData(grdTalonRTPCP.ColIndex("Base Imp")) = "SubTotal"
            grdTalonRTPCP.ColData(grdTalonRTPCP.ColIndex("Valor Rete")) = "SubTotal"


            grdTalonRTPIVACP.ColFormat(grdTalonRTPIVACP.ColIndex("Base Imp")) = "#,#0.00"
            grdTalonRTPIVACP.ColFormat(grdTalonRTPIVACP.ColIndex("Valor Rete")) = "#,#0.00"
            grdTalonRTPIVACP.ColFormat(grdTalonRTPIVACP.ColIndex("Cant.Trans")) = "#,#0"
            
            grdTalonRTPIVACP.ColData(grdTalonRTPIVACP.ColIndex("Base Imp")) = "SubTotal"
            grdTalonRTPIVACP.ColData(grdTalonRTPIVACP.ColIndex("Valor Rete")) = "SubTotal"

            


    Case "IMPCPIR"
            grdRet.ColHidden(COL_R_TRANS) = True
            grdRet.ColHidden(COL_R_NUMTRANS) = True
            grdRet.ColHidden(COL_R_RUC) = True
    Case "IMPFC"
            For i = 1 To COL_V_BASENOIVA
                grd.ColHidden(i) = False
            Next i
            
            
            
            grd.ColDataType(COL_V_FORMAPAGO) = flexDTString
            grd.ColDataType(COL_V_FORMAPAGO + 1) = flexDTString
            grd.ColDataType(COL_V_FORMAPAGO + 2) = flexDTString
            grd.ColDataType(COL_V_FORMAPAGO + 3) = flexDTString
            
            For i = COL_V_FORMAPAGO + 4 To COL_V_FORMAPAGO + 9
                grd.ColHidden(i) = True
            Next i
            
            
            grd.ColHidden(COL_V_IDPROVCLI) = True
            grd.ColHidden(COL_V_BASERETIVA) = True
            grd.ColHidden(COL_V_BASERETRENTA) = True
            
            grd.ColFormat(COL_V_BASE0) = "#,#0.00"
            grd.ColFormat(COL_V_BASEIVA) = "#,#0.00"
            grd.ColFormat(COL_V_BASENOIVA) = "#,#0.00"
            grd.ColFormat(COL_V_VALORIVA) = "#,#0.00"
            grd.ColFormat(COL_V_IVA) = "#,#0.00"
            grd.ColFormat(COL_V_ICE) = "#,#0.00"
            
            
            grd.ColDataType(COL_V_RETRENTA) = flexDTDecimal
            
            grd.ColFormat(COL_V_BASERETIVA) = "#,#0.00"
            grd.ColFormat(COL_V_RETIVA) = "#,#0.00"
            grd.ColFormat(COL_V_BASERETRENTA) = "#,#0.00"
            grd.ColFormat(COL_V_RETRENTA) = "#,#0.00"
            
    
'            grd.ColData(COL_V_CANTRANS) = "SubTotal"
            grd.ColData(COL_V_BASE0) = "SubTotal"
            grd.ColData(COL_V_BASEIVA) = "SubTotal"
            grd.ColData(COL_V_BASENOIVA) = "SubTotal"
            grd.ColData(COL_V_VALORIVA) = "SubTotal"
            grd.ColData(COL_V_CANTRANS) = "SubTotal"
            grd.ColData(COL_V_ICE) = "SubTotal"
            grd.ColData(COL_V_IVA) = "SubTotal"
            
            
            grd.ColData(COL_V_BASERETIVA) = "SubTotal"
            grd.ColData(COL_V_RETIVA) = "SubTotal"
            grd.ColData(COL_V_BASERETRENTA) = "SubTotal"
            grd.ColData(COL_V_RETRENTA) = "SubTotal"
            
            
            grdTalonFC.ColFormat(grdTalonFC.ColIndex("Base Cero")) = "#,#0.00"
            grdTalonFC.ColFormat(grdTalonFC.ColIndex("Base IVA")) = "#,#0.00"
            grdTalonFC.ColFormat(grdTalonFC.ColIndex("Base NO IVA")) = "#,#0.00"
            grdTalonFC.ColFormat(grdTalonFC.ColIndex("IVA")) = "#,#0.00"
            grdTalonFC.ColFormat(grdTalonFC.ColIndex("Cant.Trans")) = "#,#0"
            
            grdTalonFC.ColFormat(grdTalonFC.ColIndex("RETNCION ")) = "#,#0.00"
            grdTalonFC.ColFormat(grdTalonFC.ColIndex("RETENCION RENTA")) = "#,#0.00"
            
            
            
            grdTalonFC.ColData(grdTalonFC.ColIndex("Base Cero")) = "SubTotal"
            grdTalonFC.ColData(grdTalonFC.ColIndex("Base IVA")) = "SubTotal"
            grdTalonFC.ColData(grdTalonFC.ColIndex("Base NO IVA")) = "SubTotal"
            grdTalonFC.ColData(grdTalonFC.ColIndex("IVA")) = "SubTotal"
            grdTalonFC.ColData(grdTalonFC.ColIndex("Cant.Trans")) = "SubTotal"
            
            grdTalonFC.ColData(grdTalonFC.ColIndex("RETNCION ")) = "SubTotal"
            grdTalonFC.ColData(grdTalonFC.ColIndex("RETENCION RENTA")) = "SubTotal"
            
            
            
            
        Case "IMPEX"
            
           grd.ColHidden(COL_E_RUC) = True
           grd.ColFormat(COL_E_NUMSERESTA) = "000"
           grd.ColFormat(COL_E_NUMSERPUNTO) = "000"
           grd.ColFormat(COL_E_NUMSECUENCIAL) = "000000000"
            
            
            grd.ColFormat(COL_E_VALORFOB) = "##0.00"
            grd.ColFormat(COL_E_VALORFOBLOCAL) = "##0.00"
    

            grd.ColData(COL_E_VALORFOB) = "SubTotal"
            grd.ColData(COL_E_VALORFOBLOCAL) = "SubTotal"
            
    Case "IMPFCxE"
    
            For i = 1 To COL_VE_RESP
                grd.ColHidden(i) = False
            Next i
    
            grd.ColFormat(COL_VE_BASE0) = "##0.00"
            grd.ColFormat(COL_VE_BASEIVA) = "##0.00"
            grd.ColFormat(COL_VE_BASENOIVA) = "##0.00"
            grd.ColFormat(COL_VE_TOTAL) = "##0.00"
            grd.ColFormat(COL_VE_ICE) = "##0.00"
    
            grd.ColData(COL_VE_CANTRANS) = "SubTotal"
            grd.ColData(COL_VE_BASE0) = "SubTotal"
            grd.ColData(COL_VE_BASEIVA) = "SubTotal"
            grd.ColData(COL_VE_BASENOIVA) = "SubTotal"
            grd.ColData(COL_VE_TOTAL) = "SubTotal"
            grd.ColData(COL_VE_ICE) = "SubTotal"
            
    
    Case "IMPCA"
            For i = 1 To COL_A_RESP
                grd.ColHidden(i) = False
'                grd.ColFormat(i) = flexDTString
            Next i
            grd.ColFormat(COL_A_NUMESTA) = "000"
           grd.ColFormat(COL_A_NUMPUNTO) = "000"
            grd.ColFormat(COL_A_NUMSECUE) = "0000000"
''            grd.ColFormat(COL_A_NUMAUTO) = "0000000000"
''            grd.ColDataType(COL_A_NUMAUTO) = flexDTString
    End Select
    
    grd.ColSort(1) = flexSortGenericAscending
    grd.ColSort(2) = flexSortGenericAscending
    grd.ColSort(3) = flexSortGenericAscending
    grd.ColSort(4) = flexSortGenericAscending

    AsignarTituloAColKey grd
    grd.SetFocus

End Sub



''''Private Sub GeneraArchivo()
''''    Dim v As Variant, file As String, nombre As String
''''    Dim Filas As Long, Columnas As Long, i As Long, j As Long
''''    On Error GoTo ErrTrap
''''    nombre = "AT" & Format(CStr(Month(mObjCond.Fecha2)), "00") & Year(mObjCond.Fecha2) & ".XML"
''''    file = "c:\" & nombre 'txtCarpeta.Text & Nombre
''''    If ExisteArchivo(file) Then
''''        If MsgBox("El nombre del archivo " & nombre & " ya existe desea sobreescribirlo?", vbYesNo) = vbNo Then
''''            Exit Sub
''''        End If
''''    End If
''''    NumFile = FreeFile
''''    Open file For Output Access Write As #NumFile
'''''     grd.AddItem vbTab & Nombre & vbTab & "Generando  archivo..."
''''    Cadena = GeneraArchivoEncabezado
''''
''''
'''''    grd.AddItem vbTab & Nombre & vbTab & "Generando  archivo..."
''''   Print #NumFile, Cadena
''''
''''    Close NumFile
'''''    grd.textmatrix(i,grd.Rows - 1, grd.Cols - 1) = "Grabado con exito"
''''    Exit Sub
''''ErrTrap:
''''    'grd.TextMatrix(i, grd.Rows - 1, 2) = Err.Description
''''    Close NumFile
''''End Sub

Private Function GeneraArchivoEncabezadoATSXML() As String
    Dim obj As GNOpcion, cad As String, numSucursal As Integer
    cad = "<?xml version=" & """1.0""" & " encoding=" & """UTF-8""" & "" & " standalone=" & """no""" & "?>"
    cad = cad & "<!--  Generado por Ishida Asociados   -->"
    cad = cad & "<!--  Dir: Av. Gonzalez Suarez y Rayoloma Tercer Piso -->"
    cad = cad & "<!--  Telf: 098499003, 072870346, 072871094      -->"
    cad = cad & "<!--  email: ishidacue@hotmail.com, aquizhpe@ibzssoft.com    -->"
    cad = cad & "<!--  www.ibzssoft.com    -->"
    cad = cad & "<!--  Cuenca - Ecuador                -->"
    cad = cad & "<!--  SISTEMAS DE GESTION EMPRESASRIAL-->"
        
    cad = cad & "<iva>"
        
    cad = cad & "<TipoIDInformante> R </TipoIDInformante>"
    cad = cad & "<IdInformante>" & Format(gobjMain.EmpresaActual.GNOpcion.ruc, "0000000000000") & "</IdInformante>"
    cad = cad & "<razonSocial>" & UCase(gobjMain.EmpresaActual.GNOpcion.RazonSocial) & "</razonSocial>"
    cad = cad & "<Anio>" & Year(mObjCond.fecha1) & "</Anio>"
    cad = cad & "<Mes>" & IIf(Len(Month(mObjCond.fecha1)) = 1, "0" & Month(mObjCond.fecha1), Month(mObjCond.fecha1)) & "</Mes>"
    
    numSucursal = gobjMain.EmpresaActual.RecuperaNumeroSucursales
    cad = cad & "<numEstabRuc>" & Format(numSucursal, "000") & "</numEstabRuc>"
    
'    TotalVentas = gobjMain.EmpresaActual.RecuperaNumeroSucursales
    cad = cad & "<totalVentas>" & Format(TotalVentas, "#0.00") & "</totalVentas>"
    cad = cad & "<codigoOperativo>IVA</codigoOperativo>"

'    cad = cad & "<compras>"

    GeneraArchivoEncabezadoATSXML = cad
End Function

Public Function RellenaDer(ByVal s As String, lon As Long) As String
    Dim r As String
    r = "!" & String(lon, "@")
    If Len(s) = 0 Then s = " "
    RellenaDer = Format(s, r)
End Function

Public Function ValidaTelefono(ByVal Tel As String) As String
    Dim c As String
    If Len(Tel) < 6 Then Exit Function
    'asigna caracter
    Select Case Mid(Tel, 1, 2)
            Case "02", "04", "07": c = "2"
            Case "09": c = "9"
            Case Else: c = "-"  'Diego 27 Abril 2004 ' si va jeaa 02/04/04
    End Select
   
    Select Case Len(Tel)
    Case 6: Tel = "07" & c & Tel
    Case 7:
        If InStr("0249", Mid(Tel, 1, 1)) = 0 Then
            Tel = "0" & Mid(Tel, 1, 1) & c & Mid(Tel, 2, Len(Tel))
        Else
            'jeaa 2/06/04
            Tel = "07" & Tel
        End If
    Case 8: Tel = Mid(Tel, 1, 2) & c & Mid(Tel, 3, 8)
    Case 9: If Mid(Tel, 3, 1) <> c Then Tel = Mid(Tel, 1, 2) & c & Mid(Tel, 3, Len(Tel))
    End Select
    
    ValidaTelefono = Tel
End Function



Private Sub cmdExaminarCarpeta_Click(Index As Integer)
    On Error GoTo ErrTrap
    slf.OwnerHWnd = Me.hWnd
    slf.Path = txtCarpeta.Text
    If slf.Browse Then
        txtCarpeta.Text = slf.Path
        txtCarpeta_LostFocus
    End If
    Exit Sub
ErrTrap:
    MsgBox Err.Description, vbInformation
    Exit Sub
End Sub

Private Sub grd_Click()
    Dim rsRet As Recordset
    If NumProc = 1 Or NumProc = 2 Then
        Set rsRet = gobjMain.EmpresaActual.ConsANRetencionCompras2008ParaXML(grd.ValueMatrix(grd.Row, COL_C_TRANSID))
        
        If rsRet.RecordCount = 0 And grd.TextMatrix(grd.Row, COL_C_CODTIPOCOMP) <> "4" Then
            Set rsRet = gobjMain.EmpresaActual.ConsANRetencionCompras2008ParaXMLSinRetencion(grd.ValueMatrix(grd.Row, COL_C_TRANSID))
        End If
        If rsRet.RecordCount > 0 Then
            MiGetRowsRep rsRet, grdRet
        Else
            grdRet.Clear
        End If
    ElseIf NumProc = 3 Or NumProc = 4 Then
        Set rsRet = gobjMain.EmpresaActual.ConsANRetencionVentas2008ParaXML(grd.TextMatrix(grd.Row, COL_V_RUC))
        
        If rsRet.RecordCount > 0 Then
            MiGetRowsRep rsRet, grdRet
        Else
            grdRet.Clear
        End If
    
    End If
End Sub

Private Sub grd_DblClick()
    Dim gnc As GNComprobante, cad As String
    Dim pc As PCProvCli
    Select Case NumProc
    Case 1, 2
        Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(grd.ValueMatrix(grd.Row, COL_C_TRANSID))
        If Not gnc Is Nothing Then
        gnc.RecuperaDetalleTodo
        gnc.BandNoGrabaTransXML = False
        If Not gnc Is Nothing Then
            Select Case grd.col
            Case COL_C_FECHATRANS, COL_C_NUMSERESTA
                    cad = frmDatosAnexos.Inicio(gnc)
                    If cad = "O.K." Then
                        gnc.Grabar False, False
                    End If
            Case COL_C_NUMSERPUNTO, COL_C_NUMSECUENCIAL
                    cad = frmDatosAnexos.Inicio(gnc)
                    If cad = "O.K." Then
                        gnc.Grabar False, False
                    End If
            Case COL_C_NUMAUTOSRI, COL_C_FECHACADUCIDAD
                    cad = frmDatosAnexos.Inicio(gnc)
                    If cad = "O.K." Then
                        gnc.Grabar False, False
                    End If
            Case COL_C_CODSUSTENTO, COL_C_CODTIPOCOMP
                    cad = frmDatosAnexos.Inicio(gnc)
                    If cad = "O.K." Then
                        gnc.Grabar False, False
                    End If
            Case COL_C_TIPODOC, COL_C_RUC, COL_C_NOMBRE
                Set pc = gobjMain.EmpresaActual.RecuperaPCProvCli(grd.TextMatrix(grd.Row, COL_C_RUC))
                'Select Case grd.col
                'Case COL_V_RUC, COL_V_RUC, COL_V_CLIENTE
                    cad = frmDatosPC.Inicio(pc)
                            If cad = "O.K." Then
                                pc.Grabar
                                If NumProc < 3 Then
                                    grd.TextMatrix(grd.Row, COL_C_TIPODOC) = pc.codtipoDocumento
                                Else
                                    grd.TextMatrix(grd.Row, COL_V_TIPODOC) = pc.codtipoDocumento
                                End If
                                
                            End If
                'End Select
            
            End Select
            Set gnc = Nothing
        End If
        End If
    Case 3, 4
        Set pc = gobjMain.EmpresaActual.RecuperaPCProvCli(grd.TextMatrix(grd.Row, COL_V_RUC))
        If Not pc Is Nothing Then
        Select Case grd.col
        Case COL_V_RUC, COL_V_RUC, COL_V_CLIENTE
            cad = frmDatosPC.Inicio(pc)
                    If cad = "O.K." Then
                        pc.Grabar
                    End If
                    grd.TextMatrix(grd.Row, COL_V_TIPODOC) = pc.codtipoDocumento
        End Select
        End If
    End Select
    Set pc = Nothing
End Sub

Private Sub txtCarpeta_LostFocus()
    If Right$(txtCarpeta.Text, 1) <> "\" Then
        txtCarpeta.Text = txtCarpeta.Text & "\"
    End If
    'Luego a actualiza linea de comando
End Sub

Private Function GeneraArchivoATSComprasXML(ByRef cad As String) As Boolean
    Dim cadenaCP As String
    Dim i As Long, j As Long
    Dim vIR As Variant, cadenaCPIR As String
    Dim FilasIR As Long, ColumnasIR As Long, iIR As Long, jIR As Long
    Dim rsRet As Recordset, cadenaCPIVA30 As String
    Dim cadenaCPIVA70 As String, cadenaCPIVA100 As String, cadenaRET As String
    Dim rsNC As Recordset, cadenaNC As String, ret As TSRetencion
    Dim msg As String, bandIgualaFechaCompra_Reten As Boolean, resp As E_MiMsgBox
    Dim m As Integer, n As Integer, codret As String, ane As Anexos, CadenaPagoExt  As String
    Dim CadenaRGasto As String, rsRG As Recordset, contRG As Integer, totalRG As Currency
    Dim pc As PCProvCli, cadenaTipoProv As String
    Dim cadenaCPIVA10 As String
    Dim cadenaCPIVA20 As String
    Dim cadenaCPIVA50 As String
    
    On Error GoTo ErrTrap
    resp = 10
    GeneraArchivoATSComprasXML = True
    grd.Refresh
    'With grd
        
        cadenaCP = "<compras>"
            If grd.Rows < 1 Then
                prg.value = 0
                cadenaCP = cadenaCP & "</compras>"
                cad = cadenaCP
                    GeneraArchivoATSComprasXML = True
                GoTo SiguienteFila
            End If
            prg.max = grd.Rows - 1
            FilaIni = 1
            For i = 1 To grd.Rows - 1
                cadenaTipoProv = ""
                cadenaNC = ""
                bandIgualaFechaCompra_Reten = False
                If grd.IsSubtotal(i) Then

'''''                    grdTalon.Visible = True
'''''                    grdTalon.TextMatrix(FilaTalon, 1) = Right("000" & grd.TextMatrix(i - 1, COL_C_CODTIPOCOMP), 2)
'''''                    grdTalon.TextMatrix(FilaTalon, 2) = ""
'''''                    grdTalon.TextMatrix(FilaTalon, 3) = FILAINI
'''''                    grdTalon.TextMatrix(FilaTalon, 4) = i - 1
'''''                    grdTalon.TextMatrix(FilaTalon, 5) = i - 1 - FILAINI
'''''                    grdTalon.TextMatrix(FilaTalon, 6) = Format(Abs(grd.ValueMatrix(i, COL_C_BASE0)), "#0.00")
'''''                    grdTalon.TextMatrix(FilaTalon, 7) = Format(Abs(grd.ValueMatrix(i, COL_C_BASE12)), "#0.00")
'''''                    grdTalon.TextMatrix(FilaTalon, 8) = Format(Abs(grd.ValueMatrix(i, COL_C_BASENO12)), "#0.00")
'''''                    grdTalon.TextMatrix(FilaTalon, 9) = Format(IIf(Abs(grd.ValueMatrix(i, COL_C_BASE12)) = 0, "0.00", Abs(grd.ValueMatrix(i, COL_C_BASE12)) * (grd.ValueMatrix(i, COL_C_IVA))), "#0.00")
                    FilaIni = i + 1
                    FilaTalon = FilaTalon + 1
                    GoTo SiguienteFila
                End If
                grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbWhite
                prg.value = i
                DoEvents
                cadenaCP = cadenaCP & "<detalleCompras>"
                cadenaCP = cadenaCP & "<codSustento>" & grd.TextMatrix(i, COL_C_CODSUSTENTO) & "</codSustento>"
                
'''        If grd.TextMatrix(i, COL_C_NUMTRANS) = 19375 Then MsgBox "hola"
                
                Select Case grd.TextMatrix(i, COL_C_TIPODOC)
                    Case "R":
                            If Len(grd.TextMatrix(i, COL_C_RUC)) <> 13 Then
                                msg = " El Tipo de Comprobante del Proveedor " & grd.TextMatrix(i, COL_C_NOMBRE) & " es Incorrecto"
                                'MsgBox msg
                                grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                                grd.ShowCell i, grd.ColIndex("Resultado")
                                grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                                GeneraArchivoATSComprasXML = False
                                lblResp(0).Caption = "Error"
                                GoTo SiguienteFila
                            End If
                            cadenaCP = cadenaCP & "<tpIdProv>" & "01" & "</tpIdProv>"
                            
'                            Set pc = gobjMain.EmpresaActual.RecuperaPCProvCliQuick(CDbl(grd.TextMatrix(i, COL_C_IDPROV)))
 '                           cadenaTipoProv = cadenaTipoProv & "<parteRel>" & IIf(pc.BandRelacionado, "SI", "NO") & "</parteRel>"
                            If grd.ValueMatrix(i, COL_C_REL) = 0 Then
                                cadenaTipoProv = cadenaTipoProv & "<parteRel>NO</parteRel>"
                            Else
                                cadenaTipoProv = cadenaTipoProv & "<parteRel>SI</parteRel>"
                            End If
                            
                            
                    Case "C":
                            If Len(grd.TextMatrix(i, COL_C_RUC)) <> 10 Then
                                msg = " El Tipo de Comprobante del Proveedor " & grd.TextMatrix(i, COL_C_NOMBRE) & " es Incorrecto"
                                'MsgBox msg
                                grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                                grd.ShowCell i, grd.ColIndex("Resultado")
                                grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                                GeneraArchivoATSComprasXML = False
                                lblResp(0).Caption = "Error"
                                GoTo SiguienteFila
                            End If
                        cadenaCP = cadenaCP & "<tpIdProv>" & "02" & "</tpIdProv>"
                        
'                            Set pc = gobjMain.EmpresaActual.RecuperaPCProvCliQuick(CDbl(grd.TextMatrix(i, COL_C_IDPROV)))
 '                           cadenaTipoProv = cadenaTipoProv & "<parteRel>" & IIf(pc.BandRelacionado, "SI", "NO") & "</parteRel>"
                        
                              If grd.ValueMatrix(i, COL_C_REL) = 0 Then
                                cadenaTipoProv = cadenaTipoProv & "<parteRel>NO</parteRel>"
                            Else
                                cadenaTipoProv = cadenaTipoProv & "<parteRel>SI</parteRel>"
                            End If
                      
                        
                    Case "P":
                        Set pc = gobjMain.EmpresaActual.RecuperaPCProvCliQuick(CDbl(grd.TextMatrix(i, COL_C_IDPROV)))
                        cadenaCP = cadenaCP & "<tpIdProv>" & "03" & "</tpIdProv>"
                        If pc.TipoProvCli = "RPN" Then
                            cadenaTipoProv = cadenaTipoProv & "<tipoProv>01</tipoProv>"
                        ElseIf pc.TipoProvCli = "RSO" Then
                            cadenaTipoProv = cadenaTipoProv & "<tipoProv>02</tipoProv>"
                        End If
                        cadenaTipoProv = cadenaTipoProv & "<denoProv>" & grd.TextMatrix(i, COL_C_NOMBRE) & "</denoProv>"
                        cadenaTipoProv = cadenaTipoProv & "<parteRel>" & IIf(pc.BandRelacionado, "SI", "NO") & "</parteRel>"
                        Set pc = Nothing
                    Case Else
                            msg = " El Proveedor " & grd.TextMatrix(i, COL_C_NOMBRE) & " Tipo de Documento Incorrecto"
                            'MsgBox msg
                            grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                            grd.ShowCell i, grd.ColIndex("Resultado")
                            grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                            GeneraArchivoATSComprasXML = False
                            lblResp(0).Caption = "Error"
                            GoTo SiguienteFila
                End Select
                
'                If grd.TextMatrix(i, COL_C_NUMSECUENCIAL) = "000000327" Then MsgBox ""
                cadenaCP = cadenaCP & "<idProv>" & grd.TextMatrix(i, COL_C_RUC) & "</idProv>"
                If Mid$(grd.TextMatrix(i, COL_C_CODTIPOCOMP), 1, 1) = "0" Then
                    cadenaCP = cadenaCP & "<tipoComprobante>" & Format(Mid$(grd.TextMatrix(i, COL_C_CODTIPOCOMP), 2, 1), "00") & "</tipoComprobante>"
                Else
                    If grd.TextMatrix(i, COL_C_CODTIPOCOMP) = "2" Then
                        If grd.TextMatrix(i, COL_C_CODSUSTENTO) = "01" Then
                            msg = " El Sustento " & grd.TextMatrix(i, COL_C_CODSUSTENTO) & ", no va con comprobante " & grd.TextMatrix(i, COL_C_CODTIPOCOMP)
                            'MsgBox msg
                            grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                            grd.ShowCell i, grd.ColIndex("Resultado")
                            grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                            GeneraArchivoATSComprasXML = False
                            lblResp(0).Caption = "Error"
                            GoTo SiguienteFila
                        End If
                        cadenaCP = cadenaCP & "<tipoComprobante>" & Format(grd.TextMatrix(i, COL_C_CODTIPOCOMP), "00") & "</tipoComprobante>"
                    Else
                        cadenaCP = cadenaCP & "<tipoComprobante>" & Format(grd.TextMatrix(i, COL_C_CODTIPOCOMP), "00") & "</tipoComprobante>"
                    End If
                End If
                
                cadenaCP = cadenaCP & cadenaTipoProv
                
                cadenaCP = cadenaCP & "<fechaRegistro>" & grd.TextMatrix(i, COL_C_FECHAREGISTRO) & "</fechaRegistro>"
                If Len(grd.TextMatrix(i, COL_C_NUMSERESTA)) <> 3 Or grd.ValueMatrix(i, COL_C_NUMSERESTA) = 0 Then
                            msg = " El Numero de Serie Establecimiento " & grd.TextMatrix(i, COL_C_NUMSERESTA) & " Incorrecto"
                            'MsgBox msg
                            grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                            grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                            grd.ShowCell i, grd.ColIndex("Resultado")
                            GeneraArchivoATSComprasXML = False
                            lblResp(0).Caption = "Error"
                            GoTo SiguienteFila
                Else
                    cadenaCP = cadenaCP & "<establecimiento>" & grd.TextMatrix(i, COL_C_NUMSERESTA) & "</establecimiento>"
                End If
                If Len(grd.TextMatrix(i, COL_C_NUMSERPUNTO)) <> 3 Or grd.ValueMatrix(i, COL_C_NUMSERPUNTO) = 0 Then
                            msg = " El Numero de Serie Punto " & grd.TextMatrix(i, COL_C_NUMSERPUNTO) & " Incorrecto"
                            'MsgBox msg
                            grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                            grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                            grd.ShowCell i, grd.ColIndex("Resultado")
                            GeneraArchivoATSComprasXML = False
                            lblResp(0).Caption = "Error"
                            GoTo SiguienteFila
                Else
                    cadenaCP = cadenaCP & "<puntoEmision>" & grd.TextMatrix(i, COL_C_NUMSERPUNTO) & "</puntoEmision>"
                End If
                If grd.TextMatrix(i, COL_C_NUMSECUENCIAL) <> "000000000" Then
                    cadenaCP = cadenaCP & "<secuencial>" & grd.TextMatrix(i, COL_C_NUMSECUENCIAL) & "</secuencial>"
                Else
                            msg = " El Numero de Secuencia no puede ser " & grd.TextMatrix(i, COL_C_NUMSECUENCIAL) & " esta Incorrecto"
                            'MsgBox msg
                            grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                            grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                            grd.ShowCell i, grd.ColIndex("Resultado")
                            GeneraArchivoATSComprasXML = False
                            lblResp(0).Caption = "Error"
                            GoTo SiguienteFila
                
                End If
                cadenaCP = cadenaCP & "<fechaEmision>" & grd.TextMatrix(i, COL_C_FECHATRANS) & "</fechaEmision>"
                If Len(grd.TextMatrix(i, COL_C_NUMAUTOSRI)) > 50 Or grd.ValueMatrix(i, COL_C_NUMAUTOSRI) < 1 Then
                            msg = " El Numero de Autorización SRI " & grd.TextMatrix(i, COL_C_NUMAUTOSRI) & " Incorrecto"
                            'MsgBox msg
                            grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                            grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                            grd.ShowCell i, grd.ColIndex("Resultado")
                            GeneraArchivoATSComprasXML = False
                            lblResp(0).Caption = "Error"
                            GoTo SiguienteFila
                Else
                    cadenaCP = cadenaCP & "<autorizacion>" & grd.TextMatrix(i, COL_C_NUMAUTOSRI) & "</autorizacion>"
                End If
                cadenaCP = cadenaCP & "<baseNoGraIva>" & Format(Abs(grd.ValueMatrix(i, COL_C_BASENO12)), "#0.00") & "</baseNoGraIva>"
                cadenaCP = cadenaCP & "<baseImponible>" & Format(Abs(grd.ValueMatrix(i, COL_C_BASE0)), "#0.00") & "</baseImponible>"
                
                cadenaCP = cadenaCP & "<baseImpGrav>" & Format(Abs(grd.ValueMatrix(i, COL_C_BASE12)), "#0.00") & "</baseImpGrav>"
                cadenaCP = cadenaCP & "<baseImpExe>" & Format(0, "#0.00") & "</baseImpExe>"
                cadenaCP = cadenaCP & "<montoIce>" & Format(Abs(grd.ValueMatrix(i, COL_C_MONTOICE)), "#0.00") & "</montoIce>"
                
                cadenaCP = cadenaCP & "<montoIva>" & Format(IIf(Abs(grd.ValueMatrix(i, COL_C_BASE12)) = 0, "0.00", Abs(grd.ValueMatrix(i, COL_C_BASE12)) * (grd.ValueMatrix(i, COL_C_IVA))), "#0.00") & "</montoIva>"
                
                
                
          '      If grd.TextMatrix(i, COL_C_CODTIPOCOMP) <> "41" Then
                'retencion IVA
                Set rsRet = gobjMain.EmpresaActual.ConsANRetencionCompras2008ParaXML(grd.ValueMatrix(i, COL_C_TRANSID))
                    If rsRet.RecordCount = 0 And grd.TextMatrix(i, COL_C_CODTIPOCOMP) <> "4" And grd.TextMatrix(i, COL_C_CODTIPOCOMP) <> "5" Then
                        Set rsRet = gobjMain.EmpresaActual.ConsANRetencionCompras2008ParaXMLSinRetencion(grd.ValueMatrix(i, COL_C_TRANSID))
                    End If
                cadenaCPIR = "<air>"
                
                cadenaCPIVA10 = "<valRetBien10>0.00</valRetBien10>" 'AUC08/09/2015
                cadenaCPIVA20 = "<valRetServ20>0.00</valRetServ20>" 'AUC08/09/2015
                cadenaCPIVA50 = "<valRetServ50> 0.00 </valRetServ50>"
                cadenaCPIVA30 = "<valorRetBienes> 0.00 </valorRetBienes>"
                cadenaCPIVA70 = "<valorRetServicios> 0.00 </valorRetServicios>"
                cadenaCPIVA100 = "<valRetServ100> 0.00 </valRetServ100>"
                
                 cadenaRET = ""
                 CadenaPagoExt = ""
                 CadenaRGasto = ""
                 
                 
                 
                 If grd.TextMatrix(i, COL_C_PAGOEXTERIOR) = "0" Then
                    CadenaPagoExt = "<pagoExterior>"
                    CadenaPagoExt = CadenaPagoExt & "<pagoLocExt>01</pagoLocExt>"
                    CadenaPagoExt = CadenaPagoExt & "<paisEfecPago>NA</paisEfecPago>"
                    CadenaPagoExt = CadenaPagoExt & "<aplicConvDobTrib>NA</aplicConvDobTrib>"
                    CadenaPagoExt = CadenaPagoExt & "<pagExtSujRetNorLeg>NA</pagExtSujRetNorLeg>"
                    CadenaPagoExt = CadenaPagoExt & "</pagoExterior>"
                Else
                    CadenaPagoExt = "<pagoExterior>"
                    CadenaPagoExt = CadenaPagoExt & "<pagoLocExt>02</pagoLocExt>"
                    CadenaPagoExt = CadenaPagoExt & "<tipoRegi>01</tipoRegi>"
                    CadenaPagoExt = CadenaPagoExt & "<paisEfecPagoGen>" & grd.TextMatrix(i, COL_C_CODPAIS) & "</paisEfecPagoGen>"
                    CadenaPagoExt = CadenaPagoExt & "<paisEfecPago>" & grd.TextMatrix(i, COL_C_CODPAIS) & "</paisEfecPago>"
                    CadenaPagoExt = CadenaPagoExt & "<aplicConvDobTrib>" & IIf(grd.TextMatrix(i, COL_C_DOBLETRIB) = 0, "NO", "SI") & "</aplicConvDobTrib>"
                    If grd.TextMatrix(i, COL_C_DOBLETRIB) <> 0 Then
                        CadenaPagoExt = CadenaPagoExt & "<pagExtSujRetNorLeg>NA</pagExtSujRetNorLeg>"
                    Else
                        CadenaPagoExt = CadenaPagoExt & "<pagExtSujRetNorLeg>" & IIf(grd.TextMatrix(i, COL_C_PAGOSUJRET) = 0, "NO", "SI") & "</pagExtSujRetNorLeg>"
                    End If
                    CadenaPagoExt = CadenaPagoExt & "</pagoExterior>"
                End If

'''''              anulado ats junio 2016
                If ((Abs(grd.ValueMatrix(i, COL_C_BASE12)) + Abs(grd.ValueMatrix(i, COL_C_BASE0)) + Abs(grd.ValueMatrix(i, COL_C_BASENO12)) + Abs(grd.ValueMatrix(i, COL_C_BASE12) * (grd.ValueMatrix(i, COL_C_IVA)))) > "999.99" And grd.TextMatrix(i, COL_C_CODTIPOCOMP) <> "4") Or (grd.ValueMatrix(i, COL_C_MONTOICE) > 0) Then
                    If Len(grd.TextMatrix(i, COL_C_TIPOPAGO)) <> 0 Then
                        CadenaPagoExt = CadenaPagoExt & "<formasDePago>"
                        CadenaPagoExt = CadenaPagoExt & "<formaPago>" & grd.TextMatrix(i, COL_C_TIPOPAGO) & "</formaPago>"
                        CadenaPagoExt = CadenaPagoExt & "</formasDePago>"
                    Else
                            msg = " Falta seleccionar forma de pago"
                            'MsgBox msg
                            grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                            grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                            grd.ShowCell i, grd.ColIndex("Resultado")
                            GeneraArchivoATSComprasXML = False
                            lblResp(0).Caption = "Error"
                            GoTo SiguienteFila
                    End If
               End If

                If rsRet.RecordCount > 0 Then
                    MiGetRowsRep rsRet, grdRet
'                    If grd.TextMatrix(i, COL_C_NUMTRANS) = "3892" Then MsgBox "hola"
                        For j = 1 To grdRet.Rows - 1
                    
                        For m = 1 To grdRet.Rows - 1
                            codret = grdRet.TextMatrix(m, COL_R_CODIGORET)
                            For n = m + 1 To grdRet.Rows - 1
                                If codret = grdRet.TextMatrix(n, COL_R_CODIGORET) Then
                                            msg = " Retención  " & grdRet.TextMatrix(j, COL_R_NUMTRANS) & " Cód. de Ret.  Duplicado " & codret
                                            grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                                            grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                                            grd.ShowCell i, grd.ColIndex("Resultado")
                                            GeneraArchivoATSComprasXML = False
                                            lblResp(0).Caption = "Error"
                                            GoTo SiguienteFila
                                End If
                            Next n
                        Next m
                            
                    
                    
                            If grd.TextMatrix(i, COL_C_RUC) <> grdRet.TextMatrix(j, COL_R_RUC) Then
                                            msg = " Compra  " & grd.TextMatrix(i, COL_C_RUC) & " RTP " & grdRet.TextMatrix(j, COL_C_RUC)
                                            grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                                            grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                                            grd.ShowCell i, grd.ColIndex("Resultado")
                                            GeneraArchivoATSComprasXML = False
                                            lblResp(0).Caption = "Error"
                                            GoTo SiguienteFila
                            
                            Else
                            
                                If (grd.TextMatrix(i, COL_C_TRANS) = grdRet.TextMatrix(j, COL_R_TRANS)) And (grd.TextMatrix(i, COL_C_NUMTRANS) = grdRet.TextMatrix(j, COL_R_NUMTRANS)) Then
                                    
                                    If grdRet.TextMatrix(j, COL_R_TIPO) = -1 Then
                                        Select Case grdRet.ValueMatrix(j, COL_R_PORCEN)
                                        Case 10
                                            cadenaCPIVA10 = "<valRetBien10>" & Format(grdRet.ValueMatrix(j, COL_R_VALOR), "#0.00") & "</valRetBien10>"
                                        Case 20
                                            cadenaCPIVA20 = "<valRetServ20>" & Format(grdRet.ValueMatrix(j, COL_R_VALOR), "#0.00") & "</valRetServ20>"
                                        Case 30
                                            cadenaCPIVA30 = "<valorRetBienes>" & Format(grdRet.ValueMatrix(j, COL_R_VALOR), "#0.00") & "</valorRetBienes>"
                                        Case 50
                                            cadenaCPIVA50 = "<valRetServ50>" & Format(grdRet.ValueMatrix(j, COL_R_VALOR), "#0.00") & "</valRetServ50>"
                                        Case 70
                                            cadenaCPIVA70 = "<valorRetServicios>" & Format(grdRet.ValueMatrix(j, COL_R_VALOR), "#0.00") & "</valorRetServicios>"
                                        Case 100
                                            cadenaCPIVA100 = "<valRetServ100>" & Format(grdRet.ValueMatrix(j, COL_R_VALOR), "#0.00") & "</valRetServ100>"
                                        End Select
                                        
                                    Else
                                        'valores renta
'                                        If i = 29 Then MsgBox "HOLA"
                                        If Len(grdRet.TextMatrix(j, COL_R_CODIGOSRI)) = 0 Then
                                             Set ret = gobjMain.EmpresaActual.RecuperaTSRetencion(grdRet.TextMatrix(j, COL_R_CODIGORET))
                                             If Not ret Is Nothing Then
                                                Set ane = gobjMain.EmpresaActual.RecuperaAnexosRetIR(Mid$(ret.CodRetencion, 3, 3))
                                                If Not ane Is Nothing Then
                                                    ret.CodAnexo = ane.CodRetencion
                                                    ret.Grabar
                                                Else
                                                    msg = " Falta crear código " & Mid$(grdRet.TextMatrix(j, COL_R_CODIGORET), 3, 3)
                                                    grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                                                    grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                                                    grd.ShowCell i, grd.ColIndex("Resultado")
                                                    GeneraArchivoATSComprasXML = False
                                                    lblResp(0).Caption = "Error"
                                                    GoTo SiguienteFila
                                                
                                                
                                                End If
                                             End If
                                        
                                        End If
                                        If grd.ValueMatrix(i, 24) > -1 Then
                                            If grd.ValueMatrix(i, COL_C_MONTOICE + 1) = 0 Then
                                                If (grd.TextMatrix(i, COL_C_CODTIPOCOMP) <> "4" Or grd.TextMatrix(i, COL_C_CODTIPOCOMP) = "5") Then
                                                    If Len(grdRet.TextMatrix(j, COL_R_CODIGOSRI)) = 0 Then
                                                        msg = " Falta enlace Cat.Retenciones " & grdRet.TextMatrix(j, COL_R_CODIGORET)
                                                        grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                                                        grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                                                        grd.ShowCell i, grd.ColIndex("Resultado")
                                                        GeneraArchivoATSComprasXML = False
                                                        lblResp(0).Caption = "Error"
                                                        GoTo SiguienteFila
                                                    End If
                                                End If
                                                cadenaCPIR = cadenaCPIR & "<detalleAir>"
                                                cadenaCPIR = cadenaCPIR & "<codRetAir>" & grdRet.TextMatrix(j, COL_R_CODIGOSRI) & "</codRetAir>"
                                                cadenaCPIR = cadenaCPIR & "<baseImpAir>" & Format(grdRet.ValueMatrix(j, COL_R_BASE), "#0.00") & "</baseImpAir>"
                                                cadenaCPIR = cadenaCPIR & "<porcentajeAir>" & Format(grdRet.TextMatrix(j, COL_R_PORCEN), "#0.00") & "</porcentajeAir>"
                                                cadenaCPIR = cadenaCPIR & "<valRetAir>" & Format(grdRet.ValueMatrix(j, COL_R_VALOR), "#0.00") & "</valRetAir>"
                                                cadenaCPIR = cadenaCPIR & "</detalleAir>"
                                            End If
                                        End If
                                    End If
                                End If
                                
                totalRG = 0
                If grd.TextMatrix(i, COL_C_CODTIPOCOMP) = "41" Then
                    Set rsRG = gobjMain.EmpresaActual.ConsReembolsoGastos2013ParaXML(grd.ValueMatrix(i, COL_C_TRANSID))
                    CadenaRGasto = "<reembolsos>"
                    totalRG = 0
                    If rsRG.RecordCount > 0 Then
                        rsRG.MoveFirst
                        For contRG = 1 To rsRG.RecordCount
                            CadenaRGasto = CadenaRGasto & "<reembolso>"
                            CadenaRGasto = CadenaRGasto & "<tipoComprobanteReemb>" & Format(rsRG.Fields(16), "00") & "</tipoComprobanteReemb>"
                            Select Case rsRG.Fields(17)
                                Case "R": CadenaRGasto = CadenaRGasto & "<tpIdProvReemb>01</tpIdProvReemb>"
                                Case "C": CadenaRGasto = CadenaRGasto & "<tpIdProvReemb>02</tpIdProvReemb>"
                                Case "P": CadenaRGasto = CadenaRGasto & "<tpIdProvReemb>03</tpIdProvReemb>"
                                Case Else
                                                    msg = " Tipo de Comprobante en Reembolso de Gastos Errado " & grdRet.TextMatrix(j, COL_R_CODIGORET)
                                                    grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                                                    grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                                                    grd.ShowCell i, grd.ColIndex("Resultado")
                                                    GeneraArchivoATSComprasXML = False
                                                    lblResp(0).Caption = "Error"
                                                    GoTo SiguienteFila
                                
                            End Select
                            CadenaRGasto = CadenaRGasto & "<idProvReemb>" & rsRG.Fields(4) & "</idProvReemb>"
                            CadenaRGasto = CadenaRGasto & "<establecimientoReemb>" & rsRG.Fields(5) & "</establecimientoReemb>"
                            CadenaRGasto = CadenaRGasto & "<puntoEmisionReemb>" & rsRG.Fields(6) & "</puntoEmisionReemb>"
                            CadenaRGasto = CadenaRGasto & "<secuencialReemb>" & rsRG.Fields(7) & "</secuencialReemb>"
                            CadenaRGasto = CadenaRGasto & "<fechaEmisionReemb>" & rsRG.Fields(8) & "</fechaEmisionReemb>"
                            CadenaRGasto = CadenaRGasto & "<autorizacionReemb>" & rsRG.Fields(9) & "</autorizacionReemb>"
                            CadenaRGasto = CadenaRGasto & "<baseImponibleReemb>" & Format(rsRG.Fields(10), "#0.00") & "</baseImponibleReemb>"
                            CadenaRGasto = CadenaRGasto & "<baseImpGravReemb>" & Format(rsRG.Fields(11), "#0.00") & "</baseImpGravReemb>"
                            CadenaRGasto = CadenaRGasto & "<baseNoGraIvaReemb>" & Format(rsRG.Fields(12), "#0.00") & "</baseNoGraIvaReemb>"
                            CadenaRGasto = CadenaRGasto & "<baseImpExeReemb>" & Format(0, "#0.00") & "</baseImpExeReemb>" 'AUC
                            CadenaRGasto = CadenaRGasto & "<montoIceRemb>" & Format(rsRG.Fields(13), "#0.00") & "</montoIceRemb>"
                            CadenaRGasto = CadenaRGasto & "<montoIvaRemb>" & Format(rsRG.Fields(14), "#0.00") & "</montoIvaRemb>"
                            CadenaRGasto = CadenaRGasto & "</reembolso>"
                            '"<baseImpExeReemb>"
                            totalRG = totalRG + rsRG.Fields(10) + rsRG.Fields(11) + rsRG.Fields(12)
                            rsRG.MoveNext
                        Next contRG
                    End If
                    CadenaRGasto = CadenaRGasto & "</reembolsos>"
                End If
                
                                cadenaRET = ""
                                'If grd.TextMatrix(i, COL_C_MONTOICE) > 0 Then
                                If (grdRet.ValueMatrix(j, COL_R_PORCEN)) > 0 Then
                                    cadenaRET = "<estabRetencion1>" & grdRet.TextMatrix(j, COL_R_NUMEST) & "</estabRetencion1>"
                                    cadenaRET = cadenaRET & "<ptoEmiRetencion1>" & grdRet.TextMatrix(j, COL_R_NUMPTO) & "</ptoEmiRetencion1>"
                                    cadenaRET = cadenaRET & "<secRetencion1>" & grdRet.TextMatrix(j, COL_R_NUMRET) & "</secRetencion1>"
                                    cadenaRET = cadenaRET & "<autRetencion1>" & grdRet.TextMatrix(j, COL_R_NUMAUTO) & "</autRetencion1>"
                                    cadenaRET = cadenaRET & "<fechaEmiRet1>" & grdRet.TextMatrix(j, COL_R_FECHARET) & "</fechaEmiRet1>"
                                    If resp = mmsgSiTodo Then
                                        bandIgualaFechaCompra_Reten = True
                                    Else
                                        If Not Len(grd.TextMatrix(i, COL_C_NUMAUTOSRI)) = 37 Then
                                            If CDate(grdRet.TextMatrix(j, COL_R_FECHARET)) < CDate(grd.TextMatrix(i, COL_C_FECHAREGISTRO)) Then
                                                msg = "La fecha de la Retención " & grdRet.TextMatrix(j, COL_R_RETTRANS) & "-" & _
                                                            grdRet.TextMatrix(j, COL_R_RETNUMTRANS) & _
                                                            " no puede ser menor a la fecha de la Compra " & _
                                                            grd.TextMatrix(i, COL_C_TRANS) & "-" & _
                                                            grd.TextMatrix(i, COL_C_NUMTRANS) & Chr(13) & _
                                                            "Desea que para el anexo se igaule la fecha"
                                                resp = frmMiMsgBox.MiMsgBox(msg, "Fechas")
                                                If resp = 1 Then
                                                    bandIgualaFechaCompra_Reten = True
                                                ElseIf resp = vbYes Then
                                                    bandIgualaFechaCompra_Reten = True
                                                Else
                                                    bandIgualaFechaCompra_Reten = False
                                                End If
                                            End If
                                        End If
                                    End If
                                 '   End If
                            
                                    If CDate(grdRet.TextMatrix(j, COL_R_FECHARET)) < CDate(grd.TextMatrix(i, COL_C_FECHAREGISTRO)) And bandIgualaFechaCompra_Reten = False Then
                                        msg = "La fecha de la Retención " & grdRet.TextMatrix(j, COL_R_RETTRANS) & "-" & _
                                                    grdRet.TextMatrix(j, COL_R_RETNUMTRANS) & _
                                                    " no puede ser menor a la fecha de la Compra " & _
                                                    grd.TextMatrix(i, COL_C_TRANS) & "-" & _
                                                    grd.TextMatrix(i, COL_C_NUMTRANS)
                                                    
                                                    
                                        grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                                        grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                                        grd.ShowCell i, grd.ColIndex("Resultado")
                                        lblResp(0).Caption = "Error"
                                        GeneraArchivoATSComprasXML = False
                                        GoTo SiguienteFila
                                    End If
                                Else
                                    
                                End If
                                If CDate(grd.TextMatrix(i, COL_C_FECHAREGISTRO)) < CDate(grd.TextMatrix(i, COL_C_FECHATRANS)) Then
                                    msg = "La fecha de registro de la Transaccion " & _
                                    grd.TextMatrix(i, COL_C_TRANS) & "-" & _
                                    grd.TextMatrix(i, COL_C_NUMTRANS) _
                                    & " debe ser menor o igual a la fecha de registro "
                                    'MsgBox msg
                                    grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                                    grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                                    grd.ShowCell i, grd.ColIndex("Resultado")
                                    GeneraArchivoATSComprasXML = False
                                    lblResp(0).Caption = "Error"
                                    GoTo SiguienteFila
                                End If
                                End If
                            Next j
                        
                        Else
                            For j = grdRet.Rows - 1 To 1 Step -1
                                grdRet.RemoveItem (j)
                            Next j
                        End If
                    

                If grd.TextMatrix(i, COL_C_CODTIPOCOMP) = "4" Or grd.TextMatrix(i, COL_C_CODTIPOCOMP) = "5" Then
                    Set rsNC = gobjMain.EmpresaActual.ConsANNCCompras2008ParaXML(grd.ValueMatrix(i, COL_C_TRANSID))
                    If rsNC.RecordCount = 0 Then
                    Else
                        cadenaNC = "<docModificado>" & Format(rsNC.Fields(0), "00") & "</docModificado>"
                        cadenaNC = cadenaNC & "<estabModificado>" & rsNC.Fields(1) & "</estabModificado>"
                        cadenaNC = cadenaNC & "<ptoEmiModificado>" & rsNC.Fields(2) & "</ptoEmiModificado>"
                        cadenaNC = cadenaNC & "<secModificado>" & rsNC.Fields(3) & "</secModificado>"
                        cadenaNC = cadenaNC & "<autModificado>" & rsNC.Fields(4) & "</autModificado>"
                    End If
                Else
                    cadenaNC = ""
                End If
                cadenaCPIR = cadenaCPIR & "</air>"
                cadenaCP = cadenaCP & cadenaCPIVA10 & cadenaCPIVA20 & cadenaCPIVA30 & cadenaCPIVA50 & cadenaCPIVA70 & cadenaCPIVA100
                If grd.TextMatrix(i, COL_C_CODTIPOCOMP) <> "41" Then
                    cadenaCP = cadenaCP & "<totbasesImpReemb> 0.00 </totbasesImpReemb>" 'AUC
                Else
                    cadenaCP = cadenaCP & "<totbasesImpReemb> " & Format(totalRG, "#0.00") & "</totbasesImpReemb>" 'AUC
                End If
                cadenaCP = cadenaCP & CadenaPagoExt
                cadenaCP = cadenaCP & cadenaCPIR & cadenaRET & CadenaRGasto
                cadenaCP = cadenaCP & cadenaNC
                cadenaCP = cadenaCP & "</detalleCompras>"
                grd.ShowCell i, grd.ColIndex("Resultado")
                grd.TextMatrix(i, grd.ColIndex("Resultado")) = " OK "
                grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbWhite
     '       End With
            GoTo SiguienteFila
            'Next i
        'End If
    'End With
    Exit Function
SiguienteFila:
    Next i
    grd.ColWidth(grd.ColIndex("Resultado")) = 5000
    prg.value = 0
    If Len(lblResp(0).Caption) = 0 Then
        lblResp(0).Caption = "OK."
        cadenaCP = cadenaCP & "</compras>"
    Else
        cadenaCP = ""
    End If
    cad = cadenaCP
    'GeneraArchivoATSComprasXML = True
Exit Function
cancelado:
    GeneraArchivoATSComprasXML = False
ErrTrap:
    grd.TextMatrix(grd.Rows - 1, 2) = Err.Description
    GeneraArchivoATSComprasXML = False
End Function

Private Function BuscarVentasATS()
    Dim sql As String, rs As Recordset, i As Integer
    On Error GoTo ErrTrap
        With grd
        .Redraw = False
        .Rows = .FixedRows
        If Not frmB_Trans.Inicio(gobjMain, "IMPFC", dtpPeriodo.value) Then
            grd.SetFocus
        End If
        mObjCond.fecha1 = gobjMain.objCondicion.fecha1
        mObjCond.fecha2 = gobjMain.objCondicion.fecha2
        MiGetRowsRep gobjMain.EmpresaActual.Empresa2.ConsANVentas2016paraXML(), grd
        MiGetRowsRep gobjMain.EmpresaActual.Empresa2.ConsANVentas2016paraTALON(), grdTalonFC
'        MiGetRowsRep gobjMain.EmpresaActual.ConsANTotalRetencionVentas2008ParaXML, GrdRetVentas



        'GeneraArchivo

        ConfigCols "IMPFC"
        ConfigCols "IMPFCIR"
        AjustarAutoSize grd, -1, -1
        AjustarAutoSize grdRet, -1, -1
        AjustarAutoSize grdTalonFC, -1, -1
        grd.ColWidth(0) = "800"
        grd.ColHidden(COL_V_VALORIVA) = True
        SubTotalizar (COL_V_TIPOCOMP)
        Totalizar

        GNPoneNumFila grd, False
        GNPoneNumFila grdRet, False
        GNPoneNumFila grdTalonFC, False


        .Redraw = True

        
   End With

    Exit Function
ErrTrap:
    grd.Redraw = True
    DispErr
    Exit Function
End Function

Private Function GenerarVentasATS(ByRef cad As String) As Boolean
    On Error GoTo ErrTrap
        GenerarVentasATS = False
        GenerarVentasATS = GeneraArchivoATSVentasXML(cad)
    Exit Function
ErrTrap:
    grd.Redraw = True
    DispErr
    Exit Function
End Function



Private Function GeneraArchivoATSVentasXML(ByRef cad As String) As Boolean
    Dim cadenaFC As String, cadenaFCIVA  As String
    Dim i As Long, j As Long, k As Long
    Dim vIR As Variant, cadenaFCIR As String
    Dim FilasIR As Long, ColumnasIR As Long, iIR As Long, jIR As Long
    Dim rsRet As Recordset, cadenaFCIVA30 As String
    Dim cadenaFCIVA70 As String, cadenaFCIVA100 As String
    Dim rsNC As Recordset, cadenaNC As String
    Dim msg As String, pc As PCProvCli, bandCF As Boolean, filaCF As Integer
    Dim cadenaF As String
    Dim cadenaICE As String
    Dim cadenaFormap As String, cadfp As String
    On Error GoTo ErrTrap
    GeneraArchivoATSVentasXML = True
    bandCF = False
    filaCF = 1
    For j = 1 To grdCF.Rows - 1
        grdCF.RemoveItem 1
    Next j
    
    For j = 1 To GrdRetVentas.Rows - 1
        GrdRetVentas.TextMatrix(j, 8) = ""
    Next j
    
    
        grd.Refresh
        'Exit Function
        cadenaF = "<ventas>"

            If grd.Rows < 1 Then
                prg.value = 0
                cadenaF = cadenaFC & "</ventas>"
                cad = cadenaF
                GeneraArchivoATSVentasXML = True
                GoTo SiguienteFila
            End If


            prg.max = grd.Rows - 1
            For i = 1 To grd.Rows - 1
                If grd.IsSubtotal(i) Then GoTo SiguienteFila
                grd.ShowCell i, 1
                prg.value = i
                DoEvents
                cadenaFC = ""
                cadenaICE = ""
                If chkConsFinal.value = vbChecked Then
                    
                    If (grd.TextMatrix(i, COL_V_TIPODOC) = "F" Or grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed) Or grd.TextMatrix(i, COL_V_RUC) = "9999999999999" Then
                        bandCF = True
                        grdCF.AddItem ""
                        grdCF.TextMatrix(filaCF, 0) = i
                        grdCF.TextMatrix(filaCF, COL_V_TIPODOC) = grd.TextMatrix(i, COL_V_TIPODOC)

                        grdCF.TextMatrix(filaCF, COL_V_RUC) = grd.TextMatrix(i, COL_V_RUC)
                        grdCF.TextMatrix(filaCF, COL_V_TIPOCOMP) = grd.TextMatrix(i, COL_V_TIPOCOMP)
                        grdCF.TextMatrix(filaCF, COL_V_CLIENTE) = grd.TextMatrix(i, COL_V_TIPOCOMP)
                        grdCF.TextMatrix(filaCF, COL_V_CANTRANS) = grd.ValueMatrix(i, COL_V_CANTRANS)
                        grdCF.TextMatrix(filaCF, COL_V_BASE0) = grd.ValueMatrix(i, COL_V_BASE0)
                        grdCF.TextMatrix(filaCF, COL_V_BASEIVA) = grd.ValueMatrix(i, COL_V_BASEIVA)
                        grdCF.TextMatrix(filaCF, COL_V_BASENOIVA) = grd.ValueMatrix(i, COL_V_BASENOIVA)
                        filaCF = filaCF + 1
                            GoTo SiguienteFila
                        
                    End If
                    
                End If

                cadenaFC = cadenaFC & "<detalleVentas>"
                Select Case grd.TextMatrix(i, COL_V_TIPODOC)
                    Case "R":                     cadenaFC = cadenaFC & "<tpIdCliente>" & "04" & "</tpIdCliente>"
                    Case "C":                     cadenaFC = cadenaFC & "<tpIdCliente>" & "05" & "</tpIdCliente>"
                    Case "P":                     cadenaFC = cadenaFC & "<tpIdCliente>" & "06" & "</tpIdCliente>"
                    Case "F":                     cadenaFC = cadenaFC & "<tpIdCliente>" & "07" & "</tpIdCliente>"
                    Case "T":
                            msg = " El Cliente " & grd.TextMatrix(i, COL_V_CLIENTE) & " el tipo de Documento selecciona do es Valido"
                            grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                            grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                            grd.ShowCell i, grd.ColIndex("Resultado")
                            GeneraArchivoATSVentasXML = True
                            lblResp(1).Caption = "Error"
                            chkConsFinal.Visible = True
                            GoTo SiguienteFila

                    
                    Case Else
                            
                            'cadenaFC = Mid$(cadenaFC, 1, Len(cadenaFC) - Len("<detalleVentas>") + 1)
                            msg = " El Cliente " & grd.TextMatrix(i, COL_V_CLIENTE) & " No tiene seleccionado el tipo de Documento"
                            grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                            grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                            grd.ShowCell i, grd.ColIndex("Resultado")
                            GeneraArchivoATSVentasXML = True
                            lblResp(1).Caption = "Error"
                            chkConsFinal.Visible = True
                            GoTo SiguienteFila
                        
                End Select
                
                cadenaFC = cadenaFC & "<idCliente>" & grd.TextMatrix(i, COL_V_RUC) & "</idCliente>"
                
'                If grd.TextMatrix(i, COL_V_RUC) = "9999999999999" Then MsgBox ""
                
               
'                Set pc = gobjMain.EmpresaActual.RecuperaPCProvClixRUC(grd.TextMatrix(i, COL_V_RUC), True, False, False)
'                If Not pc Is Nothing Then
'                    Select Case pc.codtipoDocumento
'                        Case "R", "C", "P"
'                            If Not pc.BandRelacionado Then
'                                cadenaFC = cadenaFC & "<parteRelVtas>" & "NO" & "</parteRelVtas>"   'auc
'                            Else
'                                cadenaFC = cadenaFC & "<parteRelVtas>" & "SI" & "</parteRelVtas>"   'auc
'                            End If
'                        Case "F"
'                        Case Else
'                            cadenaFC = cadenaFC & "<parteRelVtas>" & "NO" & "</parteRelVtas>"
'                    End Select
'                Else
'                    If grd.TextMatrix(i, COL_V_RUC) <> "9999999999999" Then
'                        cadenaFC = cadenaFC & "<parteRelVtas>" & "NO" & "</parteRelVtas>"   'auc
'                    End If
'                End If
'                Set pc = Nothing

            If grd.TextMatrix(i, COL_V_RUC) <> "9999999999999" Then
                If grd.ValueMatrix(i, COL_V_CLIENTE) = 2 Then
                ElseIf grd.ValueMatrix(i, COL_V_CLIENTE) = 0 Then
                    cadenaFC = cadenaFC & "<parteRelVtas>" & "NO" & "</parteRelVtas>"   'auc
                Else
                    cadenaFC = cadenaFC & "<parteRelVtas>" & "SI" & "</parteRelVtas>"   'auc
                End If
            End If

            If grd.TextMatrix(i, COL_V_TIPODOC) = "P" Then
                Set pc = gobjMain.EmpresaActual.RecuperaPCProvClixRUC(grd.TextMatrix(i, COL_V_RUC), True, False, False)
                If Not pc Is Nothing Then
                    If pc.TipoProvCli = "RPN" Then
                        cadenaFC = cadenaFC & "<tipoCliente>01</tipoCliente>"
                    ElseIf pc.TipoProvCli = "RSO" Then
                        cadenaFC = cadenaFC & "<tipoCliente>02</tipoCliente>"
                    Else
                            msg = " El Cliente " & grd.TextMatrix(i, COL_V_CLIENTE) & " esta con pasaporte y falta seleccionar el tipo"
                            grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                            grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                            grd.ShowCell i, grd.ColIndex("Resultado")
                            GeneraArchivoATSVentasXML = True
                            lblResp(1).Caption = "Error"
                            chkConsFinal.Visible = True
                            GoTo SiguienteFila
                    
                    End If
                    cadenaFC = cadenaFC & "<denoCli>" & pc.nombre & " </denoCli>"
                End If
                Set pc = Nothing
            End If

                cadenaFC = cadenaFC & "<tipoComprobante>" & Format(grd.TextMatrix(i, COL_V_TIPOCOMP), "00") & "</tipoComprobante>"
                cadenaFC = cadenaFC & "<tipoEmision>" & grd.TextMatrix(i, COL_V_TIPOEMI) & "</tipoEmision>"
                cadenaFC = cadenaFC & "<numeroComprobantes>" & grd.TextMatrix(i, COL_V_CANTRANS) & "</numeroComprobantes>"
                cadenaFC = cadenaFC & "<baseNoGraIva>" & Format(Abs(grd.ValueMatrix(i, COL_V_BASENOIVA)), "#0.00") & "</baseNoGraIva>"
                cadenaFC = cadenaFC & "<baseImponible>" & Format(Abs(grd.ValueMatrix(i, COL_V_BASE0)), "#0.00") & "</baseImponible>"
                
                If grd.ValueMatrix(i, COL_V_ICE) <> 0 Then
                    cadenaFC = cadenaFC & "<baseImpGrav>" & Format(Abs(grd.ValueMatrix(i, COL_V_BASEIVA)) + Abs(grd.ValueMatrix(i, COL_V_ICE)), "#0.00") & "</baseImpGrav>"
                Else
                    cadenaFC = cadenaFC & "<baseImpGrav>" & Format(Abs(grd.ValueMatrix(i, COL_V_BASEIVA)), "#0.00") & "</baseImpGrav>"
                End If
                'cadenaFC = cadenaFC & "<montoIva>" & Format(IIf(Abs(grd.ValueMatrix(i, COL_V_BASEIVA)) = 0, "0.00", Abs(grd.ValueMatrix(i, COL_V_BASEIVA)) * (grd.ValueMatrix(i, COL_V_IVA))), "#0.00") & "</montoIva>"
                
                cadenaFC = cadenaFC & "<montoIva>" & Format(Abs(grd.ValueMatrix(i, COL_V_IVA)), "#0.00") & "</montoIva>"

                'cadenaFC = cadenaFC & "<tipoCompe> 0.00</tipoCompe>"
                'cadenaFC = cadenaFC & "<monto> 0.00</monto>"
                cadenaFC = cadenaFC & "<montoIce>" & Format(Abs(grd.ValueMatrix(i, COL_V_ICE)), "#0.00") & "</montoIce>"
               
               
                If grd.ValueMatrix(i, COL_V_RETIVA) <> 0 Or grd.ValueMatrix(i, COL_V_RETRENTA) <> 0 Then
                    cadenaFCIVA = "<valorRetIva>" & Format(grd.ValueMatrix(i, COL_V_RETIVA), "#0.00") & "</valorRetIva>"
                    cadenaFCIR = "<valorRetRenta>" & Format(grd.ValueMatrix(i, COL_V_RETRENTA), "#0.00") & "</valorRetRenta>"
                Else
                    cadenaFCIVA = "<valorRetIva>0.00</valorRetIva>"
                    cadenaFCIR = "<valorRetRenta>0.00</valorRetRenta>"
                End If
                
                
                
                
 
                'retencion IVA
                'If grd.ValueMatrix(i, COL_V_TIPOCOMP) = 18 And grd.TextMatrix(i, COL_V_TIPODOC) = "R" Then
''''                If grd.ValueMatrix(i, COL_V_BASERETIVA) <> 0 Or grd.ValueMatrix(i, COL_V_BASERETRENTA) <> 0 Then
''''
''''                    Set rsRet = gobjMain.EmpresaActual.ConsANRetencionVentas2008ParaXML(grd.TextMatrix(i, COL_V_RUC))
''''
''''                    If rsRet.RecordCount > 0 Then
''''                        MiGetRowsRep rsRet, grdRet
''''
''''                            For j = 1 To grdRet.Rows - 1
'''''                                If j = 2 Then MsgBox ""
''''                                If grd.TextMatrix(i, COL_V_RUC) = grdRet.TextMatrix(j, COL_RF_RUC) Then
''''                                    If grdRet.TextMatrix(j, COL_RF_TIPO) = -1 Then
''''                                        'valores iva
''''                                        cadenaFCIVA = "<valorRetIva>" & Format(grdRet.ValueMatrix(j, COL_RF_VALOR), "#0.00") & "</valorRetIva>"
''''                                    Else
''''                                        'valores renta
''''                                        cadenaFCIR = "<valorRetRenta>" & Format(grdRet.ValueMatrix(j, COL_RF_VALOR), "#0.00") & "</valorRetRenta>"
''''                                    End If
''''                                End If
''''                                'busca en tablas de retencion
''''                                For k = 1 To GrdRetVentas.Rows - 1
''''                                    If grdRet.TextMatrix(j, COL_RF_RUC) = GrdRetVentas.TextMatrix(k, COL_V_RUC) Then
''''                                        If grdRet.TextMatrix(j, COL_RF_TIPO) = GrdRetVentas.TextMatrix(k, 5) Then
''''                                            If grdRet.ValueMatrix(j, COL_RF_VALOR - 1) = GrdRetVentas.ValueMatrix(k, 6) Then
''''                                                GrdRetVentas.TextMatrix(k, 8) = "OK"
''''                                                GrdRetVentas.RemoveItem k
''''                                                GrdRetVentas.Refresh
''''                                                Exit For
''''                                            End If
''''                                        End If
''''                                    End If
''''                                Next k
''''                            Next j
''''
''''
''''
''''                    End If
''''                Else
''''                    For j = grdRet.Rows - 1 To 1 Step -1
''''                        grdRet.RemoveItem (j)
''''                    Next j
''''                End If
                
                cadenaFormap = ""
                If grd.ValueMatrix(i, COL_V_TIPOCOMP) = 18 Or grd.ValueMatrix(i, COL_V_TIPOCOMP) = 5 Then
                    If grd.ValueMatrix(i, COL_V_CANTRANS) > 0 Then
                        cadenaFormap = "<formasDePago>"
                        For k = COL_V_FORMAPAGO To COL_V_FORMAPAGO + 4
                            If Len(grd.TextMatrix(i, k)) > 0 Then
                                cadfp = grd.TextMatrix(i, k)
                                If InStr(1, cadenaFormap, cadfp) = 0 Then
                                    cadenaFormap = cadenaFormap & "<formaPago>" & cadfp & "</formaPago>"
                                End If
                            End If
                        Next k
                        If Len(cadenaFormap) = 14 Then
                            cadenaFormap = cadenaFormap & "<formaPago>01</formaPago>"
                        End If
                        cadenaFormap = cadenaFormap & "</formasDePago>"
                    Else
                        cadenaFormap = "<formasDePago> <formaPago>01</formaPago> </formasDePago>"
                    End If
                Else
                    cadenaFormap = ""
                End If
                
                cadenaFC = cadenaFC & cadenaFCIVA
                cadenaFC = cadenaFC & cadenaFCIR
                cadenaFC = cadenaFC & cadenaFormap
                cadenaFC = cadenaFC & "</detalleVentas>"
                cadenaF = cadenaF & cadenaFC
                grd.ShowCell i, grd.ColIndex("Resultado")
                grd.TextMatrix(i, grd.ColIndex("Resultado")) = " OK "
        
SiguienteFila:
    Next i
        SubTotalizarCF (COL_V_TIPOCOMP)
        TotalizarCF
    grd.ColWidth(grd.ColIndex("Resultado")) = 5000
    prg.value = 0
'    cadenaF = cadenaF & GeneraArchivoATSVentasXMLSoloRetencion
    If bandCF Then
        cadenaF = cadenaF & GeneraArchivoATSVentasXMLCF
        lblResp(1).Caption = "OK."
        cadenaF = cadenaF & "</ventas>"
    Else
    If Len(lblResp(1).Caption) = 0 Then
        lblResp(1).Caption = "OK."
        cadenaF = cadenaF & "</ventas>"
    Else
        cadenaFC = ""
    End If
    End If
    cad = cadenaF
    'TotalVentas = grd.ValueMatrix(grd.Rows - 1, COL_V_BASE0) + grd.ValueMatrix(grd.Rows - 1, COL_V_BASEIVA) + grd.ValueMatrix(grd.Rows - 1, COL_V_BASENOIVA) '
    TotalVentas = 0
    For i = 1 To grd.Rows - 1
        If grd.TextMatrix(i, COL_V_TIPOEMI) = "F" Then
            If grd.ValueMatrix(i, COL_V_TIPOCOMP) = 4 Then
                TotalVentas = TotalVentas + grd.ValueMatrix(i, COL_V_BASE0) + grd.ValueMatrix(i, COL_V_BASEIVA) + grd.ValueMatrix(i, COL_V_BASENOIVA) - grd.ValueMatrix(i, COL_V_ICE)
            Else
                TotalVentas = TotalVentas + grd.ValueMatrix(i, COL_V_BASE0) + grd.ValueMatrix(i, COL_V_BASEIVA) + grd.ValueMatrix(i, COL_V_BASENOIVA) + grd.ValueMatrix(i, COL_V_ICE)
            End If
        End If
    Next i
    Exit Function
cancelado:
    GeneraArchivoATSVentasXML = False
ErrTrap:
    grd.TextMatrix(grd.Rows - 1, 2) = Err.Description
    GeneraArchivoATSVentasXML = False
End Function


Private Function BuscarANuladosATS()
    On Error GoTo ErrTrap
        With grd
        .Redraw = False
        .Rows = .FixedRows
        If Not frmB_Trans.Inicio(gobjMain, "IMPAN", dtpPeriodo.value) Then
            grd.SetFocus
        End If
        mObjCond.fecha1 = gobjMain.objCondicion.fecha1
        mObjCond.fecha2 = gobjMain.objCondicion.fecha2
        MiGetRowsRep gobjMain.EmpresaActual.ConsANComprobantesAnulado2008ParaXML(), grd
        
        lblnumAnulados.Caption = grd.Rows - 1

        'GeneraArchivo

        ConfigCols "IMPCA"
        AjustarAutoSize grd, -1, -1
        grd.ColWidth(0) = "500"


        GNPoneNumFila grd, False
        GNPoneNumFila grdRet, False

        .Redraw = True
    End With

    Exit Function
ErrTrap:
    grd.Redraw = True
    DispErr
    Exit Function
End Function


Private Function GenerarANuladosATS(ByRef cad As String) As Boolean
    On Error GoTo ErrTrap
        GenerarANuladosATS = False
        GenerarANuladosATS = GeneraArchivoATSAnuladosXML(cad)
    Exit Function
ErrTrap:
    grd.Redraw = True
    DispErr
    Exit Function
End Function


Private Function GeneraArchivoATSAnuladosXML(ByRef cad As String) As Boolean
    Dim cadenaAN As String
    Dim i As Long, j As Long
    Dim msg As String
    On Error GoTo ErrTrap
    GeneraArchivoATSAnuladosXML = True
    grd.Refresh
        cadenaAN = "<anulados>"
        If grd.Rows < 1 Then
            prg.value = 0
            cadenaAN = cadenaAN & "</anulados>"
            cad = cadenaAN
            GeneraArchivoATSAnuladosXML = True
            GoTo SiguienteFila
        End If
            prg.max = grd.Rows - 1
            For i = 1 To grd.Rows - 1
                If grd.IsSubtotal(i) Then GoTo SiguienteFila
                prg.value = i
                DoEvents
                cadenaAN = cadenaAN & "<detalleAnulados>"
                If grd.ValueMatrix(i, COL_A_TIPODOC) <> 0 Then
                    cadenaAN = cadenaAN & "<tipoComprobante>" & Format(grd.TextMatrix(i, COL_A_TIPODOC), "00") & "</tipoComprobante>"
                Else
                        msg = " El Tipo de Comprobante " & grd.TextMatrix(i, COL_A_TIPODOC) & " Incorrecto"
                            'MsgBox msg
                            grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                            grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                            grd.ShowCell i, grd.ColIndex("Resultado")
                            GeneraArchivoATSAnuladosXML = False
                            lblResp(4).Caption = "Error"
                            GoTo SiguienteFila
               End If
                cadenaAN = cadenaAN & "<establecimiento>" & grd.TextMatrix(i, COL_A_NUMESTA) & "</establecimiento>"
                cadenaAN = cadenaAN & "<puntoEmision>" & grd.TextMatrix(i, COL_A_NUMPUNTO) & "</puntoEmision>"
                cadenaAN = cadenaAN & "<secuencialInicio>" & grd.TextMatrix(i, COL_A_NUMSECUE) & "</secuencialInicio>"
                cadenaAN = cadenaAN & "<secuencialFin>" & grd.TextMatrix(i, COL_A_NUMSECUE) & "</secuencialFin>"
                If (Len(grd.TextMatrix(i, COL_A_NUMAUTO)) <> 10 And Len(grd.TextMatrix(i, COL_A_NUMAUTO)) <> 37) Or (grd.ValueMatrix(i, COL_A_NUMAUTO) < 1) Then
                            msg = " El Numero de Autorización SRI " & grd.TextMatrix(i, COL_A_NUMAUTO) & " Incorrecto"
                            'MsgBox msg
                            grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                            grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                            grd.ShowCell i, grd.ColIndex("Resultado")
                            GeneraArchivoATSAnuladosXML = False
                            lblResp(4).Caption = "Error"
                            GoTo SiguienteFila
                Else
                    cadenaAN = cadenaAN & "<autorizacion>" & grd.TextMatrix(i, COL_A_NUMAUTO) & "</autorizacion>"
                End If
                
                cadenaAN = cadenaAN & "</detalleAnulados>"
                grd.ShowCell i, grd.ColIndex("Resultado")
                grd.TextMatrix(i, grd.ColIndex("Resultado")) = " OK "

SiguienteFila:
    Next i
    grd.ColWidth(grd.ColIndex("Resultado")) = 5000
    prg.value = 0
    If Len(lblResp(4).Caption) = 0 Then
        lblResp(4).Caption = "OK."
        cadenaAN = cadenaAN & "</anulados>"
    Else
        cadenaAN = ""
    End If
    cad = cadenaAN
    Exit Function
    
cancelado:
    GeneraArchivoATSAnuladosXML = False
ErrTrap:
    grd.TextMatrix(grd.Rows - 1, 2) = Err.Description
    GeneraArchivoATSAnuladosXML = False
End Function

Private Function BuscarComprasREOC()
    
    On Error GoTo ErrTrap
        With grd
        .Redraw = False
        .Rows = .FixedRows
        If Not frmB_Trans.Inicio(gobjMain, "IMPCPI", dtpPeriodo.value) Then
            grd.SetFocus
        End If
        mObjCond.fecha1 = gobjMain.objCondicion.fecha1
        mObjCond.fecha2 = gobjMain.objCondicion.fecha2
        MiGetRowsRep gobjMain.EmpresaActual.ConsANCompras2008ParaXML(), grd
        
        'GeneraArchivo
        
        ConfigCols "IMPCPI"
        ConfigCols "IMPCPIR"
        AjustarAutoSize grd, -1, -1
        AjustarAutoSize grdRet, -1, -1
        grd.ColWidth(0) = "500"
        grdRet.ColFormat(COL_R_BASE) = "#,#0.00"
        grdRet.ColFormat(COL_R_VALOR) = "#,#0.00"
        
        
        GNPoneNumFila grd, False
        GNPoneNumFila grdRet, False
        
        .Redraw = True
    End With
    Exit Function
ErrTrap:
    grd.Redraw = True
    DispErr
    Exit Function
End Function
    
Private Function GenerarComprasREOC(ByRef cad As String) As Boolean
    
    On Error GoTo ErrTrap
        GenerarComprasREOC = False
        GenerarComprasREOC = GeneraArchivoREOCComprasXML(cad)
    
    Exit Function
ErrTrap:
    DispErr
    Exit Function
End Function

Private Function GeneraArchivoEncabezadoREOCXML() As String
    Dim obj As GNOpcion, cad As String
    cad = "<?xml version=" & """1.0""" & " encoding=" & """ISO-8859-1""" & "" & " standalone=" & """yes""" & "?>"
    cad = cad & "<!--  Generado por Ishida Asociados   -->"
    cad = cad & "<!--  Dir: Av. Espana  y Elia Liut Aeropuerto Mariscal Lamar Segundo Piso -->"
    cad = cad & "<!--  Telf: 098499003, 072870346      -->"
    cad = cad & "<!--  email: ishidacue@hotmail.com    -->"
    cad = cad & "<!--  Cuenca - Ecuador                -->"
    cad = cad & "<!--  SISTEMAS DE GESTION EMPRESASRIAL-->"
        
    cad = cad & "<reoc>"
        
    cad = cad & "<numeroRuc>" & Format(gobjMain.EmpresaActual.GNOpcion.ruc, "0000000000000") & "</numeroRuc>"
    cad = cad & "<anio>" & Year(mObjCond.fecha1) & "</anio>"
    cad = cad & "<mes>" & IIf(Len(Month(mObjCond.fecha1)) = 1, "0" & Month(mObjCond.fecha1), Month(mObjCond.fecha1)) & "</mes>"
    GeneraArchivoEncabezadoREOCXML = cad
End Function



'''''Private Function GeneraArchivoREOCComprasXML(ByRef cad As String) As Boolean
'''''    Dim cadenaCP As String
'''''    Dim i As Long, j As Long
'''''    Dim vIR As Variant, cadenaCPIR As String
'''''    Dim FilasIR As Long, ColumnasIR As Long, iIR As Long, jIR As Long
'''''    Dim rsRet As Recordset, cadenaCPIVA30 As String
'''''    Dim cadenaCPIVA70 As String, cadenaCPIVA100 As String, cadenaRET As String
'''''    Dim rsNC As Recordset, cadenaNC As String
'''''
'''''    On Error GoTo ErrTrap
'''''    GeneraArchivoREOCComprasXML = False
'''''    grd.Refresh
'''''    With grd
'''''        cadenaCP = "<compras>"
'''''        If grd.Rows > 1 Then
'''''            prg.max = .Rows - 1
'''''            For i = 1 To grd.Rows - 1
'''''                prg.value = i
'''''                DoEvents
'''''                Set rsRet = gobjMain.EmpresaActual.ConsANRetencionCompras2008ParaXML(.ValueMatrix(i, COL_C_TRANSID))
'''''                 cadenaRET = ""
'''''                If rsRet.RecordCount > 0 Then
'''''                    MiGetRowsRep rsRet, grdRet
'''''                    cadenaCPIR = "<air>"
'''''                    For j = 1 To grdRet.Rows - 1
'''''                        If (.TextMatrix(i, COL_C_TRANS) = grdRet.TextMatrix(j, COL_R_TRANS)) And (.TextMatrix(i, COL_C_NUMTRANS) = grdRet.TextMatrix(j, COL_R_NUMTRANS)) And (.TextMatrix(i, COL_C_RUC) = grdRet.TextMatrix(j, COL_R_RUC)) Then
'''''
'''''                            If grdRet.TextMatrix(j, COL_R_TIPO) <> -1 Then
'''''                                If j = 1 Then
'''''                                    cadenaCP = cadenaCP & "<detalleCompras>"
'''''                                    Select Case .TextMatrix(i, COL_C_TIPODOC)
'''''                                        Case "R":                     cadenaCP = cadenaCP & "<tpIdProv>" & "01" & "</tpIdProv>"
'''''                                        Case "C":                     cadenaCP = cadenaCP & "<tpIdProv>" & "02" & "</tpIdProv>"
'''''                                        Case "P":                     cadenaCP = cadenaCP & "<tpIdProv>" & "03" & "</tpIdProv>"
'''''                                        Case Else
'''''                                                MsgBox " El Proveedor " & .TextMatrix(i, COL_C_NOMBRE) & " No tiene seleccionado el tipo de Documento"
'''''                                                .TextMatrix(i, grd.ColIndex("Resultado")) = " Error "
'''''                                                grd.ShowCell i, grd.ColIndex("Resultado")
''''''                                                GoTo cancelado
'''''                                    End Select
'''''                                    cadenaCP = cadenaCP & "<idProv>" & .TextMatrix(i, COL_C_RUC) & "</idProv>"
'''''                                    If Mid$(.TextMatrix(i, COL_C_CODTIPOCOMP), 1, 1) = "0" Then
'''''                                        cadenaCP = cadenaCP & "<tipoComp>" & Mid$(.TextMatrix(i, COL_C_CODTIPOCOMP), 2, 1) & "</tipoComp>"
'''''                                    Else
'''''                                        cadenaCP = cadenaCP & "<tipoComp>" & .TextMatrix(i, COL_C_CODTIPOCOMP) & "</tipoComp>"
'''''                                    End If
'''''                                    cadenaCP = cadenaCP & "<aut>" & .TextMatrix(i, COL_C_NUMAUTOSRI) & "</aut>"
'''''                                    cadenaCP = cadenaCP & "<estab>" & .TextMatrix(i, COL_C_NUMSERESTA) & "</estab>"
'''''                                    cadenaCP = cadenaCP & "<ptoEmi>" & .TextMatrix(i, COL_C_NUMSERPUNTO) & "</ptoEmi>"
'''''                                    cadenaCP = cadenaCP & "<sec>" & .TextMatrix(i, COL_C_NUMSECUENCIAL) & "</sec>"
'''''                                    cadenaCP = cadenaCP & "<fechaEmiCom>" & .TextMatrix(i, COL_C_FECHATRANS) & "</fechaEmiCom>"
'''''
'''''                                    cadenaCPIR = "<air>"
'''''                                End If
'''''                                'valores renta
'''''                                cadenaCPIR = cadenaCPIR & "<detalleAir>"
'''''                                cadenaCPIR = cadenaCPIR & "<codRetAir>" & grdRet.TextMatrix(j, COL_R_CODIGOSRI) & "</codRetAir>"
'''''                                cadenaCPIR = cadenaCPIR & "<porcentaje>" & grdRet.ValueMatrix(j, COL_R_PORCEN) & "</porcentaje>"
'''''                                cadenaCPIR = cadenaCPIR & "<base0>" & Format(.ValueMatrix(i, COL_C_BASE0), "#0.00") & "</base0>"
'''''                                cadenaCPIR = cadenaCPIR & "<baseGrav>" & Format(grdRet.ValueMatrix(j, COL_R_BASE), "#0.00") & "</baseGrav>"
'''''                                cadenaCPIR = cadenaCPIR & "<baseNoGrav>" & Format(.ValueMatrix(i, COL_C_BASENO12), "#0.00") & "</baseNoGrav>"
'''''                                cadenaCPIR = cadenaCPIR & "<valRetAir>" & Format(grdRet.ValueMatrix(j, COL_R_VALOR), "#0.00") & "</valRetAir>"
'''''                                cadenaCPIR = cadenaCPIR & "</detalleAir>"
'''''                                cadenaRET = cadenaRET & "<autRet>" & grdRet.TextMatrix(j, COL_R_NUMAUTO) & "</autRet>"
'''''                                cadenaRET = cadenaRET & "<estabRet>" & grdRet.TextMatrix(j, COL_R_NUMEST) & "</estabRet>"
'''''                                cadenaRET = cadenaRET & "<ptoEmiRet>" & grdRet.TextMatrix(j, COL_R_NUMPTO) & "</ptoEmiRet>"
'''''                                cadenaRET = cadenaRET & "<secRet>" & grdRet.TextMatrix(j, COL_R_NUMRET) & "</secRet>"
'''''                                cadenaRET = cadenaRET & "<fechaEmiRet>" & grdRet.TextMatrix(j, COL_R_FECHARET) & "</fechaEmiRet>"
'''''
'''''                            End If
'''''                        End If
'''''                    Next j
'''''                Else
'''''                    For j = grdRet.Rows - 1 To 1 Step -1
'''''                        grdRet.RemoveItem (j)
'''''                    Next j
'''''                End If
'''''                If rsRet.RecordCount > 0 Then
'''''                    cadenaCPIR = cadenaCPIR & "</air>"
'''''                    cadenaCP = cadenaCP & cadenaCPIR & cadenaRET
'''''                    cadenaCP = cadenaCP & "</detalleCompras>"
'''''                End If
'''''                grd.ShowCell i, grd.ColIndex("Resultado")
'''''                .TextMatrix(i, grd.ColIndex("Resultado")) = " OK "
'''''            Next i
'''''        End If
'''''    End With
'''''    prg.value = 0
'''''    cadenaCP = cadenaCP & "</compras>"
'''''    cad = cadenaCP
'''''    GeneraArchivoREOCComprasXML = True
'''''    Exit Function
'''''cancelado:
'''''    GeneraArchivoREOCComprasXML = False
'''''ErrTrap:
'''''    grd.TextMatrix(grd.Rows - 1, 2) = Err.Description
'''''    GeneraArchivoREOCComprasXML = False
'''''End Function

Private Function GeneraArchivoREOCComprasXML(ByRef cad As String) As Boolean
    Dim cadenaCP As String
    Dim i As Long, j As Long
    Dim vIR As Variant, cadenaCPIR As String
    Dim FilasIR As Long, ColumnasIR As Long, iIR As Long, jIR As Long
    Dim rsRet As Recordset, cadenaCPIVA30 As String
    Dim cadenaCPIVA70 As String, cadenaCPIVA100 As String, cadenaRET As String
    Dim rsNC As Recordset, cadenaNC As String
    Dim msg As String
    
    On Error GoTo ErrTrap
    GeneraArchivoREOCComprasXML = True
    grd.Refresh
    'With grd
        cadenaCP = "<compras>"
            If grd.Rows < 1 Then
                prg.value = 0
                cadenaCP = cadenaCP & "</compras>"
                cad = cadenaCP
                GeneraArchivoREOCComprasXML = True
                GoTo SiguienteFila
            End If
            prg.max = grd.Rows - 1
            For i = 1 To grd.Rows - 1
                If grd.IsSubtotal(i) Then GoTo SiguienteFila
                prg.value = i
                DoEvents
                cadenaCP = cadenaCP & "<detalleCompras>"
                Select Case grd.TextMatrix(i, COL_C_TIPODOC)
                    Case "R":                     cadenaCP = cadenaCP & "<tpIdProv>" & "01" & "</tpIdProv>"
                    Case "C":                     cadenaCP = cadenaCP & "<tpIdProv>" & "02" & "</tpIdProv>"
                    Case "P":                     cadenaCP = cadenaCP & "<tpIdProv>" & "03" & "</tpIdProv>"
                    Case Else
                            msg = " El Proveedor " & grd.TextMatrix(i, COL_C_NOMBRE) & " Tipo de Documento Incorrecto"
                            grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                            grd.ShowCell i, grd.ColIndex("Resultado")
                            grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                            GeneraArchivoREOCComprasXML = False
                            lblPasos(5).Caption = "Error"
                            GoTo SiguienteFila
                End Select
                
                
                
                cadenaCP = cadenaCP & "<idProv>" & grd.TextMatrix(i, COL_C_RUC) & "</idProv>"
                If Mid$(grd.TextMatrix(i, COL_C_CODTIPOCOMP), 1, 1) = "0" Then
                    cadenaCP = cadenaCP & "<tipoComprobante>" & Mid$(grd.TextMatrix(i, COL_C_CODTIPOCOMP), 2, 1) & "</tipoComprobante>"
                Else
                    cadenaCP = cadenaCP & "<tipoComp>" & grd.TextMatrix(i, COL_C_CODTIPOCOMP) & "</tipoComp>"
                End If
                If Len(grd.TextMatrix(i, COL_C_NUMAUTOSRI)) <> 10 Or grd.ValueMatrix(i, COL_C_NUMAUTOSRI) < 1 Then
                            msg = " El Numero de Autorización SRI " & grd.TextMatrix(i, COL_C_NUMAUTOSRI) & " Incorrecto"
                            'MsgBox msg
                            grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                            grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                            grd.ShowCell i, grd.ColIndex("Resultado")
                            GeneraArchivoREOCComprasXML = False
                            lblPasos(5).Caption = "Error"
                            GoTo SiguienteFila
                Else
                    cadenaCP = cadenaCP & "<aut>" & grd.TextMatrix(i, COL_C_NUMAUTOSRI) & "</aut>"
                End If
                
                
                If Len(grd.TextMatrix(i, COL_C_NUMSERESTA)) <> 3 Or grd.ValueMatrix(i, COL_C_NUMSERESTA) = 0 Then
                            msg = " El Numero de Serie Establecimiento " & grd.TextMatrix(i, COL_C_NUMSERESTA) & " Incorrecto"
                            grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                            grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                            grd.ShowCell i, grd.ColIndex("Resultado")
                            GeneraArchivoREOCComprasXML = False
                            lblPasos(5).Caption = "Error"
                            GoTo SiguienteFila
                Else
                    cadenaCP = cadenaCP & "<estab>" & grd.TextMatrix(i, COL_C_NUMSERESTA) & "</estab>"
                End If
                If Len(grd.TextMatrix(i, COL_C_NUMSERPUNTO)) <> 3 Or grd.ValueMatrix(i, COL_C_NUMSERPUNTO) = 0 Then
                            msg = " El Numero de Serie Punto " & grd.TextMatrix(i, COL_C_NUMSERPUNTO) & " Incorrecto"
                            grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                            grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                            grd.ShowCell i, grd.ColIndex("Resultado")
                            GeneraArchivoREOCComprasXML = False
                            lblPasos(5).Caption = "Error"
                            GoTo SiguienteFila
                Else
                    cadenaCP = cadenaCP & "<ptoEmi>" & grd.TextMatrix(i, COL_C_NUMSERPUNTO) & "</ptoEmi>"
                End If
'                If grd.ValueMatrix(i, COL_C_NUMSECUENCIAL) = 7925 Then MsgBox "hola"
                cadenaCP = cadenaCP & "<sec>" & grd.TextMatrix(i, COL_C_NUMSECUENCIAL) & "</sec>"
                cadenaCP = cadenaCP & "<fechaEmiCom>" & grd.TextMatrix(i, COL_C_FECHATRANS) & "</fechaEmiCom>"
                
                Set rsRet = gobjMain.EmpresaActual.ConsANRetencionCompras2008ParaXML(grd.ValueMatrix(i, COL_C_TRANSID))
                If rsRet.RecordCount = 0 And grd.TextMatrix(i, COL_C_CODTIPOCOMP) <> "4" Then
                    Set rsRet = gobjMain.EmpresaActual.ConsANRetencionCompras2008ParaXMLSinRetencion(grd.ValueMatrix(i, COL_C_TRANSID))
                End If
                
                cadenaCPIR = "<air>"
                 cadenaRET = ""
                If rsRet.RecordCount > 0 Then
                    MiGetRowsRep rsRet, grdRet
                    
                    For j = 1 To grdRet.Rows - 1
                        If (grd.TextMatrix(i, COL_C_TRANS) = grdRet.TextMatrix(j, COL_R_TRANS)) And (grd.TextMatrix(i, COL_C_NUMTRANS) = grdRet.TextMatrix(j, COL_R_NUMTRANS)) And (grd.TextMatrix(i, COL_C_RUC) = grdRet.TextMatrix(j, COL_R_RUC)) Then
                            
                            If grdRet.TextMatrix(j, COL_R_TIPO) <> -1 Then
                                'valores renta
                                cadenaCPIR = cadenaCPIR & "<detalleAir>"
                                cadenaCPIR = cadenaCPIR & "<codRetAir>" & grdRet.TextMatrix(j, COL_R_CODIGOSRI) & "</codRetAir>"
                                cadenaCPIR = cadenaCPIR & "<porcentaje>" & Format(grdRet.ValueMatrix(j, COL_R_PORCEN), "#0.0") & "</porcentaje>"
                                cadenaCPIR = cadenaCPIR & "<base0>" & Format(grd.ValueMatrix(i, COL_C_BASE0), "#0.00") & "</base0>"
                                cadenaCPIR = cadenaCPIR & "<baseGrav>" & Format(grdRet.ValueMatrix(j, COL_R_BASE), "#0.00") & "</baseGrav>"
                                cadenaCPIR = cadenaCPIR & "<baseNoGrav>" & Format(grd.ValueMatrix(i, COL_C_BASENO12), "#0.00") & "</baseNoGrav>"
                                cadenaCPIR = cadenaCPIR & "<valRetAir>" & Format(grdRet.ValueMatrix(j, COL_R_VALOR), "#0.00") & "</valRetAir>"
                                cadenaCPIR = cadenaCPIR & "</detalleAir>"
                                If grdRet.ValueMatrix(j, COL_R_PORCEN) = "0" Then
'                                    cadenaRET = cadenaRET & "<autRet></autRet>"
'                                    cadenaRET = cadenaRET & "<estabRet></estabRet>"
'                                    cadenaRET = cadenaRET & "<ptoEmiRet></ptoEmiRet>"
'                                    cadenaRET = cadenaRET & "<secRet></secRet>"
'                                    cadenaRET = cadenaRET & "<fechaEmiRet></fechaEmiRet>"
                                
                                Else
                                    cadenaRET = cadenaRET & "<autRet>" & grdRet.TextMatrix(j, COL_R_NUMAUTO) & "</autRet>"
                                    cadenaRET = cadenaRET & "<estabRet>" & grdRet.TextMatrix(j, COL_R_NUMEST) & "</estabRet>"
                                    cadenaRET = cadenaRET & "<ptoEmiRet>" & grdRet.TextMatrix(j, COL_R_NUMPTO) & "</ptoEmiRet>"
                                    cadenaRET = cadenaRET & "<secRet>" & grdRet.TextMatrix(j, COL_R_NUMRET) & "</secRet>"
                                    cadenaRET = cadenaRET & "<fechaEmiRet>" & grdRet.TextMatrix(j, COL_R_FECHARET) & "</fechaEmiRet>"
                                End If
                            
                            
                            End If
                        End If
                    
                        If CDate(grdRet.TextMatrix(j, COL_R_FECHARET)) < CDate(grd.TextMatrix(i, COL_C_FECHAREGISTRO)) Then
                            msg = "La fecha de la Retención " & grdRet.TextMatrix(j, COL_R_RETTRANS) & "-" & _
                                        grdRet.TextMatrix(j, COL_R_RETNUMTRANS) & _
                                        " no puede ser menor a la fecha de la Compra " & _
                                        grd.TextMatrix(i, COL_C_TRANS) & "-" & _
                                        grd.TextMatrix(i, COL_C_NUMTRANS)
                            'MsgBox msg
                            grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                            grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                            grd.ShowCell i, grd.ColIndex("Resultado")
                            lblPasos(5).Caption = "Error"
                            GeneraArchivoREOCComprasXML = False
                            GoTo SiguienteFila
                        End If
                        If CDate(grd.TextMatrix(i, COL_C_FECHAREGISTRO)) < CDate(grd.TextMatrix(i, COL_C_FECHATRANS)) Then
                            msg = "La fecha de registro de la Transaccion " & _
                            grd.TextMatrix(i, COL_C_TRANS) & "-" & _
                            grd.TextMatrix(i, COL_C_NUMTRANS) _
                            & " debe ser menor o igual a la fecha de registro "
                            grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                            grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                            grd.ShowCell i, grd.ColIndex("Resultado")
                            GeneraArchivoREOCComprasXML = False
                            lblPasos(5).Caption = "Error"
                            GoTo SiguienteFila
                            
                        End If
                    Next j
                Else
                    For j = grdRet.Rows - 1 To 1 Step -1
                        grdRet.RemoveItem (j)
                    Next j
                End If

                cadenaCP = cadenaCP & cadenaCPIR & "</air>" & cadenaRET
                cadenaCP = cadenaCP & "</detalleCompras>"
                grd.ShowCell i, grd.ColIndex("Resultado")
                grd.TextMatrix(i, grd.ColIndex("Resultado")) = " OK "
                grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbWhite
            GoTo SiguienteFila
    Exit Function
SiguienteFila:
    Next i
    grd.ColWidth(grd.ColIndex("Resultado")) = 5000
    prg.value = 0
    If Len(lblPasos(5).Caption) = 0 Then
        lblPasos(5).Caption = "OK."
        cadenaCP = cadenaCP & "</compras>"
    Else
        cadenaCP = ""
    End If
    cad = cadenaCP
Exit Function
cancelado:
    GeneraArchivoREOCComprasXML = False
ErrTrap:
    grd.TextMatrix(grd.Rows - 1, 2) = Err.Description
    GeneraArchivoREOCComprasXML = False
End Function

Public Sub Exportar(tag As String)
    Dim file As String, NumFile As Integer, Cadena As String
    Dim Filas As Long, Columnas As Long, i As Long, j As Long
    Dim pos As Integer
'    If grd.Rows = grd.FixedRows Then Exit Sub
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
        For j = 2 To grd.Cols - 1
            Select Case tag          ' jeaa 04/11/03 para que se no se guarden las columnas ocultas
                Case "IMPCP"
                        If j = COL_C_NOMBRE Then j = j + 1  'columna nombre
            End Select
                If pos = 0 Then
                    Cadena = Cadena & grd.TextMatrix(i, j) & ","
                Else
                    Cadena = Cadena & Mid$(grd.TextMatrix(i, j), 1, pos - 1) & Mid$(grd.TextMatrix(i, j), pos + 1, Len(grd.TextMatrix(i, j)) - 1) & ","
                End If


        Next j
        Cadena = Mid(Cadena, 1, Len(Cadena) - 1)
        Print #NumFile, Cadena
        Cadena = ""
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

Private Sub AbrirArchivo(cad As String)
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
        
        ConfigCols cad
        .ShowOpen
        
        Select Case UCase$(Right$(dlg1.filename, 4))
        Case ".TXT", ".CSV"
            VisualizarTexto dlg1.filename, cad
        Case Else
        End Select
    End With
    Exit Sub
ErrTrap:
    If Err.Number <> 32755 Then DispErr
    Exit Sub
End Sub

Private Sub VisualizarTexto(ByVal archi As String, cad As String)
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
            Select Case cad          ' jeaa 31-10-03 para aumentar las columnas ocultas
                Case "IMPCPI"
                    s = vbTab & s    '1,2' ' jeaa 31-10-03 para aumentar las columnas TRANSID
            
            End Select
           grd.AddItem s
        Loop
    Close #f
    RemueveSpace
    Select Case cad
        Case "IMPCPI"
            grd.Select 1, 1, 1, 1
    End Select
    grd.Sort = flexSortUseColSort
' poner numero
    GNPoneNumFila grd, False
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

Private Sub grd_KeyDown(KeyCode As Integer, Shift As Integer)
    If grd.IsSubtotal(grd.Row) Then Exit Sub
    Select Case KeyCode
    Case vbKeyInsert
        AgregarFila
    Case vbKeyDelete
        EliminarFila
    End Select
End Sub

Private Sub AgregarFila()
    On Error GoTo ErrTrap
    With grd
        .AddItem "", .Row + 1
        GNPoneNumFila grd, False
        .Row = .Row + 1
        .col = .FixedCols
    End With
    
    AjustarAutoSize grd, -1, -1
    grd.SetFocus
    Exit Sub
ErrTrap:
    MsgBox Err.Description
    grd.SetFocus
    Exit Sub
End Sub

Private Sub EliminarFila()
    On Error GoTo ErrTrap
    If grd.Row <> grd.FixedRows - 1 And Not grd.IsSubtotal(grd.Row) Then
        grd.RemoveItem grd.Row
        GNPoneNumFila grd, False
    End If
    grd.SetFocus
    Exit Sub
ErrTrap:
    MsgBox Err.Description
    grd.SetFocus
    Exit Sub
End Sub

Private Sub SubTotalizar(col As Long)
    Dim i As Long
    With grd
        For i = 1 To .Cols - 1
            If i = COL_C_CODTIPOCOMP Then i = i + 1
            If grd.ColData(i) = "SubTotal" Then
                    .subtotal flexSTSum, col, i, , grd.GridColor, vbBlack, , "Subtotal", col, True
            End If
        Next i
        .subtotal flexSTCount, col, col, , grd.GridColor, vbBlack, , "Subtotal", col, True

    End With
''    col = 2
''    With grdTalonCP
''        For i = 1 To .Cols - 1
''            'If i = 1 Then i = i + 1
''            If .ColData(i) = "SubTotal" Then
''                    .subtotal flexSTSum, col, i, , .GridColor, vbBlack, , "Subtotal", col, True
''            End If
''        Next i
''        .subtotal flexSTCount, col, col, , .GridColor, vbBlack, , "Subtotal", col, True
''
''    End With
    
End Sub

Private Sub Totalizar()
    Dim i As Long
    With grd
        For i = 1 To .Cols - 1
            If i = COL_C_CODTIPOCOMP Then i = i + 1
            If grd.ColData(i) = "SubTotal" Then
                
                .subtotal flexSTSum, -1, i, "#,#0.00", .BackColorSel, vbYellow, vbBlack, "Total"
            End If
        Next i
'        .subtotal flexSTCount, -1, COL_C_CODTIPOCOMP, "#,#0", .BackColorSel, vbYellow, vbBlack, "Total"
    End With
    
    With grdTalonCP
        For i = 1 To .Cols - 1
            'If i = COL_C_CODTIPOCOMP Then i = i + 1
            If grdTalonCP.ColData(i) = "SubTotal" Then

                .subtotal flexSTSum, -1, i, "#,#0.00", .BackColorSel, vbYellow, vbBlack, "Total"
            End If
        Next i


    End With
    
    With grdTalonRTPCP
        For i = 1 To .Cols - 1
            'If i = COL_C_CODTIPOCOMP Then i = i + 1
            If grdTalonRTPCP.ColData(i) = "SubTotal" Then

                .subtotal flexSTSum, -1, i, "#,#0.00", .BackColorSel, vbYellow, vbBlack, "Total"
            End If
        Next i


    End With
    
    With grdTalonRTPIVACP
        For i = 1 To .Cols - 1
            'If i = COL_C_CODTIPOCOMP Then i = i + 1
            If grdTalonRTPIVACP.ColData(i) = "SubTotal" Then

                .subtotal flexSTSum, -1, i, "#,#0.00", .BackColorSel, vbYellow, vbBlack, "Total"
            End If
        Next i


    End With
    
    
    With grdTalonFC
        For i = 1 To .Cols - 1
            'If i = COL_C_CODTIPOCOMP Then i = i + 1
            If grdTalonFC.ColData(i) = "SubTotal" Then

                .subtotal flexSTSum, -1, i, "#,#0.00", .BackColorSel, vbYellow, vbBlack, "Total"
            End If
        Next i


    End With
    
    
End Sub


Private Sub SubTotalizarCF(col As Long)
    Dim i As Long
    col = 6
    With grdCF
        For i = 5 To .Cols - 1
'            If i = COL_V_CANTRANS Then i = i + 1
'            If grdCF.ColData(i) = "SubTotal" Then
                If i = COL_V_TIPOCOMP Then
                    .subtotal flexSTMax, 6, i, , grdCF.GridColor, vbBlack, , "Subtotal", col, True
                Else
                    .subtotal flexSTSum, 6, i, , grdCF.GridColor, vbBlack, , "Subtotal", col, True
                End If
'            End If
        Next i
        
        '.subtotal flexSTCount, col, col, , grdCF.GridColor, vbBlack, , "Subtotal", col, True
'        .subtotal flexSTMax, 6, 6, , grdCF.GridColor, vbBlack, , "Subtotal", 5, True
        '.subtotal flexSTCount, col, col, , grdCF.GridColor, vbBlack, , "Subtotal", col, True

    End With
End Sub

Private Sub TotalizarCF()
    Dim i As Long
    With grdCF
        For i = COL_V_TIPOCOMP To .Cols - 1
'            If i = COL_V_CANTRANS Then i = i + 1
'            If grdCF.ColData(i) = "SubTotal" Then
                If i = COL_V_TIPOCOMP Then

                    .subtotal flexSTMax, -1, i, "#,#0.00", .BackColorSel, vbYellow, vbBlack, "Total"
                Else
                    .subtotal flexSTSum, -1, i, "#,#0.00", .BackColorSel, vbYellow, vbBlack, "Total"
                End If
 '           End If
        Next i
'        .subtotal flexSTMax, COL_V_TIPOCOMP, COL_V_TIPOCOMP, , .BackColorSel, vbYellow, vbBlack, "Total"

    End With
End Sub


Private Function GeneraArchivoATSVentasXMLCF() As String
    Dim cadenaFC As String, cadenaFCIVA  As String
    Dim i As Long, j As Long
    Dim vIR As Variant, cadenaFCIR As String
    Dim FilasIR As Long, ColumnasIR As Long, iIR As Long, jIR As Long
    Dim rsRet As Recordset, cadenaFCIVA30 As String
    Dim cadenaFCIVA70 As String, cadenaFCIVA100 As String
    Dim rsNC As Recordset, cadenaNC As String
    Dim msg As String, pc As PCProvCli, bandCF As Boolean, filaCF As Integer
    Dim BandFact As Boolean
    
    On Error GoTo ErrTrap
    GeneraArchivoATSVentasXMLCF = ""
            BandFact = False
            For i = 1 To grdCF.Rows - 2
               If grdCF.IsSubtotal(i) Then
                    cadenaFC = cadenaFC & "<detalleVentas>"
                    cadenaFC = cadenaFC & "<tpIdCliente>07</tpIdCliente>"
                    cadenaFC = cadenaFC & "<idCliente>9999999999999</idCliente>"
                    If BandFact = False Then
                        cadenaFC = cadenaFC & "<tipoComprobante>18</tipoComprobante>"
                        BandFact = True
                    Else
                        cadenaFC = cadenaFC & "<tipoComprobante>04</tipoComprobante>"
                    End If
                    cadenaFC = cadenaFC & "<numeroComprobantes>" & Format(grdCF.TextMatrix(i, COL_V_CANTRANS), "#0") & "</numeroComprobantes>"
                    cadenaFC = cadenaFC & "<baseNoGraIva>" & Format(Abs(grdCF.ValueMatrix(i, COL_V_BASENOIVA)), "#0.00") & "</baseNoGraIva>"
                    cadenaFC = cadenaFC & "<baseImponible>" & Format(Abs(grdCF.ValueMatrix(i, COL_V_BASE0)), "#0.00") & "</baseImponible>"
                    cadenaFC = cadenaFC & "<baseImpGrav>" & Format(Abs(grdCF.ValueMatrix(i, COL_V_BASEIVA)), "#0.00") & "</baseImpGrav>"
                    cadenaFC = cadenaFC & "<montoIva>" & Format(IIf(Abs(grdCF.ValueMatrix(i, COL_V_BASEIVA)) = 0, "0.00", Abs(grdCF.ValueMatrix(i, COL_V_BASEIVA)) * (grd.ValueMatrix(i, COL_V_IVA))), "#0.00") & "</montoIva>"
                    cadenaFCIVA = "<valorRetIva> 0.00 </valorRetIva>"
                    cadenaFCIR = "<valorRetRenta> 0.00 </valorRetRenta>"
                    
                     cadenaFC = cadenaFC & cadenaFCIVA
                    cadenaFC = cadenaFC & cadenaFCIR
                    cadenaFC = cadenaFC & "<formasDePago> <formaPago>01</formaPago> </formasDePago>"
                    cadenaFC = cadenaFC & "</detalleVentas>"
                    
                End If
                grdCF.ShowCell i, grdCF.ColIndex("Resultado")
                grdCF.TextMatrix(i, grdCF.ColIndex("Resultado")) = " OK "
        
        Next i
    
    
    grdCF.ColWidth(grd.ColIndex("Resultado")) = 5000
    prg.value = 0
    GeneraArchivoATSVentasXMLCF = cadenaFC
    Exit Function
cancelado:
    GeneraArchivoATSVentasXMLCF = ""
ErrTrap:
    grdCF.TextMatrix(grd.Rows - 1, 2) = Err.Description
    GeneraArchivoATSVentasXMLCF = ""
End Function

Private Function GeneraArchivoATSVentasXMLSoloRetencion() As String
    Dim cadenaFC As String, cadenaFCIVA  As String
    Dim i As Long, j As Long
    Dim vIR As Variant, cadenaFCIR As String
    Dim FilasIR As Long, ColumnasIR As Long, iIR As Long, jIR As Long
    Dim rsRet As Recordset, cadenaFCIVA30 As String
    Dim cadenaFCIVA70 As String, cadenaFCIVA100 As String
    Dim rsNC As Recordset, cadenaNC As String
    Dim msg As String, pc As PCProvCli, bandCF As Boolean, filaCF As Integer
    Dim BandFact As Boolean
    
    On Error GoTo ErrTrap
        GeneraArchivoATSVentasXMLSoloRetencion = ""
            BandFact = False
            For i = 1 To GrdRetVentas.Rows - 1
               If GrdRetVentas.TextMatrix(i, 8) <> "OK" Then
                    cadenaFC = cadenaFC & "<detalleVentas>"
                Select Case GrdRetVentas.TextMatrix(i, COL_V_TIPODOC)
                    Case "R":                     cadenaFC = cadenaFC & "<tpIdCliente>" & "04" & "</tpIdCliente>"
                    Case "C":                     cadenaFC = cadenaFC & "<tpIdCliente>" & "05" & "</tpIdCliente>"
                    Case "P":                     cadenaFC = cadenaFC & "<tpIdCliente>" & "06" & "</tpIdCliente>"
                    Case "F":                     cadenaFC = cadenaFC & "<tpIdCliente>" & "07" & "</tpIdCliente>"
                End Select
                    
                  
                    cadenaFC = cadenaFC & "<idCliente>" & GrdRetVentas.TextMatrix(i, COL_V_RUC) & "</idCliente>"
                    
                    
                    
                    Set pc = gobjMain.EmpresaActual.RecuperaPCProvClixRUC(GrdRetVentas.TextMatrix(i, COL_V_RUC), True, False, False)
                    If Not pc Is Nothing Then
                        Select Case pc.codtipoDocumento
                            Case "R", "C", "P"
                                If Not pc.BandRelacionado Then
                                    cadenaFC = cadenaFC & "<parteRelVtas>" & "NO" & "</parteRelVtas>"   'auc
                                Else
                                    cadenaFC = cadenaFC & "<parteRelVtas>" & "SI" & "</parteRelVtas>"   'auc
                                End If
                            Case "F"
                            Case Else
                                cadenaFC = cadenaFC & "<parteRelVtas>" & "NO" & "</parteRelVtas>"
                        End Select
                    Else
                        If GrdRetVentas.TextMatrix(i, COL_V_RUC) <> "9999999999999" Then
                            cadenaFC = cadenaFC & "<parteRelVtas>" & "NO" & "</parteRelVtas>"   'auc
                        End If
                    End If
                    Set pc = Nothing
                    
                    
                    cadenaFC = cadenaFC & "<tipoComprobante>18</tipoComprobante>"
                    cadenaFC = cadenaFC & "<tipoEmision>F</tipoEmision>"
                    cadenaFC = cadenaFC & "<numeroComprobantes>0</numeroComprobantes>"
                    cadenaFC = cadenaFC & "<baseNoGraIva>0.00</baseNoGraIva>"
                    cadenaFC = cadenaFC & "<baseImponible>0.00</baseImponible>"
                    cadenaFC = cadenaFC & "<baseImpGrav>0.00</baseImpGrav>"
                    cadenaFC = cadenaFC & "<montoIva>0.00</montoIva>"
                    cadenaFC = cadenaFC & "<montoIce>0.00</montoIce>"
                    cadenaFCIVA = "<valorRetIva> 0.00 </valorRetIva>"
                    cadenaFCIR = "<valorRetRenta> 0.00 </valorRetRenta>"
                    
                    If GrdRetVentas.TextMatrix(i, 5) = -1 Then
                        'valores iva
                        cadenaFCIVA = "<valorRetIva>" & Format(GrdRetVentas.ValueMatrix(i, 7), "#0.00") & "</valorRetIva>"
                        If i + 1 < GrdRetVentas.Rows - 1 Then
                            If GrdRetVentas.TextMatrix(i, COL_V_RUC) = GrdRetVentas.TextMatrix(i + 1, COL_V_RUC) And GrdRetVentas.TextMatrix(i + 1, 5) = 0 Then
                                cadenaFCIR = "<valorRetRenta>" & Format(GrdRetVentas.ValueMatrix(i + 1, 7), "#0.00") & "</valorRetRenta>"
                                GrdRetVentas.TextMatrix(i + 1, 8) = "OK"
                            End If
                        End If
                        
                    Else
                        'valores renta
                        cadenaFCIR = "<valorRetRenta>" & Format(GrdRetVentas.ValueMatrix(i, 7), "#0.00") & "</valorRetRenta>"
                        If i + 1 <= GrdRetVentas.Rows - 1 Then
                            If GrdRetVentas.TextMatrix(i, COL_V_RUC) = GrdRetVentas.TextMatrix(i + 1, COL_V_RUC) And GrdRetVentas.TextMatrix(i + 1, 5) = -1 Then
                                cadenaFCIVA = "<valorRetIva>" & Format(GrdRetVentas.ValueMatrix(i + 1, 7), "#0.00") & "</valorRetIva>"
                                GrdRetVentas.TextMatrix(i + 1, 8) = "OK"
                            End If
                        End If
                        
                        
                        
                    End If
                    
                    'cadenaFC = "<formasDePago> <formaPago>01</formaPago> </formasDePago>"
                    cadenaFC = cadenaFC & cadenaFCIVA
                    cadenaFC = cadenaFC & cadenaFCIR
                    cadenaFC = cadenaFC & "<formasDePago> <formaPago>01</formaPago> </formasDePago>"
                    cadenaFC = cadenaFC & "</detalleVentas>"
                    
                End If
                GrdRetVentas.ShowCell i, 8
                GrdRetVentas.TextMatrix(i, 8) = " OK "
        
        Next i
    
    
    GrdRetVentas.ColWidth(8) = 5000
    prg.value = 0
    GeneraArchivoATSVentasXMLSoloRetencion = cadenaFC
    Exit Function
cancelado:
    GeneraArchivoATSVentasXMLSoloRetencion = ""
ErrTrap:
    grdCF.TextMatrix(grd.Rows - 1, 2) = Err.Description
    GeneraArchivoATSVentasXMLSoloRetencion = ""
End Function

Private Function BuscarVentasEstablecimientoATS()
    On Error GoTo ErrTrap
        With grd
        .Redraw = False
        .Rows = .FixedRows
        If Not frmB_Trans.Inicio(gobjMain, "IMPFCxE", dtpPeriodo.value) Then
            grd.SetFocus
        End If
        mObjCond.fecha1 = gobjMain.objCondicion.fecha1
        mObjCond.fecha2 = gobjMain.objCondicion.fecha2
        'MiGetRowsRep gobjMain.EmpresaActual.ConsANVentasxEstablecimiento2013paraXML(), grd
        MiGetRowsRep gobjMain.EmpresaActual.Empresa2.ConsANVentasxEstablecimiento2016paraXML(), grd
        

        'GeneraArchivo

        ConfigCols "IMPFCxE"

        AjustarAutoSize grd, -1, -1
        grd.ColWidth(0) = "500"

        SubTotalizar (COL_VE_SUC)
        Totalizar

        GNPoneNumFila grd, False


        .Redraw = True
   End With

    Exit Function
ErrTrap:
    grd.Redraw = True
    DispErr
    Exit Function
End Function


Private Function GenerarVentasEstablecimientoATS(ByRef cad As String) As Boolean
    On Error GoTo ErrTrap
        GenerarVentasEstablecimientoATS = False
        GenerarVentasEstablecimientoATS = GeneraArchivoATSVentasEstablecimientoXML(cad)
    Exit Function
ErrTrap:
    grd.Redraw = True
    DispErr
    Exit Function
End Function

Private Function GeneraArchivoATSVentasEstablecimientoXML(ByRef cad As String) As Boolean
    Dim cadenaFC As String, cadenaFCIVA  As String
    Dim i As Long, j As Long
    Dim vIR As Variant, cadenaFCIR As String
    Dim FilasIR As Long, ColumnasIR As Long, iIR As Long, jIR As Long
    Dim rsRet As Recordset, cadenaFCIVA30 As String
    Dim cadenaFCIVA70 As String, cadenaFCIVA100 As String
    Dim rsNC As Recordset, cadenaNC As String
    Dim msg As String, pc As PCProvCli, bandCF As Boolean, filaCF As Integer
    Dim cadenaF As String, k As Integer
    
    On Error GoTo ErrTrap
    GeneraArchivoATSVentasEstablecimientoXML = True
    bandCF = False
    filaCF = 1
   
    
        grd.Refresh
        cadenaF = "<ventasEstablecimiento>"

            If grd.Rows = 1 Then
                prg.value = 0
'                cadenaF = cadenaF & "</ventasEstablecimiento>"
                cad = cadenaF
                GeneraArchivoATSVentasEstablecimientoXML = True
                GoTo SiguienteFila
            Else
            End If


            prg.max = grd.Rows - 1
            For i = 1 To grd.Rows - 2
                If grd.IsSubtotal(i) Then 'GoTo SiguienteFila
                grd.ShowCell i, 1
                prg.value = i
                DoEvents
                cadenaFC = ""
                cadenaFC = cadenaFC & "<ventaEst>"
                cadenaFC = cadenaFC & "<codEstab>" & grd.TextMatrix(i - 1, COL_VE_SUC) & "</codEstab>"
                cadenaFC = cadenaFC & "<ventasEstab>" & Format(grd.ValueMatrix(i, COL_VE_TOTAL) + grd.ValueMatrix(i, COL_VE_ICE), "#0.00") & "</ventasEstab>"
                cadenaFC = cadenaFC & "</ventaEst>"
                cadenaF = cadenaF & cadenaFC
                grd.ShowCell i, grd.ColIndex("Resultado")
                grd.TextMatrix(i, grd.ColIndex("Resultado")) = " OK "
                End If
            
        
SiguienteFila:
    Next i
    

    
        grd.ColWidth(grd.ColIndex("Resultado")) = 5000
        prg.value = 0
        
        
        lblResp(2).Caption = "OK."
    
    cad = cadenaF & "</ventasEstablecimiento>"

    Exit Function
cancelado:
    GeneraArchivoATSVentasEstablecimientoXML = False
ErrTrap:
    grd.TextMatrix(grd.Rows - 1, 2) = Err.Description
    GeneraArchivoATSVentasEstablecimientoXML = False
End Function

Private Sub TotalizarVentaestablecimiento()
    Dim i As Long
    With grd
        For i = 1 To .Cols - 1
            If grd.ColData(i) = "SubTotal" Then
                .subtotal flexSTSum, -1, i, "#,#0.00", .BackColorSel, vbYellow, vbBlack, "Total"
            End If
        Next i
    End With
End Sub


Private Function BuscarExportacionesATS()
    On Error GoTo ErrTrap
        With grd
        .Redraw = False
        .Rows = .FixedRows
        If Not frmB_Trans.Inicio(gobjMain, "IMPFCEXP", dtpPeriodo.value) Then
            grd.SetFocus
        End If
        mObjCond.fecha1 = gobjMain.objCondicion.fecha1
        mObjCond.fecha2 = gobjMain.objCondicion.fecha2
        MiGetRowsRep gobjMain.EmpresaActual.ConsANExportacion2015paraXML(), grd

        'GeneraArchivo

        ConfigCols "IMPEX"

        AjustarAutoSize grd, -1, -1
        AjustarAutoSize grdRet, -1, -1
        grd.ColWidth(0) = "500"
        grd.ColHidden(COL_V_VALORIVA) = True
        SubTotalizar (COL_V_TIPOCOMP)
        Totalizar

        GNPoneNumFila grd, False
        GNPoneNumFila grdRet, False

        .Redraw = True
   End With

    Exit Function
ErrTrap:
    grd.Redraw = True
    DispErr
    Exit Function
End Function

Private Function GenerarExportacionATS(ByRef cad As String) As Boolean
    
    On Error GoTo ErrTrap
        GenerarExportacionATS = False
        GenerarExportacionATS = GeneraArchivoATSExportacionXML(cad)
    Exit Function
ErrTrap:
    grd.Redraw = True
    DispErr
    Exit Function
End Function

Private Function GeneraArchivoATSExportacionXML(ByRef cad As String) As Boolean
    Dim cadenaEX As String
    Dim i As Long, j As Long, resp As Integer
    Dim pc As PCProvCli
    Dim msg
    On Error GoTo ErrTrap
    resp = 10
    GeneraArchivoATSExportacionXML = True
    grd.Refresh
    'With grd
        
        cadenaEX = "<exportaciones>"
            If grd.Rows < 1 Then
                prg.value = 0
                cadenaEX = cadenaEX & "</exportaciones>"
                cad = cadenaEX
                GeneraArchivoATSExportacionXML = True
                GoTo SiguienteFila
            End If
            prg.max = grd.Rows - 1
            For i = 1 To grd.Rows - 1
                If grd.IsSubtotal(i) Then GoTo SiguienteFila
                grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbWhite
                prg.value = i
                DoEvents
                Set pc = gobjMain.EmpresaActual.RecuperaPCProvClixRUC(grd.TextMatrix(i, 1), True, False, False)
                
                cadenaEX = cadenaEX & "<detalleExportaciones>"
                If pc.codtipoDocumento = "R" Then
                    cadenaEX = cadenaEX & "<tpIdClienteEx>" & "20" & "</tpIdClienteEx>"
                ElseIf pc.codtipoDocumento = "P" Then
                    cadenaEX = cadenaEX & "<tpIdClienteEx>" & "21" & "</tpIdClienteEx>"
               Else
                                msg = " El Tipo de documento del Cliente " & grd.TextMatrix(i, COL_E_NOMBRE) & " es Incorrecto"
                                'MsgBox msg
                                grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                                grd.ShowCell i, grd.ColIndex("Resultado")
                                grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                                GeneraArchivoATSExportacionXML = False
                                lblResp(0).Caption = "Error"
                                GoTo SiguienteFila
                    
                End If
                cadenaEX = cadenaEX & "<idClienteEx>" & pc.ruc & "</idClienteEx>"
                cadenaEX = cadenaEX & "<parteRelExp>" & IIf(pc.BandRelacionado, "SI", "NO") & "</parteRelExp>"
                If pc.TipoProvCli = "RPN" Then
                    cadenaEX = cadenaEX & "<tipoCli>01</tipoCli>"
                ElseIf pc.TipoProvCli = "RSO" Then
                    cadenaEX = cadenaEX & "<tipoCli>02</tipoCli>"
                End If
                
                
                cadenaEX = cadenaEX & "<denoExpCli>" & grd.TextMatrix(i, COL_E_NOMBRE) & "</denoExpCli>"
                
                
                cadenaEX = cadenaEX & "<tipoRegi>01</tipoRegi>"
                cadenaEX = cadenaEX & "<paisEfecPagoGen>" & grd.TextMatrix(i, COL_E_DESTINO) & "</paisEfecPagoGen>"
'                    cadenaEX = cadenaEX & "<parteRel>" & IIf(pc.BandRelacionado, "SI", "NO") & "</parteRel>"
                
                cadenaEX = cadenaEX & "<paisEfecExp>" & grd.TextMatrix(i, COL_E_DESTINO) & "</paisEfecExp>"
                
                
                
'                cadenaEX = cadenaEX & "<pagoRegFis>" & "NO" & " </pagoRegFis>"
                cadenaEX = cadenaEX & "<exportacionDe>" & grd.TextMatrix(i, COL_E_REFERENDO) & "</exportacionDe>"
                cadenaEX = cadenaEX & "<tipoComprobante>" & grd.TextMatrix(i, COL_E_TIPOCOMPROBANTE) & "</tipoComprobante>"
                cadenaEX = cadenaEX & "<distAduanero>" & grd.TextMatrix(i, COL_E_DISTRITO) & "</distAduanero>"
                cadenaEX = cadenaEX & "<anio>" & grd.TextMatrix(i, COL_E_ANIO) & "</anio>"
                cadenaEX = cadenaEX & "<regimen>" & grd.TextMatrix(i, COL_E_REGIMEN) & "</regimen>"
                cadenaEX = cadenaEX & "<correlativo>" & grd.TextMatrix(i, COL_E_CORRELATIVO) & "</correlativo>"
                cadenaEX = cadenaEX & "<docTransp>" & grd.TextMatrix(i, COL_E_DOCTRANSPORTE) & "</docTransp>"
                cadenaEX = cadenaEX & "<fechaEmbarque>" & grd.TextMatrix(i, COL_E_FECHAEMBARQUE) & "</fechaEmbarque>"
                cadenaEX = cadenaEX & "<valorFOB>" & Format(Abs(grd.ValueMatrix(i, COL_E_VALORFOB)), "#0.00") & "</valorFOB>"
                cadenaEX = cadenaEX & "<valorFOBComprobante>" & Format(Abs(grd.ValueMatrix(i, COL_E_VALORFOBLOCAL)), "#0.00") & "</valorFOBComprobante>"
                cadenaEX = cadenaEX & "<establecimiento>" & grd.TextMatrix(i, COL_E_NUMSERESTA) & "</establecimiento>"
                cadenaEX = cadenaEX & "<puntoEmision>" & grd.TextMatrix(i, COL_E_NUMSERPUNTO) & "</puntoEmision>"
                cadenaEX = cadenaEX & "<secuencial>" & grd.TextMatrix(i, COL_E_NUMSECUENCIAL) & "</secuencial>"
                cadenaEX = cadenaEX & "<autorizacion>" & grd.TextMatrix(i, COL_E_NUMAUTOSRI) & "</autorizacion>"
                cadenaEX = cadenaEX & "<fechaEmision>" & grd.TextMatrix(i, COL_E_FECHATRANS) & "</fechaEmision>"
                
                cadenaEX = cadenaEX & "</detalleExportaciones>"
                grd.ShowCell i, grd.ColIndex("Resultado")
                grd.TextMatrix(i, grd.ColIndex("Resultado")) = " OK "
                grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbWhite
                Set pc = Nothing
            GoTo SiguienteFila
    Exit Function
SiguienteFila:
    Next i
    grd.ColWidth(grd.ColIndex("Resultado")) = 5000
    prg.value = 0
    If Len(lblResp(3).Caption) = 0 Then
        lblResp(3).Caption = "OK."
        cadenaEX = cadenaEX & "</exportaciones>"
    Else
        cadenaEX = ""
    End If
    cad = cadenaEX
    
Exit Function
cancelado:
    GeneraArchivoATSExportacionXML = False
ErrTrap:
    grd.TextMatrix(grd.Rows - 1, 2) = Err.Description
    GeneraArchivoATSExportacionXML = False
End Function

 Public Sub Inicio2016(ByVal tag As String)
    On Error GoTo ErrTrap
    Set mObjCond = New RepCondicion
    Select Case tag
        Case "FAT2016"
            Me.Caption = "Anexo Transaccional 2016"
    End Select
    TotalVentas = 0
    dtpPeriodo.value = CDate("01/" & IIf(Month(Date) - 1 <> 0, Month(Date) - 1, 12) & "/" & Year(Date))
    mObjCond.fecha1 = dtpPeriodo.value
    cboTipo.ListIndex = 0
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("RutaATS-REOC")) > 0 Then
        txtCarpeta.Text = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("RutaATS-REOC")
    End If
    Me.tag = tag
    Me.Show
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub cmdexcel_Click()
     ExpExcel
End Sub

Private Sub ExpExcel()
Dim ex As Excel.Application, ws As Worksheet, wkb As Workbook
Dim fila As Long, NumCol As Long, col As Integer
Dim colu() As Long, v() As Long, mayor As Long
    
    MensajeStatus MSG_PREPARA, vbHourglass


    fila = 4
    NumCol = 0
       
    Dim i   As Integer
    Dim j   As Integer
       
       
    Set ex = New Excel.Application  'Crea un instancia nueva de excel
    Set wkb = ex.Workbooks.Add  'Insertar un libro nuevo
    Set ws = ex.Worksheets.Add  'Inserta una nueva hoja
    With ws
        .Name = Left(Me.Caption, 25)
        .Range("A1").Font.Name = "Times Roman"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Cells(1) = gobjMain.EmpresaActual.GNOpcion.NombreEmpresa
    End With
       
       ws.Cells(2, 1) = Me.Caption
       
        For i = 1 To grd.Cols - 1
           If grd.ColHidden(i) = False Then
               ReDim Preserve colu(NumCol) 'para saber la posicion de la columan en la grilla
                    colu(NumCol) = i 'guarda la posicion
                    NumCol = NumCol + 1 'Para saber el número de columnas se exportan a Excel
                    ws.Cells(fila, NumCol) = grd.TextMatrix(0, i)
                ReDim Preserve v(NumCol)
                v(NumCol - 1) = 0 'encera el vector
           End If
        Next i
       
       ws.Range(ws.Cells(fila, 1), ws.Cells(fila, NumCol)).Font.Bold = True
       ws.Range(ws.Cells(fila, 1), ws.Cells(fila, NumCol)).Borders.LineStyle = 1
       
        For i = grd.FixedRows To grd.Rows - 1
            If Not grd.RowHidden(i) Then
                fila = fila + 1
                j = 1
                mayor = 0
                 For col = 1 To grd.Cols - 1
                    
                     If Not grd.ColHidden(col) And Not grd.RowHidden(i) Then
                        If InStr(1, grd.TextMatrix(i, col), "/") > 0 Then
                             ws.Cells(fila, j) = "'" & grd.TextMatrix(i, col)
                        ElseIf grd.TextMatrix(0, col) = "RUC" Or grd.TextMatrix(0, col) = "Estab" Or grd.TextMatrix(0, col) = "Punto" Or grd.TextMatrix(0, col) = "Secuencial" Or grd.TextMatrix(0, col) = "Sustento" Or grd.TextMatrix(0, col) = "Tipo Pago" Or grd.TextMatrix(0, col) = "AutSRI" Or grd.TextMatrix(0, col) = "Forma1" Or grd.TextMatrix(0, col) = "Forma2" Then
                             ws.Cells(fila, j) = "'" & grd.TextMatrix(i, col)
                        ElseIf InStr(1, grd.TextMatrix(i, col), ".") = 0 And Mid$(grd.TextMatrix(i, col), 1, 1) = "0" And Len(grd.TextMatrix(i, col)) > 4 Then
                            ws.Cells(fila, j) = "'" & grd.TextMatrix(i, col)
                        Else
                            ws.Cells(fila, j) = grd.TextMatrix(i, col)
                        End If
                        j = j + 1
                     End If

               
                 Next col
            End If
            ws.Range(ws.Cells(fila, 1), ws.Cells(fila, NumCol)).Borders.LineStyle = 1
       Next i
       
    ex.Visible = True
    ws.Activate
    Set ws = Nothing
    Set wkb = Nothing
    Set ex = Nothing
    'ex.Quit
    MensajeStatus
End Sub


