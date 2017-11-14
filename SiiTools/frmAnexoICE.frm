VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{ED5A9B02-5BDB-48C7-BAB1-642DCC8C9E4D}#2.0#0"; "SelFold.ocx"
Begin VB.Form frmAnexoICE 
   Caption         =   "Anexo ICE"
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
      Left            =   120
      TabIndex        =   9
      Top             =   5100
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
      FormatString    =   $"frmAnexoICE.frx":0000
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
      Height          =   1095
      Left            =   60
      TabIndex        =   17
      Top             =   60
      Width           =   5595
      Begin VB.TextBox txtCarpeta 
         Height          =   320
         Left            =   840
         TabIndex        =   3
         Text            =   "c:\"
         Top             =   600
         Width           =   4170
      End
      Begin VB.CommandButton cmdExaminarCarpeta 
         Caption         =   "..."
         Height          =   320
         Index           =   0
         Left            =   4980
         TabIndex        =   4
         Top             =   600
         Width           =   372
      End
      Begin SelFold.SelFolder slf 
         Left            =   4140
         Top             =   420
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
         Top             =   240
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
         Format          =   2031619
         CurrentDate     =   37356
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         ItemData        =   "frmAnexoICE.frx":0063
         Left            =   840
         List            =   "frmAnexoICE.frx":006D
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.Label Label2 
         Caption         =   "Mes:"
         Height          =   255
         Left            =   60
         TabIndex        =   19
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   60
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label Label1 
         Caption         =   "Ubicacion:"
         Height          =   255
         Left            =   0
         TabIndex        =   18
         Top             =   660
         Width           =   870
      End
   End
   Begin VB.Frame fraPasos 
      Height          =   1095
      Left            =   5760
      TabIndex        =   12
      Top             =   60
      Width           =   8115
      Begin VB.CheckBox chkConsFinal 
         Caption         =   "Cargar como Consumidor Final"
         Height          =   195
         Left            =   5340
         TabIndex        =   40
         Top             =   660
         Width           =   2535
      End
      Begin VB.CheckBox chkSoloError 
         Caption         =   "Solo con error"
         Height          =   195
         Left            =   5340
         TabIndex        =   35
         Top             =   300
         Visible         =   0   'False
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
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "Generar Archivo"
         Height          =   330
         Index           =   10
         Left            =   2940
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   600
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
         Visible         =   0   'False
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
         Visible         =   0   'False
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
         Visible         =   0   'False
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
         Visible         =   0   'False
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
         Visible         =   0   'False
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
         Visible         =   0   'False
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
         Visible         =   0   'False
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
      Begin VB.Label lblResp 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   2
         Left            =   4440
         TabIndex        =   39
         Top             =   960
         Visible         =   0   'False
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
         Visible         =   0   'False
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
         Top             =   600
         Width           =   2805
      End
      Begin VB.Label lblResp 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   5
         Left            =   4440
         TabIndex        =   33
         Top             =   600
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
         Visible         =   0   'False
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
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblResp 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   1
         Left            =   4440
         TabIndex        =   23
         Top             =   240
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
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.Label lblPasos 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1. Pasar Ventas"
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
         Visible         =   0   'False
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
         Visible         =   0   'False
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
      Height          =   6810
      Left            =   15120
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   3435
      _cx             =   6059
      _cy             =   12012
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
      FormatString    =   $"frmAnexoICE.frx":008C
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
      Left            =   60
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1200
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
      FormatString    =   $"frmAnexoICE.frx":00EF
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
      FormatString    =   $"frmAnexoICE.frx":0152
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
      Left            =   480
      TabIndex        =   41
      Top             =   420
      Width           =   2805
   End
End
Attribute VB_Name = "frmAnexoICE"
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
Const COL_C_NUMSERESTA = 10
Const COL_C_NUMSERPUNTO = 11
Const COL_C_NUMSECUENCIAL = 12
Const COL_C_NUMAUTOSRI = 13
Const COL_C_FECHACADUCIDAD = 14
Const COL_C_CODSUSTENTO = 15
Const COL_C_CODTIPOCOMP = 16
Const COL_C_TIPOPAGO = 17
Const COL_C_PAGOEXTERIOR = 18
Const COL_C_CODPAIS = 19
Const COL_C_DOBLETRIB = 20
Const COL_C_PAGOSUJRET = 21
Const COL_C_BASE0 = 22
Const COL_C_BASE12 = 23
Const COL_C_BASENO12 = 24
Const COL_C_CODICE = 25
Const COL_C_MONTOICE = 26
'Const COL_C_MONTOIVA = 27
Const COL_C_ENOTRARET = 27
Const COL_C_IVA = 28
Const COL_C_RESP = 29

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
Const COL_V_RESP = 13

Const COL_VE_SUC = 1
Const COL_VE_TIPOCOMP = 2
Const COL_VE_CANTRANS = 3
Const COL_VE_BASE0 = 4
Const COL_VE_BASEIVA = 5
Const COL_VE_BASENOIVA = 6
Const COL_VE_TOTAL = 7
Const COL_VE_RESP = 8


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

Public Sub Inicio(ByVal tag As String)
    On Error GoTo ErrTrap
    Set mObjCond = New RepCondicion
    Me.Caption = "Anexo ICE"

    TotalVentas = 0
    dtpPeriodo.value = CDate("01/" & IIf(Month(Date) - 1 <> 0, Month(Date) - 1, 12) & "/" & Year(Date))
    mObjCond.fecha1 = dtpPeriodo.value
    cboTipo.ListIndex = 0
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("RutaICE")) > 0 Then
        txtCarpeta.Text = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("RutaICE")
    End If
    Me.tag = tag
    Me.Show
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Function GenerarVentasICE(ByRef cad As String) As Boolean
    On Error GoTo ErrTrap
        GenerarVentasICE = False
        GenerarVentasICE = GeneraArchivoICEVentasXML(cad)
    Exit Function
ErrTrap:
    grd.Redraw = True
    DispErr
    Exit Function
End Function



Private Function GeneraArchivoICEVentasXML(ByRef cad As String) As Boolean
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
    GeneraArchivoICEVentasXML = True
    bandCF = False
    filaCF = 1
    For j = 1 To grdCF.Rows - 1
        grdCF.RemoveItem 1
    Next j
    
    For j = 1 To GrdRetVentas.Rows - 1
        GrdRetVentas.TextMatrix(j, 8) = ""
    Next j
    
    
        grd.Refresh
        cadenaF = "<ventas>"

            If grd.Rows < 1 Then
                prg.value = 0
                cadenaF = cadenaFC & "</ventas>"
                cad = cadenaF
                GeneraArchivoICEVentasXML = True
                GoTo SiguienteFila
            End If


            prg.max = grd.Rows - 1
            For i = 1 To grd.Rows - 1
                If grd.IsSubtotal(i) Then GoTo SiguienteFila
                grd.ShowCell i, 1
                prg.value = i
                DoEvents
                cadenaFC = ""
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
                
                cadenaFC = cadenaFC & "<vta>"
                cadenaFC = cadenaFC & "<codProdICE>3092-37-003516-068-000000-00-593-000000</codProdICE>"
                Select Case grd.TextMatrix(i, COL_V_TIPODOC)
                    Case "R":                     cadenaFC = cadenaFC & "<tipoIdCliente>" & "R" & "</tipoIdCliente>"
                    Case "C":                     cadenaFC = cadenaFC & "<tipoIdCliente>" & "C" & "</tipoIdCliente>"
                    Case "P":                     cadenaFC = cadenaFC & "<tipoIdCliente>" & "P" & "</tipoIdCliente>"
                    Case "F":                     cadenaFC = cadenaFC & "<tipoIdCliente>" & "F" & "</tipoIdCliente>"
                    Case "T":
                            msg = " El Cliente " & grd.TextMatrix(i, COL_V_CLIENTE) & " el tipo de Documento selecciona do es Valido"
                            grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                            grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                            grd.ShowCell i, grd.ColIndex("Resultado")
                            GeneraArchivoICEVentasXML = True
                            lblResp(1).Caption = "Error"
                            chkConsFinal.Visible = True
                            GoTo SiguienteFila
                    Case Else
                            
                            msg = " El Cliente " & grd.TextMatrix(i, COL_V_CLIENTE) & " No tiene seleccionado el tipo de Documento"
                            grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                            grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                            grd.ShowCell i, grd.ColIndex("Resultado")
                            GeneraArchivoICEVentasXML = True
                            lblResp(1).Caption = "Error"
                            chkConsFinal.Visible = True
                            GoTo SiguienteFila
                        
                End Select
                
                cadenaFC = cadenaFC & "<idCliente>" & grd.TextMatrix(i, COL_V_RUC) & "</idCliente>"
                cadenaFC = cadenaFC & "<tipoVentaICE>1</tipoVentaICE>"
                cadenaFC = cadenaFC & "<ventaICE>" & Format(Abs(grd.ValueMatrix(i, COL_V_CANTRANS)), "#0") & "</ventaICE>"
                cadenaFC = cadenaFC & "<devICE>0</devICE> <cantProdBajaICE>0</cantProdBajaICE>"
                 cadenaFC = cadenaFC & "</vta>"
                cadenaF = cadenaF & cadenaFC
                grd.ShowCell i, grd.ColIndex("Resultado")
                grd.TextMatrix(i, grd.ColIndex("Resultado")) = " OK "
        
SiguienteFila:
    Next i
        SubTotalizarCF (COL_V_TIPOCOMP)
        TotalizarCF
    grd.ColWidth(grd.ColIndex("Resultado")) = 5000
    prg.value = 0

    If bandCF Then
        cadenaF = cadenaF & GeneraArchivoICEVentasXMLCF
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
'    TotalVentas = grd.ValueMatrix(grd.Rows - 1, COL_V_BASE0) + grd.ValueMatrix(grd.Rows - 1, COL_V_BASEIVA) + grd.ValueMatrix(grd.Rows - 1, COL_V_BASENOIVA)
    Exit Function
cancelado:
    GeneraArchivoICEVentasXML = False
ErrTrap:
    grd.TextMatrix(grd.Rows - 1, 2) = Err.Description
    GeneraArchivoICEVentasXML = False
End Function

Private Function GeneraArchivoICEVentasXMLCF() As String
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
    GeneraArchivoICEVentasXMLCF = ""
            BandFact = False
            For i = 1 To grdCF.Rows - 2
               If grdCF.IsSubtotal(i) Then
                    cadenaFC = cadenaFC & "<vta>"
                    cadenaFC = cadenaFC & "<codProdICE>3092-37-003516-068-000000-00-593-000000</codProdICE>"
                    cadenaFC = cadenaFC & "<tipoIdCliente>F</tipoIdCliente>"
                    cadenaFC = cadenaFC & "<idCliente>9999999999999</idCliente>"
                    cadenaFC = cadenaFC & "<tipoVentaICE>1</tipoVentaICE>"
                    cadenaFC = cadenaFC & "<ventaICE>" & Format(Abs(grdCF.ValueMatrix(i, COL_V_CANTRANS)), "#0") & "</ventaICE>"
                    cadenaFC = cadenaFC & "<devICE>0</devICE> <cantProdBajaICE>0</cantProdBajaICE>"
                    cadenaFC = cadenaFC & "</vta>"
                End If
                grdCF.ShowCell i, grdCF.ColIndex("Resultado")
                grdCF.TextMatrix(i, grdCF.ColIndex("Resultado")) = " OK "
        
        Next i
    
    
    grdCF.ColWidth(grd.ColIndex("Resultado")) = 5000
    prg.value = 0
    GeneraArchivoICEVentasXMLCF = cadenaFC
    Exit Function
cancelado:
    GeneraArchivoICEVentasXMLCF = ""
ErrTrap:
    grdCF.TextMatrix(grd.Rows - 1, 2) = Err.Description
    GeneraArchivoICEVentasXMLCF = ""
End Function



Private Sub cmdPasos_Click(Index As Integer)
    Dim r As Boolean, cad As String, nombre As String, file As String
    NumProc = Index + 1
    
    If Index Mod 2 = 0 Then
        cmdPasos(0).BackColor = vbButtonFace
        cmdPasos(2).BackColor = vbButtonFace
        cmdPasos(4).BackColor = vbButtonFace
        cmdPasos(6).BackColor = vbButtonFace
        cmdPasos(8).BackColor = vbButtonFace
    End If
    
    
    Select Case Index + 1
    Case 1      '2. Busca Ventas
            BuscarVentasICE
            cadVentas = ""
            cmdPasos(2).BackColor = &HFFFF00
    Case 2      '2. Generar ventas
        lblResp(1).Caption = ""
        cadVentas = ""
        If cboTipo.ListIndex = 0 Then
            r = GenerarVentasICE(cadVentas)
        End If
    
    Case 11      '8. Generar Archivo
        If cboTipo.ListIndex = 0 Then
            'nombre = "ICE-" & Format(CStr(Month(dtpPeriodo.value)), "00") & Year(dtpPeriodo.value) & ".XML"
            nombre = "ICE-" & Format(CStr(Month(dtpPeriodo.value)), "00") & Year(dtpPeriodo.value) & "-" & gobjMain.EmpresaActual.GNOpcion.ruc & ".XML"
            file = txtCarpeta.Text & nombre
            If ExisteArchivo(file) Then
                If MsgBox("El nombre del archivo " & nombre & " ya existe desea sobreescribirlo?", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
            NumFile = FreeFile
            Open file For Output Access Write As #NumFile
            cadEncabezado = GeneraArchivoEncabezadoICEXML
            Cadena = cadEncabezado & cadVentas & "</ice>"
            Print #NumFile, Cadena
            Close NumFile
            
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "RutaICE", txtCarpeta.Text
            gobjMain.EmpresaActual.GNOpcion.Grabar
            
            r = True
            
        End If
        lblResp(5).Caption = "OK."
    
    End Select
    
    
    
    If r Then
        If Index < cmdPasos.Count - 1 Then
            If cboTipo.ListIndex = 0 Then
            If Index <> 10 Then
            Select Case Index
                Case 1
                    If lblResp(1).Caption <> "Error" Then
                        lblResp(1).BackColor = vbBlue
                        lblResp(1).ForeColor = vbYellow
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
                    If lblResp(1).Caption <> "Error" Then
                        lblResp(1).BackColor = vbBlue
                        lblResp(1).ForeColor = vbYellow
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
                    If lblResp(1).Caption <> "Error" Then
                        lblResp(1).BackColor = vbBlue
                        lblResp(1).ForeColor = vbYellow
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
    grd.Move 0, fraPasos.Height + 100, Me.ScaleWidth - 4000, (Me.ScaleHeight - (fraPasos.Height + picBoton.Height) - 105)
    grd.Move 0, fraPasos.Height + 100, Me.ScaleWidth, (Me.ScaleHeight - (fraPasos.Height + picBoton.Height) - 105)



    GrdRetVentas.Visible = False
    grdRet.Visible = False
    
    grdCF.Visible = False
    grdCF.Move 0, grd.Top + grd.Height + 100, Me.ScaleWidth, (Me.ScaleHeight - (fraPasos.Height + picBoton.Height) - 200) * 0.25
    
'    grdRet.Move 0, grd.Top + grd.Height + 100, Me.ScaleWidth, (Me.ScaleHeight - (fraPasos.Height + picBoton.Height) - 200) * 0.25
 '   GrdRetVentas.Move grd.Left + grd.Width, fraPasos.Height + 100, Me.ScaleWidth / 2, (Me.ScaleHeight - (fraPasos.Height + picBoton.Height) - 105) * 0.75
  '  grdCF.Height = 4000
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


Private Sub ConfigCols(cad As String)
    Dim s As String, i As Integer
    Select Case cad
    Case "IMPCPI"           'Compras
        s = "^#|TransId|^Fecha|^FechaReg|<CodTrans|<Num|^Doc.|idproveedorref|<RUC|<Proveedor|^Estab|^Punto|^Secuencial|>AutSRI|^Caducidad|^Sustento|^TipoComp|^Tipo Pago|^Pago Exterior|>CodPaisSRI  |^BandDobleTributa |^BandPagoSujRet|>Base Cero|>Base IVA|>Base NO IVA|<Cod ICE|>Val ICE|^BANDREToTRO|>IVA"
        '|>Base Ser|<Cod Ser|>Val Ser|>Base Bien|<Cod Bien|>Val Bien|>Base IR|<Cod IR|>Val IR|<NumDocRef|^NumSerieEstabRet|^NumSeriePuntoRet|^NumSecuencialRet|^NumAutSRIRet|^FechaEmisionRet|>Base ICE|<Cod ICE|>Val ICE"
        grd.FormatString = s & "|<         Resultado           "
        AsignarTituloAColKey grd
    
    Case "IMPCPIR"           ' Retencion Compras
        s = "^#|Tipo Ret|<Codigo|<Codigo SRI|>Porcen|<CodTrans|<NumTrans|<RUC|<CodTrans Ret|<NumTrans Ret|^FechaEmisionRet|^NumSerieEstabRet|^NumSeriePuntoRet|^NumSecuencialRet|^NumAutSRIRet....................................................|>Base Ret|>Valor Ret"
        grdRet.FormatString = s
        AsignarTituloAColKey grdRet
    Case "IMPFC"

        s = "^#|<Fecha|^Doc|<IdProvcli|<RUC|<Cliente|^Tipo Comp|^Cant Trans |>Base 0|>Base IVA|>Base NO IVA|>Valor IVA|>%IVA"
        grd.FormatString = s & "|<         Resultado           "
        AsignarTituloAColKey grd
    
        grdCF.FormatString = s & "|<         Resultado           "
        AsignarTituloAColKey grdCF
        
        s = "^#|<Fecha|^Doc|<IdProvcli|<RUC|<Cliente|^Tipo Comp|^Cant Trans |>Base 0|>Base IVA|>Base NO IVA|>IVA "
        GrdRetVentas.FormatString = s & "|<         Resultado           "
        
    Case "IMPFCxE"

        s = "^#|>Establecimiento|^Tipo Comp|>Cant. Documentos|>Base 0|>Base IVA|>Base NO IVA|>Total"
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
            
            


    Case "IMPCPIR"
            grdRet.ColHidden(COL_R_TRANS) = True
            grdRet.ColHidden(COL_R_NUMTRANS) = True
            grdRet.ColHidden(COL_R_RUC) = True
    Case "IMPFC"
            For i = 1 To COL_V_BASENOIVA
                grd.ColHidden(i) = False
            Next i
            grd.ColHidden(COL_V_IDPROVCLI) = True
            grd.ColHidden(COL_V_BASE0) = True
            grd.ColHidden(COL_V_BASEIVA) = True
            grd.ColHidden(COL_V_BASENOIVA) = True
            grd.ColHidden(COL_V_VALORIVA) = True
            grd.ColHidden(COL_V_TIPOCOMP) = True
            grd.ColHidden(COL_V_VALORIVA) = True
            grd.ColHidden(COL_V_IVA) = True
            
            
            grd.ColFormat(COL_V_BASE0) = "##0.00"
            grd.ColFormat(COL_V_BASEIVA) = "##0.00"
            grd.ColFormat(COL_V_BASENOIVA) = "##0.00"
            grd.ColFormat(COL_V_VALORIVA) = "##0.00"
    
'            grd.ColData(COL_V_CANTRANS) = "SubTotal"
            grd.ColData(COL_V_BASE0) = "SubTotal"
            grd.ColData(COL_V_BASEIVA) = "SubTotal"
            grd.ColData(COL_V_BASENOIVA) = "SubTotal"
            grd.ColData(COL_V_VALORIVA) = "SubTotal"
            grd.ColData(COL_V_CANTRANS) = "SubTotal"
            
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
    
            grd.ColData(COL_VE_CANTRANS) = "SubTotal"
            grd.ColData(COL_VE_BASE0) = "SubTotal"
            grd.ColData(COL_VE_BASEIVA) = "SubTotal"
            grd.ColData(COL_VE_BASENOIVA) = "SubTotal"
            grd.ColData(COL_VE_TOTAL) = "SubTotal"
            
    
    Case "IMPCA"
            For i = 1 To COL_A_RESP
                grd.ColHidden(i) = False
'                grd.ColFormat(i) = flexDTString
            Next i
            grd.ColFormat(COL_A_NUMESTA) = "000"
           grd.ColFormat(COL_A_NUMPUNTO) = "000"
            grd.ColFormat(COL_A_NUMSECUE) = "0000000"
    End Select
    
    grd.ColSort(1) = flexSortGenericAscending
    grd.ColSort(2) = flexSortGenericAscending
    grd.ColSort(3) = flexSortGenericAscending
    grd.ColSort(4) = flexSortGenericAscending

    AsignarTituloAColKey grd
    grd.SetFocus

End Sub



Private Function GeneraArchivoEncabezadoICEXML() As String
    Dim obj As GNOpcion, cad As String, numSucursal As Integer
    cad = "<?xml version=" & """1.0""" & " encoding=" & """UTF-8""" & "" & " standalone=" & """no""" & "?>"
    cad = cad & "<!--  Generado por Ishida Asociados   -->"
    cad = cad & "<!--  Dir: Av. Gonzalez Suarez y Rayoloma Tercer Piso -->"
    cad = cad & "<!--  Telf: 098499003, 072870346, 072871094      -->"
    cad = cad & "<!--  email: ishidacue@hotmail.com, aquizhpe@ibzssoft.com    -->"
    cad = cad & "<!--  www.ibzssoft.com    -->"
    cad = cad & "<!--  Cuenca - Ecuador                -->"
    cad = cad & "<!--  SISTEMAS DE GESTION EMPRESASRIAL-->"
        
    cad = cad & "<ice>"
        
    cad = cad & "<TipoIDInformante> R </TipoIDInformante>"
    cad = cad & "<IdInformante>" & Format(gobjMain.EmpresaActual.GNOpcion.ruc, "0000000000000") & "</IdInformante>"
    cad = cad & "<razonSocial>" & UCase(gobjMain.EmpresaActual.GNOpcion.RazonSocial) & "</razonSocial>"
    cad = cad & "<Anio>" & Year(mObjCond.fecha1) & "</Anio>"
    cad = cad & "<Mes>" & IIf(Len(Month(mObjCond.fecha1)) = 1, "0" & Month(mObjCond.fecha1), Month(mObjCond.fecha1)) & "</Mes>"
    cad = cad & "<actImport>02</actImport>"
    cad = cad & "<codigoOperativo>ICE</codigoOperativo>"
    

    GeneraArchivoEncabezadoICEXML = cad
End Function

Public Function RellenaDer(ByVal s As String, lon As Long) As String
    Dim r As String
    r = "!" & String(lon, "@")
    If Len(s) = 0 Then s = " "
    RellenaDer = Format(s, r)
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

Private Sub txtCarpeta_LostFocus()
    If Right$(txtCarpeta.Text, 1) <> "\" Then
        txtCarpeta.Text = txtCarpeta.Text & "\"
    End If
    'Luego a actualiza linea de comando
End Sub


Private Function BuscarVentasICE()
    On Error GoTo ErrTrap
        With grd
        .Redraw = False
        .Rows = .FixedRows
        If Not frmB_Trans.Inicio(gobjMain, "IMPFCICE", dtpPeriodo.value) Then
            grd.SetFocus
        End If
        mObjCond.fecha1 = gobjMain.objCondicion.fecha1
        mObjCond.fecha2 = gobjMain.objCondicion.fecha2
        MiGetRowsRep gobjMain.EmpresaActual.ConsANVentas2013paraXML(), grd
        MiGetRowsRep gobjMain.EmpresaActual.ConsANTotalRetencionVentas2008ParaXML, GrdRetVentas

        'GeneraArchivo

        ConfigCols "IMPFC"
        ConfigCols "IMPFCIR"
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
                    cadenaFC = cadenaFC & "<numeroComprobantes>0</numeroComprobantes>"
                    cadenaFC = cadenaFC & "<baseNoGraIva>0.00</baseNoGraIva>"
                    cadenaFC = cadenaFC & "<baseImponible>0.00</baseImponible>"
                    cadenaFC = cadenaFC & "<baseImpGrav>0.00</baseImpGrav>"
                    cadenaFC = cadenaFC & "<montoIva>0.00</montoIva>"
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
                    
                    
                     cadenaFC = cadenaFC & cadenaFCIVA
                    cadenaFC = cadenaFC & cadenaFCIR
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

