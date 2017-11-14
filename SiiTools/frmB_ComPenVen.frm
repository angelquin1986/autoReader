VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{50067EB3-D6AF-11D3-8297-000021C5085D}#1.0#0"; "NTextBox.ocx"
Begin VB.Form frmB_ComPenVen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda"
   ClientHeight    =   6825
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5625
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6195
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   10927
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Parametros"
      TabPicture(0)   =   "frmB_ComPenVen.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraVenta"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraVendedor"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraFecha"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Vendedores"
      TabPicture(1)   =   "frmB_ComPenVen.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tabla Comisiones 1-2"
      TabPicture(2)   =   "frmB_ComPenVen.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "FrnmTablaComisionesB"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "FrnmTablaComisiones"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Tabla Comisiones 3-4"
      TabPicture(3)   =   "frmB_ComPenVen.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(1)=   "Frame3"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame4 
         Caption         =   "Tabla de Comisiones 3"
         Height          =   2595
         Left            =   -74880
         TabIndex        =   40
         Top             =   720
         Width           =   5055
         Begin VB.TextBox txtArchiOrigenC 
            Height          =   372
            Left            =   720
            TabIndex        =   42
            Top             =   300
            Width           =   3675
         End
         Begin VB.CommandButton cmdBuscaArchiC 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   4500
            TabIndex        =   41
            Top             =   300
            Width           =   372
         End
         Begin MSComDlg.CommonDialog dlg1C 
            Left            =   4560
            Top             =   180
            _ExtentX        =   688
            _ExtentY        =   688
            _Version        =   393216
         End
         Begin VSFlex7LCtl.VSFlexGrid grdComisionesC 
            Height          =   1275
            Left            =   120
            TabIndex        =   43
            Top             =   720
            Width           =   4785
            _cx             =   8440
            _cy             =   2249
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
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   11
            Cols            =   3
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
         Begin NTextBoxProy.NTextBox ntxPorcenVendedorC 
            Height          =   255
            Left            =   1740
            TabIndex        =   72
            Top             =   2160
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   450
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
         Begin NTextBoxProy.NTextBox ntxPorcenCobradorC 
            Height          =   315
            Left            =   4140
            TabIndex        =   73
            Top             =   2130
            Width           =   555
            _ExtentX        =   979
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
         Begin VB.Label Label22 
            Caption         =   "%"
            Height          =   255
            Left            =   4740
            TabIndex        =   63
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label Label21 
            Caption         =   "%"
            Height          =   255
            Left            =   2280
            TabIndex        =   62
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label Label20 
            Caption         =   "Porcentaje Vendedor"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   2160
            Width           =   1815
         End
         Begin VB.Label Label19 
            Caption         =   "Porcentaje Cobrador"
            Height          =   255
            Left            =   2640
            TabIndex        =   60
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Archivo"
            Height          =   195
            Left            =   60
            TabIndex        =   44
            Top             =   420
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tabla de Comisiones 4"
         Height          =   2655
         Left            =   -74880
         TabIndex        =   35
         Top             =   3420
         Width           =   5055
         Begin VB.CommandButton cmdBuscaArchiD 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   4500
            TabIndex        =   37
            Top             =   300
            Width           =   372
         End
         Begin VB.TextBox txtArchiOrigenD 
            Height          =   372
            Left            =   720
            TabIndex        =   36
            Top             =   300
            Width           =   3675
         End
         Begin MSComDlg.CommonDialog dlg1D 
            Left            =   4560
            Top             =   180
            _ExtentX        =   688
            _ExtentY        =   688
            _Version        =   393216
         End
         Begin VSFlex7LCtl.VSFlexGrid grdComisionesD 
            Height          =   1275
            Left            =   180
            TabIndex        =   38
            Top             =   780
            Width           =   4785
            _cx             =   8440
            _cy             =   2249
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
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   11
            Cols            =   3
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
         Begin NTextBoxProy.NTextBox ntxPorcenVendedorD 
            Height          =   255
            Left            =   1740
            TabIndex        =   74
            Top             =   2220
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   450
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
         Begin NTextBoxProy.NTextBox ntxPorcenCobradorD 
            Height          =   315
            Left            =   4200
            TabIndex        =   75
            Top             =   2190
            Width           =   555
            _ExtentX        =   979
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
         Begin VB.Label Label26 
            Caption         =   "%"
            Height          =   255
            Left            =   4800
            TabIndex        =   67
            Top             =   2220
            Width           =   135
         End
         Begin VB.Label Label25 
            Caption         =   "%"
            Height          =   255
            Left            =   2340
            TabIndex        =   66
            Top             =   2220
            Width           =   135
         End
         Begin VB.Label Label24 
            Caption         =   "Porcentaje Vendedor"
            Height          =   255
            Left            =   180
            TabIndex        =   65
            Top             =   2220
            Width           =   1815
         End
         Begin VB.Label Label23 
            Caption         =   "Porcentaje Cobrador"
            Height          =   255
            Left            =   2700
            TabIndex        =   64
            Top             =   2220
            Width           =   1575
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Archivo"
            Height          =   195
            Left            =   60
            TabIndex        =   39
            Top             =   420
            Width           =   540
         End
      End
      Begin VB.Frame FrnmTablaComisiones 
         Caption         =   "Tabla de Comisiones 1"
         Height          =   2595
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   5055
         Begin VB.TextBox txtArchiOrigen 
            Height          =   372
            Left            =   720
            TabIndex        =   32
            Top             =   300
            Width           =   3675
         End
         Begin VB.CommandButton cmdBuscaArchi 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   4500
            TabIndex        =   31
            Top             =   300
            Width           =   372
         End
         Begin MSComDlg.CommonDialog dlg1 
            Left            =   4560
            Top             =   180
            _ExtentX        =   688
            _ExtentY        =   688
            _Version        =   393216
         End
         Begin VSFlex7LCtl.VSFlexGrid grdComisiones 
            Height          =   1275
            Left            =   120
            TabIndex        =   33
            Top             =   720
            Width           =   4785
            _cx             =   8440
            _cy             =   2249
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
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   11
            Cols            =   3
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
         Begin NTextBoxProy.NTextBox ntxPorcenVendedor 
            Height          =   255
            Left            =   1740
            TabIndex        =   68
            Top             =   2130
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   450
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
         Begin NTextBoxProy.NTextBox ntxPorcenCobrador 
            Height          =   315
            Left            =   4200
            TabIndex        =   69
            Top             =   2100
            Width           =   555
            _ExtentX        =   979
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
         Begin VB.Label Label14 
            Caption         =   "%"
            Height          =   255
            Left            =   4800
            TabIndex        =   55
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label Label13 
            Caption         =   "%"
            Height          =   255
            Left            =   2340
            TabIndex        =   54
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label Label12 
            Caption         =   "Porcentaje Cobrador"
            Height          =   255
            Left            =   2700
            TabIndex        =   53
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label Label11 
            Caption         =   "Porcentaje Vendedor"
            Height          =   255
            Left            =   180
            TabIndex        =   52
            Top             =   2160
            Width           =   1815
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Archivo"
            Height          =   195
            Left            =   60
            TabIndex        =   34
            Top             =   420
            Width           =   540
         End
      End
      Begin VB.Frame FrnmTablaComisionesB 
         Caption         =   "Tabla de Comisiones 2"
         Height          =   2655
         Left            =   120
         TabIndex        =   25
         Top             =   3420
         Width           =   5055
         Begin VB.CommandButton cmdBuscaArchiB 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   4500
            TabIndex        =   27
            Top             =   300
            Width           =   372
         End
         Begin VB.TextBox txtArchiOrigenB 
            Height          =   372
            Left            =   720
            TabIndex        =   26
            Top             =   300
            Width           =   3675
         End
         Begin MSComDlg.CommonDialog dlg1b 
            Left            =   4560
            Top             =   180
            _ExtentX        =   688
            _ExtentY        =   688
            _Version        =   393216
         End
         Begin VSFlex7LCtl.VSFlexGrid grdComisionesB 
            Height          =   1275
            Left            =   180
            TabIndex        =   28
            Top             =   780
            Width           =   4785
            _cx             =   8440
            _cy             =   2249
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
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   11
            Cols            =   3
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
         Begin NTextBoxProy.NTextBox ntxPorcenVendedorB 
            Height          =   255
            Left            =   1680
            TabIndex        =   70
            Top             =   2160
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   450
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
         Begin NTextBoxProy.NTextBox ntxPorcenCobradorB 
            Height          =   315
            Left            =   4200
            TabIndex        =   71
            Top             =   2130
            Width           =   555
            _ExtentX        =   979
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
         Begin VB.Label Label18 
            Caption         =   "%"
            Height          =   255
            Left            =   4800
            TabIndex        =   59
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label Label17 
            Caption         =   "%"
            Height          =   255
            Left            =   2340
            TabIndex        =   58
            Top             =   2160
            Width           =   135
         End
         Begin VB.Label Label16 
            Caption         =   "Porcentaje Vendedor"
            Height          =   255
            Left            =   180
            TabIndex        =   57
            Top             =   2160
            Width           =   1815
         End
         Begin VB.Label Label15 
            Caption         =   "Porcentaje Cobrador"
            Height          =   255
            Left            =   2700
            TabIndex        =   56
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Archivo"
            Height          =   195
            Left            =   60
            TabIndex        =   29
            Top             =   420
            Width           =   540
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Clasificación de Vendores"
         Height          =   5325
         Left            =   -74940
         TabIndex        =   20
         Top             =   660
         Width           =   5115
         Begin VB.ListBox lstD 
            Height          =   1035
            Left            =   2940
            TabIndex        =   50
            Top             =   3960
            Width           =   2055
         End
         Begin VB.ListBox lstC 
            Height          =   1035
            Left            =   2940
            TabIndex        =   48
            Top             =   2760
            Width           =   2055
         End
         Begin VB.ListBox lstB 
            Height          =   1035
            Left            =   2940
            TabIndex        =   45
            Top             =   1560
            Width           =   2055
         End
         Begin VB.ListBox lstVendedores 
            Height          =   4350
            Left            =   180
            TabIndex        =   22
            Top             =   540
            Width           =   2055
         End
         Begin VB.ListBox lstA 
            Height          =   1035
            ItemData        =   "frmB_ComPenVen.frx":0070
            Left            =   2940
            List            =   "frmB_ComPenVen.frx":0072
            TabIndex        =   21
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label lblDragDrop 
            Caption         =   "Label11"
            Height          =   135
            Left            =   180
            TabIndex        =   51
            Top             =   480
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "TABLA 4"
            Height          =   195
            Left            =   3600
            TabIndex        =   49
            Top             =   3780
            Width           =   645
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "TABLA 3"
            Height          =   195
            Left            =   3540
            TabIndex        =   47
            Top             =   2580
            Width           =   645
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "TABLA 1"
            Height          =   195
            Left            =   3600
            TabIndex        =   46
            Top             =   120
            Width           =   645
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "VENDEDORES"
            Height          =   195
            Left            =   720
            TabIndex        =   24
            Top             =   300
            Width           =   1125
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "TABLA 2"
            Height          =   195
            Left            =   3540
            TabIndex        =   23
            Top             =   1380
            Width           =   645
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Recargos Descuentos antes IVA"
         Height          =   1605
         Left            =   -74880
         TabIndex        =   15
         Top             =   2580
         Width           =   5115
         Begin VB.ListBox lstFuente 
            Height          =   1230
            Left            =   120
            TabIndex        =   19
            Top             =   255
            Width           =   2055
         End
         Begin VB.ListBox lstDestino 
            Height          =   1230
            Left            =   2880
            TabIndex        =   18
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&>>"
            Height          =   375
            Left            =   2220
            TabIndex        =   17
            Top             =   480
            Width           =   615
         End
         Begin VB.CommandButton cmdResta 
            Caption         =   "&<<"
            Height          =   375
            Left            =   2220
            TabIndex        =   16
            Top             =   900
            Width           =   615
         End
      End
      Begin VB.Frame fraVenta 
         Caption         =   "Transacciones de Venta"
         Height          =   1575
         Left            =   -74880
         TabIndex        =   13
         Top             =   4200
         Width           =   5115
         Begin VB.ListBox lst 
            Height          =   1230
            IntegralHeight  =   0   'False
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   14
            Top             =   240
            Width           =   4875
         End
      End
      Begin VB.Frame fraVendedor 
         Caption         =   "Vendedor"
         Height          =   735
         Left            =   -74880
         TabIndex        =   10
         Top             =   1080
         Width           =   5115
         Begin FlexComboProy.FlexCombo fcbDesde1 
            Height          =   375
            Left            =   840
            TabIndex        =   0
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            ColWidth1       =   1500
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
         Begin FlexComboProy.FlexCombo fcbHasta1 
            Height          =   375
            Left            =   3240
            TabIndex        =   1
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            ColWidth1       =   1500
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
         Begin VB.Label lblDesde 
            Caption         =   "Desde"
            Height          =   252
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   612
         End
         Begin VB.Label lblHasta 
            Caption         =   "Hasta"
            Height          =   255
            Left            =   2640
            TabIndex        =   11
            Top             =   300
            Width           =   615
         End
      End
      Begin VB.Frame fraFecha 
         Caption         =   "Rango de Fecha Venta"
         Height          =   675
         Left            =   -74880
         TabIndex        =   7
         Top             =   1860
         Width           =   5115
         Begin MSComCtl2.DTPicker dtpHasta 
            Height          =   360
            Left            =   3240
            TabIndex        =   3
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   635
            _Version        =   393216
            Format          =   105578497
            CurrentDate     =   36891
         End
         Begin MSComCtl2.DTPicker dtpDesde 
            Height          =   360
            Left            =   840
            TabIndex        =   2
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   635
            _Version        =   393216
            Format          =   105578497
            CurrentDate     =   36526
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "D&esde  "
            Height          =   195
            Left            =   180
            TabIndex        =   9
            Top             =   300
            Width           =   570
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "H&asta  "
            Height          =   195
            Left            =   2580
            TabIndex        =   8
            Top             =   300
            Width           =   510
         End
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   400
      Left            =   2760
      TabIndex        =   5
      Top             =   6360
      Width           =   1200
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar -F5"
      Height          =   400
      Left            =   1320
      TabIndex        =   4
      Top             =   6360
      Width           =   1320
   End
End
Attribute VB_Name = "frmB_ComPenVen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BandAceptado As Boolean
'Private WithEvents mobjReporte As ReporteMain

Public Function InicioComPenVen(ByRef objcond As Condicion, _
                                                    Optional ByRef T1 As String, Optional ByRef T2 As String, _
                                                    Optional ByRef T3 As String, Optional ByRef T4 As String) As Boolean
                           
    Dim i As Integer, KeyT As String, KeyTC As String, pos As Integer
    Dim KeyRecargo As String, KeyVenta As String
    Dim keyDias As String, keyComision As String
    Dim keyPenal As String, keyTipoCom As String
    SSTab1.Tab = 0
    RecuperarConfig
'    LeerIntervalos
'    LeerIntervalosB
'    LeerIntervalosC
'    LeerIntervalosD
    
    LeerIntervalosGnOpcion "A"
    LeerIntervalosGnOpcion "B"
    LeerIntervalosGnOpcion "C"
    LeerIntervalosGnOpcion "D"
    
    Dim trans As String
    Me.tag = Name  'nombre del reporte
    'Prepara la lista de monedas
    CargaTablaComisiones
    CargaTablaComisionesB
    CargaTablaComisionesC
    CargaTablaComisionesD
    txtArchiOrigen.Text = gConfigura.Archivo
    txtArchiOrigenB.Text = gConfigura.ArchivoB
    txtArchiOrigenC.Text = gConfigura.ArchivoC
    txtArchiOrigenD.Text = gConfigura.ArchivoD
    ntxPorcenVendedor.value = gConfigura.PorcenVendedorA
    ntxPorcenVendedorB.value = gConfigura.PorcenVendedorB
    ntxPorcenVendedorC.value = gConfigura.PorcenVendedorC
    ntxPorcenVendedorD.value = gConfigura.PorcenVendedorD
    
    ntxPorcenCobrador.value = gConfigura.PorcenCobradorA
    ntxPorcenCobradorB.value = gConfigura.PorcenCobradorB
    ntxPorcenCobradorC.value = gConfigura.PorcenCobradorC
    ntxPorcenCobradorD.value = gConfigura.PorcenCobradorD
    
    CargaRecargo
    cargaVendedores lstVendedores, 0
    cargaVendedores lstA, 1
    cargaVendedores lstB, 2
    cargaVendedores lstC, 3
    cargaVendedores lstD, 4
    With objcond
        dtpDesde.value = IIf(.fecha1 = 0, dtpDesde.value, dtpDesde.value)
        dtpHasta.value = IIf(.fecha2 = 0, dtpHasta.value, dtpHasta.value)
        CargaTipoTrans "", lst
        'Prepara la lista de vendedores que irá en el combo
        fcbDesde1.SetData gobjMain.EmpresaActual.ListaFCVendedorN(False, False, True, False)
        fcbHasta1.SetData gobjMain.EmpresaActual.ListaFCVendedorN(False, False, True, False)
        BandAceptado = False
        'Valores predeterminados
        fcbDesde1.KeyText = .CodCentro1 '  Vendedor1
        fcbHasta1.KeyText = .CodCentro2  '.Vendedor2
        trans = GetSetting(APPNAME, App.Title, "Trans_Comisiones", "_VACIO_")
        RecuperaSelec KeyT, lst, trans
        trans = GetSetting(APPNAME, App.Title, "RecarDesc_Comisiones", "_VACIO_")
        RecuperaSelecRec KeyRecargo, lstFuente, lstDestino, trans
        
        pos = InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "HORMI")
        If pos > 0 Then
            Label7.Caption = "QUITO"
            Label4.Caption = "GUAYAQUIL"
            Label9.Caption = "CUENCA"
            Label10.Caption = "OTROS"
        End If
        
'        ntxPorcenVendedor
        Me.Show vbModal, frmMain
        'Si aplastó el botón 'Aceptar'
        If BandAceptado Then
            'Devuelve los valores de condición para la búsqueda
            .CodCentro1 = Trim$(fcbDesde1.Text)
            .CodCentro2 = Trim$(fcbHasta1.Text)
            .fecha1 = dtpDesde.value
            .fecha2 = dtpHasta.value
            .CodTrans = PreparaCadena(lst) ' .TipoTrans
            .Servicios = PreparaCadRec(lstDestino)
            T1 = PreparaCadRec(lstA)
            T2 = PreparaCadRec(lstB)
            T3 = PreparaCadRec(lstC)
            T4 = PreparaCadRec(lstD)
            GrabaConfig
            GrabaIntervalosA    '23/11/2000
            GrabaIntervalosB
            GrabaIntervalosC
            GrabaIntervalosD
            
            
            ActualizaVendedor ' pone a todos en tipotabla=0
            GrabarVendedores lstA, 1
            GrabarVendedores lstB, 2
            GrabarVendedores lstC, 3
            GrabarVendedores lstD, 4
            SaveSetting APPNAME, App.Title, "Trans_Comisiones", .CodTrans
            SaveSetting APPNAME, App.Title, "RecarDesc_Comisiones", .Servicios
        End If
    End With
    'Devuelve true/false
    Unload Me
    InicioComPenVen = BandAceptado
End Function


Private Function PreparaCadena(lst As ListBox) As String
    Dim Cadena As String, i As Integer
    Cadena = ""
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            If Cadena = "" Then
                Cadena = Left(lst.List(i), lst.ItemData(i))
            Else
                Cadena = Cadena & "," & _
                              Left(lst.List(i), lst.ItemData(i))
            End If
        End If
    Next i
    PreparaCadena = Cadena
End Function

Private Sub PreparaListaTransIV()
    Dim rs As Recordset
   'Prepara la lista de tipos de transaccion
    lst.Clear
    Set rs = gobjMain.EmpresaActual.ListaGNTrans("IV", False, True)
    With rs
        If Not (.EOF) Then
            .MoveFirst
            Do Until .EOF
                lst.AddItem !CodTrans & "  " & !NombreTrans
                lst.ItemData(lst.NewIndex) = Len(!CodTrans)
                .MoveNext
            Loop
        End If
    End With
    rs.Close
    Set rs = Nothing
End Sub

Private Sub cmdAceptar_Click()
    BandAceptado = True
    dtpDesde.SetFocus
    Me.Hide
End Sub


Private Sub cmdAddA_Click()
    Dim i As Long, ix As Long
    On Error GoTo ErrTrap
    With lstA
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                'ix = mobjGrupo.AgregarUsuario(.List(i))
                ix = .ItemData(i)
                lstB.AddItem .List(i)
                lstB.ItemData(lstB.NewIndex) = ix
                .RemoveItem i
            End If
        Next i
    End With
    Exit Sub
ErrTrap:
    DispErr

End Sub

Private Sub cmdBuscaArchi_Click()
    txtArchiOrigen.Text = Archivo
End Sub

Private Sub cmdBuscaArchiB_Click()
    txtArchiOrigenB.Text = ArchivoB
End Sub

Private Sub cmdBuscaArchiC_Click()
    txtArchiOrigenC.Text = ArchivoC
End Sub

Private Sub cmdBuscaArchiD_Click()
    txtArchiOrigenD.Text = ArchivoD
End Sub

Private Sub cmdCancelar_Click()
    BandAceptado = False
    dtpDesde.SetFocus
    Me.Hide
End Sub


Private Sub cmdRestaA_Click()
    Dim i As Long, ix As Long
    On Error GoTo ErrTrap
    With lstB
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                ix = .ItemData(i)
                lstA.AddItem .List(i)
                lstA.ItemData(lstA.NewIndex) = ix
                .RemoveItem i
            End If
        Next i
    End With
    Exit Sub
ErrTrap:
    DispErr
End Sub

Private Sub fcbDesde1_Selected(ByVal Text As String, ByVal KeyText As String)
    fcbHasta1.Text = fcbDesde1.Text
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF5
        cmdAceptar_Click
        KeyCode = 0
    Case Else
        MoverCampo Me, KeyCode, Shift, False
    End Select
End Sub


Private Sub Form_Load()
    Dim mes As Integer, anio As Integer
    'Establece los rangos de Fecha  siempre  al rango
    'del año actual
    mes = Month(Date) - 1
    anio = Year(Date)
    If mes < 1 Then
        mes = 12
        anio = anio - 1
    End If
    dtpDesde.value = CDate("01/" & mes & "/" & anio)
    dtpHasta.value = DateAdd("d", -1, DateAdd("m", 1, (dtpDesde.value)))
    SSTab1.Tab = 0
End Sub

Private Function PreparaCadRec(lst As ListBox) As String
    Dim Cadena As String, i As Integer
    Cadena = ""
    For i = 0 To lst.ListCount - 1
        If Cadena = "" Then
            Cadena = Left(lst.List(i), lst.ItemData(i))
        Else
            Cadena = Cadena & "," & _
                          Left(lst.List(i), lst.ItemData(i))
        End If
    Next i
    PreparaCadRec = Cadena
End Function


Private Sub CargaTablaComisiones()
    Dim i As Integer
'    LeerIntervalos
    ConfigColsComisiones
    For i = 1 To 10
        grdComisiones.TextMatrix(i, 0) = gComisiones(i).desde
        grdComisiones.TextMatrix(i, 1) = gComisiones(i).hasta
        grdComisiones.TextMatrix(i, 2) = gComisiones(i).Comision
        grdComisiones.TextMatrix(i, 3) = gComisiones(i).ComisionC
        grdComisiones.TextMatrix(i, 4) = gComisiones(i).ComisionSC
    Next i
    ConfigColsComisiones
End Sub

Private Sub ConfigColsComisiones() 'grilla para el Impuesto a la Renta
    With grdComisiones
        .FormatString = ">Desde|>Hasta|>%Comi V|>% Comi C|>% Comi SC"
        .ColWidth(0) = 800
        .ColWidth(1) = 800
        .ColWidth(2) = 800
        .ColWidth(3) = 800
        .ColWidth(4) = 800
        .ColDataType(0) = flexDTCurrency
        .ColDataType(1) = flexDTCurrency
        .ColDataType(2) = flexDTCurrency
        .ColDataType(3) = flexDTCurrency
        .ColDataType(4) = flexDTCurrency
'        .ColFormat(0) = mobjGNComp.FormatoMoneda
'        .ColFormat(1) = mobjGNComp.FormatoMoneda
'        .ColFormat(2) = mobjGNComp.FormatoMoneda
    End With
    With grdComisionesB
        .FormatString = ">Desde|>Hasta|>% Comi V|>% Comi C|>% Comi SC"
        .ColWidth(0) = 800
        .ColWidth(1) = 800
        .ColWidth(2) = 800
        .ColWidth(3) = 800
        .ColWidth(4) = 800
        .ColDataType(0) = flexDTCurrency
        .ColDataType(1) = flexDTCurrency
        .ColDataType(2) = flexDTCurrency
        .ColDataType(3) = flexDTCurrency
        .ColDataType(4) = flexDTCurrency
    End With

    With grdComisionesC
        .FormatString = ">Desde|>Hasta|>% Comi V|>% Comi C|>% Comi SC"
        .ColWidth(0) = 800
        .ColWidth(1) = 800
        .ColWidth(2) = 800
        .ColWidth(3) = 800
        .ColWidth(4) = 800
        .ColDataType(0) = flexDTCurrency
        .ColDataType(1) = flexDTCurrency
        .ColDataType(2) = flexDTCurrency
        .ColDataType(3) = flexDTCurrency
        .ColDataType(4) = flexDTCurrency
    End With
    With grdComisionesD
        .FormatString = ">Desde|>Hasta|>% Comi V|>% Comi C|>% Comi SC"
        .ColWidth(0) = 800
        .ColWidth(1) = 800
        .ColWidth(2) = 800
        .ColWidth(3) = 800
        .ColWidth(4) = 800
        .ColDataType(0) = flexDTCurrency
        .ColDataType(1) = flexDTCurrency
        .ColDataType(2) = flexDTCurrency
        .ColDataType(3) = flexDTCurrency
        .ColDataType(4) = flexDTCurrency
    End With

End Sub

Private Sub GrabaIntervalosA()
    Dim i As Integer
    With grdComisiones
        For i = .FixedRows To .Rows - 1
            gComisiones(i).desde = .ValueMatrix(i, 0) ' CCur(Format("0" & .TextMatrix(i, 0), mobjGNComp.FormatoMoneda))
            gComisiones(i).hasta = .ValueMatrix(i, 1) 'CCur(Format("0" & .TextMatrix(i, 1), mobjGNComp.FormatoMoneda))
            gComisiones(i).Comision = .ValueMatrix(i, 2) ' CCur(Format("0" & .TextMatrix(i, 2), mobjGNComp.FormatoMoneda))
            gComisiones(i).ComisionC = .ValueMatrix(i, 3) ' CCur(Format("0" & .TextMatrix(i, 2), mobjGNComp.FormatoMoneda))
            gComisiones(i).ComisionSC = .ValueMatrix(i, 4) ' CCur(Format("0" & .TextMatrix(i, 2), mobjGNComp.FormatoMoneda))
        Next i
    End With
'    EscribirIntervalosA
    EscribirIntervalosGnOpcion "A"
End Sub

Private Sub grdComisiones_BeforeEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
    grdComisiones.EditMaxLength = 12 'Hasta 99,999,999,999
End Sub

Private Sub grdComisiones_KeyPressEdit(ByVal Row As Long, ByVal col As Long, KeyAscii As Integer)
    'Acepta sólo númericos
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And (KeyAscii <> vbKeyBack) And (KeyAscii <> Asc(".")) Then
        KeyAscii = 0
    End If
End Sub

Private Function Archivo() As String
    Dim i As Integer
    On Error GoTo mensaje
    
    With dlg1
        .DialogTitle = "Abrir"
        .CancelError = True
'        .InitDir = gobjRol.EmpresaActual.Ruta
        .flags = cdlOFNFileMustExist
        .DefaultExt = "txt"
        .Filter = "Archivos Texto|*.txt"
        .filename = gConfigura.Archivo
        .ShowOpen
        Archivo = .filename
        gConfigura.Archivo = Archivo
    End With
    Exit Function
    
mensaje:
    MsgBox "Se ha selccionado Cancelar"
    If Len(txtArchiOrigen.Text) > 0 Then Archivo = txtArchiOrigen.Text
    Exit Function
End Function

Private Sub cmdAdd_Click()
    Dim i As Long, ix As Long
    On Error GoTo ErrTrap
    With lstFuente
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                'ix = mobjGrupo.AgregarUsuario(.List(i))
                ix = .ItemData(i)
                lstDestino.AddItem .List(i)
                lstDestino.ItemData(lstDestino.NewIndex) = ix
                .RemoveItem i
            End If
        Next i
    End With
    Exit Sub
ErrTrap:
    DispErr
End Sub

Private Sub cmdResta_Click()
    Dim i As Long, ix As Long
    On Error GoTo ErrTrap
    With lstDestino
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                ix = .ItemData(i)
                lstFuente.AddItem .List(i)
                lstFuente.ItemData(lstFuente.NewIndex) = ix
                .RemoveItem i
            End If
        Next i
    End With
    Exit Sub
ErrTrap:
    DispErr
End Sub


Private Sub lstA_DblClick()
    Regresa lstA, lstVendedores
End Sub

Private Sub lstB_dblClick()
    Regresa lstB, lstVendedores
End Sub

Private Sub lstC_Click()
    Regresa lstC, lstVendedores
End Sub

Private Sub lstD_Click()
    Regresa lstD, lstVendedores
End Sub

Private Sub lstFuente_DblClick()
    cmdAdd_Click
End Sub

Private Sub lstDestino_DblClick()
    cmdResta_Click
End Sub

Private Sub CargaRecargo()
    Dim rs As Recordset
    Set rs = gobjMain.EmpresaActual.ListaIVRecargo(True)
    With rs
        If Not (.EOF) Then
            .MoveFirst
            Do Until .EOF
                lstFuente.AddItem !codRecargo & "  " & !Descripcion
                lstFuente.ItemData(lstFuente.NewIndex) = Len(!codRecargo)
               .MoveNext
           Loop
           
            lstFuente.AddItem "SUBT" & "  " & "Subtotal"
            lstFuente.ItemData(lstFuente.NewIndex) = Len("SUBT")
            
        End If
    End With
    rs.Close
End Sub

Public Sub RecuperaSelecRec(ByVal Key As String, lstF As ListBox, lstD As ListBox, trans As String)
Dim s As String, Vector As Variant, ix As Long
Dim i As Integer, j As Integer, Selec As Integer

    s = trans           '  jeaa 20/09/2003
    If s <> "_VACIO_" Then
        Vector = Split(s, ",")
         Selec = UBound(Vector, 1)
         For i = 0 To Selec
            For j = lstF.ListCount - 1 To 0 Step -1
                If Vector(i) = Left(lstF.List(j), lstF.ItemData(j)) Then
                    'ix = mobjGrupo.AgregarUsuario(.List(i))
                    ix = lstF.ItemData(j)
                    lstD.AddItem lstF.List(j)
                    lstD.ItemData(lstD.NewIndex) = ix
                    lstF.RemoveItem j
                End If
            Next j
         Next i
    End If
End Sub


Private Function ArchivoB() As String
    Dim i As Integer
    On Error GoTo mensaje
    
    With dlg1b
        .DialogTitle = "Abrir"
        .CancelError = True
'        .InitDir = gobjRol.EmpresaActual.Ruta
        .flags = cdlOFNFileMustExist
        .DefaultExt = "txt"
        .Filter = "Archivos Texto|*.txt"
        .filename = gConfigura.ArchivoB
        .ShowOpen
        ArchivoB = .filename
        gConfigura.ArchivoB = ArchivoB
    End With
    Exit Function
    
mensaje:
    MsgBox "Se ha selccionado Cancelar"
    If Len(txtArchiOrigenB.Text) > 0 Then ArchivoB = txtArchiOrigenB.Text
    Exit Function
End Function

Private Sub grdComisionesb_BeforeEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
    grdComisionesB.EditMaxLength = 12 'Hasta 99,999,999,999
End Sub

Private Sub grdComisionesb_KeyPressEdit(ByVal Row As Long, ByVal col As Long, KeyAscii As Integer)
    'Acepta sólo númericos
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And (KeyAscii <> vbKeyBack) And (KeyAscii <> Asc(".")) Then
        KeyAscii = 0
    End If
End Sub

Private Sub GrabaIntervalosB()
    Dim i As Integer
    With grdComisionesB
        For i = .FixedRows To .Rows - 1
            gComisionesB(i).desde = .ValueMatrix(i, 0)
            gComisionesB(i).hasta = .ValueMatrix(i, 1)
            gComisionesB(i).Comision = .ValueMatrix(i, 2)
            gComisionesB(i).ComisionC = .ValueMatrix(i, 3)
            gComisionesB(i).ComisionSC = .ValueMatrix(i, 4)
        Next i
    End With
'    EscribirIntervalosB
    EscribirIntervalosGnOpcion "B"
End Sub

Private Sub CargaTablaComisionesB()
    Dim i As Integer
    ConfigColsComisiones
    For i = 1 To 10
        grdComisionesB.TextMatrix(i, 0) = gComisionesB(i).desde
        grdComisionesB.TextMatrix(i, 1) = gComisionesB(i).hasta
        grdComisionesB.TextMatrix(i, 2) = gComisionesB(i).Comision
        grdComisionesB.TextMatrix(i, 3) = gComisionesB(i).ComisionC
        grdComisionesB.TextMatrix(i, 4) = gComisionesB(i).ComisionSC
    Next i
    ConfigColsComisiones
End Sub

Private Sub cargaVendedores(lst As ListBox, Num As Integer)
    Dim rs As Recordset, sql As String

    sql = "SELECT CodVendedor,Nombre FROM FCVendedor "
    sql = sql & "WHERE BandValida=1 and BandVendedor=1 "
    sql = sql & " and tipoTabla=" & Num
    sql = sql & " or tipoTabla is null "
    sql = sql & " ORDER BY CodVendedor"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
   
    With rs
        If Not (.EOF) Then
            .MoveFirst
            Do Until .EOF
                lst.AddItem !CodVendedor & "  " & !nombre
                lst.ItemData(lst.NewIndex) = Len(!CodVendedor)
               .MoveNext
           Loop
        End If
    End With
    rs.Close
End Sub


Private Sub GrabarVendedores(lst As ListBox, valor As Integer)
    Dim Cadena As String, i As Integer, NumReg As Long
    Dim sql As String, rs As Recordset
    Cadena = ""
    
'            sql = "update fcvendedor set TipoTabla=0 "
'            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
    
    
    For i = 0 To lst.ListCount - 1
            sql = "update fcvendedor set TipoTabla=" & valor
            sql = sql & " where codvendedor='" & Left(lst.List(i), lst.ItemData(i)) & "'"
            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
    Next i
End Sub


Private Sub Regresa(lst As ListBox, lst1 As ListBox)
    Dim i As Long, ix As Long
    On Error GoTo ErrTrap
    With lst
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                ix = .ItemData(i)
                lst1.AddItem .List(i)
                lst1.ItemData(lst1.NewIndex) = ix
                .RemoveItem i
            End If
        Next i
    End With
    Exit Sub
ErrTrap:
    DispErr
End Sub

Private Sub lstVendedores_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim DY   ' Declara variable.
   DY = TextHeight("A")   ' Obtiene el alto de una línea.
   lblDragDrop.Move lstVendedores.Left, lstVendedores.Top + y - DY / 2, lstVendedores.Width, DY
   lblDragDrop.Drag   ' Ar
   lblDragDrop.Caption = lstVendedores.Text
End Sub

Private Sub Form_DragOver(Source As Control, x As Single, y As Single, State As Integer)
   ' Cambia el puntero a no colocar.
   If State = 0 Then Source.MousePointer = 12
   ' Utiliza el puntero predeterminado del mouse.
   If State = 1 Then Source.MousePointer = 0
End Sub

Private Sub lstA_DragDrop(Source As Control, x As Single, y As Single)
   On Error Resume Next
   carga lstA, lstVendedores
End Sub

Private Sub lstB_DragDrop(Source As Control, x As Single, y As Single)
   On Error Resume Next
   carga lstB, lstVendedores
End Sub

Private Sub lstC_DragDrop(Source As Control, x As Single, y As Single)
   On Error Resume Next
   carga lstC, lstVendedores
End Sub

Private Sub lstD_DragDrop(Source As Control, x As Single, y As Single)
   On Error Resume Next
   carga lstD, lstVendedores
End Sub


Private Sub carga(lst As ListBox, lst1 As ListBox)
    Dim i As Long, ix As Long
    With lst1
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                ix = .ItemData(i)
                lst.AddItem .List(i)
                lst.ItemData(lst.NewIndex) = ix
                .RemoveItem i
            End If
        Next i
    End With
End Sub



Public Sub RecuperaSeleccion(ByVal Key As String, lst As ListBox, lst1 As ListBox, Optional s As String)
Dim Vector As Variant
Dim i As Integer, j As Integer, Selec As Integer, ix As Long, max As Integer, pos As Integer
Dim trans As String
    If s <> "_VACIO_" Then
        With lst1
            Vector = Split(s, ",")
             Selec = UBound(Vector, 1)
             For i = 0 To Selec
                max = .ListCount - 1
                j = 0
                For j = 0 To max
                    pos = InStr(1, .List(j), " ")
                    'If Vector(i) = Left(.List(j), .ItemData(j)) Then
                    trans = Trim$(Mid$(.List(j), 1, pos - 1))
                    If Vector(i) = trans Then
                        ix = .ItemData(i)
                        lst.AddItem .List(j)
                        lst.ItemData(lst.NewIndex) = ix
                        .RemoveItem j
                        j = max
                    End If
                Next j
             Next i
        End With
    End If
End Sub

Private Function PreparaCadena1(lst As ListBox) As String
    Dim Cadena As String, i As Integer, pos As Integer
    Cadena = ""
    For i = 0 To lst.ListCount - 1
        'If lst.Selected(i) Then
            If Cadena = "" Then
                pos = InStr(1, lst.List(i), " ")
                'cadena = Left(lst.List(i), lst.ItemData(i))
                Cadena = Trim(Mid$(lst.List(i), 1, pos))
            Else
                'cadena = cadena & "," & _
                              Left(lst.List(i), lst.ItemData(i))
                Cadena = Cadena & "," & Trim(Mid$(lst.List(i), 1, pos))
            End If
        'End If
    Next i
    PreparaCadena1 = Cadena
End Function


Private Sub GrabaIntervalosC()
    Dim i As Integer
    With grdComisionesC
        For i = .FixedRows To .Rows - 1
            gComisionesC(i).desde = .ValueMatrix(i, 0) ' CCur(Format("0" & .TextMatrix(i, 0), mobjGNComp.FormatoMoneda))
            gComisionesC(i).hasta = .ValueMatrix(i, 1) 'CCur(Format("0" & .TextMatrix(i, 1), mobjGNComp.FormatoMoneda))
            gComisionesC(i).Comision = .ValueMatrix(i, 2) ' CCur(Format("0" & .TextMatrix(i, 2), mobjGNComp.FormatoMoneda))
            gComisionesC(i).ComisionC = .ValueMatrix(i, 3)
            gComisionesC(i).ComisionSC = .ValueMatrix(i, 4)
        Next i
    End With
'    EscribirIntervalosC
    EscribirIntervalosGnOpcion "C"
End Sub

Private Sub GrabaIntervalosD()
    Dim i As Integer
    With grdComisionesD
        For i = .FixedRows To .Rows - 1
            gComisionesD(i).desde = .ValueMatrix(i, 0) ' CCur(Format("0" & .TextMatrix(i, 0), mobjGNComp.FormatoMoneda))
            gComisionesD(i).hasta = .ValueMatrix(i, 1) 'CCur(Format("0" & .TextMatrix(i, 1), mobjGNComp.FormatoMoneda))
            gComisionesD(i).Comision = .ValueMatrix(i, 2) ' CCur(Format("0" & .TextMatrix(i, 2), mobjGNComp.FormatoMoneda))
            gComisionesD(i).ComisionC = .ValueMatrix(i, 3)
            gComisionesD(i).ComisionSC = .ValueMatrix(i, 4)
        Next i
    End With
'    EscribirIntervalosD
    EscribirIntervalosGnOpcion "D"
End Sub


Private Sub CargaTablaComisionesC()
    Dim i As Integer
    ConfigColsComisiones
    For i = 1 To 10
        grdComisionesC.TextMatrix(i, 0) = gComisionesC(i).desde
        grdComisionesC.TextMatrix(i, 1) = gComisionesC(i).hasta
        grdComisionesC.TextMatrix(i, 2) = gComisionesC(i).Comision
        grdComisionesC.TextMatrix(i, 3) = gComisionesC(i).ComisionC
        grdComisionesC.TextMatrix(i, 4) = gComisionesC(i).ComisionSC
    Next i
    ConfigColsComisiones
End Sub

Private Sub CargaTablaComisionesD()
    Dim i As Integer
    ConfigColsComisiones
    For i = 1 To 10
        grdComisionesD.TextMatrix(i, 0) = gComisionesD(i).desde
        grdComisionesD.TextMatrix(i, 1) = gComisionesD(i).hasta
        grdComisionesD.TextMatrix(i, 2) = gComisionesD(i).Comision
        grdComisionesD.TextMatrix(i, 3) = gComisionesD(i).ComisionC
        grdComisionesD.TextMatrix(i, 4) = gComisionesD(i).ComisionSC
    Next i
    ConfigColsComisiones
End Sub


Private Sub ActualizaVendedor()
    Dim Cadena As String, i As Integer, NumReg As Long
    Dim sql As String, rs As Recordset
    Cadena = ""
    sql = "update fcvendedor set TipoTabla=0 where bandvendedor=1 and bandcobrador=0 "
    gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
End Sub

Private Sub ntxPorcenCobrador_Change()
    gConfigura.PorcenCobradorA = ntxPorcenCobrador.value
End Sub

Private Sub ntxPorcenCobradorb_Change()
    gConfigura.PorcenCobradorB = ntxPorcenCobradorB.value
End Sub

Private Sub ntxPorcenCobradorc_Change()
    gConfigura.PorcenCobradorC = ntxPorcenCobradorC.value
End Sub

Private Sub ntxPorcenCobradord_Change()
    gConfigura.PorcenCobradorD = ntxPorcenCobradorD.value
End Sub



Private Sub ntxPorcenVendedor_Change()
    gConfigura.PorcenVendedorA = ntxPorcenVendedor.value
End Sub

Private Sub ntxPorcenVendedorb_Change()
    gConfigura.PorcenVendedorB = ntxPorcenVendedorB.value
End Sub

Private Sub ntxPorcenVendedorc_Change()
    gConfigura.PorcenVendedorC = ntxPorcenVendedorC.value
End Sub

Private Sub ntxPorcenVendedord_Change()
    gConfigura.PorcenVendedorD = ntxPorcenVendedorD.value
End Sub


Private Sub ntxPorcenVendedor_LostFocus()
    ntxPorcenCobrador.value = 100 - ntxPorcenVendedor.value
End Sub

Private Sub ntxPorcenCobrador_LostFocus()
    ntxPorcenVendedor.value = 100 - ntxPorcenCobrador.value
End Sub

Private Function ArchivoC() As String
    Dim i As Integer
    On Error GoTo mensaje
    
    With dlg1b
        .DialogTitle = "Abrir"
        .CancelError = True
'        .InitDir = gobjRol.EmpresaActual.Ruta
        .flags = cdlOFNFileMustExist
        .DefaultExt = "txt"
        .Filter = "Archivos Texto|*.txt"
        .filename = gConfigura.ArchivoB
        .ShowOpen
        ArchivoC = .filename
        gConfigura.ArchivoC = ArchivoC
    End With
    Exit Function
    
mensaje:
    MsgBox "Se ha selccionado Cancelar"
    If Len(txtArchiOrigenC.Text) > 0 Then ArchivoC = txtArchiOrigenC.Text
    Exit Function
End Function

Private Function ArchivoD() As String
    Dim i As Integer
    On Error GoTo mensaje
    
    With dlg1b
        .DialogTitle = "Abrir"
        .CancelError = True
'        .InitDir = gobjRol.EmpresaActual.Ruta
        .flags = cdlOFNFileMustExist
        .DefaultExt = "txt"
        .Filter = "Archivos Texto|*.txt"
        .filename = gConfigura.ArchivoD
        .ShowOpen
        ArchivoD = .filename
        gConfigura.ArchivoD = ArchivoD
    End With
    Exit Function
    
mensaje:
    MsgBox "Se ha selccionado Cancelar"
    If Len(txtArchiOrigenD.Text) > 0 Then ArchivoD = txtArchiOrigenD.Text
    Exit Function
End Function


Public Function InicioComPenVenJefe(ByRef objcond As Condicion, _
                                                    Optional ByRef T1 As String, Optional ByRef T2 As String, _
                                                    Optional ByRef T3 As String, Optional ByRef T4 As String) As Boolean
                           
    Dim i As Integer, KeyT As String, KeyTC As String, pos As Integer
    Dim KeyRecargo As String, KeyVenta As String
    Dim keyDias As String, keyComision As String
    Dim keyPenal As String, keyTipoCom As String
    SSTab1.Tab = 0
    RecuperarConfig
''    LeerIntervalosJefe "A"
''    LeerIntervalosJefe "B"
''    LeerIntervalosJefe "C"
''    LeerIntervalosJefe "D"
    
    LeerIntervaloJefesGnOpcion "A"
    LeerIntervaloJefesGnOpcion "B"
    LeerIntervaloJefesGnOpcion "C"
    LeerIntervaloJefesGnOpcion "D"
    
    
    Dim trans As String
    Me.tag = Name  'nombre del reporte
    'Prepara la lista de monedas
    CargaTablaComisionesJefe "A"
    CargaTablaComisionesJefe "B"
    CargaTablaComisionesJefe "C"
    CargaTablaComisionesJefe "D"
    txtArchiOrigen.Text = gConfiguraJefe.ArchivoJefeA
    txtArchiOrigenB.Text = gConfiguraJefe.ArchivoJefeB
    txtArchiOrigenC.Text = gConfiguraJefe.ArchivoJefeC
    txtArchiOrigenD.Text = gConfiguraJefe.ArchivoJefeD
    
    CargaRecargo
    CargaJefeVENDEDORES lstVendedores, 0
    CargaJefeVENDEDORES lstA, 1
    CargaJefeVENDEDORES lstB, 2
    CargaJefeVENDEDORES lstC, 3
    CargaJefeVENDEDORES lstD, 4
    With objcond
        dtpDesde.value = IIf(.fecha1 = 0, dtpDesde.value, dtpDesde.value)
        dtpHasta.value = IIf(.fecha2 = 0, dtpHasta.value, dtpHasta.value)
        CargaTipoTrans "", lst
        'Prepara la lista de vendedores que irá en el combo
        fcbDesde1.SetData gobjMain.EmpresaActual.ListaFCVendedorN(False, False, True, False)
        fcbHasta1.SetData gobjMain.EmpresaActual.ListaFCVendedorN(False, False, True, False)
        BandAceptado = False
        'Valores predeterminados
        fcbDesde1.KeyText = .CodCentro1 '  Vendedor1
        fcbHasta1.KeyText = .CodCentro2  '.Vendedor2
        trans = GetSetting(APPNAME, App.Title, "Trans_Comisiones", "_VACIO_")
        RecuperaSelec KeyT, lst, trans
        trans = GetSetting(APPNAME, App.Title, "RecarDesc_Comisiones", "_VACIO_")
        RecuperaSelecRec KeyRecargo, lstFuente, lstDestino, trans
        
        pos = InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "HORMI")
        If pos > 0 Then
            Label7.Caption = "QUITO"
            Label4.Caption = "GUAYAQUIL"
            Label9.Caption = "CUENCA"
            Label10.Caption = "OTROS"
        End If
        
'        ntxPorcenVendedor
        Me.Show vbModal, frmMain
        'Si aplastó el botón 'Aceptar'
        If BandAceptado Then
            'Devuelve los valores de condición para la búsqueda
            .CodCentro1 = Trim$(fcbDesde1.Text)
            .CodCentro2 = Trim$(fcbHasta1.Text)
            .fecha1 = dtpDesde.value
            .fecha2 = dtpHasta.value
            .CodTrans = PreparaCadena(lst) ' .TipoTrans
            .Servicios = PreparaCadRec(lstDestino)
            T1 = PreparaCadRec(lstA)
            T2 = PreparaCadRec(lstB)
            T3 = PreparaCadRec(lstC)
            T4 = PreparaCadRec(lstD)
            GrabaConfig
            GrabaIntervalosJefe "A"
            GrabaIntervalosJefe "B"
            GrabaIntervalosJefe "C"
            GrabaIntervalosJefe "D"
            ActualizaVendedorJefe ' pone a todos en tipotabla=0
            GrabarVendedores lstA, 1
            GrabarVendedores lstB, 2
            GrabarVendedores lstC, 3
            GrabarVendedores lstD, 4
            SaveSetting APPNAME, App.Title, "Trans_Comisiones", .CodTrans
            SaveSetting APPNAME, App.Title, "RecarDesc_Comisiones", .Servicios
        End If
    End With
    'Devuelve true/false
    Unload Me
    InicioComPenVenJefe = BandAceptado
End Function


Private Sub CargaTablaComisionesJefe(ByVal TipoTabla As String)
    Dim i As Integer
    ConfigColsComisiones
'    LeerIntervalos
        Select Case TipoTabla
            Case "A":
                For i = 1 To 10
                    grdComisiones.TextMatrix(i, 0) = gComisionesJefe(i).desde
                    grdComisiones.TextMatrix(i, 1) = gComisionesJefe(i).hasta
                    grdComisiones.TextMatrix(i, 2) = gComisionesJefe(i).Comision
                    grdComisiones.TextMatrix(i, 3) = gComisionesJefe(i).ComisionC
                    grdComisiones.TextMatrix(i, 4) = gComisionesJefe(i).ComisionSC
                Next i
            Case "B":
                For i = 1 To 10
                    grdComisionesB.TextMatrix(i, 0) = gComisionesJefeB(i).desde
                    grdComisionesB.TextMatrix(i, 1) = gComisionesJefeB(i).hasta
                    grdComisionesB.TextMatrix(i, 2) = gComisionesJefeB(i).Comision
                    grdComisionesB.TextMatrix(i, 3) = gComisionesJefeB(i).ComisionC
                    grdComisionesB.TextMatrix(i, 4) = gComisionesJefeB(i).ComisionSC
                Next i
            Case "C":
                For i = 1 To 10
                    grdComisionesC.TextMatrix(i, 0) = gComisionesJefeC(i).desde
                    grdComisionesC.TextMatrix(i, 1) = gComisionesJefeC(i).hasta
                    grdComisionesC.TextMatrix(i, 2) = gComisionesJefeC(i).Comision
                    grdComisionesC.TextMatrix(i, 3) = gComisionesJefeC(i).ComisionC
                    grdComisionesC.TextMatrix(i, 4) = gComisionesJefeC(i).ComisionSC
                Next i
            Case "D":
                For i = 1 To 10
                    grdComisionesD.TextMatrix(i, 0) = gComisionesJefeD(i).desde
                    grdComisionesD.TextMatrix(i, 1) = gComisionesJefeD(i).hasta
                    grdComisionesD.TextMatrix(i, 2) = gComisionesJefeD(i).Comision
                    grdComisionesD.TextMatrix(i, 3) = gComisionesJefeD(i).ComisionC
                    grdComisionesD.TextMatrix(i, 4) = gComisionesJefeD(i).ComisionSC
                Next i
        End Select
    ConfigColsComisiones
End Sub


Private Sub GrabaIntervalosJefe(ByVal TipoTabla As String)
    Dim i As Integer
        Select Case TipoTabla
        Case "A":
            With grdComisiones
                For i = .FixedRows To .Rows - 1
                    gComisionesJefe(i).desde = .ValueMatrix(i, 0)
                    gComisionesJefe(i).hasta = .ValueMatrix(i, 1)
                    gComisionesJefe(i).Comision = .ValueMatrix(i, 2)
                    gComisionesJefe(i).ComisionC = .ValueMatrix(i, 3)
                    gComisionesJefe(i).ComisionSC = .ValueMatrix(i, 4)
                Next i
            End With
        Case "B":
            With grdComisionesB
                For i = .FixedRows To .Rows - 1
                    gComisionesJefeB(i).desde = .ValueMatrix(i, 0)
                    gComisionesJefeB(i).hasta = .ValueMatrix(i, 1)
                    gComisionesJefeB(i).Comision = .ValueMatrix(i, 2)
                    gComisionesJefeB(i).ComisionC = .ValueMatrix(i, 3)
                    gComisionesJefeB(i).ComisionSC = .ValueMatrix(i, 4)
                Next i
            End With
        Case "C":
            With grdComisionesC
                For i = .FixedRows To .Rows - 1
                    gComisionesJefeC(i).desde = .ValueMatrix(i, 0)
                    gComisionesJefeC(i).hasta = .ValueMatrix(i, 1)
                    gComisionesJefeC(i).Comision = .ValueMatrix(i, 2)
                    gComisionesJefeC(i).ComisionC = .ValueMatrix(i, 3)
                    gComisionesJefeC(i).ComisionSC = .ValueMatrix(i, 4)
                Next i
            End With
        Case "D":
            With grdComisionesD
                For i = .FixedRows To .Rows - 1
                    gComisionesJefeD(i).desde = .ValueMatrix(i, 0)
                    gComisionesJefeD(i).hasta = .ValueMatrix(i, 1)
                    gComisionesJefeD(i).Comision = .ValueMatrix(i, 2)
                    gComisionesJefeD(i).ComisionC = .ValueMatrix(i, 3)
                    gComisionesJefeD(i).ComisionSC = .ValueMatrix(i, 4)
                Next i
            End With
        End Select
    
'    EscribirIntervalosJefe TipoTabla
    EscribirIntervalosJefeGnOpcion TipoTabla
End Sub

Private Sub CargaJefeVENDEDORES(lst As ListBox, Num As Integer)
    Dim rs As Recordset, sql As String

    sql = "SELECT CodVendedor,Nombre FROM FCVendedor "
    sql = sql & "WHERE BandValida=1 and BandVendedor=1 and BandCobrador=1"
'    If Num <> 0 Then
        sql = sql & " and tipoTabla=" & Num
        sql = sql & " or tipoTabla is null "
 '   End If
    sql = sql & " ORDER BY CodVendedor"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
   
    With rs
        If Not (.EOF) Then
            .MoveFirst
            Do Until .EOF
                lst.AddItem !CodVendedor & "  " & !nombre
                lst.ItemData(lst.NewIndex) = Len(!CodVendedor)
               .MoveNext
           Loop
        End If
    End With
    rs.Close
End Sub

Private Sub ActualizaVendedorJefe()
    Dim Cadena As String, i As Integer, NumReg As Long
    Dim sql As String, rs As Recordset
    Cadena = ""
    sql = "update fcvendedor set TipoTabla=0 where bandvendedor=1 and bandcobrador=1 "
    gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
End Sub

