VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{1B04A20A-C295-476C-BA28-DC6D9110E7A3}#1.0#0"; "vspdf.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "flexcombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmGeneraNotificacion 
   Caption         =   "Generar Notificacaiones"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12585
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin VB.Frame fraPadre 
      Caption         =   "Padres / Alumos"
      Height          =   1215
      Left            =   16200
      TabIndex        =   39
      Top             =   60
      Visible         =   0   'False
      Width           =   3075
      Begin FlexComboProy.FlexCombo fcbCli 
         Height          =   315
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
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
      Begin FlexComboProy.FlexCombo fcbGar 
         Height          =   315
         Left            =   120
         TabIndex        =   41
         Top             =   780
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
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
   Begin VB.Frame FraGrupo 
      Caption         =   "Cursos/Bustas"
      Height          =   1215
      Left            =   13740
      TabIndex        =   36
      Top             =   60
      Visible         =   0   'False
      Width           =   2415
      Begin FlexComboProy.FlexCombo fcbGrupo 
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
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
      Begin FlexComboProy.FlexCombo fcbChofer 
         Height          =   315
         Left            =   120
         TabIndex        =   38
         Top             =   780
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   556
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
   Begin VB.Frame FraFecha 
      Caption         =   "Fecha Maxima de Pago"
      Height          =   1215
      Left            =   11100
      TabIndex        =   34
      Top             =   60
      Visible         =   0   'False
      Width           =   2595
      Begin MSComCtl2.DTPicker dtpFechaPago 
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   661
         _Version        =   393216
         Format          =   106299393
         CurrentDate     =   42333
      End
      Begin MSComCtl2.DTPicker dtpFechaCorte1 
         Height          =   375
         Left            =   120
         TabIndex        =   42
         Top             =   780
         Visible         =   0   'False
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   661
         _Version        =   393216
         Format          =   106299393
         CurrentDate     =   42333
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha de Corte"
         Height          =   375
         Left            =   120
         TabIndex        =   43
         Top             =   580
         Visible         =   0   'False
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdBuscar1 
      Caption         =   "&Buscar"
      Height          =   372
      Left            =   300
      TabIndex        =   23
      Top             =   1440
      Width           =   1452
   End
   Begin VB.CommandButton cmdTransLimpiar 
      Caption         =   "Limpiar"
      Height          =   405
      Left            =   1920
      TabIndex        =   10
      Top             =   1440
      Width           =   732
   End
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   852
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   20160
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   10080
      Width           =   20160
      Begin VB.CommandButton cmdImprimir2 
         Caption         =   "Imprimir "
         Height          =   372
         Left            =   2880
         TabIndex        =   33
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton BTNOPEN 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   13800
         TabIndex        =   15
         Top             =   120
         Width           =   372
      End
      Begin VB.TextBox txtPlantilla 
         Height          =   300
         Left            =   9480
         TabIndex        =   13
         Text            =   "txtPlantilla"
         Top             =   120
         Width           =   4335
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir "
         Height          =   372
         Left            =   2880
         TabIndex        =   12
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdGenNoti 
         Caption         =   "Grabar Notificacion"
         Height          =   372
         Left            =   840
         TabIndex        =   11
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdCancelar 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   6600
         TabIndex        =   4
         Top             =   120
         Width           =   1695
      End
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   540
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
      Begin VSPDFLibCtl.VSPDF pdf 
         Left            =   120
         OleObjectBlob   =   "GeneraNotificaciones.frx":0000
         Top             =   240
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Plantilla"
         Height          =   195
         Left            =   8880
         TabIndex        =   14
         Top             =   120
         Width           =   540
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   1920
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
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   372
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   1452
   End
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   10680
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpFechaCorte 
      Height          =   300
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
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
      Format          =   106430465
      CurrentDate     =   36348
   End
   Begin VB.Frame fraCodTrans 
      Caption         =   "s"
      Height          =   1215
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.ListBox lstTrans 
         Columns         =   6
         Height          =   852
         IntegralHeight  =   0   'False
         Left            =   120
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame frmForma 
      Caption         =   "Formas de Cobro"
      Height          =   1215
      Left            =   5400
      TabIndex        =   8
      Top             =   120
      Width           =   3015
      Begin VB.ListBox lstForma 
         Columns         =   3
         Height          =   855
         IntegralHeight  =   0   'False
         Left            =   120
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.PictureBox Pic 
      BorderStyle     =   0  'None
      Height          =   1395
      Left            =   0
      ScaleHeight     =   1395
      ScaleWidth      =   14055
      TabIndex        =   24
      Top             =   -120
      Width           =   14055
      Begin VB.OptionButton Opt 
         Caption         =   "Notificacion3"
         Height          =   255
         Index           =   2
         Left            =   4020
         TabIndex        =   27
         Top             =   1020
         Width           =   1275
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Notificacion2"
         Height          =   255
         Index           =   1
         Left            =   2400
         TabIndex        =   26
         Top             =   1020
         Width           =   1275
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Notificacion1"
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   25
         Top             =   1020
         Width           =   1275
      End
      Begin VB.Frame fraCli 
         Caption         =   "Datos de cliente"
         Height          =   855
         Left            =   1080
         TabIndex        =   28
         Top             =   60
         Width           =   6495
         Begin FlexComboProy.FlexCombo fcbCliente2 
            Height          =   375
            Left            =   4200
            TabIndex        =   31
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
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
         Begin FlexComboProy.FlexCombo fcbCliente1 
            Height          =   375
            Left            =   1020
            TabIndex        =   32
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
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
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Left            =   3420
            TabIndex        =   30
            Top             =   300
            Width           =   420
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Left            =   60
            TabIndex        =   29
            Top             =   300
            Width           =   465
         End
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "3∫ Notificacion Realizada"
      Height          =   195
      Left            =   9240
      TabIndex        =   22
      Top             =   960
      Width           =   1785
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8640
      TabIndex        =   21
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "2∫ Notificacion Realizada"
      Height          =   195
      Left            =   9240
      TabIndex        =   20
      Top             =   600
      Width           =   1785
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8640
      TabIndex        =   19
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "1∫ Notificacion Realizada"
      Height          =   195
      Left            =   9240
      TabIndex        =   18
      Top             =   240
      Width           =   1785
   End
   Begin VB.Label lblColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8640
      TabIndex        =   17
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Fecha de Corte  "
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1185
   End
End
Attribute VB_Name = "frmGeneraNotificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Constantes para las columnas
Private Const COL_NUMFILA = 0
Private Const COL_TID = 1
Private Const COL_CODPROVCLI = 2
Private Const COL_NOMBRE = 3
Private Const COL_TRANS = 4
Private Const COL_NUMDOCREF = 5
Private Const COL_FEMISION = 6
Private Const COL_FVENCI = 7
Private Const COL_DIASVENCI = 8
Private Const COL_VALOR = 9
Private Const COL_SALDO = 10
Private Const COL_BANDNOTI1 = 11
Private Const COL_FECHANOTI1 = 12
Private Const COL_BANDNOTI2 = 13
Private Const COL_FECHANOTI2 = 14
Private Const COL_BANDNOTI3 = 15
Private Const COL_FECHANOTI3 = 16
Private Const COL_IMP = 17
Private Const COL_SEC = 18
Private Const MSG_EX = "Ya Existe."
Private mProcesando As Boolean
Private mCancelado As Boolean
Private mColItems As Collection

Private mobjImp As Object

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Type OPENFILENAME
  lStructSize As Long
  hwndOwner As Long
  hInstance As Long
  lpstrFilter As String
  lpstrCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  lpstrFile As String
  nMaxFile As Long
  lpstrFileTitle As String
  nMaxFileTitle As Long
  lpstrInitialDir As String
  lpstrTitle As String
  flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type
  
Private Const OFN_READONLY = &H1
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_SHOWHELP = &H10
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOLONGNAMES = &H40000 ' force no long names for 4.x modules
Private Const OFN_EXPLORER = &H80000 ' new look commdlg
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_LONGNAMES = &H200000 ' force long names for 3.x modules
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Dim Position As Long
Dim pageNo As Long
Dim lineNo As Long
Dim pageHeight As Long
Dim pageWidth As Long
Dim location(1 To 5000) As Long
Dim pageObj(1 To 5000) As Long
Dim lines As Long
Dim obj As Long
Dim Tpages As Long
Dim encoding As Long
Dim resources As Long
Dim pages As Variant
Dim author As String
Dim creator As String
Dim keywords As String
Dim subject As String
Dim Title As String
Dim BaseFont As String
Dim pointSize As Currency
Dim vertSpace As Currency
Dim rotate As Integer
Dim info As Long
Dim root As Long
Dim npagex As Double
Dim npagey As Long
Dim filetxt As String
Dim filepdf As String
Dim linelen As Long
Dim cache As String
Dim cmdline As String
Dim CREADOR As String
Const APPNAME = "Text-PDF v1.0"
Dim Vector As String

Public Sub Inicio()
    Dim i As Integer
    On Error GoTo ErrTrap
    Me.Show
    Me.ZOrder
    BandHabilita False
    cmdBuscar1.Visible = False
    cmdBuscar.Visible = True
    Pic.Visible = False
    cmdImprimir2.Visible = False
    cmdImprimir.Visible = True
    cmdGenNoti.Visible = True
    dtpFechaCorte.value = Date
    grd.Rows = 1
    CargaTrans
    CargaTodasFormasCobroPago lstForma
    recuperaConfiguracion
    
    If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "CATA") <> 0 Then
        FraFecha.Visible = True
        dtpFechaPago.value = Date
        dtpFechaCorte1.value = Date
        fcbGrupo.SetData gobjMain.EmpresaActual.ListaPCGrupo(1, True, False)
        fcbChofer.SetData gobjMain.EmpresaActual.ListaFCVendedor(True, False)
        fcbCli.SetData gobjMain.EmpresaActual.ListaPCProvCli(False, True, False)
        fcbGar.SetData gobjMain.EmpresaActual.ListaPCGar(False)
        dtpFechaCorte1.Visible = True
        Label10.Visible = True
        FraGrupo.Visible = True
        fraPadre.Visible = True
            
    End If
    
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub BandHabilita(ByVal Estado As Boolean)
    cmdGenNoti.Enabled = Estado
    cmdImprimir.Enabled = Not Estado
    cmdCancelar.Enabled = Not Estado
End Sub
Private Sub CargaTrans()
    Dim i As Long, v As Variant
    Dim s As String
        lstTrans.Clear
        'v = gobjMain.GrupoActual.PermisoActual.ListaTransxTipoTransNew(False, "", "2", "18")
        v = gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "")
        For i = LBound(v, 2) To UBound(v, 2)
            lstTrans.AddItem v(0, i)        '& " " & v(1, i)
        Next i
            s = GetSetting(APPNAME, App.Title, "KeyTNoti", "")
            RecuperaTrans "KeyTNoti", lstTrans, s
End Sub

Private Sub cmdArchivo_Click()
Dim arhivoPlantillanNoti As String
On Error GoTo ErrTrap
    With dlg1
        .InitDir = App.Path
        .CancelError = True
        .Filter = "Texto (Separado por coma *.txt)|*.csv|Texto (Separado por tabuladores *.cvs)|*.txt|Todos *.*|*.*"
        .flags = cdlOFNFileMustExist
        .ShowOpen
        txtPlantilla.Text = .filename
        arhivoPlantillanNoti = .filename
    End With
    GrabaConfiguracion
    Exit Sub
ErrTrap:
End Sub

Private Sub cmdBuscar_Click()
grd.Rows = 1
    With gobjMain.objCondicion
        .FechaCorte = dtpFechaCorte.value
        .CodTrans = PreparaCadena(lstTrans)
        .codforma = PreparaCadena(lstForma)
        If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "CATA") <> 0 Then
            BuscarEstudiantes
            BorrarNoti
            PonerColor
            ConfigColsEstudiante
'            GeneraNotificaciones
'            AsignaSecuencial
            BandHabilita False
            
        Else
            Buscar
            BorrarNoti
            PonerColor
            ConfigCols
            GeneraNotificaciones
            AsignaSecuencial
            BandHabilita True
            
        End If
        
        SaveSetting APPNAME, App.Title, "KeyFormaNoti", .codforma
        SaveSetting APPNAME, App.Title, "KeyTNoti", .CodTrans
    End With
End Sub
Private Sub PonerColor()
Dim i As Long
For i = 1 To grd.Rows - 1
    If grd.ValueMatrix(i, COL_BANDNOTI1) = -1 Then
        grd.Cell(flexcpBackColor, i, COL_BANDNOTI1, i, COL_FECHANOTI1) = &HC0FFFF
     Else
        grd.Cell(flexcpBackColor, i, COL_BANDNOTI1, i, COL_FECHANOTI1) = vbWhite
    End If
    If grd.ValueMatrix(i, COL_BANDNOTI2) = -1 Then 'segunda notificacion
        grd.Cell(flexcpBackColor, i, COL_BANDNOTI2, i, COL_FECHANOTI2) = &HC0FFC0
    Else
        grd.Cell(flexcpBackColor, i, COL_BANDNOTI2, i, COL_FECHANOTI2) = vbWhite
    End If
    If grd.ValueMatrix(i, COL_BANDNOTI3) = -1 Then 'tercera notificacion
        grd.Cell(flexcpBackColor, i, COL_BANDNOTI3, i, COL_FECHANOTI3) = &HFFC0FF
    Else
        grd.Cell(flexcpBackColor, i, COL_BANDNOTI3, i, COL_FECHANOTI3) = vbWhite
    End If
Next
End Sub
Private Sub BorrarNoti()
Dim i As Long, j As Long
For i = 1 To grd.Rows - 1
    For j = COL_FECHANOTI1 To grd.Cols - 1
        If grd.TextMatrix(i, j) = "01/01/1900" Then
            grd.TextMatrix(i, j) = ""
        End If
    Next
    grd.Cell(flexcpBackColor, i, COL_DIASVENCI, i, COL_DIASVENCI) = &HC0C0FF
Next
End Sub
Private Sub Buscar()
Dim aux As String
Dim sql As String
Dim Condicion As String
MensajeStatus "Procesando", vbHourglass
With gobjMain.objCondicion
'1) Prepara los  documentos  Asignados  menores a la fecha
        VerificaExistenciaTabla 1
        'aux = IIf(.NumMoneda > 0, "/Cotizacion" & .NumMoneda + 1, "")
        sql = "SELECT " & _
            "pck.IdAsignado, " & _
            "(pck.Debe + pck.Haber)  AS Valor " & _
            "INTO tmp1 " & _
            "From " & _
            "GNtrans gt INNER JOIN " & _
                "(GNComprobante gc INNER JOIN PCKardex pck " & _
                "ON gc.transID = pck.transID) " & _
                          "ON gt.Codtrans = gc.Codtrans " & _
            "Where (pck.IdAsignado <> 0) " & _
            "AND (gc.Estado <> 3) " & _
            "AND (gt.AfectaSaldoPC=1) " & _
            "AND (gc.Fechatrans<= " & FechaYMD(.FechaCorte, gobjMain.EmpresaActual.TipoDB) & ")"
        gobjMain.EmpresaActual.EjecutarSQL sql, 1
        '2)Agrupa  estos  documentos por IdAsignado
        VerificaExistenciaTabla 2
        sql = "SELECT " & _
              "IdAsignado," & _
              "isnull(Sum(Valor),0) AS VCancelado " & _
              "INTO tmp2 " & _
              "FROM tmp1 " & _
              "GROUP BY IdAsignado"
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
        '3) Agrupa los documentos con su valor cancelado por ID
        VerificaExistenciaTabla 3
        sql = "SELECT " & _
                "pck.Id, " & _
                "pck.Debe + pck.Haber AS Valor, " & _
                "isnull(vw.VCancelado,0) AS VCancelado, " & _
                "(pck.Debe + pck.Haber) - isnull(vw.VCancelado,0)  AS Saldo " & _
                "INTO tmp3 " & _
            "FROM GNtrans INNER JOIN  GNComprobante gc INNER JOIN (tmp2 vw RIGHT JOIN PCKardex pck  ON vw.IdAsignado = pck.Id) " & _
            "ON gc.TransID = pck.TransID  ON  GNTrans.CodTrans = gc.CodTrans " & _
            "Where (pck.IdAsignado = 0) And (gc.Estado <> 3)  " & _
                    "AND (pck.debe >0) " & _
                    " AND ((GNtrans.AfectaSaldoPC) = " & CadenaBool(True, gobjMain.EmpresaActual.TipoDB) & ") "
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
        '4) Finalmente une  con el documento  Padre
       VerificaExistenciaTabla 4
        sql = " SELECT vwConsPCDocSaldo.Id, PCProvCli.IdProvCli, PCProvCli.CodProvCli, PCProvCli.Nombre," & _
                " GNComprobante.CodTrans, GNComprobante.NumTrans, " & _
                " GNComprobante.CodTrans + ' ' + CONVERT(varchar, NumTrans)    AS Trans, " & _
                " CodForma, CodForma + pckardex.NumLetra AS Doc, " & _
                " PCKardex.FechaEmision, PCKardex.FechaVenci, " & _
                "(PCKardex.Debe + PCKardex.Haber) " & aux & " AS Valor, " & _
                "(PCKardex.Debe + PCKardex.Haber) " & aux & " - IsNull(vwConsPCDocSaldo.VCancelado,0) AS Saldo, " & _
                " PCKardex.Observacion, GNComprobante.CodUsuarioAutoriza " & _
                "INTO tmp4 " & _
            "FROM PCProvCli INNER JOIN  " & _
            " (GNTrans INNER JOIN " & _
                " (GNComprobante  INNER JOIN " & _
                    "(TSFormaCobroPago INNER JOIN " & _
                       "(PCKardex left JOIN FcVendedor  FCV  on PCKardex.idvendedor= fcv.idvendedor INNER JOIN " & _
                    " tmp3  vwConsPCDocSaldo ON PCKardex.Id = vwConsPCDocSaldo.Id) " & _
                " ON TSFormaCobroPago.IdForma = PCKardex.IdForma) ON " & _
            " GNComprobante.TransID = PCKardex.TransID) ON " & _
          " GNTrans.CodTrans = GNComprobante.CodTrans) ON " & _
         " PCProvCli.IdProvCli = PCKardex.IdProvCli " & _
            "Where (PCKardex.IdAsignado = 0) And (GNComprobante.Estado <> 3) " & _
            "AND (GNComprobante.Fechatrans<=" & FechaYMD(.FechaCorte, gobjMain.EmpresaActual.TipoDB) & ")" & _
            "AND (PCKardex.Debe >0) " & _
            "AND GNCOMPROBANTE.CodTrans IN (" & .CodTrans & ")"
       gobjMain.EmpresaActual.EjecutarSQL sql, 1
       
       ' CONSULTAS PARA PCNOTIFICACION
            VerificaExistenciaTabla 5
            sql = "select * into tmp5 from pcknotificacion  where bandnoti = 1 "
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            VerificaExistenciaTabla 6
            sql = "select * into tmp6 from pcknotificacion  where bandnoti = 2 "
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            VerificaExistenciaTabla 7
            sql = "select * into tmp7 from pcknotificacion  where bandnoti = 3 "
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            VerificaExistenciaTabla 8
            sql = " select  t5.idpckardex,t5.fechanoti1 as fechanoti1,case when t5.bandnoti= 1 then -1 else 0 end as bandnoti1,"
            sql = sql & "t6.fechanoti1 as fechanoti2,case when t6.bandnoti= 2 then -1 else 0 end as bandnoti2,"
            sql = sql & "t7.fechanoti1 as fechanoti3,case when t7.bandnoti= 3 then -1 else 0 end as bandnoti3 into tmp8"
            sql = sql & " from tmp5 t5 full join tmp6 t6 on t6.idpckardex = t5.idpckardex full join tmp7 t7 on t7.idpckardex=t6.idpckardex"
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            
    '4  Consulta  final
        Condicion = " WHERE vwConsPCCartera.Saldo >0 "
        Condicion = Condicion & " AND vwConsPCCartera.FechaEmision <= " & FechaYMD(.FechaCorte, gobjMain.EmpresaActual.TipoDB)
        
        If Len(.CodTrans) > 0 Then
            Condicion = Condicion & " AND vwConsPCCartera.CodTrans IN (" & .CodTrans & ") "
        End If
            Condicion = Condicion & " AND BandCliente=1"
        If Len(gobjMain.objCondicion.codforma) <> 0 Then
            Condicion = Condicion & " AND vwConsPCCartera.CodForma in ( " & .codforma & ")"
        End If
        sql = "SELECT vwConsPCCartera.id,vwConsPCCartera.CodProvCli, vwConsPCCartera.Nombre," & _
            " vwConsPCCartera.trans, " & _
            " vwConsPCCartera.Doc, " & _
            " vwConsPCCartera.FechaEmision, vwConsPCCartera.FechaVenci, " & _
            " datediff(dd,vwConsPCCartera.FechaVenci, " & FechaYMD(gobjMain.objCondicion.FechaCorte, gobjMain.EmpresaActual.TipoDB) & ") as DiasVencido," & _
            " vwConsPCCartera.Valor , vwConsPCCartera.Saldo," & _
            " PCKn.bandnoti1, " & _
            " PCKn.fechanoti1," & _
            " PCKn.bandnoti2, " & _
            " PCKn.fechanoti2," & _
            " PCKn.bandnoti3, " & _
            " PCKn.fechanoti3," & _
            " '' as imp, '' as sec "
        sql = sql & " FROM " & _
        " PcProvCli inner join TMP4 vwConsPCCartera ON PcProvCli.IdProvCli = vwConsPCCartera.IdProvCli " & _
        " left join tmp8 pckn on pckn.idpckardex = vwConsPCCartera.id "
        sql = sql & Condicion
        sql = sql & " ORDER BY vwConsPCCartera.CodProvCli,vwConsPCCartera.Fechavenci "
End With
        grd.Redraw = False
        MensajeStatus MSG_PREPARA, vbHourglass
        MiGetRowsRep gobjMain.EmpresaActual.OpenRecordset(sql), grd
        MensajeStatus "Listo", vbNormal
End Sub

Private Function PreparaCadena(ByVal lst As ListBox) As String
    Dim i As Long, s As String
    
    With lst
        'Si est· seleccionado solo una
        If .SelCount = 1 Then
            For i = 0 To .ListCount - 1
                If .Selected(i) Then
                    s = "'" & .List(i) & "'"
                    Exit For
                End If
            Next i
        'Si est· TODO o NINGUNO, no hay condiciÛn
        ElseIf (.SelCount < .ListCount) And (.SelCount > 0) Then
            For i = 0 To .ListCount - 1
                If .Selected(i) Then
                    s = s & "'" & .List(i) & "', "
                End If
            Next i
            If Len(s) > 0 Then s = Left$(s, Len(s) - 2)    'Quita la ultima ", "
        End If
    End With
    PreparaCadena = s
End Function

Private Sub ConfigCols()
    With grd
        .FormatString = "^#|tid|<CodProvCli|<Nombre|<Trans|<Doc Ref|<FechaEmision|<FechaVenci|^Dias Ven|>Valor|>Saldo|^1∫ Noti|<Fecha 1∫ Noti|^2∫ Noti|<Fecha 2∫ Noti|^3∫ Noti|<Fecha 3∫ Noti|^Imp|^Sec"
        .ColWidth(COL_NUMFILA) = 500
        .ColWidth(COL_CODPROVCLI) = 1100
        .ColWidth(COL_NOMBRE) = 2500
        .ColWidth(COL_TRANS) = 800
        .ColWidth(COL_NUMDOCREF) = 800
        .ColWidth(COL_FEMISION) = 1000
        .ColWidth(COL_FVENCI) = 1000
        .ColWidth(COL_VALOR) = 1000
        .ColWidth(COL_SALDO) = 1000
        .ColWidth(COL_BANDNOTI1) = 1000
        .ColWidth(COL_BANDNOTI2) = 1000
        .ColWidth(COL_BANDNOTI3) = 1000
        .ColWidth(COL_FECHANOTI1) = 1200
        .ColWidth(COL_FECHANOTI2) = 1200
        .ColWidth(COL_FECHANOTI3) = 1200
        .ColFormat(COL_VALOR) = gobjMain.EmpresaActual.GNOpcion.FormatoCantidad
        .ColFormat(COL_SALDO) = gobjMain.EmpresaActual.GNOpcion.FormatoCantidad
        .ColFormat(COL_FECHANOTI1) = gobjMain.EmpresaActual.GNOpcion.FormatoFecha
        .ColFormat(COL_FECHANOTI2) = gobjMain.EmpresaActual.GNOpcion.FormatoFecha
        .ColFormat(COL_FECHANOTI3) = gobjMain.EmpresaActual.GNOpcion.FormatoFecha
        
'        .ColHidden(COL_TID) = True
'        .ColHidden(COL_IMP) = True
'        .ColHidden(COL_SEC) = True
         
        .ColDataType(COL_FEMISION) = flexDTDate
        .ColDataType(COL_FVENCI) = flexDTDate
        .ColDataType(COL_FECHANOTI1) = flexDTDate
        .ColDataType(COL_FECHANOTI2) = flexDTDate
        .ColDataType(COL_FECHANOTI3) = flexDTDate
        
        .ColDataType(COL_BANDNOTI1) = flexDTBoolean
        .ColDataType(COL_BANDNOTI2) = flexDTBoolean
        .ColDataType(COL_BANDNOTI3) = flexDTBoolean
        .ColDataType(COL_IMP) = flexDTBoolean
        
        GNPoneNumFila grd, False
        grd.subtotal flexSTSum, 2, COL_VALOR, , grd.GridColor, , , "Subtotal", 2, True
        grd.subtotal flexSTSum, 2, COL_SALDO, , grd.GridColor, , , "Subtotal", 2, True
        grd.subtotal flexSTSum, -1, COL_VALOR, , grd.BackColorSel, vbYellow, , "Total", -1, True
        grd.subtotal flexSTSum, -1, COL_SALDO, , grd.BackColorSel, vbYellow, , "Total", -1, True
    End With
End Sub

Private Sub cmdBuscar1_Click()
'cargar datos para imprimir
Dim sql As String
    Dim rs As Recordset
    grd.Rows = 1
    If Opt(0).value = False And Opt(1).value = False And Opt(2).value = False Then: MsgBox "indique un numero de notificacion para imprimir": Exit Sub
    sql = "Select gn.codtrans,gn.numtrans,pc.codprovcli,pc.nombre,pckn.sec,'0' as Imp , FechaNoti1 "
'    If Opt(0).value = True Then
'        sql = sql & ", FechaNoti1 "
'    ElseIf Opt(1).value = True Then
'        sql = sql & ", FechaNoti2"
'    ElseIf Opt(2).value = True Then
'        sql = sql & ", FechaNoti3"
'    End If
    
    sql = sql & " from pcknotificacion pckn inner join pckardex  pck inner join pcprovcli pc"
    sql = sql & " on pc.idprovcli = pck.idprovcli"
    sql = sql & " Inner Join gncomprobante gn on gn.transid = pck.transid"
    sql = sql & " on pck.id= pckn.idpckardex "
    
    If Opt(0).value = True Then
        sql = sql & "Where bandnoti = 1"
    ElseIf Opt(1).value = True Then
        sql = sql & "Where bandnoti = 2"
    ElseIf Opt(2).value = True Then
        sql = sql & "Where bandnoti = 3"
    End If
    If Len(fcbCliente1.KeyText) > 0 And Len(fcbCliente2.KeyText) > 0 Then
        sql = sql & " And (pc.codprovcli between '" & fcbCliente1.KeyText & "' AND '" & fcbCliente2.KeyText & "')"
    End If
    
    sql = sql & " Group by gn.codtrans,gn.numtrans,pc.codprovcli,pc.nombre,pckn.sec, FechaNoti1 "
    
'    If Opt(0).value = True Then
'        sql = sql & ", FechaNoti1 "
'    ElseIf Opt(1).value = True Then
'        sql = sql & ", FechaNoti2"
'    ElseIf Opt(2).value = True Then
'        sql = sql & ", FechaNoti3"
'    End If
    
    sql = sql & " Order by pckn.sec"

    MiGetRowsRep gobjMain.EmpresaActual.OpenRecordset(sql), grd
    BorraRepetidos
    ConfigColsImp
End Sub
Private Sub BorraRepetidos()
Dim i As Long, j As Long
i = 1
Do While i <= grd.Rows - 1
    For j = grd.Rows - 1 To i + 1 Step -1
        If grd.TextMatrix(i, 5) = grd.TextMatrix(j, 5) Then
            grd.RemoveItem i
        End If
    Next
    i = i + 1
Loop
End Sub

Private Sub cmdCancelar_Click()
    If mProcesando Then
        mCancelado = True
    Else
        Unload Me
    End If
End Sub

Private Sub cmdGenNoti_Click()
     GrabarNotificacion
     BandHabilita False
End Sub

Private Sub cmdImprimir_Click()
Dim Filas  As Integer, i As Integer, j As Integer, v As Variant, s As String
If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "CATA") <> 0 Then

        Filas = 0
        ReDim v(26, 1)
            For i = 1 To grd.Rows - 1
'                If Not grd.IsSubtotal(i) Then
                    ReDim Preserve v(26, Filas)
                    For j = 1 To grd.Cols - 1
                        v(j - 1, Filas) = grd.TextMatrix(i, j)
                    Next j
                        Filas = Filas + 1
'                End If
            Next i

        s = "400" 'ntxMargIzq.value
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "ImpNoti_MarIzq", s
            
        s = "800" 'ntxMargSup.value
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "ImpNoti_MarSup", s
            
        s = 5 'ntxNUmEtiq.value
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "ImpNoti_NumEtiq", s
            
        'Graba en la base
        gobjMain.EmpresaActual.GNOpcion.Grabar


    FrmImprimeEtiketas.InicioNotificaciones v, dtpFechaPago.value
Else
    If txtPlantilla.Text <> "" Then
        Imprimir True
    Else
        MsgBox "Por favor especifique el nombre del archivo."
    End If
End If
End Sub

Private Sub Imprimir(Directo)
    Dim gc As GNComprobante
    Dim CodTrans As String
    
    Dim bandSigue As Boolean
    Dim numtrans As Long
    CodTrans = ""
    Dim i As Long, x As Long
    Dim j As Long, k As Long
    prg1.min = 0
    prg1.max = grd.Rows - 1
        Do While i <> grd.Rows - 1
                DoEvents
                prg1.value = i
                grd.Row = i
            x = grd.CellTop
            If Not grd.IsSubtotal(i) Then
                If grd.ValueMatrix(i, COL_IMP) <> -1 Then
                   ' For j = 12 To COL_FECHANOTI3 'grd.Cols - 2
                        'Select Case j
                         '   Case 12
                                If grd.Cell(flexcpBackColor, i, COL_FECHANOTI1, i, COL_FECHANOTI1) = vbWhite And (grd.ValueMatrix(i, COL_BANDNOTI1)) = -1 Then
                                    For k = 1 To Len(grd.TextMatrix(i, COL_TRANS))
                                        If Mid$(grd.TextMatrix(i, COL_TRANS), k, 1) <> " " Then
                                            CodTrans = CodTrans & Mid$(grd.TextMatrix(i, COL_TRANS), k, 1)
                                        Else
                                            numtrans = Right(grd.TextMatrix(i, COL_TRANS), Len(grd.TextMatrix(i, COL_TRANS)) - Len(CodTrans))
                                            Exit For
                                        End If
                                    Next
                                    Set gc = gobjMain.EmpresaActual.RecuperaGNComprobante(0, CodTrans, numtrans)
                                    If VerificaRepetidos(grd.TextMatrix(i, COL_CODPROVCLI), i, grd.TextMatrix(i, COL_SEC)) Then
                                        If Not ImprimeTrans(gc, mobjImp, txtPlantilla.Text, "NOTI1", Vector, grd.ValueMatrix(i, COL_SEC)) Then
                                            Me.Show
                                        End If
                                        
                                    End If
                                    Set gc = Nothing
                                    CodTrans = ""
                              '  End If
                            'Case 14
                                ElseIf grd.Cell(flexcpBackColor, i, COL_FECHANOTI2, i, COL_FECHANOTI2) = vbWhite And (grd.ValueMatrix(i, COL_BANDNOTI2)) = -1 Then 'IMPRIME SOLO LA NOTIFICACION ACTUAL
                                    For k = 1 To Len(grd.TextMatrix(i, COL_TRANS))
                                        If Mid$(grd.TextMatrix(i, COL_TRANS), k, 1) <> " " Then
                                            CodTrans = CodTrans & Mid$(grd.TextMatrix(i, COL_TRANS), k, 1)
                                        Else
                                            numtrans = Right(grd.TextMatrix(i, COL_TRANS), Len(grd.TextMatrix(i, COL_TRANS)) - Len(CodTrans))
                                            Exit For
                                        End If
                                    Next
                                    Set gc = gobjMain.EmpresaActual.RecuperaGNComprobante(0, CodTrans, numtrans)
                                        If VerificaRepetidos(grd.TextMatrix(i, COL_CODPROVCLI), i, grd.TextMatrix(i, COL_SEC)) Then
                                            If Not ImprimeTrans(gc, mobjImp, txtPlantilla.Text, "NOTI2", Vector, grd.ValueMatrix(i, COL_SEC)) Then
                                            Me.Show
                                            End If
                                        End If
                                    Set gc = Nothing
                                    CodTrans = ""
                                'End If
                            'Case 16
                                ElseIf grd.Cell(flexcpBackColor, i, COL_FECHANOTI3, i, COL_FECHANOTI3) = vbWhite And (grd.ValueMatrix(i, COL_BANDNOTI3)) = -1 Then
                                    For k = 1 To Len(grd.TextMatrix(i, COL_TRANS))
                                        If Mid$(grd.TextMatrix(i, COL_TRANS), k, 1) <> " " Then
                                            CodTrans = CodTrans & Mid$(grd.TextMatrix(i, COL_TRANS), k, 1)
                                        Else
                                            numtrans = Right(grd.TextMatrix(i, COL_TRANS), Len(grd.TextMatrix(i, COL_TRANS)) - Len(CodTrans))
                                            Exit For
                                        End If
                                    Next
                                    Set gc = gobjMain.EmpresaActual.RecuperaGNComprobante(0, CodTrans, numtrans)
                                    If VerificaRepetidos(grd.TextMatrix(i, COL_CODPROVCLI), i, grd.TextMatrix(i, COL_SEC)) Then
                                        If Not ImprimeTrans(gc, mobjImp, txtPlantilla.Text, "NOTI3", Vector, grd.ValueMatrix(i, COL_SEC)) Then
                                            Me.Show vbModal
                                        End If
                                    End If
                                    Set gc = Nothing
                                    CodTrans = ""
                                End If
                        'End Select
                    'Next
                End If
            End If
        'Next
            i = i + 1
            'Beep 10
        Loop
    prg1.value = prg1.min
    Screen.MousePointer = vbNormal
End Sub


Private Sub cmdImprimir2_Click()
Dim Filas  As Integer, i As Integer, j As Integer, v As Variant, s As String
If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "CATA") <> 0 Then

        Filas = 0
        ReDim v(15, 1)
            For i = 1 To grd.Rows - 1
                If Not grd.IsSubtotal(i) Then
                    ReDim Preserve v(15, Filas)
                    For j = 1 To grd.Cols - 1
                        v(j - 1, Filas) = grd.TextMatrix(i, j)
                    Next j
                        Filas = Filas + 1
                End If
            Next i

        s = "400" 'ntxMargIzq.value
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "ImpNoti_MarIzq", s
            
        s = "300" 'ntxMargSup.value
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "ImpNoti_MarSup", s
            
        s = 6 'ntxNUmEtiq.value
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "ImpNoti_NumEtiq", s
            
        'Graba en la base
        gobjMain.EmpresaActual.GNOpcion.Grabar


    FrmImprimeEtiketas.InicioNotificaciones v, dtpFechaPago.value
Else
    If txtPlantilla.Text <> "" Then
        ReImprimir True
    Else
        MsgBox "Por favor especifique el nombre del archivo."
    End If
End If
End Sub

Private Sub cmdTransLimpiar_Click()
    Dim i As Long, aux As Long
    aux = lstTrans.ListIndex
    For i = 0 To lstTrans.ListCount - 1
        lstTrans.Selected(i) = False
    Next i
    lstTrans.ListIndex = aux
    'LIMPIA FORMAS DE COBRO
    aux = lstForma.ListIndex
    For i = 0 To lstForma.ListCount - 1
        lstForma.Selected(i) = False
    Next i
    lstForma.ListIndex = aux
End Sub


Private Sub fcbCliente1_Selected(ByVal Text As String, ByVal KeyText As String)
fcbCliente2.KeyText = KeyText
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF9
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
    Unload Me
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    grd.Move 0, grd.Top, Me.ScaleWidth, Me.ScaleHeight - grd.Top - pic1.Height - 80
    prg1.Width = Me.ScaleWidth - (prg1.Left * 2)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set mColItems = Nothing         '*** MAKOTO 31/ago/00
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
                If Trim(Mid(Trim(Vector(i)), 2, Len(Trim(Vector(i))) - 2)) = lst.List(j) Then
                    lst.Selected(j) = True
                End If
            Next j
         Next i
    End If
End Sub
Private Sub GeneraNotificaciones()
Dim i As Long
    For i = 1 To grd.Rows - 1
            If Not grd.IsSubtotal(i) Then
           ' If DiaUno = 1 And Len(grd.TextMatrix(i, COL_FECHANOTI1)) > 0 And Len(grd.TextMatrix(i, COL_FECHANOTI2)) = 0 Then
            If Len(grd.TextMatrix(i, COL_FECHANOTI1)) > 0 And Len(grd.TextMatrix(i, COL_FECHANOTI2)) = 0 Then
                grd.TextMatrix(i, COL_BANDNOTI2) = -1
                grd.TextMatrix(i, COL_FECHANOTI2) = gobjMain.objCondicion.FechaCorte 'DateAdd("d", 1, gobjMain.objCondicion.FechaCorte)
            ElseIf Len(grd.TextMatrix(i, COL_FECHANOTI2)) > 0 And _
                 Len(grd.TextMatrix(i, COL_FECHANOTI1)) > 0 And _
                    Len(grd.TextMatrix(i, COL_FECHANOTI2)) And grd.Cell(flexcpBackColor, i, COL_BANDNOTI2, i, COL_FECHANOTI2) = &HC0FFC0 Then
                    'If DiaUno = DatePart("d", DateAdd("m", 1, grd.TextMatrix(i, COL_FECHANOTI2))) Then
                        grd.TextMatrix(i, COL_BANDNOTI3) = -1
                        grd.TextMatrix(i, COL_FECHANOTI3) = gobjMain.objCondicion.FechaCorte 'DateAdd("d", 1, gobjMain.objCondicion.FechaCorte)
                    'End If
               ' End If
            ElseIf grd.ValueMatrix(i, COL_DIASVENCI) > 0 And Len(grd.TextMatrix(i, COL_FECHANOTI1)) = 0 Then
                grd.TextMatrix(i, COL_BANDNOTI1) = -1
                grd.TextMatrix(i, COL_FECHANOTI1) = gobjMain.objCondicion.FechaCorte 'DateAdd("d", 1, gobjMain.objCondicion.FechaCorte)
            End If
        End If
    Next
End Sub
Private Function DiaUno() As Integer
    DiaUno = DatePart("d", DateAdd("d", 1, gobjMain.objCondicion.FechaCorte))
End Function
Private Function SiguienteMes() As Integer
    SiguienteMes = DatePart("m", DateAdd("m", 1, gobjMain.objCondicion.FechaCorte))
End Function


    Public Sub GrabarNotificacion()
Dim i As Long, x As Long, j As Long
'Dim EsNuevo As Boolean
Dim rs As Recordset, sql As String
If grd.Rows = 1 Then Exit Sub
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    For i = 1 To grd.Rows - 1
        DoEvents
        prg1.value = i
        grd.Row = i
        x = grd.CellTop
        If Not grd.IsSubtotal(i) Then
            If grd.ValueMatrix(i, COL_DIASVENCI) > 0 Then
                'Si es nuevo
            '    If EsNuevo(grd.TextMatrix(i, COL_TID)) Then
                    For j = 12 To COL_FECHANOTI3 'grd.Cols - 2
                        Select Case j
                            Case 12
                                If grd.Cell(flexcpBackColor, i, j, i, j) = vbWhite And (grd.ValueMatrix(i, j - 1)) = -1 Then
                                        sql = "SELECT * FROM pckNotificacion WHERE 1=0"
                                    Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
                                    rs.AddNew
                                    With rs
                                        !idpckardex = grd.ValueMatrix(i, COL_TID)
                                        !bandNoti = 1
                                        'If Len(grd.TextMatrix(i, COL_BANDNOTI1)) > 0 Then !bandNoti1 = grd.TextMatrix(i, COL_BANDNOTI1)
                                        'If Len(grd.TextMatrix(i, COL_BANDNOTI2)) > 0 Then !bandNoti2 = grd.TextMatrix(i, COL_BANDNOTI2)
                                        'If Len(grd.TextMatrix(i, COL_BANDNOTI3)) > 0 Then !bandNoti3 = grd.TextMatrix(i, COL_BANDNOTI3)
                                        If Len(grd.TextMatrix(i, COL_FECHANOTI1)) > 0 Then !FechaNoti1 = grd.TextMatrix(i, COL_FECHANOTI1)
                                        'If Len(grd.TextMatrix(i, COL_FECHANOTI2)) > 0 Then !FechaNoti2 = grd.TextMatrix(i, COL_FECHANOTI2)
                                        'If Len(grd.TextMatrix(i, COL_FECHANOTI3)) > 0 Then !FechaNoti3 = grd.TextMatrix(i, COL_FECHANOTI3)
                                        If grd.ValueMatrix(i, COL_SALDO) > 0 Then !valor = grd.ValueMatrix(i, COL_SALDO)
                                        If grd.ValueMatrix(i, COL_SEC) > 0 Then !sec = grd.ValueMatrix(i, COL_SEC)
                                        .Update
                                        .Close
                                    End With
                                    Set rs = Nothing
                                End If
                            Case 14
                                If grd.Cell(flexcpBackColor, i, j, i, j) = vbWhite And (grd.ValueMatrix(i, j - 1)) = -1 Then 'IMPRIME SOLO LA NOTIFICACION ACTUAL
                                    sql = "SELECT * FROM pckNotificacion WHERE 1=0"
                                    Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
                                    rs.AddNew
                                    With rs
                                        !idpckardex = grd.ValueMatrix(i, COL_TID)
                                        !bandNoti = 2
                                        'If Len(grd.TextMatrix(i, COL_BANDNOTI1)) > 0 Then !bandNoti1 = grd.TextMatrix(i, COL_BANDNOTI1)
                                        If Len(grd.TextMatrix(i, COL_BANDNOTI2)) > 0 Then !bandnoti2 = grd.TextMatrix(i, COL_BANDNOTI2)
                                        'If Len(grd.TextMatrix(i, COL_BANDNOTI3)) > 0 Then !bandNoti3 = grd.TextMatrix(i, COL_BANDNOTI3)
                                        If Len(grd.TextMatrix(i, COL_FECHANOTI2)) > 0 Then !FechaNoti1 = grd.TextMatrix(i, COL_FECHANOTI2)
                                        'If Len(grd.TextMatrix(i, COL_FECHANOTI2)) > 0 Then !FechaNoti2 = grd.TextMatrix(i, COL_FECHANOTI2)
                                        'If Len(grd.TextMatrix(i, COL_FECHANOTI3)) > 0 Then !FechaNoti3 = grd.TextMatrix(i, COL_FECHANOTI3)
                                        If grd.ValueMatrix(i, COL_SALDO) > 0 Then !valor = grd.ValueMatrix(i, COL_SALDO)
                                        If grd.ValueMatrix(i, COL_SEC) > 0 Then !sec = grd.ValueMatrix(i, COL_SEC)
                                        .Update
                                        .Close
                                    End With
                                    Set rs = Nothing
                                End If
                            Case 16
                                If grd.Cell(flexcpBackColor, i, j, i, j) = vbWhite And (grd.ValueMatrix(i, j - 1)) = -1 Then
                                    sql = "SELECT * FROM pckNotificacion WHERE 1=0"
                                    Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
                                    rs.AddNew
                                    With rs
                                        !idpckardex = grd.ValueMatrix(i, COL_TID)
                                        !bandNoti = 3
                                        'If Len(grd.TextMatrix(i, COL_BANDNOTI1)) > 0 Then !bandNoti1 = grd.TextMatrix(i, COL_BANDNOTI1)
                                        'If Len(grd.TextMatrix(i, COL_BANDNOTI2)) > 0 Then !bandNoti2 = grd.TextMatrix(i, COL_BANDNOTI2)
                                        'If Len(grd.TextMatrix(i, COL_BANDNOTI3)) > 0 Then !bandNoti3 = grd.TextMatrix(i, COL_BANDNOTI3)
                                        If Len(grd.TextMatrix(i, COL_FECHANOTI3)) > 0 Then !FechaNoti1 = grd.TextMatrix(i, COL_FECHANOTI3)
                                        'If Len(grd.TextMatrix(i, COL_FECHANOTI2)) > 0 Then !FechaNoti2 = grd.TextMatrix(i, COL_FECHANOTI2)
                                        'If Len(grd.TextMatrix(i, COL_FECHANOTI3)) > 0 Then !FechaNoti3 = grd.TextMatrix(i, COL_FECHANOTI3)
                                        If grd.ValueMatrix(i, COL_SALDO) > 0 Then !valor = grd.ValueMatrix(i, COL_SALDO)
                                        If grd.ValueMatrix(i, COL_SEC) > 0 Then !sec = grd.ValueMatrix(i, COL_SEC)
                                        .Update
                                        .Close
                                    End With
                                    Set rs = Nothing
                                End If
                        End Select
                    Next
                End If
'                Else
'                    sql = "Update  pckNotificacion " & _
'                        "Set  FechaNoti1 = '" & grd.TextMatrix(i, COL_FECHANOTI1) & "'," & _
'                        " bandNoti1 = " & grd.TextMatrix(i, COL_BANDNOTI1) & "," & _
'                        " FechaNoti2 = '" & grd.TextMatrix(i, COL_FECHANOTI2) & "'," & _
'                        " bandNoti2 = " & grd.TextMatrix(i, COL_BANDNOTI2) & "," & _
'                        " FechaNoti3 = '" & grd.TextMatrix(i, COL_FECHANOTI3) & "'," & _
'                        " bandNoti3 = " & grd.ValueMatrix(i, COL_BANDNOTI3) & "," & _
'                        " sec = " & grd.ValueMatrix(i, COL_SEC) & _
'                        "Where idpckardex = " & grd.TextMatrix(i, COL_TID)
'                        gobjMain.EmpresaActual.Execute sql, 1
'                End If
            End If
        'End If
    Next
    prg1.value = prg1.min
    Screen.MousePointer = vbNormal
End Sub

Private Function EsNuevo(ByVal tid As Long) As Boolean
On Error GoTo CapturaError
Dim sql As String
Dim rs As Recordset
sql = "select * from pckNotificacion where idpckardex = " & tid
Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
If rs.RecordCount = 0 Then
    EsNuevo = True
End If
Exit Function
CapturaError:
    MsgBox Err.Description
    Exit Function
    
End Function
Private Sub recuperaConfiguracion()
   txtPlantilla = GetSetting(APPNAME, App.Title, "Ruta PlantillaNoti", "")
End Sub
Private Sub GrabaConfiguracion()
Dim i As Integer
    SaveSetting APPNAME, App.Title, "Ruta PlantillaNoti", txtPlantilla.Text
End Sub
Public Sub CargaTodasFormasCobroPago(ByRef lst As ListBox)
    Dim rs As Recordset, Vector As Variant
    Dim numMod As Integer, i As Integer
    Dim s As String
    'Prepara la lista de tipos de transaccion
    lst.Clear
    Set rs = gobjMain.EmpresaActual.ListaTSFormaCobroPago(True, True, True)
    With rs
        If Not (.EOF) Then
            .MoveFirst
            Do Until .EOF
                lst.AddItem !codforma
                lst.ItemData(lst.NewIndex) = Len(!codforma) & " "
                .MoveNext
            Loop
        End If
    End With
    rs.Close
    s = GetSetting(APPNAME, App.Title, "KeyFormaNoti", "")
     RecuperaTrans "KeyFormaNoti", lstForma, s
    Set rs = Nothing
End Sub

Private Sub Form_Load()
  CREADOR = APPNAME
  cmdline = LCase(Command)
  If cmdline Like """*""" Then
    cmdline = Mid(cmdline, 2, Len(cmdline) - 2)
  End If
  
  If FileExists(cmdline) Then
    txtPlantilla.Text = cmdline
    btnConvert_Click
  End If
End Sub

Private Sub btnOpen_Click()
  Dim filename As String
  On Local Error Resume Next
  filename = OpenDialog(Me, "Text files (*.txt)|*.txt|All files (*.*)|*.*", _
                   "Select a text file", "")
  If Len(filename) Then
    txtPlantilla.Text = filename
    filename = txtPlantilla.Text
    GrabaConfiguracion
  End If
End Sub

Private Sub btnConvert_Click()
Dim i As Long, j As Long, x As Long
Dim cadSalida As String
  If txtPlantilla.Text <> "" Then
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
          For i = 1 To grd.Rows - 2
            DoEvents
            prg1.value = i
            grd.Row = i
            x = grd.CellTop
            If Not grd.IsSubtotal(i) Then
                For j = 12 To grd.Cols - 1
                    Select Case j
                        Case 12
                            If grd.Cell(flexcpBackColor, i, j, i, j) = vbWhite And (grd.ValueMatrix(i, j - 1)) = -1 Then
                                cadSalida = grd.TextMatrix(i, COL_TRANS) & "_" & Trim(grd.TextMatrix(i, COL_NOMBRE)) & "_Noti1.pdf"
                                ConvertToPDF txtPlantilla.Text, RecuperaCadena & cadSalida, _
                                        "Ishida y Asociados", CREADOR, "", _
                                        "", "", _
                                        "Times-Roman", 10, 0, _
                                        8.5, 11, grd.TextMatrix(i, COL_NOMBRE), gobjMain.objCondicion.FechaCorte, "Primera NotificaciÛn"
                                If FileExists(cmdline) Then
                                    Unload Me
                                Else
                                    ShellExecute 0, "Close", RecuperaCadena & cadSalida, vbNullString, vbNullString, 3
                                End If
                            End If
                        Case 14
                            If grd.Cell(flexcpBackColor, i, j, i, j) = vbWhite And (grd.ValueMatrix(i, j - 1)) = -1 Then
                                cadSalida = grd.TextMatrix(i, COL_TRANS) & "_" & Trim(grd.TextMatrix(i, COL_NOMBRE)) & "_Noti2.pdf"
                                ConvertToPDF txtPlantilla.Text, RecuperaCadena & cadSalida, _
                                        "Ishida y Asociados", CREADOR, "", _
                                        "", "", _
                                        "Times-Roman", 10, 0, _
                                        8.5, 11, grd.TextMatrix(i, COL_NOMBRE), gobjMain.objCondicion.FechaCorte, "Segunda NotificaciÛn"
                                If FileExists(cmdline) Then
                                    Unload Me
                                Else
                                    ShellExecute 0, "Close", RecuperaCadena & cadSalida, vbNullString, vbNullString, 3
                                End If
                            End If
                        Case 16
                            If grd.Cell(flexcpBackColor, i, j, i, j) = vbWhite And (grd.ValueMatrix(i, j - 1)) = -1 Then
                                cadSalida = grd.TextMatrix(i, COL_TRANS) & "_" & Trim(grd.TextMatrix(i, COL_NOMBRE)) & "_Noti3.pdf"
                                ConvertToPDF txtPlantilla.Text, RecuperaCadena & cadSalida, _
                                        "Ishida y Asociados", CREADOR, "", _
                                        "", "", _
                                        "Times-Roman", 10, 0, _
                                        8.5, 11, grd.TextMatrix(i, COL_NOMBRE), gobjMain.objCondicion.FechaCorte, "Tercera NotificaciÛn"
                                If FileExists(cmdline) Then
                                    Unload Me
                                Else
                                    ShellExecute 0, "Close", RecuperaCadena & cadSalida, vbNullString, vbNullString, 3
                                End If
                            End If
                        End Select
                    Next
                End If
            Next
    Else
        MsgBox "Por favor especifique el nombre del archivo."
    End If
    prg1.value = prg1.min
    Screen.MousePointer = vbNormal
End Sub

Public Sub ConvertToPDF(filename As String, outputfile As String, _
                        Optional TextAuthor As String, Optional TextCreator As String, Optional TextKeywords As String, _
                        Optional TextSubject As String, Optional TextTitle As String, _
                        Optional FontName As String = "Courier", Optional FontSize As Integer = 10, Optional Rotation As Integer, _
                        Optional pwidth As Single = 8.5, Optional pheight As Single = 11, Optional nombre As String, Optional fecha As Date, Optional NOTI As String)
  On Error GoTo er
  If Not FileExists(filename) Then
    MsgBox "Archivo '" & filename & "' no existe"
    Exit Sub
  ElseIf FileExists(outputfile) Then
    Kill outputfile
  End If
  
  initialize FontName, FontSize, Rotation, pwidth, pheight
  
  author = TextAuthor
  creator = TextCreator
  keywords = TextKeywords
  subject = TextSubject
  Title = TextTitle
  filetxt = filename
  filepdf = outputfile
  
  Call WriteStart
  Call WriteHead
  Call WritePages(nombre, fecha, NOTI)
  Call endpdf
  Exit Sub
er:
  MsgBox Err.Description
End Sub

Private Sub initialize(FontName As String, FontSize As Integer, Rotation As Integer, pwidth As Single, pheight As Single)
  pageHeight = 72 * pheight
  pageWidth = 72 * pwidth

  BaseFont = FontName ' Courier, Times-Roman, Arial
  pointSize = FontSize ' Font Size; Don't change it
  vertSpace = FontSize * 1.2 ' Vertical spacing
  rotate = Rotation ' degrees to rotate; try setting 90,180,etc
  lines = (pageHeight - 72) / vertSpace ' no of lines on one page
  
  Select Case LCase(FontName)
   Case "courier": linelen = 1.5 * pageWidth / pointSize
   Case "arial": linelen = 2 * pageWidth / pointSize
  'Case "Times-Roman": linelen = 2.2 * pageWidth / pointSize
   Case Else: linelen = 2.2 * pageWidth / pointSize
  End Select

  obj = 0
  npagex = pageWidth / 2
  npagey = 25
  pageNo = 0
  Position = 0
  cache = ""
End Sub

Private Sub writepdf(stre As String, Optional flush As Boolean)
  On Local Error Resume Next
  Position = Position + Len(stre)
  cache = cache & stre & vbCr
  If Len(cache) > 32000 Or flush Then
    Open filepdf For Append As #1
    Print #1, cache;
    Close #1
    cache = ""
  End If
End Sub
  
Private Sub WriteStart()
  writepdf ("%PDF-1.2")
  writepdf ("%‚„œ”")
End Sub

Private Sub WriteHead()
  Dim CreationDate As String
  On Error GoTo er
    CreationDate = "D:" & Format(Now, "YYYYMMDDHHNNSS")
    obj = obj + 1
    location(obj) = Position
    info = obj
    
    writepdf (obj & " 0 obj")
    writepdf ("<<")
    writepdf ("/Author (" & author & ")")
    writepdf ("/CreationDate (" & CreationDate & ")")
    writepdf ("/Creator (" & creator & ")")
    writepdf ("/Producer (" & APPNAME & ")")
    writepdf ("/Title (" & Title & ")")
    writepdf ("/Subject (" & subject & ")")
    writepdf ("/Keywords (" & keywords & ")")
    writepdf (">>")
    writepdf ("endobj")
    
    obj = obj + 1
    root = obj
    obj = obj + 1
    Tpages = obj
    encoding = obj + 2
    resources = obj + 3
    
    obj = obj + 1
    location(obj) = Position
    writepdf (obj & " 0 obj")
    writepdf ("<<")
    writepdf ("/Type /Font")
    writepdf ("/Subtype /Type1")
    writepdf ("/Name /F1")
    writepdf ("/Encoding " & encoding & " 0 R")
    writepdf ("/BaseFont /" & BaseFont)
    writepdf (">>")
    writepdf ("endobj")
    
    obj = obj + 1
    location(obj) = Position
    writepdf (obj & " 0 obj")
    writepdf ("<<")
    writepdf ("/Type /Encoding")
    writepdf ("/BaseEncoding /WinAnsiEncoding")
    writepdf (">>")
    writepdf ("endobj")
    
    obj = obj + 1
    location(obj) = Position
    writepdf (obj & " 0 obj")
    writepdf ("<<")
    writepdf ("  /Font << /F1 " & obj - 2 & " 0 R >>")
    writepdf ("  /ProcSet [ /PDF /Text ]")
    writepdf (">>")
    writepdf ("endobj")
  Exit Sub
er:
  MsgBox Err.Description
End Sub
  
Private Sub WritePages(ByVal nombre As String, ByVal fecha As Date, ByVal NOTI As String)
  Dim i As Integer
  Dim line As String, tmpline As String, beginstream As String
  On Error GoTo er
    Open filetxt For Input As #2
      beginstream = StartPage
      lineNo = -1
      Do Until EOF(2)
        Line Input #2, line
        
        lineNo = lineNo + 1
        
        'page break
        If lineNo >= lines Or InStr(line, Chr(12)) > 0 Then
            
        
          writepdf ("1 0 0 1 " & npagex & " " & npagey & " Tm")
          writepdf ("(" & pageNo & ") Tj")
          writepdf ("/F1 " & pointSize & " Tf")
          endpage (beginstream)
          beginstream = StartPage
        End If
        
        If InStr(line, "SR") > 0 Then
            line = line & "    " & nombre
        End If
        If InStr(line, "FECHA") > 0 Then
            line = Format(fecha, "dddd, MMM d yyyy")
        End If
        
        If InStr(line, "NOTIFICACION") > 0 Then
            line = NOTI
        End If
        
        line = ReplaceText(ReplaceText(line, "(", "\("), ")", "\)")
        line = Trim(line)
        
        If Len(line) > linelen Then
          
          'word wrap
          Do While Len(line) > linelen
            tmpline = Left(line, linelen)
            For i = Len(tmpline) To Len(tmpline) \ 2 Step -1
              If InStr("*&^%$#,. ;<=>[])}!""", Mid(tmpline, i, 1)) Then
                tmpline = Left(tmpline, i)
                Exit For
              End If
            Next
            
            line = Mid$(line, Len(tmpline) + 1)
            writepdf ("T* (" & tmpline & vbCrLf & ") Tj")
            lineNo = lineNo + 1
            
            'page break
            If lineNo >= lines Or InStr(line, Chr(12)) > 0 Then
              writepdf ("1 0 0 1 " & npagex & " " & npagey & " Tm")
              writepdf ("(" & pageNo & ") Tj")
              writepdf ("/F1 " & pointSize & " Tf")
              endpage (beginstream)
              beginstream = StartPage
            End If
          Loop
          
          lineNo = lineNo + 1
          writepdf ("T* (" & line & vbCrLf & ") Tj")
        
        Else
          
          writepdf ("T* (" & line & vbCrLf & ") Tj")
        
        End If
      Loop
    Close #2
    writepdf ("1 0 0 1 " & npagex & " " & npagey & " Tm")
    writepdf ("(" & pageNo & ") Tj")
    writepdf ("/F1 " & pointSize & " Tf")
    endpage (beginstream)
  Exit Sub
er:
  MsgBox Err.Description
  Close
End Sub

Private Function StartPage() As String
  Dim strmpos As Long
  On Error GoTo er
  obj = obj + 1
  location(obj) = Position
  pageNo = pageNo + 1
  pageObj(pageNo) = obj
  
  writepdf (obj & " 0 obj")
  writepdf ("<<")
  writepdf ("/Type /Page")
  writepdf ("/Parent " & Tpages & " 0 R")
  writepdf ("/Resources " & resources & " 0 R")
  obj = obj + 1
  writepdf ("/Contents " & obj & " 0 R")
  writepdf ("/Rotate " & rotate)
  writepdf (">>")
  writepdf ("endobj")
  
  location(obj) = Position
  writepdf (obj & " 0 obj")
  writepdf ("<<")
  writepdf ("/Length " & obj + 1 & " 0 R")
  writepdf (">>")
  writepdf ("stream")
  strmpos = Position
  writepdf ("BT")
  writepdf ("/F1 " & pointSize & " Tf")
  writepdf ("1 0 0 1 50 " & pageHeight - 40 & " Tm")
  writepdf (vertSpace & " TL")
  
  StartPage = strmpos
  Exit Function
er:
  MsgBox Err.Description
End Function

Function endpage(streamstart As Long) As String
  Dim streamEnd As Long
  On Error GoTo er
    writepdf ("ET")
    streamEnd = Position
    writepdf ("endstream")
    writepdf ("endobj")
    obj = obj + 1
    location(obj) = Position
    writepdf (obj & " 0 obj")
    writepdf (streamEnd - streamstart)
    writepdf "endobj"
    lineNo = 0
  Exit Function
er:
  MsgBox Err.Description
End Function

Sub endpdf()
  Dim ty As String, i As Integer, xreF As Long
  On Error GoTo er
    location(root) = Position
    writepdf (root & " 0 obj")
    writepdf ("<<")
    writepdf ("/Type /Catalog")
    writepdf ("/Pages " & Tpages & " 0 R")
    writepdf (">>")
    writepdf ("endobj")
    location(Tpages) = Position
    writepdf (Tpages & " 0 obj")
    writepdf ("<<")
    writepdf ("/Type /Pages")
    writepdf ("/Count " & pageNo)
    writepdf ("/MediaBox [ 0 0 " & pageWidth & " " & pageHeight & " ]")
    ty = ("/Kids [ ")
    For i = 1 To pageNo
      ty = ty & pageObj(i) & " 0 R "
    Next i
    ty = ty & "]"
    writepdf (ty)
    writepdf (">>")
    writepdf ("endobj")
    xreF = Position
    writepdf ("0 " & obj + 1)
    writepdf ("0000000000 65535 f ")
    For i = 1 To obj
      writepdf (Format(location(i), "0000000000") & " 00000 n ")
    Next i
    writepdf ("trailer")
    writepdf ("<<")
    writepdf ("/Size " & obj + 1)
    writepdf ("/Root " & root & " 0 R")
    writepdf ("/Info " & info & " 0 R")
    writepdf (">>")
    writepdf ("startxref")
    writepdf (xreF)
    writepdf "%%EOF", True
  Exit Sub
er:
  MsgBox Err.Description
End Sub

Public Function FileExists(ByVal filename As String) As Boolean
  On Error Resume Next
  FileExists = FileLen(filename) > 0
  Err.Clear
End Function

Public Function ReplaceText(Text As String, TextToReplace As String, NewText As String) As String
  Dim mtext As String, SpacePos As Long
  mtext = Text
  SpacePos = InStr(mtext, TextToReplace)
  Do While SpacePos
    mtext = Left(mtext, SpacePos - 1) & NewText & Mid(mtext, SpacePos + Len(TextToReplace))
    SpacePos = InStr(SpacePos + Len(NewText), mtext, TextToReplace)
  Loop
  ReplaceText = mtext
End Function

Function SaveDialog(Form1 As Form, Filter As String, Title As String, InitDir As String, DefaultFilename As String) As String
  Dim ofn As OPENFILENAME
  Dim a As Long
  On Local Error Resume Next
  ofn.lStructSize = Len(ofn)
  ofn.hwndOwner = Form1.hWnd
  ofn.hInstance = App.hInstance
  If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
  For a = 1 To Len(Filter)
      If Mid$(Filter, a, 1) = "|" Then Mid$(Filter, a, 1) = Chr$(0)
  Next
  ofn.lpstrFilter = Filter
  ofn.lpstrFile = Space$(254)
  Mid(ofn.lpstrFile, 1, 254) = DefaultFilename
  ofn.nMaxFile = 255
  ofn.lpstrFileTitle = Space$(254)
  ofn.nMaxFileTitle = 255
  ofn.lpstrInitialDir = InitDir
  ofn.lpstrTitle = Title
  ofn.lpstrDefExt = "pdf"
  ofn.flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
  a = GetSaveFileName(ofn)
  If (a) Then
      SaveDialog = Trim$(ofn.lpstrFile)
  Else
      SaveDialog = ""
  End If
End Function

Function OpenDialog(Form1 As Form, Filter As String, Title As String, InitDir As String) As String
  Dim ofn As OPENFILENAME
  Dim a As Long
  On Local Error Resume Next
  ofn.lStructSize = Len(ofn)
  ofn.hwndOwner = Form1.hWnd
  ofn.hInstance = App.hInstance
  If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"

  For a = 1 To Len(Filter)
      If Mid$(Filter, a, 1) = "|" Then Mid$(Filter, a, 1) = Chr$(0)
  Next
  ofn.lpstrFilter = Filter
  ofn.lpstrFile = Space$(254)
  ofn.nMaxFile = 255
  ofn.lpstrFileTitle = Space$(254)
  ofn.nMaxFileTitle = 255
  ofn.lpstrInitialDir = InitDir
  ofn.lpstrTitle = Title
  ofn.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
  a = GetOpenFileName(ofn)
  If (a) Then
      OpenDialog = Trim$(ofn.lpstrFile)
  Else
      OpenDialog = ""
  End If
End Function

Private Function RecuperaCadena() As String
Dim i As Long
Dim s As String, cad As String
s = txtPlantilla.Text
For i = Len(s) To 1 Step -1
    If Mid$(s, i, 1) <> "\" Then
        cad = cad & Mid$(s, i, 1)
    Else
        Exit For
    End If
Next
    s = Left(s, Len(s) - Len(cad))
    RecuperaCadena = s
End Function

Public Sub InicioAnuladas()
    Dim i As Integer
    On Error GoTo ErrTrap
    Me.Show
    Me.ZOrder
    BandHabilita False
    dtpFechaCorte.value = Date
    CargaTrans
    CargaTodasFormasCobroPago lstForma
    recuperaConfiguracion
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Function VerificaRepetidos(ByVal codcliente As String, ByVal filaAct As Long, ByVal sec As Long) As Boolean
Dim i As Long
Vector = ""
    For i = filaAct To grd.Rows - 1
        If grd.TextMatrix(i, COL_CODPROVCLI) = codcliente And grd.ValueMatrix(i, COL_SEC) = sec Then
            Vector = Vector & grd.ValueMatrix(i, COL_TID) & ","
            grd.TextMatrix(i, COL_IMP) = -1
            VerificaRepetidos = True
        End If
    Next
'    VerificaRepetidos = False
End Function

Private Sub AsignaSecuencial()
Dim sql As String
Dim i As Long, ind As Long
Dim suma As Integer
Dim sec As Integer
Dim rs As Recordset
Dim band1 As Boolean
Dim band2 As Boolean
Dim band3 As Boolean
Dim x As Long


UltimaSecuencia ind
suma = grd.ValueMatrix(ind, COL_SEC)
If ind = 1 Then ind = ind + 1
 
'28
For i = ind To grd.Rows - 1
'            DoEvents
'            grd.Row = i
'            x = grd.CellTop
    If Not grd.IsSubtotal(i) Then 'NOTI1
            If Len(grd.TextMatrix(i, COL_FECHANOTI2)) > 0 Then Exit For
        If grd.Cell(flexcpBackColor, i, COL_BANDNOTI1, i, COL_BANDNOTI1) = vbWhite And Len(grd.TextMatrix(i, COL_FECHANOTI1)) > 0 Then
            If grd.TextMatrix(i, COL_CODPROVCLI) = grd.TextMatrix(i - 1, COL_CODPROVCLI) Then
                If Len(grd.TextMatrix(i - 1, COL_FECHANOTI1)) = 0 Then
                    suma = suma + 1
                    grd.TextMatrix(i, COL_SEC) = suma
                Else
                    grd.TextMatrix(i, COL_SEC) = grd.TextMatrix(i - 1, COL_SEC)
                End If
            Else
                suma = suma + 1
                grd.TextMatrix(i, COL_SEC) = suma

            End If
        End If
        band1 = True
    End If
Next

For i = ind To grd.Rows - 1
'            DoEvents
'            grd.Row = i
'            x = grd.CellTop
    If Not grd.IsSubtotal(i) Then 'NOTI2
    If Len(grd.TextMatrix(i, COL_FECHANOTI3)) > 0 Then Exit For
'        If band1 = True Then Exit For
        If grd.Cell(flexcpBackColor, i, COL_BANDNOTI2, i, COL_BANDNOTI2) = vbWhite And Len(grd.TextMatrix(i, COL_FECHANOTI2)) > 0 Then
            If grd.TextMatrix(i, COL_CODPROVCLI) = grd.TextMatrix(i - 1, COL_CODPROVCLI) Then
'                If Len(grd.TextMatrix(i - 1, COL_FECHANOTI2)) = 0 Then
'                    suma = suma + 1
'                    grd.TextMatrix(i, COL_SEC) = suma
'                Else
                    grd.TextMatrix(i, COL_SEC) = grd.TextMatrix(i - 1, COL_SEC)
'                End If
            Else
                suma = suma + 1
                grd.TextMatrix(i, COL_SEC) = suma
            End If
        ElseIf grd.Cell(flexcpBackColor, i, COL_BANDNOTI1, i, COL_BANDNOTI1) = vbWhite And Len(grd.TextMatrix(i, COL_FECHANOTI1)) > 0 Then
                suma = suma + 1
                grd.TextMatrix(i, COL_SEC) = suma
                
        End If
        band2 = True
    End If
Next

For i = ind To grd.Rows - 1
'            DoEvents
'            grd.Row = i
'            x = grd.CellTop
'            If band1 = True Then Exit For
 '           If band2 = True Then Exit For
        If Not grd.IsSubtotal(i) Then 'NOTI3
        If grd.Cell(flexcpBackColor, i, COL_BANDNOTI3, i, COL_BANDNOTI3) = vbWhite And Len(grd.TextMatrix(i, COL_FECHANOTI3)) > 0 Then
            If grd.TextMatrix(i, COL_CODPROVCLI) = grd.TextMatrix(i - 1, COL_CODPROVCLI) Then
                
'                If Len(grd.TextMatrix(i - 1, COL_FECHANOTI3)) = 0 Then
'                    suma = suma + 1
'                    grd.TextMatrix(i, COL_SEC) = suma
'                Else
                    grd.TextMatrix(i, COL_SEC) = grd.TextMatrix(i - 1, COL_SEC)
'                End If
            Else
                suma = suma + 1
                grd.TextMatrix(i, COL_SEC) = suma
          
            End If
        ElseIf grd.Cell(flexcpBackColor, i, COL_BANDNOTI2, i, COL_BANDNOTI2) = vbWhite And Len(grd.TextMatrix(i, COL_FECHANOTI2)) > 0 Then
                suma = suma + 1
                grd.TextMatrix(i, COL_SEC) = suma
        ElseIf grd.Cell(flexcpBackColor, i, COL_BANDNOTI1, i, COL_BANDNOTI1) = vbWhite And Len(grd.TextMatrix(i, COL_FECHANOTI1)) > 0 Then
                suma = suma + 1
                grd.TextMatrix(i, COL_SEC) = suma
                
        End If
            band3 = True
    End If
Next


End Sub

Private Sub UltimaSecuencia(ByRef fila As Long)
Dim sql As String
Dim i As Long
Dim suma As Integer
Dim sec As Integer
Dim rs As Recordset
On Error GoTo brek
sql = "Select top 1 max(sec)+1 as sec from pcknotificacion"
Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
If rs.EOF Or IsNull(rs!sec) Then
    grd.TextMatrix(1, COL_SEC) = 1
    fila = 1
Else
    For i = 1 To grd.Rows - 1  'NOTI1
        If Not grd.IsSubtotal(i) Then
            If Len(grd.TextMatrix(i, COL_FECHANOTI2)) > 0 Then Exit For
            If grd.Cell(flexcpBackColor, i, COL_BANDNOTI1, i, COL_BANDNOTI1) = vbWhite Then
                grd.TextMatrix(i, COL_SEC) = rs!sec
                fila = i
                GoTo brek
            End If
        End If
    Next
    For i = 1 To grd.Rows - 1  'NOTI2
        If Not grd.IsSubtotal(i) Then
            If Len(grd.TextMatrix(i, COL_FECHANOTI3)) > 0 Then Exit For
            If grd.Cell(flexcpBackColor, i, COL_BANDNOTI2, i, COL_BANDNOTI2) = vbWhite Then
                grd.TextMatrix(i, COL_SEC) = rs!sec
                fila = i
                GoTo brek
            End If
        End If
    Next
    For i = 1 To grd.Rows - 1  'NOTI3
        If Not grd.IsSubtotal(i) Then
            If grd.Cell(flexcpBackColor, i, COL_BANDNOTI3, i, COL_BANDNOTI3) = vbWhite Then
                grd.TextMatrix(i, COL_SEC) = rs!sec
                fila = i
                GoTo brek
            End If
        End If
    Next
End If
brek:
    Set rs = Nothing
    Exit Sub
End Sub

Public Sub InicioImpresion()
    Dim i As Integer
    On Error GoTo ErrTrap
    Me.Show
    Me.ZOrder
    grd.Rows = 1
    BandHabilita False
    dtpFechaCorte.value = Date
    cmdBuscar1.Visible = True
    cmdBuscar.Visible = False
    cmdImprimir.Visible = False
    cmdImprimir2.Visible = True
    Pic.Visible = True
    cmdTransLimpiar.Visible = False
    cmdGenNoti.Visible = False
    recuperaConfiguracion
    CargaClientes
    grd.Enabled = True
    grd.Editable = flexEDKbd
    
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub CargaClientes()
    Dim v() As Variant
    Dim sql  As String, rs As Recordset, cond As String
    fcbCliente1.Clear
    fcbCliente2.Clear
    cond = " WHERE bandCliente = 1"
    sql = "SELECT CodProvCli, Nombre FROM PCProvCli "
    sql = sql & cond
    sql = sql & " ORDER BY Nombre"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not rs.EOF Then
        v = MiGetRows(rs)
        fcbCliente1.SetData v
        fcbCliente2.SetData v
    End If
    fcbCliente1.Text = ""
    fcbCliente2.Text = ""
End Sub

Private Sub ConfigColsImp()
    With grd
        .FormatString = "^#|<CodTrans|<NumTrans|<CodProvCli|<Nombre|>Num Noti|^Imp"
        .ColWidth(0) = 500
        .ColWidth(3) = 1500
        .ColWidth(4) = 5500
        .ColWidth(5) = 800
        .ColWidth(6) = 800
        .ColHidden(1) = True
        .ColHidden(2) = True
        .ColDataType(6) = flexDTBoolean
    End With
    GNPoneNumFila grd, False
End Sub

Private Sub ReImprimir(Directo)
    Dim gc As GNComprobante
    Dim CodTrans As String
    Dim numtrans As Long
    CodTrans = ""
    Dim i As Long
    Dim j As Long
    Dim sec As Integer
    Dim x As Long
    prg1.min = 0
    prg1.max = grd.Rows - 1
    i = 1
    Do While i <> grd.Rows
        DoEvents
        prg1.value = i
        grd.Row = i
        x = grd.CellTop
        If Not grd.IsSubtotal(i) Then
            If grd.ValueMatrix(i, 6) = -1 Then
                CodTrans = grd.TextMatrix(i, 1)
                numtrans = grd.ValueMatrix(i, 2)
                Set gc = gobjMain.EmpresaActual.RecuperaGNComprobante(0, CodTrans, numtrans)
                If Opt(0) = True Then
                    Vector = RecuperaId(grd.TextMatrix(i, 3), 1)
                    If Not ImprimeTrans(gc, mobjImp, txtPlantilla.Text, "NOTI1", Vector, grd.ValueMatrix(i, 5)) Then
                        Me.Show
                    End If
                ElseIf Opt(1) = True Then
                    Vector = RecuperaId(grd.TextMatrix(i, 3), 2)
                    If Not ImprimeTrans(gc, mobjImp, txtPlantilla.Text, "NOTI2", Vector, grd.ValueMatrix(i, 5)) Then
                        Me.Show
                    End If
                ElseIf Opt(2) = True Then
                    Vector = RecuperaId(grd.TextMatrix(i, 3), 3)
                    If Not ImprimeTrans(gc, mobjImp, txtPlantilla.Text, "NOTI3", Vector, grd.ValueMatrix(i, 5)) Then
                        Me.Show vbModal
                    End If
                End If
            End If
        End If
        i = i + 1
    Loop
    prg1.value = prg1.min
    Screen.MousePointer = vbNormal
End Sub

Private Function RecuperaId(ByVal CodCli As String, ByVal bandNoti As Integer) As String
Dim sql As String
Dim cad As String
Dim rs As Recordset
    sql = "Select idpckardex from pcknotificacion pckn "
    sql = sql & "Inner join pckardex pck "
    sql = sql & "Inner join pcprovcli pc "
    sql = sql & "ON pc.idprovcli = pck.idprovcli "
    sql = sql & "ON pck.id= pckn.idpckardex "
    sql = sql & "Where pc.codprovcli = '" & CodCli & "'"
    If bandNoti = 1 Then
        sql = sql & " AND pckn.bandNoti = 1 "
    ElseIf bandNoti = 2 Then
        sql = sql & " AND pckn.bandnoti = 2 "
    ElseIf bandNoti = 3 Then
        sql = sql & " AND pckn.bandnoti = 3 "
    End If
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    Do While Not rs.EOF
        cad = cad & rs!idpckardex & ","
        rs.MoveNext
    Loop
    RecuperaId = cad
End Function

Private Sub BuscarEstudiantes()
Dim aux As String
Dim sql As String
Dim Condicion As String
MensajeStatus "Procesando", vbHourglass
With gobjMain.objCondicion
'1) Prepara los  documentos  Asignados  menores a la fecha
        VerificaExistenciaTabla 1
        'aux = IIf(.NumMoneda > 0, "/Cotizacion" & .NumMoneda + 1, "")
        sql = "SELECT " & _
            "pck.IdAsignado, " & _
            "(pck.Debe + pck.Haber)  AS Valor " & _
            "INTO tmp1 " & _
            "From " & _
            "GNtrans gt INNER JOIN " & _
                "(GNComprobante gc INNER JOIN PCKardex pck " & _
                "ON gc.transID = pck.transID) " & _
                          "ON gt.Codtrans = gc.Codtrans " & _
            "Where (pck.IdAsignado <> 0) " & _
            "AND (gc.Estado <> 3) " & _
            "AND (gt.AfectaSaldoPC=1) " & _
            "AND (gc.Fechatrans<= " & FechaYMD(.FechaCorte, gobjMain.EmpresaActual.TipoDB) & ")"
        gobjMain.EmpresaActual.EjecutarSQL sql, 1
        '2)Agrupa  estos  documentos por IdAsignado
        VerificaExistenciaTabla 2
        sql = "SELECT " & _
              "IdAsignado," & _
              "isnull(Sum(Valor),0) AS VCancelado " & _
              "INTO tmp2 " & _
              "FROM tmp1 " & _
              "GROUP BY IdAsignado"
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
        '3) Agrupa los documentos con su valor cancelado por ID
        VerificaExistenciaTabla 3
        sql = "SELECT " & _
                "pck.Id, " & _
                "pck.Debe + pck.Haber AS Valor, " & _
                "isnull(vw.VCancelado,0) AS VCancelado, " & _
                "(pck.Debe + pck.Haber) - isnull(vw.VCancelado,0)  AS Saldo " & _
                "INTO tmp3 " & _
            "FROM GNtrans INNER JOIN  GNComprobante gc INNER JOIN (tmp2 vw RIGHT JOIN PCKardex pck  ON vw.IdAsignado = pck.Id) " & _
            "ON gc.TransID = pck.TransID  ON  GNTrans.CodTrans = gc.CodTrans " & _
            "Where (pck.IdAsignado = 0) And (gc.Estado <> 3)  " & _
                    "AND (pck.debe >0) " & _
                    " AND ((GNtrans.AfectaSaldoPC) = " & CadenaBool(True, gobjMain.EmpresaActual.TipoDB) & ") "
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
        '4) Finalmente une  con el documento  Padre
       VerificaExistenciaTabla 4
       
        sql = " SELECT vwConsPCDocSaldo.Id, PCProvCli.IdProvCli, PCProvCli.CodProvCli, PCProvCli.Nombre,"
        sql = sql & " GNComprobante.CodTrans, GNComprobante.NumTrans, "
        sql = sql & " GNComprobante.CodTrans + ' ' + CONVERT(varchar, NumTrans)    AS Trans, "
        sql = sql & " CodForma, CodForma + pckardex.NumLetra AS Doc, "
        sql = sql & " PCKardex.FechaEmision, PCKardex.FechaVenci, "
        sql = sql & " (PCKardex.Debe + PCKardex.Haber) " & aux & " AS Valor, "
        sql = sql & " (PCKardex.Debe + PCKardex.Haber) " & aux & " - IsNull(vwConsPCDocSaldo.VCancelado,0) AS Saldo, "
        sql = sql & " PCKardex.Observacion, GNComprobante.CodUsuarioAutoriza, pcgar.nombre as Estudiante, "
        sql = sql & " pcg1.descripcion as deacGrupo1 ,"
        sql = sql & " pcg2.descripcion as deacGrupo2 ,"
        sql = sql & " pcg3.descripcion as deacGrupo3 ,"
        sql = sql & " pcg4.descripcion as deacGrupo4 ,"
        sql = sql & " fcvam.codvendedor as codbusetaAM , fcvam.nombre as busetaAM, fcvpm.nombre as busetaPM, fcvpm.codvendedor as codbusetaPM"
        sql = sql & " INTO tmp4 "
        sql = sql & " FROM PCProvCli  INNER JOIN  "
        sql = sql & " (GNTrans INNER JOIN "
        sql = sql & " (GNComprobante  "
        sql = sql & " INNER JOIN FCVENDEDOR FCgnc ON GNComprobante.IDVENDEDOR= FCgnc.IDVENDEDOR"
        sql = sql & " left join pcprovcli pcgar "
        sql = sql & " left join pcgrupo1 pcg1 on pcgar.idgrupo1=pcg1.idgrupo1 "
        sql = sql & " left join pcgrupo2 pcg2 on pcgar.idgrupo2=pcg2.idgrupo2 "
        sql = sql & " left join pcgrupo3 pcg3 on pcgar.idgrupo3=pcg3.idgrupo3  "
        sql = sql & " left join pcgrupo4 pcg4 on pcgar.idgrupo4=pcg4.idgrupo4  "
        sql = sql & " left join PCTransporte pct on pct.idprovcli=pcgar.idprovcli  "
        sql = sql & " left join gnvehiculo gnvam left join fcvendedor fcvam on gnvam.codvehiculo =fcvam.codvendedor on pct.idvehiculoam = gnvam.idvehiculo "
        sql = sql & " left join gnvehiculo gnvpm left join fcvendedor fcvpm on gnvpm.codvehiculo =fcvpm.codvendedor on pct.idvehiculopm = gnvpm.idvehiculo  "
        sql = sql & " on GNComprobante.idgaranteref = pcgar.idprovcli INNER JOIN "
        sql = sql & " (TSFormaCobroPago INNER JOIN "
        sql = sql & " (PCKardex left JOIN FcVendedor  FCV  on PCKardex.idvendedor= fcv.idvendedor INNER JOIN "
        sql = sql & " tmp3  vwConsPCDocSaldo ON PCKardex.Id = vwConsPCDocSaldo.Id) "
        sql = sql & " ON TSFormaCobroPago.IdForma = PCKardex.IdForma) ON "
        sql = sql & " GNComprobante.TransID = PCKardex.TransID) ON "
        sql = sql & " GNTrans.CodTrans = GNComprobante.CodTrans) ON "
        sql = sql & " PCProvCli.IdProvCli = PCKardex.IdProvCli "
        sql = sql & " Where (PCKardex.IdAsignado = 0) And (GNComprobante.Estado <> 3) "
        sql = sql & " AND (GNComprobante.Fechatrans<=" & FechaYMD(.FechaCorte, gobjMain.EmpresaActual.TipoDB) & ")"
        sql = sql & " AND (PCKardex.Debe >0) "
        sql = sql & " AND GNCOMPROBANTE.CodTrans IN (" & .CodTrans & ")"
         If Len(fcbGrupo.KeyText) > 0 Then
            sql = sql & " AND pcg1.codgrupo1 = '" & fcbGrupo.KeyText & "'"
         End If
        
         If Len(fcbChofer.KeyText) > 0 Then
            sql = sql & " AND FCgnc.CODVENDEDOR = '" & fcbChofer.KeyText & "'"
         End If
         
         If Len(fcbGar.KeyText) > 0 Then
            sql = sql & " AND PCGAR.CODPROVCLI= '" & fcbGar.KeyText & "'"
        End If

         If Len(fcbCli.KeyText) > 0 Then
            sql = sql & " AND PCPROVCLI.CODPROVCLI= '" & fcbCli.KeyText & "'"
        End If
        
        sql = sql & " AND (PCKardex.FechaVenci<=" & FechaYMD(dtpFechaCorte1.value, gobjMain.EmpresaActual.TipoDB) & ")"

        
        
        
       gobjMain.EmpresaActual.EjecutarSQL sql, 1
       
       ' CONSULTAS PARA PCNOTIFICACION
            VerificaExistenciaTabla 5
            sql = "select * into tmp5 from pcknotificacion  where bandnoti = 1 "
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            VerificaExistenciaTabla 6
            sql = "select * into tmp6 from pcknotificacion  where bandnoti = 2 "
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            VerificaExistenciaTabla 7
            sql = "select * into tmp7 from pcknotificacion  where bandnoti = 3 "
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            VerificaExistenciaTabla 8
            sql = " select  t5.idpckardex,t5.fechanoti1 as fechanoti1,case when t5.bandnoti= 1 then -1 else 0 end as bandnoti1,"
            sql = sql & "t6.fechanoti1 as fechanoti2,case when t6.bandnoti= 2 then -1 else 0 end as bandnoti2,"
            sql = sql & "t7.fechanoti1 as fechanoti3,case when t7.bandnoti= 3 then -1 else 0 end as bandnoti3 into tmp8"
            sql = sql & " from tmp5 t5 full join tmp6 t6 on t6.idpckardex = t5.idpckardex full join tmp7 t7 on t7.idpckardex=t6.idpckardex"
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            
    '4  Consulta  final
        Condicion = " WHERE vwConsPCCartera.Saldo >0 "
        Condicion = Condicion & " AND vwConsPCCartera.FechaEmision <= " & FechaYMD(.FechaCorte, gobjMain.EmpresaActual.TipoDB)
        
        If Len(.CodTrans) > 0 Then
            Condicion = Condicion & " AND vwConsPCCartera.CodTrans IN (" & .CodTrans & ") "
        End If
            Condicion = Condicion & " AND BandCliente=1"
        If Len(gobjMain.objCondicion.codforma) <> 0 Then
            Condicion = Condicion & " AND vwConsPCCartera.CodForma in ( " & .codforma & ")"
        End If
        sql = "SELECT vwConsPCCartera.id,  vwConsPCCartera.Estudiante, vwConsPCCartera.nombre," & _
            " vwConsPCCartera.trans, " & _
            " vwConsPCCartera.Doc, " & _
            " vwConsPCCartera.FechaEmision, vwConsPCCartera.FechaVenci, " & _
            " datediff(dd,vwConsPCCartera.FechaVenci, " & FechaYMD(gobjMain.objCondicion.FechaCorte, gobjMain.EmpresaActual.TipoDB) & ") as DiasVencido," & _
            " vwConsPCCartera.Valor , vwConsPCCartera.Saldo," & _
            " PCKn.bandnoti1, " & _
            " PCKn.fechanoti1," & _
            " PCKn.bandnoti2, " & _
            " PCKn.fechanoti2," & _
            " PCKn.bandnoti3, " & _
            " PCKn.fechanoti3," & _
            " '' as imp, '' as sec, "
        sql = sql & " deacGrupo1 ,"
        sql = sql & " deacGrupo2 ,"
        sql = sql & " deacGrupo3 ,"
        sql = sql & " deacGrupo4 ,"
        sql = sql & " codbusetaAM, busetaAM, codbusetaPM, busetaPM"
            
        sql = sql & " FROM " & _
        " PcProvCli inner join TMP4 vwConsPCCartera ON PcProvCli.IdProvCli = vwConsPCCartera.IdProvCli " & _
        " left join tmp8 pckn on pckn.idpckardex = vwConsPCCartera.id "
        sql = sql & Condicion
        'sql = sql & " ORDER BY vwConsPCCartera.CodProvCli,vwConsPCCartera.Fechavenci "
        sql = sql & " ORDER BY codbusetaAM, Estudiante,vwConsPCCartera.Fechavenci "
        
        
End With
        grd.Redraw = False
        MensajeStatus MSG_PREPARA, vbHourglass
        MiGetRowsRep gobjMain.EmpresaActual.OpenRecordset(sql), grd
        MensajeStatus "Listo", vbNormal
End Sub


Private Sub ConfigColsEstudiante()
    With grd
        .FormatString = "^#|tid|<Nombre Estudiante|<Nombre Factura|<Trans|<Doc Ref|<FechaEmision|<FechaVenci|^Dias Ven|>Valor|>Saldo|^1∫ Noti|<Fecha 1∫ Noti|^2∫ Noti|<Fecha 2∫ Noti|^3∫ Noti|<Fecha 3∫ Noti|^Imp|^Sec"
        .ColWidth(COL_NUMFILA) = 500
        .ColWidth(COL_TID) = 0
        .ColWidth(COL_CODPROVCLI) = 3500
        .ColWidth(COL_NOMBRE) = 3500
        .ColWidth(COL_TRANS) = 800
        .ColWidth(COL_NUMDOCREF) = 0
        .ColWidth(COL_FEMISION) = 1000
        .ColWidth(COL_FVENCI) = 1000
        .ColWidth(COL_VALOR) = 1000
        .ColWidth(COL_SALDO) = 1000
        .ColWidth(COL_BANDNOTI1) = 1000
        .ColWidth(COL_BANDNOTI2) = 1000
        .ColWidth(COL_BANDNOTI3) = 1000
        .ColWidth(COL_FECHANOTI1) = 1200
        .ColWidth(COL_FECHANOTI2) = 1200
        .ColWidth(COL_FECHANOTI3) = 1200
        .ColFormat(COL_VALOR) = gobjMain.EmpresaActual.GNOpcion.FormatoCantidad
        .ColFormat(COL_SALDO) = gobjMain.EmpresaActual.GNOpcion.FormatoCantidad
        .ColFormat(COL_FECHANOTI1) = gobjMain.EmpresaActual.GNOpcion.FormatoFecha
        .ColFormat(COL_FECHANOTI2) = gobjMain.EmpresaActual.GNOpcion.FormatoFecha
        .ColFormat(COL_FECHANOTI3) = gobjMain.EmpresaActual.GNOpcion.FormatoFecha
        
        .ColHidden(COL_TID) = True
'        .ColHidden(COL_IMP) = True
        .ColHidden(COL_SEC) = True
         
        .ColDataType(COL_FEMISION) = flexDTDate
        .ColDataType(COL_FVENCI) = flexDTDate
        .ColDataType(COL_FECHANOTI1) = flexDTDate
        .ColDataType(COL_FECHANOTI2) = flexDTDate
        .ColDataType(COL_FECHANOTI3) = flexDTDate
        
        .ColDataType(COL_BANDNOTI1) = flexDTBoolean
        .ColDataType(COL_BANDNOTI2) = flexDTBoolean
        .ColDataType(COL_BANDNOTI3) = flexDTBoolean
        .ColDataType(COL_IMP) = flexDTBoolean
        
        GNPoneNumFila grd, False
        grd.subtotal flexSTSum, 2, COL_VALOR, , grd.GridColor, , , "Subtotal", 2, True
        grd.subtotal flexSTSum, 2, COL_SALDO, , grd.GridColor, , , "Subtotal", 2, True
        grd.subtotal flexSTSum, -1, COL_VALOR, , grd.BackColorSel, vbYellow, , "Total", -1, True
        grd.subtotal flexSTSum, -1, COL_SALDO, , grd.BackColorSel, vbYellow, , "Total", -1, True
    End With
End Sub


