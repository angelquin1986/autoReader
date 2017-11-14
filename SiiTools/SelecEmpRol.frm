VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmSelecEmpRol 
   Caption         =   "Selección de empresa"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   6585
   StartUpPosition =   1  'CenterOwner
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Align           =   1  'Align Top
      Height          =   1812
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6588
      _cx             =   11620
      _cy             =   3196
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
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
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
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.PictureBox picBoton 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   408
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   6585
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2328
      Width           =   6588
      Begin VB.PictureBox pic1 
         BorderStyle     =   0  'None
         Height          =   440
         Left            =   1664
         ScaleHeight     =   435
         ScaleWidth      =   3255
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   3260
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "&Aceptar"
            Default         =   -1  'True
            Height          =   372
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   1332
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "&Cancelar"
            Height          =   372
            Left            =   1920
            TabIndex        =   4
            Top             =   0
            Width           =   1332
         End
      End
   End
End
Attribute VB_Name = "frmSelecEmpRol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAceptar_Click()
    Dim cod As String
    On Error GoTo ErrTrap
    
    If grd.Row < grd.FixedRows Then Exit Sub    '*** MAKOTO 08/sep/00
    cod = grd.TextMatrix(grd.Row, 0)
    If Len(cod) = 0 Then Exit Sub
    
    If AbrirEmpresaSii(cod, True) Then Unload Me
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim v As Variant
    
    MainRoles 'CARGA BASES DE ROLES
    grd.Rows = grd.FixedRows
    'v = gobjRol.ListaEmpresas(True, False)
    'grd.Rows = grd.FixedRows
    grd.LoadArray gobjRol.ListaEmpresas(True, False)
    
    If Not IsEmpty(v) Then
        grd.LoadArray v
    End If
    With grd
            .FormatString = "<Código|<Empresa|<Tipo|<Ruta|<Servidor|<Device|<Archivo"
        .ColWidth(0) = 1000     'Codigo
        .ColWidth(1) = 2000     'Descripcion
        .ColWidth(2) = 0        'Tipo
#If DAOLIB Then                             '*** MAKOTO 30/jun/2000
        .ColWidth(3) = 2000     'Ruta
        .ColWidth(4) = 0        'Servidor
#Else                                       '*** MAKOTO 30/jun/2000
        .ColWidth(3) = 0        'Ruta
        .ColWidth(4) = 1500     'Servidor
#End If
        .ColWidth(5) = 0        'Device
        .ColWidth(6) = 1600     'NombreDB
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    'Alto del grid
    grd.Height = Me.ScaleHeight - picBoton.Height
    
    'Centra los botones
    pic1.Left = (Me.ScaleWidth - pic1.Width) / 2
End Sub

Private Sub grd_DblClick()
    cmdAceptar_Click
End Sub
