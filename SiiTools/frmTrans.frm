VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmTrans 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Transacciones"
   ClientHeight    =   3840
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   6240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   924
      Left            =   0
      ScaleHeight     =   930
      ScaleWidth      =   6240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2916
      Width           =   6240
      Begin VB.PictureBox picBotones 
         BorderStyle     =   0  'None
         Height          =   492
         Left            =   1368
         ScaleHeight     =   495
         ScaleWidth      =   3495
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   444
         Width           =   3492
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "&Cancelar"
            Height          =   372
            Left            =   2160
            TabIndex        =   4
            Top             =   0
            Width           =   1332
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "&Aceptar"
            Default         =   -1  'True
            Height          =   372
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   1332
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Mentenga presionada la tecla CTRL para seleccionar varias filas.  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   5844
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Align           =   1  'Align Top
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6240
      _cx             =   11007
      _cy             =   4683
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
      FocusRect       =   4
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   3
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
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmTrans.frx":0000
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "frmTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mAceptado As Boolean

'Devuelve transacciones seleccionadas como cadena separado por coma
Public Function Seleccionar( _
                    ByRef cad As String) As Boolean
    mAceptado = False
    CargarTrans
    Visualizar cad, grd.ColIndex("Código")
    Me.Show vbModal
    
    If mAceptado Then
        'Si está esleccionado todas, es lo mismo que no seleccionar ningúna
        If grd.SelectedRows = grd.Rows - grd.FixedRows Then
            cad = ""
        Else
            cad = FlexCodigosSeleccionados(grd, grd.ColIndex("Código"), True)
        End If
    End If
    
    Seleccionar = mAceptado
    Unload Me
End Function

Public Function SeleccionarCat( _
                    ByRef cad As String) As Boolean
    mAceptado = False
    CargarCatalogo
    Visualizar cad, grd.ColIndex("Tabla")
    Me.Show vbModal
    
    If mAceptado Then
        'Si está esleccionado todas, es lo mismo que no seleccionar ningúna
        If grd.SelectedRows = grd.Rows - grd.FixedRows Then
            cad = ""
        Else
            cad = FlexCodigosSeleccionados(grd, grd.ColIndex("Tabla"), True)
        End If
    End If
    
    SeleccionarCat = mAceptado
    Unload Me
End Function


Private Sub Visualizar(ByVal cad As String, Columna As Long)
    Dim i As Long, v As Variant, j As Long, cod As String
    
    v = Split(cad, ",")
    If UBound(v, 1) < 0 Then Exit Sub
    
    With grd
        For i = .FixedRows To .Rows - 1
            For j = LBound(v, 1) To UBound(v, 1)
                cod = Trim$(v(j))                   'Quita espacios del extremo
                cod = Right$(cod, Len(cod) - 1)     'Quita primer "'"
                cod = Left$(cod, Len(cod) - 1)      'Quita ultimo "'"
                If Trim$(.TextMatrix(i, Columna)) = Trim$(cod) Then
                    .IsSelected(i) = True
                    Exit For
                End If
            Next j
        Next i
    End With
End Sub

Private Sub cmdAceptar_Click()
    mAceptado = True
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    mAceptado = False
    Me.Hide
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    grd.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - pic1.Height
    picbotones.Move (Me.ScaleWidth - picbotones.Width) / 2
End Sub


Private Sub CargarTrans()
    With grd
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .FormatString = "^|<Código|<Descripción"
        .LoadArray gobjMain.EmpresaActual.ListaGNTrans("", False, False)
        AsignarTituloAColKey grd
        
        GNPoneNumFila grd, False
        AjustarAutoSize grd, -1, -1
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub CargarCatalogo()
    CargarCatalogos grd
End Sub


