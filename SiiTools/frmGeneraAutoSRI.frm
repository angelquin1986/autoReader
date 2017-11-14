VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIVGenAutoSRI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generación de Archivo de Autorizaciones SRI"
   ClientHeight    =   7035
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   8775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin FlexComboProy.FlexCombo fcbTipoTramite 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   2100
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   556
      ColWidth1       =   3400
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
   Begin VB.Frame Frame1 
      Caption         =   "Datos Anteriores"
      Height          =   675
      Left            =   240
      TabIndex        =   22
      Top             =   1020
      Width           =   8115
      Begin VB.Label lblAutorizaOld 
         Caption         =   "Autorizacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   1275
         TabIndex        =   26
         Top             =   240
         Width           =   1845
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Autorización:"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lblFechaHastaOld 
         Caption         =   "06/11/2008"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   4500
         TabIndex        =   24
         Top             =   240
         Width           =   1845
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Valida Hasta"
         Height          =   195
         Left            =   3435
         TabIndex        =   23
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.TextBox txtDestino 
      Height          =   320
      Left            =   960
      TabIndex        =   10
      Top             =   6120
      Width           =   7215
   End
   Begin VB.CommandButton cmdExplorar 
      Caption         =   "..."
      Height          =   310
      Left            =   8220
      TabIndex        =   8
      Top             =   6120
      Width           =   372
   End
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   8775
      TabIndex        =   4
      Top             =   6495
      Width           =   8775
      Begin VB.PictureBox picBotones 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2760
         ScaleHeight     =   375
         ScaleWidth      =   3075
         TabIndex        =   5
         Top             =   60
         Width           =   3075
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "&Aceptar"
            Default         =   -1  'True
            Height          =   372
            Left            =   120
            TabIndex        =   2
            Top             =   0
            Width           =   1332
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "&Cancelar"
            Height          =   372
            Left            =   1560
            TabIndex        =   3
            Top             =   0
            Width           =   1332
         End
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdTrans 
      Height          =   3195
      Left            =   240
      TabIndex        =   1
      Top             =   2820
      Width           =   8400
      _cx             =   14817
      _cy             =   5636
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   5000
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
      ShowComboButton =   0   'False
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   0
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   8340
      Top             =   60
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "mdb"
      DialogTitle     =   "Destino de exportación"
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   330
      Left            =   7260
      TabIndex        =   15
      Top             =   2100
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Format          =   106692609
      CurrentDate     =   36902
   End
   Begin VSFlex7LCtl.VSFlexGrid grdcomp 
      Height          =   1695
      Left            =   5760
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   4320
      Width           =   2640
      _cx             =   4657
      _cy             =   2990
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   5000
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
      ShowComboButton =   0   'False
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   0
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7LCtl.VSFlexGrid grdPunto 
      Height          =   1695
      Left            =   3060
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   4320
      Width           =   2640
      _cx             =   4657
      _cy             =   2990
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   5000
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
      ShowComboButton =   0   'False
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   0
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label lblFechaActual 
      Caption         =   "06/11/2008"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   21
      Top             =   1800
      Width           =   1485
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Valida Hasta"
      Height          =   195
      Left            =   3720
      TabIndex        =   20
      Top             =   780
      Width           =   900
   End
   Begin VB.Label lblFechaHasta 
      Caption         =   "06/11/2008"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   4740
      TabIndex        =   19
      Top             =   720
      Width           =   1845
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Autorización:"
      Height          =   195
      Left            =   525
      TabIndex        =   18
      Top             =   720
      Width           =   915
   End
   Begin VB.Label lblAutoriza 
      Caption         =   "Autorizacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   1560
      TabIndex        =   17
      Top             =   720
      Width           =   1845
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Tramite:"
      Height          =   195
      Index           =   2
      Left            =   300
      TabIndex        =   16
      Top             =   1800
      Width           =   1065
   End
   Begin VB.Label lblRUC 
      Caption         =   "RUC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   1560
      TabIndex        =   14
      Top             =   420
      Width           =   1845
   End
   Begin VB.Label lblRazonSocial 
      Caption         =   "Razón Social"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   1560
      TabIndex        =   13
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "RUC:"
      Height          =   195
      Left            =   1050
      TabIndex        =   12
      Top             =   420
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Razón Social:"
      Height          =   195
      Index           =   0
      Left            =   450
      TabIndex        =   11
      Top             =   120
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Destino  "
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   9
      Top             =   6120
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Trámite:"
      Height          =   195
      Index           =   1
      Left            =   300
      TabIndex        =   7
      Top             =   2160
      Width           =   1155
   End
   Begin VB.Label lblG1 
      Caption         =   "Transacciones"
      Height          =   255
      Left            =   300
      TabIndex        =   6
      Top             =   2460
      Width           =   1215
   End
End
Attribute VB_Name = "frmIVGenAutoSRI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Const COL_CHEK = 1
'Const COL_DESC = 2
'Const COL_IDGRUPO = 3
'Const COL_TAM = 4
'Const COL_POS = 5

Const COL_CHEK = 1
Const COL_CODTRANS = 2
Const COL_TIPOTRANS = 3
Const COL_SERESTAB = 4
Const COL_SERPUNTO = 5
Const COL_AUTORIZACION = 6
Const COL_FECHACADUCIDAD = 7
Const COL_SECUENCIA = 8
Const COL_CODTIPOCOM = 9
Const COL_NUMTRANSULTIMO = 10
Const COL_AUTORIZACIONOLD = 11
Const COL_FECHACADUCIDADOLD = 12
Const COL_NUMTRANINICIO = 13
Const COL_NUMTRANINICIOOLD = 14
Const COL_CODSUCURSAL = 15


Private mAceptado As Boolean
Dim suma As Integer, tamprefi As Integer
Dim nombre As String
Dim BandAnulaTodosComprobante As Boolean

Public Sub Inicio()
    
'''    CargarTrans
    ConfigColsHistorialAutorizaciones
    Visualizar
    lblRazonSocial.Caption = gobjMain.EmpresaActual.GNOpcion.NombreEmpresa
    lblRUC.Caption = gobjMain.EmpresaActual.GNOpcion.ruc
    lblAutoriza.Caption = gobjMain.EmpresaActual.GNOpcion.NumAutorizacion_AutoImp
    lblFechaHasta.Caption = Format(gobjMain.EmpresaActual.GNOpcion.FechaCaducidad_AutoImp, "mmm/yyyy")
    lblAutorizaOld.Caption = gobjMain.EmpresaActual.GNOpcion.NumAutorizacion_AutoImpOld
    lblFechaHastaOld.Caption = Format(gobjMain.EmpresaActual.GNOpcion.FechaCaducidad_AutoImpOld, "mmm/yyyy")
    
    lblFechaActual.Caption = Format(Date, "dd/mm/yyyy")
    dtpFecha.Visible = False
    dtpFecha.value = Date
    Me.Show vbModal
    If mAceptado Then
    End If
    
  
    Unload Me
End Sub




Private Sub Visualizar()
    Dim i As Long, v As Variant, j As Long, cod As String, vv As Variant
    Dim fila As Long, ant As String

    With grdTrans
'        For i = 0 To UBound(v)
'            .Select v(i), COL_CHEK
'            .Text = "-1"
'        Next i
'        GNPoneNumFila grdTrans, False
    End With
End Sub




'Private Sub cboTipoTramite_Click()
'    CargarTrans
'End Sub


Private Sub cmdAceptar_Click()
    If Generar Then
        mAceptado = True
        Me.Hide
    End If
End Sub


Private Sub cmdCancelar_Click()
    mAceptado = False
    Me.Hide
End Sub


Private Sub cmdExplorar_Click()
    On Error GoTo ErrTrap
    
    With dlg1
        If Len(.filename) = 0 Then
            '***Diego 25/09/2003 cambio  para VATEX
            .InitDir = App.Path
'            .InitDir = txtDestino.Text
        Else
            .InitDir = .filename
        End If
        .flags = cdlOFNPathMustExist
        .Filter = "Archivos xml (*.xml)|*.xml|Predefinido " & _
                  "|Todos (*.*)|*.*"
        .ShowSave
        txtDestino.Text = .filename
        nombre = .FileTitle
        
        
'        .FileTitle
 '       .Name
    End With
    
    Exit Sub
ErrTrap:
    If Err.Number <> 32755 Then
        DispErr
    End If
    Exit Sub
End Sub

Private Sub fcbTipoTramite_Selected(ByVal Text As String, ByVal KeyText As String)
    CargarTrans
End Sub

Private Sub Form_Load()
    fcbTipoTramite.SetData ListaTramitesSRI
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picbotones.Move (Me.ScaleWidth - picbotones.Width) / 2
End Sub


Private Sub CargarTrans()
    Dim sql As String, rs As Recordset, i As Integer, sqlAux As String, j As Integer, k As Integer
    Dim gnsuc  As GNSucursal
    Dim rsSuc As Recordset, codbase As String
    Dim v As Variant, contbase As Integer, W As Variant, contTrans As Integer
    With grdTrans
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If fcbTipoTramite.KeyText = "" Then
            MsgBox "No seleccionó tipo de tramite"
            Exit Sub
        End If
        ConfigColsHistorialAutorizaciones
        AsignarTituloAColKey grdTrans
        'Agrega una columna para CheckBox
  
        
        .ColDataType(.ColIndex("Sel")) = flexDTBoolean

      
        
        sqlAux = "SELECT * FROM GnSucursal WHERE bandvalida=1 and numpuntos>0 "
        Set rsSuc = gobjMain.EmpresaActual.OpenRecordset(sqlAux)
        
        sql = ""
        ReDim v(10, 3)
        contbase = 0
        For i = 1 To rsSuc.RecordCount
            Set gnsuc = gobjMain.EmpresaActual.RecuperaGNSucursal(rsSuc.Fields("CodSucursal"))
            If Len(gnsuc.BaseDatos) > 0 Then
                v(contbase, 0) = gnsuc.BaseDatos
                v(contbase, 1) = gnsuc.Servidor
                v(contbase, 2) = gnsuc.CodSucursal
                contbase = contbase + 1
            Else
                If i = 1 Then
                    v(contbase, 0) = gobjMain.EmpresaActual.NombreDB
                    v(contbase, 1) = gobjMain.EmpresaActual.Server
                    v(contbase, 2) = gobjMain.EmpresaActual.CodEmpresa
                    contbase = contbase + 1
                End If
            End If
            rsSuc.MoveNext
        Next i
        If rsSuc.RecordCount = 0 Then Exit Sub
        rsSuc.MoveFirst
        
        For i = 0 To 0 'contbase - 1
        
            Set gnsuc = gobjMain.EmpresaActual.RecuperaGNSucursal(rsSuc.Fields("CodSucursal"))
        
            If gnsuc.BandValida Then
                sql = sql & " SELECT  "
                If fcbTipoTramite.KeyText = "10" Then
                    sql = sql & " case when BandNuevoAutoImpresor=1 then 1 else 0 end as sel, "
                Else
                    sql = sql & " 0 as sel ,"
                End If
                sql = sql & " CodTrans, LEFT(ac.Descripcion,50) as descr,  "
                sql = sql & " NumSerieEstablecimiento, NumSeriePunto,  "
                sql = sql & " NumAutorizacion,   FechaCaducidad,  "
                sql = sql & " NumTransSiguiente,  "
                sql = sql & " case  when AnexoCodTipoTrans ='2' then CONVERT(varchar, TipoTrans)  "
                sql = sql & " else CONVERT(varchar, ac. CodComprobante) end as CodComprobante,"
                sql = sql & " NumTransUltimo , NumAutorizacionOld, FechaCaducidadOld, "
                sql = sql & " NumTransInicio , NumTransInicioOld, "
                sql = sql & "gns.codsucursal as  Sucursal, '"
                sql = sql & v(i, 0) & "' as Base "
                sql = sql & " FROM gntrans G"
                sql = sql & " Left JOIN Anexo_Comprobantes AC "
                sql = sql & " ON G.AnexoCodTipoComp=ac.CodComprobante"
                sql = sql & " Left JOIN GnSucursal GNS "
                sql = sql & " ON Gns.idsucursal=g.idsucursal"
                
                sql = sql & " where SUBSTRING(Opcion, 152, 1) = 'S'"
''                If fcbTipoTramite.KeyText = "10" Then
''                    sql = sql & " and BandNuevoAutoImpresor = 1"
''                End If
            
            End If
            
            rsSuc.MoveNext
            sql = sql & " union all "
        Next i
        sql = Mid$(sql, 1, Len(sql) - 11)
        sql = sql & " Order by CodComprobante, NumSerieEstablecimiento, NumSeriePunto"
        Set gnsuc = Nothing
        Set rsSuc = Nothing
        
        
        
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
        
        If rs.RecordCount > 0 Then
            While Not rs.EOF
                .AddItem vbTab & rs.Fields("sel") & vbTab & rs.Fields("CodTrans") & _
                vbTab & rs.Fields("descr") & _
                vbTab & rs.Fields("NumSerieEstablecimiento") & _
                vbTab & rs.Fields("NumSeriePunto") & _
                vbTab & rs.Fields("NumAutorizacion") & _
                vbTab & rs.Fields("FechaCaducidad") & _
                vbTab & rs.Fields("NumTransSiguiente") & _
                vbTab & rs.Fields("CodComprobante") & _
                vbTab & rs.Fields("NumTransUltimo") & _
                vbTab & rs.Fields("NumAutorizacionOld") & _
                vbTab & rs.Fields("FechaCaducidadOld") & _
                vbTab & rs.Fields("NumTransInicio") & _
                vbTab & rs.Fields("NumTransInicioOld") & _
                vbTab & rs.Fields("Sucursal") & _
                vbTab & rs.Fields("Base")
                
                
                rs.MoveNext
            Wend
        End If
        
        .Refresh
        .Redraw = flexRDBuffered

        

        'verifica si existe los comprobantes autoimpresos para cada sucursal
        
        ReDim W(10, 6)
        
        sql = " select CodComprobante from Anexo_Comprobantes where BandAutoImpresor=1 and BandAutorizaActiva=1"
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
        
        
    contTrans = 1
    grdcomp.Rows = contbase + 1
    grdcomp.Cols = rs.RecordCount + 1
    
    grdPunto.Rows = contbase + 1
    grdPunto.Cols = rs.RecordCount + 1
    
    If Not rs Is Nothing Then
        If Not rs.BOF Then
            rs.MoveFirst
        End If
        For i = 1 To rs.RecordCount
            grdcomp.TextMatrix(0, i) = rs.Fields("CodComprobante")
            grdcomp.ColWidth(i) = 300
            grdPunto.TextMatrix(0, i) = rs.Fields("CodComprobante")
            grdPunto.ColWidth(i) = 300
        
            rs.MoveNext
        Next i
        For k = 0 To contbase - 1
            grdcomp.TextMatrix(k + 1, 0) = v(k, 2)
            grdPunto.TextMatrix(k + 1, 0) = v(k, 2)
        Next k
        
        
        For i = 1 To contbase
            For k = 1 To rs.RecordCount
                grdcomp.TextMatrix(i, k) = 0
            Next k
        Next i
            
        For i = 1 To contbase
            For k = 1 To rs.RecordCount
                grdPunto.TextMatrix(i, k) = 0
            Next k
        Next i
            
    End If
    
    
        For i = 1 To grdcomp.Rows
            For k = 1 To grdcomp.Cols
                For j = 1 To grdTrans.Rows - 1
                    If grdcomp.TextMatrix(i - 1, 0) = grdTrans.TextMatrix(j, COL_CODSUCURSAL) Then
                        If grdcomp.TextMatrix(0, k - 1) = grdTrans.TextMatrix(j, COL_CODTIPOCOM) Then
                            grdcomp.TextMatrix(i - 1, k - 1) = 1
                            Exit For
                        End If
                    End If
                Next j
            Next k
        Next i
    
        
        For i = 1 To grdPunto.Rows
            For k = 1 To grdPunto.Cols
                For j = 1 To grdTrans.Rows - 1
                    If grdPunto.TextMatrix(i - 1, 0) = grdTrans.TextMatrix(j, COL_CODSUCURSAL) Then
                        If grdPunto.TextMatrix(0, k - 1) = grdTrans.TextMatrix(j, COL_CODTIPOCOM) Then
                            grdPunto.TextMatrix(i - 1, k - 1) = grdPunto.ValueMatrix(i - 1, k - 1) + 1
'                            Exit For
                        End If
                    End If
                Next j
            Next k
        Next i

        Dim cont_suc As Integer
        
        cont_suc = 1

       grdPunto.Refresh
        grdPunto.Redraw = flexRDBuffered

    
       .Refresh

        .Redraw = flexRDBuffered
        .ColHidden(COL_TIPOTRANS) = True
        .ColHidden(8) = True
        .ColHidden(10) = True
        .ColHidden(11) = True
        .ColHidden(12) = True
        .ColHidden(13) = True
        .ColHidden(14) = True
        .ColHidden(16) = True

        
        Select Case CInt(fcbTipoTramite.KeyText)
            Case 6, 7, 8, 9: .ColHidden(1) = True
                .Editable = flexEDNone
            Case 10:
                .Editable = flexEDNone
            Case 11: .ColHidden(1) = False
            .Editable = flexEDKbdMouse
        End Select
        
        
    End With
End Sub

Private Sub ConfigColsHistorialAutorizaciones()
    Dim s As String
    With grdTrans
        s = "^|^Sel|<Cod.Trans|<Tipo|^Num.Ser.Est"
        s = s & "|^Num.Ser.Punto|^No.Autorización|^Fecha Cad.|>NumTransSiguiente"
        s = s & "|^Cod Comp|>NumTransUltimo|^NumAutorizacionOld|^FechaCaducidadOld"
        s = s & "|>NumTransInicio|>NumTransInicioOld|<Sucursal|<Base Datos"
        .FormatString = s
        
        .ColWidth(0) = 200      '#
        .ColWidth(1) = 400      'sel
        .ColWidth(2) = 700     'CodTrans
        .ColWidth(3) = 2000     'Tipo Trans
        .ColWidth(4) = 800     '# esta
        .ColWidth(5) = 800     '# punto
        .ColWidth(6) = 1250     'Autorizacion
        .ColWidth(7) = 1100     'Fecha Caducidad
        .ColWidth(8) = 600     'TipoDocumento
        .ColWidth(COL_CODSUCURSAL) = 1200     'TipoDocumento
        .ColFormat(7) = "dd/mm/yyyy"
        .ColFormat(12) = "dd/mm/yyyy"
        .Refresh
    End With


End Sub

Private Function Generar() As Boolean
    Dim s As String, i As Integer, Selec As Integer, k As Integer
    Dim file As String, Cadena As String
    Dim NumFile As Integer
    Dim gnsuc  As GNSucursal
    On Error GoTo ErrTrap
    Generar = False
    Selec = 0
    'Verifica si está especificado el destino
    s = Trim$(txtDestino.Text)
    If Len(s) = 0 Then
        MsgBox "Debe especificar el archivo de destino.", vbInformation
        txtDestino.SetFocus
        Exit Function
    End If
    
    'Si aun no está hecho la seleccion del tramite
    If Len(fcbTipoTramite.KeyText) = 0 Then
        MsgBox "Debe especificar el tipo de Tramite.", vbInformation
        fcbTipoTramite.SetFocus
        Exit Function
    End If
    
    For i = 1 To grdTrans.Rows - 1
        If grdTrans.ValueMatrix(i, COL_CHEK) = "-1" Then
            Selec = Selec + 1
       End If
    Next i
    If CInt(fcbTipoTramite.KeyText) > 10 Then
        If Selec = 0 Then
            MsgBox "Debe seleccionar las Transacciones ", vbInformation
            grdTrans.SetFocus
            Exit Function
        End If
    End If
'    file = txtDestino.Text
'    If ExisteArchivo(file) Then
'        If MsgBox("El nombre del archivo " & nombre & " ya existe desea sobreescribirlo?", vbYesNo) = vbNo Then
'            Exit Function
'        End If
'    End If
'    NumFile = FreeFile
'    Open file For Output Access Write As #NumFile
        
    'verifica si faltan comprobantes en cada sucursal
        For i = 1 To grdcomp.Rows - 1
            For k = 1 To grdcomp.Cols - 1
                If grdcomp.ValueMatrix(i, k) = 0 Then
                    MsgBox "Falta el comprobante con código de comprobante " & grdcomp.TextMatrix(0, k) & " en la agencia: " & grdcomp.TextMatrix(i, 0)
                    MsgBox "El archivo NO se generó"
                    GoTo salida
                End If
            Next k
        Next i
        
    'verifica si faltan comprobantes en cada sucursal
        Dim max As Integer
        
        For i = 1 To grdPunto.Rows - 1
            Set gnsuc = gobjMain.EmpresaActual.RecuperaGNSucursal(grdPunto.TextMatrix(i, 0))
            max = gnsuc.NumPuntos
            
            For k = 1 To grdPunto.Cols - 1
'                If k = 1 Then
'                    max = grdPunto.ValueMatrix(i, k)
'                Else
                    If grdPunto.ValueMatrix(i, k) <> max Then
                        MsgBox "La cantidad de puntos de emisión para la sucursal " & grdPunto.TextMatrix(i, 0) & " es " & max & Chr(13) & "La cantidad de puntos configurados con el comprobante " & grdcomp.TextMatrix(0, k) & " es " & grdPunto.TextMatrix(i, k)
                        MsgBox "El archivo NO se generó"
                        GoTo salida
'                    End If
                End If
            Next k
        Next i
    
        
        
    Select Case CInt(fcbTipoTramite.KeyText)
     Case 6:
        Cadena = GeneraArchivo_AutorizacionNueva
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "TramitesPosiblesSRI", "B,C,D,E,F"
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "RealizarReporteRangos", "0"
        gobjMain.EmpresaActual.GNOpcion.Grabar
     Case 7:
        Cadena = GeneraArchivo_CambioSoftware
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "TramitesPosiblesSRI", "B,C,D,E,F"
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "RealizarReporteRangos", "0"
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "ValidacionAutoimpresores", "0"
        gobjMain.EmpresaActual.GNOpcion.Grabar
     Case 8:
        Cadena = GeneraArchivo_Renovacion
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "TramitesPosiblesSRI", "B,C,D,E,F"
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "RealizarReporteRangos", "0"
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "ValidacionAutoimpresores", "0"
        gobjMain.EmpresaActual.GNOpcion.Grabar
     Case 9:
        Cadena = GeneraArchivo_Baja
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "TramitesPosiblesSRI", "A"
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "RealizarReporteRangos", "0"
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "ValidacionAutoimpresores", "0"
        gobjMain.EmpresaActual.GNOpcion.Grabar
     
     Case 10:
        Cadena = GeneraArchivo_Inclucion
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "TramitesPosiblesSRI", "B,C,D,E,F"
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "RealizarReporteRangos", "0"
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "ValidacionAutoimpresores", "0"
        gobjMain.EmpresaActual.GNOpcion.Grabar
     
     Case 11:
        Cadena = GeneraArchivo_Eliminacion
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "ValidacionAutoimpresores", "0"
        gobjMain.EmpresaActual.GNOpcion.Grabar
    End Select
    
    If Len(Cadena) > 0 Then
        MsgBox "El archivo se ha generado con éxito"
        Generar = True
        Exit Function
    Else
        MsgBox "El archivo NO se generó"
    End If
    

salida:
    Generar = False
    Exit Function
ErrTrap:
    DispErr
    GoTo salida
End Function


Private Function GeneraArchivo_AutorizacionNueva() As String
    Dim obj As GNOpcion, cad As String, i As Integer
    Dim aut As String
    Dim file As String, Cadena As String
    Dim NumFile As Integer, Num As Long
    Dim sql As String, rs As Recordset

    On Error GoTo ErrTrap
    file = txtDestino.Text
    If ExisteArchivo(file) Then
        If MsgBox("El nombre del archivo " & nombre & " ya existe desea sobreescribirlo?", vbYesNo) = vbNo Then
            Exit Function
        End If
    End If
    NumFile = FreeFile
    Open file For Output Access Write As #NumFile
    
    cad = "<?xml version=" & """1.0""" & " encoding=" & """UTF-8""" & "?>"
    Print #NumFile, cad
    cad = "<autorizacion>"
    Print #NumFile, cad
    cad = "<codTipoTra>6</codTipoTra>"
    Print #NumFile, cad
    cad = "<ruc>" & Trim$(lblRUC.Caption) & "</ruc>"
    Print #NumFile, cad
    For i = 1 To grdTrans.Rows - 1
            cad = "<numAut>" & Trim$(lblAutoriza.Caption) & "</numAut>"
            Print #NumFile, cad
            i = grdTrans.Rows - 1
    Next i
    cad = "<fecha>" & Format(dtpFecha.value, "dd/mm/yyyy") & "</fecha>"
    Print #NumFile, cad
    cad = "<detalles>"
    Print #NumFile, cad
    For i = 1 To grdTrans.Rows - 1
            cad = "  <detalle>"
            Print #NumFile, cad
            If Len(grdTrans.TextMatrix(i, 9)) <> 0 Then
                cad = "<codDoc>" & grdTrans.TextMatrix(i, COL_CODTIPOCOM) & "</codDoc>"
                Print #NumFile, cad
            Else
                MsgBox "Debe especificar el tipo de Documento en la TRansaccion." & grdTrans.TextMatrix(i, 2), vbInformation
                cad = ""
            End If
            cad = "<estab>" & grdTrans.TextMatrix(i, COL_SERESTAB) & "</estab>"
            Print #NumFile, cad
            cad = "<ptoEmi>" & grdTrans.TextMatrix(i, COL_SERPUNTO) & "</ptoEmi>"
            Print #NumFile, cad
            Num = gobjMain.EmpresaActual.RecuperaNumeroTransaccionReporteRango(grdTrans.TextMatrix(i, COL_CODTRANS), grdTrans.TextMatrix(i, COL_AUTORIZACION), False)
            cad = "<inicio>" & Num & "</inicio>"
            Print #NumFile, cad
            cad = "  </detalle>"
            Print #NumFile, cad
        'End If
        
            sql = " Update Gntrans "
            sql = sql & " set BandNuevoAutoImpresor=0 "
            sql = sql & " where codtrans='" & grdTrans.TextMatrix(i, COL_CODTRANS) & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
        
        
    Next i
    cad = "  </detalles>"
    Print #NumFile, cad
    cad = "</autorizacion>"
    Print #NumFile, cad
    Close NumFile
salida:
    GeneraArchivo_AutorizacionNueva = False
    Exit Function
ErrTrap:
    DispErr
    GoTo salida

End Function

Private Function GeneraArchivo_CambioSoftware() As String
    Dim obj As GNOpcion, cad As String, i As Integer
    Dim aut As String
    Dim file As String, Cadena As String
    Dim NumFile As Integer, Num As Long

    On Error GoTo ErrTrap
    file = txtDestino.Text
    If ExisteArchivo(file) Then
        If MsgBox("El nombre del archivo " & nombre & " ya existe desea sobreescribirlo?", vbYesNo) = vbNo Then
            Exit Function
        End If
    End If
    NumFile = FreeFile
    Open file For Output Access Write As #NumFile
    cad = "<?xml version=" & """1.0""" & " encoding=" & """UTF-8""" & "?>"
    Print #NumFile, cad
    cad = "<autorizacion>"
    Print #NumFile, cad
    cad = "<codTipoTra>7</codTipoTra>"
    Print #NumFile, cad
    cad = "<ruc>" & Trim$(lblRUC.Caption) & "</ruc>"
    Print #NumFile, cad
    cad = "<fecha>" & Format(dtpFecha.value, "dd/mm/yyyy") & "</fecha>"
    Print #NumFile, cad
    cad = "<autOld>" & Trim$(lblAutorizaOld.Caption) & "</autOld>"
    Print #NumFile, cad
    cad = "<autNew>" & Trim$(lblAutoriza.Caption) & "</autNew>"
    Print #NumFile, cad
   cad = "<detalles>"
    Print #NumFile, cad
    For i = 1 To grdTrans.Rows - 1
            cad = "  <detalle>"
            Print #NumFile, cad
            If Len(grdTrans.TextMatrix(i, 9)) <> 0 Then
                cad = "<codDoc>" & grdTrans.TextMatrix(i, COL_CODTIPOCOM) & "</codDoc>"
                Print #NumFile, cad
            Else
                MsgBox "Debe especificar el tipo de Documento en la TRansaccion." & grdTrans.TextMatrix(i, 2), vbInformation
                cad = ""
            End If
            cad = "<estab>" & grdTrans.TextMatrix(i, COL_SERESTAB) & "</estab>"
            Print #NumFile, cad
            cad = "<ptoEmi>" & grdTrans.TextMatrix(i, COL_SERPUNTO) & "</ptoEmi>"
            Print #NumFile, cad
            Num = gobjMain.EmpresaActual.RecuperaNumeroTransaccionReporteRango(grdTrans.TextMatrix(i, COL_CODTRANS), lblAutorizaOld.Caption, True)
            cad = "<finOld>" & Num & "</finOld>"
            Print #NumFile, cad
            Num = gobjMain.EmpresaActual.RecuperaNumeroTransaccionReporteRango(grdTrans.TextMatrix(i, COL_CODTRANS), lblAutoriza.Caption, False)
            cad = "<iniNew>" & Num & "</iniNew>"
            Print #NumFile, cad
            cad = "  </detalle>"
            Print #NumFile, cad
    Next i
    cad = "  </detalles>"
    Print #NumFile, cad
    cad = "</autorizacion>"
    Print #NumFile, cad
        Close NumFile
salida:
    GeneraArchivo_CambioSoftware = False
    Exit Function
ErrTrap:
    DispErr
    GoTo salida
End Function

Private Function GeneraArchivo_Renovacion() As String
    Dim obj As GNOpcion, cad As String, i As Integer
    Dim aut As String
    Dim file As String, Cadena As String
    Dim NumFile As Integer, Num As Long

    On Error GoTo ErrTrap
    file = txtDestino.Text
    If ExisteArchivo(file) Then
        If MsgBox("El nombre del archivo " & nombre & " ya existe desea sobreescribirlo?", vbYesNo) = vbNo Then
            Exit Function
        End If
    End If
    NumFile = FreeFile
    Open file For Output Access Write As #NumFile
    cad = "<?xml version=" & """1.0""" & " encoding=" & """UTF-8""" & "?>"
    Print #NumFile, cad
    cad = "<autorizacion>"
    Print #NumFile, cad
    cad = "<codTipoTra>8</codTipoTra>"
    Print #NumFile, cad
    cad = "<ruc>" & Trim$(lblRUC.Caption) & "</ruc>"
    Print #NumFile, cad
    cad = "<fecha>" & Format(dtpFecha.value, "dd/mm/yyyy") & "</fecha>"
    Print #NumFile, cad
    cad = "<autOld>" & Trim$(lblAutorizaOld.Caption) & "</autOld>"
    Print #NumFile, cad
    cad = "<autNew>" & Trim$(lblAutoriza.Caption) & "</autNew>"
    Print #NumFile, cad
    cad = "<detalles>"
    Print #NumFile, cad
    For i = 1 To grdTrans.Rows - 1
            cad = "  <detalle>"
            Print #NumFile, cad
            If Len(grdTrans.TextMatrix(i, 9)) <> 0 Then
                cad = "<codDoc>" & grdTrans.TextMatrix(i, COL_CODTIPOCOM) & "</codDoc>"
                Print #NumFile, cad
            Else
                MsgBox "Debe especificar el tipo de Documento en la Transaccion." & grdTrans.TextMatrix(i, 2), vbInformation
                cad = ""
            End If
            cad = "<estab>" & grdTrans.TextMatrix(i, COL_SERESTAB) & "</estab>"
            Print #NumFile, cad
            cad = "<ptoEmi>" & grdTrans.TextMatrix(i, COL_SERPUNTO) & "</ptoEmi>"
            Print #NumFile, cad
            Num = gobjMain.EmpresaActual.RecuperaNumeroTransaccionReporteRango(grdTrans.TextMatrix(i, COL_CODTRANS), lblAutorizaOld.Caption, True)
            cad = "<finOld>" & Num & "</finOld>"
            Print #NumFile, cad
            Num = gobjMain.EmpresaActual.RecuperaNumeroTransaccionReporteRango(grdTrans.TextMatrix(i, COL_CODTRANS), grdTrans.TextMatrix(i, COL_AUTORIZACION), False)
            cad = "<iniNew>" & Num & "</iniNew>"
            Print #NumFile, cad
            cad = "  </detalle>"
            Print #NumFile, cad
    Next i
    cad = "  </detalles>"
    Print #NumFile, cad
    cad = "</autorizacion>"
    Print #NumFile, cad
        Close NumFile
salida:
    GeneraArchivo_Renovacion = False
    Exit Function
ErrTrap:
    DispErr
    GoTo salida
End Function

Private Function GeneraArchivo_Baja() As String
    Dim obj As GNOpcion, cad As String, i As Integer
    Dim aut As String
    Dim file As String, Cadena As String
    Dim NumFile As Integer, Num As Long
    Dim sql As String, rs As Recordset
    Dim fecha As Date
    On Error GoTo ErrTrap
    file = txtDestino.Text
    If ExisteArchivo(file) Then
        If MsgBox("El nombre del archivo " & nombre & " ya existe desea sobreescribirlo?", vbYesNo) = vbNo Then
            Exit Function
        End If
    End If
    NumFile = FreeFile
    Open file For Output Access Write As #NumFile

    fecha = gobjMain.EmpresaActual.GNOpcion.FechaCaducidad_AutoImp
    cad = "<?xml version=" & """1.0""" & " encoding=" & """UTF-8""" & "?>"
    Print #NumFile, cad
    cad = "<autorizacion>"
    Print #NumFile, cad
    cad = "<codTipoTra>9</codTipoTra>"
    Print #NumFile, cad
    cad = "<ruc>" & Trim$(lblRUC.Caption) & "</ruc>"
    Print #NumFile, cad
    cad = "<numAut>" & Trim$(lblAutoriza.Caption) & "</numAut>"
    Print #NumFile, cad
    
    If dtpFecha.value > fecha Then
        cad = "<fecha>" & Format(fecha, "dd/mm/yyyy") & "</fecha>"
    Else
        cad = "<fecha>" & Format(dtpFecha.value, "dd/mm/yyyy") & "</fecha>"
    End If
    Print #NumFile, cad
    cad = "<detalles>"
    Print #NumFile, cad
    For i = 1 To grdTrans.Rows - 1
            cad = "  <detalle>"
            Print #NumFile, cad
            If Len(grdTrans.TextMatrix(i, 9)) <> 0 Then
                cad = "<codDoc>" & grdTrans.TextMatrix(i, COL_CODTIPOCOM) & "</codDoc>"
                Print #NumFile, cad
            Else
                MsgBox "Debe especificar el tipo de Documento en la TRansaccion." & grdTrans.TextMatrix(i, 2), vbInformation
                cad = ""
            End If
            cad = "<estab>" & grdTrans.TextMatrix(i, COL_SERESTAB) & "</estab>"
            Print #NumFile, cad
            cad = "<ptoEmi>" & grdTrans.TextMatrix(i, COL_SERPUNTO) & "</ptoEmi>"
            Print #NumFile, cad
            Num = gobjMain.EmpresaActual.RecuperaNumeroTransaccionReporteRango(grdTrans.TextMatrix(i, COL_CODTRANS), lblAutoriza.Caption, True)
            cad = "<fin>" & Num & "</fin>"
            Print #NumFile, cad
            cad = "  </detalle>"
            Print #NumFile, cad
            
            sql = " select opcion from Gntrans "
            sql = sql & " where codtrans='" & grdTrans.TextMatrix(i, COL_CODTRANS) & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            
            sql = " Update Gntrans "
            sql = sql & " set opcion ='" & Mid$(rs.Fields("opcion"), 1, 151) & "N" & Mid$(rs.Fields("opcion"), 153, 50) & "', "
            sql = sql & " BandValida= 0 "
            sql = sql & " where codtrans='" & grdTrans.TextMatrix(i, COL_CODTRANS) & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
            
            If VerificaAnulacionComprobantes Then
                sql = " Update Anexo_Comprobantes "
                sql = sql & " set BandAutorizaActiva =0, BandAutoImpresor=0"
                sql = sql & " where CodComprobante='" & grdTrans.TextMatrix(i, COL_CODTIPOCOM) & "'"
                Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
            End If
            

    Next i
    cad = "  </detalles>"
    Print #NumFile, cad
    cad = "</autorizacion>"
    Print #NumFile, cad
    Close NumFile
    
    CargarTrans
    
    If VerificaAnulacionComprobantesEstablecimiento Then
            For i = 1 To grdPunto.Rows - 1
                If grdPunto.Cols > 1 Then
                        sql = " Update GNSucursal "
                        sql = sql & " set NumPuntos =" & grdPunto.TextMatrix(i, 1)
                        sql = sql & " where Codsucursal='" & grdPunto.TextMatrix(i, 0) & "'"
                        Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
                Else
                        sql = " Update GNSucursal "
                        sql = sql & " set NumPuntos = 0"
                        sql = sql & " where Codsucursal='" & grdPunto.TextMatrix(i, 0) & "'"
                        Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
                
                End If
            
            Next i
    End If
    
    
salida:
    GeneraArchivo_Baja = False
    Exit Function
ErrTrap:
    DispErr
    GoTo salida
End Function

Private Function GeneraArchivo_Inclucion() As String
    Dim obj As GNOpcion, cad As String, i As Integer
    Dim aut As String
    Dim file As String, Cadena As String
    Dim NumFile As Integer, Num As Long
    Dim sql As String, rs As Recordset, bandselec As Boolean
    On Error GoTo ErrTrap
    
    For i = 1 To grdTrans.Rows - 1
        If grdTrans.ValueMatrix(i, COL_CHEK) <> "0" Then
            bandselec = True
            i = grdTrans.Rows - 1
        End If
    Next i
    If Not bandselec Then
        MsgBox " No existen Puntos para incluir"
        GeneraArchivo_Inclucion = ""
        Exit Function
    End If
    
    
    file = txtDestino.Text
    If ExisteArchivo(file) Then
        If MsgBox("El nombre del archivo " & nombre & " ya existe desea sobreescribirlo?", vbYesNo) = vbNo Then
            Exit Function
        End If
    End If
    NumFile = FreeFile
    Open file For Output Access Write As #NumFile

    If Not VerificaAnulacionComprobantes Then
        If Not VerificaAnulacionComprobantesEstablecimiento Then
            MsgBox "Debe incluir todos los comprobantes del mismo tipo o del mismo establecimiento"
            Print #NumFile, cad
            Close NumFile
            Exit Function
        End If
    End If
    
    cad = "<?xml version=" & """1.0""" & " encoding=" & """UTF-8""" & "?>"
    Print #NumFile, cad
    cad = "<autorizacion>"
    Print #NumFile, cad
    cad = "<codTipoTra>10</codTipoTra>"
    Print #NumFile, cad
    cad = "<ruc>" & Trim$(lblRUC.Caption) & "</ruc>"
    Print #NumFile, cad
    cad = "<numAut>" & Trim$(lblAutoriza.Caption) & "</numAut>"
    Print #NumFile, cad
    
    cad = "<fecha>" & Format(dtpFecha.value, "dd/mm/yyyy") & "</fecha>"
    Print #NumFile, cad
    cad = "<detalles>"
    Print #NumFile, cad
    bandselec = False

    For i = 1 To grdTrans.Rows - 1
        If grdTrans.ValueMatrix(i, COL_CHEK) <> "0" Then
            cad = "  <detalle>"
            Print #NumFile, cad
            If Len(grdTrans.TextMatrix(i, 9)) <> 0 Then
                cad = "<codDoc>" & grdTrans.TextMatrix(i, COL_CODTIPOCOM) & "</codDoc>"
                Print #NumFile, cad
            Else
                MsgBox "Debe especificar el tipo de Documento en la TRansaccion." & grdTrans.TextMatrix(i, 2), vbInformation
                cad = ""
            End If
            cad = "<estab>" & grdTrans.TextMatrix(i, COL_SERESTAB) & "</estab>"
            Print #NumFile, cad
            cad = "<ptoEmi>" & grdTrans.TextMatrix(i, COL_SERPUNTO) & "</ptoEmi>"
            Print #NumFile, cad
            Num = gobjMain.EmpresaActual.RecuperaNumeroTransaccionReporteRango(grdTrans.TextMatrix(i, COL_CODTRANS), grdTrans.TextMatrix(i, COL_AUTORIZACION), False)
            cad = "<inicio>" & Num & "</inicio>"
            Print #NumFile, cad
            cad = "  </detalle>"
            Print #NumFile, cad
 
            sql = " select opcion from Gntrans "
            sql = sql & " where codtrans='" & grdTrans.TextMatrix(i, COL_CODTRANS) & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            
            sql = " Update Gntrans "
            sql = sql & " set BandNuevoAutoImpresor=0 "
            sql = sql & " where codtrans='" & grdTrans.TextMatrix(i, COL_CODTRANS) & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
        End If
    Next i
    cad = "  </detalles>"
    Print #NumFile, cad
    cad = "</autorizacion>"
    Print #NumFile, cad
    Close NumFile
salida:
    GeneraArchivo_Inclucion = False
    Exit Function
ErrTrap:
    DispErr
    GoTo salida
End Function


Private Function GeneraArchivo_Eliminacion() As String
    Dim obj As GNOpcion, cad As String, i As Integer, j As Integer, fila As Integer
    Dim aut As String
    Dim file As String, Cadena As String
    Dim NumFile As Integer, Num As Long
    Dim gnt As GNTrans, gnsuc As GNSucursal
    Dim sql As String, rs As Recordset
    On Error GoTo ErrTrap
    
    file = txtDestino.Text
    If ExisteArchivo(file) Then
        If MsgBox("El nombre del archivo " & nombre & " ya existe desea sobreescribirlo?", vbYesNo) = vbNo Then
            Exit Function
        End If
    End If
    NumFile = FreeFile
    Open file For Output Access Write As #NumFile

    If Not VerificaAnulacionComprobantes Then
        If Not VerificaAnulacionComprobantesEstablecimiento Then
            MsgBox "Debe eliminar todos los comprobantes del mismo tipo o del mismo establecimiento"
            Print #NumFile, cad
            Close NumFile
            Exit Function
        End If
    End If
    
    cad = "<?xml version=" & """1.0""" & " encoding=" & """UTF-8""" & "?>"
    Print #NumFile, cad
    cad = "<autorizacion>"
    Print #NumFile, cad
    cad = "<codTipoTra>11</codTipoTra>"
    Print #NumFile, cad
    cad = "<ruc>" & Trim$(lblRUC.Caption) & "</ruc>"
    Print #NumFile, cad
    cad = "<numAut>" & Trim$(lblAutoriza.Caption) & "</numAut>"
    Print #NumFile, cad

    cad = "<fecha>" & Format(dtpFecha.value, "dd/mm/yyyy") & "</fecha>"
    Print #NumFile, cad
    cad = "<detalles>"
    Print #NumFile, cad
    For i = 1 To grdTrans.Rows - 1
        If grdTrans.ValueMatrix(i, COL_CHEK) = "-1" Then
            cad = "  <detalle>"
            Print #NumFile, cad
            If Len(grdTrans.TextMatrix(i, 9)) <> 0 Then
                cad = "<codDoc>" & grdTrans.TextMatrix(i, COL_CODTIPOCOM) & "</codDoc>"
                Print #NumFile, cad
            Else
                MsgBox "Debe especificar el tipo de Documento en la Transaccion." & grdTrans.TextMatrix(i, 2), vbInformation
                cad = ""
            End If
            cad = "<estab>" & grdTrans.TextMatrix(i, COL_SERESTAB) & "</estab>"
            Print #NumFile, cad
            cad = "<ptoEmi>" & grdTrans.TextMatrix(i, COL_SERPUNTO) & "</ptoEmi>"
            Print #NumFile, cad
            Num = gobjMain.EmpresaActual.RecuperaNumeroTransaccionReporteRango(grdTrans.TextMatrix(i, COL_CODTRANS), lblAutoriza.Caption, True)
            cad = "<fin>" & Num & "</fin>"
            Print #NumFile, cad
            cad = "  </detalle>"
            Print #NumFile, cad
            'quita la configuracion de autoimpresor
'            Set gnt = gobjMain.EmpresaActual.RecuperaGNTrans(grdTrans.TextMatrix(i, COL_CODTRANS))
'            gnt.IVAutoImpresor = False
'            gnt.Grabar
            
            
            sql = " select opcion from Gntrans "
            sql = sql & " where codtrans='" & grdTrans.TextMatrix(i, COL_CODTRANS) & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            
            sql = " Update Gntrans "
            sql = sql & " set opcion ='" & Mid$(rs.Fields("opcion"), 1, 151) & "N" & Mid$(rs.Fields("opcion"), 153, 50) & "', "
            sql = sql & " BandValida= 0 "
            sql = sql & " where codtrans='" & grdTrans.TextMatrix(i, COL_CODTRANS) & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
            
            If VerificaAnulacionComprobantes Then
                sql = " Update Anexo_Comprobantes "
                sql = sql & " set BandAutorizaActiva =0, BandAutoImpresor=0"
                sql = sql & " where CodComprobante='" & grdTrans.TextMatrix(i, COL_CODTIPOCOM) & "'"
                Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
            End If
            
            
            
        End If
    Next i
    cad = "  </detalles>"
    Print #NumFile, cad
    cad = "</autorizacion>"
    Print #NumFile, cad
    Close NumFile
    
    
        CargarTrans
   
    
    'actualiza numro de puntos en gnsucursal
    If VerificaAnulacionComprobantesEstablecimiento Then
            For i = 1 To grdPunto.Rows - 1
                If grdPunto.Cols > 1 Then
                        sql = " Update GNSucursal "
                        sql = sql & " set NumPuntos =" & grdPunto.TextMatrix(i, 1)
                        sql = sql & " where Codsucursal='" & grdPunto.TextMatrix(i, 0) & "'"
                        Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
                Else
                        sql = " Update GNSucursal "
                        sql = sql & " set NumPuntos = 0"
                        sql = sql & " where Codsucursal='" & grdPunto.TextMatrix(i, 0) & "'"
                        Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
                
                End If
            
            Next i
    End If
    
'     CargarTrans
     
salida:
    GeneraArchivo_Eliminacion = False
    Exit Function
ErrTrap:
    DispErr
    GoTo salida
End Function

Private Function VerificaAnulacionComprobantes() As Boolean
Dim i As Integer, j As Integer
    VerificaAnulacionComprobantes = True
    For i = 1 To grdTrans.Rows - 1
        If grdTrans.ValueMatrix(i, COL_CHEK) = "-1" Then
            For j = i + 1 To grdTrans.Rows - 1
                If grdTrans.TextMatrix(i, COL_CODTIPOCOM) = grdTrans.TextMatrix(j, COL_CODTIPOCOM) Then
                    If grdTrans.ValueMatrix(i, COL_CHEK) <> grdTrans.ValueMatrix(j, COL_CHEK) Then
                        VerificaAnulacionComprobantes = False
                        Exit Function
                    End If
                End If
            Next j
            
            For j = grdTrans.Rows - 1 To 1 Step -1
                If grdTrans.TextMatrix(i, COL_CODTIPOCOM) = grdTrans.TextMatrix(j, COL_CODTIPOCOM) Then
                    If grdTrans.ValueMatrix(i, COL_CHEK) <> grdTrans.ValueMatrix(j, COL_CHEK) Then
                        VerificaAnulacionComprobantes = False
                        Exit Function
                    End If
                End If
            Next j
            
        End If
    Next i
End Function

Private Function VerificaAnulacionComprobantesEstablecimiento() As Boolean
Dim i As Integer, j As Integer
    VerificaAnulacionComprobantesEstablecimiento = True
    For i = 1 To grdTrans.Rows - 1
        If grdTrans.ValueMatrix(i, COL_CHEK) = "-1" Then
            For j = i + 1 To grdTrans.Rows - 1
                If grdTrans.TextMatrix(i, COL_SERESTAB) = grdTrans.TextMatrix(j, COL_SERESTAB) And grdTrans.TextMatrix(i, COL_SERPUNTO) = grdTrans.TextMatrix(j, COL_SERPUNTO) Then
                    If grdTrans.ValueMatrix(i, COL_CHEK) <> grdTrans.ValueMatrix(j, COL_CHEK) Then
                        VerificaAnulacionComprobantesEstablecimiento = False
                        Exit Function
                    End If
                End If
            Next j
            
            For j = grdTrans.Rows - 1 To 1 Step -1
                If grdTrans.TextMatrix(i, COL_SERESTAB) = grdTrans.TextMatrix(j, COL_SERESTAB) And grdTrans.TextMatrix(i, COL_SERPUNTO) = grdTrans.TextMatrix(j, COL_SERPUNTO) Then
                    If grdTrans.ValueMatrix(i, COL_CHEK) <> grdTrans.ValueMatrix(j, COL_CHEK) Then
                        VerificaAnulacionComprobantesEstablecimiento = False
                        Exit Function
                    End If
                End If
            Next j
            
        End If
    Next i
End Function


