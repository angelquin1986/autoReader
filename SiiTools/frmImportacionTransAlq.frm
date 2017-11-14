VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "flexcombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmImportacionTransAlq 
   Caption         =   "Importación de datos"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11115
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   11115
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pic2 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   11055
      TabIndex        =   19
      Top             =   5280
      Width           =   11115
      Begin VB.CommandButton cmdAbrir 
         Caption         =   "&Abrir"
         Height          =   375
         Left            =   9720
         TabIndex        =   22
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label lblArchivoLocutorio 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2280
         TabIndex        =   21
         Top             =   0
         Width           =   7335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta del Archivo de Datos: "
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   1965
      End
   End
   Begin VB.PictureBox picEncabezado 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   1536
      Left            =   0
      ScaleHeight     =   1500
      ScaleWidth      =   11085
      TabIndex        =   6
      Top             =   570
      Visible         =   0   'False
      Width           =   11115
      Begin VB.TextBox txtDescripcion 
         Height          =   336
         Left            =   3600
         MaxLength       =   120
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         ToolTipText     =   "Descripción de la transacción"
         Top             =   480
         Width           =   4740
      End
      Begin VB.TextBox txtCotizacion 
         Height          =   336
         Left            =   900
         TabIndex        =   7
         Top             =   1080
         Width           =   1452
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   336
         Left            =   900
         TabIndex        =   9
         ToolTipText     =   "Fecha de la transacción"
         Top             =   360
         Width           =   1452
         _ExtentX        =   2566
         _ExtentY        =   582
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
         Format          =   60489729
         CurrentDate     =   37078
         MaxDate         =   73415
         MinDate         =   29221
      End
      Begin FlexComboProy.FlexCombo fcbResp 
         Height          =   336
         Left            =   6888
         TabIndex        =   10
         ToolTipText     =   "Responsable de la transacción"
         Top             =   120
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
      Begin FlexComboProy.FlexCombo fcbTrans 
         Height          =   336
         Left            =   3600
         TabIndex        =   11
         ToolTipText     =   "Responsable de la transacción"
         Top             =   120
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
      Begin FlexComboProy.FlexCombo fcbMoneda 
         Height          =   336
         Left            =   900
         TabIndex        =   12
         ToolTipText     =   "Responsable de la transacción"
         Top             =   720
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
      Begin FlexComboProy.FlexCombo fcbBodOrigen 
         Height          =   330
         Left            =   9600
         TabIndex        =   23
         ToolTipText     =   "Bodega Origen"
         Top             =   120
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
      Begin FlexComboProy.FlexCombo fcbBodDestino 
         Height          =   330
         Left            =   9600
         TabIndex        =   25
         ToolTipText     =   "Bodega Destino"
         Top             =   600
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
      Begin FlexComboProy.FlexCombo fcbVendedor 
         Height          =   330
         Left            =   3600
         TabIndex        =   28
         ToolTipText     =   "Responsable de la transacción"
         Top             =   840
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "&Vendedor"
         Height          =   195
         Left            =   2760
         TabIndex        =   27
         Top             =   960
         Width           =   690
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Bod. Destino"
         Height          =   195
         Left            =   8640
         TabIndex        =   26
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Bod. Origen"
         Height          =   195
         Left            =   8640
         TabIndex        =   24
         Top             =   120
         Width           =   840
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "&Moneda  "
         Height          =   192
         Left            =   204
         TabIndex        =   18
         Top             =   720
         Width           =   672
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Fecha Transaccion  "
         Height          =   192
         Left            =   960
         TabIndex        =   17
         Top             =   120
         Width           =   1464
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "&Descripción  "
         Height          =   192
         Left            =   2604
         TabIndex        =   16
         Top             =   480
         Width           =   936
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "C&otización  "
         Height          =   192
         Left            =   60
         TabIndex        =   15
         Top             =   1104
         Width           =   816
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "&Responsable  "
         Height          =   192
         Left            =   5760
         TabIndex        =   14
         Top             =   120
         Width           =   1056
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cod.Trans  "
         Height          =   192
         Left            =   2712
         TabIndex        =   13
         Top             =   120
         Width           =   828
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   1815
      Left            =   0
      TabIndex        =   2
      Top             =   2100
      Width           =   5895
      _cx             =   10393
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
      FocusRect       =   3
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
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
      AllowUserFreezing=   2
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   7095
      Top             =   2385
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   612
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   11115
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5775
      Width           =   11115
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   5880
         TabIndex        =   4
         Top             =   120
         Width           =   1212
      End
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   180
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6480
      Top             =   2415
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacionTransAlq.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacionTransAlq.frx":0114
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacionTransAlq.frx":0568
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   1005
      ButtonWidth     =   1455
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Archivo..."
            Key             =   "Archivo"
            Object.ToolTipText     =   "Abrir archivo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Importar"
            Key             =   "Importar"
            Object.ToolTipText     =   "Importar"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cod.Trans"
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   804
   End
End
Attribute VB_Name = "frmImportacionTransAlq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbooCancelado As Boolean
Private mbooEjecutando As Boolean
Private mbooErrores As Boolean     '***Angel. 22/Abril/2004
Private conEncabezado As Boolean

Private Const MSG_OK As String = "OK"
Private Const MSG_ERR As String = "Error"
Private Const MSG_PROC As String = "Procesando..."

Private UltimoNumTransImportado As String


Public Sub Inicio(ByVal Tipo As String)
    Me.tag = Tipo                   'Tipo de importación
    Form_Resize
    Me.Show
    Me.ZOrder
    conEncabezado = False
    pic2.Visible = False
    Me.Caption = "Importacion de Saldos Iniciales de Inventario Alquilados"
    CargarEncabezado
    Form_Resize
    ConfigCols
End Sub

Private Sub ConfigCols()
    Dim s As String, i As Integer
    grd.Cols = 1
  
    
        s = "^#|<Código Cliente|<nombre|<Código Item|>Cantidad|>Costo Unitario|>Fecha Alquiler|>Fecha Devolucion"
        grd.FormatString = s & "|<Resultado"
    
    
    grd.ColSort(1) = flexSortGenericAscending
    grd.ColSort(2) = flexSortGenericAscending
    grd.ColSort(3) = flexSortGenericAscending
    
    
    
    
    AsignarTituloAColKey grd
    grd.SetFocus
End Sub

Private Sub CargarEncabezado()
    picEncabezado.Visible = True
    conEncabezado = True
    dtpFecha.value = Date
    fcbResp.SetData gobjMain.EmpresaActual.ListaGNResponsable(False)
    fcbTrans.SetData gobjMain.EmpresaActual.ListaGNTrans("", False, False)
    fcbMoneda.SetData gobjMain.EmpresaActual.ListaGNMoneda
    fcbBodOrigen.SetData gobjMain.EmpresaActual.ListaIVBodega(True, False)
    fcbBodDestino.SetData gobjMain.EmpresaActual.ListaIVBodega(True, False)
    fcbVendedor.SetData gobjMain.EmpresaActual.ListaFCVendedor(True, False)
    fcbMoneda.KeyText = "USD"
    txtCotizacion.Text = "1"
    txtDescripcion.Text = "Saldo Inicial Devolucion pendiente"
    
End Sub

Private Sub EnabledEncabezado(modo As Boolean)
    dtpFecha.Enabled = modo
    fcbResp.Enabled = modo
    fcbTrans.Enabled = modo
    fcbMoneda.Enabled = modo
    txtCotizacion.Enabled = modo
    
End Sub

Private Sub cmdAbrir_Click()
    LeerArchivo lblArchivoLocutorio.Caption
End Sub

Private Sub cmdCancelar_Click()
    If mbooEjecutando Then
        mbooCancelado = True
    Else
        Unload Me
    End If
End Sub
Private Sub ImportarGNComprobante()
    Dim i As Long, j As Long, gncomp As GNComprobante, NumeroComprobante As Integer
    Dim limite As Long
    Dim codcli As String
    On Error GoTo ErrTrap
    
    ' verificar si estan todos los datos
    If Len(fcbMoneda.Text) = 0 Then
        MsgBox "Debe selecciona una tipo de Modena", vbInformation
        fcbMoneda.SetFocus
        Exit Sub
    End If
    
    If Val(txtCotizacion.Text) = 0 Then
        MsgBox "Escriba una cotizacion valida", vbInformation
        txtCotizacion.SetFocus
        Exit Sub
    End If
    
    If Len(fcbTrans.Text) = 0 Then
        MsgBox "Seleccione un tipo de transaccion", vbInformation
        fcbTrans.SetFocus
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
    If Len(fcbBodOrigen.Text) = 0 Then
        MsgBox "Debe seleccionar una bodega origen", vbInformation
        fcbBodOrigen.SetFocus
        Exit Sub
    End If
    If Len(fcbBodDestino.Text) = 0 Then
        MsgBox "Debe seleccionar una bodega Destino", vbInformation
        fcbBodDestino.SetFocus
        Exit Sub
    End If
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna fila para importar.", vbInformation
        Exit Sub
    End If
    
    'Confirmación
    If MsgBox("Está seguro que desea comenzar el proceso de importación?", _
                vbYesNo + vbQuestion) <> vbYes Then Exit Sub
                
    mbooEjecutando = True
    MensajeStatus "Importando...", vbHourglass
    EnabledEncabezado False
    grd.Enabled = False
    With grd
        prg1.Min = .FixedRows - 1
        prg1.max = .Rows - 1
        prg1.value = prg1.Min
        Set gncomp = gobjMain.EmpresaActual.CreaGNComprobante(fcbTrans.KeyText)
        NumeroComprobante = gncomp.GNTrans.NumTransSiguiente
        

        
        For i = .FixedRows To .Rows - 1
            limite = buscarLimite(i, grd.TextMatrix(i, grd.ColIndex("Código Cliente")))
            If Not grd.IsSubtotal(i) Then
            
            prg1.value = i
            DoEvents                'Para dar control a Windows
            
            'Si usuario aplastó 'Cancelar', sale del ciclo
            If mbooCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                GoTo cancelado
            End If
            
            'Si aún no está importado bien, importa la fila
            If grd.TextMatrix(i, .Cols - 1) <> MSG_OK Then
                'Si ocurre un error y no quiere seguir el usuario, sale del ciclo
                    PonerDatosComprobante gncomp, NumeroComprobante, i
                    If Not ImportarFilaDetalleGNcomp(i, gncomp) Then GoTo cancelado
            End If
            End If
            If j = limite Then
                If GrabarGNComprobante(gncomp) Then
                    NumeroComprobante = NumeroComprobante + 1
                    Set gncomp = Nothing
                    Set gncomp = gobjMain.EmpresaActual.CreaGNComprobante(fcbTrans.KeyText)
                   ' PonerDatosComprobante gncomp, NumeroComprobante, i + 1
                End If
                'j = 0
            End If
            
            j = j + 1
        
        Next i
        ' graba el ultimo que no queda grabado
       ' GrabarGNComprobante gncomp
    End With
        
cancelado:
    Set gncomp = Nothing
    MensajeStatus
    mbooEjecutando = False
    prg1.value = prg1.Min
    grd.Enabled = True
    EnabledEncabezado True
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    mbooEjecutando = False
    prg1.value = prg1.Min
    grd.Enabled = True
    EnabledEncabezado True
    Exit Sub
End Sub
Private Function buscarLimite(ByVal filaDesde As Long, ByVal codCliente As String) As Long
Dim i As Long
    For i = filaDesde To grd.Rows - 1
    If grd.IsSubtotal(i) Then
        buscarLimite = i - 1
        Exit Function
    End If
    Next
End Function
Private Function GrabarGNComprobante(ByRef gncomp As GNComprobante) As Boolean
    Dim i As Long
    i = 0
            CreaCopia gncomp
            i = gncomp.CountIVKardex
    
    If i > 0 Then
        gncomp.Grabar False, False
        MsgBox "Se ha grabado un comprobante con numero = " & gncomp.CodTrans & gncomp.numtrans, vbInformation
        GrabarGNComprobante = True
    Else
        GrabarGNComprobante = False
    End If
End Function
Private Function ImportarFilaDetalleGNcomp(ByVal i As Long, ByRef gncomp As GNComprobante) As Boolean
    ImportarFilaDetalleGNcomp = True
    If Len(grd.TextMatrix(i, 1)) = 0 Then Exit Function
    If grd.IsSubtotal(i) Then Exit Function ' no importa la fila que tiene subtotal
        ImportarFilaDetalleGNcomp = ImportarFilaInventario(i, gncomp)
    
End Function

Sub PonerDatosComprobante(ByRef gncomp As GNComprobante, ByVal Num As Integer, ByVal fil As Long)
    gncomp.CodTrans = fcbTrans.KeyText
    gncomp.numtrans = Num
    gncomp.FechaTrans = dtpFecha.value
    gncomp.CodResponsable = fcbResp.KeyText
    gncomp.CodMoneda = fcbMoneda.KeyText
    gncomp.Cotizacion(fcbMoneda.KeyText) = Val(txtCotizacion.Text)
    gncomp.Descripcion = txtDescripcion.Text
    gncomp.CodClienteRef = grd.TextMatrix(fil, grd.ColIndex("Código Cliente"))
    gncomp.nombre = grd.TextMatrix(fil, grd.ColIndex("Nombre"))
    gncomp.CodVendedor = fcbVendedor.KeyText
End Sub

Private Function ImportarFilaInventario(ByVal i As Long, ByRef gncomp As GNComprobante) As Boolean
    Dim msg As String, IVinventario As IVKardex, ix As Long
    On Error GoTo ErrTrap
    ix = gncomp.AddIVKardex
    Set IVinventario = gncomp.IVKardex(ix)
    grd.TextMatrix(i, grd.Cols - 1) = MSG_PROC
    With IVinventario
        'AÑADIR DATOS A GRABAR EN ivinventario
        .CodInventario = grd.TextMatrix(i, grd.ColIndex("Código Item"))
        .FechaLleva = grd.TextMatrix(i, grd.ColIndex("Fecha Alquiler"))
        .FechaDevol = grd.TextMatrix(i, grd.ColIndex("Fecha Devolucion"))
        .NumDias = DateDiff("d", .FechaLleva, .FechaDevol)
        .cantidad = grd.ValueMatrix(i, grd.ColIndex("Cantidad")) * -1
        .CostoTotal = .cantidad * grd.ValueMatrix(i, grd.ColIndex("Costo Unitario"))
        .CostoRealTotal = .cantidad * grd.ValueMatrix(i, grd.ColIndex("Costo Unitario"))
        .CodBodega = fcbBodOrigen.KeyText
        .bandVer = 1
        .Orden = ix
    End With
    grd.TextMatrix(i, grd.Cols - 1) = MSG_OK
    ImportarFilaInventario = True
    Exit Function
ErrTrap:
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    
    msg = "Ha ocurrido un error al tratar de importar la fila #" & i & "." & vbCr & _
          "Código : " & grd.TextMatrix(i, 1) & vbCr & _
          "Error : " & Err.Description & vbCr & vbCr & _
          "Desea continuar el proceso desde la siguiente fila?"
    If MsgBox(msg, vbYesNo + vbExclamation) = vbYes Then
        ImportarFilaInventario = True
    Else
        ImportarFilaInventario = False
    End If
    gncomp.RemoveIVKardex ix, IVinventario
    Exit Function
End Function




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF9
'        cmdImprimir_Click
        KeyCode = 0
    Case Else
        MoverCampo Me, KeyCode, Shift, True
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ImpideSonidoEnter Me, KeyAscii
End Sub

Private Sub Form_Load()
    grd.Rows = grd.FixedRows    'Limpia la grilla
End Sub

Private Sub Form_Resize()
    Dim hei As Long
    On Error Resume Next
    hei = IIf(conEncabezado, picEncabezado.Height, 0)
    grd.Move 0, tlb1.Height + hei, Me.ScaleWidth, Me.ScaleHeight - tlb1.Height - pic1.Height - hei
    cmdCancelar.Move Me.ScaleWidth - cmdCancelar.Width - 160
    prg1.Width = Me.ScaleWidth - (prg1.Left * 2) - cmdCancelar.Width - 160
End Sub




Private Sub tlb1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo ErrTrap

    Select Case Button.Key
    Case "Archivo"
        AbrirArchivo
        grd.subtotal flexSTSum, grd.ColIndex("Código Cliente"), , , grd.GridColor, vbBlack, , "-", grd.ColIndex("Código Cliente"), False
        GNPoneNumFila grd, False
    Case "Importar"
        If conEncabezado Then   ' los documentos con encabezado tiene formato gncomprobante
            ImportarGNComprobante
        End If
    
    Case "Configurar"
        frmConfiguracion.Inicio
    End Select
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub AbrirArchivo()
    Dim i As Long
    On Error GoTo ErrTrap
    With dlg1
        .CancelError = True
'        .Filter = "Texto (Separado por coma)|*.txt|Excel 97(XLS)|*.xls"
        .Filter = "Texto (Separado por coma)|*.txt"
        .Flags = cdlOFNFileMustExist
        If Len(.FileName) = 0 Then          'Solo por primera vez, ubica a la carpeta de la aplicación
            .FileName = App.Path & "\*.txt"
        End If
        .ShowOpen
        LeerArchivo (dlg1.FileName)
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
            'InsertarColumnas
        Case ".XLS"
            VisualizarExcel archi
        Case Else
        End Select
End Sub

Private Sub ReformartearColumnas()
        ConfigCols
End Sub
Private Sub InsertarColumnaDesc_y_Cost()
    Dim pos   As Integer
    pos = 2
    grd.Cols = grd.Cols + 1
    grd.ColPosition(grd.Cols - 1) = pos
    grd.TextMatrix(0, pos) = "Descripción"
    grd.ColKey(pos) = grd.TextMatrix(0, pos)
    pos = 5
    grd.Cols = grd.Cols + 1
    grd.ColPosition(grd.Cols - 1) = pos
    grd.TextMatrix(0, pos) = "Costo Unitario"
    grd.ColKey(pos) = grd.TextMatrix(0, pos)
End Sub

Private Function ponerCostoUnitarioFila(ByVal i As Long) As Currency
    If Val(grd.ValueMatrix(i, grd.ColIndex("Cantidad"))) <> 0 Then
        ponerCostoUnitarioFila = Val(grd.ValueMatrix(i, grd.ColIndex("Costo Total"))) / Val(grd.ValueMatrix(i, grd.ColIndex("Cantidad")))
    End If
End Function
Private Function ponerCostoTotal(i As Long) As Currency
    ponerCostoTotal = grd.ValueMatrix(i, grd.ColIndex("Costo Unitario")) * grd.ValueMatrix(i, grd.ColIndex("Cantidad"))
End Function


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
        Case Else
            Separador = ","
    End Select
    
    'encontro = False
    
    Open archi For Input As #f
        Do Until EOF(f)
            Line Input #f, s
            s = vbTab & Replace(s, Separador, vbTab)      'Convierte ',' a TAB
            
                grd.AddItem s
            
        Loop
    Close #f
    RemueveSpace
    ' ordenar
    If grd.Rows > 1 Then
    
            grd.Select 1, 1, 1, 2
    
    


    
    End If
    grd.Sort = flexSortUseColSort

' poner numero
    GNPoneNumFila grd, False
    
    
    
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

Private Sub VisualizarExcel(ByVal archi As String)
    MsgBox "No se dispone de ésta función por el momento...", vbInformation
End Sub


Sub PonerMensaje(j As Long, i As Long, msg As String)
    Dim x As Integer
    For x = j To i
        grd.TextMatrix(x, grd.Cols - 1) = msg
    Next x
End Sub


Private Sub CreaCopia(ByVal gc As GNComprobante)
    Dim i As Long, ivk As IVKardex, ivk2 As IVKardex, j As Long
    
    'Asegura que todo sea de la bodega de origen
    For i = 1 To gc.CountIVKardex
        Set ivk = gc.IVKardex(i)
        ivk.CodBodega = fcbBodOrigen.KeyText
        ivk.cantidad = Abs(ivk.cantidad) * -1   'Origen es siempre negativa
    Next i
    Set ivk = Nothing
    
    'Duplica IVKardex para la bodega de destino
    ' multiplicando -1 a la cantidad
    For i = 1 To gc.CountIVKardex
        Set ivk = gc.IVKardex(i)
        j = gc.AddIVKardex
        Set ivk2 = gc.IVKardex(j)
        With ivk2
            .CodBodega = fcbBodDestino.KeyText
            .CodInventario = ivk.CodInventario
            .cantidad = ivk.cantidad * -1       'Cambia de signo
            .CostoTotal = ivk.CostoTotal * -1
            .CostoRealTotal = ivk.CostoRealTotal * -1
            .PrecioTotal = ivk.PrecioTotal * -1
            .PrecioRealTotal = ivk.PrecioRealTotal * -1
            .Descuento = ivk.Descuento
            .IVA = ivk.IVA
            .Nota = ivk.Nota
            .IdPadre = ivk.IdPadre
            .bandImprimir = ivk.bandImprimir
            .bandVer = ivk.bandVer
            .NumDias = ivk.NumDias
            .FechaDevol = ivk.FechaDevol
            .Orden = j
        End With
    Next i
    Set ivk = Nothing
    Set ivk2 = Nothing
End Sub

