VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{50067EB3-D6AF-11D3-8297-000021C5085D}#1.0#0"; "NTextBox.ocx"
Object = "{ED5A9B02-5BDB-48C7-BAB1-642DCC8C9E4D}#2.0#0"; "SelFold.ocx"
Begin VB.Form frmDinardap 
   Caption         =   "dlg1"
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
   Begin VB.Frame frmfecha 
      Height          =   1815
      Left            =   60
      TabIndex        =   7
      Top             =   0
      Width           =   8895
      Begin VB.CommandButton cmdPasos 
         Caption         =   "Exportar a Excel"
         Height          =   330
         Index           =   6
         Left            =   5580
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1020
         Width           =   1455
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "Sumar Dias de Gracia"
         Height          =   330
         Index           =   5
         Left            =   6120
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   540
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.CheckBox chkSoloFacturas 
         Alignment       =   1  'Right Justify
         Caption         =   "Solo Transacciones Venta"
         Height          =   255
         Left            =   6120
         TabIndex        =   18
         Top             =   1500
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.CheckBox chksolocartera 
         Alignment       =   1  'Right Justify
         Caption         =   "Solo Transacciones con Cartera"
         Height          =   255
         Left            =   6120
         TabIndex        =   17
         Top             =   1500
         Visible         =   0   'False
         Width           =   2715
      End
      Begin NTextBoxProy.NTextBox ntxDiasGracia 
         Height          =   315
         Left            =   7500
         TabIndex        =   16
         Top             =   180
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
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
      Begin VB.CommandButton cmdPasos 
         Caption         =   "Abrir Archivo Pipe"
         Height          =   330
         Index           =   4
         Left            =   4140
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1020
         Width           =   1455
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "Abrir Archivo Excel"
         Height          =   330
         Index           =   3
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1020
         Width           =   1455
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "Buscar"
         Height          =   330
         Index           =   2
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1020
         Width           =   1455
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "Generar Archivo"
         Height          =   330
         Index           =   10
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1380
         Width           =   1455
      End
      Begin VB.TextBox txtCarpeta 
         Height          =   320
         Left            =   1080
         TabIndex        =   2
         Text            =   "c:\"
         Top             =   660
         Width           =   4170
      End
      Begin VB.CommandButton cmdExaminarCarpeta 
         Caption         =   "..."
         Height          =   320
         Index           =   0
         Left            =   5280
         TabIndex        =   3
         Top             =   660
         Width           =   372
      End
      Begin SelFold.SelFolder slf 
         Left            =   4860
         Top             =   480
         _ExtentX        =   1349
         _ExtentY        =   265
         Title           =   "Seleccione una carpeta"
         Caption         =   "Selección de carpeta"
         RootFolder      =   "\"
         Path            =   "C:\VBPROG_ESP\SII\SELFOLD"
      End
      Begin MSComCtl2.DTPicker dtpPeriodo 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   300
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
         Format          =   104988675
         CurrentDate     =   37356
      End
      Begin MSComDlg.CommonDialog dlg1 
         Left            =   3420
         Top             =   240
         _ExtentX        =   688
         _ExtentY        =   688
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Label lblDiasGracia 
         Caption         =   "Dias de Gracia"
         Height          =   195
         Left            =   6120
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblResp 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   5
         Left            =   2640
         TabIndex        =   12
         Top             =   1380
         Width           =   4365
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Corte:"
         Height          =   255
         Left            =   60
         TabIndex        =   9
         Top             =   360
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Ubicacion:"
         Height          =   255
         Left            =   60
         TabIndex        =   8
         Top             =   720
         Width           =   870
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
         TabIndex        =   5
         Top             =   60
         Width           =   1212
      End
      Begin MSComctlLib.ProgressBar prg 
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   120
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   3870
      Left            =   60
      TabIndex        =   4
      Top             =   1860
      Width           =   15015
      _cx             =   26485
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
      FormatString    =   $"frmDinardap.frx":0000
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
End
Attribute VB_Name = "frmDinardap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ex As Excel.Application, ws As Worksheet, wkb As Workbook

Private mbooProcesando As Boolean
Private mbooCancelado As Boolean
Private mEmpOrigen As Empresa
Private Const MSG_OK As String = "OK"
Private mObjCond As RepCondicion
Private mobjBusq As Busqueda

Private WithEvents mGrupo As grupo
Attribute mGrupo.VB_VarHelpID = -1
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
Const COL_V_RESP = 12

Const COL_VE_SUC = 1
Const COL_VE_TIPOCOMP = 2
Const COL_VE_CANTRANS = 3
Const COL_VE_BASE0 = 4
Const COL_VE_BASEIVA = 5
Const COL_VE_BASENOIVA = 6
Const COL_VE_TOTAL = 7
Const COL_VE_RESP = 8

    Const COL_D_CODENT = 1
    Const COL_D_FECHADATOS = 2
    Const COL_D_TIPOIDENT = 3
    Const COL_D_RUC = 4
    Const COL_D_NOMBRE = 5
    Const COL_D_TIPOSUJETO = 6
    Const COL_D_PROVINCIA = 7
    Const COL_D_CANTON = 8
    Const COL_D_PARROQUIA = 9
    Const COL_D_SEXO = 10
    Const COL_D_ESTADOCIVIL = 11
    Const COL_D_ORIGENINGRESO = 12
    Const COL_D_TRANS = 13
    Const COL_D_VALOR = 14
    Const COL_D_SALDO = 15
    Const COL_D_FECHACONCESION = 16
    Const COL_D_FECHAVENCI = 17
    Const COL_D_FECHAEXIGIBLE = 18
  Const COL_D_DIASCREDITO = 19
  Const COL_D_PERIODICIDADPAGO = 20
  Const COL_D_DIASMORA = 21
  Const COL_D_MONTOMOROSIDAD = 22
  Const COL_D_INTERESMORA = 23
  Const COL_D_SALDOXV1_30 = 24
  Const COL_D_SALDOXV31_90 = 25
  Const COL_D_SALDOXV91_180 = 26
  Const COL_D_SALDOXV181_360 = 27
  Const COL_D_SALDOXVMAS_360 = 28
  Const COL_D_SALDOVE1_30 = 29
  Const COL_D_SALDOVE31_90 = 30
  Const COL_D_SALDOVE91_180 = 31
  Const COL_D_SALDOVE181_360 = 32
  Const COL_D_SALDOVEMAS_360 = 33
  Const COL_D_VALORDEMANDA = 34
  Const COL_D_CARTERACASTIGADA = 35
  Const COL_D_CUOTACREDITO = 36
  Const COL_D_FECHAPAGO = 37
  Const COL_D_FORMACANCELACION = 38
  Const COL_D_CODTRANS = 39
  Const COL_D_NUMTRANS = 40
  Const COL_D_TRANSID = 41
  Const COL_D_TIPOTRANS = 42
  Const COL_D_RESP = 43



Private Cadena As String
Private cadenaDD  As String

Private NumFile As Integer
Private NumProc As Integer
Private TotalVentas As Currency
Private mbooEjecutando As Boolean
Private BandError As Boolean
    
Public Sub Inicio(ByVal tag As String)
    On Error GoTo ErrTrap
    Set mObjCond = New RepCondicion
'    Select Case tag
 '       Case "FAT"
    Me.Caption = "Reporte DINARDAP"
  '  End Select
    TotalVentas = 0
    dtpPeriodo.value = CDate("01/" & IIf(Month(Date) - 1 <> 0, Month(Date) - 1, 12) & "/" & Year(Date))
    mObjCond.fecha1 = dtpPeriodo.value
    
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("DiasGraciaDINARDAP")) > 0 Then
        ntxDiasGracia.value = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("DiasGraciaDINARDAP")
        mObjCond.Num1 = ntxDiasGracia.value
    End If

    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Ruta-DINARDAP")) > 0 Then
        txtCarpeta.Text = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Ruta-DINARDAP")
    End If
    
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CambiaFechaVenci-DINARDAP")) > 0 Then
        If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CambiaFechaVenci-DINARDAP") = "1" Then
            lblDiasGracia.Visible = True
            ntxDiasGracia.Visible = True
            cmdPasos(5).Visible = True
            If ntxDiasGracia.value = 0 Then
                ntxDiasGracia.Enabled = True
                cmdPasos(5).Enabled = True
            Else
                ntxDiasGracia.Enabled = False
                cmdPasos(5).Enabled = False
            End If
            
        End If
    End If
    Me.tag = tag
    Me.Show
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub





Private Sub cmdCancelar_Click()
    mbooCancelado = True
End Sub


Private Sub cmdPasos_Click(Index As Integer)
    Dim r As Boolean, cad As String, nombre As String, file As String, fecha As Date, FECHANOMBRE As String
    NumProc = Index + 1
    
    
        cmdPasos(2).BackColor = vbButtonFace
        cmdPasos(3).BackColor = vbButtonFace
        cmdPasos(4).BackColor = vbButtonFace
        cmdPasos(10).BackColor = vbButtonFace

    
    Select Case Index + 1
    Case 3      '2. Busca Ventas
            
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("DINARDAP2015")) > 0 Then
        If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("DINARDAP2015") = "1" Then
            BuscarVentasDinardap
        Else
            BuscarVentasATS
        End If
    Else
        BuscarVentasATS
    End If
            BandError = False
            cmdPasos(2).BackColor = &HFFFF00
    Case 4
            AbrirArchivoExcel
            cmdPasos(3).BackColor = &HFFFF00
    Case 5
            AbrirArchivoPipe
            cmdPasos(4).BackColor = &HFFFF00
    
    Case 6
            
            'SumarDiasFechaVencimiento
            CambiaDiasPlazo
            cmdPasos(5).Visible = False
            lblDiasGracia.Visible = False
            ntxDiasGracia.Visible = False
            
    Case 7
            
            ExportaExcel "Reporte para la Dinardap"
    
    
    Case 11      '8. Generar Archivo
    
            fecha = DateAdd("d", -1, DateAdd("m", 1, "01/" & DatePart("m", dtpPeriodo.value) & "/" & DatePart("yyyy", dtpPeriodo.value)))
            FECHANOMBRE = Mid$(DateAdd("d", -1, DateAdd("m", 1, "01/" & DatePart("m", dtpPeriodo.value) & "/" & DatePart("yyyy", dtpPeriodo.value))), 1, 2)
            FECHANOMBRE = FECHANOMBRE & Mid$(DateAdd("d", -1, DateAdd("m", 1, "01/" & DatePart("m", dtpPeriodo.value) & "/" & DatePart("yyyy", dtpPeriodo.value))), 4, 2)
            FECHANOMBRE = FECHANOMBRE & Mid$(DateAdd("d", -1, DateAdd("m", 1, "01/" & DatePart("m", dtpPeriodo.value) & "/" & DatePart("yyyy", dtpPeriodo.value))), 7, 4)
            
            nombre = Format(gobjMain.EmpresaActual.GNOpcion.ruc, "0000000000000") & FECHANOMBRE & ".txt"
            file = txtCarpeta.Text & nombre
            If ExisteArchivo(file) Then
                If MsgBox("El nombre del archivo " & nombre & " ya existe desea sobreescribirlo?", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
            NumFile = FreeFile
            Open file For Output Access Write As #NumFile
            'ExportaTxtPipe grd, nombre
            
            r = GeneraArchivoDinardap(cadenaDD)
            Cadena = cadenaDD
'            Print #numfile, Cadena
            Close NumFile
            
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "Ruta-DINARDAP", txtCarpeta.Text
            gobjMain.EmpresaActual.GNOpcion.Grabar
            

            
        If Not BandError Then
            r = False
            lblResp(5).Caption = "OK."
        Else
            r = True
            lblResp(5).Caption = "ERROR"
        End If
    
    End Select
    
    
    
        If r Then

                lblResp(5).BackColor = vbBlue
                lblResp(5).ForeColor = vbYellow

        Else
                lblResp(5).BackColor = vbBlue
                lblResp(5).ForeColor = vbYellow
        End If
End Sub

Private Sub dtpPeriodo_Change()
 Dim i As Integer

        cmdPasos(2).Enabled = True
    
        lblResp(5).BackColor = &HC0FFFF
        lblResp(5).Caption = ""

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
    grd.Move 0, frmfecha.Height + 100, Me.ScaleWidth, (Me.ScaleHeight - (frmfecha.Height + picBoton.Height) - 105)
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


Private Sub ConfigCols(cad As String)
    Dim s As String, i As Integer
    Select Case cad

    Case "IMPFC"

        s = "^#|<Cod Entidad|>Fecha Datos|^Tipo Ident|<RUC|<Nombre|^Tipo Sujeto|^Provincia|^canton|^parroquia|^Sexo|^Estado Civil|^Origen Ingreso|<Trans|>Valor|>Saldo|>Fecha Concesion|>Fecha Venci|>Fecha Exigible|>Dias Credito|>Periodicidad Pago|>Dias Mora|>Monto Morosidad|>Interes Mora|>SaldoxV1_30|>SaldoxV31_90|>SaldoxV91_180|>SaldoxV181_360|>SaldoxVmas_360|>SaldoVe1_30|>SaldoVe31_90|>SaldoVe91_180|>SaldoVe181_360|>SaldoVemas_360|>Valor Demanda|>Cartera Castigada|>Cuota Credito|>Fecha de Pago|^Forma Cancelacion|<Trans |>Num Trans|>Transid|<Tipo Trans"
        grd.FormatString = s & "|<         Resultado           "
        AsignarTituloAColKey grd
        
        grd.ColHidden(COL_D_TRANSID) = True
        
        grd.ColFormat(COL_D_VALOR) = "#,#0.00"
        grd.ColFormat(COL_D_SALDO) = "#,#0.00"
        grd.ColFormat(COL_D_MONTOMOROSIDAD) = "#,#0.00"
        grd.ColFormat(COL_D_SALDOXV1_30) = "#,#0.00"
        grd.ColFormat(COL_D_SALDOXV31_90) = "#,#0.00"
        grd.ColFormat(COL_D_SALDOXV91_180) = "#,#0.00"
        grd.ColFormat(COL_D_SALDOXV181_360) = "#,#0.00"
        grd.ColFormat(COL_D_SALDOXVMAS_360) = "#,#0.00"
        grd.ColFormat(COL_D_SALDOVE1_30) = "#,#0.00"
        grd.ColFormat(COL_D_SALDOVE31_90) = "#,#0.00"
        grd.ColFormat(COL_D_SALDOVE91_180) = "#,#0.00"
        grd.ColFormat(COL_D_SALDOVE181_360) = "#,#0.00"
        grd.ColFormat(COL_D_SALDOVEMAS_360) = "#,#0.00"
        grd.ColFormat(COL_D_VALORDEMANDA) = "#,#0.00"
        grd.ColFormat(COL_D_CARTERACASTIGADA) = "#,#0.00"
        grd.ColFormat(COL_D_CUOTACREDITO) = "#,#0.00"

            grd.ColData(COL_D_VALOR) = "SubTotal"
            grd.ColData(COL_D_SALDO) = "SubTotal"
            grd.ColData(COL_D_MONTOMOROSIDAD) = "SubTotal"
            grd.ColData(COL_D_SALDOXV1_30) = "SubTotal"
            grd.ColData(COL_D_SALDOXV31_90) = "SubTotal"
            grd.ColData(COL_D_SALDOXV91_180) = "SubTotal"
            grd.ColData(COL_D_SALDOXV181_360) = "SubTotal"
            grd.ColData(COL_D_SALDOXVMAS_360) = "SubTotal"
            grd.ColData(COL_D_SALDOVE1_30) = "SubTotal"
            grd.ColData(COL_D_SALDOVE31_90) = "SubTotal"
            grd.ColData(COL_D_SALDOVE91_180) = "SubTotal"
            grd.ColData(COL_D_SALDOVE181_360) = "SubTotal"
            grd.ColData(COL_D_SALDOVEMAS_360) = "SubTotal"
            grd.ColData(COL_D_VALORDEMANDA) = "SubTotal"
            grd.ColData(COL_D_CARTERACASTIGADA) = "SubTotal"
            grd.ColData(COL_D_CUOTACREDITO) = "SubTotal"

    
    End Select
   If grd.Rows > 1 Then
        grd.Cell(flexcpBackColor, 1, COL_D_CODENT, grd.Rows - 1, COL_D_RESP) = &H80000018
        grd.Cell(flexcpBackColor, 1, COL_D_NOMBRE, grd.Rows - 1, COL_D_NOMBRE) = vbWhite
        grd.Cell(flexcpBackColor, 1, COL_D_TRANS, grd.Rows - 1, COL_D_TRANS) = vbWhite
        grd.Cell(flexcpBackColor, 1, COL_D_FECHAVENCI, grd.Rows - 1, COL_D_FECHAVENCI) = vbWhite
        grd.Cell(flexcpBackColor, 1, COL_D_FECHAPAGO, grd.Rows - 1, COL_D_FECHAPAGO) = vbWhite
        grd.Cell(flexcpBackColor, 1, COL_D_FECHAPAGO + 1, grd.Rows - 1, COL_D_FECHAPAGO + 1) = vbWhite
        
'        SubTotalizar (COL_D_VALOR)
        Totalizar
    End If
    

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
    cad = cad & "<!--  Dir: Av. Espana  y Elia Liut Aeropuerto Mariscal Lamar Segundo Piso -->"
    cad = cad & "<!--  Telf: 098499003, 072870346      -->"
    cad = cad & "<!--  email: ishidacue@hotmail.com    -->"
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


Private Sub grd_DblClick()
    Dim gnc As GNComprobante, cad As String
    Dim pc As PCProvCli, Prov As PCProvincia, CANTO As PCCanton, fecha As Date, FechaOrig As Date, diasmora As Integer, diascredito As Integer, valor As Currency
    Dim forma As String, nombreant As String, i As Integer, trans As String
    Dim sql As String, rs As Recordset, v As Variant
            fecha = grd.TextMatrix(grd.Row, COL_D_FECHAVENCI)
            FechaOrig = grd.TextMatrix(grd.Row, COL_D_FECHAVENCI)
            valor = grd.ValueMatrix(grd.Row, COL_D_SALDO)
            
            Select Case grd.col
            Case COL_D_TRANS
                trans = grd.TextMatrix(grd.Row, COL_D_TRANS)
                cad = frmDatosPC.InicioDINARDAPNumTrans(trans)
                v = Split(trans, ";")
                grd.TextMatrix(grd.Row, COL_D_TRANS) = v(0)
                grd.Cell(flexcpBackColor, grd.Row, COL_D_TRANS, grd.Row, COL_D_TRANS) = vbRed
                If UBound(v) > 0 Then
                    If v(1) = 1 Then
                        sql = "select  id from gnoferta where transid=" & grd.ValueMatrix(grd.Row, COL_D_TRANSID)
                        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                        If rs.RecordCount > 0 Then
                            sql = "update gnoferta set BandOmitir=1 where transid=" & grd.ValueMatrix(grd.Row, COL_D_TRANSID)
                        Else
                            sql = "insert gnoferta (transid,  BandOmitir, Atencion, FormaPago) values (" & grd.ValueMatrix(grd.Row, COL_D_TRANSID) & ",1,'x Omitir DINARDAP','')"
                        End If
                        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                    End If
                End If
                
            Case COL_D_FECHAPAGO
                cad = frmDatosPC.InicioDINARDAPFechaForma(fecha, forma)
                If fecha < CDate(grd.TextMatrix(grd.Row, COL_D_FECHACONCESION)) Then
                    MsgBox "La fecha de Pago no puede ser menor a la de emision)"
                Else
                    grd.TextMatrix(grd.Row, COL_D_FECHAPAGO) = CDate(fecha)
                    grd.TextMatrix(grd.Row, COL_D_FORMACANCELACION) = forma
                    
                    grd.Cell(flexcpBackColor, grd.Row, COL_D_FECHAPAGO, grd.Row, COL_D_FORMACANCELACION) = vbRed
                End If
            
            Case COL_D_FECHAVENCI
                cad = frmDatosPC.InicioDINARDAPFecha(fecha)
                If fecha < CDate(grd.TextMatrix(grd.Row, COL_D_FECHACONCESION)) Then
                    MsgBox "La fecha de Vencimiento no puede ser menor a la de emision)"
                Else
                    If fecha <> FechaOrig Then
                        grd.Cell(flexcpBackColor, grd.Row, COL_D_FECHAVENCI, grd.Row, COL_D_FECHAVENCI) = vbRed
                        grd.TextMatrix(grd.Row, COL_D_SALDOXVMAS_360) = "0.00"
                        grd.TextMatrix(grd.Row, COL_D_SALDOXV181_360) = "0.00"
                        grd.TextMatrix(grd.Row, COL_D_SALDOXV91_180) = "0.00"
                        grd.TextMatrix(grd.Row, COL_D_SALDOXV31_90) = "0.00"
                        grd.TextMatrix(grd.Row, COL_D_SALDOXV1_30) = "0.00"
                        grd.TextMatrix(grd.Row, COL_D_SALDOVEMAS_360) = "0.00"
                        grd.TextMatrix(grd.Row, COL_D_SALDOVE181_360) = "0.00"
                        grd.TextMatrix(grd.Row, COL_D_SALDOVE91_180) = "0.00"
                        grd.TextMatrix(grd.Row, COL_D_SALDOVE31_90) = "0.00"
                        grd.TextMatrix(grd.Row, COL_D_SALDOVE1_30) = "0.00"
                        
                        
                        grd.TextMatrix(grd.Row, COL_D_FECHAVENCI) = CDate(fecha)
                        grd.TextMatrix(grd.Row, COL_D_FECHAEXIGIBLE) = CDate(fecha)
                        diascredito = DateDiff("d", CDate(grd.TextMatrix(grd.Row, COL_D_FECHACONCESION)), CDate(fecha))
                        grd.TextMatrix(grd.Row, COL_D_DIASCREDITO) = diascredito
                        diasmora = DateDiff("d", CDate(grd.TextMatrix(grd.Row, COL_D_FECHAVENCI)), grd.TextMatrix(grd.Row, COL_D_FECHADATOS))
                        
                        If diasmora > 0 Then
                            grd.TextMatrix(grd.Row, COL_D_DIASMORA) = diasmora
                            If diasmora > 360 Then
                                grd.TextMatrix(grd.Row, COL_D_SALDOVEMAS_360) = valor
                            ElseIf diasmora > 180 And diascredito <= 360 Then
                                grd.TextMatrix(grd.Row, COL_D_SALDOVE181_360) = valor
                            ElseIf diasmora > 90 And diascredito <= 180 Then
                                grd.TextMatrix(grd.Row, COL_D_SALDOVE91_180) = valor
                            ElseIf diasmora > 30 And diascredito <= 90 Then
                                grd.TextMatrix(grd.Row, COL_D_SALDOVE31_90) = valor
                            Else
                                grd.TextMatrix(grd.Row, COL_D_SALDOVE1_30) = valor
                            End If
                            
                        Else
                            grd.TextMatrix(grd.Row, COL_D_DIASMORA) = "0"
                            grd.TextMatrix(grd.Row, COL_D_MONTOMOROSIDAD) = "0.00"
                            If diascredito > 360 Then
                                grd.TextMatrix(grd.Row, COL_D_SALDOXVMAS_360) = valor
                            ElseIf diascredito > 180 And diascredito <= 360 Then
                                grd.TextMatrix(grd.Row, COL_D_SALDOXV181_360) = valor
                            ElseIf diascredito > 90 And diascredito <= 180 Then
                                grd.TextMatrix(grd.Row, COL_D_SALDOXV91_180) = valor
                            ElseIf diascredito > 30 And diascredito <= 90 Then
                                grd.TextMatrix(grd.Row, COL_D_SALDOXV31_90) = valor
                            Else
                                grd.TextMatrix(grd.Row, COL_D_SALDOXV1_30) = valor
                            End If
                            
                        End If
                        
                        
                        If grd.ValueMatrix(grd.Row, COL_D_DIASCREDITO) > 0 Then
                        End If
                    End If
                End If

            Case COL_D_NOMBRE
                    nombreant = grd.TextMatrix(grd.Row, COL_D_NOMBRE)
                    Set pc = gobjMain.EmpresaActual.RecuperaPCProvClixRUC(grd.TextMatrix(grd.Row, COL_D_RUC), True, False, False)
                    If Not pc Is Nothing Then

                        cad = frmDatosPC.InicioDINARDAP(pc)
                        If cad = "O.K." Then
                            grd.TextMatrix(grd.Row, COL_D_NOMBRE) = pc.nombre
                            If nombreant <> pc.nombre Then
                                If MsgBox("El Nombre del cliente esta cambiado al del catalogo, desea grabar el cambio de nombre", vbYesNo) = vbNo Then
                                    pc.nombre = nombreant
                                End If
                            End If
                            pc.Grabar
                            grd.TextMatrix(grd.Row, COL_D_TIPOIDENT) = pc.codtipoDocumento
                            grd.TextMatrix(grd.Row, COL_D_TIPOSUJETO) = pc.Tiposujeto
                            Set Prov = gobjMain.EmpresaActual.RecuperaPCProvincia(pc.codProvincia)
                            grd.TextMatrix(grd.Row, COL_D_PROVINCIA) = Prov.CodProvinciaSC
                            Set Prov = Nothing
                            Set CANTO = gobjMain.EmpresaActual.RecuperaPCCanton(pc.codCanton)
                            grd.TextMatrix(grd.Row, COL_D_CANTON) = CANTO.CodCantonSC
                            Set CANTO = Nothing
                            grd.TextMatrix(grd.Row, COL_D_PARROQUIA) = pc.codParroquia
                            If pc.Tiposujeto <> "J" Then
                                grd.TextMatrix(grd.Row, COL_D_SEXO) = pc.sexo
                                grd.TextMatrix(grd.Row, COL_D_ESTADOCIVIL) = pc.EstadoCivil
                                grd.TextMatrix(grd.Row, COL_D_ORIGENINGRESO) = pc.OrigenIngresos
                            End If
                            For i = grd.Row + 1 To grd.Rows - 1
                                If grd.TextMatrix(grd.Row, COL_D_RUC) = grd.TextMatrix(i, COL_D_RUC) Then
                                    grd.TextMatrix(i, COL_D_NOMBRE) = grd.TextMatrix(grd.Row, COL_D_NOMBRE)
                                    grd.TextMatrix(i, COL_D_TIPOIDENT) = grd.TextMatrix(grd.Row, COL_D_TIPOIDENT)
                                    grd.TextMatrix(i, COL_D_TIPOSUJETO) = grd.TextMatrix(grd.Row, COL_D_TIPOSUJETO)
                                    grd.TextMatrix(i, COL_D_PROVINCIA) = grd.TextMatrix(grd.Row, COL_D_PROVINCIA)
                                    grd.TextMatrix(i, COL_D_CANTON) = grd.TextMatrix(grd.Row, COL_D_CANTON)
                                    grd.TextMatrix(i, COL_D_PARROQUIA) = grd.TextMatrix(grd.Row, COL_D_PARROQUIA)
                                    grd.TextMatrix(i, COL_D_SEXO) = grd.TextMatrix(grd.Row, COL_D_SEXO)
                                    grd.TextMatrix(i, COL_D_ESTADOCIVIL) = grd.TextMatrix(grd.Row, COL_D_ESTADOCIVIL)
                                    grd.TextMatrix(i, COL_D_ORIGENINGRESO) = grd.TextMatrix(grd.Row, COL_D_ORIGENINGRESO)
                                End If
                            Next i
                            
                        End If
                    End If
                End Select
            
''            End Select
''            Set gnc = Nothing
''        End If
''    Case 3, 4
''        Set pc = gobjMain.EmpresaActual.RecuperaPCProvCli(grd.TextMatrix(grd.Row, COL_V_RUC))
''        Select Case grd.col
''        Case COL_V_RUC, COL_V_RUC, COL_V_CLIENTE
''            cad = frmDatosPC.InicioDINARDAP(pc)
''                    If cad = "O.K." Then
''                        pc.Grabar
''                    End If
'''                    grd.TextMatrix(grd.Row, COL_V_TIPODOC) = pc.CodTipoDocumento
''        End Select
''    End Select
    Set pc = Nothing
End Sub

Private Sub txtCarpeta_LostFocus()
    If Right$(txtCarpeta.Text, 1) <> "\" Then
        txtCarpeta.Text = txtCarpeta.Text & "\"
    End If
    'Luego a actualiza linea de comando
End Sub

Private Function BuscarVentasATS()
    Dim fecha1 As Date
    Dim fecha2 As Date

    On Error GoTo ErrTrap
        With grd
        .Redraw = False
        .Rows = .FixedRows
        If Not frmB_Trans.Inicio(gobjMain, "IMPDD", dtpPeriodo.value) Then
            grd.SetFocus
        End If
        
        If DatePart("m", fecha) = 12 Then
            fecha1 = "01/" & DatePart("m", dtpPeriodo.value) & "/" & DatePart("yyyy", dtpPeriodo.value)
            fecha2 = DateAdd("yyyy", 1, DateAdd("d", -1, ("01/" & DatePart("m", DateAdd("m", 1, dtpPeriodo.value)) & "/" & DatePart("yyyy", dtpPeriodo.value))))
        Else
            fecha1 = "01/" & DatePart("m", dtpPeriodo.value) & "/" & DatePart("yyyy", dtpPeriodo.value)
            fecha2 = DateAdd("d", -1, ("01/" & DatePart("m", DateAdd("m", 1, dtpPeriodo.value)) & "/" & DatePart("yyyy", dtpPeriodo.value)))
        End If
'        With objSiiMain.objCondicion
'        .fecha1 = fecha1
'        .fecha2 = fecha2

        
        gobjMain.objCondicion.fecha1 = DateAdd("d", -1, DateAdd("m", 1, fecha1))
'        gobjMain.objCondicion.fecha2 = fecha1
        MiGetRowsRep gobjMain.EmpresaActual.ConsVentasDINARDAP(), grd


        'GeneraArchivo

        ConfigCols "IMPFC"
        AjustarAutoSize grd, -1, -1
        AjustarAutoSize grd, -1, -1
        grd.ColWidth(0) = "500"


        GNPoneNumFila grd, False


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
    Dim i As Long, j As Long
    Dim vIR As Variant, cadenaFCIR As String
    Dim FilasIR As Long, ColumnasIR As Long, iIR As Long, jIR As Long
    Dim rsRet As Recordset, cadenaFCIVA30 As String
    Dim cadenaFCIVA70 As String, cadenaFCIVA100 As String
    Dim rsNC As Recordset, cadenaNC As String
    Dim msg As String, pc As PCProvCli, bandCF As Boolean, filaCF As Integer
    Dim cadenaF As String, k As Integer
    
    On Error GoTo ErrTrap
    GeneraArchivoATSVentasXML = True
    bandCF = False
    filaCF = 1
    
    
        grd.Refresh
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
'                i = 2802
                grd.ShowCell i, 1
                prg.value = i
                DoEvents
                cadenaFC = ""
'                chkConsFinal.value = vbChecked

                

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
                            GoTo SiguienteFila

                    
                    Case Else
                            
                            'cadenaFC = Mid$(cadenaFC, 1, Len(cadenaFC) - Len("<detalleVentas>") + 1)
                            msg = " El Cliente " & grd.TextMatrix(i, COL_V_CLIENTE) & " No tiene seleccionado el tipo de Documento"
                            grd.TextMatrix(i, grd.ColIndex("Resultado")) = " Error " & msg
                            grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbRed
                            grd.ShowCell i, grd.ColIndex("Resultado")
                            GeneraArchivoATSVentasXML = True
                            lblResp(1).Caption = "Error"
                            GoTo SiguienteFila
                        
                End Select
                
                cadenaFC = cadenaFC & "<idCliente>" & grd.TextMatrix(i, COL_V_RUC) & "</idCliente>"
                cadenaFC = cadenaFC & "<tipoComprobante>" & Format(grd.TextMatrix(i, COL_V_TIPOCOMP), "00") & "</tipoComprobante>"
                cadenaFC = cadenaFC & "<numeroComprobantes>" & grd.TextMatrix(i, COL_V_CANTRANS) & "</numeroComprobantes>"
                cadenaFC = cadenaFC & "<baseNoGraIva>" & Format(Abs(grd.ValueMatrix(i, COL_V_BASENOIVA)), "#0.00") & "</baseNoGraIva>"
                cadenaFC = cadenaFC & "<baseImponible>" & Format(Abs(grd.ValueMatrix(i, COL_V_BASE0)), "#0.00") & "</baseImponible>"
                cadenaFC = cadenaFC & "<baseImpGrav>" & Format(Abs(grd.ValueMatrix(i, COL_V_BASEIVA)), "#0.00") & "</baseImpGrav>"
                cadenaFC = cadenaFC & "<montoIva>" & Format(IIf(Abs(grd.ValueMatrix(i, COL_V_BASEIVA)) = 0, "0.00", Abs(grd.ValueMatrix(i, COL_V_BASEIVA)) * 0.12), "#0.00") & "</montoIva>"
                cadenaFCIVA = "<valorRetIva> 0.00 </valorRetIva>"
                cadenaFCIR = "<valorRetRenta> 0.00 </valorRetRenta>"
 
                'retencion IVA
                If grd.ValueMatrix(i, COL_V_TIPOCOMP) = 18 And grd.TextMatrix(i, COL_V_TIPODOC) = "R" Then
                    Set rsRet = gobjMain.EmpresaActual.ConsANRetencionVentas2008ParaXML(grd.TextMatrix(i, COL_V_RUC))
                    If rsRet.RecordCount > 0 Then

                        
                            
                            
                    End If
                Else
                End If
                cadenaFC = cadenaFC & cadenaFCIVA
                cadenaFC = cadenaFC & cadenaFCIR
                cadenaFC = cadenaFC & "</detalleVentas>"
                cadenaF = cadenaF & cadenaFC
                grd.ShowCell i, grd.ColIndex("Resultado")
                grd.TextMatrix(i, grd.ColIndex("Resultado")) = " OK "
        
SiguienteFila:
    Next i
    

        

    
    grd.ColWidth(grd.ColIndex("Resultado")) = 5000
    prg.value = 0
    

    
    cad = cadenaF
    TotalVentas = grd.ValueMatrix(grd.Rows - 1, COL_V_BASE0) + grd.ValueMatrix(grd.Rows - 1, COL_V_BASEIVA) + grd.ValueMatrix(grd.Rows - 1, COL_V_BASENOIVA)
    Exit Function
cancelado:
    GeneraArchivoATSVentasXML = False
ErrTrap:
    grd.TextMatrix(grd.Rows - 1, 2) = Err.Description
    GeneraArchivoATSVentasXML = False
End Function


'''Public Sub Exportar(tag As String)
'''    Dim file As String, NumFile As Integer, Cadena As String
'''    Dim Filas As Long, Columnas As Long, i As Long, j As Long
'''    Dim pos As Integer
''''    If grd.Rows = grd.FixedRows Then Exit Sub
'''    On Error GoTo errtrap
'''
'''        With dlg1
'''          .CancelError = True
'''          '.Filter = "Texto (Separado por coma)|*.txt|Excel 97(XLS)|*.xls"
'''          .Filter = "Texto (Separado por coma)|*.csv"
'''          .ShowSave
'''
'''          file = .filename
'''        End With
'''
'''
'''    If ExisteArchivo(file) Then
'''        If MsgBox("El nombre del archivo " & file & " ya existe desea sobreescribirlo?", vbYesNo) = vbNo Then
'''            Exit Sub
'''        End If
'''    End If
'''
'''    NumFile = FreeFile
'''
'''    Open file For Output Access Write As #NumFile
'''
'''    Cadena = ""
'''    For i = 0 To grd.Rows - 1
'''        For j = 2 To grd.Cols - 1
'''            Select Case tag          ' jeaa 04/11/03 para que se no se guarden las columnas ocultas
'''                Case "IMPCP"
'''                        If j = COL_C_NOMBRE Then j = j + 1  'columna nombre
'''            End Select
'''                If pos = 0 Then
'''                    Cadena = Cadena & grd.TextMatrix(i, j) & ","
'''                Else
'''                    Cadena = Cadena & Mid$(grd.TextMatrix(i, j), 1, pos - 1) & Mid$(grd.TextMatrix(i, j), pos + 1, Len(grd.TextMatrix(i, j)) - 1) & ","
'''                End If
'''
'''
'''        Next j
'''        Cadena = Mid(Cadena, 1, Len(Cadena) - 1)
'''        Print #NumFile, Cadena
'''        Cadena = ""
'''    Next i
'''
'''
'''    Close NumFile
'''    MsgBox "El archivo se ha exportado con éxito"
'''    Exit Sub
'''errtrap:
'''    If Err.Number <> 32755 Then
'''        MsgBox Err.Description
'''    End If
'''    Close NumFile
'''End Sub



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


Public Sub ExportaTxtPipe(ByVal grd As Control, CargaNombreArchivo As String)

    Dim file As String, NumFile As Integer, fila
    Dim r As Long, c As Long, Separador As String
    
    
    
    NumFile = FreeFile
    file = "DINAR" 'CargaNombreArchivo
    If file = "" Then Exit Sub
    Open file For Output Access Write As #NumFile
    With grd
            For r = 1 To .Rows - 1
                If .RowHidden(r) = False Then ' Filas  Ocultas
                    fila = ""
                    For c = 1 To .Cols - 1
                        If .ColHidden(c) = False Then
                            If .IsSubtotal(r) And Len(.TextMatrix(r, c)) > 0 And .TextMatrix(r, c) <> "Subtotal" Then
                                fila = fila & .ValueMatrix(r, c) & "|"
                            Else
                                fila = fila & .TextMatrix(r, c) & "|"
                            End If
                        End If
                    Next c
                    fila = Left(fila, Len(fila) - 1)
                    Print #NumFile, fila
                End If
                .TextMatrix(r, 39) = "OK"
                
                '.Cell(flexcpBackColor, r, 1, r, COL_D_RESP) = vbWhite
            Next r

    End With
    Close NumFile
    gobjMain.EmpresaActual.GrabaGNLogAccion "EXP-TXT", "Exporta Reporte a PIPE txt " & file, "RE"
End Sub

    

Private Function GeneraArchivoDinardap(ByRef cad As String) As Boolean

    Dim i As Long, j As Long, resp As Integer, c As Integer
    On Error GoTo ErrTrap
    resp = 10
    BandError = False
    GeneraArchivoDinardap = True
    grd.Refresh
    'With grd
        
        
            If grd.Rows < 1 Then
                prg.value = 0
                cadenaDD = cadenaDD & ""
                cad = cadenaDD
                GeneraArchivoDinardap = True
                GoTo SiguienteFila
            End If
            prg.max = grd.Rows - 1
            For i = 1 To grd.Rows - 1
                cadenaDD = ""
                If grd.IsSubtotal(i) Then GoTo SiguienteFila
                grd.Cell(flexcpBackColor, i, 1, i, grd.ColIndex("Resultado")) = vbWhite
                prg.value = i
                DoEvents
                If grd.TextMatrix(i, COL_D_TIPOSUJETO) = "J" Then
'                    If grd.TextMatrix(i, COL_D_SEXO) <> "N" And grd.TextMatrix(i, COL_D_ESTADOCIVIL) <> "N" And grd.TextMatrix(i, COL_D_ORIGENINGRESO) <> "N" Then
'                        grd.TextMatrix(i, COL_D_RESP + 2) = " ERROR TIPO SUJETO"
'                        grd.Cell(flexcpBackColor, i, 1, i, COL_D_RESP + 2) = vbRed
'                        BandError = True
'                    End If
                
                Else
                    If grd.TextMatrix(i, COL_D_SEXO) = "N" And grd.TextMatrix(i, COL_D_ESTADOCIVIL) = "N" And grd.TextMatrix(i, COL_D_ORIGENINGRESO) = "N" Then
                        grd.TextMatrix(i, COL_D_RESP + 2) = " ERROR TIPO SUJETO"
                        grd.Cell(flexcpBackColor, i, 1, i, COL_D_RESP + 2) = vbRed
                        BandError = True
                    End If
                    
                End If
                
                If Len(grd.TextMatrix(i, COL_D_TRANS)) = 0 Then
                        grd.TextMatrix(i, COL_D_RESP) = " ERROR NUMERO COMPROBANTE"
                        grd.Cell(flexcpBackColor, i, 1, i, COL_D_RESP) = vbRed
                        BandError = True
                End If
                
                
                For c = 1 To grd.Cols - 6
                        Select Case c
'                        Case 38
'                            cadenaDD = cadenaDD & "|"
                        Case 14, 15, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36
                            cadenaDD = cadenaDD & Format(grd.ValueMatrix(i, c), "#0.00") & "|"

                        Case Else
                            If grd.TextMatrix(i, c) = "01/01/1900" Then
                                cadenaDD = cadenaDD & "|"
                            Else
                                cadenaDD = cadenaDD & grd.TextMatrix(i, c) & "|"
                            End If
                        End Select

                Next c
                cadenaDD = Left(cadenaDD, Len(cadenaDD) - 1)
                Print #NumFile, cadenaDD

                grd.ShowCell i, COL_D_RESP
                If InStr(1, grd.TextMatrix(i, COL_D_RESP), "ERROR") = 0 Then
                    grd.TextMatrix(i, COL_D_RESP) = " OK "
                    grd.Cell(flexcpBackColor, i, 1, i, COL_D_RESP) = &HFFFFC0
                End If
            GoTo SiguienteFila
    Exit Function
SiguienteFila:
    Next i
    grd.ColWidth(COL_D_RESP) = 5000
    prg.value = 0
    cad = cadenaDD
    
Exit Function
cancelado:
    GeneraArchivoDinardap = False
ErrTrap:
    grd.TextMatrix(grd.Rows - 1, 2) = Err.Description
    GeneraArchivoDinardap = False
End Function



Private Sub AbrirArchivoExcel()
    Dim i As Long
    On Error GoTo ErrTrap
    With dlg1
        .CancelError = True
'        .Filter = "Texto (Separado por coma)|*.txt|Excel 97(XLS)|*.xls"
        .Filter = "Texto (Separado por coma)|*.txt"
        .flags = cdlOFNFileMustExist

        If Len(.filename) = 0 Then          'Solo por primera vez, ubica a la carpeta de la aplicación
            .filename = App.Path & "\*.txt"
        End If
        
        .ShowOpen

        LeerArchivo (dlg1.filename)
        AjustarAutoSize grd, -1, -1
    End With
    Exit Sub
ErrTrap:
    If Err.Number <> 32755 Then DispErr
    Exit Sub
End Sub

Private Sub LeerArchivo(ByVal archi As String)
    Select Case UCase$(Right$(archi, 4))
        Case ".TXT"
            ConfigCols "IMPFC"
            VisualizarTexto archi
            ConfigCols "IMPFC"
'            InsertarColumnas
        Case ".XLS"
'            VisualizarExcel archi
        Case Else
        End Select
End Sub


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
        Case "VENTASLOCUTORIOS"         ' tienen el separador como ;
            Separador = ";"
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
    End If
    grd.Sort = flexSortUseColSort

' poner numero
    GNPoneNumFila grd, False
    
    grd.Redraw = flexRDDirect
    AjustarAutoSize grd, -1, -1
    
   If grd.Rows > 0 Then
        grd.Cell(flexcpBackColor, 1, COL_D_NOMBRE, grd.Rows - 1, COL_D_NOMBRE) = &H80000018
        grd.Cell(flexcpBackColor, 1, COL_D_TRANS, grd.Rows - 1, COL_D_TRANS) = &H80000018
        grd.Cell(flexcpBackColor, 1, COL_D_FECHAVENCI, grd.Rows - 1, COL_D_FECHAVENCI) = &H80000018
        grd.Cell(flexcpBackColor, 1, COL_D_FECHAPAGO, grd.Rows - 1, COL_D_FECHAPAGO) = &H80000018
    End If
    
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


Private Sub AbrirArchivoPipe()
    Dim i As Long
    On Error GoTo ErrTrap
    With dlg1
        .CancelError = True
'        .Filter = "Texto (Separado por coma)|*.txt|Excel 97(XLS)|*.xls"
        .Filter = "Texto (Separado por coma)|*.txt"
        .flags = cdlOFNFileMustExist

        If Len(.filename) = 0 Then          'Solo por primera vez, ubica a la carpeta de la aplicación
            .filename = App.Path & "\*.txt"
        End If
        
        .ShowOpen

        LeerArchivoPipe (dlg1.filename)
        AjustarAutoSize grd, -1, -1
    End With
    Exit Sub
ErrTrap:
    If Err.Number <> 32755 Then DispErr
    Exit Sub
End Sub


Private Sub LeerArchivoPipe(ByVal archi As String)
    Select Case UCase$(Right$(archi, 4))
        Case ".TXT"
            ConfigCols "IMPFC"
            VisualizarTextoPipe archi
            ConfigCols "IMPFC"
'            InsertarColumnas
        Case ".XLS"
'            VisualizarExcel archi
        Case Else
        End Select
End Sub

Private Sub VisualizarTextoPipe(ByVal archi As String)
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
            Separador = "|"
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
    End If
    grd.Sort = flexSortUseColSort

' poner numero
    GNPoneNumFila grd, False
    
    grd.Redraw = flexRDDirect
    AjustarAutoSize grd, -1, -1
    
   If grd.Rows > 0 Then
        grd.Cell(flexcpBackColor, 1, COL_D_NOMBRE, grd.Rows - 1, COL_D_NOMBRE) = &H80000018
        grd.Cell(flexcpBackColor, 1, COL_D_TRANS, grd.Rows - 1, COL_D_TRANS) = &H80000018
        grd.Cell(flexcpBackColor, 1, COL_D_FECHAVENCI, grd.Rows - 1, COL_D_FECHAVENCI) = &H80000018
        grd.Cell(flexcpBackColor, 1, COL_D_FECHAPAGO, grd.Rows - 1, COL_D_FECHAPAGO) = &H80000018
    End If
    
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

Private Sub SubTotalizar(col As Long)
    Dim i As Long
    With grd
        For i = 1 To .Cols - 1
            'If i = COL_C_CODTIPOCOMP Then i = i + 1
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
            'If i = COL_C_CODTIPOCOMP Then i = i + 1
            If grd.ColData(i) = "SubTotal" Then
                
                .subtotal flexSTSum, -1, i, "#,#0.00", .BackColorSel, vbYellow, vbBlack, "Total"
            End If
        Next i
'        .subtotal flexSTCount, -1, COL_C_CODTIPOCOMP, "#,#0", .BackColorSel, vbYellow, vbBlack, "Total"
    End With
End Sub


Private Function BuscarVentasDinardap()
    Dim fecha1 As Date
    Dim fecha2 As Date

    On Error GoTo ErrTrap
        With grd
        .Redraw = False
        .Rows = .FixedRows
        If Not frmB_Trans.Inicio(gobjMain, "IMPDD", dtpPeriodo.value) Then
            grd.SetFocus
        End If
        
        If DatePart("m", fecha) = 12 Then
            fecha1 = "01/" & DatePart("m", dtpPeriodo.value) & "/" & DatePart("yyyy", dtpPeriodo.value)
            fecha2 = DateAdd("yyyy", 1, DateAdd("d", -1, ("01/" & DatePart("m", DateAdd("m", 1, dtpPeriodo.value)) & "/" & DatePart("yyyy", dtpPeriodo.value))))
        Else
            fecha1 = "01/" & DatePart("m", dtpPeriodo.value) & "/" & DatePart("yyyy", dtpPeriodo.value)
            fecha2 = DateAdd("d", -1, ("01/" & DatePart("m", DateAdd("m", 1, dtpPeriodo.value)) & "/" & DatePart("yyyy", dtpPeriodo.value)))
        End If
        gobjMain.objCondicion.fecha1 = DateAdd("d", -1, DateAdd("m", 1, fecha1))
        
        gobjMain.objCondicion.NumDias1 = ntxDiasGracia.value
        
        MiGetRowsRep gobjMain.EmpresaActual.ConsVentasDINARDAP2015New(), grd


        'GeneraArchivo

        ConfigCols "IMPFC"
        AjustarAutoSize grd, -1, -1
        AjustarAutoSize grd, -1, -1
        grd.ColWidth(0) = "500"
        
        SubTotalizar COL_D_NOMBRE 'COL_D_TIPOTRANS
        
'        grd.MergeCol(1) = True
'
'        .SubTotal flexSTSum, 1, 15, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        .SubTotal flexSTSum, 1, 16, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        .SubTotal flexSTSum, 1, 23, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        .SubTotal flexSTSum, 1, 25, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        .SubTotal flexSTSum, 1, 26, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        .SubTotal flexSTSum, 1, 27, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        .SubTotal flexSTSum, 1, 28, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        .SubTotal flexSTSum, 1, 29, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        .SubTotal flexSTSum, 1, 30, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        .SubTotal flexSTSum, 1, 31, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        .SubTotal flexSTSum, 1, 32, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        .SubTotal flexSTSum, 1, 33, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        .SubTotal flexSTSum, 1, 34, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        .SubTotal flexSTSum, 1, 37, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
        

        GNPoneNumFila grd, False


        .Redraw = True
   End With

    Exit Function
ErrTrap:
    grd.Redraw = True
    DispErr
    Exit Function
End Function


Private Sub SumarDiasFechaVencimiento()
    Dim i As Integer, diascredito  As Integer, diasmora  As Integer, valor As Currency
    For i = 1 To grd.Rows - 1
        If Not grd.IsSubtotal(i) Then
    
                    valor = grd.ValueMatrix(i, COL_D_SALDO)
                        grd.TextMatrix(i, COL_D_FECHAVENCI) = DateAdd("d", ntxDiasGracia, grd.TextMatrix(i, COL_D_FECHAVENCI))
                        grd.Cell(flexcpBackColor, i, COL_D_FECHAVENCI, i, COL_D_FECHAVENCI) = vbRed
                        grd.TextMatrix(i, COL_D_SALDOXVMAS_360) = ""
                        grd.TextMatrix(i, COL_D_SALDOXV181_360) = ""
                        grd.TextMatrix(i, COL_D_SALDOXV91_180) = ""
                        grd.TextMatrix(i, COL_D_SALDOXV31_90) = ""
                        grd.TextMatrix(i, COL_D_SALDOXV1_30) = ""
                        grd.TextMatrix(i, COL_D_SALDOVEMAS_360) = ""
                        grd.TextMatrix(i, COL_D_SALDOVE181_360) = ""
                        grd.TextMatrix(i, COL_D_SALDOVE91_180) = ""
                        grd.TextMatrix(i, COL_D_SALDOVE31_90) = ""
                        grd.TextMatrix(i, COL_D_SALDOVE1_30) = ""
                        
                        
'                        grd.TextMatrix(i, COL_D_FECHAVENCI) = CDate(fecha)
'                        grd.TextMatrix(i, COL_D_FECHAEXIGIBLE) = CDate(fecha)
                        diascredito = DateDiff("d", CDate(grd.TextMatrix(i, COL_D_FECHACONCESION)), CDate(grd.TextMatrix(i, COL_D_FECHAVENCI)))
                        grd.TextMatrix(i, COL_D_DIASCREDITO) = diascredito
                        grd.TextMatrix(i, COL_D_DIASCREDITO + 1) = diascredito
                        diasmora = DateDiff("d", CDate(grd.TextMatrix(i, COL_D_FECHAVENCI)), grd.TextMatrix(i, COL_D_FECHADATOS))
                        
                        If diasmora > 0 Then
                            grd.TextMatrix(i, COL_D_DIASMORA) = diasmora
                            If diasmora > 360 Then
                                grd.TextMatrix(i, COL_D_SALDOVEMAS_360) = valor
                            ElseIf diasmora > 180 And diascredito <= 360 Then
                                grd.TextMatrix(i, COL_D_SALDOVE181_360) = valor
                            ElseIf diasmora > 90 And diascredito <= 180 Then
                                grd.TextMatrix(i, COL_D_SALDOVE91_180) = valor
                            ElseIf diasmora > 30 And diascredito <= 90 Then
                                grd.TextMatrix(i, COL_D_SALDOVE31_90) = valor
                            Else
                                grd.TextMatrix(i, COL_D_SALDOVE1_30) = valor
                            End If
                            
                        Else
                            grd.TextMatrix(i, COL_D_DIASMORA) = "0"
                            grd.TextMatrix(i, COL_D_MONTOMOROSIDAD) = "0.00"
                            diascredito = Abs(diascredito)
                            If diascredito > 360 Then
                                grd.TextMatrix(i, COL_D_SALDOXVMAS_360) = valor
                            ElseIf diascredito > 180 And diascredito <= 360 Then
                                grd.TextMatrix(i, COL_D_SALDOXV181_360) = valor
                            ElseIf diascredito > 90 And diascredito <= 180 Then
                                grd.TextMatrix(i, COL_D_SALDOXV91_180) = valor
                            ElseIf diascredito > -30 And diascredito <= 90 Then
                                grd.TextMatrix(i, COL_D_SALDOXV31_90) = valor
                            Else
                                grd.TextMatrix(i, COL_D_SALDOXV1_30) = valor
                            End If
                            
                        End If
                        
                        
                        If grd.ValueMatrix(i, COL_D_DIASCREDITO) > 0 Then
                        End If
            
            
            
        End If
    Next i
    Totalizar
'        grd.Redraw = flexRDNone
'        grd.Refresh
'
    
        
'        grd.SubTotal flexSTSum, , COL_D_VALOR, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        grd.SubTotal flexSTSum, 1, COL_D_SALDO, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        grd.SubTotal flexSTSum, 1, COL_D_MONTOMOROSIDAD, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        grd.SubTotal flexSTSum, 1, COL_D_SALDOXV1_30, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        grd.SubTotal flexSTSum, 1, COL_D_SALDOXV31_90, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        grd.SubTotal flexSTSum, 1, COL_D_SALDOXV91_180, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        grd.SubTotal flexSTSum, 1, COL_D_SALDOXV181_360, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        grd.SubTotal flexSTSum, 1, COL_D_SALDOXVMAS_360, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        grd.SubTotal flexSTSum, 1, COL_D_SALDOVE1_30, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        grd.SubTotal flexSTSum, 1, COL_D_SALDOVE31_90, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        grd.SubTotal flexSTSum, 1, COL_D_SALDOVE91_180, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        grd.SubTotal flexSTSum, 1, COL_D_SALDOVE181_360, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
'        grd.SubTotal flexSTSum, 1, COL_D_SALDOVEMAS_360, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True
        'grd.SubTotal flexSTSum, 1, 37, "#,#0.00", grd.GridColor, vbBlack, , "Subtotal", , True


End Sub

Private Sub CambiaDiasPlazo()
    Dim fecha1 As Date
    Dim fecha2 As Date

    On Error GoTo ErrTrap
        With grd
        'If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("DiasGraciaDINARDAP")) > 0 Then
             gobjMain.EmpresaActual.GNOpcion.AsignarValor "DiasGraciaDINARDAP", ntxDiasGracia.value
            gobjMain.EmpresaActual.GNOpcion.GrabarGNOpcion2

        
        'gobjMain.EmpresaActual.CambiaDiasPlazoDINARDAP2015 (ntxDiasGracia.value)



   End With

    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub


Public Sub ExportaExcel(ByVal titulo As String)
    'tipo=0 Roles de Pagos; tipo=1 Reporte Bancos y Provisiones
    
    If grd.Rows < 2 Then
        MsgBox "No existe filas para exportar a Excel"
        Exit Sub
    End If

    
    Set ex = New Excel.Application  'Crea un instancia nueva de excel
    Set wkb = ex.Workbooks.Add  'Insertar un libro nuevo
    Set ws = ex.Worksheets.Add  'Inserta una nueva hoja
    With ws
        .Name = titulo
        .Range("A1").Font.Name = "Times Roman"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Cells(1) = gobjMain.EmpresaActual.GNOpcion.NombreEmpresa
    End With
        Exportar titulo
    ex.Visible = True
    ws.Activate
    Set ws = Nothing
    Set wkb = Nothing
    Set ex = Nothing
    'ex.Quit
End Sub

Private Sub Exportar(ByVal titulo As String)
        Dim fila As Long, col As Long, i As Long, j As Long
    Dim v() As Long, mayor As Long
    Dim NumCol As Integer
    Dim fmt As String
    prg.min = 0
    prg.max = grd.Rows - 1
    MensajeStatus "Está Exportando  a Excel ...", vbHourglass
    With ws
        fila = 2
        .Range("H1").Font.Name = "Arial"
        .Range("H1").Font.Size = 10
        .Range("H1").Font.Bold = True
        .Cells(fila, 1) = titulo
        
        .PageSetup.PaperSize = xlPaperLetter 'Tamaño del papel (carta)
        .PageSetup.BottomMargin = Application.CentimetersToPoints(1.5) 'Margen Superior
        .PageSetup.TopMargin = Application.CentimetersToPoints(1) 'Margen Inferior
'        .Range(.Cells(1, 13), .Cells(500, 23)).NumberFormat = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda(fmt)    'Establece el formato para los números
        .Range("A2:AZ10000").Font.Name = "Arial"    'Tipo de letra para toda la hoja
        .Range("A2:AZ10000").Font.Size = 7          'Tamaño de la letra
        
        fila = fila + 1
        NumCol = 0
        For i = 1 To grd.Cols - 1
            NumCol = NumCol + 1
            .Cells(fila, NumCol) = grd.TextMatrix(0, i) 'cabeceras
            ReDim Preserve v(NumCol)
            v(NumCol - 1) = 0
        Next i
                
        .Range(.Cells(fila, 1), .Cells(fila, NumCol)).Font.Bold = True
        .Range(.Cells(fila, 1), .Cells(fila, NumCol)).Borders.LineStyle = 12
        .Range(.Cells(fila, 3), .Cells(fila, grd.Rows - 1)).HorizontalAlignment = xlHAlignLeft
        For i = 1 To grd.Rows - 2
            prg.value = i
            fila = fila + 1
            If grd.IsSubtotal(i) = True Then
                i = i + 1
            End If
'                .Range(.Cells(fila, 1), .Cells(fila, NumCol)).Font.Bold = True

 '           Else
                j = 1
                mayor = 0

                For col = 1 To grd.Cols - 1
                
'                    Select Case Me.tag
'                        Case "F104"
'                            Select Case col
'                                'Case 1: mayor = 2
'                                Case 1, 3, 4, 6, 8, 10, 11: mayor = 4
'                                Case 2, 5, 7, 9: mayor = 25
'                            End Select
'                        Case "F103"
'                            Select Case col
'                                Case 2: mayor = 50
'                                Case 3, 4, 6, 8, 9: mayor = 4
'                                Case 5, 7, 10: mayor = 15
'                                Case 1: mayor = 10
'                            End Select
'                       End Select
                Select Case col
                    Case 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 16, 17, 18, 37
                        .Cells(fila, j) = "'" & grd.TextMatrix(i, col)
                    Case Else
                        .Cells(fila, j) = grd.TextMatrix(i, col)
                End Select

'                        mayor = Len(grd.TextMatrix(i, Col)) 'Para ajustar el ancho de columnas
                        If mayor > v(j - 1) Then            'de acuerdo a la celda más grande
                            .Columns(j).ColumnWidth = mayor '13/11/2000 ---> Angel P.
                            v(j - 1) = mayor
                        End If
                        j = j + 1
                Next col
'            End If
            .Range(.Cells(fila, 1), .Cells(fila, NumCol)).Borders.LineStyle = 1
        Next i
    End With
     prg.value = prg.min
     MensajeStatus "Listo", vbDefault
End Sub

