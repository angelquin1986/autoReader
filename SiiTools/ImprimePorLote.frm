VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{50067EB3-D6AF-11D3-8297-000021C5085D}#1.0#0"; "NTextBox.ocx"
Object = "{ED5A9B02-5BDB-48C7-BAB1-642DCC8C9E4D}#2.0#0"; "SelFold.ocx"
Begin VB.Form frmImprimePorLote 
   Caption         =   "Impresión por lote"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6660
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   6660
   WindowState     =   2  'Maximized
   Begin VB.Frame FraVerificaEXistXML 
      Height          =   1095
      Left            =   8580
      TabIndex        =   42
      Top             =   120
      Visible         =   0   'False
      Width           =   8835
      Begin VB.CommandButton cmdExaminarCarpetaRuta 
         Caption         =   "..."
         Height          =   320
         Left            =   5100
         TabIndex        =   44
         Top             =   180
         Width           =   372
      End
      Begin VB.TextBox txtCarpeta 
         Height          =   320
         Left            =   960
         TabIndex        =   43
         Text            =   "c:\"
         Top             =   180
         Width           =   4170
      End
      Begin SelFold.SelFolder slf 
         Left            =   4320
         Top             =   0
         _ExtentX        =   1349
         _ExtentY        =   265
         Title           =   "Seleccione una carpeta"
         Caption         =   "Selección de carpeta"
         RootFolder      =   "\"
         Path            =   "C:\VBPROG_ESP\SII\SELFOLD"
      End
      Begin VB.Label Label5 
         Caption         =   "Ubicacion:"
         Height          =   255
         Left            =   180
         TabIndex        =   45
         Top             =   240
         Width           =   870
      End
   End
   Begin VB.Frame FraTransElect 
      Caption         =   "Cod.&Trans"
      Height          =   1095
      Left            =   2340
      TabIndex        =   40
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      Begin VB.ListBox lstTrans 
         Columns         =   3
         Height          =   795
         IntegralHeight  =   0   'False
         Left            =   120
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   41
         Top             =   240
         Width           =   1635
      End
   End
   Begin VB.Frame fraFecha 
      Caption         =   "&Fecha (desde - hasta)"
      Height          =   1092
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   2235
      Begin MSComCtl2.DTPicker dtpFecha1 
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1995
         _ExtentX        =   3519
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
         Format          =   115802113
         CurrentDate     =   36348
      End
      Begin MSComCtl2.DTPicker dtpFecha2 
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1995
         _ExtentX        =   3519
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
         Format          =   115802113
         CurrentDate     =   36348
      End
   End
   Begin VB.Frame FraAcopiador 
      Caption         =   "Acopiador"
      Height          =   1035
      Left            =   8580
      TabIndex        =   37
      Top             =   180
      Visible         =   0   'False
      Width           =   4515
      Begin FlexComboProy.FlexCombo fcbAcopiador 
         Height          =   345
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   3795
         _ExtentX        =   6694
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
   Begin VB.Frame framConfig 
      Caption         =   "Configuración"
      Height          =   1095
      Left            =   8580
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   4995
      Begin NTextBoxProy.NTextBox ntxMargSup 
         Height          =   300
         Left            =   1620
         TabIndex        =   31
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
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
      Begin NTextBoxProy.NTextBox ntxMargIzq 
         Height          =   300
         Left            =   1620
         TabIndex        =   32
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
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
      Begin NTextBoxProy.NTextBox ntxNUmEtiq 
         Height          =   300
         Left            =   3780
         TabIndex        =   33
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
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
      Begin VB.Label Label4 
         Caption         =   "Margen Superior"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Margen Izquierdo"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   660
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "# Etiquetas"
         Height          =   195
         Left            =   2640
         TabIndex        =   34
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.Frame FraTipoTrabajo 
      Caption         =   "Tipo Trabajo"
      Height          =   1095
      Left            =   6240
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   2355
      Begin VB.ComboBox cboTipoTrabajo 
         Height          =   315
         ItemData        =   "ImprimePorLote.frx":0000
         Left            =   120
         List            =   "ImprimePorLote.frx":0013
         TabIndex        =   25
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.Frame FraConFigEgreso 
      Caption         =   "Datos para Egresos"
      Height          =   1095
      Left            =   6240
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   8115
      Begin VB.TextBox txtEgreso 
         Height          =   315
         Left            =   2220
         TabIndex        =   20
         Top             =   240
         Width           =   5415
      End
      Begin VB.TextBox txtCheque 
         Height          =   315
         Left            =   2220
         TabIndex        =   19
         Top             =   540
         Width           =   5415
      End
      Begin VB.CommandButton cmdExplorar 
         Caption         =   "..."
         Height          =   310
         Left            =   7620
         TabIndex        =   18
         Top             =   240
         Width           =   372
      End
      Begin VB.CommandButton cmdExplorarCH 
         Caption         =   "..."
         Height          =   310
         Left            =   7620
         TabIndex        =   17
         Top             =   540
         Width           =   372
      End
      Begin MSComDlg.CommonDialog dlg1 
         Left            =   6180
         Top             =   120
         _ExtentX        =   688
         _ExtentY        =   688
         _Version        =   393216
         CancelError     =   -1  'True
         DefaultExt      =   "mdb"
         DialogTitle     =   "Destino de exportación"
      End
      Begin VB.Label Label6 
         Caption         =   "Lib. Impresion Egreso"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   300
         Width           =   1875
      End
      Begin VB.Label Label7 
         Caption         =   "Lib. Impresion Cheque"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.OptionButton optAsiento 
      Caption         =   "Asiento  "
      Height          =   192
      Left            =   5400
      TabIndex        =   15
      Top             =   1440
      Width           =   1212
   End
   Begin VB.OptionButton optTrans 
      Caption         =   "Transacción  "
      Height          =   192
      Left            =   3960
      TabIndex        =   14
      Top             =   1440
      Value           =   -1  'True
      Width           =   1692
   End
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   852
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   6660
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3855
      Width           =   6660
      Begin VB.CommandButton cmdImprimiCH 
         Caption         =   "&Imprimir Cheques"
         Enabled         =   0   'False
         Height          =   372
         Left            =   3780
         TabIndex        =   23
         Top             =   0
         Width           =   1452
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   5280
         TabIndex        =   12
         Top             =   0
         Width           =   1212
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   372
         Left            =   2280
         TabIndex        =   11
         Top             =   0
         Width           =   1452
      End
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   13
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
      Left            =   1584
      TabIndex        =   8
      Top             =   1320
      Width           =   1452
   End
   Begin VB.Frame fraCodTrans 
      Caption         =   "Cod.&Trans."
      Height          =   1092
      Left            =   2322
      TabIndex        =   3
      Top             =   120
      Width           =   1932
      Begin FlexComboProy.FlexCombo fcbTrans 
         Height          =   348
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1452
         _ExtentX        =   2566
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
      Left            =   4260
      TabIndex        =   5
      Top             =   120
      Width           =   1932
      Begin VB.CommandButton cmdCargarTrans 
         Caption         =   "Cargar"
         Height          =   372
         Left            =   180
         TabIndex        =   39
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
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
   Begin VB.Frame FraBuscaOrden 
      Caption         =   "Busca Ingreso Carcasa"
      Height          =   1095
      Left            =   60
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   2235
      Begin VB.CommandButton cmdCargar 
         Caption         =   "Cargar"
         Height          =   372
         Left            =   120
         TabIndex        =   29
         Top             =   660
         Width           =   1995
      End
      Begin VB.TextBox txtNumDocRef 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   900
         TabIndex        =   27
         Top             =   300
         Width           =   1212
      End
      Begin VB.Label Label1 
         Caption         =   "# Ingreso"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmImprimePorLote"
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
Private Const COL_NUMDOCREF = 6         '*** MAKOTO 07/feb/01 Agregado
Private Const COL_NOMBRE = 7            '*** MAKOTO 07/feb/01 Agregado
Private Const COL_DESC = 8
Private Const COL_CENTROCOSTO = 9
Private Const COL_ESTADO = 10
Private Const COL_RESULTADO = 11

Private Const COL_C_TRANS = 3
Private Const COL_C_DESC = 4
Private Const COL_C_BENEFICIARIO = 5
Private Const COL_C_BANCO = 6
Private Const COL_C_NUMCHEQUE = 7
Private Const COL_C_VALOR = 8
Private Const COL_C_ESTADO = 9
Private Const COL_C_RESULTADO = 10


Private Const COL_E_NUMING = 1
Private Const COL_E_ORDEN = 2
Private Const COL_E_TIKET = 3
Private Const COL_E_MARCA = 4
Private Const COL_E_TAMANIO = 5
Private Const COL_E_SERIE = 6
Private Const COL_E_DISENIO = 7
Private Const COL_E_TRABAJO = 8
Private Const COL_E_TRANS = 9
Private Const COL_E_FECHA = 10
Private Const COL_E_CODCLI = 11
Private Const COL_E_NOMCLI = 12
Private Const COL_E_VENDE = 13
Private Const COL_E_GAR = 14
Private Const COL_E_RESULTADO = 15


Private Const MSG_NG = "Error en impresión."
Private mProcesando As Boolean
Private mCancelado As Boolean
Private mTag As String
Private mobj As Object
Private mobjImp As Object
Private NumFile As Integer

Public Sub Inicio()
    Dim i As Integer
    On Error GoTo errtrap
    
    Me.Show
    Me.ZOrder
    dtpFecha1.value = gobjMain.EmpresaActual.GNOpcion.FechaInicio
    dtpFecha2.value = Date
    CargaTrans
    Exit Sub
errtrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub CargaTrans()
    'Carga la lista de transacción
    fcbTrans.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False)
    fcbAcopiador.SetData gobjMain.EmpresaActual.ListaGNObra(False)
End Sub



Private Function Imprimir() As Boolean
    Dim s As String, tid As Long, i As Long, x As Single, res As String
    Dim gnc As GNComprobante, cambiado As Boolean, cntError As Long
    
    On Error GoTo errtrap

    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdBuscar.Enabled = False
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    
    'Limpia los mensajes
    For i = grd.FixedRows To grd.Rows - 1
        grd.TextMatrix(i, COL_C_RESULTADO) = ""
    Next i
    
    For i = grd.FixedRows To grd.Rows - 1
        DoEvents
        If mCancelado Then
            MsgBox "El proceso fue cancelado."
            Exit For
        End If
        
        prg1.value = i
        grd.Row = i
        x = grd.CellTop                 'Para visualizar la celda actual
        
        tid = grd.ValueMatrix(i, COL_TID)
        grd.TextMatrix(i, COL_C_RESULTADO) = "Procesando ..."
        grd.Refresh
        
        'Recupera la transaccion
        Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
        If Not (gnc Is Nothing) Then
            'Si la transacción no está anulado
            If gnc.Estado <> ESTADO_ANULADO Then
'                'Forzar recuperar todos los datos de transacción
'                ' para que no se pierdan al grabar de nuveo
'                gnc.RecuperaDetalleTodo
            
                'Imprime la transaccion o asiento contable
                If Me.Caption = "Busca Transacciones para Impresion de Cheques" Then
                    res = ImprimeTransLote(gnc, optAsiento.value)
                Else
                    res = ImprimeTrans(gnc, optAsiento.value)
                End If
                If Len(res) = 0 Then
                    grd.TextMatrix(i, COL_C_RESULTADO) = "Enviado."
                Else
                    grd.TextMatrix(i, COL_C_RESULTADO) = res
                    cntError = cntError + 1
                End If
                            
            'Si la transaccion está anulado
            Else
                grd.TextMatrix(i, COL_C_RESULTADO) = "Anulado."
                cntError = cntError + 1
            End If
        Else
            grd.TextMatrix(i, COL_C_RESULTADO) = "No pudo recuperar la transación."
            cntError = cntError + 1
        End If
    Next i
    
    Screen.MousePointer = 0
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdImprimir.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    
    'Si algúna transaccion no se imprimió, avisa
    If cntError Then
        MsgBox "No se pudo imprimir " & cntError & " transacciones.", vbInformation
    End If
    
    Imprimir = True
    Exit Function
errtrap:
    Screen.MousePointer = 0
    DispErr
    prg1.value = prg1.min
    Exit Function
End Function


Private Sub cmdBuscar_Click()
    Dim v As Variant, obj As Object, sql As String, s As String
    On Error GoTo errtrap
    With gobjMain.objCondicion
        .fecha1 = dtpFecha1.value
        .fecha2 = dtpFecha2.value
        .CodTrans = fcbTrans.Text
        .NumTrans1 = Val(txtNumTrans1.Text)
        .NumTrans2 = Val(txtNumTrans2.Text)
        
        'Estados no incluye anulados
        .EstadoBool(ESTADO_NOAPROBADO) = True
        .EstadoBool(ESTADO_APROBADO) = True
        .EstadoBool(ESTADO_DESPACHADO) = True
        .EstadoBool(ESTADO_ANULADO) = False
    End With
    If Me.Caption = "Busca Transacciones con problemas de Relación" Then
        Set obj = gobjMain.EmpresaActual.ConsGNTransError()
        If Not obj.EOF Then
            v = MiGetRows(obj)
            
            grd.Redraw = flexRDNone
            grd.LoadArray v
            ConfigColsTRansErradas
            grd.Redraw = flexRDDirect
        Else
            grd.Rows = grd.FixedRows
            ConfigColsTRansErradas
        End If
    ElseIf Me.Caption = "Busca Transacciones con problemas de Clientes/Proveedores en Pagos/Cobros" Then
        Set obj = gobjMain.EmpresaActual.ConsGNErrorPagosCobros()
        If Not obj.EOF Then
            v = MiGetRows(obj)
            
            grd.Redraw = flexRDNone
            grd.LoadArray v
            ConfigColsTransErradasCobroPagos
            grd.Redraw = flexRDDirect
            cmdImprimir.Enabled = True
        Else
            grd.Rows = grd.FixedRows
            ConfigColsTransErradasCobroPagos
        End If
    ElseIf Me.Caption = "Busca Transacciones para Impresion de Cheques" Then
        Set obj = gobjMain.EmpresaActual.ConsTSTransImpresionCheques()
        If Not obj.EOF Then
            v = MiGetRows(obj)
            
            grd.Redraw = flexRDNone
            grd.LoadArray v
            ConfigColsImprimeCheques
            grd.Redraw = flexRDDirect
            
            cmdImprimiCH.Enabled = True
            cmdImprimir.Enabled = True
            cmdImprimir.SetFocus
            
            s = txtEgreso.Text
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "ImpLote_LibImpPago", s
    
            s = txtCheque.Text
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "ImpLote_LibImpCheque", s
            
            'Graba en la base
            gobjMain.EmpresaActual.GNOpcion.Grabar

            
        Else
            grd.Rows = grd.FixedRows
            ConfigColsImprimeCheques
        End If
    ElseIf Me.Caption = "Busca Transacciones para Impresion de Etiquetas" Then
        gobjMain.objCondicion.nivel = cboTipoTrabajo.ListIndex
        Set obj = gobjMain.EmpresaActual.ConsIVTransImpresionEtiquetas()

        If Not obj.EOF Then
            v = MiGetRows(obj)
            Set mobj = obj.Clone
            mobj.MoveFirst
            grd.Redraw = flexRDNone
            grd.LoadArray v
            ConfigColsImprimeEtiketas
            grd.Redraw = flexRDDirect
            
            cmdImprimiCH.Enabled = True
            cmdImprimir.Enabled = True
            cmdImprimir.SetFocus
            
            s = fcbTrans.KeyText
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "ImpLote_Etiketas", s
    
            'Graba en la base
            gobjMain.EmpresaActual.GNOpcion.Grabar

        Else
            grd.Rows = grd.FixedRows
            ConfigColsImprimeEtiketas
        End If
    ElseIf Me.Caption = "Impresion Relacion Dependencia" Then 'AUC 2012
        If Len(txtNumTrans1.Text) = 0 Then MsgBox "Debe ingresar el numero de transaccion ", vbCritical: Exit Sub
        Set obj = gobjMain.EmpresaActual.ConsRelacionDep(fcbTrans.KeyText, txtNumTrans1.Text)
        If Not obj.EOF Then
            v = MiGetRows(obj)
            grd.Redraw = flexRDNone
            grd.LoadArray v
            ConfigColsRelDep
            grd.Redraw = flexRDDirect
            cmdImprimir.Enabled = True
        Else
            grd.Rows = grd.FixedRows
            ConfigColsRelDep
        End If
    ElseIf Me.Caption = "Busca Transacciones para Cambio de Acopiador" Then
        gobjMain.objCondicion.nivel = cboTipoTrabajo.ListIndex
        Set obj = gobjMain.EmpresaActual.ConsIVTransImpresionEtiquetas()

        If Not obj.EOF Then
            v = MiGetRows(obj)
            Set mobj = obj.Clone
            mobj.MoveFirst
            grd.Redraw = flexRDNone
            grd.LoadArray v
            ConfigColsCambiaAcopiador
            grd.Redraw = flexRDDirect
            
            cmdImprimiCH.Enabled = True
            cmdImprimir.Enabled = True
            cmdImprimir.SetFocus
            
            's = fcbTrans.KeyText
            'gobjMain.EmpresaActual.GNOpcion.AsignarValor "ImpLote_Etiketas", s
    
            'Graba en la base
            gobjMain.EmpresaActual.GNOpcion.Grabar
        Else
            grd.Rows = grd.FixedRows
            ConfigColsImprimeEtiketas
        End If
            
    ElseIf Me.Caption = "Generar Archivo RIDE en formato pdf" Then
        's = PreparaTransParaGnopcion(gobjMain.objCondicion.CodTrans)
        gobjMain.objCondicion.CodTrans = PreparaCodTrans
        Set obj = gobjMain.EmpresaActual.ConsGNTrans2(True)
        If Not obj.EOF Then
            v = MiGetRows(obj)
            
            grd.Redraw = flexRDNone
            grd.LoadArray v
            ConfigColsRIDE
            grd.Redraw = flexRDDirect

        Else
            grd.Rows = grd.FixedRows
            ConfigCols
        End If
    
        cmdImprimir.Enabled = True
        cmdImprimir.SetFocus
    
    ElseIf Me.Caption = "Generar Archivo xml" Then
        's = PreparaTransParaGnopcion(gobjMain.objCondicion.CodTrans)
        gobjMain.objCondicion.CodTrans = PreparaCodTrans
        Set obj = gobjMain.EmpresaActual.ConsGNTransCompElec(True, True)
        If Not obj.EOF Then
            v = MiGetRows(obj)
            
            grd.Redraw = flexRDNone
            grd.LoadArray v
            ConfigColsRIDE
            grd.Redraw = flexRDDirect
        Else
            grd.Rows = grd.FixedRows
            ConfigCols
        End If
        cmdImprimir.Enabled = True
        cmdImprimir.SetFocus
    ElseIf Me.Caption = "Recupera Archivo xml desde Base Datos" Then
        's = PreparaTransParaGnopcion(gobjMain.objCondicion.CodTrans)
        gobjMain.objCondicion.CodTrans = PreparaCodTrans
        If optTrans.value Then
            gobjMain.objCondicion.Cliente = True
        End If
        If optAsiento.value = True Then
            gobjMain.objCondicion.Proveedor = True
        End If
        
        
        Set obj = gobjMain.EmpresaActual.ConsGNTrans2conArchivoXML(True)
        If Not obj.EOF Then
            v = MiGetRows(obj)
            
            grd.Redraw = flexRDNone
            grd.LoadArray v
            ConfigColsRIDE
            grd.Redraw = flexRDDirect
        Else
            grd.Rows = grd.FixedRows
            ConfigCols
        End If
        cmdImprimir.Enabled = True
        cmdImprimir.SetFocus
    
    Else
        Set obj = gobjMain.EmpresaActual.ConsGNTrans2(True)
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
    
        cmdImprimir.Enabled = True
        cmdImprimir.SetFocus
    End If
    Exit Sub
errtrap:
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
        .ColHidden(COL_NOMBRE) = False  'True
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
End Sub

Private Sub cmdCancelar_Click()
    If mProcesando Then
        mCancelado = True
    Else
        Unload Me
    End If
End Sub


Private Sub cmdCargar_Click()
    Dim v As Variant, rs As Recordset, sql As String, s As String, i As Integer, j As Integer
        If Len(txtNumDocRef.Text) > 0 Then
            If Me.Caption = "Busca Transacciones para Cambio de Acopiador" Then
                Set rs = gobjMain.EmpresaActual.ConsIVCambioAcopiadorISO(txtNumDocRef.Text)
                If Not rs.EOF Then
                    grd.Cols = 17
                    rs.MoveFirst
                    For i = 1 To rs.RecordCount
                        s = ""
                        s = "" & vbTab
                        For j = 0 To 15
                            s = s & rs.Fields(j) & vbTab
                        Next j
                        grd.AddItem s, i
                        grd.Redraw = flexRDDirect
                        rs.MoveNext
                    Next i
    
                    If Me.Caption = "Busca Transacciones para Cambio de Acopiador" Then
                            ConfigColsCambiaAcopiador
                    Else
                            ConfigColsImprimeEtiketas
                    End If
                    grd.Redraw = flexRDDirect
                    cmdImprimiCH.Enabled = True
                    cmdImprimir.Enabled = True
                    cmdImprimir.SetFocus
                    txtNumDocRef.Text = ""
                    txtNumDocRef.SetFocus
            
                Else
                    MsgBox "Orden " & txtNumDocRef.Text & " NO existe"
                    txtNumDocRef.Text = ""
                    txtNumDocRef.SetFocus
                    
                    ConfigColsImprimeEtiketas
                    grd.Redraw = flexRDDirect
                End If
                
                
            Else
                Set rs = gobjMain.EmpresaActual.ConsIVTransImpresionEtiquetasISO(txtNumDocRef.Text)
                If Not rs.EOF Then
                    grd.Cols = COL_E_RESULTADO
                    rs.MoveFirst
                    For i = 1 To rs.RecordCount
                        s = ""
                        s = "" & vbTab
                        For j = 0 To COL_E_GAR - 1
                            s = s & rs.Fields(j) & vbTab
                        Next j
                        grd.AddItem s, i
                        grd.Redraw = flexRDDirect
                        rs.MoveNext
                    Next i
    
                    If Me.Caption = "Busca Transacciones para Cambio de Acopiador" Then
                            ConfigColsCambiaAcopiador
                    Else
                            ConfigColsImprimeEtiketas
                    End If
                    grd.Redraw = flexRDDirect
                    cmdImprimiCH.Enabled = True
                    cmdImprimir.Enabled = True
                    cmdImprimir.SetFocus
                    txtNumDocRef.Text = ""
                    txtNumDocRef.SetFocus
            
                Else
                    MsgBox "Orden " & txtNumDocRef.Text & " NO existe"
                    txtNumDocRef.Text = ""
                    txtNumDocRef.SetFocus
                    
                    ConfigColsImprimeEtiketas
                    grd.Redraw = flexRDDirect
                End If
            End If
            
        End If
    
End Sub

Private Sub cmdCargarTrans_Click()
    Dim v As Variant, rs As Recordset, sql As String, s As String, i As Integer, j As Integer
        For i = 2 To grd.Rows - 1
            grd.RemoveItem 1
            grd.Redraw = flexRDNone
        Next i
        If Me.Caption = "Busca Transacciones de Produccion para Impresion de Etiquetas" Then
            If Len(fcbTrans.Text) > 0 Then
                s = fcbTrans.KeyText
                gobjMain.EmpresaActual.GNOpcion.AsignarValor "ImpLoteProd_Codtrans", s
                
                Set rs = gobjMain.EmpresaActual.ConsIVTransImpresionEtiquetasISOxProduccionFecha(fcbTrans.KeyText, Val(txtNumTrans1.Text), dtpFecha1.value, dtpFecha2.value, cboTipoTrabajo.ListIndex)
                If Not rs.EOF Then
                    grd.Cols = COL_E_RESULTADO + 1
                    rs.MoveFirst
                    For i = 1 To rs.RecordCount
                        s = ""
                        s = "" & vbTab
                        For j = 0 To COL_E_GAR
                            s = s & rs.Fields(j) & vbTab
                        Next j
                        grd.AddItem s, i
                        grd.Redraw = flexRDDirect
                        rs.MoveNext
                    Next i
    
                    ConfigColsImprimeEtiketasProduccion
                    grd.Redraw = flexRDDirect
                    cmdImprimiCH.Enabled = True
                    cmdImprimir.Enabled = True
                    cmdImprimir.SetFocus
                    txtNumTrans1.Text = ""
                    txtNumTrans1.SetFocus
            
                Else
                    MsgBox "Orden " & txtNumDocRef.Text & " NO existe"
                    txtNumTrans1.Text = ""
                    txtNumTrans1.SetFocus
                    
                    ConfigColsImprimeEtiketas
                    grd.Redraw = flexRDDirect
                End If
            End If
        ElseIf Me.Caption = "Busca Transacciones de Facturacion para Impresion de Etiquetas" Then
            If Len(fcbTrans.Text) > 0 Then
                s = fcbTrans.KeyText
                gobjMain.EmpresaActual.GNOpcion.AsignarValor "ImpLoteFact_Codtrans", s
                
                Set rs = gobjMain.EmpresaActual.ConsIVTransImpresionEtiquetasISOxFacturacionFecha(fcbTrans.KeyText, Val(txtNumTrans1.Text), dtpFecha1.value, dtpFecha2.value, cboTipoTrabajo.ListIndex)
                If Not rs.EOF Then
                    grd.Cols = COL_E_RESULTADO + 1
                    rs.MoveFirst
                    For i = 1 To rs.RecordCount
                        s = ""
                        s = "" & vbTab
                        For j = 0 To COL_E_GAR
                            s = s & rs.Fields(j) & vbTab
                        Next j
                        grd.AddItem s, i
                        grd.Redraw = flexRDDirect
                        rs.MoveNext
                    Next i
    
                    ConfigColsImprimeEtiketasProduccion
                    grd.Redraw = flexRDDirect
                    cmdImprimiCH.Enabled = True
                    cmdImprimir.Enabled = True
                    cmdImprimir.SetFocus
                    txtNumTrans1.Text = ""
                    txtNumTrans1.SetFocus
            
                Else
                    MsgBox "Orden " & txtNumDocRef.Text & " NO existe"
                    txtNumTrans1.Text = ""
                    txtNumTrans1.SetFocus
                    
                    ConfigColsImprimeEtiketas
                    grd.Redraw = flexRDDirect
                End If
            End If
        End If
            
End Sub



Private Sub cmdExaminarCarpetaRuta_Click()
    On Error GoTo errtrap
    slf.OwnerHWnd = Me.hWnd
    slf.Path = txtCarpeta.Text
    If slf.Browse Then
        txtCarpeta.Text = slf.Path
        txtCarpeta_LostFocus
    End If
    Exit Sub
errtrap:
    MsgBox Err.Description, vbInformation
    Exit Sub
End Sub

Private Sub cmdImprimir_Click()
    Dim i As Integer, j As Integer, v As Variant
    Dim Filas  As Integer, s As String
    Dim sql As String, NumReg As Long, resp As Integer
    Dim nombre As String, file As String, tid As Long, res As String, cntError As Long, filepdf As String, nombrepdf As String
    Dim fileA As String, fileApdf As String
    Dim gnc As GNComprobante
    'Si no hay transacciones
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna transacción para imprimir."
        Exit Sub
    End If
    If Me.Caption = "Busca Transacciones para Impresion de Etiquetas" Then
        Filas = 0
        ReDim v(COL_E_GAR, 1)
            For i = 1 To grd.Rows - 1
                If Not grd.IsSubtotal(i) Then
                    ReDim Preserve v(COL_E_GAR, Filas)
                    For j = 1 To COL_E_GAR
                        v(j - 1, Filas) = grd.TextMatrix(i, j)
                    Next j
                        Filas = Filas + 1
                End If
            Next i
        
        s = ntxMargIzq.value
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "ImpLote_MarIzq", s
            
        s = ntxMargSup.value
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "ImpLote_MarSup", s
            
        s = ntxNUmEtiq.value
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "ImpLote_NumEtiq", s
            
        'Graba en la base
        gobjMain.EmpresaActual.GNOpcion.Grabar
            
        FrmImprimeEtiketas.InicioNew v
    ElseIf Me.Caption = "Busca Transacciones de Produccion para Impresion de Etiquetas " Then
        Filas = 0
        ReDim v(COL_E_GAR, 1)
            For i = 1 To grd.Rows - 1
                If Not grd.IsSubtotal(i) Then
                    ReDim Preserve v(COL_E_GAR, Filas)
                    For j = 1 To COL_E_GAR
                        v(j - 1, Filas) = grd.TextMatrix(i, j)
                    Next j
                        Filas = Filas + 1
                End If
            Next i
        
        s = ntxMargIzq.value
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "ImpLoteProd_MarIzq", s
            
        s = ntxMargSup.value
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "ImpLoteProd_MarSup", s
            
        s = ntxNUmEtiq.value
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "ImpLoteProd_NumEtiq", s
            
        'Graba en la base
        gobjMain.EmpresaActual.GNOpcion.Grabar
            
        FrmImprimeEtiketas.InicioProduccion v
    ElseIf Me.Caption = "Busca Transacciones de Facturacion para Impresion de Etiquetas" Then
        Filas = 0
        ReDim v(COL_E_GAR, 1)
            For i = 1 To grd.Rows - 1
                If Not grd.IsSubtotal(i) Then
                    ReDim Preserve v(COL_E_GAR, Filas)
                    For j = 1 To COL_E_GAR
                        v(j - 1, Filas) = grd.TextMatrix(i, j)
                    Next j
                        Filas = Filas + 1
                End If
            Next i
        
        s = ntxMargIzq.value
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "ImpLoteFact_MarIzq", s
            
        s = ntxMargSup.value
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "ImpLoteFact_MarSup", s
            
        s = ntxNUmEtiq.value
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "ImpLoteFact_NumEtiq", s
            
        'Graba en la base
        gobjMain.EmpresaActual.GNOpcion.Grabar
            
        FrmImprimeEtiketas.InicioProduccion v
 
    ElseIf Me.Caption = "Busca Transacciones con problemas de Clientes/Proveedores en Pagos/Cobros" Then
        For i = 1 To grd.Rows - 1
            sql = " Update Pckardex"
            sql = sql & " set "
            sql = sql & " IdProvcli = " & grd.ValueMatrix(i, 3)
            sql = sql & " where id =" & grd.ValueMatrix(i, 9)
                                
            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
            
           
            
            
            grd.ShowCell i, 10
            grd.TextMatrix(i, 10) = "O.K."
            
        Next i
    ElseIf Me.Caption = "Impresion Relacion Dependencia" Then 'AUC 2012
        'For i = grd.FixedRows To grd.Rows - 1
          '  If grd.ValueMatrix(i, 4) <> 0 Then
                If ImprimirRelDep Then
                    cmdCancelar.SetFocus
               End If
            'End If
        'Next
    ElseIf Me.Caption = "Busca Transacciones para Cambio de Acopiador" Then
        If Len(fcbAcopiador.KeyText) > 0 Then
            If MsgBox("Desea cambiar los tikets cargados al Acopiador: " & fcbAcopiador.KeyText, vbYesNo) = vbYes Then
                For i = 1 To grd.Rows - 1
                    If Not grd.IsSubtotal(i) Then
                        sql = " Update IVInventarioDetalleISO"
                        sql = sql & " set "
                        sql = sql & " IdObra = (select idobra from gnobra where codobra='" & fcbAcopiador.KeyText & "')"
                        sql = sql & " where id =" & grd.ValueMatrix(i, 3)
                                            
                        gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                        
                        sql = " Update gncomprobante "
                        sql = sql & " set "
                        sql = sql & " IdObra = (select idobra from gnobra where codobra='" & fcbAcopiador.KeyText & "')"
                        sql = sql & " where transid =" & grd.ValueMatrix(i, 16)
                        
                       gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                        
                        
                        grd.ShowCell i, 16
                        grd.TextMatrix(i, 15) = fcbAcopiador.KeyText
                        grd.TextMatrix(i, 16) = "O.K."
                    End If
                    
                Next i
            End If
        Else
            MsgBox "No está seleccionado ningun acopiador"
            fcbAcopiador.SetFocus
        End If
    ElseIf Me.Caption = "Generar Archivo RIDE en formato pdf" Then 'AUC 2012
         If GeneraPDF Then
             cmdCancelar.SetFocus
        End If
    ElseIf Me.Caption = "Generar Archivo xml" Then 'AUC 2012
         If GeneraXML Then
             cmdCancelar.SetFocus
        End If
        
    ElseIf Me.Caption = "Recupera Archivo xml desde Base Datos" Then

        For i = 1 To grd.Rows - 1
                If Not grd.IsSubtotal(i) Then
                    
                    nombre = Mid$(grd.TextMatrix(i, 8), 1, 39) & Right("0000000000" & grd.TextMatrix(i, COL_TID), 10) & ".xml"
                    nombrepdf = Mid$(grd.TextMatrix(i, 8), 1, 39) & Right("0000000000" & grd.TextMatrix(i, COL_TID), 10) & ".pdf"
                    
                    file = txtCarpeta.Text & nombre
                    filepdf = txtCarpeta.Text & nombrepdf
                    
                    fileA = gobjMain.EmpresaActual.GNOpcion.ComprobantesAutorizados & "\" & nombre
                    fileApdf = gobjMain.EmpresaActual.GNOpcion.ComprobantesAutorizados & "\" & nombrepdf
                    
                    
                    If ExisteArchivo(file) And ExisteArchivo(filepdf) Then
                        grd.TextMatrix(i, 16) = "O.K."
                    Else
                    
                    If ExisteArchivo(fileA) And ExisteArchivo(fileApdf) Then
                        grd.TextMatrix(i, 16) = "O.K. en Autorizados"
                    Else
                        file = gobjMain.EmpresaActual.GNOpcion.ComprobantesAutorizados & "\" & nombre
                        NumFile = FreeFile
                        Open file For Output Access Write As #NumFile
                        
                        
                        Print #NumFile, grd.TextMatrix(i, 9)
                        Close NumFile
                        

                        grd.TextMatrix(i, 16) = "Generado xml"
                        grd.Refresh
                        
                    If optAsiento.value Then
                        tid = grd.ValueMatrix(i, COL_TID)
                        
                        'Recupera la transaccion
                        Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
                        If Not (gnc Is Nothing) Then
                            'Si la transacción no está anulado
                            If gnc.Estado <> ESTADO_ANULADO And gnc.CodigoMensaje = "60" Then
                                'Imprime la transaccion o asiento contable

                                    res = ImprimeTransRide(gnc, mobjImp)

                                If res Then
                                    If grd.TextMatrix(i, COL_C_RESULTADO + 6) = "Generado xml" Then
                                        grd.TextMatrix(i, COL_C_RESULTADO + 6) = "Generado xml + pdf"
                                    Else
                                        grd.TextMatrix(i, COL_C_RESULTADO + 6) = "Generado pdf"
                                    End If
                                Else
                                    grd.TextMatrix(i, COL_C_RESULTADO + 6) = "Error"
                                    cntError = cntError + 1
                                End If
                                            
                            'Si la transaccion está anulado
                            Else
                                grd.TextMatrix(i, COL_C_RESULTADO) = "Anulado."
                                cntError = cntError + 1
                            End If
                        Else
                            grd.TextMatrix(i, COL_C_RESULTADO) = "No pudo recuperar la transación."
                            cntError = cntError + 1
                        End If
                    Else
                        grd.TextMatrix(i, COL_C_RESULTADO + 6) = "Generado xml "
                    End If
                    End If
                    End If
                        
                        
                        
                        
                        grd.ShowCell i, 16
                        End If

                
        Next i
        
'        GeneraPDF
    
    
    Else
        If Imprimir Then
            cmdCancelar.SetFocus
        End If
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF9
        cmdImprimir_Click
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
    grd.Move 0, grd.Top, Me.ScaleWidth, Me.ScaleHeight - grd.Top - pic1.Height - 80
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


Public Function ImprimeTrans(ByVal gc As GNComprobante, ByVal bandAsiento As Boolean) As String
    Dim crear As Boolean
    Static objImp As Object
    On Error GoTo errtrap

    'Si no tiene TransID quiere decir que no está grabada
    If (gc.TransID = 0) Or gc.Modificado Then
        MsgBox MSGERR_NOGRABADO
        ImprimeTrans = False
        Exit Function
    End If
    
    'Solo por primera vez o cuando cambia la librería de impresión
    '  crea una instancia del objeto para la impresión
    crear = (objImp Is Nothing)
    If Not crear Then crear = (objImp.NombreDLL <> gc.GNTrans.ArchivoReporte)
    If crear Then
        Set objImp = Nothing
        Set objImp = CreateObject(gc.GNTrans.ArchivoReporte & ".PrintTrans")
    End If
    
    MensajeStatus "Está imprimiéndo ...", vbHourglass
    
    If Not bandAsiento Then
        'Envia directamente a la impresora con el segundo parámetro 'True'
        objImp.PrintTrans gobjMain.EmpresaActual, True, 1, 0, "", 0, gc
    Else
        objImp.PrintAsiento gobjMain.EmpresaActual, True, 1, 0, "", 0, gc
    End If
    MensajeStatus "", 0
    ImprimeTrans = ""       'Sin problema
    Exit Function
errtrap:
    MensajeStatus "", 0
    Select Case Err.Number
    Case ERR_NOIMPRIME, ERR_NOIMPRIME2, ERR_NOIMPRIME3, ERR_NOHAYCODIGO
        ImprimeTrans = Err.Description
    Case Else
        ImprimeTrans = MSGERR_NOIMPRIME2
    End Select
    Exit Function
End Function

Public Sub InicioPerdidaIdAsignado()
    Dim i As Integer
    On Error GoTo errtrap
    optAsiento.Visible = False
    optTrans.Visible = False
    fraCodTrans.Visible = False
    fraNumTrans.Visible = False
    Me.Caption = "Busca Transacciones con problemas de Relación"
    Me.Show
    Me.ZOrder
    dtpFecha1.value = gobjMain.EmpresaActual.GNOpcion.FechaInicio
    dtpFecha2.value = Date
    CargaTrans
    Exit Sub
errtrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub ConfigColsTRansErradas()
    With grd
        .FormatString = "^#|<Fecha|<Trans|<#|<Descripción"
        
        .ColDataType(1) = flexDTDate    '*** MAKOTO 14/ago/2000 para que ordene bien por fecha
        
        GNPoneNumFila grd, False
        '.AutoSize 0, grd.Cols - 1
        .ColWidth(1) = 1000
        .ColWidth(2) = 900
        .ColWidth(3) = 900
        .ColWidth(4) = 6000
    End With
End Sub


Public Sub InicioCheque()
    Dim i As Integer, s As String
    On Error GoTo errtrap
    Me.Caption = "Busca Transacciones para Impresion de Cheques"
    Me.Show
    Me.ZOrder
    dtpFecha1.value = Date
    dtpFecha2.value = Date
    CargaBanco
    optTrans.Visible = False
    optAsiento.Visible = False
    fraCodTrans.Caption = "Banco"
    fraNumTrans.Caption = "# Chque (desde - hasta)"
    FraConFigEgreso.Visible = True
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_LibImpPago")) > 0 Then
        s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_LibImpPago")
        txtEgreso.Text = s
    End If

    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_LibImpCheque")) > 0 Then
        s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_LibImpCheque")
        txtCheque.Text = s
    End If
    Exit Sub
errtrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub CargaBanco()
    'Carga la lista de transacción
    fcbTrans.SetData gobjMain.EmpresaActual.ListaTSBanco(True, False)
    fcbTrans.DispCol = 1
End Sub

Private Sub ConfigColsImprimeCheques()
    With grd
        .FormatString = "^#|tid|<Fecha|<Trans|<Descripción|<Beneficiario|<Banco|># Cheque|>Valor|^Estado|<Resultado"
        .ColHidden(COL_TID) = True
        .ColHidden(COL_C_ESTADO) = True
        
        .ColDataType(COL_FECHA) = flexDTDate
        .ColDataType(COL_C_VALOR) = flexDTDouble
        .ColFormat(COL_C_VALOR) = "##,0.00"
        
        GNPoneNumFila grd, False
        .AutoSize 0, grd.Cols - 1
        
        .ColWidth(COL_C_TRANS) = 800
        .ColWidth(COL_C_DESC) = 3400
        .ColWidth(COL_C_BENEFICIARIO) = 3400
        .ColWidth(COL_C_RESULTADO) = 2000
    End With
End Sub


Private Sub cmdExplorar_Click()
    On Error GoTo errtrap
    
    With dlg1
        If Len(.filename) = 0 Then
            .InitDir = txtEgreso.Text
            '.FileName = mPlantilla.BDDestino
        Else
            .InitDir = .filename
            '.FileName = mPlantilla.BDDestino
        End If
        .flags = cdlOFNPathMustExist
        .Filter = "Base de datos Jet (*.txt)|*.txt|Predefinido *.txt |Todos (*.*)|*.*"
        .ShowSave
        txtEgreso.Text = .filename
    End With
    
    Exit Sub
errtrap:
    If Err.Number <> 32755 Then
        DispErr
    End If
    Exit Sub
End Sub

Private Sub cmdExplorarCH_Click()
On Error GoTo errtrap
    
    With dlg1
        If Len(.filename) = 0 Then
            .InitDir = txtCheque.Text
            '.FileName = mPlantilla.BDDestino
        Else
            .InitDir = .filename
            '.FileName = mPlantilla.BDDestino
        End If
        .flags = cdlOFNPathMustExist
        .Filter = "Base de datos Jet (*.txt)|*.txt|Predefinido *.txt |Todos (*.*)|*.*"
        .ShowSave
        txtCheque.Text = .filename
    End With
    
    Exit Sub
errtrap:
    If Err.Number <> 32755 Then
        DispErr
    End If
    Exit Sub
End Sub

Private Sub cmdImprimiCH_Click()
    'Si no hay transacciones
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna transacción para imprimir."
        Exit Sub
    End If
    
    If ImprimirCheque Then
        cmdCancelar.SetFocus
    End If
End Sub



Private Function ImprimirCheque() As Boolean
    Dim s As String, tid As Long, i As Long, x As Single, res As String, pos As Integer
    Dim gnc As GNComprobante, cambiado As Boolean, cntError As Long
    
    On Error GoTo errtrap

    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdBuscar.Enabled = False
    cmdImprimiCH.Enabled = False
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    
    For i = grd.FixedRows To grd.Rows - 1
        DoEvents
        If mCancelado Then
            MsgBox "El proceso fue cancelado."
            Exit For
        End If
        
        prg1.value = i
        grd.Row = i
        x = grd.CellTop                 'Para visualizar la celda actual
        pos = Len(grd.TextMatrix(i, COL_C_RESULTADO))
        'Si es verificación, procesa todas las filas sino solo las que tengan "Asiento incorrecto."
        If pos = 0 Then
        
            tid = grd.ValueMatrix(i, COL_TID)
            grd.TextMatrix(i, COL_C_RESULTADO) = "Procesando ..."
            grd.Refresh
            
            'Recupera la transaccion
            Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
            If Not (gnc Is Nothing) Then
                'Si la transacción no está anulado
                If gnc.Estado <> ESTADO_ANULADO Then
    '                'Forzar recuperar todos los datos de transacción
    '                ' para que no se pierdan al grabar de nuveo
    '                gnc.RecuperaDetalleTodo
                
                    'Imprime la transaccion o asiento contable
                    res = ImprimeCheque(gnc, False)
                    If Len(res) = 0 Then
                        grd.TextMatrix(i, COL_C_RESULTADO) = "Enviado."
                    Else
                        grd.TextMatrix(i, COL_C_RESULTADO) = res
                        cntError = cntError + 1
                    End If
                                
                'Si la transaccion está anulado
                Else
                    grd.TextMatrix(i, COL_C_RESULTADO) = "Anulado."
                    cntError = cntError + 1
                End If
            Else
                grd.TextMatrix(i, COL_C_RESULTADO) = "No pudo recuperar la transación."
                cntError = cntError + 1
            End If
        End If
    Next i
    
    Screen.MousePointer = 0
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdImprimir.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    
    'Si algúna transaccion no se imprimió, avisa
    If cntError Then
        MsgBox "No se pudo imprimir " & cntError & " transacciones.", vbInformation
    End If
    
    ImprimirCheque = True
    Exit Function
errtrap:
    Screen.MousePointer = 0
    DispErr
    prg1.value = prg1.min
    Exit Function
End Function

Public Function ImprimeCheque(ByVal gc As GNComprobante, ByVal bandAsiento As Boolean) As String
    Dim crear As Boolean
    Static objImp As Object
    On Error GoTo errtrap

    'Si no tiene TransID quiere decir que no está grabada
    If (gc.TransID = 0) Or gc.Modificado Then
        MsgBox MSGERR_NOGRABADO
        ImprimeCheque = False
        Exit Function
    End If
    
    'Solo por primera vez o cuando cambia la librería de impresión
    '  crea una instancia del objeto para la impresión
    crear = (objImp Is Nothing)
    If Not crear Then crear = (objImp.NombreDLL <> gc.GNTrans.ArchivoReporte)
    If crear Then
        Set objImp = Nothing
        Set objImp = CreateObject(gc.GNTrans.ArchivoReporte & ".PrintTrans")
    End If
    
    MensajeStatus "Está imprimiéndo ...", vbHourglass
    objImp.PrintTransLoteRuta gobjMain.EmpresaActual, True, 1, 0, txtCheque.Text, "", 0, gc
        
    MensajeStatus "", 0
    ImprimeCheque = ""       'Sin problema
    Exit Function
errtrap:
    MensajeStatus "", 0
    Select Case Err.Number
    Case ERR_NOIMPRIME, ERR_NOIMPRIME2, ERR_NOIMPRIME3, ERR_NOHAYCODIGO
        ImprimeCheque = Err.Description
    Case Else
        ImprimeCheque = MSGERR_NOIMPRIME2
    End Select
    Exit Function
End Function


Public Function ImprimeTransLote(ByVal gc As GNComprobante, ByVal bandAsiento As Boolean) As String
    Dim crear As Boolean
    Static objImp As Object
    On Error GoTo errtrap
    
    'Si no tiene TransID quiere decir que no está grabada
    If (gc.TransID = 0) Or gc.Modificado Then
        MsgBox MSGERR_NOGRABADO
        ImprimeTransLote = False
        Exit Function
    End If
    
    'Solo por primera vez o cuando cambia la librería de impresión
    '  crea una instancia del objeto para la impresión
    crear = (objImp Is Nothing)
    If Not crear Then crear = (objImp.NombreDLL <> gc.GNTrans.ArchivoReporte)
    If crear Then
        Set objImp = Nothing
        Set objImp = CreateObject(gc.GNTrans.ArchivoReporte & ".PrintTrans")
    End If
    
    MensajeStatus "Está imprimiéndo ...", vbHourglass
    objImp.PrintTransLoteRuta gobjMain.EmpresaActual, True, 1, 0, txtEgreso.Text, "", 0, gc
    
    MensajeStatus "", 0
    ImprimeTransLote = ""       'Sin problema
    Exit Function
errtrap:
    MensajeStatus "", 0
    Select Case Err.Number
    Case ERR_NOIMPRIME, ERR_NOIMPRIME2, ERR_NOIMPRIME3, ERR_NOHAYCODIGO
        ImprimeTransLote = Err.Description
    Case Else
        ImprimeTransLote = MSGERR_NOIMPRIME2
    End Select
    Exit Function
End Function

Public Sub InicioEtiqueta()
    Dim i As Integer, s As String
    On Error GoTo errtrap
    Me.Caption = "Busca Transacciones para Impresion de Etiquetas"
    Me.Show
    Me.ZOrder
    framConfig.Visible = True
    dtpFecha1.value = Date
    dtpFecha2.value = Date
    CargaTrans
    fraFecha.Visible = False
    fraNumTrans.Visible = False
    fraCodTrans.Visible = False
    txtNumTrans2.Visible = False
    optTrans.Visible = False
    optAsiento.Visible = False
    fraCodTrans.Caption = "Trans"
    fraNumTrans.Caption = "# Trans "
    cmdCargarTrans.Visible = True
    
    FraConFigEgreso.Visible = False
    FraTipoTrabajo.Visible = False
    FraBuscaOrden.Visible = True
    cboTipoTrabajo.ListIndex = 0
    cmdBuscar.Enabled = False
    cmdImprimiCH.Visible = False
    
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_CodTrans")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_CodTrans")
            fcbTrans.KeyText = s
        End If
    
    
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_MarIzq")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_MarIzq")
            ntxMargIzq.value = s
        End If
        
        
        
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_MarSup")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_MarSup")
            ntxMargSup.value = s
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_NumEtiq")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_NumEtiq")
            ntxNUmEtiq.value = s
        End If
    txtNumDocRef.SetFocus
    Exit Sub
errtrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub ConfigColsImprimeEtiketas()
    Dim s As String
    With grd
        
        s = "^#|># Ingreso|<Orden|<Tiket|<Marca|<Tamaño|<Serie|<Diseño|<Trabajo"
        s = s & "|<Trans Ingreso|<Fecha Ingreso|<Cod Cli|<Nombre|<Vendedor|<Reciclador|<Resultado"
        
        
        .FormatString = s
        
        .ColDataType(COL_E_FECHA) = flexDTDate
        
            grd.subtotal flexSTSum, COL_E_NUMING, COL_E_NUMING, , grd.BackColorFixed, , True, " ", 1, False     '5

        
        GNPoneNumFila grd, False
        .AutoSize 0, grd.Cols - 1
        
        .ColWidth(COL_E_RESULTADO) = 2000
        
        
        
    End With
End Sub

Public Sub InicioVerificaPagosErrados()
    Dim i As Integer
    On Error GoTo errtrap
    optAsiento.Visible = False
    optTrans.Visible = False
    fraCodTrans.Visible = False
    fraNumTrans.Visible = False
    Me.Caption = "Busca Transacciones con problemas de Clientes/Proveedores en Pagos/Cobros"
    Me.Show
    Me.ZOrder
    dtpFecha1.value = gobjMain.EmpresaActual.GNOpcion.FechaInicio
    dtpFecha2.value = Date
    CargaTrans
    cmdImprimir.Caption = "Corregir"
    'cmdImprimir.Visible = False
    cmdImprimiCH.Visible = False
    Exit Sub
errtrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub ConfigColsTransErradasCobroPagos()
    With grd
        .FormatString = "^#|<Fecha|<Trans|<ID|<Nombre|<Fecha Pago|<Trans Pago|<ID|<Nombre|>ID|<Resultado"
        
        .ColDataType(1) = flexDTDate    '*** MAKOTO 14/ago/2000 para que ordene bien por fecha
        
        GNPoneNumFila grd, False
        '.AutoSize 0, grd.Cols - 1
        .ColWidth(0) = 400
        .ColWidth(1) = 1000
        .ColWidth(2) = 1500
        .ColWidth(3) = 1000
        .ColWidth(4) = 4000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1500
        .ColWidth(7) = 1000
        .ColWidth(8) = 4000
        .ColWidth(9) = 1000
        
        .ColHidden(3) = True
        .ColHidden(7) = True
        .ColHidden(9) = True
        
        
    End With
End Sub

Public Sub InicioRelDep()
    Dim i As Integer
    On Error GoTo errtrap
    Me.Caption = "Impresion Relacion Dependencia"
    Me.Show
    Me.ZOrder
    dtpFecha1.value = gobjMain.EmpresaActual.GNOpcion.FechaInicio
    dtpFecha2.value = Date
    CargaTrans
    optTrans.Visible = False
    optAsiento.Visible = False
    grd.Editable = True
    grd.Enabled = True
    txtNumTrans2.Visible = False
    Exit Sub
errtrap:
    DispErr
    Unload Me
    Exit Sub
End Sub
Private Sub ConfigColsRelDep()
    With grd
        .FormatString = "^#|<idprovcli|<Codigo|<Nombre|^Imprimir"
        
        .ColDataType(4) = flexDTBoolean
        
        GNPoneNumFila grd, False
        '.AutoSize 0, grd.Cols - 1
        .ColWidth(1) = 0
        .ColWidth(2) = 2000
        .ColWidth(3) = 3500
        .ColWidth(4) = 900
        grd.ColData(2) = -1
        grd.ColData(3) = -1
    End With
End Sub
Private Function ImprimirRelDep() As Boolean
    Dim s As String, tid As Long, i As Long, x As Single, res As String
    Dim gnc As GNComprobante, cambiado As Boolean, cntError As Long
    
    On Error GoTo errtrap

    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdBuscar.Enabled = False
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    
    'Limpia los mensajes
    'For i = grd.FixedRows To grd.Rows - 1
      '  grd.TextMatrix(i, COL_C_RESULTADO) = ""
    'Next i
    Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(0, fcbTrans.KeyText, txtNumTrans1.Text)
    
    For i = grd.FixedRows To grd.Rows - 1
        If grd.ValueMatrix(i, 4) <> 0 Then
            DoEvents
            If mCancelado Then
                MsgBox "El proceso fue cancelado."
                Exit For
            End If
            
            prg1.value = i
            grd.Row = i
            x = grd.CellTop                 'Para visualizar la celda actual
        
'        tid = grd.ValueMatrix(i, COL_TID)
 '       grd.TextMatrix(i, COL_C_RESULTADO) = "Procesando ..."
            grd.Refresh
        
        'Recupera la transaccion
        
    '    If Not (gnc Is Nothing) Then
            'Si la transacción no está anulado
     '       If gnc.Estado <> ESTADO_ANULADO Then
'                'Forzar recuperar todos los datos de transacción
'                ' para que no se pierdan al grabar de nuveo
'                gnc.RecuperaDetalleTodo
            
                'Imprime la transaccion o asiento contable
                'If Me.Caption = "Busca Transacciones para Impresion de Cheques" Then
                 '   res = ImprimeTransLote(gnc, optAsiento.value)
                'Else
                    res = ImprimeRegRelDeP(gnc, grd.TextMatrix(i, 2))
               ' End If
'                If Len(res) = 0 Then
'                    grd.TextMatrix(i, COL_C_RESULTADO) = "Enviado."
'                Else
'                    grd.TextMatrix(i, COL_C_RESULTADO) = res
'                    cntError = cntError + 1
'                End If
                            
            'Si la transaccion está anulado
      '      Else
               ' grd.TextMatrix(i, COL_C_RESULTADO) = "Anulado."
       '         cntError = cntError + 1
        '    End If
       ' Else
'            grd.TextMatrix(i, COL_C_RESULTADO) = "No pudo recuperar la transación."
        '    cntError = cntError + 1
        'End If
        End If
    Next i
    
    Screen.MousePointer = 0
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdImprimir.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    
    'Si algúna transaccion no se imprimió, avisa
    If cntError Then
        MsgBox "No se pudo imprimir " & cntError & " transacciones.", vbInformation
    End If
    
    ImprimirRelDep = True
    Exit Function
errtrap:
    Screen.MousePointer = 0
    DispErr
    prg1.value = prg1.min
    Exit Function
End Function
Public Function ImprimeRegRelDeP(ByVal gc As GNComprobante, ByVal CodEmpleado As String) As String
    Dim crear As Boolean
    Static objImp As Object
    On Error GoTo errtrap

    'Si no tiene TransID quiere decir que no está grabada
    If (gc.TransID = 0) Or gc.Modificado Then
        MsgBox MSGERR_NOGRABADO
        ImprimeRegRelDeP = False
        Exit Function
    End If
    
    'Solo por primera vez o cuando cambia la librería de impresión
    '  crea una instancia del objeto para la impresión
    crear = (objImp Is Nothing)
    If Not crear Then crear = (objImp.NombreDLL <> gc.GNTrans.ArchivoReporte)
    If crear Then
        Set objImp = Nothing
        Set objImp = CreateObject(gc.GNTrans.ArchivoReporte & ".PrintTrans")
    End If
    
    MensajeStatus "Está imprimiéndo ...", vbHourglass
    'Public Sub PrintComprobanteRet(Emp As Recordset, ByRef objComp As gnComprobante)
    objImp.PrintComprobanteRet gc, CodEmpleado
    
    MensajeStatus "", 0
    ImprimeRegRelDeP = ""       'Sin problema
    Exit Function
errtrap:
    MensajeStatus "", 0
    Select Case Err.Number
    Case ERR_NOIMPRIME, ERR_NOIMPRIME2, ERR_NOIMPRIME3, ERR_NOHAYCODIGO
        ImprimeRegRelDeP = Err.Description
    Case Else
        ImprimeRegRelDeP = MSGERR_NOIMPRIME2
    End Select
    Exit Function
End Function

Public Sub InicioAcopiador()
    Dim i As Integer, s As String
    On Error GoTo errtrap
    Me.Caption = "Busca Transacciones para Cambio de Acopiador"
    Me.Show
    Me.ZOrder
    framConfig.Visible = True
    dtpFecha1.value = Date
    dtpFecha2.value = Date
    CargaTrans
    ConfigColsCambiaAcopiador
    fraFecha.Visible = False
    fraNumTrans.Visible = False
    fraCodTrans.Visible = False
    optTrans.Visible = False
    optAsiento.Visible = False
    fraCodTrans.Caption = "Trans"
    fraNumTrans.Caption = "# Orden (desde - hasta)"
    FraConFigEgreso.Visible = False
    FraTipoTrabajo.Visible = False
    FraBuscaOrden.Visible = True
    cboTipoTrabajo.ListIndex = 0
    cmdBuscar.Enabled = False
    cmdImprimiCH.Visible = False
    framConfig.Visible = False
    FraAcopiador.Visible = True
    FraAcopiador.Top = FraBuscaOrden.Top
    cmdImprimir.Caption = "Actualizar"
'        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_MarIzq")) > 0 Then
'            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_MarIzq")
'            ntxMargIzq.value = s
'        End If
'
'        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_MarSup")) > 0 Then
'            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_MarSup")
'            ntxMargSup.value = s
'        End If
'
'        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_NumEtiq")) > 0 Then
'            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLote_NumEtiq")
'            ntxNUmEtiq.value = s
'        End If
    txtNumDocRef.SetFocus
    Exit Sub
errtrap:
    DispErr
    Unload Me
    Exit Sub
End Sub


Private Sub ConfigColsCambiaAcopiador()
    Dim s As String
    With grd
        
        s = "^#|># Ingreso|<Orden|<Tiket|<Marca|<Tamaño|<Serie|<Diseño|<Trabajo"
        s = s & "|<Trans Ingreso|<Fecha Ingreso|<Cod Cli|<Nombre|<Vendedor|<Reciclador|<Acopiador|<transid|<Resultado"
        
        
        .FormatString = s
        
        .ColDataType(COL_E_FECHA) = flexDTDate
        
        grd.subtotal flexSTSum, COL_E_NUMING, COL_E_NUMING, , grd.BackColorFixed, , True, " ", 1, False     '5
        
        GNPoneNumFila grd, False
        .AutoSize 0, grd.Cols - 1
        
        .ColWidth(COL_E_RESULTADO) = 2000
        
        
        
    End With
End Sub


Public Sub InicioEtiquetaProduccion()
    Dim i As Integer, s As String
    On Error GoTo errtrap
    Me.Caption = "Busca Transacciones de Produccion para Impresion de Etiquetas"
    Me.Show
    Me.ZOrder
    FraBuscaOrden.Visible = False
    framConfig.Visible = True
    dtpFecha1.value = Date
    dtpFecha2.value = Date
    CargaTrans
    fraFecha.Visible = True
    fraNumTrans.Visible = True
    fraCodTrans.Visible = True
    txtNumTrans2.Visible = False
    optTrans.Visible = False
    optAsiento.Visible = False
    fraCodTrans.Caption = "Trans"
    fraNumTrans.Caption = "# Trans "
    cmdCargarTrans.Visible = True
    
    FraConFigEgreso.Visible = False
    FraTipoTrabajo.Visible = True
    FraBuscaOrden.Visible = False
    cboTipoTrabajo.ListIndex = 0
    cmdBuscar.Enabled = False
    cmdImprimiCH.Visible = False
    
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLoteProd_CodTrans")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLoteProd_CodTrans")
            fcbTrans.KeyText = s
        End If
    
    
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLoteProd_MarIzq")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLoteProd_MarIzq")
            ntxMargIzq.value = s
        End If
        
        
        
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLoteProd_MarSup")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLoteProd_MarSup")
            ntxMargSup.value = s
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLoteProd_NumEtiq")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLoteProd_NumEtiq")
            ntxNUmEtiq.value = s
        End If
        txtNumTrans1.SetFocus
    Exit Sub
errtrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub ConfigColsImprimeEtiketasProduccion()
    Dim s As String
    With grd
        
        s = "^#|># Ingreso|<Orden|<Tiket|<Marca|<Tamaño|<Serie|<Diseño|<Trabajo"
        s = s & "|<Trans Ingreso|<Fecha Ingreso|<Cod Cli|<Nombre|<Vendedor|<Reciclador|^Motivo|<Resultado"
        
        
        .FormatString = s
        
        .ColDataType(COL_E_FECHA) = flexDTDate
        
'        If Len(fcbTrans.Text) > 0 Then
''            grd.SubTotal flexSTSum, COL_E_TAMANIO, COL_E_TAMANIO, , grd.BackColorFixed ', , True, " ", 0, True
'        Else
'            grd.SubTotal flexSTSum, COL_E_NUMING, COL_E_NUMING, , grd.BackColorFixed, , True, " ", 1, False     '5
''        End If
        
        GNPoneNumFila grd, False
        .AutoSize 0, grd.Cols - 1
        
        .ColWidth(COL_E_RESULTADO + 1) = 2000
        
        
        
    End With
End Sub

Public Sub InicioPDF()
    Dim i As Integer
    On Error GoTo errtrap
    Me.Caption = "Generar Archivo RIDE en formato pdf"
    cmdImprimir.Caption = "Genera PDF"
    Me.Show
    Me.ZOrder
    dtpFecha1.value = gobjMain.EmpresaActual.GNOpcion.FechaInicio
    dtpFecha2.value = Date
    optTrans.Caption = "Para Eviar"
    optAsiento.Caption = "Para Portal"
    
    optTrans.Visible = False
    optAsiento.Visible = False
    
    CargaTrans
    CargaTransElectronica
    FraTransElect.Visible = True
    Exit Sub
errtrap:
    DispErr
    Unload Me
    Exit Sub
End Sub


Private Function GeneraPDF() As Boolean
    Dim s As String, tid As Long, i As Long, x As Single, res As String
    Dim gnc As GNComprobante, cambiado As Boolean, cntError As Long
    
    On Error GoTo errtrap

    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdBuscar.Enabled = False
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    
    cmdImprimir.Caption = "Genera PDF"
    'Limpia los mensajes
    For i = grd.FixedRows To grd.Rows - 1
        grd.TextMatrix(i, COL_C_RESULTADO) = ""
    Next i
    
    For i = grd.FixedRows To grd.Rows - 1
        DoEvents
        If mCancelado Then
            MsgBox "El proceso fue cancelado."
            Exit For
        End If
        
        prg1.value = i
        grd.Row = i
        x = grd.CellTop                 'Para visualizar la celda actual
        
        tid = grd.ValueMatrix(i, COL_TID)
        grd.TextMatrix(i, COL_C_RESULTADO) = "Procesando ..."
        grd.Refresh
        
        'Recupera la transaccion
        Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
        If Not (gnc Is Nothing) Then
            'Si la transacción no está anulado
            If gnc.Estado <> ESTADO_ANULADO And gnc.CodigoMensaje = "60" Then
'                'Forzar recuperar todos los datos de transacción
'                ' para que no se pierdan al grabar de nuveo
'                gnc.RecuperaDetalleTodo
            
                'Imprime la transaccion o asiento contable
                    res = ImprimeTransRide(gnc, mobjImp)
                If res Then
                        grd.TextMatrix(i, COL_C_RESULTADO + 6) = "Generado"
                Else
                    grd.TextMatrix(i, COL_C_RESULTADO + 6) = "Error"
                    cntError = cntError + 1
                End If
                            
            'Si la transaccion está anulado
            Else
                grd.TextMatrix(i, COL_C_RESULTADO) = "Anulado."
                cntError = cntError + 1
            End If
        Else
            grd.TextMatrix(i, COL_C_RESULTADO) = "No pudo recuperar la transación."
            cntError = cntError + 1
        End If
    Next i
    
    Screen.MousePointer = 0
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdImprimir.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    
    'Si algúna transaccion no se imprimió, avisa
    If cntError Then
        MsgBox "No se pudo imprimir " & cntError & " transacciones.", vbInformation
    End If
    
    GeneraPDF = True
    Exit Function
errtrap:
    Screen.MousePointer = 0
    DispErr
    prg1.value = prg1.min
    Exit Function
End Function

Private Sub ConfigColsRIDE()
Dim s As String
    With grd
       
        
s = "^#|(TId)|<Fecha|>Asiento|<Trans|<#Trans|<#Ref.|<Nombre|<Descripción|<|"
s = s & "C.Costo"
s = s & "|^Estado|^EstadoL|>Asiento PR"
s = s & "|<Autorizado|<Mensaje Comprobante Electrónico|<Resultado"

.FormatString = s

'        .ColHidden(COL_NUMFILA) = False
        .ColHidden(COL_TID) = True
'        .ColHidden(COL_FECHA) = False
        .ColHidden(COL_CODASIENTO) = True
'        .ColHidden(COL_CODTRANS) = False
'        .ColHidden(COL_NUMTRANS) = False
        .ColHidden(COL_NUMDOCREF) = True
'        .ColHidden(COL_NOMBRE) = False  'True
'        .ColHidden(COL_DESC) = False
        .ColHidden(COL_CENTROCOSTO) = True
        .ColHidden(COL_ESTADO) = True
        .ColHidden(COL_RESULTADO + 4) = True
'        .ColHidden(COL_RESULTADO + 3) = True
        .ColHidden(COL_RESULTADO + 2) = True
        .ColHidden(COL_RESULTADO + 1) = True
        .ColHidden(COL_RESULTADO) = True
        
        .ColDataType(COL_FECHA) = flexDTDate    '*** MAKOTO 14/ago/2000 para que ordene bien por fecha
        
        GNPoneNumFila grd, False
        .AutoSize 0, grd.Cols - 1
        
        .ColWidth(COL_NUMTRANS) = 1000
        .ColWidth(COL_NOMBRE) = 2400
        .ColWidth(COL_DESC) = 5000
        .ColWidth(COL_RESULTADO + 5) = 2000
    End With
End Sub

Public Function ImprimeTransRide(ByVal gc As GNComprobante, ByRef objImp As Object) As Boolean

    Dim crear As Boolean
    Dim crearRIDE As Boolean
    On Error GoTo errtrap

    'Si no tiene TransID quere decir que no está grabada
    If (gc.TransID = 0) Or gc.Modificado Then
        MsgBox MSGERR_NOGRABADO, vbInformation
        ImprimeTransRide = False
        Exit Function
    End If
    
  
    
    crearRIDE = (objImp Is Nothing)
    If Not crearRIDE Then crearRIDE = (objImp.NombreDLL <> "GNprintg")
    If crearRIDE Then
        Set objImp = Nothing
        Set objImp = CreateObject("GNprintg.PrintTrans")
    End If
    
    MensajeStatus MSG_PREPARA, vbHourglass
    objImp.GeneraTransRide gobjMain.EmpresaActual, True, 1, 0, "", 0, gc
    ImprimeTransRide = True
    MensajeStatus
    'jeaa 30/09/04
    'gc.CambiaEstadoImpresion

    
    Exit Function
errtrap:
    ImprimeTransRide = False
    MensajeStatus
    Select Case Err.Number
    Case ERR_NOIMPRIME, ERR_NOIMPRIME2, ERR_NOIMPRIME3, ERR_NOHAYCODIGO
        DispErr
    Case Else
        
        MsgBox MSGERR_NOIMPRIME2, vbInformation
        
    End Select
    ImprimeTransRide = False
    Exit Function
End Function


Private Sub CargaTransElectronica()
    Dim i As Long, v As Variant
    Dim s As String, cod As String, aux  As Integer, gt As GNTrans

    lstTrans.Clear
    v = gobjMain.GrupoActual.PermisoActual.ListaTransElectronica()
    If UBound(v) > 0 Then
    End If
    For i = LBound(v, 2) To UBound(v, 2) - 1
        lstTrans.AddItem v(0, i)        '& " " & v(1, i)
    Next i
    
    'jeaa 25/09/206
''        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaRecosteo")) > 0 Then
''            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaRecosteo")
''            RecuperaTrans "KeyT", lstTrans, s
''        End If
    
'        If Me.tag = "CostoxProveedor" Then
'                If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaRecosteoxProveedor")) > 0 Then
'                    s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaRecosteoxProveedor")
'                    RecuperaTrans "KeyT", lstTrans, s
'                End If
'
'        Else
            aux = lstTrans.ListIndex
            For i = 0 To lstTrans.ListCount - 1
                cod = lstTrans.List(i)
'                Set gt = gobjMain.EmpresaActual.RecuperaGNTrans(cod)
'                If Not (gt Is Nothing) Then
'                    'Solo marca egresos/transferencias
'                    If gt.IVReprocesaCosto Then
                        lstTrans.Selected(i) = True
'                    End If
'                End If
            Next i
        'End If
    
    
End Sub

Private Function PreparaTransParaGnopcion(cad As String) As String
    Dim v As Variant, i As Integer, s As Variant, gt As GNTrans
    
    s = ""
    v = Split(cad, ",")
    For i = 0 To UBound(v)
        v(i) = Trim(v(i))
        s = s & Mid$(v(i), 2, Len(v(i)) - 2) & ","
    Next i
    PreparaTransParaGnopcion = Mid$(s, 1, Len(s) - 1)
End Function


Private Function PreparaCodTrans() As String
    Dim i As Long, s As String
    
    With lstTrans
        'Si está seleccionado solo una
'        If lstTrans.SelCount = 1 Then
'            For i = 0 To .ListCount - 1
'                If .Selected(i) Then
'                    s = .List(i)
'                    Exit For
'                End If
'            Next i
'        'Si está TODO o NINGUNO, no hay condición
'        ElseIf (.SelCount < .ListCount) And (.SelCount > 0) Then
s = ""
            For i = 0 To .ListCount - 1
                If .Selected(i) Then
                    s = s & "'" & .List(i) & "', "
                End If
            Next i
            If Len(s) > 0 Then s = Left$(s, Len(s) - 2)    'Quita la ultima ", "
'        End If
    End With
    PreparaCodTrans = s
End Function

Public Sub InicioXML()
    Dim i As Integer
    On Error GoTo errtrap
    Me.Caption = "Generar Archivo xml"
    cmdImprimir.Caption = "Genera XML"
    Me.Show
    Me.ZOrder
    dtpFecha1.value = gobjMain.EmpresaActual.GNOpcion.FechaInicio
    dtpFecha2.value = Date
    CargaTrans
    CargaTransElectronicaLote
    FraTransElect.Visible = True
    Exit Sub
errtrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Function GeneraXML() As Boolean
    Dim s As String, tid As Long, i As Long, x As Single, res As String
    Dim gnc As GNComprobante, cambiado As Boolean, cntError As Long
    
    On Error GoTo errtrap

    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdBuscar.Enabled = False
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    
    cmdImprimir.Caption = "Generar XML"
    'Limpia los mensajes
    For i = grd.FixedRows To grd.Rows - 1
        grd.TextMatrix(i, COL_C_RESULTADO) = ""
    Next i
    
    For i = grd.FixedRows To grd.Rows - 1
        DoEvents
        If mCancelado Then
            MsgBox "El proceso fue cancelado."
            Exit For
        End If
        
        prg1.value = i
        grd.Row = i
        x = grd.CellTop                 'Para visualizar la celda actual
        
        tid = grd.ValueMatrix(i, COL_TID)
        grd.TextMatrix(i, COL_C_RESULTADO) = "Procesando ..."
        grd.Refresh
        
        'Recupera la transaccion
        Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
        If Not (gnc Is Nothing) Then
            'Si la transacción no está anulado
            If gnc.Estado <> ESTADO_ANULADO And gnc.CodigoMensaje <> "60" Then
                'Imprime la transaccion o asiento contable
                    res = GeneraComprobanteElectronico(gnc, mobjImp)
                If res Then
                        grd.TextMatrix(i, COL_C_RESULTADO + 6) = "Generado"
                Else
                    grd.TextMatrix(i, COL_C_RESULTADO + 6) = "Error"
                    cntError = cntError + 1
                End If
                            
            'Si la transaccion está anulado
            Else
                grd.TextMatrix(i, COL_C_RESULTADO) = "Anulado."
                cntError = cntError + 1
            End If
        Else
            grd.TextMatrix(i, COL_C_RESULTADO) = "No pudo recuperar la transación."
            cntError = cntError + 1
        End If
    Next i
    
    Screen.MousePointer = 0
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdImprimir.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    
    'Si algúna transaccion no se imprimió, avisa
    If cntError Then
        MsgBox "No se pudo imprimir " & cntError & " transacciones.", vbInformation
    End If
    
    GeneraXML = True
    Exit Function
errtrap:
    Screen.MousePointer = 0
    DispErr
    prg1.value = prg1.min
    Exit Function
End Function


Public Function GeneraComprobanteElectronico(ByVal gc As GNComprobante, ByRef objImp As Object) As Boolean
    Dim crear As Boolean
    Dim crearRIDE As Boolean
    On Error GoTo errtrap

    If gc Is Nothing Then Exit Function
    'Si no tiene TransID quere decir que no está grabada
    If (gc.TransID = 0) Or gc.Modificado Then
        MsgBox MSGERR_NOGRABADO, vbInformation
        GeneraComprobanteElectronico = False
        Exit Function
    End If
    
    
    If gc.CodigoMensaje = "60" Then
        MsgBox "El Documento Electrónico ya fue Autorizado por el SRI "
        Exit Function
'    ElseIf gc.CodigoMensaje = "70" Then
'        MsgBox "El Documento Electrónico está en contigencia "
'        Exit Function
    End If
    
    'Solo por primera vez o cuando cambia la librería de impresión
    '  crea una instancia del objeto para la impresión
    crear = (objImp Is Nothing)
    If Not crear Then crear = (objImp.NombreDLL <> gc.GNTrans.ArchivoReporte)
    If crear Then
        Set objImp = Nothing
        'Set objImp = CreateObject(gc.GNTrans.ArchivoReporteRIDE & ".PrintTrans")
        Set objImp = CreateObject("gnxmla.PrintTrans")
    End If
    
   
    MensajeStatus MSG_PREPARA, vbHourglass
    'jeaa 23/11/2006
    objImp.PrintTrans gobjMain.EmpresaActual, True, 1, 0, "", 0, gc
    MensajeStatus
    'jeaa 30/09/04
'    gc.CambiaEstadoImpresion
    GeneraComprobanteElectronico = True
    
    
    
    Exit Function
errtrap:
    MensajeStatus
    Select Case Err.Number
    Case ERR_NOIMPRIME, ERR_NOIMPRIME2, ERR_NOIMPRIME3, ERR_NOHAYCODIGO
        DispErr
    Case Else
        
        MsgBox MSGERR_NOIMPRIME2, vbInformation
        
    End Select
    GeneraComprobanteElectronico = False
    Exit Function
End Function


Private Sub CargaTransElectronicaLote()
    Dim i As Long, v As Variant
    Dim s As String, cod As String, aux  As Integer, gt As GNTrans

    lstTrans.Clear
    v = gobjMain.GrupoActual.PermisoActual.ListaTransElectronicaLote()
    If UBound(v) > 0 Then
    End If
    For i = LBound(v, 2) To UBound(v, 2) - 1
        lstTrans.AddItem v(0, i)        '& " " & v(1, i)
    Next i
    
    'jeaa 25/09/206
''        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaRecosteo")) > 0 Then
''            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaRecosteo")
''            RecuperaTrans "KeyT", lstTrans, s
''        End If
    
'        If Me.tag = "CostoxProveedor" Then
'                If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaRecosteoxProveedor")) > 0 Then
'                    s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaRecosteoxProveedor")
'                    RecuperaTrans "KeyT", lstTrans, s
'                End If
'
'        Else
            aux = lstTrans.ListIndex
            For i = 0 To lstTrans.ListCount - 1
                cod = lstTrans.List(i)
'                Set gt = gobjMain.EmpresaActual.RecuperaGNTrans(cod)
'                If Not (gt Is Nothing) Then
'                    'Solo marca egresos/transferencias
'                    If gt.IVReprocesaCosto Then
                        lstTrans.Selected(i) = True
'                    End If
'                End If
            Next i
        'End If
    
    
End Sub


Public Sub InicioRecuperaXMLBD()
    Dim i As Integer
    On Error GoTo errtrap
    Me.Caption = "Recupera Archivo xml desde Base Datos"
    cmdImprimir.Caption = "Recupera XML"
    optTrans.Caption = "XML"
    optAsiento.Caption = "XML + PDF"
    Me.Show
    Me.ZOrder
    dtpFecha1.value = gobjMain.EmpresaActual.GNOpcion.FechaInicio
    dtpFecha2.value = Date
    CargaTrans
    CargaTransElectronicaLote
    FraTransElect.Visible = True
    FraVerificaEXistXML.Visible = True
    Exit Sub
errtrap:
    DispErr
    Unload Me
    Exit Sub
End Sub


Private Sub CargaTransSinXML()
    Dim i As Long, v As Variant
    Dim s As String, cod As String, aux  As Integer, gt As GNTrans

    lstTrans.Clear
    v = gobjMain.GrupoActual.PermisoActual.ListaTransRecuperaXMLBD()
    If UBound(v) > 0 Then
    End If
    For i = LBound(v, 2) To UBound(v, 2) - 1
        lstTrans.AddItem v(0, i)        '& " " & v(1, i)
    Next i
    
    'jeaa 25/09/206
''        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaRecosteo")) > 0 Then
''            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaRecosteo")
''            RecuperaTrans "KeyT", lstTrans, s
''        End If
    
'        If Me.tag = "CostoxProveedor" Then
'                If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaRecosteoxProveedor")) > 0 Then
'                    s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaRecosteoxProveedor")
'                    RecuperaTrans "KeyT", lstTrans, s
'                End If
'
'        Else
            aux = lstTrans.ListIndex
            For i = 0 To lstTrans.ListCount - 1
                cod = lstTrans.List(i)
'                Set gt = gobjMain.EmpresaActual.RecuperaGNTrans(cod)
'                If Not (gt Is Nothing) Then
'                    'Solo marca egresos/transferencias
'                    If gt.IVReprocesaCosto Then
                        lstTrans.Selected(i) = True
'                    End If
'                End If
            Next i
        'End If
    
    
End Sub

Private Sub txtCarpeta_LostFocus()
    If Right$(txtCarpeta.Text, 1) <> "\" Then
        txtCarpeta.Text = txtCarpeta.Text & "\"
    End If
    'Luego a actualiza linea de comando
End Sub

Public Sub InicioEtiquetaFacturacion()
    Dim i As Integer, s As String
    On Error GoTo errtrap
    Me.Caption = "Busca Transacciones de Facturacion para Impresion de Etiquetas"
    Me.Show
    Me.ZOrder
    FraBuscaOrden.Visible = False
    framConfig.Visible = True
    dtpFecha1.value = Date
    dtpFecha2.value = Date
    CargaTrans
    fraFecha.Visible = True
    fraNumTrans.Visible = True
    fraCodTrans.Visible = True
    txtNumTrans2.Visible = False
    optTrans.Visible = False
    optAsiento.Visible = False
    fraCodTrans.Caption = "Trans"
    fraNumTrans.Caption = "# Trans "
    cmdCargarTrans.Visible = True
    
    FraConFigEgreso.Visible = False
    FraTipoTrabajo.Visible = True
    FraBuscaOrden.Visible = False
    cboTipoTrabajo.ListIndex = 0
    cmdBuscar.Enabled = False
    cmdImprimiCH.Visible = False
    
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLoteFact_CodTrans")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLoteFact_CodTrans")
            fcbTrans.KeyText = s
        End If
    
    
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLoteFact_MarIzq")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLoteFact_MarIzq")
            ntxMargIzq.value = s
        End If
        
        
        
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLoteFact_MarSup")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLoteFact_MarSup")
            ntxMargSup.value = s
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLoteFact_NumEtiq")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImpLoteFact_NumEtiq")
            ntxNUmEtiq.value = s
        End If
        FraTipoTrabajo.Visible = False
        
        txtNumTrans1.SetFocus
    Exit Sub
errtrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

