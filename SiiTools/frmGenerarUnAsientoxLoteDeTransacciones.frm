VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{50067EB3-D6AF-11D3-8297-000021C5085D}#1.0#0"; "NTextBox.ocx"
Begin VB.Form frmGenerarUnAsientoxLoteDeTransacciones 
   Caption         =   "Generación de un asiento x lote de transacciones"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8520
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   8520
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab sst1 
      Height          =   5175
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Parametros de Busqueda - F6"
      TabPicture(0)   =   "frmGenerarUnAsientoxLoteDeTransacciones.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grd"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdBuscar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraFecha"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraCodTrans"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Asiento - F7"
      TabPicture(1)   =   "frmGenerarUnAsientoxLoteDeTransacciones.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Asiento"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraEnc"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin SiiToolsA.Asiento Asiento 
         Height          =   2895
         Left            =   -74820
         TabIndex        =   32
         Top             =   2040
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   5106
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame fraEnc 
         Height          =   1575
         Left            =   -74880
         TabIndex        =   24
         Top             =   360
         Width           =   8175
         Begin NTextBoxProy.NTextBox ntxCotizacion 
            Height          =   324
            Left            =   960
            TabIndex        =   11
            Top             =   1164
            Width           =   1452
            _ExtentX        =   2566
            _ExtentY        =   582
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
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "&Proceder - F8"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3660
            TabIndex        =   15
            Top             =   1155
            Width           =   1212
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   510
            Left            =   3660
            MaxLength       =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            ToolTipText     =   "Descripción de la transacción"
            Top             =   600
            Width           =   4380
         End
         Begin MSComCtl2.DTPicker dtpFecha 
            Height          =   360
            Left            =   960
            TabIndex        =   9
            ToolTipText     =   "Fecha de la transacción"
            Top             =   492
            Width           =   1452
            _ExtentX        =   2566
            _ExtentY        =   635
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
            Format          =   105578497
            CurrentDate     =   37078
            MaxDate         =   73415
            MinDate         =   29221
         End
         Begin FlexComboProy.FlexCombo fcbResp 
            Height          =   336
            Left            =   6600
            TabIndex        =   13
            ToolTipText     =   "Responsable de la transacción"
            Top             =   240
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
         Begin NTextBoxProy.NTextBox ntxNumTrans 
            Height          =   360
            Left            =   4515
            TabIndex        =   12
            ToolTipText     =   "Número de la transacción"
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
            Text            =   "0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SeparadoPorComa =   0   'False
         End
         Begin FlexComboProy.FlexCombo fcbMoneda 
            Height          =   324
            Left            =   960
            TabIndex        =   10
            ToolTipText     =   "Responsable de la transacción"
            Top             =   840
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
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Cod.Trans  "
            Height          =   195
            Left            =   2775
            TabIndex        =   31
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "&Responsable  "
            Height          =   195
            Left            =   5580
            TabIndex        =   30
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "C&otización  "
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   1230
            Width           =   810
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "&Descripción  "
            Height          =   195
            Left            =   2670
            TabIndex        =   28
            Top             =   600
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&Fecha Transaccion  "
            Height          =   195
            Left            =   1020
            TabIndex        =   27
            Top             =   240
            Width           =   1470
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "&Moneda  "
            Height          =   195
            Left            =   270
            TabIndex        =   26
            Top             =   840
            Width           =   675
         End
         Begin VB.Label lblCodTrans 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   360
            Left            =   3660
            TabIndex        =   25
            ToolTipText     =   "Código de la transacción"
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame fraCodTrans 
         Caption         =   "Cod.&Trans."
         Height          =   1092
         Left            =   3525
         TabIndex        =   23
         Top             =   360
         Width           =   1935
         Begin FlexComboProy.FlexCombo fcbTrans 
            Height          =   348
            Left            =   240
            TabIndex        =   5
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
      Begin VB.Frame fraFecha 
         Caption         =   "&Fecha (desde - hasta)"
         Height          =   1095
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   3495
         Begin MSComCtl2.DTPicker dtpFecha2 
            Height          =   330
            Left            =   1800
            TabIndex        =   3
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            _Version        =   393216
            Format          =   105578497
            CurrentDate     =   36902
         End
         Begin MSComCtl2.DTPicker dtpHora1 
            Height          =   330
            Left            =   120
            TabIndex        =   2
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            _Version        =   393216
            Format          =   105578498
            CurrentDate     =   36902
         End
         Begin MSComCtl2.DTPicker dtpHora2 
            Height          =   330
            Left            =   1800
            TabIndex        =   4
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            _Version        =   393216
            Format          =   105578498
            CurrentDate     =   36902
         End
         Begin MSComCtl2.DTPicker dtpFecha1 
            Height          =   330
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            _Version        =   393216
            Format          =   105578497
            CurrentDate     =   36902
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "~  "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   1605
            TabIndex        =   22
            Top             =   480
            Width           =   315
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar - F5"
         Height          =   372
         Left            =   840
         TabIndex        =   7
         Top             =   1560
         Width           =   1212
      End
      Begin VB.Frame Frame5 
         Caption         =   "Cod.&Trans de Asiento."
         Height          =   1092
         Left            =   5400
         TabIndex        =   20
         Top             =   360
         Width           =   1935
         Begin FlexComboProy.FlexCombo fcbTransAsiento 
            Height          =   348
            Left            =   240
            TabIndex        =   6
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
      Begin VSFlex7LCtl.VSFlexGrid grd 
         Height          =   2775
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   8175
         _cx             =   14420
         _cy             =   4895
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
   End
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   852
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   8520
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5565
      Width           =   8520
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar -F3"
         Height          =   372
         Left            =   2880
         TabIndex        =   16
         Top             =   360
         Width           =   1332
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   4320
         TabIndex        =   17
         Top             =   360
         Width           =   1212
      End
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   60
         Width           =   8280
         _ExtentX        =   14605
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "frmGenerarUnAsientoxLoteDeTransacciones"
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
Private Const COL_NUMDOCREF = 6
Private Const COL_NOMBRE = 7
Private Const COL_DESC = 8
Private Const COL_CENTROCOSTO = 9
Private Const COL_ESTADO = 10
Private Const COL_RESULTADO = 11

Private mProcesando As Boolean
Private mCancelado As Boolean
Private mVerificado As Boolean

Private WithEvents mobjGNComp As GNComprobante
Attribute mobjGNComp.VB_VarHelpID = -1

Public Sub Inicio()
    Dim i As Integer
    On Error GoTo ErrTrap
    
    Me.Show
    Me.ZOrder
    dtpFecha1.value = gobjMain.EmpresaActual.GNOpcion.FechaInicio
    dtpFecha2.value = Date
    dtpHora1.value = CDate(0)
    dtpHora2.value = CDate(0.99999)  ' jeaa para que sean las 23:59:59
    CargarEncabezado
    CargaTrans
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub CargaTrans()
    'Carga la lista de transacción
    fcbTrans.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False)
    fcbTransAsiento.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "CT")
End Sub



Private Sub cmdAceptar_Click()
    If Not mProcesando Then
        'Si no hay transacciones
        If grd.Rows <= grd.FixedRows Then
            MsgBox "No hay ningúna transacción para procesar."
            Exit Sub
        End If
        Enc_Aceptar
        
        If GenerarAsiento() Then
            cmdCancelar.SetFocus
        End If
    End If
End Sub

Private Function GenerarAsiento() As Boolean
    Dim s As String, tid As Long, i As Long, x As Single
    Dim gnc As GNComprobante, cambiado As Boolean
    
    On Error GoTo ErrTrap

    
'    s = "Este proceso modificará los asientos de la transacción seleccionada." & vbCr & vbCr
'    s = s & "Está seguro que desea proceder?"
'    If MsgBox(s, vbYesNo + vbQuestion) <> vbYes Then Exit Function
    
    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdBuscar.Enabled = False
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    
    ' Proceeder a Generar un solo asiento por todas las transacciones que Estan.

    If Not mobjGNComp Is Nothing Then
        mobjGNComp.Generar1AsientoxLote
        Asiento.VisualizaDesdeObjeto
    End If
    
    Screen.MousePointer = 0
    GenerarAsiento = Not mCancelado
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    DispErr
salida:
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    Exit Function
End Function




Private Sub cmdBuscar_Click()
    Dim v As Variant, obj As Object
    On Error GoTo ErrTrap
    
    If Len(fcbTrans.Text) = 0 Then
        MsgBox "Seleccione solo un tipo de transacción", vbInformation
        fcbTrans.SetFocus
        Exit Sub
    End If
    
    With gobjMain.objCondicion
        .fecha1 = dtpFecha1.value
        .fecha2 = dtpFecha2.value
        .Hora1 = dtpHora1.value
        .Hora2 = dtpHora2.value
        .CodTrans = fcbTrans.Text
        
        'Estados no incluye anulados
        .EstadoBool(ESTADO_NOAPROBADO) = True
        .EstadoBool(ESTADO_APROBADO) = True
        .EstadoBool(ESTADO_DESPACHADO) = True
        .EstadoBool(ESTADO_ANULADO) = False
    End With
    Set obj = gobjMain.EmpresaActual.ConsGNTrans3(True) 'Ascendente
    
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
    cmdAceptar.Enabled = True
    mVerificado = True
    Exit Sub
ErrTrap:
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
        .ColHidden(COL_NOMBRE) = False      'True
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


Private Sub cmdGrabar_Click()
    If Grabar Then
        'cmdGrabar.Enabled = False
         Limpiar
    Else
        'cmdGrabar.Enabled = True
    End If
End Sub


Private Sub fcbTrans_BeforeSelect(ByVal Row As Long, Cancel As Boolean)
    SacaTransAsientoGnTrans fcbTrans.Text
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF3
        cmdGrabar_Click
        KeyCode = 0
    Case vbKeyF5
        If sst1.Tab = 0 Then cmdBuscar_Click
        KeyCode = 0
    Case vbKeyF6
        sst1.Tab = 0
        dtpFecha1.SetFocus
        KeyCode = 0
    Case vbKeyF7
        sst1.Tab = 1
        dtpFecha.SetFocus
        KeyCode = 0
    Case vbKeyF8
        If sst1.Tab = 1 Then cmdAceptar_Click
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
    Me.Hide         'Se pone esto para evitar el posible BUG de Windows98
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    
    sst1.Move 0, sst1.Top, Me.ScaleWidth, Me.ScaleHeight - pic1.Height - 300
    
    With grd
        .Width = Me.ScaleWidth - 200
        .Height = Me.ScaleHeight - .Top - pic1.Height - 380
    End With
    
    With Asiento
        .Width = Me.ScaleWidth - 200
        .Height = Me.ScaleHeight - .Top - pic1.Height - 380
    End With
    
    prg1.Width = Me.ScaleWidth - (prg1.Left * 2)
        
End Sub




Private Sub grd_LostFocus()
    If sst1.Tab = 0 Then
        sst1.Tab = 1
    End If
End Sub

Private Sub mobjGNComp_EstadoGeneracion1AsientoxLote(ByVal ix As Long, ByVal Estado As String, Cancel As Boolean)
    prg1.value = ix
    grd.TextMatrix(ix, COL_RESULTADO) = Estado
    Cancel = mCancelado
End Sub


Private Sub sst1_Click(PreviousTab As Integer)
    '*** Para evitar error de ciclo infinito
 
    On Error GoTo ErrTrap
    Select Case sst1.Tab
    Case 0          'Parametros de Busqueda
    
    Case 1          'Transaccion de Asiento
        If lblCodTrans.Caption <> fcbTransAsiento.KeyText Then
            lblCodTrans.Caption = fcbTransAsiento.KeyText
            SacaDatosGnTrans (lblCodTrans.Caption)
            CrearGnComprobante
        End If
        PoneDescripcion
    End Select
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub CargarEncabezado()
    dtpFecha.value = Date
    fcbResp.SetData gobjMain.EmpresaActual.ListaGNResponsable(False)
    fcbMoneda.SetData gobjMain.EmpresaActual.ListaGNMoneda
    fcbMoneda.KeyText = "USD"
    ntxCotizacion.Text = " 1"
    txtDescripcion.Text = "Asiento de Cierre de ..."
End Sub


Private Sub SacaDatosGnTrans(ByVal CodTrans As String)
    Dim gnt As GNTrans
    Set gnt = gobjMain.EmpresaActual.RecuperaGNTrans(CodTrans)
    If Not gnt Is Nothing Then
        fcbResp.KeyText = gnt.CodResponsablePre
    End If
End Sub

Private Sub SacaTransAsientoGnTrans(ByVal CodTrans As String)
    Dim gnt As GNTrans
    Set gnt = gobjMain.EmpresaActual.RecuperaGNTrans(CodTrans)
    If Not gnt Is Nothing Then
        fcbTransAsiento.KeyText = gnt.TransAsiento
    End If
End Sub

Private Sub CrearGnComprobante()
    'Eliminar el que haya tenido
    Set mobjGNComp = Nothing
    'crear el comprobante para luego grabar
    If Len(lblCodTrans.Caption) = 0 Then
        MsgBox "No hay un tipo de transacción para crear"
    Else
        Set mobjGNComp = gobjMain.EmpresaActual.CreaGNComprobante(lblCodTrans.Caption)
        Set Asiento.GNComprobante = mobjGNComp
    End If
    
End Sub

Private Sub Enc_Aceptar()
    
    If fcbResp.Vacio Then
        fcbResp.SetFocus
        MsgBox "Seleccione el responsable..", vbInformation
        Exit Sub
    End If
    
    'Cotización no puede ser 0
    If ntxCotizacion.value <= 0 Then
        ntxCotizacion.SetFocus
        MsgBox "La cotización no puede ser 0.", vbInformation
        Exit Sub
    End If

    'Si es hora automático, graba con la hora del momento de grabación
    'Solo  cuando no es modificacion
    
    
    If Not (mobjGNComp Is Nothing) Then
        mobjGNComp.FechaTrans = dtpFecha.value
        mobjGNComp.numtrans = ntxNumTrans.value
        mobjGNComp.CodResponsable = fcbResp.KeyText
        mobjGNComp.CodMoneda = fcbMoneda.Text
        mobjGNComp.Cotizacion("") = ntxCotizacion.value
        mobjGNComp.Descripcion = Trim$(txtDescripcion.Text)
        If mobjGNComp.GNTrans.HoraAuto And mobjGNComp.EsNuevo = True Then
            mobjGNComp.HoraTrans = Time
        End If
    End If
End Sub


Private Function Grabar() As Boolean
    On Error GoTo ErrTrap
    
    If mobjGNComp Is Nothing Then Exit Function
        
    Enc_Aceptar
    Asiento.Aceptar
    
    'Verifica si tiene detalle
    If mobjGNComp.CountCTLibroDetalle = 0 Then
        MsgBox "No existe ningún detalle.", vbInformation
        Asiento.SetFocus
        Exit Function
    End If

    'Verificación de datos
    mobjGNComp.VerificaDatos

    'Verifica si está cuadrado el asiento
    If Not VerificaAsiento(mobjGNComp) Then Exit Function
    
    'Manda a grabar
    '       Aquí ya no hacemos verificación de asiento por que ya está hecho en Control Asiento
    mobjGNComp.Grabar False, False
    
    MsgBox "Asiento grabado con exito " & mobjGNComp.CodTrans & " " & mobjGNComp.numtrans, vbInformation
    ntxNumTrans.Text = mobjGNComp.numtrans
    Me.Caption = "Asiento #" & mobjGNComp.CodAsiento
        
    Grabar = True
    Exit Function
ErrTrap:
    MensajeStatus
    Select Case Err.Number
    Case ERR_DESCUADRADO, ERR_INTEGRIDAD
        'Si es que el usuario seleccionó 'No' en el cuadro de dialogo,
        'No hace nada
    Case Else
        DispErr
    End Select
    Asiento.SetFocus  'Para que no se pierda el enfoque
    Exit Function
End Function


Private Sub PoneDescripcion()
    Dim gnt As GNTrans, fdesde As Date, fhasta As Date
    'pone descripcion en la transaccion
    Set gnt = gobjMain.EmpresaActual.RecuperaGNTrans(fcbTrans.Text)
    If Not gnt Is Nothing Then
        fdesde = (dtpFecha1.value + dtpHora1.value)
        fhasta = (dtpFecha2.value + dtpHora2.value)
        txtDescripcion.Text = "Resumen de Asiento de " & gnt.Descripcion & " desde " & fdesde & " hasta " & fhasta
    Else
        txtDescripcion.Text = "Error inesperado en la transaccion seleccionada"
    End If
    ntxNumTrans = 0
    Set gnt = Nothing
End Sub

Public Sub Limpiar()
    Dim i As Long
    
    With grd
        For i = .FixedRows To .Rows - 1
            .RowData(i) = 0
        Next i
        .Rows = .FixedRows
    End With
    Asiento.Limpiar
    Enc_Limpiar
End Sub

Private Sub Enc_Limpiar()
    lblCodTrans.Caption = fcbTransAsiento.KeyText
    SacaDatosGnTrans (lblCodTrans.Caption)
    CrearGnComprobante
    PoneDescripcion
End Sub

