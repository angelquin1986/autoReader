VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{50067EB3-D6AF-11D3-8297-000021C5085D}#1.0#0"; "NTextBox.ocx"
Begin VB.Form frmGenerarDepreciacion 
   Caption         =   "Generación Depreciaciones"
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
   Begin VB.Frame fraGrupos 
      Caption         =   "Rango de Grupos"
      Height          =   855
      Left            =   60
      TabIndex        =   14
      Top             =   840
      Width           =   7335
      Begin VB.ComboBox cboGrupo 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   480
         Width           =   1452
      End
      Begin FlexComboProy.FlexCombo fcbGrupoDesde 
         Height          =   300
         Left            =   1812
         TabIndex        =   16
         Top             =   480
         Width           =   1452
         _ExtentX        =   2566
         _ExtentY        =   529
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
      Begin FlexComboProy.FlexCombo fcbDesde2 
         Height          =   315
         Left            =   3300
         TabIndex        =   17
         Top             =   480
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   556
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Grupo"
         Height          =   192
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   444
      End
      Begin VB.Label lblTipo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Desde"
         Height          =   192
         Left            =   1800
         TabIndex        =   19
         Top             =   240
         Width           =   492
      End
      Begin VB.Label Label11 
         Caption         =   "Activo Fijo"
         Height          =   255
         Left            =   3300
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraCodTrans 
      Caption         =   "Cod. &Trans. Depreciación"
      Height          =   735
      Left            =   2940
      TabIndex        =   12
      Top             =   60
      Width           =   4455
      Begin NTextBoxProy.NTextBox ntxNumPer 
         Height          =   315
         Left            =   3660
         TabIndex        =   22
         Top             =   300
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   556
         Text            =   "1"
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
      End
      Begin FlexComboProy.FlexCombo fcbTrans 
         Height          =   345
         Left            =   600
         TabIndex        =   1
         Top             =   300
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         ColWidth0       =   500
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
      Begin VB.Label Label2 
         Caption         =   "Periodos a Depre"
         Height          =   255
         Left            =   2040
         TabIndex        =   21
         Top             =   360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Trans."
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar Activos - F5"
      Height          =   372
      Left            =   60
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Frame fraFecha 
      Caption         =   "&Fecha Depreciación"
      Height          =   735
      Left            =   60
      TabIndex        =   11
      Top             =   60
      Width           =   2835
      Begin MSComCtl2.DTPicker dtpFecha1 
         Height          =   330
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         _Version        =   393216
         Format          =   106692609
         CurrentDate     =   36902
      End
   End
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   852
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   8520
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5565
      Width           =   8520
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   372
         Left            =   4080
         TabIndex        =   6
         Top             =   360
         Width           =   1452
      End
      Begin VB.CommandButton cmdAsiento 
         Caption         =   "&Asiento"
         Enabled         =   0   'False
         Height          =   372
         Left            =   2880
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Proceder - F8"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   360
         Width           =   1212
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar -F3"
         Height          =   372
         Left            =   8340
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   9660
         TabIndex        =   7
         Top             =   360
         Width           =   1212
      End
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   60
         Width           =   8280
         _ExtentX        =   14605
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   2775
      Left            =   60
      TabIndex        =   3
      Top             =   2280
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
      AllowSelection  =   -1  'True
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
Attribute VB_Name = "frmGenerarDepreciacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit


'Constantes para las columnas

Private Const COL_DEP_NUMFILA = 0
Private Const COL_DEP_DEP = 1
Private Const COL_DEP_CODGRUPO = 2
Private Const COL_DEP_DESCGRUPO = 3
Private Const COL_DEP_ID = 4
Private Const COL_DEP_CODAFINV = 5
Private Const COL_DEP_DESCSAFINV = 6
Private Const COL_DEP_TIPODEPRE = 7
Private Const COL_DEP_VIDAUTIL = 8
Private Const COL_DEP_FECHA = 9
Private Const COL_DEP_FECHAINIDEP = 10
Private Const COL_DEP_COSTOCOMPRA = 11
Private Const COL_DEP_COSTORESIDUAL = 12
Private Const COL_DEP_DEPANTERIOR = 13
Private Const COL_DEP_IDCTADEPRE = 14
Private Const COL_DEP_IDCTADEPREACUM = 15
Private Const COL_DEP_RESULTADO = 16
Private Const COL_DEP_TIDIN = 17
Private Const COL_DEP_NUMTRANSIN = 18



Private mProcesando As Boolean
Private mCancelado As Boolean
Private mVerificado As Boolean

Private WithEvents mobjGNComp As GNComprobante
Attribute mobjGNComp.VB_VarHelpID = -1
Private mobjGNCompOrigen As GNComprobante
Private mobjGNCompAux As GNComprobante
Private mColItems As Collection
Private Const MSG_NG = "Asiento incorrecto."
Private mCodMoneda As String
Private numGrupo As Integer
Private mobjAF As AFinventario

Public Sub Inicio(Name As String)
    Dim i As Integer
    On Error GoTo ErrTrap
        Me.tag = Name
        For i = 1 To AFGRUPO_MAX
            cboGrupo.AddItem gobjMain.EmpresaActual.GNOpcion.EtiqAFGrupo(i)
        Next i
        If (numGrupo <= cboGrupo.ListCount) And (numGrupo > 0) Then
            cboGrupo.ListIndex = numGrupo - 1   'Selecciona lo anterior
        ElseIf cboGrupo.ListCount > 0 Then
            cboGrupo.ListIndex = 0              'Selecciona la primera
        End If
        numGrupo = cboGrupo.ListIndex + 1
        ConfigCols
        Me.Show
        Me.ZOrder
        dtpFecha1.value = Date
        CargaTrans
        Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub CargaTrans()
    Dim i As Long, v As Variant
    Dim s As String
    
    fcbTrans.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "AF")
    'jeaa 25/09/206
    If Me.tag = "Depre" Then
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransParaDepreciacion")) > 0 Then
            fcbTrans.KeyText = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransParaDepreciacion")
        End If
    ElseIf Me.tag = "DepreReval" Then
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransParaDepreciacionReval")) > 0 Then
            fcbTrans.KeyText = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransParaDepreciacionReval")
        End If
    ElseIf Me.tag = "DepreRevalA" Then
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransParaDepreciacionRevalA")) > 0 Then
            fcbTrans.KeyText = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransParaDepreciacionRevalA")
        End If
    ElseIf Me.tag = "DepreANT" Then
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransParaDepreciacion")) > 0 Then
            fcbTrans.KeyText = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransParaDepreciacion")
        End If
        
    End If
End Sub

Private Sub cmdAceptar_Click()
    If Not mProcesando Then
        'Si no hay transacciones
        If grd.Rows <= grd.FixedRows Then
            MsgBox "No hay ningún Activo Fijo para procesar."
            Exit Sub
        End If
        If Len(fcbTrans.KeyText) = 0 Then
            MsgBox "No hay ningúna transacción de Depreciación no se podrá procesar."
            fcbTrans.SetFocus
            Exit Sub
        End If
    
    End If
    
    If Me.tag = "Depre" Then
        If DepreciacionAuto(True, False) Then
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
            mVerificado = True
            cmdAsiento.Enabled = False
            cmdImprimir.Enabled = True
        Else
            cmdAceptar.Enabled = False
            cmdAsiento.Enabled = True
        End If
    ElseIf Me.tag = "DepreReval" Then
        If DepreciacionAutoReval(True, False, False) Then
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
            mVerificado = True
            cmdAsiento.Enabled = False
            cmdImprimir.Enabled = True
        Else
            cmdAceptar.Enabled = False
            cmdAsiento.Enabled = True
        End If
    ElseIf Me.tag = "DepreRevalA" Then
        If DepreciacionAutoReval(True, False, True) Then
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
            mVerificado = True
            cmdAsiento.Enabled = False
            cmdImprimir.Enabled = True
        Else
            cmdAceptar.Enabled = False
            cmdAsiento.Enabled = True
        End If
    ElseIf Me.tag = "DepreANT" Then
        If DepreciacionAuto(True, False) Then
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
            mVerificado = True
            cmdAsiento.Enabled = False
            cmdImprimir.Enabled = True
        Else
            cmdAceptar.Enabled = False
            cmdAsiento.Enabled = True
        End If
    
    End If
    
End Sub





Private Sub cmdAsiento_Click()
'''    If grd.Rows <= grd.FixedRows Then
'''        MsgBox "No hay ningúna transacción para procesar."
'''        Exit Sub
'''    End If
'''
'''
'''    If RegenerarAsiento(True, True) Then
'''        cmdCancelar.SetFocus
'''        cmdAsiento.Enabled = False
'''        cmdImprimir.Enabled = True
'''    Else
'''        cmdImprimir.Enabled = False
'''    End If

End Sub

Private Sub cmdBuscar_Click()
    Dim v As Variant, obj As Object, s As String, i As Long
    On Error GoTo ErrTrap
    With gobjMain.objCondicion
        If cboGrupo.ListIndex >= 0 Then
            numGrupo = cboGrupo.ListIndex + 1
            .Grupo1 = Trim$(fcbGrupoDesde.KeyText)
            .CodItem1 = Trim$(fcbDesde2.KeyText)
            .FechaCorte = dtpFecha1.value
        End If
    
        mCodMoneda = GetSetting(APPNAME, SECTION, Me.Name & "_" & Me.tag & "_Moneda", "USD")    '*** MAKOTO 08/sep/00
        .fecha1 = dtpFecha1.value
        gobjMain.objCondicion.CodMoneda = mCodMoneda
        
        s = fcbTrans.KeyText
        If Me.tag = "Depre" Then
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransParaDepreciacion", s
        ElseIf Me.tag = "DepreReval" Then
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransParaDepreciacionReval", s
        ElseIf Me.tag = "DepreRevalA" Then
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransParaDepreciacionRevalA", s
        ElseIf Me.tag = "DepreANT" Then
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransParaDepreciacionANT", s
        End If
    
        'Graba en la base
        gobjMain.EmpresaActual.GNOpcion.Grabar
            
        grd.Rows = 1
        If Me.tag = "Depre" Then
            Set obj = gobjMain.EmpresaActual.ConsAFInvetario(numGrupo, .Grupo1, .CodItem1)  'Ascendente
        ElseIf Me.tag = "DepreReval" Then
            Set obj = gobjMain.EmpresaActual.ConsAFInvetarioReval(numGrupo, .Grupo1, .CodItem1, True) 'Ascendente
        ElseIf Me.tag = "DepreRevalA" Then
            Set obj = gobjMain.EmpresaActual.ConsAFInvetarioReval(numGrupo, .Grupo1, .CodItem1, False) 'Ascendente
        ElseIf Me.tag = "DepreANT" Then
            'Set obj = gobjMain.EmpresaActual.ConsAFInvetarioDepAnterior(numGrupo, .Grupo1, .CodItem1)  'Ascendente
            
        End If
    End With
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
    If Me.tag = "DepreRevalA" Then
        For i = 1 To grd.Rows - 1
            grd.TextMatrix(i, COL_DEP_VIDAUTIL) = grd.ValueMatrix(i, COL_DEP_VIDAUTIL) * -1
            grd.TextMatrix(i, COL_DEP_COSTOCOMPRA) = grd.ValueMatrix(i, COL_DEP_COSTOCOMPRA) * -1
        Next i
    End If
    cmdAceptar.Enabled = True
    cmdAsiento.Enabled = False
    cmdImprimir.Enabled = True
    mVerificado = True
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub ConfigCols()
    Dim i As Integer
    With grd
        If Me.tag = "Depre" Or Me.tag = "DepreANT" Then
            .FormatString = "^#|CTA|<Cod. " & gobjMain.EmpresaActual.GNOpcion.EtiqAFGrupo(numGrupo) & _
                                "|<Desc. " & gobjMain.EmpresaActual.GNOpcion.EtiqAFGrupo(numGrupo) & _
                                "|>idInventario|<Código|<Descripcion|^Tipo Dep|>Vida Util " & _
                                "|^Fecha Compra|^Fecha Ini Dep |>Costo Compra|>Costo Residual|>Dep. Anterior" & _
                                " |>IdCuentaDepreGasto|>IdCuentaDepreAcumulada|<Resultado|tidIn"
            
        ElseIf Me.tag = "DepreReval" Or Me.tag = "DepreRevalA" Then
            .FormatString = "^#|CTA|<Cod. " & gobjMain.EmpresaActual.GNOpcion.EtiqAFGrupo(numGrupo) & _
                                "|<Desc. " & gobjMain.EmpresaActual.GNOpcion.EtiqAFGrupo(numGrupo) & _
                                "|>idInventario|<Código|<Descripcion|^Tipo Dep|>Num Dep Reval " & _
                                "|^Fecha Reval|^Fecha Ini Dep Reval|>Valor Reval|>Costo Residual|>Dep. Anterior" & _
                                " |>IdCuentaDepreGasto|>IdCuentaDepreAcumulada|<Resultado|tidIn"
        End If
           
           .ColHidden(COL_DEP_ID) = True
'''        .ColHidden(COL_DEP_NUMFILA) = False
'''        .ColHidden(COL_DEP_FECHA) = False
'''        .ColHidden(COL_DEP_FECHAINIDEP) = False
        .ColHidden(1) = True
        .ColHidden(2) = True
        .ColHidden(3) = True
        .ColHidden(COL_DEP_IDCTADEPRE) = True
        .ColHidden(COL_DEP_IDCTADEPREACUM) = True
        .ColHidden(COL_DEP_TIDIN) = True
        .ColHidden(COL_DEP_ID) = True
'''        .ColHidden(COL_DEP_CODGRUPO) = True
'''        .ColHidden(COL_DEP_DEP) = True
        
        .ColDataType(COL_DEP_FECHA) = flexDTDate
        .ColDataType(COL_DEP_FECHAINIDEP) = flexDTDate
        .ColDataType(COL_DEP_COSTOCOMPRA) = flexDTCurrency
        .ColDataType(COL_DEP_COSTORESIDUAL) = flexDTCurrency
        .ColDataType(COL_DEP_FECHA) = flexDTDate   '*** MAKOTO 14/ago/2000 para que ordene bien por fecha
        
        .ColFormat(COL_DEP_COSTOCOMPRA) = "#,#0.0000"
        .ColFormat(COL_DEP_COSTORESIDUAL) = "#,#0.0000"
        .SubtotalPosition = flexSTBelow
        grd.SubTotal flexSTClear
        For i = 1 To COL_DEP_RESULTADO
            .ColData(i) = -1
        Next i
        
        
'        If Me.tag = "Depre" Or Me.tag = "DepreReval" Then
            .ColFormat(COL_DEP_FECHA) = "dd/mm/yyyy"
            .ColFormat(COL_DEP_FECHAINIDEP) = "mmm/yyyy"
'''            grd.subtotal flexSTSum, COL_DEP_DEP, 1, , grd.GridColor, vbBlack, , "Subtotal", COL_DEP_DEP, True
 '       End If
        GNPoneNumFila grd, False
        .AutoSize 0, grd.Cols - 1
        
        .ColWidth(COL_DEP_RESULTADO) = 2000
        .ColWidth(COL_DEP_DESCGRUPO) = 2000
        
    End With
    PoneColorFilas
End Sub

Private Sub cmdCancelar_Click()
    If mProcesando Then
        mCancelado = True
    Else
        Unload Me
    End If
End Sub

Private Sub cmdImprimir_Click()
    'Si no hay transacciones
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna transacción para imprimir."
        Exit Sub
    End If
    
    If Imprimir Then
        cmdCancelar.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF3
        KeyCode = 0
    Case vbKeyF5
        KeyCode = 0
    Case vbKeyF6
        dtpFecha1.SetFocus
        KeyCode = 0
    Case vbKeyF7
        KeyCode = 0
    Case vbKeyF8
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
    With grd
'        If Me.tag = "Depre" Then
                '.Top = FraConFigEgreso.Top
                .Width = Me.ScaleWidth - 200
                .Height = Me.ScaleHeight - .Top - pic1.Height - 380
'        End If
    End With
    prg1.Width = Me.ScaleWidth - (prg1.Left * 2)
        
End Sub


Private Sub grd_CellChanged(ByVal Row As Long, ByVal col As Long)
''    grd.subtotal flexSTSum, -1, 10, , vbBlue, vbWhite
''    grd.subtotal flexSTSum, -1, 12, , vbBlue, vbWhite
'    grd.Refresh
End Sub



'jeaa 25/09/2006 elimina los apostrofes
Private Function PreparaTransParaGnopcion(cad As String) As String
    Dim v As Variant, i As Integer, s As String, pos As Integer
    s = ""
    v = Split(cad, ",")
    For i = 0 To UBound(v)
        v(i) = Trim(v(i))
        pos = InStr(1, v(i), "'")
        If pos <> 0 Then
            s = s & Mid$(v(i), 2, Len(v(i)) - 2) & ","
        Else
            s = s & v(i) & ","
        End If
    Next i
    'quita ultima coma
    PreparaTransParaGnopcion = Mid$(s, 1, Len(s) - 1)
End Function


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
                If Trim(Vector(i)) = lst.List(j) Then
                    lst.Selected(j) = True
                End If
            Next j
         Next i
    End If
End Sub

'Private Function RegenerarAsiento(bandVerificar As Boolean, bandTodo As Boolean) As Boolean
Private Function RegenerarAsiento(ByRef trans As String, FilaIni As Long, FilaFin As Long) As Boolean
    Dim s As String, tid As Long, i As Long, x As Single, pos As Integer
    Dim gnc As GNComprobante, cambiado As Boolean
    Dim bandAsiento As Boolean
    On Error GoTo ErrTrap
    bandAsiento = True
    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdBuscar.Enabled = False
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    
    'For i = grd.FixedRows To grd.Rows - 1
    
    tid = grd.ValueMatrix(FilaIni, COL_DEP_TIDIN)
    'Recupera la transaccion
    Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
    If Not (gnc Is Nothing) Then
        'Forzar recuperar todos los datos de transacción para que no se pierdan al grabar de nuveo
        gnc.RecuperaDetalleTodo

        For i = FilaIni To FilaFin - 1
            DoEvents
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit For
            End If
            prg1.value = i
            grd.Row = i
            x = grd.CellTop                 'Para visualizar la celda actual
            'Si la transacción no está anulada
                If InStr(1, grd.TextMatrix(i, COL_DEP_RESULTADO), "ERROR") = 0 Then
                If gnc.Estado <> ESTADO_ANULADO Then
                    'Recalcula costo de los items
                    If RegenerarAsientoSub(gnc, cambiado) Then
                        'Si está cambiado algo o está forzado regenerar todo
                        'Graba la transacción
                        bandAsiento = True
    '                    gnc.Grabar False, False
                        grd.TextMatrix(i, COL_DEP_RESULTADO) = "Actualizando..."
                    Else
                        bandAsiento = False
    '                    gnc.Estado = 0
    '                    gnc.Estado = ESTADO_NOAPROBADO
    '                    gnc.Grabar False, False
                        grd.TextMatrix(i, COL_DEP_RESULTADO) = "Falló al generar asiento."
                        gobjMain.EmpresaActual.CambiaEstadoGNComp tid, ESTADO_NOAPROBADO
                        Exit For
                    End If
                Else
                    'Si está anulada
                    grd.TextMatrix(i, COL_DEP_RESULTADO) = "Anulado."
                End If
            End If
        Next i
    Else
        grd.TextMatrix(i, COL_DEP_RESULTADO) = "No pudo recuperar la transación."
    End If
    If bandAsiento Then gnc.Grabar False, False
    
    
    Screen.MousePointer = 0
    RegenerarAsiento = Not mCancelado
    
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


Private Function RegenerarAsientoSub(ByVal gnc As GNComprobante, _
                                     ByRef cambiado As Boolean) As Boolean
    Dim i As Long, cta As CtCuenta, ctd As CTLibroDetalle
    Dim colCtd As Collection, a As clsAsiento
    On Error GoTo ErrTrap
    
    cambiado = False
    Set colCtd = New Collection
    
    'Guarda todos los detalles de asiento en la colección para después comparar
    With gnc
        For i = 1 To .CountCTLibroDetalle
            Set ctd = .CTLibroDetalle(i)
            Set a = New clsAsiento
            a.IdCuenta = ctd.IdCuenta
            a.Debe = ctd.Debe
            a.Haber = ctd.Haber
            colCtd.Add item:=a
        Next i
    End With
    
    'Regenera el asiento
    gnc.GeneraAsiento
    
    'Compara el asiento para saber si ha cambiado o no
    cambiado = Not CompararAsiento(gnc, colCtd)
    
    RegenerarAsientoSub = True
    GoTo salida
    Exit Function
ErrTrap:
    cambiado = False
    DispErr
    RegenerarAsientoSub = False
salida:
    Set a = Nothing
    Set colCtd = Nothing
    Set gnc = Nothing
    Exit Function
End Function


'Devuelve True si los asientos son iguales, False si no lo son
Private Function CompararAsiento(ByVal gnc As GNComprobante, ByVal col As Collection) As Boolean
    Dim a As clsAsiento, i As Long, ctd As CTLibroDetalle
    Dim encontrado As Boolean
    
    'Si número de detalles son diferentes ya no son iguales
    If col.Count <> gnc.CountCTLibroDetalle Then Exit Function
    
    For i = 1 To gnc.CountCTLibroDetalle
        Set ctd = gnc.CTLibroDetalle(i)
        encontrado = False
        For Each a In col
            If (ctd.IdCuenta = a.IdCuenta) And _
               (ctd.Debe = a.Debe) And _
               (ctd.Haber = a.Haber) And _
               (a.Comparado = False) Then
                a.Comparado = True
                encontrado = True
                Exit For
            End If
        Next a
        'Si no se encuentra uno igual
        If Not encontrado Then
            CompararAsiento = False
            Exit Function
        End If
    Next i
    CompararAsiento = True
End Function


Private Function Imprimir() As Boolean
    Dim s As String, tid As Long, i As Long, x As Single, res As String, pos As Integer
    Dim gnc As GNComprobante, cambiado As Boolean, cntError As Long
    
    On Error GoTo ErrTrap

    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdBuscar.Enabled = False
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
        pos = InStr(1, grd.TextMatrix(i, COL_DEP_RESULTADO), "OK")
        'Si es verificación, procesa todas las filas sino solo las que tengan "Asiento incorrecto."
        If pos <> 0 Then
        
            tid = grd.ValueMatrix(i, COL_DEP_TIDIN)
            grd.TextMatrix(i, COL_DEP_RESULTADO) = "Procesando ..."
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
                    res = ImprimeTrans(gnc, False)
                    If Len(res) = 0 Then
                        grd.TextMatrix(i, COL_DEP_RESULTADO) = "OK..Enviado."
                    Else
                        grd.TextMatrix(i, COL_DEP_RESULTADO) = res
                        cntError = cntError + 1
                    End If
                                
                'Si la transaccion está anulado
                Else
                    grd.TextMatrix(i, COL_DEP_RESULTADO) = "Anulado."
                    cntError = cntError + 1
                End If
            Else
                grd.TextMatrix(i, COL_DEP_RESULTADO) = "No pudo recuperar la transación."
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
    
    Imprimir = True
    Exit Function
ErrTrap:
    Screen.MousePointer = 0
    DispErr
    prg1.value = prg1.min
    Exit Function
End Function

Public Function ImprimeTrans(ByVal gc As GNComprobante, ByVal bandAsiento As Boolean) As String
    Dim crear As Boolean
    Static objImp As Object
    On Error GoTo ErrTrap

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
    If Me.tag = "Depre" Then
        objImp.PrintTrans gobjMain.EmpresaActual, True, 1, 0, "", 0, gc
    End If
    
    MensajeStatus "", 0
    ImprimeTrans = ""       'Sin problema
    Exit Function
ErrTrap:
    MensajeStatus "", 0
    Select Case Err.Number
    Case ERR_NOIMPRIME, ERR_NOIMPRIME2, ERR_NOIMPRIME3, ERR_NOHAYCODIGO
        ImprimeTrans = Err.Description
    Case Else
        ImprimeTrans = MSGERR_NOIMPRIME2
    End Select
    Exit Function
End Function

Private Function VerificaIngresoAutomatico() As String
End Function


Private Sub PoneColorFilas()
    Dim i As Long, j As Long, Elemento As String, k As Long, l As Long
    With grd
        If .Rows <= .FixedRows Then Exit Sub
        For i = 1 To .Rows - 1
            For j = 1 To COL_DEP_RESULTADO
                  If Not (grd.IsSubtotal(i)) Then
                    If .ColData(j) = -1 Then
                      .Cell(flexcpBackColor, i, j, i, j) = &H80000018 'vbYellow
                    End If
                  End If
            Next j
        Next i
    End With
End Sub



Private Sub cboGrupo_Click()
    Dim Numg As Integer
    On Error GoTo ErrTrap
    If cboGrupo.ListIndex < 0 Then Exit Sub

    'MensajeStatus MSG_PREPARA, vbHourglass

    Numg = cboGrupo.ListIndex + 1
    fcbGrupoDesde.SetData gobjMain.EmpresaActual.ListaAFGrupo(Numg, False, False)
    fcbGrupoDesde.KeyText = ""
    CargaItems
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    Exit Sub
End Sub
    
Private Sub CargaItems()
    Dim numGrupo As Integer, v() As Variant
    Dim sql  As String, rs As Recordset, cond As String
    numGrupo = cboGrupo.ListIndex + 1
    fcbDesde2.Clear
    If Len(fcbGrupoDesde.Text) > 0 Then
        cond = " WHERE codGrupo" & numGrupo & " = '" & _
                fcbGrupoDesde.Text & "'"
    End If
    sql = "SELECT CodInventario, AFInventario.Descripcion FROM AFInventario " & _
    IIf(Len(fcbGrupoDesde.Text) > 0, " INNER JOIN AFGrupo" & numGrupo & _
           " ON AFInventario.IdGrupo" & numGrupo & " = AFGrupo" & numGrupo & ".IdGrupo" & numGrupo & cond, "")
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not rs.EOF Then
        v = MiGetRows(rs)
        fcbDesde2.SetData v
    End If
    fcbDesde2.Text = ""
End Sub


Private Sub fcbGrupoDesde_LostFocus()
    CargaItems
End Sub


Private Function DepreciacionAuto(ByVal bandVerificar As Boolean, BandTodo As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, x As Single, j As Integer, filaSubTotal As Long
    Dim gnc As GNComprobante, cambiado As Boolean, TransGen As String
    Dim rs As Recordset, sql As String
    
    On Error GoTo ErrTrap
    
    'Si no es solo verificacion, confirma
    If Not bandVerificar Then
        'Confirma la actualización
        s = "Este proceso creará Depreciaciones Automáticos  de los Activos Fijos seleccionadaos" & vbCr & vbCr
        s = s & "Está seguro que desea proceder?"
        If MsgBox(s, vbYesNo + vbQuestion) <> vbYes Then Exit Function
    End If
   
    s = ""
    
    Set mColItems = Nothing     'Limpia lo anterior
    Set mColItems = New Collection
    
    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdAceptar.Enabled = False
    cmdBuscar.Enabled = False
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    
    For i = grd.FixedRows To grd.Rows - 1
        DoEvents
        If mCancelado Then
            MsgBox "El proceso fue cancelado.", vbInformation
            Exit For
        End If
        
        prg1.value = i
        grd.Row = i
        x = grd.CellTop                 'Para visualizar la celda actual
        
        If Not grd.IsSubtotal(i) Then
            If grd.ValueMatrix(i, COL_DEP_COSTOCOMPRA) <> 0 Then
            tid = grd.ValueMatrix(i, COL_DEP_ID)
            grd.TextMatrix(i, COL_DEP_RESULTADO) = "Procesando  ..."
            grd.Refresh
            
            'Recupera la transaccion
            
            Set mobjAF = gobjMain.EmpresaActual.RecuperaAFInventario(tid)
            If Not (mobjAF Is Nothing) Then
                'Si la transacción es de Inventario y es Egreso/Transferencia
                ' Y no está anulado
                If mobjAF.Estado <> 99 Then
'                    For j = i To grd.Rows - 1
'                        If grd.IsSubtotal(j) Then
'                            filaSubTotal = j
'                            j = grd.Rows - 1
'                        End If
'                    Next j
                    filaSubTotal = grd.Rows
                    If GrabarDepreAuto(TransGen, i, filaSubTotal) Then
                        'Graba la transacción
                        j = 1
                            grd.TextMatrix(j, COL_DEP_TIDIN) = mobjGNCompAux.TransID
                        If RegenerarAsiento(TransGen, i, filaSubTotal) Then
                        End If
                        For j = i To filaSubTotal - 1
                            
                            sql = " update afinventario set NumeroDepre=isnull(NumeroDepre,0) +1 where codinventario='" & grd.TextMatrix(j, COL_DEP_CODAFINV) & "'"
                            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                            
                            grd.Row = j
                            x = grd.CellTop                 'Para visualizar la celda actual
                            prg1.value = j

                            If InStr(1, grd.TextMatrix(j, COL_DEP_RESULTADO), "ERROR") = 0 Then
                                If j = filaSubTotal - 1 Then
                                    grd.TextMatrix(j, COL_DEP_RESULTADO) = "OK.. Trans " & TransGen
                                Else
                                    grd.TextMatrix(j, COL_DEP_RESULTADO) = "OK.. " '& TransGen
                                End If
                            End If
                            grd.TextMatrix(j, COL_DEP_TIDIN) = mobjGNCompAux.TransID
                            grd.Refresh
                            
                            
                            
                            
                        Next j
                        
                        i = filaSubTotal
                    Else
                            'Si no está cambiado no graba
                        grd.TextMatrix(i, COL_DEP_RESULTADO) = "Falló Proceso"
                    End If
                Else
                    'Si está anulado
                    If gnc.Estado = ESTADO_ANULADO Then
                        grd.TextMatrix(i, COL_DEP_RESULTADO) = "Anulado"
                    'Si no tiene nada que ver con recalculo de costo
                    Else
                        grd.TextMatrix(i, COL_DEP_RESULTADO) = "---"
                    End If
                End If
            Else
                grd.TextMatrix(i, COL_DEP_RESULTADO) = "No pudo recuperar la transación."
            End If
        Else
            grd.TextMatrix(i, COL_DEP_RESULTADO) = "ERROR... valor de compra igual a 0."
        End If
            
        End If
    Next i
    Screen.MousePointer = 0
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    If i < grd.Rows And i >= grd.FixedRows Then
        grd.TextMatrix(i, COL_DEP_RESULTADO) = Err.Description
    End If
    DispErr
    prg1.value = prg1.min
salida:
    Set mColItems = Nothing         'Libera el objeto de coleccion
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdBuscar.Enabled = True
    cmdAceptar.Enabled = True
    prg1.value = prg1.min
    Exit Function
End Function


Private Function GrabarDepreAuto(ByRef trans As String, FilaIni As Long, FilaFin As Long) As Boolean
    Dim Imprime As Boolean, i As Long, ix As Long, orden1 As Integer, orden2 As Integer, j As Integer
    Dim pc As PCProvCli, Cadena As String, obser As String, codforma As String, Num As Long
    Dim tsf As TSFormaCobroPago, x As Single, k As Long, m As Integer, NumDep As Integer
    Dim tid As Long, cont As Integer, NumDepReal As Integer
    Dim valor As Currency, ValorMensual As Currency, CostoDepTotal As Currency, ValorMensualAnterior As Currency
    Dim mes As String, anio As String
    Dim sumameses As Long


    On Error GoTo ErrTrap
    GrabarDepreAuto = True
    orden1 = 1
    orden2 = 1
    cont = 0
    If CreaComprobanteDepreciacionAuto(i) Then
        'Si es solo lectura, no hace nada
        If mobjGNCompAux.SoloVer Then
            MsgBox MSG_NODISPONE, vbInformation
            Exit Function
        End If
        'carga la nueva deuda a los bancos de las tarjetas
        i = 1
        For i = FilaIni To FilaFin - 1
            If (grd.ValueMatrix(i, COL_DEP_VIDAUTIL) <> grd.ValueMatrix(i, COL_DEP_DEPANTERIOR)) And grd.ValueMatrix(i, COL_DEP_COSTOCOMPRA) <> 0 Then
                NumDepReal = 0
                grd.Row = i
                prg1.value = i
                x = grd.CellTop
               Select Case grd.TextMatrix(i, COL_DEP_TIPODEPRE)
                    Case DEP_ACELERADA
                        NumDepReal = grd.ValueMatrix(i, COL_DEP_DEPANTERIOR) + gobjMain.EmpresaActual.ObtieneNumeroDepreciacion(grd.TextMatrix(i, COL_DEP_TIPODEPRE), grd.TextMatrix(i, COL_DEP_CODAFINV), CostoDepTotal)
                       
                        
                        
                        sumameses = 0
                        For m = 1 To (grd.ValueMatrix(i, COL_DEP_VIDAUTIL))
                        'For m = 1 To NumDepReal
                            sumameses = sumameses + m
                        Next m
                        ValorMensual = (grd.ValueMatrix(i, COL_DEP_COSTOCOMPRA) - grd.ValueMatrix(i, COL_DEP_COSTORESIDUAL)) / ((sumameses)) '* (DateDiff("m", grd.TextMatrix(i, COL_DEP_FECHA), Date) + 1)
    '                    NumDepReal = gobjMain.EmpresaActual.ObtieneNumeroDepreciacion(grd.TextMatrix(i, COL_DEP_TIPODEPRE), grd.TextMatrix(i, COL_DEP_CODAFINV), CostoDepTotal)
                    Case DEP_DESACELERADA
                    Case DEP_LINEAL
    
                        ValorMensual = (grd.ValueMatrix(i, COL_DEP_COSTOCOMPRA) - grd.ValueMatrix(i, COL_DEP_COSTORESIDUAL)) / (grd.ValueMatrix(i, COL_DEP_VIDAUTIL)) '- grd.ValueMatrix(i, COL_DEP_DEPANTERIOR))
                        NumDepReal = gobjMain.EmpresaActual.ObtieneNumeroDepreciacion(grd.TextMatrix(i, COL_DEP_TIPODEPRE), grd.TextMatrix(i, COL_DEP_CODAFINV), CostoDepTotal)
                    Case Else
                        ValorMensual = (grd.ValueMatrix(i, COL_DEP_COSTOCOMPRA) - grd.ValueMatrix(i, COL_DEP_COSTORESIDUAL)) / (grd.ValueMatrix(i, COL_DEP_VIDAUTIL) - grd.ValueMatrix(i, COL_DEP_DEPANTERIOR))
                End Select
                For NumDep = 1 To ntxNumPer.value
                    cont = cont + 1
                    If grd.TextMatrix(i, COL_DEP_TIPODEPRE) = "0" Then
                        valor = ValorMensual * (NumDepReal + NumDep)
                    Else
                        valor = ValorMensual
                    End If
                    If (NumDepReal + NumDep) <= grd.ValueMatrix(i, COL_DEP_VIDAUTIL) Then
                        ix = mobjGNCompAux.AddAFKardex
                        mobjGNCompAux.AFKardex(ix).CodInventario = grd.TextMatrix(i, COL_DEP_CODAFINV)
                        mobjGNCompAux.AFKardex(ix).CodBodega = mobjGNCompAux.GNTrans.CodBodegaPre ' "A01" 'mobjGNCompAux.GNTrans.IdBodegaPre '="A01"
                        mobjGNCompAux.AFKardex(ix).CostoTotal = valor * -1
                        mobjGNCompAux.AFKardex(ix).CostoRealTotal = valor * -1
                        mobjGNCompAux.AFKardex(ix).cantidad = -1
                        mobjGNCompAux.AFKardex(ix).orden = cont
                        mes = mobjGNCompAux.Empresa.DevuelveMes(DatePart("m", DateAdd("m", NumDepReal + NumDep - 1, grd.TextMatrix(i, COL_DEP_FECHA))), True)
                        'mobjGNCompAux.AFKardex(ix).Nota = " Depreciación No. " & NumDepReal + NumDep & " del mes de " & mes & "/" & DatePart("yyyy", DateAdd("m", NumDepReal + NumDep - 1, grd.TextMatrix(i, COL_DEP_FECHA)))
                    Else
                        MsgBox "El Número de Depreciaciones excedió al la vida util del Activo: " & grd.ValueMatrix(i, COL_DEP_VIDAUTIL)
                    End If
                Next NumDep
                grd.TextMatrix(i, COL_DEP_RESULTADO) = "Procesando  ... "
            Else
                If grd.ValueMatrix(i, COL_DEP_COSTOCOMPRA) = 0 Then
                    grd.TextMatrix(i, COL_DEP_RESULTADO) = "ERROR... valor de compra igual a 0."
                Else
                    grd.TextMatrix(i, COL_DEP_RESULTADO) = "ERROR Activo ya no se puede depreciar  ... "
                End If
            End If
            grd.Refresh
            
        Next i
            mobjGNCompAux.FechaTrans = dtpFecha1.value
            mobjGNCompAux.HoraTrans = Time()
            Cadena = ""
            If Len(Cadena) > 120 Then
                mobjGNCompAux.Descripcion = Mid$(Cadena, 1, 120)
            Else
                mobjGNCompAux.Descripcion = Cadena
            End If
                
            mobjGNCompAux.Estado = ESTADO_NOAPROBADO
            'Si es que algo está modificado
            If mobjGNCompAux.Modificado Then
                MensajeStatus MSG_GENERANDOASIENTO, vbHourglass
                MensajeStatus
            End If
            'Verificación de datos
            mobjGNCompAux.VerificaDatos
        
            'Verifica si está cuadrado el asiento
        
            MensajeStatus MSG_GRABANDO, vbHourglass
        
            'Manda a grabar
            '       Aquí ya no hacemos verificación de asiento por que ya está hecho en Control Asiento
            mobjGNCompAux.Grabar False, True
        '***  Oliver 26/12/2002
        'Agregado para el control ded Impresion Configurado en la Transaccion
        grd.Refresh
        
        MensajeStatus
    '    Me.caption = "Transacción " & mobjGNCompAux.codTrans & " " & mobjGNCompAux.NumTrans
        Me.Caption = mobjGNCompAux.CodTrans & " " & mobjGNCompAux.numtrans
        trans = mobjGNCompAux.CodTrans & " " & mobjGNCompAux.numtrans
        GrabarDepreAuto = True
    Else
        GrabarDepreAuto = False
    End If
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
    GrabarDepreAuto = False
    Exit Function
    
End Function


Private Function CreaComprobanteDepreciacionAuto(ByRef Num As Long) As Boolean
    Dim v As Currency, tsf As TSFormaCobroPago
    Dim i As Long
    CreaComprobanteDepreciacionAuto = False
    If Len(fcbTrans.KeyText) > 0 Then
        Set mobjGNCompAux = gobjMain.EmpresaActual.CreaGNComprobante(fcbTrans.KeyText)
    Else
        MsgBox "Falta seleccionar transacción de Depreciación"
        CreaComprobanteDepreciacionAuto = False
        Exit Function
    End If
    CreaComprobanteDepreciacionAuto = True
End Function

    
Public Sub InicioReval(Name As String)
    Dim i As Integer
    On Error GoTo ErrTrap
        Me.tag = Name
        For i = 1 To AFGRUPO_MAX
            cboGrupo.AddItem gobjMain.EmpresaActual.GNOpcion.EtiqAFGrupo(i)
        Next i
        If (numGrupo <= cboGrupo.ListCount) And (numGrupo > 0) Then
            cboGrupo.ListIndex = numGrupo - 1   'Selecciona lo anterior
        ElseIf cboGrupo.ListCount > 0 Then
            cboGrupo.ListIndex = 0              'Selecciona la primera
        End If
        numGrupo = cboGrupo.ListIndex + 1
        ConfigCols
        Me.Show
        Me.ZOrder
        dtpFecha1.value = Date
        CargaTrans
        Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Function DepreciacionAutoReval(ByVal bandVerificar As Boolean, BandTodo As Boolean, BadAumento As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, x As Single, j As Integer, filaSubTotal As Long
    Dim gnc As GNComprobante, cambiado As Boolean, TransGen As String
    Dim rs As Recordset, sql As String
    
    On Error GoTo ErrTrap
    
    'Si no es solo verificacion, confirma
    If Not bandVerificar Then
        'Confirma la actualización
        s = "Este proceso creará Depreciaciones Revalorizadas Automáticos  de los Activos Fijos seleccionadaos" & vbCr & vbCr
        s = s & "Está seguro que desea proceder?"
        If MsgBox(s, vbYesNo + vbQuestion) <> vbYes Then Exit Function
    End If
   
    s = ""
    
    Set mColItems = Nothing     'Limpia lo anterior
    Set mColItems = New Collection
    
    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdAceptar.Enabled = False
    cmdBuscar.Enabled = False
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    
    For i = grd.FixedRows To grd.Rows - 1
        DoEvents
        If mCancelado Then
            MsgBox "El proceso fue cancelado.", vbInformation
            Exit For
        End If
        
        prg1.value = i
        grd.Row = i
        x = grd.CellTop                 'Para visualizar la celda actual
        
        If Not grd.IsSubtotal(i) Then
            If grd.ValueMatrix(i, COL_DEP_COSTOCOMPRA) <> 0 Then
            tid = grd.ValueMatrix(i, COL_DEP_ID)
            grd.TextMatrix(i, COL_DEP_RESULTADO) = "Procesando  ..."
            grd.Refresh
            
            'Recupera la transaccion
            
            Set mobjAF = gobjMain.EmpresaActual.RecuperaAFInventario(tid)
            If Not (mobjAF Is Nothing) Then
                'Si la transacción es de Inventario y es Egreso/Transferencia
                ' Y no está anulado
                If mobjAF.BandValida And Not mobjAF.BandServicio Then
                    filaSubTotal = grd.Rows
                    If GrabarDepreAutoReval(TransGen, i, filaSubTotal, BadAumento) Then
                        'Graba la transacción
                        j = 1
                            grd.TextMatrix(j, COL_DEP_TIDIN) = mobjGNCompAux.TransID
                        If RegenerarAsiento(TransGen, i, filaSubTotal) Then
                        End If
                        For j = i To filaSubTotal - 1
                            
''                            sql = " update afinventario set NumeroDepre=isnull(NumeroDepre,0) +1 where codinventario='" & grd.TextMatrix(j, COL_DEP_CODAFINV) & "'"
''                            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                            
                            grd.Row = j
                            x = grd.CellTop                 'Para visualizar la celda actual
                            prg1.value = j

                            If InStr(1, grd.TextMatrix(j, COL_DEP_RESULTADO), "ERROR") = 0 Then
                                If j = filaSubTotal - 1 Then
                                    grd.TextMatrix(j, COL_DEP_RESULTADO) = "OK.. Trans " & TransGen
                                Else
                                    grd.TextMatrix(j, COL_DEP_RESULTADO) = "OK.. " '& TransGen
                                End If
                            End If
                            grd.TextMatrix(j, COL_DEP_TIDIN) = mobjGNCompAux.TransID
                            grd.Refresh
                            
                            
                            
                            
                        Next j
                        
                        i = filaSubTotal
                    Else
                            'Si no está cambiado no graba
                        grd.TextMatrix(i, COL_DEP_RESULTADO) = "Falló Proceso"
                    End If
                Else
                    'Si está anulado
                    If gnc.Estado = ESTADO_ANULADO Then
                        grd.TextMatrix(i, COL_DEP_RESULTADO) = "Anulado"
                    'Si no tiene nada que ver con recalculo de costo
                    Else
                        grd.TextMatrix(i, COL_DEP_RESULTADO) = "---"
                    End If
                End If
            Else
                grd.TextMatrix(i, COL_DEP_RESULTADO) = "No pudo recuperar la transación."
            End If
        Else
            grd.TextMatrix(i, COL_DEP_RESULTADO) = "ERROR... valor de compra igual a 0."
        End If
            
        End If
    Next i
    Screen.MousePointer = 0
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    If i < grd.Rows And i >= grd.FixedRows Then
        grd.TextMatrix(i, COL_DEP_RESULTADO) = Err.Description
    End If
    DispErr
    prg1.value = prg1.min
salida:
    Set mColItems = Nothing         'Libera el objeto de coleccion
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdBuscar.Enabled = True
    cmdAceptar.Enabled = True
    prg1.value = prg1.min
    Exit Function
End Function

Private Function GrabarDepreAutoReval(ByRef trans As String, FilaIni As Long, FilaFin As Long, BandAumento As Boolean) As Boolean
    Dim Imprime As Boolean, i As Long, ix As Long, orden1 As Integer, orden2 As Integer, j As Integer
    Dim pc As PCProvCli, Cadena As String, obser As String, codforma As String, Num As Long
    Dim tsf As TSFormaCobroPago, x As Single, k As Long, m As Integer, NumDep As Integer
    Dim tid As Long, cont As Integer, NumDepReal As Integer
    Dim valor As Currency, ValorMensual As Currency, CostoDepTotal As Currency, ValorMensualAnterior As Currency
    Dim mes As String, anio As String
    Dim sumameses As Long
    Dim bandDep As Boolean, inidepre As Date

    On Error GoTo ErrTrap
    GrabarDepreAutoReval = True
    bandDep = False
    orden1 = 1
    orden2 = 1
    cont = 0
    
    
    For i = FilaIni To FilaFin - 1
        mes = DatePart("m", CDate(grd.TextMatrix(i, COL_DEP_FECHAINIDEP)))
        anio = DatePart("yyyy", CDate(grd.TextMatrix(i, COL_DEP_FECHAINIDEP)))
        inidepre = CDate("01/" & mes & "/" & anio)
        If (grd.ValueMatrix(i, COL_DEP_VIDAUTIL) <> grd.ValueMatrix(i, COL_DEP_DEPANTERIOR)) And grd.ValueMatrix(i, COL_DEP_COSTOCOMPRA) <> 0 And inidepre <= dtpFecha1.value Then
            bandDep = True
            Exit For
        End If
    Next i
    
    If bandDep = False Then
            MsgBox "No existe ningun detalle en la depreciación"
            Exit Function
        
    End If
    
    If CreaComprobanteDepreciacionAuto(i) Then
        'Si es solo lectura, no hace nada
        If mobjGNCompAux.SoloVer Then
            MsgBox MSG_NODISPONE, vbInformation
            Exit Function
        End If
        'carga la nueva deuda a los bancos de las tarjetas
        i = 1
        For i = FilaIni To FilaFin - 1
                mes = DatePart("m", CDate(grd.TextMatrix(i, COL_DEP_FECHAINIDEP)))
                anio = DatePart("yyyy", CDate(grd.TextMatrix(i, COL_DEP_FECHAINIDEP)))
                inidepre = CDate("01/" & mes & "/" & anio)

            If (grd.ValueMatrix(i, COL_DEP_VIDAUTIL) <> grd.ValueMatrix(i, COL_DEP_DEPANTERIOR)) And grd.ValueMatrix(i, COL_DEP_COSTOCOMPRA) <> 0 And inidepre <= dtpFecha1.value Then
                NumDepReal = 0
                grd.Row = i
                prg1.value = i
                x = grd.CellTop
               Select Case grd.TextMatrix(i, COL_DEP_TIPODEPRE)
                    Case DEP_ACELERADA
                        NumDepReal = grd.ValueMatrix(i, COL_DEP_DEPANTERIOR) + gobjMain.EmpresaActual.ObtieneNumeroDepreciacion(grd.TextMatrix(i, COL_DEP_TIPODEPRE), grd.TextMatrix(i, COL_DEP_CODAFINV), CostoDepTotal)
                       
                        
                        
                        sumameses = 0
                        For m = 1 To (grd.ValueMatrix(i, COL_DEP_VIDAUTIL))
                        'For m = 1 To NumDepReal
                            sumameses = sumameses + m
                        Next m
                        ValorMensual = (grd.ValueMatrix(i, COL_DEP_COSTOCOMPRA)) / ((sumameses))  '* (DateDiff("m", grd.TextMatrix(i, COL_DEP_FECHA), Date) + 1)
    '                    NumDepReal = gobjMain.EmpresaActual.ObtieneNumeroDepreciacion(grd.TextMatrix(i, COL_DEP_TIPODEPRE), grd.TextMatrix(i, COL_DEP_CODAFINV), CostoDepTotal)
                    Case DEP_DESACELERADA
                    Case DEP_LINEAL
    
                        ValorMensual = (grd.ValueMatrix(i, COL_DEP_COSTOCOMPRA)) / (grd.ValueMatrix(i, COL_DEP_VIDAUTIL))  '- grd.ValueMatrix(i, COL_DEP_DEPANTERIOR))
                        NumDepReal = gobjMain.EmpresaActual.ObtieneNumeroDepreciacionReval(grd.TextMatrix(i, COL_DEP_TIPODEPRE), grd.TextMatrix(i, COL_DEP_CODAFINV), CostoDepTotal, fcbTrans.KeyText, BandAumento)
                    Case Else
                        ValorMensual = (grd.ValueMatrix(i, COL_DEP_COSTOCOMPRA)) / (grd.ValueMatrix(i, COL_DEP_VIDAUTIL) - grd.ValueMatrix(i, COL_DEP_DEPANTERIOR))
                End Select
                For NumDep = 1 To ntxNumPer.value
                    cont = cont + 1
                    If grd.TextMatrix(i, COL_DEP_TIPODEPRE) = "0" Then
                        valor = ValorMensual * (NumDepReal + NumDep)
                    Else
                        valor = ValorMensual
                    End If
                    If Not BandAumento Then
                    
                        If (NumDepReal + NumDep) <= grd.ValueMatrix(i, COL_DEP_VIDAUTIL) Then
                            ix = mobjGNCompAux.AddAFKardex
                            mobjGNCompAux.AFKardex(ix).CodInventario = grd.TextMatrix(i, COL_DEP_CODAFINV)
                            mobjGNCompAux.AFKardex(ix).CodBodega = mobjGNCompAux.GNTrans.CodBodegaPre ' "A01" 'mobjGNCompAux.GNTrans.IdBodegaPre '="A01"
                            mobjGNCompAux.AFKardex(ix).CostoTotal = valor * -1
                            mobjGNCompAux.AFKardex(ix).CostoRealTotal = valor * -1
                            mobjGNCompAux.AFKardex(ix).cantidad = -1
                            mobjGNCompAux.AFKardex(ix).orden = cont
                            mobjGNCompAux.AFKardex(ix).NumeroPrecio = grd.ValueMatrix(i, COL_DEP_COSTORESIDUAL)
                            mobjGNCompAux.AFKardex(ix).NumDepreReval = grd.ValueMatrix(i, COL_DEP_VIDAUTIL) * -1
                        Else
                            MsgBox "El Número de Depreciaciones excedió al la vida util del Activo: " & grd.ValueMatrix(i, COL_DEP_VIDAUTIL)
                        End If
                    Else
                        If (NumDepReal + NumDep) <= grd.ValueMatrix(i, COL_DEP_VIDAUTIL) Then
                            ix = mobjGNCompAux.AddAFKardex
                            mobjGNCompAux.AFKardex(ix).CodInventario = grd.TextMatrix(i, COL_DEP_CODAFINV)
                            mobjGNCompAux.AFKardex(ix).CodBodega = mobjGNCompAux.GNTrans.CodBodegaPre ' "A01" 'mobjGNCompAux.GNTrans.IdBodegaPre '="A01"
                            mobjGNCompAux.AFKardex(ix).CostoTotal = valor
                            mobjGNCompAux.AFKardex(ix).CostoRealTotal = valor
                            mobjGNCompAux.AFKardex(ix).cantidad = 1
                            mobjGNCompAux.AFKardex(ix).orden = cont
                            mobjGNCompAux.AFKardex(ix).NumeroPrecio = grd.ValueMatrix(i, COL_DEP_COSTORESIDUAL)
                            mobjGNCompAux.AFKardex(ix).NumDepreReval = grd.ValueMatrix(i, COL_DEP_VIDAUTIL) * -1
                            
'                            mes = mobjGNCompAux.Empresa.DevuelveMes(DatePart("m", DateAdd("m", NumDepReal + NumDep - 1, grd.TextMatrix(i, COL_DEP_FECHA))), True)
                            'mobjGNCompAux.AFKardex(ix).Nota = " Depreciación No. " & NumDepReal + NumDep & " del mes de " & mes & "/" & DatePart("yyyy", DateAdd("m", NumDepReal + NumDep - 1, grd.TextMatrix(i, COL_DEP_FECHA)))
                        Else
                                MsgBox "El Número de Depreciaciones excedió al la Revalorización del Activo: " & grd.ValueMatrix(i, COL_DEP_VIDAUTIL)
                        End If
                    
                    End If
                    mobjGNCompAux.AFKardex(ix).CodInventario = grd.TextMatrix(i, COL_DEP_CODAFINV)
                Next NumDep
                grd.TextMatrix(i, COL_DEP_RESULTADO) = "Procesando  ... "
                Else
                If grd.ValueMatrix(i, COL_DEP_COSTOCOMPRA) = 0 Then
                    grd.TextMatrix(i, COL_DEP_RESULTADO) = "ERROR... valor de compra igual a 0."
                ElseIf inidepre > dtpFecha1.value Then
                    grd.TextMatrix(i, COL_DEP_RESULTADO) = "NO se deprecia no cumple la fecha"
                Else
                    grd.TextMatrix(i, COL_DEP_RESULTADO) = "ERROR Activo ya no se puede depreciar  ... "
                End If
            End If
            grd.Refresh
            
        Next i
            mobjGNCompAux.FechaTrans = dtpFecha1.value
            mobjGNCompAux.HoraTrans = Time()
            Cadena = ""
            If Len(Cadena) > 120 Then
                mobjGNCompAux.Descripcion = Mid$(Cadena, 1, 120)
            Else
                mobjGNCompAux.Descripcion = Cadena
            End If
                
            mobjGNCompAux.Estado = ESTADO_NOAPROBADO
            'Si es que algo está modificado
            If mobjGNCompAux.Modificado Then
                MensajeStatus MSG_GENERANDOASIENTO, vbHourglass
                MensajeStatus
            End If
            'Verificación de datos
            mobjGNCompAux.VerificaDatos
        
            'Verifica si está cuadrado el asiento
        
            MensajeStatus MSG_GRABANDO, vbHourglass
        
            'Manda a grabar
            '       Aquí ya no hacemos verificación de asiento por que ya está hecho en Control Asiento
            mobjGNCompAux.Grabar False, True
        '***  Oliver 26/12/2002
        'Agregado para el control ded Impresion Configurado en la Transaccion
        grd.Refresh
        
        MensajeStatus
    '    Me.caption = "Transacción " & mobjGNCompAux.codTrans & " " & mobjGNCompAux.NumTrans
        Me.Caption = mobjGNCompAux.CodTrans & " " & mobjGNCompAux.numtrans
        trans = mobjGNCompAux.CodTrans & " " & mobjGNCompAux.numtrans
        GrabarDepreAutoReval = True
    Else
        GrabarDepreAutoReval = False
    End If
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
    GrabarDepreAutoReval = False
    Exit Function
    
End Function

