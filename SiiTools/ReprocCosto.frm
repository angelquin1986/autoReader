VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl1.ocx"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "flexcombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReprocCosto 
   Caption         =   "Reprocesamiento de costos"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5325
   ScaleWidth      =   6810
   WindowState     =   2  'Maximized
   Begin VB.Frame fraprov 
      Caption         =   "Proveedor"
      Height          =   675
      Left            =   14670
      TabIndex        =   21
      Top             =   90
      Visible         =   0   'False
      Width           =   5052
      Begin FlexComboProy.FlexCombo fcbDesde 
         Height          =   315
         Left            =   840
         TabIndex        =   22
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
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
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   252
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   612
      End
   End
   Begin VB.CheckBox chkTodo 
      Caption         =   "&Regenerar todo sin verificar"
      Enabled         =   0   'False
      Height          =   192
      Left            =   4140
      TabIndex        =   18
      Top             =   1800
      Width           =   3252
   End
   Begin VB.Frame fraFecha 
      Caption         =   "&Fecha (desde - hasta)"
      Height          =   1572
      Left            =   168
      TabIndex        =   0
      Top             =   120
      Width           =   2052
      Begin MSComCtl2.DTPicker dtpFecha1 
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1692
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
         Format          =   102825985
         CurrentDate     =   36348
      End
      Begin MSComCtl2.DTPicker dtpFecha2 
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1692
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
         Format          =   102825985
         CurrentDate     =   36348
      End
   End
   Begin VB.Frame fraCodTrans 
      Caption         =   "Cod.&Trans"
      Height          =   1572
      Left            =   2088
      TabIndex        =   3
      Top             =   120
      Width           =   10485
      Begin VB.CommandButton cmdActualizaTrans 
         Caption         =   "Act. Trans"
         Height          =   330
         Left            =   9180
         TabIndex        =   20
         Top             =   570
         Width           =   1215
      End
      Begin VB.CommandButton cmdtranselec 
         Caption         =   "Trans. Selec"
         Height          =   330
         Left            =   9180
         TabIndex        =   19
         Top             =   210
         Width           =   1215
      End
      Begin VB.CommandButton cmdTransLimpiar 
         Caption         =   "Limp."
         Height          =   330
         Left            =   1740
         TabIndex        =   16
         Top             =   1116
         Width           =   732
      End
      Begin VB.CommandButton cmdTransTodo 
         Caption         =   "Todo egresos"
         Height          =   330
         Left            =   360
         TabIndex        =   15
         Top             =   1116
         Width           =   1155
      End
      Begin VB.ListBox lstTrans 
         Columns         =   10
         Height          =   852
         IntegralHeight  =   0   'False
         Left            =   240
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   14
         Top             =   240
         Width           =   8895
      End
   End
   Begin VB.Frame fraNumTrans 
      Caption         =   "# T&rans. (desde - hasta)"
      Height          =   1572
      Left            =   12600
      TabIndex        =   4
      Top             =   90
      Width           =   2052
      Begin VB.TextBox txtNumTrans1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   360
         TabIndex        =   5
         Top             =   280
         Width           =   1212
      End
      Begin VB.TextBox txtNumTrans2 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   360
         TabIndex        =   6
         Top             =   640
         Width           =   1212
      End
   End
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   852
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   6810
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4470
      Width           =   6810
      Begin VB.CommandButton cmdCorregirIVA 
         Caption         =   "Verificar IVA Items"
         Enabled         =   0   'False
         Height          =   372
         Left            =   3000
         TabIndex        =   17
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Proceder"
         Enabled         =   0   'False
         Height          =   372
         Left            =   1605
         TabIndex        =   12
         Top             =   0
         Width           =   1212
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   4995
         TabIndex        =   11
         Top             =   0
         Width           =   1212
      End
      Begin VB.CommandButton cmdVerificar 
         Caption         =   "&Verificar"
         Enabled         =   0   'False
         Height          =   372
         Left            =   255
         TabIndex        =   10
         Top             =   0
         Width           =   1212
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
      Height          =   1935
      Left            =   120
      TabIndex        =   8
      Top             =   2220
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
      Left            =   2520
      TabIndex        =   7
      Top             =   1740
      Width           =   1452
   End
End
Attribute VB_Name = "frmReprocCosto"
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
Private Const COL_NUMDOCREF = 6     '*** MAKOTO 07/feb/01 Agregado
Private Const COL_NOMBRE = 7        '*** MAKOTO 07/feb/01 Agregado
Private Const COL_DESC = 8
Private Const COL_CENTROCOSTO = 9
Private Const COL_ESTADO = 10
Private Const COL_RESULTADO = 11

Private Const MSG_NG = "Costo incorrecto."
Private mProcesando As Boolean
Private mCancelado As Boolean
Private mVerificado As Boolean
'*** MAKOTO 31/ago/00 Agregado
'       para almacenar items con costo incorrecto detectado
Private mColItems As Collection
Private mobjGNCompAux As GNComprobante

Public Sub Inicio()
    Dim i As Integer
    On Error GoTo ErrTrap
    Me.Show
    Me.ZOrder
    dtpFecha1.value = gobjMain.EmpresaActual.GNOpcion.FechaLimiteDesde '.FechaInicio
    dtpFecha2.value = Date
    CargaTrans
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

'*** MAKOTO 31/ago/00 Modificado
Private Sub CargaTrans()
    Dim i As Long, v As Variant
    Dim s As String, cod As String, aux  As Integer, gt As GNTrans
    'Carga la lista de transacción
'    fcbTrans.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "IV")

    lstTrans.Clear
    v = gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "IV")
    For i = LBound(v, 2) To UBound(v, 2)
        lstTrans.AddItem v(0, i)        '& " " & v(1, i)
    Next i
    
    'jeaa 25/09/206
''        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaRecosteo")) > 0 Then
''            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaRecosteo")
''            RecuperaTrans "KeyT", lstTrans, s
''        End If
    
If Me.tag = "CostoxProveedor" Then
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaRecosteoxProveedor")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaRecosteoxProveedor")
            RecuperaTrans "KeyT", lstTrans, s
        End If

Else
    aux = lstTrans.ListIndex
    For i = 0 To lstTrans.ListCount - 1
        cod = lstTrans.List(i)
        Set gt = gobjMain.EmpresaActual.RecuperaGNTrans(cod)
        If Not (gt Is Nothing) Then
            'Solo marca egresos/transferencias
            If gt.IVReprocesaCosto Then
                lstTrans.Selected(i) = True
            End If
        End If
    Next i
End If
    
    
End Sub

Private Sub chkTodo_Click()
    If chkTodo.value = vbChecked Then
        cmdVerificar.Enabled = False
        cmdAceptar.Enabled = (grd.Rows > grd.FixedRows)
    Else
        cmdVerificar.Enabled = Not mVerificado
        cmdAceptar.Enabled = mVerificado
    End If
End Sub

Private Sub cmdAceptar_Click()
    'Si no hay transacciones
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna transacción para procesar.", vbExclamation
        Exit Sub
    End If
    If dtpFecha1 < gobjMain.EmpresaActual.GNOpcion.FechaLimiteDesde Then
        MsgBox "La Rango de Fecha de reproceso es menor a la Fecha Limite Aceptable  ", vbExclamation
        Exit Sub
    End If
    
    If Me.tag = "CostoxProveedor" Then
        If ReprocCostoxProv(False, (chkTodo.value = vbChecked)) Then
            cmdCancelar.SetFocus
        End If
    
    Else
            If ReprocCosto(False, (chkTodo.value = vbChecked)) Then
                cmdCancelar.SetFocus
            End If
        
    End If

End Sub

Private Function VerificaIngreso() As String
    Dim i As Long, cod As String, gnt As GNTrans
    Dim s As String
    
    For i = 0 To lstTrans.ListCount - 1
        'Si está seleccionado
        If lstTrans.Selected(i) Then
            'Recupera el objeto GNTrans
            cod = lstTrans.List(i)
            Set gnt = gobjMain.EmpresaActual.RecuperaGNTrans(cod)
            'Si la transaccion es de ingreso, devuelve el codigo
            If gnt.IVTipoTrans = "I" Then s = s & cod & ", "
        End If
    Next i
    Set gnt = Nothing
    If Len(s) > 2 Then s = Left$(s, Len(s) - 2)     'Quita la ultima ", "
    VerificaIngreso = s
End Function

Private Function ReprocCosto(ByVal bandVerificar As Boolean, BandTodo As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, X As Single
    Dim gnc As GNComprobante, cambiado As Boolean
    Dim FechaAnt As Date, UsuarioAnt As String, UsuarioModAnt As String
    Dim sql As String, NumReg As Long, TransID As Long
    
    On Error GoTo ErrTrap
    
    'Si no es solo verificacion, confirma
    If Not bandVerificar Then
        'Confirma la actualización
        s = "Este proceso modificará los costos de la transacción seleccionada." & vbCr & vbCr
        s = s & "Está seguro que desea proceder?"
        If MsgBox(s, vbYesNo + vbQuestion) <> vbYes Then Exit Function
    End If
    
    'Verifica si está seleccionado una trans. de ingreso
    s = VerificaIngreso
    If Len(s) > 0 Then
        'Si está seleccinada, confirma si está seguro
        s = "Está seleccionada una o más transacciones de ingreso. " & vbCr & _
            "(" & s & ")" & vbCr & _
            "Generalmente no se hace reprocesamiento de costo con transacciones de ingreso." & vbCr & vbCr
        s = s & "Confirma que desea proceder?" & vbCr & _
            "Aplaste 'Sí' unicamente cuando está seguro de lo que está haciendo."
        If MsgBox(s, vbYesNo + vbQuestion + vbDefaultButton2) <> vbYes Then Exit Function
    End If
    s = ""
    
    Set mColItems = Nothing     'Limpia lo anterior
    Set mColItems = New Collection
    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdVerificar.Enabled = False
    cmdCorregirIVA.Enabled = False
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
        X = grd.CellTop                 'Para visualizar la celda actual
                'Si es verificación procesa todas las filas sino solo las que tengan "Costo Incorrecto"
        If ((grd.TextMatrix(i, COL_RESULTADO) = MSG_NG) Or bandVerificar Or BandTodo) Then
            tid = grd.ValueMatrix(i, COL_TID)
            grd.TextMatrix(i, COL_RESULTADO) = "Verificando..."
            grd.Refresh
            'Recupera la transaccion
            Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
            
'            If gnc.numtrans = 50 Then MsgBox "HOLA"
            If Not (gnc Is Nothing) Then
                'Si la transacción es de Inventario y es Egreso/Transferencia
                ' Y no está anulado
                If (gnc.GNTrans.Modulo = "IV") And _
                   (gnc.Estado <> ESTADO_ANULADO) Then
'                   (gnc.GNTrans.IVTipoTrans = "E" Or gnc.GNTrans.IVTipoTrans = "T") And _      '*** MAKOTO 06/sep/00 Eliminado
                    'Forzar recuperar todos los datos de transacción para que no se pierdan al grabar de nuveo
                    gnc.RecuperaDetalleTodo
                    If gnc.GNTrans.CodPantalla = "IVBQDISOCT" Then
                        If RecalculoTicketISO(gnc, cambiado, bandVerificar) Then
                            'Si está cambiado algo
                            If cambiado Or BandTodo Then
                                'Si no es solo verificacion
                                If Not bandVerificar Then
                                    FechaAnt = gnc.FechaGrabado
                                    UsuarioAnt = gnc.codUsuario
                                    UsuarioModAnt = gnc.codUsuarioModifica
                                    grd.TextMatrix(i, COL_RESULTADO) = "Grabando..."
                                    grd.Refresh
                                    'Prorratea los recargos/descuentos si los calcula en base a costo
                                    gnc.ProrratearIVKardexRecargo
                                    gnc.GeneraAsiento       'Diego 27 Abril 2001  corregido
                                    'Graba la transacción
                                    gnc.BandReproCostos = True
                                    gnc.Grabar False, False
                                    grd.TextMatrix(i, COL_RESULTADO) = "Actualizado."
                                    sql = " Update GNComprobante"
                                    sql = sql & " set "
                                    sql = sql & " CodUsuario = '" & UsuarioAnt & "',"
                                    sql = sql & " CodUsuarioModifica = '" & UsuarioModAnt & "',"
                                    sql = sql & " FechaGrabado = '" & FechaAnt & "'"
                                    sql = sql & " where transid =" & gnc.TransID
                                    gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                                'Si es solo verificacion
                                Else
                                    grd.TextMatrix(i, COL_RESULTADO) = MSG_NG
                                End If
                            Else
                                'Si no está cambiado no graba
                                grd.TextMatrix(i, COL_RESULTADO) = "OK."
                            End If
                        Else
                            grd.TextMatrix(i, COL_RESULTADO) = "Falló al recalcular."
                        End If
                    ElseIf gnc.GNTrans.CodPantalla = "IVCAM" Then
                        If RecalculoMega(gnc, cambiado, bandVerificar) Then
                            'Si está cambiado algo
                            If cambiado Or BandTodo Then
                                'Si no es solo verificacion
                                If Not bandVerificar Then
                                    FechaAnt = gnc.FechaGrabado
                                    UsuarioAnt = gnc.codUsuario
                                    UsuarioModAnt = gnc.codUsuarioModifica
                                    grd.TextMatrix(i, COL_RESULTADO) = "Grabando..."
                                    grd.Refresh
                                    'Prorratea los recargos/descuentos si los calcula en base a costo
                                    gnc.ProrratearIVKardexRecargo
                                    gnc.GeneraAsiento       'Diego 27 Abril 2001  corregido
                                    'Graba la transacción
                                    gnc.BandReproCostos = True
                                    gnc.Grabar False, False
                                    grd.TextMatrix(i, COL_RESULTADO) = "Actualizado."
                                    sql = " Update GNComprobante"
                                    sql = sql & " set "
                                    sql = sql & " CodUsuario = '" & UsuarioAnt & "',"
                                    sql = sql & " CodUsuarioModifica = '" & UsuarioModAnt & "',"
                                    sql = sql & " FechaGrabado = '" & FechaAnt & "'"
                                    sql = sql & " where transid =" & gnc.TransID
                                    gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                                'Si es solo verificacion
                                Else
                                    grd.TextMatrix(i, COL_RESULTADO) = MSG_NG
                                End If
                            Else
                                'Si no está cambiado no graba
                                grd.TextMatrix(i, COL_RESULTADO) = "OK."
                            End If
                        Else
                            grd.TextMatrix(i, COL_RESULTADO) = "Falló al recalcular."
                        End If
                    ElseIf gnc.GNTrans.CodPantalla = "IVCAMIE" And InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "CAMARI") <> 0 Then
                        'transformaciones de varios ingresos a varios egresos CAMARI
                        If RecalculoCAMARI(gnc, cambiado, bandVerificar) Then
                            'Si está cambiado algo
                            If cambiado Or BandTodo Then
                                'Si no es solo verificacion
                                If Not bandVerificar Then
                                    FechaAnt = gnc.FechaGrabado
                                    UsuarioAnt = gnc.codUsuario
                                    UsuarioModAnt = gnc.codUsuarioModifica
                                    grd.TextMatrix(i, COL_RESULTADO) = "Grabando..."
                                    grd.Refresh
                                    'Prorratea los recargos/descuentos si los calcula en base a costo
                                    gnc.ProrratearIVKardexRecargo
                                    gnc.GeneraAsiento       'Diego 27 Abril 2001  corregido
                                    'Graba la transacción
                                    gnc.BandReproCostos = True
                                    gnc.Grabar False, False
                                    grd.TextMatrix(i, COL_RESULTADO) = "Actualizado."
                                    sql = " Update GNComprobante"
                                    sql = sql & " set "
                                    sql = sql & " CodUsuario = '" & UsuarioAnt & "',"
                                    sql = sql & " CodUsuarioModifica = '" & UsuarioModAnt & "',"
                                    sql = sql & " FechaGrabado = '" & FechaAnt & "'"
                                    sql = sql & " where transid =" & gnc.TransID
                                    gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                                'Si es solo verificacion
                                Else
                                    grd.TextMatrix(i, COL_RESULTADO) = MSG_NG
                                End If
                            Else
                                'Si no está cambiado no graba
                                grd.TextMatrix(i, COL_RESULTADO) = "OK."
                            End If
                        Else
                            grd.TextMatrix(i, COL_RESULTADO) = "Falló al recalcular."
                        End If
                    Else
                       'Recalcula costo de los items
                       If gobjMain.EmpresaActual.GNOpcion.IVKTipoDatoDouble Then
                           If RecalculoDou(gnc, cambiado, bandVerificar) Then
                               'Si está cambiado algo
                               If cambiado Or BandTodo Then
                                   'Si no es solo verificacion
                                   If Not bandVerificar Then
                                       FechaAnt = gnc.FechaGrabado
                                       UsuarioAnt = gnc.codUsuario
                                       UsuarioModAnt = gnc.codUsuarioModifica
                                       grd.TextMatrix(i, COL_RESULTADO) = "Grabando..."
                                       grd.Refresh
                                       'Prorratea los recargos/descuentos si los calcula en base a costo
                                       gnc.ProrratearIVKardexRecargo
                                       gnc.GeneraAsiento       'Diego 27 Abril 2001  corregido
                                       'Graba la transacción
                                       gnc.BandReproCostos = True
                                       gnc.Grabar False, False
                                       If gnc.GNTrans.IVAutoImpresor Then
                                            If gnc.FechaTrans > "01/07/2011" Then 'fecha del cambio por el sri
                                               TransID = gnc.ObtieneIdTransAsientoAutoimpresor(gnc.TransID, gnc.GNTrans.AsientoTrans)
                                               If TransID <> 0 Then
                                                   ModificaTransAsiento TransID, gnc
                                               Else
                                                   GrabarTransAutoNew gnc.GNTrans.AsientoTrans, gnc
                                               End If
                                           End If
                                       End If
                                       grd.TextMatrix(i, COL_RESULTADO) = "Actualizado."
                                       sql = " Update GNComprobante"
                                       sql = sql & " set "
                                       sql = sql & " CodUsuario = '" & UsuarioAnt & "',"
                                       sql = sql & " CodUsuarioModifica = '" & UsuarioModAnt & "',"
                                       sql = sql & " FechaGrabado = '" & FechaAnt & "'"
                                       sql = sql & " where transid =" & gnc.TransID
                                       gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                                   'Si es solo verificacion
                                   Else
                                       grd.TextMatrix(i, COL_RESULTADO) = MSG_NG
                                   End If
                               Else
                                   'Si no está cambiado no graba
                                   grd.TextMatrix(i, COL_RESULTADO) = "OK."
                               End If
                           Else
                               grd.TextMatrix(i, COL_RESULTADO) = "Falló al recalcular."
                           End If
                        Else
                            If Recalculo(gnc, cambiado, bandVerificar) Then
                               'Si está cambiado algo
                               If cambiado Or BandTodo Then
                                   'Si no es solo verificacion
                                   If Not bandVerificar Then
                                       FechaAnt = gnc.FechaGrabado
                                       UsuarioAnt = gnc.codUsuario
                                       UsuarioModAnt = gnc.codUsuarioModifica
                                       grd.TextMatrix(i, COL_RESULTADO) = "Grabando..."
                                       grd.Refresh
                                       'Prorratea los recargos/descuentos si los calcula en base a costo
                                       gnc.ProrratearIVKardexRecargo
                                       gnc.GeneraAsiento       'Diego 27 Abril 2001  corregido
                                       'Graba la transacción
                                       gnc.BandReproCostos = True
                                       gnc.Grabar False, False
                                       If gnc.GNTrans.IVAutoImpresor Then
                                            If gnc.FechaTrans > "01/07/2011" Then 'fecha del cambio por el sri
                                               TransID = gnc.ObtieneIdTransAsientoAutoimpresor(gnc.TransID, gnc.GNTrans.AsientoTrans)
                                               If TransID <> 0 Then
                                                   ModificaTransAsiento TransID, gnc
                                               Else
                                                   GrabarTransAutoNew gnc.GNTrans.AsientoTrans, gnc
                                               End If
                                           End If
                                       End If
                                       grd.TextMatrix(i, COL_RESULTADO) = "Actualizado."
                                       sql = " Update GNComprobante"
                                       sql = sql & " set "
                                       sql = sql & " CodUsuario = '" & UsuarioAnt & "',"
                                       sql = sql & " CodUsuarioModifica = '" & UsuarioModAnt & "',"
                                       sql = sql & " FechaGrabado = '" & FechaAnt & "'"
                                       sql = sql & " where transid =" & gnc.TransID
                                       gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                                   'Si es solo verificacion
                                   Else
                                       grd.TextMatrix(i, COL_RESULTADO) = MSG_NG
                                   End If
                               Else
                                   'Si no está cambiado no graba
                                   grd.TextMatrix(i, COL_RESULTADO) = "OK."
                               End If
                           Else
                               grd.TextMatrix(i, COL_RESULTADO) = "Falló al recalcular."
                           End If
                        End If
                    End If
                Else
                    'Si está anulado
                    If gnc.Estado = ESTADO_ANULADO Then
                        grd.TextMatrix(i, COL_RESULTADO) = "Anulado"
                    'Si no tiene nada que ver con recalculo de costo
                    Else
                        grd.TextMatrix(i, COL_RESULTADO) = "---"
                    End If
                End If
            Else
                grd.TextMatrix(i, COL_RESULTADO) = "No pudo recuperar la transación."
            End If
        End If
    Next i
    
    Screen.MousePointer = 0
    ReprocCosto = Not mCancelado
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    If i < grd.Rows And i >= grd.FixedRows Then
        grd.TextMatrix(i, COL_RESULTADO) = Err.Description
    End If
    DispErr
    prg1.value = prg1.min
salida:
    Set mColItems = Nothing         'Libera el objeto de coleccion
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdVerificar.Enabled = True
    cmdCorregirIVA.Enabled = True
    cmdBuscar.Enabled = True
    cmdAceptar.Enabled = True
    prg1.value = prg1.min
    Exit Function
End Function


Private Function Recalculo(ByVal gnc As GNComprobante, _
                           ByRef cambiado As Boolean, _
                           ByVal booVerificando As Boolean) As Boolean
    Dim item As IVinventario, ivk As IVKardex, i As Long, k As Long, n As Long, ItemIngreso As Long
    Dim ct As Currency, ctotal As Currency, s As String
    Dim CostoTotalEgreso As Currency
    Dim ivkOUT As IVKardex, itemOUT As IVinventario
    Dim CostoTotalPadre As Currency
    Dim ItemMedio As Integer
    Dim ctPrep  As Currency, acuCosto As Currency
    Dim FechaIngreso As Date
    Dim ITEMCONS As IVinventario, CONSUMO As IVConsumoDetalle
    Dim ctbanda As Currency, ctcemento As Currency, ctcojin As Currency, ctrelleno As Currency
    Dim ctb As Currency, idproc As Long, BandCostoTransfCambiado As Boolean, CostoTotalTransforma As Currency, IdPadre As Long
    On Error GoTo ErrTrap
        BandCostoTransfCambiado = False
        cambiado = False
        ItemMedio = gnc.CountIVKardex / 2
        For i = 1 To gnc.CountIVKardex
        Set ivk = gnc.IVKardex(i)
        
        Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
        'IVCAMIEP para preparaciones/transforaciones retro
        If gnc.GNTrans.CodPantalla = "IVCAMIE" Or gnc.GNTrans.CodPantalla = "IVCAMIEP" Then
            'para que solo revise los items de egreso
            If item.Tipo = CambioPresentacion And ivk.cantidad < 0 Then
            'REGENERA ITEM TIPO 3
'            ---------------------------------------- EN EL CASO DE UNA TRANS QUE ESTE UN ITEM TIPO 3
                Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
                If Not (item Is Nothing) Then
                    If ItemIncorrecto(item.CodInventario) Then      'Este item ya está marcado como incorrecto.
                        Debug.Print "Incorrecto por trans. anterior. cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                        cambiado = True
                        GoTo SiguienteItem
                    End If
                End If
                ct = item.CostoDouble2(gnc.FechaTrans, _
                    Abs(ivk.cantidad), _
                    gnc.TransID, _
                    gnc.HoraTrans)
                
                'Convierte en moneda de la transaccion
                If item.CodMoneda <> gnc.CodMoneda Then
                    ct = ct * gnc.Cotizacion(item.CodMoneda) / gnc.Cotizacion("")
                End If
                ctotal = ct * ivk.cantidad
                CostoTotalPadre = ctotal
                If ctotal <> ivk.CostoTotal Then
                    If booVerificando Then
                        'Almacena codigo de item para que de aquí en adelante todo marque como incorrecto.
                        mColItems.Add item:=item.CodInventario, key:=item.CodInventario
                        Debug.Print "Incorrecto 1 . cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                        Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotal
                    End If
                        ivk.CostoTotal = ctotal
                        ivk.CostoRealTotal = ctotal
                        cambiado = True
                End If
                                        ivk.CostoTotal = ctotal 'AUC REPROCESO DE PREPARACIONES
                        ivk.CostoRealTotal = ctotal
                        Set ivkOUT = gnc.IVKardex(gnc.CountIVKardex - gnc.CountIVKardex + 1)
                        Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                        ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
                                       Abs(ivk.cantidad), _
                                       gnc.TransID, _
                                       gnc.HoraTrans)
                        acuCosto = acuCosto + ctotal * -1
                        ivkOUT.CostoTotal = acuCosto 'ctotal * -1
                        ivkOUT.CostoRealTotal = acuCosto 'ctotal * -1
                       If Abs(acuCosto) <> Abs(ctPrep) Then
                            cambiado = True
                        Else
                            cambiado = False
                        End If
                        Set ivkOUT = Nothing
                        Set itemOUT = Nothing

                '--------------------HASTA AQUI
                GoTo SiguienteItem
            End If
        End If
'        'Solo de salida
            'Recupera el item
            Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
            
            If Not (item Is Nothing) Then
                '*** MAKOTO 31/ago/00
                If booVerificando Then
                    If ItemIncorrecto(item.CodInventario) Then      'Este item ya está marcado como incorrecto.
                        Debug.Print "Incorrecto por trans. anterior. cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                        cambiado = True
                        GoTo SiguienteItem
                    End If
                End If
                '*** MAKOTO 08/dic/00
                If (gnc.GNTrans.CodPantalla = "IVISOFAC" Or gnc.GNTrans.CodPantalla = "IVDVISO") And Not gnc.GNTrans.IVTransProd And gnc.CodTrans <> "SCI" And Not gnc.GNTrans.IVTransProd And gnc.CodTrans <> "FSCI" And ivk.TiempoEntrega <> "" Then
                    
                    IdPadre = gnc.Empresa.ObtieneCampoDetalleTicket(ivk.TiempoEntrega, "IdPadre")
                    If IdPadre = 0 Then
                        idproc = gnc.Empresa.ObtieneCampoDetalleTicket(ivk.TiempoEntrega, "TransIDProceso")
                    Else
                        idproc = gnc.Empresa.ObtieneCampoDetalleTicket(IdPadre, "TransIDProceso")
                    End If

                    If gnc.NumDias = 0 Then
                        If idproc = 0 Then
                            k = gnc.Empresa.ObtieneCampoDetalleTicket(ivk.TiempoEntrega, "Motivo")
                            If k = 3 Then
                                ct = ivk.Precio * 0.1
                            Else
                                If IdPadre = 0 Then
                                    ct = gnc.Empresa.CalculaCostoProceso(ivk.TiempoEntrega) * -1
                                Else
                                    ct = gnc.Empresa.CalculaCostoProceso(IdPadre) * -1
                                End If
                                ct = gnc.Empresa.ObtieneCampoDetalleTicket(ivk.TiempoEntrega, "ValorCarcasa")
                            End If
                        Else
                            If IdPadre = 0 Then
                                k = gnc.Empresa.ObtieneCampoDetalleTicket(ivk.TiempoEntrega, "TransidProceso")
                            Else
                                k = gnc.Empresa.ObtieneCampoDetalleTicket(IdPadre, "TransidProceso")
                            End If
                            If k <> 0 Then
                                If IdPadre = 0 Then
                                    ct = gnc.Empresa.CalculaCostoProceso(ivk.TiempoEntrega) * -1
                                Else
                                    ct = gnc.Empresa.CalculaCostoProceso(IdPadre) * -1
                                End If
                                If ct < 0 Then ct = ct * -1
                                
                            Else
                                ct = gnc.Empresa.ObtieneCampoDetalleTicket(ivk.TiempoEntrega, "ValorCarcasa")
                            End If
                        End If
                    Else
                    ct = item.CostoDouble2(gnc.FechaTrans, _
                                           Abs(ivk.cantidad), _
                                           gnc.TransID, _
                                           gnc.HoraTrans)
                    End If
                Else
                    ct = item.CostoDouble2(gnc.FechaTrans, _
                                           Abs(ivk.cantidad), _
                                           gnc.TransID, _
                                           gnc.HoraTrans)
                End If
                'Convierte en moneda de la transaccion
                If item.CodMoneda <> gnc.CodMoneda Then
                    ct = ct * gnc.Cotizacion(item.CodMoneda) / gnc.Cotizacion("")
                End If
                ctotal = ct * ivk.cantidad
                CostoTotalPadre = ctotal
                'Si el costo es diferente de lo que está grabado
                If ctotal <> ivk.CostoTotal Or BandCostoTransfCambiado Then
                '1----------------------
                    If gnc.GNTrans.IVTipoTrans = "C" And i > 1 And gnc.GNTrans.CodPantalla <> "IVCAMIE" And gnc.GNTrans.CodPantalla <> "IVCAMIEP" Then
                        Set ivkOUT = gnc.IVKardex(i - 1)
                        Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                        If Not (itemOUT Is Nothing) Then
                            '*** MAKOTO 31/ago/00
                            If booVerificando Then
                                If ItemIncorrecto(itemOUT.CodInventario) Then      'Este item ya está marcado como incorrecto.
                                    cambiado = True
                                    GoTo SiguienteItem
                                End If
                            End If
                            ct = itemOUT.CostoDouble2(gnc.FechaTrans, _
                                                   Abs(ivkOUT.cantidad), _
                                                   gnc.TransID, _
                                                   gnc.HoraTrans)
                            
                            'Convierte en moneda de la transaccion
                            If itemOUT.CodMoneda <> gnc.CodMoneda Then
                                ct = ct * gnc.Cotizacion(itemOUT.CodMoneda) / gnc.Cotizacion("")
                            End If
                            ctotal = ct * ivkOUT.cantidad * -1
                            ivk.CostoTotal = ctotal
                            ivk.CostoRealTotal = ctotal
                            Set ivkOUT = Nothing
                            Set itemOUT = Nothing
                        End If
                    
                    '2------------------------------------
                    'ElseIf gnc.GNTrans.IVTipoTrans = "C" And gnc.GNTrans.CodPantalla = "IVCAMIE" Then
                    ElseIf gnc.GNTrans.IVTipoTrans = "C" And i > ItemMedio And gnc.GNTrans.CodPantalla = "IVCAMIE" And InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "CAMARI") <> 0 Then
                    '------------ PARA CAMARI
                             ivk.CostoTotal = ctotal
                             ivk.CostoRealTotal = ctotal
                             CostoTotalTransforma = 0
                             ItemIngreso = 0
                              For n = 1 To gnc.CountIVKardex
                                If gnc.IVKardex(n).cantidad > 0 Then
                                    ItemIngreso = n
                                    
'                                    Exit For
                                Else
'                                    CostoTotalTransforma = CostoTotalTransforma + gnc.IVKardex(n).CostoRealTotal
                                End If
                              Next n
                             
                             ItemIngreso = i - ItemMedio
                             Set ivkOUT = gnc.IVKardex(ItemIngreso)
'                             Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
'                             ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
                                            Abs(ivk.cantidad), _
                                            gnc.TransID, _
                                            gnc.HoraTrans)
                             acuCosto = ivk.CostoRealTotal * -1

                             ivkOUT.CostoTotal = acuCosto  'ctotal * -1
                             ivkOUT.CostoRealTotal = acuCosto  'ctotal * -1
                            If Abs(acuCosto) <> Abs(ctPrep) Then
                                 cambiado = True
                             Else
                                 cambiado = False
                             End If
                             Set ivkOUT = Nothing
                             Set itemOUT = Nothing
                        
'''                        End If
                         'aqui para las recetas cuando los costos  cuando los costos son diferentes
                    
                    ElseIf gnc.GNTrans.IVTipoTrans = "C" And i > 1 And gnc.GNTrans.CodPantalla = "IVCAMIE" Then
                             ivk.CostoTotal = ctotal 'AUC REPROCESO DE TRANSFORMACIONES ITALIANA
                             ivk.CostoRealTotal = ctotal
                             CostoTotalTransforma = 0
                             ItemIngreso = 0
                              For n = 1 To gnc.CountIVKardex
                                If gnc.IVKardex(n).cantidad > 0 Then
                                    ItemIngreso = n
                                    
'                                    Exit For
                                Else
                                    CostoTotalTransforma = CostoTotalTransforma + gnc.IVKardex(n).CostoRealTotal
                                End If
                              Next n
                             
                             Set ivkOUT = gnc.IVKardex(ItemIngreso)
                             Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                             ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
                                            Abs(ivk.cantidad), _
                                            gnc.TransID, _
                                            gnc.HoraTrans)
                             acuCosto = acuCosto + ctotal * -1
                             acuCosto = CostoTotalTransforma
                             ivkOUT.CostoTotal = acuCosto * -1 'ctotal * -1
                             ivkOUT.CostoRealTotal = acuCosto * -1 'ctotal * -1
                            If Abs(acuCosto) <> Abs(ctPrep) Then
                                 cambiado = True
                             Else
                                 cambiado = False
                             End If
                             Set ivkOUT = Nothing
                             Set itemOUT = Nothing
                        
'''                        End If
                         'aqui para las recetas cuando los costos  cuando los costos son diferentes
                    '3------------------------------------
                    ElseIf gnc.GNTrans.IVTipoTrans = "C" And gnc.GNTrans.CodPantalla = "IVCAMIEP" Then
                        ivk.CostoTotal = ctotal 'AUC REPROCESO DE PREPARACIONES
                        ivk.CostoRealTotal = ctotal
                        Set ivkOUT = gnc.IVKardex(gnc.CountIVKardex - gnc.CountIVKardex + 1)
                        Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                        ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
                                       Abs(ivk.cantidad), _
                                       gnc.TransID, _
                                       gnc.HoraTrans)
                        acuCosto = acuCosto + ctotal * -1
                        ivkOUT.CostoTotal = acuCosto 'ctotal * -1
                        ivkOUT.CostoRealTotal = acuCosto 'ctotal * -1
                       If Abs(acuCosto) <> Abs(ctPrep) Then
                            cambiado = True
                        Else
                            cambiado = False
                        End If
                        Set ivkOUT = Nothing
                        Set itemOUT = Nothing
                    'End If
                    Else
                        BandCostoTransfCambiado = True
                        '*** MAKOTO 31/ago/00
                        If booVerificando Then
                            'Almacena codigo de item para que de aquí en adelante todo marque como incorrecto.
                            mColItems.Add item:=item.CodInventario, key:=item.CodInventario
                            Debug.Print "Incorrecto 1 . cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                            Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotal
                        End If
                            ivk.CostoTotal = ctotal
                            ivk.CostoRealTotal = ctotal
                        cambiado = True
                        'jeaa 12/09/2005 recalculo de transformacion
                        If gnc.GNTrans.IVTipoTrans = "C" Then
                            If gnc.GNTrans.CodPantalla = "IVCAMIE" Then
                                acuCosto = acuCosto + ctotal
                            
                            ElseIf gnc.CountIVKardex = i + 1 Then
                                    CostoTotalEgreso = ctotal * -1
                            Else
                                ivk.CostoTotal = CostoTotalEgreso
                                ivk.CostoRealTotal = CostoTotalEgreso
                            End If
                        Else
                        End If
                    End If
                Else
                    'Esta parte es para cuando haya diferencia entre Costo y CostoReal
                    ' en las transacciones que no debe tener diferencia.
                    If (Not gnc.GNTrans.IVRecargoEnCosto) And (ivk.costo <> ivk.CostoReal) Then
                        ivk.CostoRealTotal = ivk.CostoTotal
                        cambiado = True
                    
                        '*** MAKOTO 31/ago/00
                        If booVerificando Then
                            'Almacena codigo de item para que de aquí en adelante todo marque como incorrecto.
                            mColItems.Add item:=item.CodInventario, key:=item.CodInventario
                            Debug.Print "Incorrecto 2 Agregado. cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                        End If
                    End If
                    
                    'aqui revisa el costo del item padre
                    If gnc.GNTrans.IVTipoTrans = "C" And i > 1 And gnc.GNTrans.CodPantalla = "IVCAMIE" Then
'                        If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "ITAL") = 0 And InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "MONT") Then
'                            ivk.CostoTotal = ctotal
'                            ivk.CostoRealTotal = ctotal
'                            If i > (gnc.CountIVKardex / 2) Then
'                                Set ivkOUT = gnc.IVKardex(Abs(i - (gnc.CountIVKardex / 2)))
'                            Else
'                                Set ivkOUT = gnc.IVKardex(Abs(i + (gnc.CountIVKardex / 2)))
'                            End If
'                            Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
'                            If Abs(ivkOUT.CostoTotal) <> Abs(ivk.CostoTotal) Then
'                                cambiado = True
'                                ivkOUT.CostoTotal = ctotal * -1
'                                ivkOUT.CostoRealTotal = ctotal * -1
'
'                                If Not ItemIncorrecto(itemOUT.CodInventario) Then      'Este item ya está marcado como incorrecto.
'                                    mColItems.Add item:=itemOUT.CodInventario, Key:=itemOUT.CodInventario
'                                    Debug.Print "Incorrecto 1 . cod='" & itemOUT.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
'                                      Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotal
'                                    GoTo SiguienteItem
'                                End If
'                            End If
'                            Set ivkOUT = Nothing
'                            Set itemOUT = Nothing
'                        Else
                             ivk.CostoTotal = ctotal 'AUC REPROCESO DE TRANSFORMACIONES ITALIANA
                             ivk.CostoRealTotal = ctotal
                              For n = 1 To gnc.CountIVKardex
                                If gnc.IVKardex(n).cantidad > 0 Then
                                    ItemIngreso = n
                                    Exit For
                                End If
                              Next n
                             Set ivkOUT = gnc.IVKardex(ItemIngreso)
                             Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                             ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
                                            Abs(ivk.cantidad), _
                                            gnc.TransID, _
                                            gnc.HoraTrans)
                             acuCosto = acuCosto + ctotal * -1
                             ivkOUT.CostoTotal = acuCosto 'ctotal * -1
                             ivkOUT.CostoRealTotal = acuCosto 'ctotal * -1
                            If Abs(acuCosto) <> Abs(ctPrep) Then
                                 cambiado = True
                             Else
                                 cambiado = False
                             End If
                             Set ivkOUT = Nothing
                             Set itemOUT = Nothing
                        
                        'End If
                   'ENDIF
                    'AUC AQUI DEBERIA IR EL REPROCESO PARA LA TRANSFORMACION cuando los costos son iguales
                    ElseIf gnc.GNTrans.IVTipoTrans = "C" And gnc.GNTrans.CodPantalla = "IVCAMIEP" Then
                        ivk.CostoTotal = ctotal 'AUC REPROCESO DE PREPARACIONES
                        ivk.CostoRealTotal = ctotal
                        Set ivkOUT = gnc.IVKardex(gnc.CountIVKardex - gnc.CountIVKardex + 1)
                        Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                        ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
                                       Abs(ivk.cantidad), _
                                       gnc.TransID, _
                                       gnc.HoraTrans)
                        acuCosto = acuCosto + ctotal * -1
                        ivkOUT.CostoTotal = acuCosto 'ctotal * -1
                        ivkOUT.CostoRealTotal = acuCosto 'ctotal * -1
                       If Abs(acuCosto) <> Abs(ctPrep) Then
                            cambiado = True
                        Else
                            cambiado = False
                        End If
                        Set ivkOUT = Nothing
                        Set itemOUT = Nothing
                    End If
                  End If
            'Si no puede recuperar el item
            Else
                'Aborta el recalculo
                cambiado = False            'Para que no se grabe
                GoTo salida
            End If
'        End If                 '*** MAKOTO 06/sep/00
SiguienteItem:
    Next i
SalidaOK:
    Recalculo = True
    GoTo salida
    Exit Function
ErrTrap:
    DispErr
salida:
    Set ivk = Nothing
    Set item = Nothing
    Set gnc = Nothing
    Exit Function
End Function

Private Function ItemIncorrecto(ByVal cod As String) As Boolean
    Dim s As String
    
    On Error Resume Next
    s = mColItems.item(cod)     'Si es que encuentra en la coleccion,
    If Err.Number = 0 Then      'Este item ya está marcado como incorrecto.
        ItemIncorrecto = True
    End If
    On Error GoTo 0
End Function

Private Sub cmdActualizaTrans_Click()
    Dim v As Variant, i As Integer, s As Variant, gt As GNTrans, cod As String, cad As String
    Dim CodTrans As String
    For i = 0 To lstTrans.ListCount - 1
        cod = lstTrans.List(i)
        Set gt = gobjMain.EmpresaActual.RecuperaGNTrans(cod)
        If Not (gt Is Nothing) Then
            If gt.Modulo = "IV" Then
                gt.IVReprocesaCosto = False
                gt.Grabar
            End If
        End If
    Next i
    
    cad = PreparaCodTrans

    s = ""
    v = Split(cad, ",")
    For i = 0 To UBound(v)
        v(i) = Trim(v(i))
        CodTrans = Mid$(v(i), 2, Len(v(i)) - 2)
        'actualiza campo en gntrans
        Set gt = gobjMain.EmpresaActual.RecuperaGNTrans(CodTrans)
        If Not (gt Is Nothing) Then
             gt.IVReprocesaCosto = True
             gt.Grabar
        End If
        s = s & Mid$(v(i), 2, Len(v(i)) - 2) & ","
    Next i
    Set gt = Nothing

End Sub

Private Sub cmdBuscar_Click()
    Dim v As Variant, obj As Object, s As String
    On Error GoTo ErrTrap
    
    '*** MAKOTO 06/sep/00 Agregado
    If lstTrans.SelCount = 0 Then
        MsgBox "Seleccione una transacción, por favor.", vbInformation
        Exit Sub
    End If
    
    With gobjMain.objCondicion
        .fecha1 = dtpFecha1.value
        .fecha2 = dtpFecha2.value
'        .CodTrans = fcbTrans.Text              '*** MAKOTO 31/ago/00 Modificado
        .CodTrans = PreparaCodTrans             '***
        .NumTrans1 = Val(txtNumTrans1.Text)
        .NumTrans2 = Val(txtNumTrans2.Text)
        
        'Estados no incluye anulados
        .EstadoBool(ESTADO_NOAPROBADO) = True
        .EstadoBool(ESTADO_APROBADO) = True
        .EstadoBool(ESTADO_DESPACHADO) = True
        .EstadoBool(ESTADO_ANULADO) = False
'        SaveSetting APPNAME, App.Title, "TransCostos", .CodTrans
        'jeaa 25/09/06
        s = PreparaTransParaGnopcion(.CodTrans)
        If Me.tag = "CostoxProveedor" Then
            .CodPC1 = fcbDesde.KeyText
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransparaRecosteoxProveedor", s
            If Len(.CodPC1) > 0 Then .BandTodo = False
            
        Else
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransparaRecosteo", s
        End If
        'Graba en la base
        gobjMain.EmpresaActual.GNOpcion.Grabar
        
        
    End With
    Set obj = gobjMain.EmpresaActual.ConsGNTrans2(True)  'Orden ascendente     '*** MAKOTO 20/oct/00
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
    
    cmdVerificar.Enabled = True
    cmdVerificar.SetFocus
    
    cmdCorregirIVA.Enabled = True
    
    cmdAceptar.Enabled = False
    chkTodo.Enabled = True
    mVerificado = False
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Function PreparaCodTrans() As String
    Dim i As Long, s As String
    
    With lstTrans
        'Si está seleccionado solo una
        If lstTrans.SelCount = 1 Then
            For i = 0 To .ListCount - 1
                If .Selected(i) Then
                    s = .List(i)
                    Exit For
                End If
            Next i
        'Si está TODO o NINGUNO, no hay condición
        ElseIf (.SelCount < .ListCount) And (.SelCount > 0) Then
            For i = 0 To .ListCount - 1
                If .Selected(i) Then
                    s = s & "'" & .List(i) & "', "
                End If
            Next i
            If Len(s) > 0 Then s = Left$(s, Len(s) - 2)    'Quita la ultima ", "
        End If
    End With
    PreparaCodTrans = s
End Function

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
        
'*** MAKOTO 20/oct/00 Eliminado
'        'Ordena por fecha ascendente      '*** MAKOTO 07/oct/00 Agregado por que cambió el orden del método
'        .col = COL_FECHA
'        .Sort = flexSortGenericAscending
        
        GNPoneNumFila grd, False
        .AutoSize 0, grd.Cols - 1
        
        .ColWidth(COL_NUMTRANS) = 700
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


Private Sub cmdCorregirIVA_Click()
    Dim ListaTrans As String
    ListaTrans = PreparaCodTrans
    frmReprocIVA.Inicio ListaTrans
End Sub

Private Sub cmdtranselec_Click()
    Dim i As Long, aux As Long, gt As GNTrans
    Dim cod As String
    On Error GoTo ErrTrap
    MensajeStatus "Preparando...", vbHourglass
    
    aux = lstTrans.ListIndex
    For i = 0 To lstTrans.ListCount - 1
        cod = lstTrans.List(i)
        Set gt = gobjMain.EmpresaActual.RecuperaGNTrans(cod)
        If Not (gt Is Nothing) Then
            'Solo marca egresos/transferencias
            If gt.IVReprocesaCosto Then
                lstTrans.Selected(i) = True
            End If
        End If
    Next i
    lstTrans.ListIndex = aux
    MensajeStatus
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    Exit Sub
End Sub

Private Sub cmdTransLimpiar_Click()
    Dim i As Long, aux As Long
    
    aux = lstTrans.ListIndex
    For i = 0 To lstTrans.ListCount - 1
        lstTrans.Selected(i) = False
    Next i
    lstTrans.ListIndex = aux
End Sub

Private Sub cmdTransTodo_Click()
    Dim i As Long, aux As Long, gt As GNTrans
    Dim cod As String
    On Error GoTo ErrTrap
    MensajeStatus "Preparando...", vbHourglass
    
    aux = lstTrans.ListIndex
    For i = 0 To lstTrans.ListCount - 1
        cod = lstTrans.List(i)
        Set gt = gobjMain.EmpresaActual.RecuperaGNTrans(cod)
        If Not (gt Is Nothing) Then
            'Solo marca egresos/transferencias
            If gt.IVTipoTrans = "E" Or gt.IVTipoTrans = "T" Then
                lstTrans.Selected(i) = True
            End If
        End If
    Next i
    lstTrans.ListIndex = aux
    MensajeStatus
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    Exit Sub
End Sub

Private Sub cmdVerificar_Click()
    'Si no hay transacciones
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna transacción para verificar."
        Exit Sub
    End If
    
    If dtpFecha1 < gobjMain.EmpresaActual.GNOpcion.FechaLimiteDesde Then
        MsgBox "La Rango de Fecha de reproceso es menor a la Fecha Limite Aceptable  ", vbExclamation
        Exit Sub
    End If
    
    If Me.tag = "CostoxProveedor" Then
        If ReprocCostoxProv(True, False) Then
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
            mVerificado = True
        End If
    
    Else
        If ReprocCosto(True, False) Then
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
            mVerificado = True
        End If
    End If
End Sub

Private Sub fcbDesde_Selected(ByVal Text As String, ByVal KeyText As String)
If Len(KeyText) > 0 Then
    
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF9
        cmdAceptar_Click
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

Public Sub RecuperaTrans(ByVal key As String, lst As ListBox, Optional s As String)
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

'jeaa 25/09/2006 elimina los apostrofes
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

Private Function ModificaTransAsiento(ByVal TransID As Long, ByRef mobjGNComp As GNComprobante) As Boolean
    Dim Imprime As Boolean, i As Long, ix As Long, j As Integer
    Dim item As IVinventario, rsReceta As Recordset
    Dim Cadena As String, aux_inc As Variant

    On Error GoTo ErrTrap
    Set mobjGNCompAux = gobjMain.EmpresaActual.RecuperaGNComprobante(TransID)
    
    If Not mobjGNCompAux Is Nothing Then
    
        For i = 1 To mobjGNCompAux.CountCTLibroDetalle
            mobjGNCompAux.RemoveCTLibroDetalle 1
        Next i
    
        If mobjGNComp.CountCTLibroDetalle > 0 Then
            For i = 1 To mobjGNComp.CountCTLibroDetalle
                ix = mobjGNCompAux.AddCTLibroDetalle
                mobjGNCompAux.CTLibroDetalle(ix).BandIntegridad = mobjGNComp.CTLibroDetalle(ix).BandIntegridad
                mobjGNCompAux.CTLibroDetalle(ix).codcuenta = mobjGNComp.CTLibroDetalle(ix).codcuenta
                mobjGNCompAux.CTLibroDetalle(ix).CodGasto = mobjGNComp.CTLibroDetalle(ix).CodGasto
                mobjGNCompAux.CTLibroDetalle(ix).Debe = mobjGNComp.CTLibroDetalle(ix).Debe
                mobjGNCompAux.CTLibroDetalle(ix).Descripcion = mobjGNComp.CTLibroDetalle(ix).Descripcion
                mobjGNCompAux.CTLibroDetalle(ix).Haber = mobjGNComp.CTLibroDetalle(ix).Haber
                mobjGNCompAux.CTLibroDetalle(ix).Orden = mobjGNComp.CTLibroDetalle(ix).Orden

            Next i
        End If
    
        
        mobjGNCompAux.Grabar False, False

        '***  Oliver 26/12/2002
        'Agregado para el control ded Impresion Configurado en la Transaccion

        MensajeStatus
        ModificaTransAsiento = True
    Else
        ModificaTransAsiento = False
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
    Exit Function

End Function

Private Sub PreparaAsientoTransAuto(Aceptar As Boolean)
    mobjGNCompAux.GeneraAsiento
End Sub



Private Function GrabarTransAutoNew(ByVal CodTrans As String, ByRef mobjGNComp As GNComprobante) As Boolean
    Dim Imprime As Boolean, i As Long, ix As Long, j As Integer
    Dim item As IVinventario, rsReceta As Recordset
    Dim Cadena As String, aux_inc As Variant

    On Error GoTo ErrTrap
    Set mobjGNCompAux = gobjMain.EmpresaActual.CreaGNComprobanteAutoimpresor(CodTrans)
    
    If Not mobjGNCompAux Is Nothing Then
    
        If mobjGNCompAux.SoloVer Then
            MsgBox MSG_NODISPONE, vbInformation
            Exit Function
        End If
        
        If mobjGNComp.CountCTLibroDetalle > 0 Then
            For i = 1 To mobjGNComp.CountCTLibroDetalle
                ix = mobjGNCompAux.AddCTLibroDetalle
                mobjGNCompAux.CTLibroDetalle(ix).BandIntegridad = mobjGNComp.CTLibroDetalle(ix).BandIntegridad
                mobjGNCompAux.CTLibroDetalle(ix).codcuenta = mobjGNComp.CTLibroDetalle(ix).codcuenta
                mobjGNCompAux.CTLibroDetalle(ix).CodGasto = mobjGNComp.CTLibroDetalle(ix).CodGasto
                mobjGNCompAux.CTLibroDetalle(ix).Debe = mobjGNComp.CTLibroDetalle(ix).Debe
                mobjGNCompAux.CTLibroDetalle(ix).Descripcion = mobjGNComp.CTLibroDetalle(ix).Descripcion
                mobjGNCompAux.CTLibroDetalle(ix).Haber = mobjGNComp.CTLibroDetalle(ix).Haber
                mobjGNCompAux.CTLibroDetalle(ix).Orden = mobjGNComp.CTLibroDetalle(ix).Orden
                
            Next i
        End If
       
        mobjGNCompAux.FechaTrans = mobjGNComp.FechaTrans
        mobjGNCompAux.HoraTrans = mobjGNComp.HoraTrans
        Cadena = "Por transaccion FACTURA " & mobjGNComp.CodTrans & "-" & mobjGNComp.numtrans & " / " & mobjGNComp.NumSerieEstaSRI & "-" & mobjGNComp.NumSeriePuntoSRI & "-" & Right("000000000" + Trim(Str(mobjGNComp.numtrans)), 9)    '& mobjGNComp.codtrans & "-" & mobjGNComp.NumTrans
        If Len(Cadena) > 120 Then
            mobjGNCompAux.Descripcion = Mid$(Cadena, 1, 120)
        Else
            mobjGNCompAux.Descripcion = Cadena
        End If
            
        mobjGNCompAux.codUsuario = mobjGNComp.codUsuario
        mobjGNCompAux.IdResponsable = mobjGNComp.IdResponsable
        mobjGNCompAux.numDocRef = mobjGNComp.NumSerieEstaSRI & "-" & mobjGNComp.NumSeriePuntoSRI & "-" & Right("000000000" + Trim(Str(mobjGNComp.numtrans)), 9)
        mobjGNCompAux.idCentro = mobjGNComp.idCentro
        mobjGNCompAux.idTransFuente = mobjGNComp.Empresa.RecuperarTransIDGncomprobante(mobjGNComp.CodTrans, mobjGNComp.numtrans)
        mobjGNCompAux.CodMoneda = mobjGNComp.CodMoneda

    
    
'        If GNTrans.ImportaCTD Then
'            mobjGNCompAux.ImportaAsiento mobjGNComp, aux_inc
 '       End If
    
    
    
        'Si es que algo está modificado
        If mobjGNCompAux.Modificado Then
            MensajeStatus MSG_GENERANDOASIENTO, vbHourglass
'            PreparaAsientoAutoNew True
            MensajeStatus
        End If
        'Verificación de datos
        mobjGNCompAux.VerificaDatos
    
'        PreparaAsientoAuto True
        'Verifica si está cuadrado el asiento
        If Not VerificaAsiento(mobjGNCompAux) Then Exit Function
    

        MensajeStatus MSG_GRABANDO, vbHourglass
    
        'Manda a grabar
        '       Aquí ya no hacemos verificación de asiento por que ya está hecho en Control Asiento
        mobjGNCompAux.Grabar False, False

        '***  Oliver 26/12/2002
        'Agregado para el control ded Impresion Configurado en la Transaccion

        MensajeStatus
        GrabarTransAutoNew = True
    Else
        GrabarTransAutoNew = False
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
    Exit Function

End Function


Private Function RecalculoMega(ByVal gnc As GNComprobante, _
                           ByRef cambiado As Boolean, _
                           ByVal booVerificando As Boolean) As Boolean
    Dim item As IVinventario, ivk As IVKardex, i As Long, k As Long, X  As Long
    Dim ct As Currency, ctotal As Currency, s As String
    Dim CostoTotalPadre As Currency
    Dim CostoTotal As Currency
    Dim ctPrep  As Currency, acuCosto As Currency
    Dim FechaIngreso As Date
    Dim ctbanda As Currency, ctcemento As Currency, ctcojin As Currency, ctrelleno As Currency
    Dim ctb As Currency, idproc As Long
    Dim ItemSub As IVinventario
    On Error GoTo ErrTrap
        cambiado = False
        For i = 1 To gnc.CountIVKardex
        Set ivk = gnc.IVKardex(i)
        Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
            'para que solo revise los items de egreso
            If item.Tipo <> CambioPresentacion Then GoTo SiguienteItem
            'REGENERA ITEM TIPO 3
'            ---------------------------------------- EN EL CASO DE UNA TRANS QUE ESTE UN ITEM TIPO 3
                Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
                If Not (item Is Nothing) Then
                    If ItemIncorrecto(item.CodInventario) Then      'Este item ya está marcado como incorrecto.
                        Debug.Print "Incorrecto por trans. anterior. cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                        cambiado = True
                        GoTo SiguienteItem
                    End If
                End If
                ct = item.CostoDouble2(gnc.FechaTrans, _
                    Abs(ivk.cantidad), _
                    gnc.TransID, _
                    gnc.HoraTrans)
                'Convierte en moneda de la transaccion
                If item.CodMoneda <> gnc.CodMoneda Then
                    ct = ct * gnc.Cotizacion(item.CodMoneda) / gnc.Cotizacion("")
                End If
                ctotal = ct * ivk.cantidad
                If ctotal <> ivk.CostoTotal Then
                    If booVerificando Then
                        'Almacena codigo de item para que de aquí en adelante todo marque como incorrecto.
                        mColItems.Add item:=item.CodInventario, key:=item.CodInventario
                        Debug.Print "Incorrecto 1 . cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                        Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotal
                    End If
                        ivk.CostoTotal = ctotal
                        ivk.CostoRealTotal = ctotal
                        cambiado = True
                End If
                        ivk.CostoTotal = ctotal
                        ivk.CostoRealTotal = ctotal
                        For k = 1 To item.NumFamiliaDetalle
                           Set ItemSub = gnc.Empresa.RecuperaIVInventario(item.RecuperaDetalleFamilia(k).CodInventario)
                            For X = 1 To gnc.CountIVKardex
                                If gnc.IVKardex(X).CodInventario = ItemSub.CodInventario Then
                                    If ctotal <> gnc.IVKardex(X).CostoTotal * -1 Then
                                        If booVerificando Then
                                            'Almacena codigo de item para que de aquí en adelante todo marque como incorrecto.
                                            'mColItems.Add item:=ItemSub.CodInventario, Key:=ItemSub.CodInventario
                                            'Debug.Print "Incorrecto 1 . cod='" & ItemSub.CodInventario & "' Trans=" & gnc.CodTrans & gnc.NumTrans
                                            'Debug.Print "    dif.:" & ctotal & "," & gnc.IVKardex(x).CostoTotal
                                            If ItemIncorrecto(ItemSub.CodInventario) Then      'Este item ya está marcado como incorrecto.
                                                Debug.Print "Incorrecto por trans. anterior. cod='" & ItemSub.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                                                cambiado = True
                                                GoTo SiguienteItem
                                            End If
                                        End If
                                        gnc.IVKardex(X).CostoTotal = ctotal * -1
                                        gnc.IVKardex(X).CostoRealTotal = ctotal * -1
                                        cambiado = True
                                    End If
                                End If
                            Next X
                        Next k
                        Set ItemSub = Nothing
                '--------------------HASTA AQUI
                GoTo SiguienteItem
            If Not (item Is Nothing) Then
                '*** MAKOTO 31/ago/00
                If booVerificando Then
                    If ItemIncorrecto(item.CodInventario) Then      'Este item ya está marcado como incorrecto.
                        Debug.Print "Incorrecto por trans. anterior. cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                        cambiado = True
                        GoTo SiguienteItem
                    End If
                End If
            Else
                'Aborta el recalculo
                cambiado = False            'Para que no se grabe
                GoTo salida
           End If
'        End If                 '*** MAKOTO 06/sep/00
SiguienteItem:
    Next i
SalidaOK:
    RecalculoMega = True
    GoTo salida
    Exit Function
ErrTrap:
    DispErr
salida:
    Set ivk = Nothing
    Set item = Nothing
    Set gnc = Nothing
    Exit Function
End Function



Private Function RecalculoCAMARI(ByVal gnc As GNComprobante, _
                           ByRef cambiado As Boolean, _
                           ByVal booVerificando As Boolean) As Boolean
    Dim item As IVinventario, ivk As IVKardex, i As Long, k As Long, n As Long, ItemIngreso As Long
    Dim ct As Currency, ctotal As Currency, s As String
    Dim CostoTotalEgreso As Currency
    Dim ivkOUT As IVKardex, itemOUT As IVinventario
    Dim ivkIN As IVKardex, itemIN As IVinventario, ctIN As Currency, ctotalIN As Currency
    Dim CostoTotalPadre As Currency
    Dim ItemMedio As Integer
    Dim ctPrep  As Currency, acuCosto As Currency
    Dim FechaIngreso As Date
    Dim ITEMCONS As IVinventario, CONSUMO As IVConsumoDetalle
    Dim ctbanda As Currency, ctcemento As Currency, ctcojin As Currency, ctrelleno As Currency
    Dim ctb As Currency, idproc As Long, BandCostoTransfCambiado As Boolean, CostoTotalTransforma As Currency
    On Error GoTo ErrTrap
        BandCostoTransfCambiado = False
        cambiado = False
        ItemMedio = gnc.CountIVKardex / 2
        For i = (ItemMedio + 1) To gnc.CountIVKardex
            Set ivk = gnc.IVKardex(i)
            Set ivkIN = gnc.IVKardex(i - ItemMedio)
            Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
            Set itemIN = gnc.Empresa.RecuperaIVInventario(ivkIN.CodInventario)
            
'        If gnc.GNTrans.CodPantalla = "IsVCAMIE" Or gnc.GNTrans.CodPantalla = "IVCAMIEP" Then
            'para que solo revise los items de egreso
            'If item.Tipo = CambioPresentacion And ivk.cantidad < 0 Then
            'REGENERA ITEM TIPO 3
'            ---------------------------------------- EN EL CASO DE UNA TRANS QUE ESTE UN ITEM TIPO 3
             '   Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
                If Not (item Is Nothing) Then
                    If ItemIncorrecto(item.CodInventario) Then      'Este item ya está marcado como incorrecto.
                        Debug.Print "Incorrecto por trans. anterior. cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                        cambiado = True
                        GoTo SiguienteItem
                    End If
                End If
                
                If Not (itemIN Is Nothing) Then
                    If ItemIncorrecto(itemIN.CodInventario) Then      'Este item ya está marcado como incorrecto.
                        Debug.Print "Incorrecto por trans. anterior. cod='" & itemIN.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                        cambiado = True
                        GoTo SiguienteItem
                    End If
                End If
                
                
                ct = item.CostoDouble2(gnc.FechaTrans, _
                    Abs(ivk.cantidad), _
                    gnc.TransID, _
                    gnc.HoraTrans)
                
            ctIN = itemIN.CostoDouble2(gnc.FechaTrans, _
                    Abs(ivkIN.cantidad), _
                    gnc.TransID, _
                    gnc.HoraTrans)
                
                'Convierte en moneda de la transaccion
                If item.CodMoneda <> gnc.CodMoneda Then
                    ct = ct * gnc.Cotizacion(item.CodMoneda) / gnc.Cotizacion("")
                End If
                ctotal = ct * ivk.cantidad
                
                If itemIN.CodMoneda <> gnc.CodMoneda Then
                    ctIN = ctIN * gnc.Cotizacion(itemIN.CodMoneda) / gnc.Cotizacion("")
                End If
                ctotalIN = (ctIN * ivkIN.cantidad) * -1
                
                

                If ctotal <> ivk.CostoTotal Or ivk.CostoTotal <> (ivkIN.CostoTotal * -1) Then
                    If booVerificando Then
                        'Almacena codigo de item para que de aquí en adelante todo marque como incorrecto.
                        mColItems.Add item:=item.CodInventario, key:=item.CodInventario
                        Debug.Print "Incorrecto 1 . cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                        Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotal
                    End If
                    ivk.CostoTotal = ctotal
                    ivk.CostoRealTotal = ctotal
                    ivkIN.CostoRealTotal = ctotal * -1
                    ivkIN.CostoTotal = ctotal * -1
                    cambiado = True
                End If
                Set ivkIN = Nothing
                Set itemIN = Nothing

                '--------------------HASTA AQUI
                GoTo SiguienteItem
            'End If
'        End If
'        'Solo de salida
            'Recupera el item
            Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
            
            If Not (item Is Nothing) Then
                '*** MAKOTO 31/ago/00
                If booVerificando Then
                    If ItemIncorrecto(item.CodInventario) Then      'Este item ya está marcado como incorrecto.
                        Debug.Print "Incorrecto por trans. anterior. cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                        cambiado = True
                        GoTo SiguienteItem
                    End If
                End If
                '*** MAKOTO 08/dic/00
                If (gnc.GNTrans.CodPantalla = "IVISOFAC" Or gnc.GNTrans.CodPantalla = "IVDVISO") And Not gnc.GNTrans.IVTransProd And gnc.CodTrans <> "SCI" And Not gnc.GNTrans.IVTransProd And gnc.CodTrans <> "FSCI" And ivk.TiempoEntrega <> "" Then
                    
                    idproc = gnc.Empresa.ObtieneCampoDetalleTicket(ivk.TiempoEntrega, "TransIDProceso")
                    If gnc.NumDias = 0 Then
                        If idproc = 0 Then
                            k = gnc.Empresa.ObtieneCampoDetalleTicket(ivk.TiempoEntrega, "Motivo")
                            If k = 3 Then
                                ct = ivk.Precio * 0.1
                            Else
                                ct = gnc.Empresa.CalculaCostoProceso(ivk.TiempoEntrega) * -1
                                ct = gnc.Empresa.ObtieneCampoDetalleTicket(ivk.TiempoEntrega, "ValorCarcasa")
                            End If
                        Else
                            k = gnc.Empresa.ObtieneCampoDetalleTicket(ivk.TiempoEntrega, "TransidProceso")
                            If k <> 0 Then
                                ct = gnc.Empresa.CalculaCostoProceso(ivk.TiempoEntrega) * -1
                                If ct < 0 Then ct = ct * -1
                                
                            Else
                                ct = gnc.Empresa.ObtieneCampoDetalleTicket(ivk.TiempoEntrega, "ValorCarcasa")
                            End If
                        End If
                    Else
                    ct = item.CostoDouble2(gnc.FechaTrans, _
                                           Abs(ivk.cantidad), _
                                           gnc.TransID, _
                                           gnc.HoraTrans)
                    End If
                Else
                    ct = item.CostoDouble2(gnc.FechaTrans, _
                                           Abs(ivk.cantidad), _
                                           gnc.TransID, _
                                           gnc.HoraTrans)
                End If
                'Convierte en moneda de la transaccion
                If item.CodMoneda <> gnc.CodMoneda Then
                    ct = ct * gnc.Cotizacion(item.CodMoneda) / gnc.Cotizacion("")
                End If
                    ctotal = ct * ivk.cantidad
                CostoTotalPadre = ctotal
                'Si el costo es diferente de lo que está grabado
                If ctotal <> ivk.CostoTotal Or BandCostoTransfCambiado Then
                '1----------------------
                    If gnc.GNTrans.IVTipoTrans = "C" And i > ItemMedio Then
                    '------------ PARA CAMARI
                             ivk.CostoTotal = ctotal
                             ivk.CostoRealTotal = ctotal
                             CostoTotalTransforma = 0
                             ItemIngreso = 0
                              For n = 1 To gnc.CountIVKardex
                                If gnc.IVKardex(n).cantidad > 0 Then
                                    ItemIngreso = n
                                Else
                                End If
                              Next n
                             
                             ItemIngreso = i - ItemMedio
                             Set ivkOUT = gnc.IVKardex(ItemIngreso)
                             acuCosto = ivk.CostoRealTotal * -1
                             ivkOUT.CostoTotal = acuCosto  'ctotal * -1
                             ivkOUT.CostoRealTotal = acuCosto  'ctotal * -1
                            If Abs(acuCosto) <> Abs(ctPrep) Then
                                 cambiado = True
                             Else
                                 cambiado = False
                             End If
                             Set ivkOUT = Nothing
                             Set itemOUT = Nothing
                        
                    Else
                        BandCostoTransfCambiado = True
                        '*** MAKOTO 31/ago/00
                        If booVerificando Then
                            'Almacena codigo de item para que de aquí en adelante todo marque como incorrecto.
                            mColItems.Add item:=item.CodInventario, key:=item.CodInventario
                            Debug.Print "Incorrecto 1 . cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                            Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotal
                    End If
                                               
                        ivk.CostoTotal = ctotal
                        ivk.CostoRealTotal = ctotal
                        cambiado = True
                        'jeaa 12/09/2005 recalculo de transformacion
                        If gnc.GNTrans.IVTipoTrans = "C" Then
                            If gnc.GNTrans.CodPantalla = "IVCAMIE" Then
                                acuCosto = acuCosto + ctotal
                            
                            ElseIf gnc.CountIVKardex = i + 1 Then
                                    CostoTotalEgreso = ctotal * -1
                            Else
                                ivk.CostoTotal = CostoTotalEgreso
                                ivk.CostoRealTotal = CostoTotalEgreso
                            End If
                        Else
                        End If
                    End If
                Else
                    'Esta parte es para cuando haya diferencia entre Costo y CostoReal
                    ' en las transacciones que no debe tener diferencia.
                    If (Not gnc.GNTrans.IVRecargoEnCosto) And (ivk.costo <> ivk.CostoReal) Then
                        ivk.CostoRealTotal = ivk.CostoTotal
                        cambiado = True
                    
                        '*** MAKOTO 31/ago/00
                        If booVerificando Then
                            'Almacena codigo de item para que de aquí en adelante todo marque como incorrecto.
                            mColItems.Add item:=item.CodInventario, key:=item.CodInventario
                            Debug.Print "Incorrecto 2 Agregado. cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                        End If
                    End If
                    
                    'aqui revisa el costo del item padre
                    If gnc.GNTrans.IVTipoTrans = "C" And i > 1 And gnc.GNTrans.CodPantalla = "IVCAMIE" Then
                             ivk.CostoTotal = ctotal 'AUC REPROCESO DE TRANSFORMACIONES ITALIANA
                             ivk.CostoRealTotal = ctotal
                              For n = 1 To gnc.CountIVKardex
                                If gnc.IVKardex(n).cantidad > 0 Then
                                    ItemIngreso = n
                                    Exit For
                                End If
                              Next n
                             Set ivkOUT = gnc.IVKardex(ItemIngreso)
                             Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                             ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
                                            Abs(ivk.cantidad), _
                                            gnc.TransID, _
                                            gnc.HoraTrans)
                             acuCosto = acuCosto + ctotal * -1
                             ivkOUT.CostoTotal = acuCosto 'ctotal * -1
                             ivkOUT.CostoRealTotal = acuCosto 'ctotal * -1
                            If Abs(acuCosto) <> Abs(ctPrep) Then
                                 cambiado = True
                             Else
                                 cambiado = False
                             End If
                             Set ivkOUT = Nothing
                             Set itemOUT = Nothing
                        
                        'End If
                   'ENDIF
                    'AUC AQUI DEBERIA IR EL REPROCESO PARA LA TRANSFORMACION cuando los costos son iguales
                    ElseIf gnc.GNTrans.IVTipoTrans = "C" And gnc.GNTrans.CodPantalla = "IVCAMIEP" Then
                        ivk.CostoTotal = ctotal 'AUC REPROCESO DE PREPARACIONES
                        ivk.CostoRealTotal = ctotal
                        Set ivkOUT = gnc.IVKardex(gnc.CountIVKardex - gnc.CountIVKardex + 1)
                        Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                        ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
                                       Abs(ivk.cantidad), _
                                       gnc.TransID, _
                                       gnc.HoraTrans)
                        acuCosto = acuCosto + ctotal * -1
                        ivkOUT.CostoTotal = acuCosto 'ctotal * -1
                        ivkOUT.CostoRealTotal = acuCosto 'ctotal * -1
                       If Abs(acuCosto) <> Abs(ctPrep) Then
                            cambiado = True
                        Else
                            cambiado = False
                        End If
                        Set ivkOUT = Nothing
                        Set itemOUT = Nothing
                    End If
                  End If
            'Si no puede recuperar el item
            Else
                'Aborta el recalculo
                cambiado = False            'Para que no se grabe
                GoTo salida
            End If
'        End If                 '*** MAKOTO 06/sep/00
SiguienteItem:
    Next i
SalidaOK:
    RecalculoCAMARI = True
    GoTo salida
    Exit Function
ErrTrap:
    DispErr
salida:
    Set ivk = Nothing
    Set item = Nothing
    Set gnc = Nothing
    Exit Function
End Function


Public Sub InicioCostoxProveedor(tag As String)
    Dim i As Integer
    On Error GoTo ErrTrap
    Me.tag = tag
    Me.Show
    Me.ZOrder
    fraprov.Visible = True
    dtpFecha1.value = gobjMain.EmpresaActual.GNOpcion.FechaInicio
    dtpFecha2.value = Date
    fcbDesde.SetData gobjMain.EmpresaActual.ListaPCProvCli(True, False, False)
    CargaTrans
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Function ReprocCostoxProv(ByVal bandVerificar As Boolean, BandTodo As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, X As Single
    Dim gnc As GNComprobante, cambiado As Boolean
    Dim FechaAnt As Date, UsuarioAnt As String, UsuarioModAnt As String
    Dim sql As String, NumReg As Long, TransID As Long
    
    On Error GoTo ErrTrap
    
    'Si no es solo verificacion, confirma
    If Not bandVerificar Then
        'Confirma la actualización
        s = "Este proceso modificará los costos de la transacción seleccionada." & vbCr & vbCr
        s = s & "Está seguro que desea proceder?"
        If MsgBox(s, vbYesNo + vbQuestion) <> vbYes Then Exit Function
    End If
    
    'Verifica si está seleccionado una trans. de ingreso
    s = VerificaIngreso
    If Len(s) > 0 Then
        'Si está seleccinada, confirma si está seguro
        s = "Está seleccionada una o más transacciones de ingreso. " & vbCr & _
            "(" & s & ")" & vbCr & _
            "Generalmente no se hace reprocesamiento de costo con transacciones de ingreso." & vbCr & vbCr
        s = s & "Confirma que desea proceder?" & vbCr & _
            "Aplaste 'Sí' unicamente cuando está seguro de lo que está haciendo."
'        If MsgBox(s, vbYesNo + vbQuestion + vbDefaultButton2) <> vbYes Then Exit Function
    End If
    s = ""
    
    Set mColItems = Nothing     'Limpia lo anterior
    Set mColItems = New Collection
    
    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdVerificar.Enabled = False
    cmdCorregirIVA.Enabled = False
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
        X = grd.CellTop                 'Para visualizar la celda actual
        
        'Si es verificación procesa todas las filas sino solo las que tengan "Costo Incorrecto"
        If ((grd.TextMatrix(i, COL_RESULTADO) = MSG_NG) Or bandVerificar Or BandTodo) Then
        
            tid = grd.ValueMatrix(i, COL_TID)
            grd.TextMatrix(i, COL_RESULTADO) = "Verificando..."
            grd.Refresh
            
            'Recupera la transaccion
            
            Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
'            If gnc.numtrans = 50 Then MsgBox "HOLA"
            If Not (gnc Is Nothing) Then
                'Si la transacción es de Inventario y es Egreso/Transferencia
                ' Y no está anulado
                If (gnc.GNTrans.Modulo = "IV") And _
                   (gnc.Estado <> ESTADO_ANULADO) Then
'                   (gnc.GNTrans.IVTipoTrans = "E" Or gnc.GNTrans.IVTipoTrans = "T") And _      '*** MAKOTO 06/sep/00 Eliminado

                    'Forzar recuperar todos los datos de transacción para que no se pierdan al grabar de nuveo
                    gnc.RecuperaDetalleTodo
                       'Recalcula costo de los items
                       If RecalculoxProv(gnc, cambiado, bandVerificar) Then
                           'Si está cambiado algo
                           If cambiado Or BandTodo Then
                               'Si no es solo verificacion
                               If Not bandVerificar Then
                                   FechaAnt = gnc.FechaGrabado
                                   UsuarioAnt = gnc.codUsuario
                                   UsuarioModAnt = gnc.codUsuarioModifica
                               
                                   grd.TextMatrix(i, COL_RESULTADO) = "Grabando..."
                                   grd.Refresh
                                   
                                   'Prorratea los recargos/descuentos si los calcula en base a costo
                                   gnc.ProrratearIVKardexRecargo
                                   gnc.GeneraAsiento       'Diego 27 Abril 2001  corregido
                                   'Graba la transacción
                                   gnc.BandReproCostos = True
                                   gnc.Grabar False, False
                                   grd.TextMatrix(i, COL_RESULTADO) = "Actualizado."
                                   
                                   
                                   sql = " Update GNComprobante"
                                   sql = sql & " set "
                                   sql = sql & " CodUsuario = '" & UsuarioAnt & "',"
                                   sql = sql & " CodUsuarioModifica = '" & UsuarioModAnt & "',"
                                   sql = sql & " FechaGrabado = '" & FechaAnt & "'"
                                   sql = sql & " where transid =" & gnc.TransID
                                   
                                   gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                                   
                                   
                                   
                               'Si es solo verificacion
                               Else
                                   grd.TextMatrix(i, COL_RESULTADO) = MSG_NG
                               End If
                           Else
                               'Si no está cambiado no graba
                               grd.TextMatrix(i, COL_RESULTADO) = "OK."
                           End If
                       Else
                           grd.TextMatrix(i, COL_RESULTADO) = "Falló al recalcular."
                       End If
                Else
                    'Si está anulado
                    If gnc.Estado = ESTADO_ANULADO Then
                        grd.TextMatrix(i, COL_RESULTADO) = "Anulado"
                    'Si no tiene nada que ver con recalculo de costo
                    Else
                        grd.TextMatrix(i, COL_RESULTADO) = "---"
                    End If
                End If
            Else
                grd.TextMatrix(i, COL_RESULTADO) = "No pudo recuperar la transación."
            End If
        End If
    Next i
    
    Screen.MousePointer = 0
    ReprocCostoxProv = Not mCancelado
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    If i < grd.Rows And i >= grd.FixedRows Then
        grd.TextMatrix(i, COL_RESULTADO) = Err.Description
    End If
    DispErr
    prg1.value = prg1.min
salida:
    Set mColItems = Nothing         'Libera el objeto de coleccion
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdVerificar.Enabled = True
    cmdCorregirIVA.Enabled = True
    cmdBuscar.Enabled = True
    cmdAceptar.Enabled = True
    prg1.value = prg1.min
    Exit Function
End Function


Private Function RecalculoxProv(ByVal gnc As GNComprobante, _
                           ByRef cambiado As Boolean, _
                           ByVal booVerificando As Boolean) As Boolean
    Dim item As IVinventario, ivk As IVKardex, i As Long, k As Long, n As Long, ItemIngreso As Long
    Dim ct As Currency, ctotal As Currency, s As String
    Dim CostoTotalEgreso As Currency
    Dim ivkOUT As IVKardex, itemOUT As IVinventario
    Dim CostoTotalPadre As Currency
    Dim ItemMedio As Integer
    Dim ctPrep  As Currency, acuCosto As Currency
    Dim FechaIngreso As Date
    Dim ITEMCONS As IVinventario, CONSUMO As IVConsumoDetalle
    Dim ctbanda As Currency, ctcemento As Currency, ctcojin As Currency, ctrelleno As Currency
    Dim ctb As Currency, idproc As Long, BandCostoTransfCambiado As Boolean, CostoTotalTransforma As Currency
    Dim pckCHP As PCKardexCHP
    On Error GoTo ErrTrap
    
    
    
    
    
    
    
        BandCostoTransfCambiado = False
        cambiado = False
    
    For i = 1 To gnc.CountIVKardex
        Set ivk = gnc.IVKardex(i)

        Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
        ct = item.CostoxProveedor(gnc.CodProveedorRef, ivk.CodInventario)

        If item.CodMoneda <> gnc.CodMoneda Then
            ct = ct * gnc.Cotizacion(item.CodMoneda) / gnc.Cotizacion("")
        End If
        ctotal = ct * ivk.cantidad
        CostoTotalPadre = ctotal

If gnc.GNTrans.IVTipoTrans = "I" Then

         If ctotal <> ivk.CostoTotal Or BandCostoTransfCambiado Or gnc.PCKardexCHP(1).Haber <> Round(ctotal, 2) Then
            BandCostoTransfCambiado = True
            ivk.CostoTotal = ctotal
            ivk.CostoRealTotal = ctotal
            cambiado = True
              
            If gnc.CountPCKardexCHP > 0 Then
                If gnc.GNTrans.IVTipoTrans = "E" Then
                        gnc.PCKardexCHP(1).Debe = Round(Abs(ctotal), 2)
                ElseIf gnc.GNTrans.IVTipoTrans = "I" Then
                        gnc.PCKardexCHP(1).Haber = Round(Abs(ctotal), 2)
                End If
            End If
        End If

ElseIf gnc.GNTrans.IVTipoTrans = "E" Then

         If ctotal <> ivk.CostoTotal Or BandCostoTransfCambiado Or gnc.PCKardexCHP(1).Debe <> Round(Abs(ctotal), 2) Then
            BandCostoTransfCambiado = True
            ivk.CostoTotal = ctotal
            ivk.CostoRealTotal = ctotal
            cambiado = True

            If gnc.CountPCKardexCHP > 0 Then
                If gnc.GNTrans.IVTipoTrans = "E" Then
                        gnc.PCKardexCHP(1).Debe = Round(Abs(ctotal), 2)
                ElseIf gnc.GNTrans.IVTipoTrans = "I" Then
                        gnc.PCKardexCHP(1).Haber = Round(Abs(ctotal), 2)
                End If
            End If
        End If
End If


    
    
    
    
    
    
    
'''''''        BandCostoTransfCambiado = False
'''''''        cambiado = False
'''''''        ItemMedio = gnc.CountIVKardex / 2
'''''''        For i = 1 To gnc.CountIVKardex
'''''''        Set ivk = gnc.IVKardex(i)
'''''''
'''''''        Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
'''''''        'IVCAMIEP para preparaciones/transforaciones retro
'''''''        If gnc.GNTrans.CodPantalla = "IVCAMIE" Or gnc.GNTrans.CodPantalla = "IVCAMIEP" Then
'''''''            'para que solo revise los items de egreso
'''''''            If item.Tipo = CambioPresentacion And ivk.cantidad < 0 Then
'''''''            'REGENERA ITEM TIPO 3
''''''''            ---------------------------------------- EN EL CASO DE UNA TRANS QUE ESTE UN ITEM TIPO 3
'''''''                Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
'''''''                If Not (item Is Nothing) Then
'''''''                    If ItemIncorrecto(item.CodInventario) Then      'Este item ya está marcado como incorrecto.
'''''''                        Debug.Print "Incorrecto por trans. anterior. cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
'''''''                        cambiado = True
'''''''                        GoTo SiguienteItem
'''''''                    End If
'''''''                End If
'''''''                ct = item.CostoDouble2(gnc.FechaTrans, _
'''''''                    Abs(ivk.cantidad), _
'''''''                    gnc.TransID, _
'''''''                    gnc.HoraTrans)
'''''''
'''''''                'Convierte en moneda de la transaccion
'''''''                If item.CodMoneda <> gnc.CodMoneda Then
'''''''                    ct = ct * gnc.Cotizacion(item.CodMoneda) / gnc.Cotizacion("")
'''''''                End If
'''''''                ctotal = ct * ivk.cantidad
'''''''                CostoTotalPadre = ctotal
'''''''                If ctotal <> ivk.CostoTotal Then
'''''''                    If booVerificando Then
'''''''                        'Almacena codigo de item para que de aquí en adelante todo marque como incorrecto.
'''''''                        mColItems.Add item:=item.CodInventario, Key:=item.CodInventario
'''''''                        Debug.Print "Incorrecto 1 . cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
'''''''                        Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotal
'''''''                    End If
'''''''                        ivk.CostoTotal = ctotal
'''''''                        ivk.CostoRealTotal = ctotal
'''''''                        cambiado = True
'''''''                End If
'''''''                                        ivk.CostoTotal = ctotal 'AUC REPROCESO DE PREPARACIONES
'''''''                        ivk.CostoRealTotal = ctotal
'''''''                        Set ivkOUT = gnc.IVKardex(gnc.CountIVKardex - gnc.CountIVKardex + 1)
'''''''                        Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
'''''''                        ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
'''''''                                       Abs(ivk.cantidad), _
'''''''                                       gnc.TransID, _
'''''''                                       gnc.HoraTrans)
'''''''                        acuCosto = acuCosto + ctotal * -1
'''''''                        ivkOUT.CostoTotal = acuCosto 'ctotal * -1
'''''''                        ivkOUT.CostoRealTotal = acuCosto 'ctotal * -1
'''''''                       If Abs(acuCosto) <> Abs(ctPrep) Then
'''''''                            cambiado = True
'''''''                        Else
'''''''                            cambiado = False
'''''''                        End If
'''''''                        Set ivkOUT = Nothing
'''''''                        Set itemOUT = Nothing
'''''''
'''''''                '--------------------HASTA AQUI
'''''''                GoTo SiguienteItem
'''''''            End If
'''''''        End If
''''''''        'Solo de salida
'''''''            'Recupera el item
'''''''            Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
'''''''
'''''''            If Not (item Is Nothing) Then
'''''''                '*** MAKOTO 31/ago/00
'''''''                If booVerificando Then
'''''''                    If ItemIncorrecto(item.CodInventario) Then      'Este item ya está marcado como incorrecto.
'''''''                        Debug.Print "Incorrecto por trans. anterior. cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
'''''''                        cambiado = True
'''''''                        GoTo SiguienteItem
'''''''                    End If
'''''''                End If
'''''''                '*** MAKOTO 08/dic/00
'''''''                If (gnc.GNTrans.CodPantalla = "IVISOFAC" Or gnc.GNTrans.CodPantalla = "IVDVISO") And Not gnc.GNTrans.IVTransProd And gnc.CodTrans <> "SCI" And Not gnc.GNTrans.IVTransProd And gnc.CodTrans <> "FSCI" And ivk.TiempoEntrega <> "" Then
'''''''
'''''''                    idproc = gnc.Empresa.ObtieneCampoDetalleTicket(ivk.TiempoEntrega, "TransIDProceso")
'''''''                    If gnc.NumDias = 0 Then
'''''''                        If idproc = 0 Then
'''''''                            k = gnc.Empresa.ObtieneCampoDetalleTicket(ivk.TiempoEntrega, "Motivo")
'''''''                            If k = 3 Then
'''''''                                ct = ivk.Precio * 0.1
'''''''                            Else
'''''''                                ct = gnc.Empresa.CalculaCostoProceso(ivk.TiempoEntrega) * -1
'''''''                                ct = gnc.Empresa.ObtieneCampoDetalleTicket(ivk.TiempoEntrega, "ValorCarcasa")
'''''''                            End If
'''''''                        Else
'''''''                            k = gnc.Empresa.ObtieneCampoDetalleTicket(ivk.TiempoEntrega, "TransidProceso")
'''''''                            If k <> 0 Then
'''''''                                ct = gnc.Empresa.CalculaCostoProceso(ivk.TiempoEntrega) * -1
'''''''                                If ct < 0 Then ct = ct * -1
'''''''
'''''''                            Else
'''''''                                ct = gnc.Empresa.ObtieneCampoDetalleTicket(ivk.TiempoEntrega, "ValorCarcasa")
'''''''                            End If
'''''''                        End If
'''''''                    Else
'''''''                    ct = item.CostoDouble2(gnc.FechaTrans, _
'''''''                                           Abs(ivk.cantidad), _
'''''''                                           gnc.TransID, _
'''''''                                           gnc.HoraTrans)
'''''''                    End If
'''''''                Else
'''''''                    ct = item.CostoDouble2(gnc.FechaTrans, _
'''''''                                           Abs(ivk.cantidad), _
'''''''                                           gnc.TransID, _
'''''''                                           gnc.HoraTrans)
'''''''                End If
'''''''                'Convierte en moneda de la transaccion
'''''''                If item.CodMoneda <> gnc.CodMoneda Then
'''''''                    ct = ct * gnc.Cotizacion(item.CodMoneda) / gnc.Cotizacion("")
'''''''                End If
'''''''                    ctotal = ct * ivk.cantidad
'''''''                CostoTotalPadre = ctotal
'''''''                'Si el costo es diferente de lo que está grabado
'''''''                If ctotal <> ivk.CostoTotal Or BandCostoTransfCambiado Then
'''''''                '1----------------------
'''''''                    If gnc.GNTrans.IVTipoTrans = "C" And i > 1 And gnc.GNTrans.CodPantalla <> "IVCAMIE" And gnc.GNTrans.CodPantalla <> "IVCAMIEP" Then
'''''''                        Set ivkOUT = gnc.IVKardex(i - 1)
'''''''                        Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
'''''''                        If Not (itemOUT Is Nothing) Then
'''''''                            '*** MAKOTO 31/ago/00
'''''''                            If booVerificando Then
'''''''                                If ItemIncorrecto(itemOUT.CodInventario) Then      'Este item ya está marcado como incorrecto.
'''''''                                    cambiado = True
'''''''                                    GoTo SiguienteItem
'''''''                                End If
'''''''                            End If
'''''''                            ct = itemOUT.CostoDouble2(gnc.FechaTrans, _
'''''''                                                   Abs(ivkOUT.cantidad), _
'''''''                                                   gnc.TransID, _
'''''''                                                   gnc.HoraTrans)
'''''''
'''''''                            'Convierte en moneda de la transaccion
'''''''                            If itemOUT.CodMoneda <> gnc.CodMoneda Then
'''''''                                ct = ct * gnc.Cotizacion(itemOUT.CodMoneda) / gnc.Cotizacion("")
'''''''                            End If
'''''''                            ctotal = ct * ivkOUT.cantidad * -1
'''''''                            ivk.CostoTotal = ctotal
'''''''                            ivk.CostoRealTotal = ctotal
'''''''                            Set ivkOUT = Nothing
'''''''                            Set itemOUT = Nothing
'''''''                        End If
'''''''
'''''''                    '2------------------------------------
'''''''                    'ElseIf gnc.GNTrans.IVTipoTrans = "C" And gnc.GNTrans.CodPantalla = "IVCAMIE" Then
'''''''                    ElseIf gnc.GNTrans.IVTipoTrans = "C" And i > ItemMedio And gnc.GNTrans.CodPantalla = "IVCAMIE" And InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "CAMARI") <> 0 Then
'''''''                    '------------ PARA CAMARI
'''''''                             ivk.CostoTotal = ctotal
'''''''                             ivk.CostoRealTotal = ctotal
'''''''                             CostoTotalTransforma = 0
'''''''                             ItemIngreso = 0
'''''''                              For n = 1 To gnc.CountIVKardex
'''''''                                If gnc.IVKardex(n).cantidad > 0 Then
'''''''                                    ItemIngreso = n
'''''''
''''''''                                    Exit For
'''''''                                Else
''''''''                                    CostoTotalTransforma = CostoTotalTransforma + gnc.IVKardex(n).CostoRealTotal
'''''''                                End If
'''''''                              Next n
'''''''
'''''''                             ItemIngreso = i - ItemMedio
'''''''                             Set ivkOUT = gnc.IVKardex(ItemIngreso)
''''''''                             Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
''''''''                             ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
'''''''                                            Abs(ivk.cantidad), _
'''''''                                            gnc.TransID, _
'''''''                                            gnc.HoraTrans)
'''''''                             acuCosto = ivk.CostoRealTotal * -1
'''''''
'''''''                             ivkOUT.CostoTotal = acuCosto  'ctotal * -1
'''''''                             ivkOUT.CostoRealTotal = acuCosto  'ctotal * -1
'''''''                            If Abs(acuCosto) <> Abs(ctPrep) Then
'''''''                                 cambiado = True
'''''''                             Else
'''''''                                 cambiado = False
'''''''                             End If
'''''''                             Set ivkOUT = Nothing
'''''''                             Set itemOUT = Nothing
'''''''
''''''''''                        End If
'''''''                         'aqui para las recetas cuando los costos  cuando los costos son diferentes
'''''''
'''''''                    ElseIf gnc.GNTrans.IVTipoTrans = "C" And i > 1 And gnc.GNTrans.CodPantalla = "IVCAMIE" Then
'''''''                             ivk.CostoTotal = ctotal 'AUC REPROCESO DE TRANSFORMACIONES ITALIANA
'''''''                             ivk.CostoRealTotal = ctotal
'''''''                             CostoTotalTransforma = 0
'''''''                             ItemIngreso = 0
'''''''                              For n = 1 To gnc.CountIVKardex
'''''''                                If gnc.IVKardex(n).cantidad > 0 Then
'''''''                                    ItemIngreso = n
'''''''
''''''''                                    Exit For
'''''''                                Else
'''''''                                    CostoTotalTransforma = CostoTotalTransforma + gnc.IVKardex(n).CostoRealTotal
'''''''                                End If
'''''''                              Next n
'''''''
'''''''                             Set ivkOUT = gnc.IVKardex(ItemIngreso)
'''''''                             Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
'''''''                             ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
'''''''                                            Abs(ivk.cantidad), _
'''''''                                            gnc.TransID, _
'''''''                                            gnc.HoraTrans)
'''''''                             acuCosto = acuCosto + ctotal * -1
'''''''                             acuCosto = CostoTotalTransforma
'''''''                             ivkOUT.CostoTotal = acuCosto * -1 'ctotal * -1
'''''''                             ivkOUT.CostoRealTotal = acuCosto * -1 'ctotal * -1
'''''''                            If Abs(acuCosto) <> Abs(ctPrep) Then
'''''''                                 cambiado = True
'''''''                             Else
'''''''                                 cambiado = False
'''''''                             End If
'''''''                             Set ivkOUT = Nothing
'''''''                             Set itemOUT = Nothing
'''''''
''''''''''                        End If
'''''''                         'aqui para las recetas cuando los costos  cuando los costos son diferentes
'''''''                    '3------------------------------------
'''''''                    ElseIf gnc.GNTrans.IVTipoTrans = "C" And gnc.GNTrans.CodPantalla = "IVCAMIEP" Then
'''''''                        ivk.CostoTotal = ctotal 'AUC REPROCESO DE PREPARACIONES
'''''''                        ivk.CostoRealTotal = ctotal
'''''''                        Set ivkOUT = gnc.IVKardex(gnc.CountIVKardex - gnc.CountIVKardex + 1)
'''''''                        Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
'''''''                        ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
'''''''                                       Abs(ivk.cantidad), _
'''''''                                       gnc.TransID, _
'''''''                                       gnc.HoraTrans)
'''''''                        acuCosto = acuCosto + ctotal * -1
'''''''                        ivkOUT.CostoTotal = acuCosto 'ctotal * -1
'''''''                        ivkOUT.CostoRealTotal = acuCosto 'ctotal * -1
'''''''                       If Abs(acuCosto) <> Abs(ctPrep) Then
'''''''                            cambiado = True
'''''''                        Else
'''''''                            cambiado = False
'''''''                        End If
'''''''                        Set ivkOUT = Nothing
'''''''                        Set itemOUT = Nothing
'''''''                    'End If
'''''''                    Else
'''''''                        BandCostoTransfCambiado = True
'''''''                        '*** MAKOTO 31/ago/00
'''''''                        If booVerificando Then
'''''''                            'Almacena codigo de item para que de aquí en adelante todo marque como incorrecto.
'''''''                            mColItems.Add item:=item.CodInventario, Key:=item.CodInventario
'''''''                            Debug.Print "Incorrecto 1 . cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
'''''''                            Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotal
'''''''                        End If
'''''''
'''''''                        ivk.CostoTotal = ctotal
'''''''                        ivk.CostoRealTotal = ctotal
'''''''                        cambiado = True
'''''''                        'jeaa 12/09/2005 recalculo de transformacion
'''''''                        If gnc.GNTrans.IVTipoTrans = "C" Then
'''''''                            If gnc.GNTrans.CodPantalla = "IVCAMIE" Then
'''''''                                acuCosto = acuCosto + ctotal
'''''''
'''''''                            ElseIf gnc.CountIVKardex = i + 1 Then
'''''''                                    CostoTotalEgreso = ctotal * -1
'''''''                            Else
'''''''                                ivk.CostoTotal = CostoTotalEgreso
'''''''                                ivk.CostoRealTotal = CostoTotalEgreso
'''''''                            End If
'''''''                        Else
'''''''                        End If
'''''''                    End If
'''''''                Else
'''''''                    'Esta parte es para cuando haya diferencia entre Costo y CostoReal
'''''''                    ' en las transacciones que no debe tener diferencia.
'''''''                    If (Not gnc.GNTrans.IVRecargoEnCosto) And (ivk.costo <> ivk.CostoReal) Then
'''''''                        ivk.CostoRealTotal = ivk.CostoTotal
'''''''                        cambiado = True
'''''''
'''''''                        '*** MAKOTO 31/ago/00
'''''''                        If booVerificando Then
'''''''                            'Almacena codigo de item para que de aquí en adelante todo marque como incorrecto.
'''''''                            mColItems.Add item:=item.CodInventario, Key:=item.CodInventario
'''''''                            Debug.Print "Incorrecto 2 Agregado. cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
'''''''                        End If
'''''''                    End If
'''''''
'''''''                    'aqui revisa el costo del item padre
'''''''                    If gnc.GNTrans.IVTipoTrans = "C" And i > 1 And gnc.GNTrans.CodPantalla = "IVCAMIE" Then
''''''''                        If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "ITAL") = 0 And InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "MONT") Then
''''''''                            ivk.CostoTotal = ctotal
''''''''                            ivk.CostoRealTotal = ctotal
''''''''                            If i > (gnc.CountIVKardex / 2) Then
''''''''                                Set ivkOUT = gnc.IVKardex(Abs(i - (gnc.CountIVKardex / 2)))
''''''''                            Else
''''''''                                Set ivkOUT = gnc.IVKardex(Abs(i + (gnc.CountIVKardex / 2)))
''''''''                            End If
''''''''                            Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
''''''''                            If Abs(ivkOUT.CostoTotal) <> Abs(ivk.CostoTotal) Then
''''''''                                cambiado = True
''''''''                                ivkOUT.CostoTotal = ctotal * -1
''''''''                                ivkOUT.CostoRealTotal = ctotal * -1
''''''''
''''''''                                If Not ItemIncorrecto(itemOUT.CodInventario) Then      'Este item ya está marcado como incorrecto.
''''''''                                    mColItems.Add item:=itemOUT.CodInventario, Key:=itemOUT.CodInventario
''''''''                                    Debug.Print "Incorrecto 1 . cod='" & itemOUT.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
''''''''                                      Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotal
''''''''                                    GoTo SiguienteItem
''''''''                                End If
''''''''                            End If
''''''''                            Set ivkOUT = Nothing
''''''''                            Set itemOUT = Nothing
''''''''                        Else
'''''''                             ivk.CostoTotal = ctotal 'AUC REPROCESO DE TRANSFORMACIONES ITALIANA
'''''''                             ivk.CostoRealTotal = ctotal
'''''''                              For n = 1 To gnc.CountIVKardex
'''''''                                If gnc.IVKardex(n).cantidad > 0 Then
'''''''                                    ItemIngreso = n
'''''''                                    Exit For
'''''''                                End If
'''''''                              Next n
'''''''                             Set ivkOUT = gnc.IVKardex(ItemIngreso)
'''''''                             Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
'''''''                             ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
'''''''                                            Abs(ivk.cantidad), _
'''''''                                            gnc.TransID, _
'''''''                                            gnc.HoraTrans)
'''''''                             acuCosto = acuCosto + ctotal * -1
'''''''                             ivkOUT.CostoTotal = acuCosto 'ctotal * -1
'''''''                             ivkOUT.CostoRealTotal = acuCosto 'ctotal * -1
'''''''                            If Abs(acuCosto) <> Abs(ctPrep) Then
'''''''                                 cambiado = True
'''''''                             Else
'''''''                                 cambiado = False
'''''''                             End If
'''''''                             Set ivkOUT = Nothing
'''''''                             Set itemOUT = Nothing
'''''''
'''''''                        'End If
'''''''                   'ENDIF
'''''''                    'AUC AQUI DEBERIA IR EL REPROCESO PARA LA TRANSFORMACION cuando los costos son iguales
'''''''                    ElseIf gnc.GNTrans.IVTipoTrans = "C" And gnc.GNTrans.CodPantalla = "IVCAMIEP" Then
'''''''                        ivk.CostoTotal = ctotal 'AUC REPROCESO DE PREPARACIONES
'''''''                        ivk.CostoRealTotal = ctotal
'''''''                        Set ivkOUT = gnc.IVKardex(gnc.CountIVKardex - gnc.CountIVKardex + 1)
'''''''                        Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
'''''''                        ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
'''''''                                       Abs(ivk.cantidad), _
'''''''                                       gnc.TransID, _
'''''''                                       gnc.HoraTrans)
'''''''                        acuCosto = acuCosto + ctotal * -1
'''''''                        ivkOUT.CostoTotal = acuCosto 'ctotal * -1
'''''''                        ivkOUT.CostoRealTotal = acuCosto 'ctotal * -1
'''''''                       If Abs(acuCosto) <> Abs(ctPrep) Then
'''''''                            cambiado = True
'''''''                        Else
'''''''                            cambiado = False
'''''''                        End If
'''''''                        Set ivkOUT = Nothing
'''''''                        Set itemOUT = Nothing
'''''''                    End If
'''''''                  End If
'''''''            'Si no puede recuperar el item
'''''''            Else
'''''''                'Aborta el recalculo
'''''''                cambiado = False            'Para que no se grabe
'''''''                GoTo salida
'''''''            End If
'        End If                 '*** MAKOTO 06/sep/00
SiguienteItem:
    Next i
SalidaOK:
    RecalculoxProv = True
    GoTo salida
    Exit Function
ErrTrap:
    DispErr
salida:
    Set ivk = Nothing
    Set item = Nothing
    Set gnc = Nothing
    Exit Function
End Function


Private Function RecalculoTicketISO(ByVal gnc As GNComprobante, _
                           ByRef cambiado As Boolean, _
                           ByVal booVerificando As Boolean) As Boolean
    Dim item As IVinventario, ivk As IVKardex, i As Long, k As Long, X  As Long, ivkOUT As IVKardex, ctOut As Currency
    Dim ct As Currency, ctotal As Currency, s As String
    Dim CostoTotalPadre As Currency
    Dim CostoTotal As Currency
    Dim ctPrep  As Currency, acuCosto As Currency
    Dim FechaIngreso As Date
    Dim ctbanda As Currency, ctcemento As Currency, ctcojin As Currency, ctrelleno As Currency, posItemIn As Long, posItemOut As Long, CantTicket As Integer, numtiketProc As Integer
    Dim ctb As Currency, idproc As Long
    Dim ItemSub As IVinventario
    On Error GoTo ErrTrap
        cambiado = False
        
        CantTicket = gnc.CountIVKardex / 2
        
        For i = 1 To CantTicket
        numtiketProc = i
        posItemIn = 0
        posItemOut = 0
        For X = 1 To gnc.CountIVKardex
            If gnc.IVKardex(X).Orden = numtiketProc And gnc.IVKardex(X).cantidad = -1 Then
                posItemOut = X
                Exit For
            End If
        Next X
        
        For X = 1 To gnc.CountIVKardex
            If gnc.IVKardex(X).Orden = numtiketProc And gnc.IVKardex(X).cantidad = 1 Then
                posItemIn = X
                Exit For
            End If
        Next X
        
        

        
        Set ivk = gnc.IVKardex(posItemIn)
        Set ivkOUT = gnc.IVKardex(posItemOut)
        Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
        
                    idproc = gnc.Empresa.ObtieneCampoDetalleTicket(ivkOUT.TiempoEntrega, "TransIDProceso")
                    If gnc.NumDias = 0 Then
                            k = gnc.Empresa.ObtieneCampoDetalleTicket(ivkOUT.TiempoEntrega, "TransidProceso")
                            If k <> 0 Then
                                ctOut = gnc.Empresa.CalculaCostoProceso(ivkOUT.TiempoEntrega) * -1
                                ct = ivk.CostoRealTotal
'                                If ct < 0 Then ct = ct * -1
                            Else
                                ct = gnc.Empresa.ObtieneCampoDetalleTicket(ivk.TiempoEntrega, "ValorCarcasa")
                            End If
                    Else
                        ct = item.CostoDouble2(gnc.FechaTrans, _
                                               Abs(ivk.cantidad), _
                                               gnc.TransID, _
                                               gnc.HoraTrans)
                    End If
        
        
                Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
                If Not (item Is Nothing) Then
                    If ItemIncorrecto(item.CodInventario) Then      'Este item ya está marcado como incorrecto.
                        Debug.Print "Incorrecto por trans. anterior. cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                        cambiado = True
                        GoTo SiguienteItem
                    End If
                End If
                'ct = item.CostoDouble2(gnc.FechaTrans, _
                    Abs(ivk.cantidad), _
                    gnc.TransID, _
                    gnc.HoraTrans)
                'Convierte en moneda de la transaccion
                If item.CodMoneda <> gnc.CodMoneda Then
                    ctOut = ctOut * gnc.Cotizacion(item.CodMoneda) / gnc.Cotizacion("")
                End If
                ctotal = ctOut * ivk.cantidad
                If ctotal <> ivk.CostoTotal Then
                    If booVerificando Then
                        'Almacena codigo de item para que de aquí en adelante todo marque como incorrecto.
                        mColItems.Add item:=item.CodInventario, key:=item.CodInventario
                        Debug.Print "Incorrecto 1 . cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                        Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotal
                    End If
                        ivk.CostoTotal = ctotal
                        ivk.CostoRealTotal = ctotal
                        cambiado = True
                End If
                    ivk.CostoTotal = ctotal
                    ivk.CostoRealTotal = ctotal
                '--------------------HASTA AQUI
                GoTo SiguienteItem
            If Not (item Is Nothing) Then
                '*** MAKOTO 31/ago/00
                If booVerificando Then
                    If ItemIncorrecto(item.CodInventario) Then      'Este item ya está marcado como incorrecto.
                        Debug.Print "Incorrecto por trans. anterior. cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                        cambiado = True
                        GoTo SiguienteItem
                    End If
                End If
            Else
                'Aborta el recalculo
                cambiado = False            'Para que no se grabe
                GoTo salida
           End If
'        End If                 '*** MAKOTO 06/sep/00
SiguienteItem:
    Next i
SalidaOK:
    RecalculoTicketISO = True
    GoTo salida
    Exit Function
ErrTrap:
    DispErr
salida:
    Set ivk = Nothing
    Set item = Nothing
    Set gnc = Nothing
    Exit Function
End Function

Private Function RecalculoDou(ByVal gnc As GNComprobante, _
                           ByRef cambiado As Boolean, _
                           ByVal booVerificando As Boolean) As Boolean
    Dim item As IVinventario, ivk As IVKardex, i As Long, k As Long, n As Long, ItemIngreso As Long
    Dim ct As Double, ctotal As Double, s As String
    Dim CostoTotalEgreso As Double
    Dim ivkOUT As IVKardex, itemOUT As IVinventario
    Dim CostoTotalPadre As Double
    Dim ItemMedio As Integer
    Dim ctPrep  As Double, acuCosto As Double
    Dim FechaIngreso As Date
    Dim ITEMCONS As IVinventario, CONSUMO As IVConsumoDetalle
    Dim ctbanda As Double, ctcemento As Double, ctcojin As Double, ctrelleno As Double
    Dim ctb As Double, idproc As Long, BandCostoTransfCambiado As Boolean, CostoTotalTransforma As Double, IdPadre As Long
    Dim fmt As String
    On Error GoTo ErrTrap
    fmt = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("FormatoCantRec")
        BandCostoTransfCambiado = False
        cambiado = False
        ItemMedio = gnc.CountIVKardex / 2
        For i = 1 To gnc.CountIVKardex
        Set ivk = gnc.IVKardex(i)
        Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
        'IVCAMIEP para preparaciones/transforaciones retro
        If gnc.GNTrans.CodPantalla = "IVCAMIE" Or gnc.GNTrans.CodPantalla = "IVCAMIEP" Then
            'para que solo revise los items de egreso
            If item.Tipo = CambioPresentacion And ivk.cantidadDou < 0 Then
            'REGENERA ITEM TIPO 3
'            ---------------------------------------- EN EL CASO DE UNA TRANS QUE ESTE UN ITEM TIPO 3
                Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
                If Not (item Is Nothing) Then
                    If ItemIncorrecto(item.CodInventario) Then      'Este item ya está marcado como incorrecto.
                        Debug.Print "Incorrecto por trans. anterior. cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                        cambiado = True
                        GoTo SiguienteItem
                    End If
                End If
                ct = item.CostoDouble2(gnc.FechaTrans, _
                    Abs(ivk.cantidadDou), _
                    gnc.TransID, _
                    gnc.HoraTrans)
                
                'Convierte en moneda de la transaccion
                If item.CodMoneda <> gnc.CodMoneda Then
                    ct = ct * gnc.Cotizacion(item.CodMoneda) / gnc.Cotizacion("")
                End If
                ctotal = ct * ivk.cantidadDou
                CostoTotalPadre = ctotal
                If ctotal <> ivk.CostoTotalDou Then
                    If booVerificando Then
                        'Almacena codigo de item para que de aquí en adelante todo marque como incorrecto.
                        mColItems.Add item:=item.CodInventario, key:=item.CodInventario
                        Debug.Print "Incorrecto 1 . cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                        Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotalDou
                    End If
                        ivk.CostoTotalDou = ctotal
                        ivk.CostoRealTotaldou = ctotal
                        cambiado = True
                End If
                        ivk.CostoTotalDou = ctotal 'AUC REPROCESO DE PREPARACIONES
                        ivk.CostoRealTotaldou = ctotal
                        Set ivkOUT = gnc.IVKardex(gnc.CountIVKardex - gnc.CountIVKardex + 1)
                        Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                        ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
                                       Abs(ivk.cantidadDou), _
                                       gnc.TransID, _
                                       gnc.HoraTrans)
                        acuCosto = acuCosto + ctotal * -1
                        ivkOUT.CostoTotalDou = acuCosto 'ctotal * -1
                        ivkOUT.CostoRealTotaldou = acuCosto 'ctotal * -1
                       If Abs(acuCosto) <> Abs(ctPrep) Then
                            cambiado = True
                        Else
                            cambiado = False
                        End If
                        Set ivkOUT = Nothing
                        Set itemOUT = Nothing

                '--------------------HASTA AQUI
                GoTo SiguienteItem
            End If
        End If
        ''-AQUI PARA EL RECALCULO DE LAS RECETAS BALGRAN--------------------------------------
        If gnc.GNTrans.CodPantalla = "IVCAMIEREC" Then
            'para que solo revise los items de egreso
'            If ivk.CantidadDou > 0 Then
'            MsgBox "para"
'            End If
            If ivk.cantidadDou < 0 Then
                Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
                If Not (item Is Nothing) Then
                    If ItemIncorrecto(item.CodInventario) Then      'Este item ya está marcado como incorrecto.
                        Debug.Print "Incorrecto por trans. anterior. cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                        cambiado = True
                        GoTo SiguienteItem
                    End If
                End If
                ct = item.CostoDouble2(gnc.FechaTrans, _
                    Abs(ivk.cantidadDou), _
                    gnc.TransID, _
                    gnc.HoraTrans)
                'Convierte en moneda de la transaccion
                If item.CodMoneda <> gnc.CodMoneda Then
                    ct = ct * gnc.Cotizacion(item.CodMoneda) / gnc.Cotizacion("")
                End If
                'aqui estoy revisando este proceso
                'de decimales
                '-----------------------------------------------------------
                ctotal = Format(ct * ivk.cantidadDou, fmt)
                CostoTotalPadre = ctotal
                If ctotal <> ivk.CostoTotalDou Then
                    If booVerificando Then
                        'Almacena codigo de item para que de aquí en adelante todo marque como incorrecto.
                        mColItems.Add item:=item.CodInventario, key:=item.CodInventario
                        Debug.Print "Incorrecto 1 . cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                        Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotalDou
                    End If
                        ivk.CostoTotalDou = Format(ctotal, fmt)
                        ivk.CostoRealTotaldou = Format(ctotal, fmt)
                        cambiado = True
                End If
                        ivk.CostoTotalDou = Format(ctotal, fmt)
                        ivk.CostoRealTotaldou = Format(ctotal, fmt)
                        'Set ivkOUT = gnc.IVKardex(gnc.CountIVKardex - gnc.CountIVKardex + 1)
                        'Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                        'ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
                                       Abs(ivk.CantidadDou), _
                                       gnc.TransID, _
                                       gnc.HoraTrans)
                        acuCosto = acuCosto + ctotal * -1
                       ' ivkOUT.CostoTotalDou = acuCosto 'ctotal * -1
                        'ivkOUT.CostoRealTotalDou = acuCosto 'ctotal * -1
                       'If Abs(acuCosto) <> Abs(ctPrep) Then
                       '     cambiado = True
                       ' Else
                       '     cambiado = False
                       ' End If
                        'Set ivkOUT = Nothing
                        'Set itemOUT = Nothing
                GoTo SiguienteItem
            End If
        End If
        ''---------------------------------------HASTA AQUI LO DE BALGRAN
            'Recupera el item
            Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
            If Not (item Is Nothing) Then
                '*** MAKOTO 31/ago/00
                If booVerificando Then
                    If ItemIncorrecto(item.CodInventario) Then      'Este item ya está marcado como incorrecto.
                        Debug.Print "Incorrecto por trans. anterior. cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                        cambiado = True
                        GoTo SiguienteItem
                    End If
                End If
                '*** MAKOTO 08/dic/00
                If (gnc.GNTrans.CodPantalla = "IVISOFAC" Or gnc.GNTrans.CodPantalla = "IVDVISO") And Not gnc.GNTrans.IVTransProd And gnc.CodTrans <> "SCI" And Not gnc.GNTrans.IVTransProd And gnc.CodTrans <> "FSCI" And ivk.TiempoEntrega <> "" Then
                    IdPadre = gnc.Empresa.ObtieneCampoDetalleTicket(ivk.TiempoEntrega, "IdPadre")
                    If IdPadre = 0 Then
                        idproc = gnc.Empresa.ObtieneCampoDetalleTicket(ivk.TiempoEntrega, "TransIDProceso")
                    Else
                        idproc = gnc.Empresa.ObtieneCampoDetalleTicket(IdPadre, "TransIDProceso")
                    End If

                    If gnc.NumDias = 0 Then
                        If idproc = 0 Then
                            k = gnc.Empresa.ObtieneCampoDetalleTicket(ivk.TiempoEntrega, "Motivo")
                            If k = 3 Then
                                ct = ivk.Precio * 0.1
                            Else
                                If IdPadre = 0 Then
                                    ct = gnc.Empresa.CalculaCostoProceso(ivk.TiempoEntrega) * -1
                                Else
                                    ct = gnc.Empresa.CalculaCostoProceso(IdPadre) * -1
                                End If
                                ct = gnc.Empresa.ObtieneCampoDetalleTicket(ivk.TiempoEntrega, "ValorCarcasa")
                            End If
                        Else
                            If IdPadre = 0 Then
                                k = gnc.Empresa.ObtieneCampoDetalleTicket(ivk.TiempoEntrega, "TransidProceso")
                            Else
                                k = gnc.Empresa.ObtieneCampoDetalleTicket(IdPadre, "TransidProceso")
                            End If
                            If k <> 0 Then
                                If IdPadre = 0 Then
                                    ct = gnc.Empresa.CalculaCostoProceso(ivk.TiempoEntrega) * -1
                                Else
                                    ct = gnc.Empresa.CalculaCostoProceso(IdPadre) * -1
                                End If
                                If ct < 0 Then ct = ct * -1
                                
                            Else
                                ct = gnc.Empresa.ObtieneCampoDetalleTicket(ivk.TiempoEntrega, "ValorCarcasa")
                            End If
                        End If
                    Else
                    ct = item.CostoDouble2(gnc.FechaTrans, _
                                           Abs(ivk.cantidadDou), _
                                           gnc.TransID, _
                                           gnc.HoraTrans)
                    End If
                Else
                    ct = item.CostoDouble2(gnc.FechaTrans, _
                                           Abs(ivk.cantidadDou), _
                                           gnc.TransID, _
                                           gnc.HoraTrans)
                End If
                'Convierte en moneda de la transaccion
                If item.CodMoneda <> gnc.CodMoneda Then
                    ct = ct * gnc.Cotizacion(item.CodMoneda) / gnc.Cotizacion("")
                End If
                    ctotal = ct * ivk.cantidadDou
                CostoTotalPadre = ctotal
                'Si el costo es diferente de lo que está grabado
                If ctotal <> ivk.CostoTotalDou Or BandCostoTransfCambiado Then
                '1----------------------
                    If gnc.GNTrans.IVTipoTrans = "C" And i > 1 And gnc.GNTrans.CodPantalla <> "IVCAMIE" And gnc.GNTrans.CodPantalla <> "IVCAMIEP" Then
                        Set ivkOUT = gnc.IVKardex(i - 1)
                        Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                        If Not (itemOUT Is Nothing) Then
                            '*** MAKOTO 31/ago/00
                            If booVerificando Then
                                If ItemIncorrecto(itemOUT.CodInventario) Then      'Este item ya está marcado como incorrecto.
                                    cambiado = True
                                    GoTo SiguienteItem
                                End If
                            End If
                            ct = itemOUT.CostoDouble2(gnc.FechaTrans, _
                                                   Abs(ivkOUT.cantidadDou), _
                                                   gnc.TransID, _
                                                   gnc.HoraTrans)
                            
                            'Convierte en moneda de la transaccion
                            If itemOUT.CodMoneda <> gnc.CodMoneda Then
                                ct = ct * gnc.Cotizacion(itemOUT.CodMoneda) / gnc.Cotizacion("")
                            End If
                            ctotal = ct * ivkOUT.cantidadDou * -1
                            ivk.CostoTotalDou = ctotal
                            ivk.CostoRealTotaldou = ctotal
                            Set ivkOUT = Nothing
                            Set itemOUT = Nothing
                        End If
                    
                    '2------------------------------------
                    'ElseIf gnc.GNTrans.IVTipoTrans = "C" And gnc.GNTrans.CodPantalla = "IVCAMIE" Then
                    ElseIf gnc.GNTrans.IVTipoTrans = "C" And i > ItemMedio And gnc.GNTrans.CodPantalla = "IVCAMIE" And InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "CAMARI") <> 0 Then
                    '------------ PARA CAMARI
                             ivk.CostoTotalDou = ctotal
                             ivk.CostoRealTotaldou = ctotal
                             CostoTotalTransforma = 0
                             ItemIngreso = 0
                              For n = 1 To gnc.CountIVKardex
                                If gnc.IVKardex(n).cantidadDou > 0 Then
                                    ItemIngreso = n
'                                    Exit For
                                Else
'                                    CostoTotalTransforma = CostoTotalTransforma + gnc.IVKardex(n).CostoRealTotal
                                End If
                              Next n
                             
                             ItemIngreso = i - ItemMedio
                             Set ivkOUT = gnc.IVKardex(ItemIngreso)
'                             Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
'                             ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
                                            Abs(ivk.cantidad), _
                                            gnc.TransID, _
                                            gnc.HoraTrans)
                             acuCosto = ivk.CostoRealTotaldou * -1
                             ivkOUT.CostoTotalDou = acuCosto  'ctotal * -1
                             ivkOUT.CostoRealTotaldou = acuCosto  'ctotal * -1
                            If Abs(acuCosto) <> Abs(ctPrep) Then
                                 cambiado = True
                             Else
                                 cambiado = False
                             End If
                             Set ivkOUT = Nothing
                             Set itemOUT = Nothing
                        
'''                        End If
                         'aqui para las recetas cuando los costos  cuando los costos son diferentes
                    
                    ElseIf gnc.GNTrans.IVTipoTrans = "C" And i > 1 And gnc.GNTrans.CodPantalla = "IVCAMIE" Then
                             ivk.CostoTotalDou = ctotal 'AUC REPROCESO DE TRANSFORMACIONES ITALIANA
                             ivk.CostoRealTotaldou = ctotal
                             CostoTotalTransforma = 0
                             ItemIngreso = 0
                              For n = 1 To gnc.CountIVKardex
                                If gnc.IVKardex(n).cantidadDou > 0 Then
                                    ItemIngreso = n
'                                    Exit For
                                Else
                                    CostoTotalTransforma = CostoTotalTransforma + gnc.IVKardex(n).CostoRealTotaldou
                                End If
                              Next n
                             
                             Set ivkOUT = gnc.IVKardex(ItemIngreso)
                             Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                             ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
                                            Abs(ivk.cantidadDou), _
                                            gnc.TransID, _
                                            gnc.HoraTrans)
                             acuCosto = acuCosto + ctotal * -1
                             acuCosto = CostoTotalTransforma
                             ivkOUT.CostoTotalDou = acuCosto * -1 'ctotal * -1
                             ivkOUT.CostoRealTotaldou = acuCosto * -1 'ctotal * -1
                            If Abs(acuCosto) <> Abs(ctPrep) Then
                                 cambiado = True
                             Else
                                 cambiado = False
                             End If
                             Set ivkOUT = Nothing
                             Set itemOUT = Nothing
                        
'''                        End If
                         'aqui para las recetas cuando los costos  cuando los costos son diferentes
                    '3------------------------------------
                    ElseIf gnc.GNTrans.IVTipoTrans = "C" And gnc.GNTrans.CodPantalla = "IVCAMIEP" Then
                        ivk.CostoTotalDou = ctotal 'AUC REPROCESO DE PREPARACIONES
                        ivk.CostoRealTotaldou = ctotal
                        Set ivkOUT = gnc.IVKardex(gnc.CountIVKardex - gnc.CountIVKardex + 1)
                        Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                        ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
                                       Abs(ivk.cantidadDou), _
                                       gnc.TransID, _
                                       gnc.HoraTrans)
                        acuCosto = acuCosto + ctotal * -1
                        ivkOUT.CostoTotalDou = acuCosto 'ctotal * -1
                        ivkOUT.CostoRealTotaldou = acuCosto 'ctotal * -1
                       If Abs(acuCosto) <> Abs(ctPrep) Then
                            cambiado = True
                        Else
                            cambiado = False
                        End If
                        Set ivkOUT = Nothing
                        Set itemOUT = Nothing
                    'End If
                    ElseIf gnc.GNTrans.IVTipoTrans = "R" Then
                            ivk.CostoTotalDou = ctotal
                            ivk.CostoRealTotaldou = ctotal
                            For n = 1 To gnc.CountIVKardex
                                If gnc.IVKardex(n).cantidadDou > 0 Then
                                    ItemIngreso = n
                                    Exit For
                                End If
                             Next n
                             Set ivkOUT = gnc.IVKardex(ItemIngreso)
                             Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                             ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
                                            Abs(ivk.cantidadDou), _
                                            gnc.TransID, _
                                            gnc.HoraTrans)
                             'acuCosto = acuCosto + ctotal * -1
                             ivkOUT.CostoTotalDou = Format(acuCosto, fmt) 'ctotal * -1
                             ivkOUT.CostoRealTotaldou = Format(acuCosto, fmt) 'ctotal * -1
                            If Abs(acuCosto) <> Abs(ctPrep) Then
                                 cambiado = True
                             Else
                                 cambiado = False
                             End If
                             Set ivkOUT = Nothing
                             Set itemOUT = Nothing
                    Else
                        BandCostoTransfCambiado = True
                        '*** MAKOTO 31/ago/00
                        If booVerificando Then
                            'Almacena codigo de item para que de aquí en adelante todo marque como incorrecto.
                            mColItems.Add item:=item.CodInventario, key:=item.CodInventario
                            Debug.Print "Incorrecto 1 . cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                                Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotalDou
                        End If
                        
                            ivk.CostoTotalDou = ctotal
                            ivk.CostoRealTotaldou = ctotal
                        cambiado = True
                        'jeaa 12/09/2005 recalculo de transformacion
                        If gnc.GNTrans.IVTipoTrans = "C" Then
                            If gnc.GNTrans.CodPantalla = "IVCAMIE" Then
                                acuCosto = acuCosto + ctotal
                            
                            ElseIf gnc.CountIVKardex = i + 1 Then
                                    CostoTotalEgreso = ctotal * -1
                            Else
                                ivk.CostoTotalDou = CostoTotalEgreso
                                ivk.CostoRealTotaldou = CostoTotalEgreso
                            End If
                        Else
                        End If
                    End If
                Else
                    'Esta parte es para cuando haya diferencia entre Costo y CostoReal
                    ' en las transacciones que no debe tener diferencia.
                    If (Not gnc.GNTrans.IVRecargoEnCosto) And (ivk.costoDou <> ivk.CostoRealDou) Then
                        ivk.CostoRealTotaldou = ivk.CostoTotalDou
                        cambiado = True
                        '*** MAKOTO 31/ago/00
                        If booVerificando Then
                            'Almacena codigo de item para que de aquí en adelante todo marque como incorrecto.
                            mColItems.Add item:=item.CodInventario, key:=item.CodInventario
                            Debug.Print "Incorrecto 2 Agregado. cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                        End If
                    End If
                    
                    'aqui revisa el costo del item padre
                    If gnc.GNTrans.IVTipoTrans = "C" And i > 1 And gnc.GNTrans.CodPantalla = "IVCAMIE" Then
'                        If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "ITAL") = 0 And InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "MONT") Then
'                            ivk.CostoTotal = ctotal
'                            ivk.CostoRealTotal = ctotal
'                            If i > (gnc.CountIVKardex / 2) Then
'                                Set ivkOUT = gnc.IVKardex(Abs(i - (gnc.CountIVKardex / 2)))
'                            Else
'                                Set ivkOUT = gnc.IVKardex(Abs(i + (gnc.CountIVKardex / 2)))
'                            End If
'                            Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
'                            If Abs(ivkOUT.CostoTotal) <> Abs(ivk.CostoTotal) Then
'                                cambiado = True
'                                ivkOUT.CostoTotal = ctotal * -1
'                                ivkOUT.CostoRealTotal = ctotal * -1
'
'                                If Not ItemIncorrecto(itemOUT.CodInventario) Then      'Este item ya está marcado como incorrecto.
'                                    mColItems.Add item:=itemOUT.CodInventario, Key:=itemOUT.CodInventario
'                                    Debug.Print "Incorrecto 1 . cod='" & itemOUT.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
'                                      Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotal
'                                    GoTo SiguienteItem
'                                End If
'                            End If
'                            Set ivkOUT = Nothing
'                            Set itemOUT = Nothing
'                        Else
                             ivk.CostoTotalDou = ctotal 'AUC REPROCESO DE TRANSFORMACIONES ITALIANA
                             ivk.CostoRealTotaldou = ctotal
                              For n = 1 To gnc.CountIVKardex
                                If gnc.IVKardex(n).cantidadDou > 0 Then
                                    ItemIngreso = n
                                    Exit For
                                End If
                              Next n
                             Set ivkOUT = gnc.IVKardex(ItemIngreso)
                             Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                             ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
                                            Abs(ivk.cantidadDou), _
                                            gnc.TransID, _
                                            gnc.HoraTrans)
                             acuCosto = acuCosto + ctotal * -1
                             ivkOUT.CostoTotalDou = acuCosto 'ctotal * -1
                             ivkOUT.CostoRealTotaldou = acuCosto 'ctotal * -1
                            If Abs(acuCosto) <> Abs(ctPrep) Then
                                 cambiado = True
                             Else
                                 cambiado = False
                             End If
                             Set ivkOUT = Nothing
                             Set itemOUT = Nothing
                        'End If
                   'ENDIF
                    'AUC AQUI DEBERIA IR EL REPROCESO PARA LA TRANSFORMACION cuando los costos son iguales
                    ElseIf gnc.GNTrans.IVTipoTrans = "C" And gnc.GNTrans.CodPantalla = "IVCAMIEP" Then
                        ivk.CostoTotalDou = ctotal 'AUC REPROCESO DE PREPARACIONES
                        ivk.CostoRealTotaldou = ctotal
                        Set ivkOUT = gnc.IVKardex(gnc.CountIVKardex - gnc.CountIVKardex + 1)
                        Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                        ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
                                       Abs(ivk.cantidadDou), _
                                       gnc.TransID, _
                                       gnc.HoraTrans)
                        acuCosto = acuCosto + ctotal * -1
                        ivkOUT.CostoTotalDou = acuCosto 'ctotal * -1
                        ivkOUT.CostoRealTotaldou = acuCosto 'ctotal * -1
                        If Abs(acuCosto) <> Abs(ctPrep) Then
                            cambiado = True
                        Else
                            cambiado = False
                        End If
                        Set ivkOUT = Nothing
                        Set itemOUT = Nothing
                    ElseIf gnc.GNTrans.IVTipoTrans = "R" Then
                            ivk.CostoTotalDou = ctotal 'AUC REPROCESO DE TRANSFORMACIONES ITALIANA
                            ivk.CostoRealTotaldou = ctotal
                            For n = 1 To gnc.CountIVKardex
                                If gnc.IVKardex(n).cantidadDou > 0 Then
                                    ItemIngreso = n
                                    Exit For
                                End If
                             Next n
                             Set ivkOUT = gnc.IVKardex(ItemIngreso)
                             Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                             ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
                                            Abs(ivk.cantidadDou), _
                                            gnc.TransID, _
                                            gnc.HoraTrans)
                             acuCosto = acuCosto + ctotal * -1
                             ivkOUT.CostoTotalDou = acuCosto 'ctotal * -1
                             ivkOUT.CostoRealTotaldou = acuCosto 'ctotal * -1
                            If Abs(acuCosto) <> Abs(ctPrep) Then
                                 cambiado = True
                             Else
                                 cambiado = False
                             End If
                             Set ivkOUT = Nothing
                             Set itemOUT = Nothing
                    End If
                  End If
            'Si no puede recuperar el item
            Else
                'Aborta el recalculo
                cambiado = False            'Para que no se grabe
                GoTo salida
            End If
'        End If                 '*** MAKOTO 06/sep/00
SiguienteItem:
    Next i
SalidaOK:
    RecalculoDou = True
    GoTo salida
    Exit Function
ErrTrap:
    DispErr
salida:
    Set ivk = Nothing
    Set item = Nothing
    Set gnc = Nothing
    Exit Function
End Function

