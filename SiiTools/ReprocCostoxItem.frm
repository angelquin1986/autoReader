VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "flexcombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReprocCostoxItem 
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
   Begin VB.CheckBox chkTodo 
      Caption         =   "&Regenerar todo sin verificar"
      Enabled         =   0   'False
      Height          =   192
      Left            =   4080
      TabIndex        =   30
      Top             =   1800
      Width           =   3252
   End
   Begin VB.Frame fraitem 
      Caption         =   "Items"
      Height          =   675
      Left            =   6720
      TabIndex        =   25
      Top             =   1020
      Width           =   5052
      Begin FlexComboProy.FlexCombo fcbDesde2 
         Height          =   315
         Left            =   840
         TabIndex        =   26
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
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
      Begin FlexComboProy.FlexCombo fcbHasta2 
         Height          =   315
         Left            =   3225
         TabIndex        =   27
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
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
         Caption         =   "Hasta"
         Height          =   252
         Left            =   2760
         TabIndex        =   29
         Top             =   240
         Width           =   612
      End
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   252
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   612
      End
   End
   Begin VB.Frame fraGrupos 
      Caption         =   "Rango de Grupos"
      Height          =   915
      Left            =   6720
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   5052
      Begin VB.ComboBox cboGrupo 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   480
         Width           =   1452
      End
      Begin FlexComboProy.FlexCombo fcbGrupoHasta 
         Height          =   300
         Left            =   3360
         TabIndex        =   20
         Top             =   480
         Width           =   1572
         _ExtentX        =   2778
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
      Begin FlexComboProy.FlexCombo fcbGrupoDesde 
         Height          =   300
         Left            =   1812
         TabIndex        =   21
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Grupo"
         Height          =   192
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   444
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Desde"
         Height          =   192
         Left            =   1800
         TabIndex        =   23
         Top             =   240
         Width           =   492
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Hasta"
         Height          =   192
         Left            =   3360
         TabIndex        =   22
         Top             =   240
         Width           =   432
      End
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
         Format          =   152567809
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
         Format          =   152567809
         CurrentDate     =   36348
      End
   End
   Begin VB.Frame fraCodTrans 
      Caption         =   "Cod.&Trans"
      Height          =   1572
      Left            =   2088
      TabIndex        =   3
      Top             =   120
      Width           =   2772
      Begin VB.CommandButton cmdTransLimpiar 
         Caption         =   "Limp."
         Height          =   330
         Left            =   1800
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
         Width           =   1332
      End
      Begin VB.ListBox lstTrans 
         Columns         =   3
         Height          =   852
         IntegralHeight  =   0   'False
         Left            =   240
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   14
         Top             =   240
         Width           =   2412
      End
   End
   Begin VB.Frame fraNumTrans 
      Caption         =   "# T&rans. (desde - hasta)"
      Height          =   1572
      Left            =   4728
      TabIndex        =   4
      Top             =   120
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
      Height          =   1932
      Left            =   120
      TabIndex        =   8
      Top             =   2160
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
      Left            =   2520
      TabIndex        =   7
      Top             =   1740
      Width           =   1452
   End
End
Attribute VB_Name = "frmReprocCostoxItem"
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
Private Const COL_CUR = 9
Private Const COL_CU = 10
Private Const COL_CENTROCOSTO = 11
Private Const COL_ESTADO = 12
Private Const COL_RESULTADO = 13


Private Const MSG_NG = "Costo incorrecto."
Private mProcesando As Boolean
Private mCancelado As Boolean
Private mVerificado As Boolean
'*** MAKOTO 31/ago/00 Agregado
'       para almacenar items con costo incorrecto detectado
Private mColItems As Collection
Const IVGRUPO_MAX = 5
Dim numGrupo As Integer
Private mobjGNCompAux As GNComprobante



Public Sub Inicio()
    Dim i As Integer
    On Error GoTo ErrTrap
    
    Me.Show
    Me.ZOrder
    dtpFecha1.value = gobjMain.EmpresaActual.GNOpcion.FechaInicio
    dtpFecha2.value = Date
         For i = 1 To IVGRUPO_MAX
             cboGrupo.AddItem gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(i)
         Next i
         If (numGrupo <= cboGrupo.ListCount) And (numGrupo > 0) Then
             cboGrupo.ListIndex = numGrupo - 1   'Selecciona lo anterior
         ElseIf cboGrupo.ListCount > 0 Then
             cboGrupo.ListIndex = 0              'Selecciona la primera
         End If
    
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
    Dim s As String
    'Carga la lista de transacción
'    fcbTrans.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "IV")

    lstTrans.Clear
    v = gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "IV")
    For i = LBound(v, 2) To UBound(v, 2)
        lstTrans.AddItem v(0, i)        '& " " & v(1, i)
    Next i
    
        'jeaa 25/09/206
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaRecosteo")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaRecosteo")
            RecuperaTrans "KeyT", lstTrans, s
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
    If ReprocCosto(False, (chkTodo.value = vbChecked)) Then
        cmdCancelar.SetFocus
    End If
End Sub

'Private Sub DebugColItems()
'    Dim s As Variant
'
'    For Each s In mColItems
'        Debug.Print s
'    Next s
'End Sub

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
    Dim s As String, tid As Long, i As Long, x As Single
    Dim gnc As GNComprobante, cambiado As Boolean, TransID  As Long
    
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
        x = grd.CellTop                 'Para visualizar la celda actual
'        If i = 11 Then MsgBox "hola"
        
        'Si es verificación procesa todas las filas sino solo las que tengan "Costo Incorrecto"
        If ((grd.TextMatrix(i, COL_RESULTADO) = MSG_NG) Or bandVerificar Or BandTodo) Then
            tid = grd.ValueMatrix(i, COL_TID)
            grd.TextMatrix(i, COL_RESULTADO) = "Verificando..."
            grd.Refresh
            
            'Recupera la transaccion
            
            Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
            If Not (gnc Is Nothing) Then
                'Si la transacción es de Inventario y es Egreso/Transferencia
                ' Y no está anulado
                If (gnc.GNTrans.Modulo = "IV") And _
                   (gnc.Estado <> ESTADO_ANULADO) Then
'                   (gnc.GNTrans.IVTipoTrans = "E" Or gnc.GNTrans.IVTipoTrans = "T") And _      '*** MAKOTO 06/sep/00 Eliminado
                    'Forzar recuperar todos los datos de transacción para que no se pierdan al grabar de nuveo
                    gnc.RecuperaDetalleTodo
                    'Recalcula costo de los items
                    If gobjMain.EmpresaActual.GNOpcion.IVKTipoDatoDouble Then
                        If RecalculoxItemDou(gnc, cambiado, bandVerificar) Then
                        'Si está cambiado algo
                            If cambiado Or BandTodo Then
                                'Si no es solo verificacion
                                If Not bandVerificar Then
                                    grd.TextMatrix(i, COL_RESULTADO) = "Grabando..."
                                    grd.Refresh
                                    'Prorratea los recargos/descuentos si los calcula en base a costo
                                    gnc.ProrratearIVKardexRecargo
                                    gnc.GeneraAsiento       'Diego 27 Abril 2001  corregido
                                    'Graba la transacción
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
                        If RecalculoxItem(gnc, cambiado, bandVerificar) Then
                            'Si está cambiado algo
                            If cambiado Or BandTodo Then
                                'Si no es solo verificacion
                                If Not bandVerificar Then
                                    grd.TextMatrix(i, COL_RESULTADO) = "Grabando..."
                                    grd.Refresh
                                    'Prorratea los recargos/descuentos si los calcula en base a costo
                                    gnc.ProrratearIVKardexRecargo
                                    gnc.GeneraAsiento       'Diego 27 Abril 2001  corregido
                                    'Graba la transacción
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


'Private Function Recalculo(ByVal gnc As GNComprobante, _
'                           ByRef cambiado As Boolean, _
'                           ByVal booVerificando As Boolean) As Boolean
'    Dim item As IVinventario, ivk As IVKardex, i As Long
'    Dim ct As Currency, ctotal As Currency, s As String
'    Dim CostoTotalEgreso As Currency
'    Dim ivkOUT As IVKardex, itemOUT As IVinventario
'    Dim CostoTotalPadre As Currency
'    Dim ItemMedio As Integer
'    Dim AcuCosto As Currency
'    Dim ctPrep As Currency
'    On Error GoTo ErrTrap
'        cambiado = False
'        ItemMedio = gnc.CountIVKardex / 2
'        For i = 1 To gnc.CountIVKardex
'        Set ivk = gnc.IVKardex(i)
'        Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
'        If gnc.GNTrans.codPantalla = "IVCAMIE" Then
'            'para que solo revise los items de egreso
'            If item.Tipo = CambioPresentacion Then
'               GoTo SiguienteItem
'            End If
'        ElseIf gnc.GNTrans.codPantalla = "IVCAMIEP" Then 'AUC para retro
'            If item.Tipo = CambioPresentacion Then
'                GoTo SiguienteItem
'            End If
'        End If
''        'Solo de salida
'            'Recupera el item
''            If ivk.CodInventario <> "7861074600618" Then GoTo SiguienteItem
'            Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
'
'            If Not (item Is Nothing) Then
'                '*** MAKOTO 31/ago/00
'                If booVerificando Then
'                    If ItemIncorrecto(item.CodInventario) Then      'Este item ya está marcado como incorrecto.
'                        Debug.Print "Incorrecto por trans. anterior. cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.NumTrans
'                        cambiado = True
'                        GoTo SiguienteItem
'                    End If
'                End If
'                '*** MAKOTO 08/dic/00
'                ct = item.CostoDouble2(gnc.FechaTrans, _
'                                       Abs(ivk.cantidad), _
'                                       gnc.TransID, _
'                                       gnc.HoraTrans)
'
'                'Convierte en moneda de la transaccion
'                If item.CodMoneda <> gnc.CodMoneda Then
'                    ct = ct * gnc.Cotizacion(item.CodMoneda) / gnc.Cotizacion("")
'                End If
'                ctotal = ct * ivk.cantidad
'                CostoTotalPadre = ctotal
'                'Si el costo es diferente de lo que está grabado
'                If ctotal <> ivk.CostoTotal Then
'                    If gnc.GNTrans.IVTipoTrans = "C" And i > 1 And gnc.GNTrans.codPantalla <> "IVCAMIE" Then
'                        Set ivkOUT = gnc.IVKardex(i - 1)
'                        Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
'                        If Not (itemOUT Is Nothing) Then
'                            '*** MAKOTO 31/ago/00
'                            If booVerificando Then
'                                If ItemIncorrecto(itemOUT.CodInventario) Then      'Este item ya está marcado como incorrecto.
'                                    cambiado = True
'                                    GoTo SiguienteItem
'                                End If
'                            End If
'                            ct = itemOUT.CostoDouble2(gnc.FechaTrans, _
'                                                   Abs(ivkOUT.cantidad), _
'                                                   gnc.TransID, _
'                                                   gnc.HoraTrans)
'
'                            'Convierte en moneda de la transaccion
'                            If itemOUT.CodMoneda <> gnc.CodMoneda Then
'                                ct = ct * gnc.Cotizacion(itemOUT.CodMoneda) / gnc.Cotizacion("")
'                            End If
'                            ctotal = ct * ivkOUT.cantidad * -1
'                            ivk.CostoTotal = ctotal
'                            ivk.CostoRealTotal = ctotal
'                            Set ivkOUT = Nothing
'                            Set itemOUT = Nothing
'                        End If
'                    ElseIf gnc.GNTrans.IVTipoTrans = "C" And i > 1 And gnc.GNTrans.codPantalla = "IVCAMIE" Then
'                        Set ivkOUT = gnc.IVKardex(i - (gnc.CountIVKardex / 2))
'                        Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
'                        If Not (itemOUT Is Nothing) Then
'                            '*** MAKOTO 31/ago/00
'                            If booVerificando Then
'                                mColItems.Add item:=item.CodInventario, Key:=item.CodInventario
'                                Debug.Print "Incorrecto 1 . cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.NumTrans
'                                Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotal
'                                ivk.CostoTotal = ctotal
'                                ivk.CostoRealTotal = ctotal
'
'
'                                If ItemIncorrecto(itemOUT.CodInventario) Then      'Este item ya está marcado como incorrecto.
'                                    cambiado = True
'                                    GoTo SiguienteItem
'                                End If
'                            End If
'                            ct = itemOUT.CostoDouble2(gnc.FechaTrans, _
'                                                   Abs(ivkOUT.cantidad), _
'                                                   gnc.TransID, _
'                                                   gnc.HoraTrans)
'
'                            'Convierte en moneda de la transaccion
'                            If itemOUT.CodMoneda <> gnc.CodMoneda Then
'                                ct = ct * gnc.Cotizacion(itemOUT.CodMoneda) / gnc.Cotizacion("")
'                            End If
'                            ctotal = ct * ivkOUT.cantidad * -1
'                            If ctotal <> CostoTotalPadre Then
'                                gnc.IVKardex(i - (gnc.CountIVKardex / 2)).CostoTotal = CostoTotalPadre * -1
'                                gnc.IVKardex(i - (gnc.CountIVKardex / 2)).CostoRealTotal = CostoTotalPadre * -1
'                                gnc.IVKardex(i).CostoTotal = CostoTotalPadre
'                                gnc.IVKardex(i).CostoRealTotal = CostoTotalPadre
'
'                                cambiado = True
''                                ivkOUT.CostoTotal = CostoTotalPadre
''                                ivkOUT.CostoRealTotal = ctotal
'                            End If
'
''                            ivk.CostoTotal = ctotal
''                            ivk.CostoRealTotal = ctotal
'                            Set ivkOUT = Nothing
'                            Set itemOUT = Nothing
'                        End If
'                  ElseIf gnc.GNTrans.IVTipoTrans = "C" And i > 1 And gnc.GNTrans.codPantalla = "IVCAMIEP" Then
'                        ivk.CostoTotal = ctotal
'                        ivk.CostoRealTotal = ctotal
'                        If ItemIncorrecto(item.CodInventario) Then      'Este item ya está marcado como incorrecto.
'                                mColItems.Add item:=item.CodInventario, Key:=item.CodInventario
'                                Debug.Print "Incorrecto 1 . cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.NumTrans
'                                Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotal
'                        End If
'                            Set ivkOUT = gnc.IVKardex(gnc.CountIVKardex - gnc.CountIVKardex + 1)
'
'                        Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
'                        ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
'                                       Abs(ivk.cantidad), _
'                                       gnc.TransID, _
'                                       gnc.HoraTrans)
'                        AcuCosto = AcuCosto + ctotal * -1
'                        ivkOUT.CostoTotal = AcuCosto 'ctotal * -1
'                        ivkOUT.CostoRealTotal = AcuCosto 'ctotal * -1
'
'                        'cambiado = True
'                        If AcuCosto <> ctPrep Then
'                             cambiado = True
'                        End If
'                        If Not ItemIncorrecto(itemOUT.CodInventario) Then      'Este item ya está marcado como incorrecto.
'                            mColItems.Add item:=itemOUT.CodInventario, Key:=itemOUT.CodInventario
'                            Debug.Print "Incorrecto 1 . cod='" & itemOUT.CodInventario & "' Trans=" & gnc.CodTrans & gnc.NumTrans
'                            Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotal
'                            GoTo SiguienteItem
'                        End If
'                        Set ivkOUT = Nothing
'                        Set itemOUT = Nothing
'                    Else
'                        '*** MAKOTO 31/ago/00
'                        If booVerificando Then
'                            'Almacena codigo de item para que de aquí en adelante todo marque como incorrecto.
'                            mColItems.Add item:=item.CodInventario, Key:=item.CodInventario
'                            Debug.Print "Incorrecto 1 . cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.NumTrans
'                            Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotal
'                        End If
'
'                        ivk.CostoTotal = ctotal
'                        ivk.CostoRealTotal = ctotal
'                        cambiado = True
'                        'jeaa 12/09/2005 recalculo de transformacion
'                        If gnc.GNTrans.IVTipoTrans = "C" Then
'                            If gnc.GNTrans.codPantalla = "IVCAMIE" Then
'
'                            ElseIf gnc.CountIVKardex = i + 1 Then
'                                    CostoTotalEgreso = ctotal * -1
'                            Else
'                                ivk.CostoTotal = CostoTotalEgreso
'                                ivk.CostoRealTotal = CostoTotalEgreso
'                            End If
'                        End If
'                    End If
'                Else
'                    'Esta parte es para cuando haya diferencia entre Costo y CostoReal
'                    ' en las transacciones que no debe tener diferencia.
'                    If (Not gnc.GNTrans.IVRecargoEnCosto) And (ivk.costo <> ivk.CostoReal) Then
'                        ivk.CostoRealTotal = ivk.CostoTotal
'                        cambiado = True
'
'                        '*** MAKOTO 31/ago/00
'                        If booVerificando Then
'                            'Almacena codigo de item para que de aquí en adelante todo marque como incorrecto.
'                            mColItems.Add item:=item.CodInventario, Key:=item.CodInventario
'                            Debug.Print "Incorrecto 2 Agregado. cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.NumTrans
'                        End If
'                    End If
'                End If
'  'aqui revisa el costo del item padre
'                    If gnc.GNTrans.IVTipoTrans = "C" And i > 1 And gnc.GNTrans.codPantalla = "IVCAMIE" Then
'                        ivk.CostoTotal = ctotal
'                        ivk.CostoRealTotal = ctotal
'                        If i > (gnc.CountIVKardex / 2) Then
'                            Set ivkOUT = gnc.IVKardex(Abs(i - (gnc.CountIVKardex / 2)))
'                        Else
'                            Set ivkOUT = gnc.IVKardex(Abs(i + (gnc.CountIVKardex / 2)))
'                        End If
'                        Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
'                        If Abs(ivkOUT.CostoTotal) <> Abs(ivk.CostoTotal) Then
'                            cambiado = True
'                            ivkOUT.CostoTotal = ctotal * -1
'                            ivkOUT.CostoRealTotal = ctotal * -1
'                            If Not ItemIncorrecto(itemOUT.CodInventario) Then      'Este item ya está marcado como incorrecto.
'                                mColItems.Add item:=itemOUT.CodInventario, Key:=itemOUT.CodInventario
'                                Debug.Print "Incorrecto 1 . cod='" & itemOUT.CodInventario & "' Trans=" & gnc.CodTrans & gnc.NumTrans
'                                  Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotal
'                                GoTo SiguienteItem
'                            End If
'                        End If
'                        Set ivkOUT = Nothing
'                        Set itemOUT = Nothing
'                      ElseIf gnc.GNTrans.IVTipoTrans = "C" And i > 1 And gnc.GNTrans.codPantalla = "IVCAMIEP" Then
'                        ivk.CostoTotal = ctotal 'AUC REPROCESO DE PREPARACIONES
'                        ivk.CostoRealTotal = ctotal
'                        Set ivkOUT = gnc.IVKardex(gnc.CountIVKardex - gnc.CountIVKardex + 1)
'                        Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
'                        ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
'                                       Abs(ivk.cantidad), _
'                                       gnc.TransID, _
'                                       gnc.HoraTrans)
'                        AcuCosto = AcuCosto + ctotal * -1
'                        ivkOUT.CostoTotal = AcuCosto 'ctotal * -1
'                        ivkOUT.CostoRealTotal = AcuCosto 'ctotal * -1
'                       If Abs(AcuCosto) <> Abs(ctPrep) Then
'                            cambiado = True
'                        Else
'                            cambiado = False
'                        End If
'                        Set ivkOUT = Nothing
'                        Set itemOUT = Nothing
'                    End If
'                  End If
'
'
'            'Si no puede recuperar el item
'            Else
'                'Aborta el recalculo
'                cambiado = False            'Para que no se grabe
'                GoTo salida
'            End If
''        End If                 '*** MAKOTO 06/sep/00
'SiguienteItem:
'
'    Next i
'
'SalidaOK:
'    Recalculo = True
'    GoTo salida
'    Exit Function
'ErrTrap:
'    DispErr
'salida:
'    Set ivk = Nothing
'    Set item = Nothing
'    Set gnc = Nothing
'    Exit Function
'End Function

Private Function ItemIncorrecto(ByVal cod As String) As Boolean
    Dim s As String
    
    On Error Resume Next
    s = mColItems.item(cod)     'Si es que encuentra en la coleccion,
    If Err.Number = 0 Then      'Este item ya está marcado como incorrecto.
        ItemIncorrecto = True
    End If
    On Error GoTo 0
End Function

Private Sub cmdBuscar_Click()
    Dim v As Variant, obj As Object, s As String
    On Error GoTo ErrTrap
    
    '*** MAKOTO 06/sep/00 Agregado
    If lstTrans.SelCount = 0 Then
        MsgBox "Seleccione una transacción, por favor.", vbInformation
        Exit Sub
    End If
    
    If Len(fcbDesde2.Text) = 0 Then
        MsgBox "Seleccione un Item, por favor.", vbInformation
        fcbDesde2.SetFocus
        Exit Sub
    End If
    
    
    With gobjMain.objCondicion
        .fecha1 = dtpFecha1.value
        .fecha2 = dtpFecha2.value
'        .CodTrans = fcbTrans.Text              '*** MAKOTO 31/ago/00 Modificado
        .CodTrans = PreparaCodTrans             '***
        .NumTrans1 = Val(txtNumTrans1.Text)
        .NumTrans2 = Val(txtNumTrans2.Text)
        
             If cboGrupo.ListIndex >= 0 Then
                 numGrupo = cboGrupo.ListIndex + 1
                 .Grupo1 = Trim$(fcbGrupoDesde.KeyText)
                 .Grupo2 = Trim$(fcbGrupoHasta.KeyText)
             End If
             .CodItem1 = Trim$(fcbDesde2.Text)
             .CodItem2 = Trim$(fcbHasta2.Text)
        
        
        'Estados no incluye anulados
        .EstadoBool(ESTADO_NOAPROBADO) = True
        .EstadoBool(ESTADO_APROBADO) = True
        .EstadoBool(ESTADO_DESPACHADO) = True
        .EstadoBool(ESTADO_ANULADO) = False
            'jeaa 25/09/06
        s = PreparaTransParaGnopcion(.CodTrans)
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransparaRecosteo", s
    'Graba en la base
    gobjMain.EmpresaActual.GNOpcion.Grabar

    End With
    
    If gobjMain.EmpresaActual.GNOpcion.IVKTipoDatoDouble Then
        Set obj = gobjMain.EmpresaActual.Empresa2.ConsGNTrans22Dou(True)      'Orden ascendente
    Else
        Set obj = gobjMain.EmpresaActual.ConsGNTrans22(True)  'Orden ascendente     '*** MAKOTO 20/oct/00
    End If
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
        .FormatString = "^#|tid|<Fecha|<Asiento|<Trans|<#|<#Ref.|<Nombre|<Descripción|>Costo rEAL Unitario|>Costo Unitario|<C.Costo|<Estado|<Resultado"
        .ColHidden(COL_NUMFILA) = False
        .ColHidden(COL_TID) = True
        .ColHidden(COL_FECHA) = False
        .ColHidden(COL_CODASIENTO) = True
        .ColHidden(COL_CODTRANS) = False
        .ColHidden(COL_NUMTRANS) = False
        .ColHidden(COL_NUMDOCREF) = True
        .ColHidden(COL_NOMBRE) = False  'True
        .ColHidden(COL_DESC) = False
        .ColHidden(COL_CUR) = False
        .ColHidden(COL_CU) = False
        .ColHidden(COL_CENTROCOSTO) = True
        .ColHidden(COL_ESTADO) = True
        
        .ColDataType(COL_FECHA) = flexDTDate    '*** MAKOTO 14/ago/2000 para que ordene bien por fecha
        .ColDataType(COL_CU) = flexDTCurrency '*** MAKOTO 14/ago/2000 para que ordene bien por fecha
        .ColFormat(COL_CU) = "##,0.0000"
'*** MAKOTO 20/oct/00 Eliminado
'        'Ordena por fecha ascendente      '*** MAKOTO 07/oct/00 Agregado por que cambió el orden del método
'        .col = COL_FECHA
'        .Sort = flexSortGenericAscending
        
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


Private Sub cmdCorregirIVA_Click()
    Dim ListaTrans As String
    ListaTrans = PreparaCodTrans
    frmReprocIVA.Inicio ListaTrans
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
    
    If ReprocCosto(True, False) Then
        cmdAceptar.Enabled = True
        cmdAceptar.SetFocus
        mVerificado = True
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

Private Sub fcbGrupoDesde_Selected(ByVal Text As String, ByVal KeyText As String)
    fcbGrupoHasta.KeyText = fcbGrupoDesde.KeyText   '*** MAKOTO 27/jun/2000
End Sub

Private Sub cboGrupo_Click()
    Dim Numg As Integer
    On Error GoTo ErrTrap
    If cboGrupo.ListIndex < 0 Then Exit Sub

    'MensajeStatus MSG_PREPARA, vbHourglass

    Numg = cboGrupo.ListIndex + 1
    fcbGrupoDesde.SetData gobjMain.EmpresaActual.ListaIVGrupo(Numg, False, False)
    fcbGrupoHasta.SetData fcbGrupoDesde.GetData             '*** MAKOTO 19/feb/01 Mod.
    fcbGrupoDesde.KeyText = ""
    fcbGrupoHasta.KeyText = ""
    CargaItems
    'MensajeStatus
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    Exit Sub
End Sub


Private Sub fcbDesde2_Selected(ByVal Text As String, ByVal KeyText As String)
    fcbHasta2.KeyText = fcbDesde2.KeyText   '*** MAKOTO 27/jun/2000
End Sub

Private Sub fcbGrupoDesde_Validate(Cancel As Boolean)
    'Carga Items
    CargaItems

End Sub



Private Sub fcbGrupoHasta_Validate(Cancel As Boolean)
    'Carga Items
    CargaItems

End Sub

Private Sub CargaItems()
    Dim numGrupo As Integer, v() As Variant
    Dim sql  As String, rs As Recordset, cond As String
    numGrupo = cboGrupo.ListIndex + 1
    fcbDesde2.Clear
    fcbHasta2.Clear
    If Len(fcbGrupoDesde.Text) > 0 And Len(fcbGrupoHasta.Text) > 0 Then
        cond = " WHERE codGrupo" & numGrupo & " BETWEEN '" & _
                fcbGrupoDesde.Text & "' AND '" & fcbGrupoHasta.Text & "'"
    End If
    sql = "SELECT CodInventario, IVInventario.Descripcion FROM IVInventario " & _
    IIf(Len(fcbGrupoDesde.Text) > 0 And Len(fcbGrupoHasta.Text) > 0, " INNER JOIN IVGrupo" & numGrupo & _
           " ON IVInventario.IdGrupo" & numGrupo & " = IVGrupo" & numGrupo & ".IdGrupo" & numGrupo & cond, "")
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not rs.EOF Then
        v = MiGetRows(rs)
        fcbDesde2.SetData v
        fcbHasta2.SetData v
    End If
    fcbDesde2.Text = ""
    fcbHasta2.Text = ""
End Sub
Private Function RecalculoxItem(ByVal gnc As GNComprobante, _
                           ByRef cambiado As Boolean, _
                           ByVal booVerificando As Boolean) As Boolean
    Dim item As IVinventario, ivk As IVKardex, i As Long
    Dim ct As Currency, ctotal As Currency, s As String
    Dim CostoTotalEgreso As Currency
    Dim ivkOUT As IVKardex, itemOUT As IVinventario
    Dim CostoTotalPadre As Currency
    Dim itemAux As IVinventario
    Dim ctOut As Currency, ctotalOut As Currency
    Dim acuCosto As Currency
    Dim ctPrep As Currency
    On Error GoTo ErrTrap
    cambiado = False
    For i = 1 To gnc.CountIVKardex
        Set ivk = gnc.IVKardex(i)
        If gnc.GNTrans.CodPantalla <> "IVCAMIEP" Then
                If ivk.CodInventario <> fcbDesde2.KeyText Then GoTo SiguienteItem
        End If
         Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
        If gnc.GNTrans.CodPantalla = "IVCAMIE" Or gnc.GNTrans.CodPantalla = "IVCAMIEP" Then
           If item.Tipo = CambioPresentacion Then
                GoTo SiguienteItem
            End If
        End If
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
                ct = item.CostoDouble2(gnc.FechaTrans, _
                                       Abs(ivk.cantidad), _
                                       gnc.TransID, _
                                       gnc.HoraTrans)
                'Convierte en moneda de la transaccion
                If item.CodMoneda <> gnc.CodMoneda Then
                    ct = ct * gnc.Cotizacion(item.CodMoneda) / gnc.Cotizacion("")
                End If
                If gnc.GNTrans.CodPantalla = "IVCAMIEP" Then 'Porque cuando es una egreso en la transformacion debe ser negativo
                    ctotal = Abs(ct) * ivk.cantidad
                Else
                    ctotal = ct * ivk.cantidad
                End If
                CostoTotalPadre = ctotal
                'Si el costo es diferente de lo que está grabado
                If ctotal <> ivk.CostoTotal Or ctotal <> ivk.CostoRealTotal Then
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
                    ElseIf gnc.GNTrans.IVTipoTrans = "C" And i > 1 And gnc.GNTrans.CodPantalla = "IVCAMIE" Then
                        ivk.CostoTotal = ctotal
                        ivk.CostoRealTotal = ctotal
'                        cambiado = True
'                            If ItemIncorrecto(item.CodInventario) Then      'Este item ya está marcado como incorrecto.
'                                mColItems.Add item:=item.CodInventario, Key:=item.CodInventario
'                                Debug.Print "Incorrecto 1 . cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
'                                Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotal
'                           End If
                            If i > (gnc.CountIVKardex / 2) Then
                                Set ivkOUT = gnc.IVKardex(Abs(i - (gnc.CountIVKardex / 2)))
                            Else
                                Set ivkOUT = gnc.IVKardex(Abs(i + (gnc.CountIVKardex / 2)))
                            End If
                            Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                            ivkOUT.CostoTotal = ctotal * -1
                            ivkOUT.CostoRealTotal = ctotal * -1
                            cambiado = True
                            If Not ItemIncorrecto(itemOUT.CodInventario) Then      'Este item ya está marcado como incorrecto.
                                mColItems.Add item:=itemOUT.CodInventario, Key:=itemOUT.CodInventario
                                Debug.Print "Incorrecto 1 . cod='" & itemOUT.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                                Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotal
                                GoTo SiguienteItem
                            End If
                            Set ivkOUT = Nothing
                            Set itemOUT = Nothing
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
                    Else
                        '*** MAKOTO 31/ago/00
                        If booVerificando Then
                            'Almacena codigo de item para que de aquí en adelante todo marque como incorrecto.
                            mColItems.Add item:=item.CodInventario, Key:=item.CodInventario
                            Debug.Print "Incorrecto 1 . cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                            Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotal
                        End If
                        ivk.CostoTotal = ctotal
                        ivk.CostoRealTotal = ctotal
                        cambiado = True
                        'jeaa 12/09/2005 recalculo de transformacion
                        If gnc.GNTrans.IVTipoTrans = "C" Then
                            If gnc.GNTrans.CodPantalla = "IVCAMIE" Then
                            ElseIf gnc.CountIVKardex = i + 1 Then
                                    CostoTotalEgreso = ctotal * -1
                            Else
                                ivk.CostoTotal = CostoTotalEgreso
                                ivk.CostoRealTotal = CostoTotalEgreso
                            End If
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
                            mColItems.Add item:=item.CodInventario, Key:=item.CodInventario
                            Debug.Print "Incorrecto 2 Agregado. cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                        End If
                    End If
                    If gnc.GNTrans.IVTipoTrans = "C" And i > 1 And gnc.GNTrans.CodPantalla = "IVCAMIE" Then
                        ivk.CostoTotal = ctotal
                        ivk.CostoRealTotal = ctotal
                        If i > (gnc.CountIVKardex / 2) Then
                            Set ivkOUT = gnc.IVKardex(Abs(i - (gnc.CountIVKardex / 2)))
                        Else
                            Set ivkOUT = gnc.IVKardex(Abs(i + (gnc.CountIVKardex / 2)))
                        End If
                        Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                        If Abs(ivkOUT.CostoTotal) <> Abs(ivk.CostoTotal) Then
                            cambiado = True
                            ivkOUT.CostoTotal = ctotal * -1
                            ivkOUT.CostoRealTotal = ctotal * -1
                            If Not ItemIncorrecto(itemOUT.CodInventario) Then      'Este item ya está marcado como incorrecto.
                                mColItems.Add item:=itemOUT.CodInventario, Key:=itemOUT.CodInventario
                                Debug.Print "Incorrecto 1 . cod='" & itemOUT.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                                Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotal
                                GoTo SiguienteItem
                            End If
                        End If
                        Set ivkOUT = Nothing
                        Set itemOUT = Nothing
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
    RecalculoxItem = True
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

Private Function VerificaItemBuscado(ByVal coditem As String) As Boolean
    Dim sql As String, rs As Recordset
    sql = "select * from ivinventario"
    If Len(fcbDesde2.Text) > 0 And Len(fcbHasta2.Text) > 0 Then
        sql = sql & " where codivinventario between '" & fcbDesde2.Text & "' and '" & fcbDesde2.Text & "'"
    End If
    
    If Len(fcbGrupoDesde.Text) > 0 And Len(fcbHasta2.Text) > 0 Then
        sql = sql & " where codivinventario='" & coditem & "'"
    End If
    

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
'jeaa 25/09/2006 elimina los apostrofes

Private Function PreparaTransParaGnopcion(cad As String) As String
    Dim v As Variant, i As Integer, s As String
    s = ""
    v = Split(cad, ",")
    For i = 0 To UBound(v)
        v(i) = Trim(v(i))
        s = s & Mid$(v(i), 2, Len(v(i)) - 2) & ","
    Next i
    'quita ultima coma
    PreparaTransParaGnopcion = Mid$(s, 1, Len(s) - 1)
End Function

Private Sub chkTodo_Click()
    If chkTodo.value = vbChecked Then
        cmdVerificar.Enabled = False
        cmdAceptar.Enabled = (grd.Rows > grd.FixedRows)
    Else
        cmdVerificar.Enabled = Not mVerificado
        cmdAceptar.Enabled = mVerificado
    End If
End Sub

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



Private Function RecalculoxItemDou(ByVal gnc As GNComprobante, _
                           ByRef cambiado As Boolean, _
                           ByVal booVerificando As Boolean) As Boolean
    Dim item As IVinventario, ivk As IVKardex, i As Long
    Dim ct As Double, ctotal As Double, s As String
    Dim CostoTotalEgreso As Double
    Dim ivkOUT As IVKardex, itemOUT As IVinventario
    Dim CostoTotalPadre As Double
    Dim itemAux As IVinventario
    Dim ctOut As Double, ctotalOut As Double
    Dim acuCosto As Double
    Dim ctPrep As Double
    On Error GoTo ErrTrap
    cambiado = False
    For i = 1 To gnc.CountIVKardex
        Set ivk = gnc.IVKardex(i)
        If gnc.GNTrans.CodPantalla <> "IVCAMIEP" Then
                If ivk.CodInventario <> fcbDesde2.KeyText Then GoTo SiguienteItem
        End If
         Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
        If gnc.GNTrans.CodPantalla = "IVCAMIE" Or gnc.GNTrans.CodPantalla = "IVCAMIEP" Then
           If item.Tipo = CambioPresentacion Then
                GoTo SiguienteItem
            End If
        End If
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
                ct = item.CostoDouble2(gnc.FechaTrans, _
                                       Abs(ivk.CantidadDou), _
                                       gnc.TransID, _
                                       gnc.HoraTrans)
                'Convierte en moneda de la transaccion
                If item.CodMoneda <> gnc.CodMoneda Then
                    ct = ct * gnc.Cotizacion(item.CodMoneda) / gnc.Cotizacion("")
                End If
                If gnc.GNTrans.CodPantalla = "IVCAMIEP" Then 'Porque cuando es una egreso en la transformacion debe ser negativo
                    ctotal = Abs(ct) * ivk.CantidadDou
                Else
                    ctotal = ct * ivk.CantidadDou
                End If
                CostoTotalPadre = ctotal
                'Si el costo es diferente de lo que está grabado
                If ctotal <> ivk.CostoTotaldou Or ctotal <> ivk.CostoRealTotaldou Then
                    If gnc.GNTrans.IVTipoTrans = "C" And i > 1 And gnc.GNTrans.CodPantalla <> "IVCAMIE" And gnc.GNTrans.CodPantalla <> "IVCAMIEP" Then
                        Set ivkOUT = gnc.IVKardex(i - 1)
                        Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                        If Not (itemOUT Is Nothing) Then
                            If booVerificando Then
                                If ItemIncorrecto(itemOUT.CodInventario) Then      'Este item ya está marcado como incorrecto.
                                    cambiado = True
                                    GoTo SiguienteItem
                                End If
                            End If
                            ct = itemOUT.CostoDouble2(gnc.FechaTrans, _
                                                   Abs(ivkOUT.CantidadDou), _
                                                   gnc.TransID, _
                                                   gnc.HoraTrans)
                            'Convierte en moneda de la transaccion
                            If itemOUT.CodMoneda <> gnc.CodMoneda Then
                                ct = ct * gnc.Cotizacion(itemOUT.CodMoneda) / gnc.Cotizacion("")
                            End If
                            ctotal = ct * ivkOUT.CantidadDou * -1
                            ivk.CostoTotaldou = ctotal
                            ivk.CostoRealTotaldou = ctotal
                            Set ivkOUT = Nothing
                            Set itemOUT = Nothing
                        End If
                    ElseIf gnc.GNTrans.IVTipoTrans = "R" And i > 1 Then
                        ivk.CostoTotaldou = ctotal
                        ivk.CostoRealTotaldou = ctotal
                            If i > (gnc.CountIVKardex / 2) Then
                                Set ivkOUT = gnc.IVKardex(Abs(i - (gnc.CountIVKardex / 2)))
                            Else
                                Set ivkOUT = gnc.IVKardex(Abs(i + (gnc.CountIVKardex / 2)))
                            End If
                            Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                            ivkOUT.CostoTotaldou = ctotal * -1
                            ivkOUT.CostoRealTotaldou = ctotal * -1
                            cambiado = True
                            If Not ItemIncorrecto(itemOUT.CodInventario) Then      'Este item ya está marcado como incorrecto.
                                mColItems.Add item:=itemOUT.CodInventario, Key:=itemOUT.CodInventario
                                Debug.Print "Incorrecto 1 . cod='" & itemOUT.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                                Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotaldou
                                GoTo SiguienteItem
                            End If
                            Set ivkOUT = Nothing
                            Set itemOUT = Nothing
                    ElseIf gnc.GNTrans.IVTipoTrans = "C" And i > 1 And gnc.GNTrans.CodPantalla = "IVCAMIE" Then
                        ivk.CostoTotaldou = ctotal
                        ivk.CostoRealTotaldou = ctotal
                            If i > (gnc.CountIVKardex / 2) Then
                                Set ivkOUT = gnc.IVKardex(Abs(i - (gnc.CountIVKardex / 2)))
                            Else
                                Set ivkOUT = gnc.IVKardex(Abs(i + (gnc.CountIVKardex / 2)))
                            End If
                            Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                            ivkOUT.CostoTotaldou = ctotal * -1
                            ivkOUT.CostoRealTotaldou = ctotal * -1
                            cambiado = True
                            If Not ItemIncorrecto(itemOUT.CodInventario) Then      'Este item ya está marcado como incorrecto.
                                mColItems.Add item:=itemOUT.CodInventario, Key:=itemOUT.CodInventario
                                Debug.Print "Incorrecto 1 . cod='" & itemOUT.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                                Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotaldou
                                GoTo SiguienteItem
                            End If
                            Set ivkOUT = Nothing
                            Set itemOUT = Nothing
                    ElseIf gnc.GNTrans.IVTipoTrans = "C" And gnc.GNTrans.CodPantalla = "IVCAMIEP" Then
                        ivk.CostoTotaldou = ctotal 'AUC REPROCESO DE PREPARACIONES
                        ivk.CostoRealTotaldou = ctotal
                        Set ivkOUT = gnc.IVKardex(gnc.CountIVKardex - gnc.CountIVKardex + 1)
                        Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                        ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
                                       Abs(ivk.CantidadDou), _
                                       gnc.TransID, _
                                       gnc.HoraTrans)
                        acuCosto = acuCosto + ctotal * -1
                        ivkOUT.CostoTotaldou = acuCosto 'ctotal * -1
                        ivkOUT.CostoRealTotaldou = acuCosto 'ctotal * -1
                       If Abs(acuCosto) <> Abs(ctPrep) Then
                            cambiado = True
                        Else
                            cambiado = False
                        End If
                        Set ivkOUT = Nothing
                        Set itemOUT = Nothing
                    Else
                        If booVerificando Then
                            'Almacena codigo de item para que de aquí en adelante todo marque como incorrecto.
                            mColItems.Add item:=item.CodInventario, Key:=item.CodInventario
                            Debug.Print "Incorrecto 1 . cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                            Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotaldou
                        End If
                        ivk.CostoTotaldou = ctotal
                        ivk.CostoRealTotaldou = ctotal
                        cambiado = True
                        If gnc.GNTrans.IVTipoTrans = "C" Or gnc.GNTrans.IVTipoTrans = "R" Then
                            If gnc.GNTrans.CodPantalla = "IVCAMIE" Then
                            ElseIf gnc.CountIVKardex = i + 1 Then
                                    CostoTotalEgreso = ctotal * -1
                            Else
                                ivk.CostoTotaldou = CostoTotalEgreso
                                ivk.CostoRealTotaldou = CostoTotalEgreso
                            End If
                        End If
                    End If
                Else
                    'Esta parte es para cuando haya diferencia entre Costo y CostoReal
                    ' en las transacciones que no debe tener diferencia.
                    If (Not gnc.GNTrans.IVRecargoEnCosto) And (ivk.costoDou <> ivk.CostoRealDou) Then
                        ivk.CostoRealTotaldou = ivk.CostoTotaldou
                        cambiado = True
                        If booVerificando Then
                            'Almacena codigo de item para que de aquí en adelante todo marque como incorrecto.
                            mColItems.Add item:=item.CodInventario, Key:=item.CodInventario
                            Debug.Print "Incorrecto 2 Agregado. cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                        End If
                    End If
                    If gnc.GNTrans.IVTipoTrans = "C" And i > 1 And gnc.GNTrans.CodPantalla = "IVCAMIE" Then
                        ivk.CostoTotaldou = ctotal
                        ivk.CostoRealTotaldou = ctotal
                        If i > (gnc.CountIVKardex / 2) Then
                            Set ivkOUT = gnc.IVKardex(Abs(i - (gnc.CountIVKardex / 2)))
                        Else
                            Set ivkOUT = gnc.IVKardex(Abs(i + (gnc.CountIVKardex / 2)))
                        End If
                        Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                        If Abs(ivkOUT.CostoTotaldou) <> Abs(ivk.CostoTotaldou) Then
                            cambiado = True
                            ivkOUT.CostoTotaldou = ctotal * -1
                            ivkOUT.CostoRealTotaldou = ctotal * -1
                            If Not ItemIncorrecto(itemOUT.CodInventario) Then      'Este item ya está marcado como incorrecto.
                                mColItems.Add item:=itemOUT.CodInventario, Key:=itemOUT.CodInventario
                                Debug.Print "Incorrecto 1 . cod='" & itemOUT.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                                Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotaldou
                                GoTo SiguienteItem
                            End If
                        End If
                        Set ivkOUT = Nothing
                        Set itemOUT = Nothing
                    'AUC AQUI DEBERIA IR EL REPROCESO PARA LA TRANSFORMACION cuando los costos son iguales
                    ElseIf gnc.GNTrans.IVTipoTrans = "C" And gnc.GNTrans.CodPantalla = "IVCAMIEP" Then
                       ivk.CostoTotaldou = ctotal 'AUC REPROCESO DE PREPARACIONES
                        ivk.CostoRealTotaldou = ctotal
                        Set ivkOUT = gnc.IVKardex(gnc.CountIVKardex - gnc.CountIVKardex + 1)
                        Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                        ctPrep = itemOUT.CostoDouble2(gnc.FechaTrans, _
                                       Abs(ivk.CantidadDou), _
                                       gnc.TransID, _
                                       gnc.HoraTrans)
                        acuCosto = acuCosto + ctotal * -1
                        ivkOUT.CostoTotaldou = acuCosto 'ctotal * -1
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
SiguienteItem:
    Next i
SalidaOK:
    RecalculoxItemDou = True
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
