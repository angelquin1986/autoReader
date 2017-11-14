VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReprocCostoPeriodo 
   Caption         =   "Reprocesamiento de costos x periodo"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6510
   ScaleWidth      =   7920
   WindowState     =   2  'Maximized
   Begin VB.Frame frmBodega 
      Caption         =   "Calcular costos x Bodega"
      Height          =   855
      Left            =   165
      TabIndex        =   18
      Top             =   1785
      Width           =   6960
      Begin VB.CheckBox chkActXBodega 
         Caption         =   "Act. costos en items que no estan bodega selecionada"
         Height          =   390
         Left            =   2610
         TabIndex        =   21
         Top             =   270
         Width           =   4185
      End
      Begin FlexComboProy.FlexCombo fcbBodega 
         Height          =   330
         Left            =   810
         TabIndex        =   19
         ToolTipText     =   "Responsable de la transacción"
         Top             =   300
         Width           =   1695
         _ExtentX        =   2990
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Bodega"
         Height          =   195
         Left            =   165
         TabIndex        =   20
         Top             =   300
         Width           =   585
      End
   End
   Begin VB.Frame fraFecha 
      Caption         =   "&Fecha (desde - hasta)"
      Height          =   1710
      Left            =   165
      TabIndex        =   0
      Top             =   45
      Width           =   2052
      Begin MSComCtl2.DTPicker dtpFecha1 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   375
         Width           =   1695
         _ExtentX        =   2990
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
         Format          =   106692609
         CurrentDate     =   36348
      End
      Begin MSComCtl2.DTPicker dtpFecha2 
         Height          =   330
         Left            =   180
         TabIndex        =   2
         Top             =   870
         Width           =   1695
         _ExtentX        =   2990
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
         Format          =   106692609
         CurrentDate     =   36348
      End
   End
   Begin VB.Frame fraCodTrans 
      Caption         =   "Cod.&Trans"
      Height          =   1710
      Left            =   2235
      TabIndex        =   3
      Top             =   45
      Width           =   3000
      Begin VB.CommandButton cmdTransLimpiar 
         Caption         =   "Limp."
         Height          =   375
         Left            =   1515
         TabIndex        =   6
         Top             =   1320
         Width           =   1380
      End
      Begin VB.CommandButton cmdTransTodo 
         Caption         =   "Todo egresos"
         Height          =   375
         Left            =   90
         TabIndex        =   5
         Top             =   1320
         Width           =   1425
      End
      Begin VB.ListBox lstTrans 
         Columns         =   3
         Height          =   1065
         IntegralHeight  =   0   'False
         Left            =   90
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   240
         Width           =   2805
      End
   End
   Begin VB.Frame fraNumTrans 
      Caption         =   "#T&rans. (desde-hasta)"
      Height          =   1725
      Left            =   5280
      TabIndex        =   7
      Top             =   30
      Width           =   1830
      Begin VB.TextBox txtNumTrans1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   8
         Top             =   330
         Width           =   1305
      End
      Begin VB.TextBox txtNumTrans2 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   195
         TabIndex        =   9
         Top             =   765
         Width           =   1290
      End
   End
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   852
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   7920
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5655
      Width           =   7920
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
         TabIndex        =   15
         Top             =   0
         Width           =   1212
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   4995
         TabIndex        =   14
         Top             =   0
         Width           =   1212
      End
      Begin VB.CommandButton cmdVerificar 
         Caption         =   "&Verificar"
         Enabled         =   0   'False
         Height          =   372
         Left            =   255
         TabIndex        =   13
         Top             =   0
         Width           =   1212
      End
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   16
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
      TabIndex        =   11
      Top             =   3120
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
      Height          =   330
      Left            =   2340
      TabIndex        =   10
      Top             =   2730
      Width           =   2655
   End
End
Attribute VB_Name = "frmReprocCostoPeriodo"
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

'*** MAKOTO 31/ago/00 Agregado
'       para almacenar items con costo incorrecto detectado
Private mColItems As Collection



Public Sub Inicio()
    Dim i As Integer
    On Error GoTo ErrTrap
    
    Me.Show
    Me.ZOrder
    dtpFecha1.value = gobjMain.EmpresaActual.GNOpcion.FechaInicio
    dtpFecha2.value = Date
    FcbBodega.SetData gobjMain.EmpresaActual.ListaIVBodega(True, False)
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
    
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("BodegaparaRecosteoPeriodo")) > 0 Then
        s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("BodegaparaRecosteoPeriodo")
        FcbBodega.KeyText = s
    End If
    
    
    'jeaa 25/09/206
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaRecosteoPeriodo")) > 0 Then
        s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaRecosteoperiodo")
        RecuperaTrans "KeyT", lstTrans, s
    End If
    
End Sub



Private Sub cmdAceptar_Click()
    'Si no hay transacciones
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna transacción para procesar.", vbExclamation
        Exit Sub
    End If
    
    If ReprocCosto(False) Then
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

Private Function ReprocCosto(ByVal bandVerificar As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, x As Single
    Dim gnc As GNComprobante, cambiado As Boolean
    Dim BandRecalculo As Boolean
    
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
        
        'Si es verificación procesa todas las filas sino solo las que tengan "Costo Incorrecto"
        If ((grd.TextMatrix(i, COL_RESULTADO) = MSG_NG) Or bandVerificar) Then
        
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
                
                    If gnc.GNTrans.CodPantalla = "IVPROD" Then
                        BandRecalculo = RecalculoProduccion(gnc, cambiado, bandVerificar)
                    Else
                        BandRecalculo = RecalculoProduccion(gnc, cambiado, bandVerificar)
                    End If
                
                    'Recalcula costo de los items
                    If BandRecalculo Then
                        'Si está cambiado algo
                        If cambiado Then
                            'Si no es solo verificacion
                            If Not bandVerificar Then
                                grd.TextMatrix(i, COL_RESULTADO) = "Grabando..."
                                grd.Refresh
                                
                                'Prorratea los recargos/descuentos si los calcula en base a costo
                                gnc.ProrratearIVKardexRecargo
                                gnc.GeneraAsiento       'Diego 27 Abril 2001  corregido
                                'Graba la transacción
                                gnc.Grabar False, False
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


'''Private Function Recalculo(ByVal gnc As GNComprobante, _
'''                           ByRef cambiado As Boolean, _
'''                           ByVal booVerificando As Boolean) As Boolean
'''    Dim item As IVinventario, ivk As IVKardex, i As Long
'''    Dim ct As Currency, ctotal As Currency, s As String
'''
'''
'''    On Error GoTo Errtrap
'''    cambiado = False
'''    For i = 1 To gnc.CountIVKardex
'''        Set ivk = gnc.IVKardex(i)
'''
''''        'Solo de salida
''''        If ivk.Cantidad < 0 Then           '*** MAKOTO 06/sep/00
'''            'Recupera el item
'''            Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
'''            'jeaa 25/06/2007 por que es item de ingreso y no debe recalcular el costo
'''           If gnc.IVKardex(i).Cantidad > 0 Then
'''                GoTo SiguienteItem
'''           End If
'''
'''
'''            If Not (item Is Nothing) Then
'''                '*** MAKOTO 31/ago/00
'''                If booVerificando Then
'''                    If ItemIncorrecto(item.CodInventario) Then      'Este item ya está marcado como incorrecto.
'''                        Debug.Print "Incorrecto por trans. anterior. cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
'''                        cambiado = True
'''                        GoTo SiguienteItem
'''                    End If
'''                End If
'''
''''                ct = item.Costo(gnc.FechaTrans, Abs(ivk.Cantidad), gnc.TransID)
'''                '*** MAKOTO 08/dic/00
'''
'''                ct = item.CostoPromXPeriodo(gobjMain.objCondicion.Fecha1, gobjMain.objCondicion.Fecha2, gobjMain.objCondicion.CodBodega1)
'''
'''                'Convierte en moneda de la transaccion
'''                If item.CodMoneda <> gnc.CodMoneda Then
'''                    ct = ct * gnc.Cotizacion(item.CodMoneda) / gnc.Cotizacion("")
'''                End If
'''                ctotal = ct * ivk.Cantidad
'''
'''                'Si el costo es diferente de lo que está grabado
'''
'''                If (ctotal <> ivk.CostoTotal) Then
'''                    If ((chkActXBodega.value = vbUnchecked) And (UCase(gobjMain.objCondicion.CodBodega1) = UCase(ivk.CodBodega))) Or (chkActXBodega.value = Checked) Then
'''                    '*** MAKOTO 31/ago/00
'''                        If booVerificando Then
'''                            'Almacena codigo de item para que de aquí en adelante todo marque como incorrecto.
'''                            mColItems.Add item:=item.CodInventario, Key:=item.CodInventario
'''                            Debug.Print "Incorrecto 1 . cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
'''                            Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotal
'''                        End If
'''
'''                        ivk.CostoTotal = ctotal
'''                        ivk.CostoRealTotal = ctotal
'''                        cambiado = True
'''                    End If
'''                Else
'''                    'Esta parte es para cuando haya diferencia entre Costo y CostoReal
'''                    ' en las transacciones que no debe tener diferencia.
'''                    If (Not gnc.GNTrans.IVRecargoEnCosto) And (ivk.costo <> ivk.CostoReal) Then
'''                        ivk.CostoRealTotal = ivk.CostoTotal
'''                        cambiado = True
'''
'''                        '*** MAKOTO 31/ago/00
'''                        If booVerificando Then
'''                            'Almacena codigo de item para que de aquí en adelante todo marque como incorrecto.
'''                            mColItems.Add item:=item.CodInventario, Key:=item.CodInventario
'''                            Debug.Print "Incorrecto 2 Agregado. cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
'''                        End If
'''                    End If
'''                End If
'''
'''            'Si no puede recuperar el item
'''            Else
'''                'Aborta el recalculo
'''                cambiado = False            'Para que no se grabe
'''                GoTo salida
'''            End If
''''        End If                 '*** MAKOTO 06/sep/00
'''SiguienteItem:
'''    Next i
'''
'''SalidaOK:
'''    Recalculo = True
'''    GoTo salida
'''    Exit Function
'''Errtrap:
'''    DispErr
'''salida:
'''    Set ivk = Nothing
'''    Set item = Nothing
'''    Set gnc = Nothing
'''    Exit Function
'''End Function

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
        lstTrans.SetFocus
        Exit Sub
    End If
    
    With gobjMain.objCondicion
        .fecha1 = dtpFecha1.value
        .fecha2 = dtpFecha2.value
        .CodTrans = PreparaCodTrans             '***
        .NumTrans1 = Val(txtNumTrans1.Text)
        .NumTrans2 = Val(txtNumTrans2.Text)
        
        .CodBodega1 = FcbBodega.KeyText         '*** Asigna bodega para usarlo caundo necesita extraer el cost x periodo
        
        'Estados no incluye anulados
        .EstadoBool(ESTADO_NOAPROBADO) = True
        .EstadoBool(ESTADO_APROBADO) = True
        .EstadoBool(ESTADO_DESPACHADO) = True
        .EstadoBool(ESTADO_ANULADO) = False
        s = PreparaTransParaGnopcion(.CodTrans)
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransparaRecosteoPeriodo", s
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "BodegaparaRecosteoPeriodo", FcbBodega.KeyText
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
    
    If ReprocCosto(True) Then
        cmdAceptar.Enabled = True
        cmdAceptar.SetFocus
    End If
End Sub


Private Sub fcbBodega_Selected(ByVal Text As String, ByVal KeyText As String)
    gobjMain.objCondicion.CodBodega1 = FcbBodega.KeyText         '*** Asigna bodega para usarlo caundo necesita extraer el cost x periodo
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

Private Function RecalculoProduccion(ByVal gnc As GNComprobante, _
                           ByRef cambiado As Boolean, _
                           ByVal booVerificando As Boolean) As Boolean
    Dim item As IVinventario, ivk As IVKardex, i As Long
    Dim ct As Currency, ctotal As Currency, s As String
    Dim CostoTotalEgreso As Currency
    Dim ivkOUT As IVKardex, itemOUT As IVinventario
    Dim CostoTotalPadre As Currency, j As Long
    On Error GoTo ErrTrap
    
    cambiado = False
    For i = 1 To gnc.CountIVKardex
        Set ivk = gnc.IVKardex(i)
        
           'Recupera el item
            
            Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
            
''''            If ivk.CodBodega <> fcbBodega.KeyText Then GoTo SiguienteItem
            
            If Not (item Is Nothing) Then
                '*** MAKOTO 31/ago/00
                If booVerificando Then
                    If ItemIncorrecto(item.CodInventario) Then      'Este item ya está marcado como incorrecto.
                        Debug.Print "Incorrecto por trans. anterior. cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                        cambiado = True
                        GoTo SiguienteItem
                    End If
                End If
                
                
                ct = item.CostoPromXPeriodo(gobjMain.objCondicion.fecha1, gobjMain.objCondicion.fecha2, gobjMain.objCondicion.CodBodega1)
                
                
                'Convierte en moneda de la transaccion
                If item.CodMoneda <> gnc.CodMoneda Then
                    ct = ct * gnc.Cotizacion(item.CodMoneda) / gnc.Cotizacion("")
                End If
                ctotal = ct * ivk.cantidad
                CostoTotalPadre = ctotal
                
                If ivk.cantidad > 0 Then
                    CostoTotalEgreso = 0
                    For j = 1 To gnc.CountIVKardex
                        Set ivkOUT = gnc.IVKardex(j)
                        If ivkOUT.cantidad < 0 Then
                            CostoTotalEgreso = CostoTotalEgreso + Abs(ivkOUT.CostoRealTotal)
                        End If
                    Next j
                    
                    If CostoTotalEgreso <> ivk.CostoTotal Then
                        cambiado = True
                        GoTo SiguienteItem
                    End If
                    Set ivkOUT = Nothing
                    GoTo SiguienteItem
                End If
                
                
                'Si el costo es diferente de lo que está grabado
                If ctotal <> ivk.CostoTotal Then
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
    
    If Not booVerificando Then
        
        'Calcula el costo total de los egresos en el item de ingreso
        CostoTotalEgreso = 0
        For i = 1 To gnc.CountIVKardex
            Set ivk = gnc.IVKardex(i)
                If ivk.cantidad < 0 Then
                    CostoTotalEgreso = CostoTotalEgreso + Abs(ivk.CostoRealTotal)
                End If
        Next i
        'busca el item de Ingreso y asigan el costo total de los items de egreso
        For i = 1 To gnc.CountIVKardex
            Set ivk = gnc.IVKardex(i)
                If ivk.cantidad > 0 Then
                        ivk.CostoTotal = CostoTotalEgreso
                        ivk.CostoRealTotal = CostoTotalEgreso
                    i = gnc.CountIVKardex
                End If
        Next i
    End If
SalidaOK:
    RecalculoProduccion = True
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

