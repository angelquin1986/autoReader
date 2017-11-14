VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReprocDesperdicio 
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
         Format          =   50724865
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
         Format          =   50724865
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
Attribute VB_Name = "frmReprocDesperdicio"
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


Private Const MSG_NG = "Desperdicio incorrecto."
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
    
    'Carga la lista de transacción
'    fcbTrans.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "IV")

    lstTrans.Clear
    v = gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "IV")
    For i = LBound(v, 2) To UBound(v, 2)
        lstTrans.AddItem v(0, i)        '& " " & v(1, i)
    Next i
End Sub



Private Sub cmdAceptar_Click()
    'Si no hay transacciones
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna transacción para procesar.", vbExclamation
        Exit Sub
    End If
    
    If ReprocDesp(False) Then
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

Private Function VerificaEgreso() As String
    Dim i As Long, cod As String, gnt As GNTrans
    Dim s As String
    
    For i = 0 To lstTrans.ListCount - 1
        'Si está seleccionado
        If lstTrans.Selected(i) Then
            'Recupera el objeto GNTrans
            cod = lstTrans.List(i)
            Set gnt = gobjMain.EmpresaActual.RecuperaGNTrans(cod)
            'Si la transaccion es de ingreso, devuelve el codigo
            If gnt.IVTipoTrans = "E" Then s = s & cod & ", "
        End If
    Next i
    Set gnt = Nothing
    If Len(s) > 2 Then s = Left$(s, Len(s) - 2)     'Quita la ultima ", "
    VerificaEgreso = s
End Function



Private Function ReprocCosto(ByVal bandVerificar As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, x As Single
    Dim gnc As GNComprobante, cambiado As Boolean
    
    On Error GoTo ErrTrap
    
    'Si no es solo verificacion, confirma
    If Not bandVerificar Then
        'Confirma la actualización
        s = "Este proceso modificará los Desperdicios de la transacción seleccionada." & vbCr & vbCr
        s = s & "Está seguro que desea proceder?"
        If MsgBox(s, vbYesNo + vbQuestion) <> vbYes Then Exit Function
    End If
    
    'Verifica si está seleccionado una trans. de ingreso
    s = VerificaEgreso
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
    prg1.Min = 0
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
                
                    'Recalcula costo de los items
                    If Recalculo(gnc, cambiado, bandVerificar) Then
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
    prg1.value = prg1.Min
salida:
    Set mColItems = Nothing         'Libera el objeto de coleccion
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdVerificar.Enabled = True
    cmdCorregirIVA.Enabled = True
    cmdBuscar.Enabled = True
    cmdAceptar.Enabled = True
    prg1.value = prg1.Min
    Exit Function
End Function


Private Function Recalculo(ByVal gnc As GNComprobante, _
                           ByRef cambiado As Boolean, _
                           ByVal booVerificando As Boolean) As Boolean
    Dim item As IVinventario, ivk As IVKardex, i As Long
    Dim ct As Currency, ctotal As Currency, s As String
    Dim CostoTotalEgreso As Currency
    Dim ivkOUT As IVKardex, itemOUT As IVinventario
    Dim CostoTotalPadre As Currency
    On Error GoTo ErrTrap
    
    cambiado = False
    For i = 1 To gnc.CountIVKardex
        Set ivk = gnc.IVKardex(i)
        If gnc.GNTrans.CodPantalla = "IVCAMIE" Then
             If Not i > (gnc.CountIVKardex / 2) Then
                GoTo SiguienteItem
            End If
        End If
'        'Solo de salida
            'Recupera el item
'            If ivk.CodInventario <> "7861074600618" Then GoTo SiguienteItem
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
                                       Abs(ivk.Cantidad), _
                                       gnc.TransID, _
                                       gnc.HoraTrans)
                
                'Convierte en moneda de la transaccion
                If item.CodMoneda <> gnc.CodMoneda Then
                    ct = ct * gnc.Cotizacion(item.CodMoneda) / gnc.Cotizacion("")
                End If
                ctotal = ct * ivk.Cantidad
                CostoTotalPadre = ctotal
                'Si el costo es diferente de lo que está grabado
                If ctotal <> ivk.CostoTotal Then
                    If gnc.GNTrans.IVTipoTrans = "C" And i > 1 And gnc.GNTrans.CodPantalla <> "IVCAMIE" Then
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
                                                   Abs(ivkOUT.Cantidad), _
                                                   gnc.TransID, _
                                                   gnc.HoraTrans)
                            
                            'Convierte en moneda de la transaccion
                            If itemOUT.CodMoneda <> gnc.CodMoneda Then
                                ct = ct * gnc.Cotizacion(itemOUT.CodMoneda) / gnc.Cotizacion("")
                            End If
                            ctotal = ct * ivkOUT.Cantidad * -1
                            ivk.CostoTotal = ctotal
                            ivk.CostoRealTotal = ctotal
                            Set ivkOUT = Nothing
                            Set itemOUT = Nothing
                        End If
                    ElseIf gnc.GNTrans.IVTipoTrans = "C" And i > 1 And gnc.GNTrans.CodPantalla = "IVCAMIE" Then
                        Set ivkOUT = gnc.IVKardex(i - (gnc.CountIVKardex / 2))
                        Set itemOUT = gnc.Empresa.RecuperaIVInventario(ivkOUT.CodInventario)
                        If Not (itemOUT Is Nothing) Then
                            '*** MAKOTO 31/ago/00
                            If booVerificando Then
                                mColItems.Add item:=item.CodInventario, Key:=item.CodInventario
                                Debug.Print "Incorrecto 1 . cod='" & item.CodInventario & "' Trans=" & gnc.CodTrans & gnc.numtrans
                                Debug.Print "    dif.:" & ctotal & "," & ivk.CostoTotal
                                ivk.CostoTotal = ctotal
                                ivk.CostoRealTotal = ctotal

                                
                                If ItemIncorrecto(itemOUT.CodInventario) Then      'Este item ya está marcado como incorrecto.
                                    cambiado = True
                                    GoTo SiguienteItem
                                End If
                            End If
                            ct = itemOUT.CostoDouble2(gnc.FechaTrans, _
                                                   Abs(ivkOUT.Cantidad), _
                                                   gnc.TransID, _
                                                   gnc.HoraTrans)
                            
                            'Convierte en moneda de la transaccion
                            If itemOUT.CodMoneda <> gnc.CodMoneda Then
                                ct = ct * gnc.Cotizacion(itemOUT.CodMoneda) / gnc.Cotizacion("")
                            End If
                            ctotal = ct * ivkOUT.Cantidad * -1
                            If ctotal <> CostoTotalPadre Then
                                gnc.IVKardex(i - (gnc.CountIVKardex / 2)).CostoTotal = CostoTotalPadre * -1
                                gnc.IVKardex(i - (gnc.CountIVKardex / 2)).CostoRealTotal = CostoTotalPadre * -1
                                gnc.IVKardex(i).CostoTotal = CostoTotalPadre
                                gnc.IVKardex(i).CostoRealTotal = CostoTotalPadre
                                
                                cambiado = True
'                                ivkOUT.CostoTotal = CostoTotalPadre
'                                ivkOUT.CostoRealTotal = ctotal
                            End If
                            
'                            ivk.CostoTotal = ctotal
'                            ivk.CostoRealTotal = ctotal
                            Set ivkOUT = Nothing
                            Set itemOUT = Nothing
                        End If
                        
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

Private Sub cmdBuscar_Click()
    Dim v As Variant, obj As Object
    On Error GoTo ErrTrap
    
    '*** MAKOTO 06/sep/00 Agregado
    If lstTrans.SelCount = 0 Then
        MsgBox "Seleccione una transacción, por favor.", vbInformation
        Exit Sub
    End If
    
    With gobjMain.objCondicion
        .Fecha1 = dtpFecha1.value
        .Fecha2 = dtpFecha2.value
'        .CodTrans = fcbTrans.Text              '*** MAKOTO 31/ago/00 Modificado
        .CodTrans = PreparaCodTrans             '***
        .NumTrans1 = Val(txtNumTrans1.Text)
        .NumTrans2 = Val(txtNumTrans2.Text)
        
        'Estados no incluye anulados
        .EstadoBool(ESTADO_NOAPROBADO) = True
        .EstadoBool(ESTADO_APROBADO) = True
        .EstadoBool(ESTADO_DESPACHADO) = True
        .EstadoBool(ESTADO_ANULADO) = False
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
    
    If ReprocDesp(True) Then
        cmdAceptar.Enabled = True
        cmdAceptar.SetFocus
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

Private Function ReprocDesp(ByVal bandVerificar As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, x As Single
    Dim gnc As GNComprobante, cambiado As Boolean
    
    On Error GoTo ErrTrap
    
    'Si no es solo verificacion, confirma
    If Not bandVerificar Then
        'Confirma la actualización
        s = "Este proceso modificará los Desperdicios de la transacción seleccionada." & vbCr & vbCr
        s = s & "Está seguro que desea proceder?"
        If MsgBox(s, vbYesNo + vbQuestion) <> vbYes Then Exit Function
    End If
    
    'Verifica si está seleccionado una trans. de Egreso
    s = VerificaEgreso
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
    prg1.Min = 0
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
                
                    'Recalcula costo de los items
                    If RecalculoDesperdicio(gnc, cambiado, bandVerificar) Then
                        'Si está cambiado algo
                        If cambiado Then
                            'Si no es solo verificacion
                            If Not bandVerificar Then
                                grd.TextMatrix(i, COL_RESULTADO) = "Grabando..."
                                grd.Refresh
                                
                                'Prorratea los recargos/descuentos si los calcula en base a costo
                                gnc.ProrratearIVKardexRecargo
                                'gnc.GeneraAsiento       'Diego 27 Abril 2001  corregido
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
    ReprocDesp = Not mCancelado
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    If i < grd.Rows And i >= grd.FixedRows Then
        grd.TextMatrix(i, COL_RESULTADO) = Err.Description
    End If
    DispErr
    prg1.value = prg1.Min
salida:
    Set mColItems = Nothing         'Libera el objeto de coleccion
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdVerificar.Enabled = True
    cmdCorregirIVA.Enabled = True
    cmdBuscar.Enabled = True
    cmdAceptar.Enabled = True
    prg1.value = prg1.Min
    Exit Function
End Function


Private Function RecalculoDesperdicio(ByVal gnc As GNComprobante, _
                           ByRef cambiado As Boolean, _
                           ByVal booVerificando As Boolean) As Boolean
    Dim item As IVinventario, ivk As IVKardex, i As Long
    Dim ct As Currency, ctotal As Currency, s As String
    Dim CostoTotalEgreso As Currency
    Dim ivkDesp As IVKardex, itemOUT As IVinventario
    Dim CostoTotalPadre As Currency
    Dim j As Long, cont As Integer
    On Error GoTo ErrTrap
    
    For i = 1 To gnc.CountIVKardex
        Set ivk = gnc.IVKardex(i)
        'pregunta si es % despedicio si es negativo
        If ivk.Cantidad < 0 Then
            'borra fila para luego generar nuevamente desperdicio
            
        End If
    Next i
    
    
    'Set GncAux = gnc
    cont = 1
    cambiado = False
    For i = 1 To gnc.CountIVKardex
        Set ivk = gnc.IVKardex(i)
        Set ivkDesp = gnc.IVKardex(i)
        Set item = gnc.Empresa.RecuperaIVInventario(ivk.CodInventario)
            If Not (item Is Nothing) Then
                '*** MAKOTO 31/ago/00
                'pregunta si tiene % desperdicio
                If item.PorDesperdicio > 0 Then
                    'agrga fila con desperdicio
'                    gnc.AddIVKardex
                    'Agrega detalle
                    j = gnc.AddIVKardex
                    Set ivkDesp = gnc.IVKardex(j)
                    ivkDesp.Cantidad = item.PorDesperdicio * ivk.Cantidad * -1
                    ivkDesp.CodBodega = ivk.CodBodega
                    ivkDesp.CodInventario = ivk.CodInventario
                    ivkDesp.CostoRealTotal = 0
                    ivkDesp.CostoTotal = 0
                    ivkDesp.PrecioRealTotal = ivk.PrecioRealTotal
                    ivkDesp.PrecioTotal = ivk.PrecioTotal
                    ivkDesp.Orden = gnc.CountIVKardex
                    cont = cont + 1
                End If
                cambiado = True
                'Convierte en moneda de la transaccion
                If item.CodMoneda <> gnc.CodMoneda Then
                    ct = ct * gnc.Cotizacion(item.CodMoneda) / gnc.Cotizacion("")
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
    RecalculoDesperdicio = True
    GoTo salida
    Exit Function
ErrTrap:
    DispErr
salida:
    Set ivk = Nothing
    Set ivkDesp = Nothing
    Set item = Nothing
    Set gnc = Nothing
    Exit Function
End Function


