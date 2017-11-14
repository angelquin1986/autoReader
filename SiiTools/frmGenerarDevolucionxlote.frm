VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl1.ocx"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "flexcombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmGenerarDevolucionxlote 
   Caption         =   "Generación de devoluciones x lote"
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
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Parametros de Busqueda - F6"
      TabPicture(0)   =   "frmGenerarDevolucionxlote.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "grd"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraFecha"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraCodTrans"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtDescripcion"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdBuscar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.Frame Frame2 
         Caption         =   "Vendedor"
         Height          =   975
         Left            =   8880
         TabIndex        =   22
         Top             =   360
         Width           =   1695
         Begin FlexComboProy.FlexCombo fcbVendedor 
            Height          =   345
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   1455
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
      Begin VB.Frame Frame1 
         Caption         =   "&Fecha Transaccion  "
         Height          =   975
         Left            =   7200
         TabIndex        =   20
         Top             =   360
         Width           =   1815
         Begin MSComCtl2.DTPicker dtpFecha 
            Height          =   360
            Left            =   120
            TabIndex        =   21
            ToolTipText     =   "Fecha de la transacción"
            Top             =   360
            Width           =   1455
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
            Format          =   86704129
            CurrentDate     =   37078
            MaxDate         =   73415
            MinDate         =   29221
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   372
         Left            =   8880
         TabIndex        =   19
         Top             =   1560
         Width           =   1212
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   510
         Left            =   1080
         MaxLength       =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         ToolTipText     =   "Descripción de la transacción"
         Top             =   1440
         Width           =   6300
      End
      Begin VB.Frame fraCodTrans 
         Caption         =   "Cod.&Trans."
         Height          =   975
         Left            =   3525
         TabIndex        =   15
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
         Height          =   975
         Left            =   120
         TabIndex        =   13
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
            Format          =   86704129
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
            Format          =   86704130
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
            Format          =   86704130
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
            Format          =   86704129
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
            TabIndex        =   14
            Top             =   480
            Width           =   315
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "&Trans de Devolucion."
         Height          =   975
         Left            =   5400
         TabIndex        =   12
         Top             =   360
         Width           =   1935
         Begin FlexComboProy.FlexCombo fcbTransDev 
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
         TabIndex        =   7
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "&Descripción  "
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   930
      End
   End
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   852
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   8520
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5565
      Width           =   8520
      Begin VB.CommandButton cmdVerificar 
         Caption         =   "&Verificar"
         Height          =   372
         Left            =   1560
         TabIndex        =   16
         Top             =   360
         Width           =   1212
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Proceder -F3"
         Height          =   372
         Left            =   2880
         TabIndex        =   8
         Top             =   360
         Width           =   1332
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   4320
         TabIndex        =   9
         Top             =   360
         Width           =   1212
      End
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   11
         Top             =   60
         Width           =   8280
         _ExtentX        =   14605
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "frmGenerarDevolucionxlote"
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
Private Const MSG_NG = "No tiene Devolucion."
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
    cmdGrabar.Enabled = False
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub CargaTrans()
    'Carga la lista de transacción
    fcbTrans.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False)
    fcbTransDev.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "IV")
End Sub

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
        .ColHidden(COL_NOMBRE) = False
        .ColHidden(COL_DESC) = False
        .ColHidden(COL_CENTROCOSTO) = True
        .ColHidden(COL_ESTADO) = True
        
        .ColDataType(COL_FECHA) = flexDTDate
        
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
    If GenerarDevolucion(False, False) Then
        cmdVerificar.Enabled = True
        cmdVerificar.SetFocus
        mVerificado = True
    End If
End Sub

Private Sub cmdVerificar_Click()
    'Si no hay transacciones
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna transacción para verificar."
        Exit Sub
    End If
    If dtpFecha1 < gobjMain.EmpresaActual.GNOpcion.FechaLimiteDesde Then
        MsgBox "La Rango de Fecha de regeneración es menor a la Fecha Limite Aceptable  ", vbExclamation
        Exit Sub
    End If

    If GenerarDevolucion(True, False) Then
        cmdGrabar.Enabled = True
        cmdGrabar.SetFocus
        mVerificado = True
    End If
End Sub
Private Function GenerarDevolucion(bandVerificar As Boolean, BandTodo As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, X As Single
    Dim gnc As GNComprobante, cambiado As Boolean
    Dim CodTrans As String, numtrans As Long, Cadena As String
    
    
    On Error GoTo ErrTrap

    'Si no es solo verificacion, confirma
    If Not bandVerificar Then
        s = "Este proceso modificará los asientos de la transacción seleccionada." & vbCr & vbCr
        s = s & "Está seguro que desea proceder?"
        If MsgBox(s, vbYesNo + vbQuestion) <> vbYes Then Exit Function
    End If
    
    If Len(fcbTrans.Text) = 0 Then MsgBox "Ingrese Transaccion de Origen": Exit Function
    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdVerificar.Enabled = False
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
        
        'Si es verificación, procesa todas las filas sino solo las que tengan "Asiento incorrecto."
        If (grd.TextMatrix(i, COL_RESULTADO) = MSG_NG) And Not bandVerificar Then
            If Len(fcbTransDev.Text) = 0 Then MsgBox "Ingrese Transaccion de devolucion": Exit Function
            If Len(fcbVendedor.Text) = 0 Then MsgBox "Ingrese el vendedor": Exit Function
            tid = grd.ValueMatrix(i, COL_TID)
            grd.Refresh
            
            'Recupera la transaccion
            Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
            If Not (gnc Is Nothing) Then
                'Si la transacción no está anulada
                If gnc.Estado <> ESTADO_ANULADO Then
                    gnc.RecuperaDetalleTodo
                    If Not GenDevolucion(gnc, i, Cadena) Then
                        grd.TextMatrix(i, COL_RESULTADO) = "Error.." & Cadena
                        Exit Function
                    End If
                End If
            End If
            grd.TextMatrix(i, COL_RESULTADO) = "Grabo como.." & Cadena
        ElseIf bandVerificar Then
            tid = grd.ValueMatrix(i, COL_TID)
            grd.TextMatrix(i, COL_RESULTADO) = "Verificando..."
            grd.Refresh
            
            'Recupera la transaccion
            Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
            If Not (gnc Is Nothing) Then
                'Si la transacción no está anulada
                If gnc.Estado <> ESTADO_ANULADO Then
                    'Forzar recuperar todos los datos de transacción para que no se pierdan al grabar de nuveo
               '     gnc.RecuperaDetalleTodo
                    'Recalcula costo de los items
                    If NoExisteDevolucion(gnc, CodTrans, numtrans) Then
'                            'Si no es solo verificacion
'                            If (Not bandVerificar) Or bandTodo Then
'                                grd.TextMatrix(i, COL_RESULTADO) = "Grabando..."
'                                grd.Refresh
'
'                                'Graba la transacción
'                                gnc.Grabar False, False
'                                grd.TextMatrix(i, COL_RESULTADO) = "Actualizado."
                                
                            'Si es solo verificacion
'                            Else
                                grd.TextMatrix(i, COL_RESULTADO) = MSG_NG
'                            End If
                     Else
                            'Si no está cambiado no graba
                            grd.TextMatrix(i, COL_RESULTADO) = "OK." & CodTrans & " " & numtrans
                     End If
                Else
                    'Si está anulada
                    grd.TextMatrix(i, COL_RESULTADO) = "Anulado."
                End If
            Else
                grd.TextMatrix(i, COL_RESULTADO) = "No pudo recuperar la transación."
            End If
        End If
    Next i
    Screen.MousePointer = 0
    GenerarDevolucion = Not mCancelado
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    DispErr
salida:
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdVerificar.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    Exit Function
End Function

Private Sub fcbTrans_BeforeSelect(ByVal Row As Long, Cancel As Boolean)
    SacaTransAsientoGnTrans fcbTrans.Text
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
    prg1.Width = Me.ScaleWidth - (prg1.Left * 2)
End Sub

Private Sub grd_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyDelete
        EliminaFila
End Select
End Sub

Private Sub mobjGNComp_EstadoGeneracion1AsientoxLote(ByVal ix As Long, ByVal Estado As String, Cancel As Boolean)
    prg1.value = ix
    grd.TextMatrix(ix, COL_RESULTADO) = Estado
    Cancel = mCancelado
End Sub

Private Sub CargarEncabezado()
    dtpFecha.value = Date
    fcbVendedor.SetData gobjMain.EmpresaActual.ListaFCVendedor(True, False)
    txtDescripcion.Text = "Dev x Lote"
End Sub

Private Sub SacaTransAsientoGnTrans(ByVal CodTrans As String)
    Dim gnt As GNTrans
    Set gnt = gobjMain.EmpresaActual.RecuperaGNTrans(CodTrans)
    If Not gnt Is Nothing Then
        fcbTransDev.KeyText = gnt.TransAsiento
    End If
End Sub
Private Function NoExisteDevolucion(ByVal gnc As GNComprobante, ByRef CodTrans As String, ByRef numtrans As Long _
                                     ) As Boolean
    Dim rs As Recordset
    Dim sql As String
    On Error GoTo ErrTrap
    
    sql = "Select * from gncomprobante where codtrans ='" & fcbTransDev.KeyText & "' AND idtransfuente = " & gnc.TransID
    Set rs = gnc.Empresa.OpenRecordset(sql)
    If rs.RecordCount = 0 Then
        NoExisteDevolucion = True
    Else
        Do While Not rs.EOF
            CodTrans = rs!CodTrans
            numtrans = rs!numtrans
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
    Exit Function
ErrTrap:
    DispErr
    NoExisteDevolucion = False
    Exit Function
End Function

Public Function GenDevolucion(ByVal gc As GNComprobante, ByVal i As Long, ByRef Cadena As String) As Boolean
    Dim gnc As GNComprobante
    On Error GoTo ErrTrap
    MensajeStatus MSG_PREPARA, vbHourglass
    Set gnc = gobjMain.EmpresaActual.CreaGNComprobante(fcbTransDev.KeyText)
    gnc.CodClienteRef = TomarCodigoCli(grd.TextMatrix(i, COL_NOMBRE))  'Asigna el mismo cliente a la
    gnc.CodVendedor = fcbVendedor.KeyText
    gnc.Nombre = grd.TextMatrix(i, COL_NOMBRE)
    gnc.Descripcion = gnc.Descripcion & " " & txtDescripcion.Text & " Trans Fuente:" & gc.CodTrans & ":" & gc.numtrans
    gnc.idTransFuente = gc.TransID
    cargarIVKardex gc, gnc
        If gnc.GNTrans.UtilizarBodegaDestino Then
            If Len(gnc.GNTrans.CodBodegaPre) = 0 Then MsgBox "Debe Asignar una bodega origen " & Chr(13) & "Revise la configuracion de la trans", vbInformation: Exit Function
            If Len(gnc.GNTrans.CodBodegaDesPre) = 0 Then MsgBox "Debe Asignar una bodega Destino " & Chr(13) & "Revise la configuracion de la trans", vbInformation: Exit Function
            EliminaDestino gnc
            CreaDestino gnc
        End If
    gnc.Grabar False, False
    CambiaEstadoHistorialCliente gc.TransID
    CambiaEstadoDeFuente gc.TransID
    Cadena = gnc.CodTrans & " " & gnc.numtrans
    MensajeStatus
    Set gnc = Nothing
    GenDevolucion = True
    Exit Function
ErrTrap:
    MensajeStatus
    DispErr
    Unload Me
    GenDevolucion = False
    Exit Function
End Function
Private Function TomarCodigoCli(ByVal Nombre As String) As String
Dim sql As String
Dim rs As Recordset
Dim codigo As String
On Error GoTo CapturaError
    sql = "select codprovcli from pcprovcli where nombre = '" & Nombre & "'"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    Do While Not rs.EOF
        codigo = rs!CodProvCli
        rs.MoveNext
    Loop
    TomarCodigoCli = codigo
    Set rs = Nothing
Exit Function
CapturaError:
    MsgBox Err.Description
    Set rs = Nothing
    Exit Function
End Function

Public Sub cargarIVKardex(ByVal GnFuente As GNComprobante, ByVal gnDestino As GNComprobante)
Dim i As Long
Dim ix As Long
Dim item As IVinventario
    For i = 1 To GnFuente.CountIVKardex
        Set item = GnFuente.Empresa.RecuperaIVInventario(GnFuente.IVKardex(i).CodInventario)
            If item.Tipo <> Preparacion Then
                ix = gnDestino.AddIVKardex
                gnDestino.IVKardex(ix).CodInventario = GnFuente.IVKardex(i).CodInventario
                gnDestino.IVKardex(ix).cantidad = Abs(GnFuente.IVKardex(i).cantidad)
                gnDestino.IVKardex(ix).Orden = i 'Guarda el orden de los items
                gnDestino.IVKardex(ix).bandImprimir = True 'Si es subitem para que no imprima
            End If
        Next
        Set item = Nothing
End Sub
Private Sub CambiaEstadoHistorialCliente(ByVal tid As Long)
Dim sql As String
On Error GoTo CapturaError
        sql = "UPDATE PCHistorial set estado = " & ESTADO_DESPACHADO & _
              ",fechagrabado = '" & Now & "'" & _
              " WHERE (TransID =" & tid & " AND Estado = " & ESTADO_APROBADO & " ) "
        gobjMain.EmpresaActual.Execute sql, True
        Exit Sub
CapturaError:
    Exit Sub
    MsgBox Err.Description
End Sub

Private Sub CambiaEstadoDeFuente(ByVal tid As Long)
Dim sql As String
        sql = "UPDATE GNComprobante SET Estado=" & ESTADO_DESPACHADO & _
              "WHERE (TransID=" & tid & ") AND ((Estado=" & ESTADO_APROBADO & ") or (EStado = " & ESTADO_SEMDESPACHADO & "))"
        gobjMain.EmpresaActual.Execute sql, True
End Sub
Private Sub EliminaFila()
Dim i As Long
Dim Pregunta
If grd.Rows > 1 Then
    Pregunta = MsgBox("Desea Eliminar....", vbYesNo)
    If Pregunta = vbYes Then
            grd.RemoveItem grd.Row
            grd.SetFocus
    End If
End If
End Sub

Private Sub EliminaDestino(ByVal gnc As GNComprobante)
    Dim i As Long, ivk As IVKardex, ivk2 As IVKardex, j As Long
    'Elimina items con cantidad positiva
    '  para que se queden solo los items de origen
    For i = gnc.CountIVKardex To 1 Step -1
        If gnc.IVKardex(i).cantidad < 0 Then             'Si es de destino, elimina
            gnc.RemoveIVKardexPreparacion i
        End If
    Next i
End Sub

Private Sub CreaDestino(ByVal gnc As GNComprobante)
    Dim i As Long, ivk As IVKardex, ivk2 As IVKardex, j As Long
    'Asegura que todo sea de la bodega de origen
    For i = 1 To gnc.CountIVKardex
        Set ivk = gnc.IVKardex(i)
        ivk.CodBodega = gnc.GNTrans.CodBodegaPre
        ivk.cantidad = Abs(ivk.cantidad) * -1   'Origen es siempre negativa
    Next i
    Set ivk = Nothing
    'Duplica IVKardex para la bodega de destino
    ' multiplicando -1 a la cantidad
    For i = 1 To gnc.CountIVKardex
        Set ivk = gnc.IVKardex(i)
        j = gnc.AddIVKardex
        Set ivk2 = gnc.IVKardex(j)
        With ivk2
            .CodBodega = gnc.GNTrans.CodBodegaDesPre
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

