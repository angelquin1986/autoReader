VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCierre 
   Caption         =   "Cierre de ejercicio"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7605
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5970
   ScaleWidth      =   7605
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPasos 
      Caption         =   "GO"
      Height          =   330
      Index           =   19
      Left            =   6840
      TabIndex        =   35
      Top             =   3000
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.CommandButton cmdPasos 
      Caption         =   "GO"
      Height          =   330
      Index           =   8
      Left            =   3240
      TabIndex        =   33
      Top             =   3240
      Width           =   612
   End
   Begin VB.CommandButton cmdPasos 
      Caption         =   "GO"
      Height          =   330
      Index           =   0
      Left            =   3240
      TabIndex        =   13
      Top             =   120
      Width           =   612
   End
   Begin VB.Frame Frame2 
      Caption         =   "Empresa destino"
      Height          =   1000
      Left            =   4080
      TabIndex        =   5
      Top             =   1200
      Width           =   2772
      Begin VB.TextBox txtDestino 
         Height          =   300
         Left            =   720
         TabIndex        =   7
         Top             =   240
         Width           =   1812
      End
      Begin VB.TextBox txtDestinoBD 
         Height          =   300
         Left            =   720
         TabIndex        =   9
         Top             =   600
         Width           =   1812
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Código  "
         Height          =   192
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "B.D."
         Height          =   192
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Empresa orígen"
      Height          =   1000
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   2772
      Begin VB.Label lblOrigenBD 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   720
         TabIndex        =   4
         Top             =   600
         Width           =   1812
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "B.D."
         Height          =   192
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Código  "
         Height          =   192
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblOrigen 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   1812
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   2208
      Left            =   0
      ScaleHeight     =   2205
      ScaleWidth      =   7605
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3765
      Width           =   7605
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   288
         Left            =   5760
         TabIndex        =   30
         Top             =   1800
         Width           =   1212
      End
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   252
         Left            =   120
         TabIndex        =   31
         Top             =   1800
         Width           =   5532
         _ExtentX        =   9763
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VSFlex7LCtl.VSFlexGrid grd 
         Height          =   1692
         Left            =   120
         TabIndex        =   32
         Top             =   0
         Width           =   7332
         _cx             =   12933
         _cy             =   2984
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
         FormatString    =   $"frmCierre.frx":0000
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
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "&Opciones..."
      Enabled         =   0   'False
      Height          =   372
      Left            =   4800
      TabIndex        =   28
      Top             =   3000
      Width           =   1812
   End
   Begin VB.CommandButton cmdPasos 
      Caption         =   "GO"
      Height          =   330
      Index           =   1
      Left            =   3240
      TabIndex        =   15
      Top             =   480
      Width           =   612
   End
   Begin VB.CommandButton cmdPasos 
      Caption         =   "GO"
      Height          =   330
      Index           =   2
      Left            =   3240
      TabIndex        =   17
      Top             =   840
      Width           =   612
   End
   Begin VB.CommandButton cmdPasos 
      Caption         =   "GO"
      Height          =   330
      Index           =   3
      Left            =   3240
      TabIndex        =   19
      Top             =   1200
      Width           =   612
   End
   Begin VB.CommandButton cmdPasos 
      Caption         =   "GO"
      Height          =   456
      Index           =   4
      Left            =   3240
      TabIndex        =   21
      Top             =   1560
      Width           =   612
   End
   Begin VB.CommandButton cmdPasos 
      Caption         =   "GO"
      Height          =   330
      Index           =   5
      Left            =   3240
      TabIndex        =   23
      Top             =   2040
      Width           =   612
   End
   Begin VB.CommandButton cmdPasos 
      Caption         =   "GO"
      Height          =   330
      Index           =   6
      Left            =   3240
      TabIndex        =   25
      Top             =   2400
      Width           =   612
   End
   Begin VB.CommandButton cmdPasos 
      Caption         =   "GO"
      Height          =   456
      Index           =   7
      Left            =   3240
      TabIndex        =   27
      Top             =   2745
      Width           =   612
   End
   Begin MSComCtl2.DTPicker dtpFechaCorte 
      Height          =   300
      Left            =   4800
      TabIndex        =   11
      Top             =   2520
      Width           =   1812
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      Format          =   107020289
      CurrentDate     =   36781
   End
   Begin VB.Label lblPasos 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9. Desactivar la base de datos original"
      Height          =   336
      Index           =   8
      Left            =   120
      TabIndex        =   34
      Top             =   3240
      Width           =   3060
   End
   Begin VB.Label lblPasos 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1. Generar asiento de cierre en orígen"
      Height          =   336
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   3060
   End
   Begin VB.Label lblPasos 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8. Pasar trans. existentes con fecha posterior a la fecha de corte "
      Height          =   456
      Index           =   7
      Left            =   120
      TabIndex        =   26
      Top             =   2760
      Width           =   3060
   End
   Begin VB.Label lblPasos 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7. Pasar saldo inicial de cuenta contable"
      Height          =   336
      Index           =   6
      Left            =   120
      TabIndex        =   24
      Top             =   2400
      Width           =   3060
   End
   Begin VB.Label lblPasos 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6. Pasar saldo inicial de bancos"
      Height          =   336
      Index           =   5
      Left            =   120
      TabIndex        =   22
      Top             =   2040
      Width           =   3060
   End
   Begin VB.Label lblPasos 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5. Pasar saldo inicial de proveedores y clientes"
      Height          =   456
      Index           =   4
      Left            =   120
      TabIndex        =   20
      Top             =   1560
      Width           =   3060
   End
   Begin VB.Label lblPasos 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4. Pasar saldo inicial de inventario"
      Height          =   336
      Index           =   3
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   3060
   End
   Begin VB.Label lblPasos 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3. Resetear # de trans. en el destino"
      Height          =   336
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   3060
   End
   Begin VB.Label lblPasos 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2. Copiar datos de catálogos a destino"
      Height          =   336
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   480
      Width           =   3060
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de corte  "
      Height          =   192
      Left            =   4200
      TabIndex        =   10
      Top             =   2280
      Width           =   1152
   End
End
Attribute VB_Name = "frmCierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mbooProcesando As Boolean
Private mbooCancelado As Boolean
Private mEmpOrigen As Empresa
Private Const MSG_OK As String = "OK"

Private WithEvents mGrupo As grupo
Attribute mGrupo.VB_VarHelpID = -1

Public Sub Inicio()
    On Error GoTo ErrTrap
    
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

Private Function VerificarOpcion() As Boolean
    Dim code As String, Fcorte As Date, i As Long, TienePermiso As Boolean
    
    'Código de emrpesa destino
    code = Trim$(txtDestino.Text)
    If Len(code) = 0 Then
        MsgBox "Ingrese el código de la empresa destino.", vbExclamation
        Exit Function
    End If
    
    'Destino no puede ser la misma que origen
    If UCase$(code) = UCase$(mEmpOrigen.CodEmpresa) Then
        MsgBox "La empresa destino no puede ser la misma que la orígen.", vbExclamation
        Exit Function
    End If
    
    'Fecha de corte
    Fcorte = dtpFechaCorte.value
    If Fcorte < mEmpOrigen.GNOpcion.FechaInicio Then
        MsgBox "La fecha de corte no puede ser antes de la fecha de inicio del período.", vbExclamation
        Exit Function
    End If
    
    'Prueba si tiene acceso a la empresa destino
    TienePermiso = False
    For i = 1 To gobjMain.GrupoActual.CountPermiso
        If UCase(gobjMain.GrupoActual.Permisos(i).CodEmpresa) = UCase(code) Then
            TienePermiso = True
            Exit For
        End If
    Next i
    If Not TienePermiso Then
        MsgBox "El usuario actual '" & gobjMain.UsuarioActual.codUsuario & "' " & _
               "no tiene permiso para acceder a la empresa destino '" & code & "'. " & vbCr & vbCr & _
               "Primero deberá dar permiso necesario en el programa 'SiiConfig'.", vbInformation
        Exit Function
    End If
    
    VerificarOpcion = True
End Function

Private Sub cmdPasos_Click(Index As Integer)
    Dim r As Boolean, res As Integer
    
    Select Case Index + 1
    Case 1      '1. Generar asiento de cierre en orígen
        r = GenerarAsientoCierre
    Case 2      '2. Copiar datos a destino
        r = CopiarDatos
    Case 3      '3. Resetear # de transacciones
        r = ResetearNumTrans
    Case 4      '4. Pasar saldo inicial de inventario
        r = SaldoIV
'''        r = SaldoAF
'''        res = MsgBox("Saldo inicial Inventarios", vbYesNo)
'''        If res = vbYes Then
'''           r = SaldoIV
'''        End If
'''        res = MsgBox("Saldo inicial Activos Fijos", vbYesNo)
'''        If res = vbYes Then
'''            r = SaldoInicialAF
'''        End If
'''        res = MsgBox("Depreciaciones Acumuladas", vbYesNo)
'''        If res = vbYes Then
'''            r = SaldoAF
'''        End If


    Case 5      '5. Pasar saldo inicial de proveedores
        r = SaldoPC
    Case 6      '6. Pasar saldo inicial de bancos
        r = SaldoTS
    Case 7      '7. Pasar saldo inicial de cuenta contable
        r = SaldoCT
    Case 8      '8. Pasar trans. existentes con la fecha posterior a la fecha de corte
        r = CopiaTrans
    Case 9      '9. Desactivar la base de datos original
        r = DesactivarOrigen
    Case 10
        r = CopiarHistorial
        
    Case 20      '2. Copiar datos Fichas
        r = CopiarDatosFichas
        
    End Select
    
    If r Then
        If Index < cmdPasos.count - 1 Then cmdPasos(Index + 1).SetFocus
        lblPasos(Index).BackColor = vbBlue
        lblPasos(Index).ForeColor = vbYellow
    End If
End Sub

'1. Generar asiento de cierre en orígen
Private Function GenerarAsientoCierre() As Boolean
    Dim sql As String, rs As Recordset, i As Long, rpos As Long
    Dim gc As GNComprobante, ctd As CTLibroDetalle
    Dim Fcorte As Date
    On Error GoTo ErrTrap
        
    'Verifica las opciones
    If Not VerificarOpcion Then Exit Function
    
    Fcorte = dtpFechaCorte.value    'Fecha de corte

    'Cambia figura de cursor de mouse
    MensajeStatus "Está preparando saldos a la fecha de corte...", vbHourglass
    mensaje True, "Generando asiento de cierre..."
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True

        If mEmpOrigen.GNOpcion.ObtenerValor("PermitirDistribucionGastos") = "1" Then
            'Obtiene Saldos de cuentas contables de categoría 4 y 5 (Ingreso y Egreso)
            sql = "SELECT ct.CodCuenta, " & _
                         "Sum((ctd.Debe-ctd.Haber)/gc.Cotizacion2) AS Saldo , isnull(codgasto,'0') as codgasto " & _
                  "FROM (GNComprobante gc INNER JOIN " & _
                            "(CTLibroDetalle ctd left join gngasto gng on ctd.idgasto=gng.idgasto INNER JOIN CTCuenta ct " & _
                            "ON ctd.IdCuenta=ct.IdCuenta) " & _
                        "ON ctd.CodAsiento = gc.CodAsiento) " & _
                  "WHERE (gc.Estado IN (" & ESTADO_APROBADO & ", " & ESTADO_DESPACHADO & ", " & ESTADO_SEMDESPACHADO & ")) AND " & _
                        "(ct.TipoCuenta IN (4,5)) AND BANDNIIF=0 AND " & _
                        "(gc.FechaTrans <" & FechaYMD(Fcorte + 1, mEmpOrigen.TipoDB) & ") " & _
                  "GROUP BY ct.CodCuenta ,codgasto " & _
                  "HAVING (Sum((ctd.Debe-ctd.Haber)/gc.Cotizacion2) <> 0) " & _
                  "ORDER BY ct.CodCuenta"
        Else
            'Obtiene Saldos de cuentas contables de categoría 4 y 5 (Ingreso y Egreso)
            sql = "SELECT ct.CodCuenta, " & _
                         "Sum((ctd.Debe-ctd.Haber)/gc.Cotizacion2) AS Saldo " & _
                  "FROM (GNComprobante gc INNER JOIN " & _
                            "(CTLibroDetalle ctd left join gngasto gng on ctd.idgasto=gng.idgasto INNER JOIN CTCuenta ct " & _
                            "ON ctd.IdCuenta=ct.IdCuenta) " & _
                        "ON ctd.CodAsiento = gc.CodAsiento) " & _
                  "WHERE (gc.Estado IN (" & ESTADO_APROBADO & ", " & ESTADO_DESPACHADO & ", " & ESTADO_SEMDESPACHADO & ")) AND " & _
                        "(ct.TipoCuenta IN (4,5)) AND " & _
                        "(gc.FechaTrans <" & FechaYMD(Fcorte + 1, mEmpOrigen.TipoDB) & ") " & _
                  "GROUP BY ct.CodCuenta " & _
                  "HAVING (Sum((ctd.Debe-ctd.Haber)/gc.Cotizacion2) <> 0) " & _
                  "ORDER BY ct.CodCuenta"
        End If
          
    Set rs = mEmpOrigen.OpenRecordset(sql)
    With rs
        If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
        i = 0
        Do Until .EOF
            prg1.value = rs.AbsolutePosition
            prg1.Refresh
            DoEvents
            rpos = 0
            MensajeStatus "Agregando detalle: #" & i & " de " & rs.RecordCount, vbHourglass
            
            'Si aplastó 'Cancelar'
            If mbooCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                GoTo cancelado
            End If
            
            Set ctd = PrepararTransCT(mEmpOrigen, "CTD", _
                        "Cierre de ejercicio", _
                        Fcorte - 1, gc, True)
            ctd.codcuenta = .Fields("CodCuenta")
            ctd.Haber = .Fields("Saldo")
            ctd.Descripcion = gc.Descripcion
            ctd.Orden = i
            If Len(mEmpOrigen.GNOpcion.ObtenerValor("PermitirDistribucionGastos")) > 0 Then
                If mEmpOrigen.GNOpcion.ObtenerValor("PermitirDistribucionGastos") = "1" Then
                    If .Fields("CodGasto") <> "0" Then
                        ctd.CodGasto = .Fields("CodGasto")
                    End If
                End If
            End If
            
            i = i + 1
            .MoveNext
        Loop
        .Close
    End With
    
    'Graba la transacción si no están grabadas
    If Not (gc Is Nothing) Then GrabarTransCT gc, True
    
    mensaje False, "", "OK"
    GenerarAsientoCierre = True
    
cancelado:
    MensajeStatus
    Set ctd = Nothing
    Set gc = Nothing
    
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
    Exit Function
ErrTrap:
    mensaje False, "", Err.Description
    MensajeStatus
    DispErr
    GoTo cancelado
End Function

'Agrega un detalle de TSKardex a GNComprobante
'Si comprobante llega a tener 100 detalles,
'Graba lo anterior y crea otra instancia
Private Function PrepararTransCT(ByVal e As Empresa, _
                            ByVal codt As String, _
                            ByVal Desc As String, _
                            ByVal Fcorte As Date, _
                            ByRef gc As GNComprobante, _
                            ByVal BandCierre As Boolean) As CTLibroDetalle
    Dim j As Long, ctd As CTLibroDetalle
                            
    'Crea transaccion si no existe todavía
    If gc Is Nothing Then
        Set gc = CrearTrans(e, codt, Desc, Fcorte, "")
    End If
    
    'Si llega a tener 100 detalles
    If gc.CountTSKardex >= 100 Then
        GrabarTransCT gc, BandCierre
        
        'Crea nueva instancia de GNComprobante
        Set gc = CrearTrans(e, codt, Desc, Fcorte, "")
    End If

    'Agrega detalle
    j = gc.AddCTLibroDetalle
    Set PrepararTransCT = gc.CTLibroDetalle(j)
End Function

Private Sub GrabarTransCT( _
                ByVal gc As GNComprobante, _
                ByVal BandCierre As Boolean)
    Dim j As Long, ctd As CTLibroDetalle

    'Graba la transacción
    MensajeStatus "Grabándo la transacción...", vbHourglass
    
    'Si es asiento de cierre
    If BandCierre Then
        'Antes de grabar, cuadra el asiento con la cuenta de resultado
        If (gc.DebeTotal - gc.HaberTotal) <> 0 Then
            j = gc.AddCTLibroDetalle
            Set ctd = gc.CTLibroDetalle(j)
            ctd.codcuenta = gc.Empresa.GNOpcion.CodCuentaResultado
            ctd.Haber = gc.DebeTotal - gc.HaberTotal
            ctd.Descripcion = "Resultado del ejercicio"
            ctd.Orden = j
        End If
    End If
            
    gc.Grabar False, False
End Sub

'Crear base de datos destino
'(NO ESTA USADO)
Private Function CrearDestino() As Boolean
'    Dim s As String
'    On Error GoTo errtrap
'
'    'Verifica las opciones
'    If Not VerificarOpcion Then Exit Function
'
'    s = "Para crear la base de datos de destino, váyase al programa 'SiiConfig' " & _
'           "En el menú 'Configuración' - 'Empresas', cree una nueva empresa." & _
'           "Al grabar la nueva empresa activándo la casilla que dice " & _
'           "'Crear B.D. físicamente' se creará la base de datos de nueva empresa."
'    MsgBox s, vbInformation, "Para crear destino"
'
'    CrearDestino = True
'    Exit Function
'errtrap:
'    DispErr
'    Exit Function
End Function

'2. Copiar datos a destino
Private Function CopiarDatos() As Boolean
    Dim sql As String, n As Long, e As Empresa, rpos As Long
    On Error GoTo ErrTrap
    
    'Verifica las opciones
    If Not VerificarOpcion Then Exit Function
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    
    n = CopiarTabla("GNOpcion", "Opciones de empresa")
    n = CopiarTabla("GNOpcion2", "Opciones de avanzadas")
    n = CopiarTabla("CTCuenta", "Plan de cuenta")
    n = CopiarTabla("TSBanco", "Catálogo de bancos")
    n = CopiarTabla("TSTipoDocBanco", "Catálogo de documentos bancarios")
    n = CopiarTabla("TSFormaCobroPago", "Catálogo de forma de pagos/cobros")
    n = CopiarTabla("IVRecargo", "Catálogo de Recargos/Descuentos")
    n = CopiarTabla("CTPresupuesto", "Presupuestos")
    n = CopiarTabla("GNTrans", "Catálogo de transacciones")
    n = CopiarTabla("GNTransAsiento", "Definición de asientos por transacción")
    n = CopiarTabla("GNTransRecargo", "Definición de recargos/descuentos por transacción")
    n = CopiarTabla("GNCentroCosto", "Catálogo de centro de costo")
    n = CopiarTabla("GNResponsable", "Catálogo de responsable")
    n = CopiarTabla("FCVendedor", "Catálogo de vendedor")
    n = CopiarTabla("IVBodega", "Catálogo de bodega")
    n = CopiarTabla("IVGrupo1", "Catálogo de Grupo1 de inventario")
    n = CopiarTabla("IVGrupo2", "Catálogo de Grupo2 de inventario")
    n = CopiarTabla("IVGrupo3", "Catálogo de Grupo3 de inventario")
    n = CopiarTabla("IVGrupo4", "Catálogo de Grupo4 de inventario")
    n = CopiarTabla("IVGrupo5", "Catálogo de Grupo5 de inventario")
    n = CopiarTabla("PCProvCli", "Catálogo de proveedores/clientes")
    n = CopiarTabla("PCContacto", "Catálogo de contactos")
    n = CopiarTabla("PCGrupo1", "Catálogo de Grupo1 de proveedores/clientes")
    n = CopiarTabla("PCGrupo2", "Catálogo de Grupo2 de proveedores/clientes")
    n = CopiarTabla("PCGrupo3", "Catálogo de Grupo3 de proveedores/clientes")
    n = CopiarTabla("PCGrupo4", "Catálogo de Grupo4 de proveedores/clientes")
    n = CopiarTabla("PCprovincia", "PCprovincia")
    n = CopiarTabla("PCCanton", "PCCanton")
    n = CopiarTabla("PCParroquia", "PCParroquia")
    
    n = CopiarTabla("IVInventario", "Catálogo de inventarios")
    '***Agregado. 05/mar/02. Angel
    n = CopiarTabla("TSRetencion", "Catálogo de Retenciones")
    '***Agregado. 18/jun/03. Angel. Tabla para producción
    n = CopiarTabla("IVMateria", "Catálogo de Familias/Recetas")
    '***Agregado. 26/05/05. jeaa. Tabla descto pot item x cliente
    n = CopiarTabla("DescIVGPCG", "Catálogo de Desc Item-Clinte")
    n = CopiarTabla("InventarioProveedor", "Historial de Compras x Proveedor")
    '***Agregado. 17/06/05. jeaa. Tabla motivo de devoluciones
    n = CopiarTabla("Motivo", "Motivos de Devoluciones")
    n = CopiarTabla("IVReservacion", "Reservaciones")
    n = CopiarTabla("IVProveedorDetalle", "ProveedorDetalle")
    n = CopiarTabla("IVUnidad", "Unidad")
    '***Agregado. 26/05/05. jeaa. Tabla descto pot item x cliente
    n = CopiarTabla("TSRetAutoDetalle", "Retenciones AutoDetalle")
    n = CopiarTabla("IVTipoCompra", "TipoCompra")
    n = CopiarTabla("IVUnidad", "Unidad")
    n = CopiarTabla("IVBanco", "IvBanco")
    n = CopiarTabla("IVTarjeta", "IvTarjeta")
    n = CopiarTabla("AFBodega", "AFBodega")
    n = CopiarTabla("AFGrupo1", "Catálogo de Grupo1 de Activo Fijo")
    n = CopiarTabla("AFGrupo2", "Catálogo de Grupo2 de Activo Fijo")
    n = CopiarTabla("AFGrupo3", "Catálogo de Grupo3 de Activo Fijo")
    n = CopiarTabla("AFGrupo4", "Catálogo de Grupo4 de Activo Fijo")
    n = CopiarTabla("AFGrupo5", "Catálogo de Grupo5 de Activo Fijo")
    n = CopiarTabla("AFInventario", "AFInventario")
    n = CopiarTabla("IVDescuento", "IVDescuento")
    n = CopiarTabla("IVDescuentoDetallePC", "IVDescuentoDetallePC")
    n = CopiarTabla("IVDescuentoDetalleIV", "IVDescuentoDetalleIV")
    n = CopiarTabla("IVDescuentoDetalleFC", "IVDescuentoDetalleFC")
    n = CopiarTabla("IVPromocion", "IVPromocion")
    n = CopiarTabla("IVCondPromocionDetalle", "IVCondPromocionDetalle")
    n = CopiarTabla("IVCondPromocionDetalleIVG", "IVCondPromocionDetalleIVG")
    n = CopiarTabla("IVCondPromocionDetalleP", "IVCondPromocionDetalleP")
    n = CopiarTabla("IVCondPromocionDetalleP", "IVCondPromocionDetalleP")
    'roles
    n = CopiarTabla("Elemento", "Elemento")
    n = CopiarTabla("Empleado", "Empleado")
    n = CopiarTabla("Personal", "Personal")
    n = CopiarTabla("TipoRol", "TipoRol")
    n = CopiarTabla("CuentaPcGrupo", "CuentaPcGrupo")
    n = CopiarTabla("CuentaPersonal", "CuentaPersonal")
    n = CopiarTabla("ImpuestoRenta", "ImpuestoRenta")
    
    n = CopiarTabla("Anexo_Comprobantes", "Anexo_Comprobantes")
    n = CopiarTabla("Anexo_Distrito", "Anexo_Distrito")
    n = CopiarTabla("Anexo_FormaPago", "Anexo_FormaPago")
    n = CopiarTabla("Anexo_ICE", "Anexo_ICE")
    n = CopiarTabla("Anexo_Pais", "Anexo_Pais")
    n = CopiarTabla("Anexo_Regimen", "Anexo_Regimen")
    n = CopiarTabla("Anexo_Sustentos", "Anexo_Sustentos")
    n = CopiarTabla("Anexo_TipoDocumento", "Anexo_TipoDocumento")
    n = CopiarTabla("Anexo_TipoExportacion", "Anexo_TipoExportacion")
    n = CopiarTabla("Anexo_RetencionIR", "Anexo_RetencionIR")
    n = CopiarTabla("Anexo_RetencionIVA", "Anexo_RetencionIVA")
    n = CopiarTabla("Anexo_Transacciones", "Anexo_Transacciones")
    
    n = CopiarTabla("gnsucursal", "gnsucursal")
    n = CopiarTabla("GnVehiculo", "GnVehiculo")
    n = CopiarTabla("GNVGrupo1", "GNVGrupo1")
    n = CopiarTabla("GNVGrupo2", "GNVGrupo2")
    n = CopiarTabla("GNVGrupo3", "GNVGrupo3")
    n = CopiarTabla("GNVGrupo4", "GNVGrupo4")
    n = CopiarTabla("GNTransporte", "GNTransporte")
    
    
    
    
     
'IVReservacion
    'Modifica la fecha de período contable y rango de fecha aceptable
    mensaje True, "Modificándo las fechas de inicio y fin."
    Set e = gobjMain.EmpresaActual
#If DAOLIB Then
#Else
    e.Coneccion.DefaultDatabase = Trim$(txtDestinoBD.Text)
#End If
    sql = "UPDATE GNOpcion SET FechaInicio=" & FechaYMD(e.GNOpcion.FechaFinal + 1, e.TipoDB) & ", " & _
                 "FechaFinal=" & _
                    FechaYMD(e.GNOpcion.FechaFinal _
                              + (e.GNOpcion.FechaFinal - e.GNOpcion.FechaInicio), e.TipoDB) & ", " & _
                 "FechaLimiteDesde=" & _
                    FechaYMD(e.GNOpcion.FechaFinal + 1, e.TipoDB) & ", " & _
                 "FechaLimiteHasta=" & _
                    FechaYMD(e.GNOpcion.FechaFinal _
                              + (e.GNOpcion.FechaFinal - e.GNOpcion.FechaInicio), e.TipoDB)
    e.EjecutarSQL sql, n
    
    sql = " update gnopcion2 set valor='A' where codigo='TramitesPosiblesSRI'"
    e.EjecutarSQL sql, n
    sql = " update gnopcion2 set valor='0' where codigo='RealizarReporteRangos'"
    e.EjecutarSQL sql, n
    sql = " update gnopcion2 set valor='0' where codigo='ValidacionAutoimpresores'"
    e.EjecutarSQL sql, n
    
#If DAOLIB Then
#Else
    e.Coneccion.DefaultDatabase = e.NombreDB
#End If
    mensaje False, "", "OK"

    CopiarDatos = True
    
salida:
    mbooProcesando = False               'Desbloquea que se cierre la ventana
    Set e = Nothing
    MensajeStatus
    Exit Function
ErrTrap:
    mensaje False, "", Err.Description
    MensajeStatus
    DispErr
    GoTo salida
End Function

Private Function CopiarTabla( _
                    ByVal tabla As String, _
                    ByVal Desc As String) As Long
    Dim sql As String, e As Empresa, Campos As String
    Dim BaseOrig As String, BaseDest As String, NumReg As Long
    Dim tiene_id As Boolean, n As Long
    On Error GoTo ErrTrap

    'Sacar mensaje
    MensajeStatus "Copiando " & Desc & " (" & tabla & ") ...", vbHourglass                          'GNVersion
    mensaje True, "Copiando '" & tabla & "'..."
    DoEvents
    
    BaseOrig = "[" & gobjMain.EmpresaActual.NombreDB & "].dbo.[" & tabla & "]"
    BaseDest = "[" & Trim$(txtDestinoBD.Text) & "].dbo.[" & tabla & "]"
    Set e = gobjMain.EmpresaActual
    
    'Obtiene lista de campos
    Campos = ObtenerCampos(e, tabla, tiene_id)
        
#If DAOLIB Then
    'Pendiente
#Else
    'Si tiene columna de identity (Autonumérico), activa la inserción con valor explícito en esa columna
    If tiene_id Then
        sql = "SET IDENTITY_INSERT " & BaseDest & " ON"
        e.EjecutarSQL sql, n
    End If
    
    'Primero elimina contenido de la tabla de destino
    sql = "DELETE FROM " & BaseDest
    e.EjecutarSQL sql, n

    'Copia los datos de la tabla
    sql = "INSERT INTO " & BaseDest & " (" & Campos & ") " & _
          "SELECT " & Campos & " FROM " & BaseOrig
    e.EjecutarSQL sql, NumReg

    If tiene_id Then
        sql = "SET IDENTITY_INSERT " & BaseDest & " OFF"
        e.EjecutarSQL sql, n
    End If
#End If
    
    mensaje False, "Copiado '" & tabla & "'.", NumReg & " registros."
    CopiarTabla = NumReg
    
salida:
    MensajeStatus
    Set e = Nothing
    Exit Function
ErrTrap:
    MensajeStatus
    mensaje False, "", Err.Description
    DispErr
    If tiene_id Then
    sql = "SET IDENTITY_INSERT " & BaseDest & " OFF"
    e.EjecutarSQL sql, n
    End If
    GoTo salida
End Function

'Obtiene nombre de todos los campos de una tabla
' y devuelve en una cadena separado por comma
Private Function ObtenerCampos( _
                    ByVal e As Empresa, _
                    ByVal tabla As String, _
                    ByRef Identidad As Boolean) As String
    Dim s As String, sql As String, rs As Recordset
    
#If DAOLIB Then
    'Pendiente DAO
#Else
    sql = "sp_help " & tabla
    Set rs = e.OpenRecordset(sql)
    Set rs = rs.NextRecordset           'Salta al segundo conjunto
    With rs
        Do Until .EOF
            DoEvents
            If Len(s) > 0 Then s = s & ", "
            s = s & .Fields("Column_name")
            .MoveNext
        Loop
    End With
    
    'Verifica si tiene una columna de identidad (Autonumérico)
    Set rs = rs.NextRecordset           'Salta al tercer conjunto
    With rs
        Identidad = True
        Do Until .EOF
            DoEvents
            If InStr(.Fields("Identity"), "No identity") > 0 Then
                Identidad = False
                Exit Do
            End If
            .MoveNext
        Loop
        .Close
    End With
#End If
    Set rs = Nothing
    ObtenerCampos = s
End Function

'3. Resetear # de transacciones
Private Function ResetearNumTrans() As Boolean
    Dim sql As String, s As String, r As Boolean, n As Long
    On Error GoTo ErrTrap
    
    'Verifica las opciones
    If Not VerificarOpcion Then Exit Function
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    
    s = "Este proceso es opcional, " & _
      "desea resetear los números de transacciones?" & vbCr & vbCr & _
      "Haga click en 'No' si quiere que se sigan las " & _
      "numeraciones de transacciones. En caso contrario, " & _
      "todas las transacciones comenzarán desde el número que usted indica " & _
      "en el siguiente paso."
    If MsgBox(s, vbQuestion + vbYesNo) = vbYes Then
OtraVez:
        s = InputBox("Ingrese el número con el que comienza las " & _
            "transacciones en la nueva base de datos.", _
            "Número de transacciones", 1)
        If Len(s) = 0 Then
            MsgBox "No se resetearán los números de transacciones.", vbInformation
        Else
            If Not IsNumeric(s) Then
                MsgBox "Ingrese un valor numérico, por favor.", vbCritical
                GoTo OtraVez
            End If
            
            'Comienza el proceso
            Me.MousePointer = vbHourglass
            mensaje True, "Reseteándo números de transacciones"
            
            n = Val(s)
            sql = "UPDATE [" & Trim$(txtDestinoBD.Text) & "].dbo.GNTrans SET NumTransSiguiente=" & n
            
            'Abre la empresa destino
            gobjMain.EmpresaActual.EjecutarSQL sql, n
            r = True
        End If
    Else
        r = True
    End If
    
    mensaje False, "", "OK"
    ResetearNumTrans = r
    
salida:
    mbooProcesando = False              'Desbloquea que se cierre la ventana
    Me.MousePointer = vbNormal
    Exit Function
ErrTrap:
    mensaje False, "", Err.Description
    DispErr
    GoTo salida
End Function

Private Sub mensaje( _
                ByVal NuevaFila As Boolean, _
                ByVal proc As String, _
                Optional ByVal res As String)
    Dim rpos As Long
    
    rpos = grd.Rows - 1
    
    If NuevaFila Then
        grd.AddItem "" & vbTab & proc & vbTab & res
    Else
        If Len(proc) > 0 Then
            grd.TextMatrix(rpos, 1) = proc
        ElseIf Right$(grd.TextMatrix(rpos, 1), 3) = "..." Then
            'Quitar último '...'
            grd.TextMatrix(rpos, 1) = Left$(grd.TextMatrix(rpos, 1), Len(grd.TextMatrix(rpos, 1)) - 3)
        End If
        
        If Len(res) > 0 Then
            grd.TextMatrix(rpos, 2) = res
        End If
    End If
End Sub

Private Function AbrirDestino() As Empresa
    Dim e As Empresa, cod As String
    
    cod = Trim$(txtDestino.Text)
    Set e = gobjMain.RecuperaEmpresa(cod)
    e.Abrir
    Set AbrirDestino = e
    Set e = Nothing
End Function





'6. Pasar saldo inicial de bancos
Private Function SaldoTS() As Boolean
    Dim e As Empresa, tsk As TSKardex
    Dim j As Long, sql As String, rs As Recordset
    Dim i As Long, c As Currency, Fcorte As Date
    Dim gcIT As GNComprobante, gcET As GNComprobante
    On Error GoTo ErrTrap
    
    'Verifica las opciones
    If Not VerificarOpcion Then Exit Function
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    Fcorte = dtpFechaCorte.value    'Fecha de corte

    'Cambia figura de cursor de mouse
    MensajeStatus "Está preparando saldos a la fecha de corte...", vbHourglass
    mensaje True, "Saldo inicial de bancos..."
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True

    'Obtiene Saldos de bancos a la fecha de corte en DOLARES
    ' No incluye documentos postfechados
    sql = "SELECT ts.CodBanco, " & _
                 "Sum((tsk.Debe-tsk.Haber)/gc.Cotizacion2) AS Saldo " & _
          "FROM (GNComprobante gc INNER JOIN " & _
                    "(TSKardex tsk INNER JOIN TSBanco ts " & _
                    "ON tsk.IdBanco=ts.IdBanco) " & _
                "ON tsk.TransID = gc.TransID) " & _
          "WHERE (gc.Estado <> " & ESTADO_ANULADO & ") AND " & _
                "(tsk.FechaVenci < " & FechaYMD(Fcorte + 1, mEmpOrigen.TipoDB) & ") AND " & _
                "(gc.FechaTrans < " & FechaYMD(Fcorte + 1, mEmpOrigen.TipoDB) & ")" & _
          "GROUP BY ts.CodBanco " & _
          "HAVING (Sum((tsk.Debe-tsk.Haber)/gc.Cotizacion2) <> 0) " & _
          "ORDER BY ts.CodBanco"
    Set rs = mEmpOrigen.OpenRecordset(sql)
    
    'Abre la empresa destino
    Set e = AbrirDestino
    
    With rs
        If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
        i = 0
        Do Until .EOF
            prg1.value = rs.AbsolutePosition
            prg1.Refresh
            DoEvents
            MensajeStatus "Agregando detalle: #" & i & " de " & rs.RecordCount, vbHourglass
            
            'Si aplastó 'Cancelar'
            If mbooCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                GoTo cancelado
            End If
            
            'Si Saldo es positivo
            If .Fields("Saldo") > 0 Then
                Set tsk = PrepararTransTS(e, "IT", _
                            "Saldo inicial de bancos", _
                            Fcorte, gcIT)
            'Si Saldo es negativo
            ElseIf .Fields("Saldo") < 0 Then
                Set tsk = PrepararTransTS(e, "ET", _
                            "Saldo inicial de bancos (Negativos)", _
                            Fcorte, gcET)
            End If
            
            'Recupera datos de proveedor y asigna al objeto
            tsk.codBanco = .Fields("CodBanco")
            If .Fields("Saldo") > 0 Then
                tsk.Debe = .Fields("Saldo")
                tsk.CodTipoDoc = "NC"
            Else
                tsk.Haber = .Fields("Saldo") * -1
                tsk.CodTipoDoc = "ND"
            End If
            tsk.FechaEmision = tsk.GNComprobante.FechaTrans
            tsk.FechaVenci = tsk.FechaEmision
            tsk.nombre = "Saldo inicial"
            tsk.numdoc = "S/I"
            tsk.Observacion = ""
            tsk.Orden = i
            
            i = i + 1
            .MoveNext
        Loop
        .Close
    End With
    
    'Graba la transacción si no están grabadas
    MensajeStatus "Grabándo la transacción...", vbHourglass
    If Not (gcIT Is Nothing) Then gcIT.Grabar False, False
    If Not (gcET Is Nothing) Then gcET.Grabar False, False
    Set gcIT = Nothing
    Set gcET = Nothing
    
    'Obtiene documentos postfechados para pasarlos uno por uno
    sql = "SELECT tsk.*, ts.CodBanco, tsd.CodTipoDoc, gc.Cotizacion2 " & _
          "FROM TSTipoDocBanco tsd INNER JOIN " & _
                    "(GNComprobante gc INNER JOIN " & _
                        "(TSKardex tsk INNER JOIN TSBanco ts " & _
                        "ON tsk.IdBanco=ts.IdBanco) " & _
                    "ON tsk.TransID = gc.TransID) " & _
                "ON tsd.IdTipoDoc = tsk.IdTipoDoc " & _
          "WHERE (gc.Estado <> " & ESTADO_ANULADO & ") AND " & _
                "(tsk.FechaVenci >= " & FechaYMD(Fcorte + 1, mEmpOrigen.TipoDB) & ") AND " & _
                "(gc.FechaTrans < " & FechaYMD(Fcorte + 1, mEmpOrigen.TipoDB) & ")" & _
          "ORDER BY ts.CodBanco"
    Set rs = mEmpOrigen.OpenRecordset(sql)
    With rs
        If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
        i = 0
        Do Until .EOF
            prg1.value = rs.AbsolutePosition
            prg1.Refresh
            DoEvents
            MensajeStatus "Agregando detalle de postfechados: #" & i & " de " & rs.RecordCount, vbHourglass
            
            'Si aplastó 'Cancelar'
            If mbooCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                GoTo cancelado
            End If
            
            'Si Saldo es positivo
            If .Fields("Debe") > 0 Then
                Set tsk = PrepararTransTS(e, "IT", _
                            "Saldo inicial de bancos (Cheques recibidos)", _
                            Fcorte, gcIT)
            'Si Saldo es negativo
            ElseIf .Fields("Haber") > 0 Then
                Set tsk = PrepararTransTS(e, "ET", _
                            "Saldo inicial de bancos (Cheques emitidos)", _
                            Fcorte, gcET)
            End If
            
            'Recupera datos de proveedor y asigna al objeto
            tsk.codBanco = .Fields("CodBanco")
            tsk.CodTipoDoc = .Fields("CodTipoDoc")
            If .Fields("Debe") > 0 Then
                tsk.Debe = .Fields("Debe") / .Fields("Cotizacion2")
            Else
                tsk.Haber = .Fields("Haber") / .Fields("Cotizacion2")
            End If
            tsk.FechaEmision = .Fields("FechaEmision")
            tsk.FechaVenci = .Fields("FechaVenci")
            tsk.nombre = .Fields("Nombre")
            tsk.numdoc = .Fields("NumDoc")
            tsk.Observacion = .Fields("Observacion")
            tsk.Orden = i
            
            i = i + 1
            .MoveNext
        Loop
        .Close
    End With
    
    'Graba la transacción si no están grabadas
    MensajeStatus "Grabándo la transacción...", vbHourglass
    If Not (gcIT Is Nothing) Then gcIT.Grabar False, False
    If Not (gcET Is Nothing) Then gcET.Grabar False, False

    MensajeStatus
    mensaje False, "", "OK"
    MsgBox "El proceso terminó con éxito.", vbInformation
    SaldoTS = True
    
cancelado:
    Set rs = Nothing
    MensajeStatus
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
    
    'Libera los objetos utilizados
    Set tsk = Nothing
    Set gcIT = Nothing
    Set gcET = Nothing
    Set e = Nothing
    
    mbooProcesando = False                  'Desbloquea que se cierre la ventana
    Exit Function
ErrTrap:
    mensaje False, "", Err.Description
    MensajeStatus
    DispErr
    GoTo cancelado
End Function

'Agrega un detalle de TSKardex a GNComprobante
'Si comprobante llega a tener 100 detalles,
'Graba lo anterior y crea otra instancia
Private Function PrepararTransTS(ByVal e As Empresa, _
                            ByVal codt As String, _
                            ByVal Desc As String, _
                            ByVal Fcorte As Date, _
                            ByRef gc As GNComprobante) As TSKardex
    Dim j As Long
                            
    'Crea transaccion si no existe todavía
    If gc Is Nothing Then
        Set gc = CrearTrans(e, codt, Desc, Fcorte, "")
    End If
    
    'Si llega a tener 100 detalles
    If gc.CountTSKardex >= 100 Then
        'Graba la transacción
        MensajeStatus "Grabándo la transacción...", vbHourglass
        gc.Grabar False, False
        
        'Crea nueva instancia de GNComprobante
        Set gc = CrearTrans(e, codt, Desc, Fcorte, "")
    End If

    'Agrega detalle
    j = gc.AddTSKardex
    Set PrepararTransTS = gc.TSKardex(j)
        
End Function


'7. Pasar saldo inicial de cuenta contable
Private Function SaldoCT() As Boolean
    Dim e As Empresa, ctd As CTLibroDetalle
    Dim j As Long, sql As String, rs As Recordset
    Dim i As Long, c As Currency, Fcorte As Date
    Dim gc As GNComprobante
    On Error GoTo ErrTrap
    
    'Verifica las opciones
    If Not VerificarOpcion Then Exit Function
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    Fcorte = dtpFechaCorte.value    'Fecha de corte

    'Cambia figura de cursor de mouse
    MensajeStatus "Está preparando saldos a la fecha de corte...", vbHourglass
    mensaje True, "Saldo inicial de cuentas contables..."
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True

    'Obtiene Saldos de asiento contable
    sql = "SELECT ct.CodCuenta, " & _
                 "Sum((ctd.Debe-ctd.Haber)/gc.Cotizacion2) AS Saldo " & _
          "FROM (GNComprobante gc INNER JOIN " & _
                    "(CTLibroDetalle ctd INNER JOIN CTCuenta ct " & _
                    "ON ctd.IdCuenta=ct.IdCuenta) " & _
                "ON ctd.CodAsiento = gc.CodAsiento) " & _
          "WHERE (gc.Estado IN (" & ESTADO_APROBADO & ", " & ESTADO_DESPACHADO & ", " & ESTADO_SEMDESPACHADO & ")) AND " & _
                "(ct.TipoCuenta IN (1, 2, 3)) AND " & _
                "(gc.FechaTrans <" & FechaYMD(Fcorte + 1, mEmpOrigen.TipoDB) & ") " & _
          "GROUP BY ct.CodCuenta " & _
          "HAVING (Sum((ctd.Debe-ctd.Haber)/gc.Cotizacion2) <> 0) " & _
          "ORDER BY ct.CodCuenta"
    Set rs = mEmpOrigen.OpenRecordset(sql)
    
    'Abre la empresa destino
    Set e = AbrirDestino
    
    With rs
        If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
        i = 0
        Do Until .EOF
            prg1.value = rs.AbsolutePosition
            prg1.Refresh
            DoEvents
            MensajeStatus "Agregando detalle: #" & i & " de " & rs.RecordCount, vbHourglass
            
            'Si aplastó 'Cancelar'
            If mbooCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                GoTo cancelado
            End If
            
            Set ctd = PrepararTransCT(e, "CTD", _
                        "Saldo inicial", _
                        Fcorte, gc, False)
            ctd.codcuenta = .Fields("CodCuenta")
            ctd.Debe = .Fields("Saldo")
            ctd.Descripcion = gc.Descripcion
            ctd.Orden = i
            
            i = i + 1
            .MoveNext
        Loop
        .Close
    End With
    
    'Graba la transacción si no están grabadas
    If Not (gc Is Nothing) Then GrabarTransCT gc, False
    
    MensajeStatus
    mensaje False, "", "OK"
    MsgBox "El proceso terminó con éxito.", vbInformation
    SaldoCT = True
    
cancelado:
    Set rs = Nothing
    MensajeStatus
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
    
    'Libera los objetos utilizados
    Set ctd = Nothing
    Set gc = Nothing
    Set e = Nothing
    
    mbooProcesando = False                 'Desbloquea que se cierre la ventana
    Exit Function
ErrTrap:
    mensaje False, "", Err.Description
    MensajeStatus
    DispErr
    GoTo cancelado
End Function


'8. Pasar trans. existentes con la fecha posterior a la fecha de corte
Private Function CopiaTrans() As Boolean
    Dim i As Long, codt As String, numt As Long
    Dim empDestino As Empresa, sql As String, Num As Long
    'Verifica  errores  en la base de Origen
    
    Set empDestino = AbrirDestino
    If empDestino.NombreDB = mEmpOrigen.NombreDB Then
        MsgBox "La empresa origen y destino son las mismas" & Chr(13) & _
               "debera  seleccionar  una empresa de  destino diferente", vbExclamation
        Exit Function
    End If
    If grd.ColKey(grd.Cols - 1) <> "Resultado" Then
        If VerificaFechaVenci = False Then
            MsgBox "Se han encontrado  errores  en las siguientes  transacciones " & Chr(13) & _
                   "La fecha de vencimiento no  puede ser mayor  a la fecha de transacción " & Chr(13) & _
                    "Primero deberá  corregirlos  para  proceder a copiar las transacciones.", vbInformation
                   
            Exit Function
        End If
    End If
    If grd.FixedRows = grd.Rows Then CargaTrans 'Carga  trans solo  si la grlla esta vacia
    'Transferir  transaccion una por una
    If MsgBox("Este proceso tardará  algunos minutos " & Chr(13) & " Desea comenzar el proceso de importación?", _
                vbYesNo + vbQuestion) <> vbYes Then Exit Function
    
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True
    
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    MensajeStatus "Copiando...", vbHourglass
    
    With grd
        prg1.min = .FixedRows - 1
        prg1.max = .Rows - 1
        prg1.value = prg1.min

        For i = .FixedRows To .Rows - 1
            prg1.value = i
            DoEvents                'Para dar control a Windows
            'Si usuario aplastó 'Cancelar', sale del ciclo
            If mbooCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                GoTo cancelado
            End If
            .ShowCell i, 0          'Hace visible la fila actual
            
'            If .IsSelected(i) Then
                codt = .TextMatrix(i, .ColIndex("CodTrans"))
                numt = .TextMatrix(i, .ColIndex("NumTrans"))
                'If codt = "FC" And numt = 2524 Then Stop
                MensajeStatus "Copiando la transacción " & codt & numt & _
                            "     " & i & " de " & .Rows - .FixedRows & _
                            " (" & Format(i * 100 / (.Rows - .FixedRows), "0") & "%)", vbHourglass
                
                'Si aún no está importado bien, importa la fila
                If grd.TextMatrix(i, .Cols - 1) <> MSG_OK Then
                    If ImportarTransSub(codt, numt, empDestino) Then
                        .TextMatrix(i, .ColIndex("Resultado")) = MSG_OK
                    Else
                        .TextMatrix(i, .ColIndex("Resultado")) = "Error"
                    End If
               End If
       Next i
    End With
    'Corregir  error de Idasignado
    MensajeStatus "Reasignando relaciones ...", vbHourglass
    sql = " UPDATE b SET b.IdAsignado = c.Id " & _
           " From    " & _
           empDestino.NombreDB & ".dbo.PCKardex c INNER JOIN " & _
           mEmpOrigen.NombreDB & ".dbo.PCKardex a INNER JOIN " & empDestino.NombreDB & ".dbo.PCKardex b " & _
           " ON a.Id  = b.IdAsignado " & _
           " ON c.Guid = a.Guid " & _
           " Where a.idAsignado = 0 And b.idAsignado <> 0 And c.idAsignado = 0 "
    
    mEmpOrigen.EjecutarSQL sql, Num
    MsgBox "Proceso terminado con exito"
    
    
    MensajeStatus
    mbooProcesando = False  'Bloquea que se cierre la ventana
    CopiaTrans = True
cancelado:
    MensajeStatus
    mbooProcesando = False
    prg1.value = prg1.min
    Exit Function
ErrTrap:
    MensajeStatus
    DispErr
    mbooProcesando = False
    prg1.value = prg1.min
    Exit Function
End Function


Private Function VerificaFechaVenci() As Boolean
    Dim sql As String, rs As Recordset, v  As Variant
    Dim Fcorte As Date
    
    Fcorte = dtpFechaCorte.value
    MensajeStatus "Verificando datos en Base Origen.....", vbHourglass
    sql = "Select FechaTrans, CodTrans, Numtrans, 'Error fecha vencimiento' as Estado " & _
          "From PCKardex  PCK Inner Join GnComprobante GNC On PCK.TransID = GNC.TransID " & _
          "Where GNC.Fechatrans > PCK.FechaVenci AND GNC.FechaTrans > " & FechaYMD(Fcorte, mEmpOrigen.TipoDB)
    
    Set rs = mEmpOrigen.OpenRecordset(sql)
    grd.Rows = grd.FixedRows
    If Not rs.EOF Then
        v = MiGetRows(rs)
        With grd
            .Redraw = flexRDNone
            .LoadArray v            'Carga a la grilla
        
            .FormatString = "^#|<Fecha|<CodTrans|<NumTrans|^Resultado"
            AjustarAutoSize grd, -1, -1, 3000
            GNPoneNumFila grd, False
            MensajeStatus "Errores  en la base origen"
            .Redraw = flexRDBuffered
            VerificaFechaVenci = False
            
       End With
    Else
        MensajeStatus
        
        VerificaFechaVenci = True
    End If
End Function



Private Function ImportarTransSub( _
                ByVal codt As String, _
                ByVal numt As Long, ByRef empDestino As Empresa) As Boolean
    Dim gnDest As GNComprobante, s As String, Estado As Byte, gnOri As GNComprobante
    On Error GoTo ErrTrap
'    Abre la empresa destino
    Set gnOri = mEmpOrigen.RecuperaGNComprobante(0, codt, numt)
'    Si existe en el destino, sobreescribe
    Set gnDest = empDestino.RecuperaGNComprobante(0, codt, numt)
    If (gnDest Is Nothing) Then
        Set gnDest = empDestino.CreaGNComprobante(codt)    'Crea  gnComprobante
    End If
    Estado = gnOri.Estado

    gnDest.Clone gnOri
    gnDest.Grabar False, False
    
'    Forzar el valor de Estado original, debido a que al Grabar cambia sin querer
    On Error Resume Next
    If gnDest.Estado = 1 And Estado = 3 Then
        'Primero Cambia  a estado cero
        empDestino.CambiaEstadoGNCompCierre gnDest.TransID, 0
    End If
    'Para  que no  considere  el IdAsignado
    empDestino.CambiaEstadoGNCompCierre gnDest.TransID, Estado
    ImportarTransSub = True
salida:
    Set gnDest = Nothing
    Set gnOri = Nothing
    Exit Function
ErrTrap:
'    DispMsg "Importar la trans. " & codt & numt, "Error", Err.Description
    If MsgBox(Err.Description & vbCr & vbCr & _
                "Desea continuar con siguiente transacción?", _
                vbQuestion + vbYesNo) <> vbYes Then
'        mCancelado = True
    End If
    GoTo salida
End Function



Private Sub CargaTrans()
    Dim sql As String, rs As Recordset, v As Variant
    Dim Fcorte As Date
    On Error GoTo ErrTrap
    
    Fcorte = dtpFechaCorte.value    'Fecha de corte
    'Selecciona las transacciones de la  base de origen
    sql = "SELECT FechaTrans, CodTrans, NumTrans, Descripcion " & _
          "FROM GNComprobante " & _
                " Where FechaTrans > " & FechaYMD(Fcorte, mEmpOrigen.TipoDB) & _
                " ORDER BY FechaTrans"
' 10/12/2004  antes estaba el orden tambien opor codtrans, y numtrans
          
    'Set rs = New Recordset
    Set rs = mEmpOrigen.OpenRecordset(sql)
    If Not rs.EOF Then
        v = MiGetRows(rs)
        With grd
            .Redraw = flexRDNone
            .LoadArray v            'Carga a la grilla
            .FormatString = "^#|<Fecha|<CodTrans|<NumTrans|<Descripción|<Resultado"
            GNPoneNumFila grd, False
            AsignarTituloAColKey grd            'Para usar ColIndex
            AjustarAutoSize grd, 0, -1, 3000     'Ajusta automáticamente ancho de cols.
            If .ColWidth(.ColIndex("Descripción")) > 1400 Then .ColWidth(.ColIndex("Descripción")) = 1400
            .ColWidth(.ColIndex("Resultado")) = 1600
            'Tipo de datos
            .ColDataType(.ColIndex("Fecha")) = flexDTDate
            .ColDataType(.ColIndex("CodTrans")) = flexDTString
            .ColDataType(.ColIndex("NumTrans")) = flexDTLong
            .ColDataType(.ColIndex("Descripción")) = flexDTString
            '.ColDataType(.ColIndex("Cod.C.C.")) = flexDTString
            '.ColDataType(.ColIndex("Estado")) = flexDTShort
            .ColDataType(.ColIndex("Resultado")) = flexDTString
            
            .Redraw = flexRDDirect
        End With
        
    Else
        'Si no hay nada de resultado limpia la grilla
        grd.Rows = grd.FixedRows
    End If
    rs.Close
'    mBuscado = True
salida:
    MensajeStatus
    Set rs = Nothing
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    GoTo salida
End Sub

'9. Desactivar la base de datos origen
Private Function DesactivarOrigen() As Boolean
    Dim v As Variant, i As Long, codg As String
    Dim g As grupo, p As Permiso, j As Long, k As Long, pt As PermisoTrans
    Dim s As String
    On Error GoTo ErrTrap
    
    'Verifica las opciones
    If Not VerificarOpcion Then Exit Function
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    v = gobjMain.ListaGrupos(True)
    If Not IsEmpty(v) Then
        For i = LBound(v, 2) To UBound(v, 2)
            codg = v(0, i)
            MensajeStatus "Procesando grupo '" & codg & "'...", vbHourglass
            mensaje True, "Procesando grupo '" & codg & "'"
            Set g = gobjMain.RecuperaGrupo(codg)
            Set mGrupo = g          'Para recibir evento 'mGrupo_Procesando'
            
            For j = 1 To g.CountPermiso
                DoEvents
                Set p = g.Permisos(j)
                'Si el grupo tiene permiso para la empresa actual
                If UCase(p.CodEmpresa) = UCase(mEmpOrigen.CodEmpresa) Then
                    'Confirmación
                    s = "Desea bloquear la modificación de datos de la empresa '" & _
                        mEmpOrigen.CodEmpresa & " (" & mEmpOrigen.Descripcion & ")' " & _
                        "para el grupo de usuario '" & codg & "'?" & vbCr & vbCr & _
                        "Si aplasta 'Sí', los usuarios que pertenecen al grupo " & _
                        "no podrán realizar ningún cambio a los datos de la dicha empresa. " & vbCr & _
                        "Sin embargo si fuera necesario se podrá desbloquear de nuevo " & _
                        "usando el programa 'SiiConfig' con código de usuario que tenga " & _
                        "derecho de supervisor."
                    If MsgBox(s, vbYesNo + vbQuestion) = vbYes Then
                        'Bloquea modificación/creación de todas las transacciones
                        For k = 1 To p.CountTrans
                            Set pt = p.trans(k)
                            With pt
                                .Anular = False
                                .Aprobar = False
                                .crear = False
                                .Desaprobar = False
                                .Despachar = False
                                .Eliminar = False
                                .Modificar = False
    '                            .Ver = False           'Permiso para ver no desactivemos
                            End With
                            Set pt = Nothing
                        Next k
                        
                        'Bloquea modificación de todos los catálogos
                        With p
                            .CatAFMod = False
                            .CatBancoMod = False
                            .CatBodegaMod = False
                            .CatCentroCostoMod = False
                            .CatClienteMod = False
                            .CatInfEmpresaMod = False
                            .CatInventarioMod = False
                            .CatInventarioPrecioMod = False
                            .CatPlanCuentaMod = False
                            .CatProveedorMod = False
                            .CatResponsableMod = False
                            .CatRolMod = False
                            .CatVendedorMod = False
                        End With
                    End If
                    
                    Exit For        'Pasa a procesar siguiente grupo
                End If
                Set p = Nothing
            Next j
            
            'Si el grupo está modificado, graba el grupo
            If g.Modificado Then
                MensajeStatus "Grabando grupo '" & codg & "'...", vbHourglass
                mensaje False, "Grabando grupo '" & codg & "'"
                g.Grabar
                mensaje False, "", "OK. Desactivado."
            Else
                mensaje False, "", "OK."
            End If
            Set g = Nothing
        Next i
    End If
    
    DesactivarOrigen = True
    
salida:
    mbooProcesando = False               'Desbloquea que se cierre la ventana
    Set pt = Nothing
    Set p = Nothing
    Set g = Nothing
    Set mGrupo = Nothing
    MensajeStatus
    Exit Function
ErrTrap:
    MensajeStatus
    DispErr
    GoTo salida
End Function



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
    dtpFechaCorte.value = mEmpOrigen.GNOpcion.FechaFinal
    
    'Visualiza codigo de empresa origen (= Empresa actual)
    lblOrigen.Caption = mEmpOrigen.CodEmpresa
    lblOrigenBD.Caption = mEmpOrigen.NombreDB
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = mbooProcesando
End Sub

Private Sub Form_Resize()
'    On Error Resume Next
'    grd.Move 0, grd.Top, Me.ScaleWidth, Me.ScaleHeight - grd.Top
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





'4. Pasar saldo inicial de inventario
Private Function SaldoIV() As Boolean
    Dim e As Empresa, gc As GNComprobante, ivk As IVKardex, iv As IVinventario
    Dim j As Long, n As Long
    Dim sql As String, rs As Recordset, codOrig As String
    Dim i As Long, c As Currency, Fcorte As Date
    On Error GoTo ErrTrap
    
    'Verifica las opciones
    If Not VerificarOpcion Then Exit Function
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    
    codOrig = gobjMain.EmpresaActual.CodEmpresa
    Fcorte = dtpFechaCorte.value    'Fecha de corte

    'Cambia figura de cursor de mouse
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True
    
    'Saca las existencias a la fecha de corte
    MensajeStatus "Preparando para grabar las existencias iniciales...", vbHourglass
    mensaje True, "Saldo inicial de inventario..."
    
    sql = "SELECT ivk.IdInventario, ivk.IdBodega, " & _
                "iv.CodInventario, ivb.CodBodega, " & _
                "Sum(ivk.Cantidad) AS Exist " & _
          "FROM IVBodega ivb INNER JOIN " & _
                    "(IVInventario iv INNER JOIN " & _
                        "(GNTrans gt INNER JOIN " & _
                            "(IVKardex ivk INNER JOIN GNComprobante gc " & _
                            "ON ivk.TransID=gc.TransID) " & _
                        "ON gt.CodTrans=gc.CodTrans) " & _
                    "ON iv.IdInventario = ivk.IdInventario) " & _
                "ON ivb.IdBodega = ivk.IdBodega " & _
          "WHERE (gc.Estado<>" & ESTADO_ANULADO & ") AND " & _
                 "(gt.AfectaCantidad=" & CadenaBool(True, gobjMain.EmpresaActual.TipoDB) & ") AND " & _
                 "(gc.FechaTrans < " & FechaYMD(Fcorte + 1, gobjMain.EmpresaActual.TipoDB) & ") AND " & _
                 "(iv.BandServicio=" & CadenaBool(False, gobjMain.EmpresaActual.TipoDB) & ") " & _
          "GROUP BY ivk.IdInventario, ivk.IdBodega, iv.CodInventario, ivb.CodBodega " & _
          "HAVING Sum(ivk.Cantidad)>0"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
#If DAOLIB = 0 Then
    Set rs.ActiveConnection = Nothing
#End If
    
    'Abre la empresa destino
    Set e = AbrirDestino
    
    With rs
        If Not rs.EOF Then
            rs.MoveLast
            rs.MoveFirst
            If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
            i = 0
            Do Until .EOF
                prg1.value = rs.AbsolutePosition
                prg1.Refresh
                DoEvents
                
                'Si aplastó 'Cancelar'
                If mbooCancelado Then
                    MsgBox "El proceso fue cancelado.", vbInformation
                    GoTo cancelado
                End If
                
                'Crea transaccion 'IVSI'
                If (i Mod 100) = 0 Then
                    'Si no es primera vez
                    If Not (gc Is Nothing) Then
                        'Graba la transacción
                        MensajeStatus "Grabando la transacción en la empresa '" & gc.Empresa.CodEmpresa & "'...", vbHourglass
                        gc.Grabar False, False
                    End If
                    
                    Set gc = CrearTrans(e, _
                            "IVSI", _
                            "Saldo inicial de inventario", _
                            Fcorte, _
                            "")
                End If
                
                'Recupera datos de inventario para llama el método Costo()
                MensajeStatus "Agregando detalle #" & i & " de " & rs.RecordCount, vbHourglass
                Set iv = mEmpOrigen.RecuperaIVInventario(.Fields("IdInventario"))
                
                'Obtiene Costo del item en Moneda de item
                c = iv.costo(Fcorte, 1)
                
                'De moneda de item, covierte en moneda de trans, si es necesario
                If iv.CodMoneda <> gc.CodMoneda Then
                    c = c * gc.Cotizacion(iv.CodMoneda) / gc.Cotizacion("")
                End If
                
                'Agrega detalle
                j = gc.AddIVKardex
                Set ivk = gc.IVKardex(j)
                ivk.cantidad = .Fields("Exist")
                ivk.CodBodega = .Fields("CodBodega")
                ivk.CodInventario = .Fields("CodInventario")
                ivk.CostoRealTotal = c * ivk.cantidad
                ivk.CostoTotal = ivk.CostoRealTotal
                ivk.PrecioRealTotal = ivk.CostoRealTotal
                ivk.PrecioTotal = ivk.PrecioRealTotal
                ivk.Orden = i Mod 100
                i = i + 1
                .MoveNext
            Loop
        End If
        .Close
    End With
        
    If Not (gc Is Nothing) Then
        'Graba la transacción
        MensajeStatus "Grabando la transacción en la empresa '" & gc.Empresa.CodEmpresa & "'...", vbHourglass
        gc.Grabar False, False
    End If
    
    'Corrige las existencias para que quede bien la tabla 'IVExist'
    MensajeStatus "Arreglando las existencias...", vbHourglass
    If Not (gc Is Nothing) Then
        gc.Empresa.CorregirExistencia
    End If
    mensaje False, "", "OK"
    MensajeStatus
    MsgBox "El proceso terminó con éxito.", vbInformation
    SaldoIV = True
    
cancelado:
    mensaje False, "", Err.Description
    MensajeStatus
    Set ivk = Nothing
    Set iv = Nothing
    Set gc = Nothing
    Set rs = Nothing
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
    
    'Vuelve a abrir la empresa origen
    Set e = gobjMain.RecuperaEmpresa(codOrig)
    e.Abrir
    Set e = Nothing
    
    mbooProcesando = False                  'Desbloquea que se cierre la ventana
    Exit Function
ErrTrap:
    MensajeStatus
    MsgBox Err.Description, vbExclamation
    GoTo cancelado
End Function

Private Function CrearTrans(ByVal emp As Empresa, _
                            ByVal CodTrans As String, _
                            ByVal Desc As String, _
                            ByVal fecha As Date, _
                            ByVal numdoc As String) As GNComprobante
    Dim g As GNComprobante
    
    Set g = emp.CreaGNComprobante(CodTrans)
    With g
        .IdResponsable = 1
        .CodMoneda = "USD"
        .Cotizacion("USD") = 1   'Diego 20/08/2002
        .Descripcion = Desc
        .FechaTrans = fecha + 1
        .numDocRef = numdoc
    End With
    Set CrearTrans = g
    Set g = Nothing
End Function

'Agrega un detalle de PCKardex a GNComprobante
'Si comprobante llega a tener 100 detalles,
'Graba lo anterior y crea otra instancia
Private Function PrepararTransPC(ByVal e As Empresa, _
                            ByVal codt As String, _
                            ByVal Desc As String, _
                            ByVal Fcorte As Date, _
                            ByRef gc As GNComprobante) As PCKardex
    Dim j As Long
                            
    'Crea transaccion si no existe todavía
    If gc Is Nothing Then
        Set gc = CrearTrans(e, codt, Desc, Fcorte, "")
    End If
    
    'Si llega a tener 100 detalles
    If gc.CountPCKardex >= 100 Then
        'Graba la transacción
        MensajeStatus "Grabándo la transacción...", vbHourglass
        gc.Grabar False, False
        
        'Crea nueva instancia de GNComprobante
        Set gc = CrearTrans(e, codt, Desc, Fcorte, "")
    End If

    'Agrega detalle
    j = gc.AddPCKardex
    Set PrepararTransPC = gc.PCKardex(j)
End Function


'5. Pasar saldo inicial de proveedores/clientes
Private Function SaldoPC() As Boolean
    Dim e As Empresa, pck As PCKardex
    Dim j As Long, sql As String, rs As Recordset
    Dim i As Long, c As Currency, Fcorte As Date
    Dim gcPVNC As GNComprobante, gcPVND As GNComprobante
    Dim gcCLNC As GNComprobante, gcCLND As GNComprobante
    On Error GoTo ErrTrap
    
    'Verifica las opciones
    If Not VerificarOpcion Then Exit Function
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    Fcorte = dtpFechaCorte.value    'Fecha de corte

    'Cambia figura de cursor de mouse
    MensajeStatus "Está preparando saldos a la fecha de corte...", vbHourglass
    mensaje True, "Saldo inicial de proveedor/cliente..."
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True

    'Obtiene Saldos de proveedor/cliente por cada documento pendiente
    sql = "spConsPCSaldo3 2, " & FechaYMD(Fcorte, gobjMain.EmpresaActual.TipoDB)
    Set rs = mEmpOrigen.OpenRecordset(sql)
    UltimoRecordset rs
    
    'Abre la empresa destino
    Set e = AbrirDestino
    
    With rs
        If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
        i = 0
        Do Until .EOF
            prg1.value = rs.AbsolutePosition
            prg1.Refresh
            DoEvents
            MensajeStatus "Agregando detalle: #" & i & " de " & rs.RecordCount, vbHourglass
            
            'Si aplastó 'Cancelar'
            If mbooCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                GoTo cancelado
            End If
            
            'Si es Proveedores por cobrar (Anticipado)
            If .Fields("Saldo") > 0 And .Fields("BandProveedor") <> 0 Then
                Set pck = PrepararTransPC(e, "PVND", _
                            "Saldo inicial de proveedores (Anticipos)", _
                            Fcorte, gcPVND)
            'Si es Proveedores por pagar
            ElseIf .Fields("Saldo") < 0 And .Fields("BandProveedor") <> 0 Then
                Set pck = PrepararTransPC(e, "PVNC", _
                            "Saldo inicial de proveedores x pagar", _
                            Fcorte, gcPVNC)
            'Si es Clientes por cobrar
            ElseIf .Fields("Saldo") > 0 And .Fields("BandProveedor") = 0 Then
                Set pck = PrepararTransPC(e, "CLND", _
                            "Saldo inicial de clientes x cobrar", _
                            Fcorte, gcCLND)
            'Si es Clientes por pagar (Anticipado)
            Else
                Set pck = PrepararTransPC(e, "CLNC", _
                            "Saldo inicial de cliente (Anticipos)", _
                            Fcorte, gcCLNC)
            End If
            
            'Recupera datos de proveedor y asigna al objeto
            pck.CodProvCli = .Fields("CodProvCli")
            If .Fields("Saldo") > 0 Then   'Si es por cobrar --> Debe
                pck.Debe = .Fields("Saldo")        'Saldo en dólares
            Else                                    'Si es por pagar --> Haber
                pck.Haber = .Fields("Saldo") * -1     'Saldo en dólares
            End If
            pck.codforma = .Fields("CodForma")
            pck.FechaEmision = .Fields("FechaEmision")
            pck.FechaVenci = .Fields("FechaVenci")
            pck.NumLetra = .Fields("Trans")
            pck.Observacion = .Fields("Observacion")
            pck.Orden = i
            pck.Guid = .Fields("Guid")        '*** <== AGREGAR ESTO
            If Not IsNull(.Fields("CodVendedor")) Then pck.CodVendedor = .Fields("codvendedor") 'AUC 04/06/07
            i = i + 1
            .MoveNext
        Loop
        .Close
    End With
    
    'Graba la transacción si no están grabadas
    MensajeStatus "Grabándo la transacción...", vbHourglass
    If Not (gcPVND Is Nothing) Then gcPVND.Grabar False, False
    If Not (gcPVNC Is Nothing) Then gcPVNC.Grabar False, False
    If Not (gcCLND Is Nothing) Then gcCLND.Grabar False, False
    If Not (gcCLNC Is Nothing) Then gcCLNC.Grabar False, False

    MensajeStatus
    mensaje False, "", "OK"
    MsgBox "El proceso terminó con éxito.", vbInformation
    SaldoPC = True
    
cancelado:
    Set rs = Nothing
    MensajeStatus
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
    
    'Libera los objetos utilizados
    Set pck = Nothing
    Set gcPVND = Nothing
    Set gcPVNC = Nothing
    Set gcCLND = Nothing
    Set gcCLNC = Nothing
    Set e = Nothing
    
    mbooProcesando = False               'Desbloquea que se cierre la ventana
    Exit Function
ErrTrap:
    mensaje False, "", Err.Description
    MensajeStatus
    DispErr
    GoTo cancelado
End Function



Private Sub mGrupo_Procesando(ByVal msg As String)
    If Len(msg) > 0 Then
        MensajeStatus msg, vbHourglass
    Else
        MensajeStatus
    End If
    DoEvents
End Sub

Private Sub txtDestino_LostFocus()
    Dim Cancel As Boolean
    'Este es necesario porque al dar Enter no se genera el evento Validate
    txtDestino_Validate Cancel
    If Cancel Then txtDestino.SetFocus
End Sub

Private Sub txtDestino_Validate(Cancel As Boolean)
    Dim e As Empresa, cod As String
    On Error GoTo ErrTrap
    
    cod = Trim$(txtDestino.Text)
    Set e = gobjMain.RecuperaEmpresa(cod)
    If Not (e Is Nothing) Then
        txtDestinoBD.Text = e.NombreDB
    End If
    Exit Sub
ErrTrap:
    MsgBox "No se encuentra la empresa de destino. ('" & cod & "')", vbInformation
    Cancel = True
    Exit Sub
End Sub






'4a. Pasar saldo inicial de Activos Fijos
Private Function SaldoAF() As Boolean
    Dim e As Empresa, gc As GNComprobante, ivk As AFKardex, iv As AFinventario, IVD As AFinventario
    Dim j As Long, n As Long
    Dim sql As String, rs As Recordset, codOrig As String
    Dim i As Long, c As Currency, Fcorte As Date
    On Error GoTo ErrTrap
    
    'Verifica las opciones
    If Not VerificarOpcion Then Exit Function
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    
    codOrig = gobjMain.EmpresaActual.CodEmpresa
    Fcorte = dtpFechaCorte.value    'Fecha de corte

    'Cambia figura de cursor de mouse
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True
    
    'Saca las existencias a la fecha de corte
    MensajeStatus "Preparando para grabar los activos iniciales...", vbHourglass
    mensaje True, "Saldo inicial del activo fijo..."
    
    sql = "SELECT ivk.IdInventario, ivk.IdBodega, " & _
                "iv.CodInventario, ivb.CodBodega, " & _
                "Sum(ivk.Cantidad) AS Exist " & _
          "FROM AFBodega ivb INNER JOIN " & _
                    "(AFInventario iv INNER JOIN " & _
                        "(GNTrans gt INNER JOIN " & _
                            "(AFKardex ivk INNER JOIN GNComprobante gc " & _
                            "ON ivk.TransID=gc.TransID) " & _
                        "ON gt.CodTrans=gc.CodTrans) " & _
                    "ON iv.IdInventario = ivk.IdInventario) " & _
                "ON ivb.IdBodega = ivk.IdBodega " & _
          "WHERE (gc.Estado<>" & ESTADO_ANULADO & ") AND " & _
                 "(gt.AfectaCantidad=" & CadenaBool(True, gobjMain.EmpresaActual.TipoDB) & ") AND " & _
                 "(gc.FechaTrans < " & FechaYMD(Fcorte + 1, gobjMain.EmpresaActual.TipoDB) & ") AND " & _
                 "(iv.BandServicio=" & CadenaBool(False, gobjMain.EmpresaActual.TipoDB) & ") " & _
          "GROUP BY ivk.IdInventario, ivk.IdBodega, iv.CodInventario, ivb.CodBodega " & _
          "HAVING Sum(ivk.Cantidad)>0"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
#If DAOLIB = 0 Then
    Set rs.ActiveConnection = Nothing
#End If
    
    'Abre la empresa destino
    Set e = AbrirDestino
    
    With rs
        If Not rs.EOF Then
            rs.MoveLast
            rs.MoveFirst
            If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
            i = 0
            Do Until .EOF
                prg1.value = rs.AbsolutePosition
                prg1.Refresh
                DoEvents
                
                'Si aplastó 'Cancelar'
                If mbooCancelado Then
                    MsgBox "El proceso fue cancelado.", vbInformation
                    GoTo cancelado
                End If
                
                'Crea transaccion 'IVSAF'
                If (i Mod 100) = 0 Then
                    'Si no es primera vez
                    If Not (gc Is Nothing) Then
                        'Graba la transacción
                        MensajeStatus "Grabando la transacción en la empresa '" & gc.Empresa.CodEmpresa & "'...", vbHourglass
                        gc.Grabar False, False
                    End If
                    
                    Set gc = CrearTrans(e, _
                            "IVSAF", _
                            "Saldo inicial del activo Fijo", _
                            Fcorte, _
                            "")
                End If
                
                'Recupera datos de inventario para llama el método Costo()
                MensajeStatus "Agregando detalle #" & i & " de " & rs.RecordCount, vbHourglass
                Set iv = mEmpOrigen.RecuperaAFInventario(.Fields("IdInventario"))
                Set IVD = e.RecuperaAFInventario(.Fields("IdInventario"))
                
                'Obtiene Costo del item en Moneda de item
                c = iv.CostoUltimoIngreso '(Fcorte, 1)
                
                'De moneda de item, covierte en moneda de trans, si es necesario
                If iv.CodMoneda <> gc.CodMoneda Then
                    c = c * gc.Cotizacion(iv.CodMoneda) / gc.Cotizacion("")
                End If
                
                'Agrega detalle
                j = gc.AddAFKardex
                Set ivk = gc.AFKardex(j)
                ivk.cantidad = .Fields("Exist")
                ivk.CodBodega = .Fields("CodBodega")
                ivk.CodInventario = .Fields("CodInventario")
                ivk.CostoRealTotal = c * ivk.cantidad
                ivk.CostoTotal = ivk.CostoRealTotal
                ivk.PrecioRealTotal = ivk.CostoRealTotal
                ivk.PrecioTotal = ivk.PrecioRealTotal
                ivk.Orden = i Mod 100
                
                'ACTUALIZA DEPRECIACIONES
                
                IVD.DepAnterior = Val(iv.DepAnterior) + Val(iv.NumeroDepre)
                
                IVD.NumeroDepre = e.CalculaNumeroDepreciaciones(IVD.CodInventario)
                IVD.Grabar
                
                i = i + 1
                .MoveNext
            Loop
        End If
        .Close
    End With
        
    If Not (gc Is Nothing) Then
        'Graba la transacción
        MensajeStatus "Grabando la transacción en la empresa '" & gc.Empresa.CodEmpresa & "'...", vbHourglass
        gc.Grabar False, False
    End If
    
    'Corrige las existencias para que quede bien la tabla 'IVExist'
    MensajeStatus "Arreglando las existencias...", vbHourglass
    If Not (gc Is Nothing) Then
'        gc.Empresa.CorregirAFExistencia
    End If
    mensaje False, "", "OK"
    MensajeStatus
    MsgBox "El proceso terminó con éxito.", vbInformation
    SaldoAF = True
    
cancelado:
    mensaje False, "", Err.Description
    MensajeStatus
    Set ivk = Nothing
    Set iv = Nothing
    Set IVD = Nothing
    Set gc = Nothing
    Set rs = Nothing
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
    
    'Vuelve a abrir la empresa origen
    Set e = gobjMain.RecuperaEmpresa(codOrig)
    e.Abrir
    Set e = Nothing
    
    mbooProcesando = False                  'Desbloquea que se cierre la ventana
    Exit Function
ErrTrap:
    MensajeStatus
    MsgBox Err.Description, vbExclamation
    GoTo cancelado
End Function

Private Function CopiarHistorial() As Long
    Dim sql As String, e As Empresa, Campos As String
    Dim BaseOrig As String, BaseDest As String, NumReg As Long
    Dim tiene_id As Boolean, n As Long
    On Error GoTo ErrTrap
    'Sacar mensaje
    MensajeStatus "Copiando  Historial de Clientes...", vbHourglass                          'GNVersion
    mensaje True, "Copiando pcHistorial..."
    DoEvents
    BaseOrig = "[" & gobjMain.EmpresaActual.NombreDB & "].dbo.pcHistorial"
    BaseDest = "[" & Trim$(txtDestinoBD.Text) & "].dbo.pcHistorial"
    Set e = gobjMain.EmpresaActual
    'Obtiene lista de campos
    Campos = ObtenerCampos(e, "pchistorial", tiene_id)
#If DAOLIB Then
    'Pendiente
#Else
    'Si tiene columna de identity (Autonumérico), activa la inserción con valor explícito en esa columna
    If tiene_id Then
        sql = "SET IDENTITY_INSERT " & BaseDest & " ON"
        e.EjecutarSQL sql, n
    End If
    'Primero elimina contenido de la tabla de destino
    sql = "DELETE FROM " & BaseDest
    e.EjecutarSQL sql, n
    'Copia los datos de la tabla
    sql = "INSERT INTO " & BaseDest & " (" & Campos & ") " & _
          "SELECT " & Campos & " FROM " & BaseOrig
    e.EjecutarSQL sql, NumReg
    If tiene_id Then
        sql = "SET IDENTITY_INSERT " & BaseDest & " OFF"
        e.EjecutarSQL sql, n
    End If
#End If
    mensaje False, "Copiado pchistorial  .", NumReg & " registros."
    CopiarHistorial = NumReg
salida:
    MensajeStatus
    Set e = Nothing
    Exit Function
ErrTrap:
    MensajeStatus
    mensaje False, "", Err.Description
    DispErr
    If tiene_id Then
    sql = "SET IDENTITY_INSERT " & BaseDest & " OFF"
    e.EjecutarSQL sql, n
    End If
    GoTo salida
End Function

Private Function CopiarDatosFichas() As Boolean
    Dim sql As String, n As Long, e As Empresa, rpos As Long
    On Error GoTo ErrTrap
    
    'Verifica las opciones
    If Not VerificarOpcion Then Exit Function
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    
    n = CopiarTabla("fichaEnfermedad", "Enfermedades")
    n = CopiarTabla("FichaMedicaExmFisico", "Ficha Medicas")
    n = CopiarTabla("FichaMedEnfDetalle", "Detalle de Enfermedades")
    
    
    
    
     
'IVReservacion
    'Modifica la fecha de período contable y rango de fecha aceptable
    mensaje True, "Modificándo las fechas de inicio y fin."
    Set e = gobjMain.EmpresaActual
#If DAOLIB Then
#Else
    e.Coneccion.DefaultDatabase = Trim$(txtDestinoBD.Text)
#End If
#If DAOLIB Then
#Else
    e.Coneccion.DefaultDatabase = e.NombreDB
#End If
    mensaje False, "", "OK"

    CopiarDatosFichas = True
    
salida:
    mbooProcesando = False               'Desbloquea que se cierre la ventana
    Set e = Nothing
    MensajeStatus
    Exit Function
ErrTrap:
    mensaje False, "", Err.Description
    MensajeStatus
    DispErr
    GoTo salida
End Function


