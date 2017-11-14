VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCierreTransXEntregar 
   Caption         =   "Pendientes x Entregar"
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
   Begin VB.CommandButton cmdPasosArchivo 
      Caption         =   "GO"
      Height          =   330
      Left            =   5640
      TabIndex        =   21
      Top             =   1860
      Width           =   612
   End
   Begin VB.TextBox txtOrigen 
      Height          =   320
      Left            =   1320
      TabIndex        =   5
      Top             =   1860
      Width           =   3855
   End
   Begin VB.CommandButton cmdExplorar 
      Caption         =   "..."
      Height          =   310
      Left            =   5220
      TabIndex        =   6
      Top             =   1860
      Width           =   372
   End
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   9825
      TabIndex        =   16
      Top             =   7350
      Width           =   9825
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   288
         Left            =   9900
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   1212
      End
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   180
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Empresa destino"
      Height          =   1000
      Left            =   3000
      TabIndex        =   11
      Top             =   60
      Width           =   2772
      Begin VB.TextBox txtDestino 
         Height          =   300
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   1812
      End
      Begin VB.TextBox txtDestinoBD 
         Height          =   300
         Left            =   720
         TabIndex        =   2
         Top             =   600
         Width           =   1812
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Código  "
         Height          =   192
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "B.D."
         Height          =   192
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Empresa orígen"
      Height          =   1000
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   2772
      Begin VB.Label lblOrigenBD 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   720
         TabIndex        =   10
         Top             =   600
         Width           =   1812
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "B.D."
         Height          =   192
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Código  "
         Height          =   192
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblOrigen 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   720
         TabIndex        =   8
         Top             =   240
         Width           =   1812
      End
   End
   Begin VB.CommandButton cmdPasos 
      Caption         =   "GO"
      Height          =   330
      Left            =   5640
      TabIndex        =   4
      Top             =   1500
      Width           =   612
   End
   Begin MSComCtl2.DTPicker dtpFechaCorte 
      Height          =   300
      Left            =   3780
      TabIndex        =   3
      Top             =   1140
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      Format          =   123666433
      CurrentDate     =   36781
   End
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   6420
      Top             =   1440
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "mdb"
      DialogTitle     =   "Orígen de Importación"
   End
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   1815
      Left            =   120
      TabIndex        =   20
      Top             =   2280
      Width           =   5895
      _cx             =   10393
      _cy             =   3196
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
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      Editable        =   2
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   2
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Archivo Orígen:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   1920
      Width           =   1125
   End
   Begin VB.Label lblPasos 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pasar trans. pendientes de entrega hasta la fecha de corte "
      Height          =   330
      Left            =   120
      TabIndex        =   15
      Top             =   1500
      Width           =   5460
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de corte  "
      Height          =   195
      Left            =   2460
      TabIndex        =   14
      Top             =   1200
      Width           =   1155
   End
End
Attribute VB_Name = "frmCierreTransXEntregar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private mbooProcesando As Boolean
Private mbooCancelado As Boolean
Private mEmpOrigen As Empresa
Private Const MSG_OK As String = "OK"
Private mObjCond As RepCondicion

Private WithEvents mGrupo As grupo
Attribute mGrupo.VB_VarHelpID = -1
Const COL_CODTRANS = 2
Const COL_NUMTRANS = 3
Const COL_CODCC = 4
Const COL_CODCLI = 5
Const COL_CODITEM = 7
Const COL_CODBODEGA = 9
Const COL_SALDO = 13
Const COL_SALDOITEM = 12

Public Sub Inicio(ByVal tag As String)
    On Error GoTo ErrTrap
    Set mObjCond = New RepCondicion
    Select Case tag
        Case "Items"
            Me.Caption = "Pendientes x Entregar Items"
        Case "Familias"
            Me.Caption = "Pendientes x Entregar Familias"
            ConfigCols
        Case "ItemsHormi"
            Me.Caption = "Pendientes x Entregar Items"
    End Select
    Me.tag = tag
    Me.Show
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub



Private Sub cmdBuscar_Click()
    LeerArchivo (dlg1.filename)
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

Private Sub cmdExplorar_Click()
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
        txtOrigen.Text = dlg1.filename
        LeerArchivo (dlg1.filename)
    End With
    Exit Sub
ErrTrap:
    If Err.Number <> 32755 Then DispErr
    Exit Sub
End Sub

Private Sub cmdPasos_Click()
    Dim r As Boolean
    
    '8. Pasar trans. existentes con la fecha posterior a la fecha de corte
            If Not frmB_PendxFamilia.InicioPendientesxFamilia(mObjCond, Me.tag) Then
                grd.SetFocus
                Exit Sub
            End If
    
    Select Case Me.tag
        Case "Items"
            r = CopiaTrans
        Case "Familias"
            r = CopiaTransFamilias
        Case "ItemsHormi"
            r = CopiaTransHormi
    
    End Select
    
        
    If r Then
        'If Index < cmdPasos.Count - 1 Then cmdPasos(Index + 1).SetFocus
        lblPasos.BackColor = vbBlue
        lblPasos.ForeColor = vbYellow
    End If
End Sub

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
        j = gc.AddCTLibroDetalle
        Set ctd = gc.CTLibroDetalle(j)
        ctd.codcuenta = gc.Empresa.GNOpcion.CodCuentaResultado
        ctd.Haber = gc.DebeTotal - gc.HaberTotal
        ctd.Descripcion = "Resultado del ejercicio"
        ctd.Orden = j
    End If
            
    gc.Grabar False, False
End Sub

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

'8. Pasar trans. existentes con la fecha posterior a la fecha de corte
Private Function CopiaTransFamilias() As Boolean
    Dim i As Long, codt As String, numt As Long
    Dim empDestino As Empresa, sql As String, Num As Long, ultimaFila As Long, j As Long
    'Verifica  errores  en la base de Origen
    
    Set empDestino = AbrirDestino
    If empDestino.NombreDB = mEmpOrigen.NombreDB Then
        MsgBox "La empresa origen y destino son las mismas" & Chr(13) & _
               "debera  seleccionar  una empresa de  destino diferente", vbExclamation
        Exit Function
    End If
    
    If grd.FixedRows + 1 = grd.Rows Then CargaPendientexFamilia 'Carga  trans solo  si la grlla esta vacia
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

        i = .FixedRows
        While i <= .Rows - 1
        'For i = .FixedRows To .Rows - 1
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
                numt = .TextMatrix(i, .ColIndex("# Trans"))
                'If codt = "FC" And numt = 2524 Then Stop
                MensajeStatus "Copiando la transacción " & codt & numt & _
                            "     " & i & " de " & .Rows - .FixedRows & _
                            " (" & Format(i * 100 / (.Rows - .FixedRows), "0") & "%)", vbHourglass
                
                'Si aún no está importado bien, importa la fila
                If grd.TextMatrix(i, .Cols - 1) <> MSG_OK Then
                    If grd.ValueMatrix(i, .Cols - 2) > 0 Then
                    ultimaFila = i
                    If ImportarTransSub(codt, numt, empDestino, i) Then
                        For j = ultimaFila To i
                            .TextMatrix(j, .ColIndex("Resultado")) = MSG_OK
                        Next j
                    Else
                        For j = ultimaFila To i
                            .TextMatrix(i, .ColIndex("Resultado")) = "Error"
                        Next j
                    End If
                    Else
                        .TextMatrix(i, .ColIndex("Resultado")) = "Cantidad Negativa"
                    End If
               End If
               i = i + 1
        Wend
       'Next i
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
    CopiaTransFamilias = True
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



Private Function ImportarTransSub( _
                ByVal codt As String, _
                ByVal numt As Long, ByRef empDestino As Empresa, _
                ByRef fila As Long) As Boolean
    Dim gnDest As GNComprobante, s As String, Estado As Byte, gnOri As GNComprobante
    Dim ivk As IVKardex
    Dim i As Long, contItems As Long, j As Long
    Dim BandOtro As Boolean, BandEncontro As Boolean
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
    
'    If gnOri.CodCentro <> gnDest.CodCentro Then
'        gnDest.CodCentro = grd.TextMatrix(Fila, grd.ColIndex("Cod. CC"))
'        gnDest.CodClienteRef = grd.TextMatrix(Fila, grd.ColIndex("Cod cliente"))
        
'    End If
    'elimina todo el ivkardex para cargar solo lo pendiente
    gnDest.BorrarIVKardex
    'verifica y corrige cantidades
    If Len(grd.TextMatrix(fila, COL_CODCLI)) > 0 Then
        gnDest.CodClienteRef = grd.TextMatrix(fila, COL_CODCLI)
    End If
    If Len(grd.TextMatrix(fila, COL_CODCC)) > 0 Then
        gnDest.CodCentro = grd.TextMatrix(fila, COL_CODCC)
    End If
    i = 1
    contItems = 0
    While i <= gnOri.CountIVKardex
        BandOtro = True
        BandEncontro = False
        
        While BandOtro And i <= gnOri.CountIVKardex
            If gnOri.IVKardex(i).CodInventario = grd.TextMatrix(fila, COL_CODITEM) Then
                contItems = contItems + 1
                gnDest.AddIVKardex
                gnDest.IVKardex(contItems).CodInventario = grd.TextMatrix(fila, COL_CODITEM)
                gnDest.IVKardex(contItems).CodBodega = grd.TextMatrix(fila, COL_CODBODEGA)
                If gnDest.GNTrans.IVTipoTrans = "E" Then
                    gnDest.IVKardex(contItems).cantidad = grd.ValueMatrix(fila, COL_SALDOITEM + 1) * -1
                Else
                    gnDest.IVKardex(contItems).cantidad = grd.ValueMatrix(fila, COL_SALDOITEM + 1)
                End If
                gnDest.IVKardex(contItems).Orden = contItems
                gnDest.IVKardex(contItems).CostoTotal = gnOri.IVKardex(i).CostoTotal
                gnDest.IVKardex(contItems).CostoRealTotal = gnOri.IVKardex(i).CostoRealTotal
                gnDest.IVKardex(contItems).PrecioTotal = gnOri.IVKardex(i).PrecioTotal
                gnDest.IVKardex(contItems).PrecioRealTotal = gnOri.IVKardex(i).PrecioRealTotal
                gnDest.IVKardex(contItems).Descuento = gnOri.IVKardex(i).Descuento
                gnDest.IVKardex(contItems).IVA = gnOri.IVKardex(i).IVA
                gnDest.IVKardex(contItems).Nota = gnOri.IVKardex(i).Nota
                gnDest.IVKardex(contItems).NumeroPrecio = gnOri.IVKardex(i).NumeroPrecio
                gnDest.IVKardex(contItems).ValorRecargoItem = gnOri.IVKardex(i).ValorRecargoItem
                gnDest.IVKardex(contItems).TiempoEntrega = gnOri.IVKardex(i).TiempoEntrega
                gnDest.IVKardex(contItems).IdICE = gnOri.IVKardex(i).IdICE
                BandEncontro = True
            End If
            
            'fila = fila + 1
            i = i + 1
            If BandEncontro Then
                If Not fila = grd.Rows - 1 Then
                    If grd.TextMatrix(fila + 1, COL_CODTRANS) = codt And grd.TextMatrix(fila + 1, COL_NUMTRANS) = numt Then
                        fila = fila + 1
                        BandOtro = True
                        i = 1
                        BandEncontro = False
'                        Fila = 5
                    Else
                        BandOtro = False
                        i = gnOri.CountIVKardex + 1
                    End If
                Else
                    BandOtro = False
                End If
            End If
            
        Wend
    Wend

    
    gnDest.FechaTrans = DateAdd("d", 1, dtpFechaCorte.value)
    gnDest.Descripcion = "PENDIENTE DESDE " & gnOri.FechaTrans
    
    
    
    For j = 1 To gnDest.CountPCKardex
        gnDest.RemovePCKardex (1)
    Next j
        
    For j = 1 To gnDest.CountIVKardexRecargo
        gnDest.RemoveIVKardexRecargo (1)
    Next j
        
        
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
    If MsgBox(Err.Description & vbCr & vbCr & _
                "Desea continuar con siguiente transacción?", _
                vbQuestion + vbYesNo) <> vbYes Then
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



Private Sub cmdPasosArchivo_Click()
    CopiaTransFamiliasArchivo
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
    dtpFechaCorte.value = mEmpOrigen.GNOpcion.FechaFinal
    
    'Visualiza codigo de empresa origen (= Empresa actual)
    lblOrigen.Caption = mEmpOrigen.CodEmpresa
    lblOrigenBD.Caption = mEmpOrigen.NombreDB
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = mbooProcesando
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    prg1.Width = Me.ScaleWidth - (prg1.Left * 2)
    grd.Move 0, lblPasos.Top + lblPasos.Height + 400, Me.ScaleWidth, Me.ScaleHeight - lblPasos.Top + lblPasos.Height + 100 - pic1.Height - 1200
'    prg1.Width = Me.ScaleWidth - (prg1.Left * 2)
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
    If Len(txtDestino.Text) > 0 Then
        txtDestino_Validate Cancel
        If Cancel Then txtDestino.SetFocus
    End If
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

Public Sub MiGetRowsRep(ByVal rs As Recordset, grd As VSFlexGrid)
    grd.LoadArray MiGetRows(rs)
'    ConfigTipoDatoCol grd, rs
End Sub

Public Sub VerificaExistenciaTablaOrigen(i As Integer)
    Dim rs As Recordset
    Dim sql As String
    'verifica  si la  tabla no esta  creada
    sql = "SELECT * FROM sysobjects WHERE NAME =  'tmp" & i & "'"
    Set rs = mEmpOrigen.OpenRecordset(sql)
    If Not (rs.EOF And rs.BOF) Then
        'elimina la tabla
        mEmpOrigen.EjecutarSQL "drop table Tmp" & i, 0
    End If
End Sub


Private Function CopiaTrans() As Boolean
    Dim i As Long, codt As String, numt As Long
    Dim empDestino As Empresa, sql As String, Num As Long, ultimaFila As Long, j As Long
    'Verifica  errores  en la base de Origen
    
    Set empDestino = AbrirDestino
'    If empDestino.NombreDB = mEmpOrigen.NombreDB Then
'        MsgBox "La empresa origen y destino son las mismas" & Chr(13) & _
'               "debera  seleccionar  una empresa de  destino diferente", vbExclamation
'        Exit Function
'    End If
    
    CargaPendiente
    
    If grd.FixedRows = grd.Rows Then CargaPendiente 'Carga  trans solo  si la grlla esta vacia
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

        i = .FixedRows
        While i <= .Rows - 1
        'For i = .FixedRows To .Rows - 1
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
                numt = .TextMatrix(i, .ColIndex("# Trans"))
                'If codt = "FC" And numt = 2524 Then Stop
                MensajeStatus "Copiando la transacción " & codt & numt & _
                            "     " & i & " de " & .Rows - .FixedRows & _
                            " (" & Format(i * 100 / (.Rows - .FixedRows), "0") & "%)", vbHourglass
                
                'Si aún no está importado bien, importa la fila
                If grd.TextMatrix(i, .Cols - 1) <> MSG_OK Then
                    ultimaFila = i
                    If grd.ValueMatrix(i, .Cols - 2) > 0 Then
                        If ImportarTransSubItem(codt, numt, empDestino, i) Then
                            For j = ultimaFila To i
                                .TextMatrix(j, .ColIndex("Resultado")) = MSG_OK
                            Next j
                        Else
                            For j = ultimaFila To i
                                .TextMatrix(i, .ColIndex("Resultado")) = "Error"
                            Next j
                        End If
                    Else
                        .TextMatrix(i, .ColIndex("Resultado")) = "Cantidad Negativa"
                    End If
               End If
               i = i + 1
        Wend
       'Next i
    End With
  
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


Private Sub CargaPendiente()
    Dim sql As String, cond As String, trans As String
    Dim CondTransSal As String, CondtransFac As String, NumReg As Long
    Dim CondTransDev As String
    Dim Fcorte As Date, rs As Recordset, v As Variant
    On Error GoTo ErrTrap
    With mObjCond
        Fcorte = dtpFechaCorte.value    'Fecha de corte
        'Parametros devueltos: .TipoTrans = lista de transacciones de  salida
        '.SQLItem = lista de transacciones de Factura
        cond = " AND (GNComprobante.FechaTrans < " & _
                FechaYMD(Fcorte, gobjMain.EmpresaActual.TipoDB) & ") "
        If Len(.tipoTrans) > 0 Then
           CondTransSal = cond & " AND  GNComprobante.CodTrans IN (" & PreparaCadena(.tipoTrans) & ") "
        End If
        If Len(.CodTrans) > 0 Then
           CondTransDev = cond & " AND  GNComprobante.CodTrans IN (" & PreparaCadena(.CodTrans) & ") "
        End If
        
        
        CondtransFac = cond & " AND  GNComprobante.CodTrans IN (" & PreparaCadena(.SQLItem) & ") "
        trans = " CodTrans,NumTrans "
        'Elimina tablas temporales
        VerificaExistenciaTablaOrigen 1
        VerificaExistenciaTablaOrigen 11
        VerificaExistenciaTablaOrigen 2
        VerificaExistenciaTablaOrigen 22
        VerificaExistenciaTablaOrigen 3
        VerificaExistenciaTablaOrigen 4
        sql = "SELECT GNComprobante.idTransFuente, GNComprobante.FechaTrans, CodTrans, NumTrans , PCProvCli.CodProvcli, PCProvCli.Nombre, " & _
              "GNComprobante.NumDocRef,IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, " & _
              "sum(IVKardex.Cantidad) As Cantidad, GNComprobante.IdCentro, GNCentroCosto.CodCentro "
              
        sql = sql & "INTO tmp1  FROM IvInventario" & " INNER JOIN ((GNComprobante LEFT JOIN PCProvCli " & _
             "ON GNComprobante.IdClienteRef " & "= PCProvCli.IdProvCli) LEFT JOIN GNCentroCosto  ON GNComprobante.IdCentro=GNCentroCosto.IdCentro " & _
             "INNER JOIN IVKardex  ON GNComprobante.TransID = IVKardex.TransID) " & _
             "ON IVInventario.IdInventario = IVKardex.IdInventario "

        sql = sql & " WHERE GNComprobante.Estado<>3 and IvInventario.bandservicio=0 "
        sql = sql & CondTransSal
        sql = sql & "GROUP BY GNComprobante.FechaTrans, CodTrans,NumTrans " & _
             ", PCProvCli.CodProvcli, PCProvCli.Nombre, " & _
             "GNComprobante.NumDocRef, " & _
             "IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, " & _
             " gnComprobante.IdCentro, GNCentroCosto.CodCentro,GNComprobante.idTransFuente "
        sql = sql & " ORDER BY  PCProvCli.Nombre, GNComprobante.FechaTrans, " & trans
        grd.Redraw = False
        mEmpOrigen.EjecutarSQL sql, NumReg 'Priemer SQL Transacciones de consumo
        '----------------------------------------------------------------------------------
        sql = "SELECT IVKardex.ID, GNComprobante.TransId, GNComprobante.FechaTrans, CodTrans, NumTrans, PCProvCli.CodProvcli , PCProvCli.Nombre, " & _
              " GNComprobante.NumDocRef,IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion,  IVBodega.CodBodega, " & _
              " sum(IVKardex.Cantidad) As Cantidad, IvInventario.Unidad, GNComprobante.IdCentro, GNCentroCosto.CodCentro  "
              
        sql = sql & "INTO tmp2  FROM IvInventario" & " INNER JOIN ((GNComprobante LEFT JOIN PCProvCli " & _
             "ON GNComprobante.IdClienteRef " & "= PCProvCli.IdProvCli) LEFT JOIN GNCentroCosto  ON GNComprobante.IdCentro=GNCentroCosto.IdCentro " & _
             "INNER JOIN (IVKardex inner join IVBodega on IvKardex.Idbodega = IvBodega.Idbodega)  ON GNComprobante.TransID = IVKardex.TransID) " & _
             "ON IVInventario.IdInventario = IVKardex.IdInventario "

        sql = sql & " WHERE GNComprobante.Estado<>3 and IvInventario.bandservicio=0 "  '
        sql = sql & " AND IVInventario.Tipo <> 2 "
        sql = sql & CondtransFac
        sql = sql & "GROUP BY GNComprobante.TransId, IVKardex.ID, GNComprobante.FechaTrans, CodTrans,NumTrans" & _
             ", PCProvCli.CodProvcli, PCProvCli.Nombre, " & _
             "GNComprobante.NumDocRef, IVBodega.CodBodega, " & _
             "IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, " & _
             " IVinventario.Unidad, GNComprobante.IdCentro, GNCentroCosto.CodCentro "
        sql = sql & " ORDER BY  PCProvCli.Nombre, GNComprobante.FechaTrans, CodTrans, NumTrans"
        mEmpOrigen.EjecutarSQL sql, NumReg 'Segundo SQL Transacciones de Facturacion
        
        
        sql = "SELECT GNComprobante.FechaTrans, CodTrans, NumTrans , PCProvCli.CodProvcli, PCProvCli.Nombre, " & _
              "GNComprobante.NumDocRef,IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, " & _
              "sum(IVKardex.Cantidad) As Cantidad, GNComprobante.IdCentro, GNCentroCosto.CodCentro, GNComprobante.idtransfuente "
              
        sql = sql & "INTO tmp11  FROM IvInventario" & " INNER JOIN ((GNComprobante LEFT JOIN PCProvCli " & _
             "ON GNComprobante.IdClienteRef " & "= PCProvCli.IdProvCli) LEFT JOIN GNCentroCosto  ON GNComprobante.IdCentro=GNCentroCosto.IdCentro " & _
             "INNER JOIN IVKardex  ON GNComprobante.TransID = IVKardex.TransID) " & _
             "ON IVInventario.IdInventario = IVKardex.IdInventario "

        sql = sql & " WHERE GNComprobante.Estado<>3 and IvInventario.bandservicio=0 "
        sql = sql & CondTransDev
        sql = sql & "GROUP BY GNComprobante.FechaTrans, CodTrans,NumTrans " & _
             ", PCProvCli.CodProvcli, PCProvCli.Nombre, " & _
             "GNComprobante.NumDocRef, " & _
             "IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, " & _
             " gnComprobante.IdCentro, GNCentroCosto.CodCentro, GNComprobante.idtransfuente "
        sql = sql & " ORDER BY  PCProvCli.Nombre, GNComprobante.FechaTrans, " & trans
        grd.Redraw = False
        mEmpOrigen.EjecutarSQL sql, NumReg 'Tercer SQL Transacciones de consumo
        
        
        '--------------------------------------------------------------------------------------------------------
        '--Transforma tmp1 en Familia  totalizar por Familias  y por Centro de costo
        
        'Inner JOIN IvMateria Ivm ON tmp1.IdInventario = Ivm.IdInventario
        sql = "Select Tmp1.idTransFuente,  sum(tmp1.Cantidad) As Cantidad, tmp1.IDCentro, tmp1.IDInventario " & _
              "INTO tmp3 From tmp1  " & _
              "Group By Tmp1.idTransFuente, tmp1.IdCentro, tmp1.IdInventario "
        mEmpOrigen.EjecutarSQL sql, NumReg 'tercer SQL Transforma PTerminado a Familias
        'Unir 2 tablas para comparar
        
        'saca totales de facturado menos devoluciones
'        sql = "Select  sum(tmp2.Cantidad) As Cantidad, tmp2.IDCentro, tmp2.IDInventario " & _
'              "INTO tmp22 From tmp2  " & _
'              "Group By tmp2.IdCentro, tmp2.IdInventario "
'        mEmpOrigen.EjecutarSQL sql, NumReg 'tercer SQL Transforma PTerminado a Familias
        
        
        
        cond = "Where  "
        If mObjCond.Bandera Then
           cond = cond & "  ((isnull(tmp2.Cantidad,0) - isnull(tmp3.Cantidad,0)+isnull(tmp11.cantidad,0))) <> 0 "
        End If
        
         sql = "Select tmp2.ID, tmp2.FechaTrans, tmp2.codTrans, tmp2.NumTrans, tmp2.CodProvcli, tmp2.Nombre, tmp2.NumDocRef, " & _
               " tmp2.CodInventario, tmp2.Descripcion,  tmp2.CodBodega, isnull(tmp2.Cantidad,0)*-1 as Cantidad, " & _
               " isnull(tmp3.Cantidad,0) *-1 - isnull(tmp11.Cantidad,0) *-1 As    Despachado, (isnull(tmp2.Cantidad,0) - isnull(tmp3.Cantidad,0) +isnull(tmp11.cantidad,0))*-1 As Pendiente,  tmp2.Unidad,  " & _
               " tmp2.IdCentro, tmp2.CodCentro into tmp4 " & _
               " From tmp2   left join tmp11 on tmp2.transid=tmp11.idtransfuente and tmp2.idinventario=tmp11.idinventario left Join  tmp3 ON  tmp2.IdCentro = tmp3.IdCentro  and tmp2.transid=tmp3.idtransfuente and tmp2.idinventario=tmp3.idinventario " & _
               cond
        mEmpOrigen.EjecutarSQL sql, NumReg 'tercer SQL Transforma PTerminado a Familias
                
'Explicacion del select despues deel UNION
'Preimero quitamos los que ya estan  calculados
'Luego quitamos todos los que tiene saldo cero

        sql = "Select tmp4.FechaTrans, tmp4.codTrans, tmp4.NumTrans, tmp4.CodCentro, tmp4.CodProvcli, tmp4.Nombre,  " & _
               "tmp4.CodInventario, tmp4.Descripcion, tmp4.CodBodega, tmp4.Cantidad, " & _
               "tmp4.Despachado, tmp4.Pendiente From tmp4  " & _
               " Where tmp4.Pendiente>0 "
'               "Union " & _
'               "Select tmp2.FechaTrans, tmp2.CodTrans, tmp2.NumTrans, tmp2.CodCentro, tmp2.CodProvcli, tmp2.Nombre,  tmp2.CodInventario, " & _
'               "tmp2.Descripcion,  tmp2.Cantidad*-1, 0 As Despachado, (tmp2.Cantidad)*-1 As Pendiente " & _
'               "From tmp2 " & _
'               "Where tmp2.Id  NOT IN ( " & _
'               " Select tmp4.Id from tmp4  ) " & _
'               "AND " & _
'               "tmp2.ID NOT IN( " & _
'               "Select tmp2.ID " & _
'               "From tmp2 Left Join  tmp3 ON  tmp2.IdCentro = tmp3.IdCentro " & _
'               "Where   tmp2.IdInventario = tmp3.IDInventario AND (tmp2.Cantidad - tmp3.Cantidad)*-1 = 0) " & _
'               " and (tmp2.Cantidad)*-1 > 0 "
            sql = sql & "Order By FechaTrans "
            'MiGetRowsRep mEmpOrigen.OpenRecordset(sql), grd
            Set rs = mEmpOrigen.OpenRecordset(sql)
            
    If Not rs.EOF Then
        v = MiGetRows(rs)
        With grd
            .Redraw = flexRDNone
            .LoadArray v            'Carga a la grilla
            .FormatString = "^#|<Fecha|<CodTrans|<# Trans|Cod. CC|<Cod Cliente|<Cliente|<Cod Item|<Descripcion Item|<Cod Bodega|>Facturado|>Entregado|>Saldo|<Resultado"
            GNPoneNumFila grd, False
            AsignarTituloAColKey grd            'Para usar ColIndex
            AjustarAutoSize grd, 0, -1, 4000     'Ajusta automáticamente ancho de cols.
            .ColWidth(.ColIndex("Resultado")) = 800
            .ColWidth(.ColIndex("Cliente")) = 2500
            'Tipo de datos
            .ColDataType(.ColIndex("Resultado")) = flexDTString
            .ColHidden(4) = True
            .ColHidden(5) = True
            .Redraw = flexRDDirect
        End With
End If
            
        'Elimina las tablas temporales
'        VerificaExistenciaTablaOrigen 1
'        VerificaExistenciaTablaOrigen 2
'        VerificaExistenciaTablaOrigen 3
'        VerificaExistenciaTablaOrigen 4
    End With
    Exit Sub
ErrTrap:
'    VerificaExistenciaTablaOrigen 1
'    VerificaExistenciaTablaOrigen 2
'    VerificaExistenciaTablaOrigen 3
'    VerificaExistenciaTablaOrigen 4
    grd.Redraw = True
    DispErr
End Sub


Private Sub CargaPendienteConBodegas()
    Dim sql As String, cond As String, trans As String
    Dim CondTransSal As String, CondtransFac As String, NumReg As Long
    Dim Fcorte As Date, rs As Recordset, v As Variant
    On Error GoTo ErrTrap
    With mObjCond
        Fcorte = dtpFechaCorte.value    'Fecha de corte
        'Parametros devueltos: .TipoTrans = lista de transacciones de  salida
        '.SQLItem = lista de transacciones de Factura
        cond = " AND (GNComprobante.FechaTrans < " & _
                FechaYMD(Fcorte, gobjMain.EmpresaActual.TipoDB) & ") "
        If Len(.tipoTrans) > 0 Then
           CondTransSal = cond & " AND  GNComprobante.CodTrans IN (" & PreparaCadena(.tipoTrans) & ") "
        End If
        CondtransFac = cond & " AND  GNComprobante.CodTrans IN (" & PreparaCadena(.SQLItem) & ") "
        trans = " CodTrans,NumTrans "
        'Elimina tablas temporales
        VerificaExistenciaTablaOrigen 1
        VerificaExistenciaTablaOrigen 2
        VerificaExistenciaTablaOrigen 3
        VerificaExistenciaTablaOrigen 4
        sql = "SELECT GNComprobante.FechaTrans, CodTrans, NumTrans , PCProvCli.CodProvcli, PCProvCli.Nombre, " & _
              "GNComprobante.NumDocRef,IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, IVBodega.CodBodega, " & _
              "sum(IVKardex.Cantidad) As Cantidad, GNComprobante.IdCentro, GNCentroCosto.CodCentro "
        
              
        sql = sql & "INTO tmp1  FROM IvInventario" & " INNER JOIN ((GNComprobante LEFT JOIN PCProvCli " & _
             "ON GNComprobante.IdClienteRef " & "= PCProvCli.IdProvCli) LEFT JOIN GNCentroCosto  ON GNComprobante.IdCentro=GNCentroCosto.IdCentro " & _
             "INNER JOIN (IVKardex inner join IVBodega on IvKardex.Idbodega = IvBodega.Idbodega) ON GNComprobante.TransID = IVKardex.TransID) " & _
             "ON IVInventario.IdInventario = IVKardex.IdInventario "

        sql = sql & " WHERE GNComprobante.Estado<>3"
        sql = sql & CondTransSal
        sql = sql & "GROUP BY GNComprobante.FechaTrans, CodTrans,NumTrans " & _
             ", PCProvCli.CodProvcli, PCProvCli.Nombre, " & _
             "GNComprobante.NumDocRef, " & _
             "IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, " & _
             "IVBodega.CodBodega, gnComprobante.IdCentro, GNCentroCosto.CodCentro "
        sql = sql & " ORDER BY  PCProvCli.Nombre, GNComprobante.FechaTrans, " & trans
        grd.Redraw = False
        mEmpOrigen.EjecutarSQL sql, NumReg 'Priemer SQL Transacciones de salida PTerminado
        '----------------------------------------------------------------------------------
        sql = "SELECT IVKardex.ID, GNComprobante.FechaTrans, CodTrans, NumTrans, PCProvCli.CodProvcli , PCProvCli.Nombre, " & _
              "GNComprobante.NumDocRef,IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, IVBodega.CodBodega, " & _
              "sum(IVKardex.Cantidad) As Cantidad, IvInventario.Unidad, GNComprobante.IdCentro, GNCentroCosto.CodCentro  "
              
        sql = sql & "INTO tmp2  FROM IvInventario" & " INNER JOIN ((GNComprobante LEFT JOIN PCProvCli " & _
             "ON GNComprobante.IdClienteRef " & "= PCProvCli.IdProvCli) LEFT JOIN GNCentroCosto  ON GNComprobante.IdCentro=GNCentroCosto.IdCentro " & _
             "INNER JOIN (IVKardex inner join IVBodega on IvKardex.Idbodega = IvBodega.Idbodega) ON GNComprobante.TransID = IVKardex.TransID) " & _
             "ON IVInventario.IdInventario = IVKardex.IdInventario "

        sql = sql & " WHERE GNComprobante.Estado<>3"  '
        sql = sql & " AND IVInventario.Tipo <> 2 "
        sql = sql & CondtransFac
        sql = sql & "GROUP BY IVKardex.ID, GNComprobante.FechaTrans, CodTrans,NumTrans" & _
             ", PCProvCli.CodProvcli, PCProvCli.Nombre, " & _
             "GNComprobante.NumDocRef, " & _
             "IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, " & _
             "IVBodega.CodBodega, IVinventario.Unidad, GNComprobante.IdCentro, GNCentroCosto.CodCentro "
        sql = sql & " ORDER BY  PCProvCli.Nombre, GNComprobante.FechaTrans, CodTrans, NumTrans"
        mEmpOrigen.EjecutarSQL sql, NumReg 'Segundo SQL Transacciones de Facturacion familias
        '--------------------------------------------------------------------------------------------------------
        '--Transforma tmp1 en Familia  totalizar por Familias  y por Centro de costo
        
        'Inner JOIN IvMateria Ivm ON tmp1.IdInventario = Ivm.IdInventario
        sql = "Select tmp1.CodBodega, sum(tmp1.Cantidad) As Cantidad, tmp1.IDCentro, tmp1.IDInventario " & _
              "INTO tmp3 From tmp1  " & _
              "Group By tmp1.CodBodega,tmp1.IdCentro, tmp1.IdInventario "
        mEmpOrigen.EjecutarSQL sql, NumReg 'tercer SQL Transforma PTerminado a Familias
        'Unir 2 tablas para comparar
        
        
        
        cond = "Where tmp2.IdInventario = tmp3.IDInventario "
        If mObjCond.Bandera Then
           cond = cond & " AND (tmp2.Cantidad - tmp3.Cantidad)*-1 <> 0 "
        End If
        
         sql = "Select tmp2.ID, tmp2.FechaTrans, tmp2.codTrans, tmp2.NumTrans, tmp2.CodProvcli, tmp2.Nombre, tmp2.NumDocRef, " & _
               "tmp2.CodInventario, tmp2.Descripcion, tmp2.CodBodega, tmp2.Cantidad*-1 as Cantidad, " & _
               "tmp3.Cantidad *-1 As    Despachado, (tmp2.Cantidad - tmp3.Cantidad)*-1 As Pendiente, tmp2.Unidad,  " & _
               "tmp2.IdCentro, tmp2.CodCentro into tmp4 " & _
               "From tmp2 Left Join  tmp3 ON  tmp2.IdCentro = tmp3.IdCentro AND TMP2.CodBodega=TMP3.CodBodega " & _
               cond
        mEmpOrigen.EjecutarSQL sql, NumReg 'tercer SQL Transforma PTerminado a Familias
                
'Explicacion del select despues deel UNION
'Preimero quitamos los que ya estan  calculados
'Luego quitamos todos los que tiene saldo cero

        sql = "Select tmp4.FechaTrans, tmp4.codTrans, tmp4.NumTrans, tmp4.CodCentro, tmp4.CodProvcli, tmp4.Nombre,  " & _
               "tmp4.CodInventario, tmp4.Descripcion, tmp4.CodBodega, tmp4.Cantidad, " & _
               "tmp4.Despachado, tmp4.Pendiente From tmp4  " & _
               " Where tmp4.Pendiente>0 " & _
               "Union " & _
               "Select tmp2.FechaTrans, tmp2.CodTrans, tmp2.NumTrans, tmp2.CodCentro, tmp2.CodProvcli, tmp2.Nombre,  tmp2.CodInventario, " & _
               "tmp2.Descripcion, tmp2.CodBodega, tmp2.Cantidad*-1, 0 As Despachado, (tmp2.Cantidad)*-1 As Pendiente " & _
               "From tmp2 " & _
               "Where tmp2.Id  NOT IN ( " & _
               " Select tmp4.Id from tmp4  ) " & _
               "AND " & _
               "tmp2.ID NOT IN( " & _
               "Select tmp2.ID " & _
               "From tmp2 Left Join  tmp3 ON  tmp2.IdCentro = tmp3.IdCentro " & _
               "Where   tmp2.IdInventario = tmp3.IDInventario AND (tmp2.Cantidad - tmp3.Cantidad)*-1 = 0) " & _
               " and (tmp2.Cantidad)*-1 > 0 "
            sql = sql & "Order By FechaTrans "
            'MiGetRowsRep mEmpOrigen.OpenRecordset(sql), grd
            Set rs = mEmpOrigen.OpenRecordset(sql)
            
    If Not rs.EOF Then
        v = MiGetRows(rs)
        With grd
            .Redraw = flexRDNone
            .LoadArray v            'Carga a la grilla
            .FormatString = "^#|<Fecha|<CodTrans|<# Trans|Cod. CC|<Cod Cliente|<Cliente|<Cod Item|<Descripcion Item|<Cod Bodega|>Facturado|>Entregado|>Saldo|<Resultado"
            GNPoneNumFila grd, False
            AsignarTituloAColKey grd            'Para usar ColIndex
            AjustarAutoSize grd, 0, -1, 4000     'Ajusta automáticamente ancho de cols.
            .ColWidth(.ColIndex("Resultado")) = 800
            .ColWidth(.ColIndex("Cliente")) = 2500
            'Tipo de datos
            .ColDataType(.ColIndex("Resultado")) = flexDTString
            .ColHidden(4) = True
            .ColHidden(5) = True
            .Redraw = flexRDDirect
        End With
End If
            
        'Elimina las tablas temporales
        VerificaExistenciaTablaOrigen 1
        VerificaExistenciaTablaOrigen 2
        VerificaExistenciaTablaOrigen 3
        VerificaExistenciaTablaOrigen 4
    End With
    Exit Sub
ErrTrap:
    VerificaExistenciaTablaOrigen 1
    VerificaExistenciaTablaOrigen 2
    VerificaExistenciaTablaOrigen 3
    VerificaExistenciaTablaOrigen 4
    grd.Redraw = True
    DispErr
End Sub

'''''Private Sub CargaPendientexFamilia()
'''''    Dim sql As String, Cond As String, trans As String
'''''    Dim CondTransSal As String, CondtransFac As String, numReg As Long
'''''    Dim Fcorte As Date, rs As Recordset, v As Variant
'''''    On Error GoTo ErrTrap
'''''    With mObjCond
'''''        Fcorte = dtpFechaCorte.value    'Fecha de corte
'''''        Parametros devueltos: .TipoTrans = lista de transacciones de  salida
'''''        .SQLItem = lista de transacciones de Factura
'''''        Cond = " AND (GNComprobante.FechaTrans < " & _
'''''                FechaYMD(Fcorte, gobjMain.EmpresaActual.TipoDB) & ") "
'''''        If Len(.TipoTrans) > 0 Then
'''''           CondTransSal = Cond & " AND  GNComprobante.CodTrans IN (" & PreparaCadena(.TipoTrans) & ") "
'''''        End If
'''''        CondtransFac = Cond & " AND  GNComprobante.CodTrans IN (" & PreparaCadena(.SQLItem) & ") "
'''''        trans = " CodTrans,NumTrans "
'''''        Elimina tablas temporales
'''''        VerificaExistenciaTablaOrigen 1
'''''        VerificaExistenciaTablaOrigen 2
'''''        VerificaExistenciaTablaOrigen 3
'''''        VerificaExistenciaTablaOrigen 4
'''''        sql = "SELECT GNComprobante.FechaTrans, CodTrans, NumTrans , PCProvCli.CodProvcli, PCProvCli.Nombre, " & _
'''''              "GNComprobante.NumDocRef,IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, IVBodega.CodBodega, " & _
'''''              "sum(IVKardex.Cantidad) As Cantidad, GNComprobante.IdCentro, GNCentroCosto.CodCentro "
'''''
'''''        sql = sql & "INTO tmp1  FROM IvInventario" & " INNER JOIN ((GNComprobante LEFT JOIN PCProvCli " & _
'''''             "ON GNComprobante.IdClienteRef " & "= PCProvCli.IdProvCli) LEFT JOIN GNCentroCosto  ON GNComprobante.IdCentro=GNCentroCosto.IdCentro " & _
'''''             "INNER JOIN (IVKardex inner join IVBodega on IvKardex.Idbodega = IvBodega.Idbodega) ON GNComprobante.TransID = IVKardex.TransID) " & _
'''''             "ON IVInventario.IdInventario = IVKardex.IdInventario "
'''''
'''''        sql = sql & " WHERE GNComprobante.Estado<>3"
'''''        sql = sql & CondTransSal
'''''        sql = sql & "GROUP BY GNComprobante.FechaTrans, CodTrans,NumTrans " & _
'''''             ", PCProvCli.CodProvcli, PCProvCli.Nombre, " & _
'''''             "GNComprobante.NumDocRef, " & _
'''''             "IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, " & _
'''''             "IVBodega.CodBodega, gnComprobante.IdCentro, GNCentroCosto.CodCentro "
'''''        sql = sql & " ORDER BY  PCProvCli.Nombre, GNComprobante.FechaTrans, " & trans
'''''        grd.Redraw = False
'''''        mEmpOrigen.EjecutarSQL sql, numReg 'Priemer SQL Transacciones de salida PTerminado
'''''        ----------------------------------------------------------------------------------
'''''        sql = "SELECT IVKardex.ID, GNComprobante.FechaTrans, CodTrans, NumTrans, PCProvCli.CodProvcli , PCProvCli.Nombre, " & _
'''''              "GNComprobante.NumDocRef,IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, IVBodega.CodBodega, " & _
'''''              "sum(IVKardex.Cantidad) As Cantidad, IvInventario.Unidad, GNComprobante.IdCentro, GNCentroCosto.CodCentro  "
'''''
'''''        sql = sql & "INTO tmp2  FROM IvInventario" & " INNER JOIN ((GNComprobante LEFT JOIN PCProvCli " & _
'''''             "ON GNComprobante.IdClienteRef " & "= PCProvCli.IdProvCli) LEFT JOIN GNCentroCosto  ON GNComprobante.IdCentro=GNCentroCosto.IdCentro " & _
'''''             "INNER JOIN (IVKardex inner join IVBodega on IvKardex.Idbodega = IvBodega.Idbodega) ON GNComprobante.TransID = IVKardex.TransID) " & _
'''''             "ON IVInventario.IdInventario = IVKardex.IdInventario "
'''''
'''''        sql = sql & " WHERE GNComprobante.Estado<>3 AND IVInventario.Tipo <> " & INV_TIPONORMAL 'Solo familias
'''''        sql = sql & CondtransFac
'''''        sql = sql & "GROUP BY IVKardex.ID, GNComprobante.FechaTrans, CodTrans,NumTrans" & _
'''''             ", PCProvCli.CodProvcli, PCProvCli.Nombre, " & _
'''''             "GNComprobante.NumDocRef, " & _
'''''             "IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, " & _
'''''             "IVBodega.CodBodega, IVinventario.Unidad, GNComprobante.IdCentro, GNCentroCosto.CodCentro "
'''''        sql = sql & " ORDER BY  PCProvCli.Nombre, GNComprobante.FechaTrans, CodTrans, NumTrans"
'''''        mEmpOrigen.EjecutarSQL sql, numReg 'Segundo SQL Transacciones de Facturacion familias
'''''        --------------------------------------------------------------------------------------------------------
'''''        --Transforma tmp1 en Familia  totalizar por Familias  y por Centro de costo
'''''
'''''        sql = "Select tmp1.CodBodega, sum(tmp1.Cantidad) As Cantidad, tmp1.IDCentro,IDMateria " & _
'''''              "INTO tmp3 From tmp1 Inner JOIN IvMateria Ivm ON tmp1.IdInventario = Ivm.IdInventario " & _
'''''              "Group By tmp1.CodBodega,tmp1.IdCentro, IdMateria "
'''''        mEmpOrigen.EjecutarSQL sql, numReg 'tercer SQL Transforma PTerminado a Familias
'''''        Unir 2 tablas para comparar
'''''
'''''
'''''        Cond = "Where tmp2.IdInventario = tmp3.IDMateria "
'''''        If mObjCond.Bandera Then
'''''           Cond = Cond & " AND (tmp2.Cantidad - tmp3.Cantidad)*-1 <> 0 "
'''''        End If
'''''
'''''         sql = "Select tmp2.ID, tmp2.FechaTrans, tmp2.codTrans, tmp2.NumTrans, tmp2.CodProvcli, tmp2.Nombre, tmp2.NumDocRef, " & _
'''''               "tmp2.CodInventario, tmp2.Descripcion, tmp2.CodBodega, tmp2.Cantidad*-1 as Cantidad, " & _
'''''               "tmp3.Cantidad *-1 As    Despachado, (tmp2.Cantidad - tmp3.Cantidad)*-1 As Pendiente, tmp2.Unidad,  " & _
'''''               "tmp2.IdCentro, tmp2.CodCentro into tmp4 " & _
'''''               "From tmp2 Left Join  tmp3 ON  tmp2.IdCentro = tmp3.IdCentro AND TMP2.CodBodega=TMP3.CodBodega " & _
'''''               Cond
'''''        mEmpOrigen.EjecutarSQL sql, numReg 'tercer SQL Transforma PTerminado a Familias
'''''
'''''Explicacion del select despues deel UNION
'''''Preimero quitamos los que ya estan  calculados
'''''Luego quitamos todos los que tiene saldo cero
'''''
'''''        sql = "Select tmp4.FechaTrans, tmp4.codTrans, tmp4.NumTrans, tmp4.CodCentro, tmp4.CodProvcli, tmp4.Nombre,  " & _
'''''               "tmp4.CodInventario, tmp4.Descripcion, tmp4.CodBodega, tmp4.Cantidad, " & _
'''''               "tmp4.Despachado, tmp4.Pendiente From tmp4  " & _
'''''               " Where tmp4.Pendiente>0 " & _
'''''               "Union " & _
'''''               "Select tmp2.FechaTrans, tmp2.CodTrans, tmp2.NumTrans, tmp2.CodCentro, tmp2.CodProvcli, tmp2.Nombre,  tmp2.CodInventario, " & _
'''''               "tmp2.Descripcion, tmp2.CodBodega, tmp2.Cantidad*-1, 0 As Despachado, (tmp2.Cantidad)*-1 As Pendiente " & _
'''''               "From tmp2 " & _
'''''               "Where tmp2.Id  NOT IN ( " & _
'''''               " Select tmp4.Id from tmp4  ) " & _
'''''               "AND " & _
'''''               "tmp2.ID NOT IN( " & _
'''''               "Select tmp2.ID " & _
'''''               "From tmp2 Left Join  tmp3 ON  tmp2.IdCentro = tmp3.IdCentro " & _
'''''               "Where   tmp2.IdInventario = tmp3.IDMateria AND (tmp2.Cantidad - tmp3.Cantidad)*-1 = 0) " & _
'''''               " and (tmp2.Cantidad)*-1 > 0 "
'''''            sql = sql & "Order By FechaTrans "
'''''            MiGetRowsRep mEmpOrigen.OpenRecordset(sql), grd
'''''            Set rs = mEmpOrigen.OpenRecordset(sql)
'''''
'''''    If Not rs.EOF Then
'''''        v = MiGetRows(rs)
'''''        With grd
'''''            .Redraw = flexRDNone
'''''            .LoadArray v            'Carga a la grilla
'''''            .FormatString = "^#|<Fecha|<CodTrans|<# Trans|Cod. CC|<Cod Cliente|<Cliente|<Cod Item|<Descripcion Item|<Cod Bodega|>Facturado|>Entregado|>Saldo|<Resultado"
'''''            GNPoneNumFila grd, False
'''''            AsignarTituloAColKey grd            'Para usar ColIndex
'''''            AjustarAutoSize grd, 0, -1, 3000     'Ajusta automáticamente ancho de cols.
'''''            .ColWidth(.ColIndex("Resultado")) = 800
'''''            Tipo de datos
'''''            .ColDataType(.ColIndex("Resultado")) = flexDTString
'''''            .ColHidden(4) = True
'''''            .ColHidden(5) = True
'''''            .Redraw = flexRDDirect
'''''        End With
'''''End If
'''''
'''''        Elimina las tablas temporales
'''''        VerificaExistenciaTablaOrigen 1
'''''        VerificaExistenciaTablaOrigen 2
'''''        VerificaExistenciaTablaOrigen 3
'''''        VerificaExistenciaTablaOrigen 4
'''''    End With
'''''    Exit Sub
'''''ErrTrap:
'''''    VerificaExistenciaTablaOrigen 1
'''''    VerificaExistenciaTablaOrigen 2
'''''    VerificaExistenciaTablaOrigen 3
'''''    VerificaExistenciaTablaOrigen 4
'''''    grd.Redraw = True
'''''    DispErr
'''''End Sub



Private Sub CargaPendientexFamilia()
    Dim sql As String, cond As String, trans As String
    Dim CondTransSal As String, CondtransFac As String, NumReg As Long, codTransDev As String
    Dim Fcorte As Date, rs As Recordset, v As Variant
    On Error GoTo ErrTrap
    With mObjCond
        'Parametros devueltos: .TipoTrans = lista de transacciones de  salida
        '.SQLItem = lista de transacciones de Factura
        Fcorte = dtpFechaCorte.value    'Fecha de corte
        cond = " AND (GNComprobante.FechaTrans <= " & _
                FechaYMD(Fcorte, gobjMain.EmpresaActual.TipoDB) & ") "
                
        If Len(.tipoTrans) > 0 Then
           CondTransSal = cond & " AND  GNComprobante.CodTrans IN (" & PreparaCadena(.tipoTrans) & ") "
        End If
        CondtransFac = cond & " AND  GNComprobante.CodTrans IN (" & PreparaCadena(.SQLItem) & ") "

        If Len(.CodTrans) > 0 Then
           codTransDev = cond & " AND  GNComprobante.CodTrans IN (" & PreparaCadena(.CodTrans) & ") "
        End If


        trans = " CodTrans, NumTrans"
        'Elimina tablas temporales
        
        VerificaExistenciaTablaOrigen 1
        VerificaExistenciaTablaOrigen 11
        VerificaExistenciaTablaOrigen 12
        VerificaExistenciaTablaOrigen 111
        VerificaExistenciaTablaOrigen 122
        
        VerificaExistenciaTablaOrigen 2
        VerificaExistenciaTablaOrigen 3
        VerificaExistenciaTablaOrigen 4
        
        sql = "SELECT GNComprobante.FechaTrans, " & trans & " ,  PCProvCli.Nombre, " & _
              "GNComprobante.NumDocRef,IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, IVBodega.CodBodega, " & _
              "sum(IVKardex.Cantidad) As Cantidad, GNComprobante.IdCentro "

        sql = sql & "INTO tmp11  FROM IvInventario" & " INNER JOIN ((GNComprobante LEFT JOIN PCProvCli " & _
             "ON GNComprobante.IdClienteRef " & "= PCProvCli.IdProvCli) " & _
             "INNER JOIN (IVKardex inner join IVBodega on IvKardex.Idbodega = IvBodega.Idbodega) ON GNComprobante.TransID = IVKardex.TransID) " & _
             "ON IVInventario.IdInventario = IVKardex.IdInventario "

        sql = sql & " WHERE GNComprobante.Estado<>3"
        sql = sql & CondTransSal
        sql = sql & "GROUP BY GNComprobante.FechaTrans, " & trans & _
             ", PCProvCli.Nombre, " & _
             "GNComprobante.NumDocRef, " & _
             "IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, " & _
             "IVBodega.CodBodega, gnComprobante.IdCentro "
        sql = sql & " ORDER BY  PCProvCli.Nombre, GNComprobante.FechaTrans, " & trans
        grd.Redraw = False
        mEmpOrigen.EjecutarSQL sql, NumReg 'Priemer SQL Transacciones de salida PTerminado
        
        '----------------------------------------------------------------------------------
        'devoluciones
        '----------------------------------------------------------------------------------
        
        sql = "SELECT GNComprobante.FechaTrans, " & trans & " ,  PCProvCli.Nombre, " & _
              "GNComprobante.NumDocRef,IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, IVBodega.CodBodega, " & _
              "sum(IVKardex.Cantidad) As Cantidad, GNComprobante.IdCentro "

        sql = sql & "INTO tmp12  FROM IvInventario" & " INNER JOIN ((GNComprobante LEFT JOIN PCProvCli " & _
             "ON GNComprobante.IdClienteRef " & "= PCProvCli.IdProvCli) " & _
             "INNER JOIN (IVKardex inner join IVBodega on IvKardex.Idbodega = IvBodega.Idbodega) ON GNComprobante.TransID = IVKardex.TransID) " & _
             "ON IVInventario.IdInventario = IVKardex.IdInventario "

        sql = sql & " WHERE GNComprobante.Estado<>3"
        sql = sql & codTransDev
        sql = sql & "GROUP BY GNComprobante.FechaTrans, " & trans & _
             ", PCProvCli.Nombre, " & _
             "GNComprobante.NumDocRef, " & _
             "IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, " & _
             "IVBodega.CodBodega, gnComprobante.IdCentro "
        sql = sql & " ORDER BY  PCProvCli.Nombre, GNComprobante.FechaTrans, " & trans
        grd.Redraw = False
        mEmpOrigen.EjecutarSQL sql, NumReg 'Priemer SQL Transacciones de salida PTerminado
        
        
        '----------------------------------------------------------------------------------
        sql = "SELECT IVKardex.ID, GNComprobante.FechaTrans, " & trans & " , PCProvCli.Codprovcli, PCProvCli.Nombre, " & _
              "GNComprobante.NumDocRef,IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, IVBodega.CodBodega, " & _
              "sum(IVKardex.Cantidad) As Cantidad, IvInventario.Unidad, GNComprobante.IdCentro, codcentro  "

        sql = sql & "INTO tmp2  FROM IvInventario" & " INNER JOIN ((GNComprobante "
        sql = sql & " left join GNCentroCosto on  GNComprobante.IdCentro=GNCentroCosto.IdCentro "
        sql = sql & " LEFT JOIN PCProvCli " & _
             "ON GNComprobante.IdClienteRef " & "= PCProvCli.IdProvCli) " & _
             "INNER JOIN (IVKardex inner join IVBodega on IvKardex.Idbodega = IvBodega.Idbodega) ON GNComprobante.TransID = IVKardex.TransID) " & _
             "ON IVInventario.IdInventario = IVKardex.IdInventario "

        sql = sql & " WHERE GNComprobante.Estado<>3 AND IVInventario.Tipo <> " & INV_TIPONORMAL 'Solo familias
        sql = sql & CondtransFac
        sql = sql & "GROUP BY IVKardex.ID, GNComprobante.FechaTrans, " & trans & _
             ", PCProvCli.Nombre, PCProvCli.CodProvcli, " & _
             "GNComprobante.NumDocRef, " & _
             "IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, " & _
             "IVBodega.CodBodega, IVinventario.Unidad, GNComprobante.IdCentro, codcentro "
        sql = sql & " ORDER BY  PCProvCli.Nombre, GNComprobante.FechaTrans, " & trans
        mEmpOrigen.EjecutarSQL sql, NumReg 'Segundo SQL Transacciones de Facturacion familias
        
        'totaliza entregas
        sql = " select  CodBodega,"
        sql = sql & " idCentro, sum(IsNull(tmp11.cantidad, 0)) As cantidad ,  IDMateria "
        sql = sql & " INTO tmp111 "
        sql = sql & " from tmp11 Inner JOIN IvMateria Ivm ON tmp11.IdInventario = Ivm.IdInventario "
        sql = sql & " group by  CodBodega, IdCentro, IdMateria "
         mEmpOrigen.EjecutarSQL sql, NumReg 'Priemer SQL Transacciones de salida PTerminado

        'totaliza devoluciones
        sql = " select  CodBodega,"
        sql = sql & " idCentro, sum(IsNull(tmp12.cantidad, 0)) As cantidad ,  IDMateria "
        sql = sql & " INTO tmp122 "
        sql = sql & " from tmp12 Inner JOIN IvMateria Ivm ON tmp12.IdInventario = Ivm.IdInventario "
        sql = sql & " group by  CodBodega, IdCentro, IdMateria "
         mEmpOrigen.EjecutarSQL sql, NumReg 'Priemer SQL Transacciones de salida PTerminado


        'une entregas y devoluciones
        sql = " select"
        'sql = sql & " a.FechaTrans , a.CodTrans, a.numtrans, a.nombre, a.NumDocRef, a.IdInventario, "
        sql = sql & " a.IdMateria, "
        'sql = sql & " a.CodInventario, a.Descripcion, a.CodBodega, a.IdCentro,"
        sql = sql & "  a.CodBodega, a.IdCentro,"
        sql = sql & " isnull(a.cantidad,0) as entregado,"
        sql = sql & " isnull(b.cantidad,0) as devolucion, isnull(a.cantidad,0)+ isnull(b.cantidad,0) as saldo"
        'sql = sql & "  isnull(a.cantidad,0)+ isnull(b.cantidad,0) as cantidad "
        sql = sql & " Into tmp1 "
        sql = sql & " from tmp111 a left join tmp122 b on a.IdCentro=b.IdCentro and a.IdMateria=b.IdMateria and a.codbodega=b.codbodega "
        mEmpOrigen.EjecutarSQL sql, NumReg 'Priemer SQL Transacciones de salida PTerminado


        
        
        '--------------------------------------------------------------------------------------------------------
        '--Transforma tmp1 en Familia  totalizar por Familias  y por Centro de costo



        sql = "Select tmp1.CodBodega, sum(tmp1.Entregado) As Cantidad, sum(tmp1.devolucion) As Devolucion, tmp1.IDCentro,IDMateria " & _
              "INTO tmp3 From tmp1  " & _
              "Group By tmp1.CodBodega,tmp1.IdCentro, IdMateria "
        mEmpOrigen.EjecutarSQL sql, NumReg 'tercer SQL Transforma PTerminado a Familias
        'Unir 2 tablas para comparar


        cond = "Where tmp2.IdInventario = tmp3.IDMateria "
        'If mObjCond.Bandera Then
           cond = cond & " AND (tmp2.Cantidad - tmp3.Cantidad - tmp3.devolucion )*-1 <> 0 "
        'End If

         sql = "Select tmp2.ID, tmp2.FechaTrans, tmp2.CodTrans, tmp2.NumTrans, tmp2.Codprovcli, tmp2.Nombre, " & _
               "tmp2.CodInventario, tmp2.Descripcion, tmp2.CodBodega, tmp2.Cantidad*-1 as Cantidad, " & _
               "tmp3.Cantidad *-1 As Despachado, tmp3.Devolucion  As Devolucion, (tmp2.Cantidad - tmp3.Cantidad - tmp3.devolucion)*-1 As Pendiente, tmp2.Unidad,  " & _
               "tmp2.IdCentro, tmp2.CodCentro into tmp4 " & _
               "From tmp2 Left Join  tmp3 ON  tmp2.IdCentro = tmp3.IdCentro AND TMP2.CodBodega=TMP3.CodBodega " & _
               cond
        mEmpOrigen.EjecutarSQL sql, NumReg 'tercer SQL Transforma PTerminado a Familias

'Explicacion del select despues deel UNION
'Preimero quitamos los que ya estan  calculados
'Luego quitamos todos los que tiene saldo cero

        sql = "Select tmp4.FechaTrans, tmp4.CodTrans, tmp4.NumTrans, tmp4.CodCentro, tmp4.Codprovcli,tmp4.Nombre,  " & _
               "tmp4.CodInventario, tmp4.Descripcion, tmp4.CodBodega, tmp4.Cantidad, " & _
               "tmp4.Despachado,  tmp4.Devolucion, tmp4.Pendiente From tmp4  " & _
               "Union " & _
               "Select tmp2.FechaTrans, tmp2.CodTrans, tmp2.NumTrans, tmp2.CodCentro, tmp2.Codprovcli, tmp2.Nombre,  tmp2.CodInventario, " & _
               "tmp2.Descripcion, tmp2.CodBodega, tmp2.Cantidad*-1, 0 As Despachado, 0 as devolucion, (tmp2.Cantidad)*-1 As Pendiente " & _
               " " & _
               "From tmp2 " & _
               "Where tmp2.Id   NOT IN ( " & _
               " Select tmp4.Id from tmp4  ) " & _
               "AND " & _
               "tmp2.ID NOT IN( " & _
               "Select tmp2.ID " & _
               "From tmp2 Left Join  tmp3 ON  tmp2.IdCentro = tmp3.IdCentro " & _
               "Where   tmp2.IdInventario = tmp3.IDMateria AND (tmp2.Cantidad - tmp3.Cantidad - tmp3.devolucion)*-1 = 0) "
        sql = sql & "Order By FechaTrans "
        MensajeStatus MSG_PREPARA, vbHourglass
        'MiGetRowsRep gobjMain.EmpresaActual.OpenRecordset(sql), grd
'        MiGetRowsRep mEmpOrigen.OpenRecordset(sql), grd
        Set rs = mEmpOrigen.OpenRecordset(sql)
        
        
    If Not rs.EOF Then
        v = MiGetRows(rs)
        With grd
            .Redraw = flexRDNone
            .LoadArray v            'Carga a la grilla
            .FormatString = "^#|<Fecha|<CodTrans|<# Trans|Cod. CC|<Cod Cliente|<Cliente|<Cod Item|<Descripcion Item|<Cod Bodega|>Facturado|>Entregado|>Devuelto|>Saldo|<Resultado"
            GNPoneNumFila grd, False
            AsignarTituloAColKey grd            'Para usar ColIndex
            AjustarAutoSize grd, 0, -1, 3000     'Ajusta automáticamente ancho de cols.
            .ColWidth(.ColIndex("Resultado")) = 800
''            'Tipo de datos
            .ColDataType(.ColIndex("Resultado")) = flexDTString
            .ColHidden(4) = True
            .ColHidden(5) = True
            .ColHidden(7) = True
            .Redraw = flexRDDirect
        End With
End If
        
        
        
        'Contorlar  campo  afectacantidad
        'Elimina las tablas temporales
        VerificaExistenciaTablaOrigen 1
        VerificaExistenciaTablaOrigen 2
        VerificaExistenciaTablaOrigen 3
        VerificaExistenciaTablaOrigen 4
    End With
    Exit Sub
ErrTrap:
    VerificaExistenciaTabla 1
    VerificaExistenciaTabla 2
    VerificaExistenciaTabla 3
    VerificaExistenciaTabla 4
    grd.Redraw = True
    DispErr
End Sub


Private Sub CargaPendienteHormi()
    Dim sql As String, cond As String, trans As String
    Dim CondTransSal As String, CondtransFac As String, NumReg As Long, codTransDev As String
    Dim Fcorte As Date, rs As Recordset, v As Variant
    On Error GoTo ErrTrap
    With mObjCond
        'Parametros devueltos: .TipoTrans = lista de transacciones de  salida
        '.SQLItem = lista de transacciones de Factura
        Fcorte = dtpFechaCorte.value    'Fecha de corte
        cond = " AND (GNComprobante.FechaTrans <= " & _
                FechaYMD(Fcorte, gobjMain.EmpresaActual.TipoDB) & ") "
                
        If Len(.tipoTrans) > 0 Then
           CondTransSal = cond & " AND  GNComprobante.CodTrans IN (" & PreparaCadena(.tipoTrans) & ") "
        End If
        CondtransFac = cond & " AND  GNComprobante.CodTrans IN (" & PreparaCadena(.SQLItem) & ") "

        If Len(.CodTrans) > 0 Then
           codTransDev = cond & " AND  GNComprobante.CodTrans IN (" & PreparaCadena(.CodTrans) & ") "
        End If


        trans = " CodTrans, NumTrans"
        'Elimina tablas temporales
        
        VerificaExistenciaTablaOrigen 1
        VerificaExistenciaTablaOrigen 11
        VerificaExistenciaTablaOrigen 12
        VerificaExistenciaTablaOrigen 111
        VerificaExistenciaTablaOrigen 122
        
        VerificaExistenciaTablaOrigen 2
        VerificaExistenciaTablaOrigen 3
        VerificaExistenciaTablaOrigen 4
        
        sql = "SELECT GNComprobante.FechaTrans, " & trans & " ,  PCProvCli.Nombre, " & _
              "GNComprobante.NumDocRef,IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, IVBodega.CodBodega, " & _
              "sum(IVKardex.Cantidad) As Cantidad, GNComprobante.IdCentro "

        sql = sql & "INTO tmp11  FROM IvInventario" & " INNER JOIN ((GNComprobante LEFT JOIN PCProvCli " & _
             "ON GNComprobante.IdClienteRef " & "= PCProvCli.IdProvCli) " & _
             "INNER JOIN (IVKardex inner join IVBodega on IvKardex.Idbodega = IvBodega.Idbodega) ON GNComprobante.TransID = IVKardex.TransID) " & _
             "ON IVInventario.IdInventario = IVKardex.IdInventario "

        sql = sql & " WHERE GNComprobante.Estado<>3"
        sql = sql & CondTransSal
        sql = sql & "GROUP BY GNComprobante.FechaTrans, " & trans & _
             ", PCProvCli.Nombre, " & _
             "GNComprobante.NumDocRef, " & _
             "IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, " & _
             "IVBodega.CodBodega, gnComprobante.IdCentro "
        sql = sql & " ORDER BY  PCProvCli.Nombre, GNComprobante.FechaTrans, " & trans
        grd.Redraw = False
        mEmpOrigen.EjecutarSQL sql, NumReg 'Priemer SQL Transacciones de salida PTerminado
        
        '----------------------------------------------------------------------------------
        'devoluciones
        '----------------------------------------------------------------------------------
        sql = "SELECT GNComprobante.FechaTrans, " & trans & " ,  PCProvCli.Nombre, " & _
              "GNComprobante.NumDocRef,IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, IVBodega.CodBodega, " & _
              "sum(IVKardex.Cantidad) As Cantidad, GNComprobante.IdCentro "

        sql = sql & "INTO tmp12  FROM IvInventario" & " INNER JOIN ((GNComprobante LEFT JOIN PCProvCli " & _
             "ON GNComprobante.IdClienteRef " & "= PCProvCli.IdProvCli) " & _
             "INNER JOIN (IVKardex inner join IVBodega on IvKardex.Idbodega = IvBodega.Idbodega) ON GNComprobante.TransID = IVKardex.TransID) " & _
             "ON IVInventario.IdInventario = IVKardex.IdInventario "

        sql = sql & " WHERE GNComprobante.Estado<>3"
        sql = sql & codTransDev
        sql = sql & "GROUP BY GNComprobante.FechaTrans, " & trans & _
             ", PCProvCli.Nombre, " & _
             "GNComprobante.NumDocRef, " & _
             "IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, " & _
             "IVBodega.CodBodega, gnComprobante.IdCentro "
        sql = sql & " ORDER BY  PCProvCli.Nombre, GNComprobante.FechaTrans, " & trans
        grd.Redraw = False
         mEmpOrigen.EjecutarSQL sql, NumReg 'Priemer SQL Transacciones de salida PTerminado
        
        'totaliza entregas
        sql = " select IdInventario, CodBodega,"
        sql = sql & " idCentro, sum(IsNull(cantidad, 0)) As cantidad"
        sql = sql & " INTO tmp111 "
        sql = sql & " from tmp11"
        sql = sql & " group by IdInventario, CodBodega, IdCentro"
         mEmpOrigen.EjecutarSQL sql, NumReg 'Priemer SQL Transacciones de salida PTerminado


        'totaliza devoluciones
        sql = " select IdInventario, CodBodega,"
        sql = sql & " idCentro, sum(IsNull(cantidad, 0)) As cantidad"
        sql = sql & " INTO tmp122 "
        sql = sql & " from tmp12"
        sql = sql & " group by IdInventario, CodBodega, IdCentro"
         mEmpOrigen.EjecutarSQL sql, NumReg 'Priemer SQL Transacciones de salida PTerminado

        
        
        'une entregas y devoluciones
        sql = " select"
        'sql = sql & " a.FechaTrans , a.CodTrans, a.numtrans, a.nombre, a.NumDocRef, a.IdInventario, "
        sql = sql & " a.IdInventario, "
        'sql = sql & " a.CodInventario, a.Descripcion, a.CodBodega, a.IdCentro,"
        sql = sql & "  a.CodBodega, a.IdCentro,"
        sql = sql & " isnull(a.cantidad,0) as entregado,"
        sql = sql & " isnull(b.cantidad,0) as devolucion, isnull(a.cantidad,0)+ isnull(b.cantidad,0) as saldo"
        'sql = sql & "  isnull(a.cantidad,0)+ isnull(b.cantidad,0) as cantidad "
        sql = sql & " Into tmp1 "
        sql = sql & " from tmp111 a left join tmp122 b on a.IdCentro=b.IdCentro and a.IdInventario=b.IdInventario"
        mEmpOrigen.EjecutarSQL sql, NumReg 'Priemer SQL Transacciones de salida PTerminado
        
        
        
        
        '----------------------------------------------------------------------------------
        sql = "SELECT IVKardex.ID, GNComprobante.FechaTrans, " & trans & " , PCProvCli.Codprovcli, PCProvCli.Nombre, " & _
              "GNComprobante.NumDocRef,IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, IVBodega.CodBodega, " & _
              "sum(IVKardex.Cantidad) As Cantidad, IvInventario.Unidad, GNComprobante.IdCentro, codcentro  "

        sql = sql & "INTO tmp2  FROM IvInventario" & " INNER JOIN ((GNComprobante "
        sql = sql & " left join GNCentroCosto on  GNComprobante.IdCentro=GNCentroCosto.IdCentro "
        sql = sql & " LEFT JOIN PCProvCli " & _
             "ON GNComprobante.IdClienteRef " & "= PCProvCli.IdProvCli) " & _
             "INNER JOIN (IVKardex inner join IVBodega on IvKardex.Idbodega = IvBodega.Idbodega) ON GNComprobante.TransID = IVKardex.TransID) " & _
             "ON IVInventario.IdInventario = IVKardex.IdInventario "

        sql = sql & " WHERE GNComprobante.Estado<>3 AND IVInventario.Tipo = " & INV_TIPONORMAL '
        sql = sql & CondtransFac
        sql = sql & "GROUP BY IVKardex.ID, GNComprobante.FechaTrans, " & trans & _
             ", PCProvCli.Nombre, PCProvCli.CodProvcli, " & _
             "GNComprobante.NumDocRef, " & _
             "IVInventario.IdInventario, IVInventario.CodInventario, IVInventario.Descripcion, " & _
             "IVBodega.CodBodega, IVinventario.Unidad, GNComprobante.IdCentro, codcentro "
        sql = sql & " ORDER BY  PCProvCli.Nombre, GNComprobante.FechaTrans, " & trans
            mEmpOrigen.EjecutarSQL sql, NumReg 'Segundo SQL Transacciones de Facturacion familias
            
            
            
        '--------------------------------------------------------------------------------------------------------
        
        '--Transforma tmp1 en Familia  totalizar por Familias  y por Centro de costo

'        sql = "Select tmp1.IdInventario, tmp1.CodBodega, sum(tmp1.Cantidad) As Cantidad, tmp1.IDCentro " & _
              "INTO tmp3 From tmp1  " & _
              "Group By tmp1.IdInventario, tmp1.CodBodega,tmp1.IdCentro "
 '       mEmpOrigen.EjecutarSQL sql, NumReg 'tercer SQL Transforma PTerminado a Familias


        sql = "Select tmp1.IdInventario, tmp1.CodBodega, sum(tmp1.Entregado) As Cantidad, sum(tmp1.devolucion) as Devolucion, tmp1.IDCentro " & _
              "INTO tmp3 From tmp1  " & _
              "Group By tmp1.IdInventario, tmp1.CodBodega,tmp1.IdCentro "
        mEmpOrigen.EjecutarSQL sql, NumReg 'tercer SQL Transforma PTerminado a Familias
        'Unir 2 tablas para comparar


        cond = "Where tmp2.IdInventario = tmp3.IdInventario "
        'If mObjCond.Bandera Then
           cond = cond & " AND (tmp2.Cantidad - tmp3.Cantidad - tmp3.Devolucion)*-1 <> 0 "
        'End If

'         sql = "Select tmp2.ID, tmp2.FechaTrans, tmp2.CodTrans, tmp2.NumTrans, tmp2.Codprovcli, tmp2.Nombre, " & _
               "tmp2.CodInventario, tmp2.Descripcion, tmp2.CodBodega, tmp2.Cantidad*-1 as Cantidad, " & _
               "tmp3.Cantidad *-1 As Despachado, (tmp2.Cantidad - tmp3.Cantidad)*-1 As Pendiente, tmp2.Unidad,  " & _
               "tmp2.IdCentro, tmp2.CodCentro into tmp4 " & _
               "From tmp2 Left Join  tmp3 ON  tmp2.IdCentro = tmp3.IdCentro AND TMP2.CodBodega=TMP3.CodBodega " & _
               cond
        
         sql = "Select tmp2.ID, tmp2.FechaTrans, tmp2.CodTrans, tmp2.NumTrans, tmp2.Codprovcli, tmp2.Nombre, " & _
               "tmp2.CodInventario, tmp2.Descripcion, tmp2.CodBodega, tmp2.Cantidad*-1 as Cantidad, " & _
               "tmp3.Cantidad *-1 As Despachado, tmp3.Devolucion  As Devolucion, (tmp2.Cantidad - tmp3.Cantidad- tmp3.devolucion)*-1 As Pendiente, tmp2.Unidad,  " & _
               "tmp2.IdCentro, tmp2.CodCentro into tmp4 " & _
               "From tmp2 Left Join  tmp3 ON  tmp2.IdCentro = tmp3.IdCentro AND TMP2.CodBodega=TMP3.CodBodega " & _
               cond
        
        
        mEmpOrigen.EjecutarSQL sql, NumReg 'tercer SQL Transforma PTerminado a Familias

'Explicacion del select despues deel UNION
'Preimero quitamos los que ya estan  calculados
'Luego quitamos todos los que tiene saldo cero

        sql = "Select tmp4.FechaTrans, tmp4.CodTrans, tmp4.NumTrans, tmp4.CodCentro, tmp4.Codprovcli,tmp4.Nombre,  " & _
               "tmp4.CodInventario, tmp4.Descripcion, tmp4.CodBodega, tmp4.Cantidad, " & _
               "tmp4.Despachado, tmp4.Devolucion, tmp4.Pendiente From tmp4  " & _
               "Union " & _
               "Select tmp2.FechaTrans, tmp2.CodTrans, tmp2.NumTrans, tmp2.CodCentro, tmp2.Codprovcli, tmp2.Nombre,  tmp2.CodInventario, " & _
               "tmp2.Descripcion, tmp2.CodBodega, tmp2.Cantidad*-1, 0 As Despachado, 0 as devolucion, (tmp2.Cantidad)*-1 As Pendiente " & _
               " " & _
               "From tmp2 " & _
               "Where tmp2.Id  NOT IN ( " & _
               " Select tmp4.Id from tmp4  ) " & _
               "AND " & _
               "tmp2.ID NOT IN( " & _
               "Select tmp2.ID " & _
               "From tmp2 Left Join  tmp3 ON  tmp2.IdCentro = tmp3.IdCentro " & _
               "Where   tmp2.IdInventario = tmp3.IDInventario AND (tmp2.Cantidad - tmp3.Cantidad - tmp3.Devolucion)*-1 = 0) "
        sql = sql & "Order By FechaTrans "
        MensajeStatus MSG_PREPARA, vbHourglass
        'MiGetRowsRep gobjMain.EmpresaActual.OpenRecordset(sql), grd
        MiGetRowsRep mEmpOrigen.OpenRecordset(sql), grd
        Set rs = mEmpOrigen.OpenRecordset(sql)
        
        
    If Not rs.EOF Then
        v = MiGetRows(rs)
        With grd
            .Redraw = flexRDNone
            .LoadArray v            'Carga a la grilla
            .FormatString = "^#|<Fecha|<CodTrans|<# Trans|Cod. CC|<Cod Cliente|<Cliente|<Cod Item|<Descripcion Item|<Cod Bodega|>Facturado|>Entregado|>Devuelto|>Saldo|<Resultado"
            GNPoneNumFila grd, False
            AsignarTituloAColKey grd            'Para usar ColIndex
            AjustarAutoSize grd, 0, -1, 3200     'Ajusta automáticamente ancho de cols.
            .ColWidth(.ColIndex("Resultado")) = 800
''            'Tipo de datos
            .ColDataType(.ColIndex("Resultado")) = flexDTString
            .ColHidden(4) = True
            .ColHidden(5) = True
            .ColHidden(7) = True
            .Redraw = flexRDDirect
        End With
End If
        
        
        
        'Contorlar  campo  afectacantidad
        'Elimina las tablas temporales
        VerificaExistenciaTablaOrigen 1
        VerificaExistenciaTablaOrigen 2
        VerificaExistenciaTablaOrigen 3
        VerificaExistenciaTablaOrigen 4
    End With
    Exit Sub
ErrTrap:
    VerificaExistenciaTabla 1
    VerificaExistenciaTabla 2
    VerificaExistenciaTabla 3
    VerificaExistenciaTabla 4
    grd.Redraw = True
    DispErr
End Sub




Private Function CopiaTransHormi() As Boolean
    Dim i As Long, codt As String, numt As Long
    Dim empDestino As Empresa, sql As String, Num As Long, ultimaFila As Long, j As Long
    'Verifica  errores  en la base de Origen
    
    Set empDestino = AbrirDestino
    If empDestino.NombreDB = mEmpOrigen.NombreDB Then
        MsgBox "La empresa origen y destino son las mismas" & Chr(13) & _
               "debera  seleccionar  una empresa de  destino diferente", vbExclamation
        Exit Function
    End If
    
    'If grd.FixedRows + 1 = grd.Rows Then
    CargaPendienteHormi 'Carga  trans solo  si la grlla esta vacia
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

        i = .FixedRows
        While i <= .Rows - 1
        'For i = .FixedRows To .Rows - 1
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
                numt = .TextMatrix(i, .ColIndex("# Trans"))
                'If codt = "FC" And numt = 2524 Then Stop
                MensajeStatus "Copiando la transacción " & codt & numt & _
                            "     " & i & " de " & .Rows - .FixedRows & _
                            " (" & Format(i * 100 / (.Rows - .FixedRows), "0") & "%)", vbHourglass
                
                'Si aún no está importado bien, importa la fila
                If grd.TextMatrix(i, .Cols - 1) <> MSG_OK Then
                
                    If grd.ValueMatrix(i, .Cols - 2) > 0 Then
                        ultimaFila = i
                        If ImportarTransSub(codt, numt, empDestino, i) Then
                            For j = ultimaFila To i
                                .TextMatrix(j, .ColIndex("Resultado")) = MSG_OK
                            Next j
                        Else
                            For j = ultimaFila To i
                                .TextMatrix(i, .ColIndex("Resultado")) = "Error"
                            Next j
                        End If
                    Else
                        .TextMatrix(i, .ColIndex("Resultado")) = "Cantidad Negativa"
                    End If
               End If
               i = i + 1
        Wend
       'Next i
    End With
  
    MsgBox "Proceso terminado con exito"
    
    
    MensajeStatus
    mbooProcesando = False  'Bloquea que se cierre la ventana
    CopiaTransHormi = True
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


Private Function ImportarTransSubItem( _
                ByVal codt As String, _
                ByVal numt As Long, ByRef empDestino As Empresa, _
                ByRef fila As Long) As Boolean
    Dim gnDest As GNComprobante, s As String, Estado As Byte, gnOri As GNComprobante
    Dim ivk As IVKardex
    Dim i As Long, contItems As Long, j As Long
    Dim BandOtro As Boolean, BandEncontro As Boolean
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
    
    'elimina todo el ivkardex para cargar solo lo pendiente
    gnDest.BorrarIVKardex
    'verifica y corrige cantidades
    gnDest.CodClienteRef = grd.TextMatrix(fila, COL_CODCLI)
    gnDest.CodCentro = grd.TextMatrix(fila, COL_CODCC)
    i = 1
    contItems = 0
    While i <= gnOri.CountIVKardex
        BandOtro = True
        BandEncontro = False
        
        While BandOtro
            If gnOri.IVKardex(i).CodInventario = grd.TextMatrix(fila, COL_CODITEM) Then
                contItems = contItems + 1
                gnDest.AddIVKardex
                gnDest.IVKardex(contItems).CodInventario = grd.TextMatrix(fila, COL_CODITEM)
                gnDest.IVKardex(contItems).CodBodega = grd.TextMatrix(fila, COL_CODBODEGA)
                If gnDest.GNTrans.IVTipoTrans = "E" Then
                    gnDest.IVKardex(contItems).cantidad = grd.TextMatrix(fila, COL_SALDOITEM) * -1
                Else
                    gnDest.IVKardex(contItems).cantidad = grd.TextMatrix(fila, COL_SALDOITEM)
                End If
                gnDest.IVKardex(contItems).Orden = contItems
                gnDest.IVKardex(contItems).CostoTotal = gnOri.IVKardex(i).CostoTotal
                gnDest.IVKardex(contItems).CostoRealTotal = gnOri.IVKardex(i).CostoRealTotal
                gnDest.IVKardex(contItems).PrecioTotal = gnOri.IVKardex(i).PrecioTotal
                gnDest.IVKardex(contItems).PrecioRealTotal = gnOri.IVKardex(i).PrecioRealTotal
                gnDest.IVKardex(contItems).Descuento = gnOri.IVKardex(i).Descuento
                gnDest.IVKardex(contItems).IVA = gnOri.IVKardex(i).IVA
                gnDest.IVKardex(contItems).Nota = gnOri.IVKardex(i).Nota
                gnDest.IVKardex(contItems).NumeroPrecio = gnOri.IVKardex(i).NumeroPrecio
                gnDest.IVKardex(contItems).ValorRecargoItem = gnOri.IVKardex(i).ValorRecargoItem
                gnDest.IVKardex(contItems).TiempoEntrega = gnOri.IVKardex(i).TiempoEntrega
                gnDest.IVKardex(contItems).IdICE = gnOri.IVKardex(i).IdICE
                BandEncontro = True
            End If
            
            'fila = fila + 1
            i = i + 1
            If BandEncontro Then
                If Not fila = grd.Rows - 1 Then
                    If grd.TextMatrix(fila + 1, COL_CODTRANS) = codt And grd.TextMatrix(fila + 1, COL_NUMTRANS) = numt Then
                        fila = fila + 1
                        BandOtro = True
                        i = 1
                        BandEncontro = False
'                        Fila = 5
                    Else
                        BandOtro = False
                        i = gnOri.CountIVKardex + 1
                    End If
                Else
                    BandOtro = False
                End If
            End If
            
        Wend
    Wend

    
    gnDest.FechaTrans = DateAdd("d", 1, dtpFechaCorte.value)
    gnDest.Descripcion = "PENDIENTE DESDE " & gnOri.FechaTrans
    
    
    
    For j = 1 To gnDest.CountPCKardex
        gnDest.RemovePCKardex (1)
    Next j
        
    For j = 1 To gnDest.CountIVKardexRecargo
        gnDest.RemoveIVKardexRecargo (1)
    Next j
        
        
    gnDest.Grabar False, False
    
'    Forzar el valor de Estado original, debido a que al Grabar cambia sin querer
    On Error Resume Next
    If gnDest.Estado = 1 And Estado = 3 Then
        'Primero Cambia  a estado cero
        empDestino.CambiaEstadoGNCompCierre gnDest.TransID, 0
    End If
    'Para  que no  considere  el IdAsignado
    empDestino.CambiaEstadoGNCompCierre gnDest.TransID, Estado
    ImportarTransSubItem = True
salida:
    Set gnDest = Nothing
    Set gnOri = Nothing
    Exit Function
ErrTrap:
    If MsgBox(Err.Description & vbCr & vbCr & _
                "Desea continuar con siguiente transacción?", _
                vbQuestion + vbYesNo) <> vbYes Then
    End If
    GoTo salida
End Function


Private Sub LeerArchivo(ByVal archi As String)
    Select Case UCase$(Right$(archi, 4))
        Case ".TXT"
            ReformartearColumnas
            VisualizarTexto archi
            InsertarColumnas
        Case Else
        End Select
End Sub

Private Sub ReformartearColumnas()
' SOLO EN ESTOS CASOS
'Select Case UCase(Me.tag)
'    Case "DIARIO", "INVENTARIO"
        ConfigCols
'End Select

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
    Select Case Me.tag
        Case "PLANCUENTA"
            grd.Select 1, 1, 1, 1
        Case "ITEM"
            grd.Select 1, 1, 1, 1
        Case "PCPROV"
            grd.Select 1, 1, 1, 1
        Case "PCCLI"
            grd.Select 1, 1, 1, 1
        Case "PORPAGAR"
            grd.Select 1, 1, 1, 4
        Case "PORCOBRAR"
            grd.Select 1, 1, 1, 4
        Case "DIARIO"
            grd.Select 1, 1, 1, 1
        Case "INVENTARIO"
            grd.Select 1, 1, 1, 2
        Case "VENTASLOCUTORIOS"
            'DEB0 ORDENAR SOLO POR LA COLUMNA DE NUMERO DE NOTA DE VENTA
            For i = 0 To 18
                grd.ColSort(i) = flexSortNone
            Next i
            grd.ColSort(19) = flexSortGenericAscending
            grd.Select 1, 1, 1, 19
    End Select
    End If
    grd.Sort = flexSortUseColSort

' poner numero
    GNPoneNumFila grd, False
    
    
    grd.Redraw = flexRDDirect
    AjustarAutoSize grd, -1, -1
    
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

Private Sub InsertarColumnas()
Dim i As Integer
Dim sql As String, rs As Recordset


'Select Case UCase(Me.tag)
'    Case "DIARIO"
'        'sumar
        With grd
'
'        InsertarColumnaCuenta
        For i = .FixedRows To .Rows - 1 ' poner nombre de cuentas en columna cuenta
            If Len(.TextMatrix(i, .ColIndex("CodTrans"))) > 0 Then
                sql = " select codprovcli, CodCentro from gncomprobante g "
                sql = sql & " inner join pcprovcli p on g.idclienteref=p.idprovcli inner join "
                sql = sql & " gncentrocosto gn on g.idcentro=gn.idcentro"
                sql = sql & " where codtrans='" & .TextMatrix(i, .ColIndex("CodTrans"))
                sql = sql & " ' and numtrans=" & .TextMatrix(i, .ColIndex("# Trans"))
                sql = sql & "  and G.ESTADO <> 3"
                
                Set rs = mEmpOrigen.OpenRecordset(sql)
                If Not rs.EOF Then
                    .TextMatrix(i, .ColIndex("Cod. CC")) = rs.Fields("CodCentro")
                    .TextMatrix(i, .ColIndex("Cod Cliente")) = rs.Fields("CodProvcli")
                End If
            End If
'            DoEvents
'            .TextMatrix(i, .ColIndex("Cuenta")) = ponerCuentaFila(.TextMatrix(i, 1))
        Next i
'        Sumar
        End With
'    Case "INVENTARIO"
'        InsertarColumnaDesc_y_Cost
'        For i = grd.FixedRows To grd.Rows - 1 ' poner nombre de cuentas en columna cuenta
'            DoEvents
'            grd.TextMatrix(i, grd.ColIndex("Descripción")) = ponerDescripcionFila(grd.TextMatrix(i, 1))
'            grd.TextMatrix(i, grd.ColIndex("Costo Unitario")) = ponerCostoUnitarioFila(i)
'        Next i
'End Select
AjustarAutoSize grd, -1, -1
End Sub

Private Sub ConfigCols()
    Dim s As String, i As Integer
    grd.Cols = 1
    grd.FormatString = "^#|<Fecha|<CodTrans|<# Trans|Cod. CC|<Cod Cliente|<Cliente|<Cod Item|<Descripcion Item|<Cod Bodega|>Facturado|>Entregado|>Saldo|<Resultado"
'    grd.ColSort(1) = flexSortGenericAscending
'    grd.ColSort(2) = flexSortGenericAscending
'    grd.ColSort(3) = flexSortGenericAscending
    'grd.ColSort(4) = flexSortGenericAscending
    
    'Asigna a ColKey los títulos de columnas
    ' para luego poder referirnos a la columna con su título mismo
    AsignarTituloAColKey grd
    grd.SetFocus
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

Private Sub grd_KeyDown(KeyCode As Integer, Shift As Integer)
    If grd.IsSubtotal(grd.Row) Or Me.tag = "VENTASLOCUTORIOS" Then Exit Sub    ' si la fila es de subtotal o
    Select Case KeyCode                                                       ' esta importanto las ventas del locutorio para el Sii no las puede editar
    Case vbKeyDelete
        EliminarFila
    End Select
End Sub

Private Function CopiaTransFamiliasArchivo() As Boolean
    Dim i As Long, codt As String, numt As Long
    Dim empDestino As Empresa, sql As String, Num As Long, ultimaFila As Long, j As Long
    'Verifica  errores  en la base de Origen
    
    Set empDestino = AbrirDestino
    
    If grd.FixedRows = grd.Rows Then CargaPendientexFamilia 'Carga  trans solo  si la grlla esta vacia
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

        i = .FixedRows
        While i <= .Rows - 1
        'For i = .FixedRows To .Rows - 1
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
                numt = .ValueMatrix(i, .ColIndex("# Trans"))
                'If codt = "FC" And numt = 2524 Then Stop
                MensajeStatus "Copiando la transacción " & codt & numt & _
                            "     " & i & " de " & .Rows - .FixedRows & _
                            " (" & Format(i * 100 / (.Rows - .FixedRows), "0") & "%)", vbHourglass
                
                'Si aún no está importado bien, importa la fila
                If grd.TextMatrix(i, .Cols - 1) <> MSG_OK Then
                    If grd.ValueMatrix(i, .Cols - 2) > 0 Then
                    ultimaFila = i
                    If ImportarTransSub(codt, numt, empDestino, i) Then
                        For j = ultimaFila To i
                            .TextMatrix(j, .ColIndex("Resultado")) = MSG_OK
                        Next j
                    Else
                        For j = ultimaFila To i
                            .TextMatrix(i, .ColIndex("Resultado")) = "Error"
                        Next j
                    End If
                    Else
                        .TextMatrix(i, .ColIndex("Resultado")) = "Cantidad Negativa"
                    End If
               End If
               i = i + 1
        Wend
       'Next i
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
    CopiaTransFamiliasArchivo = True
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

