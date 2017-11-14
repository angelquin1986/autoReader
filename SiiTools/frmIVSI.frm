VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIVSI 
   Caption         =   "Creación de IVSI según IVExist"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   6105
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   612
      Left            =   4440
      TabIndex        =   2
      Top             =   480
      Width           =   1332
   End
   Begin MSComctlLib.ProgressBar prg1 
      Height          =   372
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   3612
      _ExtentX        =   6376
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar Transacciones de Saldo Inicial (IVSI) según la existencia "
      Height          =   612
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3732
   End
End
Attribute VB_Name = "frmIVSI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbooProcesando As Boolean
Private mbooCancelado As Boolean
Private mEmpOrigen As Empresa
Private WithEvents mGrupo As grupo
Attribute mGrupo.VB_VarHelpID = -1

Private Function AbrirDestino() As Empresa
    Dim e As Empresa, cod As String
    
    cod = "Hormi2001"
    Set e = gobjMain.RecuperaEmpresa(cod)
    e.Abrir
    Set AbrirDestino = e
    Set e = Nothing
End Function

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGenerar_Click()
    SaldoIV
End Sub

Private Sub mensaje( _
                ByVal NuevaFila As Boolean, _
                ByVal proc As String, _
                Optional ByVal res As String)
'
End Sub

'4. Pasar saldo inicial de inventario
Private Function SaldoIV() As Boolean
    Dim e As Empresa, gc As GNComprobante, ivk As IVKardex, iv As IVinventario
    Dim j As Long, n As Long
    Dim sql As String, rs As Recordset, codOrig As String
    Dim i As Long, c As Currency, Fcorte As Date
    On Error GoTo ErrTrap
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    
    codOrig = gobjMain.EmpresaActual.CodEmpresa
    Fcorte = #12/31/2000#

    'Cambia figura de cursor de mouse
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True
    
    'Saca las existencias a la fecha de corte
    MensajeStatus "Preparando para grabar las existencias iniciales...", vbHourglass
    mensaje True, "Saldo inicial de inventario..."
    
'    sql = "SELECT ivk.IdInventario, ivk.IdBodega, " & _
'                "iv.CodInventario, ivb.CodBodega, " & _
'                "Sum(ivk.Cantidad) AS Exist " & _
'          "FROM IVBodega ivb INNER JOIN " & _
'                    "(IVInventario iv INNER JOIN " & _
'                        "(GNTrans gt INNER JOIN " & _
'                            "(IVKardex ivk INNER JOIN GNComprobante gc " & _
'                            "ON ivk.TransID=gc.TransID) " & _
'                        "ON gt.CodTrans=gc.CodTrans) " & _
'                    "ON iv.IdInventario = ivk.IdInventario) " & _
'                "ON ivb.IdBodega = ivk.IdBodega " & _
'          "WHERE (gc.Estado<>" & ESTADO_ANULADO & ") AND " & _
'                 "(gt.AfectaCantidad=" & CadenaBool(True, gobjMain.EmpresaActual.TipoDB) & ") AND " & _
'                 "(gc.FechaTrans < " & FechaYMD(fcorte + 1, gobjMain.EmpresaActual.TipoDB) & ") AND " & _
'                 "(iv.BandServicio=" & CadenaBool(False, gobjMain.EmpresaActual.TipoDB) & ") " & _
'          "GROUP BY ivk.IdInventario, ivk.IdBodega, iv.CodInventario, ivb.CodBodega " & _
'          "HAVING Sum(ivk.Cantidad)>0"
    sql = "SELECT ive.Exist, ivb.CodBodega, iv.CodInventario " & _
          "FROM IVInventario iv INNER JOIN (IVExist ive INNER JOIN IVBodega ivb " & _
            "ON ive.IdBodega = ivb.IdBodega) " & _
            "ON iv.IdInventario = ive.IdInventario " & _
          "WHERE ive.Exist <>0 "
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
'                Set iv = mEmpOrigen.RecuperaIVInventario(.Fields("IdInventario"))
                Set iv = gobjMain.EmpresaActual.RecuperaIVInventario(.Fields("CodInventario"))
                
                'Obtiene Costo del item en Moneda de item
'                c = iv.Costo(fcorte, 1)
                c = iv.Precio(1)
                
'                'De moneda de item, covierte en moneda de trans, si es necesario
'                If iv.CodMoneda <> gc.CodMoneda Then
'                    c = c * gc.Cotizacion(iv.CodMoneda) / gc.Cotizacion("")
'                End If
                
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
                ivk.orden = i Mod 100
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
    gc.Empresa.CorregirExistencia

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
'        .Cotizacion("USD") = 25000
        .Cotizacion("USD") = 1
        .Descripcion = Desc
        .FechaTrans = fecha + 1
        .numDocRef = numdoc
    End With
    Set CrearTrans = g
    Set g = Nothing
End Function


