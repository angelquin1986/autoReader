VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReasignacionCosto 
   Caption         =   "Recalculo de IVA de items"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   8160
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   852
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   8160
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5520
      Width           =   8160
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Proceder"
         Enabled         =   0   'False
         Height          =   372
         Left            =   2688
         TabIndex        =   4
         Top             =   120
         Width           =   1212
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   4968
         TabIndex        =   3
         Top             =   120
         Width           =   1212
      End
      Begin VB.CommandButton cmdVerificar 
         Caption         =   "&Verificar"
         Enabled         =   0   'False
         Height          =   372
         Left            =   1248
         TabIndex        =   2
         Top             =   120
         Width           =   1212
      End
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   540
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   4305
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6840
      _cx             =   12065
      _cy             =   7594
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
      Cols            =   8
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
Attribute VB_Name = "frmReasignacionCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Constantes para las columnas
Private Const COL_NUMFILA = 0
Private Const COL_TID_FUENTE = 1
Private Const COL_FECHA_FUENTE = 2
Private Const COL_CODTRANS_FUENTE = 3
Private Const COL_NUMTRANS_FUENTE = 4
Private Const COL_TID = 5
Private Const COL_FECHA = 6
Private Const COL_CODTRANS = 7
Private Const COL_NUMTRANS = 8      '*** MAKOTO 07/feb/01 Agregado
Private Const COL_RESULTADO = 9

Private Const MSG_NG = "Costo Incorrecto"
Private Const MSG_DIFF = "Transacciones no iguales"
Private Const MSG_FECHADIFF = "Fechas Diferentes"
Private mProcesando As Boolean
Private mCancelado As Boolean
Private mgrdOrigen As Object

Public Sub Inicio(ByVal grdTrans As Object)
    Dim i As Integer
    On Error GoTo ErrTrap
    Set mgrdOrigen = grdTrans
    Me.Show vbModal
    
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

'*** MAKOTO 31/ago/00 Modificado
Private Sub CargaTrans()
    Dim ant_Height As Long
    Const COL_FEC = 1
    Const COL_COD = 2
    Const COL_NUM = 3
    Const msg = "Buscando registros"
    Dim i As Long, sql As String, rs As ADODB.Recordset
    Dim CodTrans As String, numtrans As Long
    

    On Error GoTo ErrTrap

    ant_Height = Me.Height
    Me.Height = pic1.Height
    mCancelado = False
    
    ConfigCols
    
    With mgrdOrigen
        mProcesando = True
        prg1.max = .Rows - 1
        prg1.min = .FixedRows
        
        For i = .FixedRows To .Rows - 1
                If mCancelado Then
                    Exit For
                End If
                DoEvents
                Me.Caption = msg & " " & i & " de " & prg1.max
                prg1.value = i
                
                'If .IsSelected(i) Then 'Pasa solo las trans. seleccionadas
                                    'porque significa que se importaron bien
                
                CodTrans = .TextMatrix(i, COL_COD)
                numtrans = .TextMatrix(i, COL_NUM)
                sql = "SELECT TransID, FechaTrans, codtrans, numtrans, IDTransfuente " & _
                      "FROM gncomprobante " & _
                      "WHERE idtransfuente = " & _
                      "(Select transid from gncomprobante where codtrans = '" & CodTrans & "'" & _
                      " and NumTrans = " & numtrans & ")"
                Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                If Not (rs.BOF And rs.EOF) Then
                    Do Until rs.EOF
                        grd.AddItem vbTab & rs!IdTransFuente & vbTab & _
                                            .TextMatrix(i, COL_FEC) & vbTab & _
                                            .TextMatrix(i, COL_COD) & vbTab & _
                                            .TextMatrix(i, COL_NUM) & vbTab & _
                                            rs!TransID & vbTab & _
                                            rs!FechaTrans & vbTab & _
                                            rs!CodTrans & vbTab & _
                                            rs!numtrans & vbTab

                        rs.MoveNext
                    Loop
                Else
                    grd.AddItem vbTab & "0" & vbTab & _
                                            .TextMatrix(i, COL_FEC) & vbTab & _
                                            .TextMatrix(i, COL_COD) & vbTab & _
                                            .TextMatrix(i, COL_NUM) & vbTab
                End If
                rs.Close
                Set rs = Nothing
                
                

            'End If
        Next i
    End With
        
    'Da formato a la grilla
    cmdVerificar.Enabled = True
    mProcesando = False
    prg1.value = prg1.min
    ConfigCols
    'Asignar el tamaño normal  ' pero haciendo un ciclo para tener un efecto chevere
    For i = Me.Height To ant_Height Step 500
        DoEvents
        Me.Height = i
    Next i
    Exit Sub
ErrTrap:
    mProcesando = False
    DispErr
End Sub



Private Sub cmdAceptar_Click()
    'Si no hay transacciones
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna transacción para procesar.", vbExclamation
        Exit Sub
    End If
    
    Grabar
    cmdCancelar.SetFocus
'    If ReprocIVA(False) Then
'        cmdCancelar.SetFocus
'    End If
End Sub

Private Sub ConfigCols()
    With grd
        .Cols = 10
        .FormatString = "^#|tidfuente|<Fecha|<CodTrans|<#|<tid|<Fecha|<CodTrans|<#|<Resultado"
        
        .ColHidden(COL_NUMFILA) = False
        .ColHidden(COL_TID) = True
        .ColHidden(COL_FECHA) = False
        .ColHidden(COL_CODTRANS) = False
        .ColHidden(COL_NUMTRANS) = False

        .ColHidden(COL_TID_FUENTE) = True

        .ColHidden(COL_FECHA_FUENTE) = False
        .ColHidden(COL_CODTRANS_FUENTE) = False
        .ColHidden(COL_NUMTRANS_FUENTE) = False
        .ColHidden(COL_RESULTADO) = False

        .ColDataType(COL_FECHA) = flexDTDate    '*** MAKOTO 14/ago/2000 para que ordene bien por fecha
        .ColDataType(COL_FECHA_FUENTE) = flexDTDate    '*** MAKOTO 14/ago/2000 para que ordene bien por fecha
        
        GNPoneNumFila grd, False
        .AutoSize 0, grd.Cols - 1
        
'        .ColWidth(COL_FECHA) = 1100
'        .ColWidth(COL_CODTRANS) = 900
'        .ColWidth(COL_NUMTRANS) = 500
'
'        .ColWidth(COL_FECHA_FUENTE) = 1100
'        .ColWidth(COL_CODTRANS_FUENTE) = 900
'        .ColWidth(COL_NUMTRANS_FUENTE) = 500
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

Private Sub cmdVerificar_Click()
    'Si no hay transacciones
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna transacción para verificar."
        Exit Sub
    End If
    
    If Verificar Then
        cmdAceptar.Enabled = True
        cmdAceptar.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    CargaTrans
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
    grd.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - grd.Top - pic1.Height - 80
    prg1.Width = Me.ScaleWidth - (prg1.Left * 2)
End Sub


Private Function Verificar() As Boolean
    Dim gncFuente As GNComprobante, gnc As GNComprobante
    Dim i As Long, j As Long
    Dim BandDiferente As Boolean, CostoDiferente As Boolean, bandFechaDiff As Boolean
    
    On Error GoTo ErrTrap
    
    If mProcesando Then Exit Function
    
    mProcesando = True
    prg1.max = grd.Rows - 1
    prg1.min = 0
        
    For i = grd.FixedRows To grd.Rows - 1
        prg1.value = i
        DoEvents
        If mCancelado Then
            Set gnc = Nothing
            Set gncFuente = Nothing
            Exit For
        End If
     If grd.TextMatrix(i, COL_TID_FUENTE) <> 0 Then
        'recuperar la transaccion fuente
        Set gncFuente = gobjMain.EmpresaActual.RecuperaGNComprobante(0, grd.TextMatrix(i, COL_CODTRANS_FUENTE), CLng(grd.TextMatrix(i, COL_NUMTRANS_FUENTE)))
        Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(0, grd.TextMatrix(i, COL_CODTRANS), CLng(grd.TextMatrix(i, COL_NUMTRANS)))
        BandDiferente = False
        
        If gncFuente.CountIVKardex <> gnc.CountIVKardex Then
            BandDiferente = True
        Else
            bandFechaDiff = False
            CostoDiferente = False
            For j = 1 To gncFuente.CountIVKardex
                DoEvents
                If gncFuente.IVKardex(j).orden = gnc.IVKardex(j).orden And _
                    gncFuente.IVKardex(j).CodInventario = gnc.IVKardex(j).CodInventario And _
                    Abs(gncFuente.IVKardex(j).cantidad) = Abs(gnc.IVKardex(j).cantidad) Then
                    If (Abs(gncFuente.IVKardex(j).CostoRealTotal) <> Abs(gnc.IVKardex(j).CostoRealTotal)) Or (Abs(gncFuente.IVKardex(j).CostoTotal) <> Abs(gnc.IVKardex(j).CostoTotal)) Then
                        CostoDiferente = True
                    End If
                Else
                    BandDiferente = True
                End If
            Next j
            If (Not BandDiferente) And (Not CostoDiferente) Then
                If gncFuente.FechaTrans <> gnc.FechaTrans Then
                    bandFechaDiff = True
                End If
            End If
        End If
        
        If BandDiferente Then
            grd.TextMatrix(i, COL_RESULTADO) = MSG_DIFF
        Else
            If CostoDiferente Then
                grd.TextMatrix(i, COL_RESULTADO) = MSG_NG
            ElseIf bandFechaDiff Then
                grd.TextMatrix(i, COL_RESULTADO) = MSG_FECHADIFF
            Else
                grd.TextMatrix(i, COL_RESULTADO) = "OK"
            End If
        End If
        Set gnc = Nothing
        Set gncFuente = Nothing
      Else
        grd.TextMatrix(i, COL_RESULTADO) = "---"
      End If
    Next i
    prg1.value = prg1.min
    Verificar = True
    mProcesando = False
Exit Function
ErrTrap:
    Verificar = False
    mProcesando = False
    DispErr
End Function




Private Sub Grabar()
    Dim gncFuente As GNComprobante, gnc As GNComprobante
    Dim i As Long, j As Long, BandDiferente As Boolean, CostoDiferente As Boolean
    
    On Error GoTo ErrTrap
    
    If mProcesando Then Exit Sub
    
    prg1.max = grd.Rows - 1
    prg1.min = 0
    mProcesando = True
    For i = grd.FixedRows To grd.Rows - 1
        prg1.value = i
        DoEvents
        If mCancelado Then
            Set gnc = Nothing
            Set gncFuente = Nothing
            Exit For
        End If
        'recuperar la transaccion fuente
        If grd.TextMatrix(i, COL_TID_FUENTE) <> 0 Then
            Set gncFuente = gobjMain.EmpresaActual.RecuperaGNComprobante(0, grd.TextMatrix(i, COL_CODTRANS_FUENTE), CLng(grd.TextMatrix(i, COL_NUMTRANS_FUENTE)))
            Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(0, grd.TextMatrix(i, COL_CODTRANS), CLng(grd.TextMatrix(i, COL_NUMTRANS)))
            If grd.TextMatrix(i, COL_RESULTADO) = MSG_NG Then
                grd.TextMatrix(i, COL_RESULTADO) = "Actualizando..."
                'ACTUALIZAR COSTOS AL DOCUMENT DEL DOC. FUENTE
                gnc.FechaTrans = gncFuente.FechaTrans  'actualiza la fecha
                For j = 1 To gncFuente.CountIVKardex
                    DoEvents
                    gnc.IVKardex(j).CostoTotal = Abs(gncFuente.IVKardex(j).CostoTotal)
                    gnc.IVKardex(j).CostoRealTotal = Abs(gncFuente.IVKardex(j).CostoRealTotal)
                Next j
                CalculaTotalRec gnc
                gnc.Grabar False, False
                grd.TextMatrix(i, COL_RESULTADO) = "Grabado"
            ElseIf grd.TextMatrix(i, COL_RESULTADO) = MSG_FECHADIFF Then
                grd.TextMatrix(i, COL_RESULTADO) = "Actualizando..."
                gnc.FechaTrans = gncFuente.FechaTrans 'solo actualiza la fecha
'                CalculaTotalRec gnc
                gnc.ProrratearIVKardexRecargo
                gnc.Grabar False, False
                grd.TextMatrix(i, COL_RESULTADO) = "Grabado"
            End If
            Set gnc = Nothing
            Set gncFuente = Nothing
        End If
    Next i
    prg1.value = prg1.min
    'Verificar = True
    mProcesando = False
Exit Sub
ErrTrap:
    'Verificar = False
    mProcesando = False
    DispErr
End Sub


Private Sub CalculaTotalRec(ByRef gnc As GNComprobante)
    Dim i As Long, obj As IVKardexRecargo
    On Error GoTo ErrTrap
    For i = 1 To gnc.CountIVKardexRecargo
        Set obj = gnc.IVKardexRecargo(i)
        If obj.BandOrigen = -2 Then
            If Abs(obj.valor) <> gnc.IVKardexIVAItemTotal Then
                obj.valor = gnc.IVKardexIVAItemTotal
                Exit For
            End If
        End If
    Next i
    Exit Sub
ErrTrap:
    DispErr
End Sub

