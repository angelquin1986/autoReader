VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImprimePorLoteEstadoCuenta 
   Caption         =   "Impresión por lote"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7770
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   7770
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   852
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   7770
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3855
      Width           =   7770
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   372
         Left            =   360
         TabIndex        =   5
         Top             =   0
         Width           =   1452
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   4968
         TabIndex        =   3
         Top             =   0
         Width           =   1212
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   372
         Left            =   2568
         TabIndex        =   2
         Top             =   0
         Width           =   1452
      End
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   4
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
      TabIndex        =   0
      Top             =   1800
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
End
Attribute VB_Name = "frmImprimePorLoteEstadoCuenta"
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
Private Const COL_NUMDOCREF = 6         '*** MAKOTO 07/feb/01 Agregado
Private Const COL_NOMBRE = 7            '*** MAKOTO 07/feb/01 Agregado
Private Const COL_DESC = 8
Private Const COL_CENTROCOSTO = 9
Private Const COL_ESTADO = 10
Private Const COL_RESULTADO = 11

Private Const MSG_NG = "Error en impresión."
Private mProcesando As Boolean
Private mCancelado As Boolean
Private mCodMoneda  As String
Private mobjBusq As Busqueda
Private mObjCond As RepCondicion

Public Sub Inicio()
    Dim i As Integer
    On Error GoTo ErrTrap
    mCodMoneda = MONEDA_SEC
    Set mobjBusq = New Busqueda
    Set mObjCond = New RepCondicion
    
    Me.Show
    Me.ZOrder
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub




Private Function Imprimir() As Boolean
    Dim s As String, tid As Long, i As Long, x As Single, res As String
    Dim gnc As GNComprobante, cambiado As Boolean, cntError As Long
    
    On Error GoTo ErrTrap

    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdBuscar.Enabled = False
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    
    'Limpia los mensajes
    For i = grd.FixedRows To grd.Rows - 1
        grd.TextMatrix(i, COL_RESULTADO) = ""
    Next i
    
    For i = grd.FixedRows To grd.Rows - 1
        DoEvents
        If mCancelado Then
            MsgBox "El proceso fue cancelado."
            Exit For
        End If
        
        prg1.value = i
        grd.Row = i
        x = grd.CellTop                 'Para visualizar la celda actual
        
        tid = grd.ValueMatrix(i, COL_TID)
        grd.TextMatrix(i, COL_RESULTADO) = "Procesando ..."
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
                            
            'Si la transaccion está anulado
            Else
                grd.TextMatrix(i, COL_RESULTADO) = "Anulado."
                cntError = cntError + 1
            End If
        Else
            grd.TextMatrix(i, COL_RESULTADO) = "No pudo recuperar la transación."
            cntError = cntError + 1
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


Private Sub cmdBuscar_Click()
    Dim v As Variant, obj As Object, sql As String
    On Error GoTo ErrTrap
    CargaPCProyCobroPago True, True
'    With gobjMain.objCondicion
'        .Fecha1 = dtpFecha1.value
'        .Fecha2 = dtpFecha2.value
'        .CodTrans = fcbFormaCobro.Text
'        .NumTrans1 = Val(txtNumTrans1.Text)
'        .NumTrans2 = Val(txtNumTrans2.Text)
'
'        'Estados no incluye anulados
'        .EstadoBool(ESTADO_NOAPROBADO) = True
'        .EstadoBool(ESTADO_APROBADO) = True
'        .EstadoBool(ESTADO_DESPACHADO) = True
'        .EstadoBool(ESTADO_ANULADO) = False
'    End With
'    If Me.Caption = "Busca Transacciones con problemas de Relación" Then
'        Set obj = gobjMain.EmpresaActual.ConsGNTransError()
'        If Not obj.EOF Then
'            v = MiGetRows(obj)
'
'            grd.Redraw = flexRDNone
'            grd.LoadArray v
'            ConfigColsTRansErradas
'            grd.Redraw = flexRDDirect
'        Else
'            grd.Rows = grd.FixedRows
'            ConfigColsTRansErradas
'        End If
'
'    Else
'        Set obj = gobjMain.EmpresaActual.ConsGNTrans2(True)
'        If Not obj.EOF Then
'            v = MiGetRows(obj)
'
'            grd.Redraw = flexRDNone
'            grd.LoadArray v
'            ConfigCols
'            grd.Redraw = flexRDDirect
'        Else
'            grd.Rows = grd.FixedRows
'            ConfigCols
'        End If
    
        cmdImprimir.Enabled = True
        cmdImprimir.SetFocus
'    End If
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
        .ColHidden(COL_NOMBRE) = False  'True
        .ColHidden(COL_DESC) = False
        .ColHidden(COL_CENTROCOSTO) = True
        .ColHidden(COL_ESTADO) = True
        
        .ColDataType(COL_FECHA) = flexDTDate    '*** MAKOTO 14/ago/2000 para que ordene bien por fecha
        
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
    Case vbKeyF9
        cmdImprimir_Click
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
    grd.Move 0, 100, Me.ScaleWidth, Me.ScaleHeight - grd.Top - pic1.Height + 200
    prg1.Width = Me.ScaleWidth - (prg1.Left * 2)
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
    
    If Not bandAsiento Then
        'Envia directamente a la impresora con el segundo parámetro 'True'
        objImp.PrintTrans gobjMain.EmpresaActual, True, 1, 0, "", 0, gc
    Else
        objImp.PrintAsiento gobjMain.EmpresaActual, True, 1, 0, "", 0, gc
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

Private Sub ConfigColsTRansErradas()
    With grd
        .FormatString = "^#|<Fecha|<Trans|<#|<Descripción"
        
        .ColDataType(1) = flexDTDate    '*** MAKOTO 14/ago/2000 para que ordene bien por fecha
        
        GNPoneNumFila grd, False
        '.AutoSize 0, grd.Cols - 1
        .ColWidth(1) = 1000
        .ColWidth(2) = 900
        .ColWidth(3) = 900
        .ColWidth(4) = 6000
    End With
End Sub

Private Sub CargaPCProyCobroPago(ByVal Busqueda As Boolean, bandCobrar As Boolean)
    Dim FechaCorte As Date
    Dim s1 As String, s2 As String, s3 As String
    Dim sql As String, sql1 As String, sql2 As String, sql3 As String, sql4 As String, sql5 As String, cond As String
    Dim Num1 As Integer, Num2 As Integer, Num3 As Integer
    On Error GoTo ErrTrap

    'Llamada a la pantalla de búsqueda (3)
    'mObjCond.Tipo = bsqPCProyPago          'Proveedores.ProyPago
    gobjMain.objCondicion.Tipo = bsqPCProyPago
    gobjMain.objCondicion.CodMoneda = mCodMoneda
    If Not mobjBusq.Show(gobjMain) Then Exit Sub
    mObjCond.Fcorte = gobjMain.objCondicion.FechaCorte
    mObjCond.Cliente1 = gobjMain.objCondicion.CodPC1
    mObjCond.Cliente2 = gobjMain.objCondicion.CodPC2
    Num1 = gobjMain.objCondicion.NumDias1
    Num2 = gobjMain.objCondicion.NumDias2
    Num3 = gobjMain.objCondicion.NumDias3
    If Len(gobjMain.objCondicion.CodPC1) > 0 And Len(gobjMain.objCondicion.CodPC2) > 0 Then
               cond = "CodProvCli BETWEEN  '" & gobjMain.objCondicion.CodPC1 & "' AND '" & gobjMain.objCondicion.CodPC2 & "' AND "
        End If
    If Len(gobjMain.objCondicion.CodBanco1) > 0 And Len(gobjMain.objCondicion.CodBanco1) > 0 Then
               cond = "Codforma BETWEEN  '" & gobjMain.objCondicion.CodBanco1 & "' AND '" & gobjMain.objCondicion.CodBanco2 & "' AND "
        End If
    
    
    s1 = "SELECT CodProvCli, Nombre, Fechatrans, Codtrans, Codtrans + CONVERT(varchar,Numtrans) AS trans, " & _
    "CodForma + NumLetra AS Doc, FechaEmision, FechaVenci, DateDiff(dd, FechaVenci, " & FechaYMD(mObjCond.Fcorte, gobjMain.EmpresaActual.TipoDB) & ") AS NumD, "

    s2 = "Case  " & mObjCond.NumMoneda & _
               " WHEN 1 THEN Saldo1 " & _
            "WHEN 2 THEN Saldo2 " & _
            "WHEN 3 THEN Saldo3 " & _
            "WHEN 4 THEN Saldo4  END "
    'aquí asignar la condición de búsqueda por proveedor  en cond
    sql1 = s1 & s2 & " as V1, null AS V2, null AS V3, null AS V4, null AS V5, Estado " & _
              "From vwConsPCProyCobroPago " & _
              "WHERE " & cond & _
              "(FechaVenci <=  " & FechaYMD(mObjCond.Fcorte, gobjMain.EmpresaActual.TipoDB) & ") AND (PorCobrar = " & IIf(bandCobrar = True, 1, 0) & ") " & _
              "AND (CASE " & mObjCond.NumMoneda & " WHEN 1 THEN Saldo1 WHEN 2 THEN Saldo2  " & _
              "WHEN 3 THEN Saldo3 WHEN 4 THEN Saldo4 END >0) "
    '"Union "
    sql2 = s1 & " null as v1," & s2 & " as V2, null AS V3, null AS V4, null AS V5, Estado " & _
            "From vwConsPCProyCobroPago " & _
            "WHERE " & cond & _
            "(FechaVenci BETWEEN DateAdd(dd, 1, " & FechaYMD(mObjCond.Fcorte, gobjMain.EmpresaActual.TipoDB) & ") AND DateAdd(dd, " & Num1 & ", " & FechaYMD(mObjCond.Fcorte, gobjMain.EmpresaActual.TipoDB) & "))" & _
           " AND (PorCobrar = " & IIf(bandCobrar = True, 1, 0) & ") " & _
            "AND (CASE 2 WHEN 1 THEN Saldo1 WHEN 2 THEN Saldo2 " & _
            "WHEN 3 THEN Saldo3 WHEN 4 THEN Saldo4 END >0) "
'Union
'/*-----------------------------------------------------------------
' * Fecha de vencimiento entre @NumD1 a @NumD2
' ------------------------------------------------------------------*/
    sql3 = s1 & " null AS V1, null as V2, " & s2 & _
            " AS V3, null AS V4, null AS V5, Estado " & _
            "From vwConsPCProyCobroPago " & _
            "WHERE (FechaVenci BETWEEN DateAdd(dd, " & Num1 & "+1, " & FechaYMD(mObjCond.Fcorte, gobjMain.EmpresaActual.TipoDB) & ") " & _
            "AND DateAdd(dd, " & Num2 & ", " & FechaYMD(mObjCond.Fcorte, gobjMain.EmpresaActual.TipoDB) & ")) " & _
            "AND (PorCobrar = " & IIf(bandCobrar = True, 1, 0) & ") " & _
            "AND (CASE 2 WHEN 1 THEN Saldo1 WHEN 2 THEN Saldo2 " & _
            "WHEN 3 THEN Saldo3 WHEN 4 THEN Saldo4 END >0) "
'Union
'/*-----------------------------------------------------------------
' * Fecha de vencimiento entre @NumD2 a @NumD3
' ------------------------------------------------------------------*/
    sql4 = s1 & "null AS V1, null AS V2, null as V3, " & s2 & _
            "AS V4, null AS V5, Estado " & _
            "From vwConsPCProyCobroPago " & _
            "WHERE " & cond & _
            "(FechaVenci BETWEEN DateAdd(dd, " & Num2 & "+1, " & FechaYMD(mObjCond.Fcorte, gobjMain.EmpresaActual.TipoDB) & ") " & _
            "AND DateAdd(dd, " & Num3 & ", " & FechaYMD(mObjCond.Fcorte, gobjMain.EmpresaActual.TipoDB) & ")) " & _
            "AND (PorCobrar = " & IIf(bandCobrar = True, 1, 0) & ") " & _
            "AND (CASE 2 WHEN 1 THEN Saldo1 WHEN 2 THEN Saldo2 " & _
            "WHEN 3 THEN Saldo3 WHEN 4 THEN Saldo4 END >0) "
'Union
'/*-----------------------------------------------------------------
' * Fecha de vencimiento mayores a @NumD3
' ------------------------------------------------------------------*/
    sql5 = s1 & " null AS V1, null AS V2, null AS V3, null as V4, " & s2 & _
        "AS V5, Estado " & _
        "From vwConsPCProyCobroPago " & _
        "WHERE " & cond & " (FechaVenci > DateAdd(dd, " & Num3 & ", " & FechaYMD(mObjCond.Fcorte, gobjMain.EmpresaActual.TipoDB) & ")) " & _
        "AND (PorCobrar = " & IIf(bandCobrar = True, 1, 0) & ") " & _
        "AND (CASE 2 WHEN 1 THEN Saldo1 WHEN 2 THEN Saldo2 " & _
        "WHEN 3 THEN Saldo3 WHEN 4 THEN Saldo4 END >0)"
    sql = sql1 & " UNION " & sql2 & " UNION " & sql3 & " UNION " & sql4 & " UNION " & sql5
    'cambiado AUC 13/07/2005 ------
        sql = sql & " ORDER BY CodProvCli, FechaVenci "
    '------------------------------
    grd.Redraw = False
    MensajeStatus MSG_PREPARA, vbHourglass
    MiGetRowsRep gobjMain.EmpresaActual.OpenRecordset(sql), grd
    grd.Redraw = True
    grd.Refresh
    Exit Sub
    
ErrTrap:
    Err.Raise Err.Number, , Err.Description
End Sub


Public Sub MiGetRowsRep(ByVal rs As Recordset, grd As VSFlexGrid)
    grd.LoadArray MiGetRows(rs)
End Sub

