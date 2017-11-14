VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCostosRealesItems 
   Caption         =   "Costos Reales x Items"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   6240
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6240
      TabIndex        =   1
      Top             =   4185
      Width           =   6240
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   180
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin MSComctlLib.ImageList img1 
      Left            =   5520
      Top             =   120
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CostosRealesItem.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CostosRealesItem.frx":0114
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CostosRealesItem.frx":0568
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CostosRealesItem.frx":067C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CostosRealesItem.frx":0790
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   1005
      ButtonWidth     =   1191
      ButtonHeight    =   953
      Style           =   1
      ImageList       =   "img1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar"
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar (F5)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Asignar"
            Key             =   "Asignar"
            Description     =   "Asignar un valor"
            Object.ToolTipText     =   "Asignar un valor (F6)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar"
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar (F3)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Imprimir"
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir (Ctrl+P)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cerrar"
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar (ESC)"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   7935
      _cx             =   13991
      _cy             =   4043
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   200
      ColWidthMax     =   6000
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
      MergeCells      =   3
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
Attribute VB_Name = "frmCostosRealesItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mProcesando As Boolean
Private mCancelado As Boolean
Private mcolItemsSelec As Collection      'Coleccion de items
'jeaa 24/09/04 asignacion de grupo a los items
'IVINVENTARIO
Const COL_TOTAL = 6
Const COL_COMI = 14
Private objcond As Condicion
Private COL_SUBTOTAL As Integer


'Private mobjItem As IVinventario

Public Sub Inicio(ByVal tag As String)
    Dim i As Integer
    On Error GoTo ErrTrap
    
    Me.tag = tag            'Guarda en la propiedad Tag para distinguir después
    Me.Show
    Me.ZOrder
    
    Select Case Me.tag
    Case "COMI"
        Me.Caption = "Actualización de Comisiones y Penalizaciones de Vendedores"
    End Select
       
    'Inicializa la grilla
    grd.Rows = grd.FixedRows
    ConfigCols
    
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF3
        Grabar
        KeyCode = 0
    Case vbKeyF5
        Select Case Me.tag
            Case "COMI": Buscar
        End Select
        KeyCode = 0
    Case vbKeyF6
        Asignar
        KeyCode = 0
    Case vbKeyP
        If Shift And vbCtrlMask Then
            Imprimir
            KeyCode = 0
        End If
    Case vbKeyEscape
        Cerrar
    Case Else
        MoverCampo Me, KeyCode, Shift, True
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ImpideSonidoEnter Me, KeyAscii
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mProcesando Then
        Cancel = 1      'No permitir cerrar mientras procesa
    Else
        Me.Hide         'Se pone esto para evitar el posible BUG de Windows98
    End If
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    grd.Move 0, tlb1.Height, Me.ScaleWidth, Me.ScaleHeight - tlb1.Height - pic1.Height - 80
    prg1.Width = Me.ScaleWidth - (prg1.Left * 2)
End Sub


Private Sub grd_BeforeEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
    If Row < grd.FixedRows Then Cancel = True
    If grd.IsSubtotal(Row) = True Then Cancel = True
    If grd.ColData(col) < 0 Then Cancel = True
    
    If Not Cancel Then
        'Longitud maxima para editar
        grd.EditMaxLength = grd.ColData(col)
    End If
End Sub

Private Sub grd_BeforeSort(ByVal col As Long, Order As Integer)
    'Impide mientras está procesando
    If mProcesando Then Order = flexSortNone
End Sub

Private Sub grd_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
    With grd
        Select Case Me.tag
        Case "COMI"
            If Not IsNumeric(.EditText) Then
                MsgBox "Debe ingresar un valor numérico.", vbInformation
                .SetFocus
                Cancel = True
            End If
        End Select
    End With
End Sub

Private Sub tlb1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Buscar":
        Select Case Me.tag
            Case "COMI": Buscar
        End Select
    Case "Asignar":     Asignar
    Case "Grabar":
        Select Case Me.tag
            Case "COMI": Grabar
        End Select
    Case "Imprimir":    Imprimir
    Case "Cerrar":      Cerrar
    End Select
End Sub

Private Sub Buscar()
    Static coditem As String, CodAlt As String, _
           Desc As String, _
           codg As String, Numg As Integer, bandIVA As Boolean, bandFraccion As Boolean
    Dim sql As String, cond As String, rs As Recordset, comodin As String
    Dim Recargo As String, v  As Variant, max As Integer, i As Integer
    Dim NumReg As Long, CadenaValores As String, CadenaAgrupa  As String
    Dim from As String
    On Error GoTo ErrTrap
    'If Me.tag <> "CUENTA" Then Exit Sub
    
    #If DAOLIB Then
        comodin = "*"
    #Else
        comodin = "%"
    #End If
'    comodin = "%"
    'Abre la pantalla de búsqueda
    Set objcond = gobjMain.objCondicion
    If Not frmB_ComPenVen.InicioComPenVen(objcond) Then
        grd.SetFocus
        Exit Sub
    End If
    'Cambia la forma de cursor
    MensajeStatus MSG_PREPARA, vbHourglass
    
    With objcond
        'Crea  tablas temporales
        v = Split(.Servicios, ",")
        max = UBound(v, 1)
        'Maximo  6 columnas de  recargo/descuento
        If max > 6 Then max = 6
        For i = 0 To max
            If v(i) <> "SUBT" Then
                VerificaExistenciaTabla i
                sql = "SELECT IVKardexRecargo.TransID, SUM(IVKardexRecargo.Valor)  AS Valor " & _
                      "Into Tmp" & i & "  FROM IVKardexRecargo INNER JOIN  IVRecargo ON " & _
                      "IVKardexRecargo.IdRecargo = IVRecargo.IdRecargo " & _
                      "WHERE ((IVRecargo.CodRecargo) = '" & v(i) & "')  GROUP BY IVKardexRecargo.TransID   "
                gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                'Debug.Print numreg
            End If
        Next i
        cond = " (GNComprobante.FechaTrans  BETWEEN " & _
                FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
        If Len(.CodCentro1) > 0 Or Len(.CodCentro2) > 0 Then
               cond = cond & " AND FCVendedor.CodVendedor BETWEEN  '" & .CodCentro1 & "' AND '" & .CodCentro2 & "' "
        End If

        If Len(.CodTrans) > 0 Then
           cond = cond & " AND GNComprobante.CodTrans IN (" & PreparaCadena(.CodTrans) & ")"
        End If
       '***Agregado. 14/03/2003. Angel
       'Condición de Estado
        If .EstadoBool(ESTADO_NOAPROBADO) = False Then
            If Len(cond) > 0 Then cond = cond & " AND "
            cond = cond & "(GNComprobante.Estado<>" & ESTADO_NOAPROBADO & ")"
        End If
        If .EstadoBool(ESTADO_APROBADO) = False Then
            If Len(cond) > 0 Then cond = cond & " AND "
            cond = cond & "(GNComprobante.Estado<>" & ESTADO_APROBADO & ")"
        End If
        If .EstadoBool(ESTADO_DESPACHADO) = False Then
            If Len(cond) > 0 Then cond = cond & " AND "
            cond = cond & "(GNComprobante.Estado<>" & ESTADO_DESPACHADO & ")"
        End If
        If .EstadoBool(ESTADO_ANULADO) = False Then
            If Len(cond) > 0 Then cond = cond & " AND "
            cond = cond & "(GNComprobante.Estado<>" & ESTADO_ANULADO & ")"
        End If
        'Carga valores  de moneda Indice Moneda 0 / 1 / 2 / 3
'        CadenaValores = "Sum(IVK.Precio0" & _
'              IIf(.NumMoneda > 0, "/Cotizacion" & .NumMoneda + 1, "") & ") As Valor0, "
'        CadenaValores = CadenaValores & "Sum(IVK.Precio12" & _
'              IIf(.NumMoneda > 0, "/Cotizacion" & .NumMoneda + 1, "") & ") As Valor12, "
'        'Descuento x Items sobre items sin IVA
'        CadenaValores = CadenaValores & "Sum(IVK.DescxItem0" & _
'              IIf(.NumMoneda > 0, "/Cotizacion" & .NumMoneda + 1, "") & ") As DescxItem, "
'        'Descuento x Items sobre items con IVA
'        CadenaValores = CadenaValores & "Sum(IVK.DescxItemIVA" & _
'              IIf(.NumMoneda > 0, "/Cotizacion" & .NumMoneda + 1, "") & ") As DescxItem, "
        ' precio total no real
        CadenaValores = CadenaValores & "Sum(IVK.PrecioTotal" & _
              IIf(.NumMoneda > 0, "/Cotizacion" & .NumMoneda + 1, "") & ") * -1 As ValorNeto,  "
                'ciclo  para  cargar  los descuentos maximo 4  descuentos/recargos
        For i = 0 To max
            If v(i) = "SUBT" Then
                CadenaValores = CadenaValores & "0 As SubTotal, "
            Else
                CadenaValores = CadenaValores & "(tmp" & i & ".valor" & _
                IIf(.NumMoneda > 0, "/Cotizacion" & .NumMoneda + 1, "") & ") As " & v(i) & "1, "
            End If
        Next i
        For i = max + 1 To 6
           CadenaValores = CadenaValores & "0 As Rec" & i & ", "
        Next i
'        ' costo real por total
        CadenaValores = CadenaValores & " 0 AS Total "


        'Cadena Agrupa
        For i = 0 To max
            If v(i) <> "SUBT" Then
                CadenaAgrupa = CadenaAgrupa & "(tmp" & i & ".valor" & _
                IIf(.NumMoneda > 0, "/Cotizacion" & .NumMoneda + 1, "") & "), "
            End If
        Next i
        'quita la ultima  coma
        If Len(CadenaAgrupa) > 2 Then           '*** MAKOTO 19/feb/01 Mod.
            CadenaAgrupa = Left(CadenaAgrupa, Len(CadenaAgrupa) - 2)
        End If


        
        sql = "SELECT FCVendedor.Nombre, " & _
              "PCProvCli.Nombre,  GNComprobante.transID,CodTrans + ' ' + CONVERT(varchar,NumTrans) AS Trans, GNComprobante.FechaTrans,  " & _
              CadenaValores & ", 0 AS Comision, ' ' as Resul "
        from = ""
        For i = 0 To max
            If v(i) <> "SUBT" Then from = from & "("
        Next i
        

        
        If Len(from) > 0 Then
            from = from & " GNComprobante "
            For i = 0 To max
                If v(i) <> "SUBT" Then
                    from = from & " LEFT JOIN Tmp" & i & " ON GNComprobante.TransID = Tmp" & i & ".TransID) "
                End If
            Next i
        End If
        sql = sql & "FROM FCVendedor INNER JOIN (( " & IIf(Len(from), from, " GNComprobante ")
        sql = sql & " LEFT JOIN PCProvCli ON GNComprobante.IdClienteRef = PCProvCli.IdProvCli) " & _
              "INNER JOIN vwConsIVKardexIVA IVK ON GNComprobante.TransID = IVK.TransID) ON " & _
              "FCVendedor.IdVendedor = GNComprobante.IdVendedor "
        sql = sql & "WHERE  " & cond & _
              " GROUP BY FCVendedor.Nombre, " & _
              "GNComprobante.FechaTrans, CodTrans + ' ' + CONVERT(varchar,NumTrans), " & _
              "GNComprobante.NumDocRef, PCProvCli.Nombre, GNComprobante.TransID "
        '*** MAKOTO 19/feb/01 Mod.
        If Len(CadenaAgrupa) > 0 Then sql = sql & ", " & CadenaAgrupa
    sql = sql & " ORDER BY FCVendedor.Nombre, FechaTrans"
    End With
    '------------------------------
    grd.Redraw = False
    MensajeStatus MSG_PREPARA, vbHourglass
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
   
    With grd
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If Not rs.EOF Then .LoadArray MiGetRows(rs)
        ConfigCols
        .Redraw = flexRDBuffered
        .SetFocus
    End With
    
    MensajeStatus
    Exit Sub
ErrTrap:
    grd.Redraw = flexRDBuffered
    MensajeStatus
    DispErr
    grd.SetFocus
    Exit Sub
End Sub

Private Sub ConfigCols()
    Dim s As String, i As Long, j As Integer
    Dim fmt As String, max As Integer, v As Variant
    With grd
    Select Case Me.tag
        Case "COMI"
            s = "^#|<Vendedor|<Cliente|<TransID|<Transacción|<Fecha Emision|>Venta Neta|>rec0|>rec1|>rec2|>rec3|>rec4|>rec5|>rec6|>Comisión|>Resultado"
        End Select
        .FormatString = s
        fmt = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda("USD")
        grd.ColFormat(6) = fmt
        grd.ColFormat(7) = fmt
        grd.ColFormat(8) = fmt
        grd.ColFormat(9) = fmt
        grd.ColFormat(10) = fmt
        grd.ColFormat(11) = fmt
        grd.ColFormat(12) = fmt
        
        grd.ColFormat(COL_COMI) = "#0.00"
        If Not objcond Is Nothing Then
            v = Split(objcond.Servicios, ",")
            max = UBound(v, 1)
            s = "^#|<Vendedor|<Cliente|<TransID|<Transacción|<Fecha Emision|>Venta Neta"
            For i = 0 To max
                s = s & "|>" & v(i)
                grd.ColHidden(i + 7) = False
            Next i
            For i = max + 1 To 6
               s = s & "|>Rec" & i
               grd.ColHidden(i + 7) = True
            Next i
            .FormatString = s
            CalculaSubtotal
        End If
        grd.ColHidden(COL_COMI) = False
        grd.ColHidden(15) = False
        
        GNPoneNumFila grd, False
        AjustarAutoSize grd, -1, -1, 4000
        AsignarTituloAColKey grd
    
        'Columnas modificables (Longitud maxima)
        Select Case Me.tag
        End Select
        'Columnas No modificables
        For i = 0 To .ColIndex("Venta Neta")
            .ColData(i) = -1
        Next i
        
        .MergeCol(1) = True
        If .Rows > .FixedRows Then
            .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("Venta Neta")) = .BackColorFrozen
        End If
        grd.SubTotal flexSTClear
        grd.SubTotal flexSTSum, 1, COL_TOTAL, , grd.GridColor, vbBlack, , "Subtotal", 1, True
        grd.SubTotal flexSTSum, 1, COL_SUBTOTAL, , grd.GridColor, vbBlack, , "Subtotal", 1, True
        grd.SubTotal flexSTSum, -1, COL_TOTAL, , grd.BackColorSel, vbYellow, , "Total", 1, True
        grd.SubTotal flexSTSum, -1, COL_SUBTOTAL, , grd.BackColorSel, vbYellow, , "Total", 1, True
    End With
End Sub


Private Sub Asignar()
    Select Case Me.tag
        Case "COMI":        AsignarComi
    End Select
End Sub


Private Sub Grabar()
    Dim i As Long, gn As GNComprobante, cod As Long, sql As String, NumReg As Long
    On Error GoTo ErrTrap
    
    'Confirmación
    If MsgBox("Está seguro que desea grabar?", vbQuestion + vbYesNo) <> vbYes Then
        grd.SetFocus
        Exit Sub
    End If
    
    'Deshabilita los botónes y menus
    Habilitar False
    mCancelado = False
    
    With grd
        prg1.min = 0
        prg1.max = 1
        If .Rows > .FixedRows Then prg1.max = .Rows - 1
        For i = .FixedRows To .Rows - 1
            'Si es que se canceló el proceso
            If mCancelado Then GoTo salida
            prg1.value = i
            If Not grd.IsSubtotal(i) Then
                cod = .TextMatrix(i, .ColIndex("TransID"))
                MensajeStatus i & " de " & .Rows - .FixedRows, vbHourglass
                DoEvents
                Select Case Me.tag
                Case "COMI"
                        sql = "update gncomprobante Set Comision = " & .ValueMatrix(i, COL_COMI)
                        sql = sql & " where transid=" & cod
                        gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                        .TextMatrix(i, 15) = "Ok"
                End Select
            End If
        Next i
    End With
    
salida:
    MensajeStatus
    Set gn = Nothing
    Habilitar True
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    GoTo salida
    Exit Sub
End Sub


Private Sub Habilitar(ByVal v As Boolean)
    mProcesando = Not v
    
    tlb1.Buttons("Buscar").Enabled = v
    tlb1.Buttons("Asignar").Enabled = v
    tlb1.Buttons("Grabar").Enabled = v
    tlb1.Buttons("Imprimir").Enabled = False        '*** MAKOTO PENDIENTE Por ahora
    
    If v Then
        tlb1.Buttons("Cerrar").Caption = "Cerrar"
    Else
        tlb1.Buttons("Cerrar").Caption = "Cancelar"
    End If
    
    frmMain.mnuFile.Enabled = v
    frmMain.mnuHerramienta.Enabled = v
    frmMain.mnuTransferir.Enabled = v
    frmMain.mnuCerrarTodas.Enabled = v
    
    prg1.value = prg1.min
End Sub

Private Sub Imprimir()

End Sub

Private Sub Cerrar()
    If mProcesando Then
        'Si está procesando, pregunta si quere abandonarlo
        If MsgBox("Desea abandonar el proceso?", vbQuestion + vbYesNo) = vbYes Then
            mCancelado = True
        End If
        
        Exit Sub
    Else
        Unload Me
    End If
End Sub


Private Sub AsignarComi()
Dim i As Integer, j As Integer, k As Integer, m As Integer
Dim band As Boolean
    band = True
    m = 1
    For i = 1 To grd.Rows - 1
        If grd.IsSubtotal(i) = True Then
            For j = 1 To 10
                If grd.TextMatrix(i, COL_TOTAL) >= gComisiones(j).desde And grd.TextMatrix(i, COL_TOTAL) <= gComisiones(j).hasta Then
                    For k = m To i - 1
                        grd.TextMatrix(k, COL_COMI) = gComisiones(j).Comision
                    Next k
                    m = k + 1
                    j = 10
                End If
            Next j
        End If
    Next i
End Sub


Private Sub CalculaSubtotal()
    Dim i As Integer, max As Integer, v As Variant
    Dim col As Integer, j As Integer, tot As Currency
        If Not objcond Is Nothing Then
            v = Split(objcond.Servicios, ",")
            max = UBound(v, 1)
            For i = 0 To max
                If v(i) = "SUBT" Then
                    COL_SUBTOTAL = 7 + i
                End If
            Next i
        End If
    
    
    For i = 1 To grd.Rows - 1
            tot = 0
            If Not grd.IsSubtotal(i) = True Then
                For j = 0 To max
                    If Len(grd.TextMatrix(i, 6 + j)) > 0 Then
                        tot = tot + grd.TextMatrix(i, 6 + j)
                    End If
                Next j
                grd.TextMatrix(i, COL_SUBTOTAL) = tot
            End If
    Next i
End Sub
