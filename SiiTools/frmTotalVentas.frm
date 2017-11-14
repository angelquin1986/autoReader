VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl1.ocx"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Begin VB.Form frmTotalVentas 
   Caption         =   "Actualizar Total  de Ventas"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7575
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4965
   ScaleWidth      =   7575
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar tlb1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "img1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar - F3"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar - F5"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exportar"
            Object.ToolTipText     =   "Excel"
            ImageIndex      =   3
         EndProperty
      EndProperty
      Begin VB.PictureBox picSucursal 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   5400
         ScaleHeight     =   375
         ScaleWidth      =   6855
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   6855
         Begin VB.CheckBox chkMostrarCostos 
            BackColor       =   &H80000018&
            Caption         =   "Mostrar Costos"
            Height          =   255
            Left            =   3900
            TabIndex        =   7
            Top             =   60
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CommandButton cmdCargar 
            Caption         =   "Cargar"
            Height          =   360
            Left            =   5460
            TabIndex        =   6
            Top             =   0
            Width           =   1095
         End
         Begin VB.CheckBox chkTodo 
            BackColor       =   &H80000018&
            Caption         =   "Todo"
            Height          =   255
            Left            =   2880
            TabIndex        =   5
            Top             =   60
            Value           =   1  'Checked
            Width           =   735
         End
         Begin FlexComboProy.FlexCombo fcbSucursal 
            Height          =   375
            Left            =   960
            TabIndex        =   3
            Top             =   0
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   661
            ColWidth1       =   3400
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
         Begin VB.Label Label1 
            Caption         =   "Sucursal"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   60
            Width           =   1455
         End
      End
   End
   Begin MSComctlLib.ImageList img1 
      Left            =   4320
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTotalVentas.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTotalVentas.frx":0120
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTotalVentas.frx":0240
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   5055
      _cx             =   8911
      _cy             =   3619
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
      Rows            =   3
      Cols            =   5
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
      AllowUserFreezing=   0
      BackColorFrozen =   12648447
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "frmTotalVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const COL_NUM = 0
Const COL_COD = 1
Const COL_NOM = 2
Const COL_TOT = 3
Const COL_RES = 4
Const COL_GRUPO = 4

Const COL_IDIV = 1
Const COL_CODIV = 2
Const COL_IDPROV = 3
Const COL_CODPROV = 4
Const COL_TRANS = 5
Const COL_CANT = 6
Const COL_COSTO = 7
Const COL_FECHA = 8
Const COL_RESUL = 9

Const COL_F_IDIV = 1
Const COL_F_CODIV = 2
Const COL_F_DESCIV = 3
Const COL_F_IDBOD = 4
Const COL_F_CODBOD = 5
Const COL_F_DESCBOD = 6
Const COL_F_FECHA = 7
Const COL_F_FECHAN = 8
Const COL_F_RESUL = 9

Const COL_DCI_IDCLI = 1
Const COL_DCI_RUC = 2
Const COL_DCI_NOMCLI = 3
Const COL_DCI_IDIV = 4
Const COL_DCI_CODIV = 5
Const COL_DCI_DESCIV = 6
Const COL_DCI_DESC = 7
Const COL_DCI_RESUL = 8

Const COL_MV_GRUPO = 4
Const COL_MV_NEWGRUPO = 5
Const COL_MV_RES = 6
Private NumPCGrupo As Integer
Private mObjCond As RepCondicion
Private BandTodo As Boolean

Public Sub Inicio(ByVal Cadena As String)
    Me.tag = Cadena
    Me.Caption = "Actualizar Total  de Ventas"
    CargaDatos
    ConfigCols
    Me.Show
    Me.ZOrder
End Sub

Private Sub CargaDatos()
    Dim sql As String, cond As String, rs As Recordset, antes As Long
    Dim objcond As Condicion
    Static Recargo As String
    
    On Error GoTo ErrTrap
    
    antes = grd.Row
    grd.Rows = 1
    Set objcond = gobjMain.objCondicion
    If Not (frmB_VxTrans.InicioVxTransaccion(objcond, Recargo, "TotalVentas")) Then
        grd.SetFocus
        Exit Sub
    End If
       
    grd.Redraw = False
    MensajeStatus MSG_PREPARA, vbHourglass
    
    With objcond
        VerificaExistenciaTabla 0
        cond = " AND gc.FechaTrans <=" & FechaYMD(.FechaCorte, gobjMain.EmpresaActual.TipoDB)
        
        If Len(.CodTrans) > 0 Then
           cond = cond & " AND GC.CodTrans IN (" & PreparaCadena(.CodTrans) & ")"
        End If
        
        sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                "From IvRecargo ivR inner join " & _
                    "IvKardexRecargo ivkR Inner join " & _
                        "GnComprobante gc Inner join PcPRovCLi on gc.IdClienteRef = PCProvCli.IdProvCli " & _
                    "On ivkr.TransID = gc.TransID " & _
                "On Ivr.IdRecargo = IvkR.IdRecargo "
        sql = sql & "WHERE gc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(Recargo) & ") " & cond & _
              "Group by IvkR.TransID"
        
        gobjMain.EmpresaActual.EjecutarSQL sql, 0
    
        sql = "SELECT PCProvCli.CodProvCli, PCProvCli.Nombre, " & _
              "ABS(SUM((PrecioTotalBase0 + (PrecioTotalBase0 * (cast(TotalDescuento as float) / cast(PrecioTotal as float))))*SignoVenta) + " & _
              "SUM((PrecioTotalBaseIVA + (PrecioTotalBaseIVA * (cast(TotalDescuento as float) / cast(PrecioTotal as float))))*SignoVenta)) As TotalVenta, " & _
              "SPACE(0) AS Resultado "

        sql = sql & "FROM tmp0 inner join " & _
                    "vwConsSUMIVKardexIVA inner join " & _
                        "GNComprobante GC  Inner JOIN PCProvCli ON GC.IdClienteRef = PCProvCli.IdProvCli " & _
                        "ON vwConsSUMIVKardexIVA.TransID = GC.TransID " & _
                    "ON tmp0.TransID = GC.TransID "
        'jeaa 30/09/04 agegado AND PRECIOTOTAL<>0 para controlar error division para cero
        sql = sql & " WHERE (gc.Estado<>3) AND PRECIOTOTAL<>0" & cond & _
                    " GROUP BY PCProvCli.CodProvCli, PCProvCli.Nombre, PCProvCli.TotalDebe "
        
        'Para filtrar solo los clientes cuyo monto de venta se ha modificado,
        'es decir solo los que han comprado
        sql = sql & "HAVING round(ABS(SUM((PrecioTotalBase0 + (PrecioTotalBase0 * " & _
                    "(cast(TotalDescuento as float) / cast(PrecioTotal as float))))*SignoVenta) + " & _
                    "SUM((PrecioTotalBaseIVA + (PrecioTotalBaseIVA * (cast(TotalDescuento as float)/" & _
                    "cast(PrecioTotal as float))))*SignoVenta)),4)<>PCProvCli.TotalDebe "
                    
        sql = sql & "ORDER BY PCProvCli.Nombre"
    End With
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    grd.LoadArray MiGetRows(rs)
    Set rs = Nothing
    
    With grd
        '# de fila
        .ColAlignment(0) = flexAlignCenterCenter
        GNPoneNumFila grd, False
        'Reubica a la fila donde estaba antes
        If .Rows > antes And antes > 0 Then .Row = antes
        .Redraw = True
    End With
    MensajeStatus "", 0
    grd.SetFocus
    If grd.Rows <> grd.FixedRows Then grd.Row = grd.FixedRows
    Exit Sub

ErrTrap:
    grd.Redraw = True
    MensajeStatus "", 0
    DispErr
    Exit Sub
End Sub

Private Sub chkTodo_Click()
    If chkTodo.value = vbChecked Then
        fcbSucursal.KeyText = ""
        fcbSucursal.Enabled = False
    Else
        fcbSucursal.Enabled = True
    End If
End Sub

Private Sub cmdCargar_Click()
    BandTodo = True
    CargaVentasxItemxSucursal
    ConfigColsVtaxItemxSuc
    BandTodo = False
End Sub

Private Sub Form_Initialize()
    Set mObjCond = New RepCondicion
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF3
        Grabar
        KeyCode = 0
    Case vbKeyF5 'Buscar
            Select Case Me.tag
            Case "VxTrans"
                CargaDatos
                ConfigCols
            Case "CxTrans"
                CargaDatosComprasProveedor
                ConfigColsComprasProveedor
            Case "VxTransProm"
                CargaDatosPromedio
                ConfigCols
            Case "CustoxActivo"
                CargaDatosCustodioActivos
                ConfigColsCustodioxActivo
            Case "PCGxMontoVenta"
                CargaDatosMontoVentas
                ConfigCols
            Case "CalculoBuffer"
                CargaPromedioVentasDiariaBufferMP3
                ConfigColsPromedioVentasDiariaBufferMP3
            Case "CalculoBufferUtilesa"
                CargaPromedioVentasDiariaBufferUti
                ConfigColsPromedioVentasDiariaBufferUti
            Case "CalculoBufferxAlmacen"
                CargaPromedioVentasDiariaBufferMP3xAlma
                ConfigColsPromedioVentasDiariaBufferMP3Alma
            Case "DescxClixItem"
                CargaVentasxItemxCliente
                ConfigColsDescxClixItem
            End Select
        'CargaDatos
        KeyCode = 0
    Case vbKeyEscape
        Unload Me
        KeyCode = 0
    End Select
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        grd.Move 0, tlb1.Height, Me.ScaleWidth, Me.ScaleHeight - tlb1.Height
    End If
End Sub

Private Function PreparaCadena(ByVal Cadena As String) As String
'Funcion que concatena apostrofes en una cadena separada por comas
Dim v As Variant, max As Integer, i As Integer
Dim Respuesta As String
    If Cadena = "" Then
        PreparaCadena = "''"
        Exit Function
    End If
    v = Split(Cadena, ",")
    max = UBound(v, 1)
    For i = 0 To max
        Respuesta = Respuesta & "'" & v(i) & "'" & ","
    Next i
    Respuesta = Left(Respuesta, Len(Respuesta) - 1) 'Quita la útima coma
    PreparaCadena = Respuesta
End Function

Private Sub ConfigCols()
    Dim fmt As String
    Dim i As Integer
    fmt = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda("USD")
    With grd
        Select Case Me.tag
        Case "VxTrans", "CxTrans"
            .FormatString = ">#|<Código|<Nombre|>Venta Total|<Resultado"
        Case "VxTransProm"
            .FormatString = ">#|<Código|<Nombre|>Límite de Crédito|<Resultado"
        Case "PCGxMontoVenta"
            .FormatString = ">#|<Código|<Nombre|>Monto de Ventas |<Grupo|<Nuevo Grupo|<Resultado"
        Case "PCGxMontoVentaCobro"
            .FormatString = ">#|<Código|<Nombre|>Monto de Ventas|>FechaUltCobro|>PromedioDias |<Grupo|<Nuevo Grupo|^Cambiar|<Resultado"
        Case "DescxClixItem"
            .FormatString = ">#<|<idcliente|<Nombre|<IdItem|<Codigo Item|<Descripcion Item|>Descuento |<Grupo|<Nuevo Grupo|<Resultado"
            
        End Select
        
        .ColWidth(COL_NUM) = 700
        .ColWidth(COL_COD) = 1500
        .ColWidth(COL_NOM) = 3500
        .ColWidth(COL_TOT) = 1500
        .ColWidth(COL_RES) = 1500
        .ColWidth(COL_MV_GRUPO) = 1500
        
        If Me.tag = "PCGxMontoVenta" Then
            .ColWidth(COL_MV_NEWGRUPO) = 1500
            .ColWidth(COL_MV_RES) = 1500
            .ColDataType(COL_MV_RES) = flexDTString
            .ColDataType(COL_MV_RES) = flexDTString
            .ColDataType(COL_MV_NEWGRUPO) = flexDTString
        End If
        
        AsignarTituloAColKey grd
       ' If Me.tag = "VxTransProm" Then
        
        If Me.tag = "PCGxMontoVentaCobro" Then
            grd.ColDataType(8) = flexDTBoolean
        '    grd.Enabled = True
            grd.Editable = flexEDKbdMouse
            For i = 0 To .Cols - 3
                .ColData(i) = -1
            Next i
            'Color de fondo
            If .Rows > .FixedRows Then
                .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 3) = .BackColorFrozen '.ColIndex("Límite de Crédito")) = .BackColorFrozen
            End If
        Else
            grd.Editable = flexEDNone
            For i = 0 To .Cols - 2 '.ColIndex("Límite de Crédito") '- 1
                .ColData(i) = -1
            Next i
            'Color de fondo
            If .Rows > .FixedRows Then
                .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 2) = .BackColorFrozen '.ColIndex("Límite de Crédito")) = .BackColorFrozen
            End If
        End If
        
        .ColDataType(COL_NUM) = flexDTLong
        .ColDataType(COL_COD) = flexDTString
        .ColDataType(COL_NOM) = flexDTString
        .ColDataType(COL_TOT) = flexDTCurrency
        .ColDataType(COL_RES) = flexDTString
        .ColFormat(COL_TOT) = fmt
        If Me.tag = "PCGxMontoVentaCobro" Then
            .ColDataType(8) = flexDTBoolean
        End If
    End With
End Sub

Private Sub Grabar()
    Dim sql As String, cod As String, i As Long
    Dim NumReg As Long, totalventa As Currency
    On Error GoTo ErrTrap
    MensajeStatus "Guardando....", 1
    With grd
        If .Rows = .FixedRows Then Exit Sub
        .ShowCell 1, 1
        For i = .FixedRows To .Rows - 1
            If Not .IsSubtotal(i) Then
                .Row = i
                .ShowCell i, 1           'Hace visible la fila actual
                cod = .TextMatrix(i, COL_COD)
                totalventa = .ValueMatrix(i, COL_TOT)
                sql = "UPDATE PCProvCli " & _
                      "SET TotalDebe = " & totalventa & _
                      "WHERE CodProvCli= '" & cod & "'"
                gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                If NumReg > 0 Then
                    .TextMatrix(i, COL_RES) = "Actualizado..."
                Else
                    .TextMatrix(i, COL_RES) = "Error al tratar de Actualizar..."
                End If
                .Redraw = True
                .Refresh
            End If
        Next i
    End With
    MensajeStatus "", 0
    Exit Sub
ErrTrap:
    MsgBox Err.Description, vbExclamation + vbOKOnly
    Exit Sub
End Sub

Private Sub Form_Terminate()
    Set mObjCond = Nothing
End Sub

Private Sub tlb1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Grabar"
        Select Case Me.tag
            Case "VxTrans"
                Grabar
            Case "CxTrans"
                GrabarComprasProveedor
            Case "VxTransProm"
                GrabarPromedio
            Case "CustoxActivo"
                GrabarCustodioxActivo
            Case "PCGxMontoVenta"
                GrabarPCGrupoxMontoVenta
            Case "FechaUltimoEgreso"
                GrabarFechaUltimoEgreso
            Case "FechaUltimoIngreso"
                GrabarFechaUltimoIngreso
            Case "CalculoBuffer"
                GrabarBuffer
            Case "CalculoBufferUtilesa"
                GrabarBufferUtilesa
            Case "CalculoBufferxAlmacen"
                GrabarBufferxALM
            Case "DescxClixItem"
                GrabarDescuentoxClixItem
            Case "VxIxSuc"
                GrabaClasificaxItem
            Case "PCGxMontoVentaCobro"
                GrabarPCGrupoxMontoVentaCobro
            End Select
    Case "Buscar"
        Select Case Me.tag
            Case "VxTrans"
                CargaDatos
                ConfigCols
            Case "CxTrans"
                CargaDatosComprasProveedor
                ConfigColsComprasProveedor
            Case "VxTransProm"
                CargaDatosPromedio
                ConfigCols
            Case "CustoxActivo"
                CargaDatosCustodioActivos
                ConfigColsCustodioxActivo
            Case "PCGxMontoVenta"
                CargaDatosMontoVentas
                ConfigCols
            Case "FechaUltimoEgreso"
                CargaDatosFechaUltimoEgreso
                ConfigColsFechaUltimoEgreso
            Case "FechaUltimoIngreso"
                CargaDatosFechaUltimoIngreso
                ConfigColsFechaUltimoIngreso
            Case "CalculoBuffer"
                CargaPromedioVentasDiariaBufferMP3
                ConfigColsPromedioVentasDiariaBufferMP3
            Case "CalculoBufferUtilesa"
                CargaPromedioVentasDiariaBufferUti
                ConfigColsPromedioVentasDiariaBufferUti
            Case "CalculoBufferxAlmacen"
                CargaPromedioVentasDiariaBufferMP3xAlma
                ConfigColsPromedioVentasDiariaBufferMP3Alma
            Case "DescxClixItem"
                CargaVentasxItemxCliente
                ConfigColsDescxClixItem
            Case "VxIxSuc"
'                CargaVentasxItemxSucursal
'                ConfigColsVtaxItemxSuc
                        
            End Select
        Case "Exportar"
            Select Case Me.tag
                Case "VxIxSuc"
                    ExpExcel
            End Select
    End Select
End Sub

'jeaa 04/03/05
Public Sub InicioComprasProveedor(ByVal Cadena As String)
    
    Me.tag = Cadena
    Me.tag = tag
    Me.Caption = "Actualizar Compras por Proveedor"
    CargaDatosComprasProveedor
    ConfigColsComprasProveedor
    Me.Show
    Me.ZOrder
End Sub

Private Sub CargaDatosComprasProveedor()
    Dim sql As String, cond As String, rs As Recordset, antes As Long
    Dim objcond As Condicion
    Static Recargo As String
    
    On Error GoTo ErrTrap
    
    antes = grd.Row
    grd.Rows = 1
    Set objcond = gobjMain.objCondicion
    If Not (frmB_CxTrans.InicioCxProveedor(objcond, Recargo, "ComprasxVendedor")) Then
        grd.SetFocus
        Exit Sub
    End If
       
    grd.Redraw = False
    MensajeStatus MSG_PREPARA, vbHourglass
    
    With objcond
        VerificaExistenciaTabla 0
        cond = " AND gnc.FechaTrans between " & FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB)
        cond = cond & " and " & FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB)
        
        If Len(.CodTrans) > 0 Then
           cond = cond & " AND gnc.CodTrans IN (" & PreparaCadena(.CodTrans) & ")"
        End If
        
        sql = " SELECT ivk.idinventario, ivi.descripcion, "
        sql = sql & " gnc.IdProveedorRef as idProveedor, isnull(pc.nombre,'Saldo Inicial'), "
        sql = sql & " Gnc.CodTrans+' '+CONVERT(varchar,Gnc.NumTrans) as TransRet,"
        sql = sql & " cantidad, (costorealtotal/cantidad) as costoRealUnitraio,"
        sql = sql & " gnc.fechatrans"
        sql = sql & " FROM (IVKARDEX ivk inner join ivinventario ivi on ivk.idinventario=ivi.idinventario) "
        sql = sql & " inner join gncomprobante gnc"
        sql = sql & " left join pcprovcli pc on gnc.IdProveedorRef=pc.IdProvCli"
        sql = sql & " on ivk.transid=gnc.transid"
        sql = sql & " WHERE gnc.Estado <> 3  " & cond
        'sql = sql & " and costorealtotal >0"
        sql = sql & " order by ivi.descripcion, gnc.fechatrans , gnc.horatrans "
    End With
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    grd.LoadArray MiGetRows(rs)
    Set rs = Nothing
    
    With grd
        '# de fila
        .ColAlignment(0) = flexAlignCenterCenter
        GNPoneNumFila grd, False
        'Reubica a la fila donde estaba antes
        If .Rows > antes And antes > 0 Then .Row = antes
        .Redraw = True
    End With
    MensajeStatus "", 0
    grd.SetFocus
    If grd.Rows <> grd.FixedRows Then grd.Row = grd.FixedRows
    Exit Sub

ErrTrap:
    grd.Redraw = True
    MensajeStatus "", 0
    DispErr
    Exit Sub
End Sub


Private Sub ConfigColsComprasProveedor()
    Dim fmt As String
    
    fmt = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda("USD")
    With grd
        .FormatString = "^#|<Id Inventario|<Des Inventario|<Id Proveedor|<Des Proveedor|<Trans|>Cantidad|>CostoUnitario|>FechaGrabado|<Resultado"
       
        .ColWidth(COL_NUM) = 700
        .ColWidth(COL_IDIV) = 1500
        .ColWidth(COL_CODIV) = 3500
        .ColWidth(COL_IDPROV) = 1500
        .ColWidth(COL_CODPROV) = 3500
        .ColWidth(COL_TRANS) = 1000
        .ColWidth(COL_CANT) = 1000
        .ColWidth(COL_COSTO) = 1500
        .ColWidth(COL_FECHA) = 1500
        .ColWidth(COL_RESUL) = 2500
        .ColDataType(COL_NUM) = flexDTLong
        .ColDataType(COL_IDIV) = flexDTString
        .ColDataType(COL_CODIV) = flexDTString
        .ColDataType(COL_IDPROV) = flexDTString
        .ColDataType(COL_CODPROV) = flexDTString
        .ColDataType(COL_TRANS) = flexDTString
        .ColDataType(COL_CANT) = flexDTCurrency
        .ColDataType(COL_COSTO) = flexDTCurrency
        .ColDataType(COL_FECHA) = flexDTDate
        .ColDataType(COL_RESUL) = flexDTString
        
        .ColFormat(COL_COSTO) = "#,###.####"


        .ColHidden(COL_IDIV) = True
        .ColHidden(COL_IDPROV) = True
        '.ColFormat(COL_COSTO) = fmt
    End With
End Sub

Private Sub GrabarComprasProveedor()
    Dim sql As String, idivi As String, i As Long
    Dim idpc As String, cont As Integer
    Dim NumReg As Long, totalventa As Currency
    Dim rs As Recordset, fecha As Date
    Dim sql1 As String
    On Error GoTo ErrTrap
    
    MensajeStatus "Guardando....", 1
    With grd
        If .Rows = .FixedRows Then Exit Sub
        .ShowCell 1, 1
        idivi = ""
        idpc = ""
        For i = .FixedRows To .Rows - 1
            If Not .IsSubtotal(i) Then
                .Row = i
                .ShowCell i, 1           'Hace visible la fila actual
                'verifica si no esta grabado
                sql = "select IdInvPro,    IdInventario, IdProveedor, Cantidad, CostoUnitario, FechaGrabado from InventarioProveedor where IdInventario= " & .TextMatrix(i, COL_IDIV) & " and  IdProveedor= " & .TextMatrix(i, COL_IDPROV)
                Set rs = New ADODB.Recordset
                rs.CursorLocation = adUseClient
                Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                If rs.RecordCount = 0 Then
                    sql = " insert into InventarioProveedor "
                    sql = sql & " ( IdInventario, IdProveedor, "
                    sql = sql & " Cantidad, CostoUnitario, fechaGrabado) values "
                    sql = sql & " (" & .TextMatrix(i, COL_IDIV) & ", " & .TextMatrix(i, COL_IDPROV) & ", "
                    sql = sql & .TextMatrix(i, COL_CANT) & ", " & .TextMatrix(i, COL_COSTO) & ",'"
                    sql = sql & .TextMatrix(i, COL_FECHA) & "')"
                    gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                    If NumReg > 0 Then
                        .TextMatrix(i, COL_RESUL) = "Guardado..."
                    End If
                Else
                    If CDate(.TextMatrix(i, COL_FECHA)) > CDate(rs.Fields("FechaGrabado")) And .TextMatrix(i, COL_COSTO) <> 0 Then
                        fecha = rs.Fields("FechaGrabado")
                        sql = " UPDATE InventarioProveedor "
                        sql = sql & " set  IdInventario = " & .TextMatrix(i, COL_IDIV) & ","
                        sql = sql & " IdProveedor = " & .TextMatrix(i, COL_IDPROV) & ", "
                        sql = sql & " Cantidad = " & .TextMatrix(i, COL_CANT) & ","
                        sql = sql & " CostoUnitario = " & .TextMatrix(i, COL_COSTO) & ","
                        sql = sql & " fechaGrabado= '" & .TextMatrix(i, COL_FECHA) & "'"
                        sql = sql & " WHERE IdInvPro= " & rs.Fields("IdInvPro")
                        sql = sql & " and IdProveedor = " & .TextMatrix(i, COL_IDPROV)
                        gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                        If NumReg > 0 Then
                            .TextMatrix(i, COL_RESUL) = "Actualizado..." & fecha
                        End If
                    Else
                        .TextMatrix(i, COL_RESUL) = "Ya existe en la tabla..."
                    End If
                    If .TextMatrix(i, COL_COSTO) <> 0 Then
                        sql1 = " UPDATE Ivinventario  "
                        sql1 = sql1 & " set  CostoUltimoIngreso = " & .TextMatrix(i, COL_COSTO)
                        sql1 = sql1 & " WHERE Idinventario= " & .TextMatrix(i, COL_IDIV)
                        gobjMain.EmpresaActual.EjecutarSQL sql1, NumReg
                    End If
                  End If
                '---
                sql = "select IdProveedorDetalle, IdInventario, idproveedor, cantidad,              costounitario,         isnull(fecha,'01/01/2001') as Fecha from IVProveedorDetalle where IdInventario= " & .TextMatrix(i, COL_IDIV) & " and  IdProveedor= " & .TextMatrix(i, COL_IDPROV)
                Set rs = New ADODB.Recordset
                rs.CursorLocation = adUseClient
                Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                If rs.RecordCount = 0 Then
                    sql = " insert into IVProveedorDetalle "
                    sql = sql & " ( IdInventario, IdProveedor, "
                    sql = sql & " Cantidad, CostoUnitario, fecha) values "
                    sql = sql & " (" & .TextMatrix(i, COL_IDIV) & ", " & .TextMatrix(i, COL_IDPROV) & ", "
                    sql = sql & .TextMatrix(i, COL_CANT) & ", " & .TextMatrix(i, COL_COSTO) & ",'"
                    sql = sql & .TextMatrix(i, COL_FECHA) & "')"
                    gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                    If NumReg > 0 Then
                        .TextMatrix(i, COL_RESUL) = "Guardado..."
                    End If
                Else
                    If .TextMatrix(i, COL_IDPROV) = rs.Fields("idproveedor") Then
                        If CDate(.TextMatrix(i, COL_FECHA)) > CDate(rs.Fields("Fecha")) And .TextMatrix(i, COL_COSTO) <> 0 Then
                            fecha = rs.Fields("Fecha")
                            sql = " UPDATE IVProveedorDetalle "
                            sql = sql & " set  IdInventario = " & .TextMatrix(i, COL_IDIV) & ","
                            sql = sql & " IdProveedor = " & .TextMatrix(i, COL_IDPROV) & ", "
                            sql = sql & " Cantidad = " & .TextMatrix(i, COL_CANT) & ","
                            sql = sql & " CostoUnitario = " & .TextMatrix(i, COL_COSTO) & ","
                            sql = sql & " fecha= '" & .TextMatrix(i, COL_FECHA) & "'"
                            sql = sql & " WHERE IdInventario= " & rs.Fields("IdInventario")
                            sql = sql & " and IdProveedor = " & .TextMatrix(i, COL_IDPROV)
                            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                            If NumReg > 0 Then
                                .TextMatrix(i, COL_RESUL) = "Actualizado..." & fecha
                            End If
                               Else
                            .TextMatrix(i, COL_RESUL) = "Ya existe en la tabla..."
                        End If
                    End If
                  End If
                sql = " UPDATE IVInventario "
                sql = sql & " set  IdProveedor = " & .TextMatrix(i, COL_IDPROV)
                sql = sql & " WHERE IdInventario= " & .TextMatrix(i, COL_IDIV)
                gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                .Redraw = True
                .Refresh
            End If
        Next i
    End With
    MensajeStatus "", 0
    Exit Sub
    
ErrTrap:
    MsgBox Err.Description, vbExclamation + vbOKOnly
    Exit Sub
End Sub



Public Sub InicioPromVentas(ByVal Cadena As String)
    Me.tag = Cadena
    Me.Caption = "Actualizar Limite de Crédito basado en Promedio Ventas"
    CargaDatosPromedio
    ConfigCols
    Me.Show
    Me.ZOrder
End Sub

Private Sub CargaDatosPromedio()
    Dim sql As String, cond As String, rs As Recordset, antes As Long
    Dim objcond As Condicion
    Static Recargo As String
    
    On Error GoTo ErrTrap
    
    antes = grd.Row
    grd.Rows = 1
    Set objcond = gobjMain.objCondicion
    If Not (frmB_VxTrans.InicioVxMesTransaccion(objcond, Recargo, "PromedioVentas")) Then
        grd.SetFocus
        Exit Sub
    End If
       
    grd.Redraw = False
    MensajeStatus MSG_PREPARA, vbHourglass
    
    With objcond
        VerificaExistenciaTabla 0
        cond = " AND gc.FechaTrans between " & FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB)
        cond = cond & " AND  " & FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB)
        
        If Len(.CodTrans) > 0 Then
           cond = cond & " AND GC.CodTrans IN (" & PreparaCadena(.CodTrans) & ")"
        End If
        
        sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                "From IvRecargo ivR inner join " & _
                    "IvKardexRecargo ivkR Inner join " & _
                        "GnComprobante gc Inner join PcPRovCLi on gc.IdClienteRef = PCProvCli.IdProvCli " & _
                    "On ivkr.TransID = gc.TransID " & _
                "On Ivr.IdRecargo = IvkR.IdRecargo "
        sql = sql & "WHERE gc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(Recargo) & ") " & cond & _
              "Group by IvkR.TransID"
        
        gobjMain.EmpresaActual.EjecutarSQL sql, 0
    
        sql = "SELECT PCProvCli.CodProvCli, PCProvCli.Nombre, " & _
              "round(ABS(SUM((PrecioTotalBase0 + (PrecioTotalBase0 * (cast(TotalDescuento as float) / cast(PrecioTotal as float))))*SignoVenta) + " & _
              "SUM((PrecioTotalBaseIVA + (PrecioTotalBaseIVA * (cast(TotalDescuento as float) / cast(PrecioTotal as float))))*SignoVenta)) "
        If DateDiff("m", .fecha1, .fecha2) > 1 Then
            sql = sql & " / (max(datediff(m,"
            sql = sql & FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & ","
            sql = sql & FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & "))),0) as Promediomes,"
        Else
            sql = sql & ",0) as Promediomes, "
        End If
        sql = sql & "SPACE(0) AS Resultado "
        sql = sql & "FROM tmp0 inner join " & _
                    "vwConsSUMIVKardexIVA inner join " & _
                        "GNComprobante GC  inner JOIN PCProvCli ON GC.IdClienteRef = PCProvCli.IdProvCli " & _
                        "ON vwConsSUMIVKardexIVA.TransID = GC.TransID " & _
                    "ON tmp0.TransID = GC.TransID "
        'jeaa 30/09/04 agegado AND PRECIOTOTAL<>0 para controlar error division para cero
        sql = sql & " WHERE (gc.Estado<>3) AND PRECIOTOTAL<>0" & cond & _
                    " GROUP BY PCProvCli.CodProvCli, PCProvCli.Nombre, PCProvCli.TotalDebe "
        
        'Para filtrar solo los clientes cuyo monto de venta se ha modificado,
        'es decir solo los que han comprado
        sql = sql & "HAVING round(ABS(SUM((PrecioTotalBase0 + (PrecioTotalBase0 * " & _
                    "(cast(TotalDescuento as float) / cast(PrecioTotal as float))))*SignoVenta) + " & _
                    "SUM((PrecioTotalBaseIVA + (PrecioTotalBaseIVA * (cast(TotalDescuento as float) "
        If DateDiff("m", .fecha1, .fecha2) > 1 Then
            sql = sql & "/"
            sql = sql & " cast(PrecioTotal as float))))*SignoVenta)),4)<>PCProvCli.TotalDebe "
        Else
            sql = sql & " )))*SignoVenta)),4)<>PCProvCli.TotalDebe "
        End If
                    
                    
        sql = sql & "ORDER BY PCProvCli.Nombre"
    End With
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    grd.LoadArray MiGetRows(rs)
    Set rs = Nothing
    
    With grd
        '# de fila
        .ColAlignment(0) = flexAlignCenterCenter
        GNPoneNumFila grd, False
        'Reubica a la fila donde estaba antes
        If .Rows > antes And antes > 0 Then .Row = antes
        .Redraw = True
    End With
    MensajeStatus "", 0
    grd.SetFocus
    If grd.Rows <> grd.FixedRows Then grd.Row = grd.FixedRows
    Exit Sub

ErrTrap:
    grd.Redraw = True
    MensajeStatus "", 0
    DispErr
    Exit Sub
End Sub


Private Sub GrabarPromedio()
    Dim sql As String, cod As String, i As Long
    Dim NumReg As Long, totalventa As Currency
    On Error GoTo ErrTrap
    
    MensajeStatus "Guardando....", 1
    With grd
        If .Rows = .FixedRows Then Exit Sub
        .ShowCell 1, 1
        For i = .FixedRows To .Rows - 1
            If Not .IsSubtotal(i) Then
                .Row = i
                .ShowCell i, 1           'Hace visible la fila actual
                cod = .TextMatrix(i, COL_COD)
                totalventa = .ValueMatrix(i, COL_TOT)
                sql = "UPDATE PCProvCli " & _
                      "SET LimiteCredito = " & totalventa & _
                      "WHERE CodProvCli= '" & cod & "'"
                gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                If NumReg > 0 Then
                    .TextMatrix(i, COL_RES) = "Actualizado..."
                Else
                    .TextMatrix(i, COL_RES) = "Error al tratar de Actualizar..."
                End If
                .Redraw = True
                .Refresh
            End If
        Next i
    End With
    MensajeStatus "", 0
    Exit Sub
    
ErrTrap:
    MsgBox Err.Description, vbExclamation + vbOKOnly
    Exit Sub
End Sub


Private Sub CargaDatosCustodioActivos()
    Dim sql As String, cond As String, rs As Recordset, antes As Long
    Dim objcond As Condicion
    Static Recargo As String
    
    On Error GoTo ErrTrap
    
    antes = grd.Row
    grd.Rows = 1
    Set objcond = gobjMain.objCondicion
    If Not (frmB_CxTrans.InicioCustodioxActivo(objcond, Recargo, "CustoxActivo")) Then
        grd.SetFocus
        Exit Sub
    End If
       
    grd.Redraw = False
    MensajeStatus MSG_PREPARA, vbHourglass
    
    With objcond
        VerificaExistenciaTabla 0
        cond = " AND gnc.FechaTrans between " & FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB)
        cond = cond & " and " & FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB)
        
        If Len(.CodTrans) > 0 Then
           cond = cond & " AND gnc.CodTrans IN (" & PreparaCadena(.CodTrans) & ")"
        End If
        
        sql = " SELECT  "
        sql = sql & " ivk.idinventario, ivi.descripcion,"
        sql = sql & " ivk.IdProvCli as idEmpleado,"
        sql = sql & " pc.nombre,  Gnc.CodTrans+' '+CONVERT(varchar,Gnc.NumTrans) as TransRet,"
        sql = sql & " cantidad, ''as nombre, gnc.fechatrans"
        sql = sql & " FROM (afKARDEXcustodio ivk "
        sql = sql & " inner join empleado pc on ivk.IdProvCli=pc.IdProvCli"
        ''sql = sql & " inner join pcprovcli pcact on ivi.Idempleado=pcact.IdProvCli"
        
        sql = sql & " inner join afinventario ivi on ivk.idinventario=ivi.idinventario)"
        
        sql = sql & " inner join gncomprobante gnc"
        sql = sql & " on ivk.transid=gnc.transid"
        sql = sql & " WHERE gnc.Estado <> 3  " & cond
        'sql = sql & " and pc.codprovcli='147' "
        sql = sql & " order by ivi.descripcion, gnc.fechatrans , gnc.horatrans "
    End With
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    grd.LoadArray MiGetRows(rs)
    Set rs = Nothing
    
    With grd
        '# de fila
        .ColAlignment(0) = flexAlignCenterCenter
        GNPoneNumFila grd, False
        'Reubica a la fila donde estaba antes
        If .Rows > antes And antes > 0 Then .Row = antes
        .Redraw = True
    End With
    MensajeStatus "", 0
    grd.SetFocus
    If grd.Rows <> grd.FixedRows Then grd.Row = grd.FixedRows
    Exit Sub

ErrTrap:
    grd.Redraw = True
    MensajeStatus "", 0
    DispErr
    Exit Sub
End Sub

Private Sub ConfigColsCustodioxActivo()
    Dim fmt As String
    
    fmt = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda("USD")
    With grd
        .FormatString = "^#|<Id Inventario|<Des Activo|<Id Proveedor|<Custodio|<Trans|>Cantidad|<Custodio en Activo|>FechaGrabado|<Resultado"
       
        .ColWidth(COL_NUM) = 700
        .ColWidth(COL_IDIV) = 1500
        .ColWidth(COL_CODIV) = 3500
        .ColWidth(COL_IDPROV) = 1500
        .ColWidth(COL_CODPROV) = 3500
        .ColWidth(COL_TRANS) = 1000
        .ColWidth(COL_CANT) = 1000
        .ColWidth(COL_COSTO) = 1500
        .ColWidth(COL_FECHA) = 1500
        .ColWidth(COL_RESUL) = 2500
        .ColDataType(COL_NUM) = flexDTLong
        .ColDataType(COL_IDIV) = flexDTString
        .ColDataType(COL_CODIV) = flexDTString
        .ColDataType(COL_IDPROV) = flexDTString
        .ColDataType(COL_CODPROV) = flexDTString
        .ColDataType(COL_TRANS) = flexDTString
        .ColDataType(COL_CANT) = flexDTCurrency
        .ColDataType(COL_COSTO) = flexDTCurrency
        .ColDataType(COL_FECHA) = flexDTDate
        .ColDataType(COL_RESUL) = flexDTString
        
        .ColFormat(COL_COSTO) = "#,###.####"


        .ColHidden(COL_IDIV) = True
        .ColHidden(COL_IDPROV) = True
        '.ColFormat(COL_COSTO) = fmt
    End With
End Sub

Public Sub InicioCustodioxActivo(ByVal Cadena As String)
    
    Me.tag = Cadena
    Me.tag = tag
    Me.Caption = "Actualizar Custodio x Activo"
    CargaDatosCustodioActivos
    ConfigColsCustodioxActivo
    Me.Show
    Me.ZOrder
End Sub


Private Sub GrabarCustodioxActivo()
    Dim sql As String, idivi As String, i As Long
    Dim idpc As String, cont As Integer
    Dim NumReg As Long, totalventa As Currency
    Dim rs As Recordset, fecha As Date
    Dim sql1 As String
    On Error GoTo ErrTrap
    
    MensajeStatus "Guardando....", 1
    With grd
        If .Rows = .FixedRows Then Exit Sub
        .ShowCell 1, 1
        idivi = ""
        idpc = ""
        For i = .FixedRows To .Rows - 1
            If Not .IsSubtotal(i) Then
                .Row = i
                .ShowCell i, 1           'Hace visible la fila actual
                If .ValueMatrix(i, COL_CANT) > 0 Then
                    sql = " UPDATE AFInventario "
                    sql = sql & " set  IdEmpleado = " & .TextMatrix(i, COL_IDPROV)
                    sql = sql & " WHERE IdInventario= " & .TextMatrix(i, COL_IDIV)
                    gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                End If
                
                .Redraw = True
                .Refresh
            End If
        Next i
    End With
    MensajeStatus "", 0
    Exit Sub
    
ErrTrap:
    MsgBox Err.Description, vbExclamation + vbOKOnly
    Exit Sub
End Sub

Public Sub InicioPcGrupoxMontoVentas(ByVal Cadena As String)
    Me.tag = Cadena
    Me.Caption = "Actualizar PcGrupo x Monto de Ventas"
    CargaDatosMontoVentas
    ConfigCols
    Me.Show
    Me.ZOrder
End Sub

Private Sub CargaDatosMontoVentas()
    Dim sql As String, cond As String, rs As Recordset, antes As Long, i As Integer, j As Long
    Dim objcond As Condicion
    Static Recargo As String
    
    On Error GoTo ErrTrap
    
    antes = grd.Row
    grd.Rows = 1
    Set objcond = gobjMain.objCondicion
    If Not (frmB_VxTrans.InicioMontoVentas(objcond, Recargo, "PCGMontoVentas")) Then
        grd.SetFocus
        Exit Sub
    End If
       
    grd.Redraw = False
    MensajeStatus MSG_PREPARA, vbHourglass
    
    With objcond
        NumPCGrupo = RecuperaSelecPCGrupo + 1
    
        VerificaExistenciaTabla 0
        cond = " AND gc.FechaTrans between " & FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB)
        cond = cond & " AND  " & FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB)
        
        If Len(.CodTrans) > 0 Then
           cond = cond & " AND GC.CodTrans IN (" & PreparaCadena(.CodTrans) & ")"
        End If
        
    
        sql = "SELECT PCProvCli.CodProvCli, PCProvCli.Nombre, "
        sql = sql & " round(ABS(SUM(PrecioRealTotal)),2), codgrupo" & NumPCGrupo & ",'' as newgrupo,"
        sql = sql & " SPACE(0) AS Resultado "
        sql = sql & " FROM  "
        sql = sql & " vwConsSUMIVKardexIVA inner join "
        sql = sql & " GNComprobante GC  "
        sql = sql & " inner JOIN PCProvCli "
        sql = sql & " left join pcgrupo" & NumPCGrupo
        sql = sql & " on PCProvCli.idgrupo" & NumPCGrupo & " = pcgrupo" & NumPCGrupo & ".idgrupo" & NumPCGrupo
        sql = sql & " ON GC.IdClienteRef = PCProvCli.IdProvCli "
        sql = sql & " ON vwConsSUMIVKardexIVA.TransID = GC.TransID "
        sql = sql & " WHERE (gc.Estado<>3) AND PRECIORealTOTAL<>0" & cond
        sql = sql & " GROUP BY PCProvCli.CodProvCli, PCProvCli.Nombre, PCProvCli.TotalDebe, "
        'sql = sql & " SignoVenta, PCProvCli.idgrupo" & NumPCGrupo & ", "
        sql = sql & " pcgrupo" & NumPCGrupo & ".codgrupo" & NumPCGrupo
        
        'Para filtrar solo los clientes cuyo monto de venta se ha modificado,
        'es decir solo los que han comprado
        sql = sql & " having round(ABS(SUM(PrecioRealTotal)),4) <> PCProvCli.TotalDebe "
        
                    
                    
        sql = sql & "ORDER BY PCProvCli.Nombre"
    End With
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    grd.LoadArray MiGetRows(rs)

    Set rs = Nothing
    
    
    
    With grd
        '# de fila
        .ColAlignment(0) = flexAlignCenterCenter
        GNPoneNumFila grd, False
        'Reubica a la fila donde estaba antes
        If .Rows > antes And antes > 0 Then .Row = antes
        .Redraw = True
        ConfigCols
        
        
    End With
    MensajeStatus "", 0
    grd.SetFocus
    If grd.Rows <> grd.FixedRows Then grd.Row = grd.FixedRows
    
    If CargarPCGrupos(RecuperaSelecPCGrupo) Then
    End If
    
    
    For i = 1 To grd.Rows - 1

        For j = 1 To 10
            If grd.ValueMatrix(i, COL_TOT) >= gMonto(j).desde And grd.ValueMatrix(i, COL_TOT) <= gMonto(j).hasta Then
                grd.TextMatrix(i, COL_MV_NEWGRUPO) = gMonto(j).grupo
                Exit For
            End If
        Next j
    
    Next i
    
    
    Exit Sub

ErrTrap:
    grd.Redraw = True
    MensajeStatus "", 0
    DispErr
    Exit Sub
End Sub


Private Function RecuperaSelecPCGrupo() As Integer
Dim s As String, Vector As Variant, ix As Long
Dim i As Integer, j As Integer, Selec As Integer
    'Recupera selecciondados  del registro de windows
    's = GetSetting(APPNAME, App.Title, "PCGMontoVenta_NumPCGrupo", "_VACIO_")
    s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("PCGMontoVenta_NumPCGrupo")
    If Len(s) > 0 Then
        RecuperaSelecPCGrupo = CInt(s) - 1
    Else
        RecuperaSelecPCGrupo = 0
    End If

End Function


Private Function CargarPCGrupos(ByVal numGrupo As Integer) As Boolean
    Dim s As Variant
    On Error GoTo ErrTrap
    With grd
        CargarPCGrupos = True
        s = gobjMain.EmpresaActual.ListaPCGrupoParaFlexGrid(numGrupo)
        If Len(s) > 1 Then
            s = Right$(s, Len(s) - 1)
            .ColComboList(COL_MV_GRUPO) = s
        End If
    End With
    Exit Function
ErrTrap:
        MsgBox "No se han definido PCGrupos", vbInformation
        CargarPCGrupos = False
    Exit Function
End Function

Private Sub GrabarPCGrupoxMontoVenta()
    Dim sql As String, cod As String, i As Long
    Dim NumReg As Long, totalventa As Currency
    On Error GoTo ErrTrap
    
    MensajeStatus "Guardando....", 1
    With grd
        If .Rows = .FixedRows Then Exit Sub
        .ShowCell 1, 1
        For i = .FixedRows To .Rows - 1
            If Not .IsSubtotal(i) Then
                .Row = i
                .ShowCell i, 1           'Hace visible la fila actual
                cod = .TextMatrix(i, COL_COD)
                If grd.TextMatrix(i, COL_MV_NEWGRUPO) <> grd.TextMatrix(i, COL_MV_GRUPO) Then
                    sql = " UPDATE PCProvCli "
                    sql = sql & " SET idGrupo" & NumPCGrupo & " = (select idgrupo" & NumPCGrupo & " from pcgrupo" & NumPCGrupo & " where codgrupo" & NumPCGrupo & "='" & grd.TextMatrix(i, COL_MV_NEWGRUPO) & "')"
                    sql = sql & " WHERE CodProvCli= '" & cod & "'"
                    gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                    If NumReg > 0 Then
                        .TextMatrix(i, COL_MV_RES) = "Actualizado..."
                    Else
                        .TextMatrix(i, COL_MV_RES) = "Error al tratar de Actualizar..."
                    End If
                Else
                        .TextMatrix(i, COL_MV_RES) = "No Hay Cambio.........."
                End If
                .Redraw = True
                .Refresh
            End If
        Next i
    End With
    MensajeStatus "", 0
    Exit Sub
    
ErrTrap:
    MsgBox Err.Description, vbExclamation + vbOKOnly
    Exit Sub
End Sub

Public Sub InicioFechaUltimoEgreso(ByVal Cadena As String)
    
    Me.tag = Cadena
    Me.tag = tag
    Me.Caption = "Actualizar Fecha Ultimo Egreso"
    CargaDatosFechaUltimoEgreso
    ConfigColsFechaUltimoEgreso
    
    Me.Show
    Me.ZOrder
End Sub


Private Sub CargaDatosFechaUltimoEgreso()
    Dim sql As String, cond As String, rs As Recordset, antes As Long, NumReg As Long
    Dim objcond As Condicion, base As String
    Static Recargo As String
    
    
    On Error GoTo ErrTrap
    
    antes = grd.Row
    grd.Rows = 1
    Set objcond = gobjMain.objCondicion
    If Not (frmB_CxTrans.InicioActualizaFechaIngreso(objcond, Recargo, "FechgaEgreso")) Then
        grd.SetFocus
        Exit Sub
    End If
       
    grd.Redraw = False
    MensajeStatus MSG_PREPARA, vbHourglass
    
    With objcond
        base = objcond.Sucursal & ".dbo."
        
''        sql = " select iv.idinventario, codinventario, iv.descripcion, i.idbodega, codbodega, ivb.descripcion, ive.FechaUltimoEgreso, max(fechatrans) as fechaUltimoEgreso"
''        sql = sql & " from " & base & "gncomprobante g inner join " & base & "gntrans gt on g.codtrans=gt.codtrans"
''        sql = sql & " inner join " & base & "ivkardex i inner join " & base & "ivbodega ivb on i.idbodega=ivb.idbodega"
''        sql = sql & " inner join " & base & "ivinventario iv"
''        sql = sql & " on i.idinventario=iv.idinventario"
''        sql = sql & " on g.transid = i.transid"
''        sql = sql & " left join ivexist ive on iv.idinventario= ive.idinventario and i.idbodega = ive.idbodega"
''        sql = sql & " Where Estado <> 3 and cantidad<0   "
''        sql = sql & " AND  g.codtrans in (" & .CodTrans & ")"
'' '       sql = sql & " AND  AnexoCodTipoTrans in ('2')"
''        sql = sql & " group by iv.idinventario, codinventario, iv.descripcion, i.idbodega, codbodega, ivb.descripcion, ive.FechaUltimoEgreso"
''        sql = sql & " having max(fechatrans) < '" & CDate("01/" & DatePart("m", Date) & "/" & DatePart("yyyy", Date)) & "'"
''        sql = sql & " order by codinventario, codbodega"

        sql = " select iv.idinventario, codinventario, iv.descripcion as descitem, i.idbodega, codbodega, ivb.descripcion as descbodega, ive.FechaUltimoEgreso, (fechatrans) as fechaUltimoEgresoNew"
        sql = sql & " into tmp0"
        sql = sql & " from " & base & "gncomprobante g inner join " & base & "gntrans gt on g.codtrans=gt.codtrans"
        sql = sql & " inner join " & base & "ivkardex i inner join " & base & "ivbodega ivb on i.idbodega=ivb.idbodega"
        sql = sql & " inner join " & base & "ivinventario iv"
        sql = sql & " on i.idinventario=iv.idinventario"
        sql = sql & " on g.transid = i.transid"
        sql = sql & " left join ivexist ive on iv.idinventario= ive.idinventario and i.idbodega = ive.idbodega"
        sql = sql & " Where Estado <> 3 and cantidad<0"
        sql = sql & " AND  g.codtrans in (" & .CodTrans & ")"
        sql = sql & " and fechatrans <= " & FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB)
        'sql = sql & " and (fechatrans) < '" & CDate("01/" & DatePart("m", Date) & "/" & DatePart("yyyy", Date)) & "'"
        sql = sql & " order by codinventario, codbodega"
        
        VerificaExistenciaTabla 0
        gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
        
        sql = " select idinventario, codinventario, descitem, idbodega, codbodega, descbodega, FechaUltimoEgreso, max(fechaUltimoEgresoNew) as fechaUltimoEgresoNew"
        sql = sql & " from tmp0"
        sql = sql & " group by idinventario, codinventario, descitem, idbodega, codbodega, descbodega, FechaUltimoEgreso"
       ' sql = sql & " having max(fechaUltimoEgresoNew) < '" & CDate("01/" & DatePart("m", Date) & "/" & DatePart("yyyy", Date)) & "'"
        'sql = sql & " ) and (max(fechaUltimoEgresoNew) > max(fechaUltimoEgreso)) "
        sql = sql & " order by codinventario, codbodega"

    End With
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    grd.LoadArray MiGetRows(rs)
    Set rs = Nothing
    
    With grd
        '# de fila
        .ColAlignment(0) = flexAlignCenterCenter
        GNPoneNumFila grd, False
        'Reubica a la fila donde estaba antes
        If .Rows > antes And antes > 0 Then .Row = antes
        .Redraw = True
    End With
    MensajeStatus "", 0
    grd.SetFocus
    If grd.Rows <> grd.FixedRows Then grd.Row = grd.FixedRows
    Exit Sub

ErrTrap:
    grd.Redraw = True
    MensajeStatus "", 0
    DispErr
    Exit Sub
End Sub


Private Sub ConfigColsFechaUltimoEgreso()
    Dim fmt As String
    
    fmt = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda("USD")
    With grd
        .FormatString = "^#|<Id Inventario|<Cod Inventario|<Des Inventario|<Id Bodega|<Cod Bodega|<Des Bodega|>Fecha Ultimo Egreso|>Fecha Ultimo Egreso New|<Resultado"
       
        .ColWidth(COL_NUM) = 700
        .ColWidth(COL_F_IDIV) = 1500
        .ColWidth(COL_F_CODIV) = 2000
        .ColWidth(COL_F_DESCIV) = 3500
        .ColWidth(COL_F_IDBOD) = 1500
        .ColWidth(COL_F_CODBOD) = 1500
        .ColWidth(COL_F_DESCBOD) = 3500
        .ColWidth(COL_F_FECHA) = 1500
        .ColWidth(COL_F_FECHAN) = 1500
        .ColWidth(COL_F_RESUL) = 2500
        .ColDataType(COL_NUM) = flexDTLong
        .ColDataType(COL_F_IDIV) = flexDTString
        .ColDataType(COL_F_CODIV) = flexDTString
        .ColDataType(COL_F_DESCIV) = flexDTString
        .ColDataType(COL_F_IDBOD) = flexDTString
        .ColDataType(COL_F_CODBOD) = flexDTString
        .ColDataType(COL_F_DESCBOD) = flexDTString
        .ColDataType(COL_F_FECHA) = flexDTDate
        .ColDataType(COL_F_RESUL) = flexDTString
        


        .ColHidden(COL_F_IDIV) = True
        .ColHidden(COL_F_IDBOD) = True

    End With
End Sub

Private Sub GrabarFechaUltimoEgreso()
    Dim sql As String, cod As String, i As Long
    Dim NumReg As Long, totalventa As Currency
    On Error GoTo ErrTrap
    
    MensajeStatus "Guardando....", 1
    With grd
        If .Rows = .FixedRows Then Exit Sub
        .ShowCell 1, 1
        For i = .FixedRows To .Rows - 1
            If Not .IsSubtotal(i) Then
                .Row = i
                .ShowCell i, 1           'Hace visible la fila actual
'                cod = .TextMatrix(i, COL_COD)
                If Len(grd.TextMatrix(i, COL_F_FECHA)) = 0 Then
                        sql = " UPDATE Ivexist "
                        sql = sql & " SET FechaUltimoEgreso ='" & grd.TextMatrix(i, COL_F_FECHAN) & "'"
                        sql = sql & " WHERE idinventario= " & grd.ValueMatrix(i, COL_F_IDIV) & " and idbodega=" & grd.ValueMatrix(i, COL_F_IDBOD)
                        gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                        If NumReg > 0 Then
                            .TextMatrix(i, COL_F_RESUL) = "Actualizado..."
                        Else
                            sql = " Insert Ivexist  (IdInventario, IdBodega, Exist, ExistMin, ExistMax, FechaUltimoIngreso, FechaUltimoEgreso) values ("
                            sql = sql & grd.ValueMatrix(i, COL_F_IDIV) & "," & grd.ValueMatrix(i, COL_F_IDBOD) & ",0,0,0,'','" & grd.TextMatrix(i, COL_F_FECHAN) & "')"
                            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                            If NumReg > 0 Then
                                .TextMatrix(i, COL_F_RESUL) = "Insertado..."
                            Else
                                .TextMatrix(i, COL_F_RESUL) = "Error al tratar de Actualizar..."
                            End If
                        End If
                
                Else
                    If (CDate(grd.TextMatrix(i, COL_F_FECHAN)) > CDate(grd.TextMatrix(i, COL_F_FECHA))) Then
                        sql = " UPDATE Ivexist "
                        sql = sql & " SET FechaUltimoEgreso ='" & grd.TextMatrix(i, COL_F_FECHAN) & "'"
                        sql = sql & " WHERE idinventario= " & grd.ValueMatrix(i, COL_F_IDIV) & " and idbodega=" & grd.ValueMatrix(i, COL_F_IDBOD)
                        gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                        If NumReg > 0 Then
                            .TextMatrix(i, COL_F_RESUL) = "Actualizado..."
                        Else
                            .TextMatrix(i, COL_F_RESUL) = "Error al tratar de Actualizar..."
                        End If
                    End If
                End If
                .Redraw = True
                .Refresh
            End If
        Next i
    End With
    MensajeStatus "", 0
    Exit Sub
    
ErrTrap:
    MsgBox Err.Description, vbExclamation + vbOKOnly
    Exit Sub
End Sub


Public Sub InicioFechaUltimoIngreso(ByVal Cadena As String)
    
    Me.tag = Cadena
    Me.tag = tag
    Me.Caption = "Actualizar Fecha Ultimo Ingreso"
    CargaDatosFechaUltimoIngreso
    ConfigColsFechaUltimoIngreso
    Me.Show
    Me.ZOrder
End Sub

Private Sub CargaDatosFechaUltimoIngreso()
    Dim sql As String, cond As String, rs As Recordset, antes As Long, NumReg As Long
    Dim objcond As Condicion, base As String
    Static Recargo As String
    
    
    On Error GoTo ErrTrap
    
    antes = grd.Row
    grd.Rows = 1
    Set objcond = gobjMain.objCondicion
    If Not (frmB_CxTrans.InicioActualizaFechaIngreso(objcond, Recargo, "FechgaIngreso")) Then
        grd.SetFocus
        Exit Sub
    End If
       
    grd.Redraw = False
    MensajeStatus MSG_PREPARA, vbHourglass
    
    With objcond
        base = objcond.Sucursal & ".dbo."
        
        sql = " select iv.idinventario, codinventario, iv.descripcion as descitem, i.idbodega, codbodega, ivb.descripcion as descbodega, ive.FechaUltimoIngreso, (fechatrans) as fechaUltimoIngresoNew"
        sql = sql & " into tmp0"
        sql = sql & " from " & base & "gncomprobante g inner join " & base & "gntrans gt on g.codtrans=gt.codtrans"
        sql = sql & " inner join " & base & "ivkardex i inner join " & base & "ivbodega ivb on i.idbodega=ivb.idbodega"
        sql = sql & " inner join " & base & "ivinventario iv"
        sql = sql & " on i.idinventario=iv.idinventario"
        sql = sql & " on g.transid = i.transid"
        sql = sql & " left join ivexist ive on iv.idinventario= ive.idinventario and i.idbodega = ive.idbodega"
        sql = sql & " Where Estado <> 3 and cantidad>0"
        sql = sql & " AND  g.codtrans in (" & .CodTrans & ")"
        'sql = sql & " and (fechatrans) < '" & CDate("01/" & DatePart("m", Date) & "/" & DatePart("yyyy", Date)) & "'"
        sql = sql & " and fechatrans <= " & FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB)
        sql = sql & " order by codinventario, codbodega"
        
        VerificaExistenciaTabla 0
        gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
        
        sql = " select idinventario, codinventario, descitem, idbodega, codbodega, descbodega, FechaUltimoIngreso, max(fechaUltimoIngresoNew) as fechaUltimoIngresoNew"
        sql = sql & " from  tmp0"
        sql = sql & " group by idinventario, codinventario, descitem, idbodega, codbodega, descbodega, FechaUltimoIngreso"
        'sql = sql & " having max(fechaUltimoIngresoNew) < '" & CDate("01/" & DatePart("m", Date) & "/" & DatePart("yyyy", Date)) & "'"
        'sql = sql & " and fechatrans <= " & FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB)
        sql = sql & " order by codinventario, codbodega"


    End With
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    grd.LoadArray MiGetRows(rs)
    Set rs = Nothing
    
    With grd
        '# de fila
        .ColAlignment(0) = flexAlignCenterCenter
        GNPoneNumFila grd, False
        'Reubica a la fila donde estaba antes
        If .Rows > antes And antes > 0 Then .Row = antes
        .Redraw = True
    End With
    MensajeStatus "", 0
    grd.SetFocus
    If grd.Rows <> grd.FixedRows Then grd.Row = grd.FixedRows
    Exit Sub

ErrTrap:
    grd.Redraw = True
    MensajeStatus "", 0
    DispErr
    Exit Sub
End Sub


Private Sub ConfigColsFechaUltimoIngreso()
    Dim fmt As String
    
    fmt = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda("USD")
    With grd
        .FormatString = "^#|<Id Inventario|<Cod Inventario|<Des Inventario|<Id Bodega|<Cod Bodega|<Des Bodega|>Fecha Ultimo Ingreso|>Fecha Ultimo Ingreso New|<Resultado"
       
        .ColWidth(COL_NUM) = 700
        .ColWidth(COL_F_IDIV) = 1500
        .ColWidth(COL_F_CODIV) = 2000
        .ColWidth(COL_F_DESCIV) = 3500
        .ColWidth(COL_F_IDBOD) = 1500
        .ColWidth(COL_F_CODBOD) = 1500
        .ColWidth(COL_F_DESCBOD) = 3500
        .ColWidth(COL_F_FECHA) = 1500
        .ColWidth(COL_F_FECHAN) = 1500
        .ColWidth(COL_F_RESUL) = 2500
        .ColDataType(COL_NUM) = flexDTLong
        .ColDataType(COL_F_IDIV) = flexDTString
        .ColDataType(COL_F_CODIV) = flexDTString
        .ColDataType(COL_F_DESCIV) = flexDTString
        .ColDataType(COL_F_IDBOD) = flexDTString
        .ColDataType(COL_F_CODBOD) = flexDTString
        .ColDataType(COL_F_DESCBOD) = flexDTString
        .ColDataType(COL_F_FECHA) = flexDTDate
        .ColDataType(COL_F_RESUL) = flexDTString
        


        .ColHidden(COL_F_IDIV) = True
        .ColHidden(COL_F_IDBOD) = True

    End With
End Sub

Private Sub GrabarFechaUltimoIngreso()
    Dim sql As String, cod As String, i As Long
    Dim NumReg As Long, totalventa As Currency
    On Error GoTo ErrTrap
    
    MensajeStatus "Guardando....", 1
    With grd
        If .Rows = .FixedRows Then Exit Sub
        .ShowCell 1, 1
        For i = .FixedRows To .Rows - 1
            If Not .IsSubtotal(i) Then
                .Row = i
                .ShowCell i, 1           'Hace visible la fila actual
'                cod = .TextMatrix(i, COL_COD)
'''''                If Len(grd.TextMatrix(i, COL_F_FECHA)) = 0 Then
'''''                    sql = " UPDATE Ivexist "
'''''                    sql = sql & " SET FechaUltimoIngreso ='" & grd.TextMatrix(i, COL_F_FECHAN) & "'"
'''''                    sql = sql & " WHERE idinventario= " & grd.ValueMatrix(i, COL_F_IDIV) & " and idbodega=" & grd.ValueMatrix(i, COL_F_IDBOD)
'''''                    gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
'''''                    If NumReg > 0 Then
'''''                        .TextMatrix(i, COL_F_RESUL) = "Actualizado..."
'''''                    Else
'''''                        .TextMatrix(i, COL_F_RESUL) = "Error al tratar de Actualizar..."
'''''                    End If
'''''                End If


                If Len(grd.TextMatrix(i, COL_F_FECHA)) = 0 Then
                        sql = " UPDATE Ivexist "
                        sql = sql & " SET FechaUltimoIngreso ='" & grd.TextMatrix(i, COL_F_FECHAN) & "'"
                        sql = sql & " WHERE idinventario= " & grd.ValueMatrix(i, COL_F_IDIV) & " and idbodega=" & grd.ValueMatrix(i, COL_F_IDBOD)
                        gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                        If NumReg > 0 Then
                            .TextMatrix(i, COL_F_RESUL) = "Actualizado..."
                        Else
                            sql = " Insert Ivexist  (IdInventario, IdBodega, Exist, ExistMin, ExistMax, FechaUltimoIngreso, FechaUltimoEgreso) values ("
                            sql = sql & grd.ValueMatrix(i, COL_F_IDIV) & "," & grd.ValueMatrix(i, COL_F_IDBOD) & ",0,0,0,'" & grd.TextMatrix(i, COL_F_FECHAN) & "','')"
                            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                            If NumReg > 0 Then
                                .TextMatrix(i, COL_F_RESUL) = "Insertado..."
                            Else
                                .TextMatrix(i, COL_F_RESUL) = "Error al tratar de Actualizar..."
                            End If
                        End If
                
                Else
                    If (CDate(grd.TextMatrix(i, COL_F_FECHAN)) > CDate(grd.TextMatrix(i, COL_F_FECHA))) Then
                        sql = " UPDATE Ivexist "
                        sql = sql & " SET FechaUltimoIngreso ='" & grd.TextMatrix(i, COL_F_FECHAN) & "'"
                        sql = sql & " WHERE idinventario= " & grd.ValueMatrix(i, COL_F_IDIV) & " and idbodega=" & grd.ValueMatrix(i, COL_F_IDBOD)
                        gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                        If NumReg > 0 Then
                            .TextMatrix(i, COL_F_RESUL) = "Actualizado..."
                        Else
                            .TextMatrix(i, COL_F_RESUL) = "Error al tratar de Actualizar..."
                        End If
                    End If
                End If

                .Redraw = True
                .Refresh
            End If
        Next i
    End With
    MensajeStatus "", 0
    Exit Sub
    
ErrTrap:
    MsgBox Err.Description, vbExclamation + vbOKOnly
    Exit Sub
End Sub


Private Sub CargaPromedioVentasDiariaBufferMP3()
    Dim sql As String, cond As String, rs As Recordset, antes As Long, NumReg As Long
    Dim objcond As Condicion, base As String
    Static Recargo As String
    
    
    On Error GoTo ErrTrap
    
    antes = grd.Row
    grd.Rows = 1
    Set objcond = gobjMain.objCondicion
    If Not (frmB_CxTrans.InicioActualizaFechaIngreso(objcond, Recargo, "FechgaEgreso")) Then
        grd.SetFocus
        Exit Sub
    End If
       
    grd.Redraw = False
    MensajeStatus MSG_PREPARA, vbHourglass
    
    With objcond
        base = objcond.Sucursal & ".dbo."
        
        sql = " select *,cant,aCUMULADO,0 as prom from vwVentaxItemxFecha"
        'sql = sql & " Where idinventario = 54298"
        sql = sql & " Where 1=1 "
        sql = sql & "  and fechatrans between '" & Date - 180 & "' and '" & Date & "'"
        
        sql = sql & " order by descripcion, fechatrans"
        
        

    End With
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    grd.LoadArray MiGetRows(rs)
    Set rs = Nothing
    
    With grd
        '# de fila
        .ColAlignment(0) = flexAlignCenterCenter
        GNPoneNumFila grd, False
        'Reubica a la fila donde estaba antes
        If .Rows > antes And antes > 0 Then .Row = antes
        .Redraw = True
    End With
    MensajeStatus "", 0
    grd.SetFocus
    If grd.Rows <> grd.FixedRows Then grd.Row = grd.FixedRows
    
    
    Exit Sub

ErrTrap:
    grd.Redraw = True
    MensajeStatus "", 0
    DispErr
    Exit Sub
End Sub

Private Sub ConfigColsPromedioVentasDiariaBufferMP3()
    Dim fmt As String
    
    fmt = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda("USD")
    With grd
        .FormatString = "^#|<Id Inventario|<CodInventario|<Descripcion|>Venta|>T.Repo|<Fecha Trans |>Acumulado|>Desv. Estandar|>Max.Prom|>Calculo para Buffer|>Valor para Buffer|<Resultado"
       
        .ColWidth(COL_NUM) = 700
        .ColWidth(COL_F_IDIV) = 1500
        .ColWidth(COL_F_CODIV) = 1500
        .ColWidth(COL_F_DESCIV) = 4000
        .ColWidth(COL_F_IDBOD) = 1500
        .ColWidth(COL_F_CODBOD) = 1500
        .ColWidth(COL_F_DESCBOD) = 1500
        .ColWidth(COL_F_FECHA) = 1500
        .ColWidth(COL_F_FECHAN) = 1500
        .ColWidth(COL_F_RESUL) = 2500
        .ColDataType(COL_NUM) = flexDTLong
        .ColDataType(COL_F_IDIV) = flexDTLong
        .ColDataType(COL_F_CODIV) = flexDTString
        .ColDataType(COL_F_DESCIV) = flexDTString
        .ColDataType(COL_F_IDBOD) = flexDTString
        .ColDataType(COL_F_CODBOD) = flexDTString
        .ColDataType(COL_F_DESCBOD) = flexDTString
        .ColDataType(COL_F_IDIV) = flexDTDate
        .ColDataType(COL_F_RESUL) = flexDTString
        


        .ColHidden(COL_F_IDIV) = True
'        .ColHidden(COL_F_IDBOD) = True
        grd.subtotal flexSTMax, 1, 1, , grd.BackColorFixed, , True, , 1, False
        grd.subtotal flexSTAverage, 1, 4, , grd.BackColorFixed, , True, , 1, False
        grd.subtotal flexSTAverage, 1, 7, , grd.BackColorFixed, , True, , 1, False
        grd.subtotal flexSTStd, 1, 8, , grd.BackColorFixed, , True, , 1, False
        grd.subtotal flexSTMax, 1, 9, , grd.BackColorFixed, , True, , 1, False
        
        CalculaBuffer
        

    End With
End Sub

Private Sub CalculaBuffer()
Dim i As Long, sql As String, rs As Recordset, valor  As Currency
    For i = 1 To grd.Rows - 1
        If grd.IsSubtotal(i) Then
            If grd.ValueMatrix(i, 8) <> 0 Then
                valor = grd.ValueMatrix(i, 4) / grd.ValueMatrix(i, 8)
                grd.TextMatrix(i, 10) = valor
                If valor > 0.8 Then
                    grd.TextMatrix(i, 11) = Round(grd.ValueMatrix(i, 7) * 1.5, 0)
                Else
                    grd.TextMatrix(i, 11) = Round(grd.ValueMatrix(i, 9), 0)
                End If
            Else
                grd.TextMatrix(i, 11) = Round(grd.ValueMatrix(i, 7) * 1.5, 0)
            End If
            If grd.ValueMatrix(i, 11) < 0 Then
                grd.TextMatrix(i, 11) = 0
            End If

        End If
       

        
        
    Next i

End Sub

Public Sub InicioCalculoBuffer(ByVal Cadena As String)
    
    Me.tag = Cadena
    Me.tag = tag
    Me.Caption = "Actualizar Buffer"
    
    CargaPromedioVentasDiariaBufferMP3
    ConfigColsPromedioVentasDiariaBufferMP3
    Me.Show
    Me.ZOrder
End Sub

Private Sub GrabarBuffer()
    Dim sql As String, cod As String, i As Long
    Dim NumReg As Long, totalventa As Currency
    On Error GoTo ErrTrap
    
    MensajeStatus "Guardando....", 1
    With grd
        If .Rows = .FixedRows Then Exit Sub
        .ShowCell 1, 1
        
        sql = "UPDATE IvInventario " & _
              "SET buffer = 0 , FechaModBuffer=  getdate() "
              
        gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
        
        
        For i = .FixedRows To .Rows - 1

            If .IsSubtotal(i) Then
                .Row = i
                .ShowCell i, 1           'Hace visible la fila actual
                cod = Trim(Mid$(.TextMatrix(i, COL_COD), 4, Len(.TextMatrix(i, COL_COD))))
                totalventa = .ValueMatrix(i, 11)
                sql = "UPDATE IvInventario " & _
                      "SET buffer = " & totalventa & ", FechaModBuffer=  getdate() " & _
                      " WHERE idInventario= " & cod
                gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                If NumReg > 0 Then
                    .TextMatrix(i, 12) = "Actualizado..."
                Else
                    .TextMatrix(i, 12) = "Error al tratar de Actualizar..."
                End If
                .Redraw = True
                .Refresh
            End If
        Next i
    End With
    MensajeStatus "", 0
    Exit Sub
    
ErrTrap:
    MsgBox Err.Description, vbExclamation + vbOKOnly
    Exit Sub
End Sub

Public Sub InicioCalculoBufferUtilesa(ByVal Cadena As String)
    
    Me.tag = Cadena
    Me.tag = tag
    Me.Caption = "Actualizar Buffer Utilesa"
    
    CargaPromedioVentasDiariaBufferUti
    ConfigColsPromedioVentasDiariaBufferUti
    Me.Show
    Me.ZOrder
End Sub

Private Sub CargaPromedioVentasDiariaBufferUti()
    Dim sql As String, cond As String, rs As Recordset, antes As Long, NumReg As Long
    Dim objcond As Condicion, base As String, dias As Integer, mes As Currency
    Static CodAlt As String, CodBodega As String, Desc As String
    Static Recargo As String
    
    
    On Error GoTo ErrTrap
    
    antes = grd.Row
    grd.Rows = 1
    Set objcond = gobjMain.objCondicion
    'If Not (frmB_CxTrans.InicioBuffer(objcond, Recargo, "Buffer")) Then
'    If Not frmB_IV.InicioBuffer(CodAlt, Desc, CodBodega, Me.tag, mObjCond) Then
'            grd.SetFocus
'        Exit Sub
'    End If
    cond = CondicionBusquedaItemBuffer
       
    grd.Redraw = False
    MensajeStatus MSG_PREPARA, vbHourglass
    
    With mObjCond

        dias = DateDiff("d", .fecha1, .fecha2)
        mes = dias / 30
        sql = " select "
        sql = sql & " i.idinventario, max(codinventario), max(ivinventario.descripcion), max(buffer), max(TiempoReposicion), max(FrecuenciaReposicion), "
        sql = sql & " max(TiempoPromVta),  "
        sql = sql & " sum(cantidad)*-1 as cant,  "
        If .Bandera Then
            sql = sql & dias & ", round(((max(TiempoPromVta)*" & mes & ")/" & dias & ") ,4), "
            sql = sql & mes & ", round((max(TiempoPromVta)*" & mes & ") ,4), "
            sql = sql & " round(round((max(TiempoPromVta) *" & mes & ") ,4) *(max(TiempoReposicion)+ max(FrecuenciaReposicion)) * 1.5,0) "
        Else
            sql = sql & dias & ", round((sum(i.cantidad) /" & dias & ") *-1,4), "
            sql = sql & mes & ", round((sum(i.cantidad) /" & mes & ") *-1,4), "
            sql = sql & " round(round((sum(i.cantidad) /" & mes & ") *-1,4) *(max(TiempoReposicion)+ max(FrecuenciaReposicion)) * 1.5,0) "
        End If
        
        
        
        sql = sql & " from gncomprobante g inner join gntrans gnt on g.codtrans=gnt.codtrans"
        sql = sql & " inner join ivkardex i"
        sql = sql & " inner join ivinventario "
        sql = sql & " left  join ivgrupo1 on ivinventario.idgrupo1= ivgrupo1.idgrupo1 "
        sql = sql & " left join ivgrupo2 on ivinventario.idgrupo2= ivgrupo2.idgrupo2 "
        sql = sql & " left   join ivgrupo3 on ivinventario.idgrupo3= ivgrupo3.idgrupo3 "
        sql = sql & " left   join ivgrupo4 on ivinventario.idgrupo4= ivgrupo4.idgrupo4 "
        sql = sql & " left   join ivgrupo5 on ivinventario.idgrupo5= ivgrupo5.idgrupo5 "
        sql = sql & " left   join ivgrupo6 on ivinventario.idgrupo6= ivgrupo6.idgrupo6 "
        sql = sql & " on i.idinventario = ivinventario.idinventario"
        sql = sql & " on g.transid=i.transid " & cond
        sql = sql & " and Estado <> 3"
        'sql = sql & "  and codtrans in (" & .CodTrans & ")"
        sql = sql & "  and AnexoCodTipoComp in (4,18 )"
        sql = sql & "  and AnexoCodTipoTrans=2"
        sql = sql & "  and fechatrans between '" & .fecha1 & "' and '" & .fecha2 & "'"
        sql = sql & "  group by i.idinventario"
        sql = sql & " order by max(ivinventario.descripcion) "
        
        

    End With
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    grd.LoadArray MiGetRows(rs)
    Set rs = Nothing
    
    With grd
        '# de fila
        .ColAlignment(0) = flexAlignCenterCenter
        GNPoneNumFila grd, False
        'Reubica a la fila donde estaba antes
        If .Rows > antes And antes > 0 Then .Row = antes
        .Redraw = True
    End With
    MensajeStatus "", 0
    grd.SetFocus
    If grd.Rows <> grd.FixedRows Then grd.Row = grd.FixedRows
    
    
    Exit Sub

ErrTrap:
    grd.Redraw = True
    MensajeStatus "", 0
    DispErr
    Exit Sub
End Sub

Private Sub ConfigColsPromedioVentasDiariaBufferUti()
    Dim fmt As String
    
    fmt = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda("USD")
    With grd
        .FormatString = "^#|<Id Inventario|<CodInventario|<Descripcion|>Buffer Ant.|>Tiempo Rep.|>Frecuencia Rep.|>Prom. Venta Cat.|>Cant. Venta|>Num. Dias|>Prom. Vta Diaria|>Num. Meses|>Prom. Vta Mensual|>Valor para Buffer|<Resultado"
       
        .ColWidth(COL_NUM) = 700
        .ColWidth(COL_F_IDIV) = 1500
        .ColWidth(COL_F_CODIV) = 1500
        .ColWidth(COL_F_DESCIV) = 4000
        .ColWidth(COL_F_IDBOD) = 1300
        .ColWidth(COL_F_CODBOD) = 1300
        .ColWidth(COL_F_DESCBOD) = 1300
        .ColWidth(COL_F_FECHA) = 1300
        .ColWidth(COL_F_FECHAN) = 1300
        .ColWidth(COL_F_RESUL) = 1300
        
        
        
        .ColDataType(COL_NUM) = flexDTLong
        .ColDataType(COL_F_IDIV) = flexDTLong
        .ColDataType(COL_F_CODIV) = flexDTString
        .ColDataType(COL_F_DESCIV) = flexDTString
        .ColDataType(COL_F_IDBOD) = flexDTString
        .ColDataType(COL_F_CODBOD) = flexDTString
        .ColDataType(COL_F_DESCBOD) = flexDTString
        .ColDataType(COL_F_IDIV) = flexDTDate
        .ColDataType(COL_F_RESUL) = flexDTString
        
        
        .ColFormat(4) = "0"
        .ColFormat(5) = "0.00"
        .ColFormat(6) = "0.00"
        .ColFormat(7) = "0"
        .ColFormat(8) = "0"
        .ColFormat(9) = "0"
        .ColFormat(10) = "0.0000"
        .ColFormat(11) = "0.00"
        .ColFormat(12) = "0.0000"
        .ColFormat(13) = "0"
        


        .ColHidden(COL_F_IDIV) = True
        

    End With
End Sub


Private Sub GrabarBufferUtilesa()
    Dim sql As String, cod As String, i As Long
    Dim NumReg As Long, totalventa As Currency
    On Error GoTo ErrTrap
    
    MensajeStatus "Guardando....", 1
    With grd
        If .Rows = .FixedRows Then Exit Sub
        .ShowCell 1, 1
        For i = .FixedRows To .Rows - 1
            If Not .IsSubtotal(i) Then
                .Row = i
                .ShowCell i, 1           'Hace visible la fila actual
                cod = .ValueMatrix(i, COL_COD)
                totalventa = .ValueMatrix(i, 13)
                sql = "UPDATE IvInventario " & _
                      "SET buffer = " & totalventa & ", FechaModBuffer=  getdate() " & _
                      " WHERE idInventario= " & cod
                gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                If NumReg > 0 Then
                    .TextMatrix(i, 14) = "Actualizado..."
                Else
                    .TextMatrix(i, 14) = "Error al tratar de Actualizar..."
                End If
                .Redraw = True
                .Refresh
            End If
        Next i
    End With
    MensajeStatus "", 0
    Exit Sub
    
ErrTrap:
    MsgBox Err.Description, vbExclamation + vbOKOnly
    Exit Sub
End Sub


Private Function CondicionBusquedaItemBuffer() As String
    
    Static CodAlt As String, CodBodega As String, Desc As String
    Dim cond As String, Bandfirst As Boolean, comodin As String
    
#If DAOLIB Then
    comodin = "*"       'DAO
#Else
    comodin = "%"       'ADO
#End If
   

    If Not frmB_IV.InicioBuffer(CodAlt, Desc, CodBodega, Me.tag, mObjCond) Then
        CondicionBusquedaItemBuffer = ""
        Exit Function
    End If
    Bandfirst = True
    With mObjCond
        If Len(.Item1) > 0 Then
            cond = cond & " (codInventario LIKE '" & .Item1 & comodin & "')"
            Bandfirst = False
        End If
        If Len(CodAlt) > 0 Then
            If Bandfirst = False Then cond = cond & " AND "
            cond = cond & " (codAlterno1 LIKE '" & CodAlt & comodin & "')"
            Bandfirst = False
        End If
        If Len(Desc) > 0 Then
            If Bandfirst = False Then cond = cond & " AND "
            cond = cond & " (IVInventario.Descripcion LIKE '" & Desc & comodin & "')"
            Bandfirst = False
        End If
        If Len(CodBodega) > 0 Then
            If Bandfirst = False Then cond = cond & " AND "
            cond = cond & " (IVBodega.CodBodega ='" & CodBodega & "')"
            Bandfirst = False
        End If
        
        If Not .Bandera2 Then   'esta activado el filtro avanzaado de grupos
           If (Len(.Grupo1) > 0) Or (Len(.Grupo2) > 0) Then
                If Bandfirst = False Then cond = cond & " AND "
                cond = cond & " (IVGrupo" & .numGrupo & ".CodGrupo" & _
                       CStr(.numGrupo) & " BETWEEN '" & .Grupo1 & "' AND '" & .Grupo2 & "')"
                Bandfirst = False
            End If
        Else
            'Condiciones de busqueda de grupos segun filtro avanzado
            If Len(.CodGrupo1) > 0 Then
                If Bandfirst = False Then cond = cond & " AND "
                cond = cond & " (IVGrupo1.CodGrupo1 = '" & .CodGrupo1 & "')"
                Bandfirst = False
            End If
            
            If Len(.CodGrupo2) > 0 Then
                If Bandfirst = False Then cond = cond & " AND "
                cond = cond & " (IVGrupo2.CodGrupo2 = '" & .CodGrupo2 & "')"
                Bandfirst = False
            End If
            
            If Len(.CodGrupo3) > 0 Then
                If Bandfirst = False Then cond = cond & " AND "
                cond = cond & " (IVGrupo3.CodGrupo3 = '" & .CodGrupo3 & "')"
                Bandfirst = False
            End If
            
            If Len(.CodGrupo4) > 0 Then
                If Bandfirst = False Then cond = cond & " AND "
                cond = cond & " (IVGrupo4.CodGrupo4 = '" & .CodGrupo4 & "')"
                Bandfirst = False
            End If
            
            If Len(.CodGrupo5) > 0 Then
                If Bandfirst = False Then cond = cond & " AND "
                cond = cond & " (IVGrupo5.CodGrupo5 = '" & .CodGrupo5 & "')"
                Bandfirst = False
            End If
        End If
        
       If Me.tag = "ExisMin" Then
            If .Bandera = False Then
                If Not Bandfirst Then cond = cond & " AND "
                cond = cond & " (IVExist.Exist>0) "
            End If
        End If
        If Me.tag = "Exis" Then
            'If .Bandera = False Then
            If Not Bandfirst Then cond = cond & " AND "
            cond = cond & " GNComprobante.FechaTrans <= " & FechaYMD(mObjCond.Fcorte, gobjMain.EmpresaActual.TipoDB) & " "
            'End If
        End If
    End With
    If Bandfirst = False Then cond = " WHERE " & cond
    CondicionBusquedaItemBuffer = cond
End Function


Public Sub InicioCalculoBufferxAlmacen(ByVal Cadena As String)
    
    Me.tag = Cadena
    Me.tag = tag
    Me.Caption = "Actualizar Buffer x Almacen"
    
    CargaPromedioVentasDiariaBufferMP3xAlma
    ConfigColsPromedioVentasDiariaBufferMP3Alma
    Me.Show
    Me.ZOrder
End Sub

Private Sub CargaPromedioVentasDiariaBufferMP3xAlma()
    Dim sql As String, cond As String, rs As Recordset, antes As Long, NumReg As Long
    Dim objcond As Condicion, base As String
    Static Recargo As String
    
    
    On Error GoTo ErrTrap
    
    antes = grd.Row
    grd.Rows = 1
    Set objcond = gobjMain.objCondicion
    If Not (frmB_CxTrans.InicioActualizaBufferxEmpresa(objcond, Recargo, "BufferxAlm_CodBodega")) Then
        grd.SetFocus
        Exit Sub
    End If
       
    grd.Redraw = False
    MensajeStatus MSG_PREPARA, vbHourglass
    
    With objcond
        base = objcond.Sucursal & ".dbo."
        
        sql = " select idinventario, codbodega, descripcion, CANT, TiempoRep, FECHATRANS, aCUMULADO, cant,aCUMULADO,0 as prom from vwVentaxItemxFechaALM"
        'sql = sql & " Where idinventario = 54298"
        sql = sql & " Where 1=1 "
        sql = sql & "  and fechatrans between '" & Date - 180 & "' and '" & Date & "'"
        sql = sql & "  and codbodega in (" & objcond.CodTrans & ")"
        sql = sql & " order by descripcion, fechatrans"
        
        

    End With
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    grd.LoadArray MiGetRows(rs)
    Set rs = Nothing
    
    With grd
        '# de fila
        .ColAlignment(0) = flexAlignCenterCenter
        GNPoneNumFila grd, False
        'Reubica a la fila donde estaba antes
        If .Rows > antes And antes > 0 Then .Row = antes
        .Redraw = True
    End With
    MensajeStatus "", 0
    grd.SetFocus
    If grd.Rows <> grd.FixedRows Then grd.Row = grd.FixedRows
    
    
    Exit Sub

ErrTrap:
    grd.Redraw = True
    MensajeStatus "", 0
    DispErr
    Exit Sub
End Sub

Private Sub ConfigColsPromedioVentasDiariaBufferMP3Alma()
    Dim fmt As String
    
    fmt = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda("USD")
    With grd
        .FormatString = "^#|<Id Inventario|<Cod Bodega|<Descripcion|>Venta|>T.Repo|<Fecha Trans |>Acumulado|>Desv. Estandar|>Max.Prom|>Calculo para Buffer|>Valor para Buffer|<Resultado"
       
        .ColWidth(COL_NUM) = 700
        .ColWidth(COL_F_IDIV) = 1500
        .ColWidth(COL_F_CODIV) = 1500
        .ColWidth(COL_F_DESCIV) = 4000
        .ColWidth(COL_F_IDBOD) = 1500
        .ColWidth(COL_F_CODBOD) = 1500
        .ColWidth(COL_F_DESCBOD) = 1500
        .ColWidth(COL_F_FECHA) = 1500
        .ColWidth(COL_F_FECHAN) = 1500
        .ColWidth(COL_F_RESUL) = 2500
        .ColDataType(COL_NUM) = flexDTLong
        .ColDataType(COL_F_IDIV) = flexDTLong
        .ColDataType(COL_F_CODIV) = flexDTString
        .ColDataType(COL_F_DESCIV) = flexDTString
        .ColDataType(COL_F_IDBOD) = flexDTString
        .ColDataType(COL_F_CODBOD) = flexDTString
        .ColDataType(COL_F_DESCBOD) = flexDTString
        .ColDataType(COL_F_IDIV) = flexDTDate
        .ColDataType(COL_F_RESUL) = flexDTString
        


        .ColHidden(COL_F_IDIV) = True
'        .ColHidden(COL_F_IDBOD) = True
        grd.subtotal flexSTMax, 1, 1, , grd.BackColorFixed, , True, , 1, False
        grd.subtotal flexSTAverage, 1, 4, , grd.BackColorFixed, , True, , 1, False
        grd.subtotal flexSTAverage, 1, 7, , grd.BackColorFixed, , True, , 1, False
        grd.subtotal flexSTStd, 1, 8, , grd.BackColorFixed, , True, , 1, False
        grd.subtotal flexSTMax, 1, 9, , grd.BackColorFixed, , True, , 1, False
        
        CalculaBufferxAlma
        

    End With
End Sub

Private Sub CalculaBufferxAlma()
Dim i As Long, sql As String, rs As Recordset, valor  As Currency
    For i = 1 To grd.Rows - 1
        If grd.IsSubtotal(i) Then
            If grd.ValueMatrix(i, 8) <> 0 Then
                valor = grd.ValueMatrix(i, 4) / grd.ValueMatrix(i, 8)
                grd.TextMatrix(i, 10) = valor
                If valor > 0.8 Then
                    grd.TextMatrix(i, 11) = Round(grd.ValueMatrix(i, 7) * 1.5, 0)
                Else
                    grd.TextMatrix(i, 11) = Round(grd.ValueMatrix(i, 9), 0)
                End If
            Else
                grd.TextMatrix(i, 11) = Round(grd.ValueMatrix(i, 7) * 1.5, 0)
            End If
            If grd.ValueMatrix(i, 11) < 0 Then
                grd.TextMatrix(i, 11) = 0
            End If

        End If
       

        
        
    Next i

End Sub


Private Sub GrabarBufferxALM()
    Dim sql As String, cod As String, i As Long
    Dim NumReg As Long, totalventa As Currency
    On Error GoTo ErrTrap
    
    MensajeStatus "Guardando....", 1
    With grd
        If .Rows = .FixedRows Then Exit Sub
        .ShowCell 1, 1
        
        
        
        For i = .FixedRows To .Rows - 1

            If .IsSubtotal(i) Then
                .Row = i
                .ShowCell i, 1           'Hace visible la fila actual
                cod = Trim(Mid$(.TextMatrix(i, COL_COD), 4, Len(.TextMatrix(i, COL_COD))))
                totalventa = .ValueMatrix(i, 11)
                sql = "UPDATE Ivexist "
                sql = sql & " SET buffer = " & totalventa
                sql = sql & " from ivexist inner join ivbodega on ivexist.idbodega = ivbodega.idbodega"
                sql = sql & " WHERE idInventario= " & cod
                sql = sql & " and codbodega='" & .TextMatrix(i - 1, 2) & "'"
                gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                If NumReg > 0 Then
                    .TextMatrix(i, 12) = "Actualizado..."
                Else
                    .TextMatrix(i, 12) = "Error al tratar de Actualizar..."
                End If
                .Redraw = True
                .Refresh
            End If
        Next i
    End With
    MensajeStatus "", 0
    Exit Sub
    
ErrTrap:
    MsgBox Err.Description, vbExclamation + vbOKOnly
    Exit Sub
End Sub


Public Sub InicioAsignaDescuentoxItem(ByVal Cadena As String)
    Me.tag = Cadena
    Me.Caption = "Actualizar Descuento x Cliente x Item"
    CargaVentasxItemxCliente
    ConfigColsDescxClixItem
    Me.Show
    Me.ZOrder
End Sub

Private Sub CargaVentasxItemxCliente()
    Dim sql As String, cond As String, rs As Recordset, antes As Long
    Dim objcond As Condicion
    Static Recargo As String
    
    On Error GoTo ErrTrap
    
    antes = grd.Row
    grd.Rows = 1
    Set objcond = gobjMain.objCondicion
    If Not (frmB_VxTrans.InicioVxMesTransaccion(objcond, Recargo, "PromedioVentas")) Then
        grd.SetFocus
        Exit Sub
    End If
       
    grd.Redraw = False
    MensajeStatus MSG_PREPARA, vbHourglass
    
    With objcond
        
        cond = " AND g.FechaTrans between " & FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB)
        cond = cond & " AND  " & FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB)
        
        
        sql = " select"
        sql = sql & " max (fechatrans) as fechatrans,"
        sql = sql & " fc.nombre as nomvend,"
        sql = sql & " pc.codprovcli as ruc, pc.nombre, ivi.codinventario, ivi.descripcion, pc.idprovcli, ivi.idinventario, "
        sql = sql & " 1- round((i.costorealtotal/i.preciorealtotal),4) as rentab_real,"
''        sql = sql & " CASE WHEN i.preciorealtotal <>0 THEN"
''        sql = sql & " 1- round((i.costorealtotal/i.preciorealtotal),4) "
''        sql = sql & " ELSE"
''        sql = sql & " 1- round((i.costorealtotal/1),4) END as rentab_real,"

        sql = sql & " ivi.precio5,"
        sql = sql & " round(i.preciorealtotal/cantidad,4) as pu_real ,"
        sql = sql & " round(i.costorealtotal/cantidad,4) as cu_real ,"
        sql = sql & " CASE WHEN precio5 <>0 THEN"
        sql = sql & " 1-round(( (i.preciorealtotal/cantidad)/precio5),4) "
        sql = sql & " ELSE"
        sql = sql & " 1-round(( (i.preciorealtotal/cantidad)/1),4) END as descto"
        sql = sql & " into tmp0"
        sql = sql & " from gncomprobante g"
        sql = sql & " inner join ivkardex i"
        sql = sql & " inner join ivinventario ivi"
        sql = sql & " on i.idinventario= ivi.idinventario"
        sql = sql & " on g.transid=i.transid"
        sql = sql & " inner join pcprovcli pc"
        sql = sql & " inner join fcvendedor fc on PC.idvendedor= fc.idvendedor"
        sql = sql & " on pc.idprovcli=g.idclienteref"
        sql = sql & " where g.estado<>3 "
        sql = sql & " AND G.CodTrans IN (" & PreparaCadena(.CodTrans) & ")" & cond
        sql = sql & " group by fc.nombre,  pc.nombre, ivi.codinventario, ivi.descripcion , ivi.precio5 , pc.idprovcli, ivi.idinventario, pc.codprovcli, "
        sql = sql & " (1- round((i.costorealtotal/i.preciorealtotal),4))"
        sql = sql & " , round(i.preciorealtotal/cantidad,4),"
        sql = sql & " round(i.costorealtotal/cantidad,4) , i.preciorealtotal, cantidad"
        sql = sql & " order by pc.nombre,  max (fechatrans) desc"
        
        VerificaExistenciaTabla 0
        gobjMain.EmpresaActual.EjecutarSQL sql, 0
    
    
        sql = " select max(fechatrans) as fechatrans ,codinventario  ,NOMBRE"
        sql = sql & " Into tmp1"
        sql = sql & " from tmp0"
        sql = sql & " Group By CodInventario , nombre"
    
        VerificaExistenciaTabla 1
        gobjMain.EmpresaActual.EjecutarSQL sql, 0
    
    
        sql = " select "
        sql = sql & " i.idprovcli, i.ruc, i.nombre, i.idinventario, i. codinventario, i.descripcion, descto*100"
        sql = sql & " from tmp1 v "
        sql = sql & " inner join  tmp0 i on v.fechatrans=i.fechatrans and v.codinventario = i.codinventario AND V.NOMBRE=I.NOMBRE"
        sql = sql & " WHERE descto>0 "
        sql = sql & " order by I.NOMBRE, v.codinventario"
    
    End With
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    grd.LoadArray MiGetRows(rs)
    Set rs = Nothing
    
    With grd
        '# de fila
        .ColAlignment(0) = flexAlignCenterCenter
        GNPoneNumFila grd, False
        'Reubica a la fila donde estaba antes
        If .Rows > antes And antes > 0 Then .Row = antes
        .Redraw = True
    End With
    MensajeStatus "", 0
    grd.SetFocus
    If grd.Rows <> grd.FixedRows Then grd.Row = grd.FixedRows
    Exit Sub

ErrTrap:
    grd.Redraw = True
    MensajeStatus "", 0
    DispErr
    Exit Sub
End Sub



Private Sub ConfigColsDescxClixItem()
    Dim fmt As String
    Dim i As Integer
    
    fmt = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda("USD")
    With grd
            .FormatString = ">#<|<idcliente|<RUC|<Nombre|<IdItem|<Codigo Item|<Descripcion Item|>Descuento |<Resultado"
        
        .ColWidth(COL_NUM) = 700
        .ColWidth(COL_DCI_RUC) = 1500
        .ColWidth(COL_DCI_NOMCLI) = 3500
        .ColWidth(COL_DCI_CODIV) = 1500
        .ColWidth(COL_DCI_DESCIV) = 3500
        .ColWidth(COL_DCI_DESC) = 1500
        
        .ColHidden(COL_DCI_IDCLI) = True
        .ColHidden(COL_DCI_IDIV) = True
        
        AsignarTituloAColKey grd
       ' If Me.tag = "VxTransProm" Then
         grd.Editable = flexEDNone
            For i = 0 To .Cols - 2 '.ColIndex("Límite de Crédito") '- 1
                .ColData(i) = -1
            Next i
            
            'Color de fondo
            If .Rows > .FixedRows Then
                .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 2) = .BackColorFrozen '.ColIndex("Límite de Crédito")) = .BackColorFrozen
            End If
        'End If

        .ColDataType(COL_NUM) = flexDTLong
        .ColDataType(COL_DCI_RUC) = flexDTString
        .ColDataType(COL_DCI_NOMCLI) = flexDTString
        .ColDataType(COL_DCI_CODIV) = flexDTString
        .ColDataType(COL_DCI_DESCIV) = flexDTString
        .ColDataType(COL_DCI_DESC) = flexDTCurrency
        .ColDataType(COL_RES) = flexDTString
        .ColFormat(COL_DCI_DESC) = fmt
        
        .MergeCol(1) = True
        .MergeCol(2) = True
        
        grd.subtotal flexSTCount, 1, COL_DCI_DESC, , grd.BackColorFixed, , True, " ", 1, False
        
    End With
End Sub

Private Sub GrabarDescuentoxClixItem()
    Dim sql As String, idivi As String, i As Long
    Dim idpc As String, cont As Integer
    Dim NumReg As Long, totalventa As Currency
    Dim rs As Recordset, fecha As Date
    Dim sql1 As String
    Dim FilaIni As Long, FilaFin As Long, j As Long
    Dim Desc As IVDescuento, ixItem As Long, ixCli As Long, ixFC As Long
    
    On Error GoTo ErrTrap
    
    
    sql = "Delete ivdescuentodetallefc"
    gobjMain.EmpresaActual.EjecutarSQL sql, 0
    sql = "Delete ivdescuentodetallepc"
    gobjMain.EmpresaActual.EjecutarSQL sql, 0
    sql = "Delete ivdescuentodetalleiv"
    gobjMain.EmpresaActual.EjecutarSQL sql, 0
    sql = "Delete ivdescuento"
    gobjMain.EmpresaActual.EjecutarSQL sql, 0
    
    MensajeStatus "Guardando....", 1
    With grd
        If .Rows = .FixedRows Then Exit Sub
        .ShowCell 1, 1
        idivi = ""
        idpc = ""
        FilaIni = 2
        For i = .FixedRows + 1 To .Rows - 1
            If .IsSubtotal(i) Then
                FilaFin = i - 1
                Set Desc = gobjMain.EmpresaActual.CreaIVDescuento
                Desc.CodDescuento = grd.TextMatrix(FilaIni, COL_DCI_RUC)
                Desc.Descripcion = grd.TextMatrix(FilaIni, COL_DCI_NOMCLI)
                Desc.BandValida = True
                Desc.BandxCliente = True
                Desc.BandxItem = True
                ixCli = Desc.AddDetalleDescuentoPC
                Desc.IVDescuentoDetallePC(ixCli).CodProvCli = grd.TextMatrix(FilaIni, COL_DCI_RUC)
                For j = FilaIni To FilaFin
                    ixItem = Desc.AddDetalleDescuentoIV
                    Desc.IVDescuentoDetalleIV(ixItem).CodInventario = grd.TextMatrix(j, COL_DCI_CODIV)
                    Desc.IVDescuentoDetalleIV(ixItem).Descuento = grd.ValueMatrix(j, COL_DCI_DESC)
                    .Row = j
                    .ShowCell j, 1           'Hace visible la fila actual
                Next j
                
                ixFC = Desc.AddDetalleDescuentoFC
                Desc.IVDescuentoDetalleFC(ixFC).codforma = "CRC"
                
                ixFC = Desc.AddDetalleDescuentoFC
                Desc.IVDescuentoDetalleFC(ixFC).codforma = "EFECT"
                    
                ixFC = Desc.AddDetalleDescuentoFC
                Desc.IVDescuentoDetalleFC(ixFC).codforma = "CHPF"
                    
                
                Desc.Grabar
                FilaIni = i + 1
                .Redraw = True
                .Refresh
            End If
        Next i
    End With
    MensajeStatus "", 0
    Exit Sub
    
ErrTrap:
    MsgBox Err.Description, vbExclamation + vbOKOnly
    Exit Sub
End Sub

Public Sub InicioVentasxItemxSuc(ByVal Cadena As String)
    Dim objcond As Condicion
    Static Recargo As String
    Me.tag = Cadena
    Me.Caption = "Clasificar Item"
    picSucursal.Visible = True
     fcbSucursal.SetData gobjMain.EmpresaActual.ListaGNSucursales(True, False) 'jeaa 10/09/2008
'    ConfigColsVtaxItemxSuc
    'CargaVentasxItemxSucursal
    grd.Rows = 1
    ConfigColsVtaxItemxSuc
    Set objcond = gobjMain.objCondicion
        If Not (frmB_VxTrans.InicioVxSucursal(objcond, Recargo, "VentasxSuc")) Then
            grd.SetFocus
            Exit Sub
        End If
        fcbSucursal.Enabled = False
    Me.Show
    Me.ZOrder
End Sub

Private Sub CargaVentasxItemxSucursal()
    Dim sql As String, cond As String, rs As Recordset, antes As Long
    Dim objcond As Condicion
    Dim rsbod As Recordset
    Dim sqlBod As String
    Dim NumReg As Long
    Dim W As Variant, s As String, i As Long, IdBod As Long
    Static Recargo As String
    Dim ivb  As IVBodega, bod As String, NEWordenbod As String, CodigoBodega As String, COLbODEGA As String, sqlbodegas As String, sqltotal As String
    Dim numbodega As Long
    On Error GoTo ErrTrap
    antes = grd.Row
    grd.Rows = 1
    Set objcond = gobjMain.objCondicion
    If Not BandTodo Then
        If Not (frmB_VxTrans.InicioVxSucursal(objcond, Recargo, "VentasxSuc")) Then
            grd.SetFocus
            Exit Sub
        End If
        chkTodo.value = Checked
    End If
    grd.Redraw = False
    MensajeStatus MSG_PREPARA, vbHourglass
    VerificaExistenciaTabla 1
    With objcond
            cond = " AND g.FechaTrans between " & FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB)
            cond = cond & " AND  " & FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB)
'            cond = ""
            cond = cond & " AND gnt.anexocodtipotrans IN (2)"
            cond = cond & " AND gnt.anexocodtipoComp IN (4,18)"
            If chkTodo.value = vbUnchecked Then
                If Len(fcbSucursal.KeyText) = 0 Then
                    MsgBox "Escoja Una sucursal": fcbSucursal.SetFocus: Exit Sub
                Else
                    .Sucursal = fcbSucursal.KeyText
                End If
            Else
                 .Sucursal = ""
            End If
            If Len(.Sucursal) > 0 Then
                cond = cond & " AND gsuc.codsucursal IN (" & PreparaCadena(.Sucursal) & ")"
            End If
            If Not (Len(.Bienes)) = 0 Then
                cond = cond & .Bienes   'Aquí se ha grabado SQL de ítems
            End If
            'aqui saca ventas proyectadas 24 semanas
            sql = "  select vw.idinventario, "
            sql = sql & "   sum(totalM3) as prom,"
            sql = sql & "   sum(totalCANT) as promC"
            sql = sql & " Into TMPvsFin"
            sql = sql & " from vwVentasItemxMesesxSucursal vw inner join ivinventario  on vw.idinventario=ivinventario.idinventario"
            sql = sql & " Left JOIN GNSucursal gsuc on gsuc.idsucursal = vw.idsucursal "
            sql = sql & " where fechatrans between  (" & FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = sql & " and (" & FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ")"
            If Len(.Sucursal) > 0 Then
                sql = sql & " AND gsuc.codsucursal IN (" & PreparaCadena(.Sucursal) & ")"
            End If
            If Not (Len(.Bienes)) = 0 Then
                sql = sql & .Bienes  'Aquí se ha grabado SQL de ítems
            End If
            'sql = sql & " and  ivinventario.codinventario =  'filo gri 29mm 2m pvc'"
            sql = sql & " group by vw.idinventario"
            VerificaExistenciaTablaTemp "TMPvsFin"
            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
            'aqui saca ventas proyectadas 8 semanas
            sql = "  select vw.idinventario, "
            sql = sql & "   sum(totalM3) as prom, "
            sql = sql & "   sum(totalCANT) as promC"
            sql = sql & " Into TMPvs8sem"
            sql = sql & " From vwVentasItemxmesesxSucursal vw inner join ivinventario  on vw.idinventario=ivinventario.idinventario"
            sql = sql & " Left JOIN GNSucursal gsuc on gsuc.idsucursal = vw.idsucursal "
            sql = sql & " where Fechatrans between    Dateadd(m, -2, " & FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ")"
            sql = sql & " and (" & FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ")"
 
            If Len(.Sucursal) > 0 Then
                sql = sql & " AND gsuc.codsucursal IN (" & PreparaCadena(.Sucursal) & ")"
            End If
            If Not (Len(.Bienes)) = 0 Then
                sql = sql & .Bienes  'Aquí se ha grabado SQL de ítems
            End If
 '           sql = sql & " and  ivinventario.codinventario =  'DESC-MDPPANEL 7X8'"
            sql = sql & " group by vw.idinventario"
            VerificaExistenciaTablaTemp "TMPvs8sem"
            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
            '-----------------------------------
                If Len(.Sucursal) > 0 Then
                    VerificaExistenciaTablaTemp "TmpSuc"
                    sql = "select ive.idinventario,ive.idgrupo"
                    sql = sql & " into TmpSuc"
                    sql = sql & " from ivexist ive"
                    sql = sql & " Inner join ivbodega ivb on ivb.idbodega = ive.idbodega"
                    sql = sql & " Inner join ivinventario on ive.idinventario = ivinventario.idInventario"
                    sql = sql & " Where  ivb.idsucursal = (select idsucursal from gnsucursal where codsucursal IN (" & PreparaCadena(.Sucursal) & "))"
                    If Not (Len(.Bienes)) = 0 Then
                        sql = sql & .Bienes 'Aquí se ha grabado SQL de ítems
                    End If
                    sql = sql & " group by ive.idinventario,ive.idgrupo"
                    gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                End If
            '-----------------------------------
            sql = " select"
            sql = sql & " ivk.idinventario, ivinventario.codinventario,  ivinventario.DESCRIPCION +'     [' +  convert(varchar,ivinventario.m3) +']' as descripcion, "
            If Len(.Sucursal) > 0 Then
                sql = sql & " GSUC.CODSUCURSAL, "
            Else
                sql = sql & " 'TODAS' as CODSUCURSAL, "
            End If
                sql = sql & " SUM(CANTIDAD)*-1 AS CANT,"
                sql = sql & " sum(preciorealtotal)*-1 as pt, "
                sql = sql & " SUM(CANTIDAD*M3)*-1 AS TOTALM3, "
                If Len(.Sucursal) > 0 Then
                    sql = sql & " ivg.IdealDias  as TiempoPromVta,'' as clasf, '' as result, ivg.codgrupo4, '' as result1,'' as result2," 'vwex.exist *m3 as existm3,  vwex.exist ,"
                Else
                    sql = sql & " TiempoPromVta,'' as clasf, '' as result, codgrupo4, '' as result1,'' as result2," 'vwex.exist *m3 as existm3,  vwex.exist ,"
                End If
                'aqui pongo todas las bodegas
                s = ""
                sqlbodegas = ""
                '-----------------------------------
                sqlBod = "Select ivb.idbodega from ivbodega ivb Inner join gnsucursal gns on gns.idsucursal = ivb.idsucursal where ivb.bandvalida = 1 "
                If Len(.Sucursal) > 0 Then
                    sqlBod = sqlBod & " And gns.codsucursal = '" & .Sucursal & "'"
                End If
                sqlBod = sqlBod & "Order by gns.codSucursal"
                Set rsbod = gobjMain.EmpresaActual.OpenRecordset(sqlBod)
                If Not rsbod Is Nothing Then
                        Do While Not rsbod.EOF
                            s = s & rsbod!IdBodega & ","
                            rsbod.MoveNext
                        Loop
                    End If
                    Set rsbod = Nothing
                      W = Split(Left(s, Len(s) - 1), ",")
                    For i = 0 To UBound(W)
                        IdBod = W(i)
                        sqltotal = sqltotal & " IsNull(vw" & i & ".Exist ,0)+"
                        sqlbodegas = sqlbodegas & " vw" & i & ".Exist as exist" & i & ", "
                    Next i
            If Len(sqlbodegas) > 0 Then
                sql = sql & "( " & Left(sqltotal, Len(sqltotal) - 1) & ") *m3 as existm3,"
                sql = sql & Left(sqltotal, Len(sqltotal) - 1) & "as exist,"
                sql = sql & sqlbodegas
            End If
                '-----------------------------------
                sql = sql & " 0 as cu , 0 as CT "
                sql = sql & " Into tmp1 from gncomprobante g inner join gntrans gnt "
            If Len(.Sucursal) > 0 Then
                sql = sql & " Left JOIN GNSucursal gsuc on gsuc.idsucursal = GNT.idsucursal "
            End If
                sql = sql & " on g.codtrans=gnt.codtrans"
                sql = sql & " inner join ivkardex ivk"
                sql = sql & " inner join VWivexistTodo vwex "
                For i = 0 To UBound(W)
                    IdBod = W(i) 'Mid$(W(I), 2, Len(W(I)) - 2)
                    sql = sql & " LEFT JOIN vwExistencia" & IdBod & " vw" & i & " on vwex.idinventario = vw" & i & ".idinventario"
                Next i
                sql = sql & "   on ivk.idinventario= vwex.idinventario"
                sql = sql & " inner join ivinventario  "
                If Len(.Sucursal) > 0 Then
                    sql = sql & " Inner join tmpsuc t Inner join ivgrupo4 ivg on ivg.idgrupo4= t.idgrupo on t.idinventario =ivinventario.idinventario"
                End If
                sql = sql & " on ivinventario.idinventario = ivk.idinventario"
                sql = sql & " on g.transid=ivk.transid"
                sql = sql & " Where estado <>3 and bandservicio = 0 "
               sql = sql & cond
                sql = sql & " Group by ivk.idinventario, ivinventario.codinventario, codalterno1, M3, ivinventario.DESCRIPCION,"
                If Len(.Sucursal) > 0 Then
                    sql = sql & " ivg.idealdias, vwex.exist, ivg.codgrupo4 "
                    For i = 0 To UBound(W)
                        IdBod = W(i)
                            sql = sql & ",vw" & i & ".Exist "
                    Next i
                Else
                    sql = sql & " TiempoPromVta, vwex.exist, codgrupo4 "
                    For i = 0 To UBound(W)
                        IdBod = W(i)
                            sql = sql & ",vw" & i & ".Exist "
                    Next i
                End If
             '   sql = sql & " "
                If Len(.Sucursal) > 0 Then
                    sql = sql & " , GSUC.CODSUCURSAL "
                End If
                If .BandTodo Then
                    sql = sql & "  order by SUM(CANTIDAD*M3)*-1  desc"
                Else
                    sql = sql & "  order by SUM(preciorealtotal)*-1  desc"
                End If
    End With
        gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
        'sacamos los q no tienen movimientos
        VerificaExistenciaTabla 2
        sql = "Select  * into tmp2 from ivinventario "
        sql = sql & "  Where bandservicio=0  And idinventario not in (select idinventario from tmp1) "
        'sql = sql & "  Where idinventario not in (select idinventario from tmp1) "
           If Not (Len(objcond.Bienes)) = 0 Then
                sql = sql & objcond.Bienes    'Aquí se ha grabado SQL de ítems
            End If
        gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
        sqltotal = ""
        sqlbodegas = ""
        sql = "Select t.idinventario,ivg1.descripcion,t.codinventario,t.DESCRIPCION,t.CODSUCURSAL,t.CANT,t.TOTALM3,t.pt,t.clasf,t.result,t.codgrupo4,t.TiempoPromVta,t.result1,t.result2,t.existm3,t.exist, "
        sql = sql & "(isnull((tf.prom / 6),0) * 0.3  + isnull((tf8.prom / 2),0)* 0.7)  as vtasemproyM3, "
       sql = sql & "case when isnull(tf.prom,0) =0 And isnull(tf8.prom,0)=0 then  10000 else "
        sql = sql & "(existm3/ ((isnull((tf.prom / 6),0) * 0.3  + isnull((tf8.prom / 2),0)* 0.7) ))*30 end as diasStock"
       For i = 0 To UBound(W)
            IdBod = W(i)
            sql = sql & ",IsNull(t.Exist" & i & ",0)  as Exist" & i
        Next i
        sql = sql & ",0 as cu , 0 as CT"
        sql = sql & " From tmp1 t "
        sql = sql & " Left join  TMPvsFin tf on tf.idinventario = t.idinventario  "
        sql = sql & " Left join  TMPvs8sem tf8 on tf8.idinventario = t.idinventario"
        sql = sql & " Left join  ivinventario iv left Join ivgrupo1 ivg1 on ivg1.idgrupo1 = iv.idgrupo1 on iv.idinventario = t.idinventario"
        sql = sql & " Union all "
        sql = sql & " Select t.idinventario,ivg1.descripcion,t.codinventario,t.DESCRIPCION +'     [' +  convert(varchar,iv.m3) +']' as descripcion "
        If Len(objcond.Sucursal) > 0 Then
                sql = sql & ",'" & fcbSucursal.KeyText & "'as CODSUCURSAL, 0 as CANT,0 as TOTALM3,0 as pt"
                sql = sql & ", ''as clasf,'' as result,'' as codgrupo4,0 as TiempoPromVta, '' as result1,'' as  result2,"
        Else
                sql = sql & ",'TODAS'as CODSUCURSAL, 0 as CANT,0 as TOTALM3,0 as pt"
                sql = sql & ", ''as clasf,'' as result, ivg4.codgrupo4,t.TiempoPromVta, '' as result1,'' as  result2,"
        End If
        'sql = sql & "   0  as existm3,0 as exist"
        For i = 0 To UBound(W)
            IdBod = W(i)
            sqltotal = sqltotal & " IsNull(vw" & i & ".Exist ,0)+"
            sqlbodegas = sqlbodegas & " vw" & i & ".Exist as exist" & i & ", "
        Next i
        If Len(sqlbodegas) > 0 Then
            sql = sql & "( " & Left(sqltotal, Len(sqltotal) - 1) & ") *t.m3 as existm3,"
            sql = sql & Left(sqltotal, Len(sqltotal) - 1) & "as exist,"
            
            sql = sql & "0 as vtasemproym3,"
            sql = sql & " case when " & Left(sqltotal, Len(sqltotal) - 1) & " <>0 then 10000 else 0 end as diasStock,"
             
            sql = sql & sqlbodegas
        End If
        sql = sql & " 0 as cu , 0 as CT"
        sql = sql & " from tmp2 t Left Join IVGrupo4  ivg4 On ivg4.idgrupo4 = t.idgrupo4 "
                For i = 0 To UBound(W)
                    IdBod = W(i) 'Mid$(W(I), 2, Len(W(I)) - 2)
                    sql = sql & " LEFT JOIN vwExistencia" & IdBod & " vw" & i & " on t.idinventario = vw" & i & ".idinventario"
                Next i
                sql = sql & " Left join  ivinventario iv left Join ivgrupo1 ivg1 on ivg1.idgrupo1 = iv.idgrupo1 on iv.idinventario = t.idinventario"
        sql = sql & " order by pt desc "
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    grd.LoadArray MiGetRows(rs)
    Set rs = Nothing
    With grd
        '# de fila
        .ColAlignment(0) = flexAlignCenterCenter
        GNPoneNumFila grd, False
        'Reubica a la fila donde estaba antes
        If .Rows > antes And antes > 0 Then .Row = antes
        .Redraw = True
    End With
    MensajeStatus "", 0
    grd.SetFocus
    If grd.Rows <> grd.FixedRows Then grd.Row = grd.FixedRows
    Exit Sub
ErrTrap:
    grd.Redraw = True
    MensajeStatus "", 0
    DispErr
    Exit Sub
End Sub

Private Sub ConfigColsVtaxItemxSuc()
    Dim fmt As String, c As Currency, item As IVinventario
    Dim i As Integer, Total As Currency
    Dim sql As String, rs As Recordset
    Dim s As String, W As Variant
    fmt = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda("USD")
    With grd
        s = ">#|<idInventario|<Familia|<Codigo|<Descripción|<Sucursal|>Cant.Uni.Ventas|>Total M3|>Valor USD Real|>Porcen|>Participa Vtas|^Cod grupo|^Politica Dias Stock|^Clasificación|<Resultado|>Exist M3|>Exist Uni|>VentasProyM3|>DiasStockReal"
        If chkTodo.value = vbChecked Then
                    sql = "Select ivb.codbodega from ivbodega ivb Inner join gnsucursal gns on gns.idsucursal = ivb.idsucursal where ivb.bandvalida = 1 Order by gns.codSucursal "
                    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                    If Not rs Is Nothing Then
                        Do While Not rs.EOF
                            s = s & "|>" & rs!CodBodega
                            rs.MoveNext
                        Loop
                    End If
                    Set rs = Nothing
        Else
                  sql = "Select ivb.codbodega from ivbodega ivb Inner join gnsucursal gns on gns.idsucursal = ivb.idsucursal where ivb.bandvalida = 1  "
                  sql = sql & " And gns.codsucursal = '" & fcbSucursal.KeyText & "'"
                  sql = sql & " Order by gns.codSucursal"
                    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                    If Not rs Is Nothing Then
                        Do While Not rs.EOF
                            s = s & "|>" & rs!CodBodega
                            rs.MoveNext
                        Loop
                    End If
                    Set rs = Nothing
        End If
        s = s & "|>C.Unitario|>C.Total" ' |>Exist|>Exist"
        .FormatString = s
         If chkTodo.value = vbChecked Then
                    sql = "Select ivb.codbodega from ivbodega ivb Inner join gnsucursal gns on gns.idsucursal = ivb.idsucursal where ivb.bandvalida = 1 Order by gns.codSucursal "
                    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                    If Not rs Is Nothing Then
                        Do While Not rs.EOF
                            .ColWidth(17 + i) = 600
                            grd.subtotal flexSTSum, -1, 19 + i, , grd.GridColor, vbBlack, , "Subtotal", 1, True
                            rs.MoveNext
                            i = i + 1
                        Loop
                    End If
                    Set rs = Nothing
            Else
            sql = "Select ivb.codbodega from ivbodega ivb Inner join gnsucursal gns on gns.idsucursal = ivb.idsucursal where ivb.bandvalida = 1  "
            sql = sql & " And gns.codsucursal = '" & fcbSucursal.KeyText & "'"
            sql = sql & "Order by gns.codSucursal"
                    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                    If Not rs Is Nothing Then
                        Do While Not rs.EOF
                            .ColWidth(17 + i) = 600
                            grd.subtotal flexSTSum, -1, 19 + i, , grd.GridColor, vbBlack, , "Subtotal", 1, True
                            rs.MoveNext
                            i = i + 1
                        Loop
                    End If
                    Set rs = Nothing
        End If
        .ColWidth(COL_NUM) = 700
        .ColWidth(1) = 1500
        .ColWidth(2) = 1500
        .ColWidth(3) = 2400 'coditem
        .ColWidth(4) = 4100 'desc
        .ColWidth(5) = 1100 'sucursal
        .ColWidth(6) = 1000
        .ColWidth(7) = 1500
        .ColWidth(8) = 1500
        .ColWidth(17) = 1000
        .ColWidth(18) = 1000
       .ColHidden(COL_DCI_IDCLI) = True
        '.ColHidden(11) = True
        AsignarTituloAColKey grd
       ' If Me.tag = "VxTransProm" Then
         grd.Editable = flexEDNone
            For i = 0 To .Cols - 3
                .ColData(i) = -1
            Next i
            'Color de fondo
            If .Rows > .FixedRows Then
                .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = .BackColorFrozen '.ColIndex("Límite de Crédito")) = .BackColorFrozen
            End If
        'End If

        .ColDataType(COL_NUM) = flexDTLong
        .ColDataType(COL_DCI_RUC + 1) = flexDTString
        .ColDataType(COL_DCI_NOMCLI + 1) = flexDTString
        .ColDataType(COL_DCI_CODIV + 1) = flexDTString
        .ColDataType(COL_DCI_DESCIV + 1) = flexDTString
        .ColDataType(COL_DCI_DESC + 1) = flexDTCurrency
        .ColDataType(COL_RES + 1) = flexDTString
        .ColFormat(COL_DCI_DESC + 1) = fmt
        
        .ColFormat(6) = fmt
        .ColFormat(7) = "#,0.000000"
        .ColFormat(8) = fmt
        .ColFormat(15) = "#,0.000000"
        .ColFormat(16) = fmt
        .ColFormat(17) = "#,0.000000"
        .ColFormat(18) = "#"
        grd.subtotal flexSTClear
        grd.subtotal flexSTSum, -1, 6, , grd.GridColor, vbBlack, , "Subtotal", 1, True
        grd.subtotal flexSTSum, -1, 7, , grd.GridColor, vbBlack, , "Subtotal", 1, True
        grd.subtotal flexSTSum, -1, 8, , grd.GridColor, vbBlack, , "Subtotal", 1, True
        grd.subtotal flexSTSum, -1, 15, , grd.GridColor, vbBlack, , "Subtotal", 1, True
        grd.subtotal flexSTSum, -1, 16, , grd.GridColor, vbBlack, , "Subtotal", 1, True
        grd.subtotal flexSTSum, -1, 17, , grd.GridColor, vbBlack, , "Subtotal", 1, True
        grd.subtotal flexSTSum, -1, 18, , grd.GridColor, vbBlack, , "Subtotal", 1, True
        i = 0
         If chkTodo.value = vbChecked Then
                    sql = "Select ivb.codbodega from ivbodega ivb Inner join gnsucursal gns on gns.idsucursal = ivb.idsucursal where ivb.bandvalida = 1 Order by gns.codSucursal "
                    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                    If Not rs Is Nothing Then
                        Do While Not rs.EOF
                            grd.subtotal flexSTSum, -1, 19 + i, , grd.GridColor, vbBlack, , "Subtotal", 1, True
                            .ColFormat(19 + i) = "#"
                            rs.MoveNext
                            i = i + 1
                        Loop
                    End If
                    Set rs = Nothing
            Else
            sql = "Select ivb.codbodega from ivbodega ivb Inner join gnsucursal gns on gns.idsucursal = ivb.idsucursal where ivb.bandvalida = 1  "
            sql = sql & " And gns.codsucursal = '" & fcbSucursal.KeyText & "'"
            sql = sql & "Order by gns.codSucursal"
                    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                    If Not rs Is Nothing Then
                        Do While Not rs.EOF
                            grd.subtotal flexSTSum, -1, 19 + i, , grd.GridColor, vbBlack, , "Subtotal", 1, True
                            rs.MoveNext
                            i = i + 1
                        Loop
                    End If
                    Set rs = Nothing
        End If
        'grd.subtotal flexSTSum, -1, 16, , grd.GridColor, vbBlack, , "Subtotal", 1, True
       ' grd.subtotal flexSTSum, -1, 17, , grd.GridColor, vbBlack, , "Subtotal", 1, True
       ' grd.subtotal flexSTSum, -1, 18, , grd.GridColor, vbBlack, , "Subtotal", 1, True
        
        If grd.Rows < 2 Then Exit Sub
        Total = grd.ValueMatrix(grd.Rows - 1, 8)
'        ConfigColsParetoItem 7, Total

        If Total = 0 Then
            Total = grd.ValueMatrix(grd.Rows - 1, 7)
            If Total <> 0 Then
                For i = 1 To grd.Rows - 1
                        grd.TextMatrix(i, 9) = Round(grd.ValueMatrix(i, 7) * 100 / Total, 4)
                    
                Next i
            End If
        
        Else
            For i = 1 To grd.Rows - 1
                    grd.TextMatrix(i, 9) = Round(grd.ValueMatrix(i, 8) * 100 / Total, 4)
            Next i
        End If
        
        grd.TextMatrix(1, 10) = grd.ValueMatrix(1, 9)
        For i = 2 To grd.Rows - 1
            grd.TextMatrix(i, 10) = grd.ValueMatrix(i - 1, 10) + grd.ValueMatrix(i, 9)
        Next i
        For i = 1 To grd.Rows - 1
            If Not grd.IsSubtotal(i) Then
                If grd.ValueMatrix(i, 9) = 0 Then
                    grd.TextMatrix(i, 13) = "SR"
                Else
                    Select Case grd.ValueMatrix(i, 10)
                    Case Is < 50
                        grd.TextMatrix(i, 13) = "A"
                    Case Is < 80
                        grd.TextMatrix(i, 13) = "B"
                    Case Is < 100
                        grd.TextMatrix(i, 13) = "C"
                    Case Else
                        If grd.TextMatrix(i, 6) > 0 Then
                            grd.TextMatrix(i, 13) = "C"
                        Else
                            grd.TextMatrix(i, 13) = "SR"
                        End If
                    End Select
                End If
            End If
        Next
        If chkMostrarCostos.value = vbChecked Then
            For i = 1 To grd.Rows - 1
                If Not grd.IsSubtotal(i) Then
                    If grd.ValueMatrix(i, COL_F_IDIV) <> 0 Then
                    Set item = gobjMain.EmpresaActual.RecuperaIVInventarioQuick(grd.ValueMatrix(i, COL_F_IDIV))
                    c = item.CostoDouble(Date, grd.ValueMatrix(i, 16), item.idinventario)
                    grd.TextMatrix(i, grd.Cols - 2) = c
                    grd.TextMatrix(i, grd.Cols - 1) = c * grd.ValueMatrix(i, 16)
                    End If
                End If
            Next i
        End If
        .ColFormat(9) = fmt
        .ColFormat(10) = fmt
        .ColFormat(11) = fmt
        .ColFormat(grd.Cols - 1) = fmt
        .ColFormat(grd.Cols - 2) = fmt

    grd.subtotal flexSTSum, -1, grd.Cols - 1, , grd.GridColor, vbBlack, , "Subtotal", 1, True
    grd.subtotal flexSTSum, -1, grd.Cols - 2, , grd.GridColor, vbBlack, , "Subtotal", 1, True
    End With
End Sub

Private Sub ConfigColsParetoItem(ByVal C_VNETO As Integer, Total As Currency)
    Dim C_DESCUENTO  As Integer, C_SUBTOT As Integer, C_IVA As Integer, C_TOTAL
    Dim i As Long, valor1 As Currency, valor2 As Currency, tsuma As Integer
    If grd.Rows <> grd.FixedRows Then
        With grd
            valor1 = Total * 0.8
            For i = .FixedRows To .Rows - 1
                If .ValueMatrix(i, C_VNETO) <> 0 Then .TextMatrix(i, C_VNETO + 2) = Round(.ValueMatrix(i, C_VNETO) / Total, 4)
                '.TextMatrix(i, C_VNETO + 4) = DateDiff("m", mObjCond.Fecha1, mObjCond.Fecha2)
                If Not .IsSubtotal(i) Then
                    .TextMatrix(i, C_VNETO + 3) = i
                    If i = 1 Then
                        .TextMatrix(i, C_VNETO + 4) = .TextMatrix(i, C_VNETO)
                        .TextMatrix(i, C_VNETO + 5) = .TextMatrix(i, C_VNETO + 2)
                    Else
                        .TextMatrix(i, C_VNETO + 4) = .ValueMatrix(i - 1, C_VNETO + 4) + .ValueMatrix(i, C_VNETO)
                        .TextMatrix(i, C_VNETO + 5) = .ValueMatrix(i - 1, C_VNETO + 5) + .ValueMatrix(i, C_VNETO + 2)
                    End If
                    If i <> 1 Then
                        If .ValueMatrix(i, C_VNETO) <> 0 Then
                            If .ValueMatrix(i - 1, C_VNETO + 4) < valor1 Then
                                .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = vbYellow
                                .TextMatrix(i, C_VNETO + 6) = "80 %"
                            Else
                                .TextMatrix(i, C_VNETO + 6) = "20 %"
                           End If
                        Else
                            .TextMatrix(i, C_VNETO + 6) = "0 %"
                            .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = &H8080FF
                        End If
                    Else
                        If .ValueMatrix(i - 1, C_VNETO + 5) < valor1 Then
                            .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = vbYellow
                            .TextMatrix(i, C_VNETO + 6) = "80 %"
                        Else
                            .TextMatrix(i, C_VNETO + 6) = "20 %"

                        End If
                   End If
                End If
            Next i
            valor1 = Total * 0.8
           .MergeCol(1) = True
            .MergeCol(2) = True
       End With
    End If
End Sub

Private Sub GrabaClasificaxItem()
    Dim sql As String, cod As String, i As Long
    Dim NumReg As Long, totalventa As Currency
    On Error GoTo ErrTrap
    MensajeStatus "Guardando....", 1
    With grd
        If .Rows = .FixedRows Then Exit Sub
        .ShowCell 1, 1
        For i = .FixedRows To .Rows - 1
            If Not .IsSubtotal(i) Then
                .Row = i
                .ShowCell i, 1           'Hace visible la fila actual
'                cod = .TextMatrix(i, COL_COD)
                If Len(grd.TextMatrix(i, 13)) > 0 Then
                        If grd.TextMatrix(i, 5) = "TODAS" Then
                            sql = " UPDATE ivinventario set idgrupo4=(select idgrupo4 from ivgrupo4 where codgrupo4='" & grd.TextMatrix(i, 13) & "'),"
                            sql = sql & " TiempoPromVta = (Select idealdias from ivgrupo4 where codgrupo4 ='" & grd.TextMatrix(i, 13) & "' )"
                            sql = sql & " WHERE idinventario= " & grd.ValueMatrix(i, COL_F_IDIV)
                        Else
                            sql = " UPDATE ivexist set idgrupo=(select idgrupo4 from ivgrupo4 where codgrupo4='" & grd.TextMatrix(i, 13) & "')"
                            sql = sql & " from ivexist ive inner join ivbodega ivb "
                             sql = sql & " inner join gnsucursal gns on ivb.idsucursal = gns.idsucursal "
                             sql = sql & " on ive.idbodega=ivb.idbodega"
                            sql = sql & " WHERE idinventario= " & grd.ValueMatrix(i, COL_F_IDIV)
                            sql = sql & " and codsucursal= '" & grd.TextMatrix(i, 5) & "'"
                        End If
                        gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                        If NumReg > 0 Then
                            .TextMatrix(i, 14) = "Actualizado..."
                        End If
                End If
                .Redraw = True
                .Refresh
            End If
        Next i
    End With
    MensajeStatus "", 0
    Exit Sub
    
ErrTrap:
    MsgBox Err.Description, vbExclamation + vbOKOnly
    Exit Sub
End Sub

Private Sub ExpExcel()
Dim ex As Excel.Application, ws As Worksheet, wkb As Workbook
Dim fila As Long, NumCol As Long, col As Integer
Dim colu() As Long, v() As Long, mayor As Long
    MensajeStatus MSG_PREPARA, vbHourglass
    fila = 4
    NumCol = 0
    Dim i   As Integer
    Dim j   As Integer
    Set ex = New Excel.Application  'Crea un instancia nueva de excel
    Set wkb = ex.Workbooks.Add  'Insertar un libro nuevo
    Set ws = ex.Worksheets.Add  'Inserta una nueva hoja
    With ws
        .Name = Left(Me.Caption, 25)
        .Range("A1").Font.Name = "Times Roman"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Cells(1) = gobjMain.EmpresaActual.GNOpcion.NombreEmpresa
    End With
       
       ws.Cells(2, 1) = Me.Caption
       
        For i = 1 To grd.Cols - 1
          If grd.ColHidden(i) = False Then
               ReDim Preserve colu(NumCol) 'para saber la posicion de la columan en la grilla
                    colu(NumCol) = i 'guarda la posicion
                    NumCol = NumCol + 1 'Para saber el número de columnas se exportan a Excel
                    ws.Cells(fila, NumCol) = grd.TextMatrix(0, i)
                ReDim Preserve v(NumCol)
                v(NumCol - 1) = 0 'encera el vector
          End If
        Next i
       
       ws.Range(ws.Cells(fila, 1), ws.Cells(fila, NumCol)).Font.Bold = True
       ws.Range(ws.Cells(fila, 1), ws.Cells(fila, NumCol)).Borders.LineStyle = 1
       
        For i = grd.FixedRows To grd.Rows - 1
            If Not grd.RowHidden(i) Then
                fila = fila + 1
                j = 1
                mayor = 0
                 For col = 1 To grd.Cols - 1
    
                     If Not grd.ColHidden(col) Then
                        If InStr(1, grd.TextMatrix(i, col), "/") > 0 Then
                             ws.Cells(fila, j) = "'" & grd.TextMatrix(i, col)
                        Else
                            ws.Cells(fila, j) = grd.TextMatrix(i, col)
                        End If
                        j = j + 1
                     End If

               
                 Next col
            End If
            ws.Range(ws.Cells(fila, 1), ws.Cells(fila, NumCol)).Borders.LineStyle = 1
       Next i
       
    ex.Visible = True
    ws.Activate
    Set ws = Nothing
    Set wkb = Nothing
    Set ex = Nothing
    'ex.Quit
    MensajeStatus
End Sub

Public Sub InicioPcGrupoxMontoVentasCobros(ByVal Cadena As String)
    Me.tag = Cadena
    Me.Caption = "Actualizar PcGrupo x Monto de Ventas y Cobros"
    CargaDatosMontoVentasCobros
    ConfigCols
    Me.Show
    Me.ZOrder
End Sub

Private Sub CargaDatosMontoVentasCobros()
    Dim sql As String, cond As String, rs As Recordset, antes As Long, i As Integer, j As Long
    Dim objcond As Condicion
    Dim NumReg As Long
    Static Recargo As String
    On Error GoTo ErrTrap
    antes = grd.Row
    grd.Rows = 1
    Set objcond = gobjMain.objCondicion
    If Not (frmB_VxTrans.InicioMontoVentasCobros(objcond, "PCGMontoVentas")) Then
        grd.SetFocus
        Exit Sub
    End If
       
    grd.Redraw = False
    MensajeStatus MSG_PREPARA, vbHourglass
    
    With objcond
        NumPCGrupo = RecuperaSelecPCGrupo + 1
    
        VerificaExistenciaTabla 0
        cond = " AND gc.FechaTrans between " & FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB)
        cond = cond & " AND  " & FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB)
        
        If Len(.CodTrans) > 0 Then
           cond = cond & " AND GC.CodTrans IN (" & PreparaCadena(.CodTrans) & ")"
        End If
        
        VerificaExistenciaTablaTemp "tmpVentas"
        sql = "SELECT  PCProvCli.idProvcli, PCProvCli.CodProvCli, PCProvCli.Nombre, "
        sql = sql & " round(ABS(SUM(PrecioRealTotal)),2) as MontoVentas, codgrupo" & NumPCGrupo & " as CodGrupo,'' as newgrupo, "
        sql = sql & " SPACE(0) AS Resultado Into  tmpVentas"
        sql = sql & " FROM  "
        sql = sql & " vwConsSUMIVKardexIVA inner join "
        sql = sql & " GNComprobante GC  "
        sql = sql & " inner JOIN PCProvCli "
        sql = sql & " left join pcgrupo" & NumPCGrupo
        sql = sql & " on PCProvCli.idgrupo" & NumPCGrupo & " = pcgrupo" & NumPCGrupo & ".idgrupo" & NumPCGrupo
        sql = sql & " ON GC.IdClienteRef = PCProvCli.IdProvCli "
        sql = sql & " ON vwConsSUMIVKardexIVA.TransID = GC.TransID "
        sql = sql & " WHERE (gc.Estado<>3) AND PRECIORealTOTAL<>0" & cond
'        sql = sql & " And pcprovcli.idprovcli = 121842"
        sql = sql & " GROUP BY PCProvCli.idProvcli,PCProvCli.CodProvCli, PCProvCli.Nombre, PCProvCli.TotalDebe, "
        sql = sql & " pcgrupo" & NumPCGrupo & ".codgrupo" & NumPCGrupo
        'Para filtrar solo los clientes cuyo monto de venta se ha modificado,
        'es decir solo los que han comprado
        sql = sql & " having round(ABS(SUM(PrecioRealTotal)),4) <> PCProvCli.TotalDebe "
        gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
        'sql = sql & "ORDER BY PCProvCli.Nombre"
        
        'aqui saco las fechas q se debe pagar
        VerificaExistenciaTablaTemp "tmpfc"
        sql = "SELECT  gc.codtrans,gc.numtrans,pck.idforma,pck.id,PCK.idprovcli,pck.debe,pck.fechaemision,pck.fechavenci into tmpfc FROM "
        sql = sql & " Pckardex pck inner join gncomprobante gc on gc.transid = pck.transid"
        sql = sql & " WHERE (gc.Estado<>3)" & cond
 '       sql = sql & " And pck.idprovcli = 121842"
        gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
        
        'aqui saco los cobros
        VerificaExistenciaTablaTemp "tmpIT"
        sql = "Select max(pck.fechaemision) as fechacobro,pck.idasignado,t.fechavenci,pck.idprovcli "
         sql = sql & " into tmpIt "
         sql = sql & " From pckardex pck "
         sql = sql & " Inner join gncomprobante gc on gc.transid = pck.transid Inner join tmpfc t"
        sql = sql & " on t.id = pck.idasignado"
        sql = sql & " WHERE (gc.Estado<>3)"
        sql = sql & " AND GC.CodTrans IN (" & PreparaCadena(.codforma) & ")"
  '      sql = sql & " And pck.idprovcli = 121842"
        sql = sql & " group by pck.idasignado,t.fechavenci,pck.idprovcli"
        gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
        
        sql = "Select t.codprovcli,t.nombre,t.montoventas,max(ti.fechacobro) as fcobro,"
        sql = sql & " case when  avg(datediff(d,ti.fechavenci,ti.fechacobro))<0 then 0 else avg(datediff(d,ti.fechavenci,ti.fechacobro)) end as PromDiasCobro,t.codGrupo,'' as NuevoGrupo,'-1' as ConfirmaCambio from tmpventas t"
        sql = sql & " Inner join tmpit ti on ti.idprovcli = t.idProvCli"
        sql = sql & " Group by t.codprovcli,t.nombre,t.montoventas,t.codGrupo"
    End With
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    grd.LoadArray MiGetRows(rs)
    Set rs = Nothing
    With grd
        '# de fila
        .ColAlignment(0) = flexAlignCenterCenter
        GNPoneNumFila grd, False
        'Reubica a la fila donde estaba antes
        If .Rows > antes And antes > 0 Then .Row = antes
        .Redraw = True
        ConfigCols
    End With
    MensajeStatus "", 0
    grd.SetFocus
    If grd.Rows <> grd.FixedRows Then grd.Row = grd.FixedRows
    If CargarPCGrupos(RecuperaSelecPCGrupo) Then
    End If
    For i = 1 To grd.Rows - 1
        For j = 1 To 10
            If Len(gMontoCobro(j).grupo) > 0 Then
                If grd.ValueMatrix(i, COL_TOT) > gMontoCobro(j).desde And grd.ValueMatrix(i, COL_TOT) <= gMontoCobro(j).hasta And grd.ValueMatrix(i, 5) <= gMontoCobro(j).diasMorosidad Then
                    grd.TextMatrix(i, 7) = gMontoCobro(j).grupo
                    Exit For
                Else
                    grd.TextMatrix(i, 7) = "NA"
                End If
            End If
        Next j
    Next i
    Exit Sub
ErrTrap:
    grd.Redraw = True
    MensajeStatus "", 0
    DispErr
    Exit Sub
End Sub

Private Sub GrabarPCGrupoxMontoVentaCobro()
    Dim sql As String, cod As String, i As Long
    Dim NumReg As Long, totalventa As Currency
    On Error GoTo ErrTrap
    MensajeStatus "Guardando....", 1
    With grd
        If .Rows = .FixedRows Then Exit Sub
        .ShowCell 1, 1
        For i = .FixedRows To .Rows - 1
            If Not .IsSubtotal(i) Then
                .Row = i
                .ShowCell i, 1           'Hace visible la fila actual
                cod = .TextMatrix(i, COL_COD)
                If grd.ValueMatrix(i, grd.Cols - 2) = -1 Then
                    If grd.TextMatrix(i, 6) <> grd.TextMatrix(i, 7) Then
                        sql = " UPDATE PCProvCli "
                        sql = sql & " SET idGrupo" & NumPCGrupo & " = (select idgrupo" & NumPCGrupo & " from pcgrupo" & NumPCGrupo & " where codgrupo" & NumPCGrupo & "='" & grd.TextMatrix(i, 7) & "')"
                        sql = sql & " WHERE CodProvCli= '" & cod & "'"
                        gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
                        If NumReg > 0 Then
                            .TextMatrix(i, grd.Cols - 1) = "Actualizado..."
                        Else
                            .TextMatrix(i, grd.Cols - 1) = "Error al tratar de Actualizar..."
                        End If
                    Else
                            .TextMatrix(i, grd.Cols - 1) = "No Hay Cambio.........."
                    End If
                End If
                .Redraw = True
                .Refresh
            End If
        Next i
    End With
    MensajeStatus "", 0
    Exit Sub
ErrTrap:
    MsgBox Err.Description, vbExclamation + vbOKOnly
    Exit Sub
End Sub


