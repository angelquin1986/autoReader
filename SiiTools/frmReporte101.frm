VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReporte101 
   Caption         =   "Formulario 101"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9525
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6195
   ScaleWidth      =   9525
   WindowState     =   2  'Maximized
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   2295
      Left            =   60
      TabIndex        =   0
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
   Begin MSComctlLib.Toolbar tlb1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   1164
      ButtonWidth     =   1402
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imlLista"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "General"
                  Text            =   "General"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "Resumen"
                  Text            =   "Resumen"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar"
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar - F5"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlLista 
      Left            =   0
      Top             =   0
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReporte101.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReporte101.frx":0114
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReporte101.frx":0228
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReporte101.frx":033C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReporte101.frx":0450
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReporte101.frx":0564
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReporte101.frx":0678
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReporte101.frx":0CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReporte101.frx":1148
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmReporte101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const COL_NUM = 0
Const COL_CAB = 1
Const COL_TRA = 2
Const COL_TOT = 3

Public Sub Inicio(ByVal tag As String)
    Dim btnX As Button
    Me.tag = tag
    Me.Show
    Me.ZOrder
    tlb1.Buttons("Imprimir").Style = tbrDropdown
    CargarDatos
End Sub

Private Sub CargarDatos()
    Dim antes As Long
    On Error GoTo ErrTrap
    
    antes = grd.Row
    grd.Rows = 1
    Select Case Me.tag
    Case "F101"
        CargarFormulario101
    
    End Select
        
    With grd
        '# de fila
        .ColAlignment(0) = flexAlignCenterCenter
        GNPoneNumFila grd, False
        'Arregla el ancho de columnas
        If grd.Rows > grd.FixedRows Then
            .AutoSize 0, .Cols - 1
        End If
        'Reubica a la fila donde estaba antes
        If .Rows > antes And antes > 0 Then .Row = antes
        .Redraw = True
    End With
    
    MensajeStatus "", 0
    Exit Sub
ErrTrap:
    grd.Redraw = True
    MensajeStatus "", 0
    DispErr
    Exit Sub
End Sub

''Private Sub CargarResumenVentas()
''    Dim sql As String, rs As Recordset, Cond As String, Bandfirst As Boolean
''    Static lst_cabina As String, lst_trafico As String
''    Static lst_destino As String, fecha1 As Date, fecha2 As Date
''
''    If fecha1 = 0 Then fecha1 = DateSerial(Year(Date), Month(Date), 1)
''    If fecha2 = 0 Then fecha2 = DateSerial(Year(Date), Month(Date) + 1, 1 - 1)
''
''    If Not frmB_V.Inicio("ResumenVentas", lst_cabina, _
''                         lst_trafico, lst_destino, _
''                         fecha1, fecha2) Then
''        grd.SetFocus
''        Exit Sub
''    End If
''
''
''    sql = "SELECT idcabina, trafico, destino, " & _
''          "SUM(duracion) As TotalMinutos, Sum(neto) as SumaNeto, " & _
''          "SUM(ice) As SumaICE, SUM(iva) As SumaIVA, Sum(total) As SumaTotal " & _
''          "FROM kardexlocutorio "
''    Bandfirst = False
''    Cond = "WHERE "
''
''    If Len(lst_cabina) > 0 Then
''        Cond = Cond & "(idcabina IN (" & lst_cabina & "))"
''        Bandfirst = True
''    End If
''
''    If Len(lst_trafico) > 0 Then
''        If Bandfirst Then Cond = Cond & " AND "
''        Cond = Cond & "(trafico IN (" & lst_trafico & "))"
''        Bandfirst = True
''    End If
''
''    If Len(lst_destino) > 0 Then
''        If Bandfirst Then Cond = Cond & " AND "
''        Cond = Cond & "(destino IN (" & lst_destino & "))"
''        Bandfirst = True
''    End If
''
''    If fecha1 <> 0 Or fecha2 <> 0 Then
''        If Bandfirst Then Cond = Cond & " AND "
''        Cond = Cond & "(fecha BETWEEN " & FechaYMD(fecha1, gobjMain.EmpresaActual.TipoDB, False) & _
''                      " and " & FechaYMD(fecha2, gobjMain.EmpresaActual.TipoDB, False) & ")"
''        Bandfirst = True
''    End If
''
''    If Bandfirst Then sql = sql & Cond
''
''    sql = sql & " GROUP BY idcabina, trafico, destino" & _
''                " ORDER BY idcabina, trafico, destino"
''    grd.Redraw = flexRDNone
''    MiGetRowsRep AbrirTablaLectura(sql, gcnLocutorio), grd
''    ConfigColsResumenVentas
''End Sub

Private Sub ConfigColsResumenVentas()
'    Dim i As Integer, j As Integer, fmt As String
'    Dim color1 As Long, color2 As Long, color3 As Long
'    With grd
'        fmt = "#0,0.0000"
'        .FormatString = ">#|<Cabina|<Tráfico|<Destino|>Total Minutos" & _
'                        "|>Neto|>ICE|>IVA|>TOTAL"
'
'        .ColDataType(COL_CAB) = flexDTString
'        .ColDataType(COL_TRA) = flexDTString
'
'        For i = COL_MIN To COL_TOT
'            .ColDataType(i) = flexDTCurrency
'            .ColFormat(i) = fmt
'        Next i
'
'        .MergeCells = flexMergeFree
'        .MergeCol(COL_CAB) = True
'        .MergeCol(COL_TRA) = True
'
'        color1 = &HE1B9C8
'        color2 = &H784B6E
'        color3 = &H1A5A
'
'        .SubtotalPosition = flexSTBelow
'        For i = COL_CAB To COL_TRA
'            For j = COL_MIN To COL_TOT
'                Select Case i
'                Case 1
'                    .Subtotal flexSTSum, i, j, fmt, color2, vbWhite
'                Case 2
'                    .Subtotal flexSTSum, i, j, fmt, color1, color3
'                End Select
'            Next j
'        Next i
'
'        For j = COL_MIN To COL_TOT
'            .Subtotal flexSTSum, -1, j, fmt, vbBlue, vbYellow, , "TOTAL", , True
'        Next j
'    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF5
        CargarDatos
    Case vbKeyEscape
        Unload Me
    End Select
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    grd.Move 0, tlb1.Height, Me.ScaleWidth, Me.ScaleHeight - tlb1.Height
End Sub

Private Sub tlb1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Imprimir"
        Imprimir 1
    Case "Buscar"
        CargarDatos
    End Select
End Sub

Private Sub Imprimir(ByVal Indice As Integer)
    Dim NumLinDesde As Long, NumLinHasta As Long, BandTodo As Boolean
    On Error GoTo ErrTrap
'que indice
'  1 s_imprimir -> Normal
'  2 s_indice   -> Con indices
'  3 s_resumen  -> Resumen

    If grd.Rows = grd.FixedRows Then Exit Sub 'Si no hay nada, no hace nada
    'If Me.Tag <> "ConsIVVentaVol" Then
  'MensajeStatus MSG_PREPARA, vbHourglass
    Select Case Indice
    '---Proveedor cliente -------------
    Case 1 '***Agregado. 26/06/2003. Angel
        MensajeStatus "", 0
        'GeneralImprimeModGrafF101 grd, "Formulario 101", frmB_V.dtpFecha1.value, frmB_V.dtpFecha2.value
        FrmImprimeEtiketas.InicioF101 grd, 0
        
    Case 2 '***Agregado. 26/06/2003. Angel
        MensajeStatus "", 0
        'ResumenImprimeModGraf grd, "Ventas x Mes Resumen", frmB_V.dtpFecha1.value, frmB_V.dtpFecha2.value
        FrmImprimeEtiketas.InicioF101 grd, 1
   End Select
    grd.SetFocus
    MensajeStatus "", 0
    Exit Sub
ErrTrap:
    DispErr
    grd.SetFocus
    MensajeStatus "", 0
End Sub

Private Sub tlb1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu
        Case "General"
                Imprimir 1
        Case "Resumen"
                Imprimir 2
    End Select
End Sub

Private Sub CargarFormulario101()
    Dim sql As String, rs As Recordset, cond As String, Bandfirst As Boolean
    Static lst_cabina As String, lst_trafico As String
    Static lst_destino As String, fecha1 As Date, fecha2 As Date
    
    If fecha1 = 0 Then fecha1 = DateSerial(Year(Date), Month(Date), 1)
    If fecha2 = 0 Then fecha2 = DateSerial(Year(Date), Month(Date) + 1, 1 - 1)
    
    If Not frmB_V.InicioF101("Formulario101", fecha1, fecha2) Then
        grd.SetFocus
        Exit Sub
    End If
    
    sql = "select"
    sql = sql & " case tipocuenta WHEN 1 THEN 'ACTIVO' WHEN 2 THEN 'PASIVO'  WHEN 3 THEN 'PATRIMONIO'  WHEN 4 THEN 'INGRESOS' WHEN 5 THEN 'COSTOS' ELSE 'GASTOS' END AS Tipo  ,"
    sql = sql & " CampoF101, ct.codcuenta, ct.nombrecuenta, sum(debe)-sum(Haber) as saldo"
    sql = sql & " from gncomprobante g "
    sql = sql & " inner join ctlibrodetalle ctl "
    sql = sql & " inner join ctcuenta ct on ctl.idcuenta=ct.idcuenta"
    sql = sql & " on g.codasiento=ctl.codasiento"
    sql = sql & " Where g.Estado <> 3 And g.Estado <> 0 And bandtotal = 0 And ct.bandvalida = 1"
    sql = sql & "and  (fechatrans BETWEEN " & FechaYMD(fecha1, gobjMain.EmpresaActual.TipoDB, False) & _
                    " and " & FechaYMD(fecha2, gobjMain.EmpresaActual.TipoDB, False) & ")"
    
    sql = sql & " group by tipocuenta, ct.codcuenta, ct.nombrecuenta, CampoF101"
    sql = sql & " order by tipocuenta, CampoF101, ct.codcuenta, ct.nombrecuenta "
    
    
    grd.Redraw = flexRDNone
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    MiGetRowsRep rs, grd
    ConfigColsFormulario101
End Sub

Private Sub ConfigColsFormulario101()
    Dim i As Integer, j As Integer, fmt As String
    Dim color1 As Long, color2 As Long, color3 As Long
    With grd
        fmt = "#0,0.00"
        .FormatString = ">#|<Tipo|<Campo F101|<Cod. Cuenta|<Nombre Cuenta|>Valor" & _
                        ""
                        
        .ColDataType(COL_CAB) = flexDTString
        .ColDataType(COL_TRA) = flexDTString
        .ColDataType(COL_TOT + 2) = flexDTCurrency
        .ColFormat(COL_TOT + 2) = fmt
        
        
        .MergeCells = flexMergeFree
        .MergeCol(COL_CAB) = True
        
        
        color1 = &HE1B9C8
        color2 = &H784B6E
        color3 = &H1A5A
        
'        .SubtotalPosition = flexSTBelow
'        For i = COL_CAB To COL_TRA
'            For j = COL_MIN To COL_TOT
'                Select Case i
'                Case 1
'                    .Subtotal flexSTSum, i, j, fmt, color2, vbWhite
'                Case 2
                    .subtotal flexSTSum, 2, 5, fmt, color1, color3
'                End Select
'            Next j
'        Next i
        
'        For j = COL_MIN To COL_TOT
            .subtotal flexSTSum, 1, 5, fmt, vbBlue, vbYellow, , "TOTAL", , True
'        Next j
    End With
End Sub


