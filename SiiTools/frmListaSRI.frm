VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmListaSRI 
   BackColor       =   &H8000000A&
   Caption         =   "Costos"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9720
   FillColor       =   &H00C00000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   9720
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   5280
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   3135
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   6135
      _cx             =   10821
      _cy             =   5530
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
      SubtotalPosition=   1
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
   Begin MSComctlLib.ImageList img1 
      Left            =   9120
      Top             =   600
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSRI.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSRI.frx":0114
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSRI.frx":0568
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSRI.frx":067C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSRI.frx":0790
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSRI.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSRI.frx":0E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSRI.frx":1B20
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSRI.frx":1F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSRI.frx":2084
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaSRI.frx":3906
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   1005
      ButtonWidth     =   3175
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "img1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Buscar"
            Key             =   "Configurar"
            Object.ToolTipText     =   "Configurar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Abrir"
            Key             =   "Abrir"
            Object.ToolTipText     =   "Abir Archivo"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Asignar"
            Key             =   "Asignar"
            Description     =   "Asignar un valor"
            Object.ToolTipText     =   "Asignar un valor "
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Comparar (F5)"
            Key             =   "Comparar"
            Object.ToolTipText     =   "Calcular "
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exp. Excel"
            Key             =   "Excel"
            Object.ToolTipText     =   "A Excel"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Corregir Autorizacion"
            Key             =   "Corregir"
            Object.ToolTipText     =   "Corregir Autorizacion"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Solo Error"
            Key             =   "SoloError"
            Object.ToolTipText     =   "Cerrar (ESC)"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   9720
      TabIndex        =   0
      Top             =   4185
      Width           =   9720
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdCos 
      Height          =   1215
      Left            =   1080
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   3615
      _cx             =   6376
      _cy             =   2143
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
      FixedCols       =   0
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
      SubtotalPosition=   1
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
Attribute VB_Name = "frmListaSRI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ex As Excel.Application, ws As Worksheet, wkb As Workbook
Private mProcesando As Boolean
Private mCancelado As Boolean
Private mcolItemsSelec As Collection      'Coleccion de items

Const COL_COMP = 1
Const COL_DESCCOMP = 2
Const COL_ESTAB = 3
Const COL_PUNTO = 4
Const COL_SECUENCIAL = 5
Const COL_RUCSRI = 6
Const COL_FECHAAUTO = 9
Const COL_EMISION = 10
Const COL_RUC = 11
Const COL_CLAVEACCESO = 12
Const COL_AUTORIZA = 13
Const COL_TOTAL = 14
Const COL_TRANSANEXO = 15
Const COL_NUMTRANSANEXO = 16
Const COL_AUTOSRIANEXO = 17
Const COL_FECHAANEXO = 18
Const COL_RUCANEXO = 19
Const COL_NOMBREANEXO = 20
Const COL_TRANSID = 21
Const COL_RESULTADO = 22

Dim v() As String

Public Sub Inicio(ByVal tag As String)
    Dim rutaPlantilla
    Dim i As Integer
    Dim valor As Currency
    On Error GoTo ErrTrap
    
    Me.tag = tag            'Guarda en la propiedad Tag para distinguir después
    Me.Show
    Me.ZOrder
'    Select Case Me.tag
'    Case "Costos"
'        Me.Caption = "Punto de Equilibrio basado en Producción "
'    Case "Produccion"
'         Me.Caption = "Punto de Equilibrio basado en Ventas "
'    End Select
       
    'Inicializa la grilla
    grd.Rows = grd.FixedRows
   ConfigCols
'    VisualizarTexto (rutaPlantilla)
    For i = 1 To 8 'Agrega filas adicionales
            grd.AddItem vbTab '& "."
        Next
   ConfigCols
'    FijarColor ': CargarConfiguarciones
            
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF3
        KeyCode = 0
    Case vbKeyF5
        'verificar
    Case vbKeyF6
        KeyCode = 0
    Case vbKeyP
        If Shift And vbCtrlMask Then
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
    
    
    If grd.CellBackColor = grd.BackColorFrozen Or grd.CellBackColor = &HC00000 Then
       Cancel = True
    End If
End Sub

Private Sub grd_BeforeSort(ByVal col As Long, Order As Integer)
    'Impide mientras está procesando
    If mProcesando Then Order = flexSortNone
End Sub

Private Sub grd_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
    
        
    End Select
End Sub

Private Sub grd_KeyPressEdit(ByVal Row As Long, ByVal col As Long, KeyAscii As Integer)
Select Case grd.ColDataType(col)
    Case flexDTCurrency, flexDTSingle, flexDTDouble
        If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
           (KeyAscii <> vbKeyBack) And _
           (KeyAscii <> Asc(".")) And _
           (KeyAscii <> Asc("-")) And _
           (KeyAscii <> vbKeyReturn) Then
            KeyAscii = 0
        End If
    Case flexDTLong
        If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
           (KeyAscii <> vbKeyBack) And _
           (KeyAscii <> vbKeyReturn) Then
            KeyAscii = 0
        End If
    End Select
End Sub

Private Sub tlb1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Configurar":
    Case "Abrir"
        AbrirArchivo
        
    Case "Comparar"
        CompararDatos
    Case "Excel": ExportaExcel ("Verificacion Comprobantes Electronicos")
    
    Case "Corregir"
        CorregirAutorizacion
    Case "SoloError":
        SoloError
    End Select
End Sub

Private Sub ConfigCols(Optional subt As Integer)
    Dim s As String, i As Long, j As Integer, s1 As String
    Dim fmt As String
    With grd
               s = "^#|<Tipo Comp|<Descrip Comp|<Estab|<Punto|<Secuencial|<RUC|<Proveedor|>Fecha Emision"
               s = s & "|<FechaAutorizacion|<Tipo Emision|<Ruc Cliente"
               s = s & "|<Clave Acceso|>Autorizacion|>Total|<Trans|<Num Trans|<Atoriza Anexo|<Fecha Anexo|<RUC Anexo|<Nombre Anexo|<Transid|<Resultado"
  
        .FormatString = s
        AjustarAutoSize grd, -1, -1, 4000
        AsignarTituloAColKey grd
        
        .ColHidden(COL_COMP) = True
        .ColHidden(COL_TRANSID) = True
        .ColHidden(COL_EMISION) = True
        .ColHidden(COL_RUC) = True
        .ColHidden(COL_TOTAL) = True
        .ColHidden(COL_CLAVEACCESO) = True
        
        
        
    End With
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


Public Sub ExportaExcel(ByVal titulo As String)
    'tipo=0 Roles de Pagos; tipo=1 Reporte Bancos y Provisiones
    Set ex = New Excel.Application  'Crea un instancia nueva de excel
    Set wkb = ex.Workbooks.Add  'Insertar un libro nuevo
    Set ws = ex.Worksheets.Add  'Inserta una nueva hoja
    With ws
        .Name = Mid$(titulo, 1, 30)
        .Range("A1").Font.Name = "Times Roman"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Cells(1) = gobjMain.EmpresaActual.GNOpcion.NombreEmpresa
    End With
        Exportar titulo
    ex.Visible = True
    ws.Activate
    Set ws = Nothing
    Set wkb = Nothing
    Set ex = Nothing
    'ex.Quit
End Sub

Private Sub Exportar(ByVal titulo As String)
        Dim fila As Long, col As Long, i As Long, j As Long
    Dim v() As Long, mayor As Long
    Dim NumCol As Integer
    Dim fmt As String
    prg1.min = 0
    prg1.max = grd.Rows - 1
    MensajeStatus "Está Exportando  a Excel ...", vbHourglass
    With ws
        fila = 2
        .Range("H1").Font.Name = "Arial"
        .Range("H1").Font.Size = 12
        .Range("H1").Font.Bold = True
        .Cells(fila, 1) = titulo
        
'        .PageSetup.PaperSize = xlPaperA4 ' = xlPaperLetter 'Tamaño del papel (carta)
        .PageSetup.BottomMargin = Application.CentimetersToPoints(1.5) 'Margen Superior
        .PageSetup.TopMargin = Application.CentimetersToPoints(1) 'Margen Inferior
        .Range(.Cells(1, 13), .Cells(500, 13)).NumberFormat = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda(fmt)   'Establece el formato para los números
        .Range("A2:AZ1000").Font.Name = "Arial"    'Tipo de letra para toda la hoja
        .Range("A2:AZ1000").Font.Size = 8          'Tamaño de la letra
        
        
        fila = fila + 2
        NumCol = 0
        For i = 1 To grd.Cols - 1
                NumCol = NumCol + 1
                .Cells(fila, NumCol) = grd.TextMatrix(0, i) 'cabeceras
                ReDim Preserve v(NumCol)
                v(NumCol - 1) = 0
        Next i
                
        .Range(.Cells(fila, 1), .Cells(fila, NumCol)).Font.Bold = True
        .Range(.Cells(fila, 1), .Cells(fila, NumCol)).Borders.LineStyle = 12
        .Range(.Cells(fila, 3), .Cells(fila, grd.Rows - 1)).HorizontalAlignment = xlHAlignLeft
        For i = 2 To grd.Rows - 1
            If grd.RowHidden(i) = False Then
            prg1.value = i
            fila = fila + 1
                j = 1
                mayor = 0
                For col = 1 To grd.Cols - 1
                                            .Cells(fila, j) = "'" & grd.TextMatrix(i, col)
                        mayor = Len(grd.TextMatrix(i, col)) 'Para ajustar el ancho de columnas
                        If mayor > v(j - 1) Then            'de acuerdo a la celda más grande
                            .Columns(j).ColumnWidth = mayor '13/11/2000 ---> Angel P.
                            v(j - 1) = mayor
                        End If
                        j = j + 1
                Next col

            .Range(.Cells(fila, 1), .Cells(fila, NumCol)).Borders.LineStyle = 1
            End If
        Next i
    End With
     prg1.value = prg1.min
     MensajeStatus "Listo", vbDefault
End Sub


Private Sub AbrirArchivo()
    Dim i As Long
    
    On Error GoTo ErrTrap
    With dlg1
        .CancelError = True
'        .Filter = "Texto (Separado por coma)|*.txt|Excel 97(XLS)|*.xls"
        .Filter = "Texto (Separado por tabuladores *.txt)|*.txt|Texto (Separado por coma *.csv)|*.csv|Todos *.*|*.*"
        .flags = cdlOFNFileMustExist
        If Len(.filename) = 0 Then          'Solo por primera vez, ubica a la carpeta de la aplicación
            .filename = App.Path & "\*.txt"
        End If
        
        .ShowOpen
        
        Select Case UCase$(Right$(dlg1.filename, 4))
        Case ".TXT", ".CSV"
            'ReformartearColumnas
            VisualizarTexto dlg1.filename
            'InsertarColumnas
        Case ".XLS"
       '     VisualizarExcel dlg1.FileName
        Case Else
        End Select
    End With
    Exit Sub
ErrTrap:
    If Err.Number <> 32755 Then DispErr
    Exit Sub
End Sub

Private Sub VisualizarTexto(ByVal archi As String)
    Dim f As Integer, s As String, i As Integer, j As Integer
    Dim Cadena
    On Error GoTo ErrTrap
    ReDim rec(0, 1)
    MensajeStatus "Está leyendo el archivo " & archi & " ...", vbHourglass
    grd.Rows = grd.FixedRows    'Limpia la grilla
    grd.Redraw = flexRDNone
    f = FreeFile                'Obtiene número disponible de archivo

Dim Separo() As String
Dim campo As Integer
Dim X As Integer
Dim cad As Variant
Dim comprobante As String


    'Abre el archivo para lectura
    Open archi For Input As #f
        Do Until EOF(f)
            Line Input #f, s
            's = vbTab & Replace(s, "VbTab", vbTab)      'Convierte ',' a TAB
            Separo = Split(s, vbTab)
           'grd.AddItem s
           
           comprobante = ""
           For j = 0 To UBound(Separo) Step 10
                If j = 0 Then
                    Cadena = Separo(0 + j) & vbTab & Separo(1 + j) & vbTab & Separo(2 + j) & vbTab & Separo(3 + j) & vbTab & Separo(4 + j) & vbTab & Separo(5 + j) & vbTab & Separo(7 + j) & vbTab & Separo(8 + j) '& vbTab & Separo(9 + j) '& vbTab & Separo(10 + j)
                    cad = Split(Separo(9 + j), Chr(10))
'                    grd.AddItem j & vbTab & Cadena & vbTab & Mid$(Separo(9 + j), 1, 20) & vbTab & Mid$(Separo(9 + j), 21, 13)
'                    comprobante = cad(2)
                    Select Case cad(2)
                    Case "Factura"
                        comprobante = "1" & vbTab & "Factura"
                    Case "Notas de Crédito"
                        comprobante = "4" & vbTab & "Notas de Crédito"
                    Case "Notas de Débito"
                        comprobante = "5" & vbTab & "Notas de Débito"
                    Case "Comprobante de Retención"
                        comprobante = "7" & vbTab & "Comprobante de Retención"
                    End Select
                
                Else
                    Cadena = Mid$(Separo(0 + j), 1, 3) & vbTab & Mid$(Separo(0 + j), 5, 3) & vbTab & Mid$(Separo(0 + j), 9, 9) & vbTab & Separo(1 + j) & vbTab & Separo(2 + j) & vbTab & Separo(3 + j) & vbTab & Separo(4 + j) & vbTab & Separo(5 + j) & vbTab & Separo(7 + j) & vbTab & Separo(8 + j) '& vbTab & Separo(9 + j) '& vbTab & Separo(10 + j)
                    cad = Split(Separo(9 + j), Chr(10))
                    If UBound(cad) > 0 Then
                        grd.AddItem j / 10 & vbTab & comprobante & vbTab & Cadena & vbTab & cad(0) & vbTab & Format(cad(1), "0.00")
                        Select Case cad(2)
                        Case "Factura"
                            comprobante = "1" & vbTab & "Factura"
                        Case "Notas de Crédito"
                            comprobante = "4" & vbTab & "Notas de Crédito"
                        Case "Notas de Débito"
                            comprobante = "5" & vbTab & "Notas de Débito"
                        Case "Comprobante de Retención"
                            comprobante = "7" & vbTab & "Comprobante de Retención"
                        End Select
                    End If
                End If
            grd.Redraw = flexRDDirect
          Next j
           
        Loop
    Close #f
    RemueveSpace
    grd.ColSort(1) = flexSortGenericAscending
    grd.Sort = flexSortUseColSort
    grd.Redraw = flexRDDirect
    AjustarAutoSize grd, -1, -1
    grd.ColWidth(grd.Cols - 1) = 5000
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


Private Sub CompararDatos()
    Dim i As Integer, sql As String, rs As Recordset
    Dim X As Single
    
    prg1.min = 0
    prg1.max = grd.Rows - 1

    
    For i = 1 To grd.Rows - 1
    
        DoEvents
        prg1.value = i
        grd.Row = i
        X = grd.CellTop                 'Para visualizar la celda actual
        
        
        If Len(grd.TextMatrix(i, COL_COMP)) = 0 Then Exit Sub
        
        If Not grd.IsSubtotal(i) Then
            
            If grd.ValueMatrix(i, COL_COMP) <> 7 Then
                sql = "select ruc, pc.nombre, a.numautsri, a.fechaanexos, CODTRANS,NUMTRANS, G.TRANSID "
                sql = sql & " from gncomprobante g "
                sql = sql & " inner join anexos a on g.transid=a.transid"
                sql = sql & " inner join pcprovcli pc on g.idproveedorref=pc.idprovcli"
                sql = sql & " Where g.Estado <> 3"
                sql = sql & " and pc.ruc='" & grd.TextMatrix(i, COL_RUCSRI) & "'"
                sql = sql & " and numserieestablecimiento='" & grd.TextMatrix(i, COL_ESTAB) & "'"
                sql = sql & " and numseriepunto='" & grd.TextMatrix(i, COL_PUNTO) & "'"
                sql = sql & " and numsecuencial='" & grd.TextMatrix(i, COL_SECUENCIAL) & "'"
                sql = sql & " and codtipocomp=" & grd.TextMatrix(i, COL_COMP)
            Else
                sql = " select ruc, pc.nombre, a.autorizacionsri as numautsri, a.fechacaducidadsri as fechaanexos , CODTRANS,NUMTRANS, G.TRANSID "
                sql = sql & " from gncomprobante g  "
                sql = sql & " inner join tskardexret a on g.transid=a.transid"
                sql = sql & " inner join pcprovcli pc on g.idclienteref=pc.idprovcli"
                sql = sql & " Where g.Estado <> 3"
                sql = sql & " and pc.ruc='" & grd.TextMatrix(i, COL_RUCSRI) & "'"
                sql = sql & " and a.numserieestasri='" & grd.TextMatrix(i, COL_ESTAB) & "'"
                sql = sql & " and a.numseriepuntosri='" & grd.TextMatrix(i, COL_PUNTO) & "'"
                sql = sql & " and a.numsecuencialSRI='" & grd.TextMatrix(i, COL_SECUENCIAL) & "'"
                
            End If
            
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If rs.RecordCount > 0 Then
                grd.TextMatrix(i, COL_TRANSANEXO) = rs.Fields("CODTRANS")
                grd.TextMatrix(i, COL_NUMTRANSANEXO) = rs.Fields("NUMTRANS")
                grd.TextMatrix(i, COL_AUTOSRIANEXO) = rs.Fields("numautsri")
                grd.TextMatrix(i, COL_FECHAANEXO) = rs.Fields("fechaanexos")
                grd.TextMatrix(i, COL_RUCANEXO) = rs.Fields("ruc")
                grd.TextMatrix(i, COL_NOMBREANEXO) = rs.Fields("Nombre")
                grd.TextMatrix(i, COL_TRANSID) = rs.Fields("Transid")
            End If
            
            If Len(grd.TextMatrix(i, COL_AUTOSRIANEXO)) > 0 Then
                If grd.TextMatrix(i, COL_AUTOSRIANEXO) = grd.TextMatrix(i, COL_AUTORIZA) Then
                    grd.TextMatrix(i, COL_RESULTADO) = "OK."
                    grd.Cell(flexcpBackColor, i, 1, i, COL_RESULTADO) = vbWhite
                Else
                    grd.TextMatrix(i, COL_RESULTADO) = "Error Número Autirización"
                    grd.Cell(flexcpForeColor, i, 1, i, COL_RESULTADO) = vbRed
                End If
            Else
                grd.TextMatrix(i, COL_RESULTADO) = "Error Comprobante NO Ingresado"
                grd.Cell(flexcpBackColor, i, 1, i, COL_RESULTADO) = vbRed
            End If
            
            If Len(grd.TextMatrix(i, COL_RUCANEXO)) > 0 Then
                If grd.TextMatrix(i, COL_RUCANEXO) <> grd.TextMatrix(i, COL_RUCSRI) Then
                    grd.TextMatrix(i, COL_RESULTADO) = "Error Número RUC diferentes"
                    grd.Cell(flexcpForeColor, i, 1, i, COL_RESULTADO) = vbBlue
                
                End If
            End If
            
            
        End If
    Next i
    AjustarAutoSize grd, -1, -1
End Sub

Private Sub CorregirAutorizacion()
    Dim i As Integer, sql As String, rs As Recordset
    Dim X As Single
    
    prg1.min = 0
    prg1.max = grd.Rows - 1

    
    For i = 1 To grd.Rows - 1
    
        DoEvents
        prg1.value = i
        grd.Row = i
        X = grd.CellTop                 'Para visualizar la celda actual
    
        If grd.Cell(flexcpForeColor, i, 1, i, COL_RESULTADO) = vbRed Then
            If grd.ValueMatrix(i, COL_COMP) <> 7 Then
                sql = "update anexos set numautsri='" & grd.TextMatrix(i, COL_AUTORIZA) & "'"
                sql = sql & " Where transid=" & grd.ValueMatrix(i, COL_TRANSID)
            Else
                sql = "update tskardexret set autorizacionsri='" & grd.TextMatrix(i, COL_AUTORIZA) & "'"
                sql = sql & " Where transid=" & grd.ValueMatrix(i, COL_TRANSID)
            End If
            
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            If Len(grd.TextMatrix(i, COL_AUTOSRIANEXO)) > 0 Then
                    grd.TextMatrix(i, COL_RESULTADO) = "OK. Corregido "
                    grd.Cell(flexcpBackColor, i, 1, i, COL_RESULTADO) = vbWhite
                    grd.Cell(flexcpForeColor, i, 1, i, COL_RESULTADO) = vbBlack
            End If
            
            
        End If
    Next i
End Sub

Private Sub SoloError()
    Dim i As Integer, sql As String, rs As Recordset
    Dim X As Single
    
    prg1.min = 0
    prg1.max = grd.Rows - 1

    
    For i = 1 To grd.Rows - 1
    
        DoEvents
        prg1.value = i
        grd.Row = i
        X = grd.CellTop                 'Para visualizar la celda actual
    
        If grd.TextMatrix(i, COL_RESULTADO) = "OK." Then
            grd.RowHidden(i) = True
        End If
    Next i
End Sub


