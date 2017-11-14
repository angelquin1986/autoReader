VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFormulario103 
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
            Picture         =   "frmFormulario103.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario103.frx":0114
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario103.frx":0568
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario103.frx":067C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario103.frx":0790
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario103.frx":0BE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario103.frx":0E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario103.frx":1B20
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario103.frx":1F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario103.frx":2084
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormulario103.frx":3906
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   953
      ButtonWidth     =   2090
      ButtonHeight    =   953
      Style           =   1
      ImageList       =   "img1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Abrir"
            Key             =   "Abrir"
            Object.ToolTipText     =   "Abrir"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo"
            Key             =   "Nuevo"
            Description     =   "Nuevo"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar (F5)"
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exp. Excel"
            Key             =   "Excel"
            Object.ToolTipText     =   "A Excel"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Guardar Resul."
            Key             =   "Guardar"
            Object.ToolTipText     =   "Guardar Resultado"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cerrar"
            Key             =   "Cerrar"
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
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   3135
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6135
      _cx             =   10821
      _cy             =   5530
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483624
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
      Rows            =   68
      Cols            =   10
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
End
Attribute VB_Name = "frmFormulario103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ex As Excel.Application, ws As Worksheet, wkb As Workbook
Private mProcesando As Boolean
Private mCancelado As Boolean
Private mcolItemsSelec As Collection      'Coleccion de items
'jeaa 24/09/04 asignacion de grupo a los items
Dim v() As String
Dim costoFijoMensual As Currency, Precio As Integer
Private Const ColorFondo = &H8000000A
Private Const ColorFondo1 = &H80000013
Private Busqueda As Boolean
Private mobjBusq As Busqueda
Private mCodMoneda  As String
Private objcond As Condicion

Public Sub Inicio(ByVal tag As String)


    Dim rutaPlantilla
    Dim i As Integer
    Dim valor As Currency
    On Error GoTo Errtrap
    
    Me.tag = tag            'Guarda en la propiedad Tag para distinguir después
    Me.Show
    Me.ZOrder
    Select Case Me.tag
    Case "F103"
        Me.Caption = "Declaración del Impuesto a la Renta "
    Case "Produccion"
         Me.Caption = "Punto de Equilibrio basado en Ventas "
    End Select
    tlb1.Buttons(3).Enabled = False
    grd.Rows = grd.FixedRows
     LlenaFormatoFormulario
    ConfigCols
    Exit Sub
Errtrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Initialize()
    Set mobjBusq = New Busqueda
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF3
        KeyCode = 0
    Case vbKeyF5
          Buscar
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


Private Sub grd_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row < grd.FixedRows Then Cancel = True
    If grd.IsSubtotal(Row) = True Then Cancel = True
    If grd.ColData(Col) < 0 Then Cancel = True
    CalcularPorcentajes
    
    If grd.CellBackColor = grd.BackColorFrozen Or grd.CellBackColor = &HC00000 Then
       Cancel = True
    End If
End Sub

Private Sub grd_BeforeSort(ByVal Col As Long, Order As Integer)
    'Impide mientras está procesando
    If mProcesando Then Order = flexSortNone
End Sub

Private Sub grd_ChangeEdit()
    'CalcularPorcentajes
End Sub

Private Sub grd_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
Select Case grd.ColDataType(Col)
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
    Case "Abrir":
            AbrirArchivo
    Case "Nuevo":
            nuevo
            tlb1.Buttons(3).Enabled = True
    Case "Buscar":
            Buscar
    Case "Excel": ExportaExcel ("Formulario 104")
    Case "Guardar": GuardarResultado
    Case "Cerrar":      Cerrar
    End Select
End Sub
Private Sub ConfigCols()
    Dim s As String, i As Long, j As Integer, s1 As String
    Dim fmt As String
    With grd
        s = "^#|<c1|<c2|^c3|^c4|>c5|^c6|>c7|^c8|^c9|>c10"
        .FormatString = s
        'AjustarAutoSize grd, -1, -1, 4000
        AsignarTituloAColKey grd
        .ColWidth(0) = 350
        .ColWidth(1) = 1000
        .ColWidth(2) = 5000
        .ColWidth(3) = 450
        .ColWidth(4) = 450
        .ColWidth(5) = 2200
        .ColWidth(6) = 650
        .ColWidth(7) = 1800
        .ColWidth(8) = 650
        .ColWidth(9) = 650
        .ColWidth(10) = 1800
        '.ColWidth(11) = 650
        
        'grilla de resultados
        'Columnas modificables (Longitud maxima)
            .ColFormat(.ColIndex("c5")) = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda(fmt)
            .ColFormat(.ColIndex("c7")) = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda(fmt)
            .ColFormat(.ColIndex("c10")) = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda(fmt)
        .MergeCol(3) = True
        .Refresh
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
        .Name = titulo
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
        Dim Fila As Long, Col As Long, i As Long, j As Long
    Dim v() As Long, mayor As Long
    Dim NumCol As Integer
    Dim fmt As String
    prg1.Min = 0
    prg1.max = grd.Rows - 1
    MensajeStatus "Está Exportando  a Excel ...", vbHourglass
    With ws
        Fila = 2
        .Range("H1").Font.Name = "Arial"
        .Range("H1").Font.Size = 10
        .Range("H1").Font.Bold = True
        .Cells(Fila, 1) = titulo
        
        .PageSetup.PaperSize = xlPaperLetter 'Tamaño del papel (carta)
        .PageSetup.BottomMargin = Application.CentimetersToPoints(1.5) 'Margen Superior
        .PageSetup.TopMargin = Application.CentimetersToPoints(1) 'Margen Inferior
'        .Range(.Cells(1, 13), .Cells(500, 23)).NumberFormat = gobjMain.EmpresaActual.GNOpcion.FormatoMoneda(fmt)    'Establece el formato para los números
        .Range("A2:AZ100").Font.Name = "Arial"    'Tipo de letra para toda la hoja
        .Range("A2:AZ100").Font.Size = 7          'Tamaño de la letra
        
        Fila = Fila + 2
        NumCol = 0
        For i = 1 To grd.Cols - 1
            NumCol = NumCol + 1
            .Cells(Fila, NumCol) = grd.TextMatrix(0, i) 'cabeceras
            ReDim Preserve v(NumCol)
            v(NumCol - 1) = 0
        Next i
                
        .Range(.Cells(Fila, 1), .Cells(Fila, NumCol)).Font.Bold = True
        .Range(.Cells(Fila, 1), .Cells(Fila, NumCol)).Borders.LineStyle = 12
        .Range(.Cells(Fila, 3), .Cells(Fila, grd.Rows - 1)).HorizontalAlignment = xlHAlignLeft
        For i = 2 To grd.Rows - 1
            prg1.value = i
            Fila = Fila + 1
            If grd.IsSubtotal(i) = True Then
                .Range(.Cells(Fila, 1), .Cells(Fila, NumCol)).Font.Bold = True
            Else
                j = 1
                mayor = 0
                For Col = 1 To grd.Cols - 1
                        Select Case Col
                            'Case 1: mayor = 2
                            Case 1, 3, 4, 6, 8, 10, 11: mayor = 4
                            Case 2, 5, 7, 9: mayor = 25
                        End Select
                
                
                        .Cells(Fila, j) = grd.TextMatrix(i, Col)
'                        mayor = Len(grd.TextMatrix(i, Col)) 'Para ajustar el ancho de columnas
                        If mayor > v(j - 1) Then            'de acuerdo a la celda más grande
                            .Columns(j).ColumnWidth = mayor '13/11/2000 ---> Angel P.
                            v(j - 1) = mayor
                        End If
                        j = j + 1
                Next Col
            End If
            .Range(.Cells(Fila, 1), .Cells(Fila, NumCol)).Borders.LineStyle = 1
        Next i
    End With
     prg1.value = prg1.Min
     MensajeStatus "Listo", vbDefault
End Sub

'Guarda resultado
Private Sub GuardarResultado()
    Dim file As String, NumFile As Integer, cadena As String
    Dim Filas As Long, Columnas As Long, i As Long, j As Long
    
    If grd.Rows = grd.FixedRows Then Exit Sub
    On Error GoTo Errtrap
    
        With dlg1
          .CancelError = True
          '.Filter = "Texto (Separado por coma)|*.txt|Excel 97(XLS)|*.xls"
          .Filter = "Texto (Separado por coma)|*.csv"
          .ShowSave
          
          file = .FileName
        End With
    
    If ExisteArchivo(file) Then
        If MsgBox("El nombre del archivo " & file & " ya existe desea sobreescribirlo?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    NumFile = FreeFile
    Open file For Output Access Write As #NumFile
    
    cadena = ""
    For i = 1 To grd.Rows - 1
'    If Not grd.IsSubtotal(i) Then
        For j = 1 To grd.Cols - 1
               cadena = cadena & grd.TextMatrix(i, j) & ","
        Next j
        cadena = Mid(cadena, 1, Len(cadena) - 1)
        Print #NumFile, cadena
        cadena = ""
'     End If
    Next i
    Close NumFile
    MsgBox "El archivo se ha exportado con éxito"
    Exit Sub
Errtrap:
    If Err.Number <> 32755 Then
        MsgBox Err.Description
    End If
    Close NumFile
End Sub

Private Sub AbrirArchivo()
    Dim i As Long
    
    On Error GoTo Errtrap
    With dlg1
        .CancelError = True
        .Filter = "Texto (Separado por coma *.csv)|*.csv|Texto (Separado por tabuladores *.txt)|*.txt|Todos *.*|*.*"
        .Flags = cdlOFNFileMustExist
        If Len(.FileName) = 0 Then          'Solo por primera vez, ubica a la carpeta de la aplicación
            .FileName = App.Path & "\*.csv"
        End If
        
        .ShowOpen
        
        Select Case UCase$(Right$(dlg1.FileName, 4))
        Case ".TXT", ".CSV"
            'ReformartearColumnas
            VisualizarTexto dlg1.FileName
            'InsertarColumnas
        Case ".XLS"
       '     VisualizarExcel dlg1.FileName
        Case Else
        End Select
    End With
    Exit Sub
Errtrap:
    If Err.Number <> 32755 Then DispErr
    Exit Sub
End Sub

Private Sub VisualizarTexto(ByVal archi As String)
    Dim f As Integer, s As String, i As Integer
    Dim cadena
    On Error GoTo Errtrap
    ReDim rec(0, 1)
    MensajeStatus "Está leyendo el archivo " & archi & " ...", vbHourglass
    grd.Rows = grd.FixedRows    'Limpia la grilla
    grd.Redraw = flexRDNone
    grd.MergeCells = flexMergeSpill
    f = FreeFile                'Obtiene número disponible de archivo
    
    'Abre el archivo para lectura
    Open archi For Input As #f
        Do Until EOF(f)
            Line Input #f, s
            s = vbTab & Replace(s, ",", vbTab)      'Convierte ',' a TAB
           grd.AddItem s
        Loop
    Close #f
    RemueveSpace
    ConfigCols
    CambiaFondoCeldasEditables

    grd.SetFocus
    MensajeStatus
    Exit Sub
Errtrap:
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

Private Sub FijarColor()
Dim j As Integer
Dim Col As Integer
Col = grd.Rows - 9
If grd.Rows = grd.FixedCols Then Exit Sub
    For j = 1 To grd.Rows - 10
         If grd.TextMatrix(j, 1) = "." Then
            grd.Cell(flexcpBackColor, j, 1, j, grd.Cols - 1) = &H8000000C  '&HC00000
            grd.Cell(flexcpForeColor, j, 1, j, grd.Cols - 1) = &HFFFF&
        End If
    Next
              grd.TextMatrix(grd.Rows - 4, grd.ColIndex("Descripción")) = GetSetting(APPNAME, App.Title, "costofijomensual", 0)
              grd.Cell(flexcpBackColor, grd.Rows - 4, grd.ColIndex("Descripción"), grd.Rows - 4, grd.ColIndex("Descripción")) = &H80000005  'color blanco
              grd.Cell(flexcpForeColor, grd.Rows - 4, grd.ColIndex("Descripción"), grd.Rows - 4, grd.ColIndex("Descripción")) = &H80000012 'color negro
     
End Sub




'''Private Sub LlenaFormatoFormulario()
'''    With grd
'''        .MergeCells = flexMergeSpill
'''
'''        .AddItem "1" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "DECLARACION DE RETENCIONES EN LA FUENTE DEL IMPUESTO A LA RENTA" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "No."
'''        .AddItem "2" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "100" & vbTab & "IDENTIFICACION DE LA DECLARACION"
'''        .AddItem "3" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "DECLARACION MENSUAL"
'''        .AddItem "4" & vbTab & "FORMULARIO 103" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "101" & vbTab & "MES " & vbTab & "01" & vbTab & "02" & vbTab & "03" & vbTab & "04" & vbTab & "05" & vbTab & "06" & vbTab & "07" & vbTab & "08" & vbTab & "09" & vbTab & "10" & vbTab & "11" & vbTab & "12"
''''        .AddItem "5" & vbTab & "200" & vbTab & "IDENTIFICACION DEL SUJETO PASIVO (AJENTE DE PERCEPCION O RETENCION)"
''''        .AddItem "6" & vbTab & vbTab & "RUC" & vbTab & vbTab & "RAZON SOCIAL DENOMINACION O APELLIDOS Y NOMBRES COMPLETOS"
''''        .AddItem "10" & vbTab & "201" & vbTab & vbTab & "202" & vbTab
''''        '  ORIGINADOS EN EL TRABAJO
''''        .AddItem "19" & vbTab & "300" & vbTab & "POR PAGOS EN EL PAIS" & vbTab & vbTab & vbTab & "% RETENCION" & vbTab & vbTab & "IMPUESTO RETENIDO"
''''        .AddItem "20" & vbTab & "VENTAS LOCALES NETAS (VENTAS BRUTAS MENOS DESCTOS Y DEVOL EN VTAS EXCLUYE ACT. FIJOS Y OTROS)" & vbTab & vbTab & vbTab & "501" & vbTab & vbTab & "531" & vbTab & vbTab & "551"
''''        .AddItem "21" & vbTab & "VENTAS DIRECTAS A EXPORTADORES" & vbTab & vbTab & vbTab & "503" & vbTab & vbTab & "533" & vbTab & vbTab & "553"
''''        .AddItem "22" & vbTab & "VENTAS DE ACTIVOS FIJOS " & vbTab & vbTab & vbTab & "505" & vbTab & vbTab & "535" & vbTab & vbTab & "555"
''''        .AddItem "23" & vbTab & "OTROS (Donaciones promociones autoconsumos etc) " & vbTab & vbTab & vbTab & "507" & vbTab & vbTab & "537" & vbTab & vbTab & "557"
''''        .AddItem "24" & vbTab & "INGRESO POR CONCEPTO DE REEMBOLSO DE GASTOS (INFORMATIVO)" & vbTab & vbTab & vbTab & "509" & vbTab & vbTab & "539" & vbTab & vbTab & "559"
''''        .AddItem "25" & vbTab & "EXPORTACION DE BIENES" & vbTab & vbTab & vbTab & "511"
''''        .AddItem "26" & vbTab & "EXPORTACION DE SERVICIOS" & vbTab & vbTab & vbTab & "513"
''''        .AddItem "27" & vbTab & "TOTAL VENTAS Y EXPORTACIONES" & vbTab & vbTab & "501+503-505+507+511+513+531+533+535+537" & vbTab & vbTab & vbTab & "549"
''''        .AddItem "28" & vbTab & "IVA PRESUNTIVO DE SALAS DE JUEGO (BINGO-MECANICOS) Y OTROS JUEGOS DE AZAR" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "597"
''''        .AddItem "29" & vbTab & "TOTAL IMPUESTO" & vbTab & vbTab & vbTab & vbTab & "551+553+555+557+597" & vbTab & vbTab & vbTab & "599"
''''        .AddItem "30" & vbTab & "TOTAL FACTURAS EMITIDAS" & vbTab & vbTab & "105" & vbTab & vbTab & "OTROS DOCUMENTOS AUTORIZADOS EMITIDOS" & vbTab & "107" & vbTab & vbTab & "TOTAL NOTAS DE CREDITO EMITIDAS" & vbTab & vbTab & "109"
''''        .AddItem "31" & vbTab & "TOTAL NOTA VENTA EMITIDAS" & vbTab & vbTab & "106" & vbTab & vbTab & "TOTAL DOC.ADUANEROS (EXP)" & vbTab & "108" & vbTab & vbTab & "TOTAL NOTAS DE DEBITO EMITIDAS" & vbTab & vbTab & "110"
'''''        'COMPRAS E IMPORTACIONES
''''        .AddItem "32" & vbTab & "600" & vbTab & "RESUMEN DE COMPRAS Y OTRAS OPERACIONES DEL PERIODO QUE DECLARA" & vbTab & vbTab & vbTab & "BASE IMPONIBLE 0%" & vbTab & vbTab & "BASE IMPONIBLE 12%" & vbTab & "IMPUESTO"
''''        .AddItem "33" & vbTab & "COMPRAS LOCALES NETAS DE BIENES (COMPRAS BRUTAS MENOS DESCTOS Y DEVOL EXCLUYE ACT. FIJOS )" & vbTab & vbTab & vbTab & "601" & vbTab & vbTab & "631" & vbTab & vbTab & "651"
''''        .AddItem "34" & vbTab & "COMPRAS LOCALES DE SERVICIOS" & vbTab & vbTab & vbTab & "603" & vbTab & vbTab & "633" & vbTab & vbTab & "653"
''''        .AddItem "35" & vbTab & "COMPRAS LOCALES DE ACTIVOS FIJOS" & vbTab & vbTab & vbTab & "605" & vbTab & vbTab & "635" & vbTab & vbTab & "655"
''''        .AddItem "36" & vbTab & "PAGO POR CONCEPTO DE REEMBOLSO DE GASTOS (INFORMATIVO)" & vbTab & vbTab & vbTab & "607" & vbTab & vbTab & "637" & vbTab & vbTab & "657"
''''        .AddItem "37" & vbTab & "IMPORTACIONES DE BIENES (EXCLUYE ACT. FIJO)" & vbTab & vbTab & vbTab & "609" & vbTab & vbTab & "639" & vbTab & vbTab & "659"
''''        .AddItem "38" & vbTab & "IMPORTACIONES DE SERVICIOS" & vbTab & vbTab & vbTab & "611" & vbTab & vbTab & "641" & vbTab & vbTab & "661"
''''        .AddItem "39" & vbTab & "IMPORTACIONES DE ACTIVOS FIJOS" & vbTab & vbTab & vbTab & "613" & vbTab & vbTab & "643" & vbTab & vbTab & "663"
''''        .AddItem "40" & vbTab & "IVA SOBRE VALOR DE LA DEPRECIACION DE ACTIVOS EN INTERACION TEMPORAL" & vbTab & vbTab & vbTab & "." & vbTab & vbTab & "645" & vbTab & vbTab & "665"
''''        .AddItem "41" & vbTab & "IVA EN ARRENDAMIENTO MERCANTIL INTERNACIONAL" & vbTab & vbTab & vbTab & "." & vbTab & vbTab & "647" & vbTab & vbTab & "667"
''''        .AddItem "42" & vbTab & "COMPRA DE BIENES O SERVICIOS CON COMPROBANTES QUE NO SUSTENTAN CREDITO TRIBUTARIO" & vbTab & vbTab & vbTab & "619" & vbTab & vbTab & "649"
''''        .AddItem "43" & vbTab & "TOTAL COMPRAS E IMPORTACIONES" & vbTab & vbTab & "601+603-605+609+611+613+619+631+633+635+639+641+643+645+647+649" & vbTab & vbTab & vbTab & "650"
''''        .AddItem "44" & vbTab & "CREDITO TRIBUTARIO DE ACUERDO A CONTABILIDAD O A REGISTRO DE INGRESOS Y GASTO" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "651+653+655+659+661+663+665+667" & vbTab & "698"
''''        .AddItem "45" & vbTab & "CREDITO TRIBUTARIO DE ACUERDO AL FACTOR DE PROPORCIONALIDAD" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "(651+653+655+659+661+663+665+667) x 301" & vbTab & "699"
''''        .AddItem "46" & vbTab & "TOTAL FACTURAS RECIBIDAS" & vbTab & vbTab & "111"
''''        .AddItem "47" & vbTab & "TOTAL NOTA VENTA RECIBIDAS" & vbTab & vbTab & "112" & vbTab & vbTab & "OTROS DOCUMENTOS AUTORIZADOS RECIBIDOS" & vbTab & "114" & vbTab & vbTab & "TOTAL NOTAS DE CREDITO RECIBIDAS" & vbTab & vbTab & "116"
''''        .AddItem "48" & vbTab & "TOTAL LIQUID. DE COMPRAS" & vbTab & vbTab & "113" & vbTab & vbTab & "TOTAL DOC.ADUANEROS (IMP)" & vbTab & "115" & vbTab & vbTab & "TOTAL NOTAS DE DEBITO RECIBIDAS" & vbTab & vbTab & "117"
'''''        'RESUMEN
''''        .AddItem "49" & vbTab & "700" & vbTab & "RESUMEN IMPOSITIVO"
''''        .AddItem "50" & vbTab & "IMPUESTO CAUSADO" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "599-698 ó 599-699>0" & vbTab & "701"
''''        .AddItem "51" & vbTab & "(-) CREDITO TRIBUTARIO DEL MES " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "599-698 ó 599-699<0" & vbTab & "702"
''''        .AddItem "52" & vbTab & "(-) SALDO DE CREDITO TRIBUTARIO APLICARSE EN ESTE MES " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "Trasladar campo 399" & vbTab & "703"
''''        .AddItem "53" & vbTab & "(-) RETENCIONES EN LA FUENTE DE IVA QUE LE HAN SIDO EFECTUADAS " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "705"
''''        .AddItem "54" & vbTab & "(=) SALDO DE CREDITO TRIBUTARIO PARA EL PROXIMO MES " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "701-702-703-705 < 0" & vbTab & "798"
''''        .AddItem "55" & vbTab & "(=) SUBTOTAL A PAGAR " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "701-702-703-705 > 0" & vbTab & "799"
'''' '        'DECLARACION DEL SUJETO PASIVO COMO AGENTE RETENCION IVA
''''        .AddItem "56" & vbTab & "800" & vbTab & "DECLARACION DEL SUJETO PASIVO COMO AGENTE RETENCION IVA" & vbTab & vbTab & vbTab & vbTab & vbTab & "VALOR DEL IVA" & vbTab & vbTab & "VALOR RETENIDO"
''''        .AddItem "57" & vbTab & "IVA CAUSADO POR LA PRESTACION DE SERVICIOS PROFESIONALES" & vbTab & vbTab & vbTab & vbTab & vbTab & "801" & vbTab & vbTab & "851" & vbTab & vbTab & "100%"
''''        .AddItem "58" & vbTab & "IVA CAUSADO POR EL ARRENDAMIENTO DE INMUEBLES A PERSONAS NATURALES" & vbTab & vbTab & vbTab & vbTab & vbTab & "803" & vbTab & vbTab & "853" & vbTab & vbTab & "100%"
''''        .AddItem "59" & vbTab & "IVA CAUSADO EN OTRAS COMPRA DE BIENES Y SERVICIOS CON EMISION DE LIQUIDACION DE COMPRAS Y PRESTACION SERVICIOS" & vbTab & vbTab & vbTab & vbTab & vbTab & "805" & vbTab & vbTab & "855" & vbTab & vbTab & "100%"
''''        .AddItem "60" & vbTab & "IVA CAUSADO EN LA DEPRECIACION DE ACTIVOS EN INTERNACION TEMPORAL" & vbTab & vbTab & vbTab & vbTab & vbTab & "807" & vbTab & vbTab & "807" & vbTab & vbTab & "100%"
''''        .AddItem "61" & vbTab & "IVA CAUSADO EN LA DISTRIBUCION DE COMBUSTIBLES" & vbTab & vbTab & vbTab & vbTab & vbTab & "809" & vbTab & vbTab & "859" & vbTab & vbTab & "100%"
''''        .AddItem "62" & vbTab & "IVA CAUSADO EN LEASING INTERNACIONAL" & vbTab & vbTab & vbTab & vbTab & vbTab & "811" & vbTab & vbTab & "861" & vbTab & vbTab & "100%"
''''        .AddItem "63" & vbTab & "IVA CAUSADO POR LA PRESTACION DE SERVICIOS" & vbTab & vbTab & vbTab & vbTab & vbTab & "813" & vbTab & vbTab & "863" & vbTab & vbTab & "70%"
''''        .AddItem "64" & vbTab & "IVA RETENIDO POR EMPRESAS EMISORAS DE TARJETAS DE CREDITO SERVICIOS" & vbTab & vbTab & vbTab & vbTab & vbTab & "815" & vbTab & vbTab & "865" & vbTab & vbTab & "70%"
''''        .AddItem "65" & vbTab & "IVA RETENIDO POR EMPRESAS EMISORAS DE TARJETAS DE CREDITO BIENES" & vbTab & vbTab & vbTab & vbTab & vbTab & "817" & vbTab & vbTab & "867" & vbTab & vbTab & "30%"
''''        .AddItem "66" & vbTab & "IVA POR LA COMPRA DE BIENES" & vbTab & vbTab & vbTab & vbTab & vbTab & "819" & vbTab & vbTab & "869" & vbTab & vbTab & "30%"
''''        .AddItem "67" & vbTab & "IVA EN CONTRATOS DE CONSTRUCCION" & vbTab & vbTab & vbTab & vbTab & vbTab & "821" & vbTab & vbTab & "871" & vbTab & vbTab & "30%"
''''        .AddItem "68" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "TOTAL RETENIDO            SUMAR 851 AL 871 " & vbTab & vbTab & "898"
''''        .AddItem "69" & vbTab & "COMPROBANTES DE RETENCION EMITIDOS" & vbTab & vbTab & vbTab & "118" & vbTab & vbTab & "TOTAL IVA A PAGAR            799+898 " & vbTab & vbTab & "899"
'''''        'A PAGAR
''''        .AddItem "70" & vbTab & "Declaro que los datos contenidos en esta declaración son verdaderos por lo que asumo la responsabilidad correspondiente (Artículo 98 de la L.R.T.I.) " & vbTab & vbTab & vbTab & vbTab & vbTab & "900" & vbTab & "VALORES A PAGAR Y FORMA DE PAGO"
''''        .AddItem "71"
''''        .AddItem "72" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "PAGO PREVIO" & vbTab & vbTab & "901"
''''        .AddItem "73" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "TOTAL IMPUESTO A PAGAR 899-901" & vbTab & vbTab & "902"
''''        .AddItem "74" & vbTab & vbTab & "____________________________" & vbTab & vbTab & vbTab & "____________________________" & vbTab & "INTERESES POR MORA" & vbTab & vbTab & "903"
''''        .AddItem "75" & vbTab & vbTab & "FIRMA SUJETO PASIVO" & vbTab & vbTab & vbTab & "FIRMA CONTADOR" & vbTab & "MULTAS" & vbTab & vbTab & "904"
''''        .AddItem "76" & vbTab & "NOMBRE:" & vbTab & vbTab & "NOMBRE:" & vbTab & "" & vbTab & vbTab & "TOTAL PAGADO     902+903+904" & vbTab & vbTab & "999"
''''        .AddItem "77" & vbTab & "198" & vbTab & "C.I. No." & vbTab & "199" & vbTab & "RUC No." & vbTab & "."
''''        .AddItem "78" & vbTab & "MEDIANTE CHEQUE DEBITO BANCARIO EFECTIVO U OTRAS FORMAS DE COBRO" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "905"
''''        .AddItem "79" & vbTab & "MEDIANTE COMPENSACIONES" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "906"
''''        .AddItem "80" & vbTab & "MEDIANTE NOTAS DE CREDITO" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "907"
'''        .Redraw = flexRDBuffered
'''        .Refresh
'''    End With
''''    CambiaFondoCeldasEditables
'''End Sub

Private Sub LLENADATOS()
    With grd
        .Redraw = flexRDBuffered
        .TextMatrix(10, 2) = gobjMain.EmpresaActual.GNOpcion.RUC
        .TextMatrix(10, 4) = gobjMain.EmpresaActual.GNOpcion.NombreEmpresa
         .Refresh
    End With

End Sub

Private Sub CambiaFondoCeldasEditables()
    With grd
        .Redraw = flexRDBuffered
        'fondo blanco
        .Cell(flexcpBackColor, 4, 6, 4, 6) = vbWhite ' formulario modifica
        .Cell(flexcpBackColor, 4, 10, 4, 10) = vbWhite ' formulario modifica
        .Cell(flexcpBackColor, 5, 10, 5, 10) = vbWhite ' formulario modifica
        .Cell(flexcpBackColor, 8, 2, 8, 2) = vbWhite 'ruc
        .Cell(flexcpBackColor, 8, 4, 8, 9) = vbWhite 'razon social
        .Cell(flexcpBackColor, 10, 7, 41, 7) = vbWhite
        .Cell(flexcpBackColor, 43, 4, 45, 4) = vbWhite
        .Cell(flexcpBackColor, 45, 7, 45, 7) = vbWhite
        .Cell(flexcpBackColor, 47, 7, 58, 7) = vbWhite
        .Cell(flexcpBackColor, 62, 10, 62, 10) = vbWhite
        .Cell(flexcpBackColor, 64, 10, 65, 10) = vbWhite
        .Cell(flexcpBackColor, 67, 10, 68, 10) = vbWhite
        .Cell(flexcpBackColor, 65, 2, 66, 2) = vbWhite
        .Cell(flexcpBackColor, 65, 4, 66, 5) = vbWhite
        .Cell(flexcpBackColor, 70, 2, 73, 2) = vbWhite
        .Cell(flexcpBackColor, 70, 5, 73, 5) = vbWhite
        
        'fondo 1
        .Cell(flexcpBackColor, 2, 4, 2, 9) = ColorFondo1
        .Cell(flexcpBackColor, 6, 1, 6, 9) = ColorFondo1
        .Cell(flexcpBackColor, 9, 1, 9, 9) = ColorFondo1
        .Cell(flexcpBackColor, 46, 1, 46, 9) = ColorFondo1
        .Cell(flexcpBackColor, 61, 6, 61, 9) = ColorFondo1
        'fondo
        .Cell(flexcpBackColor, 4, 4, 4, 4) = ColorFondo
        .Cell(flexcpBackColor, 5, 6, 5, 6) = ColorFondo
        .Cell(flexcpBackColor, 4, 8, 4, 8) = ColorFondo
        .Cell(flexcpBackColor, 6, 4, 6, 4) = ColorFondo
        .Cell(flexcpBackColor, 7, 1, 8, 1) = ColorFondo
        .Cell(flexcpBackColor, 7, 3, 8, 3) = ColorFondo
        .Cell(flexcpBackColor, 10, 1, 14, 1) = ColorFondo
        .Cell(flexcpBackColor, 10, 8, 10, 10) = ColorFondo
        .Cell(flexcpBackColor, 11, 8, 11, 8) = ColorFondo
        .Cell(flexcpBackColor, 10, 6, 41, 6) = ColorFondo
        .Cell(flexcpBackColor, 10, 9, 42, 9) = ColorFondo
        .Cell(flexcpBackColor, 17, 8, 17, 10) = ColorFondo
        .Cell(flexcpBackColor, 36, 1, 37, 1) = ColorFondo
        .Cell(flexcpBackColor, 41, 9, 41, 10) = ColorFondo
        .Cell(flexcpBackColor, 43, 3, 45, 3) = ColorFondo
        .Cell(flexcpBackColor, 45, 6, 45, 6) = ColorFondo
        .Cell(flexcpBackColor, 48, 1, 58, 1) = ColorFondo
        .Cell(flexcpBackColor, 47, 6, 58, 6) = ColorFondo
        .Cell(flexcpBackColor, 47, 9, 60, 9) = ColorFondo
        .Cell(flexcpBackColor, 47, 1, 47, 6) = ColorFondo
        .Cell(flexcpBackColor, 47, 8, 47, 8) = ColorFondo
        .Cell(flexcpBackColor, 62, 9, 66, 9) = ColorFondo
        .Cell(flexcpBackColor, 66, 1, 66, 1) = ColorFondo
        .Cell(flexcpBackColor, 66, 3, 66, 3) = ColorFondo
        .Cell(flexcpBackColor, 67, 9, 68, 9) = ColorFondo
        .Cell(flexcpBackColor, 70, 1, 73, 1) = ColorFondo
        .Cell(flexcpBackColor, 70, 3, 73, 3) = ColorFondo
        'alineacion
        .Cell(flexcpAlignment, 7, 4, 7, 4) = 1
        .Cell(flexcpAlignment, 10, 1, 14, 1) = 3
        .Cell(flexcpAlignment, 59, 7, 66, 7) = 1
        .Cell(flexcpAlignment, 66, 4, 66, 4) = 1
        .Cell(flexcpAlignment, 45, 5, 45, 5) = 1
        .Cell(flexcpAlignment, 61, 6, 61, 6) = 1
        
        'alto linea
        .RowHeight(1) = 400
        .RowHeight(2) = 320
        .RowHeight(4) = 320
        .RowHeight(6) = 320
        .RowHeight(9) = 320
        .RowHeight(46) = 320
        .RowHeight(61) = 400
       
        'tamaño letras
        .Cell(flexcpFontSize, 1, 1, 1, 6) = 14
        .Cell(flexcpFontSize, 2, 1, 2, 8) = 10
        .Cell(flexcpFontSize, 4, 2, 4, 2) = 11
        .Cell(flexcpFontSize, 6, 1, 6, 8) = 10
        .Cell(flexcpFontSize, 9, 1, 9, 10) = 10
        .Cell(flexcpFontSize, 46, 1, 46, 10) = 10
        .Cell(flexcpFontSize, 61, 6, 61, 8) = 10
        
        'negritas
        .Cell(flexcpFontBold, 2, 3, 2, 9) = True
        .Cell(flexcpFontBold, 4, 2, 4, 2) = True
        .Cell(flexcpFontBold, 6, 1, 6, 9) = True
        .Cell(flexcpFontBold, 9, 1, 9, 10) = True
        .Cell(flexcpFontBold, 46, 1, 46, 10) = True
        .Cell(flexcpFontBold, 61, 6, 61, 9) = True
'        .Cell(flexcpFontBold, 10, 1, 10, 1) = True
'        .Cell(flexcpFontBold, 63, 1, 63, 1) = True
'        .Cell(flexcpFontBold, 1, 4, 67, 4) = True
'        .Cell(flexcpFontBold, 30, 3, 31, 3) = True
'        .Cell(flexcpFontBold, 30, 4, 31, 4) = False
'        .Cell(flexcpFontBold, 20, 6, 55, 6) = True
'        .Cell(flexcpFontBold, 12, 8, 55, 8) = True
'        .Cell(flexcpFontBold, 30, 8, 31, 8) = False
'        .Cell(flexcpFontBold, 30, 10, 31, 10) = True
'        .Cell(flexcpFontBold, 65, 5, 67, 5) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .MergeCol(4) = True
        .MergeCol(5) = True
    End With
End Sub

Private Sub Buscar()
    Dim sql As String, cond As String, CadenaValores As String
    Dim OrdenadoX As String
    Dim CadenaAgrupa  As String, Recargo As String
    Dim v As Variant, max As Integer, i As Integer
    Dim from As String, NumReg As Long, f1 As String
    Dim rs As Recordset
    Dim subtotal As Currency, CompraBienesTarifa_0 As Currency, CompraServiciosTarifa_0 As Currency
    Dim CompraActivosTarifa_0 As Currency
    Dim CP_Ser As String, CP_Act As String, CP_Dev As String
    Dim VT_Bie As String, VT_Ser As String, VT_Dev As String
    Dim Reten As String, ret_real As String, ret_recib As String
    Dim MONEDA As String
    Set objcond = gobjMain.objCondicion
    If Not frmB_FormSRI103.Inicio103(objcond, Reten) Then
        grd.SetFocus
        Exit Sub
    End If
    With objcond
        If Len(Month(.Fecha1)) < 2 Then
            grd.TextMatrix(4, 6) = "0" & Month(.Fecha1)
        Else
            grd.TextMatrix(4, 6) = Month(.Fecha1)
        End If
        
        
        grd.TextMatrix(4, 9) = " AÑO " & Year(.Fecha1)
           
        'Reporte de un mes a la vez
        f1 = DateSerial(Year(.Fecha1), Month(.Fecha1), 1)
        cond = " AND GNC.FechaTrans BETWEEN " & FechaYMD(f1, gobjMain.EmpresaActual.TipoDB) & _
               " AND " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB)

            
            
            
    '********** VENTAS BIENES
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 "
            sql = sql & "From IvRecargo ivR inner join "
            sql = sql & " IvKardexRecargo ivkR Inner join "
            sql = sql & " GnComprobante gnc Inner join PcPRovCLi on gnc.IdClienteRef = PCProvCli.IdProvCli "
            sql = sql & " On ivkr.TransID = gnc.TransID "
            sql = sql & " On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gnc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(Recargo) & ") " & cond
            sql = sql & " Group by IvkR.TransID"
            
            VerificaExistenciaTabla 0
            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg

            sql = "SELECT "
            sql = sql & " isnull(SUM(PrecioTotalBase0 + (PrecioTotalBase0 * (cast(isnull(TotalDescuento,0) as float) / cast(PrecioTotal as float)))),0) as Base0, "
            sql = sql & " isnull(SUM(PrecioTotalBaseIVA + (PrecioTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(PrecioTotal as float)))),0) As BaseIVA "
            sql = sql & " FROM tmp0 Right join "
            sql = sql & "vwConsSUMIVKardexIVA inner join "
            sql = sql & "GNComprobante GNC   "
            sql = sql & "ON vwConsSUMIVKardexIVA.TransID = GNC.TransID "
            sql = sql & "ON tmp0.TransID = GNC.TransID"
            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(VT_Bie) & ")  " & cond
            
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            grd.TextMatrix(20, 5) = Round(rs.Fields("Base0"), 2)
            grd.TextMatrix(20, 7) = Round(rs.Fields("BaseIVA"), 2)
            
            
            
    '********** VENTAS ACtivos
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 "
            sql = sql & "From IvRecargo ivR inner join "
            sql = sql & " IvKardexRecargo ivkR Inner join "
            sql = sql & " GnComprobante gnc Inner join PcPRovCLi on gnc.IdClienteRef = PCProvCli.IdProvCli "
            sql = sql & " On ivkr.TransID = gnc.TransID "
            sql = sql & " On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gnc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(Recargo) & ") " & cond
            sql = sql & " Group by IvkR.TransID"
            
            VerificaExistenciaTabla 0
            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg

            sql = "SELECT "
            sql = sql & " isnull(SUM(PrecioTotalBase0 + (PrecioTotalBase0 * (cast(isnull(TotalDescuento,0) as float) / cast(PrecioTotal as float)))),0) as Base0, "
            sql = sql & " isnull(SUM(PrecioTotalBaseIVA + (PrecioTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(PrecioTotal as float)))),0) As BaseIVA "
            sql = sql & " FROM tmp0 Right join "
            sql = sql & "vwConsSUMIVKardexIVA inner join "
            sql = sql & "GNComprobante GNC   "
            sql = sql & "ON vwConsSUMIVKardexIVA.TransID = GNC.TransID "
            sql = sql & "ON tmp0.TransID = GNC.TransID"
            sql = sql & " WHERE GNC.Estado<>3  AND GNC.CodTrans IN (" & PreparaCadena(VT_Ser) & ")" & cond
            
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            grd.TextMatrix(22, 5) = Round(rs.Fields("Base0"), 2)
            grd.TextMatrix(22, 7) = Round(rs.Fields("BaseIVA"), 2)
            
            ' cantidad de comprobantes
            sql = "SELECT "
            sql = sql & " AnexoCodTipoComp, count(gnc.codtrans) as NumComp  "
            sql = sql & " FROM  gncomprobante gnc inner join gntrans gnt "
            sql = sql & " on gnc.codtrans=gnt.codtrans"
            sql = sql & " WHERE ( GNC.CodTrans IN (" & PreparaCadena(VT_Bie) & " ) "
            sql = sql & " or GNC.CodTrans IN (" & PreparaCadena(VT_Ser) & "))  " & cond
            sql = sql & "group by AnexoCodTipoComp"
            
            
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            
            rs.MoveFirst
            While Not rs.EOF
                Select Case rs.Fields("AnexoCodTipoComp")
                    Case 1:
                        grd.TextMatrix(30, 4) = rs.Fields("NumComp")
                    Case 2:
                        grd.TextMatrix(31, 4) = rs.Fields("NumComp")
                    Case 18:
                        grd.TextMatrix(30, 7) = rs.Fields("NumComp")
                    Case 4:
                        grd.TextMatrix(30, 11) = rs.Fields("NumComp")
                    Case 5:
                        grd.TextMatrix(31, 11) = rs.Fields("NumComp")
                End Select
                rs.MoveNext
            Wend
            
            
            
            
            
            
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(.CodForma) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(.CodTrans) & ")" & _
                  "Group by IvkR.TransID"
                  
            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
            
'********** compras BIENES
            
            
            '--datos de la compra bienes tarifa 12
            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0, "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBaseIVA + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS Valor12 "
            sql = sql & " Into tmp1"
            sql = sql & " from    (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(.CodTrans) & ")"
            sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
            sql = " Select  isnull(sum(Valor0),0) as ValorTotal0, isnull(sum(Valor12),0) as ValorTotal12 from tmp1  "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            grd.TextMatrix(33, 5) = Round(rs.Fields("ValorTotal0"), 2)
            grd.TextMatrix(33, 7) = Round(rs.Fields("ValorTotal12"), 2)
            
            '--datos de la compra servicios tarifa 12
            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0, "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBaseIVA + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS Valor12 "
            sql = sql & " Into tmp1"
            sql = sql & " from    (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(CP_Ser) & ")"
            sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
            sql = " Select  isnull(sum(Valor0),0) as ValorTotal0, isnull(sum(Valor12),0) as ValorTotal12 from tmp1  "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            grd.TextMatrix(34, 5) = Round(rs.Fields("ValorTotal0"), 2)
            grd.TextMatrix(34, 7) = Round(rs.Fields("ValorTotal12"), 2)


        '--datos de la compra activos tarifa 12
            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0, "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBaseIVA + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS Valor12 "
            sql = sql & " Into tmp1"
            sql = sql & " from    (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(CP_Act) & ")"
            sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1

            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
            sql = " Select  isnull(sum(Valor0),0) as ValorTotal0, isnull(sum(Valor12),0) as ValorTotal12 from tmp1  "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            grd.TextMatrix(35, 5) = Round(rs.Fields("ValorTotal0"), 2)
            grd.TextMatrix(35, 7) = Round(rs.Fields("ValorTotal12"), 2)


        '--datos de la total compra que no sustentan cretito tributario sustento 02y 07
            sql = " Select  "
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0, "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBaseIVA + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS Valor12 "
            sql = sql & " Into tmp1"
            sql = sql & " from    (( tmp0 Right join gncomprobante Gnc inner join Anexos  on anexos.transid=gnc.transid "
            sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  (GNC.CodTrans IN (" & PreparaCadena(.CodTrans) & ")"
            sql = sql & " or  GNC.CodTrans IN (" & PreparaCadena(CP_Act) & ")"
            sql = sql & " or  GNC.CodTrans IN (" & PreparaCadena(CP_Ser) & "))"
            sql = sql & " and  anexos.CodCredTrib in ('02','07')"
            sql = sql & " and GNC.Estado<>3 " & cond
            VerificaExistenciaTabla 1

            gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
            sql = " Select  isnull(sum(Valor0),0) as ValorTotal0, isnull(sum(Valor12),0) as ValorTotal12 from tmp1  "
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            grd.TextMatrix(42, 5) = Round(rs.Fields("ValorTotal0") + rs.Fields("ValorTotal12"), 2)

' cantidad de comprobantes compras
            sql = "SELECT "
            sql = sql & " CodTipoComp, count(gnc.codtrans) as NumComp  "
            sql = sql & " FROM  gncomprobante gnc inner join anexos ane  on gnc.transid=ane.transid "
            sql = sql & " WHERE ( GNC.CodTrans IN (" & PreparaCadena(.CodTrans) & " ) "
            sql = sql & " or GNC.CodTrans IN (" & PreparaCadena(CP_Ser) & ")  "
            sql = sql & " or GNC.CodTrans IN (" & PreparaCadena(CP_Act) & "))  " & cond
            sql = sql & " group by CodTipoComp"
            
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            
            rs.MoveFirst
            While Not rs.EOF
                Select Case rs.Fields("CodTipoComp")
                    Case "1":
                        grd.TextMatrix(46, 4) = rs.Fields("NumComp")
                    Case "01":
                        grd.TextMatrix(46, 4) = grd.ValueMatrix(46, 4) + rs.Fields("NumComp")
                    Case "2":
                        grd.TextMatrix(47, 4) = rs.Fields("NumComp")
                    Case "02"
                        grd.TextMatrix(47, 4) = grd.ValueMatrix(47, 4) + rs.Fields("NumComp")
                    Case "3":
                        grd.TextMatrix(48, 4) = rs.Fields("NumComp")
                    Case "03"
                        grd.TextMatrix(48, 4) = grd.ValueMatrix(48, 4) + rs.Fields("NumComp")
                    Case "4":
                        grd.TextMatrix(47, 11) = rs.Fields("NumComp")
                    Case "04"
                        grd.TextMatrix(47, 11) = grd.ValueMatrix(47, 11) + rs.Fields("NumComp")
                    Case "5":
                        grd.TextMatrix(48, 11) = rs.Fields("NumComp")
                    Case "05"
                        grd.TextMatrix(48, 11) = grd.ValueMatrix(48, 11) + rs.Fields("NumComp")
                    Case Else
                        grd.TextMatrix(47, 7) = grd.ValueMatrix(47, 7) + rs.Fields("NumComp")
                End Select
                rs.MoveNext
            Wend



                            
                       
            ' RETENCIONES RECIBIDAS
            .CodMoneda = MONEDA_PRE
            MONEDA = IIf(.NumMoneda > 0, "/Cotizacion" & .NumMoneda + 1, "")
            sql = "SELECT ISNULL(sum(Valor" & MONEDA & "),0) as TotalRetRecibidas "
            sql = sql & " FROM vwConsRetencion "
            sql = sql & " WHERE CodTrans IN (" & PreparaCadena(ret_recib) & ")"
            sql = sql & "  AND FechaTrans BETWEEN " & FechaYMD(f1, gobjMain.EmpresaActual.TipoDB)
            sql = sql & "  AND " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB)
            sql = sql & "  and  DEBE > 0 and bandIVA=1"

            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)

            grd.TextMatrix(53, 9) = Round(rs.Fields("TotalRetRecibidas"), 2)
'''
            ' RETENCIONES REALIZADAS
            sql = "SELECT CodF104," '
            sql = sql & " ISNULL(sum(Base),0) as TotalBase "
            sql = sql & " FROM vwConsRetencion "
            sql = sql & " WHERE CodTrans IN (" & PreparaCadena(ret_real) & ")"
            sql = sql & "  AND FechaTrans BETWEEN " & FechaYMD(f1, gobjMain.EmpresaActual.TipoDB)
            sql = sql & "  AND " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB)
            sql = sql & "  and  HABER > 0 and bandIVA=1 "
            sql = sql & "  group by Codf104" ', Porcentaje"

            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            rs.MoveFirst
            While Not rs.EOF
                Select Case rs.Fields("CodF104")
                    Case "801":
                        grd.TextMatrix(57, 7) = Round(rs.Fields("TotalBase"), 2)
                    Case "803":
                        grd.TextMatrix(58, 7) = Round(rs.Fields("TotalBase"), 2)
                    Case "805":
                        grd.TextMatrix(59, 7) = Round(rs.Fields("TotalBase"), 2)
                    Case "807":
                        grd.TextMatrix(60, 7) = Round(rs.Fields("TotalBase"), 2)
                    Case "809":
                        grd.TextMatrix(61, 7) = Round(rs.Fields("TotalBase"), 2)
                    Case "811":
                        grd.TextMatrix(62, 7) = Round(rs.Fields("TotalBase"), 2)
                    Case "813":
                        grd.TextMatrix(63, 7) = Round(rs.Fields("TotalBase"), 2)
                    Case "815":
                        grd.TextMatrix(64, 7) = Round(rs.Fields("TotalBase"), 2)
                    Case "817":
                        grd.TextMatrix(65, 7) = Round(rs.Fields("TotalBase"), 2)
                    Case "819":
                        grd.TextMatrix(66, 7) = Round(rs.Fields("TotalBase"), 2)
                    Case "821":
                        grd.TextMatrix(67, 7) = Round(rs.Fields("TotalBase"), 2)
                End Select
                rs.MoveNext
            Wend
                        
            'calcula numero de retenciones
            sql = "SELECT count(Trans) as numTrans "
            sql = sql & " FROM vwConsRetencion "
            sql = sql & " WHERE CodTrans IN (" & PreparaCadena(ret_real) & ")"
            sql = sql & "  AND FechaTrans BETWEEN " & FechaYMD(f1, gobjMain.EmpresaActual.TipoDB)
            sql = sql & "  AND " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB)
            sql = sql & "  and  HABER > 0 and bandIVA=1 "
'            sql = sql & "  group by Codf104" ', Porcentaje"

            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            grd.TextMatrix(69, 5) = rs.Fields("numTrans")
            
            CalcularPorcentajes
            grd.Refresh
    End With
End Sub
    

Private Sub CalcularPorcentajes()
'    Dim i As Integer, SubtotalCompras As Currency, SubtotalVentas As Currency, TotalVentas As Currency
'    Dim SubtotalRetenido As Currency, TotalCompras As Currency
'
'    'mes anterior credito tributario 399=303-305
'
'    grd.TextMatrix(16, 5) = grd.ValueMatrix(13, 5) - grd.ValueMatrix(14, 5) + grd.ValueMatrix(15, 5)
'
'    'mes anterior credito tributario 703=399
'    grd.TextMatrix(52, 9) = grd.ValueMatrix(13, 5)
'
'    TotalVentas = 0
'    'ventas 549
'
'    For i = 20 To 23
'            TotalVentas = TotalVentas + grd.ValueMatrix(i, 5) + grd.ValueMatrix(i, 7)
'    Next i
'    TotalVentas = TotalVentas + grd.ValueMatrix(25, 5) + grd.ValueMatrix(26, 5)
'   grd.TextMatrix(27, 7) = Round(TotalVentas, 2)
'
'    SubtotalVentas = 0
'    'ventas 599
'    For i = 20 To 24
'        If Len(grd.TextMatrix(i, 7)) > 0 Then
'            grd.TextMatrix(i, 9) = Round(grd.ValueMatrix(i, 7) * gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA, 2)
'            SubtotalVentas = SubtotalVentas + grd.ValueMatrix(i, 9)
'        Else
'            grd.TextMatrix(i, 9) = ""
'        End If
'    Next i
'   grd.TextMatrix(29, 9) = Round(SubtotalVentas + grd.ValueMatrix(28, 9), 2)
'
'    'calculo 301= (511+513-531+533+537)/549
'    If grd.ValueMatrix(27, 7) <> 0 Then
'        grd.TextMatrix(12, 5) = Round((grd.ValueMatrix(25, 6) + grd.ValueMatrix(26, 6) + grd.ValueMatrix(20, 7) + grd.ValueMatrix(21, 7) + grd.ValueMatrix(23, 7)) / grd.ValueMatrix(27, 7), 3)
'    End If
'
'
'    'compras 699
'    For i = 33 To 41
'        If Len(grd.TextMatrix(i, 7)) > 0 Then
'            grd.TextMatrix(i, 9) = Round(grd.ValueMatrix(i, 7) * gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA, 2)
'            SubtotalCompras = SubtotalCompras + grd.ValueMatrix(i, 9)
'        Else
'            grd.TextMatrix(i, 9) = ""
'        End If
'    Next i
'    grd.TextMatrix(44, 9) = Round(SubtotalCompras, 2)
'    grd.TextMatrix(45, 9) = Round(SubtotalCompras * grd.ValueMatrix(12, 5), 2)
'
'    'compras 650
'    TotalCompras = 0
'    For i = 33 To 42
'            If i = 36 Then i = 37
'            TotalCompras = TotalCompras + grd.ValueMatrix(i, 5) + grd.ValueMatrix(i, 7)
'    Next i
'    grd.TextMatrix(43, 7) = Round(TotalCompras, 2)
'
''
'    ' 701=599-698 ó 599-699
'
'    If (grd.ValueMatrix(29, 9) - grd.ValueMatrix(45, 9)) > 0 Then
'        '701
'        grd.TextMatrix(50, 9) = (grd.ValueMatrix(29, 9) - grd.ValueMatrix(45, 9))
'        '702
'        grd.TextMatrix(51, 9) = 0
'    Else
'        '701
'        grd.TextMatrix(50, 9) = 0
'        '702
'        grd.TextMatrix(51, 9) = Abs((grd.ValueMatrix(29, 9) - grd.ValueMatrix(45, 9)))
'    End If
'    '798= 701-702-703-704<0
'    If (grd.ValueMatrix(50, 9) - grd.ValueMatrix(51, 9) - grd.ValueMatrix(52, 9) - grd.ValueMatrix(53, 9)) < 0 Then
'        '798
'        grd.TextMatrix(54, 9) = Abs(grd.ValueMatrix(50, 9) - grd.ValueMatrix(51, 9) - grd.ValueMatrix(52, 9) - grd.ValueMatrix(53, 9))
'        '799
'        grd.TextMatrix(55, 9) = 0
'    Else
'        '798
'        grd.TextMatrix(54, 9) = 0
'        '799
'        grd.TextMatrix(55, 9) = Abs(grd.ValueMatrix(50, 9) - grd.ValueMatrix(51, 9) - grd.ValueMatrix(52, 9) - grd.ValueMatrix(53, 9))
'
'    End If
'    'retencion 100% 851-861
'    For i = 57 To 62
'        If Len(grd.TextMatrix(i, 7)) > 0 Then
'            grd.TextMatrix(i, 9) = grd.ValueMatrix(i, 7)
'        Else
'            grd.TextMatrix(i, 9) = ""
'        End If
'    Next i
'    '863    =813* 0.7
'    If Len(grd.TextMatrix(63, 7)) > 0 Then
'        grd.TextMatrix(63, 9) = Round(grd.ValueMatrix(63, 7) * 0.7, 2)
'    Else
'        grd.TextMatrix(63, 9) = ""
'    End If
'    '865    =865
'    If Len(grd.TextMatrix(64, 7)) > 0 Then
'        grd.TextMatrix(64, 9) = Round(grd.ValueMatrix(64, 7) * 0.7, 2)
'    Else
'        grd.TextMatrix(64, 9) = ""
'    End If
'
'    For i = 65 To 67
'        If Len(grd.TextMatrix(i, 7)) > 0 Then
'            grd.TextMatrix(i, 9) = Round(grd.ValueMatrix(i, 7) * 0.3, 2)
'        Else
'            grd.TextMatrix(i, 9) = ""
'        End If
'    Next i
'
''    '898=851+853+855+857+859+861+863+865+867+869
'    For i = 57 To 67
'        If Len(grd.TextMatrix(i, 9)) > 0 Then
'            SubtotalRetenido = SubtotalRetenido + grd.ValueMatrix(i, 9)
'        End If
'    Next i
'    grd.TextMatrix(68, 9) = SubtotalRetenido
'
'    '899=799+898
'    grd.TextMatrix(69, 9) = grd.ValueMatrix(55, 9) + grd.ValueMatrix(68, 9)
'
'    '902=899-901
'    grd.TextMatrix(73, 9) = grd.ValueMatrix(69, 9) - grd.ValueMatrix(72, 9)
'
'    '999=902+903+904
'    grd.TextMatrix(76, 9) = grd.ValueMatrix(73, 9) + grd.ValueMatrix(74, 9) + grd.ValueMatrix(75, 9)
'
    
End Sub

Public Sub nuevo()
    grd.Rows = grd.FixedRows
    ConfigCols
    LlenaFormatoFormulario
'    LLENADATOS
    CambiaFondoCeldasEditables
End Sub

Private Sub LlenaFormatoFormulario()
    With grd
        .MergeCells = flexMergeSpill

        .AddItem "1" & vbTab & vbTab & "DECLARACION DE RETENCIONES EN LA FUENTE DEL IMPUESTO A LA RENTA" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "No."
        .AddItem "2" & vbTab & vbTab & vbTab & vbTab & "100" & vbTab & "IDENTIFICACION DE LA DECLARACION"
        .AddItem "3" & vbTab & vbTab & vbTab & vbTab & "DECLARACION MENSUAL"
        .AddItem "4" & vbTab & vbTab & "FORMULARIO 103" & vbTab & vbTab & "101" & vbTab & "MES " & vbTab & vbTab & vbTab & "102" & vbTab & "AÑO" & vbTab
        .AddItem "5" & vbTab & "" & vbTab & "" & vbTab & vbTab & vbTab & vbTab & "104" & vbTab & "No. FORMULARIO QUE SE RETIFICA"
        .AddItem "6" & vbTab & "200" & vbTab & "IDENTIFICACION DEL SUJETO PASIVO (AJENTE DE PERCEPCION O RETENCION)"
        .AddItem "7" & vbTab & vbTab & "RUC" & vbTab & vbTab & "RAZON SOCIAL DENOMINACION O APELLIDOS Y NOMBRES COMPLETOS"
        .AddItem "8" & vbTab & "201" & vbTab & vbTab & "202" & vbTab
        .AddItem "9" & vbTab & "300" & vbTab & "POR PAGOS EN EL PAIS" & vbTab & vbTab & vbTab & vbTab & "BASE IMPONIBLE " & vbTab & vbTab & "% RET" & vbTab & "IMPUESTO RETENIDO"
        .AddItem "10" & vbTab & "ORIGINADOS" & vbTab & "EN RELACIÓN DE DEPENDENCIA QUE NO SUPERA LA BASE DESGRAVADA" & vbTab & vbTab & vbTab & vbTab & "301"
        .AddItem "11" & vbTab & "" & vbTab & "EN RELACIÓN DE DEPENDENCIA QUE SUPERA LA BASE DESGRAVADA" & vbTab & vbTab & vbTab & vbTab & "302" & vbTab & vbTab & vbTab & "352"
        .AddItem "12" & vbTab & "EN EL" & vbTab & "HONORARIOS, COMISIONES Y DIETAS A PERSONAS NATURALES" & vbTab & vbTab & vbTab & vbTab & "303" & vbTab & vbTab & "8%" & vbTab & "353"
        .AddItem "13" & vbTab & "" & vbTab & "REMUNERACION A OTROS TRABAJADORES AUTONOMOS" & vbTab & vbTab & vbTab & vbTab & "304" & vbTab & vbTab & "1%" & vbTab & "354"
        .AddItem "14" & vbTab & "TRABAJO" & vbTab & "HONORARIOS A EXTRANJEROS POR SERVICIOS OCASIONALES" & vbTab & vbTab & vbTab & vbTab & "305" & vbTab & vbTab & "25%" & vbTab & "355"
        .AddItem "15" & vbTab & "POR COMPRAS LOCALES DE MATERIA PRIMA" & vbTab & vbTab & vbTab & vbTab & vbTab & "306" & vbTab & vbTab & "1%" & vbTab & "356"
        .AddItem "16" & vbTab & "POR COMPRAS LOCALES DE BIENES NO PRODUCIDOS POR LA SOCIEDAD" & vbTab & vbTab & vbTab & vbTab & vbTab & "307" & vbTab & vbTab & "1%" & vbTab & "357"
        .AddItem "17" & vbTab & "POR COMPRAS LOCALES DE MATERIA PRIMA NO SUJETA A RETENCIÓN" & vbTab & vbTab & vbTab & vbTab & vbTab & "308"
        .AddItem "18" & vbTab & "POR SUMINISTROS Y MATERIALES" & vbTab & vbTab & vbTab & vbTab & vbTab & "309" & vbTab & vbTab & "1%" & vbTab & "359"
        .AddItem "19" & vbTab & "POR REPUESTOS Y HERRAMIENTAS" & vbTab & vbTab & vbTab & vbTab & vbTab & "310" & vbTab & vbTab & "1%" & vbTab & "360"
        .AddItem "20" & vbTab & "POR LUBRICANTES" & vbTab & vbTab & vbTab & vbTab & vbTab & "311" & vbTab & vbTab & "1%" & vbTab & "361"
        .AddItem "21" & vbTab & "POR ACTIVOS FIJOS" & vbTab & vbTab & vbTab & vbTab & vbTab & "312" & vbTab & vbTab & "1%" & vbTab & "362"
        .AddItem "22" & vbTab & "POR CONCEPTO DE SERVICIO DE TRANSPORTE PRIVADO DE PASAJEROS O SERVICIO PUBLICO O PRIVADO DE CARGA" & vbTab & vbTab & vbTab & vbTab & vbTab & "313" & vbTab & vbTab & "1%" & vbTab & "363"
        .AddItem "23" & vbTab & "POR REGALIAS, DERECHOS DE AUTOR, MARCAS, PATENTES Y SIMILARES" & vbTab & vbTab & vbTab & vbTab & vbTab & "314" & vbTab & vbTab & "8%" & vbTab & "364"
        .AddItem "24" & vbTab & "POR REMUNERACIONES A DEPORTISTAS, ENTRENADORES, CUERPO TECNICO, ARBITROS Y ARTISTAS RESIDENTES" & vbTab & vbTab & vbTab & vbTab & vbTab & "315" & vbTab & vbTab & "5%" & vbTab & "365"
        .AddItem "25" & vbTab & "POR PAGOS REALIZADOS A NOTARIOS Y REGISTRADORES DE LA PROPIEDAD O MERCANTILES" & vbTab & vbTab & vbTab & vbTab & vbTab & "316" & vbTab & vbTab & "8%" & vbTab & "366"
        .AddItem "26" & vbTab & "POR COMISIONES PAGADAS A SOCIEDADES " & vbTab & vbTab & vbTab & vbTab & vbTab & "317" & vbTab & vbTab & "5%" & vbTab & "367"
        .AddItem "27" & vbTab & "POR PROMOCION Y PUBLICIDAD" & vbTab & vbTab & vbTab & vbTab & vbTab & "318" & vbTab & vbTab & "8%" & vbTab & "368"
        .AddItem "28" & vbTab & "POR ARRENDAMIENTO MERCANTIL LOCAL" & vbTab & vbTab & vbTab & vbTab & vbTab & "319" & vbTab & vbTab & "1%" & vbTab & "369"
        .AddItem "29" & vbTab & "POR ARRENDAMIENTO DE BIENES INMUEBLES DE PROPIEDAD DE PERSONAS NATURALES" & vbTab & vbTab & vbTab & vbTab & vbTab & "320" & vbTab & vbTab & "8%" & vbTab & "370"
        .AddItem "30" & vbTab & "POR ARRENDAMIENTO DE BIENES INMUEBLES A SOCIEDADES" & vbTab & vbTab & vbTab & vbTab & vbTab & "321" & vbTab & vbTab & "5%" & vbTab & "371"
        .AddItem "31" & vbTab & "POR SEGUROS Y REASEGUROS (10% del valor de las primas facturadas)" & vbTab & vbTab & vbTab & vbTab & vbTab & "322" & vbTab & vbTab & "1%" & vbTab & "372"
        .AddItem "32" & vbTab & "POR RENDIMIENTOS FINANCIEROS" & vbTab & vbTab & vbTab & vbTab & vbTab & "323" & vbTab & vbTab & "5%" & vbTab & "373"
        .AddItem "33" & vbTab & "POR PAGOS O CREDITOS EN CUENTA REALIZADOS POR EMPRESAS EMISORAS DE TARJETAS DE CREDITO" & vbTab & vbTab & vbTab & vbTab & vbTab & "324" & vbTab & vbTab & "1%" & vbTab & "374"
        .AddItem "34" & vbTab & "POR LOTERIAS, RIFAS, APUESTAS Y SIMILARES" & vbTab & vbTab & vbTab & vbTab & vbTab & "325" & vbTab & vbTab & "15%" & vbTab & "375"
        .AddItem "35" & vbTab & "POR INTERESES Y COMISIONES EN OPERACIONES DE CREDITO ENTRE LAS INST. DEL SISTEMA FINANCIERO" & vbTab & vbTab & vbTab & vbTab & vbTab & "326" & vbTab & vbTab & "1%" & vbTab & "376"
        .AddItem "36" & vbTab & "POR VENTA DE" & vbTab & "A COMERCIALIZADORAS " & vbTab & vbTab & vbTab & vbTab & "327" & vbTab & vbTab & "2/mil" & vbTab & "377"
        .AddItem "37" & vbTab & "COMBUSTIBLES" & vbTab & "A DISTRIBUIDORES" & vbTab & vbTab & vbTab & vbTab & "328" & vbTab & vbTab & "3/mil" & vbTab & "378"
        .AddItem "38" & vbTab & "POR OTROS SERVICIOS" & vbTab & vbTab & vbTab & vbTab & vbTab & "329" & vbTab & vbTab & "1%" & vbTab & "379"
        .AddItem "39" & vbTab & "POR PAGOS DE DIVIDENDOS ANTICIPADOS" & vbTab & vbTab & vbTab & vbTab & vbTab & "330" & vbTab & vbTab & "25%" & vbTab & "380"
        .AddItem "40" & vbTab & "POR AGUA, ENERGÍA, LUZ Y TELECOMUNICACIONES" & vbTab & vbTab & vbTab & vbTab & vbTab & "331" & vbTab & vbTab & "1%" & vbTab & "381"
        .AddItem "41" & vbTab & "OTRAS COMPRAS DE BIENES Y SERVICIOS NO SUJETAS A RETENCIÓN" & vbTab & vbTab & vbTab & vbTab & vbTab & "332"
        .AddItem "42" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "SUBTOTAL          SUMAR 352 AL 381" & vbTab & vbTab & "399"
        .AddItem "43" & vbTab & "No. de empleados en nómina" & vbTab & vbTab & "105"
        .AddItem "44" & vbTab & "No. de empleados tercerizados" & vbTab & vbTab & "106"
        .AddItem "45" & vbTab & "No. de empleados bajo contrato" & vbTab & vbTab & "107" & vbTab & vbTab & "No. de comprobantes de retención emitidos" & vbTab & "108"
        .AddItem "46" & vbTab & "300" & vbTab & " 400 POR PAGOS AL EXTERIOR" & vbTab & vbTab & vbTab & vbTab & "BASE IMPONIBLE " & vbTab & vbTab & "% RET" & vbTab & "IMPUESTO RETENIDO"
        .AddItem "47" & vbTab & "CON CONVENIO DE DOBLE TRIBUTACION" & vbTab & vbTab & vbTab & vbTab & vbTab & "401" & vbTab & vbTab & vbTab & "451"
        .AddItem "48" & vbTab & "" & vbTab & "INTERESES Y COSTOS FINANCIDEROS POR FINANCIAMIENTO DE PROVEEDORES EXTERNOS (EN LA CUANTIA QUE EXCEDE A LA TASA MÁXIMA)" & vbTab & vbTab & vbTab & vbTab & "403" & vbTab & vbTab & "25%" & vbTab & "453"
        .AddItem "49" & vbTab & "SIN" & vbTab & "INTERESES DE CRÉDITOS EXTERNOS REGISTRADOS EN EL BCE (EN LA CUANTIA QUE EXCEDE A LA TASA MÁXIMA)" & vbTab & vbTab & vbTab & vbTab & "405" & vbTab & vbTab & "25%" & vbTab & "455"
        .AddItem "50" & vbTab & "" & vbTab & "INTERESES DE CRÉDITOS EXTERNOS NO REGISTRADOS EN EL BCE" & vbTab & vbTab & vbTab & vbTab & "407" & vbTab & vbTab & "25%" & vbTab & "457"
        .AddItem "51" & vbTab & "CONVENIO" & vbTab & "COMISIONES POR EXPORTACIONES" & vbTab & vbTab & vbTab & vbTab & "409" & vbTab & vbTab & "25%" & vbTab & "459"
        .AddItem "52" & vbTab & "" & vbTab & "COMISIONES PAGADAS PARA LA PROMOCION DEL TURISMO RECEPTIVO" & vbTab & vbTab & vbTab & vbTab & "411" & vbTab & vbTab & "25%" & vbTab & "461"
        .AddItem "53" & vbTab & "DOBLE" & vbTab & "EL 4% DE LAS PRIMAS DE CESIÓN O REASEGUROS CONTRATADOS CON EMPRESAS QUE NO TENGAN ESTABLECIMIENTO O REPRESENTACIÓN PERMANENTE EN EL ECUADOR" & vbTab & vbTab & vbTab & vbTab & "413" & vbTab & vbTab & "25%" & vbTab & "463"
        .AddItem "54" & vbTab & "" & vbTab & "EL 10% DE LOS PAGOS EFECTUADOS POR LAS AGENCIAS INTERNACIONALES DE PRENSA REGISTRADAS EN LA SECRETARÍA DE COMUNICACIÓN DEL ESTADO" & vbTab & vbTab & vbTab & vbTab & "415" & vbTab & vbTab & "25%" & vbTab & "465"
        .AddItem "55" & vbTab & "TRIBUTACION" & vbTab & "EL 10% DEL VALOR DE LOS CONTRATOS DE FLETAMENTO DE NAVES PARA EMPRESAS DE TRANSPORTE AÉREO O MARÍTIMO INTERNACIONAL" & vbTab & vbTab & vbTab & vbTab & "417" & vbTab & vbTab & "25%" & vbTab & "467"
        .AddItem "56" & vbTab & "" & vbTab & "EL 15% DE LOS PAGOS EFECTUADOS POR PRODUCTORAS Y DISTRIBUIDORAS DE CINTAS CINEMATOGRÁFICAS Y DE TELEVISIÓN POR CONCEPTO DE ARRENDAMIENTO DE CINTAS Y VIDEOCINTAS" & vbTab & vbTab & vbTab & vbTab & "419" & vbTab & vbTab & "25%" & vbTab & "469"
        .AddItem "57" & vbTab & "ARRENDAMIENTO" & vbTab & "POR PAGO DE INTERESES (cuando supera la tasa autorizada por el BCE)" & vbTab & vbTab & vbTab & vbTab & "421" & vbTab & vbTab & "25%" & vbTab & "471"
        .AddItem "58" & vbTab & "MERCANTTIL INT." & vbTab & "CUANDO NO SE EJERCE LA OPCION DE COMPRA (sobre la depreciación acumulada)" & vbTab & vbTab & vbTab & vbTab & "423" & vbTab & vbTab & "25%" & vbTab & "473"
        .AddItem "59" & vbTab & "Declaro que los datos contenidos en esta declaración son verdaderos por lo que asumo la responsabilidad " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "SUBTOTAL SUMAR 451 al 473" & vbTab & vbTab & "498"
        .AddItem "60" & vbTab & "correspondiente (Artículo 101 de la de la Codificación 2004-026 de la L.R.T.I.) " & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "TOTAL RETENCIONES 399 + 498" & vbTab & vbTab & "499"
        .AddItem "61" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "900 VALORES A PAGAR Y FORMA DE PAGO"
        .AddItem "62" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "PAGO PREVIO" & vbTab & vbTab & "901"
        .AddItem "63" & vbTab & vbTab & "____________________________" & vbTab & vbTab & vbTab & "____________________________" & vbTab & vbTab & "TOTAL IMPUESTO A PAGAR 499 - 901" & vbTab & vbTab & "902"
        .AddItem "64" & vbTab & vbTab & "FIRMA SUJETO PASIVO" & vbTab & vbTab & vbTab & "FIRMA CONTADOR" & vbTab & vbTab & "INTERESES POR MORA" & vbTab & vbTab & "903"
        .AddItem "65" & vbTab & "NOMBRE:" & vbTab & vbTab & "NOMBRE:" & vbTab & vbTab & vbTab & vbTab & "MULTAS" & vbTab & vbTab & "904"
        .AddItem "66" & vbTab & "198" & vbTab & "C.I. No." & vbTab & "199" & vbTab & "RUC No." & vbTab & vbTab & vbTab & "TOTAL PAGADO     902+903+904" & vbTab & vbTab & "999"
        .AddItem "67" & vbTab & "MEDIANTE CHEQUE DEBITO BANCARIO EFECTIVO U OTRAS FORMAS DE COBRO" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "905"
        .AddItem "68" & vbTab & "MEDIANTE NOTAS DE CREDITO" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "907"
        .AddItem "69" & vbTab & vbTab & "DETALLE DE NOTAS DE CREDITO"
        .AddItem "70" & vbTab & "908" & vbTab & "N/C No." & vbTab & "909" & vbTab & "USD"
        .AddItem "71" & vbTab & "910" & vbTab & "N/C No." & vbTab & "911" & vbTab & "USD"
        .AddItem "72" & vbTab & "912" & vbTab & "N/C No." & vbTab & "913" & vbTab & "USD"
        .AddItem "73" & vbTab & "914" & vbTab & "N/C No." & vbTab & "915" & vbTab & "USD"

        .Redraw = flexRDBuffered
        .Refresh
    End With
'    CambiaFondoCeldasEditables
End Sub

