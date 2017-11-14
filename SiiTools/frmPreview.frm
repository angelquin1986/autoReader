VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "Vsprint7.ocx"
Object = "{1B04A20A-C295-476C-BA28-DC6D9110E7A3}#1.0#0"; "Vspdf.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmPreview 
   Caption         =   "Area de Impresion"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   14670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   14670
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picBoton 
      Align           =   2  'Align Bottom
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   14670
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6675
      Width           =   14670
      Begin VB.CommandButton cmdExportar 
         Appearance      =   0  'Flat
         Caption         =   "Exportar PDF"
         Height          =   320
         Left            =   7680
         TabIndex        =   14
         Top             =   90
         Width           =   1095
      End
      Begin VB.CommandButton cmdPagina 
         Caption         =   "l<"
         Height          =   300
         Index           =   0
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   60
         Visible         =   0   'False
         Width           =   350
      End
      Begin VB.CommandButton cmdPagina 
         Caption         =   "<"
         Height          =   300
         Index           =   1
         Left            =   1575
         TabIndex        =   10
         Top             =   60
         Visible         =   0   'False
         Width           =   350
      End
      Begin VB.CommandButton cmdPagina 
         Caption         =   ">"
         Height          =   300
         Index           =   2
         Left            =   2535
         TabIndex        =   9
         Top             =   60
         Visible         =   0   'False
         Width           =   350
      End
      Begin VB.CommandButton cmdPagina 
         Caption         =   ">l"
         Height          =   300
         Index           =   3
         Left            =   2895
         TabIndex        =   8
         Top             =   60
         Visible         =   0   'False
         Width           =   350
      End
      Begin VB.TextBox txtPagina 
         Height          =   300
         Left            =   1935
         TabIndex        =   7
         Top             =   60
         Visible         =   0   'False
         Width           =   612
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   285
         Left            =   15
         Picture         =   "frmPreview.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.CommandButton cmdCerrar 
         Cancel          =   -1  'True
         Caption         =   "&Cerrar"
         Height          =   320
         Left            =   6600
         TabIndex        =   3
         Top             =   90
         Width           =   972
      End
      Begin VB.CommandButton cmdCambiarPosicion 
         Caption         =   "&Cambiar ahora"
         Enabled         =   0   'False
         Height          =   320
         Left            =   4800
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   1452
      End
      Begin VB.TextBox txtNumEtiqueta 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   3690
         TabIndex        =   1
         Top             =   90
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.ComboBox cboZoom 
         Height          =   315
         ItemData        =   "frmPreview.frx":0102
         Left            =   3375
         List            =   "frmPreview.frx":0118
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   60
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VSPDFLibCtl.VSPDF pdf 
         Left            =   9360
         OleObjectBlob   =   "frmPreview.frx":0139
         Top             =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pos. &inicial de la etiqueta  "
         Height          =   195
         Left            =   1755
         TabIndex        =   0
         Top             =   165
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Página  "
         Height          =   195
         Left            =   1920
         TabIndex        =   13
         Top             =   105
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Zoom  "
         Height          =   195
         Left            =   3510
         TabIndex        =   12
         Top             =   165
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   7335
      Top             =   3060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSPrinter7LibCtl.VSPrinter vp 
      Align           =   1  'Align Top
      Height          =   3630
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   14670
      _cx             =   25876
      _cy             =   6403
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _ConvInfo       =   1
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Imprimiendo..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   16.7707404103479
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   0
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   4
      NavBarColor     =   -2147483638
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   2
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private numpag As Long
Private TransNum As String
Private grdAux As VSFlexGrid  'Sirve para cuando se cambie la posición inicial
Private NumEtiq As Integer
Private mMargen As Variant    '17/05/2001 Angel.
Private mPosicion() As Variant
Private Sucursal As String
Private mEnabledExport As Long   'Cada bits desde la derecha corresponde a HTML, Texto, PDF

Public Enum EExportFormat
    spefNONE = 0
    spefHTML = 1
    spefTEXTO = 2
    spefPDF = 4
End Enum
Event ExportarTexto(ByVal archi As String)

Public Sub ShowListado(ByVal grd As VSFlexGrid)
    Dim x1 As Integer, y1 As Integer, X2 As Integer, y2 As Integer
    Me.tag = tag
    Me.Caption = "Notificaciones"
    'sucursal = trans
    picBoton.Visible = True
    x1 = Label1.Left
    y1 = Label1.Top
    X2 = txtNumEtiqueta.Left
    y2 = txtNumEtiqueta.Top
    Set grdAux = grd
    CargaListado grdAux
End Sub

Private Sub CargaListado(ByVal grd As VSFlexGrid)
    On Error GoTo ErrTrap
    Dim i As Long
    Dim j As Long
    With vp
        '.Orientation = orPortrait
        .PaperSize = pprUser
        .PhysicalPage = True
        .PaperHeight = 2500
        .PaperWidth = 2500
        .StartDoc  'Inicializa el preview
        .FontName = "Arial"
        .MarginLeft = 1000 'Margen izquierdo de la hoja.
        .HdrFontName = "Arial"
        .HdrFontBold = True
        .Header = Space(20) & gobjMain.EmpresaActual.GNOpcion.NombreEmpresa
        .FontSize = 8 'Cambia el tamaño de la letra
        For i = 1 To grd.Rows - 2
            If Not grd.IsSubtotal(i) Then
                For j = 12 To grd.Cols - 1
                    Select Case j
                        Case 12
                            If grd.Cell(flexcpBackColor, i, j, i, j) = vbWhite And (grd.ValueMatrix(i, j - 1)) = -1 Then
                                .TextBox "Cuenca   " & grd.TextMatrix(i, 12), 1500, 1700, 10000, 10000
                                .TextBox "Primera Notificación ", 8500, 1700, 10000, 10000
                                .TextBox "Sr(a).  Srs. " & grd.TextMatrix(i, 3), 1500, 2000, 10000, 10000
                                .TextBox "Transacción: " & grd.TextMatrix(i, 4), 5000, 1700, 10000, 10000
                                .TextBox cadenaImprimir, 1500, 2500, 10000, 10000
                                If siguiente(grd, i) Then vp_NewPage
                            End If
                        Case 14
                            If grd.Cell(flexcpBackColor, i, j, i, j) = vbWhite And (grd.ValueMatrix(i, j - 1)) = -1 Then
                                .TextBox "Cuenca   " & grd.TextMatrix(i, 14), 1500, 1700, 10000, 10000
                                .TextBox "Segunda Notificación ", 8500, 1700, 10000, 10000
                                .TextBox "Sr(a).  Srs. " & grd.TextMatrix(i, 3), 1500, 2000, 10000, 10000
                                .TextBox "Transacción: " & grd.TextMatrix(i, 4), 5000, 1700, 10000, 10000
                                .TextBox cadenaImprimir, 1500, 2500, 10000, 10000
                                If siguiente(grd, i) Then vp_NewPage
                            End If
                        Case 16
                            If grd.Cell(flexcpBackColor, i, j, i, j) = vbWhite And (grd.ValueMatrix(i, j - 1)) = -1 Then
                                .TextBox "Cuenca   " & grd.TextMatrix(i, 16), 1500, 1700, 10000, 10000
                                .TextBox "Tercera Notificación ", 8500, 1700, 10000, 10000
                                .TextBox "Sr(a).  Srs. " & grd.TextMatrix(i, 3), 1500, 2000, 10000, 10000
                                .TextBox "Transacción: " & grd.TextMatrix(i, 4), 5000, 1700, 10000, 10000
                                .TextBox cadenaImprimir, 1500, 2500, 10000, 10000
                                If siguiente(grd, i) Then vp_NewPage
                            End If
                    End Select
                Next
            End If
        Next i
        .EndDoc 'Finaliza la colocación de datos dentro de la hoja
    End With
    Me.Show
    Me.ZOrder
    Exit Sub
ErrTrap:
    MsgBox Err.Description
    Exit Sub
End Sub
Private Function siguiente(grd As VSFlexGrid, ByVal ind As Long) As Boolean
Dim i As Long, j As Long
Dim band As Boolean
        For i = ind + 1 To grd.Rows - 2
            If Not grd.IsSubtotal(i) Then
                For j = 12 To grd.Cols - 1
                    Select Case j
                        Case 12, 14, 16
                            If grd.Cell(flexcpBackColor, i, j, i, j) = vbWhite And (grd.ValueMatrix(i, j - 1)) = -1 Then
                            band = True
                            End If
                    End Select
                Next
            End If
        Next i
    siguiente = band
End Function


Private Sub cmdExportar_Click()
    ExportarSub
End Sub

Private Sub vp_NewPage()   'Para poner los encabezados de columnas cada vez que se cambie de página
    Dim esp As String, EncCol As String
    Dim i As Integer
    With vp
        .PaperSize = pprUser
        .PhysicalPage = True
        .PaperHeight = 2500
        .PaperWidth = 2500
        .NewPage
        .FontName = "Arial"
        .MarginLeft = 500 'Margen izquierdo de la hoja.
        .HdrFontName = "Arial"
        .HdrFontBold = True
        .Header = Space(20) & gobjMain.EmpresaActual.GNOpcion.NombreEmpresa
        .FontSize = 8 'Cambia el tamaño de la letra
        
    End With
End Sub


Private Sub cboZoom_Click()
    vp.Zoom = Val(cboZoom.Text)
End Sub




Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdImprimir_Click()
    vp.PrintDoc 'Imprime el contenido del objeto vsprinter
End Sub

Private Sub cmdPagina_Click(Index As Integer)
    With vp
        Select Case Index
            Case 0
                .PreviewPage = 1
            Case 1
                .PreviewPage = .PreviewPage - 1
            Case 2
                .PreviewPage = .PreviewPage + 1
            Case 3
                .PreviewPage = .PageCount
        End Select
    End With
    Sincronizador
End Sub

Public Sub Sincronizador()
    Dim pag As Integer, numpag As Integer
    With vp
        pag = .PreviewPage
        numpag = .PageCount
        
        cmdPagina(0).Enabled = (pag > 1)
        cmdPagina(1).Enabled = (pag > 1)
        cmdPagina(2).Enabled = (pag < numpag)
        cmdPagina(3).Enabled = (pag < numpag)
    End With
    txtPagina.Text = vp.PreviewPage
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF3
        cmdImprimir_Click
        KeyCode = 0
    Case Else
        MoverCampo Me, KeyCode, Shift, True
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ImpideSonidoEnter Me, KeyAscii
End Sub


Private Sub Form_Load()
    txtPagina.Text = vp.PreviewPage
    cboZoom.ListIndex = 1
    vp.Zoom = Val(cboZoom.Text)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState <> vbMinimized Then
        vp.Move 0, vp.Top, Me.ScaleWidth, Me.ScaleHeight - picBoton.Height - vp.Top
    End If
End Sub


Private Sub txtNumEtiqueta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdCambiarPosicion.SetFocus
    Else
        If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or (KeyAscii = vbKeyBack)) Then KeyAscii = 0
    End If
End Sub

Private Sub txtPagina_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    If Val(txtPagina.Text) >= 1 And Val(txtPagina.Text) <= vp.PageCount Then
        vp.PreviewPage = Val(txtPagina.Text)
        Sincronizador
    Else
        vp.PreviewPage = vp.PageCount
        Sincronizador
    End If
   
End Sub

Private Sub txtPagina_LostFocus()
    txtPagina_KeyPress (13)
End Sub

Private Function cadenaImprimir() As String
 Dim f As Integer, s As String, i As Integer
    Dim Cadena, s1 As String, s2 As String
    On Error GoTo ErrTrap
    Dim archi As String
    ReDim rec(0, 1)
    Dim cad As String, bod As String
    archi = GetSetting(AppName, App.Title, "Ruta PlantillaNoti", "")
    f = FreeFile                'Obtiene número disponible de archivo
    
    'Abre el archivo para lectura
    Open archi For Input As #f
        Do Until EOF(f)
        
            Line Input #f, s
            If s = "" Then
                cad = cad & Chr(13)
            Else
                cad = cad & Replace(s, ",", Chr(13))
            End If

            
        Loop
    Close #f
    cadenaImprimir = cad
    MensajeStatus
    Exit Function
ErrTrap:
'    grd.Redraw = flexRDDirect
    MensajeStatus
    DispErr
    Close       'Cierra todo
'    grd.SetFocus
    Exit Function
End Function
Private Sub ExportarPDF(ByVal archi As String)
    On Error GoTo ErrTrap
    
    'Convierte en un archivo de formato PDF
    pdf.ConvertDocument vp, archi
    
    MsgBox "El contenido fue exportado a " & vbCr & vbCr & archi, vbInformation
    Exit Sub
ErrTrap:
    MsgBox Err.Description
    Exit Sub
End Sub

Private Sub ExportarSub()
    Dim ext As String, archi As String
    archi = ShowSave
    If Len(archi) = 0 Then Exit Sub
    ext = CogerFilterIndex
    Select Case ext
    Case "htm"
 '       ExportarHTML archi
    Case "txt"
'        ExportarTexto archi
    Case "pdf"
        ExportarPDF archi
    End Select
End Sub
Private Function ShowSave() As String
    On Error GoTo ErrTrap
    
    'Si no está permitido ningún formato, sale
    If mEnabledExport = spefNONE Then Exit Function
    
    With dlg1
        .CancelError = True
        
        'Por primera vez, tipo HTML es predeterminado
        If Len(.filename) = 0 Then
            .DefaultExt = "pdf"
            .filename = App.Path
            If Right$(.filename, 1) <> "\" Then .filename = .filename & "\"
            .filename = .filename & VerificarParaFilename(Left(Me.Caption, 20))
            If Right$(.filename, 1) = "\" Then .filename = .filename & "a"
            
        'Desde la segunda vez toma el tipo anterior como predeterminado
        Else
            .DefaultExt = CogerExtension(.filename)
            If Len(.DefaultExt) = 0 Then .DefaultExt = "pdf"
            .filename = QuitarExtension(.filename)
        End If
        .InitDir = .filename
        
        .Filter = ""
        'If mEnabledExport And spefHTML Then .Filter = .Filter & "Documento HTML (*.htm)|*.htm|"
        'If mEnabledExport And spefTEXTO Then .Filter = .Filter & "Archivo Texto (*.txt)|*.txt|"
        If mEnabledExport And spefPDF Then .Filter = .Filter & "Archivo PDF (*.pdf)|*.pdf|"
        If Right$(.Filter, 1) = "|" Then .Filter = Left$(.Filter, Len(.Filter) - 1)
        
        UbicarFilterIndex .DefaultExt
        .flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist + cdlOFNHideReadOnly
        
        'Abre el cuadro de dialogo
        .ShowSave
        
        'Si el nombre de destino no tiene extensión
        If Len(CogerExtension(.filename)) = 0 Then
            'Asegura que tenga la misma extensión que está seleccionada en Tipo de Doc. de DialogBox
            .filename = .filename & "." & CogerFilterIndex
        End If
        ShowSave = .filename
    End With
    Exit Function
ErrTrap:
    If Err.Number <> 32755 Then
        MsgBox Err.Description
    End If
    Exit Function
End Function
'Asegura que la extensión sea la misma que está seleccionada en dialog
Private Function CogerFilterIndex() As String
    Dim v  As Variant, n As Long, ext As String, i As Long
    
    v = Split(dlg1.Filter, "|")
    If Not IsEmpty(v) Then
        n = (dlg1.FilterIndex - 1) * 2 + 1
        If n <= UBound(v, 1) Then ext = v(n)
    End If
    
    'Quitar '*.'
    i = InStrRev(ext, ".")
    If i > 0 Then ext = Right$(ext, Len(ext) - i)
    
    CogerFilterIndex = ext
End Function


Private Function VerificarParaFilename(ByRef fname As String) As String
    Dim s As String
    
    'Reemplaza caracteres no adecuados para nombre de archivo
    s = Replace(fname, "/", "-")
    s = Replace(fname, ":", "-")
    
    VerificarParaFilename = s
End Function
Private Function CogerExtension(ByVal archi As String) As String
    Dim i As Long
    
    i = InStrRev(archi, ".")
    If (i > 0) And (Len(archi) - i <= 3) Then
        CogerExtension = LCase$(Right$(archi, Len(archi) - i))
    End If
End Function
Private Function QuitarExtension(ByVal archi As String) As String
    Dim i As Long
    
    i = InStrRev(archi, ".")
    If (i > 0) And (Len(archi) - i <= 3) Then
        QuitarExtension = Left$(archi, i - 1)
    Else
        QuitarExtension = archi
    End If
End Function

Private Sub UbicarFilterIndex(ByVal ext As String)
    Dim v  As Variant, i As Long, n As Long
    n = 0
    v = Split(dlg1.Filter, "|")
    If Not IsEmpty(v) Then
        For i = LBound(v, 1) To UBound(v, 1)
            If v(i) = "*." & ext Then
                n = Int(i / 2) + 1
                Exit For
            End If
        Next i
    End If
    If n > 0 Then dlg1.FilterIndex = n
End Sub

Private Sub Form_Initialize()
    'mCopias = 1
    mEnabledExport = 255
End Sub

Public Property Get EnabledExport() As Long
    EnabledExport = mEnabledExport
End Property
