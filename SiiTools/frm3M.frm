VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{ED5A9B02-5BDB-48C7-BAB1-642DCC8C9E4D}#2.0#0"; "SelFold.ocx"
Begin VB.Form frm3M 
   Caption         =   "dlg1"
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
   Begin VB.Frame frmfecha 
      Height          =   1455
      Left            =   60
      TabIndex        =   7
      Top             =   0
      Width           =   8895
      Begin VB.CommandButton cmdPasos 
         Caption         =   "Abrir Archivo Pipe"
         Height          =   330
         Index           =   4
         Left            =   4140
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1020
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "Abrir Archivo Excel"
         Height          =   330
         Index           =   3
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1020
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "Buscar"
         Height          =   330
         Index           =   2
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   660
         Width           =   1455
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "Generar Archivo"
         Height          =   330
         Index           =   10
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1455
      End
      Begin VB.TextBox txtCarpeta 
         Height          =   320
         Left            =   1080
         TabIndex        =   2
         Text            =   "c:\"
         Top             =   660
         Visible         =   0   'False
         Width           =   4170
      End
      Begin VB.CommandButton cmdExaminarCarpeta 
         Caption         =   "..."
         Height          =   320
         Index           =   0
         Left            =   5280
         TabIndex        =   3
         Top             =   660
         Visible         =   0   'False
         Width           =   372
      End
      Begin SelFold.SelFolder slf 
         Left            =   4860
         Top             =   480
         _ExtentX        =   1349
         _ExtentY        =   265
         Title           =   "Seleccione una carpeta"
         Caption         =   "Selección de carpeta"
         RootFolder      =   "\"
         Path            =   "C:\VBPROG_ESP\SII\SELFOLD"
      End
      Begin MSComCtl2.DTPicker dtpPeriodo 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   300
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MMMM/yyyy"
         Format          =   106692611
         CurrentDate     =   37356
      End
      Begin MSComDlg.CommonDialog dlg1 
         Left            =   3420
         Top             =   240
         _ExtentX        =   688
         _ExtentY        =   688
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Label lblResp 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Index           =   5
         Left            =   2640
         TabIndex        =   12
         Top             =   1020
         Width           =   2925
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Corte:"
         Height          =   255
         Left            =   60
         TabIndex        =   9
         Top             =   360
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Ubicacion:"
         Height          =   255
         Left            =   60
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   870
      End
   End
   Begin VB.PictureBox picBoton 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   9825
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7365
      Width           =   9825
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   288
         Left            =   10020
         TabIndex        =   5
         Top             =   60
         Width           =   1212
      End
      Begin MSComctlLib.ProgressBar prg 
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   120
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   3870
      Left            =   60
      TabIndex        =   4
      Top             =   1500
      Width           =   15015
      _cx             =   26485
      _cy             =   6826
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
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm3M.frx":0000
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
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "frm3M"
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
Private mobjBusq As Busqueda

Private WithEvents mGrupo As grupo
Attribute mGrupo.VB_VarHelpID = -1
Const COL_V_CODIGO = 1
Const COL_V_DESC = 2
Const COL_V_RUC = 3
Const COL_V_CLIENTE = 4
Const COL_V_CANT = 5
Const COL_V_UNIDAD = 6
Const COL_V_PU = 7
Const COL_V_PT = 8
Const COL_V_MONEDA = 9
Const COL_V_FECHA = 10
Const COL_V_UBIGEO = 11
Const COL_V_NUMDOC = 12
Const COL_V_TIPODOC = 13
Const COL_V_MERCADO = 14
Const COL_V_VENDE = 15
Const COL_V_ZONA = 16
Const COL_V_PROV = 17
Const COL_V_DISTRI = 18
Const COL_V_REGALO = 19
Const COL_V_GRUPO3 = 20
Const COL_V_RESP = 21

Const COL_E_CODIGO = 1
Const COL_E_DESC = 2
Const COL_E_CANT = 3
Const COL_E_UNIDAD = 4
Const COL_E_FECHA = 5


Private Cadena As String
Private cadenaDD  As String

Private NumFile As Integer
Private NumProc As Integer
Private TotalVentas As Currency
Private mbooEjecutando As Boolean

Public Sub InicioVentas(ByVal tag As String)
    On Error GoTo ErrTrap
    Set mObjCond = New RepCondicion
    Me.Caption = "Ventas 3M"
    TotalVentas = 0
    dtpPeriodo.value = CDate("01/" & IIf(Month(Date) - 1 <> 0, Month(Date) - 1, 12) & "/" & Year(Date))
    mObjCond.fecha1 = dtpPeriodo.value

    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Ruta-DINARDAP")) > 0 Then
        txtCarpeta.Text = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Ruta-DINARDAP")
    End If
    Me.tag = tag
    Me.Show
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub





Private Sub cmdCancelar_Click()
    mbooCancelado = True
End Sub


Private Sub cmdPasos_Click(Index As Integer)
    Dim r As Boolean, cad As String, nombre As String, file As String
    NumProc = Index + 1
    
    
        cmdPasos(2).BackColor = vbButtonFace
        cmdPasos(3).BackColor = vbButtonFace
        cmdPasos(4).BackColor = vbButtonFace
        cmdPasos(10).BackColor = vbButtonFace

    
    Select Case Index + 1
    Case 3      '2. Busca Ventas
            If Me.Caption <> "Existencias 3M" Then
                BuscarVentas3M
            Else
                BuscarExist3M
            End If

            cmdPasos(2).BackColor = &HFFFF00
    
    Case 11      '8. Generar Archivo
    
            If Me.Caption <> "Existencias 3M" Then
                GeneraTxt3M
            Else
                ExportaTxt3MInventario
            End If
            Cadena = cadenaDD

            Close NumFile
            
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "Ruta-DINARDAP", txtCarpeta.Text
            gobjMain.EmpresaActual.GNOpcion.Grabar
            
            r = True
            
            
        lblResp(5).Caption = "OK."
    
    End Select
    
    
    
        If r Then

        Else
                lblResp(5).BackColor = vbBlue
                lblResp(5).ForeColor = vbYellow
        End If
End Sub

Private Sub dtpPeriodo_Change()
 Dim i As Integer

        cmdPasos(2).Enabled = True
    
        lblResp(5).BackColor = &HC0FFFF
        lblResp(5).Caption = ""

End Sub

Private Sub Form_Initialize()
'    Set mobjBusq = New Busqueda
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
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = mbooProcesando
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    grd.Move 0, frmfecha.Height + 100, Me.ScaleWidth, (Me.ScaleHeight - (frmfecha.Height + picBoton.Height) - 105)
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


Public Sub MiGetRowsRep(ByVal rs As Recordset, grd As VSFlexGrid)
    grd.LoadArray MiGetRows(rs)
End Sub


Private Sub ConfigCols(cad As String)
    Dim s As String, i As Integer
    Select Case cad

    Case "IMPFC"

        s = "^#|<Codigo|<Descripcion|<RUC|<Nombre|>Cantidad|^Unidad|>P.Unitario|>P. Total|^Moneda|^Fecha|Ubigeo|>Numero Doc.|^Tipo Doc.|<Mercado|<Vendedor|<Zona|<Provincia|<Distrito|<Regalo|<Grupo3"
        grd.FormatString = s & "|<         Resultado           "
        AsignarTituloAColKey grd
        
        grd.ColFormat(COL_V_PU) = "#,#0.00"
        grd.ColFormat(COL_V_PT) = "#,#0.00"
        grd.ColFormat(COL_V_CANT) = "#,#0"

        grd.ColData(COL_V_CANT) = "SubTotal"
        grd.ColData(COL_V_PT) = "SubTotal"

    If grd.Rows > 1 Then
         grd.Cell(flexcpBackColor, 1, COL_V_CODIGO, grd.Rows - 1, COL_V_MERCADO) = &H80000018
         grd.Cell(flexcpBackColor, 1, COL_V_VENDE, grd.Rows - 1, COL_V_VENDE) = vbWhite
         grd.Cell(flexcpBackColor, 1, COL_V_ZONA, grd.Rows - 1, COL_V_GRUPO3) = &H80000018
         Totalizar
     End If

    Case "EXIST"

        s = "^#|<Codigo|<Descripcion|>Cantidad|^Unidad|^Fecha"
        grd.FormatString = s & "|<         Resultado           "
        AsignarTituloAColKey grd
        
        grd.ColFormat(COL_E_CANT) = "#,#0"

        grd.ColData(COL_E_CANT) = "SubTotal"

    If grd.Rows > 1 Then
         grd.Cell(flexcpBackColor, 1, COL_E_CODIGO, grd.Rows - 1, COL_E_FECHA) = &H80000018
         Totalizar
     End If

    
    End Select
    

    AsignarTituloAColKey grd
    grd.SetFocus

End Sub



''''Private Sub GeneraArchivo()
''''    Dim v As Variant, file As String, nombre As String
''''    Dim Filas As Long, Columnas As Long, i As Long, j As Long
''''    On Error GoTo ErrTrap
''''    nombre = "AT" & Format(CStr(Month(mObjCond.Fecha2)), "00") & Year(mObjCond.Fecha2) & ".XML"
''''    file = "c:\" & nombre 'txtCarpeta.Text & Nombre
''''    If ExisteArchivo(file) Then
''''        If MsgBox("El nombre del archivo " & nombre & " ya existe desea sobreescribirlo?", vbYesNo) = vbNo Then
''''            Exit Sub
''''        End If
''''    End If
''''    NumFile = FreeFile
''''    Open file For Output Access Write As #NumFile
'''''     grd.AddItem vbTab & Nombre & vbTab & "Generando  archivo..."
''''    Cadena = GeneraArchivoEncabezado
''''
''''
'''''    grd.AddItem vbTab & Nombre & vbTab & "Generando  archivo..."
''''   Print #NumFile, Cadena
''''
''''    Close NumFile
'''''    grd.textmatrix(i,grd.Rows - 1, grd.Cols - 1) = "Grabado con exito"
''''    Exit Sub
''''ErrTrap:
''''    'grd.TextMatrix(i, grd.Rows - 1, 2) = Err.Description
''''    Close NumFile
''''End Sub

Private Function GeneraArchivoEncabezadoATSXML() As String
    Dim obj As GNOpcion, cad As String, numSucursal As Integer
    cad = "<?xml version=" & """1.0""" & " encoding=" & """UTF-8""" & "" & " standalone=" & """no""" & "?>"
    cad = cad & "<!--  Generado por Ishida Asociados   -->"
    cad = cad & "<!--  Dir: Av. Espana  y Elia Liut Aeropuerto Mariscal Lamar Segundo Piso -->"
    cad = cad & "<!--  Telf: 098499003, 072870346      -->"
    cad = cad & "<!--  email: ishidacue@hotmail.com    -->"
    cad = cad & "<!--  Cuenca - Ecuador                -->"
    cad = cad & "<!--  SISTEMAS DE GESTION EMPRESASRIAL-->"
        
    cad = cad & "<iva>"
        
    cad = cad & "<TipoIDInformante> R </TipoIDInformante>"
    cad = cad & "<IdInformante>" & Format(gobjMain.EmpresaActual.GNOpcion.ruc, "0000000000000") & "</IdInformante>"
    cad = cad & "<razonSocial>" & UCase(gobjMain.EmpresaActual.GNOpcion.RazonSocial) & "</razonSocial>"
    cad = cad & "<Anio>" & Year(mObjCond.fecha1) & "</Anio>"
    cad = cad & "<Mes>" & IIf(Len(Month(mObjCond.fecha1)) = 1, "0" & Month(mObjCond.fecha1), Month(mObjCond.fecha1)) & "</Mes>"
    
    numSucursal = gobjMain.EmpresaActual.RecuperaNumeroSucursales
    cad = cad & "<numEstabRuc>" & Format(numSucursal, "000") & "</numEstabRuc>"
    
'    TotalVentas = gobjMain.EmpresaActual.RecuperaNumeroSucursales
    cad = cad & "<totalVentas>" & Format(TotalVentas, "#0.00") & "</totalVentas>"
    cad = cad & "<codigoOperativo>IVA</codigoOperativo>"

'    cad = cad & "<compras>"

    GeneraArchivoEncabezadoATSXML = cad
End Function



Public Function RellenaDer(ByVal s As String, lon As Long) As String
    Dim r As String
    r = "!" & String(lon, "@")
    If Len(s) = 0 Then s = " "
    RellenaDer = Format(s, r)
End Function

Public Function ValidaTelefono(ByVal Tel As String) As String
    Dim c As String
    If Len(Tel) < 6 Then Exit Function
    'asigna caracter
    Select Case Mid(Tel, 1, 2)
            Case "02", "04", "07": c = "2"
            Case "09": c = "9"
            Case Else: c = "-"  'Diego 27 Abril 2004 ' si va jeaa 02/04/04
    End Select
   
    Select Case Len(Tel)
    Case 6: Tel = "07" & c & Tel
    Case 7:
        If InStr("0249", Mid(Tel, 1, 1)) = 0 Then
            Tel = "0" & Mid(Tel, 1, 1) & c & Mid(Tel, 2, Len(Tel))
        Else
            'jeaa 2/06/04
            Tel = "07" & Tel
        End If
    Case 8: Tel = Mid(Tel, 1, 2) & c & Mid(Tel, 3, 8)
    Case 9: If Mid(Tel, 3, 1) <> c Then Tel = Mid(Tel, 1, 2) & c & Mid(Tel, 3, Len(Tel))
    End Select
    
    ValidaTelefono = Tel
End Function



Private Sub cmdExaminarCarpeta_Click(Index As Integer)
    On Error GoTo ErrTrap
    slf.OwnerHWnd = Me.hWnd
    slf.Path = txtCarpeta.Text
    If slf.Browse Then
        txtCarpeta.Text = slf.Path
        txtCarpeta_LostFocus
    End If
    Exit Sub
ErrTrap:
    MsgBox Err.Description, vbInformation
    Exit Sub
End Sub



Private Sub txtCarpeta_LostFocus()
    If Right$(txtCarpeta.Text, 1) <> "\" Then
        txtCarpeta.Text = txtCarpeta.Text & "\"
    End If
    'Luego a actualiza linea de comando
End Sub

Private Function BuscarVentas3M()
    Dim fecha1 As Date
    Dim fecha2 As Date

    On Error GoTo ErrTrap
        With grd
        .Redraw = False
        .Rows = .FixedRows
'        If Not frmB_Trans.Inicio(gobjMain, "IMPDD", dtpPeriodo.value) Then
'            grd.SetFocus
'        End If
        
        If DatePart("m", dtpPeriodo.value) = 12 Then
            fecha1 = "01/" & DatePart("m", dtpPeriodo.value) & "/" & DatePart("yyyy", dtpPeriodo.value)
            fecha2 = DateAdd("yyyy", 1, DateAdd("d", -1, ("01/" & DatePart("m", DateAdd("m", 1, dtpPeriodo.value)) & "/" & DatePart("yyyy", dtpPeriodo.value))))
        Else
            fecha1 = "01/" & DatePart("m", dtpPeriodo.value) & "/" & DatePart("yyyy", dtpPeriodo.value)
            fecha2 = DateAdd("d", -1, ("01/" & DatePart("m", DateAdd("m", 1, dtpPeriodo.value)) & "/" & DatePart("yyyy", dtpPeriodo.value)))
        End If
'        With objSiiMain.objCondicion
'        .fecha1 = fecha1
 '       .fecha2 = fecha2

        
        gobjMain.objCondicion.fecha1 = fecha1
        gobjMain.objCondicion.fecha2 = fecha2
        MiGetRowsRep gobjMain.EmpresaActual.ConsVentas3M(), grd


        'GeneraArchivo

        ConfigCols "IMPFC"
        AjustarAutoSize grd, -1, -1
        AjustarAutoSize grd, -1, -1
        grd.ColWidth(0) = "500"


        GNPoneNumFila grd, False


        .Redraw = True
   End With

    Exit Function
ErrTrap:
    grd.Redraw = True
    DispErr
    Exit Function
End Function





'''Public Sub Exportar(tag As String)
'''    Dim file As String, NumFile As Integer, Cadena As String
'''    Dim Filas As Long, Columnas As Long, i As Long, j As Long
'''    Dim pos As Integer
''''    If grd.Rows = grd.FixedRows Then Exit Sub
'''    On Error GoTo errtrap
'''
'''        With dlg1
'''          .CancelError = True
'''          '.Filter = "Texto (Separado por coma)|*.txt|Excel 97(XLS)|*.xls"
'''          .Filter = "Texto (Separado por coma)|*.csv"
'''          .ShowSave
'''
'''          file = .filename
'''        End With
'''
'''
'''    If ExisteArchivo(file) Then
'''        If MsgBox("El nombre del archivo " & file & " ya existe desea sobreescribirlo?", vbYesNo) = vbNo Then
'''            Exit Sub
'''        End If
'''    End If
'''
'''    NumFile = FreeFile
'''
'''    Open file For Output Access Write As #NumFile
'''
'''    Cadena = ""
'''    For i = 0 To grd.Rows - 1
'''        For j = 2 To grd.Cols - 1
'''            Select Case tag          ' jeaa 04/11/03 para que se no se guarden las columnas ocultas
'''                Case "IMPCP"
'''                        If j = COL_C_NOMBRE Then j = j + 1  'columna nombre
'''            End Select
'''                If pos = 0 Then
'''                    Cadena = Cadena & grd.TextMatrix(i, j) & ","
'''                Else
'''                    Cadena = Cadena & Mid$(grd.TextMatrix(i, j), 1, pos - 1) & Mid$(grd.TextMatrix(i, j), pos + 1, Len(grd.TextMatrix(i, j)) - 1) & ","
'''                End If
'''
'''
'''        Next j
'''        Cadena = Mid(Cadena, 1, Len(Cadena) - 1)
'''        Print #NumFile, Cadena
'''        Cadena = ""
'''    Next i
'''
'''
'''    Close NumFile
'''    MsgBox "El archivo se ha exportado con éxito"
'''    Exit Sub
'''errtrap:
'''    If Err.Number <> 32755 Then
'''        MsgBox Err.Description
'''    End If
'''    Close NumFile
'''End Sub



Private Sub grd_KeyDown(KeyCode As Integer, Shift As Integer)
    If grd.IsSubtotal(grd.Row) Then Exit Sub
    Select Case KeyCode
    Case vbKeyInsert
        AgregarFila
    Case vbKeyDelete
        EliminarFila
    End Select
End Sub

Private Sub AgregarFila()
    On Error GoTo ErrTrap
    With grd
        .AddItem "", .Row + 1
        GNPoneNumFila grd, False
        .Row = .Row + 1
        .col = .FixedCols
    End With
    
    AjustarAutoSize grd, -1, -1
    grd.SetFocus
    Exit Sub
ErrTrap:
    MsgBox Err.Description
    grd.SetFocus
    Exit Sub
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


Public Sub ExportaTxtPipe(ByVal grd As Control, CargaNombreArchivo As String)

    Dim file As String, NumFile As Integer, fila
    Dim r As Long, c As Long, Separador As String
    
    
    
    NumFile = FreeFile
    file = "DINAR" 'CargaNombreArchivo
    If file = "" Then Exit Sub
    Open file For Output Access Write As #NumFile
    With grd
            For r = 1 To .Rows - 1
                If .RowHidden(r) = False Then ' Filas  Ocultas
                    fila = ""
                    For c = 1 To .Cols - 1
                        If .ColHidden(c) = False Then
                            If .IsSubtotal(r) And Len(.TextMatrix(r, c)) > 0 And .TextMatrix(r, c) <> "Subtotal" Then
                                fila = fila & .ValueMatrix(r, c) & "|"
                            Else
                                fila = fila & .TextMatrix(r, c) & "|"
                            End If
                        End If
                    Next c
                    fila = Left(fila, Len(fila) - 1)
                    Print #NumFile, fila
                End If
                .TextMatrix(r, 39) = "OK"
                
                '.Cell(flexcpBackColor, r, 1, r, COL_D_RESP) = vbWhite
            Next r

    End With
    Close NumFile
    gobjMain.EmpresaActual.GrabaGNLogAccion "EXP-TXT", "Exporta Reporte a PIPE txt " & file, "RE"
End Sub

    




Private Sub AbrirArchivoExcel()
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


        AjustarAutoSize grd, -1, -1
    End With
    Exit Sub
ErrTrap:
    If Err.Number <> 32755 Then DispErr
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





Private Sub SubTotalizar(col As Long)
    Dim i As Long
    With grd
        For i = 1 To .Cols - 1
            'If i = COL_C_CODTIPOCOMP Then i = i + 1
            If grd.ColData(i) = "SubTotal" Then
                    .SubTotal flexSTSum, col, i, , grd.GridColor, vbBlack, , "Subtotal", col, True
            End If
        Next i
        .SubTotal flexSTCount, col, col, , grd.GridColor, vbBlack, , "Subtotal", col, True

    End With
End Sub

Private Sub Totalizar()
    Dim i As Long
    With grd
        For i = 1 To .Cols - 1
            'If i = COL_C_CODTIPOCOMP Then i = i + 1
            If grd.ColData(i) = "SubTotal" Then
                
                .SubTotal flexSTSum, -1, i, "#,#0.00", .BackColorSel, vbYellow, vbBlack, "Total"
            End If
        Next i
'        .subtotal flexSTCount, -1, COL_C_CODTIPOCOMP, "#,#0", .BackColorSel, vbYellow, vbBlack, "Total"
    End With
End Sub

Public Sub GeneraTxt3M()

    Dim file As String, NumFile As Integer, fila
    Dim r As Long, c As Long, Separador As String
    
    
    
    NumFile = FreeFile
    file = CargaNombreArchivo3M(True)
    If file = "" Then Exit Sub
    'File = "\001A1 VENTA      EC" & gobjMain.EmpresaActual.GNOpcion.RUC & "_" & "_" & Right("00" & DatePart("d", Date), 2) & Right("00" & DatePart("m", Date), 2) & DatePart("yyyy", Date) & "_" & Right("00" & DatePart("h", Time), 2) & Right("00" & DatePart("n", Time), 2) & Right("00" & DatePart("s", Time), 2) & ".txt"
    Open file For Output Access Write As #NumFile
    With grd
'        If Mid(File, InStr(File, ".") + 1, 3) = "txt" Then
            fila = "001A1 VENTA      EC" & gobjMain.EmpresaActual.GNOpcion.ruc & "CUENCA" & "    PE20100119227"
            Print #NumFile, fila
            For r = 1 To .Rows - 1
                If .RowHidden(r) = False Then ' Filas  Ocultas
                    fila = ""
                    For c = 1 To .Cols - 1
                        'If .ColHidden(c) = False Then
                            If .IsSubtotal(r) And Len(.TextMatrix(r, c)) > 0 Then
                                
                            Else
                                Select Case c
                                Case 1: fila = fila & Left(.TextMatrix(r, c) + "                    ", 17)
                                Case 2: fila = fila & "" & Left(.TextMatrix(r, c) + "                    ", 40)
                                Case 3: fila = fila & "" & Left(.TextMatrix(r, c) + "                    ", 13)
                                Case 4: fila = fila & "" & Left(.TextMatrix(r, c) + "                                                                                                ", 100)
                                Case 5: fila = fila & "" & Right("                    " + Format(.TextMatrix(r, c), "#0.00"), 17)
                                Case 6: fila = fila & "" & Left(.TextMatrix(r, c) + "                    ", 5)
                                Case 7: fila = fila & "" & Right("                    " + Format(.TextMatrix(r, c), "#0.00"), 17)
                                Case 8: fila = fila & "" & Right("                    " + Format(.TextMatrix(r, c), "#0.00"), 17)
                                Case 9: fila = fila & "" & Left(.TextMatrix(r, c) + "                    ", 4)
                                Case 10: fila = fila & "" & Right("                    " + Format(.TextMatrix(r, c), "yyyy-mm-dd"), 12)
                                Case 11: fila = fila & "" & ("       ")
                                Case 12: fila = fila & "" & Right("          " + .TextMatrix(r, c), 20)
                                Case 13: fila = fila & "" & .TextMatrix(r, c)
                                Case 14: fila = fila & "" & Left(.TextMatrix(r, c) + "                                                             ", 50)
                                Case 15: fila = fila & "" & Left(.TextMatrix(r, c) + "                                                             ", 15)
                                Case 16: fila = fila & "" & Left(.TextMatrix(r, c) + "                                                             ", 20)
                                Case 17: fila = fila & "" & Left(.TextMatrix(r, c) + "                                                             ", 50)
                                Case 18: fila = fila & "" & Left(.TextMatrix(r, c) + "                                                             ", 50)
                                Case 19: fila = fila & "" & Left(.TextMatrix(r, c) + "                                                                                                                                                                  ", 100)
                                End Select
                            End If
                        'End If
                    Next c
                    'Fila = Left(Fila, Len(Fila) - 1)
                    Print #NumFile, fila
                End If
            Next r
'        Else
'            For r = 0 To .Rows - 1
'                If .RowHidden(r) = False Then ' Filas  Ocultas
'                    Fila = ""
'                    For c = 1 To .Cols - 1
'                        'If .ColHidden(c) = False Then
'                            If .IsSubtotal(r) And Len(.TextMatrix(r, c)) > 0 And .TextMatrix(r, c) <> "Subtotal" Then
'                                Fila = Fila & .ValueMatrix(r, c) & ","
'                            Else
'                                Fila = Fila & .TextMatrix(r, c) & ","
'                            End If
'                        'End If
'                    Next c
'                    Fila = Left(Fila, Len(Fila) - 1)
'                    Print #numfile, Fila
'                End If
'            Next r
'        End If
    End With
    Close NumFile
    gobjMain.EmpresaActual.GrabaGNLogAccion "EXP-TXT", "Exporta Reporte a txt " & file, "RE"
End Sub

Private Function CargaNombreArchivo3M(bandVenta As Boolean) As String
    On Error GoTo ErrTrap
    With frmMain.dlg1
        .InitDir = App.Path
        .DialogTitle = "Guardar Archivo"
        .CancelError = True
        .Filter = "Archivo de Texto|*.txt;|Texto (Separado por coma)|*.csv"
        .DefaultExt = "txt"
        

        If bandVenta Then
            .filename = txtCarpeta.Text & "001A1_VENTA_EC" & gobjMain.EmpresaActual.GNOpcion.ruc & "_CUENCA" & "_" & Right("00" & DatePart("d", Date), 2) & Right("00" & DatePart("m", Date), 2) & DatePart("yyyy", Date) & "_" & Right("00" & DatePart("h", Time), 2) & Right("00" & DatePart("n", Time), 2) & Right("00" & DatePart("s", Time), 2) & ".txt"
        Else
            .filename = txtCarpeta.Text & "001B1_INVENTARIO_EC" & gobjMain.EmpresaActual.GNOpcion.ruc & "_CUENCA" & "_" & Right("00" & DatePart("d", Date), 2) & Right("00" & DatePart("m", Date), 2) & DatePart("yyyy", Date) & "_" & Right("00" & DatePart("h", Time), 2) & Right("00" & DatePart("n", Time), 2) & Right("00" & DatePart("s", Time), 2) & ".txt"
        End If
        .flags = cdlOFNCreatePrompt Or cdlOFNOverwritePrompt Or cdlOFNHideReadOnly
        .ShowSave
        CargaNombreArchivo3M = .filename
    End With
    Exit Function
ErrTrap:
    Exit Function
End Function

Public Sub InicioExist(ByVal tag As String)
    On Error GoTo ErrTrap
    Set mObjCond = New RepCondicion
    Me.Caption = "Existencias 3M"
    
    TotalVentas = 0
    
    dtpPeriodo.Format = dtpCustom
    dtpPeriodo.CustomFormat = "dd/MM/yyyy"
    
   
    dtpPeriodo.value = CDate(DateAdd("d", -1, "01/" & IIf(Month(Date) - 1 <> 0, Month(Date), 12) & "/" & Year(Date)))
    mObjCond.fecha1 = dtpPeriodo.value

'    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Ruta-DINARDAP")) > 0 Then
 '       txtCarpeta.Text = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Ruta-DINARDAP")
  '  End If
    Me.tag = tag
    Me.Show
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Function BuscarExist3M()
    Dim fecha1 As Date
    Dim fecha2 As Date

    On Error GoTo ErrTrap
        With grd
        .Redraw = False
        .Rows = .FixedRows
'        If Not frmB_Trans.Inicio(gobjMain, "IMPDD", dtpPeriodo.value) Then
'            grd.SetFocus
'        End If
        
        If DatePart("m", dtpPeriodo.value) = 12 Then
            fecha1 = "01/" & DatePart("m", dtpPeriodo.value) & "/" & DatePart("yyyy", dtpPeriodo.value)
            fecha2 = DateAdd("yyyy", 1, DateAdd("d", -1, ("01/" & DatePart("m", DateAdd("m", 1, dtpPeriodo.value)) & "/" & DatePart("yyyy", dtpPeriodo.value))))
        Else
            fecha1 = "01/" & DatePart("m", dtpPeriodo.value) & "/" & DatePart("yyyy", dtpPeriodo.value)
            fecha2 = DateAdd("d", -1, ("01/" & DatePart("m", DateAdd("m", 1, dtpPeriodo.value)) & "/" & DatePart("yyyy", dtpPeriodo.value)))
        End If
'        With objSiiMain.objCondicion
'        .fecha1 = fecha1
'        .fecha2 = fecha2

        
        'gobjMain.objCondicion.fecha1 = fecha2 'DateAdd("d", -1, DateAdd("m", 1, fecha1))
        gobjMain.objCondicion.FechaCorte = fecha2
'        gobjMain.objCondicion.fecha2 = fecha1
        MiGetRowsRep gobjMain.EmpresaActual.ConsExist3M(), grd


        'GeneraArchivo

        ConfigCols "EXIST"
        AjustarAutoSize grd, -1, -1
        AjustarAutoSize grd, -1, -1
        grd.ColWidth(0) = "500"


        GNPoneNumFila grd, False


        .Redraw = True
   End With

    Exit Function
ErrTrap:
    grd.Redraw = True
    DispErr
    Exit Function
End Function

Public Sub ExportaTxt3MInventario()

    Dim file As String, NumFile As Integer, fila
    Dim r As Long, c As Long, Separador As String
    
    
    
    NumFile = FreeFile
    file = CargaNombreArchivo3M(False)
    If file = "" Then Exit Sub
    'File = "\001A1 VENTA      EC" & gobjMain.EmpresaActual.GNOpcion.RUC & "_" & "_" & Right("00" & DatePart("d", Date), 2) & Right("00" & DatePart("m", Date), 2) & DatePart("yyyy", Date) & "_" & Right("00" & DatePart("h", Time), 2) & Right("00" & DatePart("n", Time), 2) & Right("00" & DatePart("s", Time), 2) & ".txt"
    Open file For Output Access Write As #NumFile
    With grd
        If Mid(file, InStr(file, ".") + 1, 3) = "txt" Then
            fila = "001B1 INVENTARIO EC" & gobjMain.EmpresaActual.GNOpcion.ruc & "CUENCA" & "    PE20100119227"
            Print #NumFile, fila
            For r = 1 To .Rows - 1
                If .RowHidden(r) = False Then ' Filas  Ocultas
                    fila = ""
                    For c = 1 To .Cols - 1
                        'If .ColHidden(c) = False Then
                            If .IsSubtotal(r) And Len(.TextMatrix(r, c)) > 0 Then
                                
                            Else
                                Select Case c
                                Case 1: fila = fila & Left(.TextMatrix(r, c) + "                                                                   ", 17)
                                Case 2: fila = fila & "" & Left(.TextMatrix(r, c) + "                                                           ", 40)
                                Case 3: fila = fila & "" & Right("                                                            " + Format(.TextMatrix(r, c), "#0.00"), 17)
                                Case 4: fila = fila & "" & Left(.TextMatrix(r, c) + "                    ", 5)
                                Case 5: fila = fila & "" & Right("                                                " + Format(.TextMatrix(r, c), "yyyy-mm-dd"), 12)
                                End Select
                            End If
                        'End If
                    Next c
                    'Fila = Left(Fila, Len(Fila) - 1)
                    Print #NumFile, fila
                End If
            Next r
        Else
            For r = 0 To .Rows - 1
                If .RowHidden(r) = False Then ' Filas  Ocultas
                    fila = ""
                    For c = 1 To .Cols - 1
                        'If .ColHidden(c) = False Then
                            If .IsSubtotal(r) And Len(.TextMatrix(r, c)) > 0 And .TextMatrix(r, c) <> "Subtotal" Then
                                fila = fila & .ValueMatrix(r, c) & ","
                            Else
                                fila = fila & .TextMatrix(r, c) & ","
                            End If
                        'End If
                    Next c
                    fila = Left(fila, Len(fila) - 1)
                    Print #NumFile, fila
                End If
            Next r
        End If
    End With
    Close NumFile
    gobjMain.EmpresaActual.GrabaGNLogAccion "EXP-TXT", "Exporta Reporte a txt " & file, "RE"
End Sub


