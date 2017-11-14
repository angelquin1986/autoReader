VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{50067EB3-D6AF-11D3-8297-000021C5085D}#1.0#0"; "NTextBox.ocx"
Begin VB.Form frmIVBusquedaPNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración"
   ClientHeight    =   5370
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdArchivo 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5700
      TabIndex        =   11
      Top             =   4380
      Width           =   372
   End
   Begin VB.TextBox txtPlantilla 
      Height          =   345
      Left            =   120
      TabIndex        =   10
      Text            =   "txtPlantilla"
      Top             =   4380
      Width           =   5535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Rango Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   4
      Top             =   420
      Width           =   5895
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   330
         Left            =   840
         TabIndex        =   5
         ToolTipText     =   "Fecha de la transacción"
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
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
         CustomFormat    =   "yyyy/MM/dd"
         Format          =   106692609
         CurrentDate     =   37078
         MaxDate         =   73415
         MinDate         =   29221
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   330
         Left            =   3000
         TabIndex        =   6
         ToolTipText     =   "Fecha de la transacción"
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
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
         CustomFormat    =   "yyyy/MM/dd"
         Format          =   106692609
         CurrentDate     =   37078
         MaxDate         =   73415
         MinDate         =   29221
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2400
         TabIndex        =   7
         Top             =   360
         Width           =   420
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar -F5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1260
      TabIndex        =   0
      Top             =   4860
      Width           =   1452
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   2880
      TabIndex        =   1
      Top             =   4860
      Width           =   1452
   End
   Begin NTextBoxProy.NTextBox ntxCostosFijosMen 
      Height          =   330
      Left            =   1920
      TabIndex        =   2
      Top             =   0
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
      Text            =   "0"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AllowDecimal    =   -1  'True
   End
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   5460
      Top             =   -60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab sst 
      Height          =   2895
      Left            =   120
      TabIndex        =   12
      Top             =   1140
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5106
      _Version        =   393216
      Tabs            =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Bodegas"
      TabPicture(0)   =   "IVBusquedaPNew.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lstBodegas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Ventas"
      TabPicture(1)   =   "IVBusquedaPNew.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstVentas"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Devoluciones"
      TabPicture(2)   =   "IVBusquedaPNew.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstNotaCredito"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Producción"
      TabPicture(3)   =   "IVBusquedaPNew.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstProd"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "MEM"
      TabPicture(4)   =   "IVBusquedaPNew.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lstMEM"
      Tab(4).ControlCount=   1
      Begin VB.ListBox lstBodegas 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   60
         Style           =   1  'Checkbox
         TabIndex        =   17
         Top             =   720
         Width           =   5595
      End
      Begin VB.ListBox lstVentas 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   -74940
         Style           =   1  'Checkbox
         TabIndex        =   16
         Top             =   720
         Width           =   5595
      End
      Begin VB.ListBox lstNotaCredito 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   -74940
         Style           =   1  'Checkbox
         TabIndex        =   15
         Top             =   720
         Width           =   5595
      End
      Begin VB.ListBox lstProd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   -74940
         Style           =   1  'Checkbox
         TabIndex        =   14
         Top             =   720
         Width           =   5595
      End
      Begin VB.ListBox lstMEM 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   -74940
         Style           =   1  'Checkbox
         TabIndex        =   13
         Top             =   720
         Width           =   5595
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Ruta de Plantilla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   4140
      Width           =   1155
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Costos Fijos Mensuales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1650
   End
End
Attribute VB_Name = "frmIVBusquedaPNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BandAceptado As Boolean
Dim arhivoPlantilla As String
Dim strBodegas As String
Dim TransVentas As String
Dim TransProd As String
Dim TransMEM As String
Dim TransNotaCredito As String

Public Function Inicio(ByRef desde As String, _
                       ByRef hasta As String, _
                       ByRef CostoMen As Currency, ByRef v() As String) As Boolean
    Dim antes As String, i As Integer, j As Integer
    Dim strBod As String, Bodega As String
    Dim strVentas As String, Ventas As String
    Dim strProduccion As String, Produccion As String
    Dim strEntrega As String, Entrega As String
    Dim strNotaCredito As String, NotaCredito As String
    Dim W As Variant
    On Error GoTo ErrTrap
    
    'Cambia forma de cursor mientras se carga
    MensajeStatus MSG_PREPARA, vbHourglass
    
    CargarBodegas
    CargarTransVenta
    CargarTransNotaCredito
    CargarTransProduccion
    CargarTransMEM
    recuperaConfiguraciones
    'carga bodegas predeteminadas
         For i = 0 To lstBodegas.ListCount - 1
            Bodega = GetSetting(APPNAME, App.Title, "strBodegas", "")
            strBod = Left$(lstBodegas.List(i), lstBodegas.ItemData(i))
            If InStr(Bodega, strBod) Then
                lstBodegas.Selected(i) = True
            Else
                lstBodegas.Selected(i) = False
            End If
        Next i
        'carga ventas predeterminadas
         Ventas = GetSetting(APPNAME, App.Title, "TransVentas", "")
         
        W = Split(Ventas, ",")

        For j = 0 To UBound(W)
             For i = 0 To lstVentas.ListCount - 1
                strVentas = Left$(lstVentas.List(i), lstVentas.ItemData(i))
                If W(j) = strVentas Then
                    lstVentas.Selected(i) = True
                    Exit For
                End If
            Next i
        Next j

        'carga NotaCredito predeterminadas
         NotaCredito = GetSetting(APPNAME, App.Title, "TransNotaCredito", "")
         
        W = Split(NotaCredito, ",")

        For j = 0 To UBound(W)
             For i = 0 To lstNotaCredito.ListCount - 1
                strNotaCredito = Left$(lstNotaCredito.List(i), lstNotaCredito.ItemData(i))
                If W(j) = strNotaCredito Then
                    lstNotaCredito.Selected(i) = True
                    Exit For
                End If
            Next i
        Next j


        'carga Transacciones Produccion
         Produccion = GetSetting(APPNAME, App.Title, "TransProd", "")
         For i = 0 To lstProd.ListCount - 1
            strProduccion = Left$(lstProd.List(i), lstProd.ItemData(i))
            If InStr(Produccion, strProduccion) Then
                lstProd.Selected(i) = True
            Else
                lstProd.Selected(i) = False
            End If
        Next i
    
        'carga Transacciones Entrega
         Entrega = GetSetting(APPNAME, App.Title, "TransMEM", "")
         For i = 0 To lstMEM.ListCount - 1
            strEntrega = Left$(lstMEM.List(i), lstMEM.ItemData(i))
            If InStr(Entrega, strEntrega) Then
                lstMEM.Selected(i) = True
            Else
                lstMEM.Selected(i) = False
            End If
        Next i
    
    MensajeStatus
    BandAceptado = False
    Me.Show vbModal, frmMain
    
    If BandAceptado Then
        CostoMen = ntxCostosFijosMen.value
        desde = dtpDesde.value
        hasta = dtpHasta.value
    End If
    'Devuelve true/false
    Inicio = BandAceptado
    Exit Function
ErrTrap:
    MensajeStatus
    DispErr
    Exit Function
End Function



Private Sub cmdAceptar_Click()
   GrabaConfiguraciones
    BandAceptado = True
    'txtCodigo.SetFocus
    Me.Hide
End Sub

Private Sub cmdArchivo_Click()
 On Error GoTo ErrTrap
    With frmIVBusquedaP.dlg1
        .InitDir = App.Path
        .CancelError = True
        .Filter = "Texto (Separado por coma *.csv)|*.csv|Texto (Separado por tabuladores *.txt)|*.txt|Todos *.*|*.*"
        .flags = cdlOFNFileMustExist
        .ShowOpen
        txtPlantilla.Text = .filename
        arhivoPlantilla = .filename
    End With
    Exit Sub
ErrTrap:
End Sub

Private Sub cmdCancelar_Click()
    BandAceptado = False
 '   txtCodigo.SetFocus
    Me.Hide
End Sub

Private Sub Form_Activate()
    Dim c As Control, band As Boolean, c2 As Control
    On Error Resume Next
    If Not Me.Visible Then Exit Sub
    
    'Busca un TextBox que tenga alguna cadena
'    Set c2 = txtCodigo
    For Each c In Me.Controls
        If TypeName(c) = "TextBox" Then
            If Len(c.Text) > 0 Then 'Si encuentra,
                If (c.TabIndex < c2.TabIndex) _
                    Or (Len(c2.Text) = 0) Then Set c2 = c
            End If
        End If
    Next c
      c2.SetFocus
    End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF5
        cmdAceptar_Click
        KeyCode = 0
    Case Else
        MoverCampo Me, KeyCode, Shift, False
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ImpideSonidoEnter Me, KeyAscii
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Me.Hide         'Se pone esto para evitar el posible BUG de Windows98
End Sub



Private Sub CargarBodegas()
    Dim v As Variant, i As Long
    lstBodegas.Clear
    v = gobjMain.EmpresaActual.ListaIVBodega(False, False)
    If Not IsEmpty(v) Then
        For i = 0 To UBound(v, 2)
            lstBodegas.AddItem v(0, i) & " " & v(1, i)
            lstBodegas.ItemData(lstBodegas.NewIndex) = Len(v(0, i))
        Next i
    End If
End Sub

Private Sub CargarTransVenta()
    Dim v As Variant, i As Long
    
    lstVentas.Clear
    v = gobjMain.EmpresaActual.ListaGNTrans("", False, False)
    If Not IsEmpty(v) Then
        For i = 0 To UBound(v, 2)
            lstVentas.AddItem v(0, i) & " " & v(1, i)
            lstVentas.ItemData(lstVentas.NewIndex) = Len(v(0, i))
        Next i
    End If
End Sub

Private Sub CargarTransProduccion()
    Dim v As Variant, i As Long
    lstProd.Clear
    v = gobjMain.EmpresaActual.ListaGNTrans("", False, False)
    If Not IsEmpty(v) Then
        For i = 0 To UBound(v, 2)
            lstProd.AddItem v(0, i) & " " & v(1, i)
            lstProd.ItemData(lstProd.NewIndex) = Len(v(0, i))
        Next i
    End If
End Sub



Private Sub recuperaConfiguraciones()
Dim i As Integer
   ntxCostosFijosMen.value = GetSetting(APPNAME, App.Title, "costofijomensual", 0)
   dtpDesde.value = GetSetting(APPNAME, App.Title, "Desde", Date)
   dtpHasta.value = GetSetting(APPNAME, App.Title, "Hasta", Date)
   txtPlantilla = GetSetting(APPNAME, App.Title, "Ruta Plantilla", "")
   
   
   
      
End Sub
Private Sub GrabaConfiguraciones()
Dim i As Integer
    SaveSetting APPNAME, App.Title, "costofijomensual", ntxCostosFijosMen.value
    SaveSetting APPNAME, App.Title, "Desde", dtpDesde.value
    SaveSetting APPNAME, App.Title, "Hasta", dtpHasta.value
    SaveSetting APPNAME, App.Title, "Ruta Plantilla", txtPlantilla.Text
    SaveSetting APPNAME, App.Title, "strBodegas", strBodegas
    SaveSetting APPNAME, App.Title, "TransVentas", TransVentas
    SaveSetting APPNAME, App.Title, "TransProd", TransProd
    SaveSetting APPNAME, App.Title, "TransMEM", TransMEM
   SaveSetting APPNAME, App.Title, "TransNotaCredito", TransNotaCredito
      
End Sub

'
Private Sub lstBodegas_Click()
Dim s As String, i As Long
    On Error GoTo ErrTrap
    'If mbooIniciando Then Exit Sub

    With lstBodegas
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                If Len(s) > 0 Then s = s & ","
                s = s & Left$(.List(i), .ItemData(i))
            End If
        Next i
    End With
     strBodegas = s
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub lstProd_Click()
Dim s As String, i As Long
    On Error GoTo ErrTrap
    
    With lstProd
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                If Len(s) > 0 Then s = s & ","
                s = s & Left$(.List(i), .ItemData(i))
            End If
        Next i
    End With
     TransProd = s
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub lstVentas_Click()
Dim s As String, i As Long
    On Error GoTo ErrTrap
    With lstVentas
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                If Len(s) > 0 Then s = s & ","
                s = s & Left$(.List(i), .ItemData(i))
            End If
        Next i
    End With
     TransVentas = s
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub CargarTransMEM()
    Dim v As Variant, i As Long
    lstMEM.Clear
    v = gobjMain.EmpresaActual.ListaGNTrans("", False, False)
    If Not IsEmpty(v) Then
        For i = 0 To UBound(v, 2)
            lstMEM.AddItem v(0, i) & " " & v(1, i)
            lstMEM.ItemData(lstMEM.NewIndex) = Len(v(0, i))
        Next i
    End If
End Sub


Private Sub lstMEM_Click()
Dim s As String, i As Long
    On Error GoTo ErrTrap
    
    With lstMEM
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                If Len(s) > 0 Then s = s & ","
                s = s & Left$(.List(i), .ItemData(i))
            End If
        Next i
    End With
     TransMEM = s
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub CargarTransNotaCredito()
    Dim v As Variant, i As Long
    
    lstNotaCredito.Clear
    v = gobjMain.EmpresaActual.ListaGNTrans("", False, False)
    If Not IsEmpty(v) Then
        For i = 0 To UBound(v, 2)
            lstNotaCredito.AddItem v(0, i) & " " & v(1, i)
            lstNotaCredito.ItemData(lstNotaCredito.NewIndex) = Len(v(0, i))
        Next i
    End If
End Sub

Private Sub lstNotaCredito_Click()
Dim s As String, i As Long
    On Error GoTo ErrTrap
    With lstNotaCredito
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                If Len(s) > 0 Then s = s & ","
                s = s & Left$(.List(i), .ItemData(i))
            End If
        Next i
    End With
     TransNotaCredito = s
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub
