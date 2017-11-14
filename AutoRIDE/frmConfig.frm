VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{ED5A9B02-5BDB-48C7-BAB1-642DCC8C9E4D}#2.0#0"; "SelFold.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConfig 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
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
      Left            =   5160
      TabIndex        =   15
      Top             =   360
      Width           =   372
   End
   Begin VB.TextBox txtArchivo 
      Height          =   288
      Left            =   120
      TabIndex        =   14
      Top             =   360
      Width           =   4995
   End
   Begin TabDlg.SSTab sst1 
      Height          =   2955
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   5212
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Exportación -F5"
      TabPicture(0)   =   "frmConfig.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblubicacion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "slf"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dtpHora"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtCarpeta"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdExaminarCarpeta"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "dtpFecha"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Importar -F6"
      TabPicture(1)   =   "frmConfig.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "txtCarpetaD"
      Tab(1).Control(2)=   "cmdExaminarCarpetaD"
      Tab(1).Control(3)=   "txtCarpetaO"
      Tab(1).Control(4)=   "cmdExaminarCarpetaO"
      Tab(1).Control(5)=   "Label5"
      Tab(1).Control(6)=   "Label1"
      Tab(1).ControlCount=   7
      Begin VB.Frame Frame2 
         Caption         =   "Selecione la Plantilla a utiilizar"
         Height          =   975
         Left            =   -74880
         TabIndex        =   23
         Top             =   480
         Width           =   5115
         Begin VB.ComboBox cboPlantillaI 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   27
            Top             =   300
            Width           =   495
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   840
         End
         Begin VB.Label lblDescripcionI 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1080
            TabIndex        =   25
            Top             =   600
            Width           =   3915
         End
      End
      Begin VB.TextBox txtCarpetaD 
         Height          =   312
         Left            =   -73800
         TabIndex        =   21
         Text            =   "c:\"
         Top             =   1920
         Width           =   3510
      End
      Begin VB.CommandButton cmdExaminarCarpetaD 
         Caption         =   "..."
         Height          =   300
         Left            =   -70260
         TabIndex        =   20
         Top             =   1920
         Width           =   372
      End
      Begin VB.TextBox txtCarpetaO 
         Height          =   312
         Left            =   -73800
         TabIndex        =   18
         Text            =   "c:\"
         Top             =   1560
         Width           =   3510
      End
      Begin VB.CommandButton cmdExaminarCarpetaO 
         Caption         =   "..."
         Height          =   300
         Left            =   -70260
         TabIndex        =   17
         Top             =   1560
         Width           =   372
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   375
         Left            =   2400
         TabIndex        =   12
         Top             =   1920
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16777217
         CurrentDate     =   40647
      End
      Begin VB.CommandButton cmdExaminarCarpeta 
         Caption         =   "..."
         Height          =   300
         Left            =   4800
         TabIndex        =   8
         Top             =   1560
         Width           =   372
      End
      Begin VB.TextBox txtCarpeta 
         Height          =   312
         Left            =   1260
         TabIndex        =   7
         Text            =   "c:\"
         Top             =   1560
         Width           =   3510
      End
      Begin VB.Frame Frame1 
         Caption         =   "Selecione la Plantilla a utiilizar"
         Height          =   975
         Left            =   180
         TabIndex        =   2
         Top             =   540
         Width           =   5115
         Begin VB.ComboBox cboPlantilla 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label lblDescripcion 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1080
            TabIndex        =   6
            Top             =   600
            Width           =   3915
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   840
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Código"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   4
            Top             =   300
            Width           =   495
         End
      End
      Begin MSComCtl2.DTPicker dtpHora 
         Height          =   375
         Left            =   2400
         TabIndex        =   13
         Top             =   2280
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   661
         _Version        =   393216
         Format          =   16777218
         CurrentDate     =   40647
      End
      Begin SelFold.SelFolder slf 
         Left            =   4500
         Top             =   1920
         _ExtentX        =   1614
         _ExtentY        =   238
         Title           =   "Seleccione una carpeta"
         Caption         =   "Selección de carpeta"
         RootFolder      =   "\"
         Path            =   "c:\Vbprog_esp\Sii\SelFold"
      End
      Begin VB.Label Label5 
         Caption         =   "Destino:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   22
         Top             =   1980
         Width           =   810
      End
      Begin VB.Label Label1 
         Caption         =   "Origen:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   19
         Top             =   1620
         Width           =   810
      End
      Begin VB.Label Label2 
         Caption         =   "Ultima Hora de Exportación"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2400
         Width           =   2475
      End
      Begin VB.Label Label4 
         Caption         =   "Utima Fecha de Exportación"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   2040
         Width           =   2475
      End
      Begin VB.Label lblubicacion 
         Caption         =   "Destino:"
         Height          =   255
         Left            =   300
         TabIndex        =   9
         Top             =   1620
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar -F9"
      Height          =   372
      Left            =   2280
      TabIndex        =   0
      Top             =   3780
      Width           =   1428
   End
   Begin VB.Label Label3 
      Caption         =   "Archivo de Plantillas:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   60
      Width           =   1935
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BandAceptar  As Boolean
Dim mobjGNTrans As GNTrans
Dim mAuxOpt As Integer
Dim mOpt As String
Private NPCGRUPOCREDCLI As Integer
Private mobjGNOp As GNOpcion
Private WithEvents mobjEmpresa As Empresa
Attribute mobjEmpresa.VB_VarHelpID = -1
Private mcnPlantilla As ADODB.Connection '***Angel. 13/feb/2004
Private mPlantilla As clsPlantilla       '***Angel. 20/feb/2004
Private mPlantillaI As clsPlantilla       '***Angel. 20/feb/2004
Private mUltimaPlantilla As String
Private mUltimaPlantillaI As String
Private RutaBD As String
Private NombreBD As String
Private sucursal As String

Public Function Inicio() As Boolean
    sucursal = GetSetting(APPNAME, SECTION, "CodSucursal_Activa_" & gobjMain.EmpresaActual.CodEmpresa, "")
    AbrirBasePlantilla
    CargarComboPlantilla
    CargarComboPlantillaI
    PrepararPlantilla
    visualizar
    Me.Show vbModal
    Unload Me
    Inicio = BandAceptar
End Function


Private Sub cboPlantilla_Change()
    RecuperarPlantilla
End Sub

Private Sub cboPlantillaI_Change()
    RecuperarPlantillai
End Sub


Private Sub cboPlantilla_Click()
    RecuperarPlantilla
    PrepararPlantilla
End Sub

Private Sub cboPlantillaI_Click()
    RecuperarPlantillai
    PrepararPlantillaI
End Sub


Private Sub cmdAceptar_Click()
    sucursal = GetSetting(APPNAME, SECTION, "CodSucursal_Activa_" & gobjMain.EmpresaActual.CodEmpresa, "")
    
    gobjMain.EmpresaActual.GNOpcion.AsignarValor "Plantilla_AutoExportacion_" & sucursal, cboPlantilla.Text
    gobjMain.EmpresaActual.GNOpcion.AsignarValor "RutaArchivo_AutoExportacion_" & sucursal, txtCarpeta.Text
    gobjMain.EmpresaActual.GNOpcion.AsignarValor "FechaUltima_AutoExportacion_" & sucursal, dtpFecha.value
    gobjMain.EmpresaActual.GNOpcion.AsignarValor "HoraUltima_AutoExportacion_" & sucursal, dtpHora.value
    gobjMain.EmpresaActual.GNOpcion.AsignarValor "RutaPlantilla_" & sucursal, RutaBD
    gobjMain.EmpresaActual.GNOpcion.AsignarValor "NombrePlantilla_" & sucursal, NombreBD
    
    
    
    
    gobjMain.EmpresaActual.GNOpcion.AsignarValor "Plantilla_AutoImport", cboPlantillaI.Text
    gobjMain.EmpresaActual.GNOpcion.AsignarValor "RutaArchivoO_AutoImport", txtCarpetaO.Text
    gobjMain.EmpresaActual.GNOpcion.AsignarValor "RutaArchivoD_AutoImport", txtCarpetaD.Text
    
    gobjMain.EmpresaActual.GNOpcion.GrabarGNOpcion2
    
    
        
        
    BandAceptar = True
    
    Unload Me
    Set mobjGNOp = Nothing
    Set mobjEmpresa = Nothing
End Sub



Private Sub cmdExaminarCarpetaD_Click()
    On Error GoTo Errtrap
    slf.OwnerHWnd = Me.hWnd
    slf.Path = txtCarpetaD.Text
    If slf.Browse Then
        txtCarpetaD.Text = slf.Path
        txtCarpetaD_LostFocus
    End If
    Exit Sub
Errtrap:
    MsgBox Err.Description, vbInformation
    Exit Sub

End Sub

Private Sub cmdExaminarCarpetaO_Click()
    On Error GoTo Errtrap
    slf.OwnerHWnd = Me.hWnd
    slf.Path = txtCarpetaO.Text
    If slf.Browse Then
        txtCarpetaO.Text = slf.Path
        txtCarpetaO_LostFocus
    End If
    Exit Sub
Errtrap:
    MsgBox Err.Description, vbInformation
    Exit Sub

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF9
        cmdAceptar_Click
        KeyCode = 0
    Case vbKeyF5, vbKeyF6, vbKeyF7, vbKeyF8
    Case Else
        MoverCampo Me, KeyCode, Shift, True
    End Select
End Sub




Private Sub cmdExaminarCarpeta_Click()
    On Error GoTo Errtrap
    slf.OwnerHWnd = Me.hWnd
    slf.Path = txtCarpeta.Text
    If slf.Browse Then
        txtCarpeta.Text = slf.Path
        txtCarpeta_LostFocus
    End If
    Exit Sub
Errtrap:
    MsgBox Err.Description, vbInformation
    Exit Sub
End Sub


Private Sub txtCarpeta_LostFocus()
    If Right$(txtCarpeta.Text, 1) <> "\" Then
        txtCarpeta.Text = txtCarpeta.Text & "\"
    End If
    'Luego a actualiza linea de comando
    
End Sub

Private Sub cmdPlantilla_Click()
    If Len(cboPlantilla.Text) > 0 Then
        CargarComboPlantilla
    End If
End Sub

Private Sub cmdPlantillai_Click()
    If Len(cboPlantillaI.Text) > 0 Then
        CargarComboPlantillaI
    End If
End Sub


Private Sub CargarComboPlantilla()
    Dim sql As String, rs As Recordset
        
    sql = "SELECT CodPlantilla, Descripcion FROM Plantilla_EI " & _
          "WHERE (Tipo=0) AND (BandValida=True) ORDER BY CodPlantilla"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnPlantilla, adOpenStatic, adLockReadOnly
    cboPlantilla.Clear
    With rs
        If Not (.BOF And .EOF) Then
            Do Until .EOF
                cboPlantilla.AddItem !CodPlantilla
                .MoveNext
            Loop
        End If
    End With
    Set rs = Nothing
    
    
    
End Sub


Private Sub AbrirBasePlantilla()
    Dim s As String, RutaBD As String, NombreBD As String
    On Error GoTo Errtrap
    
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("RutaBDPlantilla")) <> 0 Then
        RutaBD = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("RutaBDPlantilla")
    Else
        RutaBD = GetSetting(APPNAME, App.Title, "RutaBDPlantilla", App.Path)
        If Right(RutaBD, 1) <> "\" Then RutaBD = RutaBD & "\"
    End If
    
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("NombreBDPlantilla")) <> 0 Then
        NombreBD = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("NombreBDPlantilla")
    Else
        NombreBD = GetSetting(APPNAME, App.Title, "NombreBDPlantilla", "ConfigSiiToolsA.mdb")
    End If

    
    
    txtArchivo.Text = RutaBD & NombreBD
    
    'NombreBD = GetSetting(APPNAME, App.Title, "NombreBDPlantilla", "ConfigSiiToolsA.mdb")
    
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "RutaBDPlantilla", RutaBD
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "NombreBDPlantilla", NombreBD
        gobjMain.EmpresaActual.GNOpcion.GrabarGNOpcion2
    

    If mcnPlantilla Is Nothing Then Set mcnPlantilla = New ADODB.Connection
    If mcnPlantilla.State <> adStateClosed Then mcnPlantilla.Close
    
    'Abre la conección con el archivo de destino
    s = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & RutaBD & NombreBD & ";" & _
        "Persist Security Info=False"
    mcnPlantilla.Open s, "admin", ""
    
    Exit Sub

Errtrap:
    MsgBox Err.Description, vbOKOnly + vbInformation
End Sub

'***Angel. 13/feb/2004
Private Sub CerrarBasePlantilla()
    'Cierra base
    On Error Resume Next
    If mcnPlantilla.State <> adStateClosed Then
        mcnPlantilla.Close
    End If
End Sub


Private Sub RecuperarPlantilla()
    Dim cod As String
    
    cod = cboPlantilla.Text
    If mPlantilla.Recuperar(cod) Then
        lblDescripcion.Caption = mPlantilla.Descripcion
    Else
        lblDescripcion.Caption = ""
    End If
End Sub

Private Sub PrepararPlantilla()
    Dim i As Integer
    On Error GoTo mensaje
    
    Set mPlantilla = New clsPlantilla
    mPlantilla.Coneccion = mcnPlantilla
    RecuperarPlantilla
    Exit Sub
    
mensaje:
    MsgBox Err.Description, vbOKOnly + vbExclamation
End Sub


Private Sub visualizar()
    Dim i As Integer

    If Len(sucursal) > 0 Then
        mUltimaPlantilla = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Plantilla_AutoExportacion_" & sucursal)
        
        If Len(mUltimaPlantilla) Then
            For i = 0 To cboPlantilla.ListCount - 1
                If cboPlantilla.List(i) = mUltimaPlantilla Then
                    cboPlantilla.ListIndex = i
                    Exit For
                End If
            Next i
        End If
        
        mUltimaPlantillaI = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Plantilla_AutoImport_" & sucursal)
        
        If Len(mUltimaPlantillaI) Then
            For i = 0 To cboPlantillaI.ListCount - 1
                If cboPlantillaI.List(i) = mUltimaPlantillaI Then
                    cboPlantillaI.ListIndex = i
                    Exit For
                End If
            Next i
        End If
    
    

        txtCarpeta.Text = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("RutaArchivo_AutoExportacion_" & sucursal)
        dtpFecha.value = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("FechaUltima_AutoExportacion_" & sucursal)
        dtpHora.value = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("HoraUltima_AutoExportacion_" & sucursal)
        RutaBD = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("RutaPlantilla_" & sucursal)
        If Right(RutaBD, 1) <> "\" Then RutaBD = RutaBD & "\"
        NombreBD = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("NombrePlantilla_" & sucursal)
        txtArchivo.Text = RutaBD & NombreBD
        
        txtCarpetaO.Text = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("RutaArchivoO_AutoImport")
        txtCarpetaD.Text = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("RutaArchivoD_AutoImport")
        
        
        
        
    End If
    

End Sub


Private Sub cmdArchivo_Click()
    Dim v As Variant, max As Integer
    Dim cadena As String, i As Integer
    On Error GoTo Errtrap
    With frmMain.dlg1
        .InitDir = App.Path
        .CancelError = True
        .Filter = "Base de Datos(mdb)|*.mdb"
        .ShowOpen
        txtArchivo.Text = .FileName
        NombreBD = .FileTitle
        v = Split(.FileName, "\")
        cadena = ""
        max = UBound(v)
        For i = 0 To max - 1
            cadena = cadena & v(i) & "\"
        Next i
        RutaBD = cadena
    End With
    Exit Sub
Errtrap:
    
End Sub

Private Sub txtCarpetaD_LostFocus()
    If Right$(txtCarpetaD.Text, 1) <> "\" Then
        txtCarpetaD.Text = txtCarpetaD.Text & "\"
    End If

End Sub


Private Sub txtCarpetaO_LostFocus()
    If Right$(txtCarpetaO.Text, 1) <> "\" Then
        txtCarpetaO.Text = txtCarpetaO.Text & "\"
    End If

End Sub

Private Sub RecuperarPlantillai()
    Dim cod As String
    
    cod = cboPlantillaI.Text
    If mPlantilla.Recuperar(cod) Then
        lblDescripcionI.Caption = mPlantilla.Descripcion
    Else
        lblDescripcionI.Caption = ""
    End If
End Sub

Private Sub PrepararPlantillaI()
    Dim i As Integer
    On Error GoTo mensaje
    
    Set mPlantillaI = New clsPlantilla
    mPlantillaI.Coneccion = mcnPlantilla
    RecuperarPlantillai
    Exit Sub
    
mensaje:
    MsgBox Err.Description, vbOKOnly + vbExclamation
End Sub

Private Sub CargarComboPlantillaI()
    Dim sql As String, rs As Recordset
        
    sql = "SELECT CodPlantilla, Descripcion FROM Plantilla_EI " & _
          "WHERE (Tipo=1) AND (BandValida=True) ORDER BY CodPlantilla"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnPlantilla, adOpenStatic, adLockReadOnly
    cboPlantillaI.Clear
    With rs
        If Not (.BOF And .EOF) Then
            Do Until .EOF
                cboPlantillaI.AddItem !CodPlantilla
                .MoveNext
            Loop
        End If
    End With
    Set rs = Nothing
End Sub


