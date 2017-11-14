VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración"
   ClientHeight    =   1230
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   5880
   HelpContextID   =   2
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   5880
   Begin VB.TextBox txtArchivo 
      Height          =   288
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   5292
   End
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
      Left            =   5400
      TabIndex        =   3
      Top             =   360
      Width           =   372
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   372
      Left            =   3060
      TabIndex        =   1
      Top             =   735
      Width           =   1332
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar -F5"
      Height          =   372
      Left            =   1500
      TabIndex        =   0
      Top             =   720
      Width           =   1332
   End
   Begin VB.Label Label3 
      Caption         =   "Archivo de Plantillas:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
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
Private RutaBD As String
Private NombreBD As String
Dim bandAceptar  As Boolean

Private Sub cmdAceptar_Click()

    bandAceptar = True
    If bandAceptar Then
        SaveSetting APPNAME, "SiiToolsA", "RutaBDPlantilla", RutaBD
        SaveSetting APPNAME, "SiiToolsA", "NombreBDPlantilla", NombreBD
        
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "RutaBDPlantilla", RutaBD
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "NombreBDPlantilla", NombreBD
        gobjMain.EmpresaActual.GNOpcion.GrabarGNOpcion2
    End If
    Unload Me
End Sub

Public Sub Inicio()
    
    
    
    
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
    Me.Show
End Sub


Private Sub cmdArchivo_Click()
    Dim v As Variant, max As Integer
    Dim Cadena As String, i As Integer
    On Error GoTo ErrTrap
    With frmMain.dlg1
        .InitDir = App.Path
        .CancelError = True
        .Filter = "Base de Datos(mdb)|*.mdb"
        .ShowOpen
        txtArchivo.Text = .filename
        NombreBD = .FileTitle
        v = Split(.filename, "\")
        Cadena = ""
        max = UBound(v)
        For i = 0 To max - 1
            Cadena = Cadena & v(i) & "\"
        Next i
        RutaBD = Cadena
    End With
    Exit Sub
ErrTrap:
    
End Sub

Private Sub cmdCancelar_Click()
    bandAceptar = False
    Unload Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF5
        cmdAceptar_Click
        KeyCode = 0
    Case vbKeyEscape
        cmdCancelar_Click
    Case Else
        MoverCampo Me, KeyCode, Shift, False
    End Select
End Sub

