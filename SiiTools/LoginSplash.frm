VERSION 5.00
Begin VB.Form frmLoginSplash 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sii Inicio de sesión"
   ClientHeight    =   5190
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   8100
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Inicio de sesión"
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   3540
      Picture         =   "LoginSplash.frx":0000
      ScaleHeight     =   1035
      ScaleWidth      =   1935
      TabIndex        =   17
      Top             =   1080
      Width           =   1935
   End
   Begin VB.PictureBox picbotones 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1665
      Left            =   0
      ScaleHeight     =   1665
      ScaleWidth      =   8100
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3525
      Width           =   8100
      Begin VB.TextBox txtNombre 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   900
         TabIndex        =   0
         Top             =   1080
         Width           =   1488
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Ingresar"
         Height          =   360
         Left            =   5640
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   2
         Tag             =   "Aceptar"
         Top             =   1080
         Width           =   1620
      End
      Begin VB.TextBox txtClave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   3360
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1080
         Width           =   1488
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "CUENCA - ECUADOR www.ibzssoft.com"
         BeginProperty Font 
            Name            =   "Sansation"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   600
         TabIndex        =   20
         Top             =   480
         Width           =   2355
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "info@ibzssoft.com ishidacue@hotmail.com"
         BeginProperty Font 
            Name            =   "Sansation"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   4440
         TabIndex        =   19
         Top             =   480
         Width           =   3555
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Aeropuerto Mariscal Lamar Segundo Piso                           Telf: 072870346/0998499003 "
         BeginProperty Font 
            Name            =   "Sansation"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   240
         Width           =   8115
      End
      Begin VB.Label Label3 
         BackColor       =   &H000000C0&
         Caption         =   "Asesoría Contable y administrativa"
         BeginProperty Font 
            Name            =   "Sansation"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5100
         TabIndex        =   14
         Top             =   0
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000C0&
         Caption         =   "Venta y mantenimiento de equipos"
         BeginProperty Font 
            Name            =   "Sansation"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1980
         TabIndex        =   13
         Top             =   0
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000C0&
         Caption         =   "Desarrollo de Software"
         BeginProperty Font 
            Name            =   "Sansation"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1995
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         Caption         =   "Usuari&o    "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Tag             =   "Usuari&o:"
         Top             =   1140
         Width           =   645
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackColor       =   &H000000C0&
         Caption         =   "&Clave    "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   2820
         TabIndex        =   4
         Tag             =   "&Contraseña:"
         Top             =   1140
         Width           =   525
      End
   End
   Begin VB.Label lblPlatform 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ishida Business Software para Windows"
      BeginProperty Font 
         Name            =   "Sansation"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   705
      Left            =   2880
      TabIndex        =   18
      Tag             =   "Plataforma"
      Top             =   2040
      Width           =   3285
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "z"
      BeginProperty Font 
         Name            =   "Sansation"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4500
      TabIndex        =   16
      Tag             =   "Producto"
      Top             =   1800
      Width           =   150
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Copyright"
      BeginProperty Font 
         Name            =   "Sansation"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Tag             =   "Copyright"
      Top             =   2100
      Width           =   2115
   End
   Begin VB.Label lblCompany 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Organización"
      BeginProperty Font 
         Name            =   "Sansation"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Tag             =   "Organización"
      Top             =   1860
      Width           =   2115
   End
   Begin VB.Label lblWarning 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Advertencia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   9
      Tag             =   "Advertencia"
      Top             =   3240
      Width           =   7935
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Versión 4"
      BeginProperty Font 
         Name            =   "Sansation"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   225
      Left            =   120
      TabIndex        =   8
      Tag             =   "Versión"
      Top             =   1620
      Width           =   2115
   End
   Begin VB.Label lblLicenseTo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "LicenciaA"
      Height          =   255
      Left            =   2370
      TabIndex        =   7
      Tag             =   "LicenciaA"
      Top             =   120
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblProductName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "IB S"
      BeginProperty Font 
         Name            =   "Sansation"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   3840
      TabIndex        =   6
      Tag             =   "Producto"
      Top             =   2700
      Width           =   960
   End
   Begin VB.Image Image2 
      Height          =   3495
      Left            =   0
      Picture         =   "LoginSplash.frx":6342
      Top             =   0
      Width           =   8580
   End
End
Attribute VB_Name = "frmLoginSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Aceptado As Boolean
Private PasswordCorrecto As Boolean
Private bandPassword As Boolean
Private usuario As usuario

Public Function Inicio() As Boolean
    Aceptado = False
    bandPassword = False
    Me.Show vbModal, frmMain
    Inicio = Aceptado
    Unload Me
End Function

Private Sub cmdAceptar_Click()
    Dim us As usuario
    Dim usClave As usuario
    Dim BandModulo As Boolean
    On Error GoTo errtrap
    If bandPassword Then
        PasswordCorrecto = gobjMain.LoginVen(Trim$(txtNombre), Trim$(txtClave))
        If PasswordCorrecto Then
            Aceptado = True
            gobjMain.EmpresaActual.GrabaGNLogAccion "CLA_VTA", "Clave para Venta:" & txtNombre.Text, "IV"
            Me.Hide
        Else
            MsgBox "Nombre o clave está incorrecto. Intente de nuevo."
            txtNombre.SetFocus
        End If
        Exit Sub
    Else
       Set us = gobjMain.Login(Trim$(txtNombre), Trim$(txtClave))
        If Not us Is Nothing Then
                'verifica si es la primera vez ingresando al sistema y pide que cambie la clave
                Set usClave = us
                If usClave.BandCambiaClave Then
                    txtClave.Text = frmCambiaClave.Modificar(usClave.codUsuario, "Usuario", "Cambio de clave primera vez")
                     
                End If
                Set usClave = Nothing
            'jeaa 23/09/2008
            Set us = gobjMain.Login(Trim$(txtNombre), Trim$(txtClave))
            Set usuario = us
            If Not us Is Nothing Then
                 If Not us.BandValida Then
                    MsgBox "El usuario " & us.NombreUsuario & " no esta Activo" & Chr(13) & _
                           "Pongase en contacto con el administrador del sistema", vbInformation
                     End
                 End If
                 'AUC verifica modulo
                BandModulo = gobjMain.PermisoModulo(Trim$(txtNombre.Text), Trim$(txtClave.Text), ModuloTools)
                If Not BandModulo Then
                    MsgBox "No tiene permiso para abrir este modulo" & Chr(13) & _
                           "Pongase en contacto con el administrador del sistema", vbInformation
                       End
                End If
                gobjMain.ModuloCargado = ModuloTools
                '-----------------
                 Aceptado = True
                 Me.Hide
            Else
                MsgBox "Nombre o clave está incorrecto. Intente de nuevo."
                txtNombre.SetFocus
            End If
            Set us = Nothing
            
        Else
            MsgBox "Nombre o clave está incorrecto. Intente de nuevo."
            txtNombre.SetFocus
        End If
        Set us = Nothing

        Exit Sub
    End If
errtrap:
    DispErr
    Exit Sub
End Sub




Private Sub cmdCancelar_Click()
    Aceptado = False
    Me.Hide
End Sub

Private Sub Form_Activate()
    Dim s As String

    s = gobjMain.UsuarioAnterior
    If Len(s) = 0 Then
        txtNombre.Text = ""
        txtNombre.SetFocus
    Else
        txtNombre.Text = s
        txtClave.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoverCampo Me, KeyCode, Shift, False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ImpideSonidoEnter Me, KeyAscii
End Sub



Private Sub Form_Load()
    'lblLicenseTo.caption = ""
    lblProductName.Caption = "SiiTools"
'    lblCompanyProduct.caption = "Ishida && Asociados"
    lblPlatform.Caption = "Ishida Business Software para Windows"
    lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    lblCopyright.Caption = "Copyright"
    lblCompany.Caption = "I && A. 1998-2017"
    lblWarning.Caption = "Advertencia: Producto protegido por las " & _
            "leyes de derechos de autor y otros tratados internacionales."
End Sub

Private Sub txtClave_GotFocus()
    txtClave.SelStart = 0
    txtClave.SelLength = Len(txtClave.Text)
End Sub

Private Sub txtNombre_GotFocus()
    txtNombre.SelStart = 0
    txtNombre.SelLength = Len(txtNombre.Text)
End Sub

