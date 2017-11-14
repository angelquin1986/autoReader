VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4575
   ClientLeft      =   30
   ClientTop       =   30
   ClientWidth     =   7860
   ControlBox      =   0   'False
   Icon            =   "Splash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      BackColor       =   &H00FFFFFF&
      Height          =   4530
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7740
      Begin VB.Label lblPlatform 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sistema Informático Integrado para Windows"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   705
         Left            =   2460
         TabIndex        =   8
         Tag             =   "Plataforma"
         Top             =   1320
         Width           =   3765
      End
      Begin VB.Label lblProductName 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "SiiTools"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   3240
         TabIndex        =   7
         Tag             =   "Producto"
         Top             =   2100
         Width           =   2310
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "LicenciaA"
         Height          =   255
         Left            =   2430
         TabIndex        =   1
         Tag             =   "LicenciaA"
         Top             =   300
         Width           =   915
      End
      Begin VB.Image Image1 
         Height          =   1815
         Left            =   120
         Picture         =   "Splash.frx":0442
         Top             =   180
         Width           =   1980
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Versión 4"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   240
         Left            =   720
         TabIndex        =   5
         Tag             =   "Versión"
         Top             =   1980
         Width           =   765
      End
      Begin VB.Label lblWarning 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Advertencia"
         Height          =   555
         Left            =   0
         TabIndex        =   2
         Tag             =   "Advertencia"
         Top             =   3000
         Width           =   4155
      End
      Begin VB.Label lblCompany 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Organización"
         BeginProperty Font 
            Name            =   "Arial Narrow"
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
         TabIndex        =   4
         Tag             =   "Organización"
         Top             =   2220
         Width           =   2115
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Copyright"
         Height          =   255
         Left            =   5280
         TabIndex        =   3
         Tag             =   "Copyright"
         Top             =   3360
         Width           =   2415
      End
      Begin VB.Image Image2 
         Height          =   4365
         Left            =   0
         Picture         =   "Splash.frx":369F
         Top             =   120
         Width           =   7740
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Ishida && Asociados"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   2508
         TabIndex        =   6
         Tag             =   "ProductoOrganización"
         Top             =   768
         Width           =   2244
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Aceptado As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)
            Aceptado = True
            Me.Hide
End Sub

Public Function Inicio() As Boolean
    Aceptado = False
    Me.Show vbModal, frmMain
    Inicio = Aceptado
    Unload Me
End Function


Private Sub Form_Load()
    lblLicenseTo.Caption = ""
    lblProductName.Caption = "SiiTools"
    lblCompanyProduct.Caption = "Ishida && Asociados"
    lblPlatform.Caption = "Sistema Informático Integrado para Windows"
    lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    lblCopyright.Caption = "Copyright"
    lblCompany.Caption = "I && A. 1998-2013"
    lblWarning.Caption = "Advertencia: Este producto está protegido por las " & _
            "leyes de derechos de autor y otros tratados internacionales."
End Sub



