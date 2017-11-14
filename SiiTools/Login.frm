VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inicio de sesión"
   ClientHeight    =   1515
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   3075
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   3075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Inicio de sesión"
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   1739
      TabIndex        =   5
      Tag             =   "Cancelar"
      Top             =   1020
      Width           =   1020
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   360
      Left            =   314
      TabIndex        =   4
      Tag             =   "Aceptar"
      Top             =   1020
      Width           =   1020
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
      Left            =   948
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   1608
   End
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
      Left            =   948
      TabIndex        =   1
      Top             =   135
      Width           =   1608
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Clave"
      Height          =   252
      Index           =   1
      Left            =   228
      TabIndex        =   2
      Tag             =   "&Contraseña:"
      Top             =   540
      Width           =   720
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Usuario"
      Height          =   252
      Index           =   0
      Left            =   228
      TabIndex        =   0
      Tag             =   "Usuari&o:"
      Top             =   156
      Width           =   720
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Aceptado As Boolean

Public Function Inicio() As Boolean
    Aceptado = False
    Me.Show vbModal, frmMain
    
    Inicio = Aceptado
    Unload Me
End Function


Private Sub cmdAceptar_Click()
    Dim op As usuario
    Dim usClave As usuario
    Dim BandModulo As Boolean
    On Error GoTo ErrTrap
    
    Set op = gobjMain.Login(Trim$(txtNombre), Trim$(txtClave))
    If Not op Is Nothing Then
            'jeaa 23/09/2008
                Set usClave = op
                If usClave.BandCambiaClave Then
                    txtClave.Text = frmCambiaClave.Modificar(usClave.codUsuario, "Usuario", "Cambio de clave primera vez")
                     
                End If
                Set usClave = Nothing
            
            If Not op.BandValida Then
               MsgBox "El usuario " & op.NombreUsuario & " no esta Activo" & Chr(13) & _
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
        '-----------------
        Aceptado = True
        Me.Hide
    Else
        MsgBox "Nombre o clave está incorrecto. Intente de nuevo."
        txtNombre.SetFocus
    End If
    Set op = Nothing
    Exit Sub
ErrTrap:
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




Private Sub txtClave_GotFocus()
    txtClave.SelStart = 0
    txtClave.SelLength = Len(txtClave.Text)
End Sub

Private Sub txtNombre_GotFocus()
    txtNombre.SelStart = 0
    txtNombre.SelLength = Len(txtNombre.Text)
End Sub
