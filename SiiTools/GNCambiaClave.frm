VERSION 5.00
Begin VB.Form frmCambiaClave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos"
   ClientHeight    =   2040
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   5145
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
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtClave 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1260
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1140
      Width           =   1695
   End
   Begin VB.TextBox txtClave 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1260
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   780
      Width           =   1695
   End
   Begin VB.TextBox txtCodigo 
      Enabled         =   0   'False
      Height          =   360
      Left            =   1260
      MaxLength       =   10
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   " "
      Top             =   60
      Width           =   1695
   End
   Begin VB.TextBox txtNombre 
      Enabled         =   0   'False
      Height          =   360
      Left            =   1260
      MaxLength       =   50
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   420
      Width           =   3735
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar -F9"
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
      Left            =   1080
      TabIndex        =   2
      Top             =   1620
      Width           =   1452
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
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
      Left            =   2640
      TabIndex        =   3
      Top             =   1620
      Width           =   1452
   End
   Begin VB.Label lblEtiqClave 
      AutoSize        =   -1  'True
      Caption         =   "Repetir C&lave"
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
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   960
   End
   Begin VB.Label lblEtiqClave 
      AutoSize        =   -1  'True
      Caption         =   "C&lave"
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
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Nombre"
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
      TabIndex        =   5
      Top             =   480
      Width           =   945
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "&Código"
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
      TabIndex        =   4
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "frmCambiaClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private tabla As String
Private mbooIniciando As Boolean

Private mobjUsuario As usuario
Private mobjUsuarioAnt As usuario
Private ClaveNew As String

Public Sub Visualiza()
    Dim codMod As String, i As Integer
    mbooIniciando = True
    'Visualiza
    Select Case tabla
    Case "Usuario"
        txtCodigo = mobjUsuario.codUsuario
        txtNombre = mobjUsuario.NombreUsuario
    End Select
    mbooIniciando = False
End Sub

Public Function Modificar(cod As String, NombreTabla As String, Caption As String) As String
    On Error GoTo ErrTrap
    tabla = NombreTabla
    
    Me.Caption = "Datos de " & Caption
    
    'Se bloquea control  "txtCodigo"
    txtCodigo.Enabled = False
    txtNombre.Enabled = False
    'txtClave(0).SetFocus
    'Visualiza
    Select Case tabla
    Case "Usuario"
        Set mobjUsuario = gobjMain.RecuperaUsuario(cod)
    End Select
    Visualiza
    
    Me.Show vbModal, frmMain
    Modificar = ClaveNew
    Exit Function
ErrTrap:
    DispErr
    Unload Me
    Exit Function
End Function

Public Sub Copiar(cod As String, NombreTabla As String, Caption As String)
End Sub



Public Sub Eliminar(cod As String, NombreTabla As String, Caption As String)
End Sub


Private Sub cmdAceptar_Click()
    If Grabar Then Unload Me
End Sub

Private Function Grabar() As Boolean
    On Error GoTo ErrTrap

    If txtClave(0).Text <> txtClave(1).Text Then
        MsgBox "Las claves no son iguales.", vbInformation
        txtClave(0).SetFocus
        Exit Function
        
    End If
    ClaveNew = mobjUsuario.Clave
    mobjUsuario.BandCambiaClave = False
    
    'Graba
    MensajeStatus MSG_GRABANDO, vbHourglass '***Agregado 11/abr/02 Angel
    Select Case tabla
    Case "Usuario"
        mobjUsuario.GrabarCambioClave
    End Select
    MensajeStatus                           '***Agregado 11/abr/02 Angel
    Grabar = True
    Exit Function
ErrTrap:
    MensajeStatus                           '***Agregado 11/abr/02 Angel
    If Err.Number = 3022 Or Err.Number = ERR_REPITECODIGO Then
        MsgBox "El código ya existe. Por favor utilice otro código.", vbExclamation
    Else
        DispErr
    End If
    Exit Function
End Function

Private Sub cmdCancelar_Click()
    Unload Me
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF9
        cmdAceptar_Click
        KeyCode = 0
    Case Else
        MoverCampo Me, KeyCode, Shift, True
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ImpideSonidoEnter Me, KeyAscii
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim rt As Integer, band As Boolean
    
    Select Case tabla
    Case "Usuario"
        If Not (mobjUsuario Is Nothing) Then band = mobjUsuario.Modificado
    Case Else
        Debug.Assert True
    End Select
    
    If band Then
        rt = MsgBox(MSG_CANCELMOD, vbYesNoCancel)
        Select Case rt
        Case vbYes           'Graba y cierra
            If Grabar Then
                Me.Hide
            Else
                Cancel = 1    'Si ocurre error al grabar,no cierra
            End If
        Case vbNo          'Cierra sin grabar
            Me.Hide
        Case vbCancel
            Cancel = 1      'No se cierra la ventana
        End Select
    End If
End Sub

Private Sub Form_Terminate()
    Set mobjUsuario = Nothing
End Sub



Private Sub grd_BeforeEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
If col = 1 Then Cancel = True
End Sub


Private Sub txtClave_Change(Index As Integer)
    On Error GoTo ErrTrap
    If mbooIniciando Then Exit Sub
    Select Case tabla
    Case "Usuario"
        mobjUsuario.Clave = txtClave(0).Text
    End Select
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

