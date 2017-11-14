VERSION 5.00
Begin VB.Form frmMiMsgBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SiiTools"
   ClientHeight    =   3060
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   372
      Left            =   3720
      TabIndex        =   2
      Top             =   2040
      Width           =   1092
   End
   Begin VB.TextBox txtMsg 
      Height          =   1692
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   5172
   End
   Begin VB.CommandButton cmdNoTodo 
      Caption         =   "No a &todo"
      Height          =   372
      Left            =   1800
      TabIndex        =   1
      Top             =   2520
      Width           =   1092
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "&No"
      Height          =   372
      Left            =   1800
      TabIndex        =   0
      Top             =   2040
      Width           =   1092
   End
   Begin VB.CommandButton cmdSiiTodo 
      Caption         =   "Sí &a todo"
      Height          =   372
      Left            =   600
      TabIndex        =   5
      Top             =   2520
      Width           =   1092
   End
   Begin VB.CommandButton cmdSi 
      Caption         =   "&Sí"
      Height          =   372
      Left            =   600
      TabIndex        =   4
      Top             =   2040
      Width           =   1092
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5280
      Y1              =   1920
      Y2              =   1920
   End
End
Attribute VB_Name = "frmMiMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Enum E_MiMsgBox
    mmsgSi = 0
    mmsgSiTodo = 1
    mmsgNo = 2
    mmsgNoTodo = 3
    mmsgCancelar = 4
End Enum

Private mRes As E_MiMsgBox


Public Function MiMsgBox( _
                    ByVal msg As String, _
                    Optional ByVal titulo As String) As E_MiMsgBox
    If Len(titulo) > 0 Then Me.Caption = titulo
    mRes = mmsgCancelar
    
    'Convierte de vbCr --> vbCrLf
    msg = Replace(msg, vbCrLf, vbCr)
    msg = Replace(msg, vbCr, vbCrLf)
    
    txtMsg.Text = msg
    Me.Show vbModal, frmMain
    
    MiMsgBox = mRes
    
    Unload Me
End Function

Private Sub cmdCancelar_Click()
    mRes = mmsgCancelar
    Me.Hide
End Sub

Private Sub cmdNo_Click()
    mRes = mmsgNo
    Me.Hide
End Sub

Private Sub cmdNoTodo_Click()
    mRes = mmsgNoTodo
    Me.Hide
End Sub

Private Sub cmdSi_Click()
    mRes = mmsgSi
    Me.Hide
End Sub

Private Sub cmdSiiTodo_Click()
    mRes = mmsgSiTodo
    Me.Hide
End Sub

