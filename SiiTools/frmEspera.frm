VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEspera 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preparando Datos....."
   ClientHeight    =   1275
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   1275
   ScaleWidth      =   4365
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   372
      Left            =   3000
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   720
      Width           =   972
   End
   Begin MSComctlLib.ProgressBar pgb 
      Height          =   372
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4092
      _ExtentX        =   7223
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Caption         =   "reg."
      Height          =   252
      Left            =   1560
      TabIndex        =   4
      Top             =   840
      Width           =   372
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   252
      Left            =   960
      TabIndex        =   3
      Top             =   840
      Width           =   492
   End
   Begin VB.Label Label1 
      Caption         =   "Quedan: "
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   732
   End
End
Attribute VB_Name = "frmEspera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BandAceptado As Boolean


Private Sub cmdCancelar_Click()
    'Genera un evento que indica que se cancelo
    'RaiseEvent Cancelar
    'Unload Me
    BandAceptado = False
End Sub


Public Sub Inicio(Maximo As Integer, Minimo As Integer)
    'Maximo  tiene que ser un valor diferente de cero
    BandAceptado = True
    Me.MousePointer = vbArrow

    With pgb
        If Maximo <= 0 Then
            .max = 1
        Else
            .max = Maximo
        End If
        .Min = Minimo
        .value = Minimo
        lblTime.Caption = Maximo
    End With
    
    Me.Show 0, frmMain
    
End Sub


Public Function Estado() As Boolean
    Estado = BandAceptado
End Function

Public Sub Cancela()
    Unload Me
End Sub


Public Sub Incrementa()
    pgb.value = pgb.value + pgb.Min + 1
    lblTime.Caption = Val(lblTime.Caption) - 1
    lblTime.Refresh
    If pgb.value = pgb.max Then
         Unload Me
    End If
End Sub





