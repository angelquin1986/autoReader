VERSION 5.00
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Begin VB.Form frmB_PlanMant 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda"
   ClientHeight    =   1440
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5355
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   5355
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraEmpresas 
      Caption         =   "Empresa"
      Height          =   2655
      Left            =   5400
      TabIndex        =   2
      Top             =   60
      Visible         =   0   'False
      Width           =   2835
      Begin VB.ListBox lstBase 
         Height          =   2310
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   400
      Left            =   2340
      TabIndex        =   1
      Top             =   855
      Width           =   1200
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar -F5"
      Height          =   400
      Left            =   1035
      TabIndex        =   0
      Top             =   855
      Width           =   1200
   End
   Begin FlexComboProy.FlexCombo fcbPlan 
      Height          =   360
      Left            =   2475
      TabIndex        =   4
      Top             =   270
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   635
      DispCol         =   1
      ColWidth1       =   2400
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblPlan 
      AutoSize        =   -1  'True
      Caption         =   "Plan Mantenimiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   540
      TabIndex        =   5
      Top             =   315
      Width           =   1680
   End
End
Attribute VB_Name = "frmB_PlanMant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BandAceptado As Boolean


Private Sub cmdAceptar_Click()
    If Len(fcbPlan.KeyText) = 0 Then
        MsgBox "Debe Escoger un Plan": fcbPlan.SetFocus
        Exit Sub
    End If
    BandAceptado = True
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    BandAceptado = False
    Me.Hide
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

Public Function Inicio(ByRef objcond As Condicion) As Boolean
       
    With objcond
        fcbPlan.SetData gobjMain.EmpresaActual.ListaIVPlan(False, True)
        BandAceptado = False
        'KeyTrans = "TCompra_Trans"
        'RecuperaSelecTrans
        Me.Show vbModal, frmMain
        'Si aplastó el botón 'Aceptar'
        If BandAceptado Then
            'Devuelve los valores de condición para la búsqueda
            .CodBanco1 = fcbPlan.KeyText
        End If
    End With
    'Devuelve true/false
    Unload Me
    Inicio = BandAceptado
End Function


Private Function PreparaCadenaIN(lst As ListBox) As String
    Dim Cadena As String, i As Integer
    Cadena = ""
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            If Cadena = "" Then
                Cadena = "'" & Left(lst.List(i), lst.ItemData(i)) & "',"
            Else
                Cadena = Cadena & "'" & _
                              Left(lst.List(i), lst.ItemData(i)) & "',"
            End If
        End If
    Next i
    PreparaCadenaIN = Mid$(Cadena, 1, Len(Cadena) - 1)
End Function



