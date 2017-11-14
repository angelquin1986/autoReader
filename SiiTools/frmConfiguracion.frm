VERSION 5.00
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "flexcombo.ocx"
Begin VB.Form frmConfiguracion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   2475
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   6270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDiferencial 
      Caption         =   "Abrir archivo de para importacion en forma diferencial"
      Height          =   390
      Left            =   165
      TabIndex        =   14
      Top             =   1440
      Width           =   4500
   End
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   6270
      TabIndex        =   11
      Top             =   1860
      Width           =   6270
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   3360
         TabIndex        =   13
         Top             =   120
         Width           =   1200
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   120
         Width           =   1200
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   6270
      TabIndex        =   10
      Top             =   1860
      Width           =   6270
   End
   Begin FlexComboProy.FlexCombo fcbCliente 
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   960
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      ColWidth1       =   2400
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FlexComboProy.FlexCombo fcbResp 
      Height          =   330
      Left            =   4440
      TabIndex        =   1
      ToolTipText     =   "Responsable de la transacción"
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ColWidth2       =   1200
      ColWidth3       =   1200
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FlexComboProy.FlexCombo fcbTrans 
      Height          =   330
      Left            =   1080
      TabIndex        =   0
      ToolTipText     =   "Responsable de la transacción"
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ColWidth2       =   1200
      ColWidth3       =   1200
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FlexComboProy.FlexCombo fcbForma 
      Height          =   330
      Left            =   1080
      TabIndex        =   2
      ToolTipText     =   "Responsable de la transacción"
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ColWidth1       =   2400
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FlexComboProy.FlexCombo fcbMoneda 
      Height          =   330
      Left            =   4440
      TabIndex        =   3
      ToolTipText     =   "Responsable de la transacción"
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ColWidth2       =   1200
      ColWidth3       =   1200
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moneda"
      Height          =   195
      Left            =   3240
      TabIndex        =   9
      Top             =   600
      Width           =   600
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Responsable  "
      Height          =   195
      Left            =   3270
      TabIndex        =   8
      Top             =   120
      Width           =   1050
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cod.Trans  "
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   825
   End
   Begin VB.Label lblforma 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forma  "
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
End
Attribute VB_Name = "frmConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Inicio()
    'carga listado de Clientes
    MensajeStatus "Recuperando Configuraciones...", vbHourglass
    CargaCliente
    CargaTrans
    CargaMoneda
    CargaResponsable
    CargaForma
    RecuperaSeleccion
    MensajeStatus "", vbNormal
    Me.Show vbModal
End Sub


Private Sub RecuperaSeleccion()
    With gConfig
        fcbCliente.KeyText = .CodCli
        fcbMoneda.KeyText = .Moneda
        fcbResp.KeyText = .Responsable
        fcbTrans.KeyText = .CodTrans
        fcbForma.KeyText = .FormaCobroPago
        chkDiferencial.value = IIf(.AbrirArchivoenFormaDiferencial, vbChecked, vbUnchecked)
    End With

End Sub

Private Sub CargaForma()
    fcbForma.SetData gobjMain.EmpresaActual.ListaTSFormaCobroPago(True, True, False)
End Sub
Private Sub CargaCliente()
    'gobjMain.EmpresaActual.ListaPCProvCli
    fcbCliente.SetData gobjMain.EmpresaActual.ListaPCProvCli(False, True, False)
End Sub

 Private Sub CargaTrans()
    fcbTrans.SetData gobjMain.EmpresaActual.ListaGNTrans("IV", True, False)
 End Sub

Private Sub CargaMoneda()
    fcbMoneda.SetData gobjMain.EmpresaActual.ListaGNMoneda
End Sub

Private Sub CargaResponsable()
    fcbResp.SetData gobjMain.EmpresaActual.ListaGNResponsable(False)
End Sub

Private Sub cmdAceptar_Click()
    'Graba en el registro de Windows
    With gConfig
        .CodCli = fcbCliente.Text
        .FormaCobroPago = fcbForma.Text
        .Moneda = fcbMoneda.Text
        .CodTrans = fcbTrans.Text
        .Responsable = fcbResp.Text
        .AbrirArchivoenFormaDiferencial = IIf(chkDiferencial.value = vbChecked, True, False)
    End With
    GuardaConfig
    RecuperaConfig
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case Else
        MoverCampo Me, KeyCode, Shift, False
    End Select
End Sub

