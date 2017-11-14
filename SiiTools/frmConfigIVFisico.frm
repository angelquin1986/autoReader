VERSION 5.00
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Begin VB.Form frmConfigIVFisico 
   Caption         =   "Configuración"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3930
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   3930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transacciones"
      Height          =   2055
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   3735
      Begin VB.CheckBox chkBandTotalizarItem 
         Caption         =   "Totalizar Items"
         Height          =   195
         Left            =   720
         TabIndex        =   2
         Top             =   960
         Width           =   2415
      End
      Begin VB.CheckBox chkBandLineaAuto 
         Caption         =   "Paso Automático de Línea"
         Height          =   195
         Left            =   720
         TabIndex        =   1
         Top             =   720
         Width           =   2415
      End
      Begin FlexComboProy.FlexCombo fcbTrans_BJ 
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
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
      Begin FlexComboProy.FlexCombo fcbTrans_AJ 
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Top             =   1320
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
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
      Begin FlexComboProy.FlexCombo fcbTrans_CF 
         Height          =   255
         Left            =   1920
         TabIndex        =   0
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Baja de Bodega"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ajuste de Bodega"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Constatación Física"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1410
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      Caption         =   "Parámetros para la constatación física de Inventario"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3690
   End
End
Attribute VB_Name = "frmConfigIVFisico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub Inicio()
    CargarDatos
    Me.Show vbModal
End Sub

Private Sub CargarDatos()
    RecuperarConfigIVFisico
    fcbTrans_CF.SetData gobjMain.EmpresaActual.ListaGNTrans("IV", False, False)
    fcbTrans_AJ.SetData gobjMain.EmpresaActual.ListaGNTrans("IV", False, False)
    fcbTrans_BJ.SetData gobjMain.EmpresaActual.ListaGNTrans("IV", False, False)
    
    fcbTrans_CF.KeyText = gConfigIVFisico.CodTrans_CF
    fcbTrans_AJ.KeyText = gConfigIVFisico.CodTrans_AJ
    fcbTrans_BJ.KeyText = gConfigIVFisico.CodTrans_BJ
    chkBandLineaAuto.value = IIf(gConfigIVFisico.BandLineaAuto, vbChecked, vbUnchecked)
    'jeaa 13/10/04
    chkBandTotalizarItem.value = IIf(gConfigIVFisico.BandTotalizarItem, vbChecked, vbUnchecked)
End Sub

Private Sub cmdAceptar_Click()
    If Me.tag <> "AjusteAutomatico" Then
        With gConfigIVFisico
            .CodTrans_CF = fcbTrans_CF.KeyText
            .CodTrans_AJ = fcbTrans_AJ.KeyText
            .CodTrans_BJ = fcbTrans_BJ.KeyText
            .BandLineaAuto = (chkBandLineaAuto.value = vbChecked)
            'jeaa 13/10/04
            .BandTotalizarItem = (chkBandTotalizarItem.value = vbChecked)
        End With
        GrabarConfigIVFisico
    Else
        With gConfigIVAjusteAutomatico
            .CodTrans_AA = fcbTrans_CF.KeyText
            .CodTrans_AAJ = fcbTrans_AJ.KeyText
            .CodTrans_ABJ = fcbTrans_BJ.KeyText
            .BandLineaAutoA = (chkBandLineaAuto.value = vbChecked)
            'jeaa 13/10/04
            .BandTotalizarItemA = (chkBandTotalizarItem.value = vbChecked)
        End With
        GrabarConfigIVAjusteAutomatico
    
    End If
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        Unload Me
        KeyCode = 0
    Case Else
        MoverCampo Me, KeyCode, Shift, True
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ImpideSonidoEnter Me, KeyAscii
End Sub

Public Sub InicioAjusteAutomatico(ByVal tag As String)
    Me.tag = tag
    Label1.Caption = "Parámetros para Ajustes Automáticos"
    CargarDatosAjustesAutomaticos
    Me.Show vbModal
End Sub

Private Sub CargarDatosAjustesAutomaticos()
    RecuperarConfigIVAjusteAutomatico
    fcbTrans_CF.SetData gobjMain.EmpresaActual.ListaGNTrans("IV", False, False)
    fcbTrans_AJ.SetData gobjMain.EmpresaActual.ListaGNTrans("IV", False, False)
    fcbTrans_BJ.SetData gobjMain.EmpresaActual.ListaGNTrans("IV", False, False)
    
    fcbTrans_CF.KeyText = gConfigIVAjusteAutomatico.CodTrans_AA
    fcbTrans_AJ.KeyText = gConfigIVAjusteAutomatico.CodTrans_AAJ
    fcbTrans_BJ.KeyText = gConfigIVAjusteAutomatico.CodTrans_ABJ
    chkBandLineaAuto.value = IIf(gConfigIVAjusteAutomatico.BandLineaAutoA, vbChecked, vbUnchecked)
    'jeaa 13/10/04
    chkBandTotalizarItem.value = IIf(gConfigIVAjusteAutomatico.BandTotalizarItemA, vbChecked, vbUnchecked)
End Sub

