VERSION 5.00
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Begin VB.Form frmPCBusqueda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Búsqueda de Prov/Cli"
   ClientHeight    =   3165
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4395
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraCliente 
      Caption         =   "Filtro"
      Height          =   975
      Left            =   60
      TabIndex        =   9
      Top             =   1080
      Width           =   4215
      Begin VB.TextBox lblFiltroCli 
         BackColor       =   &H80000018&
         Height          =   600
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmdGenCliente 
         Caption         =   "&Condiciones de Búsqueda"
         Height          =   492
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1092
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   372
      Left            =   2280
      TabIndex        =   8
      Top             =   2580
      Width           =   1452
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar -F5"
      Height          =   372
      Left            =   540
      TabIndex        =   7
      Top             =   2580
      Width           =   1452
   End
   Begin VB.TextBox txtCodigo 
      Height          =   372
      Left            =   1080
      MaxLength       =   20
      TabIndex        =   1
      Top             =   120
      Width           =   3195
   End
   Begin VB.TextBox txtDesc 
      Height          =   372
      Left            =   1080
      MaxLength       =   50
      TabIndex        =   3
      Top             =   600
      Width           =   3195
   End
   Begin VB.ComboBox cboGrupo 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   1212
   End
   Begin FlexComboProy.FlexCombo fcbGrupo 
      Height          =   348
      Left            =   2400
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   1692
      _ExtentX        =   2990
      _ExtentY        =   609
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
   Begin FlexComboProy.FlexCombo fcbvendedor 
      Height          =   360
      Left            =   1080
      TabIndex        =   12
      Top             =   2100
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   635
      DispCol         =   1
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
   Begin VB.Label lblEtiqGrupo1 
      Caption         =   "Vendedor"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2205
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Código "
      Height          =   192
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   564
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Nombre"
      Height          =   192
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   588
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "&Grupo "
      Height          =   192
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmPCBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private BandAceptado As Boolean
Private BandCliente As Boolean

Public Function Inicio(ByRef CodProvCli As String, _
                       ByRef Desc As String, _
                       ByRef CodGrupo As String, _
                       ByRef numGrupo As Integer) As Boolean
    Dim antes As String, i As Integer
    On Error GoTo errtrap
    
    'Cambia forma de cursor mientras se carga
    MensajeStatus MSG_PREPARA, vbHourglass
    
    'Prepara ComboBox de etiquetas de grupo
    cboGrupo.Clear
    For i = 1 To PCGRUPO_MAX
        cboGrupo.AddItem gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupo(i)
    Next i
    
    If (numGrupo <= cboGrupo.ListCount) And (numGrupo > 0) Then
        cboGrupo.ListIndex = numGrupo - 1   'Selecciona lo anterior
    ElseIf cboGrupo.ListCount > 0 Then
        cboGrupo.ListIndex = 0              'Selecciona la primera
    End If
    fcbGrupo.KeyText = CodGrupo     'Recupera la selección anterior
    
    fcbvendedor.SetData gobjMain.EmpresaActual.ListaFCVendedorJefe(True, False)
    
    MensajeStatus
    BandAceptado = False
    Me.Show vbModal, frmMain
    
    'Si aplastó el botón 'Aceptar'
    If BandAceptado Then
        'Devuelve los valores de condición para a búsqueda
        CodProvCli = Trim$(txtCodigo.Text)
        Desc = Trim$(txtDesc.Text)
        
        If cboGrupo.ListIndex >= 0 Then
            numGrupo = cboGrupo.ListIndex + 1
            CodGrupo = Trim$(fcbGrupo.KeyText)
        End If
        gobjMain.objCondicion.Usuario1 = fcbvendedor.KeyText
    End If
    
    'Devuelve true/false
    Inicio = BandAceptado
    
    Exit Function
errtrap:
    MensajeStatus
    DispErr
    Exit Function
End Function

Private Sub cboGrupo_Click()
    Dim Numg As Integer
    On Error GoTo errtrap
    If cboGrupo.ListIndex < 0 Then Exit Sub
    
    MensajeStatus MSG_PREPARA, vbHourglass
    
    Numg = cboGrupo.ListIndex + 1
    If InStr(1, UCase(Me.Caption), "CLIENTE") Then
        fcbGrupo.SetData gobjMain.EmpresaActual.ListaPCGrupoOrigen(Numg, False, False, 2)
    ElseIf InStr(1, UCase(Me.Caption), "PROVEEDOR") Then
        fcbGrupo.SetData gobjMain.EmpresaActual.ListaPCGrupoOrigen(Numg, False, False, 1)
    ElseIf InStr(1, UCase(Me.Caption), "EMPLEA") Then
        fcbGrupo.SetData gobjMain.EmpresaActual.ListaPCGrupoOrigen(Numg, False, False, 4)
    
    Else
        fcbGrupo.SetData gobjMain.EmpresaActual.ListaPCGrupo(Numg, False, False)
    End If
    fcbGrupo.KeyText = ""
    MensajeStatus
    Exit Sub
errtrap:
    MensajeStatus
    DispErr
    Exit Sub
End Sub

Private Sub cmdAceptar_Click()
    BandAceptado = True
    txtCodigo.SetFocus
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    BandAceptado = False
    txtCodigo.SetFocus
    Me.Hide
End Sub

Private Sub cmdGenCliente_Click()
    Dim frmCli As frmB_FiltroxCliente, EtiqCliente As String
    Set frmCli = New frmB_FiltroxCliente
    EtiqCliente = frmCli.Inicio(Me.tag, BandCliente)
    If UCase(EtiqCliente) <> "CANCELAR" Then
        lblFiltroCli.Text = EtiqCliente
    End If
End Sub

Private Sub Form_Activate()
    Dim c As Control, band As Boolean, c2 As Control
    On Error Resume Next
    If Not Me.Visible Then Exit Sub
    
    'Busca un TextBox que tenga alguna cadena
    Set c2 = txtCodigo
    For Each c In Me.Controls
        If TypeName(c) = "TextBox" Then
            If Len(c.Text) > 0 Then 'Si encuentra,
                If (c.TabIndex < c2.TabIndex) _
                    Or (Len(c2.Text) = 0) Then Set c2 = c
            End If
        End If
    Next c
    
    c2.SetFocus
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    ImpideSonidoEnter Me, KeyAscii
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Me.Hide         'Se pone esto para evitar el posible BUG de Windows98
End Sub

Private Sub txtCodigo_GotFocus()
    txtCodigo.SelStart = 0
    txtCodigo.SelLength = Len(txtCodigo.Text)
End Sub

Private Sub txtDesc_GotFocus()
    txtDesc.SelStart = 0
    txtDesc.SelLength = Len(txtDesc.Text)
End Sub


Public Function InicioGarante(ByRef CodProvCli As String, _
                       ByRef Desc As String, _
                       ByRef CodGrupo As String, _
                       ByRef numGrupo As Integer) As Boolean
    Dim antes As String, i As Integer
    On Error GoTo errtrap
    
    'Cambia forma de cursor mientras se carga
    MensajeStatus MSG_PREPARA, vbHourglass
    
    'Prepara ComboBox de etiquetas de grupo
    cboGrupo.Clear
    For i = 1 To PCGRUPO_MAX
        cboGrupo.AddItem gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupoG(i)
    Next i
    
    If (numGrupo <= cboGrupo.ListCount) And (numGrupo > 0) Then
        cboGrupo.ListIndex = numGrupo - 1   'Selecciona lo anterior
    ElseIf cboGrupo.ListCount > 0 Then
        cboGrupo.ListIndex = 0              'Selecciona la primera
    End If
    fcbGrupo.KeyText = CodGrupo     'Recupera la selección anterior
    
    MensajeStatus
    BandAceptado = False
    Me.Show vbModal, frmMain
    
    'Si aplastó el botón 'Aceptar'
    If BandAceptado Then
        'Devuelve los valores de condición para a búsqueda
        CodProvCli = Trim$(txtCodigo.Text)
        Desc = Trim$(txtDesc.Text)
        
        If cboGrupo.ListIndex >= 0 Then
            numGrupo = cboGrupo.ListIndex + 1
            CodGrupo = Trim$(fcbGrupo.KeyText)
        End If
    End If
    
    'Devuelve true/false
    InicioGarante = BandAceptado
    
    Exit Function
errtrap:
    MensajeStatus
    DispErr
    Exit Function
End Function


