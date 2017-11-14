VERSION 5.00
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAFBusqueda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Búsqueda de items"
   ClientHeight    =   3090
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   4530
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab sst 
      Height          =   2535
      Left            =   60
      TabIndex        =   16
      Top             =   60
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   4471
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Busqueda"
      TabPicture(0)   =   "AFBusqueda.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraTrans"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Filtros"
      TabPicture(1)   =   "AFBusqueda.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fcbGrupo1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fcbGrupo2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fcbGrupo3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fcbGrupo4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "fcbGrupo5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblEtiqGrupo1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblEtiqGrupo5"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblEtiqGrupo4"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblEtiqGrupo3"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lblEtiqGrupo2"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      Begin FlexComboProy.FlexCombo fcbGrupo1 
         Height          =   360
         Left            =   -73800
         TabIndex        =   9
         Top             =   540
         Width           =   3000
         _ExtentX        =   5292
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
      Begin FlexComboProy.FlexCombo fcbGrupo2 
         Height          =   360
         Left            =   -73800
         TabIndex        =   10
         Top             =   900
         Width           =   3000
         _ExtentX        =   5292
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
      Begin FlexComboProy.FlexCombo fcbGrupo3 
         Height          =   360
         Left            =   -73800
         TabIndex        =   11
         Top             =   1260
         Width           =   3000
         _ExtentX        =   5292
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
      Begin FlexComboProy.FlexCombo fcbGrupo4 
         Height          =   360
         Left            =   -73800
         TabIndex        =   12
         Top             =   1620
         Width           =   3015
         _ExtentX        =   5318
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
      Begin FlexComboProy.FlexCombo fcbGrupo5 
         Height          =   360
         Left            =   -73800
         TabIndex        =   13
         Top             =   1980
         Visible         =   0   'False
         Width           =   3000
         _ExtentX        =   5292
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
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   120
         TabIndex        =   0
         Top             =   420
         Width           =   4155
         Begin VB.TextBox txtCodAlt 
            Height          =   372
            Left            =   1020
            MaxLength       =   20
            TabIndex        =   2
            Top             =   660
            Width           =   3012
         End
         Begin VB.TextBox txtDesc 
            Height          =   372
            Left            =   1020
            MaxLength       =   50
            TabIndex        =   3
            Top             =   1080
            Width           =   3012
         End
         Begin VB.TextBox txtCodigo 
            Height          =   372
            Left            =   1020
            MaxLength       =   20
            TabIndex        =   1
            Top             =   240
            Width           =   3012
         End
         Begin VB.CheckBox chkIVA 
            Caption         =   "Solo Items IVA diferente de cero"
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
            Left            =   1050
            TabIndex        =   4
            Top             =   1500
            Width           =   2772
         End
         Begin FlexComboProy.FlexCombo FcbBodega 
            Height          =   375
            Left            =   1020
            TabIndex        =   5
            Top             =   1500
            Visible         =   0   'False
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
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
            Caption         =   "Cód. &Alterno "
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
            Left            =   60
            TabIndex        =   26
            Top             =   720
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "&Descripción "
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
            Left            =   60
            TabIndex        =   25
            Top             =   1200
            Width           =   900
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "&Código "
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
            Left            =   60
            TabIndex        =   24
            Top             =   240
            Width           =   570
         End
         Begin VB.Label lblBodega 
            AutoSize        =   -1  'True
            Caption         =   "&Bodega "
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
            Left            =   60
            TabIndex        =   28
            Top             =   1560
            Visible         =   0   'False
            Width           =   600
         End
      End
      Begin VB.Frame FraTrans 
         Caption         =   "Transacciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   120
         TabIndex        =   22
         Top             =   420
         Visible         =   0   'False
         Width           =   4155
         Begin VB.Frame fraNumTrans 
            Caption         =   "# T&rans. (desde - hasta)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   600
            TabIndex        =   27
            Top             =   240
            Width           =   3195
            Begin VB.TextBox txtNumTrans2 
               Alignment       =   1  'Right Justify
               Height          =   360
               Left            =   1800
               TabIndex        =   7
               Top             =   240
               Width           =   1212
            End
            Begin VB.TextBox txtNumTrans1 
               Alignment       =   1  'Right Justify
               Height          =   360
               Left            =   300
               TabIndex        =   6
               Top             =   240
               Width           =   1212
            End
         End
         Begin VB.Frame fraCodTrans 
            Caption         =   "Cod.&Trans."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   1260
            TabIndex        =   23
            Top             =   1020
            Width           =   1932
            Begin FlexComboProy.FlexCombo fcbTrans 
               Height          =   315
               Left            =   165
               TabIndex        =   8
               Top             =   240
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
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
         End
      End
      Begin VB.Label lblEtiqGrupo1 
         Caption         =   "Grupo1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   21
         Top             =   645
         Width           =   1035
      End
      Begin VB.Label lblEtiqGrupo5 
         Caption         =   "Grupo5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   20
         Top             =   2085
         Width           =   1035
      End
      Begin VB.Label lblEtiqGrupo4 
         Caption         =   "Grupo4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   19
         Top             =   1725
         Width           =   1035
      End
      Begin VB.Label lblEtiqGrupo3 
         Caption         =   "Grupo3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   18
         Top             =   1365
         Width           =   1035
      End
      Begin VB.Label lblEtiqGrupo2 
         Caption         =   "Grupo2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   17
         Top             =   1005
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar -F5"
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
      Left            =   600
      TabIndex        =   14
      Top             =   2640
      Width           =   1452
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
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
      Left            =   2220
      TabIndex        =   15
      Top             =   2640
      Width           =   1452
   End
End
Attribute VB_Name = "frmAFBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private BandAceptado As Boolean




Public Function Inicio(ByRef coditem As String, _
                       ByRef CodAlt As String, _
                       ByRef Desc As String, _
                       ByRef CodGrupo1 As String, _
                       ByRef CodGrupo2 As String, _
                       ByRef CodGrupo3 As String, _
                       ByRef CodGrupo4 As String, _
                       ByRef CodGrupo5 As String, _
                       ByRef numGrupo As Integer, _
                       ByRef bandIVA As Boolean, _
                       ByRef tag As String, _
                        Optional ByRef CodBodega As String) As Boolean
    Dim antes As String, i As Integer
    On Error GoTo ErrTrap
    
    Me.tag = tag
    'Cambia forma de cursor mientras se carga
    MensajeStatus MSG_PREPARA, vbHourglass
    sst.Tab = 0
    lblEtiqGrupo1 = gobjMain.EmpresaActual.GNOpcion.EtiqAFGrupo(1)
    lblEtiqGrupo2 = gobjMain.EmpresaActual.GNOpcion.EtiqAFGrupo(2)
    lblEtiqGrupo3 = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(3)
    lblEtiqGrupo4 = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(4)
    lblEtiqGrupo5 = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(5)
    
    
    
    fcbGrupo1.SetData gobjMain.EmpresaActual.ListaAFGrupo(1, False, False)
    fcbGrupo2.SetData gobjMain.EmpresaActual.ListaAFGrupo(2, False, False)
    fcbGrupo3.SetData gobjMain.EmpresaActual.ListaAFGrupo(3, False, False)
    fcbGrupo4.SetData gobjMain.EmpresaActual.ListaAFGrupo(4, False, False)
    fcbGrupo5.SetData gobjMain.EmpresaActual.ListaAFGrupo(5, False, False)
    
    FcbBodega.SetData gobjMain.EmpresaActual.ListaAFBodega(True, False)
    
    
    If tag = "COSTOUI" Then
        chkIVA.Caption = " Solo Costo Ultima Compra 0 "
    Else
        chkIVA.Caption = "Solo Items IVA diferente de cero"
    End If
    chkIVA.value = IIf(bandIVA, vbChecked, vbUnchecked)

    If tag = "AFEXIST" Then
        chkIVA.Visible = False
        lblBodega.Visible = True
        FcbBodega.Visible = True
    End If
    MensajeStatus
    BandAceptado = False
    FraTrans.Visible = False
    Me.Show vbModal, frmMain
    
    'Si aplastó el botón 'Aceptar'
    If BandAceptado Then
        'Devuelve los valores de condición para a búsqueda
        coditem = Trim$(txtCodigo.Text)
        CodAlt = Trim$(txtCodAlt.Text)
        Desc = Trim$(txtDesc.Text)
        bandIVA = (chkIVA.value = vbChecked)
        
        CodGrupo1 = fcbGrupo1.KeyText
        CodGrupo2 = fcbGrupo2.KeyText
        CodGrupo3 = fcbGrupo3.KeyText
        CodGrupo4 = fcbGrupo4.KeyText
        CodGrupo5 = fcbGrupo5.KeyText
        CodBodega = FcbBodega.KeyText
    End If
    
    'Devuelve true/false
    Inicio = BandAceptado
    
    Exit Function
ErrTrap:
    MensajeStatus
    DispErr
    Exit Function
End Function

'Private Sub cboGrupo_Click()
'    Dim Numg As Integer
'    On Error GoTo ErrTrap
'    If cboGrupo.ListIndex < 0 Then Exit Sub
'
'    MensajeStatus MSG_PREPARA, vbHourglass
'
'    Numg = cboGrupo.ListIndex + 1
'    fcbGrupo.SetData gobjMain.EmpresaActual.ListaIVGrupo(Numg, False, False)
'    fcbGrupo.KeyText = ""
'    MensajeStatus
'    Exit Sub
'ErrTrap:
'    MensajeStatus
'    DispErr
'    Exit Sub
'End Sub

Private Sub cmdAceptar_Click()
    If Me.tag = "IVEXIST" Then
        If Len(FcbBodega.KeyText) = 0 Then
            MsgBox " Debe Seleccionar Bodega"
            FcbBodega.SetFocus
            Exit Sub
        End If
    End If
    BandAceptado = True
    If txtCodigo.Visible = True Then
        txtCodigo.SetFocus
    Else
        fcbTrans.SetFocus
    End If
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    BandAceptado = False
    txtCodigo.SetFocus
    Me.Hide
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

Private Sub txtCodAlt_GotFocus()
    txtCodAlt.SelStart = 0
    txtCodAlt.SelLength = Len(txtCodAlt.Text)
End Sub

Private Sub txtCodigo_GotFocus()
    txtCodigo.SelStart = 0
    txtCodigo.SelLength = Len(txtCodigo.Text)
End Sub

Private Sub txtDesc_GotFocus()
    txtDesc.SelStart = 0
    txtDesc.SelLength = Len(txtDesc.Text)
End Sub

Public Function InicioTrans(ByRef CodTrans As String, _
                       ByRef desde As Long, ByRef hasta As Long) As Boolean
    Dim antes As String, i As Integer
    On Error GoTo ErrTrap
    
    'Cambia forma de cursor mientras se carga
    MensajeStatus MSG_PREPARA, vbHourglass
    FraTrans.Visible = True
    Frame1.Visible = False
    'Prepara ComboBox de etiquetas de grupo
    CargaTrans
    MensajeStatus
    BandAceptado = False
'    fraCodTrans.Visible = False
    Me.Show vbModal, frmMain
    
    'Si aplastó el botón 'Aceptar'
    If BandAceptado Then
        'Devuelve los valores de condición para a búsqueda
        CodTrans = fcbTrans.KeyText
        desde = IIf(Len(txtNumTrans1.Text) > 0, txtNumTrans1.Text, 0)
        hasta = IIf(Len(txtNumTrans2.Text) > 0, txtNumTrans2.Text, IIf(Len(txtNumTrans1.Text) > 0, txtNumTrans1.Text, 0))
    End If
    
    'Devuelve true/false
    InicioTrans = BandAceptado
    
    Exit Function
ErrTrap:
    MensajeStatus
    DispErr
    Exit Function
End Function


Private Sub CargaTrans()
    'Carga la lista de transacción
    fcbTrans.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False)
End Sub



Private Sub fcbGrupo1_Selected(ByVal Text As String, ByVal KeyText As String)
    CargarListadeGrupos 1
End Sub


Private Sub fcbGrupo2_Selected(ByVal Text As String, ByVal KeyText As String)
    CargarListadeGrupos 2
End Sub


Private Sub fcbGrupo3_Selected(ByVal Text As String, ByVal KeyText As String)
    CargarListadeGrupos 3
End Sub


Private Sub fcbGrupo4_Selected(ByVal Text As String, ByVal KeyText As String)
    CargarListadeGrupos 4
End Sub

Private Sub CargarListadeGrupos(Index As Byte)
    Dim sql As String, cond As String
    Dim Campos As String, Tablas As String
    On Error GoTo ErrTrap
    
    
    'ivgrupo2
    Campos = "Select distinct ivg2.CodGrupo2 , ivg2.Descripcion "
    Tablas = " From AFinventario iv " & _
       "INNER  join   afgrupo1 ivg1 on iv.Idgrupo1 = ivg1.Idgrupo1 " & _
       "INNER  join   afgrupo2 ivg2 on iv.Idgrupo2 = ivg2.Idgrupo2 "
    
    If Len(fcbGrupo1.KeyText) > 0 Then cond = " ivg1.CodGrupo1 = '" & fcbGrupo1.KeyText & "'"
    
    If Index <> 2 Then
        sql = Campos & Tablas & IIf(Len(cond) > 0, " WHERE " & cond, "")
        fcbGrupo2.SetData MiGetRows(gobjMain.EmpresaActual.OpenRecordset(sql))
        fcbGrupo2.KeyText = fcbGrupo2.Text
    End If
    
    'ivgrupo3
    Campos = "Select distinct ivg3.CodGrupo3 , ivg3.Descripcion "
    
    Tablas = Tablas & _
            "INNER join   afgrupo3 ivg3 on iv.Idgrupo3 = ivg3.Idgrupo3 "

    
    If Len(fcbGrupo2.KeyText) > 0 Then
        cond = cond & IIf(Len(cond) > 0, " AND ", "") & " ivg2.CodGrupo2 = '" & fcbGrupo2.KeyText & "'"
    End If
    
    If Index <> 3 Then
        sql = Campos & Tablas & IIf(Len(cond) > 0, " WHERE " & cond, "")
        fcbGrupo3.SetData MiGetRows(gobjMain.EmpresaActual.OpenRecordset(sql))
        fcbGrupo3.KeyText = fcbGrupo3.Text
    End If
    
    'ivgrupo4
    Campos = "Select distinct ivg4.CodGrupo4 , ivg4.Descripcion "
    
    Tablas = Tablas & _
            "INNER join   afgrupo4 ivg4 on iv.Idgrupo4 = ivg4.Idgrupo4 "
                
    If Len(fcbGrupo3.KeyText) > 0 Then
        cond = cond & IIf(Len(cond) > 0, " AND ", "") & " ivg3.CodGrupo3 = '" & fcbGrupo3.KeyText & "'"
    End If

    If Index <> 4 Then
        sql = Campos & Tablas & IIf(Len(cond) > 0, " WHERE " & cond, "")
        fcbGrupo4.SetData MiGetRows(gobjMain.EmpresaActual.OpenRecordset(sql))
        fcbGrupo4.KeyText = fcbGrupo4.Text
    End If
    
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    Exit Sub
End Sub


