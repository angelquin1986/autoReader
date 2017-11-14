VERSION 5.00
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDatosPC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos Prov/Cli"
   ClientHeight    =   6660
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4980
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   4980
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicNumtrans 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   4995
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   4995
      Begin VB.CheckBox chkNoReportaDinardap 
         Caption         =   "No Reportar"
         Height          =   255
         Left            =   1740
         TabIndex        =   43
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtnumtrans 
         Height          =   360
         Left            =   1740
         MaxLength       =   15
         TabIndex        =   42
         Top             =   60
         Width           =   3135
      End
      Begin VB.Label Label7 
         Caption         =   "No.Transaccion"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   1395
      End
   End
   Begin VB.PictureBox PicBotones 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   4980
      TabIndex        =   31
      Top             =   6165
      Width           =   4980
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar-F9"
         Height          =   372
         Left            =   1380
         TabIndex        =   33
         Top             =   60
         Width           =   972
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   372
         Left            =   2460
         TabIndex        =   32
         Top             =   60
         Width           =   972
      End
   End
   Begin VB.PictureBox picDatosSC 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4995
      Left            =   60
      ScaleHeight     =   4995
      ScaleWidth      =   4875
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   4875
      Begin VB.TextBox txtNombre 
         Height          =   360
         Left            =   840
         MaxLength       =   80
         TabIndex        =   2
         Top             =   60
         Width           =   3975
      End
      Begin VB.Frame FraPersona 
         Caption         =   "Persona"
         Enabled         =   0   'False
         Height          =   735
         Left            =   60
         TabIndex        =   27
         Top             =   1740
         Width           =   4695
         Begin VB.OptionButton optClaseSujeto 
            Caption         =   "Natural"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   1620
            TabIndex        =   7
            Top             =   300
            Width           =   975
         End
         Begin VB.OptionButton optClaseSujeto 
            Caption         =   "Juridica"
            Enabled         =   0   'False
            Height          =   255
            Index           =   2
            Left            =   3120
            TabIndex        =   8
            Top             =   300
            Width           =   915
         End
         Begin VB.OptionButton optClaseSujeto 
            Caption         =   "NO SELC"
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   6
            Top             =   300
            Width           =   1095
         End
      End
      Begin VB.PictureBox PicPerNat 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   2430
         Left            =   60
         ScaleHeight     =   2430
         ScaleWidth      =   4695
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2520
         Width           =   4695
         Begin VB.Frame FraSexo 
            Caption         =   "Sexo"
            Enabled         =   0   'False
            Height          =   735
            Left            =   0
            TabIndex        =   26
            Top             =   60
            Width           =   4635
            Begin VB.OptionButton optSexo 
               Caption         =   "Femenino"
               Enabled         =   0   'False
               Height          =   255
               Index           =   2
               Left            =   3120
               TabIndex        =   11
               Top             =   300
               Width           =   1035
            End
            Begin VB.OptionButton optSexo 
               Caption         =   "Masculino"
               Enabled         =   0   'False
               Height          =   255
               Index           =   1
               Left            =   1560
               TabIndex        =   10
               Top             =   300
               Width           =   1095
            End
            Begin VB.OptionButton optSexo 
               Caption         =   "NO SELC"
               Enabled         =   0   'False
               ForeColor       =   &H000000FF&
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   9
               Top             =   300
               Width           =   1035
            End
         End
         Begin VB.Frame FraEstadoCivil 
            Caption         =   "Estado Civil"
            Enabled         =   0   'False
            Height          =   735
            Left            =   0
            TabIndex        =   25
            Top             =   840
            Width           =   4635
            Begin FlexComboProy.FlexCombo fcbEstadoCivil 
               Height          =   375
               Left            =   120
               TabIndex        =   12
               Top             =   240
               Width           =   4395
               _ExtentX        =   7752
               _ExtentY        =   661
               Enabled         =   0   'False
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
         End
         Begin VB.Frame FraOrigenIngresos 
            Caption         =   "Origen Ingresos"
            Enabled         =   0   'False
            Height          =   735
            Left            =   0
            TabIndex        =   24
            Top             =   1620
            Width           =   4635
            Begin FlexComboProy.FlexCombo fcbOrigenIngresos 
               Height          =   375
               Left            =   120
               TabIndex        =   13
               Top             =   240
               Width           =   4395
               _ExtentX        =   7752
               _ExtentY        =   661
               Enabled         =   0   'False
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
         End
      End
      Begin FlexComboProy.FlexCombo fcbCanton 
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   840
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   661
         DispCol         =   2
         ColWidth0       =   800
         ColWidth1       =   800
         ColWidth2       =   1400
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
      Begin FlexComboProy.FlexCombo FcbParroquia 
         Height          =   375
         Left            =   840
         TabIndex        =   5
         Top             =   1260
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   661
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
      Begin FlexComboProy.FlexCombo fcbProvincia 
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   420
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   661
         DispCol         =   2
         ColWidth0       =   800
         ColWidth1       =   800
         ColWidth2       =   1400
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   0
         TabIndex        =   34
         Top             =   120
         Width           =   555
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Pro&vincia  "
         Height          =   195
         Left            =   0
         TabIndex        =   30
         Top             =   480
         Width           =   750
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Canton"
         Height          =   195
         Left            =   0
         TabIndex        =   29
         Top             =   840
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Parroquia"
         Height          =   195
         Left            =   0
         TabIndex        =   28
         Top             =   1320
         Width           =   675
      End
   End
   Begin VB.PictureBox pic1 
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   4875
      TabIndex        =   14
      Top             =   0
      Width           =   4875
      Begin VB.TextBox txtRuc 
         Height          =   360
         Left            =   2400
         MaxLength       =   20
         TabIndex        =   1
         Top             =   660
         Width           =   2415
      End
      Begin VB.TextBox txtCodTransAfectada 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1680
         TabIndex        =   18
         Top             =   3000
         Visible         =   0   'False
         Width           =   612
      End
      Begin VB.TextBox txtNumTransAfectada 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2280
         TabIndex        =   17
         Top             =   3000
         Visible         =   0   'False
         Width           =   1212
      End
      Begin FlexComboProy.FlexCombo fcbTipoDocumento 
         Height          =   375
         Left            =   0
         TabIndex        =   0
         Top             =   660
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         DispCol         =   1
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
      Begin VB.Label LblNombre 
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   660
         TabIndex        =   21
         Top             =   120
         Width           =   4035
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "RUC/CI"
         Height          =   195
         Left            =   2400
         TabIndex        =   20
         Top             =   420
         Width           =   570
      End
      Begin VB.Label lblTransAfectada 
         Caption         =   "NC/ND aplicada a:"
         Height          =   195
         Left            =   0
         TabIndex        =   19
         Top             =   3120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Nombre"
         Height          =   195
         Left            =   0
         TabIndex        =   16
         Top             =   120
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Identificación"
         Height          =   195
         Left            =   0
         TabIndex        =   15
         Top             =   420
         Width           =   1575
      End
   End
   Begin VB.PictureBox picFecha 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   4995
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   4995
      Begin VB.ComboBox cboForma 
         Height          =   315
         ItemData        =   "frmDatosPC.frx":0000
         Left            =   1740
         List            =   "frmDatosPC.frx":000D
         TabIndex        =   39
         Top             =   600
         Visible         =   0   'False
         Width           =   1875
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   375
         Left            =   1740
         TabIndex        =   36
         Top             =   180
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   661
         _Version        =   393216
         Format          =   106692609
         CurrentDate     =   41691
      End
      Begin VB.Label lblForma 
         Caption         =   "Forma Pago"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   660
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Vencimieno"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   1395
      End
   End
End
Attribute VB_Name = "frmDatosPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'El campo TransIDAfectada tiene mismo funcionamiento de PCKardex.IDAsignado, pero se lo ha colocado con el fin de controlar
'que las NC/ND solamente se harán sobre una compra a la vez.

Private mobjGNComp As GNComprobante
Private BandAceptado As Boolean
Private mVisualizando As Boolean
Private TransIDAfectada As Long
Private BandTransIDAfectada As Boolean
Private WithEvents mobjEmpresa As Empresa
Attribute mobjEmpresa.VB_VarHelpID = -1
Private pc As PCProvCli
Private FechaVenci As Date
    
Public Function Inicio(ByRef obj As PCProvCli) As String
    picDatosSC.Visible = False
    fcbTipoDocumento.SetData gobjMain.EmpresaActual.ListaAnexoTipoDocumento(True, False)
    Set pc = obj
    
    With pc
        txtruc.Text = Replace(.ruc, " ", "")
        LblNombre.Caption = .nombre
        If Len(.codtipoDocumento) > 0 Then fcbTipoDocumento.KeyText = .codtipoDocumento
    End With
    Me.Show vbModal
    
    
    Me.Height = 1980
    Inicio = IIf(BandAceptado, "O.K.", "Vacío")
    Unload Me
    Set pc = Nothing
End Function




Private Sub cmdAceptar_Click()
    If picFecha.Visible = False And pic1.Visible Then
        If Len(txtruc.Text) = 0 Then
            MsgBox "Número de identificación incorrecto"
            BandAceptado = False
            txtruc.SetFocus
            Exit Sub
        End If
        If pc.codtipoDocumento = "R" And Len(txtruc.Text) <> 13 Then
            MsgBox "Número de RUC  incorrecto"
            BandAceptado = False
            txtruc.SetFocus
            Exit Sub
        End If
        If Not pc.VerificaRUC(txtruc.Text) Then
            If fcbTipoDocumento.KeyText <> "P" And fcbTipoDocumento.KeyText <> "O" And fcbTipoDocumento.KeyText <> "T" Then
                If Len(fcbTipoDocumento.KeyText) Then
                    MsgBox "Numero de  " & (fcbTipoDocumento.Text) & " Incorrecto"
               Else
                    MsgBox "Numero de  RUC/IC  Incorrecto"
                End If
                txtruc.SetFocus
                BandAceptado = False
                Exit Sub
            End If
        End If
    End If
        If picDatosSC.Visible Then
            If Len(fcbProvincia.KeyText) = 0 Then
                MsgBox "Ingrese la Provincia.", vbInformation
                fcbProvincia.SetFocus
                BandAceptado = False
                Exit Sub

            End If
            
            If Len(fcbCanton.KeyText) = 0 Then
                MsgBox "Ingrese el Cantón.", vbInformation
                fcbCanton.SetFocus
                BandAceptado = False
                Exit Sub

            End If
            
            If Len(FcbParroquia.KeyText) = 0 Then
                MsgBox "Ingrese el Parróquia.", vbInformation
                FcbParroquia.SetFocus
                BandAceptado = False
                Exit Sub

            End If
        
        If optClaseSujeto(0).value Then
            MsgBox "Debe seleccionar el tipo de Persona.", vbInformation
            optClaseSujeto(0).SetFocus
                BandAceptado = False
                Exit Sub
        ElseIf optClaseSujeto(1).value Then
            If optSexo(0).value Then
                MsgBox "Debe seleccionar el Sexo.", vbInformation
                optSexo(0).SetFocus
                BandAceptado = False
                Exit Sub
            End If
            If Len(fcbEstadoCivil.KeyText) = 0 Then
                MsgBox "Debe seleccionar el Estado Civil", vbInformation
                fcbEstadoCivil.SetFocus
                BandAceptado = False
                Exit Sub
            End If
            If Len(fcbOrigenIngresos.KeyText) = 0 Then
                MsgBox "Debe seleccionar el Origen de los Ingresos", vbInformation
                fcbOrigenIngresos.SetFocus
                BandAceptado = False
                Exit Sub

            End If
            
        End If
            
        End If
        
    BandAceptado = True
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    Me.Hide
End Sub


Private Sub dtpFecha_Change()
    FechaVenci = dtpFecha.value
End Sub

Private Sub FcbParroquia_Selected(ByVal Text As String, ByVal KeyText As String)
    pc.codParroquia = KeyText
End Sub

Private Sub Form_Initialize()
    BandAceptado = False
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

Private Sub fcbTipoDocumento_Selected(ByVal Text As String, ByVal KeyText As String)
    On Error GoTo ErrTrap

        pc.codtipoDocumento = KeyText
        Select Case pc.codtipoDocumento
            Case "R": pc.TipoDocumento = "1"
            Case "C": pc.TipoDocumento = "2"
            Case "O": pc.TipoDocumento = "5"
            Case "P": pc.TipoDocumento = "6"
            Case "F": pc.TipoDocumento = "7"
            Case Else:  pc.TipoDocumento = "4"
        End Select
    
    
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub txtNombre_Change()
    On Error GoTo ErrTrap
        pc.nombre = txtNombre
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub

End Sub

Private Sub txtRuc_Change()
    On Error GoTo ErrTrap
        pc.ruc = txtruc
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub


Public Function InicioDINARDAP(ByRef obj As PCProvCli) As String
    Set pc = obj
    picDatosSC.Visible = True
    fcbTipoDocumento.SetData gobjMain.EmpresaActual.ListaAnexoTipoDocumento(True, False)
    fcbProvincia.SetData gobjMain.EmpresaActual.ListaPCProvincia(False, False)
    fcbCanton.SetData gobjMain.EmpresaActual.ListaPCCantonxProvincia(False, False, pc.codProvincia)
    FcbParroquia.SetData gobjMain.EmpresaActual.ListaPCParroquiaxCanton(False, False, pc.codCanton)
    fcbEstadoCivil.SetData gobjMain.EmpresaActual.ListaEstadoCivil
    fcbOrigenIngresos.SetData gobjMain.EmpresaActual.ListaOrigenIngresos

    fcbProvincia.KeyText = pc.codProvincia
    fcbCanton.KeyText = pc.codCanton
    FcbParroquia.KeyText = pc.codParroquia
    

        txtNombre.Text = pc.nombre
        picDatosSC.Enabled = True
        FraPersona.Enabled = True
        optClaseSujeto(0).Enabled = True
        optClaseSujeto(1).Enabled = True
        optClaseSujeto(2).Enabled = True
        
        If pc.Tiposujeto = "N" Then
            optClaseSujeto(1).value = True
            optClaseSujeto(0).Visible = False
            PicPerNat.Visible = True
            If pc.sexo = "M" Then
                optSexo(1).value = True
                optSexo(0).Visible = False
            ElseIf pc.sexo = "F" Then
                optSexo(2).value = True
                optSexo(0).Visible = False
            Else
                optSexo(0).value = True
                optSexo(0).Visible = True
            End If
            fcbEstadoCivil.KeyText = pc.EstadoCivil
            fcbOrigenIngresos.KeyText = pc.OrigenIngresos
            
        ElseIf pc.Tiposujeto = "J" Then
            optClaseSujeto(2).value = True
            optClaseSujeto(0).Visible = False
            PicPerNat.Enabled = False
        Else
            optClaseSujeto(0).value = True
            optClaseSujeto(0).Visible = True
            optSexo(0).value = True
            optSexo(0).Visible = True
            PicPerNat.Visible = True
        End If


    Set pc = obj
    
    With pc
        txtruc.Text = Replace(.ruc, " ", "")
        LblNombre.Caption = .nombre
        If Len(.codtipoDocumento) > 0 Then fcbTipoDocumento.KeyText = .codtipoDocumento
    End With
    
    Me.Show vbModal
    
    
    
    InicioDINARDAP = IIf(BandAceptado, "O.K.", "Vacío")
    Unload Me
    Set pc = Nothing
End Function

Private Sub fcbProvincia_Selected(ByVal Text As String, ByVal KeyText As String)
On Error GoTo ErrTrap
    pc.codProvincia = KeyText
    fcbCanton.SetData gobjMain.EmpresaActual.ListaPCCantonxProvincia(True, False, KeyText)
    fcbCanton.KeyText = ""
    FcbParroquia.KeyText = ""
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub fcbCanton_Selected(ByVal Text As String, ByVal KeyText As String)
On Error GoTo ErrTrap
    pc.codCanton = KeyText
    FcbParroquia.SetData gobjMain.EmpresaActual.ListaPCParroquiaxCanton(True, False, KeyText)
    FcbParroquia.KeyText = ""
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub optClaseSujeto_Click(Index As Integer)
    If Index = 1 Then
        PicPerNat.Enabled = True
        pc.Tiposujeto = "N"
        optClaseSujeto(0).Visible = False
        FraSexo.Enabled = True
        optSexo(0).Enabled = True
        optSexo(1).Enabled = True
        optSexo(2).Enabled = True
        FraEstadoCivil.Enabled = True
        fcbEstadoCivil.Enabled = True
        FraOrigenIngresos.Enabled = True
        fcbOrigenIngresos.Enabled = True
    ElseIf Index = 2 Then
        PicPerNat.Enabled = False
        pc.Tiposujeto = "J"
        pc.sexo = "N"
        pc.EstadoCivil = "N"
        pc.OrigenIngresos = "N"
        optClaseSujeto(0).Visible = False
        
        FraSexo.Enabled = False
        optSexo(0).Enabled = False
        optSexo(1).Enabled = False
        optSexo(2).Enabled = False
        FraEstadoCivil.Enabled = False
        fcbEstadoCivil.Enabled = False
        FraOrigenIngresos.Enabled = False
        fcbOrigenIngresos.Enabled = False
        
        
    ElseIf Index = 0 And Len(pc.Tiposujeto) > 0 Then
        MsgBox "Selección Incorrecta"
        optClaseSujeto(0).Visible = False
        optClaseSujeto(1).value = True
    Else
        PicPerNat.Visible = True
    End If
End Sub

Private Sub optSexo_Click(Index As Integer)
    If Index = 0 And Len(pc.sexo) > 0 Then
        MsgBox "Selección Incorrecta"
        optSexo(0).Visible = False
        optSexo(1).value = True
        
    ElseIf Index = 1 Then
        optSexo(0).Visible = False
        pc.sexo = "M"
    Else
        optSexo(0).Visible = False
        pc.sexo = "F"
    End If

End Sub

Private Sub fcbEstadoCivil_Selected(ByVal Text As String, ByVal KeyText As String)
    On Error GoTo ErrTrap
        pc.EstadoCivil = KeyText
    Exit Sub
ErrTrap:
    DispErr
    fcbEstadoCivil.KeyText = pc.EstadoCivil
    Exit Sub
End Sub

Private Sub fcbOrigenIngresos_Selected(ByVal Text As String, ByVal KeyText As String)
    On Error GoTo ErrTrap
        pc.OrigenIngresos = KeyText
    Exit Sub
ErrTrap:
    DispErr
    fcbOrigenIngresos.KeyText = pc.OrigenIngresos
    Exit Sub
End Sub


Public Function InicioDINARDAPFecha(ByRef fechaV As Date) As String
    pic1.Visible = False
    picFecha.Top = 1
    pic1.Left = 1
    picFecha.Visible = True
    dtpFecha.value = fechaV
    FechaVenci = fechaV
    Me.Height = 1980
    Me.Show vbModal
    fechaV = dtpFecha.value
    InicioDINARDAPFecha = IIf(BandAceptado, "O.K.", "Vacío")
    Unload Me
End Function

Public Function InicioDINARDAPFechaForma(ByRef fechaPago As Date, forma As String) As String
    pic1.Visible = False
    picFecha.Top = 1
    pic1.Left = 1
    picFecha.Visible = True
    lblForma.Visible = True
    cboForma.Visible = True
    dtpFecha.value = fechaPago
    FechaVenci = fechaPago
    Me.Height = 1980
    Me.Show vbModal
    fechaPago = dtpFecha.value
    forma = Mid$(cboForma.Text, 1, 1)
    InicioDINARDAPFechaForma = IIf(BandAceptado, "O.K.", "Vacío")
    Unload Me
End Function

Public Function InicioDINARDAPNumTrans(trans As String) As String
    pic1.Visible = False
    picFecha.Visible = False
    PicNumtrans.Top = 1
    PicNumtrans.Left = 1
    PicNumtrans.Visible = True
    txtnumtrans.MaxLength = 20
    txtnumtrans.Text = trans
    Me.Height = 1980
    
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("OmiteTrans-DINARDAP")) > 0 Then
        If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("OmiteTrans-DINARDAP") = "1" Then
            chkNoReportaDinardap.Visible = True
        End If
    End If
    
    
    Me.Show vbModal
    If BandAceptado Then
            trans = txtnumtrans.Text & ";" & IIf(chkNoReportaDinardap.value = vbChecked, 1, 0)
    End If
    InicioDINARDAPNumTrans = IIf(BandAceptado, "O.K.", "Vacío")
    Unload Me
End Function


