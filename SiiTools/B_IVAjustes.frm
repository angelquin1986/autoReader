VERSION 5.00
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmB_IVAjustes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda de items"
   ClientHeight    =   4845
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   5250
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
   ScaleHeight     =   4845
   ScaleWidth      =   5250
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab sst1 
      Height          =   4185
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   7382
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   529
      TabCaption(0)   =   "&Filtro de Items"
      TabPicture(0)   =   "B_IVAjustes.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblFecha"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblh"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbld"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblG"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label5"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "dtpFecha2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "fcbGrupoHasta"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "fcbGrupoDesde"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "fcbBodega"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "dtpFecha"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "chkExistencia"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cboGrupo"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtCodAlt"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtDesc"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtCodigo"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "chkGrupos"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "Filtro de &Grupos"
      TabPicture(1)   =   "B_IVAjustes.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fcbGrupo1"
      Tab(1).Control(1)=   "fcbGrupo2"
      Tab(1).Control(2)=   "fcbGrupo3"
      Tab(1).Control(3)=   "fcbGrupo4"
      Tab(1).Control(4)=   "fcbGrupo5"
      Tab(1).Control(5)=   "lblEtiqGrupo1"
      Tab(1).Control(6)=   "lblEtiqGrupo2"
      Tab(1).Control(7)=   "lblEtiqGrupo3"
      Tab(1).Control(8)=   "lblEtiqGrupo4"
      Tab(1).Control(9)=   "lblEtiqGrupo5"
      Tab(1).ControlCount=   10
      Begin VB.CheckBox chkGrupos 
         Caption         =   "&Utilizar Filtro Avanzado de Grupo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   12
         Top             =   2970
         Width           =   2655
      End
      Begin VB.TextBox txtCodigo 
         Height          =   360
         Left            =   1245
         MaxLength       =   20
         TabIndex        =   2
         Top             =   510
         Width           =   3135
      End
      Begin VB.TextBox txtDesc 
         Height          =   360
         Left            =   1245
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1470
         Width           =   3135
      End
      Begin VB.TextBox txtCodAlt 
         Height          =   360
         Left            =   1245
         MaxLength       =   20
         TabIndex        =   4
         Top             =   990
         Width           =   3135
      End
      Begin VB.ComboBox cboGrupo 
         Height          =   360
         Left            =   165
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   3510
         Width           =   1212
      End
      Begin VB.CheckBox chkExistencia 
         Caption         =   "&Incluir existencia cero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   2940
         TabIndex        =   11
         Top             =   3000
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   375
         Left            =   1245
         TabIndex        =   10
         Top             =   2430
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   106692609
         CurrentDate     =   36789
      End
      Begin FlexComboProy.FlexCombo fcbBodega 
         Height          =   345
         Left            =   1245
         TabIndex        =   8
         Top             =   1950
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   609
         ColWidth2       =   1200
         ColWidth3       =   1200
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
      Begin FlexComboProy.FlexCombo fcbGrupoDesde 
         Height          =   345
         Left            =   1485
         TabIndex        =   16
         Top             =   3510
         Width           =   1695
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
      Begin FlexComboProy.FlexCombo fcbGrupoHasta 
         Height          =   345
         Left            =   3285
         TabIndex        =   18
         Top             =   3510
         Width           =   1695
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
      Begin FlexComboProy.FlexCombo fcbGrupo1 
         Height          =   360
         Left            =   -73755
         TabIndex        =   20
         Top             =   510
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
         Left            =   -73755
         TabIndex        =   22
         Top             =   990
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
         Left            =   -73755
         TabIndex        =   24
         Top             =   1470
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
         Left            =   -73755
         TabIndex        =   26
         Top             =   1950
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
         Left            =   -73755
         TabIndex        =   28
         Top             =   2445
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
      Begin MSComCtl2.DTPicker dtpFecha2 
         Height          =   375
         Left            =   3000
         TabIndex        =   31
         Top             =   2460
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   106692609
         CurrentDate     =   36789
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "~"
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
         Left            =   2760
         TabIndex        =   32
         Top             =   2580
         Width           =   105
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
         Left            =   -74835
         TabIndex        =   19
         Top             =   510
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
         Left            =   -74835
         TabIndex        =   21
         Top             =   997
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
         Left            =   -74835
         TabIndex        =   23
         Top             =   1470
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
         Left            =   -74835
         TabIndex        =   25
         Top             =   1950
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
         Left            =   -74835
         TabIndex        =   27
         Top             =   2550
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Cód. Item  "
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
         Left            =   165
         TabIndex        =   1
         Top             =   510
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Descripción  "
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
         Left            =   165
         TabIndex        =   5
         Top             =   1470
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Bodega  "
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
         Left            =   165
         TabIndex        =   7
         Top             =   1950
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cód. &Alterno  "
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
         Left            =   165
         TabIndex        =   3
         Top             =   990
         Width           =   945
      End
      Begin VB.Label lblG 
         AutoSize        =   -1  'True
         Caption         =   "&Grupo "
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
         Left            =   165
         TabIndex        =   13
         Top             =   3270
         Width           =   480
      End
      Begin VB.Label lbld 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde  "
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
         Left            =   1485
         TabIndex        =   15
         Top             =   3270
         Width           =   570
      End
      Begin VB.Label lblh 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta  "
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
         Left            =   3285
         TabIndex        =   17
         Top             =   3270
         Width           =   510
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         Caption         =   "&Fecha "
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
         Left            =   165
         TabIndex        =   9
         Top             =   2550
         Width           =   495
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
      Left            =   798
      TabIndex        =   29
      Top             =   4395
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
      Left            =   2838
      TabIndex        =   30
      Top             =   4395
      Width           =   1452
   End
End
Attribute VB_Name = "frmB_IVAjustes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const IVGRUPO_MAX = 5
Private BandAceptado As Boolean

Public Function Inicio(ByRef CodAlt As String, _
                       ByRef Desc As String, _
                       ByRef CodBodega As String, _
                       ByVal Formulario As String, _
                       ByRef objcond As RepCondicion) As Boolean
                       
                        
    Dim antes As String, i As Integer, sql As String
    Dim rs As Recordset, SeleccionAnterior As String
    
    
    'visualiza
    'Cambia caption de grupos
        
    '*** Temporal mente desactivamos Grupo 5 porque no usamos
    lblEtiqGrupo5.Visible = False
    fcbGrupo5.Visible = False
    
    lblEtiqGrupo1 = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(1)
    lblEtiqGrupo2 = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(2)
    lblEtiqGrupo3 = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(3)
    lblEtiqGrupo4 = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(4)
    'lblEtiqGrupo5 = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(5)

    'Cargar todos los Grupos
    fcbGrupo1.SetData gobjMain.EmpresaActual.ListaIVGrupo(1, False, False)
    fcbGrupo2.SetData gobjMain.EmpresaActual.ListaIVGrupo(2, False, False)
    fcbGrupo3.SetData gobjMain.EmpresaActual.ListaIVGrupo(3, False, False)
    fcbGrupo4.SetData gobjMain.EmpresaActual.ListaIVGrupo(4, False, False)
    
    
    SeleccionAnterior = GetSetting(APPNAME, SECTION, "bqdITEMS", "N")    ' Ultima seleccion del chkbox sobre el filtro avanzado
    
    chkGrupos.value = IIf(SeleccionAnterior = "S", vbChecked, vbUnchecked)
    chkGrupos_Click
    
    
    With objcond
        If Formulario = "AjustesInventario" Then
            dtpFecha.value = IIf(.fecha1 = 0, gobjMain.EmpresaActual.GNOpcion.FechaLimiteDesde, .fecha1)
            dtpFecha2.value = IIf(.fecha2 = 0, Date, .fecha2)
        ElseIf Formulario <> "Exis" And Formulario <> "ConsIVExistVol" Then
            lblFecha.Enabled = False
            dtpFecha.Enabled = False
        Else
            dtpFecha.value = IIf(.Fcorte = 0, Date, .Fcorte)
            'dtpFecha.Enabled = False
        End If
        
        
        'Prepara ComboBox de bodega
        If Formulario = "ListaItemPorLinea" Or Formulario = "ListaPrecios" Then
            FcbBodega.Enabled = False
            chkExistencia.Enabled = False
        Else    'Existencias  y Existencias minimas
            chkExistencia.Enabled = True
            chkExistencia.value = IIf(.Bandera = True, vbChecked, vbUnchecked)
            FcbBodega.Enabled = True
            antes = CodBodega
            FcbBodega.SetData gobjMain.EmpresaActual.ListaIVBodega(False, False)
            FcbBodega.Text = antes
        End If
        
        'Prepara ComboBox de etiquetas de grupo
        cboGrupo.Clear
        
        txtCodAlt.Text = CodAlt
        txtDesc.Text = Desc
        txtCodigo.Text = .Item1
        
        For i = 1 To IVGRUPO_MAX
            cboGrupo.AddItem gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(i)
        Next i
        
        
        If (.numGrupo <= cboGrupo.ListCount) And (.numGrupo > 0) Then
            cboGrupo.ListIndex = .numGrupo - 1   'Selecciona lo anterior
        ElseIf cboGrupo.ListCount > 0 Then
            cboGrupo.ListIndex = 0              'Selecciona la primera
        End If
        fcbGrupoDesde.KeyText = .Grupo1 'Recupera la selección anterior
        fcbGrupoHasta.KeyText = .Grupo2
        
        '***Agregado Oliver 8/dic/2003
        fcbGrupo1.KeyText = .CodGrupo1
        fcbGrupo2.KeyText = .CodGrupo2
        fcbGrupo3.KeyText = .CodGrupo3
        fcbGrupo4.KeyText = .CodGrupo4
        'fcbGrupo5.KeyText = .CodGrupo5   'no es usado temporalmente
        CargarListadeGrupos 0
        
        BandAceptado = False
        Me.Show vbModal, frmMain
        
        'Si aplastó el botón 'Aceptar'
        If BandAceptado Then
            'Guarda la configuracion del chkBox sobre el filtrto avanzado
            SaveSetting APPNAME, SECTION, "bqdITEMS", IIf(chkGrupos.value = vbChecked, "S", "N")   ' Ultima seleccion del chkbox sobre el filtro avanzado
            
            'Devuelve los valores de condición para la búsqueda
            .Bandera = IIf(chkExistencia.value = vbChecked, True, False)
            
            
                
            .Fcorte = dtpFecha.value
            .fecha1 = dtpFecha.value
            .fecha2 = dtpFecha2.value
            .Item1 = Trim$(txtCodigo.Text)
            CodAlt = Trim$(txtCodAlt.Text)
            Desc = Trim$(txtDesc.Text)
            If FcbBodega.Text = "Todas" Then
                CodBodega = ""
            Else
                CodBodega = FcbBodega.Text
            End If
            
            If cboGrupo.ListIndex >= 0 Then
                .numGrupo = cboGrupo.ListIndex + 1
                .Grupo1 = Trim$(fcbGrupoDesde.KeyText)
                .Grupo2 = Trim$(fcbGrupoHasta.KeyText)
            End If
            
            '***Agregado Oliver 8/dic/2003 para la condiciones dee filtros de grupos avanzado
            .Bandera2 = IIf(chkGrupos.value = vbChecked, True, False)
            
            .CodGrupo1 = fcbGrupo1.KeyText
            .CodGrupo2 = fcbGrupo2.KeyText
            .CodGrupo3 = fcbGrupo3.KeyText
            .CodGrupo4 = fcbGrupo4.KeyText
            '.CodGrupo5 = fcbGrupo5.KeyText
            
        Else
            .numGrupo = 0
        End If
     End With
    'Devuelve true/false
    Inicio = BandAceptado
    Unload Me
End Function

Private Sub cboGrupo_Click()
    Dim Numg As Integer
    On Error GoTo ErrTrap
    If cboGrupo.ListIndex < 0 Then Exit Sub

    MensajeStatus MSG_PREPARA, vbHourglass

    Numg = cboGrupo.ListIndex + 1
    fcbGrupoDesde.SetData gobjMain.EmpresaActual.ListaIVGrupo(Numg, False, False)
    fcbGrupoHasta.SetData fcbGrupoDesde.GetData             '*** MAKOTO 19/feb/01 Mod.
    fcbGrupoDesde.KeyText = ""
    fcbGrupoHasta.KeyText = ""
    MensajeStatus
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    Exit Sub
End Sub

Private Sub chkGrupos_Click()
    HabilitarFiltroAvanzado (chkGrupos.value = vbUnchecked)
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

Private Sub fcbBodega_Selected(ByVal Text As String, ByVal KeyText As String)
'    fcbBodega.SetFocus     '*** MAKOTO 09/ago/2000 Eliminado por que generaba error
End Sub


Private Sub fcbGrupo1_GotFocus()
    'Cuando cambia el enfoque solo con teclado activa el el tab donde esta ejecutado
    If sst1.TabEnabled(1) = True And sst1.Tab = 0 Then
        sst1.Tab = 1
    End If
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

Private Sub fcbGrupoDesde_Selected(ByVal Text As String, ByVal KeyText As String)
    fcbGrupoHasta.KeyText = fcbGrupoDesde.KeyText
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



Private Sub CargarListadeGrupos(Index As Byte)
    Dim sql As String, cond As String
    Dim Campos As String, Tablas As String
    On Error GoTo ErrTrap
    
    
    'ivgrupo2
    Campos = "Select distinct ivg2.CodGrupo2 , ivg2.Descripcion "
    Tablas = " From Ivinventario iv " & _
       "INNER  join   ivgrupo1 ivg1 on iv.Idgrupo1 = ivg1.Idgrupo1 " & _
       "INNER  join   ivgrupo2 ivg2 on iv.Idgrupo2 = ivg2.Idgrupo2 "
    
    If Len(fcbGrupo1.KeyText) > 0 Then cond = " ivg1.CodGrupo1 = '" & fcbGrupo1.KeyText & "'"
    
    If Index <> 2 Then
        sql = Campos & Tablas & IIf(Len(cond) > 0, " WHERE " & cond, "")
        fcbGrupo2.SetData MiGetRows(gobjMain.EmpresaActual.OpenRecordset(sql))
        fcbGrupo2.KeyText = fcbGrupo2.Text
    End If
    
    'ivgrupo3
    Campos = "Select distinct ivg3.CodGrupo3 , ivg3.Descripcion "
    
    Tablas = Tablas & _
            "INNER join   ivgrupo3 ivg3 on iv.Idgrupo3 = ivg3.Idgrupo3 "

    
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
            "INNER join   ivgrupo4 ivg4 on iv.Idgrupo4 = ivg4.Idgrupo4 "
                
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


Private Sub HabilitarFiltroAvanzado(modo As Boolean)
    sst1.TabEnabled(1) = Not modo
    lblG.Enabled = modo
    lbld.Enabled = modo
    lblh.Enabled = modo
    cboGrupo.Enabled = modo
    fcbGrupoDesde.Enabled = modo
    fcbGrupoHasta.Enabled = modo
    
    ' los objetos dentrto del tab1
    lblEtiqGrupo1.Enabled = Not modo
    fcbGrupo1.Enabled = Not modo
    
    lblEtiqGrupo1.Enabled = Not modo
    fcbGrupo1.Enabled = Not modo
    
    lblEtiqGrupo2.Enabled = Not modo
    fcbGrupo2.Enabled = Not modo
    
    lblEtiqGrupo3.Enabled = Not modo
    fcbGrupo3.Enabled = Not modo
    
    lblEtiqGrupo4.Enabled = Not modo
    fcbGrupo4.Enabled = Not modo
    
    lblEtiqGrupo5.Enabled = Not modo
    fcbGrupo5.Enabled = Not modo
    
End Sub

