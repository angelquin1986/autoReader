VERSION 5.00
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmIVBusqueda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Búsqueda de items"
   ClientHeight    =   3210
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   4500
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
   ScaleHeight     =   3210
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   2340
      TabIndex        =   30
      Top             =   2760
      Width           =   1452
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
      Left            =   720
      TabIndex        =   29
      Top             =   2760
      Width           =   1452
   End
   Begin TabDlg.SSTab sst 
      Height          =   2655
      Left            =   60
      TabIndex        =   14
      Top             =   60
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   4683
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Busqueda"
      TabPicture(0)   =   "IVBusqueda.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraFormaPago"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FraTrans"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Filtos"
      TabPicture(1)   =   "IVBusqueda.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblEtiqGrupo2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblEtiqGrupo3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblEtiqGrupo4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblEtiqGrupo5"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblEtiqGrupo1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblEtiqGrupo6"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "fcbGrupo6"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "fcbGrupo5"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "fcbGrupo4"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "fcbGrupo3"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "fcbGrupo2"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "fcbGrupo1"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      Begin FlexComboProy.FlexCombo fcbGrupo1 
         Height          =   360
         Left            =   -73800
         TabIndex        =   9
         Top             =   360
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
         Top             =   720
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
         Top             =   1080
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
         Top             =   1440
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
         Top             =   1800
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
         Height          =   2115
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   4155
         Begin VB.CheckBox chkIVA0 
            Caption         =   "Solo Items  IVA 0"
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
            Left            =   2520
            TabIndex        =   50
            Top             =   1680
            Width           =   1635
         End
         Begin VB.TextBox txtCodAlt 
            Height          =   372
            Left            =   1080
            MaxLength       =   20
            TabIndex        =   2
            Top             =   600
            Width           =   3012
         End
         Begin VB.TextBox txtDesc 
            Height          =   372
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   3
            Top             =   960
            Width           =   3012
         End
         Begin VB.TextBox txtCodigo 
            Height          =   372
            Left            =   1080
            MaxLength       =   20
            TabIndex        =   1
            Top             =   240
            Width           =   3012
         End
         Begin VB.CheckBox chkIVA 
            Caption         =   "Solo Items IVA "
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
            TabIndex        =   4
            Top             =   1680
            Width           =   2772
         End
         Begin FlexComboProy.FlexCombo FcbBodega 
            Height          =   375
            Left            =   1080
            TabIndex        =   5
            Top             =   1680
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
         Begin FlexComboProy.FlexCombo fcbTipo 
            Height          =   375
            Left            =   1080
            TabIndex        =   27
            Top             =   1320
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            Enabled         =   0   'False
            DispCol         =   1
            ColWidth0       =   600
            ColWidth1       =   1000
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
         Begin VB.Label Label31 
            Caption         =   "Tipo Item"
            Enabled         =   0   'False
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
            Left            =   60
            TabIndex        =   28
            Top             =   1440
            Width           =   1035
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
            TabIndex        =   24
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
            TabIndex        =   23
            Top             =   1080
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
            TabIndex        =   22
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
            TabIndex        =   26
            Top             =   1800
            Visible         =   0   'False
            Width           =   600
         End
      End
      Begin FlexComboProy.FlexCombo fcbGrupo6 
         Height          =   360
         Left            =   -73800
         TabIndex        =   48
         Top             =   2160
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
         Height          =   2055
         Left            =   120
         TabIndex        =   20
         Top             =   420
         Visible         =   0   'False
         Width           =   4155
         Begin VB.Frame fraCodTransRel 
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
            Left            =   2100
            TabIndex        =   46
            Top             =   1260
            Visible         =   0   'False
            Width           =   1995
            Begin FlexComboProy.FlexCombo fcbTransRel 
               Height          =   315
               Left            =   165
               TabIndex        =   47
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
         Begin VB.Frame fraRUC 
            Caption         =   "RUC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2100
            TabIndex        =   44
            Top             =   1320
            Visible         =   0   'False
            Width           =   1995
            Begin VB.TextBox txtruc 
               Alignment       =   1  'Right Justify
               Height          =   360
               Left            =   120
               TabIndex        =   45
               Top             =   180
               Width           =   1755
            End
         End
         Begin VB.Frame frafecha 
            Caption         =   "# Fecha. (desde - hasta)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Left            =   2100
            TabIndex        =   41
            Top             =   240
            Visible         =   0   'False
            Width           =   1995
            Begin MSComCtl2.DTPicker dtpFechaDesde 
               Height          =   375
               Left            =   180
               TabIndex        =   42
               Top             =   240
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   661
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   103809025
               CurrentDate     =   41638
            End
            Begin MSComCtl2.DTPicker dtpFechaHasta 
               Height          =   375
               Left            =   180
               TabIndex        =   43
               Top             =   600
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   661
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   103809025
               CurrentDate     =   41638
            End
         End
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
            Height          =   1035
            Left            =   60
            TabIndex        =   25
            Top             =   240
            Width           =   1995
            Begin VB.TextBox txtNumTrans2 
               Alignment       =   1  'Right Justify
               Height          =   360
               Left            =   300
               TabIndex        =   7
               Top             =   600
               Width           =   1515
            End
            Begin VB.TextBox txtNumTrans1 
               Alignment       =   1  'Right Justify
               Height          =   360
               Left            =   300
               TabIndex        =   6
               Top             =   240
               Width           =   1515
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
            Left            =   60
            TabIndex        =   21
            Top             =   1260
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
      Begin VB.Frame FraFormaPago 
         Height          =   2115
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Visible         =   0   'False
         Width           =   4155
         Begin VB.Frame Frame3 
            Caption         =   "Rango Fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   32
            Top             =   780
            Width           =   3915
            Begin MSComCtl2.DTPicker dtpDesde 
               Height          =   330
               Left            =   780
               TabIndex        =   33
               ToolTipText     =   "Fecha de la transacción"
               Top             =   240
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   582
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "yyyy/MM/dd"
               Format          =   103809025
               CurrentDate     =   37078
               MaxDate         =   73415
               MinDate         =   29221
            End
            Begin MSComCtl2.DTPicker dtpHasta 
               Height          =   330
               Left            =   780
               TabIndex        =   34
               ToolTipText     =   "Fecha de la transacción"
               Top             =   600
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   582
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "yyyy/MM/dd"
               Format          =   103809025
               CurrentDate     =   37078
               MaxDate         =   73415
               MinDate         =   29221
            End
            Begin FlexComboProy.FlexCombo fcbTransSRI 
               Height          =   315
               Left            =   2460
               TabIndex        =   35
               Top             =   600
               Width           =   1335
               _ExtentX        =   2355
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
            Begin VB.Label Label7 
               Caption         =   "Transaccion"
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
               Left            =   2460
               TabIndex        =   38
               Top             =   300
               Width           =   1275
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Hasta"
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
               Left            =   180
               TabIndex        =   37
               Top             =   660
               Width           =   420
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Desde"
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
               TabIndex        =   36
               Top             =   300
               Width           =   465
            End
         End
         Begin FlexComboProy.FlexCombo fcbProveedor 
            Height          =   375
            Left            =   1020
            TabIndex        =   39
            Top             =   300
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            DispCol         =   1
            ColWidth0       =   600
            ColWidth1       =   1000
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
            Caption         =   "Proveedor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   795
         End
      End
      Begin VB.Label lblEtiqGrupo6 
         Caption         =   "Grupo6"
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
         TabIndex        =   49
         Top             =   2265
         Width           =   1035
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
         TabIndex        =   19
         Top             =   465
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
         TabIndex        =   18
         Top             =   1905
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
         TabIndex        =   17
         Top             =   1545
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
         TabIndex        =   16
         Top             =   1185
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
         TabIndex        =   15
         Top             =   825
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmIVBusqueda"
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
                       ByRef CodGrupo6 As String, _
                       ByRef numGrupo As Integer, _
                       ByRef bandIVA As Boolean, _
                       ByRef tag As String, _
                        Optional ByRef CodBodega As String, _
                        Optional ByRef Tipo As Integer, _
                        Optional ByRef bandIVA0 As Boolean) As Boolean
    Dim antes As String, i As Integer
    On Error GoTo errtrap
    
    Dim datos(0 To 1, 0 To 7) As Variant 'AUC agregado
        datos(0, 0) = "0": datos(1, 0) = "Normal"
        datos(0, 1) = "1": datos(1, 1) = "Receta"
        datos(0, 2) = "2": datos(1, 2) = "Familia"
        datos(0, 3) = "3": datos(1, 3) = "Cambio Presentación" 'jeaa 15/09/2005
        datos(0, 4) = "4": datos(1, 4) = "Preparación"         'AUC 30/12/05
        datos(0, 5) = "5": datos(1, 5) = "Promoción" 'AUC 26/09/07
        datos(0, 6) = "6": datos(1, 6) = "Rubro" 'AUC 26/09/07
        datos(0, 7) = "7": datos(1, 7) = "Porcentaje Preparación" 'JEAA 21/09/2009
        fcbTipo.SetData datos
        Label31.Enabled = False 'Si ver que se necesita mas adelante quita este filtro
        fcbTipo.Enabled = False
    Me.tag = tag
    'Cambia forma de cursor mientras se carga
    MensajeStatus MSG_PREPARA, vbHourglass
    sst.Tab = 0
    lblEtiqGrupo1 = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(1)
    lblEtiqGrupo2 = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(2)
    lblEtiqGrupo3 = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(3)
    lblEtiqGrupo4 = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(4)
    lblEtiqGrupo5 = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(5)
    lblEtiqGrupo6 = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(6)
        
    fcbGrupo1.SetData gobjMain.EmpresaActual.ListaIVGrupo(1, False, False)
    fcbGrupo2.SetData gobjMain.EmpresaActual.ListaIVGrupo(2, False, False)
    fcbGrupo3.SetData gobjMain.EmpresaActual.ListaIVGrupo(3, False, False)
    fcbGrupo4.SetData gobjMain.EmpresaActual.ListaIVGrupo(4, False, False)
    fcbGrupo5.SetData gobjMain.EmpresaActual.ListaIVGrupo(5, False, False)
    fcbGrupo6.SetData gobjMain.EmpresaActual.ListaIVGrupo(6, False, False)
        
    fcbBodega.SetData gobjMain.EmpresaActual.ListaIVBodega(True, False)
        
    If tag = "COSTOUI" Then
        chkIVA.Caption = " Solo Costo Ultima Compra 0 "
    Else
        chkIVA.Caption = "Solo Items IVA diferente de cero"
    End If
    chkIVA0.Visible = False
    If tag = "CUENTA" Then
        chkIVA0.Visible = True
        chkIVA.Caption = "Items con IVA "
        chkIVA0.Caption = "Items sin IVA "
    End If
    chkIVA.value = IIf(bandIVA, vbChecked, vbUnchecked)
    chkIVA0.value = IIf(bandIVA0, vbChecked, vbUnchecked)

    If tag = "IVEXIST" Or tag = "MINMAX" Then
        chkIVA.Visible = False
        lblBodega.Visible = True
        fcbBodega.Visible = True
        Label31.Enabled = True
        fcbTipo.Enabled = True
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
        bandIVA0 = (chkIVA0.value = vbChecked)
        
        CodGrupo1 = fcbGrupo1.KeyText
        CodGrupo2 = fcbGrupo2.KeyText
        CodGrupo3 = fcbGrupo3.KeyText
        CodGrupo4 = fcbGrupo4.KeyText
        CodGrupo5 = fcbGrupo5.KeyText
        CodGrupo6 = fcbGrupo6.KeyText
        CodBodega = fcbBodega.KeyText
        If Len(fcbTipo.KeyText) > 0 Then Tipo = fcbTipo.KeyText
        If Len(fcbGrupo1.KeyText) > 0 Then numGrupo = 1
        If Len(fcbGrupo2.KeyText) > 0 Then numGrupo = 2
        If Len(fcbGrupo3.KeyText) > 0 Then numGrupo = 3
        If Len(fcbGrupo4.KeyText) > 0 Then numGrupo = 4
        If Len(fcbGrupo5.KeyText) > 0 Then numGrupo = 5
        If Len(fcbGrupo6.KeyText) > 0 Then numGrupo = 6
        
    End If
    
    'Devuelve true/false
    Inicio = BandAceptado
    
    Exit Function
errtrap:
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
        If Len(fcbBodega.KeyText) = 0 Then
            MsgBox " Debe Seleccionar Bodega"
            fcbBodega.SetFocus
            Exit Sub
        End If
    End If
    BandAceptado = True
    If txtCodigo.Visible = True Then
        txtCodigo.SetFocus
    Else
        If Me.tag = "FORMAPAGOSRI" Then
            fcbTransSRI.SetFocus
        Else
            fcbTrans.SetFocus
        End If
    End If
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    BandAceptado = False
    '    txtCodigo.SetFocus
    
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
    
    If txtCodigo.Visible Then
        c2.SetFocus
    End If
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
    On Error GoTo errtrap
    
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
errtrap:
    MensajeStatus
    DispErr
    Exit Function
End Function


Private Sub CargaTrans()
    'Carga la lista de transacción
    fcbTrans.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False)
    fcbTransSRI.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False)
    fcbTransRel.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False)
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

Private Sub fcbGrupo5_Selected(ByVal Text As String, ByVal KeyText As String)
    CargarListadeGrupos 5
End Sub

Private Sub fcbGrupo6_Selected(ByVal Text As String, ByVal KeyText As String)
    CargarListadeGrupos 6
End Sub


Private Sub CargarListadeGrupos(Index As Byte)
    Dim sql As String, cond As String
    Dim Campos As String, Tablas As String
    On Error GoTo errtrap
    
    
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
    
    'ivgrupo5
    Campos = "Select distinct ivg5.CodGrupo5 , ivg5.Descripcion "
    
    Tablas = Tablas & _
            "INNER join   ivgrupo5 ivg5 on iv.Idgrupo5 = ivg5.Idgrupo5 "
                
    If Len(fcbGrupo4.KeyText) > 0 Then
        cond = cond & IIf(Len(cond) > 0, " AND ", "") & " ivg4.CodGrupo4 = '" & fcbGrupo4.KeyText & "'"
    End If

    If Index <> 5 Then
        sql = Campos & Tablas & IIf(Len(cond) > 0, " WHERE " & cond, "")
        fcbGrupo5.SetData MiGetRows(gobjMain.EmpresaActual.OpenRecordset(sql))
        fcbGrupo5.KeyText = fcbGrupo5.Text
    End If
    
    'ivgrupo6
    Campos = "Select distinct ivg6.CodGrupo6 , ivg6.Descripcion "
    
    Tablas = Tablas & _
            "INNER join   ivgrupo6 ivg6 on iv.Idgrupo6 = ivg6.Idgrupo6 "
                
    If Len(fcbGrupo5.KeyText) > 0 Then
        cond = cond & IIf(Len(cond) > 0, " AND ", "") & " ivg5.CodGrupo5 = '" & fcbGrupo5.KeyText & "'"
    End If

    If Index <> 6 Then
        sql = Campos & Tablas & IIf(Len(cond) > 0, " WHERE " & cond, "")
        fcbGrupo6.SetData MiGetRows(gobjMain.EmpresaActual.OpenRecordset(sql))
        fcbGrupo6.KeyText = fcbGrupo6.Text
    End If
    
    Exit Sub
errtrap:
    MensajeStatus
    DispErr
    Exit Sub
End Sub


Public Function InicioAF(ByRef coditem As String, _
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
    On Error GoTo errtrap
    
    Me.tag = tag
    'Cambia forma de cursor mientras se carga
    MensajeStatus MSG_PREPARA, vbHourglass
    sst.Tab = 0
    lblEtiqGrupo1 = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(1)
    lblEtiqGrupo2 = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(2)
    lblEtiqGrupo3 = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(3)
    lblEtiqGrupo4 = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(4)
    lblEtiqGrupo5 = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(5)
        
    fcbGrupo1.SetData gobjMain.EmpresaActual.ListaAFGrupo(1, False, False)
    fcbGrupo2.SetData gobjMain.EmpresaActual.ListaAFGrupo(2, False, False)
    fcbGrupo3.SetData gobjMain.EmpresaActual.ListaAFGrupo(3, False, False)
    fcbGrupo4.SetData gobjMain.EmpresaActual.ListaAFGrupo(4, False, False)
    fcbGrupo5.SetData gobjMain.EmpresaActual.ListaAFGrupo(5, False, False)
    
    fcbBodega.SetData gobjMain.EmpresaActual.ListaPCProvCli(False, True, False)
        
    If tag = "COSTOUI" Then
        chkIVA.Caption = " Solo Costo Ultima Compra 0 "
    Else
        chkIVA.Caption = "Solo Items IVA diferente de cero"
    End If
    chkIVA.value = IIf(bandIVA, vbChecked, vbUnchecked)

    If tag = "AFIVEXIST" Then
        chkIVA.Visible = False
        lblBodega.Visible = True
        fcbBodega.Visible = True
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
        CodBodega = fcbBodega.KeyText
    End If
    
    'Devuelve true/false
    InicioAF = BandAceptado
    
    Exit Function
errtrap:
    MensajeStatus
    DispErr
    Exit Function
End Function

Public Function InicioFormaPagoSRI(ByRef CodProv As String, ByRef CodTrans As String, _
                       ByRef desde As Date, ByRef hasta As Date) As Boolean
    Dim antes As String, i As Integer
    On Error GoTo errtrap
    
    'Cambia forma de cursor mientras se carga
    MensajeStatus MSG_PREPARA, vbHourglass
    FraTrans.Visible = False
    Frame1.Visible = False
    FraFormaPago.Visible = True
    'Prepara ComboBox de etiquetas de grupo
    dtpDesde.value = IIf(desde = "00:00:00", Date, desde)
    dtpHasta.value = IIf(hasta = "00:00:00", Date, hasta)
    CargaTrans
    'CargaProveedor
    CargaTransFormaPago
    Label3.Caption = "F.Pago"
    MensajeStatus
    BandAceptado = False
'    fraCodTrans.Visible = False
    Me.tag = "FORMAPAGOSRI"
    Me.Show vbModal, frmMain

    'Si aplastó el botón 'Aceptar'
    If BandAceptado Then
        'Devuelve los valores de condición para a búsqueda
        CodTrans = fcbTransSRI.KeyText
        CodProv = fcbProveedor.KeyText
        
        desde = dtpDesde.value
        hasta = dtpHasta.value
    End If
    
    'Devuelve true/false
    InicioFormaPagoSRI = BandAceptado
    
    Exit Function
errtrap:
    MensajeStatus
    DispErr
    Exit Function
End Function



Public Function InicioFormaPagoSRIVentas(ByRef CodProv As String, ByRef CodTrans As String, _
                       ByRef desde As Date, ByRef hasta As Date) As Boolean
    Dim antes As String, i As Integer
    On Error GoTo errtrap
    
    'Cambia forma de cursor mientras se carga
    MensajeStatus MSG_PREPARA, vbHourglass
    FraTrans.Visible = True
    fraNumTrans.Visible = False
    Frame1.Visible = False
    FraFormaPago.Visible = True
    fraFecha.Visible = True
    fraCodTransRel.Visible = True
    fraCodTransRel.Caption = "Forma Cobro SRI"
    'Prepara ComboBox de etiquetas de grupo
    dtpFechaDesde.value = IIf(desde = "00:00:00", "01/06/2016", desde)
    dtpFechaHasta.value = IIf(hasta = "00:00:00", Date, hasta)
    CargaTrans
    CargaProveedor
    MensajeStatus
    BandAceptado = False
'    fraCodTrans.Visible = False
    Me.tag = "FORMAPAGOSRI"
    Me.Show vbModal, frmMain

    'Si aplastó el botón 'Aceptar'
    If BandAceptado Then
        'Devuelve los valores de condición para a búsqueda
        CodProv = fcbTrans.KeyText
        CodTrans = fcbProveedor.KeyText
        
        desde = dtpFechaDesde.value
        hasta = dtpFechaHasta.value
    End If
    
    'Devuelve true/false
    InicioFormaPagoSRIVentas = BandAceptado
    
    Exit Function
errtrap:
    MensajeStatus
    DispErr
    Exit Function
End Function


Private Sub CargaProveedor()
    'Carga la lista de transacción
    fcbProveedor.SetData gobjMain.EmpresaActual.ListaPCProvCli(True, False, False)
End Sub

Public Function InicioTransNew(ByRef CodTrans As String, _
                       ByRef desde As Long, ByRef hasta As Long, ByRef ruc As String, _
                        ByRef fechadesde As Date, ByRef fechahasta As Date) As Boolean
    Dim antes As String, i As Integer
    On Error GoTo errtrap
    
    'Cambia forma de cursor mientras se carga
    MensajeStatus MSG_PREPARA, vbHourglass
    FraTrans.Visible = True
    Frame1.Visible = False
    fraFecha.Visible = True
    fraRUC.Visible = True
    'Prepara ComboBox de etiquetas de grupo
    CargaTrans
    MensajeStatus
    BandAceptado = False
    dtpFechaDesde.value = IIf(fechadesde <> "00:00:00", fechadesde, Date)
    dtpFechaHasta.value = IIf(fechahasta <> "00:00:00", fechahasta, Date)
    txtruc.Text = ruc
    
'    fraCodTrans.Visible = False
    Me.Show vbModal, frmMain
    
    'Si aplastó el botón 'Aceptar'
    If BandAceptado Then
        'Devuelve los valores de condición para a búsqueda
        CodTrans = fcbTrans.KeyText
        desde = IIf(Len(txtNumTrans1.Text) > 0, txtNumTrans1.Text, 0)
        hasta = IIf(Len(txtNumTrans2.Text) > 0, txtNumTrans2.Text, IIf(Len(txtNumTrans1.Text) > 0, txtNumTrans1.Text, 0))
        fechadesde = dtpFechaDesde.value
        fechahasta = dtpFechaHasta.value
        ruc = txtruc.Text
    End If
    
    'Devuelve true/false
    InicioTransNew = BandAceptado
    
    Exit Function
errtrap:
    MensajeStatus
    DispErr
    Exit Function
End Function


Public Function InicioTransRelacion(ByRef CodTrans As String, ByRef CodTransRel As String, _
                       ByRef desde As Long, ByRef hasta As Long) As Boolean
    Dim antes As String, i As Integer
    On Error GoTo errtrap
    
    'Cambia forma de cursor mientras se carga
    MensajeStatus MSG_PREPARA, vbHourglass
    FraTrans.Visible = True
    Frame1.Visible = False
    'Prepara ComboBox de etiquetas de grupo
    CargaTrans
    
    fraCodTransRel.Visible = True
    
    MensajeStatus
    BandAceptado = False
'    fraCodTrans.Visible = False
    Me.Show vbModal, frmMain
    
    'Si aplastó el botón 'Aceptar'
    If BandAceptado Then
        'Devuelve los valores de condición para a búsqueda
        CodTrans = fcbTrans.KeyText
        CodTransRel = fcbTransRel.KeyText
        desde = IIf(Len(txtNumTrans1.Text) > 0, txtNumTrans1.Text, 0)
        hasta = IIf(Len(txtNumTrans2.Text) > 0, txtNumTrans2.Text, IIf(Len(txtNumTrans1.Text) > 0, txtNumTrans1.Text, 0))
    End If
    
    'Devuelve true/false
    InicioTransRelacion = BandAceptado
    
    Exit Function
errtrap:
    MensajeStatus
    DispErr
    Exit Function
End Function

Public Function InicioFormaCobro(ByRef CodTrans As String, ByRef codforma As String, _
                       ByRef fechadesde As Date, ByRef fechahasta As Date) As Boolean
    Dim antes As String, i As Integer
    On Error GoTo errtrap
    
    'Cambia forma de cursor mientras se carga
    MensajeStatus MSG_PREPARA, vbHourglass
    FraTrans.Visible = True
    Frame1.Visible = False
    'Prepara ComboBox de etiquetas de grupo
    CargaTransForma
    fraFecha.Visible = True
    fraCodTransRel.Visible = True
    fraNumTrans.Visible = False
    fraCodTransRel.Visible = True
    fraCodTransRel.Caption = "Forma Cobro"
    If fechadesde = "00:00:00" Then
        dtpFechaDesde.value = "01/01/" & DatePart("yyyy", Date)
    Else
        dtpFechaDesde.value = fechadesde
    End If
    If fechahasta = "00:00:00" Then
        dtpFechaHasta.value = Date
    Else
        dtpFechaHasta.value = fechahasta
    End If
    
    
    MensajeStatus
    BandAceptado = False
'    fraCodTrans.Visible = False
    Me.Show vbModal, frmMain
    
    'Si aplastó el botón 'Aceptar'
    If BandAceptado Then
        'Devuelve los valores de condición para a búsqueda
        CodTrans = fcbTrans.KeyText
        codforma = fcbTransRel.KeyText
        fechadesde = dtpFechaDesde.value
        fechahasta = dtpFechaHasta.value
    End If
    
    'Devuelve true/false
    InicioFormaCobro = BandAceptado
    
    Exit Function
errtrap:
    MensajeStatus
    DispErr
    Exit Function
End Function


Private Sub CargaTransForma()
    'Carga la lista de transacción
    fcbTrans.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False)
    fcbTransSRI.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False)
    fcbTransRel.SetData gobjMain.EmpresaActual.ListaTSFormaCobroPago(True, True, False)
    
    fcbTrans.SetData gobjMain.EmpresaActual.ListaAnexoTipoComprobante(True, False)
End Sub


Private Sub CargaTransFormaPago()
    'Carga la lista de transacción
    fcbProveedor.SetData gobjMain.EmpresaActual.ListaAnexoFormaPago(False, False)
End Sub

