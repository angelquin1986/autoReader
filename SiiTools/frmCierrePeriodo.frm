VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCierrePeriodo 
   Caption         =   "Cierre de ejercicio"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7605
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5970
   ScaleWidth      =   7605
   WindowState     =   2  'Maximized
   Begin VB.ListBox lstEle 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2160
      Style           =   1  'Checkbox
      TabIndex        =   60
      Top             =   5100
      Width           =   2085
   End
   Begin VB.CommandButton cmdPasos 
      Caption         =   "GO"
      Height          =   645
      Index           =   16
      Left            =   4300
      TabIndex        =   58
      Top             =   5100
      Width           =   612
   End
   Begin VB.CommandButton cmdPasos 
      Caption         =   "GO"
      Height          =   330
      Index           =   8
      Left            =   4300
      TabIndex        =   51
      Top             =   6240
      Width           =   612
   End
   Begin VB.CommandButton cmdPasos 
      Caption         =   "GO"
      Height          =   330
      Index           =   10
      Left            =   4300
      TabIndex        =   50
      Top             =   4740
      Width           =   612
   End
   Begin VB.Frame frmRol 
      Caption         =   "Desde Roles (otro Modulo)"
      Height          =   4335
      Left            =   5100
      TabIndex        =   37
      Top             =   2220
      Width           =   4455
      Begin VB.CommandButton pasa 
         Caption         =   "Command1"
         Height          =   375
         Left            =   3000
         TabIndex        =   61
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "GO"
         Height          =   330
         Index           =   15
         Left            =   3720
         TabIndex        =   57
         Top             =   3480
         Width           =   612
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "GO"
         Height          =   330
         Index           =   14
         Left            =   3720
         TabIndex        =   55
         Top             =   2760
         Width           =   612
      End
      Begin VB.Frame Frame4 
         Caption         =   "Empresa orígen Roles"
         Height          =   1000
         Left            =   120
         TabIndex        =   42
         Top             =   480
         Width           =   4215
         Begin VB.TextBox txtOrigenRoles 
            Height          =   300
            Left            =   720
            TabIndex        =   44
            Top             =   360
            Width           =   2715
         End
         Begin VB.CommandButton cmdAbrirEmpRol 
            Caption         =   "Abrir"
            Height          =   330
            Left            =   3480
            TabIndex        =   43
            Top             =   360
            Width           =   612
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Código  "
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   600
         End
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "GO"
         Height          =   330
         Index           =   9
         Left            =   3720
         TabIndex        =   41
         Top             =   3120
         Width           =   612
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "GO"
         Height          =   330
         Index           =   11
         Left            =   3720
         TabIndex        =   40
         Top             =   1680
         Width           =   612
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "GO"
         Height          =   330
         Index           =   12
         Left            =   3720
         TabIndex        =   39
         Top             =   2040
         Width           =   612
      End
      Begin VB.CommandButton cmdPasos 
         Caption         =   "GO"
         Height          =   330
         Index           =   13
         Left            =   3720
         TabIndex        =   38
         Top             =   2400
         Width           =   612
      End
      Begin VB.Label lblPasos 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6.Pasar Configuracion de Cuentas"
         Height          =   330
         Index           =   15
         Left            =   120
         TabIndex        =   56
         Top             =   3480
         Width           =   3540
      End
      Begin VB.Label lblPasos 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4.Pasar datos personal(no catalogo)"
         Height          =   330
         Index           =   14
         Left            =   120
         TabIndex        =   54
         Top             =   2760
         Width           =   3540
      End
      Begin VB.Label lblPasos 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5.Pasar Saldos de Rol de Pagos (año actual)"
         Height          =   330
         Index           =   9
         Left            =   120
         TabIndex        =   49
         Top             =   3120
         Width           =   3540
      End
      Begin VB.Label lblPasos 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1.Copiar Elementos"
         Height          =   330
         Index           =   11
         Left            =   120
         TabIndex        =   48
         Top             =   1680
         Width           =   3540
      End
      Begin VB.Label lblPasos 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2.Copiar Departamentos"
         Height          =   330
         Index           =   12
         Left            =   120
         TabIndex        =   47
         Top             =   2040
         Width           =   3540
      End
      Begin VB.Label z 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3.Copia Cargos"
         Height          =   330
         Index           =   13
         Left            =   120
         TabIndex        =   46
         Top             =   2400
         Width           =   3540
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "NOTA"
      Height          =   675
      Left            =   180
      TabIndex        =   34
      Top             =   60
      Width           =   8475
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   60
         ScaleHeight     =   375
         ScaleWidth      =   8295
         TabIndex        =   35
         Top             =   240
         Width           =   8295
         Begin VB.Label Label6 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Antes de Ejecutar este proceso debe estar creada la nueva empresa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   36
            Top             =   60
            Width           =   7170
         End
      End
   End
   Begin VB.CommandButton cmdPasos 
      Caption         =   "GO"
      Height          =   330
      Index           =   1
      Left            =   4300
      TabIndex        =   5
      Top             =   2580
      Width           =   612
   End
   Begin VB.CommandButton cmdPasos 
      Caption         =   "GO"
      Height          =   330
      Index           =   0
      Left            =   4300
      TabIndex        =   4
      Top             =   2220
      Width           =   612
   End
   Begin VB.CommandButton cmdPasos 
      Caption         =   "GO"
      Height          =   330
      Index           =   2
      Left            =   4300
      TabIndex        =   6
      Top             =   2940
      Width           =   612
   End
   Begin VB.Frame Frame2 
      Caption         =   "Empresa destino"
      Height          =   1000
      Left            =   3060
      TabIndex        =   16
      Top             =   780
      Width           =   2772
      Begin VB.TextBox txtDestino 
         Height          =   300
         Left            =   720
         TabIndex        =   1
         Top             =   240
         Width           =   1812
      End
      Begin VB.TextBox txtDestinoBD 
         Height          =   300
         Left            =   720
         TabIndex        =   2
         Top             =   600
         Width           =   1812
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Código  "
         Height          =   192
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "B.D."
         Height          =   192
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Empresa orígen"
      Height          =   1000
      Left            =   180
      TabIndex        =   0
      Top             =   780
      Width           =   2772
      Begin VB.Label lblOrigenBD 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   720
         TabIndex        =   15
         Top             =   600
         Width           =   1812
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "B.D."
         Height          =   192
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Código  "
         Height          =   192
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   600
      End
      Begin VB.Label lblOrigen 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   720
         TabIndex        =   13
         Top             =   240
         Width           =   1812
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   3345
      Left            =   0
      ScaleHeight     =   3345
      ScaleWidth      =   7605
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2625
      Width           =   7605
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   288
         Left            =   5760
         TabIndex        =   28
         Top             =   3000
         Width           =   1212
      End
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   3000
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VSFlex7LCtl.VSFlexGrid grd 
         Height          =   2895
         Left            =   120
         TabIndex        =   30
         Top             =   0
         Width           =   7335
         _cx             =   12938
         _cy             =   5106
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCierrePeriodo.frx":0000
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   -1  'True
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   0
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   5
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.CommandButton cmdOpcion 
      Caption         =   "&Opciones..."
      Enabled         =   0   'False
      Height          =   372
      Left            =   9660
      TabIndex        =   26
      Top             =   2340
      Visible         =   0   'False
      Width           =   1812
   End
   Begin VB.CommandButton cmdPasos 
      Caption         =   "GO"
      Height          =   330
      Index           =   3
      Left            =   4300
      TabIndex        =   7
      Top             =   3300
      Width           =   612
   End
   Begin VB.CommandButton cmdPasos 
      Caption         =   "GO"
      Height          =   330
      Index           =   4
      Left            =   4300
      TabIndex        =   8
      Top             =   3660
      Width           =   612
   End
   Begin VB.CommandButton cmdPasos 
      Caption         =   "GO"
      Height          =   330
      Index           =   5
      Left            =   4300
      TabIndex        =   9
      Top             =   4020
      Width           =   612
   End
   Begin VB.CommandButton cmdPasos 
      Caption         =   "GO"
      Height          =   330
      Index           =   6
      Left            =   4300
      TabIndex        =   10
      Top             =   4380
      Width           =   612
   End
   Begin VB.CommandButton cmdPasos 
      Caption         =   "GO"
      Height          =   450
      Index           =   7
      Left            =   4300
      TabIndex        =   11
      Top             =   5760
      Width           =   612
   End
   Begin MSComCtl2.DTPicker dtpFechaCorte 
      Height          =   300
      Left            =   1500
      TabIndex        =   3
      Top             =   1860
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   529
      _Version        =   393216
      Format          =   91226113
      CurrentDate     =   36781
   End
   Begin VSFlex7Ctl.VSFlexGrid grd1 
      Height          =   1110
      Left            =   6000
      TabIndex        =   31
      Top             =   840
      Visible         =   0   'False
      Width           =   5010
      _cx             =   8826
      _cy             =   1947
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCierrePeriodo.frx":0063
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   1
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label lblPasos 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "9. Pasar saldo inicial de empleados"
      Height          =   645
      Index           =   16
      Left            =   180
      TabIndex        =   59
      Top             =   5100
      Width           =   1950
   End
   Begin VB.Label lblPasos 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "11. Desactivar la base de datos original"
      Height          =   330
      Index           =   8
      Left            =   180
      TabIndex        =   53
      Top             =   6240
      Width           =   4065
   End
   Begin VB.Label lblPasos 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "8.Pasar Roles en Historial de Roles"
      Height          =   330
      Index           =   10
      Left            =   180
      TabIndex        =   52
      Top             =   4740
      Width           =   4065
   End
   Begin VB.Label lblPasos 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "2. Crear Base de Empresa Nueva"
      Height          =   330
      Index           =   1
      Left            =   180
      TabIndex        =   33
      Top             =   2580
      Width           =   4060
   End
   Begin VB.Label lblPasos 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1. Respaldar Base Empresa Origen"
      Height          =   330
      Index           =   0
      Left            =   180
      TabIndex        =   32
      Top             =   2220
      Width           =   4060
   End
   Begin VB.Label lblPasos 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "3. Generar asiento de cierre en orígen"
      Height          =   330
      Index           =   2
      Left            =   180
      TabIndex        =   20
      Top             =   2940
      Width           =   4060
   End
   Begin VB.Label lblPasos 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10. Borrarr trans. existentes con fecha superior a la fecha de corte en Destino"
      Height          =   450
      Index           =   7
      Left            =   180
      TabIndex        =   25
      Top             =   5760
      Width           =   4065
   End
   Begin VB.Label lblPasos 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "7. Pasar saldo inicial de cuenta contable"
      Height          =   330
      Index           =   6
      Left            =   180
      TabIndex        =   24
      Top             =   4380
      Width           =   4065
   End
   Begin VB.Label lblPasos 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "6. Pasar saldo inicial de bancos"
      Height          =   330
      Index           =   5
      Left            =   180
      TabIndex        =   23
      Top             =   4020
      Width           =   4065
   End
   Begin VB.Label lblPasos 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "5. Pasar saldo inicial de proveedores y clientes"
      Height          =   330
      Index           =   4
      Left            =   180
      TabIndex        =   22
      Top             =   3660
      Width           =   4060
   End
   Begin VB.Label lblPasos 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "4. Pasar saldo inicial de inventario"
      Height          =   330
      Index           =   3
      Left            =   180
      TabIndex        =   21
      Top             =   3300
      Width           =   4060
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de corte  "
      Height          =   195
      Left            =   180
      TabIndex        =   19
      Top             =   1920
      Width           =   1155
   End
End
Attribute VB_Name = "frmCierrePeriodo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbooProcesando As Boolean
Private mbooCancelado As Boolean
Private mEmpOrigen As Empresa
Private mEmpDestino As Empresa
Private Const MSG_OK As String = "OK"

Private WithEvents mGrupo As grupo
Attribute mGrupo.VB_VarHelpID = -1

Private Type T_PROPS
    Servidor As String
    BaseDatos As String
    usuario As String
    Clave As String
    Conexion As Connection
    NumRegAfectados As Long
End Type

Private mProps As T_PROPS
Private RutaRespaldo As String
'Private RutaData As String

Public Sub Inicio()
    On Error GoTo errtrap
        CargaEleRol
    Me.Show
    Exit Sub
errtrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub cmdCancelar_Click()
    mbooCancelado = True
End Sub

Private Function VerificarOpcion() As Boolean
    Dim code As String, Fcorte As Date, i As Long, TienePermiso As Boolean
    
    'Código de emrpesa destino
    code = Trim$(txtDestino.Text)
    If Len(code) = 0 Then
        MsgBox "Ingrese el código de la empresa destino.", vbExclamation
        Exit Function
    End If
    
    'Destino no puede ser la misma que origen
    If UCase$(code) = UCase$(mEmpOrigen.CodEmpresa) Then
        MsgBox "La empresa destino no puede ser la misma que la orígen.", vbExclamation
        Exit Function
    End If
    
    'Fecha de corte
    Fcorte = dtpFechaCorte.value
    If Fcorte < mEmpOrigen.GNOpcion.FechaInicio Then
        MsgBox "La fecha de corte no puede ser antes de la fecha de inicio del período.", vbExclamation
        Exit Function
    End If
    
    'Prueba si tiene acceso a la empresa destino
    TienePermiso = False
'    For i = 1 To gobjMain.GrupoActual.CountPermiso
'        If UCase(gobjMain.GrupoActual.Permisos(i).CodEmpresa) = UCase(code) Then
'            TienePermiso = True
'            Exit For
'        End If
'    Next i
'    If Not TienePermiso Then
'        MsgBox "El usuario actual '" & gobjMain.UsuarioActual.CodUsuario & "' " & _
'               "no tiene permiso para acceder a la empresa destino '" & code & "'. " & vbCr & vbCr & _
'               "Primero deberá dar permiso necesario en el programa 'SiiConfig'.", vbInformation
'        Exit Function
'    End If
    
    VerificarOpcion = True
End Function



Private Sub LeerMedio()
Dim v As Variant, logYfis(0 To 1, 0 To 1) As Variant, i As Integer, j As Integer
On Error GoTo errtrap
'RESTORE Filelistonly From Disk = ''
v = InfMedio(lblOrigenBD.Caption, RutaRespaldo)
If Not IsEmpty(v) Then
For i = 0 To 1
    For j = 0 To 1
        logYfis(i, j) = v(i, j)
    Next j
Next i
grd1.LoadArray logYfis
AjustarAutoSize grd1, -1, -1
End If
Exit Sub
errtrap:
    MsgBox Err.Number & "  " & Err.Description, vbInformation

End Sub

Private Sub cmdPasos_Click(Index As Integer)
    Dim r As Boolean, res As Integer
    Dim i As Long
    
    Select Case Index + 1
    Case 1      '1. Respaldar base DAtos Anterior
        r = Respaldar
    Case 2      '2. Restaura Base Nueva
        r = Restaurar
    Case 3      '3. Generar asiento de cierre en orígen
        r = GenerarAsientoCierre
'    Case 2      '2. Copiar datos a destino
'        r = CopiarDatos
'    Case 3      '3. Resetear # de transacciones
'        r = ResetearNumTrans
    Case 4      '4. Pasar saldo inicial de inventario
        res = MsgBox("Saldo inicial Inventarios", vbYesNo)
        If res = vbYes Then
           r = SaldoIV
        End If
        res = MsgBox("Saldo inicial Activos Fijos", vbYesNo)
        If res = vbYes Then
            r = SaldoInicialAF
        End If
        res = MsgBox("Depreciaciones Acumuladas", vbYesNo)
        If res = vbYes Then
            r = SaldoAF
        End If
        res = MsgBox("Saldo inicial Custodios", vbYesNo)
        If res = vbYes Then
            r = SaldoInicialAFCustodios
        End If
        res = MsgBox("Saldo inicial Custodios", vbYesNo)
        If res = vbYes Then
            r = SaldoInicialAFCustodios
        End If
        res = MsgBox("Saldo inicial IVSeries", vbYesNo)
        If res = vbYes Then
           r = SaldoIVSerie
        End If
        
    Case 5      '5. Pasar saldo inicial de proveedores
        r = SaldoPC
    Case 6      '6. Pasar saldo inicial de bancos
        r = SaldoTS
    Case 7      '7. Pasar saldo inicial de cuenta contable
        r = SaldoCT
    Case 8      '8. Pasar trans. existentes con la fecha posterior a la fecha de corte
        r = BorraTrans
    Case 9      '9. Desactivar la base de datos original
        r = DesactivarOrigen
     Case 10 'AUC Pasar total ganado en un año como historial de roles del empleado desde el sistema de Roles
        res = MsgBox("Historial Roles", vbYesNo)
        If res = vbYes Then
           r = SaldoRoles
        End If
    Case 11 'AUC Pasar total ganado en un año como historial de roles del empleado de la misma empresa SII
        res = MsgBox("Historial Roles", vbYesNo)
        If res = vbYes Then
           r = SaldoRolesSii
        End If
    Case 12 'copia los elementos
        res = MsgBox("Elementos para Roles", vbYesNo)
        If res = vbYes Then
           r = Elementos
        End If
    Case 13 'copia los Depsa
        res = MsgBox("Departamentos para Roles", vbYesNo)
        If res = vbYes Then
           r = Departamentos
        End If
    Case 14 'copia los cargos
        res = MsgBox("cargos para Roles", vbYesNo)
        If res = vbYes Then
           r = Cargos
        End If
    Case 15 'copia los PERSONAL
        res = MsgBox("Personal", vbYesNo)
        If res = vbYes Then
           r = Personal
        End If
    Case 16
     'copia los CUENTASDEPARTAMENTOS EN CUENTASPERSONAL
        res = MsgBox("Cuentas para Asientos DEP->PERSONA", vbYesNo)
        If res = vbYes Then
           r = CuentasRol
           r = CuentasRolPre
        End If
        
        
    Case 17
       ' r = SaldoEmpSinRol 'AUC primero paso los que no estan con rol 01/01/2015 ya no necesito todo deberia tener un elemento
        'If r Then
            For i = 0 To lstEle.ListCount - 1 'Los que tienen rubros rol para separar
                If lstEle.Selected(i) Then
                    r = SaldoEmpRol(lstEle.List(i))
                End If
            Next
        'End If
    End Select
   
    If r Then
        If Index < cmdPasos.count - 1 Then cmdPasos(Index + 1).SetFocus
        lblPasos(Index).BackColor = vbBlue
        lblPasos(Index).ForeColor = vbYellow
    End If
End Sub

'1. Generar asiento de cierre en orígen
Private Function GenerarAsientoCierre() As Boolean
    Dim sql As String, rs As Recordset, i As Long, rpos As Long
    Dim gc As GNComprobante, ctd As CTLibroDetalle
    Dim Fcorte As Date
    On Error GoTo errtrap
        
    'Verifica las opciones
    If Not VerificarOpcion Then Exit Function
    
    Fcorte = dtpFechaCorte.value    'Fecha de corte

    'Cambia figura de cursor de mouse
    MensajeStatus "Está preparando saldos a la fecha de corte...", vbHourglass
    mensaje True, "Generando asiento de cierre..."
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True

        If mEmpOrigen.GNOpcion.ObtenerValor("PermitirDistribucionGastos") = "1" Then
            'Obtiene Saldos de cuentas contables de categoría 4 y 5 (Ingreso y Egreso)
            sql = "SELECT ct.CodCuenta, " & _
                         "Sum((ctd.Debe-ctd.Haber)/gc.Cotizacion2) AS Saldo , isnull(codgasto,'0') as codgasto " & _
                  "FROM (GNComprobante gc INNER JOIN " & _
                            "(CTLibroDetalle ctd left join gngasto gng on ctd.idgasto=gng.idgasto INNER JOIN CTCuenta ct " & _
                            "ON ctd.IdCuenta=ct.IdCuenta) " & _
                        "ON ctd.CodAsiento = gc.CodAsiento) " & _
                  "WHERE (gc.Estado IN (" & ESTADO_APROBADO & ", " & ESTADO_DESPACHADO & ", " & ESTADO_SEMDESPACHADO & ")) AND " & _
                        "(ct.TipoCuenta IN (4,5)) AND " & _
                        "(gc.FechaTrans <" & FechaYMD(Fcorte + 1, mEmpOrigen.TipoDB) & ") " & _
                  "GROUP BY ct.CodCuenta ,codgasto " & _
                  "HAVING (Sum((ctd.Debe-ctd.Haber)/gc.Cotizacion2) <> 0) " & _
                  "ORDER BY ct.CodCuenta"
        Else
            'Obtiene Saldos de cuentas contables de categoría 4 y 5 (Ingreso y Egreso)
            sql = "SELECT ct.CodCuenta, " & _
                         "Sum((ctd.Debe-ctd.Haber)/gc.Cotizacion2) AS Saldo " & _
                  "FROM (GNComprobante gc INNER JOIN " & _
                            "(CTLibroDetalle ctd left join gngasto gng on ctd.idgasto=gng.idgasto INNER JOIN CTCuenta ct " & _
                            "ON ctd.IdCuenta=ct.IdCuenta) " & _
                        "ON ctd.CodAsiento = gc.CodAsiento) " & _
                  "WHERE (gc.Estado IN (" & ESTADO_APROBADO & ", " & ESTADO_DESPACHADO & ", " & ESTADO_SEMDESPACHADO & ")) AND " & _
                        "(ct.TipoCuenta IN (4,5)) AND  BANDNIIF=0 AND " & _
                        "(gc.FechaTrans <" & FechaYMD(Fcorte + 1, mEmpOrigen.TipoDB) & ") " & _
                  "GROUP BY ct.CodCuenta " & _
                  "HAVING (Sum((ctd.Debe-ctd.Haber)/gc.Cotizacion2) <> 0) " & _
                  "ORDER BY ct.CodCuenta"
        End If
    
    Set rs = mEmpOrigen.OpenRecordset(sql)
    With rs
        If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
        i = 0
        Do Until .EOF
            prg1.value = rs.AbsolutePosition
            prg1.Refresh
            DoEvents
            rpos = 0
            MensajeStatus "Agregando detalle: #" & i & " de " & rs.RecordCount, vbHourglass
            
            'Si aplastó 'Cancelar'
            If mbooCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                GoTo cancelado
            End If
            
            Set ctd = PrepararTransCT(mEmpOrigen, "CTD", _
                        "Cierre de ejercicio", _
                        Fcorte - 1, gc, True)
            ctd.codcuenta = .Fields("CodCuenta")
            ctd.Haber = .Fields("Saldo")
            ctd.Descripcion = gc.Descripcion
            ctd.Orden = i
            
            If Len(mEmpOrigen.GNOpcion.ObtenerValor("PermitirDistribucionGastos")) > 0 Then
                If mEmpOrigen.GNOpcion.ObtenerValor("PermitirDistribucionGastos") = "1" Then
                    If .Fields("CodGasto") <> "0" Then
                        ctd.CodGasto = .Fields("CodGasto")
                    End If
                End If
            End If

            i = i + 1
            .MoveNext
        Loop
        .Close
    End With
    
    'Graba la transacción si no están grabadas
    If Not (gc Is Nothing) Then GrabarTransCT gc, True
    
    mensaje False, "", "OK"
    GenerarAsientoCierre = True
    
cancelado:
    MensajeStatus
    Set ctd = Nothing
    Set gc = Nothing
    
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
    Exit Function
errtrap:
    mensaje False, "", Err.Description
    MensajeStatus
    DispErr
    GoTo cancelado
End Function

'Agrega un detalle de TSKardex a GNComprobante
'Si comprobante llega a tener 100 detalles,
'Graba lo anterior y crea otra instancia
Private Function PrepararTransCT(ByVal e As Empresa, _
                            ByVal codt As String, _
                            ByVal Desc As String, _
                            ByVal Fcorte As Date, _
                            ByRef gc As GNComprobante, _
                            ByVal BandCierre As Boolean) As CTLibroDetalle
    Dim j As Long, ctd As CTLibroDetalle
                            
    'Crea transaccion si no existe todavía
    If gc Is Nothing Then
        Set gc = CrearTrans(e, codt, Desc, Fcorte, "")
    End If
    
    'Si llega a tener 100 detalles
    If gc.CountTSKardex >= 100 Then
        gc.HoraTrans = "00:00:01"
        GrabarTransCT gc, BandCierre
        
        'Crea nueva instancia de GNComprobante
        Set gc = CrearTrans(e, codt, Desc, Fcorte, "")
    End If

    'Agrega detalle
    j = gc.AddCTLibroDetalle
    Set PrepararTransCT = gc.CTLibroDetalle(j)
End Function

Private Sub GrabarTransCT( _
                ByVal gc As GNComprobante, _
                ByVal BandCierre As Boolean)
    Dim j As Long, ctd As CTLibroDetalle

    'Graba la transacción
    MensajeStatus "Grabándo la transacción...", vbHourglass
    
    'Si es asiento de cierre
    If BandCierre Then
        If (gc.DebeTotal - gc.HaberTotal) <> 0 Then
            'Antes de grabar, cuadra el asiento con la cuenta de resultado
            j = gc.AddCTLibroDetalle
            Set ctd = gc.CTLibroDetalle(j)
            ctd.codcuenta = gc.Empresa.GNOpcion.CodCuentaResultado
            ctd.Haber = gc.DebeTotal - gc.HaberTotal
            ctd.Descripcion = "Resultado del ejercicio"
            ctd.Orden = j
        End If
    End If
    gc.HoraTrans = "00:00:01"
    gc.Grabar False, False
End Sub

'Crear base de datos destino
'(NO ESTA USADO)
Private Function CrearDestino() As Boolean
'    Dim s As String
'    On Error GoTo errtrap
'
'    'Verifica las opciones
'    If Not VerificarOpcion Then Exit Function
'
'    s = "Para crear la base de datos de destino, váyase al programa 'SiiConfig' " & _
'           "En el menú 'Configuración' - 'Empresas', cree una nueva empresa." & _
'           "Al grabar la nueva empresa activándo la casilla que dice " & _
'           "'Crear B.D. físicamente' se creará la base de datos de nueva empresa."
'    MsgBox s, vbInformation, "Para crear destino"
'
'    CrearDestino = True
'    Exit Function
'errtrap:
'    DispErr
'    Exit Function
End Function

'2. Copiar datos a destino
Private Function CopiarDatos() As Boolean
    Dim sql As String, n As Long, e As Empresa, rpos As Long
    On Error GoTo errtrap
    
    'Verifica las opciones
    If Not VerificarOpcion Then Exit Function
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    
    n = CopiarTabla("GNOpcion", "Opciones de empresa")
    n = CopiarTabla("GNOpcion2", "Opciones de avanzadas")
    n = CopiarTabla("CTCuenta", "Plan de cuenta")
    n = CopiarTabla("TSBanco", "Catálogo de bancos")
    n = CopiarTabla("TSTipoDocBanco", "Catálogo de documentos bancarios")
    n = CopiarTabla("TSFormaCobroPago", "Catálogo de forma de pagos/cobros")
    n = CopiarTabla("IVRecargo", "Catálogo de Recargos/Descuentos")
    n = CopiarTabla("CTPresupuesto", "Presupuestos")
    n = CopiarTabla("GNTrans", "Catálogo de transacciones")
    n = CopiarTabla("GNTransAsiento", "Definición de asientos por transacción")
    n = CopiarTabla("GNTransRecargo", "Definición de recargos/descuentos por transacción")
    n = CopiarTabla("GNCentroCosto", "Catálogo de centro de costo")
    n = CopiarTabla("GNResponsable", "Catálogo de responsable")
    n = CopiarTabla("FCVendedor", "Catálogo de vendedor")
    n = CopiarTabla("IVBodega", "Catálogo de bodega")
    n = CopiarTabla("IVGrupo1", "Catálogo de Grupo1 de inventario")
    n = CopiarTabla("IVGrupo2", "Catálogo de Grupo2 de inventario")
    n = CopiarTabla("IVGrupo3", "Catálogo de Grupo3 de inventario")
    n = CopiarTabla("IVGrupo4", "Catálogo de Grupo4 de inventario")
    n = CopiarTabla("IVGrupo5", "Catálogo de Grupo5 de inventario")
    n = CopiarTabla("PCProvCli", "Catálogo de proveedores/clientes")
    n = CopiarTabla("PCContacto", "Catálogo de contactos")
    n = CopiarTabla("PCGrupo1", "Catálogo de Grupo1 de proveedores/clientes")
    n = CopiarTabla("PCGrupo2", "Catálogo de Grupo2 de proveedores/clientes")
    n = CopiarTabla("PCGrupo3", "Catálogo de Grupo3 de proveedores/clientes")
    n = CopiarTabla("PCGrupo4", "Catálogo de Grupo4 de proveedores/clientes")
    n = CopiarTabla("IVInventario", "Catálogo de inventarios")
    '***Agregado. 05/mar/02. Angel
    n = CopiarTabla("TSRetencion", "Catálogo de Retenciones")
    
'    n = CopiarTabla( "GNLogAccion", "Registro de historial")
'    n = CopiarTabla("GNVersion", "Registro de versiones")
    
    
    '***Agregado. 18/jun/03. Angel. Tabla para producción
    n = CopiarTabla("IVMateria", "Catálogo de Familias/Recetas")
    '***Agregado. 26/05/05. jeaa. Tabla descto pot item x cliente
    n = CopiarTabla("DescIVGPCG", "Catálogo de Desc Item-Clinte")
    n = CopiarTabla("InventarioProveedor", "Historial de Compras x Proveedor")
    '***Agregado. 17/06/05. jeaa. Tabla motivo de devoluciones
    n = CopiarTabla("Motivo", "Motivos de Devoluciones")
    n = CopiarTabla("IVReservacion", "Reservaciones")
    n = CopiarTabla("IVProveedorDetalle", "ProveedorDetalle")
    n = CopiarTabla("IVUnidad", "Unidad")
    n = CopiarTabla("TSRetAutoDetalle", "Retenciones AutoDetalle")
    n = CopiarTabla("IVTipoCompra", "TipoCompra")
    n = CopiarTabla("IVUnidad", "Unidad")
    '***Agregado. 12/09/2008 . jeaa. Tabla descto pot item x numPago
'    n = CopiarTabla("DescNumPagIVG", "Catálogo de Desc Item")
    
    n = CopiarTabla("IVBanco", "IvBanco")
    n = CopiarTabla("IVTarjeta", "IvTarjeta")
    n = CopiarTabla("AFBodega", "AFBodega")
    n = CopiarTabla("AFGrupo1", "Catálogo de Grupo1 de Activo Fijo")
    n = CopiarTabla("AFGrupo2", "Catálogo de Grupo2 de Activo Fijo")
    n = CopiarTabla("AFGrupo3", "Catálogo de Grupo3 de Activo Fijo")
    n = CopiarTabla("AFGrupo4", "Catálogo de Grupo4 de Activo Fijo")
    n = CopiarTabla("AFGrupo5", "Catálogo de Grupo5 de Activo Fijo")
    n = CopiarTabla("AFInventario", "AFInventario")
    
    'descuentos
    n = CopiarTabla("IVDescuento", "IVDescuento")
    n = CopiarTabla("IVDescuentoDetallePC", "IVDescuentoDetallePC")
    n = CopiarTabla("IVDescuentoDetalleIV", "IVDescuentoDetalleIV")
    n = CopiarTabla("IVDescuentoDetalleFC", "IVDescuentoDetalleFC")
    
    'promociones
    
    n = CopiarTabla("IVPromocion", "IVPromocion")
    n = CopiarTabla("IVCondPromocionDetalle", "IVCondPromocionDetalle")
    n = CopiarTabla("IVCondPromocionDetalleIVG", "IVCondPromocionDetalleIVG")
    n = CopiarTabla("IVCondPromocionDetalleP", "IVCondPromocionDetalleP")
     
'IVReservacion
    'Modifica la fecha de período contable y rango de fecha aceptable
    mensaje True, "Modificándo las fechas de inicio y fin."
    Set e = gobjMain.EmpresaActual
#If DAOLIB Then
#Else
    e.Coneccion.DefaultDatabase = Trim$(txtDestinoBD.Text)
#End If
    sql = "UPDATE GNOpcion SET FechaInicio=" & FechaYMD(e.GNOpcion.FechaFinal + 1, e.TipoDB) & ", " & _
                 "FechaFinal=" & _
                    FechaYMD(e.GNOpcion.FechaFinal _
                              + (e.GNOpcion.FechaFinal - e.GNOpcion.FechaInicio), e.TipoDB) & ", " & _
                 "FechaLimiteDesde=" & _
                    FechaYMD(e.GNOpcion.FechaFinal + 1, e.TipoDB) & ", " & _
                 "FechaLimiteHasta=" & _
                    FechaYMD(e.GNOpcion.FechaFinal _
                              + (e.GNOpcion.FechaFinal - e.GNOpcion.FechaInicio), e.TipoDB)
    e.EjecutarSQL sql, n
#If DAOLIB Then
#Else
    e.Coneccion.DefaultDatabase = e.NombreDB
#End If
    mensaje False, "", "OK"

    CopiarDatos = True
    
salida:
    mbooProcesando = False               'Desbloquea que se cierre la ventana
    Set e = Nothing
    MensajeStatus
    Exit Function
errtrap:
    mensaje False, "", Err.Description
    MensajeStatus
    DispErr
    GoTo salida
End Function

Private Function CopiarTabla( _
                    ByVal tabla As String, _
                    ByVal Desc As String) As Long
    Dim sql As String, e As Empresa, Campos As String
    Dim BaseOrig As String, BaseDest As String, NumReg As Long
    Dim tiene_id As Boolean, n As Long
    On Error GoTo errtrap

    'Sacar mensaje
    MensajeStatus "Copiando " & Desc & " (" & tabla & ") ...", vbHourglass                          'GNVersion
    mensaje True, "Copiando '" & tabla & "'..."
    DoEvents
    
    BaseOrig = "[" & gobjMain.EmpresaActual.NombreDB & "].dbo.[" & tabla & "]"
    BaseDest = "[" & Trim$(txtDestinoBD.Text) & "].dbo.[" & tabla & "]"
    Set e = gobjMain.EmpresaActual
    
    'Obtiene lista de campos
    Campos = ObtenerCampos(e, tabla, tiene_id)
        
#If DAOLIB Then
    'Pendiente
#Else
    'Si tiene columna de identity (Autonumérico), activa la inserción con valor explícito en esa columna
    If tiene_id Then
        sql = "SET IDENTITY_INSERT " & BaseDest & " ON"
        e.EjecutarSQL sql, n
    End If
    
    'Primero elimina contenido de la tabla de destino
    sql = "DELETE FROM " & BaseDest
    e.EjecutarSQL sql, n

    'Copia los datos de la tabla
    sql = "INSERT INTO " & BaseDest & " (" & Campos & ") " & _
          "SELECT " & Campos & " FROM " & BaseOrig
    e.EjecutarSQL sql, NumReg

    If tiene_id Then
        sql = "SET IDENTITY_INSERT " & BaseDest & " OFF"
        e.EjecutarSQL sql, n
    End If
#End If
    
    mensaje False, "Copiado '" & tabla & "'.", NumReg & " registros."
    CopiarTabla = NumReg
    
salida:
    MensajeStatus
    Set e = Nothing
    Exit Function
errtrap:
    MensajeStatus
    mensaje False, "", Err.Description
    DispErr
    If tiene_id Then
    sql = "SET IDENTITY_INSERT " & BaseDest & " OFF"
    e.EjecutarSQL sql, n
    End If
    GoTo salida
End Function

'Obtiene nombre de todos los campos de una tabla
' y devuelve en una cadena separado por comma
Private Function ObtenerCampos( _
                    ByVal e As Empresa, _
                    ByVal tabla As String, _
                    ByRef Identidad As Boolean) As String
    Dim s As String, sql As String, rs As Recordset
    
#If DAOLIB Then
    'Pendiente DAO
#Else
    sql = "sp_help " & tabla
    Set rs = e.OpenRecordset(sql)
    Set rs = rs.NextRecordset           'Salta al segundo conjunto
    With rs
        Do Until .EOF
            DoEvents
            If Len(s) > 0 Then s = s & ", "
            s = s & .Fields("Column_name")
            .MoveNext
        Loop
    End With
    
    'Verifica si tiene una columna de identidad (Autonumérico)
    Set rs = rs.NextRecordset           'Salta al tercer conjunto
    With rs
        Identidad = True
        Do Until .EOF
            DoEvents
            If InStr(.Fields("Identity"), "No identity") > 0 Then
                Identidad = False
                Exit Do
            End If
            .MoveNext
        Loop
        .Close
    End With
#End If
    Set rs = Nothing
    ObtenerCampos = s
End Function

'3. Resetear # de transacciones
Private Function ResetearNumTrans() As Boolean
    Dim sql As String, s As String, r As Boolean, n As Long
    On Error GoTo errtrap
    
    'Verifica las opciones
    If Not VerificarOpcion Then Exit Function
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    
    s = "Este proceso es opcional, " & _
      "desea resetear los números de transacciones?" & vbCr & vbCr & _
      "Haga click en 'No' si quiere que se sigan las " & _
      "numeraciones de transacciones. En caso contrario, " & _
      "todas las transacciones comenzarán desde el número que usted indica " & _
      "en el siguiente paso."
    If MsgBox(s, vbQuestion + vbYesNo) = vbYes Then
OtraVez:
        s = InputBox("Ingrese el número con el que comienza las " & _
            "transacciones en la nueva base de datos.", _
            "Número de transacciones", 1)
        If Len(s) = 0 Then
            MsgBox "No se resetearán los números de transacciones.", vbInformation
        Else
            If Not IsNumeric(s) Then
                MsgBox "Ingrese un valor numérico, por favor.", vbCritical
                GoTo OtraVez
            End If
            
            'Comienza el proceso
            Me.MousePointer = vbHourglass
            mensaje True, "Reseteándo números de transacciones"
            
            n = Val(s)
            sql = "UPDATE [" & Trim$(txtDestinoBD.Text) & "].dbo.GNTrans SET NumTransSiguiente=" & n
            
            'Abre la empresa destino
            gobjMain.EmpresaActual.EjecutarSQL sql, n
            r = True
        End If
    Else
        r = True
    End If
    
    mensaje False, "", "OK"
    ResetearNumTrans = r
    
salida:
    mbooProcesando = False              'Desbloquea que se cierre la ventana
    Me.MousePointer = vbNormal
    Exit Function
errtrap:
    mensaje False, "", Err.Description
    DispErr
    GoTo salida
End Function

Private Sub mensaje( _
                ByVal NuevaFila As Boolean, _
                ByVal proc As String, _
                Optional ByVal res As String)
    Dim rpos As Long
    
    rpos = grd.Rows - 1
    
    If NuevaFila Then
        grd.AddItem "" & vbTab & proc & vbTab & res
    Else
        If Len(proc) > 0 Then
            grd.TextMatrix(rpos, 1) = proc
        ElseIf Right$(grd.TextMatrix(rpos, 1), 3) = "..." Then
            'Quitar último '...'
            grd.TextMatrix(rpos, 1) = Left$(grd.TextMatrix(rpos, 1), Len(grd.TextMatrix(rpos, 1)) - 3)
        End If
        
        If Len(res) > 0 Then
            grd.TextMatrix(rpos, 2) = res
        End If
    End If
End Sub

Private Function AbrirDestino() As Empresa
    Dim e As Empresa, cod As String
    
    cod = Trim$(txtDestino.Text)
    Set e = gobjMain.RecuperaEmpresa(cod)
    e.Abrir
    Set AbrirDestino = e
    Set e = Nothing
End Function





'6. Pasar saldo inicial de bancos
Private Function SaldoTS() As Boolean
    Dim e As Empresa, tsk As TSKardex
    Dim j As Long, sql As String, rs As Recordset
    Dim i As Long, c As Currency, Fcorte As Date
    Dim gcIT As GNComprobante, gcET As GNComprobante
    On Error GoTo errtrap
    
    'Verifica las opciones
    If Not VerificarOpcion Then Exit Function
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    Fcorte = dtpFechaCorte.value    'Fecha de corte

    'Cambia figura de cursor de mouse
    MensajeStatus "Está preparando saldos a la fecha de corte...", vbHourglass
    mensaje True, "Saldo inicial de bancos..."
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True

    'Obtiene Saldos de bancos a la fecha de corte en DOLARES
    ' No incluye documentos postfechados
    sql = "SELECT ts.CodBanco, " & _
                 "Sum((tsk.Debe-tsk.Haber)/gc.Cotizacion2) AS Saldo " & _
          "FROM (GNComprobante gc INNER JOIN " & _
                    "(TSKardex tsk INNER JOIN TSBanco ts " & _
                    "ON tsk.IdBanco=ts.IdBanco) " & _
                "ON tsk.TransID = gc.TransID) " & _
          "WHERE (gc.Estado <> " & ESTADO_ANULADO & ") AND " & _
                "(tsk.FechaVenci < " & FechaYMD(Fcorte + 1, mEmpOrigen.TipoDB) & ") AND " & _
                "(gc.FechaTrans < " & FechaYMD(Fcorte + 1, mEmpOrigen.TipoDB) & ")" & _
          "GROUP BY ts.CodBanco " & _
          "HAVING (Sum((tsk.Debe-tsk.Haber)/gc.Cotizacion2) <> 0) " & _
          "ORDER BY ts.CodBanco"
    Set rs = mEmpOrigen.OpenRecordset(sql)
    
    'Abre la empresa destino
    Set e = AbrirDestino
    
    With rs
        If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
        i = 0
        Do Until .EOF
            prg1.value = rs.AbsolutePosition
            prg1.Refresh
            DoEvents
            MensajeStatus "Agregando detalle: #" & i & " de " & rs.RecordCount, vbHourglass
            
            'Si aplastó 'Cancelar'
            If mbooCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                GoTo cancelado
            End If
            
            'Si Saldo es positivo
            If .Fields("Saldo") > 0 Then
                Set tsk = PrepararTransTS(e, "IT", _
                            "Saldo inicial de bancos", _
                            Fcorte, gcIT)
            'Si Saldo es negativo
            ElseIf .Fields("Saldo") < 0 Then
                Set tsk = PrepararTransTS(e, "ET", _
                            "Saldo inicial de bancos (Negativos)", _
                            Fcorte, gcET)
            End If
            
            'Recupera datos de proveedor y asigna al objeto
            tsk.codBanco = .Fields("CodBanco")
            If .Fields("Saldo") > 0 Then
                tsk.Debe = .Fields("Saldo")
                tsk.CodTipoDoc = "NC"
            Else
                tsk.Haber = .Fields("Saldo") * -1
                tsk.CodTipoDoc = "ND"
            End If
            tsk.FechaEmision = tsk.GNComprobante.FechaTrans
            tsk.FechaVenci = tsk.FechaEmision
            tsk.nombre = "Saldo inicial"
            tsk.numdoc = "S/I"
            tsk.Observacion = ""
            tsk.Orden = i
            
            i = i + 1
            .MoveNext
        Loop
        .Close
    End With
    
    'Graba la transacción si no están grabadas
    MensajeStatus "Grabándo la transacción...", vbHourglass
    If Not (gcIT Is Nothing) Then
        gcIT.HoraTrans = "00:00:01"
        gcIT.Grabar False, False
    End If
    If Not (gcET Is Nothing) Then
        gcET.HoraTrans = "00:00:01"
        gcET.Grabar False, False
    End If
    Set gcIT = Nothing
    Set gcET = Nothing
    
    'Obtiene documentos postfechados para pasarlos uno por uno
    sql = "SELECT tsk.*, ts.CodBanco, tsd.CodTipoDoc, gc.Cotizacion2 " & _
          "FROM TSTipoDocBanco tsd INNER JOIN " & _
                    "(GNComprobante gc INNER JOIN " & _
                        "(TSKardex tsk INNER JOIN TSBanco ts " & _
                        "ON tsk.IdBanco=ts.IdBanco) " & _
                    "ON tsk.TransID = gc.TransID) " & _
                "ON tsd.IdTipoDoc = tsk.IdTipoDoc " & _
          "WHERE (gc.Estado <> " & ESTADO_ANULADO & ") AND " & _
                "(tsk.FechaVenci >= " & FechaYMD(Fcorte + 1, mEmpOrigen.TipoDB) & ") AND " & _
                "(gc.FechaTrans < " & FechaYMD(Fcorte + 1, mEmpOrigen.TipoDB) & ")" & _
          "ORDER BY ts.CodBanco"
    Set rs = mEmpOrigen.OpenRecordset(sql)
    With rs
        If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
        i = 0
        Do Until .EOF
            prg1.value = rs.AbsolutePosition
            prg1.Refresh
            DoEvents
            MensajeStatus "Agregando detalle de postfechados: #" & i & " de " & rs.RecordCount, vbHourglass
            
            'Si aplastó 'Cancelar'
            If mbooCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                GoTo cancelado
            End If
            
            'Si Saldo es positivo
            If .Fields("Debe") > 0 Then
                Set tsk = PrepararTransTS(e, "IT", _
                            "Saldo inicial de bancos (Cheques recibidos)", _
                            Fcorte, gcIT)
            'Si Saldo es negativo
            ElseIf .Fields("Haber") > 0 Then
                Set tsk = PrepararTransTS(e, "ET", _
                            "Saldo inicial de bancos (Cheques emitidos)", _
                            Fcorte, gcET)
            End If
            
            'Recupera datos de proveedor y asigna al objeto
            tsk.codBanco = .Fields("CodBanco")
            tsk.CodTipoDoc = .Fields("CodTipoDoc")
            If .Fields("Debe") > 0 Then
                tsk.Debe = .Fields("Debe") / .Fields("Cotizacion2")
            Else
                tsk.Haber = .Fields("Haber") / .Fields("Cotizacion2")
            End If
            tsk.FechaEmision = .Fields("FechaEmision")
            tsk.FechaVenci = .Fields("FechaVenci")
            tsk.nombre = .Fields("Nombre")
            tsk.numdoc = .Fields("NumDoc")
            tsk.Observacion = .Fields("Observacion")
            tsk.Orden = i
            
            i = i + 1
            .MoveNext
        Loop
        .Close
    End With
    
    'Graba la transacción si no están grabadas
    MensajeStatus "Grabándo la transacción...", vbHourglass
    If Not (gcIT Is Nothing) Then
        gcIT.HoraTrans = "00:00:01"
        gcIT.Grabar False, False
    End If
    If Not (gcET Is Nothing) Then
        gcET.HoraTrans = "00:00:01"
        gcET.Grabar False, False
    End If

    MensajeStatus
    mensaje False, "", "OK"
    MsgBox "El proceso terminó con éxito.", vbInformation
    SaldoTS = True
    
cancelado:
    Set rs = Nothing
    MensajeStatus
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
    
    'Libera los objetos utilizados
    Set tsk = Nothing
    Set gcIT = Nothing
    Set gcET = Nothing
    Set e = Nothing
    
    mbooProcesando = False                  'Desbloquea que se cierre la ventana
    Exit Function
errtrap:
    mensaje False, "", Err.Description
    MensajeStatus
    DispErr
    GoTo cancelado
End Function

'Agrega un detalle de TSKardex a GNComprobante
'Si comprobante llega a tener 100 detalles,
'Graba lo anterior y crea otra instancia
Private Function PrepararTransTS(ByVal e As Empresa, _
                            ByVal codt As String, _
                            ByVal Desc As String, _
                            ByVal Fcorte As Date, _
                            ByRef gc As GNComprobante) As TSKardex
    Dim j As Long
                            
    'Crea transaccion si no existe todavía
    If gc Is Nothing Then
        Set gc = CrearTrans(e, codt, Desc, Fcorte, "")
    End If
    
    'Si llega a tener 100 detalles
    If gc.CountTSKardex >= 100 Then
        'Graba la transacción
        MensajeStatus "Grabándo la transacción...", vbHourglass
        gc.HoraTrans = "00:00:01"
        gc.Grabar False, False
        
        'Crea nueva instancia de GNComprobante
        Set gc = CrearTrans(e, codt, Desc, Fcorte, "")
    End If

    'Agrega detalle
    j = gc.AddTSKardex
    Set PrepararTransTS = gc.TSKardex(j)
        
End Function


'7. Pasar saldo inicial de cuenta contable
Private Function SaldoCT() As Boolean
    Dim e As Empresa, ctd As CTLibroDetalle
    Dim j As Long, sql As String, rs As Recordset
    Dim i As Long, c As Currency, Fcorte As Date
    Dim gc As GNComprobante
    On Error GoTo errtrap
    
    'Verifica las opciones
    If Not VerificarOpcion Then Exit Function
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    Fcorte = dtpFechaCorte.value    'Fecha de corte

    'Cambia figura de cursor de mouse
    MensajeStatus "Está preparando saldos a la fecha de corte...", vbHourglass
    mensaje True, "Saldo inicial de cuentas contables..."
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True

    'Obtiene Saldos de asiento contable
    sql = "SELECT ct.CodCuenta, " & _
                 "Sum((ctd.Debe-ctd.Haber)/gc.Cotizacion2) AS Saldo " & _
          "FROM (GNComprobante gc INNER JOIN " & _
                    "(CTLibroDetalle ctd INNER JOIN CTCuenta ct " & _
                    "ON ctd.IdCuenta=ct.IdCuenta) " & _
                "ON ctd.CodAsiento = gc.CodAsiento) " & _
          "WHERE (gc.Estado IN (" & ESTADO_APROBADO & ", " & ESTADO_DESPACHADO & ", " & ESTADO_SEMDESPACHADO & ")) AND " & _
                "(ct.TipoCuenta IN (1, 2, 3)) AND " & _
                "(gc.FechaTrans <" & FechaYMD(Fcorte + 1, mEmpOrigen.TipoDB) & ") " & _
          "GROUP BY ct.CodCuenta " & _
          "HAVING (Sum((ctd.Debe-ctd.Haber)/gc.Cotizacion2) <> 0) " & _
          "ORDER BY ct.CodCuenta"
    Set rs = mEmpOrigen.OpenRecordset(sql)
    
    'Abre la empresa destino
    Set e = AbrirDestino
    
    With rs
        If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
        i = 0
        Do Until .EOF
            prg1.value = rs.AbsolutePosition
            prg1.Refresh
            DoEvents
            MensajeStatus "Agregando detalle: #" & i & " de " & rs.RecordCount, vbHourglass
            
            'Si aplastó 'Cancelar'
            If mbooCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                GoTo cancelado
            End If
            
            Set ctd = PrepararTransCT(e, "CTD", _
                        "Saldo inicial", _
                        Fcorte, gc, False)
            ctd.codcuenta = .Fields("CodCuenta")
            ctd.Debe = .Fields("Saldo")
            ctd.Descripcion = gc.Descripcion
            ctd.Orden = i
            
            i = i + 1
            .MoveNext
        Loop
        .Close
    End With
    
    'Graba la transacción si no están grabadas
    If Not (gc Is Nothing) Then
        gc.HoraTrans = "00:00:01"
        GrabarTransCT gc, False
    End If
    MensajeStatus
    mensaje False, "", "OK"
    MsgBox "El proceso terminó con éxito.", vbInformation
    SaldoCT = True
    
cancelado:
    Set rs = Nothing
    MensajeStatus
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
    
    'Libera los objetos utilizados
    Set ctd = Nothing
    Set gc = Nothing
    Set e = Nothing
    
    mbooProcesando = False                 'Desbloquea que se cierre la ventana
    Exit Function
errtrap:
    mensaje False, "", Err.Description
    MensajeStatus
    DispErr
    GoTo cancelado
End Function


'8. Pasar trans. existentes con la fecha posterior a la fecha de corte
'Private Function CopiaTrans() As Boolean
'    Dim i As Long, codt As String, numt As Long
'    Dim empDestino As Empresa, sql As String, num As Long
'    'Verifica  errores  en la base de Origen
'
'    Set empDestino = AbrirDestino
'    If empDestino.NombreDB = mEmpOrigen.NombreDB Then
'        MsgBox "La empresa origen y destino son las mismas" & Chr(13) & _
'               "debera  seleccionar  una empresa de  destino diferente", vbExclamation
'        Exit Function
'    End If
'    If grd.ColKey(grd.Cols - 1) <> "Resultado" Then
'        If VerificaFechaVenci = False Then
'            MsgBox "Se han encontrado  errores  en las siguientes  transacciones " & Chr(13) & _
'                   "La fecha de vencimiento no  puede ser mayor  a la fecha de transacción " & Chr(13) & _
'                    "Primero deberá  corregirlos  para  proceder a copiar las transacciones.", vbInformation
'
'            Exit Function
'        End If
'    End If
'    If grd.FixedRows = grd.Rows Then CargaTrans 'Carga  trans solo  si la grlla esta vacia
'    'Transferir  transaccion una por una
'    If MsgBox("Este proceso tardará  algunos minutos " & Chr(13) & " Desea comenzar el proceso de importación?", _
'                vbYesNo + vbQuestion) <> vbYes Then Exit Function
'
'    prg1.Min = 0
'    mbooCancelado = False
'    cmdCancelar.Enabled = True
'
'
'    mbooProcesando = True               'Bloquea que se cierre la ventana
'    MensajeStatus "Copiando...", vbHourglass
'
'    With grd
'        prg1.Min = .FixedRows - 1
'        prg1.max = .Rows - 1
'        prg1.value = prg1.Min
'
'        For i = .FixedRows To .Rows - 1
'            prg1.value = i
'            DoEvents                'Para dar control a Windows
'            'Si usuario aplastó 'Cancelar', sale del ciclo
'            If mbooCancelado Then
'                MsgBox "El proceso fue cancelado.", vbInformation
'                GoTo cancelado
'            End If
'            .ShowCell i, 0          'Hace visible la fila actual
'
''            If .IsSelected(i) Then
'                codt = .TextMatrix(i, .ColIndex("CodTrans"))
'                numt = .TextMatrix(i, .ColIndex("NumTrans"))
'                'If codt = "FC" And numt = 2524 Then Stop
'                MensajeStatus "Copiando la transacción " & codt & numt & _
'                            "     " & i & " de " & .Rows - .FixedRows & _
'                            " (" & Format(i * 100 / (.Rows - .FixedRows), "0") & "%)", vbHourglass
'
'                'Si aún no está importado bien, importa la fila
'                If grd.TextMatrix(i, .Cols - 1) <> MSG_OK Then
'                    If ImportarTransSub(codt, numt, empDestino) Then
'                        .TextMatrix(i, .ColIndex("Resultado")) = MSG_OK
'                    Else
'                        .TextMatrix(i, .ColIndex("Resultado")) = "Error"
'                    End If
'               End If
'       Next i
'    End With
'    'Corregir  error de Idasignado
'    MensajeStatus "Reasignando relaciones ...", vbHourglass
'    sql = " UPDATE b SET b.IdAsignado = c.Id " & _
'           " From    " & _
'           empDestino.NombreDB & ".dbo.PCKardex c INNER JOIN " & _
'           mEmpOrigen.NombreDB & ".dbo.PCKardex a INNER JOIN " & empDestino.NombreDB & ".dbo.PCKardex b " & _
'           " ON a.Id  = b.IdAsignado " & _
'           " ON c.Guid = a.Guid " & _
'           " Where a.idAsignado = 0 And b.idAsignado <> 0 And c.idAsignado = 0 "
'
'    mEmpOrigen.EjecutarSQL sql, num
'    MsgBox "Proceso terminado con exito"
'
'
'    MensajeStatus
'    mbooProcesando = False  'Bloquea que se cierre la ventana
'    CopiaTrans = True
'cancelado:
'    MensajeStatus
'    mbooProcesando = False
'    prg1.value = prg1.Min
'    Exit Function
'ErrTrap:
'    MensajeStatus
'    DispErr
'    mbooProcesando = False
'    prg1.value = prg1.Min
'    Exit Function
'End Function


Private Function VerificaFechaVenci() As Boolean
    Dim sql As String, rs As Recordset, v  As Variant
    Dim Fcorte As Date
    
    Fcorte = dtpFechaCorte.value
    MensajeStatus "Verificando datos en Base Origen.....", vbHourglass
    sql = "Select FechaTrans, CodTrans, Numtrans, 'Error fecha vencimiento' as Estado " & _
          "From PCKardex  PCK Inner Join GnComprobante GNC On PCK.TransID = GNC.TransID " & _
          "Where GNC.Fechatrans > PCK.FechaVenci AND GNC.FechaTrans > " & FechaYMD(Fcorte, mEmpOrigen.TipoDB)
    
    Set rs = mEmpOrigen.OpenRecordset(sql)
    grd.Rows = grd.FixedRows
    If Not rs.EOF Then
        v = MiGetRows(rs)
        With grd
            .Redraw = flexRDNone
            .LoadArray v            'Carga a la grilla
        
            .FormatString = "^#|<Fecha|<CodTrans|<NumTrans|^Resultado"
            AjustarAutoSize grd, -1, -1, 3000
            GNPoneNumFila grd, False
            MensajeStatus "Errores  en la base origen"
            .Redraw = flexRDBuffered
            VerificaFechaVenci = False
            
       End With
    Else
        MensajeStatus
        
        VerificaFechaVenci = True
    End If
End Function



Private Function ImportarTransSub( _
                ByVal codt As String, _
                ByVal numt As Long, ByRef empDestino As Empresa) As Boolean
    Dim gnDest As GNComprobante, s As String, Estado As Byte, gnOri As GNComprobante
    On Error GoTo errtrap
'    Abre la empresa destino
    Set gnOri = mEmpOrigen.RecuperaGNComprobante(0, codt, numt)
'    Si existe en el destino, sobreescribe
    Set gnDest = empDestino.RecuperaGNComprobante(0, codt, numt)
    If (gnDest Is Nothing) Then
        Set gnDest = empDestino.CreaGNComprobante(codt)    'Crea  gnComprobante
    End If
    Estado = gnOri.Estado
    gnDest.Clone gnOri
    gnDest.Grabar False, False
    
'    Forzar el valor de Estado original, debido a que al Grabar cambia sin querer
    On Error Resume Next
    If gnDest.Estado = 1 And Estado = 3 Then
        'Primero Cambia  a estado cero
        empDestino.CambiaEstadoGNCompCierre gnDest.TransID, 0
    End If
    'Para  que no  considere  el IdAsignado
    empDestino.CambiaEstadoGNCompCierre gnDest.TransID, Estado
    ImportarTransSub = True
salida:
    Set gnDest = Nothing
    Set gnOri = Nothing
    Exit Function
errtrap:
'    DispMsg "Importar la trans. " & codt & numt, "Error", Err.Description
    If MsgBox(Err.Description & vbCr & vbCr & _
                "Desea continuar con siguiente transacción?", _
                vbQuestion + vbYesNo) <> vbYes Then
'        mCancelado = True
    End If
    GoTo salida
End Function



'Private Sub CargaTrans()
'    Dim sql As String, rs As Recordset, v As Variant
'    Dim Fcorte As Date
'    On Error GoTo ErrTrap
'
'    Fcorte = dtpFechaCorte.value    'Fecha de corte
'    'Selecciona las transacciones de la  base de origen
'    sql = "SELECT FechaTrans, CodTrans, NumTrans, Descripcion " & _
'          "FROM GNComprobante " & _
'                " Where FechaTrans > " & FechaYMD(Fcorte, mEmpOrigen.TipoDB) & _
'                " ORDER BY FechaTrans"
'' 10/12/2004  antes estaba el orden tambien opor codtrans, y numtrans
'
'    'Set rs = New Recordset
'    Set rs = mEmpOrigen.OpenRecordset(sql)
'    If Not rs.EOF Then
'        v = MiGetRows(rs)
'        With grd
'            .Redraw = flexRDNone
'            .LoadArray v            'Carga a la grilla
'            .FormatString = "^#|<Fecha|<CodTrans|<NumTrans|<Descripción|<Resultado"
'            GNPoneNumFila grd, False
'            AsignarTituloAColKey grd            'Para usar ColIndex
'            AjustarAutoSize grd, 0, -1, 3000     'Ajusta automáticamente ancho de cols.
'            If .ColWidth(.ColIndex("Descripción")) > 1400 Then .ColWidth(.ColIndex("Descripción")) = 1400
'            .ColWidth(.ColIndex("Resultado")) = 1600
'            'Tipo de datos
'            .ColDataType(.ColIndex("Fecha")) = flexDTDate
'            .ColDataType(.ColIndex("CodTrans")) = flexDTString
'            .ColDataType(.ColIndex("NumTrans")) = flexDTLong
'            .ColDataType(.ColIndex("Descripción")) = flexDTString
'            '.ColDataType(.ColIndex("Cod.C.C.")) = flexDTString
'            '.ColDataType(.ColIndex("Estado")) = flexDTShort
'            .ColDataType(.ColIndex("Resultado")) = flexDTString
'
'            .Redraw = flexRDDirect
'        End With
'
'    Else
'        'Si no hay nada de resultado limpia la grilla
'        grd.Rows = grd.FixedRows
'    End If
'    rs.Close
''    mBuscado = True
'salida:
'    MensajeStatus
'    Set rs = Nothing
'    Exit Sub
'ErrTrap:
'    MensajeStatus
'    DispErr
'    GoTo salida
'End Sub

'9. Desactivar la base de datos origen
Private Function DesactivarOrigen() As Boolean
    Dim v As Variant, i As Long, codg As String
    Dim g As grupo, p As Permiso, j As Long, k As Long, pt As PermisoTrans
    Dim s As String
    On Error GoTo errtrap
    
    'Verifica las opciones
    If Not VerificarOpcion Then Exit Function
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    v = gobjMain.ListaGrupos(True)
    If Not IsEmpty(v) Then
        For i = LBound(v, 2) To UBound(v, 2)
            codg = v(0, i)
            MensajeStatus "Procesando grupo '" & codg & "'...", vbHourglass
            mensaje True, "Procesando grupo '" & codg & "'"
            Set g = gobjMain.RecuperaGrupo(codg)
            Set mGrupo = g          'Para recibir evento 'mGrupo_Procesando'
            
            For j = 1 To g.CountPermiso
                DoEvents
                Set p = g.Permisos(j)
                'Si el grupo tiene permiso para la empresa actual
                If UCase(p.CodEmpresa) = UCase(mEmpOrigen.CodEmpresa) Then
                    'Confirmación
                    s = "Desea bloquear la modificación de datos de la empresa '" & _
                        mEmpOrigen.CodEmpresa & " (" & mEmpOrigen.Descripcion & ")' " & _
                        "para el grupo de usuario '" & codg & "'?" & vbCr & vbCr & _
                        "Si aplasta 'Sí', los usuarios que pertenecen al grupo " & _
                        "no podrán realizar ningún cambio a los datos de la dicha empresa. " & vbCr & _
                        "Sin embargo si fuera necesario se podrá desbloquear de nuevo " & _
                        "usando el programa 'SiiConfig' con código de usuario que tenga " & _
                        "derecho de supervisor."
                    If MsgBox(s, vbYesNo + vbQuestion) = vbYes Then
                        'Bloquea modificación/creación de todas las transacciones
                        For k = 1 To p.CountTrans
                            Set pt = p.trans(k)
                            With pt
                                .Anular = False
                                .Aprobar = False
                                .crear = False
                                .Desaprobar = False
                                .Despachar = False
                                .Eliminar = False
                                .Modificar = False
    '                            .Ver = False           'Permiso para ver no desactivemos
                            End With
                            Set pt = Nothing
                        Next k
                        
                        'Bloquea modificación de todos los catálogos
                        With p
                            .CatAFMod = False
                            .CatBancoMod = False
                            .CatBodegaMod = False
                            .CatCentroCostoMod = False
                            .CatClienteMod = False
                            .CatInfEmpresaMod = False
                            .CatInventarioMod = False
                            .CatInventarioPrecioMod = False
                            .CatPlanCuentaMod = False
                            .CatProveedorMod = False
                            .CatResponsableMod = False
                            .CatRolMod = False
                            .CatVendedorMod = False
                        End With
                    End If
                    
                    Exit For        'Pasa a procesar siguiente grupo
                End If
                Set p = Nothing
            Next j
            
            'Si el grupo está modificado, graba el grupo
            If g.Modificado Then
                MensajeStatus "Grabando grupo '" & codg & "'...", vbHourglass
                mensaje False, "Grabando grupo '" & codg & "'"
                g.Grabar
                mensaje False, "", "OK. Desactivado."
            Else
                mensaje False, "", "OK."
            End If
            Set g = Nothing
        Next i
    End If
    
    DesactivarOrigen = True
    
salida:
    mbooProcesando = False               'Desbloquea que se cierre la ventana
    Set pt = Nothing
    Set p = Nothing
    Set g = Nothing
    Set mGrupo = Nothing
    MensajeStatus
    Exit Function
errtrap:
    MensajeStatus
    DispErr
    GoTo salida
End Function



Private Sub Command1_Click()
Respaldar
End Sub

Private Sub Command2_Click()
    LeerMedio
    Restaurar
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        Unload Me
    Case Else
        MoverCampo Me, KeyCode, Shift, True
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ImpideSonidoEnter Me, KeyAscii
End Sub

Private Sub Form_Load()
    'Guarda referencia a la empresa de origen
    Set mEmpOrigen = gobjMain.EmpresaActual

    'Fecha de corte asignamos predeterminadamente FechaFinal
    dtpFechaCorte.value = mEmpOrigen.GNOpcion.FechaFinal
    
    'Visualiza codigo de empresa origen (= Empresa actual)
    lblOrigen.Caption = mEmpOrigen.CodEmpresa
    lblOrigenBD.Caption = mEmpOrigen.NombreDB
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = mbooProcesando
End Sub

Private Sub Form_Resize()
'    On Error Resume Next
'    grd.Move 0, grd.Top, Me.ScaleWidth, Me.ScaleHeight - grd.Top
End Sub





Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo errtrap
    
    MensajeStatus

    'Cierra y abre de nuevo para que quede como EmpresaActual
    mEmpOrigen.Cerrar
    mEmpOrigen.Abrir
    
    'Libera la referencia
    Set mEmpOrigen = Nothing
    Exit Sub
errtrap:
    Set mEmpOrigen = Nothing
    DispErr
    Exit Sub
End Sub





'4. Pasar saldo inicial de inventario
Private Function SaldoIV() As Boolean
    Dim e As Empresa, gc As GNComprobante, ivk As IVKardex, iv As IVinventario
    Dim j As Long, n As Long
    Dim sql As String, rs As Recordset, codOrig As String
    Dim i As Long, c As Currency, Fcorte As Date
    On Error GoTo errtrap
    
    'Verifica las opciones
    If Not VerificarOpcion Then Exit Function
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    
    codOrig = gobjMain.EmpresaActual.CodEmpresa
    Fcorte = dtpFechaCorte.value    'Fecha de corte

    'Cambia figura de cursor de mouse
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True
    
    'Saca las existencias a la fecha de corte
    MensajeStatus "Preparando para grabar las existencias iniciales...", vbHourglass
    mensaje True, "Saldo inicial de inventario..."
    
    sql = "SELECT ivk.IdInventario, ivk.IdBodega, " & _
                "iv.CodInventario, ivb.CodBodega, " & _
                "Sum(ivk.Cantidad) AS Exist " & _
          "FROM IVBodega ivb INNER JOIN " & _
                    "(IVInventario iv INNER JOIN " & _
                        "(GNTrans gt INNER JOIN " & _
                            "(IVKardex ivk INNER JOIN GNComprobante gc " & _
                            "ON ivk.TransID=gc.TransID) " & _
                        "ON gt.CodTrans=gc.CodTrans) " & _
                    "ON iv.IdInventario = ivk.IdInventario) " & _
                "ON ivb.IdBodega = ivk.IdBodega " & _
          "WHERE (gc.Estado<>" & ESTADO_ANULADO & ") AND " & _
                 "(gt.AfectaCantidad=" & CadenaBool(True, gobjMain.EmpresaActual.TipoDB) & ") AND " & _
                 "(gc.FechaTrans < " & FechaYMD(Fcorte + 1, gobjMain.EmpresaActual.TipoDB) & ") AND " & _
                 "(iv.BandServicio=" & CadenaBool(False, gobjMain.EmpresaActual.TipoDB) & ") " & _
          "GROUP BY ivk.IdInventario, ivk.IdBodega, iv.CodInventario, ivb.CodBodega " & _
          "HAVING Sum(ivk.Cantidad)>0"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
#If DAOLIB = 0 Then
    Set rs.ActiveConnection = Nothing
#End If
    
    'Abre la empresa destino
    Set e = AbrirDestino
    
    With rs
        If Not rs.EOF Then
            rs.MoveLast
            rs.MoveFirst
            If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
            i = 0
            Do Until .EOF
                prg1.value = rs.AbsolutePosition
                prg1.Refresh
                DoEvents
                
                'Si aplastó 'Cancelar'
                If mbooCancelado Then
                    MsgBox "El proceso fue cancelado.", vbInformation
                    GoTo cancelado
                End If
                
                'Crea transaccion 'IVSI'
                If (i Mod 100) = 0 Then
                    'Si no es primera vez
                    If Not (gc Is Nothing) Then
                        'Graba la transacción
                        MensajeStatus "Grabando la transacción en la empresa '" & gc.Empresa.CodEmpresa & "'...", vbHourglass
                        gc.HoraTrans = "00:00:01"
                        gc.Grabar False, False
                    End If
                    
                    Set gc = CrearTrans(e, _
                            "IVSI", _
                            "Saldo inicial de inventario", _
                            Fcorte, _
                            "")
                End If
                
                'Recupera datos de inventario para llama el método Costo()
                MensajeStatus "Agregando detalle #" & i & " de " & rs.RecordCount, vbHourglass
                Set iv = mEmpOrigen.RecuperaIVInventarioQuick(.Fields("IdInventario"))
                
                'Obtiene Costo del item en Moneda de item
                c = iv.costo(Fcorte, 1)
                
                'De moneda de item, covierte en moneda de trans, si es necesario
                If iv.CodMoneda <> gc.CodMoneda Then
                    c = c * gc.Cotizacion(iv.CodMoneda) / gc.Cotizacion("")
                End If
                
                'Agrega detalle
                j = gc.AddIVKardex
                Set ivk = gc.IVKardex(j)
                ivk.cantidad = .Fields("Exist")
                ivk.CodBodega = .Fields("CodBodega")
                ivk.CodInventario = .Fields("CodInventario")
                ivk.CostoRealTotal = c * ivk.cantidad
                ivk.CostoTotal = ivk.CostoRealTotal
                ivk.PrecioRealTotal = ivk.CostoRealTotal
                ivk.PrecioTotal = ivk.PrecioRealTotal
                ivk.Orden = i Mod 100
                i = i + 1
                .MoveNext
            Loop
        End If
        .Close
    End With
        
    If Not (gc Is Nothing) Then
        'Graba la transacción
        MensajeStatus "Grabando la transacción en la empresa '" & gc.Empresa.CodEmpresa & "'...", vbHourglass
        gc.HoraTrans = "00:00:01"
        gc.Grabar False, False
    End If
    
    'Corrige las existencias para que quede bien la tabla 'IVExist'
    MensajeStatus "Arreglando las existencias...", vbHourglass
    If Not (gc Is Nothing) Then
        gc.Empresa.CorregirExistencia
    End If
    mensaje False, "", "OK"
    MensajeStatus
    MsgBox "El proceso terminó con éxito.", vbInformation
    SaldoIV = True
    
cancelado:
    mensaje False, "", Err.Description
    MensajeStatus
    Set ivk = Nothing
    Set iv = Nothing
    Set gc = Nothing
    Set rs = Nothing
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
    
    'Vuelve a abrir la empresa origen
    Set e = gobjMain.RecuperaEmpresa(codOrig)
    e.Abrir
    Set e = Nothing
    
    mbooProcesando = False                  'Desbloquea que se cierre la ventana
    Exit Function
errtrap:
    MensajeStatus
    MsgBox Err.Description, vbExclamation
    GoTo cancelado
End Function

Private Function CrearTrans(ByVal emp As Empresa, _
                            ByVal CodTrans As String, _
                            ByVal Desc As String, _
                            ByVal fecha As Date, _
                            ByVal numdoc As String) As GNComprobante
    Dim g As GNComprobante
    
    Set g = emp.CreaGNComprobante(CodTrans)
    With g
        .IdResponsable = 1
        .CodMoneda = "USD"
        .Cotizacion("USD") = 1   'Diego 20/08/2002
        .Descripcion = Desc
        .FechaTrans = fecha + 1
        .numDocRef = numdoc
    End With
    Set CrearTrans = g
    Set g = Nothing
End Function

'Agrega un detalle de PCKardex a GNComprobante
'Si comprobante llega a tener 100 detalles,
'Graba lo anterior y crea otra instancia
Private Function PrepararTransPC(ByVal e As Empresa, _
                            ByVal codt As String, _
                            ByVal Desc As String, _
                            ByVal Fcorte As Date, _
                            ByRef gc As GNComprobante) As PCKardex
    Dim j As Long, limiteFila As Integer
                            
    
    'Crea transaccion si no existe todavía
    If gc Is Nothing Then
        Set gc = CrearTrans(e, codt, Desc, Fcorte, "")
    End If
    
    If gc.GNTrans.IVNumFilaMax = 0 Then
        limiteFila = 100
    Else
        limiteFila = gc.GNTrans.IVNumFilaMax
    End If
    
    If gc.CodTrans = "CLND" Then
    
        limiteFila = 1
    End If
    
    'Si llega a tener 100 detalles
    If gc.CountPCKardex >= limiteFila Then
        'Graba la transacción
        MensajeStatus "Grabándo la transacción...", vbHourglass
        
        If limiteFila = 1 Then
            gc.CodClienteRef = gc.PCKardex(1).CodProvCli
            If Len(gc.PCKardex(1).CodVendedor) > 0 Then
                gc.CodVendedor = gc.PCKardex(1).CodVendedor
                
            End If
        End If
            
        
        gc.HoraTrans = "00:00:01"
        gc.Grabar False, False
        
        'Crea nueva instancia de GNComprobante
        Set gc = CrearTrans(e, codt, Desc, Fcorte, "")
    End If

    'Agrega detalle
    j = gc.AddPCKardex
    Set PrepararTransPC = gc.PCKardex(j)
End Function


'5. Pasar saldo inicial de proveedores/clientes
Private Function SaldoPC() As Boolean
    Dim e As Empresa, pck As PCKardex
    Dim j As Long, sql As String, rs As Recordset
    Dim i As Long, c As Currency, Fcorte As Date
    Dim gcPVNC As GNComprobante, gcPVND As GNComprobante
    Dim gcCLNC As GNComprobante, gcCLND As GNComprobante
    On Error GoTo errtrap
    
    'Verifica las opciones
    If Not VerificarOpcion Then Exit Function
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    Fcorte = dtpFechaCorte.value    'Fecha de corte

    'Cambia figura de cursor de mouse
    MensajeStatus "Está preparando saldos a la fecha de corte...", vbHourglass
    mensaje True, "Saldo inicial de proveedor/cliente..."
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True

    'Obtiene Saldos de proveedor/cliente por cada documento pendiente
    sql = "spConsPCSaldo3 2, " & FechaYMD(Fcorte, gobjMain.EmpresaActual.TipoDB)
    Set rs = mEmpOrigen.OpenRecordset(sql)
    UltimoRecordset rs
    'Abre la empresa destino
    Set e = AbrirDestino
    With rs
        If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
        i = 0
        Do Until .EOF
            prg1.value = rs.AbsolutePosition
            prg1.Refresh
            DoEvents
            MensajeStatus "Agregando detalle: #" & i & " de " & rs.RecordCount, vbHourglass
            
            'Si aplastó 'Cancelar'
            If mbooCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                GoTo cancelado
            End If
            
            'Si es Proveedores por cobrar (Anticipado)
            If .Fields("Saldo") > 0 And .Fields("BandProveedor") <> 0 Then
                Set pck = PrepararTransPC(e, "PVND", _
                            "Saldo inicial de proveedores (Anticipos)", _
                            Fcorte, gcPVND)
            'Si es Proveedores por pagar
            ElseIf .Fields("Saldo") < 0 And .Fields("BandProveedor") <> 0 Then
                Set pck = PrepararTransPC(e, "PVNC", _
                            "Saldo inicial de proveedores x pagar", _
                            Fcorte, gcPVNC)
            'Si es Clientes por cobrar
            ElseIf .Fields("Saldo") > 0 And .Fields("BandProveedor") = 0 Then
                Set pck = PrepararTransPC(e, "CLND", _
                            "Saldo inicial de clientes x cobrar", _
                            Fcorte, gcCLND)
'                            If .Fields("idEmpleadoref") <> 0 Then
'                                MsgBox "para"
'                            End If
'                            gcCLND.IdEmpleadoRef = .Fields("idEmpleadoref")
            'Si es Clientes por pagar (Anticipado)
            Else
                Set pck = PrepararTransPC(e, "CLNC", _
                            "Saldo inicial de cliente (Anticipos)", _
                            Fcorte, gcCLNC)
            End If
            'If i = 399 Then MsgBox ""
            'Recupera datos de proveedor y asigna al objeto
            pck.idGestionc = .Fields("id") 'yolita reasignacion de gestion de cartera
            pck.CodProvCli = .Fields("CodProvCli")
            If .Fields("Saldo") > 0 Then   'Si es por cobrar --> Debe
                pck.Debe = .Fields("Saldo")        'Saldo en dólares
            Else                                    'Si es por pagar --> Haber
                pck.Haber = .Fields("Saldo") * -1     'Saldo en dólares
            End If
            pck.codforma = .Fields("CodForma")
            pck.FechaEmision = .Fields("FechaEmision")
            pck.FechaVenci = .Fields("FechaVenci")
            pck.NumLetra = .Fields("Trans")
            pck.Observacion = .Fields("Observacion")
            pck.Orden = i
            pck.Guid = .Fields("Guid")        '*** <== AGREGAR ESTO
            If Not IsNull(.Fields("CodVendedor")) Then
                pck.CodVendedor = .Fields("codvendedor") 'AUC 04/06/07
                pck.GNComprobante.CodVendedor = .Fields("codvendedor")
            End If
            If Not IsNull(.Fields("codbanco")) Then
                pck.codBanco = .Fields("codbanco")
            End If
            If Not IsNull(.Fields("codtarjeta")) Then
                pck.CodTarjeta = .Fields("codtarjeta")
            End If
            If Not IsNull(.Fields("Numcheque")) Then
                pck.Numcheque = .Fields("Numcheque")
            End If
            If Not IsNull(.Fields("Numcuenta")) Then
                pck.NumCuenta = .Fields("Numcuenta")
            End If
            If Not IsNull(.Fields("TitularCta")) Then
                pck.TitularCta = .Fields("TitularCta")
            End If
            
            If .Fields("Saldo") > 0 And .Fields("BandProveedor") = 0 Then
                If .Fields("modulo") = "IV" Then
                    If Not IsNull(.Fields("IDTransOrigen")) Then
                            pck.GNComprobante.idTransFuente = .Fields("IDTransOrigen")
                    End If
                    If Not IsNull(.Fields("ValorOriginal")) Then
                        pck.GNComprobante.Atencion = .Fields("ValorOriginal")
                    End If
                    pck.GNComprobante.numDocRef = .Fields("TransSRI")
                Else
                    pck.GNComprobante.numDocRef = .Fields("Trans")
                End If
                If Not IsNull(.Fields("nombre")) Then
                    pck.GNComprobante.nombre = .Fields("nombre")
                End If
            Else
                pck.GNComprobante.numDocRef = .Fields("Trans")
            End If
            i = i + 1
            .MoveNext
        Loop
        .Close
    End With
    
    'Graba la transacción si no están grabadas
    MensajeStatus "Grabándo la transacción...", vbHourglass
    If Not (gcPVND Is Nothing) Then
        gcPVND.HoraTrans = "00:00:01"
        gcPVND.Grabar False, False
        
    End If
    If Not (gcPVNC Is Nothing) Then
        gcPVNC.HoraTrans = "00:00:01"
        gcPVNC.Grabar False, False
    End If
    If Not (gcCLND Is Nothing) Then
        gcCLND.HoraTrans = "00:00:01"
        
        
        
        gcCLND.Grabar False, False
    End If
    If Not (gcCLNC Is Nothing) Then
        gcCLNC.HoraTrans = "00:00:01"
        gcCLNC.Grabar False, False
    End If

    MensajeStatus
    mensaje False, "", "OK"
    MsgBox "El proceso terminó con éxito.", vbInformation
    SaldoPC = True
    
cancelado:
    Set rs = Nothing
    MensajeStatus
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
    
    'Libera los objetos utilizados
    Set pck = Nothing
    Set gcPVND = Nothing
    Set gcPVNC = Nothing
    Set gcCLND = Nothing
    Set gcCLNC = Nothing
    Set e = Nothing
    
    mbooProcesando = False               'Desbloquea que se cierre la ventana
    Exit Function
errtrap:
    mensaje False, "", Err.Description
    MensajeStatus
    DispErr
    GoTo cancelado
End Function



Private Sub mGrupo_Procesando(ByVal msg As String)
    If Len(msg) > 0 Then
        MensajeStatus msg, vbHourglass
    Else
        MensajeStatus
    End If
    DoEvents
End Sub

Private Sub pasa_Click()
Dim r As Boolean, res As Integer
'copia los CUENTASDEPARTAMENTOS EN CUENTASPERSONAL
        res = MsgBox("Cuentas para Asientos DEP->PERSONA", vbYesNo)
        If res = vbYes Then
           r = CuentasRol
           r = CuentasRolPre
        End If
End Sub

Private Sub txtDestino_LostFocus()
    Dim Cancel As Boolean
    'Este es necesario porque al dar Enter no se genera el evento Validate
    txtDestino_Validate Cancel
    If Cancel Then txtDestino.SetFocus
End Sub

Private Sub txtDestino_Validate(Cancel As Boolean)
    Dim e As Empresa, cod As String
    On Error GoTo errtrap
    
    cod = Trim$(txtDestino.Text)
    Set e = gobjMain.RecuperaEmpresa(cod)
    If Not (e Is Nothing) Then
        txtDestinoBD.Text = e.NombreDB
    End If
    Exit Sub
errtrap:
    MsgBox "No se encuentra la empresa de destino. ('" & cod & "')", vbInformation
    Cancel = True
    Exit Sub
End Sub



Private Function Respaldar() As Boolean
    Dim i As Long, nombre As String
    Dim Carpeta As String, sql As String, Indice As Integer
    Dim Seleccionado As Boolean, bases As String, Ruta As String
    Dim pos As Integer

    On Error GoTo errtrap
    
   
    
    Me.MousePointer = vbHourglass
    
      
    mensaje True, "Respaldando Empresa Origen ..."
    nombre = gobjMain.EmpresaActual.NombreDB
    pos = InStrRev(gobjMain.EmpresaActual.Ruta, "\")
    Ruta = Mid$(gobjMain.EmpresaActual.Ruta, 1, pos - 1)
    pos = InStrRev(Ruta, "\")
    Ruta = Mid$(Ruta, 1, pos)
    Carpeta = Ruta & "Backup\"
    RutaRespaldo = Carpeta
'    lblMensaje.Caption = "Respaldándo " & Nombre & _
                            " a " & Carpeta & "..."
    EjecutarSQLResp nombre, Carpeta
    
'    lblMensaje.Caption = ""
    Me.MousePointer = vbNormal
    
    MsgBox "Se completó el respaldo exitósamente.", vbInformation
    Respaldar = True
    mensaje False, "", "OK"
    Exit Function
errtrap:
    Respaldar = False
    Me.MousePointer = vbNormal
    MsgBox Err.Description, vbExclamation
    Exit Function
End Function

Sub EjecutarSQLResp(ByVal nombre As String, ByVal Carpeta As String)
    Dim sql As String, NumReg As Long
        Dim nombreConjunto As String, Destino As String
            
    nombreConjunto = CrearNombreConjunto("Copia de seguridad de %BD", nombre)
    Destino = Trim$(Carpeta) & nombre & "_BAK"

    sql = "BACKUP DATABASE [" & nombre & "] " & _
            "TO  DISK = N'" & Destino & "' WITH  INIT , " & _
            "NOUNLOAD , " & _
            "Name = N'" & nombreConjunto & "', " & _
            "SKIP , STATS = 10, FORMAT "
    'mCn.Execute sql
    gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
End Sub


Private Function CrearNombreConjunto(ByVal nc As String, ByVal bd As String) As String
    Dim t As String, i As Integer
    
    i = InStr(UCase$(nc), "%BD")
    If i > 1 Then
        t = Left$(nc, i - 1) & bd
    End If
    CrearNombreConjunto = t
End Function

Private Function Restaurar() As Boolean
    Dim txtdatos As String, txtregistro As String
    Dim v As Variant, logYfis(0 To 1, 0 To 1) As Variant, i As Integer, j As Integer
    On Error GoTo errtrap
        Me.MousePointer = vbHourglass
        mensaje True, "Creando Empresa Destino... "
        v = InfMedio(lblOrigenBD.Caption, RutaRespaldo)
        If Not IsEmpty(v) Then
            For i = 0 To 1
                For j = 0 To 1
                    logYfis(i, j) = v(i, j)
                Next j
            Next i
            grd1.LoadArray logYfis
            AjustarAutoSize grd1, -1, -1
        End If
         txtdatos = CambiaNombreBaseDatos(grd1.TextMatrix(1, 1)) & txtDestinoBD.Text & ".mdf"
         txtregistro = CambiaNombreBaseDatos(grd1.TextMatrix(2, 1)) & txtDestinoBD.Text & ".ldf"
        RestaurarDB txtDestinoBD.Text, RutaRespaldo & lblOrigenBD.Caption & "_BAK", grd1.TextMatrix(1, 0), grd1.TextMatrix(2, 0), txtdatos, txtregistro
        MsgBox "La restauración concluyó satisfactoriamente"
        Me.MousePointer = vbDefault
        Restaurar = True
        mensaje False, "", "OK"
    Exit Function
errtrap:
    Restaurar = False
    Me.MousePointer = vbDefault
    MsgBox Err.Number & " " & Err.Description, vbInformation
    Exit Function
End Function

Public Sub RestaurarDB(ByVal BaseDatos As String, ByVal Ruta As String, Logico1 As String, Logico2 As String, Fisico1 As String, Fisico2 As String)
    Dim sql As String, NumReg As Long
    
'Restore DataBase $$$$$$$
'From Disk = '$$$$$$'
'WITH REPLACE,
'Move '$1' TO '$1',
'Move '$2' TO '$2'

    On Error GoTo errtrap
    sql = "Restore DataBase " & BaseDatos & _
          " From Disk='" & Ruta & "' " & _
          " With Replace"
    If Len(Logico1) > 0 And Len(Logico2) > 0 _
        And Len(Fisico1) > 0 And Len(Fisico2) > 0 Then
    sql = sql & ", MOVE '" & Logico1 & "' TO '" & Fisico1 & "', " & _
                  "MOVE '" & Logico2 & "' TO '" & Fisico2 & "'"
    End If
    'mProps.Conexion.Execute sql
    gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
    Exit Sub
    
errtrap:
    'Err.Raise ERR_BASENOVALIDA, "Modulo de Restaurar Base", Err.Description
    Exit Sub
End Sub


Public Function InfMedio(ByVal nombre As String, ByVal Carpeta As String) As Variant
    Dim sql As String, rs As Recordset, medio As String
    medio = Trim$(Carpeta) & nombre & "_BAK"
    sql = "RESTORE Filelistonly From Disk='" & medio & "'"
    
    On Error GoTo errtrap
    Set rs = New Recordset
    rs.CursorLocation = adUseClient
 
    'rs.Open sql, mProps.Conexion, adOpenStatic, adLockReadOnly
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    UltimoRecordset rs
    InfMedio = MiGetRows(rs)
    
    rs.Close
    Set rs = Nothing
    Exit Function
    
errtrap:
'    Err.Raise ERR_NOMBREOBJETO, "Medio", Err.Description
    Exit Function
End Function

Private Function CambiaNombreBaseDatos(texto As String) As String
Dim pos As Integer
pos = InStrRev(texto, "\")
CambiaNombreBaseDatos = Mid$(texto, 1, pos)

End Function

'8. Borrat trans. existentes con la fecha anterior a la fecha de corte en base Destino
Private Function BorraTrans() As Boolean
    Dim i As Long, codt As String, numt As Long
    Dim empDestino As Empresa, sql As String, Num As Long
    'Verifica  errores  en la base de Origen
    
    Set empDestino = AbrirDestino
    If empDestino.NombreDB = mEmpOrigen.NombreDB Then
        MsgBox "La empresa origen y destino son las mismas" & Chr(13) & _
               "debera  seleccionar  una empresa de  destino diferente", vbExclamation
        Exit Function
    End If
    If grd.ColKey(grd.Cols - 1) <> "Resultado" Then
    End If
    'If grd.FixedRows = grd.Rows Then CargaTrans 'Carga  trans solo  si la grlla esta vacia
    'Transferir  transaccion una por una
    If MsgBox("Este proceso tardará  algunos minutos " & Chr(13) & " Desea comenzar el proceso de eliminación?", _
                vbYesNo + vbQuestion) <> vbYes Then Exit Function
    
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True
    
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    MensajeStatus "Borrando...", vbHourglass
    BorraTransacciones
    'Corregir  error de Idasignado
    MensajeStatus "Reasignando relaciones ...", vbHourglass
    sql = " UPDATE b SET b.IdAsignado = c.Id " & _
           " From    " & _
           empDestino.NombreDB & ".dbo.PCKardex c INNER JOIN " & _
           mEmpOrigen.NombreDB & ".dbo.PCKardex a INNER JOIN " & empDestino.NombreDB & ".dbo.PCKardex b " & _
           " ON a.Id  = b.IdAsignado " & _
           " ON c.Guid = a.Guid " & _
           " Where a.idAsignado = 0 And b.idAsignado <> 0 And c.idAsignado = 0 "
    
    mEmpOrigen.EjecutarSQL sql, Num
    
    
    
    sql = " Update b"
    sql = sql & " Set b.IdAsignadoPCK = c.id "
    sql = sql & " from "
    sql = sql & empDestino.NombreDB & ".dbo.PCKardex c "
    sql = sql & " INNER JOIN " & mEmpOrigen.NombreDB & ".dbo.PCKardex a "
    sql = sql & " INNER JOIN " & empDestino.NombreDB & ".dbo.PCKardexCHP B "
    sql = sql & " ON a.Id  = b.IdAsignadoPCK "
    sql = sql & " ON c.Guid = a.Guid "
    sql = sql & " Where a.idAsignado = 0 And b.IdAsignadoPCK <> 0 And c.idAsignado = 0"
    
    mEmpOrigen.EjecutarSQL sql, Num
    
    MsgBox "Proceso terminado con exito"
    
    
    MensajeStatus
    mbooProcesando = False  'Bloquea que se cierre la ventana
    BorraTrans = True
cancelado:
    MensajeStatus
    mbooProcesando = False
    prg1.value = prg1.min
    Exit Function
errtrap:
    MensajeStatus
    DispErr
    mbooProcesando = False
    prg1.value = prg1.min
    Exit Function
End Function


Private Sub BorraTransacciones()
    Dim sql As String, rs As Recordset, v As Variant, i As Integer
    Dim mesini As Integer, mesfin As Integer, mesTotal As Integer
    Dim anioini As Integer, aniofin As Integer, anioTotal As Integer
    Dim diaini As Integer, diafin As Integer, diaTotal As Integer
    Dim Fcorte As Date, n As Long
    On Error GoTo errtrap
    
    diaini = DatePart("d", gobjMain.EmpresaActual.GNOpcion.FechaInicio)
    mesini = DatePart("m", gobjMain.EmpresaActual.GNOpcion.FechaInicio)
    anioini = DatePart("yyyy", gobjMain.EmpresaActual.GNOpcion.FechaInicio)

    diafin = DatePart("d", dtpFechaCorte.value)
    mesfin = DatePart("m", dtpFechaCorte.value)
    aniofin = DatePart("yyyy", dtpFechaCorte.value)
    
    anioTotal = aniofin - anioini
    mesTotal = (mesfin - mesini) + anioTotal * 12
    diaTotal = DateDiff("d", gobjMain.EmpresaActual.GNOpcion.FechaInicio, dtpFechaCorte.value)
    
    Fcorte = dtpFechaCorte.value    'Fecha de corte
    'Selecciona las transacciones de la  base de origen
    prg1.min = 0
    prg1.max = diaTotal
    'For i = mesTotal To 0 Step -1
    MensajeStatus "Moviendo Depreciaciones Anteriores...", vbHourglass
    sql = "UPDATE [" & Trim$(txtDestinoBD.Text) & "].dbo.GNCOMPROBANTE SET FECHATRANS='" & DateAdd("d", 1, dtpFechaCorte.value) & "' WHERE ESTADO <>3 AND CODTRANS IN ('AFSI','RDPAF')"
    gobjMain.EmpresaActual.EjecutarSQL sql, n
    
    MensajeStatus "Borrando... pckardex", vbHourglass
    For i = 0 To diaTotal Step 5
        prg1.value = i
        prg1.Refresh
        
        MensajeStatus "Borrando... pckardex fecha: " & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio), vbHourglass
        
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.pckardex  from [" & Trim$(txtDestinoBD.Text) & "].dbo.pckardex"
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.transid=[" & Trim$(txtDestinoBD.Text) & "].dbo.pckardex.transid"
        sql = sql & " where fechatrans<='" & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio) & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n
    Next i
    
    
    sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.pckardex  from [" & Trim$(txtDestinoBD.Text) & "].dbo.pckardex"
    sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
    sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.transid=[" & Trim$(txtDestinoBD.Text) & "].dbo.pckardex.transid"
    sql = sql & " where fechatrans<='" & dtpFechaCorte.value & "'"
    gobjMain.EmpresaActual.EjecutarSQL sql, n
     
    MensajeStatus "Borrando... ivkardex", vbHourglass
    For i = 0 To diaTotal Step 5
        prg1.value = i
        prg1.Refresh
        
        MensajeStatus "Borrando... ivkardex fecha: " & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio), vbHourglass
        
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.ivkardex from [" & Trim$(txtDestinoBD.Text) & "].dbo.ivkardex"
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.transid=[" & Trim$(txtDestinoBD.Text) & "].dbo.ivkardex.transid"
        sql = sql & " where fechatrans<='" & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio) & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n
        
    Next i
    
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.ivkardex from [" & Trim$(txtDestinoBD.Text) & "].dbo.ivkardex"
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.transid=[" & Trim$(txtDestinoBD.Text) & "].dbo.ivkardex.transid"
        sql = sql & " where fechatrans<='" & dtpFechaCorte.value & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n
    
    
    MensajeStatus "Borrando... afkardex", vbHourglass
    For i = 0 To diaTotal Step 5
        prg1.value = i
        prg1.Refresh
        
        MensajeStatus "Borrando... afkardex fecha: " & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio), vbHourglass
        
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.afkardex from [" & Trim$(txtDestinoBD.Text) & "].dbo.afkardex"
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.transid=[" & Trim$(txtDestinoBD.Text) & "].dbo.afkardex.transid"
        sql = sql & " where fechatrans<='" & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio) & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n
        
    Next i
    
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.afkardex from [" & Trim$(txtDestinoBD.Text) & "].dbo.afkardex"
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.transid=[" & Trim$(txtDestinoBD.Text) & "].dbo.afkardex.transid"
        sql = sql & " where fechatrans<='" & dtpFechaCorte.value & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n
    
    
    
    'For i = mesTotal To 0 Step -1
    MensajeStatus "Borrando... ivkardexrecargo", vbHourglass
    For i = 0 To diaTotal Step 5
        prg1.value = i
        prg1.Refresh
        
        MensajeStatus "Borrando... ivkardexrecargo fecha: " & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio), vbHourglass
        
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.ivkardexrecargo from [" & Trim$(txtDestinoBD.Text) & "].dbo.ivkardexrecargo"
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.transid=[" & Trim$(txtDestinoBD.Text) & "].dbo.ivkardexrecargo.transid"
        sql = sql & " where fechatrans<='" & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio) & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n
    
    Next i

        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.ivkardexrecargo from [" & Trim$(txtDestinoBD.Text) & "].dbo.ivkardexrecargo"
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.transid=[" & Trim$(txtDestinoBD.Text) & "].dbo.ivkardexrecargo.transid"
        sql = sql & " where fechatrans<='" & dtpFechaCorte.value & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n


    MensajeStatus "Borrando... ivfinanciamientoitem", vbHourglass
    For i = 0 To diaTotal Step 5
        prg1.value = i
        prg1.Refresh
        
        MensajeStatus "Borrando... ivfinanciamientoitem fecha: " & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio), vbHourglass
        
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.ivfinanciamientoitem from [" & Trim$(txtDestinoBD.Text) & "].dbo.ivfinanciamientoitem "
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.transid=[" & Trim$(txtDestinoBD.Text) & "].dbo.ivfinanciamientoitem.transid"
        sql = sql & " where fechatrans<='" & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio) & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n
    
    Next i

        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.ivfinanciamientoitem from [" & Trim$(txtDestinoBD.Text) & "].dbo.ivfinanciamientoitem "
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.transid=[" & Trim$(txtDestinoBD.Text) & "].dbo.ivfinanciamientoitem.transid"
        sql = sql & " where fechatrans<='" & dtpFechaCorte.value & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n


    
    'For i = mesTotal To 0 Step -1
    MensajeStatus "Borrando... ctlibrodetalle", vbHourglass
    For i = 0 To diaTotal Step 5
        prg1.value = i
        prg1.Refresh
        
        MensajeStatus "Borrando... ctlibrodetalle fecha: " & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio), vbHourglass
        
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.ctlibrodetalle from [" & Trim$(txtDestinoBD.Text) & "].dbo.ctlibrodetalle"
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.codasiento=[" & Trim$(txtDestinoBD.Text) & "].dbo.ctlibrodetalle.codasiento"
        sql = sql & " where fechatrans<='" & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio) & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n
    Next i
    
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.ctlibrodetalle from [" & Trim$(txtDestinoBD.Text) & "].dbo.ctlibrodetalle"
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.codasiento=[" & Trim$(txtDestinoBD.Text) & "].dbo.ctlibrodetalle.codasiento"
        sql = sql & " where fechatrans<='" & dtpFechaCorte.value & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n
    
    
    MensajeStatus "Borrando... anexos", vbHourglass
    For i = 0 To diaTotal Step 5
    'For i = mesTotal To 0 Step -1
        prg1.value = i
        prg1.Refresh
        
        MensajeStatus "Borrando... anexos fecha: " & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio), vbHourglass
        
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.anexos from [" & Trim$(txtDestinoBD.Text) & "].dbo.anexos"
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.transid=[" & Trim$(txtDestinoBD.Text) & "].dbo.anexos.transid"
        sql = sql & " where fechatrans<='" & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio) & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n
    Next i
    
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.anexos from [" & Trim$(txtDestinoBD.Text) & "].dbo.anexos"
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.transid=[" & Trim$(txtDestinoBD.Text) & "].dbo.anexos.transid"
        sql = sql & " where fechatrans<='" & dtpFechaCorte.value & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n
    
    
    'For i = mesTotal To 0 Step -1
    MensajeStatus "Borrando... tskardex", vbHourglass
    For i = 0 To diaTotal Step 5
        prg1.value = i
        prg1.Refresh
        
        MensajeStatus "Borrando... tskardex fecha: " & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio), vbHourglass
        
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.tskardex from [" & Trim$(txtDestinoBD.Text) & "].dbo.tskardex"
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.transid=[" & Trim$(txtDestinoBD.Text) & "].dbo.tskardex.transid"
        sql = sql & " where fechatrans<='" & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio) & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n
    Next i
    
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.tskardex from [" & Trim$(txtDestinoBD.Text) & "].dbo.tskardex"
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.transid=[" & Trim$(txtDestinoBD.Text) & "].dbo.tskardex.transid"
        sql = sql & " where fechatrans<='" & dtpFechaCorte.value & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n
    
    
    MensajeStatus "Borrando... tskardexret", vbHourglass
    For i = 0 To diaTotal Step 5
        prg1.value = i
        prg1.Refresh
        
        MensajeStatus "Borrando... tskardexret fecha: " & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio), vbHourglass
        
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.tskardexret from [" & Trim$(txtDestinoBD.Text) & "].dbo.tskardexret"
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.transid=[" & Trim$(txtDestinoBD.Text) & "].dbo.tskardexret.transid"
        sql = sql & " where fechatrans<='" & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio) & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n
    Next i

        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.tskardexret from [" & Trim$(txtDestinoBD.Text) & "].dbo.tskardexret"
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.transid=[" & Trim$(txtDestinoBD.Text) & "].dbo.tskardexret.transid"
        sql = sql & " where fechatrans<='" & dtpFechaCorte.value & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n


    MensajeStatus "Borrando... tskardexconcilia", vbHourglass
    For i = 0 To diaTotal Step 5
        prg1.value = i
        prg1.Refresh
        
        MensajeStatus "Borrando... tskardexconcilia fecha: " & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio), vbHourglass
        
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.tskardexconcilia from [" & Trim$(txtDestinoBD.Text) & "].dbo.tskardexconcilia"
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.transid=[" & Trim$(txtDestinoBD.Text) & "].dbo.tskardexconcilia.transid"
        sql = sql & " where fechatrans<='" & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio) & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n
    Next i
    
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.tskardexconcilia from [" & Trim$(txtDestinoBD.Text) & "].dbo.tskardexconcilia"
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.transid=[" & Trim$(txtDestinoBD.Text) & "].dbo.tskardexconcilia.transid"
        sql = sql & " where fechatrans<='" & dtpFechaCorte.value & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n


    MensajeStatus "Borrando... GNOferta", vbHourglass
    For i = 0 To diaTotal Step 5
        prg1.value = i
        prg1.Refresh
        
        MensajeStatus "Borrando... GNOferta fecha: " & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio), vbHourglass
        
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.GNOferta from [" & Trim$(txtDestinoBD.Text) & "].dbo.GNOferta"
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.transid=[" & Trim$(txtDestinoBD.Text) & "].dbo.GNOferta.transid"
        sql = sql & " where fechatrans<='" & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio) & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n
    Next i
    
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.GNOferta from [" & Trim$(txtDestinoBD.Text) & "].dbo.GNOferta "
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.transid=[" & Trim$(txtDestinoBD.Text) & "].dbo.GNOferta.transid"
        sql = sql & " where fechatrans<='" & dtpFechaCorte.value & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n

    MensajeStatus "Borrando... GNFuente", vbHourglass
    For i = 0 To diaTotal Step 5
        prg1.value = i
        prg1.Refresh
        
        MensajeStatus "Borrando... GNFuente fecha: " & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio), vbHourglass
        
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.GNFuente from [" & Trim$(txtDestinoBD.Text) & "].dbo.GNFuente "
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.transid=[" & Trim$(txtDestinoBD.Text) & "].dbo.GNFuente.transid"
        sql = sql & " where fechatrans<='" & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio) & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n
    Next i
    
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.GNFuente from [" & Trim$(txtDestinoBD.Text) & "].dbo.GNFuente "
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.transid=[" & Trim$(txtDestinoBD.Text) & "].dbo.GNFuente.transid"
        sql = sql & " where fechatrans<='" & dtpFechaCorte.value & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n

    MensajeStatus "Borrando... Infocomprobante", vbHourglass
    For i = 0 To diaTotal Step 5
        prg1.value = i
        prg1.Refresh
        
        MensajeStatus "Borrando... Infocomprobantes fecha: " & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio), vbHourglass
        
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.Infocomprobantes from [" & Trim$(txtDestinoBD.Text) & "].dbo.Infocomprobantes "
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.transid=[" & Trim$(txtDestinoBD.Text) & "].dbo.Infocomprobantes.transid"
        sql = sql & " where fechatrans<='" & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio) & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n
    Next i
    
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.Infocomprobantes from [" & Trim$(txtDestinoBD.Text) & "].dbo.Infocomprobantes "
        sql = sql & " inner join [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " on [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante.transid=[" & Trim$(txtDestinoBD.Text) & "].dbo.Infocomprobantes.transid"
        sql = sql & " where fechatrans<='" & dtpFechaCorte.value & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n



    MensajeStatus "Borrando... gncomprobante", vbHourglass
    For i = 0 To diaTotal Step 5
    'For i = mesTotal To 0 Step -1
        prg1.value = i
        prg1.Refresh
        
        MensajeStatus "Borrando... gncomprobante fecha: " & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio), vbHourglass
        
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " where fechatrans<='" & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio) & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n
    Next i
        
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.gncomprobante"
        sql = sql & " where fechatrans<='" & dtpFechaCorte.value & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n
    
    
    MensajeStatus "Borrando... regauditoria", vbHourglass
    For i = 0 To diaTotal Step 5
    'For i = mesTotal To 0 Step -1
        prg1.value = i
        prg1.Refresh
        
        MensajeStatus "Borrando... regauditoria fecha: " & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio), vbHourglass
        
        sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.regauditoria"
        sql = sql & " where fechaHora<='" & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio) & "'"
        gobjMain.EmpresaActual.EjecutarSQL sql, n
    Next i
    
    sql = "Delete [" & Trim$(txtDestinoBD.Text) & "].dbo.regauditoria"
    sql = sql & " where fechaHora<='" & dtpFechaCorte.value & "'"
    gobjMain.EmpresaActual.EjecutarSQL sql, n

    
    prg1.min = 0
    MensajeStatus "Corriguiendo Existencias", vbHourglass
    gobjMain.EmpresaActual.CorregirExistenciaBaseNueva "[" & Trim$(txtDestinoBD.Text) & "].dbo."
    
    Exit Sub
salida:
    MensajeStatus
    Set rs = Nothing
    Exit Sub
errtrap:
    MensajeStatus
    DispErr
    GoTo salida
End Sub


Private Function SaldoAF() As Boolean
    Dim e As Empresa, gc As GNComprobante, afk As AFKardex, af As AFinventario
    Dim j As Long, n As Long
    Dim sql As String, rs As Recordset, codOrig As String
    Dim i As Long, c As Currency, Fcorte As Date
    On Error GoTo errtrap
    
    'Verifica las opciones
    If Not VerificarOpcion Then Exit Function
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    
    codOrig = gobjMain.EmpresaActual.CodEmpresa
    Fcorte = dtpFechaCorte.value    'Fecha de corte

    'Cambia figura de cursor de mouse
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True
    
    'Saca las existencias a la fecha de corte
    MensajeStatus "Preparando para grabar las Depreciación de activos fijos...", vbHourglass
    mensaje True, "Saldo inicial de Depreciación de activos fijos..."
    
    sql = "SELECT afk.IdInventario, afk.IdBodega, " & _
                "af.CodInventario, afb.CodBodega, " & _
                "Sum(afk.Cantidad) AS Exist, Sum(afk.costototal) AS ct " & _
          "FROM afBodega afb INNER JOIN " & _
                    "(afInventario af INNER JOIN " & _
                        "(GNTrans gt INNER JOIN " & _
                            "(afKardex afk INNER JOIN GNComprobante gc " & _
                            "ON afk.TransID=gc.TransID) " & _
                        "ON gt.CodTrans=gc.CodTrans) " & _
                    "ON af.IdInventario = afk.IdInventario) " & _
                "ON afb.IdBodega = afk.IdBodega " & _
          "WHERE (gc.Estado<>" & ESTADO_ANULADO & ") AND " & _
                 "(gt.AfectaCantidad=" & CadenaBool(True, gobjMain.EmpresaActual.TipoDB) & ") AND " & _
                 "(gc.FechaTrans < " & FechaYMD(Fcorte + 1, gobjMain.EmpresaActual.TipoDB) & ") AND gc.codtrans = 'DEPAF' AND " & _
                 "(af.BandServicio<>" & CadenaBool(True, gobjMain.EmpresaActual.TipoDB) & ") " & _
          "GROUP BY afk.IdInventario, afk.IdBodega, af.CodInventario, afb.CodBodega " & _
          "HAVING Sum(afk.Cantidad)<>0"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
#If DAOLIB = 0 Then
    Set rs.ActiveConnection = Nothing
#End If
    
    'Abre la empresa destino
    Set e = AbrirDestino
    
    
    With rs
        If Not rs.EOF Then
            rs.MoveLast
            rs.MoveFirst
            If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
            i = 0
            Do Until .EOF
                prg1.value = rs.AbsolutePosition
                prg1.Refresh
                DoEvents
                
                'Si aplastó 'Cancelar'
                If mbooCancelado Then
                    MsgBox "El proceso fue cancelado.", vbInformation
                    GoTo cancelado
                End If
                
                'Crea transaccion 'AFSI'
                If (i Mod 100) = 0 Then
                    'Si no es primera vez
                    If Not (gc Is Nothing) Then
                        'Graba la transacción
                        MensajeStatus "Grabando la transacción en la empresa '" & gc.Empresa.CodEmpresa & "'...", vbHourglass
                        gc.FechaTrans = CDate("01/01/" & DatePart("yyyy", Date))
                        gc.Descripcion = "Depreciacion acumulada de Activos Fijos del Año " & DatePart("yyyy", Fcorte)
                        gc.HoraTrans = "00:00:10"
                        gc.Grabar False, False
                    End If
                    
                    Set gc = CrearTrans(e, _
                            "AFSI", _
                            "Depreciacion acumulada de activos fijos", _
                            Fcorte, _
                            "")
                End If
                
                'Recupera datos de inventario para llama el método Costo()
                MensajeStatus "Agregando detalle #" & i & " de " & rs.RecordCount, vbHourglass
                Set af = mEmpOrigen.RecuperaAFInventario(.Fields("IdInventario"))
                
                'Obtiene Costo del item en Moneda de item
                c = af.costo(Fcorte, 1)
                
                'De moneda de item, covierte en moneda de trans, si es necesario
                If af.CodMoneda <> gc.CodMoneda Then
                    c = c * gc.Cotizacion(af.CodMoneda) / gc.Cotizacion("")
                End If
                
                'Agrega detalle
                j = gc.AddAFKardex
                Set afk = gc.AFKardex(j)
                afk.cantidad = .Fields("Exist")
                afk.CodBodega = .Fields("CodBodega")
                afk.CodBodega = .Fields("CodBodega")
                afk.idinventario = .Fields("idInventario")
'                afk.CodInventario = .Fields("CodInventario")
                
                afk.CostoRealTotal = .Fields("ct")
                afk.CostoTotal = .Fields("ct")
                afk.PrecioRealTotal = afk.PrecioRealTotal
                afk.PrecioTotal = afk.PrecioRealTotal
                afk.Orden = i Mod 100
                i = i + 1
                .MoveNext
            Loop
        End If
        .Close
    End With
        
    If Not (gc Is Nothing) Then
        'Graba la transacción
        MensajeStatus "Grabando la transacción en la empresa '" & gc.Empresa.CodEmpresa & "'...", vbHourglass
        gc.FechaTrans = CDate("01/01/" & DatePart("yyyy", Date))
        gc.HoraTrans = "00:00:01"
        gc.Descripcion = "Depreciacion acumulada de activos fijos del Año " & DatePart("yyyy", Fcorte)
        gc.Grabar False, False
    End If
    
    'Corrige las existencias para que quede bien la tabla 'IVExist'
    MensajeStatus "Arreglando las existencias...", vbHourglass
    If Not (gc Is Nothing) Then
        gc.Empresa.CorregirExistencia
    End If
    mensaje False, "", "OK"
    MensajeStatus
    MsgBox "El proceso terminó con éxito.", vbInformation
    SaldoAF = True
    
cancelado:
    mensaje False, "", Err.Description
    MensajeStatus
    Set afk = Nothing
    Set af = Nothing
    Set gc = Nothing
    Set rs = Nothing
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
    
    'Vuelve a abrir la empresa origen
    Set e = gobjMain.RecuperaEmpresa(codOrig)
    e.Abrir
    Set e = Nothing
    
    mbooProcesando = False                  'Desbloquea que se cierre la ventana
    Exit Function
errtrap:
    MensajeStatus
    MsgBox Err.Description, vbExclamation
    GoTo cancelado
End Function

Private Function SaldoInicialAF() As Boolean
    Dim e As Empresa, gc As GNComprobante, afk As AFKardex, af As AFinventario
    Dim j As Long, n As Long
    Dim sql As String, rs As Recordset, codOrig As String
    Dim i As Long, c As Currency, Fcorte As Date
    On Error GoTo errtrap
    
    'Verifica las opciones
    If Not VerificarOpcion Then Exit Function
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    
    codOrig = gobjMain.EmpresaActual.CodEmpresa
    Fcorte = dtpFechaCorte.value    'Fecha de corte

    'Cambia figura de cursor de mouse
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True
    
    'Saca las existencias a la fecha de corte
    MensajeStatus "Preparando para grabar las existencias iniciales de activos fijos...", vbHourglass
    mensaje True, "Saldo inicial de activos fijos..."
    
    sql = "SELECT af.IdInventario, 1 , "
    sql = sql & " af.CodInventario, 'B01' as codbodega, "
    sql = sql & "ISNULL((ISNULL(af.numvidautil,0)- isnull(af.depanterior,0)),0) AS Cant, costoultimoingreso "
    sql = sql & "FROM  afInventario af "
    sql = sql & "WHERE FECHACOMPRA <" & FechaYMD(Fcorte + 1, gobjMain.EmpresaActual.TipoDB)

'    sql = " select a.idinventario, af.codinventario, max('B01') as codbodega ,"
'    sql = sql & " Count (numtrans)as cant , costoultimoingreso "
'    sql = sql & " from gncomprobante g"
'    sql = sql & " inner join afkardex a"
'    sql = sql & " inner join afinventario af"
'    sql = sql & " on a.idinventario=af.idinventario"
'    sql = sql & " on g.transid=a.transid where codtrans='DEPAF'"
'    sql = sql & " and g.estado <>3"
'    sql = sql & " and g.Fechatrans < " & FechaYMD(Fcorte + 1, gobjMain.EmpresaActual.TipoDB)
'    sql = sql & " group by a.idinventario, af.codinventario"


    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
#If DAOLIB = 0 Then
    Set rs.ActiveConnection = Nothing
#End If
    
    'Abre la empresa destino
    Set e = AbrirDestino
    
    With rs
        If Not rs.EOF Then
            rs.MoveLast
            rs.MoveFirst
            If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
            i = 0
            Do Until .EOF
                prg1.value = rs.AbsolutePosition
                prg1.Refresh
                DoEvents
                
                'Si aplastó 'Cancelar'
                If mbooCancelado Then
                    MsgBox "El proceso fue cancelado.", vbInformation
                    GoTo cancelado
                End If
                
                'Crea transaccion 'IVASF'
                If (i Mod 100) = 0 Then
                    'Si no es primera vez
                    If Not (gc Is Nothing) Then
                        'Graba la transacción
                        MensajeStatus "Grabando la transacción en la empresa '" & gc.Empresa.CodEmpresa & "'...", vbHourglass
                        gc.HoraTrans = "00:00:01"
                        gc.Grabar False, False
                    End If
                    
                    Set gc = CrearTrans(e, _
                            "IVSAF", _
                            "Saldo inicial de activos fijos", _
                            Fcorte, _
                            "")
                End If
                
                'Recupera datos de inventario para llama el método Costo()
                MensajeStatus "Agregando detalle #" & i & " de " & rs.RecordCount, vbHourglass
                Set af = mEmpOrigen.RecuperaAFInventario(.Fields("IdInventario"))
                
                'Obtiene Costo del item en Moneda de item
                c = af.costo(Fcorte, 1)
                
                'De moneda de item, covierte en moneda de trans, si es necesario
                If af.CodMoneda <> gc.CodMoneda Then
                    c = c * gc.Cotizacion(af.CodMoneda) / gc.Cotizacion("")
                End If
                
                'Agrega detalle
                j = gc.AddAFKardex
                Set afk = gc.AFKardex(j)
                If .Fields("CANT") = 0 Then
                    afk.cantidad = 1
'                ElseIf .Fields("CANT") = "Nulo" Then
'                    afk.cantidad = 1
                Else
                    afk.cantidad = .Fields("CANT")
                End If
                afk.CodBodega = .Fields("CodBodega")
                afk.CodInventario = .Fields("CodInventario")
                afk.CostoRealTotal = .Fields("costoultimoingreso")
                afk.CostoTotal = .Fields("costoultimoingreso")
                afk.PrecioRealTotal = .Fields("costoultimoingreso")
                afk.PrecioTotal = .Fields("costoultimoingreso")
                afk.Orden = i Mod 100
                i = i + 1
                .MoveNext
            Loop
        End If
        .Close
    End With
        
    If Not (gc Is Nothing) Then
        'Graba la transacción
        MensajeStatus "Grabando la transacción en la empresa '" & gc.Empresa.CodEmpresa & "'...", vbHourglass
        gc.HoraTrans = "00:00:01"
        gc.Grabar False, False
    End If
    
    'Corrige las existencias para que quede bien la tabla 'IVExist'
    MensajeStatus "Arreglando las existencias...", vbHourglass
    If Not (gc Is Nothing) Then
        gc.Empresa.CorregirExistencia
    End If
    mensaje False, "", "OK"
    MensajeStatus
    MsgBox "El proceso terminó con éxito.", vbInformation
    SaldoInicialAF = True
    
cancelado:
    mensaje False, "", Err.Description
    MensajeStatus
    Set afk = Nothing
    Set af = Nothing
    Set gc = Nothing
    Set rs = Nothing
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
    
    'Vuelve a abrir la empresa origen
    Set e = gobjMain.RecuperaEmpresa(codOrig)
    e.Abrir
    Set e = Nothing
    
    mbooProcesando = False                  'Desbloquea que se cierre la ventana
    Exit Function
errtrap:
    MensajeStatus
    MsgBox Err.Description, vbExclamation
    GoTo cancelado
End Function

Private Sub cmdAbrirEmpRol_Click()
    'txtOrigenRoles.Text = frmSelecEmpRol.Inicio
    frmSelecEmpRol.Show vbModal
    txtOrigenRoles.Text = gobjRol.EmpresaActual.CodEmpresa
    'lblEmpresaSii.Caption = gobjSii.EmpresaActual.Descripcion
End Sub

Private Sub txtOrigenRoles_LostFocus()
    Dim Cancel As Boolean
    'Este es necesario porque al dar Enter no se genera el evento Validate
    txtDestino_Validate Cancel
    If Cancel Then txtDestino.SetFocus
End Sub
'AUC
Private Function SaldoRoles() As Boolean
    Dim e As Empresa, gc As GNComprobante, ivk As IVKardex, iv As IVinventario
    Dim j As Long, n As Long
    Dim sql As String, rs As Recordset, codOrig As String
    Dim i As Long, c As Currency, Fcorte As Date
    Dim ele As Elementos
    Dim r
    On Error GoTo errtrap
    'Verifica las opciones
'    If Not VerificarOpcion Then Exit Function
    mbooProcesando = True               'Bloquea que se cierre la ventana
    codOrig = gobjMain.EmpresaActual.CodEmpresa
    Fcorte = dtpFechaCorte.value    'Fecha de corte
    'Cambia figura de cursor de mouse
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True
    'Saca las existencias a la fecha de corte
    MensajeStatus "Preparando para grabar Saldos...", vbHourglass
    mensaje True, "Historial de roles..."
    VerificaExistenciaTablaRol 99
    sql = "SELECT datepart(" & "yyyy" & ",r.fechainicio) as fecharol,p.ci,rd.codelemento,  " & _
                "Sum(rd.valor) AS valor INTO Tmp99  " & _
          "FROM ROL R INNER JOIN  RolDetalle rd INNER JOIN " & _
                    "Elementos ele on ele.codelemento = rd.codelemento inner join personal p on p.codempleado = rd.codempleado " & _
                    "on r.idrol = rd.idrol " & _
                    " " & _
          "WHERE  " & _
            "(R.FechaFinal > " & FechaYMD(DateAdd("YYYY", -1, Fcorte), gobjMain.EmpresaActual.TipoDB) & ") AND " & _
            "(R.FechaFinal < " & FechaYMD(Fcorte + 1, gobjMain.EmpresaActual.TipoDB) & ") AND " & _
                 " BandExportaSii = 1 " & _
                 "GROUP BY r.fechainicio,p.ci,rd.codelemento,ele.orden,ele.codelemento" & _
          "  order by p.ci,ele.codelemento"
            gobjRol.EmpresaActual.EjecutarSQL sql, 0
            sql = "select fecharol, ci, codelemento, Sum(valor) AS valor  from tmp99" & _
                  " group by fecharol,ci,codelemento " & _
                  " order by ci,codelemento"
    Set rs = gobjRol.EmpresaActual.OpenRecordset(sql)
#If DAOLIB = 0 Then
    Set rs.ActiveConnection = Nothing
#End If
    'Abre la empresa destino
    Set e = AbrirDestino
    With rs
        If Not rs.EOF Then
            rs.MoveLast
            rs.MoveFirst
            If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
            i = 1
            Do Until .EOF
                prg1.value = rs.AbsolutePosition
                prg1.Refresh
                DoEvents
                'Si aplastó 'Cancelar'
                If mbooCancelado Then
                    MsgBox "El proceso fue cancelado.", vbInformation
                    GoTo cancelado
                End If
                    MensajeStatus "Grabando Historial de Roles en la empresa '" & gobjMain.EmpresaActual.CodEmpresa & "'...", vbHourglass
                    Dim idpc As Long
                    idpc = gobjMain.EmpresaActual.RecuperaIdProvCli(rs!ci)
                  '  If rs!ci = "0103480851" Then
                  '      MsgBox "parar"
                  '  End If
                    Set ele = gobjMain.EmpresaActual.RecuperarElemento(rs!Codelemento)
                        If idpc = 0 Then
                            r = MsgBox("El empleado con cedula " & rs!ci & " No existe Desea continuar con el siguiente ", vbYesNo)
                            If r = vbYes Then
                                GoTo continua
                            Else
                                GoTo cancelado
                            End If
                        End If
                        'AUC quitado para crear el historial de isollanta porque son dos empresas con los mismos elementos
'                    If VerificaRol("31/12/" & rs!fecharol, idpc, ele.idElemento) Then
'                      r = MsgBox("El Historial de Este Empleado ya Existe " & idpc & ", " & ele.codElemento & " Desea continuar con el siguiente ", vbYesNo)
'                        If r = vbYes Then
'                            GoTo continua
'                        Else
'                            GoTo cancelado
'                        End If
'                    End If
                    sql = "INSERT INTO HISTORIALROL (fecharol,IdEmpleado,IdElemento,Orden,Valor)"
                    sql = sql & " VALUES ( '31/12/" & rs!fecharol & "'," & idpc & "," & ele.idElemento & "," & i & "," & rs!valor & ")"
                
                MensajeStatus "Agregando detalle #" & i & " de " & rs.RecordCount, vbHourglass
                gobjMain.EmpresaActual.Execute sql
                i = i + 1
continua:
                rs.MoveNext
                Set ele = Nothing
                idpc = 0
            Loop
        End If
        .Close
    End With
    mensaje False, "", "OK"
    MensajeStatus
    MsgBox "El proceso terminó con éxito.", vbInformation
    SaldoRoles = True
cancelado:
    Set ele = Nothing
    Set rs = Nothing
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
    mbooProcesando = False                  'Desbloquea que se cierre la ventana
    Exit Function
errtrap:
    MensajeStatus
    MsgBox Err.Description, vbExclamation
    GoTo cancelado
End Function
Public Sub VerificaExistenciaTablaRol(i As Integer)
    Dim rs As Recordset
    Dim sql As String
    'verifica  si la  tabla no esta  creada
    sql = "SELECT * FROM sysobjects WHERE NAME =  'tmp" & i & "'"
    Set rs = gobjRol.EmpresaActual.OpenRecordset(sql)
    If Not (rs.EOF And rs.BOF) Then
        'elimina la tabla
        gobjRol.EmpresaActual.EjecutarSQL "drop table Tmp" & i, 0
    End If
End Sub
Private Function SaldoRolesSii() As Boolean
    Dim e As Empresa, gc As GNComprobante, ivk As IVKardex, iv As IVinventario
    Dim j As Long, n As Long
    Dim sql As String, rs As Recordset, codOrig As String
    Dim i As Long, c As Currency, Fcorte As Date
    Dim idSec As Byte
    Dim pc As PCProvCli
    Dim pcg As PCGRUPO
    Dim ele As Elementos
    Dim pcc As PCCanton
    idSec = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("seccion") + 1
    On Error GoTo errtrap
    'Verifica las opciones
'    If Not VerificarOpcion Then Exit Function
    mbooProcesando = True               'Bloquea que se cierre la ventana
    codOrig = gobjMain.EmpresaActual.CodEmpresa
    Fcorte = dtpFechaCorte.value    'Fecha de corte
    'Cambia figura de cursor de mouse
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True
    'Saca las existencias a la fecha de corte
    MensajeStatus "Preparando para grabar Saldos...", vbHourglass
    mensaje True, "Historial de roles..."
    VerificaExistenciaTablaSii 99
    sql = "SELECT datepart(yyyy,gc.fechadevol) as fecharol,pc.codprovcli as codEmpleado, ele.codelemento, pcG.CodGrupo" & idSec & " as codGrupo , " & _
                "pcc.codcanton,rd.bandpagoprov,Sum(rd.valor) AS valor  INTO Tmp99  " & _
          "FROM gncomprobante gc INNER JOIN  RolDetalle rd INNER JOIN Elemento ele  " & _
                    "on ele.idelemento = rd.idElemento on gc.transid = rd.transid  " & _
                    "INNER JOIN Empleado pc on pc.idprovcli = rd.idempleado " & _
                    "INNER JOIN PcGrupo" & idSec & " pcg on pcg.idGrupo" & idSec & "  = rd.idGrupo" & idSec & _
                    " LEFT JOIN PcCanton pcc on pcc.idcanton = rd.idcanton " & _
          "WHERE gc.estado <> 3 AND  " & _
            "(gc.Fechadevol > " & FechaYMD(DateAdd("YYYY", -1, Fcorte), gobjMain.EmpresaActual.TipoDB) & ") AND " & _
            "(gc.Fechadevol < " & FechaYMD(Fcorte + 1, gobjMain.EmpresaActual.TipoDB) & ") " & _
                 "GROUP BY gc.fechadevol, pc.codprovcli,ele.codelemento,pcG.CodGrupo" & idSec & ",pcc.codcanton,rd.bandpagoprov,ele.orden,ele.codelemento " & _
          " HAVING Sum(rd.valor)>0 order by pc.codprovcli,ele.codelemento "
            gobjMain.EmpresaActual.EjecutarSQL sql, 0
            sql = "select fecharol, codempleado, codelemento, Sum(valor) AS valor  " & _
                  "  from tmp99  " & _
                  " group by fecharol,codempleado,codelemento" & _
                  " order by codempleado,codelemento"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
'#If DAOLIB = 0 Then
'    Set rs.ActiveConnection = Nothing
'#End If
    'Abre la empresa destino
    Set e = AbrirDestino
    Dim r
    With rs
        If Not rs.EOF Then
            rs.MoveLast
            rs.MoveFirst
            If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
            i = 1
            Do Until .EOF
                prg1.value = rs.AbsolutePosition
                prg1.Refresh
                DoEvents
                'Si aplastó 'Cancelar'
                If mbooCancelado Then
                    MsgBox "El proceso fue cancelado.", vbInformation
                    GoTo cancelado
                End If
                    MensajeStatus "Grabando Historial de Roles en la empresa '" & gobjMain.EmpresaActual.CodEmpresa & "'...", vbHourglass
                    Set pc = gobjMain.EmpresaActual.RecuperaEmpleado(rs!CodEmpleado)
                    Set ele = gobjMain.EmpresaActual.RecuperarElemento(rs!Codelemento)
'                    Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idSec, rs!CodGrupo, 4)
'                    If pcg.idgrupo = 0 Then
'                        MensajeStatus "El codigo del grupo " & rs!codDepartamento & " No existe": GoTo cancelado
'                    End If
 '                   Set pcc = gobjMain.EmpresaActual.RecuperaPCCanton(rs!codCanton)
                    If VerificaRol("31/12/" & rs!fecharol, pc.IdProvCli, ele.idElemento) Then
                      r = MsgBox("El Historial de Este Empleado ya Existe .. Desea continuar con el siguiente ", vbYesNo)
                        If r = vbYes Then
                            GoTo continua
                        Else
                            GoTo cancelado
                        End If
                    End If
'                Select Case idSec
'                    Case 1
                        sql = "INSERT INTO HISTORIALROL (fecharol,IdEmpleado,IdElemento,IdCanton,Orden,Valor,IdGrupo1,BandPagoProv)"
                        sql = sql & " VALUES ( '31/12/" & rs!fecharol & "'," & pc.IdProvCli & "," & ele.idElemento & "," & 0 & "," & i & "," & rs!valor & ",0,0)"
'                    Case 2
'                        sql = "INSERT INTO HISTORIALROL (fecharol,IdEmpleado,IdElemento,IdCanton,Orden,Valor,IdGrupo2,BandPagoProv)"
'                        sql = sql & " VALUES ( '31/12/" & rs!fecharol & "'," & pc.IdProvCli & "," & ele.idElemento & "," & pcc.Idcanton & "," & i & "," & rs!valor & "," & pcg.idgrupo & "," & CInt(rs!BandPagoProv) & ")"
'                    Case 3
'                        sql = "INSERT INTO HISTORIALROL (fecharol,IdEmpleado,IdElemento,IdCanton,Orden,Valor,IdGrupo3,BandPagoProv)"
'                        sql = sql & " VALUES ( '31/12/" & rs!fecharol & "'," & pc.IdProvCli & "," & ele.idElemento & "," & pcc.Idcanton & "," & i & "," & rs!valor & "," & pcg.idgrupo & "," & CInt(rs!BandPagoProv) & ")"
'                    Case 4
'                        sql = "INSERT INTO HISTORIALROL (fecharol,IdEmpleado,IdElemento,IdCanton,Orden,Valor,IdGrupo4,BandPagoProv)"
'                        sql = sql & " VALUES ( '31/12/" & rs!fecharol & "'," & pc.IdProvCli & "," & ele.idElemento & "," & pcc.Idcanton & "," & i & "," & rs!valor & "," & pcg.idgrupo & "," & CInt(rs!BandPagoProv) & ")"
'                End Select
                MensajeStatus "Agregando detalle #" & i & " de " & rs.RecordCount, vbHourglass
                gobjMain.EmpresaActual.Execute sql
                i = i + 1
continua:                                    rs.MoveNext
                Set pc = Nothing
                Set pcc = Nothing
                Set pcg = Nothing
                Set ele = Nothing
            Loop
        End If
        .Close
    End With
    mensaje False, "", "OK"
    MensajeStatus
    MsgBox "El proceso terminó con éxito.", vbInformation
    SaldoRolesSii = True
cancelado:
 '   mensaje False, "", Err.Description
    MensajeStatus "", vbNormal
    Set pc = Nothing
    Set pcc = Nothing
    Set pcg = Nothing
    Set ele = Nothing
    Set rs = Nothing
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
'    'Vuelve a abrir la empresa origen
'    Set e = gobjMain.RecuperaEmpresa(codOrig)
'    e.Abrir
'    Set e = Nothing
    mbooProcesando = False                  'Desbloquea que se cierre la ventana
    Exit Function
errtrap:
    MensajeStatus
    MsgBox Err.Description, vbExclamation
    GoTo cancelado
End Function
Public Sub VerificaExistenciaTablaSii(i As Integer)
    Dim rs As Recordset
    Dim sql As String
    'verifica  si la  tabla no esta  creada
    sql = "SELECT * FROM sysobjects WHERE NAME =  'tmp" & i & "'"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not (rs.EOF And rs.BOF) Then
        'elimina la tabla
        gobjMain.EmpresaActual.EjecutarSQL "drop table Tmp" & i, 0
    End If
End Sub
Private Function VerificaRol(ByVal f As Date, ByVal idempleado As Long, ByVal idElemento As Long) As Boolean
Dim sql As String
Dim rs As Recordset
On Error GoTo CapturaError
    sql = "Select * from HistorialRol where fecharol = '" & f & "'"
    sql = sql & " And idEmpleado = " & idempleado
    sql = sql & " And idElemento = " & idElemento
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        VerificaRol = True
    End If
    Set rs = Nothing
    Exit Function
CapturaError:
    Set rs = Nothing
    MsgBox Err.Description
    Exit Function
End Function
Private Function Elementos() As Boolean
    Dim e As Empresa, gc As GNComprobante, ivk As IVKardex, iv As IVinventario
    Dim j As Long, n As Long
    Dim sql As String, rs As Recordset, codOrig As String
    Dim i As Long, c As Currency, Fcorte As Date
    Dim idDep As Byte
    Dim pc As PCProvCli
    Dim pcg As PCGRUPO
    Dim ele As Elementos
    Dim pcc As PCCanton
'    idDep = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Departamento") + 1
    On Error GoTo errtrap
    'Verifica las opciones
'    If Not VerificarOpcion Then Exit Function
    mbooProcesando = True               'Bloquea que se cierre la ventana
    codOrig = gobjMain.EmpresaActual.CodEmpresa
    Fcorte = dtpFechaCorte.value    'Fecha de corte
    'Cambia figura de cursor de mouse
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True
    'Saca las existencias a la fecha de corte
    MensajeStatus "Preparando para grabar Elementos...", vbHourglass
    mensaje True, "Importando elementos para roles..."
    sql = "SELECT * from Elementos"
    Set rs = gobjRol.EmpresaActual.OpenRecordset(sql)
#If DAOLIB = 0 Then
    Set rs.ActiveConnection = Nothing
#End If
    'Abre la empresa destino
    Set e = AbrirDestino
    Dim r
    With rs
        If Not rs.EOF Then
            rs.MoveLast
            rs.MoveFirst
            If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
            i = 1
            Do Until .EOF
                prg1.value = rs.AbsolutePosition
                prg1.Refresh
                DoEvents
                'Si aplastó 'Cancelar'
                If mbooCancelado Then
                    MsgBox "El proceso fue cancelado.", vbInformation
                    GoTo cancelado
                End If
                    MensajeStatus "Grabando Empresa de Roles en la empresa '" & gobjMain.EmpresaActual.CodEmpresa & "'...", vbHourglass
                    If Not ExisteElemento(rs!Codelemento) Then
                        sql = "INSERT INTO ELEMENTO (CodElemento,Nombre,Descripcion,Formula,Meses,Tipo,Editable,Orden,BandActivo,Visible,Imprimir,debe,haber,afectaemp," & _
                            "bandacumular,bandmostrarenprovision,bandmostrarenreporte,bandVALIDARasignar)"
                        sql = sql & " VALUES ( '" & rs!Codelemento & "','" & rs!nombre & "','" & rs!Descripcion & "','" & rs!Formula & "','" & rs!Meses & "','" & rs!Tipo & "','" & Abs(CInt(rs!Editable)) & "','" & rs!Orden & "','" & Abs(CInt(rs!BandActivo)) & "','" & Abs(CInt(rs!Visible)) & "','" & Abs(CInt(rs!Imprimir)) & "',0,0,0," & _
                            "'" & Abs(CInt(rs!BandAcumular)) & "','" & Abs(CInt(rs!bandmostrarenprovision)) & "','" & Abs(CInt(rs!bandmostrarenreporte)) & "','" & Abs(CInt(rs!BandVALIDARAsignar)) & "')"
                    MensajeStatus "Agregando Elementos #" & i & " de " & rs.RecordCount, vbHourglass
                    gobjMain.EmpresaActual.Execute sql
                    End If
                i = i + 1
                rs.MoveNext
                Set pc = Nothing
                Set pcc = Nothing
                Set pcg = Nothing
                Set ele = Nothing
            Loop
        End If
        .Close
    End With
    mensaje False, "", "OK"
    MensajeStatus
    MsgBox "El proceso terminó con éxito.", vbInformation
    Elementos = True
cancelado:
    MensajeStatus "", vbNormal
    Set pc = Nothing
    Set pcc = Nothing
    Set pcg = Nothing
    Set ele = Nothing
    Set rs = Nothing
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
    mbooProcesando = False
    Exit Function
errtrap:
    MensajeStatus
    MsgBox Err.Description, vbExclamation
    GoTo cancelado
End Function
Private Function Departamentos() As Boolean
    Dim e As Empresa, gc As GNComprobante
    Dim j As Long, n As Long
    Dim sql As String, rs As Recordset, codOrig As String
    Dim i As Long, c As Currency, Fcorte As Date
    Dim idDep As Byte
    idDep = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Departamento") + 1
    On Error GoTo errtrap
    'Verifica las opciones
'    If Not VerificarOpcion Then Exit Function
    mbooProcesando = True               'Bloquea que se cierre la ventana
    codOrig = gobjMain.EmpresaActual.CodEmpresa
    Fcorte = dtpFechaCorte.value    'Fecha de corte
    'Cambia figura de cursor de mouse
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True
    'Saca las existencias a la fecha de corte
    MensajeStatus "Preparando para grabar Departametnos...", vbHourglass
    mensaje True, "Importando Departamentos para roles..."
    sql = "SELECT * from Departamento"
    Set rs = gobjRol.EmpresaActual.OpenRecordset(sql)
#If DAOLIB = 0 Then
    Set rs.ActiveConnection = Nothing
#End If
    'Abre la empresa destino
    Set e = AbrirDestino
    Dim r
    With rs
        If Not rs.EOF Then
            rs.MoveLast
            rs.MoveFirst
            If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
            i = 1
            Do Until .EOF
                prg1.value = rs.AbsolutePosition
                prg1.Refresh
                DoEvents
                'Si aplastó 'Cancelar'
                If mbooCancelado Then
                    MsgBox "El proceso fue cancelado.", vbInformation
                    GoTo cancelado
                End If
                    MensajeStatus "Grabando Empresa de Roles en la empresa '" & gobjMain.EmpresaActual.CodEmpresa & "'...", vbHourglass
                    Select Case idDep
                        Case 1
                            If VerificaDepartamento(rs!codDepartamento, idDep) Then
                                sql = "update pcgrupo1 set codgrupo1 = '" & rs!codDepartamento & "',Descripcion = '" & rs!nombre & "',origen = 4" ' (CodGrupo2,Descripcion,BandValida,preciosDisponibles,Origen)"
                                sql = sql & " where codgrupo1 = '" & rs!codDepartamento & "'"
                            Else
                                sql = "INSERT INTO PCGRUPO1 (CodGrupo1,Descripcion,BandValida,preciosDisponibles,Origen)"
                                sql = sql & " VALUES('" & rs!codDepartamento & "','" & rs!nombre & "',1,0,4)"
                            End If
                        Case 2
                            If VerificaDepartamento(rs!codDepartamento, idDep) Then
                                sql = "update pcgrupo2 set codgrupo2 = '" & rs!codDepartamento & "',Descripcion = '" & rs!nombre & "',origen = 4" ' (CodGrupo2,Descripcion,BandValida,preciosDisponibles,Origen)"
                                sql = sql & " where codgrupo2 = '" & rs!codDepartamento & "'"
                            Else
                                sql = "INSERT INTO PCGRUPO2 (CodGrupo2,Descripcion,BandValida,preciosDisponibles,Origen)"
                                sql = sql & " VALUES('" & rs!codDepartamento & "','" & rs!nombre & "',1,0,4)"
                            End If
                        Case 3
                            If VerificaDepartamento(rs!codDepartamento, idDep) Then
                                sql = "update pcgrupo3 set codgrupo3 = '" & rs!codDepartamento & "',Descripcion = '" & rs!nombre & "',origen = 4" ' (CodGrupo2,Descripcion,BandValida,preciosDisponibles,Origen)"
                                sql = sql & " where codgrupo3 = '" & rs!codDepartamento & "'"
                            Else
                                sql = "INSERT INTO PCGRUPO3 (CodGrupo3,Descripcion,BandValida,preciosDisponibles,Origen)"
                                sql = sql & " VALUES('" & rs!codDepartamento & "','" & rs!nombre & "',1,0,4)"
                            End If
                        Case 4
                            If VerificaDepartamento(rs!Departamento, idDep) Then
                                sql = "update pcgrupo1 set codgrupo4 = '" & rs!codDepartamento & "',Descripcion = '" & rs!nombre & "',origen = 4" ' (CodGrupo2,Descripcion,BandValida,preciosDisponibles,Origen)"
                                sql = sql & " where codgrupo4 = '" & rs!Departamento & "'"
                            Else
                                sql = "INSERT INTO PCGRUPO4 (CodGrupo4,Descripcion,BandValida,preciosDisponibles,Origen)"
                                sql = sql & " VALUES('" & rs!codDepartamento & "','" & rs!nombre & "',1,0,4)"
                            End If
                    End Select
                MensajeStatus "Agregando Dep #" & i & " de " & rs.RecordCount, vbHourglass
                gobjMain.EmpresaActual.Execute sql
                i = i + 1
                rs.MoveNext
            Loop
        End If
        .Close
    End With
    mensaje False, "", "OK"
    MensajeStatus
    MsgBox "El proceso terminó con éxito.", vbInformation
    Departamentos = True
cancelado:
    MensajeStatus "", vbNormal
    Set rs = Nothing
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
    mbooProcesando = False
Exit Function
errtrap:
    MensajeStatus
    MsgBox Err.Description, vbExclamation
    GoTo cancelado
End Function
Private Function Cargos() As Boolean
    Dim e As Empresa, gc As GNComprobante
    Dim j As Long, n As Long
    Dim sql As String, rs As Recordset, codOrig As String
    Dim i As Long, c As Currency, Fcorte As Date
    Dim idcargo As Byte
    Dim resp
    idcargo = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Cargo") + 1
    On Error GoTo errtrap
    mbooProcesando = True               'Bloquea que se cierre la ventana
    codOrig = gobjMain.EmpresaActual.CodEmpresa
    Fcorte = dtpFechaCorte.value    'Fecha de corte
    'Cambia figura de cursor de mouse
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True
    'Saca las existencias a la fecha de corte
    MensajeStatus "Preparando para grabar Cargos...", vbHourglass
    mensaje True, "Importando Cargos para roles..."
    sql = "SELECT * from Cargo"
    Set rs = gobjRol.EmpresaActual.OpenRecordset(sql)
#If DAOLIB = 0 Then
    Set rs.ActiveConnection = Nothing
#End If
    'Abre la empresa destino
    Set e = AbrirDestino
    Dim r
    With rs
        If Not rs.EOF Then
            rs.MoveLast
            rs.MoveFirst
            If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
            i = 1
            Do Until .EOF
                prg1.value = rs.AbsolutePosition
                prg1.Refresh
                DoEvents
                'Si aplastó 'Cancelar'
                If mbooCancelado Then
                    MsgBox "El proceso fue cancelado.", vbInformation
                    GoTo cancelado
                End If
                    MensajeStatus "Grabando Empresa de Roles en la empresa '" & gobjMain.EmpresaActual.CodEmpresa & "'...", vbHourglass
                    Select Case idcargo
                        Case 1
                            If VerificaCargo(rs!codcargo, idcargo) Then
                                sql = "update pcgrupo1 set codgrupo1 = '" & rs!codcargo & "',Descripcion = '" & rs!Descripcion & "',origen = 4" ' (CodGrupo2,Descripcion,BandValida,preciosDisponibles,Origen)"
                                sql = sql & " where codgrupo1 = '" & rs!codcargo & "'"
                            Else
                                sql = "INSERT INTO PCGRUPO1 (CodGrupo1,Descripcion,BandValida,preciosDisponibles,Origen)"
                                sql = sql & " VALUES('" & rs!codcargo & "','" & rs!Descripcion & "',1,0,4)"
                            End If
                        Case 2
                            If VerificaCargo(rs!codcargo, idcargo) Then
                                sql = "update pcgrupo2 set codgrupo2 = '" & rs!codcargo & "',Descripcion = '" & rs!Descripcion & "',origen = 4" ' (CodGrupo2,Descripcion,BandValida,preciosDisponibles,Origen)"
                                sql = sql & " where codgrupo2 = '" & rs!codcargo & "'"
                            Else
                                sql = "INSERT INTO PCGRUPO2 (CodGrupo2,Descripcion,BandValida,preciosDisponibles,Origen)"
                                sql = sql & " VALUES('" & rs!codcargo & "','" & rs!Descripcion & "',1,0,4)"
                            End If
                        Case 3
                            If VerificaCargo(rs!codcargo, idcargo) Then
                                sql = "update pcgrupo3 set codgrupo3 = '" & rs!codcargo & "',Descripcion = '" & rs!Descripcion & "',origen = 4" ' (CodGrupo3,Descripcion,BandValida,preciosDisponibles,Origen)"
                                sql = sql & " where codgrupo3 = '" & rs!codcargo & "'"
                            Else
                                sql = "INSERT INTO PCGRUPO3 (CodGrupo3,Descripcion,BandValida,preciosDisponibles,Origen)"
                                sql = sql & " VALUES('" & rs!codcargo & "','" & rs!Descripcion & "',1,0,4)"
                            End If
                        Case 4
                            If VerificaCargo(rs!codcargo, idcargo) Then
                                sql = "update pcgrupo4 set codgrupo4 = '" & rs!codcargo & "',Descripcion = '" & rs!Descripcion & "',origen = 4" ' (CodGrupo2,Descripcion,BandValida,preciosDisponibles,Origen)"
                                sql = sql & " where codgrupo4 = '" & rs!codcargo & "'"
                            Else
                                sql = "INSERT INTO PCGRUPO4 (CodGrupo4,Descripcion,BandValida,preciosDisponibles,Origen)"
                                sql = sql & " VALUES('" & rs!codcargo & "','" & rs!Descripcion & "',1,0,4)"
                            End If
                    End Select
                MensajeStatus "Agregando cargo #" & i & " de " & rs.RecordCount, vbHourglass
                gobjMain.EmpresaActual.Execute sql
                i = i + 1
siguiente:              rs.MoveNext
            Loop
        End If
        .Close
    End With
    mensaje False, "", "OK"
    MensajeStatus
    MsgBox "El proceso terminó con éxito.", vbInformation
    Cargos = True
cancelado:
    MensajeStatus "", vbNormal
    Set rs = Nothing
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
    mbooProcesando = False                  'Desbloquea que se cierre la ventana
    Exit Function
errtrap:
'    MensajeStatus
    resp = MsgBox(rs!Descripcion & " " & Err.Description & " Desea Continuar....", vbYesNo)
    If resp = vbYes Then
        GoTo siguiente
    Else
        GoTo cancelado
    End If
End Function
Private Function VerificaDepartamento(ByVal codigo As String, ByVal idDep As Byte) As Boolean
    Dim sql As String
    Dim rs As Recordset
    Select Case idDep
        Case 1:        sql = "Select * from pcgrupo1 where codgrupo1 = '" & codigo & "'"
        Case 2:        sql = "Select * from pcgrupo2 where codgrupo2 = '" & codigo & "'"
        Case 3:        sql = "Select * from pcgrupo3 where codgrupo3 = '" & codigo & "'"
        Case 4:        sql = "Select * from pcgrupo4 where codgrupo4 = '" & codigo & "'"
    End Select
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        VerificaDepartamento = True
    End If
    Set rs = Nothing
End Function
Private Function VerificaCargo(ByVal codigo As String, ByVal idcargo As Byte) As Boolean
    Dim sql As String
    Dim rs As Recordset
    Select Case idcargo
        Case 1:        sql = "Select * from pcgrupo1 where codgrupo1 = '" & codigo & "'"
        Case 2:        sql = "Select * from pcgrupo2 where codgrupo2 = '" & codigo & "'"
        Case 3:        sql = "Select * from pcgrupo3 where codgrupo3 = '" & codigo & "'"
        Case 4:        sql = "Select * from pcgrupo4 where codgrupo4 = '" & codigo & "'"
    End Select
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        VerificaCargo = True
    End If
    Set rs = Nothing
End Function



Private Function Personal() As Boolean
    Dim e As Empresa, gc As GNComprobante
    Dim resp
    Dim pc As PCProvCli
    Dim j As Long, n As Long
    Dim sql As String, rs As Recordset, codOrig As String
    Dim i As Long, c As Currency, Fcorte As Date, rsAux As Recordset
    Dim idDep As Byte
    idDep = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Departamento") + 1
    On Error GoTo errtrap
    'Verifica las opciones
'    If Not VerificarOpcion Then Exit Function
    mbooProcesando = True               'Bloquea que se cierre la ventana
    codOrig = gobjMain.EmpresaActual.CodEmpresa
    Fcorte = dtpFechaCorte.value    'Fecha de corte
    'Cambia figura de cursor de mouse
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True
    'Saca las existencias a la fecha de corte
    MensajeStatus "Preparando para grabar personal...", vbHourglass
    mensaje True, "Importando personla para roles..."
    sql = "SELECT * from Personal"
    Set rs = gobjRol.EmpresaActual.OpenRecordset(sql)
#If DAOLIB = 0 Then
    Set rs.ActiveConnection = Nothing
#End If
    'Abre la empresa destino
    Set e = AbrirDestino
    Dim r
    With rs
        If Not rs.EOF Then
        
            rs.MoveLast
            rs.MoveFirst
            If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
            i = 1
            Do Until .EOF
                prg1.value = rs.AbsolutePosition
                prg1.Refresh
                DoEvents
                'Si aplastó 'Cancelar'
                If mbooCancelado Then
                    MsgBox "El proceso fue cancelado.", vbInformation
                    GoTo cancelado
                End If
                    MensajeStatus "Grabando Empresa de Roles en la empresa '" & gobjMain.EmpresaActual.CodEmpresa & "'...", vbHourglass
                    Set pc = gobjMain.EmpresaActual.RecuperaEmpleado(rs!CodEmpleado)
                    If Not pc Is Nothing Then
                    sql = "select idempleado from personal where idempleado =" & pc.IdProvCli
                        Set rsAux = gobjMain.EmpresaActual.OpenRecordset(sql)
                        If rsAux.RecordCount = 0 Then
                            If Not pc Is Nothing Then
                                sql = "INSERT INTO Personal (idEmpleado,Sexo, EstadoCivil, NumCargas,Salario,TipoSalario, FechaIngreso,FechaEgreso,bandActivo,Varios1,Varios2,Varios3,Varios4,Varios5,Varios6,Varios7,Varios8,PagarSeguro,Contador,FechaPagoSeguro,bandFR,bandPagarHE )"
                                sql = sql & " VALUES(" & pc.IdProvCli & "," & rs!sexo & "," & rs!EstadoCivil & "," & rs!NumCargas & "," & rs!Salario & "," & rs!TipoSalario & ",'" & rs!FechaIngreso & "','" & rs!FechaEgreso & "'," & IIf(rs!BandActivo, 1, 0) & "," & rs!Varios1 & "," & rs!Varios2 & "," & rs!Varios3 & "," & rs!Varios4 & "," & rs!Varios5 & "," & rs!Varios6 & "," & rs!Varios7 & "," & rs!Varios8 & "," & _
                                "" & rs!PagarSeguro & "," & rs!contador & ",'" & rs!FechaPagoSeguro & " '," & IIf(rs!BandFR, 1, 0) & "," & IIf(rs!BandPagarHE, 1, 0) & ")"
                            End If
                                
                            MensajeStatus "Agregando Empleado #" & i & " de " & rs.RecordCount, vbHourglass
                            gobjMain.EmpresaActual.Execute sql
                            i = i + 1
                        End If
                    Else
                        MsgBox "Eel empleado con codigo: " & rs!CodEmpleado & " no esta como empleado en Sii4a"
                    End If
siguiente:                 rs.MoveNext
                Set pc = Nothing
            Loop
        End If
        .Close
    End With
    mensaje False, "", "OK"
    MensajeStatus
    MsgBox "El proceso terminó con éxito.", vbInformation
    Personal = True
cancelado:
    MensajeStatus "", vbNormal
    Set rs = Nothing
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
    mbooProcesando = False
Exit Function
errtrap:
    resp = MsgBox(Err.Description & " Desea Continuar....", vbYesNo)
    If resp = vbYes Then
        GoTo siguiente
    Else
        GoTo cancelado
    End If
End Function

Private Function CuentasRol() As Boolean
    Dim e As Empresa, gc As GNComprobante
    Dim resp
    Dim pcg As PCGRUPO
    Dim ele As Elementos
    Dim j As Long, n As Long
    Dim sql As String, rs As Recordset, codOrig As String
    Dim i As Long, c As Currency, Fcorte As Date
    Dim idDep As Byte
    Dim idSec As Byte
    Dim idG As Byte
    Dim rsE As Recordset
    idDep = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Departamento") + 1
    idSec = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Seccion") + 1
    On Error GoTo errtrap
    'Verifica las opciones
'    If Not VerificarOpcion Then Exit Function
    mbooProcesando = True               'Bloquea que se cierre la ventana
    codOrig = gobjMain.EmpresaActual.CodEmpresa
    Fcorte = dtpFechaCorte.value    'Fecha de corte
    'Cambia figura de cursor de mouse
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True
    'Saca las existencias a la fecha de corte
    MensajeStatus "Preparando para grabar personal...", vbHourglass
    mensaje True, "Importando personla para roles..."
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SeContabilizaPor")) > 0 Then
        If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SeContabilizaPor") = 1 Then
            sql = "SELECT ctd.* from cuentadepartamento ctd inner join elementos e on e.codelemento = ctd.codelemento"
        Else
            sql = "SELECT ctd.* from cuentaPersonal ctd inner join elementos e on e.codelemento = ctd.codelemento"
        End If
    End If
        Set rs = gobjRol.EmpresaActual.OpenRecordset(sql)
#If DAOLIB = 0 Then
    Set rs.ActiveConnection = Nothing
#End If
    'Abre la empresa destino
    Set e = AbrirDestino
    Dim r
    With rs
        If Not rs.EOF Then
            rs.MoveLast
            rs.MoveFirst
            If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
            i = 1
            Do Until .EOF
                prg1.value = rs.AbsolutePosition
                prg1.Refresh
                DoEvents
                'Si aplastó 'Cancelar'
                If mbooCancelado Then
                    MsgBox "El proceso fue cancelado.", vbInformation
                    GoTo cancelado
                End If
                    MensajeStatus "Grabando Empresa de Roles en la empresa '" & gobjMain.EmpresaActual.CodEmpresa & "'...", vbHourglass
                    
                    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SeContabilizaPor")) > 0 Then
                        If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SeContabilizaPor") = 1 Then
                            idG = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ConAsiento") + 1
                            Select Case idG
                                Case 1
                                    Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, rs!codDepartamento, 4)
                                    Set ele = gobjMain.EmpresaActual.RecuperarElemento(rs!Codelemento)
                                    If Not pcg Is Nothing Then
                                        'sql = "Select idgrupo" & idG & " from pcprovcli where idgrupo" & idG & "= " & pcg.idgrupo
                                       ' Set rsE = gobjMain.EmpresaActual.OpenRecordset(sql)
                                        'Do While Not rsE.EOF
                                            sql = "INSERT INTO Cuentapcgrupo (idPCGrupo,idElemento,IdCuenta)"
                                            sql = sql & " VALUES(" & pcg.idgrupo & "," & ele.idElemento & "," & rs!IdCuenta & ")"
                                            MensajeStatus "Agregando CuentasPcGrupo  ", vbHourglass
                                            gobjMain.EmpresaActual.Execute sql
                                        '    rsE.MoveNext
                                       ' Loop
                                        'Set rsE = Nothing
                                    End If
                            End Select
                        Else
                            Set ele = gobjMain.EmpresaActual.RecuperarElemento(rs!Codelemento)
                           If Not ele Is Nothing Then
                                sql = "Select idprovcli from Empleado where bandempleado = 1 AND  ruc = '" & rs!CodEmpleado & "'"
                                Set rsE = gobjMain.EmpresaActual.OpenRecordset(sql)
                                Do While Not rsE.EOF
                                    If Not ExisteCuenta(rsE!IdProvCli, ele.idElemento) Then
                                        sql = "INSERT INTO CuentaPersonal (idEmpleado,idelemento, idcuenta)"
                                        sql = sql & " VALUES(" & rsE!IdProvCli & "," & ele.idElemento & "," & rs!IdCuenta & ")"
                                        MensajeStatus "Agregando CUENTASPERSONAL #" & i & " de " & rs.RecordCount, vbHourglass
                                        gobjMain.EmpresaActual.Execute sql
                                    End If
        
                                rsE.MoveNext
                                Loop
                                Set rsE = Nothing
                            End If
                        End If
                    End If
                i = i + 1
siguiente:
            rs.MoveNext
                Set pcg = Nothing
                Set ele = Nothing
            Loop
        End If
        .Close
    End With
    mensaje False, "", "OK"
    MsgBox "El proceso terminó con éxito.", vbInformation
    CuentasRol = True
cancelado:
    MensajeStatus "", vbNormal
    Set rs = Nothing
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
    mbooProcesando = False
Exit Function
errtrap:
    resp = MsgBox(Err.Description & " Desea Continuar....", vbYesNo)
    If resp = vbYes Then
        GoTo siguiente
    Else
        GoTo cancelado
    End If
End Function

Private Function CuentasRolPre() As Boolean
    Dim e As Empresa, gc As GNComprobante
    Dim resp
    Dim pcg As PCGRUPO
    Dim ele As Elementos
    Dim j As Long, n As Long
    Dim sql As String, rs As Recordset, codOrig As String
    Dim i As Long, c As Currency, Fcorte As Date
    Dim idDep As Byte
    Dim rsE As Recordset
    idDep = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Departamento") + 1
    On Error GoTo errtrap
    'Verifica las opciones
'    If Not VerificarOpcion Then Exit Function
    mbooProcesando = True               'Bloquea que se cierre la ventana
    codOrig = gobjMain.EmpresaActual.CodEmpresa
    Fcorte = dtpFechaCorte.value    'Fecha de corte
    'Cambia figura de cursor de mouse
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True
    'Saca las existencias a la fecha de corte
    MensajeStatus "Preparando para grabar personal...", vbHourglass
    mensaje True, "Importando cuentas de presupuesto desde roles roles..."
    sql = "SELECT * from cuentadepartamentoPre"
    Set rs = gobjRol.EmpresaActual.OpenRecordset(sql)
#If DAOLIB = 0 Then
    Set rs.ActiveConnection = Nothing
#End If
    'Abre la empresa destino
    Set e = AbrirDestino
    Dim r
    With rs
        If Not rs.EOF Then
            rs.MoveLast
            rs.MoveFirst
            If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
            i = 1
            Do Until .EOF
                prg1.value = rs.AbsolutePosition
                prg1.Refresh
                DoEvents
                'Si aplastó 'Cancelar'
                If mbooCancelado Then
                    MsgBox "El proceso fue cancelado.", vbInformation
                    GoTo cancelado
                End If
                    MensajeStatus "Grabando Empresa de Roles en la empresa '" & gobjMain.EmpresaActual.CodEmpresa & "'...", vbHourglass
                    Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idDep, rs!codDepartamento, 4)
                    Set ele = gobjMain.EmpresaActual.RecuperarElemento(rs!Codelemento)
                    If Not pcg Is Nothing Then
                        sql = "Select idprovcli from Empleado where idgrupo" & idDep & "= " & pcg.idgrupo
                        Set rsE = gobjMain.EmpresaActual.OpenRecordset(sql)
                        Do While Not rsE.EOF
                            sql = "INSERT INTO CuentaPersonalPre (idEmpleado,idelemento, idcuenta)"
                            sql = sql & " VALUES(" & rsE!IdProvCli & "," & ele.idElemento & "," & rs!IdCuenta & ")"
                            MensajeStatus "Agregando CUENTASPERSONALpre #" & i & " de " & rs.RecordCount, vbHourglass
                            gobjMain.EmpresaActual.Execute sql

                        rsE.MoveNext
                        Loop
                        Set rsE = Nothing
                    End If
                    
                i = i + 1
siguiente:                 rs.MoveNext
                Set pcg = Nothing
                Set ele = Nothing
            Loop
        End If
        .Close
    End With
    mensaje False, "", "OK"
    MsgBox "El proceso terminó con éxito.", vbInformation
    CuentasRolPre = True
cancelado:
    MensajeStatus "", vbNormal
    Set rs = Nothing
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
    mbooProcesando = False
Exit Function
errtrap:
    resp = MsgBox(Err.Description & " Desea Continuar....", vbYesNo)
    If resp = vbYes Then
        GoTo siguiente
    Else
        GoTo cancelado
    End If
End Function



Private Function ExisteElemento(ByVal Codelemento As String) As Boolean
Dim sql As String
Dim rs As Recordset
sql = "Select * from elemento where codelemento = '" & Codelemento & "'"
Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
If rs.RecordCount > 0 Then
    ExisteElemento = True
    Exit Function
End If
ExisteElemento = False
End Function


Private Function ExisteCuenta(ByVal idempleado As Long, ByVal idElemento As Long) As Boolean
Dim sql As String
Dim rs As Recordset
sql = "Select * from CuentaPersonal where idempleado = " & idempleado
sql = sql & " AND idelemento = " & idElemento
Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
If rs.RecordCount > 0 Then
    ExisteCuenta = True
    Exit Function
End If
ExisteCuenta = False
End Function

Private Function SaldoInicialAFCustodios() As Boolean
    Dim e As Empresa, gc As GNComprobante, afk As AFKardexCustodio, af As AFinventario
    Dim j As Long, n As Long
    Dim sql As String, rs As Recordset, codOrig As String
    Dim i As Long, c As Currency, Fcorte As Date
    On Error GoTo errtrap
    
    'Verifica las opciones
    If Not VerificarOpcion Then Exit Function
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    
    codOrig = gobjMain.EmpresaActual.CodEmpresa
    Fcorte = dtpFechaCorte.value    'Fecha de corte

    'Cambia figura de cursor de mouse
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True
    
    'Saca las existencias a la fecha de corte
    MensajeStatus "Preparando para grabar las existencias iniciales de custodios activos fijos...", vbHourglass
    mensaje True, "Saldo inicial de custodios de activos fijos..."
    
'    sql = "SELECT af.IdInventario, 1 , "
'    sql = sql & " af.CodInventario, 'B01' as codbodega, "
'    sql = sql & "ISNULL((ISNULL(af.numvidautil,0)- isnull(af.depanterior,0)),0) AS Cant, costoultimoingreso "
'    sql = sql & "FROM  afInventario af "
'    sql = sql & "WHERE FECHACOMPRA <" & FechaYMD(Fcorte + 1, gobjMain.EmpresaActual.TipoDB)

'    sql = " select a.idinventario, af.codinventario, max('B01') as codbodega ,"
'    sql = sql & " Count (numtrans)as cant , costoultimoingreso "
'    sql = sql & " from gncomprobante g"
'    sql = sql & " inner join afkardex a"
'    sql = sql & " inner join afinventario af"
'    sql = sql & " on a.idinventario=af.idinventario"
'    sql = sql & " on g.transid=a.transid where codtrans='DEPAF'"
'    sql = sql & " and g.estado <>3"
'    sql = sql & " and g.Fechatrans < " & FechaYMD(Fcorte + 1, gobjMain.EmpresaActual.TipoDB)
'    sql = sql & " group by a.idinventario, af.codinventario"

sql = " select"
sql = sql & " afe.idinventario,1,af.codinventario, pc.codprovcli as codbodega, 1 as cant ,0"
sql = sql & " from afInventario af"
sql = sql & " inner join afexistcustodio afe"
sql = sql & " inner join empleado pc"
sql = sql & " on afe.idprovcli=pc.idprovcli"
sql = sql & " on af.idinventario=afe.idinventario"
sql = sql & " Where afe.exist > 0 order by pc.nombre"


    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
#If DAOLIB = 0 Then
    Set rs.ActiveConnection = Nothing
#End If
    
    'Abre la empresa destino
    Set e = AbrirDestino
    
    With rs
        If Not rs.EOF Then
            rs.MoveLast
            rs.MoveFirst
            If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
            i = 0
            Do Until .EOF
                prg1.value = rs.AbsolutePosition
                prg1.Refresh
                DoEvents
                
                'Si aplastó 'Cancelar'
                If mbooCancelado Then
                    MsgBox "El proceso fue cancelado.", vbInformation
                    GoTo cancelado
                End If
                
                'Crea transaccion 'AFSI'
                If (i Mod 100) = 0 Then
                    'Si no es primera vez
                    If Not (gc Is Nothing) Then
                        'Graba la transacción
                        MensajeStatus "Grabando la transacción en la empresa '" & gc.Empresa.CodEmpresa & "'...", vbHourglass
                        gc.HoraTrans = "00:00:01"
                        gc.Grabar False, False
                    End If
                    
                    Set gc = CrearTrans(e, _
                            "IVCAF", _
                            "Saldo inicial de Custodios activos fijos", _
                            Fcorte, _
                            "")
                End If
                
                'Recupera datos de inventario para llama el método Costo()
                MensajeStatus "Agregando detalle #" & i & " de " & rs.RecordCount, vbHourglass
                Set af = mEmpOrigen.RecuperaAFInventario(.Fields("IdInventario"))
                
                'Obtiene Costo del item en Moneda de item
                c = af.costo(Fcorte, 1)
                
                'De moneda de item, covierte en moneda de trans, si es necesario
                If af.CodMoneda <> gc.CodMoneda Then
                    c = c * gc.Cotizacion(af.CodMoneda) / gc.Cotizacion("")
                End If
                
                'Agrega detalle
                j = gc.AddAFKardexCustodio
                Set afk = gc.AFKardexCustodio(j)
                If .Fields("CANT") = 0 Then
                    afk.cantidad = 1
'                ElseIf .Fields("CANT") = "Nulo" Then
'                    afk.cantidad = 1
                Else
                    afk.cantidad = .Fields("CANT")
                End If
                afk.CodEmpleado = .Fields("CodBodega")
                'afk.CodInventario = .Fields("CodInventario")
                afk.idinventario = .Fields("idInventario")
'                afk.CostoRealTotal = .Fields("costoultimoingreso")
 '               afk.CostoTotal = .Fields("costoultimoingreso")
'                afk.PrecioRealTotal = .Fields("costoultimoingreso")
'                afk.PrecioTotal = .Fields("costoultimoingreso")
                afk.Orden = i Mod 100
                i = i + 1
                .MoveNext
            Loop
        End If
        .Close
    End With
        
    If Not (gc Is Nothing) Then
        'Graba la transacción
        MensajeStatus "Grabando la transacción en la empresa '" & gc.Empresa.CodEmpresa & "'...", vbHourglass
        gc.HoraTrans = "00:00:01"
        gc.Grabar False, False
    End If
    
    'Corrige las existencias para que quede bien la tabla 'IVExist'
    MensajeStatus "Arreglando las existencias...", vbHourglass
    If Not (gc Is Nothing) Then
        gc.Empresa.CorregirExistenciaAFCustodio
    End If
    mensaje False, "", "OK"
    MensajeStatus
    MsgBox "El proceso terminó con éxito.", vbInformation
    SaldoInicialAFCustodios = True
    
cancelado:
    mensaje False, "", Err.Description
    MensajeStatus
    Set afk = Nothing
    Set af = Nothing
    Set gc = Nothing
    Set rs = Nothing
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
    
    'Vuelve a abrir la empresa origen
    Set e = gobjMain.RecuperaEmpresa(codOrig)
    e.Abrir
    Set e = Nothing
    
    mbooProcesando = False                  'Desbloquea que se cierre la ventana
    Exit Function
errtrap:
    MensajeStatus
    MsgBox Err.Description, vbExclamation
    GoTo cancelado
End Function


'16. Pasar saldo inicial de Empleados
Private Function SaldoEmpRol(ByVal Codelemento As String) As Boolean
    Dim e As Empresa, pck As PCKardex
    Dim j As Long, sql As String, rs As Recordset
    Dim i As Long, c As Currency, Fcorte As Date
    Dim gcCLNE As GNComprobante
    Dim gcPVNE As GNComprobante
    On Error GoTo errtrap
    
    'Verifica las opciones
    If Not VerificarOpcion Then Exit Function
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    Fcorte = dtpFechaCorte.value    'Fecha de corte

    'Cambia figura de cursor de mouse
    MensajeStatus "Está preparando saldos a la fecha de corte...", vbHourglass
    mensaje True, "Saldo inicial de Empleados..." & Codelemento
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True

    'Obtiene Saldos de proveedor/cliente por cada documento pendiente
    sql = "spConsEmpSaldoRol '" & Codelemento & "', 2," & FechaYMD(Fcorte, gobjMain.EmpresaActual.TipoDB)
    Set rs = mEmpOrigen.OpenRecordset(sql)
    UltimoRecordset rs
    'Abre la empresa destino
    Set e = AbrirDestino
    
    With rs
        If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
        i = 0
        Do Until .EOF
            prg1.value = rs.AbsolutePosition
            prg1.Refresh
            DoEvents
            MensajeStatus "Agregando detalle: #" & i & " de " & rs.RecordCount, vbHourglass
            
            'Si aplastó 'Cancelar'
            If mbooCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                GoTo cancelado
            End If
            
            'Si es empleados por pagar
            If .Fields("Saldo") < 0 And .Fields("BandEmpleado") <> 0 Then
                Set pck = PrepararTransPC(e, "CLNE", _
                            "Saldo inicial de Empleados x pagar " & Codelemento, _
                            Fcorte, gcCLNE)
'            'Si es Empleados por pagar (Anticipado)
            Else
     
                Set pck = PrepararTransPC(e, "PVNE", _
                            "Saldo inicial de Empleados (Anticipos)", _
                                   Fcorte, gcPVNE)
            End If
            'Recupera datos de proveedor y asigna al objeto
            pck.CodEmpleado = .Fields("CodProvCli")
            If .Fields("Saldo") > 0 Then   'Si es por cobrar --> Debe
                pck.Debe = .Fields("Saldo")        'Saldo en dólares
            Else                                    'Si es por pagar --> Haber
                pck.Haber = .Fields("Saldo") * -1     'Saldo en dólares
            End If
            pck.codforma = .Fields("CodForma")
            pck.FechaEmision = .Fields("FechaEmision")
            pck.FechaVenci = .Fields("FechaVenci")
            pck.NumLetra = .Fields("Trans") & "_" & Codelemento
            pck.Observacion = .Fields("Observacion")
            pck.Orden = i
            pck.Guid = .Fields("Guid")        '*** <== AGREGAR ESTO
            pck.idElemento = .Fields("Idelemento")
            If Not IsNull(.Fields("CodVendedor")) Then pck.CodVendedor = .Fields("codvendedor") 'AUC 04/06/07
            i = i + 1
            .MoveNext
        Loop
        .Close
    End With
    
    'Graba la transacción si no están grabadas
    MensajeStatus "Grabándo la transacción...", vbHourglass
    
    If Not (gcCLNE Is Nothing) Then
        gcCLNE.HoraTrans = "00:00:01"
        gcCLNE.Grabar False, False
    End If
    If Not (gcPVNE Is Nothing) Then
        gcPVNE.HoraTrans = "00:00:01"
        gcPVNE.Grabar False, False
    End If
    
    MensajeStatus
    mensaje False, "", "OK"
    MsgBox "El proceso para " & Codelemento & " terminó con éxito.", vbInformation
    SaldoEmpRol = True
    
cancelado:
    Set rs = Nothing
    MensajeStatus
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
    
    'Libera los objetos utilizados
    Set pck = Nothing
    Set gcCLNE = Nothing
    Set gcPVNE = Nothing
    Set e = Nothing
    mbooProcesando = False               'Desbloquea que se cierre la ventana
    Exit Function
errtrap:
    mensaje False, "", Err.Description
    MensajeStatus
    DispErr
    GoTo cancelado
End Function

Private Sub CargaEleRol()
    Dim i As Long
    Dim rs As Recordset
    lstEle.Clear
    Set rs = gobjMain.EmpresaActual.ListaElementosAfectaSaldoEmp
    Do While Not rs.EOF
      lstEle.AddItem rs!Codelemento
      rs.MoveNext
    Loop
    For i = 0 To lstEle.ListCount - 1
        lstEle.Selected(i) = True
    Next
End Sub

Private Function SaldoEmpSinRol() As Boolean
    Dim e As Empresa, pck As PCKardex
    Dim j As Long, sql As String, rs As Recordset
    Dim i As Long, c As Currency, Fcorte As Date
    Dim gcCLNE As GNComprobante
    Dim gcPVNE As GNComprobante
    On Error GoTo errtrap
    
    'Verifica las opciones
    If Not VerificarOpcion Then Exit Function
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    Fcorte = dtpFechaCorte.value    'Fecha de corte

    'Cambia figura de cursor de mouse
    MensajeStatus "Está preparando saldos a la fecha de corte...", vbHourglass
    mensaje True, "Saldo inicial de Empleados..."
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True

    'Obtiene Saldos de proveedor/cliente por cada documento pendiente
    sql = "spConsEmpSaldoSinRol 2," & FechaYMD(Fcorte, gobjMain.EmpresaActual.TipoDB)
    Set rs = mEmpOrigen.OpenRecordset(sql)
    UltimoRecordset rs
    'Abre la empresa destino
    Set e = AbrirDestino
    With rs
        If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
        i = 0
        Do Until .EOF
            prg1.value = rs.AbsolutePosition
            prg1.Refresh
            DoEvents
            MensajeStatus "Agregando detalle: #" & i & " de " & rs.RecordCount, vbHourglass
            
            'Si aplastó 'Cancelar'
            If mbooCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                GoTo cancelado
            End If
            
            
            'Si es Empleados por pagar
            If .Fields("Saldo") < 0 And .Fields("BandEmpleado") <> 0 Then
                Set pck = PrepararTransPC(e, "CLNE", _
                            "Saldo inicial de Empleados x pagar", _
                            Fcorte, gcCLNE)
                            
            'Si es Empleados por cobrar (Anticipado)
            Else
                Set pck = PrepararTransPC(e, "PVNE", _
                            "Saldo inicial de Empleados (Anticipos)", _
                            Fcorte, gcPVNE)
            End If
            
            'Recupera datos de proveedor y asigna al objeto
            pck.CodEmpleado = .Fields("CodProvCli")
            If .Fields("Saldo") > 0 Then   'Si es por cobrar --> Debe
                pck.Debe = .Fields("Saldo")        'Saldo en dólares
            Else                                    'Si es por pagar --> Haber
                pck.Haber = .Fields("Saldo") * -1     'Saldo en dólares
            End If
            pck.codforma = .Fields("CodForma")
            pck.FechaEmision = .Fields("FechaEmision")
            pck.FechaVenci = .Fields("FechaVenci")
            pck.NumLetra = .Fields("Trans")
            pck.Observacion = .Fields("Observacion")
            pck.Orden = i
            pck.Guid = .Fields("Guid")        '*** <== AGREGAR ESTO
            If Not IsNull(.Fields("CodVendedor")) Then pck.CodVendedor = .Fields("codvendedor") 'AUC 04/06/07
            i = i + 1
            .MoveNext
        Loop
        .Close
    End With
    
    'Graba la transacción si no están grabadas
    MensajeStatus "Grabándo la transacción...", vbHourglass
    
    If Not (gcCLNE Is Nothing) Then
        gcCLNE.HoraTrans = "00:00:01"
        gcCLNE.Grabar False, False
    End If
    If Not (gcPVNE Is Nothing) Then
        gcPVNE.HoraTrans = "00:00:01"
        gcPVNE.Grabar False, False
    End If
    
    MensajeStatus
    mensaje False, "", "OK"
    MsgBox "El proceso para terminó con éxito.", vbInformation
    SaldoEmpSinRol = True
    
cancelado:
    Set rs = Nothing
    MensajeStatus
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
    
    'Libera los objetos utilizados
    Set pck = Nothing
    Set gcCLNE = Nothing
    Set gcPVNE = Nothing
    Set e = Nothing
    mbooProcesando = False               'Desbloquea que se cierre la ventana
    Exit Function
errtrap:
    mensaje False, "", Err.Description
    MensajeStatus
    DispErr
    GoTo cancelado
End Function


Private Function SaldoIVSerie() As Boolean
    Dim e As Empresa, gc As GNComprobante
    Dim j As Long, n As Long
    Dim sql As String, rs As Recordset, codOrig As String
    Dim i As Long, Fcorte As Date
    Dim ivks As IVKardexSerie
    On Error GoTo errtrap
    
    'Verifica las opciones
    If Not VerificarOpcion Then Exit Function
    
    mbooProcesando = True               'Bloquea que se cierre la ventana
    
    codOrig = gobjMain.EmpresaActual.CodEmpresa
    Fcorte = dtpFechaCorte.value    'Fecha de corte

    'Cambia figura de cursor de mouse
    prg1.min = 0
    mbooCancelado = False
    cmdCancelar.Enabled = True
    
    'Saca las existencias a la fecha de corte
    MensajeStatus "Preparando para grabar las existencias iniciales de IVSeries...", vbHourglass
    mensaje True, "Saldo inicial de IVSeries..."
    
    sql = "SELECT ivk.IdSerie, ivk.IdBodega, iv.Campo1, ivb.CodBodega, Sum(ivk.Cantidad) AS Exist  " & _
                "FROM IVBodega ivb INNER JOIN (IVSerie iv " & _
                "INNER JOIN (GNTrans gt INNER JOIN (IVKardexSerie ivk INNER JOIN GNComprobante gc ON ivk.TransID=gc.TransID)  " & _
          "ON gt.CodTrans=gc.CodTrans) ON iv.IdSerie = ivk.IdSerie) ON ivb.IdBodega = ivk.IdBodega " & _
          "WHERE (gc.Estado<>" & ESTADO_ANULADO & ") AND " & _
                 "(gt.AfectaCantidad=" & CadenaBool(True, gobjMain.EmpresaActual.TipoDB) & ") AND " & _
                 "(gc.FechaTrans < " & FechaYMD(Fcorte + 1, gobjMain.EmpresaActual.TipoDB) & ") " & _
          "GROUP BY ivk.IdSerie, ivk.IdBodega, iv.Campo1, ivb.CodBodega " & _
          "HAVING Sum(ivk.Cantidad)>0"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
#If DAOLIB = 0 Then
    Set rs.ActiveConnection = Nothing
#End If
    
    'Abre la empresa destino
    Set e = AbrirDestino
    
    With rs
        If Not rs.EOF Then
            rs.MoveLast
            rs.MoveFirst
            If rs.RecordCount > 0 Then prg1.max = rs.RecordCount
            i = 0
            Do Until .EOF
                prg1.value = rs.AbsolutePosition
                prg1.Refresh
                DoEvents
                
                'Si aplastó 'Cancelar'
                If mbooCancelado Then
                    MsgBox "El proceso fue cancelado.", vbInformation
                    GoTo cancelado
                End If
                
                'Crea transaccion 'IVSI'
                If (i Mod 100) = 0 Then
                    'Si no es primera vez
                    If Not (gc Is Nothing) Then
                        'Graba la transacción
                        MensajeStatus "Grabando la transacción en la empresa '" & gc.Empresa.CodEmpresa & "'...", vbHourglass
                        gc.HoraTrans = "00:00:01"
                        gc.Grabar False, False
                    End If
                    
                    Set gc = CrearTrans(e, _
                            "IVIS", _
                            "Saldo inicial de IVSeries", _
                            Fcorte, _
                            "")
                End If
                
                'Recupera datos de inventario para llama el método Costo()
                MensajeStatus "Agregando detalle #" & i & " de " & rs.RecordCount, vbHourglass
                
                'Agrega detalle
                j = gc.AddIVKNumSerie
                Set ivks = gc.IVKNumSerie(j)
                ivks.cantidad = .Fields("Exist")
                ivks.CodBodega = .Fields("CodBodega")
                ivks.IdSerie = .Fields("idserie")
                ivks.Orden = i Mod 100
                i = i + 1
                .MoveNext
            Loop
        End If
        .Close
    End With
        
    If Not (gc Is Nothing) Then
        'Graba la transacción
        MensajeStatus "Grabando la transacción en la empresa '" & gc.Empresa.CodEmpresa & "'...", vbHourglass
        gc.HoraTrans = "00:00:01"
        gc.Grabar False, False
    End If
    
    'Corrige las existencias para que quede bien la tabla 'IVExist'
    MensajeStatus "Arreglando las existencias...", vbHourglass
    If Not (gc Is Nothing) Then
        gc.Empresa.CorregirExistenciaSerie
    End If
    mensaje False, "", "OK"
    MensajeStatus
    MsgBox "El proceso terminó con éxito.", vbInformation
    SaldoIVSerie = True
    
cancelado:
    mensaje False, "", Err.Description
    MensajeStatus
    Set ivks = Nothing
    'Set iv = Nothing
    Set gc = Nothing
    Set rs = Nothing
    prg1.value = prg1.min
    cmdCancelar.Enabled = False
    
    'Vuelve a abrir la empresa origen
    Set e = gobjMain.RecuperaEmpresa(codOrig)
    e.Abrir
    Set e = Nothing
    mbooProcesando = False                  'Desbloquea que se cierre la ventana
    Exit Function
errtrap:
    MensajeStatus
    MsgBox Err.Description, vbExclamation
    GoTo cancelado
End Function


