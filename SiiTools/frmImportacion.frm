VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportacion 
   Caption         =   "Importación de datos"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11115
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pic2 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   20100
      TabIndex        =   21
      Top             =   9840
      Width           =   20160
      Begin VB.CommandButton cmdAbrir 
         Caption         =   "&Abrir"
         Height          =   375
         Left            =   9720
         TabIndex        =   24
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label lblArchivoLocutorio 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   2280
         TabIndex        =   23
         Top             =   0
         Width           =   7335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta del Archivo de Datos: "
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   1965
      End
   End
   Begin VB.PictureBox picEncabezado 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   1536
      Left            =   0
      ScaleHeight     =   1500
      ScaleWidth      =   20130
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   20160
      Begin VB.TextBox txtDescripcion 
         Height          =   336
         Left            =   3600
         MaxLength       =   120
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         ToolTipText     =   "Descripción de la transacción"
         Top             =   480
         Width           =   4740
      End
      Begin VB.TextBox txtCotizacion 
         Height          =   336
         Left            =   900
         TabIndex        =   7
         Top             =   1080
         Width           =   1452
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   336
         Left            =   900
         TabIndex        =   9
         ToolTipText     =   "Fecha de la transacción"
         Top             =   360
         Width           =   1452
         _ExtentX        =   2566
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
      Begin FlexComboProy.FlexCombo fcbResp 
         Height          =   336
         Left            =   6888
         TabIndex        =   10
         ToolTipText     =   "Responsable de la transacción"
         Top             =   120
         Width           =   1452
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
         Height          =   336
         Left            =   3600
         TabIndex        =   11
         ToolTipText     =   "Responsable de la transacción"
         Top             =   120
         Width           =   1452
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
         Height          =   336
         Left            =   3600
         TabIndex        =   12
         ToolTipText     =   "Responsable de la transacción"
         Top             =   840
         Width           =   1452
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
         Height          =   336
         Left            =   900
         TabIndex        =   13
         ToolTipText     =   "Responsable de la transacción"
         Top             =   720
         Width           =   1452
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
      Begin FlexComboProy.FlexCombo fcbElemento 
         Height          =   330
         Left            =   6900
         TabIndex        =   25
         ToolTipText     =   "Responsable de la transacción"
         Top             =   900
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         Enabled         =   0   'False
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
      Begin VB.Label lblElemento 
         AutoSize        =   -1  'True
         Caption         =   "Rubro Ro de Pagos"
         Enabled         =   0   'False
         Height          =   195
         Left            =   5400
         TabIndex        =   26
         Top             =   960
         Width           =   1410
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "&Moneda  "
         Height          =   192
         Left            =   204
         TabIndex        =   20
         Top             =   720
         Width           =   672
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Fecha Transaccion  "
         Height          =   192
         Left            =   960
         TabIndex        =   19
         Top             =   120
         Width           =   1464
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "&Descripción  "
         Height          =   192
         Left            =   2604
         TabIndex        =   18
         Top             =   480
         Width           =   936
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "C&otización  "
         Height          =   192
         Left            =   60
         TabIndex        =   17
         Top             =   1104
         Width           =   816
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "&Responsable  "
         Height          =   192
         Left            =   5760
         TabIndex        =   16
         Top             =   120
         Width           =   1056
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cod.Trans  "
         Height          =   192
         Left            =   2712
         TabIndex        =   15
         Top             =   120
         Width           =   828
      End
      Begin VB.Label lblforma 
         AutoSize        =   -1  'True
         Caption         =   "Forma  "
         Height          =   192
         Left            =   3000
         TabIndex        =   14
         Top             =   840
         Width           =   540
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   1815
      Left            =   0
      TabIndex        =   2
      Top             =   2100
      Width           =   5895
      _cx             =   10393
      _cy             =   3196
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
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
      Editable        =   2
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   2
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   7095
      Top             =   2385
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   612
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   20160
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   10335
      Width           =   20160
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   5880
         TabIndex        =   4
         Top             =   120
         Width           =   1212
      End
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   180
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6480
      Top             =   2415
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacion.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacion.frx":0114
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacion.frx":0568
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20160
      _ExtentX        =   35560
      _ExtentY        =   1058
      ButtonWidth     =   1693
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Archivo..."
            Key             =   "Archivo"
            Object.ToolTipText     =   "Abrir archivo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Importar"
            Key             =   "Importar"
            Object.ToolTipText     =   "Importar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Grabar"
            Key             =   "Grabar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "separador"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Configurar"
            Key             =   "Configurar"
            Description     =   "Configurar"
            Object.ToolTipText     =   "Configurar opciones para importacion de Ventas para Locutorios"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cod.Trans"
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   804
   End
End
Attribute VB_Name = "frmImportacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbooCancelado As Boolean
Private mbooEjecutando As Boolean
Private mbooErrores As Boolean     '***Angel. 22/Abril/2004
Private conEncabezado As Boolean

Private Const MSG_OK As String = "OK"
Private Const MSG_ERR As String = "Error"
Private Const MSG_PROC As String = "Procesando..."

Private UltimoNumTransImportado As String

'Tipo : PLANCUENTA
'       ITEM
'       PCPROV
'       PCCLI
'       PORPAGAR
'       PORCOBRAR
'       DIARIO
'       INVENTARIO
'       AFITEM
'       CODPCGRUPO1,2,3,4,
'       CODIVGRUPO1,2,3,4,5
'       CODAFGRUPO1,2,3,4,5
Public Sub Inicio(ByVal Tipo As String)
    Me.tag = Tipo                   'Tipo de importación
    Form_Resize
    Me.Show
    Me.ZOrder
    conEncabezado = False
    pic2.Visible = False
    Select Case Me.tag
        Case "PLANCUENTA"
            Me.Caption = "Importación de datos para Plan de Cuentas"
        Case "PLANPRCUENTA"
            Me.Caption = "Importación de datos para Plan de Cuentas Presupuesto"
        Case "PLANENFERME"
            Me.Caption = "Importación de datos para Plan de Enfermedades"
        
        
        Case "ITEM"
            Me.Caption = "Importación de datos para Items"
        Case "PCPROV"
            Me.Caption = "Importación de datos para Proveedores"
        Case "PCCLI"
            Me.Caption = "Importación de datos para Clientes"
         Case "PCEMP"
            Me.Caption = "Importación de datos para Clientes"
        
        Case "PORPAGAR"
            Me.Caption = "Importación de datos para Cuentas por Pagar"
            CargarEncabezado
            Form_Resize
        Case "PORCOBRAR"
            Me.Caption = "Importación de datos para Cuentas por Cobrar"
            CargarEncabezado
            Form_Resize
        Case "DIARIO"
            Me.Caption = "Importación de Saldos Iniciales para Cuentas Contables"
            CargarEncabezado
            Form_Resize
        Case "INVENTARIO"
            Me.Caption = "Importacion de Saldos Iniciales de Inventario"
            CargarEncabezado
            Form_Resize
        Case "EXISTENCIA MINIMA"
            Me.Caption = "Editando Existencia Minima de Inventarios"
            tlb1.Buttons(2).Enabled = False
            tlb1.Buttons(3).Enabled = True
        Case "VENTASLOCUTORIOS"
            RecuperaConfig
            RecuperarUltimoNumImp
            Me.Caption = "Importacion de Ventas de Locutorios"
            tlb1.Buttons(5).Enabled = True
            pic2.Visible = True
            lblArchivoLocutorio.Caption = GetSetting(APPNAME, App.Title, "ArchivoLocutorio", "C:\Telemic\Datos\datos.txt")
        Case "AFITEM"
            Me.Caption = "Importación de datos para Activos Fijos"
        Case "AFINVENTARIO"
            Me.Caption = "Importacion de Saldos Iniciales de Activos Fijos"
            CargarEncabezado
            Form_Resize
        Case "PCGRUPO1"
            Me.Caption = "Importacion de Grupo 1 de Proveedor Cliente"
        Case "PCGRUPO2"
            Me.Caption = "Importacion de Grupo 2 de Proveedor Cliente"
        Case "PCGRUPO3"
            Me.Caption = "Importacion de Grupo 3 de Proveedor Cliente"
        Case "PCGRUPO4"
            Me.Caption = "Importacion de Grupo 4 de Proveedor Cliente"
        Case "IVGRUPO1"
            Me.Caption = "Importacion de Grupo 1 de Inventario"
        Case "IVGRUPO2"
            Me.Caption = "Importacion de Grupo 2 de Inventario"
        Case "IVGRUPO3"
            Me.Caption = "Importacion de Grupo 3 de Inventario"
        Case "IVGRUPO4"
            Me.Caption = "Importacion de Grupo 4 de Inventario"
        Case "IVGRUPO5"
            Me.Caption = "Importacion de Grupo 5 de Inventario"
        Case "IVGRUPO6"
            Me.Caption = "Importacion de Grupo 6 de Inventario"
        Case "AFGRUPO1"
            Me.Caption = "Importacion de Grupo 1 de Activo Fijo"
        Case "AFGRUPO2"
            Me.Caption = "Importacion de Grupo 2 de Activo Fijo"
        Case "AFGRUPO3"
            Me.Caption = "Importacion de Grupo 3 de Activo Fijo"
        Case "AFGRUPO4"
            Me.Caption = "Importacion de Grupo 4 de Activo Fijo"
        Case "AFGRUPO5"
            Me.Caption = "Importacion de Grupo 5 de Activo Fijo"
        Case "IVUNIDAD"
            Me.Caption = "Importacion de Unidades de Inventario"
        Case "AFINVENTARIOC"
            Me.Caption = "Importacion de Saldos Iniciales de Activos Fijos Custodios"
            CargarEncabezado
            Form_Resize
        Case "PLANCUENTASC"
            Me.Caption = "Importación de datos para Plan de Cuentas SC"
        Case "PLANCUENTAFE"
            Me.Caption = "Importación de datos para Plan de Cuentas Flujo Efectivo"
        Case "PRDIARIO"
            Me.Caption = "Importación de Saldos Iniciales para Cuentas Contables Presupuesto"
            CargarEncabezado
            Form_Resize
        Case "PORPAGAREMP"
             Me.Caption = "Importación de datos para Cuentas por Pagar para empleados"
            CargarEncabezado
            Form_Resize
        Case "PORCOBRAREMP"
             Me.Caption = "Importación de datos para Cuentas por Cobrar para empleados"
            CargarEncabezado
            Form_Resize
        Case "INVENTARIOSERIES"
            Me.Caption = "Importacion de Saldos Iniciales de Inventario Num Series"
            CargarEncabezado
            Form_Resize
    End Select
    ConfigCols
End Sub

Private Sub ConfigCols()
    Dim s As String, i As Integer
    grd.Cols = 1
    Select Case UCase$(Me.tag)
    Case "PLANCUENTA", "PLANPRCUENTA", "PLANCUENTASC", "PLANCUENTAFE", "PLANENFERME"
        s = "^#|<Código|<Nombre de cuenta|<Tipo|<Cód. Cuenta a sumar"
        grd.FormatString = s & "|<Resultado"
        grd.ColComboList(3) = "|1|2|3|4|5"      'Listado predeterminado de Tipo
    Case "ITEM"                                 ' actualizado Oliver 4/17/2001  - para que tenga existencia minima
        s = "^#|<Código|<CodAlterno1|<CodAlterno2|<Descripción|<Descripción2" & _
            "|>Precio1|>Precio2|>Precio3|>Precio4|>Precio5|>Precio6|>Precio7" & _
            "|>Existencia Minima|<Unidad medida|>%IVA|<Moneda" & _
            "|<Grupo1|<Grupo2|<Grupo3|<Grupo4|<Grupo5|<Grupo6" & _
            "|<Cód.Cuenta Activo|<Cód.Cuenta Costo|<Cód.Cuenta Venta" & _
            "|<Cód.Proveedor|<Servicio(S/N)|<Obsercación"
        Dim item As IVinventario
        grd.FormatString = s & "|<Resultado"
        grd.ColComboList(27) = "S|N"
    Case "PCPROV", "PCCLI"
        s = "^#|<Código|<Nombre|<Dirección|<Teléfono|<Teléfono2|<Teléfono3|<RUC|<Email|<Vendedor"
        s = s & "|<" & gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupo(1)
        s = s & "|<" & gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupo(2)
        s = s & "|<" & gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupo(3)
        s = s & "|<" & gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupo(4)
        grd.FormatString = s & "|<Resultado"
      Case "PCEMP"
        s = "^#|<Código|<Nombre|<Dirección|<Teléfono|<CI"
        s = s & "|<" & gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupoE(1)
        s = s & "|<" & gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupoE(2)
        s = s & "|<" & gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupoE(3)
        s = s & "|<" & gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupoE(4)
        grd.FormatString = s & "|>Sueldo Basico|>Fecha Ingreso|<Resultado"

    Case "PORPAGAR", "PORPAGAREMP"
        s = "^#|<Código Prov/Cli|<NumDoc|<Fecha|<Fecha Vencimiento|<Valor|<Observacion"
        grd.FormatString = s & "|<Resultado"
        grd.ColDataType(3) = flexDTDate
        grd.ColDataType(4) = flexDTDate

    Case "PORCOBRAR"
        s = "^#|<Código Prov/Cli|<NumDoc|<Fecha|<Fecha Vencimiento|<Valor|<Observacion"
        grd.FormatString = s & "|<Resultado"
        grd.ColDataType(3) = flexDTDate
        grd.ColDataType(4) = flexDTDate

    Case "DIARIO", "PRDIARIO"
        s = "^#|<Código Cuenta|>Debe|>Haber"
        grd.FormatString = s & "|<Resultado"

    Case "INVENTARIO"
        s = "^#|<Código Item|<Código Bodega|>Cantidad|>Costo Total"
        grd.FormatString = s & "|<Resultado"
    
    Case "EXISTENCIA MINIMA"
        s = "^#|<Código Item|>Precio1|>Precio2|>Precio3|>Precio4|>Precio5"
        grd.FormatString = s & "|<Resultado"
    Case "VENTASLOCUTORIOS"
        s = "^#|<IdCabina|<Cabina|<#Marcado|<Destino|<Trafico|<Fecha|<Hora|>Duracion|^x1|>ValorMinuto|>ICE|>Neto|>IVA|>Total|^Modificacion|^x2|^x3|^x4|>#NotaVenta|^x5|<Operador|<#Turno"
        grd.FormatString = s & "|<Resultado"
        
        grd.ColData(8) = "SubTotal"
        grd.ColData(11) = "SubTotal"
        grd.ColData(12) = "SubTotal"
        grd.ColData(13) = "SubTotal"
        grd.ColData(14) = "SubTotal"
        
        grd.Editable = flexEDNone
    Case "AFITEM"                                 ' jeaa 29/12/2008
        s = "^#|<Código|<CodAlterno1|<CodAlterno2|<Descripción" & _
            "|<Marca|<Numero Serie|>Vida Util|>Dep Anteriores" & _
            "|<Tipo Depre|>Fecha Compra|>CostoResidual|>Cod. Proveedor" & _
            "|<Unidad medida|>% IVA|<Moneda" & _
            "|<Grupo1|<Grupo2|<Grupo3|<Grupo4|<Grupo5" & _
            "|<Cód.Cuenta Activo|<Cód.Cuenta Costo|<Cód.Cuenta Venta" & _
            "|<Cód.Cuenta Depre Gasto|<Cód.Cuenta Depre Acumulada|<Cód.Cuenta Revaloriza|<Cód.Cuenta Dep Revaloriza" & _
            "|<Servicio(S/N)|<Observación"
        Dim AFitem As AFinventario
        grd.FormatString = s & "|<Resultado"
        grd.ColComboList(28) = "S|N"
    Case "AFINVENTARIO"
        s = "^#|<Código Item|<Código Bodega|>Cantidad|>Costo Compra"
        grd.FormatString = s & "|<Resultado"
    Case "PCGRUPO1", "PCGRUPO2", "PCGRUPO3", "PCGRUPO4"
        s = "^#|<Código|<Descripción"
        grd.FormatString = s & "|<Resultado"
    Case "IVGRUPO1", "IVGRUPO2", "IVGRUPO3", "IVGRUPO4", "IVGRUPO5", "IVGRUPO6"
        s = "^#|<Código|<Descripción"
        grd.FormatString = s & "|<Resultado"
    Case "AFGRUPO1", "AFGRUPO2", "AFGRUPO3", "AFGRUPO4", "AFGRUPO5"
        s = "^#|<Código|<Descripción"
        grd.FormatString = s & "|<Resultado"
    Case "AFINVENTARIOC"
        s = "^#|<Código Item|<Código Empleado|>Cantidad|>Costo Compra"
        grd.FormatString = s & "|<Resultado"
    Case "INVENTARIOSERIES"
        s = "^#|<Código Item|<Código Bodega"
'            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("EtiCampoSerie_Campo1")) > 0 Then
'                s = s & "|" & gobjMain.EmpresaActual.GNOpcion.ObtenerValor("EtiCampoSerie_Campo1")
'            Else
                s = s & "|<Campo1"
'            End If
'            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("EtiCampoSerie_Campo2")) > 0 Then
'                s = s & "|" & gobjMain.EmpresaActual.GNOpcion.ObtenerValor("EtiCampoSerie_Campo2")
'            Else
                s = s & "|<Campo2"
'            End If
'            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("EtiCampoSerie_Campo3")) > 0 Then
'                s = s & "|" & gobjMain.EmpresaActual.GNOpcion.ObtenerValor("EtiCampoSerie_Campo3")
'            Else
                s = s & "|<Campo3"
'            End If
'            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("EtiCampoSerie_Campo4")) > 0 Then
'                s = s & "|" & gobjMain.EmpresaActual.GNOpcion.ObtenerValor("EtiCampoSerie_Campo4")
'            Else
                s = s & "|<Campo4"
'            End If
'            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("EtiCampoSerie_Campo5")) > 0 Then
'                s = s & "|" & gobjMain.EmpresaActual.GNOpcion.ObtenerValor("EtiCampoSerie_Campo5")
'            Else
                s = s & "|<Campo5"
'            End If
            s = s & "|>Cantidad|>FechaCreacion"
            grd.FormatString = s & "|<Resultado"
          grd.ColHidden(9) = True
    End Select
    
    grd.ColSort(1) = flexSortGenericAscending
    grd.ColSort(2) = flexSortGenericAscending
    grd.ColSort(3) = flexSortGenericAscending
    'grd.ColSort(4) = flexSortGenericAscending
    
    'Asigna a ColKey los títulos de columnas
    ' para luego poder referirnos a la columna con su título mismo
    AsignarTituloAColKey grd
    grd.SetFocus
End Sub

Private Sub CargarEncabezado()
    picEncabezado.Visible = True
    conEncabezado = True
    dtpFecha.value = Date
    fcbResp.SetData gobjMain.EmpresaActual.ListaGNResponsable(False)
    fcbTrans.SetData gobjMain.EmpresaActual.ListaGNTrans("", False, False)
    fcbMoneda.SetData gobjMain.EmpresaActual.ListaGNMoneda
    fcbMoneda.KeyText = "USD"
    txtCotizacion.Text = "1"
    Select Case UCase(Me.tag)
    Case "PORCOBRAR"
        txtDescripcion.Text = "Saldo inicial de Cuentas x Cobrar"
        fcbForma.SetData gobjMain.EmpresaActual.ListaTSFormaCobroPago(True, True, False)
        fcbTrans.KeyText = "CLND"
        fcbForma.KeyText = "CRC"
    Case "PORPAGAR"
        txtDescripcion.Text = "Saldo inicial de Cuentas x Pagar"
        fcbForma.SetData gobjMain.EmpresaActual.ListaTSFormaCobroPago(False, True, False)
        fcbTrans.KeyText = "PVNC"
        fcbForma.KeyText = "CRP"
    Case "DIARIO"
        txtDescripcion.Text = "Saldo Inicial de Contabilidad"
        lblforma.Visible = False
        fcbForma.Visible = False
        fcbTrans.KeyText = "CTD"
    Case "INVENTARIO"
        txtDescripcion.Text = "Saldo Inicial de inventarios"
        lblforma.Visible = False
        fcbForma.Visible = False
        fcbTrans.KeyText = "IVSI"
    Case "AFINVENTARIO"
        txtDescripcion.Text = "Saldo Inicial de Activos Fijos"
        lblforma.Visible = False
        fcbForma.Visible = False
        fcbTrans.KeyText = "AFSI"
    Case "AFINVENTARIOC"
        txtDescripcion.Text = "Saldo Inicial de Activos Fijos Custodios"
        lblforma.Visible = False
        fcbForma.Visible = False
        fcbTrans.KeyText = "AFSIC"
    Case "PRDIARIO"
        txtDescripcion.Text = "Saldo Inicial de Presupuesto"
        lblforma.Visible = False
        fcbForma.Visible = False
        fcbTrans.KeyText = "PRICT"
    Case "PORPAGAREMP"
        txtDescripcion.Text = "Saldo inicial de Cuentas x Pagar Empleados"
        fcbForma.SetData gobjMain.EmpresaActual.ListaTSFormaCobroPago(False, True, False)
        fcbElemento.SetData gobjMain.EmpresaActual.ListaElementosParaFlex(True)
        fcbTrans.KeyText = "PVNE"
        fcbForma.KeyText = "CRP"
        lblElemento.Enabled = True
        fcbElemento.Enabled = True
        fcbElemento.KeyText = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("EleAplicaAnti")
   Case "PORCOBRAREMP"
        txtDescripcion.Text = "Saldo inicial de Cuentas x Cobrar Empleados"
        fcbForma.SetData gobjMain.EmpresaActual.ListaTSFormaCobroPago(False, True, False)
        fcbElemento.SetData gobjMain.EmpresaActual.ListaElementosParaFlex(True)
        fcbTrans.KeyText = "PVNE"
        fcbForma.KeyText = "CRC"
        lblElemento.Enabled = True
        fcbElemento.Enabled = True
        fcbElemento.KeyText = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("EleAplicaAnti")
   Case "INVENTARIOSERIES"
        txtDescripcion.Text = "Saldo Inicial de inventarios Num Series"
        lblforma.Visible = False
        fcbForma.Visible = False
        fcbTrans.KeyText = "IVIS"
    End Select
End Sub

Private Sub EnabledEncabezado(modo As Boolean)
    dtpFecha.Enabled = modo
    fcbResp.Enabled = modo
    fcbTrans.Enabled = modo
    fcbMoneda.Enabled = modo
    txtCotizacion.Enabled = modo
    fcbForma.Enabled = modo
End Sub



Private Sub cmdAbrir_Click()
    LeerArchivo lblArchivoLocutorio.Caption
End Sub

Private Sub cmdCancelar_Click()
    If mbooEjecutando Then
        mbooCancelado = True
    Else
        Unload Me
    End If
End Sub

Private Sub Importar()
    Dim i As Long
    On Error GoTo errtrap
    
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna fila para importar.", vbExclamation
        Exit Sub
    End If
    
    'Confirmación
    If MsgBox("Está seguro que desea comenzar el proceso de importación?", _
                vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    
    mbooEjecutando = True
    MensajeStatus "Importando...", vbHourglass
    
    With grd
        prg1.min = .FixedRows - 1
        prg1.max = .Rows - 1
        prg1.value = prg1.min
        
        For i = .FixedRows To .Rows - 1
            prg1.value = i
            DoEvents                'Para dar control a Windows
            
            'Si usuario aplastó 'Cancelar', sale del ciclo
            If mbooCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                GoTo cancelado
            End If
            
            'Si aún no está importado bien, importa la fila
            If grd.TextMatrix(i, .Cols - 1) <> MSG_OK Then
                'Si ocurre un error y no quiere seguir el usuario, sale del ciclo
                If Not ImportarFila(i) Then GoTo cancelado
            End If
        Next i
    End With
    
cancelado:
    MensajeStatus
    mbooEjecutando = False
    prg1.value = prg1.min
    Exit Sub
errtrap:
    MensajeStatus
    DispErr
    mbooEjecutando = False
    prg1.value = prg1.min
    Exit Sub
End Sub

Private Sub ImportarGNComprobante()
    Dim i As Long, j As Long, gncomp As GNComprobante, NumeroComprobante As Integer
    Dim limite As Long
    On Error GoTo errtrap
    
    ' verificar si estan todos los datos
    If Len(fcbMoneda.Text) = 0 Then
        MsgBox "Debe selecciona una tipo de Modena", vbInformation
        fcbMoneda.SetFocus
        Exit Sub
    End If
    
    If Val(txtCotizacion.Text) = 0 Then
        MsgBox "Escriba una cotizacion valida", vbInformation
        txtCotizacion.SetFocus
        Exit Sub
    End If
    
    If Len(fcbTrans.Text) = 0 Then
        MsgBox "Seleccione un tipo de transaccion", vbInformation
        fcbTrans.SetFocus
        Exit Sub
    End If
    
    If Len(txtDescripcion.Text) = 0 Then
        MsgBox "Debe escribir una Descripcion para estas transaciones", vbInformation
        txtDescripcion.SetFocus
        Exit Sub
    End If
    
    If Len(fcbResp.Text) = 0 Then
        MsgBox "Debe seleccionar un responsable", vbInformation
        fcbResp.SetFocus
        Exit Sub
    End If
    
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna fila para importar.", vbInformation
        Exit Sub
    End If
    
    'Confirmación
    If MsgBox("Está seguro que desea comenzar el proceso de importación?", _
                vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    
    ' limite = 100
    
    ' Verificar si no tiene errores si es Diario
    If Me.tag = "DIARIO" And Me.tag = "PRDIARIO" Then
        limite = -1
        For i = grd.FixedRows To grd.Rows - 1
            If Left(grd.TextMatrix(i, grd.ColIndex("Cuenta")), Len(MSG_ERR)) = MSG_ERR Then
                If MsgBox("Existen cuentas con error las cuales no seran importadas" & vbCr & _
                          "esta seguro que desea importar?", vbYesNo + vbQuestion) <> vbYes Then
                          Exit Sub
                Else
                    Exit For
                End If
            End If
        Next i
    End If
    
    
    
    
    mbooEjecutando = True
    MensajeStatus "Importando...", vbHourglass
    
    EnabledEncabezado False
    grd.Enabled = False
    With grd
        prg1.min = .FixedRows - 1
        prg1.max = .Rows - 1
        prg1.value = prg1.min
        Set gncomp = gobjMain.EmpresaActual.CreaGNComprobante(fcbTrans.KeyText)
        
        NumeroComprobante = 1
        PonerDatosComprobante gncomp, NumeroComprobante

'*********************************************
'borrar mensajes de msg_ok en ultima columna
'*********************************************
        limite = Val(GetSetting(APPNAME, SECTION, "LimitedeImportacion", "100"))
        'no puede quedar el limite en 0  o 1, seria absurdo
        If limite < 3 Then limite = 100 ''entonces asigno valor por defecto en 100
        
        .Cell(flexcpText, .FixedRows, .Cols - 1, .Rows - 1, .Cols - 1) = ""
        j = 1
        For i = .FixedRows To .Rows - 1
            prg1.value = i
            DoEvents                'Para dar control a Windows
            
            'Si usuario aplastó 'Cancelar', sale del ciclo
            If mbooCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                GoTo cancelado
            End If

            'Si aún no está importado bien, importa la fila
            If grd.TextMatrix(i, .Cols - 1) <> MSG_OK Then
                'Si ocurre un error y no quiere seguir el usuario, sale del ciclo
                If Not ImportarFilaDetalleGNcomp(i, gncomp) Then GoTo cancelado
            End If
            
            If gncomp.GNTrans.IVNumFilaMax <> 0 Then limite = gncomp.GNTrans.IVNumFilaMax
            
            If j = limite Then
                If GrabarGNComprobante(gncomp) Then
                    NumeroComprobante = NumeroComprobante + 1
                    Set gncomp = Nothing
                    Set gncomp = gobjMain.EmpresaActual.CreaGNComprobante(fcbTrans.KeyText)
                    PonerDatosComprobante gncomp, NumeroComprobante
                End If
                j = 0
            End If
            j = j + 1
        Next i
        
        ' graba el ultimo que no queda grabado
        
        GrabarGNComprobante gncomp
        
    End With
        
cancelado:
    Set gncomp = Nothing
    MensajeStatus
    mbooEjecutando = False
    prg1.value = prg1.min
    grd.Enabled = True
    EnabledEncabezado True
    Exit Sub
errtrap:
    MensajeStatus
    DispErr
    mbooEjecutando = False
    prg1.value = prg1.min
    grd.Enabled = True
    EnabledEncabezado True
    Exit Sub
End Sub

Private Function GrabarGNComprobante(ByRef gncomp As GNComprobante) As Boolean
    Dim i As Long
    i = 0
    Select Case UCase(Me.tag)
        Case "PORCOBRAR"
            i = gncomp.CountPCKardex
        Case "PORPAGAR"
            i = gncomp.CountPCKardex
        Case "DIARIO"
            i = gncomp.CountCTLibroDetalle
        Case "INVENTARIO"
            i = gncomp.CountIVKardex
        Case "AFINVENTARIO"
            i = gncomp.CountAFKardex
        Case "AFINVENTARIOC"
            i = gncomp.CountAFKardexCustodio
        Case "PRDIARIO"
            i = gncomp.CountPRLibroDetalle
        Case "PORPAGAREMP"
            i = gncomp.CountPCKardex
        Case "INVENTARIOSERIES"
            i = gncomp.CountIVKNumSerie
    End Select
    If i > 0 Then
        gncomp.Grabar False, False
        MsgBox "Se ha grabado un comprobante con numero = " & gncomp.CodTrans & gncomp.numtrans, vbInformation
        GrabarGNComprobante = True
    Else
        GrabarGNComprobante = False
    End If
End Function

Private Function ImportarFilaDetalleGNcomp(ByVal i As Long, ByRef gncomp As GNComprobante) As Boolean
    ImportarFilaDetalleGNcomp = True
    If Len(grd.TextMatrix(i, 1)) = 0 Then Exit Function
    If grd.IsSubtotal(i) Then Exit Function ' no importa la fila que tiene subtotal
    Select Case UCase$(Me.tag)
    Case "PORCOBRAR"
        ImportarFilaDetalleGNcomp = ImportarFilaPorCobrarPagar(i, gncomp, True)
    Case "PORPAGAR"
        ImportarFilaDetalleGNcomp = ImportarFilaPorCobrarPagar(i, gncomp, False)
    Case "DIARIO"
        If Left(grd.TextMatrix(i, grd.ColIndex("Cuenta")), Len(MSG_ERR)) <> MSG_ERR Then
            ImportarFilaDetalleGNcomp = ImportarFilaDiario(i, gncomp)
        Else
            grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
        End If
    Case "INVENTARIO"
        ImportarFilaDetalleGNcomp = ImportarFilaInventario(i, gncomp)
    Case "AFINVENTARIO"
        ImportarFilaDetalleGNcomp = ImportarFilaAFInventario(i, gncomp)
    Case "AFINVENTARIOC"
        ImportarFilaDetalleGNcomp = ImportarFilaAFInventarioC(i, gncomp)
    Case "PRDIARIO"
        If Left(grd.TextMatrix(i, grd.ColIndex("Cuenta")), Len(MSG_ERR)) <> MSG_ERR Then
            ImportarFilaDetalleGNcomp = ImportarFilaPRDiario(i, gncomp)
        Else
            grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
        End If
    Case "PORPAGAREMP"
        ImportarFilaDetalleGNcomp = ImportarFilaPorCobrarPagarEmp(i, gncomp, False)
    Case "INVENTARIOSERIES"
        ImportarFilaDetalleGNcomp = ImportarFilaInventarioSerie(i, gncomp)
    End Select

End Function


Sub PonerDatosComprobante(ByRef gncomp As GNComprobante, ByVal Num As Integer)
'    gncomp.CodTrans = fcbTrans.KeyText
    gncomp.numtrans = Num
    gncomp.FechaTrans = dtpFecha.value
    gncomp.CodResponsable = fcbResp.KeyText
    gncomp.CodMoneda = fcbMoneda.KeyText
    gncomp.Cotizacion(fcbMoneda.KeyText) = Val(txtCotizacion.Text)
    gncomp.Descripcion = txtDescripcion.Text
End Sub

Private Function ImportarFila(ByVal i As Long) As Boolean
    ImportarFila = True
    If Len(grd.TextMatrix(i, 1)) = 0 Then Exit Function
    Select Case UCase$(Me.tag)
    Case "PLANCUENTA"
        ImportarFila = ImportarFilaCuenta(i)
    Case "PLANPRCUENTA"
        ImportarFila = ImportarFilaPRCuenta(i)
    Case "PLANCUENTASC"
        ImportarFila = ImportarFilaCuentaSC(i)
    Case "PLANCUENTAFE"
        ImportarFila = ImportarFilaCuentaFE(i)
    
    Case "PLANENFERME"
        ImportarFila = ImportarFilaEnfermedades(i)
    
    Case "ITEM"
        ImportarFila = ImportarFilaItem(i)
    Case "PCPROV"
        ImportarFila = ImportarFilaProvCli(i, True)  'MISMA FUNCION SOLO QUE
    Case "PCCLI"
        ImportarFila = ImportarFilaProvCli(i, False) ' ENVIA TRUE PARA PROVEEDOR Y FALSE PARA CLIENTES
    Case "PCEMP"
        ImportarFila = ImportarFilaEmp(i) ' ENVIA TRUE PARA PROVEEDOR Y FALSE PARA CLIENTES
    Case "AFITEM"
        ImportarFila = ImportarFilaAFItem(i)
    Case "PCGRUPO1"
        ImportarFila = ImportarFilaPCGrupo(1, i)
    Case "PCGRUPO2"
        ImportarFila = ImportarFilaPCGrupo(2, i)
    Case "PCGRUPO3"
        ImportarFila = ImportarFilaPCGrupo(3, i)
    Case "PCGRUPO4"
        ImportarFila = ImportarFilaPCGrupo(4, i)
    Case "IVGRUPO1"
        ImportarFila = ImportarFilaIVGrupo(1, i)
    Case "IVGRUPO2"
        ImportarFila = ImportarFilaIVGrupo(2, i)
    Case "IVGRUPO3"
        ImportarFila = ImportarFilaIVGrupo(3, i)
    Case "IVGRUPO4"
        ImportarFila = ImportarFilaIVGrupo(4, i)
    Case "IVGRUPO5"
        ImportarFila = ImportarFilaIVGrupo(5, i)
    Case "IVGRUPO6"
        ImportarFila = ImportarFilaIVGrupo(6, i)
    
    Case "AFGRUPO1"
        ImportarFila = ImportarFilaAFGrupo(1, i)
    Case "AFGRUPO2"
        ImportarFila = ImportarFilaAFGrupo(2, i)
    Case "AFGRUPO3"
        ImportarFila = ImportarFilaAFGrupo(3, i)
    Case "AFGRUPO4"
        ImportarFila = ImportarFilaAFGrupo(4, i)
    Case "AFGRUPO5"
        ImportarFila = ImportarFilaAFGrupo(5, i)
    
    End Select

End Function

Private Function ImportarFilaCuenta(ByVal i As Long) As Boolean
    Dim ct As CtCuenta, msg As String
    Dim ct_Aux As CtCuenta, nivelPadre As Integer
    On Error GoTo errtrap
    
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_PROC
    
    Set ct = gobjMain.EmpresaActual.CreaCTCuenta
    With ct
        .codcuenta = grd.TextMatrix(i, grd.ColIndex("Código"))
        .NombreCuenta = grd.TextMatrix(i, grd.ColIndex("Nombre de cuenta"))
        .TipoCuenta = Val(grd.TextMatrix(i, grd.ColIndex("Tipo")))
        .CodCuentaSuma = grd.TextMatrix(i, grd.ColIndex("Cód. Cuenta a sumar"))
        'jeaa 24/09/04 para modificar el campo la cuenta de total la cta cuentaSuma y obtener el nivel del padre
        If Len(.CodCuentaSuma) > 0 Then
            Set ct_Aux = gobjMain.EmpresaActual.RecuperaCTCuenta(.CodCuentaSuma)
            If Not ct_Aux Is Nothing Then
                ct_Aux.BandTotal = True
                .nivel = ct_Aux.nivel + 1
                ct_Aux.Grabar
                Set ct_Aux = Nothing
            End If
        End If
        .Grabar
        
        'Saca mensaje en columna de resultado
        grd.TextMatrix(i, grd.Cols - 1) = MSG_OK
    End With
    
    ImportarFilaCuenta = True
    Exit Function
errtrap:
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    
    msg = "Ha ocurrido un error al tratar de importar la fila #" & i & "." & vbCr & _
          "Código : " & grd.TextMatrix(i, 1) & vbCr & _
          "Error : " & Err.Description & vbCr & vbCr & _
          "Desea continuar el proceso desde la siguiente fila?"
    If MsgBox(msg, vbYesNo + vbExclamation) = vbYes Then
        ImportarFilaCuenta = True
    Else
        ImportarFilaCuenta = False
    End If
    Exit Function
End Function

Private Function ImportarFilaItem(ByVal i As Long) As Boolean
    Dim msg As String, iv As IVinventario
    On Error GoTo errtrap
    
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_PROC
    
    Set iv = gobjMain.EmpresaActual.CreaIVInventario
    With iv
        .CodInventario = grd.TextMatrix(i, grd.ColIndex("Código"))
        .CodAlterno1 = grd.TextMatrix(i, grd.ColIndex("CodAlterno1"))
        .CodAlterno2 = (grd.TextMatrix(i, grd.ColIndex("CodAlterno2")))
        .Descripcion = grd.TextMatrix(i, grd.ColIndex("Descripción"))
        .Descripcion2 = grd.TextMatrix(i, grd.ColIndex("Descripción2"))
        .Precio(1) = grd.ValueMatrix(i, grd.ColIndex("Precio1"))
        .Precio(2) = grd.ValueMatrix(i, grd.ColIndex("Precio2"))
        .Precio(3) = grd.ValueMatrix(i, grd.ColIndex("Precio3"))
        .Precio(4) = grd.ValueMatrix(i, grd.ColIndex("Precio4"))
        .Precio(5) = grd.ValueMatrix(i, grd.ColIndex("Precio5"))
        .Precio(6) = grd.ValueMatrix(i, grd.ColIndex("Precio6"))
        .Precio(7) = grd.ValueMatrix(i, grd.ColIndex("Precio7"))
        .CodUnidad = grd.TextMatrix(i, grd.ColIndex("Unidad medida"))
        .CodUnidadConteo = grd.TextMatrix(i, grd.ColIndex("Unidad medida"))
        .ExistenciaMinima = grd.ValueMatrix(i, grd.ColIndex("Existencia Minima"))
        .PorcentajeIVA = grd.ValueMatrix(i, grd.ColIndex("%IVA")) / 100
        If grd.ValueMatrix(i, grd.ColIndex("%IVA")) / 100 = 0 Then
            .bandIVA = False
        Else
            .bandIVA = True
        End If
        .CodMoneda = grd.TextMatrix(i, grd.ColIndex("Moneda"))
        .CodGrupo(1) = grd.TextMatrix(i, grd.ColIndex("Grupo1"))
        .CodGrupo(2) = grd.TextMatrix(i, grd.ColIndex("Grupo2"))
        .CodGrupo(3) = grd.TextMatrix(i, grd.ColIndex("Grupo3"))
        .CodGrupo(4) = grd.TextMatrix(i, grd.ColIndex("Grupo4"))
        .CodGrupo(5) = grd.TextMatrix(i, grd.ColIndex("Grupo5"))
        .CodGrupo(6) = grd.TextMatrix(i, grd.ColIndex("Grupo6"))
        .CodCuentaActivo = grd.TextMatrix(i, grd.ColIndex("Cód.Cuenta Activo"))
        .CodCuentaCosto = grd.TextMatrix(i, grd.ColIndex("Cód.Cuenta Costo"))
        .CodCuentaVenta = grd.TextMatrix(i, grd.ColIndex("Cód.Cuenta Venta"))
        .codProveedor = grd.TextMatrix(i, grd.ColIndex("Cód.Proveedor"))
        If grd.TextMatrix(i, grd.ColIndex("Servicio(S/N)")) = "S" Then
            .BandServicio = True
        Else
            .BandServicio = False
        End If
        
        .Observacion = grd.TextMatrix(i, grd.ColIndex("Obsercación"))
        .Grabar
        
        'Saca mensaje en columna de resultado
        grd.TextMatrix(i, grd.Cols - 1) = MSG_OK
    End With
    
    ImportarFilaItem = True
    Exit Function
errtrap:
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    
    msg = "Ha ocurrido un error al tratar de importar la fila #" & i & "." & vbCr & _
          "Código : " & grd.TextMatrix(i, 1) & vbCr & _
          "Error : " & Err.Description & vbCr & vbCr & _
          "Desea continuar el proceso desde la siguiente fila?"
    If MsgBox(msg, vbYesNo + vbExclamation) = vbYes Then
        ImportarFilaItem = True
    Else
        ImportarFilaItem = False
    End If
    Exit Function
End Function

Private Function ImportarFilaProvCli(ByVal i As Long, BandProv As Boolean) As Boolean
    Dim msg As String, pc As PCProvCli, Cliprov As String, cad As String
    On Error GoTo errtrap
    
    grd.TextMatrix(i, grd.Cols - 1) = MSG_PROC
    Set pc = gobjMain.EmpresaActual.RecuperaPCProvCli(grd.TextMatrix(i, grd.ColIndex("Código")))
    
    If Not (pc Is Nothing) Then
        If pc.BandCliente Then Cliprov = "Cliente"
        If pc.BandProveedor Then
            If Len(Cliprov) > 0 Then Cliprov = Cliprov & " y "
            Cliprov = Cliprov & "Proveedor"
        End If
        
        cad = IIf((BandProv = True), "Proveedor", "Cliente")
        
        msg = "Ya existe un " & Cliprov & " con el codigo " & _
                pc.CodProvCli & " y nombre : " & pc.nombre & vbCr & vbCr & _
                "esta seguro que desea sobreescribirlos por datos de " & cad
                
        If MsgBox(msg, vbYesNo + vbExclamation) = vbNo Then
            grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
            ImportarFilaProvCli = True
            Exit Function
        End If
    Else
        Set pc = gobjMain.EmpresaActual.CreaPCProvCli
    End If
    

    With pc

        .CodProvCli = grd.TextMatrix(i, grd.ColIndex("Código"))
        .nombre = grd.TextMatrix(i, grd.ColIndex("Nombre"))
        .Direccion1 = grd.TextMatrix(i, grd.ColIndex("Dirección"))
        .Telefono1 = grd.TextMatrix(i, grd.ColIndex("Teléfono"))
        .Telefono2 = grd.TextMatrix(i, grd.ColIndex("Teléfono2"))
        .Telefono3 = grd.TextMatrix(i, grd.ColIndex("Teléfono3"))
        .CodVendedor = grd.TextMatrix(i, grd.ColIndex("Vendedor"))
        .Email = grd.TextMatrix(i, grd.ColIndex("Email"))
        If BandProv Then
            .BandProveedor = True
        Else
            .BandCliente = True
        End If
        .ruc = grd.TextMatrix(i, grd.ColIndex("RUC"))
        If Len(.ruc) = 13 Then
            .codtipoDocumento = "R"
        ElseIf Len(.ruc) = 10 Then
            .codtipoDocumento = "C"
        End If
        .CodGrupo1 = grd.TextMatrix(i, grd.ColIndex(gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupo(1)))
        .CodGrupo2 = grd.TextMatrix(i, grd.ColIndex(gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupo(2)))
        .CodGrupo3 = grd.TextMatrix(i, grd.ColIndex(gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupo(3)))
        .CodGrupo4 = grd.TextMatrix(i, grd.ColIndex(gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupo(4)))
        .Grabar

        'Saca mensaje en columna de resultado
        grd.TextMatrix(i, grd.Cols - 1) = MSG_OK
    End With
    ImportarFilaProvCli = True
    Exit Function

errtrap:
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    
    msg = "Ha ocurrido un error al tratar de importar la fila #" & i & "." & vbCr & _
          "Error : " & Err.Description & vbCr & vbCr & _
          "Desea continuar el proceso desde la siguiente fila?"
    If MsgBox(msg, vbYesNo + vbExclamation) = vbYes Then
        ImportarFilaProvCli = True
    Else
        ImportarFilaProvCli = False
    End If
    Exit Function
End Function


Private Function ImportarFilaPorCobrarPagar(ByVal i As Long, ByRef gncomp As GNComprobante, ByVal bandCobrar As Boolean) As Boolean
    Dim msg As String, kardex As PCKardex, ix As Long
    On Error GoTo errtrap
    ix = gncomp.AddPCKardex
    Set kardex = gncomp.PCKardex(ix)
    With kardex
        .CodProvCli = grd.TextMatrix(i, grd.ColIndex("Código Prov/Cli"))
        .NumLetra = grd.TextMatrix(i, grd.ColIndex("numdoc"))
        .FechaEmision = grd.TextMatrix(i, grd.ColIndex("Fecha"))
        .FechaVenci = grd.TextMatrix(i, grd.ColIndex("Fecha Vencimiento"))
        If bandCobrar Then
            .Debe = grd.TextMatrix(i, grd.ColIndex("Valor"))
        Else
            .Haber = grd.TextMatrix(i, grd.ColIndex("Valor"))
        End If
        .Observacion = grd.TextMatrix(i, grd.ColIndex("Observacion"))
        .codforma = fcbForma.KeyText
        .Orden = ix
        grd.TextMatrix(i, grd.Cols - 1) = MSG_OK
    End With
    ImportarFilaPorCobrarPagar = True
    Exit Function
errtrap:
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    
    msg = "Ha ocurrido un error al tratar de importar la fila #" & i & "." & vbCr & _
          "Código : " & grd.TextMatrix(i, 1) & vbCr & _
          "Error : " & Err.Description & vbCr & vbCr & _
          "Desea continuar el proceso desde la siguiente fila?"
    If MsgBox(msg, vbYesNo + vbExclamation) = vbYes Then
        ImportarFilaPorCobrarPagar = True
    Else
        ImportarFilaPorCobrarPagar = False
    End If
    gncomp.RemovePCKardex ix, kardex
    Exit Function
End Function

Private Function ImportarFilaDiario(ByVal i As Long, ByRef gncomp As GNComprobante) As Boolean
    Dim msg As String, CTdiario As CTLibroDetalle, ix As Long
    On Error GoTo errtrap
    ix = gncomp.AddCTLibroDetalle
    Set CTdiario = gncomp.CTLibroDetalle(ix)
    grd.TextMatrix(i, grd.Cols - 1) = MSG_PROC
    With CTdiario
        'AÑADIR DATOS A GRABAR EN CTDIARIO
        .codcuenta = grd.TextMatrix(i, grd.ColIndex("Código Cuenta"))
        .Descripcion = txtDescripcion.Text
        .Debe = Val(grd.TextMatrix(i, grd.ColIndex("Debe")))
        .Haber = Val(grd.TextMatrix(i, grd.ColIndex("Haber")))
        .Orden = ix
'        Debug.Print i; ix
    End With
    grd.TextMatrix(i, grd.Cols - 1) = MSG_OK
    ImportarFilaDiario = True
    Exit Function
errtrap:
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    
    msg = "Ha ocurrido un error al tratar de importar la fila #" & i & "." & vbCr & _
          "Código : " & grd.TextMatrix(i, 1) & vbCr & _
          "Error : " & Err.Description & vbCr & vbCr & _
          "Desea continuar el proceso desde la siguiente fila?"
    If MsgBox(msg, vbYesNo + vbExclamation) = vbYes Then
        ImportarFilaDiario = True
    Else
        ImportarFilaDiario = False
    End If
    gncomp.RemoveCTLibroDetalle ix, CTdiario
    Exit Function
End Function

Private Function ImportarFilaInventario(ByVal i As Long, ByRef gncomp As GNComprobante) As Boolean
    Dim msg As String, IVinventario As IVKardex, ix As Long, item As IVinventario, arancel As IVRecargoArancel
    Dim ICE As IVRecargoICE
    On Error GoTo errtrap
    ix = gncomp.AddIVKardex
    Set IVinventario = gncomp.IVKardex(ix)
    grd.TextMatrix(i, grd.Cols - 1) = MSG_PROC
    With IVinventario
        'AÑADIR DATOS A GRABAR EN ivinventario
        .CodInventario = grd.TextMatrix(i, grd.ColIndex("Código Item"))
        .CodBodega = grd.TextMatrix(i, grd.ColIndex("Código Bodega"))
        If gncomp.GNTrans.IVTipoTrans = "I" Or gncomp.GNTrans.IVTipoTrans = "X" Then
            If gncomp.Empresa.GNOpcion.IVKTipoDatoDouble Then
                .CantidadDou = grd.ValueMatrix(i, grd.ColIndex("Cantidad"))
                .CostoTotaldou = grd.ValueMatrix(i, grd.ColIndex("Costo Total"))
                .CostoRealTotaldou = grd.ValueMatrix(i, grd.ColIndex("Costo Total"))
            Else
                .cantidad = grd.ValueMatrix(i, grd.ColIndex("Cantidad"))
                .CostoTotal = grd.ValueMatrix(i, grd.ColIndex("Costo Total"))
                .CostoRealTotal = grd.ValueMatrix(i, grd.ColIndex("Costo Total"))
            End If
        ElseIf gncomp.GNTrans.IVTipoTrans = "E" Then
            If gncomp.Empresa.GNOpcion.IVKTipoDatoDouble Then
                .CantidadDou = grd.ValueMatrix(i, grd.ColIndex("Cantidad")) * -1
                .CostoTotaldou = grd.ValueMatrix(i, grd.ColIndex("Costo Total")) * -1
                .CostoRealTotaldou = grd.ValueMatrix(i, grd.ColIndex("Costo Total")) * -1
            Else
                .cantidad = grd.ValueMatrix(i, grd.ColIndex("Cantidad")) * -1
                .CostoTotal = grd.ValueMatrix(i, grd.ColIndex("Costo Total")) * -1
                .CostoRealTotal = grd.ValueMatrix(i, grd.ColIndex("Costo Total")) * -1
            End If
        End If
        Set item = gncomp.Empresa.RecuperaIVInventario(.CodInventario)
        If Not item Is Nothing Then
            .IVA = item.PorcentajeIVA
            If Len(item.CodArancel) > 0 Then
                Set arancel = gncomp.Empresa.RecuperaARANCEL(item.CodArancel)
                If Not arancel Is Nothing Then
                    .arancel = arancel.porcentaje
                    .FODIN = arancel.RecarPorcentaje
                End If
            End If
            If Len(item.CodICE) > 0 Then
            Set ICE = gncomp.Empresa.RecuperaICE(item.CodICE)
                If Not ICE Is Nothing Then
                    .ICE = ICE.porcentaje
                    .CodICE = item.CodICE
                End If
            End If
            
        End If
        Set ICE = Nothing
        Set arancel = Nothing
        Set item = Nothing
        .Orden = ix
    End With
    grd.TextMatrix(i, grd.Cols - 1) = MSG_OK
    ImportarFilaInventario = True
    Exit Function
errtrap:
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    
    msg = "Ha ocurrido un error al tratar de importar la fila #" & i & "." & vbCr & _
          "Código : " & grd.TextMatrix(i, 1) & vbCr & _
          "Error : " & Err.Description & vbCr & vbCr & _
          "Desea continuar el proceso desde la siguiente fila?"
    If MsgBox(msg, vbYesNo + vbExclamation) = vbYes Then
        ImportarFilaInventario = True
    Else
        ImportarFilaInventario = False
    End If
    gncomp.RemoveIVKardex ix, IVinventario
    Exit Function
End Function






Private Sub fcbTrans_Selected(ByVal Text As String, ByVal KeyText As String)
    Dim gnt As GNTrans
    Set gnt = gobjMain.EmpresaActual.RecuperaGNTrans(fcbTrans.KeyText)
    txtDescripcion.Text = gnt.NombreTrans
    Set gnt = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF9
'        cmdImprimir_Click
        KeyCode = 0
    Case Else
        MoverCampo Me, KeyCode, Shift, True
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ImpideSonidoEnter Me, KeyAscii
End Sub

Private Sub Form_Load()
    grd.Rows = grd.FixedRows    'Limpia la grilla
End Sub

Private Sub Form_Resize()
    Dim hei As Long
    On Error Resume Next
    hei = IIf(conEncabezado, picEncabezado.Height, 0)
    If Me.tag = "VENTASLOCUTORIOS" Then
        grd.Move 0, tlb1.Height + hei, Me.ScaleWidth, Me.ScaleHeight - (tlb1.Height + pic1.Height + pic2.Height + hei + 200)
    Else
        grd.Move 0, tlb1.Height + hei, Me.ScaleWidth, Me.ScaleHeight - tlb1.Height - pic1.Height - hei
    End If
    cmdCancelar.Move Me.ScaleWidth - cmdCancelar.Width - 160
    prg1.Width = Me.ScaleWidth - (prg1.Left * 2) - cmdCancelar.Width - 160
End Sub



Private Sub Form_Unload(Cancel As Integer)
    If Me.tag = "VENTASLOCUTORIOS" Then
        'graba el ultimo numero de nota de venta de Locutorios Importado
        GrabarUltimoNumImp
    End If
End Sub

Private Sub grd_AfterEdit(ByVal Row As Long, ByVal col As Long)

    If Me.tag = "DIARIO" Then
        If col = 1 Then
            grd.TextMatrix(Row, grd.ColIndex("Cuenta")) = ponerCuentaFila(grd.TextMatrix(Row, col))
        End If
        Sumar
    ElseIf Me.tag = "PRDIARIO" Then
        If col = 1 Then
            grd.TextMatrix(Row, grd.ColIndex("Cuenta")) = ponerCuentaFila(grd.TextMatrix(Row, col))
        End If
        Sumar
    ElseIf Me.tag = "INVENTARIO" Then
        Select Case col
        Case 1
            grd.TextMatrix(Row, grd.ColIndex("Descripción")) = ponerDescripcionFila(grd.TextMatrix(Row, col))
        Case 5
            grd.TextMatrix(Row, grd.ColIndex("Costo Total")) = ponerCostoTotal(Row)
        Case 6
             grd.TextMatrix(Row, grd.ColIndex("Costo Unitario")) = ponerCostoUnitarioFila(Row)
        End Select
    ElseIf Me.tag = "AFINVENTARIO" Then
        Select Case col
        Case 1
            grd.TextMatrix(Row, grd.ColIndex("Descripción")) = ponerDescripcionFilaAF(grd.TextMatrix(Row, col))
''        Case 5
''            grd.TextMatrix(Row, grd.ColIndex("Costo Total")) = ponerCostoTotal(Row)
''        Case 6
''             grd.TextMatrix(Row, grd.ColIndex("Costo Unitario")) = ponerCostoUnitarioFila(Row)
        End Select
    
    End If
End Sub

Private Sub grd_BeforeEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
    If grd.IsSubtotal(Row) Then Cancel = True
    If grd.ColIndex("Resultado") = col Then Cancel = True
    If grd.ColIndex("Cuenta") = col Then Cancel = True
    If Me.tag = "INVENTARIO" And grd.ColIndex("Descripción") = col Then
        Cancel = True
    ElseIf Me.tag = "AFINVENTARIO" And grd.ColIndex("Descripción") = col Then
        Cancel = True
    End If
End Sub

Private Sub grd_KeyDown(KeyCode As Integer, Shift As Integer)
    If grd.IsSubtotal(grd.Row) Or Me.tag = "VENTASLOCUTORIOS" Then Exit Sub    ' si la fila es de subtotal o
    Select Case KeyCode                                                       ' esta importanto las ventas del locutorio para el Sii no las puede editar
    Case vbKeyInsert
        AgregarFila
    Case vbKeyDelete
        EliminarFila
    End Select
End Sub

Private Sub AgregarFila()
    On Error GoTo errtrap
    If Me.tag = "DIARIO" Then
        If grd.ColIndex("Cuenta") = -1 Then InsertarColumnaCuenta
    ElseIf Me.tag = "INVENTARIO" Then
        If grd.ColIndex("Descripción") = -1 Then InsertarColumnaDesc_y_Cost
    ElseIf Me.tag = "AFINVENTARIO" Then
        If grd.ColIndex("Descripción") = -1 Then InsertarColumnaDesc
    ElseIf Me.tag = "AFINVENTARIOC" Then
        If grd.ColIndex("Descripción") = -1 Then InsertarColumnaDesc
    ElseIf Me.tag = "PRDIARIO" Then
        If grd.ColIndex("Cuenta") = -1 Then InsertarColumnaCuenta
     ElseIf Me.tag = "INVENTARIOSERIES" Then
        If grd.ColIndex("Descripción") = -1 Then InsertarColumnaDescSeries
    End If
    With grd
        .AddItem "", .Row + 1
        GNPoneNumFila grd, False
        .Row = .Row + 1
        .col = .FixedCols
    End With
    
    AjustarAutoSize grd, -1, -1
    grd.SetFocus
    Exit Sub
errtrap:
    MsgBox Err.Description
    grd.SetFocus
    Exit Sub
End Sub

Private Sub EliminarFila()
    On Error GoTo errtrap
    If grd.Row <> grd.FixedRows - 1 And Not grd.IsSubtotal(grd.Row) Then
        grd.RemoveItem grd.Row
        GNPoneNumFila grd, False
    End If
    grd.SetFocus
    Exit Sub
errtrap:
    MsgBox Err.Description
    grd.SetFocus
    Exit Sub
End Sub


Private Sub tlb1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo errtrap

    Select Case Button.Key
    Case "Archivo"
        AbrirArchivo
    Case "Importar"
        If conEncabezado Then   ' los documentos con encabezado tiene formato gncomprobante
            ImportarGNComprobante
        Else
            If Me.tag = "VENTASLOCUTORIOS" Then
                ImportarVentasLocutorio
            Else
                Importar
            End If
        End If
    Case "Grabar"
        Grabar
    Case "Configurar"
        frmConfiguracion.Inicio
    End Select
    Exit Sub
errtrap:
    DispErr
    Exit Sub
End Sub

Private Sub AbrirArchivo()
    Dim i As Long
    On Error GoTo errtrap
    With dlg1
        .CancelError = True
'        .Filter = "Texto (Separado por coma)|*.txt|Excel 97(XLS)|*.xls"
        .Filter = "Texto (Separado por coma)|*.txt"
        .flags = cdlOFNFileMustExist
        If Me.tag = "VENTASLOCUTORIOS" Then .filename = lblArchivoLocutorio.Caption
        If Len(.filename) = 0 Then          'Solo por primera vez, ubica a la carpeta de la aplicación
            .filename = App.Path & "\*.txt"
        End If
        
        .ShowOpen
        If Me.tag = "VENTASLOCUTORIOS" Then lblArchivoLocutorio.Caption = dlg1.filename
        LeerArchivo (dlg1.filename)
    End With
    Exit Sub
errtrap:
    If Err.Number <> 32755 Then DispErr
    Exit Sub
End Sub

Private Sub LeerArchivo(ByVal archi As String)
    Select Case UCase$(Right$(archi, 4))
        Case ".TXT"
            ReformartearColumnas
            VisualizarTexto archi
            InsertarColumnas
        Case ".XLS"
            VisualizarExcel archi
        Case Else
        End Select
End Sub

Private Sub ReformartearColumnas()
' SOLO EN ESTOS CASOS
Select Case UCase(Me.tag)
    Case "DIARIO", "INVENTARIO", "AFINVENTARIO", "AFINVENTARIOC", "PRDIARIO"
        ConfigCols
End Select

End Sub

Private Sub InsertarColumnas()
Dim i As Integer
Select Case UCase(Me.tag)
    Case "DIARIO"
        'sumar
        With grd
        
        InsertarColumnaCuenta
        For i = .FixedRows To .Rows - 1 ' poner nombre de cuentas en columna cuenta
            DoEvents
            .TextMatrix(i, .ColIndex("Cuenta")) = ponerCuentaFila(.TextMatrix(i, 1))
        Next i
        Sumar
        End With
    Case "INVENTARIO"
        InsertarColumnaDesc_y_Cost
        For i = grd.FixedRows To grd.Rows - 1 ' poner nombre de cuentas en columna cuenta
            DoEvents
            grd.TextMatrix(i, grd.ColIndex("Descripción")) = ponerDescripcionFila(grd.TextMatrix(i, 1))
            grd.TextMatrix(i, grd.ColIndex("Costo Unitario")) = ponerCostoUnitarioFila(i)
        Next i
    Case "AFINVENTARIO"
        InsertarColumnaDesc_y_Cost
        For i = grd.FixedRows To grd.Rows - 1 ' poner nombre de cuentas en columna cuenta
            DoEvents
            grd.TextMatrix(i, grd.ColIndex("Descripción")) = ponerDescripcionFilaAF(grd.TextMatrix(i, 1))
'            grd.TextMatrix(i, grd.ColIndex("Costo Unitario")) = ponerCostoUnitarioFila(i)
        Next i
    Case "AFINVENTARIOC"
        InsertarColumnaDesc_y_Cost
        For i = grd.FixedRows To grd.Rows - 1 ' poner nombre de cuentas en columna cuenta
            DoEvents
            grd.TextMatrix(i, grd.ColIndex("Descripción")) = ponerDescripcionFilaAF(grd.TextMatrix(i, 1))
'            grd.TextMatrix(i, grd.ColIndex("Costo Unitario")) = ponerCostoUnitarioFila(i)
        Next i
    Case "PRDIARIO"
        'sumar
        With grd
        
        InsertarColumnaCuenta
        For i = .FixedRows To .Rows - 1 ' poner nombre de cuentas en columna cuenta
            DoEvents
            .TextMatrix(i, .ColIndex("Cuenta")) = ponerPRCuentaFila(.TextMatrix(i, 1))
        Next i
        Sumar
        End With
    Case "INVENTARIOSERIES"
        InsertarColumnaDescSeries
        For i = grd.FixedRows To grd.Rows - 1 ' poner nombre de cuentas en columna cuenta
            DoEvents
            grd.TextMatrix(i, grd.ColIndex("Descripción")) = ponerDescripcionFila(grd.TextMatrix(i, 1))
        Next i

End Select
AjustarAutoSize grd, -1, -1
End Sub
Private Sub Sumar()
    With grd
        .subtotal flexSTSum, -1, .ColIndex("Debe"), , .BackColorFrozen, , True, " ", , True
        .subtotal flexSTSum, -1, .ColIndex("Haber"), , .BackColorFrozen, , True, " ", , True
        grd.TextMatrix(grd.Rows - 1, grd.ColIndex("Cuenta")) = "Suman"
        .Refresh
    End With
End Sub

Private Sub InsertarColumnaCuenta()
    Const pos = 2
    grd.Cols = grd.Cols + 1
    grd.ColPosition(grd.Cols - 1) = pos
    grd.TextMatrix(0, pos) = "Cuenta"
    grd.ColKey(pos) = grd.TextMatrix(0, pos)
End Sub
Private Sub InsertarColumnaDesc_y_Cost()
    Dim pos   As Integer
    pos = 2
    grd.Cols = grd.Cols + 1
    grd.ColPosition(grd.Cols - 1) = pos
    grd.TextMatrix(0, pos) = "Descripción"
    grd.ColKey(pos) = grd.TextMatrix(0, pos)
    pos = 5
    grd.Cols = grd.Cols + 1
    grd.ColPosition(grd.Cols - 1) = pos
    grd.TextMatrix(0, pos) = "Costo Unitario"
    grd.ColKey(pos) = grd.TextMatrix(0, pos)
End Sub


Private Function ponerCostoUnitarioFila(ByVal i As Long) As Currency
    If Val(grd.ValueMatrix(i, grd.ColIndex("Cantidad"))) <> 0 Then
        ponerCostoUnitarioFila = Val(grd.ValueMatrix(i, grd.ColIndex("Costo Total"))) / Val(grd.ValueMatrix(i, grd.ColIndex("Cantidad")))
    End If
End Function
Private Function ponerCostoTotal(i As Long) As Currency
    ponerCostoTotal = grd.ValueMatrix(i, grd.ColIndex("Costo Unitario")) * grd.ValueMatrix(i, grd.ColIndex("Cantidad"))
End Function


Private Function ponerCuentaFila(codcuenta As String) As String
    Dim ct As CtCuenta
    Set ct = gobjMain.EmpresaActual.RecuperaCTCuenta(codcuenta)
    If Not (ct Is Nothing) Then
        If ct.BandTotal = False Then
            ponerCuentaFila = ct.NombreCuenta
        Else
            ponerCuentaFila = MSG_ERR & " (Cuenta de Mayor)"
        End If
    Else
        ponerCuentaFila = MSG_ERR
    End If
    Set ct = Nothing
End Function

Private Function ponerDescripcionFila(coditem As String) As String
    Dim iv As IVinventario
    Set iv = gobjMain.EmpresaActual.RecuperaIVInventario(coditem)
    If Not (iv Is Nothing) Then
        ponerDescripcionFila = iv.Descripcion
    Else
        ponerDescripcionFila = MSG_ERR
    End If
    Set iv = Nothing
End Function


Private Sub VisualizarTexto(ByVal archi As String)
    Dim f As Integer, s As String, Separador As String, i As Integer
    Dim v As Variant
    ' dim   encontro As Boolean  no  esta el archivo ordenado
    On Error GoTo errtrap
    
    MensajeStatus "Está leyendo el archivo " & archi & " ...", vbHourglass
    grd.Rows = grd.FixedRows    'Limpia la grilla
    grd.Redraw = flexRDNone
    f = FreeFile                'Obtiene número disponible de archivo
    
    'Abre el archivo para lectura
    '*** Agregado Oliver 26/03/2004   agrege una opcion especial porque
    Select Case Me.tag                  ' para importar el archivo de ventas de los locutorios
        Case "VENTASLOCUTORIOS"         ' tienen el separador como ;
            Separador = ";"
        Case Else
            Separador = ","
    End Select
    
    'encontro = False
    
    Open archi For Input As #f
        Do Until EOF(f)
            Line Input #f, s
            s = vbTab & Replace(s, Separador, vbTab)      'Convierte ',' a TAB
            
            If Me.tag = "VENTASLOCUTORIOS" And gConfig.AbrirArchivoenFormaDiferencial Then
                v = Split(s, vbTab)
                'Debug.Print s
                If Not IsEmpty(v) Then
                       If UBound(v) >= 19 Then
                            If Val(v(19)) > Val(UltimoNumTransImportado) Then
                                'encontro = True   no esta ordenado el archivo
                                grd.AddItem s
                            End If
                       End If
                End If
            Else
                grd.AddItem s
            End If
        Loop
    Close #f
    RemueveSpace
    ' ordenar
    If grd.Rows > 1 Then
    Select Case Me.tag
        Case "PLANCUENTA"
            grd.Select 1, 1, 1, 1
        Case "PLANPRCUENTA"
            grd.Select 1, 1, 1, 1
        Case "PLANCUENTASC"
            grd.Select 1, 1, 1, 1
        Case "PLANCUENTAFE"
            grd.Select 1, 1, 1, 1
        Case "ITEM"
            grd.Select 1, 1, 1, 1
        Case "PCPROV"
            grd.Select 1, 1, 1, 1
        Case "PCCLI"
            grd.Select 1, 1, 1, 1
        Case "PORPAGAR"
            grd.Select 1, 1, 1, 4
        Case "PORCOBRAR"
            grd.Select 1, 1, 1, 4
        Case "DIARIO"
            grd.Select 1, 1, 1, 1
        Case "INVENTARIO"
            grd.Select 1, 1, 1, 2
        Case "VENTASLOCUTORIOS"
            'DEB0 ORDENAR SOLO POR LA COLUMNA DE NUMERO DE NOTA DE VENTA
            For i = 0 To 18
                grd.ColSort(i) = flexSortNone
            Next i
            grd.ColSort(19) = flexSortGenericAscending
            grd.Select 1, 1, 1, 19
        Case "AFITEM"
            grd.Select 1, 1, 1, 1
        Case "AFINVENTARIO"
            grd.Select 1, 1, 1, 2
        Case "AFINVENTARIOC"
            grd.Select 1, 1, 1, 2
        Case "PRDIARIO"
            grd.Select 1, 1, 1, 1
        Case "INVENTARIOSERIES"
            grd.Select 1, 1, 1, 2
        Case "PLANENFERME"
            grd.Select 1, 1, 1, 1
            
    End Select
    End If
    grd.Sort = flexSortUseColSort

' poner numero
    GNPoneNumFila grd, False
    
    If Me.tag = "VENTASLOCUTORIOS" Then
        SubTotalizar (19)
        'Totalizar
        'Almacena la ruta del archivo de importaciòn
        SaveSetting APPNAME, App.Title, "ArchivoLocutorio", lblArchivoLocutorio.Caption
    End If
    
    grd.Redraw = flexRDDirect
    AjustarAutoSize grd, -1, -1
    
    grd.SetFocus
    MensajeStatus
    Exit Sub
errtrap:
    grd.Redraw = flexRDDirect
    MensajeStatus
    DispErr
    Close       'Cierra todo
    grd.SetFocus
    Exit Sub
End Sub

Private Sub RemueveSpace()
    Dim i As Long, j As Long
    
    With grd
        .Redraw = flexRDNone
        For i = .FixedRows To .Rows - 1
            For j = .FixedCols To .Cols - 1
                .TextMatrix(i, j) = Trim$(.TextMatrix(i, j))
            Next j
        Next i
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub VisualizarExcel(ByVal archi As String)
    MsgBox "No se dispone de ésta función por el momento...", vbInformation
End Sub



' solo para editar existencia minima  y proveedores


Private Sub Grabar()
    Dim i As Long
    On Error GoTo errtrap
    
    
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna fila para importar.", vbInformation
        Exit Sub
    End If
        
    'Confirmación
    If MsgBox("Está seguro que desea grabar?", vbQuestion + vbYesNo) <> vbYes Then
        grd.SetFocus
        Exit Sub
    End If
    
    'Deshabilita los botónes y menus

    mbooEjecutando = True
    mbooCancelado = False
    
    With grd
        prg1.min = 0
        prg1.max = 1
        If .Rows > .FixedRows Then prg1.max = .Rows - 1
        For i = .FixedRows To .Rows - 1
            'Si es que se canceló el proceso
            If mbooCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                GoTo cancelado
            End If
            
            If grd.TextMatrix(i, .Cols - 1) <> MSG_OK Then
                'Si ocurre un error y no quiere seguir el usuario, sale del ciclo
                If Not GrabarFila(i) Then GoTo cancelado
            End If
            
        Next i
    End With
    
cancelado:
    MensajeStatus
    mbooEjecutando = False
    prg1.value = prg1.min
    Exit Sub
errtrap:
    MensajeStatus
    DispErr
    mbooEjecutando = False
    prg1.value = prg1.min
    Exit Sub
End Sub

Private Function GrabarFila(ByVal i As Long) As Boolean
    Dim iv As IVinventario, cod As String, msg As String
    GrabarFila = True
    prg1.value = i
With grd
    grd.TextMatrix(i, grd.Cols - 1) = MSG_PROC
    cod = .TextMatrix(i, .ColIndex("Código Item"))
    MensajeStatus i & " de " & .Rows - .FixedRows, vbHourglass
    DoEvents
    On Error GoTo errtrap
    'Recupera el objeto de Inventario
    Set iv = gobjMain.EmpresaActual.RecuperaIVInventario(cod)

    Select Case Me.tag
    Case "EXISTENCIA MINIMA"
        ' "^#|<Código Item|<Existencia Minima|<Proveedor"
        'iv.ExistenciaMinima = .ValueMatrix(i, .ColIndex("Existencia Minima"))
        'iv.CodProveedor = .TextMatrix(i, .ColIndex("Proveedor"))
        'iv.CodInventario = .TextMatrix(i, .ColIndex("NuevoCod"))
        iv.Precio(1) = .ValueMatrix(i, .ColIndex("precio1"))
        iv.Precio(2) = .ValueMatrix(i, .ColIndex("precio2"))
        iv.Precio(3) = .ValueMatrix(i, .ColIndex("precio3"))
        iv.Precio(4) = .ValueMatrix(i, .ColIndex("precio4"))
        iv.Precio(5) = .ValueMatrix(i, .ColIndex("precio5"))
    End Select

End With
    iv.Grabar
    grd.TextMatrix(i, grd.Cols - 1) = MSG_OK
    Exit Function
errtrap:
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    
    msg = "Ha ocurrido un error al tratar de importar la fila #" & i & "." & vbCr & _
          "Código : " & grd.TextMatrix(i, 1) & vbCr & _
          "Error : " & Err.Description & vbCr & vbCr & _
          "Desea continuar el proceso desde la siguiente fila?"
    If MsgBox(msg, vbYesNo + vbExclamation) = vbYes Then
        GrabarFila = True
    Else
        GrabarFila = False
    End If
    Exit Function
End Function




Private Sub ImportarVentasLocutorio()
    Dim i As Long, j As Long, gncomp As GNComprobante
    Dim numtrans  As String, msg As String
    Dim resp As E_MiMsgBox
    


    On Error GoTo errtrap
    ' verificar si estan todos los datos configurados
    With gConfig
        If Len(.MONEDA) = 0 Then
            MsgBox "Debe configurar un tipo de Modena", vbInformation
            Exit Sub
        End If
        
        If Len(.CodTrans) = 0 Then
            MsgBox "Debe configurar un tipo de transaccion", vbInformation
            Exit Sub
        End If
        
        If Len(.CodCli) = 0 Then
            MsgBox "Debe configurar un cliente predeterminado", vbInformation
            Exit Sub
        End If
        
        If Len(.Responsable) = 0 Then
            MsgBox "Debe configurar un responsable", vbInformation
            Exit Sub
        End If
        
        If Len(.FormaCobroPago) = 0 Then
            MsgBox "Debe configurar una forma de cobro predeterminada"
            Exit Sub
        End If
    End With
    
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna fila para importar.", vbInformation
        Exit Sub
    End If
    
    
    'Confirmación
    If MsgBox("Está seguro que desea comenzar el proceso de importación?", _
                vbYesNo + vbQuestion) <> vbYes Then Exit Sub
                
    mbooErrores = False   '***Angel. 22/Abril/2004
    mbooEjecutando = True
    MensajeStatus "Importando...", vbHourglass
    
'    grd.Enabled = False
    With grd
        prg1.min = .FixedRows - 1
        prg1.max = .Rows - 1
        prg1.value = prg1.min
        
        'antes de comenzar a importa pone en NumTrans esta en blanco
        numtrans = ""
        mbooCancelado = False
        For i = .FixedRows To .Rows - 1
            prg1.value = i
            DoEvents                'Para dar control a Windows
            
            'Si usuario aplastó 'Cancelar', sale del ciclo
            If mbooCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                GoTo cancelado
            End If

            'Si aún no está importado bien, importa la fila
            If grd.TextMatrix(i, .Cols - 1) <> MSG_OK Then
                
                'grabar el comprobante cuando es una fila de subtotal
                If grd.IsSubtotal(i) Then
                    numtrans = ""
                    If Not gncomp Is Nothing Then
                        'Prepara IvKardexRecargo
                        'Prepara x Cobrar
                        msg = MSG_OK
                        GrabarComprobante gncomp, i, msg
                        
                        'gncomp.Grabar False, False   'GRABO EL COMPROBANTE ANTERIOR
                        PonerMensaje j, i - 1, msg   'poner menjade de ok a todas las filas q pertenecen al mismo comprobante
                    Else
                        PonerMensaje j, i - 1, "Eligio no Sobreescribir"
                    End If
                    Set gncomp = Nothing             'LIBRERO LA MEMORIA PARA CREAR OTRO
                Else
                    If Len(numtrans) = 0 Then
                        numtrans = grd.TextMatrix(i, grd.ColIndex("#NOTAVENTA"))
                        ImportarTransSub gncomp, numtrans, resp
                        j = i
                    End If
                    If Not gncomp Is Nothing Then   ' solo puede importar el ivkardex si el objeto gncomp esta creado
                        ImportarIvKardex gncomp, i  ' porque si eligio no sobreescribir el objeto no se creo
                    End If
                End If
            End If
        Next i
        
'        ' graba el ultimo que no queda grabado
'         gncomp.Grabar False, False   'GRABO EL COMPROBANTE ANTERIOR
'         PonerMensaje j, i - 1, MSG_OK   'poner menjade de ok a todas las filas
    End With
    If Not mbooErrores Then '***Angel. 22/Abril/2004
        MsgBox "Proceso finalizado con éxito"
        gConfig.AbrirArchivoenFormaDiferencial = True
    Else
        MsgBox "Proceso finalizado. Errores presentados, revise la información"
        gConfig.AbrirArchivoenFormaDiferencial = False
    End If
    GuardaConfig
    RecuperaConfig
        
cancelado:
    If Not (gncomp Is Nothing) Then Set gncomp = Nothing
    MensajeStatus
    mbooEjecutando = False
    prg1.value = prg1.min
 '   grd.Enabled = True
 '   EnabledEncabezado True
    Exit Sub
errtrap:
    MensajeStatus
    DispErr
    mbooEjecutando = False
    prg1.value = prg1.min
    grd.Enabled = True
    EnabledEncabezado True
    Exit Sub
End Sub

Private Function ImportarTransSub( _
                ByRef gc As GNComprobante, _
                ByVal numt As String, _
                ByRef resp As E_MiMsgBox) As Boolean
    Dim s As String, Estado As Byte, numtrans  As Long
    Dim sql As String
    On Error GoTo errtrap
    
    
    
    
    numtrans = CLng(numt)
    'Recuperar la transacción en el destino
    Set gc = gobjMain.EmpresaActual.RecuperaGNComprobante(0, gConfig.CodTrans, numtrans)
    'Si existe en el destino,
    If Not (gc Is Nothing) Then
        If (resp = mmsgSi) Or (resp = mmsgNo) Then
            'Confirma si quiere sobre escribir lo existente
            s = "La transacción " & gConfig.CodTrans & numtrans & " ya existe en la base destino." & vbCr & vbCr & _
                "Desea sobreescribirla?"
            resp = frmMiMsgBox.MiMsgBox(s, gConfig.CodTrans & numtrans)
        End If
        
        Select Case resp
        Case mmsgNoTodo, mmsgNo
            Set gc = Nothing   ' librero de memoria porque ya abrio y puede grabar  algo, pero
            GoTo salida        ' eligio no sobreescribir, para que el proceso q graba no pueda hacerlo
        Case mmsgCancelar
            mbooCancelado = True
            GoTo salida
        End Select
        
        
    'Si no existe,
    Else
        'Crea como nueva
        Set gc = gobjMain.EmpresaActual.CreaGNComprobante(gConfig.CodTrans)
    End If
    
    gc.numtrans = numtrans          'Asigna el número de trans.
    gc.CodResponsable = gConfig.Responsable
    gc.CodMoneda = gConfig.MONEDA
    gc.Cotizacion(gConfig.MONEDA) = 1
    gc.CodClienteRef = gConfig.CodCli
    
    'Primero limpia
    gc.BorrarIVKardex
    gc.BorrarAFKardex
    gc.BorrarIVKardexRecargo
    gc.BorrarPCKardex
    gc.BorrarTSKardex
    
    
    ImportarTransSub = True
salida:
    Exit Function
errtrap:
    'Saca mensaje en columna de resultado
    Set gc = Nothing
    'grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    If MsgBox(Err.Description & vbCr & vbCr & _
                "Desea continuar con siguiente transacción?", _
                vbQuestion + vbYesNo) <> vbYes Then
        mbooCancelado = True
    End If
    GoTo salida
    mbooErrores = True '***Angel. 22/Abril/2004
End Function

Sub ImportarIvKardex(ByRef gc As GNComprobante, _
                        ByVal i As Long)
    Dim ivk As IVKardex, ix As Long, iv As IVinventario, IVA As Currency
    
'    s = "^#|<IdCabina|<Cabina|<#Marcado|<Destino|<Trafico|<Fecha|<Hora|>Duracion|^x1|>ValorMinuto|>ICE|>Neto|>IVA|>Total|^Modificacion|^x2|^x3|^x4|>#NotaVenta|^x5|<Operador|<#Turno"
    
    

    On Error GoTo errtrap
        Set iv = gc.GNTrans.Empresa.RecuperaIVInventario("-")
        IVA = iv.PorcentajeIVA
        
        ix = gc.AddIVKardex
        Set ivk = gc.IVKardex(ix)
        grd.TextMatrix(i, grd.Cols - 1) = MSG_PROC
        With grd
            ivk.CodBodega = gc.GNTrans.CodBodegaPre
            ivk.CodInventario = "-"
            ivk.Nota = Left(.TextMatrix(i, 2) & Space(4), 4) & _
                       Left(.TextMatrix(i, 3) & Space(20), 20) & _
                       Left(.TextMatrix(i, 4) & Space(19), 19) & _
                       Left(.TextMatrix(i, 5) & Space(19), 19) & _
                       Left(.TextMatrix(i, 6) & Space(10), 10) & _
                       Left(.TextMatrix(i, 7) & Space(8), 8)
            ivk.IVA = IVA
            ivk.cantidad = .ValueMatrix(i, 8) * -1
            If ivk.cantidad = 0 Then ivk.cantidad = -0.0001   ' asigno una cantidad super pequena para no perder los items que solo marco y salio
            ivk.PrecioTotal = .ValueMatrix(i, .ColIndex("Neto")) * -1
            ivk.PrecioRealTotal = ivk.PrecioTotal
        End With
    Exit Sub
errtrap:
    'Saca mensaje en columna de resultado
    Set ivk = Nothing
    grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    If MsgBox(Err.Description & vbCr & vbCr & _
                "Desea continuar con siguiente transacción?", _
                vbQuestion + vbYesNo) <> vbYes Then
        mbooCancelado = True
    End If
    mbooErrores = True  '***Angel. 22/Abril/2004
End Sub

Sub PonerMensaje(j As Long, i As Long, msg As String)
    Dim X As Integer
    For X = j To i
        grd.TextMatrix(X, grd.Cols - 1) = msg
    Next X
End Sub

Private Function GrabarComprobante(gc As GNComprobante, i As Long, ByRef msg As String) As Boolean
    Dim j As Long, ivkr As IVKardexRecargo, pck As PCKardex, ix As Long, v As Currency
    Dim pc As PCProvCli, CodUsuarioAnterior As String, codvd As String
    Dim VENDEDOR As FCVendedor
    
    On Error GoTo errtrap
    
    ''Para las Auditorias cambio el nombre de usuario y lod ejo como estaba
    CodUsuarioAnterior = gobjMain.UsuarioActual.codUsuario
    
    '*** Actualizando Datos de Cabecera
    gc.FechaTrans = grd.TextMatrix(i - 1, grd.ColIndex("Fecha"))
    gc.HoraTrans = grd.TextMatrix(i - 1, grd.ColIndex("Hora"))
    gc.Descripcion = "Ventas Turno " & grd.TextMatrix(i - 1, grd.ColIndex("#Turno"))
    gc.numDocRef = grd.TextMatrix(i - 1, grd.ColIndex("#NotaVenta"))
    
    codvd = grd.TextMatrix(i - 1, grd.ColIndex("x5")) 'modificado por la columna x5
    'Controla que no sea mas de 10 caracteres
    If Len(codvd) > 10 Then codvd = Mid$(codvd, 1, 10)
    
    Set VENDEDOR = gc.GNTrans.Empresa.RecuperaFCVendedor(codvd)
    If VENDEDOR Is Nothing Then
        Set VENDEDOR = gc.GNTrans.Empresa.CreaFCVendedor
        VENDEDOR.CodVendedor = codvd 'grd.TextMatrix(i - 1, grd.ColIndex("x5"))
        VENDEDOR.nombre = codvd 'grd.TextMatrix(i - 1, grd.ColIndex("x5"))
        VENDEDOR.Grabar
        Set VENDEDOR = Nothing
    End If
    
    'Cambio el usuario para q grabe la auditoria
    gobjMain.UsuarioActual.codUsuario = codvd 'VENDEDOR.CodVendedor
    gc.CodVendedor = codvd 'grd.TextMatrix(i - 1, grd.ColIndex("x5"))
    
    
    
    '''Recuperar Nombre del Cliente
    Set pc = gc.GNTrans.Empresa.RecuperaPCProvCli(gc.CodClienteRef)
    If Not pc Is Nothing Then
        gc.nombre = pc.nombre  'asigno el nombre a la transaccion porq no es automatico
    End If
    Set pc = Nothing
    
    
    '*** Crear los Recargos Descuentos con los totales de IVA &  ICE
    j = gc.AddIVKardexRecargo
    Set ivkr = gc.IVKardexRecargo(j)
    ivkr.codRecargo = "IVA"
    ivkr.Orden = 1
    ivkr.BandModificable = True
    ivkr.valor = grd.ValueMatrix(i, grd.ColIndex("IVA"))
    j = gc.AddIVKardexRecargo
    Set ivkr = Nothing
    Set ivkr = gc.IVKardexRecargo(j)
    ivkr.codRecargo = "ICE"
    ivkr.Orden = 2
    ivkr.BandModificable = True
    ivkr.valor = grd.ValueMatrix(i, grd.ColIndex("ICE"))
    Set ivkr = Nothing
    
    gc.ProrratearIVKardexRecargo
    
    
    '*** Crear Forma de Cobro
    v = Abs(gc.IVKardexTotal(True)) + gc.IVRecargoTotal(True, True)
    If v <> 0 Then
        ix = gc.AddPCKardex
        Set pck = gc.PCKardex(ix)
        pck.codforma = gConfig.FormaCobroPago               'Cobro al contado
        pck.Debe = v
        pck.FechaEmision = gc.FechaTrans
        pck.FechaVenci = pck.FechaEmision
        pck.CodProvCli = gc.CodClienteRef
        pck.Orden = gc.CountPCKardex
        Set pck = Nothing
    End If
    gc.GeneraAsiento
    gc.Grabar False, False
    UltimoNumTransImportado = grd.TextMatrix(i - 1, grd.ColIndex("#NotaVenta"))
    gobjMain.UsuarioActual.codUsuario = CodUsuarioAnterior
    GrabarComprobante = True
    Exit Function
errtrap:
    GrabarComprobante = False
    gobjMain.UsuarioActual.codUsuario = CodUsuarioAnterior
    msg = Err.Description
    mbooErrores = True '***Angel. 22/Abril/2004
End Function


Private Sub SubTotalizar(col As Long)
    Dim i As Long
    With grd
        For i = 1 To .Cols - 1
            If grd.ColData(i) = "SubTotal" Then
                .subtotal flexSTSum, col, i, "#,#0.0000", grd.GridColor, vbBlack, , "Subtotal", col, True
            End If
        Next i
    End With
End Sub

Private Sub Totalizar()
    Dim i As Long
    With grd
        For i = 1 To .Cols - 1
            If grd.ColData(i) = "SubTotal" Then
                .subtotal flexSTSum, -1, i, "#,#0.0000", .BackColorSel, vbYellow, vbBlack, "Total"
            End If
        Next i
    End With
End Sub



Private Sub GrabarUltimoNumImp()
    SaveSetting APPNAME, App.Title, "UltimoRegistroImportadoVentasLocutorios", UltimoNumTransImportado
End Sub


Private Sub RecuperarUltimoNumImp()
    UltimoNumTransImportado = GetSetting(APPNAME, App.Title, "UltimoRegistroImportadoVentasLocutorios", "")
End Sub

Private Function ImportarFilaAFItem(ByVal i As Long) As Boolean
    Dim msg As String, af As AFinventario
    On Error GoTo errtrap
    
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_PROC
    
    Set af = gobjMain.EmpresaActual.CreaAFInventario
    With af
        .CodInventario = grd.TextMatrix(i, grd.ColIndex("Código"))
        .CodAlterno1 = grd.TextMatrix(i, grd.ColIndex("CodAlterno1"))
        .CodAlterno2 = Val(grd.TextMatrix(i, grd.ColIndex("CodAlterno2")))
        .Descripcion = grd.TextMatrix(i, grd.ColIndex("Descripción"))
        .NumSerie = grd.TextMatrix(i, grd.ColIndex("Numero Serie"))
        .VidaUtil = grd.TextMatrix(i, grd.ColIndex("Vida Util"))
        .tipodepre = grd.TextMatrix(i, grd.ColIndex("Tipo Depre"))
        .DepAnterior = grd.TextMatrix(i, grd.ColIndex("Dep Anteriores"))
        .Marca = grd.TextMatrix(i, grd.ColIndex("Marca"))
        .CostoResidual = grd.ValueMatrix(i, grd.ColIndex("CostoResidual"))
        .CodUnidad = grd.TextMatrix(i, grd.ColIndex("Unidad medida"))
        .CodUnidadConteo = grd.TextMatrix(i, grd.ColIndex("Unidad medida"))
        .PorcentajeIVA = grd.ValueMatrix(i, grd.ColIndex("% IVA")) / 100
        If grd.ValueMatrix(i, grd.ColIndex("% IVA")) / 100 = 0 Then
            .bandIVA = False
        Else
            .bandIVA = True
        End If
        .CodMoneda = grd.TextMatrix(i, grd.ColIndex("Moneda"))
        .CodGrupo(1) = grd.TextMatrix(i, grd.ColIndex("Grupo1"))
        .CodGrupo(2) = grd.TextMatrix(i, grd.ColIndex("Grupo2"))
        .CodGrupo(3) = grd.TextMatrix(i, grd.ColIndex("Grupo3"))
        .CodGrupo(4) = grd.TextMatrix(i, grd.ColIndex("Grupo4"))
        .CodGrupo(5) = grd.TextMatrix(i, grd.ColIndex("Grupo5"))
        .CodCuentaActivo = grd.TextMatrix(i, grd.ColIndex("Cód.Cuenta Activo"))
        .CodCuentaCosto = grd.TextMatrix(i, grd.ColIndex("Cód.Cuenta Costo"))
        .CodCuentaVenta = grd.TextMatrix(i, grd.ColIndex("Cód.Cuenta Venta"))
        .CodCuentaDepreAcumulada = grd.TextMatrix(i, grd.ColIndex("Cód.Cuenta Depre Acumulada"))
        .CodCuentaDepreGasto = grd.TextMatrix(i, grd.ColIndex("Cód.Cuenta Depre Gasto"))
        .CodCuentaDepRevaloriza = grd.TextMatrix(i, grd.ColIndex("Cód.Cuenta Dep Revaloriza"))
        .CodCuentaRevaloriza = grd.TextMatrix(i, grd.ColIndex("Cód.Cuenta Revaloriza"))
        If Len(grd.TextMatrix(i, grd.ColIndex("Fecha Compra"))) > 0 Then
            .FechaCompra = grd.TextMatrix(i, grd.ColIndex("Fecha Compra"))
        End If
        If grd.TextMatrix(i, grd.ColIndex("Servicio(S/N)")) = "S" Then
            .BandServicio = True
        Else
            .BandServicio = False
        End If
        
        .Observacion = grd.TextMatrix(i, grd.ColIndex("Observación"))
        .Grabar
        
        'Saca mensaje en columna de resultado
        grd.TextMatrix(i, grd.Cols - 1) = MSG_OK
    End With
    
    ImportarFilaAFItem = True
    Exit Function
errtrap:
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    
    msg = "Ha ocurrido un error al tratar de importar la fila #" & i & "." & vbCr & _
          "Código : " & grd.TextMatrix(i, 1) & vbCr & _
          "Error : " & Err.Description & vbCr & vbCr & _
          "Desea continuar el proceso desde la siguiente fila?"
    If MsgBox(msg, vbYesNo + vbExclamation) = vbYes Then
        ImportarFilaAFItem = True
    Else
        ImportarFilaAFItem = False
    End If
    Exit Function
End Function


Private Function ImportarFilaAFInventario(ByVal i As Long, ByRef gncomp As GNComprobante) As Boolean
    Dim msg As String, AFinventario As AFKardex, ix As Long
    On Error GoTo errtrap
    ix = gncomp.AddAFKardex
    Set AFinventario = gncomp.AFKardex(ix)
    grd.TextMatrix(i, grd.Cols - 1) = MSG_PROC
    With AFinventario
        'AÑADIR DATOS A GRABAR EN afinventario
        .CodInventario = grd.TextMatrix(i, grd.ColIndex("Código Item"))
        .CodBodega = grd.TextMatrix(i, grd.ColIndex("Código Bodega"))
        .cantidad = grd.ValueMatrix(i, grd.ColIndex("Cantidad"))
        .CostoTotal = grd.ValueMatrix(i, grd.ColIndex("Costo Compra"))
        .CostoRealTotal = grd.ValueMatrix(i, grd.ColIndex("Costo Compra"))
        .Orden = ix
    End With
    grd.TextMatrix(i, grd.Cols - 1) = MSG_OK
    ImportarFilaAFInventario = True
    Exit Function
errtrap:
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    
    msg = "Ha ocurrido un error al tratar de importar la fila #" & i & "." & vbCr & _
          "Código : " & grd.TextMatrix(i, 1) & vbCr & _
          "Error : " & Err.Description & vbCr & vbCr & _
          "Desea continuar el proceso desde la siguiente fila?"
    If MsgBox(msg, vbYesNo + vbExclamation) = vbYes Then
        ImportarFilaAFInventario = True
    Else
        ImportarFilaAFInventario = False
    End If
    gncomp.RemoveAFKardex ix, AFinventario
    Exit Function
End Function


Private Function ponerDescripcionFilaAF(coditem As String) As String
    Dim iv As AFinventario
    Set iv = gobjMain.EmpresaActual.RecuperaAFInventario(coditem)
    If Not (iv Is Nothing) Then
        ponerDescripcionFilaAF = iv.Descripcion
    Else
        ponerDescripcionFilaAF = MSG_ERR
    End If
    Set iv = Nothing
End Function

Private Sub InsertarColumnaDesc()
    Dim pos   As Integer
    pos = 2
    grd.Cols = grd.Cols + 1
    grd.ColPosition(grd.Cols - 1) = pos
    grd.TextMatrix(0, pos) = "Descripción"
    grd.ColKey(pos) = grd.TextMatrix(0, pos)
End Sub

Private Function ImportarFilaPCGrupo(ByVal numGrupo As Byte, ByVal i As Long) As Boolean
    Dim msg As String, pcg As PCGRUPO
    On Error GoTo errtrap
    
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_PROC
    
    Set pcg = gobjMain.EmpresaActual.CreaPCGrupo(numGrupo)
    With pcg
        .CodGrupo = grd.TextMatrix(i, grd.ColIndex("Código"))
        .Descripcion = grd.TextMatrix(i, grd.ColIndex("Descripción"))
        .BandValida = True
        .Grabar
        
        'Saca mensaje en columna de resultado
        grd.TextMatrix(i, grd.Cols - 1) = MSG_OK
    End With
    
    ImportarFilaPCGrupo = True
    Exit Function
errtrap:
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    
    msg = "Ha ocurrido un error al tratar de importar la fila #" & i & "." & vbCr & _
          "Código : " & grd.TextMatrix(i, 1) & vbCr & _
          "Error : " & Err.Description & vbCr & vbCr & _
          "Desea continuar el proceso desde la siguiente fila?"
    If MsgBox(msg, vbYesNo + vbExclamation) = vbYes Then
        ImportarFilaPCGrupo = True
    Else
        ImportarFilaPCGrupo = False
    End If
    Exit Function
End Function

Private Function ImportarFilaIVGrupo(ByVal numGrupo As Byte, ByVal i As Long) As Boolean
    Dim msg As String, ivg As ivgrupo
    On Error GoTo errtrap
    
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_PROC
    
    Set ivg = gobjMain.EmpresaActual.CreaIVGrupo(numGrupo)
    With ivg
        .CodGrupo = grd.TextMatrix(i, grd.ColIndex("Código"))
        .Descripcion = grd.TextMatrix(i, grd.ColIndex("Descripción"))
        .BandValida = True
        .Grabar
        
        
        'Saca mensaje en columna de resultado
        grd.TextMatrix(i, grd.Cols - 1) = MSG_OK
    End With
    
    ImportarFilaIVGrupo = True
    Exit Function
errtrap:
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    
    msg = "Ha ocurrido un error al tratar de importar la fila #" & i & "." & vbCr & _
          "Código : " & grd.TextMatrix(i, 1) & vbCr & _
          "Error : " & Err.Description & vbCr & vbCr & _
          "Desea continuar el proceso desde la siguiente fila?"
    If MsgBox(msg, vbYesNo + vbExclamation) = vbYes Then
        ImportarFilaIVGrupo = True
    Else
        ImportarFilaIVGrupo = False
    End If
    Exit Function
End Function

Private Function ImportarFilaAFGrupo(ByVal numGrupo As Byte, ByVal i As Long) As Boolean
    Dim msg As String, afg As AFGrupo
    On Error GoTo errtrap
    
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_PROC
    
    Set afg = gobjMain.EmpresaActual.CreaAFGrupo(numGrupo)
    With afg
        .CodGrupo = grd.TextMatrix(i, grd.ColIndex("Código"))
        .Descripcion = grd.TextMatrix(i, grd.ColIndex("Descripción"))
        .BandValida = True
        .Grabar
        
        'Saca mensaje en columna de resultado
        grd.TextMatrix(i, grd.Cols - 1) = MSG_OK
    End With
    
    ImportarFilaAFGrupo = True
    Exit Function
errtrap:
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    
    msg = "Ha ocurrido un error al tratar de importar la fila #" & i & "." & vbCr & _
          "Código : " & grd.TextMatrix(i, 1) & vbCr & _
          "Error : " & Err.Description & vbCr & vbCr & _
          "Desea continuar el proceso desde la siguiente fila?"
    If MsgBox(msg, vbYesNo + vbExclamation) = vbYes Then
        ImportarFilaAFGrupo = True
    Else
        ImportarFilaAFGrupo = False
    End If
    Exit Function
End Function


Private Function ImportarFilaPRCuenta(ByVal i As Long) As Boolean
    Dim ct As prCuenta, msg As String
    Dim ct_Aux As prCuenta, nivelPadre As Integer
    On Error GoTo errtrap
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_PROC
    Set ct = gobjMain.EmpresaActual.CreaPRCuenta
    
    With ct
        .codcuenta = grd.TextMatrix(i, grd.ColIndex("Código"))
        .NombreCuenta = grd.TextMatrix(i, grd.ColIndex("Nombre de cuenta"))
        .TipoCuenta = Val(grd.TextMatrix(i, grd.ColIndex("Tipo")))
        .CodCuentaSuma = grd.TextMatrix(i, grd.ColIndex("Cód. Cuenta a sumar"))
        'jeaa 24/09/04 para modificar el campo la cuenta de total la cta cuentaSuma y obtener el nivel del padre
        If Len(.CodCuentaSuma) > 0 Then
            Set ct_Aux = gobjMain.EmpresaActual.RecuperaPRCuenta(.CodCuentaSuma)
            If Not ct_Aux Is Nothing Then
                ct_Aux.BandTotal = True
                .nivel = ct_Aux.nivel + 1
                ct_Aux.Grabar
                Set ct_Aux = Nothing
            End If
        End If
        .Grabar
        'Saca mensaje en columna de resultado
        grd.TextMatrix(i, grd.Cols - 1) = MSG_OK
    End With
    ImportarFilaPRCuenta = True
    Exit Function
errtrap:
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    msg = "Ha ocurrido un error al tratar de importar la fila #" & i & "." & vbCr & _
          "Código : " & grd.TextMatrix(i, 1) & vbCr & _
          "Error : " & Err.Description & vbCr & vbCr & _
          "Desea continuar el proceso desde la siguiente fila?"
    If MsgBox(msg, vbYesNo + vbExclamation) = vbYes Then
        ImportarFilaPRCuenta = True
    Else
        ImportarFilaPRCuenta = False
    End If
    Exit Function
End Function


Private Function ImportarFilaAFInventarioC(ByVal i As Long, ByRef gncomp As GNComprobante) As Boolean
    Dim msg As String, AFinventario As AFKardexCustodio, ix As Long
    On Error GoTo errtrap
    ix = gncomp.AddAFKardexCustodio
    Set AFinventario = gncomp.AFKardexCustodio(ix)
    grd.TextMatrix(i, grd.Cols - 1) = MSG_PROC
    With AFinventario
        'AÑADIR DATOS A GRABAR EN afinventario
        .CodInventario = grd.TextMatrix(i, grd.ColIndex("Código Item"))
        .CodEmpleado = grd.TextMatrix(i, grd.ColIndex("Código Empleado"))
        .cantidad = grd.ValueMatrix(i, grd.ColIndex("Cantidad"))
        .Orden = ix
    End With
    grd.TextMatrix(i, grd.Cols - 1) = MSG_OK
    ImportarFilaAFInventarioC = True
    Exit Function
errtrap:
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    
    msg = "Ha ocurrido un error al tratar de importar la fila #" & i & "." & vbCr & _
          "Código : " & grd.TextMatrix(i, 1) & vbCr & _
          "Error : " & Err.Description & vbCr & vbCr & _
          "Desea continuar el proceso desde la siguiente fila?"
    If MsgBox(msg, vbYesNo + vbExclamation) = vbYes Then
        ImportarFilaAFInventarioC = True
    Else
        ImportarFilaAFInventarioC = False
    End If
    gncomp.RemoveAFKardexCustodio ix, AFinventario
    Exit Function
End Function

Private Function ImportarFilaEmp(ByVal i As Long) As Boolean
    Dim msg As String, pc As PCProvCli, Cliprov As String, cad As String
    Dim Per As Personal
    On Error GoTo errtrap
    grd.TextMatrix(i, grd.Cols - 1) = MSG_PROC
    Set pc = gobjMain.EmpresaActual.RecuperaEmpleado(grd.TextMatrix(i, grd.ColIndex("Código")))
    If Not (pc Is Nothing) Then
        msg = "Ya existe un  Empleado con el codigo " & _
                pc.CodProvCli & " y nombre : " & pc.nombre & vbCr & vbCr & _
                "esta seguro que desea sobreescribirlos por datos de " & cad
        If MsgBox(msg, vbYesNo + vbExclamation) = vbNo Then
            grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
            ImportarFilaEmp = True
            Exit Function
        End If
    Else
        Set pc = gobjMain.EmpresaActual.CreaPCProvCli
    End If
    With pc
        .CodProvCli = grd.TextMatrix(i, grd.ColIndex("Código"))
        .nombre = grd.TextMatrix(i, grd.ColIndex("Nombre"))
        .Direccion1 = grd.TextMatrix(i, grd.ColIndex("Dirección"))
        .Telefono1 = grd.TextMatrix(i, grd.ColIndex("Teléfono"))
        .bandEmpleado = True
        
        .ruc = grd.TextMatrix(i, grd.ColIndex("CI"))
        If Len(.ruc) = 13 Then
            .codtipoDocumento = "R"
        Else
            .codtipoDocumento = "P"
        End If
        
        .CodGrupo1 = grd.TextMatrix(i, grd.ColIndex(gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupoE(1)))
        .CodGrupo2 = grd.TextMatrix(i, grd.ColIndex(gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupoE(2)))
        .CodGrupo3 = grd.TextMatrix(i, grd.ColIndex(gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupoE(3)))
        .CodGrupo4 = grd.TextMatrix(i, grd.ColIndex(gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupoE(4)))
        
        '.Grabar
            .GrabarEmpleado
        
        Set Per = gobjMain.EmpresaActual.RecuperarEmpleado(pc.IdProvCli)
            If Not Per Is Nothing Then
            Per.Salario = grd.TextMatrix(i, grd.ColIndex("Sueldo Basico"))
            Per.FechaIngreso = grd.TextMatrix(i, grd.ColIndex("Fecha Ingreso"))
            Per.BandActivo = True
            Per.Grabar pc.IdProvCli
            Set Per = Nothing
            End If
        'Saca mensaje en columna de resultado
        grd.TextMatrix(i, grd.Cols - 1) = MSG_OK
    End With
    ImportarFilaEmp = True
   Exit Function
errtrap:
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    msg = "Ha ocurrido un error al tratar de importar la fila #" & i & "." & vbCr & _
          "Error : " & Err.Description & vbCr & vbCr & _
          "Desea continuar el proceso desde la siguiente fila?"
    If MsgBox(msg, vbYesNo + vbExclamation) = vbYes Then
        ImportarFilaEmp = True
   Else
        ImportarFilaEmp = False
   End If

    Exit Function
End Function

Private Function ImportarFilaCuentaSC(ByVal i As Long) As Boolean
    Dim ct As ctCuentaSC, msg As String
    Dim ct_Aux As ctCuentaSC, nivelPadre As Integer
    On Error GoTo errtrap
    
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_PROC
    
    Set ct = gobjMain.EmpresaActual.CreaCTCuentaSC
    With ct
        .codcuenta = grd.TextMatrix(i, grd.ColIndex("Código"))
        .NombreCuenta = grd.TextMatrix(i, grd.ColIndex("Nombre de cuenta"))
        .TipoCuenta = Val(grd.TextMatrix(i, grd.ColIndex("Tipo")))
        .CodCuentaSuma = grd.TextMatrix(i, grd.ColIndex("Cód. Cuenta a sumar"))
        'jeaa 24/09/04 para modificar el campo la cuenta de total la cta cuentaSuma y obtener el nivel del padre
        If Len(.CodCuentaSuma) > 0 Then
            Set ct_Aux = gobjMain.EmpresaActual.RecuperaCTCuentaSC(.CodCuentaSuma)
            If Not ct_Aux Is Nothing Then
                ct_Aux.BandTotal = True
                .nivel = ct_Aux.nivel + 1
                ct_Aux.Grabar
                Set ct_Aux = Nothing
            End If
        End If
        .Grabar
        
        'Saca mensaje en columna de resultado
        grd.TextMatrix(i, grd.Cols - 1) = MSG_OK
    End With
    
    ImportarFilaCuentaSC = True
    Exit Function
errtrap:
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    
    msg = "Ha ocurrido un error al tratar de importar la fila #" & i & "." & vbCr & _
          "Código : " & grd.TextMatrix(i, 1) & vbCr & _
          "Error : " & Err.Description & vbCr & vbCr & _
          "Desea continuar el proceso desde la siguiente fila?"
    If MsgBox(msg, vbYesNo + vbExclamation) = vbYes Then
        ImportarFilaCuentaSC = True
    Else
        ImportarFilaCuentaSC = False
    End If
    Exit Function
End Function


Sub ImportarAvKardex(ByRef gc As GNComprobante, _
                        ByVal i As Long)
    Dim ivk As AFKardex, ix As Long, iv As AFinventario, IVA As Currency
    
'    s = "^#|<IdCabina|<Cabina|<#Marcado|<Destino|<Trafico|<Fecha|<Hora|>Duracion|^x1|>ValorMinuto|>ICE|>Neto|>IVA|>Total|^Modificacion|^x2|^x3|^x4|>#NotaVenta|^x5|<Operador|<#Turno"
    
    

    On Error GoTo errtrap
        Set iv = gc.GNTrans.Empresa.RecuperaAFInventario("-")
        IVA = iv.PorcentajeIVA
        
        ix = gc.AddAFKardex
        Set ivk = gc.AFKardex(ix)
        grd.TextMatrix(i, grd.Cols - 1) = MSG_PROC
        With grd
            ivk.CodBodega = gc.GNTrans.CodBodegaPre
            ivk.CodInventario = "-"
            ivk.Nota = Left(.TextMatrix(i, 2) & Space(4), 4) & _
                       Left(.TextMatrix(i, 3) & Space(20), 20) & _
                       Left(.TextMatrix(i, 4) & Space(19), 19) & _
                       Left(.TextMatrix(i, 5) & Space(19), 19) & _
                       Left(.TextMatrix(i, 6) & Space(10), 10) & _
                       Left(.TextMatrix(i, 7) & Space(8), 8)
            ivk.IVA = IVA
            ivk.cantidad = .ValueMatrix(i, 8) * -1
            If ivk.cantidad = 0 Then ivk.cantidad = -0.0001   ' asigno una cantidad super pequena para no perder los items que solo marco y salio
            ivk.PrecioTotal = .ValueMatrix(i, .ColIndex("Neto")) * -1
            ivk.PrecioRealTotal = ivk.PrecioTotal
        End With
    Exit Sub
errtrap:
    'Saca mensaje en columna de resultado
    Set ivk = Nothing
    grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    If MsgBox(Err.Description & vbCr & vbCr & _
                "Desea continuar con siguiente transacción?", _
                vbQuestion + vbYesNo) <> vbYes Then
        mbooCancelado = True
    End If
    mbooErrores = True  '***Angel. 22/Abril/2004
End Sub


Private Function ImportarFilaCuentaFE(ByVal i As Long) As Boolean
    Dim ct As ctCuentaFE, msg As String
    Dim ct_Aux As ctCuentaFE, nivelPadre As Integer
    On Error GoTo errtrap
    
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_PROC
    
    Set ct = gobjMain.EmpresaActual.CreaCTCuentaFE
    With ct
        .codcuenta = grd.TextMatrix(i, grd.ColIndex("Código"))
        .NombreCuenta = grd.TextMatrix(i, grd.ColIndex("Nombre de cuenta"))
        .TipoCuenta = Val(grd.TextMatrix(i, grd.ColIndex("Tipo")))
        .CodCuentaSuma = grd.TextMatrix(i, grd.ColIndex("Cód. Cuenta a sumar"))
        'jeaa 24/09/04 para modificar el campo la cuenta de total la cta cuentaSuma y obtener el nivel del padre
        If Len(.CodCuentaSuma) > 0 Then
            Set ct_Aux = gobjMain.EmpresaActual.RecuperaCTCuentaFE(.CodCuentaSuma)
            If Not ct_Aux Is Nothing Then
                ct_Aux.BandTotal = True
                .nivel = ct_Aux.nivel + 1
                ct_Aux.Grabar
                Set ct_Aux = Nothing
            End If
        End If
        .Grabar
        
        'Saca mensaje en columna de resultado
        grd.TextMatrix(i, grd.Cols - 1) = MSG_OK
    End With
    
    ImportarFilaCuentaFE = True
    Exit Function
errtrap:
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    
    msg = "Ha ocurrido un error al tratar de importar la fila #" & i & "." & vbCr & _
          "Código : " & grd.TextMatrix(i, 1) & vbCr & _
          "Error : " & Err.Description & vbCr & vbCr & _
          "Desea continuar el proceso desde la siguiente fila?"
    If MsgBox(msg, vbYesNo + vbExclamation) = vbYes Then
        ImportarFilaCuentaFE = True
    Else
        ImportarFilaCuentaFE = False
    End If
    Exit Function
End Function


Private Function ImportarFilaPRDiario(ByVal i As Long, ByRef gncomp As GNComprobante) As Boolean
    Dim msg As String, CTdiario As PRLibroDetalle, ix As Long
    On Error GoTo errtrap
    ix = gncomp.AddPRLibroDetalle
    Set CTdiario = gncomp.PRLibroDetalle(ix)
    grd.TextMatrix(i, grd.Cols - 1) = MSG_PROC
    With CTdiario
        'AÑADIR DATOS A GRABAR EN CTDIARIO

        .codcuenta = grd.TextMatrix(i, grd.ColIndex("Código Cuenta"))
        .Descripcion = txtDescripcion.Text
        .Debe = Val(grd.TextMatrix(i, grd.ColIndex("Debe")))
        .Haber = Val(grd.TextMatrix(i, grd.ColIndex("Haber")))
        .Orden = ix

'        Debug.Print i; ix
    End With
    grd.TextMatrix(i, grd.Cols - 1) = MSG_OK
    ImportarFilaPRDiario = True
    Exit Function
errtrap:
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    
    msg = "Ha ocurrido un error al tratar de importar la fila #" & i & "." & vbCr & _
          "Código : " & grd.TextMatrix(i, 1) & vbCr & _
          "Error : " & Err.Description & vbCr & vbCr & _
          "Desea continuar el proceso desde la siguiente fila?"
    If MsgBox(msg, vbYesNo + vbExclamation) = vbYes Then
        ImportarFilaPRDiario = True
    Else
        ImportarFilaPRDiario = False
    End If
    gncomp.RemovePRLibroDetalle ix, CTdiario
    Exit Function
End Function


Private Function ponerPRCuentaFila(codcuenta As String) As String
    Dim ct As prCuenta
    Set ct = gobjMain.EmpresaActual.RecuperaPRCuenta(codcuenta)
    If Not (ct Is Nothing) Then
        If ct.BandTotal = False Then
            ponerPRCuentaFila = ct.NombreCuenta
        Else
            ponerPRCuentaFila = MSG_ERR & " (Cuenta de Mayor)"
        End If
    Else
        ponerPRCuentaFila = MSG_ERR
    End If
    Set ct = Nothing
End Function

Private Function ImportarFilaPorCobrarPagarEmp(ByVal i As Long, ByRef gncomp As GNComprobante, ByVal bandCobrar As Boolean) As Boolean
    Dim msg As String, kardex As PCKardex, ix As Long
    On Error GoTo errtrap
    ix = gncomp.AddPCKardex
    Set kardex = gncomp.PCKardex(ix)
    With kardex
        .CodEmpleado = grd.TextMatrix(i, grd.ColIndex("Código Prov/Cli"))
        .NumLetra = grd.TextMatrix(i, grd.ColIndex("numdoc"))
        .FechaEmision = grd.TextMatrix(i, grd.ColIndex("Fecha"))
        .FechaVenci = grd.TextMatrix(i, grd.ColIndex("Fecha Vencimiento"))
        .Codelemento = fcbElemento.KeyText
        If bandCobrar Then
            .Debe = grd.TextMatrix(i, grd.ColIndex("Valor"))
        Else
            .Haber = grd.TextMatrix(i, grd.ColIndex("Valor"))
        End If
        .Observacion = grd.TextMatrix(i, grd.ColIndex("Observacion"))
        .codforma = fcbForma.KeyText
        .Orden = ix
        grd.TextMatrix(i, grd.Cols - 1) = MSG_OK
    End With
    ImportarFilaPorCobrarPagarEmp = True
    Exit Function
errtrap:
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    msg = "Ha ocurrido un error al tratar de importar la fila #" & i & "." & vbCr & _
          "Código : " & grd.TextMatrix(i, 1) & vbCr & _
          "Error : " & Err.Description & vbCr & vbCr & _
          "Desea continuar el proceso desde la siguiente fila?"
    If MsgBox(msg, vbYesNo + vbExclamation) = vbYes Then
        ImportarFilaPorCobrarPagarEmp = True
    Else
        ImportarFilaPorCobrarPagarEmp = False
    End If
    gncomp.RemovePCKardex ix, kardex
    Exit Function
End Function


Private Sub InsertarColumnaDescSeries()
    Dim pos   As Integer
    pos = 2
    grd.Cols = grd.Cols + 1
    grd.ColPosition(grd.Cols - 1) = pos
    grd.TextMatrix(0, pos) = "Descripción"
    grd.ColKey(pos) = grd.TextMatrix(0, pos)
    
End Sub

Private Function ImportarFilaInventarioSerie(ByVal i As Long, ByRef gncomp As GNComprobante) As Boolean
    Dim msg As String, ix As Long
    Dim idInve As Long, IdSerie As Long
    Dim ivS As IVNumSerie
    On Error GoTo errtrap
            idInve = 0
            IdSerie = 0
            idInve = gncomp.Empresa.RecuperaIdInventario(grd.TextMatrix(i, grd.ColIndex("Código Item")))
    
            Set ivS = gncomp.Empresa.CreaIVSerie
            ivS.campo1 = grd.TextMatrix(i, grd.ColIndex("Campo1"))
            ivS.campo2 = grd.TextMatrix(i, grd.ColIndex("Campo2"))
            ivS.campo3 = grd.TextMatrix(i, grd.ColIndex("Campo3"))
            ivS.Campo4 = grd.TextMatrix(i, grd.ColIndex("Campo4"))
            ivS.Campo5 = grd.TextMatrix(i, grd.ColIndex("Campo5"))
            ivS.idinventario = idInve

            ivS.FechaCreacion = gncomp.FechaTrans
            ivS.GrabarIVSerieNew IdSerie

            ix = gncomp.AddIVKNumSerie
            gncomp.IVKNumSerie(ix).IdSerie = IdSerie
            gncomp.IVKNumSerie(ix).cantidad = grd.ValueMatrix(i, grd.ColIndex("Cantidad"))
            gncomp.IVKNumSerie(ix).IdBodega = gncomp.Empresa.RecuperaIdBodega(grd.TextMatrix(i, grd.ColIndex("Código Bodega")))
            gncomp.IVKNumSerie(ix).Orden = ix

    

    grd.TextMatrix(i, grd.Cols - 1) = MSG_PROC

    grd.TextMatrix(i, grd.Cols - 1) = MSG_OK
    ImportarFilaInventarioSerie = True
    Exit Function
errtrap:
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    
    msg = "Ha ocurrido un error al tratar de importar la fila #" & i & "." & vbCr & _
          "Código : " & grd.TextMatrix(i, 1) & vbCr & _
          "Error : " & Err.Description & vbCr & vbCr & _
          "Desea continuar el proceso desde la siguiente fila?"
    If MsgBox(msg, vbYesNo + vbExclamation) = vbYes Then
        ImportarFilaInventarioSerie = True
    Else
        ImportarFilaInventarioSerie = False
    End If
    'gncomp.RemoveIVKardex ix, IVinventario
    Set ivS = Nothing
    Exit Function
End Function


Private Function ImportarFilaEnfermedades(ByVal i As Long) As Boolean
    Dim ct As PCListaEnfermedad, msg As String
    Dim ct_Aux As PCListaEnfermedad, nivelPadre As Integer
    On Error GoTo errtrap
    
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_PROC
    
    Set ct = gobjMain.EmpresaActual.CreaListaEnfermedad
    With ct
        .codcuenta = grd.TextMatrix(i, grd.ColIndex("Código"))
        .NombreCuenta = Mid$(grd.TextMatrix(i, grd.ColIndex("Nombre de cuenta")), 1, 79)
        .TipoCuenta = Val(grd.TextMatrix(i, grd.ColIndex("Tipo")))
        .CodCuentaSuma = grd.TextMatrix(i, grd.ColIndex("Cód. Cuenta a sumar"))
        'jeaa 24/09/04 para modificar el campo la cuenta de total la cta cuentaSuma y obtener el nivel del padre
        If Len(.CodCuentaSuma) > 0 Then
            Set ct_Aux = gobjMain.EmpresaActual.RecuperaListaEnfermedad(.CodCuentaSuma)
            If Not ct_Aux Is Nothing Then
                ct_Aux.BandTotal = True
                .nivel = ct_Aux.nivel + 1
                ct_Aux.Grabar
                Set ct_Aux = Nothing
            End If
        End If
        .Grabar
        
        'Saca mensaje en columna de resultado
        grd.TextMatrix(i, grd.Cols - 1) = MSG_OK
    End With
    
    ImportarFilaEnfermedades = True
    Exit Function
errtrap:
    'Saca mensaje en columna de resultado
    grd.TextMatrix(i, grd.Cols - 1) = MSG_ERR
    
    msg = "Ha ocurrido un error al tratar de importar la fila #" & i & "." & vbCr & _
          "Código : " & grd.TextMatrix(i, 1) & vbCr & _
          "Error : " & Err.Description & vbCr & vbCr & _
          "Desea continuar el proceso desde la siguiente fila?"
    If MsgBox(msg, vbYesNo + vbExclamation) = vbYes Then
        ImportarFilaEnfermedades = True
    Else
        ImportarFilaEnfermedades = False
    End If
    Exit Function
End Function


