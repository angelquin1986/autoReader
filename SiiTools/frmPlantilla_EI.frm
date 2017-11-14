VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPlantilla_EI 
   Caption         =   "Datos de la Plantilla"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   6000
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3075
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5424
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Opciones"
      TabPicture(0)   =   "frmPlantilla_EI.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "optTipoFecha(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "optTipoFecha(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkIgnorarDocAsignado"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkIgnorarContable"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkActualizaCat"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chkLimitarFechaHora"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Catálogos"
      TabPicture(1)   =   "frmPlantilla_EI.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdQuitarCat"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdMarcarCat"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "grdCat"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label6"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Transacciones"
      TabPicture(2)   =   "frmPlantilla_EI.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label7"
      Tab(2).Control(1)=   "Label8"
      Tab(2).Control(2)=   "grdTrans"
      Tab(2).Control(3)=   "cmdMarcarTrans"
      Tab(2).Control(4)=   "cmdQuitarTrans"
      Tab(2).Control(5)=   "fcbSucursal"
      Tab(2).ControlCount=   6
      Begin FlexComboProy.FlexCombo fcbSucursal 
         Height          =   375
         Left            =   -73980
         TabIndex        =   32
         Top             =   480
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   661
         ColWidth1       =   3400
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
      Begin VB.CommandButton cmdQuitarTrans 
         Caption         =   "Quitar Todos"
         Height          =   495
         Left            =   -70200
         TabIndex        =   18
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton cmdMarcarTrans 
         Caption         =   "Marcar Todos"
         Height          =   495
         Left            =   -70200
         TabIndex        =   17
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmdQuitarCat 
         Caption         =   "Quitar Todos"
         Height          =   495
         Left            =   -70200
         TabIndex        =   15
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton cmdMarcarCat 
         Caption         =   "Marcar Todos"
         Height          =   495
         Left            =   -70200
         TabIndex        =   14
         Top             =   840
         Width           =   855
      End
      Begin VB.Frame Frame2 
         Height          =   2175
         Left            =   3120
         TabIndex        =   29
         Top             =   360
         Width           =   2535
         Begin VB.ListBox lstBodega 
            Height          =   1185
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   12
            Top             =   600
            Width           =   2295
         End
         Begin VB.CheckBox chkFiltrarBodega 
            Caption         =   "Filtrar x Bodega"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.CheckBox chkLimitarFechaHora 
         Caption         =   "Limitar Rango Fecha/Hora"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   2535
      End
      Begin VB.CheckBox chkActualizaCat 
         Caption         =   "Actualizar solo catálogos"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CheckBox chkIgnorarContable 
         Caption         =   "Ignorar Aspecto Contable"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   2295
      End
      Begin VB.CheckBox chkIgnorarDocAsignado 
         Caption         =   "Ignorar Doc. Asignado"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   2175
      End
      Begin VB.OptionButton optTipoFecha 
         Caption         =   "Fecha Modificación"
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   6
         Top             =   1200
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton optTipoFecha 
         Caption         =   "Fecha Transacción"
         Height          =   195
         Index           =   1
         Left            =   960
         TabIndex        =   7
         Top             =   1440
         Width           =   1815
      End
      Begin VSFlex7LCtl.VSFlexGrid grdCat 
         Height          =   1695
         Left            =   -74880
         TabIndex        =   13
         Top             =   480
         Width           =   4560
         _cx             =   8043
         _cy             =   2990
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
         FocusRect       =   4
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPlantilla_EI.frx":0054
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   -1  'True
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   2
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
         WordWrap        =   -1  'True
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
      Begin VSFlex7LCtl.VSFlexGrid grdTrans 
         Height          =   1695
         Left            =   -74880
         TabIndex        =   16
         Top             =   960
         Width           =   4560
         _cx             =   8043
         _cy             =   2990
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
         FocusRect       =   4
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   5000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPlantilla_EI.frx":00C9
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   -1  'True
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   2
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
         WordWrap        =   -1  'True
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
      Begin VB.Label Label8 
         Caption         =   "Sucursal"
         Height          =   255
         Left            =   -74820
         TabIndex        =   33
         Top             =   540
         Width           =   795
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Mantenga presionada la tecla CTRL para seleccionar varias filas.  "
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -74880
         TabIndex        =   31
         Top             =   2760
         Width           =   4695
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Mantenga presionada la tecla CTRL para seleccionar varias filas.  "
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   -74880
         TabIndex        =   30
         Top             =   2280
         Width           =   4695
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Fecha"
         Height          =   195
         Left            =   360
         TabIndex        =   28
         Top             =   960
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3000
      TabIndex        =   23
      Top             =   5280
      Width           =   1200
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar - F9"
      Height          =   375
      Left            =   1680
      TabIndex        =   22
      Top             =   5280
      Width           =   1200
   End
   Begin VB.CheckBox chkValida 
      Caption         =   "Válida"
      Height          =   195
      Left            =   4920
      TabIndex        =   21
      Top             =   4920
      Width           =   975
   End
   Begin VB.CheckBox chkActualizarCostos 
      Caption         =   "Actualizar Costos"
      Height          =   195
      Left            =   3000
      TabIndex        =   20
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CheckBox chkGuardarResultado 
      Caption         =   "Permitir Guardar Resultados"
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Top             =   4920
      Width           =   2535
   End
   Begin VB.TextBox txtPrefijoNombre 
      Height          =   375
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox txtDesc 
      Height          =   375
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   2
      Top             =   840
      Width           =   4695
   End
   Begin VB.TextBox txtCodigo 
      Height          =   375
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.ComboBox cboTipo 
      Height          =   315
      ItemData        =   "frmPlantilla_EI.frx":013E
      Left            =   1200
      List            =   "frmPlantilla_EI.frx":0148
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prefijo Archivo"
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   1320
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   960
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Plantilla"
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   240
      Width           =   900
   End
End
Attribute VB_Name = "frmPlantilla_EI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbooIniciando As Boolean
Private mobjPlantilla As clsPlantilla
Private mBandAceptado As Boolean
Const EXPOR = 0
Const IMPOR = 1

Public Function Inicio(ByVal tag As String, obj As clsPlantilla) As Boolean
    mbooIniciando = True
    Me.tag = tag
    Set mobjPlantilla = obj
    CargarBodega
    CargarCatalogo
    'CargarTrans
    CargarTransxSucursal
    fcbSucursal.SetData gobjMain.EmpresaActual.ListaGNSucursales(True, False) 'jeaa 10/09/2008
    Select Case Me.tag
    Case "Agregar"  'modificado por Diego 04/03/2004
        cboTipo.ListIndex = mobjPlantilla.Tipo
    Case "Modificar"
        CargarDatos
    Case "Copiar"
        CargarDatos
        Me.Caption = Me.Caption & " (Copiado)"
    End Select
    mbooIniciando = False
    HabilitarControles
    Me.Show vbModal
    Inicio = mBandAceptado
    Unload Me
End Function

Private Sub CargarDatos()
    With mobjPlantilla
        cboTipo.ListIndex = .Tipo
        txtCodigo.Text = .CodPlantilla
        txtDesc.Text = .Descripcion
        txtPrefijoNombre.Text = .PrefijoNombreArchivo
        chkLimitarFechaHora.value = IIf(.BandRangoFechaHora, vbChecked, vbUnchecked)
        If .BandTipoFecha Then
            optTipoFecha(0).value = True
            optTipoFecha(1).value = False
        Else
            optTipoFecha(0).value = False
            optTipoFecha(1).value = True
        End If
        chkActualizaCat.value = IIf(.BandActualizaCatalogos, vbChecked, vbUnchecked)
        chkIgnorarContable.value = IIf(.BandIgnorarContabilidad, vbChecked, vbUnchecked)
        chkIgnorarDocAsignado.value = IIf(.BandIgnorarDocAsignado, vbChecked, vbUnchecked)
        chkFiltrarBodega.value = IIf(.BandFiltrarxBodega, vbChecked, vbUnchecked)
        chkGuardarResultado.value = IIf(.BandGuardarResultado, vbChecked, vbUnchecked)
        chkActualizarCostos.value = IIf(.BandActualizarCosto, vbChecked, vbUnchecked)
        chkValida.value = IIf(.BandValida, vbChecked, vbUnchecked)
        
        'Recupera Bodegas Seleccionadas
        VisualizarSeleccion "BODEGA", .ListaBodegas, 0
        'Recupera Catalogos Seleccionados
        VisualizarSeleccion "CATALOGO", .ListaCatalogos, grdCat.ColIndex("Tabla")
        'Recupera Trans. Seleccionadas
        VisualizarSeleccion "TRANS", .ListaTransacciones, grdTrans.ColIndex("Código")
    End With
End Sub

Private Sub cboTipo_Click()
    On Error GoTo ErrTrap
    If Not mbooIniciando Then
        mobjPlantilla.Tipo = cboTipo.ListIndex
        HabilitarControles
    End If
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub chkActualizaCat_Click()
    On Error GoTo ErrTrap
    If Not mbooIniciando Then mobjPlantilla.BandActualizaCatalogos = _
                              (chkActualizaCat.value = vbChecked)
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub chkActualizarCostos_Click()
    On Error GoTo ErrTrap
    If Not mbooIniciando Then mobjPlantilla.BandActualizarCosto = _
                              (chkActualizarCostos.value = vbChecked)
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub chkFiltrarBodega_Click()
    On Error GoTo ErrTrap
    If Not mbooIniciando Then
        mobjPlantilla.BandFiltrarxBodega = _
                              (chkFiltrarBodega.value = vbChecked)
        lstBodega.Enabled = (chkFiltrarBodega.value = vbChecked)
    End If
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub chkGuardarResultado_Click()
    On Error GoTo ErrTrap
    If Not mbooIniciando Then mobjPlantilla.BandGuardarResultado = _
                              (chkGuardarResultado.value = vbChecked)
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub chkIgnorarContable_Click()
    On Error GoTo ErrTrap
    If Not mbooIniciando Then mobjPlantilla.BandIgnorarContabilidad = _
                              (chkIgnorarContable.value = vbChecked)
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub chkIgnorarDocAsignado_Click()
    On Error GoTo ErrTrap
    If Not mbooIniciando Then mobjPlantilla.BandIgnorarDocAsignado = _
                              (chkIgnorarDocAsignado.value = vbChecked)
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub chkLimitarFechaHora_Click()
    On Error GoTo ErrTrap
    If Not mbooIniciando Then mobjPlantilla.BandRangoFechaHora = _
                              (chkLimitarFechaHora.value = vbChecked)
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub chkValida_Click()
    On Error GoTo ErrTrap
    If Not mbooIniciando Then mobjPlantilla.BandValida = _
                              (chkValida.value = vbChecked)
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub cmdCancelar_Click()
    mBandAceptado = False
    If mobjPlantilla.Modificado Then
        Unload Me
    Else
        Me.Hide
    End If
End Sub

Private Sub cmdGrabar_Click()
    If Grabar Then
        mBandAceptado = True
        Me.Hide
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdMarcarCat_Click()
    MarcarDesmarcar "CATALOGO", True
End Sub

Private Sub cmdMarcarTrans_Click()
    MarcarDesmarcar "TRANS", True
End Sub

Private Sub cmdQuitarCat_Click()
    MarcarDesmarcar "CATALOGO", False
End Sub

Private Sub cmdQuitarTrans_Click()
    MarcarDesmarcar "TRANS", False
End Sub

Private Sub fcbSucursal_Selected(ByVal Text As String, ByVal KeyText As String)
    CargarTransxSucursal
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF9
        cmdGrabar_Click
        KeyCode = 0
    Case Else
        MoverCampo Me, KeyCode, Shift, True
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ImpideSonidoEnter Me, KeyAscii
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim rt As Integer
    If Not (mobjPlantilla Is Nothing) Then
        If mobjPlantilla.Modificado Then
            Me.ZOrder
            rt = MsgBox(MSG_CANCELMOD, vbYesNoCancel + vbQuestion)
            Select Case rt
            Case vbYes           'Graba y cierra
                If Grabar Then
                    Me.Hide
                Else
                    Cancel = 1    'Si ocurre error al grabar,no cierra
                End If
            Case vbNo          'Cierra sin grabar
                Me.Hide
            Case vbCancel
                Cancel = 1      'No se cierra la ventana
            End Select
        End If
    End If
End Sub

Private Sub Form_Terminate()
    Set mobjPlantilla = Nothing
End Sub

Private Sub optTipoFecha_Click(Index As Integer)
    On Error GoTo ErrTrap
    If Not mbooIniciando Then
        Select Case Index
        Case 0
            mobjPlantilla.BandTipoFecha = optTipoFecha(Index).value
        Case 1
            mobjPlantilla.BandTipoFecha = Not (optTipoFecha(Index).value)
        End Select
    End If
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub

End Sub

Private Sub txtCodigo_Change()
    On Error GoTo ErrTrap
    If Not mbooIniciando Then mobjPlantilla.CodPlantilla = txtCodigo.Text
    Exit Sub
ErrTrap:
    DispErr
    txtCodigo.Text = mobjPlantilla.CodPlantilla
    Exit Sub
End Sub

Private Sub txtDesc_Change()
    On Error GoTo ErrTrap
    If Not mbooIniciando Then mobjPlantilla.Descripcion = txtDesc.Text
    Exit Sub
ErrTrap:
    DispErr
    txtDesc.Text = mobjPlantilla.Descripcion
End Sub

Private Sub txtPrefijoNombre_Change()
    On Error GoTo ErrTrap
    If Not mbooIniciando Then mobjPlantilla.PrefijoNombreArchivo = txtPrefijoNombre.Text
    Exit Sub
ErrTrap:
    DispErr
    txtPrefijoNombre.Text = mobjPlantilla.PrefijoNombreArchivo
End Sub

Private Sub CargarTrans()
    With grdTrans
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .FormatString = "^|<Código|<Descripción"
        .LoadArray gobjMain.EmpresaActual.ListaGNTrans("", False, False)
        AsignarTituloAColKey grdTrans
        GNPoneNumFila grdTrans, False
        AjustarAutoSize grdTrans, -1, -1
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub CargarCatalogo()
    CargarCatalogos grdCat
End Sub

Private Sub CargarBodega()
    Dim rs As Recordset
    
    Set rs = gobjMain.EmpresaActual.ListaIVBodega(True, True)
    With rs
        If Not (.BOF And .EOF) Then
            Do Until .EOF
                lstBodega.AddItem !CodBodega
                .MoveNext
            Loop
        End If
    End With
    Set rs = Nothing
End Sub

Private Function Grabar() As Boolean
    On Error GoTo mensaje
    
    Grabar = False
    'Pasa Datos al objeto
    mobjPlantilla.ListaTransacciones = FlexCodigosSeleccionados(grdTrans, grdTrans.ColIndex("Código"), True)
    mobjPlantilla.ListaCatalogos = FlexCodigosSeleccionados(grdCat, grdCat.ColIndex("Tabla"), True)
    mobjPlantilla.ListaBodegas = BodegasSeleccionadas
    mobjPlantilla.Grabar
    Grabar = True
    Exit Function
    
mensaje:
    MsgBox Err.Description, vbOKOnly + vbExclamation
    Grabar = False
    Exit Function
End Function

Private Sub VisualizarSeleccion(ByVal op As String, _
                                ByVal cad As String, _
                                Columna As Long)
    Dim i As Long, v As Variant, j As Long, cod As String
    
    v = Split(cad, ",")
    If UBound(v, 1) < 0 Then Exit Sub
    Select Case op
    Case "TRANS"
        With grdTrans
            For i = .FixedRows To .Rows - 1
                For j = LBound(v, 1) To UBound(v, 1)
                    cod = Trim$(v(j))                   'Quita espacios del extremo
                    cod = Right$(cod, Len(cod) - 1)     'Quita primer "'"
                    cod = Left$(cod, Len(cod) - 1)      'Quita ultimo "'"
                    If Trim$(.TextMatrix(i, Columna)) = Trim$(cod) Then
                        .IsSelected(i) = True
                        Exit For
                    End If
                Next j
            Next i
        End With
    Case "CATALOGO"
        With grdCat
            For i = .FixedRows To .Rows - 1
                For j = LBound(v, 1) To UBound(v, 1)
                    cod = Trim$(v(j))                   'Quita espacios del extremo
                    cod = Right$(cod, Len(cod) - 1)     'Quita primer "'"
                    cod = Left$(cod, Len(cod) - 1)      'Quita ultimo "'"
                    If Trim$(.TextMatrix(i, Columna)) = Trim$(cod) Then
                        .IsSelected(i) = True
                        Exit For
                    End If
                Next j
            Next i
        End With
    Case "BODEGA"
        With lstBodega
            For i = 0 To .ListCount - 1
                For j = LBound(v, 1) To UBound(v, 1)
                    cod = Trim$(v(j))                   'Quita espacios del extremo
                    cod = Right$(cod, Len(cod) - 1)     'Quita primer "'"
                    cod = Left$(cod, Len(cod) - 1)      'Quita ultimo "'"
                    If Trim$(.List(i)) = Trim$(cod) Then
                        .Selected(i) = True
                        Exit For
                    End If
                Next j
            Next i
        End With
    End Select
End Sub

Public Function BodegasSeleccionadas() As String
    Dim i As Long, s As String
    
    With lstBodega
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                If Len(s) > 0 Then s = s & ","
                s = s & "'" & .List(i) & "'"
            End If
        Next i
    End With
    BodegasSeleccionadas = s
End Function

Private Sub MarcarDesmarcar(ByVal op As String, ByVal Bandera As Boolean)
    Dim i As Integer
    Select Case op
    Case "CATALOGO"
        With grdCat
            For i = .FixedRows To .Rows - 1
                .IsSelected(i) = Bandera
            Next i
        End With
    Case "TRANS"
        With grdTrans
            For i = .FixedRows To .Rows - 1
                .IsSelected(i) = Bandera
            Next i
        End With
    End Select
End Sub

Private Sub HabilitarControles()
    Dim i As Integer, band As Boolean
       
    Select Case cboTipo.ListIndex
    Case EXPOR
        chkIgnorarDocAsignado.value = vbUnchecked
        chkActualizaCat.value = vbUnchecked
        chkGuardarResultado.value = vbUnchecked
        chkActualizarCostos.value = vbUnchecked
        band = True
    Case IMPOR
        chkLimitarFechaHora.value = vbUnchecked
        chkFiltrarBodega.value = vbUnchecked
        'Limpia Bodegas seleccionadas
        For i = 0 To lstBodega.ListCount - 1
            lstBodega.Selected(i) = False
        Next i
        'Limpia Catalogos selecccionados
        MarcarDesmarcar "CATALOGO", False
        
        'Limpia Transacciones seleccionadas
        MarcarDesmarcar "TRANS", False
        
        band = False
    End Select
    If cboTipo.ListIndex < 0 Then Exit Sub
    'Controles no disponibles para Exportar
    chkIgnorarDocAsignado.Enabled = Not band
    chkActualizaCat.Enabled = Not band
    chkGuardarResultado.Enabled = Not band
    chkActualizarCostos.Enabled = Not band
    
    'Controles no disponibles para Importar
    chkLimitarFechaHora.Enabled = band
    optTipoFecha(0).Enabled = band
    optTipoFecha(1).Enabled = band
    chkFiltrarBodega.Enabled = band
    lstBodega.Enabled = (chkFiltrarBodega.value = vbChecked)
    grdCat.Enabled = band
    cmdMarcarCat.Enabled = band
    cmdQuitarCat.Enabled = band
    grdTrans.Enabled = band
    cmdMarcarTrans.Enabled = band
    cmdQuitarTrans.Enabled = band
    SSTab1.TabEnabled(1) = band
    SSTab1.TabEnabled(2) = band
End Sub

Private Sub CargarTransxSucursal()
    With grdTrans
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .FormatString = "^|<Código|<Descripción|<Sucursal"
        .LoadArray gobjMain.EmpresaActual.ListaGNTransxSucursal("", False, False, fcbSucursal.KeyText)
        AsignarTituloAColKey grdTrans
        GNPoneNumFila grdTrans, False
        AjustarAutoSize grdTrans, -1, -1
        .Redraw = flexRDBuffered
    End With
End Sub

