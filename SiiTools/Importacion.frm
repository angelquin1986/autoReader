VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmImportacionDatos 
   Caption         =   "Transacciones para importar"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   HasDC           =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   8055
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1935
      ScaleWidth      =   8055
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3090
      Width           =   8055
      Begin VB.Frame Frame1 
         Caption         =   "&Búsqueda"
         Height          =   1410
         Left            =   15
         TabIndex        =   6
         Top             =   -30
         Width           =   7995
         Begin VB.Frame fraFecha 
            Caption         =   "&Fecha (desde - hasta)"
            Height          =   1092
            Left            =   105
            TabIndex        =   16
            Top             =   165
            Width           =   1932
            Begin MSComCtl2.DTPicker dtpFecha1 
               Height          =   300
               Left            =   120
               TabIndex        =   17
               Top             =   240
               Width           =   1692
               _ExtentX        =   2990
               _ExtentY        =   529
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
               Format          =   106692609
               CurrentDate     =   36348
            End
            Begin MSComCtl2.DTPicker dtpFecha2 
               Height          =   300
               Left            =   120
               TabIndex        =   18
               Top             =   600
               Width           =   1692
               _ExtentX        =   2990
               _ExtentY        =   529
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
               Format          =   106692609
               CurrentDate     =   36348
            End
         End
         Begin VB.Frame fraEstado 
            Caption         =   "Estado"
            Height          =   1095
            Left            =   4845
            TabIndex        =   12
            Top             =   165
            Width           =   1710
            Begin VB.CheckBox chkEstado 
               Caption         =   "&No aprobados"
               Height          =   192
               Index           =   0
               Left            =   105
               TabIndex        =   15
               Top             =   225
               Value           =   1  'Checked
               Width           =   1410
            End
            Begin VB.CheckBox chkEstado 
               Caption         =   "&Aprobados"
               Height          =   192
               Index           =   1
               Left            =   105
               TabIndex        =   14
               Top             =   480
               Value           =   1  'Checked
               Width           =   1332
            End
            Begin VB.CheckBox chkEstado 
               Caption         =   "&Despachados"
               Height          =   192
               Index           =   2
               Left            =   105
               TabIndex        =   13
               Top             =   735
               Width           =   1452
            End
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "Buscar -F5"
            Height          =   600
            Left            =   6615
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   480
            Width           =   1305
         End
         Begin FlexComboProy.FlexCombo fcbNombre 
            Height          =   345
            Left            =   2325
            TabIndex        =   8
            Top             =   375
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   609
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
         Begin FlexComboProy.FlexCombo fcbCentro 
            Height          =   330
            Left            =   2325
            TabIndex        =   10
            Top             =   930
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   582
            ColWidth0       =   1800
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Centro de costo  "
            Height          =   195
            Left            =   2340
            TabIndex        =   11
            Top             =   705
            Width           =   1200
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "&Nombre  "
            Height          =   195
            Left            =   2355
            TabIndex        =   9
            Top             =   150
            Width           =   945
         End
      End
      Begin VB.CheckBox chkIncremental 
         Caption         =   "&Importación incremental (Items/Asiento)"
         Height          =   312
         Left            =   150
         TabIndex        =   1
         Top             =   1485
         Width           =   3732
      End
      Begin VB.PictureBox pic2 
         BorderStyle     =   0  'None
         Height          =   390
         Left            =   4395
         ScaleHeight     =   390
         ScaleWidth      =   3540
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1470
         Width           =   3540
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "&Cancelar"
            Height          =   375
            Left            =   2220
            TabIndex        =   3
            Top             =   0
            Width           =   1305
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "&Aceptar -F9"
            Default         =   -1  'True
            Height          =   390
            Left            =   450
            TabIndex        =   2
            Top             =   0
            Width           =   1305
         End
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grd 
      Align           =   1  'Align Top
      Height          =   2910
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      _cx             =   14208
      _cy             =   5133
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
      Rows            =   1
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
      AutoSearch      =   2
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   0
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
      PicturesOver    =   0   'False
      FillStyle       =   1
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
End
Attribute VB_Name = "frmImportacionDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Aceptado As Boolean
Private mCondEstado As Integer      'Para conservar el estado original de la condición
Private mGNComp As GNComprobante
Private mPrimero As Boolean

Public Function Inicio( _
                    ByVal gc As GNComprobante, _
                    ByRef Incremental As Boolean, _
                    ByRef TransIDs As String) As Boolean
    Dim id As Long, i As Long
    On Error GoTo ErrTrap
    
    Aceptado = False
    Me.Caption = "Importación de datos para " & _
                 gc.CodTrans & " (" & gc.GNTrans.Descripcion & ")"
    Set mGNComp = gc
    
    'Ajusta el tamaño de la ventan
    Me.Width = Screen.Width * 0.93
    Me.Height = Screen.Height * 0.7
    
    'Carga lista de centro de costo
    CargarCentro
    'Carga lista de Clientes y Proveedores
    CargaProvCli
        
    With gobjMain.objCondicion
        dtpFecha1.value = .fecha1
        dtpFecha2.value = .fecha2
        fcbNombre.KeyText = ""
    End With
'    'Deshabilita el CheckBox de Desaprobado si la trans. no permite importar de Desaprobados
'    chkEstado(ESTADO_NOAPROBADO).Enabled = Not gc.GNTrans.ImportaSoloAprobado
                
    'Visualiza la condicion inicial
    If Not mPrimero Then
        chkEstado(ESTADO_NOAPROBADO).value = _
                IIf(gc.GNTrans.ImportaSoloAprobado, vbUnchecked, vbChecked)
        chkEstado(ESTADO_APROBADO).value = vbChecked
        chkEstado(ESTADO_DESPACHADO).value = vbUnchecked
        chkIncremental.value = IIf(Incremental, vbChecked, vbUnchecked) '*** MAKOTO 15/dic/00
        
        mPrimero = True
    End If
    
    chkEstado_Click 0 'carga lo que esta en pantalla par el obj. que busque bien
    
    'Carga condición de estado y obtiene listado de transacciones
    gobjMain.objCondicion.TopN = 30
    CargarDatos          'Aquí se carga datos iniciales
    gobjMain.objCondicion.TopN = 0    'para la proxima ya pueda buscar todos
     Me.Show vbModal, frmMain
    
    If Aceptado Then
        'Asigna TransID de fuente al objeto
        id = grd.ValueMatrix(grd.Row, 1)
        gc.IdTransFuente = id
        Incremental = (chkIncremental.value = vbChecked)
    
        TransIDs = ""
        With grd
            For i = 0 To .SelectedRows - 1
                If Len(TransIDs) > 0 Then TransIDs = TransIDs & ", "
                TransIDs = TransIDs & .ValueMatrix(.SelectedRow(i), .ColIndex("TID"))
            Next i
        End With
    End If
    
    Inicio = Aceptado
'    Unload Me
    Me.Hide                         '*** MAKOTO 15/dic/00 Para que quede las condiciones
    Set mGNComp = Nothing
    'Recupera la condición de estado
    gobjMain.objCondicion.Estado = mCondEstado
    Exit Function
ErrTrap:
    DispErr
'    Unload Me
    Me.Hide                         '*** MAKOTO 15/dic/00 Para que quede las condiciones
    Set mGNComp = Nothing
    'Recupera la condición de estado
    gobjMain.objCondicion.Estado = mCondEstado
    Exit Function
End Function

Private Sub CargarCentro()
    Dim s As String
    s = fcbCentro.KeyText
    fcbCentro.SetData mGNComp.Empresa.ListaGNCentroCosto2(False, False)
    If Len(s) > 0 Then
        fcbCentro.KeyText = s
'    ElseIf Len(mGNComp.CodCentro) > 0 Then
'        fcbCentro.KeyText = mGNComp.CodCentro
    End If
End Sub

Private Sub CargaProvCli()
    fcbNombre.SetData gobjMain.EmpresaActual.ListaPCProvCli(True, True, False)
End Sub

Private Sub CargarDatos()
    On Error GoTo ErrTrap
    MensajeStatus MSG_PREPARA, vbHourglass
    
    With gobjMain.objCondicion
        .fecha1 = dtpFecha1.value
        .fecha2 = dtpFecha2.value
        .CodPC1 = fcbNombre.KeyText
    End With
    'Obtiene listado de transacciones de fuente
    Set grd.DataSource = mGNComp.ListaTransFuente2(fcbCentro.KeyText)
    GNPoneNumFila grd, False
    ConfigCols
    
    MensajeStatus
    If Me.Visible Then grd.SetFocus
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    If Me.Visible Then grd.SetFocus
    Exit Sub
End Sub

Private Sub ConfigCols()
    With grd
        .FormatString = "^#|<TID|<Fecha|<Cod.Trans|<#Trans" & _
                        "|<Nombre|<Descripcion|<Cod.CC|<Desc.CC|^Estado"
        AsignarTituloAColKey grd            '*** MAKOTO 15/dic/00
        AjustarAutoSize grd, -1, -1, 5000   '*** MAKOTO 15/dic/00
        .ColHidden(1) = True
        
        If .Rows > .FixedRows Then
            .Row = .FixedRows
            .col = 2
        End If
    End With
End Sub


Private Sub cmdAceptar_Click()
    If grd.Rows <= grd.FixedRows Then Exit Sub
    If grd.Row < grd.FixedRows Then             '*** MAKOTO 23/oct/00
        MsgBox "Seleccione un comprobante, por favor.", vbInformation
        Exit Sub
    End If
    Aceptado = True
    Me.Hide
End Sub

Private Sub cmdBuscar_Click()
    CargarDatos
End Sub

Private Sub cmdCancelar_Click()
    Aceptado = False
    Me.Hide
End Sub

Private Sub chkEstado_Click(Index As Integer)
    'No hay como desactivar todo
    If chkEstado(0).value = vbUnchecked And _
        chkEstado(1).value = vbUnchecked And _
        chkEstado(2).value = vbUnchecked Then
'        MsgBox "No se puede desactivar todas las opciones.", vbInformation
        chkEstado(Index).value = vbChecked
        Exit Sub
    End If

    With gobjMain.objCondicion
        .EstadoBool(ESTADO_NOAPROBADO) = (chkEstado(ESTADO_NOAPROBADO).value = vbChecked)
        .EstadoBool(ESTADO_APROBADO) = (chkEstado(ESTADO_APROBADO).value = vbChecked)
        .EstadoBool(ESTADO_DESPACHADO) = (chkEstado(ESTADO_DESPACHADO).value = vbChecked)
        .EstadoBool(ESTADO_ANULADO) = False
    End With
   
   ' CargarDatos
End Sub


''*** MAKOTO 15/dic/00 Agregado
'Private Sub fcbCentro_Selected(ByVal Text As String, ByVal KeyText As String)
'   CargarDatos
'End Sub

Private Sub Form_Activate()
    grd.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF5
        cmdBuscar_Click
        KeyCode = 0
    Case vbKeyF9
        cmdAceptar_Click
        KeyCode = 0
    Case Else
        MoverCampo Me, KeyCode, Shift, True
    End Select
End Sub

Private Sub Form_Load()
    'Ajusta el tamaño de la ventana a la pantalla
    Me.Width = Screen.Width * 0.85
    Me.Height = Screen.Height * 0.6
    
    'Guarda la condición de estado para luego recuperar
    mCondEstado = gobjMain.objCondicion.Estado
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    grd.Height = Me.ScaleHeight - pic1.Height
    'pic2.Left = (Me.ScaleWidth - pic2.Width) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Libera la referencia al objeto
    Set mGNComp = Nothing
End Sub


Private Sub grd_DblClick()
    cmdAceptar_Click
End Sub

Private Sub grd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdAceptar_Click
        KeyAscii = 0
    End If
End Sub
