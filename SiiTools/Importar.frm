VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmImportar 
   Caption         =   "Importación"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6885
   ScaleWidth      =   11535
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Selecione la Plantilla a utiilizar"
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   8175
      Begin VB.ComboBox cboPlantilla 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton cmdPlantilla 
         Caption         =   "..."
         Height          =   315
         Left            =   7680
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         Height          =   195
         Left            =   3240
         TabIndex        =   16
         Top             =   360
         Width           =   840
      End
      Begin VB.Label lblDescripcion 
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   4200
         TabIndex        =   15
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.CommandButton cmdGuardarRes 
      Caption         =   "&Guardar Res."
      Height          =   380
      Left            =   5070
      TabIndex        =   7
      Top             =   5640
      Width           =   1452
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "&Actualizar costo en trans. relacionadas"
      Height          =   380
      Left            =   2040
      TabIndex        =   6
      Top             =   5640
      Width           =   2970
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar... -F5"
      Height          =   380
      Left            =   6360
      TabIndex        =   3
      Top             =   1200
      Width           =   1452
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   372
      Left            =   6600
      TabIndex        =   8
      Top             =   5640
      Width           =   1092
   End
   Begin MSComDlg.CommonDialog dlg1 
      Left            =   5760
      Top             =   1080
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "mdb"
      DialogTitle     =   "Orígen de Importación"
   End
   Begin VB.CommandButton cmdImportar 
      Caption         =   "&Importar -F9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   480
      TabIndex        =   5
      Top             =   5640
      Width           =   1452
   End
   Begin VSFlex7LCtl.VSFlexGrid grdMsg 
      Align           =   2  'Align Bottom
      Height          =   1455
      Left            =   0
      TabIndex        =   9
      Top             =   5430
      Width           =   11535
      _cx             =   20346
      _cy             =   2566
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
      FocusRect       =   2
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Importar.frx":0000
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
   Begin TabDlg.SSTab sst1 
      Height          =   3855
      Left            =   240
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1680
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   6800
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   529
      TabCaption(0)   =   "Transacciones"
      TabPicture(0)   =   "Importar.frx":006D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdTrans"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Catálogos"
      TabPicture(1)   =   "Importar.frx":0089
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdCat"
      Tab(1).ControlCount=   1
      Begin VSFlex7LCtl.VSFlexGrid grdTrans 
         Height          =   3315
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   9675
         _cx             =   17066
         _cy             =   5847
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
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   5
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7LCtl.VSFlexGrid grdCat 
         Height          =   3315
         Left            =   -74880
         TabIndex        =   13
         Top             =   480
         Width           =   9675
         _cx             =   17066
         _cy             =   5847
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
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   3
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.CommandButton cmdExplorar 
      Caption         =   "..."
      Height          =   310
      Left            =   7800
      TabIndex        =   11
      Top             =   855
      Width           =   372
   End
   Begin VB.TextBox txtOrigen 
      Height          =   320
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   6975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Orígen  "
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   555
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuGrabarGrilla 
         Caption         =   "Grabar registro"
      End
   End
End
Attribute VB_Name = "frmImportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const HEIGHT_MIN = 7425   '6344
Const WIDTH_MIN = 8655   '7764

Private mcnOrigen As ADODB.Connection
Private mCancelado As Boolean
Private mEjecutando As Boolean
Private mBuscado As Boolean

Private mcnPlantilla As ADODB.Connection '***Angel. 27/feb/2004
Private mPlantilla As clsPlantilla       '***Angel. 27/feb/2004
Private mUltimaPlantilla As String       '***Angel. 27/feb/2004
Private mRutaBDDestino As String         '***Angel. 27/feb/2004
Private Const NUMERROR_DUPLI = -2147217873  '***Angel. 17/dic/2003

Public Sub Inicio()
    Me.Show
    Me.ZOrder
    mBuscado = False
End Sub

Private Sub cboPlantilla_Click()
    mBuscado = False
    Habilitar False
    RecuperarPlantilla
End Sub

Private Sub cmdActualizar_Click()
    If grdTrans.Rows = grdTrans.FixedRows Then Exit Sub 'Si no hay nada, no hace nada
    frmReasignacionCosto.Inicio grdTrans
End Sub

Private Sub cmdBuscar_Click()
    Dim sql As String, rs As Recordset, v As Variant
    On Error GoTo ErrTrap
    
    If Not AbrirOrigen Then Exit Sub
    grdTrans.Rows = grdTrans.FixedRows      'Para limpiar la selección
    
    
    'Obtiene lista de transacciones en el orígen        '*** MAKOTO 06/mar/01 Aumentado 'Nombre'
    sql = "SELECT FechaTrans, CodTrans, NumTrans, Nombre, Descripcion, " & _
                 "CodCentro, Estado " & _
          "FROM GNComprobante ORDER BY FechaTrans, HoraTrans " '**** Oliver, recupera en orden de fechatrans +  HoraTrans
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
        v = MiGetRows(rs)
        
        With grdTrans
            .Redraw = flexRDNone
            .LoadArray v            'Carga a la grilla
        
            .FormatString = "^#|<Fecha|<CodTrans|<NumTrans|<Nombre|<Descripción|<Cod.C.C.|^Estado"
            GNPoneNumFila grdTrans, False
            AsignarTituloAColKey grdTrans           'Para usar ColIndex
            AjustarAutoSize grdTrans, -1, -1, 3000  'Ajusta automáticamente ancho de cols.
            If .ColWidth(.ColIndex("Nombre")) > 1400 Then .ColWidth(.ColIndex("Nombre")) = 1400
            
            'Tipo de datos
            .ColDataType(.ColIndex("Fecha")) = flexDTDate
            .ColDataType(.ColIndex("CodTrans")) = flexDTString
            .ColDataType(.ColIndex("NumTrans")) = flexDTLong
            .ColDataType(.ColIndex("Descripción")) = flexDTString
            .ColDataType(.ColIndex("Cod.C.C.")) = flexDTString
            .ColDataType(.ColIndex("Estado")) = flexDTShort
            
            .Redraw = flexRDDirect
        End With
        
    Else
        'Si no hay nada de resultado limpia la grilla
        grdTrans.Rows = grdTrans.FixedRows
    End If
    rs.Close
    cmdImportar.SetFocus
    mBuscado = True             '*** MAKOTO 14/mar/01 Agregado
salida:
    MensajeStatus
    Set rs = Nothing
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    GoTo salida
End Sub

Private Sub cmdCancelar_Click()
    If mEjecutando Then
        mCancelado = True
    Else
        Unload Me
    End If
End Sub

Private Sub cmdExplorar_Click()
    On Error GoTo ErrTrap
    
    With dlg1
        If Len(.filename) = 0 Then
            '.InitDir = App.Path
            .InitDir = txtOrigen.Text
            '.FileName = gobjMain.EmpresaActual.CodEmpresa & _
                        Format(Date, "dd-mm-yyyy") & ".mdb"
            .filename = mPlantilla.BDDestino
        Else
            .InitDir = .filename
        End If
        .flags = cdlOFNPathMustExist + cdlOFNFileMustExist
        '.Filter = "Base de datos Jet (*.mdb)|Predefinido (" & Trim$(mPlantilla.PrefijoNombreArchivo) & "*.mdb)" & "|*.mdb|Todos (*.*)|*.*"
        .Filter = "Base de datos Jet (*.mdb)|*.mdb|Predefinido (" & _
                  Trim$(mPlantilla.PrefijoNombreArchivo) & "*.mdb)|" & _
                  Trim$(mPlantilla.PrefijoNombreArchivo) & "*.mdb" & _
                  "|Todos (*.*)|*.*"
        .ShowOpen
        txtOrigen.Text = .filename
    End With
    
    Exit Sub
ErrTrap:
    If Err.Number <> 32755 Then
        DispErr
    End If
    Exit Sub
End Sub

Private Function AbrirOrigen() As Boolean
    Dim s As String
    On Error GoTo ErrTrap

    AbrirOrigen = True
    s = Trim$(txtOrigen.Text)

    If mcnOrigen Is Nothing Then Set mcnOrigen = New ADODB.Connection
    If mcnOrigen.State <> adStateClosed Then mcnOrigen.Close
    
    'Abre la conección con el archivo de destino
    s = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & s & ";" & _
        "Jet OLEDB:Database Password='aq9021'" & _
        ";" & "Persist Security Info=False"
    mcnOrigen.Open s, "admin", ""
    
    Exit Function
ErrTrap:
    AbrirOrigen = False
    DispErr
    Exit Function
End Function

Private Sub cmdImportar_Click()
    Dim s As String
    On Error GoTo ErrTrap
    
    'Verifica si está especificado el origen
    s = Trim$(txtOrigen.Text)
    If Len(s) = 0 Then
        MsgBox "Debe especificar el archivo de orígen.", vbInformation
        txtOrigen.SetFocus
        Exit Sub
    End If
    
    'Si aun no está hecho la búsqueda, llamarlo automaticamente
    If Not mBuscado Then
        cmdBuscar_Click
    End If
    
    'Limpia la selcción
    LimpiarSeleccion grdCat
    LimpiarSeleccion grdTrans
    
    mCancelado = False
    Habilitar False
    
    If Not AbrirOrigen Then GoTo salida
    
    'Importa Catálogos
    ImportarCatalogo
    
    'Importa Transacciones
     If Not mPlantilla.BandActualizaCatalogos Then ImportarTrans
    
    'Guarda parámetros en el registro de sistema
    GuardarConfig
    'se activa boton guardar resultados 26/01/04
    If mPlantilla.BandGuardarResultado Then cmdGuardarRes.Enabled = True
    
salida:
    Habilitar True
    Exit Sub
ErrTrap:
    DispErr
    GoTo salida
End Sub

Private Sub Habilitar(ByVal v As Boolean)
    mEjecutando = Not v
    cmdBuscar.Enabled = v
    cmdExplorar.Enabled = v
    cmdImportar.Enabled = v
    txtOrigen.Enabled = v
        
    frmMain.mnuFile.Enabled = v
    frmMain.mnuHerramienta.Enabled = v
    frmMain.mnuTransferir.Enabled = v
    frmMain.mnuCerrarTodas.Enabled = v
End Sub

Private Sub ImportarCatalogo()
    Dim i As Long
    
    sst1.Tab = 1
    
    With grdCat
        For i = .FixedRows To .Rows - 1
            DoEvents
            
            'Si el usuario canceló la operación
            If mCancelado Then Exit For
            
            If Not ImportarCatalogoSub(.TextMatrix(i, .ColIndex("Catálogo")), _
                                .TextMatrix(i, .ColIndex("Tabla"))) Then
                'Si ocurrió algún error y no quizo continuar
                Exit For
            Else
                'Sí exportó sin problema, resalta la fila
                .IsSelected(i) = True
            End If
        Next i
    End With
    MensajeStatus
End Sub

Private Function ImportarCatalogoSub( _
                ByVal Desc As String, _
                ByVal tabla As String) As Boolean
    Dim NumReg As Long
    On Error GoTo ErrTrap
                
    Select Case tabla
    'Case "GNResponsable"
    Case "GNResp"
        NumReg = GrabarGNResponsable
    Case "CTCuenta"
        If Not mPlantilla.BandIgnorarContabilidad Then     '*** MAKOTO 14/mar/01 Agregado
            NumReg = GrabarCTCuenta
        End If
    Case "GNCentroCosto"
        NumReg = GrabarGNCentroCosto
    Case "TSBanco"
        NumReg = GrabarTSBanco
    Case "TSRetencion"                  '*** MAKOTO 12/feb/01 Agregado
        NumReg = GrabarTSRetencion
    Case "IVBodega"
        NumReg = GrabarIVBodega
    'Case "IVGrupo1", "IVGrupo2", "IVGrupo3", "IVGrupo4", "IVGrupo5"
    Case "IVG1", "IVG2", "IVG3", "IVG4", "IVG5"
        NumReg = GrabarIVGrupo(Val(Right$(tabla, 1)))
    'Case "IVInventario"
    Case "IVInv"
        NumReg = GrabarIVInventario
    Case "FCVendedor"
        NumReg = GrabarFCVendedor
    'Case "PCGrupo1", "PCGrupo2", "PCGrupo3", "PCGrupo4"
    Case "PCG1", "PCG2", "PCG3", "PCG4"
        NumReg = GrabarPCGrupo(Val(Right$(tabla, 1)))
    Case "DiasCred"
        NumReg = GrabarPCDiasCredito
    Case "PCProvCli(P)", "PCProvCli(C)"
        NumReg = GrabarPCProvCli(Right$(tabla, 3) = "(P)")
    Case "PCProvCli(G)"
        NumReg = GrabarPCGarante(True)
    Case "DescIVGPCG"  '**** jeaa 05/01/05
        NumReg = GrabarDesctoIVGrupoPCGrupo
    Case "TSFormaC_P"
        NumReg = GrabarTSFormaCobroPago
    Case "Motivo"
        NumReg = GrabarMotivo
    Case "TCompra"
        NumReg = GrabarTipoCompra
    Case "IVU" 'jeaa 17/04/2006
        NumReg = GrabarIVUnidad
    Case "Exist" 'jeaa 17/04/2006
        NumReg = GrabarIVExist
    Case "DescNumPagIVG"  '**** jeaa 12/09/2008
        NumReg = GrabarDesctoNumPagosIVGrupo
    Case "PCHistorial" 'jeaa 17/04/2006
        NumReg = GrabarPCHistorial
    Case "IVBanco" 'jeaa 17/04/2006
        NumReg = GrabarIVBanco
    Case "IVTarjeta" 'jeaa 17/04/2006
        NumReg = GrabarIVTarjeta
    Case "PLAIVGPCG"  '**** jeaa 05/01/05
        NumReg = GrabarPlazoIVGrupoPCGrupo
    Case "PCParroquia" 'jeaa 17/04/2006
        NumReg = GrabarPCParroquia
    
    End Select
    
    'Si fue cancelado devuelve numreg en negativo
    If NumReg < 0 Then
        DispMsg "Importar datos de " & Desc, "Cancelado", Abs(NumReg) & " registros."
    Else
        DispMsg "Importar datos de " & Desc, "OK", NumReg & " registros."
    End If
    ImportarCatalogoSub = True
salida:
    Exit Function
ErrTrap:
    DispMsg "Importar datos de " & Desc, "Error", Err.Description
    If MsgBox(Err.Description & vbCr & vbCr & _
                "Desea continuar con siguiente catálogo?", _
                vbQuestion + vbYesNo) = vbYes Then
        ImportarCatalogoSub = True
    End If
    GoTo salida
End Function

 
Private Function GrabarCTCuenta() As Long
    Dim sql As String, rs As Recordset, ct As CtCuenta, i As Long
    Dim s As String, resp As E_MiMsgBox
    
    On Error GoTo ErrTrap
    
    resp = mmsgSi
    
    'Abre el orígen
    sql = "SELECT * FROM CTCuenta ORDER BY CodCuenta"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    
    With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Plan de cuenta... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
        
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            
            'Primero busca si existe ya el mismo código
            Set ct = gobjMain.EmpresaActual.RecuperaCTCuenta(.Fields("CodCuenta"))
            If ct Is Nothing Then
                'Si no existe el código, crea nuevo
                Set ct = gobjMain.EmpresaActual.CreaCTCuenta
            Else
                'Si no se ha hecho pregunta, ó no ha contestado para Todo
                If (resp = mmsgSi) Or (resp = mmsgNo) Then
                    'Pregunta si quiere sobreescribir o no
                    s = "El registro '" & ct.codcuenta & "' (" & ct.NombreCuenta & ") ya existe en el destino." & vbCr & vbCr & _
                        "Desea sobreescribirlo?"
                    resp = frmMiMsgBox.MiMsgBox(s, "Plan de cuenta")
                End If
                Select Case resp
                Case mmsgCancelar
                    mCancelado = True
                    Exit Do
                Case mmsgNo, mmsgNoTodo
                    GoTo siguiente
                End Select
            End If
            
            ct.codcuenta = .Fields("CodCuenta")
            ct.NombreCuenta = .Fields("NombreCuenta")
            ct.nivel = .Fields("Nivel")
            
            'CodCuentaSuma --> IdCuentaSuma (Hace internamente en el objeto)
            ct.CodCuentaSuma = .Fields("CodCuentaSuma")
                                            
            ct.TipoCuenta = .Fields("TipoCuenta")
'            ct.BandDeudor = .Fields("BandDeudor")
            ct.BandTotal = .Fields("BandTotal")
            ct.BandValida = .Fields("BandValida")

            ct.Grabar
            
            Dim tabla As String, cod As String, campo As String
            campo = "CodCuenta"
            cod = .Fields("CodCuenta")
            tabla = "CTCuenta"
            CambiaFechaenTabla campo, cod, tabla
            
            GrabarCTCuenta = GrabarCTCuenta + 1
siguiente:
            .MoveNext
        Loop
        
        .Close
    End With
    

salida:
    Set rs = Nothing
    Set ct = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarCTCuenta = GrabarCTCuenta * -1
    Exit Function

ErrTrap:
    If Not (ct Is Nothing) Then
        s = Err.Description & ": " & Err.Source & vbCr & ct.codcuenta & ", " & ct.CodCuentaSuma
    End If
    DispMsg "Importar datos de Plan de cuenta", "Error", s
    If MsgBox(s & vbCr & vbCr & _
                "Desea continuar con el siguiente registro?", _
                vbQuestion + vbYesNo) = vbYes Then
        Resume siguiente
    Else
        mCancelado = True
    End If
    GoTo salida
End Function

Private Function GrabarGNResponsable() As Long
    Dim sql As String, rs As Recordset, cc As GNResponsable, i As Long
    Dim s As String, resp As E_MiMsgBox
    
    On Error GoTo ErrTrap
    
    resp = mmsgSi
    
    'Abre el orígen
    sql = "SELECT * FROM GNResponsable ORDER BY CodResponsable"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    
    With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Responsables ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
        
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            
            'Primero busca si existe ya el mismo código
            Set cc = gobjMain.EmpresaActual.RecuperaGNResponsable(.Fields("CodResponsable"))
            If cc Is Nothing Then
                'Si no existe el código, crea nuevo
                Set cc = gobjMain.EmpresaActual.CreaGNResponsable
            Else
                'Si no se ha hecho pregunta, ó no ha contestado para Todo
                If (resp = mmsgSi) Or (resp = mmsgNo) Then
                    'Pregunta si quiere sobreescribir o no
                    s = "El registro '" & cc.CodResponsable & "' (" & cc.nombre & ") ya existe en el destino." & vbCr & vbCr & _
                        "Desea sobreescribirlo?"
                    resp = frmMiMsgBox.MiMsgBox(s, "Responsable")
                End If
                Select Case resp
                Case mmsgCancelar
                    mCancelado = True
                    Exit Do
                Case mmsgNo, mmsgNoTodo
                    GoTo siguiente
                End Select
            End If
            
            cc.CodResponsable = .Fields("CodResponsable")
            cc.nombre = .Fields("Nombre")
            cc.BandValida = .Fields("BandValida")
            cc.Grabar
            GrabarGNResponsable = GrabarGNResponsable + 1
siguiente:
            .MoveNext
        Loop
        
        .Close
    End With
    

salida:
    Set rs = Nothing
    Set cc = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarGNResponsable = GrabarGNResponsable * -1
    Exit Function

ErrTrap:
    If Not (cc Is Nothing) Then
        s = Err.Description & ": " & Err.Source & vbCr & cc.CodResponsable & ", " & cc.CodResponsable
    End If
    DispMsg "Importar datos de Responsable", "Error", s
    If MsgBox(s & vbCr & vbCr & _
                "Desea continuar con el siguiente registro?", _
                vbQuestion + vbYesNo) = vbYes Then
        Resume siguiente
    Else
        mCancelado = True
    End If
    GoTo salida
End Function

Private Function GrabarGNCentroCosto() As Long
    Dim sql As String, rs As Recordset, cc As GNCentroCosto, i As Long
    Dim s As String, resp As E_MiMsgBox
    
    On Error GoTo ErrTrap
    
    resp = mmsgSi
    
    'Abre el orígen
    sql = "SELECT * FROM GNCentroCosto ORDER BY CodCentro"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    
    With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Centro de costo ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
        
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            
            'Primero busca si existe ya el mismo código
            Set cc = gobjMain.EmpresaActual.RecuperaGNCentroCosto(.Fields("CodCentro"))
            If cc Is Nothing Then
                'Si no existe el código, crea nuevo
                Set cc = gobjMain.EmpresaActual.CreaGNCentroCosto
            Else
                'Si no se ha hecho pregunta, ó no ha contestado para Todo
                If (resp = mmsgSi) Or (resp = mmsgNo) Then
                    'Pregunta si quiere sobreescribir o no
                    s = "El registro '" & cc.CodCentro & "' (" & cc.Descripcion & ") ya existe en el destino." & vbCr & vbCr & _
                        "Desea sobreescribirlo?"
                    resp = frmMiMsgBox.MiMsgBox(s, "Centro de costo")
                End If
                Select Case resp
                Case mmsgCancelar
                    mCancelado = True
                    Exit Do
                Case mmsgNo, mmsgNoTodo
                    GoTo siguiente
                End Select
            End If
            
            cc.CodCentro = .Fields("CodCentro")
            If Not IsNull(.Fields("Descripcion")) Then cc.Descripcion = .Fields("Descripcion")
            If Not IsNull(.Fields("Nombre")) Then cc.nombre = .Fields("Nombre")           '*** MAKOTO 14/feb/01 Agregado
            If Not IsNull(.Fields("FechaInicio")) Then cc.FechaInicio = .Fields("FechaInicio")
            If Not IsNull(.Fields("FechaFinal")) Then cc.FechaFinal = .Fields("FechaFinal")
            
            If Not IsNull(.Fields("CodProveedor")) Then cc.codProveedor = .Fields("CodProveedor")   '*** MAKOTO 06/mar/01 Agregado
            If Not IsNull(.Fields("CodCliente")) Then cc.codcliente = .Fields("CodCliente")       '*** MAKOTO 06/mar/01 Agregado
            
            cc.Grabar
            
            Dim tabla As String, cod As String, campo As String
            campo = "CodCentro"
            cod = .Fields("CodCentro")
            tabla = "gncentrocosto"
            CambiaFechaenTabla campo, cod, tabla
            
            GrabarGNCentroCosto = GrabarGNCentroCosto + 1
siguiente:
            .MoveNext
        Loop
        
        .Close
    End With
    

salida:
    Set rs = Nothing
    Set cc = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarGNCentroCosto = GrabarGNCentroCosto * -1
    Exit Function

ErrTrap:
    If Not (cc Is Nothing) Then
        s = Err.Description & ": " & Err.Source & vbCr & cc.CodCentro & ", " & cc.Descripcion
    End If
    DispMsg "Importar datos de Centro de costo", "Error", s
    If MsgBox(s & vbCr & vbCr & _
                "Desea continuar con el siguiente registro?", _
                vbQuestion + vbYesNo) = vbYes Then
        Resume siguiente
    Else
        mCancelado = True
    End If
    GoTo salida
End Function

Private Function GrabarTSBanco() As Long
    Dim sql As String, rs As Recordset, ts As TSBanco, i As Long
    Dim s As String, resp As E_MiMsgBox
    
    On Error GoTo ErrTrap
    
    resp = mmsgSi
    
    'Abre el orígen
    sql = "SELECT * FROM TSBanco ORDER BY CodBanco"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    
    With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Bancos ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
        
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            
            'Primero busca si existe ya el mismo código
            Set ts = gobjMain.EmpresaActual.RecuperaTSBanco(.Fields("CodBanco"))
            If ts Is Nothing Then
                'Si no existe el código, crea nuevo
                Set ts = gobjMain.EmpresaActual.CreaTSBanco
            Else
                'Si no se ha hecho pregunta, ó no ha contestado para Todo
                If (resp = mmsgSi) Or (resp = mmsgNo) Then
                    'Pregunta si quiere sobreescribir o no
                    s = "El registro '" & ts.codBanco & "' (" & ts.Descripcion & ") ya existe en el destino." & vbCr & vbCr & _
                        "Desea sobreescribirlo?"
                    resp = frmMiMsgBox.MiMsgBox(s, "Banco")
                End If
                Select Case resp
                Case mmsgCancelar
                    mCancelado = True
                    Exit Do
                Case mmsgNo, mmsgNoTodo
                    GoTo siguiente
                End Select
            End If
            
            ts.codBanco = .Fields("CodBanco")
            ts.Descripcion = .Fields("Descripcion")
            
            If Not mPlantilla.BandIgnorarContabilidad Then     '*** MAKOTO 14/mar/01 Agregado
                If Len(.Fields("CodCuentaContable")) > 0 Then
                    ts.CodCuentaContable = .Fields("CodCuentaContable")
                End If
            End If
            
            If Not IsNull(.Fields("Nombre")) Then ts.nombre = .Fields("Nombre")
            If Not IsNull(.Fields("NumCuenta")) Then ts.NumCuenta = .Fields("NumCuenta")
            If Not IsNull(.Fields("BandValida")) Then ts.BandValida = .Fields("BandValida")
                        
            ts.Grabar
            GrabarTSBanco = GrabarTSBanco + 1
siguiente:
            .MoveNext
        Loop
        
        .Close
    End With
    

salida:
    Set rs = Nothing
    Set ts = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarTSBanco = GrabarTSBanco * -1
    Exit Function

ErrTrap:
    If Not (ts Is Nothing) Then
        s = Err.Description & ": " & Err.Source & vbCr & ts.codBanco & ", " & ts.Descripcion
    End If
    DispMsg "Importar datos de Banco", "Error", s
    If MsgBox(s & vbCr & vbCr & _
                "Desea continuar con el siguiente registro?", _
                vbQuestion + vbYesNo) = vbYes Then
        Resume siguiente
    Else
        mCancelado = True
    End If
    GoTo salida
End Function
        
'*** MAKOTO 12/feb/01 Agregado
Private Function GrabarTSRetencion() As Long
    Dim sql As String, rs As Recordset, ts As TSRetencion, i As Long
    Dim s As String, resp As E_MiMsgBox
    
    On Error GoTo ErrTrap
    
    resp = mmsgSi
    
    'Abre el orígen
    sql = "SELECT * FROM TSRetencion ORDER BY CodRetencion"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    
    With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Retenciones ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
        
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            
            'Primero busca si existe ya el mismo código
            Set ts = gobjMain.EmpresaActual.RecuperaTSRetencion(.Fields("CodRetencion"))
            If ts Is Nothing Then
                'Si no existe el código, crea nuevo
                Set ts = gobjMain.EmpresaActual.CreaTSRetencion
            Else
                'Si no se ha hecho pregunta, ó no ha contestado para Todo
                If (resp = mmsgSi) Or (resp = mmsgNo) Then
                    'Pregunta si quiere sobreescribir o no
                    s = "El registro '" & ts.CodRetencion & "' (" & ts.Descripcion & ") ya existe en el destino." & vbCr & vbCr & _
                        "Desea sobreescribirlo?"
                    resp = frmMiMsgBox.MiMsgBox(s, "Retenciones")
                End If
                Select Case resp
                Case mmsgCancelar
                    mCancelado = True
                    Exit Do
                Case mmsgNo, mmsgNoTodo
                    GoTo siguiente
                End Select
            End If
            
            ts.CodRetencion = .Fields("CodRetencion")
            ts.Descripcion = .Fields("Descripcion")
            
            If Not mPlantilla.BandIgnorarContabilidad Then     '*** MAKOTO 14/mar/01 Agregado
                If Len(.Fields("CodCuentaActivo")) > 0 Then
                    ts.CodCuentaActivo = .Fields("CodCuentaActivo")
                End If
                If Len(.Fields("CodCuentaPasivo")) > 0 Then
                    ts.CodCuentaPasivo = .Fields("CodCuentaPasivo")
                End If
            End If
                
            ts.porcentaje = .Fields("Porcentaje")
            ts.BandValida = .Fields("BandValida")
            'jeaa 21/09/2005
            If Len(.Fields("CodSRI")) > 0 Then
                ts.CodSRI = .Fields("CodSRI")
            End If
            'jeaa 08/07/2008
            ts.bandIVA = .Fields("BandIVA")
            ts.BandCompras = .Fields("BandCompras")
            ts.BandVentas = .Fields("BandVentas")
                        
            ts.Grabar
            
            Dim tabla As String, cod As String, campo As String
            campo = "CodRetencion"
            cod = .Fields("CodRetencion")
            tabla = "TSRetencion"
            CambiaFechaenTabla campo, cod, tabla
            
            
            GrabarTSRetencion = GrabarTSRetencion + 1
siguiente:
            .MoveNext
        Loop
        
        .Close
    End With
    

salida:
    Set rs = Nothing
    Set ts = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarTSRetencion = GrabarTSRetencion * -1
    Exit Function

ErrTrap:
    If Not (ts Is Nothing) Then
        s = Err.Description & ": " & Err.Source & vbCr & ts.CodRetencion & ", " & ts.Descripcion
    End If
    DispMsg "Importar datos de Banco", "Error", s
    If MsgBox(s & vbCr & vbCr & _
                "Desea continuar con el siguiente registro?", _
                vbQuestion + vbYesNo) = vbYes Then
        Resume siguiente
    Else
        mCancelado = True
    End If
    GoTo salida
End Function

Private Function GrabarIVBodega() As Long
    Dim sql As String, rs As Recordset, obj As IVBodega, i As Long
    Dim s As String, resp As E_MiMsgBox
    
    On Error GoTo ErrTrap
    
    resp = mmsgSi
    
    'Abre el orígen
    sql = "SELECT * FROM IVBodega ORDER BY CodBodega"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    
    With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Bodegas ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
        
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            
            'Primero busca si existe ya el mismo código
            Set obj = gobjMain.EmpresaActual.RecuperaIVBodega(.Fields("CodBodega"))
            If obj Is Nothing Then
                'Si no existe el código, crea nuevo
                Set obj = gobjMain.EmpresaActual.CreaIVBodega
            Else
                'Si no se ha hecho pregunta, ó no ha contestado para Todo
                If (resp = mmsgSi) Or (resp = mmsgNo) Then
                    'Pregunta si quiere sobreescribir o no
                    s = "El registro '" & obj.CodBodega & "' (" & obj.Descripcion & ") ya existe en el destino." & vbCr & vbCr & _
                        "Desea sobreescribirlo?"
                    resp = frmMiMsgBox.MiMsgBox(s, "Bodega")
                End If
                Select Case resp
                Case mmsgCancelar
                    mCancelado = True
                    Exit Do
                Case mmsgNo, mmsgNoTodo
                    GoTo siguiente
                End Select
            End If
            
            obj.CodBodega = .Fields("CodBodega")
            obj.Descripcion = .Fields("Descripcion")
            obj.BandValida = .Fields("BandValida")
            obj.Grabar
            GrabarIVBodega = GrabarIVBodega + 1
siguiente:
            .MoveNext
        Loop
        
        .Close
    End With
    

salida:
    Set rs = Nothing
    Set obj = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarIVBodega = GrabarIVBodega * -1
    Exit Function

ErrTrap:
    If Not (obj Is Nothing) Then
        s = Err.Description & ": " & Err.Source & vbCr & obj.CodBodega & ", " & obj.Descripcion
    End If
    DispMsg "Importar datos de Bodega ", "Error", s
    If MsgBox(s & vbCr & vbCr & _
                "Desea continuar con el siguiente registro?", _
                vbQuestion + vbYesNo) = vbYes Then
        Resume siguiente
    Else
        mCancelado = True
    End If
    GoTo salida
End Function

Private Function GrabarIVGrupo(ByVal numGrupo As Byte) As Long
    Dim sql As String, rs As Recordset, obj As ivgrupo, i As Long
    Dim s As String, resp As E_MiMsgBox
    Dim codigo As String, bandCambiaCodigo As Boolean
    
    On Error GoTo ErrTrap
    
    resp = mmsgSi
    
    'Abre el orígen
    sql = "SELECT * FROM IVGrupo" & numGrupo & " ORDER BY CodGrupo" & numGrupo
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    
    With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando " & _
                    gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(CInt(numGrupo)) & " ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
        
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            
            '***Angel. 23/dic/2003
            codigo = .Fields("CodGrupo" & numGrupo)
            bandCambiaCodigo = False
Recupera_OtraVez:
            
            'Primero busca si existe ya el mismo código
            Set obj = gobjMain.EmpresaActual.RecuperaIVGrupo(numGrupo, codigo)
            If obj Is Nothing Then
                'Si no existe el código, crea nuevo
                Set obj = gobjMain.EmpresaActual.CreaIVGrupo(numGrupo)
            Else
                If bandCambiaCodigo = False Then '***Angel. 23/dic/2003
                    'Si no se ha hecho pregunta, ó no ha contestado para Todo
                    If (resp = mmsgSi) Or (resp = mmsgNo) Then
                        'Pregunta si quiere sobreescribir o no
                        s = "El registro '" & obj.CodGrupo & "' (" & obj.Descripcion & ") ya existe en el destino." & vbCr & vbCr & _
                            "Desea sobreescribirlo?"
                        resp = frmMiMsgBox.MiMsgBox(s, gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(CInt(numGrupo)))
                    End If
                    Select Case resp
                    Case mmsgCancelar
                        mCancelado = True
                        Exit Do
                    Case mmsgNo, mmsgNoTodo
                        GoTo siguiente
                    End Select
                End If
            End If
            
            obj.CodGrupo = .Fields("CodGrupo" & numGrupo)
            obj.Descripcion = .Fields("Descripcion")
            obj.BandValida = .Fields("BandValida")
            obj.Grabar
            
            Dim tabla As String, cod As String, campo As String
            campo = "CodGrupo" & numGrupo
            cod = .Fields("CodGrupo" & numGrupo)
            tabla = "IVGrupo" & numGrupo
            CambiaFechaenTabla campo, cod, tabla
            
            GrabarIVGrupo = GrabarIVGrupo + 1
siguiente:
            .MoveNext
        Loop
        
        .Close
    End With
    

salida:
    Set rs = Nothing
    Set obj = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarIVGrupo = GrabarIVGrupo * -1
    Exit Function

ErrTrap:
    If Err.Number = NUMERROR_DUPLI Then '***Angel. 23/dic/2003
        Dim consulta As String, datos_destino As String
        s = rs.Fields("Descripcion")
        consulta = "SELECT * FROM IVGrupo" & numGrupo & " WHERE Descripcion='" & s & "'"
        codigo = BuscarxDesc_Nombre(consulta, "CodGrupo" & numGrupo, "Descripcion", "", datos_destino)
        If Len(codigo) = 0 Then Resume presentar_error
        
        s = "Se ha detectado dos registros de " & gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(CInt(numGrupo)) & _
            " con descripciones idénticas." & vbCrLf & _
            "Es posible que se haya modificado el código. " & vbCrLf & vbCrLf & _
            "Origen:  " & vbTab & Trim$(rs.Fields("CodGrupo" & numGrupo)) & vbTab & Trim$(rs.Fields("Descripcion")) & vbCrLf & _
            "Destino: " & vbTab & datos_destino & vbCrLf & vbCrLf & _
            "¿Desea sobreescribir el código en el destino?"
            
        If MsgBox(s, vbYesNo + vbQuestion) = vbYes Then
            bandCambiaCodigo = True
            Resume Recupera_OtraVez
        Else
            Resume siguiente
        End If
    Else
presentar_error:
        If Not (obj Is Nothing) Then
            s = Err.Description & ": " & Err.Source & vbCr & obj.CodGrupo & ", " & obj.Descripcion
        End If
        DispMsg "Importar datos de " & _
                gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(CInt(numGrupo)), "Error", s
        If MsgBox(s & vbCr & vbCr & _
                    "Desea continuar con el siguiente registro?", _
                    vbQuestion + vbYesNo) = vbYes Then
            Resume siguiente
        Else
            mCancelado = True
        End If
        GoTo salida
    End If
End Function
        
Private Function GrabarIVInventario() As Long
    Dim sql As String, rs As Recordset, obj As IVinventario, i As Long
    Dim s As String, resp As E_MiMsgBox, j As Integer, MSGERR As String, codfamilia As String, codProveedor As String
    
    On Error GoTo ErrTrap
    
    resp = mmsgSi
    
    'Abre el orígen
    'sql = "SELECT * FROM IVInventario ORDER BY CodInventario"
    'modificado jeaa 13-05-2004 para que primero cree los hijos y luego los padres
    sql = "SELECT * FROM IVInventario ORDER BY TIPO"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    
    With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Inventario ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
        
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            
            'Primero busca si existe ya el mismo código
            Set obj = gobjMain.EmpresaActual.RecuperaIVInventario(.Fields("CodInventario"))
            If obj Is Nothing Then
                'Si no existe el código, crea nuevo
                Set obj = gobjMain.EmpresaActual.CreaIVInventario
            Else
                'Si no se ha hecho pregunta, ó no ha contestado para Todo
                If (resp = mmsgSi) Or (resp = mmsgNo) Then
                    'Pregunta si quiere sobreescribir o no
                    s = "El registro '" & obj.CodInventario & "' (" & _
                        obj.Descripcion & ") ya existe en el destino." & vbCr & vbCr & _
                        "Desea sobreescribirlo?"
                    resp = frmMiMsgBox.MiMsgBox(s, "Inventario")
                End If
                Select Case resp
                Case mmsgCancelar
                    mCancelado = True
                    Exit Do
                Case mmsgNo, mmsgNoTodo
                    GoTo siguiente
                End Select
            End If
            
            obj.CodInventario = .Fields("CodInventario")
            If Not IsNull(.Fields("CodAlterno1")) Then obj.CodAlterno1 = .Fields("CodAlterno1")
            If Not IsNull(.Fields("CodAlterno2")) Then obj.CodAlterno2 = .Fields("CodAlterno2")
            If Not IsNull(.Fields("Descripcion")) Then obj.Descripcion = .Fields("Descripcion")
            If Not IsNull(.Fields("DescripcionDetalle")) Then obj.DescripcionDetalle = .Fields("DescripcionDetalle")
            
            If Not mPlantilla.BandIgnorarContabilidad Then     '*** MAKOTO 14/mar/01 Agregado
                If Len(.Fields("CodCuentaActivo")) > 0 Then obj.CodCuentaActivo = .Fields("CodCuentaActivo")
                If Len(.Fields("CodCuentaCosto")) > 0 Then obj.CodCuentaCosto = .Fields("CodCuentaCosto")
                If Len(.Fields("CodCuentaVenta")) > 0 Then obj.CodCuentaVenta = .Fields("CodCuentaVenta")
            End If
            
            For j = 1 To IVGRUPO_MAX
                If Len(.Fields("CodGrupo" & j)) > 0 Then obj.CodGrupo(j) = .Fields("CodGrupo" & j)
            Next j
            For j = 1 To 5
                obj.Precio(j) = .Fields("Precio" & j)
                obj.cantLimite(j) = .Fields("CantLimite" & j)
                obj.Comision(j) = .Fields("Comision" & j)
                obj.Descuento(j) = .Fields("Descuento" & j) '***Agregado. 05/ago/2004. Angel
            Next j
            obj.CodMoneda = .Fields("CodMoneda")
            
            'Código de proveedor
            If Len(.Fields("CodProveedor")) > 0 Then
                'Ignorar si no existe el proveedor
                On Error Resume Next
                obj.codProveedor = .Fields("CodProveedor")
                Err.Clear
                On Error GoTo ErrTrap
            End If
            
            obj.ExistenciaMaxima = .Fields("ExistenciaMaxima")
            obj.ExistenciaMinima = .Fields("ExistenciaMinima")
            If Not IsNull(.Fields("Observacion")) Then obj.Observacion = .Fields("Observacion")
            obj.PorcentajeIVA = .Fields("PorcentajeIVA")
            If Not IsNull(.Fields("Unidad")) Then obj.Unidad = .Fields("Unidad")
            obj.UnidadMinimaCompra = .Fields("UnidadMinimaCompra")
            obj.UnidadMinimaVenta = .Fields("UnidadMinimaVenta")
            obj.BandServicio = .Fields("BandServicio")
            obj.BandValida = .Fields("BandValida")
            'Diego 19/02/2003
            obj.Tipo = .Fields("Tipo")
            obj.ValorRecargo = .Fields("ValorRecargo") '***Agregado. 05/ago/2004. Angel
            obj.bandFraccion = .Fields("BandFraccion")  '******* Agregado JEAA 09/04/2005
            obj.BandArea = .Fields("BandArea") '******* Agregado JEAA 15/09/2005
            If Not IsNull(.Fields("CodUnidad")) Then obj.CodUnidad = .Fields("CodUnidad") '******* Agregado JEAA 17/04/2006
            If Not IsNull(.Fields("CodUnidadConteo")) Then obj.CodUnidadConteo = .Fields("CodUnidadConteo") '******* Agregado JEAA 17/04/2006
            If Not IsNull(.Fields("CostoUltimoIngreso")) Then obj.CostoUltimoIngreso = .Fields("CostoUltimoIngreso") '******* Agregado JEAA 17/04/2006
            If Not IsNull(.Fields("PorcentajeICE")) Then obj.PorcentajeICE = .Fields("PorcentajeICE") '******* Agregado JEAA 17/04/2006
            If Not IsNull(.Fields("PorDesperdicio")) Then obj.PorDesperdicio = .Fields("PorDesperdicio") '******* Agregado JEAA 17/04/2006
            If Not IsNull(.Fields("CantRelUnidad")) Then obj.CantRelUnidad = .Fields("CantRelUnidad") '******* Agregado JEAA 03/10/2007
            If Not IsNull(.Fields("CantRelUnidadCont")) Then obj.CantRelUnidadCont = .Fields("CantRelUnidadCont") '******* Agregado JEAA 03/10/2007
            'jeaa 08/07/2008
            If Not IsNull(.Fields("Descripcion2")) Then obj.Descripcion2 = .Fields("Descripcion2")
            If Not IsNull(.Fields("PesoNeto")) Then obj.PesoNeto = .Fields("PesoNeto")
            If Not IsNull(.Fields("PesoBruto")) Then obj.PesoBruto = .Fields("PesoBruto")
            If Not IsNull(.Fields("CodUnidadPeso")) Then obj.CodUnidadPeso = .Fields("CodUnidadPeso")
            obj.BandConversion = .Fields("BandConversion")
            obj.BandRepGastos = .Fields("BandRepGastos")
            obj.BandNoSeFactura = .Fields("BandNoSeFactura")
            obj.TiempoReposicion = .Fields("TiempoReposicion")
            obj.TiempoPromVta = .Fields("TiempoPromVta")
            
            
            obj.Grabar
            
            
            Set obj = Nothing
            If .Fields("Tipo") <> INV_TIPONORMAL Then
                Set obj = gobjMain.EmpresaActual.RecuperaIVInventario(.Fields("CodInventario"))
                If Not obj Is Nothing Then
                    codfamilia = obj.CodInventario
                    GrabaIVFamilia obj, MSGERR
                    
                    If Not (obj Is Nothing) Then
                        obj.Grabar
                    Else
                        s = "Error al asignar : " & codfamilia & ", como Familia " & vbCr & _
                        "Codigo de Item no encontrado: " & MSGERR
                        DispMsg "Importar datos de Inventario", "Error", s
                        If MsgBox(s & vbCr & vbCr & _
                                    "Desea continuar con el siguiente registro?", _
                                    vbQuestion + vbYesNo) = vbYes Then
                        Else
                            mCancelado = True
                        End If
                        'jeaa 01/10/2005
                        'GoTo salida
                        GoTo siguiente
                    End If
                End If
            End If
            'AUC 25/11/05 desde aqui para importar datos de proveedores en detalleproveedor
            Set obj = gobjMain.EmpresaActual.RecuperaIVInventario(.Fields("CodInventario"))
            If Not obj Is Nothing Then
                codProveedor = obj.CodInventario
                GrabaIVProveedor obj, MSGERR
                If Not (obj Is Nothing) Then
                    obj.Grabar
                Else
                    s = "Error al asignar : " & codfamilia & ", como Familia " & vbCr & _
                    "Codigo de Item no encontrado: " & MSGERR
                    DispMsg "Importar datos de Inventario", "Error", s
                    If MsgBox(s & vbCr & vbCr & _
                                "Desea continuar con el siguiente registro?", _
                                vbQuestion + vbYesNo) = vbYes Then
                    Else
                        mCancelado = True
                    End If
                    'jeaa 01/10/2005
                    'GoTo salida
                    GoTo siguiente
                End If
            End If
            Dim tabla As String, cod As String, campo As String
            campo = "CodInventario"
            cod = .Fields("CodInventario")
            tabla = "IvInventario"
            CambiaFechaenTabla campo, cod, tabla
                'AUC Hasta aqui proveedores
            GrabarIVInventario = GrabarIVInventario + 1
siguiente:
            .MoveNext
        Loop
        
        .Close
    End With
    

salida:
    Set rs = Nothing
    Set obj = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarIVInventario = GrabarIVInventario * -1
    Exit Function

ErrTrap:
    If Not (obj Is Nothing) Then
        s = Err.Description & ": " & Err.Source & vbCr & obj.CodInventario & ", " & obj.Descripcion
    End If
    DispMsg "Importar datos de Inventario", "Error", s
    If MsgBox(s & vbCr & vbCr & _
                "Desea continuar con el siguiente registro?", _
                vbQuestion + vbYesNo) = vbYes Then
        Resume siguiente
    Else
        mCancelado = True
    End If
    GoTo salida
End Function
        
Private Sub GrabaIVFamilia(ByRef objItem As IVinventario, ByRef MSGERR As String)
    Dim objfam As IVinventario, objHijo As IVinventario
    Dim sql As String, rs As Recordset, i As Long
    Dim ix As Long, objf As IVFamiliaDetalle
    'Abre el orígen
    MSGERR = "" '' limpio los mensajes de error
    sql = "SELECT * FROM IVMateria Where CodMateria = '" & objItem.CodInventario & "'"
    'sql = "SELECT * FROM IVMateria Where codMateria = '" & objItem.CodInventario & "'"
        
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        rs.MoveLast
        rs.MoveFirst
    End If
    
    'Primero borra lo anterior si es necesario
    For i = objItem.NumFamiliaDetalle To 1 Step -1
        objItem.RemoveDetalleFamilia i
    Next i
    
    With rs
        Do Until .EOF
            'Agrega detalle Familia
            ix = objItem.AddDetalleFamilia   'Aumenta  item  a la coleccion
            Set objf = objItem.RecuperaDetalleFamilia(ix)
            
            
            Set objHijo = gobjMain.EmpresaActual.RecuperaIVInventario(!CodInventario)
            If Not objHijo Is Nothing Then
                objf.CodInventario = !CodInventario
                objf.cantidad = !cantidad
                objf.Precio = !TarifaJornal
                objf.BandPrincipal = !BandPrincipal
                objf.BandModificar = !BandModificar
                objf.TarifaJornal = !TarifaJornal
                objf.Rendimiento = !Rendimiento
                objf.Orden = !Orden
                 objf.xCuanto = !xCuanto
            Else ''No encontro el hijo y sale del ciclo y borra el objpadre para que genere el error en la funcion que le esta llamando
                MSGERR = !CodInventario
                Set objItem = Nothing
                Exit Do
            End If
            .MoveNext
        Loop
    End With
End Sub

Private Function GrabarFCVendedor() As Long
    Dim sql As String, rs As Recordset, obj As FCVendedor, i As Long
    Dim s As String, resp As E_MiMsgBox
    
    On Error GoTo ErrTrap
    
    resp = mmsgSi
    
    'Abre el orígen
    sql = "SELECT * FROM FCVendedor ORDER BY CodVendedor"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        rs.MoveLast
        rs.MoveFirst
    End If
    
    With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Vendedores ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
        
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            
            'Primero busca si existe ya el mismo código
            Set obj = gobjMain.EmpresaActual.RecuperaFCVendedor(.Fields("CodVendedor"))
            If obj Is Nothing Then

                'Si no existe el código, crea nuevo
                Set obj = gobjMain.EmpresaActual.CreaFCVendedor
            Else
                'Si no se ha hecho pregunta, ó no ha contestado para Todo
                If (resp = mmsgSi) Or (resp = mmsgNo) Then
                    'Pregunta si quiere sobreescribir o no
                    s = "El registro '" & obj.CodVendedor & "' (" & _
                            obj.nombre & ") ya existe en el destino." & vbCr & vbCr & _
                        "Desea sobreescribirlo?"
                    resp = frmMiMsgBox.MiMsgBox(s, "Vendedor")
                End If
                Select Case resp
                Case mmsgCancelar
                    mCancelado = True
                    Exit Do
                Case mmsgNo, mmsgNoTodo
                    GoTo siguiente
                End Select
            End If
            
            obj.CodVendedor = .Fields("CodVendedor")
            obj.nombre = .Fields("Nombre")
            obj.BandValida = .Fields("BandValida")
            obj.Grabar
            GrabarFCVendedor = GrabarFCVendedor + 1
siguiente:
            .MoveNext
        Loop
        
        .Close
    End With
    

salida:
    Set rs = Nothing
    Set obj = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarFCVendedor = GrabarFCVendedor * -1
    Exit Function

ErrTrap:
    If Not (obj Is Nothing) Then
        s = Err.Description & ": " & Err.Source & vbCr & obj.CodVendedor & ", " & obj.nombre
    End If
    DispMsg "Importar datos de Vendedor", "Error", s
    If MsgBox(s & vbCr & vbCr & _
                "Desea continuar con el siguiente registro?", _
                vbQuestion + vbYesNo) = vbYes Then
        Resume siguiente
    Else
        mCancelado = True
    End If
    GoTo salida
End Function
        
Private Function GrabarPCGrupo(ByVal numGrupo As Integer) As Long
    Dim sql As String, rs As Recordset, obj As PCGRUPO, i As Long
    Dim s As String, resp As E_MiMsgBox
    Dim codigo As String, bandCambiaCodigo As Boolean
    
    On Error GoTo ErrTrap
    
    resp = mmsgSi
    
    'Abre el orígen
    sql = "SELECT * FROM PCGrupo" & numGrupo & " ORDER BY CodGrupo" & numGrupo
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        rs.MoveLast
        rs.MoveFirst
    End If
    
    With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando " & _
                    gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupo(numGrupo) & " ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
        
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            
            '***Angel. 23/dic/2003
            codigo = .Fields("CodGrupo" & numGrupo)
            bandCambiaCodigo = False
Recupera_OtraVez:

            'Primero busca si existe ya el mismo código
            Set obj = gobjMain.EmpresaActual.RecuperaPCGrupo(CByte(numGrupo), codigo)
            If obj Is Nothing = 0 Then
                'Si no existe el código, crea nuevo
                Set obj = gobjMain.EmpresaActual.CreaPCGrupo(CByte(numGrupo))
            Else
                If bandCambiaCodigo = False Then
                    'Si no se ha hecho pregunta, ó no ha contestado para Todo
                    If (resp = mmsgSi) Or (resp = mmsgNo) Then
                        'Pregunta si quiere sobreescribir o no
                        s = "El registro '" & obj.CodGrupo & "' (" & obj.Descripcion & ") ya existe en el destino." & vbCr & vbCr & _
                            "Desea sobreescribirlo?"
                        resp = frmMiMsgBox.MiMsgBox(s, gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupo(numGrupo))
                    End If
                    Select Case resp
                    Case mmsgCancelar
                        mCancelado = True
                        Exit Do
                    Case mmsgNo, mmsgNoTodo
                        GoTo siguiente
                    End Select
                End If
            End If
            
            obj.CodGrupo = .Fields("CodGrupo" & numGrupo)
            obj.Descripcion = .Fields("Descripcion")
            obj.BandValida = .Fields("BandValida")
            obj.Grabar
            
            Dim tabla As String, cod As String, campo As String
            campo = "CodGrupo" & numGrupo
            cod = .Fields("CodGrupo" & numGrupo)
            tabla = "PCGrupo" & numGrupo
            CambiaFechaenTabla campo, cod, tabla
            
            
            GrabarPCGrupo = GrabarPCGrupo + 1
siguiente:
            .MoveNext
        Loop
        
        .Close
    End With
    

salida:
    Set rs = Nothing
    Set obj = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarPCGrupo = GrabarPCGrupo * -1
    Exit Function

ErrTrap:
    If Err.Number = NUMERROR_DUPLI Then '***Angel. 23/dic/2003
        Dim consulta As String, datos_destino As String
        s = rs.Fields("Descripcion")
        consulta = "SELECT * FROM PCGrupo" & numGrupo & " WHERE Descripcion='" & s & "'"
        codigo = BuscarxDesc_Nombre(consulta, "CodGrupo" & numGrupo, "Descripcion", "", datos_destino)
        If Len(codigo) = 0 Then Resume presentar_error
        
        s = "Se ha detectado dos registros de " & gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupo(numGrupo) & _
            " con descripciones idénticas." & vbCrLf & _
            "Es posible que se haya modificado el código. " & vbCrLf & vbCrLf & _
            "Origen:  " & vbTab & Trim$(rs.Fields("CodGrupo" & numGrupo)) & vbTab & Trim$(rs.Fields("Descripcion")) & vbCrLf & _
            "Destino: " & vbTab & datos_destino & vbCrLf & vbCrLf & _
            "¿Desea sobreescribir el código en el destino?"
            
        If MsgBox(s, vbYesNo + vbQuestion) = vbYes Then
            bandCambiaCodigo = True
            Resume Recupera_OtraVez
        Else
            Resume siguiente
        End If
    Else
presentar_error:
        If Not (obj Is Nothing) Then
            s = Err.Description & ": " & Err.Source & vbCr & obj.CodGrupo & ", " & obj.Descripcion
        End If
        DispMsg "Importar datos de " & _
                gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupo(numGrupo), "Error", s
        If MsgBox(s & vbCr & vbCr & _
                    "Desea continuar con el siguiente registro?", _
                    vbQuestion + vbYesNo) = vbYes Then
            Resume siguiente
        Else
            mCancelado = True
        End If
        GoTo salida
    End If
End Function
        
Private Function GrabarPCProvCli(ByVal Proveedor As Boolean) As Long
    Dim sql As String, rs As Recordset, obj As PCProvCli, i As Long
    Dim s As String, resp As E_MiMsgBox, j As Long, Desc As String
    Dim codigo As String, bandCambiaCodigo As Boolean
    Dim rs2 As Recordset
    
    On Error GoTo ErrTrap
    
    Desc = IIf(Proveedor, "Proveedor", "Cliente")
    resp = mmsgSi
    
    'Abre el orígen
    sql = "SELECT * FROM PCProvCli "
    sql = sql & "WHERE " & IIf(Proveedor, "BandProveedor<>0", "BandCliente<>0")
    sql = sql & " ORDER BY CodProvCli"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    
    With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando " & Desc & " ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
        
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            
            '***Angel. 22/dic/2003
            codigo = .Fields("CodProvCli")
            bandCambiaCodigo = False
Recupera_OtraVez:

            'Primero busca si existe ya el mismo código
            Set obj = gobjMain.EmpresaActual.RecuperaPCProvCli(codigo)
            If obj Is Nothing Then
                'Si no existe el código, crea nuevo
                Set obj = gobjMain.EmpresaActual.CreaPCProvCli
            Else
                If bandCambiaCodigo = False Then '***Angel. 22/dic/2003
                    'Si no se ha hecho pregunta, ó no ha contestado para Todo
                    If (resp = mmsgSi) Or (resp = mmsgNo) Then
                        'Pregunta si quiere sobreescribir o no
                        s = "El registro '" & obj.CodProvCli & "' (" & _
                            obj.nombre & ") ya existe en el destino." & vbCr & vbCr & _
                            "Desea sobreescribirlo?"
                        resp = frmMiMsgBox.MiMsgBox(s, IIf(Proveedor, "Proveedor", "Cliente"))
                    End If
                    Select Case resp
                    Case mmsgCancelar
                        mCancelado = True
                        Exit Do
                    Case mmsgNo, mmsgNoTodo
                        GoTo siguiente
                    End Select
                End If
            End If
            
            obj.CodProvCli = .Fields("CodProvCli")
            obj.nombre = .Fields("Nombre")
            
            If Not mPlantilla.BandIgnorarContabilidad Then     '*** MAKOTO 14/mar/01 Agregado
                If Len(.Fields("CodCuentaContable")) > 0 Then
                    obj.CodCuentaContable = .Fields("CodCuentaContable")
                End If
                If Len(.Fields("CodCuentaContable2")) > 0 Then
                    obj.CodCuentaContable2 = .Fields("CodCuentaContable2")
                End If
            End If
            
            obj.BandCliente = .Fields("BandCliente")
            obj.BandProveedor = .Fields("BandProveedor")
            If Not IsNull(.Fields("CodPostal")) Then obj.CodPostal = .Fields("CodPostal")
            If Not IsNull(.Fields("Direccion1")) Then obj.Direccion1 = .Fields("Direccion1")
            If Not IsNull(.Fields("Direccion2")) Then obj.Direccion2 = .Fields("Direccion2")
            If Not IsNull(.Fields("Ciudad")) Then obj.Ciudad = .Fields("Ciudad")
            If Not IsNull(.Fields("Provincia")) Then obj.Provincia = .Fields("Provincia")
            If Not IsNull(.Fields("Pais")) Then obj.Pais = .Fields("Pais")
            If Not IsNull(.Fields("Telefono1")) Then obj.Telefono1 = .Fields("Telefono1")
            If Not IsNull(.Fields("Telefono2")) Then obj.Telefono2 = .Fields("Telefono2")
            If Not IsNull(.Fields("Telefono3")) Then obj.Telefono3 = .Fields("Telefono3")
            If Not IsNull(.Fields("Fax")) Then obj.Fax = .Fields("Fax")
            If Not IsNull(.Fields("RUC")) Then obj.ruc = .Fields("RUC")
            If Not IsNull(.Fields("EMail")) Then obj.Email = .Fields("EMail")
            If Not IsNull(.Fields("Estado")) Then obj.Estado = .Fields("Estado")
            If Not IsNull(.Fields("CodVendedor")) Then obj.CodVendedor = .Fields("CodVendedor")
            If Not IsNull(.Fields("LimiteCredito")) Then obj.LimiteCredito = .Fields("LimiteCredito")
            
            If Not IsNull(.Fields("CodGrupo1")) Then obj.CodGrupo1 = .Fields("CodGrupo1")
            If Not IsNull(.Fields("CodGrupo2")) Then obj.CodGrupo2 = .Fields("CodGrupo2")
            If Not IsNull(.Fields("CodGrupo3")) Then obj.CodGrupo3 = .Fields("CodGrupo3")
            If Not IsNull(.Fields("CodGrupo4")) Then obj.CodGrupo4 = .Fields("CodGrupo4")
            
            If Not IsNull(.Fields("Banco")) Then obj.banco = .Fields("Banco")
            If Not IsNull(.Fields("NumCuenta")) Then obj.NumCuenta = .Fields("NumCuenta")
            If Not IsNull(.Fields("Swit")) Then obj.Swit = .Fields("Swit")
            If Not IsNull(.Fields("DirecBanco")) Then obj.DirecBanco = .Fields("DirecBanco")
            If Not IsNull(.Fields("TelBanco")) Then obj.TelBanco = .Fields("TelBanco")
            
            '***Agregado. 08/sep/2003. Angel
            '***Campos referentes a Anexos
            If Not IsNull(.Fields("TipoDocumento")) Then
                If Len(.Fields("TipoDocumento")) > 0 Then obj.TipoDocumento = .Fields("TipoDocumento")
                'jeaa 17/05/2007
                Select Case obj.TipoDocumento
                        Case "1", "01": obj.codtipoDocumento = "R"
                        Case "2", "02": obj.codtipoDocumento = "C"
                        Case "5", "05": obj.codtipoDocumento = "O"
                        Case "6", "06": obj.codtipoDocumento = "P"
                        Case "7", "07": obj.codtipoDocumento = "F"
                        Case Else: obj.codtipoDocumento = "T"
                End Select
            End If
            If Not IsNull(.Fields("TipoComprobante")) Then
                If Len(.Fields("TipoComprobante")) > 0 Then obj.TipoComprobante = .Fields("TipoComprobante")
            End If
            If Not IsNull(.Fields("NumAutSRI")) Then
                If Len(.Fields("NumAutSRI")) > 0 Then obj.NumAutSRI = .Fields("NumAutSRI")
            End If
            '***Agregado. 08/sep/2003. Angel
            '***Campos necesarios para tarjetas de descuentos
            If Not IsNull(.Fields("NombreAlterno")) Then obj.NombreAlterno = .Fields("NombreAlterno")
            If Not IsNull(.Fields("FechaNacimiento")) Then obj.FechaNacimiento = .Fields("FechaNacimiento")
            If Not IsNull(.Fields("FechaEntrega")) Then obj.FechaEntrega = .Fields("FechaEntrega")
            If Not IsNull(.Fields("FechaExpiracion")) Then obj.FechaExpiracion = .Fields("FechaExpiracion")
            If Not IsNull(.Fields("TotalDebe")) Then obj.TotalDebe = .Fields("TotalDebe")
            If Not IsNull(.Fields("TotalHaber")) Then obj.TotalHaber = .Fields("TotalHaber")
            If Not IsNull(.Fields("Observacion")) Then obj.Observacion = .Fields("observacion")
            If Not IsNull(.Fields("TipoProvCli")) Then obj.TipoProvCli = .Fields("TipoProvCli") 'jeaa 17/01/2008
            obj.BandEmpresaPublica = .Fields("BandEmpresaPublica")
            obj.BandGarante = .Fields("BandGarante")
            
            If Not IsNull(.Fields("CodProvincia")) Then obj.codProvincia = .Fields("CodProvincia")
            If Not IsNull(.Fields("CodCanton")) Then obj.codCanton = .Fields("CodCanton")
            If Not IsNull(.Fields("CodParroquia")) Then obj.codParroquia = .Fields("CodParroquia")
            If Not IsNull(.Fields("CodDiasCredito")) Then obj.CodDiasCredito = .Fields("CodDiasCredito")
            
            If Not IsNull(.Fields("TipoSujeto")) Then obj.Tiposujeto = .Fields("TipoSujeto")
            If Not IsNull(.Fields("Sexo")) Then obj.sexo = .Fields("Sexo")
            If Not IsNull(.Fields("EstadoCivil")) Then obj.EstadoCivil = .Fields("EstadoCivil")
            If Not IsNull(.Fields("OrigenIngresos")) Then obj.OrigenIngresos = .Fields("OrigenIngresos")
            
            
            
            
            'Agrega los contactos
            AgregarContactos obj
            
            obj.Grabar
            
            Dim tabla As String, cod As String, campo As String
            campo = "CodProvCli"
            cod = .Fields("CodProvCli")
            tabla = "PcProvcli"
            CambiaFechaenTabla campo, cod, tabla
            
            
            GrabarPCProvCli = GrabarPCProvCli + 1
siguiente:
            .MoveNext
        Loop
        
        .Close
    End With

salida:
    Set rs = Nothing
    Set obj = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarPCProvCli = GrabarPCProvCli * -1
    Exit Function

ErrTrap:
    If Err.Number = NUMERROR_DUPLI Then '***Angel. 22/dic/2003
        Dim consulta As String, datos_destino As String
        s = rs.Fields("Nombre")
        consulta = "select * from pcprovcli where nombre='" & s & "'"
        codigo = BuscarxDesc_Nombre(consulta, "CodProvCli", "Nombre", "RUC", datos_destino)
        If Len(codigo) = 0 Then Resume presentar_error
        
        s = "Se ha detectado dos registros de " & Desc & " con nombres idénticos." & vbCrLf & _
            "Es posible que se haya modificado el código. " & vbCrLf & vbCrLf & _
            "Origen:  " & vbTab & Trim$(rs.Fields("CodProvCli")) & vbTab & Trim$(rs.Fields("Nombre")) & vbTab & Trim$(rs.Fields("RUC")) & vbCrLf & _
            "Destino: " & vbTab & datos_destino & vbCrLf & vbCrLf & _
            "¿Desea sobreescribir el código en el destino?"
            
        If MsgBox(s, vbYesNo + vbQuestion) = vbYes Then
            bandCambiaCodigo = True
            Resume Recupera_OtraVez
        Else
            Resume siguiente
        End If
    Else
presentar_error:
        If Not (obj Is Nothing) Then
            s = "Error: " & Err.Description & vbCrLf & _
                "Origen: " & Err.Source & vbCrLf & vbCrLf & _
                "Registro Afectado:" & vbCrLf & _
                "   Código: " & obj.CodProvCli & vbCrLf & _
                "   Nombre: " & obj.nombre
        End If
        DispMsg "Importar datos de " & Desc, "Error", s
        If MsgBox(s & vbCr & vbCr & _
                    "Desea continuar con el siguiente registro?", _
                    vbQuestion + vbYesNo) = vbYes Then
            Resume siguiente
        Else
            mCancelado = True
        End If
        GoTo salida
    End If
End Function

Private Sub AgregarContactos(ByVal pc As PCProvCli)
    Dim rs As Recordset, sql As String, j As Long
    
    'Borra los contactos si ha recuperado algo
    Do Until pc.CountContacto = 0
        pc.RemoveContacto pc.CountContacto
    Loop
    
    'Obtiene los contactos del prov/cli
    sql = "SELECT * FROM PCContacto " & _
          "WHERE CodProvCli = '" & pc.CodProvCli & "' " & _
          "ORDER BY Orden"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    With rs
        Do Until .EOF
            j = pc.AddContacto
            pc.Contactos(j).Cargo = .Fields("Cargo")
            pc.Contactos(j).Email = .Fields("EMail")
            pc.Contactos(j).nombre = .Fields("Nombre")
            pc.Contactos(j).Telefono1 = .Fields("Telefono1")
            pc.Contactos(j).Telefono2 = .Fields("Telefono2")
            pc.Contactos(j).titulo = .Fields("Titulo")
            pc.Contactos(j).Orden = j
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
End Sub

Private Sub ImportarTrans()
    Dim i As Long, codt As String, numt As Long
    Dim resp As E_MiMsgBox
    Dim s As String
    Dim fechaT As Date
    sst1.Tab = 0
    resp = mmsgSi
    
    With grdTrans
        For i = .FixedRows To .Rows - 1
            DoEvents
            
            'Si el usuario canceló la operación
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit For
            End If
            
            .ShowCell i, 0          'Hace visible la fila actual
            
'            If .IsSelected(i) Then
                codt = .TextMatrix(i, .ColIndex("CodTrans"))
                numt = .TextMatrix(i, .ColIndex("NumTrans"))
                fechaT = .TextMatrix(i, .ColIndex("fecha"))
                MensajeStatus "Importando la transacción " & codt & numt & _
                            "     " & i & " de " & .Rows - .FixedRows & _
                            " (" & Format(i * 100 / (.Rows - .FixedRows), "0") & "%)", vbHourglass
               
                If ImportarTransSub(codt, numt, resp, fechaT) Then
                    'Sí se grabó sin problema, quita la selección
                    .IsSelected(i) = True
                End If
'            End If
        Next i
    End With
    MensajeStatus
    
End Sub

Private Function ImportarTransSub( _
                ByVal codt As String, _
                ByVal numt As Long, _
                ByRef resp As E_MiMsgBox, _
                ByVal fecha As Date) As Boolean
    Dim gc As GNComprobante, s As String, Estado As Byte, EstadoOriginal As Integer
    Dim sql As String
    On Error GoTo ErrTrap
                
        If fecha < gobjMain.EmpresaActual.GNOpcion.FechaLimiteDesde Then
            If (resp = mmsgSi) Or (resp = mmsgNo) Then
                'Confirma si quiere sobre escribir lo existente
                s = "La fecha de la transacción " & codt & numt & " está fuera de Rango Aceptable" & vbCr & vbCr & _
                    "Desea Importarla?"
                resp = frmMiMsgBox.MiMsgBox(s, codt & numt)
            End If
        Select Case resp
        Case mmsgNoTodo, mmsgNo
            DispMsg "Importar la trans. " & codt & numt, "Saltado", "Eligió no sobreescribir."
            GoTo salida
        Case mmsgCancelar
            DispMsg "Importar la trans. " & codt & numt, "Cancelado"
            mCancelado = True
            GoTo salida
        End Select
    End If
        
    'Recuperar la transacción en el destino
    Set gc = gobjMain.EmpresaActual.RecuperaGNComprobante(0, codt, numt)
    'Si existe en el destino,
    If Not (gc Is Nothing) Then
        If (resp = mmsgSi) Or (resp = mmsgNo) Then
            'Confirma si quiere sobre escribir lo existente
            s = "La transacción " & codt & numt & " ya existe en la base destino." & vbCr & vbCr & _
                "Desea sobreescribirla?"
            resp = frmMiMsgBox.MiMsgBox(s, codt & numt)
        End If
        Select Case resp
        Case mmsgNoTodo, mmsgNo
            DispMsg "Importar la trans. " & codt & numt, "Saltado", "Eligió no sobreescribir."
            GoTo salida
        Case mmsgCancelar
            DispMsg "Importar la trans. " & codt & numt, "Cancelado"
            mCancelado = True
            GoTo salida
        End Select
        
    'Si no existe,
    Else
        'Crea como nueva
        Set gc = gobjMain.EmpresaActual.CreaGNComprobante(codt)
        gc.numtrans = numt          'Asigna el número de trans.
    End If
    
    'Grabar en las tablas
    PrepararGNComprobante gc, Estado
    PrepararGNOferta gc, Estado
    EstadoOriginal = gc.Estado
    If gc.Estado <> 3 Then
            PrepararIVKardex gc
            PrepararIVKardexRecargo gc
'********************************************************* OJO
' jeaa anulado por dista
'            PrepararAFKardex gc
'            PrepararAFKardexRecargo gc
            
'                PrepararAFKardex gc
            
    Else
'        If Not gc.GNTrans.IVVerificaLimite Then
'            PrepararIVKardex gc
'            PrepararIVKardexRecargo gc
'        Else
            BorraIVKardex gc
'            PrepararIVKardex gc
'        End If
    End If

    'jeaa 08-05-2007
    'si estado estado =3 y si es de tesoreria NO CREA PCKARDEX
    If gc.Estado <> 3 Then
        PrepararPCKardex gc
        PrepararPCKardexCHP gc
    Else
        
        If gc.GNTrans.Modulo = "IV" Then
            PrepararPCKardex gc
            
        ElseIf gc.GNTrans.Modulo = "TS" Then
            BorraPCKardex gc
            BorraPCKardexCHP gc
        Else
            BorraPCKardex gc
            PrepararPCKardex gc
        End If
    End If
    PrepararTSKardex gc
    PrepararTSKardexRet gc
    
    If Not mPlantilla.BandIgnorarContabilidad Then     '*** MAKOTO 14/mar/01 Agregado
        PrepararCTLibroDetalle gc
        PrepararPRLibroDetalle gc
    End If
    
    'Graba la transacción
    gc.Grabar False, False

    'Forzar el valor de Estado original, debido a que al Grabar cambia sin querer
    'ALEX mayo-2002
    On Error Resume Next
    If Estado = ESTADO_ANULADO Then
        gc.Empresa.CambiaEstadoGNComp gc.TransID, ESTADO_NOAPROBADO
        gc.Empresa.CambiaEstadoGNComp gc.TransID, Estado
    End If
    
    CambiaUsuariosenGNComprobante gc.CodTrans, gc.numtrans
    CambiaFechaenGNComprobante gc.CodTrans, gc.numtrans
    CambiaEstadoGNComprobante gc.CodTrans, gc.numtrans, EstadoOriginal
    
    Err.Clear
    On Error GoTo ErrTrap
    
    
    If gc.numtrans <> numt Then
        DispMsg "Importar la trans. " & codt & numt, "OK", "Grabado como " & gc.CodTrans & gc.numtrans & "."
    Else
        If Estado <> 3 Then
            DispMsg "Importar la trans. " & codt & numt, "OK", "Grabado con el mismo número."
        Else
            DispMsg "Importar la trans. " & codt & numt, "OK", "Trans. Anulada grabado con el mismo número."
        End If
    End If

    ImportarTransSub = True
salida:
    Set gc = Nothing
    Exit Function
ErrTrap:
    DispMsg "Importar la trans. " & codt & numt, "Error", Err.Description
    If MsgBox(Err.Description & vbCr & vbCr & _
                "Desea continuar con siguiente transacción?", _
                vbQuestion + vbYesNo) <> vbYes Then
        mCancelado = True
    End If
    GoTo salida
End Function

        
Private Sub PrepararGNComprobante( _
                ByVal gc As GNComprobante, _
                ByRef Estado As Byte)
    Dim sql As String, rs As Recordset, id As Long
    Dim LoQueNoExiste As String
    On Error GoTo ErrTrap
    'Abre el orígen para recuperar registro
    sql = "SELECT * FROM GNComprobante " & _
          "WHERE CodTrans = '" & gc.CodTrans & "' AND NumTrans = " & gc.numtrans
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly

    With gc
'        .CodAsiento = rs.Fields("CodAsiento")
        .FechaTrans = rs.Fields("FechaTrans")
        .HoraTrans = rs.Fields("HoraTrans")
        If Not IsNull(rs.Fields("Descripcion")) Then .Descripcion = rs.Fields("Descripcion")
        LoQueNoExiste = "Código de Codusuario: " & rs.Fields("CodUsuario")
        .codUsuario = rs.Fields("CodUsuario")      'Solo lectura
        On Error Resume Next
        
        '***Agregado. 06/ago/2004. Angel
        LoQueNoExiste = "Código de CodUsuarioMod: " & rs.Fields("CodUsuarioModifica")
        If Err.Number = 3265 Then
            .codUsuarioModifica = rs.Fields("CodUsuario")       'Solo lectura
        Else
            .codUsuarioModifica = rs.Fields("CodUsuarioModifica")       'Solo lectura
        End If
        Err.Clear
        On Error GoTo ErrTrap
        
        LoQueNoExiste = "Código de CodResponsable: " & rs.Fields("CodResponsable")
        .CodResponsable = rs.Fields("CodResponsable")
        If Not IsNull(rs.Fields("NumDocRef")) Then .numDocRef = rs.Fields("NumDocRef")
        If Not IsNull(rs.Fields("Nombre")) Then .nombre = rs.Fields("Nombre")
        Estado = rs.Fields("Estado")        'Devuelve el estado para forzar a grabar con este valor
        .Estado = rs.Fields("Estado")
'        rs.Fields("PosID") = .PosID                'Solo lectura
        .NumTransCierrePOS = rs.Fields("NumTransCierrePOS")
        LoQueNoExiste = "Código de Centro de Costo: " & rs.Fields("CodCentro")
        If Not IsNull(rs.Fields("CodCentro")) Then .CodCentro = rs.Fields("CodCentro")

        'CodTransFuente + NumTransFuente --> IdTransFuente
        id = RecuperarCampo("GNComprobante", "TransID", _
                    "CodTrans='" & rs.Fields("CodTransFuente") & "' AND " & _
                    "NumTrans=" & rs.Fields("NumTransFuente"))
        .idTransFuente = id

        .CodMoneda = rs.Fields("CodMoneda")
        .Cotizacion(2) = rs.Fields("Cotizacion2")
        .Cotizacion(3) = rs.Fields("Cotizacion3")
        .Cotizacion(4) = rs.Fields("Cotizacion4")

        'Codxxxx --> IdProveedorRef,IdClienteRef, IdVendedor (Hace dentro el objeto)
        LoQueNoExiste = "Código de  Proveedor: " & rs.Fields("CodProveedorRef")
        If Not IsNull(rs.Fields("CodProveedorRef")) Then .CodProveedorRef = rs.Fields("CodProveedorRef")
        LoQueNoExiste = "Código de Cliente: " & rs.Fields("CodClienteRef")
        If Not IsNull(rs.Fields("CodClienteRef")) Then .CodClienteRef = rs.Fields("CodClienteRef")
        LoQueNoExiste = "Código de Vendedor: " & rs.Fields("CodVendedor")
        If Not IsNull(rs.Fields("CodVendedor")) Then .CodVendedor = rs.Fields("CodVendedor")
        'jeaa 12/09/08
        LoQueNoExiste = "Impresion: " & rs.Fields("Impresion")
        If Err.Number <> 3265 Then
            If Not IsNull(rs.Fields("Impresion")) Then .Impresion = rs.Fields("Impresion")
        End If
        'jeaa 17/06/2005
        LoQueNoExiste = "Código de Motivo: " & rs.Fields("CodMotivo")
        If Not IsNull(rs.Fields("CodMotivo")) Then .CodMotivo = rs.Fields("CodMotivo")
        'jeaa 17/04/2006
        LoQueNoExiste = "Observacion: " & rs.Fields("Observacion")
        If Not IsNull(rs.Fields("Observacion")) Then .Observacion = rs.Fields("Observacion")
        LoQueNoExiste = "Comision: " & rs.Fields("Comision")
        If Not IsNull(rs.Fields("Comision")) Then .Comision = rs.Fields("Comision")
        LoQueNoExiste = "Fecha Devol: " & rs.Fields("FechaDevol")
        If Not IsNull(rs.Fields("FechaDevol")) Then .FechaDevol = rs.Fields("FechaDevol")
        If Not IsNull(rs!NumDias) Then gc.NumDias = rs!NumDias
        If Not IsNull(rs!AutorizacionSRI) Then gc.AutorizacionSRI = rs!AutorizacionSRI
        If Not IsNull(rs!FechaCaducidadSRI) Then gc.FechaCaducidadSRI = rs!FechaCaducidadSRI
        
        '***Agregado. 26/09/2008
        If Not IsNull(rs!CodUsuarioAutoriza) Then gc.CodUsuarioAutoriza = rs!CodUsuarioAutoriza
        gc.Estado1 = rs!Estado1
        gc.Estado2 = rs!Estado2
        LoQueNoExiste = "Código de  Garante: " & rs.Fields("CodGaranteRef")
        If Not IsNull(rs.Fields("CodGaranteRef")) Then .CodGaranteRef = rs.Fields("CodGaranteRef")
        
        Err.Clear
        LoQueNoExiste = ""
    End With
    rs.Close
'----AUC 08/11/2005----desde aqui prepara anexos
    If gc.Empresa.GNOpcion.ObtenerValor("PermiteControlAspectosAnexos") = "1" And _
          gc.GNTrans.IVVisibleAnexos And gc.Estado <> ESTADO_ANULADO Then
          sql = "SELECT * FROM Anexos " & _
        "WHERE CodTrans = '" & gc.CodTrans & "' AND NumTrans = " & gc.numtrans
          Set rs = New Recordset
          rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
        With gc
            If Not IsNull(rs!CodCredTrib) Then gc.CodCredTrib = rs!CodCredTrib
            If Not IsNull(rs!CodTipoComp) Then gc.CodTipoComp = rs!CodTipoComp
            If Not IsNull(rs!NumAutSRI) Then gc.NumAutSRI = rs!NumAutSRI
            If Not IsNull(rs!NumSerie) Then gc.NumSerie = rs!NumSerie
            If Not IsNull(rs!NumSecuencial) Then gc.NumSecuencial = rs!NumSecuencial
            If Not IsNull(rs!BandDevolucion) Then gc.BandDevolucion = rs!BandDevolucion
            If Not IsNull(rs!TransIDAfectada) Then gc.TransIDAfectada = rs!TransIDAfectada
            If Not IsNull(rs!FechaAnexos) Then gc.FechaAnexos = rs!FechaAnexos
            'jeaa 17/04/06
            If Not IsNull(rs!NumSerieEstablecimiento) Then gc.NumSerieEstablecimiento = rs!NumSerieEstablecimiento
            If Not IsNull(rs!NumSeriePunto) Then gc.NumSeriePunto = rs!NumSeriePunto
            If Not IsNull(rs!FechaCaducidad) Then gc.FechaCaducidad = rs!FechaCaducidad
            If Not IsNull(rs!BandFactElec) Then gc.BandFactElec = rs!BandFactElec
            rs.Close
        End With
    End If
    Set rs = Nothing
    
'----AUC 08/11/2005----desde aqui prepara anexos
    If gc.GNTrans.IVAplicaFinaciamiento And gc.Estado <> ESTADO_ANULADO Then
          sql = "SELECT * FROM GnFinanciamiento " & _
        "WHERE CodTrans = '" & gc.CodTrans & "' AND NumTrans = " & gc.numtrans
          Set rs = New Recordset
          rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
        With gc
            If Not IsNull(rs!TasaMensual) Then gc.TasaMensual = rs!TasaMensual
            If Not IsNull(rs!MesesGracia) Then gc.MesesGracia = rs!MesesGracia
            If Not IsNull(rs!ValorEntrada) Then gc.ValorEntrada = rs!ValorEntrada
            If Not IsNull(rs!FechaPrimerPago) Then gc.FechaPrimerPago = rs!FechaPrimerPago
            If Not IsNull(rs!DiaPago) Then gc.DiaPago = rs!DiaPago
            If Not IsNull(rs!NumeroPagos) Then gc.NumeroPagos = rs!NumeroPagos
            If Not IsNull(rs!ValorSegundaEntrada) Then gc.ValorSegundaEntrada = rs!ValorSegundaEntrada
            If Not IsNull(rs!FechaSegundoPago) Then gc.FechaSegundoPago = rs!FechaSegundoPago
            If Not IsNull(rs!ValorIntereses) Then gc.ValorIntereses = rs!ValorIntereses
            rs.Close
        End With
    End If
    Set rs = Nothing
    
    
    
    Set gc = Nothing
    '--------Hasta aqui
    Exit Sub
ErrTrap:
    If Err.Number = -2147220960 Then
        Set gc = Nothing
        Set rs = Nothing
        Err.Raise Err.Number, "Importacion.PrepararGNComprobante", "No existe " & LoQueNoExiste
        Exit Sub
    End If
    Set gc = Nothing
    Set rs = Nothing
    Err.Raise Err.Number, "Importacion", Err.Description
End Sub

Private Sub PrepararIVKardex(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset, ivk As IVKardex, i As Long
    Dim LoQueNoExiste As String
    On Error GoTo ErrTrap
   'Primero limpia
    gc.BorrarIVKardex
    'Abre el destino para agregar registro
    sql = "SELECT * FROM IVKardex " & _
          "WHERE CodTrans = '" & gc.CodTrans & "' AND NumTrans = " & gc.numtrans & _
          " ORDER BY Orden"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    Do Until rs.EOF
        DoEvents
        i = gc.AddIVKardex
        Set ivk = gc.IVKardex(i)
        With ivk
            LoQueNoExiste = "Código de Item: " & rs.Fields("CodInventario")
            .CodInventario = rs.Fields("CodInventario")
            LoQueNoExiste = "Código de Bodega: " & rs.Fields("CodBodega")
            .CodBodega = rs.Fields("CodBodega")
            LoQueNoExiste = "?"
           .cantidad = rs.Fields("Cantidad")
            .CostoTotal = rs.Fields("CostoTotal")
            .CostoRealTotal = rs.Fields("CostoRealTotal")
            .PrecioTotal = rs.Fields("PrecioTotal")
            .PrecioRealTotal = rs.Fields("PrecioRealTotal")
            If Not IsNull(rs.Fields("Descuento")) Then .Descuento = rs.Fields("Descuento")
            If Not IsNull(rs.Fields("IVA")) Then .IVA = rs.Fields("IVA")
            .Orden = rs.Fields("Orden")
            If Not IsNull(rs.Fields("Nota")) Then .Nota = rs.Fields("Nota")
            
            On Error Resume Next
            LoQueNoExiste = "Numero Precio: " & rs.Fields("NumeroPrecio")
            If Err.Number <> 3265 Then
                '***Agregado. 11/sep/2003. Angel
                If Not IsNull(rs.Fields("NumeroPrecio")) Then .NumeroPrecio = rs.Fields("NumeroPrecio")
            End If
            Err.Clear
            
            LoQueNoExiste = "ValorRecargoItem: " & rs.Fields("ValorRecargoItem")
            If Err.Number <> 3265 Then
                '***Agregado. 05/ago/2004. Angel
                If Not IsNull(rs.Fields("ValorRecargoItem")) Then .ValorRecargoItem = rs.Fields("ValorRecargoItem")
            End If
            'jeaa 22/09/2005
            Err.Clear
            LoQueNoExiste = "TiempoEntrega: " & rs.Fields("TiempoEntrega")
            If Err.Number <> 3265 Then
                '***Agregado. 05/ago/2004. Angel
                If Not IsNull(rs.Fields("TiempoEntrega")) Then .TiempoEntrega = rs.Fields("TiempoEntrega")
            End If
            Err.Clear
            LoQueNoExiste = "bandImprimir: " & rs.Fields("bandImprimir")
            If Err.Number <> 3265 Then
                '***Agregado. 05/ago/2004. Angel
                If Not IsNull(rs.Fields("bandImprimir")) Then .bandImprimir = rs.Fields("bandImprimir")
            End If
            Err.Clear
            LoQueNoExiste = "idPadre: " & rs.Fields("idPadre")
            If Err.Number <> 3265 Then
                '***Agregado. 05/ago/2004. Angel
                If Not IsNull(rs.Fields("idPadre")) Then .IdPadre = rs.Fields("idPadre")
            End If
            Err.Clear
            LoQueNoExiste = "bandVer: " & rs.Fields("bandVer")
            If Err.Number <> 3265 Then
                '***Agregado. 05/ago/2004. Angel
                If Not IsNull(rs.Fields("bandVer")) Then .bandVer = rs.Fields("bandVer")
            End If
            Err.Clear
            LoQueNoExiste = "idPadreSub: " & rs.Fields("idPadreSub")
            If Err.Number <> 3265 Then
                '***Agregado. 05/ago/2004. Angel
                If Not IsNull(rs.Fields("idPadreSub")) Then .idpadresub = rs.Fields("idPadreSub")
            End If
            Err.Clear
            
            
            On Error GoTo ErrTrap
            
        End With
        rs.MoveNext
    Loop
    rs.Close

    Set gc = Nothing
    Set ivk = Nothing
    Set rs = Nothing
    Exit Sub
ErrTrap:
    If Err.Number = -2147220960 Then
        Set gc = Nothing
        Set ivk = Nothing
        Set rs = Nothing
        Err.Raise Err.Number, "Importacion.PrepararIVKardex", "No existe " & LoQueNoExiste
        Exit Sub
    End If
    Set gc = Nothing
    Set ivk = Nothing
    Set rs = Nothing
    Err.Raise Err.Number, "Importacion", Err.Description
End Sub

Private Sub PrepararIVKardexRecargo(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset, ivkr As IVKardexRecargo, i As Long
    Dim LoQueNoExiste As String
    On Error GoTo ErrTrap
   'Primero limpia
    gc.BorrarIVKardexRecargo
    'Abre el destino para agregar registro
    sql = "SELECT * FROM IVKardexRecargo " & _
          "WHERE CodTrans = '" & gc.CodTrans & "' AND NumTrans = " & gc.numtrans & _
          " ORDER BY Orden"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    Do Until rs.EOF
        DoEvents
        i = gc.AddIVKardexRecargo
        Set ivkr = gc.IVKardexRecargo(i)
        With ivkr
            LoQueNoExiste = "Código de Recargo: " & rs.Fields("CodRecargo")
           .codRecargo = rs.Fields("CodRecargo")
            .porcentaje = rs.Fields("Porcentaje")
            .valor = rs.Fields("Valor")
            .BandModificable = rs.Fields("BandModificable")
            .BandOrigen = rs.Fields("BandOrigen")
            .BandProrrateado = rs.Fields("BandProrrateado")
            .AfectaIvaItem = rs.Fields("AfectaIvaItem")
            .Orden = rs.Fields("Orden")
        End With
        rs.MoveNext
    Loop
    rs.Close
    Set gc = Nothing
    Set ivkr = Nothing
    Set rs = Nothing
    Exit Sub
ErrTrap:
    If Err.Number = -2147220960 Then Err.Raise Err.Number, "Importacion.PrepararPCKardex", "No existe " & LoQueNoExiste
End Sub

Private Sub PrepararPCKardex(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset, pck As PCKardex, i As Long
    Dim idAsignado As Long, LoQueNoExiste As String
   Dim v() As String
    On Error GoTo ErrTrap
   'Primero limpia
    gc.BorrarPCKardex
    'Abre el destino para agregar registro
    sql = "SELECT * FROM PCKardex " & _
          "WHERE CodTrans = '" & gc.CodTrans & "' AND NumTrans = " & gc.numtrans & _
          " ORDER BY Orden"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    Do Until rs.EOF
        DoEvents
        i = gc.AddPCKardex
        Set pck = gc.PCKardex(i)
        With pck
            'Desactiva la verificación de saldo de doc.asignado
            'Para que no genere error cuando asigna valor de Debe/Haber
            .BandNoVerificarSaldo = True            '*** MAKOTO 22/mar/01 Agregado
            'en esta sección pide el valor de cliente, no lo encuentra y emite error .....  es aquí en donde hay que controlar
            LoQueNoExiste = "Código de Cliente/Proveedor: " & rs.Fields("CodProvCli")
           .CodProvCli = rs.Fields("CodProvCli")
           
           '***Angel. 13/nov/2003. Para que se importe sin necesidad de IdAsignado
           If Not (mPlantilla.BandIgnorarDocAsignado) Then
                If Len(rs.Fields("GuidAsignado")) > 0 Then      '*** MAKOTO 16/mar/01
                    '***Angel. 13/nov/2003. Agregado Mensaje
                    LoQueNoExiste = "PCKardex: " & _
                                    "No se encuentra el documento asignado. " & vbCr & _
                                    "(" & gc.CodTrans & _
                                    gc.numtrans & ") " & vbCr & rs.Fields("GuidAsignado")
                    .SetIdAsignadoPorGuid rs.Fields("GuidAsignado")
                End If
           End If
           
            LoQueNoExiste = "Código de Forma de Pago: " & rs.Fields("CodForma")
            .codforma = rs.Fields("CodForma")
            If Not IsNull(rs.Fields("NumLetra")) Then .NumLetra = rs.Fields("NumLetra")
            .Debe = rs.Fields("Debe")
            .Haber = rs.Fields("Haber")
            .FechaEmision = rs.Fields("FechaEmision")
            '--------- OJO ----------------------------
            'provicional para distablasa hasta actualizar sucursales
            If Mid$(gc.CodTrans, 1, 3) = "CTC" Then
                .FechaVenci = DateAdd("m", 1, rs.Fields("FechaEmision"))
            Else
                .FechaVenci = rs.Fields("FechaVenci")
            End If
            If Not IsNull(rs.Fields("Observacion")) Then .Observacion = rs.Fields("Observacion")
            .Orden = rs.Fields("Orden")
            .Guid = rs.Fields("guid")
            LoQueNoExiste = "Código de Tarjeta " & rs.Fields("CodTarjeta")
            If Not IsNull(rs.Fields("CodTarjeta")) Then .CodTarjeta = rs.Fields("CodTarjeta")
            LoQueNoExiste = "Código de Banco " & rs.Fields("CodBanco")
            If Not IsNull(rs.Fields("CodBanco")) Then .codBanco = rs.Fields("CodBanco")
            If Not IsNull(rs.Fields("NumCuenta")) Then .NumCuenta = rs.Fields("NumCuenta")
            If Not IsNull(rs.Fields("NumCheque")) Then .Numcheque = rs.Fields("NumCheque")
            If Not IsNull(rs.Fields("TitularCta")) Then .TitularCta = rs.Fields("TitularCta")

            .SetIdFromGuid
        End With
        rs.MoveNext
    Loop
    rs.Close
    Set gc = Nothing
    Set pck = Nothing
    Set rs = Nothing
    Exit Sub
ErrTrap:
    If Err.Number = -2147220960 Then Err.Raise Err.Number, "Importacion.PrepararPCKardex", "No existe " & LoQueNoExiste
End Sub
Private Sub PrepararTSKardex(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset, tsk As TSKardex, i As Long
    Dim LoQueNoExiste As String
   'Primero limpia
    gc.BorrarTSKardex
    'Abre el destino para agregar registro
    sql = "SELECT * FROM TSKardex " & _
          "WHERE CodTrans = '" & gc.CodTrans & "' AND NumTrans = " & gc.numtrans & _
          " ORDER BY Orden"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    Do Until rs.EOF
        DoEvents
        i = gc.AddTSKardex
        Set tsk = gc.TSKardex(i)
        With tsk
            LoQueNoExiste = "Código de Banco: " & rs.Fields("CodBanco")
           .codBanco = rs.Fields("CodBanco")
            .Debe = rs.Fields("Debe")
            .Haber = rs.Fields("Haber")
            If Not IsNull(rs.Fields("Nombre")) Then .nombre = rs.Fields("Nombre")
            .CodTipoDoc = rs.Fields("CodTipoDoc")
            If Not IsNull(rs.Fields("NumDoc")) Then .numdoc = rs.Fields("NumDoc")
            .FechaEmision = rs.Fields("FechaEmision")
            .FechaVenci = rs.Fields("FechaVenci")
            If Not IsNull(rs.Fields("Observacion")) Then .Observacion = rs.Fields("Observacion")
            If Not IsNull(rs.Fields("BandConciliado")) Then .BandConciliado = rs.Fields("BandConciliado")
            .Orden = rs.Fields("Orden")
        End With
        rs.MoveNext
    Loop
    rs.Close
    Set gc = Nothing
    Set tsk = Nothing
    Set rs = Nothing
    Exit Sub
ErrTrap:
    If Err.Number = -2147220960 Then Err.Raise Err.Number, "Importacion.PrepararPCKardex", "No existe " & LoQueNoExiste
End Sub
'*** MAKOTO 12/feb/01 Agregado
Private Sub PrepararTSKardexRet(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset, tskr As TSKardexRet, i As Long
    Dim LoQueNoExiste As String
    On Error GoTo ErrTrap
   'Primero limpia
    gc.BorrarTSKardexRet
    'Abre el destino para agregar registro
    sql = "SELECT * FROM TSKardexRet " & _
          "WHERE CodTrans = '" & gc.CodTrans & "' AND NumTrans = " & gc.numtrans & _
          " ORDER BY Orden"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    Do Until rs.EOF
        DoEvents
        i = gc.AddTSKardexRet
        Set tskr = gc.TSKardexRet(i)
        With tskr
            LoQueNoExiste = "Código de Retención: " & rs.Fields("CodRetencion")
           .CodRetencion = rs.Fields("CodRetencion")
            .valor = rs.Fields("Debe") + rs.Fields("Haber")
            .base = rs.Fields("Base")
            If Not IsNull(rs.Fields("NumDoc")) Then .numdoc = rs.Fields("NumDoc")
            If Not IsNull(rs.Fields("Observacion")) Then .Observacion = rs.Fields("Observacion")
            .Orden = rs.Fields("Orden")
        End With
        rs.MoveNext
    Loop
    rs.Close
    Set gc = Nothing
    Set tskr = Nothing
    Set rs = Nothing
        Exit Sub
ErrTrap:
    If Err.Number = -2147220960 Then Err.Raise Err.Number, "Importacion.PrepararPCKardex", "No existe " & LoQueNoExiste
End Sub
Private Sub PrepararCTLibroDetalle(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset, ctd As CTLibroDetalle, i As Long
    Dim LoQueNoExiste As String
    On Error GoTo ErrTrap
   'Primero limpia
    gc.BorrarCTLibroDetalle
    'Abre el destino para agregar registro
    sql = "SELECT * FROM CTLibroDetalle " & _
          "WHERE CodTrans = '" & gc.CodTrans & "' AND NumTrans = " & gc.numtrans & _
          " ORDER BY Orden"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    Do Until rs.EOF
        DoEvents
        i = gc.AddCTLibroDetalle
        Set ctd = gc.CTLibroDetalle(i)
        With ctd
            LoQueNoExiste = "Código de Cuenta: " & rs.Fields("CodCuenta")
           .codcuenta = rs.Fields("CodCuenta")
            If Not IsNull(rs.Fields("Descripcion")) Then .Descripcion = rs.Fields("Descripcion")
            .Debe = rs.Fields("Debe")
            .Haber = rs.Fields("Haber")
            .BandIntegridad = rs.Fields("BandIntegridad")
            .Orden = rs.Fields("Orden")
        End With
        rs.MoveNext
    Loop
    rs.Close
    Set gc = Nothing
    Set ctd = Nothing
    Set rs = Nothing
            Exit Sub
ErrTrap:
    If Err.Number = -2147220960 Then Err.Raise Err.Number, "Importacion.PrepararPCKardex", "No existe " & LoQueNoExiste
End Sub

Private Sub cmdPlantilla_Click()
    If Len(cboPlantilla.Text) > 0 Then mUltimaPlantilla = cboPlantilla.Text
    
    frmListaPlantilla.Inicio "PlantillasImportar", mcnPlantilla
    'Para que actualizar si hubo cambios en la plantilla seleccionada
    CargarComboPlantilla
    PrepararPlantilla
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF5
        cmdBuscar_Click
        KeyCode = 0
    Case vbKeyF9
        cmdImportar_Click
        KeyCode = 0
    Case Else
        MoverCampo Me, KeyCode, Shift, True
    End Select
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    ImpideSonidoEnter Me, KeyAscii
End Sub
Private Sub Form_Load()
    sst1.Tab = 0
    grdTrans.Rows = grdTrans.FixedRows
    grdMsg.Rows = grdMsg.FixedRows
    CargarCatalogos grdCat
    
    'Recupera parámetros guardados en el registro
    RecuperarConfig
    
    '***Angel. 27/feb/2004
    AbrirBasePlantilla
    CargarComboPlantilla
    PrepararPlantilla
    
    '***Angel. 27/feb/2004
    '***Solo supervisor puede acceder a trabajar con las plantillas
    cmdPlantilla.Enabled = gobjMain.UsuarioActual.BandSupervisor
End Sub
Private Sub GuardarConfig()
    Dim pos As Integer, s As String
    SaveSetting APPNAME, App.Title, Me.Name & ".UltimaPlantilla", cboPlantilla.Text
    pos = InStrRev(txtOrigen.Text, "\")
    If pos > 0 Then
        s = Mid$(txtOrigen.Text, 1, pos)
    Else
        s = ""
    End If
    SaveSetting APPNAME, App.Title, Me.Name & ".RutaBDDestino", s
End Sub
Private Sub RecuperarConfig()
    mUltimaPlantilla = GetSetting(APPNAME, App.Title, Me.Name & ".UltimaPlantilla", "")
    mRutaBDDestino = GetSetting(APPNAME, App.Title, Me.Name & ".RutaBDDestino", App.Path)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'No permitir cerrar la ventana mientras esté ejecutando
    Cancel = mEjecutando
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    'Tamaño minimo de la ventana
    If Me.Height < HEIGHT_MIN Then Me.Height = HEIGHT_MIN
    If Me.Width < WIDTH_MIN Then Me.Width = WIDTH_MIN
End Sub
Private Sub DispMsg( _
                ByVal proc As String, _
                ByVal resultado As String, _
                Optional ByVal msg As String)
    Dim x As Single
    grdMsg.AddItem proc & vbTab & resultado & vbTab & msg
    grdMsg.Row = grdMsg.Rows - 1        'Ubica a la ultima fila
    x = grdMsg.CellTop                     'Para visualizar la fila actual
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'Libera el objeto de nivel de modulo
    Set mcnOrigen = Nothing
    CerrarBasePlantilla '***Angel. 27/feb/2004
End Sub

Private Sub grdCat_BeforeSort(ByVal col As Long, Order As Integer)
    'Impide que cambie el orden de filas mientras esté procesando
    If mEjecutando Then Order = flexSortNone
End Sub

Private Sub grdMsg_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mEjecutando Then Exit Sub
    If Button And vbRightButton Then
        Me.PopupMenu mnuPopUp
    End If
End Sub

Private Sub grdTrans_BeforeSort(ByVal col As Long, Order As Integer)
    'Impide que cambie el orden de filas mientras esté procesando
    If mEjecutando Then Order = flexSortNone
End Sub

Private Sub mnuGrabarGrilla_Click()
    Dim archi As String
    On Error GoTo ErrTrap
    archi = App.Path & "\res.txt"
    grdMsg.SaveGrid archi, flexFileCommaText, True
    MsgBox "El resultado fue grabado en '" & archi & "'.", vbInformation
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub
Private Sub sst1_Click(PreviousTab As Integer)
    If (sst1.Tab = 1) And (PreviousTab <> 1) And (Not mEjecutando) Then
        MostrarNumReg
    End If
End Sub
Private Sub MostrarNumReg()
    Dim i As Long, sql As String, rs As Recordset, tabla As String
    On Error GoTo ErrTrap
    If Len(txtOrigen.Text) = 0 Then Exit Sub
    If Not AbrirOrigen Then Exit Sub
    Set rs = New Recordset
    With grdCat
        .Cols = 4
        .TextMatrix(0, .Cols - 1) = "Núm.Reg."
        For i = .FixedRows To .Rows - 1
            tabla = .TextMatrix(i, .ColIndex("tabla"))
            sql = "SELECT Count(*) AS Cnt FROM " & tabla
            If tabla = "PCProvCli(P)" Then
                sql = "SELECT Count(*) AS Cnt FROM PCProvCli WHERE BandProveedor<>0"
            ElseIf tabla = "PCProvCli(C)" Then
                sql = "SELECT Count(*) AS Cnt FROM PCProvCli WHERE BandCliente<>0"
            ElseIf tabla = "PCProvCli(G)" Then
                sql = "SELECT Count(*) AS Cnt FROM PCProvCli WHERE BandGarante<>0"
            ElseIf tabla = "TSFormaC_P" Then
                sql = "SELECT Count(*) AS Cnt FROM TSFormaCobroPago"
            ElseIf tabla = "GNResp" Then
                sql = "SELECT Count(*) AS Cnt FROM GnResponsable"
            ElseIf tabla = "IVInv" Then
                sql = "SELECT count (*) as Cnt from IVInventario"
            ElseIf tabla = "IVG1" Then
                sql = "SELECT count (*) as Cnt from IVGrupo1"
            ElseIf tabla = "IVG2" Then
                sql = "SELECT count (*) as Cnt from IVGrupo2"
            ElseIf tabla = "IVG3" Then
                sql = "SELECT count (*) as Cnt from IVGrupo3"
            ElseIf tabla = "IVG4" Then
                sql = "SELECT count (*) as Cnt from IVGrupo4"
            ElseIf tabla = "IVG5" Then
                sql = "SELECT count (*) as Cnt from IVGrupo5"
            ElseIf tabla = "IVG6" Then
                sql = "SELECT count (*) as Cnt from IVGrupo6"
            ElseIf tabla = "PCG1" Then
                sql = "SELECT count (*) as Cnt from PCGrupo1"
            ElseIf tabla = "PCG2" Then
                sql = "SELECT count (*) as Cnt from PCGrupo2"
            ElseIf tabla = "PCG3" Then
                sql = "SELECT count (*) as Cnt from PCGrupo3"
            ElseIf tabla = "PCG4" Then
                sql = "SELECT count (*) as Cnt from PCGrupo4"
            ElseIf tabla = "TCompra" Then
                sql = "SELECT count (*) as Cnt from IvTipoCompra"
            ElseIf tabla = "IVU" Then
                sql = "SELECT count (*) as Cnt from IvUnidad"
            ElseIf tabla = "PCHistorial" Then
                sql = "SELECT count (*) as Cnt from PCHistorial"
            ElseIf tabla = "Exist" Then
                sql = "SELECT count (*) as Cnt from IvExist "
            ElseIf tabla = "DescIVGPCG" Then
                sql = "SELECT count (*) as Cnt from DescIVGPCG"
                
            ElseIf tabla = "DescNumPagIVG" Then
                sql = "SELECT count (*) as Cnt from DescNumPagIVG"
            ElseIf tabla = "DiasCred" Then
                sql = "SELECT count (*) as Cnt from PcDiasCredito"
                
            ElseIf tabla = "PLAIVGPCG" Then
                sql = "SELECT count (*) as Cnt from PlazoIVGPCG"

            End If
            rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
            If Not rs.EOF Then
                .TextMatrix(i, .Cols - 1) = rs.Fields("Cnt")
            End If
            rs.Close
        Next i
    End With
    Set rs = Nothing
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

''***Angel. 22/dic/2003
'Funcion para buscar los registros de las diferentes tablas que sirven como catalogos
'en base al nombre o descripcion.
Private Function BuscarxDesc_Nombre(ByVal sql As String, _
                                    ByVal campo1 As String, _
                                    ByVal campo2 As String, _
                                    ByVal campo3 As String, _
                                    ByRef datos As String) As String
    Dim rs As ADODB.Recordset
    
    BuscarxDesc_Nombre = ""
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not (rs Is Nothing) Then
        datos = rs.Fields(campo1) & vbTab & rs.Fields(campo2)
        If Len(campo3) > 0 Then datos = datos & vbTab & rs.Fields(campo3)
        BuscarxDesc_Nombre = rs.Fields(campo1)
    End If
    Set rs = Nothing
End Function





Private Sub CmdGuardarRes_Click()   '************ jeaa 26-01-04 para guardar el resultado de la grilla de mensajes
    Dim nomarchivo As String, NUMARCHIVO As Integer, s As String
    Dim numlinea As Long, resp As Integer
    On Error GoTo ErrTrap
    NUMARCHIVO = FreeFile()
    With dlg1
        .filename = ""
        .DialogTitle = "Guardar Resultado de la Importación"
         If Len(.filename) = 0 Then
            .InitDir = App.Path
            .filename = gobjMain.EmpresaActual.CodEmpresa & _
                        Format(Date, "dd-mm-yyyy") & ".txt"
        Else
            .InitDir = .filename
        End If
        .flags = cdlOFNPathMustExist + cdlOFNFileMustExist
        .Filter = "Archivos de Texto (*.txt)|*.doc|Todos (*.*)|*.*"
        .ShowSave
        nomarchivo = .filename
    End With
    'para revisar si existe el archivo
    Open nomarchivo For Input As #NUMARCHIVO
    numlinea = 0
    Do Until EOF(NUMARCHIVO)
       Line Input #NUMARCHIVO, s
        numlinea = 1
        Exit Do
    Loop
    Close #NUMARCHIVO
    If numlinea <> 0 Then
        resp = MsgBox("Archivo ya existe desea sobrescribirlo", vbYesNoCancel)
        If resp = vbYes Then
            If grabaarchivo(nomarchivo) Then
                    MsgBox "Se guardo el archivo con éxito"
                    Exit Sub
            End If
        End If
    Else
            If grabaarchivo(nomarchivo) Then
                    MsgBox "Se guardo el archivo con éxito"
                    Exit Sub
            End If
    End If
    Exit Sub
ErrTrap:
    If Err.Number <> 32755 Then
            If grabaarchivo(nomarchivo) Then
                    MsgBox "Se guardo el archivo con éxito"
                    Exit Sub
            End If
        DispErr
    End If
    Exit Sub
End Sub
Private Function grabaarchivo(ByVal nomarchivo As String) As Boolean
    Dim msg As String, texto As Variant, fil As Long, col As Integer
    Dim NUMARCHIVO As Integer
    On Error GoTo ErrTrap
    NUMARCHIVO = FreeFile()
    Open nomarchivo For Output As #NUMARCHIVO
    msg = "Importación del archivo: " & txtOrigen.Text
    Write #NUMARCHIVO, msg
    msg = "Fecha de Importación: " & Date & ", Hora: " & Time
    Write #NUMARCHIVO, msg
    msg = "*********************************************************************"
    Write #NUMARCHIVO, msg
    msg = ""
    For fil = 1 To grdMsg.Rows - 1
        For col = 0 To 2
            grdMsg.Select fil, col
            msg = msg & " " & grdMsg.Text
        Next col
    Write #NUMARCHIVO, msg
    msg = " "
 Next fil
 grabaarchivo = True
 Close #NUMARCHIVO
    cmdGuardarRes.Enabled = False
    Exit Function
ErrTrap:
    grabaarchivo = False
    If Err.Number <> 32755 Then
        DispErr
    End If
    Exit Function
End Function

'***Angel. 27/feb/2004
Private Sub AbrirBasePlantilla()
    Dim s As String, RutaBD As String, NombreBD As String
    On Error GoTo ErrTrap
    RutaBD = GetSetting(APPNAME, App.Title, "RutaBDPlantilla", App.Path)
    If Right(RutaBD, 1) <> "\" Then RutaBD = RutaBD & "\"
    NombreBD = GetSetting(APPNAME, App.Title, "NombreBDPlantilla", "ConfigSiiToolsA.mdb")

    If mcnPlantilla Is Nothing Then Set mcnPlantilla = New ADODB.Connection
    If mcnPlantilla.State <> adStateClosed Then mcnPlantilla.Close
    
    'Abre la conección con el archivo de destino
    s = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & RutaBD & NombreBD & ";" & _
        "Persist Security Info=False"
    mcnPlantilla.Open s, "admin", ""
    
    Exit Sub

ErrTrap:
    MsgBox Err.Description, vbOKOnly + vbInformation
End Sub

'***Angel. 27/feb/2004
Private Sub CerrarBasePlantilla()
    'Cierra base
    On Error Resume Next
    If mcnPlantilla.State <> adStateClosed Then
        mcnPlantilla.Close
    End If
End Sub

'***Angel. 27/feb/2004
Private Sub CargarComboPlantilla()
    Dim sql As String, rs As Recordset
        
    sql = "SELECT CodPlantilla, Descripcion FROM Plantilla_EI " & _
          "WHERE (Tipo=1) AND (BandValida=True) ORDER BY CodPlantilla"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnPlantilla, adOpenStatic, adLockReadOnly
    
    cboPlantilla.Clear
    With rs
        If Not (.BOF And .EOF) Then
            Do Until .EOF
                cboPlantilla.AddItem !CodPlantilla
                .MoveNext
            Loop
        End If
    End With
    Set rs = Nothing
End Sub

'***Angel. 27/feb/2004
Private Sub PrepararPlantilla()
    Dim i As Integer
    On Error GoTo mensaje
    
    Set mPlantilla = New clsPlantilla
    mPlantilla.Coneccion = mcnPlantilla
    If Right$(mRutaBDDestino, 1) <> "\" Then mRutaBDDestino = mRutaBDDestino & "\"
    If Len(mUltimaPlantilla) Then
        For i = 0 To cboPlantilla.ListCount - 1
            If cboPlantilla.List(i) = mUltimaPlantilla Then
                cboPlantilla.ListIndex = i
                Exit For
            End If
        Next i
    End If
    RecuperarPlantilla
    Exit Sub
    
mensaje:
    MsgBox Err.Description, vbOKOnly + vbExclamation
End Sub

'***Angel. 27/feb/2004
Private Sub RecuperarPlantilla()
    Dim cod As String
    
    cod = cboPlantilla.Text
    If mPlantilla.Recuperar(cod) Then
        Habilitar True
        lblDescripcion.Caption = mPlantilla.Descripcion
        txtOrigen.Text = mRutaBDDestino & mPlantilla.BDDestino
        cmdActualizar.Visible = mPlantilla.BandActualizarCosto
        cmdGuardarRes.Visible = mPlantilla.BandGuardarResultado
        If mPlantilla.BandGuardarResultado Then
            cmdGuardarRes.Enabled = False
        End If
    Else
        mBuscado = False
        Habilitar False
        lblDescripcion.Caption = ""
        txtOrigen.Text = ""
        cmdActualizar.Visible = False
        cmdGuardarRes.Enabled = False
        
        'Para que permita cerrar el formulario
        mEjecutando = False
        frmMain.mnuFile.Enabled = True
        frmMain.mnuHerramienta.Enabled = True
        frmMain.mnuTransferir.Enabled = True
        frmMain.mnuCerrarTodas.Enabled = True
    End If
End Sub


'jeaa 04/01/2005
Private Function GrabarDesctoIVGrupoPCGrupo() As Long
    Dim sql As String, rs As Recordset, obj As DesctoPcGrupoxIVGrupo, i As Long
    Dim s As String, resp As E_MiMsgBox
    On Error GoTo ErrTrap
    resp = mmsgSi
    'Abre el orígen
    sql = "SELECT * FROM DescIVGPCG "
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
   With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Descuento IVGrupo por PCGrupo ... " & _
                   i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            'Primero busca si existe ya el mismo código
            Set obj = gobjMain.EmpresaActual.RecuperaDesctPCxIV(.Fields("CodPCGrupo") & "," & .Fields("CodIVGrupo"))
            If obj Is Nothing Then
               'Si no existe el código, crea nuevo
                Set obj = gobjMain.EmpresaActual.CreaDesctPCxIV
           Else
                'Si no se ha hecho pregunta, ó no ha contestado para Todo
                If (resp = mmsgSi) Or (resp = mmsgNo) Then
                    'Pregunta si quiere sobreescribir o no
                    s = "El registro '" & obj.CodPCGrupo & "' (" & obj.CodPCGrupo & ") ya existe en el destino." & vbCr & vbCr & _
                       "Desea sobreescribirlo?"
                    resp = frmMiMsgBox.MiMsgBox(s, "DescIVGPCG")
               End If
                Select Case resp
                Case mmsgCancelar
                    mCancelado = True
                    Exit Do
                Case mmsgNo, mmsgNoTodo
                    GoTo siguiente
                End Select
            End If
            obj.CodPCGrupo = .Fields("CodPCGrupo")
            obj.CodIVGrupo = .Fields("CodIVGrupo")
            obj.valor = .Fields("Valor")
            obj.Grabar
            GrabarDesctoIVGrupoPCGrupo = GrabarDesctoIVGrupoPCGrupo + 1
siguiente:
            .MoveNext
        Loop
        .Close
    End With
salida:
    Set rs = Nothing
    Set obj = Nothing
   'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarDesctoIVGrupoPCGrupo = GrabarDesctoIVGrupoPCGrupo * -1
   Exit Function
ErrTrap:
    If Not (obj Is Nothing) Then
        s = Err.Description & ": " & Err.Source & vbCr & obj.CodPCGrupo & ", " & obj.CodIVGrupo
   End If
    DispMsg "Importar datos de Descuentos IVGrupo por PCGrupo ", "Error", s
   If MsgBox(s & vbCr & vbCr & _
                "Desea continuar con el siguiente registro?", _
                vbQuestion + vbYesNo) = vbYes Then
        Resume siguiente
    Else
        mCancelado = True
    End If
    GoTo salida
End Function

Private Function GrabarTSFormaCobroPago() As Long
    Dim sql As String, rs As Recordset, ts As TSFormaCobroPago, i As Long
    Dim s As String, resp As E_MiMsgBox
    
    On Error GoTo ErrTrap
    
    resp = mmsgSi
    
    'Abre el orígen
    sql = "SELECT * FROM TSFormaCobroPago ORDER BY CodForma"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    
    With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Formas de Cobro/Pago ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
        
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            
            'Primero busca si existe ya el mismo código
            Set ts = gobjMain.EmpresaActual.RecuperaTSFormaCobroPago(.Fields("CodForma"))
            If ts Is Nothing Then
                'Si no existe el código, crea nuevo
                Set ts = gobjMain.EmpresaActual.CreaTSFormaCobroPago
            Else
                'Si no se ha hecho pregunta, ó no ha contestado para Todo
                If (resp = mmsgSi) Or (resp = mmsgNo) Then
                    'Pregunta si quiere sobreescribir o no
                    s = "El registro '" & ts.codforma & "' (" & ts.NombreForma & ") ya existe en el destino." & vbCr & vbCr & _
                        "Desea sobreescribirlo?"
                    resp = frmMiMsgBox.MiMsgBox(s, "Banco")
                End If
                Select Case resp
                Case mmsgCancelar
                    mCancelado = True
                    Exit Do
                Case mmsgNo, mmsgNoTodo
                    GoTo siguiente
                End Select
            End If
            
            
            
            
            
            ts.codforma = .Fields("CodForma")
            ts.NombreForma = .Fields("NombreForma")
            If Not IsNull(.Fields("Plazo")) Then ts.Plazo = .Fields("Plazo")
            If Not IsNull(.Fields("CambiaFechaVenci")) Then ts.CambiaFechaVenci = .Fields("CambiaFechaVenci")
            If Not IsNull(.Fields("PermiteAbono")) Then ts.PermiteAbono = .Fields("PermiteAbono")
            If Not IsNull(.Fields("BandCobro")) Then ts.BandCobro = .Fields("BandCobro")
            If Not IsNull(.Fields("CodBanco")) Then ts.codBanco = .Fields("CodBanco")
            If Not IsNull(.Fields("CodTipoDoc")) Then ts.CodTipoDoc = .Fields("CodTipoDoc")
            If Not IsNull(.Fields("BandValida")) Then ts.BandValida = .Fields("BandValida")
            If Not IsNull(.Fields("ConsiderarComoEfectivo")) Then ts.ConsiderarComoEfectivo = .Fields("ConsiderarComoEfectivo")
            'jeaa 05/05/2008
            If Not IsNull(.Fields("IngresoAutomatico")) Then ts.IngresoAutomatico = .Fields("IngresoAutomatico")
            If Not IsNull(.Fields("CodProvCli")) Then ts.CodProvCli = .Fields("CodProvCli")
            If Not IsNull(.Fields("CodFormaTC")) Then ts.CodFormaTC = .Fields("CodFormaTC")
            If Not IsNull(.Fields("DeudaMismoCliente")) Then ts.DeudaMismoCliente = .Fields("DeudaMismoCliente")
            
            'ts.FechaGrabado = .Fields("FechaGrabado")
            ts.Grabar
            GrabarTSFormaCobroPago = GrabarTSFormaCobroPago + 1
siguiente:
            .MoveNext
        Loop
        
        .Close
    End With
    

salida:
    Set rs = Nothing
    Set ts = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarTSFormaCobroPago = GrabarTSFormaCobroPago * -1
    Exit Function

ErrTrap:
    If Not (ts Is Nothing) Then
        s = Err.Description & ": " & Err.Source & vbCr & ts.codforma & ", " & ts.NombreForma
    End If
    DispMsg "Importar datos de Forma de Cobro/Pago", "Error", s
    If MsgBox(s & vbCr & vbCr & _
                "Desea continuar con el siguiente registro?", _
                vbQuestion + vbYesNo) = vbYes Then
        Resume siguiente
    Else
        mCancelado = True
    End If
    GoTo salida
End Function
        

Private Function GrabarMotivo() As Long
    Dim sql As String, rs As Recordset, obj As Motivo, i As Long
    Dim s As String, resp As E_MiMsgBox
    
    On Error GoTo ErrTrap
    
    resp = mmsgSi
    
    'Abre el orígen
    sql = "SELECT * FROM Motivo ORDER BY CodMotivo"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    
    With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Motivos de Devolucion ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
        
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            
            'Primero busca si existe ya el mismo código
            Set obj = gobjMain.EmpresaActual.RecuperaMotivo(.Fields("CodMotivo"))
            If obj Is Nothing Then
                'Si no existe el código, crea nuevo
                Set obj = gobjMain.EmpresaActual.CreaMotivo
            Else
                'Si no se ha hecho pregunta, ó no ha contestado para Todo
                If (resp = mmsgSi) Or (resp = mmsgNo) Then
                    'Pregunta si quiere sobreescribir o no
                    s = "El registro '" & obj.CodMotivo & "' (" & _
                            obj.Descripcion & ") ya existe en el destino." & vbCr & vbCr & _
                        "Desea sobreescribirlo?"
                    resp = frmMiMsgBox.MiMsgBox(s, "Motivo")
                End If
                Select Case resp
                Case mmsgCancelar
                    mCancelado = True
                    Exit Do
                Case mmsgNo, mmsgNoTodo
                    GoTo siguiente
                End Select
            End If
            
            obj.CodMotivo = .Fields("CodMotivo")
            obj.Descripcion = .Fields("Descripcion")
            obj.BandValida = .Fields("BandValida")
            obj.Grabar
            GrabarMotivo = GrabarMotivo + 1
siguiente:
            .MoveNext
        Loop
        
        .Close
    End With
    

salida:
    Set rs = Nothing
    Set obj = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarMotivo = GrabarMotivo * -1
    Exit Function

ErrTrap:
    If Not (obj Is Nothing) Then
        s = Err.Description & ": " & Err.Source & vbCr & obj.CodMotivo & ", " & obj.Descripcion
    End If
    DispMsg "Importar datos de Motivo Devolucion", "Error", s
    If MsgBox(s & vbCr & vbCr & _
                "Desea continuar con el siguiente registro?", _
                vbQuestion + vbYesNo) = vbYes Then
        Resume siguiente
    Else
        mCancelado = True
    End If
    GoTo salida
End Function

'AUC 25/11/2005 para importar ivproveedodetalle
Private Sub GrabaIVProveedor(ByRef objItem As IVinventario, ByRef MSGERR As String)
    Dim objProv As IVinventario, objHijo As IVinventario
    Dim sql As String, rs As Recordset, i As Long
    Dim ix As Integer, objP As IVDetalleProveedor
    'Abre el orígen
    MSGERR = "" '' limpio los mensajes de error
    sql = "SELECT * FROM IVProveedorDetalle Where CodInventario = '" & objItem.CodInventario & "'"
    'sql = "SELECT * FROM IVMateria Where codMateria = '" & objItem.CodInventario & "'"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    'Primero borra lo anterior si es necesario
    For i = objItem.NumProveedorDetalle To 1 Step -1
        objItem.RemoveDetalleProveedor i
    Next i
    With rs
        Do Until .EOF
            'Agrega detalle PROVEEDOR
            ix = objItem.AddDetalleProveedor    'Aumenta  item  a la coleccion
            Set objP = objItem.RecuperaDetalleProveedor(ix)
            Set objHijo = gobjMain.EmpresaActual.RecuperaIVInventario(!CodInventario)
            If Not objHijo Is Nothing Then
                objP.CodInventario = !CodInventario
                objP.codProveedor = !codProveedor
            Else ''No encontro el hijo y sale del ciclo y borra el objpadre para que genere el error en la funcion que le esta llamando
                MSGERR = !CodInventario
                Set objItem = Nothing
                Exit Do
            End If
            .MoveNext
        Loop
    End With
End Sub

Private Function GrabarTipoCompra() As Long
    Dim sql As String, rs As Recordset, obj As IVTipoCompra, i As Long
    Dim s As String, resp As E_MiMsgBox
    
    On Error GoTo ErrTrap
    
    resp = mmsgSi
    
    'Abre el orígen
    sql = "SELECT * FROM IvTipoCompra ORDER BY CodTipoCompra "
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    
    With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Tipo de Compra ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
        
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            
            'Primero busca si existe ya el mismo código
            Set obj = gobjMain.EmpresaActual.RecuperaIVTipoCompra(.Fields("CodTipoCompra"))
            If obj Is Nothing Then
                'Si no existe el código, crea nuevo
                Set obj = gobjMain.EmpresaActual.CreaIVTipoCompra
            Else
                'Si no se ha hecho pregunta, ó no ha contestado para Todo
                If (resp = mmsgSi) Or (resp = mmsgNo) Then
                    'Pregunta si quiere sobreescribir o no
                    s = "El registro '" & obj.CodTipoCompra & "' (" & _
                            obj.Descripcion & ") ya existe en el destino." & vbCr & vbCr & _
                        "Desea sobreescribirlo?"
                    resp = frmMiMsgBox.MiMsgBox(s, "IVTipoCompra")
                End If
                Select Case resp
                Case mmsgCancelar
                    mCancelado = True
                    Exit Do
                Case mmsgNo, mmsgNoTodo
                    GoTo siguiente
                End Select
            End If
            
            obj.CodTipoCompra = .Fields("CodTipoCompra")
            obj.Descripcion = .Fields("Descripcion")
            obj.BandValida = .Fields("BandValida")
            obj.Grabar
            GrabarTipoCompra = GrabarTipoCompra + 1
siguiente:
            .MoveNext
        Loop
        
        .Close
    End With
    

salida:
    Set rs = Nothing
    Set obj = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarTipoCompra = GrabarTipoCompra * -1
    Exit Function

ErrTrap:
    If Not (obj Is Nothing) Then
        s = Err.Description & ": " & Err.Source & vbCr & obj.CodTipoCompra & ", " & obj.Descripcion
    End If
    DispMsg "Importar datos de Tipo de Compra", "Error", s
    If MsgBox(s & vbCr & vbCr & _
                "Desea continuar con el siguiente registro?", _
                vbQuestion + vbYesNo) = vbYes Then
        Resume siguiente
    Else
        mCancelado = True
    End If
    GoTo salida
End Function

Private Function GrabarIVUnidad() As Long
    Dim sql As String, rs As Recordset, obj As IVUnidad, i As Long
    Dim s As String, resp As E_MiMsgBox
    
    On Error GoTo ErrTrap
    
    resp = mmsgSi
    
    'Abre el orígen
    sql = "SELECT * FROM IVUnidad ORDER BY CodUnidad"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    
    With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Unidad ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
        
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            
            'Primero busca si existe ya el mismo código
            Set obj = gobjMain.EmpresaActual.RecuperaIVUnidad(.Fields("CodUnidad"))
            If obj Is Nothing Then
                'Si no existe el código, crea nuevo
                Set obj = gobjMain.EmpresaActual.CreaIVUnidad
            Else
                'Si no se ha hecho pregunta, ó no ha contestado para Todo
                If (resp = mmsgSi) Or (resp = mmsgNo) Then
                    'Pregunta si quiere sobreescribir o no
                    s = "El registro '" & obj.CodUnidad & "' (" & obj.Descripcion & ") ya existe en el destino." & vbCr & vbCr & _
                        "Desea sobreescribirlo?"
                    resp = frmMiMsgBox.MiMsgBox(s, "Unidad")
                End If
                Select Case resp
                Case mmsgCancelar
                    mCancelado = True
                    Exit Do
                Case mmsgNo, mmsgNoTodo
                    GoTo siguiente
                End Select
            End If
            
            obj.CodUnidad = .Fields("CodUnidad")
            obj.Descripcion = .Fields("Descripcion")
            obj.BandValida = .Fields("BandValida")
            obj.Grabar
            GrabarIVUnidad = GrabarIVUnidad + 1
siguiente:
            .MoveNext
        Loop
        
        .Close
    End With
    

salida:
    Set rs = Nothing
    Set obj = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarIVUnidad = GrabarIVUnidad * -1
    Exit Function

ErrTrap:
    If Not (obj Is Nothing) Then
        s = Err.Description & ": " & Err.Source & vbCr & obj.CodUnidad & ", " & obj.Descripcion
    End If
    DispMsg "Importar datos de Unidad ", "Error", s
    If MsgBox(s & vbCr & vbCr & _
                "Desea continuar con el siguiente registro?", _
                vbQuestion + vbYesNo) = vbYes Then
        Resume siguiente
    Else
        mCancelado = True
    End If
    GoTo salida
End Function


Private Sub BorraIVKardex(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset, ivk As IVKardex, i As Long
    Dim LoQueNoExiste As String
    On Error GoTo ErrTrap
   'Primero limpia
    gc.BorrarIVKardex
    'Abre el destino para agregar registro
    Set gc = Nothing
    Set ivk = Nothing
    Set rs = Nothing
    Exit Sub
ErrTrap:
    If Err.Number = -2147220960 Then
        Set gc = Nothing
        Set ivk = Nothing
        Set rs = Nothing
        Err.Raise Err.Number, "Importacion.PrepararIVKardex", "No existe " & LoQueNoExiste
        Exit Sub
    End If
    Set gc = Nothing
    Set ivk = Nothing
    Set rs = Nothing
    Err.Raise Err.Number, "Importacion", Err.Description
End Sub


Private Sub BorraPCKardex(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset, pck As PCKardex, i As Long
    Dim idAsignado As Long, LoQueNoExiste As String
   Dim v() As String
    On Error GoTo ErrTrap
   'Primero limpia
    gc.BorrarPCKardex
    Set gc = Nothing
    Set pck = Nothing
    Set rs = Nothing
    Exit Sub
ErrTrap:
    If Err.Number = -2147220960 Then Err.Raise Err.Number, "Importacion.PrepararPCKardex", "No existe " & LoQueNoExiste
End Sub


Private Function GrabarIVExist() As Long
    Dim sql As String, rs As Recordset, obj As IVUnidad, i As Long
    Dim s As String, resp As E_MiMsgBox
    Dim rsDest As Recordset, sqlDest As String
    Dim rsAux As Recordset
    On Error GoTo ErrTrap
    
    resp = mmsgSi
    
    'Abre el orígen
    sql = "SELECT * FROM IVExist ORDER BY CodInventario, CodBodega"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    
    With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Existencia ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
        
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            
            'Primero busca si existe ya el mismo código
            sqlDest = " select codinventario, codbodega, exist, existmin, existmax"
            sqlDest = sqlDest & " from ivexist ive"
            sqlDest = sqlDest & "    join ivinventario ivi on ive.idinventario=ivi.idinventario"
            sqlDest = sqlDest & " inner join ivbodega ivb on ivb.idbodega = ive.idbodega"
            sqlDest = sqlDest & " where codInventario='" & .Fields("CodInventario") & "' "
            sqlDest = sqlDest & " and codbodega='" & .Fields("CodBodega") & "'"
            'Set obj = gobjMain.EmpresaActual.RecuperaIVUnidad(.Fields("CodUnidad"))
            Set rsDest = gobjMain.EmpresaActual.OpenRecordset(sqlDest)
            
            If rsDest.RecordCount = 0 Then
                'Si no existe el código, crea nuevo
                'Set obj = gobjMain.EmpresaActual.CreaIVUnidad
                sql = "insert into ivexist (IdInventario, IdBodega, Exist, ExistMin, ExistMax)"
                sql = sql & " (select idinventario ,"
                sql = sql & " (select idbodega from IVBodegA where codBodega='" & .Fields("codBodega") & "'), "
                sql = sql & .Fields("exist") & "," & .Fields("existmin") & "," & .Fields("existMax")
                sql = sql & " from IVinventario where codInventario='" & .Fields("codInventario") & "')"
                
                Set rsAux = gobjMain.EmpresaActual.OpenRecordset(sql)
            Else
                'Si no se ha hecho pregunta, ó no ha contestado para Todo
                If (resp = mmsgSi) Or (resp = mmsgNo) Then
                    'Pregunta si quiere sobreescribir o no
                    's = "El registro '" & obj.CodUnidad & "' (" & obj.Descripcion & ") ya existe en el destino." & vbCr & vbCr & _
                        "Desea sobreescribirlo?"
                    'resp = frmMiMsgBox.MiMsgBox(s, "Unidad")
                    resp = mmsgSiTodo
                End If
                Select Case resp
                Case mmsgCancelar
                    mCancelado = True
                    Exit Do
                Case mmsgNo, mmsgNoTodo
                    GoTo siguiente
                End Select
            End If
            
                sql = " Update  ivexist "
                sql = sql & " set "
                sql = sql & " Exist=" & .Fields("exist") & ", "
                sql = sql & " ExistMin=" & .Fields("existmin") & ", "
                sql = sql & " ExistMax=" & .Fields("existmax")
                sql = sql & " from ivexist ive inner join ivinventario ivi on ive.idinventario=ivi.idinventario "
                sql = sql & " inner join ivbodega ivb on ivb.idbodega = ive.idbodega"
                sql = sql & " where codInventario='" & .Fields("codInventario") & "'"
                sql = sql & " and codbodega='" & .Fields("codBodega") & "'"

                Set rsAux = gobjMain.EmpresaActual.OpenRecordset(sql)
                GrabarIVExist = GrabarIVExist + 1
siguiente:
            .MoveNext
        Loop
        
        .Close
    End With
    

salida:
    Set rs = Nothing
    Set rsDest = Nothing
    Set rsAux = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarIVExist = GrabarIVExist * -1
    Exit Function

ErrTrap:
    If Not (obj Is Nothing) Then
        s = Err.Description & ": " & Err.Source & vbCr & rs.Fields("CodInventario") & ", " & rs.Fields("CodBodega")
    End If
    DispMsg "Importar datos de Existencia ", "Error", s
    If MsgBox(s & vbCr & vbCr & _
                "Desea continuar con el siguiente registro?", _
                vbQuestion + vbYesNo) = vbYes Then
        Resume siguiente
    Else
        mCancelado = True
    End If
    GoTo salida
End Function


Private Sub CambiaUsuariosenGNComprobante(CodTrans As String, numtrans As Long)
    Dim sql As String, rs As Recordset, id As Long, NumReg As Long
    Dim LoQueNoExiste As String
    Dim codUsuario As String, codUsuarioModifica As String, FechaGrabado As Date
    On Error GoTo ErrTrap
    'Abre el orígen para recuperar registro
    sql = "SELECT * FROM GNComprobante " & _
          "WHERE CodTrans = '" & CodTrans & "' AND NumTrans = " & numtrans
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly



        On Error Resume Next
        sql = " Update GNComprobante"
        sql = sql & " set CodUsuario='" & UCase(rs.Fields("CodUsuario")) & "', "
        sql = sql & " CodUsuarioModifica='" & UCase(rs.Fields("CodUsuarioModifica")) & "'"
        sql = sql & " where codtrans='" & CodTrans & "' and numtrans=" & numtrans
        
        gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
        
    Set rs = Nothing

    '--------Hasta aqui
    Exit Sub
ErrTrap:
    If Err.Number = -2147220960 Then
        Set rs = Nothing
        Err.Raise Err.Number, "Importacion.PrepararGNComprobante", "No existe " & LoQueNoExiste
        Exit Sub
    End If
    Set rs = Nothing
    Err.Raise Err.Number, "Importacion", Err.Description
End Sub

Private Sub CambiaFechaenGNComprobante(CodTrans As String, numtrans As Long)
    Dim sql As String, rs As Recordset, id As Long, NumReg As Long
    Dim LoQueNoExiste As String
    Dim codUsuario As String, codUsuarioModifica As String, FechaGrabado As Date
    On Error GoTo ErrTrap
    'Abre el orígen para recuperar registro
    sql = "SELECT * FROM GNComprobante " & _
          "WHERE CodTrans = '" & CodTrans & "' AND NumTrans = " & numtrans
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly



        On Error Resume Next
        
        sql = " Update GNComprobante"
        sql = sql & " set "
        sql = sql & " FechaGrabado = '" & rs.Fields("FechaGrabado") & "'"
        sql = sql & " where codtrans='" & CodTrans & "' and numtrans=" & numtrans
        
        gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
        
    Set rs = Nothing

    '--------Hasta aqui
    Exit Sub
ErrTrap:
    If Err.Number = -2147220960 Then
        Set rs = Nothing
        Err.Raise Err.Number, "Importacion.PrepararGNComprobante", "No existe " & LoQueNoExiste
        Exit Sub
    End If
    Set rs = Nothing
    Err.Raise Err.Number, "Importacion", Err.Description
End Sub


Private Sub CambiaFechaenTabla(campo As String, cod As String, tabla As String)
    Dim sql As String, rs As Recordset, id As Long, NumReg As Long
    On Error GoTo ErrTrap
    'Abre el orígen para recuperar registro
    sql = "SELECT * FROM  " & tabla
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly



        On Error Resume Next
        
        sql = " Update " & tabla
        sql = sql & " set "
        sql = sql & " FechaGrabado = '" & rs.Fields("FechaGrabado") & "'"
        sql = sql & " where " & campo & "='" & cod & "'"
        
        gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
        
    Set rs = Nothing

    '--------Hasta aqui
    Exit Sub
ErrTrap:
    If Err.Number = -2147220960 Then
        Set rs = Nothing
        Err.Raise Err.Number, "Importacion.PrepararGNComprobante", "No existe "
        Exit Sub
    End If
    Set rs = Nothing
    Err.Raise Err.Number, "Importacion", Err.Description
End Sub


'jeaa 04/01/2005
Private Function GrabarDesctoNumPagosIVGrupo() As Long
    Dim sql As String, rs As Recordset, obj As DesctoNumPagosxIVGrupo, i As Long
    Dim obj1 As DesctoPcGrupoxIVGrupo
    Dim s As String, resp As E_MiMsgBox
    On Error GoTo ErrTrap
    resp = mmsgSi
    'Abre el orígen
    sql = "SELECT * FROM DescNumPagIVG"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
   With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Descuento NumPagos por IVGrupo  ... " & _
                   i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            'Primero busca si existe ya el mismo código
            Set obj = gobjMain.EmpresaActual.RecuperaDesctNumPAgosxIV(.Fields("NumPagos") & "," & .Fields("CodIVGrupo"))
            If obj Is Nothing Then
               'Si no existe el código, crea nuevo
                Set obj = gobjMain.EmpresaActual.CreaDesctNumPagosxIV
           Else
                'Si no se ha hecho pregunta, ó no ha contestado para Todo
                If (resp = mmsgSi) Or (resp = mmsgNo) Then
                    'Pregunta si quiere sobreescribir o no
                    s = "El registro '" & obj.NumPagos & "' (" & obj.NumPagos & ") ya existe en el destino." & vbCr & vbCr & _
                       "Desea sobreescribirlo?"
                    resp = frmMiMsgBox.MiMsgBox(s, "DescNumPagIVG")
               End If
                Select Case resp
                Case mmsgCancelar
                    mCancelado = True
                    Exit Do
                Case mmsgNo, mmsgNoTodo
                    GoTo siguiente
                End Select
            End If
            obj.CodIVGrupo = .Fields("CodIVGrupo")
            obj.valor = .Fields("Valor")
            obj.NumPagos = .Fields("NumPagos")
            obj.BandOmiteRecDesc = .Fields("BandOmiteRecDesc")
            obj.Grabar
            GrabarDesctoNumPagosIVGrupo = GrabarDesctoNumPagosIVGrupo + 1
siguiente:
            .MoveNext
        Loop
        .Close
    End With
salida:
    Set rs = Nothing
    Set obj = Nothing
   'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarDesctoNumPagosIVGrupo = GrabarDesctoNumPagosIVGrupo * -1
   Exit Function
ErrTrap:
    If Not (obj Is Nothing) Then
        s = Err.Description & ": " & Err.Source & vbCr & obj.NumPagos & ", " & obj.CodIVGrupo
   End If
    DispMsg "Importar datos de Descuentos NumGrupo por PCGrupo ", "Error", s
   If MsgBox(s & vbCr & vbCr & _
                "Desea continuar con el siguiente registro?", _
                vbQuestion + vbYesNo) = vbYes Then
        Resume siguiente
    Else
        mCancelado = True
    End If
    GoTo salida
End Function


Private Sub CambiaEstadoGNComprobante(CodTrans As String, numtrans As Long, Estado As Integer)
    Dim sql As String, rs As Recordset, id As Long, NumReg As Long
    Dim LoQueNoExiste As String
    Dim codUsuario As String, codUsuarioModifica As String, FechaGrabado As Date
    On Error GoTo ErrTrap
    'Abre el orígen para recuperar registro
    sql = "SELECT * FROM GNComprobante " & _
          "WHERE CodTrans = '" & CodTrans & "' AND NumTrans = " & numtrans
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly



        On Error Resume Next
        
        sql = " Update GNComprobante"
        sql = sql & " set "
        sql = sql & " Estado = " & Estado
        sql = sql & " where codtrans='" & CodTrans & "' and numtrans=" & numtrans
        
        gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
        
    Set rs = Nothing

    '--------Hasta aqui
    Exit Sub
ErrTrap:
    If Err.Number = -2147220960 Then
        Set rs = Nothing
        Err.Raise Err.Number, "Importacion.PrepararGNComprobante", "No existe " & LoQueNoExiste
        Exit Sub
    End If
    Set rs = Nothing
    Err.Raise Err.Number, "Importacion", Err.Description
End Sub


Private Function GrabarPCHistorial() As Long
    Dim sql As String, rs As Recordset, obj As PCHistorial, i As Long
    Dim s As String, resp As E_MiMsgBox
    Dim mbooEsNuevo As Boolean
    On Error GoTo ErrTrap
    resp = mmsgSi
    'Abre el orígen
    sql = "SELECT * FROM PCHistorial ORDER BY FechaTrans"

    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando PCHistorial ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            'Primero busca si existe ya el mismo código
            Set obj = gobjMain.EmpresaActual.RecuperaPCHistorial(.Fields("Trans"))
            If obj Is Nothing Then
                'Si no existe el código, crea nuevo
                Set obj = gobjMain.EmpresaActual.CreaPCHistorial
                mbooEsNuevo = True
            Else
                    'Si no se ha hecho pregunta, ó no ha contestado para Todo
                    If (resp = mmsgSi) Or (resp = mmsgNo) Then
                        'Pregunta si quiere sobreescribir o no
                        s = "El registro  de PCHistorial : (" & obj.trans & ") ya existe en el destino." & vbCr & vbCr & _
                            "Desea sobreescribirlo?"
                        resp = frmMiMsgBox.MiMsgBox(s, "PCHistorial")
                    End If
                    Select Case resp
                    Case mmsgCancelar
                        mCancelado = True
                        Exit Do
                    Case mmsgNo, mmsgNoTodo
                        GoTo siguiente
                    End Select
                    End If
            obj.IdProvCli = .Fields("IdProvCli")
            obj.TransID = .Fields("TransId")
            obj.FechaTrans = .Fields("FechaTrans")
            If mbooEsNuevo = False Then
                If Val(obj.Estado) <> .Fields!Estado Then
                    If MsgBox("Estado de Trans. Origen " & IIf(.Fields!Estado = 2, "Devuelto", "NoDevuelto") & vbCr & _
                        "Estado de Trans. Destino " & IIf(obj.Estado = 2, "Devuelto", "NoDevuelto") & vbCr & vbCr & _
                        "Cliente: " & obj.trans & "...." & vbCr & _
                        "Desea Actualizar ? ", _
                        vbQuestion + vbYesNo) = vbYes Then
                        obj.Estado = .Fields("Estado")
                    End If
                End If
            Else
                obj.Estado = .Fields("Estado")
            End If
            obj.trans = .Fields("Trans")
            obj.Descripcion = .Fields("Descripcion")
            obj.FechaGrabado = .Fields("FechaGrabado")
            obj.valor = .Fields("Valor")
            obj.Grabar (mbooEsNuevo)
            'Para fechagrabado deja con la originarl
            Dim tabla As String, cod As String, campo As String
            campo = "trans"
            cod = .Fields("trans")
            tabla = "PcHistorial"
            CambiaFechaenTabla campo, cod, tabla
            GrabarPCHistorial = GrabarPCHistorial + 1
            mbooEsNuevo = False
siguiente:
            .MoveNext
        Loop
        .Close
    End With
salida:
    Set rs = Nothing
    Set obj = Nothing
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarPCHistorial = GrabarPCHistorial * -1
    Exit Function
ErrTrap:
    If Not (obj Is Nothing) Then
        s = Err.Description & ": " & Err.Source & vbCr & obj.trans & ", " & obj.Descripcion
    End If
    DispMsg "Importar datos de PCHistorial ", "Error", s
    If MsgBox(s & vbCr & vbCr & _
                "Desea continuar con el siguiente registro?", _
                vbQuestion + vbYesNo) = vbYes Then
        Resume siguiente
    Else
        mCancelado = True
    End If
    GoTo salida
End Function


Private Function GrabarPCGarante(ByVal Garante As Boolean) As Long
    Dim sql As String, rs As Recordset, obj As PCProvCli, i As Long
    Dim s As String, resp As E_MiMsgBox, j As Long, Desc As String
    Dim codigo As String, bandCambiaCodigo As Boolean
    Dim rs2 As Recordset
    
    On Error GoTo ErrTrap
    
    Desc = IIf(Garante, "Garante", "")
    resp = mmsgSi
    
    'Abre el orígen
    sql = "SELECT * FROM PCProvCli "
    sql = sql & "WHERE BandGarante<>0"
    sql = sql & " ORDER BY CodProvCli"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    
    With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando " & Desc & " ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
        
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            
            '***Angel. 22/dic/2003
            codigo = .Fields("CodProvCli")
            bandCambiaCodigo = False
Recupera_OtraVez:

            'Primero busca si existe ya el mismo código
            Set obj = gobjMain.EmpresaActual.RecuperaPCProvCli(codigo)
            If obj Is Nothing Then
                'Si no existe el código, crea nuevo
                Set obj = gobjMain.EmpresaActual.CreaPCProvCli
            Else
                If bandCambiaCodigo = False Then '***Angel. 22/dic/2003
                    'Si no se ha hecho pregunta, ó no ha contestado para Todo
                    If (resp = mmsgSi) Or (resp = mmsgNo) Then
                        'Pregunta si quiere sobreescribir o no
                        s = "El registro '" & obj.CodProvCli & "' (" & _
                            obj.nombre & ") ya existe en el destino." & vbCr & vbCr & _
                            "Desea sobreescribirlo?"
                        resp = frmMiMsgBox.MiMsgBox(s, IIf(Garante, "Garante", ""))
                    End If
                    Select Case resp
                    Case mmsgCancelar
                        mCancelado = True
                        Exit Do
                    Case mmsgNo, mmsgNoTodo
                        GoTo siguiente
                    End Select
                End If
            End If
            
            obj.CodProvCli = .Fields("CodProvCli")
            obj.nombre = .Fields("Nombre")
            
            If Not mPlantilla.BandIgnorarContabilidad Then     '*** MAKOTO 14/mar/01 Agregado
                If Len(.Fields("CodCuentaContable")) > 0 Then
                    obj.CodCuentaContable = .Fields("CodCuentaContable")
                End If
                If Len(.Fields("CodCuentaContable2")) > 0 Then
                    obj.CodCuentaContable2 = .Fields("CodCuentaContable2")
                End If
            End If
            
            obj.BandCliente = .Fields("BandCliente")
            obj.BandProveedor = .Fields("BandProveedor")
            If Not IsNull(.Fields("CodPostal")) Then obj.CodPostal = .Fields("CodPostal")
            If Not IsNull(.Fields("Direccion1")) Then obj.Direccion1 = .Fields("Direccion1")
            If Not IsNull(.Fields("Direccion2")) Then obj.Direccion2 = .Fields("Direccion2")
            If Not IsNull(.Fields("Ciudad")) Then obj.Ciudad = .Fields("Ciudad")
            If Not IsNull(.Fields("Provincia")) Then obj.Provincia = .Fields("Provincia")
            If Not IsNull(.Fields("Pais")) Then obj.Pais = .Fields("Pais")
            If Not IsNull(.Fields("Telefono1")) Then obj.Telefono1 = .Fields("Telefono1")
            If Not IsNull(.Fields("Telefono2")) Then obj.Telefono2 = .Fields("Telefono2")
            If Not IsNull(.Fields("Telefono3")) Then obj.Telefono3 = .Fields("Telefono3")
            If Not IsNull(.Fields("Fax")) Then obj.Fax = .Fields("Fax")
            If Not IsNull(.Fields("RUC")) Then obj.ruc = .Fields("RUC")
            If Not IsNull(.Fields("EMail")) Then obj.Email = .Fields("EMail")
            If Not IsNull(.Fields("Estado")) Then obj.Estado = .Fields("Estado")
            If Not IsNull(.Fields("CodVendedor")) Then obj.CodVendedor = .Fields("CodVendedor")
            If Not IsNull(.Fields("LimiteCredito")) Then obj.LimiteCredito = .Fields("LimiteCredito")
            
            If Not IsNull(.Fields("CodGrupo1")) Then obj.CodGrupo1 = .Fields("CodGrupo1")
            If Not IsNull(.Fields("CodGrupo2")) Then obj.CodGrupo2 = .Fields("CodGrupo2")
            If Not IsNull(.Fields("CodGrupo3")) Then obj.CodGrupo3 = .Fields("CodGrupo3")
            If Not IsNull(.Fields("CodGrupo4")) Then obj.CodGrupo4 = .Fields("CodGrupo4")
            
            If Not IsNull(.Fields("Banco")) Then obj.banco = .Fields("Banco")
            If Not IsNull(.Fields("NumCuenta")) Then obj.NumCuenta = .Fields("NumCuenta")
            If Not IsNull(.Fields("Swit")) Then obj.Swit = .Fields("Swit")
            If Not IsNull(.Fields("DirecBanco")) Then obj.DirecBanco = .Fields("DirecBanco")
            If Not IsNull(.Fields("TelBanco")) Then obj.TelBanco = .Fields("TelBanco")
            
            '***Agregado. 08/sep/2003. Angel
            '***Campos referentes a Anexos
            If Not IsNull(.Fields("TipoDocumento")) Then
                If Len(.Fields("TipoDocumento")) > 0 Then obj.TipoDocumento = .Fields("TipoDocumento")
                'jeaa 17/05/2007
                Select Case obj.TipoDocumento
                        Case "1", "01": obj.codtipoDocumento = "R"
                        Case "2", "02": obj.codtipoDocumento = "C"
                        Case "5", "05": obj.codtipoDocumento = "O"
                        Case "6", "06": obj.codtipoDocumento = "P"
                        Case "7", "07": obj.codtipoDocumento = "F"
                        Case Else: obj.codtipoDocumento = "T"
                End Select
            End If
            If Not IsNull(.Fields("TipoComprobante")) Then
                If Len(.Fields("TipoComprobante")) > 0 Then obj.TipoComprobante = .Fields("TipoComprobante")
            End If
            If Not IsNull(.Fields("NumAutSRI")) Then
                If Len(.Fields("NumAutSRI")) > 0 Then obj.NumAutSRI = .Fields("NumAutSRI")
            End If
            '***Agregado. 08/sep/2003. Angel
            '***Campos necesarios para tarjetas de descuentos
            If Not IsNull(.Fields("NombreAlterno")) Then obj.NombreAlterno = .Fields("NombreAlterno")
            If Not IsNull(.Fields("FechaNacimiento")) Then obj.FechaNacimiento = .Fields("FechaNacimiento")
            If Not IsNull(.Fields("FechaEntrega")) Then obj.FechaEntrega = .Fields("FechaEntrega")
            If Not IsNull(.Fields("FechaExpiracion")) Then obj.FechaExpiracion = .Fields("FechaExpiracion")
            If Not IsNull(.Fields("TotalDebe")) Then obj.TotalDebe = .Fields("TotalDebe")
            If Not IsNull(.Fields("TotalHaber")) Then obj.TotalHaber = .Fields("TotalHaber")
            If Not IsNull(.Fields("Observacion")) Then obj.Observacion = .Fields("observacion")
            If Not IsNull(.Fields("TipoProvCli")) Then obj.TipoProvCli = .Fields("TipoProvCli") 'jeaa 17/01/2008
            obj.BandEmpresaPublica = .Fields("BandEmpresaPublica")
            obj.BandGarante = .Fields("BandGarante")
            'Agrega los contactos
            AgregarContactos obj
            
            obj.Grabar
            
            Dim tabla As String, cod As String, campo As String
            campo = "CodProvCli"
            cod = .Fields("CodProvCli")
            tabla = "PcProvcli"
            CambiaFechaenTabla campo, cod, tabla
            
            
            GrabarPCGarante = GrabarPCGarante + 1
siguiente:
            .MoveNext
        Loop
        
        .Close
    End With

salida:
    Set rs = Nothing
    Set obj = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarPCGarante = GrabarPCGarante * -1
    Exit Function

ErrTrap:
    If Err.Number = NUMERROR_DUPLI Then '***Angel. 22/dic/2003
        Dim consulta As String, datos_destino As String
        s = rs.Fields("Nombre")
        consulta = "select * from pcprovcli where nombre='" & s & "'"
        codigo = BuscarxDesc_Nombre(consulta, "CodProvCli", "Nombre", "RUC", datos_destino)
        If Len(codigo) = 0 Then Resume presentar_error
        
        s = "Se ha detectado dos registros de " & Desc & " con nombres idénticos." & vbCrLf & _
            "Es posible que se haya modificado el código. " & vbCrLf & vbCrLf & _
            "Origen:  " & vbTab & Trim$(rs.Fields("CodProvCli")) & vbTab & Trim$(rs.Fields("Nombre")) & vbTab & Trim$(rs.Fields("RUC")) & vbCrLf & _
            "Destino: " & vbTab & datos_destino & vbCrLf & vbCrLf & _
            "¿Desea sobreescribir el código en el destino?"
            
        If MsgBox(s, vbYesNo + vbQuestion) = vbYes Then
            bandCambiaCodigo = True
            Resume Recupera_OtraVez
        Else
            Resume siguiente
        End If
    Else
presentar_error:
        If Not (obj Is Nothing) Then
            s = "Error: " & Err.Description & vbCrLf & _
                "Origen: " & Err.Source & vbCrLf & vbCrLf & _
                "Registro Afectado:" & vbCrLf & _
                "   Código: " & obj.CodProvCli & vbCrLf & _
                "   Nombre: " & obj.nombre
        End If
        DispMsg "Importar datos de " & Desc, "Error", s
        If MsgBox(s & vbCr & vbCr & _
                    "Desea continuar con el siguiente registro?", _
                    vbQuestion + vbYesNo) = vbYes Then
            Resume siguiente
        Else
            mCancelado = True
        End If
        GoTo salida
    End If
End Function


Private Function GrabarIVBanco() As Long
    Dim sql As String, rs As Recordset, obj As IVBanco, i As Long
    Dim s As String, resp As E_MiMsgBox
    
    On Error GoTo ErrTrap
    
    resp = mmsgSi
    
    'Abre el orígen
    sql = "SELECT * FROM IvBanco ORDER BY CodBanco"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    
    With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando IvBancos ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
        
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            
            'Primero busca si existe ya el mismo código
            Set obj = gobjMain.EmpresaActual.RecuperaIVBanco(.Fields("CodBanco"))
            If obj Is Nothing Then
                'Si no existe el código, crea nuevo
                Set obj = gobjMain.EmpresaActual.CreaIVBanco
            Else
                'Si no se ha hecho pregunta, ó no ha contestado para Todo
                If (resp = mmsgSi) Or (resp = mmsgNo) Then
                    'Pregunta si quiere sobreescribir o no
                    s = "El registro '" & obj.codBanco & "' (" & _
                            obj.Descripcion & ") ya existe en el destino." & vbCr & vbCr & _
                        "Desea sobreescribirlo?"
                    resp = frmMiMsgBox.MiMsgBox(s, "IVBanco")
                End If
                Select Case resp
                Case mmsgCancelar
                    mCancelado = True
                    Exit Do
                Case mmsgNo, mmsgNoTodo
                    GoTo siguiente
                End Select
            End If
            
            obj.codBanco = .Fields("CodBanco")
            If Not IsNull(.Fields("CodForma")) Then obj.codforma = .Fields("CodForma") 'jeaa 17/01/2008
            obj.Descripcion = .Fields("Descripcion")
            If Not IsNull(.Fields("CodCliente")) Then obj.codcliente = .Fields("CodCliente") 'jeaa 17/01/2008
            obj.BandValida = .Fields("BandValida")
            obj.Grabar
            GrabarIVBanco = GrabarIVBanco + 1
siguiente:
            .MoveNext
        Loop
        
        .Close
    End With
    

salida:
    Set rs = Nothing
    Set obj = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarIVBanco = GrabarIVBanco * -1
    Exit Function

ErrTrap:
    If Not (obj Is Nothing) Then
        s = Err.Description & ": " & Err.Source & vbCr & obj.codBanco & ", " & obj.Descripcion
    End If
    DispMsg "Importar datos de IVBanco", "Error", s
    If MsgBox(s & vbCr & vbCr & _
                "Desea continuar con el siguiente registro?", _
                vbQuestion + vbYesNo) = vbYes Then
        Resume siguiente
    Else
        mCancelado = True
    End If
    GoTo salida
End Function

Private Function GrabarIVTarjeta() As Long
    Dim sql As String, rs As Recordset, obj As IVTarjeta, i As Long
    Dim s As String, resp As E_MiMsgBox
    
    On Error GoTo ErrTrap
    
    resp = mmsgSi
    
    'Abre el orígen
    sql = "SELECT * FROM IvTarjeta ORDER BY CodTarjeta"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    
    With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando IvTarjetas ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
        
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            
            'Primero busca si existe ya el mismo código
            Set obj = gobjMain.EmpresaActual.RecuperaIVTarjeta(.Fields("CodTarjeta"))
            If obj Is Nothing Then
                'Si no existe el código, crea nuevo
                Set obj = gobjMain.EmpresaActual.CreaIVTarjeta
            Else
                'Si no se ha hecho pregunta, ó no ha contestado para Todo
                If (resp = mmsgSi) Or (resp = mmsgNo) Then
                    'Pregunta si quiere sobreescribir o no
                    s = "El registro '" & obj.CodTarjeta & "' (" & _
                            obj.Descripcion & ") ya existe en el destino." & vbCr & vbCr & _
                        "Desea sobreescribirlo?"
                    resp = frmMiMsgBox.MiMsgBox(s, "IVTarjeta")
                End If
                Select Case resp
                Case mmsgCancelar
                    mCancelado = True
                    Exit Do
                Case mmsgNo, mmsgNoTodo
                    GoTo siguiente
                End Select
            End If
            
            obj.CodTarjeta = .Fields("CodTarjeta")
            If Not IsNull(.Fields("CodForma")) Then obj.codforma = .Fields("CodForma") 'jeaa 17/01/2008
            obj.Descripcion = .Fields("Descripcion")
            obj.BandValida = .Fields("BandValida")
            obj.Grabar
            GrabarIVTarjeta = GrabarIVTarjeta + 1
siguiente:
            .MoveNext
        Loop
        
        .Close
    End With
    

salida:
    Set rs = Nothing
    Set obj = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarIVTarjeta = GrabarIVTarjeta * -1
    Exit Function

ErrTrap:
    If Not (obj Is Nothing) Then
        s = Err.Description & ": " & Err.Source & vbCr & obj.CodTarjeta & ", " & obj.Descripcion
    End If
    DispMsg "Importar datos de IVTarjeta", "Error", s
    If MsgBox(s & vbCr & vbCr & _
                "Desea continuar con el siguiente registro?", _
                vbQuestion + vbYesNo) = vbYes Then
        Resume siguiente
    Else
        mCancelado = True
    End If
    GoTo salida
End Function



Private Sub PrepararPRLibroDetalle(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset, ctd As PRLibroDetalle, i As Long
    Dim LoQueNoExiste As String
    On Error GoTo ErrTrap
   'Primero limpia
    gc.BorrarPRLibroDetalle
    'Abre el destino para agregar registro
    sql = "SELECT * FROM PRLibroDetalle " & _
          "WHERE CodTrans = '" & gc.CodTrans & "' AND NumTrans = " & gc.numtrans & _
          " ORDER BY Orden"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    Do Until rs.EOF
        DoEvents
        i = gc.AddPRLibroDetalle
        Set ctd = gc.PRLibroDetalle(i)
        With ctd
            LoQueNoExiste = "Código de Cuenta: " & rs.Fields("CodCuenta")
           .codcuenta = rs.Fields("CodCuenta")
            If Not IsNull(rs.Fields("Descripcion")) Then .Descripcion = rs.Fields("Descripcion")
            .Debe = rs.Fields("Debe")
            .Haber = rs.Fields("Haber")
            .BandIntegridad = rs.Fields("BandIntegridad")
            .FechaEjec = CDate("01/" & DatePart("m", rs.Fields("FechaEjec")) & "/" & DatePart("yyyy", rs.Fields("FechaEjec")))
            .Orden = rs.Fields("Orden")
        End With
        rs.MoveNext
    Loop
    rs.Close
    Set gc = Nothing
    Set ctd = Nothing
    Set rs = Nothing
            Exit Sub
ErrTrap:
    If Err.Number = -2147220960 Then Err.Raise Err.Number, "Importacion.PrepararPCKardex", "No existe " & LoQueNoExiste
End Sub


''Private Sub PrepararAFKardex(ByVal gc As GNComprobante)
''    Dim sql As String, rs As Recordset, ivk As AFKardex, i As Long
''    Dim LoQueNoExiste As String
''    On Error GoTo Errtrap
''   'Primero limpia
''    gc.BorrarAFKardex
''    'Abre el destino para agregar registro
''    sql = "SELECT * FROM AFKardex " & _
''          "WHERE CodTrans = '" & gc.CodTrans & "' AND NumTrans = " & gc.numtrans & _
''          " ORDER BY Orden"
''    Set rs = New Recordset
''    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
''    Do Until rs.EOF
''        DoEvents
''        i = gc.AddAFKardex
''        Set ivk = gc.AFKardex(i)
''        With ivk
''            LoQueNoExiste = "Código de Item: " & rs.Fields("CodInventario")
''            .CodInventario = rs.Fields("CodInventario")
''            LoQueNoExiste = "Código de Bodega: " & rs.Fields("CodBodega")
''            .CodBodega = rs.Fields("CodBodega")
''            LoQueNoExiste = "?"
''           .cantidad = rs.Fields("Cantidad")
''            .CostoTotal = rs.Fields("CostoTotal")
''            .CostoRealTotal = rs.Fields("CostoRealTotal")
''            .PrecioTotal = rs.Fields("PrecioTotal")
''            .PrecioRealTotal = rs.Fields("PrecioRealTotal")
''            If Not IsNull(rs.Fields("Descuento")) Then .Descuento = rs.Fields("Descuento")
''            If Not IsNull(rs.Fields("IVA")) Then .IVA = rs.Fields("IVA")
''            .orden = rs.Fields("Orden")
''            If Not IsNull(rs.Fields("Nota")) Then .Nota = rs.Fields("Nota")
''
''            On Error Resume Next
''            LoQueNoExiste = "Numero Precio: " & rs.Fields("NumeroPrecio")
''            If Err.Number <> 3265 Then
''                '***Agregado. 11/sep/2003. Angel
''                If Not IsNull(rs.Fields("NumeroPrecio")) Then .NumeroPrecio = rs.Fields("NumeroPrecio")
''            End If
''
''
''            'jeaa 22/09/2005
''            Err.Clear
''            LoQueNoExiste = "TiempoEntrega: " & rs.Fields("TiempoEntrega")
''            If Err.Number <> 3265 Then
''                '***Agregado. 05/ago/2004. Angel
''                If Not IsNull(rs.Fields("TiempoEntrega")) Then .TiempoEntrega = rs.Fields("TiempoEntrega")
''            End If
''            Err.Clear
''
''
''            On Error GoTo Errtrap
''
''        End With
''        rs.MoveNext
''    Loop
''    rs.Close
''
''    Set gc = Nothing
''    Set ivk = Nothing
''    Set rs = Nothing
''    Exit Sub
''Errtrap:
''    If Err.Number = -2147220960 Then
''        Set gc = Nothing
''        Set ivk = Nothing
''        Set rs = Nothing
''        Err.Raise Err.Number, "Importacion.PrepararAFKardex", "No existe " & LoQueNoExiste
''        Exit Sub
''    End If
''    Set gc = Nothing
''    Set ivk = Nothing
''    Set rs = Nothing
''    Err.Raise Err.Number, "Importacion", Err.Description
''End Sub
''
''
''
''Private Sub PrepararAFKardexRecargo(ByVal gc As GNComprobante)
''    Dim sql As String, rs As Recordset, ivkr As AFKardexRecargo, i As Long
''    Dim LoQueNoExiste As String
''    On Error GoTo Errtrap
''   'Primero limpia
''    gc.BorrarAFKardexRecargo
''    'Abre el destino para agregar registro
''    sql = "SELECT * FROM afKardexRecargo " & _
''          "WHERE CodTrans = '" & gc.CodTrans & "' AND NumTrans = " & gc.numtrans & _
''          " ORDER BY Orden"
''    Set rs = New Recordset
''    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
''    Do Until rs.EOF
''        DoEvents
''        i = gc.AddAFKardexRecargo
''        Set ivkr = gc.AFKardexRecargo(i)
''        With ivkr
''            LoQueNoExiste = "Código de Recargo: " & rs.Fields("CodRecargo")
''           .codRecargo = rs.Fields("CodRecargo")
''            .porcentaje = rs.Fields("Porcentaje")
''            .valor = rs.Fields("Valor")
''            .BandModificable = rs.Fields("BandModificable")
''            .BandOrigen = rs.Fields("BandOrigen")
''            .BandProrrateado = rs.Fields("BandProrrateado")
''            .AfectaIvaItem = rs.Fields("AfectaIvaItem")
''            .orden = rs.Fields("Orden")
''        End With
''        rs.MoveNext
''    Loop
''    rs.Close
''    Set gc = Nothing
''    Set ivkr = Nothing
''    Set rs = Nothing
''    Exit Sub
''Errtrap:
''    If Err.Number = -2147220960 Then Err.Raise Err.Number, "Importacion.PrepararPCKardex", "No existe " & LoQueNoExiste
''End Sub
''
''

Private Function GrabarPCDiasCredito() As Long
    Dim sql As String, rs As Recordset, obj As PCDiasCredito, i As Long
    Dim s As String, resp As E_MiMsgBox
    Dim codigo As String, bandCambiaCodigo As Boolean
    
    On Error GoTo ErrTrap
    
    resp = mmsgSi
    
    'Abre el orígen
    sql = "SELECT * FROM PCDiasCredito ORDER BY CodDiasCredito "
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        rs.MoveLast
        rs.MoveFirst
    End If
    
    With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando DiasCredito " & _
            " ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
        
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            
            '***Angel. 23/dic/2003
            codigo = .Fields("CodDiasCredito")
            bandCambiaCodigo = False
Recupera_OtraVez:

            'Primero busca si existe ya el mismo código
            Set obj = gobjMain.EmpresaActual.RecuperaPCDiasCredito(codigo)
            If obj Is Nothing Then
                'Si no existe el código, crea nuevo
                Set obj = gobjMain.EmpresaActual.CreaPCDiasCredito
            Else
                If bandCambiaCodigo = False Then
                    'Si no se ha hecho pregunta, ó no ha contestado para Todo
                    If (resp = mmsgSi) Or (resp = mmsgNo) Then
                        'Pregunta si quiere sobreescribir o no
                        s = "El registro '" & obj.CodDiasCredito & "' (" & obj.Descripcion & ") ya existe en el destino." & vbCr & vbCr & _
                            "Desea sobreescribirlo?"
                        resp = frmMiMsgBox.MiMsgBox(s, "DiasCredito")
                    End If
                    Select Case resp
                    Case mmsgCancelar
                        mCancelado = True
                        Exit Do
                    Case mmsgNo, mmsgNoTodo
                        GoTo siguiente
                    End Select
                End If
            End If
            
            obj.CodDiasCredito = .Fields("CodDiasCredito")
            obj.Descripcion = .Fields("Descripcion")
            obj.BandValida = .Fields("BandValida")
            obj.Grabar
            
            Dim tabla As String, cod As String, campo As String
            campo = "CodDiasCredito"
            cod = .Fields("CodDiasCredito")
            tabla = "PCDiasCredito"
            CambiaFechaenTabla campo, cod, tabla
            
            
            GrabarPCDiasCredito = GrabarPCDiasCredito + 1
siguiente:
            .MoveNext
        Loop
        
        .Close
    End With
    

salida:
    Set rs = Nothing
    Set obj = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarPCDiasCredito = GrabarPCDiasCredito * -1
    Exit Function

ErrTrap:
    If Err.Number = NUMERROR_DUPLI Then '***Angel. 23/dic/2003
        Dim consulta As String, datos_destino As String
        s = rs.Fields("Descripcion")
        consulta = "SELECT * FROM PCDiasCredito WHERE Descripcion='" & s & "'"
        codigo = BuscarxDesc_Nombre(consulta, "CodDiasCredito", "Descripcion", "", datos_destino)
        If Len(codigo) = 0 Then Resume presentar_error
        
        s = "Se ha detectado dos registros de DiasCredito " & _
            " con descripciones idénticas." & vbCrLf & _
            "Es posible que se haya modificado el código. " & vbCrLf & vbCrLf & _
            "Origen:  " & vbTab & Trim$(rs.Fields("CodDiasCredito")) & vbTab & Trim$(rs.Fields("Descripcion")) & vbCrLf & _
            "Destino: " & vbTab & datos_destino & vbCrLf & vbCrLf & _
            "¿Desea sobreescribir el código en el destino?"
            
        If MsgBox(s, vbYesNo + vbQuestion) = vbYes Then
            bandCambiaCodigo = True
            Resume Recupera_OtraVez
        Else
            Resume siguiente
        End If
    Else
presentar_error:
        If Not (obj Is Nothing) Then
            s = Err.Description & ": " & Err.Source & vbCr & obj.CodDiasCredito & ", " & obj.Descripcion
        End If
        DispMsg "Importar datos de DiasCredito" & _
                 "Error", s
        If MsgBox(s & vbCr & vbCr & _
                    "Desea continuar con el siguiente registro?", _
                    vbQuestion + vbYesNo) = vbYes Then
            Resume siguiente
        Else
            mCancelado = True
        End If
        GoTo salida
    End If
End Function


Private Function GrabarPlazoIVGrupoPCGrupo() As Long
    Dim sql As String, rs As Recordset, obj As PlazoPcGrupoxIVGrupo, i As Long
    Dim s As String, resp As E_MiMsgBox
    On Error GoTo ErrTrap
    resp = mmsgSi
    'Abre el orígen
    sql = "SELECT * FROM PlazoIVGPCG "
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
   With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Plazo IVGrupo por PCGrupo ... " & _
                   i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            'Primero busca si existe ya el mismo código
            Set obj = gobjMain.EmpresaActual.RecuperaPlazoPCxIV(.Fields("CodPCGrupo") & "," & .Fields("CodIVGrupo"))
            If obj Is Nothing Then
               'Si no existe el código, crea nuevo
                Set obj = gobjMain.EmpresaActual.CreaPlazoPCxIV
           Else
                'Si no se ha hecho pregunta, ó no ha contestado para Todo
                If (resp = mmsgSi) Or (resp = mmsgNo) Then
                    'Pregunta si quiere sobreescribir o no
                    s = "El registro '" & obj.CodPCGrupo & "' (" & obj.CodPCGrupo & ") ya existe en el destino." & vbCr & vbCr & _
                       "Desea sobreescribirlo?"
                    resp = frmMiMsgBox.MiMsgBox(s, "PlazoIVGPCG")
               End If
                Select Case resp
                Case mmsgCancelar
                    mCancelado = True
                    Exit Do
                Case mmsgNo, mmsgNoTodo
                    GoTo siguiente
                End Select
            End If
            obj.CodPCGrupo = .Fields("CodPCGrupo")
            obj.CodIVGrupo = .Fields("CodIVGrupo")
            obj.valor = .Fields("Valor")
            obj.Grabar
            GrabarPlazoIVGrupoPCGrupo = GrabarPlazoIVGrupoPCGrupo + 1
siguiente:
            .MoveNext
        Loop
        .Close
    End With
salida:
    Set rs = Nothing
    Set obj = Nothing
   'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarPlazoIVGrupoPCGrupo = GrabarPlazoIVGrupoPCGrupo * -1
   Exit Function
ErrTrap:
    If Not (obj Is Nothing) Then
        s = Err.Description & ": " & Err.Source & vbCr & obj.CodPCGrupo & ", " & obj.CodIVGrupo
   End If
    DispMsg "Importar datos de Plazo IVGrupo por PCGrupo ", "Error", s
   If MsgBox(s & vbCr & vbCr & _
                "Desea continuar con el siguiente registro?", _
                vbQuestion + vbYesNo) = vbYes Then
        Resume siguiente
    Else
        mCancelado = True
    End If
    GoTo salida
End Function


Private Sub PrepararPCKardexCHP(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset, pck As PCKardexCHP, i As Long
    Dim idAsignado As Long, LoQueNoExiste As String
   Dim v() As String
    On Error GoTo ErrTrap
   'Primero limpia
    gc.BorrarPCKardexCHP
    'Abre el destino para agregar registro
    sql = "SELECT * FROM PCKardexCHP " & _
          "WHERE CodTrans = '" & gc.CodTrans & "' AND NumTrans = " & gc.numtrans & _
          " ORDER BY Orden"
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    Do Until rs.EOF
        DoEvents
        i = gc.AddPCKardexCHP
        Set pck = gc.PCKardexCHP(i)
        With pck
            'Desactiva la verificación de saldo de doc.asignado
            'Para que no genere error cuando asigna valor de Debe/Haber
            .BandNoVerificarSaldo = True            '*** MAKOTO 22/mar/01 Agregado
            'en esta sección pide el valor de cliente, no lo encuentra y emite error .....  es aquí en donde hay que controlar
            LoQueNoExiste = "Código de Cliente/Proveedor: " & rs.Fields("CodProvCli")
           .CodProvCli = rs.Fields("CodProvCli")
           
           '***Angel. 13/nov/2003. Para que se importe sin necesidad de IdAsignado
           If Not (mPlantilla.BandIgnorarDocAsignado) Then
                If Len(rs.Fields("GuidAsignadoPCK")) > 0 Then      '*** MAKOTO 16/mar/01
                    '***Angel. 13/nov/2003. Agregado Mensaje
                    LoQueNoExiste = "PCKardexCHP: " & _
                                    "No se encuentra el documento asignado. " & vbCr & _
                                    "(" & gc.CodTrans & _
                                    gc.numtrans & ") " & vbCr & rs.Fields("GuidAsignadoPCK")
                    If gc.GNTrans.CodPantalla = "TSICHP" Then
                        .SetIdAsignadoPCKPorGuidCHP rs.Fields("GuidAsignadoPCK")
                    Else
                        .SetIdAsignadoPCKPorGuid rs.Fields("GuidAsignadoPCK")
                    End If
                End If
           End If
           
            LoQueNoExiste = "Código de Forma de Pago: " & rs.Fields("CodForma")
            .codforma = rs.Fields("CodForma")
            If Not IsNull(rs.Fields("NumLetra")) Then .NumLetra = rs.Fields("NumLetra")
            .Debe = rs.Fields("Debe")
            .Haber = rs.Fields("Haber")
            .FechaEmision = rs.Fields("FechaEmision")
            .FechaVenci = rs.Fields("FechaVenci")
            If Not IsNull(rs.Fields("Observacion")) Then .Observacion = rs.Fields("Observacion")
            .Orden = rs.Fields("Orden")
            .Guid = rs.Fields("guid")
            LoQueNoExiste = "Código de Tarjeta " & rs.Fields("CodTarjeta")
            If Not IsNull(rs.Fields("CodTarjeta")) Then .CodTarjeta = rs.Fields("CodTarjeta")
            LoQueNoExiste = "Código de Banco " & rs.Fields("CodBanco")
            If Not IsNull(rs.Fields("CodBanco")) Then .codBanco = rs.Fields("CodBanco")
            If Not IsNull(rs.Fields("NumCuenta")) Then .NumCuenta = rs.Fields("NumCuenta")
            If Not IsNull(rs.Fields("NumCheque")) Then .Numcheque = rs.Fields("NumCheque")
            If Not IsNull(rs.Fields("TitularCta")) Then .TitularCta = rs.Fields("TitularCta")

            .SetIdFromGuid
        End With
        rs.MoveNext
    Loop
    rs.Close
    Set gc = Nothing
    Set pck = Nothing
    Set rs = Nothing
    Exit Sub
ErrTrap:
    If Err.Number = -2147220960 Then Err.Raise Err.Number, "Importacion.PrepararPCKardexCHP", "No existe " & LoQueNoExiste
End Sub


Private Sub BorraPCKardexCHP(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset, pckCHP As PCKardexCHP, i As Long
    Dim idAsignado As Long, LoQueNoExiste As String
   Dim v() As String
    On Error GoTo ErrTrap
   'Primero limpia
    gc.BorrarPCKardexCHP
    Set gc = Nothing
    Set pckCHP = Nothing
    Set rs = Nothing
    Exit Sub
ErrTrap:
    If Err.Number = -2147220960 Then Err.Raise Err.Number, "Importacion.PrepararPCKardexCHP", "No existe " & LoQueNoExiste
End Sub

Private Function GrabarPCParroquia() As Long
    Dim sql As String, rs As Recordset, obj As PCParroquia, i As Long
    Dim s As String, resp As E_MiMsgBox
    Dim codigo As String, bandCambiaCodigo As Boolean
    
    On Error GoTo ErrTrap
    
    resp = mmsgSi
    
    'Abre el orígen
    sql = "SELECT * FROM PCParroquia ORDER BY CodParroquia "
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        rs.MoveLast
        rs.MoveFirst
    End If
    
    With rs
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Parroquia " & _
            " ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
        
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            
            '***Angel. 23/dic/2003
            codigo = .Fields("CodParroquia")
            bandCambiaCodigo = False
Recupera_OtraVez:

            'Primero busca si existe ya el mismo código
            Set obj = gobjMain.EmpresaActual.RecuperaPCParroquia(codigo)
            If obj Is Nothing Then
                'Si no existe el código, crea nuevo
                Set obj = gobjMain.EmpresaActual.CreaPCParroquia
            Else
                If bandCambiaCodigo = False Then
                    'Si no se ha hecho pregunta, ó no ha contestado para Todo
                    If (resp = mmsgSi) Or (resp = mmsgNo) Then
                        'Pregunta si quiere sobreescribir o no
                        s = "El registro '" & obj.codParroquia & "' (" & obj.Descripcion & ") ya existe en el destino." & vbCr & vbCr & _
                            "Desea sobreescribirlo?"
                        resp = frmMiMsgBox.MiMsgBox(s, "Parroquia")
                    End If
                    Select Case resp
                    Case mmsgCancelar
                        mCancelado = True
                        Exit Do
                    Case mmsgNo, mmsgNoTodo
                        GoTo siguiente
                    End Select
                End If
            End If
            
            obj.codParroquia = .Fields("CodParroquia")
            obj.Descripcion = .Fields("Descripcion")
            obj.codCanton = .Fields("codCanton")
            obj.BandValida = .Fields("BandValida")
            obj.Grabar
            
            Dim tabla As String, cod As String, campo As String
            campo = "CodParroquia"
            cod = .Fields("CodParroquia")
            tabla = "PCParroquia"
            CambiaFechaenTabla campo, cod, tabla
            
            
            GrabarPCParroquia = GrabarPCParroquia + 1
siguiente:
            .MoveNext
        Loop
        
        .Close
    End With
    

salida:
    Set rs = Nothing
    Set obj = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarPCParroquia = GrabarPCParroquia * -1
    Exit Function

ErrTrap:
'    If Err.Number = NUMERROR_DUPLI Then '***Angel. 23/dic/2003
'        Dim consulta As String, datos_destino As String
'        s = rs.Fields("Descripcion")
'        consulta = "SELECT * FROM PCParroquia WHERE Descripcion='" & s & "'"
'        codigo = BuscarxDesc_Nombre(consulta, "CodParroquia", "Descripcion", "", datos_destino)
'        If Len(codigo) = 0 Then Resume presentar_error
'
'        s = "Se ha detectado dos registros de Parroquia " & _
'            " con descripciones idénticas." & vbCrLf & _
'            "Es posible que se haya modificado el código. " & vbCrLf & vbCrLf & _
'            "Origen:  " & vbTab & Trim$(rs.Fields("CodParroquia")) & vbTab & Trim$(rs.Fields("Descripcion")) & vbCrLf & _
'            "Destino: " & vbTab & datos_destino & vbCrLf & vbCrLf & _
'            "¿Desea sobreescribir el código en el destino?"
            
'        If MsgBox(s, vbYesNo + vbQuestion) = vbYes Then
'            bandCambiaCodigo = True
'            Resume Recupera_OtraVez
'        Else
'            Resume siguiente
'        End If
''    Else
'presentar_error:
'        If Not (obj Is Nothing) Then
'            s = Err.Description & ": " & Err.Source & vbCr & obj.codParroquia & ", " & obj.Descripcion
'        End If
'        DispMsg "Importar datos de Parroquia" & _
'                 "Error", s
'        If MsgBox(s & vbCr & vbCr & _
'                    "Desea continuar con el siguiente registro?", _
'                    vbQuestion + vbYesNo) = vbYes Then
'            Resume siguiente
'        Else
'            mCancelado = True
'        End If
'        GoTo salida
'    End If
End Function

Private Sub PrepararGNOferta( _
                ByVal gc As GNComprobante, _
                ByRef Estado As Byte)
    Dim sql As String, rs As Recordset, id As Long
    Dim LoQueNoExiste As String
    On Error GoTo ErrTrap
    'Abre el orígen para recuperar registro
    sql = "SELECT * FROM GNOferta" & _
          " WHERE CodTrans = '" & gc.CodTrans & "' AND NumTrans = " & gc.numtrans
    Set rs = New Recordset
    rs.Open sql, mcnOrigen, adOpenStatic, adLockReadOnly

    With gc

        .Atencion = rs.Fields("Atencion")
        .FormaPago = rs.Fields("FormaPago")
        .TiempoEntrega = rs.Fields("TiempoEntrega")
        .Validez = rs.Fields("Validez")
        
        .Detalles = rs.Fields("Detalles")
        .FechaValidez = rs.Fields("FechaValidez")
        .Observaciones = rs.Fields("Observaciones")
        .FechaEntrega = rs.Fields("FechaEntrega")
        .TiempoEstimadoEntrega = rs.Fields("TiempoEstimadoEntrega")
        LoQueNoExiste = "Código de CodGarante2: " & rs.Fields("CodGarante2")
        .CodGaranteRef2 = rs.Fields("CodGarante2")
        LoQueNoExiste = "Código de Inventario: " & rs.Fields("CodInventario")
        If Not IsNull(rs.Fields("CodInventario")) Then .CodInventario = rs.Fields("CodInventario")
        LoQueNoExiste = "Código de  Empleado: " & rs.Fields("CodEmpleadoRef")
        If Not IsNull(rs.Fields("CodEmpleadoRef")) Then .CodEmpleadoRef = rs.Fields("CodEmpleadoRef")
        .NumDireccion = rs.Fields("NumDireccion")
        .DirTransporte = rs.Fields("DireccionTransporte")
        .Opcion = rs.Fields("Opcion")
        'LoQueNoExiste = "Código de PCAgencia: " & rs.Fields("CodAgenciaref")
        'If Not IsNull(rs.Fields("CodAgencia")) Then .Coda= rs.Fields("CodEmpleadoRef")
       
              
        
        LoQueNoExiste = "Código de Cobrador: " & rs.Fields("CodCobrador")
        If Not IsNull(rs.Fields("CodCobrador")) Then .CodCobrador = rs.Fields("CodCobrador")
        
        LoQueNoExiste = "Código de Agencia Curier: " & rs.Fields("CodAgenciaCurier")
        If Not IsNull(rs.Fields("CodAgenciaCurier")) Then .CodAgeCurier = rs.Fields("CodAgenciaCurier")
        
        LoQueNoExiste = "Código de Destinatario: " & rs.Fields("CodDestinatario")
        If Not IsNull(rs.Fields("CodDestinatario")) Then .CodDestinatario = rs.Fields("CodDestinatario")

        .EstadoGuia = rs.Fields("EstadoGuia")
        Err.Clear
        LoQueNoExiste = ""
    End With
    rs.Close
    Set gc = Nothing
    Exit Sub
ErrTrap:
    If Err.Number = -2147220960 Then
        Set gc = Nothing
        Set rs = Nothing
        Err.Raise Err.Number, "Importacion.PrepararGNOferta", "No existe " & LoQueNoExiste
        Exit Sub
    End If
    Set gc = Nothing
    Set rs = Nothing
    Err.Raise Err.Number, "Importacion", Err.Description
End Sub

