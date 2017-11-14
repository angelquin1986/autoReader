VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmExportar 
   Caption         =   "Exportación"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7710
   ScaleWidth      =   11925
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1935
      ScaleWidth      =   11925
      TabIndex        =   15
      Top             =   0
      Width           =   11925
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar -F5"
         Height          =   516
         Left            =   6840
         TabIndex        =   8
         Top             =   1200
         Width           =   1332
      End
      Begin VB.TextBox txtDestino 
         Height          =   320
         Left            =   840
         TabIndex        =   2
         Top             =   840
         Width           =   6972
      End
      Begin VB.CommandButton cmdExplorar 
         Caption         =   "..."
         Height          =   310
         Left            =   7815
         TabIndex        =   3
         Top             =   840
         Width           =   372
      End
      Begin VB.Frame Frame1 
         Caption         =   "Selecione la Plantilla a utiilizar"
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   0
         Width           =   8175
         Begin VB.CommandButton cmdPlantilla 
            Caption         =   "..."
            Height          =   315
            Left            =   7680
            TabIndex        =   1
            Top             =   240
            Width           =   375
         End
         Begin VB.ComboBox cboPlantilla 
            Height          =   315
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label lblDescripcion 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   4200
            TabIndex        =   19
            Top             =   240
            Width           =   3495
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
            Height          =   195
            Left            =   3240
            TabIndex        =   18
            Top             =   360
            Width           =   840
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
      End
      Begin MSComCtl2.DTPicker dtpFecha2 
         Height          =   330
         Left            =   2760
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   93192193
         CurrentDate     =   36902
      End
      Begin MSComCtl2.DTPicker dtpHora1 
         Height          =   330
         Left            =   840
         TabIndex        =   6
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   93192194
         CurrentDate     =   36902
      End
      Begin MSComCtl2.DTPicker dtpHora2 
         Height          =   330
         Left            =   2760
         TabIndex        =   7
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   93192194
         CurrentDate     =   36902
      End
      Begin MSComCtl2.DTPicker dtpFecha1 
         Height          =   330
         Left            =   840
         TabIndex        =   4
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         Format          =   93192193
         CurrentDate     =   36902
      End
      Begin MSComDlg.CommonDialog dlg1 
         Left            =   4320
         Top             =   1200
         _ExtentX        =   688
         _ExtentY        =   688
         _Version        =   393216
         CancelError     =   -1  'True
         DefaultExt      =   "mdb"
         DialogTitle     =   "Destino de exportación"
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hora "
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   1560
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "~  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   2445
         TabIndex        =   22
         Top             =   1320
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha  "
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Destino  "
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   372
      Left            =   7560
      TabIndex        =   10
      Top             =   5820
      Width           =   1092
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "&Exportar -F9"
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
      Left            =   6000
      TabIndex        =   9
      Top             =   5820
      Width           =   1452
   End
   Begin VSFlex7LCtl.VSFlexGrid grdMsg 
      Align           =   2  'Align Bottom
      Height          =   1275
      Left            =   0
      TabIndex        =   14
      Top             =   6435
      Width           =   11925
      _cx             =   21034
      _cy             =   2249
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
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Exportar.frx":0000
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
      Height          =   3735
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2040
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   529
      TabCaption(0)   =   "Transacciones"
      TabPicture(0)   =   "Exportar.frx":006D
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grdTrans"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Catálogos"
      TabPicture(1)   =   "Exportar.frx":0089
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "grdCat"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VSFlex7LCtl.VSFlexGrid grdTrans 
         Height          =   3195
         Left            =   -74880
         TabIndex        =   12
         Top             =   420
         Width           =   9735
         _cx             =   17171
         _cy             =   5636
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
         Height          =   3195
         Left            =   120
         TabIndex        =   13
         Top             =   420
         Width           =   9735
         _cx             =   17171
         _cy             =   5636
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
End
Attribute VB_Name = "frmExportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const HEIGHT_MIN = 8115    '6544
Const WIDTH_MIN = 8595     '7764

Private mcnDestino As ADODB.Connection
Private mCancelado As Boolean
Private mEjecutando As Boolean
Private mBuscado As Boolean

Private mcnPlantilla As ADODB.Connection '***Angel. 13/feb/2004
Private mPlantilla As clsPlantilla       '***Angel. 20/feb/2004
Private mUltimaPlantilla As String       '***Angel. 20/feb/2004
Private mRutaBDDestino As String

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

Private Sub cmdBuscar_Click()
    Dim v As Variant, rs As Recordset, sql As String, cond As String
    On Error GoTo ErrTrap
    
    MensajeStatus "Está buscando...", vbHourglass
    
    'Primero busca número de registros de Catálogos
    MostrarNumReg
    
    'Limpia la grilla
    grdTrans.Rows = grdTrans.FixedRows      'Para limpiar la selección
    
    'Obtiene listado de transacciones               '*** MAKOTO 06/mar/01 Aumentado 'Nombre'
    sql = "SELECT gc.FechaTrans, gc.CodTrans, gc.NumTrans, gc.Nombre, gc.Descripcion, " & _
                 "cc.CodCentro, gc.Estado " & _
          "FROM GNComprobante gc LEFT JOIN GNCentroCosto cc ON gc.IdCentro = cc.IdCentro "
    
    'Condición de FechaGrabado
    If mPlantilla.BandRangoFechaHora Then
        If Len(cond) > 0 Then cond = cond & " AND "
        cond = cond & HacerCondicion(False, "gc.")
    End If
    
    'Condición de CodTrans
    If Len(mPlantilla.ListaTransacciones) > 0 Then
        If Len(cond) > 0 Then cond = cond & " AND "
        cond = cond & "gc.CodTrans IN (" & mPlantilla.ListaTransacciones & ")"
    Else
        If Len(cond) > 0 Then cond = cond & " AND "
        cond = cond & "gc.CodTrans IN ('-----')"
    End If
    
    If Len(cond) > 0 Then sql = sql & " WHERE " & cond
'    sql = sql & " ORDER BY gc.FechaTrans, gc.CodTrans, gc.NumTrans "
    ''sql = sql & " ORDER BY gc.FechaTrans, gc.HoraTrans "   '*** OLIVER , CAMBIO PARA MEJOR IMPORTACION DE DOC X COBRAR Y PAGAR
    sql = sql & " ORDER BY gc.TransID "  ' jeaa 10-may-2013
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
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
    
    mBuscado = True             '*** MAKOTO 14/mar/01 Agregado
    Habilitar True
    cmdExportar.SetFocus
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
            '***Diego 25/09/2003 cambio  para VATEX
            '.InitDir = App.Path
            .InitDir = txtDestino.Text
            'txtDestino.Text
            .filename = mPlantilla.BDDestino
        Else
            .InitDir = .filename
            .filename = mPlantilla.BDDestino
        End If
        .flags = cdlOFNPathMustExist
        '.Filter = "Base de datos Jet (*.mdb)|*.mdb|Todos (*.*)|*.*"
        .Filter = "Base de datos Jet (*.mdb)|*.mdb|Predefinido (" & _
                  Trim$(mPlantilla.PrefijoNombreArchivo) & "*.mdb)|" & _
                  Trim$(mPlantilla.PrefijoNombreArchivo) & "*.mdb" & _
                  "|Todos (*.*)|*.*"
        .ShowSave
        txtDestino.Text = .filename
    End With
    
    Exit Sub
ErrTrap:
    If Err.Number <> 32755 Then
        DispErr
    End If
    Exit Sub
End Sub

Private Sub cmdExportar_Click()
    Dim s As String
    On Error GoTo ErrTrap
        
    'Verifica si está especificado el destino
    s = Trim$(txtDestino.Text)
    If Len(s) = 0 Then
        MsgBox "Debe especificar el archivo de destino.", vbInformation
        txtDestino.SetFocus
        Exit Sub
    End If
    
    'Si aun no está hecho la búsqueda, llamarlo automaticamente
    If Not mBuscado Then
        cmdBuscar_Click
    End If
    
    'Verifica si existe el archivo de destino
    If ExisteArchivo(s) Then
        Kill s      'Si existelo elimina para sobreescribir
    End If
    
    If Not ExisteArchivo(s) Then
        'Si no existe, lo crea copiando del archivo ARCHIVO_MODELO
                FileCopy App.Path & IIf(Right$(App.Path, 1) <> "\", "\", "") & _
                    ARCHIVO_MODELO, s
    End If

    
'    'Verifica si está seleccionada algúna fila
'    If Not VerificarSeleccionado(grdTrans, grdCat, "exportar") Then Exit Sub
    LimpiarSeleccion grdCat
    LimpiarSeleccion grdTrans
    
    mCancelado = False
    Habilitar False
    
    'Abre la conección con el archivo de destino
    Set mcnDestino = New ADODB.Connection
    s = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & s & ";" & _
     "Jet OLEDB:Database Password='aq9021'" & _
         ";" & "Persist Security Info=False"
    mcnDestino.Open s, "admin", ""
    
    
'''    Private Sub Form_Load()
'''Dim cnn As ADODB.Connection
'''Dim sBase As String
'''Set cnn = New ADODB.Connection
'''
'''sBase = "D:\VBProg_esp\Sii4\SiiTools\_trans.mdb"
'''cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0; " & _
'''     "Data Source=" & sBase & ";" & _
'''     "Jet OLEDB:Database Password='aq9021'"
'''
'''
'''End Sub



    
    'Exporta Catálogos
    ExportarCatalogo
    
    'Exporta Transacciones
    ExportarTrans
    
    'Cierra la BD destino
    mcnDestino.Close
    
    'Guarda en el registro de sistema las configuraciones
    GuardarConfig
    ActualizaRangoFechas  '***Angel. 20/feb/2004
    
salida:
    Set mcnDestino = Nothing
    Habilitar True
    Exit Sub
ErrTrap:
    DispErr
    GoTo salida
End Sub

Private Sub GuardarConfig()
    Dim pos As Integer, s As String
    
    SaveSetting APPNAME, App.Title, Me.Name & ".UltimaPlantilla", cboPlantilla.Text
    pos = InStrRev(txtDestino.Text, "\")
    If pos > 0 Then
        s = Mid$(txtDestino.Text, 1, pos)
    Else
        s = ""
    End If
    SaveSetting APPNAME, App.Title, Me.Name & ".RutaBDDestino", s
End Sub

Private Sub RecuperarConfig()
    mUltimaPlantilla = GetSetting(APPNAME, App.Title, Me.Name & ".UltimaPlantilla", "")
    mRutaBDDestino = GetSetting(APPNAME, App.Title, Me.Name & ".RutaBDDestino", App.Path)
End Sub

Private Sub Habilitar(ByVal v As Boolean)
    mEjecutando = Not v
    cmdBuscar.Enabled = v
    cmdExplorar.Enabled = v
    cmdExportar.Enabled = v And mBuscado
    txtDestino.Enabled = v
    dtpFecha1.Enabled = v And mPlantilla.BandRangoFechaHora
    dtpFecha2.Enabled = v And mPlantilla.BandRangoFechaHora
    dtpHora1.Enabled = v And mPlantilla.BandRangoFechaHora
    dtpHora2.Enabled = v And mPlantilla.BandRangoFechaHora
        
    frmMain.mnuFile.Enabled = v
    frmMain.mnuHerramienta.Enabled = v
    frmMain.mnuTransferir.Enabled = v
    frmMain.mnuCerrarTodas.Enabled = v
End Sub

Private Sub ExportarCatalogo()
    Dim i As Long
    
    sst1.Tab = 1
    
    With grdCat
        For i = .FixedRows To .Rows - 1
            DoEvents
            
            'Si el usuario canceló la operación
            If mCancelado Then Exit For
            
            'Verificando si el Catalogo esta Selecionado para Exportar
            If InStr(mPlantilla.ListaCatalogos, .TextMatrix(i, .ColIndex("Tabla"))) > 0 Then

                If Not ExportarCatalogoSub(.TextMatrix(i, .ColIndex("Catálogo")), _
                                    .TextMatrix(i, .ColIndex("Tabla"))) Then
                    'Si ocurrió algún error y no quizo continuar
                    Exit For
                Else
                    'Sí exportó sin problema, quita la selección
                    .IsSelected(i) = True
                End If
            End If
        Next i
    End With
    MensajeStatus
End Sub

Private Function ExportarCatalogoSub( _
                ByVal Desc As String, _
                ByVal tabla As String) As Boolean
    Dim NumReg As Long
    On Error GoTo ErrTrap

    Select Case tabla
    'Case "GNResponsable" para recortar tamaño de palabra
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
    Case "TSRetencion"
        NumReg = GrabarTSRetencion
    Case "IVBodega"
        NumReg = GrabarIVBodega
    'Case "IVGrupo1", "IVGrupo2", "IVGrupo3", "IVGrupo4", "IVGrupo5"
    Case "IVG1", "IVG2", "IVG3", "IVG4", "IVG5", "IVG6"
        NumReg = GrabarIVGrupo(Val(Right$(tabla, 1)))
    Case "IVInv"
    ' antes Case "IVInventario para comprimir datos"
        NumReg = GrabarIVInventario
    Case "FCVendedor"
        NumReg = GrabarFCVendedor
    'Case "PCGrupo1", "PCGrupo2", "PCGrupo3", "PCGrupo4"
    Case "PCG1", "PCG2", "PCG3", "PCG4"
        NumReg = GrabarPCGrupo(Val(Right$(tabla, 1)))
    Case "PCProvCli(P)", "PCProvCli(C)"
        NumReg = GrabarPCProvCli((Right$(tabla, 3) = "(P)"))
    Case "DescIVGPCG"
        NumReg = GrabarDesctoIVGrupoPCGrupo
    Case "TSFormaC_P"
        NumReg = GrabarTSFormaCobroPago
    'jeaa 17/06/2005
    Case "Motivo"
        NumReg = GrabarMotivo
    'jeaa 17/06/2005
    Case "TCompra"
        NumReg = GrabarTipoCompra
    Case "IVU"
        NumReg = GrabarIVUnidad
    'jeaa 18/02/2008
    Case "Exist"
        NumReg = GrabarIVExist
    Case "DescNumPagIVG"
        NumReg = GrabarDesctoNumPagosIVGrupo
    Case "PCHistorial"
        NumReg = GrabarPCHistorial
    Case "PCProvCli(G)"
        NumReg = GrabarPCGarante(True)
    Case "IVBanco" ' jeaa 22/07/2009
        NumReg = GrabarIVBanco()
    Case "IVTarjeta"
        NumReg = GrabarIVTarjeta()
    Case "PCProvincia"
        NumReg = GrabarPCProvincia
    Case "PCCanton"
        NumReg = GrabarPCCanton
    Case "PCParroquia"
        NumReg = GrabarPCParroquia
    Case "DiasCred"
        NumReg = GrabarPCDiasCredito
    Case "PLAIVGPCG"
        NumReg = GrabarPlazoIVGrupoPCGrupo
    
    End Select
    
    'Si fue cancelado devuelve numreg en negativo
    If NumReg < 0 Then
        DispMsg "Exportar datos de " & Desc, "Cancelado", Abs(NumReg) & " registros."
    Else
        DispMsg "Exportar datos de " & Desc, "OK", NumReg & " registros."
    End If
    
    ExportarCatalogoSub = True
salida:
    Exit Function
ErrTrap:
    DispMsg "Exportar datos de " & Desc, "Error", Err.Description
    If MsgBox(Err.Description & vbCr & vbCr & _
                "Desea continuar con siguiente catálogo?", _
                vbQuestion + vbYesNo) = vbYes Then
        ExportarCatalogoSub = True
    End If
    GoTo salida
End Function


Private Function GrabarCTCuenta() As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    
    'Borra de destino registros de la tabla
    sql = "DELETE FROM CTCuenta"
    mcnDestino.Execute sql
    
    'Abre el orígen
    sql = "SELECT * FROM CTCuenta " & HacerCondicion(True)
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    'Abre el destino
    Set rs2 = New Recordset
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    With rs1
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
        
            rs2.AddNew
            rs2.Fields("CodCuenta") = .Fields("CodCuenta")
            rs2.Fields("NombreCuenta") = .Fields("NombreCuenta")
            rs2.Fields("Nivel") = .Fields("Nivel")
            
            'IdCuentaSuma --> CodCuentaSuma
            rs2.Fields("CodCuentaSuma") = RecuperarCampo("CTCuenta", _
                                            "CodCuenta", _
                                            "IdCuenta=" & .Fields("IdCuentaSuma"))
                                            
            rs2.Fields("TipoCuenta") = .Fields("TipoCuenta")
            rs2.Fields("BandDeudor") = .Fields("BandDeudor")
            rs2.Fields("BandTotal") = .Fields("BandTotal")
            rs2.Fields("BandValida") = .Fields("BandValida")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            rs2.Fields("IDLocal") = .Fields("IDLocal")  'Diego 12/03/2004
            rs2.Update
            .MoveNext
        Loop
        
        GrabarCTCuenta = i
        .Close
        rs2.Close
    End With
    
    Set rs1 = Nothing
    Set rs2 = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarCTCuenta = GrabarCTCuenta * -1
End Function

Private Function GrabarGNResponsable() As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    
    'Borra de destino registros de la tabla
    sql = "DELETE FROM GNResponsable"
    mcnDestino.Execute sql
    
    'Abre el orígen
    sql = "SELECT * FROM GNResponsable " & HacerCondicion(True)
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    'Abre el destino
    Set rs2 = New Recordset
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de Responsables ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
        
            rs2.AddNew
            rs2.Fields("CodResponsable") = .Fields("CodResponsable")
            rs2.Fields("Nombre") = .Fields("Nombre")
            rs2.Fields("BandValida") = .Fields("BandValida")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            rs2.Update
            .MoveNext
        Loop
        
        GrabarGNResponsable = i
        .Close
        rs2.Close
    End With
    
    Set rs1 = Nothing
    Set rs2 = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarGNResponsable = GrabarGNResponsable * -1
End Function

Private Function GrabarGNCentroCosto() As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    
    'Borra de destino registros de la tabla
    sql = "DELETE FROM GNCentroCosto"
    mcnDestino.Execute sql
    
    'Abre el orígen
'    sql = "SELECT * FROM GNCentroCosto " & HacerCondicion(True)
    sql = "SELECT cc.*, " & _
                 "pc1.CodProvCli AS CodProveedor, " & _
                 "pc2.CodProvCli AS CodCliente " & _
          "FROM PCProvCli pc1 RIGHT JOIN " & _
                    "(PCProvCli pc2 RIGHT JOIN GNCentroCosto cc " & _
                    "ON pc2.IdProvCli = cc.IdCliente) " & _
               "ON pc1.IdProvCli = cc.IdProveedor " & _
                 HacerCondicion(True, "cc.")
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    'Abre el destino
    Set rs2 = New Recordset
    sql = "SELECT * FROM GNCentroCosto " & HacerCondicion(True)
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de Centro de costo... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
        
            rs2.AddNew
            rs2.Fields("CodCentro") = .Fields("CodCentro")
            rs2.Fields("Descripcion") = .Fields("Descripcion")
            rs2.Fields("Nombre") = .Fields("Nombre")        '*** MAKOTO 14/feb/01 Agregado
            rs2.Fields("FechaInicio") = .Fields("FechaInicio")
            rs2.Fields("FechaFinal") = .Fields("FechaFinal")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")

            rs2.Fields("CodProveedor") = .Fields("CodProveedor")    '*** MAKOTO 06/mar/01 Agregado
            rs2.Fields("CodCliente") = .Fields("CodCliente")        '*** MAKOTO 06/mar/01 Agregado
            
            rs2.Update
            .MoveNext
        Loop
        
        GrabarGNCentroCosto = i
        .Close
        rs2.Close
    End With
    
    Set rs1 = Nothing
    Set rs2 = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarGNCentroCosto = GrabarGNCentroCosto * -1
End Function

Private Function HacerCondicion( _
                    ByVal ConWhere As Boolean, _
                    Optional ByVal prefijo As String) As String
    If mPlantilla.BandRangoFechaHora Then
        If ConWhere Then HacerCondicion = " WHERE "
        If Len(prefijo) > 0 And prefijo <> "cc." Then
            If mPlantilla.BandTipoFecha Then
                HacerCondicion = HacerCondicion & " " & prefijo & "FechaGrabado BETWEEN "
            Else
                HacerCondicion = HacerCondicion & " (" & prefijo & "FechaTrans + " & prefijo & "HoraTrans) BETWEEN "
            End If
        Else
            HacerCondicion = HacerCondicion & " " & prefijo & "FechaGrabado BETWEEN "
        End If
        
        HacerCondicion = HacerCondicion & FechaYMD(dtpFecha1.value + dtpHora1.value, gobjMain.EmpresaActual.TipoDB, True) & _
                " AND " & FechaYMD(dtpFecha2.value + dtpHora2.value, gobjMain.EmpresaActual.TipoDB, True)
    End If
End Function

Private Function GrabarTSBanco() As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    
    'Borra de destino registros de la tabla
    sql = "DELETE FROM TSBanco"
    mcnDestino.Execute sql
    
    'Abre el orígen
    sql = "SELECT * FROM TSBanco " & HacerCondicion(True)
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    'Abre el destino
    Set rs2 = New Recordset
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de bancos... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
        
            rs2.AddNew
            rs2.Fields("CodBanco") = .Fields("CodBanco")
            rs2.Fields("Descripcion") = .Fields("Descripcion")
            
            If Not mPlantilla.BandIgnorarContabilidad Then     '*** MAKOTO 14/mar/01 Agregado
                'IdCuentaContable --> CodCuentaContable
                rs2.Fields("CodCuentaContable") = RecuperarCampo("CTCuenta", _
                                                "CodCuenta", _
                                                "IdCuenta=" & .Fields("IdCuentaContable"))
            End If
            
            rs2.Fields("NumCuenta") = .Fields("NumCuenta")
            rs2.Fields("Nombre") = .Fields("Nombre")
            rs2.Fields("BandValida") = .Fields("BandValida")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            rs2.Update
            .MoveNext
        Loop
        
        GrabarTSBanco = i
        .Close
        rs2.Close
    End With
    
    Set rs1 = Nothing
    Set rs2 = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarTSBanco = GrabarTSBanco * -1
End Function
        
'*** MAKOTO 12/feb/01 Agregado
Private Function GrabarTSRetencion() As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    
    'Borra de destino registros de la tabla
    sql = "DELETE FROM TSRetencion"
    mcnDestino.Execute sql
    
    'Abre el orígen
    sql = "SELECT * FROM TSRetencion " & HacerCondicion(True)
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    'Abre el destino
    Set rs2 = New Recordset
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de Retenciones... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
        
            rs2.AddNew
            rs2.Fields("CodRetencion") = .Fields("CodRetencion")
            rs2.Fields("Descripcion") = .Fields("Descripcion")
            
            If Not mPlantilla.BandIgnorarContabilidad Then     '*** MAKOTO 14/mar/01 Agregado
                'IdCuentaContable --> CodCuentaContable
                rs2.Fields("CodCuentaActivo") = RecuperarCampo("CTCuenta", _
                                            "CodCuenta", _
                                            "IdCuenta=" & .Fields("IdCuentaActivo"))
                rs2.Fields("CodCuentaPasivo") = RecuperarCampo("CTCuenta", _
                                            "CodCuenta", _
                                            "IdCuenta=" & .Fields("IdCuentaPasivo"))
            End If
            
            rs2.Fields("Porcentaje") = .Fields("Porcentaje")
            rs2.Fields("BandValida") = .Fields("BandValida")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            'jeaa 21/09/2005
            rs2.Fields("CodSRI") = .Fields("CodSRI")
            'jeaa 08/07/2008
            rs2.Fields("BandIVA") = .Fields("BandIVA")
            rs2.Fields("BandCompras") = .Fields("BandCompras")
            rs2.Fields("BandVentas") = .Fields("BandVentas")
            rs2.Update
            .MoveNext
        Loop
        
        GrabarTSRetencion = i
        .Close
        rs2.Close
    End With
    
    Set rs1 = Nothing
    Set rs2 = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarTSRetencion = GrabarTSRetencion * -1
End Function
        
Private Function GrabarIVBodega() As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    
    'Borra de destino registros de la tabla
    sql = "DELETE FROM IVBodega"
    mcnDestino.Execute sql
    
    'Abre el orígen
    sql = "SELECT * FROM IVBodega " & HacerCondicion(True)
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    'Abre el destino
    Set rs2 = New Recordset
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de bodegas... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
        
            rs2.AddNew
            rs2.Fields("CodBodega") = .Fields("CodBodega")
            rs2.Fields("Descripcion") = .Fields("Descripcion")
            rs2.Fields("BandValida") = .Fields("BandValida")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            rs2.Update
            .MoveNext
        Loop
        
        GrabarIVBodega = i
        .Close
        rs2.Close
    End With
    
    Set rs1 = Nothing
    Set rs2 = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarIVBodega = GrabarIVBodega * -1
End Function

Private Function GrabarIVGrupo(ByVal numGrupo As Integer) As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    
    'Borra de destino registros de la tabla
    sql = "DELETE FROM IVGrupo" & numGrupo
    mcnDestino.Execute sql
    
    'Abre el orígen
    sql = "SELECT * FROM IVGrupo" & numGrupo & " " & HacerCondicion(True)
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    'Abre el destino
    Set rs2 = New Recordset
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de " & _
                    gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(numGrupo) & "... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
        
            rs2.AddNew
            rs2.Fields("CodGrupo" & numGrupo) = .Fields("CodGrupo" & numGrupo)
            rs2.Fields("Descripcion") = .Fields("Descripcion")
            rs2.Fields("BandValida") = .Fields("BandValida")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            rs2.Update
            .MoveNext
        Loop
        
        GrabarIVGrupo = i
        .Close
        rs2.Close
    End With
    
    Set rs1 = Nothing
    Set rs2 = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarIVGrupo = GrabarIVGrupo * -1
End Function
        
Private Function GrabarIVInventario() As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    Dim j As Integer
    
    'Borra de destino registros de la tabla
    sql = "DELETE FROM IVInventario"
    mcnDestino.Execute sql
    'Borra de destino registros de IVMateria
    sql = "DELETE FROM IVMateria"
    mcnDestino.Execute sql
    
    'AUC 25/11/2005 Borra de destino los registros IVproveedorDetalle
    sql = "DELETE FROM IVProveedorDetalle"
    mcnDestino.Execute sql
    
    
    'Abre el orígen
    sql = "SELECT * FROM IVInventario " & HacerCondicion(True) & " ORDER BY TIPO " ''agrergado el orden by tipo para cuando importe primero sean lo shijos uy despues las familias
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    'Abre el destino
    Set rs2 = New Recordset
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de inventarios... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
        
            rs2.AddNew
            rs2.Fields("CodInventario") = .Fields("CodInventario")
            rs2.Fields("CodAlterno1") = .Fields("CodAlterno1")
            rs2.Fields("CodAlterno2") = .Fields("CodAlterno2")
            rs2.Fields("Descripcion") = .Fields("Descripcion")
            rs2.Fields("DescripcionDetalle") = .Fields("DescripcionDetalle")
            rs2.Fields("Precio1") = .Fields("Precio1")
            rs2.Fields("Precio2") = .Fields("Precio2")
            rs2.Fields("Precio3") = .Fields("Precio3")
            rs2.Fields("Precio4") = .Fields("Precio4")
            rs2.Fields("CodMoneda") = IIf(Len(.Fields("CodMoneda")) = 0, "USD", .Fields("CodMoneda"))
            rs2.Fields("PorcentajeIVA") = .Fields("PorcentajeIVA")
            rs2.Fields("Comision1") = .Fields("Comision1")
            rs2.Fields("Comision2") = .Fields("Comision2")
            rs2.Fields("Comision3") = .Fields("Comision3")
            rs2.Fields("Comision4") = .Fields("Comision4")
            rs2.Fields("CantLimite1") = .Fields("CantLimite1")
            rs2.Fields("CantLimite2") = .Fields("CantLimite2")
            rs2.Fields("CantLimite3") = .Fields("CantLimite3")
            rs2.Fields("CantLimite4") = .Fields("CantLimite4")
            rs2.Fields("Descuento1") = .Fields("Descuento1")     '***Agregado. 03/ago/2004. Angel
            rs2.Fields("Descuento2") = .Fields("Descuento2")     '***Agregado. 03/ago/2004. Angel
            rs2.Fields("Descuento3") = .Fields("Descuento3")     '***Agregado. 03/ago/2004. Angel
            rs2.Fields("Descuento4") = .Fields("Descuento4")     '***Agregado. 03/ago/2004. Angel
            
            If Not mPlantilla.BandIgnorarContabilidad Then      '*** MAKOTO 14/mar/01 Agregado
                'IdCuentaActivo --> CodCuentaActivo
                rs2.Fields("CodCuentaActivo") = RecuperarCampo("CTCuenta", _
                                            "CodCuenta", "IdCuenta=" & .Fields("IdCuentaActivo"))
                rs2.Fields("CodCuentaCosto") = RecuperarCampo("CTCuenta", _
                                            "CodCuenta", "IdCuenta=" & .Fields("IdCuentaCosto"))
                rs2.Fields("CodCuentaVenta") = RecuperarCampo("CTCuenta", _
                                            "CodCuenta", "IdCuenta=" & .Fields("IdCuentaVenta"))
            End If
            
            rs2.Fields("Unidad") = .Fields("Unidad")
            
            'IdGrupo1-5 --> CodGrupo1-5
            For j = 1 To IVGRUPO_MAX
                rs2.Fields("CodGrupo" & j) = RecuperarCampo("IVGrupo" & j, _
                                            "CodGrupo" & j, "IdGrupo" & j & "=" & .Fields("IdGrupo" & j))
            Next j
            
            'IdProvCli --> CodProvCli
            rs2.Fields("CodProveedor") = RecuperarCampo("PCProvCli", _
                                            "CodProvCli", "IdProvCli=" & .Fields("IdProveedor"))
            
            rs2.Fields("Observacion") = .Fields("Observacion")
            rs2.Fields("ExistenciaMinima") = .Fields("ExistenciaMinima")
            rs2.Fields("ExistenciaMaxima") = .Fields("ExistenciaMaxima")
            rs2.Fields("UnidadMinimaCompra") = .Fields("UnidadMinimaCompra")
            rs2.Fields("UnidadMinimaVenta") = .Fields("UnidadMinimaVenta")
            rs2.Fields("BandValida") = .Fields("BandValida")
            rs2.Fields("BandServicio") = .Fields("BandServicio")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            If IsNull(.Fields("Tipo")) Then
                rs2.Fields("Tipo") = INV_TIPONORMAL     'Diego 05/11/2003
            Else
                rs2.Fields("Tipo") = .Fields("Tipo")    'Diego 19/02/2003
            End If
            If .Fields("Tipo") <> INV_TIPONORMAL Then GrabaFamilia .Fields("CodInventario")
            GrabaProveedor .Fields("CodInventario") 'AUC 25/11/05
            rs2.Fields("ValorRecargo") = .Fields("ValorRecargo")  '***Agregado. 03/ago/2004. Angel
            rs2.Fields("BandFraccion") = .Fields("BandFraccion")    ' *** agregado jeaa 09/04/2005
            rs2.Fields("BandArea") = .Fields("BandArea")    ' *** agregado jeaa 15/09/2005
            rs2.Fields("BandVenta") = .Fields("BandVenta")    ' *** agregado jeaa 26/12/2005
            
            If Not IsNull(.Fields("IdUnidad")) Then
                rs2.Fields("CodUnidad") = RecuperarCampo("IVUnidad", _
                                            "CodUnidad", "IdUnidad=" & .Fields("IdUnidad"))
            End If

            If Not IsNull(.Fields("IdUnidad")) Then
                rs2.Fields("CodUnidadConteo") = RecuperarCampo("IVUnidad", _
                                            "CodUnidad", "IdUnidad=" & .Fields("IdUnidadConteo"))
            End If
            rs2.Fields("CostoUltimoIngreso") = IIf(IsNull(.Fields("CostoUltimoIngreso")), 0, .Fields("CostoUltimoIngreso")) '******* Agregado JEAA 17/04/2006
            rs2.Fields("PorcentajeICE") = IIf(IsNull(.Fields("PorcentajeICE")), 0, .Fields("PorcentajeICE")) '******* Agregado JEAA 17/04/2006
            rs2.Fields("PorDesperdicio") = IIf(IsNull(.Fields("PorDesperdicio")), 0, .Fields("PorDesperdicio")) '******* Agregado JEAA 17/04/2006
            rs2.Fields("Precio5") = IIf(IsNull(.Fields("Precio5")), 0, .Fields("Precio5"))
            rs2.Fields("Comision5") = IIf(IsNull(.Fields("Comision5")), 0, .Fields("Comision5"))
            rs2.Fields("CantLimite5") = IIf(IsNull(.Fields("CantLimite5")), 0, .Fields("CantLimite5"))
            rs2.Fields("Descuento5") = IIf(IsNull(.Fields("Descuento5")), 0, .Fields("Descuento5"))
            rs2.Fields("CantRelUnidad") = IIf(IsNull(.Fields("CantRelUnidad")), 0, .Fields("CantRelUnidad"))
            rs2.Fields("CantRelUnidadCont") = IIf(IsNull(.Fields("CantRelUnidadCont")), 0, .Fields("CantRelUnidadCont"))
            'jeaa 09/07/2008
            rs2.Fields("Descripcion2") = IIf(IsNull(.Fields("Descripcion2")), 0, .Fields("Descripcion2"))
            rs2.Fields("PesoNeto") = IIf(IsNull(.Fields("PesoNeto")), 0, .Fields("PesoNeto"))
            rs2.Fields("PesoBruto") = IIf(IsNull(.Fields("PesoBruto")), 0, .Fields("PesoBruto"))
            
            If Not IsNull(.Fields("IdUnidadPeso")) Then
                rs2.Fields("CodUnidadPeso") = RecuperarCampo("IVUnidad", _
                                            "CodUnidad", "IdUnidad=" & .Fields("IdUnidadPeso"))
            End If
            rs2.Fields("BandConversion") = .Fields("BandConversion")
            rs2.Fields("BandRepGastos") = .Fields("BandRepGastos")
            rs2.Fields("BandNoSeFactura") = .Fields("BandNoSeFactura")
            rs2.Fields("TiempoReposicion") = IIf(IsNull(.Fields("TiempoReposicion")), 0, .Fields("TiempoReposicion"))
            rs2.Fields("TiempoPromVta") = IIf(IsNull(.Fields("TiempoPromVta")), 0, .Fields("TiempoPromVta"))
           
            
            rs2.Update
            .MoveNext
        Loop
        
        GrabarIVInventario = i
        .Close
        rs2.Close
    End With
    
    Set rs1 = Nothing
    Set rs2 = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarIVInventario = GrabarIVInventario * -1
End Function
        
Private Sub GrabaFamilia(ByVal CodMateria As String)
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    Dim j As Integer
    Dim rsAux As Recordset, sqlAux As String
    
    sqlAux = "Select tipo from ivinventario where codinventario='" & CodMateria & "'"
    Set rsAux = gobjMain.EmpresaActual.OpenRecordset(sqlAux)

    'Abre el orígen
    sql = "Select " & _
          "IV.CodInventario, " & _
          "IV1.CodInventario as CodMateria, " & _
          "IVM.Cantidad,ivm.bandPrincipal,ivm.bandModificar,ivm.TarifaJornal,ivm.Rendimiento,ivm.Orden,xCuanto " & _
          "FROM IvInventario IV1 INNER JOIN ( " & _
          "IvMateria IVM INNER JOIN IVInventario IV "

    Select Case rsAux.Fields("Tipo")
        Case 3, 4, 5, 6
            sql = sql & " ON IVM.IdMateria = IV.IdInventario) "
            sql = sql & " ON IV1.IdInventario = IVM.Idinventario "
        Case Else
            sql = sql & " ON IVM.IdInventario = IV.IdInventario) "
            sql = sql & "ON IV1.IdInventario = IVM.IdMateria "
    End Select
    
    sql = sql & "WHERE IV1.CodInventario = '" & CodMateria & "'"
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
    sql = "Select * from IVMateria Where 1=0"
    'Abre el destino
    Set rs2 = New Recordset
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    With rs1
        Do Until .EOF
            i = i + 1
            DoEvents
            rs2.AddNew
            rs2.Fields("CodInventario") = .Fields("CodInventario")
            rs2.Fields("CodMateria") = .Fields("CodMateria")
            rs2.Fields("Cantidad") = .Fields("Cantidad")
            'AUC 12/12/07
            rs2.Fields("bandPrincipal") = .Fields("bandPrincipal")
            rs2.Fields("bandModificar") = .Fields("bandModificar")
            rs2.Fields("TarifaJornal") = .Fields("TarifaJornal")
            rs2.Fields("Rendimiento") = .Fields("Rendimiento")
            rs2.Fields("Orden") = .Fields("Orden")
            rs2.Fields("xCuanto") = .Fields("xCuanto")
            rs2.Update
            .MoveNext
        Loop
        .Close
        rs2.Close
    End With
    Set rs1 = Nothing
    Set rs2 = Nothing
End Sub

Private Function GrabarFCVendedor() As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    
    'Borra de destino registros de la tabla
    sql = "DELETE FROM FCVendedor"
    mcnDestino.Execute sql
    
    'Abre el orígen
    sql = "SELECT * FROM FCVendedor " & HacerCondicion(True)
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    'Abre el destino
    Set rs2 = New Recordset
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de vendedores... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
        
            rs2.AddNew
            rs2.Fields("CodVendedor") = .Fields("CodVendedor")
            rs2.Fields("Nombre") = .Fields("Nombre")
            rs2.Fields("BandValida") = .Fields("BandValida")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            rs2.Update
            .MoveNext
        Loop
        
        GrabarFCVendedor = i
        .Close
        rs2.Close
    End With
    
    Set rs1 = Nothing
    Set rs2 = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarFCVendedor = GrabarFCVendedor * -1
End Function
        
Private Function GrabarPCGrupo(ByVal numGrupo As Integer) As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    
    'Borra de destino registros de la tabla
    sql = "DELETE FROM PCGrupo" & numGrupo
    mcnDestino.Execute sql
    
    'Abre el orígen
    sql = "SELECT * FROM PCGrupo" & numGrupo & " " & HacerCondicion(True)
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    'Abre el destino
    Set rs2 = New Recordset
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de " & _
                    gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupo(numGrupo) & "... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
        
            rs2.AddNew
            rs2.Fields("CodGrupo" & numGrupo) = .Fields("CodGrupo" & numGrupo)
            rs2.Fields("Descripcion") = .Fields("Descripcion")
            rs2.Fields("BandValida") = .Fields("BandValida")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            rs2.Update
            .MoveNext
        Loop
        
        GrabarPCGrupo = i
        .Close
        rs2.Close
    End With
    
    Set rs1 = Nothing
    Set rs2 = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarPCGrupo = GrabarPCGrupo * -1
End Function
        
Private Function GrabarPCProvCli(ByVal Proveedor As Boolean) As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    Dim Desc As String, j As Integer, cond As String
    
        


    'Borra de destino registros de la tabla
    sql = "DELETE FROM PCProvCli WHERE " & IIf(Proveedor, "BandProveedor", "BandCliente") & "<>0"
    mcnDestino.Execute sql
    
    'Abre el orígen
    sql = "SELECT * FROM PCProvCli "
    cond = HacerCondicion(False)
    If Len(cond) > 0 Then cond = cond & " AND "
    cond = cond & IIf(Proveedor, "BandProveedor<>0", "BandCliente<>0")
    
    If Len(cond) > 0 Then sql = sql & " WHERE " & cond
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    'Abre el destino
    Set rs2 = New Recordset
    sql = "SELECT * FROM PCProvCli WHERE 1=0"
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de " & Desc & "... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
        
            rs2.AddNew

            rs2.Fields("CodProvCli") = .Fields("CodProvCli")
            rs2.Fields("Nombre") = .Fields("Nombre")
            rs2.Fields("BandProveedor") = .Fields("BandProveedor")
            rs2.Fields("BandCliente") = .Fields("BandCliente")
            
            If Not mPlantilla.BandIgnorarContabilidad Then     '*** MAKOTO 14/mar/01 Agregado
                'IdCuentaContable --> CodCuentaContable
                rs2.Fields("CodCuentaContable") = RecuperarCampo("CTCuenta", _
                                            "CodCuenta", _
                                            "IdCuenta=" & .Fields("IdCuentaContable"))
                rs2.Fields("CodCuentaContable2") = RecuperarCampo("CTCuenta", _
                                            "CodCuenta", _
                                            "IdCuenta=" & .Fields("IdCuentaContable2"))
            End If
            
            rs2.Fields("Direccion1") = .Fields("Direccion1")
            rs2.Fields("Direccion2") = .Fields("Direccion2")
            rs2.Fields("CodPostal") = .Fields("CodPostal")
            rs2.Fields("Ciudad") = .Fields("Ciudad")
            rs2.Fields("Provincia") = .Fields("Provincia")
            rs2.Fields("Pais") = .Fields("Pais")
            rs2.Fields("Telefono1") = .Fields("Telefono1")
            rs2.Fields("Telefono2") = .Fields("Telefono2")
            rs2.Fields("Telefono3") = .Fields("Telefono3")
            rs2.Fields("Fax") = .Fields("Fax")
            rs2.Fields("RUC") = .Fields("RUC")
            rs2.Fields("EMail") = .Fields("EMail")
            rs2.Fields("LimiteCredito") = .Fields("LimiteCredito")
            rs2.Fields("Banco") = .Fields("Banco")
            rs2.Fields("NumCuenta") = .Fields("NumCuenta")
            rs2.Fields("Swit") = .Fields("Swit")
            rs2.Fields("DirecBanco") = .Fields("DirecBanco")
            rs2.Fields("TelBanco") = .Fields("TelBanco")
            
            rs2.Fields("CodVendedor") = RecuperarCampo("FCVendedor", _
                                            "CodVendedor", _
                                            "IdVendedor=" & .Fields("IdVendedor"))
            'IdGrupo1-4 --> CodGrupo1-4
            For j = 1 To PCGRUPO_MAX
                If Len(.Fields("IdGrupo" & j)) > 0 Then
                    rs2.Fields("CodGrupo" & j) = RecuperarCampo("PCGrupo" & j, _
                                            "CodGrupo" & j, "IdGrupo" & j & "=" & .Fields("IdGrupo" & j))
                Else
                    rs2.Fields("CodGrupo" & j) = RecuperarCampo("PCGrupo" & j, _
                                            "CodGrupo" & j, "")
                End If
            Next j
            
            rs2.Fields("Estado") = .Fields("Estado")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            
            '***Agregado. 08/sep/2003. Angel
            '***Campos referentes a Anexos
            rs2.Fields("TipoDocumento") = .Fields("TipoDocumento")
            rs2.Fields("TipoComprobante") = .Fields("TipoComprobante")
            rs2.Fields("NumAutSRI") = .Fields("NumAutSRI")
            
            '***Agregado. 08/sep/2003. Angel
            '***Campos necesarios para tarjetas de descuentos
            rs2.Fields("NombreAlterno") = .Fields("NombreAlterno")
            rs2.Fields("FechaNacimiento") = .Fields("FechaNacimiento")
            rs2.Fields("FechaEntrega") = .Fields("FechaEntrega")
            rs2.Fields("FechaExpiracion") = .Fields("FechaExpiracion")
            rs2.Fields("TotalDebe") = .Fields("TotalDebe")
            rs2.Fields("TotalHaber") = .Fields("TotalHaber")
            rs2.Fields("Observacion") = .Fields("Observacion")
            rs2.Fields("TipoProvCli") = .Fields("TipoProvCli") 'JEAA 22/12/2005
            rs2.Fields("BandEmpresaPublica") = .Fields("BandEmpresaPublica")   'jeaa 17/01/2008
            'rs.Fields("CodGarante") = RecuperarCampo("PCProvCli", "CodProvCli", "IdProvCli = " & .IdGarante) 'jeaa 10/06/2009
            rs2.Fields("BandGarante") = .Fields("BandGarante") 'jeaa 10/06/2009
            
            rs2.Fields("FechaCreacion") = .Fields("FechaCreacion")
            rs2.Fields("CodProvincia") = CopiarCodProvincia(.Fields("IDProvincia"))
            rs2.Fields("CodCanton") = CopiarCodCanton(.Fields("IdCanton"))
            rs2.Fields("CodParroquia") = CopiarCodParroquia(.Fields("idParroquia"))
            If .Fields("idDiasCredito") <> Null Then
                rs2.Fields("CodDiasCredito") = CopiarCodDiasCredito(.Fields("idDiasCredito"))
            End If
            
            rs2.Fields("TipoSujeto") = .Fields("TipoSujeto")
            rs2.Fields("Sexo") = .Fields("Sexo")
            rs2.Fields("EstadoCivil") = .Fields("EstadoCivil")
            rs2.Fields("OrigenIngresos") = .Fields("OrigenIngresos")
            
                       
            
            
            rs2.Update
            
            'Graba los contactos del prov/cli actual
            GrabarContactos .Fields("CodProvCli")
            
            .MoveNext
        Loop
        
        GrabarPCProvCli = i
        .Close
        rs2.Close
    End With
    
    Set rs1 = Nothing
    Set rs2 = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarPCProvCli = GrabarPCProvCli * -1
End Function

Private Sub GrabarContactos(ByVal codPC As String)
    Dim i As Long, rs As Recordset, sql As String
    Dim pc As PCProvCli
    
    'Recupera PCProvCli para obtener contactos
    Set pc = gobjMain.EmpresaActual.RecuperaPCProvCli(codPC)
    If pc Is Nothing Then Exit Sub
    
    'Primero borra lo exisntente del Prov/Cli actual
    sql = "DELETE FROM PCContacto WHERE CodProvCli = '" & pc.CodProvCli & "'"
    mcnDestino.Execute sql
    
    sql = "SELECT * FROM PCContacto WHERE 1=0"
    Set rs = New Recordset
    rs.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    With rs
        For i = 1 To pc.CountContacto
            .AddNew
            .Fields("CodProvCli") = pc.CodProvCli
            .Fields("Cargo") = pc.Contactos(i).Cargo
            .Fields("EMail") = pc.Contactos(i).Email
            .Fields("Nombre") = pc.Contactos(i).nombre
            .Fields("Orden") = pc.Contactos(i).Orden
            .Fields("Telefono1") = pc.Contactos(i).Telefono1
            .Fields("Telefono2") = pc.Contactos(i).Telefono2
            .Fields("Titulo") = pc.Contactos(i).titulo
            .Update
        Next i
        
        .Close
    End With
    Set rs = Nothing
    Set pc = Nothing
End Sub


Private Sub ExportarTrans()
    Dim i As Long, sql As String, codt As String, numt As Long
    
    sst1.Tab = 0
    
    'Elimina registros en las tablas de transacción
    sql = "DELETE FROM GNComprobante"
    mcnDestino.Execute sql

    sql = "DELETE FROM IVKardex"
    mcnDestino.Execute sql

    sql = "DELETE FROM IVKardexRecargo"
    mcnDestino.Execute sql

    sql = "DELETE FROM PCKardex"
    mcnDestino.Execute sql


    sql = "DELETE FROM TSKardex"
    mcnDestino.Execute sql

    sql = "DELETE FROM TSKardexRet"         '*** MAKOTO 12/feb/01 Agregado
    mcnDestino.Execute sql
    
    sql = "DELETE FROM CTLibroDetalle"
    mcnDestino.Execute sql
    
    sql = "DELETE FROM Anexos" 'Agregado Auc 07/11/2005
    mcnDestino.Execute sql
    
    sql = "DELETE FROM PRLibroDetalle"
    mcnDestino.Execute sql
    
    sql = "DELETE FROM AFKardex"
    mcnDestino.Execute sql
    
    sql = "DELETE FROM PCKardexCHP"
    mcnDestino.Execute sql
    
    
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
                MensajeStatus "Exportando la transacción " & codt & numt & _
                            "     " & i & " de " & .Rows - .FixedRows & _
                            " (" & Format(i * 100 / (.Rows - .FixedRows), "0") & "%)", vbHourglass
    
                If ExportarTransSub(codt, numt) Then
                    'Sí exportó sin problema, quita la selección
                    .IsSelected(i) = True
                End If
'            End If
        Next i
    End With
    MensajeStatus
End Sub

Private Function ExportarTransSub( _
                ByVal codt As String, _
                ByVal numt As Long) As Boolean
    Dim gc As GNComprobante
    On Error GoTo ErrTrap
    
    'Recuperar la transacción
    Set gc = gobjMain.EmpresaActual.RecuperaGNComprobante(0, codt, numt)
    If Not (gc Is Nothing) Then
        'Grabar en las tablas
        GrabarGNComprobante gc
        GrabarGNOferta gc
        'filtrar  por bodega
        If gc.GNTrans.Modulo = "IV" Then   '***Diego
            If gc.CountIVKardex > 0 Then
                GrabarIVKardex gc
            End If
        ElseIf gc.GNTrans.Modulo = "AF" Then   '***Diego
            If gc.CountAFKardex > 0 Then
                GrabarAFKardex gc
            End If
        
        End If
        GrabarIVKardexRecargo gc
        GrabarAFKardexRecargo gc
        GrabarPCKardex gc
        GrabarTSKardex gc
        GrabarTSKardexRet gc            '*** MAKOTO 12/feb/01 Agregado
                    '----AUC 08/11/2005---- aqui exporta anexos
        If gc.Empresa.GNOpcion.ObtenerValor("PermiteControlAspectosAnexos") = "1" And _
        gc.GNTrans.IVVisibleAnexos Then GrabarAnexos gc
            
        If Not mPlantilla.BandIgnorarContabilidad Then     '*** MAKOTO 14/mar/01 Agregado
            GrabarCTLibroDetalle gc
            GrabarPRLibroDetalle gc
        End If
        
        If gc.GNTrans.IVAplicaFinaciamiento Then
            GrabarFinanciamiento gc
        End If
        
        GrabarPCKardexCHP gc
        
        DispMsg "Exportar la trans. " & codt & numt, "OK"
    Else
        'Sacar mensaje
        DispMsg "Exportar la trans. " & codt & numt, "Error", "No se pudo recuperar."
    End If
    
    ExportarTransSub = True
salida:
    Set gc = Nothing
    Exit Function
ErrTrap:
    DispMsg "Exportar la trans. " & codt & numt, "Error", Err.Description
    If ERR_IVFILTROBODEGA = Err.Number Then GoTo salida
    If MsgBox(Err.Description & vbCr & vbCr & _
                "Desea continuar con siguiente transacción?", _
                vbQuestion + vbYesNo) <> vbYes Then
        mCancelado = True
    End If
    GoTo salida
End Function

        
Private Sub GrabarGNComprobante(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset
    
    'Borra de destino si existe el mismo CodTrans, NumTrans
    sql = "DELETE FROM GNComprobante " & _
          "WHERE CodTrans='" & gc.CodTrans & "' AND NumTrans=" & gc.numtrans
    mcnDestino.Execute sql
    
    'Abre el destino para agregar registro
    sql = "SELECT * FROM GNComprobante WHERE 1=0"
    Set rs = New Recordset
    rs.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    With gc
        rs.AddNew
        rs.Fields("CodTrans") = .CodTrans
        rs.Fields("NumTrans") = .numtrans
        rs.Fields("CodAsiento") = .CodAsiento
        rs.Fields("FechaTrans") = .FechaTrans
        rs.Fields("HoraTrans") = .HoraTrans
        rs.Fields("Descripcion") = .Descripcion
        rs.Fields("CodUsuario") = .codUsuario
        rs.Fields("CodUsuarioModifica") = .codUsuarioModifica '***Agregado. 09/ago/2004. Angel
        rs.Fields("CodResponsable") = .CodResponsable
        rs.Fields("NumDocRef") = .numDocRef
        rs.Fields("Estado") = .Estado
        rs.Fields("PosID") = .PosID
        rs.Fields("NumTransCierrePOS") = .NumTransCierrePOS
        rs.Fields("CodCentro") = .CodCentro
        
        'IdTransFuente --> CodTransFuente + NumTransFuente
        rs.Fields("CodTransFuente") = RecuperarCampo("GNComprobante", "CodTrans", "TransID = " & .idTransFuente)
        rs.Fields("NumTransFuente") = RecuperarCampo("GNComprobante", "NumTrans", "TransID = " & .idTransFuente)
        
        rs.Fields("CodMoneda") = .CodMoneda
        rs.Fields("Cotizacion2") = .Cotizacion(2)
        rs.Fields("Cotizacion3") = .Cotizacion(3)
        rs.Fields("Cotizacion4") = .Cotizacion(4)
        
        'IdProveedorRef,IdClienteRef, IdVendedor --> Codxxxx
        rs.Fields("CodProveedorRef") = RecuperarCampo("PCProvCli", "CodProvCli", "IdProvCli = " & .IdProveedorRef)
        rs.Fields("CodClienteRef") = RecuperarCampo("PCProvCli", "CodProvCli", "IdProvCli = " & .IdClienteRef)
        rs.Fields("CodVendedor") = RecuperarCampo("FCVendedor", "CodVendedor", "IdVendedor = " & .IdVendedor)
        rs.Fields("Nombre") = .nombre       '*** MAKOTO 05/feb/01 Agregado
        
        rs.Fields("FechaGrabado") = .FechaGrabado
        rs.Fields("Impresion") = .Impresion '***Agregado. 12/10/2004. jeaa
        rs.Fields("CodMotivo") = RecuperarCampo("Motivo", "CodMotivo", "IdMotivo = " & .IdMotivo) '***Agregado. 17/jun/2005. jeaa
        
        'jeaa 17/04/2006
        rs.Fields("Observacion") = .Observacion
        rs.Fields("Comision") = .Comision
        rs.Fields("FechaDevol") = .FechaDevol
        
        rs.Fields("AutorizacionSRI") = .AutorizacionSRI
        rs.Fields("FechaCaducidadSRI") = .FechaCaducidadSRI
        rs.Fields("CodUsuarioAutoriza") = .CodUsuarioAutoriza '***Agregado. 26/09/2008
        rs.Fields("Estado1") = .Estado1 '***Agregado. 26/09/2008
        rs.Fields("Estado2") = .Estado2 '***Agregado. 26/09/2008
        rs.Fields("numDias") = .NumDias
        rs.Fields("CodGaranteRef") = RecuperarCampo("PCProvCli", "CodProvCli", "IdProvCli = " & .IdGaranteRef)
        
        rs.Update
        rs.Close
    End With
    
    Set gc = Nothing
    Set rs = Nothing
End Sub
        
Private Sub GrabarIVKardex(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset, ivk As IVKardex, i As Long
    Dim cont As Long, v As Variant, r As Long
    'Borra de destino si existe el mismo CodTrans, NumTrans
    'mPlantilla.ListaBodegas = fcbBodega.Text
    
    sql = "DELETE FROM IVKardex " & _
          "WHERE CodTrans='" & gc.CodTrans & "' AND NumTrans=" & gc.numtrans
    mcnDestino.Execute sql
    
    'Abre el destino para agregar registro
    sql = "SELECT * FROM IVKardex WHERE 1=0"
    Set rs = New Recordset
    rs.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    cont = 0
    If gc.CountIVKardex > 0 Then
        For i = 1 To gc.CountIVKardex
            DoEvents
            Set ivk = gc.IVKardex(i)
            If Not (mPlantilla.BandFiltrarxBodega) Then
                AgregaFilaivk ivk, rs, gc
                cont = cont + 1
            Else
                v = Split(mPlantilla.ListaBodegas, ",")
                For r = LBound(v, 1) To UBound(v, 1)
                    If Mid$(v(r), 2, Len(v(r)) - 2) = ivk.CodBodega Then
                        AgregaFilaivk ivk, rs, gc
                        cont = cont + 1
                        Exit For
                    End If
                Next r
            End If
        Next i
    End If
    rs.Close
    'Confirma que haya  por lo menos una fila en la transacion
    If cont = 0 Then
        'Genera Error
        'Borra la tabla gncomprobante
        sql = "DELETE FROM GNComprobante " & _
              "WHERE CodTrans='" & gc.CodTrans & "' AND NumTrans=" & gc.numtrans
        mcnDestino.Execute sql
        Err.Raise ERR_IVFILTROBODEGA, "GrabarIVKardex", MSGERR_IVFILTROBODEGA
    End If
    Set gc = Nothing
    Set ivk = Nothing
    Set rs = Nothing
End Sub


Private Sub AgregaFilaivk(ByVal ivk As IVKardex, ByRef rs As Recordset, ByVal gc As GNComprobante)
    With ivk
        rs.AddNew
        rs.Fields("CodTrans") = gc.CodTrans
        rs.Fields("NumTrans") = gc.numtrans
        rs.Fields("CodInventario") = .CodInventario
        rs.Fields("CodBodega") = .CodBodega
        rs.Fields("Cantidad") = .cantidad
        rs.Fields("CostoTotal") = .CostoTotal
        rs.Fields("CostoRealTotal") = .CostoRealTotal
        rs.Fields("PrecioTotal") = .PrecioTotal
        rs.Fields("PrecioRealTotal") = .PrecioRealTotal
        rs.Fields("Descuento") = .Descuento
        rs.Fields("IVA") = .IVA
        rs.Fields("Orden") = .Orden
        rs.Fields("Nota") = .Nota
        rs.Fields("NumeroPrecio") = .NumeroPrecio         '***Agregado. 11/sep/2003. Angel
        rs.Fields("ValorRecargoItem") = .ValorRecargoItem '***Agregado. 03/ago/2004. Angel
        rs.Fields("TiempoEntrega") = .TiempoEntrega '***Agregado. 23/09/2005
        rs.Fields("bandImprimir") = .bandImprimir
        rs.Fields("idPadre") = .IdPadre
        rs.Fields("bandver") = .bandVer
        rs.Fields("idPadreSub") = .idpadresub
        rs.Update
    End With
End Sub

Private Sub GrabarIVKardexRecargo(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset, ivkr As IVKardexRecargo, i As Long
    
    'Borra de destino si existe el mismo CodTrans, NumTrans
    sql = "DELETE FROM IVKardexRecargo " & _
          "WHERE CodTrans='" & gc.CodTrans & "' AND NumTrans=" & gc.numtrans
    mcnDestino.Execute sql
    
    'Abre el destino para agregar registros
    sql = "SELECT * FROM IVKardexRecargo WHERE 1=0"
    Set rs = New Recordset
    rs.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    For i = 1 To gc.CountIVKardexRecargo
        DoEvents
        
        Set ivkr = gc.IVKardexRecargo(i)
        With ivkr
            rs.AddNew
            rs.Fields("CodTrans") = gc.CodTrans
            rs.Fields("NumTrans") = gc.numtrans
            rs.Fields("CodRecargo") = .codRecargo
            rs.Fields("Porcentaje") = .porcentaje
            rs.Fields("Valor") = .valor
            rs.Fields("BandModificable") = .BandModificable
            rs.Fields("BandOrigen") = .BandOrigen
            rs.Fields("BandProrrateado") = .BandProrrateado
            rs.Fields("AfectaIvaItem") = .AfectaIvaItem
            rs.Fields("Orden") = .Orden
            rs.Update
        End With
    Next i
    
    rs.Close
    Set gc = Nothing
    Set ivkr = Nothing
    Set rs = Nothing
End Sub
        
Private Sub GrabarPCKardex(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset, pck As PCKardex, i As Long
    Dim CodTransAsignado As String, NumTransAsignado As Long, OrdenAsignado As Long
    Dim GuidAsignado As String, pcd As PCDocAsignado
    
    'Borra de destino si existe el mismo CodTrans, NumTrans
    sql = "DELETE FROM PCKardex " & _
          "WHERE CodTrans='" & gc.CodTrans & "' AND NumTrans=" & gc.numtrans
    mcnDestino.Execute sql
    
    'Abre el destino para agregar registros
    sql = "SELECT * FROM PCKardex WHERE 1=0"
    Set rs = New Recordset
    rs.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    For i = 1 To gc.CountPCKardex
        DoEvents
        
        Set pck = gc.PCKardex(i)
        With pck
            rs.AddNew
            rs.Fields("CodTrans") = gc.CodTrans
            rs.Fields("NumTrans") = gc.numtrans
            rs.Fields("CodProvCli") = .CodProvCli
            rs.Fields("CodForma") = .codforma
            
            GuidAsignado = ""
            'Si tiene un documento asignado
            If .idAsignado <> 0 Then
                Set pcd = .RecuperaPCDocAsignado    'Recupera doc. asignado original
                If Not (pcd Is Nothing) Then GuidAsignado = pcd.Guid         '*** MAKOTO 16/mar/01
            End If
            Set pcd = Nothing
            rs.Fields("GuidAsignado") = GuidAsignado        '*** MAKOTO 16/mar/01
            
            rs.Fields("NumLetra") = .NumLetra
            rs.Fields("Debe") = .Debe
            rs.Fields("Haber") = .Haber
            rs.Fields("FechaEmision") = .FechaEmision
            rs.Fields("FechaVenci") = .FechaVenci
            rs.Fields("Observacion") = .Observacion
            rs.Fields("Orden") = .Orden
            'jeaa 22/07/2009
            rs.Fields("CodTarjeta") = .CodTarjeta
            rs.Fields("CodBanco") = .codBanco
            rs.Fields("NumCuenta") = .NumCuenta
            rs.Fields("NumCheque") = .Numcheque
            rs.Fields("TitularCta") = .TitularCta
            
            rs.Fields("Orden") = .Orden
            
            
            rs.Fields("Guid") = .Guid       '*** MAKOTO 16/mar/01 Agregado
            rs.Update
        End With
    Next i
    
    rs.Close
    Set gc = Nothing
    Set pck = Nothing
    Set rs = Nothing
End Sub
        
Private Sub GrabarTSKardex(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset, tsk As TSKardex, i As Long
    
    'Borra de destino si existe el mismo CodTrans, NumTrans
    sql = "DELETE FROM TSKardex " & _
          "WHERE CodTrans='" & gc.CodTrans & "' AND NumTrans=" & gc.numtrans
    mcnDestino.Execute sql
    
    'Abre el destino para agregar registro
    sql = "SELECT * FROM TSKardex WHERE 1=0"
    Set rs = New Recordset
    rs.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    For i = 1 To gc.CountTSKardex
        DoEvents
        
        Set tsk = gc.TSKardex(i)
        With tsk
            rs.AddNew
            rs.Fields("CodTrans") = gc.CodTrans
            rs.Fields("NumTrans") = gc.numtrans
            rs.Fields("CodBanco") = .codBanco
            rs.Fields("Debe") = .Debe
            rs.Fields("Haber") = .Haber
            rs.Fields("Nombre") = .nombre
            rs.Fields("CodTipoDoc") = .CodTipoDoc
            rs.Fields("NumDoc") = .numdoc
            rs.Fields("FechaEmision") = .FechaEmision
            rs.Fields("FechaVenci") = .FechaVenci
            rs.Fields("Observacion") = .Observacion
            rs.Fields("BandConciliado") = .BandConciliado
            rs.Fields("Orden") = .Orden
            rs.Update
        End With
    Next i
    
    rs.Close
    Set gc = Nothing
    Set tsk = Nothing
    Set rs = Nothing
End Sub
        
        
'*** MAKOTO 12/feb/01 Agergado
Private Sub GrabarTSKardexRet(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset, tskr As TSKardexRet, i As Long
    
    'Borra de destino si existe el mismo CodTrans, NumTrans
    sql = "DELETE FROM TSKardexRet " & _
          "WHERE CodTrans='" & gc.CodTrans & "' AND NumTrans=" & gc.numtrans
    mcnDestino.Execute sql
    
    'Abre el destino para agregar registro
    sql = "SELECT * FROM TSKardexRet WHERE 1=0"
    Set rs = New Recordset
    rs.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    For i = 1 To gc.CountTSKardexRet
        DoEvents
        
        Set tskr = gc.TSKardexRet(i)
        With tskr
            rs.AddNew
            rs.Fields("CodTrans") = gc.CodTrans
            rs.Fields("NumTrans") = gc.numtrans
            rs.Fields("CodRetencion") = .CodRetencion
            rs.Fields("Debe") = .Debe
            rs.Fields("Haber") = .Haber
            rs.Fields("Base") = .base
            rs.Fields("NumDoc") = .numdoc
            rs.Fields("Observacion") = .Observacion
            rs.Fields("Orden") = .Orden
            rs.Update
        End With
    Next i
    
    rs.Close
    Set gc = Nothing
    Set tskr = Nothing
    Set rs = Nothing
End Sub
        
Private Sub GrabarCTLibroDetalle(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset, ctd As CTLibroDetalle, i As Long
    
    'Borra de destino si existe el mismo CodTrans, NumTrans
    sql = "DELETE FROM CTLibroDetalle " & _
          "WHERE CodTrans='" & gc.CodTrans & "' AND NumTrans=" & gc.numtrans
    mcnDestino.Execute sql
    
    'Abre el destino para agregar registro
    sql = "SELECT * FROM CTLibroDetalle WHERE 1=0"
    Set rs = New Recordset
    rs.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    For i = 1 To gc.CountCTLibroDetalle
        DoEvents
        
        Set ctd = gc.CTLibroDetalle(i)
        With ctd
            rs.AddNew
            rs.Fields("CodTrans") = gc.CodTrans
            rs.Fields("NumTrans") = gc.numtrans
            rs.Fields("CodCuenta") = .codcuenta
            rs.Fields("Descripcion") = .Descripcion
            rs.Fields("Debe") = .Debe
            rs.Fields("Haber") = .Haber
            rs.Fields("BandIntegridad") = .BandIntegridad
            rs.Fields("Orden") = .Orden
            rs.Update
        End With
    Next i
    
    rs.Close
    Set gc = Nothing
    Set ctd = Nothing
    Set rs = Nothing
End Sub

Private Sub cmdPlantilla_Click()
    If Len(cboPlantilla.Text) > 0 Then mUltimaPlantilla = cboPlantilla.Text
    
    frmListaPlantilla.Inicio "PlantillasExportar", mcnPlantilla
    'Para que actualizar si hubo cambios en la plantilla seleccionada
    CargarComboPlantilla
    PrepararPlantilla
End Sub

Private Sub dtpFecha1_Change()
    'Deshabilitar el botón 'Exportar'           '*** MAKOTO 14/mar/01
    mBuscado = False
    Habilitar True
End Sub

Private Sub dtpFecha2_Change()
    'Deshabilitar el botón 'Exportar'           '*** MAKOTO 14/mar/01
    mBuscado = False
    Habilitar True
End Sub

Private Sub dtpHora1_Change()
    'Deshabilitar el botón 'Exportar'           '*** MAKOTO 14/mar/01
    mBuscado = False
    Habilitar True
End Sub

Private Sub dtpHora2_Change()
    'Deshabilitar el botón 'Exportar'           '*** MAKOTO 14/mar/01
    mBuscado = False
    Habilitar True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF5
        cmdBuscar_Click
        KeyCode = 0
    Case vbKeyF9
        cmdExportar_Click
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
    grdTrans.Rows = grdTrans.FixedRows      'Limpia
    grdMsg.Rows = grdMsg.FixedRows          'Limpia
    
    dtpFecha1.value = Date
    dtpFecha2.value = Date
    dtpHora1.value = Time
    dtpHora2.value = Time
    
    CargarCatalogos grdCat
        
    'Recupera los parámetros guardados en el registro
    RecuperarConfig
    
    '***Angel. 20/feb/2004
    AbrirBasePlantilla
    CargarComboPlantilla
    PrepararPlantilla
    
    '***Angel. 26/feb/2004
    '***Solo supervisor puede acceder a trabajar con las plantillas
    cmdPlantilla.Enabled = gobjMain.UsuarioActual.BandSupervisor
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
    Set mcnDestino = Nothing
    CerrarBasePlantilla '***Angel. 13/feb/2004
End Sub

Private Sub grdCat_BeforeSort(ByVal col As Long, Order As Integer)
    'Impide que cambie el orden de filas mientras esté procesando
    If mEjecutando Then Order = flexSortNone
End Sub


Private Sub grdTrans_BeforeSort(ByVal col As Long, Order As Integer)
    'Impide que cambie el orden de filas mientras esté procesando
    If mEjecutando Then Order = flexSortNone
End Sub

Private Sub sst1_Click(PreviousTab As Integer)
    If (sst1.Tab = 1) And (PreviousTab <> 1) And (Not mEjecutando) Then
        MostrarNumReg
    End If
End Sub

Private Sub MostrarNumReg()
    Dim i As Long, sql As String, rs As Recordset, tabla As String, cond As String

    With grdCat
        .Cols = 4
        .TextMatrix(0, .Cols - 1) = "Núm.Reg."
    
        For i = .FixedRows To .Rows - 1
            cond = ""
            tabla = .TextMatrix(i, .ColIndex("tabla"))
            'jeaa 17/05/2005
            If tabla = "GNResp" Then tabla = "GNResponsable"
            sql = "SELECT Count(*) AS Cnt FROM " & tabla
            'AUC 24/10/2005
            If tabla = "IVInv" Then tabla = "IVinventario"
            sql = "SELECT Count(*) AS Cnt FROM " & tabla
            'jeaa 26/12/2005
            If tabla = "IVG1" Then tabla = "IVGrupo1"
            If tabla = "IVG2" Then tabla = "IVGrupo2"
            If tabla = "IVG3" Then tabla = "IVGrupo3"
            If tabla = "IVG4" Then tabla = "IVGrupo4"
            If tabla = "IVG5" Then tabla = "IVGrupo5"
            If tabla = "IVG6" Then tabla = "IVGrupo6"
'            'jeaa 26/12/2005
            If tabla = "PCG1" Then tabla = "PCGrupo1"
            If tabla = "PCG2" Then tabla = "PCGrupo2"
            If tabla = "PCG3" Then tabla = "PCGrupo3"
            If tabla = "PCG4" Then tabla = "PCGrupo4"
            If tabla = "PCG5" Then tabla = "PCGrupo5"
            If tabla = "TCompra" Then tabla = "IVTipoCompra"
            If tabla = "IVU" Then tabla = "IVUnidad"
            If tabla = "Exist" Then tabla = "IVExist"
            If tabla = "DescNumPagIVG" Then tabla = "DescIVGPCG"
            If tabla = "PCProvincia" Then tabla = "PCProvincia"
            If tabla = "PCCanton" Then tabla = "PCCanton"
            If tabla = "PCParroquia" Then tabla = "PCParroquia"
            If tabla = "DiasCred" Then tabla = "PCDiasCredito"
            If tabla = "PLAIVGPCG" Then tabla = "PlazoIVGPCG"
            
            
            sql = "SELECT Count(*) AS Cnt FROM " & tabla
            
            If tabla = "PCProvCli(P)" Then
                sql = "SELECT Count(*) AS Cnt FROM PCProvCli"
                cond = "BandProveedor<>0"
            ElseIf tabla = "PCProvCli(C)" Then
                sql = "SELECT Count(*) AS Cnt FROM PCProvCli"
                cond = "BandCliente<>0"
            ElseIf tabla = "PCProvCli(G)" Then
                sql = "SELECT Count(*) AS Cnt FROM PCProvCli"
                cond = "BandGarante<>0"
            ElseIf tabla = "TSFormaC_P" Then
                sql = "SELECT Count(*) as Cnt FROM TSFormaCobroPago"
            ElseIf tabla = "IVExist" Then
                sql = "SELECT Count(*) as Exist FROM IVExist"
            If tabla = "PCHistorial" Then
                tabla = "PCHistorial"
            End If
            End If
            If tabla <> "IVExist" Then
            If mPlantilla.BandRangoFechaHora Then
                If Len(cond) > 0 Then cond = cond & " AND "
                cond = cond & HacerCondicion(False)
            End If
            
            If Len(cond) > 0 Then sql = sql & " WHERE " & cond
            End If
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If Not rs.EOF Then
                If tabla <> "IVExist" Then
                    .TextMatrix(i, .Cols - 1) = rs.Fields("Cnt")
                Else
                    .TextMatrix(i, .Cols - 1) = rs.Fields("Exist")
                End If
            End If
            rs.Close
        Next i
    End With
    
    Set rs = Nothing
End Sub

'***Angel. 13/feb/2004
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

'***Angel. 13/feb/2004
Private Sub CerrarBasePlantilla()
    'Cierra base
    On Error Resume Next
    If mcnPlantilla.State <> adStateClosed Then
        mcnPlantilla.Close
    End If
End Sub

'***Angel. 20/feb/2004
Private Sub CargarComboPlantilla()
    Dim sql As String, rs As Recordset
        
    sql = "SELECT CodPlantilla, Descripcion FROM Plantilla_EI " & _
          "WHERE (Tipo=0) AND (BandValida=True) ORDER BY CodPlantilla"
    
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

'***Angel. 20/feb/2004
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

'***Angel. 20/feb/2004
Private Sub RecuperarPlantilla()
    Dim cod As String
    
    cod = cboPlantilla.Text
    If mPlantilla.Recuperar(cod) Then
        Habilitar True
        lblDescripcion.Caption = mPlantilla.Descripcion
        txtDestino.Text = mRutaBDDestino & mPlantilla.BDDestino
        dtpFecha1.value = Format(mPlantilla.FechaDesde, "dd/MM/yyyy")
        dtpFecha2.value = Format(mPlantilla.FechaHasta, "dd/MM/yyyy")
        dtpHora1.value = Format(mPlantilla.HoraDesde, "HH:mm:ss")
        dtpHora2.value = Format(mPlantilla.HoraHasta, "HH:mm:ss")
    Else
        mBuscado = False
        Habilitar False
        lblDescripcion.Caption = ""
        txtDestino.Text = ""
        dtpFecha1.value = mPlantilla.FechaDesde
        dtpFecha2.value = mPlantilla.FechaHasta
        dtpHora1.value = mPlantilla.HoraDesde
        dtpHora2.value = mPlantilla.HoraHasta
        
        'Para que permita cerrar el formulario
        mEjecutando = False
        frmMain.mnuFile.Enabled = True
        frmMain.mnuHerramienta.Enabled = True
        frmMain.mnuTransferir.Enabled = True
        frmMain.mnuCerrarTodas.Enabled = True
    End If
End Sub

'***Angel. 20/feb/2004
Private Sub ActualizaRangoFechas()
    Dim sql As String, cod As String
    
    If Len(mPlantilla.CodPlantilla) > 0 Then
        mPlantilla.FechaDesde = dtpFecha1.value
        mPlantilla.FechaHasta = dtpFecha2.value
        mPlantilla.HoraDesde = dtpHora1.value
        mPlantilla.HoraHasta = dtpHora2.value
        mPlantilla.Grabar
    End If
End Sub


'jeaa 04/01/2005
Private Function GrabarDesctoIVGrupoPCGrupo() As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    'Borra de destino registros de la tabla
    sql = "DELETE FROM DescIVGPCG"
    mcnDestino.Execute sql
    'Abre el orígen
    sql = "SELECT * FROM DescIVGPCG" & " " & HacerCondicion(True)
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
   'Abre el destino
    Set rs2 = New Recordset
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de Descto PCGrupo x IVGrupo... " & _
                   i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            rs2.AddNew
            rs2.Fields("CodPCGrupo") = .Fields("CodPCGrupo")
            rs2.Fields("CodIVGrupo") = .Fields("CodIVGrupo")
            rs2.Fields("Valor") = .Fields("Valor")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            rs2.Update
            .MoveNext
        Loop
        GrabarDesctoIVGrupoPCGrupo = i
       .Close
        rs2.Close
    End With
    Set rs1 = Nothing
    Set rs2 = Nothing
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarDesctoIVGrupoPCGrupo = GrabarDesctoIVGrupoPCGrupo * -1
End Function

'08/01/2005 ----OLIVER
Private Function GrabarTSFormaCobroPago() As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    
    'Borra de destino registros de la tabla
    sql = "DELETE FROM TSFormaCobroPago"
    mcnDestino.Execute sql
    
    'Abre el orígen
    sql = "SELECT * FROM TSFormaCobroPago " & HacerCondicion(True)
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    'Abre el destino
    Set rs2 = New Recordset
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de formas de cobro/pago... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            
            rs2.AddNew
            
            rs2.Fields("CodForma") = .Fields("CodForma")
            rs2.Fields("NombreForma") = .Fields("NombreForma")
            rs2.Fields("Plazo") = .Fields("Plazo")
            rs2.Fields("CambiaFechaVenci") = .Fields("CambiaFechaVenci")
            rs2.Fields("PermiteAbono") = .Fields("PermiteAbono")
            rs2.Fields("BandCobro") = .Fields("BandCobro")
            
            rs2.Fields("CodBanco") = RecuperarCampo("TSBanco", _
                                "CodBanco", _
                                "IdBanco=" & .Fields("IdBanco"))
            
            rs2.Fields("CodTipoDoc") = RecuperarCampo("TSTipoDocBanco", _
                                "CodTipoDoc", _
                                "IdTipoDoc=" & .Fields("IdTipoDoc"))
                                
                                                
            rs2.Fields("BandValida") = .Fields("BandValida")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            rs2.Fields("ConsiderarComoEfectivo") = .Fields("ConsiderarComoEfectivo")
            'jeaa 05/05/2008
            rs2.Fields("IngresoAutomatico") = .Fields("IngresoAutomatico")
            rs2.Fields("CodProvCli") = RecuperarCampo("PcProvCli", "codProvCli", "IdProvCli=" & .Fields("IdProvCli"))
            If .Fields("IdFormaTC") <> 0 Then
                rs2.Fields("CodFormaTC") = RecuperarCampo("TsFormaCobroPago", "CodForma", "IdForma=" & .Fields("IdFormaTC"))
            Else
                rs2.Fields("CodFormaTC") = ""
            End If
            rs2.Fields("DeudaMismoCliente") = .Fields("DeudaMismoCliente")
            rs2.Update
            .MoveNext
        Loop
        GrabarTSFormaCobroPago = i
        .Close
        rs2.Close
    End With
    
    Set rs1 = Nothing
    Set rs2 = Nothing
    
    'Si Fue Cancelado, De Vuelve Numero de Registros en Negativo
    If mCancelado Then GrabarTSFormaCobroPago = GrabarTSFormaCobroPago * -1
End Function

Private Function GrabarMotivo() As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    
    'Borra de destino registros de la tabla
    sql = "DELETE FROM Motivo"
    mcnDestino.Execute sql
    
    'Abre el orígen
    sql = "SELECT * FROM Motivo " & HacerCondicion(True)
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    'Abre el destino
    Set rs2 = New Recordset
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de Motivos de Devolucion... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
        
            rs2.AddNew
            rs2.Fields("CodMotivo") = .Fields("CodMotivo")
            rs2.Fields("Descripcion") = .Fields("Descripcion")
            rs2.Fields("BandValida") = .Fields("BandValida")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            rs2.Update
            .MoveNext
        Loop
        
        GrabarMotivo = i
        .Close
        rs2.Close
    End With
    
    Set rs1 = Nothing
    Set rs2 = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarMotivo = GrabarMotivo * -1
End Function

'AUC Agregado 07/11/2005 para exportar los anexos
Private Sub GrabarAnexos(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset, rsAnexos As Recordset
    Set rsAnexos = gobjMain.EmpresaActual.RecuperarAnexosExportar(gc.CodTrans, gc.numtrans)
    
    If rsAnexos.RecordCount > 0 Then
        'Borra de destino si existe el mismo CodTrans, NumTrans
        sql = "DELETE FROM Anexos " & _
              "WHERE CodTrans='" & gc.CodTrans & "' AND NumTrans=" & gc.numtrans
        mcnDestino.Execute sql
        'Abre el destino para agregar registro
        sql = "SELECT * FROM Anexos WHERE 1=0"
        Set rs = New Recordset
        rs.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
        With gc
            rs.AddNew
            rs.Fields("CodTrans") = rsAnexos.Fields("CodTrans")
            rs.Fields("NumTrans") = rsAnexos.Fields("numtrans")
            rs.Fields("CodCredTrib") = rsAnexos.Fields("CodCredTrib")
            rs.Fields("Codtipocomp") = rsAnexos.Fields("CodTipoComp")
            rs.Fields("numautsri") = rsAnexos.Fields("NumAutSRI")
            rs.Fields("NumSerie") = rsAnexos.Fields("NumSerie")
            If Len(rsAnexos.Fields("NumSecuencial")) > 9 Then
                rs.Fields("NumSecuencial") = Right$(rsAnexos.Fields("NumSecuencial"), 5)
            Else
                rs.Fields("NumSecuencial") = rsAnexos.Fields("NumSecuencial")
            End If
            rs.Fields("BandDevolucion") = rsAnexos.Fields("BandDevolucion")
            rs.Fields("TransIDAfectada") = rsAnexos.Fields("TransIDAfectada")
            rs.Fields("FechaAnexos") = rsAnexos.Fields("FechaAnexos")
            'jeaa 17/04/2006)
            rs.Fields("NumSerieEstablecimiento") = rsAnexos.Fields("NumSerieEstablecimiento")
            rs.Fields("NumSeriePunto") = rsAnexos.Fields("NumSeriePunto")
            rs.Fields("FechaCaducidad") = rsAnexos.Fields("FechaCaducidad")
            rs.Fields("BandFactElec") = rsAnexos.Fields("BandFactElec")
            
            rs.Update
            rs.Close
        End With
    End If
    Set gc = Nothing
    Set rs = Nothing
End Sub


Private Sub GrabaProveedor(ByVal CodInventario As String)

    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    Dim j As Integer
    Dim rsAux As Recordset, sqlAux As String
    'Abre el orígen
    sql = "Select IV.CodInventario, PC.CodProvCli,IVPD.idinventario,IVPD.idproveedor FROM IvInventario IV,PCProvCli PC ,IVProveedorDetalle IVPD  " & _
          "WHERE IV.IdInventario = IVPD.IdInventario AND PC.IdProvCli = IVPD.IdProveedor AND IV.CodInventario = '" & CodInventario & "' "
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
    sql = "Select * from IVProveedorDetalle Where 1=0"
    'Abre el destino
    Set rs2 = New Recordset
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    With rs1
        Do Until .EOF
            i = i + 1
            DoEvents
            rs2.AddNew
            rs2.Fields("CodInventario") = .Fields("CodInventario")
            rs2.Fields("CodProveedor") = .Fields("CodProvCli")
            rs2.Update
            .MoveNext
        Loop
        .Close
        rs2.Close
    End With
    Set rs1 = Nothing
    Set rs2 = Nothing
End Sub

'jeaa 26/12/2005
Private Function GrabarTipoCompra() As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    
    'Borra de destino registros de la tabla
    sql = "DELETE FROM IVTipoCompra"
    mcnDestino.Execute sql
    
    'Abre el orígen
    sql = "SELECT * FROM IVTipoCompra " & HacerCondicion(True)
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    'Abre el destino
    Set rs2 = New Recordset
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de Tipo de Compras... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
        
            rs2.AddNew
            rs2.Fields("CodTipoCompra") = .Fields("CodTipoCompra")
            rs2.Fields("Descripcion") = .Fields("Descripcion")
            rs2.Fields("BandValida") = .Fields("BandValida")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            rs2.Update
            .MoveNext
        Loop
        
        GrabarTipoCompra = i
        .Close
        rs2.Close
    End With
    
    Set rs1 = Nothing
    Set rs2 = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarTipoCompra = GrabarTipoCompra * -1
End Function


Private Function GrabarIVUnidad() As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    
    'Borra de destino registros de la tabla
    sql = "DELETE FROM IVUnidad"
    mcnDestino.Execute sql
    
    'Abre el orígen
    sql = "SELECT * FROM IVUnidad " & HacerCondicion(True)
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    'Abre el destino
    Set rs2 = New Recordset
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de Unidadess... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
        
            rs2.AddNew
            rs2.Fields("CodUnidad") = .Fields("CodUnidad")
            rs2.Fields("Descripcion") = .Fields("Descripcion")
            rs2.Fields("BandValida") = .Fields("BandValida")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            rs2.Update
            .MoveNext
        Loop
        
        GrabarIVUnidad = i
        .Close
        rs2.Close
    End With
    
    Set rs1 = Nothing
    Set rs2 = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarIVUnidad = GrabarIVUnidad * -1
End Function


Private Function GrabarIVExist() As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long, sql2 As String
    Dim v As Variant, r As Integer, cond As String
    
    'Borra de destino registros de la tabla
    sql = "DELETE FROM IVExist"
    mcnDestino.Execute sql
    
    'Abre el orígen
'    sql = "SELECT * FROM IVExist "
    sql = " select codinventario, codbodega, exist, existmin, existmax"
    sql = sql & " from ivexist ive"
    sql = sql & " inner join ivinventario ivi"
    sql = sql & " on ive.idinventario=ivi.idinventario"
    sql = sql & " inner join ivbodega ivb on ivb.idbodega = ive.idbodega"
    v = Split(mPlantilla.ListaBodegas, ",")
    For r = LBound(v, 1) To UBound(v, 1)
        If r = LBound(v, 1) Then
            cond = v(r) & ","
        Else
            cond = cond & v(r) & ","
        End If
            
    Next r
    If Len(cond) > 0 Then
        cond = Mid$(cond, 1, Len(cond) - 1)
        sql = sql & " where codbodega  in (" & cond & ")"
    End If

    
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    'Abre el destino
    Set rs2 = New Recordset
    
    sql2 = " select codinventario, codbodega, exist, existmin, existmax"
    sql2 = sql2 & " from ivexist "
    rs2.Open sql2, mcnDestino, adOpenDynamic, adLockPessimistic
    
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de IVExist... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
        
            rs2.AddNew
            rs2.Fields("CodInventario") = .Fields("CodInventario")
            rs2.Fields("CodBodega") = .Fields("CodBodega")
            rs2.Fields("Exist") = .Fields("Exist")
            rs2.Fields("ExistMin") = .Fields("ExistMin")
            rs2.Fields("ExistMax") = .Fields("ExistMax")
            rs2.Update
            .MoveNext
        Loop
        
        GrabarIVExist = i
        .Close
        rs2.Close
    End With
    
    Set rs1 = Nothing
    Set rs2 = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarIVExist = GrabarIVExist * -1
End Function

'jeaa 04/01/2005
Private Function GrabarDesctoNumPagosIVGrupo() As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    'Borra de destino registros de la tabla
    sql = "DELETE FROM DescNumPagIVG"
    mcnDestino.Execute sql
    'Abre el orígen
    sql = "SELECT * FROM DescIVGPCG" & " " & HacerCondicion(True)
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
   'Abre el destino
    Set rs2 = New Recordset
    sql = "SELECT * FROM DescNumPagIVG" & " " & HacerCondicion(True)
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de Descto NumPagos  x IVGrupo... " & _
                   i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            rs2.AddNew
            rs2.Fields("CodIVGrupo") = .Fields("CodIVGrupo")
            rs2.Fields("Valor") = .Fields("Valor")
            'JEAA 12/09/2008
            rs2.Fields("NumPagos") = .Fields("NumPagos")
            rs2.Fields("BandOmiteRecDesc") = .Fields("BandOmiteRecDesc")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            rs2.Update
            .MoveNext
        Loop
        GrabarDesctoNumPagosIVGrupo = i
       .Close
        rs2.Close
    End With
    Set rs1 = Nothing
    Set rs2 = Nothing
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarDesctoNumPagosIVGrupo = GrabarDesctoNumPagosIVGrupo * -1
End Function


Private Function GrabarPCHistorial() As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    Dim Desc As String, j As Integer, cond As String
    'Borra de destino registros de la tabla
    sql = "DELETE FROM PCHistorial " ' & IIf(Proveedor, "BandProveedor", "BandCliente") & "<>0"
    mcnDestino.Execute sql
    'Abre el orígen
    sql = "SELECT * FROM PCHistorial "
    cond = HacerCondicion(False)
    If Len(cond) > 0 Then sql = sql & " WHERE " & cond
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
    'Abre el destino
    Set rs2 = New Recordset
    sql = "SELECT * FROM PCHistorial WHERE 1=0"
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de " & Desc & "... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            rs2.AddNew
            rs2.Fields("IdProvCli") = .Fields("IdProvCli")
            rs2.Fields("TransId") = .Fields("TransId")
            rs2.Fields("FechaTrans") = .Fields("FechaTrans")
            rs2.Fields("Estado") = .Fields("Estado")
            rs2.Fields("Trans") = .Fields("Trans")
            rs2.Fields("Descripcion") = .Fields("Descripcion")
            rs2.Fields("Valor") = .Fields("Valor")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            rs2.Update
            .MoveNext
        Loop
        GrabarPCHistorial = i
        .Close
        rs2.Close
    End With
    Set rs1 = Nothing
    Set rs2 = Nothing
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarPCHistorial = GrabarPCHistorial * -1
End Function


Private Sub GrabarFinanciamiento(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset, rsFinanciamiento As Recordset
    Set rsFinanciamiento = gobjMain.EmpresaActual.RecuperarFinanciamientoExportar(gc.CodTrans, gc.numtrans)
    
    If rsFinanciamiento.RecordCount > 0 Then
        'Borra de destino si existe el mismo CodTrans, NumTrans
        sql = "DELETE FROM GnFinanciamiento " & _
              "WHERE CodTrans='" & gc.CodTrans & "' AND NumTrans=" & gc.numtrans
        mcnDestino.Execute sql
        'Abre el destino para agregar registro
        sql = "SELECT * FROM GnFinanciamiento WHERE 1=0"
        Set rs = New Recordset
        rs.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
        With gc
            rs.AddNew
            rs.Fields("CodTrans") = rsFinanciamiento.Fields("CodTrans")
            rs.Fields("NumTrans") = rsFinanciamiento.Fields("numtrans")
            rs.Fields("TasaMensual") = rsFinanciamiento.Fields("TasaMensual")
            rs.Fields("MesesGracia") = rsFinanciamiento.Fields("MesesGracia")
            rs.Fields("ValorEntrada") = rsFinanciamiento.Fields("ValorEntrada")
            rs.Fields("FechaPrimerPago") = rsFinanciamiento.Fields("FechaPrimerPago")
            rs.Fields("DiaPago") = rsFinanciamiento.Fields("DiaPago")
            rs.Fields("NumeroPagos") = rsFinanciamiento.Fields("NumeroPagos")
            rs.Fields("ValorSegundaEntrada") = rsFinanciamiento.Fields("ValorSegundaEntrada")
            rs.Fields("FechaSegundoPago") = rsFinanciamiento.Fields("FechaSegundoPago")
            rs.Fields("ValorIntereses") = rsFinanciamiento.Fields("ValorIntereses")
            rs.Update
            rs.Close
        End With
    End If
    Set gc = Nothing
    Set rs = Nothing
    Set rsFinanciamiento = Nothing
End Sub



Private Function GrabarPCGarante(ByVal PGarante As Boolean) As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    Dim Desc As String, j As Integer, cond As String
    
        


    'Borra de destino registros de la tabla
    sql = "DELETE FROM PCProvCli WHERE BandGarante <>0"
    mcnDestino.Execute sql
    
    'Abre el orígen
    sql = "SELECT * FROM PCProvCli "
    cond = HacerCondicion(False)
    If Len(cond) > 0 Then cond = cond & " AND "
    cond = cond & " BandGarante<>0 "
    
    If Len(cond) > 0 Then sql = sql & " WHERE " & cond
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    'Abre el destino
    Set rs2 = New Recordset
    sql = "SELECT * FROM PCProvCli WHERE 1=0"
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de " & Desc & "... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
        
            rs2.AddNew

            rs2.Fields("CodProvCli") = .Fields("CodProvCli")
            rs2.Fields("Nombre") = .Fields("Nombre")
            rs2.Fields("BandProveedor") = .Fields("BandProveedor")
            rs2.Fields("BandCliente") = .Fields("BandCliente")
            rs2.Fields("BandGarante") = .Fields("BandGarante")
            
            If Not mPlantilla.BandIgnorarContabilidad Then     '*** MAKOTO 14/mar/01 Agregado
                'IdCuentaContable --> CodCuentaContable
                rs2.Fields("CodCuentaContable") = RecuperarCampo("CTCuenta", _
                                            "CodCuenta", _
                                            "IdCuenta=" & .Fields("IdCuentaContable"))
                rs2.Fields("CodCuentaContable2") = RecuperarCampo("CTCuenta", _
                                            "CodCuenta", _
                                            "IdCuenta=" & .Fields("IdCuentaContable2"))
            End If
            
            rs2.Fields("Direccion1") = .Fields("Direccion1")
            rs2.Fields("Direccion2") = .Fields("Direccion2")
            rs2.Fields("CodPostal") = .Fields("CodPostal")
            rs2.Fields("Ciudad") = .Fields("Ciudad")
            rs2.Fields("Provincia") = .Fields("Provincia")
            rs2.Fields("Pais") = .Fields("Pais")
            rs2.Fields("Telefono1") = .Fields("Telefono1")
            rs2.Fields("Telefono2") = .Fields("Telefono2")
            rs2.Fields("Telefono3") = .Fields("Telefono3")
            rs2.Fields("Fax") = .Fields("Fax")
            rs2.Fields("RUC") = .Fields("RUC")
            rs2.Fields("EMail") = .Fields("EMail")
            rs2.Fields("LimiteCredito") = .Fields("LimiteCredito")
            rs2.Fields("Banco") = .Fields("Banco")
            rs2.Fields("NumCuenta") = .Fields("NumCuenta")
            rs2.Fields("Swit") = .Fields("Swit")
            rs2.Fields("DirecBanco") = .Fields("DirecBanco")
            rs2.Fields("TelBanco") = .Fields("TelBanco")
            
            rs2.Fields("CodVendedor") = RecuperarCampo("FCVendedor", _
                                            "CodVendedor", _
                                            "IdVendedor=" & .Fields("IdVendedor"))
            'IdGrupo1-4 --> CodGrupo1-4
            For j = 1 To PCGRUPO_MAX
                If Len(.Fields("IdGrupo" & j)) > 0 Then
                    rs2.Fields("CodGrupo" & j) = RecuperarCampo("PCGrupo" & j, _
                                            "CodGrupo" & j, "IdGrupo" & j & "=" & .Fields("IdGrupo" & j))
                Else
                    rs2.Fields("CodGrupo" & j) = RecuperarCampo("PCGrupo" & j, _
                                            "CodGrupo" & j, "")
                End If
            Next j
            
            rs2.Fields("Estado") = .Fields("Estado")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            
            '***Agregado. 08/sep/2003. Angel
            '***Campos referentes a Anexos
            rs2.Fields("TipoDocumento") = .Fields("TipoDocumento")
            rs2.Fields("TipoComprobante") = .Fields("TipoComprobante")
            rs2.Fields("NumAutSRI") = .Fields("NumAutSRI")
            
            '***Agregado. 08/sep/2003. Angel
            '***Campos necesarios para tarjetas de descuentos
            rs2.Fields("NombreAlterno") = .Fields("NombreAlterno")
            rs2.Fields("FechaNacimiento") = .Fields("FechaNacimiento")
            rs2.Fields("FechaEntrega") = .Fields("FechaEntrega")
            rs2.Fields("FechaExpiracion") = .Fields("FechaExpiracion")
            rs2.Fields("TotalDebe") = .Fields("TotalDebe")
            rs2.Fields("TotalHaber") = .Fields("TotalHaber")
            rs2.Fields("Observacion") = .Fields("Observacion")
            rs2.Fields("TipoProvCli") = .Fields("TipoProvCli") 'JEAA 22/12/2005
            rs2.Fields("BandEmpresaPublica") = .Fields("BandEmpresaPublica")   'jeaa 17/01/2008
            'rs.Fields("CodGarante") = RecuperarCampo("PCProvCli", "CodProvCli", "IdProvCli = " & .IdGarante) 'jeaa 10/06/2009
            rs2.Update
            
            'Graba los contactos del prov/cli actual
            GrabarContactos .Fields("CodProvCli")
            
            .MoveNext
        Loop
        
        GrabarPCGarante = i
        .Close
        rs2.Close
    End With
    
    Set rs1 = Nothing
    Set rs2 = Nothing
    
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarPCGarante = GrabarPCGarante * -1
End Function


'jeaa 22/07/2009
Private Function GrabarIVBanco() As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    'Borra de destino registros de la tabla
    sql = "DELETE FROM IVBanco"
    mcnDestino.Execute sql
    'Abre el orígen
    sql = "SELECT CodBanco, Descripcion, ivb.BandValida,  ivb.FechaGrabado ,tsf.codforma, codprovcli   "
    sql = sql & " FROM IVBanco Ivb"
    sql = sql & " left join TsformaCobroPago tsf on Ivb.idforma=tsf.idforma"
    sql = sql & " left join Pcprovcli pc on pc.idprovcli=ivb.idcliente"
    sql = sql & " " & HacerCondicion(True, "ivb.")
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
   'Abre el destino
    Set rs2 = New Recordset
    sql = "SELECT * FROM IVBanco" & " " & HacerCondicion(True)
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de IVBanco... " & _
                   i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            rs2.AddNew
            rs2.Fields("CodBanco") = .Fields("CodBanco")
            rs2.Fields("Descripcion") = .Fields("Descripcion")
            rs2.Fields("CodForma") = .Fields("CodForma")
            rs2.Fields("CodCliente") = .Fields("CodProvcli")
            rs2.Fields("BandValida") = .Fields("BandValida")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            rs2.Update
            .MoveNext
        Loop
        GrabarIVBanco = i
       .Close
        rs2.Close
    End With
    Set rs1 = Nothing
    Set rs2 = Nothing
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarIVBanco = GrabarIVBanco * -1
End Function

Private Function GrabarIVTarjeta() As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    'Borra de destino registros de la tabla
    sql = "DELETE FROM IVTarjeta"
    mcnDestino.Execute sql
    'Abre el orígen
    sql = "SELECT CodTarjeta, Descripcion, ivt.BandValida,  ivt.FechaGrabado ,tsf.codforma   "
    sql = sql & " FROM IVTarjeta Ivt"
    sql = sql & " left join TsformaCobroPago tsf on Ivt.idforma=tsf.idforma"
    sql = sql & " " & HacerCondicion(True, "ivt.")
    
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
   'Abre el destino
    Set rs2 = New Recordset
    

    sql = "SELECT * FROM IVTarjeta" & " " & HacerCondicion(True)
    
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de IVTarjeta... " & _
                   i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            rs2.AddNew
            rs2.Fields("CodTarjeta") = .Fields("CodTarjeta")
            rs2.Fields("Descripcion") = .Fields("Descripcion")
            rs2.Fields("CodForma") = .Fields("CodForma")
            rs2.Fields("BandValida") = .Fields("BandValida")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            rs2.Update
            .MoveNext
        Loop
        GrabarIVTarjeta = i
       .Close
        rs2.Close
    End With
    Set rs1 = Nothing
    Set rs2 = Nothing
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarIVTarjeta = GrabarIVTarjeta * -1
End Function



Private Sub GrabarPRLibroDetalle(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset, ctd As PRLibroDetalle, i As Long
    
    'Borra de destino si existe el mismo CodTrans, NumTrans
    sql = "DELETE FROM PRLibroDetalle " & _
          "WHERE CodTrans='" & gc.CodTrans & "' AND NumTrans=" & gc.numtrans
    mcnDestino.Execute sql
    
    'Abre el destino para agregar registro
    sql = "SELECT * FROM PRLibroDetalle WHERE 1=0"
    Set rs = New Recordset
    rs.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    For i = 1 To gc.CountPRLibroDetalle
        DoEvents
        
        Set ctd = gc.PRLibroDetalle(i)
        With ctd
            rs.AddNew
            rs.Fields("CodTrans") = gc.CodTrans
            rs.Fields("NumTrans") = gc.numtrans
            rs.Fields("CodCuenta") = .codcuenta
            rs.Fields("Descripcion") = .Descripcion
            rs.Fields("Debe") = .Debe
            rs.Fields("Haber") = .Haber
            rs.Fields("BandIntegridad") = .BandIntegridad
            rs.Fields("FechaEjec") = .FechaEjec
            rs.Fields("Orden") = .Orden
            rs.Update
        End With
    Next i
    
    rs.Close
    Set gc = Nothing
    Set ctd = Nothing
    Set rs = Nothing
End Sub


Private Function GrabarPCProvincia() As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    'Borra de destino registros de la tabla
    sql = "DELETE FROM PCPROVINCIA"
    mcnDestino.Execute sql
    'Abre el orígen
    sql = "SELECT * FROM PCPROVINCIA" & HacerCondicion(True)
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
    'Abre el destino
    Set rs2 = New Recordset
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de " & _
                    "PCProvincia ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            rs2.AddNew
            rs2.Fields("CodProvincia") = .Fields("CodProvincia")
            rs2.Fields("Descripcion") = .Fields("Descripcion")
            rs2.Fields("BandValida") = .Fields("BandValida")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            rs2.Update
            .MoveNext
        Loop
        GrabarPCProvincia = i
        .Close
        rs2.Close
    End With
    Set rs1 = Nothing
    Set rs2 = Nothing
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarPCProvincia = GrabarPCProvincia * -1
End Function

Private Function GrabarPCCanton() As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    'Borra de destino registros de la tabla
    sql = "DELETE FROM PCCANTON"
    mcnDestino.Execute sql
    'Abre el orígen
    sql = "SELECT * FROM PCCANTON" & HacerCondicion(True)
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
    'Abre el destino
    Set rs2 = New Recordset
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de " & _
                    "PCCanton ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            rs2.AddNew
            rs2.Fields("CodCanton") = .Fields("CodCanton")
            rs2.Fields("CodProvincia") = CopiarCodProvincia(.Fields("idProvincia"))
            rs2.Fields("Descripcion") = .Fields("Descripcion")
            rs2.Fields("BandValida") = .Fields("BandValida")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            rs2.Update
            .MoveNext
        Loop
        GrabarPCCanton = i
        .Close
        rs2.Close
    End With
    Set rs1 = Nothing
    Set rs2 = Nothing
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarPCCanton = GrabarPCCanton * -1
End Function

Private Function GrabarPCParroquia() As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    'Borra de destino registros de la tabla
    sql = "DELETE FROM PCParroquia"
    mcnDestino.Execute sql
    'Abre el orígen
    sql = "SELECT * FROM PCParroquia" & HacerCondicion(True)
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
    'Abre el destino
    Set rs2 = New Recordset
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de " & _
                    "PCParroquia ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            rs2.AddNew
            rs2.Fields("CodParroquia") = .Fields("CodParroquia")
            rs2.Fields("CodCanton") = CopiarCodCanton(.Fields("idcanton"))
            rs2.Fields("Descripcion") = .Fields("Descripcion")
            rs2.Fields("BandValida") = .Fields("BandValida")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            rs2.Update
            .MoveNext
        Loop
        GrabarPCParroquia = i
        .Close
        rs2.Close
    End With
    Set rs1 = Nothing
    Set rs2 = Nothing
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarPCParroquia = GrabarPCParroquia * -1
End Function

Private Function CopiarCodProvincia(ByVal IdProvincia As Long) As String
Dim sql As String
Dim s As String
Dim rs As Recordset
    sql = "Select codprovincia from pcprovincia where idprovincia  = " & IdProvincia
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    Do While Not rs.EOF
        s = rs!codProvincia
        rs.MoveNext
    Loop
    CopiarCodProvincia = s
End Function

Private Function CopiarCodCanton(ByVal Idcanton As Long) As String
Dim sql As String
Dim s As String
Dim rs As Recordset
    sql = "Select codcanton from pccanton where idcanton  = " & Idcanton
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    Do While Not rs.EOF
        s = rs!codCanton
        rs.MoveNext
    Loop
    CopiarCodCanton = s
End Function

Private Function CopiarCodParroquia(ByVal IdParroquia As Long) As String
Dim sql As String
Dim s As String
Dim rs As Recordset
    sql = "Select codParroquia from pcparroquia where idparroquia  = " & IdParroquia
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    Do While Not rs.EOF
        s = rs!codParroquia
        rs.MoveNext
    Loop
    CopiarCodParroquia = s
End Function


Private Sub GrabarAFKardex(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset, ivk As AFKardex, i As Long
    Dim cont As Long, v As Variant, r As Long
    'Borra de destino si existe el mismo CodTrans, NumTrans
    'mPlantilla.ListaBodegas = fcbBodega.Text
    
    sql = "DELETE FROM AFKardex " & _
          "WHERE CodTrans='" & gc.CodTrans & "' AND NumTrans=" & gc.numtrans
    mcnDestino.Execute sql
    
    'Abre el destino para agregar registro
    sql = "SELECT * FROM AFKardex WHERE 1=0"
    Set rs = New Recordset
    rs.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    cont = 0
    If gc.CountAFKardex > 0 Then
        For i = 1 To gc.CountAFKardex
            DoEvents
            Set ivk = gc.AFKardex(i)
            If Not (mPlantilla.BandFiltrarxBodega) Then
                AgregaFilaafk ivk, rs, gc
                cont = cont + 1
            Else
                v = Split(mPlantilla.ListaBodegas, ",")
                For r = LBound(v, 1) To UBound(v, 1)
                    If Mid$(v(r), 2, Len(v(r)) - 2) = ivk.CodBodega Then
                        AgregaFilaafk ivk, rs, gc
                        cont = cont + 1
                        Exit For
                    End If
                Next r
            End If
        Next i
    End If
    rs.Close
    'Confirma que haya  por lo menos una fila en la transacion
    If cont = 0 Then
        'Genera Error
        'Borra la tabla gncomprobante
        sql = "DELETE FROM GNComprobante " & _
              "WHERE CodTrans='" & gc.CodTrans & "' AND NumTrans=" & gc.numtrans
        mcnDestino.Execute sql
        Err.Raise ERR_IVFILTROBODEGA, "GrabarAFKardex", MSGERR_IVFILTROBODEGA
    End If
    Set gc = Nothing
    Set ivk = Nothing
    Set rs = Nothing
End Sub


Private Sub AgregaFilaafk(ByVal ivk As AFKardex, ByRef rs As Recordset, ByVal gc As GNComprobante)
    With ivk
        rs.AddNew
        rs.Fields("CodTrans") = gc.CodTrans
        rs.Fields("NumTrans") = gc.numtrans
        rs.Fields("CodInventario") = .CodInventario
        rs.Fields("CodBodega") = .CodBodega
        rs.Fields("Cantidad") = .cantidad
        rs.Fields("CostoTotal") = .CostoTotal
        rs.Fields("CostoRealTotal") = .CostoRealTotal
        rs.Fields("PrecioTotal") = .PrecioTotal
        rs.Fields("PrecioRealTotal") = .PrecioRealTotal
        rs.Fields("Descuento") = .Descuento
        rs.Fields("IVA") = .IVA
        rs.Fields("Orden") = .Orden
        rs.Fields("Nota") = .Nota
        rs.Fields("NumeroPrecio") = .NumeroPrecio         '***Agregado. 11/sep/2003. Angel
        rs.Fields("TiempoEntrega") = .TiempoEntrega '***Agregado. 23/09/2005
        rs.Update
    End With
End Sub


Private Sub GrabarAFKardexRecargo(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset, ivkr As AFKardexRecargo, i As Long
    
    'Borra de destino si existe el mismo CodTrans, NumTrans
    sql = "DELETE FROM AFKardexRecargo " & _
          "WHERE CodTrans='" & gc.CodTrans & "' AND NumTrans=" & gc.numtrans
    mcnDestino.Execute sql
    
    'Abre el destino para agregar registros
    sql = "SELECT * FROM AFKardexRecargo WHERE 1=0"
    Set rs = New Recordset
    rs.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    For i = 1 To gc.CountAFKardexRecargo
        DoEvents
        
        Set ivkr = gc.AFKardexRecargo(i)
        With ivkr
            rs.AddNew
            rs.Fields("CodTrans") = gc.CodTrans
            rs.Fields("NumTrans") = gc.numtrans
            rs.Fields("CodRecargo") = .codRecargo
            rs.Fields("Porcentaje") = .porcentaje
            rs.Fields("Valor") = .valor
            rs.Fields("BandModificable") = .BandModificable
            rs.Fields("BandOrigen") = .BandOrigen
            rs.Fields("BandProrrateado") = .BandProrrateado
            rs.Fields("AfectaIvaItem") = .AfectaIvaItem
            rs.Fields("Orden") = .Orden
            rs.Update
        End With
    Next i
    
    rs.Close
    Set gc = Nothing
    Set ivkr = Nothing
    Set rs = Nothing
End Sub


Private Function GrabarPCDiasCredito() As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    'Borra de destino registros de la tabla
    sql = "DELETE FROM PCDiasCredito"
    mcnDestino.Execute sql
    'Abre el orígen
    sql = "SELECT * FROM PCDiasCredito" & HacerCondicion(True)
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
    'Abre el destino
    Set rs2 = New Recordset
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de " & _
                    "PCDiasCredito ... " & _
                    i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            rs2.AddNew
            rs2.Fields("CodDiasCredito") = .Fields("CodDiasCredito")
            rs2.Fields("Descripcion") = .Fields("Descripcion")
            rs2.Fields("BandValida") = .Fields("BandValida")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            rs2.Update
            .MoveNext
        Loop
        GrabarPCDiasCredito = i
        .Close
        rs2.Close
    End With
    Set rs1 = Nothing
    Set rs2 = Nothing
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarPCDiasCredito = GrabarPCDiasCredito * -1
End Function


Private Function CopiarCodDiasCredito(ByVal IdDiasCredito As Long) As String
Dim sql As String
Dim s As String
Dim rs As Recordset
    sql = "Select codDiasCredito from pcDiasCredito where idDiasCredito  = " & IdDiasCredito
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    Do While Not rs.EOF
        s = rs!CodDiasCredito
        rs.MoveNext
    Loop
    CopiarCodDiasCredito = s
End Function


Private Function GrabarPlazoIVGrupoPCGrupo() As Long
    Dim sql As String, rs1 As Recordset, rs2 As Recordset, i As Long
    'Borra de destino registros de la tabla
    sql = "DELETE FROM PlazoIVGPCG"
    mcnDestino.Execute sql
    'Abre el orígen
    sql = "SELECT * FROM PlazoIVGPCG" & " " & HacerCondicion(True)
    Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
   'Abre el destino
    Set rs2 = New Recordset
    sql = "SELECT * FROM PlazoIVGPCG" & " " & HacerCondicion(True)
    rs2.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    With rs1
        Do Until .EOF
            i = i + 1
            MensajeStatus "Grabando Catálogo de Plazo IVGrupo x PCGrupo... " & _
                   i & " de " & .RecordCount & _
                    " (" & Format(i * 100 / .RecordCount, "0") & "%)", vbHourglass
            DoEvents
            If mCancelado Then
                MsgBox "El proceso fue cancelado.", vbInformation
                Exit Do
            End If
            rs2.AddNew
            rs2.Fields("CodIVGrupo") = .Fields("CodIVGrupo")
            rs2.Fields("CodPCGrupo") = .Fields("CodPCGrupo")
            rs2.Fields("Valor") = .Fields("Valor")
            rs2.Fields("FechaGrabado") = .Fields("FechaGrabado")
            rs2.Update
            .MoveNext
        Loop
        GrabarPlazoIVGrupoPCGrupo = i
       .Close
        rs2.Close
    End With
    Set rs1 = Nothing
    Set rs2 = Nothing
    'Si fue cancelado, devuelve numero de registros en negativo
    If mCancelado Then GrabarPlazoIVGrupoPCGrupo = GrabarPlazoIVGrupoPCGrupo * -1
End Function


Private Sub GrabarPCKardexCHP(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset, pck As PCKardex, i As Long
    Dim CodTransAsignado As String, NumTransAsignado As Long, OrdenAsignado As Long
    Dim GuidAsignado As String, pcd As PCDocAsignado
    Dim pckCHP As PCKardexCHP, pcdCHP As PCDocAsignadoCHP
    
    'Borra de destino si existe el mismo CodTrans, NumTrans
    sql = "DELETE FROM PCKardexCHP " & _
          "WHERE CodTrans='" & gc.CodTrans & "' AND NumTrans=" & gc.numtrans
    mcnDestino.Execute sql
    
    'Abre el destino para agregar registros
    sql = "SELECT * FROM PCKardexCHP WHERE 1=0"
    Set rs = New Recordset
    rs.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    For i = 1 To gc.CountPCKardexCHP
        DoEvents
        
        Set pckCHP = gc.PCKardexCHP(i)
        With pckCHP
            rs.AddNew
            rs.Fields("CodTrans") = gc.CodTrans
            rs.Fields("NumTrans") = gc.numtrans
            rs.Fields("CodProvCli") = .CodProvCli
            rs.Fields("CodForma") = .codforma
            
            If gc.GNTrans.CodPantalla = "TSICHP" Then
            GuidAsignado = ""
            'Si tiene un documento asignado
            If .IdAsignadoPCK <> 0 Then
                Set pcdCHP = .RecuperaPCDocAsignadoOriginal    'Recupera doc. asignado original
                If Not (pcdCHP Is Nothing) Then GuidAsignado = pcdCHP.Guid         '*** MAKOTO 16/mar/01
            End If
            Set pcdCHP = Nothing
            rs.Fields("GuidAsignadoPCK") = GuidAsignado        '*** MAKOTO 16/mar/01
            Else
            GuidAsignado = ""
            'Si tiene un documento asignado
            If .idAsignado <> 0 Then
                Set pcdCHP = .RecuperaPCDocAsignadoOriginalCHP   'Recupera doc. asignado original
                If Not (pcdCHP Is Nothing) Then GuidAsignado = pcdCHP.Guid         '*** MAKOTO 16/mar/01
            End If
            End If
            Set pcd = Nothing
            rs.Fields("GuidAsignadoPCK") = GuidAsignado        '*** MAKOTO 16/mar/01
            rs.Fields("NumLetra") = .NumLetra
            rs.Fields("Debe") = .Debe
            rs.Fields("Haber") = .Haber
            rs.Fields("FechaEmision") = .FechaEmision
            rs.Fields("FechaVenci") = .FechaVenci
            rs.Fields("Observacion") = .Observacion
            rs.Fields("Orden") = .Orden
            'jeaa 22/07/2009
            rs.Fields("CodTarjeta") = .CodTarjeta
            rs.Fields("CodBanco") = .codBanco
            rs.Fields("NumCuenta") = .NumCuenta
            rs.Fields("NumCheque") = .Numcheque
            rs.Fields("TitularCta") = .TitularCta
            
            rs.Fields("Orden") = .Orden
            
            
            rs.Fields("Guid") = .Guid       '*** MAKOTO 16/mar/01 Agregado
            rs.Update
        End With
    Next i
    
    rs.Close
    Set gc = Nothing
    Set pck = Nothing
    Set rs = Nothing
End Sub


Private Sub GrabarGNOferta(ByVal gc As GNComprobante)
    Dim sql As String, rs As Recordset
    
    'Borra de destino si existe el mismo CodTrans, NumTrans
    sql = "DELETE FROM GNOferta " & _
          "WHERE CodTrans='" & gc.CodTrans & "' AND NumTrans=" & gc.numtrans
    mcnDestino.Execute sql
    
    'Abre el destino para agregar registro
    sql = "SELECT * FROM GNOferta WHERE 1=0"
    Set rs = New Recordset
    rs.Open sql, mcnDestino, adOpenDynamic, adLockPessimistic
    
    With gc
        rs.AddNew
        rs.Fields("CodTrans") = .CodTrans
        rs.Fields("NumTrans") = .numtrans
        rs.Fields("Atencion") = .Atencion
        rs.Fields("FormaPago") = .FormaPago
        rs.Fields("TiempoEntrega") = .TiempoEntrega
        rs.Fields("Validez") = .Validez
        rs.Fields("Detalles") = .Detalles
        rs.Fields("FechaValidez") = .FechaValidez
        rs.Fields("Observaciones") = .Observaciones
        rs.Fields("FechaEntrega") = .FechaEntrega
        rs.Fields("TiempoEstimadoEntrega") = .TiempoEstimadoEntrega
                
        rs.Fields("CodGarante2") = RecuperarCampo("PCProvCli", "CodProvcli", "idProvcli= " & .IdGaranteRef2)
        rs.Fields("CodInventario") = RecuperarCampo("IVInventario", "CodInventario", "idInventario= " & .idinventario)
        rs.Fields("CodEmpleadoRef") = RecuperarCampo("Empleado", "CodProvcli", "idprovcli= " & .IdEmpleadoRef)
        rs.Fields("NumDireccion") = .NumDireccion
        rs.Fields("DireccionTransporte") = .DirTransporte
        rs.Fields("Opcion") = .Opcion
        rs.Fields("CodAgencia") = RecuperarCampo("PCAgencia", "CodAgencia", "idagencia= " & .IdAgencia)
        'rs.Fields("BandOmitir") = .BandOmitir
        'rs.Fields("CodPlan") = .CodPlan
        rs.Fields("CodCobrador") = RecuperarCampo("FCVendedor", "CodVendedor", "idVendedor= " & .IdCobrador)
        rs.Fields("CodAgenciaCurier") = RecuperarCampo("GNAgenciaCurier", "CodAgeCurier", "idagecurier= " & .IdAgeCurier)
        rs.Fields("CodDestinatario") = RecuperarCampo("PCProvcli", "CodProvcli", "idprovcli= " & .IdDestinatario)
        rs.Fields("EstadoGuia") = .EstadoGuia
                
        rs.Update
        rs.Close
    End With
    
    Set gc = Nothing
    Set rs = Nothing
End Sub


