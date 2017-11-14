VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.UserControl IVGNBusqueda 
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5460
   ClipControls    =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   5460
   Begin VSFlex7Ctl.VSFlexGrid grd 
      Align           =   1  'Align Top
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5460
      _cx             =   9631
      _cy             =   3625
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
      Cols            =   29
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
      BackColorFrozen =   8421504
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Menu mnuDetalle 
      Caption         =   "&Detalle"
      Begin VB.Menu mnuAgregar 
         Caption         =   "&Agregar fila"
      End
      Begin VB.Menu mnuEliminar 
         Caption         =   "&Eliminar fila"
      End
      Begin VB.Menu lin1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTotalizar 
         Caption         =   "&Totalizar repetidos"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuGrabarPrecio 
         Caption         =   "Grabar precios"
      End
      Begin VB.Menu mnuOptimizarCantidad 
         Caption         =   "Optimizar cantidades"
      End
      Begin VB.Menu lin2 
         Caption         =   "-"
      End
      Begin VB.Menu mnucrearitem 
         Caption         =   "&Crear nuevo Item"
      End
      Begin VB.Menu mnuKardexItemCli 
         Caption         =   "Kardex Item x Cliente"
         Shortcut        =   ^K
      End
   End
End
Attribute VB_Name = "IVGNBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Option Explicit

Private Const MAXLEN_NOTA As Integer = 80

'Ubicación de columnas
Private Const COL_NUMFILA = 0
Private Const COL_CODBODEGA = 1
Private Const COL_G1 = 2
Private Const COL_G2 = 3
Private Const COL_G3 = 4
Private Const COL_G4 = 5
Private Const COL_G5 = 6
Private Const COL_CODITEM = 7
Private Const COL_CODALT = 8
Private Const COL_DESC = 9
Private Const COL_EXIST = 10
Private Const COL_CANT = 11
Private Const COL_UNIDAD = 12       '*** MAKOTO 22/jul/00
Private Const COL_CU = 13
Private Const COL_CUR = 14
Private Const COL_CT = 15
Private Const COL_CTR = 16
Private Const COL_UTIL = 17
Private Const COL_PU = 18
Private Const COL_PUR = 19
Private Const COL_PUIVA = 20    '******** ******** jeaa 22-Dic-03 22-12-03
Private Const COL_PT = 21
Private Const COL_PTR = 22
Private Const COL_PTIVA = 23    '******** ******** jeaa 22-Dic-03 22-12-03
Private Const COL_PORDCNT = 24
Private Const COL_PORIVA = 25
Private Const COL_VALIVA = 26
Private Const COL_OBS = 27          '*************** 'jeaa 22-04-2005
Private Const COL_TENTREGA = 28          '*************** 'jeaa 22-04-2005
'Private Const COL_NOTA = 25         '*** MAKOTO 16/dic/00 Eliminado

'Objeto de comprobante
Private WithEvents mobjGNComp As GNComprobante
Attribute mobjGNComp.VB_VarHelpID = -1

Private ItemsImportados As Variant

'Event Declarations:
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event TotalizadoItem()
Event AgregarFilaAuto(ByRef Cancel As Boolean)  '*** MAKOTO 12/dic/00 Agregado
Event OcultarColumna()

'*** MAKOTO 09/nov/00 Agregado
Private mTransBodega As Boolean     'Si está True, visualiza solo los item
                                    'con cantidad negativa. (En apariencia, SIN signo '-')
'*** MAKOTO 14/nov/00 Agregado
Private mCodBodegaOrigen As String  'Código de bodega de origen
                                    'para mostrar la existencia
                                    


Private gobjSiiXML As Object            '*** Agregado Esteban 06/02/2004 para que crear clientes con un objeto creado con siixml
Private numfila As Integer
Dim ANCHO_COLS(0 To 29) As Long
Private mItemsSigno As Integer  'AUC 13/12/05
Private mbooItems As Boolean
Dim ItemsTmp() As Variant
Private PrecioSelec As Currency

                                    

'*** MAKOTO 09/nov/00 Agregado
Public Property Get TransBodega() As Boolean
    TransBodega = mTransBodega
End Property

'*** MAKOTO 09/nov/00 Agregado
Public Property Let TransBodega(ByVal value As Boolean)
    mTransBodega = value
End Property

'*** MAKOTO 14/nov/00 Agregado
Public Property Get CodBodegaOrigen() As String
    'Disponible solo si está en modo de Transferencia bodega
    If Not mTransBodega Then Err.Raise ERR_INVALIDO, "IVGNBusqueda", MSGERR_INVALIDO
    
    CodBodegaOrigen = mCodBodegaOrigen
End Property

'*** MAKOTO 14/nov/00 Agregado
Public Property Let CodBodegaOrigen(ByVal value As String)
    'Disponible solo si está en modo de Transferencia bodega
    If Not mTransBodega Then Err.Raise ERR_INVALIDO, "IVGNBusqueda", MSGERR_INVALIDO
    
    mCodBodegaOrigen = value
End Property



Private Sub ConfigCols()
    Dim s As String, i As Integer

    s = "^#|<Bodega"
    With mobjGNComp.Empresa.GNOpcion
        s = s & "|<" & .EtiqGrupo(1)
        s = s & "|<" & .EtiqGrupo(2)
        s = s & "|<" & .EtiqGrupo(3)
        s = s & "|<" & .EtiqGrupo(4)
        s = s & "|<" & .EtiqGrupo(5)
    End With
    '*** MAKOTO 22/jul/00 Agregado 'Unidad'
    s = s & "|<Cod.Item|<Cod.Alterno|<Descripción|>Exist|>Cant|<Unid|>Costo U.|>Costo U.R." & _
            "|>Costo T.|>Costo T.R.|>%Util|>Precio U.|>Precio U.R.|>P. Unitario + IVA|>Precio T.|>Precio T.R." & _
            "|>P. Total + IVA|>%Dscnt|>%IVA|>IVA|<Observación"
    If mobjGNComp.GNTrans.IVVerificaFechaCompra Then
        s = s & "|<Fecha Ult.Compra"
    Else
        s = s & "|<Tiem.Entrega"
    End If

    With grd
        .FormatString = s
         GetColsWidth
        .ColWidth(COL_NUMFILA) = 500
        .ColWidth(COL_CODBODEGA) = 800              'Cod.Bodega
        .ColWidth(COL_G1) = 800                     'Cod.Grupo1
        .ColWidth(COL_G2) = 800                     'Cod.Grupo2
        .ColWidth(COL_G3) = 800                     'Cod.Grupo3
        .ColWidth(COL_G4) = 800                     'Cod.Grupo4
        .ColWidth(COL_G5) = 800                     'Cod.Grupo5
        .ColWidth(COL_CODITEM) = 1800               'Cod.Item
        .ColWidth(COL_CODALT) = 1800                'Cod.Alterno
        .ColWidth(COL_DESC) = 2000                  'Descripcion
        .ColWidth(COL_EXIST) = COLANCHO_CANT        'Existencia
        .ColWidth(COL_CANT) = COLANCHO_CANT         'Cantidad
        .ColWidth(COL_UNIDAD) = 600                 'Unidad de medida   '*** MAKOTO 22/jul/00
        .ColWidth(COL_CU) = COLANCHO_CUR            'Costo U.
        .ColWidth(COL_CUR) = COLANCHO_CUR           'Costo U.Real
        .ColWidth(COL_CT) = COLANCHO_CUR            'Costo T.
        .ColWidth(COL_CTR) = COLANCHO_CUR           'Costo T.Real
        .ColWidth(COL_UTIL) = 1000                  '%Utilidad
        .ColWidth(COL_PU) = COLANCHO_CUR            'Precio U.
        .ColWidth(COL_PUR) = COLANCHO_CUR           'Precio U.Real
        .ColWidth(COL_PUIVA) = COLANCHO_CUR           'Precio U.+IVA    '******** jeaa 22-Dic-03
        .ColWidth(COL_PT) = COLANCHO_CUR            'Precio T.
        .ColWidth(COL_PTR) = COLANCHO_CUR           'Precio T.Real
        .ColWidth(COL_PTIVA) = COLANCHO_CUR           'Precio T.+IVA    '******** jeaa 22-Dic-03
        .ColWidth(COL_PORDCNT) = 1000               '%Descuento
        .ColWidth(COL_PORIVA) = 1000                '%IVA
        .ColWidth(COL_VALIVA) = COLANCHO_CUR        'Valor de IVA
        .ColWidth(COL_OBS) = 2000        'Observacion ******** jeaa 22/04/2005
        .ColWidth(COL_TENTREGA) = 2000        'Tiempo de Entrega ******** jeaa 22/09/2005
        
        .ColDataType(COL_DESC) = flexDTString
        .ColDataType(COL_EXIST) = flexDTDouble
        .ColDataType(COL_CANT) = flexDTDouble
        .ColDataType(COL_UNIDAD) = flexDTString     '*** MAKOTO 22/jul/00
        .ColDataType(COL_CU) = flexDTCurrency
        .ColDataType(COL_CUR) = flexDTCurrency
        .ColDataType(COL_CT) = flexDTCurrency
        .ColDataType(COL_CTR) = flexDTCurrency
        .ColDataType(COL_UTIL) = flexDTSingle
        .ColDataType(COL_PU) = flexDTCurrency
        .ColDataType(COL_PUIVA) = flexDTCurrency    '******** jeaa 22-Dic-03
        .ColDataType(COL_PUR) = flexDTCurrency
        .ColDataType(COL_PT) = flexDTCurrency
        .ColDataType(COL_PTR) = flexDTCurrency
        .ColDataType(COL_PTIVA) = flexDTCurrency    '******** jeaa 22-Dic-03
        .ColDataType(COL_PORDCNT) = flexDTSingle
        .ColDataType(COL_PORIVA) = flexDTSingle
        .ColDataType(COL_VALIVA) = flexDTCurrency
        .ColDataType(COL_OBS) = flexDTString            '******** jeaa 22/04/2005
        .ColDataType(COL_TENTREGA) = flexDTString            '******** jeaa 22/09/2005
        
        ConfigColsVisible
        
        'No modificables siempre
'        .ColData(COL_DESC) = -1                     'Descripcion de item
        .ColData(COL_EXIST) = -1                    'Existencia
        .ColData(COL_UNIDAD) = -1                    'Unidad de medida '*** MAKOTO 22/jul/00
        .ColData(COL_CUR) = -1
        .ColData(COL_CTR) = -1
        .ColData(COL_PUR) = -1
        .ColData(COL_PTR) = -1
        .ColData(COL_PORIVA) = -1
        .ColData(COL_VALIVA) = -1
        .ColData(COL_PUIVA) = -1
        .ColData(COL_PTIVA) = -1
        For i = 0 To .Cols - 1
             .ColWidth(i) = ANCHO_COLS(i)
         Next i
    
    End With
    ConfigColsFormato
End Sub

Private Sub ConfigColsVisible()
    Dim v As Long, v2 As Long, v3 As Long
    
    v = mobjGNComp.GNTrans.ColVisible
    v2 = mobjGNComp.GNTrans.ColEditable
    v3 = mobjGNComp.GNTrans.ColSeleccionable
    
    With grd
        '.Cols = COL_VALIVA + 1            '*** MAKOTO 16/dic/00 Agregado
        '.Cols = COL_OBS + 1                     '*** jeaa 22/04/05 modificado
        .Cols = COL_TENTREGA + 1                     '*** jeaa 22/09/05 modificado
        .ColHidden(COL_CODBODEGA) = Not CBool(v And &H80000001)
        .ColHidden(COL_G1) = Not CBool(v And &H80000002)
        .ColHidden(COL_G2) = Not CBool(v And &H80000004)
        .ColHidden(COL_G3) = Not CBool(v And &H80000008)
        .ColHidden(COL_G4) = Not CBool(v And &H80000010)
        .ColHidden(COL_G5) = Not CBool(v And &H80000020)
        .ColHidden(COL_CODITEM) = Not CBool(v And &H80000040)
        .ColHidden(COL_CODALT) = Not CBool(v And &H80000080)
        .ColHidden(COL_DESC) = Not CBool(v And &H80000100)
        .ColHidden(COL_EXIST) = Not CBool(v And &H80000200)
        .ColHidden(COL_CANT) = Not CBool(v And &H80000400)
        .ColHidden(COL_CU) = Not CBool(v And &H80000800)
        .ColHidden(COL_CUR) = Not CBool(v And &H80001000)
        .ColHidden(COL_CT) = Not CBool(v And &H80002000)
        .ColHidden(COL_CTR) = Not CBool(v And &H80004000)
        .ColHidden(COL_UTIL) = Not CBool(v And &H80008000)
        .ColHidden(COL_PU) = Not CBool(v And &H80010000)
        .ColHidden(COL_PUR) = Not CBool(v And &H80020000)
        .ColHidden(COL_PUIVA) = Not CBool(v And &H80040000) '******** jeaa 22-Dic-03
        .ColHidden(COL_PT) = Not CBool(v And &H80080000)
        .ColHidden(COL_PTR) = Not CBool(v And &H80100000)
        .ColHidden(COL_PTIVA) = Not CBool(v And &H80200000) '******** jeaa 22-Dic-03
        .ColHidden(COL_PORDCNT) = Not CBool(v And &H80400000)
        .ColHidden(COL_PORIVA) = Not CBool(v And &H80800000)
        .ColHidden(COL_VALIVA) = Not CBool(v And &H81000000)
        .ColHidden(COL_UNIDAD) = Not CBool(v And &H82000000)   '*** MAKOTO 22/jul/00
        .ColHidden(COL_OBS) = Not CBool(v And &H84000000)           '******** jeaa 22/04/2005
        .ColHidden(COL_TENTREGA) = Not CBool(v And &H88000000)           '******** jeaa 22/09/2005
        'jeaa 07/01/2009
        If mobjGNComp.GNTrans.IVVerificaFechaCompra Then
            .ColHidden(COL_TENTREGA) = False
        Else
            .ColHidden(COL_TENTREGA) = Not CBool(v And &H88000000)           '******** jeaa 22/09/2005
        End If

        
        .ColData(COL_CODBODEGA) = CInt(Not (CBool(v2 And &H80000001) Or CBool(v3 And &H80000001)))
        .ColData(COL_G1) = CInt(Not (CBool(v2 And &H80000002) Or CBool(v3 And &H80000002)))
        .ColData(COL_G2) = CInt(Not (CBool(v2 And &H80000004) Or CBool(v3 And &H80000004)))
        .ColData(COL_G3) = CInt(Not (CBool(v2 And &H80000008) Or CBool(v3 And &H80000008)))
        .ColData(COL_G4) = CInt(Not (CBool(v2 And &H80000010) Or CBool(v3 And &H80000010)))
        .ColData(COL_G5) = CInt(Not (CBool(v2 And &H80000020) Or CBool(v3 And &H80000020)))
        
        'Diego 27/02/2002 Para  quela columna sea solo  editable
'        .ColData(COL_CODITEM) = CInt(Not (CBool(v2 And &H80000040) Or CBool(v3 And &H80000040)))
'        .ColData(COL_CODALT) = CInt(Not (CBool(v2 And &H80000080) Or CBool(v3 And &H80000080)))

        If (CBool(v2 And &H80000040) = True) And (CBool(v3 And &H80000040) = False) Then
          .ColData(COL_CODITEM) = 1   'Solo  Editable
        Else
          .ColData(COL_CODITEM) = CInt(Not (CBool(v2 And &H80000040) Or CBool(v3 And &H80000040)))
          'Carga amemoria la pantalla  de busquea
        End If
        
        If (CBool(v2 And &H80000080) = True) And (CBool(v3 And &H80000080) = False) Then
          .ColData(COL_CODALT) = 1  'Solo Editable
        Else
          .ColData(COL_CODALT) = CInt(Not (CBool(v2 And &H80000080) Or CBool(v3 And &H80000080)))
        End If
        .ColData(COL_DESC) = CInt(Not (CBool(v2 And &H80000100) Or CBool(v3 And &H80000100)))
        .ColData(COL_EXIST) = CInt(Not CBool(v2 And &H80000200))
        .ColData(COL_CANT) = CInt(Not CBool(v2 And &H80000400))
        .ColData(COL_CU) = CInt(Not CBool(v2 And &H80000800))
        .ColData(COL_CUR) = CInt(Not CBool(v2 And &H80001000))
        .ColData(COL_CT) = CInt(Not CBool(v2 And &H80002000))
        .ColData(COL_CTR) = CInt(Not CBool(v2 And &H80004000))
        .ColData(COL_UTIL) = CInt(Not CBool(v2 And &H80008000))
'        .ColData(COL_PU) = CInt(Not (CBool(v2 And &H80010000) Or CBool(v3 And &H80010000)))
        .ColData(COL_PU) = CInt(Not CBool(v2 And &H80010000))
        If (Not CBool(v2 And &H80010000)) And CBool(v3 And &H80010000) Then   'para el caso de que es solo selecionable
            .ColData(COL_PU) = 1    'Solo Editable
        End If
        .ColData(COL_PUR) = CInt(Not CBool(v2 And &H80020000))
        .ColData(COL_PUIVA) = CInt(Not CBool(v2 And &H80040000))    '******** jeaa 22-Dic-03
        If (Not CBool(v2 And &H80040000)) And CBool(v3 And &H80040000) Then   'para el caso de que es solo selecionable
            .ColData(COL_PUIVA) = 1    'Solo Editable
        End If
        .ColData(COL_PT) = CInt(Not CBool(v2 And &H80080000))
        .ColData(COL_PTR) = CInt(Not CBool(v2 And &H80100000))
        .ColData(COL_PUIVA) = CInt(Not CBool(v2 And &H80200000))    '******** jeaa 22-Dic-03
        .ColData(COL_PORDCNT) = CInt(Not CBool(v2 And &H80400000))
        .ColData(COL_PORIVA) = CInt(Not CBool(v2 And &H80800000))
        .ColData(COL_VALIVA) = CInt(Not CBool(v2 And &H8100000))
        .ColData(COL_UNIDAD) = CInt(Not CBool(v2 And &H82000000))   '*** MAKOTO 22/jul/00
        .ColData(COL_OBS) = CInt(Not (CBool(v2 And &H84000000) Or CBool(v3 And &H84000000)))        'jeaa 22/04/2005
        .ColData(COL_TENTREGA) = CInt(Not (CBool(v2 And &H88000000) Or CBool(v3 And &H88000000)))        'jeaa 22/09/2005
        'jeaa 23/05/2008
        If Not gobjMain.GrupoActual.PermisoActual.CatInventarioCostoVer Then
            .ColHidden(COL_CU) = True
            .ColHidden(COL_CUR) = True
            .ColHidden(COL_CT) = True
            .ColHidden(COL_CTR) = True
            .ColHidden(COL_UTIL) = True
        End If
        
        
        
    End With
End Sub

Private Sub ConfigColsFormato()
    With grd
        .ColFormat(COL_EXIST) = mobjGNComp.Empresa.GNOpcion.FormatoCantidad
        .ColFormat(COL_CANT) = .ColFormat(COL_EXIST)
        
        '*** MAKOTO 31/ene/01 Mod. para aplicar formato de costo
        .ColFormat(COL_CU) = mobjGNComp.FormatoCosto
        .ColFormat(COL_CUR) = .ColFormat(COL_CU)
        .ColFormat(COL_CT) = .ColFormat(COL_CU)
        .ColFormat(COL_CTR) = .ColFormat(COL_CU)
        
        .ColFormat(COL_UTIL) = "#,0.00"             '*** MAKOTO 30/nov/00 Modificado #,#.00 --> #,0.00
        .ColFormat(COL_PU) = mobjGNComp.FormatoPU   '*** MAKOTO 13/feb/01 Mod.
        .ColFormat(COL_PUIVA) = mobjGNComp.FormatoPU   '******** jeaa 22-Dic-03
        .ColFormat(COL_PUR) = .ColFormat(COL_PU)    '***    "
        .ColFormat(COL_PT) = mobjGNComp.FormatoMoneda
        .ColFormat(COL_PTR) = .ColFormat(COL_PT)
        .ColFormat(COL_PTIVA) = mobjGNComp.FormatoMoneda '******** jeaa 22-Dic-03
        .ColFormat(COL_PORDCNT) = "#,0.00"          '*** MAKOTO 30/nov/00 Modificado
        .ColFormat(COL_PORIVA) = .ColFormat(COL_PORDCNT)
        
        '*** MAKOTO 13/feb/01 Mod. Valor de IVA usamos formato de PrecioTotal siempre
        ' independientemente de formato de costo
        '.ColFormat(COL_VALIVA) = .ColFormat(COL_PT) 'jeaa 09/03/2007 para que sea siempre con cuatro decimales
        .ColFormat(COL_VALIVA) = "#,0.0000" '.ColFormat(COL_PT)
        
        .ScrollBars = flexScrollBarBoth
        .Refresh
    End With
End Sub



Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Sub Refresh()
    If mobjGNComp Is Nothing Then Exit Sub
    'Cuando es solo ver, deshabilita grid
    If mobjGNComp.SoloVer Then
        grd.Editable = flexEDNone
        Exit Sub
    Else
        grd.Editable = flexEDKbdMouse
    End If
    
    'Actualiza lista de bodegas

    If grd.ColData(COL_CODBODEGA) = 0 Then
        grd.ColComboList(COL_CODBODEGA) = mobjGNComp.Empresa.ListaIVBodegaParaFlexGrid
    End If
    'Actualiza lista de Grupo1,2,3,4,5
    If grd.ColData(COL_G1) = 0 Then
        grd.ColComboList(COL_G1) = mobjGNComp.Empresa.ListaIVGrupoParaFlexGrid(1)
    End If
    If grd.ColData(COL_G2) = 0 Then
        grd.ColComboList(COL_G2) = mobjGNComp.Empresa.ListaIVGrupoParaFlexGrid(2)
    End If
    If grd.ColData(COL_G3) = 0 Then
        grd.ColComboList(COL_G3) = mobjGNComp.Empresa.ListaIVGrupoParaFlexGrid(3)
    End If
    If grd.ColData(COL_G4) = 0 Then
        grd.ColComboList(COL_G4) = mobjGNComp.Empresa.ListaIVGrupoParaFlexGrid(4)
    End If
    If grd.ColData(COL_G5) = 0 Then
        grd.ColComboList(COL_G5) = mobjGNComp.Empresa.ListaIVGrupoParaFlexGrid(5)
    End If
    
    'Si se muestra la columna de código alterno     '**** En BeforeEdit
'    If Not grd.ColHidden(COL_CODALT) Then
'        'Actualiza la lista de CodAlterno
'        grd.ColComboList(COL_CODALT) = mobjGNComp.Empresa.ListaIVCodAlternoParaFlex
'    End If
    
    'Llama a VisualizaTotal para que actualice valores prorrateados
    VisualizaTotal
    ConfigColsFormato       'Llama esta para actualizar formato de moneda

    '*** MAKOTO 30/nov/00 Agregado
    'Si no tiene permiso para modificar precios, desactiva el menú para grabar precios
     mnuGrabarPrecio.Enabled = gobjMain.GrupoActual.PermisoActual.CatInventarioPrecioMod And (Not grd.ColHidden(COL_PU))
'    mnuOptimizarCantidad.Enabled = Not grd.ColHidden(COL_EXIST)       '*** MAKOTO 16/dic/00
End Sub

'*** MAKOTO 12/ene/01 Agregado para permitir ordenar items por cualquier columna
Private Sub grd_AfterSort(ByVal COL As Long, Order As Integer)
    If Not mobjGNComp.SoloVer Then
        Aceptar             'Para re-asignar ordenes de detalles
    End If
End Sub

Private Sub grd_Click()
    RaiseEvent Click
End Sub

Private Sub grd_DblClick()
    Dim iv As IVinventario, CodItem As String, frm As IRegistro

    RaiseEvent DblClick

    If grd.Rows <= grd.FixedRows Then Exit Sub
    If grd.Row < grd.FixedRows Then Exit Sub            '*** MAKOTO 27/dic/00 Agregado

    '*** Oliver Si no tiene permisos para modificar no debe aparecer la ventana de item
    If Not gobjMain.GrupoActual.PermisoActual.CatInventarioMod Then Exit Sub
    
    On Error GoTo errtrap
    CodItem = grd.TextMatrix(grd.Row, COL_CODITEM)
    Set iv = Nothing
    Set frm = Nothing
    Exit Sub
errtrap:
    Set iv = Nothing
    Set frm = Nothing
    DispErr
    Exit Sub
End Sub

'*** MAKOTO 12/dic/00 Agregado
Private Sub grd_GotFocus()
    Dim Cancel As Boolean
    FlexGridGotFocusColor grd
    
    If grd.Editable And grd.Rows <= grd.FixedRows Then
        RaiseEvent AgregarFilaAuto(Cancel)  'Pregunta al contenedor si permite agregar la primera fila automáticamente o no
        If Not Cancel Then AgregaFila       'Si dice que sí }, agrega la primera fila
    End If
End Sub

Private Sub grd_LostFocus()
    FlexGridGotFocusColor grd
End Sub

Private Sub grd_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim cad As String, cod As String, bodega As String, bod As String
    Dim MatItem As Variant, i As Integer, v As Variant
    Dim iv As IVinventario, Area As String, obs As String
    Dim Dim1 As Currency, Dim2 As Currency
    Dim PU As Currency
    If mobjGNComp Is Nothing Then Exit Sub
    If mobjGNComp.SoloVer Then Exit Sub
    'ReDim MatItem(100) As Variant
    Select Case KeyCode
    Case vbKeyInsert
        AgregaFila
          PrecioSelec = 0
        grd.SetFocus
        KeyCode = 0
    Case vbKeyDelete
        EliminaFila
        grd.SetFocus
        KeyCode = 0
    Case vbKeyReturn
        'Cuando aplasta CTRL+ENTER
        If (Shift And vbCtrlMask) Then
            grd_DblClick     'Abre el registro del Item
        Else
            'jeaa 19/09/2008
            
            If Not grd.IsSubtotal(grd.Row) Then
            If grd.ColData(grd.COL) = 0 And grd.COL = COL_DESC Then
              'Modificado para la pantalla de busqueda tiene parametro de que bodega
              '
              If Not grd.ColHidden(COL_CODBODEGA) Then
                bodega = grd.TextMatrix(grd.Row, COL_CODBODEGA)
              Else
                'jeaa 27/04/04 para visualizar todas las bodegas
                If mobjGNComp.GNTrans.IVVisualizaTodaBodega Then
                    bodega = "%" '
                Else
                    bodega = BodegaPredeterminada
                End If
              End If
              grd.SetFocus
            End If
        End If
        End If
    Case 84 ' totaliza items
        If mobjGNComp.GNTrans.codPantalla <> "IVBQDPR" Then
        'Cuando aplasta CTRL+T
            If (Shift And vbCtrlMask) Then
                TotalizarItem
                KeyCode = 0
            End If
        End If
    Case 75  'kardex de item
        If mobjGNComp.GNTrans.codPantalla = "IVBQDPR" Then
            If grd.Cell(flexcpBackColor, grd.Row, COL_CODBODEGA, grd.Row, COL_TENTREGA) <> &H80000013 Then
                
            End If
        End If
    Case TECLA_CLICKDERECHO                     '*** MAKOTO 30/nov/00
        grd_MouseDown vbRightButton, Shift, 0, 0
    End Select

    RaiseEvent KeyDown(KeyCode, Shift)
End Sub


Private Sub grd_KeyPress(KeyAscii As Integer)
'   Dim CAD  As String
'   If KeyAscii = vbKeyReturn Then
'      If grd.ColData(grd.Col) = 0 And grd.Col = COL_DESC Then
'        CAD = frmBDescripcion.Inicio(grd.TextMatrix(grd.Row, grd.Col), mobjGNComp)
'        If Len(CAD) Then
'            grd.TextMatrix(grd.Row, grd.Col) = CAD
'            grd_AfterEdit grd.Row, grd.Col
'            grd_ValidateEdit grd.Row, grd.Col, False
'        End If
'
'      End If
'    End If
    
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub grd_KeyPressEdit(ByVal Row As Long, ByVal COL As Long, KeyAscii As Integer)
    Dim NoNeg As Boolean
    
'*** MAKOTO 27/ene/01 Eliminado, porque creó la propiedad 'IVPermitirSignoNegativo'
'    'Sólo cuando es transferencia permitimos el sigono '-'      '*** MAKOTO 15/oct/00
'    If mobjGNComp.GNTrans.IVTipoTrans = "T" _
'        And Not mTransBodega Then               '*** MAKOTO 09/nov/00 Modificado para transferencia bodega
'        NoNeg = False
'    Else
'        '*** MAKOTO 25/ene/01 Mod.
''        NoNeg = True
        NoNeg = Not mobjGNComp.GNTrans.IVPermitirSignoNegativo
'    End If
    
    '*** MAKOTO 03/oct/2000
    ValidarTeclaFlexGrid grd, Row, COL, KeyAscii, NoNeg
  
    
End Sub

Private Sub grd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mobjGNComp Is Nothing Then Exit Sub
    If mobjGNComp.SoloVer Then Exit Sub
    numfila = grd.Row
    If Not gobjMain.GrupoActual.PermisoActual.CatInventarioMod Then
        mnucrearitem.Enabled = False
    End If
    If mobjGNComp.GNTrans.codPantalla = "IVBQDPR" Then
        mnuTotalizar.Visible = False
    End If
    If mobjGNComp.GNTrans.codPantalla <> "IVBQDPR" Then
            mnuKardexItemCli.Visible = False
    Else
        If grd.Cell(flexcpBackColor, grd.Row, COL_CODBODEGA, grd.Row, COL_TENTREGA) = &H80000013 Then
            mnuKardexItemCli.Visible = False
        Else
            mnuKardexItemCli.Visible = True
        End If
    End If
    
    
    If Button And vbRightButton Then
        UserControl.PopupMenu mnuDetalle, , x, y
    End If
End Sub


Private Sub grd_AfterEdit(ByVal Row As Long, ByVal COL As Long)
    Dim obj As IVKardex, cod As String, bod As String
    Dim itemCom As IVinventario, i As Integer, mover As Boolean
    Dim itemP As IVinventario, j As Long, k As Long  'AUC para item de promocion
    Dim auxCant As Currency
    Dim codSub As String, cantSub As Currency, xCuanto As Currency, Precio As Currency, descriSub As String

    On Error GoTo errtrap

    If Not IsObject(grd.RowData(Row)) Then Exit Sub
    mover = True
    With grd
        Set obj = .RowData(Row)
        Select Case COL
        Case COL_CODBODEGA
            'Visualiza la existencia de la bodega seleccionada
            If Len(.TextMatrix(Row, COL_CODBODEGA)) > 0 Then
                bod = .TextMatrix(Row, COL_CODBODEGA)
                VisualizaItem Row, .TextMatrix(Row, COL_CODITEM), bod
            Else
                VisualizaItem Row, .TextMatrix(Row, COL_CODITEM)
            End If
            obj.CodBodega = Trim$(.Text)
'            itemCom = mobjGNComp.Empresa.RecuperaIVInventario(grd.TextMatrix(grd.Row, COL_CODITEM))
'            VerificarExistencia Row, itemCom
        Case COL_G1, COL_G2, COL_G3, COL_G4, COL_G5
            BorraItem Row
        Case COL_CODITEM
            obj.CodInventario = Trim$(.Text)
            '----------Agregado AUC 14/12/05
            If mobjGNComp.GNTrans.TotalizaItemsRepetidos Then
                TotalizarItem
            End If
            If mobjGNComp.GNTrans.IVAvanceLinea Then
                If Len(obj.CodInventario) > 0 Then
                'Si la fila actual es la última fila, inserta  automaticamente una fila
                    For i = grd.FixedRows To grd.Rows - 1       'Busca primera fila de SubTotal
                        If grd.IsSubtotal(i) Then Exit For
                    Next i
                    .Row = i
                    'If mobjGNComp.GNTrans.IVAvanceLinea Then
'                    TotalizarItem
                    AgregaFila
                    'jeaa se coloca siempre en la ultima fila insertada
                    mover = False
                Else
                    mover = True
                End If
            End If
        
        Case COL_CODALT
            obj.CodInventario = Trim$(.TextMatrix(Row, COL_CODITEM))
            '----------Agregado AUC 14/12/05
            If mobjGNComp.GNTrans.TotalizaItemsRepetidos Then
                TotalizarItem
            End If
            If mobjGNComp.GNTrans.IVAvanceLinea Then
                If Len(obj.CodInventario) > 0 Then
                'Si la fila actual es la última fila, inserta  automaticamente una fila
                    For i = grd.FixedRows To grd.Rows - 1       'Busca primera fila de SubTotal
                        If grd.IsSubtotal(i) Then Exit For
                    Next i
                    .Row = i
                    'If mobjGNComp.GNTrans.IVAvanceLinea Then
'                    TotalizarItem
                    AgregaFila
                    'jeaa se coloca siempre en la ultima fila insertada
                    mover = False
                Else
                    mover = True
                End If
            End If
        Case COL_DESC
            If obj.CodInventario <> "-" And obj.CodInventario <> "+" Then            '*** MAKOTO 16/oct/00
                cod = CogeSoloCodigo(Trim$(.Text))    'coge  en  este formato XXXX[codigo]
                'jeaa 18/08/04
                If mobjGNComp.GNTrans.IVTipoTrans = "E" Or mobjGNComp.GNTrans.IVTipoTrans = "Y" Then
                    bod = CogeSoloBodega(Trim$(.Text))    'coge  en  este formato XXXX¨{codbodega}
                End If
                If Len(cod) > 0 Then                    '*** MAKOTO 14/dic/00 Corregido
                    obj.CodInventario = cod
                End If
                If Len(.TextMatrix(Row, COL_CODBODEGA)) > 0 Then
                    bod = .TextMatrix(Row, COL_CODBODEGA)
                    obj.CodBodega = bod
                End If
            Else
                .TextMatrix(Row, COL_DESC) = obj.Nota
            End If
            '----------Agregado AUC 14/12/05
            If mobjGNComp.GNTrans.TotalizaItemsRepetidos Then
                TotalizarItem
            End If
            If mobjGNComp.GNTrans.IVAvanceLinea Then
                If Len(obj.CodInventario) > 0 Then
                'Si la fila actual es la última fila, inserta  automaticamente una fila
                    For i = grd.FixedRows To grd.Rows - 1       'Busca primera fila de SubTotal
                        If grd.IsSubtotal(i) Then Exit For
                    Next i
                    .Row = i
                    'If mobjGNComp.GNTrans.IVAvanceLinea Then
'                    TotalizarItem
                    AgregaFila
                    'jeaa se coloca siempre en la ultima fila insertada
                    mover = False
                Else
                    mover = True
                End If
            End If
        
        
        Case COL_CANT
            'Para que recalcule el costo. (Cuando es FIFO,LIFO importa la cantidad)
            'Y para que haga la verificación de cantidad límite de items    '*** MAKOTO 15/oct/00
            '************ NO SE QUIEN PUSO ESTO AQUI jeaa 01/09/2008
           ' If grd.Cell(flexcpBackColor, grd.Row, COL_CODITEM, grd.Row, COL_CODITEM) <> &H80000018 Then
           '*******************************************
            bod = CogeSoloBodega(.TextMatrix(Row, COL_DESC))    'coge  en  este formato XXXX¨{codbodega}
            If Len(.TextMatrix(Row, COL_CODBODEGA)) > 0 Then
                bod = .TextMatrix(Row, COL_CODBODEGA)
            End If
            If Len(bod) > 0 Then
                If Not VisualizaItem(Row, .TextMatrix(Row, COL_CODITEM), bod) Then
    '                .TextMatrix(Row, COL_CANT) = 0      'Borra la cantidad si está mal
                End If
            Else
                If Not VisualizaItem(Row, .TextMatrix(Row, COL_CODITEM)) Then
    '                .TextMatrix(Row, COL_CANT) = 0      'Borra la cantidad si está mal
                End If
            
            End If
            'jeaa 06/04/2005 para control de venta en fracciones
            If (Int(.TextMatrix(Row, COL_CANT)) - .TextMatrix(Row, COL_CANT)) <> 0 Then
               'Recupera el item
                Set itemCom = mobjGNComp.Empresa.RecuperaIVInventario(grd.TextMatrix(grd.Row, COL_CODITEM))
                If Not (itemCom Is Nothing) Then
                    If itemCom.bandFraccion = False Then
                        MsgBox "Este Item solo solo puede facturarse en unidades "
                        If Int(.TextMatrix(Row, COL_CANT)) = 0 Then
                            .TextMatrix(Row, COL_CANT) = 1
                        Else
                            .TextMatrix(Row, COL_CANT) = Int(.TextMatrix(Row, COL_CANT))
                        End If
                    End If
                End If
                Set itemCom = Nothing
            End If
'                'jeaa 14/09/2005 se visualizara pantalla para digitar largo y ancho y calculara el area
'                Set itemCom = mobjGNComp.Empresa.RecuperaIVInventario(obj.CodInventario)
'                If Not (itemCom Is Nothing) Then
'                    If itemCom.BandArea Then
'                            obj.Nota = .TextMatrix(Row, COL_OBS)
'                    End If
'                End If
            CalculaDetalle Row, COL
            'AUC Despliega las promociones----------------
        'ITEMS PROMOCION
        If mobjGNComp.GNTrans.IVVisItemPromocion Then
            If obj.CodInventario <> "" Then
               Set itemP = mobjGNComp.Empresa.RecuperaIVInventario(obj.CodInventario) 'ecupera el item padre
                If Not (itemP Is Nothing) Then
                    If itemP.tipo = Promocion Then
                        If itemP.NumFamiliaDetalle > 0 Then
                            For j = 1 To itemP.NumFamiliaDetalle 'ciclo para detalles de items
                                xCuanto = itemP.RecuperaDetalleFamilia(j).xCuanto
                                If .ValueMatrix(Row, COL_CANT) >= xCuanto Then
                                    AgregaFilaPromocion
                                    
                                    codSub = itemP.RecuperaDetalleFamilia(j).CodInventario
                                    descriSub = itemP.RecuperaDetalleFamilia(j).Descripcion
                                    cantSub = itemP.RecuperaDetalleFamilia(j).Cantidad
                                    Precio = itemP.RecuperaDetalleFamilia(j).Precio
                                    
                                    auxCant = 0
                                    For k = xCuanto To .ValueMatrix(Row, COL_CANT) Step xCuanto
                                        auxCant = auxCant + cantSub
                                    Next k
                                    If mobjGNComp.GNTrans.IVTipoTrans = "E" Or mobjGNComp.GNTrans.IVTipoTrans = "Y" Then
                                        auxCant = auxCant * -1
                                    End If
                                     AgregaItemsPromocion Row + j, COL_DESC, descriSub & "[" & codSub & "]{" & .TextMatrix(Row, COL_CODBODEGA) & "}", auxCant, Precio
                                     grd.TextMatrix(grd.Row, COL_CODBODEGA) = bod
                                     grd.TextMatrix(grd.Row, COL_PU) = Precio
                                     .Cell(flexcpBackColor, .Row, COL_CODBODEGA, .Row, COL_TENTREGA) = &H80000013    'CELESTE
                                Else
                                    Exit For
                                End If
                            Next j
                        End If
                    End If
                End If
                Set itemP = Nothing
            End If
        End If
'        Else
'            CalculaDetalle Row, COL
'            grd.COL = grd.COL + 1
'        End If
            
            '------------------------------
        Case COL_CU, COL_CT, COL_PU, COL_PT, COL_PORDCNT, COL_UTIL, COL_PUIVA, COL_PTIVA
            'si es Descuento por Item
            If Not grd.ColHidden(COL_PORDCNT) Then
                    If COL = COL_PU Then
                        Set itemCom = mobjGNComp.Empresa.RecuperaIVInventario(.TextMatrix(Row, COL_CODITEM))
                        If Not itemCom Is Nothing Then
                            If grd.ComboIndex > -1 Then
                                grd.TextMatrix(Row, COL_PORDCNT) = itemCom.Descuento(grd.ComboIndex + 1) * 100
                                grd.Cell(flexcpData, Row, COL_PORDCNT) = itemCom.Descuento(grd.ComboIndex + 1) * 100
                            End If
                            Set itemCom = Nothing
                        End If
                    End If
                'jeaa 23-03-04
                'si esta activo el limite
                If COL = COL_PORDCNT And mobjGNComp.Empresa.GNOpcion.ObtenerValor("LimiteDescto") = "1" Then
                    'jeaa 21/09/05 descuento por Cliente/Grupo Item 0 es por item 1 por clieente grupo
                    If Val(mobjGNComp.Empresa.GNOpcion.ObtenerValor("TipoDescuentoItems")) = 1 Then
                        Dim AuxDesct As Double, ivgrupo As Integer
                        Set itemCom = mobjGNComp.Empresa.RecuperaIVInventario(.TextMatrix(Row, COL_CODITEM))
                        If Not itemCom Is Nothing Then
                            ivgrupo = CInt(mobjGNComp.Empresa.GNOpcion.ObtenerValor("DescuentoNumIVGrupo")) + 1
                            AuxDesct = mobjGNComp.Empresa.DesctoPCxIV(itemCom.CodGrupo(ivgrupo), mobjGNComp.CodClienteRef)
                            If CDbl(.ValueMatrix(Row, COL_PORDCNT)) > AuxDesct Then
                                    If Not mobjGNComp.GNTrans.IVOmitirPorcentajeDesct Then
                                        MsgBox "No puede dar un descuento mayor al " & CDbl(AuxDesct) & "% a este Item.", vbExclamation
                                        .TextMatrix(Row, COL_PORDCNT) = AuxDesct
                                    Else
                                        MsgBox "El Porcentaje de Descuento exede al límite asignado en este Item.", vbExclamation
                                        .TextMatrix(Row, COL_PORDCNT) = "0"
                                    End If
                                    
                            End If
                        End If
                        Set itemCom = Nothing
                    Else
                        If COL = COL_PORDCNT And mobjGNComp.Empresa.GNOpcion.ObtenerValor("LimiteDescto") = "1" Then
                            If CDbl(.ValueMatrix(Row, COL_PORDCNT)) > CDbl(grd.Cell(flexcpData, Row, COL_PORDCNT)) Then
                                    MsgBox "No puede dar un descuento mayor al " & CDbl(grd.Cell(flexcpData, Row, COL_PORDCNT)) & "% a este Item.", vbExclamation
                                    .TextMatrix(Row, COL_PORDCNT) = grd.Cell(flexcpData, Row, COL_PORDCNT)
                            End If
                        End If
                    End If
                End If
            End If
            CalculaDetalle Row, COL
        '********* jeaa 22/04/2005
        Case COL_OBS
            If obj.CodInventario <> "-" Then
                obj.Nota = .TextMatrix(Row, COL_OBS)
            End If
        '********* jeaa 22/09/2005
        Case COL_TENTREGA
            If obj.CodInventario <> "-" Then
                obj.TiempoEntrega = .TextMatrix(Row, COL_TENTREGA)
            End If
        
        End Select
'        .ComboList = ""     'Limpia combo para que no se quede el boton de DropDown
    End With

    VisualizaTotal
    
    If mobjGNComp.GNTrans.IVAvanceLinea Then
        If mover = True Then
            MueveColumna
        End If
    Else
        MueveColumna
    End If
    
'    RaiseEvent DespuesdeEditarGrd
    Exit Sub
errtrap:
    DispErr
    Exit Sub
End Sub

Private Sub grd_CellChanged(ByVal Row As Long, ByVal COL As Long)
    '*** MAKOTO 29/ene/01 Agregado.
    FlexGridRedondear grd, Row, COL
End Sub

Private Function CogeSoloCodigo(Desc As String) As String
    Dim s As String, i As Long, f As Long
    i = InStrRev(Desc, "[")
    f = InStrRev(Desc, "]")
    If i > 0 Then s = Mid$(Desc, i + 1, f - i)
    If Len(s) > 0 Then s = Left$(s, Len(s) - 1)
    CogeSoloCodigo = s
End Function

Private Sub CalculaDetalle(Row As Long, COL As Long)

    Dim cu As Currency, ct As Currency, PU As Currency, PT As Currency
    Dim cur As Currency, ctr As Currency, pur As Currency, ptr As Currency
    Dim poriva As Currency, cant As Currency, pordes As Currency
    Dim obj As IVKardex, signo As Integer, ut As Single, PUIVA As Currency
    Dim PTIVA As Currency
    Dim pc As PCProvCli
    With grd
        If mobjGNComp.GNTrans.IVTipoTrans = "E" Or mobjGNComp.GNTrans.IVTipoTrans = "Y" Then
            signo = -1
        Else
            signo = 1
        End If
    
        'signo = IIf(mobjGNComp.GNTrans.IVTipoTrans = "E", -1, 1) '-1 si es egreso
        If mTransBodega Then signo = -1                 '*** MAKOTO 09/nov/00 Origen es siempre negativo
'        cant = .ValueMatrix(Row, COL_CANT) * signo
'        cu = .ValueMatrix(Row, COL_CU)
'        cur = .ValueMatrix(Row, COL_CUR)
'        ct = .ValueMatrix(Row, COL_CT) * signo
'        pu = .ValueMatrix(Row, COL_PU)
'        pt = .ValueMatrix(Row, COL_PT) * signo
'        pordes = .ValueMatrix(Row, COL_PORDCNT)
'        poriva = .ValueMatrix(Row, COL_PORIVA)
'        ut = .ValueMatrix(Row, COL_UTIL) / 100
        cant = MiCCur(.Cell(flexcpTextDisplay, Row, COL_CANT)) * signo
        cu = MiCCur(.Cell(flexcpTextDisplay, Row, COL_CU))
        cur = MiCCur(.Cell(flexcpTextDisplay, Row, COL_CUR))
        ct = MiCCur(.Cell(flexcpTextDisplay, Row, COL_CT)) * signo
        PU = MiCCur(.Cell(flexcpTextDisplay, Row, COL_PU))
        PT = MiCCur(.Cell(flexcpTextDisplay, Row, COL_PT)) * signo
        pordes = MiCCur(.Cell(flexcpTextDisplay, Row, COL_PORDCNT))
        poriva = MiCCur(.Cell(flexcpTextDisplay, Row, COL_PORIVA))
        ut = MiCCur(.Cell(flexcpTextDisplay, Row, COL_UTIL)) / 100
        PUIVA = MiCCur(.Cell(flexcpTextDisplay, Row, COL_PUIVA))          '******** jeaa 22-Dic-03
        PTIVA = MiCCur(.Cell(flexcpTextDisplay, Row, COL_PTIVA)) * signo  '******** jeaa 22-Dic-03
        Select Case COL
        Case COL_CANT
            ct = cu * cant
            PT = PU * cant
        Case COL_CU
            ct = cu * cant
        Case COL_CT
            If cant Then cu = ct / cant Else cu = 0
        Case COL_PU
            PT = PU * cant
        Case COL_PT
            If cant Then PU = PT / cant Else PU = 0
        Case COL_UTIL
            PU = cur * (1# + ut)
            PT = PU * cant
        Case COL_PTIVA  '******** jeaa 22-Dic-03
            If cant Then PUIVA = PTIVA / cant Else PUIVA = 0
        End Select
        
        Set obj = .RowData(Row)
        
        obj.Cantidad = cant
        obj.Descuento = pordes / 100
        obj.IVA = poriva / 100
'''        'jeaa 17/01/2008 SRI EQUIDAD TRIBUTARIA
'''        Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(mobjGNComp.IdClienteRef)
'''        If pc Is Nothing Then
'''            Exit Sub
'''        End If
'''        If mobjGNComp.Empresa.GNOpcion.ObtenerValor("OmitirIVA") = "1" And mobjGNComp.GNTrans.IVOmitirIVA And pc.BandEmpresaPublica Then
'''            obj.IVA = 0
'''        Else
'''            obj.IVA = poriva / 100
'''        End If
        obj.CostoTotal = ct
        obj.PrecioTotal = PT
        
        'Graba en el objeto la nota libre           '*** MAKOTO 16/oct/00
        If .TextMatrix(Row, COL_CODITEM) = "-" Then
            obj.Nota = .TextMatrix(Row, COL_DESC)
        End If
        
        'Graba en el objeto la observacion libre           '*** jeaa 25/04/2005
        If .TextMatrix(Row, COL_CODITEM) = "+" Then
            obj.Nota = .TextMatrix(Row, COL_DESC)
        End If
        Set pc = Nothing
    End With
End Sub

Private Sub BorraItem(Row As Long)
    Dim i As Long
    
    With grd
        For i = COL_CODITEM To .Cols - 1
            If i <> COL_CANT Then           'No se borra la cantidad
                .TextMatrix(Row, i) = ""
            End If
        Next i
    End With
End Sub

Private Sub VisualizaTotal()
    Dim i As Long, obj As IVKardex, cot As Double
    Dim por As Double, bandCalculado As Boolean, signo As Currency
    Dim itemCom As IVinventario, j As Integer
    
    If (Not mobjGNComp.SoloVer) And (mobjGNComp.Modificado) Then
        'Prorratea los recargos que deben ser prorrateado
        mobjGNComp.ProrratearIVKardexRecargo
    End If
    
    cot = mobjGNComp.Cotizacion("")
    'signo = IIf(mobjGNComp.GNTrans.IVTipoTrans = "E", -1, 1) '-1 si es egreso
    If mobjGNComp.GNTrans.IVTipoTrans = "E" Or mobjGNComp.GNTrans.IVTipoTrans = "Y" Then
        signo = -1
    Else
        signo = 1
    End If

    If Me.TransBodega Then signo = -1
    
    With grd
        For i = .FixedRows To .Rows - 1
            If Not .IsSubtotal(i) Then
                If Not IsEmpty(.RowData(i)) Then    '*** MAKOTO 14/sep/00
                    Set obj = .RowData(i)
                    .TextMatrix(i, COL_CU) = obj.costo
                    .TextMatrix(i, COL_CUR) = obj.CostoReal
                    .TextMatrix(i, COL_UTIL) = CalculaUtilidad(obj)
                    .TextMatrix(i, COL_PU) = obj.Precio
'                    .TextMatrix(i, COL_PUR) = Abs(obj.PrecioReal)
                    .TextMatrix(i, COL_PUR) = obj.PrecioReal       '*** MAKOTO 20/ene/01 Mod.
                    .TextMatrix(i, COL_PUIVA) = obj.PrecioReal + (obj.PrecioReal * obj.IVA) ' ******** jeaa 22-Dic-03
                    'Visualiza Costo y Precio con signos
                    .TextMatrix(i, COL_CT) = obj.CostoTotal * signo
                    .TextMatrix(i, COL_CTR) = obj.CostoRealTotal * signo
                    .TextMatrix(i, COL_PT) = obj.PrecioTotal * signo
                    .TextMatrix(i, COL_PTR) = obj.PrecioRealTotal * signo
                    .TextMatrix(i, COL_PTIVA) = (obj.PrecioRealTotal + (obj.PrecioRealTotal * obj.IVA)) * signo ' ******** jeaa 22-Dic-03
                    
                    
                    '*** MAKOTO 13/dic/00       '*** MAKOTO 26/ene/01 Mod. Quitado ABS()
'                    .TextMatrix(i, COL_VALIVA) = Abs(obj.CalcularIvaItem(por, bandCalculado))
                    .TextMatrix(i, COL_VALIVA) = obj.CalcularIvaItem(por, bandCalculado) * signo
                    'jeaa 23-03-04
                    'esta seleccionado item
                    If Len(obj.CodInventario) > 0 Then
                        ' activado descuento por item
                        If Not grd.ColHidden(COL_PORDCNT) Then
                            'recupera el item para en base al precio escogido saber el % descuento
                            Set itemCom = mobjGNComp.Empresa.RecuperaIVInventario(obj.CodInventario)
                            For j = 1 To 5
                                'si el precio recuperado es igual al precio recupera de la misma fila el % descuento
                                If obj.Precio = itemCom.Precio(j) Then
                                    ' para saber el limite del % descuento
                                    grd.Cell(flexcpData, i, COL_PORDCNT) = itemCom.Descuento(j) * 100
                                    Set itemCom = Nothing
                                    Exit For
                                End If
                            Next j
                        End If
                    End If
                End If
            End If
        Next i
    
        .subtotal flexSTSum, -1, COL_CANT, , .BackColorFrozen, vbYellow, , " ", , True
        .subtotal flexSTSum, -1, COL_CT, , .BackColorFrozen, vbYellow, , " ", , True
        .subtotal flexSTSum, -1, COL_CTR, , .BackColorFrozen, vbYellow, , " ", , True
        .subtotal flexSTSum, -1, COL_PT, , .BackColorFrozen, vbYellow, , " ", , True
        .subtotal flexSTSum, -1, COL_PTR, , .BackColorFrozen, vbYellow, , " ", , True
        .subtotal flexSTSum, -1, COL_PTIVA, , .BackColorFrozen, vbYellow, , " ", , True  'JEAA
        .subtotal flexSTSum, -1, COL_VALIVA, , .BackColorFrozen, vbYellow, , " ", , True
        .Refresh
    End With
End Sub

Private Function CalculaUtilidad(obj As IVKardex) As Single
    Dim ut As Single
    If obj.CostoRealTotal <> 0 Then
        ut = (Abs(obj.PrecioRealTotal) - Abs(obj.CostoRealTotal)) _
                    / Abs(obj.CostoRealTotal) * 100
    End If
    CalculaUtilidad = ut
End Function

Private Sub MueveColumna()
    Dim c As Long
    With grd
        If .Rows > .FixedRows Then
            For c = .COL + 1 To .Cols - 1
                If .ColData(c) >= 0 And .ColWidth(c) > 0 And (Not .ColHidden(c)) Then
                    .COL = c
                    Exit Sub
                End If
            Next c
    
            If .Row < .Rows - 1 Then .Row = .Row + 1
    
            For c = .FixedCols To .Cols - 1
                If .ColData(c) >= 0 And .ColWidth(c) > 0 And (Not .ColHidden(c)) Then
                    .COL = c
                    Exit Sub
                End If
            Next c
        End If
    End With
End Sub

Private Sub grd_BeforeEdit(ByVal Row As Long, ByVal COL As Long, Cancel As Boolean)
    Static r_antes As Long, c_antes As Long
    On Error GoTo errtrap
    If mobjGNComp.SoloVer Then Exit Sub
    
    
    'Cuando es una columna no modificable
    If grd.Rows > grd.FixedRows Then
        Cancel = (grd.ColData(COL) < 0) Or grd.IsSubtotal(Row) Or grd.ColHidden(COL)
    Else
        Cancel = True
    End If
    If Cancel Then Exit Sub
    If grd.Cell(flexcpBackColor, Row, COL_CODBODEGA, Row, COL_TENTREGA) = &H80000013 Then
        Cancel = True
    End If
   
    If Row = r_antes And COL = c_antes Then Exit Sub    'Si no cambia sale
    r_antes = Row: c_antes = COL

    grd.ComboList = ""
    Select Case COL
    Case COL_CODBODEGA
        grd.EditMaxLength = 10
    Case COL_G1, COL_G2, COL_G3, COL_G4, COL_G5
        PreparaComboGrupo COL - COL_G1 + 1
        grd.EditMaxLength = 20
    Case COL_CODITEM
        'Prepara la lista de items
        If grd.ColData(COL_CODITEM) = 0 Then  '1 Solo editable  0 Seleccionable
          PreparaComboItem   'PROBLEMA DE TIEMPO DE RESPUESTA  DIEGO 02/2002
        End If
        grd.EditMaxLength = 20
    Case COL_CODALT
        'Si se muestra la columna de código alterno
        If grd.ColData(COL_CODALT) = 0 Then  '1 Solo editable  0 Seleccionable
          If Not grd.ColHidden(COL_CODALT) Then PreparaComboCodAlterno
        End If
        grd.EditMaxLength = 20
    Case COL_DESC
        'Si se muestra la columna de Descripcion
        'Diego YA NO
        'If Not grd.ColHidden(COL_DESC) Then frmBDescripcion.Inicio
        grd.EditMaxLength = 80
    Case COL_CANT                   '*** MAKOTO 06/feb/01 Modificado
        grd.EditMaxLength = 14  'Hasta 99,999,999,999,999 sucres
    Case COL_CU, COL_CT, COL_CUR, COL_CTR, COL_PU, COL_PT, COL_PUIVA, COL_PTIVA
        grd.EditMaxLength = 14  'Hasta 99,999,999,999,999 sucres
    Case COL_PORDCNT, COL_UTIL
        grd.EditMaxLength = 5
    'JEAA 22/04/2005
    Case COL_OBS
       grd.EditMaxLength = 80
    'JEAA 22/09/2005
    Case COL_TENTREGA
       grd.EditMaxLength = 80
    End Select
    
    'Prepara la lista de precios de venta
    If COL = COL_PU Then
        grd.ComboList = grd.Cell(flexcpData, Row, COL_PU)
''        'Si no está preparada la lista (en caso de modificación)
''        If Len(grd.ComboList) = 0 And Len(grd.TextMatrix(Row, COL_CODITEM)) Then
'        If Len(grd.TextMatrix(Row, COL_CODITEM)) > 0 Then
'            Dim iv As IVInventario
'            'Recupera el objeto IVInventario (item) y obtiene lista de precios
'            Set iv = mobjGNComp.Empresa.RecuperaIVInventario(grd.TextMatrix(Row, COL_CODITEM))
'            If Not (iv Is Nothing) Then
'                grd.ComboList = iv.ListaPrecioParaFlex(mobjGNComp)
''                grd.Cell(flexcpData, Row, COL_PU) = grd.ComboList
'            End If
'        End If
    End If
    If COL = COL_PUIVA Then
        grd.ComboList = grd.Cell(flexcpData, Row, COL_PUIVA)
    End If
   
    Exit Sub
errtrap:
    DispErr
    Exit Sub
End Sub

Private Sub PreparaComboItem()
    Dim codg1 As String, codg2 As String, codg3 As String, codg4 As String, codg5 As String, r As Long
    
    ''!#' Significa que no hay condición
    codg1 = "!#": codg2 = "!#": codg3 = "!#": codg4 = "!#": codg5 = "!#"
    With grd
        r = .Row
        If Not .ColHidden(COL_G1) Then codg1 = Trim$(.TextMatrix(r, COL_G1))
        If Not .ColHidden(COL_G2) Then codg2 = Trim$(.TextMatrix(r, COL_G2))
        If Not .ColHidden(COL_G3) Then codg3 = Trim$(.TextMatrix(r, COL_G3))
        If Not .ColHidden(COL_G4) Then codg4 = Trim$(.TextMatrix(r, COL_G4))
        If Not .ColHidden(COL_G5) Then codg5 = Trim$(.TextMatrix(r, COL_G5))
        .ComboList = mobjGNComp.Empresa.ListaIVItemParaFlex("", codg1, codg2, codg3, codg4, codg5)
    End With
End Sub

Private Sub PreparaComboCodAlterno()
    Dim codg1 As String, codg2 As String, codg3 As String, codg4 As String, codg5 As String, r As Long
    
    ''!#' Significa que no hay condición
    codg1 = "!#": codg2 = "!#": codg3 = "!#": codg4 = "!#": codg5 = "!#"
    With grd
        r = .Row
        If Not .ColHidden(COL_G1) Then codg1 = .TextMatrix(r, COL_G1)
        If Not .ColHidden(COL_G2) Then codg2 = .TextMatrix(r, COL_G2)
        If Not .ColHidden(COL_G3) Then codg3 = .TextMatrix(r, COL_G3)
        If Not .ColHidden(COL_G4) Then codg4 = .TextMatrix(r, COL_G4)
        If Not .ColHidden(COL_G5) Then codg5 = .TextMatrix(r, COL_G5)
        .ComboList = mobjGNComp.Empresa.ListaivCodAlternoParaFlexPorGrupo(codg1, codg2, codg3, codg4, codg5)
    End With
End Sub

Private Sub PreparaComboDescripcion()
    Dim sql As String, rs As Recordset
    'Saca todos  los  items  del  Inventario aumentar   existencia

'    sql = "SELECT Descripcion + ' [' + CodInventario + ']', " & _
          "ISNULL(( " & _
            "SELECT Exist from IvExist " & _
            "WHERE IvExist.IdInventario = IVInventario.IdInventario and " & _
            "IvExist.IdBodega = " & mobjGNComp.GNTrans.IdBodegaPre & " ),0)" & _
          "FROM IvInventario ORDER BY Descripcion"
'    sql = "SELECT Descripcion + ' [' + CodInventario + ']' " & _
'          "FROM IvInventario ORDER BY Descripcion"
'
'    Set rs = mobjGNComp.Empresa.OpenRecordset(sql)
'    frmBDescripcion.grdItem.SetData (MiGetRows(rs))
'    frmBDescripcion.fcb.ColWidth0 = 3300
'    frmBDescripcion.fcb.ColWidth1 = 100
    
End Sub

Private Sub PreparaComboGrupo(numg As Integer)
    Dim codg1 As String, codg2 As String, codg3 As String, codg4 As String, codg5 As String, r As Long
    With grd
        r = .Row
        If Not .ColHidden(COL_G1) Then codg1 = Trim$(.TextMatrix(r, COL_G1))
        If Not .ColHidden(COL_G2) Then codg2 = Trim$(.TextMatrix(r, COL_G2))
        If Not .ColHidden(COL_G3) Then codg3 = Trim$(.TextMatrix(r, COL_G3))
        If Not .ColHidden(COL_G4) Then codg4 = Trim$(.TextMatrix(r, COL_G4))
        If Not .ColHidden(COL_G5) Then codg5 = Trim$(.TextMatrix(r, COL_G5))
    End With
End Sub

Private Sub AgregaFila()
    Dim r As Long, r2 As Long, ix As Long, COL As Integer
    On Error GoTo errtrap

    'Verifica si ya está número maximo de filas
    If (mobjGNComp.GNTrans.IVNumFilaMax > 0) And _
        (mobjGNComp.CountIVKardex >= mobjGNComp.GNTrans.IVNumFilaMax) Then
        MsgBox "No se puede agregar más filas porque está limitado hasta " & _
         mobjGNComp.GNTrans.IVNumFilaMax & " filas." & vbCr & vbCr & _
        "Si hay más detalle de items, regístrelos en otro comprobante." & vbCr & _
        "Si quiere cambiar el límite, váyase a la configuración de la transacción, por favor.", vbInformation
        Exit Sub
    End If

    'Agrega nuevo objeto IVKardex al comprobante        '*** MAKOTO 14/oct/00 Modificado
    ix = mobjGNComp.AddIVKardex
    
    With grd
        r2 = .Rows - 1
        If .IsSubtotal(.Rows - 1) Then r2 = r2 - 1
        'Si no es la primera fila
        If r2 > 0 Then
            'Si no está en la fila de total
            If Not .IsSubtotal(.Row) Then
                .AddItem "", .Row + 1
                r = .Row + 1
            'Si está en la fila de total
            Else
                .AddItem "", .Row
                r = .Row
            End If
        'Si es la primera fila
        Else
            'Si no está en la fila de total
            If (.Row < .Rows - 1) Or (.Row = 0) Then
'            If Not .IsSubtotal(.Row) Then
                .AddItem ""
                r = .Rows - 1
            'Si está en la fila de total
            Else
                .AddItem "", .Row
                r = .Row
            End If
        End If
        numfila = r
        'Asigna el indice de nuevo objeto a la fila nueva
        .RowData(r) = mobjGNComp.IVKardex(ix)
        
        'Visualiza los valores predeterminados
        .TextMatrix(r, COL_CODBODEGA) = mobjGNComp.IVKardex(ix).CodBodega
        .TextMatrix(r, COL_CANT) = mobjGNComp.IVKardex(ix).Cantidad
        '.TextMatrix(r, COL_CANT) = mobjGNComp.GNTrans.CantidadPre
        
        'Copia de la fila anterior
        If r > .FixedRows Then
'            .TextMatrix(r, COL_CODBODEGA) = .TextMatrix(r - 1, COL_CODBODEGA)  '*** MAKOTO 16/dic/00 Eliminado por que lo hace en objeto GNComprobante.AddIVKardex
            .TextMatrix(r, COL_G1) = .TextMatrix(r - 1, COL_G1)
            .TextMatrix(r, COL_G2) = .TextMatrix(r - 1, COL_G2)
            .TextMatrix(r, COL_G3) = .TextMatrix(r - 1, COL_G3)
            .TextMatrix(r, COL_G4) = .TextMatrix(r - 1, COL_G4)
            .TextMatrix(r, COL_G5) = .TextMatrix(r - 1, COL_G5)
        End If
        
        'Ubica cursor en la primera columna visible y editable/seleccionable
        ' y vacio
        If .Rows > .FixedRows Then
            .Row = r
            .COL = .FixedCols
            For ix = .FixedCols To .Cols - 1
                If (Val(.ColData(ix)) >= 0) And _
                   (.ColWidth(ix) > 0) And _
                   (.ColHidden(ix) = False) And _
                   (Len(.TextMatrix(r, ix)) = 0) Then
                    .COL = ix
                    Exit For
                End If
            Next ix
        End If
    End With
    
    ' *********** jeaa 24-12-2003 para cambiar el fondo de las celdas que no son modificables
    For COL = 1 To 28 ' 27
            If grd.ColData(COL) = -1 Then
                grd.Cell(flexcpBackColor, grd.Row, COL, grd.Row, COL) = &H80000018
            Else
                grd.Cell(flexcpBackColor, grd.Row, COL, grd.Row, COL) = vbWhite
            End If
        Next
    PoneNumFila
    VisualizaTotal
    Exit Sub
errtrap:
    DispErr
    Exit Sub
End Sub

Private Sub EliminaFila()
    Dim msg As String, r As Long
    On Error GoTo errtrap

    
    If grd.Row < grd.FixedRows Then Exit Sub        '*** MAKOTO 07/feb/01 Mod.
    If grd.Rows <= grd.FixedRows Then Exit Sub
    If grd.IsSubtotal(grd.Row) Then Exit Sub
    
    'jeaa 10/09/2008
    r = grd.Row
    If mobjGNComp.GNTrans.IVNoEliminarItemsImportar Then
        msg = "La configuración de la Transacción no permite  eliminar la fila #" & r & "?"
        If MsgBox(msg, vbNo + vbQuestion) <> vbYes Then Exit Sub
    End If
    
    
'    r = grd.Row
    msg = "Desea eliminar la fila #" & r & "?"
    If MsgBox(msg, vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    'Remueve de la colección de objeto
    mobjGNComp.RemoveIVKardex 0, grd.RowData(r)
    
    'Elimina del grid
    grd.RemoveItem r
    PoneNumFila
    grd.subtotal flexSTClear
    VisualizaTotal
    
    Exit Sub
errtrap:
    DispErr
    Exit Sub
End Sub

Private Sub PoneNumFila()
    Dim i As Long
    With grd
        For i = .FixedRows To .Rows - 1
            If Not .IsSubtotal(i) Then .TextMatrix(i, 0) = i
        Next i
    End With
End Sub

Private Sub grd_SelChange()
'If grd.Cell(flexcpBackColor, grd.Row, COL_CODITEM, grd.Row, COL_CODITEM) <> &H80000018 Then
    
End Sub

Private Sub grd_StartEdit(ByVal Row As Long, ByVal COL As Long, Cancel As Boolean)
    If mobjGNComp.GNTrans.codPantalla = "IVCAMIEPRD" Then
        If grd.Cell(flexcpBackColor, Row, COL_CODITEM, grd.Row, COL_CODITEM) = &H80000018 Then
                MsgBox "No debe modificar este Item "
                grd.SetFocus
                Exit Sub
        End If
    End If
End Sub

Private Sub grd_ValidateEdit(ByVal Row As Long, ByVal COL As Long, Cancel As Boolean)
    Dim msg As String, cod As String, bod As String
'    Dim v2 As Long, v3 As Long

    With grd
        Select Case COL
        Case COL_CODITEM
            If Len(.TextMatrix(Row, COL_CODBODEGA)) > 0 Then
                    bod = .TextMatrix(Row, COL_CODBODEGA)
                    If Not VisualizaItem(Row, .EditText, bod) Then
                        .TextMatrix(Row, COL) = ""
                        Cancel = True
                    End If
            Else
                    If Not VisualizaItem(Row, .EditText) Then
                        .TextMatrix(Row, COL) = ""
                        Cancel = True
                    End If
            End If
        Case COL_CODALT
            If Not VisualizaCodAlt(Row, .EditText) Then
                .TextMatrix(Row, COL) = ""
                Cancel = True
            End If
        Case COL_DESC                   '*** MAKOTO 15/oct/00
            cod = CogeSoloCodigo(Trim$(.EditText))
            'jeaa 18/08/04
            If mobjGNComp.GNTrans.IVTipoTrans = "E" Or mobjGNComp.GNTrans.IVTipoTrans = "Y" Then
                bod = CogeSoloBodega(Trim$(.EditText))
            End If
            If Len(cod) > 0 Then
                'Visualiza la existencia de la bodega seleccionada
                If Len(bod) > 0 Then
                    If Not VisualizaItem(Row, cod, bod) Then
                        .TextMatrix(Row, COL) = ""
                        Cancel = True
                    End If
                Else
                    'jeaa 18/08/04
                    If Not VisualizaItem(Row, cod, grd.TextMatrix(Row, COL_CODBODEGA)) Then
                        .TextMatrix(Row, COL) = ""
                        Cancel = True
                    End If
                End If
            End If
        
        Case COL_CANT
            If Len(.EditText) > 0 Then
                If Not IsNumeric(.EditText) Then
                    MsgBox "Ingrese un valor numérico.", vbExclamation
                    .TextMatrix(Row, COL) = ""
                    Cancel = True
                End If
            End If
            
        Case COL_CU, COL_CT, COL_CTR, COL_PU, COL_PT, COL_PUIVA, COL_PTIVA
            If Len(.EditText) > 0 Then
                If Not IsNumeric(.EditText) Then
                    MsgBox "Ingrese un valor numérico.", vbExclamation
                    .TextMatrix(Row, COL) = ""
                    Cancel = True
                ElseIf CCur(.EditText) < 0 Then
                    '*** MAKOTO 26/ene/01 Mod.
                    If ((COL <> COL_CT) And (COL <> COL_PT) And (COL <> COL_PTIVA)) Or _
                        (Not mobjGNComp.GNTrans.IVPermitirSignoNegativo) Then
                        MsgBox "Ingrese un valor positivo.", vbExclamation
                        .TextMatrix(Row, COL) = ""
                        Cancel = True
                    End If
                End If
            End If
'            '***  0liver 28/01/03   Agregado para controlar el caso de que el precio se solo seleccionable
'            '                       revisar que escribio es un precio valido
'            If Col = COL_PU Then
'
'                v2 = mobjGNComp.GNTrans.ColEditable
'                v3 = mobjGNComp.GNTrans.ColSeleccionable
'                'Si  No Editable           And  'Seleccionable
'                If Not CBool(v2 And &H80010000) And CBool(v3 And &H80010000) Then
'
'                    If Not ExistePrecio(.EditText, .Cell(flexcpData, Row, Col), .TextMatrix(Row, COL_CODITEM)) Then
'                       MsgBox "No puede escribir un precio que no esta en la lista", vbInformation
'                       Cancel = True
'                    End If
'                End If
'            End If
        End Select
    End With
End Sub

Private Function VisualizaItem(Row As Long, cod As String, Optional CodBodega As String) As Boolean
    Dim item As IVinventario, c As Currency, P As Currency, msg As String, ListaPrecio As String
    Dim saldo As Currency, EncontroItemEnDocFuente As Boolean
    On Error GoTo errtrap

    If Len(cod) = 0 Then Exit Function
    
    MensajeStatus MSG_PREPARA, vbHourglass
    
    '********************************** VERIFICACION DE LIMITE DE CANTIDAD CON PRESPUESTO
    'Item con código '-' es especial
    If cod <> "-" And cod <> "+" Then
        'Verifica con el límite establecido         '*** MAKOTO 15/oct/00 Agregado
        If Not VerificarLimiteitem(cod, Row, saldo, "IVVerificaLimite", CodBodega) Then
            'Si está configurado para que no permita grabar superando el límite
            If mobjGNComp.GNTrans.IVVerificaLimiteNoGrabar Then   'AUC cambiado 26/11/07
                If saldo >= 0 Then
                    'Si hay saldo, modifica la cantidad
                    grd.TextMatrix(Row, COL_CANT) = saldo
                Else
                    grd.TextMatrix(Row, COL_CANT) = 0
                    'Si no hay saldo, no permite seleccionar ése item
                    VisualizaItem = False
                    MensajeStatus
                    Exit Function
                End If
            End If
        End If
    End If
    '**********************************
    
    '********************VERIFICACION DE LIMITE DE CANTIDAD CON TRANSFUENTE
    If cod <> "-" And cod <> "+" Then
        If Not VerificarLimiteitem(cod, Row, saldo, "IVVerificaItemFuente", CodBodega) Then
            If saldo > 0 Then
                'Si hay saldo, modifica la cantidad
                grd.TextMatrix(Row, COL_CANT) = saldo
            Else
                'Si no hay saldo, no permite seleccionar ése item
                VisualizaItem = False
                MensajeStatus
                Exit Function
            End If
        End If
    End If
    
    'Recupera el item seleccionado
        Set item = mobjGNComp.Empresa.RecuperaIVInventario(cod)
        With item
        If Not grd.ColHidden(COL_CODBODEGA) Then grd.TextMatrix(Row, COL_CODBODEGA) = CodBodega
        If Not grd.ColHidden(COL_G1) Then grd.TextMatrix(Row, COL_G1) = .CodGrupo(1)
        If Not grd.ColHidden(COL_G2) Then grd.TextMatrix(Row, COL_G2) = .CodGrupo(2)
        If Not grd.ColHidden(COL_G3) Then grd.TextMatrix(Row, COL_G3) = .CodGrupo(3)
        If Not grd.ColHidden(COL_G4) Then grd.TextMatrix(Row, COL_G4) = .CodGrupo(4)
        If Not grd.ColHidden(COL_G5) Then grd.TextMatrix(Row, COL_G5) = .CodGrupo(5)
        grd.TextMatrix(Row, COL_CODITEM) = .CodInventario
        grd.TextMatrix(Row, COL_CODALT) = .CodAlterno1
        If cod = "-" Or cod = "+" Then                       '*** MAKOTO 16/oct/00 Item '-' es especial
            If Len(grd.TextMatrix(Row, COL_DESC)) = 0 Then      'Sólo cuando no está ingresado nada
                grd.TextMatrix(Row, COL_DESC) = .Descripcion    '   visualizamos la descripcion
            End If
        Else
            grd.TextMatrix(Row, COL_DESC) = .Descripcion
        End If
        If Len(.CODUnidad) <> 0 Then
            grd.TextMatrix(Row, COL_UNIDAD) = .CODUnidad            '*** MAKOTO 22/jul/00
        Else
            grd.TextMatrix(Row, COL_UNIDAD) = .Unidad            '*** MAKOTO 22/jul/00
        End If
        grd.TextMatrix(Row, COL_PORIVA) = .PorcentajeIVA * 100
        
        'Si el C.U. No es modificable o está en 0, visualiza el costo calculado
        If (grd.ColHidden(COL_CU) = True) _
            Or (grd.ColData(COL_CU) < 0) _
            Or (grd.ValueMatrix(Row, COL_CU) = 0) Then
'            c = .costo(mobjGNComp.FechaTrans, grd.ValueMatrix(Row, COL_CANT), mobjGNComp.TransID)
            c = .CostoDouble2(mobjGNComp.FechaTrans, _
                                grd.ValueMatrix(Row, COL_CANT), _
                                mobjGNComp.TransID, _
                                mobjGNComp.HoraTrans)  '*** MAKOTO 08/dic/00 Agregado Hora
            
            'Si el costo calculado está en otra moneda, convierte en moneda de trans.
            If mobjGNComp.CodMoneda <> .CodMoneda Then
                c = c * mobjGNComp.Cotizacion(.CodMoneda) / mobjGNComp.Cotizacion("")
            End If
            grd.TextMatrix(Row, COL_CU) = c
        End If
        
        
        'Si P.U. no está oculto , visualiza el Precio1
        If (Not grd.ColHidden(COL_PU)) Then
'        'Si P.U. no está oculto Y está en 0, visualiza el Precio1
'        If (Not grd.ColHidden(COL_PU)) And (grd.ValueMatrix(Row, COL_PU) = 0) Then
            
            '***19/09/2003  oliver
            'Agregado condicion para el caso en el que es documento importado y el precio no es modificable
            ' debe respetar el precio del documento fuente
'''''            If grd.ColData(COL_PU) = -1 And mobjGNComp.IdTransFuente <> 0 Then
'''''                p = PrecioIVK_DocFuente(.CodInventario, EncontroItemEnDocFuente)
'''''            End If
'''''            If Not EncontroItemEnDocFuente Then   'Si no encuentra en item en Doc Fuete pone al precio predeterminado
'''''                If mobjGNComp.GNTrans.IVPrecioPre > 0 Then   ' Agregado Oliver
'''''                    p = .Precio(mobjGNComp.GNTrans.IVPrecioPre) ' para sacar el precio predeterminado
'''''                Else
'''''                    p = 0                                       ' en caso de no tener precio predeterminado
'''''                                                                ' no saca precio
'''''                End If
'''''            End If
'''''            p = p * mobjGNComp.Cotizacion(.CodMoneda) / mobjGNComp.Cotizacion("")  'Convierte en moneda del comprobante
            If mobjGNComp.GNTrans.codPantalla <> "IVDVBQD" Then
'                If mobjGNComp.PrecioRef = 0 Then
                If Not mobjGNComp.GNTrans.IVVisualizarNumPagos Then
                    grd.TextMatrix(Row, COL_PU) = precio_predeterminado(item, EncontroItemEnDocFuente)
                Else
                    grd.TextMatrix(Row, COL_PU) = Precio_PredeterminadoMasIntereses(item, EncontroItemEnDocFuente)
                End If
'                Else
'                    grd.TextMatrix(Row, COL_PU) = precio_predeterminado(item, EncontroItemEnDocFuente)
'                End If
            End If
        End If
                
        'Si P.U. no está oculto Y es modificable/seleccionable, guarda la lista de precios
        
        If (Not grd.ColHidden(COL_PU)) And (grd.ColData(COL_PU) >= 0) Then
            ListaPrecio = .ListaPrecioParaFlex(mobjGNComp)
            If grd.ColData(COL_PU) > 0 Then ListaPrecio = Mid$(ListaPrecio, 2) 'Quita el |  para que sea solo seleccionable
            grd.Cell(flexcpData, Row, COL_PU) = ListaPrecio
        End If


'--------------------------- Si P.UIVA. no está oculto , visualiza el Precio1
        If (Not grd.ColHidden(COL_PUIVA)) Then
'''''''        'Si P.U. no está oculto Y está en 0, visualiza el Precio1
''''''        '************************
''''''        'coloca el precio predetermnado en .precio
''''''        '************************
        If (grd.ColHidden(COL_PU)) Then
                'jeaa 02/09/2008
                If Not mobjGNComp.GNTrans.IVVisualizarNumPagos Then
                    grd.TextMatrix(Row, COL_PU) = precio_predeterminado(item, EncontroItemEnDocFuente)
                Else
                    grd.TextMatrix(Row, COL_PU) = Precio_PredeterminadoMasIntereses(item, EncontroItemEnDocFuente)
                End If
        End If
'''''''******************
            '***19/09/2003  oliver
            'Agregado condicion para el caso en el que es documento importado y el precio no es modificable
            ' debe respetar el precio del documento fuente
            If grd.ColData(COL_PUIVA) = -1 And mobjGNComp.IdTransFuente <> 0 Then
                P = PrecioIVK_DocFuente(.CodInventario, EncontroItemEnDocFuente)
            End If
            If Not EncontroItemEnDocFuente Then   'Si no encuentra en item en Doc Fuete pone al precio predeterminado
                If mobjGNComp.GNTrans.IVPrecioPre > 0 Then   ' Agregado Oliver

                        P = .Precio(mobjGNComp.GNTrans.IVPrecioPre) + (.Precio(mobjGNComp.GNTrans.IVPrecioPre) * .PorcentajeIVA) '*  para sacar el precio MAS iva
                Else
                    If PrecioSelec = 0 Then
                        P = 0                                       ' en caso de no tener precio predeterminado                                                                ' no saca precio
                    Else
                        P = PrecioSelec + (PrecioSelec * .PorcentajeIVA) '*  para sacar el precio MAS iva
                    End If
                    
                End If
            End If
            P = P * mobjGNComp.Cotizacion(.CodMoneda) / mobjGNComp.Cotizacion("")  'Convierte en moneda del comprobante

            grd.TextMatrix(Row, COL_PUIVA) = P
        End If
                
        '****************************** VISUALIZACION DE EXISTENCIA
        'Visualiza la existencia en la bodega seleccionada
        If mTransBodega And Len(mCodBodegaOrigen) > 0 Then                   '*** MAKOTO 14/nov/00
            'En caso de pantalla de transferencia, coge la existencia de bodega de orígen
            grd.TextMatrix(Row, COL_EXIST) = .Existencia(mCodBodegaOrigen)
        'Si columna de bodega está visible                  '*** MAKOTO 15/dic/00
        ElseIf Len(grd.TextMatrix(Row, COL_CODBODEGA)) > 0 Then
        'jeaa 28/04/04 modificado para que solo tome existencia de la bodega predeterminada
        'Not grd.ColHidden(COL_CODBODEGA) Then
        
            'Visualiza existencia de la bodega seleccionada
            grd.TextMatrix(Row, COL_EXIST) = .Existencia(grd.TextMatrix(Row, COL_CODBODEGA))
        'Si columna de bodega está oculta
        Else
            'Visualiza la suma de todas las bodegas         '*** MAKOTO 15/dic/00
            grd.TextMatrix(Row, COL_EXIST) = .Existencia("")
        End If
        '****************************** VISUALIZACION DE DESCUENTO X ITEM  ******** jeaa 22-Dic-03-01-12-03
        'jeaa 03/01/05 descuento por Cliente/Grupo Item 0 es por item 1 por clieente grupo
        If Val(mobjGNComp.Empresa.GNOpcion.ObtenerValor("TipoDescuentoItems")) = 1 Then
            Dim AuxDesct As Double, ivgrupo As Integer
            ivgrupo = CInt(mobjGNComp.Empresa.GNOpcion.ObtenerValor("DescuentoNumIVGrupo")) + 1
            If Not mobjGNComp.GNTrans.IVVisualizarNumPagos Then
                AuxDesct = mobjGNComp.Empresa.DesctoPCxIV(.CodGrupo(ivgrupo), mobjGNComp.CodClienteRef)
            Else
                If mobjGNComp.Estado1 = 0 Then
                    AuxDesct = mobjGNComp.Empresa.DesctoNumPagosxIV(.CodGrupo(ivgrupo), mobjGNComp.numdias)
                    '99= SIN DESCUENTO
                    If AuxDesct = 99 Then
                        AuxDesct = 0
                    End If
                Else
                    AuxDesct = 0
                End If
                
            End If
        Else
            'jeaa 10/02/05
            If Not grd.ColHidden(COL_PORDCNT) Then
                If mobjGNComp.GNTrans.IVPrecioPre > 0 Then
                    AuxDesct = .Descuento(mobjGNComp.GNTrans.IVPrecioPre) * 100
                Else
                    AuxDesct = 0
                End If
            End If
        End If
        If Not grd.ColHidden(COL_PORDCNT) Then 'Corregido 25/03/2004 jeaa
            'jeaa 03/10/2005
            If Not mobjGNComp.GNTrans.IVOmitirPorcentajeDesct Then
                If mobjGNComp.GNTrans.IVPrecioPre > 0 Then
                    grd.TextMatrix(Row, COL_PORDCNT) = AuxDesct
                Else
                    grd.TextMatrix(Row, COL_PORDCNT) = "0"
                End If
            Else
                grd.TextMatrix(Row, COL_PORDCNT) = "0"  'jeaa 03/10/2005
            End If
        End If
            
        'JEAA 05/01/2009
        ' VISUALIZA LA FECHA DE LA ULTIMA COMPRA
        If mobjGNComp.GNTrans.IVVerificaFechaCompra Then
            If Not grd.ColHidden(COL_TENTREGA) Then
                grd.TextMatrix(Row, COL_TENTREGA) = Format(.FechaUltimoIngreso, .Empresa.GNOpcion.FormatoFecha)
                If DateDiff("d", .FechaUltimoIngreso, mobjGNComp.FechaTrans) > Val(mobjGNComp.GNTrans.IVNumDiasVerificaFechaCompra) Then
                    grd.Cell(flexcpBackColor, Row, COL_TENTREGA, Row, COL_TENTREGA) = vbRed
                End If
            End If
        End If
            
            
'        Else
'            If Not grd.ColHidden(COL_PORDCNT) Then 'Corregido 25/03/2004 jeaa
'                If mobjGNComp.GNTrans.IVPrecioPre > 0 Then
'                    grd.TextMatrix(Row, COL_PORDCNT) = .Descuento(mobjGNComp.GNTrans.IVPrecioPre) * 100
'                Else
'                    grd.TextMatrix(Row, COL_PORDCNT) = "0"
'                End If
'            End If
'        End If
        
        '****************************** VERIFICACION DE EXISTENCIA NEGATIVA
        '*** MAKOTO 06/feb/01 Mod.
        VerificarExistencia Row, item
        '*** ANGEL 21/mar/03 Agregado
        VerificarExisMaxMin Row, item
        '****************************** INGRESO DE NOTA LIBRE
        '*** MAKOTO 16/oct/00
        'Cuando selecciona item '-', ingresa la nota libre
        If (.CodInventario = "-") And (grd.COL < COL_CANT) Then
            msg = grd.TextMatrix(Row, COL_DESC)
            Do
                msg = InputBox("Ingrese una nota", , msg)
                If Len(msg) > MAXLEN_NOTA Then
                    MsgBox "La longitud máxima de la nota es de " & _
                            MAXLEN_NOTA & " caracteres.", vbInformation
                    msg = Left$(msg, MAXLEN_NOTA)        'Automáticamente corta hasta MAXLEN_NOTA létras
                Else
                    Exit Do
                End If
            Loop
            grd.TextMatrix(Row, COL_DESC) = msg
        End If
        '****************************** INGRESO DE OBSERVACION LIBRE
        '*** JEAA 22/04/2005
        'Cuando selecciona item '+', ingresa la observacion libre
        If (.CodInventario = "+") And (grd.COL < COL_CANT) Then
            msg = grd.TextMatrix(Row, COL_DESC)
            Do
                msg = InputBox("Ingrese una observación", , msg)
                If Len(msg) > MAXLEN_NOTA Then
                    MsgBox "La longitud máxima de la nota es de " & _
                            MAXLEN_NOTA & " caracteres.", vbInformation
                    msg = Left$(msg, MAXLEN_NOTA)        'Automáticamente corta hasta MAXLEN_NOTA létras
                Else
                    Exit Do
                End If
            Loop
            grd.TextMatrix(Row, COL_DESC) = msg
        End If
        
        
        
        '******************************
        'jeaa 23-03-04 para guardad el descuento por item
        If Not grd.ColHidden(COL_PORDCNT) Then
            grd.Cell(flexcpData, Row, COL_PORDCNT) = grd.TextMatrix(Row, COL_PORDCNT)
        End If
        
        '*************************************
'         devolucion automatica
        If mobjGNComp.GNTrans.codPantalla = "IVDVBQD" Then
            If grd.TextMatrix(Row, COL_CANT) > 0 Then
                grd.TextMatrix(Row, COL_CANT) = grd.TextMatrix(Row, COL_CANT)
            End If
        End If
        
        
        
        '***Agregado. 05/ago/04. Angel. Para refrescar el recargo x item
        Dim obj As IVKardex
        Set obj = grd.RowData(Row)
        obj.ValorRecargoItem = .ValorRecargo
'        If mobjGNComp.GNTrans.IVVisItemPromocion Then
'            obj.IdPadre = item.RecuperaID(cod)
'        End If
        'Calcula detalles (PT,CT etc...)
        ' y graba la nota libre en la propiedad Nota        '*** MAKOTO 16/oct/00
        
        '***************************************************
        '***Agregado. 22/04/05. Esteban, campo observacion
        If Not grd.ColHidden(COL_OBS) Then
            grd.TextMatrix(Row, COL_OBS) = .Observacion
        End If
        '***************************************************
        
        
        CalculaDetalle Row, COL_CANT
       
        VisualizaItem = True
        
        'jeaa 06/04/2005
        'para control de venta en fracciones
        If (Int(grd.TextMatrix(Row, COL_CANT)) - grd.TextMatrix(Row, COL_CANT)) <> 0 Then
             If .bandFraccion = False Then
                MsgBox "Este Item solo solo puede facturarse en unidades "
                If Int(grd.TextMatrix(Row, COL_CANT)) = 0 Then
                    grd.TextMatrix(Row, COL_CANT) = 1
                Else
                    grd.TextMatrix(Row, COL_CANT) = Int(grd.TextMatrix(Row, COL_CANT))
                End If
            End If
        End If
            
        
    End With
    Set item = Nothing
    MensajeStatus
    Exit Function
errtrap:
    MensajeStatus
    'Si no encuentra el codigo de item
    If Err.Number = 3021 Or Err.Number = 91 Then
        MsgBox MSG_ERR_NOENCUENTRA & "(" & cod & ")", vbInformation
    Else
        DispErr
    End If
    Exit Function
End Function

'*** MAKOTO 06/feb/01 Agregado
Private Function VerificarExistencia( _
                    ByVal Row As Long, _
                    ByVal item As IVinventario) As Boolean
    Dim cant_ant As Currency, msg As String, cod_ant As String, exist As Currency
    Dim sumaCant As Currency, i As Long, codb As String, cod As String
    Dim codb_ant As String
    
    If (mobjGNComp.GNTrans.IVTipoTrans <> "T") Then
        'Si la transacción NO es egreso
        ' ó NO está configurado para verificar existencia negativa
        ' ó el item es de servicio, no hace la verificación
        If ((mobjGNComp.GNTrans.IVTipoTrans <> "E") And (mobjGNComp.GNTrans.IVTipoTrans <> "Y")) Or _
           (Not mobjGNComp.GNTrans.IVVerificaExist) Or _
           (item.BandServicio = True) Then
            VerificarExistencia = True
            Exit Function
        End If
    End If
    
    With grd
        exist = .ValueMatrix(Row, COL_EXIST)
    
        'En caso de modificación está guardada el codigo de item original en la propiedad .Cell(flexcpData,,)
        If Not IsEmpty(.Cell(flexcpData, Row, COL_CODITEM)) Then
            'Obtiene el código  de item original
            cod_ant = .Cell(flexcpData, Row, COL_CODITEM)
            codb_ant = .Cell(flexcpData, Row, COL_CODBODEGA)
            
            'Si ha cambiado de item ó de bodega, cantidad anterior debe ser 0
            If cod_ant = item.CodInventario And _
                codb_ant = .TextMatrix(Row, COL_CODBODEGA) Then
                If Not IsEmpty(.Cell(flexcpData, Row, COL_CANT)) Then
                    'Obtiene la cantidad original
                    cant_ant = .Cell(flexcpData, Row, COL_CANT)
                End If
            End If
        End If
        
        'Obtiene la suma de cantidad del mismo ítem dentro de la misma transacción
        For i = .FixedRows To .Rows - 1
            If (Not .IsSubtotal(i)) And (i <> Row) Then
                codb = .TextMatrix(i, COL_CODBODEGA)
                cod = .TextMatrix(i, COL_CODITEM)
                If codb = .TextMatrix(Row, COL_CODBODEGA) And _
                    cod = .TextMatrix(Row, COL_CODITEM) Then
                    sumaCant = sumaCant + .ValueMatrix(i, COL_CANT)
                    
                    If Not IsEmpty(.Cell(flexcpData, i, COL_CANT)) Then
                        If .Cell(flexcpData, i, COL_CODBODEGA) = codb And _
                           .Cell(flexcpData, i, COL_CODITEM) = cod Then
                            'Resta la cantidad original, si es que la tiene
                            sumaCant = sumaCant - .Cell(flexcpData, i, COL_CANT)
                        End If
                    End If
                End If
            End If
        Next i
       
        'Si la cantidad está más que la existencia + cantidad original(en caso de modificacion)
        If exist + cant_ant < .ValueMatrix(Row, COL_CANT) + sumaCant Then
            'Si la transacción NO afecta la cantidad, saca mensaje de advertencia
            If Not mobjGNComp.GNTrans.AfectaCantidad Then
                msg = "La cantidad es mayor a la existencia actual." & vbCr & vbCr & _
                      "Confirma que la cantidad está bien?"
                If MsgBox(msg, vbYesNo + vbQuestion + vbDefaultButton2) <> vbYes Then
                    'Corrige la cantidad para que no sea mayor que la existencia
                    .TextMatrix(Row, COL_CANT) = exist + cant_ant - sumaCant
                End If
            'Si afecta la cantaidad no permite ni pregunta
            Else
                msg = "La cantidad no puede ser mayor a la existencia en ésta transacción." & vbCr & vbCr
                msg = msg & "Existencia actual: " & Format$(exist, .ColFormat(COL_EXIST)) & vbCr
                If cant_ant <> 0 Then msg = msg & "Cantidad original: " & Format$(cant_ant, .ColFormat(COL_CANT)) & vbCr
                If sumaCant <> 0 Then msg = msg & "Cant. en otras filas: " & Format$(sumaCant, .ColFormat(COL_CANT)) & vbCr
                msg = msg & "Cantidad máxima: " & Format$(exist + cant_ant - sumaCant, .ColFormat(COL_CANT))
                
                MsgBox msg, vbExclamation
                'Corrige la cantidad para que no sea mayor que la existencia
                .TextMatrix(Row, COL_CANT) = exist + cant_ant - sumaCant
            End If
        End If
    End With
    VerificarExistencia = True
End Function

'*** MAKOTO 14/oct/00 Agregado
'Devuelve saldo de cantidad que se puede utilizar
Private Function VerificarLimiteitem( _
                    ByVal cod As String, _
                    ByVal Row As Long, _
                    ByRef saldo As Currency, tipo As String, Optional CodBodega As String) As Currency
    'Tipo:
    'IVVerificaLimite:  verifica  limite de cantidad de Item  con  transaccion  establecida
    'IVVerificaItemsFuente: verifica limite de  items  con transaccion  fuente
    Const TIPOFUENTE As String = "IVVerificaItemFuente"
    Const TIPONORMAL As String = "IVVerificaLimite"
    Dim cantLimite As Currency, cantGrabada As Currency, msg As String
    Dim fmt As String, i As Long, cant As Currency, cantOtras As Currency
    Dim gnc As GNComprobante
    On Error GoTo errtrap
   'Si no está configurado para verificar, sale no más
    If Not mobjGNComp.GNTrans.IVVerificaLimite And tipo = TIPONORMAL Then
        VerificarLimiteitem = True
        Exit Function
    End If
    If Not mobjGNComp.GNTrans.IVVerificaItemsFuente And tipo = TIPOFUENTE Then
        ' si no esta configurado para que verifique con la transaccion fuente
        VerificarLimiteitem = True
        Exit Function
    ElseIf mobjGNComp.IdTransFuente = 0 And tipo = TIPOFUENTE And _
                                mobjGNComp.GNTrans.IVVerificaItemsFuente Then
        'Sale  si no ha sido importacion
        VerificarLimiteitem = True
        Exit Function
    End If
    'Calcula cantidad utilizada en las filas del mismo comprobante
    With grd
'        cant = .ValueMatrix(Row, COL_CANT)      'La fila actual
        cant = MiCCur(.Cell(flexcpTextDisplay, Row, COL_CANT))      '*** MAKOTO 29/ene/01 Mod.
        For i = .FixedRows To .Rows - 1         'Cantidad de las otras filas del mismo item
            If (Not .IsSubtotal(i)) And (i <> Row) Then
                If .TextMatrix(i, COL_CODITEM) = cod Then
                    cantOtras = cantOtras + .ValueMatrix(i, COL_CANT)
                End If
            End If
        Next i
    End With
    'Verifica el límite
    If tipo = TIPONORMAL Then
        Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(mobjGNComp.IdTransFuente)
        If Not gnc Is Nothing Then
             mobjGNComp.VerificarLimiteItemConTrans gnc.codtrans, gnc.NumTrans, mobjGNComp.IdTransFuente, cod, cantLimite, cantGrabada
            Set gnc = Nothing
        End If
   ElseIf tipo = TIPOFUENTE Then
        mobjGNComp.VerificaItemConFuente cod, cantLimite
    End If
    'Devuelve saldo de cantidad para que pueda corregir en la pantalla
    If cantLimite = 0 Then
        saldo = 0
    Else
        If cantGrabada > 0 Then
            saldo = cantLimite + cantGrabada
        Else
            saldo = cantLimite - cantGrabada
        End If
        If tipo = TIPONORMAL Then
            'solo  transacciones iguales Ej: Egreso / Egreso
            '                                Ingreso/ Ingreso
            If mobjGNComp.GNTrans.IVTipoTrans = "I" Or mobjGNComp.GNTrans.IVTipoTrans = "X" Then
                'If saldo < 0 Then saldo = 0 modificado jeaa 17/01/05
                If saldo < 0 Then saldo = Abs(saldo)
            Else
                If saldo > 0 Then saldo = 0
            End If
        ElseIf tipo = TIPOFUENTE Then
            'solo si las transacciones son diferentes Egreso /Ingreso
            '                                         Ingreso / Egreso
            If mobjGNComp.GNTrans.IVTipoTrans = "I" Or mobjGNComp.GNTrans.IVTipoTrans = "X" Then
                If saldo > 0 Then saldo = 0
            Else
                'If saldo < 0 Then saldo = 0 modificado jeaa 17/01/05
                If saldo < 0 Then saldo = Abs(saldo)
            End If
        End If
        saldo = Abs(saldo) - cantOtras     'Devuelve sin signo
    End If
    'Si está superando el límite, saca mensaje
    If cant > saldo Then
        fmt = mobjGNComp.Empresa.GNOpcion.FormatoCantidad
        If tipo = TIPONORMAL Then
            msg = "Ha intentado registrar la cantidad mayor al límite " & _
                  "establecido en la transacción '" & _
                    mobjGNComp.GNTrans.IVVerificaLimiteCon & "' y '" & _
                    mobjGNComp.GNTrans.IVVerificaLimiteCon & "M'." & vbCr & vbCr & _
                  "    Código de item: " & cod & vbCr & _
                  "    Cantidad límite: " & Format(Abs(cantLimite), fmt) & _
                  "    Cantidad utilizada: " & Format(Abs(cantGrabada) + cantOtras, fmt) & _
                  "    Saldo: " & Format(saldo, fmt)
        Else
                msg = "Ha intentado registrar la cantidad mayor al límite " & _
                      "establecido en la transacción fuente" & _
                      vbCr & vbCr & _
                      "    Código de item: " & cod & vbCr & _
                      "    Cantidad límite: " & Format(Abs(cantLimite), fmt) & _
                      "    Cantidad utilizada: " & Format(cantOtras, fmt) & _
                      "    Saldo: " & Format(saldo, fmt)
        End If
        MsgBox msg, vbInformation
        Exit Function
    End If
    'Si no está superando , devuelve True
    VerificarLimiteitem = True
    Exit Function
errtrap:
    DispErr
    Exit Function
End Function



Private Function VisualizaCodAlt(Row As Long, CodAlt As String) As Boolean
    Dim n As Long, s As String
    On Error GoTo errtrap
    
    If Len(CodAlt) = 0 Then Exit Function

    'Obtiene el numero de items que coincide con el codigo alterno
    n = mobjGNComp.Empresa.BuscaIVCodAlterno(CodAlt, s)
    
    'Si no hay nada, salta al errtrap
    If n = 0 Then Err.Raise 3021
    
    'Si hay más de un registro, saca mensaje
    If n > 1 Then
        MsgBox "Existen " & n & " registro con el mismo código alterno." & vbCr & _
               "Selccione un item de la lista.", vbInformation
    End If
    
    'Visualiza el item (El primero si hay varios)
    VisualizaCodAlt = VisualizaItem(Row, s)
    Exit Function
errtrap:
    'Si no encuentra el codigo
    If Err.Number = 3021 Then
        MsgBox MSG_ERR_NOENCUENTRA & "(" & CodAlt & ")", vbInformation
    Else
        DispErr
    End If
    Exit Function
End Function



Public Property Get GNComprobante() As GNComprobante
    Set GNComprobante = mobjGNComp
End Property

Public Property Set GNComprobante(obj As GNComprobante)
    Set mobjGNComp = obj

    If Not mobjGNComp.EsNuevo Then
        Visualizar
    Else
        ConfigCols
        Limpiar
    End If
    Refresh
End Property

Public Property Get Rows() As Long
    Rows = grd.Rows
End Property

Public Property Get Cols() As Long
    Cols = grd.Cols
End Property

Public Sub Limpiar()
    Dim i As Long
    
    With grd
        For i = .FixedRows To .Rows - 1
            .RowData(i) = 0
        Next i
        .Rows = .FixedRows
    End With
End Sub

Public Sub Visualizar()
    Dim i As Long, neg As Boolean, ivk As IVKardex, COL As Integer, fil As Integer, ListaPrecio As String, item As IVinventario
    
    grd.Redraw = flexRDNone
    ConfigColsVisible          'Para configurar visible o no cada columna --> para ver ColHidden de CodBodega      '*** MAKOTO 16/dic/00
    
    
    'Visualiza los detalles que está en GNComprobante
    '*** MAKOTO 16/dic/00 Modificado para que saque existencia por item cuando está oculta la columna de CodBodega
    Set grd.DataSource = mobjGNComp.ListaIVKardex2(Not grd.ColHidden(COL_CODBODEGA))
    ConfigCols
    
    'Prepara vertor para cargar Codigos de Items y precios solo si el documento ha sido importado
    If mobjGNComp.IdTransFuente <> 0 Then
        ReDim ItemsImportados(1, mobjGNComp.CountIVKardex)
    End If
    
    'Asigna referencia al objeto IVKardex a cada fila de grid
    With grd
        For i = mobjGNComp.CountIVKardex To 1 Step -1
            Set ivk = mobjGNComp.IVKardex(i)
            .RowData(i) = ivk
            If mobjGNComp.GNTrans.IVTipoTrans <> "C" Then    'AUC preguntar que hace
               '*** MAKOTO 09/nov/00 Agregado, Tratamiento especial para transferencia de bodega
                If mTransBodega Then
                    'Si es destino(=ingreso), elimina la fila
                    If ivk.Cantidad > 0 Then
                        .RowData(i) = 0
                        .RemoveItem i
                    'Si es orígen(=egreso), visualiza sin signo
                    Else
                        .TextMatrix(i, COL_CANT) = Abs(ivk.Cantidad)     'Recupera la cantidad SIN signo
                    End If
                End If
            Else   ' auc PARA CAMBIO DE PRESENTACION I.E
                    If mTransBodega Or mItemsSigno = 1 Then        '*** mItemsSigno --> ALEX 21/ene/03 Agregado, Tratamiento especial para módulo producción
                        'Si es destino(=ingreso), elimina la fila                   '*** Producción: ficha de egreso, visualiza solo cant. negativas, pero muestra con signo positivo
                       If ivk.Cantidad < 0 Then
                            .RowData(i) = 0
                            .RemoveItem i
                        'Si es orígen(=egreso), visualiza sin signo
                        End If
                    Else        '** para que funcione mItemsSigno en ctrl prod. mtransBodega siempre = false
                        If mItemsSigno = -1 Then
                            If ivk.Cantidad > 0 Then        'ficha de ingreso en prod., visualiza solo positivos
                                .RowData(i) = 0
                                .RemoveItem i
                            Else
                                .TextMatrix(i, COL_CANT) = ivk.Cantidad
                           End If
                        End If
                    End If
            End If
            If mobjGNComp.IdTransFuente <> 0 Then   '*** Oliver 26/sep/2003 para tener un respaldos de items Importados
                ItemsImportados(0, i - 1) = ivk.CodInventario
                ItemsImportados(1, i - 1) = ivk.Precio
            End If
            '*******************************************************************
            'Recupera el item seleccionado con su configuracion del precio  jeaa 26/01/04
            '*******************************************************************
            Set item = mobjGNComp.Empresa.RecuperaIVInventario(ivk.CodInventario)
            If (Not grd.ColHidden(COL_PU)) And (grd.ColData(COL_PU) >= 0) Then
                ListaPrecio = item.ListaPrecioParaFlex(mobjGNComp)
                    If grd.ColData(COL_PU) > 0 Then ListaPrecio = Mid$(ListaPrecio, 2) 'Quita el |  para que sea solo seleccionable
                    grd.Cell(flexcpData, i, COL_PU) = ListaPrecio
                    If Not grd.ColHidden(COL_PORDCNT) Then
                        If mobjGNComp.GNTrans.IVPrecioPre > 0 Then
                            grd.Cell(flexcpData, i, COL_PORDCNT) = item.Descuento(mobjGNComp.GNTrans.IVPrecioPre) * 100
                        End If
                    End If
            End If

        Next i
    End With
    Set item = Nothing
    '*** MAKOTO 06/feb/01 Agregado
    'Guarda las cantidades originales para la restricción de existencia negativa
    GuardarCantidadOrig
' *********** jeaa 24-12-2003 para cambiar el fondo de las celdas que no son modificables
    If Not mobjGNComp.SoloVer Then
        For fil = 1 To grd.Rows - 1
            For COL = 1 To 28
                    If grd.ColData(COL) = -1 Then
                        grd.Cell(flexcpBackColor, fil, COL, fil, COL) = &H80000018
                    Else
                        grd.Cell(flexcpBackColor, fil, COL, fil, COL) = vbWhite
                    End If
            Next COL
        Next fil
    Else
        For fil = 1 To grd.Rows - 1
            For COL = 1 To 28
                        grd.Cell(flexcpBackColor, fil, COL, fil, COL) = &H80000018
            Next COL
        Next fil
    End If
    PoneNumFila
    VisualizaTotal
    grd.Redraw = True
    grd.Refresh
    Set ivk = Nothing
End Sub

'*** MAKOTO 06/feb/01 Agregado
'En caso de modificación, hay que guardar cantidad original
' y Código de item original
'para que funcione bien el restricción de existencia negativa.
'Esta subrutina guarda la cantidad y código de cada fila en la propiedad 'Cell(flexcpData,,)'
Private Sub GuardarCantidadOrig()
    Dim i As Long
    With grd
        For i = .FixedRows To .Rows - 1
            If Not .IsSubtotal(i) Then
                .Cell(flexcpData, i, COL_CANT) = .ValueMatrix(i, COL_CANT)
                .Cell(flexcpData, i, COL_CODITEM) = .TextMatrix(i, COL_CODITEM)
                .Cell(flexcpData, i, COL_CODBODEGA) = .TextMatrix(i, COL_CODBODEGA)
            End If
        Next i
    End With
End Sub

Public Sub Aceptar()
    Dim i As Long, obj As IVKardex
    On Error GoTo errtrap

    'Pasa los detalles al objeto GNComprobante
    With mobjGNComp
        For i = grd.FixedRows To grd.Rows - 1
            If Not grd.IsSubtotal(i) Then
                Set obj = grd.RowData(i)
                obj.orden = i
            End If
        Next i
        Set obj = Nothing
    End With
    
    Exit Sub
errtrap:
    Select Case Err.Number
    Case Else
        DispErr
    End Select
    Exit Sub
End Sub

Private Sub mnuAgregar_Click()
    AgregaFila
    grd.SetFocus
End Sub


Private Sub mnuEliminar_Click()
    EliminaFila
    grd.SetFocus
End Sub

'*** MAKOTO 30/nov/00 Agregado
Private Sub mnuGrabarPrecio_Click()
    GrabarPrecio
End Sub

'*** MAKOTO 30/nov/00 Agregado
Private Sub GrabarPrecio()
    Dim ivk As IVKardex, i As Long, iv As IVinventario
    Dim s As String, Num As Integer, grabado As Boolean
    On Error GoTo errtrap
    
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No existe ningún item.", vbInformation
        Exit Sub
    End If
    
    'Especifíca número de precio
    Do
        If Len(s) = 0 Then s = "1"      'Valor predeterminado
        s = InputBox("A cuál precio desea grabar? (1-4)", , s)
        If Len(s) = 0 Then Exit Sub     'Si es que cancela, sale
        
        If IsNumeric(s) Then
            Num = Val(s)
            If Num >= 1 And Num <= 4 Then Exit Do
        Else
            MsgBox "Por favor, ingrese un valor numérico.", vbInformation
            s = ""
        End If
    Loop
    
    'Confirmación
    If MsgBox("Está seguro de que desea grabar los precios en Precio" & Num & " de cada item?", _
                vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    
    For i = 1 To mobjGNComp.CountIVKardex
        Set ivk = mobjGNComp.IVKardex(i)
        If Not (ivk Is Nothing) Then
            If Len(ivk.CodInventario) > 0 Then
                MensajeStatus "Está grabando... " & i & ": " & ivk.CodInventario, vbHourglass
                Set iv = mobjGNComp.Empresa.RecuperaIVInventario(ivk.CodInventario)
                iv.Precio(Num) = ivk.Precio
                iv.Grabar
                grabado = True
            End If
        End If
    Next i
    
    Set ivk = Nothing
    Set iv = Nothing
    MensajeStatus
    If grabado Then MsgBox "Los precios han sido grabados.", vbInformation
    Exit Sub
errtrap:
    MensajeStatus
    DispErr
    Exit Sub
End Sub


Private Sub mnuOptimizarCantidad_Click()
    OptimizarCantidad
End Sub

Private Sub OptimizarCantidad()
    Dim iv As IVinventario, ivk As IVKardex
    Dim exist As Currency, i As Long, msg As String
    On Error GoTo errtrap
    
    If mobjGNComp.GNTrans.IVTipoTrans = "I" Or mobjGNComp.GNTrans.IVTipoTrans = "X" Then
        msg = "Automáticamente cambiará la cantidad de cada fila " & _
              "restándo la existencia actual "
        If grd.ColHidden(COL_CODBODEGA) Then
            msg = msg & "por ítem."
        Else
            msg = msg & "por bodega."
        End If
        msg = msg & vbCr & "Esto generalmente sirve por ejemplo para " & _
              "crear pedidos a proveedor(ordenes de compra) " & _
              "con la cantidad óptima, suponiéndo que las cantidades ingresadas " & _
              "representan el monto de demanda."
    Else
        msg = "Comparará la cantidad de cada fila con la existencia actual y " & _
              "la cambiará si es mayor a la existencia."
    End If
    msg = msg & vbCr & vbCr & "Desea continuar?"
    If MsgBox(msg, vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    MensajeStatus MSG_PREPARA, vbHourglass
    
    For i = 1 To mobjGNComp.CountIVKardex
        Set ivk = mobjGNComp.IVKardex(i)
        MensajeStatus "Procesándo #" & i & ": '" & ivk.CodInventario & "'...", vbHourglass
        
        Set iv = mobjGNComp.Empresa.RecuperaIVInventario(ivk.CodInventario)
        If Not (iv Is Nothing) Then
            If grd.ColHidden(COL_CODBODEGA) Then
                exist = iv.Existencia("")
            Else
                exist = iv.Existencia(ivk.CodBodega)
            End If
        End If
        If exist < 0 Then exist = 0         'Si está negativo no hace nada
        
        Set iv = Nothing
        
        'Si la trans es de ingreso suponiendo que en la cantidad está ingresada la cantidad demandada
        If mobjGNComp.GNTrans.IVTipoTrans = "I" Or mobjGNComp.GNTrans.IVTipoTrans = "X" Then
            If ivk.Cantidad > exist Then    ' y hay menos existencia que la demanda (ejm. en Orden de compra)
                ivk.Cantidad = ivk.Cantidad - exist     'Ajusta la cantidad restando lo que hay en stock
            Else
                ivk.Cantidad = 0    'Si hay más existencia que la demanda, no es necesario comprar más por eso pone 0
            End If
            
        'Si la trans es de egreso (En este caso casi no hay mucho sentido de éste menú)
        Else
            If Abs(ivk.Cantidad) > exist Then
                ivk.Cantidad = exist * -1       'Limita a la existencia
            End If
        End If
    Next i
    Set ivk = Nothing
    
    'Actualiza la pantalla
    VisualizaDesdeObjeto
        
    'Genera evento para avisar al form
    RaiseEvent TotalizadoItem       'Aprovechamos usando el mismo evento por que debe hacer lo mismo
    
    MensajeStatus
    Exit Sub
errtrap:
    MensajeStatus
    DispErr
    Exit Sub
End Sub

Private Sub mnuTotalizar_Click()
    TotalizarItem
End Sub

Public Sub TotalizarItem()
    Dim i As Long, cod As String, cant As Currency
    On Error GoTo errtrap
    MensajeStatus MSG_PREPARA, vbHourglass
    
    'Totaliza items repetidos
    
    If mobjGNComp.GNTrans.IVTotItemSignosDif Then
        If mobjGNComp.TotalizaItemRepetidoSignosDiferentes Then
            'Actualiza la pantalla
            VisualizaDesdeObjeto
            With grd
                For i = .Rows - 1 To .FixedRows Step -1
                    If Not .IsSubtotal(i) Then
                        cod = .TextMatrix(i, COL_CODITEM)
                        cant = .ValueMatrix(i, COL_CANT)
                        If (Len(cod) = 0) Or (cant = 0) Then
                            'Remueve de la colección de objeto IVKardex
                            mobjGNComp.RemoveIVKardex 0, grd.RowData(i)
                            'Elimina del grid
                            grd.RemoveItem i
                        End If
                    End If
                Next i
            End With
            VisualizaDesdeObjeto
            'Genera evento para avisar al form
            RaiseEvent TotalizadoItem
        End If
    Else
        If mobjGNComp.TotalizaItemRepetido Then
            'Actualiza la pantalla
            VisualizaDesdeObjeto
            'Genera evento para avisar al form
            RaiseEvent TotalizadoItem
        End If
    End If
    MensajeStatus
    Exit Sub
errtrap:
    MensajeStatus
    DispErr
    Exit Sub
End Sub


Private Sub mobjGNComp_CotizacionCambiado()
    mobjGNComp_MonedaCambiado
End Sub

Private Sub mobjGNComp_Grabado()
    '*** MAKOTO 06/feb/01 Agregado
    GuardarCantidadOrig
End Sub

Private Sub mobjGNComp_MonedaCambiado()
    Dim r As Long
    On Error GoTo errtrap
    
    ConfigColsFormato
    
    'Reasigna todos los valores
    With grd
        For r = .FixedRows To .Rows - 1
            If Not .IsSubtotal(r) Then
                If Not .RowData(r) Is Nothing Then
                    CalculaDetalle r, COL_CANT
                End If
            End If
        Next r
    End With
    
    mobjGNComp.ProrratearIVKardexRecargo
    Exit Sub
errtrap:
    DispErr
    Exit Sub
End Sub

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
    mTransBodega = False
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_Resize()
    'Ajusta el tamaño del grid
    grd.Height = UserControl.ScaleHeight
End Sub

Private Sub UserControl_Terminate()
    Set mobjGNComp = Nothing
    'Cierra los   formularios los  baja de memoria
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
End Sub



Public Sub VisualizaDesdeObjeto()
    Dim ivk As IVKardex, iv As IVinventario, i As Long, s As String, ut As Single, COL As Integer, fil As Integer
    
    With grd
        .Redraw = False
        Limpiar
        'Prepara vertor para cargar Codigos de Items y precios solo si el documento ha sido importado
        If mobjGNComp.IdTransFuente <> 0 Then
            ReDim ItemsImportados(1, mobjGNComp.CountIVKardex)
        End If
        For i = 1 To mobjGNComp.CountIVKardex
            Set ivk = mobjGNComp.IVKardex(i)
            
            
            'Recupera el item
            Set iv = mobjGNComp.Empresa.RecuperaIVInventario(ivk.CodInventario)
            If Not (iv Is Nothing) Then
                'Visualiza si no es transferencia, ó solo egresos en case de transferencia
                If (Not mTransBodega) Or (ivk.Cantidad < 0) Then
                    s = .Rows & vbTab & ivk.CodBodega
                    s = s & vbTab & iv.CodGrupo(1) & vbTab & iv.CodGrupo(2) & vbTab & iv.CodGrupo(3) & vbTab & _
                            iv.CodGrupo(4) & vbTab & iv.CodGrupo(5) & vbTab & _
                            iv.CodInventario & vbTab & iv.CodAlterno1 & vbTab
                    'Item '-' es especial  '+' observacion especial             '*** MAKOTO 16/oct/00
                    If iv.CodInventario = "-" Or iv.CodInventario = "+" Then
                        s = s & ivk.Nota & vbTab
                    Else
                        s = s & iv.Descripcion & vbTab
                    End If
                    
                    If grd.ColHidden(COL_CODBODEGA) Then        '*** MAKOTO 15/dic/00
                        s = s & iv.Existencia("") & vbTab       'Suma de todas las bodegas
                    Else
                        s = s & iv.Existencia(ivk.CodBodega) & vbTab
                    End If
                    
                    '*** MAKOTO 26/ene/01 Mod. para poder visualizar negativos
'                    s = s & Abs(ivk.Cantidad) & vbTab & iv.Unidad & vbTab   '*** MAKOTO 22/jul/00
                    If mobjGNComp.GNTrans.IVTipoTrans = "E" Or mobjGNComp.GNTrans.IVTipoTrans = "Y" Then
                        'Si es egreso multiplica por -1
                        s = s & ivk.Cantidad * -1
                    Else
                        s = s & ivk.Cantidad
                    End If
                    s = s & vbTab & iv.Unidad & vbTab  '*** MAKOTO 22/jul/00
                    s = s & ivk.costo & vbTab & ivk.CostoReal & vbTab & _
                            ivk.CostoTotal & vbTab & ivk.CostoRealTotal & vbTab
                    ut = CalculaUtilidad(ivk)
                    s = s & 0 & vbTab
                    'MODIFICADO OLIVER, NO ESTA VIZUALIZANDO LOS NUEVOS CAMPOS AGREGADOS DE PRECIO+IVA
                    s = s & ivk.Precio & vbTab & ivk.PrecioReal & vbTab & _
                            ivk.PrecioReal * (1 + ivk.IVA) & vbTab & _
                            ivk.PrecioTotal & vbTab & ivk.PrecioRealTotal & vbTab & _
                            ivk.PrecioRealTotal * (1 + ivk.IVA) & vbTab
                    's = s & ivk.Descuento * 100 & vbTab & ivk.IVA * 100 & vbTab & 0 & vbTab & iv.Observacion
                    If mobjGNComp.GNTrans.IVBodegasDestinos Then
                        s = s & ivk.Descuento * 100 & vbTab & ivk.IVA * 100 & vbTab & 0 & vbTab & ivk.Nota
                    Else
                        s = s & ivk.Descuento * 100 & vbTab & ivk.IVA * 100 & vbTab & 0 & vbTab & iv.Observacion
                    End If


                Else
                    s = ""
                End If
            Else
                s = .Rows & vbTab & ivk.CodBodega
                s = s & vbTab & vbTab & vbTab & vbTab & _
                         vbTab & vbTab & _
                        ivk.CodInventario & vbTab & vbTab
                'Item '-' es especial, '+' observacion especial             '*** MAKOTO 16/oct/00
                If ivk.CodInventario = "-" Or ivk.CodInventario = "+" Then s = s & ivk.Nota
                s = s & vbTab & 0 & vbTab
                s = s & Abs(ivk.Cantidad) & vbTab & vbTab    '*** MAKOTO 22/jul/00
                s = s & ivk.costo & vbTab & ivk.CostoReal & vbTab & _
                        ivk.CostoTotal & vbTab & ivk.CostoRealTotal & vbTab
                ut = CalculaUtilidad(ivk) & 0 & vbTab
                'MODIFICADO OLIVER, NO ESTA VIZUALIZANDO LOS NUEVOS CAMPOS AGREGADOS DE PRECIO+IVA
                s = s & ivk.Precio & vbTab & ivk.PrecioReal & vbTab & _
                        ivk.PrecioReal * (1 + ivk.IVA) & vbTab & _
                        ivk.PrecioTotal & vbTab & ivk.PrecioRealTotal & vbTab & _
                        ivk.PrecioRealTotal * (1 + ivk.IVA) & vbTab
                s = s & ivk.Descuento * 100 & vbTab & ivk.IVA * 100 & vbTab & 0 & vbTab & iv.Observacion
            End If
            
            If Len(s) > 0 Then          '*** MAKOTO 09/nov/00 para no agregar items de destino en Trans. Bodegas
                .AddItem s
                .RowData(.Rows - 1) = ivk
            End If
            
            If mobjGNComp.IdTransFuente <> 0 Then   '*** Oliver 26/sep/2003 para tener un respaldos de items Importados
                ItemsImportados(0, i - 1) = ivk.CodInventario
                ItemsImportados(1, i - 1) = ivk.Precio
            End If
        Next i
        
    ' *********** jeaa 24-12-2003 para cambiar el fondo de las celdas que no son modificables
        For fil = 1 To grd.Rows - 1
            For COL = 1 To 26
                    If grd.ColData(COL) = -1 Then
                        grd.Cell(flexcpBackColor, fil, COL, fil, COL) = &H80000018
                    Else
                        grd.Cell(flexcpBackColor, fil, COL, fil, COL) = vbWhite
                    End If
            Next COL
        Next fil
    
        '*** MAKOTO 06/feb/01 Agregado
        'Guarda las cantidades originales para la restricción de existencia negativa
        GuardarCantidadOrig
        
        VisualizaTotal
        .Redraw = True
    End With
End Sub



Private Function MiGetRows(ByVal obj As Object) As Variant
    Dim v As Variant
    
    If Not obj.EOF Then
#If DAOLIB Then
        obj.MoveLast
        obj.MoveFirst
        v = obj.GetRows(obj.RecordCount)
#Else
        v = obj.GetRows
#End If
    End If
    MiGetRows = v
End Function




Private Function BodegaPredeterminada() As String
    Dim sql As String, rs As Recordset
    BodegaPredeterminada = ""
    sql = "Select * from IVbodega WHERE IdBodega = " & mobjGNComp.GNTrans.IdBodegaPre
    Set rs = mobjGNComp.Empresa.OpenRecordset(sql)
    
    If Not rs.EOF Then
        BodegaPredeterminada = rs!CodBodega
    End If
    
    Set rs = Nothing
End Function


Private Function ExistePrecio(Precio As String, ListaPrecios As String, CodItem As String) As Boolean
    Dim P As Variant, i As Integer, Existe As Boolean, rs As Recordset, sql As String
    Existe = False
    
    If Len(ListaPrecios) = 0 Then
        sql = "Select Precio1, precio2, precio3, precio4 from ivinventario WHERE CodInventario = '" & CodItem & "'"
        Set rs = mobjGNComp.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then
            
        Else
        End If
    Else
    P = Split(ListaPrecios, "|")
    For i = 1 To UBound(P)
        If Precio = P(i) Then
            Existe = True
        End If
    Next i
    End If
    ExistePrecio = True
End Function
'*** Angel 21/Mar/03 Agregado
Private Sub VerificarExisMaxMin( _
                    ByVal Row As Long, _
                    ByVal item As IVinventario)
    Dim cant_ant As Currency, msg As String, cod_ant As String, exist As Currency
    Dim sumaCant As Currency, i As Long, codb As String, cod As String
    Dim codb_ant As String, cant_maxmin As Currency, exis_mod As Currency
    ' ó NO está configurado para alertar limites de existencia maxima y mínima
    ' ó el item es de servicio, no hace la verificación
    If (Not mobjGNComp.GNTrans.IVAlertarExisMaxMin) Or _
       (item.BandServicio = True) Then
        Exit Sub
    End If
    Select Case mobjGNComp.GNTrans.IVTipoTrans
    Case "I", "X"
        cant_maxmin = item.ExistenciaMaxima
    Case "E", "Y"
        cant_maxmin = item.ExistenciaMinima
    Case "T"
        Exit Sub 'Por lo pronto no hace nada. Buscar alternativa
    End Select
    With grd
        exist = .ValueMatrix(Row, COL_EXIST)
        'En caso de modificación está guardada el codigo de item original en la propiedad .Cell(flexcpData,,)
        If Not IsEmpty(.Cell(flexcpData, Row, COL_CODITEM)) Then
            'Obtiene el código  de item original
            cod_ant = .Cell(flexcpData, Row, COL_CODITEM)
            codb_ant = .Cell(flexcpData, Row, COL_CODBODEGA)
            'Si ha cambiado de item ó de bodega, cantidad anterior debe ser 0
            If cod_ant = item.CodInventario And _
                codb_ant = .TextMatrix(Row, COL_CODBODEGA) Then
                If Not IsEmpty(.Cell(flexcpData, Row, COL_CANT)) Then
                    'Obtiene la cantidad original
                    cant_ant = .Cell(flexcpData, Row, COL_CANT)
                End If
            End If
        End If
        'Obtiene la suma de cantidad del mismo ítem dentro de la misma transacción
        For i = .FixedRows To .Rows - 1
            If (Not .IsSubtotal(i)) And (i <> Row) Then
                codb = .TextMatrix(i, COL_CODBODEGA)
                cod = .TextMatrix(i, COL_CODITEM)
                If codb = .TextMatrix(Row, COL_CODBODEGA) And _
                    cod = .TextMatrix(Row, COL_CODITEM) Then
                    sumaCant = sumaCant + .ValueMatrix(i, COL_CANT)
                    If Not IsEmpty(.Cell(flexcpData, i, COL_CANT)) Then
                        If .Cell(flexcpData, i, COL_CODBODEGA) = codb And _
                           .Cell(flexcpData, i, COL_CODITEM) = cod Then
                            'Resta la cantidad original, si es que la tiene
                            sumaCant = sumaCant - .Cell(flexcpData, i, COL_CANT)
                        End If
                    End If
                End If
            End If
        Next i
        Select Case mobjGNComp.GNTrans.IVTipoTrans
        Case "I", "X"
            exis_mod = .ValueMatrix(Row, COL_CANT) + sumaCant + (exist - cant_ant)
            If exis_mod >= cant_maxmin Then
                msg = "Alerta. Se ha llegado al stock máximo del item " & vbCr & vbCr
                'Si la cantidad está más que la existencia + cantidad original(en caso de modificacion)
                If cant_ant <> 0 Then
                    msg = msg & "Existencia Actual: " & Format$(exist, .ColFormat(COL_EXIST)) & vbCr
                    msg = msg & "Cantidad original: " & Format$(cant_ant, .ColFormat(COL_CANT)) & vbCr
                    msg = msg & "Cantidad modificada: " & Format$(.ValueMatrix(Row, COL_CANT), .ColFormat(COL_CANT)) & vbCr
                    If sumaCant <> 0 Then msg = msg & "Cant. en otras filas: " & Format$(sumaCant, .ColFormat(COL_CANT)) & vbCr
                    msg = msg & "Existencia Total x Grabar: " & Format$(exis_mod, .ColFormat(COL_CANT)) & vbCr
                Else
                    msg = msg & "Existencia Actual: " & Format$(exist, .ColFormat(COL_EXIST)) & vbCr
                    msg = msg & "Cantidad x Ingresar: " & Format$(.ValueMatrix(Row, COL_CANT), .ColFormat(COL_EXIST)) & vbCr
                    If sumaCant <> 0 Then msg = msg & "Cant. en otras filas: " & Format$(sumaCant, .ColFormat(COL_CANT)) & vbCr
                    msg = msg & "Existencia Total x Grabar: " & Format$(exis_mod, .ColFormat(COL_CANT)) & vbCr
                End If
                'msg = msg & "Diferencia:          " & Format$((exist - cant_ant) + (.ValueMatrix(Row, COL_CANT) + sumaCant), .ColFormat(COL_EXIST)) & vbCr
                msg = msg & "Stock Máximo:   " & Format$(cant_maxmin, .ColFormat(COL_EXIST))
                MsgBox msg, vbExclamation
            End If
        Case "E"
            exis_mod = (exist + cant_ant) - (.ValueMatrix(Row, COL_CANT) + sumaCant)
            If exis_mod <= cant_maxmin Then
                msg = "Alerta. Se ha llegado al stock mínimo del item " & vbCr & vbCr
                'Si la cantidad está más que la existencia + cantidad original(en caso de modificacion)
                If cant_ant <> 0 Then
                    msg = msg & "Existencia Actual: " & Format$(exist, .ColFormat(COL_EXIST)) & vbCr
                    msg = msg & "Cantidad original: " & Format$(cant_ant, .ColFormat(COL_CANT)) & vbCr
                    msg = msg & "Cantidad modificada: " & Format$(.ValueMatrix(Row, COL_CANT), .ColFormat(COL_CANT)) & vbCr
                    If sumaCant <> 0 Then msg = msg & "Cant. en otras filas: " & Format$(sumaCant, .ColFormat(COL_CANT)) & vbCr
                    msg = msg & "Existencia Total x Grabar: " & Format$(exis_mod, .ColFormat(COL_CANT)) & vbCr
                Else
                    msg = msg & "Existencia Actual: " & Format$(exist, .ColFormat(COL_EXIST)) & vbCr
                    msg = msg & "Cantidad x Egresar: " & Format$(.ValueMatrix(Row, COL_CANT), .ColFormat(COL_EXIST)) & vbCr
                    If sumaCant <> 0 Then msg = msg & "Cant. en otras filas: " & Format$(sumaCant, .ColFormat(COL_CANT)) & vbCr
                    msg = msg & "Existencia Total x Grabar: " & Format$(exis_mod, .ColFormat(COL_CANT)) & vbCr
                End If
                'msg = msg & "Diferencia:          " & Format$((exist - cant_ant) + (.ValueMatrix(Row, COL_CANT) + sumaCant), .ColFormat(COL_EXIST)) & vbCr
                msg = msg & "Stock Mínimo:   " & Format$(cant_maxmin, .ColFormat(COL_EXIST))
                MsgBox msg, vbExclamation
            End If
        End Select
0    End With
End Sub

'*** 22 sep 2003 Agregado Oliver
'para recuperar el precio del item del documento fuente

Private Function PrecioIVK_DocFuente(CodItem As String, ByRef Encontro As Boolean) As Currency
    Dim i As Long, PU As Currency
    PU = 0
    Encontro = False
    For i = 0 To UBound(ItemsImportados, 2) - 1
        
        If ItemsImportados(0, i) = CodItem Then
            PU = ItemsImportados(1, i)
            Encontro = True
            Exit For
        End If
    Next i
    PrecioIVK_DocFuente = PU
End Function

Private Sub CalculaPrecio(CodInventario As String, Precio As Variant, CodMoneda As String)
    Dim P As Currency
        If mobjGNComp.GNTrans.IVPrecioPre > 0 Then   ' Agregado Oliver
            P = Precio(1)  '*  para sacar el precio MAS iva
        Else
            P = 0                                       ' en caso de no tener precio predeterminado
        End If
        P = P * mobjGNComp.Cotizacion(CodMoneda) / mobjGNComp.Cotizacion("")  'Convierte en moneda del comprobante
End Sub


Private Function precio_predeterminado(item As Object, EncontroItemEnDocFuente As Boolean) As Currency
    Dim P As Currency
    With item
        If grd.ColData(COL_PU) = -1 And mobjGNComp.IdTransFuente <> 0 Then
            P = PrecioIVK_DocFuente(.CodInventario, EncontroItemEnDocFuente)
        End If
        If Not EncontroItemEnDocFuente Then   'Si no encuentra en item en Doc Fuete pone al precio predeterminado
        
            If mobjGNComp.GNTrans.codPantalla = "IVBQDPR" Then
                If mobjGNComp.GNTrans.IVRetornaNumPrecio Then
                    If PrecioSelec = 0 Then
                        P = .Precio(mobjGNComp.GNTrans.IVPrecioPre) ' para sacar el precio predeterminado
                    Else
                        P = .Precio(PrecioSelec) ' para sacar el precio predeterminado
                    End If
                Else
                    If PrecioSelec = 0 Then
                        P = .Precio(mobjGNComp.GNTrans.IVPrecioPre) ' para sacar el precio predeterminado
                    Else
                        P = PrecioSelec
                    End If
                End If
'                If mobjGNComp.GnTrans.IVPrecioPre > 0 Then   ' Agregado Oliver
'                    P = .Precio(mobjGNComp.GnTrans.IVPrecioPre) ' para sacar el precio predeterminado
'                Else
'                    If PrecioSelec = 0 Then
'                        If mobjGNComp.PrecioRef = 0 Then
'                            P = 0                                       ' en caso de no tener precio predeterminado no saca precio
'                        Else
'                                P = .Precio(mobjGNComp.PrecioRef) ' para sacar el precio predeterminado
'                        End If
'                    Else
'                        P = PrecioSelec
'                    End If
'                End If
            
            
            
            
            Else
                If mobjGNComp.GNTrans.IVPrecioPre > 0 Then   ' Agregado Oliver
                    P = .Precio(mobjGNComp.GNTrans.IVPrecioPre) ' para sacar el precio predeterminado
                Else
                    If PrecioSelec = 0 Then
                        If mobjGNComp.PrecioRef = 0 Then
                            P = 0                                       ' en caso de no tener precio predeterminado no saca precio
                        Else
                                P = .Precio(mobjGNComp.PrecioRef) ' para sacar el precio predeterminado
                        End If
                    Else
                        P = PrecioSelec
                    End If
                End If
            End If
            P = P * mobjGNComp.Cotizacion(.CodMoneda) / mobjGNComp.Cotizacion("")  'Convierte en moneda del comprobante
        End If
    End With
    precio_predeterminado = P
End Function

Private Function CogeSoloBodega(Desc As String) As String
    Dim s As String, i As Long
    i = InStrRev(Desc, "{")
    If i > 0 Then s = Mid$(Desc, i + 1)
    If Len(s) > 0 Then s = Left$(s, Len(s) - 1)
    CogeSoloBodega = s
End Function

Public Sub ActualizaDsctosPCxIV()
    Dim i As Integer, j As Long, k As Long
    Dim item As IVinventario
    For i = 1 To grd.Rows - 2
        'Recupera el item seleccionado
        If Len(grd.TextMatrix(i, COL_CODITEM)) > 0 Then
            Set item = mobjGNComp.Empresa.RecuperaIVInventario(grd.TextMatrix(i, COL_CODITEM))
            With item
                If mobjGNComp.Empresa.GNOpcion.ObtenerValor("TipoDescuentoItems") = 1 Then
                    Dim AuxDesct As Double, ivgrupo As Integer
                    ivgrupo = CInt(mobjGNComp.Empresa.GNOpcion.ObtenerValor("DescuentoNumIVGrupo")) + 1
                    AuxDesct = mobjGNComp.Empresa.DesctoPCxIV(.CodGrupo(ivgrupo), mobjGNComp.CodClienteRef)
                Else
                    AuxDesct = .Descuento(mobjGNComp.GNTrans.IVPrecioPre) * 100
                End If
                If Not grd.ColHidden(COL_PORDCNT) Then 'Corregido 25/03/2004 jeaa
                    If mobjGNComp.GNTrans.IVPrecioPre > 0 Then
                        grd.TextMatrix(i, COL_PORDCNT) = AuxDesct
                    Else
                        grd.TextMatrix(i, COL_PORDCNT) = "0"
                    End If
                End If
            End With
        End If
        k = i
        For j = 0 To 26
            CalculaDetalle k, j
        Next j
    Next i
End Sub

'jeaa 05/08/2005
Public Sub HabilitarCtrlsGrupoCaja(ByVal bandHabilita As Boolean)
Dim COL As Integer, fil As Integer
    For COL = 1 To 27
        If Not bandHabilita Then
            grd.ColData(COL) = -1
            For fil = 1 To grd.Rows - 2
                            grd.Cell(flexcpBackColor, fil, COL, fil, COL) = &H80000018
            Next fil
        End If
    Next COL
End Sub

Private Sub grd_AfterUserResize(ByVal Row As Long, ByVal COL As Long)
    With grd
        SaveSetting APPNAME, "SiiConfigCols", "config_col_item_" & mobjGNComp.GNTrans.codtrans & "_" & COL, .ColWidth(COL)
        ANCHO_COLS(COL) = .ColWidth(COL)
    End With
End Sub
Private Sub GetColsWidth()
    Dim i As Integer
    With grd
            For i = 0 To 28
                ANCHO_COLS(i) = GetSetting(APPNAME, "SiiConfigCols", "config_col_item_" & mobjGNComp.GNTrans.codtrans & "_" & i, 1200)
            Next i
    End With
End Sub

'*** AUC 13/12/05
Public Property Get ItemsSigno() As Integer
    ItemsSigno = mItemsSigno
End Property
'*** AUC 13/12/05
Public Property Let ItemsSigno(ByVal value As Integer)
    mItemsSigno = value
End Property

Public Sub VisualizaDesdeObjetoParaCambioPre() 'AUC
    Dim ivk As IVKardex, iv As IVinventario, i As Long, s As String, ut As Single
    Dim bandSalta As Boolean, COL As Integer, fil As Integer, CodItemPadre As String
    With grd
        bandSalta = False
        .Redraw = False
        Limpiar
        'Prepara vertor para cargar Codigos de Items y precios solo si el documento ha sido importado
        If mobjGNComp.IdTransFuente <> 0 Then
            ReDim ItemsImportados(1, mobjGNComp.CountIVKardex)
        End If
            For i = 1 To mobjGNComp.CountIVKardex
            If mobjGNComp.GNTrans.IVTipoTrans <> "C" Then
                If mItemsSigno = -1 Then        'Solo carga items negativos
                   bandSalta = mobjGNComp.IVKardex(i).Cantidad > 0
                ElseIf mItemsSigno = 1 Then  'Solo carga items positivos
                       bandSalta = mobjGNComp.IVKardex(i).Cantidad < 0
                End If
            Else    'para Transformacion jeaa 25/08/2005
                If mItemsSigno = -1 Then        'Solo carga items negativos
                   bandSalta = mobjGNComp.IVKardex(i).Cantidad > 0
                ElseIf mItemsSigno = 1 Then 'Solo carga items positivos
                       bandSalta = mobjGNComp.IVKardex(i).Cantidad < 0
                End If
            End If
            If bandSalta = False Then  ' SI CANTIDAD ES NEGATIVA
                Set ivk = mobjGNComp.IVKardex(i)
                'Recupera el item
                Set iv = mobjGNComp.Empresa.RecuperaIVInventario(ivk.CodInventario)
                If Not (iv Is Nothing) Then
                    'Visualiza si no es transferencia, ó solo egresos en case de transferencia
                    If (Not mTransBodega) Or (ivk.Cantidad < 0) Then
                        s = .Rows & vbTab & ivk.CodBodega
                        s = s & vbTab & iv.CodGrupo(1) & vbTab & iv.CodGrupo(2) & vbTab & iv.CodGrupo(3) & vbTab & _
                                iv.CodGrupo(4) & vbTab & iv.CodGrupo(5) & vbTab & _
                                iv.CodInventario & vbTab & iv.CodAlterno1 & vbTab
                        'Item '-' es especial               '*** MAKOTO 16/oct/00
                        If iv.CodInventario = "-" Then
                            s = s & ivk.Nota & vbTab
                        Else
                            s = s & iv.Descripcion & vbTab
                        End If
                        If grd.ColHidden(COL_CODBODEGA) Then        '*** MAKOTO 15/dic/00
                            s = s & iv.Existencia("") & vbTab       'Suma de todas las bodegas
                        Else
                            s = s & iv.Existencia(ivk.CodBodega) & vbTab
                        End If
                        '*** MAKOTO 26/ene/01 Mod. para poder visualizar negativos
    '                    s = s & Abs(ivk.Cantidad) & vbTab & iv.Unidad & vbTab   '*** MAKOTO 22/jul/00
                        'oooooooooooooooooooooooooooooojo cambiar Tipo Trans no olvidar
                        If mobjGNComp.GNTrans.IVTipoTrans = "E" Or mobjGNComp.GNTrans.IVTipoTrans = "Y" Then
                            'Si es egreso multiplica por -1
                            s = s & ivk.Cantidad * -1
'                        ElseIf mobjGNComp.GNTrans.IVTipoTrans = "C" Then
'                        ivk.Cantidad = Abs(ivk.Cantidad)
'                            s = s & Abs(ivk.Cantidad)
                        Else
                            s = s & ivk.Cantidad
                        End If
                        'AUC para cambio de presentacion IE
'                        If mobjGNComp.GNTrans.IVTipoTrans = "C" Then
'                            s = s & vbTab & iv.Unidad & vbTab  '*** MAKOTO 22/jul/00
'                            s = s & iv.CostoUltimoIngreso & vbTab & ivk.CostoReal & vbTab & _
'                                ivk.CostoTotal & vbTab & ivk.CostoRealTotal & vbTab
'                               ' mbooItems = True
'                        Else
                            s = s & vbTab & iv.Unidad & vbTab  '*** MAKOTO 22/jul/00
                            s = s & ivk.costo & vbTab & ivk.CostoReal & vbTab & _
                                ivk.CostoTotal & vbTab & ivk.CostoRealTotal & vbTab
                        'End If
                        ut = CalculaUtilidad(ivk)
                        s = s & 0 & vbTab
                        'MODIFICADO oLIVER, NO ESTA VIZUALIZANDO LOS NUEVOS CAMPOS AGREGADOS DE PRECIO+IVA
                        s = s & ivk.Precio & vbTab & ivk.PrecioReal & vbTab & _
                                ivk.PrecioReal * (1 + ivk.IVA) & vbTab & _
                                ivk.PrecioTotal & vbTab & ivk.PrecioRealTotal & vbTab & _
                                ivk.PrecioRealTotal * (1 + ivk.IVA) & vbTab
                        s = s & ivk.Descuento * 100 & vbTab & ivk.IVA * 100 & vbTab & 0 & vbTab & iv.Observacion
'''                        If Len(CodItemPadre) > 0 Then
'''                            If mobjGNComp.Empresa.RevisaPrinicialPreparacion(ivk.CodInventario, CodItemPadre) Then
'''                                'CambiaColorPrincipal
'''                                grd.Cell(flexcpBackColor, i, 1, i, 28) = vbCyan '&HC0FFFF
'''                            Else
'''                                'CambiaColorPrincipal
'''                                grd.Cell(flexcpBackColor, .Rows - 1, .Cols - 7, .Rows - 1, .Cols - 6) = vbWhite
'''                            End If
'''                        End If
                    Else
                        s = ""
                    End If
                Else
                    s = .Rows & vbTab & ivk.CodBodega
                    s = s & vbTab & vbTab & vbTab & vbTab & _
                             vbTab & vbTab & _
                            ivk.CodInventario & vbTab & vbTab
                    'Item '-' es especial               '*** MAKOTO 16/oct/00
                    If ivk.CodInventario = "-" Then s = s & ivk.Nota
                    s = s & vbTab & 0 & vbTab
                    s = s & Abs(ivk.Cantidad) & vbTab & vbTab    '*** MAKOTO 22/jul/00
                    s = s & ivk.costo & vbTab & ivk.CostoReal & vbTab & _
                            ivk.CostoTotal & vbTab & ivk.CostoRealTotal & vbTab
                    ut = CalculaUtilidad(ivk) & 0 & vbTab
                    'MODIFICADO OLIVER, NO ESTA VIZUALIZANDO LOS NUEVOS CAMPOS AGREGADOS DE PRECIO + IVA
                    s = s & ivk.Precio & vbTab & ivk.PrecioReal & vbTab & _
                            ivk.PrecioReal * (1 + ivk.IVA) & vbTab & _
                            ivk.PrecioTotal & vbTab & ivk.PrecioRealTotal & vbTab & _
                            ivk.PrecioRealTotal * (1 + ivk.IVA) & vbTab
                    ' AUC ver despues s = s & ivk.Descuento * 100 & vbTab & ivk.IVA * 100 & vbTab & 0 & vbTab & iv.Observacion
                End If
                If Len(s) > 0 Then          '*** MAKOTO 09/nov/00 para no agregar items de destino en Trans. Bodegas
                    .AddItem s
                        If Len(CodItemPadre) > 0 Then
                            If mobjGNComp.Empresa.RevisaPrinicialPreparacion(ivk.CodInventario, CodItemPadre) Then
                                'es principal
                                grd.Cell(flexcpBackColor, i - 1, 1, i - 1, 28) = &H80000018
                            Else
                                'no es principal
                                grd.Cell(flexcpBackColor, i - 1, 1, i - 1, 28) = vbWhite
                            End If
                        End If
                    grd.Refresh
                    
                '    MsgBox s
                    .RowData(.Rows - 1) = ivk
                End If
                If mobjGNComp.IdTransFuente <> 0 Then   '*** Oliver 26/sep/2003 para tener un respaldos de items Importados
                    ItemsImportados(0, i - 1) = ivk.CodInventario
                    ItemsImportados(1, i - 1) = ivk.Precio
                End If
            Else
                Set ivk = mobjGNComp.IVKardex(i)
                CodItemPadre = ivk.CodInventario
            End If
        Next i
    ' *********** jeaa 24-12-2003 para cambiar el fondo de las celdas que no son modificables
    For fil = 1 To grd.Rows - 1
        For COL = 1 To 26
                If grd.Cell(flexcpBackColor, fil, COL, fil, COL) <> &H80000018 Then
                    If grd.ColData(COL) = -1 Then
                        grd.Cell(flexcpBackColor, fil, COL, fil, COL) = &H80000018
                    Else
                        grd.Cell(flexcpBackColor, fil, COL, fil, COL) = vbWhite
                    End If
                End If
        Next COL
    Next fil
        '*** MAKOTO 06/feb/01 Agregado
        'Guarda las cantidades originales para la restricción de existencia negativa
        GuardarCantidadOrig
        VisualizaTotal
        .Redraw = True
    End With
End Sub

Private Sub GrabarUltimoCostoIngreso()
   Dim itemCom As IVinventario
   Call mobjGNComp.Empresa.GrabarCostoUltimoIngreso(grd.TextMatrix(grd.Row, COL_CODITEM), grd.TextMatrix(grd.Row, COL_CU))
End Sub

Public Sub OcultarColumnas()
    Dim i As Integer, j As Long, k As Long
    Dim item As IVinventario
    For i = COL_UNIDAD To COL_TENTREGA
        grd.ColHidden(i) = True
    Next i
    grd.ColHidden(COL_EXIST) = True
End Sub


'AUC 29/08/07 para guias de transporte
Public Sub CargaItemsTemporal()
Dim i As Long
Dim ix As Long
Dim obj As IVKardex
Dim item As IVinventario
'primero elimina el ivkardex de los devueltos
For i = 1 To grd.Rows - 1
    If Not grd.IsSubtotal(i) Then
            ReDim Preserve ItemsTmp(11, i)
            ItemsTmp(1, i) = grd.TextMatrix(i, COL_CODBODEGA)
            ItemsTmp(2, i) = grd.TextMatrix(i, COL_CODITEM)
            ItemsTmp(3, i) = (grd.TextMatrix(i, COL_CANT)) * -1
            ItemsTmp(4, i) = grd.ValueMatrix(i, COL_PTR)
            ItemsTmp(5, i) = grd.ValueMatrix(i, COL_PT)
            ItemsTmp(6, i) = grd.ValueMatrix(i, COL_CTR)
            ItemsTmp(7, i) = grd.ValueMatrix(i, COL_CT)
            Set item = mobjGNComp.Empresa.RecuperaIVInventario(grd.TextMatrix(i, COL_CODITEM))
            ItemsTmp(8, i) = item.PorcentajeIVA
            ItemsTmp(9, i) = i 'Guarda el orden de los items
            ItemsTmp(10, i) = True 'Si es subitem para que no imprima
            ItemsTmp(11, i) = True 'Si es subitem para que no imprima
    End If
Next
End Sub

'AUC 29/08/07 para guias de transporte
Public Sub cargarValores()
Dim i As Long
Dim ix As Long
Dim obj As IVKardex
Dim item As IVinventario
'primero elimina el ivkardex de los devueltos
For i = mobjGNComp.CountIVKardex To 1 Step -1
    mobjGNComp.RemoveIVKardex (i)
Next
 For i = 1 To UBound(ItemsTmp, 2)
           ix = mobjGNComp.AddIVKardex
            mobjGNComp.IVKardex(ix).CodBodega = ItemsTmp(1, i)
            mobjGNComp.IVKardex(ix).CodInventario = ItemsTmp(2, i)
            mobjGNComp.IVKardex(ix).Cantidad = ItemsTmp(3, i)
            mobjGNComp.IVKardex(ix).PrecioRealTotal = ItemsTmp(4, i)
            mobjGNComp.IVKardex(ix).PrecioTotal = ItemsTmp(5, i)
            mobjGNComp.IVKardex(ix).CostoRealTotal = ItemsTmp(6, i)
            mobjGNComp.IVKardex(ix).CostoTotal = ItemsTmp(7, i)
            Set item = mobjGNComp.Empresa.RecuperaIVInventario(ItemsTmp(2, i))
            mobjGNComp.IVKardex(ix).IVA = ItemsTmp(8, i)
            mobjGNComp.IVKardex(ix).orden = ItemsTmp(9, i)
           ' mobjGNComp.IVKardex(ix).bandImprimir = ItemsTmp(9, i) 'True 'Si es subitem para que no imprima
           ' mobjGNComp.IVKardex(ix).bandVer = ItemsTmp(10, i) 'True  'Si es subitem para que no imprima
  '  End If
Next
End Sub

Private Function VisualizaItemPromocion(Row As Long, cod As String, Optional CodBodega As String, Optional Precio As Currency, Optional codp As String) As Boolean
    Dim item As IVinventario, c As Currency, P As Currency, msg As String, ListaPrecio As String
    Dim saldo As Currency, EncontroItemEnDocFuente As Boolean
    On Error GoTo errtrap
    If Len(cod) = 0 Then Exit Function
    MensajeStatus MSG_PREPARA, vbHourglass
    '********************************** VERIFICACION DE LIMITE DE CANTIDAD CON PRESPUESTO
    'Item con código '-' es especial
    If cod <> "-" And cod <> "+" Then
        'Verifica con el límite establecido         '*** MAKOTO 15/oct/00 Agregado
        If Not VerificarLimiteitem(cod, Row, saldo, "IVVerificaLimite", CodBodega) Then
            'Si está configurado para que no permita grabar superando el límite
            If mobjGNComp.GNTrans.IVVerificaLimiteNoGrabar Then
                If saldo >= 0 Then
                    'Si hay saldo, modifica la cantidad
                    grd.TextMatrix(Row, COL_CANT) = saldo
                Else
                    'Si no hay saldo, no permite seleccionar ése item
                    VisualizaItemPromocion = False
                    MensajeStatus
                    Exit Function
                End If
            End If
        End If
    End If
    '**********************************
    '********************VERIFICACION DE LIMITE DE CANTIDAD CON TRANSFUENTE
    If cod <> "-" And cod <> "+" Then
        If Not VerificarLimiteitem(cod, Row, saldo, "IVVerificaItemFuente", CodBodega) Then
            If saldo > 0 Then
                'Si hay saldo, modifica la cantidad
                grd.TextMatrix(Row, COL_CANT) = saldo
            Else
                'Si no hay saldo, no permite seleccionar ése item
                VisualizaItemPromocion = False
                MensajeStatus
                Exit Function
            End If
        End If
    End If
    'Recupera el item seleccionado
        Set item = mobjGNComp.Empresa.RecuperaIVInventario(cod)
        With item
        If Not grd.ColHidden(COL_CODBODEGA) Then grd.TextMatrix(Row, COL_CODBODEGA) = CodBodega
        If Not grd.ColHidden(COL_G1) Then grd.TextMatrix(Row, COL_G1) = .CodGrupo(1)
        If Not grd.ColHidden(COL_G2) Then grd.TextMatrix(Row, COL_G2) = .CodGrupo(2)
        If Not grd.ColHidden(COL_G3) Then grd.TextMatrix(Row, COL_G3) = .CodGrupo(3)
        If Not grd.ColHidden(COL_G4) Then grd.TextMatrix(Row, COL_G4) = .CodGrupo(4)
        If Not grd.ColHidden(COL_G5) Then grd.TextMatrix(Row, COL_G5) = .CodGrupo(5)
        grd.TextMatrix(Row, COL_CODITEM) = .CodInventario
        grd.TextMatrix(Row, COL_CODALT) = .CodAlterno1
        If cod = "-" Or cod = "+" Then                       '*** MAKOTO 16/oct/00 Item '-' es especial
            If Len(grd.TextMatrix(Row, COL_DESC)) = 0 Then      'Sólo cuando no está ingresado nada
                grd.TextMatrix(Row, COL_DESC) = .Descripcion    '   visualizamos la descripcion
            End If
        Else
            grd.TextMatrix(Row, COL_DESC) = .Descripcion
        End If
        If Len(.CODUnidad) <> 0 Then
            grd.TextMatrix(Row, COL_UNIDAD) = .CODUnidad            '*** MAKOTO 22/jul/00
        Else
            grd.TextMatrix(Row, COL_UNIDAD) = .Unidad            '*** MAKOTO 22/jul/00
        End If
        grd.TextMatrix(Row, COL_PORIVA) = .PorcentajeIVA * 100
        'Si el C.U. No es modificable o está en 0, visualiza el costo calculado
        If (grd.ColHidden(COL_CU) = True) _
            Or (grd.ColData(COL_CU) < 0) _
            Or (grd.ValueMatrix(Row, COL_CU) = 0) Then
'            c = .costo(mobjGNComp.FechaTrans, grd.ValueMatrix(Row, COL_CANT), mobjGNComp.TransID)
            c = .CostoDouble2(mobjGNComp.FechaTrans, _
                                grd.ValueMatrix(Row, COL_CANT), _
                                mobjGNComp.TransID, _
                                mobjGNComp.HoraTrans)  '*** MAKOTO 08/dic/00 Agregado Hora
            'Si el costo calculado está en otra moneda, convierte en moneda de trans.
            If mobjGNComp.CodMoneda <> .CodMoneda Then
                c = c * mobjGNComp.Cotizacion(.CodMoneda) / mobjGNComp.Cotizacion("")
            End If
            grd.TextMatrix(Row, COL_CU) = c
        End If
        'Si P.U. no está oculto , visualiza el Precio1
        If (Not grd.ColHidden(COL_PU)) Then
'        'Si P.U. no está oculto Y está en 0, visualiza el Precio1
'        If (Not grd.ColHidden(COL_PU)) And (grd.ValueMatrix(Row, COL_PU) = 0) Then
            '***19/09/2003  oliver
            'Agregado condicion para el caso en el que es documento importado y el precio no es modificable
            ' debe respetar el precio del documento fuente
            If mobjGNComp.GNTrans.codPantalla <> "IVDVBQD" Then
'                If mobjGNComp.PrecioRef = 0 Then
                 If mobjGNComp.GNTrans.IVVisItemPromocion Then
                        grd.TextMatrix(Row, COL_PU) = Precio 'AUC para la promocion con precio 0
                 Else
                        grd.TextMatrix(Row, COL_PU) = precio_predeterminado(item, EncontroItemEnDocFuente)
                 End If
            End If
        End If
        'Si P.U. no está oculto Y es modificable/seleccionable, guarda la lista de precios
        If (Not grd.ColHidden(COL_PU)) And (grd.ColData(COL_PU) >= 0) Then
            ListaPrecio = .ListaPrecioParaFlex(mobjGNComp)
            If grd.ColData(COL_PU) > 0 Then ListaPrecio = Mid$(ListaPrecio, 2) 'Quita el |  para que sea solo seleccionable
            grd.Cell(flexcpData, Row, COL_PU) = ListaPrecio
        End If
'--------------------------- Si P.UIVA. no está oculto , visualiza el Precio1
        If (Not grd.ColHidden(COL_PUIVA)) Then
'''''''        'Si P.U. no está oculto Y está en 0, visualiza el Precio1
''''''        '************************
''''''        'coloca el precio predetermnado en .precio
''''''        '************************
        If (grd.ColHidden(COL_PU)) Then
                  grd.TextMatrix(Row, COL_PU) = precio_predeterminado(item, EncontroItemEnDocFuente)
        End If
'''''''******************
            '***19/09/2003  oliver
            'Agregado condicion para el caso en el que es documento importado y el precio no es modificable
            ' debe respetar el precio del documento fuente
            If grd.ColData(COL_PUIVA) = -1 And mobjGNComp.IdTransFuente <> 0 Then
                P = PrecioIVK_DocFuente(.CodInventario, EncontroItemEnDocFuente)
            End If
            If Not EncontroItemEnDocFuente Then   'Si no encuentra en item en Doc Fuete pone al precio predeterminado
                If mobjGNComp.GNTrans.IVPrecioPre > 0 Then   ' Agregado Oliver
                    P = .Precio(mobjGNComp.GNTrans.IVPrecioPre) + (.Precio(mobjGNComp.GNTrans.IVPrecioPre) * .PorcentajeIVA) '*  para sacar el precio MAS iva
                Else
                    P = 0                                       ' en caso de no tener precio predeterminado
                                                                ' no saca precio
                End If
            End If
            P = P * mobjGNComp.Cotizacion(.CodMoneda) / mobjGNComp.Cotizacion("")  'Convierte en moneda del comprobante
            grd.TextMatrix(Row, COL_PUIVA) = P
        End If
        '****************************** VISUALIZACION DE EXISTENCIA
        'Visualiza la existencia en la bodega seleccionada
        If mTransBodega And Len(mCodBodegaOrigen) > 0 Then                   '*** MAKOTO 14/nov/00
            'En caso de pantalla de transferencia, coge la existencia de bodega de orígen
            grd.TextMatrix(Row, COL_EXIST) = .Existencia(mCodBodegaOrigen)
        'Si columna de bodega está visible                  '*** MAKOTO 15/dic/00
        ElseIf Len(grd.TextMatrix(Row, COL_CODBODEGA)) > 0 Then
        'jeaa 28/04/04 modificado para que solo tome existencia de la bodega predeterminada
        'Not grd.ColHidden(COL_CODBODEGA) Then
            'Visualiza existencia de la bodega seleccionada
            grd.TextMatrix(Row, COL_EXIST) = .Existencia(grd.TextMatrix(Row, COL_CODBODEGA))
        'Si columna de bodega está oculta
        Else
            'Visualiza la suma de todas las bodegas         '*** MAKOTO 15/dic/00
            grd.TextMatrix(Row, COL_EXIST) = .Existencia("")
        End If
        '****************************** VISUALIZACION DE DESCUENTO X ITEM  ******** jeaa 22-Dic-03-01-12-03
        'jeaa 03/01/05 descuento por Cliente/Grupo Item 0 es por item 1 por clieente grupo
        If Val(mobjGNComp.Empresa.GNOpcion.ObtenerValor("TipoDescuentoItems")) = 1 Then
            Dim AuxDesct As Double, ivgrupo As Integer
            ivgrupo = CInt(mobjGNComp.Empresa.GNOpcion.ObtenerValor("DescuentoNumIVGrupo")) + 1
            AuxDesct = mobjGNComp.Empresa.DesctoPCxIV(.CodGrupo(ivgrupo), mobjGNComp.CodClienteRef)
        Else
            'jeaa 10/02/05
            If Not grd.ColHidden(COL_PORDCNT) Then
                If mobjGNComp.GNTrans.IVPrecioPre > 0 Then
                    AuxDesct = .Descuento(mobjGNComp.GNTrans.IVPrecioPre) * 100
                Else
                    AuxDesct = 0
                End If
            End If
        End If
        If Not grd.ColHidden(COL_PORDCNT) Then 'Corregido 25/03/2004 jeaa
            'jeaa 03/10/2005
            If Not mobjGNComp.GNTrans.IVOmitirPorcentajeDesct Then
                If mobjGNComp.GNTrans.IVPrecioPre > 0 Then
                    grd.TextMatrix(Row, COL_PORDCNT) = AuxDesct
                Else
                    grd.TextMatrix(Row, COL_PORDCNT) = "0"
                End If
            Else
                grd.TextMatrix(Row, COL_PORDCNT) = "0"  'jeaa 03/10/2005
            End If
        End If
        '****************************** VERIFICACION DE EXISTENCIA NEGATIVA
        '*** MAKOTO 06/feb/01 Mod.
        VerificarExistencia Row, item
        '*** ANGEL 21/mar/03 Agregado
        VerificarExisMaxMin Row, item
        '****************************** INGRESO DE NOTA LIBRE
        '*** MAKOTO 16/oct/00
        'Cuando selecciona item '-', ingresa la nota libre
        If (.CodInventario = "-") And (grd.COL < COL_CANT) Then
            msg = grd.TextMatrix(Row, COL_DESC)
            Do
                msg = InputBox("Ingrese una nota", , msg)
                If Len(msg) > MAXLEN_NOTA Then
                    MsgBox "La longitud máxima de la nota es de " & _
                            MAXLEN_NOTA & " caracteres.", vbInformation
                    msg = Left$(msg, MAXLEN_NOTA)        'Automáticamente corta hasta MAXLEN_NOTA létras
                Else
                    Exit Do
                End If
            Loop
            grd.TextMatrix(Row, COL_DESC) = msg
        End If
        '****************************** INGRESO DE OBSERVACION LIBRE
        '*** JEAA 22/04/2005
        'Cuando selecciona item '+', ingresa la observacion libre
        If (.CodInventario = "+") And (grd.COL < COL_CANT) Then
            msg = grd.TextMatrix(Row, COL_DESC)
            Do
                msg = InputBox("Ingrese una observación", , msg)
                If Len(msg) > MAXLEN_NOTA Then
                    MsgBox "La longitud máxima de la nota es de " & _
                            MAXLEN_NOTA & " caracteres.", vbInformation
                    msg = Left$(msg, MAXLEN_NOTA)        'Automáticamente corta hasta MAXLEN_NOTA létras
                Else
                    Exit Do
                End If
            Loop
            grd.TextMatrix(Row, COL_DESC) = msg
        End If
        '******************************
        'jeaa 23-03-04 para guardad el descuento por item
        If Not grd.ColHidden(COL_PORDCNT) Then
            grd.Cell(flexcpData, Row, COL_PORDCNT) = grd.TextMatrix(Row, COL_PORDCNT)
        End If
        '*************************************
'         devolucion automatica
        If mobjGNComp.GNTrans.codPantalla = "IVDVBQD" Then
            If grd.TextMatrix(Row, COL_CANT) > 0 Then
                grd.TextMatrix(Row, COL_CANT) = grd.TextMatrix(Row, COL_CANT)
            End If
        End If
        '***Agregado. 05/ago/04. Angel. Para refrescar el recargo x item
        Dim obj As IVKardex
        Set obj = grd.RowData(Row)
        obj.ValorRecargoItem = .ValorRecargo
        obj.IdPadre = item.RecuperaID(codp)
        
        'Calcula detalles (PT,CT etc...)
        
        ' y graba la nota libre en la propiedad Nota        '*** MAKOTO 16/oct/00
        '***************************************************
        '***Agregado. 22/04/05. Esteban, campo observacion
        If Not grd.ColHidden(COL_OBS) Then
            grd.TextMatrix(Row, COL_OBS) = .Observacion
        End If
        '***************************************************
        CalculaDetalle Row, COL_CANT
        VisualizaItemPromocion = True
        'jeaa 06/04/2005
        'para control de venta en fracciones
        If (Int(grd.TextMatrix(Row, COL_CANT)) - grd.TextMatrix(Row, COL_CANT)) <> 0 Then
             If .bandFraccion = False Then
                MsgBox "Este Item solo solo puede facturarse en unidades "
                If Int(grd.TextMatrix(Row, COL_CANT)) = 0 Then
                    grd.TextMatrix(Row, COL_CANT) = 1
                Else
                    grd.TextMatrix(Row, COL_CANT) = Int(grd.TextMatrix(Row, COL_CANT))
                End If
            End If
        End If
    End With
    Set item = Nothing
    MensajeStatus
    Exit Function
errtrap:
    MensajeStatus
    'Si no encuentra el codigo de item
    If Err.Number = 3021 Or Err.Number = 91 Then
        MsgBox MSG_ERR_NOENCUENTRA & "(" & cod & ")", vbInformation
    Else
        DispErr
    End If
    Exit Function
End Function

Private Sub PonerColorFilas() 'Auc 22/03/06
    Dim i As Long, j As Long
    With grd
        If .Rows <= .FixedRows Then Exit Sub
        'Si la fila no es editable cambia de color
        For i = 1 To grd.Rows - 1
            If Not .IsSubtotal(i) Then
                .Cell(flexcpBackColor, i, COL_CODBODEGA, i, COL_TENTREGA) = &HFFFFC0         'vbYellow
            End If
        Next
    End With
End Sub


Private Sub AgregaFilaPromocion()
    Dim r As Long, r2 As Long, ix As Long, COL As Integer
    Dim itemP As IVinventario
    On Error GoTo errtrap

    'Verifica si ya está número maximo de filas
    If (mobjGNComp.GNTrans.IVNumFilaMax > 0) And _
        (mobjGNComp.CountIVKardex >= mobjGNComp.GNTrans.IVNumFilaMax) Then
        MsgBox "No se puede agregar más filas porque está limitado hasta " & _
         mobjGNComp.GNTrans.IVNumFilaMax & " filas." & vbCr & vbCr & _
        "Si hay más detalle de items, regístrelos en otro comprobante." & vbCr & _
        "Si quiere cambiar el límite, váyase a la configuración de la transacción, por favor.", vbInformation
        Exit Sub
    End If

    'Agrega nuevo objeto IVKardex al comprobante        '*** MAKOTO 14/oct/00 Modificado
    ix = mobjGNComp.AddIVKardex
    
    With grd
        r2 = .Rows - 1
        If .IsSubtotal(.Rows - 1) Then r2 = r2 - 1
        'Si no es la primera fila
        If r2 > 0 Then
            'Si no está en la fila de total
            If Not .IsSubtotal(.Row) Then
                .AddItem "", .Row + 1
                r = .Row + 1
            'Si está en la fila de total
            Else
                .AddItem "", .Row
                r = .Row
            End If
        'Si es la primera fila
        Else
            'Si no está en la fila de total
            If (.Row < .Rows - 1) Or (.Row = 0) Then
'            If Not .IsSubtotal(.Row) Then
                .AddItem ""
                r = .Rows - 1
            'Si está en la fila de total
            Else
                .AddItem "", .Row
                r = .Row
            End If
        End If
        numfila = r
        'Asigna el indice de nuevo objeto a la fila nueva
        .RowData(r) = mobjGNComp.IVKardex(ix)
        
        'Visualiza los valores predeterminados
        .TextMatrix(r, COL_CODBODEGA) = mobjGNComp.IVKardex(ix).CodBodega
        .TextMatrix(r, COL_CANT) = mobjGNComp.IVKardex(ix).Cantidad
        '.TextMatrix(r, COL_CANT) = mobjGNComp.GNTrans.CantidadPre
        
        'Copia de la fila anterior
        If r > .FixedRows Then
'            .TextMatrix(r, COL_CODBODEGA) = .TextMatrix(r - 1, COL_CODBODEGA)  '*** MAKOTO 16/dic/00 Eliminado por que lo hace en objeto GNComprobante.AddIVKardex
            .TextMatrix(r, COL_G1) = .TextMatrix(r - 1, COL_G1)
            .TextMatrix(r, COL_G2) = .TextMatrix(r - 1, COL_G2)
            .TextMatrix(r, COL_G3) = .TextMatrix(r - 1, COL_G3)
            .TextMatrix(r, COL_G4) = .TextMatrix(r - 1, COL_G4)
            .TextMatrix(r, COL_G5) = .TextMatrix(r - 1, COL_G5)
        End If
        
        'Ubica cursor en la primera columna visible y editable/seleccionable
        ' y vacio
        If .Rows > .FixedRows Then
            .Row = r
            .COL = .FixedCols
            For ix = .FixedCols To .Cols - 1
                If (Val(.ColData(ix)) >= 0) And _
                   (.ColWidth(ix) > 0) And _
                   (.ColHidden(ix) = False) And _
                   (Len(.TextMatrix(r, ix)) = 0) Then
                    .COL = ix
                    Exit For
                End If
            Next ix
        End If
    End With
    PoneNumFila
    VisualizaTotal
    Exit Sub
errtrap:
    DispErr
    Exit Sub
End Sub

Private Sub AgregaItemsPromocion(ByVal Row As Long, ByVal COL As Long, codItemPromo As String, cantPromo As Currency, PrecioPromo As Currency)
    Dim obj As IVKardex, cod As String, bod As String
    Dim mover As Boolean
    On Error GoTo errtrap
    If Not IsObject(grd.RowData(Row)) Then Exit Sub
    mover = True
    With grd
        Set obj = .RowData(Row)
        .TextMatrix(Row, COL_DESC) = codItemPromo
        If obj.CodInventario <> "-" And obj.CodInventario <> "+" Then
            cod = CogeSoloCodigo(Trim$(.Text))
            If mobjGNComp.GNTrans.IVTipoTrans = "E" Or mobjGNComp.GNTrans.IVTipoTrans = "Y" Then
                bod = CogeSoloBodega(Trim$(.Text))
            End If
            .TextMatrix(Row, COL_CODBODEGA) = bod
            If Len(cod) > 0 Then
            obj.CodInventario = cod
            End If
            If Len(.TextMatrix(Row, COL_CODBODEGA)) > 0 Then
                bod = .TextMatrix(Row, COL_CODBODEGA)
                obj.CodBodega = bod
            End If
            VisualizaItem Row, cod
            If mobjGNComp.GNTrans.IVTipoTrans = "E" Or mobjGNComp.GNTrans.IVTipoTrans = "Y" Then
                .TextMatrix(Row, COL_CANT) = cantPromo * -1
            Else
                .TextMatrix(Row, COL_CANT) = cantPromo
            End If
            obj.Cantidad = cantPromo
            .TextMatrix(Row, COL_PU) = PrecioPromo
            CalculaDetalle Row, COL_PU
        Else
            .TextMatrix(Row, COL_DESC) = obj.Nota
        End If
    End With

    VisualizaTotal
    
    If mobjGNComp.GNTrans.IVAvanceLinea Then
        If mover = True Then
            MueveColumna
        End If
    Else
        MueveColumna
    End If
    Exit Sub
errtrap:
    DispErr
    Exit Sub
End Sub


'jeaa 02/09/2008
Private Function Precio_PredeterminadoMasIntereses(item As Object, EncontroItemEnDocFuente As Boolean) As Currency
    Dim P As Currency
    With item
        If grd.ColData(COL_PU) = -1 And mobjGNComp.IdTransFuente <> 0 Then
            P = PrecioIVK_DocFuente(.CodInventario, EncontroItemEnDocFuente)
        End If
        If Not EncontroItemEnDocFuente Then   'Si no encuentra en item en Doc Fuete pone al precio predeterminado
        
            If mobjGNComp.GNTrans.codPantalla = "IVBQDPR" Then
                If mobjGNComp.GNTrans.IVRetornaNumPrecio Then
                    If PrecioSelec = 0 Then
                        P = .Precio(mobjGNComp.GNTrans.IVPrecioPre) ' para sacar el precio predeterminado
                    Else
                        P = .Precio(PrecioSelec) ' para sacar el precio predeterminado
                    End If
                Else
                    If PrecioSelec = 0 Then
                        P = .Precio(mobjGNComp.GNTrans.IVPrecioPre) ' para sacar el precio predeterminado
                    Else
                        P = PrecioSelec
                    End If
                End If
            Else
                If mobjGNComp.GNTrans.IVPrecioPre > 0 Then   ' Agregado Oliver
                    P = .Precio(mobjGNComp.GNTrans.IVPrecioPre) ' para sacar el precio predeterminado
                Else
                    If PrecioSelec = 0 Then
                        If mobjGNComp.PrecioRef = 0 Then
                            P = 0                                       ' en caso de no tener precio predeterminado no saca precio
                        Else
                                P = .Precio(mobjGNComp.PrecioRef) ' para sacar el precio predeterminado
                        End If
                    Else
                        P = PrecioSelec
                    End If
                End If
            End If
            P = P * mobjGNComp.Cotizacion(.CodMoneda) / mobjGNComp.Cotizacion("")  'Convierte en moneda del comprobante
        End If

    Dim PrecioIVA As Currency
    Dim ValorInteres As Currency
    Dim Plazo As Integer
    Dim CodIVA As String
    Dim tasa As Double
    Dim poriva As Double
    
            Dim AuxDesct As Double, ivgrupo As Integer
            ivgrupo = CInt(mobjGNComp.Empresa.GNOpcion.ObtenerValor("DescuentoNumIVGrupo")) + 1
            If Not mobjGNComp.GNTrans.IVVisualizarNumPagos Then
                AuxDesct = mobjGNComp.Empresa.DesctoPCxIV(.CodGrupo(ivgrupo), mobjGNComp.CodClienteRef)
            Else
                AuxDesct = mobjGNComp.Empresa.DesctoNumPagosxIV(.CodGrupo(ivgrupo), mobjGNComp.numdias)
            End If
    
    If AuxDesct <> 99 And mobjGNComp.Estado1 <> 1 Then
        If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("TipoCalculoInteres")) > 0 Then
            If mobjGNComp.Empresa.GNOpcion.ObtenerValor("TipoCalculoInteres") = 0 Then
                tasa = IIf(Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("TasaMensual")) > 0, mobjGNComp.Empresa.GNOpcion.ObtenerValor("TasaMensual"), 0)
            Else
                tasa = ObtieneInteresTabla
            End If
        Else
            tasa = IIf(Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("TasaMensual")) > 0, mobjGNComp.Empresa.GNOpcion.ObtenerValor("TasaMensual"), 0)
        End If
        'tasa = IIf(Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("TasaMensual")) > 0, mobjGNComp.Empresa.GNOpcion.ObtenerValor("TasaMensual"), 0)
        CodIVA = IIf(Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("RecDescParaIVA")) > 0, mobjGNComp.Empresa.GNOpcion.ObtenerValor("RecDescParaIVA"), "IVA")
        If mobjGNComp.VerificaConfigIVAenTrans(CodIVA) Then
            poriva = mobjGNComp.Empresa.GNOpcion.PorcentajeIVA
        Else
            poriva = 0
        End If
        Plazo = mobjGNComp.numdias
        'ntxPrecioIVA.value =
        If tasa <> 0 Then
            P = P + (P * tasa / 100)
    '        P = (P * Plazo * tasa / 100)
        Else
            P = P
    '        ValorInteres = 0
        End If
    End If
    End With
    
    Precio_PredeterminadoMasIntereses = P
End Function

Private Function ObtieneInteresTabla() As Double
    Dim sql As String, rs As Recordset
    sql = " select interes from GNInteres where " & mobjGNComp.numdias & " between desde and hasta"
    Set rs = mobjGNComp.Empresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        ObtieneInteresTabla = rs.Fields("interes")
    End If
End Function

Public Sub ConfigColsVisible_PrecioModificable()
    Dim v As Long, v2 As Long, v3 As Long
    Dim i As Integer
    v = mobjGNComp.GNTrans.ColVisible
    v2 = mobjGNComp.GNTrans.ColEditable
    v3 = mobjGNComp.GNTrans.ColSeleccionable
    
    With grd
        .ColData(COL_PU) = 1
        If (Not CBool(v2 And &H80010000)) And CBool(v3 And &H80010000) Then   'para el caso de que es solo selecionable
            .ColData(COL_PU) = 2
        End If
        
        For i = 1 To .Rows - 1
            If Not .IsSubtotal(i) Then
                grd.Cell(flexcpBackColor, i, COL_PU, i, COL_PU) = vbWhite
            End If
        Next i
        
    End With
End Sub

