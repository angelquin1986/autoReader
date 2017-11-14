VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.UserControl PCDocCHP 
   ClientHeight    =   2145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4905
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   2145
   ScaleWidth      =   4905
   Begin VSFlex7Ctl.VSFlexGrid grd 
      Align           =   1  'Align Top
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4905
      _cx             =   8652
      _cy             =   3625
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Cols            =   10
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
      ExplorerBar     =   0
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
      Begin VB.CommandButton cmdBotonDatos 
         Caption         =   "Datos"
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   1620
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Menu mnuDetalle 
      Caption         =   "&Detalle"
      Begin VB.Menu mnuAgregar 
         Caption         =   "&Agregar fila"
      End
      Begin VB.Menu mnuEliminar 
         Caption         =   "&Eliminar fila"
      End
      Begin VB.Menu mnulin1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActNumRef 
         Caption         =   "Actualiza Num Refe"
      End
   End
End
Attribute VB_Name = "PCDocCHP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit


'Ubicación de columnas
Private Const COL_NUMFILA = 0
Private Const COL_CODPROVCLI = 1
Private Const COL_PROVCLI = 2
Private Const COL_CODFORMA = 3
Private Const COL_NUMLETRA = 4
Private Const COL_VALOR = 5
Private Const COL_FECHAEMI = 6
Private Const COL_PLAZO = 7
Private Const COL_FECHAVENCI = 8
Private Const COL_OBSERVA = 9
'AUC 01/06/07
Private Const COL_CODVEN = 10
Private Const COL_VENDEDOR = 11
Private Const COL_DATOSADI = 12

Private WithEvents mobjGNComp As GNComprobante
Attribute mobjGNComp.VB_VarHelpID = -1
Private mobjGNTrans As GNTrans 'AUC 19/10/2005
Attribute mobjGNTrans.VB_VarHelpID = -1
Private mbooPorCobrar As Boolean
Private mbooModoProveedor As Boolean
Private mbooSI As Boolean
Private mbooProvCliVisible As Boolean
Private mstrCodProvCli As String
Private mstrCodProvCli2 As String           'Agregado para usar cuando se grabe los PCKardexCHP y no asigne todo al un solo cliente o proveedor
Private gncAux As GNComprobante


'Eventos
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event PorAgregarFila(ByRef valorPre As Currency)
Event AgregarFilaAuto(ByRef Cancel As Boolean)  '*** MAKOTO 12/dic/00 Agregado
Event CambiaNombreTrans()
Event VerificaPendiente(ByRef CodPCProvCli As String)       'jeaa 31/10/2005
Private gVerificaPendientes As Boolean '**** Agregado Esteban 31/10/2005
Event AgregarFilasPrestamo(ByRef Cancel As Boolean)  '*** MAKOTO 12/dic/00 Agregado
Event AgregarFilaDocs1(ByRef Cancel As Boolean, codforma As String, valor As Currency)  '*** jeaa 17/04/2008

Event PorAgregarFilaconPagoInicial(ByRef valorPre As Currency)
Event ValorPagoInicial(ByRef valorPre As Currency)
Event PorAgregarFilaporCredito(ByRef valorPre As Currency)
Event PorAgregarFilaDescXFormaCP(ByRef valorPre As Currency, codforma As String)

'Agregado AUC 01/11/2005
Private mIDPcGrupo1 As Long
Private mIDPcGrupo2 As Long
Private mIDPcGrupo3 As Long
Private midpcgrupo4 As Long
Private mCodforma As String
Private BandTarjeta As Boolean
Private numfila As Long
Private numTotalFila As Long
Private OtrasFilas As Long
Dim ANCHO_COLS(0 To 12) As Long
Private FilaGrilla As Long


Private Sub ConfigCols()
    Dim i As Integer
    With grd
        .FormatString = "^#|<Código|<Nombre|<Forma|<#Doc|>Valor" & _
                        "|<F.Emisión|>Plazo|<F.Venci.|<Observación|<CodVend.|<Vendedor|<Datos Adic."
                    'AUC 01/06/07 agregado codven  y vendedor
        GetColsWidth
        .ColWidth(COL_NUMFILA) = 500
        .ColWidth(COL_CODPROVCLI) = 1200                'Cod.Proveedor/Cliente
        .ColWidth(COL_PROVCLI) = 1800                   'Proveedor/Cliente
        .ColWidth(COL_CODFORMA) = 1200                   'CodForma
        .ColWidth(COL_NUMLETRA) = 1800                   'NumLetra
        .ColWidth(COL_VALOR) = COLANCHO_CUR             'valor
        .ColWidth(COL_FECHAEMI) = COLANCHO_FECHA        'F.Emisión
        .ColWidth(COL_PLAZO) = 1200  'antes 600                     'Plazo
        .ColWidth(COL_FECHAVENCI) = COLANCHO_FECHA      'F.Vencimiento
        .ColWidth(COL_OBSERVA) = 2000               'Observación
        'AUC 01/06/07
        .ColWidth(COL_CODVEN) = 1200                'Cod.vendedor
        .ColWidth(COL_VENDEDOR) = 1800               'vendedor
        .ColWidth(COL_DATOSADI) = 2800               'datos adicionales
        
        If mobjGNComp.GNTrans.CodPantalla = "PCGN" Then
            .ColHidden(COL_CODPROVCLI) = False
            .ColHidden(COL_PROVCLI) = False
        Else
            .ColHidden(COL_CODPROVCLI) = Not mbooProvCliVisible
            .ColHidden(COL_PROVCLI) = Not mbooProvCliVisible
        End If
        
        
        If mobjGNComp.GNTrans.IVDatosAdicionales Or mobjGNComp.GNTrans.TSDatosAdicionales Or mobjGNComp.GNTrans.TSDatosAdicionalesCHR Then
            .ColHidden(COL_DATOSADI) = False
        Else
            .ColHidden(COL_DATOSADI) = True
        End If
        
        .ColDataType(COL_VALOR) = flexDTCurrency
        .ColDataType(COL_FECHAEMI) = flexDTDate
        .ColDataType(COL_PLAZO) = flexDTShort
        .ColDataType(COL_FECHAVENCI) = flexDTDate

        
        
        'No modificables/Longitud maxima de campo
        .ColData(COL_CODPROVCLI) = 20
'        .ColData(COL_PROVCLI) = -1          'No modificable
        .ColData(COL_PROVCLI) = 40          '*** MAKOTO 18/jul/00 Modificado. para permitir seleccionar por nombre
        .ColData(COL_CODFORMA) = 5
        .ColData(COL_NUMLETRA) = 20
        .ColData(COL_VALOR) = 17        '9,999,999,999,999
        .ColData(COL_FECHAEMI) = 14     '1999/12/31
        .ColData(COL_PLAZO) = 4
        .ColData(COL_FECHAVENCI) = 14
        .ColData(COL_OBSERVA) = 80
        'AUC 01/06/07
        .ColData(COL_CODVEN) = 20
        .ColData(COL_VENDEDOR) = 40
        'AUC 19/06/07
        If mobjGNComp.GNTrans.Modulo = "IV" Then
            .ColHidden(COL_CODVEN) = True
            .ColHidden(COL_VENDEDOR) = True
            If Not mobjGNComp.GNTrans.IVDatosAdicionales Then
                .ColHidden(COL_DATOSADI) = True
            End If
        ElseIf mobjGNComp.GNTrans.Modulo = "TS" Then
            If Not mobjGNComp.GNTrans.TSDatosAdicionales And Not mobjGNComp.GNTrans.TSDatosAdicionalesCHR Then
                .ColHidden(COL_DATOSADI) = True
            End If
            
            
            If Not mobjGNComp.GNTrans.TSPideCobrador Then
                .ColHidden(COL_CODVEN) = True
                .ColHidden(COL_VENDEDOR) = True
            End If
        ElseIf mobjGNComp.GNTrans.Modulo = "AF" Then
            .ColHidden(COL_CODVEN) = True
            .ColHidden(COL_VENDEDOR) = True
        End If
        
         'If mobjGNComp.Empresa.GNOpcion.ObtenerValor("BloquearCuotas") = "1" Then
         If mobjGNComp.GNTrans.ivBloquearCuotas Then
            For i = COL_NUMLETRA To COL_VENDEDOR
                .ColData(i) = -1
            Next i
         End If
         
        
        
        
        If .Rows > .FixedRows Then .Row = .FixedRows
        If .Rows > .FixedRows Then .col = COL_CODFORMA
       If mobjGNComp.GNTrans.CodPantalla = "IVRT" Then
           For i = 0 To .Cols - 1
                .ColWidth(i) = ANCHO_COLS(i)
            Next i
        End If
        
    End With
    ConfigColsFormato
End Sub

Private Sub ConfigColsFormato()
    If mobjGNComp Is Nothing Then Exit Sub
    With grd
        .ColFormat(COL_FECHAEMI) = mobjGNComp.Empresa.GNOpcion.FormatoFecha
        .ColFormat(COL_FECHAVENCI) = .ColFormat(COL_FECHAEMI)
        .ColFormat(COL_VALOR) = mobjGNComp.FormatoMoneda
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
    Dim i As Integer

    cmdBotonDatos.Visible = False
    If mobjGNComp Is Nothing Then Exit Sub
    'Cuando es solo ver, deshabilita grid
    If mobjGNComp.SoloVer Then
        grd.Editable = flexEDNone
    Else
        grd.Editable = flexEDKbdMouse
    
'        Actualiza las listas
'        grd.ColComboList(COL_CODPROVCLI) = mobjGNComp.Empresa.ListaPCProvCliParaFlex(mbooBandProveedor)    '*** MAKOTO 23/oct/00
'        grd.ColComboList(COL_CODPROVCLI) = mobjGNComp.Empresa.ListaPCProvCliParaFlexEx(True, True)
'
         ActualizarFormaCobroPago 'AUC cambiado 1/11/2005

'    grd.ColComboList(COL_CODFORMA) = mobjGNComp.Empresa.ListaTSFormaCobroPagoParaFlex(mbooPorCobrar)
    


'    If mobjGNComp.GNTrans.TSVisibleF8 Or mobjGNComp.GNTrans.IVProvCliPorFila Or mobjGNComp.GNTrans.codPantalla = "TSIE" Then
    If mobjGNComp.GNTrans.Modulo = "IV" Then
        If mbooProvCliVisible And mobjGNComp.GNTrans.TSProvCliPorFila Then     'Agregado Oliver 17/dic/2003 solo es nececsario cuando esta la columna de Nombre visible
            '*** MAKOTO 18/jul/00 Agregado. Para permitir seleccionar por Nombre de prov/cli
        '    grd.ColComboList(COL_PROVCLI) = mobjGNComp.Empresa.ListaPCProvCliParaFlex2(mbooModoProveedor)
        
           grd.ColComboList(COL_CODPROVCLI) = mobjGNComp.Empresa.ListaPCProvCliParaFlexEx(True, True)
           grd.ColComboList(COL_PROVCLI) = mobjGNComp.Empresa.ListaPCProvCliParaFlex2Ex(True, True)        '*** MAKOTO 23/oct/00
           
        End If
    ElseIf mobjGNComp.GNTrans.Modulo = "PV" Or mobjGNComp.GNTrans.Modulo = "CL" Then
            If mbooProvCliVisible Then    'Agregado Oliver 17/dic/2003 solo es nececsario cuando esta la columna de Nombre visible
                grd.ColComboList(COL_CODPROVCLI) = mobjGNComp.Empresa.ListaPCProvCliParaFlexEx(mobjGNComp.GNTrans.ProveedorVisible, mobjGNComp.GNTrans.ClienteVisible)
                grd.ColComboList(COL_PROVCLI) = mobjGNComp.Empresa.ListaPCProvCliParaFlex2Ex(mobjGNComp.GNTrans.ProveedorVisible, mobjGNComp.GNTrans.ClienteVisible)        '*** MAKOTO 23/oct/00
            End If
    Else
        If mobjGNComp.GNTrans.TSProvCliPorFila Then
            If mbooProvCliVisible Then    'Agregado Oliver 17/dic/2003 solo es nececsario cuando esta la columna de Nombre visible
                
                grd.ColComboList(COL_CODPROVCLI) = mobjGNComp.Empresa.ListaPCProvCliParaFlexEx(mobjGNComp.GNTrans.ProveedorVisible, mobjGNComp.GNTrans.ClienteVisible)
                grd.ColComboList(COL_PROVCLI) = mobjGNComp.Empresa.ListaPCProvCliParaFlex2Ex(mobjGNComp.GNTrans.ProveedorVisible, mobjGNComp.GNTrans.ClienteVisible)        '*** MAKOTO 23/oct/00
            End If
        End If
    End If
            'AUC 01/06/007
        grd.ColComboList(COL_CODVEN) = mobjGNComp.Empresa.ListaFCVendedorParaFlex
        grd.ColComboList(COL_VENDEDOR) = mobjGNComp.Empresa.ListaFCVendedorParaFlex2
    ConfigColsFormato       'Llama esta para actualizar formato de moneda

        If mobjGNComp.GNTrans.ivBloquearCuotas Then
            For i = 1 To COL_VENDEDOR
                grd.ColData(i) = -1
            Next i
            HabilitarCtrlsGrupoCajaFormaCobro False
         End If

    End If
End Sub


Private Sub cmdBotonDatos_Click()
Dim cod As String, i As Long
Dim tsf As TSFormaCobroPago
    If mobjGNComp.GNTrans.Modulo = "IV" Then
        For i = 1 To mobjGNComp.CountPCKardex
            Set tsf = mobjGNComp.Empresa.RecuperaTSFormaCobroPago(mobjGNComp.PCKardexCHP(i).codforma)
            If Not tsf Is Nothing Then
                If tsf.DatosAdicionales Then
                    If InStr(1, UCase(tsf.NombreForma), "TAR") > 0 Then
                        If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "ITAL") > 0 Then
                            If Not Len(mobjGNComp.PCKardexCHP(i).Numcheque) > 0 Then
'                                cmdBotonDatos.Caption = frmDatosCobro.InicioCHP(mobjGNComp, True, i)
                            Else
 '                               cmdBotonDatos.Caption = frmDatosCobro.InicioCHP(mobjGNComp, True, i)
                            End If
                        Else
'                            If Not Len(mobjGNComp.PCKardexCHP(i).numCheque) > 0 Then
'                                cmdBotonDatos.Caption = frmDatosCobro.InicioCHP(mobjGNComp, True, numfila)
'                            End If
                        End If
                    Else
                        If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "ITAL") > 0 Then
                            If Not Len(mobjGNComp.PCKardexCHP(i).Numcheque) > 0 Then
'                                cmdBotonDatos.Caption = frmDatosCobro.InicioCHP(mobjGNComp, False, numfila)
                            End If
                        Else
                            If Not Len(mobjGNComp.PCKardexCHP(i).Numcheque) > 0 Then
'                                cmdBotonDatos.Caption = frmDatosCobro.InicioCHP(mobjGNComp, False, numfila)
                            End If
                        End If
                        grd.TextMatrix(FilaGrilla, COL_DATOSADI) = cmdBotonDatos.Caption
                    End If
                    
                End If
            End If
        Next i
    Else
        If mobjGNComp.EsNuevo Then
            numTotalFila = 0
            For i = 1 To mobjGNComp.CountPCKardexCHP
                If mobjGNComp.PCKardexCHP(i).IdAsignadoPCK = 0 Then
                    numTotalFila = i
                    Exit For
                End If
            Next i
            'numTotalFila = numTotalFila + numfila
            numTotalFila = i
        Else
            numTotalFila = 1
            For i = 1 To mobjGNComp.CountPCKardexCHP
                If mobjGNComp.PCKardexCHP(i).IdAsignadoPCK = 0 Then
                    numTotalFila = numTotalFila + 1
                End If
            Next i
'            numTotalFila = numfila
        End If

'        cmdBotonDatos.Caption = frmDatosCobro.InicioCHP(mobjGNComp, BandTarjeta, numTotalFila)
        grd.TextMatrix(FilaGrilla, COL_DATOSADI) = cmdBotonDatos.Caption
        grd.SetFocus
    End If
End Sub

Private Sub grd_Click()
    RaiseEvent Click
    numfila = grd.Row
'    numTotalFila = grd.Row + OtrasFilas
End Sub

Private Sub grd_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub grd_GotFocus()
    Dim Cancel As Boolean
    FlexGridGotFocusColor grd
    
    If grd.Editable And grd.Rows <= grd.FixedRows Then
        If mobjGNComp.GNTrans.CodPantalla = "GENROL" Then
            RaiseEvent AgregarFilaAuto(Cancel)
                If Not Cancel Then
                    AgregaFilaRol
                End If
        Else
            RaiseEvent AgregarFilaAuto(Cancel)  'Pregunta al contenedor si permite agregar la primera fila automáticamente o no
            If Not Cancel Then
                AgregaFila       'Si dice que sí }, agrega la primera fila
            End If
        End If
    End If
End Sub

Private Sub grd_LostFocus()
    FlexGridGotFocusColor grd
End Sub

Private Sub grd_KeyDown(KeyCode As Integer, Shift As Integer)
    If mobjGNComp Is Nothing Then Exit Sub
    If mobjGNComp.SoloVer Then Exit Sub
    
    Select Case KeyCode
    Case vbKeyInsert
        AgregaFila
        KeyCode = 0
    Case vbKeyDelete
        EliminaFila
        grd.SetFocus
        KeyCode = 0
    Case vbKeyReturn
    
    Case TECLA_CLICKDERECHO                     '*** MAKOTO 30/nov/00
        grd_MouseDown vbRightButton, Shift, 0, 0
    End Select

    RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub grd_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    
End Sub

Private Sub grd_KeyPressEdit(ByVal Row As Long, ByVal col As Long, KeyAscii As Integer)
    '*** MAKOTO 03/oct/2000
    ValidarTeclaFlexGrid grd, Row, col, KeyAscii, True
End Sub

Private Sub grd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mobjGNComp Is Nothing Then Exit Sub
    If mobjGNComp.SoloVer Then Exit Sub
    
    If Button And vbRightButton Then
        UserControl.PopupMenu mnuDetalle, , x, y
    End If
End Sub

Private Sub grd_AfterEdit(ByVal Row As Long, ByVal col As Long)
    Dim obj As PCKardexCHP, cod As String, i As Long, id As Long, cad  As String
    Dim tsf As TSFormaCobroPago, ix As Long, numP As Integer, gncp As GNComprobante
    Dim AuxDesct  As Long, ivgrupo As Integer, pc As PCProvCli, AuxDesctOri  As Long
    Dim CodFormaAnt As String, codFormaNue As String, j As Long, mensaje As String
    On Error GoTo ErrTrap
    numfila = Row
    
    OtrasFilas = 0
    For ix = 1 To mobjGNComp.CountPCKardexCHP
        'verifica si es anticipo
        If mobjGNComp.PCKardexCHP(ix).idAsignado <> 0 Then
            OtrasFilas = OtrasFilas + 1
        End If
    Next ix
    numfila = Row + OtrasFilas
    
    
    numTotalFila = Row + OtrasFilas
    If Not IsObject(grd.RowData(Row)) Then Exit Sub
    With grd
        Set obj = .RowData(Row)
        Select Case col
    'Case COL_CODFORMA

        
        
        
        Case COL_CODPROVCLI
            obj.CodProvCli = Trim$(.Text)
            CambiarNombreTrans          '*** MAKOTO 17/feb/01 Agregado
            RaiseEvent VerificaPendiente(obj.CodProvCli) '''    'jeaa 31/10/2005
        Case COL_PROVCLI                    '*** MAKOTO 18/jul/00 Agregado para permitir seleccionar por nombre
            cod = CogeSoloCodigo(Trim$(.Text))
            If Len(cod) > 0 Then
                If VisualizaProvCli(Row, cod) Then
                    obj.CodProvCli = cod
                    CambiarNombreTrans          '*** MAKOTO 17/feb/01 Agregado
                    RaiseEvent VerificaPendiente(obj.CodProvCli) '''    'jeaa 31/10/2005
                End If
            End If
        Case COL_CODFORMA
            If mobjGNComp.GNTrans.IVDescXFormaCP Then
                If Row = 1 Then
                    If Trim$(.Text) <> mobjGNComp.codforma Then
                        MsgBox " Forma de Pago Diferente a la seleccionada anteriormente"
                        .TextMatrix(Row, col) = mobjGNComp.codforma
                        obj.codforma = mobjGNComp.codforma
                    End If
                    
                End If
            Else
                CodFormaAnt = obj.codforma
                obj.codforma = Trim$(.Text)
                codFormaNue = Trim$(.Text)
                Set tsf = mobjGNComp.Empresa.RecuperaTSFormaCobroPago(obj.codforma)
                If Not tsf Is Nothing Then
                    If tsf.ConsiderarComoEfectivo Then
                        grd.TextMatrix(Row, COL_FECHAVENCI) = mobjGNComp.FechaTrans
                        obj.FechaVenci = mobjGNComp.FechaTrans
                        grd.TextMatrix(Row, COL_PLAZO) = "0"
                    End If
                    
                End If
                
            End If
            If mobjGNComp.GNTrans.IVControlaPrecioxFormaCobro Then
                Set tsf = mobjGNComp.Empresa.RecuperaTSFormaCobroPago(obj.codforma)
                If Not tsf Is Nothing Then
                        If tsf.ControlPrecios Then
                            For ix = 1 To mobjGNComp.CountIVKardex
                                numP = BuscaNumeroPrecio(ix)
                                If Mid$(tsf.ListaPrecios, numP, 1) = "0" Then
                                    grd.TextMatrix(ix, COL_CODFORMA) = CodFormaAnt
                                    obj.codforma = CodFormaAnt
                                    Err.Raise ERR_INVALIDO, "GNComprobante.RemovePCKardex", _
                                    "Con la forma de cobro " & codFormaNue & " no puede seleccionar el precio " & numP & Chr(13) & _
                                    " en la fila " & ix & " código Item: " & mobjGNComp.IVKardex(ix).CodInventario
                                
                                End If
                            Next ix
                        End If
                End If
                Set tsf = Nothing
            End If
            
        If mobjGNComp.GNTrans.IVDatosAdicionales Or mobjGNComp.GNTrans.TSDatosAdicionales Or mobjGNComp.GNTrans.TSDatosAdicionalesCHR Then
            Set tsf = mobjGNComp.Empresa.RecuperaTSFormaCobroPago(grd.TextMatrix(Row, COL_CODFORMA))
            If Not tsf Is Nothing Then
                    If tsf.DatosAdicionales Then
                        cmdBotonDatos_Click

                        BandTarjeta = InStr(1, UCase(tsf.NombreForma), "TARJETA") > 0
                        If tsf.DatosAdicionales Then
                            If grd.TextMatrix(Row, COL_DATOSADI) = "NO" Or grd.TextMatrix(Row, COL_DATOSADI) = "Vacío" Then
                                grd.TextMatrix(Row, COL_DATOSADI) = "Vacío"
                            Else
                                grd.TextMatrix(Row, COL_DATOSADI) = "O.K."
                                obj.NumLetra = obj.Numcheque
                                grd.TextMatrix(Row, COL_NUMLETRA) = obj.NumLetra
                                cad = ""
                                For i = 1 To mobjGNComp.CountPCKardexCHP
                                    If mobjGNComp.PCKardexCHP(i).idAsignado <> 0 Then
                                        id = mobjGNComp.Empresa.RecuperarIDGncomprobantexIdAsignado(mobjGNComp.PCKardexCHP(i).idAsignado)
                                        Set gncp = mobjGNComp.Empresa.RecuperaGNComprobante(id)
                                        If Not gncp Is Nothing Then
                                            cad = cad & gncp.CodTrans & "-" & gncp.numtrans & "/"
                                            obj.CodVendedor = gncp.CodVendedor
                                        End If
                                    End If
                                Next i
                                Set gncp = Nothing
                                If Len(cad) > 0 Then
                                cad = Mid$(cad, 1, Len(cad) - 1)
                                cad = "Por pago de: " & cad
                                End If
                                grd.TextMatrix(Row, COL_OBSERVA) = Left(cad, 80)
                                obj.Observacion = Left(cad, 80)
                                
                            End If
                             
                        Else
'                            grd.Cell(flexcpData, Row, COL_DATOSADI) = -1
                             grd.TextMatrix(Row, COL_DATOSADI) = "NO"
                            grd.Cell(flexcpBackColor, Row, COL_DATOSADI, Row, COL_DATOSADI) = &H80000018
                        End If
                    End If
                End If
                Set tsf = Nothing
            End If
        VisualizarPlazo obj, Row
        If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "ITAL") > 0 Then
            If Row = 1 Then
                If mobjGNComp.codforma <> .Text Then
                Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(mobjGNComp.CodClienteRef)
                If Not pc Is Nothing Then
                    AuxDesctOri = mobjGNComp.Empresa.VerificaDesctoPCxIVxFecha(mobjGNComp.CodDescuento, mobjGNComp.codforma)
                    AuxDesct = mobjGNComp.Empresa.VerificaDesctoPCxIVxFecha(mobjGNComp.CodDescuento, Trim$(.Text))
                    If AuxDesctOri <> AuxDesct Then
                        MsgBox "Selecciono otra forma de cobro diferente a la inicial, " & Chr(13) & "debe actualizar la forma de cobro,  tienen diferente % de descuento"
                        .Text = mobjGNComp.codforma
                        Exit Sub
                    End If
                End If
                End If
            End If
        End If
        obj.codforma = Trim$(.Text)
        Case COL_NUMLETRA
            obj.NumLetra = Trim$(.Text)
        Case COL_VALOR
            '*** Asignamos en ValidateEdit para verificar y cancelar si es necesario
            '*** MAKOTO 29/ene/01 Mod.
            '*** Sin embargo hay que asignar de nuevo para que guarde con el valor redondeado
            Dim CodTrans As String, numtrans As Long, valor As Currency, valorAntes As Currency
            For ix = 1 To mobjGNComp.CountPCKardex
                If ix = Row Then
                    If mobjGNComp.PCKardexCHP(ix).id <> 0 Then
                        If mobjGNComp.Empresa.VerificarCambioCobroPago(mobjGNComp.PCKardexCHP(ix).id, CodTrans, numtrans, valor) Then
                            If valor <> grd.ValueMatrix(ix, COL_VALOR) Then
                                grd.TextMatrix(ix, COL_VALOR) = mobjGNComp.PCKardexCHP(ix).Debe
                                Err.Raise ERR_INVALIDO, "GNComprobante.RemovePCKardex", _
                                "No se puede Modificar el documento debido a que existen cobros o pagos asignados " & Chr(13) & "con la Transacción: " & CodTrans & "-" & numtrans & "del Cliente: " & mobjGNComp.PCKardexCHP(ix).CodProvCli
                            End If
                            'Exit For
                        End If
                    End If
                    'If mcolPCKardex.item(ix) Is obj Then Exit For
                End If
            Next ix
            
            
            If mbooPorCobrar Then
                obj.Debe = MiCCur(.Cell(flexcpTextDisplay, Row, col))
            Else
                obj.Haber = MiCCur(.Cell(flexcpTextDisplay, Row, col))
            End If
        Case COL_FECHAEMI
            If Len(.Text) = 0 Then .Text = obj.GNComprobante.FechaTrans
            .TextMatrix(Row, COL_FECHAVENCI) = CDate(.Text) + .ValueMatrix(Row, COL_PLAZO)
            obj.FechaVenci = CDate(.TextMatrix(Row, COL_FECHAVENCI))
            obj.FechaEmision = CDate(.Text)
        Case COL_PLAZO
            .TextMatrix(Row, COL_FECHAVENCI) = CDate(.TextMatrix(Row, COL_FECHAEMI)) + .value
            obj.FechaVenci = CDate(.TextMatrix(Row, COL_FECHAVENCI))
            If mobjGNComp.GNTrans.IVPlazoProvCli Then
                Set tsf = mobjGNComp.Empresa.RecuperaTSFormaCobroPago(obj.codforma)
                If Not (tsf Is Nothing) Then
                    If Not tsf.ConsiderarComoEfectivo Then
                        Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(obj.CodProvCli)
                        If Not (pc Is Nothing) Then
                            If .ValueMatrix(Row, COL_PLAZO) > pc.DiasPlazo Then
                                MsgBox "El plazo máximo para el cliente " & pc.nombre & ", es de " & pc.DiasPlazo & " días"
                                obj.FechaVenci = obj.FechaEmision + pc.DiasPlazo
                                .TextMatrix(Row, COL_FECHAVENCI) = obj.FechaVenci
                                .TextMatrix(Row, COL_PLAZO) = pc.DiasPlazo
                            End If
                        End If
                        Set pc = Nothing
                    End If
                End If
            End If
            
            If mobjGNComp.GNTrans.TSVerificaNuevoPlazo Then
                Set tsf = mobjGNComp.Empresa.RecuperaTSFormaCobroPago(obj.codforma)
                If Not (tsf Is Nothing) Then
                    For i = 1 To mobjGNComp.CountPCKardex
                        id = mobjGNComp.Empresa.RecuperarIDGncomprobantexIdAsignado(mobjGNComp.PCKardexCHP(i).idAsignado)
                        Set gncp = mobjGNComp.Empresa.RecuperaGNComprobante(id)
                        If Not gncp Is Nothing Then
                            For j = 1 To gncp.CountPCKardex
                                If gncp.PCKardexCHP(j).id = mobjGNComp.PCKardexCHP(i).idAsignado Then
                                    If CDate(grd.TextMatrix(Row, COL_FECHAVENCI)) > (DateAdd("d", tsf.NuevoPlazo, gncp.PCKardexCHP(j).FechaVenci)) Then
                                        MsgBox "No se puede dar más días de crédito del pago original, para la transsacción" & Chr(13) & gncp.CodTrans & "-" & gncp.numtrans & "  vence el " & DateAdd("d", tsf.NuevoPlazo, gncp.PCKardexCHP(j).FechaVenci)
                                        .TextMatrix(Row, COL_FECHAVENCI) = DateAdd("d", tsf.NuevoPlazo, gncp.PCKardexCHP(j).FechaVenci)
                                        .TextMatrix(Row, COL_PLAZO) = CDate(.TextMatrix(Row, COL_FECHAVENCI)) - CDate(.TextMatrix(Row, COL_FECHAEMI))
                                        obj.FechaVenci = CDate(.Text)
                                        j = gncp.CountPCKardex
                                        i = mobjGNComp.CountPCKardex
                                    End If
                                End If
                            Next j
                        End If
                        Set gncp = Nothing
                    Next i
                End If
                Set tsf = Nothing
            End If
            
        Case COL_FECHAVENCI
            .TextMatrix(Row, COL_PLAZO) = CDate(.Text) - CDate(.TextMatrix(Row, COL_FECHAEMI))
            obj.FechaVenci = CDate(.Text)
            If mobjGNComp.GNTrans.IVPlazoProvCli Then
                Set tsf = mobjGNComp.Empresa.RecuperaTSFormaCobroPago(obj.codforma)
                If Not (tsf Is Nothing) Then
                    If Not tsf.ConsiderarComoEfectivo Then
                        Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(obj.CodProvCli)
                        If Not (pc Is Nothing) Then
                            If .ValueMatrix(Row, COL_PLAZO) > pc.DiasPlazo Then
                                MsgBox "El plazo máximo para el cliente " & pc.nombre & ", es de " & pc.DiasPlazo & " días"
                                obj.FechaVenci = obj.FechaEmision + pc.DiasPlazo
                                .TextMatrix(Row, COL_FECHAVENCI) = obj.FechaVenci
                                .TextMatrix(Row, COL_PLAZO) = pc.DiasPlazo
                            End If
                        End If
                        Set pc = Nothing
                    End If
                End If
            End If
            
'            If mobjGNComp.GNTrans.TSVerificaNuevoPlazo Then
'                Set TSF = mobjGNComp.Empresa.RecuperaTSFormaCobroPago(obj.codforma)
'                If Not (TSF Is Nothing) Then
'                    For i = 1 To mobjGNComp.CountPCKardex
'                        id = mobjGNComp.Empresa.RecuperarIDGncomprobantexIdAsignado(mobjGNComp.PCKardexCHP(i).idasignado)
'                        Set gncp = mobjGNComp.Empresa.RecuperaGNComprobante(id)
'                        If Not gncp Is Nothing Then
'                            For j = 1 To gncp.CountPCKardex
'                                If gncp.PCKardexCHP(j).id = mobjGNComp.PCKardexCHP(i).idasignado Then
'                                    If CDate(grd.TextMatrix(Row, COL_FECHAVENCI)) > (DateAdd("d", TSF.NuevoPlazo, gncp.PCKardexCHP(j).FechaVenci)) Then
'                                        mensaje = "No se puede dar más credito del pago original, o sea hasta " & DateAdd("d", TSF.NuevoPlazo, gncp.PCKardexCHP(j).FechaVenci)
'                                        frmMensajeAutoriza.Inicio mensaje, mobjGNComp, bandaut
'                                        .TextMatrix(Row, COL_FECHAVENCI) = DateAdd("d", TSF.NuevoPlazo, gncp.PCKardexCHP(j).FechaVenci)
'                                        .TextMatrix(Row, COL_PLAZO) = CDate(.TextMatrix(Row, COL_FECHAVENCI)) - CDate(.TextMatrix(Row, COL_FECHAEMI))
'                                        obj.FechaVenci = CDate(.Text)
'                                        j = gncp.CountPCKardex
'                                        i = mobjGNComp.CountPCKardex
'                                    End If
'                                End If
'                            Next j
'                        End If
'                        Set gncp = Nothing
'                    Next i
'                End If
'                Set TSF = Nothing
'            End If
            
            
        Case COL_OBSERVA
            obj.Observacion = Trim$(.Text)
            'AUC 01/06/07 agregado vendedor
        Case COL_CODVEN
            obj.CodVendedor = Trim$(.Text)
        Case COL_VENDEDOR
            cod = CogeSoloCodigo(Trim$(.Text))
            If Len(cod) > 0 Then
                If VisualizaVendedor(Row, cod) Then
                    obj.CodVendedor = cod
                End If
            End If
        Case COL_DATOSADI
            numfila = Row
            numTotalFila = Row + OtrasFilas
            If cmdBotonDatos.Caption = "O.K." Then
                grd.TextMatrix(Row, col) = "O.K."
            End If
        End Select
    End With

    VisualizaTotal
    MueveColumna
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub grd_CellChanged(ByVal Row As Long, ByVal col As Long)
    '*** MAKOTO 29/ene/01 Agregado.
    FlexGridRedondear grd, Row, col
    FilaGrilla = Row
End Sub


'*** MAKOTO 18/jul/00
'Para coger solo código de prov/cli cuando selecciona nombre
''xxxxxxxxxxxxxxxxx [nnnnnnn]'    --> Devuelve solo 'nnnnnnn'
Private Function CogeSoloCodigo(ByVal Desc As String) As String
    Dim s As String, i As Long
    i = InStrRev(Desc, "[")
    If i > 0 Then s = Mid$(Desc, i + 1)
    If Len(s) > 0 Then s = Left$(s, Len(s) - 1)
    CogeSoloCodigo = s
End Function


Private Sub VisualizaTotal()
    grd.subtotal flexSTSum, -1, COL_VALOR, , grd.BackColorFrozen, vbYellow, , "Total", , True
    grd.Refresh
End Sub

Private Sub MueveColumna()
    Dim c As Long
    With grd
        If .Rows > .FixedRows Then
            For c = .col + 1 To .Cols - 1
                If .ColData(c) >= 0 And .ColWidth(c) > 0 And (Not .ColHidden(c)) Then
                    While .ColData(c) = 14
                            c = c + 1
                    Wend
                    .col = c
                    Exit Sub
                End If
            Next c
    
            If .Row < .Rows - 1 Then .Row = .Row + 1
    
            For c = .FixedCols To .Cols - 1
                If .ColData(c) >= 0 And .ColWidth(c) > 0 And (Not .ColHidden(c)) Then
                    .col = c
                    Exit Sub
                End If
            Next c
        End If
    End With
End Sub

Private Sub grd_BeforeEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
    Dim tsf As TSFormaCobroPago, cod As String
    If mobjGNComp.SoloVer Then Exit Sub
    cmdBotonDatos.Visible = False
    numfila = Row
    numTotalFila = Row + OtrasFilas
    If Not grd.IsSubtotal(Row) Then
        If cmdBotonDatos.Caption = "O.K." Then
            grd.TextMatrix(Row, COL_DATOSADI) = "O.K."
            cmdBotonDatos.Caption = "Datos"
        Else
            If grd.TextMatrix(Row, COL_DATOSADI) <> "O.K." And grd.TextMatrix(Row, COL_DATOSADI) <> "NO" Then
                grd.TextMatrix(Row, COL_DATOSADI) = "Vacío"
            End If
        End If
    End If
    'Cuando es una columna no modificable
    If grd.Rows > grd.FixedRows Then
        Cancel = (grd.ColData(col) < 0) Or grd.IsSubtotal(Row) Or grd.ColHidden(col)
    Else
        Cancel = True
    End If
    
    If grd.Cell(flexcpData, Row, col) = -1 Then Cancel = True
    If Cancel Then Exit Sub
    
    Select Case col
    Case COL_CODFORMA
'''        If mobjGNComp.GNTrans.codPantalla = "IVPVTS" Then
'''            Cancel = True
'''        End If
''If mobjGNComp.CodFormnaCP <> grd.TextMatrix(Row, Col) Then
''    MsgBox "Selecciono otra forma de cobro diferente a la inicial"
''
''End If
    Case COL_FECHAEMI
        If mobjGNComp.CodTrans <> "CLND" And mobjGNComp.CodTrans <> "CLNC" And mobjGNComp.CodTrans <> "PVND" And mobjGNComp.CodTrans <> "PVNC" Then
            If Not mobjGNComp.GNTrans.IVDesbloquearFechas Then
                Cancel = True
            End If
        End If
    Case COL_DATOSADI
            Set tsf = mobjGNComp.Empresa.RecuperaTSFormaCobroPago(grd.TextMatrix(Row, COL_CODFORMA))
            If Not tsf Is Nothing Then
                BandTarjeta = InStr(1, UCase(tsf.NombreForma), "TARJETA") > 0
                If tsf.DatosAdicionales Then
'                        cod = frmDatosCobro.Inicio(mobjGNComp)
                    With grd

                        ' -- Redimensionar y posicionar el boton
                        cmdBotonDatos.Move (.Left + .CellLeft), _
                                       (.Top - 10 + (.RowHeight(0) * (.Row - .TopRow + 1))), _
                                       (.CellWidth), _
                                       (.CellHeight - 10)

                        ' -- Hacer visible y pasarle el foco
                        cmdBotonDatos.Visible = True
                        cmdBotonDatos.Enabled = True
                        cmdBotonDatos.SetFocus
                    End With
                Else
''                    grd.Cell(flexcpData, Row, COL_DATOSADI) = -1
                    grd.Cell(flexcpBackColor, Row, COL_DATOSADI, Row, COL_DATOSADI) = &H80000018
                End If
            End If
            Set tsf = Nothing
    End Select
    
    'Longitud maxima
    grd.EditMaxLength = grd.ColData(col)
End Sub

'*** MAKOTO 17/feb/01 Agregado
Private Sub CambiarNombreTrans()
    mobjGNComp.AsignarNombreTrans
    RaiseEvent CambiaNombreTrans
End Sub

Public Sub AgregaFila()
    Dim r As Long, r2 As Long, ix As Long, i As Long, v As Currency
    Dim pck As PCKardexCHP
    Dim Ven As FCVendedor
    Dim tsf As TSFormaCobroPago
    
    cmdBotonDatos.Visible = False
    RaiseEvent PorAgregarFila(v)        'Para calcular el valor predeterminado

    On Error GoTo ErrTrap
    BandTarjeta = False
    'Llama a agregar un objeto PCKardexCHP antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
    ix = mobjGNComp.AddPCKardexCHP
    
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
        numTotalFila = r + OtrasFilas
        'Asigna la referencia al nuevo objeto a la fila nueva
        Set pck = mobjGNComp.PCKardexCHP(ix)
        .RowData(r) = pck
        cmdBotonDatos.Visible = False
        'Proporciona el valor predeterminado        '*** MAKOTO 05/oct/00 Modificado
        If v > 0 Then
            If mbooPorCobrar Then pck.Debe = Abs(v)
            .TextMatrix(r, COL_VALOR) = pck.Debe
            pck.Debe = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
        ElseIf v < 0 Then
            If Not mbooPorCobrar Then pck.Haber = Abs(v)
            .TextMatrix(r, COL_VALOR) = pck.Haber
            pck.Haber = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
        End If
        .TextMatrix(r, COL_CODPROVCLI) = pck.CodProvCli     '*** MAKOTO 14/oct/00
        VisualizaProvCli r, pck.CodProvCli                  '***
        
        
        '.TextMatrix(r, COL_NOMBRE) = mobjGNComp.
        '***Agregado. 17/Ago/2004. Angel
        '***Para que se inserte la fila pero con la forma de pago predeterminada
        If Len(mobjGNComp.GNTrans.CodFormaPre) > 0 Then
            If Not mobjGNComp.GNTrans.IVDescXFormaCP Then
                If Len(mobjGNComp.codforma) = 0 Then
                    .TextMatrix(r, COL_CODFORMA) = mobjGNComp.GNTrans.CodFormaPre
                Else
                    .TextMatrix(r, COL_CODFORMA) = mobjGNComp.codforma
                End If
            Else
                If Len(mobjGNComp.codforma) > 0 Then
                    .TextMatrix(r, COL_CODFORMA) = mobjGNComp.codforma
                    BloqueaColumnaCodForma mobjGNComp.codforma
                End If
            End If
            pck.codforma = Trim$(.TextMatrix(r, COL_CODFORMA))
        Else
            If mbooPorCobrar Then
                If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroAnticipo")) > 0 Then
                    If mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroAnticipo") <> pck.codforma Then
                        pck.codforma = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroAnticipo")
                    End If
                End If
            Else
                If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoAnticipo")) > 0 Then
                    If mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoAnticipo") <> pck.codforma Then
                        pck.codforma = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoAnticipo")
                    End If
                End If
            End If
            .TextMatrix(r, COL_CODFORMA) = pck.codforma
        End If
        If mobjGNComp.GNTrans.IVDatosAdicionales Or mobjGNComp.GNTrans.TSDatosAdicionales Or mobjGNComp.GNTrans.TSDatosAdicionalesCHR Then
            Set tsf = mobjGNComp.Empresa.RecuperaTSFormaCobroPago(grd.TextMatrix(r, COL_CODFORMA))
            If Not tsf Is Nothing Then
                BandTarjeta = InStr(1, UCase(tsf.NombreForma), "TARJETA") > 0
                If tsf.DatosAdicionales Then
                        ' -- Redimensionar y posicionar el boton
                    cmdBotonDatos.Move (.Left + .CellLeft), _
                                   (.Top - 10 + (.RowHeight(0) * (.Row - .TopRow + 1))), _
                                   (.CellWidth), _
                                   (.CellHeight - 10)
                                       
                    ' -- Hacer visible y pasarle el foco
                    cmdBotonDatos.Visible = True
                    cmdBotonDatos.Enabled = True
                    cmdBotonDatos.SetFocus
                Else
'                    grd.Cell(flexcpData, r, COL_DATOSADI) = -1
                    grd.Cell(flexcpBackColor, r, COL_DATOSADI, r, COL_DATOSADI) = &H80000018
                    grd.TextMatrix(r, COL_DATOSADI) = "NO"
                End If
            End If
            Set tsf = Nothing
        End If
        
        .TextMatrix(r, COL_NUMLETRA) = pck.NumLetra
        pck.NumLetra = .TextMatrix(r, COL_NUMLETRA)
        .TextMatrix(r, COL_FECHAEMI) = pck.FechaEmision
        
        
        
        
        
        If mobjGNComp.GNTrans.CodPantalla = "TSIEE" Then
            pck.Orden = mobjGNComp.CountPCKardex
        ElseIf mobjGNComp.GNTrans.CodPantalla = "TSICHP" Then
            pck.Orden = mobjGNComp.CountPCKardex + 1
        End If
        
        'AUC 19/07/06
        If mobjGNComp.GNTrans.IVPideVendedor And Len(mobjGNComp.CodVendedor) > 0 Then
            Set Ven = mobjGNComp.Empresa.RecuperaFCVendedor(mobjGNComp.CodVendedor)
            .TextMatrix(r, COL_CODVEN) = Ven.CodVendedor
            .TextMatrix(r, COL_VENDEDOR) = Ven.nombre
            pck.IdVendedor = GNComprobante.IdVendedor
        Else
            If Len(mobjGNComp.CodVendedor) > 0 Then
                .TextMatrix(r, COL_VENDEDOR) = mobjGNComp.CodVendedor
                pck.IdVendedor = GNComprobante.IdVendedor
                
            End If
        End If
        '------------
        VisualizarPlazo pck, r
        '.TextMatrix(r, COL_PLAZO) = pck.FechaVenci - pck.FechaEmision
        '.TextMatrix(r, COL_FECHAVENCI) = pck.FechaVenci
        If Not GNComprobante.GNTrans.IVDesbloquearFechas Then
            .Cell(flexcpBackColor, r, COL_FECHAEMI) = grd.BackColorFixed
            .Cell(flexcpBackColor, r, COL_FECHAEMI) = grd.BackColorFixed
         End If
        
        .Row = r
        If .Rows > .FixedRows Then
            .col = .FixedCols
            'Busca la primera columna
            For i = .FixedCols To .Cols - 1
                If .ColData(i) >= 0 And (Not .ColHidden(i)) And .ColWidth(i) > 0 Then
                    .col = i
                    Exit For
                End If
            Next i
        End If
    End With
    
    PoneNumFila
    VisualizaTotal
salida:
    Set pck = Nothing
    Set Ven = Nothing
'    grd.SetFocus
    Exit Sub
ErrTrap:
    Set pck = Nothing
    Set Ven = Nothing
    DispErr
    GoTo salida
End Sub




Private Sub EliminaFila()
    Dim msg As String, r As Long
    On Error GoTo ErrTrap

    If grd.Row < grd.FixedRows Then Exit Sub        '*** MAKOTO 07/feb/01 Mod.
    If grd.Rows <= grd.FixedRows Then Exit Sub
    If grd.IsSubtotal(grd.Row) Then Exit Sub
    
    r = grd.Row
    msg = "Desea eliminar la fila #" & r & "?"
    If MsgBox(msg, vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    'Remueve de la colección de objeto
    mobjGNComp.RemovePCKardexCHP 0, grd.RowData(r)
    
    'Elimina del grid
    grd.RemoveItem r
    PoneNumFila
    grd.subtotal flexSTClear
    VisualizaTotal
    
    CambiarNombreTrans          '*** MAKOTO 17/feb/01 Agregado
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub PoneNumFila()
    Dim i As Long
    With grd
        For i = .FixedRows To .Rows - 1
            If Not .IsSubtotal(i) Then .TextMatrix(i, 0) = i
                If .ColData(COL_CODFORMA) = -1 And Not .IsSubtotal(i) Then
                    .Cell(flexcpBackColor, i, COL_CODFORMA, i, COL_CODFORMA) = &H80000018
                End If
        Next i
    End With
End Sub

Private Sub grd_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
    Dim f1 As Date, f2 As Date, obj As PCKardexCHP
    On Error GoTo ErrTrap
    
    With grd
        Set obj = .RowData(Row)
        Select Case col
        Case COL_CODPROVCLI
            If Not VisualizaProvCli(Row, .EditText) Then Cancel = True
        Case COL_VALOR
            '*** MAKOTO 29/ene/01 Modificado
            If Len(.EditText) > 0 Then
                If Not IsNumeric(.EditText) Then
                    MsgBox "Ingrese un valor numérico.", vbExclamation
                    Cancel = True
                Else
                    If mbooPorCobrar Then
                        obj.Debe = Val(.EditText)
                    Else
                        obj.Haber = Val(.EditText)
                    End If
                End If
            End If
        Case COL_FECHAEMI
            If Len(.EditText) > 0 Then
                If Not IsDate(.EditText) Then
                    MsgBox "Fecha incorrecta. (ejm. 1999/12/31)", vbExclamation
                    Cancel = True
                End If
            End If
        Case COL_FECHAVENCI
            If Len(.EditText) > 0 Then
                If Not IsDate(.EditText) Then
                    MsgBox "Fecha incorrecta. (ejm. 1999/12/31)", vbExclamation
                    Cancel = True
                End If
                
                If IsDate(.TextMatrix(Row, COL_FECHAEMI)) And _
                   IsDate(.TextMatrix(Row, COL_FECHAVENCI)) Then
                    f1 = CDate(.TextMatrix(Row, COL_FECHAEMI))
                    f2 = CDate(.EditText)
                    If f1 > f2 Then
                        MsgBox "La fecha de vencimiento no puede ser menor a la fecha de emisión.", vbExclamation
                        Cancel = True
                    End If
                End If
            End If
        Case COL_CODVEN 'AUC 01/06/07
            If Not VisualizaVendedor(Row, .EditText) Then Cancel = True
        End Select
    End With
    Exit Sub
ErrTrap:
    Cancel = True
    DispErr
    Exit Sub
End Sub

Private Function VisualizaProvCli(Row As Long, cod As String) As Boolean
    Dim pc As PCProvCli
    On Error GoTo ErrTrap

    If Len(cod) = 0 Then Exit Function
    
    Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(cod)
    'jeaa 31/10/2005
    If mobjGNComp.GNTrans.IVProvCliPorFila Then
        RaiseEvent VerificaPendiente(pc.CodProvCli)
    End If
    With pc
        grd.TextMatrix(Row, COL_CODPROVCLI) = .CodProvCli   '*** MAKOTO 18/jul/00 Agregado para cuando seleccione por nombre
        grd.TextMatrix(Row, COL_PROVCLI) = .nombre
        VisualizaProvCli = True
    End With
    Set pc = Nothing
    Exit Function
ErrTrap:
    'Si no encuentra el codigo
    If Err.Number = 3021 Then
        MsgBox MSG_ERR_NOENCUENTRA & "(" & cod & ")", vbInformation
    Else
        DispErr
    End If
    Set pc = Nothing
    Exit Function
End Function





Public Property Get GNComprobante() As GNComprobante
    Set GNComprobante = mobjGNComp
End Property

Public Property Set GNComprobante(obj As GNComprobante)
    Set mobjGNComp = obj

    If Not mobjGNComp.EsNuevo Then
'        Visualizar
        ConfigCols
        If mobjGNComp.GNTrans.CodPantalla <> "TSIEE" And mobjGNComp.GNTrans.CodPantalla <> "TSIECC" Then
            VisualizaDesdeObjeto
        Else
            If mobjGNComp.Empresa.mbooBandPorCobrar Then
                VisualizaDesdeObjetoAuto (True)
                mobjGNComp.Empresa.mbooBandPorCobrar = False
            Else
                VisualizaDesdeObjetoAuto (False)
                mobjGNComp.Empresa.mbooBandPorCobrar = True
            End If
            
        End If
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
    cmdBotonDatos.Visible = False
    With grd
        For i = .FixedRows To .Rows - 1
            .RowData(i) = 0
        Next i
        .Rows = .FixedRows
    End With
End Sub

'Public Sub Visualizar()
'    Dim i As Long
'
'    grd.Redraw = False
'
'    'Visualiza los detalles que está en GNComprobante
'    Set grd.DataSource = mobjGNComp.ListaPCKardex
'    ConfigCols
'
'    'Asigna referencia al objeto PCKardexCHP a cada fila de grid
'    With grd
'        For i = 1 To mobjGNComp.CountPCKardex
'            .RowData(i) = mobjGNComp.PCKardexCHP(i)
'        Next i
'    End With
'
'    PoneNumFila
'    VisualizaTotal
'    grd.Redraw = True
'    grd.Refresh
'End Sub

Public Sub Aceptar()
    Dim i As Long, obj As PCKardexCHP
    On Error GoTo ErrTrap

    'Pasa los detalles al objeto
    With grd
        For i = .FixedRows To .Rows - 1
            If Not .IsSubtotal(i) Then
                Set obj = .RowData(i)
                
                '*** MAKOTO 12/oct/00 Modificado
                If Len(mstrCodProvCli2) > 0 Then
                    If Not obj.GNComprobante.GNTrans.IVProvCliPorFila Then
                        obj.CodProvCli = mstrCodProvCli2
                    End If
                End If
                obj.Orden = i
            End If
        Next i
        Set obj = Nothing
    End With
    Exit Sub
ErrTrap:
    Select Case Err.Number
    Case Else
        DispErr
    End Select
    Exit Sub
End Sub

Public Property Let ProvCliVisible(value As Boolean)
    mbooProvCliVisible = value
    PropertyChanged "ProvCliVisible"
    
    grd.ColHidden(COL_CODPROVCLI) = Not value
    grd.ColHidden(COL_PROVCLI) = Not value
End Property

Public Property Get ProvCliVisible() As Boolean
    ProvCliVisible = mbooProvCliVisible
End Property

Public Property Let CodProvCli(value As String)
    Dim i As Long, pck As PCKardexCHP
    On Error GoTo ErrTrap
    
    If Len(value) > 0 And _
        (Not mobjGNComp.GNTrans.IVProvCliPorFila) Then  '*** MAKOTO 12/oct/00 Modificado
        With grd
            'Asigna el codigo de prov/cli a todos los detalles de PCKardexCHP
            For i = .FixedRows To .Rows - 1
                If Not .IsSubtotal(i) Then
                    Set pck = .RowData(i)
                    If Not pck Is Nothing Then pck.CodProvCli = value
                End If
            Next i
        End With
    End If

    mstrCodProvCli = value
    mstrCodProvCli2 = value
    
    PropertyChanged "CodProvCli"
    Exit Property
ErrTrap:
    DispErr
    Exit Property
End Property

Public Property Get CodProvCli() As String
    CodProvCli = mstrCodProvCli
End Property
Public Property Let PorCobrar(value As Boolean)
    mbooPorCobrar = value
    PropertyChanged "PorCobrar"
End Property

Public Property Get PorCobrar() As Boolean
    PorCobrar = mbooPorCobrar
End Property

Public Property Let ModoProveedor(value As Boolean)
    mbooModoProveedor = value
    PropertyChanged "ModoProveedor"
End Property

Public Property Get ModoProveedor() As Boolean
    ModoProveedor = mbooModoProveedor
End Property
Private Sub mnuActNumRef_Click()
    ActualizaNumeroReferencia mobjGNComp.numDocRef
End Sub

Private Sub mnuAgregar_Click()
    AgregaFila
    grd.SetFocus
End Sub

Private Sub mnuEliminar_Click()
    EliminaFila
    grd.SetFocus
End Sub

Private Sub mobjGNComp_CotizacionCambiado()
    mobjGNComp_MonedaCambiado
End Sub

Private Sub mobjGNComp_MonedaCambiado()
    Dim r As Long, pck As PCKardexCHP
    On Error GoTo ErrTrap
    
    ConfigColsFormato
    
    'Reasigna todos los valores
    With grd
        For r = .FixedRows To .Rows - 1
            If Not .IsSubtotal(r) Then
                Set pck = .RowData(r)
                If Not pck Is Nothing Then
                    If mbooPorCobrar Then
'                        pck.Debe = .ValueMatrix(r, COL_VALOR)
                        pck.Debe = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))   '*** MAKOTO 29/ene/01 Mod.
                    Else
'                        pck.Haber = .ValueMatrix(r, COL_VALOR)
                        pck.Haber = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))  '*** MAKOTO 29/ene/01 Mod.
                    End If
                End If
            End If
        Next r
    End With
    
'    VisualizaDesdeObjeto
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Public Sub VisualizaDesdeObjeto()
    Dim pck As PCKardexCHP, pc As PCProvCli, i As Long, s As String
    Dim r As Long, NombrePC As String, Ven As FCVendedor, nombreVen As String
    With grd
        .Redraw = False
        Limpiar
        For i = 1 To mobjGNComp.CountPCKardexCHP
            Set pck = mobjGNComp.PCKardexCHP(i)
            If Len(mstrCodProvCli) = 0 Then   'Agregado esta condicion para que cuando se recupera carge a la variable de la propiedad Docs.codProvCli
                'Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(pck.CodProvCli)
                'If Not (pc Is Nothing) Then
                mstrCodProvCli = pck.CodProvCli
                'Set pc = Nothing
            End If
            If mobjGNComp.GNTrans.CodPantalla = "IVBQD2PCK" Then
                If pck.idAsignado = 0 Then
                    'Si hay que mostrar el nombre de proveedor/cliente
                    If Not .ColHidden(COL_PROVCLI) Then
                        'Recupera PCProvCli
                        Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(pck.CodProvCli)
                        If Not (pc Is Nothing) Then NombrePC = pc.nombre
                    End If
                        'AUC 01/06/07
                        Set Ven = mobjGNComp.Empresa.RecuperaFCVendedor(pck.CodVendedor)
                        If Not (Ven Is Nothing) Then
                            nombreVen = Ven.nombre
                        Else
                            nombreVen = ""
                        End If
                    
                    s = .Rows & vbTab & _
                            pck.CodProvCli & vbTab & _
                            NombrePC & vbTab & _
                            pck.codforma & vbTab & _
                            pck.NumLetra & vbTab & _
                            NullSiZero(pck.Debe + pck.Haber) & vbTab & _
                            pck.FechaEmision & vbTab & _
                            pck.FechaVenci - pck.FechaEmision & vbTab & _
                            pck.FechaVenci & vbTab & _
                            pck.Observacion & vbTab & pck.CodVendedor & vbTab & nombreVen
                            If mobjGNComp.SoloVer Then
                                If Len(pck.codBanco) > 0 Then
                                    If Len(pck.CodTarjeta) > 0 Then
                                        s = s & vbTab & pck.codBanco & "/" & pck.CodTarjeta
                                    Else
                                        s = s & vbTab & pck.codBanco & "/" & pck.Numcheque
                                    End If
                                End If
                                
                                
                            End If
                            .AddItem s
                            r = .Rows - 1
                            .RowData(r) = pck
                            '*** Oliver 29/01/2003 Agregado para que bloquee las columnas que no se modifiquen
                            BloquearPlazo pck, r
                End If
            Else
                If pck.IdAsignadoPCK = 0 Then
                    'Si hay que mostrar el nombre de proveedor/cliente
                    If Not .ColHidden(COL_PROVCLI) Then
                        'Recupera PCProvCli
                        Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(pck.CodProvCli)
                        If Not (pc Is Nothing) Then NombrePC = pc.nombre
                    End If
                        'AUC 01/06/07
                        Set Ven = mobjGNComp.Empresa.RecuperaFCVendedor(pck.CodVendedor)
                        If Not (Ven Is Nothing) Then
                            nombreVen = Ven.nombre
                        Else
                            nombreVen = ""
                        End If
                    
                    s = .Rows & vbTab & _
                            pck.CodProvCli & vbTab & _
                            NombrePC & vbTab & _
                            pck.codforma & vbTab & _
                            pck.NumLetra & vbTab & _
                            NullSiZero(pck.Debe + pck.Haber) & vbTab & _
                            pck.FechaEmision & vbTab & _
                            pck.FechaVenci - pck.FechaEmision & vbTab & _
                            pck.FechaVenci & vbTab & _
                            pck.Observacion & vbTab & pck.CodVendedor & vbTab & nombreVen
                            If mobjGNComp.SoloVer Then
                                If Len(pck.codBanco) > 0 Then
                                    If Len(pck.CodTarjeta) > 0 Then
                                        s = s & vbTab & pck.codBanco & "/" & pck.CodTarjeta
                                    Else
                                        s = s & vbTab & pck.codBanco & "/" & pck.Numcheque
                                    End If
                                End If
                                
                                
                            End If
                            .AddItem s
                            r = .Rows - 1
                            .RowData(r) = pck
                            '*** Oliver 29/01/2003 Agregado para que bloquee las columnas que no se modifiquen
                            BloquearPlazo pck, r
                End If
            End If
        Next i
        PoneNumFila
        VisualizaTotal
        .Redraw = True
    End With
    Set Ven = Nothing
End Sub


'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()

End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    ProvCliVisible = PropBag.ReadProperty("ProvCliVisible", True)
    PorCobrar = PropBag.ReadProperty("PorCobrar", True)
    grd.FontSize = PropBag.ReadProperty("FontSize", 10)
End Sub

Private Sub UserControl_Resize()
    'Ajusta el tamaño del grid
    grd.Height = UserControl.ScaleHeight
End Sub

Private Sub UserControl_Terminate()
    Set mobjGNComp = Nothing
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Enabled", UserControl.Enabled, True
    PropBag.WriteProperty "ProvCliVisible", ProvCliVisible
    PropBag.WriteProperty "PorCobrar", PorCobrar
    PropBag.WriteProperty "FontSize", grd.FontSize, 0
End Sub

Private Sub VisualizarPlazo(ByRef obj As PCKardexCHP, Row As Long)
    Dim tsf As TSFormaCobroPago
    Dim pc As PCProvCli, Plazo As PlazoPcGrupoxIVGrupo
    Dim item As IVinventario, ivgrupo As ivgrupo, numGrupo As Integer
    With grd
    Set tsf = mobjGNComp.Empresa.RecuperaTSFormaCobroPago(obj.codforma)
    If mobjGNComp.GNTrans.IVFiltroGrupoAlFact And mobjGNComp.GNTrans.IVUtilizarFiltroIvGDiasCred And Not tsf.ConsiderarComoEfectivo Then
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("FactGrupoItem")) > 0 Then
                If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("FactGrupoItem") = "1" Then
                    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("cboFacturaxGrupo")) > 0 Then
                        numGrupo = CInt(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("cboGrupoItem")) + 1
                        Set item = mobjGNComp.Empresa.RecuperaIVInventario(mobjGNComp.IVKardex(1).CodInventario)
                        Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(obj.CodProvCli)
                        Set Plazo = mobjGNComp.Empresa.RecuperaPlazoPCxIV(pc.CodDiasCredito & "," & item.CodGrupo(numGrupo))
                        If Not Plazo Is Nothing Then
                            obj.FechaVenci = DateAdd("d", Plazo.valor, obj.FechaEmision)
                        End If
                        Set item = Nothing
                        Set pc = Nothing
                        Set Plazo = Nothing

                    End If
                Else
                End If
            End If
    Else
        'cambiar plazo que tiene esta forma de cobro
        If Not mobjGNComp.GNTrans.IVPlazoProvCli Then
            'Set TSF = mobjGNComp.Empresa.RecuperaTSFormaCobroPago(obj.codforma)
            If Not (tsf Is Nothing) Then
                obj.FechaVenci = obj.FechaEmision + tsf.Plazo
            Else
                obj.FechaVenci = obj.FechaEmision
            End If
            'Set TSF = Nothing
        Else
            'Set TSF = mobjGNComp.Empresa.RecuperaTSFormaCobroPago(obj.codforma)
            If Not (tsf Is Nothing) Then
                If Not tsf.ConsiderarComoEfectivo Then
                    Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(obj.CodProvCli)
                    If Not (pc Is Nothing) Then
                        obj.FechaVenci = obj.FechaEmision + pc.DiasPlazo
                    Else
                        obj.FechaVenci = obj.FechaEmision
                    End If
                    Set pc = Nothing
                Else
                    obj.FechaVenci = obj.FechaEmision
                End If
            End If
        End If
    End If
        'Visualizando los cambios en las fechas
        .TextMatrix(Row, COL_PLAZO) = obj.FechaVenci - obj.FechaEmision
        .TextMatrix(Row, COL_FECHAVENCI) = obj.FechaVenci
    End With
    Set tsf = Nothing
    BloquearPlazo obj, Row
End Sub

'Agregado  para controlar si el plazo permite modificacion o no
Private Sub BloquearPlazo(ByRef obj As PCKardexCHP, Row As Long)
    Dim tsf As TSFormaCobroPago, i As Integer
    Dim item As IVinventario, numGrupo As Integer
    With grd
        'cambiar plazo que tiene esta forma de cobro
        Set tsf = mobjGNComp.Empresa.RecuperaTSFormaCobroPago(obj.codforma)
        
        
        
            If mobjGNComp.CodTrans <> "CLND" And mobjGNComp.CodTrans <> "CLNC" And mobjGNComp.CodTrans <> "PVND" And mobjGNComp.CodTrans <> "PVNC" Then
                If Not mobjGNComp.GNTrans.IVDesbloquearFechas Then
                    .Cell(flexcpData, Row, COL_FECHAEMI) = -1
                    .Cell(flexcpBackColor, Row, COL_FECHAEMI) = grd.BackColorFixed
                    .Cell(flexcpBackColor, Row, COL_FECHAEMI) = grd.BackColorFixed
                End If
            End If
        
        If Not (tsf Is Nothing) Then
           If Not tsf.CambiaFechaVenci Then
                .Cell(flexcpData, Row, COL_PLAZO) = -1
                .Cell(flexcpData, Row, COL_FECHAVENCI) = -1
                .Cell(flexcpBackColor, Row, COL_PLAZO) = grd.BackColorFixed
                .Cell(flexcpBackColor, Row, COL_FECHAVENCI) = grd.BackColorFixed
            
            Else
                .Cell(flexcpData, Row, COL_PLAZO) = 0
                .Cell(flexcpData, Row, COL_FECHAVENCI) = 0
                
                .Cell(flexcpBackColor, Row, COL_PLAZO) = vbWhite
                .Cell(flexcpBackColor, Row, COL_FECHAVENCI) = vbWhite
            End If
        Else
            .Cell(flexcpData, Row, COL_PLAZO) = -1
            .Cell(flexcpData, Row, COL_FECHAVENCI) = -1
        End If
        If mobjGNComp.GNTrans.ivBloquearCuotas Then
            For i = COL_NUMLETRA To COL_VENDEDOR
                grd.ColData(i) = -1
            Next i
            HabilitarCtrlsGrupoCajaFormaCobro False
        End If
        
        
        If mobjGNComp.GNTrans.IVFiltroGrupoAlFact And mobjGNComp.GNTrans.IVUtilizarFiltroIvGDiasCred Then
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("FactGrupoItem")) > 0 Then
                If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("FactGrupoItem") = "1" Then

                        .Cell(flexcpData, Row, COL_PLAZO) = -1
                        .Cell(flexcpData, Row, COL_FECHAVENCI) = -1
                        .Cell(flexcpBackColor, Row, COL_PLAZO) = grd.BackColorFixed
                        .Cell(flexcpBackColor, Row, COL_FECHAVENCI) = grd.BackColorFixed

                Else
                End If
            End If
        End If
        
       Set tsf = Nothing
    End With
End Sub

Public Sub EliminaFilaDocs(Optional ByVal fila As Integer)
    Dim msg As String, r As Long
    Dim Cancel As Boolean
    
    On Error GoTo ErrTrap
    If fila = 0 Then
        r = mobjGNComp.CountPCKardex
        grd.RemoveItem 1 'borra la linea de subtotal
    Else
        r = fila
    End If
    If r > 0 Then
        mobjGNComp.RemovePCKardex r, grd.RowData(r)
        grd.RemoveItem r
    End If
'Anulado 28/10/04 jeaa
'    If fila <> 0 Then
'        If grd.Rows - 1 = grd.FixedRows Then
'            RaiseEvent AgregarFilaAuto(Cancel)  'Pregunta al contenedor si permite agregar la primera fila automáticamente o no
'            If Not Cancel Then
'                AgregaFila       'Si dice que sí }, agrega la primera fila
'            End If
'        End If
'    End If
    
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Public Sub AgregaFilaPago(ByVal capital As Double, _
                                                ByVal NumPagos As Integer, _
                                                ByVal Pago As Double, _
                                                ByVal tasa As Double, _
                                                ByVal PeriodoAnio As Double, _
                                                ByVal DiasPeriodo As Double, _
                                                ByVal BandAmortiza As Boolean)
    Dim r As Long, r2 As Long, ix As Long, i As Long, v As Currency
    Dim j As Integer
    Dim pck As PCKardexCHP
    Dim mBandcapital  As Boolean, ContPago As Long
    Dim mRedondeo As Integer, mPosPunto As Integer
    Dim auxcapital As Double
    RaiseEvent PorAgregarFila(v)        'Para calcular el valor predeterminado

    On Error GoTo ErrTrap
        mPosPunto = InStr(1, mobjGNComp.FormatoMoneda, ".")
        If mPosPunto > 0 Then
            mRedondeo = Len(Mid$(mobjGNComp.FormatoMoneda, mPosPunto, Len(mobjGNComp.FormatoMoneda) - mPosPunto))
        End If
        
        mBandcapital = False
        ContPago = 1
        auxcapital = capital
        For j = 1 To NumPagos * 2
        DoEvents
'        Llama a agregar un objeto PCKardexCHP antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
            ix = mobjGNComp.AddPCKardex
            
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
    
                'Asigna la referencia al nuevo objeto a la fila nueva
                Set pck = mobjGNComp.PCKardexCHP(ix)
                .RowData(r) = pck
'                If v > 0 Then
                If mBandcapital Then
                    If BandAmortiza Then
                        .TextMatrix(r, COL_VALOR) = Round(Pago - (capital * (tasa / PeriodoAnio)), mRedondeo)
                        pck.Debe = Round(MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR)), mRedondeo)
                        capital = Round(capital - pck.Debe, mRedondeo)
                    Else
                        .TextMatrix(r, COL_VALOR) = Round(Pago - (auxcapital * (tasa / PeriodoAnio)), mRedondeo)
                        pck.Debe = Round(MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR)), mRedondeo)
                        capital = Round(capital - pck.Debe, mRedondeo)
                    End If
                    'si el saldo es diferente de 0 iguala el pago sumandolo o restandolo
                    If j = NumPagos * 2 And capital <> 0 Then
                        .TextMatrix(r, COL_VALOR) = Round(.ValueMatrix(r, COL_VALOR) + capital, mRedondeo)
                        pck.Debe = Round(MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR)), mRedondeo)
                    End If
                    
                    
                    If Len(mobjGNComp.GNTrans.CodFormaPre) > 0 Then
                        .TextMatrix(r, COL_CODFORMA) = mobjGNComp.GNTrans.CodFormaPre
                        pck.codforma = Trim$(.TextMatrix(r, COL_CODFORMA))
                    Else
                        .TextMatrix(r, COL_CODFORMA) = pck.codforma
                    End If
                    
                    mBandcapital = False
                    pck.NumLetra = pck.NumLetra & " - Cap - " & ContPago & "/" & NumPagos
                Else
                    If BandAmortiza Then
                        .TextMatrix(r, COL_VALOR) = Round(capital * (tasa / PeriodoAnio), mRedondeo)
                        pck.Debe = Round(MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR)), mRedondeo)
                    Else
                        .TextMatrix(r, COL_VALOR) = Round(auxcapital * (tasa / PeriodoAnio), mRedondeo)
                        pck.Debe = Round(MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR)), mRedondeo)
                    End If
                    
                    If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FornmaCobroInteresFinanciamiento")) > 0 Then
'                    If Len(mobjGNComp.GNTrans.IVInteresxCobrar) > 0 Then
                        .TextMatrix(r, COL_CODFORMA) = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FornmaCobroInteresFinanciamiento")
                       pck.codforma = Trim$(.TextMatrix(r, COL_CODFORMA))
                    Else
                        .TextMatrix(r, COL_CODFORMA) = pck.codforma
                    End If
                    pck.NumLetra = pck.NumLetra & " - Inter - " & ContPago & "/" & NumPagos
                    mBandcapital = True
                End If
               .TextMatrix(r, COL_NUMLETRA) = pck.NumLetra
                .TextMatrix(r, COL_CODPROVCLI) = pck.CodProvCli     '*** MAKOTO 14/oct/00
                VisualizaProvCli r, pck.CodProvCli
                
                
                '***Agregado. 17/Ago/2004. Angel
                '***Para que se inserte la fila pero con la forma de pago predeterminada
                
                
'                .TextMatrix(r, COL_NUMLETRA) = ContPago
                .TextMatrix(r, COL_FECHAEMI) = pck.FechaEmision
                VisualizarPlazoMasDias pck, r, ContPago * DiasPeriodo
                
                .Row = r
                If .Rows > .FixedRows Then
                    .col = .FixedCols
                    'Busca la primera columna
                    For i = .FixedCols To .Cols - 1
                        If .ColData(i) >= 0 And (Not .ColHidden(i)) And .ColWidth(i) > 0 Then
                            .col = i
                            Exit For
                        End If
                    Next i
                End If
        End With
        If mBandcapital = False Then
            ContPago = ContPago + 1
        End If
    Next j
    PoneNumFila
    VisualizaTotal
salida:
    Set pck = Nothing
    grd.SetFocus
    Exit Sub
ErrTrap:
    Set pck = Nothing
    DispErr
    GoTo salida
End Sub


Private Sub VisualizarPlazoMasDias(ByRef obj As PCKardexCHP, Row As Long, Optional dias As Long)
    Dim tsf As TSFormaCobroPago
    With grd
        'cambiar plazo que tiene esta forma de cobro
        Set tsf = mobjGNComp.Empresa.RecuperaTSFormaCobroPago(obj.codforma)
        If Not (tsf Is Nothing) Then
            obj.FechaVenci = obj.FechaEmision + tsf.Plazo + dias
        Else
            obj.FechaVenci = obj.FechaEmision + dias
        End If
        Set tsf = Nothing
        'Visualizando los cambios en las fechas
        .TextMatrix(Row, COL_PLAZO) = obj.FechaVenci - obj.FechaEmision
        .TextMatrix(Row, COL_FECHAVENCI) = obj.FechaVenci
    End With
    BloquearPlazo obj, Row
End Sub

Public Sub AgregaFilaEntrada(ByVal valor As Double)
    Dim r As Long, r2 As Long, ix As Long, i As Long, v As Currency
    Dim pck As PCKardexCHP

    RaiseEvent PorAgregarFilaconPagoInicial(v)        'Para calcular el valor predeterminado

    On Error GoTo ErrTrap
    
    'Llama a agregar un objeto PCKardexCHP antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
    ix = mobjGNComp.AddPCKardex
    
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

        'Asigna la referencia al nuevo objeto a la fila nueva
        Set pck = mobjGNComp.PCKardexCHP(ix)
        .RowData(r) = pck
        
        'Proporciona el valor predeterminado        '*** MAKOTO 05/oct/00 Modificado
'        If v > 0 Then
            If mbooPorCobrar Then pck.Debe = Abs(v)
            .TextMatrix(r, COL_VALOR) = v
            pck.Debe = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
'        ElseIf v < 0 Then
'            If Not mbooPorCobrar Then pck.Haber = Abs(v)
'            .TextMatrix(r, COL_VALOR) = Valor
'            pck.Haber = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
'        End If
        .TextMatrix(r, COL_CODPROVCLI) = pck.CodProvCli     '*** MAKOTO 14/oct/00
        VisualizaProvCli r, pck.CodProvCli                  '***
        
        '***Agregado. 17/Ago/2004. Angel
        '***Para que se inserte la fila pero con la forma de pago predeterminada en la configuracion IVFIN
        If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FornmaCobroCuotaInicial")) > 0 Then
            .TextMatrix(r, COL_CODFORMA) = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FornmaCobroCuotaInicial")
            pck.codforma = Trim$(.TextMatrix(r, COL_CODFORMA))
        Else
            .TextMatrix(r, COL_CODFORMA) = mobjGNComp.GNTrans.CodFormaPre 'pck.codforma
        End If
        pck.NumLetra = "Cuota Inicial"
        .TextMatrix(r, COL_NUMLETRA) = pck.NumLetra
        .TextMatrix(r, COL_FECHAEMI) = pck.FechaEmision
        'VisualizarPlazo pck, r
        '.TextMatrix(r, COL_PLAZO) = pck.FechaVenci - pck.FechaEmision
        .TextMatrix(r, COL_FECHAVENCI) = pck.FechaEmision
        pck.FechaVenci = pck.FechaEmision
        .TextMatrix(r, COL_PLAZO) = pck.FechaVenci - pck.FechaEmision
        
        .Row = r
        If .Rows > .FixedRows Then
            .col = .FixedCols
            'Busca la primera columna
            For i = .FixedCols To .Cols - 1
                If .ColData(i) >= 0 And (Not .ColHidden(i)) And .ColWidth(i) > 0 Then
                    .col = i
                    Exit For
                End If
            Next i
        End If
    End With
    
    PoneNumFila
    VisualizaTotal
salida:
    Set pck = Nothing
    grd.SetFocus
    Exit Sub
ErrTrap:
    Set pck = Nothing
    DispErr
    GoTo salida
End Sub

Public Sub EliminaFilaDocsSubtotal(Optional ByVal fila As Integer)
    Dim msg As String, r As Long
    Dim Cancel As Boolean
    
    On Error GoTo ErrTrap
    If fila = 0 Then
        r = mobjGNComp.CountPCKardex
        grd.RemoveItem 1 'borra la linea de subtotal
    Else
        r = fila
    End If
    If r > 0 Then
        grd.RemoveItem r
    End If
    
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub
'jeaa 21/07/2005
Private Sub ActualizaNumeroReferencia(ByVal numDocRef As String)
    Dim pck As PCKardexCHP, i As Long, msg As String
On Error GoTo ErrTrap
    msg = "Automáticamente cambiará la coumna #Doc " & Chr(13) & "es decir, el Número de Referencia de cada fila "
    msg = msg & vbCr & vbCr & "Desea continuar?"
    If MsgBox(msg, vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    MensajeStatus MSG_PREPARA, vbHourglass
    
    For i = 1 To mobjGNComp.CountPCKardex
        Set pck = mobjGNComp.PCKardexCHP(i)
        MensajeStatus "Procesándo #" & i & ": '" & pck.NumLetra & "'...", vbHourglass
        pck.NumLetra = numDocRef
    Next i
    Set pck = Nothing
    
    'Actualiza la pantalla
    VisualizaDesdeObjeto
        
    MensajeStatus
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    Exit Sub
End Sub
'jeaa 05/08/2005
Public Sub HabilitarCtrlsGrupoCaja(ByVal BandHabilita As Boolean)
Dim col As Integer, fil As Integer
    For col = 1 To 9
        If Not BandHabilita Then
            grd.ColData(col) = -1
            For fil = 1 To grd.Rows - 2
                        grd.Cell(flexcpBackColor, fil, col, fil, col) = &H80000018
            Next fil
        End If
    Next col
End Sub

Public Sub ActualizarFormaCobroPago()
    Dim pc As PCProvCli
    Dim NumGrupoControl As String
    
     Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(mobjGNComp.IdClienteRef)
     If (pc Is Nothing) Then
        'grd.ComboList = ""
        grd.ColComboList(COL_CODFORMA) = mobjGNComp.Empresa.ListaTSFormaCobroPagoParaFlex(mbooPorCobrar)
    Else
    'Selecciona  la forma de cobro pago si esta seleccionado
        If mobjGNComp.GNTrans.IVControlaCreditos And mobjGNComp.GNTrans.ClienteVisible Then
            Select Case mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroCli")
            Case 0
                NumGrupoControl = pc.CodGrupo1
            Case 1
                NumGrupoControl = pc.CodGrupo2
            Case 2
                NumGrupoControl = pc.CodGrupo3
            Case 3
                NumGrupoControl = pc.CodGrupo4
            End Select
            If mobjGNComp.BandCierre Then
                grd.ColComboList(COL_CODFORMA) = mobjGNComp.Empresa.ListaTSFormaCobroPagoParaFlexContolFormasSoloEfectivo(mbooPorCobrar, NumGrupoControl, True)
            
            Else
                grd.ColComboList(COL_CODFORMA) = mobjGNComp.Empresa.ListaTSFormaCobroPagoParaFlexContolFormas(mbooPorCobrar, NumGrupoControl)
            End If
        Else
           grd.ColComboList(COL_CODFORMA) = mobjGNComp.Empresa.ListaTSFormaCobroPagoParaFlex(mbooPorCobrar)
        End If
    End If
        Set pc = Nothing
End Sub
'AUC 01/06/07
Private Function VisualizaVendedor(Row As Long, cod As String) As Boolean
    Dim Ven As FCVendedor
    On Error GoTo ErrTrap

    If Len(cod) = 0 Then Exit Function
    
    Set Ven = mobjGNComp.Empresa.RecuperaFCVendedor(cod)
    
    With Ven
        grd.TextMatrix(Row, COL_CODVEN) = .CodVendedor
        grd.TextMatrix(Row, COL_VENDEDOR) = .nombre
        VisualizaVendedor = True
    End With
    Set Ven = Nothing
    Exit Function
ErrTrap:
    
    If Err.Number = 3021 Then
        MsgBox MSG_ERR_NOENCUENTRA & "(" & cod & ")", vbInformation
    Else
        DispErr
    End If
    Set Ven = Nothing
    Exit Function
End Function

Public Sub AgregaFilaPrestamo(ByVal valor As Double, ByVal dias As Integer, ByVal Num As Integer, ByVal numpago As Integer, ByVal FormaCobro As String)
    Dim r As Long, r2 As Long, ix As Long, i As Long, v As Currency
    Dim pck As PCKardexCHP

    RaiseEvent PorAgregarFila(v)        'Para calcular el valor predeterminado

    On Error GoTo ErrTrap
    
    'Llama a agregar un objeto PCKardexCHP antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
    ix = mobjGNComp.AddPCKardex
    
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

        'Asigna la referencia al nuevo objeto a la fila nueva
        Set pck = mobjGNComp.PCKardexCHP(ix)
        .RowData(r) = pck
        
        'Proporciona el valor predeterminado        '*** MAKOTO 05/oct/00 Modificado
            If mbooPorCobrar Then pck.Debe = Abs(v)
            .TextMatrix(r, COL_VALOR) = valor
            pck.Debe = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
        .TextMatrix(r, COL_CODPROVCLI) = pck.CodProvCli     '*** MAKOTO 14/oct/00
        VisualizaProvCli r, pck.CodProvCli                  '***
        
        '***Agregado. 17/Ago/2004. Angel
        '***Para que se inserte la fila pero con la forma de pago predeterminada en la configuracion IVFIN
        If Len(FormaCobro) > 0 Then
            .TextMatrix(r, COL_CODFORMA) = FormaCobro
            pck.codforma = Trim$(.TextMatrix(r, COL_CODFORMA))
        Else
            .TextMatrix(r, COL_CODFORMA) = pck.codforma
        End If
        pck.NumLetra = "Pago " & Num + 1 & " de " & numpago
        pck.Observacion = "Pago " & Num + 1 & " de " & numpago & " Cuota de " & mobjGNComp.Descripcion & " " & mobjGNComp.CodTrans & "-" & mobjGNComp.numtrans + 1
        .TextMatrix(r, COL_NUMLETRA) = pck.NumLetra
        .TextMatrix(r, COL_FECHAEMI) = pck.FechaEmision
        'VisualizarPlazo pck, r
        '.TextMatrix(r, COL_PLAZO) = pck.FechaVenci - pck.FechaEmision
        .TextMatrix(r, COL_FECHAVENCI) = DateAdd("d", dias, pck.FechaEmision)
        pck.FechaVenci = DateAdd("d", dias, pck.FechaEmision)
        .TextMatrix(r, COL_PLAZO) = pck.FechaVenci - pck.FechaEmision
        
        .Row = r
        If .Rows > .FixedRows Then
            .col = .FixedCols
            'Busca la primera columna
            For i = .FixedCols To .Cols - 1
                If .ColData(i) >= 0 And (Not .ColHidden(i)) And .ColWidth(i) > 0 Then
                    .col = i
                    Exit For
                End If
            Next i
        End If
    End With
    
    PoneNumFila
    VisualizaTotal
salida:
    Set pck = Nothing
    grd.SetFocus
    Exit Sub
ErrTrap:
    Set pck = Nothing
    DispErr
    GoTo salida
End Sub

Private Sub AgregaFilaDoc1()
    Dim r As Long, r2 As Long, ix As Long, i As Long, v As Currency
    Dim pck As PCKardexCHP
    Dim Ven As FCVendedor
    RaiseEvent PorAgregarFila(v)        'Para calcular el valor predeterminado
    On Error GoTo ErrTrap
    'Llama a agregar un objeto PCKardexCHP antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
    ix = mobjGNComp.AddPCKardex
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
        'Asigna la referencia al nuevo objeto a la fila nueva
        Set pck = mobjGNComp.PCKardexCHP(ix)
        .RowData(r) = pck
        'Proporciona el valor predeterminado        '*** MAKOTO 05/oct/00 Modificado
        If v > 0 Then
            If mbooPorCobrar Then pck.Debe = Abs(v)
            .TextMatrix(r, COL_VALOR) = pck.Debe
            pck.Debe = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
        ElseIf v < 0 Then
            If Not mbooPorCobrar Then pck.Haber = Abs(v)
            .TextMatrix(r, COL_VALOR) = pck.Haber
            pck.Haber = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
        End If
        .TextMatrix(r, COL_CODPROVCLI) = pck.CodProvCli     '*** MAKOTO 14/oct/00
        VisualizaProvCli r, pck.CodProvCli                  '***
        '.TextMatrix(r, COL_NOMBRE) = mobjGNComp.
        '***Agregado. 17/Ago/2004. Angel
        '***Para que se inserte la fila pero con la forma de pago predeterminada
        If Len(mobjGNComp.GNTrans.CodFormaPre) > 0 Then
            .TextMatrix(r, COL_CODFORMA) = mobjGNComp.GNTrans.CodFormaPre
            pck.codforma = Trim$(.TextMatrix(r, COL_CODFORMA))
        Else
            .TextMatrix(r, COL_CODFORMA) = pck.codforma
        End If
        .TextMatrix(r, COL_NUMLETRA) = pck.NumLetra
        .TextMatrix(r, COL_FECHAEMI) = pck.FechaEmision
        'AUC 19/07/06
        If mobjGNComp.GNTrans.IVPideVendedor And Len(mobjGNComp.CodVendedor) > 0 Then
            Set Ven = mobjGNComp.Empresa.RecuperaFCVendedor(mobjGNComp.CodVendedor)
            .TextMatrix(r, COL_CODVEN) = Ven.CodVendedor
            .TextMatrix(r, COL_VENDEDOR) = Ven.nombre
            pck.IdVendedor = GNComprobante.IdVendedor
        End If
        '------------
        VisualizarPlazo pck, r
        '.TextMatrix(r, COL_PLAZO) = pck.FechaVenci - pck.FechaEmision
        '.TextMatrix(r, COL_FECHAVENCI) = pck.FechaVenci
        .Row = r
        If .Rows > .FixedRows Then
            .col = .FixedCols
            'Busca la primera columna
            For i = .FixedCols To .Cols - 1
                If .ColData(i) >= 0 And (Not .ColHidden(i)) And .ColWidth(i) > 0 Then
                    .col = i
                    Exit For
                End If
            Next i
        End If
    End With
    PoneNumFila
    VisualizaTotal
salida:
    Set pck = Nothing
    Set Ven = Nothing
    grd.SetFocus
    Exit Sub
ErrTrap:
    Set pck = Nothing
    Set Ven = Nothing
    DispErr
    GoTo salida
End Sub


Public Sub VisualizaDesdeObjetoAuto(ByVal band As Boolean)
    Dim pck As PCKardexCHP, pc As PCProvCli, i As Long, s As String
    Dim r As Long, NombrePC As String, Ven As FCVendedor, nombreVen As String
    Dim tsf As TSFormaCobroPago
    With grd
        .Redraw = False
        Limpiar
        For i = 1 To mobjGNComp.CountPCKardex
            Set pck = mobjGNComp.PCKardexCHP(i)
            If Len(mstrCodProvCli) = 0 Then   'Agregado esta condicion para que cuando se recupera carge a la variable de la propiedad Docs.codProvCli
                mstrCodProvCli = pck.CodProvCli
            End If
            If pck.idAsignado = 0 Then
                'Si hay que mostrar el nombre de proveedor/cliente
                If Not .ColHidden(COL_PROVCLI) Then
                    'Recupera PCProvCli
                    Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(pck.CodProvCli)
                    If Not (pc Is Nothing) Then NombrePC = pc.nombre
                End If
                    'AUC 01/06/07
                    Set Ven = mobjGNComp.Empresa.RecuperaFCVendedor(pck.CodVendedor)
                    If Not (Ven Is Nothing) Then
                        nombreVen = Ven.nombre
                    Else
                        nombreVen = ""
                    End If
                
'                If mobjGNComp.GNTrans.codPantalla = "TSIEE" Then
                    Set tsf = mobjGNComp.Empresa.RecuperaTSFormaCobroPago(pck.codforma)
                    
                    If tsf.BandCobro = band And pck.Debe > 0 Then
                    s = .Rows & vbTab & _
                            pck.CodProvCli & vbTab & _
                            NombrePC & vbTab & _
                            pck.codforma & vbTab & _
                            pck.NumLetra & vbTab & _
                            NullSiZero(pck.Debe + pck.Haber) & vbTab & _
                            pck.FechaEmision & vbTab & _
                            pck.FechaVenci - pck.FechaEmision & vbTab & _
                            pck.FechaVenci & vbTab & _
                            pck.Observacion & vbTab & pck.CodVendedor & vbTab & nombreVen
                            If mobjGNComp.SoloVer Then
                                If Len(pck.codBanco) > 0 Then
                                    If Len(pck.CodTarjeta) > 0 Then
                                        s = s & vbTab & pck.codBanco & "/" & pck.CodTarjeta
                                    Else
                                        s = s & vbTab & pck.codBanco & "/" & pck.Numcheque
                                    End If
                                End If
                            End If
                            
                            .AddItem s
                            r = .Rows - 1
                            .RowData(r) = pck
                            '*** Oliver 29/01/2003 Agregado para que bloquee las columnas que no se modifiquen
                            BloquearPlazo pck, r
                    Else
                            If tsf.BandCobro = band And Not band Then
                                s = .Rows & vbTab & _
                                pck.CodProvCli & vbTab & _
                                NombrePC & vbTab & _
                                pck.codforma & vbTab & _
                                pck.NumLetra & vbTab & _
                                NullSiZero(pck.Debe + pck.Haber) & vbTab & _
                                pck.FechaEmision & vbTab & _
                                pck.FechaVenci - pck.FechaEmision & vbTab & _
                                pck.FechaVenci & vbTab & _
                                pck.Observacion & vbTab & pck.CodVendedor & vbTab & nombreVen
                                .AddItem s
                                r = .Rows - 1
                                .RowData(r) = pck
                                '*** Oliver 29/01/2003 Agregado para que bloquee las columnas que no se modifiquen
                                BloquearPlazo pck, r
                            Else
                                If tsf.BandCobro And pck.Haber > 0 And Not band Then
                                
                                s = .Rows & vbTab & _
                                pck.CodProvCli & vbTab & _
                                NombrePC & vbTab & _
                                pck.codforma & vbTab & _
                                pck.NumLetra & vbTab & _
                                NullSiZero(pck.Debe + pck.Haber) & vbTab & _
                                pck.FechaEmision & vbTab & _
                                pck.FechaVenci - pck.FechaEmision & vbTab & _
                                pck.FechaVenci & vbTab & _
                                pck.Observacion & vbTab & pck.CodVendedor & vbTab & nombreVen
                                .AddItem s
                                r = .Rows - 1
                                .RowData(r) = pck
'                                '*** Oliver 29/01/2003 Agregado para que bloquee las columnas que no se modifiquen
                                BloquearPlazo pck, r
                                End If
                            End If
                    End If
'                Else
'
'                s = .Rows & vbTab & _
'                        pck.CodProvCli & vbTab & _
'                        NombrePC & vbTab & _
'                        pck.codForma & vbTab & _
'                        pck.NumLetra & vbTab & _
'                        NullSiZero(pck.Debe + pck.Haber) & vbTab & _
'                        pck.FechaEmision & vbTab & _
'                        pck.FechaVenci - pck.FechaEmision & vbTab & _
'                        pck.FechaVenci & vbTab & _
'                        pck.Observacion & vbTab & pck.CodVendedor & vbTab & nombreVen
'                        .AddItem s
'                        r = .Rows - 1
'                        .RowData(r) = pck
'                        '*** Oliver 29/01/2003 Agregado para que bloquee las columnas que no se modifiquen
'                        BloquearPlazo pck, r
'
'                End If
            End If
        Next i
        PoneNumFila
        VisualizaTotal
        .Redraw = True
    End With
    Set Ven = Nothing
End Sub


Public Sub AgregaFilaTC(ByVal codforma As String, valor As Currency)
    Dim r As Long, r2 As Long, ix As Long, i As Long, v As Currency
    Dim pck As PCKardexCHP
    Dim tsf As TSFormaCobroPago
    Dim pc As PCProvCli, obser As String
    Dim GNCompAux As GNComprobante
    RaiseEvent PorAgregarFila(v)        'Para calcular el valor predeterminado

    On Error GoTo ErrTrap
    
    'Llama a agregar un objeto PCKardexCHP antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
    ix = mobjGNComp.AddPCKardex
    
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

        'Asigna la referencia al nuevo objeto a la fila nueva
        Set pck = mobjGNComp.PCKardexCHP(ix)
        .RowData(r) = pck
        
        
        Set tsf = mobjGNComp.Empresa.RecuperaTSFormaCobroPago(codforma)
        If tsf.DeudaMismoCliente Then
            Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(mobjGNComp.CodClienteRef)
        Else
            Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(tsf.CodProvCli)
        End If
        Set GNCompAux = mobjGNComp.Empresa.RecuperaGNComprobante(mobjGNComp.idTransFuente)

        
        'Proporciona el valor predeterminado        '*** MAKOTO 05/oct/00 Modificado
        If mbooPorCobrar Then pck.Debe = Abs(v)
         .TextMatrix(r, COL_VALOR) = valor
         pck.Debe = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
         
         If tsf.DeudaMismoCliente Then
            pck.CodProvCli = GNCompAux.CodClienteRef
        Else
            pck.CodProvCli = tsf.CodProvCli
        End If
        .TextMatrix(r, COL_CODPROVCLI) = pck.CodProvCli     '*** MAKOTO 14/oct/00
        
         VisualizaProvCli r, pck.CodProvCli                 '***
        
        .TextMatrix(r, COL_CODFORMA) = tsf.CodFormaTC
        
        pck.codforma = tsf.CodFormaTC
        
        pck.NumLetra = GNCompAux.CodTrans & " " & GNCompAux.numtrans
        
        obser = "Por pago con: " & tsf.codforma & " de " & mobjGNComp.CodTrans & "-" & mobjGNComp.numtrans & " Cliente: " & mobjGNComp.CodClienteRef & " - " & mobjGNComp.nombre
        pck.Observacion = IIf(Len(obser) > 80, Left(obser, 80), obser)
        
        
'        pck.Observacion = "Por pago con: " & tsf.codforma & " de " & GNCompAux.codtrans & "-" & GNCompAux.NumTrans & " Cliente: " & GNCompAux.CodClienteRef & " - " & GNCompAux.nombre
        mobjGNComp.Descripcion = pck.Observacion
        .TextMatrix(r, COL_OBSERVA) = pck.Observacion
        .TextMatrix(r, COL_NUMLETRA) = pck.NumLetra
        .TextMatrix(r, COL_FECHAEMI) = pck.FechaEmision
        'VisualizarPlazo pck, r
        '.TextMatrix(r, COL_PLAZO) = pck.FechaVenci - pck.FechaEmision
        .TextMatrix(r, COL_FECHAVENCI) = pck.FechaEmision
        pck.FechaVenci = pck.FechaEmision
        .TextMatrix(r, COL_PLAZO) = pck.FechaVenci - pck.FechaEmision
        
        .Row = r
        If .Rows > .FixedRows Then
            .col = .FixedCols
            'Busca la primera columna
            For i = .FixedCols To .Cols - 1
                If .ColData(i) >= 0 And (Not .ColHidden(i)) And .ColWidth(i) > 0 Then
                    .col = i
                    Exit For
                End If
            Next i
        End If
    End With
    
    PoneNumFila
    VisualizaTotal
salida:
    Set pck = Nothing
    Set tsf = Nothing
    Set pc = Nothing
    Set GNCompAux = Nothing
    
    grd.SetFocus
    Exit Sub
ErrTrap:
    Set pck = Nothing
    DispErr
    GoTo salida
End Sub

Public Sub AgregaFilaPagoInicial()
    Dim r As Long, r2 As Long, ix As Long, i As Long, ValorIni As Currency, saldo As Currency
    Dim pck As PCKardexCHP, Ven  As FCVendedor

'    RaiseEvent ValorPagoInicial(ValorIni)
'    RaiseEvent PorAgregarFilaconPagoInicial(Saldo)        'Para calcular el valor predeterminado

    On Error GoTo ErrTrap
    If Not EliminaTodaslasFilasDocs Then Exit Sub
    ValorIni = mobjGNComp.ValorEntrada
    If ValorIni = 0 Then Exit Sub
    
    'Llama a agregar un objeto PCKardexCHP antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
    ix = mobjGNComp.AddPCKardex
    
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

        'Asigna la referencia al nuevo objeto a la fila nueva
        Set pck = mobjGNComp.PCKardexCHP(ix)
        .RowData(r) = pck
        
        'Proporciona el valor predeterminado        '*** MAKOTO 05/oct/00 Modificado
'        If v > 0 Then
            If mbooPorCobrar Then pck.Debe = Abs(mobjGNComp.ValorEntrada)
            .TextMatrix(r, COL_VALOR) = mobjGNComp.ValorEntrada
            pck.Debe = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
'        ElseIf v < 0 Then
'            If Not mbooPorCobrar Then pck.Haber = Abs(v)
'            .TextMatrix(r, COL_VALOR) = Valor
'            pck.Haber = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
'        End If
        .TextMatrix(r, COL_CODPROVCLI) = pck.CodProvCli     '*** MAKOTO 14/oct/00
        VisualizaProvCli r, pck.CodProvCli                  '***
        
        '***Agregado. 17/Ago/2004. Angel
        '***Para que se inserte la fila pero con la forma de pago predeterminada en la configuracion IVFIN
        If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FornmaCobroCuotaInicial")) > 0 Then
            If mobjGNComp.FechaPrimerPago = pck.FechaEmision Then
                .TextMatrix(r, COL_CODFORMA) = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FornmaCobroCuotaInicial")
            Else
                .TextMatrix(r, COL_CODFORMA) = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FornmaCobroOtrasCuotas")
            End If
            pck.codforma = Trim$(.TextMatrix(r, COL_CODFORMA))
        Else
            .TextMatrix(r, COL_CODFORMA) = pck.codforma
        End If
        pck.NumLetra = "Cuota Inicial"
        .TextMatrix(r, COL_NUMLETRA) = pck.NumLetra
        
        .TextMatrix(r, COL_FECHAEMI) = pck.FechaEmision
        
       If mobjGNComp.FechaPrimerPago <> pck.FechaEmision Then
            .TextMatrix(r, COL_FECHAVENCI) = mobjGNComp.FechaPrimerPago
            pck.FechaVenci = mobjGNComp.FechaPrimerPago
        Else
            .TextMatrix(r, COL_FECHAVENCI) = pck.FechaEmision
            pck.FechaVenci = pck.FechaEmision
        End If
        .TextMatrix(r, COL_PLAZO) = pck.FechaVenci - pck.FechaEmision
        
        If mobjGNComp.GNTrans.IVPideVendedor And Len(mobjGNComp.CodVendedor) > 0 Then
            Set Ven = mobjGNComp.Empresa.RecuperaFCVendedor(mobjGNComp.CodVendedor)
            .TextMatrix(r, COL_CODVEN) = Ven.CodVendedor
            .TextMatrix(r, COL_VENDEDOR) = Ven.nombre
            pck.IdVendedor = GNComprobante.IdVendedor
        End If

        .Cell(flexcpBackColor, r, COL_FECHAEMI) = grd.BackColorFixed
        .Cell(flexcpBackColor, r, COL_FECHAEMI) = grd.BackColorFixed

        
        .Row = r
        If .Rows > .FixedRows Then
            .col = .FixedCols
            'Busca la primera columna
            For i = .FixedCols To .Cols - 1
                If .ColData(i) >= 0 And (Not .ColHidden(i)) And .ColWidth(i) > 0 Then
                    .col = i
                    Exit For
                End If
            Next i
        End If
        'If mobjGNComp.Empresa.GNOpcion.ObtenerValor("BloquearCuotas") = "1" Then
''        If mobjGNComp.GNTrans.IVBloquearCuotas Then
''            For i = COL_NUMLETRA To COL_VENDEDOR
''                .ColData(i) = -1
''            Next i
''         End If
'        HabilitarCtrlsGrupoCaja False
    End With


    PoneNumFila
    VisualizaTotal
    
    If mobjGNComp.GNTrans.ivBloquearCuotas Then
        For i = COL_NUMLETRA To COL_VENDEDOR
            grd.ColData(i) = -1
        Next i
        HabilitarCtrlsGrupoCajaFormaCobro False
     End If
    
salida:
    Set pck = Nothing
    grd.SetFocus
    Exit Sub
ErrTrap:
    Set pck = Nothing
    DispErr
    GoTo salida
End Sub

Public Sub AgregaFilasPorCredito(ByVal valor As Double)
    Dim r As Long, r2 As Long, ix As Long, i As Long, ValorIni As Currency, saldo As Currency
    Dim pck As PCKardexCHP, Ven As FCVendedor
    Dim NumPagos As Integer, ValorCuota As Currency, CuotaFinal As Currency, TotalCuota As Currency
    Dim numdecimales As Integer, valorCuotasGeneradas As Currency
    Dim intervalo As Integer, pc As PCProvCli
'    RaiseEvent PorAgregarFilaconPagoInicial(Saldo)        'Para calcular el valor predeterminado
    On Error GoTo ErrTrap
'    If valor <= 0 Then Exit Sub
    
    If mobjGNComp.GNTrans.IVGeneraPagosxCliente Then
        If mobjGNComp.GNTrans.IVTipoTrans = "I" Then
            Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(mobjGNComp.CodProveedorRef)
        Else
            Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(mobjGNComp.CodClienteRef)
        End If
        If Not pc Is Nothing Then
            intervalo = pc.intervalo
'            valor = pc.intervalo
'            NumPagos = pc.NumPagos
        End If
        Set pc = Nothing
    Else
        If valor <= 0 Then Exit Sub
    End If
    TotalCuota = 0
    valorCuotasGeneradas = valor
    'If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("NumDecimalesCuotas")) > 0 Then
    If Len(mobjGNComp.GNTrans.IVNumDecimalesCuotas) > 0 Then
        'numdecimales = mobjGNComp.Empresa.GNOpcion.ObtenerValor("NumDecimalesCuotas")
        numdecimales = mobjGNComp.GNTrans.IVNumDecimalesCuotas
    Else
        numdecimales = 2
    End If
    ValorCuota = Round(valor / mobjGNComp.NumeroPagos, numdecimales)
    For NumPagos = 1 To mobjGNComp.NumeroPagos
        valorCuotasGeneradas = valorCuotasGeneradas - ValorCuota
        'Llama a agregar un objeto PCKardexCHP antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
        ix = mobjGNComp.AddPCKardex
        
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
    
            'Asigna la referencia al nuevo objeto a la fila nueva
            Set pck = mobjGNComp.PCKardexCHP(ix)
            .RowData(r) = pck
            
            'Proporciona el valor predeterminado        '*** MAKOTO 05/oct/00 Modificado
            
            If valorCuotasGeneradas < 0 Then
                ValorCuota = Round(valor - TotalCuota, 2)
'                NumPagos = mobjGNComp.NumeroPagos
            End If
            
            If NumPagos = mobjGNComp.NumeroPagos Then
                ValorCuota = Round(valor - TotalCuota, 2)
            End If
            If mbooPorCobrar Then
                pck.Debe = Abs(ValorCuota)
            Else
                pck.Haber = Abs(ValorCuota)
            End If
            .TextMatrix(r, COL_VALOR) = ValorCuota
            TotalCuota = TotalCuota + ValorCuota
            If mbooPorCobrar Then
                pck.Debe = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
            Else
                pck.Haber = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
            End If
            .TextMatrix(r, COL_CODPROVCLI) = pck.CodProvCli     '*** MAKOTO 14/oct/00
            VisualizaProvCli r, pck.CodProvCli                  '***
            
            '***Agregado. 17/Ago/2004. Angel
            '***Para que se inserte la fila pero con la forma de pago predeterminada en la configuracion IVFIN
            If mbooPorCobrar Then
                If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FornmaCobroOtrasCuotas")) > 0 Then
                    .TextMatrix(r, COL_CODFORMA) = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FornmaCobroOtrasCuotas")
                    pck.codforma = Trim$(.TextMatrix(r, COL_CODFORMA))
                Else
                    .TextMatrix(r, COL_CODFORMA) = pck.codforma
                End If
            Else
                If Len(mobjGNComp.GNTrans.CodFormaPre) > 0 Then
                    If Not mobjGNComp.GNTrans.IVDescXFormaCP Then
                        .TextMatrix(r, COL_CODFORMA) = mobjGNComp.GNTrans.CodFormaPre
                    Else
                        If Len(mobjGNComp.CodFormnaCP) > 0 Then
                            .TextMatrix(r, COL_CODFORMA) = mobjGNComp.CodFormnaCP
                            BloqueaColumnaCodForma mobjGNComp.CodFormnaCP
                        End If
                    End If
                    pck.codforma = Trim$(.TextMatrix(r, COL_CODFORMA))
                Else
                    If mbooPorCobrar Then
                        If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroAnticipo")) > 0 Then
                            If mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroAnticipo") <> pck.codforma Then
                                pck.codforma = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroAnticipo")
                            End If
                        End If
                    Else
                        If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoAnticipo")) > 0 Then
                            If mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoAnticipo") <> pck.codforma Then
                                pck.codforma = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoAnticipo")
                            End If
                        End If
                    End If
                    .TextMatrix(r, COL_CODFORMA) = pck.codforma
                End If
            
            End If
            'mobjGNComp.codtrans & "-" & mobjGNComp.GNTrans.NumTransSiguiente &
            If mobjGNComp.GNTrans.IVPideNumDoc Then
                If Len(mobjGNComp.numDocRef) > 0 Then
                    pck.NumLetra = Mid$(mobjGNComp.numDocRef & " - " & NumPagos & "/" & mobjGNComp.NumeroPagos, 1, 20)
                Else
                    If mobjGNComp.GNTrans.CodPantalla = "TSIER" Then
                        pck.NumLetra = Mid$("Refinan - " & NumPagos & "/" & mobjGNComp.NumeroPagos, 1, 20)
                    Else
                        pck.NumLetra = Mid$(NumPagos & "/" & mobjGNComp.NumeroPagos, 1, 20)
                    End If
                End If
            Else
                pck.NumLetra = "Cuota " & NumPagos & "/" & mobjGNComp.NumeroPagos
            End If
            .TextMatrix(r, COL_NUMLETRA) = pck.NumLetra
            .TextMatrix(r, COL_FECHAEMI) = pck.FechaEmision

            
            If mbooPorCobrar Then
                If mobjGNComp.DiaPago <> DatePart("d", pck.FechaEmision) Then
                    .TextMatrix(r, COL_FECHAVENCI) = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision))
                    pck.FechaVenci = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision))
                Else
                    If mobjGNComp.GNTrans.CodPantalla = "TSIER" Then
                        .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, mobjGNComp.FechaPrimerPago)
                        pck.FechaVenci = DateAdd("m", NumPagos - 1 + mobjGNComp.MesesGracia, mobjGNComp.FechaPrimerPago)
                    Else
                        .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
                        pck.FechaVenci = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
                    End If
                End If
            Else
'                If mobjGNComp.DiaPago <> DatePart("d", pck.FechaEmision) Then
'                    .TextMatrix(r, COL_FECHAVENCI) = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision))
'                    pck.FechaVenci = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision))
'                Else

                If mobjGNComp.GNTrans.IVGeneraPagosxCliente Then
                    .TextMatrix(r, COL_FECHAVENCI) = DateAdd("d", NumPagos * intervalo, pck.FechaEmision)
                    pck.FechaVenci = DateAdd("d", NumPagos * intervalo, pck.FechaEmision)
                Else
                    .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
                    pck.FechaVenci = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
                End If
'                End If
            
            End If
            
            .TextMatrix(r, COL_PLAZO) = pck.FechaVenci - pck.FechaEmision
            
        If mobjGNComp.GNTrans.IVPideVendedor And Len(mobjGNComp.CodVendedor) > 0 Then
            Set Ven = mobjGNComp.Empresa.RecuperaFCVendedor(mobjGNComp.CodVendedor)
            .TextMatrix(r, COL_CODVEN) = Ven.CodVendedor
            .TextMatrix(r, COL_VENDEDOR) = Ven.nombre
            pck.IdVendedor = GNComprobante.IdVendedor
        End If
            
            
            If valorCuotasGeneradas < 0 Then
                NumPagos = mobjGNComp.NumeroPagos
            End If
            
            
        .Cell(flexcpBackColor, r, COL_FECHAEMI) = grd.BackColorFixed
        .Cell(flexcpBackColor, r, COL_FECHAEMI) = grd.BackColorFixed

            .Row = r
            If .Rows > .FixedRows Then
                .col = .FixedCols
                'Busca la primera columna
                For i = .FixedCols To .Cols - 1
                    If .ColData(i) >= 0 And (Not .ColHidden(i)) And .ColWidth(i) > 0 Then
                        .col = i
                        Exit For
                    End If
                Next i
            End If
''''         HabilitarCtrlsGrupoCaja False
            
        End With
    
        PoneNumFila
        VisualizaTotal
    Next NumPagos
    'If mobjGNComp.Empresa.GNOpcion.ObtenerValor("BloquearCuotas") = "1" Then
    If mobjGNComp.GNTrans.ivBloquearCuotas Then
        For i = COL_NUMLETRA To COL_VENDEDOR
            grd.ColData(i) = -1
        Next i
        HabilitarCtrlsGrupoCajaFormaCobro False
     End If
    
salida:
    Set pck = Nothing
    grd.SetFocus
    Exit Sub
ErrTrap:
    Set pck = Nothing
    DispErr
    GoTo salida
End Sub

Public Sub EliminaFilaTodoDocs()
    Dim msg As String, r As Long, i As Long
    Dim Cancel As Boolean
    
    On Error GoTo ErrTrap
        r = mobjGNComp.CountPCKardex
        'For i = r To 1 Step -1
        i = 1
        Do While mobjGNComp.CountPCKardex <> 0
'            grd.RemoveItem i 'borra la linea de subtotal
'            If grd.Rows > 1 Then
                    mobjGNComp.RemovePCKardex i, grd.RowData(i)
                    grd.RemoveItem i
'            End If
        Loop
        'Next i
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub AgregaFilaDescXFormaCP(codforma)
    Dim r As Long, r2 As Long, ix As Long, i As Long, v As Currency
    Dim pck As PCKardexCHP
    Dim Ven As FCVendedor
    RaiseEvent PorAgregarFila(v)        'Para calcular el valor predeterminado

    On Error GoTo ErrTrap
    
    'Llama a agregar un objeto PCKardexCHP antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
    ix = mobjGNComp.AddPCKardex
    
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

        'Asigna la referencia al nuevo objeto a la fila nueva
        Set pck = mobjGNComp.PCKardexCHP(ix)
        .RowData(r) = pck
        
        'Proporciona el valor predeterminado        '*** MAKOTO 05/oct/00 Modificado
        If v > 0 Then
            If mbooPorCobrar Then pck.Debe = Abs(v)
            .TextMatrix(r, COL_VALOR) = pck.Debe
            pck.Debe = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
        ElseIf v < 0 Then
            If Not mbooPorCobrar Then pck.Haber = Abs(v)
            .TextMatrix(r, COL_VALOR) = pck.Haber
            pck.Haber = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
        End If
        .TextMatrix(r, COL_CODPROVCLI) = pck.CodProvCli     '*** MAKOTO 14/oct/00
        VisualizaProvCli r, pck.CodProvCli                  '***
        
        
        '.TextMatrix(r, COL_NOMBRE) = mobjGNComp.
        '***Agregado. 17/Ago/2004. Angel
        '***Para que se inserte la fila pero con la forma de pago predeterminada
        .TextMatrix(r, COL_CODFORMA) = codforma
        pck.codforma = Trim$(.TextMatrix(r, COL_CODFORMA))
        
        .TextMatrix(r, COL_NUMLETRA) = pck.NumLetra
        .TextMatrix(r, COL_FECHAEMI) = pck.FechaEmision
        
        'AUC 19/07/06
        If mobjGNComp.GNTrans.IVPideVendedor And Len(mobjGNComp.CodVendedor) > 0 Then
            Set Ven = mobjGNComp.Empresa.RecuperaFCVendedor(mobjGNComp.CodVendedor)
            .TextMatrix(r, COL_CODVEN) = Ven.CodVendedor
            .TextMatrix(r, COL_VENDEDOR) = Ven.nombre
            pck.IdVendedor = GNComprobante.IdVendedor
        End If
        '------------
        VisualizarPlazo pck, r
        '.TextMatrix(r, COL_PLAZO) = pck.FechaVenci - pck.FechaEmision
        '.TextMatrix(r, COL_FECHAVENCI) = pck.FechaVenci
        
        .Row = r
        If .Rows > .FixedRows Then
            .col = .FixedCols
            'Busca la primera columna
            For i = .FixedCols To .Cols - 1
                If .ColData(i) >= 0 And (Not .ColHidden(i)) And .ColWidth(i) > 0 Then
                    .col = i
                    Exit For
                End If
            Next i
        End If
    End With
    
    PoneNumFila
    VisualizaTotal
salida:
    Set pck = Nothing
    Set Ven = Nothing
    grd.SetFocus
    Exit Sub
ErrTrap:
    Set pck = Nothing
    Set Ven = Nothing
    DispErr
    GoTo salida
End Sub

Private Sub BloqueaColumnaCodForma(codforma As String)
    Dim tsf As TSFormaCobroPago
    If mobjGNComp.GNTrans.IVDescXFormaCP Then
        Set tsf = mobjGNComp.Empresa.RecuperaTSFormaCobroPago(codforma)
        If tsf.PorDesc <> 0 Then
            grd.ColData(COL_CODFORMA) = -1
                
        Else
            grd.ColData(COL_CODFORMA) = 1
        End If
        Set tsf = Nothing
    End If
End Sub

Public Sub CambiaFecha()
    Dim i As Long
    With grd
        For i = .FixedRows To .Rows - 1
            If Not .IsSubtotal(i) Then
                grd.TextMatrix(i, COL_FECHAEMI) = mobjGNComp.FechaTrans
                grd.TextMatrix(i, COL_FECHAVENCI) = DateAdd("d", grd.TextMatrix(i, COL_PLAZO), mobjGNComp.FechaTrans)
            End If
        Next i
    End With
End Sub

Public Sub AgregaFilaSegundoPagoInicial()
    Dim r As Long, r2 As Long, ix As Long, i As Long, ValorIni As Currency, saldo As Currency
    Dim pck As PCKardexCHP, Ven  As FCVendedor

'    RaiseEvent ValorPagoInicial(ValorIni)
'    RaiseEvent PorAgregarFilaconPagoInicial(Saldo)        'Para calcular el valor predeterminado

    On Error GoTo ErrTrap
'    EliminaFilaTodoDocs
    ValorIni = mobjGNComp.ValorSegundaEntrada
    If ValorIni = 0 Then Exit Sub
    
    'Llama a agregar un objeto PCKardexCHP antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
    ix = mobjGNComp.AddPCKardex
    
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

        'Asigna la referencia al nuevo objeto a la fila nueva
        Set pck = mobjGNComp.PCKardexCHP(ix)
        .RowData(r) = pck
        
        'Proporciona el valor predeterminado        '*** MAKOTO 05/oct/00 Modificado
        If mbooPorCobrar Then pck.Debe = Abs(mobjGNComp.ValorSegundaEntrada)
        .TextMatrix(r, COL_VALOR) = mobjGNComp.ValorSegundaEntrada
        pck.Debe = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
        .TextMatrix(r, COL_CODPROVCLI) = pck.CodProvCli     '*** MAKOTO 14/oct/00
        VisualizaProvCli r, pck.CodProvCli                  '***
        
        '***Agregado. 17/Ago/2004. Angel
        '***Para que se inserte la fila pero con la forma de pago predeterminada en la configuracion IVFIN
            If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FornmaCobroOtrasCuotas")) > 0 Then
                .TextMatrix(r, COL_CODFORMA) = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FornmaCobroOtrasCuotas")
                pck.codforma = Trim$(.TextMatrix(r, COL_CODFORMA))
            Else
                .TextMatrix(r, COL_CODFORMA) = pck.codforma
            End If

        pck.NumLetra = "Segunda Cuota Ini"
        .TextMatrix(r, COL_NUMLETRA) = pck.NumLetra
        
        .TextMatrix(r, COL_FECHAEMI) = pck.FechaEmision
        
       If mobjGNComp.FechaSegundoPago <> pck.FechaEmision Then
            .TextMatrix(r, COL_FECHAVENCI) = mobjGNComp.FechaSegundoPago
            pck.FechaVenci = mobjGNComp.FechaSegundoPago
        Else
            .TextMatrix(r, COL_FECHAVENCI) = pck.FechaEmision
            pck.FechaVenci = pck.FechaEmision
        End If
        .TextMatrix(r, COL_PLAZO) = pck.FechaVenci - pck.FechaEmision
        
        If mobjGNComp.GNTrans.IVPideVendedor And Len(mobjGNComp.CodVendedor) > 0 Then
            Set Ven = mobjGNComp.Empresa.RecuperaFCVendedor(mobjGNComp.CodVendedor)
            .TextMatrix(r, COL_CODVEN) = Ven.CodVendedor
            .TextMatrix(r, COL_VENDEDOR) = Ven.nombre
            pck.IdVendedor = GNComprobante.IdVendedor
        End If

        .Cell(flexcpBackColor, r, COL_FECHAEMI) = grd.BackColorFixed
        .Cell(flexcpBackColor, r, COL_FECHAEMI) = grd.BackColorFixed

        
        .Row = r
        If .Rows > .FixedRows Then
            .col = .FixedCols
            'Busca la primera columna
            For i = .FixedCols To .Cols - 1
                If .ColData(i) >= 0 And (Not .ColHidden(i)) And .ColWidth(i) > 0 Then
                    .col = i
                    Exit For
                End If
            Next i
        End If
    
        'If mobjGNComp.Empresa.GNOpcion.ObtenerValor("BloquearCuotas") = "1" Then
'''        If mobjGNComp.GNTrans.IVBloquearCuotas Then
'''            For i = 1 To COL_VENDEDOR
'''                .ColData(i) = -1
'''            Next i
'''         End If
'''        HabilitarCtrlsGrupoCajaFormaCobro False
    
    End With



    PoneNumFila
    VisualizaTotal
    If mobjGNComp.GNTrans.ivBloquearCuotas Then
        For i = COL_NUMLETRA To COL_VENDEDOR
            grd.ColData(i) = -1
        Next i
        HabilitarCtrlsGrupoCajaFormaCobro False
     End If
    
salida:
    Set pck = Nothing
    grd.SetFocus
    Exit Sub
ErrTrap:
    Set pck = Nothing
    DispErr
    GoTo salida
End Sub

Public Function EliminaTodaslasFilasDocs() As Boolean
    Dim msg As String, r As Long, i As Long
    Dim Cancel As Boolean
    EliminaTodaslasFilasDocs = True
    On Error GoTo ErrTrap
        r = mobjGNComp.CountPCKardex
        'For i = r To 1 Step -1
        i = 1
        Do While mobjGNComp.CountPCKardex <> 0
'            grd.RemoveItem i 'borra la linea de subtotal
'            If grd.Rows > 1 Then
                mobjGNComp.RemovePCKardex i, grd.RowData(i)
                grd.RemoveItem i
'            End If
        Loop
        'elimina del del subtotal
        If grd.Rows <> 1 Then
            grd.RemoveItem i
        End If
        'Next i
    Exit Function
ErrTrap:
    EliminaTodaslasFilasDocs = False
    DispErr
    Exit Function
End Function

Public Sub HabilitarCtrlsGrupoCajaFormaCobro(ByVal BandHabilita As Boolean)
Dim col As Integer, fil As Integer
    For col = 1 To 9
        If Not BandHabilita Then
            If col <> COL_CODFORMA Then 'And mobjGNComp.GNTrans.IVDescXPCGrupo Then
                grd.ColData(col) = -1
                For fil = 1 To grd.Rows - 2
                            grd.Cell(flexcpBackColor, fil, col, fil, col) = &H80000018
                Next fil
            End If
        End If
    Next col
    grd.Enabled = BandHabilita
End Sub

Public Sub EliminaFilaTodoDocsGenerados()
    Dim msg As String, r As Long, i As Long, fila As Integer
    Dim Cancel As Boolean
    
    On Error GoTo ErrTrap
        fila = grd.Rows - 2
        r = mobjGNComp.CountPCKardex
        For i = r To 1 Step -1
            If mobjGNComp.PCKardexCHP(i).Debe > 0 Then
                If mobjGNComp.PCKardexCHP(i).Debe > 0 Then
                    mobjGNComp.RemovePCKardex i, grd.RowData(fila)
                    grd.RemoveItem fila
                    fila = fila - 1
                End If
            End If
        Next i
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Public Sub AgregaFilasPorCreditOtroProvCliente(ByVal CodProvCli As String, ByVal valor As Double)
    Dim r As Long, r2 As Long, ix As Long, i As Long, ValorIni As Currency, saldo As Currency
    Dim pck As PCKardexCHP, Ven As FCVendedor
    Dim NumPagos As Integer, ValorCuota As Currency, CuotaFinal As Currency, TotalCuota As Currency
    Dim numdecimales As Integer, valorCuotasGeneradas As Currency
'    RaiseEvent PorAgregarFilaconPagoInicial(Saldo)        'Para calcular el valor predeterminado
    On Error GoTo ErrTrap
    If valor <= 0 Then Exit Sub
    TotalCuota = 0
    valorCuotasGeneradas = valor
    'If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("NumDecimalesCuotas")) > 0 Then
    If Len(mobjGNComp.GNTrans.IVNumDecimalesCuotas) > 0 Then
        'numdecimales = mobjGNComp.Empresa.GNOpcion.ObtenerValor("NumDecimalesCuotas")
        numdecimales = mobjGNComp.GNTrans.IVNumDecimalesCuotas
    Else
        numdecimales = 2
    End If
    ValorCuota = Round(valor / mobjGNComp.NumeroPagos, numdecimales)
    For NumPagos = 1 To mobjGNComp.NumeroPagos
        valorCuotasGeneradas = valorCuotasGeneradas - ValorCuota
        'Llama a agregar un objeto PCKardexCHP antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
        ix = mobjGNComp.AddPCKardex
        
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
    
            'Asigna la referencia al nuevo objeto a la fila nueva
            Set pck = mobjGNComp.PCKardexCHP(ix)
            .RowData(r) = pck
            
            'Proporciona el valor predeterminado        '*** MAKOTO 05/oct/00 Modificado
            
            If valorCuotasGeneradas < 0 Then
                ValorCuota = Round(valor - TotalCuota, 2)
'                NumPagos = mobjGNComp.NumeroPagos
            End If
            
            pck.CodProvCli = CodProvCli
            If NumPagos = mobjGNComp.NumeroPagos Then
                ValorCuota = Round(valor - TotalCuota, 2)
            End If
            If mbooPorCobrar Then
                pck.Debe = Abs(ValorCuota)
            Else
                pck.Haber = Abs(ValorCuota)
            End If
            .TextMatrix(r, COL_VALOR) = ValorCuota
            TotalCuota = TotalCuota + ValorCuota
            If mbooPorCobrar Then
                pck.Debe = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
            Else
                pck.Haber = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
            End If
            .TextMatrix(r, COL_CODPROVCLI) = pck.CodProvCli     '*** MAKOTO 14/oct/00
            VisualizaProvCli r, pck.CodProvCli                  '***
            
            '***Agregado. 17/Ago/2004. Angel
            '***Para que se inserte la fila pero con la forma de pago predeterminada en la configuracion IVFIN
            If mbooPorCobrar Then
                If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FornmaCobroOtrasCuotas")) > 0 Then
                    .TextMatrix(r, COL_CODFORMA) = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FornmaCobroOtrasCuotas")
                    pck.codforma = Trim$(.TextMatrix(r, COL_CODFORMA))
                Else
                    .TextMatrix(r, COL_CODFORMA) = pck.codforma
                End If
            Else
                If Len(mobjGNComp.GNTrans.CodFormaPre) > 0 Then
                    If Not mobjGNComp.GNTrans.IVDescXFormaCP Then
                        .TextMatrix(r, COL_CODFORMA) = mobjGNComp.GNTrans.CodFormaPre
                    Else
                        If Len(mobjGNComp.CodFormnaCP) > 0 Then
                            .TextMatrix(r, COL_CODFORMA) = mobjGNComp.CodFormnaCP
                            BloqueaColumnaCodForma mobjGNComp.CodFormnaCP
                        End If
                    End If
                    pck.codforma = Trim$(.TextMatrix(r, COL_CODFORMA))
                Else
                    If mbooPorCobrar Then
                        If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroAnticipo")) > 0 Then
                            If mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroAnticipo") <> pck.codforma Then
                                pck.codforma = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroAnticipo")
                            End If
                        End If
                    Else
                        If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoAnticipo")) > 0 Then
                            If mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoAnticipo") <> pck.codforma Then
                                pck.codforma = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoAnticipo")
                            End If
                        End If
                    End If
                    .TextMatrix(r, COL_CODFORMA) = pck.codforma
                End If
            
            End If
            'mobjGNComp.codtrans & "-" & mobjGNComp.GNTrans.NumTransSiguiente &
            If mobjGNComp.GNTrans.IVPideNumDoc Then
                If Len(mobjGNComp.numDocRef) > 0 Then
                    pck.NumLetra = Mid$(mobjGNComp.numDocRef & " - " & NumPagos & "/" & mobjGNComp.NumeroPagos, 1, 20)
                Else
                    If mobjGNComp.GNTrans.CodPantalla = "TSIER" Then
                        pck.NumLetra = Mid$("Refinan - " & NumPagos & "/" & mobjGNComp.NumeroPagos, 1, 20)
                    Else
                        pck.NumLetra = Mid$(NumPagos & "/" & mobjGNComp.NumeroPagos, 1, 20)
                    End If
                End If
            Else
                pck.NumLetra = "Cuota " & NumPagos & "/" & mobjGNComp.NumeroPagos
            End If
            .TextMatrix(r, COL_NUMLETRA) = pck.NumLetra
            .TextMatrix(r, COL_FECHAEMI) = pck.FechaEmision

            
            If mbooPorCobrar Then
                If mobjGNComp.DiaPago <> DatePart("d", pck.FechaEmision) Then
                    .TextMatrix(r, COL_FECHAVENCI) = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision))
                    pck.FechaVenci = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision))
                Else
                    If mobjGNComp.GNTrans.CodPantalla = "TSIER" Then
                        .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, mobjGNComp.FechaPrimerPago)
                        pck.FechaVenci = DateAdd("m", NumPagos - 1 + mobjGNComp.MesesGracia, mobjGNComp.FechaPrimerPago)
                    Else
                        .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
                        pck.FechaVenci = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
                    End If
                End If
            Else
                .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
                pck.FechaVenci = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
            
            End If
            
            .TextMatrix(r, COL_PLAZO) = pck.FechaVenci - pck.FechaEmision
            
        If mobjGNComp.GNTrans.IVPideVendedor And Len(mobjGNComp.CodVendedor) > 0 Then
            Set Ven = mobjGNComp.Empresa.RecuperaFCVendedor(mobjGNComp.CodVendedor)
            .TextMatrix(r, COL_CODVEN) = Ven.CodVendedor
            .TextMatrix(r, COL_VENDEDOR) = Ven.nombre
            pck.IdVendedor = GNComprobante.IdVendedor
        End If
            
            
            If valorCuotasGeneradas < 0 Then
                NumPagos = mobjGNComp.NumeroPagos
            End If
            
            
        .Cell(flexcpBackColor, r, COL_FECHAEMI) = grd.BackColorFixed
        .Cell(flexcpBackColor, r, COL_FECHAEMI) = grd.BackColorFixed

            .Row = r
            If .Rows > .FixedRows Then
                .col = .FixedCols
                'Busca la primera columna
                For i = .FixedCols To .Cols - 1
                    If .ColData(i) >= 0 And (Not .ColHidden(i)) And .ColWidth(i) > 0 Then
                        .col = i
                        Exit For
                    End If
                Next i
            End If
            
        End With
    
        PoneNumFila
        VisualizaTotal
    Next NumPagos
    'If mobjGNComp.Empresa.GNOpcion.ObtenerValor("BloquearCuotas") = "1" Then
    If mobjGNComp.GNTrans.ivBloquearCuotas Then
        For i = COL_NUMLETRA To COL_VENDEDOR
            grd.ColData(i) = -1
        Next i
        HabilitarCtrlsGrupoCajaFormaCobro False
     End If
    
salida:
    Set pck = Nothing
    grd.SetFocus
    Exit Sub
ErrTrap:
    Set pck = Nothing
    DispErr
    GoTo salida
End Sub

Public Sub OcultaBotonDatos(ByVal BandVisible As Boolean)
    cmdBotonDatos.Visible = BandVisible
End Sub


Private Function BuscaNumeroPrecio(fila As Long) As Integer
    Dim iv As IVinventario, i As Integer, j As Long, ivpadre As IVinventario
    
    Set iv = mobjGNComp.Empresa.RecuperaIVInventario(mobjGNComp.IVKardex(fila).CodInventario)
    If Not (iv Is Nothing) Then
        If mobjGNComp.IVKardex(fila).IdPadre = 0 Then
            For i = 1 To 6
                If mobjGNComp.IVKardex(fila).Precio = iv.Precio(i) Then
                    BuscaNumeroPrecio = i
                    Exit For
                End If
            Next i
        Else
            'ubica padrea
            For j = 1 To mobjGNComp.CountIVKardex
                If mobjGNComp.IVKardex(j).idinventario = mobjGNComp.IVKardex(fila).IdPadre Then
                    Set ivpadre = mobjGNComp.Empresa.RecuperaIVInventario(mobjGNComp.IVKardex(j).CodInventario)
                    If Not (iv Is Nothing) Then
                        For i = 1 To 6
                            If mobjGNComp.IVKardex(j).Precio = ivpadre.Precio(i) Then
                                BuscaNumeroPrecio = i
                                Exit For
                            End If
                        Next i
                    End If
                    'BuscaNumeroPrecio = mProps.objGNComprobante.IVKardex(j).NumeroPrecio
                    Exit For
                End If
            Next j
        End If
    End If
    Set iv = Nothing
End Function

Public Sub CargarDatosAdicionales()
    cmdBotonDatos_Click
End Sub

Private Sub grd_AfterUserResize(ByVal Row As Long, ByVal col As Long)
    With grd
         SaveSetting APPNAME, "SiiConfigCols", "config_col_DocsT_" & mobjGNComp.GNTrans.CodTrans & "_" & col, .ColWidth(col)
        ANCHO_COLS(col) = .ColWidth(col)
    End With
End Sub

Private Sub GetColsWidth()
    Dim i As Integer
    With grd
            For i = 0 To .Cols - 1
                   ANCHO_COLS(i) = GetSetting(APPNAME, "SiiConfigCols", "config_col_DocsT_" & mobjGNComp.GNTrans.CodTrans & "_" & i, 1200)

            Next i
    End With
End Sub

Public Sub CambiaVendedor()
    Dim i As Long
    With grd
        For i = .FixedRows To .Rows - 1
            If Not .IsSubtotal(i) Then
                grd.TextMatrix(i, COL_CODVEN) = mobjGNComp.CodVendedor
            End If
        Next i
    End With
End Sub
Public Sub AgregaFilaRol()
    Dim r As Long, r2 As Long, ix As Long, i As Long, v As Currency
    Dim rd As RolDetalle
    Dim pck As PCKardexCHP, k As Long
    Dim ele As Elementos
    Dim tsf As TSFormaCobroPago
    cmdBotonDatos.Visible = False
    'RaiseEvent PorAgregarFilaRol

    On Error GoTo ErrTrap
    grd.Rows = 1
    BandTarjeta = False
    For k = 1 To mobjGNComp.CountRolDetalle

        If mobjGNComp.RolDetalle(k).BandAfectaSaldoEmp And (mobjGNComp.RolDetalle(k).IdTipoRol = mobjGNComp.IdTipoRol _
                Or mobjGNComp.RolDetalle(k).IdTipoRol1 = mobjGNComp.IdTipoRol Or mobjGNComp.RolDetalle(k).IdTipoRol2 = mobjGNComp.IdTipoRol _
                Or mobjGNComp.RolDetalle(k).IdTipoRol3 = mobjGNComp.IdTipoRol) Then
                
            Set ele = mobjGNComp.Empresa.RecuperarElemento(mobjGNComp.RolDetalle(k).Codelemento)
            ix = mobjGNComp.AddPCKardex
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
                numTotalFila = r + OtrasFilas
                'Asigna la referencia al nuevo objeto a la fila nueva
                Set pck = mobjGNComp.PCKardexCHP(ix)
                .RowData(r) = pck
               
                cmdBotonDatos.Visible = False
                'Proporciona el valor predeterminado        '*** MAKOTO 05/oct/00 Modificado
                
                    If Not mbooPorCobrar Then pck.Haber = mobjGNComp.RolDetalle(k).valor 'Abs(v)
                    .TextMatrix(r, COL_VALOR) = pck.Haber
                    pck.Haber = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
                    pck.idElemento = mobjGNComp.RolDetalle(k).idElemento
                'End If
                .TextMatrix(r, COL_CODPROVCLI) = mobjGNComp.RolDetalle(k).CodEmpleado  'pck.CodProvCli     '*** MAKOTO 14/oct/00
                pck.CodProvCli = mobjGNComp.RolDetalle(k).CodEmpleado
                VisualizaProvCli r, mobjGNComp.RolDetalle(k).CodEmpleado
                        If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoEmpleado")) > 0 Then
                            If mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoEmpleado") <> pck.codforma Then
                                pck.codforma = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoEmpleado")
                            End If
                        End If
'                    End If
                    .TextMatrix(r, COL_CODFORMA) = pck.codforma
                'End If
                If mobjGNComp.GNTrans.IVDatosAdicionales Or mobjGNComp.GNTrans.TSDatosAdicionales Or mobjGNComp.GNTrans.TSDatosAdicionalesCHR Then
                    Set tsf = mobjGNComp.Empresa.RecuperaTSFormaCobroPago(grd.TextMatrix(r, COL_CODFORMA))
                    If Not tsf Is Nothing Then
                        BandTarjeta = InStr(1, UCase(tsf.NombreForma), "TARJETA") > 0
                        If tsf.DatosAdicionales Then
                            ' -- Redimensionar y posicionar el boton
                            cmdBotonDatos.Move (.Left + .CellLeft), _
                                       (.Top - 10 + (.RowHeight(0) * (.Row - .TopRow + 1))), _
                                       (.CellWidth), _
                                       (.CellHeight - 10)
                                           
                        ' -- Hacer visible y pasarle el foco
                            cmdBotonDatos.Visible = True
                            cmdBotonDatos.Enabled = True
                            cmdBotonDatos.SetFocus
                        Else
    '                        grd.Cell(flexcpData, r, COL_DATOSADI) = -1
                            grd.Cell(flexcpBackColor, r, COL_DATOSADI, r, COL_DATOSADI) = &H80000018
                            grd.TextMatrix(r, COL_DATOSADI) = "NO"
                        End If
                    End If
                    Set tsf = Nothing
                End If
            
                .TextMatrix(r, COL_NUMLETRA) = ele.Codelemento
                pck.NumLetra = ele.Codelemento
                .TextMatrix(r, COL_FECHAEMI) = pck.FechaEmision
                '------------
                VisualizarPlazo pck, r
                If Not GNComprobante.GNTrans.IVDesbloquearFechas Then
                    .Cell(flexcpBackColor, r, COL_FECHAEMI) = grd.BackColorFixed
                    .Cell(flexcpBackColor, r, COL_FECHAEMI) = grd.BackColorFixed
                End If
                .Row = r
                If .Rows > .FixedRows Then
                    .col = .FixedCols
                    'Busca la primera columna
                    For i = .FixedCols To .Cols - 1
                        If .ColData(i) >= 0 And (Not .ColHidden(i)) And .ColWidth(i) > 0 Then
                            .col = i
                            Exit For
                        End If
                    Next i
                End If
            End With
        End If
    Next k
    PoneNumFila
    VisualizaTotal
    
    
salida:
    Set pck = Nothing
    Set ele = Nothing
    grd.SetFocus
    Exit Sub
ErrTrap:
    Set pck = Nothing
    Set ele = Nothing
    DispErr
    GoTo salida
End Sub

Public Function EliminaFilaDocsRol() As Boolean
   Dim msg As String, r As Long, i As Long, fila As Integer
    Dim Cancel As Boolean
    On Error GoTo ErrTrap
        If grd.Rows = 1 Then: Exit Function
        fila = grd.Rows - 1
        r = mobjGNComp.CountPCKardex
       For i = r To 1 Step -1
            mobjGNComp.RemovePCKardex i ', grd.RowData(fila)
                   'grd.RemoveItem fila
                    'fila = fila - 1
        Next i
        grd.Rows = 1
        EliminaFilaDocsRol = True
    Exit Function
ErrTrap:
    EliminaFilaDocsRol = False
    DispErr
    Exit Function
End Function



Public Sub AgregaFilasxPCGrupo(valor As Currency, NumeroPagos As Integer, intervalo As Integer, BandDias As Boolean)
    Dim r As Long, r2 As Long, ix As Long, i As Long, ValorIni As Currency, saldo As Currency
    Dim pck As PCKardexCHP, Ven As FCVendedor, NumPagos  As Integer
    Dim ValorCuota As Currency, CuotaFinal As Currency, TotalCuota As Currency
    Dim numdecimales As Integer, valorCuotasGeneradas As Currency
'    RaiseEvent PorAgregarFilaconPagoInicial(Saldo)        'Para calcular el valor predeterminado
    On Error GoTo ErrTrap
    If valor <= 0 Then Exit Sub
    
    If mobjGNComp.CountPCKardex > 0 Then
        For i = 1 To mobjGNComp.CountPCKardex
            If mobjGNComp.PCKardexCHP(i).idAsignado = 0 Then
                mobjGNComp.BorrarPCKardex
           End If
        Next i
    End If
  
    
    
    TotalCuota = 0
    valorCuotasGeneradas = valor



    If Len(mobjGNComp.GNTrans.IVNumDecimalesCuotas) > 0 Then
        numdecimales = mobjGNComp.GNTrans.IVNumDecimalesCuotas
    Else
        numdecimales = 2
    End If
    If NumeroPagos = 0 Then NumeroPagos = 1
    ValorCuota = Round(valor / NumeroPagos, numdecimales)
    For NumPagos = 1 To NumeroPagos
        valorCuotasGeneradas = valorCuotasGeneradas - ValorCuota
        'Llama a agregar un objeto PCKardexCHP antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
        ix = mobjGNComp.AddPCKardex
        
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
    
            'Asigna la referencia al nuevo objeto a la fila nueva
            Set pck = mobjGNComp.PCKardexCHP(ix)
            .RowData(r) = pck
            
            'Proporciona el valor predeterminado        '*** MAKOTO 05/oct/00 Modificado
            
            If NumPagos = NumeroPagos Then
            'If valorCuotasGeneradas < 0 Then
                ValorCuota = Round(valor - TotalCuota, 2)
'                NumPagos = mobjGNComp.NumeroPagos
            End If
            
            If NumPagos = mobjGNComp.NumeroPagos Then
                ValorCuota = Round(valor - TotalCuota, 2)
            End If
            If mbooPorCobrar Then
                pck.Debe = Abs(ValorCuota)
            Else
                pck.Haber = Abs(ValorCuota)
            End If
            .TextMatrix(r, COL_VALOR) = ValorCuota
            TotalCuota = TotalCuota + ValorCuota
            If mbooPorCobrar Then
                pck.Debe = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
            Else
                pck.Haber = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
            End If
            .TextMatrix(r, COL_CODPROVCLI) = pck.CodProvCli     '*** MAKOTO 14/oct/00
            VisualizaProvCli r, pck.CodProvCli                  '***
            
            '***Agregado. 17/Ago/2004. Angel
            '***Para que se inserte la fila pero con la forma de pago predeterminada en la configuracion IVFIN
            If mbooPorCobrar Then
                If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FornmaCobroOtrasCuotas")) > 0 Then
                    .TextMatrix(r, COL_CODFORMA) = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FornmaCobroOtrasCuotas")
                    pck.codforma = Trim$(.TextMatrix(r, COL_CODFORMA))
                Else
                    .TextMatrix(r, COL_CODFORMA) = pck.codforma
                End If
            Else
                If Len(mobjGNComp.GNTrans.CodFormaPre) > 0 Then
                    If Not mobjGNComp.GNTrans.IVDescXFormaCP Then
                        .TextMatrix(r, COL_CODFORMA) = mobjGNComp.GNTrans.CodFormaPre
                    Else
                        If Len(mobjGNComp.CodFormnaCP) > 0 Then
                            .TextMatrix(r, COL_CODFORMA) = mobjGNComp.CodFormnaCP
                            BloqueaColumnaCodForma mobjGNComp.CodFormnaCP
                        End If
                    End If
                    pck.codforma = Trim$(.TextMatrix(r, COL_CODFORMA))
                Else
                    If mbooPorCobrar Then
                        If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroAnticipo")) > 0 Then
                            If mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroAnticipo") <> pck.codforma Then
                                pck.codforma = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroAnticipo")
                            End If
                        End If
                    Else
                        If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoAnticipo")) > 0 Then
                            If mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoAnticipo") <> pck.codforma Then
                                pck.codforma = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoAnticipo")
                            End If
                        End If
                    End If
                    .TextMatrix(r, COL_CODFORMA) = pck.codforma
                End If
            
            End If
            'mobjGNComp.codtrans & "-" & mobjGNComp.GNTrans.NumTransSiguiente &
            If mobjGNComp.GNTrans.IVPideNumDoc Then
                If Len(mobjGNComp.numDocRef) > 0 Then
                    pck.NumLetra = Mid$(mobjGNComp.numDocRef & " - " & NumPagos & "/" & NumeroPagos, 1, 20)
                Else
                    If mobjGNComp.GNTrans.CodPantalla = "TSIER" Then
                        pck.NumLetra = Mid$("Refinan - " & NumPagos & "/" & NumeroPagos, 1, 20)
                    Else
                        pck.NumLetra = Mid$(NumPagos & "/" & NumeroPagos, 1, 20)
                    End If
                End If
            Else
                pck.NumLetra = "Cuota " & NumPagos & "/" & NumeroPagos
            End If
            .TextMatrix(r, COL_NUMLETRA) = pck.NumLetra
            .TextMatrix(r, COL_FECHAEMI) = pck.FechaEmision
            pck.Orden = NumPagos

            mobjGNComp.DiaPago = DatePart("d", pck.FechaEmision)
            
            If Not BandDias Then
            
                If mbooPorCobrar Then
                    If mobjGNComp.DiaPago <> DatePart("d", pck.FechaEmision) Then
                        .TextMatrix(r, COL_FECHAVENCI) = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision))
                        
                        pck.FechaVenci = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision))
                    Else
                        If mobjGNComp.GNTrans.CodPantalla = "TSIER" Then
                            .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, mobjGNComp.FechaPrimerPago)
                            pck.FechaVenci = DateAdd("m", NumPagos - 1 + mobjGNComp.MesesGracia, mobjGNComp.FechaPrimerPago)
                        Else
                            .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
                            pck.FechaVenci = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
                        End If
                    End If
                Else
                        .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
                        pck.FechaVenci = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
                End If
            Else
            
                If mbooPorCobrar Then
                    If mobjGNComp.DiaPago <> DatePart("d", pck.FechaEmision) Then
                        .TextMatrix(r, COL_FECHAVENCI) = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision))
                        pck.FechaVenci = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision))
                    Else
                        If mobjGNComp.GNTrans.CodPantalla = "TSIER" Then
                            .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, mobjGNComp.FechaPrimerPago)
                            pck.FechaVenci = DateAdd("m", NumPagos - 1 + mobjGNComp.MesesGracia, mobjGNComp.FechaPrimerPago)
                        Else
                            .TextMatrix(r, COL_FECHAVENCI) = DateAdd("d", (NumPagos * intervalo) + mobjGNComp.MesesGracia, pck.FechaEmision)
                            pck.FechaVenci = DateAdd("d", (NumPagos * intervalo) + mobjGNComp.MesesGracia, pck.FechaEmision)
                        End If
                    End If
                Else
                        .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
                        pck.FechaVenci = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
                End If
            
            End If
            
            .TextMatrix(r, COL_PLAZO) = pck.FechaVenci - pck.FechaEmision
            
        If mobjGNComp.GNTrans.IVPideVendedor And Len(mobjGNComp.CodVendedor) > 0 Then
            Set Ven = mobjGNComp.Empresa.RecuperaFCVendedor(mobjGNComp.CodVendedor)
            .TextMatrix(r, COL_CODVEN) = Ven.CodVendedor
            .TextMatrix(r, COL_VENDEDOR) = Ven.nombre
            pck.IdVendedor = GNComprobante.IdVendedor
        End If
            
            
            If valorCuotasGeneradas < 0 Then
                NumPagos = NumeroPagos
            End If
            
            
        .Cell(flexcpBackColor, r, COL_FECHAEMI) = grd.BackColorFixed
        .Cell(flexcpBackColor, r, COL_FECHAEMI) = grd.BackColorFixed

            .Row = r
            If .Rows > .FixedRows Then
                .col = .FixedCols
                'Busca la primera columna
                For i = .FixedCols To .Cols - 1
                    If .ColData(i) >= 0 And (Not .ColHidden(i)) And .ColWidth(i) > 0 Then
                        .col = i
                        Exit For
                    End If
                Next i
            End If
            
        End With
    
        PoneNumFila
        VisualizaTotal
    Next NumPagos
    'If mobjGNComp.Empresa.GNOpcion.ObtenerValor("BloquearCuotas") = "1" Then
    If mobjGNComp.GNTrans.ivBloquearCuotas Then
        For i = COL_NUMLETRA To COL_VENDEDOR
            grd.ColData(i) = -1
        Next i
        HabilitarCtrlsGrupoCajaFormaCobro False
     End If
    
salida:
    Set pck = Nothing
'    grd.SetFocus
    Exit Sub
ErrTrap:
    Set pck = Nothing
    DispErr
    GoTo salida
End Sub


Public Sub AgregaFilasPorCreditoIntervalo(ByVal valor As Double, intervalo As Integer, BandDias As Boolean)
    Dim r As Long, r2 As Long, ix As Long, i As Long, ValorIni As Currency, saldo As Currency
    Dim pck As PCKardexCHP, Ven As FCVendedor
    Dim NumPagos As Integer, ValorCuota As Currency, CuotaFinal As Currency, TotalCuota As Currency
    Dim numdecimales As Integer, valorCuotasGeneradas As Currency
'    RaiseEvent PorAgregarFilaconPagoInicial(Saldo)        'Para calcular el valor predeterminado
    On Error GoTo ErrTrap
    If valor <= 0 Then Exit Sub
    TotalCuota = 0
    valorCuotasGeneradas = valor
    'If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("NumDecimalesCuotas")) > 0 Then
    If Len(mobjGNComp.GNTrans.IVNumDecimalesCuotas) > 0 Then
        'numdecimales = mobjGNComp.Empresa.GNOpcion.ObtenerValor("NumDecimalesCuotas")
        numdecimales = mobjGNComp.GNTrans.IVNumDecimalesCuotas
    Else
        numdecimales = 2
    End If
    ValorCuota = Round(valor / mobjGNComp.NumeroPagos, numdecimales)
    If mobjGNComp.DiaPago = 0 Then mobjGNComp.DiaPago = DatePart("d", mobjGNComp.FechaTrans)
    For NumPagos = 1 To mobjGNComp.NumeroPagos
        valorCuotasGeneradas = valorCuotasGeneradas - ValorCuota
        'Llama a agregar un objeto PCKardexCHP antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
        ix = mobjGNComp.AddPCKardex
        
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
    
            'Asigna la referencia al nuevo objeto a la fila nueva
            Set pck = mobjGNComp.PCKardexCHP(ix)
            .RowData(r) = pck
            
            'Proporciona el valor predeterminado        '*** MAKOTO 05/oct/00 Modificado
            
            If valorCuotasGeneradas < 0 Then
                ValorCuota = Round(valor - TotalCuota, 2)
'                NumPagos = mobjGNComp.NumeroPagos
            End If
            
            If NumPagos = mobjGNComp.NumeroPagos Then
                ValorCuota = Round(valor - TotalCuota, 2)
            End If
            If mbooPorCobrar Then
                pck.Debe = Abs(ValorCuota)
            Else
                pck.Haber = Abs(ValorCuota)
            End If
            .TextMatrix(r, COL_VALOR) = ValorCuota
            TotalCuota = TotalCuota + ValorCuota
            If mbooPorCobrar Then
                pck.Debe = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
            Else
                pck.Haber = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
            End If
            .TextMatrix(r, COL_CODPROVCLI) = pck.CodProvCli     '*** MAKOTO 14/oct/00
            VisualizaProvCli r, pck.CodProvCli                  '***
            
            '***Agregado. 17/Ago/2004. Angel
            '***Para que se inserte la fila pero con la forma de pago predeterminada en la configuracion IVFIN
            If mbooPorCobrar Then
                If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FornmaCobroOtrasCuotas")) > 0 Then
                    .TextMatrix(r, COL_CODFORMA) = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FornmaCobroOtrasCuotas")
                    pck.codforma = Trim$(.TextMatrix(r, COL_CODFORMA))
                Else
                    .TextMatrix(r, COL_CODFORMA) = pck.codforma
                End If
            Else
                If Len(mobjGNComp.GNTrans.CodFormaPre) > 0 Then
                    If Not mobjGNComp.GNTrans.IVDescXFormaCP Then
                        .TextMatrix(r, COL_CODFORMA) = mobjGNComp.GNTrans.CodFormaPre
                    Else
                        If Len(mobjGNComp.CodFormnaCP) > 0 Then
                            .TextMatrix(r, COL_CODFORMA) = mobjGNComp.CodFormnaCP
                            BloqueaColumnaCodForma mobjGNComp.CodFormnaCP
                        End If
                    End If
                    pck.codforma = Trim$(.TextMatrix(r, COL_CODFORMA))
                Else
                    If mbooPorCobrar Then
                        If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroAnticipo")) > 0 Then
                            If mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroAnticipo") <> pck.codforma Then
                                pck.codforma = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroAnticipo")
                            End If
                        End If
                    Else
                        If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoAnticipo")) > 0 Then
                            If mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoAnticipo") <> pck.codforma Then
                                pck.codforma = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoAnticipo")
                            End If
                        End If
                    End If
                    .TextMatrix(r, COL_CODFORMA) = pck.codforma
                End If
            
            End If
            'mobjGNComp.codtrans & "-" & mobjGNComp.GNTrans.NumTransSiguiente &
            If mobjGNComp.GNTrans.IVPideNumDoc Then
                If Len(mobjGNComp.numDocRef) > 0 Then
                    pck.NumLetra = Mid$(mobjGNComp.numDocRef & " - " & NumPagos & "/" & mobjGNComp.NumeroPagos, 1, 20)
                Else
                    If mobjGNComp.GNTrans.CodPantalla = "TSIER" Then
                        pck.NumLetra = Mid$("Refinan - " & NumPagos & "/" & mobjGNComp.NumeroPagos, 1, 20)
                    Else
                        pck.NumLetra = Mid$(NumPagos & "/" & mobjGNComp.NumeroPagos, 1, 20)
                    End If
                End If
            Else
                pck.NumLetra = "Cuota " & NumPagos & "/" & mobjGNComp.NumeroPagos
            End If
            .TextMatrix(r, COL_NUMLETRA) = pck.NumLetra
            .TextMatrix(r, COL_FECHAEMI) = pck.FechaEmision

            If Not BandDias Then
                If mbooPorCobrar Then
                    If mobjGNComp.DiaPago <> DatePart("d", pck.FechaEmision) Then
                        .TextMatrix(r, COL_FECHAVENCI) = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision))
                        pck.FechaVenci = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision))
                    Else
                        If mobjGNComp.GNTrans.CodPantalla = "TSIER" Then
                            .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, mobjGNComp.FechaPrimerPago)
                            pck.FechaVenci = DateAdd("m", NumPagos - 1 + mobjGNComp.MesesGracia, mobjGNComp.FechaPrimerPago)
                        Else
                            .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", (NumPagos * intervalo) + mobjGNComp.MesesGracia, pck.FechaEmision)
                            pck.FechaVenci = DateAdd("m", (NumPagos * intervalo) + mobjGNComp.MesesGracia, pck.FechaEmision)
                        End If
                    End If
                Else
                    .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
                    pck.FechaVenci = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
                End If
        Else
                If mbooPorCobrar Then
                    If mobjGNComp.DiaPago <> DatePart("d", pck.FechaEmision) Then
                        .TextMatrix(r, COL_FECHAVENCI) = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision))
                        pck.FechaVenci = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision))
                    Else
                        If mobjGNComp.GNTrans.CodPantalla = "TSIER" Then
                            .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, mobjGNComp.FechaPrimerPago)
                            pck.FechaVenci = DateAdd("m", NumPagos - 1 + mobjGNComp.MesesGracia, mobjGNComp.FechaPrimerPago)
                        Else
                            .TextMatrix(r, COL_FECHAVENCI) = DateAdd("d", (NumPagos * intervalo) + mobjGNComp.MesesGracia, pck.FechaEmision)
                            pck.FechaVenci = DateAdd("d", (NumPagos * intervalo) + mobjGNComp.MesesGracia, pck.FechaEmision)
                        End If
                    End If
                Else
                        .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
                        pck.FechaVenci = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
                End If
        End If
            .TextMatrix(r, COL_PLAZO) = pck.FechaVenci - pck.FechaEmision
            
        If mobjGNComp.GNTrans.IVPideVendedor And Len(mobjGNComp.CodVendedor) > 0 Then
            Set Ven = mobjGNComp.Empresa.RecuperaFCVendedor(mobjGNComp.CodVendedor)
            .TextMatrix(r, COL_CODVEN) = Ven.CodVendedor
            .TextMatrix(r, COL_VENDEDOR) = Ven.nombre
            pck.IdVendedor = GNComprobante.IdVendedor
        End If
            
            
            If valorCuotasGeneradas < 0 Then
                NumPagos = mobjGNComp.NumeroPagos
            End If
            
            
        .Cell(flexcpBackColor, r, COL_FECHAEMI) = grd.BackColorFixed
        .Cell(flexcpBackColor, r, COL_FECHAEMI) = grd.BackColorFixed

            .Row = r
            If .Rows > .FixedRows Then
                .col = .FixedCols
                'Busca la primera columna
                For i = .FixedCols To .Cols - 1
                    If .ColData(i) >= 0 And (Not .ColHidden(i)) And .ColWidth(i) > 0 Then
                        .col = i
                        Exit For
                    End If
                Next i
            End If
''''         HabilitarCtrlsGrupoCaja False
            
        End With
    
        PoneNumFila
        VisualizaTotal
    Next NumPagos
    'If mobjGNComp.Empresa.GNOpcion.ObtenerValor("BloquearCuotas") = "1" Then
    If mobjGNComp.GNTrans.ivBloquearCuotas Then
        For i = COL_NUMLETRA To COL_VENDEDOR
            grd.ColData(i) = -1
        Next i
        HabilitarCtrlsGrupoCajaFormaCobro False
     End If
    
salida:
    Set pck = Nothing
'    grd.SetFocus
    Exit Sub
ErrTrap:
    Set pck = Nothing
    DispErr
    GoTo salida
End Sub

Public Sub AgregaFilasPorCreditoIntervaloPagosIguales(ByVal valor As Double, intervalo As Integer, BandDias As Boolean)
    Dim r As Long, r2 As Long, ix As Long, i As Long, ValorIni As Currency, saldo As Currency
    Dim pck As PCKardexCHP, Ven As FCVendedor
    Dim NumPagos As Integer, ValorCuota As Currency, CuotaFinal As Currency, TotalCuota As Currency
    Dim numdecimales As Integer, valorCuotasGeneradas As Currency
'    RaiseEvent PorAgregarFilaconPagoInicial(Saldo)        'Para calcular el valor predeterminado
    On Error GoTo ErrTrap
    If valor <= 0 Then Exit Sub
    If mobjGNComp.NumeroPagos = 0 Then: MsgBox "Tiene que asignar Numero de pagos al proveedor", vbInformation: Exit Sub
    If intervalo = 0 Then MsgBox "Tiene que asignar un intervalo de pagos al proveedor", vbInformation: Exit Sub
    
    TotalCuota = 0
    'valorCuotasGeneradas = valor
'    'If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("NumDecimalesCuotas")) > 0 Then
'    If Len(mobjGNComp.GNTrans.ivNumDecimalesCuotas) > 0 Then
'        'numdecimales = mobjGNComp.Empresa.GNOpcion.ObtenerValor("NumDecimalesCuotas")
'        numdecimales = mobjGNComp.GNTrans.ivNumDecimalesCuotas
'    Else
        numdecimales = 2
'    End If
    ValorCuota = Round(valor / mobjGNComp.NumeroPagos, numdecimales)
    For NumPagos = 1 To mobjGNComp.NumeroPagos
      '  valorCuotasGeneradas = valorCuotasGeneradas - ValorCuota
        'Llama a agregar un objeto PCKardexCHP antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
        ix = mobjGNComp.AddPCKardex
        
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
    
            'Asigna la referencia al nuevo objeto a la fila nueva
            Set pck = mobjGNComp.PCKardexCHP(ix)
            .RowData(r) = pck
            
            'Proporciona el valor predeterminado        '*** MAKOTO 05/oct/00 Modificado
            
            'If valorCuotasGeneradas < 0 Then
             '   ValorCuota = Round(valor - TotalCuota, 2)
'                NumPagos = mobjGNComp.NumeroPagos
            'End If
            
            If NumPagos = mobjGNComp.NumeroPagos Then
                ValorCuota = Round(valor - TotalCuota, 2)
            End If
            If mbooPorCobrar Then
                pck.Debe = Abs(ValorCuota)
            Else
                pck.Haber = Abs(ValorCuota)
            End If
            .TextMatrix(r, COL_VALOR) = ValorCuota
            TotalCuota = TotalCuota + ValorCuota
            If mbooPorCobrar Then
                pck.Debe = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
            Else
                pck.Haber = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
            End If
            .TextMatrix(r, COL_CODPROVCLI) = pck.CodProvCli     '*** MAKOTO 14/oct/00
            VisualizaProvCli r, pck.CodProvCli                  '***
            
            '***Agregado. 17/Ago/2004. Angel
            '***Para que se inserte la fila pero con la forma de pago predeterminada en la configuracion IVFIN
            If mbooPorCobrar Then
                If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FornmaCobroOtrasCuotas")) > 0 Then
                    .TextMatrix(r, COL_CODFORMA) = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FornmaCobroOtrasCuotas")
                    pck.codforma = Trim$(.TextMatrix(r, COL_CODFORMA))
                Else
                    .TextMatrix(r, COL_CODFORMA) = pck.codforma
                End If
            Else
                If Len(mobjGNComp.GNTrans.CodFormaPre) > 0 Then
                    If Not mobjGNComp.GNTrans.IVDescXFormaCP Then
                        .TextMatrix(r, COL_CODFORMA) = mobjGNComp.GNTrans.CodFormaPre
                    Else
                        If Len(mobjGNComp.CodFormnaCP) > 0 Then
                            .TextMatrix(r, COL_CODFORMA) = mobjGNComp.CodFormnaCP
                            BloqueaColumnaCodForma mobjGNComp.CodFormnaCP
                        End If
                    End If
                    pck.codforma = Trim$(.TextMatrix(r, COL_CODFORMA))
                Else
                    If mbooPorCobrar Then
                        If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroAnticipo")) > 0 Then
                            If mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroAnticipo") <> pck.codforma Then
                                pck.codforma = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroAnticipo")
                            End If
                        End If
                    Else
                        If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoAnticipo")) > 0 Then
                            If mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoAnticipo") <> pck.codforma Then
                                pck.codforma = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoAnticipo")
                            End If
                        End If
                    End If
                    .TextMatrix(r, COL_CODFORMA) = pck.codforma
                End If
            
            End If
            'mobjGNComp.codtrans & "-" & mobjGNComp.GNTrans.NumTransSiguiente &
            If mobjGNComp.GNTrans.IVPideNumDoc Then
                If Len(mobjGNComp.numDocRef) > 0 Then
                    pck.NumLetra = Mid$(mobjGNComp.numDocRef & " - " & NumPagos & "/" & mobjGNComp.NumeroPagos, 1, 20)
                Else
                    If mobjGNComp.GNTrans.CodPantalla = "TSIER" Then
                        pck.NumLetra = Mid$("Refinan - " & NumPagos & "/" & mobjGNComp.NumeroPagos, 1, 20)
                    Else
                        pck.NumLetra = Mid$(NumPagos & "/" & mobjGNComp.NumeroPagos, 1, 20)
                    End If
                End If
            Else
                pck.NumLetra = "Cuota " & NumPagos & "/" & mobjGNComp.NumeroPagos
            End If
            .TextMatrix(r, COL_NUMLETRA) = pck.NumLetra
            .TextMatrix(r, COL_FECHAEMI) = pck.FechaEmision

            If Not BandDias Then
                If mbooPorCobrar Then
                    If mobjGNComp.DiaPago <> DatePart("d", pck.FechaEmision) Then
                        .TextMatrix(r, COL_FECHAVENCI) = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision))
                        pck.FechaVenci = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision))
                    Else
                        If mobjGNComp.GNTrans.CodPantalla = "TSIER" Then
                            .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, mobjGNComp.FechaPrimerPago)
                            pck.FechaVenci = DateAdd("m", NumPagos - 1 + mobjGNComp.MesesGracia, mobjGNComp.FechaPrimerPago)
                        Else
                            .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", (NumPagos * intervalo) + mobjGNComp.MesesGracia, pck.FechaEmision)
                            pck.FechaVenci = DateAdd("m", (NumPagos * intervalo) + mobjGNComp.MesesGracia, pck.FechaEmision)
                        End If
                    End If
                Else
                    .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
                    pck.FechaVenci = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
                End If
        Else
                If mbooPorCobrar Then
                    If mobjGNComp.DiaPago <> DatePart("d", pck.FechaEmision) Then
                        .TextMatrix(r, COL_FECHAVENCI) = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision))
                        pck.FechaVenci = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision))
                    Else
                        If mobjGNComp.GNTrans.CodPantalla = "TSIER" Then
                            .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, mobjGNComp.FechaPrimerPago)
                            pck.FechaVenci = DateAdd("m", NumPagos - 1 + mobjGNComp.MesesGracia, mobjGNComp.FechaPrimerPago)
                        Else
                            .TextMatrix(r, COL_FECHAVENCI) = DateAdd("d", (NumPagos * intervalo) + mobjGNComp.MesesGracia, pck.FechaEmision)
                            pck.FechaVenci = DateAdd("d", (NumPagos * intervalo) + mobjGNComp.MesesGracia, pck.FechaEmision)
                        End If
                    End If
                Else
                        .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
                        pck.FechaVenci = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
                End If
        End If
            .TextMatrix(r, COL_PLAZO) = pck.FechaVenci - pck.FechaEmision
            
        If mobjGNComp.GNTrans.IVPideVendedor And Len(mobjGNComp.CodVendedor) > 0 Then
            Set Ven = mobjGNComp.Empresa.RecuperaFCVendedor(mobjGNComp.CodVendedor)
            .TextMatrix(r, COL_CODVEN) = Ven.CodVendedor
            .TextMatrix(r, COL_VENDEDOR) = Ven.nombre
            pck.IdVendedor = GNComprobante.IdVendedor
        End If
            
            
            If valorCuotasGeneradas < 0 Then
                NumPagos = mobjGNComp.NumeroPagos
            End If
            
            
        .Cell(flexcpBackColor, r, COL_FECHAEMI) = grd.BackColorFixed
        .Cell(flexcpBackColor, r, COL_FECHAEMI) = grd.BackColorFixed

            .Row = r
            If .Rows > .FixedRows Then
                .col = .FixedCols
                'Busca la primera columna
                For i = .FixedCols To .Cols - 1
                    If .ColData(i) >= 0 And (Not .ColHidden(i)) And .ColWidth(i) > 0 Then
                        .col = i
                        Exit For
                    End If
                Next i
            End If
''''         HabilitarCtrlsGrupoCaja False
            
        End With
    
        PoneNumFila
        VisualizaTotal
    Next NumPagos
    'If mobjGNComp.Empresa.GNOpcion.ObtenerValor("BloquearCuotas") = "1" Then
    If mobjGNComp.GNTrans.ivBloquearCuotas Then
        For i = COL_NUMLETRA To COL_VENDEDOR
            grd.ColData(i) = -1
        Next i
        HabilitarCtrlsGrupoCajaFormaCobro False
     End If
    
salida:
    Set pck = Nothing
'    grd.SetFocus
    Exit Sub
ErrTrap:
    Set pck = Nothing
    DispErr
    GoTo salida
End Sub
Public Property Get FontSize() As Single
    FontSize = grd.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    grd.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property


Public Sub ConfigCols_PlazoModificable()
    Dim i As Integer
    
    With grd
    
    
    
        
        For i = 1 To .Rows - 1
            If Not .IsSubtotal(i) Then
                        .Cell(flexcpData, i, COL_PLAZO) = 1
                        .Cell(flexcpData, i, COL_FECHAVENCI) = 1
                        .Cell(flexcpBackColor, i, COL_PLAZO) = vbWhite
                        .Cell(flexcpBackColor, i, COL_FECHAVENCI) = vbWhite

            End If
        Next i
        
    End With
End Sub


'Public Sub AgregaFilasxCheque(valor As Currency, NumeroPagos As Integer, intervalo As Integer, BandDias As Boolean)
Public Sub AgregaFilasxPagoconCheque(codBanco As String, numche As String, numcta As String, titular As String, FechaVenci As Date)
    Dim r As Long, r2 As Long, ix As Long, i As Long, ValorIni As Currency, saldo As Currency
    Dim pck As PCKardexCHP, Ven As FCVendedor, NumPagos  As Integer
    Dim ValorCuota As Currency, CuotaFinal As Currency, TotalCuota As Currency
    Dim numdecimales As Integer, valorCuotasGeneradas As Currency
    Dim valor As Currency, NumeroPagos As Integer, BandDias As Integer, intervalo As Integer, j As Long
    Dim gcaux As GNComprobante, TransID As Long
'    RaiseEvent PorAgregarFilaconPagoInicial(Saldo)        'Para calcular el valor predeterminado
    On Error GoTo ErrTrap
     valor = 0
    NumeroPagos = 0
    If mobjGNComp.CountPCKardex > 0 Then
        For i = 1 To mobjGNComp.CountPCKardex
            If mobjGNComp.PCKardexCHP(i).idAsignado = 0 Then
                mobjGNComp.BorrarPCKardex
            Else
                valor = valor + mobjGNComp.PCKardexCHP(i).Haber
                NumeroPagos = NumeroPagos + 1
           End If
        Next i
    End If
  
    
    
'    TotalCuota = valor
 '   valorCuotasGeneradas = valor



'    If Len(mobjGNComp.GNTrans.ivNumDecimalesCuotas) > 0 Then
'        numdecimales = mobjGNComp.GNTrans.ivNumDecimalesCuotas
'    Else
    numdecimales = 2
'    End If
'    If NumeroPagos = 0 Then NumeroPagos = 1
'    ValorCuota = Round(valor / NumeroPagos, numdecimales)
    For NumPagos = 1 To NumeroPagos
'        valorCuotasGeneradas = valorCuotasGeneradas - ValorCuota
        'Llama a agregar un objeto PCKardexCHP antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
        ix = mobjGNComp.AddPCKardex
        i = NumPagos
        
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
    
            'Asigna la referencia al nuevo objeto a la fila nueva
            Set pck = mobjGNComp.PCKardexCHP(ix)
            .RowData(r) = pck
            
            'Proporciona el valor predeterminado        '*** MAKOTO 05/oct/00 Modificado
            
'            If NumPagos = NumeroPagos Then
'            'If valorCuotasGeneradas < 0 Then
'                ValorCuota = Round(valor - TotalCuota, 2)
''                NumPagos = mobjGNComp.NumeroPagos
'            End If
            
'            If NumPagos = mobjGNComp.NumeroPagos Then
'                ValorCuota = Round(valor - TotalCuota, 2)
'            End If
            .TextMatrix(r, COL_CODPROVCLI) = pck.CodProvCli     '*** MAKOTO 14/oct/00
            VisualizaProvCli r, pck.CodProvCli                  '***
            
            '***Agregado. 17/Ago/2004. Angel
            '***Para que se inserte la fila pero con la forma de pago predeterminada en la configuracion IVFIN
            If mbooPorCobrar Then
                If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FornmaCobroOtrasCuotas")) > 0 Then
                    .TextMatrix(r, COL_CODFORMA) = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FornmaCobroOtrasCuotas")
                    pck.codforma = Trim$(.TextMatrix(r, COL_CODFORMA))
                Else
                    .TextMatrix(r, COL_CODFORMA) = "CHR" 'pck.codforma
                End If
            Else
                If Len(mobjGNComp.GNTrans.CodFormaPre) > 0 Then
                    If Not mobjGNComp.GNTrans.IVDescXFormaCP Then
                        .TextMatrix(r, COL_CODFORMA) = mobjGNComp.GNTrans.CodFormaPre
                    Else
                        If Len(mobjGNComp.CodFormnaCP) > 0 Then
                            .TextMatrix(r, COL_CODFORMA) = mobjGNComp.CodFormnaCP
                            BloqueaColumnaCodForma mobjGNComp.CodFormnaCP
                        End If
                    End If
                    pck.codforma = Trim$(.TextMatrix(r, COL_CODFORMA))
                Else
                    If mbooPorCobrar Then
                        If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroAnticipo")) > 0 Then
                            If mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroAnticipo") <> pck.codforma Then
                                pck.codforma = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroAnticipo")
                            End If
                        End If
                    Else
                        If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoAnticipo")) > 0 Then
                            If mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoAnticipo") <> pck.codforma Then
                                pck.codforma = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoAnticipo")
                            End If
                        End If
                    End If
                    .TextMatrix(r, COL_CODFORMA) = pck.codforma
                End If
            
            End If
            'mobjGNComp.codtrans & "-" & mobjGNComp.GNTrans.NumTransSiguiente &
            TransID = mobjGNComp.Empresa.RecuperarIDGncomprobantexIdAsignado(mobjGNComp.PCKardexCHP(i).idAsignado)
            Set gncAux = mobjGNComp.Empresa.RecuperaGNComprobante(TransID)
'            i = NumPagos
            
                For j = 1 To gncAux.CountPCKardex
                    If gncAux.PCKardexCHP(j).id = mobjGNComp.PCKardexCHP(i).idAsignado Then
                    
                        If mbooPorCobrar Then
                            pck.Debe = Abs(mobjGNComp.PCKardexCHP(i).Haber)
                        Else
                            pck.Haber = Abs(ValorCuota)
                        End If
                        .TextMatrix(r, COL_VALOR) = Abs(mobjGNComp.PCKardexCHP(i).Haber)
                        'TotalCuota = TotalCuota + ValorCuota
                        If mbooPorCobrar Then
                            pck.Debe = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
                        Else
                            pck.Haber = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
                        End If
                    
                        If gncAux.GNTrans.Modulo <> "IV" Then
                            pck.NumLetra = gncAux.PCKardexCHP(j).NumLetra
                            pck.CodVendedor = gncAux.PCKardexCHP(j).CodVendedor
                        Else
                            pck.NumLetra = gncAux.CodTrans & " " & gncAux.numtrans
                            pck.CodVendedor = gncAux.CodVendedor
                            
                        End If
                        mobjGNComp.CodVendedor = gncAux.CodVendedor
                        
                        
                        pck.FechaEmision = gncAux.PCKardexCHP(j).FechaEmision
                        pck.FechaVenci = gncAux.PCKardexCHP(j).FechaVenci
                        pck.CodProvCli = gncAux.PCKardexCHP(j).CodProvCli
                        
                        pck.codBanco = codBanco
                        pck.Numcheque = numche
                        pck.NumCuenta = numcta
                        pck.TitularCta = titular
                        pck.Observacion = codBanco & "-" & numche & "-" & FechaVenci
                        Exit For
                    End If
                Next j
            'Else
              '  pck.NumLetra = gncAux.codtrans & " " & gncAux.NumTrans
            'End If
'            If mobjGNComp.GNTrans.IVPideNumDoc Then
'                If Len(mobjGNComp.numDocRef) > 0 Then
'                    pck.NumLetra = Mid$(mobjGNComp.numDocRef & " - " & NumPagos & "/" & NumeroPagos, 1, 20)
'                Else
'                    If mobjGNComp.GNTrans.codPantalla = "TSIER" Then
'                        pck.NumLetra = Mid$("Refinan - " & NumPagos & "/" & NumeroPagos, 1, 20)
'                    Else
'                        pck.NumLetra = Mid$(NumPagos & "/" & NumeroPagos, 1, 20)
'                    End If
'                End If
'            Else
'                pck.NumLetra = "Cuota " & NumPagos & "/" & NumeroPagos
'            End If
            .TextMatrix(r, COL_NUMLETRA) = pck.NumLetra
            .TextMatrix(r, COL_FECHAEMI) = pck.FechaEmision
            .TextMatrix(r, COL_FECHAVENCI) = pck.FechaVenci
            pck.Orden = NumPagos

'            mobjGNComp.DiaPago = DatePart("d", pck.FechaEmision)
'
'            If Not BandDias Then
'
'                If mbooPorCobrar Then
'                    If mobjGNComp.DiaPago <> DatePart("d", pck.FechaEmision) Then
'                        .TextMatrix(r, COL_FECHAVENCI) = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision))
'
'                        pck.FechaVenci = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision))
'                    Else
'                        If mobjGNComp.GNTrans.codPantalla = "TSIER" Then
'                            .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, mobjGNComp.FechaPrimerPago)
'                            pck.FechaVenci = DateAdd("m", NumPagos - 1 + mobjGNComp.MesesGracia, mobjGNComp.FechaPrimerPago)
'                        Else
'                            .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
'                            pck.FechaVenci = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
'                        End If
'                    End If
'                Else
'                        .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
'                        pck.FechaVenci = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
'                End If
'            Else
'
'                If mbooPorCobrar Then
'                    If mobjGNComp.DiaPago <> DatePart("d", pck.FechaEmision) Then
'                        .TextMatrix(r, COL_FECHAVENCI) = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision))
'                        pck.FechaVenci = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision))
'                    Else
'                        If mobjGNComp.GNTrans.codPantalla = "TSIER" Then
'                            .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, mobjGNComp.FechaPrimerPago)
'                            pck.FechaVenci = DateAdd("m", NumPagos - 1 + mobjGNComp.MesesGracia, mobjGNComp.FechaPrimerPago)
'                        Else
'                            .TextMatrix(r, COL_FECHAVENCI) = DateAdd("d", (NumPagos * intervalo) + mobjGNComp.MesesGracia, pck.FechaEmision)
'                            pck.FechaVenci = DateAdd("d", (NumPagos * intervalo) + mobjGNComp.MesesGracia, pck.FechaEmision)
'                        End If
'                    End If
'                Else
'                        .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
'                        pck.FechaVenci = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pck.FechaEmision)
'                End If
'
'            End If
            
            .TextMatrix(r, COL_PLAZO) = pck.FechaVenci - pck.FechaEmision
            
'        If mobjGNComp.GNTrans.IVPideVendedor And Len(mobjGNComp.CodVendedor) > 0 Then
'            Set Ven = mobjGNComp.Empresa.RecuperaFCVendedor(mobjGNComp.CodVendedor)
'            .TextMatrix(r, COL_CODVEN) = Ven.CodVendedor
'            .TextMatrix(r, COL_VENDEDOR) = Ven.nombre
'            pck.IdVendedor = GNComprobante.IdVendedor
'        End If
            
            
            If valorCuotasGeneradas < 0 Then
                NumPagos = NumeroPagos
            End If
            
            
        .Cell(flexcpBackColor, r, COL_FECHAEMI) = grd.BackColorFixed
        .Cell(flexcpBackColor, r, COL_FECHAEMI) = grd.BackColorFixed

            .Row = r
            If .Rows > .FixedRows Then
                .col = .FixedCols
                'Busca la primera columna
                For i = .FixedCols To .Cols - 1
                    If .ColData(i) >= 0 And (Not .ColHidden(i)) And .ColWidth(i) > 0 Then
                        .col = i
                        Exit For
                    End If
                Next i
            End If
            
        End With
    
        PoneNumFila
        VisualizaTotal
    Next NumPagos
    'If mobjGNComp.Empresa.GNOpcion.ObtenerValor("BloquearCuotas") = "1" Then
    If mobjGNComp.GNTrans.ivBloquearCuotas Then
        For i = COL_NUMLETRA To COL_VENDEDOR
            grd.ColData(i) = -1
        Next i
        HabilitarCtrlsGrupoCajaFormaCobro False
     End If
    
salida:
    Set pck = Nothing
'    grd.SetFocus
    Exit Sub
ErrTrap:
    Set pck = Nothing
    DispErr
    GoTo salida
End Sub

Public Sub AgregaFilasxPagoconChequeCHP(codBanco As String, numche As String, numcta As String, titular As String, FechaVenci As Date)
    Dim r As Long, r2 As Long, ix As Long, i As Long, ValorIni As Currency, saldo As Currency
    Dim pckCHP As PCKardexCHP, Ven As FCVendedor, NumPagos  As Integer, pck As PCKardexCHP
    Dim ValorCuota As Currency, CuotaFinal As Currency, TotalCuota As Currency
    Dim numdecimales As Integer, valorCuotasGeneradas As Currency
    Dim valor As Currency, NumeroPagos As Integer, BandDias As Integer, intervalo As Integer, j As Long
    Dim gcaux As GNComprobante, TransID As Long
'    RaiseEvent PorAgregarFilaconPagoInicial(Saldo)        'Para calcular el valor predeterminado
    On Error GoTo ErrTrap
     valor = 0
    NumeroPagos = 0
    If mobjGNComp.CountPCKardexCHP > 0 Then
        For i = 1 To mobjGNComp.CountPCKardexCHP
            If mobjGNComp.PCKardexCHP(i).idAsignado = 0 Then
                mobjGNComp.BorrarPCKardexCHP
            Else
                valor = valor + mobjGNComp.PCKardexCHP(i).Haber
                NumeroPagos = NumeroPagos + 1
           End If
        Next i
    End If
  
  
    If mobjGNComp.CountPCKardex > 0 Then
        For i = 1 To mobjGNComp.CountPCKardex
                valor = valor + mobjGNComp.PCKardexCHP(i).Haber
                NumeroPagos = NumeroPagos + 1
           
        Next i
    End If
  
    
    
'    TotalCuota = valor
 '   valorCuotasGeneradas = valor



'    If Len(mobjGNComp.GNTrans.ivNumDecimalesCuotas) > 0 Then
'        numdecimales = mobjGNComp.GNTrans.ivNumDecimalesCuotas
'    Else
    numdecimales = 2
'    End If
'    If NumeroPagos = 0 Then NumeroPagos = 1
'    ValorCuota = Round(valor / NumeroPagos, numdecimales)
    For NumPagos = 1 To NumeroPagos
'        valorCuotasGeneradas = valorCuotasGeneradas - ValorCuota
        'Llama a agregar un objeto PCKardexCHP antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
        ix = mobjGNComp.AddPCKardexCHP
        i = NumPagos
        
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
    
            'Asigna la referencia al nuevo objeto a la fila nueva
            Set pckCHP = mobjGNComp.PCKardexCHP(ix)
            .RowData(r) = pckCHP
            
            'Proporciona el valor predeterminado        '*** MAKOTO 05/oct/00 Modificado
            
'            If NumPagos = NumeroPagos Then
'            'If valorCuotasGeneradas < 0 Then
'                ValorCuota = Round(valor - TotalCuota, 2)
''                NumPagos = mobjGNComp.NumeroPagos
'            End If
            
'            If NumPagos = mobjGNComp.NumeroPagos Then
'                ValorCuota = Round(valor - TotalCuota, 2)
'            End If
            .TextMatrix(r, COL_CODPROVCLI) = pckCHP.CodProvCli     '*** MAKOTO 14/oct/00
            VisualizaProvCli r, pckCHP.CodProvCli                  '***
            
            '***Agregado. 17/Ago/2004. Angel
            '***Para que se inserte la fila pero con la forma de pago predeterminada en la configuracion IVFIN
            If mbooPorCobrar Then
                If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FornmaCobroOtrasCuotas")) > 0 Then
                    .TextMatrix(r, COL_CODFORMA) = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FornmaCobroOtrasCuotas")
                    pckCHP.codforma = Trim$(.TextMatrix(r, COL_CODFORMA))
                Else
                    .TextMatrix(r, COL_CODFORMA) = "CHR" 'pckCHP.codforma
                    pckCHP.codforma = "CHR"
                End If
            Else
                If Len(mobjGNComp.GNTrans.CodFormaPre) > 0 Then
                    If Not mobjGNComp.GNTrans.IVDescXFormaCP Then
                        .TextMatrix(r, COL_CODFORMA) = mobjGNComp.GNTrans.CodFormaPre
                    Else
                        If Len(mobjGNComp.CodFormnaCP) > 0 Then
                            .TextMatrix(r, COL_CODFORMA) = mobjGNComp.CodFormnaCP
                            BloqueaColumnaCodForma mobjGNComp.CodFormnaCP
                        End If
                    End If
                    pckCHP.codforma = Trim$(.TextMatrix(r, COL_CODFORMA))
                Else
                    If mbooPorCobrar Then
                        If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroAnticipo")) > 0 Then
                            If mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroAnticipo") <> pckCHP.codforma Then
                                pckCHP.codforma = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaCobroAnticipo")
                            End If
                        End If
                    Else
                        If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoAnticipo")) > 0 Then
                            If mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoAnticipo") <> pckCHP.codforma Then
                                pckCHP.codforma = mobjGNComp.Empresa.GNOpcion.ObtenerValor("FormaPagoAnticipo")
                            End If
                        End If
                    End If
                    .TextMatrix(r, COL_CODFORMA) = pckCHP.codforma
                End If
            
            End If
            'mobjGNComp.codtrans & "-" & mobjGNComp.GNTrans.NumTransSiguiente &
            TransID = mobjGNComp.Empresa.RecuperarIDGncomprobantexIdAsignado(mobjGNComp.PCKardexCHP(i).idAsignado)
            Set gncAux = mobjGNComp.Empresa.RecuperaGNComprobante(TransID)
'            i = NumPagos
            
                For j = 1 To gncAux.CountPCKardex
                    If gncAux.PCKardexCHP(j).id = mobjGNComp.PCKardexCHP(i).idAsignado Then
                    
                        If mbooPorCobrar Then
                            pckCHP.Debe = Abs(mobjGNComp.PCKardexCHP(i).Haber)
                        Else
                            pckCHP.Haber = Abs(ValorCuota)
                        End If
                        .TextMatrix(r, COL_VALOR) = Abs(mobjGNComp.PCKardexCHP(i).Haber)
                        'TotalCuota = TotalCuota + ValorCuota
                        If mbooPorCobrar Then
                            pckCHP.Debe = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
                        Else
                            pckCHP.Haber = MiCCur(.Cell(flexcpTextDisplay, r, COL_VALOR))
                        End If
                    
                        If gncAux.GNTrans.Modulo <> "IV" Then
                            pckCHP.NumLetra = gncAux.PCKardexCHP(j).NumLetra
                            pckCHP.CodVendedor = gncAux.PCKardexCHP(j).CodVendedor
                        Else
                            pckCHP.NumLetra = gncAux.CodTrans & " " & gncAux.numtrans
                            pckCHP.CodVendedor = gncAux.CodVendedor
                            
                        End If
                        mobjGNComp.CodVendedor = gncAux.CodVendedor
                        
                        
                        pckCHP.FechaEmision = gncAux.PCKardexCHP(j).FechaEmision
                        pckCHP.FechaVenci = gncAux.PCKardexCHP(j).FechaVenci
                        pckCHP.CodProvCli = gncAux.PCKardexCHP(j).CodProvCli
                        
                        pckCHP.codBanco = codBanco
                        pckCHP.Numcheque = numche
                        pckCHP.NumCuenta = numcta
                        pckCHP.TitularCta = titular
                        pckCHP.Observacion = codBanco & "-" & numche & "-" & FechaVenci
                        pckCHP.FechaVenci = FechaVenci
                        Exit For
                    End If
                Next j
            'Else
              '  pckCHP.NumLetra = gncAux.codtrans & " " & gncAux.NumTrans
            'End If
'            If mobjGNComp.GNTrans.IVPideNumDoc Then
'                If Len(mobjGNComp.numDocRef) > 0 Then
'                    pckCHP.NumLetra = Mid$(mobjGNComp.numDocRef & " - " & NumPagos & "/" & NumeroPagos, 1, 20)
'                Else
'                    If mobjGNComp.GNTrans.codPantalla = "TSIER" Then
'                        pckCHP.NumLetra = Mid$("Refinan - " & NumPagos & "/" & NumeroPagos, 1, 20)
'                    Else
'                        pckCHP.NumLetra = Mid$(NumPagos & "/" & NumeroPagos, 1, 20)
'                    End If
'                End If
'            Else
'                pckCHP.NumLetra = "Cuota " & NumPagos & "/" & NumeroPagos
'            End If
            .TextMatrix(r, COL_NUMLETRA) = pckCHP.NumLetra
            .TextMatrix(r, COL_FECHAEMI) = pckCHP.FechaEmision
            .TextMatrix(r, COL_FECHAVENCI) = pckCHP.FechaVenci
            pckCHP.Orden = NumPagos

'            mobjGNComp.DiaPago = DatePart("d", pckCHP.FechaEmision)
'
'            If Not BandDias Then
'
'                If mbooPorCobrar Then
'                    If mobjGNComp.DiaPago <> DatePart("d", pckCHP.FechaEmision) Then
'                        .TextMatrix(r, COL_FECHAVENCI) = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pckCHP.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pckCHP.FechaEmision))
'
'                        pckCHP.FechaVenci = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pckCHP.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pckCHP.FechaEmision))
'                    Else
'                        If mobjGNComp.GNTrans.codPantalla = "TSIER" Then
'                            .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, mobjGNComp.FechaPrimerPago)
'                            pckCHP.FechaVenci = DateAdd("m", NumPagos - 1 + mobjGNComp.MesesGracia, mobjGNComp.FechaPrimerPago)
'                        Else
'                            .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pckCHP.FechaEmision)
'                            pckCHP.FechaVenci = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pckCHP.FechaEmision)
'                        End If
'                    End If
'                Else
'                        .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pckCHP.FechaEmision)
'                        pckCHP.FechaVenci = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pckCHP.FechaEmision)
'                End If
'            Else
'
'                If mbooPorCobrar Then
'                    If mobjGNComp.DiaPago <> DatePart("d", pckCHP.FechaEmision) Then
'                        .TextMatrix(r, COL_FECHAVENCI) = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pckCHP.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pckCHP.FechaEmision))
'                        pckCHP.FechaVenci = mobjGNComp.DiaPago & "/" & DatePart("m", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pckCHP.FechaEmision)) & "/" & DatePart("yyyy", DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pckCHP.FechaEmision))
'                    Else
'                        If mobjGNComp.GNTrans.codPantalla = "TSIER" Then
'                            .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, mobjGNComp.FechaPrimerPago)
'                            pckCHP.FechaVenci = DateAdd("m", NumPagos - 1 + mobjGNComp.MesesGracia, mobjGNComp.FechaPrimerPago)
'                        Else
'                            .TextMatrix(r, COL_FECHAVENCI) = DateAdd("d", (NumPagos * intervalo) + mobjGNComp.MesesGracia, pckCHP.FechaEmision)
'                            pckCHP.FechaVenci = DateAdd("d", (NumPagos * intervalo) + mobjGNComp.MesesGracia, pckCHP.FechaEmision)
'                        End If
'                    End If
'                Else
'                        .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pckCHP.FechaEmision)
'                        pckCHP.FechaVenci = DateAdd("m", NumPagos + mobjGNComp.MesesGracia, pckCHP.FechaEmision)
'                End If
'
'            End If
            
            .TextMatrix(r, COL_PLAZO) = pckCHP.FechaVenci - pckCHP.FechaEmision
            
'        If mobjGNComp.GNTrans.IVPideVendedor And Len(mobjGNComp.CodVendedor) > 0 Then
'            Set Ven = mobjGNComp.Empresa.RecuperaFCVendedor(mobjGNComp.CodVendedor)
'            .TextMatrix(r, COL_CODVEN) = Ven.CodVendedor
'            .TextMatrix(r, COL_VENDEDOR) = Ven.nombre
'            pckCHP.IdVendedor = GNComprobante.IdVendedor
'        End If
            
            
            If valorCuotasGeneradas < 0 Then
                NumPagos = NumeroPagos
            End If
            
            
        .Cell(flexcpBackColor, r, COL_FECHAEMI) = grd.BackColorFixed
        .Cell(flexcpBackColor, r, COL_FECHAEMI) = grd.BackColorFixed

            .Row = r
            If .Rows > .FixedRows Then
                .col = .FixedCols
                'Busca la primera columna
                For i = .FixedCols To .Cols - 1
                    If .ColData(i) >= 0 And (Not .ColHidden(i)) And .ColWidth(i) > 0 Then
                        .col = i
                        Exit For
                    End If
                Next i
            End If
            
        End With
    
        PoneNumFila
        VisualizaTotal
    Next NumPagos
    'If mobjGNComp.Empresa.GNOpcion.ObtenerValor("BloquearCuotas") = "1" Then
    If mobjGNComp.GNTrans.ivBloquearCuotas Then
        For i = COL_NUMLETRA To COL_VENDEDOR
            grd.ColData(i) = -1
        Next i
        HabilitarCtrlsGrupoCajaFormaCobro False
     End If
    
salida:
    Set pckCHP = Nothing
'    grd.SetFocus
    Exit Sub
ErrTrap:
    Set pckCHP = Nothing
    DispErr
    GoTo salida
End Sub

