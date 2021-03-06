VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.UserControl PCDoc 
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
Attribute VB_Name = "PCDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'Ubicaci�n de columnas
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

Private WithEvents mobjGNComp As GNComprobante
Attribute mobjGNComp.VB_VarHelpID = -1
Private mobjGNTrans As GNTrans 'AUC 19/10/2005
Attribute mobjGNTrans.VB_VarHelpID = -1
Private mbooPorCobrar As Boolean
Private mbooModoProveedor As Boolean
Private mbooSI As Boolean
Private mbooProvCliVisible As Boolean
Private mstrCodProvCli As String
Private mstrCodProvCli2 As String           'Agregado para usar cuando se grabe los pcKardex y no asigne todo al un solo cliente o proveedor
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
Private mIDPcGrupo4 As Long
Private mCodforma As String

Private Sub ConfigCols()
    With grd
        .FormatString = "^#|<C�digo|<Nombre|<Forma|<#Doc|>Valor" & _
                        "|<F.Emisi�n|>Plazo|<F.Venci.|<Observaci�n|<CodVend.|<Vendedor"
                    'AUC 01/06/07 agregado codven  y vendedor
        .ColWidth(COL_NUMFILA) = 500
        .ColWidth(COL_CODPROVCLI) = 1200                'Cod.Proveedor/Cliente
        .ColWidth(COL_PROVCLI) = 1800                   'Proveedor/Cliente
        .ColWidth(COL_CODFORMA) = 1200                   'CodForma
        .ColWidth(COL_NUMLETRA) = 800                   'NumLetra
        .ColWidth(COL_VALOR) = COLANCHO_CUR             'valor
        .ColWidth(COL_FECHAEMI) = COLANCHO_FECHA        'F.Emisi�n
        .ColWidth(COL_PLAZO) = 1200  'antes 600                     'Plazo
        .ColWidth(COL_FECHAVENCI) = COLANCHO_FECHA      'F.Vencimiento
        .ColWidth(COL_OBSERVA) = 2000               'Observaci�n
        'AUC 01/06/07
        .ColWidth(COL_CODVEN) = 1200                'Cod.vendedor
        .ColWidth(COL_VENDEDOR) = 1800               'vendedor
        
        .ColHidden(COL_CODPROVCLI) = Not mbooProvCliVisible
        .ColHidden(COL_PROVCLI) = Not mbooProvCliVisible
        
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
        ElseIf mobjGNComp.GNTrans.Modulo = "TS" Then
            If Not mobjGNComp.GNTrans.TSPideCobrador Then
                .ColHidden(COL_CODVEN) = True
                .ColHidden(COL_VENDEDOR) = True
            End If
        ElseIf mobjGNComp.GNTrans.Modulo = "AF" Then
            .ColHidden(COL_CODVEN) = True
            .ColHidden(COL_VENDEDOR) = True
        End If
        
        If .Rows > .FixedRows Then .Row = .FixedRows
        If .Rows > .FixedRows Then .col = COL_CODFORMA
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
        If mbooProvCliVisible Then    'Agregado Oliver 17/dic/2003 solo es nececsario cuando esta la columna de Nombre visible
            '*** MAKOTO 18/jul/00 Agregado. Para permitir seleccionar por Nombre de prov/cli
        '    grd.ColComboList(COL_PROVCLI) = mobjGNComp.Empresa.ListaPCProvCliParaFlex2(mbooModoProveedor)
            grd.ColComboList(COL_CODPROVCLI) = mobjGNComp.Empresa.ListaPCProvCliParaFlexEx(True, True)
            grd.ColComboList(COL_PROVCLI) = mobjGNComp.Empresa.ListaPCProvCliParaFlex2Ex(True, True)        '*** MAKOTO 23/oct/00
        End If
            'AUC 01/06/007
            grd.ColComboList(COL_CODVEN) = mobjGNComp.Empresa.ListaFCVendedorParaFlex
            grd.ColComboList(COL_VENDEDOR) = mobjGNComp.Empresa.ListaFCVendedorParaFlex2

    ConfigColsFormato       'Llama esta para actualizar formato de moneda
    End If   'Modificado Oliver 17/dic/2003 porque no nesetita cargar la lista si esta en modo ver
    
End Sub

Private Sub grd_Click()
    RaiseEvent Click
End Sub

Private Sub grd_DblClick()
    RaiseEvent DblClick
End Sub

'*** MAKOTO 12/dic/00 Agregado
Private Sub grd_GotFocus()
    Dim Cancel As Boolean
    FlexGridGotFocusColor grd
    
    If grd.Editable And grd.Rows <= grd.FixedRows Then
        RaiseEvent AgregarFilaAuto(Cancel)  'Pregunta al contenedor si permite agregar la primera fila autom�ticamente o no
        If Not Cancel Then
            AgregaFila       'Si dice que s� }, agrega la primera fila
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
    Dim obj As PCKardex, cod As String
    Dim tsf As TSFormaCobroPago, ix As Long
    On Error GoTo ErrTrap

    If Not IsObject(grd.RowData(Row)) Then Exit Sub
    With grd
        Set obj = .RowData(Row)
        Select Case col
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
                    If Trim$(.Text) <> mobjGNComp.CodFormnaCP Then
                        MsgBox " Forma de Pago Diferente a la seleccionada anteriormente"
                        .TextMatrix(Row, col) = mobjGNComp.CodFormnaCP
                        obj.codforma = mobjGNComp.CodFormnaCP
                    End If
                    
                End If
            Else
                obj.codforma = Trim$(.Text)
            End If
''            Set tsf = mobjGNComp.Empresa.RecuperaTSFormaCobroPago(obj.codForma)
''            If Not tsf Is Nothing Then
''                    If tsf.IngresoAutomatico Then
''                        RaiseEvent AgregarFilaDocs1(True, obj.codForma, obj.Debe)
''                    End If
''            End If
''            Set tsf = Nothing
            VisualizarPlazo obj, Row
        Case COL_NUMLETRA
            obj.NumLetra = Trim$(.Text)
        Case COL_VALOR
            '*** Asignamos en ValidateEdit para verificar y cancelar si es necesario
            '*** MAKOTO 29/ene/01 Mod.
            '*** Sin embargo hay que asignar de nuevo para que guarde con el valor redondeado
            Dim CodTrans As String, numtrans As Long, valor As Currency
            For ix = 1 To mobjGNComp.CountPCKardex
                    If mobjGNComp.PCKardex(ix).id <> 0 Then
                    If mobjGNComp.Empresa.VerificarCambioCobroPago(mobjGNComp.PCKardex(ix).id, CodTrans, numtrans, valor) Then
                        If valor <> grd.ValueMatrix(ix, COL_VALOR) Then
                            grd.TextMatrix(ix, COL_VALOR) = mobjGNComp.PCKardex(ix).Debe
                            Err.Raise ERR_INVALIDO, "GNComprobante.RemovePCKardex", _
                            "No se puede Modificar el documento debido a que existen cobros o pagos asignados " & Chr(13) & "con la Transacci�n: " & CodTrans & "-" & numtrans & "del Cliente: " & mobjGNComp.PCKardex(ix).CodProvCli
                        End If
                        'Exit For
                    End If
                End If
                'If mcolPCKardex.item(ix) Is obj Then Exit For
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
        Case COL_FECHAVENCI
            .TextMatrix(Row, COL_PLAZO) = CDate(.Text) - CDate(.TextMatrix(Row, COL_FECHAEMI))
            obj.FechaVenci = CDate(.Text)
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
    
End Sub


'*** MAKOTO 18/jul/00
'Para coger solo c�digo de prov/cli cuando selecciona nombre
''xxxxxxxxxxxxxxxxx [nnnnnnn]'    --> Devuelve solo 'nnnnnnn'
Private Function CogeSoloCodigo(ByVal Desc As String) As String
    Dim s As String, i As Long
    i = InStrRev(Desc, "[")
    If i > 0 Then s = Mid$(Desc, i + 1)
    If Len(s) > 0 Then s = Left$(s, Len(s) - 1)
    CogeSoloCodigo = s
End Function


Private Sub VisualizaTotal()
    grd.SubTotal flexSTSum, -1, COL_VALOR, , grd.BackColorFrozen, vbYellow, , "Total", , True
    grd.Refresh
End Sub

Private Sub MueveColumna()
    Dim c As Long
    With grd
        If .Rows > .FixedRows Then
            For c = .col + 1 To .Cols - 1
                If .ColData(c) >= 0 And .ColWidth(c) > 0 And (Not .ColHidden(c)) Then
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
    If mobjGNComp.SoloVer Then Exit Sub
    
    'Cuando es una columna no modificable
    If grd.Rows > grd.FixedRows Then
        Cancel = (grd.ColData(col) < 0) Or grd.IsSubtotal(Row) Or grd.ColHidden(col)
    Else
        Cancel = True
    End If
    
    If grd.Cell(flexcpData, Row, col) = -1 Then Cancel = True
    If Cancel Then Exit Sub
    
    Select Case col
    Case COL_FECHAEMI
        If mobjGNComp.CodTrans <> "CLND" And mobjGNComp.CodTrans <> "CLNC" And mobjGNComp.CodTrans <> "PVND" And mobjGNComp.CodTrans <> "PVNC" Then
             Cancel = True
        End If
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
    Dim pck As PCKardex
    Dim Ven As FCVendedor
    RaiseEvent PorAgregarFila(v)        'Para calcular el valor predeterminado

    On Error GoTo ErrTrap
    
    'Llama a agregar un objeto PCKardex antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
    ix = mobjGNComp.AddPCKardex
    
    With grd
        r2 = .Rows - 1
        If .IsSubtotal(.Rows - 1) Then r2 = r2 - 1
        'Si no es la primera fila
        If r2 > 0 Then
            'Si no est� en la fila de total
            If Not .IsSubtotal(.Row) Then
                .AddItem "", .Row + 1
                r = .Row + 1
            'Si est� en la fila de total
            Else
                .AddItem "", .Row
                r = .Row
            End If
        'Si es la primera fila
        Else
            'Si no est� en la fila de total
            If (.Row < .Rows - 1) Or (.Row = 0) Then
'            If Not .IsSubtotal(.Row) Then
                .AddItem ""
                r = .Rows - 1
            'Si est� en la fila de total
            Else
                .AddItem "", .Row
                r = .Row
            End If
        End If

        'Asigna la referencia al nuevo objeto a la fila nueva
        Set pck = mobjGNComp.PCKardex(ix)
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

Private Sub EliminaFila()
    Dim msg As String, r As Long
    On Error GoTo ErrTrap

    If grd.Row < grd.FixedRows Then Exit Sub        '*** MAKOTO 07/feb/01 Mod.
    If grd.Rows <= grd.FixedRows Then Exit Sub
    If grd.IsSubtotal(grd.Row) Then Exit Sub
    
    r = grd.Row
    msg = "Desea eliminar la fila #" & r & "?"
    If MsgBox(msg, vbYesNo + vbQuestion) <> vbYes Then Exit Sub

    'Remueve de la colecci�n de objeto
    mobjGNComp.RemovePCKardex 0, grd.RowData(r)
    
    'Elimina del grid
    grd.RemoveItem r
    PoneNumFila
    grd.SubTotal flexSTClear
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
    Dim f1 As Date, f2 As Date, obj As PCKardex
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
                    MsgBox "Ingrese un valor num�rico.", vbExclamation
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
                        MsgBox "La fecha de vencimiento no puede ser menor a la fecha de emisi�n.", vbExclamation
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
        If mobjGNComp.GNTrans.CodPantalla <> "TSIEE" Then
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
'    'Visualiza los detalles que est� en GNComprobante
'    Set grd.DataSource = mobjGNComp.ListaPCKardex
'    ConfigCols
'
'    'Asigna referencia al objeto PCKardex a cada fila de grid
'    With grd
'        For i = 1 To mobjGNComp.CountPCKardex
'            .RowData(i) = mobjGNComp.PCKardex(i)
'        Next i
'    End With
'
'    PoneNumFila
'    VisualizaTotal
'    grd.Redraw = True
'    grd.Refresh
'End Sub

Public Sub Aceptar()
    Dim i As Long, obj As PCKardex
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
                obj.orden = i
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
    Dim i As Long, pck As PCKardex
    On Error GoTo ErrTrap
    
    If Len(value) > 0 And _
        (Not mobjGNComp.GNTrans.IVProvCliPorFila) Then  '*** MAKOTO 12/oct/00 Modificado
        With grd
            'Asigna el codigo de prov/cli a todos los detalles de PCKardex
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
    Dim r As Long, pck As PCKardex
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
    Dim pck As PCKardex, pc As PCProvCli, i As Long, s As String
    Dim r As Long, NombrePC As String, Ven As FCVendedor, nombreVen As String
    With grd
        .Redraw = False
        Limpiar
        For i = 1 To mobjGNComp.CountPCKardex
            Set pck = mobjGNComp.PCKardex(i)
            If Len(mstrCodProvCli) = 0 Then   'Agregado esta condicion para que cuando se recupera carge a la variable de la propiedad Docs.codProvCli
                'Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(pck.CodProvCli)
                'If Not (pc Is Nothing) Then
                mstrCodProvCli = pck.CodProvCli
                'Set pc = Nothing
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

'Cargar valores de propiedad desde el almac�n
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    ProvCliVisible = PropBag.ReadProperty("ProvCliVisible", True)
    PorCobrar = PropBag.ReadProperty("PorCobrar", True)
End Sub

Private Sub UserControl_Resize()
    'Ajusta el tama�o del grid
    grd.Height = UserControl.ScaleHeight
End Sub

Private Sub UserControl_Terminate()
    Set mobjGNComp = Nothing
End Sub

'Escribir valores de propiedad en el almac�n
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Enabled", UserControl.Enabled, True
    PropBag.WriteProperty "ProvCliVisible", ProvCliVisible
    PropBag.WriteProperty "PorCobrar", PorCobrar
End Sub

Private Sub VisualizarPlazo(ByRef obj As PCKardex, Row As Long)
    Dim tsf As TSFormaCobroPago
    With grd
        'cambiar plazo que tiene esta forma de cobro
        Set tsf = mobjGNComp.Empresa.RecuperaTSFormaCobroPago(obj.codforma)
        If Not (tsf Is Nothing) Then
            obj.FechaVenci = obj.FechaEmision + tsf.Plazo
        Else
            obj.FechaVenci = obj.FechaEmision
        End If
        Set tsf = Nothing
        'Visualizando los cambios en las fechas
        .TextMatrix(Row, COL_PLAZO) = obj.FechaVenci - obj.FechaEmision
        .TextMatrix(Row, COL_FECHAVENCI) = obj.FechaVenci
    End With
    BloquearPlazo obj, Row
End Sub

'Agregado  para controlar si el plazo permite modificacion o no
Private Sub BloquearPlazo(ByRef obj As PCKardex, Row As Long)
    Dim tsf As TSFormaCobroPago
    With grd
        'cambiar plazo que tiene esta forma de cobro
        Set tsf = mobjGNComp.Empresa.RecuperaTSFormaCobroPago(obj.codforma)
        
        If mobjGNComp.CodTrans <> "CLND" And mobjGNComp.CodTrans <> "CLNC" And mobjGNComp.CodTrans <> "PVND" And mobjGNComp.CodTrans <> "PVNC" Then
            .Cell(flexcpData, Row, COL_FECHAEMI) = -1
            .Cell(flexcpBackColor, Row, COL_FECHAEMI) = grd.BackColorFixed
            .Cell(flexcpBackColor, Row, COL_FECHAEMI) = grd.BackColorFixed
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
'            RaiseEvent AgregarFilaAuto(Cancel)  'Pregunta al contenedor si permite agregar la primera fila autom�ticamente o no
'            If Not Cancel Then
'                AgregaFila       'Si dice que s� }, agrega la primera fila
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
    Dim pck As PCKardex
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
'        Llama a agregar un objeto PCKardex antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
            ix = mobjGNComp.AddPCKardex
            
            With grd
                r2 = .Rows - 1
                If .IsSubtotal(.Rows - 1) Then r2 = r2 - 1
                'Si no es la primera fila
                If r2 > 0 Then
                    'Si no est� en la fila de total
                    If Not .IsSubtotal(.Row) Then
                        .AddItem "", .Row + 1
                        r = .Row + 1
                    'Si est� en la fila de total
                    Else
                        .AddItem "", .Row
                        r = .Row
                    End If
                'Si es la primera fila
                Else
                    'Si no est� en la fila de total
                    If (.Row < .Rows - 1) Or (.Row = 0) Then
        '            If Not .IsSubtotal(.Row) Then
                        .AddItem ""
                        r = .Rows - 1
                    'Si est� en la fila de total
                    Else
                        .AddItem "", .Row
                        r = .Row
                    End If
                End If
    
                'Asigna la referencia al nuevo objeto a la fila nueva
                Set pck = mobjGNComp.PCKardex(ix)
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


Private Sub VisualizarPlazoMasDias(ByRef obj As PCKardex, Row As Long, Optional dias As Long)
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
    Dim pck As PCKardex

    RaiseEvent PorAgregarFilaconPagoInicial(v)        'Para calcular el valor predeterminado

    On Error GoTo ErrTrap
    
    'Llama a agregar un objeto PCKardex antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
    ix = mobjGNComp.AddPCKardex
    
    With grd
        r2 = .Rows - 1
        If .IsSubtotal(.Rows - 1) Then r2 = r2 - 1
        'Si no es la primera fila
        If r2 > 0 Then
            'Si no est� en la fila de total
            If Not .IsSubtotal(.Row) Then
                .AddItem "", .Row + 1
                r = .Row + 1
            'Si est� en la fila de total
            Else
                .AddItem "", .Row
                r = .Row
            End If
        'Si es la primera fila
        Else
            'Si no est� en la fila de total
            If (.Row < .Rows - 1) Or (.Row = 0) Then
'            If Not .IsSubtotal(.Row) Then
                .AddItem ""
                r = .Rows - 1
            'Si est� en la fila de total
            Else
                .AddItem "", .Row
                r = .Row
            End If
        End If

        'Asigna la referencia al nuevo objeto a la fila nueva
        Set pck = mobjGNComp.PCKardex(ix)
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
            .TextMatrix(r, COL_CODFORMA) = pck.codforma
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
    Dim pck As PCKardex, i As Long, msg As String
On Error GoTo ErrTrap
    msg = "Autom�ticamente cambiar� la coumna #Doc " & Chr(13) & "es decir, el N�mero de Referencia de cada fila "
    msg = msg & vbCr & vbCr & "Desea continuar?"
    If MsgBox(msg, vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    
    MensajeStatus MSG_PREPARA, vbHourglass
    
    For i = 1 To mobjGNComp.CountPCKardex
        Set pck = mobjGNComp.PCKardex(i)
        MensajeStatus "Proces�ndo #" & i & ": '" & pck.NumLetra & "'...", vbHourglass
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
            grd.ColComboList(COL_CODFORMA) = mobjGNComp.Empresa.ListaTSFormaCobroPagoParaFlexContolFormas(mbooPorCobrar, NumGrupoControl)
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
    Dim pck As PCKardex

    RaiseEvent PorAgregarFila(v)        'Para calcular el valor predeterminado

    On Error GoTo ErrTrap
    
    'Llama a agregar un objeto PCKardex antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
    ix = mobjGNComp.AddPCKardex
    
    With grd
        r2 = .Rows - 1
        If .IsSubtotal(.Rows - 1) Then r2 = r2 - 1
        'Si no es la primera fila
        If r2 > 0 Then
            'Si no est� en la fila de total
            If Not .IsSubtotal(.Row) Then
                .AddItem "", .Row + 1
                r = .Row + 1
            'Si est� en la fila de total
            Else
                .AddItem "", .Row
                r = .Row
            End If
        'Si es la primera fila
        Else
            'Si no est� en la fila de total
            If (.Row < .Rows - 1) Or (.Row = 0) Then
'            If Not .IsSubtotal(.Row) Then
                .AddItem ""
                r = .Rows - 1
            'Si est� en la fila de total
            Else
                .AddItem "", .Row
                r = .Row
            End If
        End If

        'Asigna la referencia al nuevo objeto a la fila nueva
        Set pck = mobjGNComp.PCKardex(ix)
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
    Dim pck As PCKardex
    Dim Ven As FCVendedor
    RaiseEvent PorAgregarFila(v)        'Para calcular el valor predeterminado
    On Error GoTo ErrTrap
    'Llama a agregar un objeto PCKardex antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
    ix = mobjGNComp.AddPCKardex
    With grd
        r2 = .Rows - 1
        If .IsSubtotal(.Rows - 1) Then r2 = r2 - 1
        'Si no es la primera fila
        If r2 > 0 Then
            'Si no est� en la fila de total
            If Not .IsSubtotal(.Row) Then
                .AddItem "", .Row + 1
                r = .Row + 1
            'Si est� en la fila de total
            Else
                .AddItem "", .Row
                r = .Row
            End If
        'Si es la primera fila
        Else
            'Si no est� en la fila de total
            If (.Row < .Rows - 1) Or (.Row = 0) Then
'            If Not .IsSubtotal(.Row) Then
                .AddItem ""
                r = .Rows - 1
            'Si est� en la fila de total
            Else
                .AddItem "", .Row
                r = .Row
            End If
        End If
        'Asigna la referencia al nuevo objeto a la fila nueva
        Set pck = mobjGNComp.PCKardex(ix)
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
    Dim pck As PCKardex, pc As PCProvCli, i As Long, s As String
    Dim r As Long, NombrePC As String, Ven As FCVendedor, nombreVen As String
    Dim tsf As TSFormaCobroPago
    With grd
        .Redraw = False
        Limpiar
        For i = 1 To mobjGNComp.CountPCKardex
            Set pck = mobjGNComp.PCKardex(i)
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
'                                s = .Rows & vbTab & _
                                pck.CodProvCli & vbTab & _
                                NombrePC & vbTab & _
                                pck.codforma & vbTab & _
                                pck.NumLetra & vbTab & _
                                NullSiZero(pck.Debe + pck.Haber) & vbTab & _
                                pck.FechaEmision & vbTab & _
                                pck.FechaVenci - pck.FechaEmision & vbTab & _
                                pck.FechaVenci & vbTab & _
                                pck.Observacion & vbTab & pck.CodVendedor & vbTab & nombreVen
'                                .AddItem s
'                                r = .Rows - 1
'                                .RowData(r) = pck
'                                '*** Oliver 29/01/2003 Agregado para que bloquee las columnas que no se modifiquen
'                                BloquearPlazo pck, r
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
    Dim pck As PCKardex
    Dim tsf As TSFormaCobroPago
    Dim pc As PCProvCli, obser As String
    Dim GNCompAux As GNComprobante
    RaiseEvent PorAgregarFila(v)        'Para calcular el valor predeterminado

    On Error GoTo ErrTrap
    
    'Llama a agregar un objeto PCKardex antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
    ix = mobjGNComp.AddPCKardex
    
    With grd
        r2 = .Rows - 1
        If .IsSubtotal(.Rows - 1) Then r2 = r2 - 1
        'Si no es la primera fila
        If r2 > 0 Then
            'Si no est� en la fila de total
            If Not .IsSubtotal(.Row) Then
                .AddItem "", .Row + 1
                r = .Row + 1
            'Si est� en la fila de total
            Else
                .AddItem "", .Row
                r = .Row
            End If
        'Si es la primera fila
        Else
            'Si no est� en la fila de total
            If (.Row < .Rows - 1) Or (.Row = 0) Then
'            If Not .IsSubtotal(.Row) Then
                .AddItem ""
                r = .Rows - 1
            'Si est� en la fila de total
            Else
                .AddItem "", .Row
                r = .Row
            End If
        End If

        'Asigna la referencia al nuevo objeto a la fila nueva
        Set pck = mobjGNComp.PCKardex(ix)
        .RowData(r) = pck
        
        
        Set tsf = mobjGNComp.Empresa.RecuperaTSFormaCobroPago(codforma)
        If tsf.DeudaMismoCliente Then
            Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(mobjGNComp.CodClienteRef)
        Else
            Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(tsf.CodProvCli)
        End If
        Set GNCompAux = mobjGNComp.Empresa.RecuperaGNComprobante(mobjGNComp.IdTransFuente)

        
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

Public Sub AgregaFilaPagoInicial(ByVal valor As Double)
    Dim r As Long, r2 As Long, ix As Long, i As Long, ValorIni As Currency, saldo As Currency
    Dim pck As PCKardex

'    RaiseEvent ValorPagoInicial(ValorIni)
'    RaiseEvent PorAgregarFilaconPagoInicial(Saldo)        'Para calcular el valor predeterminado

    On Error GoTo ErrTrap
    EliminaFilaTodoDocs
    ValorIni = valor
    If ValorIni = 0 Then Exit Sub
    
    'Llama a agregar un objeto PCKardex antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
    ix = mobjGNComp.AddPCKardex
    
    With grd
        r2 = .Rows - 1
        If .IsSubtotal(.Rows - 1) Then r2 = r2 - 1
        'Si no es la primera fila
        If r2 > 0 Then
            'Si no est� en la fila de total
            If Not .IsSubtotal(.Row) Then
                .AddItem "", .Row + 1
                r = .Row + 1
            'Si est� en la fila de total
            Else
                .AddItem "", .Row
                r = .Row
            End If
        'Si es la primera fila
        Else
            'Si no est� en la fila de total
            If (.Row < .Rows - 1) Or (.Row = 0) Then
'            If Not .IsSubtotal(.Row) Then
                .AddItem ""
                r = .Rows - 1
            'Si est� en la fila de total
            Else
                .AddItem "", .Row
                r = .Row
            End If
        End If

        'Asigna la referencia al nuevo objeto a la fila nueva
        Set pck = mobjGNComp.PCKardex(ix)
        .RowData(r) = pck
        
        'Proporciona el valor predeterminado        '*** MAKOTO 05/oct/00 Modificado
'        If v > 0 Then
            If mbooPorCobrar Then pck.Debe = Abs(ValorIni)
            .TextMatrix(r, COL_VALOR) = ValorIni
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
            .TextMatrix(r, COL_CODFORMA) = pck.codforma
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

Public Sub AgregaFilasPorCredito(ByVal valor As Double)
    Dim r As Long, r2 As Long, ix As Long, i As Long, ValorIni As Currency, saldo As Currency
    Dim pck As PCKardex
    Dim NumPagos As Integer, ValorCuota As Currency, CuotaFinal As Currency, TotalCuota As Currency
'    RaiseEvent PorAgregarFilaconPagoInicial(Saldo)        'Para calcular el valor predeterminado
    On Error GoTo ErrTrap
    If valor <= 0 Then Exit Sub
    TotalCuota = 0
    For NumPagos = 1 To mobjGNComp.NumDias
        ValorCuota = Round(valor / mobjGNComp.NumDias, 2)
        'Llama a agregar un objeto PCKardex antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
        ix = mobjGNComp.AddPCKardex
        
        With grd
            r2 = .Rows - 1
            If .IsSubtotal(.Rows - 1) Then r2 = r2 - 1
            'Si no es la primera fila
            If r2 > 0 Then
                'Si no est� en la fila de total
                If Not .IsSubtotal(.Row) Then
                    .AddItem "", .Row + 1
                    r = .Row + 1
                'Si est� en la fila de total
                Else
                    .AddItem "", .Row
                    r = .Row
                End If
            'Si es la primera fila
            Else
                'Si no est� en la fila de total
                If (.Row < .Rows - 1) Or (.Row = 0) Then
    '            If Not .IsSubtotal(.Row) Then
                    .AddItem ""
                    r = .Rows - 1
                'Si est� en la fila de total
                Else
                    .AddItem "", .Row
                    r = .Row
                End If
            End If
    
            'Asigna la referencia al nuevo objeto a la fila nueva
            Set pck = mobjGNComp.PCKardex(ix)
            .RowData(r) = pck
            
            'Proporciona el valor predeterminado        '*** MAKOTO 05/oct/00 Modificado
            If NumPagos = mobjGNComp.NumDias Then
                ValorCuota = Round(valor - TotalCuota, 2)
            End If
                If mbooPorCobrar Then pck.Debe = Abs(ValorCuota)
                .TextMatrix(r, COL_VALOR) = ValorCuota
                TotalCuota = TotalCuota + ValorCuota
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
            'mobjGNComp.codtrans & "-" & mobjGNComp.GNTrans.NumTransSiguiente &
            pck.NumLetra = " Cuota " & NumPagos & "/" & mobjGNComp.NumDias
            .TextMatrix(r, COL_NUMLETRA) = pck.NumLetra
            .TextMatrix(r, COL_FECHAEMI) = pck.FechaEmision
            .TextMatrix(r, COL_FECHAVENCI) = DateAdd("m", NumPagos, pck.FechaEmision)
            pck.FechaVenci = DateAdd("m", NumPagos, pck.FechaEmision)
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
    Next NumPagos
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
    Dim pck As PCKardex
    Dim Ven As FCVendedor
    RaiseEvent PorAgregarFila(v)        'Para calcular el valor predeterminado

    On Error GoTo ErrTrap
    
    'Llama a agregar un objeto PCKardex antes de agregar la fila    '*** MAKOTO 14/oct/00 Modificado
    ix = mobjGNComp.AddPCKardex
    
    With grd
        r2 = .Rows - 1
        If .IsSubtotal(.Rows - 1) Then r2 = r2 - 1
        'Si no es la primera fila
        If r2 > 0 Then
            'Si no est� en la fila de total
            If Not .IsSubtotal(.Row) Then
                .AddItem "", .Row + 1
                r = .Row + 1
            'Si est� en la fila de total
            Else
                .AddItem "", .Row
                r = .Row
            End If
        'Si es la primera fila
        Else
            'Si no est� en la fila de total
            If (.Row < .Rows - 1) Or (.Row = 0) Then
'            If Not .IsSubtotal(.Row) Then
                .AddItem ""
                r = .Rows - 1
            'Si est� en la fila de total
            Else
                .AddItem "", .Row
                r = .Row
            End If
        End If

        'Asigna la referencia al nuevo objeto a la fila nueva
        Set pck = mobjGNComp.PCKardex(ix)
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

