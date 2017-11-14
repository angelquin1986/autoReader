VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmB_TransConsol 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8415
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5610
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   400
      Left            =   2580
      TabIndex        =   1
      Top             =   7920
      Width           =   1200
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar -F5"
      Height          =   400
      Left            =   1320
      TabIndex        =   0
      Top             =   7920
      Width           =   1200
   End
   Begin VB.Frame fraCobro 
      Caption         =   "Códigos de Retenciones de IVA"
      Height          =   1695
      Left            =   60
      TabIndex        =   2
      Top             =   6120
      Width           =   5475
      Begin VB.CommandButton cmdResta 
         Caption         =   "&<<"
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&>>"
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   600
         Width           =   615
      End
      Begin VB.ListBox lstServicios 
         Height          =   1035
         Left            =   3180
         TabIndex        =   4
         Top             =   510
         Width           =   2175
      End
      Begin VB.ListBox lstBienes 
         Height          =   1035
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   2115
      End
      Begin VB.Label Label3 
         Caption         =   "BIENES"
         Height          =   255
         Left            =   105
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "SERVICIOS"
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdTrans 
      Height          =   3600
      Left            =   60
      TabIndex        =   9
      Top             =   300
      Width           =   5475
      _cx             =   9657
      _cy             =   6350
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
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
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
      Editable        =   2
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
   Begin VSFlex7LCtl.VSFlexGrid grdTransRet 
      Height          =   1815
      Left            =   60
      TabIndex        =   10
      Top             =   4260
      Width           =   5475
      _cx             =   9657
      _cy             =   3201
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
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
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
      Editable        =   2
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
   Begin VB.Label Label2 
      Caption         =   "Transacciones de Retencion en Compras"
      Height          =   255
      Left            =   60
      TabIndex        =   12
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Transacciones de Compras"
      Height          =   255
      Left            =   60
      TabIndex        =   11
      Top             =   60
      Width           =   2535
   End
End
Attribute VB_Name = "frmB_TransConsol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BandAceptado As Boolean
Private fecha1 As Date
Private fecha2 As Date
Dim Key As String

'grdsucursal
Const COLS_SEL = 0
Const COLS_COD = 1
Const COLS_DES = 2
Const COLS_NOM = 3
'grdtrans
Const COLT_SEL = 0
Const COLT_SUC = 1
Const COLT_COD = 2
Const COLT_DES = 3
Const IVGRUPO_MAX = 5



Public Function Inicio(ByVal objSiiMain As SiiMain, cad As String, fecha As Date, ListaEmpresas As String) As Boolean
    Dim i As Integer, j As Integer, v As Variant, W As Variant
    Dim trans As String, KeyTrans As String, KeyTransRet As String, KeyRecargo As String, KeyAnulado As String
    Me.tag = Name  'nombre del reporte
    CargaRecargos
    

    If DatePart("m", fecha) = 12 Then
        fecha1 = "01/" & DatePart("m", fecha) & "/" & DatePart("yyyy", fecha)
        fecha2 = DateAdd("yyyy", 1, DateAdd("d", -1, ("01/" & DatePart("m", DateAdd("m", 1, fecha)) & "/" & DatePart("yyyy", fecha))))
    Else
        fecha1 = "01/" & DatePart("m", fecha) & "/" & DatePart("yyyy", fecha)
        fecha2 = DateAdd("d", -1, ("01/" & DatePart("m", DateAdd("m", 1, fecha)) & "/" & DatePart("yyyy", fecha)))
    End If
    ConfigCols
'    CargarListaEmpresas
    With objSiiMain.objCondicion
        .fecha1 = fecha1
        .fecha2 = fecha2
        BandAceptado = False
        Me.Caption = "Condiciones de Busqueda "
        Select Case cad
                
           Case "IMPCPI2016"
                grdTrans.Rows = 1
                v = Split(ListaEmpresas, ";")
                For i = 0 To UBound(v)
                    W = Split(v(i), ",")
                    If UBound(W) > 0 Then
                        CargarListaTransEmpresasSeleccionadas "IV", "1,2,3,4", "1", W(0), W(1), grdTrans
                        CargarListaTransEmpresasSeleccionadas "TS", "7", "1", W(0), W(1), grdTransRet
                    End If
                Next i
                
                KeyTrans = "ATS2016_Trans_CPI"
                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyTrans)
                RecuperaSelecTransGrd trans, grdTrans, trans, 2, 1

                KeyTransRet = "ATS2016_Trans_RTP"
                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyTransRet)
                RecuperaSelecTransGrd trans, grdTransRet, trans, 2, 1
                
                KeyRecargo = "ATS2016_Trans_REC"
                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyRecargo)
                RecuperaSelecRec KeyRecargo, trans
                
            Case "IMPFC2016"

                grdTrans.Rows = 1
                v = Split(ListaEmpresas, ";")
                For i = 0 To UBound(v)
                    W = Split(v(i), ",")
                    If UBound(W) > 0 Then
                        CargarListaTransEmpresasSeleccionadas "IV", "18,4", "2", W(0), W(1), grdTrans
                        CargarListaTransEmpresasSeleccionadas "TS", "7", "", W(0), W(1), grdTransRet
                    End If
                Next i
                
                KeyTrans = "ATS2016_Trans_FC"
                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyTrans)
                RecuperaSelecTransGrd trans, grdTrans, trans, 2, 1

                KeyTransRet = "ATS2016_Trans_RTC"
                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyTransRet)
                RecuperaSelecTransGrd trans, grdTransRet, trans, 2, 1
                
                KeyRecargo = "ATS2016_Trans_REC_FC"
                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyRecargo)
                RecuperaSelecRec KeyRecargo, trans

            Case "IMPFC2016xE"

                grdTransRet.Visible = False
                
                grdTrans.Height = 5700
                grdTrans.Rows = 1
                v = Split(ListaEmpresas, ";")
                For i = 0 To UBound(v)
                    W = Split(v(i), ",")
                    If UBound(W) > 0 Then
                        CargarListaTransEmpresasSeleccionadas "IV", " 18,4", "2", W(0), W(1), grdTrans
                    End If
                Next i
                
                KeyTrans = "ATS2016_Trans_FC"
                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyTrans)
                RecuperaSelecTransGrd trans, grdTrans, trans, 2, 1

                KeyRecargo = "ATS2016_Trans_REC_FC"
                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyRecargo)
                RecuperaSelecRec KeyRecargo, trans


            
            Case "IMPAN2016"
''''                CargaTipoTransxTipoTrans "IV", lst, "'1','2'"
''''                fraCobro.Visible = False
''''                fraRetencion.Visible = False
''''                fra.Height = 5800
''''                lst.Height = 5500
''''                KeyAnulado = Name & "_Trans" & "_" & objSiiMain.EmpresaActual.CodEmpresa & "_ANU"
''''                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyAnulado)
''''                'trans = GetSetting(APPNAME, App.Title, KeyAnulado)
''''                RecuperaSelec KeyAnulado, lst, trans

                grdTransRet.Visible = False
                fraCobro.Visible = False

                grdTrans.Height = 7400
                
                
                grdTrans.Rows = 1
                v = Split(ListaEmpresas, ";")
                For i = 0 To UBound(v)
                    W = Split(v(i), ",")
                    If UBound(W) > 0 Then
                        CargarListaTransEmpresasSeleccionadas "", "", "1,2", W(0), W(1), grdTrans
                    End If
                Next i
                
                
                
                KeyAnulado = "ATS2016_Trans_ANU"
                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyAnulado)
                RecuperaSelecTransGrd trans, grdTrans, trans, 2, 1



            Case "IMPDD"
'''''                CargaTipoTransxTipoTrans "", lst, "'2' "
'''''                'CargaTipoTrans "TS", lstRet
'''''                KeyTrans = Name & "_Trans" & "_" & objSiiMain.EmpresaActual.CodEmpresa & "_FCDD"
'''''                trans = GetSetting(APPNAME, App.Title, KeyTrans)
'''''                RecuperaSelec KeyTrans, lst, trans
'''''
'''''
'''''                fraRetencion.Visible = False
'''''                fra.Caption = "Transacciones de Venta"
'''''
'''''                fraCobro.Visible = False
'''''                lst.Height = 5600
'''''                fra.Height = 5900
            Case "IMPFCICE"
'''''                CargaTipoTransxTipoTrans "IV", lst, "'2'"
'''''                'CargaTipoTrans "TS", lstRet
'''''                KeyTrans = Name & "_Trans" & "_" & objSiiMain.EmpresaActual.CodEmpresa & "_FC"
'''''                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyTrans)
'''''                'trans = GetSetting(APPNAME, App.Title, KeyTrans)
'''''                gobjMain.EmpresaActual.GNOpcion.ObtenerValor (KeyTrans)
'''''                RecuperaSelec KeyTrans, lst, trans
'''''
''''''                KeyTransRet = Name & "_Trans" & "_" & objSiiMain.EmpresaActual.CodEmpresa & "_RTC"
''''''                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyTransRet)
''''''                'trans = GetSetting(APPNAME, App.Title, KeyTransRet)
''''''                RecuperaSelec KeyTransRet, lstRet, trans
'''''
'''''                fraRetencion.Visible = False
'''''                fra.Caption = ""
'''''
'''''                fraCobro.Visible = True
'''''                fraCobro.Caption = "Recargos Descuentos antes del IVA"
'''''                CargaRecargos
'''''                KeyRecargo = Name & "_Trans" & "_" & objSiiMain.EmpresaActual.CodEmpresa & "_REC_FC"
'''''                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyRecargo)
'''''                'trans = GetSetting(APPNAME, App.Title, KeyRecargo)
'''''                RecuperaSelecRec KeyRecargo, trans
'''''
                
        End Select
            
        BandAceptado = False
        Me.Show vbModal, frmMain
        If BandAceptado Then
            .fecha1 = fecha1
            .fecha2 = fecha2
        Select Case cad
            Case "IMPCPI2016"
                .CodTrans = ListaEmpresasyTrans(grdTrans)
                .CodRetencion1 = ListaEmpresasyTrans(grdTransRet)
                .Servicios = PreparaCadReTencion(lstServicios)    'servicios
                gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyTrans, .CodTrans
                gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyTransRet, .CodRetencion1
                gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyRecargo, .Servicios
            Case "IMPFC2016"
                .CodGrupo = ListaEmpresasyTrans(grdTrans)
                .CodRetencion2 = ListaEmpresasyTrans(grdTransRet)
                .Sucursal = PreparaCadReTencion(lstServicios)    'servicios
                gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyTrans, .CodGrupo
                gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyTransRet, .CodRetencion2
                gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyRecargo, .Sucursal
            Case "IMPFC2016xE"
                .CodGrupo = ListaEmpresasyTrans(grdTrans)
                .CodRetencion2 = ListaEmpresasyTrans(grdTransRet)
                .Sucursal = PreparaCadReTencion(lstServicios)    'servicios

            Case "IMPAN2016"
                    .CodPC2 = ListaEmpresasyTrans(grdTrans)
                    gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyAnulado, .CodPC2
'''''            Case "IMPDD"
'''''                .CodGrupo = PreparaCadena
'''''                SaveSetting APPNAME, App.Title, KeyTrans, .CodGrupo
'''''            Case "IMPFCICE"
'''''                .CodGrupo = PreparaCadena
'''''                '.CodRetencion2 = PreparaCadenaRet
'''''                .Sucursal = PreparaCadReTencion(lstServicios)    'servicios
'''''                'SaveSetting APPNAME, App.Title, KeyTrans, .CodGrupo
'''''                'SaveSetting APPNAME, App.Title, KeyTransRet, .CodRetencion2
'''''                'SaveSetting APPNAME, App.Title, KeyRecargo, .Sucursal
'''''                gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyTrans, .CodGrupo
'''''                'gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyTransRet, .CodRetencion2
'''''                gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyRecargo, .Sucursal
'''''                gobjMain.EmpresaActual.GNOpcion.GrabarGNOpcion2
                
            End Select
            gobjMain.EmpresaActual.GNOpcion.GrabarSoloGnOpcion2
        End If
    End With
    Inicio = BandAceptado
    Unload Me
End Function

Private Sub cmdAceptar_Click()
    Dim Cadena  As String
    If Key = "CarteraxCxV_Trans" Then
        'Verificar si se han seleccionado transacciones
        'en reprote  de Cartera por cobrara por Vendedor
'        Cadena = PreparaCadena
        If Cadena = "" Then
            MsgBox "Seleccione las transacciones  a incluir "
            'lst.SetFocus
            Exit Sub
        End If
    End If
    
    
    BandAceptado = True
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    BandAceptado = False
    Me.Hide
End Sub


Private Sub dtpDesde_LostFocus()
Dim valor As Long
    
''''    Select Case Month(dtpDesde.value)
''''    Case 1, 3, 5, 7, 8, 10, 12
''''        dtpHasta.value = CDate("31/" & Month(dtpDesde.value) & "/" & Year(dtpDesde.value))
''''    Case 4, 6, 9, 11
''''        dtpHasta.value = CDate("30/" & Month(dtpDesde.value) & "/" & Year(dtpDesde.value))
''''    Case 2
''''        valor = Year(dtpDesde.value) / 4
''''        If Int(valor) * 4 = Year(dtpDesde.value) Then
''''            dtpHasta.value = CDate("29/" & Month(dtpDesde.value) & "/" & Year(dtpDesde.value))
''''        Else
''''            dtpHasta.value = CDate("28/" & Month(dtpDesde.value) & "/" & Year(dtpDesde.value))
''''        End If
''''    End Select
End Sub

Private Sub CargaTrans(ListaEmpresas As String)
Dim i As Integer, v As Variant, W As Variant
            grdTrans.Rows = 1
            v = Split(ListaEmpresas, ";")
'            If UBound(v) > 0 Then
            For i = 0 To UBound(v)
                W = Split(v(i), ",")
                If UBound(W) > 0 Then
'                    CargarListaTransEmpresasSeleccionadas w(0), w(1), grdTrans
'                    CargarListaTransRetEmpresasSeleccionadas w(0), w(1), grdTransRet
                End If
            Next i
 '           End If


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

'''Private Function PreparaCadena() As String
'''    Dim Cadena As String, i As Integer
'''    Cadena = ""
'''    For i = 0 To lst.ListCount - 1
'''        If lst.Selected(i) Then
'''            If Cadena = "" Then
'''                Cadena = Left(lst.List(i), lst.ItemData(i))
'''            Else
'''                Cadena = Cadena & "," & _
'''                              Left(lst.List(i), lst.ItemData(i))
'''            End If
'''        End If
'''    Next i
'''    PreparaCadena = Cadena
'''End Function
'''
'''
'''Private Function PreparaCadenaRet() As String
'''    Dim Cadena As String, i As Integer
'''    Cadena = ""
'''    For i = 0 To lstRet.ListCount - 1
'''        If lstRet.Selected(i) Then
'''            If Cadena = "" Then
'''                Cadena = Left(lstRet.List(i), lstRet.ItemData(i))
'''            Else
'''                Cadena = Cadena & "," & _
'''                              Left(lstRet.List(i), lstRet.ItemData(i))
'''            End If
'''        End If
'''    Next i
'''    PreparaCadenaRet = Cadena
'''End Function
'''
'''




Private Sub Form_Load()
    Dim resto As Integer, valor As Long
    'Establece los rangos de Fecha  siempre  al rango
    'del año actual
''''    dtpDesde.value = CDate("01/" & Month(Date) & "/" & Year(Date))
    
''''    Select Case Month(dtpDesde.value)
''''    Case 1, 3, 5, 7, 8, 10, 12
''''        dtpHasta.value = CDate("31/" & Month(dtpDesde.value) & "/" & Year(dtpDesde.value))
''''    Case 4, 6, 9, 11
''''        dtpHasta.value = CDate("30/" & Month(dtpDesde.value) & "/" & Year(dtpDesde.value))
''''    Case 2
''''        valor = Year(dtpDesde.value)
''''        If Int(valor) * 4 = Year(dtpDesde.value) Then
''''            dtpHasta.value = CDate("29/" & Month(dtpDesde.value) & "/" & Year(dtpDesde.value))
''''        Else
''''            dtpHasta.value = CDate("28/" & Month(dtpDesde.value) & "/" & Year(dtpDesde.value))
''''        End If
''''
''''    End Select
''''    dtpHasta.value = CDate("30/12/" & Year(dtpDesde.value))
End Sub


Private Sub cmdAdd_Click()
    Dim i As Long, ix As Long
    On Error GoTo ErrTrap
    With lstBienes
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                'ix = mobjGrupo.AgregarUsuario(.List(i))
                ix = .ItemData(i)
                lstServicios.AddItem .List(i)
                lstServicios.ItemData(lstServicios.NewIndex) = ix
                .RemoveItem i
            End If
        Next i
    End With
    Exit Sub
ErrTrap:
    DispErr
End Sub

Private Sub cmdResta_Click()
    Dim i As Long, ix As Long
    On Error GoTo ErrTrap
    With lstServicios
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                ix = .ItemData(i)
                lstBienes.AddItem .List(i)
                lstBienes.ItemData(lstBienes.NewIndex) = ix
                .RemoveItem i
            End If
        Next i
    End With
    Exit Sub
ErrTrap:
    DispErr
End Sub


Private Sub lstBienes_DblClick()
    cmdAdd_Click
End Sub

Private Sub lstServicios_DblClick()
    cmdResta_Click
End Sub

Public Sub RecuperaSelecRetencion(ByVal Key As String, ByRef lstBienes As ListBox, ByRef lstServicios As ListBox)
Dim s As String, Vector As Variant, ix As Long
Dim i As Integer, j As Integer, Selec As Integer
    'Recupera seleccionados  del registro de windows
    s = GetSetting(APPNAME, App.Title, Key, "_VACIO_")
    If s <> "_VACIO_" Then
        Vector = Split(s, ",")
         Selec = UBound(Vector, 1)
         For i = 0 To Selec
            For j = lstBienes.ListCount - 1 To 0 Step -1
                If Vector(i) = Left(lstBienes.List(j), lstBienes.ItemData(j)) Then
                    'ix = mobjGrupo.AgregarUsuario(.List(i))
                    ix = lstBienes.ItemData(j)
                    lstServicios.AddItem lstBienes.List(j)
                    lstServicios.ItemData(lstServicios.NewIndex) = ix
                    lstBienes.RemoveItem j
                End If
            Next j
         Next i
    End If
End Sub

Public Function PreparaCadReTencion(lst As ListBox) As String
    Dim Cadena As String, i As Integer
    Cadena = ""
    For i = 0 To lst.ListCount - 1
        If Cadena = "" Then
            Cadena = Left(lst.List(i), lst.ItemData(i))
        Else
            Cadena = Cadena & "," & _
                          Left(lst.List(i), lst.ItemData(i))
        End If
    Next i
    PreparaCadReTencion = Cadena
End Function


'Public Sub RecuperaSelecRetencion(ByVal Key As String, ByRef lstBienes As ListBox, ByRef lstServicios As ListBox)
'Dim s As String, Vector As Variant, ix As Long
'Dim i As Integer, j As Integer, Selec As Integer
'    'Recupera seleccionados  del registro de windows
'    s = GetSetting(APPNAME, App.Title, Key, "_VACIO_")
'    If s <> "_VACIO_" Then
'        Vector = Split(s, ",")
'         Selec = UBound(Vector, 1)
'         For i = 0 To Selec
'            For j = lstBienes.ListCount - 1 To 0 Step -1
'                If Vector(i) = Left(lstBienes.List(j), lstBienes.ItemData(j)) Then
'                    'ix = mobjGrupo.AgregarUsuario(.List(i))
'                    ix = lstBienes.ItemData(j)
'                    lstServicios.AddItem lstBienes.List(j)
'                    lstServicios.ItemData(lstServicios.NewIndex) = ix
'                    lstBienes.RemoveItem j
'                End If
'            Next j
'         Next i
'    End If
'End Sub
'
Private Sub CargaRecargos()
    Dim rs As Recordset
    lstBienes.Clear
    Set rs = gobjMain.EmpresaActual.ListaIVRecargo(True)
    With rs
        If Not (.EOF) Then
            .MoveFirst
            Do Until .EOF
                lstBienes.AddItem !codRecargo & "  " & !Descripcion
                lstBienes.ItemData(lstBienes.NewIndex) = Len(!codRecargo)
               .MoveNext
           Loop
           
'            lstBienes.AddItem "SUBT" & "  " & "Subtotal"
'            lstBienes.ItemData(lstBienes.NewIndex) = Len("SUBT")
'
        End If
    End With
    rs.Close
End Sub


Public Sub RecuperaSelecRec(ByVal Key As String, trans As String)
Dim s As String, Vector As Variant, ix As Long
Dim i As Integer, j As Integer, Selec As Integer
    'Recupera selecciondados  del registro de windows
    lstServicios.Clear
    s = trans           '  jeaa 20/09/2003
    'GetSetting(APPNAME, App.Title, Key, "_VACIO_")
    If s <> "_VACIO_" Then
        Vector = Split(s, ",")
         Selec = UBound(Vector, 1)
         For i = 0 To Selec
            For j = lstBienes.ListCount - 1 To 0 Step -1
                If Vector(i) = Left(lstBienes.List(j), lstBienes.ItemData(j)) Then
                    'ix = mobjGrupo.AgregarUsuario(.List(i))
                    ix = lstBienes.ItemData(j)
                    lstServicios.AddItem lstBienes.List(j)
                    lstServicios.ItemData(lstServicios.NewIndex) = ix
                    lstBienes.RemoveItem j
                End If
            Next j
         Next i
    End If
End Sub


Private Sub CargarListaEmpresas()
    Dim rs As Recordset, i As Integer, aux As String, j As Integer
    Dim sql As String
    On Error GoTo ErrTrap

    Set rs = gobjMain.ListaEmpresas(False, False)
    i = 1
    While Not rs.EOF
'            If Trim$(gobjMain.EmpresaActual.NombreDB) <> rs!NombreDB Then
'                grdSucursal.AddItem "0" & vbTab & rs!CodEmpresa & vbTab & rs!Descripcion & vbTab & rs!NombreDB, i
                i = i + 1
 '           End If
            rs.MoveNext
    Wend
Salir:
    Set rs = Nothing
    Exit Sub

ErrTrap:
    MsgBox Err.Description, vbExclamation + vbOKOnly
    GoTo Salir
End Sub

''''Private Sub RecuperarEmpSeleccionadas()
''''    Dim Cadena As String, i As Integer, j As Integer, v As Variant
''''
''''
''''    'Recupera las empresas Sucursales
''''    Cadena = GetSetting(APPNAME, SECTION, "IVConsolidarSucursal", "")
''''    If Len(Cadena) > 0 Then
''''        v = Split(Cadena, ";")
''''    Else
''''        Exit Sub
''''    End If
''''
''''    'Recuperar del sistema las empresas seleccionadas para la consolidación
''''    For i = LBound(v) To UBound(v)
''''        For j = 1 To grdSucursal.Rows - 1
''''            If grdSucursal.TextMatrix(j, COLS_NOM) = v(i) Then
''''                grdSucursal.TextMatrix(j, COLS_SEL) = -1
''''                Exit For
''''            End If
''''        Next j
''''    Next i
''''End Sub
''''
''''Private Sub GuardarEmpSeleccionadas()
''''    Dim i As Integer, Cadena As String
''''    Cadena = ""
''''    'Guarda en registro del sistema las empresas seleccionadas para la consolidación
''''    For i = 1 To grdSucursal.Rows - 1
''''        If grdSucursal.TextMatrix(i, COLS_SEL) = vbChecked Or grdSucursal.TextMatrix(i, COLS_SEL) = -1 Then
''''            Cadena = Cadena & grdSucursal.TextMatrix(i, COLS_COD) & "," & grdSucursal.TextMatrix(i, COLS_NOM) & ";"
''''        End If
''''    Next i
''''
''''    If Len(Cadena) > 0 Then Cadena = Mid$(Cadena, 1, Len(Cadena) - 1)
''''
''''    SaveSetting APPNAME, SECTION, "IVConsolidarSucursal", Trim$(Cadena)
''''End Sub

''''Private Function ListaEmpresas() As String
''''    Dim i As Long, Cadena As String, count As Integer
''''    Dim v As Variant
''''    Cadena = Trim$(gobjMain.EmpresaActual.NombreDB) & ","
''''    count = 1
''''    For i = 1 To grdSucursal.Rows - 1
''''        If grdSucursal.TextMatrix(i, COLS_SEL) = -1 Or grdSucursal.TextMatrix(i, COLS_SEL) = vbChecked Then
''''            'maximo 15 empresa seleccionadas
''''            If count > 14 Then
''''                MsgBox "Puede seleccionar máximo 15 Empresas"
''''                Exit For
''''            End If
''''            Cadena = Cadena & Trim$(grdSucursal.TextMatrix(i, COLS_COD)) & "," & Trim$(grdSucursal.TextMatrix(i, COLS_NOM)) & ";"
''''            count = count + 1
''''        End If
''''    Next i
''''    If Len(Cadena) > 0 Then Cadena = Mid$(Cadena, 1, Len(Cadena) - 1) 'Quita la última coma
''''
''''    ListaEmpresas = Cadena
''''
''''End Function

Private Sub ConfigCols()
''''    With grdSucursal
''''         .FormatString = "|<Código|<Nombre de la Empresa|<Nombre"
''''        .ColWidth(COLS_SEL) = 600
''''        .ColWidth(COLS_COD) = 1200
''''        .ColWidth(COLS_DES) = 3420
''''        .ColWidth(COLS_NOM) = 2500
''''
''''        .ColDataType(COLS_SEL) = flexDTBoolean
''''        .ColDataType(COLS_COD) = flexDTString
''''        .ColDataType(COLS_DES) = flexDTString
''''        .ColDataType(COLS_NOM) = flexDTString
''''
''''        .ColHidden(COLS_COD) = True
''''        .ColHidden(COLS_NOM) = True
''''        .ColData(COLS_COD) = -1
''''        .ColData(COLS_DES) = -1
''''    End With
    'jeaa 12/07/04
    With grdTrans
        .FormatString = "|<Sucursal|<CodTrans|<Descripción"
        .ColWidth(COLT_SEL) = 600
        .ColWidth(COLT_SUC) = 1000
        .ColWidth(COLT_COD) = 1000
        .ColWidth(COLT_DES) = 2220
        
        .ColDataType(COLT_SEL) = flexDTBoolean
        .ColDataType(COLT_SUC) = flexDTString
        .ColDataType(COLT_COD) = flexDTString
        .ColDataType(COLT_DES) = flexDTString
        
        .ColData(COLT_COD) = -1
        .ColData(COLT_SUC) = -1
        .ColData(COLT_DES) = -1
        .col = COLT_COD
        .Sort = flexSortGenericAscending
    End With

    With grdTransRet
        .FormatString = "|<Sucursal|<CodTrans|<Descripción"
        .ColWidth(COLT_SEL) = 600
        .ColWidth(COLT_SUC) = 1000
        .ColWidth(COLT_COD) = 1000
        .ColWidth(COLT_DES) = 2220
        
        .ColDataType(COLT_SEL) = flexDTBoolean
        .ColDataType(COLT_SUC) = flexDTString
        .ColDataType(COLT_COD) = flexDTString
        .ColDataType(COLT_DES) = flexDTString
        
        .ColData(COLT_COD) = -1
        .ColData(COLT_SUC) = -1
        .ColData(COLT_DES) = -1
        .col = COLT_COD
        .Sort = flexSortGenericAscending
    End With

End Sub

Private Sub CargarListaTransEmpresasSeleccionadas(ByVal Modulo As String, ByVal TipoComp As String, ByVal tipoTrans As String, ByVal CodigoBD As String, ByVal NombreBD As String, grd As VSFlexGrid)
    Dim emp_matriz As Empresa, sql As String, rs As Recordset
    On Error GoTo ErrTrap
        sql = "SELECT '0' as Seleccion, '" & CodigoBD & "' as suc, g.CodTrans, NombreTrans "
        sql = sql & " FROM  " & NombreBD & ".dbo.GnTrans g"
        sql = sql & " INNER JOIN " & NombreBD & ".dbo.Anexo_Comprobantes ac"
        sql = sql & " ON G.AnexoCodTipoComp=AC.ID"
        sql = sql & " LEFT join " & NombreBD & ".dbo.Anexos_Transacciones at"
        sql = sql & " ON G.AnexoCodTipoTrans=At.ID"
        sql = sql & " WHERE 1=1"
        If Len(Modulo) > 0 Then
            sql = sql & " AND modulo='" & Modulo & "' "
        End If
        If Len(TipoComp) > 0 Then
            sql = sql & " AND AnexoCodTipoComp in (" & TipoComp & ")"
        End If
        If Len(tipoTrans) > 0 Then
            sql = sql & " and G.AnexoCodTipoTrans in (" & tipoTrans & ")"
        End If
        sql = sql & " Order by G.AnexoCodTipoComp desc"
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
        If rs.RecordCount > 0 Then
            rs.MoveLast
            rs.MoveFirst
            While Not (rs.EOF)
                grd.AddItem rs.Fields("Seleccion") & vbTab & rs.Fields("suc") & vbTab & rs.Fields("codTrans") & vbTab & rs.Fields("NombreTrans")
                rs.MoveNext
            Wend
        End If
    ConfigCols
    Exit Sub
ErrTrap:
    MsgBox "La base de datos [" & CodigoBD & "] seleccioanda no pertenece al servidor principal o no existe", vbInformation
End Sub

Private Sub CargarListaTransRetEmpresasSeleccionadas(ByVal CodigoBD As String, ByVal NombreBD As String, grd As VSFlexGrid)
    Dim emp_matriz As Empresa, sql As String, rs As Recordset
    On Error GoTo ErrTrap
        sql = "SELECT AnexoCodTipoComp as Seleccion, '" & CodigoBD & "' as suc, g.CodTrans, NombreTrans "
        sql = sql & " FROM  " & NombreBD & ".dbo.GnTrans g"
        sql = sql & " INNER JOIN " & NombreBD & ".dbo.Anexo_Comprobantes ac"
        sql = sql & " ON G.AnexoCodTipoComp=AC.ID"
        sql = sql & " inner join " & NombreBD & ".dbo.Anexos_Transacciones at"
        sql = sql & " ON G.AnexoCodTipoTrans=At.ID"
        sql = sql & " WHERE modulo='TS' "
        sql = sql & " AND AnexoCodTipoComp in (7)"
        sql = sql & " and G.AnexoCodTipoTrans in (1)"
        sql = sql & " Order by G.AnexoCodTipoComp desc"
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
        If rs.RecordCount > 0 Then
            rs.MoveLast
            rs.MoveFirst
            While Not (rs.EOF)
                grd.AddItem rs.Fields("Seleccion") & vbTab & rs.Fields("suc") & vbTab & rs.Fields("codTrans") & vbTab & rs.Fields("NombreTrans")
                rs.MoveNext
            Wend
        End If
    ConfigCols
    Exit Sub
ErrTrap:
    MsgBox "La base de datos [" & CodigoBD & "] seleccioanda no pertenece al servidor principal o no existe", vbInformation
End Sub


Private Function ListaEmpresasyTrans(grd As VSFlexGrid) As String
    Dim i As Long, Cadena As String, count As Integer
    Dim v As Variant
    count = 1
  '  grd.Col = COLT_SUC
    grd.Sort = flexSortGenericAscending
    For i = 1 To grd.Rows - 1
        If grd.TextMatrix(i, COLT_SEL) <> 0 Then
            Cadena = Cadena & Trim$(grd.TextMatrix(i, COLT_SUC)) & "," & Trim$(grd.TextMatrix(i, COLT_COD)) & ";"
            count = count + 1
        End If
    Next i
    If Len(Cadena) > 0 Then Cadena = Mid$(Cadena, 1, Len(Cadena) - 1) 'Quita la última coma
   ListaEmpresasyTrans = Cadena
End Function

Public Sub RecuperaSelecTransGrd(ByVal Key As String, grd As VSFlexGrid, s As String, Columna As Integer, colCodEmp As Integer)
Dim Vector As Variant, VectorAux As Variant
Dim i As Integer, j As Integer, Selec As Integer


    If s <> "_VACIO_" Then
        Vector = Split(s, ";")
         Selec = UBound(Vector, 1)
         For i = 0 To Selec
            For j = 1 To grd.Rows - 1
                VectorAux = Split(Vector(i), ",")
                If VectorAux(0) = grd.TextMatrix(j, colCodEmp) And VectorAux(1) = grd.TextMatrix(j, Columna) Then
                    grd.Select j, 0
                    grd.Text = 1
                End If
            Next j
         Next i
    End If
End Sub

