VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmB_VxTrans 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda"
   ClientHeight    =   5535
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab sst1 
      Height          =   4875
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   8599
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Parametros"
      TabPicture(0)   =   "frmB_VxTrans.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraCobros"
      Tab(0).Control(1)=   "FraSucursal"
      Tab(0).Control(2)=   "dtpFechaCorte"
      Tab(0).Control(3)=   "dtpFechaHasta"
      Tab(0).Control(4)=   "fraVenta"
      Tab(0).Control(5)=   "fraRecarDscto"
      Tab(0).Control(6)=   "fraitem"
      Tab(0).Control(7)=   "lblFechaCorte"
      Tab(0).Control(8)=   "lblFechaHasta"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Tabla Monto Ventas"
      TabPicture(1)   =   "frmB_VxTrans.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FrnmTablaComisiones"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fraCobros 
         Caption         =   "Transacciones de Cobro"
         Height          =   1812
         Left            =   -74880
         TabIndex        =   24
         Top             =   2880
         Visible         =   0   'False
         Width           =   5175
         Begin VB.ListBox lstCobros 
            Height          =   1368
            IntegralHeight  =   0   'False
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   25
            Top             =   312
            Width           =   4935
         End
      End
      Begin VB.Frame FraSucursal 
         Caption         =   "Sucursal"
         Height          =   855
         Left            =   -74880
         TabIndex        =   18
         Top             =   840
         Visible         =   0   'False
         Width           =   5175
         Begin FlexComboProy.FlexCombo fcbSucursal 
            Height          =   375
            Left            =   420
            TabIndex        =   19
            Top             =   240
            Width           =   4335
            _ExtentX        =   7646
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
      End
      Begin VB.Frame FrnmTablaComisiones 
         Caption         =   "Tabla de Monto Venta"
         Height          =   4395
         Left            =   60
         TabIndex        =   14
         Top             =   360
         Width           =   5355
         Begin VB.ComboBox cboGrupo 
            Height          =   315
            Left            =   1560
            TabIndex        =   17
            Text            =   "Combo1"
            Top             =   360
            Width           =   2295
         End
         Begin VSFlex7LCtl.VSFlexGrid grdComisiones 
            Height          =   3555
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   5085
            _cx             =   8969
            _cy             =   6271
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
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   11
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   0
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
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Asignar a Grupo"
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   420
            Width           =   1140
         End
      End
      Begin MSComCtl2.DTPicker dtpFechaCorte 
         Height          =   360
         Left            =   -73740
         TabIndex        =   10
         Top             =   420
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
         _Version        =   393216
         Format          =   39256065
         CurrentDate     =   36526
      End
      Begin MSComCtl2.DTPicker dtpFechaHasta 
         Height          =   360
         Left            =   -71040
         TabIndex        =   11
         Top             =   420
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
         _Version        =   393216
         Format          =   39256065
         CurrentDate     =   36526
      End
      Begin VB.Frame fraVenta 
         Caption         =   "Transacciones de Venta"
         Height          =   1812
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   5175
         Begin VB.ListBox lst 
            Height          =   1368
            IntegralHeight  =   0   'False
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   9
            Top             =   312
            Width           =   4935
         End
      End
      Begin VB.Frame fraRecarDscto 
         Caption         =   "Recargos ó Descuentos antes del Valor Neto"
         Height          =   2025
         Left            =   -74880
         TabIndex        =   3
         Top             =   840
         Width           =   5175
         Begin VB.ListBox lstFuente 
            Height          =   1425
            Left            =   120
            TabIndex        =   7
            Top             =   255
            Width           =   2055
         End
         Begin VB.ListBox lstDestino 
            Height          =   1425
            Left            =   3000
            TabIndex        =   6
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&>>"
            Height          =   375
            Left            =   2280
            TabIndex        =   5
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton cmdResta 
            Caption         =   "&<<"
            Height          =   375
            Left            =   2280
            TabIndex        =   4
            Top             =   840
            Width           =   615
         End
      End
      Begin VB.Frame fraitem 
         Caption         =   "Filtro de Items"
         Height          =   3030
         Left            =   -74880
         TabIndex        =   20
         Top             =   1740
         Visible         =   0   'False
         Width           =   5175
         Begin VB.TextBox lblFiltroItem 
            BackColor       =   &H80000018&
            Height          =   2445
            Left            =   1440
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   240
            Width           =   3495
         End
         Begin VB.CheckBox chkm3 
            Caption         =   "Utilizar M3"
            Height          =   255
            Left            =   300
            TabIndex        =   23
            Top             =   2700
            Width           =   1455
         End
         Begin VB.CommandButton cmdGenItem 
            Caption         =   "Filtro de  Items"
            Height          =   975
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   1092
         End
      End
      Begin VB.Label lblFechaCorte 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Corte"
         Height          =   195
         Left            =   -74880
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Corte"
         Height          =   195
         Left            =   -72180
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   400
      Left            =   2880
      TabIndex        =   1
      Top             =   5040
      Width           =   1200
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar -F5"
      Height          =   400
      Left            =   1560
      TabIndex        =   0
      Top             =   5040
      Width           =   1200
   End
End
Attribute VB_Name = "frmB_VxTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private BandAceptado As Boolean
Const PCGRUPO = 2
Const PCGRUPOCOBRO = 3
Private KeyItm As String

Public Function InicioVxTransaccion(ByRef objcond As Condicion, _
                                    ByRef Recargo As String, _
                                    ByVal tag As String) As Boolean
    Dim KeyTrans As String, KeyRecargo As String
    Me.tag = tag
    fraVenta.Caption = "Transacciones"
    lstFuente.Enabled = True
    lstDestino.Enabled = True
    cmdAdd.Enabled = True
    cmdResta.Enabled = True
    sst1.TabVisible(1) = False
    With objcond
        dtpFechaCorte.Format = dtpCustom
        dtpFechaCorte.CustomFormat = "dd/MM/yyyy"
        dtpFechaCorte.value = IIf(.FechaCorte = 0, Date, .FechaCorte)
        
        CargaTipoTrans "IV", lst
        CargaRecargo

        BandAceptado = False
        KeyTrans = "TVenta_Trans"
        KeyRecargo = "TVenta_Recar"

        RecuperaSelecTrans
        RecuperaSelecRecar

        Me.Show vbModal, frmMain
        'Si aplastó el botón 'Aceptar'
        If BandAceptado Then
            'Devuelve los valores de condición para la búsqueda
            .FechaCorte = dtpFechaCorte.value
            .CodTrans = PreparaCadena(lst)
            Recargo = PreparaCadRec(lstDestino)
            'grabar las formas de cobro a visualizar
            SaveSetting APPNAME, App.Title, KeyTrans, .CodTrans
            SaveSetting APPNAME, App.Title, KeyRecargo, Recargo
        End If
    End With
    'Devuelve true/false
    Unload Me
    InicioVxTransaccion = BandAceptado
End Function

Public Sub RecuperaSelecRecar()
Dim s As String, Vector As Variant, ix As Long
Dim i As Integer, j As Integer, Selec As Integer
    'Recupera selecciondados  del registro de windows
    s = GetSetting(APPNAME, App.Title, "TVenta_Recar", "_VACIO_")
    If s <> "_VACIO_" Then
        Vector = Split(s, ",")
         Selec = UBound(Vector, 1)
         For i = 0 To Selec
            For j = lstFuente.ListCount - 1 To 0 Step -1
                If Vector(i) = Left(lstFuente.List(j), lstFuente.ItemData(j)) Then
                    ix = lstFuente.ItemData(j)
                    lstDestino.AddItem lstFuente.List(j)
                    lstDestino.ItemData(lstDestino.NewIndex) = ix
                    lstFuente.RemoveItem j
                End If
            Next j
         Next i
    End If
End Sub

Private Function PreparaCadena(lst As ListBox) As String
    Dim Cadena As String, i As Integer
    Cadena = ""
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            If Cadena = "" Then
                Cadena = Left(lst.List(i), lst.ItemData(i))
            Else
                Cadena = Cadena & "," & _
                              Left(lst.List(i), lst.ItemData(i))
            End If
        End If
    Next i
    PreparaCadena = Cadena
End Function

Private Function PreparaCadRec(lst As ListBox) As String
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
    PreparaCadRec = Cadena
End Function

Private Sub PreparaListaTransIV()
    Dim rs As Recordset
   'Prepara la lista de tipos de transaccion
    lst.Clear
    Set rs = gobjMain.EmpresaActual.ListaGNTrans("IV", False, True)
    With rs
        If Not (.EOF) Then
            .MoveFirst
            Do Until .EOF
                lst.AddItem !CodTrans & "  " & !NombreTrans
                lst.ItemData(lst.NewIndex) = Len(!CodTrans)
                .MoveNext
            Loop
        End If
    End With
    rs.Close
    Set rs = Nothing
End Sub


Private Sub cboGrupo_Change()
'        If Not CargarPCGrupos(cboGrupo.ListIndex + 1) Then
'                grdComisiones.Rows = 1
'                Exit Sub
'        End If
End Sub

Private Sub cboGrupo_Click()
ConfigColsClasificador
    If Not CargarPCGruposCobros(cboGrupo.ListIndex + 1) Then
                grdComisiones.Rows = 1
                Exit Sub
        End If
End Sub

Private Sub cmdAceptar_Click()
    BandAceptado = True
'    dtpFechaCorte.SetFocus
    Me.Hide
End Sub

Private Sub cmdAdd_Click()
    Dim i As Long, ix As Long
    On Error GoTo ErrTrap
    With lstFuente
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                ix = .ItemData(i)
                lstDestino.AddItem .List(i)
                lstDestino.ItemData(lstDestino.NewIndex) = ix
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
    With lstDestino
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                ix = .ItemData(i)
                lstFuente.AddItem .List(i)
                lstFuente.ItemData(lstFuente.NewIndex) = ix
                .RemoveItem i
            End If
        Next i
    End With
    Exit Sub
ErrTrap:
    DispErr
End Sub

Private Sub cmdCancelar_Click()
    BandAceptado = False
    dtpFechaCorte.SetFocus
    Me.Hide
End Sub

Private Sub dtpFechaHasta_Click()
    dtpFechaCorte.value = DateAdd("m", -6, dtpFechaHasta.value)
End Sub

Private Sub dtpFechaHasta_KeyDown(KeyCode As Integer, Shift As Integer)
    dtpFechaCorte.value = DateAdd("m", -6, dtpFechaHasta.value)
End Sub

Private Sub dtpFechaHasta_LostFocus()
dtpFechaCorte.value = DateAdd("m", -6, dtpFechaHasta.value)
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

Private Sub CargaRecargo()
    Dim rs As Recordset
    Set rs = gobjMain.EmpresaActual.ListaIVRecargo(True)
    With rs
        If Not (.EOF) Then
            .MoveFirst
            Do Until .EOF
                lstFuente.AddItem !codRecargo & "  " & !Descripcion
                lstFuente.ItemData(lstFuente.NewIndex) = Len(!codRecargo)
               .MoveNext
           Loop
            lstFuente.AddItem "SUBT" & "  " & "Subtotal"
            lstFuente.ItemData(lstFuente.NewIndex) = Len("SUBT")
        End If
    End With
    rs.Close
End Sub

Private Sub Form_Load()
    'Establece los rangos de Fecha  siempre  al rango
    'del año actual
    dtpFechaCorte.value = Date
End Sub

Private Sub Label2_Click()

End Sub

Private Sub lstDestino_DblClick()
    cmdResta_Click
End Sub

Private Sub lstFuente_DblClick()
    cmdAdd_Click
End Sub

Private Sub CargaTipoTrans(ByRef Modulo As String, ByRef lst As ListBox)
    Dim rs As Recordset, Vector As Variant
    Dim numMod As Integer, i As Integer
    'Prepara la lista de tipos de transaccion
    lst.Clear
    Vector = Split(Modulo, ",")
    numMod = UBound(Vector, 1)
    If numMod = -1 Then
        Set rs = gobjMain.EmpresaActual.ListaGNTrans("", False, True)
        With rs
            If Not (.EOF) Then
                .MoveFirst
                Do Until .EOF
                    lst.AddItem !CodTrans & "  " & !NombreTrans
                    lst.ItemData(lst.NewIndex) = Len(!CodTrans)
                    .MoveNext
                Loop
            End If
        End With
        rs.Close
    Else
        For i = 0 To numMod
            Set rs = gobjMain.EmpresaActual.ListaGNTrans(CStr(Vector(i)), False, True)
            With rs
                If Not (.EOF) Then
                    .MoveFirst
                    Do Until .EOF
                        lst.AddItem !CodTrans & "  " & !NombreTrans
                        lst.ItemData(lst.NewIndex) = Len(!CodTrans)
                        .MoveNext
                    Loop
                End If
            End With
            rs.Close
        Next i
    End If
    Set rs = Nothing
End Sub

Private Sub RecuperaSelecTrans()
    Dim Vector As Variant, s As String
    Dim i As Integer, j As Integer, Selec As Integer
    s = GetSetting(APPNAME, App.Title, "TVenta_Trans", "_VACIO_")
    If s <> "_VACIO_" Then
        Vector = Split(s, ",")
         Selec = UBound(Vector, 1)
         For i = 0 To Selec
            For j = 0 To lst.ListCount - 1
                If Vector(i) = Left(lst.List(j), lst.ItemData(j)) Then
                    lst.Selected(j) = True
                End If
            Next j
         Next i
    End If
End Sub


Public Function InicioVxMesTransaccion(ByRef objcond As Condicion, _
                                    ByRef Recargo As String, _
                                    ByVal tag As String) As Boolean
    Dim KeyTrans As String, KeyRecargo As String
    Me.tag = tag
    fraVenta.Caption = "Transacciones"
    lstFuente.Enabled = True
    lstDestino.Enabled = True
    cmdAdd.Enabled = True
    cmdResta.Enabled = True
    lblFechaCorte.Caption = "Fecha desde"
    lblFechaHasta.Visible = True
    dtpFechaHasta.Visible = True
    sst1.TabVisible(1) = False
    With objcond
        dtpFechaCorte.Format = dtpCustom
        dtpFechaCorte.CustomFormat = "dd/MMM/yyyy"
        
        dtpFechaHasta.Format = dtpCustom
        dtpFechaHasta.CustomFormat = "dd/MMM/yyyy"
        
        dtpFechaCorte.value = IIf(.fecha1 = 0, Date, .fecha1)
        dtpFechaHasta.value = IIf(.fecha2 = 0, Date, .fecha2)
        
        
        CargaTipoTrans "IV", lst
        CargaRecargo

        BandAceptado = False
        KeyTrans = "PromVenta_Trans"
        KeyRecargo = "PromVenta_Recar"

        RecuperaSelecTransProm
        RecuperaSelecRecarProm

        Me.Show vbModal, frmMain
        'Si aplastó el botón 'Aceptar'
        If BandAceptado Then
            'Devuelve los valores de condición para la búsqueda
            .fecha1 = dtpFechaCorte.value
            '.Fecha2 = DateAdd("d", -1, CDate("01/" & DatePart("m", dtpFechaCorte.value) & "/" & DatePart("yyyy", dtpFechaHasta.value)))
            .fecha2 = dtpFechaHasta.value
            
            .CodTrans = PreparaCadena(lst)
            Recargo = PreparaCadRec(lstDestino)
            'grabar las formas de cobro a visualizar
            SaveSetting APPNAME, App.Title, KeyTrans, .CodTrans
            SaveSetting APPNAME, App.Title, KeyRecargo, Recargo
        End If
    End With
    'Devuelve true/false
    Unload Me
    InicioVxMesTransaccion = BandAceptado
End Function


Private Sub RecuperaSelecTransProm()
    Dim Vector As Variant, s As String
    Dim i As Integer, j As Integer, Selec As Integer
    s = GetSetting(APPNAME, App.Title, "PromVenta_Trans", "_VACIO_")
    If s <> "_VACIO_" Then
        Vector = Split(s, ",")
         Selec = UBound(Vector, 1)
         For i = 0 To Selec
            For j = 0 To lst.ListCount - 1
                If Vector(i) = Left(lst.List(j), lst.ItemData(j)) Then
                    lst.Selected(j) = True
                End If
            Next j
         Next i
    End If
End Sub

Public Sub RecuperaSelecRecarProm()
Dim s As String, Vector As Variant, ix As Long
Dim i As Integer, j As Integer, Selec As Integer
    'Recupera selecciondados  del registro de windows
    s = GetSetting(APPNAME, App.Title, "PromVenta_Recar", "_VACIO_")
    If s <> "_VACIO_" Then
        Vector = Split(s, ",")
         Selec = UBound(Vector, 1)
         For i = 0 To Selec
            For j = lstFuente.ListCount - 1 To 0 Step -1
                If Vector(i) = Left(lstFuente.List(j), lstFuente.ItemData(j)) Then
                    ix = lstFuente.ItemData(j)
                    lstDestino.AddItem lstFuente.List(j)
                    lstDestino.ItemData(lstDestino.NewIndex) = ix
                    lstFuente.RemoveItem j
                End If
            Next j
         Next i
    End If
End Sub

Private Sub ConfigColsComisiones() 'grilla para el Impuesto a la Renta
    With grdComisiones
        .FormatString = ">Desde|>Hasta|<Grupo"
        .ColWidth(0) = 1000
        .ColWidth(1) = 1000
        .ColWidth(2) = 2000
        .ColDataType(0) = flexDTCurrency
        .ColDataType(1) = flexDTCurrency
    End With
End Sub

Private Sub ConfigColsClasificador()
    With grdComisiones
        .FormatString = ">Desde|>Hasta|>DiasMorosidad|<Grupo"
        .ColWidth(0) = 1000
        .ColWidth(1) = 1000
        .ColWidth(2) = 1000
        .ColWidth(3) = 2000
        .ColDataType(0) = flexDTCurrency
        .ColDataType(1) = flexDTCurrency
        .ColDataType(2) = flexDTCurrency
    End With
End Sub


Private Sub GrabaIntervalos()
    Dim i As Integer
    With grdComisiones
        For i = .FixedRows To .Rows - 1
            gMonto(i).desde = .ValueMatrix(i, 0)
            gMonto(i).hasta = .ValueMatrix(i, 1)
            gMonto(i).grupo = .TextMatrix(i, 2)
        Next i
    End With
    EscribirIntervalosGnOpcionMontoVentas
End Sub

Private Sub grdComisiones_BeforeEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
    grdComisiones.EditMaxLength = 12 'Hasta 99,999,999,999
End Sub

'Private Sub grdComisiones_KeyPressEdit(ByVal Row As Long, ByVal col As Long, KeyAscii As Integer)
'    'Acepta sólo númericos
'    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And (KeyAscii <> vbKeyBack) And (KeyAscii <> Asc(".")) Then
'        KeyAscii = 0
'    End If
'End Sub

Private Sub CargaTablaComisiones()
    Dim i As Integer
'    LeerIntervalos
    ConfigColsComisiones
    grdComisiones.Rows = 11
    If Not CargarPCGrupos(cboGrupo.ListIndex + 1) Then
    End If
    
    For i = 1 To 10
        grdComisiones.TextMatrix(i, 0) = gMonto(i).desde
        grdComisiones.TextMatrix(i, 1) = gMonto(i).hasta
        grdComisiones.TextMatrix(i, 2) = gMonto(i).grupo
    Next i
    ConfigColsComisiones
End Sub

Public Function InicioMontoVentas(ByRef objcond As Condicion, _
                                    ByRef Recargo As String, _
                                    ByVal tag As String) As Boolean
    Dim KeyTrans As String, KeyRecargo As String, i As Integer
    Me.tag = tag
    fraVenta.Caption = "Transacciones"
    lstFuente.Enabled = True
    lstDestino.Enabled = True
    cmdAdd.Enabled = True
    cmdResta.Enabled = True
    lblFechaCorte.Caption = "Fecha desde"
    lblFechaHasta.Visible = True
    dtpFechaHasta.Visible = True
    fraVenta.Top = fraRecarDscto.Top
    fraVenta.Height = fraRecarDscto.Height * 1.9
    lst.Height = fraRecarDscto.Height * 1.7
    fraRecarDscto.Visible = False
    sst1.TabVisible(1) = True
    
    
    
    With objcond
        dtpFechaCorte.Format = dtpCustom
        dtpFechaCorte.CustomFormat = "MMM/yyyy"
        
        dtpFechaHasta.Format = dtpCustom
        dtpFechaHasta.CustomFormat = "MMM/yyyy"
        
        dtpFechaCorte.value = IIf(.fecha1 = 0, Date, .fecha1)
        dtpFechaHasta.value = IIf(.fecha2 = 0, Date, .fecha2)
        
        cboGrupo.Clear
         For i = 1 To PCGRUPO_MAX
             cboGrupo.AddItem gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupoC(i)
         Next i
         If (.numGrupo <= cboGrupo.ListCount) And (.numGrupo > 0) Then
             cboGrupo.ListIndex = .numGrupo - 1   'Selecciona lo anterior
         ElseIf cboGrupo.ListCount > 0 Then
             cboGrupo.ListIndex = 0              'Selecciona la primera
         End If
        
        
        CargaTipoTrans "IV", lst
        'CargaRecargo

        BandAceptado = False
        KeyTrans = "PCGMontoVenta_Trans"
        KeyRecargo = "PCGMontoVenta_NumPCGrupo"

        RecuperaSelecTransPromGnOpcion
        cboGrupo.ListIndex = RecuperaSelecPCGrupo

        LeerIntervalosGnOpcionMontoVentas
        CargaTablaComisiones

        Me.Show vbModal, frmMain
        'Si aplastó el botón 'Aceptar'
        If BandAceptado Then
            'Devuelve los valores de condición para la búsqueda
            .fecha1 = "01/ " & DatePart("m", dtpFechaCorte.value) & "/" & DatePart("yyyy", dtpFechaCorte.value)
            .fecha2 = DateAdd("d", -1, "01/ " & DatePart("m", DateAdd("m", 1, dtpFechaHasta.value)) & "/" & DatePart("yyyy", dtpFechaCorte.value))
            .CodTrans = PreparaCadena(lst)
            
            If cboGrupo.ListIndex >= 0 Then
                 .numGrupo = cboGrupo.ListIndex + 1
             End If
            
            
         '   Recargo = PreparaCadRec(lstDestino)
            'grabar las formas de cobro a visualizar
'            SaveSetting APPNAME, App.Title, KeyTrans, .CodTrans
'            SaveSetting APPNAME, App.Title, KeyRecargo, .numGrupo
            
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "PCGMontoVenta_Trans", .CodTrans
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "PCGMontoVenta_NumPCGrupo", .numGrupo
            gobjMain.EmpresaActual.GNOpcion.GrabarGNOpcion2
            
            
            GrabaIntervalos
        End If
    End With
    'Devuelve true/false
    Unload Me
    InicioMontoVentas = BandAceptado
End Function


Public Function RecuperaSelecPCGrupo() As Integer
Dim s As String, Vector As Variant, ix As Long
Dim i As Integer, j As Integer, Selec As Integer
    'Recupera selecciondados  del registro de windows
    's = GetSetting(APPNAME, App.Title, "PCGMontoVenta_NumPCGrupo", "_VACIO_")
    s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("PCGMontoVenta_NumPCGrupo")
    If Len(s) > 0 Then
        RecuperaSelecPCGrupo = CInt(s) - 1
    Else
        RecuperaSelecPCGrupo = 0
    End If
End Function

Private Sub RecuperaSelecTransPromGnOpcion()
    Dim Vector As Variant, s As String
    Dim i As Integer, j As Integer, Selec As Integer
    's = GetSetting(APPNAME, App.Title, "PromVenta_Trans", "_VACIO_")
    s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("PCGMontoVenta_Trans")
    If Len(s) > 0 Then
        Vector = Split(s, ",")
         Selec = UBound(Vector, 1)
         For i = 0 To Selec
            For j = 0 To lst.ListCount - 1
                If Vector(i) = Left(lst.List(j), lst.ItemData(j)) Then
                    lst.Selected(j) = True
                End If
            Next j
         Next i
    End If
End Sub
Private Sub RecuperaSelecTransPromGnOpcionCobro()
    Dim Vector As Variant, s As String
    Dim i As Integer, j As Integer, Selec As Integer
    's = GetSetting(APPNAME, App.Title, "PromVenta_Trans", "_VACIO_")
    s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("PCGMontoVenta_TransCobros")
    If Len(s) > 0 Then
        Vector = Split(s, ",")
         Selec = UBound(Vector, 1)
         For i = 0 To Selec
            For j = 0 To lstCobros.ListCount - 1
                If Vector(i) = Left(lstCobros.List(j), lstCobros.ItemData(j)) Then
                    lstCobros.Selected(j) = True
                End If
            Next j
         Next i
    End If
End Sub

Private Function CargarPCGrupos(ByVal numGrupo As Integer) As Boolean
    Dim s As Variant
    On Error GoTo ErrTrap
    With grdComisiones
        CargarPCGrupos = True
        s = gobjMain.EmpresaActual.ListaPCGrupoOrigenParaFlexGrid(numGrupo, 2)
        If Len(s) > 1 Then
            s = Right$(s, Len(s) - 1)
            .ColComboList(PCGRUPO) = s
        End If
    End With
    Exit Function
ErrTrap:
        MsgBox "No se han definido PCGrupos", vbInformation
        CargarPCGrupos = False
    Exit Function
End Function


Public Function InicioVxSucursal(ByRef objcond As Condicion, _
                                    ByRef Recargo As String, _
                                    ByVal tag As String) As Boolean
    Dim KeyTrans As String, KeyRecargo As String
    Me.tag = tag
    lblFechaCorte.Caption = "Fecha desde"
    lblFechaHasta.Visible = True
    dtpFechaHasta.Visible = True
    sst1.TabVisible(1) = False
    fraRecarDscto.Visible = False
    fraVenta.Visible = False
    With objcond
        dtpFechaCorte.Format = dtpCustom
        dtpFechaCorte.CustomFormat = "dd/MMM/yyyy"
        
        dtpFechaHasta.Format = dtpCustom
        dtpFechaHasta.CustomFormat = "dd/MMM/yyyy"
        

        dtpFechaHasta.value = DateAdd("d", -1, Date)
        dtpFechaCorte.value = DateAdd("m", -6, dtpFechaHasta.value)
        
          dtpFechaCorte.Enabled = False
     '   dtpFechaHasta.Enabled = False
        
        'FraSucursal.Visible = True
        'fcbSucursal.SetData gobjMain.EmpresaActual.ListaGNSucursales(True, False) 'jeaa 10/09/2008
        
        fraItem.Visible = True
        fcbSucursal.KeyText = .Sucursal

        Me.Show vbModal, frmMain
        'Si aplastó el botón 'Aceptar'
        If BandAceptado Then
            'Devuelve los valores de condición para la búsqueda
            .fecha1 = dtpFechaCorte.value
            .fecha2 = dtpFechaHasta.value
            .Sucursal = fcbSucursal.KeyText
            .BandTodo = (chkm3.value = vbChecked)
            
        End If
    End With
    'Devuelve true/false
    Unload Me
    InicioVxSucursal = BandAceptado
End Function


Private Sub cmdGenItem_Click()
    Dim frmItem As frmB_FiltroxItem, EtiqItem As String
    Set frmItem = New frmB_FiltroxItem
    EtiqItem = frmItem.Inicio(Me.tag, KeyItm)
    lblFiltroItem.Text = EtiqItem
End Sub

Public Function InicioMontoVentasCobros(ByRef objcond As Condicion, _
                                    ByVal tag As String) As Boolean
    Dim KeyTrans As String, KeyRecargo As String, i As Integer
    Me.tag = tag
    fraVenta.Caption = "Transacciones Venta"
    fraCobros.Visible = True
    lstFuente.Enabled = True
    lstDestino.Enabled = True
    cmdAdd.Enabled = True
    cmdResta.Enabled = True
    lblFechaCorte.Caption = "Fecha desde"
    lblFechaHasta.Visible = True
    dtpFechaHasta.Visible = True
    fraVenta.Top = fraRecarDscto.Top
    'fraVenta.Height = fraRecarDscto.Height * 1.9
    'lst.Height = fraRecarDscto.Height * 1.7
    fraCobros.Visible = True
    fraRecarDscto.Visible = False
    sst1.TabVisible(1) = True
    With objcond
        dtpFechaCorte.Format = dtpCustom
        dtpFechaCorte.CustomFormat = "MMM/yyyy"
        
        dtpFechaHasta.Format = dtpCustom
        dtpFechaHasta.CustomFormat = "MMM/yyyy"
        
        dtpFechaCorte.value = IIf(.fecha1 = 0, Date, .fecha1)
        dtpFechaHasta.value = IIf(.fecha2 = 0, Date, .fecha2)
        
        cboGrupo.Clear
         For i = 1 To PCGRUPO_MAX
             cboGrupo.AddItem gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupoC(i)
         Next i
         If (.numGrupo <= cboGrupo.ListCount) And (.numGrupo > 0) Then
             cboGrupo.ListIndex = .numGrupo - 1   'Selecciona lo anterior
         ElseIf cboGrupo.ListCount > 0 Then
             cboGrupo.ListIndex = 0              'Selecciona la primera
         End If
        CargaTipoTrans "IV", lst
        CargaTipoTrans "TS", lstCobros
        
        BandAceptado = False
        KeyTrans = "PCGMontoVenta_Trans"
        KeyRecargo = "PCGMontoVenta_NumPCGrupo"

       RecuperaSelecTransPromGnOpcion
       RecuperaSelecTransPromGnOpcionCobro
        cboGrupo.ListIndex = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("PCGMontoVenta_NumPCGrupo") - 1

        LeerIntervalosGnOpcionMontoVentasCobros
        CargaTablaClasificador

        Me.Show vbModal, frmMain
        'Si aplastó el botón 'Aceptar'
        If BandAceptado Then
            'Devuelve los valores de condición para la búsqueda
            .fecha1 = "01/ " & DatePart("m", dtpFechaCorte.value) & "/" & DatePart("yyyy", dtpFechaCorte.value)
            .fecha2 = DateAdd("d", -1, "01/ " & DatePart("m", DateAdd("m", 1, dtpFechaHasta.value)) & "/" & DatePart("yyyy", dtpFechaCorte.value))
            .CodTrans = PreparaCadena(lst)
            .codforma = PreparaCadena(lstCobros)
            If cboGrupo.ListIndex >= 0 Then
                 .numGrupo = cboGrupo.ListIndex + 1
             End If
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "PCGMontoVenta_Trans", .CodTrans
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "PCGMontoVenta_TransCobros", .codforma
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "PCGMontoVenta_NumPCGrupo", .numGrupo
            gobjMain.EmpresaActual.GNOpcion.GrabarGNOpcion2
            GrabaIntervalosCobro
        End If
    End With
    'Devuelve true/false
    Unload Me
    InicioMontoVentasCobros = BandAceptado
End Function

Private Sub CargaTablaClasificador()
    Dim i As Integer
    ConfigColsClasificador
    grdComisiones.Rows = 10
    If Not CargarPCGruposCobros(cboGrupo.ListIndex + 1) Then
    End If
    For i = 1 To 9
        grdComisiones.TextMatrix(i, 0) = gMontoCobro(i).desde
        grdComisiones.TextMatrix(i, 1) = gMontoCobro(i).hasta
        grdComisiones.TextMatrix(i, 2) = gMontoCobro(i).diasMorosidad
        grdComisiones.TextMatrix(i, 3) = gMontoCobro(i).grupo
    Next i
    ConfigColsClasificador
End Sub
Private Function CargarPCGruposCobros(ByVal numGrupo As Integer) As Boolean
    Dim s As Variant
    On Error GoTo ErrTrap
    With grdComisiones
        CargarPCGruposCobros = True
        s = gobjMain.EmpresaActual.ListaPCGrupoOrigenParaFlexGrid(numGrupo, 2)
        If Len(s) > 1 Then
            s = Right$(s, Len(s) - 1)
            .ColComboList(PCGRUPOCOBRO) = s
        End If
    End With
    Exit Function
ErrTrap:
        MsgBox "No se han definido PCGrupos", vbInformation
        CargarPCGruposCobros = False
    Exit Function
End Function

Private Sub GrabaIntervalosCobro()
    Dim i As Integer
    With grdComisiones
        For i = .FixedRows To .Rows - 1
            gMontoCobro(i).desde = .ValueMatrix(i, 0)
            gMontoCobro(i).hasta = .ValueMatrix(i, 1)
            gMontoCobro(i).diasMorosidad = .ValueMatrix(i, 2)
            gMontoCobro(i).grupo = .TextMatrix(i, 3)
        Next i
    End With
    EscribirIntervalosGnOpcionMontoVentasCobro
End Sub

