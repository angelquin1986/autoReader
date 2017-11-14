VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmB_CxTrans 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda"
   ClientHeight    =   3315
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkProm 
      Caption         =   "Prom Vta Catalogo"
      Height          =   255
      Left            =   180
      TabIndex        =   18
      Top             =   2820
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame FraEmpresas 
      Caption         =   "Empresa"
      Height          =   2655
      Left            =   5400
      TabIndex        =   11
      Top             =   60
      Visible         =   0   'False
      Width           =   2835
      Begin VB.ListBox lstBase 
         Height          =   2310
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   12
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   400
      Left            =   4020
      TabIndex        =   3
      Top             =   2820
      Width           =   1200
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar -F5"
      Height          =   400
      Left            =   2580
      TabIndex        =   2
      Top             =   2820
      Width           =   1200
   End
   Begin MSComCtl2.DTPicker dtpFechaCorte 
      Height          =   360
      Left            =   1440
      TabIndex        =   0
      Top             =   180
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      _Version        =   393216
      Format          =   114360321
      CurrentDate     =   36526
   End
   Begin VB.Frame fraFecha 
      Caption         =   "Rango de Fechas"
      Height          =   675
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   5175
      Begin MSComCtl2.DTPicker DTPFechaDesde 
         Height          =   360
         Left            =   1080
         TabIndex        =   7
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   635
         _Version        =   393216
         Format          =   114360321
         CurrentDate     =   36526
      End
      Begin MSComCtl2.DTPicker DTPFechaHasta 
         Height          =   360
         Left            =   3240
         TabIndex        =   9
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   635
         _Version        =   393216
         Format          =   114360321
         CurrentDate     =   36526
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
         Height          =   195
         Left            =   2760
         TabIndex        =   10
         Top             =   300
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
         Height          =   195
         Left            =   540
         TabIndex        =   8
         Top             =   300
         Width           =   465
      End
   End
   Begin VB.Frame fraVenta 
      Caption         =   "Transacciones de Venta"
      Height          =   2115
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   5175
      Begin VB.ListBox lst 
         Height          =   1725
         IntegralHeight  =   0   'False
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame FraEmpresa 
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   -60
      Visible         =   0   'False
      Width           =   5175
      Begin VB.ComboBox cboBaseMatriz 
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   300
         Width           =   2115
      End
      Begin MSComCtl2.DTPicker DTPFechaHasta1 
         Height          =   360
         Left            =   3720
         TabIndex        =   16
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   635
         _Version        =   393216
         Format          =   114360321
         CurrentDate     =   36526
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
         Height          =   195
         Left            =   3240
         TabIndex        =   17
         Top             =   300
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   360
         Width           =   795
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Corte"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmB_CxTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BandAceptado As Boolean

Public Function InicioVxTransaccion(ByRef objcond As Condicion, _
                                    ByRef Recargo As String, _
                                    ByVal tag As String) As Boolean
    Dim KeyTrans As String, KeyRecargo As String
    Me.tag = tag
    fraVenta.Caption = "Transacciones"
    Label1.Visible = True
    dtpFechaCorte.Visible = True
    fraFecha.Visible = False
   
    With objcond
        dtpFechaCorte.Format = dtpCustom
        dtpFechaCorte.CustomFormat = "dd/MM/yyyy"
        dtpFechaCorte.value = IIf(.FechaCorte = 0, Date, .FechaCorte)
        
        CargaTipoTrans "IV", lst

        BandAceptado = False
        KeyTrans = "TCompra_Trans"
        RecuperaSelecTrans

        Me.Show vbModal, frmMain
        'Si aplastó el botón 'Aceptar'
        If BandAceptado Then
            'Devuelve los valores de condición para la búsqueda
            .FechaCorte = dtpFechaCorte.value
            .CodTrans = PreparaCadena(lst)
            'grabar las formas de cobro a visualizar
            SaveSetting APPNAME, App.Title, KeyTrans, .CodTrans
        End If
    End With
    'Devuelve true/false
    Unload Me
    InicioVxTransaccion = BandAceptado
End Function

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


Private Sub cmdAceptar_Click()
    BandAceptado = True
    If Me.tag = "UltimoCosto" Then
        DTPFechaDesde.SetFocus
    ElseIf Me.tag = "Buffer" Then
        DTPFechaDesde.SetFocus
    ElseIf Me.tag <> "ComprasxVendedor" And Me.tag <> "CustoxActivo" Then
        If Me.tag <> "FechgaIngreso" And Me.tag <> "FechgaEgreso" And Me.tag <> "BufferxAlm_CodBodega" Then
            dtpFechaCorte.SetFocus
        End If

    Else
        DTPFechaDesde.SetFocus
    End If
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    BandAceptado = False
    If Me.tag = "Buffer" Then
        DTPFechaDesde.SetFocus
    ElseIf Me.tag <> "ComprasxVendedor" And Me.tag <> "CustoxActivo" Then
        If Me.tag <> "FechgaIngreso" And Me.tag <> "FechgaEgreso" And Me.tag <> "BufferxAlm_CodBodega" Then
            dtpFechaCorte.SetFocus
        End If
    Else
        DTPFechaDesde.SetFocus
    End If
    Me.Hide
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

Private Sub Form_Load()
    'Establece los rangos de Fecha  siempre  al rango
    'del año actual
    dtpFechaCorte.value = Date
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
    If Me.tag = "ComprasxVendedor" Then
        s = GetSetting(APPNAME, App.Title, "TCompra_Trans", "_VACIO_")
    ElseIf Me.tag = "CustoxActivo" Then
        s = GetSetting(APPNAME, App.Title, "TCusto_Trans", "_VACIO_")
    ElseIf Me.tag = "FechgaEgreso" Then
        s = GetSetting(APPNAME, App.Title, "FechaTransEgreso_Trans", "_VACIO_")
    ElseIf Me.tag = "FechgaIngreso" Then
        s = GetSetting(APPNAME, App.Title, "FechaTransIngreso_Trans", "_VACIO_")
    ElseIf Me.tag = "Buffer" Then
        s = GetSetting(APPNAME, App.Title, "Buffer_Trans", "_VACIO_")
    ElseIf Me.tag = "BufferxAlm_CodBodega" Then
        s = GetSetting(APPNAME, App.Title, "BufferxAlm_CodBodega", "_VACIO_")
    
    
    Else
        s = GetSetting(APPNAME, App.Title, "TVenta_Trans", "_VACIO_")
    End If
    If s <> "_VACIO_" Then
        Vector = Split(s, ",")
         Selec = UBound(Vector, 1)
         For i = 0 To Selec
            For j = 0 To lst.ListCount - 1
                If Mid$(Vector(i), 2, Len(Vector(i)) - 2) = Left(lst.List(j), lst.ItemData(j)) Then
                    lst.Selected(j) = True
                End If
            Next j
         Next i
    End If
End Sub

Public Function InicioCxProveedor(ByRef objcond As Condicion, _
                                    ByRef Recargo As String, _
                                    ByVal tag As String) As Boolean
    Dim KeyTrans As String, KeyRecargo As String
    Me.tag = tag
    fraVenta.Caption = "Transacciones"
    Label1.Visible = False
    dtpFechaCorte.Visible = False
    fraFecha.Visible = True
    With objcond
        DTPFechaHasta.Format = dtpCustom
        DTPFechaDesde.Format = dtpCustom
        DTPFechaHasta.CustomFormat = "dd/MM/yyyy"
        DTPFechaDesde.CustomFormat = "dd/MM/yyyy"
        DTPFechaDesde.value = IIf(.fecha1 = 0, Date, .fecha1)
        DTPFechaHasta.value = IIf(.fecha2 = 0, Date, .fecha2)
        
        CargaTipoTrans "IV", lst

        BandAceptado = False
        KeyTrans = "TCompra_Trans"
        RecuperaSelecTrans

        Me.Show vbModal, frmMain
        'Si aplastó el botón 'Aceptar'
        If BandAceptado Then
            'Devuelve los valores de condición para la búsqueda
            .fecha1 = DTPFechaDesde.value
            .fecha2 = DTPFechaHasta.value
            If Me.tag = "UltimoCosto" Then
                .CodTrans = PreparaCadenaIN(lst)
            Else
                .CodTrans = PreparaCadena(lst)
            End If
            'grabar las formas de cobro a visualizar
            SaveSetting APPNAME, App.Title, KeyTrans, .CodTrans
        End If
    End With
    'Devuelve true/false
    Unload Me
    InicioCxProveedor = BandAceptado
End Function


Private Function PreparaCadenaIN(lst As ListBox) As String
    Dim Cadena As String, i As Integer
    Cadena = ""
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            If Cadena = "" Then
                Cadena = "'" & Left(lst.List(i), lst.ItemData(i)) & "',"
            Else
                Cadena = Cadena & "'" & _
                              Left(lst.List(i), lst.ItemData(i)) & "',"
            End If
        End If
    Next i
    PreparaCadenaIN = Mid$(Cadena, 1, Len(Cadena) - 1)
End Function


Public Function InicioCustodioxActivo(ByRef objcond As Condicion, _
                                    ByRef Recargo As String, _
                                    ByVal tag As String) As Boolean
    Dim KeyTrans As String, KeyRecargo As String
    Me.tag = tag
    fraVenta.Caption = "Transacciones"
    Label1.Visible = False
    dtpFechaCorte.Visible = False
    fraFecha.Visible = True
    With objcond
        DTPFechaHasta.Format = dtpCustom
        DTPFechaDesde.Format = dtpCustom
        DTPFechaHasta.CustomFormat = "dd/MM/yyyy"
        DTPFechaDesde.CustomFormat = "dd/MM/yyyy"
        DTPFechaDesde.value = IIf(.fecha1 = 0, Date, .fecha1)
        DTPFechaHasta.value = IIf(.fecha2 = 0, Date, .fecha2)
        
        CargaTipoTrans "AF", lst

        BandAceptado = False
        KeyTrans = "TCusto_Trans"
        RecuperaSelecTrans

        Me.Show vbModal, frmMain
        'Si aplastó el botón 'Aceptar'
        If BandAceptado Then
            'Devuelve los valores de condición para la búsqueda
            .fecha1 = DTPFechaDesde.value
            .fecha2 = DTPFechaHasta.value
            If Me.tag = "UltimoCosto" Then
                .CodTrans = PreparaCadenaIN(lst)
            Else
                .CodTrans = PreparaCadena(lst)
            End If
            'grabar las formas de cobro a visualizar
            SaveSetting APPNAME, App.Title, KeyTrans, .CodTrans
        End If
    End With
    'Devuelve true/false
    Unload Me
    InicioCustodioxActivo = BandAceptado
End Function


Private Sub CargaListaEmpresas()
    Dim rs As Recordset, i As Integer
    On Error GoTo ErrTrap

    lstBase.Clear
    'Lista las empresas existentes
    Set rs = gobjMain.ListaEmpresas(False, False)
    With rs
        Do Until .EOF
            lstBase.AddItem rs!CodEmpresa
            cboBaseMatriz.AddItem rs!CodEmpresa
            .MoveNext
        Loop
        .Close
    End With
Salir:
    Set rs = Nothing
    Exit Sub

ErrTrap:
    MsgBox Err.Description, vbExclamation + vbOKOnly
    GoTo Salir
End Sub

Private Sub RecuperarEmpSeleccionadas()
    Dim Cadena As String, i As Integer, j As Integer, v As Variant
       
    If cboBaseMatriz.ListCount > 0 Then
        cboBaseMatriz.ListIndex = Val(GetSetting(APPNAME, SECTION, "IndBaseMatriz", 0))
    End If
    Cadena = GetSetting(APPNAME, SECTION, "EmpresasSelec", "")
    If Len(Cadena) > 0 Then
        v = Split(Cadena, ",")
    Else
        Exit Sub
    End If
    
    'Recuperar del sistema las empresas seleccionadas para la consolidación
    For i = LBound(v) To UBound(v)
        For j = 0 To lstBase.ListCount - 1
            If lstBase.List(j) = v(i) Then
                lstBase.Selected(j) = True
                Exit For
            End If
        Next j
    Next i
End Sub

Private Sub GuardarEmpSeleccionadas()
    Dim i As Integer, Cadena As String
    
    'Guarda en registro del sistema las empresas seleccionadas para la consolidación
    For i = 0 To lstBase.ListCount - 1
        If lstBase.Selected(i) = True Then Cadena = Cadena & lstBase.List(i) & ","
    Next i
    
    If Len(Cadena) > 0 Then Cadena = Mid$(Cadena, 1, Len(Cadena) - 1)
    
    SaveSetting APPNAME, SECTION, "IndBaseMatriz", cboBaseMatriz.ListIndex
    SaveSetting APPNAME, SECTION, "EmpresasSelec", Trim$(Cadena)
End Sub

Private Function ListaEmpresas() As String
    Dim i As Long, Cadena As String
    For i = 0 To lstBase.ListCount - 1
        If lstBase.Selected(i) = True Then
            Cadena = Cadena & Trim$(lstBase.List(i)) & ","
        End If
    Next i
    If Len(Cadena) > 0 Then Cadena = Mid$(Cadena, 1, Len(Cadena) - 1) 'Quita la última coma
    ListaEmpresas = Cadena
End Function


Public Function InicioActualizaFechaIngreso(ByRef objcond As Condicion, _
                                    ByRef Recargo As String, _
                                    ByVal tag As String) As Boolean
    Dim KeyTrans As String, KeyRecargo As String
    Me.tag = tag
        
    
    fraVenta.Caption = "Transacciones"
    Label1.Visible = False
    dtpFechaCorte.Visible = False
    fraFecha.Visible = False
    FraEmpresa.Visible = True
    CargaListaEmpresas
    With objcond
    
        DTPFechaHasta1.Format = dtpCustom
        DTPFechaHasta1.CustomFormat = "dd/MM/yyyy"
        DTPFechaHasta1.value = IIf(.fecha2 = 0, Date, .fecha2)
    
       
        
        
        
        CargaTipoTrans "IV", lst

        BandAceptado = False
        If Me.tag = "FechgaIngreso" Then
            KeyTrans = "FechaTransIngreso_Trans"
        ElseIf Me.tag = "FechgaEgreso" Then
            KeyTrans = "FechaTransEgreso_Trans"
        End If
        RecuperaSelecTrans
        If Len(.Sucursal) > 0 Then
            cboBaseMatriz.Text = .Sucursal
        End If

        Me.Show vbModal, frmMain
        'Si aplastó el botón 'Aceptar'
        If BandAceptado Then
            'Devuelve los valores de condición para la búsqueda
'            .fecha1 = DTPFechaDesde.value
            .fecha2 = DTPFechaHasta1.value
            .Sucursal = cboBaseMatriz.Text
            If Me.tag = "FechgaIngreso" Then
                .CodTrans = PreparaCadenaIN(lst)
            ElseIf Me.tag = "FechgaEgreso" Then
                .CodTrans = PreparaCadenaIN(lst)
            End If
            'grabar las formas de cobro a visualizar
            SaveSetting APPNAME, App.Title, KeyTrans, .CodTrans

        End If
    End With
    'Devuelve true/false
    Unload Me
    InicioActualizaFechaIngreso = BandAceptado
End Function


Public Function InicioBuffer(ByRef objcond As Condicion, _
                                    ByRef Recargo As String, _
                                    ByVal tag As String) As Boolean
    Dim KeyTrans As String, KeyRecargo As String
    Me.tag = tag
        
    
    fraVenta.Caption = "Transacciones"
    Label1.Visible = False
    dtpFechaCorte.Visible = False
    fraFecha.Visible = True
    FraEmpresa.Visible = False
    CargaListaEmpresas
    If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "UTIL") > 0 Then
        chkProm.Visible = True
    End If
    KeyTrans = "Buffer_Trans"
    With objcond

        DTPFechaDesde.Format = dtpCustom
        DTPFechaDesde.CustomFormat = "dd/MM/yyyy"
        DTPFechaDesde.value = IIf(.fecha1 = 0, Date, .fecha1)
        
        
        DTPFechaHasta.Format = dtpCustom
        DTPFechaHasta.CustomFormat = "dd/MM/yyyy"
        DTPFechaHasta.value = IIf(.fecha2 = 0, Date, .fecha2)
    
       
        
        
        
        CargaTipoTrans "IV", lst
        BandAceptado = False
        RecuperaSelecTrans
        If Len(.Sucursal) > 0 Then
            cboBaseMatriz.Text = .Sucursal
        End If

        Me.Show vbModal, frmMain
        'Si aplastó el botón 'Aceptar'
        If BandAceptado Then
            'Devuelve los valores de condición para la búsqueda
            .fecha1 = DTPFechaDesde.value
            .fecha2 = DTPFechaHasta.value
            .CodTrans = PreparaCadenaIN(lst)
            .BandAnticipo = (chkProm.value = vbChecked)
            'grabar las formas de cobro a visualizar
            SaveSetting APPNAME, App.Title, KeyTrans, .CodTrans

        End If
    End With
    'Devuelve true/false
    Unload Me
    InicioBuffer = BandAceptado
End Function




Public Function InicioActualizaBufferxEmpresa(ByRef objcond As Condicion, _
                                    ByRef Recargo As String, _
                                    ByVal tag As String) As Boolean
    Dim KeyTrans As String, KeyRecargo As String
    Me.tag = tag
        
    
    fraVenta.Caption = "Bodegas"
    Label1.Visible = False
    dtpFechaCorte.Visible = False
    fraFecha.Visible = False
    FraEmpresa.Visible = True
    CargaListaEmpresas
    With objcond
    
    DTPFechaHasta1.Format = dtpCustom
    DTPFechaHasta1.CustomFormat = "dd/MM/yyyy"
    DTPFechaHasta1.value = IIf(.fecha2 = 0, Date, .fecha2)
    
       
        
        
        
        CargaBodegas lst

        BandAceptado = False
        KeyTrans = "BufferxAlm_CodBodega"
        RecuperaSelecTrans
        If Len(.Sucursal) > 0 Then
            cboBaseMatriz.Text = .Sucursal
        End If

        Me.Show vbModal, frmMain
        'Si aplastó el botón 'Aceptar'
        If BandAceptado Then
            'Devuelve los valores de condición para la búsqueda
'            .fecha1 = DTPFechaDesde.value
            .fecha2 = DTPFechaHasta1.value
            .Sucursal = cboBaseMatriz.Text
            .CodTrans = PreparaCadenaIN(lst)
            'grabar las formas de cobro a visualizar
            SaveSetting APPNAME, App.Title, KeyTrans, .CodTrans

        End If
    End With
    'Devuelve true/false
    Unload Me
    InicioActualizaBufferxEmpresa = BandAceptado
End Function


Private Sub CargaBodegas(ByRef lst As ListBox)
    Dim rs As Recordset, Vector As Variant
    Dim numMod As Integer, i As Integer
    'Prepara la lista de tipos de transaccion
    lst.Clear
        Set rs = gobjMain.EmpresaActual.ListaIVBodega(True, True)
        With rs
            If Not (.EOF) Then
                .MoveFirst
                Do Until .EOF
                    lst.AddItem !CodBodega & "  " & !Descripcion
                    lst.ItemData(lst.NewIndex) = Len(!CodBodega)
                    .MoveNext
                Loop
            End If
        End With
        rs.Close
    Set rs = Nothing
End Sub

