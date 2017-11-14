VERSION 5.00
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Begin VB.Form frmB_FiltroxCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtros por Cliente"
   ClientHeight    =   5790
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5595
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   372
      Left            =   2880
      TabIndex        =   1
      Top             =   5340
      Width           =   1092
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar -F5"
      Height          =   372
      Left            =   1680
      TabIndex        =   0
      Top             =   5340
      Width           =   1092
   End
   Begin VB.Frame FraGrupos 
      Caption         =   "Grupos"
      Height          =   5295
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   5475
      Begin VB.Frame fraCliente 
         Caption         =   "Cliente"
         Height          =   900
         Left            =   60
         TabIndex        =   18
         Top             =   4260
         Width           =   5355
         Begin FlexComboProy.FlexCombo fcbHasta 
            Height          =   315
            Left            =   2700
            TabIndex        =   19
            Top             =   480
            Width           =   2600
            _ExtentX        =   4577
            _ExtentY        =   556
            ColWidth1       =   3400
            ColWidth2       =   3400
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
         Begin FlexComboProy.FlexCombo fcbDesde 
            Height          =   315
            Left            =   60
            TabIndex        =   20
            Top             =   480
            Width           =   2600
            _ExtentX        =   4577
            _ExtentY        =   556
            ColWidth1       =   3400
            ColWidth2       =   3400
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
         Begin VB.Label Label1 
            Caption         =   "&Desde:"
            Height          =   255
            Left            =   60
            TabIndex        =   22
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "&Hasta"
            Height          =   255
            Left            =   2700
            TabIndex        =   21
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.ListBox lstGrupo 
         Height          =   960
         Index           =   3
         Left            =   1020
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   3240
         Width           =   4335
      End
      Begin VB.ListBox lstGrupo 
         Height          =   960
         Index           =   2
         Left            =   1020
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   2250
         Width           =   4335
      End
      Begin VB.ListBox lstGrupo 
         Height          =   960
         Index           =   1
         Left            =   1020
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   1260
         Width           =   4335
      End
      Begin VB.ListBox lstGrupo 
         Height          =   960
         Index           =   0
         Left            =   1020
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   285
         Width           =   4335
      End
      Begin VB.Label lbl1 
         Caption         =   "Label1"
         Height          =   255
         Index           =   3
         Left            =   180
         TabIndex        =   10
         Top             =   3285
         Width           =   615
      End
      Begin VB.Label lbl1 
         Caption         =   "Label1"
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Top             =   2250
         Width           =   615
      End
      Begin VB.Label lbl1 
         Caption         =   "Label1"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lbl1 
         Caption         =   "Label1"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame FraProvincia 
      Caption         =   "Provincias / Cantones"
      Height          =   5205
      Left            =   60
      TabIndex        =   11
      Top             =   0
      Width           =   5475
      Begin VB.ListBox lstCanton 
         Height          =   1410
         Left            =   840
         Style           =   1  'Checkbox
         TabIndex        =   14
         Top             =   2100
         Width           =   4515
      End
      Begin VB.ListBox lstParroquia 
         Height          =   1410
         Left            =   840
         Style           =   1  'Checkbox
         TabIndex        =   13
         Top             =   3540
         Width           =   4515
      End
      Begin VB.ListBox lstProv 
         Height          =   1860
         Left            =   840
         Style           =   1  'Checkbox
         TabIndex        =   12
         Top             =   240
         Width           =   4515
      End
      Begin VB.Label lbl1 
         Caption         =   "Parroquia"
         Height          =   255
         Index           =   7
         Left            =   60
         TabIndex        =   17
         Top             =   3540
         Width           =   975
      End
      Begin VB.Label lbl1 
         Caption         =   "Cantón"
         Height          =   255
         Index           =   6
         Left            =   60
         TabIndex        =   16
         Top             =   2130
         Width           =   675
      End
      Begin VB.Label lbl1 
         Caption         =   "Provincia"
         Height          =   255
         Index           =   5
         Left            =   60
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame FraPC 
      Caption         =   "Clientes"
      Height          =   5175
      Left            =   60
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   5475
      Begin VB.ListBox lstPC 
         Height          =   4785
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   24
         Top             =   240
         Width           =   5235
      End
   End
End
Attribute VB_Name = "frmB_FiltroxCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BandAceptado As Boolean
Private Const numGrupo = 4
Private sql As String
Private Key As String
Private KeyProv As String
Private Const SeparaLista = 50
Private trans As String
Private BandProv As Boolean

Public Function Inicio(ByVal tag As String, bandCompra As Boolean) As String
    Dim i As Integer
    Me.tag = tag  'nombre del reporte
    MensajeStatus MSG_PREPARA, vbHourglass
    FraProvincia.Visible = False
    FraGrupos.Visible = True
    If bandCompra Then
        Me.Caption = "Filtros por Proveedor"
        CargaClientes (False)
        fraCliente.Caption = "Proveedor"
        CargaGrupos False, True, False
        
    Else
        Me.Caption = "Filtros por Cliente"
        CargaClientes (True)
        fraCliente.Caption = "Cliente"
        CargaGrupos True, False, False
    End If

    With gobjMain.objCondicion
        fcbDesde.Text = .CodPC1
        fcbHasta.Text = .CodPC2
        MensajeStatus
        Me.Show vbModal
        .CodPC1 = fcbDesde.KeyText
        .CodPC2 = fcbHasta.KeyText
        If BandAceptado Then '"Ha aceptado, pero debemos armar cadena"
            .Bienes = ArmarSqlCliente  'Aqui  guarda  los IDs  en el  registor de Windows
'            .Servicios = ArmarSqlClienteP   'Aqui  guarda  los IDs  en el  registor de Windows
'            .CodComp = ArmarSqlxCliente
'            .SQLProvCliente = ArmarSqlProvCliente  'Aqui  guarda  los IDs  en el  registor de Windows
'            .SQLCantCliente = ArmarSqlCantCliente
                Inicio = ArmarEtiqueta
            If bandCompra Then
                If Len(.CodPC1) > 0 Or Len(.CodPC2) > 0 Then
                    Inicio = Inicio & "Y proveedores desde : " & .CodPC1 & " hasta " & .CodPC2
                End If
            Else
                If Len(.CodPC1) > 0 Or Len(.CodPC2) > 0 Then
                    Inicio = Inicio & "Y clientes desde : " & .CodPC1 & " hasta " & .CodPC2
                End If
            End If
            'guarda  solo  la cadena
       Else
            Inicio = "Cancelar"
        End If
    End With
    Unload Me
End Function


Private Sub cmdAceptar_Click()
    BandAceptado = True
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    BandAceptado = False
    sql = "Ha cancelado, no necesitamos armar cadena"
    Me.Hide
End Sub

Private Sub fcbDesde_Selected(ByVal Text As String, ByVal KeyText As String)
        fcbHasta.KeyText = fcbDesde.KeyText
End Sub

Private Sub fcbDesdeG_Selected(ByVal Text As String, ByVal KeyText As String)
'fcbHastaG.KeyText = KeyText
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF5
        cmdAceptar_Click
        KeyCode = 0
    Case vbKeyEscape
        cmdCancelar_Click
        KeyCode = 0
    Case Else
        MoverCampo Me, KeyCode, Shift, False
    End Select
End Sub

Private Sub Form_Load()
'    lbl1(0).Caption = gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupo(1)
'    lbl1(1).Caption = gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupo(2)
'    lbl1(2).Caption = gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupo(3)
'    lbl1(3).Caption = gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupo(4)
'    'f BandProv Then
'        CargaListaGrupos
'        CargaProvincias
'    'End If
End Sub

Private Sub CargaClientes(ByRef BandCliente As Boolean)
    Dim v() As Variant
    Dim sql  As String, rs As Recordset, cond As String
    fcbDesde.Clear
    fcbHasta.Clear
    If BandCliente Then
    
        cond = " WHERE bandCliente = 1"
        If Me.tag = "VxItemVen" Then
            cond = cond & ""
        End If
    Else
        cond = " WHERE BandProveedor = 1"
    End If
    sql = "SELECT CodProvCli, Nombre, NombreAlterno FROM PCProvCli "
    sql = sql & cond
    sql = sql & " ORDER BY Nombre"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not rs.EOF Then
        v = MiGetRows(rs)
        fcbDesde.SetData v
        fcbHasta.SetData v
    End If
    fcbDesde.Text = ""
    fcbHasta.Text = ""
End Sub

Private Sub CargaListaGrupos(ByVal bandCli As Boolean, ByVal BandProv As Boolean, BandEmp As Boolean)
    Dim i As Long
    Dim sql  As String, rs As Recordset, cond As String
    For i = 1 To numGrupo
        sql = "SELECT CodGrupo" & i & " as Codgrupo, Descripcion, IDGrupo" & i & " as IDgrupo FROM PCGrupo" & i & " "
        If bandCli Then
            sql = sql & "where origen = 2 "
        ElseIf BandProv Then
            sql = sql & "where origen = 1 "
        Else
            sql = sql & "where origen = 4 "
        End If
        
        sql = sql & "ORDER BY CodGrupo" & i
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
        Do While Not rs.EOF
            lstGrupo(i - 1).AddItem rs!Descripcion & Space(SeparaLista) & " [" & rs!idgrupo & "]"
            lstGrupo(i - 1).ItemData(lstGrupo(i - 1).NewIndex) = Len(rs!CodGrupo)
            rs.MoveNext
        Loop
        Set rs = Nothing
    Next i
End Sub

Private Function ArmarSqlCliente() As String
    Dim i As Long, s As String, codigos As String
    s = ""
    For i = 1 To numGrupo
        codigos = PreparaListaGrupo(lstGrupo(i - 1))
        If Len(codigos) > 0 Then
            s = s & " vwPCProvCli.IdGrupo" & i & " in (" & codigos & ") AND "
        End If

           
    Next i
    If Len(s) > 0 Then
        If Mid(s, Len(s) - 3, 4) = "AND " Then s = Mid(s, 1, Len(s) - 4)
        ArmarSqlCliente = "AND (" & s & ")"
    End If
End Function

Private Function PreparaListaGrupo(ByVal lst As ListBox) As String
    Dim i As Long, s As String
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) = True Then s = s & "'" & CogeSoloCodigo(lst.List(i)) & "',"
    Next i
    If Len(s) > 0 Then
        If Mid(s, Len(s), 1) = "," Then s = Mid(s, 1, Len(s) - 1)
    End If
    PreparaListaGrupo = s
End Function

Public Function CogeSoloCodigo(Desc As String) As String
    Dim s As String, i As Long
    i = InStrRev(Desc, "[")
    If i > 0 Then s = Mid$(Desc, i + 1)
    If Len(s) > 0 Then s = Left$(s, Len(s) - 1)
    CogeSoloCodigo = s
End Function

Private Sub RecuperaDatos()
    Dim i As Long
    For i = 1 To numGrupo
        RecuperaGrupoSelec Key & i, lstGrupo(i - 1), trans
        'RecuperaGrupoSelec Key, lst, trans
    Next i
End Sub


Public Sub RecuperaGrupoSelec(ByVal Key As String, lst As ListBox, s As String)
Dim Vector As Variant
Dim i As Integer, j As Integer, Selec As Integer



    If s <> "_VACIO_" Then
        Vector = Split(s, ",")
         Selec = UBound(Vector, 1)
         For i = 0 To Selec
            For j = 0 To lst.ListCount - 1
                If Vector(i) = "'" & CogeSoloCodigo(lst.List(j)) & "'" Then
                    lst.Selected(j) = True
                End If
            Next j
         Next i
    End If
End Sub

Private Function ArmarEtiqueta() As String
    Dim i As Long, s As String, lst As ListBox, X As Long, cod As String
    
    For X = 1 To numGrupo
        Set lst = lstGrupo(X - 1)
        For i = 0 To lst.ListCount - 1
            If lst.Selected(i) = True Then
                cod = lst.List(i)
                s = s & Mid(cod, 1, Len(cod) - (SeparaLista + 3 + Len(CogeSoloCodigo(cod)))) & ","
            End If
        Next i
        If Len(s) > 0 Then
            If Mid(s, Len(s), 1) = "," Then s = Mid(s, 1, Len(s) - 1) '& vbCrLf
'            s = lbl1(X - 1).Caption & ": " & s
        End If
        Set lst = Nothing
        ArmarEtiqueta = ArmarEtiqueta & s & IIf(Len(s) > 0, "; ", "")
        s = ""
    Next X
'    ArmarEtiqueta = "Clientes de " & vbCrLf & ArmarEtiqueta
End Function

Private Sub CargaClientesxVen(ByRef BandCliente As Boolean)
    Dim v() As Variant
    Dim sql  As String, rs As Recordset, cond As String
    fcbDesde.Clear
    fcbHasta.Clear
    If BandCliente Then
        cond = " WHERE bandCliente = 1 AND v.codvendedor = '" & gobjMain.UsuarioActual.codUsuario & "'"
    Else
        cond = " WHERE BandProveedor = 1"
    End If
    sql = "SELECT CodProvCli,pc.Nombre FROM PCProvCli pc "
    sql = sql & "INNER JOIN FCVendedor v on v.idvendedor =  pc.idvendedor "
    sql = sql & cond
    sql = sql & " ORDER BY pc.nombre"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not rs.EOF Then
        v = MiGetRows(rs)
        fcbDesde.SetData v
        fcbHasta.SetData v
    End If
    fcbDesde.Text = ""
    fcbHasta.Text = ""
End Sub

Private Sub CargaProvincias()
    Dim i As Long
    Dim sql  As String, rs As Recordset, cond As String

        sql = "SELECT Codprovincia, Descripcion,  IDProvincia FROM PcProvincia  ORDER BY Descripcion"
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
        Do While Not rs.EOF
            lstProv.AddItem rs!Descripcion & Space(SeparaLista) & " [" & rs!IdProvincia & "]"
            lstProv.ItemData(lstProv.NewIndex) = Len(rs!codProvincia)

            rs.MoveNext
        Loop
        Set rs = Nothing

End Sub

Private Sub CargaCanton(ByVal provin As String)
    Dim i As Long
    Dim sql  As String, rs As Recordset, cond As String
'    For i = 1 To numGrupo
        sql = "SELECT CodCanton, Descripcion,  IDCanton FROM PcCanton  "
        If Len(provin) > 0 Then
            sql = sql & " where idprovincia in (" & provin & ")"
        End If
        sql = sql & " ORDER BY Descripcion"
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
        lstCanton.Clear
        Do While Not rs.EOF
'            lstCanton.AddItem rs!CodCanton & " " & Left(rs!Descripcion, 20) & " " & " [" & rs!IdCanton & "]"
'            lstCanton.ItemData(lstCanton.NewIndex) = Len(rs!CodCanton)
            
            lstCanton.AddItem rs!Descripcion & Space(SeparaLista) & " [" & rs!Idcanton & "]"
            lstCanton.ItemData(lstCanton.NewIndex) = Len(rs!codCanton)

            
            rs.MoveNext
        Loop
        Set rs = Nothing
'    Next i
End Sub

    Private Sub CargaParroquia(ByVal Canton As String)
        Dim i As Long
        Dim sql  As String, rs As Recordset, cond As String
            sql = "SELECT CodParroquia, Descripcion,  IDParroquia FROM PcParroquia "
            If Len(Canton) > 0 Then
                sql = sql & " where idcanton in (" & Canton & ")"
            End If
            sql = sql & " ORDER BY Descripcion"
            lstParroquia.Clear
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            Do While Not rs.EOF
                'lstParroquia.AddItem rs!CodParroquia & " " & Left(rs!Descripcion, 20) & " " & " [" & rs!IDParroquia & "]"
                'lstParroquia.ItemData(lstParroquia.NewIndex) = Len(rs!CodParroquia)
                lstParroquia.AddItem rs!Descripcion & Space(SeparaLista) & " [" & rs!IdParroquia & "]"
                lstParroquia.ItemData(lstParroquia.NewIndex) = Len(rs!codParroquia)
                
                rs.MoveNext
            Loop
            Set rs = Nothing
    End Sub


Private Sub List1_Click()

End Sub

Private Sub lstCanton_Click()
  Dim cant As String, i As Integer, cod As String, s As String
    Dim codigos As String
    codigos = PreparaListaGrupo(lstCanton)
    CargaParroquia codigos
End Sub

Private Sub lstProv_Click()
    Dim cant As String, i As Integer, cod As String, s As String
    Dim codigos As String
    codigos = PreparaListaGrupo(lstProv)
    CargaCanton codigos
End Sub

Private Function ArmarSqlProvCliente() As String
    Dim i As Long, s As String, codigos As String, codigosCant As String, codigosParr As String
        
        codigos = PreparaListaGrupo(lstProv)
        
        If Len(codigos) > 0 Then
            s = s & " PCProvCli.IdProvincia" & " in (" & codigos & ") AND "
        End If
        
        codigosCant = PreparaListaGrupo(lstCanton)
        
        If Len(codigosCant) > 0 Then
            s = s & " PCProvCli.IdCanton" & " in (" & codigosCant & ") AND "
        End If
        
        codigosParr = PreparaListaGrupo(lstParroquia)
        If Len(codigosParr) > 0 Then
            s = s & " PCProvCli.Idparroquia" & " in (" & codigosParr & ") AND "
        End If
        

    If Len(s) > 0 Then
        If Mid(s, Len(s) - 3, 4) = "AND " Then s = Mid(s, 1, Len(s) - 4)
        ArmarSqlProvCliente = "AND (" & s & ")"
    End If
End Function

Private Function ArmarSqlCantCliente() As String
    Dim i As Long, s As String, codigos As String
        codigos = PreparaListaGrupo(lstCanton)
        If Len(codigos) > 0 Then
            s = s & " PCProvCli.IdCanton" & i & " in (" & codigos & ") AND "
        End If


    If Len(s) > 0 Then
        If Mid(s, Len(s) - 3, 4) = "AND " Then s = Mid(s, 1, Len(s) - 4)
        ArmarSqlCantCliente = "AND (" & s & ")"
    End If
End Function


Public Function InicioNuevo(ByVal tag As String, ByRef objcond As RepCondicion, bandCompra As Boolean, BandProvincia As Boolean) As String
    Dim i As Integer

    BandProv = BandProvincia
    Me.tag = tag  'nombre del reporte

    MensajeStatus MSG_PREPARA, vbHourglass
    BandProv = BandProvincia
    FraProvincia.Visible = BandProvincia
    FraGrupos.Visible = Not BandProvincia
    If bandCompra Then
        Me.Caption = "Filtros por Proveedor"
    End If
    Select Case tag
    Case "VxVolxZona" 'Reporte específico para una empresa
        lstGrupo(1).Enabled = False
        lstGrupo(2).Enabled = False
        Key = "I_FiltroxCli_VVZ"
        CargaClientes (True)
        fraCliente.Caption = "Cliente"
        CargaGrupos True, False, False
        
    Case "ConsIVVentaCat"
        Key = "I_FiltroxCli_IVC"
        CargaClientes (True)
        fraCliente.Caption = "Cliente"
        CargaGrupos True, False, False
    Case "ConsIVVentaDescCat", "ConsIVVentaDescCat4"
        Key = "I_FiltroxCli_IVC"
        CargaClientes (True)
        fraCliente.Caption = "Cliente"
        CargaGrupos True, False, False
    Case "Lista_Cli"
        Key = "I_FiltroxCli_LC"   'Para reporte de LIsta Clientes
        CargaClientes (True)
        fraCliente.Caption = "Cliente"
        CargaGrupos True, False, False
    Case "Lista_Prov", "ConsCompraCar"
        Key = "I_FiltroxCli_LP"   'Para reporte de LIsta Proveedores
        CargaClientes (False)
        fraCliente.Caption = "Proveedor"
        CargaGrupos False, True, False
    Case "Cartera_Cliente"
        Key = "I_FiltroxCli_CC"   'Para reporte de CARTERA LIsta Clientes
        If gobjMain.GrupoActual.PermisoActual.ConsRepVen Then
            CargaClientesxVen (True)
        Else
            CargaClientes (True)
        End If
        CargaGrupos True, False, False
        fraCliente.Caption = "Cliente"
        Me.Caption = "Filtros por Proveedor"
    Case "Cartera_Prov", "ConsIVPedP"
        Key = "I_FiltroxCli_CP"   'Para reporte de CARTERA LIsta Proveedores
        CargaClientes (False)
        fraCliente.Caption = "Proveedor"
        Me.Caption = "Filtros por Proveedor"
        CargaGrupos False, True, False
    Case "CarteraFC_Cliente", "frmB_CXVD"
        Key = "I_FiltroxCli_CCFC"   'Para reporte de CARTERA LIsta Clientes FECHA cORTE
        If gobjMain.GrupoActual.PermisoActual.ConsRepVen Then
            CargaClientesxVen (True)
        Else
            CargaClientes (True)
        End If
        fraCliente.Caption = "Cliente"
        CargaGrupos True, False, False
    Case "CarteraFC_Prov"
        Key = "I_FiltroxCli_CPFC"   'Para reporte de CARTERA LIsta Proveedores FECHACORTE
        fraCliente.Caption = "Proveedores"
        CargaClientes (False)
        Me.Caption = "Filtros por Proveedor"
        CargaGrupos False, True, False
    Case "CarteraDV_Cliente"
        Key = "I_FiltroxCli_CCDV"   'Para reporte de CARTERA LIsta Clientes DIAS vencidos
        If gobjMain.GrupoActual.PermisoActual.ConsRepVen Then
            CargaClientesxVen (True)
        Else
            CargaClientes (True)
        End If
        fraCliente.Caption = "Cliente"
        CargaGrupos True, False, False
    Case "CarteraDV_Prov"
        Key = "I_FiltroxCli_CPDV"   'Para reporte de CARTERA LIsta Proveedores dis vencidos
        CargaClientes (False)
        fraCliente.Caption = "Proveedores"
        Me.Caption = "Filtros por Proveedor"
        CargaGrupos False, True, False
    Case "CarteraDVR_Cliente"
        Key = "I_FiltroxCli_CCDVR"   'Para reporte de CARTERA LIsta Clientes DIAS vencidos
        CargaClientes (True)
        fraCliente.Caption = "Cliente"
        CargaGrupos True, False, False
    Case "CxItem"
        If bandCompra Then
            Key = "I_CxItem_Trans"   'Para reporte de compras x mes
            fraCliente.Caption = "Proveedores"
            CargaClientes (False)
            Me.Caption = "Filtros por Proveedor"
            CargaGrupos False, True, False
        Else
            Key = "I_VxItem_Trans"   'Para reporte de compras x mes
            fraCliente.Caption = "Clientes"
            CargaClientes (True)
            CargaGrupos True, False, False
        End If
    Case "TotalPagos", "TotalPagosCAO"
        If bandCompra Then
            Key = "TotalPagosP"   'Para reporte de compras x mes
            fraCliente.Caption = "Proveedores"
            CargaClientes (False)
            CargaGrupos False, True, False
        Else
            Key = "TotalPagosC"   'Para reporte de compras x mes
            fraCliente.Caption = "Clientes"
            If gobjMain.GrupoActual.PermisoActual.ConsRepVen Then
                CargaClientesxVen (True)
            Else
                CargaClientes (True)
            End If
            CargaGrupos True, False, False
        End If
    Case "VxGeneralFranq"
        Key = "Etiq_VxItemFranq_Cli"
        fraCliente.Caption = "Clientes"
        CargaClientes (True)
        CargaGrupos True, False, False
    Case "CarterayChequesCons"
            fraCliente.Caption = "Proveedores"
            CargaClientes (False)
            CargaGrupos False, True, False
    Case "CarteraDV_ClienteH"
        Key = "Etiq_CarteraDV_ClienteH"
        fraCliente.Caption = "Clientes"
        CargaClientes (True)
        CargaGrupos True, False, False
    Case Else
        Key = "I_FiltroxCli_G"   'Para reporte de ventas General
        If gobjMain.GrupoActual.PermisoActual.ConsRepVen Then
            CargaClientesxVen (True)
        Else
            CargaClientes (True)
        End If
        fraCliente.Caption = "Cliente"
        CargaGrupos True, False, False
    End Select
    'trans = mobjReporte.RecuperarConfigBusqueda(Me.Tag, Key)
    
    RecuperaDatos
    With objcond
        fcbDesde.Text = .Cliente1
        fcbHasta.Text = .Cliente2
        MensajeStatus
        Me.Show vbModal
        .Cliente1 = fcbDesde.Text
        .Cliente2 = fcbHasta.Text
        If BandAceptado Then '"Ha aceptado, pero debemos armar cadena"
            If Not BandProvincia Then
                .SQLCliente = ArmarSqlCliente  'Aqui  guarda  los IDs  en el  registor de Windows
                .SQLClienteP = ArmarSqlClienteP   'Para los paretos

                InicioNuevo = ArmarEtiqueta
                If bandCompra Then
                    If Len(.Cliente1) > 0 Or Len(.Cliente2) > 0 Then
                        InicioNuevo = InicioNuevo & "Y proveedores desde : " & .Cliente1 & " hasta " & .Cliente2
                    End If
                Else
                    If Len(.Cliente1) > 0 Or Len(.Cliente2) > 0 Then
                        InicioNuevo = InicioNuevo & "Y clientes desde : " & .Cliente1 & " hasta " & .Cliente2
                    End If
                End If
            Else
                .SQLProvCliente = ArmarSqlProvCliente  'Aqui  guarda  los IDs  en el  registor de Windows
                InicioNuevo = ArmarEtiquetaProv
            End If
            
            
       Else
            InicioNuevo = "Cancelar"
        End If
    End With
    Unload Me
End Function

Private Function ArmarEtiquetaProv() As String
    Dim i As Long, s As String, lst As ListBox, X As Long, cod As String
    
'    For X = 1 To numGrupo
        Set lst = lstProv
        For i = 0 To lst.ListCount - 1
            If lst.Selected(i) = True Then
                cod = lst.List(i)
                s = s & Mid(cod, 1, Len(cod) - (SeparaLista + 3 + Len(CogeSoloCodigo(cod)))) & ","
            End If
        Next i
        
        Set lst = lstCanton
        For i = 0 To lst.ListCount - 1
            If lst.Selected(i) = True Then
                cod = lst.List(i)
                s = s & Mid(cod, 1, Len(cod) - (SeparaLista + 3 + Len(CogeSoloCodigo(cod)))) & ","
            End If
        Next i
        
        Set lst = lstParroquia
        For i = 0 To lst.ListCount - 1
            If lst.Selected(i) = True Then
                cod = lst.List(i)
                s = s & Mid(cod, 1, Len(cod) - (SeparaLista + 3 + Len(CogeSoloCodigo(cod)))) & ","
            End If
        Next i

        
        If Len(s) > 0 Then
            If Mid(s, Len(s), 1) = "," Then s = Mid(s, 1, Len(s) - 1) '& vbCrLf
'            s = lbl1(X - 1).Caption & ": " & s
        End If
        Set lst = Nothing
        ArmarEtiquetaProv = ArmarEtiqueta & s & IIf(Len(s) > 0, "; ", "")
        s = ""
'    Next X
'    ArmarEtiqueta = "Clientes de " & vbCrLf & ArmarEtiqueta
End Function


Private Function ArmarSqlClienteP() As String
    Dim i As Long, s As String, codigos As String
    For i = 1 To numGrupo
        codigos = PreparaListaGrupo(lstGrupo(i - 1))
        If Len(codigos) > 0 Then
            s = s & " vwConsVxParetos.IdGrupo" & i & " in (" & codigos & ") AND "
        End If

        Next i
    If Len(s) > 0 Then
        If Mid(s, Len(s) - 3, 4) = "AND " Then s = Mid(s, 1, Len(s) - 4)
        ArmarSqlClienteP = "AND (" & s & ")"
    End If
End Function

Private Sub CargaGrupos(ByVal BandCliente As Boolean, ByVal BandProv As Boolean, ByVal BandEmp As Boolean)
    If BandCliente Then
        lbl1(0).Caption = gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupoC(1)
        lbl1(1).Caption = gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupoC(2)
        lbl1(2).Caption = gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupoC(3)
        lbl1(3).Caption = gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupoC(4)
        CargaListaGrupos True, False, False
    ElseIf BandProv Then
        lbl1(0).Caption = gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupoP(1)
        lbl1(1).Caption = gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupoP(2)
        lbl1(2).Caption = gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupoP(3)
        lbl1(3).Caption = gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupoP(4)
        CargaListaGrupos False, True, False
    Else
        lbl1(0).Caption = gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupoE(1)
        lbl1(1).Caption = gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupoE(2)
        lbl1(2).Caption = gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupoE(3)
        lbl1(3).Caption = gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupoE(4)
        CargaListaGrupos False, False, True
    End If
        CargaProvincias
        
End Sub

Private Sub CargaLSTClientes()
    Dim i As Long
    Dim sql  As String, rs As Recordset, cond As String

        sql = "SELECT RUC, Nombre,  IDProvcli FROM PcProvcli where bandcliente=1  ORDER BY Nombre"
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
        
        Do While Not rs.EOF
            lstPC.AddItem rs!nombre & Space(SeparaLista) & " [" & rs!IdProvCli & "]"
            lstPC.ItemData(lstPC.NewIndex) = Len(rs!ruc)

            rs.MoveNext
        Loop
        Set rs = Nothing

End Sub

Private Function ArmarSqlxCliente() As String
    Dim i As Long, s As String, codigos As String
    For i = 1 To 1
        codigos = PreparaListaGrupo(lstPC)
        If Len(codigos) > 0 Then
            s = s & " pcFac.IdProvcli  in (" & codigos & ") AND "
        End If

        Next i
    If Len(s) > 0 Then
        If Mid(s, Len(s) - 3, 4) = "AND " Then s = Mid(s, 1, Len(s) - 4)
        ArmarSqlxCliente = "AND (" & s & ")"
    End If
End Function

Private Function ArmarEtiquetaxCli() As String
    Dim i As Long, s As String, lst As ListBox, X As Long, cod As String
    
    For X = 1 To 1
        Set lst = lstPC
        For i = 0 To lst.ListCount - 1
            If lst.Selected(i) = True Then
                cod = lst.List(i)
                s = s & Mid(cod, 1, Len(cod) - (SeparaLista + 3 + Len(CogeSoloCodigo(cod)))) & ","
            End If
        Next i
        If Len(s) > 0 Then
            If Mid(s, Len(s), 1) = "," Then s = Mid(s, 1, Len(s) - 1) '& vbCrLf
'            s = lbl1(X - 1).Caption & ": " & s
        End If
        Set lst = Nothing
        ArmarEtiquetaxCli = ArmarEtiquetaxCli & s & IIf(Len(s) > 0, "; ", "")
        s = ""
    Next X
'    ArmarEtiqueta = "Clientes de " & vbCrLf & ArmarEtiqueta
End Function


Private Sub RecuperaDatosxCli()
    Dim i As Long

        RecuperaGrupoSelec Key, lstPC, trans
        'RecuperaGrupoSelec Key, lst, trans
End Sub

