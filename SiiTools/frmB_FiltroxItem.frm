VERSION 5.00
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Begin VB.Form frmB_FiltroxItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtro por Item"
   ClientHeight    =   6555
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4590
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstGrupo 
      Height          =   960
      Index           =   4
      Left            =   840
      Style           =   1  'Checkbox
      TabIndex        =   16
      Top             =   4140
      Width           =   3495
   End
   Begin VB.ListBox lstGrupo 
      Height          =   960
      Index           =   3
      Left            =   840
      Style           =   1  'Checkbox
      TabIndex        =   14
      Top             =   3120
      Width           =   3495
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   372
      Left            =   2400
      TabIndex        =   12
      Top             =   6000
      Width           =   1092
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   372
      Left            =   1080
      TabIndex        =   11
      Top             =   6000
      Width           =   1092
   End
   Begin VB.Frame fraItem 
      Caption         =   "Items"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   4335
      Begin FlexComboProy.FlexCombo fcbHasta 
         Height          =   252
         Left            =   2760
         TabIndex        =   7
         Top             =   240
         Width           =   1332
         _ExtentX        =   2355
         _ExtentY        =   450
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
      Begin FlexComboProy.FlexCombo fcbDesde 
         Height          =   252
         Left            =   720
         TabIndex        =   8
         Top             =   240
         Width           =   1332
         _ExtentX        =   2355
         _ExtentY        =   450
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
      Begin VB.Label Label1 
         Caption         =   "&Desde:"
         Height          =   252
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   612
      End
      Begin VB.Label Label2 
         Caption         =   "&Hasta"
         Height          =   252
         Left            =   2160
         TabIndex        =   9
         Top             =   240
         Width           =   492
      End
   End
   Begin VB.ListBox lstGrupo 
      Height          =   960
      Index           =   2
      Left            =   840
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   2100
      Width           =   3495
   End
   Begin VB.ListBox lstGrupo 
      Height          =   960
      Index           =   1
      Left            =   840
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   1080
      Width           =   3495
   End
   Begin VB.ListBox lstGrupo 
      Height          =   960
      Index           =   0
      Left            =   840
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   60
      Width           =   3495
   End
   Begin VB.Label lbl1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   15
      Top             =   4140
      Width           =   615
   End
   Begin VB.Label lbl1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label lbl1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   2100
      Width           =   615
   End
   Begin VB.Label lbl1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lbl1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   495
   End
End
Attribute VB_Name = "frmB_FiltroxItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BandAceptado As Boolean
Private Const numGrupo = 5
Private sql As String
Private key As String
Private Const SeparaLista = 50
Private trans As String

Public Function Inicio(ByVal tag As String, ByVal KeyItm As String) As String
    Dim i As Integer, s As String
    Me.tag = tag
    
    MensajeStatus MSG_PREPARA, vbHourglass
    
   
   If Me.tag = "BufferGYP" Then
        For i = 1 To numGrupo
            s = ""
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Pedidos_BufferGYPConsTipoTrans_ItemG_" & i)) > 0 Then
                s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Pedidos_BufferGYPConsTipoTrans_ItemG_" & i)
            End If
    
            RecuperaGrupoSelec key & i, lstGrupo(i - 1), s
        Next i
    ElseIf Me.tag = "IVPareto" Then
        For i = 1 To numGrupo
            s = ""
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("IVPareto_ItemG_" & i)) > 0 Then
                s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("IVPareto_ItemG_" & i)
            End If
    
            RecuperaGrupoSelec key & i, lstGrupo(i - 1), s
        Next i
    ElseIf Me.tag = "VentasxSuc" Then
    Else
        For i = 1 To numGrupo
            s = ""
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Pedidos_BufferConsTipoTrans_ItemG_" & i)) > 0 Then
                s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Pedidos_BufferConsTipoTrans_ItemG_" & i)
            End If
    
            RecuperaGrupoSelec key & i, lstGrupo(i - 1), s
        Next i
    End If
    
   'RecuperaDatos
    With gobjMain.objCondicion
        fcbDesde.Text = .CodItem1
        fcbHasta.Text = .CodItem2
        MensajeStatus
        Me.Show vbModal
        .CodItem1 = fcbDesde.Text
        .CodItem2 = fcbHasta.Text
        If BandAceptado Then
            .Bienes = ArmarSqlItem                           '"Ha aceptado, pero debemos armar cadena"
            .CodVehiculo = ArmarSqlItemP
            Inicio = ArmarEtiqueta
            If Len(.CodItem1) > 0 Or Len(.CodItem2) > 0 Then
                Inicio = Inicio & "e ítems desde : " & .CodItem1 & " hasta " & .CodItem2
            End If
       Else

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
    Me.Hide
End Sub

Private Sub fcbDesde_Selected(ByVal Text As String, ByVal KeyText As String)
    fcbHasta.KeyText = fcbDesde.KeyText
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
    lbl1(0).Caption = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(1)
    lbl1(1).Caption = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(2)
    lbl1(2).Caption = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(3)
    lbl1(3).Caption = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(4)
    lbl1(4).Caption = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(5)
    CargaItems
    CargaListaGruposNew
End Sub

Private Sub CargaItems()
    Dim v() As Variant
    Dim sql  As String, rs As Recordset, cond As String
    fcbDesde.Clear
    fcbHasta.Clear
    sql = "SELECT CodInventario, Descripcion FROM IVInventario ORDER BY CodInventario"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not rs.EOF Then
        v = MiGetRows(rs)
        fcbDesde.SetData v
        fcbHasta.SetData v
    End If
    fcbDesde.Text = ""
    fcbHasta.Text = ""
End Sub

Private Sub CargaListaGrupos()
    Dim i As Long
    Dim sql  As String, rs As Recordset, cond As String
    For i = 1 To numGrupo
        sql = "SELECT CodGrupo" & i & " as Codgrupo, Descripcion,  IDGrupo" & i & " as IDgrupo FROM IVGrupo" & i & " ORDER BY CodGrupo" & i
        
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
        Do While Not rs.EOF
            lstGrupo(i - 1).AddItem rs!CodGrupo & " " & Left(rs!Descripcion, 20) & " " & " [" & rs!idgrupo & "]"
            lstGrupo(i - 1).ItemData(lstGrupo(i - 1).NewIndex) = Len(rs!CodGrupo)
            rs.MoveNext
        Loop
        Set rs = Nothing
    Next i
End Sub

Private Function ArmarSqlItem() As String
    Dim i As Long, s As String, codigos As String, cod As String, j As Integer
    Dim v As Variant
    For i = 1 To numGrupo
        cod = ""
        codigos = PreparaListaGrupo(lstGrupo(i - 1))
        If Len(codigos) > 0 Then
            s = s & " IVInventario.IdGrupo" & i & " in (" & codigos & ") AND "

        
            v = Split(codigos, ",")
            For j = 0 To UBound(v)
                cod = cod & Mid(v(j), 2, Len(v(j)) - 2) & ","
            Next j
            cod = Mid$(cod, 1, Len(cod) - 1)
        End If
        If Len(cod) > 0 Then
            If Me.tag = "BufferGYP" Then
                gobjMain.EmpresaActual.GNOpcion.AsignarValor "Pedidos_BufferGYPConsTipoTrans_ItemG_" & i, cod
            ElseIf Me.tag = "IVPareto" Then
                gobjMain.EmpresaActual.GNOpcion.AsignarValor "IVPareto_ItemG_" & i, cod
            
            Else
                gobjMain.EmpresaActual.GNOpcion.AsignarValor "Pedidos_BufferConsTipoTrans_ItemG_" & i, cod
            End If
        Else
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "Pedidos_BufferConsTipoTrans_ItemG_" & i, ""
        End If
        
        'End If
    Next i
    If Len(s) > 0 Then
        If Mid(s, Len(s) - 3, 4) = "AND " Then s = Mid(s, 1, Len(s) - 4)
        ArmarSqlItem = "AND (" & s & ")"
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
        RecuperaGrupoSelec key & i, lstGrupo(i - 1), trans
    Next i
End Sub


Public Sub RecuperaGrupoSelec(ByVal key As String, lst As ListBox, s As String)
Dim Vector As Variant
Dim i As Integer, j As Integer, Selec As Integer
    'Recupera selecciondados  del registro de windows
    ''''s = mobjReporte.RecuperarConfigBusqueda(Me.tag, Key)
'        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Pedidos_BufferConsTipoTrans_ItemG_" & i)) > 0 Then
'            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Pedidos_BufferConsTipoTrans_Item")
'        End If
    
    
    If s <> "_VACIO_" Then
        Vector = Split(s, ",")
         Selec = UBound(Vector, 1)
         For i = 0 To Selec
            For j = 0 To lst.ListCount - 1
                If Vector(i) = CogeSoloCodigo(lst.List(j)) Then
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
                s = s & Left(cod, lst.ItemData(i)) & ","    '      Mid(cod, 1, Len(cod) - (Len(CogeSoloCodigo(cod)))) & ","
            End If
        Next i
        If Len(s) > 0 Then
            'quita la ultima coma
            If Mid(s, Len(s), 1) = "," Then s = Mid(s, 1, Len(s) - 1)  '& vbCrLf
'            s = lbl1(X - 1).Caption & ": " & s
        End If
        Set lst = Nothing
        ArmarEtiqueta = ArmarEtiqueta & s & IIf(Len(s) > 0, "; ", "")
        s = ""
    Next X
'    ArmarEtiqueta = "Items de " & vbCrLf & ArmarEtiqueta
End Function

Public Function InicioOculto(ByVal tag As String, ByVal KeyItm As String) As String
    'Inicializa Variables
    Dim i As Integer, s As String
    For i = 1 To numGrupo
        RecuperaGrupoSelec key & i, lstGrupo(i - 1), s
    Next i

    'RecuperaDatos
    With gobjMain.objCondicion
        .Bienes = ArmarSqlItem                           '"Ha aceptado, pero debemos armar cadena"
'        .SQLItem = ArmarSqlItem
        InicioOculto = ArmarEtiqueta
        If Len(.CodItem1) > 0 Or Len(.CodItem2) > 0 Then
            InicioOculto = InicioOculto & "e ítems desde : " & .CodItem1 & " hasta " & .CodItem2
        End If
        'SaveSetting APPNAME, App.Title, KeyItm, InicioOculto
        'SaveSetting APPNAME, App.Title, "Etiq_VxItem2_Itm", InicioOculto
'        mobjReporte.GrabarConfigBusqueda Me.tag, Key, InicioOculto
    End With
    Unload Me
End Function


Private Function ArmarSqlItemP() As String
    Dim i As Long, s As String, codigos As String
    For i = 1 To numGrupo
        codigos = PreparaListaGrupo(lstGrupo(i - 1))
        If Len(codigos) > 0 Then
            s = s & " vwConsVxParetos.IVIdGrupo" & i & " in (" & codigos & ") AND "
        End If
    Next i
    If Len(s) > 0 Then
        If Mid(s, Len(s) - 3, 4) = "AND " Then s = Mid(s, 1, Len(s) - 4)
        ArmarSqlItemP = "AND (" & s & ")"
    End If
End Function



Private Sub CargaListaGruposNew()
    Dim i As Long
    Dim sql  As String, rs As Recordset, cond As String
    For i = 1 To numGrupo
        sql = "SELECT CodGrupo" & i & " as Codgrupo, Descripcion,  IDGrupo" & i & " as IDgrupo FROM IVGrupo" & i & " ORDER BY Descripcion " ' & i
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
        Do While Not rs.EOF
            'lstGrupo(i - 1).AddItem rs!CodGrupo & " " & Left(rs!Descripcion, 20) & " " & " [" & rs!idgrupo & "]"
            lstGrupo(i - 1).AddItem Left(rs!Descripcion, 20) & " " & rs!CodGrupo & " " & " [" & rs!idgrupo & "]"
            lstGrupo(i - 1).ItemData(lstGrupo(i - 1).NewIndex) = Len(rs!CodGrupo)
            rs.MoveNext
        Loop
        Set rs = Nothing
    Next i
End Sub

