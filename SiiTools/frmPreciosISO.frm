VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPreciosISO 
   Caption         =   "Actualización de precios Isollanta"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   8160
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicBusca 
      Align           =   1  'Align Top
      Height          =   1635
      Left            =   0
      ScaleHeight     =   1575
      ScaleWidth      =   8100
      TabIndex        =   5
      Top             =   0
      Width           =   8160
      Begin VB.Frame fraPrecio 
         Caption         =   "Acción"
         Height          =   1395
         Left            =   6240
         TabIndex        =   13
         Top             =   120
         Width           =   6000
         Begin VB.CommandButton cmdRedondear 
            Caption         =   "Re&dondear precios nuevos"
            Height          =   375
            Left            =   3420
            TabIndex        =   23
            Top             =   900
            Width           =   2412
         End
         Begin VB.CommandButton cmdCalcular 
            Caption         =   "Calcular"
            Height          =   375
            Left            =   240
            TabIndex        =   22
            Top             =   900
            Width           =   1815
         End
         Begin VB.ComboBox cboRedondeo 
            Height          =   315
            Left            =   4080
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   480
            Width           =   1812
         End
         Begin VB.OptionButton optBajar 
            Caption         =   "&Bajar"
            Height          =   240
            Left            =   240
            TabIndex        =   16
            Top             =   480
            Width           =   972
         End
         Begin VB.OptionButton optAlzar 
            Caption         =   "Al&zar"
            Height          =   240
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Value           =   -1  'True
            Width           =   972
         End
         Begin VB.TextBox txtPorcent 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   1320
            TabIndex        =   14
            Top             =   240
            Width           =   700
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   4080
            TabIndex        =   18
            Top             =   120
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            Format          =   41222145
            CurrentDate     =   40654
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "R&edondeo"
            Height          =   240
            Left            =   2520
            TabIndex        =   21
            Top             =   480
            Width           =   1140
         End
         Begin VB.Label Label1 
            Caption         =   "%"
            Height          =   252
            Left            =   2100
            TabIndex        =   20
            Top             =   360
            Width           =   372
         End
         Begin VB.Label Label5 
            Caption         =   "&Fecha Lista Precios"
            Height          =   255
            Left            =   2520
            TabIndex        =   19
            Top             =   180
            Visible         =   0   'False
            Width           =   1455
         End
      End
      Begin VB.Frame fraItem 
         Caption         =   "&Rango de items"
         Height          =   1455
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   6012
         Begin FlexComboProy.FlexCombo fcbDesde 
            Height          =   345
            Left            =   900
            TabIndex        =   7
            Top             =   300
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   609
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
         Begin FlexComboProy.FlexCombo fcbHasta 
            Height          =   345
            Left            =   900
            TabIndex        =   8
            Top             =   660
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   609
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
         Begin FlexComboProy.FlexCombo fcbBanda 
            Height          =   345
            Left            =   900
            TabIndex        =   9
            Top             =   1020
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   609
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
         Begin VB.Label Label3 
            Caption         =   "&Diseño"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   660
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "&Tamaño"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "&Banda"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1020
            Width           =   735
         End
      End
   End
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   852
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   8160
      TabIndex        =   0
      Top             =   7710
      Width           =   8160
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   372
         Left            =   1320
         TabIndex        =   3
         Top             =   0
         Width           =   1452
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   3480
         TabIndex        =   2
         Top             =   0
         Width           =   1452
      End
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   540
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   3375
      Left            =   60
      TabIndex        =   4
      Top             =   1740
      Width           =   12075
      _cx             =   21299
      _cy             =   5953
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
      Rows            =   3
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
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
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
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   12648447
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "frmPreciosISO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private numGrupo As Integer
Private IVA As Currency

Const COL_TAMANIO = 1
Const COL_TRABAJO = 2
Const COL_DISENIO = 3
Const COL_PRECIOACT = 4
Const COL_PRECIOANT = 5
Const COL_PRECIO1 = 6
Const COL_PRECIO1IVA = 7


Public Sub Inicio()
    Dim i As Integer
    On Error GoTo ErrTrap

    
    IVA = gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA
    
    Me.Show
    Me.ZOrder
    Me.Caption = "Actualización de precios"
    
    cboRedondeo.Clear
    For i = -2 To 6
        cboRedondeo.AddItem Format(10 ^ i, "#,0.00")
    Next i
    cboRedondeo.ListIndex = 0
    optPorGrupo_Click 0
    ConfigColsIVESPISO
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub


Private Sub cmdAceptar_Click()
'    If fcbDesde.Vacio Then
'        MsgBox "Seleccione el inicio de rango."
'        fcbDesde.SetFocus
'        Exit Sub
'    End If
'    If fcbHasta.Vacio Then
'        MsgBox "Seleccione el fin de rango."
'        fcbHasta.SetFocus
'        Exit Sub
'    End If
    
    
    ActualizaPrecio
End Sub

Private Function ActualizaPrecio() As Boolean
    Dim upd As String, betw As String, s As String, p As Single
    Dim v As Variant, i As Long, cod As String, j As Integer
    Dim item As IVinventario, cap As String, msg As String
    Dim sql As String, rs As Recordset
    On Error GoTo ErrTrap
    cap = Me.Caption

    s = "Está seguro que desea "
    If optAlzar.value Then
        s = s & "ALZAR"
    ElseIf optBajar.value Then
        s = s & "BAJAR"
    Else
        s = s & "CALCULAR DE COSTO"
    End If
    s = s & " los precios por " & Val(txtPorcent.Text) & " por ciento?"
    If MsgBox(s, vbYesNo + vbQuestion) <> vbYes Then Exit Function

    
    
    Screen.MousePointer = vbHourglass
    
    If grd.Rows > 0 Then
        prg1.min = 0
        prg1.max = grd.Rows
        prg1.value = 0
        For i = 1 To grd.Rows - 1
            
            sql = "UPDATE ivespprodiso SET "
            sql = sql & " precio= " & Round(grd.ValueMatrix(i, COL_PRECIO1), 2)
            sql = sql & " , porcentaje= " & grd.ValueMatrix(i, COL_PRECIOACT)
            sql = sql & " from"
            sql = sql & " ivespprodiso ives"
            sql = sql & " inner join vwIVInventarioRecuperar ivtam on ives.idtamanio = ivtam.idinventario"
            sql = sql & " inner join vwIVInventarioRecuperar ivtra on ives.idtrabajo = ivtra.idinventario"
            sql = sql & " inner join vwIVInventarioRecuperar ivdis on ives.iddisenio = ivdis.idinventario"
            sql = sql & " Where"
            sql = sql & " ivtam.codinventario='" & grd.TextMatrix(i, COL_TAMANIO) & "'"
            sql = sql & " AND ivtra.codinventario ='" & grd.TextMatrix(i, COL_TRABAJO) & "'"
            sql = sql & " AND ivdis.codinventario ='" & grd.TextMatrix(i, COL_DISENIO) & "'"
            
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            DoEvents
            prg1.value = i + 1
        Next i
    End If
    Me.Caption = cap
    Screen.MousePointer = 0
    MsgBox "La actualización de precios se ha finalizado.", vbInformation
    prg1.value = prg1.min
    ActualizaPrecio = True
    Exit Function
ErrTrap:
    Me.Caption = cap
    Screen.MousePointer = 0
    DispErr
    prg1.value = prg1.min
    Exit Function
End Function


Private Sub cmdCalcular_Click()

    Dim s As String, v As Single
    Dim i As Long, p As Single
    's = InputBox("Ingrese el valor de IVA (%)", "Asignar un valor", "15")
'    If IsNumeric(s) Then
'        v = CSng(s)
'    Else
'        MsgBox "Debe ingresar un valor numérico. (ejm. 15 para 15%)", vbInformation
'        grd.SetFocus
'        Exit Sub
'    End If
    
    
    With grd
        prg1.min = 0
        prg1.max = grd.Rows
        prg1.value = 0
    
        For i = .FixedRows To .Rows - 1
        
            If optAlzar.value Then
                p = 1 + (Val(txtPorcent.Text) / 100#)   'Alzar
            ElseIf optBajar.value Then
                p = 1 - (Val(txtPorcent.Text) / 100#)   'Bajar
            End If
        
        
            .TextMatrix(i, 6) = .ValueMatrix(i, 4) * p
            .TextMatrix(i, 7) = .ValueMatrix(i, 6) * (1 + IVA)
'            .TextMatrix(i, 8) = .ValueMatrix(i, 4)
'            .TextMatrix(i, 9) = .ValueMatrix(i, 5) * (1 + iva)
            DoEvents
            prg1.value = i + 1
            
        Next i
    End With
End Sub



Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub fcbBanda_Selected(ByVal Text As String, ByVal KeyText As String)
    CargarDatos numGrupo + 1
End Sub

Private Sub fcbDesde_Selected(ByVal Text As String, ByVal KeyText As String)
    fcbHasta.KeyText = fcbDesde.KeyText
    CargarDatos numGrupo + 1
End Sub

Private Sub fcbHasta_Selected(ByVal Text As String, ByVal KeyText As String)
    CargarDatos numGrupo + 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF9
        cmdAceptar_Click
        KeyCode = 0
    Case Else
        MoverCampo Me, KeyCode, Shift, True
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ImpideSonidoEnter Me, KeyAscii
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Me.Hide         'Se pone esto para evitar el posible BUG de Windows98
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    grd.Move 0, PicBusca.Height, Me.ScaleWidth, Me.ScaleHeight - PicBusca.Height - pic1.Height - 80
    prg1.Width = Me.ScaleWidth - (prg1.Left * 2)
End Sub

Private Sub optPorCodigo_Click()
    Dim v As Variant, i As Long
    On Error GoTo ErrTrap
    Screen.MousePointer = vbHourglass
    fcbDesde.Clear
    fcbHasta.Clear
    
    v = gobjMain.EmpresaActual.ListaIVInventarioSimple
    fcbDesde.SetData v
    fcbHasta.SetData v
    Screen.MousePointer = 0
    Exit Sub
ErrTrap:
    Screen.MousePointer = 0
    DispErr
    Exit Sub
End Sub

Private Sub optPorGrupo_Click(Index As Integer)
    Dim v As Variant, i As Long
    On Error GoTo ErrTrap
    Screen.MousePointer = vbHourglass
     numGrupo = Index
    fcbDesde.Clear
    fcbHasta.Clear
    
    Select Case Index
        Case 0
            
            v = gobjMain.EmpresaActual.ListaIVInventarioSoloIvGrupo(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("NumIvGrupoTamanio") + 1, False, False, gobjMain.EmpresaActual.GNOpcion.ObtenerValor("IdGruposTamanio"))
    End Select

    'v = gobjMain.EmpresaActual.ListaIVGrupo(Index + 1, False, False)
    v = gobjMain.EmpresaActual.ListaIVInventarioSoloIvGrupo(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("NumIvGrupoTamanio") + 1, False, False, gobjMain.EmpresaActual.GNOpcion.ObtenerValor("IdGruposTamanio"))
    fcbDesde.SetData v
    v = gobjMain.EmpresaActual.ListaIVInventarioSoloIvGrupo(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("NumIvGrupoTrabajo") + 1, False, False, gobjMain.EmpresaActual.GNOpcion.ObtenerValor("IdGruposTrabajo"))
    fcbHasta.SetData v
    v = gobjMain.EmpresaActual.ListaIVInventarioSoloIvGrupo(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("NumIvGrupoDisenio") + 1, False, False, gobjMain.EmpresaActual.GNOpcion.ObtenerValor("IdGruposDisenio"))
    fcbBanda.SetData v
    
    Screen.MousePointer = 0
    Exit Sub
ErrTrap:
    Screen.MousePointer = 0
    DispErr
    Exit Sub
End Sub

Private Sub grd_CellChanged(ByVal Row As Long, ByVal col As Long)
    Dim p As Currency
    
    If Row > 0 Then
    Select Case col
        Case 6
            grd.TextMatrix(Row, 7) = grd.ValueMatrix(Row, 6) * (1 + IVA)
        Case 7
            grd.TextMatrix(Row, 6) = grd.ValueMatrix(Row, 7) / (1 + IVA)
    End Select
    End If

End Sub

Private Sub txtPorcent_KeyPress(KeyAscii As Integer)
    'Acepta solo numericos, BackSpace y punto decimal
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
        (KeyAscii <> vbKeyBack) And _
        (KeyAscii <> Asc(".")) Then
        KeyAscii = 0
    End If
End Sub

Private Sub CargarDatos(numGrupo As Integer)
    Dim sql As String, rs As Recordset

        sql = "SELECT ivtam.codinventario,"
        sql = sql & " ivtra.codinventario,"
        sql = sql & " ivdis.codinventario,"
        sql = sql & " precio, porcentaje, 0,0 "
        sql = sql & " from"
        sql = sql & " ivespprodiso ives"
        sql = sql & " inner join vwIVInventarioRecuperar ivtam on ives.idtamanio = ivtam.idinventario"
        sql = sql & " inner join vwIVInventarioRecuperar ivtra on ives.idtrabajo = ivtra.idinventario"
        sql = sql & " inner join vwIVInventarioRecuperar ivdis on ives.iddisenio = ivdis.idinventario"
    
        If Len(fcbDesde.KeyText) > 0 Or Len(fcbHasta.KeyText) > 0 Or Len(fcbBanda.KeyText) > 0 Then

            If Len(fcbDesde.KeyText) > 0 Then
                If InStr(1, sql, "where", vbTextCompare) Then
                    sql = sql & " AND "
                Else
                    sql = sql & " Where"
                End If

                sql = sql & " ivtam.codinventario ='" & fcbDesde.KeyText & "'"
            End If
            If Len(fcbHasta.KeyText) > 0 Then
                If InStr(1, sql, "where", vbTextCompare) Then
                    sql = sql & " AND "
                Else
                    sql = sql & " Where"
                End If
                
                sql = sql & " ivtra.codinventario ='" & fcbHasta.KeyText & "'"
            End If
            If Len(fcbBanda.KeyText) > 0 Then
                If InStr(1, sql, "where", vbTextCompare) Then
                    sql = sql & " AND "
                Else
                    sql = sql & " Where"
                End If
                sql = sql & " ivdis.codinventario ='" & fcbBanda.KeyText & "'"
            End If
        End If
        sql = sql & " order by ivtam.codinventario, ivtra.codinventario, ivdis.codinventario"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    With grd
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If Not rs.EOF Then .LoadArray MiGetRows(rs)
        ConfigColsIVESPISO
        GNPoneNumFila grd, False
        AjustarAutoSize grd, -1, -1, 4000
        
        .Redraw = flexRDBuffered
        .SetFocus
        
        
    End With

End Sub

Private Sub ConfigColsIVESPISO()
    With grd
        .FormatString = "^#|<Medidas|<Trabajo|<Diseño|>Precio Actual|>Precio Anterior|>Nuevo Precio|>Nuevo Precio + IVA"
        .ColWidth(0) = 600      '#
        .ColWidth(1) = 2000
        .ColWidth(2) = 2000
        .ColWidth(3) = 2000
        .ColWidth(4) = 1500
        .ColWidth(5) = 1500
        .ColWidth(6) = 1500
        .ColWidth(7) = 1500
        .Refresh
        .Redraw = flexRDNone
    
    'Tipo de datos
        .ColDataType(1) = flexDTString
        .ColDataType(2) = flexDTString
        .ColDataType(3) = flexDTString
        .ColDataType(4) = flexDTCurrency
        .ColDataType(5) = flexDTCurrency
        .ColDataType(6) = flexDTCurrency
        .ColDataType(7) = flexDTCurrency
        
        .ColFormat(4) = gobjMain.EmpresaActual.GNOpcion.FormatoCantidad
        .ColFormat(5) = gobjMain.EmpresaActual.GNOpcion.FormatoCantidad
        .ColFormat(6) = gobjMain.EmpresaActual.GNOpcion.FormatoCantidad
        .ColFormat(7) = gobjMain.EmpresaActual.GNOpcion.FormatoCantidad
        
        .ColData(1) = -1
        .ColData(2) = -1
        .ColData(3) = -1
        .ColData(4) = -1
        .ColData(5) = -1
        

        .Refresh
        .Redraw = flexRDNone
        
            If .Rows > .FixedRows Then
                .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, 5) = .BackColorFrozen
            End If
        
    End With
End Sub

Private Sub cmdRedondear_Click()
    Dim p As Currency, i As Long, d As Integer
    Dim col_antes As Long, row_antes As Long
    On Error GoTo ErrTrap
    
    If cboRedondeo.ListIndex >= 0 Then
        d = cboRedondeo.ListIndex - 2
    Else
        MsgBox "Seleccione límite de redondeo.", vbInformation
        cboRedondeo.SetFocus
        Exit Sub
    End If
    
    With grd
        .Redraw = flexRDNone
        
        'Guarda la posición del cursor para reubicar después
        col_antes = .col
        row_antes = .Row
        
        For i = .FixedRows To .Rows - 1
            .Row = i
            
            p = .ValueMatrix(i, COL_PRECIO1)
            .TextMatrix(i, COL_PRECIO1) = Redondear(p, d)
            
            p = .ValueMatrix(i, COL_PRECIO1 + 1)
            .TextMatrix(i, COL_PRECIO1 + 1) = Redondear(p, d)

            
            .col = COL_PRECIO1
            
            
            
'            grd_AfterEdit .Row, .Col         'Actualiza la utilidad
            

        Next i
        
        'Reubica el cursor a la posicion anterior
        .col = col_antes
        .Row = row_antes
        col_antes = .CellTop        'Para mover en la vista la celda actual
        
        .Redraw = flexRDBuffered
        .SetFocus
    End With
    Exit Sub
ErrTrap:
    DispErr
    grd.SetFocus
    Exit Sub
End Sub


