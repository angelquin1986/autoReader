VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "flexcombo.ocx"
Begin VB.Form frmGnSecuencia 
   Caption         =   "Cambiar Secuencial"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   8415
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picpie 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   8355
      TabIndex        =   6
      Top             =   4875
      Width           =   8415
      Begin VB.CommandButton cmdProcederSeries 
         Caption         =   "Proceder Series"
         Height          =   315
         Left            =   5310
         TabIndex        =   14
         Top             =   30
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.CommandButton cmdAsignarSerie 
         Caption         =   "Asignar Series"
         Height          =   315
         Left            =   3570
         TabIndex        =   13
         Top             =   30
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.CommandButton cmdProceder 
         Caption         =   "Proceder"
         Height          =   315
         Left            =   1800
         TabIndex        =   8
         Top             =   30
         Width           =   1605
      End
      Begin VB.CommandButton cmdAsignar 
         Caption         =   "Asignar"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   30
         Width           =   1605
      End
   End
   Begin VB.PictureBox picEnc 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   1245
      Left            =   0
      ScaleHeight     =   1245
      ScaleWidth      =   8415
      TabIndex        =   1
      Top             =   0
      Width           =   8415
      Begin VB.CommandButton cmdbuscar1 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   6630
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   6660
         TabIndex        =   3
         Top             =   150
         Width           =   1155
      End
      Begin VB.OptionButton opt 
         Caption         =   "+100 Mil"
         Height          =   555
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   570
         Value           =   -1  'True
         Width           =   945
      End
      Begin FlexComboProy.FlexCombo fcbCodTrans 
         Height          =   375
         Left            =   1110
         TabIndex        =   4
         Top             =   120
         Width           =   2115
         _ExtentX        =   3731
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
      Begin FlexComboProy.FlexCombo fcbCodTrans2 
         Height          =   375
         Left            =   4500
         TabIndex        =   9
         Top             =   150
         Width           =   2115
         _ExtentX        =   3731
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
      Begin VB.Label lbltit 
         AutoSize        =   -1  'True
         Caption         =   "Ingresar la transaccion del numero de serie"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1230
         TabIndex        =   12
         Top             =   600
         Width           =   4515
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Trans Destino"
         Height          =   195
         Left            =   3450
         TabIndex        =   10
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Trans Origen"
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   150
         Width           =   915
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grd 
      Height          =   2055
      Left            =   90
      TabIndex        =   0
      Top             =   1350
      Width           =   8415
      _cx             =   14843
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
      Cols            =   7
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
End
Attribute VB_Name = "frmGnSecuencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Indice As Integer
Dim NumMax As Long
Private Const COL_RESULT = 5

Public Sub Inicio()
    Dim i As Integer
    On Error GoTo ErrTrap
    Me.Show
    Me.ZOrder
   ' ConfigCols
    CargaTrans
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub CargaTrans()
Dim v As Variant
v = gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "IV")
fcbCodTrans.SetData v
fcbCodTrans2.SetData fcbCodTrans.GetData
End Sub

Private Sub cmdAsignar_Click()
Dim sql As String
Dim rs As Recordset
Dim i As Long
If Len(fcbCodTrans2.KeyText) = 0 Then
    MsgBox "Debe seleccionar transacccion a donde se va ha pasar los datos ... "
        fcbCodTrans2.SetFocus
    Exit Sub
End If
'sql = "Select numtranssiguiente from gnTrans where codtrans = '" & fcbCodTrans2.KeyText & "'"
sql = "select top 1 numtrans from gncomprobante where codtrans ='" & fcbCodTrans2.KeyText & "'"
sql = sql & " order by numtrans desc"
Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
If rs.RecordCount > 0 Then
    
    If IsNull(rs!numtrans) Then
        NumMax = 0
    Else
        NumMax = rs!numtrans
    End If
Else
    NumMax = 0
End If

    For i = 1 To grd.Rows - 1
        grd.TextMatrix(i, 4) = grd.ValueMatrix(i, grd.ColIndex("NumTrans")) + NumMax
    Next
End Sub

Private Sub cmdAsignarSerie_Click()
Dim i As Long
Dim rs As Recordset
Dim sql As String
Dim idKardexAsignar As Long
Dim idTransFuente As Long
For i = 1 To grd.Rows - 1
    If Not grd.IsSubtotal(i) Then
        idTransFuente = grd.ValueMatrix(i, 4)
        If idTransFuente <> grd.ValueMatrix(i - 1, 4) Then
            sql = "select g.transid,g.codtrans,g.numtrans,ivk.id,iv.codinventario,ivk.cantidad,ivk.orden"
            sql = sql & " from ivkardex ivk "
            sql = sql & " Inner Join gncomprobante g on g.transid = ivk.transid"
            sql = sql & " Inner join IvInventario iv on iv.idinventario = ivk.idinventario"
            sql = sql & " Where g.Transid =" & idTransFuente
            sql = sql & " order by iv.codinventario, ivk.orden"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If Not rs.EOF Then
            Do While Not rs.EOF
                AsignarIdIvKardex rs!TransID, rs!CodInventario, rs!id, i, rs!Orden
                rs.MoveNext
            Loop
            End If
        Else
'            grd.TextMatrix(i, 7) = idKardexAsignar
        End If
    End If
Next
End Sub
Private Function AsignarIdIvKardex(ByVal TransID As Long, ByVal CodInventario As String, ByVal idIvKardex As Long, ByVal fila As Long, ByVal Orden As Long)
Dim i As Long
For i = fila To grd.Rows - 1
    If Not grd.IsSubtotal(i) Then
        If grd.TextMatrix(i, 5) = CodInventario Then
            grd.TextMatrix(i, 7) = idIvKardex
            grd.TextMatrix(i, 8) = TransID
            grd.TextMatrix(i, 9) = Orden
        End If
    End If
Next
End Function
Private Sub cmdBuscar_Click()
Dim sql As String
Dim rs As Recordset

If Len(fcbCodTrans.KeyText) = 0 Then
    MsgBox "Debe escoger una transaccion"
    fcbCodTrans.SetFocus
    Exit Sub
End If

sql = "Select transid,codtrans,numtrans,'' as NuevoNum,'' as result from gncomprobante where codtrans = '" & fcbCodTrans.KeyText & "'"
Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
Set grd.DataSource = rs
ConfigCols
End Sub

Private Sub cmdBuscar1_Click() 'para series yolita
Dim sql As String
Dim rs As Recordset

If Len(fcbCodTrans.KeyText) = 0 Then
    MsgBox "Debe escoger una transaccion"
    fcbCodTrans.SetFocus
    Exit Sub
End If

sql = "select g.codtrans,g.numtrans,ivk.id,g.idtransfuente,iv.codinventario,ivk.idserie,ivk.idivkardex "
sql = sql & "from ivkardexserie ivk Inner Join gncomprobante g on g.transid = ivk.transid "
sql = sql & "Inner Join ivSerie ivs Inner Join ivinventario iv on iv.idinventario = ivs.idinventario "
sql = sql & "on ivs.idserie = ivk.idserie "
sql = sql & " Where g.codtrans = '" & fcbCodTrans.KeyText & "'"
sql = sql & " And g.estado <> 3"
sql = sql & " Order by g.codtrans,g.numtrans,iv.codinventario"
Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
Set grd.DataSource = rs
ConfigColsSerie

End Sub

Private Sub cmdProceder_Click()
Dim i As Long
Dim sql As String, x As Single

If Len(fcbCodTrans2.KeyText) = 0 Then
    MsgBox "Debe seleccionar transacccion a donde se va ha pasar los datos ... "
        fcbCodTrans2.SetFocus
    Exit Sub
End If

For i = 1 To grd.Rows - 1
    DoEvents
    If grd.TextMatrix(i, COL_RESULT) <> "OK." Then
        sql = "Update gncomprobante set numtrans = " & grd.ValueMatrix(i, 4)
        sql = sql & ",codtrans = '" & fcbCodTrans2.KeyText & "'"
        sql = sql & " Where transid = " & grd.ValueMatrix(i, 1)
        gobjMain.EmpresaActual.EjecutarSQL sql, 1
        grd.Row = i
        x = grd.CellTop                 'Para visualizar la celda actual
        grd.IsSelected(i) = True
        grd.TextMatrix(i, COL_RESULT) = "OK."
        grd.Redraw = flexRDDirect
    End If
Next
sql = "Update gnTrans set  numtransSiguiente = 1"
sql = sql & " Where codtrans  = '" & fcbCodTrans.KeyText & "'"
gobjMain.EmpresaActual.EjecutarSQL sql, 1

sql = "Update gnTrans set  numtransSiguiente = " & NumMax + grd.Rows
sql = sql & " Where codtrans  = '" & fcbCodTrans2.KeyText & "'"
gobjMain.EmpresaActual.EjecutarSQL sql, 1

End Sub

Private Sub cmdProcederSeries_Click()
Dim i As Long
Dim sql As String, x As Single


For i = 1 To grd.Rows - 1
    DoEvents
    If Not grd.IsSubtotal(i) Then
        If grd.TextMatrix(i, 9) <> "OK." Then
            sql = "Update ivkardexSerie set transid = " & grd.ValueMatrix(i, 4)
            sql = sql & ",idivkardex = " & grd.ValueMatrix(i, 7)
            sql = sql & ",ordenivkardex = " & grd.ValueMatrix(i, 9)
            sql = sql & " Where id = " & grd.ValueMatrix(i, 3)
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            grd.Row = i
            x = grd.CellTop                 'Para visualizar la celda actual
            grd.IsSelected(i) = True
            grd.TextMatrix(i, 10) = "OK."
            grd.Redraw = flexRDDirect
        End If
    End If
Next
End Sub

Private Sub Form_Resize()
picEnc.Top = 0
grd.Move 0, picEnc.Height, Me.ScaleWidth, Me.ScaleHeight - picEnc.Height - picpie.Height
End Sub
Private Sub ConfigCols()
    Dim s As String, i As Long, j As Integer
    With grd
        s = "^#|<Transid|<CodTrans|>NumTrans|>NuevoNumTrans|<Result."
        .FormatString = s
        GNPoneNumFila grd, False
        AjustarAutoSize grd, -1, -1, 4000
        AsignarTituloAColKey grd
        .ColData(.ColIndex("CodTrans")) = -1
        .ColData(.ColIndex("NumTrans")) = -1
        .ColData(.ColIndex("NuevoNumTrans")) = -1
        .ColHidden(1) = True
        grd.Editable = flexEDNone
'Color de fondo
'        If .Rows > .FixedRows Then
'            .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("Descripción")) = .BackColorFrozen
'        End If
    End With
End Sub

Private Sub ConfigColsSerie()
    Dim s As String, i As Long, j As Integer
    With grd
        s = "^#|>CodTrans|>NumTrans|<id|>idTransFuente|<CodInventario|>IdSerie|<idIvKardex|<Transid|<OrdenIvkardex|<Resultado"
        .FormatString = s
        GNPoneNumFila grd, False
        AjustarAutoSize grd, -1, -1, 4000
        AsignarTituloAColKey grd
'        .ColData(.ColIndex("CodTrans")) = -1
'        .ColData(.ColIndex("NumTrans")) = -1
'        .ColData(.ColIndex("NuevoNumTrans")) = -1
'        .ColHidden(1) = True
        grd.Editable = flexEDNone
'        grd.SubTotal flexSTClear
        grd.SubTotal flexSTNone, 2, 0, , grd.GridColor, vbBlack, , "Total", 1, True

    End With
End Sub


Private Sub opt_Click(Index As Integer)
Select Case Index
    Case 0: Indice = 0
    Case 1: Indice = 1
    Case 2: Indice = 2
    Case 3: Indice = 3
End Select
End Sub
Public Sub InicioSerie()
    Dim i As Integer
    On Error GoTo ErrTrap
    Me.Show
    Me.ZOrder
   ' ConfigCols
   opt(0).Visible = False
   cmdBuscar.Visible = False
   cmdbuscar1.Visible = True
   fcbCodTrans2.Visible = False
   Label2.Visible = False
   cmdAsignar.Visible = False
   cmdProceder.Visible = False
   cmdAsignarSerie.Visible = True
   cmdProcederSeries.Visible = True
    CargaTransSerie
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub CargaTransSerie()
Dim v As Variant
v = gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "IV")
fcbCodTrans.SetData v

End Sub

