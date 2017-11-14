VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmResumir 
   Caption         =   "Resumir Kardex Inventario"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   6795
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   852
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   6795
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4728
      Width           =   6792
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   4116
         TabIndex        =   8
         Top             =   36
         Width           =   1212
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Proceder"
         Height          =   372
         Left            =   2772
         TabIndex        =   7
         Top             =   36
         Width           =   1212
      End
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   540
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Frame fraCodTrans 
      Caption         =   "Cod.&Trans"
      Height          =   1572
      Left            =   2244
      TabIndex        =   3
      Top             =   96
      Width           =   2772
      Begin VB.ListBox lstTrans 
         Columns         =   3
         Height          =   852
         IntegralHeight  =   0   'False
         Left            =   240
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   240
         Width           =   2412
      End
      Begin VB.CommandButton cmdTransLimpiar 
         Caption         =   "Limp."
         Height          =   330
         Left            =   1800
         TabIndex        =   4
         Top             =   1116
         Width           =   732
      End
   End
   Begin VB.Frame fraFecha 
      Caption         =   "&Fecha (desde - hasta)"
      Height          =   1572
      Left            =   192
      TabIndex        =   0
      Top             =   96
      Width           =   2052
      Begin MSComCtl2.DTPicker dtpFecha1 
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1692
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   106692609
         CurrentDate     =   36348
      End
      Begin MSComCtl2.DTPicker dtpFecha2 
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1692
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   106692609
         CurrentDate     =   36348
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   1932
      Left            =   144
      TabIndex        =   10
      Top             =   2136
      Width           =   6372
      _cx             =   11239
      _cy             =   3408
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
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   100
      ColWidthMax     =   4000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmResumir.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   2
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
      Editable        =   0
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
      AllowUserFreezing=   3
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "frmResumir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Const MSG_ERR = "Error al Resumir"
Private mProcesando As Boolean
Private mCancelado As Boolean

Private mColItems As Collection



Public Sub InicioResumenIvK()
    Dim i As Integer
    On Error GoTo ErrTrap
    
    Me.Show
    Me.ZOrder
    dtpFecha1.value = DateAdd("M", -1, Date)
    dtpFecha2.value = Date
    CargaTrans
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

'*** MAKOTO 31/ago/00 Modificado
Private Sub CargaTrans()
    Dim i As Long, v As Variant
    
    'Carga la lista de transacción
'    fcbTrans.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "IV")

    lstTrans.Clear
    v = gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "IV")
    For i = LBound(v, 2) To UBound(v, 2)
        lstTrans.AddItem v(0, i)        '& " " & v(1, i)
    Next i
End Sub



Private Sub cmdAceptar_Click()
    ResumirIVK
End Sub

'Private Function VerificaIngreso() As String
'    Dim i As Long, cod As String, gnt As GNTrans
'    Dim s As String
'
'    For i = 0 To lstTrans.ListCount - 1
'        'Si está seleccionado
'        If lstTrans.Selected(i) Then
'            'Recupera el objeto GNTrans
'            cod = lstTrans.List(i)
'            Set gnt = gobjMain.EmpresaActual.RecuperaGNTrans(cod)
'            'Si la transaccion es de ingreso, devuelve el codigo
'            If gnt.IVTipoTrans = "I" Then s = s & cod & ", "
'        End If
'    Next i
'    Set gnt = Nothing
'    If Len(s) > 2 Then s = Left$(s, Len(s) - 2)     'Quita la ultima ", "
'    VerificaIngreso = s
'End Function



Private Function PreparaCodTrans() As String
    Dim i As Long, s As String
    
    With lstTrans
        'Si está seleccionado solo una
        If lstTrans.SelCount = 1 Then
            For i = 0 To .ListCount - 1
                If .Selected(i) Then
                    s = "'" & .List(i) & "'"    ' /* Agregado Olvier viernes 13 2001 */
                    Exit For
                End If
            Next i
        'Si está TODO o NINGUNO, no hay condición
        ElseIf (.SelCount < .ListCount) And (.SelCount > 0) Then
            For i = 0 To .ListCount - 1
                If .Selected(i) Then
                    s = s & "'" & .List(i) & "', "
                End If
            Next i
            If Len(s) > 0 Then s = Left$(s, Len(s) - 2)    'Quita la ultima ", "
        End If
    End With
    PreparaCodTrans = s
End Function


Private Sub ResumirIVK()
    Dim rs As Recordset, sql As String, Condicion As String, Tablas As String, NumReg As Long, TransID As Long
    Dim rs2 As Recordset, sql2 As String, fila As Integer
    Dim v As Variant, i As Integer, dia As Date, Registros As Long, RegistrosResumidos As Long
    On Error GoTo ErrTrap


    If lstTrans.SelCount = 0 Then
        MsgBox "Seleccione una transacción, por favor.", vbInformation
        Exit Sub
    End If
    cmdAceptar.Enabled = False
    
    grd.Clear
    grd.Rows = 1
    grd.FormatString = ">Tipo|<Transaccion|<Fecha|>Total Registros|>Resumiendo a|<Estado"
    grd.ColWidth(1) = 1000
    grd.ColWidth(2) = 1000
    grd.ColWidth(3) = 1200
    grd.ColWidth(4) = 1200
    grd.ColWidth(5) = 1500
    
    With gobjMain.objCondicion
        .fecha1 = dtpFecha1.value
        .fecha2 = dtpFecha2.value
        .CodTrans = PreparaCodTrans
        v = Split(.CodTrans, ",")
    End With
    mProcesando = True
For i = LBound(v) To UBound(v)
  dia = gobjMain.objCondicion.fecha1
  Do While dia <= gobjMain.objCondicion.fecha2
  '************ saco el TRANSID menor para todo el grupo segun la condicion de busqueda
  sql = "Select  Min(TransId) as TransId, count(TransID) as Registros " & _
       "FROM GNComprobante "
  Condicion = " WHERE gncomprobante.estado <> 3 AND codTrans = " & v(i) & " AND fechaTrans = '" & dia & "'"
  sql = sql & Condicion
  Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
  If Not IsNull(rs!TransID) Then   ' si tiene datos para resumir
    TransID = rs!TransID
    
  '*************   Resumiendo IVKARDEX *****************
    sql = "Select Min(ivk.ID) as ID, ivk.IdInventario, ivk.Idbodega," & _
                "Sum(ivk.Cantidad) as cantidad," & _
                "Sum(ivk.CostoTotal) as costototal," & _
                "Sum(ivk.CostoRealTotal)as costorealtotal," & _
                "Sum(ivk.PrecioTotal) as PrecioTotal," & _
                "Sum(ivk.PrecioRealTotal) as PrecioRealTotal," & _
                "Max(ivk.Descuento) as Descuento," & _
                "Max(ivk.IVA) As IVA"
    Tablas = " FROM GNComprobante inner join Ivkardex ivk ON GnComprobante.TransID = ivk.TransID"
    Condicion = " WHERE gncomprobante.estado <> 3 AND codTrans = " & v(i) & " AND fechaTrans = '" & dia & "'"
    sql = sql & Tablas & Condicion & " Group by ivk.IdInventario,Ivk.Idbodega"
    
    ' ******* solo para sacar cuantos registros ba ha resumir
    Set rs = gobjMain.EmpresaActual.OpenRecordset("Select Count(IvK.TransID) as TotalRegistros" & Tablas & Condicion)
    Registros = rs!TotalRegistros
    grd.AddItem "K" & vbTab & v(i) & vbTab & dia & vbTab & Registros & vbTab & "0" & vbTab & "Procesando ..."
    fila = grd.Rows - 1
    RegistrosResumidos = 0
    ''           **********************************
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
    End If
    Do Until rs.EOF
        DoEvents
        RegistrosResumidos = RegistrosResumidos + 1
        grd.TextMatrix(fila, 4) = RegistrosResumidos
        If mCancelado Then
            GoTo salida
        End If
        sql2 = "Select * from ResumenIVKardex WHERE ID = " & rs!id
        Set rs2 = gobjMain.EmpresaActual.OpenRecordset(sql2)
        If rs2.EOF Then
            'nuevo
            sql2 = "Insert into ResumenIVKardex (ID,TransID,IdInventario,IdBodega,Cantidad,CostoTotal,CostoRealTotal,PrecioTotal,PrecioRealTotal,Descuento,IVA) " & _
                   " Values ( " & _
                   rs!id & "," & _
                   TransID & "," & _
                   rs!idinventario & "," & _
                   rs!IdBodega & "," & _
                   rs!cantidad & "," & _
                   rs!CostoTotal & "," & _
                   rs!CostoRealTotal & "," & _
                   rs!PrecioTotal & "," & _
                   rs!PrecioRealTotal & "," & _
                   rs!Descuento & "," & _
                   rs!IVA & ")"
            gobjMain.EmpresaActual.EjecutarSQL sql2, NumReg
        Else
            'editar
            sql2 = "UPDATE ResumenIvKardex SET " & _
                        "TransID = " & TransID & "," & _
                        "IdInventario = " & rs!idinventario & "," & _
                        "IdBodega=" & rs!IdBodega & "," & _
                        "Cantidad=" & rs!cantidad & "," & _
                        "CostoTotal=" & rs!CostoTotal & "," & _
                        "CostoRealTotal=" & rs!CostoRealTotal & "," & _
                        "PrecioTotal=" & rs!PrecioTotal & "," & _
                        "PrecioRealTotal=" & rs!PrecioRealTotal & "," & _
                        "Descuento=" & rs!Descuento & "," & _
                        "IVA=" & rs!IVA & _
                    " WHERE ID = " & rs!id
            gobjMain.EmpresaActual.EjecutarSQL sql2, NumReg
        End If
        rs.MoveNext
    Loop
        
    grd.TextMatrix(fila, 5) = "OK"
    
    
    ' ****** resumiendo IVkarderRecargo
    
    sql = "Select " & _
          "Min(IvkR.ID) as ID," & _
          "IvkR.IdRecargo," & _
          "MAX(ivkR.Porcentaje) as Porcentaje," & _
          "SUM(ivkR.Valor) as Valor," & _
          "MAX(Orden) As Orden"
    Tablas = " FROM GNComprobante inner join IvkardexRecargo ivkR ON GnComprobante.TransID = ivkR.TransID "
    Condicion = " WHERE gncomprobante.estado <> 3 AND codTrans = " & v(i) & " AND fechaTrans = '" & dia & "'"
    sql = sql & Tablas & Condicion & " Group by CodTrans, ivkr.IdRecargo Order by  IDRecargo"
    
    

    ' ******* solo para sacar cuantos registros ba ha resumir
    Set rs = gobjMain.EmpresaActual.OpenRecordset("Select Count(IvkR.ID)     as TotalRegistros" & Tablas & Condicion)
    Registros = rs!TotalRegistros
    grd.AddItem "R" & vbTab & v(i) & vbTab & dia & vbTab & Registros & vbTab & "0" & vbTab & "Procesando ..."
    fila = grd.Rows - 1
    RegistrosResumidos = 0
    ''           **********************************
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveFirst
    End If
    Do Until rs.EOF
        RegistrosResumidos = RegistrosResumidos + 1
        grd.TextMatrix(fila, 4) = RegistrosResumidos
        DoEvents
        If mCancelado Then
            GoTo salida
        End If
        sql2 = "Select * from ResumenIVKardexRecargo WHERE ID = " & rs!id
        Set rs2 = gobjMain.EmpresaActual.OpenRecordset(sql2)
        If rs2.EOF Then
            'nuevo
            sql2 = "Insert into ResumenIVKardexRecargo (ID,TransID,IdRecargo,Porcentaje,Valor,Orden) " & _
                   " Values ( " & _
                   rs!id & "," & _
                   TransID & "," & _
                   rs!IdRecargo & "," & _
                   rs!porcentaje & "," & _
                   rs!valor & "," & _
                   rs!orden & ")"
            gobjMain.EmpresaActual.EjecutarSQL sql2, NumReg
        Else
            'editar
            sql2 = "UPDATE ResumenIvKardexRecargo SET " & _
                        "TransID = " & TransID & "," & _
                        "IdRecargo = " & rs!IdRecargo & "," & _
                        "Porcentaje=" & rs!porcentaje & "," & _
                        "Valor=" & rs!valor & "," & _
                        "Orden=" & rs!orden & _
                    " WHERE ID = " & rs!id
            gobjMain.EmpresaActual.EjecutarSQL sql2, NumReg
        End If

        rs.MoveNext
    Loop
  
  grd.TextMatrix(fila, 5) = "OK"
  End If
  dia = DateAdd("D", 1, dia)
  'MsgBox v(I) & " " & DIA
  Loop
Next i

salida:
    If mCancelado Then
        grd.TextMatrix(fila, 5) = "Cancelado .."
    End If
    mProcesando = False
    mCancelado = False
    cmdAceptar.Enabled = True
    Exit Sub
ErrTrap:
    DispErr
    mProcesando = False
    mCancelado = False
    cmdAceptar.Enabled = True
    Exit Sub
End Sub

Private Sub cmdCancelar_Click()
    If mProcesando Then
        mCancelado = True
    Else
        Unload Me
    End If
End Sub


Private Sub cmdTransLimpiar_Click()
    Dim i As Long, aux As Long
    
    aux = lstTrans.ListIndex
    For i = 0 To lstTrans.ListCount - 1
        lstTrans.Selected(i) = False
    Next i
    lstTrans.ListIndex = aux
End Sub


'Private Sub cmdTransTodo_Click()
'    Dim i As Long, aux As Long, gt As GNTrans
'    Dim cod As String
'    On Error GoTo errtrap
'    MensajeStatus "Preparando...", vbHourglass
'
'    aux = lstTrans.ListIndex
'    For i = 0 To lstTrans.ListCount - 1
'        cod = lstTrans.List(i)
'        Set gt = gobjMain.EmpresaActual.RecuperaGNTrans(cod)
'        If Not (gt Is Nothing) Then
'            'Solo marca egresos/transferencias
'            If gt.IVTipoTrans = "E" Or gt.IVTipoTrans = "T" Then
'                lstTrans.Selected(i) = True
'            End If
'        End If
'    Next i
'    lstTrans.ListIndex = aux
'    MensajeStatus
'    Exit Sub
'errtrap:
'    MensajeStatus
'    DispErr
'    Exit Sub
'End Sub


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
    If mProcesando Then
        Cancel = True
        Exit Sub
    End If
    Me.Hide         'Se pone esto para evitar el posible BUG de Windows98
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    grd.Move 0, grd.Top, Me.ScaleWidth, Me.ScaleHeight - grd.Top - pic1.Height - 80
    prg1.Width = Me.ScaleWidth - (prg1.Left * 2)
End Sub


