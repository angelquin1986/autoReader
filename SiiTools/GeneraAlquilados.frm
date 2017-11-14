VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGeneraAlquilados 
   Caption         =   "Generar Alquilados"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5325
   ScaleWidth      =   6810
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   2280
      ScaleHeight     =   1155
      ScaleWidth      =   11115
      TabIndex        =   23
      Top             =   0
      Width           =   11175
      Begin FlexComboProy.FlexCombo fcbTrans 
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
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
         AutoSize        =   -1  'True
         Caption         =   "Seleccione transaccion"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   1665
      End
   End
   Begin VB.Frame fraCodTrans 
      Caption         =   "Cod.&Trans"
      Height          =   1215
      Left            =   2400
      TabIndex        =   15
      Top             =   0
      Width           =   8295
      Begin VB.ListBox lstTrans 
         Columns         =   6
         Height          =   495
         IntegralHeight  =   0   'False
         Left            =   240
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   22
         Top             =   240
         Width           =   7935
      End
      Begin VB.CommandButton cmdTransLimpiar 
         Caption         =   "Limp."
         Height          =   330
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   732
      End
      Begin VB.CheckBox chkEstado 
         Caption         =   "&No aprobados"
         Height          =   252
         Index           =   0
         Left            =   960
         TabIndex        =   20
         Top             =   720
         Width           =   1335
      End
      Begin VB.CheckBox chkEstado 
         Caption         =   "&Aprobados"
         Height          =   252
         Index           =   1
         Left            =   2400
         TabIndex        =   19
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkEstado 
         Caption         =   "&Despachados"
         Height          =   252
         Index           =   2
         Left            =   3600
         TabIndex        =   18
         Top             =   720
         Width           =   1452
      End
      Begin VB.CheckBox chkEstado 
         Caption         =   "S&emi Despachados"
         Height          =   252
         Index           =   4
         Left            =   5040
         TabIndex        =   17
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox chkEstado 
         Caption         =   "A&nulados"
         Height          =   252
         Index           =   3
         Left            =   6960
         TabIndex        =   16
         Top             =   720
         Width           =   1092
      End
   End
   Begin VB.CommandButton cmdBuscaColas 
      Caption         =   "&Buscar"
      Height          =   372
      Left            =   1920
      TabIndex        =   12
      Top             =   1320
      Width           =   1452
   End
   Begin VB.Frame fraFecha 
      Caption         =   "&Fecha (desde - hasta)"
      Height          =   1095
      Left            =   168
      TabIndex        =   0
      Top             =   120
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
         Format          =   88801281
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
         Format          =   88801281
         CurrentDate     =   36348
      End
   End
   Begin VB.Frame fraNumTrans 
      Caption         =   "# T&rans. (desde - hasta)"
      Height          =   1215
      Left            =   11040
      TabIndex        =   3
      Top             =   0
      Width           =   2052
      Begin VB.TextBox txtNumTrans1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   360
         TabIndex        =   4
         Top             =   280
         Width           =   1212
      End
      Begin VB.TextBox txtNumTrans2 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   360
         TabIndex        =   5
         Top             =   640
         Width           =   1212
      End
   End
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   852
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   6810
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4470
      Width           =   6810
      Begin VB.CommandButton cmdverificarP 
         Caption         =   "Verificar Form."
         Height          =   372
         Left            =   4080
         TabIndex        =   27
         ToolTipText     =   "Verifica y arregla segun formula"
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdGrabarColas 
         Caption         =   "Grabar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5640
         TabIndex        =   13
         Top             =   0
         Width           =   1365
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Proceder"
         Enabled         =   0   'False
         Height          =   372
         Left            =   2040
         TabIndex        =   11
         Top             =   0
         Width           =   1212
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   7800
         TabIndex        =   9
         Top             =   0
         Width           =   1212
      End
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   540
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   1935
      Left            =   4680
      TabIndex        =   7
      Top             =   6240
      Width           =   7695
      _cx             =   13573
      _cy             =   3413
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   100
      ColWidthMax     =   4000
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   372
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   1452
   End
   Begin VSFlex7LCtl.VSFlexGrid grdOP 
      Height          =   5895
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   4575
      _cx             =   8070
      _cy             =   10398
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   100
      ColWidthMax     =   4000
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin VSFlex7LCtl.VSFlexGrid grdFormula 
      Height          =   1935
      Left            =   4680
      TabIndex        =   24
      Top             =   1800
      Width           =   7695
      _cx             =   13573
      _cy             =   3413
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
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   100
      ColWidthMax     =   4000
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      Caption         =   "Trans"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4680
      TabIndex        =   28
      Top             =   1440
      Width           =   690
   End
End
Attribute VB_Name = "frmGeneraAlquilados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'AUC 15/11/07 Creado para historial de cliente
Option Explicit

'Constantes para las columnas
Private Const COL_NUMFILA = 0
Private Const COL_TID = 1
Private Const COL_FECHA = 2
Private Const COL_CODASIENTO = 3
Private Const COL_CODTRANS = 4
Private Const COL_NUMTRANS = 5
Private Const COL_NUMDOCREF = 6
Private Const COL_NOMBRE = 7
Private Const COL_DESC = 8
Private Const COL_CENTROCOSTO = 9
Private Const COL_ESTADO = 10
Private Const COL_RESULTADO = 11
Private Const MSG_EX = "Ya Existe."
Private mProcesando As Boolean
Private mCancelado As Boolean
Private Const COL_COLASMOD = 22
Private Const COL_COLASRES = 23
Private gc As GNComprobante
Private mColItems As Collection
Private mobjivkProceso As IVKProceso
Public Sub Inicio()
    Dim i As Integer
    On Error GoTo errtrap
    Me.Show
    Me.ZOrder
    cmdBuscar.Visible = True
    cmdBuscaColas.Visible = False
    cmdAceptar.Visible = True
    cmdGrabarColas.Visible = False
    dtpFecha1.value = gobjMain.EmpresaActual.GNOpcion.FechaInicio
    dtpFecha2.value = Date
    CargaTrans
    Exit Sub
errtrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub CargaTrans()
    Dim i As Long, v As Variant
    Dim s As String
    lstTrans.Clear
    v = gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "IV")
    For i = LBound(v, 2) To UBound(v, 2)
        lstTrans.AddItem v(0, i)        '& " " & v(1, i)
    Next i
   
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaRecosteo")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaRecosteo")
            RecuperaTrans "KeyT", lstTrans, s
        End If
    
End Sub

Private Function VerificaIngreso() As String
    Dim i As Long, cod As String, gnt As GNTrans
    Dim s As String
    
    For i = 0 To lstTrans.ListCount - 1
        'Si está seleccionado
        If lstTrans.Selected(i) Then
            'Recupera el objeto GNTrans
            cod = lstTrans.List(i)
            Set gnt = gobjMain.EmpresaActual.RecuperaGNTrans(cod)
            'Si la transaccion es de ingreso, devuelve el codigo
            If gnt.IVTipoTrans = "I" Then s = s & cod & ", "
        End If
    Next i
    Set gnt = Nothing
    If Len(s) > 2 Then s = Left$(s, Len(s) - 2)     'Quita la ultima ", "
    VerificaIngreso = s
End Function

Private Function GeneraHistorial(bandVerificar As Boolean, BandTodo As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, X As Single, xi As Long
    Dim cambiado As Boolean, rs As Recordset
    Dim obj
    Dim sql As String, t As Currency, tdesc As Currency
    Dim gn As GNComprobante, Orden As Long
    Dim trans As String
    
    On Error GoTo errtrap
    mCancelado = False
    mProcesando = True
    frmMain.mnuFile.Enabled = False
    cmdAceptar.Enabled = False
    cmdBuscar.Enabled = False
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    
    For i = grd.FixedRows To grd.Rows - 1
        DoEvents
        If mCancelado Then
            MsgBox "El proceso fue cancelado.", vbInformation
            Exit For
        End If
        
        prg1.value = i
        grd.Row = i
        X = grd.CellTop                 'Para visualizar la celda actual
        tid = grd.ValueMatrix(i, COL_TID)
        trans = Trim(grd.TextMatrix(i, COL_CODTRANS)) & " " & grd.ValueMatrix(i, COL_NUMTRANS)
        
        'verifica si ya existe la transaccion
        If grd.TextMatrix(i, COL_RESULTADO) <> "OK." Then
             
                    Set gn = gobjMain.EmpresaActual.RecuperaGNComprobante(tid, grd.TextMatrix(i, COL_CODTRANS), grd.TextMatrix(i, COL_NUMTRANS))
                    If Not gobjMain.EmpresaActual.VerificaAlquilado(gn.TransID, gn.GNTrans.IdBodegaDesPre) Then     'AUC deberia buscar por codtrans y numtrans
                    'If obj Is Nothing Then
                        gn.RecuperaDetalleTodo
                        Orden = gn.CountIVKardex + 1
                        For xi = 1 To gn.CountIVKardex
                        
                            sql = "insert into IVKARDEX (TransID,IdInventario,IdBodega,Cantidad,CostoTotal,CostoRealTotal,PrecioTotal," & _
                                "PrecioRealTotal,Descuento,IVA,Orden,Nota,NumeroPrecio,ValorRecargoItem,TiempoEntrega,bandImprimir," & _
                                "idPadre,idPadreSub,FechaDevol,NumDias,fechaLleva)"
                            sql = sql & "VALUES(" & gn.TransID & "," & gn.IVKardex(xi).idinventario & "," & gn.GNTrans.IdBodegaDesPre & "," & Abs(gn.IVKardex(xi).cantidad) & _
                                "," & Abs(gn.IVKardex(xi).CostoTotal) & "," & Abs(gn.IVKardex(xi).CostoRealTotal) & "," & Abs(gn.IVKardex(xi).PrecioTotal) & _
                                "," & Abs(gn.IVKardex(xi).PrecioRealTotal) & "," & Abs(gn.IVKardex(xi).Descuento) & "," & Abs(gn.IVKardex(xi).IVA) & "," & Orden & ","
                            sql = sql & "'" & gn.IVKardex(xi).Nota & "'," & Abs(gn.IVKardex(xi).NumeroPrecio) & "," & Abs(gn.IVKardex(xi).ValorRecargoItem) & ",'" & gn.IVKardex(xi).TiempoEntrega & "',"
                            sql = sql & 0 & "," & gn.IVKardex(xi).IdPadre & "," & gn.IVKardex(xi).idpadresub & ",'" & gn.IVKardex(xi).FechaDevol & "'," & gn.IVKardex(xi).NumDias & ",'" & gn.IVKardex(xi).FechaLleva & "')"
                            Orden = Orden + 1
                            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                        Next xi
                        grd.TextMatrix(i, COL_RESULTADO) = "OK."
                    Else
                        grd.TextMatrix(i, COL_RESULTADO) = "SI EXISTE"
                End If
        End If
        Set obj = Nothing
        Set rs = Nothing
        Set gn = Nothing
siguiente:
    Next i
    Screen.MousePointer = 0
    GoTo salida
    If mCancelado Then Exit Function
errtrap:
    s = Err.Description & ":   '" & grd.TextMatrix(i, COL_CODTRANS) & " " & grd.TextMatrix(i, COL_NUMTRANS) & "' , " & grd.TextMatrix(i, COL_DESC)
    grd.TextMatrix(i, COL_RESULTADO) = "Error " & s
    If MsgBox(s & vbCr & vbCr & _
                "Desea continuar con el siguiente registro?", _
                vbQuestion + vbYesNo) = vbYes Then
        Resume siguiente
    Else
        mCancelado = True
    End If
    GoTo salida
salida:
    Screen.MousePointer = 0
    GeneraHistorial = True
    frmMain.mnuFile.Enabled = True
    cmdAceptar.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    Exit Function
End Function

Private Sub cmdBuscaColas_Click()
 Dim v As Variant, obj As Object, s As String
 Dim rs As Recordset
    On Error GoTo errtrap
    MensajeStatus "Preparando...", vbHourglass
'    If lstTrans.SelCount = 0 Then
'        MsgBox "Seleccione una transacción, por favor.", vbInformation
'        Exit Sub
'    End If
   grdOP.Rows = 1
   grdFormula.Rows = 1
   grd.Rows = 1
    
    With gobjMain.objCondicion
        .fecha1 = dtpFecha1.value
        .fecha2 = dtpFecha2.value
        .CodTrans = fcbTrans.KeyText
        .NumTrans1 = Val(txtNumTrans1.Text)
        .NumTrans2 = Val(txtNumTrans2.Text)
        s = fcbTrans.KeyText   'PreparaTransParaGnopcion(.CodTrans)
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransparaProduccion", s
        'Graba en la base
        gobjMain.EmpresaActual.GNOpcion.Grabar
    End With
    'Set obj = gobjMain.EmpresaActual.ListaColas
    Set obj = gobjMain.EmpresaActual.Empresa2.ListaGNTrans
    If Not obj.EOF Then
        v = MiGetRows(obj)
        grdOP.Redraw = flexRDNone
        grdOP.LoadArray v
        'ConfigColsColas
        ConfigColsTrans
        grdOP.Redraw = flexRDDirect
    Else
        grdOP.Rows = grd.FixedRows
        ConfigColsTrans
    End If
    cmdGrabarColas.Enabled = True
    cmdGrabarColas.SetFocus
    MensajeStatus "Preparando...", vbNormal
    Exit Sub
errtrap:
    DispErr
    MensajeStatus "Preparando...", vbNormal
    Exit Sub
End Sub

Private Sub cmdBuscar_Click()
    Dim v As Variant, obj As Object, s As String
    On Error GoTo errtrap
    MensajeStatus "Preparando...", vbHourglass
    '*** MAKOTO 06/sep/00 Agregado
    If lstTrans.SelCount = 0 Then
        MsgBox "Seleccione una transacción, por favor.", vbInformation
        Exit Sub
    End If
    
    With gobjMain.objCondicion
        .fecha1 = dtpFecha1.value
        .fecha2 = dtpFecha2.value
'        .CodTrans = fcbTrans.Text              '*** MAKOTO 31/ago/00 Modificado
        .CodTrans = PreparaCodTrans             '***
        
        .NumTrans1 = Val(txtNumTrans1.Text)
        .NumTrans2 = Val(txtNumTrans2.Text)
        
        .EstadoBool(ESTADO_NOAPROBADO) = (chkEstado(ESTADO_NOAPROBADO).value = vbChecked)
        .EstadoBool(ESTADO_APROBADO) = (chkEstado(ESTADO_APROBADO).value = vbChecked)
        .EstadoBool(ESTADO_DESPACHADO) = (chkEstado(ESTADO_DESPACHADO).value = vbChecked)
        .EstadoBool(ESTADO_SEMDESPACHADO) = (chkEstado(ESTADO_SEMDESPACHADO).value = vbChecked) 'AUC
        .EstadoBool(ESTADO_ANULADO) = (chkEstado(ESTADO_ANULADO).value = vbChecked)
'       'jeaa 25/09/06
        s = PreparaTransParaGnopcion(.CodTrans)
        gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransparaRecosteo", s
    'Graba en la base
    gobjMain.EmpresaActual.GNOpcion.Grabar
    End With
    Set obj = gobjMain.EmpresaActual.ConsGNTrans2(True)  'Orden ascendente     '*** MAKOTO 20/oct/00
    If Not obj.EOF Then
        v = MiGetRows(obj)
        grd.Redraw = flexRDNone
        grd.LoadArray v
        ConfigCols
        grd.Redraw = flexRDDirect
    Else
        grd.Rows = grd.FixedRows
        ConfigCols
    End If
    cmdAceptar.Enabled = True
    cmdAceptar.SetFocus
    MensajeStatus "Preparando...", vbNormal
    Exit Sub
errtrap:
    DispErr
    MensajeStatus "Preparando...", vbNormal
    Exit Sub
End Sub

Private Function PreparaCodTrans() As String
    Dim i As Long, s As String
    
    With lstTrans
        'Si está seleccionado solo una
        If lstTrans.SelCount = 1 Then
            For i = 0 To .ListCount - 1
                If .Selected(i) Then
                    s = .List(i)
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

Private Sub ConfigCols()
    With grd
        .FormatString = "^#|tid|<Fecha|<Asiento|<Trans|<#|<#Ref.|<Nombre|<Descripción|<C.Costo|<Estado|<Resultado"
        .ColHidden(COL_NUMFILA) = False
        .ColHidden(COL_TID) = True
        .ColHidden(COL_FECHA) = False
        .ColHidden(COL_CODASIENTO) = True
        .ColHidden(COL_CODTRANS) = False
        .ColHidden(COL_NUMTRANS) = False
        .ColHidden(COL_NUMDOCREF) = True
        .ColHidden(COL_NOMBRE) = False  'True
        .ColHidden(COL_DESC) = False
        .ColHidden(COL_CENTROCOSTO) = True
        '.ColHidden(COL_ESTADO) = True
        .ColDataType(COL_FECHA) = flexDTDate    '*** MAKOTO 14/ago/2000 para que ordene bien por fecha
        .ColDataType(COL_NUMTRANS) = flexDTCurrency
        GNPoneNumFila grd, False
        .AutoSize 0, grd.Cols - 1
        .ColWidth(COL_NUMTRANS) = 1000
        .ColWidth(COL_NOMBRE) = 2500
        .ColWidth(COL_DESC) = 2400
        .ColWidth(COL_RESULTADO) = 2000
        .ColWidth(COL_ESTADO) = 500
    End With
End Sub

Private Sub cmdCancelar_Click()
    If mProcesando Then
        mCancelado = True
    Else
        Unload Me
    End If
End Sub

Private Sub cmdGrabarColas_Click()
Dim i As Long
Dim sql As String
MensajeStatus "Grabando el cambio...", vbHourglass
gc.Grabar False, False
MensajeStatus "Listo...", vbNormal
End Sub

Private Sub cmdTransLimpiar_Click()
    Dim i As Long, aux As Long
    aux = lstTrans.ListIndex
    For i = 0 To lstTrans.ListCount - 1
        lstTrans.Selected(i) = False
    Next i
    lstTrans.ListIndex = aux
End Sub

Private Sub cmdAceptar_Click()
    'Si no hay transacciones
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna transacción para verificar."
        Exit Sub
    End If
    If dtpFecha1 < gobjMain.EmpresaActual.GNOpcion.FechaLimiteDesde Then
        MsgBox "La Rango de Fecha de reproceso es menor a la Fecha Limite Aceptable  ", vbExclamation
        Exit Sub
    End If
    If GeneraHistorial(True, False) Then
        mProcesando = False
    End If
End Sub

Private Sub cmdverificarP_Click()
Dim i As Long, j As Long, k As Long
Dim obj As Object
Dim v As Variant
Dim msg As Variant
Dim bandNoCorrige As Boolean
For i = 1 To grdOP.Rows - 1
lblTitulo = grdOP.TextMatrix(i, 3) & " " & grdOP.TextMatrix(i, 4)
Vuelve_a_cargar:
    'primero cargo las formulas
    Set obj = gobjMain.EmpresaActual.Empresa2.ListaFormulaRubro(grdOP.ValueMatrix(i, 1), True)
    If Not obj.EOF Then
        v = MiGetRows(obj)
        grdFormula.Redraw = flexRDNone
        grdFormula.LoadArray v
        ConfigColsFormula
        grdFormula.Redraw = flexRDDirect
    Else
        grdFormula.Rows = grdFormula.FixedRows
        ConfigColsFormula
    End If
    'aqui verifico con el ivkproceso
    Set obj = gobjMain.EmpresaActual.Empresa2.ListaColas(grdOP.ValueMatrix(i, 1))
    If Not obj.EOF Then
        v = MiGetRows(obj)
        grd.Redraw = flexRDNone
        grd.LoadArray v
        ConfigColsColas
        grd.Redraw = flexRDDirect
    Else
        grd.Rows = grd.FixedRows
        ConfigColsColas
    End If
    'primero reviso si tienen el mismo numero de filas
    If grd.Rows <> grdFormula.Rows Then
        grdOP.TextMatrix(i, grdOP.Cols - 1) = "Error."
    Else
        For j = 1 To grdFormula.Rows - 1
            'verfica proceso y secuencia
            If RevisarFormula(grdFormula.TextMatrix(j, 6), grdFormula.ValueMatrix(j, 7)) Then
                grdFormula.TextMatrix(j, grdFormula.Cols - 1) = "Ok."
            Else
                msg = MsgBox("Desea Proceder a corregir", vbYesNo)
                If msg = vbYes Then
                    CorrigeIVKProceso
                    GoTo Vuelve_a_cargar
'                    grdOP.TextMatrix(i, grdOP.Cols - 1) = "Ok."
                ElseIf msg = vbNo Then
                    bandNoCorrige = True
                    Exit For
                End If
            End If
        Next
        If bandNoCorrige Then
            grdOP.TextMatrix(i, grdOP.Cols - 1) = "NoCorrige."
        Else
            grdOP.TextMatrix(i, grdOP.Cols - 1) = "Ok."
        End If
    End If
Next
End Sub

Private Sub CorrigeIVKProceso()
Dim i As Long
For i = 1 To grdFormula.Rows - 1
    ArregaIVKProceso grdFormula.TextMatrix(i, 6), grdFormula.ValueMatrix(i, 7)
Next i
ActualizaCambios
End Sub
Private Sub ArregaIVKProceso(ByVal codproceso As String, ByVal sec As Integer) 'Arreglo secuencia
Dim i As Long
For i = 1 To grd.Rows - 1
    If grd.TextMatrix(i, 2) = codproceso Then
        grd.TextMatrix(i, 18) = sec
        grd.TextMatrix(i, COL_COLASMOD) = "1"
    End If
Next
End Sub

Private Function RevisarFormula(ByVal codproceso As String, Secuencia As Integer) As Boolean
Dim i As Long
Dim bandSi As Boolean
    For i = 1 To grd.Rows - 1
        If grd.TextMatrix(i, grd.Cols - 1) <> "Ok." Then
            If grd.TextMatrix(i, 2) = codproceso And grd.ValueMatrix(i, 18) = Secuencia Then
                grd.TextMatrix(i, grd.Cols - 1) = "Ok."
                bandSi = True
                Exit For
            End If
        End If
    Next
    RevisarFormula = bandSi
End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF9
        'cmdaceptar1_Click
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
    grdOP.Left = 0
    grdFormula.Move grdOP.Width, grdFormula.Top, Me.ScaleWidth - grdOP.Width, (Me.ScaleHeight - grdFormula.Top - pic1.Height - 80) / 2
    grd.Move grdOP.Width, grdFormula.Top + grdFormula.Height, Me.ScaleWidth - grdOP.Width, grdFormula.Height
    grdOP.Top = grdFormula.Top
    grdOP.Height = grd.Height + grdFormula.Height
    prg1.Width = Me.ScaleWidth - (prg1.Left * 2)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set mColItems = Nothing         '*** MAKOTO 31/ago/00
End Sub

Private Sub grd_AfterEdit(ByVal Row As Long, ByVal col As Long)
Dim obj As IVKProceso
If Not IsObject(grd.RowData(Row)) Then Exit Sub
Set obj = grd.RowData(Row)
    Select Case col
        Case 1
        Case 2: obj.codproceso = grd.TextMatrix(Row, col)
        Case 3: obj.cantidad = grd.ValueMatrix(Row, col)
        Case 4: obj.Orden = grd.ValueMatrix(Row, col)
        Case 5: obj.DescProceso = grd.TextMatrix(Row, col)
        Case 6: obj.FechaInicio = grd.TextMatrix(Row, col)
        Case 7: obj.HoraInicio = grd.TextMatrix(Row, col)
        Case 8: obj.FechaToma = grd.TextMatrix(Row, col)
        Case 9: obj.HoraToma = grd.TextMatrix(Row, col)
        Case 10: obj.FechaFinal = grd.TextMatrix(Row, col)
        Case 11: obj.HoraFinal = grd.TextMatrix(Row, col)
        Case 12: obj.FechaFinEspera = grd.TextMatrix(Row, col)
        Case 13: obj.CodEstado = grd.TextMatrix(Row, col)
        Case 14: obj.idCentroDet = grd.ValueMatrix(Row, col)
        Case 15: obj.CodEstado1 = grd.TextMatrix(Row, col)
        Case 16: obj.codUsuario = grd.TextMatrix(Row, col)
        Case 17: obj.idkpAsignado = grd.ValueMatrix(Row, col)
        Case 18: obj.Secuencia = grd.ValueMatrix(Row, col)
        Case 19: obj.BandGarantia = grd.ValueMatrix(Row, col)
        Case 20: obj.BandUrgente = grd.ValueMatrix(Row, col)
        Case 21: obj.OrdenUrgente = grd.ValueMatrix(Row, col)
        Case 22
        
    End Select
End Sub

Private Sub grd_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyInsert
        AgregaFila
End Select
End Sub

Private Sub grd_RowColChange()
Dim i As Long
Dim j As Long
'If mobjGNComp.SoloVer Then Exit Sub
    For i = 1 To grd.Rows - 1
        If Not grd.IsSubtotal(i) Then
            If i = grd.Row Then
                grd.Cell(flexcpBackColor, i, 2, i, grd.Cols - 1) = &H80000000
            Else
                For j = 1 To grd.Cols - 1
                    If grd.ColData(j) = -1 Then
                        grd.Cell(flexcpBackColor, i, j, i, j) = &H80000018
                    Else
                        grd.Cell(flexcpBackColor, i, j, i, j) = vbWhite
                    End If
                Next
               ' grd.Cell(flexcpBackColor, i, 1, i, grd.Cols - 1) = &H80000005
                'PonerColor
            End If
        End If
    Next
End Sub

Private Sub grdOP_Click()
Dim obj As Object
Dim v As Variant
Dim ivk As IVKProceso
Dim i As Long
'primero cargo las formulas
    lblTitulo = grdOP.TextMatrix(grdOP.Row, 3) & " " & grdOP.TextMatrix(grdOP.Row, 4)
    Set obj = gobjMain.EmpresaActual.Empresa2.ListaFormulaRubro(grdOP.ValueMatrix(grdOP.Row, 1), True)
    If Not obj.EOF Then
        v = MiGetRows(obj)
        grdFormula.Redraw = flexRDNone
        grdFormula.LoadArray v
        ConfigColsFormula
        grdFormula.Redraw = flexRDDirect
    Else
        grdFormula.Rows = grdFormula.FixedRows
        ConfigColsFormula
    End If
        'aqui va los procesos
        Set obj = gobjMain.EmpresaActual.Empresa2.ListaColas(grdOP.ValueMatrix(grdOP.Row, 1))
        If Not obj.EOF Then
            v = MiGetRows(obj)
            grd.Redraw = flexRDNone
            grd.LoadArray v
            ConfigColsColas
            grd.Redraw = flexRDDirect
        Else
            grd.Rows = grd.FixedRows
            ConfigColsColas
        End If
        Set gc = gobjMain.EmpresaActual.RecuperaGNComprobante(grdOP.ValueMatrix(grdOP.Row, 1))
        If Not gc Is Nothing Then
            For i = gc.CountIVKProceso To 1 Step -1
                Set ivk = gc.IVKProceso(i)
                grd.RowData(i) = ivk
            Next
        End If
        grd.ColComboList(2) = gobjMain.EmpresaActual.Empresa2.ListaProcesoFlexGrid(True)
        grd.ColComboList(16) = gobjMain.ListaUsuariosFlexCombo
        grd.ColComboList(13) = gobjMain.EmpresaActual.Empresa2.ListaGNEstadoProdFlexGrid
        grd.ColComboList(15) = gobjMain.EmpresaActual.Empresa2.ListaGNEstadoProdFlexGrid
End Sub

Private Sub txtNumTrans1_KeyPress(KeyAscii As Integer)
    'Acepta solo numericos
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNumTrans2_KeyPress(KeyAscii As Integer)
    'Acepta solo numericos
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Public Sub RecuperaTrans(ByVal Key As String, lst As ListBox, Optional s As String)
Dim Vector As Variant
Dim i As Integer, j As Integer, Selec As Integer
'Dim S As String
    If s <> "_VACIO_" Then
        Vector = Split(s, ",")
         Selec = UBound(Vector, 1)
         For i = 0 To Selec
            For j = 0 To lst.ListCount - 1
'                If Vector(i) = Left(lst.List(j), lst.ItemData(j)) Then
                If Trim(Vector(i)) = lst.List(j) Then
                    lst.Selected(j) = True
                End If
            Next j
         Next i
    End If
End Sub

Private Function PreparaTransParaGnopcion(cad As String) As String
    Dim v As Variant, i As Integer, s As String
    s = ""
    v = Split(cad, ",")
    For i = 0 To UBound(v)
        v(i) = Trim(v(i))
        s = s & Mid$(v(i), 2, Len(v(i)) - 2) & ","
    Next i
    'quita ultima coma
    PreparaTransParaGnopcion = Mid$(s, 1, Len(s) - 1)
End Function
        
Public Sub InicioColas(ByVal colas As String)
    Dim i As Integer
     Me.tag = colas
    On Error GoTo errtrap
    Me.Show
    Me.ZOrder
    cmdBuscar.Visible = False
    cmdBuscaColas.Visible = True
    cmdAceptar.Visible = False
    cmdGrabarColas.Visible = True
    chkEstado(0).Visible = False
    chkEstado(1).Visible = False
    chkEstado(2).Visible = False
    chkEstado(3).Visible = False
    chkEstado(4).Visible = False
    
    grd.Editable = flexEDKbdMouse
    lstTrans.Columns = 10
    grd.AutoSearch = flexSearchNone
    
    dtpFecha1.value = Date
    dtpFecha2.value = Date
    'CargaTransColas
    fcbTrans.SetData gobjMain.EmpresaActual.ListaGNTrans("IV", False, False)
    If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaProduccion")) > 0 Then
        fcbTrans.KeyText = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaProduccion")
    End If
    Exit Sub
errtrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub CargaTransColas()
    Dim i As Long, v As Variant
    Dim s As String
    lstTrans.Clear
     v = gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "IV")
    For i = LBound(v, 2) To UBound(v, 2)
        lstTrans.AddItem v(0, i)        '& " " & v(1, i)
    Next i
End Sub
Private Sub ConfigColsColas()
    With grd                                                                   '
        .FormatString = "^#|Idkp|<CodProceso|^Cant|<Orden|<Descripcion|<FechaIni|<HoraIni|<FechaToma|<HoraToma|<FechaFin|<HoraFin|<FechaFinEsp|<Estado|<IdCentroDet|<Estado1|<CodUsu|<idkpasignado|^Sec|^BandGar|^BandUrg|^OrdenUrg|<Modifica|<Resultado"
        .ColHidden(1) = True
        .ColHidden(4) = True 'Orden
        .ColHidden(5) = True 'Descripcion
        .ColHidden(14) = True
        .ColHidden(17) = True
        .ColHidden(21) = True
        .ColDataType(6) = flexDTDate
        .ColDataType(7) = flexDTDate
        .ColDataType(8) = flexDTDate
        .ColDataType(9) = flexDTDate
        .ColDataType(10) = flexDTDate
        .ColDataType(11) = flexDTDate
        .ColDataType(12) = flexDTDate
        .ColFormat(6) = gobjMain.EmpresaActual.GNOpcion.FormatoFecha
        .ColFormat(7) = "HH:MM:SS"
        .ColFormat(8) = gobjMain.EmpresaActual.GNOpcion.FormatoFecha
        .ColFormat(9) = "HH:MM:SS"
        .ColFormat(10) = gobjMain.EmpresaActual.GNOpcion.FormatoFecha
        .ColFormat(11) = "HH:MM:SS"
        .ColFormat(12) = gobjMain.EmpresaActual.GNOpcion.FormatoFecha
        GNPoneNumFila grd, False
        .AutoSize 0, grd.Cols - 1
    End With
End Sub
Private Sub ConfigColsTrans()
    With grdOP
        .FormatString = "^#|<idtrans|<Fecha|<CodTrans|<NumTrans|<Resultado"
        .ColHidden(1) = True
        .ColDataType(2) = flexDTDate
        GNPoneNumFila grdOP, False
        .AutoSize 0, grdOP.Cols - 1
    End With
End Sub

Private Sub ConfigColsFormula()
    With grdFormula
        .FormatString = "^#|<CodItemPadre|<DescItemPadre|<CodItem|<Descripcion|>Cantidad|<CodProceso|<Orden|<Resultado"
        '.ColHidden(1) = True
        .ColDataType(5) = flexDTDate
        GNPoneNumFila grdFormula, False
        .AutoSize 0, grdFormula.Cols - 1
    End With
End Sub

Private Sub AgregaFila()
Dim ix As Long
Dim r2 As Long
Dim r As Long
    ix = gc.AddIVKProceso
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
        'Asigna el indice de nuevo objeto a la fila nueva
        .RowData(r) = gc.IVKProceso(ix)
        grd.TextMatrix(r, 3) = grd.TextMatrix(r - 1, 3)
        gc.IVKProceso(ix).cantidad = grd.TextMatrix(r, 3)
        grd.TextMatrix(r, 4) = grd.TextMatrix(r - 1, 4)
        gc.IVKProceso(ix).Orden = grd.TextMatrix(r, 4)
        grd.TextMatrix(r, 19) = grd.TextMatrix(r - 1, 19)
        gc.IVKProceso(ix).BandGarantia = grd.TextMatrix(r, 19)
        grd.TextMatrix(r, 20) = grd.TextMatrix(r - 1, 20)
        gc.IVKProceso(ix).BandUrgente = grd.TextMatrix(r, 20)
        grd.TextMatrix(r, 21) = grd.TextMatrix(r - 1, 21)
        gc.IVKProceso(ix).OrdenUrgente = grd.TextMatrix(r, 21)
    End With
    'PoneNumFila
    Exit Sub
errtrap:
    DispErr
    Exit Sub
End Sub

Private Sub ActualizaCambios()
Dim i As Long
Dim sql  As String
For i = 1 To grd.Rows - 1
    If Not grd.IsSubtotal(i) Then
        If grd.ValueMatrix(i, COL_COLASMOD) = 1 Then
            sql = " Update ivkProceso "
            sql = sql & "Set idproceso = (Select idproceso from ivproceso where codproceso = '" & grd.TextMatrix(i, 2) & "')"
            sql = sql & ",cantidad = " & grd.ValueMatrix(i, 3)
            sql = sql & ",Orden = " & grd.TextMatrix(i, 4)
            If Len(grd.TextMatrix(i, 6)) > 0 Then
                sql = sql & ",Descripcion = '" & grd.TextMatrix(i, 5) & "'"
            End If
            If Len(grd.TextMatrix(i, 6)) > 0 Then
                sql = sql & ",FechaInicio = '" & grd.TextMatrix(i, 6) & "'"
            End If
            If Len(grd.TextMatrix(i, 7)) > 0 Then
                sql = sql & ",HoraInicio = '" & grd.TextMatrix(i, 7) & "'"
            End If
            If Len(grd.TextMatrix(i, 8)) > 0 Then
                sql = sql & ",FechaToma= '" & grd.TextMatrix(i, 8) & "'"
            End If
            If Len(grd.TextMatrix(i, 9)) > 0 Then
                sql = sql & ",HoraToma= '" & grd.TextMatrix(i, 9) & "'"
            End If
            If Len(grd.TextMatrix(i, 10)) > 0 Then
                sql = sql & ",FechaFin= '" & grd.TextMatrix(i, 10) & "'"
            End If
            If Len(grd.TextMatrix(i, 11)) > 0 Then
                sql = sql & ",HoraFin= '" & grd.TextMatrix(i, 11) & "'"
            End If
            If Len(grd.TextMatrix(i, 12)) > 0 Then
                sql = sql & ",FechaFinEspera= '" & grd.TextMatrix(i, 12) & "'"
            End If
            If Len(grd.TextMatrix(i, 13)) > 0 Then
                sql = sql & ",Estado = (Select valor from gnestadoprod where codEstado = '" & grd.TextMatrix(i, 13) & "')"
            End If
            If Len(grd.TextMatrix(i, 14)) > 0 Then
                sql = sql & ",IdCentroDet= " & grd.ValueMatrix(i, 14)
            End If
            If Len(grd.TextMatrix(i, 15)) > 0 Then
                sql = sql & ",Estado1 = (Select valor from gnestadoprod where codEstado = '" & grd.TextMatrix(i, 15) & "')"
            End If
            If Len(grd.TextMatrix(i, 16)) > 0 Then
                sql = sql & ",CodUsuario= '" & grd.TextMatrix(i, 16) & "'"
            End If
            If Len(grd.TextMatrix(i, 17)) > 0 Then
                sql = sql & ",idkpAsignado= " & grd.ValueMatrix(i, 17)
            End If
            If Len(grd.TextMatrix(i, 18)) > 0 Then
                sql = sql & ",Secuencia= " & grd.ValueMatrix(i, 18)
            End If
            If Len(grd.TextMatrix(i, 19)) > 0 Then
                sql = sql & ",BandGarantia= " & grd.ValueMatrix(i, 19)
            End If
            If Len(grd.TextMatrix(i, 20)) > 0 Then
                sql = sql & ",BandUrgente= " & grd.ValueMatrix(i, 20)
            End If
            If Len(grd.TextMatrix(i, 21)) > 0 Then
                sql = sql & ",OrdenUrgente= " & grd.ValueMatrix(i, 21)
            End If
            sql = sql & "  WHERE idkp  = " & grd.ValueMatrix(i, 1)
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
            grd.TextMatrix(i, grd.Cols - 1) = "OK"
        End If
    End If
Next

End Sub
