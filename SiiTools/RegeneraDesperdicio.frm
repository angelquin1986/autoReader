VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRegeneraDesperdicio 
   Caption         =   "Regeneración de Recetas"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6585
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   6585
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   852
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   6585
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3855
      Width           =   6585
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Proceder"
         Enabled         =   0   'False
         Height          =   372
         Left            =   1800
         TabIndex        =   12
         Top             =   0
         Width           =   1212
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   6540
         TabIndex        =   11
         Top             =   60
         Width           =   1212
      End
      Begin VB.CommandButton cmdVerificar 
         Caption         =   "&Verificar"
         Enabled         =   0   'False
         Height          =   372
         Left            =   288
         TabIndex        =   10
         Top             =   0
         Width           =   1212
      End
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   540
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   4395
      Left            =   120
      TabIndex        =   8
      Top             =   1980
      Width           =   6735
      _cx             =   11880
      _cy             =   7752
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
      Left            =   9540
      TabIndex        =   7
      Top             =   1320
      Width           =   1212
   End
   Begin VB.Frame fraFecha 
      Caption         =   "&Fecha (desde - hasta)"
      Height          =   1455
      Left            =   402
      TabIndex        =   0
      Top             =   120
      Width           =   1932
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
   Begin VB.Frame fraCodTrans 
      Caption         =   "Cod.&Trans."
      Height          =   1455
      Left            =   2340
      TabIndex        =   3
      Top             =   120
      Width           =   4155
      Begin VB.ListBox lstTrans 
         Columns         =   3
         Height          =   1095
         IntegralHeight  =   0   'False
         Left            =   60
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   17
         Top             =   240
         Width           =   4035
      End
   End
   Begin VB.Frame fraNumTrans 
      Caption         =   "# T&rans. (desde - hasta)"
      Height          =   1092
      Left            =   6540
      TabIndex        =   4
      Top             =   180
      Width           =   1932
      Begin VB.TextBox txtNumTrans1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   360
         TabIndex        =   5
         Top             =   280
         Width           =   1212
      End
      Begin VB.TextBox txtNumTrans2 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   360
         TabIndex        =   6
         Top             =   640
         Width           =   1212
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grdReceta 
      Height          =   4575
      Left            =   120
      TabIndex        =   14
      Top             =   6480
      Width           =   6735
      _cx             =   11880
      _cy             =   8070
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
   Begin VSFlex7LCtl.VSFlexGrid grdRecetaIng 
      Height          =   4995
      Left            =   6960
      TabIndex        =   18
      Top             =   6480
      Width           =   6675
      _cx             =   11774
      _cy             =   8811
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
   Begin VSFlex7LCtl.VSFlexGrid grdOut 
      Height          =   4395
      Left            =   6900
      TabIndex        =   20
      Top             =   1980
      Width           =   6735
      _cx             =   11880
      _cy             =   7752
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
   Begin VB.Label Label3 
      Caption         =   "Lista de Items Ingreso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12060
      TabIndex        =   19
      Top             =   900
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Lista de Items Egreso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   16
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Lista de Transacciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   2115
   End
End
Attribute VB_Name = "frmRegeneraDesperdicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Constantes para las columnas
Private Const COL_NUMFILA = 0
Private Const COL_TID = 1
Private Const COL_FECHA = 2
Private Const COL_CODASIENTO = 3
Private Const COL_CODTRANS = 4
Private Const COL_NUMTRANS = 5
Private Const COL_NUMDOCREF = 6     '*** MAKOTO 07/feb/01 Agregado
Private Const COL_NOMBRE = 7        '*** MAKOTO 07/feb/01 Agregado
Private Const COL_DESC = 8
Private Const COL_CENTROCOSTO = 9
Private Const COL_ESTADO = 10
Private Const COL_RESULTADO = 11

Private Const TIPORECETA = 4
Private Const COL_VENTA_ID = 1
Private Const COL_VENTA_IDINV = 2
Private Const COL_VENTA_CANT = 5
Private Const COL_VENTA_TIPO = 6
Private Const COL_VENTA_IDIDPADRE = 7

Private Const COL_ITEM_ID = 1
Private Const COL_ITEM_IDINV = 2
Private Const COL_ITEM_CANT = 5
Private Const COL_ITEM_TIPO = 6
Private Const COL_ITEM_IDIDPADRE = 7

Private Const COL_RECETA_IDINV = 1
Private Const COL_RECETA_CANT = 4

Private Const MSG_NG = "Receta incorrecta."
Private mProcesando As Boolean
Private mCancelado As Boolean
Private mVerificado As Boolean
Private num_fila_trans As Long
Private num_fila_itemVenta As Long
Private num_fila_Receta As Long
Private num_fila_IVkitem As Long
Private IdPadreT As Long
Private mobjGNComp As GNComprobante
Private mobjGNCompAux As GNComprobante


Private Sub CargaTrans()
    Dim i As Long, v As Variant
    Dim s As String
    'Carga la lista de transacción
'    fcbTrans.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "IV")

    lstTrans.Clear
    v = gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "IV")
    For i = LBound(v, 2) To UBound(v, 2)
        lstTrans.AddItem v(0, i)        '& " " & v(1, i)
    Next i
    
    If tag = "Desperdicio" Then
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaDesperdicio")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaDesperdicio")
            RecuperaTrans "KeyT", lstTrans, s
        End If
    End If
End Sub

Private Sub cmdAceptar_Click()
     'Si no hay transacciones
'    If grd.Rows <= grd.FixedRows Then
'        MsgBox "No hay ningúna transacción para procesar."
'        Exit Sub
'    End If
'
'    If dtpFecha1 < gobjMain.EmpresaActual.GNOpcion.FechaLimiteDesde Then
'        MsgBox "La Rango de Fecha de regeneración es menor a la Fecha Limite Aceptable  ", vbExclamation
'        Exit Sub
'    End If
'
'    If grdReceta.Rows = 1 Then MsgBox "No hay nada que procesar": Exit Sub
'    CorrigeDesperdicio
    grdReceta.Rows = 1
    mVerificado = True
    'Si no hay transacciones
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna transacción para verificar."
        Exit Sub
    End If
    If Me.tag = "Desperdicio" Then
        If RegenerarDesperdicio(True, False) Then
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
            mVerificado = True
        End If
    End If

End Sub

Private Function RegenerarAsientoSub(ByVal gnc As GNComprobante, _
                                     ByRef cambiado As Boolean) As Boolean
    Dim i As Long, cta As CtCuenta, ctd As CTLibroDetalle
    Dim colCtd As Collection, a As clsAsiento
    On Error GoTo ErrTrap
    
    cambiado = False
    Set colCtd = New Collection
    
    'Guarda todos los detalles de asiento en la colección para después comparar
    With gnc
        For i = 1 To .CountCTLibroDetalle
            Set ctd = .CTLibroDetalle(i)
            Set a = New clsAsiento
            a.IdCuenta = ctd.IdCuenta
            a.Debe = ctd.Debe
            a.Haber = ctd.Haber
            colCtd.Add item:=a
        Next i
    End With
    
    'Regenera el asiento
    gnc.GeneraAsiento
    
    'Compara el asiento para saber si ha cambiado o no
    cambiado = Not CompararAsiento(gnc, colCtd)
    
    RegenerarAsientoSub = True
    GoTo salida
    Exit Function
ErrTrap:
    cambiado = False
    DispErr
    RegenerarAsientoSub = False
salida:
    Set a = Nothing
    Set colCtd = Nothing
    Set gnc = Nothing
    Exit Function
End Function


'Devuelve True si los asientos son iguales, False si no lo son
Private Function CompararAsiento(ByVal gnc As GNComprobante, ByVal col As Collection) As Boolean
    Dim a As clsAsiento, i As Long, ctd As CTLibroDetalle
    Dim encontrado As Boolean
    
    'Si número de detalles son diferentes ya no son iguales
    If col.Count <> gnc.CountCTLibroDetalle Then Exit Function
    
    For i = 1 To gnc.CountCTLibroDetalle
        Set ctd = gnc.CTLibroDetalle(i)
        encontrado = False
        For Each a In col
            If (ctd.IdCuenta = a.IdCuenta) And _
               (ctd.Debe = a.Debe) And _
               (ctd.Haber = a.Haber) And _
               (a.Comparado = False) Then
                a.Comparado = True
                encontrado = True
                Exit For
            End If
        Next a
        'Si no se encuentra uno igual
        If Not encontrado Then
            CompararAsiento = False
            Exit Function
        End If
    Next i
    CompararAsiento = True
End Function

Private Sub cmdBuscar_Click()
    Dim v As Variant, obj As Object, s As String
    On Error GoTo ErrTrap
'    If Len(fcbGrupoDesde.KeyText) = 0 Then
'        MsgBox "deberia escoger el grupo para aplicar"
'        Exit Sub
'    End If
    grd.Rows = 1
    grdReceta.Rows = 1
    grdRecetaIng.Rows = 1
    mVerificado = False
    With gobjMain.objCondicion
        .fecha1 = dtpFecha1.value
        .fecha2 = dtpFecha2.value
        .CodTrans = PreparaCodTrans
        .NumTrans1 = Val(txtNumTrans1.Text)
        .NumTrans2 = Val(txtNumTrans2.Text)
        'Estados no incluye anulados
        s = PreparaTransParaGnopcion(.CodTrans)
        If Me.tag = "Desperdicio" Then
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransparaDesperdicio", s
        End If
        gobjMain.EmpresaActual.GNOpcion.Grabar
        
    End With
    Set obj = gobjMain.EmpresaActual.ConsGNTransTransformacion(True)
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
    cmdVerificar.Enabled = True
    cmdVerificar.SetFocus
    cmdAceptar.Enabled = False
    mVerificado = False
    Exit Sub
ErrTrap:
    DispErr
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
        .FormatString = "^#|tid|<Fecha|<Asiento|<Trans|>#|<#Ref.|<Nombre|<Descripción|<C.Costo|<Estado|<Resultado|||||<Cant"
        .ColHidden(COL_NUMFILA) = False
        .ColHidden(COL_TID) = False
        .ColHidden(COL_FECHA) = False
        .ColHidden(COL_CODASIENTO) = True
        .ColHidden(COL_CODTRANS) = False
        .ColHidden(COL_NUMTRANS) = False
        .ColHidden(COL_NUMDOCREF) = True
        .ColHidden(COL_NOMBRE) = False      'True
        .ColHidden(COL_DESC) = False
        .ColHidden(COL_CENTROCOSTO) = True
        .ColHidden(1) = True
        .ColHidden(8) = True
        .ColHidden(9) = True
        .ColHidden(10) = True
        If .Cols > 12 Then
            .ColHidden(12) = True
            .ColHidden(13) = True
            .ColHidden(14) = True
            .ColHidden(15) = True
        End If
        '.ColHidden(COL_ESTADO) = True
        .ColDataType(COL_FECHA) = flexDTDate    '*** MAKOTO 14/ago/2000 para que ordene bien por fecha
        GNPoneNumFila grd, False
        .AutoSize 0, grd.Cols - 1
        .ColWidth(COL_NUMTRANS) = 700
        .ColWidth(COL_NOMBRE) = 2000
        .ColWidth(COL_DESC) = 2400
        .ColWidth(COL_RESULTADO) = 2000
        
    End With
    
    With grdOut
        .FormatString = "^#|tid|<Fecha|<Asiento|<Trans|>#|<#Ref.|<Nombre|<Descripción|<C.Costo|<Estado|<Resultado|||||<Cant"
        .ColHidden(COL_NUMFILA) = False
        .ColHidden(COL_TID) = False
        .ColHidden(COL_FECHA) = False
        .ColHidden(COL_CODASIENTO) = True
        .ColHidden(COL_CODTRANS) = False
        .ColHidden(COL_NUMTRANS) = False
        .ColHidden(COL_NUMDOCREF) = True
        .ColHidden(COL_NOMBRE) = False      'True
        .ColHidden(COL_DESC) = False
        .ColHidden(COL_CENTROCOSTO) = True
        .ColHidden(1) = True
        .ColHidden(8) = True
        .ColHidden(9) = True
        .ColHidden(10) = True
        If .Cols > 12 Then
            .ColHidden(12) = True
            .ColHidden(13) = True
            .ColHidden(14) = True
            .ColHidden(15) = True
        End If
        '.ColHidden(COL_ESTADO) = True
        .ColDataType(COL_FECHA) = flexDTDate    '*** MAKOTO 14/ago/2000 para que ordene bien por fecha
        GNPoneNumFila grd, False
        .AutoSize 0, grd.Cols - 1
        .ColWidth(COL_NUMTRANS) = 700
        .ColWidth(COL_NOMBRE) = 2000
        .ColWidth(COL_DESC) = 2400
        .ColWidth(COL_RESULTADO) = 2000
        
    End With
    
    
    With grdReceta
        .FormatString = "^#|ID|Trans|idInventario|<Cod.Inv|<Descripcion|>Cantidad|>CT|>IdPadre|<Resultado|>Cant.Real|>Id|>PU|>Ast "
        GNPoneNumFila grd, False
        .AutoSize 0, .Cols - 1
        .ColWidth(1) = 700
        .ColWidth(2) = 700
        .ColWidth(3) = 700
        .ColWidth(4) = 800
        .ColWidth(5) = 3500
        .ColWidth(6) = 800
        .ColWidth(7) = 1500
        .ColWidth(8) = 500
        .ColWidth(9) = 800
        .ColHidden(1) = True
'        .ColHidden(2) = True
'        .ColHidden(5) = True
        .ColHidden(3) = True
'        .ColHidden(7) = True
        .ColHidden(8) = True
        .ColHidden(10) = True
        .ColHidden(11) = True
        .ColHidden(12) = True
        .ColHidden(13) = True

    End With
    
With grdRecetaIng
        .FormatString = "^#|ID|Trans|idInventario|<Cod.Inv|<Descripcion|>Cantidad|>CT|>IdPadre|<Resultado|>Cant.Real|>Id|>PU|>Ast "
        GNPoneNumFila grd, False
        .AutoSize 0, .Cols - 1
        .ColWidth(1) = 700
        .ColWidth(2) = 700
        .ColWidth(3) = 700
        .ColWidth(4) = 800
        .ColWidth(5) = 3500
        .ColWidth(6) = 800
        .ColWidth(7) = 1500
        .ColWidth(8) = 500
        .ColWidth(9) = 800
        .ColHidden(1) = True
'        .ColHidden(2) = True
'        .ColHidden(4) = True
        .ColHidden(3) = True
'        .ColHidden(7) = True
        .ColHidden(8) = True
        .ColHidden(10) = True
        .ColHidden(11) = True
        .ColHidden(12) = True
        .ColHidden(13) = True

    End With
    
End Sub

Private Sub cmdCancelar_Click()
    If mProcesando Then
        mCancelado = True
    Else
        Unload Me
    End If
End Sub



Private Sub cmdVerificar_Click()
    grdReceta.Rows = 1
    mVerificado = False
    'Si no hay transacciones
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna transacción para verificar."
        Exit Sub
    End If
    If Me.tag = "Desperdicio" Then
        If RegenerarDesperdicio(True, False) Then
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
            mVerificado = True
        End If
    End If
    
    
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
    If mProcesando Then
        Cancel = True
        Exit Sub
    End If
    Me.Hide         'Se pone esto para evitar el posible BUG de Windows98
End Sub

Private Sub Form_Resize()
    On Error Resume Next
'    grd.Move 0, grd.Top, Me.ScaleWidth, (Me.ScaleHeight - grd.Top - pic1.Height - 80)
    grd.Move 0, grd.Top, Me.ScaleWidth / 2, (Me.ScaleHeight - grd.Top - pic1.Height - 80) * 0.7
    grdReceta.Move 0, grd.Top + grd.Height, Me.ScaleWidth / 2, (Me.ScaleHeight - grd.Top - pic1.Height - 80) * 0.3
    
    grdOut.Move grd.Width, grd.Top, Me.ScaleWidth / 2, (Me.ScaleHeight - grd.Top - pic1.Height - 80) * 0.7
    grdRecetaIng.Move grd.Width, grdOut.Top + grdOut.Height, Me.ScaleWidth / 2, (Me.ScaleHeight - grd.Top - pic1.Height - 80) * 0.3
    Label2.Left = grdReceta.Left
    prg1.Width = Me.ScaleWidth - (prg1.Left * 2)
    Label3.Top = grdRecetaIng.Top - 200
    
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

'jeaa 25/09/2006 elimina los apostrofes
Private Function PreparaTransParaGnopcion(cad As String) As String
    Dim v As Variant, i As Integer, s As String
    s = ""
    v = Split(cad, ",")
    If UBound(v) > 0 Then
        For i = 0 To UBound(v)
            v(i) = Trim(v(i))
            s = s & Mid$(v(i), 2, Len(v(i)) - 2) & ","
        Next i
    Else
        s = cad & ","
    End If
    'quita ultima coma
    PreparaTransParaGnopcion = Mid$(s, 1, Len(s) - 1)
End Function

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






Public Sub InicioDesperdicio(ByVal tag As String)
    Dim i As Integer
    On Error GoTo ErrTrap
    Me.tag = tag
    Me.Show
    Me.ZOrder
    dtpFecha1.value = gobjMain.EmpresaActual.GNOpcion.FechaLimiteDesde
    dtpFecha2.value = Date
    CargaTrans
    
    
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Function RegenerarDesperdicio(bandVerificar As Boolean, BandTodo As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, x As Single
    Dim gnc As GNComprobante, cambiado As Boolean
    On Error GoTo ErrTrap

    'Si no es solo verificacion, confirma
    If Not bandVerificar Then
        s = "Desea volver a Verificar la transacción seleccionada." & vbCr & vbCr
        s = s & "Está seguro que desea proceder?"
        If MsgBox(s, vbYesNo + vbQuestion) <> vbYes Then
            Exit Function
        Else
            grdReceta.Rows = 1
        End If
    End If
    
    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdVerificar.Enabled = False
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
        x = grd.CellTop                 'Para visualizar la celda actual
        'Si es verificación, procesa todas las filas sino solo las que tengan "Asiento incorrecto."
        If (grd.TextMatrix(i, COL_RESULTADO) = MSG_NG) Or bandVerificar Or BandTodo Then
            tid = grd.ValueMatrix(i, COL_TID)
            grd.TextMatrix(i, COL_RESULTADO) = "Verificando..."
            grd.Refresh
            'Recupera la transaccion
            Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
            If Not (gnc Is Nothing) Then
                'Si la transacción no está anulada
                If gnc.Estado <> ESTADO_ANULADO Then
                    'Forzar recuperar todos los datos de transacción para que no se pierdan al grabar de nuveo
                    gnc.RecuperaDetalleTodo
                    cargaItemsKardexDesperdicio gnc, i, grd.TextMatrix(i, 14)
                    ArreglaDesperdicio gnc, i, mVerificado, grd.TextMatrix(i, 14)
                    num_fila_trans = i
                Else
                    'Si está anulada
                    grd.TextMatrix(i, COL_RESULTADO) = "Anulado."
                End If
            Else
                grd.TextMatrix(i, COL_RESULTADO) = "No pudo recuperar la transación."
            End If
        End If
    Next i
    Screen.MousePointer = 0
    RegenerarDesperdicio = Not mCancelado
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    DispErr
salida:
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdVerificar.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    Exit Function
End Function

Private Sub cargaItemsKardexDesperdicio(ByRef gnc As GNComprobante, ByVal i As Long, Optional coditem As String)
''''    Dim j As Long, rs1 As Recordset
''''    Dim item As IVinventario
''''    Dim mbooBand As Boolean, idInve As Long, codInve As String, Desc As String, Tipo As Integer
''''    Dim x As Long, s As String
''''
''''    Dim gncOut As GNComprobante
''''    'carga la el detalle transaccion
''''
''''    For j = 1 To grdReceta.Rows - 1
''''        grdReceta.RemoveItem 1
''''    Next j
''''
''''    For j = 1 To grdRecetaIng.Rows - 1
''''        grdRecetaIng.RemoveItem 1
''''    Next j
''''
''''    For j = 1 To grdOut.Rows - 1
''''        grdOut.RemoveItem 1
''''    Next j
''''
''''
''''
''''
''''
''''    For j = 1 To gnc.CountIVKardex
''''
''''        Set item = gnc.Empresa.RecuperaIVInventarioQuick(gnc.IVKardex(j).CodInventario)
''''        s = gnc.IVKardex(j).orden & vbTab & gnc.IVKardex(j).id & vbTab & gnc.numtrans & vbTab & gnc.IVKardex(j).idinventario & vbTab & gnc.IVKardex(j).CodInventario & vbTab & item.Descripcion & vbTab & gnc.IVKardex(j).cantidad & vbTab & gnc.IVKardex(j).CostoTotal & vbTab & gnc.TransID & vbTab & (gnc.IVKardex(j).DescuentoOriginal * 100)
''''
''''
''''        If gnc.IVKardex(j).cantidad > 1 Then
''''            If Len(s) > 0 Then          '*** MAKOTO 09/nov/00 para no agregar items de destino en Trans. Bodegas
''''                With grdReceta
''''                    .AddItem s
''''                End With
''''            End If
''''        End If
''''    Next j
''''    Set item = Nothing
''''
''''    If grdReceta.ValueMatrix(1, 9) = 0 Then
''''        grd.TextMatrix(i, COL_RESULTADO) = "OK Sin Desperdicio " & grdReceta.ValueMatrix(1, 9)
''''        Exit Sub
''''    End If
''''
''''
''''        Set gncOut = gobjMain.EmpresaActual.RecuperaGNComprobantexIdTrandfuente(gnc.TransID)
''''
''''
''''
''''        If Not gncOut Is Nothing Then
''''            s = "1" & vbTab & gncOut.TransID & vbTab & gncOut.FechaTrans & vbTab & vbTab & gncOut.CodTrans & vbTab & gncOut.numtrans & vbTab & gncOut.nombre & vbTab & vbTab & gncOut.Descripcion & vbTab & gncOut.CodCentro & vbTab & gncOut.Estado & vbTab & gncOut.numtrans
''''        Else
''''            grd.TextMatrix(i, COL_RESULTADO) = "Error"
''''            Exit Sub
''''        End If
''''        grdOut.AddItem s
''''
''''
''''
''''
''''    For j = 1 To gncOut.CountIVKardex
''''
''''        Set item = gncOut.Empresa.RecuperaIVInventarioQuick(gncOut.IVKardex(j).CodInventario)
''''        s = gncOut.IVKardex(j).orden & vbTab & gncOut.IVKardex(j).id & vbTab & gncOut.numtrans & vbTab & gncOut.IVKardex(j).idinventario & vbTab & gncOut.IVKardex(j).CodInventario & vbTab & item.Descripcion & vbTab & gncOut.IVKardex(j).cantidad & vbTab & gncOut.IVKardex(j).CostoTotal & vbTab & gncOut.TransID & vbTab & gncOut.IVKardex(j).DescuentoOriginal
''''
''''
''''        If gncOut.IVKardex(j).cantidad < 1 Then
''''            If Len(s) > 0 Then          '*** MAKOTO 09/nov/00 para no agregar items de destino en Trans. Bodegas
''''                With grdRecetaIng
''''                    .AddItem s
''''                End With
''''            End If
''''        End If
''''    Next j
''''    Set item = Nothing
''''    Dim val1 As Currency, val2 As Currency
''''
''''    val1 = grdRecetaIng.ValueMatrix(1, 6)
''''    val2 = Round((grdReceta.ValueMatrix(1, 9) / 100) * grdReceta.ValueMatrix(1, 6), 0)
''''    'grdReceta.subtotal flexSTSum, 2, 7, , , , , , , True
''''    If (val1 + val2) <> 0 Then
''''        'MsgBox gnc.numtrans
''''        grd.TextMatrix(i, COL_RESULTADO) = "Error"
''''    Else
''''        If (Int(val1) + val2) <> 0 Then
''''            grd.TextMatrix(i, COL_RESULTADO) = "Error"
''''        Else
''''            grd.TextMatrix(i, COL_RESULTADO) = "OK"
''''        End If
''''    End If
    
    
End Sub


Private Function CorrigeDesperdicio()  'ByVal idInven As Long, ByVal cant As Currency, ByVal idPadres As Long, gnc As GNComprobante) As Boolean
    Dim i As Long, sql As String, rs As Recordset, fila As Integer
    Dim j As Long, idsubItem As Long, CantSubItem As Currency, costo As Currency
    Dim item As IVinventario, tid As Long, cant As Currency
    Dim gnc As GNComprobante, x As Long
    Dim c As Currency, orden As Integer
    Dim IdBod As Long, BODAUX As String, idAsignado As Long
    Dim ivk As IVKardex, gnt As GNTrans, ix As Long
    On Error GoTo ErrTrap
    
    
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    
    For i = 1 To grd.Rows - 1
     DoEvents
     prg1.value = i
     grd.Row = i
     x = grd.CellTop
        
        If grd.TextMatrix(i, 11) = "Error" Then
            grdReceta.Refresh
            If grdOut.Rows > 1 Then
                Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(grdOut.ValueMatrix(1, 1))
                fila = gnc.CountIVKardex
                For j = 1 To fila
                    Set ivk = gnc.IVKardex(j)
                    gnc.IVKardex(j).cantidad = Round((grdReceta.ValueMatrix(1, 9) / 100 * grdReceta.ValueMatrix(1, 6)) * -1, 0)
                    gnc.IVKardex(j).CostoRealTotal = (grdReceta.ValueMatrix(1, 7) / grdReceta.ValueMatrix(1, 6)) * gnc.IVKardex(j).cantidad
                    gnc.IVKardex(j).CostoTotal = ivk.CostoRealTotal
                Next j
                
                If gnc.CountPCKardexCHP > 0 Then
                    gnc.PCKardexCHP(1).Debe = Abs(ivk.CostoRealTotal)
                End If
                
                gnc.Grabar False, False
                grd.TextMatrix(i, 11) = "Corregido  "
                x = grd.CellTop                 'Para visualizar la celda actual
            End If

        
    End If
Next i
        Screen.MousePointer = 0
        CorrigeDesperdicio = True
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    DispErr
salida:
    Set gnc = Nothing
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdVerificar.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    Exit Function

End Function

Private Sub ArreglaDesperdicio(ByRef gnc As GNComprobante, ByVal i As Long, ByVal bandcorrige As Boolean, Optional coditem As String)

    Dim j As Long, rs1 As Recordset
    Dim item As IVinventario
    Dim mbooBand As Boolean, idInve As Long, codInve As String, Desc As String, Tipo As Integer
    Dim x As Long, s As String
    Dim fila  As Integer, ivk As IVKardex
    
    Dim gncOut As GNComprobante
    'carga la el detalle transaccion
    
    For j = 1 To grdReceta.Rows - 1
        grdReceta.RemoveItem 1
    Next j
    
    For j = 1 To grdRecetaIng.Rows - 1
        grdRecetaIng.RemoveItem 1
    Next j
    
    For j = 1 To grdOut.Rows - 1
        grdOut.RemoveItem 1
    Next j
    
    
    
    
    
    For j = 1 To gnc.CountIVKardex
        
        Set item = gnc.Empresa.RecuperaIVInventarioQuick(gnc.IVKardex(j).CodInventario)
        s = gnc.IVKardex(j).orden & vbTab & gnc.IVKardex(j).id & vbTab & gnc.numtrans & vbTab & gnc.IVKardex(j).idinventario & vbTab & gnc.IVKardex(j).CodInventario & vbTab & item.Descripcion & vbTab & gnc.IVKardex(j).cantidad & vbTab & gnc.IVKardex(j).CostoTotal & vbTab & gnc.TransID & vbTab & (gnc.IVKardex(j).DescuentoOriginal * 100)
        
        
        If gnc.IVKardex(j).cantidad > 0 Then
            If Len(s) > 0 Then          '*** MAKOTO 09/nov/00 para no agregar items de destino en Trans. Bodegas
                With grdReceta
                    .AddItem s
                End With
            End If
        End If
    Next j
    Set item = Nothing
    
    If grdReceta.ValueMatrix(1, 9) = 0 Then
        grd.TextMatrix(i, COL_RESULTADO) = "OK Sin Desperdicio " & grdReceta.ValueMatrix(1, 9)
        Exit Sub
    End If

    
        Set gncOut = gobjMain.EmpresaActual.RecuperaGNComprobantexIdTrandfuente(gnc.TransID)
        
        
        
        If Not gncOut Is Nothing Then
            s = "1" & vbTab & gncOut.TransID & vbTab & gncOut.FechaTrans & vbTab & vbTab & gncOut.CodTrans & vbTab & gncOut.numtrans & vbTab & gncOut.nombre & vbTab & vbTab & gncOut.Descripcion & vbTab & gncOut.CodCentro & vbTab & gncOut.Estado & vbTab & gncOut.numtrans
        Else
            grd.TextMatrix(i, COL_RESULTADO) = "Error"
            If bandcorrige Then
                GrabarTransDesperdicioAuto gnc, "ELD"
                grd.TextMatrix(i, COL_RESULTADO) = "OK se creo Desperdicio"
            End If
            
            Exit Sub
        End If
        grdOut.AddItem s
    
    
    
    
    For j = 1 To gncOut.CountIVKardex
        
        Set item = gncOut.Empresa.RecuperaIVInventarioQuick(gncOut.IVKardex(j).CodInventario)
        s = gncOut.IVKardex(j).orden & vbTab & gncOut.IVKardex(j).id & vbTab & gncOut.numtrans & vbTab & gncOut.IVKardex(j).idinventario & vbTab & gncOut.IVKardex(j).CodInventario & vbTab & item.Descripcion & vbTab & gncOut.IVKardex(j).cantidad & vbTab & gncOut.IVKardex(j).CostoTotal & vbTab & gncOut.TransID & vbTab & gncOut.IVKardex(j).DescuentoOriginal
        
        
        If gncOut.IVKardex(j).cantidad < 0 Then
            If Len(s) > 0 Then          '*** MAKOTO 09/nov/00 para no agregar items de destino en Trans. Bodegas
                With grdRecetaIng
                    .AddItem s
                End With
            End If
        End If
    Next j
    Set item = Nothing
    Dim val1 As Currency, val2 As Currency
    
    val1 = grdRecetaIng.ValueMatrix(1, 6)
    val2 = Round((grdReceta.ValueMatrix(1, 9) / 100) * grdReceta.ValueMatrix(1, 6), 0)
    'grdReceta.subtotal flexSTSum, 2, 7, , , , , , , True
    If (val1 + val2) <> 0 Then
        'MsgBox gnc.numtrans
        grd.TextMatrix(i, COL_RESULTADO) = "Error"
    Else
        If (Int(val1) + val2) <> 0 Then
            grd.TextMatrix(i, COL_RESULTADO) = "Error"
        Else
            grd.TextMatrix(i, COL_RESULTADO) = "OK"
        End If
    End If
    
    If bandcorrige Then
        If grd.TextMatrix(i, 11) = "Error" Then
            grdReceta.Refresh
            If grdOut.Rows > 1 Then
                Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(grdOut.ValueMatrix(1, 1))
                fila = gnc.CountIVKardex
                For j = 1 To fila
                    Set ivk = gnc.IVKardex(j)
                    gnc.IVKardex(j).cantidad = Round((grdReceta.ValueMatrix(1, 9) / 100 * grdReceta.ValueMatrix(1, 6)) * -1, 0)
                    gnc.IVKardex(j).CostoRealTotal = (grdReceta.ValueMatrix(1, 7) / grdReceta.ValueMatrix(1, 6)) * gnc.IVKardex(j).cantidad
                    gnc.IVKardex(j).CostoTotal = ivk.CostoRealTotal
                Next j
                
                If gnc.CountPCKardexCHP > 0 Then
                    gnc.PCKardexCHP(1).Debe = Abs(ivk.CostoRealTotal)
                End If
                
                gnc.Grabar False, False
                grd.TextMatrix(i, 11) = "Corregido  "
                x = grd.CellTop                 'Para visualizar la celda actual
            End If
        End If
    End If


End Sub


Private Function GrabarTransDesperdicioAuto(gc As GNComprobante, ByVal CodTrans As String) As Boolean
    Dim Imprime As Boolean, i As Long, ix As Long, j As Integer
    Dim item As IVinventario, rsReceta As Recordset
    Dim Cadena As String, peso As Currency, c As Currency, v As Variant
    Dim codRelleno As String, codCemento As String, codCojin As String ', codParche As String
    Dim porRelleno As Currency, porCemento As Currency, porCojin As Currency
    Dim costoRelleno As Currency, costoCemento As Currency, costoCojin As Currency
    Dim costoPintura As Currency, porPintura As Currency, codPintura As String
    Dim bandMerma As Boolean

    On Error GoTo ErrTrap
    
    Set mobjGNComp = gc
    bandMerma = False
    
    For i = 1 To mobjGNComp.CountIVKardex
        If mobjGNComp.IVKardex(i).DescuentoOriginal <> 0 Then
            bandMerma = True
        End If
    Next i
    
    If Not bandMerma Then
        GrabarTransDesperdicioAuto = False
        Exit Function
    End If
    
    Set mobjGNCompAux = mobjGNComp.Empresa.CreaGNComprobante(CodTrans)
    
    If Not mobjGNCompAux Is Nothing Then
    
        If mobjGNCompAux.SoloVer Then
            MsgBox MSG_NODISPONE, vbInformation
            Exit Function
        End If
        
        
        'carga los hijos de los items seleccionados
        For i = 1 To mobjGNComp.CountIVKardex
            If mobjGNComp.IVKardex(i).cantidad > 0 Then
                If mobjGNComp.IVKardex(i).DescuentoOriginal <> 0 Then
                    ix = mobjGNCompAux.AddIVKardex
                    mobjGNCompAux.IVKardex(ix).CodBodega = mobjGNComp.IVKardex(i).CodBodega
                    mobjGNCompAux.IVKardex(ix).CodInventario = mobjGNComp.IVKardex(i).CodInventario
                    mobjGNCompAux.IVKardex(ix).cantidad = Round((mobjGNComp.IVKardex(i).cantidad * (mobjGNComp.IVKardex(i).DescuentoOriginal)), 0) * -1
                    mobjGNCompAux.IVKardex(ix).CostoRealTotal = mobjGNCompAux.IVKardex(ix).cantidad * (mobjGNComp.IVKardex(i).CostoRealTotal / mobjGNComp.IVKardex(i).cantidad)
                    mobjGNCompAux.IVKardex(ix).CostoTotal = mobjGNCompAux.IVKardex(ix).cantidad * (mobjGNComp.IVKardex(i).CostoTotal / mobjGNComp.IVKardex(i).cantidad)
                End If
            End If
        Next i
        
        If mobjGNComp.CountPCKardexCHP > 0 Then
            ix = mobjGNCompAux.AddPCKardexCHP
            mobjGNCompAux.PCKardexCHP(ix).CodProvCli = mobjGNComp.PCKardexCHP(1).CodProvCli
            mobjGNCompAux.PCKardexCHP(ix).Debe = Round(Abs(mobjGNCompAux.IVKardex(1).CostoTotal), 2)
            mobjGNCompAux.PCKardexCHP(ix).idAsignado = mobjGNComp.PCKardexCHP(1).id
            mobjGNCompAux.PCKardexCHP(ix).codforma = mobjGNComp.PCKardexCHP(1).codforma
        End If
        
        
        Set item = Nothing
        
        
        
        mobjGNCompAux.FechaTrans = mobjGNComp.FechaTrans
        mobjGNCompAux.HoraTrans = DateAdd("s", 1, mobjGNComp.HoraTrans)
        mobjGNCompAux.CodProveedorRef = mobjGNComp.CodProveedorRef
        mobjGNCompAux.nombre = mobjGNComp.nombre
        
        Cadena = "Por Desperdicio en el ingreso " & mobjGNComp.CodTrans & "-" & mobjGNComp.numtrans & ", porcentaje: " & mobjGNComp.IVKardex(1).DescuentoOriginal * 100
        If Len(Cadena) > 120 Then
            mobjGNCompAux.Descripcion = Mid$(Cadena, 1, 120)
        Else
            mobjGNCompAux.Descripcion = Cadena
        End If
            
        mobjGNCompAux.codUsuario = mobjGNComp.codUsuario
        mobjGNCompAux.IdResponsable = mobjGNComp.IdResponsable
        mobjGNCompAux.numDocRef = mobjGNComp.CodTrans & " " & mobjGNComp.numtrans
        mobjGNCompAux.idCentro = mobjGNComp.idCentro
        mobjGNCompAux.IdTransFuente = mobjGNComp.Empresa.RecuperarTransIDGncomprobante(mobjGNComp.CodTrans, mobjGNComp.numtrans)
        mobjGNCompAux.CodMoneda = mobjGNComp.CodMoneda
        'mobjGNCompAux.CodVendedor = fcbVendedor.KeyText
    
        'Si es que algo está modificado
        If mobjGNCompAux.Modificado Then
            MensajeStatus MSG_GENERANDOASIENTO, vbHourglass
            MensajeStatus
        End If
        If mobjGNCompAux.GNTrans.AfectaSaldoPC And _
           mobjGNCompAux.GNTrans.TSVerificaTotalCuadrado Then
            'Verifica si está cuadrado el total de transacción y total de PCKardexCHP.
            If Not TotalCuadrado Then Exit Function
        End If
        'Verificación de datos
        mobjGNCompAux.VerificaDatos
    
        PreparaAsientoBajaAuto True
        'Verifica si está cuadrado el asiento
        If Not VerificaAsiento(mobjGNCompAux) Then Exit Function
    
        'Verifica si tiene detalle de banco
        If (mobjGNCompAux.CountIVKardex = 0) Then
'            MsgBox "No existe ningún detalle.", vbInformation
            Exit Function
        End If

        MensajeStatus MSG_GRABANDO, vbHourglass
    
        'Manda a grabar
        '       Aquí ya no hacemos verificación de asiento por que ya está hecho en Control Asiento
        mobjGNCompAux.Grabar False, False

        '***  Oliver 26/12/2002
        'Agregado para el control ded Impresion Configurado en la Transaccion
        

        MensajeStatus
        GrabarTransDesperdicioAuto = True
    Else
        GrabarTransDesperdicioAuto = False
    End If
    Exit Function
ErrTrap:
    MensajeStatus
    Select Case Err.Number
    Case ERR_DESCUADRADO, ERR_INTEGRIDAD
        'Si es que el usuario seleccionó 'No' en el cuadro de dialogo,
        'No hace nada
    Case Else
        DispErr
        GrabarTransDesperdicioAuto = False
    End Select

    Exit Function
End Function

Private Sub PreparaAsientoBajaAuto(Aceptar As Boolean)
    If mobjGNCompAux.SoloVer Then Exit Sub

    mobjGNCompAux.GeneraAsiento
End Sub

Private Sub CargarRecProv()
Dim ix As Long
Dim t As Currency
Dim orden1 As Integer
Dim obser As String
Dim ts As TSFormaCobroPago
'            Recargos.Refresh
            For ix = mobjGNComp.CountPCKardexCHP To 1 Step -1
                mobjGNComp.RemovePCKardexCHP (ix)
            Next
             t = mobjGNComp.IVKardexTotal(True)
             t = MiCCur(Format$(t, mobjGNComp.FormatoMoneda))  'Redondea al formato de moneda
                t = t + mobjGNComp.IVRecargoTotal(True, False) * Sgn(t)
                'GENERA NUEVA DEUDA
                ix = mobjGNComp.AddPCKardexCHP
                mobjGNComp.PCKardexCHP(ix).Haber = t
                mobjGNComp.PCKardexCHP(ix).CodProvCli = mobjGNComp.CodProveedorRef
                
                Set ts = mobjGNComp.Empresa.RecuperaTSFormaCobroPago(mobjGNComp.GNTrans.CodFormaPre)
                mobjGNComp.PCKardexCHP(ix).codforma = mobjGNComp.GNTrans.CodFormaPre
                
                mobjGNComp.PCKardexCHP(ix).NumLetra = mobjGNComp.CodTrans & " " & mobjGNComp.GNTrans.NumTransSiguiente
                mobjGNComp.PCKardexCHP(ix).FechaEmision = mobjGNComp.FechaTrans
                If Not ts Is Nothing Then
                    mobjGNComp.PCKardexCHP(ix).FechaVenci = mobjGNComp.PCKardexCHP(ix).FechaEmision + ts.Plazo
                Else
                    mobjGNComp.PCKardexCHP(ix).FechaVenci = mobjGNComp.PCKardexCHP(ix).FechaEmision + 30
                End If
                
                obser = "Por pago con: " & mobjGNComp.GNTrans.CodFormaPre & " de " & mobjGNComp.CodTrans & "-" & mobjGNComp.numtrans & " prov: " & mobjGNComp.CodProveedorRef & " - " & mobjGNComp.nombre
                mobjGNComp.PCKardexCHP(ix).Observacion = IIf(Len(obser) > 80, Left(obser, 80), obser)
                'mobjGNComp.pckardexchp(ix).CodVendedor = mobjGNComp.CodVendedor
                mobjGNComp.PCKardexCHP(ix).orden = orden1
                orden1 = orden1 + 1
                Set ts = Nothing
End Sub







Private Function TotalCuadrado() As Boolean
    Dim t As Currency, p As Currency
    With mobjGNComp
        t = .IVKardexTotal(True)
        t = MiCCur(Format$(t, .FormatoMoneda))  'Redondea al formato de moneda
        t = t + .IVRecargoTotal(True, False) * Sgn(t)
        p = .PCKardexCHPHaberTotal - .PCKardexCHPDebeTotal + .TSKardexHaberTotal - .TSKardexDebeTotal
        
        If t <> p Then
            MsgBox "El valor total de transacción (" & Format(t, "#,0.0000") & _
                   ") y forma de pago/cobro (" & Format(p, "#,0.0000") & _
                   ") no están cuadrados por la diferencia de " & _
                        Format(t - p, "#,0.0000") & " " & _
                        mobjGNComp.CodMoneda & "." & vbCr & vbCr & _
                   "Para grabar la transacción tiene que estar cuadrado.", vbInformation
            TotalCuadrado = False
        Else
            TotalCuadrado = True
        End If
    End With
End Function


