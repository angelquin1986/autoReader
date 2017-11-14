VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRegeneraTransformacion 
   Caption         =   "Regeneración de Recetas"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   11775
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   852
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   11775
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3855
      Width           =   11775
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
      Height          =   7935
      Left            =   120
      TabIndex        =   8
      Top             =   1980
      Width           =   6735
      _cx             =   11880
      _cy             =   13996
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
         Format          =   109838337
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
         Format          =   109838337
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
      Height          =   4995
      Left            =   6960
      TabIndex        =   14
      Top             =   1980
      Width           =   8055
      _cx             =   14208
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
   Begin VSFlex7LCtl.VSFlexGrid grdRecetaIng 
      Height          =   4995
      Left            =   6960
      TabIndex        =   18
      Top             =   7320
      Width           =   8055
      _cx             =   14208
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
      Left            =   6960
      TabIndex        =   19
      Top             =   7020
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
Attribute VB_Name = "frmRegeneraTransformacion"
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

Public Sub Inicio()
    Dim i As Integer
    On Error GoTo errtrap
    Me.Show
    Me.ZOrder
    dtpFecha1.value = gobjMain.EmpresaActual.GNOpcion.FechaLimiteDesde
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
    'Carga la lista de transacción
'    fcbTrans.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "IV")

    lstTrans.Clear
    v = gobjMain.GrupoActual.PermisoActual.ListaTrans(False, "IV")
    For i = LBound(v, 2) To UBound(v, 2)
        lstTrans.AddItem v(0, i)        '& " " & v(1, i)
    Next i
    
    If tag = "NotaEntrega" Then
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaNotaEntrega")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaNotaEntrega")
            RecuperaTrans "KeyT", lstTrans, s
        End If
    ElseIf tag = "Desperdicio" Then
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaDesperdicio")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaDesperdicio")
            RecuperaTrans "KeyT", lstTrans, s
        End If
    
    ElseIf tag = "FacturaISO" Then
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaFacturaISO")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaFacturaISO")
            RecuperaTrans "KeyT", lstTrans, s
        End If
    Else
        'jeaa 25/09/206
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaTransformacion")) > 0 Then
            s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("TransparaTransformacion")
            RecuperaTrans "KeyT", lstTrans, s
        End If
    End If
End Sub

Private Sub cmdAceptar_Click()
     'Si no hay transacciones
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna transacción para procesar."
        Exit Sub
    End If
    
    If dtpFecha1 < gobjMain.EmpresaActual.GNOpcion.FechaLimiteDesde Then
        MsgBox "La Rango de Fecha de regeneración es menor a la Fecha Limite Aceptable  ", vbExclamation
        Exit Sub
    End If
    
'    If RegenerarReceta(False, True) Then
'        cmdAceptar.Enabled = True
'        cmdAceptar.SetFocus
'        mVerificado = True
'    End If
    If grdReceta.Rows = 1 Then MsgBox "No hay nada que procesar": Exit Sub
    CorrigeItemsFaltantes
End Sub

Private Function RegenerarAsientoSub(ByVal gnc As GNComprobante, _
                                     ByRef cambiado As Boolean) As Boolean
    Dim i As Long, cta As CtCuenta, ctd As CTLibroDetalle
    Dim colCtd As Collection, a As clsAsiento
    On Error GoTo errtrap
    
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
errtrap:
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
    If col.count <> gnc.CountCTLibroDetalle Then Exit Function
    
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
    On Error GoTo errtrap
'    If Len(fcbGrupoDesde.KeyText) = 0 Then
'        MsgBox "deberia escoger el grupo para aplicar"
'        Exit Sub
'    End If
    grd.Rows = 1
    grdReceta.Rows = 1
    grdRecetaIng.Rows = 1
    With gobjMain.objCondicion
        .fecha1 = dtpFecha1.value
        .fecha2 = dtpFecha2.value
        .CodTrans = PreparaCodTrans
        .NumTrans1 = Val(txtNumTrans1.Text)
        .NumTrans2 = Val(txtNumTrans2.Text)
        'Estados no incluye anulados
        s = PreparaTransParaGnopcion(.CodTrans)
        If Me.tag = "NotaEntrega" Then
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransparaNotaEntrega", s
        ElseIf Me.tag = "Desperdicio" Then
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransparaDesperdicio", s
        ElseIf Me.tag = "FacturaISO" Then
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransparaFacturaISO", s
        Else
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "TransparaTransformacion", s
        End If
        gobjMain.EmpresaActual.GNOpcion.GrabarSoloGnOpcion2
        
    End With
    If Me.tag <> "Contratos" Then
        Set obj = gobjMain.EmpresaActual.ConsGNTransTransformacion(True)
    Else
        Set obj = gobjMain.EmpresaActual.Empresa2.RecuperarContratosErrados
    End If
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
errtrap:
    DispErr
    Exit Sub
End Sub

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


Private Sub Ordenar()
Dim idesde As Long, ihasta As Long
Dim jdesde As Long, jhastas As Long
Dim i As Long, j As Long
For i = 1 To grdReceta.Rows - 1
    ihasta = indDesde(grdReceta.ValueMatrix(i, 1), i, 1)
    OrdenaBurbuja i, ihasta
    i = ihasta
Next
        
For j = 1 To grdReceta.Rows - 1
    jhastas = indDesde(grdReceta.ValueMatrix(j, 8), j, 8)
    OrdenaBurbujaPrecio j, jhastas
    j = jhastas
Next
    
End Sub
Private Function indDesde(TransID As Currency, ind As Long, col As Long) As Currency
Dim i As Long, cont As Long
cont = 0
Dim paso As Long
For i = ind To grdReceta.Rows - 1
    If TransID = grdReceta.ValueMatrix(i, col) Then
        indDesde = i
        cont = cont + 1
    End If
    If cont = 1 Then
        If Not IDSiguienteIgual_Burbuja(i, grdReceta.ValueMatrix(i, col), col) Then
            indDesde = i
            Exit Function
        End If
    End If
    If cont > 1 Then
        If Not IDAnteriorIgual_Burbuja(i, grdReceta.ValueMatrix(i, col), col) Then
                indDesde = i - 1
                Exit Function
            End If
    End If
Next
End Function

Private Sub cmdVerificar_Click()
    grdReceta.Rows = 1
    'Si no hay transacciones
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna transacción para verificar."
        Exit Sub
    End If
    If Me.tag <> "FacturaISO" And Me.tag <> "Contratos" Then
        If dtpFecha1 < gobjMain.EmpresaActual.GNOpcion.FechaLimiteDesde Then
            MsgBox "La Rango de Fecha de regeneración es menor a la Fecha Limite Aceptable  ", vbExclamation
            Exit Sub
        End If
    End If
    If Me.tag = "NotaEntrega" Then
        
        If RegenerarNotaEntrega(True, False) Then
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
            mVerificado = True
        End If
    ElseIf Me.tag = "Desperdicio" Then
        
        If RegenerarDesperdicio(True, False) Then
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
            mVerificado = True
        End If
        
    ElseIf Me.tag = "FacturaISO" Then
        
        If RegenerarFacturaISO(True, False) Then
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
            mVerificado = True
        End If
        
    ElseIf Me.tag = "Contratos" Then
        If RegenerarContratos(True, False) Then
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
            mVerificado = True
        End If
    
    Else
        If RegenerarReceta(True, False) Then
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
            mVerificado = True
        End If
    End If
    
    
End Sub



Private Sub Command1_Click()
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
    grd.Move 0, grd.Top, Me.ScaleWidth / 2, (Me.ScaleHeight - grd.Top - pic1.Height - 80)
    grdReceta.Move grd.Width, grd.Top, Me.ScaleWidth / 2, (Me.ScaleHeight - grd.Top - pic1.Height - 80) * 0.5
    grdRecetaIng.Move grd.Width, grdReceta.Top + grdReceta.Height + 250, Me.ScaleWidth / 2, (Me.ScaleHeight - grd.Top - pic1.Height - 80) * 0.5
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

Private Function RegenerarReceta(bandVerificar As Boolean, BandTodo As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, x As Single
    Dim gnc As GNComprobante, cambiado As Boolean
    On Error GoTo errtrap

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
                    cargaItemsKardex gnc, i, grd.TextMatrix(i, 14)
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
    RegenerarReceta = Not mCancelado
    GoTo salida
errtrap:
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

Private Sub cargaItemsKardex(ByRef gnc As GNComprobante, ByVal i As Long, Optional coditem As String)
    Dim j As Long, rs1 As Recordset
    Dim item As IVinventario
    Dim mbooBand As Boolean, idInve As Long, codInve As String, Desc As String, Tipo As Integer
    Dim x As Long, s As String
    'carga la el detalle transaccion
    
    For j = 1 To grdReceta.Rows - 1
        grdReceta.RemoveItem 1
    Next j
    
    For j = 1 To grdRecetaIng.Rows - 1
        grdRecetaIng.RemoveItem 1
    Next j
    
    
    For j = 1 To gnc.CountIVKardex
        
        Set item = gnc.Empresa.RecuperaIVInventarioQuick(gnc.IVKardex(j).CodInventario)
        s = gnc.IVKardex(j).Orden & vbTab & gnc.IVKardex(j).id & vbTab & gnc.TransID & vbTab & gnc.IVKardex(j).idinventario & vbTab & gnc.IVKardex(j).CodInventario & vbTab & item.Descripcion & vbTab & gnc.IVKardex(j).cantidad & vbTab & gnc.IVKardex(j).CostoTotal
        
        
        If gnc.IVKardex(j).cantidad < 1 Then
            If Len(s) > 0 Then          '*** MAKOTO 09/nov/00 para no agregar items de destino en Trans. Bodegas
                With grdReceta
                    .AddItem s
                End With
            End If
        Else
            If Len(s) > 0 Then          '*** MAKOTO 09/nov/00 para no agregar items de destino en Trans. Bodegas
                With grdRecetaIng
                    .AddItem s
                End With
            End If
            
        End If
    Next j
    Set item = Nothing
    'grdReceta.subtotal flexSTSum, 2, 7, , , , , , , True
    If grdRecetaIng.Rows > 2 Then
        grd.TextMatrix(i, COL_RESULTADO) = "Error"
    Else
        grd.TextMatrix(i, COL_RESULTADO) = "OK"
    End If
End Sub

Private Sub RecursivoReceta(ByVal gnc As GNComprobante, ByVal i As Long, ByVal coditem As String, ByVal cant As Currency)
Dim mbooBand As Boolean, idInve As Long, codInve As String, Desc As String, Tipo As Integer
Dim id As Long, IdPadre As Long, cantR As Currency
Dim x As Long, rs1 As Recordset
Dim IdBodega As Long, sql As String
Dim rs As Recordset

With grdReceta
    Set rs1 = gnc.Empresa.rsReceta(coditem)
    If rs1.RecordCount > 0 Then
    Do While Not rs1.EOF
        mbooBand = False
        For x = i To gnc.CountIVKardex
            cantR = rs1!cantidad
            idInve = rs1!idmateria
            codInve = rs1!CodInventario
            Desc = rs1!Descripcion
            Tipo = rs1!Tipo
            sql = "select idbodega from ivbodega where codbodega = '" & gnc.IVKardex(i).CodBodega & " '"
            Set rs = gnc.Empresa.OpenRecordset(sql)
            If rs.RecordCount > 0 Then IdBodega = rs!IdBodega
            id = gnc.TransID
            IdPadre = gnc.IVKardex(x).IdPadre
            If rs1!idmateria = gnc.IVKardex(x).idinventario Then
                mbooBand = True
            End If
            Set rs = Nothing
        Next x
        If mbooBand = False Then
            .AddItem .Rows + 1 & vbTab & id & vbTab & gnc.CodTrans & " " & gnc.numtrans & vbTab & idInve & vbTab & codInve & vbTab & Desc & vbTab & cant * cantR & vbTab & IdBodega & vbTab & IdPadreT
        End If
        
        rs1.MoveNext
    Loop
    End If
End With
End Sub

Private Function CorrigeItemsFaltantes()  'ByVal idInven As Long, ByVal cant As Currency, ByVal idPadres As Long, gnc As GNComprobante) As Boolean
    Dim i As Long, sql As String, rs As Recordset, fila As Integer
    Dim j As Long, idsubItem As Long, CantSubItem As Currency, costo As Currency
    Dim item As IVinventario, tid As Long, cant As Currency
    Dim gnc As GNComprobante, x As Long
    Dim c As Currency, Orden As Integer
    Dim IdBod As Long, BODAUX As String
    Dim ivk As IVKardex, gnt As GNTrans, ix As Long
    On Error GoTo errtrap
    
    
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    
    For i = 1 To grd.Rows - 1
     DoEvents
     prg1.value = i
     grd.Row = i
     x = grd.CellTop
        
        If grd.TextMatrix(i, 11) = "Error" Then
'            tid = grdReceta.ValueMatrix(i, 1)
'            cant = grdReceta.ValueMatrix(i, 6) * -1
            grdReceta.Refresh
            'Recupera la transaccion
            
            Set gnt = gobjMain.EmpresaActual.RecuperaGNTrans(grdReceta.TextMatrix(i, 9))
            Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(grdReceta.TextMatrix(i, 8))
            fila = gnc.CountIVKardex
            For j = 1 To fila
                ix = gnc.AddIVKardex
                Set ivk = gnc.IVKardex(j)
                gnc.IVKardex(ix).CodBodega = gnt.CodBodegaDesPre
                gnc.IVKardex(ix).CodInventario = ivk.CodInventario
                gnc.IVKardex(ix).cantidad = ivk.cantidad * -1
                gnc.IVKardex(ix).CostoRealTotal = ivk.CostoRealTotal * -1
                gnc.IVKardex(ix).CostoTotal = ivk.CostoTotal * -1
                gnc.IVKardex(ix).Orden = x
                gnc.IVKardex(ix).PrecioRealTotal = ivk.PrecioRealTotal
                gnc.IVKardex(ix).PrecioTotal = ivk.PrecioTotal
                gnc.IVKardex(ix).IVA = ivk.IVA
                gnc.IVKardex(ix).NumeroPrecio = ivk.NumeroPrecio
                gnc.IVKardex(ix).Descuento = ivk.Descuento
                gnc.IVKardex(ix).FechaLleva = ivk.FechaLleva
            Next j
'        Next i
            
        gnc.Grabar False, False
                    grd.TextMatrix(i, 11) = "Corregido  "
            x = grd.CellTop                 'Para visualizar la celda actual

        
    End If
Next i
        Screen.MousePointer = 0
        CorrigeItemsFaltantes = True
    GoTo salida
errtrap:
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

Private Function RegenerarRecetaCant(bandVerificar As Boolean, BandTodo As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, x As Single
    Dim gnc As GNComprobante, cambiado As Boolean
    On Error GoTo errtrap

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
                    cargaItemsKardexTodo gnc, i, grd.TextMatrix(i, 14)
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
    RegenerarRecetaCant = Not mCancelado
    GoTo salida
errtrap:
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

Private Sub cargaItemsKardexTodo(ByRef gnc As GNComprobante, ByVal i As Long, Optional coditem As String)
    Dim j As Long, rs1 As Recordset
    Dim item As IVinventario, sql As String
    Dim mbooBand As Boolean, idInve As Long, codInve As String, Desc As String, Tipo As Integer
    Dim x As Long, id As Long, rs As Recordset, IdBodega As Long, idpadresub As Long
    'carga la el detalle transaccion
    id = gnc.TransID
    For j = 1 To gnc.CountIVKardex

            Set item = gnc.Empresa.RecuperaIVInventario(gnc.IVKardex(j).CodInventario)
            
            sql = "select idbodega from ivbodega where codbodega = '" & gnc.IVKardex(j).CodBodega & " '"
            Set rs = gnc.Empresa.OpenRecordset(sql)
            If rs.RecordCount > 0 Then IdBodega = rs!IdBodega
            ''''''''''-----------


'            idpadresub = recuperaidSub(gnc, gnc.IVKardex(j).CodInventario, gnc.IVKardex(j).idpadre)
            grdReceta.AddItem grdReceta.Rows & vbTab & id & vbTab & gnc.CodTrans & " " & gnc.numtrans & vbTab & item.RecuperaId(gnc.IVKardex(j).CodInventario) & vbTab & item.CodInventario & vbTab & item.Descripcion & vbTab & Abs(gnc.IVKardex(j).cantidad) & vbTab & IdBodega & vbTab & gnc.IVKardex(j).IdPadre & vbTab & vbTab & vbTab & gnc.IVKardex(j).id & vbTab & Abs(gnc.IVKardex(j).PrecioTotal)
                                
    Next j
    Set rs = Nothing
    Set item = Nothing
End Sub

Private Function recuperaidSub(ByVal gnc As GNComprobante, ByVal coditem As String, ByVal IdPadre As Long) As Long
    Dim sql As String, rs As Recordset
    sql = "Select * from ivmateria ivm inner join ivinventario iv on iv.idinventario = ivm.idmateria"
    sql = sql & " where iv.codinventario = '" & coditem & "'"
    sql = sql & " AND ivm.idinventario = " & IdPadre
    Set rs = gnc.Empresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then recuperaidSub = rs!idinventario
End Function
Private Function RegenerarCantReal(bandVerificar As Boolean, BandTodo As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, x As Single, ihasta As Long
    Dim gnc As GNComprobante, cambiado As Boolean, NOesPreparacion As Boolean, ihastaId As Long
    On Error GoTo errtrap

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
    prg1.max = grdReceta.Rows - 1
    
    For i = grdReceta.FixedRows To grdReceta.Rows - 1
        DoEvents
        If mCancelado Then
            MsgBox "El proceso fue cancelado.", vbInformation
            Exit For
        End If
        
        prg1.value = i
        grdReceta.Row = i
        x = grdReceta.CellTop                 'Para visualizar la celda actual
        tid = grdReceta.TextMatrix(i, 1)
        Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
        'Si es verificación, procesa todas las filas sino solo las que tengan "Asiento incorrecto."
        
        If Len((grdReceta.TextMatrix(i, 9))) = 0 Then 'Or bandVerificar Or bandTodo Then
            ihasta = hastaIndice(grdReceta.ValueMatrix(i, 8), i)
            RevisarCantidad gnc, grdReceta.ValueMatrix(i, 6), grdReceta.ValueMatrix(i, 8), ihasta, NOesPreparacion, grdReceta.ValueMatrix(i, grdReceta.Cols - 2)
            If NOesPreparacion Then
                grdReceta.TextMatrix(i, 10) = grdReceta.ValueMatrix(i, 6)
                grdReceta.TextMatrix(i, 9) = "Verificando..."
            End If
        End If
'        If i = 300 Then
'                GoTo salida
'        End If
        Set gnc = Nothing
    Next i
    Screen.MousePointer = 0
    RegenerarCantReal = Not mCancelado
    GoTo salida
errtrap:
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

Private Function seRepite(iditem As Long, ind As Long, Optional j As Long) As Boolean
Dim i As Long
     For i = j To ind
        If grdReceta.ValueMatrix(i, 3) = iditem And grdReceta.TextMatrix(i, grdReceta.Cols - 1) = "ok" Then
            seRepite = True
            Exit Function
        End If
     Next
End Function

Private Function hastaIndice(ByVal IdPadre As Long, ByVal ind As Long) As Long
Dim i As Long, cont As Long
cont = 0
Dim paso As Long
For i = ind To grdReceta.Rows - 1
    If IdPadre = grdReceta.ValueMatrix(i, 8) And Len(grdReceta.TextMatrix(i, 9)) = 0 Then
        hastaIndice = i
        cont = cont + 1
    End If
    If cont = 1 Then
        If Not SiguienteIgual(i, grdReceta.ValueMatrix(i, 8)) Then
            hastaIndice = i
            Exit Function
        End If
    End If
    If cont > 1 Then
        If Not AnteriorIgual(i, grdReceta.ValueMatrix(i, 8)) Then
                hastaIndice = i - 1
                Exit Function
            End If
    End If
    
Next
End Function
Private Function SiguienteIgual(i As Long, valor As Long) As Boolean
    If i = grdReceta.Rows - 1 Then Exit Function
    If grdReceta.ValueMatrix(i + 1, 8) = valor Then
        SiguienteIgual = True
    Else
        SiguienteIgual = False
    End If

End Function
Private Function AnteriorIgual(i As Long, valor As Long) As Boolean
    If grdReceta.ValueMatrix(i - 1, 8) = valor Then
        AnteriorIgual = True
    Else
        AnteriorIgual = False
    End If
End Function
Private Sub RevisarCantidad(ByVal gnc As GNComprobante, ByVal cant As Currency, ByVal IdPadre As Long, ihasta As Long, ByRef NOesPreparacion As Boolean, ByVal pv As Currency)
Dim sql As String, rs As Recordset
Dim i As Long
sql = "select iv1.tipo,ivm.* from ivmateria ivm inner join ivinventario iv on iv.idinventario = ivm.idinventario inner join ivinventario iv1 on iv1.idinventario = ivm.idmateria Where ivm.idinventario = " & IdPadre
Set rs = gnc.Empresa.OpenRecordset(sql)
Do While Not rs.EOF
    For i = 1 To ihasta
        If grdReceta.TextMatrix(i, 3) = rs!idmateria And Len(grdReceta.TextMatrix(i, 9)) = 0 Then
          '  If grdReceta.ValueMatrix(i, grdReceta.Cols - 2) <> 0 Then  'Compara para ver si tiene precio
                grdReceta.TextMatrix(i, 10) = rs!cantidad * cant    'si tiene precio es padre
           ' Else
            '    grdReceta.TextMatrix(i, 10) = rs!cantidad
            'End If
                grdReceta.TextMatrix(i, 9) = "Verificando..."
                grdReceta.Refresh
                If rs!Tipo = 4 Then
             '       grdReceta.TextMatrix(i, 10) = rs!cantidad * cant
                    IdPadre = rs!idmateria
                    RevisarCantidad gnc, cant, IdPadre, ihasta, NOesPreparacion, pv
                End If
        End If
    Next
        rs.MoveNext
Loop
If rs.RecordCount = 0 Then
    NOesPreparacion = True
End If
End Sub

Private Function IDhastaIndice(ByVal TransID As Long, ByVal ind As Long) As Long
Dim i As Long, cont As Long
cont = 0
Dim paso As Long
For i = ind To grdReceta.Rows - 1
    If TransID = grdReceta.ValueMatrix(i, 1) Then
        IDhastaIndice = i
        cont = cont + 1
    End If
    If cont = 1 Then
        If Not IDSiguienteIgual(i, grdReceta.ValueMatrix(i, 1)) Then
            IDhastaIndice = i
            Exit Function
        End If
    End If
    If cont > 1 Then
        If Not IDAnteriorIgual(i, grdReceta.ValueMatrix(i, 1)) Then
                IDhastaIndice = i - 1
                Exit Function
            End If
    End If
Next
End Function


Private Function IDSiguienteIgual(i As Long, valor As Long) As Boolean
    If i = grdReceta.Rows - 1 Then Exit Function
    If grdReceta.ValueMatrix(i + 1, 1) = valor Then
        IDSiguienteIgual = True
    Else
        IDSiguienteIgual = False
    End If

End Function
Private Function IDAnteriorIgual(i As Long, valor As Long) As Boolean
    If grdReceta.ValueMatrix(i - 1, 1) = valor Then
        IDAnteriorIgual = True
    Else
        IDAnteriorIgual = False
    End If
End Function



Public Sub OrdenaBurbuja(fila As Long, col As Long)   'Procedimiento que utiliza el metodo
Dim i As Integer, j As Integer 'de la burbuja para ordenar
Dim tamaño As Integer
With grdReceta
tamaño = col
    For i = fila To tamaño - 1
        For j = fila To tamaño - 1
             If Val(.TextMatrix(j + 1, 8)) >= Val(.TextMatrix(j, 8)) Then
                Call cambio(j, j + 1)
            End If
        Next
    
    Next
End With
End Sub
Public Sub OrdenaBurbujaPrecio(fila As Long, col As Long)   'Procedimiento que utiliza el metodo
Dim i As Integer, j As Integer 'de la burbuja para ordenar
Dim tamaño As Integer
With grdReceta
tamaño = col
    For i = fila To tamaño - 1
        For j = fila To tamaño - 1
             If Val(.TextMatrix(j + 1, 12)) >= Val(.TextMatrix(j, 12)) Then
                Call cambio(j, j + 1)
            End If
        Next
    
    Next
End With
End Sub

'AUC 11/07/07
Private Sub cambio(ByVal a As Integer, ByVal b As Integer)
Dim col1 As String
Dim col2 As String
Dim col3 As String
Dim col4 As String
Dim col5 As String
Dim col6 As String
Dim col7 As String
Dim col8 As String
Dim col9 As String
Dim col10 As String
Dim col11 As String
Dim col12 As String
Dim col13 As String

    With grdReceta
       col1 = .TextMatrix(a, 1)
       col2 = .TextMatrix(a, 2)
       col3 = .TextMatrix(a, 3)
       col4 = .TextMatrix(a, 4)
       col5 = .TextMatrix(a, 5)
       col6 = .TextMatrix(a, 6)
       col7 = .TextMatrix(a, 7)
       col8 = .TextMatrix(a, 8)
       col9 = .TextMatrix(a, 9)
       col10 = .TextMatrix(a, 10)
       col11 = .TextMatrix(a, 11)
       col12 = .TextMatrix(a, 12)
       col13 = .TextMatrix(a, 13)
      .TextMatrix(a, 1) = .TextMatrix(b, 1)
      .TextMatrix(a, 2) = .TextMatrix(b, 2)
      .TextMatrix(a, 3) = .TextMatrix(b, 3)
      .TextMatrix(a, 4) = .TextMatrix(b, 4)
      .TextMatrix(a, 5) = .TextMatrix(b, 5)
      .TextMatrix(a, 6) = .TextMatrix(b, 6)
      .TextMatrix(a, 7) = .TextMatrix(b, 7)
      .TextMatrix(a, 8) = .TextMatrix(b, 8)
      .TextMatrix(a, 9) = .TextMatrix(b, 9)
      .TextMatrix(a, 10) = .TextMatrix(b, 10)
      .TextMatrix(a, 11) = .TextMatrix(b, 11)
      .TextMatrix(a, 12) = .TextMatrix(b, 12)
      .TextMatrix(a, 13) = .TextMatrix(b, 13)
      
      .TextMatrix(b, 1) = col1
      .TextMatrix(b, 2) = col2
      .TextMatrix(b, 3) = col3
      .TextMatrix(b, 4) = col4
      .TextMatrix(b, 5) = col5
      .TextMatrix(b, 6) = col6
      .TextMatrix(b, 7) = col7
      .TextMatrix(b, 8) = col8
      .TextMatrix(b, 9) = col9
      .TextMatrix(b, 10) = col10
      .TextMatrix(b, 11) = col11
      .TextMatrix(b, 12) = col12
      .TextMatrix(b, 13) = col13
      
    End With
End Sub
















Private Function IDSiguienteIgual_Burbuja(i As Long, valor As Long, col As Long) As Boolean
    If i = grdReceta.Rows - 1 Then Exit Function
    If grdReceta.ValueMatrix(i + 1, col) = valor Then
        IDSiguienteIgual_Burbuja = True
    Else
        IDSiguienteIgual_Burbuja = False
    End If

End Function
Private Function IDAnteriorIgual_Burbuja(i As Long, valor As Long, col As Long) As Boolean
    If grdReceta.ValueMatrix(i - 1, col) = valor Then
        IDAnteriorIgual_Burbuja = True
    Else
        IDAnteriorIgual_Burbuja = False
    End If
End Function


Public Sub InicioNotaEntrega(ByVal tag As String)
    Dim i As Integer
    On Error GoTo errtrap
    Me.tag = tag
    Me.Show
    Me.ZOrder
    dtpFecha1.value = gobjMain.EmpresaActual.GNOpcion.FechaLimiteDesde
    dtpFecha2.value = Date
    CargaTrans
    
    
    Exit Sub
errtrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

''Private Sub cargaItemsKardexNotaEntrega(ByRef gnc As GNComprobante, ByVal i As Long, Optional coditem As String)
''    Dim j As Long, rs1 As Recordset, k As Long
''    Dim item As IVinventario
''    Dim mbooBand As Boolean, idInve As Long, codInve As String, Desc As String, Tipo As Integer
''    Dim x As Long, s As String, sql As String
''    'carga la el detalle transaccion
''
''    For j = 1 To grdReceta.Rows - 1
''        grdReceta.RemoveItem 1
''    Next j
''
''    For j = 1 To grdRecetaIng.Rows - 1
''        grdRecetaIng.RemoveItem 1
''    Next j
''
''
''    For j = 1 To gnc.CountIVKardex
''
''        Set item = gnc.Empresa.RecuperaIVInventarioQuick(gnc.IVKardex(j).CodInventario)
''        s = gnc.IVKardex(j).orden & vbTab & gnc.IVKardex(j).id & vbTab & gnc.numtrans & vbTab & gnc.IVKardex(j).idinventario & vbTab & gnc.IVKardex(j).CodInventario & vbTab & item.Descripcion & vbTab & gnc.IVKardex(j).cantidad & vbTab & gnc.TransID
''
''
''        If gnc.IVKardex(j).cantidad > 0 Then
''            If Len(s) > 0 Then          '*** MAKOTO 09/nov/00 para no agregar items de destino en Trans. Bodegas
''                With grdReceta
''                    .AddItem s
''                End With
''            End If
''        End If
''    Next j
''
''
''    For k = 1 To grdReceta.Rows - 1
''        sql = "select  codtrans, numtrans from gncomprobante g inner join ivkardex i on g.transid = i.transid"
''        sql = sql & " where g.estado <> 3 and left(codtrans,2)= 'FC'  and i.numdias=" & grdReceta.ValueMatrix(k, 7)
''        sql = sql & " and idinventario=" & grdReceta.ValueMatrix(k, 3)
''        Set rs1 = gobjMain.EmpresaActual.OpenRecordset(sql)
''        If rs1.RecordCount > 0 Then
''            grdReceta.TextMatrix(k, 9) = "O.K., " & rs1.Fields("codtrans") & "-" & rs1.Fields("numtrans")
''            grd.TextMatrix(i, COL_RESULTADO) = "O.K. " & rs1.Fields("codtrans") & "-" & rs1.Fields("numtrans")
''        Else
''
'''            sql = "spConsIVSaldoNotasEntrega " & grdReceta.ValueMatrix(k, 7) & ",'NE','NEM'"
''            Set rs1 = gobjMain.EmpresaActual.ConsIVSaldoNotasEntrega(grdReceta.ValueMatrix(k, 7), "NE")
''            If rs1.RecordCount > 0 Then
''                grdReceta.TextMatrix(k, 9) = "O.K., " & rs1.Fields("trans")
''                grd.TextMatrix(i, COL_RESULTADO) = "O.K. " & rs1.Fields("trans")
''
''            Else
''                grd.TextMatrix(i, COL_RESULTADO) = "Por Facturar"
''            End If
''        End If
''
''    Next k
''
''
''    Set item = Nothing
''    If grdRecetaIng.Rows > 2 Then
'''        MsgBox gnc.numtrans
''        grd.TextMatrix(i, COL_RESULTADO) = "Error"
''    End If
''End Sub

Private Function RegenerarNotaEntrega(bandVerificar As Boolean, BandTodo As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, x As Single
    Dim gnc As GNComprobante, cambiado As Boolean
    On Error GoTo errtrap

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
                    cargaItemsKardexNotaEntrega gnc, i, grd.TextMatrix(i, 14)
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
    RegenerarNotaEntrega = Not mCancelado
    GoTo salida
errtrap:
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

Private Sub cargaItemsKardexNotaEntrega(ByRef gnc As GNComprobante, ByVal i As Long, Optional coditem As String)
    Dim j As Long, rs1 As Recordset
    Dim item As IVinventario
    Dim mbooBand As Boolean, idInve As Long, codInve As String, Desc As String, Tipo As Integer
    Dim x As Long, s As String
    'carga la el detalle transaccion
    
    For j = 1 To grdReceta.Rows - 1
        grdReceta.RemoveItem 1
    Next j
    
    For j = 1 To grdRecetaIng.Rows - 1
        grdRecetaIng.RemoveItem 1
    Next j
    
    
    For j = 1 To gnc.CountIVKardex
        
        Set item = gnc.Empresa.RecuperaIVInventarioQuick(gnc.IVKardex(j).CodInventario)
        s = gnc.IVKardex(j).Orden & vbTab & gnc.IVKardex(j).id & vbTab & gnc.numtrans & vbTab & gnc.IVKardex(j).idinventario & vbTab & gnc.IVKardex(j).CodInventario & vbTab & item.Descripcion & vbTab & gnc.IVKardex(j).cantidad & vbTab & gnc.IVKardex(j).CostoTotal & vbTab & gnc.TransID & vbTab & gnc.CodTrans
        
        
        If gnc.IVKardex(j).cantidad < 1 Then
            If Len(s) > 0 Then          '*** MAKOTO 09/nov/00 para no agregar items de destino en Trans. Bodegas
                With grdReceta
                    .AddItem s
                End With
            End If
        Else
            If Len(s) > 0 Then          '*** MAKOTO 09/nov/00 para no agregar items de destino en Trans. Bodegas
                With grdRecetaIng
                    .AddItem s
                End With
            End If
            
        End If
    Next j
    Set item = Nothing
    'grdReceta.subtotal flexSTSum, 2, 7, , , , , , , True
    If grdRecetaIng.Rows <> grdReceta.Rows Then
        'MsgBox gnc.numtrans
        grd.TextMatrix(i, COL_RESULTADO) = "Error"
    Else
        grd.TextMatrix(i, COL_RESULTADO) = "OK"
    End If
End Sub

Public Sub InicioFacturaISO(ByVal tag As String)
    Dim i As Integer
    On Error GoTo errtrap
    Me.tag = tag
    Me.Show
    Me.ZOrder
    dtpFecha1.value = gobjMain.EmpresaActual.GNOpcion.FechaLimiteDesde
    dtpFecha2.value = Date
    CargaTrans
    
    
    Exit Sub
errtrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Function RegenerarFacturaISO(bandVerificar As Boolean, BandTodo As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, x As Single
    Dim gnc As GNComprobante, cambiado As Boolean
    On Error GoTo errtrap

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
                    cargaItemsKardexFacturaISO gnc, i, grd.TextMatrix(i, 14)
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
    RegenerarFacturaISO = Not mCancelado
    GoTo salida
errtrap:
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


Private Sub cargaItemsKardexFacturaISO(ByRef gnc As GNComprobante, ByVal i As Long, Optional coditem As String)
    Dim j As Long, rs1 As Recordset
    Dim item As IVinventario, bandFactura As String
    Dim mbooBand As Boolean, idInve As Long, codInve As String, Desc As String, Tipo As Integer
    Dim x As Long, s As String
    'carga la el detalle transaccion
    
    For j = 1 To grdReceta.Rows - 1
        grdReceta.RemoveItem 1
    Next j
    
    For j = 1 To grdRecetaIng.Rows - 1
        grdRecetaIng.RemoveItem 1
    Next j
    
    
    For j = 1 To gnc.CountIVKardex
        
        'Set item = gnc.Empresa.RecuperaIVInventarioQuick(gnc.IVKardex(j).CodInventario)
        s = gnc.IVKardex(j).Orden & vbTab & gnc.IVKardex(j).id & vbTab & gnc.TransID & vbTab & gnc.IVKardex(j).idinventario & vbTab & gnc.IVKardex(j).CodInventario & vbTab & gnc.IVKardex(j).TiempoEntrega & vbTab & gnc.IVKardex(j).cantidad & vbTab & gnc.IVKardex(j).CostoTotal & vbTab & gnc.TransID & vbTab & gnc.CodTrans
        
        
        If Len(s) > 0 Then          '*** MAKOTO 09/nov/00 para no agregar items de destino en Trans. Bodegas
            With grdReceta
                .AddItem s
            End With
        End If
    Next j
    
    For j = 1 To grdReceta.Rows - 1
        If grdReceta.ValueMatrix(j, 5) <> 0 Then
            bandFactura = IIf(gnc.Empresa.BuscaDatosInventarioISO(grdReceta.ValueMatrix(j, 5), "bandFactura") = "Verdadero", "OK", "Error")
            grdRecetaIng.AddItem vbTab & vbTab & grdReceta.ValueMatrix(j, 5) & vbTab & vbTab & bandFactura
            If bandFactura = "Error" Then
                grd.TextMatrix(i, COL_RESULTADO) = "Error ticket " & grdReceta.ValueMatrix(j, 5)
                gnc.ActualizaTransIdNew grdReceta.ValueMatrix(j, 5), "BandFactura", 1
                gnc.ActualizaTransIdNew grdReceta.ValueMatrix(j, 5), "TransIDFactura", grdReceta.ValueMatrix(j, 3)
                
            Else
                grd.TextMatrix(i, COL_RESULTADO) = "OK"
            End If
        Else
            grd.TextMatrix(i, COL_RESULTADO) = "OK"
        End If
    Next j
    
    
    Set item = Nothing
    'grdReceta.subtotal flexSTSum, 2, 7, , , , , , , True
'''    If grdRecetaIng.Rows <> grdReceta.Rows Then
'''        'MsgBox gnc.numtrans
'''        grd.TextMatrix(i, COL_RESULTADO) = "Error"
'''    Else
'''        grd.TextMatrix(i, COL_RESULTADO) = "OK"
'''    End If
End Sub


Public Sub InicioDesperdicio(ByVal tag As String)
    Dim i As Integer
    On Error GoTo errtrap
    Me.tag = tag
    Me.Show
    Me.ZOrder
    dtpFecha1.value = gobjMain.EmpresaActual.GNOpcion.FechaLimiteDesde
    dtpFecha2.value = Date
    CargaTrans
    
    
    Exit Sub
errtrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Function RegenerarDesperdicio(bandVerificar As Boolean, BandTodo As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, x As Single
    Dim gnc As GNComprobante, cambiado As Boolean
    On Error GoTo errtrap

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
errtrap:
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
    Dim j As Long, rs1 As Recordset
    Dim item As IVinventario
    Dim mbooBand As Boolean, idInve As Long, codInve As String, Desc As String, Tipo As Integer
    Dim x As Long, s As String
    'carga la el detalle transaccion
    
    For j = 1 To grdReceta.Rows - 1
        grdReceta.RemoveItem 1
    Next j
    
    For j = 1 To grdRecetaIng.Rows - 1
        grdRecetaIng.RemoveItem 1
    Next j
    
    
    For j = 1 To gnc.CountIVKardex
        
        Set item = gnc.Empresa.RecuperaIVInventarioQuick(gnc.IVKardex(j).CodInventario)
        s = gnc.IVKardex(j).Orden & vbTab & gnc.IVKardex(j).id & vbTab & gnc.numtrans & vbTab & gnc.IVKardex(j).idinventario & vbTab & gnc.IVKardex(j).CodInventario & vbTab & item.Descripcion & vbTab & gnc.IVKardex(j).cantidad & vbTab & gnc.IVKardex(j).CostoTotal & vbTab & gnc.TransID & vbTab & gnc.CodTrans
        
        
        If gnc.IVKardex(j).cantidad < 1 Then
            If Len(s) > 0 Then          '*** MAKOTO 09/nov/00 para no agregar items de destino en Trans. Bodegas
                With grdReceta
                    .AddItem s
                End With
            End If
        Else
            If Len(s) > 0 Then          '*** MAKOTO 09/nov/00 para no agregar items de destino en Trans. Bodegas
                With grdRecetaIng
                    .AddItem s
                End With
            End If
            
        End If
    Next j
    Set item = Nothing
    'grdReceta.subtotal flexSTSum, 2, 7, , , , , , , True
    If grdRecetaIng.Rows <> grdReceta.Rows Then
        'MsgBox gnc.numtrans
        grd.TextMatrix(i, COL_RESULTADO) = "Error"
    Else
        grd.TextMatrix(i, COL_RESULTADO) = "OK"
    End If
End Sub


Public Sub InicioContratos(ByVal tag As String)
    Dim i As Integer
    On Error GoTo errtrap
    Me.tag = tag
    Me.Show
    Me.ZOrder

    
    
    Exit Sub
errtrap:
    DispErr
    Unload Me
    Exit Sub
End Sub


Private Function RegenerarContratos(bandVerificar As Boolean, BandTodo As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, x As Single
    Dim gnc As GNComprobante, cambiado As Boolean
    
    Dim pc As PCProvCli
    On Error GoTo errtrap

    'Si no es solo verificacion, confirma
    If Not bandVerificar Then
'        s = "Desea volver a Verificar la transacción seleccionada." & vbCr & vbCr
'        s = s & "Está seguro que desea proceder?"
'        If MsgBox(s, vbYesNo + vbQuestion) <> vbYes Then
'            Exit Function
'        Else
'            grdReceta.Rows = 1
'        End If
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
            tid = grd.ValueMatrix(i, 4)
            grd.TextMatrix(i, COL_RESULTADO) = "Verificando..."
            grd.Refresh
            'Recupera la transaccion
            Set pc = gobjMain.EmpresaActual.RecuperaPCProvCli(tid)
            If Not (pc Is Nothing) Then
                'Si la transacción no está anulada
'                If gnc.Estado <> ESTADO_ANULADO Then
 '                   'Forzar recuperar todos los datos de transacción para que no se pierdan al grabar de nuveo
                    'gnc.RecuperaDetalleTodo
                    cargaContratos pc, i, grd.TextMatrix(i, 14)
                    num_fila_trans = i
'                Else
'                    'Si está anulada
'                    grd.TextMatrix(i, COL_RESULTADO) = "Anulado."
'                End If
            Else
                grd.TextMatrix(i, COL_RESULTADO) = "No pudo recuperar la transación."
            End If
        End If
    Next i
    Screen.MousePointer = 0
    RegenerarContratos = Not mCancelado
    GoTo salida
errtrap:
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


Private Sub cargaContratos(ByRef pc As PCProvCli, ByVal i As Long, Optional coditem As String)
    Dim j As Long, rs1 As Recordset, obj As IVDetalleItem, col As Integer, ix As Long
    Dim item As IVinventario, bandFactura As String
    Dim mbooBand As Boolean, idInve As Long, codInve As String, Desc As String, Tipo As Integer
    Dim x As Long, s As String
    'carga la el detalle transaccion
    
    For j = 1 To grdReceta.Rows - 1
        grdReceta.RemoveItem 1
    Next j
    
    For j = 1 To grdRecetaIng.Rows - 1
        grdRecetaIng.RemoveItem 1
    Next j
    
    
    For j = 1 To pc.NumDetalleItem
        
        'Set item = gnc.Empresa.RecuperaIVInventarioQuick(gnc.IVKardex(j).CodInventario)
        's = gnc.IVKardex(j).Orden & vbTab & gnc.IVKardex(j).id & vbTab & gnc.TransID & vbTab & gnc.IVKardex(j).idinventario & vbTab & gnc.IVKardex(j).CodInventario & vbTab & gnc.IVKardex(j).TiempoEntrega & vbTab & gnc.IVKardex(j).cantidad & vbTab & gnc.IVKardex(j).CostoTotal & vbTab & gnc.TransID & vbTab & gnc.CodTrans
        Set obj = pc.RecuperaDetalleItem(j)
        s = vbTab & obj.CodInventario & vbTab & obj.Contrato & vbTab & obj.CodInventario & vbTab & obj.Descripcion & vbTab & obj.Descripcion & vbTab & obj.cantidad & vbTab & obj.PU & vbTab & obj.fechaIni & vbTab & obj.FechaFin & vbTab & obj.Frecuencia & vbTab & obj.referencia & vbTab & IIf(obj.BandProntoPago, 1, 0) & vbTab & IIf(obj.bandPublicidad, 1, 0) & vbTab & obj.Plazo
        
        
        If Len(s) > 0 Then          '*** MAKOTO 09/nov/00 para no agregar items de destino en Trans. Bodegas
            With grdReceta
                .AddItem s, j
            End With
        End If
    Next j
    
'    grdReceta.Sort = flexSortUseColSort
'    grdReceta.ColSort(2) = flexSortGenericAscending
'    grdReceta.Redraw = flexRDBuffered
    
    grdReceta.Select 1, 2, 1, 2
    grdReceta.Sort = flexSortUseColSort

        

    For j = grdReceta.Rows - 1 To 1 Step -1
        If grdReceta.TextMatrix(j, 2) = grdReceta.TextMatrix(j - 1, 2) And grdReceta.TextMatrix(j, 3) = grdReceta.TextMatrix(j - 1, 3) Then
            grdReceta.TextMatrix(j, 7) = "1"
        End If
        If Len(grdReceta.TextMatrix(j, 2)) = 0 Then
            grdReceta.TextMatrix(j, 7) = "1"
        End If
    Next j
    
    s = ""
    For j = 1 To grdReceta.Rows - 1
        s = ""
        If grdReceta.TextMatrix(j, 7) = "0" Then
            For col = 1 To grdReceta.Cols - 1
                s = s & grdReceta.TextMatrix(j, col) & vbTab
            Next col
        If Len(s) > 0 Then          '*** MAKOTO 09/nov/00 para no agregar items de destino en Trans. Bodegas
            With grdRecetaIng
                .AddItem s
            End With
        End If
        End If
    Next j

'    If grdRecetaIng.Rows > 2 Then
'        MsgBox "mas de dos contratos"
'    End If

        For ix = 1 To pc.NumDetalleItem
            pc.RemoveDetalleItem 1
        Next ix
        
        
        For i = 1 To grdRecetaIng.Rows - 1
            ix = pc.AddDetalleItem
            Set obj = pc.RecuperaDetalleItem(ix)
            obj.CodInventario = grdRecetaIng.TextMatrix(i, 2)
            obj.Contrato = grdRecetaIng.TextMatrix(i, 1)
            obj.fechaIni = grdRecetaIng.TextMatrix(i, 8)
            obj.Frecuencia = "1;2;3;4;5;6;7;8;9;10;11;12"
            obj.FechaFin = DateAdd("yyyy", 100, grdRecetaIng.TextMatrix(i, 8))
            obj.referencia = Int(grdRecetaIng.TextMatrix(i, 1))
            'obj.Frecuencia = grdDetalleFact.TextMatrix(i, 7)
        Next i

    pc.Grabar

    
    Set item = Nothing
    'grdReceta.subtotal flexSTSum, 2, 7, , , , , , , True
'''    If grdRecetaIng.Rows <> grdReceta.Rows Then
'''        'MsgBox gnc.numtrans
'''        grd.TextMatrix(i, COL_RESULTADO) = "Error"
'''    Else
'''        grd.TextMatrix(i, COL_RESULTADO) = "OK"
'''    End If
End Sub

