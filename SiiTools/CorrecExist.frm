VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Begin VB.Form frmCorrecExist 
   Caption         =   "Corrección de existencias"
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
      Height          =   492
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6585
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4212
      Width           =   6585
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Corregir"
         Enabled         =   0   'False
         Height          =   372
         Left            =   2688
         TabIndex        =   1
         Top             =   0
         Width           =   1212
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   4968
         TabIndex        =   2
         Top             =   0
         Width           =   1212
      End
      Begin VB.CommandButton cmdVerificar 
         Caption         =   "&Verificar"
         Height          =   372
         Left            =   1248
         TabIndex        =   0
         Top             =   0
         Width           =   1212
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Align           =   1  'Align Top
      Height          =   3255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6585
      _cx             =   11615
      _cy             =   5741
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
End
Attribute VB_Name = "frmCorrecExist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objcond As Condicion

Public Sub Inicio(ByVal tag As String)
    Dim i As Integer
    On Error GoTo ErrTrap
    Me.tag = tag
    Me.Show
    Me.ZOrder
    ConfigCols
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub cmdAceptar_Click()
    On Error GoTo ErrTrap
    If Me.tag = "Comprometido" Then
        If grd.Rows <= grd.FixedRows Then
            MsgBox "No hay ningún item con comprometido incorrecto"
            Exit Sub
        End If
        cmdCancelar.Enabled = False
        Screen.MousePointer = vbHourglass
        gobjMain.EmpresaActual.CorregirComprometido
        Screen.MousePointer = 0
        cmdCancelar.Enabled = True
        MsgBox "Los comprometidos han sido corregido."
    ElseIf Me.tag = "AFExistencias" Then
        If grd.Rows <= grd.FixedRows Then
            MsgBox "No hay ningún activo con las depreciacines incorrecta."
            Exit Sub
        End If
        cmdCancelar.Enabled = False
        Screen.MousePointer = vbHourglass
        gobjMain.EmpresaActual.CorregirAFExistencia
        
    ElseIf Me.tag = "AFExistenciasCustodio" Then
        
        If grd.Rows <= grd.FixedRows Then
            MsgBox "No hay ningún activo con las depreciacines incorrecta."
            Exit Sub
        End If
        
        cmdCancelar.Enabled = False
        Screen.MousePointer = vbHourglass
        gobjMain.EmpresaActual.CorregirExistenciaAFCustodio
        Screen.MousePointer = 0
        cmdCancelar.Enabled = True
        MsgBox "Las Existencias de los Custodios han sido corregido."
    
    ElseIf Me.tag = "ExistenciasDocum" Then
        If grd.Rows <= grd.FixedRows Then
            MsgBox "No hay ningún Documento con existencia incorrecta."
            Exit Sub
        End If
        cmdCancelar.Enabled = False
        Screen.MousePointer = vbHourglass
        gobjMain.EmpresaActual.CorregirExistenciaDocum
        Screen.MousePointer = 0
        cmdCancelar.Enabled = True
        MsgBox "Las existencias de Documentos han sido corregido."
    ElseIf Me.tag = "ExistenciasDocumRuta" Then
        If grd.Rows <= grd.FixedRows Then
            MsgBox "No hay ningún Documento con existencia incorrecta."
            Exit Sub
        End If
        cmdCancelar.Enabled = False
        Screen.MousePointer = vbHourglass
        gobjMain.EmpresaActual.Empresa2.CorregirExistenciaDocumRuta
        Screen.MousePointer = 0
        cmdCancelar.Enabled = True
        MsgBox "Las existencias de Documentos han sido corregido."
    ElseIf Me.tag = "ExistenciasSerie" Then
        If grd.Rows <= grd.FixedRows Then
            MsgBox "No hay ningún item con existencia incorrecta de Series."
            Exit Sub
        End If
        cmdCancelar.Enabled = False
        Screen.MousePointer = vbHourglass
        gobjMain.EmpresaActual.CorregirExistenciaSerie
        Screen.MousePointer = 0
        cmdCancelar.Enabled = True
        MsgBox "Las existencias de series han sido corregido."

    Else
    
        If grd.Rows <= grd.FixedRows Then
            MsgBox "No hay ningún item con existencia incorrecta."
            Exit Sub
        End If
        
        cmdCancelar.Enabled = False
        Screen.MousePointer = vbHourglass
        gobjMain.EmpresaActual.CorregirExistencia
        
        
        Screen.MousePointer = 0
        cmdCancelar.Enabled = True
        MsgBox "Las existencias han sido corregido."

    
    End If
    If cmdCancelar.Enabled Then
        cmdCancelar.SetFocus
    End If
    Exit Sub
ErrTrap:
    cmdCancelar.Enabled = True
    Screen.MousePointer = 0
    DispErr
    Exit Sub
End Sub


Private Sub ConfigCols()
    With grd
        If Me.tag = "Comprometido" Then
            .FormatString = "^#|Id|<Código|<Descripcion|<IdBodega|<Bodega|>Vendido|>Entregado|>Compr|>Compr.Correcta|>Diferencia"
        ElseIf Me.tag = "AFExistencias" Then
            .FormatString = "^#|Id|<Código|<IdBodega|<Bodega|>Exist|>Exist.Correcta|>Diferencia"
        ElseIf Me.tag = "AFExistenciasCustodio" Then
            .FormatString = "^#|Id|<Código|<Idprovcli|<Custodio|>Exist|>Exist.Correcta|>Diferencia"
        ElseIf Me.tag = "ExistenciasDocum" Then
            .FormatString = "^#|Id|<Código|<IdProvcli|<Empleado|>Exist|>Exist.Correcta|>Diferencia"
         ElseIf Me.tag = "ExistenciasDocumRuta" Then
            .FormatString = "^#|Transid|Id|<Codigo|>Exist|>Exist.Correcta|>Diferencia"
        Else
            .FormatString = "^#|Id|<Código|<IdBodega|<Bodega|>Exist|>Exist.Correcta|>Diferencia"
        End If
        .ColHidden(1) = True
        If Me.tag <> "ExistenciasDocumRuta" Then
            .ColHidden(4) = True
        End If
        GNPoneNumFila grd, False
        .AutoSize 0, grd.Cols - 1
    End With
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub


Private Sub cmdVerificar_Click()
    Dim v As Variant, rs As Object
    On Error GoTo ErrTrap
    If Me.tag = "Comprometido" Then
        Set objcond = gobjMain.objCondicion
        If Not (frmB_VxTransComp.Inicio(objcond, "Comprometido")) Then
            grd.SetFocus
            Exit Sub
        End If
   
        Set rs = gobjMain.EmpresaActual.ConsIVCorrecCompr(objcond.CodTrans, objcond.Bienes)
        If Not rs.EOF Then
            v = MiGetRows(rs)
            grd.Redraw = flexRDNone
            grd.LoadArray v
            ConfigCols
            grd.Redraw = flexRDDirect
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
        Else
            grd.Rows = grd.FixedRows
            ConfigCols
            MsgBox "No hay ningún item con el comprometido incorrecto."
            cmdCancelar.SetFocus
        End If
    ElseIf Me.tag = "AFExistencias" Then
    
    Set rs = gobjMain.EmpresaActual.ConsAFCorrecExist
        If Not rs.EOF Then
            v = MiGetRows(rs)
            
            grd.Redraw = flexRDNone
            grd.LoadArray v
            ConfigCols
            grd.Redraw = flexRDDirect
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
        Else
            grd.Rows = grd.FixedRows
            ConfigCols
            MsgBox "No hay ningún item con Depreciaciones incorrecta."
            If cmdCancelar.Enabled Then
                cmdCancelar.SetFocus
            End If
        End If
    ElseIf Me.tag = "AFExistenciasCustodio" Then
        
        Set rs = gobjMain.EmpresaActual.ConsAFCorrecExistCustodio
        If Not rs.EOF Then
            v = MiGetRows(rs)
            
            grd.Redraw = flexRDNone
            grd.LoadArray v
            ConfigCols
            grd.Redraw = flexRDDirect
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
        Else
            grd.Rows = grd.FixedRows
            ConfigCols
            MsgBox "No hay ningún Activo con existencia incorrecta."
            cmdCancelar.SetFocus
        End If
        
    ElseIf Me.tag = "ExistenciasDocum" Then
        
        Set rs = gobjMain.EmpresaActual.ConsIVCorrecExistDocum
        If Not rs.EOF Then
            v = MiGetRows(rs)
            grd.Redraw = flexRDNone
            grd.LoadArray v
            ConfigCols
            grd.Redraw = flexRDDirect
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
        Else
            grd.Rows = grd.FixedRows
            ConfigCols
            MsgBox "No hay ningún documento con existencia incorrecta."
            cmdCancelar.SetFocus
        End If
       ElseIf Me.tag = "ExistenciasDocumRuta" Then
        Set rs = gobjMain.EmpresaActual.Empresa2.ConsIVCorrecExistDocumRuta
            If Not rs.EOF Then
                v = MiGetRows(rs)
                grd.Redraw = flexRDNone
                grd.LoadArray v
                ConfigCols
                grd.Redraw = flexRDDirect
                cmdAceptar.Enabled = True
                cmdAceptar.SetFocus
            Else
                grd.Rows = grd.FixedRows
                ConfigCols
                MsgBox "No hay ningún documento con existencia incorrecta."
                cmdCancelar.SetFocus
            End If
    ElseIf Me.tag = "ExistenciasSerie" Then
    Set rs = gobjMain.EmpresaActual.ConsIVCorrecExistSerie
        If Not rs.EOF Then
            v = MiGetRows(rs)
            
            grd.Redraw = flexRDNone
            grd.LoadArray v
            ConfigCols
            grd.Redraw = flexRDDirect
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
        Else
            grd.Rows = grd.FixedRows
            ConfigCols
            MsgBox "No hay ningún Num Serie con existencia incorrecta."
            cmdCancelar.SetFocus
        End If
Else
        If gobjMain.EmpresaActual.GNOpcion.IVKTipoDatoDouble Then
            Set rs = gobjMain.EmpresaActual.ConsIVCorrecExistDou
        Else
            Set rs = gobjMain.EmpresaActual.ConsIVCorrecExist
        End If
        If Not rs.EOF Then
            v = MiGetRows(rs)
            
            grd.Redraw = flexRDNone
            grd.LoadArray v
            ConfigCols
            grd.Redraw = flexRDDirect
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
        Else
            grd.Rows = grd.FixedRows
            ConfigCols
            MsgBox "No hay ningún item con existencia incorrecta."
            cmdCancelar.SetFocus
        End If
    End If
        
    
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
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
    grd.Move 0, grd.Top, Me.ScaleWidth, Me.ScaleHeight - grd.Top - pic1.Height - 80
'    prg1.Width = Me.ScaleWidth - (prg1.Left * 2)
End Sub


Public Sub InicioComprometido(ByVal tag As String)
    Dim i As Integer
    On Error GoTo ErrTrap
    Me.tag = tag
    Me.Show
    Me.ZOrder
    Me.Caption = "Correción de Comprometido"
    ConfigCols
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Public Sub InicioDepreciaciones(ByVal tag As String)
    Dim i As Integer
    On Error GoTo ErrTrap
    Me.tag = tag
    Me.Show
    Me.ZOrder
    Me.Caption = "Correción del Número Depreciaciones"
    ConfigCols
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub


Public Sub InicioCustodios(ByVal tag As String)
    Dim i As Integer
    On Error GoTo ErrTrap
    Me.tag = tag
    Me.Show
    Me.ZOrder
    ConfigCols
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Public Sub InicioDocum(ByVal tag As String)
    Dim i As Integer
    On Error GoTo ErrTrap
    Me.tag = tag
    Me.Show
    Me.ZOrder
    ConfigCols
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

