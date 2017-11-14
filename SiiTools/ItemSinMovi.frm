VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmItemSinMovi 
   Caption         =   "Eliminación de items sin movimiento"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6465
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4470
   ScaleWidth      =   6465
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6465
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3972
      Width           =   6468
      Begin VB.CommandButton cmdSelecTodo 
         Caption         =   "&Selec. Todo"
         Height          =   372
         Left            =   1440
         TabIndex        =   1
         Top             =   0
         Width           =   1212
      End
      Begin VB.CommandButton cmdVerificar 
         Caption         =   "&Verificar"
         Height          =   372
         Left            =   168
         TabIndex        =   0
         Top             =   0
         Width           =   1212
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   4968
         TabIndex        =   3
         Top             =   0
         Width           =   1212
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   372
         Left            =   3648
         TabIndex        =   2
         Top             =   0
         Width           =   1212
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Align           =   1  'Align Top
      Height          =   3252
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6468
      _cx             =   11409
      _cy             =   5736
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
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   3
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
Attribute VB_Name = "frmItemSinMovi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private nombre As String


Public Sub Inicio()
    On Error GoTo ErrTrap
    nombre = "Inventario"
    Me.Show
    Me.ZOrder
    ConfigCols
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub ConfigCols()
    With grd
        .FormatString = "^#|Id|<Código|<Descripción|<Fecha de grabación"
        .ColHidden(1) = True
        .ColDataType(4) = flexDTDate        'Fecha grabado
        GNPoneNumFila grd, False
        .AutoSize 0, grd.Cols - 1
    End With
End Sub


Private Sub cmdAceptar_Click()
    Dim i As Long, iv As IVinventario, id As Long, r As Long, cod As String, gc As GNCentroCosto
    On Error GoTo ErrTrap
    If nombre = "Inventario" Then
        'Si no hay nada en la grilla
        If grd.Rows <= grd.FixedRows Then
            MsgBox "No hay ningún item.", vbInformation
            Exit Sub
        End If
        
        'Confirma si está seguro
        If MsgBox("Está seguro de que desea eliminar los items seleccionados? " & vbCr & _
                  "Este proceso no podrá deshacerse.", _
                  vbYesNo + vbQuestion) <> vbYes Then Exit Sub
        
        cmdCancelar.Enabled = False
        Screen.MousePointer = vbHourglass
        
        With grd
            For i = .SelectedRows - 1 To 0 Step -1
                r = .SelectedRow(i)
                Debug.Print "Deleting row: " & r
                id = .ValueMatrix(r, 1)       'Col 1 = idInventario
                cod = .TextMatrix(r, 2)
                Set iv = gobjMain.EmpresaActual.RecuperaIVInventario(id)
                If Not (iv Is Nothing) Then
                    iv.Eliminar
                    .RemoveItem r
                    grd.Refresh
                Else
                    MsgBox "No se puede encontrar el item '" & cod & "'", vbInformation
                    Exit For
                End If
            Next i
        End With
    ElseIf nombre = "CentroCosto" Then
        'Si no hay nada en la grilla
        If grd.Rows <= grd.FixedRows Then
            MsgBox "No hay ningún Centro.", vbInformation
            Exit Sub
        End If
        
        'Confirma si está seguro
        If MsgBox("Está seguro de que desea eliminar los Centro de Costo seleccionados? " & vbCr & _
                  "Este proceso no podrá deshacerse.", _
                  vbYesNo + vbQuestion) <> vbYes Then Exit Sub
        
        cmdCancelar.Enabled = False
        Screen.MousePointer = vbHourglass
        
        With grd
            For i = .SelectedRows - 1 To 0 Step -1
                r = .SelectedRow(i)
                Debug.Print "Deleting row: " & r
                id = .ValueMatrix(r, 1)       'Col 1 = idInventario
                cod = .TextMatrix(r, 2)
                Set gc = gobjMain.EmpresaActual.RecuperaGNCentroCosto(id)
                If Not (gc Is Nothing) Then
                    gc.Eliminar
                    .RemoveItem r
                    grd.Refresh
                Else
                    MsgBox "No se puede encontrar el Centro de Costo '" & cod & "'", vbInformation
                    Exit For
                End If
            Next i
        End With
    
    End If
    
    Screen.MousePointer = 0
    cmdCancelar.Enabled = True
    cmdCancelar.SetFocus
    Exit Sub
ErrTrap:
    cmdCancelar.Enabled = True
    Screen.MousePointer = 0
    DispErr
    Exit Sub

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdSelecTodo_Click()
    Dim i As Long
    On Error GoTo ErrTrap
    With grd
        For i = .FixedRows To .Rows - 1
            .IsSelected(i) = True
        Next i
    End With
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub cmdVerificar_Click()
    Dim v As Variant, rs As Object, sql As String
    On Error GoTo ErrTrap
    If nombre = "Inventario" Then
        sql = "SELECT iv.IdInventario, iv.CodInventario, " & _
                     "iv.Descripcion, iv.FechaGrabado " & _
              "FROM IVInventario iv " & _
              "WHERE NOT EXISTS " & _
                    "(SELECT * FROM IVKardex ivk " & _
                    "WHERE ivk.IdInventario=iv.IdInventario) " & _
              "ORDER BY CodInventario"
    ElseIf nombre = "CentroCosto" Then
    
        sql = " SELECT gc.Idcentro, gc.Codcentro, gc.Descripcion, gc.FechaGrabado"
        sql = sql & " FROM gncentrocosto gc"
        sql = sql & " WHERE NOT EXISTS (SELECT * FROM gncomprobante g WHERE gc.Idcentro=g.Idcentro)"
        sql = sql & " ORDER BY idcentro"
    
    End If
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    
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
        If nombre = "Inventario" Then
            MsgBox "No hay ningún item sin movimiento.", vbInformation
        ElseIf nombre = "CentroCosto" Then
            MsgBox "No hay ningún Centro sin movimiento.", vbInformation
        End If
        cmdCancelar.SetFocus
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



Public Sub InicioCentroCosto()
    On Error GoTo ErrTrap
    nombre = "CentroCosto"
    Me.Show
    Me.ZOrder
    ConfigCols
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

