VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListaPlantilla 
   Caption         =   "Listado de Plantillas"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VSFlex7Ctl.VSFlexGrid grd 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5895
      _cx             =   10398
      _cy             =   9763
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
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmListaPlantilla.frx":0000
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
      MergeCells      =   4
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
      ExplorerBar     =   7
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
      AllowUserFreezing=   3
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSComctlLib.ImageList imlLista 
      Left            =   5640
      Top             =   0
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaPlantilla.frx":009F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaPlantilla.frx":01B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaPlantilla.frx":02C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListaPlantilla.frx":03DB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "imlLista"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   200
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Agregar"
            Object.ToolTipText     =   "Agregar (INSERT)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Modificar"
            Object.ToolTipText     =   "Modificar (ENTER)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Eliminar"
            Object.ToolTipText     =   "Eliminar (DEL)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copiar"
            Object.ToolTipText     =   "Copiar (CTRL+INS)"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmListaPlantilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mCn As ADODB.Connection
Const COL_NUM = 0
Const COL_IDP = 1
Const COL_COD = 2
Const COL_DES = 3
Const COL_VAL = 4

Public Sub Inicio(ByVal tag As String, ByVal cn As ADODB.Connection)
    Me.tag = tag
    Set mCn = cn
    CargarListaPlantilla
    Me.Show vbModal
End Sub

Private Sub CargarListaPlantilla()
    Dim sql As String, rs As Recordset
    
    sql = "SELECT IdPlantilla, CodPlantilla, Descripcion, BandValida " & _
          "FROM Plantilla_EI "
    If Me.tag = "PlantillasExportar" Then
        sql = sql & "WHERE Tipo=0 "
        Me.Caption = "Listado de Plantillas para Exportación"
    Else
        sql = sql & "WHERE Tipo=1 "
        Me.Caption = "Listado de Plantillas para Importación"
    End If
    sql = sql & "ORDER BY CodPlantilla"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mCn, adOpenStatic, adLockReadOnly
    
    With rs
        If Not (.BOF And .EOF) Then
            .MoveLast
            .MoveFirst
            grd.LoadArray .GetRows(.RecordCount)
        End If
    End With
    ConfigCols
    GNPoneNumFila grd, False
    AjustarAutoSize grd, -1, -1
    Set rs = Nothing
End Sub

Private Sub ConfigCols()
    With grd
        .FormatString = ">#|IdPlantilla|<Codigo|<Descripcion|Válida"
        
        .ColWidth(COL_NUM) = 700
        .ColWidth(COL_COD) = 1500
        .ColWidth(COL_DES) = 3000
        
        .ColHidden(COL_IDP) = True
        
        .ColDataType(COL_VAL) = flexDTBoolean
    End With
End Sub

Private Sub Form_Activate()
    grd.SetFocus
End Sub

Private Sub Form_Resize()
    If Me.WindowState <> vbMinimized Then
        grd.Move 0, tlb1.Height, Me.ScaleWidth, Me.ScaleHeight - tlb1.Height
    End If
End Sub

Private Sub Agregar()
    Dim obj As clsPlantilla
    
    Set obj = New clsPlantilla
    obj.Coneccion = mCn
    If Me.tag = "PlantillasExportar" Then  'Diego 04/03/2004
        obj.Tipo = 0
     Else
        obj.Tipo = 1
     End If
    If frmPlantilla_EI.Inicio("Agregar", obj) Then CargarListaPlantilla
End Sub

Private Sub Modificar()
    Dim obj As clsPlantilla, id As Long
    
    If grd.Rows <= grd.FixedRows Then Exit Sub
    If grd.Row < grd.FixedRows Then Exit Sub
    
    id = grd.ValueMatrix(grd.Row, COL_IDP)
    Set obj = New clsPlantilla
    obj.Coneccion = mCn
    obj.Recuperar id
    If frmPlantilla_EI.Inicio("Modificar", obj) Then CargarListaPlantilla
End Sub

Private Sub Eliminar()
    Dim sql As String, id As Long, r As Integer
    If grd.Rows <= grd.FixedRows Then Exit Sub
    If grd.Row < grd.FixedRows Then Exit Sub
    
    id = grd.ValueMatrix(grd.Row, COL_IDP)
    sql = "DELETE FROM Plantilla_EI WHERE IdPlantilla = " & id
    r = MsgBox("Desea Eliminar la Plantilla: " & vbCrLf & _
             "Codigo:      " & grd.TextMatrix(grd.Row, COL_COD) & vbCrLf & _
             "Descripción: " & grd.TextMatrix(grd.Row, COL_DES), vbQuestion + vbYesNo)
    If r = vbYes Then
        mCn.Execute sql
        CargarListaPlantilla
    End If
End Sub

Private Sub Copiar()
    Dim obj As clsPlantilla, id As Long
    
    If grd.Rows <= grd.FixedRows Then Exit Sub
    If grd.Row < grd.FixedRows Then Exit Sub
    
    id = grd.ValueMatrix(grd.Row, COL_IDP)
    Set obj = New clsPlantilla
    obj.Coneccion = mCn
    obj.Recuperar id
    Set obj = obj.Clone
    If frmPlantilla_EI.Inicio("Copiar", obj) Then CargarListaPlantilla
End Sub

Private Sub grd_DblClick()
    If (grd.FixedRows <> grd.Rows) And (grd.Row <> 0) Then Modificar
End Sub

Private Sub grd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyInsert
        If (Shift And vbCtrlMask) Then
            Copiar
        Else
            Agregar
        End If
        KeyCode = 0
    Case vbKeyReturn
        Modificar
        KeyCode = 0
    Case vbKeyDelete
        Eliminar
        KeyCode = 0
    Case vbKeyEscape
        Unload Me
        KeyCode = 0
    End Select
End Sub

Private Sub tlb1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Agregar"
        Agregar
    Case "Modificar"
        Modificar
    Case "Eliminar"
        Eliminar
    Case "Copiar"
        Copiar
    End Select
End Sub
