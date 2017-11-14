VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAFItemIVA_Cuenta 
   Caption         =   "Actualización de IVA / Cuenta contable / PROcLI"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6240
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   6240
   WindowState     =   2  'Maximized
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   2052
      Left            =   480
      TabIndex        =   3
      Top             =   840
      Width           =   5052
      _cx             =   8911
      _cy             =   3619
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
   Begin MSComctlLib.ImageList img1 
      Left            =   7560
      Top             =   120
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AFItemIVA_Cuenta.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AFItemIVA_Cuenta.frx":0114
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AFItemIVA_Cuenta.frx":0568
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AFItemIVA_Cuenta.frx":067C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "AFItemIVA_Cuenta.frx":0790
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   1005
      ButtonWidth     =   1402
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "img1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar"
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar (F5)"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Asignar"
            Key             =   "Asignar"
            Description     =   "Asignar un valor"
            Object.ToolTipText     =   "Asignar un valor (F6)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grabar"
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar (F3)"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Imprimir"
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir (Ctrl+P)"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cerrar"
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar (ESC)"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   492
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   6240
      TabIndex        =   0
      Top             =   4185
      Width           =   6240
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "frmAFItemIVA_Cuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit

Private mProcesando As Boolean
Private mCancelado As Boolean
Private mcolItemsSelec As Collection      'Coleccion de items
'jeaa 24/09/04 asignacion de grupo a los items
'AFInventario
Const AFGRUPO1 = 4
Const AFGRUPO2 = 5
Const AFGRUPO3 = 6
Const AFGRUPO4 = 7
Const AFGRUPO5 = 8
'PC_PROV_CLI
Const PCGRUPO1 = 3
Const PCGRUPO2 = 4
Const PCGRUPO3 = 5
Const PCGRUPO4 = 6 'Agregado AUC 03/10/2005


'Private mobjItem As AFInventario

Public Sub Inicio(ByVal tag As String)
    Dim I As Integer
    On Error GoTo ErrTrap
    
    Me.tag = tag            'Guarda en la propiedad Tag para distinguir después
    Me.Show
    Me.ZOrder
    
    Select Case Me.tag
    Case "ITEM_VIDAUTIL"
        Me.Caption = "Actualización de Vida Util de ítems"
    Case "CUENTA"
        Me.Caption = "Actualización de cuentas contables de ítems"
    Case "ITEM_AFGRUPOS"    'jeaa 24/09/04 asignacion de grupo a los items
        Me.Caption = "Asiganción de Grupos a Items "
    Case "PCGRUPOS_PROV"   'jeaa 24/09/04 asignacion de grupo a los prov
        Me.Caption = "Asiganción de Grupo a Proveedores"
    Case "PCGRUPOS_CLI"    'jeaa 24/09/04 asignacion de grupo a los cli
        Me.Caption = "Asiganción de Grupo a Clientes"
    Case "FRACCION"    'jeaa 13/04/05 asignacion de bandera fraccion
        Me.Caption = "Asiganción Bandera para Venta en Fracción"
    Case "AREA"    'jeaa 15/09/05 asignacion de bandera AREA
        Me.Caption = "Asiganción Bandera para Venta por Areas"
    Case "VENTA"    'jeaa 26/12/05 asignacion de bandera VENTA
        Me.Caption = "Asiganción Bandera para Venta "
    Case "ITEM_VIDAUTIL"
        Me.Caption = "Actualización de Vida Util de ítems"
    Case "COSTOUI"
        Me.Caption = "Actualización del Costo Ultimo Ingreso de ítems"
    Case "SRI"
        Me.Caption = "Actualización de Numero de autorizacion SRI y Fecha Caducidad"
    Case "MINMAX"
        Me.Caption = "Actualización de Existencias Mínimas y Máximas"
    Case "IVEXIST"
        Me.Caption = "Llenado de Items en Tabla Existencias "
    Case "CUENTA_PRESUP"
        Me.Caption = "Actualizacion de Presupuesto a cuentas contables "   'jeaa-09/01/2009
    Case "DIASREPO"
        Me.Caption = "Actualizacion de Dias de Reposicion "   'jeaa-09/01/2009
    
    End Select
       
    'Inicializa la grilla
    grd.Rows = grd.FixedRows
    ConfigCols
    
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF3
        Grabar
        KeyCode = 0
    Case vbKeyF5
        Select Case Me.tag
            Case "ITEM_VIDAUTIL", "CUENTA", "ITEM_AFGRUPOS", "FRACCION", "AREA", "VENTA": Buscar
            Case "CUENTA_PROV", "PCGRUPOS_PROV": BuscarPC True
            Case "CUENTA_CLI", "PCGRUPOS_CLI": BuscarPC False
            Case "CUENTA_LOCAL": BuscarCuenta 'jeaa 21/01/04
            Case "ITEM_FAMILIA": BuscarItemFamilia 'jeaa 21/01/04
            Case "COSTOUI": BuscarCostoUltimo
            Case "SRI": BuscarSRI
            Case "MINMAX": BuscarMINMAX
            Case "IVEXIST": BuscarIvExist
            Case "CUENTA_PRESUP": BuscarCuenta 'jeaa 21/01/04
            Case "DIASREPO": Buscar 'jeaa 21/01/04
        End Select
        KeyCode = 0
    Case vbKeyF6
        Asignar
        KeyCode = 0
    Case vbKeyP
        If Shift And vbCtrlMask Then
            Imprimir
            KeyCode = 0
        End If
    Case vbKeyEscape
        Cerrar
    Case Else
        MoverCampo Me, KeyCode, Shift, True
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ImpideSonidoEnter Me, KeyAscii
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mProcesando Then
        Cancel = 1      'No permitir cerrar mientras procesa
    Else
        Me.Hide         'Se pone esto para evitar el posible BUG de Windows98
    End If
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    grd.Move 0, tlb1.Height, Me.ScaleWidth, Me.ScaleHeight - tlb1.Height - pic1.Height - 80
    prg1.Width = Me.ScaleWidth - (prg1.Left * 2)
End Sub



Private Sub grd_AfterEdit(ByVal Row As Long, ByVal col As Long)
    Select Case Me.tag
        Case "CUENTA_LOCAL"
            RecuperaCodLocal grd.TextMatrix(Row, grd.ColIndex("Sucursal")), Row
        Case "ITEM_FAMILIA"
            RecuperaCodFamilia grd.TextMatrix(Row, grd.ColIndex("Familia")), Row
        Case "CUENTA_PRESUP"
'            RecuperaCodLocal grd.TextMatrix(Row, grd.ColIndex("Presupuesto")), Row
    
    End Select
End Sub

Private Sub grd_BeforeEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
    If Row < grd.FixedRows Then Cancel = True
    If grd.IsSubtotal(Row) = True Then Cancel = True
    If grd.ColData(col) < 0 Then Cancel = True
    
    If Not Cancel Then
        'Longitud maxima para editar
        grd.EditMaxLength = grd.ColData(col)
    End If
End Sub

Private Sub grd_BeforeSort(ByVal col As Long, Order As Integer)
    'Impide mientras está procesando
    If mProcesando Then Order = flexSortNone
End Sub

Private Sub grd_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
    With grd
        Select Case Me.tag
        Case "ITEM_VIDAUTIL"
            If Not IsNumeric(.EditText) Then
                MsgBox "Debe ingresar un valor numérico.", vbInformation
                .SetFocus
                Cancel = True
            End If
        Case "CUENTA", "CUENTA_PROV", "CUENTA_CLI", "CUENTA_LOCAL", "ITEM_FAMILIA", "ITEM_RECETA", "MINMAX" 'jeaa 21/01/04
            If Right$(.EditText, 1) = "*" Then
                MsgBox "No puede seleccionar una cuenta de mayor.", vbInformation
                .SetFocus
                Cancel = True
            End If
        Case "CUENTA_PRESUP"
            If Right$(.EditText, 1) = "*" Then
                MsgBox "No puede seleccionar una cuenta de mayor.", vbInformation
                .SetFocus
                Cancel = True
            End If
        
        Case "COSTOUI"
            If Not IsNumeric(.EditText) Then
                MsgBox "Debe ingresar un valor numérico.", vbInformation
                .SetFocus
                Cancel = True
            End If
        Case "SRI"
            If col <> 8 Then
                If Not IsNumeric(.EditText) Then
                    MsgBox "Debe ingresar un valor numérico.", vbInformation
                    .SetFocus
                    Cancel = True
                End If
            End If
        Case "IVEXIST" 'jeaa 21/01/04
            If Right$(.EditText, 1) = "*" Then
                MsgBox "No puede seleccionar una cuenta de mayor.", vbInformation
                .SetFocus
                Cancel = True
            End If
        Case "DIASREPO"
            If Not IsNumeric(.EditText) Then
                MsgBox "Debe ingresar un valor numérico.", vbInformation
                .SetFocus
                Cancel = True
            End If
        End Select
    End With
End Sub

Private Sub tlb1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Buscar":
        Select Case Me.tag
            Case "ITEM_VIDAUTIL", "CUENTA", "ITEM_AFGRUPOS", "FRACCION", "AREA", "VENTA": Buscar
            Case "CUENTA_PROV", "PCGRUPOS_PROV": BuscarPC True
            Case "CUENTA_CLI", "PCGRUPOS_CLI": BuscarPC False
            Case "CUENTA_LOCAL": BuscarCuenta 'jeaa 21/01/04
            Case "ITEM_FAMILIA": BuscarItemFamilia
            Case "COSTOUI": BuscarCostoUltimo
            Case "SRI": BuscarSRI
            Case "MINMAX": BuscarMINMAX
            Case "IVEXIST": BuscarIvExist
            Case "CUENTA_PRESUP": BuscarCuenta 'jeaa 21/01/04
            Case "DIASREPO": Buscar
        End Select
    Case "Asignar":     Asignar
    Case "Grabar":
        Select Case Me.tag
            Case "ITEM_VIDAUTIL", "CUENTA", "ITEM_AFGRUPOS", "FRACCION", "AREA", "VENTA", "COSTOUI", "MINMAX", "IVEXIST": Grabar
            Case "CUENTA_PROV", "CUENTA_CLI", "PCGRUPOS_PROV", "PCGRUPOS_CLI": GrabarPC
            Case "CUENTA_LOCAL": GrabarLocal 'jeaa 21/01/04
            Case "SRI": Grabar
            Case "CUENTA_PRESUP": GrabarPresupuesto 'jeaa 09/01/2009
            Case "DIASREPO": Grabar
        End Select
    Case "Imprimir":    Imprimir
    Case "Cerrar":      Cerrar
    End Select
End Sub

Private Sub Buscar()
    Static coditem As String, CodAlt As String, _
           Desc As String, _
           codg As String, Numg As Integer, bandIVA As Boolean, bandFraccion As Boolean
    Dim codg1 As String, codg2 As String, codg3 As String, codg4 As String, codg5 As String
    Dim sql As String, cond As String, rs As Recordset, comodin As String
    On Error GoTo ErrTrap
    'If Me.tag <> "CUENTA" Then Exit Sub
    
    #If DAOLIB Then
        comodin = "*"
    #Else
        comodin = "%"
    #End If
'    comodin = "%"
    'Abre la pantalla de búsqueda
    If Not frmAFBusqueda.Inicio( _
                coditem, _
                CodAlt, _
                Desc, _
                codg1, codg2, codg3, codg4, codg5, _
                Numg, _
                bandIVA, _
                Me.tag) Then
      'if not frmAFBusqueda.InicioTrans (
        'Si fue cancelada la busqueda, sale no mas
        grd.SetFocus
        Exit Sub
    End If
    
    'Cambia la forma de cursor
    MensajeStatus MSG_PREPARA, vbHourglass
    
    'Compone la cadena de SQL
    sql = "SELECT CodInventario, CodAlterno1, Descripcion "
    Select Case Me.tag
    Case "ITEM_VIDAUTIL"
        sql = sql & ", Vidautil "
    Case "CUENTA"
        sql = sql & ", CodCuentaActivo, CodCuentaCosto, CodCuentaVenta, CodCuentaDepreGasto, CodCuentaDepreAcumulada, CodCuentaRevaloriza, CodCuentaDepRevaloriza "
    Case "ITEM_AFGRUPOS"    'jeaa 24/09/04 asignacion de grupo a los items
        sql = sql & ", codGrupo1, codgrupo2, codGrupo3, codGrupo4 , codGrupo5 "
    Case "FRACCION"
        sql = sql & ", BandFraccion "
    Case "AREA" 'jeaa 15/09/2005
        sql = sql & ", BandArea "
    Case "VENTA" 'jeaa 26/12/2005
        sql = sql & ", BandVenta "
    Case "COSTOUI"
        sql = sql & ", CostoUltimoIngreso "
    Case "MINMAX"
        sql = sql & ", CodCuentaActivo, CodCuentaCosto, CodCuentaVenta "
    Case "IVEXIST"
        sql = sql & ", CodCuentaActivo, CodCuentaCosto, CodCuentaVenta "
    Case "DIASREPO"
        sql = sql & ", TiempoReposicion "
    
    End Select
    sql = sql & "FROM vwAFInventarioRecuperar "
    'CodInventario
    If Len(coditem) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodInventario LIKE '" & coditem & comodin & "') "
    End If
    
    'CodAlterno
    If Len(CodAlt) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "((CodAlterno1 LIKE '" & CodAlt & comodin & "') " & _
                      "OR (CodAlterno2 LIKE '" & CodAlt & comodin & "')) "
    End If
    
    'Descripcion
    If Len(Desc) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(Descripcion LIKE '" & Desc & comodin & "') "
    End If
    
'    'Grupo
'    If Len(codg) > 0 Then
'        If Len(Cond) > 0 Then Cond = Cond & "AND "
'        Cond = Cond & "(CodGrupo" & Numg & " = '" & codg & "') "
'    End If
    
    If Len(codg1) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodGrupo1" & " = '" & codg1 & "') "
    End If
    
    If Len(codg2) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodGrupo2" & " = '" & codg2 & "') "
    End If
    
    If Len(codg3) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodGrupo3" & " = '" & codg3 & "') "
    End If
    
    If Len(codg4) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodGrupo4" & " = '" & codg4 & "') "
    End If
    
    If Len(codg5) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodGrupo5" & " = '" & codg5 & "') "
    End If
    
    
    
    If bandIVA Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & " PorcentajeIVA <> 0 "
        If Me.tag = "COSTOUI" Then
                If Len(cond) > 0 Then cond = cond & "AND "
                cond = cond & " (costoultimoingreso Is Null Or costoultimoingreso = 0) "
        End If
    End If
    
    If Len(cond) > 0 Then sql = sql & " WHERE " & cond
    sql = sql & " ORDER BY CodInventario "
    
    
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    With grd
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If Not rs.EOF Then .LoadArray MiGetRows(rs)
        ConfigCols
        .Redraw = flexRDBuffered
        .SetFocus
    End With
    
    MensajeStatus
    Exit Sub
ErrTrap:
    grd.Redraw = flexRDBuffered
    MensajeStatus
    DispErr
    grd.SetFocus
    Exit Sub
End Sub

Private Sub BuscarPC(ByVal bandprov As Boolean)
    Static CodProvCli As String, _
           nombre As String, _
           codg As String, Numg As Integer
    Dim sql As String, cond As String, rs As Recordset, comodin As String
    Dim provCli As String
    On Error GoTo ErrTrap

#If DAOLIB Then
    comodin = "*"
#Else
    comodin = "%"
#End If
    provCli = IIf(Me.tag = "CUENTA_PROV" Or Me.tag = "PCGRUPOS_PROV", "Proveedores", "Clientes")
    
    
    'Abre la pantalla de búsqueda
    frmPCBusqueda.Caption = "Busqueda de " & provCli ' primero cambia el titulo de ventanda de busqueda
    If Not frmPCBusqueda.Inicio( _
                CodProvCli, _
                nombre, _
                codg, _
                Numg) Then
        'Si fue cancelada la busqueda, sale no mas
        grd.SetFocus
        Exit Sub
    End If
    
    'Cambia la forma de cursor
    MensajeStatus MSG_PREPARA, vbHourglass
    Select Case Me.tag 'jeaa 24/09/04 asignacion de grupo a los items
        Case "CUENTA_PROV", "CUENTA_CLI"
           'Compone la cadena de SQL
            sql = "SELECT CodProvCli, Nombre" & _
            ", CodCuentaContable, CodCuentaContable2 " & _
            "FROM vwPCProvCli "
        Case "PCGRUPOS_PROV", "PCGRUPOS_CLI"
            'Compone la cadena de SQL
            sql = "SELECT CodProvCli, Nombre" & _
            ", codgrupo1, codgrupo2, codgrupo3 ,codgrupo4 " & _
            "FROM vwPCProvCli "
        End Select
    ' si Busca Proveedor o cliente mediante bandera de prov
    If Len(cond) > 0 Then cond = cond & "AND "
    If bandprov Then
        cond = cond & "(BandProveedor = 1) "
    Else
        cond = cond & "(BandCliente = 1) "
    End If
    
    'CodProveedor/cliente
    If Len(CodProvCli) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(codProvCli LIKE '" & CodProvCli & comodin & "') "
    End If

    
    'Nombre
    If Len(nombre) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(Nombre LIKE '" & nombre & comodin & "') "
    End If

    'Grupo
    If Len(codg) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodGrupo" & Numg & " = '" & codg & "') "
    End If

    If Len(cond) > 0 Then sql = sql & " WHERE " & cond
    sql = sql & " ORDER BY CodProvCli "
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)

    With grd
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If Not rs.EOF Then .LoadArray MiGetRows(rs)
        ConfigCols
        .Redraw = flexRDBuffered
        .SetFocus
    End With

    MensajeStatus
    Exit Sub
ErrTrap:
    grd.Redraw = flexRDBuffered
    MensajeStatus
    DispErr
    grd.SetFocus
    Exit Sub
End Sub

Private Sub ConfigCols()
    Dim s As String, I As Long, j As Integer
    With grd
    Select Case Me.tag
        Case "CUENTA_PROV", "CUENTA_CLI"
            s = "^#|<Código|<Descripción|<Cuenta Contable1|<Cuenta Contable2"
        Case "ITEM_VIDAUTIL"
            s = "^#|<Código|<Cód.Alterno|<Descripción|>VidaUtil"
        Case "CUENTA"
            s = "^#|<Código|<Cód.Alterno|<Descripción|<Cta Activo|<Cta Costo|<Cta Venta|<Cta Dep. Gasto|<Cta Dep Acum|<Cta Revalo|<Cta Dep. Revalo"
        Case "CUENTA_LOCAL"     'jeaa 21/01/04
            s = "^#|<Código|<Descripción|^CodLocal|^Sucursal"
            grd.ColHidden(3) = True
        Case "ITEM_FAMILIA"
            s = "^#|<Código|<Descripción|>IdRecuperado|>CodFamilia|<Familia|<Id"
        Case "ITEM_AFGRUPOS" 'jeaa 24/09/04 asignacion de grupo a los items
            s = "^#|<Código|<Cód.Alterno|<Descripción"
            For j = 1 To 5
                s = s & "|<" & gobjMain.EmpresaActual.GNOpcion.EtiqAFGrupo(j)
            Next j
        Case "PCGRUPOS_PROV", "PCGRUPOS_CLI" 'jeaa 24/09/04 asignacion de grupo a los items
            s = "^#|<Código|<Descripción"
            For j = 1 To 4 'Cambiado AUC 03/10/2005 Antes 3
                s = s & "|>" & gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupo(j)
            Next j
        'jeaa 13/04/2005
        Case "FRACCION"
            s = "^#|<Código|<Cód.Alterno|<Descripción|>Venta por Fraccion"
        'jeaa 15/09/2005
        Case "AREA"
            s = "^#|<Código|<Cód.Alterno|<Descripción|>Venta por Area"
        'jeaa 26/12/2005
        Case "VENTA"
            s = "^#|<Código|<Cód.Alterno|<Descripción|>Venta"
        Case "COSTOUI"
            s = "^#|<Código|<Cód.Alterno|<Descripción|>Costo Ultimo Ingreso|>Fecha Grabado"
        Case "SRI"
            s = "^#|<Transid|<Fecha Trans|<Código|<Num.Trans|<#Doc Ref.|<Descripción|<Autorizacion SRI|>Fecha Caducidad"
            grd.ColHidden(1) = True
        Case "MINMAX"
            s = "^#|<idInventario|<Código|<Cód.Alterno|<Descripción|<idbodega|<CodBodega|>Existencia|>Existencia Mínima|>Existencia Máxima"
        Case "IVEXIST"
            s = "^#|<idInventario|<Código|<Cód.Alterno|<Descripción|<idbodega|<CodBodega|>Existencia"
        Case "CUENTA_PRESUP"     'jeaa 21/01/04
            s = "^#|<Código|<Descripción|>Presupuesto"
        Case "DIASREPO"
            s = "^#|<Código|<Cód.Alterno|<Descripción|>Dias"
        
        End Select
        .FormatString = s
        GNPoneNumFila grd, False
        AjustarAutoSize grd, -1, -1, 4000
        AsignarTituloAColKey grd
    
        'Columnas modificables (Longitud maxima)
        Select Case Me.tag
        Case "ITEM_VIDAUTIL"
            .ColData(.ColIndex("VidaUtil")) = 5
        Case "CUENTA"
            .ColData(.ColIndex("Cta Activo")) = 20
            .ColData(.ColIndex("Cta Costo")) = 20
            .ColData(.ColIndex("Cta Venta")) = 20
            .ColData(.ColIndex("Cta Dep. Gasto")) = 20
            .ColData(.ColIndex("Cta Dep Acum")) = 20
            .ColData(.ColIndex("Cta Revalo")) = 20
            .ColData(.ColIndex("Cta Dep. Revalo")) = 20
            
            CargarCuentas
        
        Case "ITEM_AFGRUPOS" 'jeaa 24/09/04 asignacion de grupo a los items
            .ColData(AFGRUPO1) = 20
            .ColData(AFGRUPO2) = 20
            .ColData(AFGRUPO3) = 20
            .ColData(AFGRUPO4) = 20
            .ColData(AFGRUPO5) = 20
            .ColWidth(AFGRUPO1) = 2000
            .ColWidth(AFGRUPO2) = 2000
            .ColWidth(AFGRUPO3) = 2000
            .ColWidth(AFGRUPO4) = 2000
            .ColWidth(AFGRUPO5) = 2000
            If Not CargarAFGRUPOs Then
                grd.Rows = 1
                Exit Sub
            End If
        Case "PCGRUPOS_PROV", "PCGRUPOS_CLI"
            .TextMatrix(0, .ColIndex("Descripción")) = "Nombre"  ' cambio caption de Descripcion a Nombre
            .ColData(PCGRUPO1) = 20
            .ColData(PCGRUPO2) = 20
            .ColData(PCGRUPO3) = 20
            .ColData(PCGRUPO4) = 20
            .ColWidth(PCGRUPO1) = 2000
            .ColWidth(PCGRUPO2) = 2000
            .ColWidth(PCGRUPO3) = 2000
            .ColData(PCGRUPO4) = 2000
            If Not CargarPCGrupos Then
                grd.Rows = 1
                Exit Sub
            End If
        Case "FRACCION"
            .ColData(.ColIndex("Venta por Fraccion")) = 5
            .ColDataType(4) = flexDTBoolean
        Case "AREA" 'jeaa 15/09/2005
            .ColData(.ColIndex("Venta por Area")) = 5
            .ColDataType(4) = flexDTBoolean
        Case "VENTA" 'jeaa 26/12/2005
            .ColData(.ColIndex("Venta")) = 5
            .ColDataType(4) = flexDTBoolean
        Case "COSTOUI"
            .ColData(.ColIndex("Costo Ultimo Ingreso")) = 5
            .ColFormat(.ColIndex("Costo Ultimo Ingreso")) = "##.0000"
        Case "SRI"
            .ColData(.ColIndex("Autorizacion SRI")) = 10
            .ColData(.ColIndex("Fecha Caducidad")) = 10
        Case "MINMAX"
            .ColData(.ColIndex("Existencia")) = 20
            .ColData(.ColIndex("Existencia Mínima")) = 20
            .ColData(.ColIndex("Existencia Máxima")) = 20
            .ColHidden(.ColIndex("idInventario")) = True
            .ColHidden(.ColIndex("idBodega")) = True
        Case "IVEXIST"
            .ColData(.ColIndex("Existencia")) = 20
            .ColData(.ColIndex("Existencia Mínima")) = 20
            .ColData(.ColIndex("Existencia Máxima")) = 20
            .ColHidden(.ColIndex("idInventario")) = True
            .ColHidden(.ColIndex("idBodega")) = True
        Case "DIASREPO"
            .ColData(.ColIndex("DIAS")) = 5
            
        End Select
        'Columnas No modificables
        
        Select Case Me.tag
        Case "MINMAX"
            For I = 0 To .ColIndex("Existencia")
                .ColData(I) = -1
            Next I
            
            'Color de fondo
            If .Rows > .FixedRows Then
                .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("Existencia")) = .BackColorFrozen
            End If
        Case "EXIST"
            For I = 0 To .ColIndex("Existencia")
                .ColData(I) = -1
            Next I
            
            'Color de fondo
            If .Rows > .FixedRows Then
                .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("Existencia")) = .BackColorFrozen
            End If
        
        
        Case Else
            For I = 0 To .ColIndex("Descripción")
                .ColData(I) = -1
            Next I
            
            'Color de fondo
            If .Rows > .FixedRows Then
                .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("Descripción")) = .BackColorFrozen
            End If
        End Select
    End With
End Sub

Private Sub CargarCuentas()
    Dim s As String
    With grd
        s = gobjMain.EmpresaActual.ListaCTCuentaParaFlexGrid(0)
        s = Right$(s, Len(s) - 1)
        .ColComboList(.ColIndex("Cta Activo")) = s
        .ColComboList(.ColIndex("Cta Costo")) = s
        .ColComboList(.ColIndex("Cta Venta")) = s
        .ColComboList(.ColIndex("Cta Dep. Gasto")) = s
        .ColComboList(.ColIndex("Cta Dep Acum")) = s
        .ColComboList(.ColIndex("Cta Revalo")) = s
        .ColComboList(.ColIndex("Cta Dep. Revalo")) = s
        
        
    End With
End Sub

Private Sub Asignar()
    Select Case Me.tag
        Case "ITEM_VIDAUTIL":        AsignarVidaUtil
        Case "CUENTA":     AsignarCuenta
        Case "CUENTA_LOCAL": AsignarLocal
        Case "ITEM_FAMILIA": AsignarFamilia
        Case "ITEM_AFGRUPOS": AsignarAFGRUPOs   'jeaa 24/09/04 asignacion de grupo a los items
        Case "PCGRUPOS_PROV", "PCGRUPOS_CLI": AsignarPCGrupos  'jeaa 24/09/04 asignacion de grupo a los prov-cli
        Case "FRACCION":  AsignarFraccion
        Case "AREA":  AsignarArea  'jeaa 15/09/2005
        Case "VENTA":  AsignarVenta 'jeaa 26/12/2005
        Case "COSTOUI":        AsignarCostoUI
        Case "SRI":        AsignarSRI
        Case "MINMAX":     AsignarMinMax
        Case "IVEXIST":     AsignarIVExist
        Case "CUENTA_PRESUP": AsignarPresupuesto
        Case "DIASREPO":        AsignarDias
    End Select
End Sub

Private Sub AsignarIVA()
    Dim s As String, v As Single
    Dim I As Long
    
    s = InputBox("Ingrese el valor de IVA (%)", "Asignar un valor", "15")
    If IsNumeric(s) Then
        v = CSng(s)
    Else
        MsgBox "Debe ingresar un valor numérico. (ejm. 15 para 15%)", vbInformation
        grd.SetFocus
        Exit Sub
    End If
    
    With grd
        For I = .FixedRows To .Rows - 1
            .TextMatrix(I, .ColIndex("VidaUtil")) = v
        Next I
    End With
End Sub

Private Sub AsignarCostoUI()
    Dim s As String, v As Single
    Dim I As Long
    
    s = InputBox("Ingrese el valor de Costo ULtimo Ingreso ", "Asignar un valor", "15")
    If IsNumeric(s) Then
        v = CSng(s)
    Else
        MsgBox "Debe ingresar un valor numérico. (ejm. 1.1544) ", vbInformation
        grd.SetFocus
        Exit Sub
    End If
    
    With grd
        For I = .FixedRows To .Rows - 1
            .TextMatrix(I, .ColIndex("Costo Ultimo Ingreso")) = v
        Next I
    End With
End Sub


Private Sub AsignarCuenta()
    Dim activo As String, costo As String, venta As String
    Dim DepGasto As String, CtaDepAcum As String, CtaRevalo As String, CtaDepRevalo As String
    Dim cta As String, nomcta  As String
    Dim I As Long, s As String, j As Integer
    
    With grd
        'Obtiene cuentas de la fila actual
        For j = 1 To 7
            Select Case j
            Case 1
                cta = .TextMatrix(.Row, .ColIndex("Cta Activo"))
                nomcta = "Cta Activo"
            Case 2
                cta = .TextMatrix(.Row, .ColIndex("Cta Costo"))
                nomcta = "Cta Costo"
            Case 3
                cta = .TextMatrix(.Row, .ColIndex("Cta Venta"))
                nomcta = "Cta Venta"
            Case 4
                cta = .TextMatrix(.Row, .ColIndex("Cta Dep. Gasto"))
                nomcta = "Cta Dep. Gasto"
            Case 5
                cta = .TextMatrix(.Row, .ColIndex("Cta Dep Acum"))
                nomcta = "Cta Dep Acum"
            Case 6
                cta = .TextMatrix(.Row, .ColIndex("Cta Revalo"))
                nomcta = "Cta Revalo"
            Case 7
                cta = .TextMatrix(.Row, .ColIndex("Cta Dep. Revalo"))
                nomcta = "Cta Dep. Revalo"
            End Select
        
        'Confirma las cuentas
        s = "Está seguro que desea asignar los siguientes códigos " & _
            "en todos los ítems que están visualizados?" & vbCr & vbCr & nomcta & ": " & cta
            If MsgBox(s, vbQuestion + vbYesNo) = vbYes Then
                For I = .FixedRows To .Rows - 1
                    Select Case j
                    Case 1
                        .TextMatrix(1, .ColIndex("Cta Activo")) = cta
                    Case 2
                         .TextMatrix(I, .ColIndex("Cta Costo")) = cta
                    Case 3
                         .TextMatrix(I, .ColIndex("Cta Venta")) = cta
                    Case 4
                         .TextMatrix(I, .ColIndex("Cta Dep. Gasto")) = cta
                    Case 5
                         .TextMatrix(I, .ColIndex("Cta Dep Acum")) = cta
                    Case 6
                         .TextMatrix(I, .ColIndex("Cta Revalo")) = cta
                    Case 7
                         .TextMatrix(I, .ColIndex("Cta Dep. Revalo")) = cta
                    End Select
                Next I
            
            
            
'''            "    Cta de Activo:  " & activo & vbCr & _
'''            "    Cta de Costo:   " & costo & vbCr & _
'''            "    Cta de Venta:   " & venta

            .SetFocus
'            Exit Sub
        End If
        
        'Copia a todas las filas los mismos códigos de cuenta
''''        For i = .FixedRows To .Rows - 1
''''            .TextMatrix(i, .ColIndex("Cta Activo")) = activo
''''            .TextMatrix(i, .ColIndex("Cta Costo")) = costo
''''            .TextMatrix(i, .ColIndex("Cta Venta")) = venta
''''        Next i
        Next j
    End With
End Sub


Private Sub Grabar()
    Dim I As Long, iv As AFinventario, cod As String
    Dim gnc As GNComprobante
    Dim sql As String, rs As Recordset
    Dim IdBodega As Integer
    On Error GoTo ErrTrap
    
    'Confirmación
    If MsgBox("Está seguro que desea grabar?", vbQuestion + vbYesNo) <> vbYes Then
        grd.SetFocus
        Exit Sub
    End If
    
    'Deshabilita los botónes y menus
    Habilitar False
    mCancelado = False
    
    With grd
        If Me.tag = "IVEXIST" Then
            sql = "select idbodega from ivbodega where codbodega='" & .TextMatrix(2, .ColIndex("CodBodega")) & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            IdBodega = rs.Fields("idbodega")
        End If
        
            
        prg1.min = 0
        prg1.max = 1
        If .Rows > .FixedRows Then prg1.max = .Rows - 1
        For I = .FixedRows To .Rows - 1
            'Si es que se canceló el proceso
            If mCancelado Then GoTo salida
        
            prg1.value = I
            cod = .TextMatrix(I, .ColIndex("Código"))
            MensajeStatus I & " de " & .Rows - .FixedRows, vbHourglass
            DoEvents
            
            'Recupera el objeto de Inventario
            If Me.tag <> "SRI" Then
                Set iv = gobjMain.EmpresaActual.RecuperaAFInventario(cod)
            End If
            
            Select Case Me.tag
            Case "ITEM_VIDAUTIL"
                If iv.VidaUtil <> .ValueMatrix(I, .ColIndex("VidaUtil")) Then
                    iv.VidaUtil = .ValueMatrix(I, .ColIndex("VidaUtil"))
                End If
            Case "CUENTA"
                If iv.CodCuentaActivo <> .TextMatrix(I, .ColIndex("Cta Activo")) Then
                    iv.CodCuentaActivo = .TextMatrix(I, .ColIndex("Cta Activo"))
                End If
                If iv.CodCuentaCosto <> .TextMatrix(I, .ColIndex("Cta Costo")) Then
                    iv.CodCuentaCosto = .TextMatrix(I, .ColIndex("Cta Costo"))
                End If
                If iv.CodCuentaVenta <> .TextMatrix(I, .ColIndex("Cta Venta")) Then
                    iv.CodCuentaVenta = .TextMatrix(I, .ColIndex("Cta Venta"))
                End If
                If iv.CodCuentaDepreGasto <> .TextMatrix(I, .ColIndex("Cta Dep. Gasto")) Then
                    iv.CodCuentaDepreGasto = .TextMatrix(I, .ColIndex("Cta Dep. Gasto"))
                End If
                If iv.CodCuentaDepreAcumulada <> .TextMatrix(I, .ColIndex("Cta Dep Acum")) Then
                    iv.CodCuentaDepreAcumulada = .TextMatrix(I, .ColIndex("Cta Dep Acum"))
                End If
                If iv.CodCuentaRevaloriza <> .TextMatrix(I, .ColIndex("Cta Revalo")) Then
                    iv.CodCuentaRevaloriza = .TextMatrix(I, .ColIndex("Cta Revalo"))
                End If
                If iv.CodCuentaDepRevaloriza <> .TextMatrix(I, .ColIndex("Cta Dep. Revalo")) Then
                    iv.CodCuentaDepRevaloriza = .TextMatrix(I, .ColIndex("Cta Dep. Revalo"))
                End If
            
            
            Case "ITEM_AFGRUPOS"    'jeaa 24/09/04 asignacion de grupo a los items
                If iv.CodGrupo(1) <> .TextMatrix(I, AFGRUPO1) Then
                    iv.CodGrupo(1) = .TextMatrix(I, AFGRUPO1)
                End If
                If iv.CodGrupo(2) <> .TextMatrix(I, AFGRUPO2) Then
                    iv.CodGrupo(2) = .TextMatrix(I, AFGRUPO2)
                End If
                If iv.CodGrupo(3) <> .TextMatrix(I, AFGRUPO3) Then
                    iv.CodGrupo(3) = .TextMatrix(I, AFGRUPO3)
                End If
                If iv.CodGrupo(4) <> .TextMatrix(I, AFGRUPO4) Then
                    iv.CodGrupo(4) = .TextMatrix(I, AFGRUPO4)
                End If
                If iv.CodGrupo(5) <> .TextMatrix(I, AFGRUPO5) Then
                    iv.CodGrupo(5) = .TextMatrix(I, AFGRUPO5)
                End If
            
            Case "COSTOUI"
                If iv.CostoUltimoIngreso <> .ValueMatrix(I, .ColIndex("Costo Ultimo Ingreso")) Then
                    'iv.costoultimoingreso = .ValueMatrix(i, .ColIndex("Costo Ultimo Ingreso"))
                'End If
                    sql = " UPDATE AFInventario "
                    sql = sql & " SET CostoUltimoIngreso= " & .ValueMatrix(I, .ColIndex("Costo Ultimo Ingreso"))
                    sql = sql & " where codinventario='" & iv.CodInventario & "'"
                    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                End If
            Case "IVEXIST"
                    sql = " insert ivexist (IdInventario,idbodega,exist) values ("
                    sql = sql & .ValueMatrix(I, .ColIndex("idInventario")) & ","
                    sql = sql & IdBodega & ","
                    sql = sql & " 0) "

                    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
         
            End Select
            If Me.tag <> "COSTOUI" And Me.tag <> "SRI" Then
                iv.Grabar
            End If
        Next I
    End With
    
salida:
    MensajeStatus
    Set iv = Nothing
    Habilitar True
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    GoTo salida
    Exit Sub
End Sub

Private Sub GrabarPC()
    Dim I As Long, pc As PCProvCli, cod As String
    On Error GoTo ErrTrap
    
    'Confirmación
    If MsgBox("Está seguro que desea grabar?", vbQuestion + vbYesNo) <> vbYes Then
        grd.SetFocus
        Exit Sub
    End If
    
    'Deshabilita los botónes y menus
    Habilitar False
    mCancelado = False
    
    With grd
        prg1.min = 0
        prg1.max = 1
        If .Rows > .FixedRows Then prg1.max = .Rows - 1
        For I = .FixedRows To .Rows - 1
            'Si es que se canceló el proceso
            If mCancelado Then GoTo salida
        
            prg1.value = I
            cod = .TextMatrix(I, .ColIndex("Código"))
            MensajeStatus I & " de " & .Rows - .FixedRows, vbHourglass
            DoEvents
            
            'Recupera el objeto de Inventario
            Set pc = gobjMain.EmpresaActual.RecuperaPCProvCli(cod)
            Select Case Me.tag
                Case "CUENTA_PROV", "CUENTA_CLI"
                    If pc.CodCuentaContable <> .TextMatrix(I, .ColIndex("Cta Contable1")) Then
                        pc.CodCuentaContable = .TextMatrix(I, .ColIndex("Cta Contable1"))
                    End If
                    
                    If pc.CodCuentaContable2 <> .TextMatrix(I, .ColIndex("Cta Contable2")) Then
                        pc.CodCuentaContable2 = .TextMatrix(I, .ColIndex("Cta Contable2"))
                    End If
                Case "PCGRUPOS_PROV ", "PCGRUPOS_CLI"
                    If pc.CodGrupo1 <> .TextMatrix(I, PCGRUPO1) Then
                        pc.CodGrupo1 = .TextMatrix(I, PCGRUPO1)
                    End If
                    If pc.CodGrupo2 <> .TextMatrix(I, PCGRUPO2) Then
                        pc.CodGrupo2 = .TextMatrix(I, PCGRUPO2)
                    End If
                    If pc.CodGrupo3 <> .TextMatrix(I, PCGRUPO3) Then
                        pc.CodGrupo3 = .TextMatrix(I, PCGRUPO3)
                    End If
                    'AUC 03/10/2005
                  If pc.CodGrupo4 <> .TextMatrix(I, PCGRUPO4) Then
                        pc.CodGrupo4 = .TextMatrix(I, PCGRUPO4)
                    End If
            
            End Select
            pc.Grabar
        Next I
    End With
    
salida:
    MensajeStatus
    Set pc = Nothing
    Habilitar True
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    GoTo salida
    Exit Sub
End Sub

Private Sub Habilitar(ByVal v As Boolean)
    mProcesando = Not v
    
    tlb1.Buttons("Buscar").Enabled = v
    tlb1.Buttons("Asignar").Enabled = v
    tlb1.Buttons("Grabar").Enabled = v
'    tlb1.Buttons("Imprimir").Enabled = v
    tlb1.Buttons("Imprimir").Enabled = False        '*** MAKOTO PENDIENTE Por ahora
    
    If v Then
        tlb1.Buttons("Cerrar").Caption = "Cerrar"
    Else
        tlb1.Buttons("Cerrar").Caption = "Cancelar"
    End If
    
    frmMain.mnuFile.Enabled = v
    frmMain.mnuHerramienta.Enabled = v
    frmMain.mnuTransferir.Enabled = v
    frmMain.mnuCerrarTodas.Enabled = v
    
    prg1.value = prg1.min
End Sub

Private Sub Imprimir()

End Sub

Private Sub Cerrar()
    If mProcesando Then
        'Si está procesando, pregunta si quere abandonarlo
        If MsgBox("Desea abandonar el proceso?", vbQuestion + vbYesNo) = vbYes Then
            mCancelado = True
        End If
        
        Exit Sub
    Else
        Unload Me
    End If
End Sub

Private Function CargarLocales() As Boolean
    Dim s As String
    On Error GoTo ErrTrap
        With grd
            CargarLocales = True
            s = gobjMain.EmpresaActual.ListaCTLocalesParaFlexGrid(0)
            s = Right$(s, Len(s) - 1)
            .ColComboList(.ColIndex("Sucursal")) = s
        End With
        Exit Function
ErrTrap:
        MsgBox "No se han definido Locales", vbInformation
        CargarLocales = False
    Exit Function
End Function

Private Sub BuscarCuenta()
    Dim sql As String, cad As String, rs As Recordset
    Static codcuenta As String, nombre As String, codg As String, Numg As Integer
    Dim cond As String, comodin As String
    Dim CtCuenta As String
    On Error GoTo ErrTrap
    #If DAOLIB Then
        comodin = "*"
    #Else
        comodin = "%"
    #End If
    
    'Abre la pantalla de búsqueda
    If Me.tag = "CUENTA_LOCAL" Then
        frmPCBusqueda.Caption = "Busqueda de Cuentas Contables"
        frmPCBusqueda.cboGrupo.Visible = False
        frmPCBusqueda.fcbGrupo.Visible = False
        frmPCBusqueda.Label5.Visible = False
        frmPCBusqueda.cmdAceptar.Top = frmPCBusqueda.fcbGrupo.Top
        frmPCBusqueda.cmdCancelar.Top = frmPCBusqueda.cmdAceptar.Top
        frmPCBusqueda.Height = 2000
    ElseIf Me.tag = "CUENTA_PRESUP" Then
        frmPCBusqueda.Caption = "Busqueda de Cuentas Contables"
        frmPCBusqueda.cboGrupo.Visible = False
        frmPCBusqueda.fcbGrupo.Visible = False
        frmPCBusqueda.Label5.Visible = False
        frmPCBusqueda.cmdAceptar.Top = frmPCBusqueda.fcbGrupo.Top
        frmPCBusqueda.cmdCancelar.Top = frmPCBusqueda.cmdAceptar.Top
        frmPCBusqueda.Height = 2000
    
    End If
    If Not frmPCBusqueda.Inicio( _
                codcuenta, _
                nombre, _
                codg, _
                Numg) Then
        'Si fue cancelada la busqueda, sale no mas
        grd.SetFocus
        Exit Sub
    End If
    
'CodProveedor/cliente
    If Len(codcuenta) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(codCuenta LIKE '" & codcuenta & comodin & "') "
    End If

    
    'Nombre
    If Len(nombre) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(NombreCuenta LIKE '" & comodin & nombre & comodin & "') "
    End If

    
    sql = "SELECT CodCuenta, NombreCuenta "
    If Me.tag = "CUENTA_LOCAL" Then
        sql = sql & " ,codlocal,nombre FROM ctlocal right join ctcuenta "
        sql = sql & " on ctlocal.idlocal = ctcuenta.idlocal "
    ElseIf Me.tag = "CUENTA_PRESUP" Then
        sql = sql & " , isnull(valPresupuesto,0) as valPresupuesto  FROM  ctcuenta "
    End If

    If Len(cond) > 0 Then
        sql = sql & " where bandtotal=0 and " & cond
    Else
        sql = sql & " where bandtotal=0"
    End If
    sql = sql & " ORDER BY CodCuenta"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
   
    With grd
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If Not rs.EOF Then .LoadArray MiGetRows(rs)
        ConfigCols
        .Redraw = flexRDBuffered
        .SetFocus
    End With
    Set rs = Nothing
    Exit Sub
ErrTrap:
    grd.Redraw = flexRDBuffered
    MensajeStatus
    DispErr
    grd.SetFocus
    Exit Sub

End Sub

Private Sub AsignarLocal()
    Dim Sucursal As String, codlocal As String
    Dim I As Long, s As String
    
    With grd
        'Obtiene cuentas de la fila actual
        codlocal = .TextMatrix(.Row, .ColIndex("CodLocal"))
        Sucursal = .TextMatrix(.Row, .ColIndex("Sucursal"))
        'Confirma las cuentas
        s = "Está seguro que desea asignar las siguientes sucursales " & _
            "en todos las cuentas que están visualizados?" & vbCr & vbCr & _
            "    Sucursal :  " & Sucursal & vbCr
        If MsgBox(s, vbQuestion + vbYesNo) <> vbYes Then
            .SetFocus
            Exit Sub
        End If
        
        'Copia a todas las filas los mismos códigos de cuenta
        For I = .Row To .Rows - 1
            .TextMatrix(I, .ColIndex("Sucursal")) = Sucursal
            .TextMatrix(I, .ColIndex("CodLocal")) = codlocal
        Next I
    End With
    End Sub

Private Sub GrabarLocal()
    Dim I As Long, ct As CtCuenta, cod As String
    On Error GoTo ErrTrap
    
    'Confirmación
    If MsgBox("Está seguro que desea grabar?", vbQuestion + vbYesNo) <> vbYes Then
        grd.SetFocus
        Exit Sub
    End If
    
    'Deshabilita los botónes y menus
    Habilitar False
    mCancelado = False
    
    With grd
        prg1.min = 0
        prg1.max = 1
        If .Rows > .FixedRows Then prg1.max = .Rows - 1
        For I = .FixedRows To .Rows - 1
            'Si es que se canceló el proceso
            If mCancelado Then GoTo salida
        
            prg1.value = I
            cod = .TextMatrix(I, .ColIndex("Código"))
            MensajeStatus I & " de " & .Rows - .FixedRows, vbHourglass
            DoEvents
            
            'Recupera el objeto de Inventario
            Set ct = gobjMain.EmpresaActual.RecuperaCTCuenta(cod)
            ct.codlocal = .TextMatrix(I, .ColIndex("CodLocal"))
            ct.Grabar
        Next I
    End With
    
salida:
    MensajeStatus
    Set ct = Nothing
    Habilitar True
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    GoTo salida
    Exit Sub
End Sub

Private Sub RecuperaCodLocal(ByVal nombre As String, I As Long)
On Error GoTo ErrTrap
    Dim codlocal As String, rs As Recordset, sql As String
    sql = "SELECT codlocal FROM ctlocal where nombre = '" & nombre & "'"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not rs.EOF Then
        grd.TextMatrix(I, grd.ColIndex("CodLocal")) = rs.Fields("codlocal")
    End If
        Exit Sub
ErrTrap:
        MsgBox "No se han definido Locales.", vbInformation
    Exit Sub
End Sub

Private Function CargarFamilias() As Boolean
    Dim s As String
    On Error GoTo ErrTrap
    With grd
        CargarFamilias = True
        s = gobjMain.EmpresaActual.ListaIVItemFamiliaParaFlex
        s = Right$(s, Len(s) - 1)
        .ColComboList(.ColIndex("Familia")) = s
    End With
    Exit Function
ErrTrap:
        MsgBox "No se han definido Familias.", vbInformation
        CargarFamilias = False
    Exit Function
End Function

Private Sub RecuperaCodFamilia(ByVal nombre As String, I As Long)
    Dim codlocal As String, rs As Recordset, sql As String
    sql = "SELECT idinventario,codinventario FROM AFInventario where descripcion = '" & nombre & "'"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not rs.EOF Then
        grd.TextMatrix(I, grd.ColIndex("CodFamilia")) = rs.Fields("CodInventario")
        grd.TextMatrix(I, grd.ColIndex("Id")) = rs.Fields("Idinventario")
    End If
End Sub

Private Sub AsignarFamilia()
    Dim familia As String, codfamilia As String
    Dim I As Long, s As String
    
    With grd
        'Obtiene cuentas de la fila actual
        codfamilia = .TextMatrix(.Row, .ColIndex("CodFamilia"))
        familia = .TextMatrix(.Row, .ColIndex("Familia"))
        'Confirma las cuentas
        s = "Está seguro que desea asignar las siguientes Familias " & _
            "en todos las items que están visualizados?" & vbCr & vbCr & _
            "    Sucursal :  " & familia & vbCr
        If MsgBox(s, vbQuestion + vbYesNo) <> vbYes Then
            .SetFocus
            Exit Sub
        End If
        
        'Copia a todas las filas los mismos códigos de cuenta
        For I = .Row To .Rows - 1
            .TextMatrix(I, .ColIndex("Familia")) = familia
            .TextMatrix(I, .ColIndex("CodFamilia")) = codfamilia
            RecuperaCodFamilia familia, I
        Next I
    End With
    End Sub


'''Private Sub EliminaFila(ByVal idfamilia As Long, ByVal codhijo As String)
'''    Dim msg As String, r As Long, i As Long
'''    Dim mobjIV As AFinventario, mobjIVF As IVFamiliaDetalle, rs As Recordset
'''    On Error GoTo errtrap
'''        'recupero el item
'''        Set mobjIV = gobjMain.EmpresaActual.RecuperaAFInventario(idfamilia)
'''        'boy recorrer la coleccion
'''        For i = 1 To mobjIV.NumFamiliaDetalle
'''                'recupero un item de la coleccio
'''                Set mobjIVF = mobjIV.RecuperaDetalleFamilia(i)
'''                'comparo si es igual al parametro
'''                If mobjIVF.CodInventario = codhijo Then
'''                    'elimino de la coleccion
'''                    mobjIV.RemoveDetalleFamilia (i)
'''                    'grabo el item
'''                    mobjIV.Grabar
'''                    Set mobjIVF = Nothing
'''                    Set mobjIV = Nothing
'''                    Exit Sub
'''                End If
'''        Next i
'''    Set mobjIVF = Nothing
'''    Set mobjIV = Nothing
'''    Exit Sub
'''errtrap:
'''    DispErr
'''    Exit Sub
'''End Sub


Private Sub BuscarItemFamilia()
    Static coditem As String, CodAlt As String, _
           Desc As String, _
           codg As String, Numg As Integer, bandIVA As Boolean, bandFraccion As Boolean
    Dim codg1 As String, codg2 As String, codg3 As String, codg4 As String, codg5 As String
    Dim sql As String, cond As String, rs As Recordset, comodin As String
    On Error GoTo ErrTrap
   
    #If DAOLIB Then
        comodin = "*"
    #Else
        comodin = "%"
    #End If
'    comodin = "%"
    'Abre la pantalla de búsqueda
    frmAFBusqueda.chkIVA.Visible = False
    frmAFBusqueda.cmdAceptar.Top = frmAFBusqueda.chkIVA.Top
    frmAFBusqueda.cmdCancelar.Top = frmAFBusqueda.chkIVA.Top
    frmAFBusqueda.Height = 3000
    If Not frmAFBusqueda.Inicio( _
                coditem, _
                CodAlt, _
                Desc, _
                codg1, codg2, codg3, codg4, codg5, _
                Numg, _
                bandIVA, _
                Me.tag) Then
        'Si fue cancelada la busqueda, sale no mas
        grd.SetFocus
        Exit Sub
    End If
    'Cambia la forma de cursor
    MensajeStatus MSG_PREPARA, vbHourglass
    sql = " SELECT vwConsIVFamilia.CodInventario, vwConsIVFamilia.Descripcion, vwConsIVFamilia.IdFamilia, vwConsIVFamilia.CodFamilia, vwConsIVFamilia.Familia "
    sql = sql & "FROM vwConsIVFamilia INNER JOIN vwAFInventarioRecuperar ON vwConsIVFamilia.IdInventario=vwAFInventarioRecuperar.IdInventario"
    'que no sea padre de familia o receta
    sql = sql & " WHERE vwAFInventarioRecuperar.tipo= '0' "
    ' que no pertenesca a ninguna otra familia ni receta
'    sql = sql & "  and vwConsIVFamilia.idInventario not in(select idinventario from ivmateria)"
    'CodInventario
    If Len(coditem) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(vwConsIVFamilia.CodInventario LIKE '" & coditem & comodin & "') "
    End If
    
    'CodAlterno
    If Len(CodAlt) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "((vwAFInventarioRecuperar.CodAlterno1 LIKE '" & CodAlt & comodin & "') " & _
                      "OR (vwAFInventarioRecuperar.CodAlterno2 LIKE '" & CodAlt & comodin & "')) "
    End If
    
    'Descripcion
    If Len(Desc) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(vwConsIVFamilia.Descripcion LIKE '" & Desc & comodin & "') "
    End If
    
    'Grupo
    If Len(codg) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(vwAFInventarioRecuperar.CodGrupo" & Numg & " = '" & codg & "') "
    End If
    If Len(cond) > 0 Then sql = sql & " and  " & cond
    sql = sql & " ORDER BY vwConsIVFamilia.CodInventario "
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    With grd
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If Not rs.EOF Then .LoadArray MiGetRows(rs)
        ConfigCols
        .Redraw = flexRDBuffered
        .SetFocus
    End With
    
    MensajeStatus
    Exit Sub
ErrTrap:
    grd.Redraw = flexRDBuffered
    MensajeStatus
    DispErr
    grd.SetFocus
    Exit Sub
End Sub

'jeaa 24/09/04 asignacion de grupo a los items
Private Function CargarAFGRUPOs() As Boolean
    Dim s As String
    On Error GoTo ErrTrap
    With grd
        CargarAFGRUPOs = True
        'fcbGrupoDesde.SetData gobjMain.EmpresaActual.ListaAFGRUPO(Numg, False, False)
        
        s = gobjMain.EmpresaActual.ListaAFGrupoParaFlexGrid(1)
        If Len(s) <> 0 Then
            s = Right$(s, Len(s) - 1)
            .ColComboList(AFGRUPO1) = s
        End If
        s = gobjMain.EmpresaActual.ListaAFGrupoParaFlexGrid(2)
        If Len(s) <> 0 Then
            s = Right$(s, Len(s) - 1)
            .ColComboList(AFGRUPO2) = s
        End If
        s = gobjMain.EmpresaActual.ListaAFGrupoParaFlexGrid(3)
        If Len(s) <> 0 Then
            s = Right$(s, Len(s) - 1)
            .ColComboList(AFGRUPO3) = s
        End If
        s = gobjMain.EmpresaActual.ListaAFGrupoParaFlexGrid(4)
        If Len(s) <> 0 Then
            s = Right$(s, Len(s) - 1)
            .ColComboList(AFGRUPO4) = s
        End If
        s = gobjMain.EmpresaActual.ListaAFGrupoParaFlexGrid(5)
        If Len(s) <> 0 Then
            s = Right$(s, Len(s) - 1)
            .ColComboList(AFGRUPO5) = s
        End If
    
    End With
    Exit Function
ErrTrap:
        MsgBox "No se han definido AFGRUPOs", vbInformation
        CargarAFGRUPOs = False
    Exit Function
End Function

'jeaa 24/09/04 asignacion de grupo a los items
Private Sub AsignarAFGRUPOs()
    Dim ValorGrupo As String
    Dim I As Long, s As String, j As Integer, grupo As String
    
    With grd
        For j = 1 To 5
            'Obtiene cuentas de la fila actual
            Select Case j
                Case 1
                    ValorGrupo = .TextMatrix(.Row, AFGRUPO1)
                Case 2
                    ValorGrupo = .TextMatrix(.Row, AFGRUPO2)
                Case 3
                    ValorGrupo = .TextMatrix(.Row, AFGRUPO3)
                Case 4
                    ValorGrupo = .TextMatrix(.Row, AFGRUPO4)
                Case 5
                    ValorGrupo = .TextMatrix(.Row, AFGRUPO5)
            
            End Select
            grupo = gobjMain.EmpresaActual.GNOpcion.EtiqAFGrupo(j) & ":  " & ValorGrupo & vbCr
            'Confirma las GRUPOS
            s = "Está seguro que desea asignar los siguientes códigos " & _
                "en todos los ítems que están visualizados?" & vbCr & vbCr & grupo
            If MsgBox(s, vbQuestion + vbYesNo) = vbYes Then
                'Copia a todas las filas los mismos códigos de cuenta
                For I = .FixedRows To .Rows - 1
                    Select Case j
                        Case 1
                            .TextMatrix(I, AFGRUPO1) = ValorGrupo
                        Case 2
                            .TextMatrix(I, AFGRUPO2) = ValorGrupo
                        Case 3
                            .TextMatrix(I, AFGRUPO3) = ValorGrupo
                        Case 4
                            .TextMatrix(I, AFGRUPO4) = ValorGrupo
                        Case 5
                            .TextMatrix(I, AFGRUPO5) = ValorGrupo
                        
                        End Select
                Next I
            End If
        Next j
    End With
End Sub

'jeaa 24/09/04 asignacion de grupo a los prov_cli
Private Function CargarPCGrupos() As Boolean
    Dim s As Variant
    On Error GoTo ErrTrap
    With grd
        CargarPCGrupos = True
        s = gobjMain.EmpresaActual.ListaPCGrupoParaFlexGrid(1)
        s = Right$(s, Len(s) - 1)
        .ColComboList(PCGRUPO1) = s
        s = gobjMain.EmpresaActual.ListaPCGrupoParaFlexGrid(2)
        s = Right$(s, Len(s) - 1)
        .ColComboList(PCGRUPO2) = s
        s = gobjMain.EmpresaActual.ListaPCGrupoParaFlexGrid(3)
        s = Right$(s, Len(s) - 1)
        .ColComboList(PCGRUPO3) = s
        'AUC 03/10/2005
        s = gobjMain.EmpresaActual.ListaPCGrupoParaFlexGrid(4)
        If Len(s) > 0 Then
            s = Right$(s, Len(s) - 1)
            .ColComboList(PCGRUPO4) = s
        End If
    End With
    Exit Function
ErrTrap:
        MsgBox "No se han definido PCGrupos", vbInformation
        CargarPCGrupos = False
    Exit Function
End Function

'jeaa 24/09/04 asignacion de grupo a los items
Private Sub AsignarPCGrupos()
    Dim ValorGrupo As String
    Dim I As Long, s As String, j As Integer, grupo As String
    With grd
        For j = 1 To 4 'AUC  03/10/2005 antes 3
           'Obtiene cuentas de la fila actual
            Select Case j
                Case 1
                    ValorGrupo = .TextMatrix(.Row, PCGRUPO1)
                Case 2
                    ValorGrupo = .TextMatrix(.Row, PCGRUPO2)
                Case 3
                    ValorGrupo = .TextMatrix(.Row, PCGRUPO3)
                Case 4 'AUC 03/10/2005
                    ValorGrupo = .TextMatrix(.Row, PCGRUPO4)
           End Select
            grupo = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(j) & ":  " & ValorGrupo & vbCr
            'Confirma las GRUPOS
            s = "Está seguro que desea asignar los siguientes códigos " & _
                "en todos los Prov/cli que están visualizados?" & vbCr & vbCr & grupo
            If MsgBox(s, vbQuestion + vbYesNo) = vbYes Then
                'Copia a todas las filas los mismos códigos de cuenta
                For I = .FixedRows To .Rows - 1
                    Select Case j
                        Case 1
                            .TextMatrix(I, PCGRUPO1) = ValorGrupo
                        Case 2
                            .TextMatrix(I, PCGRUPO2) = ValorGrupo
                        Case 3
                            .TextMatrix(I, PCGRUPO3) = ValorGrupo
                         Case 4 'auc 03/10/2005
                           .TextMatrix(I, PCGRUPO4) = ValorGrupo
                       End Select
                Next I
            End If
        Next j
    End With
End Sub
'jeaa 13/04/2005
Private Sub AsignarFraccion()
    Dim band As Boolean, I As Integer
    With grd
        'Obtiene Bandera de la fila actual
        band = .TextMatrix(.Row, .ColIndex("Venta por Fraccion"))
        For I = .Row To .Rows - 1
            .TextMatrix(I, .ColIndex("Venta por Fraccion")) = IIf(band, vbChecked, vbUnchecked)
        Next I
    End With
End Sub

'jeaa 13/04/2005
Private Sub AsignarArea()
    Dim band As Boolean, I As Integer
    With grd
        'Obtiene Bandera de la fila actual
        band = .TextMatrix(.Row, .ColIndex("Venta por Area"))
        For I = .Row To .Rows - 1
            .TextMatrix(I, .ColIndex("Venta por Area")) = IIf(band, vbChecked, vbUnchecked)
        Next I
    End With
End Sub

'jeaa 26/12/2005
Private Sub AsignarVenta()
    Dim band As Boolean, I As Integer
    With grd
        'Obtiene Bandera de la fila actual
        band = .TextMatrix(.Row, .ColIndex("Venta"))
        For I = .Row To .Rows - 1
            .TextMatrix(I, .ColIndex("Venta")) = IIf(band, vbChecked, vbUnchecked)
        Next I
    End With
End Sub


Private Sub BuscarSRI()
    Static CodTrans As String, desde As Long, hasta As Long
    Dim sql As String, cond As String, rs As Recordset, comodin As String
    On Error GoTo ErrTrap
    'If Me.tag <> "CUENTA" Then Exit Sub
    
    #If DAOLIB Then
        comodin = "*"
    #Else
        comodin = "%"
    #End If
'    comodin = "%"
    'Abre la pantalla de búsqueda
    If Not frmAFBusqueda.InicioTrans( _
                CodTrans, _
                desde, hasta) Then
        'Si fue cancelada la busqueda, sale no mas
        grd.SetFocus
        Exit Sub
    End If
    
    'Cambia la forma de cursor
    MensajeStatus MSG_PREPARA, vbHourglass
    
    'Compone la cadena de SQL
    sql = "SELECT transid, fechatrans, gnc.codtrans, numtrans , NumDocRef, gnt.descripcion, AutorizacionSRI, FechaCaducidadSRI "
    sql = sql & " FROM gncomprobante gnc inner join gntrans gnt on gnc.codtrans=gnt.codtrans"
    'CodInventario
    If Len(CodTrans) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(gnc.Codtrans = '" & CodTrans & "') "
    End If
    
    'CodAlterno
    If desde <> 0 And hasta <> 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "numtrans between " & desde & " and " & hasta
    End If
    
    
    
    
    If Len(cond) > 0 Then sql = sql & " WHERE " & cond
    sql = sql & " ORDER BY Numtrans "
    
    
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    With grd
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If Not rs.EOF Then .LoadArray MiGetRows(rs)
        ConfigCols
        .Redraw = flexRDBuffered
        .SetFocus
    End With
    
    MensajeStatus
    Exit Sub
ErrTrap:
    grd.Redraw = flexRDBuffered
    MensajeStatus
    DispErr
    grd.SetFocus
    Exit Sub
End Sub


'jeaa 24/09/04 asignacion de grupo a los items
Private Sub AsignarSRI()
    Dim NumSRI As String, fecha As String
    Dim I As Long, s As String
    
    With grd
        'Obtiene cuentas de la fila actual
        NumSRI = .TextMatrix(.Row, .ColIndex("Autorizacion SRI"))
        fecha = .TextMatrix(.Row, .ColIndex("Fecha Caducidad"))
        
        'Confirma las cuentas
        s = "Está seguro que desea asignar los siguientes códigos " & _
            "en todos los ítems que están visualizados?" & vbCr & vbCr & _
            "    Autorizacion SRI:  " & NumSRI & vbCr & _
            "    Fecha Caducidad:   " & fecha
        If MsgBox(s, vbQuestion + vbYesNo) <> vbYes Then
            .SetFocus
            Exit Sub
        End If
        
        'Copia a todas las filas los mismos códigos de cuenta
        For I = .FixedRows To .Rows - 1
            .TextMatrix(I, .ColIndex("Autorizacion SRI")) = NumSRI
            .TextMatrix(I, .ColIndex("Fecha Caducidad")) = fecha
        Next I
    End With

End Sub

Private Sub BuscarCostoUltimo()
    Static CodTrans As String, desde As Long, hasta As Long
    Dim sql As String, cond As String, rs As Recordset, comodin As String
    Dim objcond As Condicion
    Dim Recargo As String
    On Error GoTo ErrTrap
    'If Me.tag <> "CUENTA" Then Exit Sub
    
    #If DAOLIB Then
        comodin = "*"
    #Else
        comodin = "%"
    #End If
    Set objcond = gobjMain.objCondicion
'    comodin = "%"
    'Abre la pantalla de búsqueda
    CodTrans = "UltimoCosto"
    If Not frmB_CxTrans.InicioCxProveedor(objcond, Recargo, _
                CodTrans) Then
        'Si fue cancelada la busqueda, sale no mas
        grd.SetFocus
        Exit Sub
    End If
    
    'Cambia la forma de cursor
    MensajeStatus MSG_PREPARA, vbHourglass
    
    'Compone la cadena de SQL
    sql = "SELECT   IV.codInventario,IV.codAlterno1,IV.Descripcion,ivk.CostoRealTotal/ivk.cantidad as CostoRealTotal,gnc.FechaGrabado "
    sql = sql & " FROM AFInventario IV inner join IVKardex ivk inner join gncomprobante  gnc on gnc.Transid=ivk.transid"
    sql = sql & "  ON iv.idinventario = ivk.idinventario"
    'CodInventario
    If Len(objcond.CodTrans) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & " gnc.Codtrans in (" & objcond.CodTrans & ") "
    End If
    
    'CodAlterno
    If objcond.fecha1 <> 0 And objcond.fecha2 <> 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "gnc.fechatrans between '" & objcond.fecha1 & "' And  '" & objcond.fecha2 & "'"
    End If
    If Len(cond) > 0 Then sql = sql & " WHERE " & cond
    sql = sql & " AND ivk.CostoRealTotal <> 0 ORDER BY gnc.FechaGrabado "
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    
    With grd
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If Not rs.EOF Then .LoadArray MiGetRows(rs)
        ConfigCols
        .Redraw = flexRDBuffered
        .SetFocus
    End With
    
    MensajeStatus
    Exit Sub
ErrTrap:
    grd.Redraw = flexRDBuffered
    MensajeStatus
    DispErr
    grd.SetFocus
    Exit Sub
End Sub


Private Sub BuscarMINMAX()
Static coditem As String, CodAlt As String, _
           Desc As String, _
           codg As String, Numg As Integer, bandIVA As Boolean, bandFraccion As Boolean
    Dim codg1 As String, codg2 As String, codg3 As String, codg4 As String, codg5 As String
    Static CodTrans As String, desde As Long, hasta As Long
    Dim sql As String, cond As String, rs As Recordset, comodin As String
    On Error GoTo ErrTrap
    
    #If DAOLIB Then
        comodin = "*"
    #Else
        comodin = "%"
    #End If
'    comodin = "%"
    'Abre la pantalla de búsqueda
    If Not frmAFBusqueda.Inicio( _
                coditem, _
                CodAlt, _
                Desc, _
                codg1, codg2, codg3, codg4, codg5, _
                Numg, _
                bandIVA, _
                Me.tag) Then
      'if not frmAFBusqueda.InicioTrans (
        'Si fue cancelada la busqueda, sale no mas
        grd.SetFocus
        Exit Sub
    End If
    
    'Cambia la forma de cursor
    MensajeStatus MSG_PREPARA, vbHourglass
    
    'Compone la cadena de SQL
    sql = "SELECT"
    sql = sql & " IVI.IdInventario , CodInventario, CodAlterno1, IVI.Descripcion, "
    sql = sql & " ive.IdBodega, CodBodega, exist, existmin, existmax"
    sql = sql & " from    AFGRUPO5"
    sql = sql & " RIGHT JOIN (AFGRUPO4"
    sql = sql & " RIGHT JOIN (AFGRUPO3"
    sql = sql & " RIGHT JOIN (AFGRUPO2"
    sql = sql & " RIGHT JOIN (AFGRUPO1"
    sql = sql & " RIGHT JOIN AFInventario ivi"
    sql = sql & " inner join ivexist ive"
    sql = sql & " inner join ivbodega ivb"
    sql = sql & " on ive.idbodega=ivb.idbodega"
    sql = sql & " on ivi.idinventario=ive.idinventario"
    sql = sql & " ON AFGRUPO1.IdGrupo1 = IVI.IdGrupo1)"
    sql = sql & " ON AFGRUPO2.IdGrupo2 = IVI.IdGrupo2)"
    sql = sql & " ON AFGRUPO3.IdGrupo3 = IVI.IdGrupo3)"
    sql = sql & " ON AFGRUPO4.IdGrupo4 = IVI.IdGrupo4)"
    sql = sql & " ON AFGRUPO5.IdGrupo5 = IVI.IdGrupo5"

        'CodInventario
    If Len(coditem) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodInventario LIKE '" & coditem & comodin & "') "
    End If
    
    'CodAlterno
    If Len(CodAlt) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "((CodAlterno1 LIKE '" & CodAlt & comodin & "') " & _
                      "OR (CodAlterno2 LIKE '" & CodAlt & comodin & "')) "
    End If
    
    'Descripcion
    If Len(Desc) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(ivi.Descripcion LIKE '" & Desc & comodin & "') "
    End If
    
    
    
    If Len(codg1) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodGrupo1" & " = '" & codg1 & "') "
    End If
    
    If Len(codg2) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodGrupo2" & " = '" & codg2 & "') "
    End If
    
    If Len(codg3) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodGrupo3" & " = '" & codg3 & "') "
    End If
    
    If Len(codg4) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodGrupo4" & " = '" & codg4 & "') "
    End If
    
    If Len(codg5) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodGrupo5" & " = '" & codg5 & "') "
    End If
    
   
    If Len(cond) > 0 Then sql = sql & " WHERE " & cond
    sql = sql & " ORDER BY CodInventario, codbodega "
    
    
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    With grd
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If Not rs.EOF Then .LoadArray MiGetRows(rs)
        ConfigCols
        .Redraw = flexRDBuffered
        .SetFocus
    End With
    
    MensajeStatus
    Exit Sub
ErrTrap:
    grd.Redraw = flexRDBuffered
    MensajeStatus
    DispErr
    grd.SetFocus
    Exit Sub
End Sub

Private Sub AsignarMinMax()
    Dim min As Integer, max As Integer
    Dim I As Long, s As String
    
    With grd
        min = .ValueMatrix(.Row, .ColIndex("Existencia Mínima"))
        max = .ValueMatrix(.Row, .ColIndex("Existencia Máxima"))
        
        'Confirma las cuentas
        s = "Está seguro que desea asignar los siguientes códigos " & _
            "en todos los ítems que están visualizados?" & vbCr & vbCr & _
            "    Existencia Mínima:  " & min & vbCr & _
            "    Existencia Máxima:   " & max & vbCr
        If MsgBox(s, vbQuestion + vbYesNo) <> vbYes Then
            .SetFocus
            Exit Sub
        End If
        
        'Copia a todas las filas los mismos códigos de cuenta
        For I = .FixedRows To .Rows - 1
            .TextMatrix(I, .ColIndex("Existencia Mínima")) = min
            .TextMatrix(I, .ColIndex("Existencia Máxima")) = max
        Next I
    End With
End Sub


Private Sub BuscarIvExist()
Static coditem As String, CodAlt As String, _
           Desc As String, _
           codg As String, Numg As Integer, bandIVA As Boolean, bandFraccion As Boolean
    Dim codg1 As String, codg2 As String, codg3 As String, codg4 As String, codg5 As String
    Dim CodBodega As String
    Static CodTrans As String, desde As Long, hasta As Long
    Dim sql As String, cond As String, rs As Recordset, comodin As String
    On Error GoTo ErrTrap
    
    #If DAOLIB Then
        comodin = "*"
    #Else
        comodin = "%"
    #End If
'    comodin = "%"
    'Abre la pantalla de búsqueda
    If Not frmAFBusqueda.Inicio( _
                coditem, _
                CodAlt, _
                Desc, _
                codg1, codg2, codg3, codg4, codg5, _
                Numg, _
                bandIVA, _
                Me.tag, CodBodega) Then
      'if not frmAFBusqueda.InicioTrans (
        'Si fue cancelada la busqueda, sale no mas
        grd.SetFocus
        Exit Sub
    End If
    
    'Cambia la forma de cursor
    MensajeStatus MSG_PREPARA, vbHourglass
    
    'Compone la cadena de SQL
    sql = "SELECT"
    sql = sql & " IVI.IdInventario , CodInventario, CodAlterno1, IVI.Descripcion, 0,'" & CodBodega & "',0 "
    sql = sql & " from    AFGRUPO5"
    sql = sql & " RIGHT JOIN (AFGRUPO4"
    sql = sql & " RIGHT JOIN (AFGRUPO3"
    sql = sql & " RIGHT JOIN (AFGRUPO2"
    sql = sql & " RIGHT JOIN (AFGRUPO1"
    sql = sql & " RIGHT JOIN AFInventario ivi"
    sql = sql & " ON AFGRUPO1.IdGrupo1 = IVI.IdGrupo1)"
    sql = sql & " ON AFGRUPO2.IdGrupo2 = IVI.IdGrupo2)"
    sql = sql & " ON AFGRUPO3.IdGrupo3 = IVI.IdGrupo3)"
    sql = sql & " ON AFGRUPO4.IdGrupo4 = IVI.IdGrupo4)"
    sql = sql & " ON AFGRUPO5.IdGrupo5 = IVI.IdGrupo5"

If Len(cond) > 0 Then cond = cond & "AND "
cond = cond & " ivi.idinventario not in"
cond = cond & " ( select idinventario from ivexist ive inner join ivbodega ivb on ive.idbodega=ivb.idbodega"
cond = cond & " where codbodega = '" & CodBodega & "')"
        'CodInventario
    If Len(coditem) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodInventario LIKE '" & coditem & comodin & "') "
    End If
    
    'CodAlterno
    If Len(CodAlt) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "((CodAlterno1 LIKE '" & CodAlt & comodin & "') " & _
                      "OR (CodAlterno2 LIKE '" & CodAlt & comodin & "')) "
    End If
    
    'Descripcion
    If Len(Desc) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(ivi.Descripcion LIKE '" & Desc & comodin & "') "
    End If
    
    
    
    If Len(codg1) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodGrupo1" & " = '" & codg1 & "') "
    End If
    
    If Len(codg2) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodGrupo2" & " = '" & codg2 & "') "
    End If
    
    If Len(codg3) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodGrupo3" & " = '" & codg3 & "') "
    End If
    
    If Len(codg4) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodGrupo4" & " = '" & codg4 & "') "
    End If
    
    If Len(codg5) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodGrupo5" & " = '" & codg5 & "') "
    End If
    
   
    If Len(cond) > 0 Then sql = sql & " WHERE " & cond
    sql = sql & " ORDER BY CodInventario "
    
    
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    With grd
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If Not rs.EOF Then .LoadArray MiGetRows(rs)
        ConfigCols
        .Redraw = flexRDBuffered
        .SetFocus
    End With
    
    MensajeStatus
    Exit Sub
ErrTrap:
    grd.Redraw = flexRDBuffered
    MensajeStatus
    DispErr
    grd.SetFocus
    Exit Sub
End Sub


Private Sub AsignarIVExist()
    Dim bod As String
    Dim I As Long, s As String
    
    With grd
        bod = .TextMatrix(.Row, .ColIndex("CodBodega"))
        
        'Confirma las cuentas
        s = "Está seguro que desea asignar los siguientes códigos " & _
            "en todos los ítems que están visualizados?" & vbCr & vbCr & _
            "    Bodega:  " & bod
        If MsgBox(s, vbQuestion + vbYesNo) <> vbYes Then
            .SetFocus
            Exit Sub
        End If
        
        'Copia a todas las filas los mismos códigos de cuenta
        For I = .FixedRows To .Rows - 1
            .TextMatrix(I, .ColIndex("CodBodega")) = bod
        Next I
    End With
End Sub


Private Sub AsignarPresupuesto()
    Dim Presupuesto As Currency
    Dim I As Long, s As String
    
    With grd
        'Obtiene cuentas de la fila actual
        Presupuesto = .ValueMatrix(.Row, .ColIndex("Presupuesto"))
        'Confirma las cuentas
        s = "Está seguro que desea asignar el presupuesto " & _
            "en todos las cuentas que están visualizados?" & vbCr & vbCr & _
            "    Presupuesto :  " & Presupuesto & vbCr
        If MsgBox(s, vbQuestion + vbYesNo) <> vbYes Then
            .SetFocus
            Exit Sub
        End If
        
        'Copia a todas las filas los mismos códigos de cuenta
        For I = .Row To .Rows - 1
            .TextMatrix(I, .ColIndex("Presupuesto")) = Presupuesto
        Next I
    End With
End Sub


Private Sub GrabarPresupuesto()
    Dim I As Long, ct As CtCuenta, cod As String
    On Error GoTo ErrTrap
    
    'Confirmación
    If MsgBox("Está seguro que desea grabar?", vbQuestion + vbYesNo) <> vbYes Then
        grd.SetFocus
        Exit Sub
    End If
    
    'Deshabilita los botónes y menus
    Habilitar False
    mCancelado = False
    
    With grd
        prg1.min = 0
        prg1.max = 1
        If .Rows > .FixedRows Then prg1.max = .Rows - 1
        For I = .FixedRows To .Rows - 1
            'Si es que se canceló el proceso
            If mCancelado Then GoTo salida
        
            prg1.value = I
            cod = .TextMatrix(I, .ColIndex("Código"))
            MensajeStatus I & " de " & .Rows - .FixedRows, vbHourglass
            DoEvents
            
            'Recupera el objeto de Inventario
            Set ct = gobjMain.EmpresaActual.RecuperaCTCuenta(cod)
            ct.ValPresupuesto = .ValueMatrix(I, .ColIndex("Presupuesto"))
            ct.Grabar
        Next I
    End With
    
salida:
    MensajeStatus
    Set ct = Nothing
    Habilitar True
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    GoTo salida
    Exit Sub
End Sub

Private Sub AsignarDias()
    Dim s As String, v As Single
    Dim I As Long
    
    s = InputBox("Ingrese el valor de Dias", "Asignar un valor", "15")
    If IsNumeric(s) Then
        v = CSng(s)
    Else
        MsgBox "Debe ingresar un valor numérico. (ejm. 15)", vbInformation
        grd.SetFocus
        Exit Sub
    End If
    
    With grd
        For I = .FixedRows To .Rows - 1
            .TextMatrix(I, .ColIndex("Dias")) = v
        Next I
    End With
End Sub


Private Sub AsignarVidaUtil()
    Dim s As String, v As Single
    Dim I As Long
    
    s = InputBox("Ingrese el valor de VidaUtil ", "Asignar un valor", "15")
    If IsNumeric(s) Then
        v = CSng(s)
    Else
        MsgBox "Debe ingresar un valor numérico. (ejm. 15 para 15)", vbInformation
        grd.SetFocus
        Exit Sub
    End If
    
    With grd
        For I = .FixedRows To .Rows - 1
            .TextMatrix(I, .ColIndex("VidaUtil")) = v
        Next I
    End With
End Sub


