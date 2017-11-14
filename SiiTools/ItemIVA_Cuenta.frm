VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl1.ocx"
Begin VB.Form frmItemIVA_Cuenta 
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
            Picture         =   "ItemIVA_Cuenta.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ItemIVA_Cuenta.frx":0114
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ItemIVA_Cuenta.frx":0568
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ItemIVA_Cuenta.frx":067C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ItemIVA_Cuenta.frx":0790
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
Attribute VB_Name = "frmItemIVA_Cuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mProcesando As Boolean
Private mCancelado As Boolean
Private mcolItemsSelec As Collection      'Coleccion de items
'jeaa 24/09/04 asignacion de grupo a los items
'IVINVENTARIO
Const IVGRUPO1 = 4
Const IVGRUPO2 = 5
Const IVGRUPO3 = 6
Const IVGRUPO4 = 7
Const IVGRUPO5 = 8
Const IVGRUPO6 = 9
'PC_PROV_CLI
Const PCGRUPO1 = 3
Const PCGRUPO2 = 4
Const PCGRUPO3 = 5
Const PCGRUPO4 = 6 'Agregado AUC 03/10/2005
'descxItem
Const DESC1 = 4
Const DESC2 = 5
Const DESC3 = 6
Const DESC4 = 7
Const DESC5 = 8
Private CodTransRel As String
Private WithEvents mobjGNComp As GNComprobante
Attribute mobjGNComp.VB_VarHelpID = -1

'Private mobjItem As IVinventario

Public Sub Inicio(ByVal tag As String)
    Dim i As Integer
    On Error GoTo ErrTrap
    
    Me.tag = tag            'Guarda en la propiedad Tag para distinguir después
    Me.Show
    Me.ZOrder
    
    Select Case Me.tag
    Case "IVA"
        Me.Caption = "Actualización de IVA de ítems"
    Case "CUENTA"
        Me.Caption = "Actualización de cuentas contables de ítems"
    Case "CUENTA_PROV"
        Me.Caption = "Actualización de cuentas contables de Proveedores"
    Case "CUENTA_CLI"
        Me.Caption = "Actualización de cuentas contables de Cliente"
    Case "CUENTA_EMP"
        Me.Caption = "Actualización de cuentas contables de Empleados"
    Case "CUENTA_LOCAL"
        Me.Caption = "Actualizacion de cuentas contables de Locales"   'jeaa-21/01/04
    Case "ITEM_FAMILIA"
        Me.Caption = "Asiganción de Items a una Familia"
    Case "ITEM_IVGRUPOS"    'jeaa 24/09/04 asignacion de grupo a los items
        Me.Caption = "Asiganción de Grupos a Items "
    Case "PCGRUPOS_PROV"   'jeaa 24/09/04 asignacion de grupo a los prov
        Me.Caption = "Asiganción de Grupo a Proveedores"
    Case "PCGRUPOS_CLI"    'jeaa 24/09/04 asignacion de grupo a los cli
        Me.Caption = "Asiganción de Grupo a Clientes"
    Case "PCGRUPOS_GAR"    'jeaa 24/09/04 asignacion de grupo a los cli
        Me.Caption = "Asiganción de Grupo a Garantes"
    Case "PCGRUPOS_EMP"
        Me.Caption = "Asiganción de Grupo a Empleados"
    Case "FRACCION"    'jeaa 13/04/05 asignacion de bandera fraccion
        Me.Caption = "Asiganción Bandera para Venta en Fracción"
    Case "AREA"    'jeaa 15/09/05 asignacion de bandera AREA
        Me.Caption = "Asiganción Bandera para Venta por Areas"
    Case "VENTA"    'jeaa 26/12/05 asignacion de bandera VENTA
        Me.Caption = "Asiganción Bandera para Venta "
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
        Me.Caption = "Actualizacion de Tiempo de Reposicion "   'jeaa-09/01/2009
    Case "PORDESC"
        Me.Caption = "Actualización de % Descuento"
    Case "PORCOMI"
        Me.Caption = "Actualización de % Comisión"
    Case "IVEXISTNEG"
        Me.Caption = "Llenado de Items en Tabla Existencias "
    Case "COSTOREF"
        Me.Caption = "Actualización del Costo Referencial"
    Case "PROVINCIAS", "PROVINCIAS_PROV", "PROVINCIASEMP"
        Me.Caption = "Llenado Provincias Cantones Parroquia "
    Case "ARANCEL"
        Me.Caption = "Actualización de Arancel "
    Case "AFEXIST"
        Me.Caption = "Llenado de Activos en Tabla Existencias "
    Case "CUENTASC"
        Me.Caption = "Relacionador Ctas. Contables de la Super Compañías"
    Case "PCCLIRUC"
        Me.Caption = "Verifica CI/RUC de Clientes"
    Case "CUENTAFE"
        Me.Caption = "Relacionador Ctas. Contables de la Flujo Efectivo"
    Case "FORMAPAGOSRI"
        Me.Caption = "Actualización de Forma de Pago SRI "
    Case "DINARDAP"
        Me.Caption = "Actualización de datos de clientes para DINARDAP "
    Case "DIVNOMEMP"
        Me.Caption = "Dividir el Nombre a Empleados"
    Case "CUENTA101"
        Me.Caption = "Relacionador Ctas. Contables para Formulario 101"
    Case "COMPROB"
        Me.Caption = "Relacionador de Comprobantes"
    Case "EMPDOC"
        Me.Caption = "Asignar Empleado - Entrega Documentos"
    Case "PCCLIRUCFCEL"
        Me.Caption = "Verifica CI/RUC, Nombre de Clientes para Facturacion Electrónica"
    Case "ITEMFCEL"
        Me.Caption = "Verifica Código, Descripción de Items para Facturacion Electrónica"
    Case "PCEMAIL"
        Me.Caption = "Verifica dirección de correo electrónico"
    Case "LECTURAS"
        Me.Caption = "Ingreso de Lecturas"
    Case "PCPARR"
        Me.Caption = "Verifica Parroquias Incorrectas"
    Case "FORMASRI"
        Me.Caption = "Asigna Forma Cobro SRI"
    Case "ESCOPIARTC"
        Me.Caption = "Asigna valor de Copia/Original en Retencion de Clientes"
    Case "PCCLI_VENDEDOR"
        Me.Caption = "Asiganción de Vendedores as Clientes"
    Case "PCAGENCIA"
        Me.Caption = "Creacion de Agencia a Clientes"
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
            Case "IVA", "CUENTA", "ITEM_IVGRUPOS", "FRACCION", "AREA", "VENTA", "PORDESC", "PORCOMI", "ARANCEL": Buscar
            Case "CUENTA_PROV", "PCGRUPOS_PROV": BuscarPC True
            Case "CUENTA_CLI", "PCGRUPOS_CLI", "PCCLI_VENDEDOR", "PCAGENCIA": BuscarPC False
            Case "PCGRUPOS_GAR": BuscarPCgar
            Case "CUENTA_EMP", "PCGRUPOS_EMP", "DIVNOMEMP": BuscarEmp
            Case "CUENTA_LOCAL": BuscarCuenta 'jeaa 21/01/04
            Case "ITEM_FAMILIA": BuscarItemFamilia 'jeaa 21/01/04
            Case "COSTOUI": BuscarCostoUltimo
            Case "SRI": BuscarSRI
            Case "MINMAX": BuscarMINMAX
            Case "IVEXIST": BuscarIvExist
            Case "CUENTA_PRESUP", "CUENTASC", "CUENTAFE", "CUENTA101": BuscarCuenta
            Case "DIASREPO": Buscar 'jeaa 21/01/04
            Case "IVEXISTNEG": BuscarIvExist
            Case "COSTOREF": Buscar
            Case "PROVINCIAS": BuscarProvincias False
            Case "PROVINCIAS_PROV": BuscarProvincias True
            Case "PROVINCIASEMP": BuscarProvinciasEmp
            Case "VENDE": BuscarVendedor
            Case "AFEXIST": BuscarAFExist
            Case "PCCLIRUC": BuscarPC False
            Case "FORMAPAGOSRI": BuscarFormaPagoSRI
            Case "DINARDAP": BuscarDINARDAP
            Case "COMPROB"": BuscarComprobanteRelacionado"
            Case "EMPDOC": BuscarEmpleado
            Case "PCCLIRUCFCEL": BuscarPC False
            Case "ITEMFCEL": Buscar
            Case "PCEMAIL": BuscarPC False
            Case "LECTURAS": BuscarPC False
            Case "PCPARR": BuscarDINARDAPParr
            Case "FORMASRI": BuscarFormaCobro
            Case "ESCOPIARTC": BuscarTransRTC
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
        Case "DINARDAP", "PCPARR"
            Select Case col
                Case grd.ColIndex("Provincia")
                    RecuperaProvincia grd.TextMatrix(Row, grd.ColIndex("Provincia")), Row
                Case grd.ColIndex("Canton")
                    RecuperaCanton grd.TextMatrix(Row, grd.ColIndex("Canton")), Row
                Case grd.ColIndex("Parroquia")
                    RecuperaParroquia grd.TextMatrix(Row, grd.ColIndex("Parroquia")), Row
                End Select
            
                
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
    Select Case Me.tag
        Case "PROVINCIAS", "PROVINCIAS_PROV", "PROVINCIASEMP"
            Select Case col
                Case grd.ColIndex("Provincia")
                    grd.ColComboList(grd.ColIndex("Provincia")) = gobjMain.EmpresaActual.ListaPCProvinciaParaFlex(True)
                Case grd.ColIndex("Canton")
                    grd.ColComboList(grd.ColIndex("Canton")) = gobjMain.EmpresaActual.ListaPCCantonxProvinciaFlex(True, grd.TextMatrix(Row, grd.ColIndex("Provincia")))
                Case grd.ColIndex("Parroquia")
                    grd.ColComboList(grd.ColIndex("Parroquia")) = gobjMain.EmpresaActual.ListaPCParroquiaxCantonFlex(True, grd.TextMatrix(Row, grd.ColIndex("CantonNue")))
            End Select
        
        Case "DINARDAP", "PCPARR"
            Select Case col
                Case grd.ColIndex("Provincia")
                    grd.ColComboList(grd.ColIndex("Provincia")) = gobjMain.EmpresaActual.ListaPCProvinciaParaFlex(True)
                Case grd.ColIndex("Canton")
                    grd.ColComboList(grd.ColIndex("Canton")) = gobjMain.EmpresaActual.ListaPCCantonxProvinciaFlex(True, grd.TextMatrix(Row, grd.ColIndex("Provincia")))
                Case grd.ColIndex("Parroquia")
                    grd.ColComboList(grd.ColIndex("Parroquia")) = gobjMain.EmpresaActual.ListaPCParroquiaxCantonFlex(True, grd.TextMatrix(Row, grd.ColIndex("Canton")))
                Case grd.ColIndex("Tipo Sujeto")
                    grd.ColComboList(grd.ColIndex("Tipo Sujeto")) = gobjMain.EmpresaActual.ListaTipoSujetoFlex()
            
            End Select
        Case "PCCLI_VENDEDOR"
            grd.ColComboList(grd.ColIndex("Cod Vendedor New")) = gobjMain.EmpresaActual.ListaFCVendedorParaFlex
        Case "PCAGENCIA"

    End Select
    
End Sub

Private Sub grd_BeforeSort(ByVal col As Long, Order As Integer)
    'Impide mientras está procesando
    If mProcesando Then Order = flexSortNone
End Sub

Private Sub grd_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
    With grd
        Select Case Me.tag
        Case "IVA", "PORDESC", "PORCOMI"
            If Not IsNumeric(.EditText) Then
                MsgBox "Debe ingresar un valor numérico.", vbInformation
                .SetFocus
                Cancel = True
            End If
        Case "CUENTA", "CUENTA_PROV", "CUENTA_CLI", "CUENTA_LOCAL", "ITEM_FAMILIA", "ITEM_RECETA", "MINMAX", "CUENTA_EMP", "ARANCEL", "CUENTASC", "CUENTAFE", "CUENTA101" 'jeaa 21/01/04
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
        Case "COSTOREF"
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
        Case "IVEXISTNEG" 'jeaa 21/01/04
            If Right$(.EditText, 1) = "*" Then
                MsgBox "No puede seleccionar una cuenta de mayor.", vbInformation
                .SetFocus
                Cancel = True
            End If
        Case "FORMAPAGOSRI"
            If col <> 8 Then
                If Not IsNumeric(.EditText) Then
                    MsgBox "Debe ingresar un valor numérico.", vbInformation
                    .SetFocus
                    Cancel = True
                End If
            End If
        Case "FECHAINICIAL"
            If Not IsDate(.EditText) Then
                MsgBox "Debe ingresar una fecha", vbInformation
                .SetFocus
                Cancel = True
            End If

        End Select
    End With
End Sub

Private Sub tlb1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.key
    Case "Buscar":
        Select Case Me.tag
            Case "IVA", "CUENTA", "ITEM_IVGRUPOS", "FRACCION", "AREA", "VENTA", "PORDESC", "PORCOMI", "ARANCEL": Buscar
            Case "CUENTA_PROV", "PCGRUPOS_PROV": BuscarPC True
            Case "CUENTA_CLI", "PCGRUPOS_CLI", "PCCLI_VENDEDOR", "PCAGENCIA": BuscarPC False
            Case "PCGRUPOS_GAR":: BuscarPCgar
            Case "CUENTA_EMP", "PCGRUPOS_EMP", "DIVNOMEMP": BuscarEmp
            Case "CUENTA_LOCAL": BuscarCuenta 'jeaa 21/01/04
            Case "ITEM_FAMILIA": BuscarItemFamilia
            Case "COSTOUI": BuscarCostoUltimo
            Case "SRI": BuscarSRI
            Case "MINMAX": BuscarMINMAX
            Case "IVEXIST": BuscarIvExist
            Case "CUENTA_PRESUP", "CUENTASC", "CUENTAFE", "CUENTA101": BuscarCuenta 'jeaa 21/01/04
            Case "DIASREPO": Buscar
            Case "IVEXISTNEG": BuscarIvExistNegativa
            Case "COSTOREF": Buscar
            Case "PROVINCIAS": BuscarProvincias False
            Case "PROVINCIAS_PROV": BuscarProvincias True
            Case "PROVINCIASEMP": BuscarProvinciasEmp
            Case "VENDE": BuscarVendedor
            Case "AFEXIST": BuscarAFExist
            Case "PCCLIRUC": BuscarPC False
            Case "FORMAPAGOSRI": BuscarFormaPagoSRI
            Case "DINARDAP": BuscarDINARDAP
            Case "COMPROB": BuscarComprobanteRelacionado
            Case "PCCLIRUCFCEL": BuscarPC False
            Case "ITEMFCEL": Buscar
            Case "PCEMAIL": BuscarPC False
            Case "LECTURAS": BuscarPC False
            Case "PCPARR": BuscarDINARDAPParr
            Case "FECHAINICIAL": BuscarFechaInicial
            Case "FORMASRI": BuscarFormaCobro
            Case "ESCOPIARTC": BuscarTransRTC
        End Select
    Case "Asignar":     Asignar
    Case "Grabar":
        Select Case Me.tag
            Case "IVA", "CUENTA", "ITEM_IVGRUPOS", "FRACCION", "AREA", "VENTA", "COSTOUI", "MINMAX", "IVEXIST", "PORDESC", "PORCOMI", "IVEXISTNEG", "ARANCEL", "AFEXIST": Grabar
            Case "CUENTA_PROV", "CUENTA_CLI", "PCGRUPOS_PROV", "PCGRUPOS_CLI", "PROVINCIAS", "PROVINCIAS_PROV", "PCCLI_VENDEDOR": GrabarPC
            
            Case "PCGRUPOS_GAR": GrabarPC
            Case "PCGRUPOS_EMP", "PROVINCIASEMP", "CUENTA_EMP": GrabarEmp
            Case "CUENTA_LOCAL": GrabarLocal 'jeaa 21/01/04
            Case "ITEM_FAMILIA": GrabarFamilia
            Case "SRI": Grabar
            Case "CUENTA_PRESUP", "CUENTASC", "CUENTA101": GrabarPresupuesto 'jeaa 09/01/2009
            Case "DIASREPO": Grabar
            Case "VENDE": Grabar
            Case "PCCLIRUC": GrabarPCBandRUCValido
            Case "CUENTAFE": GrabarFlujoEfectivo 'jeaa 09/01/2009
            Case "FORMAPAGOSRI": Grabar
            Case "DINARDAP", "PCPARR": GrabarPC
            Case "DIVNOMEMP": Grabar
            Case "COMPROB": Grabar
            Case "EMPDOC": Grabar
            Case "PCCLIRUCFCEL": GrabarPCBandRUCValido
            Case "ITEMFCEL": Grabar
            Case "PCEMAIL": GrabarPCemail
            Case "LECTURAS": GrabarLectura
            Case "FECHAINICIAL": Grabar
            Case "FORMASRI": Grabar
            Case "ESCOPIARTC": Grabar
            Case "PCAGENCIA":            GrabarPCAgencia
        End Select
    Case "Imprimir":    Imprimir
    Case "Cerrar":      Cerrar
    End Select
End Sub

Private Sub Buscar()
    Static coditem As String, CodAlt As String, _
           Desc As String, _
           codg As String, Numg As Integer, bandIVA As Boolean, bandFraccion As Boolean, bandIVA0 As Boolean
    Dim codg1 As String, codg2 As String, codg3 As String, codg4 As String, codg5 As String, codg6 As String
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
    If Not frmIVBusqueda.Inicio( _
                coditem, _
                CodAlt, _
                Desc, _
                codg1, codg2, codg3, codg4, codg5, codg6, _
                Numg, _
                bandIVA, _
                Me.tag, , , bandIVA0) Then
      'if not frmivbusqueda.InicioTrans (
        'Si fue cancelada la busqueda, sale no mas
        grd.SetFocus
        Exit Sub
    End If
    
    'Cambia la forma de cursor
    MensajeStatus MSG_PREPARA, vbHourglass
    
    'Compone la cadena de SQL
    sql = "SELECT CodInventario, CodAlterno1, Descripcion "
    Select Case Me.tag
    Case "IVA"
        sql = sql & ", PorcentajeIVA * 100 AS PorIVA "
    Case "CUENTA"
        sql = sql & ", CodCuentaActivo, CodCuentaCosto, CodCuentaVenta , CodCuentaDevolucion,  CodCuentaDiferida "
    Case "ITEM_IVGRUPOS"    'jeaa 24/09/04 asignacion de grupo a los items
        sql = sql & ", codGrupo1, codgrupo2, codGrupo3, codGrupo4, codgrupo5, codgrupo6 "
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
'    Case "IVEXIST"
 '       sql = sql & ", CodCuentaActivo, CodCuentaCosto, CodCuentaVenta "
    
    Case "DIASREPO"
        sql = sql & ", TiempoReposicion, FrecuenciaReposicion,TiempoPromvta "
    Case "PORDESC"
        sql = sql & ", Descuento1 * 100 AS PorDesc1"
        sql = sql & ", Descuento2 * 100 AS PorDesc2"
        sql = sql & ", Descuento3 * 100 AS PorDesc3"
        sql = sql & ", Descuento4 * 100 AS PorDesc4"
        sql = sql & ", Descuento5 * 100 AS PorDesc5 "
    Case "PORCOMI"
        sql = sql & ", Comision1 * 100 AS PorComi1"
        sql = sql & ", Comision2 * 100 AS PorComi2"
        sql = sql & ", Comision3 * 100 AS PorComi3"
        sql = sql & ", Comision4 * 100 AS PorComi4"
        sql = sql & ", Comision5 * 100 AS PorComi5 "
    Case "IVEXISTNEG"
        sql = sql & ", CodCuentaActivo, CodCuentaCosto, CodCuentaVenta "
    Case "COSTOREF"
        sql = sql & ", CostoReferencial "
    Case "ARANCEL"
        sql = sql & ", CodArancel "
    Case "CUENTASC"
        sql = sql & ", CodCuentaSC, Campo101 "
    Case "CUENTAFE"
        sql = sql & ", CodCuentaFE "
    Case "CUENTA101"
        sql = sql & " , CampoF101"
    End Select
    sql = sql & "FROM vwIVInventarioRecuperar "
    'CodInventario
    If Len(coditem) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodInventario LIKE '" & coditem & comodin & "') "
    End If
    
    'CodAlterno
    If Len(CodAlt) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
'        cond = cond & "((CodAlterno1 LIKE '" & CodAlt & comodin & "') " & _
'                      "OR (CodAlterno2 LIKE '" & CodAlt & comodin & "')) "

        cond = cond & "(CodAlterno1 LIKE '" & CodAlt & comodin & "') "
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
    
    If Len(codg6) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodGrupo6" & " = '" & codg6 & "') "
    End If
    
    
    
    If bandIVA Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & " BANDIVA=1 "
        If Me.tag = "COSTOUI" Then
                If Len(cond) > 0 Then cond = cond & "AND "
                cond = cond & " (costoultimoingreso Is Null Or costoultimoingreso = 0) "
        End If
    End If
    
    If bandIVA0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & " BANDIVA= 0 "
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

Private Sub BuscarPC(ByVal BandProv As Boolean)
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
        Case "PCGRUPOS_PROV", "PCGRUPOS_CLI", "PCGRUPOS_GAR"
            'Compone la cadena de SQL
            sql = "SELECT CodProvCli, Nombre" & _
            ", codgrupo1, codgrupo2, codgrupo3 ,codgrupo4 " & _
            "FROM vwPCProvCli "
        Case "PCCLIRUC"
            sql = "SELECT CodProvCli, Nombre" & _
            ", ruc, CodTipoDocumento " & _
            "FROM vwPCProvCli "
        Case "PCCLIRUCFCEL"
            sql = "SELECT CodProvCli, Nombre" & _
            ", ruc, CodTipoDocumento,direccion1, telefono1, email, estado " & _
            "FROM vwPCProvCli "
        Case "PCEMAIL"
            sql = "SELECT CodProvCli, Nombre" & _
            ", ruc, email " & _
            "FROM vwPCProvCli "
        Case "LECTURAS"
            sql = "SELECT codcentro,  Codprovcli, w.Nombre, valor1 " & _
            "FROM vwPCProvCli w inner join GNCENTROCOSTO gncc on gncc.IdCliente = w.idprovcli"
        Case "PCCLI_VENDEDOR"
            sql = "SELECT CodProvCli, Nombre" & _
            ", codvendedor " & _
            "FROM vwPCProvCli "
        Case "PCAGENCIA"
        
            sql = "SELECT CodProvCli, Nombre" & _
            " " & _
            "FROM vwPCProvCli "
        

        End Select
    ' si Busca Proveedor o cliente mediante bandera de prov
    If Len(cond) > 0 Then cond = cond & "AND "
    If BandProv Then
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

    Select Case Me.tag
        Case "PCCLIRUC"
        If Len(cond) > 0 Then
            If Len(cond) > 0 Then cond = cond & "AND "
            cond = cond & "((BandRUCValido <> 1) or (BandRUCValido is null))"
        End If
        Case "PCCLIRUCFCEL"
        If Len(cond) > 0 Then
            If Len(cond) > 0 Then cond = cond & "AND "
            cond = cond & "((BandRUCValido <> 1) or (BandRUCValido is null))"
        End If
        Case "PCEMAIL"
        If Len(cond) > 0 Then
            If Len(cond) > 0 Then cond = cond & "AND "
            cond = cond & "(LEN(EMAIL) > 0) "
        End If
    
    End Select


    If Len(cond) > 0 Then sql = sql & " WHERE " & cond
    
    
    If Len(gobjMain.objCondicion.Bienes) > 0 Then
        sql = sql & gobjMain.objCondicion.Bienes
    End If
    
    If Len(gobjMain.objCondicion.Usuario1) > 0 Then
        sql = sql & "and codvendedor='" & gobjMain.objCondicion.Usuario1 & "'"
    End If
    
    
    sql = sql & " ORDER BY Nombre "
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
    Dim s As String, i As Long, j As Integer
    With grd
    Select Case Me.tag
        Case "CUENTA_PROV", "CUENTA_CLI", "CUENTA_EMP"
            s = "^#|<Código|<Descripción|<Cuenta Contable1|<Cuenta Contable2"
        Case "PCCLIRUC"
            s = "^#|<Código|<Nombre|<RUC|^Tipo Documento|<Verificado"
        Case "IVA"
            s = "^#|<Código|<Cód.Alterno|<Descripción|>IVA"
        Case "CUENTA"
            s = "^#|<Código|<Cód.Alterno|<Descripción|<Cuenta Activo|<Cuenta Costo|<Cuenta Venta|<Cuenta Devolucion|<Cuenta Diferida"
        Case "CUENTA_LOCAL"     'jeaa 21/01/04
            s = "^#|<Código|<Descripción|^CodLocal|^Sucursal"
            grd.ColHidden(3) = True
        Case "ITEM_FAMILIA"
            s = "^#|<Código|<Descripción|>IdRecuperado|>CodFamilia|<Familia|<Id"
        Case "ITEM_IVGRUPOS" 'jeaa 24/09/04 asignacion de grupo a los items
            s = "^#|<Código|<Cód.Alterno|<Descripción"
            For j = 1 To 6
                s = s & "|<" & gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(j)
            Next j
        Case "PCGRUPOS_PROV", "PCGRUPOS_CLI", "PCGRUPOS_EMP", "PCGRUPOS_GAR"  'jeaa 24/09/04 asignacion de grupo a los items
            s = "^#|<Código|<Descripción"
            For j = 1 To 4 'Cambiado AUC 03/10/2005 Antes 3
                s = s & "|<" & gobjMain.EmpresaActual.GNOpcion.EtiqPCGrupo(j)
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
            s = "^#|<Código|<Cód.Alterno|<Descripción|>Tiempo Repo|>Frecuencia Repo|>Prom Vta"
        Case "PORDESC"
            s = "^#|<Código|<Cód.Alterno|<Descripción|>Desc 1|>Desc 2|>Desc 3|>Desc 4|>Desc 5"
        Case "PORCOMI"
            s = "^#|<Código|<Cód.Alterno|<Descripción|>Comi 1|>Comi 2|>Comi 3|>Comi 4|>Comi 5"
        Case "IVEXISTNEG"
            s = "^#|<idInventario|<Código|<Cód.Alterno|<Descripción|<idbodega|<CodBodega|>Ajuste"
        Case "COSTOREF"
            s = "^#|<Código|<Cód.Alterno|<Descripción|>Costo Referencial"
        Case "PROVINCIAS", "PROVINCIAS_PROV", "PROVINCIASEMP"
            s = "^#|<Código|<Nombre|<ProvinciaAnt|<CiudadAnt|<ProvinciaNue|<CantonNue|<Parroquia"
        Case "VENDE"
            s = "^#|<Transid|<Fecha Trans|<Código|<Num.Trans|Nombre|<CodVendedor"
            grd.ColHidden(1) = True
        Case "ARANCEL"
            s = "^#|<Código|<Cód.Alterno|<Descripción|<CodArancel"
        Case "AFEXIST"
            s = "^#|<idInventario|<Código|<Cód.Alterno|<Descripción|<idProvcli|<Cod Custodio|>Existencia"
        Case "CUENTASC"
            s = "^#|<Código|<Descripción|<Cuenta SC|>Campo F101"
        Case "CUENTAFE"
            s = "^#|<Código|<Descripción|<Cuenta FE"
        Case "FORMAPAGOSRI"
            s = "^#|<Transid|<Fecha Trans|<Código|<Num.Trans|<Proveedor|>RUC|^Num. Esta|^Num Punto|^Secuencial|>Cod Sustento|>Forma Pago SRI|>Nueva Forma Pago SRI"
            grd.ColHidden(1) = True
        Case "DINARDAP", "PCPARR"
            s = "^#|<Código|<RUC|<Nombre|<Provincia|<Desc. Prov|<Canton|<Desc. Canton|<Parroquia|<Desc. Parroq|<Tipo Sujeto|<Sexo|<Estado Civil|<Origen Ingresos"
            's = "^#|>Código|<RUC|<Nombre|<Provincia|<Canton|<Parroquia|<Tipo Sujeto|<Sexo|<Estado Civil|<Origen Ingresos"
        Case "DIVNOMEMP" 'jeaa 24/09/04 asignacion de grupo a los items
            s = "^#|idempleado|<Código|<Nombre Completo|<Apellido|<Nombre"
        Case "CUENTA101"
            s = "^#|<Código|<Descripción|<Campo F101"
        Case "COMPROB"
            s = "^#|<Transid|<Fecha Trans|<Código|<Num.Trans|<#Doc Ref.|<Descripción|<Comprobante Relacionado"
            grd.ColHidden(1) = True
        Case "EMPDOC"
            s = "^#|<Transid|<Fecha Trans|<Código|<Num.Trans|Nombre|<CodEmpleado"
            grd.ColHidden(1) = True
        Case "PCCLIRUCFCEL"
            s = "^#|<Código|<Nombre|<RUC|^Tipo Documento|<Dirección|<Telefono|<E-mail|^Estado|<Verificado"
        
        Case "ITEMFCEL"
            s = "^#|<Código|<Cód.Alterno|<Descripción|<Verificado"
        Case "PCEMAIL"
            s = "^#|<Código|<Nombre|<RUC|<E-mail|<Verificado"
        Case "LECTURAS"
            s = "^#|<No. Medidor|<RUC|<Nombre|>Lectura Anteior|>Lectura Nueva|<Resultado"
        Case "FECHAINICIAL"
            s = "^#|<Código|<Plan|<Código Plan|<Fecha Inicial"
        Case "FORMASRI"
            s = "^#|<id|<IdForma|<Fecha Trans|<Código|<Num.Trans|<Cod Forma|<Cod Forma SRI"
            grd.ColHidden(1) = True
            grd.ColHidden(2) = True
        Case "ESCOPIARTC"
            s = "^#|<Transid|<Fecha Trans|<Cliente|<Código|<Num.Trans|^Copia/Original|^Copia/Original New"
            grd.ColHidden(1) = True
        Case "PCCLI_VENDEDOR"
            s = "^#|<Código|<Nombre|<Cod Vendedor|<Cod Vendedor New"
        Case "PCAGENCIA"
            s = "^#|<Código|<Nombre|<Resultado"
        
        End Select
        .FormatString = s
        GNPoneNumFila grd, False
        AjustarAutoSize grd, -1, -1, 4000
        AsignarTituloAColKey grd
    
        'Columnas modificables (Longitud maxima)
        Select Case Me.tag
        Case "IVA"
            .ColData(.ColIndex("IVA")) = 5
        Case "CUENTA"
            .ColData(.ColIndex("Cuenta Activo")) = 20
            .ColData(.ColIndex("Cuenta Costo")) = 20
            .ColData(.ColIndex("Cuenta Venta")) = 20
            .ColData(.ColIndex("Cuenta Devolucion")) = 20
            .ColData(.ColIndex("Cuenta Diferida")) = 20
            CargarCuentas
        Case "CUENTA_PROV", "CUENTA_CLI", "CUENTA_EMP"
            .TextMatrix(0, .ColIndex("Descripción")) = "Nombre"  ' cambio caption de Descripcion a Nombre
            .ColData(.ColIndex("Cuenta Contable1")) = 20
            .ColData(.ColIndex("Cuenta Contable2")) = 20
            'Carga lista de Cuentas contables
            CargarCuentas
        Case "CUENTA_LOCAL" 'jeaa 21/01/04
            .TextMatrix(0, .ColIndex("Descripción")) = "Nombre de la Cuenta"  ' cambio caption de Descripcion a Nombre
             .ColData(.ColIndex("Sucursal")) = 20
             .ColWidth(.ColIndex("Sucursal")) = 3000
             If Not CargarLocales Then
                grd.Rows = 1
                Exit Sub
            End If
        Case "CUENTA_PREUP" 'jeaa 21/01/04
            .TextMatrix(0, .ColIndex("Descripción")) = "Nombre de la Cuenta"  ' cambio caption de Descripcion a Nombre
             .ColData(.ColIndex("Sucursal")) = 20
             .ColWidth(.ColIndex("Sucursal")) = 3000
             If Not CargarLocales Then
                grd.Rows = 1
                Exit Sub
            End If
        
        Case "ITEM_FAMILIA"
            grd.ColHidden(.ColIndex("IdRecuperado")) = True
            grd.ColHidden(.ColIndex("CodFamilia")) = True
            grd.ColHidden(.ColIndex("Id")) = True
            .ColData(.ColIndex("Familia")) = 20
            .ColWidth(.ColIndex("Familia")) = 3000
            If Not CargarFamilias Then
                grd.Rows = 1
                Exit Sub
            End If
        Case "ITEM_IVGRUPOS" 'jeaa 24/09/04 asignacion de grupo a los items
            .ColData(IVGRUPO1) = 20
            .ColData(IVGRUPO2) = 20
            .ColData(IVGRUPO3) = 20
            .ColData(IVGRUPO4) = 20
            .ColData(IVGRUPO5) = 20
            .ColData(IVGRUPO6) = 20
            .ColWidth(IVGRUPO1) = 2000
            .ColWidth(IVGRUPO2) = 2000
            .ColWidth(IVGRUPO3) = 2000
            .ColWidth(IVGRUPO4) = 2000
            .ColWidth(IVGRUPO5) = 2000
            .ColWidth(IVGRUPO6) = 2000
            If Not CargarIVGrupos Then
                grd.Rows = 1
                Exit Sub
            End If
        Case "PCGRUPOS_PROV", "PCGRUPOS_CLI", "PCGRUPOS_GAR"
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
            
        Case "PCGRUPOS_EMP"
            .TextMatrix(0, .ColIndex("Descripción")) = "Nombre"  ' cambio caption de Descripcion a Nombre
            .ColData(PCGRUPO1) = 20
            .ColData(PCGRUPO2) = 20
            .ColData(PCGRUPO3) = 20
            .ColData(PCGRUPO4) = 20
            .ColWidth(PCGRUPO1) = 2000
            .ColWidth(PCGRUPO2) = 2000
            .ColWidth(PCGRUPO3) = 2000
            .ColData(PCGRUPO4) = 2000
            If Not CargarEmpGrupos Then
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
            .ColWidth(.ColIndex("Autorizacion SRI")) = 1500
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
        Case "IVEXISTNEG"
            .ColData(.ColIndex("Existencia")) = 20
            .ColData(.ColIndex("Existencia Mínima")) = 20
            .ColData(.ColIndex("Existencia Máxima")) = 20
            .ColHidden(.ColIndex("idInventario")) = True
            .ColHidden(.ColIndex("idBodega")) = True
        
        Case "DIASREPO"
            .ColData(.ColIndex("Tiempo Repo")) = 5
            .ColData(.ColIndex("Frecuencia Repo")) = 5
            .ColData(.ColIndex("Prom Vta")) = 5
        Case "PORDESC"
            .ColData(DESC1) = 4
            .ColData(DESC2) = 5
            .ColData(DESC3) = 6
            .ColData(DESC4) = 7
            .ColData(DESC5) = 8
        Case "PORCOMI"
            .ColData(DESC1) = 4
            .ColData(DESC2) = 5
            .ColData(DESC3) = 6
            .ColData(DESC4) = 7
            .ColData(DESC5) = 8
        Case "COSTOREF"
            .ColData(.ColIndex("Costo Referencial")) = 5
            .ColFormat(.ColIndex("Costo Referencial")) = "##.0000"
        Case "PROVINCIAS", "PROVINCIAS_PROV", "PROVINCIASEMP"
        'Codigo|<Nombre|<ProvinciaAnt|<CiudadAnt
            .ColData(.ColIndex("Codigo")) = -1
            .ColData(.ColIndex("Nombre")) = -1
            .ColData(.ColIndex("ProvinciaAnt")) = -1
            .ColData(.ColIndex("CiudadAnt")) = -1
            .ColData(.ColIndex("Provincia")) = 20
            .ColData(.ColIndex("Canton")) = 20
            .ColData(.ColIndex("Parroquia")) = 20
        Case "VENDE"
            .ColData(.ColIndex("CodVendedor")) = 10
            If Not CargarVendedor Then
                grd.Rows = 1
                Exit Sub
            End If
        Case "ARANCEL"
            .ColData(.ColIndex("CodArancel")) = 20
            .ColWidth(.ColIndex("CodArancel")) = 3000
            CargarArancel
        Case "AFEXIST"
            .ColData(.ColIndex("Existencia")) = 20
            .ColData(.ColIndex("Existencia Mínima")) = 20
            .ColData(.ColIndex("Existencia Máxima")) = 20
            .ColHidden(.ColIndex("idInventario")) = True
            .ColHidden(.ColIndex("idProvcli")) = True
        Case "CUENTASC"
            .ColData(.ColIndex("Cuenta SC")) = 20
            .ColData(.ColIndex("Campo F101")) = 20
            CargarCuentasSC
        Case "PCCLIRUC"
            .ColData(.ColIndex("Codigo")) = -1
            .ColData(.ColIndex("Nombre")) = -1
            .ColData(.ColIndex("RUC")) = 20
            .ColData(.ColIndex("Tipo Documento")) = 20
            .ColWidth(.ColIndex("Verificado")) = 4000
            CargarTipoDocumentos
        Case "CUENTAFE"
            .ColData(.ColIndex("Cuenta FE")) = 20
            CargarCuentasFE
        Case "FORMAPAGOSRI"
             If Not CargarFormaCobroSRI Then
                grd.Rows = 1
                Exit Sub
            End If
        
            .ColData(.ColIndex("Nueva Forma Pago SRI")) = 2
        Case "DINARDAP", "PCPARR"
            CargarProvincias
'            .ColHidden(.ColIndex("idprovcli")) = True
        Case "DIVNOMEMP"
'            .TextMatrix(0, .ColIndex("Nombre")) = "Nombre"  ' cambio caption de Descripcion a Nombre
            .ColData(3) = 20
            .ColData(4) = 20
            .ColWidth(4) = 2000
            .ColWidth(5) = 2000
        Case "COMPROB"
            .ColData(.ColIndex("Comprobante Relacionado")) = 20
            CargarComprobantes
         Case "EMPDOC"
            .ColData(.ColIndex("CodEmpleado")) = 10
            If Not CargarEmpleado Then
                grd.Rows = 1
                Exit Sub
            End If
        Case "PCCLIRUCFCEL"
            .ColData(.ColIndex("Codigo")) = -1
            .ColData(.ColIndex("Nombre")) = 200
            .ColData(.ColIndex("RUC")) = 13
            .ColData(.ColIndex("Tipo Documento")) = 1
            .ColData(.ColIndex("Dirección")) = 200
            .ColData(.ColIndex("Telefono")) = 20
            .ColData(.ColIndex("E-mail")) = 20
            .ColWidth(.ColIndex("Verificado")) = 4000
            CargarTipoDocumentos
        Case "ITEMFCEL"
            .ColData(.ColIndex("Código")) = 20
            .ColData(.ColIndex("Cód.Alterno")) = 20
            .ColData(.ColIndex("Descripción")) = 100
        Case "PCEMAIL"
            .ColData(.ColIndex("Codigo")) = -1
            .ColData(.ColIndex("Nombre")) = 200
            .ColData(.ColIndex("RUC")) = 13
            .ColData(.ColIndex("E-mail")) = 20
            .ColWidth(.ColIndex("Verificado")) = 4000
        Case "LECTURAS"
            .ColData(.ColIndex("No. Medidor")) = -1
            .ColData(.ColIndex("Ruc")) = -1
            .ColData(.ColIndex("Nombre")) = -1
            .ColData(.ColIndex("Lectura Anterior")) = -1
            
'            .ColWidth(1) = 1000
'            .ColWidth(2) = 2000
'            .ColWidth(3) = 2000
'            .ColWidth(4) = 1000
'            .ColWidth(5) = 1000
            
            
            .ColData(.ColIndex("Lectura Nueva")) = 20
'            .ColWidth(.ColIndex("Verificado")) = 400
        Case "FORMASRI"
            .ColData(.ColIndex("Cod Forma")) = 20
            CargarFormasSRI
        Case "ESCOPIARTC"
            .ColData(.ColIndex("Copia/Original")) = 20
            .ColData(.ColIndex("Copia/Original New")) = 20
        Case "PCCLI_VENDEDOR"
            .TextMatrix(0, .ColIndex("Nombre")) = "Nombre"  ' cambio caption de Descripcion a Nombre
            .ColData(PCGRUPO1) = 20
            .ColData(PCGRUPO2) = 20
            .ColData(.ColIndex("Cod Vendedor")) = -1

        End Select
        'Columnas No modificables
        
        Select Case Me.tag
        Case "MINMAX"
            For i = 0 To .ColIndex("Existencia")
                .ColData(i) = -1
            Next i
            
            'Color de fondo
            If .Rows > .FixedRows Then
                .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("Existencia")) = .BackColorFrozen
            End If
        Case "EXIST"
            For i = 0 To .ColIndex("Existencia")
                .ColData(i) = -1
            Next i
            
            'Color de fondo
            If .Rows > .FixedRows Then
                .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("Existencia")) = .BackColorFrozen
            End If
        
        Case "PROVINCIAS", "PROVINCIAS_PROV", "PROVINCIASEMP"
            If .Rows > .FixedRows Then
                .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("CiudadAnt")) = .BackColorFrozen
            End If
        Case "VENDE", "EMPDOC"
            If .Rows > .FixedRows Then
                .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("Nombre")) = .BackColorFrozen
            End If
        Case "PCCLIRUC"
            If .Rows > .FixedRows Then
                .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("Código")) = .BackColorFrozen
            End If
        
        Case "PCCLIRUCFCEL"
            If .Rows > .FixedRows Then
                .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("Código")) = .BackColorFrozen
                .Cell(flexcpBackColor, .FixedRows, .ColIndex("Estado"), .Rows - 1, .ColIndex("Estado")) = .BackColorFrozen
            End If
        Case "PCEMAIL"
            If .Rows > .FixedRows Then
                .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("Código")) = .BackColorFrozen
            End If
        
        
        Case "FORMAPAGOSRI"
            If .Rows > .FixedRows Then
                .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("Forma Pago SRI")) = .BackColorFrozen
            End If
        Case "DINARDAP", "PCPARR"
            If .Rows > .FixedRows Then
                .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("Nombre")) = .BackColorFrozen
                .Cell(flexcpBackColor, .FixedRows, .ColIndex("Desc. Prov"), .Rows - 1, .ColIndex("Desc. Prov")) = .BackColorFrozen
                .Cell(flexcpBackColor, .FixedRows, .ColIndex("Desc. Canton"), .Rows - 1, .ColIndex("Desc. Canton")) = .BackColorFrozen
                .Cell(flexcpBackColor, .FixedRows, .ColIndex("Desc. Parroq"), .Rows - 1, .ColIndex("Desc. Parroq")) = .BackColorFrozen
            End If
        Case "DIVNOMEMP"
            If .Rows > .FixedRows Then
                .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("Nombre Completo")) = .BackColorFrozen
            End If
        Case "ITEMFCEL"
        Case "LECTURAS"
            If .Rows > .FixedRows Then
                .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("Lectura Anteior")) = .BackColorFrozen
            End If
       Case "FECHAINICIAL"
                .ColDataType(.ColIndex("Fecha Inicial")) = flexDTDate
        Case "FORMASRI"
            If .Rows > .FixedRows Then
                .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("Cod Forma")) = .BackColorFrozen
            End If
        Case "ESCOPIARTC"
            If .Rows > .FixedRows Then
                .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("Copia/Original")) = .BackColorFrozen
            End If
        Case "PCCLI_VENDEDOR"
        
            If .Rows > .FixedRows Then
                .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .ColIndex("cod Vendedor")) = .BackColorFrozen
            End If
        
        Case "PCAGENCIA"
        Case Else
            For i = 0 To .ColIndex("Descripción")
                .ColData(i) = -1
            Next i
            
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
    If Me.tag = "CUENTA_PROV" Or Me.tag = "CUENTA_CLI" Or Me.tag = "CUENTA_EMP" Then
        .ColComboList(.ColIndex("Cuenta Contable1")) = s
        .ColComboList(.ColIndex("Cuenta Contable2")) = s
    Else
        .ColComboList(.ColIndex("Cuenta Activo")) = s
        .ColComboList(.ColIndex("Cuenta Costo")) = s
        .ColComboList(.ColIndex("Cuenta Venta")) = s
        .ColComboList(.ColIndex("Cuenta Devolucion")) = s
        .ColComboList(.ColIndex("Cuenta Diferida")) = s
    End If
        
    End With
End Sub

Private Sub Asignar()
    Select Case Me.tag
        Case "IVA":        AsignarIVA
        Case "CUENTA":     AsignarCuenta
        Case "CUENTA_PROV", "CUENTA_CLI", "CUENTA_EMP": AsignarCuentaPC
        Case "CUENTA_LOCAL": AsignarLocal
        Case "ITEM_FAMILIA": AsignarFamilia
        Case "ITEM_IVGRUPOS": AsignarIVGrupos   'jeaa 24/09/04 asignacion de grupo a los items
        Case "PCGRUPOS_PROV", "PCGRUPOS_CLI", "PCGRUPOS_EMP", "PCGRUPOS_GAR": AsignarPCGrupos  'jeaa 24/09/04 asignacion de grupo a los prov-cli
        Case "PCCLI_VENDEDOR": AsignarFCVendedor
        Case "FRACCION":  AsignarFraccion
        Case "AREA":  AsignarArea  'jeaa 15/09/2005
        Case "VENTA":  AsignarVenta 'jeaa 26/12/2005
        Case "COSTOUI":        AsignarCostoUI
        Case "SRI":        AsignarSRI
        Case "MINMAX":     AsignarMinMax
        Case "IVEXIST":     AsignarIVExist
        Case "CUENTA_PRESUP": AsignarPresupuesto
        Case "DIASREPO":        AsignarDias
        Case "PORDESC":        AsignarPorDescuento
        Case "PORCOMI":        AsignarPorComision
        Case "IVEXISTNEG":     AsignarIVExist
        Case "COSTOREF":        AsignarCostoReferencial
        Case "PROVINCIAS", "PROVINCIAS_PROV", "PROVINCIASEMP":    AsignarProvincias
        Case "VENDE":        AsignarVendedor
        Case "ARANCEL":     AsignarArancel
        Case "AFEXIST":     AsignarAFExist
        Case "CUENTASC":     AsignarCuentaSC
        Case "CUENTAFE": AsignarCuentaFE
        Case "FORMAPAGOSRI":        AsignarFormaPagoSRI
        Case "DINARDAP", "PCPARR":        AsignarDINARDAP
        Case "DIVNOMEMP":    AsignarApellidoNombre
        Case "CUENTA101":     AsignarCuenta101
        Case "EMPDOC":        AsignarEmpleado 'AUC para kardex de documentos
        Case "FECHAINICIAL":        AsignarFechaInicial
        Case "FORMASRI":        AsignarFormaCobro
'        Case "ESCOPIARTC": AsignarFormaCobro
    End Select
End Sub

Private Sub AsignarIVA()
    Dim s As String, v As Single
    Dim i As Long
    
    s = InputBox("Ingrese el valor de IVA (%)", "Asignar un valor", "15")
    If IsNumeric(s) Then
        v = CSng(s)
    Else
        MsgBox "Debe ingresar un valor numérico. (ejm. 15 para 15%)", vbInformation
        grd.SetFocus
        Exit Sub
    End If
    
    With grd
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("IVA")) = v
        Next i
    End With
End Sub

Private Sub AsignarCostoUI()
    Dim s As String, v As Single
    Dim i As Long
    
    s = InputBox("Ingrese el valor de Costo ULtimo Ingreso ", "Asignar un valor", "15")
    If IsNumeric(s) Then
        v = CSng(s)
    Else
        MsgBox "Debe ingresar un valor numérico. (ejm. 1.1544) ", vbInformation
        grd.SetFocus
        Exit Sub
    End If
    
    With grd
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("Costo Ultimo Ingreso")) = v
        Next i
    End With
End Sub


Private Sub AsignarCuenta()
    Dim activo As String, costo As String, venta As String, Devol As String, Difer As String
    Dim i As Long, s As String
    
    With grd
        'Obtiene cuentas de la fila actual
        activo = .TextMatrix(.Row, .ColIndex("Cuenta Activo"))
        costo = .TextMatrix(.Row, .ColIndex("Cuenta Costo"))
        venta = .TextMatrix(.Row, .ColIndex("Cuenta Venta"))
        Devol = .TextMatrix(.Row, .ColIndex("Cuenta Devolucion"))
        Difer = .TextMatrix(.Row, .ColIndex("Cuenta Diferida"))
        
        'Confirma las cuentas
        s = "Está seguro que desea asignar los siguientes códigos " & _
            "en todos los ítems que están visualizados?" & vbCr & vbCr & _
            "    Cuenta de Activo:  " & activo & vbCr & _
            "    Cuenta de Costo:   " & costo & vbCr & _
            "    Cuenta de Venta:   " & venta & vbCr & _
            "    Cuenta de Devolucion:   " & Devol & vbCr & _
            "    Cuenta de Diferida:   " & Difer
        If MsgBox(s, vbQuestion + vbYesNo) <> vbYes Then
            .SetFocus
            Exit Sub
        End If
        
        'Copia a todas las filas los mismos códigos de cuenta
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("Cuenta Activo")) = activo
            .TextMatrix(i, .ColIndex("Cuenta Costo")) = costo
            .TextMatrix(i, .ColIndex("Cuenta Venta")) = venta
            .TextMatrix(i, .ColIndex("Cuenta Devolucion")) = Devol
            .TextMatrix(i, .ColIndex("Cuenta Diferida")) = Difer
        Next i
    End With
End Sub

Private Sub AsignarCuentaPC()
    Dim Cuenta1 As String, Cuenta2 As String
    Dim i As Long, s As String, provCli As String
    provCli = IIf(Me.tag = "CUENTA_PROV", "Proveedores", "Clientes")
    With grd
        'Obtiene cuentas de la fila actual
        Cuenta1 = .TextMatrix(.Row, .ColIndex("Cuenta Contable1"))
        Cuenta2 = .TextMatrix(.Row, .ColIndex("Cuenta Contable2"))
        
        
        'Confirma las cuentas
        s = "Está seguro que desea asignar los siguientes códigos " & _
            "en todos los " & provCli & " que están visualizados?" & vbCr & vbCr & _
            "    Cuenta Contable1:  " & Cuenta1 & vbCr & _
            "    Cuenta Contable2:   " & Cuenta2
            
        If MsgBox(s, vbQuestion + vbYesNo) <> vbYes Then
            .SetFocus
            Exit Sub
        End If
        
        'Copia a todas las filas los mismos códigos de cuenta
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("Cuenta Contable1")) = Cuenta1
            .TextMatrix(i, .ColIndex("Cuenta Contable2")) = Cuenta2
        Next i
    End With
End Sub


Private Sub Grabar()
    Dim i As Long, iv As IVinventario, cod As String, X As Single
    Dim gnc As GNComprobante
    Dim sql As String, rs As Recordset
    Dim IdBodega As Integer, W As Variant
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
        ElseIf Me.tag = "AFEXIST" Then
            sql = "select idProvcli from Pcprovcli where codProvcli='" & .TextMatrix(2, .ColIndex("Cod Custodio")) & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            IdBodega = rs.Fields("idProvcli")
        End If
        
        prg1.min = 0
        prg1.max = 1
        If .Rows > .FixedRows Then prg1.max = .Rows - 1
        For i = .FixedRows To .Rows - 1
            'Si es que se canceló el proceso
            If mCancelado Then GoTo salida
        
            prg1.value = i
            grd.Row = i
            X = grd.CellTop
            
            cod = .TextMatrix(i, .ColIndex("Código"))
            MensajeStatus i & " de " & .Rows - .FixedRows, vbHourglass
            DoEvents
            
            'Recupera el objeto de Inventario
            If Me.tag <> "COMPROB" And Me.tag <> "SRI" And Me.tag <> "VENDE" And Me.tag <> "FORMAPAGOSRI" And Me.tag <> "DINARDAP" And Me.tag <> "EMPDOC" And Me.tag <> "FORMASRI" Then
                Set iv = gobjMain.EmpresaActual.RecuperaIVInventario(cod)
            End If
            
            Select Case Me.tag
            Case "IVA"
                If iv.PorcentajeIVA <> .ValueMatrix(i, .ColIndex("IVA")) / 100 Then
                    iv.PorcentajeIVA = .ValueMatrix(i, .ColIndex("IVA")) / 100
                End If
                If .ValueMatrix(i, .ColIndex("IVA")) = 0 Then
                    If iv.bandIVA <> False Then
                        iv.bandIVA = False
                    End If
                Else
                    If iv.bandIVA <> True Then
                        iv.bandIVA = True
                    End If
                End If
                
            Case "CUENTA"
                If iv.CodCuentaActivo <> .TextMatrix(i, .ColIndex("Cuenta Activo")) Then
                    iv.CodCuentaActivo = .TextMatrix(i, .ColIndex("Cuenta Activo"))
                End If
                If iv.CodCuentaCosto <> .TextMatrix(i, .ColIndex("Cuenta Costo")) Then
                    iv.CodCuentaCosto = .TextMatrix(i, .ColIndex("Cuenta Costo"))
                End If
                If iv.CodCuentaVenta <> .TextMatrix(i, .ColIndex("Cuenta Venta")) Then
                    iv.CodCuentaVenta = .TextMatrix(i, .ColIndex("Cuenta Venta"))
                End If
            
                If iv.CodCuentaDevolucion <> .TextMatrix(i, .ColIndex("Cuenta Devolucion")) Then
                    iv.CodCuentaDevolucion = .TextMatrix(i, .ColIndex("Cuenta Devolucion"))
                End If
            
                If iv.CodCuentaDiferida <> .TextMatrix(i, .ColIndex("Cuenta Diferida")) Then
                    iv.CodCuentaDiferida = .TextMatrix(i, .ColIndex("Cuenta Diferida"))
                End If
            
            
            Case "ITEM_IVGRUPOS"    'jeaa 24/09/04 asignacion de grupo a los items
                If iv.CodGrupo(1) <> .TextMatrix(i, IVGRUPO1) Then
                    iv.CodGrupo(1) = .TextMatrix(i, IVGRUPO1)
                End If
                If iv.CodGrupo(2) <> .TextMatrix(i, IVGRUPO2) Then
                    iv.CodGrupo(2) = .TextMatrix(i, IVGRUPO2)
                End If
                If iv.CodGrupo(3) <> .TextMatrix(i, IVGRUPO3) Then
                    iv.CodGrupo(3) = .TextMatrix(i, IVGRUPO3)
                End If
                If iv.CodGrupo(4) <> .TextMatrix(i, IVGRUPO4) Then
                    iv.CodGrupo(4) = .TextMatrix(i, IVGRUPO4)
                End If
                If iv.CodGrupo(5) <> .TextMatrix(i, IVGRUPO5) Then
                    iv.CodGrupo(5) = .TextMatrix(i, IVGRUPO5)
                End If
            
                If iv.CodGrupo(6) <> .TextMatrix(i, IVGRUPO6) Then
                    iv.CodGrupo(6) = .TextMatrix(i, IVGRUPO6)
                End If
            
            Case "FRACCION"
                If iv.bandFraccion <> .ValueMatrix(i, .ColIndex("Venta por Fraccion")) Then
                    iv.bandFraccion = .ValueMatrix(i, .ColIndex("Venta por Fraccion"))
                End If
            Case "AREA" 'jeaa 15/09/2005
                If iv.BandArea <> .ValueMatrix(i, .ColIndex("Venta por Area")) Then
                    iv.BandArea = .ValueMatrix(i, .ColIndex("Venta por Area"))
                End If
            Case "VENTA" 'jeaa 26/12/2005
                If iv.bandVenta <> .ValueMatrix(i, .ColIndex("Venta")) Then
                    iv.bandVenta = .ValueMatrix(i, .ColIndex("Venta"))
                End If
            Case "COSTOUI"
                If iv.CostoUltimoIngreso <> .ValueMatrix(i, .ColIndex("Costo Ultimo Ingreso")) Then
                    sql = " UPDATE IvInventario "
                    sql = sql & " SET CostoUltimoIngreso= " & .ValueMatrix(i, .ColIndex("Costo Ultimo Ingreso"))
                    sql = sql & " where codinventario='" & iv.CodInventario & "'"
                    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                End If
            Case "SRI"
                    sql = " UPDATE GnComprobante "
                    sql = sql & " SET AutorizacionSRI= '" & .TextMatrix(i, .ColIndex("Autorizacion SRI")) & "', "
                    sql = sql & " FechaCaducidadSRI= '" & .TextMatrix(i, .ColIndex("Fecha Caducidad")) & "' "
                    sql = sql & " where transid=" & .ValueMatrix(i, .ColIndex("Transid"))
                    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                    
            Case "MINMAX"
                    sql = " UPDATE ivexist "
                    sql = sql & " SET ExistMin= '" & .ValueMatrix(i, .ColIndex("Existencia Mínima")) & "', "
                    sql = sql & " ExistMax= '" & .ValueMatrix(i, .ColIndex("Existencia Máxima")) & "' "
                    sql = sql & " where idinventario=" & .TextMatrix(i, .ColIndex("idInventario"))
                    sql = sql & " and idbodega=" & .TextMatrix(i, .ColIndex("idbodega"))
                    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
           
            Case "IVEXIST"
                    sql = " insert ivexist (IdInventario,idbodega,exist) values ("
                    sql = sql & .ValueMatrix(i, .ColIndex("idInventario")) & ","
                    sql = sql & IdBodega & ","
                    sql = sql & " 0) "

                    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
              Case "DIASREPO"
                If iv.TiempoReposicion <> .ValueMatrix(i, .ColIndex("Tiempo Repo")) Then
                    iv.TiempoReposicion = .ValueMatrix(i, .ColIndex("Tiempo Repo"))
                End If
            
                If iv.FrecuenciaReposicion <> .ValueMatrix(i, .ColIndex("Frecuencia Repo")) Then
                    iv.FrecuenciaReposicion = .ValueMatrix(i, .ColIndex("Frecuencia Repo"))
                End If
            
                If iv.TiempoPromVta <> .ValueMatrix(i, .ColIndex("Prom Vta")) Then
                    iv.TiempoPromVta = .ValueMatrix(i, .ColIndex("Prom Vta"))
                End If
            
            
            Case "PORDESC"
                If iv.Descuento(1) <> .ValueMatrix(i, DESC1) / 100 Then
                    iv.Descuento(1) = .ValueMatrix(i, DESC1) / 100
                End If
                If iv.Descuento(2) <> .ValueMatrix(i, DESC2) / 100 Then
                    iv.Descuento(2) = .ValueMatrix(i, DESC2) / 100
                End If
                If iv.Descuento(3) <> .ValueMatrix(i, DESC3) / 100 Then
                    iv.Descuento(3) = .ValueMatrix(i, DESC3) / 100
                End If
                If iv.Descuento(4) <> .ValueMatrix(i, DESC4) / 100 Then
                    iv.Descuento(4) = .ValueMatrix(i, DESC4) / 100
                End If
                If iv.Descuento(5) <> .ValueMatrix(i, DESC5) / 100 Then
                    iv.Descuento(5) = .ValueMatrix(i, DESC5) / 100
                End If
            Case "PORCOMI"
                If iv.Comision(1) <> .ValueMatrix(i, DESC1) / 100 Then
                    iv.Comision(1) = .ValueMatrix(i, DESC1) / 100
                End If
                If iv.Comision(2) <> .ValueMatrix(i, DESC2) / 100 Then
                    iv.Comision(2) = .ValueMatrix(i, DESC2) / 100
                End If
                If iv.Comision(3) <> .ValueMatrix(i, DESC3) / 100 Then
                    iv.Comision(3) = .ValueMatrix(i, DESC3) / 100
                End If
                If iv.Comision(4) <> .ValueMatrix(i, DESC4) / 100 Then
                    iv.Comision(4) = .ValueMatrix(i, DESC4) / 100
                End If
                If iv.Comision(5) <> .ValueMatrix(i, DESC5) / 100 Then
                    iv.Comision(5) = .ValueMatrix(i, DESC5) / 100
                End If
            Case "IVEXISTNEG"
                    sql = " insert ivexist (IdInventario,idbodega,exist) values ("
                    sql = sql & .ValueMatrix(i, .ColIndex("idInventario")) & ","
                    sql = sql & IdBodega & ","
                    sql = sql & " 0) "

                    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            Case "COSTOREF"
                If iv.CostoReferencial <> .ValueMatrix(i, .ColIndex("Costo Referencial")) Then
                    sql = " UPDATE IvInventario "
                    sql = sql & " SET Costoreferencial= " & .ValueMatrix(i, .ColIndex("Costo Referencial"))
                    sql = sql & " where codinventario='" & iv.CodInventario & "'"
                    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                End If
            Case "VENDE"
                    sql = " UPDATE GnComprobante "
                    sql = sql & " SET Idvendedor= (select idvendedor from fcvendedor where codvendedor='" & .TextMatrix(i, .ColIndex("CodVendedor")) & "') "
                    sql = sql & " where transid=" & .ValueMatrix(i, .ColIndex("Transid"))
                    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                    sql = " UPDATE PcKardex "
                    sql = sql & " SET Idvendedor= (select idvendedor from fcvendedor where codvendedor='" & .TextMatrix(i, .ColIndex("CodVendedor")) & "') "
                    sql = sql & " where transid=" & .ValueMatrix(i, .ColIndex("Transid"))
                    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            Case "ARANCEL"
                If iv.CodArancel <> .TextMatrix(i, .ColIndex("CodArancel")) Then
                    iv.CodArancel = .TextMatrix(i, .ColIndex("CodArancel"))
                End If
            Case "AFEXIST"
                    sql = " insert afexistCustodio (IdInventario,idProvcli,exist) values ("
                    sql = sql & .ValueMatrix(i, .ColIndex("idInventario")) & ","
                    sql = sql & IdBodega & ","
                    sql = sql & " 0) "

                    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            Case "FORMAPAGOSRI"
                    sql = " UPDATE anexos "
                    'sql = sql & " SET codformapagoSRI= '" & .TextMatrix(i, .ColIndex("Nueva Forma Pago SRI")) & "'"
                    sql = sql & " SET codformapagoSRI = (select id from anexo_formapago where codformapago='" & .TextMatrix(i, .ColIndex("Nueva Forma Pago SRI")) & "') "
                    sql = sql & " where transid=" & .ValueMatrix(i, .ColIndex("Transid"))
                    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            
            Case "DIVNOMEMP"
                    sql = " UPDATE personal "
                    sql = sql & " SET papellido= '" & .TextMatrix(i, .ColIndex("Apellido")) & "',"
                    sql = sql & " pnombre= '" & .TextMatrix(i, .ColIndex("Nombre")) & "'"
                    sql = sql & " where idempleado=" & .ValueMatrix(i, .ColIndex("idEmpleado"))
                    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            Case "COMPROB"
                    If Len(.TextMatrix(i, .ColIndex("Comprobante Relacionado"))) > 0 Then
                        W = Split(.TextMatrix(i, .ColIndex("Comprobante Relacionado")), "-")
                        sql = " UPDATE GnComprobante "
                        sql = sql & " SET Idtransfuente = (select transid from gncomprobante where codtrans='" & W(0) & "' and numtrans = " & W(1) & ") "
                        sql = sql & " where transid=" & .ValueMatrix(i, .ColIndex("Transid"))
                    Else
                        sql = " UPDATE GnComprobante "
                        sql = sql & " SET Idtransfuente = 0 "
                        sql = sql & " where transid=" & .ValueMatrix(i, .ColIndex("Transid"))
                    
                    End If
                    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            Case "EMPDOC"
            '---------------
                Dim IdProvCli As Integer, TransID As Long
                sql = "SELECT IdProvCli FROM Empleado where codProvCli='" & .TextMatrix(i, .ColIndex("CodEmpleado")) & "'"
                'recupera idProvCli
                Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                IdProvCli = rs.Fields("IdProvCli")
                'recupera Transid
                TransID = .ValueMatrix(i, .ColIndex("Transid"))
                sql = "INSERT INTO GNExistDocumento (Transid,IdProvcli,Exist)  " & _
                          "VALUES(" & TransID & "," & IdProvCli & "," & 1 & ")"
                gobjMain.EmpresaActual.EjecutarSQL sql, 1
                rs.Close
                Set rs = Nothing
            '---------------
            Case "ITEMFCEL"
                    iv.CodInventario = .TextMatrix(i, .ColIndex("Código"))
                    iv.CodAlterno1 = .TextMatrix(i, .ColIndex("Cód.Alterno"))
                    iv.Descripcion = .TextMatrix(i, .ColIndex("Descripción"))
                    grd.TextMatrix(i, .ColIndex("Verificado")) = "OK"
            Case "FECHAINICIAL" 'FECHA INICIAL MANTENIMIENTO
                Dim gnv As GnVehiculo
                Dim ivp As IVPlan
                If Len(grd.TextMatrix(i, 4)) = 0 Then
                    Err.Raise ERR_INVALIDO, "FechaInicial", _
                        "Debe Ingresar una fecha Inicial " & vbCr & _
                        "Digite una fecha"
                End If
                Set gnv = gobjMain.EmpresaActual.RecuperaGNVehiculo(grd.TextMatrix(i, 1))
                Set ivp = gobjMain.EmpresaActual.RecuperaIVPLAN(grd.TextMatrix(i, 2))
                sql = "Insert into GNVehiculoPlan(idvehiculo,idPlan,FechaProx,bandValida) values ("
                sql = sql & gnv.IdVehiculo & "," & ivp.IdPlan & ",'" & grd.TextMatrix(i, 4) & "',1)"
                gobjMain.EmpresaActual.EjecutarSQL sql, 1
            Case "FORMASRI"
                    If Len(.TextMatrix(i, .ColIndex("Cod Forma SRI"))) > 0 Then
                        sql = " UPDATE pckardex "
                        sql = sql & " SET IdformaSri = (select id from anexo_formapago where codformapago='" & .TextMatrix(i, .ColIndex("Cod Forma SRI")) & "') " & " where id=" & .ValueMatrix(i, .ColIndex("id"))
                    End If
                    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            Case "ESCOPIARTC":

                    If .ValueMatrix(i, .ColIndex("Copia/Original New")) <> .ValueMatrix(i, .ColIndex("Copia/Original")) Then
                        sql = " UPDATE gncomprobante "
                        sql = sql & " SET estado1 = " & IIf(.ValueMatrix(i, .ColIndex("Copia/Original New")) = 0, 0, 1)
                        sql = sql & " where transid =" & .ValueMatrix(i, .ColIndex("transid"))
                        
                        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                    End If
                    

            End Select
            If Me.tag <> "COMPROB" And Me.tag <> "COSTOUI" And Me.tag <> "SRI" And Me.tag <> "VENDE" And Me.tag <> "AFEXIST" And Me.tag <> "FORMAPAGOSRI" And Me.tag <> "DINARDAP" And Me.tag <> "DIVNOMEMP" And Me.tag <> "EMPDOC" And Me.tag <> "FECHAINICIAL" And Me.tag <> "FORMASRI" And Me.tag <> "ESCOPIARTC" Then
                iv.Grabar
            End If
        Next i
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
    Dim i As Long, pc As PCProvCli, cod As String
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
        For i = .FixedRows To .Rows - 1
            'Si es que se canceló el proceso
            If mCancelado Then GoTo salida
        
            prg1.value = i
            cod = .TextMatrix(i, .ColIndex("Código"))
            MensajeStatus i & " de " & .Rows - .FixedRows, vbHourglass
            DoEvents
            
            'Recupera el objeto de Inventario
            Set pc = gobjMain.EmpresaActual.RecuperaPCProvCli(cod)
            Select Case Me.tag
                Case "CUENTA_PROV", "CUENTA_CLI"
                    If pc.CodCuentaContable <> .TextMatrix(i, .ColIndex("Cuenta Contable1")) Then
                        pc.CodCuentaContable = .TextMatrix(i, .ColIndex("Cuenta Contable1"))
                    End If
                    
                    If pc.CodCuentaContable2 <> .TextMatrix(i, .ColIndex("Cuenta Contable2")) Then
                        pc.CodCuentaContable2 = .TextMatrix(i, .ColIndex("Cuenta Contable2"))
                    End If
                    
                
                Case "PCGRUPOS_PROV ", "PCGRUPOS_CLI", "PCGRUPOS_GAR"
                    If pc.CodGrupo1 <> .TextMatrix(i, PCGRUPO1) Then
                        pc.CodGrupo1 = .TextMatrix(i, PCGRUPO1)
                    End If
                    If pc.CodGrupo2 <> .TextMatrix(i, PCGRUPO2) Then
                        pc.CodGrupo2 = .TextMatrix(i, PCGRUPO2)
                    End If
                    If pc.CodGrupo3 <> .TextMatrix(i, PCGRUPO3) Then
                        pc.CodGrupo3 = .TextMatrix(i, PCGRUPO3)
                    End If
                    'AUC 03/10/2005
                  If pc.CodGrupo4 <> .TextMatrix(i, PCGRUPO4) Then
                        pc.CodGrupo4 = .TextMatrix(i, PCGRUPO4)
                  End If
                Case "PROVINCIAS", "PROVINCIAS_PROV"
                        If Len(.TextMatrix(i, .ColIndex("ProvinciaNue"))) > 0 Then pc.codProvincia = .TextMatrix(i, .ColIndex("ProvinciaNue"))
                        If Len(.TextMatrix(i, .ColIndex("CantonNue"))) > 0 Then pc.codCanton = .TextMatrix(i, .ColIndex("CantonNue"))
                        If Len(.TextMatrix(i, .ColIndex("Parroquia"))) > 0 Then pc.codParroquia = .TextMatrix(i, .ColIndex("Parroquia"))
                        pc.Provincia = "" 'ENCERO LO ANTERIOR
                        pc.Ciudad = "" 'ENCERO LO ANTERIOR
                Case "DINARDAP", "PCPARR"
                        'sql = " UPDATE Pcprovcli "
                        If Len(.TextMatrix(i, .ColIndex("Provincia"))) > 0 Then
                            
                            pc.codProvincia = .TextMatrix(i, .ColIndex("Provincia"))
                        End If
                        If Len(.TextMatrix(i, .ColIndex("Canton"))) > 0 Then pc.codCanton = .TextMatrix(i, .ColIndex("Canton"))
                        If Len(.TextMatrix(i, .ColIndex("Parroquia"))) > 0 Then pc.codParroquia = .TextMatrix(i, .ColIndex("Parroquia"))
                        If Len(.TextMatrix(i, .ColIndex("Tipo Sujeto"))) > 0 Then pc.Tiposujeto = .TextMatrix(i, .ColIndex("Tipo Sujeto"))
                        If Len(.TextMatrix(i, .ColIndex("Sexo"))) > 0 Then pc.sexo = .TextMatrix(i, .ColIndex("Sexo"))
                        If Len(.TextMatrix(i, .ColIndex("Estado Civil"))) > 0 Then pc.EstadoCivil = .TextMatrix(i, .ColIndex("Estado Civil"))
                        If Len(.TextMatrix(i, .ColIndex("Origen Ingresos"))) > 0 Then pc.OrigenIngresos = .TextMatrix(i, .ColIndex("Origen Ingresos"))
                Case "PCCLI_VENDEDOR"
                    If .TextMatrix(i, PCGRUPO1) <> .TextMatrix(i, PCGRUPO2) Then
                        pc.CodVendedor = .TextMatrix(i, PCGRUPO2)
                    End If
                Case "PCAGENCIA"
                        
            End Select
            pc.Grabar
        Next i
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
    ElseIf Me.tag = "CUENTASC" Then
        frmPCBusqueda.Caption = "Busqueda de Cuentas Contables SC"
        frmPCBusqueda.cboGrupo.Visible = False
        frmPCBusqueda.fcbGrupo.Visible = False
        frmPCBusqueda.Label5.Visible = False
        frmPCBusqueda.cmdAceptar.Top = frmPCBusqueda.fcbGrupo.Top
        frmPCBusqueda.cmdCancelar.Top = frmPCBusqueda.cmdAceptar.Top
        frmPCBusqueda.Height = 2000
    ElseIf Me.tag = "CUENTAFE" Then
        frmPCBusqueda.Caption = "Busqueda de Cuentas Contables FE"
        frmPCBusqueda.cboGrupo.Visible = False
        frmPCBusqueda.fcbGrupo.Visible = False
        frmPCBusqueda.Label5.Visible = False
        frmPCBusqueda.cmdAceptar.Top = frmPCBusqueda.fcbGrupo.Top
        frmPCBusqueda.cmdCancelar.Top = frmPCBusqueda.cmdAceptar.Top
        frmPCBusqueda.Height = 2000
    ElseIf Me.tag = "CUENTA101" Then
        frmPCBusqueda.Caption = "Busqueda de Cuentas Contables SC"
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
        cond = cond & "(ctcuenta.codCuenta LIKE '" & codcuenta & comodin & "') "
    End If

    
    'Nombre
    If Len(nombre) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(ctcuenta.NombreCuenta LIKE '" & comodin & nombre & comodin & "') "
    End If

    
    sql = "SELECT ctcuenta.CodCuenta, ctcuenta.NombreCuenta "
    If Me.tag = "CUENTA_LOCAL" Then
        sql = sql & " ,codlocal,nombre FROM ctlocal right join ctcuenta "
        sql = sql & " on ctlocal.idlocal = ctcuenta.idlocal "
    ElseIf Me.tag = "CUENTA_PRESUP" Then
        sql = sql & " , isnull(valPresupuesto,0) as valPresupuesto  FROM  ctcuenta "
    ElseIf Me.tag = "CUENTASC" Then
        sql = sql & " , CTCUENTASC.CODCUENTA, CAMPO101  FROM  ctcuenta LEFT JOIN CTCUENTASC ON CTCUENTA.IDCUENTASC = CTCUENTASC.IDCUENTA "
    ElseIf Me.tag = "CUENTAFE" Then
        sql = sql & " , CTCUENTAFE.CODCUENTA FROM  ctcuenta LEFT JOIN CTCUENTAFE ON CTCUENTA.IDCUENTAFE = CTCUENTAFE.IDCUENTA "
    ElseIf Me.tag = "CUENTA101" Then
        sql = sql & " , CTCUENTA.CAMPOF101  FROM  ctcuenta "
    
    End If

    If Len(cond) > 0 Then
        sql = sql & " where ctcuenta.bandtotal=0 and " & cond
    Else
        sql = sql & " where ctcuenta.bandtotal=0"
    End If
    sql = sql & " ORDER BY ctcuenta.CodCuenta"
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
    Dim i As Long, s As String
    
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
        For i = .Row To .Rows - 1
            .TextMatrix(i, .ColIndex("Sucursal")) = Sucursal
            .TextMatrix(i, .ColIndex("CodLocal")) = codlocal
        Next i
    End With
    End Sub

Private Sub GrabarLocal()
    Dim i As Long, ct As CtCuenta, cod As String
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
        For i = .FixedRows To .Rows - 1
            'Si es que se canceló el proceso
            If mCancelado Then GoTo salida
        
            prg1.value = i
            cod = .TextMatrix(i, .ColIndex("Código"))
            MensajeStatus i & " de " & .Rows - .FixedRows, vbHourglass
            DoEvents
            
            'Recupera el objeto de Inventario
            Set ct = gobjMain.EmpresaActual.RecuperaCTCuenta(cod)
            ct.codlocal = .TextMatrix(i, .ColIndex("CodLocal"))
            ct.Grabar
        Next i
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

Private Sub RecuperaCodLocal(ByVal nombre As String, i As Long)
On Error GoTo ErrTrap
    Dim codlocal As String, rs As Recordset, sql As String
    sql = "SELECT codlocal FROM ctlocal where nombre = '" & nombre & "'"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not rs.EOF Then
        grd.TextMatrix(i, grd.ColIndex("CodLocal")) = rs.Fields("codlocal")
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

Private Sub RecuperaCodFamilia(ByVal nombre As String, i As Long)
    Dim codlocal As String, rs As Recordset, sql As String
    sql = "SELECT idinventario,codinventario FROM ivinventario where descripcion = '" & nombre & "'"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not rs.EOF Then
        grd.TextMatrix(i, grd.ColIndex("CodFamilia")) = rs.Fields("CodInventario")
        grd.TextMatrix(i, grd.ColIndex("Id")) = rs.Fields("Idinventario")
    End If
End Sub

Private Sub AsignarFamilia()
    Dim familia As String, codfamilia As String
    Dim i As Long, s As String
    
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
        For i = .Row To .Rows - 1
            .TextMatrix(i, .ColIndex("Familia")) = familia
            .TextMatrix(i, .ColIndex("CodFamilia")) = codfamilia
            RecuperaCodFamilia familia, i
        Next i
    End With
    End Sub


Private Sub GrabarFamilia()
    Dim i As Long, cod As String, ix As Long, mobjIV As IVinventario
    Dim obj As IVFamiliaDetalle, msg As String
    Dim IdPadre As String, codPadre As String, codhijo As String
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
        
        For i = .FixedRows To .Rows - 1
            DoEvents
            If Len(.TextMatrix(i, .ColIndex("IdRecuperado"))) > 0 Then
                EliminaFila .TextMatrix(i, .ColIndex("IdRecuperado")), .TextMatrix(i, .ColIndex("Código"))
            End If
            ' compara si se ha cambiado de familia
'            If .TextMatrix(i, .ColIndex("id")) <> .TextMatrix(i, .ColIndex("IdRecuperado")) And Len(.TextMatrix(i, .ColIndex("Id"))) > 0 And Len(.TextMatrix(i, .ColIndex("IdRecuperado"))) > 0 Then
'                  msg = "El item " & _
'                            .TextMatrix(i, .ColIndex("Descripción")) & " ya esta asignado a una Familia " & vbCr & vbCr & _
'                            "desea cambiar a la Familia: " & .TextMatrix(i, .ColIndex("Familia"))
'                    If MsgBox(msg, vbYesNo + vbQuestion) = vbYes Then
'                        'elimina de la coleccion
'                        EliminaFila .TextMatrix(i, .ColIndex("IdRecuperado")), .TextMatrix(i, .ColIndex("Código"))
'                    End If
'            End If
                        
            codPadre = .TextMatrix(i, .ColIndex("CodFamilia"))
            Set mobjIV = gobjMain.EmpresaActual.RecuperaIVInventario(codPadre)
            If Not mobjIV Is Nothing Then
            'para que modificado sea true
                mobjIV.Descripcion = mobjIV.Descripcion & "."
                'para dejar como estaba
                mobjIV.Descripcion = Mid$(mobjIV.Descripcion, 1, Len(mobjIV.Descripcion) - 1)
                codhijo = .TextMatrix(i, .ColIndex("Código"))
                ix = mobjIV.AddDetalleFamilia  'Aumenta  item  a la coleccion
                Set obj = mobjIV.RecuperaDetalleFamilia(ix)
                If Not obj Is Nothing Then
                'Si es que se canceló el proceso
                    If mCancelado Then GoTo salida
                    prg1.value = i
                    MensajeStatus i & " de " & .Rows - .FixedRows, vbHourglass
                    obj.CodInventario = codhijo
                    obj.cantidad = 0
                End If
                mobjIV.Grabar
            End If
        Next i
    End With
    Set obj = Nothing
    Set mobjIV = Nothing
salida:
    MensajeStatus
    Set obj = Nothing
    Set mobjIV = Nothing
    Habilitar True
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    GoTo salida
    Exit Sub
End Sub

Private Sub EliminaFila(ByVal idfamilia As Long, ByVal codhijo As String)
    Dim msg As String, r As Long, i As Long
    Dim mobjIV As IVinventario, mobjIVF As IVFamiliaDetalle, rs As Recordset
    On Error GoTo ErrTrap
        'recupero el item
        Set mobjIV = gobjMain.EmpresaActual.RecuperaIVInventario(idfamilia)
        'boy recorrer la coleccion
        For i = 1 To mobjIV.NumFamiliaDetalle
                'recupero un item de la coleccio
                Set mobjIVF = mobjIV.RecuperaDetalleFamilia(i)
                'comparo si es igual al parametro
                If mobjIVF.CodInventario = codhijo Then
                    'elimino de la coleccion
                    mobjIV.RemoveDetalleFamilia (i)
                    'grabo el item
                    mobjIV.Grabar
                    Set mobjIVF = Nothing
                    Set mobjIV = Nothing
                    Exit Sub
                End If
        Next i
    Set mobjIVF = Nothing
    Set mobjIV = Nothing
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub


Private Sub BuscarItemFamilia()
    Static coditem As String, CodAlt As String, _
           Desc As String, _
           codg As String, Numg As Integer, bandIVA As Boolean, bandFraccion As Boolean
    Dim codg1 As String, codg2 As String, codg3 As String, codg4 As String, codg5 As String, codg6 As String
    Dim sql As String, cond As String, rs As Recordset, comodin As String
    On Error GoTo ErrTrap
   
    #If DAOLIB Then
        comodin = "*"
    #Else
        comodin = "%"
    #End If
'    comodin = "%"
    'Abre la pantalla de búsqueda
    frmIVBusqueda.chkIVA.Visible = False
    frmIVBusqueda.cmdAceptar.Top = frmIVBusqueda.Frame1.Height + 650
    frmIVBusqueda.cmdCancelar.Top = frmIVBusqueda.Frame1.Height + 650
    frmIVBusqueda.Height = 3700
    
    If Not frmIVBusqueda.Inicio( _
                coditem, _
                CodAlt, _
                Desc, _
                codg1, codg2, codg3, codg4, codg5, codg6, _
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
    sql = sql & "FROM vwConsIVFamilia INNER JOIN vwIVInventarioRecuperar ON vwConsIVFamilia.IdInventario=vwIVInventarioRecuperar.IdInventario"
    'que no sea padre de familia o receta
    sql = sql & " WHERE vwIVInventarioRecuperar.tipo= '0' AND bandConversionUni = 0"
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
        cond = cond & "((vwIVInventarioRecuperar.CodAlterno1 LIKE '" & CodAlt & comodin & "') " & _
                      "OR (vwIVInventarioRecuperar.CodAlterno2 LIKE '" & CodAlt & comodin & "')) "
    End If
    
    'Descripcion
    If Len(Desc) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(vwConsIVFamilia.Descripcion LIKE '" & Desc & comodin & "') "
    End If
    
    'Grupo
    If Len(codg1) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(vwIVInventarioRecuperar.CodGrupo" & Numg & " = '" & codg1 & "') "
    End If
    If Len(codg2) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(vwIVInventarioRecuperar.CodGrupo" & Numg & " = '" & codg2 & "') "
    End If
    If Len(codg3) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(vwIVInventarioRecuperar.CodGrupo" & Numg & " = '" & codg3 & "') "
    End If
    If Len(codg4) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(vwIVInventarioRecuperar.CodGrupo" & Numg & " = '" & codg4 & "') "
    End If
    
    If Len(codg5) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(vwIVInventarioRecuperar.CodGrupo" & Numg & " = '" & codg5 & "') "
    End If
    
    If Len(codg6) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(vwIVInventarioRecuperar.CodGrupo" & Numg & " = '" & codg6 & "') "
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
Private Function CargarIVGrupos() As Boolean
    Dim s As String
    On Error GoTo ErrTrap
    With grd
        CargarIVGrupos = True
        'fcbGrupoDesde.SetData gobjMain.EmpresaActual.ListaIVGrupo(Numg, False, False)
        
        s = gobjMain.EmpresaActual.ListaIVGrupoParaFlexGrid(1)
        If Len(s) <> 0 Then
            s = Right$(s, Len(s) - 1)
            .ColComboList(IVGRUPO1) = s
        End If
        s = gobjMain.EmpresaActual.ListaIVGrupoParaFlexGrid(2)
        If Len(s) <> 0 Then
            s = Right$(s, Len(s) - 1)
            .ColComboList(IVGRUPO2) = s
        End If
        s = gobjMain.EmpresaActual.ListaIVGrupoParaFlexGrid(3)
        If Len(s) <> 0 Then
            s = Right$(s, Len(s) - 1)
            .ColComboList(IVGRUPO3) = s
        End If
        s = gobjMain.EmpresaActual.ListaIVGrupoParaFlexGrid(4)
        If Len(s) <> 0 Then
            s = Right$(s, Len(s) - 1)
            .ColComboList(IVGRUPO4) = s
        End If
        s = gobjMain.EmpresaActual.ListaIVGrupoParaFlexGrid(5)
        If Len(s) <> 0 Then
            s = Right$(s, Len(s) - 1)
            .ColComboList(IVGRUPO5) = s
        End If
    
        s = gobjMain.EmpresaActual.ListaIVGrupoParaFlexGrid(6)
        If Len(s) <> 0 Then
            s = Right$(s, Len(s) - 1)
            .ColComboList(IVGRUPO6) = s
        End If
    
    End With
    Exit Function
ErrTrap:
        MsgBox "No se han definido IVGrupos", vbInformation
        CargarIVGrupos = False
    Exit Function
End Function

'jeaa 24/09/04 asignacion de grupo a los items
Private Sub AsignarIVGrupos()
    Dim ValorGrupo As String
    Dim i As Long, s As String, j As Integer, grupo As String
    
    With grd
        For j = 1 To 6
            'Obtiene cuentas de la fila actual
            Select Case j
                Case 1
                    ValorGrupo = .TextMatrix(.Row, IVGRUPO1)
                Case 2
                    ValorGrupo = .TextMatrix(.Row, IVGRUPO2)
                Case 3
                    ValorGrupo = .TextMatrix(.Row, IVGRUPO3)
                Case 4
                    ValorGrupo = .TextMatrix(.Row, IVGRUPO4)
                Case 5
                    ValorGrupo = .TextMatrix(.Row, IVGRUPO5)
                Case 6
                    ValorGrupo = .TextMatrix(.Row, IVGRUPO6)
            
            End Select
            grupo = gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(j) & ":  " & ValorGrupo & vbCr
            'Confirma las GRUPOS
            s = "Está seguro que desea asignar los siguientes códigos " & _
                "en todos los ítems que están visualizados?" & vbCr & vbCr & grupo
            If MsgBox(s, vbQuestion + vbYesNo) = vbYes Then
                'Copia a todas las filas los mismos códigos de cuenta
                For i = .FixedRows To .Rows - 1
                    Select Case j
                        Case 1
                            .TextMatrix(i, IVGRUPO1) = ValorGrupo
                        Case 2
                            .TextMatrix(i, IVGRUPO2) = ValorGrupo
                        Case 3
                            .TextMatrix(i, IVGRUPO3) = ValorGrupo
                        Case 4
                            .TextMatrix(i, IVGRUPO4) = ValorGrupo
                        Case 5
                            .TextMatrix(i, IVGRUPO5) = ValorGrupo
                        Case 6
                            .TextMatrix(i, IVGRUPO6) = ValorGrupo
                        
                        End Select
                Next i
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
        If Me.tag = "PCGRUPOS_CLI" Then
            s = gobjMain.EmpresaActual.ListaPCGrupoOrigenParaFlexGrid(1, 2)
        ElseIf Me.tag = "PCGRUPOS_PROV" Then
            s = gobjMain.EmpresaActual.ListaPCGrupoOrigenParaFlexGrid(1, 1)
        ElseIf Me.tag = "PCGRUPOS_EMP" Then
            s = gobjMain.EmpresaActual.ListaPCGrupoOrigenParaFlexGrid(1, 4)
        ElseIf Me.tag = "PCGRUPOS_GAR" Then
            s = gobjMain.EmpresaActual.ListaPCGrupoOrigenParaFlexGrid(1, 3)
            
        Else
            s = gobjMain.EmpresaActual.ListaPCGrupoParaFlexGrid(1)
        End If
        If Len(s) > 1 Then
            s = Right$(s, Len(s) - 1)
            .ColComboList(PCGRUPO1) = s
        End If
'        s = gobjMain.EmpresaActual.ListaPCGrupoParaFlexGrid(2)
        If Me.tag = "PCGRUPOS_CLI" Then
            s = gobjMain.EmpresaActual.ListaPCGrupoOrigenParaFlexGrid(2, 2)
        ElseIf Me.tag = "PCGRUPOS_PROV" Then
            s = gobjMain.EmpresaActual.ListaPCGrupoOrigenParaFlexGrid(2, 1)
        ElseIf Me.tag = "PCGRUPOS_EMP" Then
            s = gobjMain.EmpresaActual.ListaPCGrupoOrigenParaFlexGrid(2, 4)
        ElseIf Me.tag = "PCGRUPOS_GAR" Then
            s = gobjMain.EmpresaActual.ListaPCGrupoOrigenParaFlexGrid(2, 3)
            
        Else
            s = gobjMain.EmpresaActual.ListaPCGrupoParaFlexGrid(2)
        End If

        If Len(s) > 1 Then
            s = Right$(s, Len(s) - 1)
            .ColComboList(PCGRUPO2) = s
        End If
        's = gobjMain.EmpresaActual.ListaPCGrupoParaFlexGrid(3)
        If Me.tag = "PCGRUPOS_CLI" Then
            s = gobjMain.EmpresaActual.ListaPCGrupoOrigenParaFlexGrid(3, 2)
        ElseIf Me.tag = "PCGRUPOS_PROV" Then
            s = gobjMain.EmpresaActual.ListaPCGrupoOrigenParaFlexGrid(3, 1)
        ElseIf Me.tag = "PCGRUPOS_EMP" Then
            s = gobjMain.EmpresaActual.ListaPCGrupoOrigenParaFlexGrid(3, 4)
        ElseIf Me.tag = "PCGRUPOS_GAR" Then
            s = gobjMain.EmpresaActual.ListaPCGrupoOrigenParaFlexGrid(3, 3)
            
        Else
            s = gobjMain.EmpresaActual.ListaPCGrupoParaFlexGrid(3)
        End If
        
        If Len(s) > 1 Then
            s = Right$(s, Len(s) - 1)
            .ColComboList(PCGRUPO3) = s
        End If
        'AUC 03/10/2005
        's = gobjMain.EmpresaActual.ListaPCGrupoParaFlexGrid(4)
        If Me.tag = "PCGRUPOS_CLI" Then
            s = gobjMain.EmpresaActual.ListaPCGrupoOrigenParaFlexGrid(4, 2)
        ElseIf Me.tag = "PCGRUPOS_PROV" Then
            s = gobjMain.EmpresaActual.ListaPCGrupoOrigenParaFlexGrid(4, 1)
        ElseIf Me.tag = "PCGRUPOS_EMP" Then
            s = gobjMain.EmpresaActual.ListaPCGrupoOrigenParaFlexGrid(4, 4)
        ElseIf Me.tag = "PCGRUPOS_GAR" Then
            s = gobjMain.EmpresaActual.ListaPCGrupoOrigenParaFlexGrid(4, 3)
            
        Else
            s = gobjMain.EmpresaActual.ListaPCGrupoParaFlexGrid(4)
        End If
        
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
    Dim i As Long, s As String, j As Integer, grupo As String
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
                For i = grd.Row To .Rows - 1
                    Select Case j
                        Case 1
                            .TextMatrix(i, PCGRUPO1) = ValorGrupo
                        Case 2
                            .TextMatrix(i, PCGRUPO2) = ValorGrupo
                        Case 3
                            .TextMatrix(i, PCGRUPO3) = ValorGrupo
                         Case 4 'auc 03/10/2005
                           .TextMatrix(i, PCGRUPO4) = ValorGrupo
                       End Select
                Next i
            End If
        Next j
    End With
End Sub
'jeaa 13/04/2005
Private Sub AsignarFraccion()
    Dim band As Boolean, i As Integer
    With grd
        'Obtiene Bandera de la fila actual
        band = .TextMatrix(.Row, .ColIndex("Venta por Fraccion"))
        For i = .Row To .Rows - 1
            .TextMatrix(i, .ColIndex("Venta por Fraccion")) = IIf(band, vbChecked, vbUnchecked)
        Next i
    End With
End Sub

'jeaa 13/04/2005
Private Sub AsignarArea()
    Dim band As Boolean, i As Integer
    With grd
        'Obtiene Bandera de la fila actual
        band = .TextMatrix(.Row, .ColIndex("Venta por Area"))
        For i = .Row To .Rows - 1
            .TextMatrix(i, .ColIndex("Venta por Area")) = IIf(band, vbChecked, vbUnchecked)
        Next i
    End With
End Sub

'jeaa 26/12/2005
Private Sub AsignarVenta()
    Dim band As Boolean, i As Integer
    With grd
        'Obtiene Bandera de la fila actual
        band = .TextMatrix(.Row, .ColIndex("Venta"))
        For i = .Row To .Rows - 1
            .TextMatrix(i, .ColIndex("Venta")) = IIf(band, vbChecked, vbUnchecked)
        Next i
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
    If Not frmIVBusqueda.InicioTrans( _
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
    Dim i As Long, s As String
    
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
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("Autorizacion SRI")) = NumSRI
            .TextMatrix(i, .ColIndex("Fecha Caducidad")) = fecha
        Next i
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
    sql = sql & " FROM IvInventario IV inner join IVKardex ivk inner join gncomprobante  gnc on gnc.Transid=ivk.transid"
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
    Dim codg1 As String, codg2 As String, codg3 As String, codg4 As String, codg5 As String, codg6 As String
    Static CodTrans As String, desde As Long, hasta As Long
    Dim sql As String, cond As String, rs As Recordset, comodin As String, codbod As String
    Dim Tipo As Integer
    On Error GoTo ErrTrap
    codbod = ""
    
    #If DAOLIB Then
        comodin = "*"
    #Else
        comodin = "%"
    #End If
'    comodin = "%"
    'Abre la pantalla de búsqueda
    If Not frmIVBusqueda.Inicio( _
                coditem, _
                CodAlt, _
                Desc, _
                codg1, codg2, codg3, codg4, codg5, codg6, _
                Numg, _
                bandIVA, _
                Me.tag, codbod, Tipo) Then
      'if not frmivbusqueda.InicioTrans (
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
    sql = sql & " from    IVGrupo6"
    sql = sql & " RIGHT JOIN (IVGrupo5"
    sql = sql & " RIGHT JOIN (IVGrupo4"
    sql = sql & " RIGHT JOIN (IVGrupo3"
    sql = sql & " RIGHT JOIN (IVGrupo2"
    sql = sql & " RIGHT JOIN (IVGrupo1"
    sql = sql & " RIGHT JOIN IVInventario ivi"
    sql = sql & " inner join ivexist ive"
    sql = sql & " inner join ivbodega ivb"
    sql = sql & " on ive.idbodega=ivb.idbodega"
    sql = sql & " on ivi.idinventario=ive.idinventario"
    sql = sql & " ON IVGrupo1.IdGrupo1 = IVI.IdGrupo1)"
    sql = sql & " ON IVGrupo2.IdGrupo2 = IVI.IdGrupo2)"
    sql = sql & " ON IVGrupo3.IdGrupo3 = IVI.IdGrupo3)"
    sql = sql & " ON IVGrupo4.IdGrupo4 = IVI.IdGrupo4)"
    sql = sql & " ON IVGrupo5.IdGrupo5 = IVI.IdGrupo5)"
    sql = sql & " ON IVGrupo6.IdGrupo6 = IVI.IdGrupo6"

        'CodInventario
    If Len(coditem) > 0 Then
        'If Len(Cond) > 0 Then
        cond = cond & "AND "
        cond = cond & "(CodInventario LIKE '" & coditem & comodin & "') "
    End If
    
    'CodAlterno
    If Len(CodAlt) > 0 Then
        'If Len(Cond) > 0 Then
        cond = cond & "AND "
        cond = cond & "((CodAlterno1 LIKE '" & CodAlt & comodin & "') " & _
                      "OR (CodAlterno2 LIKE '" & CodAlt & comodin & "')) "
    End If
    
    'Descripcion
    If Len(Desc) > 0 Then
        'If Len(Cond) > 0 Then
        cond = cond & "AND "
        cond = cond & "(ivi.Descripcion LIKE '" & Desc & comodin & "') "
    End If
    
    
    
    If Len(codg1) > 0 Then
        'If Len(Cond) > 0 Then
        cond = cond & "AND "
        cond = cond & "(CodGrupo1" & " = '" & codg1 & "') "
    End If
    
    If Len(codg2) > 0 Then
        'If Len(Cond) > 0 Then
        cond = cond & "AND "
        cond = cond & "(CodGrupo2" & " = '" & codg2 & "') "
    End If
    
    If Len(codg3) > 0 Then
        'If Len(Cond) > 0 Then
        cond = cond & "AND "
        cond = cond & "(CodGrupo3" & " = '" & codg3 & "') "
    End If
    
    If Len(codg4) > 0 Then
        'If Len(Cond) > 0 Then
        cond = cond & "AND "
        cond = cond & "(CodGrupo4" & " = '" & codg4 & "') "
    End If
    
    If Len(codg5) > 0 Then
        'If Len(Cond) > 0 Then
        cond = cond & "AND "
        cond = cond & "(CodGrupo5" & " = '" & codg5 & "') "
    End If
    
    If Len(codg6) > 0 Then
        'If Len(Cond) > 0 Then
        cond = cond & "AND "
        cond = cond & "(CodGrupo6" & " = '" & codg6 & "') "
    End If
    
    If Len(codbod) > 0 Then
        'If Len(codbod) > 0 Then
        cond = cond & "AND "
        cond = cond & "(CodBodega" & " = '" & codbod & "') "
    End If
    
   If Tipo <> 0 Then
        'If Len(Cond) > 0 Then Cond = Cond & "AND "
        cond = cond & "AND "
        cond = cond & "(Tipo" & " = " & Tipo & ") "
    End If
   
    If Len(cond) > 0 Then sql = sql & " WHERE 1=1 " & cond
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
    Dim i As Long, s As String
    
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
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("Existencia Mínima")) = min
            .TextMatrix(i, .ColIndex("Existencia Máxima")) = max
        Next i
    End With
End Sub


Private Sub BuscarIvExist()
Static coditem As String, CodAlt As String, _
           Desc As String, _
           codg As String, Numg As Integer, bandIVA As Boolean, bandFraccion As Boolean
    Dim codg1 As String, codg2 As String, codg3 As String, codg4 As String, codg5 As String, codg6 As String
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
    If Not frmIVBusqueda.Inicio( _
                coditem, _
                CodAlt, _
                Desc, _
                codg1, codg2, codg3, codg4, codg5, codg6, _
                Numg, _
                bandIVA, _
                Me.tag, CodBodega) Then
      'if not frmivbusqueda.InicioTrans (
        'Si fue cancelada la busqueda, sale no mas
        grd.SetFocus
        Exit Sub
    End If
    
    'Cambia la forma de cursor
    MensajeStatus MSG_PREPARA, vbHourglass
    
    'Compone la cadena de SQL
    sql = "SELECT"
    sql = sql & " IVI.IdInventario , CodInventario, CodAlterno1, IVI.Descripcion, 0,'" & CodBodega & "',0 "
    sql = sql & " from    IVGrupo6"
    sql = sql & " RIGHT JOIN (IVGrupo5"
    sql = sql & " RIGHT JOIN (IVGrupo4"
    sql = sql & " RIGHT JOIN (IVGrupo3"
    sql = sql & " RIGHT JOIN (IVGrupo2"
    sql = sql & " RIGHT JOIN (IVGrupo1"
    sql = sql & " RIGHT JOIN IVInventario ivi"
    sql = sql & " ON IVGrupo1.IdGrupo1 = IVI.IdGrupo1)"
    sql = sql & " ON IVGrupo2.IdGrupo2 = IVI.IdGrupo2)"
    sql = sql & " ON IVGrupo3.IdGrupo3 = IVI.IdGrupo3)"
    sql = sql & " ON IVGrupo4.IdGrupo4 = IVI.IdGrupo4)"
    sql = sql & " ON IVGrupo5.IdGrupo5 = IVI.IdGrupo5)"
    sql = sql & " ON IVGrupo6.IdGrupo6 = IVI.IdGrupo6"

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
    
    If Len(codg6) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodGrupo6" & " = '" & codg6 & "') "
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
    Dim i As Long, s As String
    
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
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("CodBodega")) = bod
        Next i
    End With
End Sub


Private Sub AsignarPresupuesto()
    Dim Presupuesto As Currency
    Dim i As Long, s As String
    
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
        For i = .Row To .Rows - 1
            .TextMatrix(i, .ColIndex("Presupuesto")) = Presupuesto
        Next i
    End With
End Sub


Private Sub GrabarPresupuesto()
    Dim i As Long, ct As CtCuenta, cod As String
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
        For i = .FixedRows To .Rows - 1
            'Si es que se canceló el proceso
            If mCancelado Then GoTo salida
        
            prg1.value = i
            cod = .TextMatrix(i, .ColIndex("Código"))
            MensajeStatus i & " de " & .Rows - .FixedRows, vbHourglass
            DoEvents
            
            'Recupera el objeto de Inventario
            Set ct = gobjMain.EmpresaActual.RecuperaCTCuenta(cod)
            If Me.tag = "CUENTA_PRESUP" Then
                ct.ValPresupuesto = .ValueMatrix(i, .ColIndex("Presupuesto"))
            ElseIf Me.tag = "CUENTA101" Then
                ct.CampoF101 = .ValueMatrix(i, .ColIndex("Campo F101"))

            Else
                ct.CodCuentaSC = .ValueMatrix(i, .ColIndex("Cuenta SC"))
                ct.Campo101 = .ValueMatrix(i, .ColIndex("Campo F101"))
            End If
            ct.Grabar
        Next i
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
    Dim i As Long
    If MsgBox("Desea Ingrese el valor de Tiempo Reposicion", vbYesNo) = vbYes Then
        s = InputBox("Ingrese el valor de Tiempo Reposicion", "Asignar un valor", "")
        If IsNumeric(s) Then
            v = CSng(s)
        Else
            MsgBox "Debe ingresar un valor numérico. (ejm. 15)", vbInformation
            grd.SetFocus
            Exit Sub
        End If
        
        With grd
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Tiempo Repo")) = v
            Next i
        End With
    End If

    If MsgBox("Desea Ingrese el valor de Frecuencia Reposicion", vbYesNo) = vbYes Then
        s = InputBox("Ingrese el valor de Frecuencia Reposicion", "Asignar un valor", "")
        If IsNumeric(s) Then
            v = CSng(s)
        Else
            MsgBox "Debe ingresar un valor numérico. (ejm. 15)", vbInformation
            grd.SetFocus
            Exit Sub
        End If
        
        With grd
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Frecuencia Repo")) = v
            Next i
        End With
    End If

    If MsgBox("Desea Ingrese el valor de Promedio Ventas", vbYesNo) = vbYes Then
        s = InputBox("Ingrese el valor de Promedio Ventas", "Asignar un valor", "")
        If IsNumeric(s) Then
            v = CSng(s)
        Else
            MsgBox "Debe ingresar un valor numérico. (ejm. 15)", vbInformation
            grd.SetFocus
            Exit Sub
        End If
        
        With grd
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Prom Vta")) = v
            Next i
        End With
    End If

End Sub


Private Sub AsignarPorDescuento()
    Dim ValorDesc As Currency
    Dim i As Long, s As String, j As Integer, grupo As String
    
    With grd
        For j = 1 To 5
            'Obtiene cuentas de la fila actual
            Select Case j
                Case 1
                    ValorDesc = .TextMatrix(.Row, DESC1)
                Case 2
                    ValorDesc = .TextMatrix(.Row, DESC2)
                Case 3
                    ValorDesc = .TextMatrix(.Row, DESC3)
                Case 4
                    ValorDesc = .TextMatrix(.Row, DESC4)
                Case 5
                    ValorDesc = .TextMatrix(.Row, DESC5)
            End Select
            'Confirma las GRUPOS
            s = "Está seguro que desea asignar los % Descuento " & j & _
                " en todos los ítems que están visualizados?" & vbCr & vbCr & grupo
            If MsgBox(s, vbQuestion + vbYesNo) = vbYes Then
                'Copia a todas las filas los mismos códigos de cuenta
                For i = .FixedRows To .Rows - 1
                    Select Case j
                        Case 1
                            .TextMatrix(i, DESC1) = ValorDesc
                        Case 2
                            .TextMatrix(i, DESC2) = ValorDesc
                        Case 3
                            .TextMatrix(i, DESC3) = ValorDesc
                        Case 4
                            .TextMatrix(i, DESC4) = ValorDesc
                        Case 5
                            .TextMatrix(i, DESC5) = ValorDesc
                        
                        End Select
                Next i
            End If
        Next j
    End With
End Sub


Private Sub AsignarPorComision()
    Dim ValorComision As Currency
    Dim i As Long, s As String, j As Integer, grupo As String
    
    With grd
        For j = 1 To 5
            'Obtiene cuentas de la fila actual
            Select Case j
                Case 1
                    ValorComision = .TextMatrix(.Row, DESC1)
                Case 2
                    ValorComision = .TextMatrix(.Row, DESC2)
                Case 3
                    ValorComision = .TextMatrix(.Row, DESC3)
                Case 4
                    ValorComision = .TextMatrix(.Row, DESC4)
                Case 5
                    ValorComision = .TextMatrix(.Row, DESC5)
            End Select
            'Confirma las GRUPOS
            s = "Está seguro que desea asignar los % Comisión " & j & _
                " en todos los ítems que están visualizados?" & vbCr & vbCr & grupo
            If MsgBox(s, vbQuestion + vbYesNo) = vbYes Then
                'Copia a todas las filas los mismos códigos de cuenta
                For i = .FixedRows To .Rows - 1
                    Select Case j
                        Case 1
                            .TextMatrix(i, DESC1) = ValorComision
                        Case 2
                            .TextMatrix(i, DESC2) = ValorComision
                        Case 3
                            .TextMatrix(i, DESC3) = ValorComision
                        Case 4
                            .TextMatrix(i, DESC4) = ValorComision
                        Case 5
                            .TextMatrix(i, DESC5) = ValorComision
                        
                        End Select
                Next i
            End If
        Next j
    End With
End Sub




Private Sub BuscarIvExistNegativa()
Static coditem As String, CodAlt As String, _
           Desc As String, _
           codg As String, Numg As Integer, bandIVA As Boolean, bandFraccion As Boolean
    Dim codg1 As String, codg2 As String, codg3 As String, codg4 As String, codg5 As String, codg6 As String
    Dim CodBodega As String, sum As String
    Static CodTrans As String, desde As Long, hasta As Long
    Dim sql As String, cond As String, rs As Recordset, comodin As String
    Dim fechaIni As Date, fechahasta As Date
    Dim dias As Integer, i As Integer, NumReg As Long, j As Integer
    On Error GoTo ErrTrap
    
    #If DAOLIB Then
        comodin = "*"
    #Else
        comodin = "%"
    #End If
'    comodin = "%"
    'Abre la pantalla de búsqueda
    If Not frmIVBusqueda.Inicio( _
                coditem, _
                CodAlt, _
                Desc, _
                codg1, codg2, codg3, codg4, codg5, codg6, _
                Numg, _
                bandIVA, _
                Me.tag, CodBodega) Then
      'if not frmivbusqueda.InicioTrans (
        'Si fue cancelada la busqueda, sale no mas
        grd.SetFocus
        Exit Sub
    End If
    
    'Cambia la forma de cursor
    MensajeStatus MSG_PREPARA, vbHourglass
    dias = DateDiff("d", gobjMain.EmpresaActual.GNOpcion.FechaInicio, DateAdd("d", -1, CDate("01" & "/" & DatePart("m", DateAdd("m", 1, gobjMain.EmpresaActual.GNOpcion.FechaInicio)) & "/" & DatePart("yyyy", gobjMain.EmpresaActual.GNOpcion.FechaInicio))))
    
    For i = 0 To dias
        sql = " SELECT"
        sql = sql & " IVInventario.idInventario,"
        sql = sql & " IVBodega.idBodega, SUM(IVKardex.Cantidad)*-1 AS Existencia"
        sql = sql & " into tmp" & i
        sql = sql & " From (IVInventario"
        sql = sql & " INNER JOIN (IVBodega"
        sql = sql & " INNER JOIN (IVKardex"
        sql = sql & " INNER JOIN (GNtrans"
        sql = sql & " INNER JOIN GNComprobante"
        sql = sql & " ON GNtrans.Codtrans = GNCOmprobante.Codtrans)"
        sql = sql & " ON IVKardex.transID = GNComprobante.transID)"
        sql = sql & " ON IVBodega.IdBodega = IVKArdex.IdBodega)"
        sql = sql & " ON IVInventario.IdInventario = IVKardex.IdInventario)"
        sql = sql & " WHERE  GNComprobante.FechaTrans <= '" & DateAdd("d", i, gobjMain.EmpresaActual.GNOpcion.FechaInicio) & "'"
        sql = sql & " AND  ((GNtrans.AfectaCantidad) = 1)"
        sql = sql & " AND GNComprobante.Estado <> 3"
        sql = sql & " AND BandServicio = 0"
        sql = sql & " GROUP BY IVInventario.IdInventario,  IVBodega.idBodega"
        sql = sql & " Having Sum(IVKardex.cantidad) < 0"
         VerificaExistenciaTablaTemp "tmp" & i
         gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
    Next i
        sum = ""
        sql = " SELECT"
        sql = sql & " IVInventario.idInventario, IVInventario.CodInventario, IVInventario.CodAlterno1,"
        sql = sql & " IVInventario.Descripcion, IVBodega.idBodega, IVBodega.CodBodega, 0 AS Existencia, "
        For i = 0 To dias
            sql = sql & " tmp" & i & ". existencia as exist" & i & ", "
        Next i
       
       sql = Mid(sql, 1, Len(sql) - 2)
        sql = sql & " into t" & 1
        sql = sql & " From IVexist"
        sql = sql & " INNER JOIN IVBodega"
        sql = sql & " ON IVBodega.IdBodega = IVexist.IdBodega"
        sql = sql & " INNER JOIN IVInventario"
        sql = sql & " ON IVInventario.IdInventario = IVexist.IdInventario"
      For i = 0 To dias
            sql = sql & " left JOIN tmp" & i
            sql = sql & " ON IVInventario.IdInventario = tmp" & i & ".idInventario "
            sql = sql & " and IVexist.idbodega = tmp" & i & ".idbodega "
    Next i
     sql = sql & " where BandServicio = 0"
    VerificaExistenciaTablaTemporal 1
    gobjMain.EmpresaActual.EjecutarSQL sql, NumReg
        sql = " SELECT"
        sql = sql & " idInventario, CodInventario, CodAlterno1,"
        sql = sql & " Descripcion, idbodega, CodBodega,  "
        sql = sql & "0 as Mayor, "
        For i = 0 To dias
            sql = sql & "isnull(exist" & i & ",0), "
        Next i
'        For i = 0 To dias
'            sql = sql & "exist" & i & "+ "
'        Next i
        sql = Mid(sql, 1, Len(sql) - 2)
  '      sql = sql & " as existt "
        sql = sql & " from t1"
        sql = sql & " where ("
        For i = 0 To dias
            sql = sql & "isnull(exist" & i & ",0) + "
        Next i
        sql = Mid(sql, 1, Len(sql) - 2)
        sql = sql & " ) <>0"
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    With grd
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If Not rs.EOF Then .LoadArray MiGetRows(rs)
        ConfigCols
        .Redraw = flexRDBuffered
        .SetFocus
        For i = 1 To grd.Rows - 1
            For j = 8 To grd.Cols - 1
                If j = 8 Then
                    grd.TextMatrix(i, 7) = grd.TextMatrix(i, j)
                Else
                    If grd.ValueMatrix(i, j) >= grd.ValueMatrix(i, 7) Then
                        grd.TextMatrix(i, 7) = grd.TextMatrix(i, j)
                    End If
                End If
            Next j
        Next i
    End With
    
    For j = 8 To grd.Cols - 1
        grd.ColHidden(j) = True
    Next j
    MensajeStatus
    Exit Sub
ErrTrap:
    grd.Redraw = flexRDBuffered
    MensajeStatus
    DispErr
    grd.SetFocus
    Exit Sub
End Sub


Private Sub BuscarCostoReferencial()
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
    sql = "SELECT   IV.codInventario,IV.codAlterno1,IV.Descripcion,ivk.CostoRealTotal/ivk.cantidad as CostoRealTotal,gnc.FechaGrabado, iv.CostoReferencial "
    sql = sql & " FROM IvInventario IV inner join IVKardex ivk inner join gncomprobante  gnc on gnc.Transid=ivk.transid"
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


Private Sub AsignarCostoReferencial()
    Dim s As String, v As Single
    Dim i As Long
    
    s = InputBox("Ingrese el valor de Costo Referencial ", "Asignar un valor", "15")
    If IsNumeric(s) Then
        v = CSng(s)
    Else
        MsgBox "Debe ingresar un valor numérico. (ejm. 1.1544) ", vbInformation
        grd.SetFocus
        Exit Sub
    End If
    
    With grd
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("Costo Referencial")) = v
        Next i
    End With
End Sub

Private Function CopiarProvincia(ByVal Desc As String) As String
Dim sql As String
Dim s As String
Dim rs As Recordset
    sql = "Select codprovincia from pcprovincia where descripcion  = '" & Desc & "'"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    Do While Not rs.EOF
        s = rs!codProvincia
        rs.MoveNext
    Loop
    CopiarProvincia = s
End Function

Private Sub BuscarProvincias(ByVal BandProv As Boolean)
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
    provCli = IIf(BandProv, "Proveedores", "Clientes")
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
            sql = "SELECT CodProvCli, Nombre  " & _
            ", provincia,ciudad,codProvincia,codCanton,codParroquia " & _
            "FROM vwPCProvCli "
    ' si Busca Proveedor o cliente mediante bandera de prov
    If Len(cond) > 0 Then cond = cond & "AND "
    If BandProv Then
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
    CopiarExistentes
    CargarProvincias
   MensajeStatus
    Exit Sub
ErrTrap:
    grd.Redraw = flexRDBuffered
    MensajeStatus
    DispErr
    grd.SetFocus
    Exit Sub
End Sub

Private Sub CopiarExistentes()
Dim i As Long
With grd
    For i = 1 To .Rows - 1
        If Len(.TextMatrix(i, .ColIndex("ProvinciaAnt"))) > 0 Then
            .TextMatrix(i, .ColIndex("ProvinciaNue")) = CopiarProvincia(.TextMatrix(i, .ColIndex("ProvinciaAnt")))
            .TextMatrix(i, .ColIndex("CantonNue")) = CopiarCanton(.TextMatrix(i, .ColIndex("CiudadAnt")))
        End If
    Next
End With
End Sub

Private Function CopiarCanton(ByVal Desc As String) As String
Dim sql As String
Dim s As String
Dim rs As Recordset
On Error GoTo CapturaError
    sql = "Select codCanton from pccanton where descripcion  = '" & Desc & "'"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    Do While Not rs.EOF
        s = rs!codCanton
        rs.MoveNext
    Loop
    CopiarCanton = s
    Exit Function
CapturaError:
    MsgBox Err.Description
    Exit Function
End Function

Private Sub CargarProvincias()
    Dim Provincias As String
    Dim Canton As String
    Dim Parroquia As String
    With grd
        Provincias = gobjMain.EmpresaActual.ListaPCProvinciaParaFlex(True)
        .ColComboList(.ColIndex("Provincia")) = Provincias
   End With
End Sub

Private Sub AsignarProvincias()
    Dim codProvincia As String
    Dim codCanton As String
    Dim codParroquia As String
    Dim i As Long, s As String, j As Integer
    With grd
        For j = 1 To 3
            Select Case j
                Case 1
                    codProvincia = .TextMatrix(.Row, .ColIndex("ProvinciaNue"))
                Case 2
                    codCanton = .TextMatrix(.Row, .ColIndex("CantonNue"))
                Case 3
                    codParroquia = .TextMatrix(.Row, .ColIndex("Parroquia"))
           End Select
        Next j
            'Confirma las GRUPOS
            s = "Está por asignar la misma provincia al resto de filas .. Desea Continuar"
            If MsgBox(s, vbQuestion + vbYesNo) = vbYes Then
                For i = .Row To .Rows - 1
                    .TextMatrix(i, .ColIndex("ProvinciaNue")) = codProvincia
                Next i
            End If
            s = "Está por asingnar el mismo canton al resto de filas.. Desea Continuar"
            If MsgBox(s, vbQuestion + vbYesNo) = vbYes Then
                For i = .Row To .Rows - 1
                    .TextMatrix(i, .ColIndex("CantonNue")) = codCanton
                Next i
            End If
            s = "Está por asingnar la misma Parroquia al resto de filas.. Desea Continuar"
            If MsgBox(s, vbQuestion + vbYesNo) = vbYes Then
                For i = .Row To .Rows - 1
                    .TextMatrix(i, .ColIndex("parroquia")) = codParroquia
                Next i
            End If
    End With
End Sub

Private Sub AsignarVendedor()
    Dim CodVende As String, fecha As String
    Dim i As Long, s As String
    With grd
        'Obtiene cuentas de la fila actual
        CodVende = .TextMatrix(.Row, .ColIndex("CodVendedor"))
        'Confirma las cuentas
        s = "Está seguro que desea asignar los siguientes códigos " & _
            "en todos los ítems que están visualizados?" & vbCr & vbCr & _
            "    Codigo Vendedor:  " & CodVende & vbCr
        If MsgBox(s, vbQuestion + vbYesNo) <> vbYes Then
            .SetFocus
            Exit Sub
        End If
        'Copia a todas las filas los mismos códigos de cuenta
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("CodVendedo")) = CodVende
        Next i
    End With
End Sub

Private Sub BuscarVendedor()
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
    If Not frmIVBusqueda.InicioTrans( _
                CodTrans, _
                desde, hasta) Then
        'Si fue cancelada la busqueda, sale no mas
        grd.SetFocus
        Exit Sub
    End If
    'Cambia la forma de cursor
    MensajeStatus MSG_PREPARA, vbHourglass
    'Compone la cadena de SQL
    sql = "SELECT transid, fechatrans, gnc.codtrans, numtrans , gnc.nombre, codVendedor "
    sql = sql & " FROM gncomprobante gnc inner join fcvendedor fc on gnc.IdVendedor=fc.IdVendedor"
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

Private Function CargarVendedor() As Boolean
    Dim s As String
    On Error GoTo ErrTrap
    With grd
        CargarVendedor = True
        s = gobjMain.EmpresaActual.ListaFCVendedorParaFlex
        s = Right$(s, Len(s) - 1)
        .ColComboList(.ColIndex("CodVendedor")) = s
    End With
    Exit Function
ErrTrap:
        MsgBox "No  hay Vendedores.", vbInformation
        CargarVendedor = False
    Exit Function
End Function


Private Sub BuscarProvinciasEmp()
    Static CodEmp As String, _
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
    
    'Abre la pantalla de búsqueda
    frmPCBusqueda.Caption = "Busqueda de " & provCli ' primero cambia el titulo de ventanda de busqueda
    If Not frmPCBusqueda.Inicio( _
                CodEmp, _
                nombre, _
                codg, _
                Numg) Then
       'Si fue cancelada la busqueda, sale no mas
        grd.SetFocus
        Exit Sub
    End If
    'Cambia la forma de cursor
    MensajeStatus MSG_PREPARA, vbHourglass
            sql = "SELECT CodProvCli, Nombre  " & _
            ", provincia,ciudad,codProvincia,codCanton,codParroquia " & _
            "FROM vwEmpleado vwe INNER JOIN PersonaL p on p.idempleado = vwe.idprovcli  "
    ' si Busca Proveedor o cliente mediante bandera de prov
    If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(BandEmpleado = 1 AND P.bandActivo = 1) "
    
    'CodProveedor/cliente
    If Len(CodEmp) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(codProvCli LIKE '" & CodEmp & comodin & "') "
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
    sql = sql & " ORDER BY nombre "
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
   With grd
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If Not rs.EOF Then .LoadArray MiGetRows(rs)
        ConfigCols
        .Redraw = flexRDBuffered
        .SetFocus
    End With
    CopiarExistentes
    CargarProvincias
   MensajeStatus
    Exit Sub
ErrTrap:
    grd.Redraw = flexRDBuffered
    MensajeStatus
    DispErr
    grd.SetFocus
    Exit Sub
End Sub


Private Sub BuscarEmp()
    Static CodEmp As String, _
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
    
    'Abre la pantalla de búsqueda
    frmPCBusqueda.Caption = "Busqueda de Empleados "
    If Not frmPCBusqueda.Inicio( _
                CodEmp, _
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
        Case "CUENTA_EMP"
           'Compone la cadena de SQL
            sql = "SELECT CodProvcli, Nombre" & _
            ", CodCuentaContable, CodCuentaContable2 " & _
            "FROM vwEmpleado "
        
        Case "PCGRUPOS_EMP"
            'Compone la cadena de SQL
            sql = "SELECT CodProvcli, Nombre" & _
            ", codgrupo1, codgrupo2, codgrupo3 ,codgrupo4 " & _
            "FROM vwEmpleado "
        
        Case "DIVNOMEMP"
            'Compone la cadena de SQL
            sql = "SELECT idempleado, CodProvcli, Nombre, P.papellido, P.pnombre " & _
            "FROM vwEmpleado vw INNER JOIN PERSONAL P ON P.idempleado = vw.idprovcli "
        
        End Select
    ' si Busca Proveedor o cliente mediante bandera de prov
    If Len(cond) > 0 Then cond = cond & "AND "
    
        cond = cond & "(BandEmpleado = 1 ) "
    
    
    'CodProveedor/cliente
    If Len(CodEmp) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(Codprovcli LIKE '" & CodEmp & comodin & "') "
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
    sql = sql & " ORDER BY nombre "
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



Private Sub AsignarArancel()
    Dim arancel As String
    Dim i As Long, s As String
    
    With grd
        'Obtiene cuentas de la fila actual
        arancel = .TextMatrix(.Row, .ColIndex("CodArancel"))
        
        'Confirma las cuentas
        s = "Está seguro que desea asignar los siguientes códigos " & _
            "en todos los ítems que están visualizados?" & vbCr & vbCr & _
            "    Arancel:  " & arancel
        If MsgBox(s, vbQuestion + vbYesNo) <> vbYes Then
            .SetFocus
            Exit Sub
        End If
        
        'Copia a todas las filas los mismos códigos de cuenta
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("CodArancel")) = arancel
        Next i
    End With
End Sub


Private Sub CargarArancel()
    Dim s As String
    With grd
        s = gobjMain.EmpresaActual.ListaArancelParaFlexGrid
        s = Right$(s, Len(s) - 1)
        .ColComboList(.ColIndex("CodArancel")) = s
       
    End With
End Sub

Private Sub BuscarAFExist()
Static coditem As String, CodAlt As String, _
           Desc As String, _
           codg As String, Numg As Integer, bandIVA As Boolean, bandFraccion As Boolean
    Dim codg1 As String, codg2 As String, codg3 As String, codg4 As String, codg5 As String, codg6 As String
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
    If Not frmIVBusqueda.InicioAF( _
                coditem, _
                CodAlt, _
                Desc, _
                codg1, codg2, codg3, codg4, codg5, _
                Numg, _
                bandIVA, _
                Me.tag, CodBodega) Then
      'if not frmivbusqueda.InicioTrans (
        'Si fue cancelada la busqueda, sale no mas
        grd.SetFocus
        Exit Sub
    End If
    
    'Cambia la forma de cursor
    MensajeStatus MSG_PREPARA, vbHourglass
    
    'Compone la cadena de SQL
    sql = "SELECT"
    sql = sql & " af.IdInventario , CodInventario, CodAlterno1, af.Descripcion, 0,pc.codprovcli ,0 "
    sql = sql & " from    AFGrupo5"
    sql = sql & " RIGHT JOIN (AFGrupo4"
    sql = sql & " RIGHT JOIN (AFGrupo3"
    sql = sql & " RIGHT JOIN (AFGrupo2"
    sql = sql & " RIGHT JOIN (AFGrupo1"
    sql = sql & " RIGHT JOIN AFInventario af"
    sql = sql & " ON AFGrupo1.IdGrupo1 = af.IdGrupo1)"
    sql = sql & " ON AFGrupo2.IdGrupo2 = af.IdGrupo2)"
    sql = sql & " ON AFGrupo3.IdGrupo3 = af.IdGrupo3)"
    sql = sql & " ON AFGrupo4.IdGrupo4 = af.IdGrupo4)"
    sql = sql & " ON AFGrupo5.IdGrupo5 = af.IdGrupo5"
    sql = sql & " inner join empleado pc"
    sql = sql & " ON pc.idprovcli = af.Idempleado "
    
If Len(cond) > 0 Then cond = cond & "AND "
cond = cond & " af.idinventario not in"
cond = cond & " ( select idinventario from afexistcustodio afc inner join pcprovcli pc on afc.idprovcli=pc.idprovcli"
cond = cond & " where codprovcli = '" & CodBodega & "')"
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
        cond = cond & "(af.Descripcion LIKE '" & Desc & comodin & "') "
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
    
    If Len(codg6) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodGrupo6" & " = '" & codg6 & "') "
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


Private Sub AsignarAFExist()
    Dim bod As String
    Dim i As Long, s As String
    
    With grd
        bod = .TextMatrix(.Row, .ColIndex("Cod Custodio"))
        
        'Confirma las cuentas
        s = "Está seguro que desea asignar los siguientes códigos " & _
            "en todos los Activos Fijos que están visualizados?" & vbCr & vbCr & _
            "    Bodega:  " & bod
        If MsgBox(s, vbQuestion + vbYesNo) <> vbYes Then
            .SetFocus
            Exit Sub
        End If
        
        'Copia a todas las filas los mismos códigos de cuenta
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("Cod Custodio")) = bod
        Next i
    End With
End Sub


Private Sub AsignarCuentaSC()
    Dim activo As String, costo As String, venta As String
    Dim i As Long, s As String
    
    With grd
        'Obtiene cuentas de la fila actual
        activo = .TextMatrix(.Row, .ColIndex("Cuenta SC"))
        
        'Confirma las cuentas
        s = "Está seguro que desea asignar los siguientes códigos " & _
            "en todos los ítems que están visualizados?" & vbCr & vbCr & _
            "    Cuenta SC:  " & activo
        If MsgBox(s, vbQuestion + vbYesNo) <> vbYes Then
            .SetFocus
            Exit Sub
        End If
        
        'Copia a todas las filas los mismos códigos de cuenta
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("Cuenta SC")) = activo
        Next i
    End With
End Sub


Private Sub CargarCuentasSC()
    Dim s As String
    With grd
        s = gobjMain.EmpresaActual.ListaCTCuentaSCParaFlexGrid(0)
        s = Right$(s, Len(s) - 1)
        .ColComboList(.ColIndex("Cuenta SC")) = s
        
    End With
End Sub

Private Sub CargarTipoDocumentos()
    Dim s As String
    With grd
        s = gobjMain.EmpresaActual.ListaAnexoTipoDocumentoFlexGrid(0)
        s = Right$(s, Len(s) - 1)
        .ColComboList(.ColIndex("Tipo Documento")) = s
        
    End With
End Sub


Private Sub GrabarPCBandRUCValido()
    Dim i As Long, pc As PCProvCli, cod As String, X As Single, bandOk As Boolean
    On Error GoTo ErrTrap
    
    'Confirmación
    If MsgBox("Está seguro que desea Verificar y grabar?", vbQuestion + vbYesNo) <> vbYes Then
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
        bandOk = False
        For i = .FixedRows To .Rows - 1
            bandOk = False
            'Si es que se canceló el proceso
            If mCancelado Then GoTo salida
        
            prg1.value = i
            grd.Row = i
            X = grd.CellTop
            cod = .TextMatrix(i, .ColIndex("Código"))
            MensajeStatus i & " de " & .Rows - .FixedRows, vbHourglass
            DoEvents
            
            
            'Recupera el objeto de Inventario
            Set pc = gobjMain.EmpresaActual.RecuperaPCProvCli(cod)
'            If Len(.TextMatrix(i, .ColIndex("RUC"))) = 13 And (.TextMatrix(i, .ColIndex("Tipo Documento"))) = "C" Then MsgBox "hola"
            
            Select Case Len(.TextMatrix(i, .ColIndex("RUC")))
                Case 13
                    If .TextMatrix(i, .ColIndex("RUC")) <> "9999999999999" Then
                        pc.codtipoDocumento = "R"
                        grd.TextMatrix(i, .ColIndex("Tipo Documento")) = "R"
                    ElseIf .TextMatrix(i, .ColIndex("RUC")) = "9999999999999" Then
                        pc.codtipoDocumento = "F"
                        grd.TextMatrix(i, .ColIndex("Tipo Documento")) = "F"
                    Else
                        pc.codtipoDocumento = ""
                        grd.TextMatrix(i, .ColIndex("Tipo Documento")) = ""
                        grd.TextMatrix(i, .ColIndex("Verificado")) = "Error Tipo Documento"
                    End If
                Case 10
                    If Mid$(.TextMatrix(i, .ColIndex("RUC")), 1, 3) <> "019" And Mid$(.TextMatrix(i, .ColIndex("RUC")), 1, 3) <> "013" And Mid$(.TextMatrix(i, .ColIndex("RUC")), 1, 3) <> "099" Then
                        If .TextMatrix(i, .ColIndex("Tipo Documento")) = "C" Then
                            pc.codtipoDocumento = "C"
                            grd.TextMatrix(i, .ColIndex("Tipo Documento")) = "C"
                        Else
                            pc.codtipoDocumento = ""
                            grd.TextMatrix(i, .ColIndex("Tipo Documento")) = ""
                            grd.TextMatrix(i, .ColIndex("Verificado")) = "Error Tipo Documento"
                        End If
                    End If
                Case Is < 10
                    If Len(.TextMatrix(i, .ColIndex("RUC"))) > 0 Then
                        If .TextMatrix(i, .ColIndex("Tipo Documento")) = "P" Then
                            pc.codtipoDocumento = "P"
                            grd.TextMatrix(i, .ColIndex("Tipo Documento")) = "P"
                        Else
                            pc.codtipoDocumento = ""
                            grd.TextMatrix(i, .ColIndex("Tipo Documento")) = ""
                            grd.TextMatrix(i, .ColIndex("Verificado")) = "Error Tipo Documento"
                        End If
                    
                    Else
                        grd.TextMatrix(i, .ColIndex("Verificado")) = "Error CI/RUC"
                    End If
                Case Else
            End Select
            
            
            
            
            
            If pc.VerificaRUC(.TextMatrix(i, .ColIndex("RUC"))) Then
                If .TextMatrix(i, .ColIndex("Tipo Documento")) = "F" Or .TextMatrix(i, .ColIndex("Tipo Documento")) = "C" Or .TextMatrix(i, .ColIndex("Tipo Documento")) = "R" Then
                    Select Case Len(.TextMatrix(i, .ColIndex("RUC")))
                    Case 13
                        If .TextMatrix(i, .ColIndex("Tipo Documento")) = "R" And .TextMatrix(i, .ColIndex("RUC")) <> "9999999999999" Then
                            pc.BandRUCValido = True
                            bandOk = True
                        ElseIf .TextMatrix(i, .ColIndex("Tipo Documento")) = "F" And .TextMatrix(i, .ColIndex("RUC")) = "9999999999999" Then
                                pc.BandRUCValido = True
                                bandOk = True
                        ElseIf .TextMatrix(i, .ColIndex("Tipo Documento")) <> "R" And .TextMatrix(i, .ColIndex("RUC")) <> "9999999999999" Then
                            .TextMatrix(i, .ColIndex("Tipo Documento")) = "R"
                            pc.codtipoDocumento = "R"
                            pc.BandRUCValido = True
                            bandOk = True
                        Else
                            grd.TextMatrix(i, .ColIndex("Verificado")) = "Error Tipo Documento"
                            bandOk = False
                                    
                        End If
                    Case 10
                        If .TextMatrix(i, .ColIndex("Tipo Documento")) = "C" Then
                            If Mid$(.TextMatrix(i, .ColIndex("RUC")), 1, 3) = "019" Then
                                grd.TextMatrix(i, .ColIndex("Verificado")) = "Error CI/RUC "
                                bandOk = False
                            Else
                                pc.BandRUCValido = True
                                bandOk = True
                            End If
                        Else
                            grd.TextMatrix(i, .ColIndex("Verificado")) = "Error Tipo Documento"
                            bandOk = False
                        End If
                    Case Else
                            grd.TextMatrix(i, .ColIndex("Verificado")) = "Error Tipo Documento"
                            bandOk = False
                        
                    End Select
                ElseIf Len(.TextMatrix(i, .ColIndex("Tipo Documento"))) = 0 Then
                    Select Case Len(.TextMatrix(i, .ColIndex("RUC")))
                    Case 13
                        .TextMatrix(i, .ColIndex("Tipo Documento")) = "R"
                        pc.codtipoDocumento = "R"
                    Case 10
                        .TextMatrix(i, .ColIndex("Tipo Documento")) = "C"
                        pc.codtipoDocumento = "C"
                    End Select
                    If pc.VerificaRUC(.TextMatrix(i, .ColIndex("RUC"))) Then
                            pc.BandRUCValido = True
                            bandOk = True
                    End If
                ElseIf .TextMatrix(i, .ColIndex("Tipo Documento")) = "T" Or .TextMatrix(i, .ColIndex("Tipo Documento")) = "O" Or .TextMatrix(i, .ColIndex("Tipo Documento")) = "P" Then
                    Select Case Len(.TextMatrix(i, .ColIndex("RUC")))
                    Case 13
                        .TextMatrix(i, .ColIndex("Tipo Documento")) = "R"
                        pc.codtipoDocumento = "R"
                    Case 10
                        .TextMatrix(i, .ColIndex("Tipo Documento")) = "C"
                        pc.codtipoDocumento = "C"
                    End Select
                    If pc.VerificaRUC(.TextMatrix(i, .ColIndex("RUC"))) Then
                            pc.BandRUCValido = True
                            bandOk = True
                    End If
                
                Else
                        grd.TextMatrix(i, .ColIndex("Verificado")) = "Error Tipo Documento"
                        bandOk = False
                End If
                
                
                If bandOk Then
                    pc.nombre = grd.TextMatrix(i, .ColIndex("Nombre"))
                    pc.Grabar
                    grd.TextMatrix(i, .ColIndex("Verificado")) = "OK."
                End If
            Else
                If .TextMatrix(i, .ColIndex("Tipo Documento")) = "P" Then
                        pc.BandRUCValido = True
                        pc.Grabar
                        grd.TextMatrix(i, .ColIndex("Verificado")) = "OK."
                        
                Else
                    grd.TextMatrix(i, .ColIndex("Verificado")) = "Error CI/RUC "
                    pc.Estado = 2
                    pc.Grabar
                End If
            End If
            
            
            
        Next i
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



Private Sub GrabarFlujoEfectivo()
    Dim i As Long, ct As CtCuenta, cod As String
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
        For i = .FixedRows To .Rows - 1
            'Si es que se canceló el proceso
            If mCancelado Then GoTo salida
        
            prg1.value = i
            cod = .TextMatrix(i, .ColIndex("Código"))
            MensajeStatus i & " de " & .Rows - .FixedRows, vbHourglass
            DoEvents
            
            'Recupera el objeto de Inventario
            Set ct = gobjMain.EmpresaActual.RecuperaCTCuenta(cod)
            ct.CodCuentaFE = .TextMatrix(i, .ColIndex("Cuenta FE"))
            ct.Grabar
        Next i
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

Private Sub CargarCuentasFE()
    Dim s As String
    With grd
        s = gobjMain.EmpresaActual.ListaCTCuentaFEParaFlexGrid(0)
        s = Right$(s, Len(s) - 1)
        .ColComboList(.ColIndex("Cuenta FE")) = s
        
    End With
End Sub

Private Sub AsignarCuentaFE()
    Dim activo As String, costo As String, venta As String
    Dim i As Long, s As String
    
    With grd
        'Obtiene cuentas de la fila actual
        activo = .TextMatrix(.Row, .ColIndex("Cuenta FE"))
        
        'Confirma las cuentas
        s = "Está seguro que desea asignar los siguientes códigos " & _
            "en todos los ítems que están visualizados?" & vbCr & vbCr & _
            "    Cuenta FE:  " & activo
        If MsgBox(s, vbQuestion + vbYesNo) <> vbYes Then
            .SetFocus
            Exit Sub
        End If
        
        'Copia a todas las filas los mismos códigos de cuenta
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("Cuenta FE")) = activo
        Next i
    End With
End Sub


Private Sub BuscarFormaPagoSRI()
    Static CodTrans As String, desde As Date, hasta As Date
    Dim sql As String, cond As String, rs As Recordset, comodin As String, CodProv As String
    On Error GoTo ErrTrap
    'If Me.tag <> "CUENTA" Then Exit Sub
    
    #If DAOLIB Then
        comodin = "*"
    #Else
        comodin = "%"
    #End If
'    comodin = "%"
    'Abre la pantalla de búsqueda
    If Not frmIVBusqueda.InicioFormaPagoSRI( _
                CodProv, CodTrans, _
                desde, hasta) Then
        'Si fue cancelada la busqueda, sale no mas
        grd.SetFocus
        Exit Sub
    End If
    
    'Cambia la forma de cursor
    MensajeStatus MSG_PREPARA, vbHourglass
    
    'Compone la cadena de SQL
    
    sql = "select gnc.transid, fechatrans, codtrans, numtrans, p.nombre, p.ruc, NumSerieEstablecimiento ,NumSeriePunto,NumSecuencial, a.codcredtrib, a.codformapagoSRI"
    sql = sql & " from gncomprobante gnc inner join anexos a"
    sql = sql & " on gnc.transid= a.transid"
    sql = sql & " inner join pcprovcli p"
    sql = sql & " on gnc.idproveedorref = p.idprovcli"
'    sql = sql & "where g.estado<>3 and  codtrans='cpsyn'"
    
    'CodInventario
    If Len(CodTrans) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(gnc.Codtrans = '" & CodTrans & "') "
    End If
    
    If Len(CodProv) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
'        cond = cond & "(p.Codprovcli= '" & CodProv & "') "
        cond = cond & " (a.codformapagoSRI='" & CodProv & "') "
    End If
    
    'CodAlterno
    'If desde <> 0 And hasta <> 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "fechatrans between '" & desde & "' and '" & hasta & "'"
    'End If
    
    
    
    
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


Private Sub AsignarFormaPagoSRI()
    Dim NumSRI As String, fecha As String
    Dim i As Long, s As String
    
    With grd
        'Obtiene cuentas de la fila actual
        NumSRI = .TextMatrix(.Row, .ColIndex("Nueva Forma Pago SRI"))
        'fecha = .TextMatrix(.Row, .ColIndex("Fecha Caducidad"))
        
        'Confirma las cuentas
        s = "Está seguro que desea asignar los siguientes códigos " & _
            "en todos los ítems que están visualizados?" & vbCr & vbCr & _
            "    Forma Pago SRI:  " & NumSRI & vbCr & _
            "    :   "
        If MsgBox(s, vbQuestion + vbYesNo) <> vbYes Then
            .SetFocus
            Exit Sub
        End If
        
        'Copia a todas las filas los mismos códigos de cuenta
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("Nueva Forma Pago SRI")) = NumSRI
            '.TextMatrix(i, .ColIndex("Fecha Caducidad")) = fecha
        Next i
    End With

End Sub


Private Function CargarFormaCobroSRI() As Boolean
    Dim s As String
    On Error GoTo ErrTrap
        With grd
            CargarFormaCobroSRI = True
            s = gobjMain.EmpresaActual.ListaAnexoFormaPagoParaFlexGrid(0)
            s = Right$(s, Len(s) - 1)
            .ColComboList(.ColIndex("Nueva Forma Pago SRI")) = s
        End With
        Exit Function
ErrTrap:
        MsgBox "No se han definido Forma Pago SRI", vbInformation
        CargarFormaCobroSRI = False
    Exit Function
End Function

Private Sub BuscarDINARDAP()
    Static CodTrans As String, desde As Long, hasta As Long, fechadesde As Date, fechahasta As Date
    Dim sql As String, cond As String, rs As Recordset, comodin As String, ruc As String
    On Error GoTo ErrTrap
    'If Me.tag <> "CUENTA" Then Exit Sub
    
    #If DAOLIB Then
        comodin = "*"
    #Else
        comodin = "%"
    #End If
'    comodin = "%"
    'Abre la pantalla de búsqueda
    If Not frmIVBusqueda.InicioTransNew( _
                CodTrans, _
                desde, hasta, ruc, fechadesde, fechahasta) Then
        'Si fue cancelada la busqueda, sale no mas
        grd.SetFocus
        Exit Sub
    End If
    
    'Cambia la forma de cursor
    MensajeStatus MSG_PREPARA, vbHourglass
    
    'Compone la cadena de SQL
    sql = "SELECT  distinct codprovcli, ruc, pc.nombre, "
    sql = sql & " pcprov.CodProvincia, "
    sql = sql & " pcprov.descripcion, "
    sql = sql & " pccan.codcanton,"
    sql = sql & " pccan.Descripcion , "
    sql = sql & " pcparr.codparroquia, "
    sql = sql & " pcparr.descripcion, "
    sql = sql & " Tiposujeto, sexo, estadocivil, Origeningresos "
    sql = sql & " FROM gncomprobante gnc inner join pcprovcli pc "
    sql = sql & " left join pcprovincia pcprov on pc.idprovincia = pcprov.idprovincia "
    sql = sql & " left join pccanton pccan on pc.idcanton = pccan.idcanton "
    sql = sql & " left join pcparroquia pcparr on pc.idParroquia = pcparr.idparroquia "
    
    sql = sql & " on gnc.IdClienteref=pc.idprovcli"
    'CodInventario
    If Len(CodTrans) > 0 Then
        'If Len(Cond) > 0 Then Cond = Cond & "AND "
        cond = cond & "and (gnc.Codtrans = '" & CodTrans & "') "
    End If
    
    
    If Len(ruc) > 0 Then
        'If Len(ruc) > 0 Then Cond = Cond & "AND "
        cond = cond & "and (ruc like '" & ruc & "%') "
    End If
    
    'CodAlterno
    If desde <> 0 And hasta <> 0 Then
        'If Len(Cond) > 0 Then Cond = Cond & "AND "
        cond = cond & " and numtrans between " & desde & " and " & hasta
    End If
    
    
    
    
    If Len(cond) > 0 Then sql = sql & " WHERE 1=1 " & cond
    sql = sql & " and gnc.fechatrans between '" & fechadesde & "' And  '" & fechahasta & "'"
    sql = sql & " ORDER BY pc.nombre "
    
    
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    With grd
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If Not rs.EOF Then .LoadArray MiGetRows(rs)
        ConfigCols
        .Redraw = flexRDBuffered
        .SetFocus
    End With
    CargarProvincias
    CargarTipoSujeto
    MensajeStatus
    Exit Sub
ErrTrap:
    grd.Redraw = flexRDBuffered
    MensajeStatus
    DispErr
    grd.SetFocus
    Exit Sub
End Sub

Private Sub CargarTipoSujeto()
    Dim Tiposujeto As String, sexo  As String, EstadoCivil As String, OrigenIngresos As String
    With grd
        Tiposujeto = gobjMain.EmpresaActual.ListaTipoSujetoFlex()
        .ColComboList(.ColIndex("Tipo Sujeto")) = Tiposujeto
        
        sexo = gobjMain.EmpresaActual.ListaSexoFlex()
        .ColComboList(.ColIndex("Sexo")) = sexo
        
        EstadoCivil = gobjMain.EmpresaActual.ListaEstadoCivilFlex()
        .ColComboList(.ColIndex("Estado Civil")) = EstadoCivil
        
        OrigenIngresos = gobjMain.EmpresaActual.ListaOrigenIngresosFlex()
        .ColComboList(.ColIndex("Origen Ingresos")) = OrigenIngresos
        
        
   End With
End Sub


'jeaa 24/09/04 asignacion de grupo a los items
Private Sub AsignarDINARDAP()
    Dim NumSRI As String, fecha As String
    Dim i As Long, s As String
    Dim Prov As String, Canton As String, parro As String, Tipo As String, Sex As String, estao As String, origen As String
    Dim cad As String, campo As String, campofin As String, j As Integer
    With grd
        
        For j = 1 To 7
            Select Case j
                Case 1
                    cad = .TextMatrix(.Row, .ColIndex("Provincia"))
                    campo = "Provincia"
                Case 2
                    cad = .TextMatrix(.Row, .ColIndex("Canton"))
                    campo = "Canton"
                Case 3
                    cad = .TextMatrix(.Row, .ColIndex("Parroquia"))
                    campo = "Parroquia"
                Case 4
                    cad = .TextMatrix(.Row, .ColIndex("Tipo Sujeto"))
                    campo = "Tipo Sujeto"
                Case 5
                    cad = .TextMatrix(.Row, .ColIndex("Sexo"))
                    campo = "Sexo"
                Case 6
                    cad = .TextMatrix(.Row, .ColIndex("Estado Civil"))
                    campo = "Estado Civil"
                Case 7
                    cad = .TextMatrix(.Row, .ColIndex("Origen Ingresos"))
                    campo = "Origen Ingresos"
                End Select
            
            'Confirma las cuentas
            campofin = campo & ": " & cad
            s = "Está seguro que desea asignar los siguientes códigos " & _
                "en todos los ítems que están visualizados?" & vbCr & vbCr & campofin
                
            If MsgBox(s, vbQuestion + vbYesNo) = vbYes Then
                .SetFocus
                'Copia a todas las filas los mismos códigos de cuenta
                For i = .FixedRows To .Rows - 1
                    Select Case j
                        Case 1
                            .TextMatrix(i, .ColIndex("Provincia")) = cad
                            RecuperaProvincia grd.TextMatrix(i, grd.ColIndex("Provincia")), i
                        Case 2
                            .TextMatrix(i, .ColIndex("Canton")) = cad
                            RecuperaCanton grd.TextMatrix(i, grd.ColIndex("Canton")), i
                        Case 3
                            .TextMatrix(i, .ColIndex("Parroquia")) = cad
                            RecuperaParroquia grd.TextMatrix(i, grd.ColIndex("Parroquia")), i
                        Case 4
                            .TextMatrix(i, .ColIndex("Tipo Sujeto")) = cad
                        Case 5
                            .TextMatrix(i, .ColIndex("Sexo")) = cad
                        Case 6
                            .TextMatrix(i, .ColIndex("Estado Civil")) = cad
                        Case 7
                            .TextMatrix(i, .ColIndex("Origen Ingresos")) = cad
                        End Select
                Next i
            End If
        Next j
    End With

End Sub


Private Sub RecuperaProvincia(ByVal codigo As String, i As Long)
    Dim codlocal As String, rs As Recordset, sql As String
    
    sql = "SELECT descripcion FROM pcprovincia where CodProvincia = '" & codigo & "'"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not rs.EOF Then
        grd.TextMatrix(i, grd.ColIndex("Desc. Prov")) = rs.Fields("descripcion")
    End If
    Set rs = Nothing

End Sub


Private Sub RecuperaCanton(ByVal codigo As String, i As Long)
    Dim codlocal As String, rs As Recordset, sql As String
    
   
    sql = "SELECT descripcion FROM pccanton where Codcanton = '" & codigo & "'"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not rs.EOF Then
        grd.TextMatrix(i, grd.ColIndex("Desc. Canton")) = rs.Fields("descripcion")
    End If
    
    Set rs = Nothing

End Sub


Private Sub RecuperaParroquia(ByVal codigo As String, i As Long)
    Dim codlocal As String, rs As Recordset, sql As String
   
    sql = "SELECT descripcion FROM pcparroquia where CodParroquia = '" & codigo & "'"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not rs.EOF Then
        grd.TextMatrix(i, grd.ColIndex("Desc. Parroq")) = rs.Fields("descripcion")
    End If
    Set rs = Nothing

End Sub



Private Sub AsignarApellidoNombre()
    Dim apellido As String, nombre As String
    Dim i As Long, s As String, v As Variant
    
    With grd
        'Obtiene cuentas de la fila actual
        For i = 1 To grd.Rows - 1
         v = Split(.TextMatrix(i, .ColIndex("Nombre Completo")), " ")
         apellido = ""
         nombre = ""
        If UBound(v) > 0 Then
            apellido = v(0)
            If UBound(v) > 0 Then
                apellido = apellido + " " + v(1)
                If UBound(v) > 1 Then
                    nombre = v(2)
                    If UBound(v) > 2 Then
                        nombre = nombre + " " + v(3)
                    End If
                End If
            End If
        End If
        
        'Confirma las cuentas
'        s = "Está seguro que desea asignar los siguientes códigos " & _
'            "en todos los ítems que están visualizados?" & vbCr & vbCr & _
'            "    Arancel:  " & arancel
'        If MsgBox(s, vbQuestion + vbYesNo) <> vbYes Then
'            .SetFocus
'            Exit Sub
'        End If
        
        'Copia a todas las filas los mismos códigos de cuenta
        'For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("Apellido")) = apellido
            .TextMatrix(i, .ColIndex("Nombre")) = nombre
        Next i
        
    End With
End Sub


Private Sub AsignarCuenta101()
    Dim activo As String, costo As String, venta As String
    Dim i As Long, s As String
    
    With grd
        'Obtiene cuentas de la fila actual
        activo = .TextMatrix(.Row, .ColIndex("Campo F101"))
        
        'Confirma las cuentas
        s = "Está seguro que desea asignar los siguientes códigos " & _
            "en todos las cuentas que están visualizados hacia abajo?" & vbCr & vbCr & _
            "    Campo F 101:  " & activo
            
        If MsgBox(s, vbQuestion + vbYesNo) <> vbYes Then
            .SetFocus
            Exit Sub
        End If
        
        'Copia a todas las filas los mismos códigos de cuenta
        For i = .Row To .Rows - 1
            .TextMatrix(i, .ColIndex("Campo F101")) = activo
        Next i
    End With
End Sub


Private Sub GrabarEmp()
    Dim i As Long, cod As String, emp As PCProvCli
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
        For i = .FixedRows To .Rows - 1
            'Si es que se canceló el proceso
            If mCancelado Then GoTo salida
        
            prg1.value = i
            cod = .TextMatrix(i, .ColIndex("Código"))
            MensajeStatus i & " de " & .Rows - .FixedRows, vbHourglass
            DoEvents
            
            'Recupera el objeto de Inventario
            Set emp = gobjMain.EmpresaActual.RecuperaEmpleado(cod)
            Select Case Me.tag
                
                    
                Case "CUENTA_EMP"
                    Set emp = gobjMain.EmpresaActual.RecuperaEmpleado(cod)
                    If emp.CodCuentaContable <> .TextMatrix(i, .ColIndex("Cuenta Contable1")) Then
                        emp.CodCuentaContable = .TextMatrix(i, .ColIndex("Cuenta Contable1"))
                    End If
                    
                    If emp.CodCuentaContable2 <> .TextMatrix(i, .ColIndex("Cuenta Contable2")) Then
                        emp.CodCuentaContable2 = .TextMatrix(i, .ColIndex("Cuenta Contable2"))
                    End If
                Case "PCGRUPOS_EMP"
                    If emp.CodGrupo1 <> .TextMatrix(i, PCGRUPO1) Then
                       emp.CodGrupo1 = .TextMatrix(i, PCGRUPO1)
                    End If
                    If emp.CodGrupo2 <> .TextMatrix(i, PCGRUPO2) Then
                        emp.CodGrupo2 = .TextMatrix(i, PCGRUPO2)
                    End If
                    If emp.CodGrupo3 <> .TextMatrix(i, PCGRUPO3) Then
                        emp.CodGrupo3 = .TextMatrix(i, PCGRUPO3)
                    End If
                    'AUC 03/10/2005
                  If emp.CodGrupo4 <> .TextMatrix(i, PCGRUPO4) Then
                        emp.CodGrupo4 = .TextMatrix(i, PCGRUPO4)
                  End If
                
                Case "PROVINCIASEMP"
                        If Len(.TextMatrix(i, .ColIndex("Provincia"))) > 0 Then emp.codProvincia = .TextMatrix(i, .ColIndex("Provincia"))
                        If Len(.TextMatrix(i, .ColIndex("Canton"))) > 0 Then emp.codCanton = .TextMatrix(i, .ColIndex("Canton"))
                        If Len(.TextMatrix(i, .ColIndex("Parroquia"))) > 0 Then emp.codParroquia = .TextMatrix(i, .ColIndex("Parroquia"))
                        emp.Provincia = "" 'ENCERO LO ANTERIOR
                        emp.Ciudad = "" 'ENCERO LO ANTERIOR
            End Select
            emp.GrabarEmpleado
        Next i
    End With
    
salida:
    MensajeStatus
    Set emp = Nothing
    Habilitar True
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    GoTo salida
    Exit Sub
End Sub

Private Function CargarEmpGrupos() As Boolean
    Dim s As Variant
    On Error GoTo ErrTrap
    With grd
        CargarEmpGrupos = True
        s = gobjMain.EmpresaActual.ListaPCGrupoOrigenParaFlexGrid(1, 4)
        If Len(s) > 1 Then
            s = Right$(s, Len(s) - 1)
            .ColComboList(PCGRUPO1) = s
        End If
        s = gobjMain.EmpresaActual.ListaPCGrupoOrigenParaFlexGrid(2, 4)
        If Len(s) > 1 Then
            s = Right$(s, Len(s) - 1)
            .ColComboList(PCGRUPO2) = s
        End If
        s = gobjMain.EmpresaActual.ListaPCGrupoOrigenParaFlexGrid(3, 4)
        If Len(s) > 1 Then
            s = Right$(s, Len(s) - 1)
            .ColComboList(PCGRUPO3) = s
        End If
        'AUC 03/10/2005
        s = gobjMain.EmpresaActual.ListaPCGrupoOrigenParaFlexGrid(4, 4)
        If Len(s) > 0 Then
            s = Right$(s, Len(s) - 1)
            .ColComboList(PCGRUPO4) = s
        End If
    End With
    Exit Function
ErrTrap:
        MsgBox "No se han definido PCGrupos", vbInformation
        CargarEmpGrupos = False
    Exit Function
End Function

Private Sub BuscarComprobanteRelacionado()
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
    If Not frmIVBusqueda.InicioTransRelacion( _
                CodTrans, CodTransRel, _
                desde, hasta) Then
        'Si fue cancelada la busqueda, sale no mas
        grd.SetFocus
        Exit Sub
    End If
    
    'Cambia la forma de cursor
    MensajeStatus MSG_PREPARA, vbHourglass
    
    'Compone la cadena de SQL
    sql = "SELECT gnc.transid, gnc.fechatrans, gnc.codtrans, gnc.numtrans , gnc.NumDocRef, gnc.descripcion,  gncr.codtrans + '-' + convert(varchar,gncr.numtrans) "
    sql = sql & " FROM gncomprobante gnc "
    sql = sql & " left join gncomprobante gncr on gnc.idtransFuente=gncr.transid"
'    sql = sql & " inner join gntrans gnt on gnc.codtrans=gnt.codtrans"
    
    'CodInventario
    If Len(CodTrans) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(gnc.Codtrans = '" & CodTrans & "') "
    End If
    
    'CodAlterno
    If desde <> 0 And hasta <> 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "gnc.numtrans between " & desde & " and " & hasta
    End If
    
    
    
    
    If Len(cond) > 0 Then sql = sql & " WHERE " & cond
    sql = sql & " ORDER BY gnc.Numtrans "
    
    
    
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


Private Sub CargarComprobantes()
    Dim s As String
    With grd
        If Len(CodTransRel) > 0 Then
            s = gobjMain.EmpresaActual.ListaGnComprobanteFlexGrid(CodTransRel)
            s = Right$(s, Len(s) - 1)
        End If
        .ColComboList(.ColIndex("Comprobante Relacionado")) = s
    
        
    End With
End Sub

Private Sub BuscarEmpleado()
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
    If Not frmIVBusqueda.InicioTrans( _
                CodTrans, _
                desde, hasta) Then
        'Si fue cancelada la busqueda, sale no mas
        grd.SetFocus
        Exit Sub
    End If
    'Cambia la forma de cursor
    MensajeStatus MSG_PREPARA, vbHourglass
    'Compone la cadena de SQL
    sql = "SELECT transid, fechatrans, gnc.codtrans, numtrans , gnc.nombre "
    sql = sql & " FROM gncomprobante gnc" 'inner join fcvendedor fc on gnc.IdVendedor=fc.IdVendedor"
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


Private Function CargarEmpleado() As Boolean
    Dim s As String
    On Error GoTo ErrTrap
    With grd
        CargarEmpleado = True
        s = gobjMain.EmpresaActual.ListaEmpleadoFlexGrid
        s = Right$(s, Len(s) - 1)
        .ColComboList(.ColIndex("CodEmpleado")) = s
    End With
    Exit Function
ErrTrap:
        MsgBox "No  hay Empleados.", vbInformation
        CargarEmpleado = False
    Exit Function
End Function

Private Sub AsignarEmpleado()
    Dim CodEmpleado As String, fecha As String
    Dim i As Long, s As String
    With grd
        'Obtiene cuentas de la fila actual
        CodEmpleado = .TextMatrix(.Row, .ColIndex("CodEmpleado"))
        'Confirma las cuentas
        s = "Está seguro que desea asignar los siguientes códigos " & _
            "en todos los ítems que están visualizados?" & vbCr & vbCr & _
            "    Codigo Empleado:  " & CodEmpleado & vbCr
        If MsgBox(s, vbQuestion + vbYesNo) <> vbYes Then
            .SetFocus
            Exit Sub
        End If
        'Copia a todas las filas los mismos códigos de cuenta
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("CodEmpleado")) = CodEmpleado
        Next i
    End With
End Sub

Private Sub GrabarPCemail()
    Dim i As Long, pc As PCProvCli, cod As String, X As Single, bandOk As Boolean
    On Error GoTo ErrTrap
    
    'Confirmación
    If MsgBox("Está seguro que desea Verificar y grabar?", vbQuestion + vbYesNo) <> vbYes Then
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
        bandOk = False
        For i = .FixedRows To .Rows - 1
            bandOk = False
            'Si es que se canceló el proceso
            If mCancelado Then GoTo salida
        
            prg1.value = i
            grd.Row = i
            X = grd.CellTop
            cod = .TextMatrix(i, .ColIndex("Código"))
            MensajeStatus i & " de " & .Rows - .FixedRows, vbHourglass
            DoEvents
            
            
            'Recupera el objeto de Inventario
            Set pc = gobjMain.EmpresaActual.RecuperaPCProvCli(cod)
            
            Dim v As Variant, cad As String, j As Integer
    If InStr(1, pc.Email, ",") > 0 Then
        cad = ""
        v = Split(pc.Email, ",")
        For j = 0 To UBound(v)
            If Not Validar_Email(v(j)) Then
                grd.TextMatrix(i, .ColIndex("Verificado")) = "La dirección de Correo " & v(j) & " no es valida"
            Else
                If Not ValidarDominio_Email(v(j)) Then
                    grd.TextMatrix(i, .ColIndex("Verificado")) = "La dirección de Correo " & v(j) & " no es valida, el dominio esta"
                Else
                    cad = cad & v(j) & ","
                End If
            End If
        Next j
        pc.Email = Mid$(cad, 1, Len(cad) - 1)
    Else
            If Len(pc.Email) > 0 Then
                If Not Validar_Email(pc.Email) Then
                    grd.TextMatrix(i, .ColIndex("Verificado")) = "La dirección de Correo no es valida"
                    pc.Email = ""
                    pc.Grabar
                End If
                If Not ValidarDominio_Email(pc.Email) Then
                    grd.TextMatrix(i, .ColIndex("Verificado")) = "La dirección de Correo no es valida, el dominio esta"
                    pc.Email = ""
                    pc.Grabar
                End If
            End If
    
    End If
                        
                        
            
            
                        
            
            
            
            
        Next i
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

Public Function Validar_Email(ByVal Email As String) As Boolean
    
    Dim i As Integer, iLen As Integer, caracter As String
    Dim pos As Integer, bp As Boolean, ipos As Integer, iPos2 As Integer

    On Local Error GoTo Err_Sub

    Email = Trim$(Email)

    If Email = vbNullString Then
        Exit Function
    End If

    Email = LCase$(Email)
    iLen = Len(Email)

    
    For i = 1 To iLen
        caracter = Mid(Email, i, 1)

        If (Not (caracter Like "[a-z]")) And (Not (caracter Like "[0-9]")) Then
            
            If InStr(1, "_-" & "." & "@", caracter) > 0 Then
                If bp = True Then
                   Exit Function
                Else
                    bp = True
                   
                    If i = 1 Or i = iLen Then
                        Exit Function
                    End If
                    
                    If caracter = "@" Then
                        If ipos = 0 Then
                            ipos = i
                        Else
                            
                            Exit Function
                        End If
                    End If
                    If caracter = "." Then
                        iPos2 = i
                    End If
                    
                End If
            Else
                
                Exit Function
            End If
        Else
            bp = False
        End If
    Next i
    If ipos = 0 Or iPos2 = 0 Then
        Exit Function
    End If
    
    If iPos2 < ipos Then
        Exit Function
    End If

    
    Validar_Email = True

    Exit Function
Err_Sub:
    On Local Error Resume Next
    
    Validar_Email = False
End Function


Public Function ValidarDominio_Email(ByVal Email As String) As Boolean

Dim strTmp As String
Dim n As Long
Dim sEXT As String
Dim MensajeError As String

'MensajeError = ""
ValidarDominio_Email = True

sEXT = Email

Do While InStr(1, sEXT, ".") <> 0
   sEXT = Right(sEXT, Len(sEXT) - InStr(1, sEXT, "."))
Loop

If Email = "" Then
   ValidarDominio_Email = False
   'MensajeError = 'MensajeError & "No se indicó ninguna dirección de " & _
                  "mail para verificar!" & vbNewLine
ElseIf InStr(1, Email, "@") = 0 Then
   ValidarDominio_Email = False
   'MensajeError = 'MensajeError & "La dirección de email con contiene el signo @" & vbNewLine
ElseIf InStr(1, Email, "@") = 1 Then
   ValidarDominio_Email = False
   'MensajeError = 'MensajeError & "El @ No puede estar al principio" & vbNewLine
ElseIf InStr(1, Email, "@") = Len(Email) Then
   ValidarDominio_Email = False
   'MensajeError = 'MensajeError & "El @ no puede estar al final de la dirección" & vbNewLine
ElseIf EXTisOK(sEXT) = False Then
   ValidarDominio_Email = False
   'MensajeError = 'MensajeError & "La dirección no tiene un dominio válido, "
   'MensajeError = 'MensajeError & "por ejemplo : "
   'MensajeError = 'MensajeError & ".com, .net, .gov, .org, .edu, .biz, .gob etc.. " & vbNewLine
ElseIf Len(Email) < 6 Then
   ValidarDominio_Email = False
   'MensajeError = 'MensajeError & "La dirección no puede ser menor a 6 caracteres." & vbNewLine
End If
strTmp = Email
Do While InStr(1, strTmp, "@") <> 0
   n = 1
   strTmp = Right(strTmp, Len(strTmp) - InStr(1, strTmp, "@"))
Loop
If n > 1 Then
   ValidarDominio_Email = False
   'MensajeError = 'MensajeError & "Solo puede haber un @ en la dirección de email" & vbNewLine
End If

    Dim pos As Integer

    pos = InStr(1, Email, "@")

    If Mid(Email, pos + 1, 1) = "." Then
        ValidarDominio_Email = False
        'MensajeError = 'MensajeError & "El punto no puede estar seguido del @" & vbNewLine
    End If

    If MensajeError <> "" Then
        MsgBox MensajeError, vbCritical
    End If

End Function


Public Function EXTisOK(ByVal sEXT As String) As Boolean
Dim ext As String, X As Long
EXTisOK = False
If Left(sEXT, 1) <> "." Then sEXT = "." & sEXT
    sEXT = UCase(sEXT) 'just to avoid errors
    ext = ext & ".COM.EDU.GOV.NET.BIZ.ORG.TV"
    ext = ext & ".AF.AL.DZ.As.AD.AO.AI.AQ.AG.AP.AR.AM.AW.AU.AT.AZ.BS.BH.BD.BB.BY"
    ext = ext & ".BE.BZ.BJ.BM.BT.BO.BA.BW.BV.BR.IO.BN.BG.BF.MM.BI.KH.CM.CA.CV.KY"
    ext = ext & ".CF.TD.CL.CN.CX.CC.CO.KM.CG.CD.CK.CR.CI.HR.CU.CY.CZ.DK.DJ.DM.DO"
    ext = ext & ".TP.EC.EG.SV.GQ.ER.EE.ET.FK.FO.FJ.FI.CS.SU.FR.FX.GF.PF.TF.GA.GM.GE.DE"
    ext = ext & ".GH.GI.GB.GR.GL.GD.GP.GU.GT.GN.GW.GY.HT.HM.HN.HK.HU.IS.IN.ID.IR.IQ"
    ext = ext & ".IE.IL.IT.JM.JP.JO.KZ.KE.KI.KW.KG.LA.LV.LB.LS.LR.LY.LI.LT.LU.MO.MK.MG"
    ext = ext & ".MW.MY.MV.ML.MT.MH.MQ.MR.MU.YT.MX.FM.MD.MC.MN.MS.MA.MZ.NA"
    ext = ext & ".NR.NP.NL.AN.NT.NC.NZ.NI.NE.NG.NU.NF.KP.MP.NO.OM.PK.PW.PA.PG.PY"
    ext = ext & ".PE.PH.PN.PL.PT.PR.QA.RE.RO.RU.RW.GS.SH.KN.LC.PM.ST.VC.SM.SA.SN.SC"
    ext = ext & ".SL.SG.SK.SI.SB.SO.ZA.KR.ES.LK.SD.SR.SJ.SZ.SE.CH.SY.TJ.TW.TZ.TH.TG.TK"
    ext = ext & ".TO.TT.TN.TR.TM.TC.TV.UG.UA.AE.UK.US.UY.UM.UZ.VU.VA.VE.VN.VG.VI"
    ext = ext & ".WF.WS.EH.YE.YU.ZR.ZM.ZW.GOB"
    ext = UCase(ext) 'just to avoid errors
    If InStr(1, ext, sEXT, 0) <> 0 Then
        EXTisOK = True
    End If
End Function


Private Sub GrabarLectura()
    Dim i As Long, cod As String, emp As PCProvCli, sql As String, rs As Recordset
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
        For i = .FixedRows To .Rows - 1
            'Si es que se canceló el proceso
            If mCancelado Then GoTo salida
        
            prg1.value = i
            'cod = .TextMatrix(i, .ColIndex("Código"))
            MensajeStatus i & " de " & .Rows - .FixedRows, vbHourglass
            If grd.ValueMatrix(i, 5) > grd.ValueMatrix(i, 4) Then
                DoEvents
    '            If CrearTransacciones Then
                    If GenerarConsumo(i) Then
                        sql = " UPDATE GnCentroCosto "
                        sql = sql & " SET Valor1= " & .ValueMatrix(i, .ColIndex("Lectura Nueva"))
                        sql = sql & " where codcentro='" & .TextMatrix(i, .ColIndex("No. Medidor")) & "'"
                        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
                        grd.TextMatrix(i, .ColIndex("Resultado")) = "OK"
                    End If
            ElseIf grd.ValueMatrix(i, 5) = grd.ValueMatrix(i, 4) Then
                grd.TextMatrix(i, .ColIndex("Resultado")) = "No existe consumo"
            Else
                grd.TextMatrix(i, .ColIndex("Resultado")) = "Error la lectura Actual es Menor a la Anterior"
            End If
 '           End If
            'Recupera el objeto de Inventario
        Next i
    End With
    
salida:
    MensajeStatus
    Set emp = Nothing
    Habilitar True
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    GoTo salida
    Exit Sub
End Sub


Private Function CrearTransacciones() As Boolean
    On Error GoTo mensaje
    CrearTransacciones = True
    'Transaccion para conteo fisico

            Set mobjGNComp = gobjMain.EmpresaActual.CreaGNComprobante("CSM")

            Exit Function
    
mensaje:
    DispErr
    CrearTransacciones = False
End Function



Private Function GenerarConsumo(fila As Long) As Boolean
    Dim s As String, tid As Long, i As Long, X As Single
    Dim gnc As GNComprobante, cambiado As Boolean
    
    On Error GoTo ErrTrap
    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False

    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    ProcesarDatos (fila)
    Screen.MousePointer = 0
    GenerarConsumo = Not mCancelado
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    DispErr
salida:
    mProcesando = False
    frmMain.mnuFile.Enabled = True

    prg1.value = prg1.min
    Exit Function
End Function

Private Sub ProcesarDatos(fila As Integer)
    Dim ix As Long, ivk As IVKardex, dif As Currency
    Dim i As Long, signo As Integer, cant As String
    Dim iv As IVinventario, c As Currency
    Dim sql As String, rs As Recordset
    Dim item As IVinventario, CONSUMO As Currency, coditem As String
        If CrearTransacciones Then
                With mobjGNComp
                    CONSUMO = grd.ValueMatrix(fila, 5) - grd.ValueMatrix(fila, 4)
                    .CodClienteRef = grd.TextMatrix(fila, 2)
                    .nombre = grd.TextMatrix(fila, 3)
                    .CodCentro = grd.TextMatrix(fila, 1)
                    .FechaTrans = Date
                    .Descripcion = "Consumo de Agua del Mes " & MonthName(DatePart("m", Date)) & "/" & DatePart("yyyy", Date)
                    .HoraTrans = Time
                    Select Case CONSUMO
                    Case Is <= 5: coditem = "DE 00-05"
                    Case Is <= 10: coditem = "DE 00-10"
                    Case Is <= 15: coditem = "DE 00-15"
                    Case Else
                        coditem = "MAYOR 15"
                    End Select
                    Set item = .Empresa.RecuperaIVInventario(coditem)
                    ix = .AddIVKardex
                    Set ivk = .IVKardex(ix)
                    .IVKardex(ix).cantidad = CONSUMO * -1
                    .IVKardex(ix).CodBodega = ivk.CodBodega
                    .IVKardex(ix).CodInventario = item.CodInventario
                    If .FechaTrans >= gobjMain.EmpresaActual.GNOpcion.FechaIVA Then
                        .IVKardex(ix).IVA = IIf(item.bandIVA = True, gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA, 0)
                    Else
                        .IVKardex(ix).IVA = IIf(item.bandIVA = True, gobjMain.EmpresaActual.GNOpcion.PorcentajeIVAAnt, 0)
                    End If
                    .IVKardex(ix).Nota = "Esta lectura: " & grd.ValueMatrix(fila, 5) & ", Lectura Anterior:" & grd.ValueMatrix(fila, 4)
                    'Si el costo calculado está en otra moneda, convierte en moneda de trans.
                    
                    .IVKardex(ix).CostoTotal = "0"
                    .IVKardex(ix).PrecioTotal = item.Precio(1) * CONSUMO * -1
                    .Grabar False, False
                End With
           
           End If
            
        
    
End Sub

Private Sub BuscarDINARDAPParr()
    Static CodTrans As String, desde As Long, hasta As Long, fechadesde As Date, fechahasta As Date
    Dim sql As String, cond As String, rs As Recordset, comodin As String, ruc As String
    On Error GoTo ErrTrap
    'If Me.tag <> "CUENTA" Then Exit Sub
    #If DAOLIB Then
        comodin = "*"
    #Else
        comodin = "%"
    #End If
'    comodin = "%"
    'Abre la pantalla de búsqueda
    If Not frmIVBusqueda.InicioTransNew( _
                CodTrans, _
                desde, hasta, ruc, fechadesde, fechahasta) Then
        'Si fue cancelada la busqueda, sale no mas
        grd.SetFocus
        Exit Sub
    End If
    'Cambia la forma de cursor
    MensajeStatus MSG_PREPARA, vbHourglass
    
    'Compone la cadena de SQL
    sql = "SELECT  distinct codprovcli, ruc, pc.nombre, "
    sql = sql & " pcprov.CodProvincia, "
    sql = sql & " pcprov.descripcion, "
    sql = sql & " pccan.codcanton,"
    sql = sql & " pccan.Descripcion , "
    sql = sql & " pcparr.codparroquia, "
    sql = sql & " pcparr.descripcion, "
    sql = sql & " Tiposujeto, sexo, estadocivil, Origeningresos "
    sql = sql & " FROM gncomprobante gnc inner join pcprovcli pc "
    sql = sql & " left join pcprovincia pcprov on pc.idprovincia = pcprov.idprovincia "
    sql = sql & " left join pccanton pccan on pc.idcanton = pccan.idcanton "
    sql = sql & " left join pcparroquia pcparr "
    sql = sql & " inner join pccanton pcc1 on pcparr.idcanton = pcc1.idcanton"
    sql = sql & " on pc.idParroquia = pcparr.idparroquia "
    
    sql = sql & " on gnc.IdClienteref=pc.idprovcli"
    sql = sql & " WHERE pccan.codcanton <> pcc1.codcanton "
    'CodInventario
'    If Len(CodTrans) > 0 Then
'        'If Len(Cond) > 0 Then Cond = Cond & "AND "
'        Cond = Cond & "and (gnc.Codtrans = '" & CodTrans & "') "
'    End If
'
'
'    If Len(ruc) > 0 Then
'        'If Len(ruc) > 0 Then Cond = Cond & "AND "
'        Cond = Cond & "and (ruc like '" & ruc & "%') "
'    End If
'
'    'CodAlterno
'    If desde <> 0 And hasta <> 0 Then
'        'If Len(Cond) > 0 Then Cond = Cond & "AND "
'        Cond = Cond & " and numtrans between " & desde & " and " & hasta
'    End If'
    
'    If Len(Cond) > 0 Then sql = sql & Cond
'    sql = sql & " and gnc.fechatrans between '" & FechaDesde & "' And  '" & FechaHasta & "'"
    sql = sql & " ORDER BY pc.nombre "
    
    
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    With grd
        .Redraw = flexRDNone
        .Rows = .FixedRows
        If Not rs.EOF Then .LoadArray MiGetRows(rs)
        ConfigCols
        .Redraw = flexRDBuffered
        .SetFocus
    End With
    CargarProvincias
    CargarTipoSujeto
    MensajeStatus
    Exit Sub
ErrTrap:
    grd.Redraw = flexRDBuffered
    MensajeStatus
    DispErr
    grd.SetFocus
    Exit Sub
End Sub

Private Sub BuscarFechaInicial() 'AUC plan mantenimiento
    Dim sql As String, cond As String
    Dim objcond As Condicion
    Dim rs As Recordset
    Dim i As Long
    Dim ivp As IVPlan
    Dim iv As IVinventario
    Dim gnv As GnVehiculo
    On Error GoTo ErrTrap
    'If Me.tag <> "CUENTA" Then Exit Sub
    Set objcond = gobjMain.objCondicion
'    comodin = "%"
    'Abre la pantalla de búsqueda
    'CodTrans = "UltimoCosto"
    If Not frmB_PlanMant.Inicio(objcond _
                ) Then
        'Si fue cancelada la busqueda, sale no mas
       grd.SetFocus
        Exit Sub
    End If
    'Cambia la forma de cursor
    MensajeStatus MSG_PREPARA, vbHourglass
    If Len(objcond.CodBanco1) > 0 Then
        Set ivp = gobjMain.EmpresaActual.RecuperaIVPLAN(objcond.CodBanco1)
        Set iv = gobjMain.EmpresaActual.RecuperaIVInventarioQuick(ivp.idinventario)
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = " Where CHARINDEX('" & ivp.IdPlan & "',PLANMANT) > 0"
    End If
    'Compone la cadena de SQL
    sql = "SELECT gnv.codvehiculo,'" & objcond.CodBanco1 & "','" & iv.CodInventario & "', '' as  fechaProx FROM GNVEHICULO GNV"

    If Len(cond) > 0 Then sql = sql & cond

    'sql = sql & "And gnvp.idplan <> 2 "
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
'    With grd
        grd.Redraw = flexRDNone
        grd.Rows = grd.FixedRows
        If Not rs.EOF Then grd.LoadArray MiGetRows(rs)
        ConfigCols
        grd.Redraw = flexRDBuffered
        grd.SetFocus
'    End With
    'cargar fechas existentes
    For i = 1 To grd.Rows - 1
        Set ivp = gobjMain.EmpresaActual.RecuperaIVPLAN(objcond.CodBanco1)
        Set gnv = gobjMain.EmpresaActual.RecuperaGNVehiculo(grd.TextMatrix(i, 1))
        sql = "Select fechaProx from gnvehiculoPlan gnvp inner join ivplan ivp on ivp.idplan = gnvp.idplan"
        sql = sql & " Inner Join GnVehiculo gnv on gnv.idVehiculo = gnvp.idvehiculo "
        sql = sql & " where gnv.codvehiculo='" & gnv.CodVehiculo & "'"
        sql = sql & " And ivp.codPlan ='" & ivp.CodPlan & "'"
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
        If rs.RecordCount > 0 Then
            grd.TextMatrix(i, 4) = rs!fechaprox
        End If
        Set ivp = Nothing
        Set iv = Nothing
    Next
    'borramos los que ya existen
    For i = grd.Rows - 1 To 1 Step -1
        If Not grd.IsSubtotal(i) Then
            If Len(grd.TextMatrix(i, 4)) > 0 Then
                grd.RemoveItem i
            End If
        End If
    Next
    MensajeStatus
    Set ivp = Nothing
    Set iv = Nothing
    Exit Sub
ErrTrap:
    grd.Redraw = flexRDBuffered
    MensajeStatus
    DispErr
    grd.SetFocus
    Exit Sub
End Sub
Private Sub AsignarFechaInicial()
    Dim s As String, v As Date
    Dim i As Long
    s = InputBox("Ingrese una Fecha", "Asignar un valor", Date)
    If IsDate(s) Then
        v = (s)
    Else
        MsgBox "Debe ingresar una fecha. (ejm. dd/mm/yyyy) ", vbInformation
       grd.SetFocus
        Exit Sub
    End If
    With grd
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("Fecha Inicial")) = v
        Next i
    End With
End Sub
Private Sub BuscarPCgar()
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
    
    
    'Abre la pantalla de búsqueda
    frmPCBusqueda.Caption = "Busqueda de Garantes"
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
        Case "PCGRUPOS_PROV", "PCGRUPOS_CLI", "PCGRUPOS_GAR"
            'Compone la cadena de SQL
            sql = "SELECT CodProvCli, Nombre" & _
            ", codgrupo1, codgrupo2, codgrupo3 ,codgrupo4 " & _
            "FROM vwPCProvCli "
        

        End Select
    ' si Busca Proveedor o cliente mediante bandera de prov
    If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(BandGarante = 1) "
    
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
    
    sql = sql & " ORDER BY Nombre "
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


Private Sub BuscarFormaCobro()
    Static CodTrans As String, fechadesde As Date, fechahasta As Date
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
    If Not frmIVBusqueda.InicioFormaCobro( _
                CodTrans, CodTransRel, _
                fechadesde, fechahasta) Then
        'Si fue cancelada la busqueda, sale no mas
        grd.SetFocus
        Exit Sub
    End If
    
    'Cambia la forma de cursor
    MensajeStatus MSG_PREPARA, vbHourglass
    
    'Compone la cadena de SQL
    sql = "SELECT pck.id, pck.idforma, gnc.fechatrans, gnc.codtrans, gnc.numtrans , codforma,  a.codformapago "
    sql = sql & " FROM gncomprobante gnc "
    sql = sql & " inner join gntrans gt on gnc.codtrans=gt.codtrans"
    sql = sql & " inner join pckardex pck "
    sql = sql & " left join Anexo_FormaPago a  "
    sql = sql & " on pck.idformasri=a.id "
    sql = sql & " inner join tsformacobropago tsf "
    sql = sql & " on pck.idforma = tsf.idforma"
    sql = sql & " on gnc.transid=pck.transid"
'    sql = sql & " inner join gntrans gnt on gnc.codtrans=gnt.codtrans"
    
    'CodInventario
    If Len(CodTrans) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(gt.anexocodtipocomp = '" & CodTrans & "') "
    End If
    
    If Len(CodTransRel) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(tsf.CodForma = '" & CodTransRel & "') "
    End If
    
    'CodAlterno
    If fechadesde <> 0 And fechahasta <> 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "gnc.fechatrans between '" & fechadesde & "' and '" & fechahasta & "'"
    End If
    
    
    
    
    If Len(cond) > 0 Then sql = sql & " WHERE  " & cond & " and anexocodtipotrans =2"
    sql = sql & " ORDER BY gnc.Numtrans "
    
    
    
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


Private Sub CargarFormasSRI()
    Dim s As String
    With grd

            s = gobjMain.EmpresaActual.ListaAnexoFormaPagoParaFlexGrid(1)
            s = Right$(s, Len(s) - 1)

        .ColComboList(.ColIndex("Cod Forma SRI")) = s
    
        
    End With
End Sub

Private Sub AsignarFormaCobro()
    Dim FormaSRI As String
    Dim i As Long, s As String
    
    With grd
        'Obtiene cuentas de la fila actual
        FormaSRI = .TextMatrix(.Row, .ColIndex("Cod Forma SRI"))
        
        s = "Está seguro que desea asignar los siguientes códigos " & _
            "en todos los ítems que están visualizados?" & vbCr & vbCr & _
            "    Forma Cobro SRI:  " & FormaSRI
        If MsgBox(s, vbQuestion + vbYesNo) <> vbYes Then
            .SetFocus
            Exit Sub
        End If
        
        'Copia a todas las filas los mismos códigos de cuenta
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("Cod Forma SRI")) = FormaSRI
        Next i
    End With

End Sub


Private Sub BuscarTransRTC()
    Static CodTrans As String, desde As Long, hasta As Long
    Dim sql As String, cond As String, rs As Recordset, comodin As String
    Dim ruc As String, fechadesde As Date, fechahasta As Date
    On Error GoTo ErrTrap
    'If Me.tag <> "CUENTA" Then Exit Sub
    
    #If DAOLIB Then
        comodin = "*"
    #Else
        comodin = "%"
    #End If
'    comodin = "%"
    'Abre la pantalla de búsqueda
    If Not frmIVBusqueda.InicioTransNew( _
                CodTrans, desde, hasta, ruc, fechadesde, fechahasta) Then
        'Si fue cancelada la busqueda, sale no mas
        grd.SetFocus
        Exit Sub
    End If
    
    'Cambia la forma de cursor
    MensajeStatus MSG_PREPARA, vbHourglass
    
    'Compone la cadena de SQL
    sql = "SELECT transid, fechatrans, nombre, gnc.codtrans, numtrans , estado1 "
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
    
    
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "fechatrans  between '" & fechadesde & "' and '" & fechahasta & "'"
    
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & " estado1 <> 1 "
    
    
    
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



Private Sub AsignarFCVendedor()
    Dim ValorGrupo As String
    Dim i As Long, s As String, j As Integer, grupo As String
    With grd
            ValorGrupo = .TextMatrix(.Row, PCGRUPO2)
            'Confirma las GRUPOS
            s = "Está seguro que desea asignar el siguient código de vendedor " & ValorGrupo & _
                ", en todos los Clientes que están visualizados?"
            If MsgBox(s, vbQuestion + vbYesNo) = vbYes Then
                'Copia a todas las filas los mismos códigos de cuenta
                For i = .FixedRows To .Rows - 1
                            .TextMatrix(i, PCGRUPO2) = ValorGrupo
                Next i
            End If

    End With
End Sub


Private Sub GrabarPCAgencia()
    Dim i As Long, pc As PCProvCli, cod As String, X As Single, bandOk As Boolean, pca As PCAgencia
    On Error GoTo ErrTrap
    
    'Confirmación
    If MsgBox("Está seguro que desea Verificar y grabar?", vbQuestion + vbYesNo) <> vbYes Then
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
        bandOk = False
        For i = .FixedRows To .Rows - 1
            bandOk = False
            'Si es que se canceló el proceso
            If mCancelado Then GoTo salida
        
            prg1.value = i
            grd.Row = i
            X = grd.CellTop
            cod = .TextMatrix(i, .ColIndex("Código"))
            MensajeStatus i & " de " & .Rows - .FixedRows, vbHourglass
            DoEvents
            
            If grd.TextMatrix(i, .ColIndex("Resultado")) <> "OK" Then
            
            'Recupera el objeto de Inventario
            Set pc = gobjMain.EmpresaActual.RecuperaPCProvCliQuick(cod)
            
            
            
            
                If Not gobjMain.EmpresaActual.Empresa2.RecuperarAgenciaxCodprovcli(pc.IdProvCli) Then
            
                    Set pca = gobjMain.EmpresaActual.CreaAgencia
                    pca.IdProvCli = pc.IdProvCli
                    pca.Direccion = pc.Direccion1
                    pca.Telefono = Val(Mid$(pc.Telefono1, 1, 10))
                    pca.Contacto = pc.banco
                    pca.IdProvincia = pc.IdProvincia
                    pca.IdCiudad = pc.Idcanton
                    pca.IdGrupo1 = pc.IdGrupo1
                    pca.IdGrupo2 = pc.IdGrupo2
                    pca.IdGrupo3 = pc.IdGrupo3
                    pca.IdGrupo4 = pc.IdGrupo4
                    pca.IdVendedor = pc.IdVendedor
                    pca.Orden = 1
                    pca.CodAgencia = "001"
                    pca.Descripcion = pc.NombreAlterno
                    pca.Grabar
                    grd.TextMatrix(i, .ColIndex("Resultado")) = "OK"
                End If
            End If
        Next i
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

