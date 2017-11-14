VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmB_FormSRI 
   Caption         =   "Condiciones de Busqueda"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8490
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   8490
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkFactor 
      Caption         =   "Utilizar Factor Proporcionalidad"
      Height          =   255
      Left            =   2940
      TabIndex        =   1
      Top             =   330
      Width           =   2595
   End
   Begin VB.Frame fraRecargos 
      Caption         =   "Recargos y Descuentos antes del IVA"
      Height          =   1395
      Left            =   120
      TabIndex        =   39
      Top             =   6240
      Width           =   8295
      Begin VB.ListBox lstRecar1 
         Height          =   1035
         Left            =   120
         TabIndex        =   9
         Top             =   270
         Width           =   3675
      End
      Begin VB.ListBox lstRecar2 
         Height          =   1035
         Left            =   4500
         TabIndex        =   12
         Top             =   285
         Width           =   3675
      End
      Begin VB.CommandButton cmdAdd2 
         Caption         =   "&>>"
         Height          =   375
         Left            =   3840
         TabIndex        =   10
         Top             =   420
         Width           =   615
      End
      Begin VB.CommandButton cmdResta2 
         Caption         =   "&<<"
         Height          =   375
         Left            =   3840
         TabIndex        =   11
         Top             =   825
         Width           =   615
      End
   End
   Begin TabDlg.SSTab sst1 
      Height          =   5475
      Left            =   60
      TabIndex        =   38
      Top             =   720
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   9657
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Ventas"
      TabPicture(0)   =   "frmB_FomSRI.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblDragDrop"
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(2)=   "lstTransVentas"
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(5)=   "Frame4"
      Tab(0).Control(6)=   "Frame6"
      Tab(0).Control(7)=   "Frame7"
      Tab(0).Control(8)=   "Frame8"
      Tab(0).Control(9)=   "Frame12"
      Tab(0).Control(10)=   "Frame13"
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Compras"
      TabPicture(1)   =   "frmB_FomSRI.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraTrans1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lstTransCompras"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame9"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame10"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame11"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Retenciones"
      TabPicture(2)   =   "frmB_FomSRI.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraRetRecibidas"
      Tab(2).Control(1)=   "fraRetRealizada"
      Tab(2).Control(2)=   "Fra103"
      Tab(2).ControlCount=   3
      Begin VB.Frame Frame13 
         Caption         =   "Trans Notas Crédito Exp Ser"
         Height          =   960
         Left            =   -69540
         TabIndex        =   63
         Top             =   4320
         Width           =   2715
         Begin VB.ListBox lstNC_ExpS 
            Height          =   645
            IntegralHeight  =   0   'False
            Left            =   120
            TabIndex        =   64
            Top             =   225
            Width           =   2505
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Trans Notas Crédito Exp Bie"
         Height          =   960
         Left            =   -72300
         TabIndex        =   61
         Top             =   4320
         Width           =   2715
         Begin VB.ListBox lstNC_ExpB 
            Height          =   645
            IntegralHeight  =   0   'False
            Left            =   120
            TabIndex        =   62
            Top             =   225
            Width           =   2505
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Tran. de Imp. Act Fijos (505-506)"
         Height          =   1065
         Left            =   5520
         TabIndex        =   59
         Top             =   4260
         Width           =   2715
         Begin VB.ListBox lstRise 
            Height          =   765
            IntegralHeight  =   0   'False
            Left            =   120
            TabIndex        =   60
            Top             =   225
            Width           =   2500
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Tran. de Imp. Bienes (504)"
         Height          =   1080
         Left            =   2700
         TabIndex        =   58
         Top             =   4260
         Width           =   2715
         Begin VB.ListBox lstTransBie 
            Height          =   765
            IntegralHeight  =   0   'False
            Left            =   120
            TabIndex        =   17
            Top             =   225
            Width           =   2500
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Tran. de NC Compras"
         Height          =   2340
         Left            =   5520
         TabIndex        =   57
         Top             =   480
         Width           =   2715
         Begin VB.ListBox lstTransNCCompra 
            Height          =   1965
            IntegralHeight  =   0   'False
            Left            =   120
            TabIndex        =   16
            Top             =   225
            Width           =   2500
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Trans Notas Crédito"
         Height          =   1980
         Left            =   -69540
         TabIndex        =   56
         Top             =   420
         Width           =   2715
         Begin VB.ListBox lstNC_ventas 
            Height          =   1665
            IntegralHeight  =   0   'False
            Left            =   60
            TabIndex        =   4
            Top             =   240
            Width           =   2565
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Tran de Rep. Gastos"
         Height          =   960
         Left            =   -69540
         TabIndex        =   55
         Top             =   2400
         Width           =   2715
         Begin VB.ListBox LstRepGast 
            Height          =   645
            IntegralHeight  =   0   'False
            Left            =   120
            TabIndex        =   8
            Top             =   225
            Width           =   2500
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Tran de Exp. Servicios"
         Height          =   960
         Left            =   -69540
         TabIndex        =   54
         Top             =   3360
         Width           =   2715
         Begin VB.ListBox LstExpSer 
            Height          =   645
            IntegralHeight  =   0   'False
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   2505
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Tran de Exp. Bienes"
         Height          =   960
         Left            =   -72300
         TabIndex        =   5
         Top             =   3360
         Width           =   2715
         Begin VB.ListBox lstExpBien 
            Height          =   645
            IntegralHeight  =   0   'False
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   2505
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Trans de Ventas Activos"
         Height          =   960
         Left            =   -72300
         TabIndex        =   49
         Top             =   2400
         Width           =   2715
         Begin VB.ListBox LstVentas0 
            Height          =   645
            IntegralHeight  =   0   'False
            Left            =   120
            TabIndex        =   6
            Top             =   225
            Width           =   2500
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Trans de Ventas Netas"
         Height          =   1980
         Left            =   -72300
         TabIndex        =   48
         Top             =   420
         Width           =   2715
         Begin VB.ListBox LstVentas12 
            Height          =   1665
            IntegralHeight  =   0   'False
            Left            =   120
            TabIndex        =   3
            Top             =   225
            Width           =   2500
         End
      End
      Begin VB.ListBox lstTransVentas 
         BackColor       =   &H80000018&
         Height          =   4665
         IntegralHeight  =   0   'False
         Left            =   -74880
         TabIndex        =   2
         Top             =   600
         Width           =   2500
      End
      Begin VB.Frame fraRetRecibidas 
         Caption         =   "Transacciones de Retenciones Recibidas"
         Height          =   2340
         Left            =   -74880
         TabIndex        =   43
         Top             =   420
         Width           =   8055
         Begin VB.ListBox lstRetRecibidas 
            Height          =   1965
            IntegralHeight  =   0   'False
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   19
            Top             =   240
            Width           =   7755
         End
      End
      Begin VB.Frame fraRetRealizada 
         Caption         =   "Transacciones de Retenciones Realizadas"
         Height          =   2460
         Left            =   -74880
         TabIndex        =   41
         Top             =   2880
         Width           =   8055
         Begin VB.ListBox LstRetencion 
            Height          =   1005
            IntegralHeight  =   0   'False
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   21
            Top             =   1320
            Width           =   7785
         End
         Begin VB.ListBox lstRetRealizadas 
            Height          =   746
            IntegralHeight  =   0   'False
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   20
            Top             =   240
            Width           =   7815
         End
         Begin VB.Label lblTipoRetencion 
            Caption         =   "Tipos de Retencion IVA"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   1080
            Width           =   2415
         End
      End
      Begin VB.ListBox lstTransCompras 
         BackColor       =   &H80000018&
         Height          =   4725
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   2500
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tran. de Compras Activos"
         Height          =   1320
         Left            =   2700
         TabIndex        =   44
         Top             =   2880
         Width           =   2715
         Begin VB.ListBox lstTransAct 
            Height          =   945
            IntegralHeight  =   0   'False
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   2500
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tran. de Compras Reembolso"
         Height          =   1320
         Left            =   5520
         TabIndex        =   45
         Top             =   2880
         Width           =   2715
         Begin VB.ListBox lstTransRem 
            Height          =   1005
            IntegralHeight  =   0   'False
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   2500
         End
      End
      Begin VB.Frame fraTrans1 
         Caption         =   "Tran. de Compras Netas"
         Height          =   2325
         Left            =   2700
         TabIndex        =   46
         Top             =   480
         Width           =   2715
         Begin VB.ListBox lstTransCompraNeta 
            Height          =   1965
            IntegralHeight  =   0   'False
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   2500
         End
      End
      Begin VB.Frame Fra103 
         Caption         =   "Codigos Retenciones"
         Height          =   3615
         Left            =   -74880
         TabIndex        =   51
         Top             =   420
         Visible         =   0   'False
         Width           =   5355
         Begin VB.ListBox LstRetencion103 
            Height          =   3285
            IntegralHeight  =   0   'False
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   52
            Top             =   240
            Width           =   5085
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Transacciones de Venta"
         Height          =   255
         Left            =   -74880
         TabIndex        =   50
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblDragDrop 
         Caption         =   "Label4"
         Height          =   255
         Left            =   -74280
         TabIndex        =   40
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Transacciones de Compra"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   360
         Width           =   2715
      End
   End
   Begin VB.Frame fraCobro 
      Caption         =   "Códigos de Retenciones de IVA"
      Height          =   1590
      Left            =   10020
      TabIndex        =   31
      Top             =   3420
      Width           =   5175
      Begin VB.ListBox lstBienes 
         Height          =   1035
         Left            =   120
         TabIndex        =   35
         Top             =   450
         Width           =   2055
      End
      Begin VB.ListBox lstServicios 
         Height          =   1035
         Left            =   3000
         TabIndex        =   34
         Top             =   450
         Width           =   2055
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&>>"
         Height          =   375
         Left            =   2280
         TabIndex        =   33
         Top             =   570
         Width           =   615
      End
      Begin VB.CommandButton cmdResta 
         Caption         =   "&<<"
         Height          =   375
         Left            =   2280
         TabIndex        =   32
         Top             =   1050
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "SERVICIOS"
         Height          =   255
         Left            =   3000
         TabIndex        =   37
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "BIENES"
         Height          =   255
         Left            =   105
         TabIndex        =   36
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   8490
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   7710
      Width           =   8490
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   4380
         TabIndex        =   23
         Top             =   60
         Width           =   1452
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar -F5"
         Height          =   372
         Left            =   2655
         TabIndex        =   22
         Top             =   75
         Width           =   1452
      End
   End
   Begin VB.Frame fraFecha 
      Caption         =   "Rango de Fecha"
      Height          =   705
      Left            =   60
      TabIndex        =   27
      Top             =   0
      Width           =   2715
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   360
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         Format          =   106692611
         UpDown          =   -1  'True
         CurrentDate     =   36891
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   360
         Left            =   840
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   1692
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         Format          =   106692609
         CurrentDate     =   36526
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "H&asta  "
         Height          =   195
         Left            =   2760
         TabIndex        =   29
         Top             =   270
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D&esde  "
         Height          =   192
         Left            =   240
         TabIndex        =   28
         Top             =   276
         Width           =   564
      End
   End
   Begin VB.Frame fraTrans2 
      Caption         =   "Transacciones de Retención"
      Height          =   1275
      Left            =   9840
      TabIndex        =   24
      Top             =   5400
      Width           =   5175
      Begin VB.ListBox lstTrans2 
         Height          =   975
         IntegralHeight  =   0   'False
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   26
         Top             =   210
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmB_FormSRI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private BandAceptado As Boolean
Private mobjSiiMain As SiiMain

Public Function Inicio104(ByRef objcond As Condicion, ByRef Recargo As String, _
                                            ByRef CP_Ser As String, ByRef CP_Act As String, _
                                            ByRef CP_Dev As String, ByRef VT_Dev As String, _
                                            ByRef VT_12 As String, ByRef VT_Act As String, _
                                            ByRef VT_ExpB As String, ByRef VT_expS As String, _
                                            ByRef VT_RepGas As String, _
                                            ByRef ret_real As String, ByRef ret_recib As String, _
                                            ByRef Reten As String, ByRef NC_Ventas As String, ByRef NC_Compras As String, _
                                            ByRef CP_Bie As String) As Boolean
    Dim KeyTrans As String, KeyRecargo As String, KeyTransRet As String, KeyRet As String, KeyT As String
    Dim KeyTNC As String, KeyTNCC As String
    Dim trans As String, s As String
    On Error GoTo ErrTrap
    Screen.MousePointer = vbHourglass
    BandAceptado = False
    'visualizar la condicion anterior
    With objcond
        dtpHasta.Format = dtpCustom
        dtpHasta.CustomFormat = "MMMM yyyy"
        If .fecha1 <> 0 Then
            dtpHasta.value = .fecha1
        Else
            dtpHasta.value = Now
        End If
        chkFactor.value = IIf(.IncluirCero, vbChecked, vbUnchecked)
        
        'ventas
        CargaTransxTipoTrans lstTransVentas, 2, 4
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_Ventas")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_Ventas")
            RecuperaSeleccion KeyT, LstVentas12, lstTransVentas, trans
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasAct")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasAct")
            RecuperaSeleccion KeyT, LstVentas0, lstTransVentas, trans
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasExpB")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasExpB")
            RecuperaSeleccion KeyT, lstExpBien, lstTransVentas, trans
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasExpS")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasExpS")
            RecuperaSeleccion KeyT, LstExpSer, lstTransVentas, trans
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasRepGas")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasRepGas")
            RecuperaSeleccion KeyT, LstRepGast, lstTransVentas, trans
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_NCVentas")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_NCVentas")
            RecuperaSeleccion KeyTNC, lstNC_ventas, lstTransVentas, trans
        End If
              
              
              
        'compras
        CargaTransxTipoTrans lstTransCompras, 1
       
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_Compras")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_Compras")
            RecuperaSeleccion KeyT, lstTransCompraNeta, lstTransCompras, trans
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasSer")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasSer")
            RecuperaSeleccion KeyT, lstTransRem, lstTransCompras, trans
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasAct")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasAct")
            RecuperaSeleccion KeyT, lstTransAct, lstTransCompras, trans
        End If
       
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_NCCompras")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_NCCompras")
            RecuperaSeleccion KeyTNCC, lstTransNCCompra, lstTransCompras, trans
        End If
       
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasBie")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasBie")
            RecuperaSeleccion KeyT, lstTransBie, lstTransCompras, trans
        End If
               
        'Retenciones
        CargaTransxTipoComprobante lstRetRecibidas, 7
        CargaTransxTipoComprobante lstRetRealizadas, 7
        CargaTransRetencion LstRetencion, True
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_RetenRecibidas")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_RetenRecibidas")
            RecuperaSelec KeyT, lstRetRecibidas, trans
        End If

        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_RetenRealizadas")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_RetenRealizadas")
            RecuperaSelec KeyT, lstRetRealizadas, trans
        End If

        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_Reten")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_Reten")
            RecuperaSelec KeyT, LstRetencion, trans
        End If
                          
        
        'carga recargos
        CargaRecargo
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_RecarDesc")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_RecarDesc")
            RecuperaSelecRec KeyRecargo, trans
        End If
        
        BandAceptado = False
        Screen.MousePointer = 0
        sst1.Tab = 0
       Me.Show vbModal
'        Si aplastó el botón 'Aceptar'
        If BandAceptado Then

'            'Devuelve las condiciones de búsqueda en el objeto objCondicion en SiiMain
            .fecha1 = dtpHasta.value
            .CodTrans = PreparaCadena(lstTransCompraNeta)
            .IncluirCero = (chkFactor.value = vbChecked)
            Recargo = PreparaCadRec(lstRecar2)
            
            'ventas
            VT_12 = PreparaCadena(LstVentas12)
            VT_Act = PreparaCadena(LstVentas0)
            VT_ExpB = PreparaCadena(lstExpBien)
            VT_expS = PreparaCadena(LstExpSer)
            VT_RepGas = PreparaCadena(LstRepGast)
            NC_Ventas = PreparaCadena(lstNC_ventas)
            NC_Compras = PreparaCadena(lstTransNCCompra)
            
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_Ventas", VT_12
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_VentasAct", VT_Act
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_VentasExpB", VT_ExpB
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_VentasExpS", VT_expS
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_VentasExpRepGas", VT_RepGas
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_NCVentas", NC_Ventas
            
            CP_Ser = PreparaCadena(lstTransRem)
            CP_Act = PreparaCadena(lstTransAct)
            CP_Bie = PreparaCadena(lstTransBie)
            
            
            'compras
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_Compras", .CodTrans
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_ComprasSer", CP_Ser
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_ComprasAct", CP_Act
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_NCCompras", NC_Compras
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_ComprasBie", CP_Bie
            
                   
            'retenciones
            ret_real = PreparaCadenaSeleccion(lstRetRealizadas)
            ret_recib = PreparaCadenaSeleccion(lstRetRecibidas)
            Reten = PreparaCadenaSeleccion(LstRetencion)
            
'            SaveSetting APPNAME, App.Title, "Reten_Recibidas", ret_recib
'            SaveSetting APPNAME, App.Title, "Reten_Realizadas", ret_real
'            SaveSetting APPNAME, App.Title, "Reten", Reten
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_RetenRecibidas", ret_recib
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_RetenRealizadas", ret_real
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_Reten", Reten
            'recargos descuentos
'            SaveSetting APPNAME, App.Title, "RecarDesc", Recargo
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_RecarDesc", Recargo
            'Graba en la base
            gobjMain.EmpresaActual.GNOpcion.Grabar
        End If
    End With
    Inicio104 = BandAceptado
    Unload Me
    'Set objSiiMain = Nothing
    Exit Function
ErrTrap:
    Screen.MousePointer = 0
    DispErr
    Exit Function
End Function

Private Sub Habilita()
    dtpDesde.Enabled = True
    dtpHasta.Enabled = True
    fraFecha.Enabled = True
    fraTrans1.Enabled = True
    fraTrans2.Enabled = True
    fraCobro.Enabled = True
    fraRecargos.Enabled = True
    Label2.Visible = True
    Label3.Visible = True
End Sub



'Private Sub cboGrupo_Click()
'    Dim Numg As Integer
'    On Error GoTo Errtrap
'    Numg = cboGrupo.ListIndex + 1
'    fcbGrupoDesde.SetData mobjSiiMain.EmpresaActual.ListaPCGrupo(Numg, True, False)
'    fcbGrupoHasta.SetData mobjSiiMain.EmpresaActual.ListaPCGrupo(Numg, False, False)
'    fcbGrupoDesde.KeyText = ""
'    fcbGrupoHasta.KeyText = ""
'    Exit Sub
'Errtrap:
'    DispErr
'    Exit Sub
'End Sub
Private Sub cmdAceptar_Click()
    Dim msg As String, ctl As Control
    On Error Resume Next
    If Len(msg) > 0 Then
        MsgBox msg, vbInformation
        ctl.SetFocus
        Exit Sub
    End If

    BandAceptado = True
    Me.Hide
End Sub

Private Sub cmdAdd2_Click()
    Dim i As Long, ix As Long
    On Error GoTo ErrTrap
    With lstRecar1
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                'ix = mobjGrupo.AgregarUsuario(.List(i))
                ix = .ItemData(i)
                lstRecar2.AddItem .List(i)
                lstRecar2.ItemData(lstRecar2.NewIndex) = ix
                .RemoveItem i
            End If
        Next i
    End With
    Exit Sub
ErrTrap:
    DispErr
End Sub

Private Sub cmdCancelar_Click()
    BandAceptado = False
    Me.Hide
End Sub


Private Sub cmdResta2_Click()
    Dim i As Long, ix As Long
    On Error GoTo ErrTrap
    With lstRecar2
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                ix = .ItemData(i)
                lstRecar1.AddItem .List(i)
                lstRecar1.ItemData(lstRecar1.NewIndex) = ix
                .RemoveItem i
            End If
        Next i
    End With
    Exit Sub
ErrTrap:
    DispErr
End Sub

'Private Sub fcbdesde1_Selected(ByVal Text As String, ByVal KeyText As String)
'    fcbHasta1.KeyText = fcbDesde1.KeyText
'End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF5
        cmdAceptar_Click
    Case vbKeyEscape
        cmdCancelar_Click
    Case Else
        MoverCampo Me, KeyCode, Shift, False
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    ImpideSonidoEnter Me, KeyAscii
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Me.Hide         'Se pone esto para evitar el posible BUG de Windows98
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjSiiMain = Nothing
End Sub


Private Sub cmdAdd_Click()
    Dim i As Long, ix As Long
    On Error GoTo ErrTrap
    With lstBienes
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                'ix = mobjGrupo.AgregarUsuario(.List(i))
                ix = .ItemData(i)
                lstServicios.AddItem .List(i)
                lstServicios.ItemData(lstServicios.NewIndex) = ix
                .RemoveItem i
            End If
        Next i
    End With
    Exit Sub
ErrTrap:
    DispErr
End Sub

Private Sub cmdResta_Click()
    Dim i As Long, ix As Long
    On Error GoTo ErrTrap
    With lstServicios
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                ix = .ItemData(i)
                lstBienes.AddItem .List(i)
                lstBienes.ItemData(lstBienes.NewIndex) = ix
                .RemoveItem i
            End If
        Next i
    End With
    Exit Sub
ErrTrap:
    DispErr
End Sub


Private Sub List1_Click()

End Sub

Private Sub lstBienes_DblClick()
    cmdAdd_Click
End Sub



Private Sub lstNC_ventas_DblClick()
    Regresa lstNC_ventas, lstTransVentas
End Sub

Private Sub lstRecar1_DblClick()
    cmdAdd2_Click
End Sub

Private Sub lstRecar2_DblClick()
    cmdResta2_Click
End Sub

Private Sub lstRise_DblClick()
Regresa lstRise, lstTransCompras
End Sub

Private Sub lstRise_DragDrop(Source As Control, x As Single, y As Single)
   On Error Resume Next
carga lstRise, lstTransCompras
End Sub

Private Sub lstServicios_DblClick()
    cmdResta_Click
End Sub

Private Function PreparaCadena(lst As ListBox) As String
    Dim Cadena As String, i As Integer, pos As Integer
    Cadena = ""
    For i = 0 To lst.ListCount - 1
        'If lst.Selected(i) Then
            pos = InStr(1, lst.List(i), " ")
            If Cadena = "" Then
                
                'cadena = Left(lst.List(i), lst.ItemData(i))
                Cadena = Trim(Mid$(lst.List(i), 1, pos))
            Else
                'cadena = cadena & "," & _
                              Left(lst.List(i), lst.ItemData(i))
                Cadena = Cadena & "," & Trim(Mid$(lst.List(i), 1, pos))
            End If
        'End If
    Next i
    PreparaCadena = Cadena
End Function



Private Sub CargaRecargo()
    Dim rs As Recordset
    Set rs = gobjMain.EmpresaActual.ListaIVRecargo(True)
    With rs
        If Not (.EOF) Then
            .MoveFirst
            Do Until .EOF
                lstRecar1.AddItem !codRecargo & "  " & !Descripcion
                lstRecar1.ItemData(lstRecar1.NewIndex) = Len(!codRecargo)
               .MoveNext
           Loop
           
            lstRecar1.AddItem "SUBT" & "  " & "Subtotal"
            lstRecar1.ItemData(lstRecar1.NewIndex) = Len("SUBT")
            
        End If
    End With
    rs.Close
End Sub

Private Function PreparaCadRec(lst As ListBox) As String
    Dim Cadena As String, i As Integer
    Cadena = ""
    For i = 0 To lst.ListCount - 1
        If Cadena = "" Then
            Cadena = Left(lst.List(i), lst.ItemData(i))
        Else
            Cadena = Cadena & "," & _
                          Left(lst.List(i), lst.ItemData(i))
        End If
    Next i
    PreparaCadRec = Cadena
End Function


Public Sub RecuperaSelecRec(ByVal Key As String, trans As String)
Dim s As String, Vector As Variant, ix As Long
Dim i As Integer, j As Integer, Selec As Integer

    s = trans           '  jeaa 20/09/2003
    If s <> "_VACIO_" Then
        Vector = Split(s, ",")
         Selec = UBound(Vector, 1)
         For i = 0 To Selec
            For j = lstRecar1.ListCount - 1 To 0 Step -1
                If Vector(i) = Left(lstRecar1.List(j), lstRecar1.ItemData(j)) Then
                    'ix = mobjGrupo.AgregarUsuario(.List(i))
                    ix = lstRecar1.ItemData(j)
                    lstRecar2.AddItem lstRecar1.List(j)
                    lstRecar2.ItemData(lstRecar2.NewIndex) = ix
                    lstRecar1.RemoveItem j
                End If
            Next j
         Next i
    End If
End Sub


Private Sub lstTransAct_DblClick()
    Regresa lstTransAct, lstTransCompras
End Sub

Private Sub lstTransNCCompra_DblClick()
    Regresa lstTransNCCompra, lstTransCompras
End Sub


Private Sub lstTransCompraNeta_DblClick()
    Regresa lstTransCompraNeta, lstTransCompras
End Sub

Private Sub lstTransbie_DblClick()
    Regresa lstTransBie, lstTransCompras
End Sub


Private Sub lstTransRem_DblClick()
    Regresa lstTransRem, lstTransCompras
End Sub

Private Sub lstTransVentas_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim DY   ' Declara variable.
   DY = TextHeight("A")   ' Obtiene el alto de una línea.
   lblDragDrop.Move lstTransVentas.Left, lstTransVentas.Top + y - DY / 2, lstTransVentas.Width, DY
   lblDragDrop.Drag   ' Ar
   lblDragDrop.Caption = lstTransVentas.Text
End Sub

Private Sub lstTransCompras_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim DY   ' Declara variable.
   DY = TextHeight("A")   ' Obtiene el alto de una línea.
   lblDragDrop.Move lstTransCompras.Left, lstTransCompras.Top + y - DY / 2, lstTransCompras.Width, DY
   lblDragDrop.Drag   ' Ar
   lblDragDrop.Caption = lstTransCompras.Text
End Sub


Private Sub Form_DragOver(Source As Control, x As Single, y As Single, State As Integer)
   ' Cambia el puntero a no colocar.
   If State = 0 Then Source.MousePointer = 12
   ' Utiliza el puntero predeterminado del mouse.
   If State = 1 Then Source.MousePointer = 0
End Sub

Private Sub lstTransCompraNeta_DragDrop(Source As Control, x As Single, y As Single)
   On Error Resume Next
   carga lstTransCompraNeta, lstTransCompras '1
End Sub

Private Sub lstTransbie_DragDrop(Source As Control, x As Single, y As Single)
   On Error Resume Next
   carga lstTransBie, lstTransCompras '1
End Sub


Private Sub lstTransRem_DragDrop(Source As Control, x As Single, y As Single)
   On Error Resume Next
   carga lstTransRem, lstTransCompras '2
End Sub

Private Sub lstTransact_DragDrop(Source As Control, x As Single, y As Single)
   On Error Resume Next
   carga lstTransAct, lstTransCompras '3
End Sub

Private Sub lstTransNCCompra_DragDrop(Source As Control, x As Single, y As Single)
   On Error Resume Next
   carga lstTransNCCompra, lstTransCompras '3
End Sub


Private Sub LstVentas0_DblClick()
    Regresa LstVentas0, lstTransVentas
End Sub

Private Sub LstVentas12_DblClick()
    Regresa LstVentas12, lstTransVentas
End Sub

Private Sub LstNCVENTAS_DblClick()
    Regresa lstNC_ventas, lstTransVentas
End Sub


Private Sub Lstexpbien_DblClick()
    Regresa lstExpBien, lstTransVentas
End Sub

Private Sub LstExpSer_DblClick()
    Regresa LstExpSer, lstTransVentas
End Sub

Private Sub LstRepGast_DblClick()
    Regresa LstRepGast, lstTransVentas
End Sub



Private Sub LstVentas12_DragDrop(Source As Control, x As Single, y As Single)
   On Error Resume Next
   carga LstVentas12, lstTransVentas '1
End Sub

Private Sub LstVentas0_DragDrop(Source As Control, x As Single, y As Single)
   On Error Resume Next
   carga LstVentas0, lstTransVentas '2
End Sub

Private Sub LstExpBien_DragDrop(Source As Control, x As Single, y As Single)
   On Error Resume Next
   carga lstExpBien, lstTransVentas '2
End Sub

Private Sub Lstnc_Ventas_DragDrop(Source As Control, x As Single, y As Single)
   On Error Resume Next
   carga lstNC_ventas, lstTransVentas '2
   
End Sub


Private Sub LstExpSer_DragDrop(Source As Control, x As Single, y As Single)
   On Error Resume Next
   carga LstExpSer, lstTransVentas '2
End Sub

Private Sub Lstrepgast_DragDrop(Source As Control, x As Single, y As Single)
   On Error Resume Next
   carga LstRepGast, lstTransVentas '2
End Sub



Private Sub carga(lst As ListBox, lst1 As ListBox)
    Dim i As Long, ix As Long
    With lst1
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                ix = .ItemData(i)
                lst.AddItem .List(i)
                lst.ItemData(lst.NewIndex) = ix
                .RemoveItem i
            End If
        Next i
    End With
End Sub


Public Sub RecuperaSeleccion(ByVal Key As String, lst As ListBox, lst1 As ListBox, Optional s As String)
Dim Vector As Variant
Dim i As Integer, j As Integer, Selec As Integer, ix As Long, max As Integer, pos As Integer
Dim trans As String
    If s <> "_VACIO_" Then
        With lst1
            Vector = Split(s, ",")
             Selec = UBound(Vector, 1)
             For i = 0 To Selec
                max = .ListCount - 1
                j = 0
                For j = 0 To max
                    pos = InStr(1, .List(j), " ")
                    'If Vector(i) = Left(.List(j), .ItemData(j)) Then
                    trans = Trim$(Mid$(.List(j), 1, pos - 1))
                    If Vector(i) = trans Then
                        ix = .ItemData(j)
                        lst.AddItem .List(j)
                        lst.ItemData(lst.NewIndex) = ix
                        .RemoveItem j
                        j = max
                    End If
                Next j
             Next i
        End With
    End If
End Sub


Private Sub Regresa(lst As ListBox, lst1 As ListBox)
    Dim i As Long, ix As Long
    On Error GoTo ErrTrap
    With lst
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                ix = .ItemData(i)
                lst1.AddItem .List(i)
                lst1.ItemData(lst1.NewIndex) = ix
                .RemoveItem i
            End If
        Next i
    End With
    Exit Sub
ErrTrap:
    DispErr
End Sub

Private Function PreparaCadenaSeleccion(lst As ListBox) As String
    Dim Cadena As String, i As Integer
    Cadena = ""
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            If Cadena = "" Then
                Cadena = Left(lst.List(i), lst.ItemData(i))
            Else
                Cadena = Cadena & "," & _
                              Left(lst.List(i), lst.ItemData(i))
            End If
        End If
    Next i
    PreparaCadenaSeleccion = Cadena
End Function


'jeaa 25/09/2006 elimina los apostrofes
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

Public Function Inicio103(ByRef objcond As Condicion, ByRef Recargo As String, _
                                            ByRef CP_Ser As String, ByRef CP_Act As String, _
                                            ByRef CP_Dev As String, ByRef VT_Dev As String, _
                                            ByRef VT_12 As String, ByRef VT_Act As String, _
                                            ByRef ret_real As String, ByRef ret_recib As String, _
                                            Reten As String) As Boolean
    Dim KeyTrans As String, KeyRecargo As String, KeyTransRet As String, KeyRet As String, KeyT As String
    Dim trans As String, s As String
    On Error GoTo ErrTrap
    Screen.MousePointer = vbHourglass
    BandAceptado = False
    'visualizar la condicion anterior
    With objcond
        dtpHasta.Format = dtpCustom
        dtpHasta.CustomFormat = "MMMM yyyy"
'        dtpHasta.value = .Fecha1
        If .fecha1 <> 0 Then
            dtpHasta.value = .fecha1
        Else
            dtpHasta.value = Now
        End If
        sst1.TabEnabled(0) = False
        sst1.TabEnabled(1) = False
        sst1.Tab = 1
        fraRetRecibidas.Visible = False
        fraRetRealizada.Visible = False
        Fra103.Visible = True
        fraRecargos.Visible = False
'        'ventas
'        CargaTransxTipoTrans lstTransVentas, 2
'
'        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_Ventas")) > 0 Then
'            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_Ventas")
'            RecuperaSeleccion KeyT, LstVentas12, lstTransVentas, trans
'        End If
'
'        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasAct")) > 0 Then
'            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasAct")
'            RecuperaSeleccion KeyT, LstVentas0, lstTransVentas, trans
'        End If
'
'        'compras
'        CargaTransxTipoTrans lstTransCompras, 1
'
'        'trans = GetSetting(APPNAME, App.Title, "Trans_Com", "_VACIO_")
'        'RecuperaSeleccion KeyT, lstTransCompraNeta, lstTransCompras, trans
'        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_Compras")) > 0 Then
'            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_Compras")
'            RecuperaSeleccion KeyT, lstTransCompraNeta, lstTransCompras, trans
'        End If
'
'
'        'trans = GetSetting(APPNAME, App.Title, "Trans_Ser", "_VACIO_")
'        'RecuperaSeleccion KeyT, lstTransRem, lstTransCompras, trans
'        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasSer")) > 0 Then
'            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasSer")
'            RecuperaSeleccion KeyT, lstTransRem, lstTransCompras, trans
'        End If
''        trans = GetSetting(APPNAME, App.Title, "Trans_Act", "_VACIO_")
''        RecuperaSeleccion KeyT, lstTransAct, lstTransCompras, trans
'        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasAct")) > 0 Then
'            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasAct")
'            RecuperaSeleccion KeyT, lstTransAct, lstTransCompras, trans
'        End If
       
        
        'Retenciones
'        CargaTransxTipoComprobante lstRetRecibidas, 7
'        CargaTransxTipoComprobante lstRetRealizadas, 7
        lblTipoRetencion.Caption = "Tipos de Retencion RENTA"
        CargaTransRetencion LstRetencion103, False
        
'        trans = GetSetting(APPNAME, App.Title, "Reten_Recibidas", "_VACIO_")
'        RecuperaSelec KeyT, lstRetRecibidas, trans
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_RetenRecibidas")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_RetenRecibidas")
            RecuperaSelec KeyT, lstRetRecibidas, trans
        End If

'        trans = GetSetting(APPNAME, App.Title, "Reten_Realizadas", "_VACIO_")
'        RecuperaSelec KeyT, lstRetRealizadas, trans
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_RetenRealizadas")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_RetenRealizadas")
            RecuperaSelec KeyT, lstRetRealizadas, trans
        End If

'        trans = GetSetting(APPNAME, App.Title, "Reten", "_VACIO_")
'        RecuperaSelec KeyT, LstRetencion, trans
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_Reten")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_Reten")
            RecuperaSelec KeyT, LstRetencion, trans
        End If
       
        
        
        
        'carga recargos
        CargaRecargo
'        trans = GetSetting(APPNAME, App.Title, "RecarDesc", "_VACIO_")
'        RecuperaSelecRec KeyRecargo, trans
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_RecarDesc")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_RecarDesc")
            RecuperaSelecRec KeyRecargo, trans
        End If
        
        BandAceptado = False
        Screen.MousePointer = 0
        sst1.Tab = 0
        Me.Show vbModal
'        Si aplastó el botón 'Aceptar'
        If BandAceptado Then



'            'Devuelve las condiciones de búsqueda en el objeto objCondicion en SiiMain
            .fecha1 = dtpHasta.value
            .CodTrans = PreparaCadena(lstTransCompraNeta)
            Recargo = PreparaCadRec(lstRecar2)
            
            'ventas
            VT_12 = PreparaCadena(LstVentas12)
            VT_Act = PreparaCadena(LstVentas0)
            SaveSetting APPNAME, App.Title, "Trans_Ventas12", VT_12
            SaveSetting APPNAME, App.Title, "Trans_VentasAct", VT_Act
            
'            s = PreparaTransParaGnopcion(VT_12)
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_Ventas", VT_12
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_VentasAct", VT_Act
            
            
            CP_Ser = PreparaCadena(lstTransRem)
            CP_Act = PreparaCadena(lstTransAct)
            
            'compras
'            SaveSetting APPNAME, App.Title, "Trans_Com", .CodTrans
'            SaveSetting APPNAME, App.Title, "Trans_Ser", CP_Ser
'            SaveSetting APPNAME, App.Title, "Trans_Act", CP_Act

            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_Compras", .CodTrans
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_ComprasSer", CP_Ser
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_ComprasAct", CP_Act

            

            
            'retenciones
            ret_real = PreparaCadenaSeleccion(lstRetRealizadas)
            ret_recib = PreparaCadenaSeleccion(lstRetRecibidas)
            Reten = PreparaCadenaSeleccion(LstRetencion)
            
'            SaveSetting APPNAME, App.Title, "Reten_Recibidas", ret_recib
'            SaveSetting APPNAME, App.Title, "Reten_Realizadas", ret_real
'            SaveSetting APPNAME, App.Title, "Reten", Reten
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_RetenRecibidas", ret_recib
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_RetenRealizadas", ret_real
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_Reten", Reten
            
            
            'recargos descuentos
'            SaveSetting APPNAME, App.Title, "RecarDesc", Recargo
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_RecarDesc", Recargo
        'Graba en la base
        gobjMain.EmpresaActual.GNOpcion.Grabar
            
        End If
    End With
    Inicio103 = BandAceptado
    Unload Me
    'Set objSiiMain = Nothing
    Exit Function
ErrTrap:
    Screen.MousePointer = 0
    DispErr
    Exit Function
End Function


Public Function Inicio1042010(ByRef objcond As Condicion, ByRef Recargo As String, _
                                            ByRef CP_Ser As String, ByRef CP_Act As String, _
                                            ByRef CP_Dev As String, ByRef VT_Dev As String, _
                                            ByRef VT_12 As String, ByRef VT_Act As String, _
                                            ByRef VT_ExpB As String, ByRef VT_expS As String, _
                                            ByRef VT_RepGas As String, _
                                            ByRef ret_real As String, ByRef ret_recib As String, _
                                            ByRef Reten As String, ByRef NC_Ventas As String, ByRef NC_Compras As String, _
                                            ByRef CP_Bie As String, ByRef CP_Rise As String) As Boolean
    Dim KeyTrans As String, KeyRecargo As String, KeyTransRet As String, KeyRet As String, KeyT As String, KeyRise As String
    Dim KeyTNC As String, KeyTNCC As String
    Dim trans As String, s As String
    On Error GoTo ErrTrap
    Screen.MousePointer = vbHourglass
    BandAceptado = False
    'visualizar la condicion anterior
    With objcond
        dtpHasta.Format = dtpCustom
        dtpHasta.CustomFormat = "MMMM yyyy"
'        dtpHasta.value = .Fecha1
        If .fecha1 <> 0 Then
            dtpHasta.value = .fecha1
        Else
            dtpHasta.value = Now
        End If
        chkFactor.value = IIf(.IncluirCero, vbChecked, vbUnchecked)
        
        'ventas
        CargaTransxTipoTrans lstTransVentas, 2, 4
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_Ventas")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_Ventas")
            RecuperaSeleccion KeyT, LstVentas12, lstTransVentas, trans
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasAct")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasAct")
            RecuperaSeleccion KeyT, LstVentas0, lstTransVentas, trans
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasExpB")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasExpB")
            RecuperaSeleccion KeyT, lstExpBien, lstTransVentas, trans
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasExpS")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasExpS")
            RecuperaSeleccion KeyT, LstExpSer, lstTransVentas, trans
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasRepGas")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasRepGas")
            RecuperaSeleccion KeyT, LstRepGast, lstTransVentas, trans
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_NCVentas")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_NCVentas")
            RecuperaSeleccion KeyTNC, lstNC_ventas, lstTransVentas, trans
        End If
              
        'compras
        CargaTransxTipoTrans lstTransCompras, 1
       
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_Compras")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_Compras")
            RecuperaSeleccion KeyT, lstTransCompraNeta, lstTransCompras, trans
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasSer")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasSer")
            RecuperaSeleccion KeyT, lstTransRem, lstTransCompras, trans
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasAct")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasAct")
            RecuperaSeleccion KeyT, lstTransAct, lstTransCompras, trans
        End If
       
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_NCCompras")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_NCCompras")
            RecuperaSeleccion KeyTNCC, lstTransNCCompra, lstTransCompras, trans
        End If
       
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasBie")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasBie")
            RecuperaSeleccion KeyT, lstTransBie, lstTransCompras, trans
        End If
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasRISE")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasRISE")
            RecuperaSeleccion KeyT, lstRise, lstTransCompras, trans
        End If
        'Retenciones
        CargaTransxTipoComprobante lstRetRecibidas, 7
        CargaTransxTipoComprobante lstRetRealizadas, 7
        CargaTransRetencion LstRetencion, True
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_RetenRecibidas")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_RetenRecibidas")
            RecuperaSelec KeyT, lstRetRecibidas, trans
        End If

        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_RetenRealizadas")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_RetenRealizadas")
            RecuperaSelec KeyT, lstRetRealizadas, trans
        End If

        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_Reten")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_Reten")
            RecuperaSelec KeyT, LstRetencion, trans
        End If
                    
        
        
        'carga recargos
        CargaRecargo
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_RecarDesc")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_RecarDesc")
            RecuperaSelecRec KeyRecargo, trans
        End If
        
        BandAceptado = False
        Screen.MousePointer = 0
        sst1.Tab = 0
       Me.Show vbModal
'        Si aplastó el botón 'Aceptar'
        If BandAceptado Then

'            'Devuelve las condiciones de búsqueda en el objeto objCondicion en SiiMain
            .fecha1 = dtpHasta.value
            .CodTrans = PreparaCadena(lstTransCompraNeta)
            .IncluirCero = (chkFactor.value = vbChecked)
            Recargo = PreparaCadRec(lstRecar2)
            
            'ventas
            VT_12 = PreparaCadena(LstVentas12)
            VT_Act = PreparaCadena(LstVentas0)
            VT_ExpB = PreparaCadena(lstExpBien)
            VT_expS = PreparaCadena(LstExpSer)
            VT_RepGas = PreparaCadena(LstRepGast)
            NC_Ventas = PreparaCadena(lstNC_ventas)
            NC_Compras = PreparaCadena(lstTransNCCompra)
            
            
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_Ventas", VT_12
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_VentasAct", VT_Act
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_VentasExpB", VT_ExpB
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_VentasExpS", VT_expS
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_VentasExpRepGas", VT_RepGas
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_NCVentas", NC_Ventas
            
            CP_Ser = PreparaCadena(lstTransRem)
            CP_Act = PreparaCadena(lstTransAct)
            CP_Bie = PreparaCadena(lstTransBie)
            CP_Rise = PreparaCadena(lstRise)
             'CP_SER le voy a utilizar para reposiciones de gastos en compras
            'compras
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_Compras", .CodTrans
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_ComprasSer", CP_Ser
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_ComprasAct", CP_Act
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_NCCompras", NC_Compras
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_ComprasBie", CP_Bie
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_ComprasRise", CP_Rise
         
           
            'retenciones
            ret_real = PreparaCadenaSeleccion(lstRetRealizadas)
            ret_recib = PreparaCadenaSeleccion(lstRetRecibidas)
            Reten = PreparaCadenaSeleccion(LstRetencion)
            
'            SaveSetting APPNAME, App.Title, "Reten_Recibidas", ret_recib
'            SaveSetting APPNAME, App.Title, "Reten_Realizadas", ret_real
'            SaveSetting APPNAME, App.Title, "Reten", Reten
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_RetenRecibidas", ret_recib
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_RetenRealizadas", ret_real
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_Reten", Reten
            
            
            'recargos descuentos
'            SaveSetting APPNAME, App.Title, "RecarDesc", Recargo
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_RecarDesc", Recargo
        'Graba en la base
        gobjMain.EmpresaActual.GNOpcion.Grabar
            
        End If
    End With
    Inicio1042010 = BandAceptado
    Unload Me
    'Set objSiiMain = Nothing
    Exit Function
ErrTrap:
    Screen.MousePointer = 0
    DispErr
    Exit Function
End Function


Private Sub lstNC_ExpB_DblClick()
    Regresa lstNC_ExpB, lstTransVentas
End Sub

Private Sub lstNC_Exps_DblClick()
    Regresa lstNC_ExpS, lstTransVentas
End Sub


Private Sub lstNC_ExpB_DragDrop(Source As Control, x As Single, y As Single)
   On Error Resume Next
   carga lstNC_ExpB, lstTransVentas '2
End Sub

Private Sub lstNC_Exps_DragDrop(Source As Control, x As Single, y As Single)
   On Error Resume Next
   carga lstNC_ExpS, lstTransVentas '2
End Sub

Public Function Inicio104_2013(ByRef objcond As Condicion, ByRef Recargo As String, _
                                            ByRef CP_Ser As String, _
                                            ByRef CP_Act As String, _
                                            ByRef CP_Dev As String, _
                                            ByRef VT_Dev As String, _
                                            ByRef VT_12 As String, _
                                            ByRef VT_Act As String, _
                                            ByRef VT_ExpB As String, _
                                            ByRef VT_expS As String, _
                                            ByRef VT_RepGas As String, _
                                            ByRef ret_real As String, _
                                            ByRef ret_recib As String, _
                                            ByRef Reten As String, _
                                            ByRef NC_ExpB As String, _
                                            ByRef NC_ExpS As String, _
                                            ByRef NC_Ventas As String, _
                                            ByRef NC_Compras As String, _
                                            ByRef CP_Bie As String, _
                                            ByRef CP_Rise As String) As Boolean
    Dim KeyTrans As String, KeyRecargo As String, KeyTransRet As String, KeyRet As String, KeyT As String
    Dim KeyTNC As String, KeyTNCC As String, KeyTNCEB As String, KeyTNCES As String
    Dim trans As String, s As String
    On Error GoTo ErrTrap
    Screen.MousePointer = vbHourglass
    BandAceptado = False
    'visualizar la condicion anterior
    
        With objcond
        dtpHasta.Format = dtpCustom
        dtpHasta.CustomFormat = "MMMM yyyy"
'        dtpHasta.value = .Fecha1
        If .fecha1 <> 0 Then
            dtpHasta.value = .fecha1
        Else
            dtpHasta.value = Now
        End If
        chkFactor.value = IIf(.IncluirCero, vbChecked, vbUnchecked)
        
        'ventas
        CargaTransxTipoTrans lstTransVentas, 2, 4
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_Ventas")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_Ventas")
            RecuperaSeleccion KeyT, LstVentas12, lstTransVentas, trans
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasAct")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasAct")
            RecuperaSeleccion KeyT, LstVentas0, lstTransVentas, trans
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasExpB")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasExpB")
            RecuperaSeleccion KeyT, lstExpBien, lstTransVentas, trans
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasExpS")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasExpS")
            RecuperaSeleccion KeyT, LstExpSer, lstTransVentas, trans
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasRepGas")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_VentasRepGas")
            RecuperaSeleccion KeyT, LstRepGast, lstTransVentas, trans
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_NCVentas")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_NCVentas")
            RecuperaSeleccion KeyTNC, lstNC_ventas, lstTransVentas, trans
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_NCExpB")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_NCExpB")
            RecuperaSeleccion KeyTNCEB, lstNC_ExpB, lstTransVentas, trans
        End If

        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_NCExpS")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_NCExpS")
            RecuperaSeleccion KeyTNCES, lstNC_ExpS, lstTransVentas, trans
        End If
        
              
        'compras
        CargaTransxTipoTrans lstTransCompras, 1
       
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_Compras")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_Compras")
            RecuperaSeleccion KeyT, lstTransCompraNeta, lstTransCompras, trans
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasSer")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasSer")
            RecuperaSeleccion KeyT, lstTransRem, lstTransCompras, trans
        End If
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasAct")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasAct")
            RecuperaSeleccion KeyT, lstTransAct, lstTransCompras, trans
        End If
       
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_NCCompras")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_NCCompras")
            RecuperaSeleccion KeyTNCC, lstTransNCCompra, lstTransCompras, trans
        End If
       
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasBie")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasBie")
            RecuperaSeleccion KeyT, lstTransBie, lstTransCompras, trans
        End If
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasRISE")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_ComprasRISE")
            RecuperaSeleccion KeyT, lstRise, lstTransCompras, trans
        End If
        'Retenciones
        CargaTransxTipoComprobante lstRetRecibidas, 7
        CargaTransxTipoComprobante lstRetRealizadas, 7
        CargaTransRetencion LstRetencion, True
        
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_RetenRecibidas")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_RetenRecibidas")
            RecuperaSelec KeyT, lstRetRecibidas, trans
        End If

        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_RetenRealizadas")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_RetenRealizadas")
            RecuperaSelec KeyT, lstRetRealizadas, trans
        End If

        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_Reten")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_Reten")
            RecuperaSelec KeyT, LstRetencion, trans
        End If
                    
        
        
        'carga recargos
        CargaRecargo
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_RecarDesc")) > 0 Then
            trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("104_RecarDesc")
            RecuperaSelecRec KeyRecargo, trans
        End If
        
        BandAceptado = False
        Screen.MousePointer = 0
        sst1.Tab = 0
       Me.Show vbModal
'        Si aplastó el botón 'Aceptar'
        If BandAceptado Then

'            'Devuelve las condiciones de búsqueda en el objeto objCondicion en SiiMain
            .fecha1 = dtpHasta.value
            .CodTrans = PreparaCadena(lstTransCompraNeta)
            .IncluirCero = (chkFactor.value = vbChecked)
            Recargo = PreparaCadRec(lstRecar2)
            
            'ventas
            VT_12 = PreparaCadena(LstVentas12)
            VT_Act = PreparaCadena(LstVentas0)
            VT_ExpB = PreparaCadena(lstExpBien)
            VT_expS = PreparaCadena(LstExpSer)
            VT_RepGas = PreparaCadena(LstRepGast)
            NC_Ventas = PreparaCadena(lstNC_ventas)
            NC_Compras = PreparaCadena(lstTransNCCompra)
            NC_ExpB = PreparaCadena(lstNC_ExpB)
            NC_ExpS = PreparaCadena(lstNC_ExpS)

            
            
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_Ventas", VT_12
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_VentasAct", VT_Act
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_VentasExpB", VT_ExpB
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_VentasExpS", VT_expS
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_VentasExpRepGas", VT_RepGas
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_NCVentas", NC_Ventas
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_NCExpB", NC_ExpB
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_NCExpS", NC_ExpS

            
            CP_Ser = PreparaCadena(lstTransRem)
            CP_Act = PreparaCadena(lstTransAct)
            CP_Bie = PreparaCadena(lstTransBie)
            CP_Rise = PreparaCadena(lstRise)
             'CP_SER le voy a utilizar para reposiciones de gastos en compras
            'compras
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_Compras", .CodTrans
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_ComprasSer", CP_Ser
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_ComprasAct", CP_Act
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_NCCompras", NC_Compras
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_ComprasBie", CP_Bie
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_ComprasRise", CP_Rise
         
           
            'retenciones
            ret_real = PreparaCadenaSeleccion(lstRetRealizadas)
            ret_recib = PreparaCadenaSeleccion(lstRetRecibidas)
            Reten = PreparaCadenaSeleccion(LstRetencion)
            
'            SaveSetting APPNAME, App.Title, "Reten_Recibidas", ret_recib
'            SaveSetting APPNAME, App.Title, "Reten_Realizadas", ret_real
'            SaveSetting APPNAME, App.Title, "Reten", Reten
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_RetenRecibidas", ret_recib
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_RetenRealizadas", ret_real
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_Reten", Reten
            
            
            'recargos descuentos
'            SaveSetting APPNAME, App.Title, "RecarDesc", Recargo
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "104_RecarDesc", Recargo
        'Graba en la base
        gobjMain.EmpresaActual.GNOpcion.Grabar
            
        End If
    End With

    
    
    Inicio104_2013 = BandAceptado
    Unload Me
    'Set objSiiMain = Nothing
    Exit Function
ErrTrap:
    Screen.MousePointer = 0
    DispErr
    Exit Function
End Function


