VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmB_Form104_old 
   Caption         =   "Condiciones de Busqueda"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5580
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   5580
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraRecargos 
      Caption         =   "Recargos y Descuentos antes del IVA"
      Height          =   1575
      Left            =   180
      TabIndex        =   32
      Top             =   5820
      Width           =   5175
      Begin VB.ListBox lstRecar1 
         Height          =   1230
         Left            =   120
         TabIndex        =   36
         Top             =   270
         Width           =   2055
      End
      Begin VB.ListBox lstRecar2 
         Height          =   1230
         Left            =   3000
         TabIndex        =   35
         Top             =   285
         Width           =   2055
      End
      Begin VB.CommandButton cmdAdd2 
         Caption         =   "&>>"
         Height          =   375
         Left            =   2280
         TabIndex        =   34
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdResta2 
         Caption         =   "&<<"
         Height          =   375
         Left            =   2280
         TabIndex        =   33
         Top             =   825
         Width           =   615
      End
   End
   Begin TabDlg.SSTab sst1 
      Height          =   5055
      Left            =   60
      TabIndex        =   17
      Top             =   720
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   8916
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Compras"
      TabPicture(0)   =   "frmB_Fom104old.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTrans1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Ventas"
      TabPicture(1)   =   "frmB_Fom104old.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(2)=   "Frame6"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmB_Fom104old.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame Frame6 
         Caption         =   "Transacciones Devoluciones"
         Height          =   1080
         Left            =   -74880
         TabIndex        =   30
         Top             =   2700
         Width           =   5175
         Begin VB.ListBox LstDevVentas 
            Height          =   765
            IntegralHeight  =   0   'False
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   31
            Top             =   225
            Width           =   4935
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Transacciones de Ventas Netas"
         Height          =   1140
         Left            =   -74880
         TabIndex        =   28
         Top             =   360
         Width           =   5175
         Begin VB.ListBox LstVentas12 
            Height          =   765
            IntegralHeight  =   0   'False
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   29
            Top             =   225
            Width           =   4935
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Transacciones de Ventas Activos"
         Height          =   1080
         Left            =   -74880
         TabIndex        =   26
         Top             =   1560
         Width           =   5175
         Begin VB.ListBox LstVentas0 
            Height          =   765
            IntegralHeight  =   0   'False
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   27
            Top             =   225
            Width           =   4935
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Transacciones Devoluciones"
         Height          =   1080
         Left            =   120
         TabIndex        =   24
         Top             =   3900
         Width           =   5175
         Begin VB.ListBox lstTransDev 
            Height          =   765
            IntegralHeight  =   0   'False
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   25
            Top             =   225
            Width           =   4935
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Transacciones de Compras Activos"
         Height          =   1140
         Left            =   120
         TabIndex        =   22
         Top             =   2700
         Width           =   5175
         Begin VB.ListBox lstTransAct 
            Height          =   765
            IntegralHeight  =   0   'False
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   23
            Top             =   225
            Width           =   4935
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Transacciones de Compras Servicios"
         Height          =   1080
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   5175
         Begin VB.ListBox lstTransSer 
            Height          =   765
            IntegralHeight  =   0   'False
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   21
            Top             =   225
            Width           =   4935
         End
      End
      Begin VB.Frame fraTrans1 
         Caption         =   "Transacciones de Compras Netas"
         Height          =   1140
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   5175
         Begin VB.ListBox lstTransCompraNeta 
            Height          =   765
            IntegralHeight  =   0   'False
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   19
            Top             =   225
            Width           =   4935
         End
      End
   End
   Begin VB.Frame fraCobro 
      Caption         =   "Códigos de Retenciones de IVA"
      Height          =   1590
      Left            =   6420
      TabIndex        =   10
      Top             =   3420
      Width           =   5175
      Begin VB.ListBox lstBienes 
         Height          =   1035
         Left            =   120
         TabIndex        =   14
         Top             =   450
         Width           =   2055
      End
      Begin VB.ListBox lstServicios 
         Height          =   1035
         Left            =   3000
         TabIndex        =   13
         Top             =   450
         Width           =   2055
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&>>"
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   570
         Width           =   615
      End
      Begin VB.CommandButton cmdResta 
         Caption         =   "&<<"
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   1050
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "SERVICIOS"
         Height          =   255
         Left            =   3000
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "BIENES"
         Height          =   255
         Left            =   105
         TabIndex        =   15
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
      ScaleWidth      =   5580
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7470
      Width           =   5580
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   2715
         TabIndex        =   5
         Top             =   75
         Width           =   1452
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar -F5"
         Height          =   372
         Left            =   1035
         TabIndex        =   4
         Top             =   75
         Width           =   1452
      End
   End
   Begin VB.Frame fraFecha 
      Caption         =   "Rango de Fecha"
      Height          =   705
      Left            =   180
      TabIndex        =   6
      Top             =   -15
      Width           =   5175
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   360
         Left            =   840
         TabIndex        =   2
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         Format          =   51576835
         UpDown          =   -1  'True
         CurrentDate     =   36891
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   360
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   1692
         _ExtentX        =   2990
         _ExtentY        =   635
         _Version        =   393216
         Format          =   51576833
         CurrentDate     =   36526
      End
      Begin VB.Label lblFechaHasta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "H&asta  "
         Height          =   195
         Left            =   2760
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   276
         Width           =   564
      End
   End
   Begin VB.Frame fraTrans2 
      Caption         =   "Transacciones de Retención"
      Height          =   1275
      Left            =   6120
      TabIndex        =   0
      Top             =   5400
      Width           =   5175
      Begin VB.ListBox lstTrans2 
         Height          =   975
         IntegralHeight  =   0   'False
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   210
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmB_Form104_old"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private BandAceptado As Boolean
Private mobjSiiMain As SiiMain


Public Function Inicio(ByRef objcond As Condicion, ByRef Recargo As String, _
                                            ByRef CP_Ser As String, ByRef CP_Act As String, _
                                            ByRef CP_Dev As String, ByRef VT_Dev As String, _
                                            ByRef VT_12 As String, ByRef VT_0 As String) As Boolean
    Dim KeyTrans As String, KeyRecargo As String, KeyTransRet As String, KeyRet As String, KeyT As String
    Dim trans As String
    On Error GoTo Errtrap
    Screen.MousePointer = vbHourglass
    BandAceptado = False
    'visualizar la condicion anterior
    With objcond
        dtpHasta.Format = dtpCustom
        dtpHasta.CustomFormat = "MMMM yyyy"
        dtpHasta.value = .Fecha1
        If .Fecha1 <> 0 Then dtpHasta.value = .Fecha1

        'compras
        CargaTipoTrans "IV", lstTransCompraNeta
        CargaTipoTrans "IV", lstTransSer
        CargaTipoTrans "IV", lstTransAct
        CargaTipoTrans "IV", lstTransDev
        
        trans = GetSetting(APPNAME, App.Title, "Trans_Com", "_VACIO_")
        RecuperaSelec KeyT, lstTransCompraNeta, trans
        trans = GetSetting(APPNAME, App.Title, "Trans_Ser", "_VACIO_")
        RecuperaSelec KeyT, lstTransSer, trans
        trans = GetSetting(APPNAME, App.Title, "Trans_Act", "_VACIO_")
        RecuperaSelec KeyT, lstTransAct, trans
        trans = GetSetting(APPNAME, App.Title, "Trans_Dev", "_VACIO_")
        RecuperaSelec KeyT, lstTransDev, trans
       
        'ventas
        CargaTipoTrans "IV", LstVentas12
        CargaTipoTrans "IV", LstVentas0
        CargaTipoTrans "IV", LstDevVentas
        
        trans = GetSetting(APPNAME, App.Title, "Trans_Ventas12", "_VACIO_")
        RecuperaSelec KeyT, LstVentas12, trans
        trans = GetSetting(APPNAME, App.Title, "Trans_Ventas0", "_VACIO_")
        RecuperaSelec KeyT, LstVentas0, trans
        trans = GetSetting(APPNAME, App.Title, "Trans_DevVentas", "_VACIO_")
        RecuperaSelec KeyT, LstDevVentas, trans
        
        
        
        
        'carga recargos
        CargaRecargo
        trans = GetSetting(APPNAME, App.Title, "RecarDesc", "_VACIO_")
        RecuperaSelecRec KeyRecargo, trans
        
        BandAceptado = False
        Screen.MousePointer = 0
        sst1.Tab = 0
        Me.Show vbModal
'        Si aplastó el botón 'Aceptar'
        If BandAceptado Then
'            'Devuelve las condiciones de búsqueda en el objeto objCondicion en SiiMain
            .Fecha1 = dtpHasta.value
            .CodTrans = PreparaCadena(lstTransCompraNeta)
            Recargo = PreparaCadRec(lstRecar2)
            CP_Ser = PreparaCadena(lstTransSer)
            CP_Act = PreparaCadena(lstTransAct)
            CP_Dev = PreparaCadena(lstTransDev)
            
            'compras
            SaveSetting APPNAME, App.Title, "Trans_Com", .CodTrans
            SaveSetting APPNAME, App.Title, "Recar", Recargo
            SaveSetting APPNAME, App.Title, "Trans_Ser", CP_Ser
            SaveSetting APPNAME, App.Title, "Trans_Act", CP_Act
            SaveSetting APPNAME, App.Title, "Trans_Dev", CP_Dev
            
            'ventas
            VT_12 = PreparaCadena(LstVentas12)
            VT_0 = PreparaCadena(LstVentas0)
            VT_Dev = PreparaCadena(LstDevVentas)
            SaveSetting APPNAME, App.Title, "Trans_Ventas12", VT_12
            SaveSetting APPNAME, App.Title, "Trans_Ventas0", VT_0
            SaveSetting APPNAME, App.Title, "Trans_DevVentas", VT_Dev
            
            'recargos descuentos
            SaveSetting APPNAME, App.Title, "RecarDesc", Recargo
            
            
        End If
    End With
    Inicio = BandAceptado
    Unload Me
    'Set objSiiMain = Nothing
    Exit Function
Errtrap:
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
    On Error GoTo Errtrap
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
Errtrap:
    DispErr
End Sub

Private Sub cmdCancelar_Click()
    BandAceptado = False
    Me.Hide
End Sub


Private Sub cmdResta2_Click()
    Dim i As Long, ix As Long
    On Error GoTo Errtrap
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
Errtrap:
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
    On Error GoTo Errtrap
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
Errtrap:
    DispErr
End Sub

Private Sub cmdResta_Click()
    Dim i As Long, ix As Long
    On Error GoTo Errtrap
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
Errtrap:
    DispErr
End Sub


Private Sub lstBienes_DblClick()
    cmdAdd_Click
End Sub

Private Sub lstRecar1_DblClick()
    cmdAdd2_Click
End Sub

Private Sub lstRecar2_DblClick()
    cmdResta2_Click
End Sub

Private Sub lstServicios_DblClick()
    cmdResta_Click
End Sub

Private Function PreparaCadena(lst As ListBox) As String
    Dim cadena As String, i As Integer
    cadena = ""
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            If cadena = "" Then
                cadena = Left(lst.List(i), lst.ItemData(i))
            Else
                cadena = cadena & "," & _
                              Left(lst.List(i), lst.ItemData(i))
            End If
        End If
    Next i
    PreparaCadena = cadena
End Function


Private Sub CargaRecargo()
    Dim rs As Recordset
    Set rs = gobjMain.EmpresaActual.ListaIVRecargo(True)
    With rs
        If Not (.EOF) Then
            .MoveFirst
            Do Until .EOF
                lstRecar1.AddItem !CodRecargo & "  " & !Descripcion
                lstRecar1.ItemData(lstRecar1.NewIndex) = Len(!CodRecargo)
               .MoveNext
           Loop
           
            lstRecar1.AddItem "SUBT" & "  " & "Subtotal"
            lstRecar1.ItemData(lstRecar1.NewIndex) = Len("SUBT")
            
        End If
    End With
    rs.Close
End Sub

Private Function PreparaCadRec(lst As ListBox) As String
    Dim cadena As String, i As Integer
    cadena = ""
    For i = 0 To lst.ListCount - 1
        If cadena = "" Then
            cadena = Left(lst.List(i), lst.ItemData(i))
        Else
            cadena = cadena & "," & _
                          Left(lst.List(i), lst.ItemData(i))
        End If
    Next i
    PreparaCadRec = cadena
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


