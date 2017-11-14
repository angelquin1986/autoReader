VERSION 5.00
Begin VB.Form frmB_Trans 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6450
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5385
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   400
      Left            =   2640
      TabIndex        =   2
      Top             =   6000
      Width           =   1200
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar -F5"
      Height          =   400
      Left            =   1320
      TabIndex        =   1
      Top             =   6000
      Width           =   1200
   End
   Begin VB.Frame fra 
      Caption         =   "Transacciones"
      Height          =   2235
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   5175
      Begin VB.ListBox lst 
         Height          =   1845
         IntegralHeight  =   0   'False
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame fraCobro 
      Caption         =   "Códigos de Retenciones de IVA"
      Height          =   1695
      Left            =   60
      TabIndex        =   5
      Top             =   2340
      Width           =   5175
      Begin VB.CommandButton cmdResta 
         Caption         =   "&<<"
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&>>"
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   600
         Width           =   615
      End
      Begin VB.ListBox lstServicios 
         Height          =   1035
         Left            =   3000
         TabIndex        =   7
         Top             =   510
         Width           =   2055
      End
      Begin VB.ListBox lstBienes 
         Height          =   1035
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "BIENES"
         Height          =   255
         Left            =   105
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "SERVICIOS"
         Height          =   255
         Left            =   3000
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame fraRetencion 
      Caption         =   "Transaccion Retencion"
      Height          =   1815
      Left            =   60
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   5175
      Begin VB.ListBox lstRet 
         Height          =   1485
         IntegralHeight  =   0   'False
         Left            =   60
         Style           =   1  'Checkbox
         TabIndex        =   12
         Top             =   240
         Width           =   4995
      End
   End
End
Attribute VB_Name = "frmB_Trans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BandAceptado As Boolean
Private fecha1 As Date
Private fecha2 As Date
Dim Key As String


Public Function Inicio(ByVal objSiiMain As SiiMain, cad As String, fecha As Date) As Boolean
    Dim i As Integer, j As Integer
    Dim trans As String, KeyTrans As String, KeyTransRet As String, KeyRecargo As String, KeyAnulado As String
    Me.tag = Name  'nombre del reporte
    If DatePart("m", fecha) = 12 Then
        fecha1 = "01/" & DatePart("m", fecha) & "/" & DatePart("yyyy", fecha)
        fecha2 = DateAdd("yyyy", 1, DateAdd("d", -1, ("01/" & DatePart("m", DateAdd("m", 1, fecha)) & "/" & DatePart("yyyy", fecha))))
    Else
        fecha1 = "01/" & DatePart("m", fecha) & "/" & DatePart("yyyy", fecha)
        fecha2 = DateAdd("d", -1, ("01/" & DatePart("m", DateAdd("m", 1, fecha)) & "/" & DatePart("yyyy", fecha)))
    End If
    With objSiiMain.objCondicion
'        If .fecha1 <> 0 Then dtpDesde.value = fecha1
'        If .fecha2 <> 0 Then dtpHasta.value = fecha2
        .fecha1 = fecha1
        .fecha2 = fecha2
        
        ''''CargaRetencion lstBienes
        BandAceptado = False
        Me.Caption = "Condiciones de Busqueda "
        Select Case cad
           Case "IMPCPI"
                CargaTipoTransxTipoTrans "IV", lst, "'1'"
                CargaTipoTransxTipoTrans "IV", lstRet, "'1'"
                KeyTrans = Name & "_Trans" & "_" & objSiiMain.EmpresaActual.CodEmpresa & "_CPI"
                trans = GetSetting(APPNAME, App.Title, KeyTrans)
                RecuperaSelec KeyTrans, lst, trans
                
                KeyTransRet = Name & "_Trans" & "_" & objSiiMain.EmpresaActual.CodEmpresa & "_RTP"
                trans = GetSetting(APPNAME, App.Title, KeyTransRet)
                RecuperaSelec KeyTransRet, lstRet, trans
                        
                fraRetencion.Visible = True
                fra.Caption = "Transacciones de Compra"
                
                fraCobro.Visible = True
                fraCobro.Caption = "Recargos Descuentos antes del IVA"
                CargaRecargos
                KeyRecargo = Name & "_Trans" & "_" & objSiiMain.EmpresaActual.CodEmpresa & "_REC"
                trans = GetSetting(APPNAME, App.Title, KeyRecargo)
                RecuperaSelecRec KeyRecargo, trans
                
           Case "IMPCPI2013"
                CargaTipoTransxTipoTransTipoComp "IV", lst, "'1'", "'1','2','3','4'"
                CargaTipoTransxTipoTransTipoComp "IV", lstRet, "'1'", "'7'"
                

                KeyTrans = Name & "_Trans" & "_" & objSiiMain.EmpresaActual.CodEmpresa & "_CPI"
                'trans = GetSetting(APPNAME, App.Title, KeyTrans)
                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyTrans)
                RecuperaSelec KeyTrans, lst, trans
                

                KeyTransRet = Name & "_Trans" & "_" & objSiiMain.EmpresaActual.CodEmpresa & "_RTP"
                'trans = GetSetting(APPNAME, App.Title, KeyTransRet)
                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyTransRet)
                RecuperaSelec KeyTransRet, lstRet, trans
                        
                fraRetencion.Visible = True
                fra.Caption = "Transacciones de Compra"
                
                fraCobro.Visible = True
                fraCobro.Caption = "Recargos Descuentos antes del IVA"
                CargaRecargos
                
                
                KeyRecargo = Name & "_Trans" & "_" & objSiiMain.EmpresaActual.CodEmpresa & "_REC"
                'trans = GetSetting(APPNAME, App.Title, KeyRecargo)
                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyRecargo)
                RecuperaSelecRec KeyRecargo, trans
                
            Case "IMPFC"
                CargaTipoTransxTipoTrans "IV", lst, "'2'"
                CargaTipoTrans "TS", lstRet
                KeyTrans = Name & "_Trans" & "_" & objSiiMain.EmpresaActual.CodEmpresa & "_FC"
                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyTrans)
                'trans = GetSetting(APPNAME, App.Title, KeyTrans)
                gobjMain.EmpresaActual.GNOpcion.ObtenerValor (KeyTrans)
                RecuperaSelec KeyTrans, lst, trans
                
                KeyTransRet = Name & "_Trans" & "_" & objSiiMain.EmpresaActual.CodEmpresa & "_RTC"
                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyTransRet)
                'trans = GetSetting(APPNAME, App.Title, KeyTransRet)
                RecuperaSelec KeyTransRet, lstRet, trans
                        
                fraRetencion.Visible = True
                fra.Caption = "Transacciones de Compra"
                
                fraCobro.Visible = True
                fraCobro.Caption = "Recargos Descuentos antes del IVA"
                CargaRecargos
                KeyRecargo = Name & "_Trans" & "_" & objSiiMain.EmpresaActual.CodEmpresa & "_REC_FC"
                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyRecargo)
                'trans = GetSetting(APPNAME, App.Title, KeyRecargo)
                RecuperaSelecRec KeyRecargo, trans
            Case "IMPFCxE"
                CargaTipoTransxTipoTrans "IV", lst, "'2'"
                lstRet.Enabled = False
                KeyTrans = Name & "_Trans" & "_" & objSiiMain.EmpresaActual.CodEmpresa & "_FC"
                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyTrans)
                'trans = GetSetting(APPNAME, App.Title, KeyTrans)
                RecuperaSelec KeyTrans, lst, trans
                
                fraRetencion.Visible = False
                fra.Caption = ""
                
                fraCobro.Visible = True
                fraCobro.Caption = "Recargos Descuentos antes del IVA"
                CargaRecargos
                KeyRecargo = Name & "_Trans" & "_" & objSiiMain.EmpresaActual.CodEmpresa & "_REC_FC"
                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyRecargo)
                'trans = GetSetting(APPNAME, App.Title, KeyRecargo)
                RecuperaSelecRec KeyRecargo, trans
            Case "IMPFCEXP"
                CargaTipoTransxTipoTrans "IV", lst, "'2'"
                CargaTipoTrans "TS", lstRet
                KeyTrans = Name & "_Trans" & "_" & objSiiMain.EmpresaActual.CodEmpresa & "_FCEXP"
                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyTrans)
                'trans = GetSetting(APPNAME, App.Title, KeyTrans)
                gobjMain.EmpresaActual.GNOpcion.ObtenerValor (KeyTrans)
                RecuperaSelec KeyTrans, lst, trans
                
                KeyTransRet = Name & "_Trans" & "_" & objSiiMain.EmpresaActual.CodEmpresa & "_RTCEXP"
                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyTransRet)
                'trans = GetSetting(APPNAME, App.Title, KeyTransRet)
                RecuperaSelec KeyTransRet, lstRet, trans
                        
                fraRetencion.Visible = True
                fra.Caption = "Transacciones de Compra"
                
                fraCobro.Visible = True
                fraCobro.Caption = "Recargos Descuentos antes del IVA"
                CargaRecargos
                KeyRecargo = Name & "_Trans" & "_" & objSiiMain.EmpresaActual.CodEmpresa & "_REC_FCEXP"
                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyRecargo)
                'trans = GetSetting(APPNAME, App.Title, KeyRecargo)
                RecuperaSelecRec KeyRecargo, trans
            
            Case "IMPAN"
                CargaTipoTransxTipoTrans "IV", lst, "'1','2'"
                fraCobro.Visible = False
                fraRetencion.Visible = False
                fra.Height = 5800
                lst.Height = 5500
                KeyAnulado = Name & "_Trans" & "_" & objSiiMain.EmpresaActual.CodEmpresa & "_ANU"
                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyAnulado)
                'trans = GetSetting(APPNAME, App.Title, KeyAnulado)
                RecuperaSelec KeyAnulado, lst, trans
            Case "IMPDD"
                CargaTipoTransxTipoTrans "", lst, "'2' "
                'CargaTipoTrans "TS", lstRet
                KeyTrans = Name & "_Trans" & "_" & objSiiMain.EmpresaActual.CodEmpresa & "_FCDD"
                trans = GetSetting(APPNAME, App.Title, KeyTrans)
                RecuperaSelec KeyTrans, lst, trans
                
                        
                fraRetencion.Visible = False
                fra.Caption = "Transacciones de Venta"
                
                fraCobro.Visible = False
                lst.Height = 5600
                fra.Height = 5900
            Case "IMPFCICE"
                CargaTipoTransxTipoTrans "IV", lst, "'2'"
                'CargaTipoTrans "TS", lstRet
                KeyTrans = Name & "_Trans" & "_" & objSiiMain.EmpresaActual.CodEmpresa & "_FC"
                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyTrans)
                'trans = GetSetting(APPNAME, App.Title, KeyTrans)
                gobjMain.EmpresaActual.GNOpcion.ObtenerValor (KeyTrans)
                RecuperaSelec KeyTrans, lst, trans
                
'                KeyTransRet = Name & "_Trans" & "_" & objSiiMain.EmpresaActual.CodEmpresa & "_RTC"
'                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyTransRet)
'                'trans = GetSetting(APPNAME, App.Title, KeyTransRet)
'                RecuperaSelec KeyTransRet, lstRet, trans
                        
                fraRetencion.Visible = False
                fra.Caption = ""
                
                fraCobro.Visible = True
                fraCobro.Caption = "Recargos Descuentos antes del IVA"
                CargaRecargos
                KeyRecargo = Name & "_Trans" & "_" & objSiiMain.EmpresaActual.CodEmpresa & "_REC_FC"
                trans = gobjMain.EmpresaActual.GNOpcion.ObtenerValor(KeyRecargo)
                'trans = GetSetting(APPNAME, App.Title, KeyRecargo)
                RecuperaSelecRec KeyRecargo, trans
                
                
        End Select
            
        BandAceptado = False
        Me.Show vbModal, frmMain
        If BandAceptado Then
            .fecha1 = fecha1
            .fecha2 = fecha2
        Select Case cad
            Case "IMPCPI"
                .CodTrans = PreparaCadena
                .CodRetencion1 = PreparaCadenaRet
                .Servicios = PreparaCadReTencion(lstServicios)    'servicios
                gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyTrans, .CodTrans
                gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyTransRet, .CodRetencion1
                gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyRecargo, .Servicios
                'SaveSetting APPNAME, App.Title, KeyTrans, .CodTrans
                'SaveSetting APPNAME, App.Title, KeyTransRet, .CodRetencion1
                'SaveSetting APPNAME, App.Title, KeyRecargo, .Servicios
                gobjMain.EmpresaActual.GNOpcion.GrabarSoloGnOpcion2
            Case "IMPCPI2013"
                .CodTrans = PreparaCadena
                .CodRetencion1 = PreparaCadenaRet
                .Servicios = PreparaCadReTencion(lstServicios)    'servicios
                'SaveSetting APPNAME, App.Title, KeyTrans, .CodTrans
                'SaveSetting APPNAME, App.Title, KeyTransRet, .CodRetencion1
                'SaveSetting APPNAME, App.Title, KeyRecargo, .Servicios
                gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyTrans, .CodTrans
                gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyTransRet, .CodRetencion1
                gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyRecargo, .Servicios
                gobjMain.EmpresaActual.GNOpcion.GrabarSoloGnOpcion2
            Case "IMPFC"
                .CodGrupo = PreparaCadena
                .CodRetencion2 = PreparaCadenaRet
                .Sucursal = PreparaCadReTencion(lstServicios)    'servicios
                'SaveSetting APPNAME, App.Title, KeyTrans, .CodGrupo
                'SaveSetting APPNAME, App.Title, KeyTransRet, .CodRetencion2
                'SaveSetting APPNAME, App.Title, KeyRecargo, .Sucursal
                gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyTrans, .CodGrupo
                gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyTransRet, .CodRetencion2
                gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyRecargo, .Sucursal
                gobjMain.EmpresaActual.GNOpcion.GrabarSoloGnOpcion2
            Case "IMPFCEXP"
                .CodGrupo = PreparaCadena
                .CodRetencion2 = PreparaCadenaRet
                .Sucursal = PreparaCadReTencion(lstServicios)    'servicios
                'SaveSetting APPNAME, App.Title, KeyTrans, .CodGrupo
                'SaveSetting APPNAME, App.Title, KeyTransRet, .CodRetencion2
                'SaveSetting APPNAME, App.Title, KeyRecargo, .Sucursal
                gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyTrans, .CodGrupo
                gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyTransRet, .CodRetencion2
                gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyRecargo, .Sucursal
                gobjMain.EmpresaActual.GNOpcion.GrabarSoloGnOpcion2
            
            Case "IMPFCxE"
                .CodGrupo = PreparaCadena
                .CodRetencion2 = PreparaCadenaRet
                .Sucursal = PreparaCadReTencion(lstServicios)    'servicios
            Case "IMPAN"
                .CodPC2 = PreparaCadena
                gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyAnulado, .CodPC2
                gobjMain.EmpresaActual.GNOpcion.GrabarSoloGnOpcion2
                'SaveSetting APPNAME, App.Title, KeyAnulado, .CodPC2
            Case "IMPDD"
                .CodGrupo = PreparaCadena
                SaveSetting APPNAME, App.Title, KeyTrans, .CodGrupo
            Case "IMPFCICE"
                .CodGrupo = PreparaCadena
                '.CodRetencion2 = PreparaCadenaRet
                .Sucursal = PreparaCadReTencion(lstServicios)    'servicios
                'SaveSetting APPNAME, App.Title, KeyTrans, .CodGrupo
                'SaveSetting APPNAME, App.Title, KeyTransRet, .CodRetencion2
                'SaveSetting APPNAME, App.Title, KeyRecargo, .Sucursal
                gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyTrans, .CodGrupo
                'gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyTransRet, .CodRetencion2
                gobjMain.EmpresaActual.GNOpcion.AsignarValor KeyRecargo, .Sucursal
                gobjMain.EmpresaActual.GNOpcion.GrabarSoloGnOpcion2
                
            End Select
        End If
    End With
    Inicio = BandAceptado
    Unload Me
End Function

Private Sub cmdAceptar_Click()
    Dim Cadena  As String
    If Key = "CarteraxCxV_Trans" Then
        'Verificar si se han seleccionado transacciones
        'en reprote  de Cartera por cobrara por Vendedor
        Cadena = PreparaCadena
        If Cadena = "" Then
            MsgBox "Seleccione las transacciones  a incluir "
            lst.SetFocus
            Exit Sub
        End If
    End If
    
    
    BandAceptado = True
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    BandAceptado = False
    Me.Hide
End Sub


Private Sub dtpDesde_LostFocus()
Dim valor As Long
    
''''    Select Case Month(dtpDesde.value)
''''    Case 1, 3, 5, 7, 8, 10, 12
''''        dtpHasta.value = CDate("31/" & Month(dtpDesde.value) & "/" & Year(dtpDesde.value))
''''    Case 4, 6, 9, 11
''''        dtpHasta.value = CDate("30/" & Month(dtpDesde.value) & "/" & Year(dtpDesde.value))
''''    Case 2
''''        valor = Year(dtpDesde.value) / 4
''''        If Int(valor) * 4 = Year(dtpDesde.value) Then
''''            dtpHasta.value = CDate("29/" & Month(dtpDesde.value) & "/" & Year(dtpDesde.value))
''''        Else
''''            dtpHasta.value = CDate("28/" & Month(dtpDesde.value) & "/" & Year(dtpDesde.value))
''''        End If
''''    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF5
        cmdAceptar_Click
        KeyCode = 0
    Case Else
        MoverCampo Me, KeyCode, Shift, False
    End Select
End Sub

Private Function PreparaCadena() As String
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
    PreparaCadena = Cadena
End Function


Private Function PreparaCadenaRet() As String
    Dim Cadena As String, i As Integer
    Cadena = ""
    For i = 0 To lstRet.ListCount - 1
        If lstRet.Selected(i) Then
            If Cadena = "" Then
                Cadena = Left(lstRet.List(i), lstRet.ItemData(i))
            Else
                Cadena = Cadena & "," & _
                              Left(lstRet.List(i), lstRet.ItemData(i))
            End If
        End If
    Next i
    PreparaCadenaRet = Cadena
End Function






Private Sub Form_Load()
    Dim resto As Integer, valor As Long
    'Establece los rangos de Fecha  siempre  al rango
    'del año actual
''''    dtpDesde.value = CDate("01/" & Month(Date) & "/" & Year(Date))
    
''''    Select Case Month(dtpDesde.value)
''''    Case 1, 3, 5, 7, 8, 10, 12
''''        dtpHasta.value = CDate("31/" & Month(dtpDesde.value) & "/" & Year(dtpDesde.value))
''''    Case 4, 6, 9, 11
''''        dtpHasta.value = CDate("30/" & Month(dtpDesde.value) & "/" & Year(dtpDesde.value))
''''    Case 2
''''        valor = Year(dtpDesde.value)
''''        If Int(valor) * 4 = Year(dtpDesde.value) Then
''''            dtpHasta.value = CDate("29/" & Month(dtpDesde.value) & "/" & Year(dtpDesde.value))
''''        Else
''''            dtpHasta.value = CDate("28/" & Month(dtpDesde.value) & "/" & Year(dtpDesde.value))
''''        End If
''''
''''    End Select
''''    dtpHasta.value = CDate("30/12/" & Year(dtpDesde.value))
End Sub


Private Sub cmdAdd_Click()
    Dim i As Long, ix As Long
    On Error GoTo errtrap
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
errtrap:
    DispErr
End Sub

Private Sub cmdResta_Click()
    Dim i As Long, ix As Long
    On Error GoTo errtrap
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
errtrap:
    DispErr
End Sub

Private Sub lstBienes_DblClick()
    cmdAdd_Click
End Sub

Private Sub lstServicios_DblClick()
    cmdResta_Click
End Sub

Public Sub RecuperaSelecRetencion(ByVal Key As String, ByRef lstBienes As ListBox, ByRef lstServicios As ListBox)
Dim s As String, Vector As Variant, ix As Long
Dim i As Integer, j As Integer, Selec As Integer
    'Recupera seleccionados  del registro de windows
    s = GetSetting(APPNAME, App.Title, Key, "_VACIO_")
    If s <> "_VACIO_" Then
        Vector = Split(s, ",")
         Selec = UBound(Vector, 1)
         For i = 0 To Selec
            For j = lstBienes.ListCount - 1 To 0 Step -1
                If Vector(i) = Left(lstBienes.List(j), lstBienes.ItemData(j)) Then
                    'ix = mobjGrupo.AgregarUsuario(.List(i))
                    ix = lstBienes.ItemData(j)
                    lstServicios.AddItem lstBienes.List(j)
                    lstServicios.ItemData(lstServicios.NewIndex) = ix
                    lstBienes.RemoveItem j
                End If
            Next j
         Next i
    End If
End Sub

Public Function PreparaCadReTencion(lst As ListBox) As String
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
    PreparaCadReTencion = Cadena
End Function


'Public Sub RecuperaSelecRetencion(ByVal Key As String, ByRef lstBienes As ListBox, ByRef lstServicios As ListBox)
'Dim s As String, Vector As Variant, ix As Long
'Dim i As Integer, j As Integer, Selec As Integer
'    'Recupera seleccionados  del registro de windows
'    s = GetSetting(APPNAME, App.Title, Key, "_VACIO_")
'    If s <> "_VACIO_" Then
'        Vector = Split(s, ",")
'         Selec = UBound(Vector, 1)
'         For i = 0 To Selec
'            For j = lstBienes.ListCount - 1 To 0 Step -1
'                If Vector(i) = Left(lstBienes.List(j), lstBienes.ItemData(j)) Then
'                    'ix = mobjGrupo.AgregarUsuario(.List(i))
'                    ix = lstBienes.ItemData(j)
'                    lstServicios.AddItem lstBienes.List(j)
'                    lstServicios.ItemData(lstServicios.NewIndex) = ix
'                    lstBienes.RemoveItem j
'                End If
'            Next j
'         Next i
'    End If
'End Sub
'
Private Sub CargaRecargos()
    Dim rs As Recordset
    lstBienes.Clear
    Set rs = gobjMain.EmpresaActual.ListaIVRecargo(True)
    With rs
        If Not (.EOF) Then
            .MoveFirst
            Do Until .EOF
                lstBienes.AddItem !codRecargo & "  " & !Descripcion
                lstBienes.ItemData(lstBienes.NewIndex) = Len(!codRecargo)
               .MoveNext
           Loop
           
'            lstBienes.AddItem "SUBT" & "  " & "Subtotal"
'            lstBienes.ItemData(lstBienes.NewIndex) = Len("SUBT")
'
        End If
    End With
    rs.Close
End Sub


Public Sub RecuperaSelecRec(ByVal Key As String, trans As String)
Dim s As String, Vector As Variant, ix As Long
Dim i As Integer, j As Integer, Selec As Integer
    'Recupera selecciondados  del registro de windows
    lstServicios.Clear
    s = trans           '  jeaa 20/09/2003
    'GetSetting(APPNAME, App.Title, Key, "_VACIO_")
    If s <> "_VACIO_" Then
        Vector = Split(s, ",")
         Selec = UBound(Vector, 1)
         For i = 0 To Selec
            For j = lstBienes.ListCount - 1 To 0 Step -1
                If Vector(i) = Left(lstBienes.List(j), lstBienes.ItemData(j)) Then
                    'ix = mobjGrupo.AgregarUsuario(.List(i))
                    ix = lstBienes.ItemData(j)
                    lstServicios.AddItem lstBienes.List(j)
                    lstServicios.ItemData(lstServicios.NewIndex) = ix
                    lstBienes.RemoveItem j
                End If
            Next j
         Next i
    End If
End Sub

