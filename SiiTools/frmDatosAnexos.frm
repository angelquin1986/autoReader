VERSION 5.00
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDatosAnexos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos para Anexos"
   ClientHeight    =   4530
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   6555
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic1 
      BorderStyle     =   0  'None
      Height          =   3915
      Left            =   60
      ScaleHeight     =   3915
      ScaleWidth      =   6375
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   60
      Width           =   6375
      Begin VB.CheckBox chkRetencionOtro 
         Caption         =   "Retención APLICADA en otra Compra con el mismo No. Comprobante"
         Height          =   195
         Left            =   60
         TabIndex        =   14
         Top             =   3600
         Width           =   6135
      End
      Begin VB.CheckBox chkFacturaElec 
         Alignment       =   1  'Right Justify
         Caption         =   "Factura Electrónica"
         Height          =   195
         Left            =   4140
         TabIndex        =   17
         Top             =   480
         Width           =   2115
      End
      Begin VB.TextBox txtNumAutSRI 
         Height          =   372
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1500
         Width           =   4635
      End
      Begin VB.TextBox txtNumSecuencial 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2520
         MaxLength       =   9
         TabIndex        =   7
         Top             =   1860
         Width           =   1035
      End
      Begin VB.TextBox txtCodTransAfectada 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1680
         TabIndex        =   9
         Top             =   2580
         Width           =   612
      End
      Begin VB.CheckBox chkBandDevolucion 
         Caption         =   "La compra tiene derecho a Devolución"
         Height          =   252
         Left            =   60
         TabIndex        =   11
         Top             =   3060
         Width           =   3255
      End
      Begin VB.TextBox txtNumTransAfectada 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2280
         TabIndex        =   10
         Top             =   2580
         Width           =   1275
      End
      Begin VB.TextBox txtnumSeriePunto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   2100
         MaxLength       =   3
         TabIndex        =   6
         Top             =   1860
         Width           =   435
      End
      Begin VB.CheckBox chkNoReteRenta 
         Caption         =   "Compra NO sujeta a retención IR"
         Height          =   255
         Left            =   60
         TabIndex        =   12
         Top             =   3300
         Width           =   2655
      End
      Begin VB.TextBox txtnumSerieEstablecimiento 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1680
         MaxLength       =   3
         TabIndex        =   5
         Top             =   1860
         Width           =   435
      End
      Begin VB.CheckBox chkNOCreditoTributario 
         Caption         =   "NO tiene derecho a Crédito Tributario"
         Height          =   252
         Left            =   60
         TabIndex        =   0
         Top             =   60
         Width           =   3435
      End
      Begin FlexComboProy.FlexCombo fcbTipoRetencion 
         Height          =   375
         Left            =   4440
         TabIndex        =   13
         Top             =   3000
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Enabled         =   0   'False
         ColWidth0       =   500
         ColWidth1       =   3800
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpFechaAnexos 
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   420
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   106692609
         CurrentDate     =   37556
      End
      Begin FlexComboProy.FlexCombo fcbCredTributario 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   780
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   661
         DispCol         =   1
         ColWidth0       =   400
         ColWidth1       =   4000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPFechaCaducidad 
         Height          =   375
         Left            =   4560
         TabIndex        =   8
         Top             =   1860
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   661
         _Version        =   393216
         OLEDropMode     =   1
         Format          =   106692611
         CurrentDate     =   39385
      End
      Begin FlexComboProy.FlexCombo fcbTipoComprobante 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   1140
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   661
         DispCol         =   1
         ColWidth0       =   500
         ColWidth1       =   3800
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin FlexComboProy.FlexCombo fcbFormaPagoSRI 
         Height          =   375
         Left            =   1680
         TabIndex        =   30
         Top             =   2220
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   661
         DispCol         =   1
         ColWidth0       =   500
         ColWidth1       =   3800
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtNumSerie 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   4320
         MaxLength       =   6
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3900
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblTipoRet 
         Caption         =   "Tipo de Retención"
         Height          =   195
         Left            =   4440
         TabIndex        =   20
         Top             =   2760
         Width           =   1515
      End
      Begin VB.Label lblFormaPago 
         Caption         =   "For.de Pago"
         Height          =   195
         Left            =   0
         TabIndex        =   31
         Top             =   2280
         Width           =   1515
      End
      Begin VB.Label lblTransAfectada 
         Caption         =   "NC/ND aplicada a:"
         Height          =   195
         Left            =   0
         TabIndex        =   29
         Top             =   2700
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "# Est-Pun-Secuencial:"
         Height          =   195
         Left            =   0
         TabIndex        =   28
         Top             =   1980
         Width           =   1620
      End
      Begin VB.Label Label4 
         Caption         =   "# Serie Establecimien:"
         Height          =   195
         Left            =   4440
         TabIndex        =   27
         Top             =   3840
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label Label3 
         Caption         =   "# Aut. SRI"
         Height          =   195
         Left            =   0
         TabIndex        =   26
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Comp."
         Height          =   195
         Left            =   0
         TabIndex        =   25
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Sustento Tributario"
         Height          =   195
         Left            =   0
         TabIndex        =   24
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Transaccion"
         Height          =   195
         Left            =   0
         TabIndex        =   23
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "# Serie Punto:"
         Height          =   195
         Left            =   4440
         TabIndex        =   22
         Top             =   4260
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "F.Caducidad"
         Height          =   195
         Left            =   3600
         TabIndex        =   21
         Top             =   1920
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   372
      Left            =   3240
      TabIndex        =   16
      Top             =   4080
      Width           =   972
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar-F9"
      Height          =   372
      Left            =   2160
      TabIndex        =   15
      Top             =   4080
      Width           =   972
   End
End
Attribute VB_Name = "frmDatosAnexos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'El campo TransIDAfectada tiene mismo funcionamiento de PCKardex.IDAsignado, pero se lo ha colocado con el fin de controlar
'que las NC/ND solamente se harán sobre una compra a la vez.

Private mobjGNComp As GNComprobante
Private BandAceptado As Boolean
Private mVisualizando As Boolean
Private TransIDAfectada As Long
Private BandTransIDAfectada As Boolean
Private WithEvents mobjEmpresa As Empresa
Attribute mobjEmpresa.VB_VarHelpID = -1
Private pc As PCProvCli
    
Public Function Inicio(ByVal obj As GNComprobante, Optional ByVal IDTransAfectada As Long, _
                                    Optional ByVal BandTransAfectada As Boolean) As String
    
    Set mobjGNComp = obj
    TransIDAfectada = IDTransAfectada
    BandTransIDAfectada = BandTransAfectada
    DTPFechaCaducidad.Format = dtpCustom
    DTPFechaCaducidad.CustomFormat = "dd/MM/yyyy"
    HabilitaAnexos
        'buscar la forma de identificar si el proveedor original cambio, en el caso de modificación una vez grabado o no
        'Debe el formulario padre, indicarle que se realizó ese cambio?
    Set pc = mobjGNComp.Empresa.RecuperaPCProvCli(mobjGNComp.IdProveedorRef)
    
    
    If (mobjGNComp.TransID = 0 And Not BandAceptado) Or mobjGNComp.NumAutSRI = "-1" Then  'Inicia nuevo y recupera de catálogo de proveedor
        VisualizaNuevo
    Else
        mVisualizando = True
        VisualizaDesdeObjeto
        mVisualizando = False
    End If
    
    Me.Show vbModal
    Inicio = IIf(BandAceptado Or mobjGNComp.SoloVer, "O.K.", "Vacío")
    Unload Me
    Set pc = Nothing
End Function

'Agregado Alex Sept/2002
Private Sub VisualizaNuevo()
    Dim objPC As PCProvCli
    Dim i As Integer, largo As Integer
    Dim dia As Integer, mes As String, fechaAux As Date, anio As String
    'jeaa 06/07/2005
    If mobjGNComp.GNTrans.IVValNumDoc Then
        If Len(mobjGNComp.numDocRef) = 6 Then
            txtNumSerie.Text = Mid$(mobjGNComp.numDocRef, 1, 7)
            txtnumSerieEstablecimiento.Text = ""
            txtnumSeriePunto.Text = ""
        Else
            txtNumSerie.Text = ""
            txtnumSerieEstablecimiento.Text = ""
            txtnumSeriePunto.Text = ""
            
        End If
        If Len(mobjGNComp.numDocRef) = 15 Then
            txtnumSerieEstablecimiento.Text = Mid$(mobjGNComp.numDocRef, 1, 3)
            txtnumSeriePunto.Text = Mid$(mobjGNComp.numDocRef, 4, 3)
            txtNumSecuencial.Text = Mid$(mobjGNComp.numDocRef, 7, 9)
        Else
            txtNumSecuencial.Text = mobjGNComp.numDocRef
            txtnumSerieEstablecimiento.Text = ""
            txtnumSeriePunto.Text = ""
        End If
    Else
            If mobjGNComp.GNTrans.IVVisibleAnexos Then
                txtNumSecuencial.Text = mobjGNComp.numDocRef
                largo = Len(mobjGNComp.numDocRef) + 1
                If largo < 7 Then
                    For i = largo To 7
                        txtNumSecuencial.Text = "0" & txtNumSecuencial.Text
                    Next i
                End If
            Else
                txtNumSecuencial.Text = mobjGNComp.numDocRef
            End If
            txtnumSerieEstablecimiento.Text = ""
            txtnumSeriePunto.Text = ""
    End If
    dtpFechaAnexos.value = mobjGNComp.FechaTrans
    If Len(mobjGNComp.CodProveedorRef) > 0 Then
        Set objPC = mobjGNComp.Empresa.RecuperaPCProvCli(mobjGNComp.CodProveedorRef)
        txtNumAutSRI.Text = objPC.NumAutSRI
        fcbTipoComprobante.KeyText = objPC.TipoComprobante
        
        If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("ActualizaDatosProvCompra")) > 0 Then
            If mobjGNComp.Empresa.GNOpcion.ObtenerValor("ActualizaDatosProvCompra") = "1" Then
                txtnumSerieEstablecimiento.Text = objPC.NumSerie
                txtnumSeriePunto.Text = objPC.NumPunto
                DTPFechaCaducidad.value = objPC.FechaCaducidad
            End If
        End If
        
        If Len(mobjGNComp.GNTrans.AnexoSustento) > 0 Then
            fcbCredTributario.KeyText = mobjGNComp.GNTrans.AnexoSustento
            fcbCredTributario_Selected mobjGNComp.GNTrans.AnexoSustento, mobjGNComp.GNTrans.AnexoSustento
        End If
        
    End If
    
    
    
    mes = DatePart("m", DateAdd("m", -1, Date))
    anio = DatePart("yyyy", Date)
    Select Case mes
     Case 1, 3, 5, 7, 8, 10, 12
        fechaAux = "31/" & mes & "/" & anio
    Case 4, 6, 9, 11
        fechaAux = "30/" & mes & "/" & anio
    Case 2
        fechaAux = "28/" & mes & "/" & anio
    End Select
        If Len(mobjGNComp.Empresa.GNOpcion.ObtenerValor("ActualizaDatosProvCompra")) > 0 Then
        If mobjGNComp.Empresa.GNOpcion.ObtenerValor("ActualizaDatosProvCompra") = "0" Then
           DTPFechaCaducidad.value = fechaAux
        End If
    End If
    
    
    
    If chkNoReteRenta.value = vbChecked Then
        lblTipoRet.Enabled = True
        fcbTipoRetencion.Enabled = True
        If mobjGNComp.EsNuevo Then
        If pc.TipoProvCli = "RISE" Or pc.TipoProvCli = "ARTE" Then
            fcbTipoRetencion.KeyText = "332"
        Else
            fcbTipoRetencion.KeyText = mobjGNComp.GNTrans.AnexoCodTipoRetencion
        End If
    Else
        fcbTipoRetencion.KeyText = mobjGNComp.CodTipoRetencion
    End If

    Else
        lblTipoRet.Enabled = False
        fcbTipoRetencion.Enabled = False
    End If
    
End Sub

Private Sub HabilitaAnexos()
    Dim obj As GNComprobante

    'Deshabilita campos
    If mobjGNComp.SoloVer Then
        fcbTipoComprobante.Enabled = False
        fcbCredTributario.Enabled = False
        txtNumAutSRI.Enabled = False
        txtNumSerie.Enabled = False
        txtNumSecuencial.Enabled = False
        cmdAceptar.Enabled = False
        txtnumSerieEstablecimiento.Enabled = False
        txtnumSeriePunto.Enabled = False
        dtpFechaAnexos.Enabled = False
        DTPFechaCaducidad.Enabled = False
        txtCodTransAfectada.Enabled = False
        txtNumTransAfectada.Enabled = False
        chkRetencionOtro.Enabled = False
        chkFacturaElec.Enabled = False
        
    Else
        lblTransAfectada.Enabled = BandTransIDAfectada
        txtCodTransAfectada.Enabled = BandTransIDAfectada
        txtNumTransAfectada.Enabled = BandTransIDAfectada
        chkRetencionOtro.Enabled = Not (BandTransIDAfectada)
        If BandTransIDAfectada Then
            If TransIDAfectada <> 0 Then
                Set obj = mobjGNComp.Empresa.RecuperaGNComprobante(TransIDAfectada)
                txtCodTransAfectada.Text = obj.CodTrans
                txtNumTransAfectada.Text = obj.numtrans
                Set obj = Nothing
            End If
        End If
        dtpFechaAnexos.Enabled = True 'Not (BandTransIDAfectada)
        DTPFechaCaducidad.Enabled = True ' Not (BandTransIDAfectada)

    End If
    
End Sub

Private Sub VisualizaDesdeObjeto()
    Dim obj As GNComprobante
    'Recupera datos de GnComprobante
    Dim CodSecuenci As String

    Select Case pc.codtipoDocumento
    Case "R": CodSecuenci = "01"
    Case "C": CodSecuenci = "02"
    Case "P": CodSecuenci = "03"
'    Case "F": CodSecuenci = "07"
    Case Else: CodSecuenci = ""
    End Select
    If Len(CodSecuenci) > 0 Then
        fcbTipoComprobante.SetData mobjGNComp.Empresa.ListaAnexoTipoComprobanteValidado(True, False, fcbCredTributario.KeyText, CodSecuenci)
'    Else
'        MsgBox "El Proveedor no tiene asignado el Tipo de Documento" & Chr(13) & "Debe primero arreglar el Proveedor "
    End If

    
    dtpFechaAnexos.value = IIf(mobjGNComp.FechaAnexos = 0, mobjGNComp.FechaTrans, mobjGNComp.FechaAnexos)
    fcbCredTributario.KeyText = mobjGNComp.CodCredTrib
    fcbTipoComprobante.KeyText = mobjGNComp.CodTipoComp
    
    chkFacturaElec.value = IIf(mobjGNComp.BandFactElec, vbChecked, vbUnchecked)
    If mobjGNComp.BandFactElec Then
        txtNumAutSRI.MaxLength = 37
    Else
        txtNumAutSRI.MaxLength = 10
    End If
    
    
    txtNumAutSRI.Text = mobjGNComp.NumAutSRI
    txtNumSerie.Text = mobjGNComp.NumSerie
    txtNumSecuencial.Text = mobjGNComp.NumSecuencial

    'txtNumSecuencial.Text = mobjGNComp.NumSecuencial
    'jeaa 06/07/2005
'        If mobjGNComp.GNTrans.IVValNumDoc Then
'            If Len(mobjGNComp.numDocRef) = 15 Then
'                txtNumSerie.Text = Mid$(mobjGNComp.numDocRef, 1, 6)
'                txtnumSerieEstablecimiento.Text = Mid$(mobjGNComp.numDocRef, 1, 3)
'                txtnumSeriePunto.Text = Mid$(mobjGNComp.numDocRef, 4, 3)
'                txtNumSecuencial.Text = Mid$(mobjGNComp.numDocRef, 7, 9)
'            Else
'                txtNumSecuencial.Text = mobjGNComp.numDocRef
'                txtnumSerieEstablecimiento.Text = ""
'                txtnumSeriePunto.Text = ""
'            End If
'        Else
'            txtNumSecuencial.Text = mobjGNComp.numDocRef
'            txtnumSerieEstablecimiento.Text = ""
'            txtnumSeriePunto.Text = ""
'        End If
    
    If BandTransIDAfectada Then
        Set obj = mobjGNComp.Empresa.RecuperaGNComprobante(mobjGNComp.TransIDAfectada)
        If Not (obj Is Nothing) Then
            txtCodTransAfectada.Text = obj.CodTrans
            txtNumTransAfectada.Text = obj.numtrans
        End If
        Set obj = Nothing
    End If
    DTPFechaCaducidad.value = IIf(mobjGNComp.FechaCaducidad = 0, mobjGNComp.FechaTrans, mobjGNComp.FechaCaducidad)
    txtnumSerieEstablecimiento.Text = mobjGNComp.NumSerieEstablecimiento
    txtnumSeriePunto.Text = mobjGNComp.NumSeriePunto
    If mobjGNComp.BandCompraSinRetencion Then
        chkNoReteRenta.value = vbChecked
        fcbTipoRetencion.KeyText = mobjGNComp.CodTipoRetencion
    Else
        chkNoReteRenta.value = vbUnchecked
        fcbTipoRetencion.KeyText = ""
    
    End If
    chkFacturaElec.value = IIf(mobjGNComp.BandFactElec, vbChecked, vbUnchecked)
    chkRetencionOtro.value = IIf(mobjGNComp.BandRetOtro, vbChecked, vbUnchecked)
    fcbFormaPagoSRI.KeyText = mobjGNComp.CodFormaPagoSRI
End Sub

Private Function VerificaDatos() As Boolean
Dim obj As GNComprobante
        If dtpFechaAnexos.value > mobjGNComp.FechaTrans Then
            MsgBox "La fecha del documento de compra no puede ser mayor que la" & vbCrLf & _
                            "fecha de la transacción"
            dtpFechaAnexos.SetFocus
            Exit Function
        End If
        If Len(fcbCredTributario.KeyText) = 0 Then
            MsgBox "Ingrese crédito tributario", vbInformation
            fcbCredTributario.SetFocus
            Exit Function
        End If
        If Len(fcbTipoComprobante.KeyText) = 0 Then
            MsgBox "Ingrese tipo de comprobante", vbInformation
            fcbTipoComprobante.SetFocus
            Exit Function
        End If
        If Len(txtNumAutSRI.Text) = 0 Then
            MsgBox "Ingrese número Aut. SRI", vbInformation
            txtNumAutSRI.SetFocus
            Exit Function
        End If
        
        If Len(txtnumSerieEstablecimiento.Text) = 0 Then
            MsgBox "Ingrese número de serie Establecimiento del comprobante", vbInformation
            txtnumSerieEstablecimiento.SetFocus
            Exit Function
        End If
        If Len(txtnumSerieEstablecimiento.Text) <> 3 Then
            MsgBox "Número de serie Establecimiento del comprobante incorrecto" & Chr(13) & "Debe tener 3 caracteres", vbInformation
            txtnumSerieEstablecimiento.SetFocus
            Exit Function
        End If
        
        If Len(txtnumSeriePunto.Text) = 0 Then
            MsgBox "Ingrese número de serie Punto del comprobante", vbInformation
            txtnumSeriePunto.SetFocus
            Exit Function
        End If
        If Len(txtnumSeriePunto.Text) <> 3 Then
            MsgBox "Número de serie Punto del comprobante incorrecto" & Chr(13) & "Debe tener 3 caracteres", vbInformation
            txtnumSeriePunto.SetFocus
            Exit Function
        End If
        
        If DTPFechaCaducidad.value < mobjGNComp.FechaTrans Then
            MsgBox "La fecha de Caducidad del documento de compra no puede ser menor que la" & vbCrLf & _
                            "fecha de la transacción"
            DTPFechaCaducidad.SetFocus
            Exit Function
        End If
        
        'jeaa 17/09/2007
        If DateDiff("m", mobjGNComp.FechaTrans, DTPFechaCaducidad.value) > 12 Then
            MsgBox "La fecha de Caducidad del documento de compra no puede mayor a un año que la" & vbCrLf & _
                            "fecha de la transacción"
            DTPFechaCaducidad.SetFocus
            Exit Function
        End If
        
        
        If Len(txtNumSecuencial.Text) = 0 Then
            MsgBox "Ingrese número secuencial del comprobante", vbInformation
            txtNumSecuencial.SetFocus
            Exit Function
        End If
        'jeaa 06/05/2005
        If Len(txtNumSecuencial.Text) <> 9 Then
            MsgBox "Número secuencial del comprobante Incorrecto" & Chr(13) & "Debe tener 9 caracteres", vbInformation
            txtNumSecuencial.SetFocus
            Exit Function
        End If
        If txtCodTransAfectada.Enabled Then
            If txtCodTransAfectada.Text = "" Then
                MsgBox "Ingrese código de transacción afectada", vbInformation
                txtCodTransAfectada.SetFocus
                Exit Function
            End If
            If txtNumTransAfectada.Text = "" Then
                MsgBox "Ingrese número de transacción afectada", vbInformation
                txtNumTransAfectada.SetFocus
                Exit Function
            End If
            'verifica la validez de la transacción especificada como afectada
            Set obj = mobjGNComp.Empresa.RecuperaGNComprobante(0, txtCodTransAfectada.Text, Val(txtNumTransAfectada.Text))
            If obj Is Nothing Then
                MsgBox "La transacción especificada como afectada" & vbCrLf & _
                            "no existe, por favor vuelva a ingresar", vbInformation
                txtCodTransAfectada.SetFocus
                Exit Function
            End If
            'si existe, asignar a objeto
            mobjGNComp.TransIDAfectada = obj.TransID
            Set obj = Nothing
        End If

        
        
        VerificaDatos = True
End Function

Private Sub AsignaDatosAObjeto()
            
        'Asigna los valores correspondientes al objeto
        With mobjGNComp
            .NumAutSRI = txtNumAutSRI.Text
            If Len(txtNumSerie.Text) = 0 Then
                .NumSerie = txtnumSerieEstablecimiento.Text & txtnumSeriePunto.Text
            Else
                .NumSerie = txtNumSerie.Text
            End If
            .NumSerieEstablecimiento = txtnumSerieEstablecimiento.Text
            .NumSeriePunto = txtnumSeriePunto.Text
            .NumSecuencial = txtNumSecuencial.Text
            If mobjGNComp.GNTrans.IVValNumDoc Then
                .numDocRef = txtnumSerieEstablecimiento.Text & txtnumSeriePunto.Text & txtNumSecuencial.Text
            Else
                .numDocRef = txtNumSecuencial.Text
            End If
            .CodCredTrib = fcbCredTributario.KeyText
            .CodTipoComp = fcbTipoComprobante.KeyText

            .FechaAnexos = dtpFechaAnexos.value
            .FechaCaducidad = DTPFechaCaducidad.value
            .CodTipoTrans = mobjGNComp.GNTrans.AnexoCodTipoTrans
            .BandCompraSinRetencion = IIf(chkNoReteRenta.value, vbChecked, vbUnchecked)
            .CodTipoRetencion = fcbTipoRetencion.KeyText
            .BandFactElec = IIf(chkFacturaElec.value = vbChecked, True, False)
            .BandRetOtro = IIf(chkRetencionOtro.value = vbChecked, True, False)
            .CodFormaPagoSRI = fcbFormaPagoSRI.KeyText
            
        End With
End Sub



Private Sub chkFacturaElec_Click()
    If chkFacturaElec.value = vbChecked Then
        txtNumAutSRI.MaxLength = 37
    Else
        txtNumAutSRI.MaxLength = 10
        txtNumAutSRI.Text = ""
        
    End If
End Sub

Private Sub cmdAceptar_Click()
    If Not VerificaDatos Then Exit Sub
    AsignaDatosAObjeto
    BandAceptado = True
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    If (mobjGNComp.NumAutSRI = "" Or mobjGNComp.NumAutSRI = "-1") And BandAceptado Then BandAceptado = False
    Me.Hide
End Sub

Private Sub DTPFechaCaducidad_Change()
'    Dim dia As Integer, mes As String, fechaAux As Date, anio As String
'    dia = DatePart("dd", DTPFechaCaducidad.value)
'    mes = DatePart("m", DTPFechaCaducidad.value)
'    anio = DatePart("yyyy", DTPFechaCaducidad.value)
'    Select Case mes
'     Case 1, 3, 5, 7, 8, 10, 12
'        fechaAux = "31/" & mes & "/" & anio
'    Case 4, 6, 9, 11
'        fechaAux = "30/" & mes & "/" & anio
'    Case 2
'        fechaAux = "28/" & mes & "/" & anio
'    End Select
'    DTPFechaCaducidad.value = fechaAux
End Sub




Private Sub fcbCredTributario_Selected(ByVal Text As String, ByVal KeyText As String)
    Dim CodSecuenci As String
    On Error GoTo ErrTrap
    mobjGNComp.CodCredTrib = KeyText
    Select Case pc.codtipoDocumento
        Case "R": CodSecuenci = "01"
        Case "C": CodSecuenci = "02"
        Case "P": CodSecuenci = "03"
        Case "F": CodSecuenci = "07"
        Case Else: CodSecuenci = ""
    End Select
    
    If Len(CodSecuenci) > 0 Then
            fcbTipoComprobante.SetData mobjGNComp.Empresa.ListaAnexoTipoComprobanteValidado(True, False, fcbCredTributario.KeyText, CodSecuenci)
    Else
            MsgBox "El Proveedor no tiene asignado el Tipo de Documento" & Chr(13) & "Debe primero arreglar el Proveedor "
    End If
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub fcbTipoComprobante_Selected(ByVal Text As String, ByVal KeyText As String)
    On Error GoTo ErrTrap
'    mobjGNComp.CodTipoComp = KeyText
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub Form_Initialize()
    BandAceptado = False
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
Private Sub Form_Load()
    fcbCredTributario.SetData mobjGNComp.Empresa.ListaAnexoTipoSustento(True, False)
    fcbTipoRetencion.SetData mobjGNComp.Empresa.ListaAnexoTipoRetencionPorcentaje0(True, False)
    fcbFormaPagoSRI.SetData mobjGNComp.Empresa.ListaAnexoFormaPago(True, False)
End Sub


Private Sub txtNumAutSRI_KeyPress(KeyAscii As Integer)
    If mVisualizando Then Exit Sub
    'Acepta solo numericos
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
       (KeyAscii <> vbKeyBack) And _
       (KeyAscii <> Asc(".")) And _
       (KeyAscii <> vbKeyReturn) And _
       (KeyAscii <> 22) Then               '22 = CTRL+v (CTRL+c es automático)
        KeyAscii = 0
    End If
End Sub


Private Sub txtNumSecuencial_KeyPress(KeyAscii As Integer)
    If mVisualizando Then Exit Sub
    'Acepta solo numericos
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
       (KeyAscii <> vbKeyBack) And _
       (KeyAscii <> Asc(".")) And _
       (KeyAscii <> vbKeyReturn) And _
       (KeyAscii <> 22) Then               '22 = CTRL+v (CTRL+c es automático)
        KeyAscii = 0
    End If
End Sub


Private Sub txtNumSecuencial_Validate(Cancel As Boolean)
    If Len(txtNumSecuencial.Text) <> 9 Then
        MsgBox "Número de secuencia Incorrecto"
        Cancel = False
    End If
End Sub

Private Sub txtNumSerie_KeyPress(KeyAscii As Integer)
    If mVisualizando Then Exit Sub
    'Acepta solo numericos
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
       (KeyAscii <> vbKeyBack) And _
       (KeyAscii <> Asc(".")) And _
       (KeyAscii <> vbKeyReturn) And _
       (KeyAscii <> 22) Then               '22 = CTRL+v (CTRL+c es automático)
        KeyAscii = 0
    End If
End Sub

Private Sub txtNumSerieestablecimiento_KeyPress(KeyAscii As Integer)
    If mVisualizando Then Exit Sub
    'Acepta solo numericos
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
       (KeyAscii <> vbKeyBack) And _
       (KeyAscii <> Asc(".")) And _
       (KeyAscii <> vbKeyReturn) And _
       (KeyAscii <> 22) Then               '22 = CTRL+v (CTRL+c es automático)
        KeyAscii = 0
    End If
End Sub


Private Sub txtNumSeriePunto_KeyPress(KeyAscii As Integer)
    If mVisualizando Then Exit Sub
    'Acepta solo numericos
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
       (KeyAscii <> vbKeyBack) And _
       (KeyAscii <> Asc(".")) And _
       (KeyAscii <> vbKeyReturn) And _
       (KeyAscii <> 22) Then               '22 = CTRL+v (CTRL+c es automático)
        KeyAscii = 0
    End If
End Sub

Private Sub txtNumSeriePunto_LostFocus()
    txtnumSeriePunto.Text = Format(Val(txtnumSeriePunto.Text), "000")
End Sub

Private Sub txtNumSerieEstablecimiento_LostFocus()
    txtnumSerieEstablecimiento.Text = Format(Val(txtnumSerieEstablecimiento.Text), "000")
End Sub

Private Sub txtNumAutsri_LostFocus()
'    txtNumAutSRI.Text = Format(Val(txtNumAutSRI.Text), "0000000000")
End Sub

Private Sub txtNumsecuencial_LostFocus()
    txtNumSecuencial.Text = Format(Val(txtNumSecuencial.Text), "000000000")
End Sub

Private Sub chkNoReteRenta_Click()
    If chkNoReteRenta.value = vbChecked Then
        lblTipoRet.Enabled = True
        fcbTipoRetencion.Enabled = True
        If mobjGNComp.EsNuevo Then
            fcbTipoRetencion.KeyText = mobjGNComp.GNTrans.AnexoCodTipoRetencion
        Else
            fcbTipoRetencion.KeyText = mobjGNComp.CodTipoRetencion
        End If

    Else
        lblTipoRet.Enabled = False
        fcbTipoRetencion.Enabled = False
        fcbTipoRetencion.KeyText = ""
    End If
End Sub

