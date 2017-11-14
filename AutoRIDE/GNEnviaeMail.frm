VERSION 5.00
Begin VB.Form frmGNEnviaeMail 
   Caption         =   "Envia Correo Electrónico de Comprobante Electrónico"
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtcopiamail 
      Height          =   435
      Left            =   60
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1080
      Width           =   6435
   End
   Begin VB.CheckBox chkReenvio 
      Caption         =   "Reenviar Correo"
      Height          =   195
      Left            =   4920
      TabIndex        =   11
      Top             =   60
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.TextBox txtmailEmpresa 
      Height          =   435
      Left            =   60
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   300
      Width           =   6435
   End
   Begin VB.TextBox txtemclicatalogo 
      Height          =   435
      Left            =   60
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1800
      Width           =   6435
   End
   Begin VB.TextBox txtmail 
      Height          =   675
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "GNEnviaeMail.frx":0000
      Top             =   2580
      Width           =   6435
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Enviar"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "con copia a:"
      Height          =   195
      Left            =   60
      TabIndex        =   13
      Top             =   780
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "e-mail Empresa"
      Height          =   195
      Left            =   60
      TabIndex        =   10
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label lblservidorcorreo 
      Caption         =   "Label3"
      Height          =   315
      Left            =   180
      TabIndex        =   8
      Top             =   4440
      Width           =   5175
   End
   Begin VB.Label lblclavecorreo 
      Caption         =   "Label3"
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   4320
      Width           =   5235
   End
   Begin VB.Label lblmailsaliente 
      Caption         =   "Label3"
      Height          =   315
      Left            =   300
      TabIndex        =   6
      Top             =   3900
      Width           =   5115
   End
   Begin VB.Label lblidcli 
      Caption         =   "Label3"
      Height          =   315
      Left            =   180
      TabIndex        =   5
      Top             =   4740
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Label Label2 
      Caption         =   "e-mail Cliente"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Mensaje"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   2340
      Width           =   1635
   End
End
Attribute VB_Name = "frmGNEnviaeMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private gnc As GNComprobante
Private pc As PCProvCli
Private WithEvents oMail As clsCDOmail
Attribute oMail.VB_VarHelpID = -1

Private Sub cmdEnviar_Click()
    Set oMail = New clsCDOmail
    Dim Nombre As String, MensajeAsunto As String
    With oMail
         'datos para enviar
        .servidor = gnc.Empresa.GNOpcion.ServidorCorreo '465
        .puerto = gnc.Empresa.GNOpcion.PuertoCorreo
        .UseAuntentificacion = True
        .ssl = True
        .usuario = gnc.Empresa.GNOpcion.NombreUsuario
        .PassWord = gnc.Empresa.GNOpcion.PasswordCorreo
        .Copia = gnc.Empresa.GNOpcion.CopiaCorreo
        
        MensajeAsunto = gnc.Empresa.GNOpcion.RazonSocial & " envía la siguiente "
        
        Select Case gnc.GNTrans.AnexoCodTipoComp
                Case 18
                    MensajeAsunto = MensajeAsunto & "Factura Electrónica"
                Case 7
                    MensajeAsunto = MensajeAsunto & "Retención Electrónica"
                Case 6
                    MensajeAsunto = MensajeAsunto & "Guía de Remisión Electrónica"
                Case 5
                    MensajeAsunto = MensajeAsunto & "Nota de Débito Electrónica"
                Case 4
                    MensajeAsunto = MensajeAsunto & "Nota de Crédito Electrónica"
            End Select
        MensajeAsunto = MensajeAsunto & " del cliente: " & pc.Nombre & " ,comprobante: " & gnc.NumSerieEstaSRI & "-" & gnc.NumSeriePuntoSRI & "-" & Right("000000000" & gnc.NumTrans, 9)
        .asunto = MensajeAsunto
'        If chkReenvio.value = vbChecked Then
'            Nombre = gnc.Empresa.GNOpcion.ComprobantesEnviados & "\" & Mid$(gnc.ClaveAcceso, 1, 39) & Right("0000000000" & gnc.transid, 10)
'        Else
            GeneraPDF
            Nombre = gnc.Empresa.GNOpcion.ComprobantesAutorizados & "\" & Mid$(gnc.ClaveAcceso, 1, 39) & Right("0000000000" & gnc.transid, 10)
'        End If
        .Adjunto = Nombre & ".xml" '& ";" & nombre & ".pdf"
        .de = txtmailEmpresa.Text
        .para = pc.EMail '"javabril@hotmail.com" ' txtemclicatalogo.Text   ''
        
        .mensaje = txtmail.Text
        
        .Enviar_Backup ' manda el mail
    
    End With
    
    Set oMail = Nothing
    Unload Me
End Sub

' envio completo
Private Sub oMail_EnvioCompleto()
'    MsgBox "Correo enviado", vbInformation
End Sub
' error al enviar
Private Sub oMail_Error(Descripcion As String, Numero As Variant)
    MsgBox Descripcion, vbCritical, Numero
End Sub

Public Function Inicio(ByVal gn As GNComprobante) As Boolean

    Set gnc = gn
    lblidcli.Caption = gn.IdClienteRef
    Set pc = gn.Empresa.RecuperaPCProvCliQuick(gn.IdClienteRef)
    txtemclicatalogo.Text = pc.EMail
    lblservidorcorreo.Caption = gnc.Empresa.GNOpcion.ServidorCorreo
    lblmailsaliente.Caption = gnc.Empresa.GNOpcion.CuentaCorreo
    lblclavecorreo.Caption = gnc.Empresa.GNOpcion.PasswordCorreo
    txtmail.Text = gnc.Empresa.GNOpcion.MensajeCorreo
    txtmailEmpresa.Text = gnc.Empresa.GNOpcion.CuentaCorreo
    txtcopiamail.Text = gnc.Empresa.GNOpcion.CopiaCorreo
    Me.Show vbModal


End Function


Private Sub GeneraPDF()
    Dim id As Long
    Dim mobjxml As Object
    
        If GeneraRidePDF(gnc, mobjxml) Then
        End If
    
End Sub

Public Sub AutoEnvia(ByVal gc As GNComprobante)
    Set oMail = New clsCDOmail
    Dim Nombre As String, MensajeAsunto As String, pc As PCProvCli
    Set gnc = gc
    With oMail
        If gnc.IdClienteRef <> 0 Then
            Set pc = gnc.Empresa.RecuperaPCProvCliQuick(gnc.IdClienteRef)
        ElseIf gnc.IdProveedorRef <> 0 Then
            Set pc = gnc.Empresa.RecuperaPCProvCliQuick(gnc.IdProveedorRef)
        End If
        
         'datos para enviar
        .servidor = gnc.Empresa.GNOpcion.ServidorCorreo '465
        .puerto = gnc.Empresa.GNOpcion.PuertoCorreo
        .UseAuntentificacion = True
        .ssl = True
        .usuario = gnc.Empresa.GNOpcion.NombreUsuario
        .PassWord = gnc.Empresa.GNOpcion.PasswordCorreo
        .Copia = gnc.Empresa.GNOpcion.CopiaCorreo
        
        MensajeAsunto = gnc.Empresa.GNOpcion.RazonSocial & " envía la siguiente "
        
        Select Case gnc.GNTrans.AnexoCodTipoComp
                Case 18
                    MensajeAsunto = MensajeAsunto & "Factura Electrónica"
                Case 7
                    MensajeAsunto = MensajeAsunto & "Retención Electrónica"
                Case 6
                    MensajeAsunto = MensajeAsunto & "Guía de Remisión Electrónica"
                Case 5
                    MensajeAsunto = MensajeAsunto & "Nota de Débito Electrónica"
                Case 4
                    MensajeAsunto = MensajeAsunto & "Nota de Crédito Electrónica"
            End Select
        MensajeAsunto = MensajeAsunto & " del cliente: " & pc.Nombre & " ,comprobante: " & gnc.NumSerieEstaSRI & "-" & gnc.NumSeriePuntoSRI & "-" & Right("000000000" & gnc.NumTrans, 9)
        .asunto = MensajeAsunto
'        If chkReenvio.value = vbChecked Then
'            nombre = gnc.Empresa.GNOpcion.ComprobantesEnviados & "\" & Mid$(gnc.ClaveAcceso, 1, 39) & Right("0000000000" & gnc.transid, 10)
'        Else
            GeneraPDF
            Nombre = gnc.Empresa.GNOpcion.ComprobantesAutorizados & "\" & Mid$(gnc.ClaveAcceso, 1, 39) & Right("0000000000" & gnc.transid, 10)
'        End If
        .Adjunto = Nombre & ".xml" '& ";" & nombre & ".pdf"
        .de = gnc.Empresa.GNOpcion.CuentaCorreo
        .para = pc.EMail
        
        .mensaje = gnc.Empresa.GNOpcion.MensajeCorreo
        
        .Enviar_Backup ' manda el mail
    
    End With
    
    Set oMail = Nothing
    Unload Me
End Sub

Public Function GeneraRidePDF(ByVal gc As GNComprobante, ByRef objImp As Object) As Boolean

    Dim crear As Boolean
    Dim crearRIDE As Boolean
    On Error GoTo Errtrap

    'Si no tiene TransID quere decir que no está grabada
    If (gc.transid = 0) Or gc.Modificado Then
        MsgBox MSGERR_NOGRABADO, vbInformation
        GeneraRidePDF = False
        Exit Function
    End If
    
  
    
    crearRIDE = (objImp Is Nothing)
    If Not crearRIDE Then crearRIDE = (objImp.NombreDLL <> "GNprintg")
    If crearRIDE Then
        Set objImp = Nothing
        Set objImp = CreateObject("GNprintg.PrintTrans")
    End If
    
    MensajeStatus MSG_PREPARA, vbHourglass
    objImp.GeneraTransRide gobjMain.EmpresaActual, True, 1, 0, "", 0, gc
    GeneraRidePDF = True
    MensajeStatus
    'jeaa 30/09/04
    'gc.CambiaEstadoImpresion

    
    Exit Function
Errtrap:
    GeneraRidePDF = False
    MensajeStatus
    Select Case Err.Number
    Case ERR_NOIMPRIME, ERR_NOIMPRIME2, ERR_NOIMPRIME3, ERR_NOHAYCODIGO
        DispErr
    Case Else
        
        MsgBox MSGERR_NOIMPRIME2, vbInformation
        
    End Select
    GeneraRidePDF = False
    Exit Function
End Function




Public Sub GeneraPDFAuto(ByVal gc As GNComprobante)
    Set oMail = New clsCDOmail
    
    Dim Cifrado As Integer
    Dim Proxy As Integer
    Dim servidorProxy As String
    Dim asunto As String
    Dim strArchivoXML As String
    Dim strArchivoPDF As String
    
    Dim filename
    Dim filenameDestino
    Dim valor As String
    
    Dim Nombre As String, MensajeAsunto As String, pc As PCProvCli, nombredestino As String
    Set gnc = gc
    With oMail
        If gnc.IdClienteRef <> 0 Then
            Set pc = gnc.Empresa.RecuperaPCProvCliQuick(gnc.IdClienteRef)
        ElseIf gnc.IdProveedorRef <> 0 Then
            Set pc = gnc.Empresa.RecuperaPCProvCliQuick(gnc.IdProveedorRef)
        End If
        GeneraPDF
        
        
        Nombre = gc.Empresa.GNOpcion.ComprobantesAutorizados & "\" & Mid$(gc.ClaveAcceso, 1, 39) & Right("0000000000" & gc.transid, 10)
        filename = Dir(Nombre & ".xml", vbArchive)
        strArchivoXML = Nombre & ".xml"
        strArchivoPDF = Nombre & ".pdf"
        

    
    valor = "c:\ia\EnviarCorreoOculto """ & gc.Empresa.GNOpcion.ServidorCorreo & """ " & gc.Empresa.GNOpcion.PuertoCorreo & " """ & gc.Empresa.GNOpcion.CuentaCorreo & _
            """ """ & gc.Empresa.GNOpcion.PasswordCorreo & """ """ & gc.Empresa.GNOpcion.NombreUsuario & """ """ & pc.EMail & _
            """ """ & gc.Empresa.GNOpcion.CopiaCorreo & """ " & Cifrado & " " & Proxy & " " & gc.Empresa.GNOpcion.PuertoCorreo & " """ & servidorProxy & """ """ & gc.Empresa.GNOpcion.NombreEmpresa & _
            """ """ & asunto & """ """ & gc.Empresa.GNOpcion.MensajeCorreo & """ """ & strArchivoXML & """ """ & strArchivoPDF & """ "
            
    Shell valor, vbHide
    nombredestino = gc.Empresa.GNOpcion.ComprobantesEnviados & "\" & Mid$(gc.ClaveAcceso, 1, 39) & Right("0000000000" & gc.transid, 10)
    filenameDestino = Dir(Nombre & ".xml", vbArchive)
            
    valor = " move " & strArchivoXML & " " & nombredestino & ".xml"
    Shell valor, vbHide
        
    valor = " move " & strArchivoPDF & " " & nombredestino & ".pdf"
    Shell valor, vbHide
        
        Set pc = Nothing
    End With
    
    Unload Me
End Sub

