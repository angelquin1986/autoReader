VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Crea RIDE pdf"
   ClientHeight    =   4470
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10350
   Icon            =   "frmMaint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   10350
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerSeg 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4740
      Top             =   60
   End
   Begin VB.TextBox txtClave 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   3465
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   5070
      Width           =   1488
   End
   Begin VB.TextBox txtNombre 
      Height          =   360
      Left            =   1020
      TabIndex        =   7
      Top             =   5040
      Width           =   1488
   End
   Begin VB.TextBox txtTiempoEspera 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Text            =   "1"
      Top             =   120
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   8460
      Top             =   3600
   End
   Begin TabDlg.SSTab sst1 
      Height          =   315
      Left            =   10980
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   529
      TabCaption(0)   =   "Transacciones"
      TabPicture(0)   =   "frmMaint.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grdTrans"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Catálogos"
      TabPicture(1)   =   "frmMaint.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdCat"
      Tab(1).ControlCount=   1
      Begin VSFlex7LCtl.VSFlexGrid grdTrans 
         Height          =   3195
         Left            =   120
         TabIndex        =   1
         Top             =   420
         Width           =   6500
         _cx             =   11465
         _cy             =   5636
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7LCtl.VSFlexGrid grdCat 
         Height          =   3195
         Left            =   -74880
         TabIndex        =   2
         Top             =   420
         Width           =   6500
         _cx             =   11465
         _cy             =   5636
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   3
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grd 
      Height          =   3075
      Left            =   180
      TabIndex        =   14
      Top             =   540
      Width           =   9915
      _cx             =   17489
      _cy             =   5424
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
      FocusRect       =   4
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
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
      AutoSearch      =   2
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   0
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
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
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label lblseg 
      Alignment       =   1  'Right Justify
      Caption         =   "41"
      Height          =   315
      Left            =   3780
      TabIndex        =   13
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblRutraDestino 
      Caption         =   "Label3"
      Height          =   255
      Left            =   60
      TabIndex        =   12
      Top             =   4140
      Width           =   6015
   End
   Begin VB.Label lblRutraOrigen 
      Caption         =   "Label3"
      Height          =   255
      Left            =   60
      TabIndex        =   11
      Top             =   3840
      Width           =   6015
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&Clave    "
      Height          =   195
      Index           =   1
      Left            =   2625
      TabIndex        =   10
      Tag             =   "&Contraseña:"
      Top             =   5160
      Width           =   570
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Usuari&o    "
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   9
      Tag             =   "Usuari&o:"
      Top             =   5100
      Width           =   705
   End
   Begin VB.Label Label2 
      Caption         =   "minutos"
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   180
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Tiempo de Espera"
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   180
      Width           =   1515
   End
   Begin VB.Label labelhora 
      Caption         =   "5"
      Height          =   255
      Left            =   5880
      TabIndex        =   4
      Top             =   120
      Width           =   195
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const APPNAME = "Ishida y Asociados"
Const SECTION = "Sii"
Const KEY_PARA_CIFRAR = "chRIstIANQuIzhpEQuIzHpEa1161970UPeBzJtGjmf;"
'"QuIzHpEaMabl0050869TOdAyIsFile;"

Public MaxRegistro As String

Public Usuario_RIDEpdf As String
Public Password_Usuario_RIDEpdf As String
Public TiempoEspera_RIDEpdf As String
Private rs As Recordset
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private FilaComp As Integer
Private gnc As GNComprobante
Private MaxFila As Integer

Private Const COL_NUMFILA = 0
Private Const COL_CODTRANS = 1
Private Const COL_FECHA = 2
Private Const COL_TID = 3
Private Const COL_PDF = 4
Private Const COL_NOMBRE = 5
Private Const COL_TIPO = 6


Public Sub Inicio()
    Me.Caption = "Crea RIDE pdf" & gobjMain.EmpresaActual.CodEmpresa
    FilaComp = 1
    
    Me.Show
    Me.ZOrder
End Sub


Public Sub AutoInicio()
    Dim i As Integer
    
'    lblseg.Caption = "41"
    lblseg.Caption = "21"
    
    
    TimerSeg.Enabled = True
    
    Set rs = Nothing
    Set rs = gobjMain.EmpresaActual.CargaTransaccionesParaGenerarRide

    
    lblRutraOrigen.Caption = gobjMain.EmpresaActual.GNOpcion.ComprobantesAutorizados
    lblRutraDestino.Caption = gobjMain.EmpresaActual.GNOpcion.ComprobantesEnviados
    
    If rs.RecordCount > 0 Then
        MaxFila = rs.RecordCount
        Set grd.DataSource = rs
        
        grd.FormatString = "#|<Trans|<Fecha Trans|>g.transid|<pdf|<nombre|>Tipo"
        grd.ColWidth(1) = 1500
        grd.ColWidth(2) = 1500
        grd.ColWidth(COL_TID) = 0
        grd.ColWidth(COL_PDF) = 0
        grd.ColWidth(COL_NOMBRE) = 6000
        grd.ColWidth(COL_TIPO) = 0
        
        FilaComp = 1
        TimerSeg.Enabled = True
        
    Else
        TimerSeg.Enabled = False
        Timer1.Enabled = True
        Set rs = Nothing
        If Timer1.Enabled = False Then
            AutoInicio
        End If
    End If
    Set rs = Nothing

End Sub

Private Sub Envio()
    Dim comando As String
    Dim origen As String, destino As String
    Dim transid As Long
    Dim i As Integer
    
        If grd.ValueMatrix(i, COL_PDF) = 0 Then
            If gobjMain.EmpresaActual.ActualizaBanderaGeneraRide(grd.ValueMatrix(i, COL_TID)) Then
                grd.TextMatrix(i, COL_PDF) = 1
                grd.Refresh
            End If
        End If
    
End Sub

Private Sub Form_Activate()
    Me.Caption = "Crea RIDE pdf [" & gobjMain.EmpresaActual.CodEmpresa & "]"
    AutoInicio
End Sub

Private Sub Timer1_Timer()
    
    If txtTiempoEspera.Text <> "" And labelhora.Caption > "0" Then
        labelhora.Caption = Val(labelhora.Caption) - 1
    End If
    If labelhora.Caption = "0" Then
        Timer1.Enabled = False
        labelhora.Caption = txtTiempoEspera.Text
        AutoInicio
    End If
End Sub

Private Sub Recuperar()
    Dim i As Integer, s As String, objCifrar As Sii4Seg.clsCifrar
    On Error Resume Next
    
    
'    Usuario_IshidaMovil = GetSetting(APPNAME, SECTION, "Usuario_RIDEpdf", "")
    
    'Clave de servidor está cifrado             '*** MAKOTO 19/mar/01
    s = GetSetting(APPNAME, SECTION, "Password_Usuario_RIDEpdf", "")
    Set objCifrar = New Sii4Seg.clsCifrar
    If Not (objCifrar Is Nothing) Then
        s = objCifrar.Decifrar(s, KEY_PARA_CIFRAR)
    End If
'    Password_Usuario_IshidaMovil = s
'    MaxRegistro_IshidaMovil = GetSetting(APPNAME, SECTION, "TiempoEspera_RIDEpdf", "5")
    'Recupera

    
End Sub

Private Sub Recupera()
    Dim i As Integer, s As String, objCifrar As Sii4Seg.clsCifrar
    On Error Resume Next
    
'    txtNombre.Text = Usuario_IshidaMovil
 '   txtClave.Text = Password_Usuario_IshidaMovil
  '  txtNumReg.Text = MaxRegistro_IshidaMovil
    
    
End Sub


Public Sub GeneraPDFAuto()
'    Set oMail = New clsCDOmail

    Dim Cifrado As Integer
    Dim Proxy As Integer
    Dim servidorProxy As String
    Dim asunto As String
    Dim strArchivoXML As String
    Dim strArchivoPDF As String
    Dim mobjxml As Object
    Dim filename
    Dim valor As String
    Dim Nombre As String, MensajeAsunto As String, pc As PCProvCli, nombredestino  As String, email As String
    On Error GoTo Errtrap
    

    GeneraRidePDF mobjxml
        
    
        
        Set pc = Nothing

        

    Exit Sub
Errtrap:
    If Err.Number <> 32755 Then
        MsgBox Err.Description
    End If
    Exit Sub
        
    'End With
    
'    Unload Me
End Sub


Public Function GeneraRidePDF(ByRef objImp As Object) As Boolean
    Dim crear As Boolean
    Dim crearRIDE As Boolean

    
    On Error GoTo Errtrap

    If grd.ValueMatrix(FilaComp, COL_PDF) = 0 Then


      'Si no tiene TransID quere decir que no está grabada
      If (gnc.transid = 0) Or gnc.Modificado Then
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
      objImp.GeneraTransRide gobjMain.EmpresaActual, True, 1, 0, "", 0, gnc
'      If gnc.Empresa.ActualizaBanderaGeneraRide(gnc.transid) Then
'      End If
      GeneraRidePDF = True
      MensajeStatus
      grd.TextMatrix(FilaComp, COL_PDF) = "1"
      'jeaa 30/09/04
      'gc.CambiaEstadoImpresion
      Set objImp = Nothing
    End If
    
    
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

Public Function MoverArchivos()
    Dim ret As String, hPID As Variant
    Dim nombreArchivo As String
    Dim Nombre As String, nombredestino  As String, valor As String, valorpdf As String
    Dim strArchivoXML As String
    Dim strArchivoPDF As String

    

    nombreArchivo = grd.TextMatrix(FilaComp, COL_NOMBRE)
    Nombre = lblRutraOrigen.Caption & nombreArchivo '  gc.Empresa.GNOpcion.ComprobantesAutorizados & "\" & Mid$(gc.ClaveAcceso, 1, 39) & Right("0000000000" & gc.transid, 10)
    strArchivoXML = Nombre & ".xml"
    strArchivoPDF = Nombre & ".pdf"

    nombredestino = lblRutraDestino & nombreArchivo 'gc.Empresa.GNOpcion.ComprobantesEnviados & "\" & Mid$(gc.ClaveAcceso, 1, 39) & Right("0000000000" & gc.transid, 10)
            
'    valor = "  move """ & strArchivoXML & """ """ & nombredestino & ".xml"""
'    valorpdf = " move """ & strArchivoPDF & """ """ & nombredestino & ".pdf"""
    
    MoveFile Nombre & ".xml", nombredestino & ".xml"
    MoveFile Nombre & ".PDF", nombredestino & ".PDF"
           
End Function





Public Sub EmailAuto()
'    Set oMail = New clsCDOmail

    Dim Cifrado As Integer
    Dim Proxy As Integer
    Dim servidorProxy As String
    Dim asunto As String
    Dim strArchivoXML As String
    Dim strArchivoPDF As String
    Dim mobjxml As Object
    Dim filename
    Dim valor As String
    Dim Nombre As String, MensajeAsunto As String, pc As PCProvCli, nombredestino  As String, email As String
    
    Dim texto As String, v As Variant, trans As GNTrans, cadena  As String, w As Variant, X As Variant, ptotal As Currency
    
    On Error GoTo Errtrap
    
    
        
    Cifrado = gnc.Empresa.GNOpcion.BandConexionSegura
    Proxy = 0
    servidorProxy = "servidorproxy"
    asunto = "Envió de Comprobante Electrónico." 'ENLAZAR A LA TABLA GNOPCION

    'With oMail
        If gnc.IdClienteRef <> 0 Then
            Set pc = gnc.Empresa.RecuperaPCProvCliQuick(gnc.IdClienteRef)
        ElseIf gnc.IdProveedorRef <> 0 Then
            Set pc = gnc.Empresa.RecuperaPCProvCliQuick(gnc.IdProveedorRef)
        End If
        
        email = pc.email
    '    email = "javabril@hotmail.com"
        
    Set trans = gnc.Empresa.RecuperaGNTrans(gnc.CodTrans)   ' ************** JEAA 20-8-03
    If trans.TipoTrans = "1" Then
        cadena = "FACTURA"
    ElseIf trans.TipoTrans = "4" Then
        cadena = "NOTA DE CREDITO "
    ElseIf trans.TipoTrans = "5" Then
        cadena = "NOTA DE DEBITO "
    ElseIf trans.TipoTrans = "6" Then
        cadena = "GUIA DE REMISION"
    ElseIf trans.TipoTrans = "7" Then
        cadena = "COMPROBANTE DE RETENCION"
    
    End If


    If trans.TipoTrans <> "7" Then
        ptotal = Format(Abs(gnc.IVKardexPTotal(True)) + gnc.IVRecargoTotal(True, False), "#,0.00")
    Else
        ptotal = Format(Abs(TotalRetencion(gnc)), "#,0.00")
    End If

    If InStr(1, gnc.Empresa.GNOpcion.MensajeCorreo, "<p>") > 0 Then
        texto = "<i><b>Estimado(a),</b></i><br/><p>" & gnc.Nombre & "<p> "
        texto = texto & "<p><b>" & gnc.Empresa.GNOpcion.RazonSocial & "</b> "
        If InStr(1, gnc.Empresa.GNOpcion.MensajeCorreo, "[DETALLETRANS]") > 0 Then
            v = Split(gnc.Empresa.GNOpcion.MensajeCorreo, "[DETALLETRANS]")
            texto = texto & v(0) & "</p>"
            texto = texto & "<p><b>Comprobante  : </b>" & cadena & "</p>"
            texto = texto & "<p><b>Número       : </b>" & gnc.NumSerieEstaSRI & "-" & gnc.NumSerieEstaSRI & "-" & Right("000000000" & gnc.NumTrans, 9) & "</p>"
            texto = texto & "<p><b>Fecha Emision: </b>" & gnc.FechaTrans & "</p>"
            If trans.TipoTrans <> "6" Then
                texto = texto & "<p><b>Total        : </b>" & ptotal & "</p>"
            End If
            
            If InStr(1, v(1), "[LOGO]") > 0 Then
                w = Split(v(1), "[LOGO]")
                texto = texto & w(0) & "</p>"
                X = Split(w(1), ",")
                texto = texto & "<img src='" & X(0) & "'" & " heigth=" & X(1) & " with= " & X(2) & "/>"
            Else
                texto = texto & v(1) & "</p>"
            End If
        End If
    Else
        texto = gnc.Empresa.GNOpcion.MensajeCorreo
    End If
        
        
'        If gnc.Empresa.GNOpcion.BandCorreoAutomatico Then
            Nombre = gnc.Empresa.GNOpcion.ComprobantesAutorizados & "\" & Mid$(gnc.ClaveAcceso, 1, 39) & Right("0000000000" & gnc.transid, 10)
            grd.TextMatrix(FilaComp, COL_NOMBRE) = "\" & Mid$(gnc.ClaveAcceso, 1, 39) & Right("0000000000" & gnc.transid, 10)
            filename = Dir(Nombre & ".xml", vbArchive)
            
            strArchivoXML = Nombre & ".xml"
            strArchivoPDF = Nombre & ".pdf"
            
            If Len(email) > 0 And grd.ValueMatrix(FilaComp, COL_TIPO) <> 6 Then
                valor = "c:\ia\EnviarCorreoOculto.exe """ & gnc.Empresa.GNOpcion.ServidorCorreo & """ " & gnc.Empresa.GNOpcion.PuertoCorreo & " """ & gnc.Empresa.GNOpcion.CuentaCorreo & _
                        """ """ & gnc.Empresa.GNOpcion.PasswordCorreo & """ """ & gnc.Empresa.GNOpcion.NombreUsuario & """ """ & email & _
                        """ """ & gnc.Empresa.GNOpcion.CopiaCorreo & """ " & Cifrado & " " & Proxy & " " & gnc.Empresa.GNOpcion.PuertoCorreo & " """ & "servidorproxy" & """ """ & gnc.Empresa.GNOpcion.NombreEmpresa & _
                        """ """ & asunto & """ """ & texto & """ """ & strArchivoPDF & """ """ & strArchivoXML & """ "
                        
                Shell valor, vbHide
            End If
'        End If
    
        
        Set pc = Nothing
        Set trans = Nothing
        
      If gnc.Empresa.ActualizaBanderaGeneraRide(gnc.transid) Then
      End If


    Exit Sub
Errtrap:

        Set pc = Nothing
        Set trans = Nothing

    grd.TextMatrix(FilaComp, COL_PDF) = "0"
    If Err.Number <> 32755 Then
        MsgBox Err.Description
    End If
    Exit Sub
        
    'End With
    
'    Unload Me
End Sub





Private Sub TimerSeg_Timer()
    Dim mobjxml As Object, i As Integer, X As Single
    lblseg.Caption = Val(lblseg.Caption) - 1

    grd.Row = FilaComp
    X = grd.CellTop                 'Para visualizar la celda actual

'    If lblseg.Caption = "40" Then
            If lblseg.Caption = "20" Then
                If grd.ValueMatrix(FilaComp, COL_TID) <> 0 Then
                    Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(grd.ValueMatrix(FilaComp, COL_TID))
                End If
'            ElseIf lblseg.Caption = "25" Then
            ElseIf lblseg.Caption = "15" Then
                If Not gnc Is Nothing Then
                        GeneraRidePDF mobjxml
                        If Not gobjMain.EmpresaActual.GNOpcion.BandCorreoAutomatico Then
                            lblseg.Caption = "4"
                        End If
                End If
'            ElseIf lblseg.Caption = "15" Then
            ElseIf lblseg.Caption = "10" Then
                    If Not gnc Is Nothing Then
                        If grd.ValueMatrix(FilaComp, COL_TIPO) <> 6 Then
                            If gobjMain.EmpresaActual.GNOpcion.BandCorreoAutomatico Then
                                EmailAuto
                            End If
                        Else
                            grd.TextMatrix(FilaComp, COL_NOMBRE) = "\" & Mid$(gnc.ClaveAcceso, 1, 39) & Right("0000000000" & gnc.transid, 10)
                            If gobjMain.EmpresaActual.ActualizaBanderaGeneraRide(gnc.transid) Then
                            End If
                            lblseg.Caption = "6"
                           
                        End If
                    End If
                        
            ElseIf lblseg.Caption = "5" Then
                If Not gnc Is Nothing Then
                    If gnc.Empresa.GNOpcion.BandCorreoAutomatico Then
                        If grd.ValueMatrix(FilaComp, COL_PDF) <> 0 Then
                            MoverArchivos
                       End If
                    End If
                    lblseg.Caption = "1"
                End If
            
            ElseIf lblseg.Caption = "0" Then
'                lblseg.Caption = "41"
                lblseg.Caption = "21"
                FilaComp = FilaComp + 1
                    
                If FilaComp <= MaxFila Then
                    grd.Row = FilaComp
                    X = grd.CellTop                 'Para visualizar la celda actual
                End If
                
                
                If FilaComp = (MaxFila + 1) Then
                    FilaComp = 1
                    For i = 1 To grd.Rows - 1
                         grd.RemoveItem 1
                     Next i
        '            grd.Row = FilaComp
        '            X = grd.CellTop                 'Para visualizar la celda actual
                If FilaComp <> 1 Then
                    grd.Row = FilaComp
                    X = grd.CellTop                 'Para visualizar la celda actual
                End If
                    
                   AutoInicio
                End If
                
            End If
'            Set gnc = Nothing
   
    

End Sub


Private Function TotalRetencion(ByRef GnComp As GNComprobante) As Currency
    Dim tsk As TSKardexRet, i As Long
    Dim total As Currency, v As Variant
'     v = SeparaParamVar("TOTRET")
    For i = 1 To GnComp.CountTSKardexRet
        Set tsk = GnComp.TSKardexRet(i)
        total = total + Abs(tsk.Debe - tsk.Haber)   'Preguntar
    Next i
    TotalRetencion = total
End Function

