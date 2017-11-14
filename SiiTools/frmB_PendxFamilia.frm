VERSION 5.00
Begin VB.Form frmB_PendxFamilia 
   Caption         =   "Condiciones de Busqueda"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3450
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Transacciones de Devol."
      Height          =   2775
      Left            =   4740
      TabIndex        =   6
      Top             =   120
      Width           =   2235
      Begin VB.ListBox lstDev 
         Height          =   2355
         IntegralHeight  =   0   'False
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   240
         Width           =   1995
      End
   End
   Begin VB.Frame fra1 
      Caption         =   "Transacciones de Salida"
      Height          =   2775
      Left            =   2460
      TabIndex        =   4
      Top             =   120
      Width           =   2235
      Begin VB.ListBox lstSalida 
         Height          =   2355
         IntegralHeight  =   0   'False
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   240
         Width           =   1995
      End
   End
   Begin VB.Frame fra2 
      Caption         =   "Transacciones con Familias"
      Height          =   2775
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2235
      Begin VB.ListBox lst 
         Height          =   2445
         IntegralHeight  =   0   'False
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   192
         Width           =   1995
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   400
      Left            =   3660
      TabIndex        =   3
      Top             =   3000
      Width           =   1320
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar -F5"
      Height          =   400
      Left            =   2220
      TabIndex        =   2
      Top             =   3000
      Width           =   1320
   End
End
Attribute VB_Name = "frmB_PendxFamilia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BandAceptado As Boolean



Public Function InicioPendientesxFamilia(ByRef objcond As RepCondicion, ByVal tag As String) As Boolean
    Dim i As Integer, KeyS As String, KeyF As String, s As String, KeyD
    Dim trans As String
    With objcond
        CargaTipoTrans "IV", lstSalida
        CargaTipoTrans "IV", lst    'Carga  lista para facturas
        CargaTipoTrans "IV", lstDev    'Carga  lista para facturas
        BandAceptado = False
        KeyS = "PendientexFamilia_TransSalida"
        KeyF = "PendientexFamilia_TransFactura"
        KeyD = "PendientexFamilia_TransDev"

    'jeaa 06/06/2007
        Select Case tag
        Case "Items"
            fra2.Caption = "Transaccion de Venta"
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarFacturar")) > 0 Then
                s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarFacturar")
                RecuperaTrans "KeyF", lst, s
            End If
    
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarSalida")) > 0 Then
                s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarSalida")
                RecuperaTrans "KeyS", lstSalida, s
            End If
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarDev")) > 0 Then
                s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarDev")
                RecuperaTrans "KeyD", lstDev, s
            End If
        
        
        Case "Familias"
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarFacturarF")) > 0 Then
                s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarFacturarF")
                RecuperaTrans "KeyF", lst, s
            End If
    
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarSalidaF")) > 0 Then
                s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarSalidaF")
                RecuperaTrans "KeyS", lstSalida, s
            End If
            
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarDevF")) > 0 Then
                s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarDevF")
                RecuperaTrans "KeyD", lstDev, s
            End If
        
        
        Case "ItemsHormi"
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarFacturarItem")) > 0 Then
                s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarFacturarItem")
                RecuperaTrans "KeyF", lst, s
            End If
    
    
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarSalidaItem")) > 0 Then
                s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarSalidaItem")
                RecuperaTrans "KeyS", lstSalida, s
            End If
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarDevItem")) > 0 Then
                s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarDevItem")
                RecuperaTrans "KeyD", lstDev, s
            End If
        
        
        
        End Select

        'Valores predeterminados
        Me.Show vbModal, frmMain
        'Si aplastó el botón 'Aceptar'
        If BandAceptado Then
             .SQLItem = PreparaCadena(lst)
            .tipoTrans = PreparaCadena(lstSalida)
            .CodTrans = PreparaCadena(lstDev)
            .Bandera = True
            Select Case tag
            Case "Items"
                s = PreparaTransParaGnopcion(.SQLItem)
                gobjMain.EmpresaActual.GNOpcion.AsignarValor "CierreTransXEntregarFacturar", s
                s = PreparaTransParaGnopcion(.tipoTrans)
                gobjMain.EmpresaActual.GNOpcion.AsignarValor "CierreTransXEntregarSalida", s
                s = PreparaTransParaGnopcion(.CodTrans)
                gobjMain.EmpresaActual.GNOpcion.AsignarValor "CierreTransXEntregarDev", s
                
            Case "Familias"
                s = PreparaTransParaGnopcion(.SQLItem)
                gobjMain.EmpresaActual.GNOpcion.AsignarValor "CierreTransXEntregarFacturarF", s
                s = PreparaTransParaGnopcion(.tipoTrans)
                gobjMain.EmpresaActual.GNOpcion.AsignarValor "CierreTransXEntregarSalidaF", s
                s = PreparaTransParaGnopcion(.CodTrans)
                gobjMain.EmpresaActual.GNOpcion.AsignarValor "CierreTransXEntregarDevF", s
            
            Case "ItemsHormi"
                s = PreparaTransParaGnopcion(.SQLItem)
                gobjMain.EmpresaActual.GNOpcion.AsignarValor "CierreTransXEntregarFacturarItem", s
                s = PreparaTransParaGnopcion(.tipoTrans)
                gobjMain.EmpresaActual.GNOpcion.AsignarValor "CierreTransXEntregarSalidaItem", s
                s = PreparaTransParaGnopcion(.CodTrans)
                gobjMain.EmpresaActual.GNOpcion.AsignarValor "CierreTransXEntregarDevItem", s
            
            End Select
            'Graba en la base
            gobjMain.EmpresaActual.GNOpcion.GrabarSoloGnOpcion2
        End If
    End With
    Unload Me
    InicioPendientesxFamilia = BandAceptado
End Function



Private Function PreparaCadena(lst As ListBox) As String
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



Private Sub cmdAceptar_Click()
    BandAceptado = True
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    BandAceptado = False
    Me.Hide
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

'jeaa 25/09/2006 elimina los apostrofes
Private Function PreparaTransParaGnopcion(cad As String) As String
    Dim v As Variant, i As Integer, s As String
    s = ""
    v = Split(cad, ",")
    For i = 0 To UBound(v)
        v(i) = Trim(v(i))
        s = s & Trim$(v(i)) & ","
    Next i
    'quita ultima coma
    PreparaTransParaGnopcion = Mid$(s, 1, Len(s) - 1)
End Function

Public Sub RecuperaTrans(ByVal Key As String, lst As ListBox, Optional s As String)
Dim Vector As Variant
Dim i As Integer, j As Integer, Selec As Integer, pos As Integer, CodTrans As String
    If s <> "_VACIO_" Then
        Vector = Split(s, ",")
         Selec = UBound(Vector, 1)
         For i = 0 To Selec
            For j = 0 To lst.ListCount - 1
                pos = InStr(1, lst.List(j), " ")
                If pos <> 0 Then
                    CodTrans = Trim$(Mid$(lst.List(j), 1, pos - 1))
                    If Trim(Vector(i)) = CodTrans Then
                        lst.Selected(j) = True
                    End If
                End If
            Next j
         Next i
    End If
End Sub
'AUC Pendientes de producir mahaivkaiv
Public Function InicioPendientesProduccion(ByRef objcond As RepCondicion, ByVal tag As String) As Boolean
    Dim i As Integer, KeyP As String, s As String
    Dim trans As String
    Me.Width = fra2.Width + 300
    fra2.Caption = "Pendientes Producir"
    cmdAceptar.Left = 10
    cmdCancelar.Left = cmdAceptar.Width + 10
    With objcond
        'CargaTipoTrans "IV", lstSalida
        CargaTipoTrans "IV", lst
        'CargaTipoTrans "IV", lstDev    'Carga  lista para facturas
        BandAceptado = False
        KeyP = "PendientexProducir"
        
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXProducir")) > 0 Then
                s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXProducir")
                RecuperaTrans "KeyP", lst, s
            End If
    
        

        'Valores predeterminados
        Me.Show vbModal, frmMain
        'Si aplastó el botón 'Aceptar'
        If BandAceptado Then
            .tipoTrans = PreparaCadena(lst)
            .Bandera = True
            s = PreparaTransParaGnopcion(.tipoTrans)
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "CierreTransXProducir", s
            'Graba en la base
            gobjMain.EmpresaActual.GNOpcion.Grabar
        End If
    End With
    Unload Me
    InicioPendientesProduccion = BandAceptado
End Function

Public Function InicioPendientesxTicket(ByRef objcond As RepCondicion, ByVal tag As String) As Boolean
    Dim i As Integer, KeyS As String, KeyF As String, s As String, KeyD
    Dim trans As String
    With objcond
        CargaTipoTrans "IV", lstSalida
        CargaTipoTrans "IV", lst    'Carga  lista para facturas
        CargaTipoTrans "IV", lstDev    'Carga  lista para facturas
        BandAceptado = False
        KeyS = "PendienteXTicket_TransSalida"
        KeyF = "PendienteXTicket_TransIngreso"
        KeyD = "PendienteXTicket_TransDev"

    'jeaa 06/06/2007
        Select Case tag
        Case "Ingreso"
            fra2.Caption = "Transaccion de Ingreso"
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXTicketIngreso")) > 0 Then
                s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXTicketIngreso")
                RecuperaTrans "KeyF", lst, s
            End If
    
        Case "Familias"
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarFacturarF")) > 0 Then
                s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarFacturarF")
                RecuperaTrans "KeyF", lst, s
            End If
    
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarSalidaF")) > 0 Then
                s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarSalidaF")
                RecuperaTrans "KeyS", lstSalida, s
            End If
            
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarDevF")) > 0 Then
                s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarDevF")
                RecuperaTrans "KeyD", lstDev, s
            End If
        
        
        Case "ItemsHormi"
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarFacturarItem")) > 0 Then
                s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarFacturarItem")
                RecuperaTrans "KeyF", lst, s
            End If
    
    
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarSalidaItem")) > 0 Then
                s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarSalidaItem")
                RecuperaTrans "KeyS", lstSalida, s
            End If
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarDevItem")) > 0 Then
                s = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("CierreTransXEntregarDevItem")
                RecuperaTrans "KeyD", lstDev, s
            End If
        
        
        
        End Select

        'Valores predeterminados
        Me.Show vbModal, frmMain
        'Si aplastó el botón 'Aceptar'
        If BandAceptado Then
             .SQLItem = PreparaCadena(lst)
            .tipoTrans = PreparaCadena(lstSalida)
            .CodTrans = PreparaCadena(lstDev)
            .Bandera = True
            Select Case tag
            Case "Ingreso"
                s = PreparaTransParaGnopcion(.SQLItem)
                gobjMain.EmpresaActual.GNOpcion.AsignarValor "CierreTransXTicketIngreso", s
                
                
            Case "Familias"
                s = PreparaTransParaGnopcion(.SQLItem)
                gobjMain.EmpresaActual.GNOpcion.AsignarValor "CierreTransXEntregarFacturarF", s
                s = PreparaTransParaGnopcion(.tipoTrans)
                gobjMain.EmpresaActual.GNOpcion.AsignarValor "CierreTransXEntregarSalidaF", s
                s = PreparaTransParaGnopcion(.CodTrans)
                gobjMain.EmpresaActual.GNOpcion.AsignarValor "CierreTransXEntregarDevF", s
            
            Case "ItemsHormi"
                s = PreparaTransParaGnopcion(.SQLItem)
                gobjMain.EmpresaActual.GNOpcion.AsignarValor "CierreTransXEntregarFacturarItem", s
                s = PreparaTransParaGnopcion(.tipoTrans)
                gobjMain.EmpresaActual.GNOpcion.AsignarValor "CierreTransXEntregarSalidaItem", s
                s = PreparaTransParaGnopcion(.CodTrans)
                gobjMain.EmpresaActual.GNOpcion.AsignarValor "CierreTransXEntregarDevItem", s
            
            End Select
            'Graba en la base
            gobjMain.EmpresaActual.GNOpcion.Grabar
        End If
    End With
    Unload Me
    InicioPendientesxTicket = BandAceptado
End Function

