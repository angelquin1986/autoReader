VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmB_FormSRI103 
   Caption         =   "Condiciones de Busqueda"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5685
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkReembolso 
      Caption         =   "Totalizar Reembolso y (332)"
      Height          =   255
      Left            =   2850
      TabIndex        =   18
      Top             =   360
      Width           =   2415
   End
   Begin VB.ListBox LstRetencion103 
      Height          =   3765
      IntegralHeight  =   0   'False
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   17
      Top             =   840
      Width           =   5325
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
      ScaleWidth      =   5685
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4800
      Width           =   5685
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
      Caption         =   " Fecha"
      Height          =   705
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   2715
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   360
         Left            =   840
         TabIndex        =   2
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
         TabIndex        =   1
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
Attribute VB_Name = "frmB_FormSRI103"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private BandAceptado As Boolean
Private mobjSiiMain As SiiMain


Private Sub Habilita()
    dtpDesde.Enabled = True
    dtpHasta.Enabled = True
    fraFecha.Enabled = True
    fraTrans2.Enabled = True
    fraCobro.Enabled = True
    
    Label2.Visible = True
    Label3.Visible = True
End Sub



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


Private Sub cmdCancelar_Click()
    BandAceptado = False
    Me.Hide
End Sub


Private Sub Form_DragOver(Source As Control, x As Single, y As Single, State As Integer)
 ' Cambia el puntero a no colocar.
   If State = 0 Then Source.MousePointer = 12
   ' Utiliza el puntero predeterminado del mouse.
   If State = 1 Then Source.MousePointer = 0
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


Private Sub lstBienes_DblClick()
    cmdAdd_Click
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

Private Function PreparaCadenaCP(lst As ListBox) As String
    Dim Cadena As String, i As Integer, pos As Integer
    Cadena = ""
    For i = 0 To lst.ListCount - 1

            pos = InStr(1, lst.List(i), " ")
            If Cadena = "" Then
                
                Cadena = Trim(Mid$(lst.List(i), 1, pos))
            Else
                
                Cadena = Cadena & "," & Trim(Mid$(lst.List(i), 1, pos))
            End If
        'End If
    Next i
    PreparaCadenaCP = Cadena
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

Public Function Inicio103(ByRef objcond As Condicion, ByRef Reten As String) As Boolean
    Dim KeyTrans As String, KeyRecargo As String, KeyTransRet As String, KeyRet As String, KeyT As String
    Dim trans As String, s As String
    Dim bandRembolso As String
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
       
        CargaTransRetencion LstRetencion103, False
        Reten = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("103_TipoRetencion")
        bandRembolso = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("103_BandRembolso")
        If Len(bandRembolso) > 0 Then
            If bandRembolso Then
                chkReembolso.value = vbChecked
            Else
                chkReembolso.value = vbUnchecked
            End If
        End If
        cargaConfigLista Reten
        BandAceptado = False
        Screen.MousePointer = 0
        Me.Show vbModal
'        Si aplastó el botón 'Aceptar'
         If BandAceptado Then
'            'Devuelve las condiciones de búsqueda en el objeto objCondicion en SiiMain
            .fecha1 = dtpHasta.value
            'retenciones
            Reten = PreparaCadenaSeleccion(LstRetencion103)
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "103_TipoRetencion", Reten
            gobjMain.EmpresaActual.GNOpcion.AsignarValor "103_BandRembolso", (chkReembolso.value = vbChecked)
            .Estado1Bool(1) = (chkReembolso.value = vbChecked)
            'Graba en la base
            gobjMain.EmpresaActual.GNOpcion.Grabar
        End If
    End With
    Inicio103 = BandAceptado
        Unload Me
    Exit Function
ErrTrap:
    Screen.MousePointer = 0
    DispErr
    Exit Function
End Function
Private Sub cargaConfigLista(ByVal codigo As String)
Dim v
Dim i As Integer, j As Integer
v = Split(codigo, ",")
    For i = 0 To UBound(v)
        For j = 0 To LstRetencion103.ListCount - 1
            If v(i) = Left(LstRetencion103.List(j), LstRetencion103.ItemData(j)) Then
                LstRetencion103.Selected(j) = True
            End If
        Next
    Next
End Sub




