VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Begin VB.Form frmPrecios 
   Caption         =   "Actualización de precios"
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
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   852
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   6240
      TabIndex        =   23
      Top             =   3828
      Width           =   6240
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   372
         Left            =   1320
         TabIndex        =   26
         Top             =   0
         Width           =   1452
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   3480
         TabIndex        =   25
         Top             =   0
         Width           =   1452
      End
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   24
         Top             =   540
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Frame fraItem 
      Caption         =   "&Rango de items"
      Height          =   1932
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   6012
      Begin FlexComboProy.FlexCombo fcbDesde 
         Height          =   348
         Left            =   840
         TabIndex        =   20
         Top             =   1080
         Width           =   4932
         _ExtentX        =   8705
         _ExtentY        =   609
         ColWidth1       =   2400
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
      Begin VB.OptionButton optPorGrupo 
         Caption         =   "Por Grupo5"
         Height          =   240
         Index           =   4
         Left            =   3960
         TabIndex        =   18
         Top             =   480
         Width           =   1812
      End
      Begin VB.OptionButton optPorGrupo 
         Caption         =   "Por Grupo4"
         Height          =   240
         Index           =   3
         Left            =   3960
         TabIndex        =   17
         Top             =   240
         Width           =   1812
      End
      Begin VB.OptionButton optPorCodigo 
         Caption         =   "Por Código"
         Height          =   240
         Left            =   840
         TabIndex        =   13
         Top             =   240
         Width           =   1452
      End
      Begin VB.OptionButton optPorGrupo 
         Caption         =   "Por Grupo1"
         Height          =   240
         Index           =   0
         Left            =   2520
         TabIndex        =   14
         Top             =   240
         Width           =   1572
      End
      Begin VB.OptionButton optPorGrupo 
         Caption         =   "Por Grupo2"
         Height          =   240
         Index           =   1
         Left            =   2520
         TabIndex        =   15
         Top             =   480
         Width           =   1692
      End
      Begin VB.OptionButton optPorGrupo 
         Caption         =   "Por Grupo3"
         Height          =   240
         Index           =   2
         Left            =   2520
         TabIndex        =   16
         Top             =   720
         Width           =   1812
      End
      Begin FlexComboProy.FlexCombo fcbHasta 
         Height          =   348
         Left            =   840
         TabIndex        =   22
         Top             =   1440
         Width           =   4932
         _ExtentX        =   8705
         _ExtentY        =   609
         ColWidth1       =   2400
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
      Begin VB.Label Label2 
         Caption         =   "&Desde"
         Height          =   252
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   732
      End
      Begin VB.Label Label3 
         Caption         =   "&Hasta"
         Height          =   252
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   732
      End
   End
   Begin VB.Frame fraPrecio 
      Caption         =   "Acción"
      Height          =   1332
      Left            =   132
      TabIndex        =   0
      Top             =   120
      Width           =   6000
      Begin VB.CheckBox chkPrecio 
         Caption         =   "Precio&4"
         Height          =   252
         Index           =   3
         Left            =   2520
         TabIndex        =   9
         Top             =   960
         Value           =   1  'Checked
         Width           =   1212
      End
      Begin VB.TextBox txtPorcent 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   700
      End
      Begin VB.CheckBox chkPrecio 
         Caption         =   "Precio&1"
         Height          =   252
         Index           =   0
         Left            =   2520
         TabIndex        =   6
         Top             =   240
         Value           =   1  'Checked
         Width           =   1212
      End
      Begin VB.CheckBox chkPrecio 
         Caption         =   "Precio&2"
         Height          =   252
         Index           =   1
         Left            =   2520
         TabIndex        =   7
         Top             =   480
         Value           =   1  'Checked
         Width           =   1212
      End
      Begin VB.CheckBox chkPrecio 
         Caption         =   "Precio&3"
         Height          =   252
         Index           =   2
         Left            =   2520
         TabIndex        =   8
         Top             =   720
         Value           =   1  'Checked
         Width           =   1212
      End
      Begin VB.OptionButton optAlzar 
         Caption         =   "Al&zar"
         Height          =   240
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   972
      End
      Begin VB.OptionButton optBajar 
         Caption         =   "&Bajar"
         Height          =   240
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   972
      End
      Begin VB.ComboBox cboRedondeo 
         Height          =   288
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   480
         Width           =   1812
      End
      Begin VB.OptionButton optCalcular 
         Caption         =   "&Calcular de costo"
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   2052
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         Height          =   252
         Left            =   2100
         TabIndex        =   5
         Top             =   360
         Width           =   372
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "R&edondeo"
         Height          =   240
         Left            =   4080
         TabIndex        =   10
         Top             =   240
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub Inicio()
    Dim i As Integer
    On Error GoTo ErrTrap
    
    Me.Show
    Me.ZOrder
    Me.Caption = "Actualización de precios"
    
    cboRedondeo.Clear
    For i = -2 To 6
        cboRedondeo.AddItem Format(10 ^ i, "#,0.00")
    Next i
    cboRedondeo.ListIndex = 0
    
    optPorGrupo(0).Caption = "Por " & gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(1)
    optPorGrupo(1).Caption = "Por " & gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(2)
    optPorGrupo(2).Caption = "Por " & gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(3)
    optPorGrupo(3).Caption = "Por " & gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(4)
    optPorGrupo(4).Caption = "Por " & gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(5)
    
'    optPorCodigo.Value = True
'    optPorCodigo_Click          'Carga items
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub


Private Sub cmdAceptar_Click()
    If fcbDesde.Vacio Then
        MsgBox "Seleccione el inicio de rango."
        fcbDesde.SetFocus
        Exit Sub
    End If
    If fcbHasta.Vacio Then
        MsgBox "Seleccione el fin de rango."
        fcbHasta.SetFocus
        Exit Sub
    End If
    
    If (chkPrecio(0).value <> vbChecked) And _
       (chkPrecio(1).value <> vbChecked) And _
       (chkPrecio(2).value <> vbChecked) And _
       (chkPrecio(3).value <> vbChecked) Then
        MsgBox "Al menos uno de los precios debe seleccionar para actualizar."
        chkPrecio(0).SetFocus
        Exit Sub
    End If
    
    ActualizaPrecio
End Sub

Private Function ActualizaPrecio() As Boolean
    Dim upd As String, betw As String, s As String, p As Single
    Dim v As Variant, i As Long, cod As String, j As Integer
    Dim item As IVinventario, cap As String, msg As String
    Dim codItemDesde As String, CodItemHasta  As String, _
        codG1Desde As String, CodG1Hasta As String, _
        codG2Desde As String, CodG2Hasta As String, _
        codG3Desde As String, CodG3Hasta As String, _
        codG4Desde As String, CodG4Hasta As String, _
        codG5Desde As String, CodG5Hasta As String
    
    On Error GoTo ErrTrap
    cap = Me.Caption

    s = "Está seguro que desea "
    If optAlzar.value Then
        s = s & "ALZAR"
    ElseIf optBajar.value Then
        s = s & "BAJAR"
    Else
        s = s & "CALCULAR DE COSTO"
    End If
    s = s & " los precios por " & Val(txtPorcent.Text) & " por ciento?"
    If MsgBox(s, vbYesNo + vbQuestion) <> vbYes Then Exit Function

    If optAlzar.value Then
        p = 1 + (Val(txtPorcent.Text) / 100#)   'Alzar
    ElseIf optBajar.value Then
        p = 1 - (Val(txtPorcent.Text) / 100#)   'Bajar
    Else
        p = Val(txtPorcent.Text) / 100#         'Calcular
    End If
    
    If optPorCodigo.value Then
        codItemDesde = fcbDesde.Text
        CodItemHasta = fcbHasta.Text
    ElseIf optPorGrupo(0).value Then
        codG1Desde = fcbDesde.Text
        CodG1Hasta = fcbHasta.Text
    ElseIf optPorGrupo(1).value Then
        codG2Desde = fcbDesde.Text
        CodG2Hasta = fcbHasta.Text
    ElseIf optPorGrupo(2).value Then
        codG3Desde = fcbDesde.Text
        CodG3Hasta = fcbHasta.Text
    ElseIf optPorGrupo(3).value Then
        codG4Desde = fcbDesde.Text
        CodG4Hasta = fcbHasta.Text
    ElseIf optPorGrupo(4).value Then
        codG5Desde = fcbDesde.Text
        CodG5Hasta = fcbHasta.Text
    End If
    
    
    Screen.MousePointer = vbHourglass
    
    v = gobjMain.EmpresaActual.ListaIVInventarioPorRango( _
                                codItemDesde, CodItemHasta, _
                                codG1Desde, CodG1Hasta, _
                                codG2Desde, CodG2Hasta, _
                                codG3Desde, CodG3Hasta, _
                                codG4Desde, CodG4Hasta, _
                                codG5Desde, CodG5Hasta)
    If Not IsEmpty(v) Then
        prg1.min = 0
        prg1.max = UBound(v, 2) + 1
        prg1.value = 0
        For i = 0 To UBound(v, 2)
            cod = v(0, i)
            Set item = gobjMain.EmpresaActual.RecuperaIVInventario(cod)
            If Not (item Is Nothing) Then
                Me.Caption = item.CodInventario
                
                For j = 1 To 4
                    'Si está seleccionado para modificar
                    If chkPrecio(j - 1).value = vbChecked Then
                        If Not optCalcular.value Then
                            'Sube/baja
                            item.Precio(j) = item.Precio(j) * p
                        Else
                            'Calcula de costo
                            item.Precio(j) = item.costo(Now, 0) * p
                        End If
                        
                        'Redondea el precio
                        item.RedondearPrecio j, (cboRedondeo.ListIndex - 2)
                    End If
                Next j
                
                'Grabar el cambio
                item.Grabar
            Else
                msg = "No se encuentra el item '" & cod & "'." & vbCr & vbCr & _
                      "Desea continuar el proceso?"
                If MsgBox(msg, vbYesNo + vbQuestion) <> vbYes Then Exit For
            End If
            DoEvents
            prg1.value = i + 1
        Next i
    End If
    Me.Caption = cap
    Screen.MousePointer = 0
    MsgBox "La actualización de precios se ha finalizado.", vbInformation
    prg1.value = prg1.min
    ActualizaPrecio = True
    Exit Function
ErrTrap:
    Me.Caption = cap
    Screen.MousePointer = 0
    DispErr
    prg1.value = prg1.min
    Exit Function
End Function


Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub fcbDesde_Selected(ByVal Text As String, ByVal KeyText As String)
    fcbHasta.KeyText = fcbDesde.KeyText
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    ImpideSonidoEnter Me, KeyAscii
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Me.Hide         'Se pone esto para evitar el posible BUG de Windows98
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    prg1.Width = Me.ScaleWidth - (prg1.Left * 2)
End Sub

Private Sub optPorCodigo_Click()
    Dim v As Variant, i As Long
    On Error GoTo ErrTrap
    Screen.MousePointer = vbHourglass
    fcbDesde.Clear
    fcbHasta.Clear
    
    v = gobjMain.EmpresaActual.ListaIVInventarioSimple
    fcbDesde.SetData v
    fcbHasta.SetData v
    Screen.MousePointer = 0
    Exit Sub
ErrTrap:
    Screen.MousePointer = 0
    DispErr
    Exit Sub
End Sub

Private Sub optPorGrupo_Click(Index As Integer)
    Dim v As Variant, i As Long
    On Error GoTo ErrTrap
    Screen.MousePointer = vbHourglass
    
    fcbDesde.Clear
    fcbHasta.Clear
    
    v = gobjMain.EmpresaActual.ListaIVGrupo(Index + 1, False, False)
    fcbDesde.SetData v
    fcbHasta.SetData v
    Screen.MousePointer = 0
    Exit Sub
ErrTrap:
    Screen.MousePointer = 0
    DispErr
    Exit Sub
End Sub

Private Sub txtPorcent_KeyPress(KeyAscii As Integer)
    'Acepta solo numericos, BackSpace y punto decimal
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
        (KeyAscii <> vbKeyBack) And _
        (KeyAscii <> Asc(".")) Then
        KeyAscii = 0
    End If
End Sub
