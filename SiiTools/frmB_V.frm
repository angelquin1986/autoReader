VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmB_V 
   Caption         =   "Condiciones de Búsqueda"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4920
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   2580
      TabIndex        =   6
      Top             =   2700
      Width           =   1400
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar F5"
      Height          =   495
      Left            =   900
      TabIndex        =   5
      Top             =   2700
      Width           =   1400
   End
   Begin VB.Frame Frame4 
      Caption         =   "Fecha Transacción"
      Height          =   2415
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   4575
      Begin MSComCtl2.DTPicker dtpFecha1 
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   106692609
         CurrentDate     =   38010
      End
      Begin MSComCtl2.DTPicker dtpFecha2 
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   106692609
         CurrentDate     =   38010
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   420
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   15
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   465
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Destino"
      Height          =   2415
      Left            =   2460
      TabIndex        =   9
      Top             =   120
      Width           =   2295
      Begin VB.ListBox lstDestino 
         Height          =   1860
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Trafico"
      Height          =   2415
      Left            =   2520
      TabIndex        =   8
      Top             =   120
      Width           =   2295
      Begin VB.ListBox lstTrafico 
         Height          =   1860
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cabinas"
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2295
      Begin VB.ListBox lstCabinas 
         Height          =   1860
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmB_V"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mBandAceptado As Boolean

Public Function Inicio(ByVal tag As String, _
                       ByRef lst_cabina As String, _
                       ByRef lst_trafico As String, _
                       ByRef lst_destino As String, _
                       ByRef fecha1 As Date, _
                       ByRef fecha2 As Date) As Boolean
    Me.tag = tag
    CargarListas
    dtpFecha1.value = fecha1
    dtpFecha2.value = fecha2
    RecuperaSelCabinas lst_cabina
    RecuperaSelTrafico lst_trafico
    RecuperaSelDestino lst_destino
        
    Me.Show vbModal
    If mBandAceptado Then
        lst_cabina = ElementosSeleccionados(lstCabinas)
        lst_trafico = ElementosSeleccionados(lstTrafico)
        lst_destino = ElementosSeleccionados(lstDestino)
        fecha1 = dtpFecha1.value
        fecha2 = dtpFecha2.value
    End If
    Inicio = mBandAceptado
End Function

Private Sub RecuperaSelCabinas(lista As String)
    Dim v As Variant, i As Integer, n As Integer, s As String
    If Len(lista) = 0 Then Exit Sub
    v = Split(lista, ",")
    With lstCabinas
        For n = LBound(v, 1) To UBound(v, 1)
            s = v(n)
            'Quita las comillas simples
            s = Mid$(s, 2, Len(s) - 2)
            For i = 0 To .ListCount - 1
                If .List(i) = s Then
                    .Selected(i) = True
                    Exit For
                End If
            Next i
        Next n
    End With
End Sub

Private Sub RecuperaSelTrafico(lista As String)
    Dim v As Variant, i As Integer, n As Integer, s As String
    If Len(lista) = 0 Then Exit Sub
    v = Split(lista, ",")
    With lstTrafico
        For n = LBound(v, 1) To UBound(v, 1)
            s = v(n)
            'Quita las comillas simples
            s = Mid$(s, 2, Len(s) - 2)
            For i = 0 To .ListCount - 1
                If .List(i) = s Then
                    .Selected(i) = True
                    Exit For
                End If
            Next i
        Next n
    End With
End Sub

Private Sub RecuperaSelDestino(lista As String)
    Dim v As Variant, i As Integer, n As Integer, s As String
    If Len(lista) = 0 Then Exit Sub
    v = Split(lista, ",")
    With lstDestino
        For n = LBound(v, 1) To UBound(v, 1)
            s = v(n)
            'Quita las comillas simples
            s = Mid$(s, 2, Len(s) - 2)
            For i = 0 To .ListCount - 1
                If .List(i) = s Then
                    .Selected(i) = True
                    Exit For
                End If
            Next i
        Next n
    End With
End Sub

Private Function ElementosSeleccionados(ByVal lst As ListBox) As String
    Dim i As Integer, s As String
    With lst
        For i = 0 To lst.ListCount - 1
            If .Selected(i) Then
                s = s & "'" & .List(i) & "'" & ","
            End If
        Next i
    End With
    If Right(s, 1) = "," Then s = Mid$(s, 1, Len(s) - 1)
    ElementosSeleccionados = s
End Function

Private Sub CargarListas()
End Sub

Private Sub CargarDestinos(ByVal cond_traf As String)
    
End Sub

Private Sub cmdAceptar_Click()
    mBandAceptado = True
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    mBandAceptado = False
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

Private Sub lstTrafico_Click()
    Dim lista As String, i As Integer
    lista = ""
    For i = 0 To lstTrafico.ListCount - 1
        If lstTrafico.Selected(i) Then
            lista = lista & "'" & lstTrafico.List(i) & "'" & ","
        End If
    Next i
    If Right(lista, 1) = "," Then lista = Mid$(lista, 1, Len(lista) - 1)
    CargarDestinos lista
End Sub

Public Function InicioF101(ByVal tag As String, _
                       ByRef fecha1 As Date, _
                       ByRef fecha2 As Date) As Boolean
    Me.tag = tag
''    CargarListas
    dtpFecha1.value = fecha1
    dtpFecha2.value = fecha2
''    RecuperaSelCabinas lst_cabina
''    RecuperaSelTrafico lst_trafico
''    RecuperaSelDestino lst_destino
        
    Me.Show vbModal
    If mBandAceptado Then
''        lst_cabina = ElementosSeleccionados(lstCabinas)
''        lst_trafico = ElementosSeleccionados(lstTrafico)
''        lst_destino = ElementosSeleccionados(lstDestino)
        fecha1 = dtpFecha1.value
        fecha2 = dtpFecha2.value
    End If
    InicioF101 = mBandAceptado
End Function


