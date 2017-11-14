VERSION 5.00
Begin VB.Form frmB_VxTransComp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda"
   ClientHeight    =   4410
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5430
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraEntrega 
      Caption         =   "Transacciones de Entrega"
      Height          =   1812
      Left            =   120
      TabIndex        =   4
      Top             =   1980
      Width           =   5175
      Begin VB.ListBox lstE 
         Height          =   1368
         IntegralHeight  =   0   'False
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   312
         Width           =   4935
      End
   End
   Begin VB.Frame fraVenta 
      Caption         =   "Transacciones de Venta y Devolucion"
      Height          =   1812
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5175
      Begin VB.ListBox lst 
         Height          =   1368
         IntegralHeight  =   0   'False
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   312
         Width           =   4935
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   400
      Left            =   2760
      TabIndex        =   2
      Top             =   3900
      Width           =   1200
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar -F5"
      Height          =   400
      Left            =   1380
      TabIndex        =   1
      Top             =   3900
      Width           =   1200
   End
End
Attribute VB_Name = "frmB_VxTransComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private BandAceptado As Boolean

Public Function Inicio(ByRef objcond As Condicion, _
                                    ByVal tag As String) As Boolean
    Dim KeyTrans As String, KeyTransE As String
    Dim s As String
    Me.tag = tag
   
    With objcond
        
        CargaTipoTrans "IV", lst
        CargaTipoTrans "IV", lstE

        BandAceptado = False
        'KeyTrans = "TVentaCompr_Trans"
        'KeyTransE = "TEntregaCompr_Trans"

        KeyTrans = GetSetting(APPNAME, App.Title, "TVentaCompr_Trans", "_VACIO_")
        KeyTransE = GetSetting(APPNAME, App.Title, "TEntregaCompr_Trans", "_VACIO_")
        
        RecuperaSelecTrans KeyTrans, lst
        RecuperaSelecTrans KeyTransE, lstE

        Me.Show vbModal, frmMain
        'Si aplastó el botón 'Aceptar'
        If BandAceptado Then
            'Devuelve los valores de condición para la búsqueda
            .CodTrans = PreparaCadena(lst)
            .Bienes = PreparaCadena(lstE)
            'grabar las formas de cobro a visualizar
            SaveSetting APPNAME, App.Title, "TVentaCompr_Trans", .CodTrans
            SaveSetting APPNAME, App.Title, "TEntregaCompr_Trans", .Bienes
        End If
    End With
    'Devuelve true/false
    Unload Me
    Inicio = BandAceptado
End Function


Private Function PreparaCadena(lst As ListBox) As String
    Dim Cadena As String, i As Integer
    Cadena = ""
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            If Cadena = "" Then
                Cadena = "'" & Left(lst.List(i), lst.ItemData(i)) & "'"
            Else
                Cadena = Cadena & "," & _
                              "'" & Left(lst.List(i), lst.ItemData(i)) & "'"
            End If
        End If
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

Private Sub PreparaListaTransIV()
    Dim rs As Recordset
   'Prepara la lista de tipos de transaccion
    lst.Clear
    Set rs = gobjMain.EmpresaActual.ListaGNTrans("IV", False, True)
    With rs
        If Not (.EOF) Then
            .MoveFirst
            Do Until .EOF
                lst.AddItem !CodTrans & "  " & !NombreTrans
                lst.ItemData(lst.NewIndex) = Len(!CodTrans)
                .MoveNext
            Loop
        End If
    End With
    rs.Close
    Set rs = Nothing
End Sub


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


Private Sub Form_Load()
    'Establece los rangos de Fecha  siempre  al rango
    'del año actual
End Sub


Private Sub CargaTipoTrans(ByRef Modulo As String, ByRef lst As ListBox)
    Dim rs As Recordset, Vector As Variant
    Dim numMod As Integer, i As Integer
    'Prepara la lista de tipos de transaccion
    lst.Clear
    Vector = Split(Modulo, ",")
    numMod = UBound(Vector, 1)
    If numMod = -1 Then
        Set rs = gobjMain.EmpresaActual.ListaGNTrans("", False, True)
        With rs
            If Not (.EOF) Then
                .MoveFirst
                Do Until .EOF
                    lst.AddItem !CodTrans & "  " & !NombreTrans
                    lst.ItemData(lst.NewIndex) = Len(!CodTrans)
                    .MoveNext
                Loop
            End If
        End With
        rs.Close
    Else
        For i = 0 To numMod
            Set rs = gobjMain.EmpresaActual.ListaGNTrans(CStr(Vector(i)), False, True)
            With rs
                If Not (.EOF) Then
                    .MoveFirst
                    Do Until .EOF
                        lst.AddItem !CodTrans & "  " & !NombreTrans
                        lst.ItemData(lst.NewIndex) = Len(!CodTrans)
                        .MoveNext
                    Loop
                End If
            End With
            rs.Close
        Next i
    End If
    Set rs = Nothing
End Sub

Private Sub RecuperaSelecTrans(s As String, lst As ListBox)
    Dim Vector As Variant
    Dim i As Integer, j As Integer, Selec As Integer
        If s <> "_VACIO_" Then
        Vector = Split(s, ",")
         Selec = UBound(Vector, 1)
         For i = 0 To Selec
            For j = 0 To lst.ListCount - 1
                If Mid$(Vector(i), 2, Len(Vector(i)) - 2) = Left(lst.List(j), lst.ItemData(j)) Then
                    lst.Selected(j) = True
                End If
            Next j
         Next i
    End If
End Sub

