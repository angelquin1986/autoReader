VERSION 5.00
Begin VB.Form frmExportarOpcion2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones de Exportación de datos"
   ClientHeight    =   2700
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   2880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   2880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCatalogo 
      Caption         =   "&Catalogos ..."
      Height          =   372
      Left            =   360
      TabIndex        =   5
      Top             =   972
      Width           =   2052
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   372
      Left            =   1452
      TabIndex        =   4
      Top             =   2112
      Width           =   1068
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   372
      Left            =   240
      TabIndex        =   3
      Top             =   2112
      Width           =   1068
   End
   Begin VB.CommandButton cmdCodTrans 
      Caption         =   "&Transacciones ..."
      Height          =   372
      Left            =   372
      TabIndex        =   2
      Top             =   1452
      Width           =   2052
   End
   Begin VB.CheckBox chkContable 
      Caption         =   "Ignorar el aspecto contable "
      Height          =   192
      Left            =   252
      TabIndex        =   1
      Top             =   600
      Width           =   2400
   End
   Begin VB.CheckBox chkLimitarFecha 
      Caption         =   "Limitar rango de fecha/hora"
      Height          =   192
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Value           =   1  'Checked
      Width           =   2400
   End
End
Attribute VB_Name = "frmExportarOpcion2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mAceptado As Boolean


Public Function Inicio( _
                    ByVal bandExportar As Boolean, _
                    ByRef LimitarFecha As Boolean, _
                    ByRef IgnorarContable As Boolean, _
                    ByRef Catalogo As String, _
                    ByRef CodTrans As String) As Boolean
    On Error GoTo errtrap
    
    'Cuando es de Exportación
    If bandExportar Then
        Me.Caption = "Opciones de Exportación"
        chkLimitarFecha.value = IIf(LimitarFecha, vbChecked, vbUnchecked)
        cmdCatalogo.tag = Catalogo
        cmdCodTrans.tag = CodTrans
    'Cuando es de Importación
    Else
        Me.Caption = "Opciones de Importación"
        chkLimitarFecha.value = vbUnchecked
        chkLimitarFecha.Enabled = False
        cmdCatalogo.tag = Catalogo
        cmdCodTrans.tag = CodTrans
    End If
    
    chkContable.value = IIf(IgnorarContable, vbChecked, vbUnchecked)
        
    mAceptado = False
    Me.Show vbModal
    
    If mAceptado Then
        If chkLimitarFecha.Enabled Then LimitarFecha = (chkLimitarFecha.value = vbChecked)
        If chkContable.Enabled Then IgnorarContable = (chkContable.value = vbChecked)
        If cmdCatalogo.Enabled Then Catalogo = cmdCatalogo.tag
        If cmdCodTrans.Enabled Then CodTrans = cmdCodTrans.tag
    End If
    
    Inicio = True
    Unload Me
    Exit Function
errtrap:
    DispErr
    Exit Function
End Function


Private Sub cmdAceptar_Click()
    mAceptado = True
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    mAceptado = False
    Me.Hide
End Sub

Private Sub cmdCatalogo_Click()
    Dim s As String
    'Abre la pantalla de búsqueda para seleccionar Códigos de transacciones
    s = cmdCatalogo.tag
    If frmTrans.SeleccionarCat(s) Then
        cmdCatalogo.tag = s
    End If
End Sub

Private Sub cmdCodTrans_Click()
    Dim s As String
    'Abre la pantalla de búsqueda para seleccionar Códigos de transacciones
    s = cmdCodTrans.tag
    If frmTrans.Seleccionar(s) Then
        cmdCodTrans.tag = s
    End If
End Sub

