VERSION 5.00
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Begin VB.Form frmExportarOpcion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones de Exportación de datos"
   ClientHeight    =   2850
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkDocAsignado 
      Caption         =   "Ignorar Documentos Asignados"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Frame fraIV 
      Caption         =   "Modulo IV"
      Height          =   1095
      Left            =   2760
      TabIndex        =   7
      Top             =   120
      Width           =   2055
      Begin FlexComboProy.FlexCombo fcbBodega 
         Height          =   375
         Left            =   840
         TabIndex        =   9
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
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
      Begin VB.CheckBox chkFiltroBodega 
         Caption         =   "Filtrar por Bodega"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblBodega 
         Caption         =   "Bodega"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkSoloCatalogo 
      Caption         =   "A&ctualizar solo catalogos"
      Height          =   192
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   372
      Left            =   2694
      TabIndex        =   6
      Top             =   2400
      Width           =   1332
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   372
      Left            =   654
      TabIndex        =   5
      Top             =   2400
      Width           =   1332
   End
   Begin VB.CommandButton cmdCodTrans 
      Caption         =   "&Transacciones ..."
      Height          =   372
      Left            =   1314
      TabIndex        =   4
      Top             =   1800
      Width           =   2052
   End
   Begin VB.CheckBox chkContable 
      Caption         =   "Ignorar el aspecto contable "
      Height          =   192
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin VB.CheckBox chkLimitarFecha 
      Caption         =   "Limitar rango de fecha/hora"
      Height          =   192
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Value           =   1  'Checked
      Width           =   2895
   End
End
Attribute VB_Name = "frmExportarOpcion"
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
                    ByRef soloCatalogo As Boolean, _
                    ByRef IgnorarDocAsignado As Boolean, _
                    ByRef CodTrans As String) As Boolean
    On Error GoTo ErrTrap
    
    'Cuando es de Exportación
    If bandExportar Then
        Me.Caption = "Opciones de Exportación"
        chkLimitarFecha.value = IIf(LimitarFecha, vbChecked, vbUnchecked)
        cmdCodTrans.tag = CodTrans
        'Prepara la bodega
        FcbBodega.SetData gobjMain.EmpresaActual.ListaIVBodega(False, False)
        'fcbBodega.KeyText = objcond.Bodega

    'Cuando es de Importación
    Else
        Me.Caption = "Opciones de Importación"
        chkLimitarFecha.value = vbUnchecked
        chkLimitarFecha.Enabled = False
        cmdCodTrans.Enabled = False
        fraIV.Visible = False   'Diego
    End If
    chkContable.value = IIf(IgnorarContable, vbChecked, vbUnchecked)
    chkSoloCatalogo.value = IIf(soloCatalogo, vbChecked, vbUnchecked)
    chkDocAsignado.value = IIf(IgnorarDocAsignado, vbChecked, vbUnchecked) '***Angel. 13/nov/2003
    
    mAceptado = False
    Me.Show vbModal
    
    If mAceptado Then
        If chkLimitarFecha.Enabled Then LimitarFecha = (chkLimitarFecha.value = vbChecked)
        If chkContable.Enabled Then IgnorarContable = (chkContable.value = vbChecked)
        If chkSoloCatalogo.Enabled Then soloCatalogo = (chkSoloCatalogo.value = vbChecked)
        If cmdCodTrans.Enabled Then CodTrans = cmdCodTrans.tag
        If chkDocAsignado.Enabled Then IgnorarDocAsignado = (chkDocAsignado.value = vbChecked) '***Angel. 13/nov/2003
    End If
    
    Inicio = True
    Unload Me
    Exit Function
ErrTrap:
    DispErr
    Exit Function
End Function


Private Sub chkFiltroBodega_Click()
    If chkFiltroBodega.value = vbChecked Then
        FcbBodega.Enabled = True
        lblBodega.Enabled = True
    Else
        FcbBodega.Enabled = False
        lblBodega.Enabled = False
    End If
End Sub

Private Sub cmdAceptar_Click()
    mAceptado = True
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
    mAceptado = False
    Me.Hide
End Sub

Private Sub cmdCodTrans_Click()
    Dim s As String
    
    'Abre la pantalla de búsqueda para seleccionar Códigos de transacciones
    s = cmdCodTrans.tag
    If frmTrans.Seleccionar(s) Then
        cmdCodTrans.tag = s
    End If
End Sub

