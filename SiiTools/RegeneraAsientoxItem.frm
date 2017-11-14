VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{C4EBE568-AA77-11D3-8306-000021C5085D}#5.3#0"; "FlexCombo.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRegeneraAsientoxItem 
   Caption         =   "Regeneración de asientos"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6585
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   6585
   WindowState     =   2  'Maximized
   Begin VB.Frame fraGrupos 
      Caption         =   "Rango de Grupos"
      Height          =   915
      Left            =   6240
      TabIndex        =   22
      Top             =   180
      Visible         =   0   'False
      Width           =   5052
      Begin VB.ComboBox cboGrupo 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   480
         Width           =   1452
      End
      Begin FlexComboProy.FlexCombo fcbGrupoHasta 
         Height          =   300
         Left            =   3360
         TabIndex        =   24
         Top             =   480
         Width           =   1572
         _ExtentX        =   2778
         _ExtentY        =   529
         ColWidth2       =   1200
         ColWidth3       =   1200
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
      Begin FlexComboProy.FlexCombo fcbGrupoDesde 
         Height          =   300
         Left            =   1812
         TabIndex        =   25
         Top             =   480
         Width           =   1452
         _ExtentX        =   2566
         _ExtentY        =   529
         ColWidth2       =   1200
         ColWidth3       =   1200
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Hasta"
         Height          =   192
         Left            =   3360
         TabIndex        =   28
         Top             =   240
         Width           =   432
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Desde"
         Height          =   192
         Left            =   1800
         TabIndex        =   27
         Top             =   240
         Width           =   492
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Grupo"
         Height          =   192
         Left            =   240
         TabIndex        =   26
         Top             =   240
         Width           =   444
      End
   End
   Begin VB.Frame fraitem 
      Caption         =   "Items"
      Height          =   675
      Left            =   6240
      TabIndex        =   17
      Top             =   1080
      Width           =   5052
      Begin FlexComboProy.FlexCombo fcbDesde2 
         Height          =   315
         Left            =   840
         TabIndex        =   18
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         ColWidth2       =   1200
         ColWidth3       =   1200
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
      Begin FlexComboProy.FlexCombo fcbHasta2 
         Height          =   315
         Left            =   3225
         TabIndex        =   19
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         ColWidth2       =   1200
         ColWidth3       =   1200
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
      Begin VB.Label Label6 
         Caption         =   "Desde"
         Height          =   252
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   612
      End
      Begin VB.Label Label7 
         Caption         =   "Hasta"
         Height          =   252
         Left            =   2760
         TabIndex        =   20
         Top             =   240
         Width           =   612
      End
   End
   Begin VB.CheckBox chkTodo 
      Caption         =   "&Regenerar todo sin verificar"
      Enabled         =   0   'False
      Height          =   192
      Left            =   3120
      TabIndex        =   15
      Top             =   1440
      Width           =   2955
   End
   Begin VB.PictureBox pic1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   852
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   6585
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3855
      Width           =   6585
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Proceder"
         Enabled         =   0   'False
         Height          =   372
         Left            =   1728
         TabIndex        =   13
         Top             =   0
         Width           =   1212
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   372
         Left            =   4968
         TabIndex        =   12
         Top             =   0
         Width           =   1212
      End
      Begin VB.CommandButton cmdVerificar 
         Caption         =   "&Verificar"
         Enabled         =   0   'False
         Height          =   372
         Left            =   288
         TabIndex        =   11
         Top             =   0
         Width           =   1212
      End
      Begin MSComctlLib.ProgressBar prg1 
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   540
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid grd 
      Height          =   1932
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   6372
      _cx             =   11239
      _cy             =   3408
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
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   100
      ColWidthMax     =   4000
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
      AllowUserFreezing=   3
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   372
      Left            =   1704
      TabIndex        =   8
      Top             =   1320
      Width           =   1212
   End
   Begin VB.Frame fraFecha 
      Caption         =   "&Fecha (desde - hasta)"
      Height          =   1092
      Left            =   402
      TabIndex        =   0
      Top             =   120
      Width           =   1932
      Begin MSComCtl2.DTPicker dtpFecha1 
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1692
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   41222145
         CurrentDate     =   36348
      End
      Begin MSComCtl2.DTPicker dtpFecha2 
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1692
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   41222145
         CurrentDate     =   36348
      End
   End
   Begin VB.Frame fraCodTrans 
      Caption         =   "Cod.&Trans."
      Height          =   1092
      Left            =   2322
      TabIndex        =   3
      Top             =   120
      Width           =   1932
      Begin VB.CheckBox chkNoAprobadas 
         Caption         =   "Solo no aprobados"
         Height          =   255
         Left            =   150
         TabIndex        =   16
         Top             =   780
         Width           =   1665
      End
      Begin FlexComboProy.FlexCombo fcbTrans 
         Height          =   345
         Left            =   165
         TabIndex        =   4
         Top             =   360
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   609
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
   End
   Begin VB.Frame fraNumTrans 
      Caption         =   "# T&rans. (desde - hasta)"
      Height          =   1092
      Left            =   4242
      TabIndex        =   5
      Top             =   120
      Width           =   1932
      Begin VB.TextBox txtNumTrans1 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   360
         TabIndex        =   6
         Top             =   280
         Width           =   1212
      End
      Begin VB.TextBox txtNumTrans2 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   360
         TabIndex        =   7
         Top             =   640
         Width           =   1212
      End
   End
End
Attribute VB_Name = "frmRegeneraAsientoxItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Constantes para las columnas
Private Const COL_NUMFILA = 0
Private Const COL_TID = 1
Private Const COL_FECHA = 2
Private Const COL_CODASIENTO = 3
Private Const COL_CODTRANS = 4
Private Const COL_NUMTRANS = 5
Private Const COL_NUMDOCREF = 6     '*** MAKOTO 07/feb/01 Agregado
Private Const COL_NOMBRE = 7        '*** MAKOTO 07/feb/01 Agregado
Private Const COL_DESC = 8
Private Const COL_CENTROCOSTO = 9
Private Const COL_ESTADO = 10
Private Const COL_RESULTADO = 11

Private Const MSG_NG = "Asiento incorrecto."
Private mProcesando As Boolean
Private mCancelado As Boolean
Private mVerificado As Boolean
Const IVGRUPO_MAX = 5
Dim numGrupo As Integer
Private mobjGNCompAux As GNComprobante

Public Sub Inicio()
    Dim i As Integer
    On Error GoTo ErrTrap
    
    Me.Show
    Me.ZOrder
    dtpFecha1.value = gobjMain.EmpresaActual.GNOpcion.FechaInicio
    dtpFecha2.value = Date
    
    For i = 1 To IVGRUPO_MAX
        cboGrupo.AddItem gobjMain.EmpresaActual.GNOpcion.EtiqGrupo(i)
    Next i
    If (numGrupo <= cboGrupo.ListCount) And (numGrupo > 0) Then
        cboGrupo.ListIndex = numGrupo - 1   'Selecciona lo anterior
    ElseIf cboGrupo.ListCount > 0 Then
        cboGrupo.ListIndex = 0              'Selecciona la primera
    End If
    
    
    CargaTrans
    Exit Sub
ErrTrap:
    DispErr
    Unload Me
    Exit Sub
End Sub

Private Sub CargaTrans()
    'Carga la lista de transacción
    fcbTrans.SetData gobjMain.GrupoActual.PermisoActual.ListaTrans(False)
End Sub



Private Sub cmdAceptar_Click()
    'Si no hay transacciones
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna transacción para procesar."
        Exit Sub
    End If
    
    If dtpFecha1 < gobjMain.EmpresaActual.GNOpcion.FechaLimiteDesde Then
        MsgBox "La Rango de Fecha de regeneración es menor a la Fecha Limite Aceptable  ", vbExclamation
        Exit Sub
    End If

    
    If RegenerarAsiento(False, (chkTodo.value = vbChecked)) Then
        cmdCancelar.SetFocus
    End If
End Sub

Private Function RegenerarAsiento(bandVerificar As Boolean, BandTodo As Boolean) As Boolean
    Dim s As String, tid As Long, i As Long, x As Single
    Dim gnc As GNComprobante, cambiado As Boolean, TransID  As Long
    
    On Error GoTo ErrTrap

    'Si no es solo verificacion, confirma
    If Not bandVerificar Then
        s = "Este proceso modificará los asientos de la transacción seleccionada." & vbCr & vbCr
        s = s & "Está seguro que desea proceder?"
        If MsgBox(s, vbYesNo + vbQuestion) <> vbYes Then Exit Function
    End If
    
    mProcesando = True
    mCancelado = False
    frmMain.mnuFile.Enabled = False
    cmdVerificar.Enabled = False
    cmdBuscar.Enabled = False
    Screen.MousePointer = vbHourglass
    prg1.min = 0
    prg1.max = grd.Rows - 1
    
    For i = grd.FixedRows To grd.Rows - 1
        DoEvents
        If mCancelado Then
            MsgBox "El proceso fue cancelado.", vbInformation
            Exit For
        End If
        
        prg1.value = i
        grd.Row = i
        x = grd.CellTop                 'Para visualizar la celda actual
        
        'Si es verificación, procesa todas las filas sino solo las que tengan "Asiento incorrecto."
        If (grd.TextMatrix(i, COL_RESULTADO) = MSG_NG) Or bandVerificar Or BandTodo Then
        
            tid = grd.ValueMatrix(i, COL_TID)
            grd.TextMatrix(i, COL_RESULTADO) = "Verificando..."
            grd.Refresh
            
            'Recupera la transaccion
            Set gnc = gobjMain.EmpresaActual.RecuperaGNComprobante(tid)
            If Not (gnc Is Nothing) Then
                'Si la transacción no está anulada
                If gnc.Estado <> ESTADO_ANULADO Then
                
                    'Forzar recuperar todos los datos de transacción para que no se pierdan al grabar de nuveo
                    gnc.RecuperaDetalleTodo
                
                    'Recalcula costo de los items
                    If RegenerarAsientoSub(gnc, cambiado) Then
                        'Si está cambiado algo o está forzado regenerar todo
                        If cambiado Or BandTodo Then
                            'Si no es solo verificacion
                            If (Not bandVerificar) Or BandTodo Then
                                grd.TextMatrix(i, COL_RESULTADO) = "Grabando..."
                                grd.Refresh
                                
                                'Graba la transacción
                                gnc.Grabar False, False
                                
                                If gnc.GNTrans.IVAutoImpresor And Len(gnc.GNTrans.AsientoTrans) > 0 Then
                                    If gnc.FechaTrans > "01/07/2011" Then 'fecha del cambio por el sri
                                        'verifca si hay transaccion de asiento
                                        If gnc.VerificasiHayAsientoAutoImpresor(gnc.TransID, gnc.GNTrans.AsientoTrans) Then
                                            GrabarTransAutoNew gnc.GNTrans.AsientoTrans, gnc
                                        Else
                                            ModificaTransAsiento TransID, gnc
                                        End If
                                    End If
                                End If
                                
                                grd.TextMatrix(i, COL_RESULTADO) = "Actualizado."
                                
                            'Si es solo verificacion
                            Else
                                grd.TextMatrix(i, COL_RESULTADO) = MSG_NG
                            End If
                        Else
                            'Si no está cambiado no graba
                            grd.TextMatrix(i, COL_RESULTADO) = "OK."
                        End If
                    Else
                        grd.TextMatrix(i, COL_RESULTADO) = "Falló al regenerar."
                    End If
                Else
                    'Si está anulada
                    grd.TextMatrix(i, COL_RESULTADO) = "Anulado."
                End If
            Else
                grd.TextMatrix(i, COL_RESULTADO) = "No pudo recuperar la transación."
            End If
        End If
    Next i
    
    Screen.MousePointer = 0
    RegenerarAsiento = Not mCancelado
    GoTo salida
ErrTrap:
    Screen.MousePointer = 0
    DispErr
salida:
    mProcesando = False
    frmMain.mnuFile.Enabled = True
    cmdVerificar.Enabled = True
    cmdBuscar.Enabled = True
    prg1.value = prg1.min
    Exit Function
End Function


Private Function RegenerarAsientoSub(ByVal gnc As GNComprobante, _
                                     ByRef cambiado As Boolean) As Boolean
    Dim i As Long, cta As CtCuenta, ctd As CTLibroDetalle
    Dim colCtd As Collection, a As clsAsiento
    On Error GoTo ErrTrap
    
    cambiado = False
    Set colCtd = New Collection
    
    'Guarda todos los detalles de asiento en la colección para después comparar
    With gnc
        For i = 1 To .CountCTLibroDetalle
            Set ctd = .CTLibroDetalle(i)
            Set a = New clsAsiento
            a.IdCuenta = ctd.IdCuenta
            a.Debe = ctd.Debe
            a.Haber = ctd.Haber
            colCtd.Add item:=a
        Next i
    End With
    
    'Regenera el asiento
    gnc.GeneraAsiento
    
    'Compara el asiento para saber si ha cambiado o no
    cambiado = Not CompararAsiento(gnc, colCtd)
    
    RegenerarAsientoSub = True
    GoTo salida
    Exit Function
ErrTrap:
    cambiado = False
    DispErr
    RegenerarAsientoSub = False
salida:
    Set a = Nothing
    Set colCtd = Nothing
    Set gnc = Nothing
    Exit Function
End Function


'Devuelve True si los asientos son iguales, False si no lo son
Private Function CompararAsiento(ByVal gnc As GNComprobante, ByVal col As Collection) As Boolean
    Dim a As clsAsiento, i As Long, ctd As CTLibroDetalle
    Dim encontrado As Boolean
    
    'Si número de detalles son diferentes ya no son iguales
    If col.Count <> gnc.CountCTLibroDetalle Then Exit Function
    
    For i = 1 To gnc.CountCTLibroDetalle
        Set ctd = gnc.CTLibroDetalle(i)
        encontrado = False
        For Each a In col
            If (ctd.IdCuenta = a.IdCuenta) And _
               (ctd.Debe = a.Debe) And _
               (ctd.Haber = a.Haber) And _
               (a.Comparado = False) Then
                a.Comparado = True
                encontrado = True
                Exit For
            End If
        Next a
        'Si no se encuentra uno igual
        If Not encontrado Then
            CompararAsiento = False
            Exit Function
        End If
    Next i
    CompararAsiento = True
End Function

Private Sub cmdBuscar_Click()
    Dim v As Variant, obj As Object
    On Error GoTo ErrTrap
    
    If Len(fcbDesde2.Text) = 0 Then
        MsgBox "Seleccione un Item, por favor.", vbInformation
        fcbDesde2.SetFocus
        Exit Sub
    End If

    
    With gobjMain.objCondicion
        .fecha1 = dtpFecha1.value
        .fecha2 = dtpFecha2.value
        .CodTrans = fcbTrans.Text
        .NumTrans1 = Val(txtNumTrans1.Text)
        .NumTrans2 = Val(txtNumTrans2.Text)
        .CodItem1 = Trim$(fcbDesde2.Text)
        .CodItem2 = Trim$(fcbHasta2.Text)
        
        
        'Estados no incluye anulados
        If chkNoAprobadas.value = vbChecked Then
            .EstadoBool(ESTADO_NOAPROBADO) = True
            .EstadoBool(ESTADO_APROBADO) = False
            .EstadoBool(ESTADO_DESPACHADO) = False
            .EstadoBool(ESTADO_ANULADO) = False
        Else
            .EstadoBool(ESTADO_NOAPROBADO) = True
            .EstadoBool(ESTADO_APROBADO) = True
            .EstadoBool(ESTADO_DESPACHADO) = True
            .EstadoBool(ESTADO_ANULADO) = False
        End If
    End With
    Set obj = gobjMain.EmpresaActual.ConsGNTrans22(True) 'Ascendente     '*** MAKOTO 20/oct/00
    If Not obj.EOF Then
        v = MiGetRows(obj)
        
        grd.Redraw = flexRDNone
        grd.LoadArray v
        ConfigCols
        grd.Redraw = flexRDDirect
    Else
        grd.Rows = grd.FixedRows
        ConfigCols
    End If
    cmdVerificar.Enabled = True
    cmdVerificar.SetFocus
    cmdAceptar.Enabled = False
    chkTodo.Enabled = True
    mVerificado = False
    Exit Sub
ErrTrap:
    DispErr
    Exit Sub
End Sub

Private Sub ConfigCols()
    With grd
        .FormatString = "^#|tid|<Fecha|<Asiento|<Trans|<#|<#Ref.|<Nombre|<Descripción|<C.Costo|<Estado|<Resultado"
        .ColHidden(COL_NUMFILA) = False
        .ColHidden(COL_TID) = True
        .ColHidden(COL_FECHA) = False
        .ColHidden(COL_CODASIENTO) = True
        .ColHidden(COL_CODTRANS) = False
        .ColHidden(COL_NUMTRANS) = False
        .ColHidden(COL_NUMDOCREF) = True
        .ColHidden(COL_NOMBRE) = False      'True
        .ColHidden(COL_DESC) = False
        .ColHidden(COL_CENTROCOSTO) = True
        .ColHidden(COL_ESTADO) = True
        
        .ColDataType(COL_FECHA) = flexDTDate    '*** MAKOTO 14/ago/2000 para que ordene bien por fecha
        
        GNPoneNumFila grd, False
        .AutoSize 0, grd.Cols - 1
        
        .ColWidth(COL_NUMTRANS) = 500
        .ColWidth(COL_NOMBRE) = 1400
        .ColWidth(COL_DESC) = 2400
        .ColWidth(COL_RESULTADO) = 2000
    End With
End Sub

Private Sub cmdCancelar_Click()
    If mProcesando Then
        mCancelado = True
    Else
        Unload Me
    End If
End Sub


Private Sub cmdVerificar_Click()
    'Si no hay transacciones
    If grd.Rows <= grd.FixedRows Then
        MsgBox "No hay ningúna transacción para verificar."
        Exit Sub
    End If
    
    If dtpFecha1 < gobjMain.EmpresaActual.GNOpcion.FechaLimiteDesde Then
        MsgBox "La Rango de Fecha de regeneración es menor a la Fecha Limite Aceptable  ", vbExclamation
        Exit Sub
    End If

    If RegenerarAsiento(True, False) Then
        cmdAceptar.Enabled = True
        cmdAceptar.SetFocus
        mVerificado = True
    End If
End Sub

Private Sub chkTodo_Click()
    If chkTodo.value = vbChecked Then
        cmdVerificar.Enabled = False
        cmdAceptar.Enabled = (grd.Rows > grd.FixedRows)
    Else
        cmdVerificar.Enabled = Not mVerificado
        cmdAceptar.Enabled = mVerificado
    End If
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
    If mProcesando Then
        Cancel = True
        Exit Sub
    End If
    Me.Hide         'Se pone esto para evitar el posible BUG de Windows98
End Sub



Private Sub Form_Resize()
    On Error Resume Next
    grd.Move 0, grd.Top, Me.ScaleWidth, Me.ScaleHeight - grd.Top - pic1.Height - 80
    prg1.Width = Me.ScaleWidth - (prg1.Left * 2)
End Sub


Private Sub txtNumTrans1_KeyPress(KeyAscii As Integer)
    'Acepta solo numericos
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNumTrans2_KeyPress(KeyAscii As Integer)
    'Acepta solo numericos
    If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End Sub

Private Sub fcbDesde2_Selected(ByVal Text As String, ByVal KeyText As String)
    fcbHasta2.KeyText = fcbDesde2.KeyText   '*** MAKOTO 27/jun/2000
End Sub

Private Sub CargaItems()
    Dim numGrupo As Integer, v() As Variant
    Dim sql  As String, rs As Recordset, cond As String
    numGrupo = cboGrupo.ListIndex + 1
    fcbDesde2.Clear
    fcbHasta2.Clear
    If Len(fcbGrupoDesde.Text) > 0 And Len(fcbGrupoHasta.Text) > 0 Then
        cond = " WHERE codGrupo" & numGrupo & " BETWEEN '" & _
                fcbGrupoDesde.Text & "' AND '" & fcbGrupoHasta.Text & "'"
    End If
    sql = "SELECT CodInventario, IVInventario.Descripcion FROM IVInventario " & _
    IIf(Len(fcbGrupoDesde.Text) > 0 And Len(fcbGrupoHasta.Text) > 0, " INNER JOIN IVGrupo" & numGrupo & _
           " ON IVInventario.IdGrupo" & numGrupo & " = IVGrupo" & numGrupo & ".IdGrupo" & numGrupo & cond, "")
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not rs.EOF Then
        v = MiGetRows(rs)
        fcbDesde2.SetData v
        fcbHasta2.SetData v
    End If
    fcbDesde2.Text = ""
    fcbHasta2.Text = ""
End Sub


Private Sub fcbGrupoDesde_Validate(Cancel As Boolean)
    'Carga Items
    CargaItems

End Sub



Private Sub fcbGrupoHasta_Validate(Cancel As Boolean)
    'Carga Items
    CargaItems

End Sub

Private Sub cboGrupo_Click()
    Dim Numg As Integer
    On Error GoTo ErrTrap
    If cboGrupo.ListIndex < 0 Then Exit Sub

    'MensajeStatus MSG_PREPARA, vbHourglass

    Numg = cboGrupo.ListIndex + 1
    fcbGrupoDesde.SetData gobjMain.EmpresaActual.ListaIVGrupo(Numg, False, False)
    fcbGrupoHasta.SetData fcbGrupoDesde.GetData             '*** MAKOTO 19/feb/01 Mod.
    fcbGrupoDesde.KeyText = ""
    fcbGrupoHasta.KeyText = ""
    CargaItems
    'MensajeStatus
    Exit Sub
ErrTrap:
    MensajeStatus
    DispErr
    Exit Sub
End Sub


Private Function GrabarTransAutoNew(ByVal CodTrans As String, ByRef mobjGNComp As GNComprobante) As Boolean
    Dim Imprime As Boolean, i As Long, ix As Long, j As Integer
    Dim item As IVinventario, rsReceta As Recordset
    Dim Cadena As String, aux_inc As Variant

    On Error GoTo ErrTrap
    Set mobjGNCompAux = gobjMain.EmpresaActual.CreaGNComprobanteAutoimpresor(CodTrans)
    
    If Not mobjGNCompAux Is Nothing Then
    
        If mobjGNCompAux.SoloVer Then
            MsgBox MSG_NODISPONE, vbInformation
            Exit Function
        End If
        
        If mobjGNComp.CountCTLibroDetalle > 0 Then
            For i = 1 To mobjGNComp.CountCTLibroDetalle
                ix = mobjGNCompAux.AddCTLibroDetalle
                mobjGNCompAux.CTLibroDetalle(ix).BandIntegridad = mobjGNComp.CTLibroDetalle(ix).BandIntegridad
                mobjGNCompAux.CTLibroDetalle(ix).codcuenta = mobjGNComp.CTLibroDetalle(ix).codcuenta
                mobjGNCompAux.CTLibroDetalle(ix).CodGasto = mobjGNComp.CTLibroDetalle(ix).CodGasto
                mobjGNCompAux.CTLibroDetalle(ix).Debe = mobjGNComp.CTLibroDetalle(ix).Debe
                mobjGNCompAux.CTLibroDetalle(ix).Descripcion = mobjGNComp.CTLibroDetalle(ix).Descripcion
                mobjGNCompAux.CTLibroDetalle(ix).Haber = mobjGNComp.CTLibroDetalle(ix).Haber
                mobjGNCompAux.CTLibroDetalle(ix).orden = mobjGNComp.CTLibroDetalle(ix).orden
                
            Next i
        End If
       
        mobjGNCompAux.FechaTrans = mobjGNComp.FechaTrans
        mobjGNCompAux.HoraTrans = mobjGNComp.HoraTrans
        Cadena = "Por transaccion FACTURA " & mobjGNComp.CodTrans & "-" & mobjGNComp.numtrans & " / " & mobjGNComp.NumSerieEstaSRI & "-" & mobjGNComp.NumSeriePuntoSRI & "-" & Right("000000000" + Trim(Str(mobjGNComp.numtrans)), 9)    '& mobjGNComp.codtrans & "-" & mobjGNComp.NumTrans
        If Len(Cadena) > 120 Then
            mobjGNCompAux.Descripcion = Mid$(Cadena, 1, 120)
        Else
            mobjGNCompAux.Descripcion = Cadena
        End If
            
        mobjGNCompAux.codUsuario = mobjGNComp.codUsuario
        mobjGNCompAux.IdResponsable = mobjGNComp.IdResponsable
        mobjGNCompAux.numDocRef = mobjGNComp.NumSerieEstaSRI & "-" & mobjGNComp.NumSeriePuntoSRI & "-" & Right("000000000" + Trim(Str(mobjGNComp.numtrans)), 9)
        mobjGNCompAux.idCentro = mobjGNComp.idCentro
        mobjGNCompAux.IdTransFuente = mobjGNComp.Empresa.RecuperarTransIDGncomprobante(mobjGNComp.CodTrans, mobjGNComp.numtrans)
        mobjGNCompAux.CodMoneda = mobjGNComp.CodMoneda

    
    
'        If GNTrans.ImportaCTD Then
'            mobjGNCompAux.ImportaAsiento mobjGNComp, aux_inc
 '       End If
    
    
    
        'Si es que algo está modificado
        If mobjGNCompAux.Modificado Then
            MensajeStatus MSG_GENERANDOASIENTO, vbHourglass
'            PreparaAsientoAutoNew True
            MensajeStatus
        End If
        'Verificación de datos
        mobjGNCompAux.VerificaDatos
    
'        PreparaAsientoAuto True
        'Verifica si está cuadrado el asiento
        If Not VerificaAsiento(mobjGNCompAux) Then Exit Function
    

        MensajeStatus MSG_GRABANDO, vbHourglass
    
        'Manda a grabar
        '       Aquí ya no hacemos verificación de asiento por que ya está hecho en Control Asiento
        mobjGNCompAux.Grabar False, False

        '***  Oliver 26/12/2002
        'Agregado para el control ded Impresion Configurado en la Transaccion

        MensajeStatus
        GrabarTransAutoNew = True
    Else
        GrabarTransAutoNew = False
    End If
    Exit Function
ErrTrap:
    MensajeStatus
    Select Case Err.Number
    Case ERR_DESCUADRADO, ERR_INTEGRIDAD
        'Si es que el usuario seleccionó 'No' en el cuadro de dialogo,
        'No hace nada
    Case Else
        DispErr
    End Select
    Exit Function

End Function

Private Sub PreparaAsientoTransAuto(Aceptar As Boolean)
    mobjGNCompAux.GeneraAsiento
End Sub


Private Function ModificaTransAsiento(ByVal TransID As Long, ByRef mobjGNComp As GNComprobante) As Boolean
    Dim Imprime As Boolean, i As Long, ix As Long, j As Integer
    Dim item As IVinventario, rsReceta As Recordset
    Dim Cadena As String, aux_inc As Variant

    On Error GoTo ErrTrap
    Set mobjGNCompAux = gobjMain.EmpresaActual.RecuperaGNComprobante(TransID)
    
    If Not mobjGNCompAux Is Nothing Then
    
        For i = 1 To mobjGNCompAux.CountCTLibroDetalle
            mobjGNCompAux.RemoveCTLibroDetalle 1
        Next i
    
        If mobjGNComp.CountCTLibroDetalle > 0 Then
            For i = 1 To mobjGNComp.CountCTLibroDetalle
                ix = mobjGNCompAux.AddCTLibroDetalle
                mobjGNCompAux.CTLibroDetalle(ix).BandIntegridad = mobjGNComp.CTLibroDetalle(ix).BandIntegridad
                mobjGNCompAux.CTLibroDetalle(ix).codcuenta = mobjGNComp.CTLibroDetalle(ix).codcuenta
                mobjGNCompAux.CTLibroDetalle(ix).CodGasto = mobjGNComp.CTLibroDetalle(ix).CodGasto
                mobjGNCompAux.CTLibroDetalle(ix).Debe = mobjGNComp.CTLibroDetalle(ix).Debe
                mobjGNCompAux.CTLibroDetalle(ix).Descripcion = mobjGNComp.CTLibroDetalle(ix).Descripcion
                mobjGNCompAux.CTLibroDetalle(ix).Haber = mobjGNComp.CTLibroDetalle(ix).Haber
                mobjGNCompAux.CTLibroDetalle(ix).orden = mobjGNComp.CTLibroDetalle(ix).orden

            Next i
        End If
    
        
        mobjGNCompAux.Grabar False, False

        '***  Oliver 26/12/2002
        'Agregado para el control ded Impresion Configurado en la Transaccion

        MensajeStatus
        ModificaTransAsiento = True
    Else
        ModificaTransAsiento = False
    End If
    Exit Function
ErrTrap:
    MensajeStatus
    Select Case Err.Number
    Case ERR_DESCUADRADO, ERR_INTEGRIDAD
        'Si es que el usuario seleccionó 'No' en el cuadro de dialogo,
        'No hace nada
    Case Else
        DispErr
    End Select
    Exit Function

End Function



