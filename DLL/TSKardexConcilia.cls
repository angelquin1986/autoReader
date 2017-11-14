VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TSKardexConcilia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Private Type T_PROP
    Id As Long
    IdBanco As Long
    CodBanco As String
    IdTSKardex As Long
    Debe As Currency
    Haber As Currency
    Saldo As Currency
    SaldoConfirmado As Currency
    Diferencia As Currency

    FechaGrabado As Date
    BandConciliado As Boolean
    Orden As Integer
    objGNComprobante As GNComprobante           'Referencia al objeto padre
End Type

Private mProps As T_PROP


Public Sub VerificaDatos()
    'Verifica datos
    If mProps.Debe = 0 And mProps.Haber = 0 Then
        Err.Raise ERR_INVALIDO, "TSKardex_VerificaDatos", _
            "Se intentó grabar sin valor de Debe ni de Haber en detalle de Bancos."
    End If
'''    If mProps.Debe <> 0 And mProps.Haber <> 0 Then
'''        Err.Raise ERR_INVALIDO, "TSKardex_VerificaDatos", _
'''            "Se intentó grabar con Debe y Haber al mismo tiempo en detalle de Bancos."
'''    End If
    
    If mProps.IdBanco = 0 Then
        Err.Raise ERR_INVALIDO, "TSKardexConcilia_VerificaDatos", _
            "Debe seleccionar un banco."
    End If
'    If mProps.IdTSKardex = 0 Then
'        Err.Raise ERR_INVALIDO, "TSKardexConcilia_VerificaDatos", _
'            "Debe seleccionar un Detalle del documento de Bancos."
'    End If
End Sub

Friend Sub Grabar()
    Dim sql As String, rs As Recordset
    
    VerificaDatos
   
    sql = "SELECT * FROM TSKardexConcilia WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !TransID = mProps.objGNComprobante.TransID
        !IdBanco = mProps.IdBanco
        !Debe = mProps.Debe
        !Haber = mProps.Haber
        !Saldo = mProps.Saldo
        !SaldoConfirmado = mProps.SaldoConfirmado
        !Diferencia = mProps.Diferencia
        !IdTSKardex = mProps.IdTSKardex
        !FechaGrabado = mProps.FechaGrabado
        !BandConciliado = mProps.BandConciliado
        !Orden = mProps.Orden
        .Update
        
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.Id = !Id             'Guarda el Id(Autonumerico)
        .Close
    End With

    'If mProps.IdTSKardex = 507 Then MsgBox "Hola"
'    If mProps.BandConciliado Then
        sql = "update tskardex set BandConciliado=" & IIf(mProps.BandConciliado, 1, 0)
        sql = sql & " where id=" & mProps.IdTSKardex
        Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
'    End If
    'cambia estado en tskardex
    Set rs = Nothing
End Sub

Friend Sub Recuperar(rs As Recordset)
    With rs
        If Not .EOF Then
            mProps.Id = !Id
            If Not IsNull(!IdBanco) Then mProps.IdBanco = !IdBanco
'            If Not IsNull(!CodBanco) Then mProps.CodBanco = !CodBanco
            If Not IsNull(!Debe) Then mProps.Debe = !Debe
            If Not IsNull(!Haber) Then mProps.Haber = !Haber
            If Not IsNull(!Saldo) Then mProps.Saldo = !Saldo
            If Not IsNull(!SaldoConfirmado) Then mProps.SaldoConfirmado = !SaldoConfirmado
            If Not IsNull(!Diferencia) Then mProps.Diferencia = !Diferencia
            If Not IsNull(!IdTSKardex) Then mProps.IdTSKardex = !IdTSKardex
            If Not IsNull(!FechaGrabado) Then mProps.FechaGrabado = !FechaGrabado
            If Not IsNull(!BandConciliado) Then mProps.BandConciliado = !BandConciliado
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
        End If
    End With
End Sub

Friend Function Clone() As TSKardexConcilia
    Dim nuevo As TSKardexConcilia
    
    Set nuevo = New TSKardexConcilia
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROP)          'Para el método Clone
    mProps = value
    mProps.Id = 0
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Set GNComprobante(value As GNComprobante)
    Set mProps.objGNComprobante = value
End Property

Public Property Get GNComprobante() As GNComprobante
    Set GNComprobante = mProps.objGNComprobante
End Property



Public Property Get Id() As Long
    Id = mProps.Id
End Property

Friend Property Get IdBanco() As Long
    IdBanco = mProps.IdBanco
End Property


Public Property Let CodBanco(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodBanco Then
        If Len(value) > 0 Then
            'Actualiza IdBanco también
            sql = "SELECT IdBanco,IdCuentaContable FROM TSBanco WHERE CodBanco='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "TSKardex_CodBanco", MSGERR_NOHAYCODIGO
            Else
                mProps.IdBanco = rs!IdBanco
            End If
            rs.Close
        Else
            mProps.IdBanco = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodBanco = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodBanco() As String
    CodBanco = Trim$(mProps.CodBanco)
End Property

Public Property Let IdTSKardex(ByVal value As Long)
    mProps.IdTSKardex = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get IdTSKardex() As Long
    IdTSKardex = mProps.IdTSKardex
End Property



Public Property Let FechaGrabado(ByVal value As Date)
    mProps.FechaGrabado = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get FechaGrabado() As Date
    FechaGrabado = mProps.FechaGrabado
End Property

Public Property Let BandConciliado(ByVal value As Boolean)
    mProps.BandConciliado = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get BandConciliado() As Boolean
    BandConciliado = mProps.BandConciliado
End Property


Public Property Let Orden(ByVal value As Integer)
    mProps.Orden = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property

Private Sub Class_Terminate()
    Set mProps.objGNComprobante = Nothing
End Sub


'*** oliver 29/may/2001


Public Function Verificar_ExisteDocumento(CodBanco As String, CodTipoDoc As String, NumDoc As String) As Boolean
    Dim rs As Recordset, sql As String, numReg As Long
    On Error GoTo ErrTrap
    sql = "SELECT Count(TransID) as Cuantos FROM " & _
           "(TSKardex INNER JOIN TSBanco ON TSKardex.IdBanco = TSBanco.IdBanco) " & _
           "INNER JOIN TSTipoDocBanco ON TSKardex.IdTsKardex = TSTipoDocBanco.IdTsKardex " & _
           "WHERE TSBanco.CodBanco = '" & CodBanco & "' AND " & _
           "TSTipoDocBanco.CodTipoDoc = '" & CodTipoDoc & "' AND " & _
           "TSKARDEX.NumDoc='" & NumDoc & "'"
    'Debug.Print sql
    Set rs = Me.GNComprobante.Empresa.OpenRecordset(sql)
        Verificar_ExisteDocumento = IIf(rs!Cuantos = 0, False, True)
    Exit Function
ErrTrap:
        MsgBox Err.Number & " " & Err.Description, vbInformation
        Verificar_ExisteDocumento = False
End Function

Public Property Let Debe(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "TSKardex_Debe", MSGERR_INVALIDO
    ElseIf value > 0 Then
        mProps.Debe = value
        mProps.Haber = 0
    Else
        mProps.Debe = 0
    End If
    
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Debe() As Currency
    'Devuelve en Moneda ex.
    Debe = mProps.Debe / mProps.objGNComprobante.Cotizacion("")
End Property

Public Property Let Haber(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "TSKardex_Haber", MSGERR_INVALIDO
    ElseIf value > 0 Then
        mProps.Haber = value
'        mProps.Debe = 0
    Else
        mProps.Haber = 0
    End If
    
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Haber() As Currency
    'Devuelve en Moneda ex.
    Haber = mProps.Haber / mProps.objGNComprobante.Cotizacion("")
End Property


Public Property Let Saldo(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    mProps.Saldo = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Saldo() As Currency
    'Devuelve en Moneda ex.
    Saldo = mProps.Saldo / mProps.objGNComprobante.Cotizacion("")
End Property


Public Property Let SaldoConfirmado(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    mProps.SaldoConfirmado = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get SaldoConfirmado() As Currency
    'Devuelve en Moneda ex.
    SaldoConfirmado = mProps.SaldoConfirmado / mProps.objGNComprobante.Cotizacion("")
End Property

Public Property Let Diferencia(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    mProps.Diferencia = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Diferencia() As Currency
    'Devuelve en Moneda ex.
    Diferencia = mProps.Diferencia / mProps.objGNComprobante.Cotizacion("")
End Property

