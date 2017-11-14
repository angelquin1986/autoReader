VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GNRepoGasto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Private Type T_PROP
    id As Long
    IdBanco As Long
    CodBanco As String
    Debe As Currency
    Haber As Currency
    Nombre As String
    IdTipoDoc As Long
    CodTipoDoc As String
    NumDoc As String
    FechaEmision As Date
    FechaVenci As Date
    Observacion As String
    BandConciliado As Boolean
    Orden As Integer
    
    auxIdCuenta As Long          'Auxiliar para almacenar codigo de cuenta de banco. (No se graba en la base)
    auxBandCalculoBanco As Integer  'Auxiliar para tipo de documento
    auxBandCalculoDoc As Integer    'Auxiliar para tipo de documento
    
    IdBancoRef As Long
    CodBancoRef As String
    IdTarjeta As Long
    CodTarjeta As String
    NumCuenta  As String
    numCheque  As String
    TitularCta  As String
    
    objGNComprobante As gnComprobante           'Referencia al objeto padre
End Type

Private mProps As T_PROP


Public Sub VerificaDatos()
    'Verifica datos
    If mProps.Debe = 0 And mProps.Haber = 0 Then
        Err.Raise ERR_INVALIDO, "TSKardex_VerificaDatos", _
            "Se intentó grabar sin valor de Debe ni de Haber en detalle de Bancos."
    End If
    If mProps.Debe <> 0 And mProps.Haber <> 0 Then
        Err.Raise ERR_INVALIDO, "TSKardex_VerificaDatos", _
            "Se intentó grabar con Debe y Haber al mismo tiempo en detalle de Bancos."
    End If
    If mProps.IdBanco = 0 Then
        Err.Raise ERR_INVALIDO, "TSKardex_VerificaDatos", _
            "Debe seleccionar un banco."
    End If
    If mProps.IdTipoDoc = 0 Then
            Err.Raise ERR_INVALIDO, "TSKardex_VerificaDatos", _
            "Debe seleccionar un tipo de documento en detalle de Bancos."
    End If
    If Len(mProps.NumDoc) = 0 Then
        Err.Raise ERR_INVALIDO, "TSKardex_VerificaDatos", _
            "Debe ingresar un número de documento en detalle de Bancos."
    End If
    If mProps.FechaEmision > mProps.FechaVenci Then
        Err.Raise ERR_INVALIDO, "TSKardex_VerificaDatos", _
            "Fecha de vencimiento no puede ser menor a la fecha de emisión en detalle de Bancos."
    End If
    'jeaa 10/05/2005
    If mProps.objGNComprobante.GNTrans.VerificaFechas Then
        If mProps.objGNComprobante.FechaTrans <> mProps.FechaEmision Then
            Err.Raise ERR_INVALIDO, "TSKardex_VerificaDatos", _
                "Fecha del comprobante no puede ser diferente a la fecha de emisión en detalle de Bancos."
        End If
    End If
    
    
    
    
    
End Sub

Friend Sub grabar()
    Dim sql As String, rs As Recordset
    Dim IdTSKardex As Long
    
    IdTSKardex = mProps.id
    VerificaDatos
    
'    If mProps.objGNComprobante.EsNuevo Then
'        mProps.NumDoc = AsignaNumCheque(mProps.CodBanco)       'Asigna un numero cogiendo de NumTransSiguiente de GNTrans
'    End If
    
    sql = "SELECT * FROM TSKardex WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !transid = mProps.objGNComprobante.transid
        !IdBanco = mProps.IdBanco
        !Debe = mProps.Debe
        !Haber = mProps.Haber
        !Nombre = mProps.Nombre
        !IdTipoDoc = mProps.IdTipoDoc
        !NumDoc = mProps.NumDoc
        !FechaEmision = mProps.FechaEmision
        !FechaVenci = mProps.FechaVenci
        !Observacion = mProps.Observacion
        !BandConciliado = mProps.BandConciliado
        !Orden = mProps.Orden
        
        !IdBancoRef = mProps.IdBancoRef   'jeaa 20/07/2009
        !IdTarjeta = mProps.IdTarjeta
        !NumCuenta = mProps.NumCuenta
        !numCheque = mProps.numCheque
        !TitularCta = mProps.TitularCta
        
        
        
        .Update
        
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.id = !id             'Guarda el Id(Autonumerico)
        .Close
        If mProps.CodTipoDoc = "CH-E" Then
            GrabaNuevoNumCheque mProps.CodBanco, CDbl(mProps.NumDoc)
        End If
        'Modificar IdTsKardex en transaccion  de Conciliacion
        If IdTSKardex <> 0 Then
            ModificaNuevoId IdTSKardex
        End If
        
    End With
    Set rs = Nothing
End Sub

Friend Sub recuperar(rs As Recordset)
    With rs
        If Not .EOF Then
            mProps.id = !id
            If Not IsNull(!IdBanco) Then mProps.IdBanco = !IdBanco
            If Not IsNull(!CodBanco) Then mProps.CodBanco = !CodBanco
            If Not IsNull(!Nombre) Then mProps.Nombre = !Nombre
            If Not IsNull(!Debe) Then mProps.Debe = !Debe
            If Not IsNull(!Haber) Then mProps.Haber = !Haber
            If Not IsNull(!IdTipoDoc) Then mProps.IdTipoDoc = !IdTipoDoc
            If Not IsNull(!CodTipoDoc) Then mProps.CodTipoDoc = !CodTipoDoc
            If Not IsNull(!NumDoc) Then mProps.NumDoc = !NumDoc
            If Not IsNull(!FechaEmision) Then mProps.FechaEmision = !FechaEmision
            If Not IsNull(!FechaVenci) Then mProps.FechaVenci = !FechaVenci
            If Not IsNull(!Observacion) Then mProps.Observacion = !Observacion
            If Not IsNull(!BandConciliado) Then mProps.BandConciliado = !BandConciliado
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
                
            If Not IsNull(!IdCuentaContable) Then mProps.auxIdCuenta = !IdCuentaContable
            
            
            If Not IsNull(!IdBancoRef) Then mProps.IdBancoRef = !IdBancoRef   ' jeaa 20/07/2009
            If Not IsNull(!CodBancoRef) Then mProps.CodBancoRef = !CodBancoRef
            If Not IsNull(!IdTarjeta) Then mProps.IdTarjeta = !IdTarjeta
            If Not IsNull(!CodTarjeta) Then mProps.CodTarjeta = !CodTarjeta
            If Not IsNull(!NumCuenta) Then mProps.NumCuenta = !NumCuenta
            If Not IsNull(!numCheque) Then mProps.numCheque = !numCheque
            If Not IsNull(!TitularCta) Then mProps.TitularCta = !TitularCta
            

            
        End If
    End With
End Sub

Friend Function Clone() As TSKardex
    Dim nuevo As TSKardex
    
    Set nuevo = New TSKardex
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(Value As T_PROP)          'Para el método Clone
    mProps = Value
    mProps.id = 0
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Set gnComprobante(Value As gnComprobante)
    Set mProps.objGNComprobante = Value
End Property

Public Property Get gnComprobante() As gnComprobante
    Set gnComprobante = mProps.objGNComprobante
End Property



Public Property Get id() As Long
    id = mProps.id
End Property

Public Sub SetIdCuentaAux(auxIdCuenta As Long)
    mProps.auxIdCuenta = auxIdCuenta
End Sub

'Public Property Let IdBanco(ByVal value As Long)
'    mProps.IdBanco = value
'    mProps.objGNComprobante.SetModificado
'End Property

Friend Property Get IdBanco() As Long
    IdBanco = mProps.IdBanco
End Property


Public Property Let CodBanco(ByVal Value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If Value <> mProps.CodBanco Then
        If Len(Value) > 0 Then
            'Actualiza IdBanco también
            sql = "SELECT IdBanco,IdCuentaContable FROM TSBanco WHERE CodBanco='" & Value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "TSKardex_CodBanco", MSGERR_NOHAYCODIGO
            Else
                mProps.IdBanco = rs!IdBanco
                mProps.auxIdCuenta = rs!IdCuentaContable
            End If
            rs.Close
        Else
            mProps.IdBanco = 0
            mProps.auxIdCuenta = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodBanco = Value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodBanco() As String
    CodBanco = Trim$(mProps.CodBanco)
End Property



Public Property Let Debe(ByVal Value As Currency)
    'Convierte en sucres
    Value = Value * mProps.objGNComprobante.Cotizacion("")
    
    If Value < 0 Then
        Err.Raise ERR_INVALIDO, "TSKardex_Debe", MSGERR_INVALIDO
    ElseIf Value > 0 Then
        mProps.Debe = Value
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

Public Property Let Haber(ByVal Value As Currency)
    'Convierte en sucres
    Value = Value * mProps.objGNComprobante.Cotizacion("")
    
    If Value < 0 Then
        Err.Raise ERR_INVALIDO, "TSKardex_Haber", MSGERR_INVALIDO
    ElseIf Value > 0 Then
        mProps.Haber = Value
        mProps.Debe = 0
    Else
        mProps.Haber = 0
    End If
    
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Haber() As Currency
    'Devuelve en Moneda ex.
    Haber = mProps.Haber / mProps.objGNComprobante.Cotizacion("")
End Property


Public Property Let Nombre(ByVal Value As String)
    If Len(Value) > 40 Then
        Err.Raise ERR_INVALIDO, "TSKardex_Nombre", MSGERR_INVALIDO
    End If
    
    mProps.Nombre = Value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Nombre() As String
    Nombre = Trim$(mProps.Nombre)
End Property


Public Sub SetIdTipoDocAux(auxCalculoBanco As Integer, _
                            auxCalculoDoc As Integer)
    mProps.auxBandCalculoBanco = auxCalculoBanco
    mProps.auxBandCalculoDoc = auxCalculoDoc
End Sub

Friend Property Let IdTipoDoc(ByVal Value As Long)
    mProps.IdTipoDoc = Value
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Get IdTipoDoc() As Long
    IdTipoDoc = mProps.IdTipoDoc
End Property


Public Property Let CodTipoDoc(ByVal Value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If Value <> mProps.CodTipoDoc Then
        If Len(Value) > 0 Then
            'Actualiza IdTipoDoc también
            sql = "SELECT IdTipoDoc FROM TSTipoDocBanco WHERE CodTipoDoc='" & Value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "TSKardex_CodTipoDoc", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdTipoDoc = rs!IdTipoDoc
            End If
            rs.Close
        Else
            IdTipoDoc = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodTipoDoc = Value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodTipoDoc() As String
    CodTipoDoc = Trim$(mProps.CodTipoDoc)
End Property



Public Property Let NumDoc(ByVal Value As String)
    If Len(Value) > 20 Then
        Err.Raise ERR_INVALIDO, "TSKardex_NumDoc", MSGERR_INVALIDO
    End If
    
    mProps.NumDoc = Value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get NumDoc() As String
    NumDoc = Trim$(mProps.NumDoc)
End Property

Public Property Let FechaEmision(ByVal Value As Date)
    mProps.FechaEmision = Value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get FechaEmision() As Date
    FechaEmision = mProps.FechaEmision
End Property

Public Property Let FechaVenci(ByVal Value As Date)
    mProps.FechaVenci = Value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get FechaVenci() As Date
    FechaVenci = mProps.FechaVenci
End Property

Public Property Let Observacion(ByVal Value As String)
    If Len(Value) > 80 Then
        Err.Raise ERR_INVALIDO, "TSKardex_Observacion", MSGERR_INVALIDO
    End If
    
    mProps.Observacion = Value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Observacion() As String
    Observacion = Trim$(mProps.Observacion)
End Property

Public Property Let BandConciliado(ByVal Value As Date)
    mProps.BandConciliado = Value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get BandConciliado() As Date
    BandConciliado = mProps.BandConciliado
End Property


Public Property Let Orden(ByVal Value As Integer)
    mProps.Orden = Value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property






Friend Property Get auxIdCuenta() As Long
    'Si no está asignado IdBanco no devuelve nada
    If IdBanco = 0 Then Exit Property
    auxIdCuenta = mProps.auxIdCuenta
End Property

Friend Property Get auxBandCalculoBanco() As Integer
    Dim objTipoDoc As TSTipoDocBanco

    'Si no está asignado IdTipoDoc no devuelve nada
    If IdTipoDoc = 0 Then Exit Property
    
    If mProps.auxBandCalculoBanco < -1 Then
        Set objTipoDoc = mProps.objGNComprobante.Empresa.RecuperaTSTipoDocBanco(IdTipoDoc)
        mProps.auxBandCalculoBanco = objTipoDoc.BandCalculoBanco
        mProps.auxBandCalculoDoc = objTipoDoc.BandCalculoDoc
        auxBandCalculoBanco = objTipoDoc.BandCalculoBanco
    Else
        auxBandCalculoBanco = mProps.auxBandCalculoBanco
    End If
End Property

Friend Property Get auxBandCalculoDoc() As Integer
    Dim objTipoDoc As TSTipoDocBanco

    'Si no está asignado IdTipoDoc no devuelve nada
    If IdTipoDoc = 0 Then Exit Property
    
    If mProps.auxBandCalculoDoc < -1 Then
        Set objTipoDoc = mProps.objGNComprobante.Empresa.RecuperaTSTipoDocBanco(IdTipoDoc)
        mProps.auxBandCalculoBanco = objTipoDoc.BandCalculoBanco
        mProps.auxBandCalculoDoc = objTipoDoc.BandCalculoDoc
        auxBandCalculoDoc = objTipoDoc.BandCalculoDoc
    Else
        auxBandCalculoDoc = mProps.auxBandCalculoDoc
    End If
End Property


Private Sub Class_Initialize()
    mProps.auxBandCalculoBanco = -9     'Indica que No está asignado
    mProps.auxBandCalculoDoc = -9       'Indica que No está asignado
End Sub

Private Sub Class_Terminate()
    Set mProps.objGNComprobante = Nothing
End Sub


'*** oliver 29/may/2001


Public Function Verificar_ExisteDocumento(CodBanco As String, CodTipoDoc As String, NumDoc As String) As Boolean
Dim rs As Recordset, sql As String, numReg As Long
On Error GoTo ErrTrap
sql = "SELECT Count(TransID) as Cuantos FROM " & _
       "(TSKardex INNER JOIN TSBanco ON TSKardex.IdBanco = TSBanco.IdBanco) " & _
       "INNER JOIN TSTipoDocBanco ON TSKardex.IdTipoDoc = TSTipoDocBanco.IdTipoDoc " & _
       "WHERE TSBanco.CodBanco = '" & CodBanco & "' AND " & _
       "TSTipoDocBanco.CodTipoDoc = '" & CodTipoDoc & "' AND " & _
       "TSKARDEX.NumDoc='" & NumDoc & "'"
'Debug.Print sql
Set rs = Me.gnComprobante.Empresa.OpenRecordset(sql)
    Verificar_ExisteDocumento = IIf(rs!Cuantos = 0, False, True)
Exit Function
ErrTrap:
    MsgBox Err.Number & " " & Err.Description, vbInformation
    Verificar_ExisteDocumento = False
End Function

Public Function AsignaNumCheque(CodB As String) As Long
    Dim num As Long, sql As String, rs As Recordset, rs2 As Recordset
    
    'Obtiene el numero siguiente
    sql = "SELECT NumChequeSiguiente FROM TSBanco WHERE CodBanco='" & CodB & "'"

#If DAOLIB Then
    Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
    rs.Edit                         'Para que otros usuarios no tengan accesso
#Else
    Set rs = New Recordset
    rs.CursorLocation = adUseServer
    Set rs = Me.gnComprobante.Empresa.OpenRecordset(sql)
#End If
    If Not IsNull(rs!NumChequeSiguiente) Then
        num = rs!NumChequeSiguiente          '*** MAKOTO 16/ene/01 Mod.
    End If
    
    'Asegura que no esté ocupado el número
'    Do
        sql = "SELECT NumDoc "
        sql = sql & " From GNComprobante g  "
        sql = sql & " inner join Tskardex t"
        sql = sql & " inner join TSTipoDocBanco TS on t.IdTipoDoc=ts.IdTipoDoc"
        sql = sql & " inner join tsbanco tb on t.idbanco=tb.idbanco"
        sql = sql & " on g.transid=t.transid"
        sql = sql & " WHERE CodBanco='" & CodB & "' AND NumDoc='" & num & "'"
        sql = sql & " and CodTipoDoc='CH-E'"
        Set rs2 = Me.gnComprobante.Empresa.OpenRecordset(sql)
        'If rs2.EOF Then Exit Do
        If rs2.RecordCount > 0 Then
            MsgBox "El siguiente número de Cheque es: " & rs2.Fields("NumDoc") & " pero ya existe" & Chr(13) & "Ingrese el número de documento"
            num = 0
        End If

    AsignaNumCheque = num
    num = num + 1

    Set rs = Nothing
    Set rs2 = Nothing
    
End Function


Public Sub GrabaNumChequeNuevo(CodB As String)
End Sub


Public Sub GrabaNuevoNumCheque(CodB As String, numCheque As Long)
    Dim num As Long, sql As String, rs As Recordset, rs2 As Recordset
    sql = "UPDATE TSBanco "
    sql = sql & " SET NumChequeSiguiente = " & numCheque + 1
    sql = sql & " WHERE CodBanco='" & CodB & "'"
    Set rs = Me.gnComprobante.Empresa.OpenRecordset(sql)
    Set rs = Nothing
End Sub

Public Sub ModificaNuevoId(Id_old As Long)
    Dim num As Long, sql As String, rs As Recordset, rs2 As Recordset
    sql = "UPDATE tskardexconcilia "
    sql = sql & " SET IdTsKardex = " & mProps.id
    sql = sql & " WHERE IdTsKardex=" & Id_old
    Set rs = Me.gnComprobante.Empresa.OpenRecordset(sql)
    Set rs = Nothing
End Sub


Public Function CalculaSaldo(ByVal id As Long) As Currency
    Dim sql As String, campo As String, NumMon As Integer
    Dim rs As Recordset, transid As Long



    NumMon = Me.gnComprobante.Empresa.GNOpcion.IndiceMoneda(Me.gnComprobante.CodMoneda)
    campo = "Saldo" & NumMon
    
    
    
    
    sql = "SELECT " & campo & " FROM vwConsPCDocSaldo WHERE Id=" & id
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
    If Not rs.EOF Then
        CalculaSaldo = rs.Fields(campo)
    End If
    rs.Close
    Set rs = Nothing
End Function


Friend Property Let IdBancoRef(ByVal Value As Long)
    mProps.IdBancoRef = Value
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Get IdBancoRef() As Long
    IdBancoRef = mProps.IdBancoRef
End Property

Public Property Get CodBancoRef() As String
    CodBancoRef = Trim$(mProps.CodBancoRef)
End Property

Public Property Let CodBancoRef(ByVal Value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If Value <> mProps.CodBancoRef Then
        If Len(Value) > 0 Then
            'Actualiza IdBancoRef también
            sql = "SELECT IdBanco FROM IVBanco WHERE CodBanco='" & Value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCKardex_CodBancoRef", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdBancoRef = rs!IdBanco
            End If
            rs.Close
        Else
            mProps.IdBancoRef = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodBancoRef = Value
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Let IdTarjeta(ByVal Value As Long)
    mProps.IdTarjeta = Value
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Get IdTarjeta() As Long
    IdTarjeta = mProps.IdTarjeta
End Property

Public Property Get CodTarjeta() As String
    CodTarjeta = Trim$(mProps.CodTarjeta)
End Property

Public Property Let CodTarjeta(ByVal Value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If Value <> mProps.CodTarjeta Then
        If Len(Value) > 0 Then
            'Actualiza IdTarjeta también
            sql = "SELECT IdTarjeta FROM IVTarjeta WHERE CodTarjeta='" & Value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCKardex_CodTarjeta", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdTarjeta = rs!IdTarjeta
            End If
            rs.Close
        Else
            mProps.IdTarjeta = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodTarjeta = Value
    mProps.objGNComprobante.SetModificado
End Property


Public Property Let NumCuenta(ByVal Value As String)
    If Len(Value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCKardex_NumCuenta", MSGERR_INVALIDO
    End If
    
    mProps.NumCuenta = Value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get NumCuenta() As String
    NumCuenta = Trim$(mProps.NumCuenta)
End Property

Public Property Let numCheque(ByVal Value As String)
    If Len(Value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCKardex_NumCheque", MSGERR_INVALIDO
    End If
    
    mProps.numCheque = Value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get numCheque() As String
    numCheque = Trim$(mProps.numCheque)
End Property

Public Property Let TitularCta(ByVal Value As String)
    If Len(Value) > 50 Then
        Err.Raise ERR_INVALIDO, "PCKardex_TitularCta", MSGERR_INVALIDO
    End If
    
    mProps.TitularCta = Value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get TitularCta() As String
    TitularCta = Trim$(mProps.TitularCta)
End Property




