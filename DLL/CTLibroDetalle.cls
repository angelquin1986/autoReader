VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTLibroDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"GNComprobante"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
    Option Explicit


Private Type T_PROP
    id As Long
    IdCuenta As Long
    CodCuenta As String
    auxNombreCuenta As String
    Descripcion As String
    Debe As Currency
    Haber As Currency
    BandIntegridad As Byte              'Indica si está integrado o no al llamar VerificaIntegridad de GNComprobante.
                                        '0=Está integrado, 1=AF,2=TS,3=IV,4=PC,5=RL
    Orden As Integer
    idGasto As Long             'jeaa 07/01/2008
    CodGasto As String          'jeaa 07/01/2008
    BandNIIF As Byte
   
    objGNComprobante As GNComprobante       'Referencia al objeto padre
End Type

Private mProps As T_PROP

Public Sub VerificaDatos()

    If (IdCuenta = 0) Or (Len(mProps.CodCuenta) = 0) Then
        Err.Raise ERR_INVALIDO, "CTLibroDetalle.Verificación", _
            "Tiene que asignar una cuenta contable en detalle de asiento."
    End If
    
    If (Debe = 0 And Haber = 0) Then
        Err.Raise ERR_INVALIDO, "CTLibroDetalle.Verificación", _
            "Tiene que asignar un valor a Debe ó Haber en detalle de asiento."
    End If
    
    If (Debe <> 0 And Haber <> 0) Then
        Err.Raise ERR_INVALIDO, "CTLibroDetalle.Verificación", _
            "No puede asignar valor a Debe y Haber al mismo tiempo en detalle de asiento."
    End If
    
    If Len(mProps.objGNComprobante.GNTrans.CodGastoPre) > 0 And mProps.BandIntegridad <> 100 Then
        If mProps.objGNComprobante.GNTrans.OmitirGastoxZona Then
            If Len(mProps.CodGasto) = 0 Then
            Err.Raise ERR_INVALIDO, "CTLibroDetalle.Verificación", _
                "Falta seleccionar el Gasto o Zona para la cuenta: " & mProps.CodCuenta
            End If
        End If
    End If
    

End Sub


Friend Sub Grabar()
    Dim sql As String, rs As Recordset

    VerificaDatos
    sql = "SELECT * FROM CTLibroDetalle WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !CodAsiento = mProps.objGNComprobante.CodAsiento
        !IdCuenta = mProps.IdCuenta
        !Descripcion = mProps.Descripcion
        !Debe = mProps.Debe
        !Haber = mProps.Haber
        !BandIntegridad = mProps.BandIntegridad
        !Orden = mProps.Orden
        !idGasto = mProps.idGasto  ' jeaa 07/01/2008
        !BandNIIF = mProps.BandNIIF

        .Update
        
#If DAOLIB Then
        rs.Bookmark = rs.LastModified   'Para actualizar
#End If
        .Move 0                     'Para actualizar
        mProps.id = !id             'Guarda el Id(Autonumerico)
        .Close
    End With
    Set rs = Nothing
End Sub

Friend Sub Recuperar(rs As Recordset)
    With rs
        If Not .EOF Then
            mProps.id = !id
            mProps.IdCuenta = !IdCuenta
            If Not IsNull(!CodCuenta) Then mProps.CodCuenta = !CodCuenta
            If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
            If Not IsNull(!Debe) Then mProps.Debe = !Debe
            If Not IsNull(!Haber) Then mProps.Haber = !Haber
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            If Not IsNull(!idGasto) Then mProps.idGasto = !idGasto
            If Not IsNull(!CodGasto) Then mProps.CodGasto = !CodGasto
            
            If Not IsNull(!BandIntegridad) Then
                If !BandIntegridad = INTEG_AUTO Then
                    mProps.BandIntegridad = INTEG_AUTO
                    
                'Cuand es agregado por usuario
                Else
                    'No se debe recuperar de la base por que tiene que ser verificado nuevamente
                    mProps.BandIntegridad = INTEG_NADA
                End If
            End If
            If Not IsNull(!BandNIIF) Then mProps.BandNIIF = !BandNIIF

            
        End If
    End With
End Sub

Friend Function Clone() As CTLibroDetalle
    Dim nuevo As CTLibroDetalle
    
    Set nuevo = New CTLibroDetalle
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROP)          'Para el método Clone
    mProps = value
    mProps.id = 0
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Set GNComprobante(value As GNComprobante)
    Set mProps.objGNComprobante = value
End Property

Public Property Get GNComprobante() As GNComprobante
    Set GNComprobante = mProps.objGNComprobante
End Property



Public Property Get id() As Long
    id = mProps.id
End Property

'Llama solo desde GeneraAsiento
Friend Sub SetIdCuentaAuto(IdCta As Long)
    Dim ct As CTCuenta

    mProps.IdCuenta = IdCta
    mProps.objGNComprobante.SetModificado
    
    'Sincloniza CodCuenta también
    On Error Resume Next
    Set ct = mProps.objGNComprobante.Empresa.RecuperaCTCuenta(mProps.IdCuenta)
    If Err.Number = 0 Then
        If Not ct Is Nothing Then
            mProps.CodCuenta = ct.CodCuenta
            mProps.auxNombreCuenta = ct.NombreCuenta
        End If
    End If
    Err.Clear
    On Error GoTo 0
    Set ct = Nothing
End Sub

Public Property Let IdCuenta(ByVal value As Long)
    Dim ct As CTCuenta

    'Recupera el registro por el id
    Set ct = mProps.objGNComprobante.Empresa.RecuperaCTCuenta(value)
    If Not (ct Is Nothing) Then
        'Sincloniza CodCuenta también
        mProps.CodCuenta = ct.CodCuenta
        mProps.auxNombreCuenta = ct.NombreCuenta
    Else
        'Si no recupera, genera un error
        Err.Raise ERR_NOHAYCODIGO, "CTLibroDetalle_IdCuenta", MSGERR_NOHAYCODIGO
    End If
    
    mProps.IdCuenta = value
    mProps.objGNComprobante.SetModificado
    
    'Si es que usuario cambia la cuenta
    mProps.BandIntegridad = INTEG_NADA  'Tiene que ser verificado
    
    Set ct = Nothing
End Property

Public Property Get IdCuenta() As Long
    IdCuenta = mProps.IdCuenta
End Property



Public Property Let CodCuenta(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCuenta Then
        'Actualiza IdCuenta también
        sql = "SELECT IdCuenta,NombreCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
        Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "CTLibroDetalle_CodCuenta", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            mProps.IdCuenta = rs!IdCuenta
            mProps.auxNombreCuenta = rs!NombreCuenta
        End If
        rs.Close
        Set rs = Nothing
        
        mProps.CodCuenta = value
        mProps.objGNComprobante.SetModificado
    End If
End Property

Public Property Get CodCuenta() As String
    CodCuenta = Trim$(mProps.CodCuenta)
End Property


Public Property Get auxNombreCuenta() As String
    Dim ct As CTCuenta
    
    If Len(mProps.auxNombreCuenta) > 0 Then
        auxNombreCuenta = mProps.auxNombreCuenta
    Else
        On Error Resume Next
        Set ct = mProps.objGNComprobante.Empresa.RecuperaCTCuenta(mProps.IdCuenta)
        If Err.Number = 0 Then auxNombreCuenta = ct.NombreCuenta
        Err.Clear
        On Error GoTo 0
    End If
    Set ct = Nothing
End Property


Public Property Let Descripcion(ByVal value As String)
    If value = mProps.Descripcion Then Exit Property
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "CTLibroDetalle_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property



Public Property Let Debe(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
'    value = RedondeaMoneda(value, Me.GNComprobante.FormatoMoneda)   'Redondea
    
    If value > 0 Then
        mProps.Debe = value
        mProps.Haber = 0
    ElseIf value < 0 Then
        mProps.Debe = 0
        mProps.Haber = -value
    Else
        mProps.Debe = 0
    End If
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Debe() As Currency
    'Devuelve en Moneda Ex.
    Debe = mProps.Debe / mProps.objGNComprobante.Cotizacion("")
End Property

'Devuelve el valor en sucre
' Es para sacar el total de Debe exactamente
Friend Property Get DebeSucre() As Currency
    DebeSucre = mProps.Debe
End Property

Public Property Let Haber(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
'    value = RedondeaMoneda(value, Me.GNComprobante.FormatoMoneda)   'Redondea
    
    If value > 0 Then
        mProps.Debe = 0
        mProps.Haber = value
    ElseIf value < 0 Then
        mProps.Debe = -value
        mProps.Haber = 0
    Else
        mProps.Haber = 0
    End If
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Haber() As Currency
    'Devuelve en Moneda Ex.
    Haber = mProps.Haber / mProps.objGNComprobante.Cotizacion("")
End Property

'Devuelve el valor de Haber en sucre
' Es para sacar el total de Haber exactamente
Friend Property Get HaberSucre() As Currency
    HaberSucre = mProps.Haber
End Property


'*** MAKOTO 12/ene/01 Mod. Antes era Friend ahora Public
'*** Por que necesito asignar desde UI(SiiTools) para importación de transacción
Public Property Let BandIntegridad(ByVal value As Byte)
    If mProps.BandIntegridad <> value Then
        mProps.BandIntegridad = value
        mProps.objGNComprobante.SetModificado
    End If
End Property

Public Property Get BandIntegridad() As Byte
    BandIntegridad = mProps.BandIntegridad
End Property

Public Sub SetIntegrado()               'Llama cuando hace la integración de asiento
    BandIntegridad = INTEG_INTEGRADO
End Sub

Public Property Let Orden(ByVal value As Integer)
    If value = mProps.Orden Then Exit Property
    mProps.Orden = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property




Private Sub Class_Initialize()
    'Valores predeterminados
    mProps.BandIntegridad = INTEG_NADA
    mProps.BandNIIF = False

End Sub

Private Sub Class_Terminate()
    Set mProps.objGNComprobante = Nothing
End Sub

Public Property Let idGasto(ByVal value As Long)
    mProps.IdCuenta = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get idGasto() As Long
    idGasto = mProps.idGasto
End Property



Public Property Let CodGasto(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodGasto Then
        'Actualiza IdGasto también
        sql = "SELECT IdGasto,Nombre FROM GNGasto WHERE CodGasto='" & value & "'"
        Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "CTLibroDetalle_CodGasto", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            mProps.idGasto = rs!idGasto
        End If
        rs.Close
        
        Set rs = Nothing
        
        mProps.CodGasto = value
        mProps.objGNComprobante.SetModificado
    End If
End Property

Public Property Get CodGasto() As String
    CodGasto = Trim$(mProps.CodGasto)
End Property

Friend Sub GrabarxCosto(ByVal id As Long)
    Dim sql As String, rs As Recordset

    VerificaDatos
    sql = "SELECT * FROM CTLibroDetalle WHERE codasiento=" & id
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    sql = "delete CTLibroDetalle WHERE codasiento=" & id
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    
    With rs
'        !CodAsiento = mProps.objGNComprobante.CodAsiento
'        !IdCuenta = mProps.IdCuenta
'        !Descripcion = mProps.Descripcion
'        !Debe = mProps.Debe
'        !Haber = mProps.Haber
'        !BandIntegridad = mProps.BandIntegridad
'        !Orden = mProps.Orden
'        !IDGasto = mProps.IDGasto  ' jeaa 07/01/2008
'        .Update
        
#If DAOLIB Then
        rs.Bookmark = rs.LastModified   'Para actualizar
#End If
''        .Move 0                     'Para actualizar
''        mProps.Id = !Id             'Guarda el Id(Autonumerico)
'        .Close
    End With
    Set rs = Nothing
End Sub


Friend Sub RecuperarxCosto(rs As Recordset)
    With rs
        If Not .EOF Then
            mProps.id = !id
            mProps.IdCuenta = !IdCuenta
            If Not IsNull(!CodCuenta) Then mProps.CodCuenta = !CodCuenta
            If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
            If Not IsNull(!Debe) Then mProps.Debe = !Debe
            If Not IsNull(!Haber) Then mProps.Haber = !Haber
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            If Not IsNull(!idGasto) Then mProps.idGasto = !idGasto
            If Not IsNull(!CodGasto) Then mProps.CodGasto = !CodGasto
            
            If Not IsNull(!BandIntegridad) Then
                If !BandIntegridad = INTEG_AUTO Then
                    mProps.BandIntegridad = INTEG_AUTO
                    
                'Cuand es agregado por usuario
                Else
                    'No se debe recuperar de la base por que tiene que ser verificado nuevamente
                    mProps.BandIntegridad = INTEG_NADA
                End If
            End If
            If Not IsNull(!BandNIIF) Then mProps.BandNIIF = !BandNIIF

            
        End If
    End With
End Sub

Public Property Let BandNIIF(ByVal value As Byte)
    If mProps.BandNIIF <> value Then
        mProps.BandNIIF = value
        mProps.objGNComprobante.SetModificado
    End If
End Property

Public Property Get BandNIIF() As Byte
    BandNIIF = mProps.BandNIIF
End Property



