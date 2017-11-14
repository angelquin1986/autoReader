VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CcKardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit




Private Type T_PROP
    Id As Long
    idCentro As Long
    IdRecargo As Long
    CodRecargo As String
    CodCentro As String
    IdAsignado As Long
    Debe As Currency
    Haber As Currency
    FechaEmision As Date
    Observacion As String
    Orden As Integer
    Guid As String                          '*** MAKOTO 16/mar/01 Agregado para Exportación/Importación
    objGNComprobante As gnComprobante           'Referencia al objeto padre
    auxIdCuenta As Long          'Auxiliar para almacenar Id de cuenta de proveedor/cliente. (No se graba en la base)
    auxIdCuenta2 As Long          '*** MAKOTO 12/feb/01 Agregado
    BandPorCobrar As Boolean          'True:Por cobrar, False:Por pagar (Esto no se guarda en la base)
    NoVerificarSaldo As Boolean
End Type

Private mProps As T_PROP

Public Sub VerificaDatos()
    Dim cck As CcKardex, ix As Long, ccd As CcDocAsignado, v As Currency
    Dim msg As String

    'Verifica datos
    If mProps.Debe = 0 And mProps.Haber = 0 Then
        Err.Raise ERR_INVALIDO, "CcKardex.VerificaDatos", _
            "Se intentó grabar sin valor de Debe ni de Haber " & _
                "en detalle de Documentos de Importacion"
    End If
    If (mProps.Debe <> 0) And (mProps.Haber <> 0) Then
        Err.Raise ERR_INVALIDO, "PCKardex.VerificaDatos", _
            "Se intentó grabar con Debe y Haber al mismo tiempo " & _
                "en detalle de Documentos de Proveedor/Cliente."
    End If
    If mProps.IdAsignado = 0 And mProps.IdRecargo = 0 Then
        Err.Raise ERR_INVALIDO, "CcKardex_VerificaDatos", _
            "Debe seleccionar código de Recargo del documento " & _
                "en detalle de Documentos de Importacion."
    End If
    If mProps.idCentro = 0 Then
        Err.Raise ERR_INVALIDO, "CcKardex.VerificaDatos", _
            "Debe seleccionar un Centro de Costo " & _
                "en detalle de Documentos de Importacion"
    End If
    
       'Verifica si repite el mismo doc. asignado en el comprobante
    If mProps.IdAsignado And mProps.objGNComprobante.EsNuevo Then
        For ix = 1 To mProps.objGNComprobante.CountCcKardex
            Set cck = mProps.objGNComprobante.CcKardex(ix)
            If (cck.IdAsignado = mProps.IdAsignado) And (mProps.IdAsignado <> -1) And (Not (cck Is Me)) Then
                msg = "Existen documentos repetidos para el mismo comprobante." & vbCr & vbCr
                Set ccd = RecuperaCcDocAsignado
                If ccd.Id Then msg = msg & ccd.Trans & " " & ": "
                msg = msg & Format(Me.Debe + Me.Haber, Me.gnComprobante.FormatoMoneda) & " " & _
                            Me.gnComprobante.CodMoneda & " y " & _
                            Format(cck.Debe + cck.Haber, cck.gnComprobante.FormatoMoneda) & _
                            " " & cck.gnComprobante.CodMoneda
                Err.Raise ERR_INVALIDO, "CcKardex_VerificaDatos", msg
            End If
            If cck.IdAsignado = mProps.IdAsignado Then
                v = v + cck.Debe + cck.Haber
            End If
        Next ix
        Set cck = Nothing
        Set ccd = Nothing
    End If
End Sub

Friend Sub Grabar()
    Dim sql As String
    Dim id_antes As Long, rs As Recordset
    
    id_antes = mProps.Id        'Guarda el id anterior
    
    VerificaDatos
    
    'Graba en PCKardex
    sql = "SELECT * FROM CcKardex WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !TransID = mProps.objGNComprobante.TransID
        !idCentro = mProps.idCentro
        !IdAsignado = mProps.IdAsignado
        !IdRecargo = mProps.IdRecargo
        !Debe = mProps.Debe
        !Haber = mProps.Haber
        !FechaEmision = mProps.FechaEmision
        !Observacion = mProps.Observacion
        !Orden = mProps.Orden
        
        If Len(mProps.Guid) > 0 Then !Guid = mProps.Guid        '*** MAKOTO 16/mar/01 Agregado
        .Update
        
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.Id = !Id             'Guarda el Id(Autonumerico)
        .Close
    End With
    
    'Si está modificando
    If id_antes <> 0 Then
        'Actualiza IdAsignado de CcKardex relacionado
        '  para que apunte de nuevo a éste registro.
        sql = "UPDATE CCKardex SET IdAsignado=" & mProps.Id & _
              " WHERE IdAsignado=" & id_antes
#If DAOLIB Then
        mProps.objGNComprobante.Empresa.Database.Execute sql, dbFailOnError
#Else
        mProps.objGNComprobante.Empresa.Coneccion.Execute sql
#End If
    End If

    'Si asigna Id al campo IdAsignado de otros CcKardexs en la misma transacción
    ' para descargar cobros/pagos al contado
    AsignaIDCcK

    Set rs = Nothing
End Sub

'Asigna IdAsignado en los CcKardex que está en el mismo comprobante y que tengan ahora -1 como IdAsignado
' Es para descargar cobro/pago al contado
Private Sub AsignaIDCcK()
    Dim i As Long, cck As CcKardex
    
    For i = 1 To mProps.objGNComprobante.CountCcKardex
        Set cck = mProps.objGNComprobante.CcKardex(i)
        If (cck.IdAsignado = -1) And _
            (cck.CodRecargo = Me.CodRecargo) And _
           ((cck.Debe + cck.Haber) = (Me.Debe + Me.Haber)) Then
            cck.IdAsignado = mProps.Id
            cck.CodCentro = mProps.CodCentro
            cck.CodRecargo = mProps.CodRecargo
            Exit For
        End If
    Next i
    Set cck = Nothing
End Sub

Friend Sub Recuperar(rs As Recordset)
    With rs
        If Not .EOF Then
            mProps.Id = !Id
            If Not IsNull(!idCentro) Then mProps.idCentro = !idCentro
            If Not IsNull(!CodCentro) Then mProps.CodCentro = !CodCentro
            If Not IsNull(!IdRecargo) Then mProps.IdRecargo = !IdRecargo
            If Not IsNull(!CodRecargo) Then mProps.CodRecargo = !CodRecargo
            If Not IsNull(!IdAsignado) Then mProps.IdAsignado = !IdAsignado
            If Not IsNull(!Debe) Then mProps.Debe = !Debe
            If Not IsNull(!Haber) Then mProps.Haber = !Haber
            If Not IsNull(!FechaEmision) Then mProps.FechaEmision = !FechaEmision
            If Not IsNull(!Observacion) Then mProps.Observacion = !Observacion
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            If Not IsNull(!Guid) Then mProps.Guid = !Guid       '*** MAKOTO 16/mar/01 Agregado
'            If Not IsNull(!IdCuentaContable) Then mProps.auxIdCuenta = !IdCuentaContable
 '           If Not IsNull(!IdCuentaContable2) Then mProps.auxIdCuenta2 = !IdCuentaContable2
        End If
    End With
End Sub

Friend Function Clone() As CcKardex
    Dim nuevo As CcKardex
    
    Set nuevo = New CcKardex
    nuevo.Props = mProps
    '*** No sirve porque cuendo clona por el cierre no funciona
    'nuevo.Guid = "" '*** borro el GUID del njuevo objeto para que no se repita nunca.  ***Oliver 10/12/2004'
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROP)          'Para el método Clone
    mProps = value
    mProps.Id = 0
    mProps.objGNComprobante.SetModificado
End Property


Friend Property Set gnComprobante(value As gnComprobante)
    Set mProps.objGNComprobante = value
    
    'Coge la fecha de emisión
    Me.FechaEmision = value.FechaTrans
End Property

Public Property Get gnComprobante() As gnComprobante
    Set gnComprobante = mProps.objGNComprobante
End Property

Friend Property Let BandPorCobrar(ByVal value As Boolean)
    mProps.BandPorCobrar = value
End Property

Public Property Get BandPorCobrar() As Boolean
    BandPorCobrar = mProps.BandPorCobrar
End Property


Public Property Get Id() As Long
    Id = mProps.Id
End Property


Friend Property Get idCentro() As Long
    idCentro = mProps.idCentro
End Property

Public Property Let CodCentro(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCentro Then
        If Len(value) > 0 Then
            'Actualiza IdCentro también
            sql = "SELECT IdCentro " & _
                  "FROM GNCentroCosto WHERE CodCentro='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCKardex_CodCentro", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.idCentro = rs!idCentro
            End If
            rs.Close
        Else
            mProps.idCentro = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodCentro = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodCentro() As String
    CodCentro = Trim$(mProps.CodCentro)
End Property

Friend Property Let IdRecargo(ByVal value As Long)
    mProps.IdRecargo = value
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Get IdRecargo() As Long
    IdRecargo = mProps.IdRecargo
End Property


Public Property Let CodRecargo(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodRecargo Then
        If Len(value) > 0 Then
            'Actualiza IdRecargo también
            sql = "SELECT IdRecargo FROM IvRecargo WHERE CodRecargo='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "CcKardex_CodRecargo", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdRecargo = rs!IdRecargo
            End If
            rs.Close
        Else
            mProps.IdRecargo = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodRecargo = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodRecargo() As String
    CodRecargo = Trim$(mProps.CodRecargo)
End Property




Public Property Let IdAsignado(ByVal value As Long)
    mProps.IdAsignado = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get IdAsignado() As Long
    IdAsignado = mProps.IdAsignado
End Property

'*** MAKOTO 12/ene/01 Agregado para la necesidad de SiiTools (Importar trans.)
Public Sub SetIdAsignadoPorTrans( _
                ByVal CodTrans As String, _
                ByVal NumTrans As Long, _
                ByVal Orden As Long)
    Dim sql As String, rs As Recordset
    
    'Si  el documento asignado NO está en la misma transacción
    If Not ((CodTrans = Me.gnComprobante.CodTrans) _
            And (NumTrans = Me.gnComprobante.NumTrans)) Then
            
        'Busca CodTrans, NumTrans y Orden del doc. asignado
        sql = "SELECT cck.Id FROM CcKardex cck INNER JOIN GNComprobante gc " & _
                        "ON cck.TransID = gc.TransID " & _
              "WHERE gc.CodTrans = '" & CodTrans & "' AND gc.NumTrans = " & NumTrans & _
                     " AND cck.Orden = " & Orden
        Set rs = Me.gnComprobante.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then
            mProps.IdAsignado = rs.Fields("Id")
            mProps.objGNComprobante.SetModificado
        Else
            Err.Raise ERR_NOHAYCODIGO, "CcKardex", _
                        "No se encuentra el documento asignado. " & _
                        "(" & CodTrans & NumTrans & " - #" & Orden & ")"
        End If
        rs.Close
        Set rs = Nothing
        
    'Si el doc.asignado es de la misma transacción como en caso de pago/cobro a CONTADO
    Else
        'IdAsignado queda pendiente hasta que grabe la transacción
        mProps.IdAsignado = -1
    End If
End Sub

'*** MAKOTO 16/mar/01 Agregado para la necesidad de SiiTools (Importar trans.)
Public Sub SetIdAsignadoPorGuid(ByVal GuidAsignado As String)
    Dim sql As String, rs As Recordset
    
    'Busca Guid del doc. asignado
    sql = "SELECT Id, TransID FROM CcKardex WHERE Guid = '" & GuidAsignado & "'"
    Set rs = Me.gnComprobante.Empresa.OpenRecordset(sql)
    If Not rs.EOF Then
        'Si  el documento asignado NO está en la misma transacción
        If rs.Fields("TransID") <> Me.gnComprobante.TransID Then
            mProps.IdAsignado = rs.Fields("Id")
        Else
            'Si el doc.asignado es de la misma transacción como en caso de pago/cobro a CONTADO
            'IdAsignado queda pendiente hasta que grabe la transacción
            mProps.IdAsignado = -1
        End If
        mProps.objGNComprobante.SetModificado
    
    Else
        'Puede ser que no encuentre guid en la BD en caso de ser nueva transaccion
        'En ese caso buscamos dentro de la transaccion que está en la memoria
        If BuscarGuidEnTrans(GuidAsignado) Then
            'Si el doc.asignado es de la misma transacción como en caso de pago/cobro a CONTADO
            'IdAsignado queda pendiente hasta que grabe la transacción
            mProps.IdAsignado = -1
        Else
            Err.Raise ERR_NOHAYCODIGO, "CcKardex", _
                    "No se encuentra el documento asignado. " & vbCr & _
                    "(" & Me.gnComprobante.CodTrans & _
                          Me.gnComprobante.NumTrans & ") " & vbCr & GuidAsignado
        End If
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Function BuscarGuidEnTrans(ByVal GuidBuscar As String) As Boolean
    Dim i As Long, cck As CcKardex
    
    BuscarGuidEnTrans = False
    For i = 1 To Me.gnComprobante.CountCcKardex
        Set cck = Me.gnComprobante.CcKardex(i)
        If Not (cck Is Me) Then
            If cck.Guid = GuidBuscar Then
                BuscarGuidEnTrans = True
                GoTo salida
            End If
        End If
    Next i
salida:
    Set cck = Nothing
End Function

Private Sub VerificaSaldo(ByVal value As Currency, ByVal bandDebe As Boolean)
    Dim Vdoc As Currency, Saldo As Currency, ccd As CcDocAsignado, cancelado As Currency

    'Si no necesita verificar, sale.            '*** MAKOTO 22/mar/01 Agregado
    If mProps.NoVerificarSaldo Then Exit Sub

    Vdoc = IIf(bandDebe, Debe, Haber)
    
    'Si quiere reducir el valor de doc. por cobrar/pagar
    If (mProps.IdAsignado = 0) And (value < Vdoc) Then
'    If (mProps.IdAsignado = 0) Then
        'Obtiene el valor cancelado
        cancelado = CalculaCancelado - CalculaCanceladoEnTrans
        
        If value < cancelado Then
            'No puede ser menor que valor cobrado
            Err.Raise ERR_INVALIDO, "CcKardex.VerificaSaldo", _
                "No se puede reducir el valor debido a que ya está cancelado la cantidad de " & _
                cancelado & " " & mProps.objGNComprobante.CodMoneda
        End If
    
    'Si quiere aumentar el valor de cobro/pago
    ElseIf (mProps.IdAsignado <> 0) And (value > Vdoc) Then
'    ElseIf (mProps.IdAsignado <> 0) Then
        'Obtiene el doc. por pagar asignado
        Set ccd = RecuperaCcDocAsignado
        If ccd.Id Then      'Si es que encuentra
        
            'Convierte en moneda del doc. asignado
            value = value * Me.gnComprobante.Cotizacion("")
            value = value / mProps.objGNComprobante.Cotizacion(ccd.CodMoneda)
            If value > ccd.Valor Then
                'No puede ser mayor al valor de documento asignado
                Err.Raise ERR_INVALIDO, "CcKardex_VerificaSaldo", _
                    "No se puede asignar un valor mayor al del documento. " & vbCr & vbCr & _
                    "Valor del documento: " & vbTab & ccd.Valor & vbCr & _
                    "Valor asignado:        " & vbTab & value & _
                    " " & mProps.objGNComprobante.CodMoneda
            ElseIf value > ccd.Saldo + Vdoc Then
                'No puede ser mayor al saldo del documento asignado
                Err.Raise ERR_INVALIDO, "CcKardex.VerificaSaldo", _
                    "No se puede asignar un valor mayor al saldo del documento. " & vbCr & vbCr & _
                    "Saldo:      " & vbTab & ccd.Saldo + Vdoc & vbCr & _
                    "Asignado: " & vbTab & value & vbCr & _
                    "Dif.:          " & vbTab & value - ccd.Saldo - Vdoc & _
                    " " & mProps.objGNComprobante.CodMoneda
            End If
        End If
        Set ccd = Nothing
    End If
End Sub


Public Property Let Debe(ByVal value As Currency)
    Dim vSucre As Currency
        
    'Convierte en sucres
    vSucre = value * mProps.objGNComprobante.Cotizacion("")
    
    If vSucre < 0 Then
        Err.Raise ERR_INVALIDO, "CcKardex_Debe", MSGERR_INVALIDO
    Else
        VerificaSaldo value, True
        
        If vSucre > 0 Then
            mProps.Haber = 0
            mProps.Debe = vSucre
        Else
            mProps.Debe = 0
        End If
    End If
    
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Debe() As Currency
    'Devuelve en moneda ex.
    Debe = mProps.Debe / mProps.objGNComprobante.Cotizacion("")
End Property

Public Property Let Haber(ByVal value As Currency)
    Dim vSucre As Currency
    
    'Convierte en sucres
    vSucre = value * mProps.objGNComprobante.Cotizacion("")
    
    If vSucre < 0 Then
        Err.Raise ERR_INVALIDO, "CcKardex_Haber", MSGERR_INVALIDO
    Else
        VerificaSaldo value, False      'Hay que mandar el valor en moneda de trans.
        
        If vSucre > 0 Then
            mProps.Debe = 0
            mProps.Haber = vSucre
        Else
            mProps.Haber = 0
        End If
    End If
    
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Haber() As Currency
    Haber = mProps.Haber / mProps.objGNComprobante.Cotizacion("")
End Property

'*** MAKOTO 22/mar/01 Agregado para no hacer la verificación de saldo
'En casos especiales como por ejemplo Importacion de datos de SiiTools.
Public Property Let BandNoVerificarSaldo(ByVal value As Boolean)
    mProps.NoVerificarSaldo = value
End Property

Public Property Get BandNoVerificarSaldo() As Boolean
    BandNoVerificarSaldo = mProps.NoVerificarSaldo
End Property

Public Property Let FechaEmision(ByVal value As Date)
    mProps.FechaEmision = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get FechaEmision() As Date
    FechaEmision = mProps.FechaEmision
End Property


Public Property Let Observacion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "CcKardex_NombreEmitante", MSGERR_INVALIDO
    End If
    
    mProps.Observacion = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Observacion() As String
    Observacion = Trim$(mProps.Observacion)
End Property


Public Property Let Orden(ByVal value As Integer)
    mProps.Orden = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property


'*** MAKOTO 16/mar/01 Agregado
Public Property Let Guid(ByVal value As String)
    mProps.Guid = value
    mProps.objGNComprobante.SetModificado
End Property

'*** MAKOTO 16/mar/01 Agregado
Public Property Get Guid() As String
    Guid = mProps.Guid
End Property


Friend Property Get auxIdCuenta() As Long
    'Si no está asignado IdCentro no devuelve nada
    If idCentro = 0 Then Exit Property
    auxIdCuenta = mProps.auxIdCuenta
End Property

Friend Property Get auxIdCuenta2() As Long
    'Si no está asignado IdCentro no devuelve nada
    If idCentro = 0 Then Exit Property
    auxIdCuenta2 = mProps.auxIdCuenta2
End Property

Public Function RecuperaCcDocAsignado() As CcDocAsignado
    Dim ccd As CcDocAsignado

    If mProps.IdAsignado = 0 Then Exit Function
    
    Set ccd = New CcDocAsignado
    Set ccd.CcKardex = Me
    ccd.Recuperar mProps.IdAsignado
    Set RecuperaCcDocAsignado = ccd
    Set ccd = Nothing
End Function

Public Function CalculaSaldo() As Currency
    Dim sql As String, campo As String, NumMon As Integer
    Dim rs As Recordset

    NumMon = Me.gnComprobante.Empresa.GNOpcion.IndiceMoneda(Me.gnComprobante.CodMoneda)
    campo = "Saldo" & NumMon
    
    sql = "SELECT " & campo & " FROM vwConsPCDocSaldo WHERE Id=" & mProps.Id
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
    If Not rs.EOF Then
        CalculaSaldo = rs.Fields(campo)
    End If
    rs.Close
    Set rs = Nothing
End Function

Public Function CalculaCancelado() As Currency
    Dim sql As String, campo As String, NumMon As Integer
    Dim rs As Recordset

    NumMon = Me.gnComprobante.Empresa.GNOpcion.IndiceMoneda(Me.gnComprobante.CodMoneda)
    campo = "VCancelado" & NumMon
    
    sql = "SELECT " & campo & " FROM vwConsPCDocSaldo WHERE Id=" & mProps.Id
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
    If Not rs.EOF Then
        CalculaCancelado = rs.Fields(campo)
    End If
    rs.Close
    Set rs = Nothing
End Function

Public Function CalculaCanceladoEnTrans() As Currency
    Dim cck As CcKardex, ix As Long, c As Currency
    
    For ix = 1 To Me.gnComprobante.CountCcKardex
        Set cck = Me.gnComprobante.CcKardex(ix)
        If cck.IdAsignado = Me.Id Then
            c = c + cck.Debe + cck.Haber
        End If
    Next ix
    Set cck = Nothing
    CalculaCanceladoEnTrans = c
End Function


'Devuelve True si es nuevo o no tiene cobros/pagos asignados
'         False si tiene cobros/pagos asignados
Public Function VerificaEliminacion() As Boolean
    Dim sql As String, rs As Recordset
    
    VerificaEliminacion = True

    'Si es nuevo no hay para que consultar la base
    If mProps.Id = 0 Then Exit Function

    'Busca en la PCKardex registros asignados que sea de otro comprobante
    sql = "SELECT Id FROM CcKardex WHERE IdAsignado=" & mProps.Id & _
                                    " AND TransID<>" & Me.gnComprobante.TransID

    
    'Si existe, devuelve False sino True
    If Not rs.EOF Then
        VerificaEliminacion = False
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Sub Class_Initialize()
    With mProps
        .FechaEmision = Date
        'Activa la verificación de saldo de doc. asignado
        .NoVerificarSaldo = True                '*** MAKOTO 22/mar/01 Agregado
    End With
End Sub

Private Sub Class_Terminate()
    Set mProps.objGNComprobante = Nothing
End Sub


'*** MAKOTO 16/mar/01 Agregado
'Recibe guid y lo busca en CcKardex
'Guarda id del registro encontrado en la propiedad id del objeto.
'
'Esto sirve para que no se pierda enlace entre Doc.Por Cobrar/pagar y
'Pagos/Cobros asignados al momento de sobreescbirir en el proceso de
'Exportacion/Importacion
'
'Al llamar al método Grabar, el valor de id será reemplazado por otro valor
'generado por BD mismo, sin embargo el valor asignado en éste momento
'por éste método sirve como 'Valor anterior' para re-asignar los valores
'de IdAsignado del Cobros/Pagos originalmente asignados a éste documento.
Public Sub SetIdFromGuid()
    Dim sql As String, rs As Recordset
    
    sql = "SELECT Id FROM cCKardex WHERE guid='" & mProps.Guid & "'"
    Set rs = Me.gnComprobante.Empresa.OpenRecordset(sql)
    If rs.EOF Then
'No debe generar ningún error porque cuando es transacción nueva no ha de existir
'        Err.Raise ERR_NOHAYCODIGO, "PCKardex.SetIdFromGuid", _
'            MSGERR_NOHAYCODIGO & vbCr & "Guid=" & mProps.Guid
    Else
        mProps.Id = rs.Fields("Id")
    End If
    rs.Close
    Set rs = Nothing
End Sub
'JEAA 31/08/2005
'Devuelve True si es nuevo o no tiene cobros/pagos asignados Y EN TRANS EL DOUMENTO ASIGNADO
'         False si tiene cobros/pagos asignados
Public Function VerificaEliminacionNEW(ByRef Trans As String) As Boolean
    Dim sql As String, rs As Recordset
    
    VerificaEliminacionNEW = True

    'Si es nuevo no hay para que consultar la base
    If mProps.Id = 0 Then Exit Function


    'Busca en la CcKardex registros asignados que sea de otro comprobante
    'JEAA       modificado para que visualice cual es el comprobante
    sql = " SELECT gnc.codtrans, gnc.numtrans " & _
            " FROM CcKardex cck INNER JOIN GNComprobante gnc " & _
            " ON gnc.transid=cck.transid " & _
            " WHERE IdAsignado=" & mProps.Id & _
            " AND cck.TransID<>" & Me.gnComprobante.TransID
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
    
    'Si existe, devuelve False sino True
    If Not rs.EOF Then
        Trans = rs.Fields(0) & "-" & rs.Fields(1)
        VerificaEliminacionNEW = False
    End If
    rs.Close
    Set rs = Nothing
End Function



