VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TSKardexRet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"GNComprobante"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit



Private Type T_PROP
    Id As Long
    IdRetencion As Long
    CodRetencion As String
    Debe As Currency
    Haber As Currency
    Base As Currency
    NumDoc As String
    Observacion As String
    Orden As Integer
    
    auxIdCuentaActivo As Long          'Auxiliar para almacenar codigo de cuenta de banco. (No se graba en la base)
    auxIdCuentaPasivo As Long          'Auxiliar para almacenar codigo de cuenta de banco. (No se graba en la base)
    
    NumSerieEstaSRI As String
    NumSeriePuntoSRI As String
    NumSecuencialSRI As String
    AutorizacionSRI As String
    FechaCaducidadSRI As Date
    
    objGNComprobante As GNComprobante           'Referencia al objeto padre
End Type

Private mProps As T_PROP


Public Sub VerificaDatos()
    'Verifica datos
    Dim tsr As TSRetencion, pc As PCProvCli
    
    If mProps.Debe = 0 And mProps.Haber = 0 Then
        Set tsr = mProps.objGNComprobante.Empresa.RecuperaTSRetencion(mProps.CodRetencion)
        If Not tsr Is Nothing Then
            If tsr.Porcentaje <> 0 Then
                Err.Raise ERR_INVALIDO, "TSKardexRet.VerificaDatos", _
                    "Se intentó grabar sin valor de Debe ni de Haber en detalle de Retención."
            End If
        End If
        Set tsr = Nothing
    End If
    If mProps.Debe <> 0 And mProps.Haber <> 0 Then
        Err.Raise ERR_INVALIDO, "TSKardexRet.VerificaDatos", _
            "Se intentó grabar con Debe y Haber al mismo tiempo en detalle de Retención."
    End If
    If mProps.IdRetencion = 0 Then
        Err.Raise ERR_INVALIDO, "TSKardexRet.VerificaDatos", _
            "Debe seleccionar un código de retención."
    End If
    If Len(mProps.NumDoc) = 0 Then
        Err.Raise ERR_INVALIDO, "TSKardexRet.VerificaDatos", _
            "Debe ingresar un número de documento en detalle de Retención."
    End If
    
    If Me.GNComprobante.GNTrans.TSDatosRetencion Then
        If Len(mProps.AutorizacionSRI) = 0 Then
        Err.Raise ERR_INVALIDO, "TSKardexRet.VerificaDatos", _
            "Debe ingresar un número de Autorizacion en detalle de Retención."
        End If
        If Len(mProps.NumSerieEstaSRI) = 0 Then
        Err.Raise ERR_INVALIDO, "TSKardexRet.VerificaDatos", _
            "Debe ingresar un número de Serie Establecimiento en detalle de Retención."
        End If
        If Len(mProps.NumSeriePuntoSRI) = 0 Then
        Err.Raise ERR_INVALIDO, "TSKardexRet.VerificaDatos", _
            "Debe ingresar un número de Serie Punto en detalle de Retención."
        End If
        If Len(mProps.NumSecuencialSRI) = 0 Then
        Err.Raise ERR_INVALIDO, "TSKardexRet.VerificaDatos", _
            "Debe ingresar un número de Secuencial en detalle de Retención."
        End If
        If mProps.FechaCaducidadSRI < mProps.objGNComprobante.FechaTrans Then
            Err.Raise ERR_INVALIDO, "TSKardexRet.VerificaDatos", _
            "La fecha de Caducidad " & mProps.FechaCaducidadSRI & " del documento " & vbCrLf & _
                            "NO puede ser menor que la fecha de la transacción" & mProps.objGNComprobante.FechaTrans
        End If
        'jeaa 17/09/2007
        If DateDiff("m", mProps.objGNComprobante.FechaTrans, mProps.FechaCaducidadSRI) > 12 Then
            Err.Raise ERR_INVALIDO, "TSKardexRet.VerificaDatos", _
            "La fecha de Caducidad del documento de compra no puede mayor a un año que la" & vbCrLf & _
                            "fecha de la transacción"
        End If
    End If
    
    If Len(mProps.objGNComprobante.CodClienteRef) > 0 Then
        Set pc = mProps.objGNComprobante.Empresa.RecuperaPCProvCliQuick(mProps.objGNComprobante.CodClienteRef)
        If Not pc Is Nothing Then
            If pc.codtipoDocumento <> "R" And pc.codtipoDocumento <> "F" Then
            Err.Raise ERR_INVALIDO, "TSKardexRet.VerificaDatos", _
            "Con el Tipo de Comprobante " & pc.codtipoDocumento & " del Cliente no se puede recibir la retención " & vbCrLf & _
            "Debe cambiarlo en el Catalogo de Clientes"
            End If
        End If
        Set pc = Nothing
    End If
    
End Sub

Friend Sub Grabar()
    Dim sql As String, rs As Recordset
    
    VerificaDatos
    
    sql = "SELECT * FROM TSKardexRet WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !transid = mProps.objGNComprobante.transid
        !IdRetencion = mProps.IdRetencion
        !Debe = mProps.Debe
        !Haber = mProps.Haber
        !Base = mProps.Base
        !NumDoc = mProps.NumDoc
        !Observacion = mProps.Observacion
        !Orden = mProps.Orden
        !NumSecuencialSRI = mProps.NumSecuencialSRI
        !NumSerieEstaSRI = mProps.NumSerieEstaSRI
        !NumSeriePuntoSRI = mProps.NumSeriePuntoSRI
        !AutorizacionSRI = mProps.AutorizacionSRI
        !FechaCaducidadSRI = mProps.FechaCaducidadSRI

        
        .Update
        
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.Id = !Id             'Guarda el Id(Autonumerico)
        .Close
    End With
    Set rs = Nothing
End Sub

Friend Sub Recuperar(rs As Recordset)
    With rs
        If Not .EOF Then
            mProps.Id = !Id
            If Not IsNull(!IdRetencion) Then mProps.IdRetencion = !IdRetencion
            If Not IsNull(!CodRetencion) Then mProps.CodRetencion = !CodRetencion
            If Not IsNull(!Debe) Then mProps.Debe = !Debe
            If Not IsNull(!Haber) Then mProps.Haber = !Haber
            If Not IsNull(!Base) Then mProps.Base = !Base
            If Not IsNull(!NumDoc) Then mProps.NumDoc = !NumDoc
            If Not IsNull(!Observacion) Then mProps.Observacion = !Observacion
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            If Not IsNull(!IdCuentaActivo) Then mProps.auxIdCuentaActivo = !IdCuentaActivo
            If Not IsNull(!IdCuentaPasivo) Then mProps.auxIdCuentaPasivo = !IdCuentaPasivo
            If Not IsNull(!NumSerieEstaSRI) Then mProps.NumSerieEstaSRI = !NumSerieEstaSRI
            If Not IsNull(!NumSeriePuntoSRI) Then mProps.NumSeriePuntoSRI = !NumSeriePuntoSRI
            If Not IsNull(!NumSecuencialSRI) Then mProps.NumSecuencialSRI = !NumSecuencialSRI
            If Not IsNull(!AutorizacionSRI) Then mProps.AutorizacionSRI = !AutorizacionSRI
            If Not IsNull(!FechaCaducidadSRI) Then mProps.FechaCaducidadSRI = !FechaCaducidadSRI
            
        End If
    End With
End Sub

Friend Function Clone() As TSKardexRet
    Dim nuevo As TSKardexRet
    
    Set nuevo = New TSKardexRet
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

'Public Sub SetIdCuentaAux(auxIdCuenta As Long)
'    mProps.auxIdCuenta = auxIdCuenta
'End Sub

Friend Property Get IdRetencion() As Long
    IdRetencion = mProps.IdRetencion
End Property


Public Property Let CodRetencion(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodRetencion Then
        If Len(value) > 0 Then
            'Actualiza IdRetencion también
            sql = "SELECT IdRetencion, IdCuentaActivo, IdCuentaPasivo " & _
                  "FROM TSRetencion WHERE CodRetencion='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "TSKardexRet.CodRetencion", MSGERR_NOHAYCODIGO
            Else
                mProps.IdRetencion = rs!IdRetencion
                mProps.auxIdCuentaActivo = rs!IdCuentaActivo
                mProps.auxIdCuentaPasivo = rs!IdCuentaPasivo
            End If
            rs.Close
        Else
            mProps.IdRetencion = 0
            mProps.auxIdCuentaActivo = 0
            mProps.auxIdCuentaPasivo = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodRetencion = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodRetencion() As String
    CodRetencion = Trim$(mProps.CodRetencion)
End Property



Public Property Let Debe(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")

    If value < 0 Then
        Err.Raise ERR_INVALIDO, "TSKardexRet.Debe", MSGERR_INVALIDO
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
        Err.Raise ERR_INVALIDO, "TSKardexRet.Haber", MSGERR_INVALIDO
    ElseIf value > 0 Then
        mProps.Haber = value
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

Public Property Let Valor(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "TSKardexRet.Valor", MSGERR_INVALIDO
    End If
    
        
    If mProps.objGNComprobante.GNTrans.TSCobroAuto Then
        mProps.Debe = value
        mProps.Haber = 0
    Else
        'Si descarga de cuentas por pagar (= Retencion a proveedor)
        If Not mProps.objGNComprobante.GNTrans.TSCobro Then
            mProps.Debe = 0
            mProps.Haber = value
        'Si descarga de cuentas por cobrar (= Retencion de cliente)
        Else
            mProps.Debe = value
            mProps.Haber = 0
        End If
    End If
    
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Valor() As Currency
    'Devuelve en Moneda ex.
    Valor = (mProps.Debe + mProps.Haber) / mProps.objGNComprobante.Cotizacion("")
End Property


Public Property Let Base(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "TSKardexRet.Base", MSGERR_INVALIDO
    End If
    
    mProps.Base = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Base() As Currency
    Base = mProps.Base / mProps.objGNComprobante.Cotizacion("")
End Property


Public Property Get Porcentaje() As Currency
    If Me.Base <> 0 Then
        Porcentaje = Me.Valor * 100 / Me.Base
    End If
End Property



Public Property Let NumDoc(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "TSKardexRet.NumDoc", MSGERR_INVALIDO
    End If
    
    mProps.NumDoc = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get NumDoc() As String
    NumDoc = Trim$(mProps.NumDoc)
End Property

Public Property Let Observacion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "TSKardexRet.Observacion", MSGERR_INVALIDO
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

Friend Property Get auxIdCuentaActivo() As Long
    'Si no está asignado IdRetencion no devuelve nada
    If IdRetencion = 0 Then Exit Property
    auxIdCuentaActivo = mProps.auxIdCuentaActivo
End Property

Friend Property Get auxIdCuentaPasivo() As Long
    'Si no está asignado IdRetencion no devuelve nada
    If IdRetencion = 0 Then Exit Property
    auxIdCuentaPasivo = mProps.auxIdCuentaPasivo
End Property

Private Sub Class_Terminate()
    Set mProps.objGNComprobante = Nothing
End Sub

Public Property Let AutorizacionSRI(ByVal value As String)
    'Cuando SoloVer está establecido a True, genera error
    mProps.AutorizacionSRI = value
End Property

Public Property Get AutorizacionSRI() As String
    AutorizacionSRI = mProps.AutorizacionSRI
End Property

Public Property Let FechaCaducidadSRI(ByVal value As Date)
    'Cuando SoloVer está establecido a True, genera error
    mProps.FechaCaducidadSRI = value
End Property

Public Property Get FechaCaducidadSRI() As Date
    FechaCaducidadSRI = mProps.FechaCaducidadSRI
End Property

Public Property Let NumSerieEstaSRI(ByVal value As String)
    mProps.NumSerieEstaSRI = value
End Property

Public Property Get NumSerieEstaSRI() As String
    NumSerieEstaSRI = mProps.NumSerieEstaSRI
End Property

'jeaa 16/07/2007
Public Property Let NumSeriePuntoSRI(ByVal value As String)
    mProps.NumSeriePuntoSRI = value
End Property

Public Property Get NumSeriePuntoSRI() As String
    NumSeriePuntoSRI = mProps.NumSeriePuntoSRI
End Property

Public Property Let NumSecuencialSRI(ByVal value As String)
    mProps.NumSecuencialSRI = value
End Property

Public Property Get NumSecuencialSRI() As String
    NumSecuencialSRI = mProps.NumSecuencialSRI
End Property


