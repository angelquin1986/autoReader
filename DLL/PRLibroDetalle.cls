VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PRLibroDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
    Option Explicit


Private Type T_PROP
    Id As Long
    IdCuenta As Long
    CodCuenta As String
    auxNombreCuenta As String
    Descripcion As String
    Debe As Currency
    Haber As Currency
    Saldo As Currency
    BandIntegridad As Byte              'Indica si está integrado o no al llamar VerificaIntegridad de GNComprobante.
                                        '0=Está integrado, 1=AF,2=TS,3=IV,4=PC,5=RL
    Orden As Integer
    idGasto As Long             'jeaa 07/01/2008
    CodGasto As String          'jeaa 07/01/2008
    FechaEjec As Date
    
    objGNComprobante As GNComprobante       'Referencia al objeto padre
End Type

'Private mcolDetalleGasto As Collection      'Coleccion de detalles de familia

Private mProps As T_PROP

Public Sub VerificaDatos()

    If (IdCuenta = 0) Or (Len(mProps.CodCuenta) = 0) Then
        Err.Raise ERR_INVALIDO, "PRLibroDetalle.Verificación", _
            "Tiene que asignar una cuenta contable en detalle de asiento."
    End If
    
    If (Debe = 0 And Haber = 0) Then
        Err.Raise ERR_INVALIDO, "PRLibroDetalle.Verificación", _
            "Tiene que asignar un valor a Debe ó Haber en detalle de asiento."
    End If
    
    If (Debe <> 0 And Haber <> 0) Then
        Err.Raise ERR_INVALIDO, "PRLibroDetalle.Verificación", _
            "No puede asignar valor a Debe y Haber al mismo tiempo en detalle de asiento."
    End If
    
    If mProps.objGNComprobante.GNTrans.IVDatosAdicionales Then
        If mProps.CodCuenta <> "9999.01" Then
            If mProps.FechaEjec = "00:00:00" Then
                Err.Raise ERR_INVALIDO, "PRLibroDetalle.Verificación", _
                "Falta Ingresar la fecha de Ejecución"
                
            End If
        End If
    End If

End Sub


Friend Sub Grabar()
    Dim sql As String, rs As Recordset

    VerificaDatos
    sql = "SELECT * FROM PRLibroDetalle WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !CodAsiento = mProps.objGNComprobante.CodPRAsiento
        !IdCuenta = mProps.IdCuenta
        !Descripcion = mProps.Descripcion
        !Debe = mProps.Debe
        !Haber = mProps.Haber
        !BandIntegridad = mProps.BandIntegridad
        !Orden = mProps.Orden
        !idGasto = mProps.idGasto  ' jeaa 07/01/2008
        !FechaEjec = mProps.FechaEjec
        .Update
        
#If DAOLIB Then
        rs.Bookmark = rs.LastModified   'Para actualizar
#End If
        .Move 0                     'Para actualizar
        mProps.Id = !Id             'Guarda el Id(Autonumerico)
        .Close
        
'''        If mProps.CodCuenta <> "9999.01" Then
'''            GrabaDetalleGasto
'''        End If
        
    End With
    Set rs = Nothing
End Sub

Friend Sub Recuperar(rs As Recordset)
    With rs
        If Not .EOF Then
            mProps.Id = !Id
            mProps.IdCuenta = !IdCuenta
            If Not IsNull(!CodCuenta) Then mProps.CodCuenta = !CodCuenta
            If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
            If Not IsNull(!Debe) Then mProps.Debe = !Debe
            If Not IsNull(!Haber) Then mProps.Haber = !Haber
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            If Not IsNull(!idGasto) Then mProps.idGasto = !idGasto
            If Not IsNull(!CodGasto) Then mProps.CodGasto = !CodGasto
            If Not IsNull(!FechaEjec) Then mProps.FechaEjec = !FechaEjec
            
            If Not IsNull(!BandIntegridad) Then
                If !BandIntegridad = INTEG_AUTO Then
                    mProps.BandIntegridad = INTEG_AUTO
                    
                'Cuand es agregado por usuario
                Else
                    'No se debe recuperar de la base por que tiene que ser verificado nuevamente
                    mProps.BandIntegridad = INTEG_NADA
                End If
            End If
        End If
    End With
End Sub

Friend Function Clone() As PRLibroDetalle
    Dim nuevo As PRLibroDetalle
    
    Set nuevo = New PRLibroDetalle
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

'Llama solo desde GeneraAsiento
Friend Sub SetIdCuentaAuto(IdCta As Long)
    Dim ct As PRCuenta

    mProps.IdCuenta = IdCta
    mProps.objGNComprobante.SetModificado
    
    'Sincloniza CodCuenta también
    On Error Resume Next
    Set ct = mProps.objGNComprobante.Empresa.RecuperaPRCuenta(mProps.IdCuenta)
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
    Dim ct As PRCuenta

    'Recupera el registro por el id
    Set ct = mProps.objGNComprobante.Empresa.RecuperaPRCuenta(value)
    If Not (ct Is Nothing) Then
        'Sincloniza CodCuenta también
        mProps.CodCuenta = ct.CodCuenta
        mProps.auxNombreCuenta = ct.NombreCuenta
    Else
        'Si no recupera, genera un error
        Err.Raise ERR_NOHAYCODIGO, "PRLibroDetalle_IdCuenta", MSGERR_NOHAYCODIGO
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
        sql = "SELECT IdCuenta,NombreCuenta FROM PRCuenta WHERE CodCuenta='" & value & "'"
        Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "PRLibroDetalle_CodCuenta", MSGERR_NOHAYCODIGO
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
    Dim ct As PRCuenta
    
    If Len(mProps.auxNombreCuenta) > 0 Then
        auxNombreCuenta = mProps.auxNombreCuenta
    Else
        On Error Resume Next
        Set ct = mProps.objGNComprobante.Empresa.RecuperaPRCuenta(mProps.IdCuenta)
        If Err.Number = 0 Then auxNombreCuenta = ct.NombreCuenta
        Err.Clear
        On Error GoTo 0
    End If
    Set ct = Nothing
End Property


Public Property Let Descripcion(ByVal value As String)
    If value = mProps.Descripcion Then Exit Property
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PRLibroDetalle_Descripcion", MSGERR_INVALIDO
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
    mProps.FechaEjec = Date
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
            Err.Raise ERR_NOHAYCODIGO, "PRLibroDetalle_CodGasto", MSGERR_NOHAYCODIGO
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



Public Property Let Sadlo(ByVal value As Currency)
    value = value * mProps.objGNComprobante.Cotizacion("")
    mProps.Saldo = value

End Property

Public Property Get Saldo() As Currency
    'Devuelve en Moneda Ex.
    Saldo = mProps.Saldo / mProps.objGNComprobante.Cotizacion("")
End Property

Public Function CalculaSaldo(ByVal CodCuenta As String) As Currency
    Dim sql As String, rs As Recordset, Trans As String, ValorTotal As Currency, ValorComp As Currency
    'Cuando cambia el valor
        'Actualiza IdGasto también
        sql = " SELECT"
        sql = sql & " case when isnull(sum(ctl.Debe) / 1, 0) > isnull(sum(ctl.Haber) / 1, 0) then "
        sql = sql & " (isnull(sum(ctl.Debe) / " & mProps.objGNComprobante.Cotizacion("") & ", 0) - isnull(sum(ctl.Haber) / " & mProps.objGNComprobante.Cotizacion("") & ", 0))*-1  else"
        sql = sql & " (isnull(sum(ctl.Debe) / " & mProps.objGNComprobante.Cotizacion("") & ", 0) - isnull(sum(ctl.Haber) / " & mProps.objGNComprobante.Cotizacion("") & ", 0))  end as saldo"
        sql = sql & " FROM PRCuenta ct"
        sql = sql & " INNER JOIN PRLibroDetalle ctl"
        sql = sql & " INNER JOIN GnComprobante gc"
        sql = sql & " ON gc.CodPRAsiento = ctl.CodAsiento"
        sql = sql & " ON ct.IdCuenta = ctl.IdCuenta"
        sql = sql & " where gc.estado <> 3  and ct.codcuenta='" & CodCuenta & "'"
        sql = sql & " and gc.transid<>" & mProps.objGNComprobante.transid
        
'        If Len(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("TransVerificaSaldoPresupuesto")) > 0 Then
'            Trans = PreparaCadena(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("TransVerificaSaldoPresupuesto"))
'        End If
        
        If mProps.objGNComprobante.GNTrans.IVTipoTransCompromete Then
            If Len(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("TransVerificaSaldoPresupuesto")) > 0 Then
                Trans = PreparaCadena(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("TransVerificaSaldoPresupuesto"))
            End If
        Else
            If Len(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("TransVerificaSaldoPresupuestoEject")) > 0 Then
                Trans = PreparaCadena(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("TransVerificaSaldoPresupuestoEject"))
            End If
        
        End If
        
        
        sql = sql & " and codtrans in (" & Trans & ")"
        
       
       ValorTotal = 0
        Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
        If rs.RecordCount > 0 Then
            ValorTotal = rs.Fields("saldo") * -1
        End If
        
        If mProps.objGNComprobante.CodTrans = "RI" Or mProps.objGNComprobante.CodTrans = "DRI" Then
         sql = " SELECT"
         sql = sql & " case when isnull(sum(ctl.Debe) / 1, 0) > isnull(sum(ctl.Haber) / 1, 0) then "
         sql = sql & " (isnull(sum(ctl.Debe) / " & mProps.objGNComprobante.Cotizacion("") & ", 0) - isnull(sum(ctl.Haber) / " & mProps.objGNComprobante.Cotizacion("") & ", 0))*-1  else"
         sql = sql & " (isnull(sum(ctl.Debe) / " & mProps.objGNComprobante.Cotizacion("") & ", 0) - isnull(sum(ctl.Haber) / " & mProps.objGNComprobante.Cotizacion("") & ", 0))  end as saldo"
         sql = sql & " FROM PRCuenta ct"
         sql = sql & " INNER JOIN PRLibroDetalle ctl"
         sql = sql & " INNER JOIN GnComprobante gc"
         sql = sql & " INNER JOIN gnrequerimiento GNR ON GC.TRANSID=GNR.TRANSID"
         sql = sql & " ON gc.CodPRAsiento = ctl.CodAsiento"
         sql = sql & " ON ct.IdCuenta = ctl.IdCuenta"
         sql = sql & " where gc.estado <> 3  and ct.codcuenta='" & CodCuenta & "'"
         sql = sql & " and gnr.BandCertifica=1"
         sql = sql & " and  gc.codtrans not in('PRICT','PRPLM')"
         
        ValorComp = 0
         Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
         If rs.RecordCount > 0 Then
             ValorComp = rs.Fields("saldo") * -1
         End If
        Else
            ValorComp = 0
        End If
        
        
        
'        If rs.EOF Then
'            Err.Raise ERR_NOHAYCODIGO, "PRLibroDetalle_CodGasto", MSGERR_NOHAYCODIGO
'            Exit Function
'        Else
            CalculaSaldo = (ValorTotal - ValorComp) * -1
'        End If
        rs.Close
        Set rs = Nothing
End Function

Public Function VerificaSaldo(ByVal CodCuenta As String, ByVal BandComprometido As Boolean, ByVal Valor As Currency, ByRef SaldoCuenta As Currency) As Boolean
    Dim sql As String, rs As Recordset, Trans As String
    Dim ValorTotal As Currency, ValorComp As Currency, Saldo As Currency, TipoCuenta As Integer
    'Cuando cambia el valor
        'Actualiza IdGasto también
''''        sql = " SELECT"
''''        sql = sql & " MAX(CT.TIPOCUENTA) AS TC, case when MAX(CT.TIPOCUENTA)=1 THEN" 'snull(sum(ctl.Debe) / 1, 0) > isnull(sum(ctl.Haber) / 1, 0) then "
''''        sql = sql & " (isnull(sum(ctl.Debe) / " & mProps.objGNComprobante.Cotizacion("") & ", 0) - isnull(sum(ctl.Haber) / " & mProps.objGNComprobante.Cotizacion("") & ", 0)) *-1  else"
''''        sql = sql & " (isnull(sum(ctl.Debe) / " & mProps.objGNComprobante.Cotizacion("") & ", 0) - isnull(sum(ctl.Haber) / " & mProps.objGNComprobante.Cotizacion("") & ", 0))  end as saldo"
''''        sql = sql & " FROM PRCuenta ct"
''''        sql = sql & " INNER JOIN PRLibroDetalle ctl"
''''        sql = sql & " INNER JOIN GnComprobante gc"
''''        sql = sql & " ON gc.CodPRAsiento = ctl.CodAsiento"
''''        sql = sql & " ON ct.IdCuenta = ctl.IdCuenta"
''''        sql = sql & " where gc.estado <> 3 "
''''
''''        If mProps.objGNComprobante.GNTrans.IVTipoTransCompromete Then
''''            If Len(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("TransVerificaSaldoPresupuesto")) > 0 Then
''''                Trans = PreparaCadena(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("TransVerificaSaldoPresupuesto"))
''''            End If
''''        Else
''''            If Len(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("TransVerificaSaldoPresupuestoEject")) > 0 Then
''''                Trans = PreparaCadena(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("TransVerificaSaldoPresupuestoEject"))
''''            End If
''''
''''        End If
''''        sql = sql & " and codtrans in (" & Trans & ")"
''''
''''
''''
''''        sql = sql & " and ct.codcuenta='" & CodCuenta & "'"
''''        sql = sql & " and gc.transid<>" & mProps.objGNComprobante.TransID
'        If Not BandComprometido Then
            sql = " SELECT isnull(MAX(CT.TIPOCUENTA),0) AS TC, "
            sql = sql & " case when isnull(sum(ctl.Debe) / 1, 0) > isnull(sum(ctl.Haber) / 1, 0) then "
            sql = sql & " (isnull(sum(ctl.Debe) / " & mProps.objGNComprobante.Cotizacion("") & ", 0) - isnull(sum(ctl.Haber) / " & mProps.objGNComprobante.Cotizacion("") & ", 0))*-1  else"
            sql = sql & " (isnull(sum(ctl.Debe) / " & mProps.objGNComprobante.Cotizacion("") & ", 0) - isnull(sum(ctl.Haber) / " & mProps.objGNComprobante.Cotizacion("") & ", 0))  end as saldo"
            sql = sql & " FROM PRCuenta ct"
            sql = sql & " INNER JOIN PRLibroDetalle ctl"
            sql = sql & " INNER JOIN GnComprobante gc"
            sql = sql & " ON gc.CodPRAsiento = ctl.CodAsiento"
            sql = sql & " ON ct.IdCuenta = ctl.IdCuenta"
            sql = sql & " where gc.estado <> 3  and ct.codcuenta='" & CodCuenta & "'"
            sql = sql & " and gc.transid<>" & mProps.objGNComprobante.transid
            
            If mProps.objGNComprobante.GNTrans.IVTipoTransCompromete Then
                If Len(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("TransVerificaSaldoPresupuesto")) > 0 Then
                    Trans = PreparaCadena(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("TransVerificaSaldoPresupuesto"))
                End If
            Else
                If Len(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("TransVerificaSaldoPresupuestoEject")) > 0 Then
                    Trans = PreparaCadena(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("TransVerificaSaldoPresupuestoEject"))
                End If
            
            End If
            
            
            sql = sql & " and codtrans in (" & Trans & ")"
            
            
            
            ValorTotal = 0
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.RecordCount > 0 Then
                ValorTotal = rs.Fields("saldo") * -1
                TipoCuenta = rs.Fields("TC")
            End If
'        Else
'            sql = " SELECT CT.TIPOCUENTA AS TC "
'            sql = sql & " FROM PRCuenta ct"
'            sql = sql & " where ct.codcuenta='" & CodCuenta & "'"
'            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
'
'            ValorTotal = 0
'            TipoCuenta = rs.Fields("TC")
'        End If

        If BandComprometido Then
            sql = " SELECT"
            sql = sql & " case when isnull(sum(ctl.Debe) / 1, 0) > isnull(sum(ctl.Haber) / 1, 0) then "
            sql = sql & " (isnull(sum(ctl.Debe) / " & mProps.objGNComprobante.Cotizacion("") & ", 0) - isnull(sum(ctl.Haber) / " & mProps.objGNComprobante.Cotizacion("") & ", 0))*-1  else"
            sql = sql & " (isnull(sum(ctl.Debe) / " & mProps.objGNComprobante.Cotizacion("") & ", 0) - isnull(sum(ctl.Haber) / " & mProps.objGNComprobante.Cotizacion("") & ", 0))  end as saldo"
            sql = sql & " FROM PRCuenta ct"
            sql = sql & " INNER JOIN PRLibroDetalle ctl"
            sql = sql & " INNER JOIN GnComprobante gc"
            sql = sql & " INNER JOIN gnrequerimiento GNR ON GC.TRANSID=GNR.TRANSID"
            sql = sql & " ON gc.CodPRAsiento = ctl.CodAsiento"
            sql = sql & " ON ct.IdCuenta = ctl.IdCuenta"
            sql = sql & " where gc.estado <> 3  and ct.codcuenta='" & CodCuenta & "'"
            sql = sql & " and gnr.BandCertifica=1"
            sql = sql & " and  gc.codtrans not in('PRICT','PRPLM')"
            
           ValorComp = 0
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.RecordCount > 0 Then
                ValorComp = rs.Fields("saldo") * -1
            End If
        Else
            ValorComp = 0
        End If
        
        If BandComprometido Then
            Saldo = (ValorTotal - ValorComp) * -1
        Else
            Saldo = (ValorTotal - ValorComp) * -1
        End If

'        Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
'        If rs.EOF Then
'            Err.Raise ERR_NOHAYCODIGO, "PRLibroDetalle_CodGasto", MSGERR_NOHAYCODIGO
'            Exit Function
'        Else
            
            
            If TipoCuenta = 1 Then
                SaldoCuenta = (Saldo)
                If Saldo - Valor >= 0 Then
                    VerificaSaldo = True
                Else
                    VerificaSaldo = False
                End If
            
            Else
                SaldoCuenta = Saldo
                If SaldoCuenta < 0 Then
                    If (SaldoCuenta + Valor) <= 0 Then
                        VerificaSaldo = True
                        SaldoCuenta = Abs(SaldoCuenta)
                    Else
                        VerificaSaldo = False
                        SaldoCuenta = Abs(SaldoCuenta)
                    End If
                Else
                        VerificaSaldo = False
                        SaldoCuenta = SaldoCuenta * -1
                End If
            End If
'        End If
        rs.Close
        Set rs = Nothing
End Function

Public Property Let FechaEjec(ByVal value As Date)
    mProps.FechaEjec = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get FechaEjec() As Date
    FechaEjec = mProps.FechaEjec
End Property



'''
'''
'''
'''Public Sub GrabaDetalleGasto()
'''    Dim rs1 As Recordset, numReg As Long, i As Long
'''    Dim obj As PRDetalleGasto, sql As String, band As Boolean
'''    Dim NumHijos As Integer, CountNumHijos As Integer, cuenta As PRCuenta
'''
'''    On Error GoTo CapturaError
'''
'''
'''
'''
'''
'''
'''
'''
'''
'''    'Elimina  todos  los registros
'''    sql = "Delete Ivcd " & _
'''          "From prlibrodetalle Ivk " & _
'''          "INNER JOIN PRDetalleGasto Ivcd ON IVk.Id = Ivcd.IdPRLibroDetalleRef "
'''          sql = sql & " Where Ivcd.IdPRLibroDetalleRef = " & mProps.Id
'''
'''    mProps.objGNComprobante.Empresa.EjecutarSQL sql, numReg
'''    'guarda  todos  los registros
'''    band = True
'''
'''    If Not mcolDetalleGasto Is Nothing Then
'''        sql = "SELECT * FROM PRDetalleGasto WHERE 1=0"
'''        Set rs1 = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
'''        NumHijos = mcolDetalleGasto.Count
'''        CountNumHijos = 0
'''        For i = 1 To Me.NumFamiliaDetalle 'mcolDetalleGasto.Count
'''            If CountNumHijos < NumHijos Then
'''                CountNumHijos = CountNumHijos + 1
'''                Set obj = mcolDetalleGasto.item(i)
'''''''                If obj.idCuenta = 0 Then
'''''''                        obj.Recuperar obj.CodCuenta
'''''''                        Set item = mProps.objGNComprobante.Empresa.RecuperaIVInventario(obj.CodInventario)
'''''''                        If verificacionDuiplicado(obj.IdInventario) Then
'''''''                            RemoveDetalleGasto i
'''''''                            i = i - 1
'''''''                        End If
'''''''                End If
'''                If band Then
'''                    rs1.AddNew
'''                    rs1!TransID = mProps.objGNComprobante.TransID
'''                    rs1!IdCuenta = mProps.IdCuenta
'''                    rs1!IdPRLibroDetalleRef = mProps.Id
'''                    rs1!Valor = obj.Valor
'''                    rs1.Update
'''                End If
'''            End If
'''            band = True
'''            Set cuenta = Nothing
'''        Next i
'''    End If
'''
'''
'''    Set rs1 = Nothing
'''    Set obj = Nothing
'''    Exit Sub
'''CapturaError:
'''    MsgBox Err.Description
'''    Set rs1 = Nothing
'''    Set obj = Nothing
'''    Exit Sub
'''End Sub
'''
'''
'''Public Property Get NumFamiliaDetalle() As Long
'''    If mcolDetalleGasto Is Nothing Then
'''        NumFamiliaDetalle = 0
'''    Else
'''        NumFamiliaDetalle = mcolDetalleGasto.Count
'''    End If
'''    'If ObtieneUltimo(False) Then PrecioUltimoEgreso = mUEgreso.Precio
'''End Property
'''
'''Private Sub CargaDetalleGasto()
'''    Dim rs As Recordset, sql As String, obj As PRDetalleGasto
'''    Dim i As Long
'''    Set mcolDetalleGasto = New Collection
'''        sql = "Select prd.Idcuenta, prc.Codcuenta, prc.nombrecuenta, "
'''        sql = sql & " prd.valor , prd.orden "
'''        sql = sql & " From  prlibrodetalle prl "
'''        sql = sql & " INNER JOIN PRDetalleGasto prd "
'''        sql = sql & " INNER JOIN Prcuenta prc "
'''        sql = sql & " ON prd.Idcuenta = prc.Idcuenta "
'''        sql = sql & " ON prl.Id = prd.IdPRLibroDetalleRef"
'''        sql = sql & " Where prd.IdPRLibroDetalleRef = " & mProps.Id
'''        sql = sql & " order by prd.orden"
'''
'''    Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
'''    With rs
'''        Do Until .EOF
'''            i = AddDetalleGasto
'''            Set obj = mcolDetalleGasto.item(i)
'''                  obj.IdCuenta = !IdCuenta
'''                  obj.CodCuenta = !CodCuenta
'''                 obj.Valor = !Valor
'''                obj.Orden = !Orden
'''           .MoveNext
'''        Loop
'''        .Close
'''    End With
'''    Set rs = Nothing
'''    Set obj = Nothing
'''End Sub
'''
'''Public Function RecuperaDetalleGasto(ix As Long) As PRDetalleGasto
'''    Set RecuperaDetalleGasto = mcolDetalleGasto.item(ix)
'''End Function
'''
'''
'''Public Function AddDetalleGasto() As Long
'''    Dim nuevo As PRDetalleGasto
'''
'''    If mcolDetalleGasto Is Nothing Then CargaDetalleGasto
'''    Set nuevo = New PRDetalleGasto
'''    Set nuevo.PRLibroDetalle = Me  'Me  olvide  del padre
'''    'Proporciona valores predeterminados
'''
'''
'''    mcolDetalleGasto.Add nuevo
'''    AddDetalleGasto = mcolDetalleGasto.Count    'Devuelve el indice de nuevo elemento
'''    Set nuevo = Nothing
'''
'''
'''End Function
'''
'''Public Sub RemoveDetalleGasto(ByVal ix As Long, Optional obj As PRDetalleGasto)
'''    If mcolDetalleGasto Is Nothing Then CargaDetalleGasto
'''
'''    'Si está especificado un objeto
'''    If Not (obj Is Nothing) Then
'''        'Busca el indice del objeto
'''        For ix = 1 To mcolDetalleGasto.Count
'''            If mcolDetalleGasto.item(ix) Is obj Then
'''                mcolDetalleGasto.Remove ix
'''                Exit For
'''            End If
'''        Next ix
'''        Debug.Assert True     'No se puede dar
'''    Else
'''        mcolDetalleGasto.Remove ix
'''    End If
'''
'''End Sub
'''
'''Private Function verificacionDuiplicado(ByVal IdInventario As String) As Boolean
'''    Dim rs As Recordset, sql As String
'''    verificacionDuiplicado = False
'''    'Recupera los detalles de existencia
'''    sql = "SELECT IdInventario FROM ivmateria where idinventario =" & IdInventario
'''    Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
'''    If rs.RecordCount > 0 Then
'''        verificacionDuiplicado = True
'''    End If
'''    Set rs = Nothing
'''End Function
'''
'''
