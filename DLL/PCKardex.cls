VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCKardex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"PCKardex"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Type T_PROP
    id As Long
    IdProvcli As Long
    CodProvcli As String
    
    IDForma As Long
    CodForma As String
    
    IdAsignado As Long
    NumLetra As String
    Debe As Currency
    Haber As Currency
    FechaEmision As Date
    FechaVenci As Date
    Observacion As String
    Orden As Integer
    Guid As String                          '*** MAKOTO 16/mar/01 Agregado para Exportación/Importación
    objGNComprobante As GNComprobante           'Referencia al objeto padre
    auxIdCuenta As Long          'Auxiliar para almacenar Id de cuenta de proveedor/cliente. (No se graba en la base)
    auxIdCuenta2 As Long          '*** MAKOTO 12/feb/01 Agregado
    BandPorCobrar As Boolean          'True:Por cobrar, False:Por pagar (Esto no se guarda en la base)
    NoVerificarSaldo As Boolean
    CodGrupo As String
    IdVendedor As Long 'jeaa 25/05/2007
    CodVendedor As String
    NombreVendedor As String
    IdCobrador As Long 'jeaa 25/05/2007
    CodCobrador As String
    NombreCobrador As String
    BandNoAfectaPckardex As Boolean 'AUC para saber si grabo on o en Pckardex temporal no se guarda en la base
    IdBanco As Long
    CodBanco As String
    IdTarjeta As Long
    CodTarjeta As String
    NumCuenta  As String
    numCheque  As String
    TitularCta  As String
    idElemento As Long
    auxIdCuentaXIII As Long          'Auxiliar para almacenar Id de cuenta del empleado xiii
    auxIdCuentaXIV As Long          'Auxiliar para almacenar Id de cuenta del empleado xiV
    auxIdCuentaVac As Long          'Auxiliar para almacenar Id de cuenta del empleado Vac
    auxIdCuentaNetoRol As Long          'Auxiliar para almacenar Id de cuenta del empleado Vac
    RubroRol As String
    EsRubroRol As Boolean
    ValorCapital As Currency
    ValorInteres As Currency
    OrdenCuota As Integer
    ValorCapital1 As Currency
    ValorInteres1 As Currency
    TasaInteres1 As Currency
    ValorCapital2 As Currency
    ValorInteres2 As Currency
    TasaInteres2 As Currency
    ValorCapital3 As Currency
    ValorInteres3 As Currency
    TasaInteres3 As Currency
    CodEmpleado As String
    Codelemento As String
    auxIdCuenta3 As Long
    auxIdCuenta4 As Long
    Porc As Currency
    DebePar As Currency
    DebeImPar As Currency
    CodFormaPar As String
    CodFormaImPar As String
    FechaVenciPar As Date
    FechaVenciImPar As Date
    FechaEmisionPar As Date
    FechaEmisionImPar As Date
    OrdenPar As Integer
    OrdenImPar As Integer
    Cuota As Currency
    idGestionc As Long
    IdEquipo As Long
    CodEquipo As String
    IDFormaSRI As Long
    BandPosf As Boolean
    
End Type

Private mProps As T_PROP
Private mcolTransIDsPorPagar As Collection 'AUC 06/12/06


Public Sub VerificaDatos()
    Dim pck As PCKardex, ix As Long, pcd As PCDocAsignado, v As Currency
    Dim tsf As TSFormaCobroPago, IdProvcli As Long
    Dim msg As String, nump As Integer, pc As PCProvCli

    'Verifica datos
    If mProps.Debe = 0 And mProps.Haber = 0 Then
        Err.Raise ERR_INVALIDO, "Pckardex.VerificaDatos", _
            "Se intentó grabar sin valor de Debe ni de Haber " & _
                "en detalle de Documentos de Proveedor/Cliente."
    End If
    If (mProps.Debe <> 0) And (mProps.Haber <> 0) Then
        Err.Raise ERR_INVALIDO, "Pckardex.VerificaDatos", _
            "Se intentó grabar con Debe y Haber al mismo tiempo " & _
                "en detalle de Documentos de de Proveedor/Cliente."
    End If
    If mProps.IdAsignado = 0 And mProps.IDForma = 0 Then
        Err.Raise ERR_INVALIDO, "Pckardex_VerificaDatos", _
            "Debe seleccionar código de Forma de documento " & _
                "en detalle de Documentos de Proveedor/Cliente."
    End If
    If mProps.IdProvcli = 0 Then
        Err.Raise ERR_INVALIDO, "Pckardex.VerificaDatos", _
            "Debe seleccionar un cliente o un proveedor " & _
                "en detalle de Documentos de Proveedor/Cliente."
    End If
    If Len(mProps.NumLetra) = 0 Then
        Err.Raise ERR_INVALIDO, "Pckardex.VerificaDatos", _
            "Debe ingresar un número de documento " & _
                "en detalle de Documentos de Proveedor/Cliente."
    End If
    If mProps.FechaEmision > mProps.FechaVenci Then
        Err.Raise ERR_INVALIDO, "Pckardex.VerificaDatos", _
            "Fecha de vencimiento no puede ser menor a la fecha de emisión " & _
                "en detalle de Documentos de Proveedor/Cliente."
    End If
    
    If mProps.objGNComprobante.GNTrans.CodPantalla <> "TSCC" And _
         mProps.objGNComprobante.GNTrans.CodPantalla <> "TSCCA" And _
         mProps.objGNComprobante.GNTrans.CodPantalla <> "TSIEE" Then
        'Verifica si la transaccion tiene diferente cliente en gncomprobant y Pckardex
        'And (mProps.objGNComprobante.GNTrans.CodPantalla <> "TSIEE")
        If Not (mProps.objGNComprobante.GNTrans.IVProvCliPorFila) And Not (mProps.objGNComprobante.GNTrans.TSProvCliPorFila) Then
            If mProps.objGNComprobante.GNTrans.ClienteVisible And Not mProps.objGNComprobante.GNTrans.ProveedorVisible Then
                If mProps.objGNComprobante.IdClienteRef <> 0 Then
                    mProps.IdProvcli = mProps.objGNComprobante.IdClienteRef
                End If
             ElseIf mProps.objGNComprobante.GNTrans.ProveedorVisible And Not mProps.objGNComprobante.GNTrans.ClienteVisible Then
                If mProps.objGNComprobante.IdProveedorRef <> 0 Then
                    If mProps.IdAsignado = 0 Then
                        mProps.IdProvcli = mProps.objGNComprobante.IdProveedorRef
                    End If
                End If
            ElseIf mProps.objGNComprobante.GNTrans.ProveedorVisible And mProps.objGNComprobante.GNTrans.ClienteVisible Then
                If mProps.objGNComprobante.IdClienteRef <> 0 And mProps.objGNComprobante.IdProveedorRef = 0 Then
                    mProps.IdProvcli = mProps.objGNComprobante.IdClienteRef
                ElseIf mProps.objGNComprobante.IdClienteRef = 0 And mProps.objGNComprobante.IdProveedorRef <> 0 Then
                    mProps.IdProvcli = mProps.objGNComprobante.IdProveedorRef
                End If
             End If
        End If
    End If
    '.FechaEmision > mProps.FechaVenci Then
    
    ' si es modificacion y la transaccion es varios cliente
    If Not mProps.objGNComprobante.EsNuevo Then
        If (mProps.objGNComprobante.GNTrans.IVProvCliPorFila) Or (mProps.objGNComprobante.GNTrans.TSProvCliPorFila) Then
            
'            For ix = 1 To mProps.objGNComprobante.CountPckardex
                'IdProvCli = ObtieneIdCliente(mProps.objGNComprobante.Pckardex(ix).IdAsignado)
                IdProvcli = ObtieneIdCliente(Me.IdAsignado)
                If IdProvcli <> 0 Then
                    If mProps.objGNComprobante.IdClienteRef = 0 And mProps.objGNComprobante.IdProveedorRef = 0 Then
                        'If mProps.objGNComprobante.Pckardex(ix).IdProvCli <> IdProvCli Then
                        If Me.IdProvcli <> IdProvcli Then
                            mProps.IdProvcli = IdProvcli
                        End If
                    End If
                End If
        End If
    End If
    
    
    'jeaa 10/05/2005
    If mProps.objGNComprobante.GNTrans.VerificaFechas Then
        If mProps.objGNComprobante.FechaTrans <> mProps.FechaEmision Then
            Err.Raise ERR_INVALIDO, "Pckardex.VerificaDatos", _
                "Fecha de emision del comprobante no puede ser diferente a la fecha" & Chr(13) & "de emisión en el detalle de Documentos de Proveedor/Cliente."
        End If
    End If
    
    
    
    'Verifica si repite el mismo doc. asignado en el comprobante
    If mProps.IdAsignado And mProps.objGNComprobante.EsNuevo Then
        If mProps.objGNComprobante.GNTrans.CodPantalla <> "TSIDEP" Then
            For ix = 1 To mProps.objGNComprobante.CountPCKardex
                Set pck = mProps.objGNComprobante.PCKardex(ix)
                If (pck.IdAsignado = mProps.IdAsignado) And (mProps.IdAsignado <> -1) And (Not (pck Is Me)) Then
                    msg = "Existen documentos repetidos para el mismo comprobante." & vbCr & vbCr
                    Set pcd = RecuperaPCDocAsignado
                    If pcd.id Then msg = msg & pcd.Trans & " " & pcd.Doc & ": "
                    msg = msg & Format(Me.Debe + Me.Haber, Me.GNComprobante.FormatoMoneda) & " " & _
                                Me.GNComprobante.CodMoneda & " y " & _
                                Format(pck.Debe + pck.Haber, pck.GNComprobante.FormatoMoneda) & _
                                " " & pck.GNComprobante.CodMoneda
                    Err.Raise ERR_INVALIDO, "Pckardex_VerificaDatos", msg
                End If
                If pck.IdAsignado = mProps.IdAsignado Then
                    v = v + pck.Debe + pck.Haber
                End If
                If mProps.objGNComprobante.GNTrans.TSPago Then
                    If Len(pck.CodVendedor) = 0 Then
                        Set pc = mProps.objGNComprobante.Empresa.RecuperaPCProvCli(pck.CodProvcli)
                        If Len(pc.CodVendedor) > 0 Then
                            pck.CodVendedor = pc.CodVendedor
                        End If
                    End If
                End If
            Next ix
        End If
        Set pc = Nothing
        'Si el valor está más que el saldo del documento asignado
        Set pcd = RecuperaPCDocAsignado
        If v > pcd.Saldo And pcd.id Then
            Err.Raise ERR_INVALIDO, "Pckardex_VerificaDatos", _
                "Valor que Cobra/Paga es mayor al saldo del documento " & vbCr & vbCr & _
                pcd.NombreProvCli & " " & pcd.Trans & " " & pcd.Doc & vbCr & _
                "Saldo es de: " & pcd.Saldo & " " & pcd.CodMoneda & vbCr & vbCr & _
                "Cobro/Pago de : " & v & " " & mProps.objGNComprobante.CodMoneda
        End If
        Set pck = Nothing
        Set pcd = Nothing
    End If
    If mProps.objGNComprobante.GNTrans.IVControlaPrecioxFormaCobro Then
        Set tsf = mProps.objGNComprobante.Empresa.RecuperaTSFormaCobroPago(mProps.CodForma)
        If Not tsf Is Nothing Then
                If tsf.ControlPrecios Then
                        For ix = 1 To mProps.objGNComprobante.CountIVKardex
                            nump = BuscaNumeroPrecio(ix)
                            If nump <> 0 Then
                                If Mid$(tsf.ListaPrecios, nump, 1) = "0" Then
                                    Err.Raise ERR_INVALIDO, "GNComprobante.RemovePckardex", _
                                    "Con la forma de cobro " & mProps.CodForma & " no puede seleccionar el precio " & nump & Chr(13) & _
                                    " en la fila " & ix & " código Item: " & mProps.objGNComprobante.IVKardex(ix).CodInventario
                                End If
                            End If
                        Next ix
                End If
        End If
    End If
      If Len(mProps.CodEmpleado) > 0 And InStr(1, mProps.objGNComprobante.CodTrans, "ROL") > 0 Then
        If mProps.idElemento = 0 Then
            Err.Raise ERR_INVALIDO, "Pckardex.VerificaDatos", _
            "Se intentó grabar sin rubro/rol " & _
                "en detalle de Documentos de Empleado."
        End If
    End If
       If Me.GNComprobante.GNTrans.CodPantalla = "IVBQDCRC" Then
            mProps.NumLetra = Me.GNComprobante.CodTrans & Left(Me.GNComprobante.CodVendedor, 2) & Left(Me.GNComprobante.CodCobrador, 2) & Weekday(Me.GNComprobante.FechaFuente, vbMonday) & Me.GNComprobante.NumTrans & "-"
        End If
           
End Sub

Friend Sub Grabar()
    Dim sql As String
    Dim id_antes As Long, rs As Recordset
    Dim id_antesC As Long
    id_antes = mProps.id        'Guarda el id anterior
    id_antesC = mProps.idGestionc
    VerificaDatos
    'Graba en Pckardex
    sql = "SELECT * FROM Pckardex WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !transid = mProps.objGNComprobante.transid
        !IdProvcli = mProps.IdProvcli
        !IDForma = mProps.IDForma
        !IdAsignado = mProps.IdAsignado
        !NumLetra = mProps.NumLetra
        !Debe = mProps.Debe
        !Haber = mProps.Haber
        !FechaEmision = mProps.FechaEmision
        !FechaVenci = mProps.FechaVenci
        !Observacion = mProps.Observacion
        !Orden = mProps.Orden
        !IdVendedor = mProps.IdVendedor
        !IdCobrador = mProps.IdCobrador
        If Len(mProps.Guid) > 0 Then !Guid = mProps.Guid        '*** MAKOTO 16/mar/01 Agregado
        !IdBanco = mProps.IdBanco   'jeaa 20/07/2009
        !IdTarjeta = mProps.IdTarjeta
        !NumCuenta = mProps.NumCuenta
        !numCheque = mProps.numCheque
        !TitularCta = mProps.TitularCta
        !idElemento = mProps.idElemento
        !IdEquipo = mProps.IdEquipo
        !IDFormaSRI = mProps.IDFormaSRI
        !BandPosf = mProps.BandPosf
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.id = !id             'Guarda el Id(Autonumerico)
        .Close
    End With
    
    'Si está modificando
    If id_antes <> 0 Then
        'Actualiza IdAsignado de Pckardex relacionado
        '  para que apunte de nuevo a éste registro.
        sql = "UPDATE Pckardex SET IdAsignado=" & mProps.id & _
              " WHERE IdAsignado=" & id_antes
              
        #If DAOLIB Then
                mProps.objGNComprobante.Empresa.Database.Execute sql, dbFailOnError
        #Else
                mProps.objGNComprobante.Empresa.Coneccion.Execute sql
        #End If
              
              
        sql = "UPDATE PckardexAmort SET Idpckardex=" & mProps.id & _
              " WHERE idPckardex=" & id_antes
        #If DAOLIB Then
                mProps.objGNComprobante.Empresa.Database.Execute sql, dbFailOnError
        #Else
                mProps.objGNComprobante.Empresa.Coneccion.Execute sql
        #End If

        sql = "UPDATE Pckardexchp SET IdAsignadoPCK=" & mProps.id & _
              " WHERE IdAsignadoPCK=" & id_antes
        #If DAOLIB Then
                mProps.objGNComprobante.Empresa.Database.Execute sql, dbFailOnError
        #Else
                mProps.objGNComprobante.Empresa.Coneccion.Execute sql
        #End If
       UpdateGnCall id_antes
    End If
    If id_antesC <> 0 Then
        UpdateGnCall id_antesC
        End If
    'Si asigna Id al campo IdAsignado de otros Pckardexs en la misma transacción
    ' para descargar cobros/pagos al contado
        AsignaIDPCK
        GrabarPCFinanciamiento
        GrabarPCKardexAmort
        
    Set rs = Nothing
End Sub

'Asigna IdAsignado en los Pckardex que está en el mismo comprobante y que tengan ahora -1 como IdAsignado
' Es para descargar cobro/pago al contado
Private Sub AsignaIDPCK()
    Dim i As Long, pck As PCKardex
    
    For i = 1 To mProps.objGNComprobante.CountPCKardex
        Set pck = mProps.objGNComprobante.PCKardex(i)
        If (pck.IdAsignado = -1) And _
           (pck.CodForma = Me.CodForma) And _
           ((pck.Debe + pck.Haber) = (Me.Debe + Me.Haber)) Then
            pck.IdAsignado = mProps.id
            pck.CodProvcli = mProps.CodProvcli
            pck.CodForma = mProps.CodForma
            pck.NumLetra = mProps.NumLetra
            Exit For
        End If
    Next i
    Set pck = Nothing
End Sub

Friend Sub Recuperar(rs As Recordset)
    With rs
        If Not .EOF Then
            mProps.id = !id
            If Not IsNull(!IdProvcli) Then mProps.IdProvcli = !IdProvcli
            'If Not IsNull(!CodProvCli) Then mProps.CodProvCli = !CodProvCli
            If !idElemento = 0 Then
                If Not IsNull(!CodProvcli) Then mProps.CodProvcli = !CodProvcli
            Else
                If Not IsNull(!CodProvcli) Then mProps.CodEmpleado = !CodProvcli
            End If
            If Not IsNull(!IDForma) Then mProps.IDForma = !IDForma
            If Not IsNull(!CodForma) Then mProps.CodForma = !CodForma
            If Not IsNull(!IdAsignado) Then mProps.IdAsignado = !IdAsignado
            If Not IsNull(!NumLetra) Then mProps.NumLetra = !NumLetra
            If Not IsNull(!Debe) Then mProps.Debe = !Debe
            If Not IsNull(!Haber) Then mProps.Haber = !Haber
            If Not IsNull(!FechaEmision) Then mProps.FechaEmision = !FechaEmision
            If Not IsNull(!FechaVenci) Then mProps.FechaVenci = !FechaVenci
            If Not IsNull(!Observacion) Then mProps.Observacion = !Observacion
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            If Not IsNull(!IdVendedor) Then mProps.IdVendedor = !IdVendedor       'jeaa 25/05/2007
            If Not IsNull(!CodVendedor) Then mProps.CodVendedor = !CodVendedor  'jeaa 25/05/2007
            If Not IsNull(!NombreVendedor) Then mProps.NombreVendedor = !NombreVendedor  'jeaa 25/05/2007
            
            If Not IsNull(!IdCobrador) Then mProps.IdCobrador = !IdCobrador       'jeaa 25/05/2007
            If Not IsNull(!CodCobrador) Then mProps.CodCobrador = !CodCobrador  'jeaa 25/05/2007
            If Not IsNull(!NombreCobrador) Then mProps.NombreCobrador = !NombreCobrador  'jeaa 25/05/2007
            
            
            If Not IsNull(!Guid) Then mProps.Guid = !Guid       '*** MAKOTO 16/mar/01 Agregado
            
            If Not IsNull(!IdCuentaContable) Then mProps.auxIdCuenta = !IdCuentaContable
            If Not IsNull(!IdCuentaContable2) Then mProps.auxIdCuenta2 = !IdCuentaContable2
                        
            If Not IsNull(!IdBanco) Then mProps.IdBanco = !IdBanco   ' jeaa 20/07/2009
            If Not IsNull(!CodBanco) Then mProps.CodBanco = !CodBanco
            If Not IsNull(!IdTarjeta) Then mProps.IdTarjeta = !IdTarjeta
            If Not IsNull(!CodTarjeta) Then mProps.CodTarjeta = !CodTarjeta
            If Not IsNull(!NumCuenta) Then mProps.NumCuenta = !NumCuenta
            If Not IsNull(!numCheque) Then mProps.numCheque = !numCheque
            If Not IsNull(!TitularCta) Then mProps.TitularCta = !TitularCta
            If Not IsNull(!idElemento) Then mProps.idElemento = !idElemento
            If Not IsNull(!Codelemento) Then mProps.Codelemento = !Codelemento
            If Not IsNull(!ValorCapital_a) Then mProps.ValorCapital = !ValorCapital_a
            If Not IsNull(!ValorInteres_a) Then mProps.ValorInteres = !ValorInteres_a
            If Not IsNull(!OrdenCuota_a) Then mProps.OrdenCuota = !OrdenCuota_a
            If Not IsNull(!IdCuentaDiferida) Then mProps.auxIdCuenta3 = !IdCuentaDiferida
            If Not IsNull(!IdCuentaActivo) Then mProps.auxIdCuenta4 = !IdCuentaActivo
            If Not IsNull(!IdEquipo) Then mProps.IdEquipo = !IdEquipo
            If Not IsNull(!CodEquipo) Then mProps.CodEquipo = !CodEquipo
            If Not IsNull(!IDFormaSRI) Then mProps.IDFormaSRI = !IDFormaSRI
            If Not IsNull(!BandPosf) Then mProps.BandPosf = !BandPosf
            
        End If
    End With
End Sub

Friend Function Clone() As PCKardex
    Dim nuevo As PCKardex
    
    Set nuevo = New PCKardex
    nuevo.Props = mProps
    '*** No sirve porque cuendo clona por el cierre no funciona
    'nuevo.Guid = "" '*** borro el GUID del njuevo objeto para que no se repita nunca.  ***Oliver 10/12/2004'
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
    
    'Coge la fecha de emisión
    Me.FechaEmision = value.FechaTrans
End Property

Public Property Get GNComprobante() As GNComprobante
    Set GNComprobante = mProps.objGNComprobante
End Property

Friend Property Let BandPorCobrar(ByVal value As Boolean)
    mProps.BandPorCobrar = value
End Property

Public Property Get BandPorCobrar() As Boolean
    BandPorCobrar = mProps.BandPorCobrar
End Property


Public Property Get id() As Long
    id = mProps.id
End Property

Public Property Let idGestionc(ByVal value As Long) 'solo para cierre
    mProps.idGestionc = value
End Property

'Friend Property Let IdProvCli(ByVal value As Long)
'    mProps.IdProvCli = value
'    mProps.objGNComprobante.SetModificado
'End Property

Friend Property Get IdProvcli() As Long
    IdProvcli = mProps.IdProvcli
End Property

Public Property Let CodProvcli(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodProvcli Then
        If Len(value) > 0 Then
            'Actualiza IdProvCli también
            sql = "SELECT IdProvCli,IdCuentaContable, IdCuentaContable2 " & _
                  "FROM PCProvCli WHERE CodProvCli='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Pckardex_CodProvCli", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdProvcli = rs!IdProvcli
                mProps.auxIdCuenta = rs!IdCuentaContable
                mProps.auxIdCuenta2 = rs!IdCuentaContable2
            End If
            rs.Close
        Else
            mProps.IdProvcli = 0
            mProps.auxIdCuenta = 0
            mProps.auxIdCuenta2 = 0
        End If
    End If
    
    Set rs = Nothing
    mProps.CodProvcli = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodProvcli() As String
    CodProvcli = Trim$(mProps.CodProvcli)
End Property

Friend Property Let IDForma(ByVal value As Long)
    mProps.IDForma = value
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Get IDForma() As Long
    IDForma = mProps.IDForma
End Property

Public Property Let CodForma(ByVal value As String)
    Dim sql As String, rs As Recordset
    Dim sqlSRI As String, rsSRI As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodForma Then
        If Len(value) > 0 Then
            'Actualiza IdForma también
            sql = "SELECT IdForma,isnull(IdCuentaDiferida,0) as IdCuentaDiferida,isnull(IdCuentaActivo,0) as IdCuentaActivo FROM TSFormaCobroPago WHERE CodForma='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Pckardex_CodForma", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IDForma = rs!IDForma
                mProps.auxIdCuenta3 = rs!IdCuentaDiferida
                mProps.auxIdCuenta4 = rs!IdCuentaActivo
            End If
            rs.Close
            
            sqlSRI = "SELECT idformaSRI FROM TSFormaCobroPago WHERE CodForma='" & value & "'"
            Set rsSRI = mProps.objGNComprobante.Empresa.OpenRecordset(sqlSRI)
            If rsSRI.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Pckardex_CodFormaSRI", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                If Not IsNull(rsSRI!IDFormaSRI) Then
                    mProps.IDFormaSRI = rsSRI!IDFormaSRI
                Else
                    mProps.IDFormaSRI = 0
                End If
            End If
            rsSRI.Close
            
            
        Else
            mProps.IDForma = 0
            mProps.auxIdCuenta3 = 0
            mProps.auxIdCuenta4 = 0
        End If
        
        
    End If
    
    

    Set rs = Nothing
    mProps.CodForma = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Let CodGrupo(ByVal value As String)
    Dim sql As String, rs As Recordset

    
    'Cuando cambia el valor
    If value <> mProps.CodGrupo Then
        If Len(value) > 0 Then
            'Actualiza IdForma también
            sql = "SELECT codgrupo4 FROM pcgrupo4 WHERE Codgrupo4='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCgrupo4_codgrupo", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.CodGrupo = rs!CodGrupo
            End If
            rs.Close
        Else
            mProps.CodGrupo = ""
        End If
    End If
    
    

    Set rs = Nothing
    mProps.CodGrupo = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodForma() As String
    CodForma = Trim$(mProps.CodForma)
End Property

Public Property Get CodGrupo() As String
    CodGrupo = Trim$(mProps.CodGrupo)
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
    If Not ((CodTrans = Me.GNComprobante.CodTrans) _
            And (NumTrans = Me.GNComprobante.NumTrans)) Then
            
        'Busca CodTrans, NumTrans y Orden del doc. asignado
        sql = "SELECT pck.Id FROM Pckardex pck INNER JOIN GNComprobante gc " & _
                        "ON pck.TransID = gc.TransID " & _
              "WHERE gc.CodTrans = '" & CodTrans & "' AND gc.NumTrans = " & NumTrans & _
                     " AND pck.Orden = " & Orden
        Set rs = Me.GNComprobante.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then
            mProps.IdAsignado = rs.Fields("Id")
            mProps.objGNComprobante.SetModificado
        Else
            Err.Raise ERR_NOHAYCODIGO, "Pckardex", _
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
    sql = "SELECT Id, TransID FROM Pckardex WHERE Guid = '" & GuidAsignado & "'"
    Set rs = Me.GNComprobante.Empresa.OpenRecordset(sql)
    If Not rs.EOF Then
        'Si  el documento asignado NO está en la misma transacción
        If rs.Fields("TransID") <> Me.GNComprobante.transid Then
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
            Err.Raise ERR_NOHAYCODIGO, "Pckardex", _
                    "No se encuentra el documento asignado. " & vbCr & _
                    "(" & Me.GNComprobante.CodTrans & _
                          Me.GNComprobante.NumTrans & ") " & vbCr & GuidAsignado
        End If
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Function BuscarGuidEnTrans(ByVal GuidBuscar As String) As Boolean
    Dim i As Long, pck As PCKardex
    
    BuscarGuidEnTrans = False
    For i = 1 To Me.GNComprobante.CountPCKardex
        Set pck = Me.GNComprobante.PCKardex(i)
        If Not (pck Is Me) Then
            If pck.Guid = GuidBuscar Then
                BuscarGuidEnTrans = True
                GoTo salida
            End If
        End If
    Next i
salida:
    Set pck = Nothing
End Function
Public Property Get NumLetra() As String
'AUC quitado no funciona bien
'If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "LIRA") > 0 And (Me.gnComprobante.EsNuevo Or Me.gnComprobante.Modificado) Then
'    NumLetra = NumLetraNext'
'Else
    NumLetra = Trim$(mProps.NumLetra)
'End If
End Property
Public Property Let NumLetra(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "Pckardex_NumLetra", MSGERR_INVALIDO
    End If
    
    mProps.NumLetra = value
    mProps.objGNComprobante.SetModificado
End Property

Private Sub VerificaSaldo(ByVal value As Currency, ByVal bandDebe As Boolean)
    Dim Vdoc As Currency, Saldo As Currency, pcd As PCDocAsignado, cancelado As Currency

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
            Err.Raise ERR_INVALIDO, "Pckardex.VerificaSaldo", _
                "No se puede reducir el valor debido a que ya está cancelado la cantidad de " & _
                cancelado & " " & mProps.objGNComprobante.CodMoneda
        End If
    
    'Si quiere aumentar el valor de cobro/pago
    ElseIf (mProps.IdAsignado <> 0) And (value > Vdoc) Then
'    ElseIf (mProps.IdAsignado <> 0) Then
        'Obtiene el doc. por pagar asignado
        Set pcd = RecuperaPCDocAsignado
        If pcd.id Then      'Si es que encuentra
        
            'Convierte en moneda del doc. asignado
            value = value * Me.GNComprobante.Cotizacion("")
            value = value / mProps.objGNComprobante.Cotizacion(pcd.CodMoneda)
            If value > pcd.Valor Then
                'No puede ser mayor al valor de documento asignado
                Err.Raise ERR_INVALIDO, "Pckardex_VerificaSaldo", _
                    "No se puede asignar un valor mayor al del documento. " & vbCr & vbCr & _
                    "Valor del documento: " & vbTab & pcd.Valor & vbCr & _
                    "Valor asignado:        " & vbTab & value & _
                    " " & mProps.objGNComprobante.CodMoneda
            ElseIf value > pcd.Saldo + Vdoc Then
                'No puede ser mayor al saldo del documento asignado
'Debug.Print Me.GNComprobante.CodTrans, Me.GNComprobante.NumTrans
                Err.Raise ERR_INVALIDO, "Pckardex.VerificaSaldo", _
                    "No se puede asignar un valor mayor al saldo del documento. " & vbCr & vbCr & _
                    "Saldo:      " & vbTab & pcd.Saldo + Vdoc & vbCr & _
                    "Asignado: " & vbTab & value & vbCr & _
                    "Dif.:          " & vbTab & value - pcd.Saldo - Vdoc & _
                    " " & mProps.objGNComprobante.CodMoneda
            End If
        End If
        Set pcd = Nothing
    End If
End Sub


Public Property Let Debe(ByVal value As Currency)
    Dim vSucre As Currency
        
    'Convierte en sucres
    vSucre = value * mProps.objGNComprobante.Cotizacion("")
    
    If vSucre < 0 Then
        Err.Raise ERR_INVALIDO, "Pckardex_Debe", MSGERR_INVALIDO
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
        Err.Raise ERR_INVALIDO, "Pckardex_Haber", MSGERR_INVALIDO
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

Public Property Let FechaVenci(ByVal value As Date)
    mProps.FechaVenci = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get FechaVenci() As Date
    FechaVenci = mProps.FechaVenci
End Property

'Public Property Let CotizacionAsignado(value As Currency)
'    mProps.CotizacionAsignado = value
'    mProps.objGNComprobante.SetModificado
'End Property
'
'Public Property Get CotizacionAsignado() As Currency
'    CotizacionAsignado = mProps.CotizacionAsignado
'End Property
'



Public Property Let Observacion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Pckardex_NombreEmitante", MSGERR_INVALIDO
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
    'Si no está asignado IdProvCli no devuelve nada
    If IdProvcli = 0 Then Exit Property
    auxIdCuenta = mProps.auxIdCuenta
End Property

Friend Property Get auxIdCuenta2() As Long
    'Si no está asignado IdProvCli no devuelve nada
    If IdProvcli = 0 Then Exit Property
    auxIdCuenta2 = mProps.auxIdCuenta2
End Property

'Public Sub SetCodProvCliAux(auxIdCuenta As Long)
'    mProps.auxIdCuenta = auxIdCuenta
'End Sub

Public Function RecuperaPCDocAsignado() As PCDocAsignado
    Dim pcd As PCDocAsignado

    If mProps.IdAsignado = 0 Then Exit Function
    
    Set pcd = New PCDocAsignado
    Set pcd.PCKardex = Me
    pcd.Recuperar mProps.IdAsignado
    Set RecuperaPCDocAsignado = pcd
    Set pcd = Nothing
End Function

Public Function CalculaSaldo() As Currency
    Dim sql As String, campo As String, NumMon As Integer
    Dim rs As Recordset

    NumMon = Me.GNComprobante.Empresa.GNOpcion.IndiceMoneda(Me.GNComprobante.CodMoneda)
    campo = "Saldo" & NumMon
    
    sql = "SELECT " & campo & " FROM vwConsPCDocSaldo WHERE Id=" & mProps.id
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

    NumMon = Me.GNComprobante.Empresa.GNOpcion.IndiceMoneda(Me.GNComprobante.CodMoneda)
    campo = "VCancelado" & NumMon
    
    sql = "SELECT " & campo & " FROM vwConsPCDocSaldo WHERE Id=" & mProps.id
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
    If Not rs.EOF Then
        CalculaCancelado = rs.Fields(campo)
    End If
    rs.Close
    Set rs = Nothing
End Function

Public Function CalculaCanceladoEnTrans() As Currency
    Dim pck As PCKardex, ix As Long, c As Currency
    
    For ix = 1 To Me.GNComprobante.CountPCKardex
        Set pck = Me.GNComprobante.PCKardex(ix)
        If pck.IdAsignado = Me.id Then
            c = c + pck.Debe + pck.Haber
        End If
    Next ix
    Set pck = Nothing
    CalculaCanceladoEnTrans = c
End Function


'Devuelve True si es nuevo o no tiene cobros/pagos asignados
'         False si tiene cobros/pagos asignados
Public Function VerificaEliminacion() As Boolean
    Dim sql As String, rs As Recordset
    
    VerificaEliminacion = True

    'Si es nuevo no hay para que consultar la base
    If mProps.id = 0 Then Exit Function

    'Busca en la Pckardex registros asignados que sea de otro comprobante
    sql = "SELECT Id FROM Pckardex WHERE IdAsignado=" & mProps.id & _
                                    " AND TransID<>" & Me.GNComprobante.transid

    
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
        .FechaVenci = .FechaEmision
        .NumLetra = "1"
        
        'Activa la verificación de saldo de doc. asignado
        .NoVerificarSaldo = True                '*** MAKOTO 22/mar/01 Agregado
    End With
End Sub

Private Sub Class_Terminate()
    Set mProps.objGNComprobante = Nothing
End Sub


'*** MAKOTO 16/mar/01 Agregado
'Recibe guid y lo busca en Pckardex
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
    
    sql = "SELECT Id FROM Pckardex WHERE guid='" & mProps.Guid & "'"
    Set rs = Me.GNComprobante.Empresa.OpenRecordset(sql)
    If rs.EOF Then
'No debe generar ningún error porque cuando es transacción nueva no ha de existir
'        Err.Raise ERR_NOHAYCODIGO, "Pckardex.SetIdFromGuid", _
'            MSGERR_NOHAYCODIGO & vbCr & "Guid=" & mProps.Guid
    Else
        mProps.id = rs.Fields("Id")
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
    If mProps.id = 0 Then Exit Function


    'Busca en la Pckardex registros asignados que sea de otro comprobante
    'JEAA       modificado para que visualice cual es el comprobante
    sql = " SELECT gnc.codtrans, gnc.numtrans " & _
            " FROM Pckardex pck INNER JOIN GNComprobante gnc " & _
            " ON gnc.transid=pck.transid " & _
            " WHERE IdAsignado=" & mProps.id & _
            " AND pck.TransID<>" & Me.GNComprobante.transid
    sql = sql & " and estado <>3 "
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
    
    'Si existe, devuelve False sino True
    If Not rs.EOF Then
        Trans = rs.Fields(0) & "-" & rs.Fields(1)
        VerificaEliminacionNEW = False
    End If
    rs.Close
    Set rs = Nothing
End Function

Friend Property Get IdCobrador() As Long
    IdCobrador = mProps.IdCobrador
End Property

Public Property Let CodCobrador(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCobrador Then
        If Len(value) > 0 Then
            'Actualiza IdProvCli también
            sql = "SELECT IdVendedor, Nombre " & _
                  "FROM FCVendedor WHERE CodVendedor='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Pckardex_CodCobrador", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdCobrador = rs!IdVendedor
                mProps.NombreCobrador = rs!Nombre
            End If
            rs.Close
        Else
            mProps.IdCobrador = 0
            mProps.NombreCobrador = ""
        End If
    End If

    Set rs = Nothing
    mProps.CodCobrador = value
    mProps.objGNComprobante.SetModificado
End Property


Public Property Get CodCobrador() As String
    CodCobrador = Trim$(mProps.CodCobrador)
End Property

Friend Property Get IdVendedor() As Long
    IdVendedor = mProps.IdVendedor
End Property

Public Property Let CodVendedor(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodVendedor Then
        If Len(value) > 0 Then
            'Actualiza IdProvCli también
            sql = "SELECT IdVendedor,Nombre " & _
                  "FROM FCVendedor WHERE CodVendedor='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Pckardex_CodVendedor", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdVendedor = rs!IdVendedor
                mProps.NombreVendedor = rs!Nombre
            End If
            rs.Close
        Else
            mProps.IdVendedor = 0
            mProps.NombreVendedor = ""
        End If
    End If

    Set rs = Nothing
    mProps.CodVendedor = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodVendedor() As String
    CodVendedor = Trim$(mProps.CodVendedor)
End Property

Public Property Let NombreCobrador(ByVal value As String)
    mProps.NombreCobrador = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get NombreCobrador() As String
    NombreCobrador = Trim$(mProps.NombreCobrador)
End Property

Public Property Let NombreVendedor(ByVal value As String)
    mProps.NombreVendedor = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get NombreVendedor() As String
    NombreVendedor = Trim$(mProps.NombreVendedor)
End Property

'AUC 19/06/07
Public Property Let IdVendedor(ByVal value As Long)
    mProps.IdVendedor = value
    mProps.objGNComprobante.SetModificado
End Property
Public Property Let IdCobrador(ByVal value As Long)
    mProps.IdCobrador = value
    mProps.objGNComprobante.SetModificado
End Property

Public Function RecuperaPCDocAsignadoT() As PCDocAsignado
    Dim pcd As PCDocAsignado

    If mProps.IdAsignado = 0 Then Exit Function
    
    Set pcd = New PCDocAsignado
    Set pcd.PCKardex = Me
    pcd.Recuperar mProps.IdAsignado
    Set RecuperaPCDocAsignadoT = pcd
    Set pcd = Nothing
End Function

Public Function ObtieneSaldodePago(ByVal transid As String) As Currency
    Dim sql As String, rs As Recordset
    
    sql = " SELECT Saldo" & mProps.objGNComprobante.Cotizacion(mProps.objGNComprobante.CodMoneda) & _
            " FROM vwConsPCDocSaldo " & _
            " WHERE Id=" & transid
    
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
    
    'Si existe, devuelve False sino True
    If Not rs.EOF Then
        ObtieneSaldodePago = rs.Fields(0)
    End If
    rs.Close
    Set rs = Nothing
End Function

'AUC 08/12/06
Public Property Let BandNoAfectaPckardex(ByVal value As Boolean)
    mProps.BandNoAfectaPckardex = value
End Property

Public Property Get BandNoAfectaPckardex() As Boolean
    BandNoAfectaPckardex = mProps.BandNoAfectaPckardex
End Property

Friend Property Let IdBanco(ByVal value As Long)
    mProps.IdBanco = value
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Get IdBanco() As Long
    IdBanco = mProps.IdBanco
End Property

Public Property Get CodBanco() As String
    CodBanco = Trim$(mProps.CodBanco)
End Property

Public Property Let CodBanco(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodBanco Then
        If Len(value) > 0 Then
            'Actualiza IdBanco también
            sql = "SELECT IdBanco FROM IVBanco WHERE CodBanco='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Pckardex_CodBanco", MSGERR_NOHAYCODIGO
                Exit Property
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

Friend Property Let IdTarjeta(ByVal value As Long)
    mProps.IdTarjeta = value
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Get IdTarjeta() As Long
    IdTarjeta = mProps.IdTarjeta
End Property

Public Property Get CodTarjeta() As String
    CodTarjeta = Trim$(mProps.CodTarjeta)
End Property

Public Property Let CodTarjeta(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodTarjeta Then
        If Len(value) > 0 Then
            'Actualiza IdTarjeta también
            sql = "SELECT IdTarjeta FROM IVTarjeta WHERE CodTarjeta='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Pckardex_CodTarjeta", MSGERR_NOHAYCODIGO
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
    mProps.CodTarjeta = value
    mProps.objGNComprobante.SetModificado
End Property


Public Property Let NumCuenta(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "Pckardex_NumCuenta", MSGERR_INVALIDO
    End If
    
    mProps.NumCuenta = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get NumCuenta() As String
    NumCuenta = Trim$(mProps.NumCuenta)
End Property

Public Property Let numCheque(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "Pckardex_NumCheque", MSGERR_INVALIDO
    End If
    
    mProps.numCheque = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get numCheque() As String
    numCheque = Trim$(mProps.numCheque)
End Property

Public Property Let TitularCta(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "Pckardex_TitularCta", MSGERR_INVALIDO
    End If
    
    mProps.TitularCta = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get TitularCta() As String
    TitularCta = Trim$(mProps.TitularCta)
End Property


Private Function BuscaNumeroPrecio(fila As Long) As Integer
    Dim iv As IVInventario, i As Integer, j As Long, ivpadre As IVInventario
    Set iv = mProps.objGNComprobante.Empresa.RecuperaIVInventario(mProps.objGNComprobante.IVKardex(fila).CodInventario)
    If mProps.objGNComprobante.IVKardex(fila).idpadre = 0 Then
        If Not (iv Is Nothing) Then
            If mProps.objGNComprobante.EsNuevo Then
               For i = 1 To 5
                    If mProps.objGNComprobante.IVKardex(fila).NumeroPrecio <> 0 Then
                        If mProps.objGNComprobante.IVKardex(fila).Precio = iv.Precio(i) Then
                            BuscaNumeroPrecio = i
                            Exit For
                        End If
                    Else
                        BuscaNumeroPrecio = mProps.objGNComprobante.IVKardex(fila).NumeroPrecio
                    End If
                Next i
            Else
                    If mProps.objGNComprobante.IVKardex(fila).NumeroPrecio <> 0 Then
'''                         jeaa  anulado 03/12/2010
'''                        BuscaNumeroPrecio = mProps.objGNComprobante.IVKardex(fila).NumeroPrecio
'''                    Else
                        For i = 1 To 5
                            If mProps.objGNComprobante.IVKardex(fila).Precio = iv.Precio(i) Then
                                BuscaNumeroPrecio = i
                                Exit For
                            End If
                        Next i
'                        BuscaNumeroPrecio = 0
                    End If
            End If
       End If
    Else
        If Not (iv Is Nothing) Then
            For j = 1 To mProps.objGNComprobante.CountIVKardex
                If mProps.objGNComprobante.IVKardex(j).idInventario = mProps.objGNComprobante.IVKardex(fila).idpadre Then
                    Set ivpadre = mProps.objGNComprobante.Empresa.RecuperaIVInventario(mProps.objGNComprobante.IVKardex(j).CodInventario)
                    If Not (iv Is Nothing) Then
                        For i = 1 To 5
                            If mProps.objGNComprobante.IVKardex(j).Precio = ivpadre.Precio(i) Then
                                BuscaNumeroPrecio = i
                                Exit For
                            End If
                        Next i
                    End If
                    'BuscaNumeroPrecio = mProps.objGNComprobante.IVKardex(j).NumeroPrecio
                    Exit For
                End If
            Next j
        End If
    End If
    Set iv = Nothing
End Function

Public Function ObtieneIdCliente(ByVal IdProvcli As String) As Long
    Dim sql As String, rs As Recordset
    
    sql = " SELECT Saldo" & mProps.objGNComprobante.Cotizacion(mProps.objGNComprobante.CodMoneda) & _
            " FROM vwConsPCDocSaldo " & _
            " WHERE Id=" & IdProvcli
    
   sql = " select IDProvcli from gncomprobante g"
   sql = sql & " inner join Pckardex P on g.transid=P.transid"
   sql = sql & " Where id = " & IdProvcli
    
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
    
    'Si existe, devuelve False sino True
    If Not rs.EOF Then
        ObtieneIdCliente = rs.Fields("idProvCli")
    Else
        ObtieneIdCliente = 0
    End If
    rs.Close
    Set rs = Nothing
End Function

Public Property Let idElemento(ByVal value As Long)
    mProps.idElemento = value
    mProps.objGNComprobante.SetModificado
End Property
Public Property Get idElemento() As Long
    idElemento = mProps.idElemento
End Property

Private Function NumLetraNext() As Integer
    Dim sql As String
    Dim rs As Recordset
    On Error GoTo CapturaError
    sql = "Select numletra from tsformacobropago where codforma = '" & mProps.CodForma & "'"
    Set rs = Me.GNComprobante.Empresa.OpenRecordset(sql)
    If Not rs Is Nothing Then
        NumLetraNext = rs!NumLetra
    End If
    Set rs = Nothing
    Exit Function
CapturaError:
    Set rs = Nothing
    MsgBox Err.Description
    Exit Function
End Function

Friend Property Get auxIdCuentaXIII() As Long
    'Si no está asignado IdProvCli no devuelve nada
    If IdProvcli = 0 Then Exit Property
    auxIdCuentaXIII = mProps.auxIdCuentaXIII
End Property
Friend Property Get auxIdCuentaXIV() As Long
    'Si no está asignado IdProvCli no devuelve nada
    If IdProvcli = 0 Then Exit Property
    auxIdCuentaXIV = mProps.auxIdCuentaXIV
End Property
Friend Property Get auxIdCuentaVac() As Long
    'Si no está asignado IdProvCli no devuelve nada
    If IdProvcli = 0 Then Exit Property
    auxIdCuentaVac = mProps.auxIdCuentaVac
End Property
Friend Property Get auxIdCuentaNetoRol() As Long
    'Si no está asignado IdProvCli no devuelve nada
    If IdProvcli = 0 Then Exit Property
    auxIdCuentaNetoRol = mProps.auxIdCuentaNetoRol
End Property

Public Property Let RubroRol(ByVal value As String)
    Dim sql As String, rs As Recordset, rsR As Recordset
    Dim msg As String
    Dim pcg As PcGrupo
    Dim pc As PCProvCli
    Dim idG As Byte
    'Cuando cambia el valor
    If value <> mProps.RubroRol Then
        If Len(value) > 0 Then
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ConAsiento")) > 0 Then
            idG = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ConAsiento") + 1
                If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SeContabilizaPor")) > 0 Then
                    If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SeContabilizaPor") = 0 Then
                        sql = "Select cp.*,e.tipo from cuentapersonal cp "
                        sql = sql & "Inner join elemento e on e.idelemento= cp.idelemento "
                        sql = sql & "Where codelemento = '" & value & "' AND cp.idempleado = " & mProps.IdProvcli
                    ElseIf gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SeContabilizaPor") = 1 Then
                        
                        
                        sql = "Select top 1 idgrupo1,idgrupo2,idgrupo3,idgrupo4 from roldetalle where transid = " & mProps.objGNComprobante.transid
                        sql = sql & " And IdEmpleado =" & mProps.IdProvcli
                        Set rsR = gobjMain.EmpresaActual.OpenRecordset(sql)
                        If Not rsR.EOF Then
                            If mProps.objGNComprobante.GNTrans.CodPantalla = "GENROL" Then
                                Select Case idG
                                    Case 1: Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, rsR!IdGrupo1, 4)
                                    Case 2: Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, rsR!IdGrupo2, 4)
                                    Case 3: Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, rsR!IdGrupo3, 4)
                                    Case 4: Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, rsR!IdGrupo4, 4)
                                End Select
                            Else
                                Set pc = gobjMain.EmpresaActual.RecuperaEmpleado(mProps.IdProvcli)
                                Select Case idG
                                    Case 1: Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, pc.CodGrupo1, 4)
                                    Case 2: Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, pc.CodGrupo2, 4)
                                    Case 3: Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, pc.CodGrupo3, 4)
                                    Case 4: Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, pc.CodGrupo4, 4)
                                End Select
                            
                            End If
                        Else
                                Set pc = gobjMain.EmpresaActual.RecuperaEmpleado(mProps.IdProvcli)
                                Select Case idG
                                    Case 1: Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, pc.CodGrupo1, 4)
                                    Case 2: Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, pc.CodGrupo2, 4)
                                    Case 3: Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, pc.CodGrupo3, 4)
                                    Case 4: Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, pc.CodGrupo4, 4)
                                End Select

                        End If
                            sql = "Select cp.*,e.tipo from cuentapcgrupo cp "
                            sql = sql & "Inner join elemento e on e.idelemento= cp.idelemento "
                            sql = sql & " Inner  join pcgrupo" & idG & " pcg on pcg.idgrupo" & idG & " = cp.idpcgrupo "
                            sql = sql & "Where codelemento = '" & value & "' AND cp.idpcgrupo = " & pcg.IdGrupo
                            Set rsR = Nothing
                            Set pc = Nothing
'                            Set pcg = Nothing
                    End If
                End If
        End If
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                msg = "El siguiente Elemento " & value & " no tiene asignado una cuenta contable, " & vbCr & _
                "por lo que no se puede generar ésta transacción. " & vbCr & vbCr & _
                "Para poder Continuar, por favor habra datos del dicho Elemento" & vbCr & _
                "y asigne una cuenta contable correspondiente."
                Err.Raise ERR_NOHAYCODIGO, "Pckardex_RubroRol", msg
                Exit Property
            Else
                Select Case rs!Tipo
                    Case 7: mProps.auxIdCuentaXIII = rs!IdCuenta  '-> decimo tercero
                    Case 8: mProps.auxIdCuentaXIV = rs!IdCuenta ' -> decimo cuarto
                    Case 9: mProps.auxIdCuentaVac = rs!IdCuenta ' -> vac
                    Case 5, 3: mProps.auxIdCuentaNetoRol = rs!IdCuenta ' -> netorol
                    '3 Para poder cruzar anticipos no se si funcionara hay que analizar
                End Select
                mProps.EsRubroRol = True
            End If
        End If
            rs.Close
    Else
        mProps.auxIdCuentaXIII = 0
        mProps.auxIdCuentaXIV = 0
        mProps.auxIdCuentaVac = 0
        mProps.auxIdCuentaNetoRol = 0
        mProps.EsRubroRol = False
    End If
    Set rs = Nothing
'    mProps.objGNComprobante.SetModificado
End Property
Public Property Get EsRubroRol() As Boolean
    EsRubroRol = Trim$(mProps.EsRubroRol)
End Property

Private Sub GrabarPCFinanciamiento()
    Dim sql As String, rs As Recordset
    Dim AuxDesc As Integer, SecuencialItem As Long
    Dim rsA As Recordset
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM PCFinanciamiento WHERE Transid=" & mProps.id
        Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
        If Not rs.EOF Then
            #If DAOLIB Then
                rs.Edit
            #End If
        Else
'            If mProps.Cantidad > 0 Then Exit Sub
            sql = "SELECT * FROM PCFinanciamiento WHERE 1=0"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
            rs.AddNew
        End If
    With rs
        !IdPCKardex = mProps.id                       'Es el mismo asignado a GnComprobante
        !transid = mProps.objGNComprobante.transid
        !ValorCapital1 = mProps.ValorCapital1
        !ValorInteres1 = mProps.ValorInteres1
        !TasaInteres1 = mProps.TasaInteres1
        !ValorCapital2 = mProps.ValorCapital2
        !ValorInteres2 = mProps.ValorInteres2
        !TasaInteres2 = mProps.TasaInteres2
        !ValorCapital3 = mProps.ValorCapital3
        !ValorInteres3 = mProps.ValorInteres3
        !TasaInteres3 = mProps.TasaInteres3
        .Update
        #If DAOLIB Then
                rs.Bookmark = rs.LastModified
        #End If
        .Move 0             'Para actualizar
        .Close
    End With
End Sub
Public Sub RecuperarPCFinanciamiento(ByVal transid As Long, ByVal IdIvkardex As Long)
    Dim sql As String, rs As Recordset
    If transid = 0 Then transid = mProps.id
    sql = " SELECT pcFinan.* "
    sql = sql & " FROM PCFinanciamiento ivFinan"
    sql = sql & " Where TransID = " & transid
    sql = sql & " and IdIVKardex= " & IdIvkardex
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
    If rs.EOF Then Exit Sub
    With rs
        If Not IsNull(!ValorCapital1) Then mProps.ValorCapital1 = !ValorCapital1
        If Not IsNull(!ValorInteres1) Then mProps.ValorInteres1 = !ValorInteres1
        If Not IsNull(!TasaInteres1) Then mProps.TasaInteres1 = !ValorTasaInteres1
        If Not IsNull(!ValorCapital2) Then mProps.ValorCapital2 = !ValorCapital2
        If Not IsNull(!ValorInteres2) Then mProps.ValorInteres2 = !ValorInteres2
        If Not IsNull(!TasaInteres2) Then mProps.TasaInteres2 = !ValorTasaInteres2
        If Not IsNull(!ValorCapital3) Then mProps.ValorCapital3 = !ValorCapital3
        If Not IsNull(!ValorInteres3) Then mProps.ValorInteres3 = !ValorInteres3
        If Not IsNull(!TasaInteres3) Then mProps.TasaInteres3 = !ValorTasaInteres3
        .Close
    End With
    Set rs = Nothing
End Sub
Public Property Let ValorCapital1(ByVal value As Currency)
    mProps.ValorCapital1 = value
End Property
Public Property Get ValorCapital1() As Currency
    ValorCapital1 = mProps.ValorCapital1
End Property
Public Property Let ValorCapital2(ByVal value As Currency)
    mProps.ValorCapital2 = value
End Property
Public Property Get ValorCapital2() As Currency
    ValorCapital2 = mProps.ValorCapital2
End Property
Public Property Let ValorCapital3(ByVal value As Currency)
    mProps.ValorCapital3 = value
End Property
Public Property Get ValorCapital3() As Currency
    ValorCapital3 = mProps.ValorCapital3
End Property
Public Property Let ValorInteres1(ByVal value As Currency)
    mProps.ValorInteres1 = value
End Property
Public Property Get ValorInteres1() As Currency
    ValorInteres1 = mProps.ValorInteres1
End Property
Public Property Let ValorInteres2(ByVal value As Currency)
    mProps.ValorInteres2 = value
End Property
Public Property Get ValorInteres2() As Currency
    ValorInteres2 = mProps.ValorInteres2
End Property
Public Property Let ValorInteres3(ByVal value As Currency)
    mProps.ValorInteres3 = value
End Property
Public Property Get ValorInteres3() As Currency
    ValorInteres3 = mProps.ValorInteres3
End Property
Public Property Let OrdenCuota(ByVal value As Integer)
    mProps.OrdenCuota = value
    mProps.objGNComprobante.SetModificado
End Property
Public Property Get OrdenCuota() As Integer
    OrdenCuota = mProps.OrdenCuota
End Property
Public Property Let ValorCapital(ByVal value As Currency)
    mProps.ValorCapital = value
End Property
Public Property Get ValorCapital() As Currency
    ValorCapital = mProps.ValorCapital
End Property
Public Property Let ValorInteres(ByVal value As Currency)
    mProps.ValorInteres = value
End Property
Public Property Get ValorInteres() As Currency
    ValorInteres = mProps.ValorInteres
End Property
Public Property Let TasaInteres1(ByVal value As Currency)
    mProps.TasaInteres1 = value
End Property
Public Property Get TasaInteres1() As Currency
    TasaInteres1 = mProps.TasaInteres1
End Property
Public Property Let TasaInteres2(ByVal value As Currency)
    mProps.TasaInteres2 = value
End Property
Public Property Get TasaInteres2() As Currency
    TasaInteres2 = mProps.TasaInteres2
End Property
Public Property Let TasaInteres3(ByVal value As Currency)
    mProps.TasaInteres3 = value
End Property
Public Property Get TasaInteres3() As Currency
    TasaInteres3 = mProps.TasaInteres3
End Property
Public Property Let CodEmpleado(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
        If value <> mProps.CodEmpleado Then
        If Len(value) > 0 Then
            'Actualiza IdProvCli también
            sql = "SELECT IdProvCli,IdCuentaContable, IdCuentaContable2 " & _
                  "FROM Empleado WHERE CodProvCli='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Pckardex_CodProvCli", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdProvcli = rs!IdProvcli
                mProps.auxIdCuenta = rs!IdCuentaContable
                mProps.auxIdCuenta2 = rs!IdCuentaContable2
            End If
            rs.Close
        Else
            mProps.IdProvcli = 0
            mProps.auxIdCuenta = 0
            mProps.auxIdCuenta2 = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodEmpleado = value
    mProps.objGNComprobante.SetModificado
End Property


Public Property Get CodEmpleado() As String
    CodEmpleado = Trim$(mProps.CodEmpleado)
End Property

Private Sub GrabarPCKardexAmort()
    Dim sql As String, rs As Recordset
    Dim AuxDesc As Integer, SecuencialItem As Long
    Dim rsA As Recordset
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM PCKardexAmort WHERE idPcKardex=" & mProps.id
        Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
        If Not rs.EOF Then
            #If DAOLIB Then
                rs.Edit
            #End If
        Else
'            If mProps.Cantidad > 0 Then Exit Sub
            sql = "SELECT * FROM PCKardexAmort WHERE 1=0"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
            rs.AddNew
        End If
        
    With rs
        !IdPCKardex = mProps.id
        !ValorCapital = mProps.ValorCapital
        !ValorInteres = mProps.ValorInteres
        !OrdenCuota = mProps.OrdenCuota
        .Update
        #If DAOLIB Then
                rs.Bookmark = rs.LastModifiedasfasdf
        #End If
        .Move 0             'Para actualizar
        .Close
    End With
End Sub

Public Property Let Codelemento(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.Codelemento Then
        If Len(value) > 0 Then
            'Actualiza IdProvCli también
            sql = "SELECT IdElemento " & _
                  "FROM Elemento WHERE Codelemento='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Pckardex_CodElemento", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.idElemento = rs!idElemento
            End If
            rs.Close
        Else
            mProps.idElemento = 0
        End If
    End If
    Set rs = Nothing
    mProps.Codelemento = value
    mProps.objGNComprobante.SetModificado
End Property
Public Property Get Codelemento() As String
    Codelemento = Trim$(mProps.Codelemento)
End Property


Friend Property Get auxIdCuenta3() As Long
    If IDForma = 0 Then Exit Property
    auxIdCuenta3 = mProps.auxIdCuenta3
End Property

Friend Property Get auxIdCuenta4() As Long
    If IDForma = 0 Then Exit Property
    auxIdCuenta4 = mProps.auxIdCuenta4
End Property

Public Property Let Porc(ByVal value As Currency)
    mProps.Porc = value
End Property
Public Property Get Porc() As Currency
    Porc = mProps.Porc
End Property

Private Sub UpdateGnCall(ByVal id_antes As Long)
Dim sql As String
Dim rs As Recordset
'verifica que exista




If Len(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("FornmaCobroOtrasCuotas")) > 0 Then
    If mProps.CodForma = mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("FornmaCobroOtrasCuotas") Then
        
        sql = " Select * from gncall Where IdpcKardex1= " & id_antes
        Set rs = Me.GNComprobante.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then
            sql = "UPDATE GnCall SET IdpcKardex1=" & mProps.id & _
                  " WHERE IdpcKardex1=" & id_antes
        #If DAOLIB Then
                mProps.objGNComprobante.Empresa.Database.Execute sql, dbFailOnError
        #Else
                mProps.objGNComprobante.Empresa.Coneccion.Execute sql
        #End If
      End If
    End If
End If
If Len(mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("FormaCobroCuotaInt")) > 0 Then
    If mProps.CodForma = mProps.objGNComprobante.Empresa.GNOpcion.ObtenerValor("FormaCobroCuotaInt") Then
        sql = " Select * from gncall Where IdpcKardex1= " & id_antes
        Set rs = Me.GNComprobante.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then
        sql = "UPDATE GnCall SET IdpcKardex2=" & mProps.id & _
              " WHERE IdpcKardex2=" & id_antes
        #If DAOLIB Then
                mProps.objGNComprobante.Empresa.Database.Execute sql, dbFailOnError
        #Else
                mProps.objGNComprobante.Empresa.Coneccion.Execute sql
        #End If
        End If
    End If
End If
Set rs = Nothing
End Sub

Public Function VerificaEliminacionGestion(ByRef Trans As String) As Boolean
    Dim sql As String, rs As Recordset
    
    VerificaEliminacionGestion = True

    'Si es nuevo no hay para que consultar la base
    If mProps.id = 0 Then Exit Function

    
    sql = "SELECT pcg.codgestion, gnc.observacion " & _
            " FROM gncall gnc Inner join pcgestion pcg " & _
            " On pcg.id= gnc.idgestion " & _
            " WHERE Idpckardex1=" & mProps.id & _
            " OR idpckardex2=" & mProps.id
            
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
    
    'Si existe, devuelve False sino True
    If Not rs.EOF Then
        Trans = rs.Fields(0) & " - " & rs.Fields(1)
        VerificaEliminacionGestion = False
    End If
    rs.Close
    Set rs = Nothing
End Function


Friend Sub Recuperar2Col(rs As Recordset)
    With rs
        If Not .EOF Then
            'mProps.Id = !Id
            If Not IsNull(!IdProvClipar) Then mProps.IdProvcli = !IdProvClipar
'            If !idElemento = 0 Then
'                If Not IsNull(!CodProvCli) Then mProps.CodProvCli = !CodProvCli
'            Else
'                If Not IsNull(!CodProvCli) Then mProps.CodEmpleado = !CodProvCli
'            End If
            If Not IsNull(!IDFormapar) Then mProps.IDForma = !IDFormapar
            If Not IsNull(!CodFormaPar) Then mProps.CodFormaPar = !CodFormaPar
            If Not IsNull(!CodFormaImPar) Then mProps.CodFormaImPar = !CodFormaImPar
            If Not IsNull(!IdAsignadopar) Then mProps.IdAsignado = !IdAsignadopar
            If Not IsNull(!NumLetrapar) Then mProps.NumLetra = !NumLetrapar
            If Not IsNull(!DebePar) Then mProps.DebePar = !DebePar
            If Not IsNull(!DebeImPar) Then mProps.DebeImPar = !DebeImPar
            'If Not IsNull(!Haberpar) Then mProps.Haberpar = !Haberpar
            If Not IsNull(!FechaEmisionPar) Then mProps.FechaEmisionPar = !FechaEmisionPar
            If Not IsNull(!FechaVenciPar) Then mProps.FechaVenciPar = !FechaVenciPar
            If Not IsNull(!FechaEmisionImPar) Then mProps.FechaEmisionImPar = !FechaEmisionImPar
            If Not IsNull(!FechaVenciImPar) Then mProps.FechaVenciImPar = !FechaVenciImPar
            If Not IsNull(!Obspar) Then mProps.Observacion = !Obspar
            If Not IsNull(!OrdenPar) Then mProps.OrdenPar = !OrdenPar
            If Not IsNull(!OrdenImPar) Then mProps.OrdenImPar = !OrdenImPar
            If Not IsNull(!Cuota) Then mProps.Cuota = !Cuota
        End If
    End With
End Sub
Public Property Let DebePar(ByVal value As Currency)
    Dim vSucre As Currency
    'Convierte en sucres
    vSucre = value * mProps.objGNComprobante.Cotizacion("")
    If vSucre < 0 Then
        Err.Raise ERR_INVALIDO, "Pckardex_DebePar", MSGERR_INVALIDO
    Else
        VerificaSaldo value, True
        If vSucre > 0 Then
            mProps.Haber = 0
            mProps.DebePar = vSucre
        Else
            mProps.DebePar = 0
        End If
    End If
    mProps.objGNComprobante.SetModificado
End Property
Public Property Get DebePar() As Currency
    DebePar = mProps.DebePar / mProps.objGNComprobante.Cotizacion("")
End Property
Public Property Let DebeImPar(ByVal value As Currency)
    Dim vSucre As Currency
    'Convierte en sucres
    vSucre = value * mProps.objGNComprobante.Cotizacion("")
    If vSucre < 0 Then
        Err.Raise ERR_INVALIDO, "Pckardex_DebePar", MSGERR_INVALIDO
    Else
        VerificaSaldo value, True
        If vSucre > 0 Then
            mProps.Haber = 0
            mProps.DebeImPar = vSucre
        Else
            mProps.DebeImPar = 0
        End If
    End If
    mProps.objGNComprobante.SetModificado
End Property
Public Property Get DebeImPar() As Currency
    DebeImPar = mProps.DebeImPar / mProps.objGNComprobante.Cotizacion("")
End Property
Public Property Let CodFormaPar(ByVal value As String)
    mProps.CodForma = value
    mProps.objGNComprobante.SetModificado
End Property
Public Property Get CodFormaPar() As String
    CodFormaPar = Trim$(mProps.CodFormaPar)
End Property
Public Property Let CodFormaImPar(ByVal value As String)
    mProps.CodFormaImPar = value
    mProps.objGNComprobante.SetModificado
End Property
Public Property Get CodFormaImPar() As String
    CodFormaImPar = Trim$(mProps.CodFormaImPar)
End Property
Public Property Let FechaEmisionPar(ByVal value As Date)
    mProps.FechaEmisionPar = value
    mProps.objGNComprobante.SetModificado
End Property
Public Property Get FechaEmisionPar() As Date
    FechaEmisionPar = mProps.FechaEmisionPar
End Property
Public Property Let FechaVenciPar(ByVal value As Date)
    mProps.FechaVenciPar = value
    mProps.objGNComprobante.SetModificado
End Property
Public Property Get FechaVenciPar() As Date
    FechaVenciPar = mProps.FechaVenciPar
End Property
Public Property Let FechaEmisionImPar(ByVal value As Date)
    mProps.FechaEmisionImPar = value
    mProps.objGNComprobante.SetModificado
End Property
Public Property Get FechaEmisionImPar() As Date
    FechaEmisionImPar = mProps.FechaEmisionImPar
End Property
Public Property Let FechaVenciImPar(ByVal value As Date)
    mProps.FechaVenciImPar = value
    mProps.objGNComprobante.SetModificado
End Property
Public Property Get FechaVenciImPar() As Date
    FechaVenciImPar = mProps.FechaVenciImPar
End Property

Public Property Let OrdenPar(ByVal value As Integer)
    mProps.OrdenPar = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get OrdenPar() As Integer
    OrdenPar = mProps.OrdenPar
End Property


Public Property Let OrdenImPar(ByVal value As Integer)
    mProps.OrdenImPar = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get OrdenImPar() As Integer
    OrdenImPar = mProps.OrdenImPar
End Property
Public Property Let Cuota(ByVal value As Currency)
    mProps.Cuota = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Cuota() As Currency
    Cuota = mProps.Cuota
End Property

Friend Property Let IdEquipo(ByVal value As Long)
    mProps.IdEquipo = value
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Get IdEquipo() As Long
    IdEquipo = mProps.IdEquipo
End Property

Public Property Get CodEquipo() As String
    CodEquipo = Trim$(mProps.CodEquipo)
End Property

Public Property Let CodEquipo(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodEquipo Then
        If Len(value) > 0 Then
            'Actualiza IdEquipo también
            sql = "SELECT IdEquipo FROM IVEquipo WHERE CodEquipo='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Pckardex_CodEquipo", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdEquipo = rs!IdEquipo
            End If
            rs.Close
        Else
            mProps.IdEquipo = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodEquipo = value
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Let IDFormaSRI(ByVal value As Long)
    mProps.IDFormaSRI = value
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Get IDFormaSRI() As Long
    IDFormaSRI = mProps.IDFormaSRI
End Property

Public Property Let BandPosf(ByVal value As Boolean) 'Funciona con la misma bandera del gncomprobante para los chposf para diferenciar el ef de chp
    mProps.BandPosf = value
End Property

Public Property Get BandPosf() As Boolean
    BandPosf = mProps.BandPosf
    mProps.objGNComprobante.SetModificado
End Property

Public Function VerificaEliminacionCHPNEW(ByRef Trans As String) As Boolean
    Dim sql As String, rs As Recordset
    
    VerificaEliminacionCHPNEW = True

    'Si es nuevo no hay para que consultar la base
    If mProps.id = 0 Then Exit Function


    'Busca en la Pckardex registros asignados que sea de otro comprobante
    'JEAA       modificado para que visualice cual es el comprobante
    sql = " SELECT gnc.codtrans, gnc.numtrans " & _
            " FROM Pckardexchp pck INNER JOIN GNComprobante gnc " & _
            " ON gnc.transid=pck.transid " & _
            " WHERE IdAsignadopck=" & mProps.id & _
            " AND pck.TransID<>" & Me.GNComprobante.transid
    sql = sql & " and estado <>3 "
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
    
    'Si existe, devuelve False sino True
    If Not rs.EOF Then
        Trans = rs.Fields(0) & "-" & rs.Fields(1)
        VerificaEliminacionCHPNEW = False
    End If
    rs.Close
    Set rs = Nothing
End Function


