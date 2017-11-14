VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCKardexCHP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
    BandNoAfectaPCKardexCHP As Boolean 'AUC para saber si grabo on o en PCKardexCHP temporal no se guarda en la base
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
    IdAsignadoPCK As Long
End Type

Private mProps As T_PROP
Private mcolTransIDsPorPagar As Collection 'AUC 06/12/06


Public Sub VerificaDatos()
    Dim pck As PCKardexCHP, ix As Long, pcd As PCDocAsignadoCHP, v As Currency
    Dim tsf As TSFormaCobroPago, IdProvcli As Long, pcdo As PCDocAsignado
    Dim msg As String, nump As Integer, IdBanco As Long, NumCta As String, NumCHE As String, Titu As String

    'Verifica datos
    If mProps.Debe = 0 And mProps.Haber = 0 Then
            Err.Raise ERR_INVALIDO, "PCKardexCHP.VerificaDatos", _
            "Se intentó grabar sin valor de Debe ni de Haber " & _
                "en detalle de Documentos de Proveedor/Cliente."
    End If
    If (mProps.Debe <> 0) And (mProps.Haber <> 0) Then
        Err.Raise ERR_INVALIDO, "PCKardexCHP.VerificaDatos", _
            "Se intentó grabar con Debe y Haber al mismo tiempo " & _
                "en detalle de Documentos de de Proveedor/Cliente."
    End If
    If mProps.objGNComprobante.GNTrans.CodPantalla <> "TSIE2PCK" Then
        If mProps.IdAsignado = 0 And mProps.IDForma = 0 Then
            Err.Raise ERR_INVALIDO, "PCKardexCHP_VerificaDatos", _
                "Debe seleccionar código de Forma de documento " & _
                    "en detalle de Documentos de Proveedor/Cliente."
        End If
    End If
    
    If mProps.objGNComprobante.GNTrans.CodPantalla <> "IVBQD2PCK" And mProps.objGNComprobante.GNTrans.CodPantalla <> "TSIE2PCK" Then
        If mProps.IdAsignadoPCK = 0 And mProps.IDForma = 0 Then
            Err.Raise ERR_INVALIDO, "PCKardexCHP_VerificaDatos", _
                "Debe seleccionar código de Forma de documento " & _
                    "en detalle de Documentos de Proveedor/Cliente."
        End If
    End If
    
    If mProps.IdProvcli = 0 Then
        Err.Raise ERR_INVALIDO, "PCKardexCHP.VerificaDatos", _
            "Debe seleccionar un cliente o un proveedor " & _
                "en detalle de Documentos de Proveedor/Cliente."
    End If
    If Len(mProps.NumLetra) = 0 Then
        Err.Raise ERR_INVALIDO, "PCKardexCHP.VerificaDatos", _
            "Debe ingresar un número de documento " & _
                "en detalle de Documentos de Proveedor/Cliente."
    End If
    If mProps.FechaEmision > mProps.FechaVenci Then
        Err.Raise ERR_INVALIDO, "PCKardexCHP.VerificaDatos", _
            "Fecha de vencimiento no puede ser menor a la fecha de emisión " & _
                "en detalle de Documentos de Proveedor/Cliente."
    End If
    
    If mProps.objGNComprobante.GNTrans.CodPantalla <> "TSCC" And _
         mProps.objGNComprobante.GNTrans.CodPantalla <> "TSCCA" And _
         mProps.objGNComprobante.GNTrans.CodPantalla <> "TSICHP" And _
         mProps.objGNComprobante.GNTrans.CodPantalla <> "TSIEE" Then
        'Verifica si la transaccion tiene diferente cliente en gncomprobant y PCKardexCHP
        'And (mProps.objGNComprobante.GNTrans.CodPantalla <> "TSIEE")
        If Not (mProps.objGNComprobante.GNTrans.IVProvCliPorFila) And Not (mProps.objGNComprobante.GNTrans.TSProvCliPorFila) Then
            If mProps.objGNComprobante.GNTrans.ClienteVisible And Not mProps.objGNComprobante.GNTrans.ProveedorVisible Then
                If mProps.objGNComprobante.IdClienteRef <> 0 Then
                    mProps.IdProvcli = mProps.objGNComprobante.IdClienteRef
                End If
             ElseIf mProps.objGNComprobante.GNTrans.ProveedorVisible And Not mProps.objGNComprobante.GNTrans.ClienteVisible Then
                If mProps.objGNComprobante.IdProveedorRef <> 0 Then
                    mProps.IdProvcli = mProps.objGNComprobante.IdProveedorRef
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
    
    For ix = 1 To mProps.objGNComprobante.CountPCKardexCHP
        If mProps.objGNComprobante.PCKardexCHP(ix).IdBanco > 0 Then
            IdBanco = mProps.objGNComprobante.PCKardexCHP(ix).IdBanco
            NumCta = mProps.objGNComprobante.PCKardexCHP(ix).NumCuenta
            NumCHE = mProps.objGNComprobante.PCKardexCHP(ix).numCheque
            Titu = mProps.objGNComprobante.PCKardexCHP(ix).TitularCta
            Exit For
        End If
    Next ix
    
    If IdBanco <> 0 Then
        For ix = 1 To mProps.objGNComprobante.CountPCKardexCHP
            If mProps.objGNComprobante.PCKardexCHP(ix).IdBanco = 0 Then
                mProps.objGNComprobante.PCKardexCHP(ix).IdBanco = IdBanco
                mProps.objGNComprobante.PCKardexCHP(ix).NumCuenta = NumCta
                mProps.objGNComprobante.PCKardexCHP(ix).numCheque = NumCHE
                mProps.objGNComprobante.PCKardexCHP(ix).TitularCta = Titu
            End If
        Next ix
    End If
    
    ' si es modificacion y la transaccion es varios cliente
    If Not mProps.objGNComprobante.EsNuevo Then
        If (mProps.objGNComprobante.GNTrans.IVProvCliPorFila) Or (mProps.objGNComprobante.GNTrans.TSProvCliPorFila) Then
            
'            For ix = 1 To mProps.objGNComprobante.CountPCKardexCHP
                'IdProvCli = ObtieneIdCliente(mProps.objGNComprobante.PCKardexCHP(ix).IdAsignado)
                IdProvcli = ObtieneIdCliente(Me.IdAsignado)
                If IdProvcli <> 0 Then
                    If mProps.objGNComprobante.IdClienteRef = 0 And mProps.objGNComprobante.IdProveedorRef = 0 Then
                        'If mProps.objGNComprobante.PCKardexCHP(ix).IdProvCli <> IdProvCli Then
                        If Me.IdProvcli <> IdProvcli Then
                            mProps.IdProvcli = IdProvcli
                        End If
                    End If
                End If
                
                
 '           Next ix
            
            
            
        End If
    End If
    
    
    
    'jeaa 10/05/2005
    If mProps.objGNComprobante.GNTrans.VerificaFechas Then
        If mProps.objGNComprobante.FechaTrans <> mProps.FechaEmision Then
            Err.Raise ERR_INVALIDO, "PCKardexCHP.VerificaDatos", _
                "Fecha de emision del comprobante no puede ser diferente a la fecha" & Chr(13) & "de emisión en el detalle de Documentos de Proveedor/Cliente."
        End If
    End If
    
    'Verifica si repite el mismo doc. asignado en el comprobante
    If mProps.IdAsignado And mProps.objGNComprobante.EsNuevo Then
        For ix = 1 To mProps.objGNComprobante.CountPCKardexCHP
            Set pck = mProps.objGNComprobante.PCKardexCHP(ix)
            If (pck.IdAsignado = mProps.IdAsignado) And (mProps.IdAsignado <> -1) And (Not (pck Is Me)) Then
                msg = "Existen documentos repetidos para el mismo comprobante." & vbCr & vbCr
                Set pcd = RecuperaPCDocAsignado
                If pcd.id Then msg = msg & pcd.Trans & " " & pcd.Doc & ": "
                msg = msg & Format(Me.Debe + Me.Haber, Me.GNComprobante.FormatoMoneda) & " " & _
                            Me.GNComprobante.CodMoneda & " y " & _
                            Format(pck.Debe + pck.Haber, pck.GNComprobante.FormatoMoneda) & _
                            " " & pck.GNComprobante.CodMoneda
                Err.Raise ERR_INVALIDO, "PCKardexCHP_VerificaDatos", msg
            End If
            If pck.IdAsignado = mProps.IdAsignado Then
                v = v + pck.Debe + pck.Haber
            End If
        Next ix
        'Si el valor está más que el saldo del documento asignado
        If mProps.objGNComprobante.GNTrans.CodPantalla <> "TSIDEP" Then
            Set pcd = RecuperaPCDocAsignado
            
            If v > pcd.Saldo And pcd.id Then
                Err.Raise ERR_INVALIDO, "PCKardexCHP_VerificaDatos", _
                    "Valor que Cobra/Paga es mayor al saldo del documento " & vbCr & vbCr & _
                    pcd.NombreProvCli & " " & pcd.Trans & " " & pcd.Doc & vbCr & _
                    "Saldo es de: " & pcd.Saldo & " " & pcd.CodMoneda & vbCr & vbCr & _
                    "Cobro/Pago de : " & v & " " & mProps.objGNComprobante.CodMoneda
            End If
        Else
            Set pcd = RecuperaPCDocAsignadoOri
            If v > pcd.Saldo And pcd.id Then
                Err.Raise ERR_INVALIDO, "PCKardexCHP_VerificaDatos", _
                    "Valor que Cobra/Paga es mayor al saldo del documento " & vbCr & vbCr & _
                    pcd.NombreProvCli & " " & pcd.Trans & " " & pcd.Doc & vbCr & _
                    "Saldo es de: " & pcd.Saldo & " " & pcd.CodMoneda & vbCr & vbCr & _
                    "Cobro/Pago de : " & v & " " & mProps.objGNComprobante.CodMoneda
            End If
        
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
                                    Err.Raise ERR_INVALIDO, "GNComprobante.RemovePCKardexCHP", _
                                    "Con la forma de cobro " & mProps.CodForma & " no puede seleccionar el precio " & nump & Chr(13) & _
                                    " en la fila " & ix & " código Item: " & mProps.objGNComprobante.IVKardex(ix).CodInventario
                                End If
                            End If
                        Next ix
                End If
        End If
    End If

End Sub

Friend Sub Grabar()
    Dim sql As String
    Dim id_antes As Long, rs As Recordset
    
    id_antes = mProps.id        'Guarda el id anterior
    
    VerificaDatos
    
    'Graba en PCKardexCHP
    sql = "SELECT * FROM PCKardexCHP WHERE 1=0"
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
        !IdAsignadoPCK = mProps.IdAsignadoPCK
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
        'Actualiza IdAsignado de PCKardexCHP relacionado
        '  para que apunte de nuevo a éste registro.
        sql = "UPDATE PCKardexCHP SET IdAsignado=" & mProps.id & _
              " WHERE IdAsignado=" & id_antes
#If DAOLIB Then
        mProps.objGNComprobante.Empresa.Database.Execute sql, dbFailOnError
#Else
        mProps.objGNComprobante.Empresa.Coneccion.Execute sql
#End If
    End If

    'Si asigna Id al campo IdAsignado de otros PCKardexCHPs en la misma transacción
    ' para descargar cobros/pagos al contado
    AsignaIDPCK

    Set rs = Nothing
End Sub

'Asigna IdAsignado en los PCKardexCHP que está en el mismo comprobante y que tengan ahora -1 como IdAsignado
' Es para descargar cobro/pago al contado
Private Sub AsignaIDPCK()
    Dim i As Long, pck As PCKardexCHP
    
    For i = 1 To mProps.objGNComprobante.CountPCKardexCHP
        Set pck = mProps.objGNComprobante.PCKardexCHP(i)
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
            If Not IsNull(!CodProvcli) Then mProps.CodProvcli = !CodProvcli
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
            If Not IsNull(!IdAsignadoPCK) Then mProps.IdAsignadoPCK = !IdAsignadoPCK
            
        End If
    End With
End Sub

Friend Function Clone() As PCKardexCHP
    Dim nuevo As PCKardexCHP
    
    Set nuevo = New PCKardexCHP
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
                Err.Raise ERR_NOHAYCODIGO, "PCKardexCHP_CodProvCli", MSGERR_NOHAYCODIGO
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
    
    'Cuando cambia el valor
    If value <> mProps.CodForma Then
        If Len(value) > 0 Then
            'Actualiza IdForma también
            sql = "SELECT IdForma FROM TSFormaCobroPago WHERE CodForma='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCKardexCHP_CodForma", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IDForma = rs!IDForma
            End If
            rs.Close
        Else
            mProps.IDForma = 0
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

Public Property Let IdAsignadoPCK(ByVal value As Long)
    mProps.IdAsignadoPCK = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get IdAsignadoPCK() As Long
    IdAsignadoPCK = mProps.IdAsignadoPCK
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
        sql = "SELECT pck.Id FROM PCKardexCHP pck INNER JOIN GNComprobante gc " & _
                        "ON pck.TransID = gc.TransID " & _
              "WHERE gc.CodTrans = '" & CodTrans & "' AND gc.NumTrans = " & NumTrans & _
                     " AND pck.Orden = " & Orden
        Set rs = Me.GNComprobante.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then
            mProps.IdAsignado = rs.Fields("Id")
            mProps.objGNComprobante.SetModificado
        Else
            Err.Raise ERR_NOHAYCODIGO, "PCKardexCHP", _
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
    sql = "SELECT Id, TransID FROM PCKardexCHP WHERE Guid = '" & GuidAsignado & "'"
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
            Err.Raise ERR_NOHAYCODIGO, "PCKardexCHP", _
                    "No se encuentra el documento asignado. " & vbCr & _
                    "(" & Me.GNComprobante.CodTrans & _
                          Me.GNComprobante.NumTrans & ") " & vbCr & GuidAsignado
        End If
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Function BuscarGuidEnTrans(ByVal GuidBuscar As String) As Boolean
    Dim i As Long, pck As PCKardexCHP
    
    BuscarGuidEnTrans = False
    For i = 1 To Me.GNComprobante.CountPCKardexCHP
        Set pck = Me.GNComprobante.PCKardexCHP(i)
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
        Err.Raise ERR_INVALIDO, "PCKardexCHP_NumLetra", MSGERR_INVALIDO
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
            Err.Raise ERR_INVALIDO, "PCKardexCHP.VerificaSaldo", _
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
                Err.Raise ERR_INVALIDO, "PCKardexCHP_VerificaSaldo", _
                    "No se puede asignar un valor mayor al del documento. " & vbCr & vbCr & _
                    "Valor del documento: " & vbTab & pcd.Valor & vbCr & _
                    "Valor asignado:        " & vbTab & value & _
                    " " & mProps.objGNComprobante.CodMoneda
            ElseIf value > pcd.Saldo + Vdoc Then
                'No puede ser mayor al saldo del documento asignado
'Debug.Print Me.GNComprobante.CodTrans, Me.GNComprobante.NumTrans
                Err.Raise ERR_INVALIDO, "PCKardexCHP.VerificaSaldo", _
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
        Err.Raise ERR_INVALIDO, "PCKardexCHP_Debe", MSGERR_INVALIDO
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
        Err.Raise ERR_INVALIDO, "PCKardexCHP_Haber", MSGERR_INVALIDO
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
        Err.Raise ERR_INVALIDO, "PCKardexCHP_NombreEmitante", MSGERR_INVALIDO
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

Public Function RecuperaPCDocAsignado() As PCDocAsignadoCHP
    Dim pcd As PCDocAsignadoCHP

    If mProps.IdAsignado = 0 Then Exit Function
    
    Set pcd = New PCDocAsignadoCHP
    Set pcd.PCKardexCHP = Me
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
    Dim pck As PCKardexCHP, ix As Long, c As Currency
    
    For ix = 1 To Me.GNComprobante.CountPCKardexCHP
        Set pck = Me.GNComprobante.PCKardexCHP(ix)
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

    'Busca en la PCKardexCHP registros asignados que sea de otro comprobante
    sql = "SELECT Id FROM PCKardexCHP WHERE IdAsignado=" & mProps.id & _
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
'Recibe guid y lo busca en PCKardexCHP
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
    
    sql = "SELECT Id FROM PCKardexCHP WHERE guid='" & mProps.Guid & "'"
    Set rs = Me.GNComprobante.Empresa.OpenRecordset(sql)
    If rs.EOF Then
'No debe generar ningún error porque cuando es transacción nueva no ha de existir
'        Err.Raise ERR_NOHAYCODIGO, "PCKardexCHP.SetIdFromGuid", _
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


    'Busca en la PCKardexCHP registros asignados que sea de otro comprobante
    'JEAA       modificado para que visualice cual es el comprobante
    sql = " SELECT gnc.codtrans, gnc.numtrans " & _
            " FROM PCKardexCHP pck INNER JOIN GNComprobante gnc " & _
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
                Err.Raise ERR_NOHAYCODIGO, "PCKardexCHP_CodCobrador", MSGERR_NOHAYCODIGO
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
                Err.Raise ERR_NOHAYCODIGO, "PCKardexCHP_CodVendedor", MSGERR_NOHAYCODIGO
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

Public Function RecuperaPCDocAsignadoT() As PCDocAsignadoCHP
    Dim pcd As PCDocAsignadoCHP

    If mProps.IdAsignado = 0 Then Exit Function
    
    Set pcd = New PCDocAsignado
    Set pcd.PCKardexCHP = Me
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
Public Property Let BandNoAfectaPCKardexCHP(ByVal value As Boolean)
    mProps.BandNoAfectaPCKardexCHP = value
End Property

Public Property Get BandNoAfectaPCKardexCHP() As Boolean
    BandNoAfectaPCKardexCHP = mProps.BandNoAfectaPCKardexCHP
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
                Err.Raise ERR_NOHAYCODIGO, "PCKardexCHP_CodBanco", MSGERR_NOHAYCODIGO
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
                Err.Raise ERR_NOHAYCODIGO, "PCKardexCHP_CodTarjeta", MSGERR_NOHAYCODIGO
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
        Err.Raise ERR_INVALIDO, "PCKardexCHP_NumCuenta", MSGERR_INVALIDO
    End If
    
    mProps.NumCuenta = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get NumCuenta() As String
    NumCuenta = Trim$(mProps.NumCuenta)
End Property

Public Property Let numCheque(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCKardexCHP_NumCheque", MSGERR_INVALIDO
    End If
    
    mProps.numCheque = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get numCheque() As String
    numCheque = Trim$(mProps.numCheque)
End Property

Public Property Let TitularCta(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "PCKardexCHP_TitularCta", MSGERR_INVALIDO
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
                If mProps.objGNComprobante.IVKardex(j).IdInventario = mProps.objGNComprobante.IVKardex(fila).idpadre Then
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
   sql = sql & " inner join PCKardexCHP P on g.transid=P.transid"
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
'    Dim sql As String, rs As Recordset
'    Dim msg As String
'    Dim pcg As PcGrupo
'    Dim pc As PCProvCli
'    Dim idG As Byte
'    'Cuando cambia el valor
'    If value <> mProps.RubroRol Then
'        If Len(value) > 0 Then
'        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ConAsiento")) > 0 Then
'            idG = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ConAsiento") + 1
'                If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SeContabilizaPor")) > 0 Then
'                    If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SeContabilizaPor") = 0 Then
'                        sql = "Select cp.*,e.tipo from cuentapersonal cp "
'                        sql = sql & "Inner join elemento e on e.idelemento= cp.idelemento "
'                        sql = sql & "Where codelemento = '" & value & "' AND cp.idempleado = " & mProps.IdProvCli
'                    ElseIf gobjMain.EmpresaActual.GNOpcion.ObtenerValor("SeContabilizaPor") = 1 Then
'                        Set pc = gobjMain.EmpresaActual.RecuperaPCProvCli(mProps.IdProvCli)
'                        Select Case idG
'                            Case 1: Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, pc.CodGrupo1, 4)
'                            Case 2: Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, pc.CodGrupo2, 4)
'                            Case 3: Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, pc.CodGrupo3, 4)
'                            Case 4: Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, pc.CodGrupo4, 4)
'                        End Select
'                            sql = "Select cp.*,e.tipo from cuentapcgrupo cp "
'                            sql = sql & "Inner join elemento e on e.idelemento= cp.idelemento "
'                            sql = sql & " Inner  join pcgrupo" & idG & " pcg on pcg.idgrupo" & idG & " = cp.idpcgrupo "
'                            sql = sql & "Where codelemento = '" & value & "' AND cp.idpcgrupo = " & pcg.IdGrupo
'                            Set pc = Nothing
'                            Set pcg = Nothing
'                    End If
'                End If
'        End If
'            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
'            If rs.EOF Then
'                msg = "El siguiente Elemento " & value & " no tiene asignado una cuenta contable, " & vbCr & _
'                "por lo que no se puede generar ésta transacción. " & vbCr & vbCr & _
'                "Para poder Continuar, por favor habra datos del dicho Elemento" & vbCr & _
'                "y asigne una cuenta contable correspondiente."
'                Err.Raise ERR_NOHAYCODIGO, "PCKardexCHP_RubroRol", msg
'                Exit Property
'            Else
'                Select Case rs!Tipo
'                    Case 7: mProps.auxIdCuentaXIII = rs!IdCuenta  '-> decimo tercero
'                    Case 8: mProps.auxIdCuentaXIV = rs!IdCuenta ' -> decimo cuarto
'                    Case 9: mProps.auxIdCuentaVac = rs!IdCuenta ' -> vac
'                    Case 5: mProps.auxIdCuentaNetoRol = rs!IdCuenta  ' -> netorol
'                End Select
'                mProps.EsRubroRol = True
'            End If
'        End If
'            rs.Close
'    Else
'        mProps.auxIdCuentaXIII = 0
'        mProps.auxIdCuentaXIV = 0
'        mProps.auxIdCuentaVac = 0
'        mProps.auxIdCuentaNetoRol = 0
'        mProps.EsRubroRol = False
'    End If
'    Set rs = Nothing
'    mProps.objGNComprobante.SetModificado
    Dim sql As String, rs As Recordset, rsR As Recordset
    Dim msg As String
    Dim pcg As PcGrupo
 '   Dim pc As PCProvCli
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
                        
                        'Set pc = gobjMain.EmpresaActual.RecuperaEmpleado(mProps.IdProvCli)
                        sql = "Select top 1 idgrupo1,idgrupo2,idgrupo3,idgrupo4 from roldetalle where transid = " & mProps.objGNComprobante.transid
                        sql = sql & " And IdEmpleado =" & mProps.IdProvcli
                        Set rsR = gobjMain.EmpresaActual.OpenRecordset(sql)
                        If Not rsR.EOF Then
                            Select Case idG
'                                Case 1: Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, pc.CodGrupo1, 4)
'                                Case 2: Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, pc.CodGrupo2, 4)
'                                Case 3: Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, pc.CodGrupo3, 4)
'                                Case 4: Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, pc.CodGrupo4, 4)
                                Case 1: Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, rsR!IdGrupo1, 4)
                                Case 2: Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, rsR!IdGrupo2, 4)
                                Case 3: Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, rsR!IdGrupo3, 4)
                                Case 4: Set pcg = gobjMain.EmpresaActual.RecuperaPCGrupoOrigen(idG, rsR!IdGrupo4, 4)
                            End Select
                        End If
                            sql = "Select cp.*,e.tipo from cuentapcgrupo cp "
                            sql = sql & "Inner join elemento e on e.idelemento= cp.idelemento "
                            sql = sql & " Inner  join pcgrupo" & idG & " pcg on pcg.idgrupo" & idG & " = cp.idpcgrupo "
                            sql = sql & "Where codelemento = '" & value & "' AND cp.idpcgrupo = " & pcg.IdGrupo
                            Set rsR = Nothing
                            'Set pc = Nothing
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




Public Function RecuperaPCDocAsignadoOri() As PCDocAsignadoCHP
    Dim pcd As PCDocAsignadoCHP

    If mProps.IdAsignado = 0 Then Exit Function
    
    Set pcd = New PCDocAsignadoCHP
    Set pcd.PCKardexCHP = Me
    pcd.RecuperarOri mProps.IdAsignado
    Set RecuperaPCDocAsignadoOri = pcd
    Set pcd = Nothing
End Function


Public Function RecuperaPCDocAsignadoOriginal() As PCDocAsignadoCHP
    Dim pcd As PCDocAsignadoCHP

    If mProps.IdAsignadoPCK = 0 Then Exit Function
    
    Set pcd = New PCDocAsignadoCHP
    Set pcd.PCKardexCHP = Me
    pcd.RecuperarOriginal mProps.IdAsignadoPCK
    Set RecuperaPCDocAsignadoOriginal = pcd
    Set pcd = Nothing
End Function

Public Sub SetIdAsignadoPCKPorGuid(ByVal GuidAsignado As String)
    Dim sql As String, rs As Recordset
    
    'Busca Guid del doc. asignado
    sql = "SELECT Id, TransID FROM PCKardexchp WHERE Guid = '" & GuidAsignado & "'"
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
            Err.Raise ERR_NOHAYCODIGO, "PCKardexCHP", _
                    "No se encuentra el documento asignado. " & vbCr & _
                    "(" & Me.GNComprobante.CodTrans & _
                          Me.GNComprobante.NumTrans & ") " & vbCr & GuidAsignado
        End If
    End If
    
    rs.Close
    Set rs = Nothing
End Sub


Public Function RecuperaPCDocAsignadoOriginalCHP() As PCDocAsignadoCHP
    Dim pcd As PCDocAsignadoCHP

    If mProps.IdAsignado = 0 Then Exit Function
    
    
    Set pcd = New PCDocAsignadoCHP
    Set pcd.PCKardexCHP = Me
    pcd.RecuperarOri mProps.IdAsignado
    Set RecuperaPCDocAsignadoOriginalCHP = pcd
    Set pcd = Nothing
End Function

Public Sub SetIdAsignadoPCKPorGuidCHP(ByVal GuidAsignado As String)
    Dim sql As String, rs As Recordset
    
    'Busca Guid del doc. asignado
    sql = "SELECT Id, TransID FROM PCKardex WHERE Guid = '" & GuidAsignado & "'"
    Set rs = Me.GNComprobante.Empresa.OpenRecordset(sql)
    If Not rs.EOF Then
        'Si  el documento asignado NO está en la misma transacción
        If rs.Fields("TransID") <> Me.GNComprobante.transid Then
            mProps.IdAsignadoPCK = rs.Fields("Id")
        Else
            'Si el doc.asignado es de la misma transacción como en caso de pago/cobro a CONTADO
            'IdAsignado queda pendiente hasta que grabe la transacción
            mProps.IdAsignadoPCK = -1
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
            Err.Raise ERR_NOHAYCODIGO, "PCKardexCHP", _
                    "No se encuentra el documento asignado. " & vbCr & _
                    "(" & Me.GNComprobante.CodTrans & _
                          Me.GNComprobante.NumTrans & ") " & vbCr & GuidAsignado
        End If
    End If
    
    rs.Close
    Set rs = Nothing
End Sub


