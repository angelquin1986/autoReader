VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCProvCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Empresa"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type T_PROPS
    IdProvcli As Long
    CodProvcli As String
    Nombre As String
    BandProveedor As Boolean
    BandCliente As Boolean
    IdCuentaContable As Long
    IdCuentaContable2 As Long
    CodCuentaContable As String
    CodCuentaContable2 As String        '*** MAKOTO 12/feb/01 Agregado
    Direccion1 As String
    Direccion2 As String
    CodPostal As String
    Ciudad As String
    Provincia As String
    Pais As String
    Telefono1 As String
    Telefono2 As String
    Telefono3 As String
    Fax As String
    RUC As String
    EMail As String
    LimiteCredito As Currency
    banco As String
    NumCuenta  As String
    Swit As String
    DirecBanco As String
    TelBanco As String
    IdVendedor As Long
    CodVendedor As String
    IdGrupo1 As Long
    CodGrupo1 As String
    IdGrupo2 As Long
    CodGrupo2 As String
    IdGrupo3 As Long
    CodGrupo3 As String
    '-agregado AUC 22/09/2005
    IdGrupo4 As Long
    CodGrupo4 As String
   '----
    Estado As Byte
    TipoDocumento As String             'Agregado Alex: Sept/2002
    TipoComprobante As String           'Agregado Alex: Sept/2002
    NumAutSRI As String                 'Agregado Alex: Sept/2002
    
    NombreAlterno As String             '***Agregado. 05/sep/2003. Angel
    FechaNacimiento As Date             '***Agregado. 05/sep/2003. Angel
    FechaEntrega As Date                '***Agregado. 05/sep/2003. Angel
    FechaExpiracion As Date             '***Agregado. 05/sep/2003. Angel
    TotalDebe As Currency               '***Agregado. 08/sep/2003. Angel
    TotalHaber As Currency              '***Agregado. 08/sep/2003. Angel
        
    fechagrabado As Date
    Observacion As String         'Agregado AUC 31/08/2005
    TipoProvCli As String  ' jeaa 17/12/2005
    ReferidoPor As String         'Agregado jeaa 08/12/2006

    IdTipoDocumento As Long
    codtipoDocumento As String
    BandEmpresaPublica As Boolean   'jeaa 16/01/2008 cambio por EQUIDAD TRIBUTARIA

    IdGarante As Long
    CodGarante As String
    BandGarante As Boolean
    DiasPlazo As Currency 'AUC 11/10/2010
    
    CodigoUsuario As String  'para Cierre de caja
    bandEmpleado As Boolean   'AUC 18/07/2011

    IdProvincia As Long
    CodProvincia As String
    IdCanton As Long
    CodCanton As String
    IDParroquia As Long
    CodParroqui As String
    FechaCreacion As Date

    FechaLote1 As Date 'facturacion x lote
    FechaLote2 As Date
    BandLote As Boolean
    
    NumSerie As String
    NumPunto As String
    FechaCaducidad As Date
    IdDiasCredito As Long
    CodDiasCredito As String
    IsCopy As Boolean
    NumPagos As Integer
    Intervalo As Integer
    PorDescNeto As Currency
    PorDescPP As Currency
     IdTurno As Long
    Codturno As String
    TipoTurno As Byte
    BandRUCValido As Boolean
    
    CodigoReloj As String
    BandRelacionado As Boolean
    
    BandSolicitudCredito As Boolean
    DescripcionDetalle As String        'memo
    NombreConyuge As String
    RucConyuge As String
    BandSuperGarante As Boolean
    BandOmitirPendientes As Boolean
    
    TipoSujeto As String
    Sexo As String
    EstadoCivil As String
    OrigenIngresos As String
    idSolCredito As Long
    Password As String
    CodigoRelojP As String
    BandOmitirDinardap As Boolean
    
    idpadre As Long
    CodPadre As String
    IdMadre As Long
    CodMadre As String
    IdOtro As Long
    CodOtro As String
    DireccionAM As String
    DireccionPM As String
    BandMismaDireccion As Boolean
    NumRepresentante As Integer
    NumEmiteFactura As Integer
    IdInventario As Long
    CodInventario As String
    IdVehiculoAM As Long
    CodVehiculoAM As String
    IdVehiculoPM As Long
    CodVehiculoPM As String
    BandTransporte As Boolean
    FechaInscripcion As Date
    IdCobrador As Long
    CodCobrador As String
    BandRebate As Boolean
    IdPais As Long
    CodPais As String
    CordenadaX As Long
   CordenadaY As Long
   posGoogleMaps As String
   IDCallePri As Long
   CodCallePri As String
   IdCalleSec As Long
   CodCalleSec As String
    IdGarante2 As Long
    CodGarante2 As String
    NumCasa As String
    bandCompensaIVA As Boolean
    BandEsConsumidorFinal As Boolean
    BandIVCodigoIgual As Boolean
    
    IdCuentaContableGasto As Long
    CodCuentaContableGasto As String
    Garantia As String
    Predio As String
    BandNCPP As Boolean
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mProps As T_PROPS
Private mcolContactos As Collection
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean
Private mcolDetalleItem As Collection  'Coleccion de detalles proveedor AUC 22/11/05
Private mcolItemsDetalle As Collection  'Coleccion de detalles Items
Private mcolAgencia As Collection



Public Sub Grabar()
    Dim sql As String, fgrabado As Date, i As Long
    Dim rs As Recordset
    Dim rsSuc As Recordset, gnSuc As GNSucursal
    Dim NumGrupo As Integer, ivg As ivgrupo, campo As String, k As Long, n As Byte
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
       
    'Verifica si no repite el código
    sql = "SELECT CodProvCli FROM PCProvCli " & _
          "WHERE (CodProvCli='" & mProps.CodProvcli & "') AND " & _
                "(IdProvCli<>" & mProps.IdProvcli & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "PCProvCli", MSGERR_REPITECODIGO
    End If
    rs.Close
'*** Eliminado porque puse dentro de una funcion VerificaRUCRepetido
'    '*** Agregado Oliver 30/jun/2003 para verificar control de duplicidad de RUC o CI
'    'Verifica si se repite el RUC
'    sql = "SELECT  RUC FROM PCProvCli " & _
'          "WHERE (RUC='" & mProps.RUC & "') AND " & _
'               " ((BandCliente=" & IIf(mProps.BandCliente, 1, 0) & ") OR " & _
'               " (BandProveedor=" & IIf(mProps.BandProveedor, 1, 0) & ")) AND " & _
'               " (IdProvCli<>" & mProps.IdProvCli & ")"
'    Set rs = mProps.objEmpresa.OpenRecordset(sql)
'    If Not rs.EOF Then      'Si ya existe el codigo
'        Err.Raise ERR_REPITECODIGO, "PCProvCli", MSGERR_REPITERUC
'    End If
'    rs.Close
    
''''    If mProps.BandCliente Then
''''        If Len(mProps.objEmpresa.GNOpcion.ObtenerValor("BandClienteCIRUC")) > 0 Then
''''            If mProps.objEmpresa.GNOpcion.ObtenerValor("BandClienteCIRUC") = "1" Then
''''                If mProps.VerificaCIyRUCRepetido(mProps.RUC, rucRepetido, Nombre) And Not mProps.BandEsConsumidorFinal Then
''''                    If Len(rucRepetido) = 13 Then
''''                        MsgBox "Ya existe creado un cliente con RUC: " & rucRepetido & "-" & Nombre, vbInformation
''''                    ElseIf Len(rucRepetido) = 10 Then
''''                        MsgBox "Ya existe creado un cliente con CI: " & rucRepetido & "-" & Nombre, vbInformation
''''                    End If
''''                    Exit Function
''''                End If
''''            End If
''''        End If
''''
    
    
    On Error GoTo ErrTrap

#If DAOLIB Then
    gobjWS.BeginTrans                           '======= Comienza la transacción
#Else
    mProps.objEmpresa.Coneccion.BeginTrans      '======= Comienza la transacción
#End If
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM PCProvCli WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Elimina todos los Contactos
        sql = "DELETE FROM PCContacto WHERE IdProvCli=" & mProps.IdProvcli
        mProps.objEmpresa.Execute sql, True
    
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM PCProvCli WHERE IdProvCli=" & mProps.IdProvcli
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "PCProvCli_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en la base
    With rs
        !CodProvcli = mProps.CodProvcli
        ValidaCodigo mProps.Nombre
        If mProps.BandCliente Then
            !Nombre = QuitarEspacios(mProps.Nombre)
        Else
            !Nombre = (mProps.Nombre)
        End If
        !BandProveedor = mProps.BandProveedor
        !BandCliente = mProps.BandCliente
        !IdCuentaContable = mProps.IdCuentaContable
        !IdCuentaContable2 = mProps.IdCuentaContable2
        !Direccion1 = mProps.Direccion1
        !Direccion2 = mProps.Direccion2
        !CodPostal = mProps.CodPostal
        !Ciudad = mProps.Ciudad
        !Provincia = mProps.Provincia
        !Pais = mProps.Pais
        !Telefono1 = mProps.Telefono1
        !Telefono2 = mProps.Telefono2
        !Telefono3 = mProps.Telefono3
        !Fax = mProps.Fax
        If Not !RUC = mProps.RUC Then
            Me.Empresa.GrabaAccionModulo !RUC, mProps.RUC, "RUC", "MOD_CLI", "CL"
        End If
        !RUC = mProps.RUC
        !EMail = mProps.EMail
        If Not !LimiteCredito = mProps.LimiteCredito Then
            Me.Empresa.GrabaAccionModulo !LimiteCredito, mProps.LimiteCredito, "LimiteCredito", "MOD_CLI", "CL"
        End If
        !LimiteCredito = mProps.LimiteCredito
        !banco = mProps.banco
        !NumCuenta = mProps.NumCuenta
        !Swit = mProps.Swit
        !DirecBanco = mProps.DirecBanco
        !TelBanco = mProps.TelBanco
        !IdVendedor = mProps.IdVendedor
        !IdGrupo1 = mProps.IdGrupo1
        !IdGrupo2 = mProps.IdGrupo2
        !IdGrupo3 = mProps.IdGrupo3
        
        !IdGrupo4 = mProps.IdGrupo4 'agregado AUC 22/09/2005
        
        If Not !Estado = mProps.Estado Then
            Me.Empresa.GrabaAccionModulo !Estado, mProps.Estado, "Estado", "MOD_CLI", "CL"
        End If
        !Estado = mProps.Estado
        !TipoDocumento = mProps.TipoDocumento           'Agregado Alex: Sept/2002
        !TipoComprobante = mProps.TipoComprobante       'Agregado Alex: Sept/2002
        !NumAutSRI = mProps.NumAutSRI                   'Agregado Alex: Sept/2002
        If mProps.BandCliente Then
            !NombreAlterno = QuitarEspacios(mProps.NombreAlterno)
        Else
            !NombreAlterno = mProps.NombreAlterno           '***Agregado. 05/sep/2003. Angel
        End If
        !FechaNacimiento = mProps.FechaNacimiento       '***Agregado. 05/sep/2003. Angel
        !FechaEntrega = mProps.FechaEntrega             '***Agregado. 05/sep/2003. Angel
        !FechaExpiracion = mProps.FechaExpiracion       '***Agregado. 05/sep/2003. Angel
        !TotalDebe = mProps.TotalDebe                   '***Agregado. 08/sep/2003. Angel
        !TotalHaber = mProps.TotalHaber                 '***Agregado. 08/sep/2003. Angel
        !Observacion = mProps.Observacion  'Agregado AUC 31/08/2005
        !TipoProvCli = mProps.TipoProvCli       ' jeaa 17/12/2005
        !ReferidoPor = mProps.ReferidoPor  'Agregado jeaa 12/08/2006
        !IdTipoDocumento = mProps.IdTipoDocumento   'Agregado jeaa 15/05/2007
        !BandEmpresaPublica = mProps.BandEmpresaPublica   'jeaa 16/01/2008 cambio por EQUIDAD TRIBUTARIA
        !IdGarante = mProps.IdGarante
        !BandGarante = mProps.BandGarante
        !bandEmpleado = mProps.bandEmpleado
        If Not !DiasPlazo = mProps.DiasPlazo Then
            Me.Empresa.GrabaAccionModulo !DiasPlazo, mProps.DiasPlazo, "DiasPlazo", "MOD_CLI", "CL"
        End If
        !DiasPlazo = mProps.DiasPlazo
        
        If Not !CodigoUsuario = mProps.CodigoUsuario Then
            Me.Empresa.GrabaAccionModulo !CodigoUsuario, mProps.CodigoUsuario, "CodigoUsuario", "MOD_CLI", "CL"
        End If
        
        !CodigoUsuario = mProps.CodigoUsuario
        !IdProvincia = mProps.IdProvincia
        !IdCanton = mProps.IdCanton
        !IDParroquia = mProps.IDParroquia
        !FechaCreacion = mProps.FechaCreacion
        !FechaLote1 = mProps.FechaLote1
        !FechaLote2 = mProps.FechaLote2
        !BandLote = mProps.BandLote
        !bandEmpleado = mProps.bandEmpleado
        !NumPunto = mProps.NumPunto
        !NumSerie = mProps.NumSerie
        !FechaCaducidad = mProps.FechaCaducidad
        !IdDiasCredito = mProps.IdDiasCredito
        !NumPagos = mProps.NumPagos
        !Intervalo = mProps.Intervalo
        !PorDescNeto = mProps.PorDescNeto
        !PorDescPP = mProps.PorDescPP
        !BandRUCValido = mProps.BandRUCValido
        !IdTurno = mProps.IdTurno
        !TipoTurno = mProps.TipoTurno
        !BandRelacionado = mProps.BandRelacionado
        
        If Not !BandOmitirPendientes = mProps.BandOmitirPendientes Then
            Me.Empresa.GrabaAccionModulo !BandOmitirPendientes, mProps.BandOmitirPendientes, "BandOmitirPendientes", "MOD_CLI", "CL"
        End If
        !BandOmitirPendientes = mProps.BandOmitirPendientes
        
        !TipoSujeto = mProps.TipoSujeto
        !Sexo = mProps.Sexo
        !EstadoCivil = mProps.EstadoCivil
        !OrigenIngresos = mProps.OrigenIngresos
        !BandSuperGarante = mProps.BandSuperGarante
        !BandSolicitudCredito = mProps.BandSolicitudCredito
         .Fields("DescripcionDetalle").AppendChunk mProps.DescripcionDetalle
        !Password = mProps.Password
        !CodigoReloj = mProps.CodigoRelojP
        !fechagrabado = Now
        !BandOmitirDinardap = BandOmitirDinardap
         !IdCobrador = mProps.IdCobrador
        GrabarItemsDetalle (mProps.CodProvcli)
        !BandRebate = mProps.BandRebate
         !IdPais = mProps.IdPais
         !IDCallePri = mProps.IDCallePri
         !IdCalleSec = mProps.IdCalleSec
         
        !CordenadaX = mProps.CordenadaX
        !CordenadaY = mProps.CordenadaY
        !posGoogleMaps = mProps.posGoogleMaps
        !IdGarante2 = mProps.IdGarante2
        !NumCasa = mProps.NumCasa
        !bandCompensaIVA = mProps.bandCompensaIVA
        !BandEsConsumidorFinal = mProps.BandEsConsumidorFinal
        !BandIVCodigoIgual = mProps.BandIVCodigoIgual
        !IdCuentaContableGasto = mProps.IdCuentaContableGasto
        !Garantia = mProps.Garantia
        !Predio = mProps.Predio
        !BandNCPP = mProps.BandNCPP
        GrabarItemsDetalle (mProps.CodProvcli)
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        fgrabado = !fechagrabado
        mProps.IdProvcli = !IdProvcli
        .Close
    End With
    
    'Graba todos los PCContacto
    For i = 1 To CountContacto
        Contactos(i).Grabar
    Next i
    'Graba todos los PCAgencia
'    For i = 1 To CountAgencia
'        Agencia(i).Grabar
'    Next i
#If DAOLIB Then
    gobjWS.CommitTrans                  '======== Completa la transacción
#Else
    mProps.objEmpresa.Coneccion.CommitTrans      '======== Completa la transacción
#End If
    mProps.fechagrabado = fgrabado
    GrabarIVDetalle 'guarda configuracion para facturacion x lote x cliente
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_PC", "Proveedor/Cliente Nuevo: " & _
                                                         mProps.CodProvcli, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_PC", "Proveedor/Cliente Modificado: " & _
                                                         mProps.CodProvcli, "DT"
    End If
    
        'actualiza Inventario en Bases de Datos Sucursales
    sql = "SELECT * FROM GnSucursal WHERE bandvalida=1"
    Set rsSuc = mProps.objEmpresa.OpenRecordset(sql)
    
    For i = 1 To rsSuc.RecordCount
        Set gnSuc = mProps.objEmpresa.RecuperaGNSucursal(rsSuc.Fields("CodSucursal"))
        If gnSuc.BandCliente Or gnSuc.BandProveedor Then
            GrabarEnSucursales gnSuc.BaseDatos, gnSuc.servidor
        End If
        rsSuc.MoveNext
    Next i
    Set gnSuc = Nothing
    Set rsSuc = Nothing

    ' si esta activado opcion de anexar a ivgrupo
    If Len(mProps.objEmpresa.GNOpcion.ObtenerValor("NumAnexarCatProvxIVG")) > 0 Then
        NumGrupo = mProps.objEmpresa.GNOpcion.ObtenerValor("NumAnexarCatProvxIVG")
    Else
        NumGrupo = 1
    End If
    n = NumGrupo
    If Len(mProps.objEmpresa.GNOpcion.ObtenerValor("AnexarCatProvxIVG")) > 0 Then
        If mProps.objEmpresa.GNOpcion.ObtenerValor("AnexarCatProvxIVG") = "1" Then
            If mProps.BandProveedor Then
                If mbooEsNuevo Then
                    sql = "SELECT * FROM ivgrupo" & NumGrupo & " WHERE 1=0"
                    Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
                    rs.AddNew
                    Set ivg = mProps.objEmpresa.CreaIVGrupo(n)
                Else
                    Set ivg = mProps.objEmpresa.RecuperaIVGrupo(n, mProps.CodProvcli)
                End If
                If Not ivg Is Nothing Then
                With ivg
                    .CodGrupo = mProps.CodProvcli
                    .Descripcion = mProps.Nombre
                    .BandValida = IIf(mProps.Estado = 0, 1, 0)
                    .Grabar
                End With
                End If
            End If
        End If
    End If


    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PROVCLI
    Exit Sub
ErrTrap:
#If DAOLIB Then
    gobjWS.Rollback                   '======== Anula la transacción
#Else
    mProps.objEmpresa.Coneccion.RollbackTrans  '======== Anula la transacción
#End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

    Friend Sub Recuperar(key As Variant)
    Dim sql As String, i As Long
    Dim rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM vwPCProvCli WHERE CodProvCli='" & key & "'"
    Else
        sql = "SELECT * FROM vwPCProvCli WHERE IdProvCli=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IdProvcli = !IdProvcli
        mProps.CodProvcli = !CodProvcli
        If Not IsNull(!Nombre) Then mProps.Nombre = !Nombre
        If Not IsNull(!BandProveedor) Then mProps.BandProveedor = !BandProveedor
        If Not IsNull(!BandCliente) Then mProps.BandCliente = !BandCliente
        If Not IsNull(!IdCuentaContable) Then mProps.IdCuentaContable = !IdCuentaContable
        If Not IsNull(!IdCuentaContable2) Then mProps.IdCuentaContable2 = !IdCuentaContable2
        If Not IsNull(!CodCuentaContable) Then mProps.CodCuentaContable = !CodCuentaContable
        If Not IsNull(!CodCuentaContable2) Then mProps.CodCuentaContable2 = !CodCuentaContable2
        If Not IsNull(!Direccion1) Then mProps.Direccion1 = !Direccion1
        If Not IsNull(!Direccion2) Then mProps.Direccion2 = !Direccion2
        If Not IsNull(!CodPostal) Then mProps.CodPostal = !CodPostal
        If Not IsNull(!Ciudad) Then mProps.Ciudad = !Ciudad
        If Not IsNull(!Provincia) Then mProps.Provincia = !Provincia
        If Not IsNull(!Pais) Then
            If Len(!Pais) = 0 Then
                mProps.Pais = "ECUADOR"
            Else
                mProps.Pais = !Pais
            End If
        End If
        If Not IsNull(!Telefono1) Then mProps.Telefono1 = !Telefono1
        If Not IsNull(!Telefono2) Then mProps.Telefono2 = !Telefono2
        If Not IsNull(!Telefono3) Then mProps.Telefono3 = !Telefono3
        If Not IsNull(!Fax) Then mProps.Fax = !Fax
        If Not IsNull(!RUC) Then mProps.RUC = !RUC
        If Not IsNull(!EMail) Then mProps.EMail = !EMail
        If Not IsNull(!LimiteCredito) Then mProps.LimiteCredito = !LimiteCredito
        If Not IsNull(!banco) Then mProps.banco = !banco
        If Not IsNull(!NumCuenta) Then mProps.NumCuenta = !NumCuenta
        If Not IsNull(!Swit) Then mProps.Swit = !Swit
        If Not IsNull(!DirecBanco) Then mProps.DirecBanco = !DirecBanco
        If Not IsNull(!TelBanco) Then mProps.TelBanco = !TelBanco
        If Not IsNull(!IdVendedor) Then mProps.IdVendedor = !IdVendedor
        If Not IsNull(!CodVendedor) Then mProps.CodVendedor = !CodVendedor
        If Not IsNull(!IdGrupo1) Then mProps.IdGrupo1 = !IdGrupo1
        If Not IsNull(!CodGrupo1) Then mProps.CodGrupo1 = !CodGrupo1
        If Not IsNull(!IdGrupo2) Then mProps.IdGrupo2 = !IdGrupo2
        If Not IsNull(!CodGrupo2) Then mProps.CodGrupo2 = !CodGrupo2
        
        If Not IsNull(!IdGrupo3) Then mProps.IdGrupo3 = !IdGrupo3
        If Not IsNull(!CodGrupo3) Then mProps.CodGrupo3 = !CodGrupo3
        
        If Not IsNull(!IdGrupo4) Then mProps.IdGrupo4 = !IdGrupo4 'Agregadoa AUC 22/09/2005
        If Not IsNull(!CodGrupo4) Then mProps.CodGrupo4 = !CodGrupo4
        
        If Not IsNull(!Estado) Then mProps.Estado = !Estado
        If Not IsNull(!TipoDocumento) Then mProps.TipoDocumento = !TipoDocumento          'Agregado Alex Sept/2002
        If Not IsNull(!TipoComprobante) Then mProps.TipoComprobante = !TipoComprobante           'Agregado Alex Sept/2002
        If Not IsNull(!NumAutSRI) Then mProps.NumAutSRI = !NumAutSRI             'Agregado Alex Sept/2002
        If Not IsNull(!NombreAlterno) Then mProps.NombreAlterno = !NombreAlterno          '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!FechaNacimiento) Then mProps.FechaNacimiento = !FechaNacimiento    '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!FechaEntrega) Then mProps.FechaEntrega = !FechaEntrega             '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!FechaExpiracion) Then mProps.FechaExpiracion = !FechaExpiracion    '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!TotalDebe) Then mProps.TotalDebe = !TotalDebe                      '***Agregado. 08/sep/2003. Angel
        If Not IsNull(!TotalHaber) Then mProps.TotalHaber = !TotalHaber                   '***Agregado. 08/sep/2003. Angel
        If Not IsNull(!Observacion) Then mProps.Observacion = !Observacion      'Agregado AUC 31/08/2005
        If Not IsNull(!TipoProvCli) Then mProps.TipoProvCli = Trim$(!TipoProvCli)  ' jeaa 17/12/2005
        If Not IsNull(!ReferidoPor) Then mProps.ReferidoPor = !ReferidoPor      'Agregado jeaa 12/08/2006
        'jeaa 15/05/2007
        If Not IsNull(!IdTipoDocumento) Then mProps.IdTipoDocumento = !IdTipoDocumento 'Agregado jeaa 15/05/2007
        If Not IsNull(!codtipoDocumento) Then mProps.codtipoDocumento = !codtipoDocumento 'Agregado jeaa 15/05/2007
        If Not IsNull(!BandEmpresaPublica) Then mProps.BandEmpresaPublica = !BandEmpresaPublica   'jeaa 16/01/2008 cambio por EQUIDAD TRIBUTARIA
        If Not IsNull(!IdGarante) Then mProps.IdGarante = !IdGarante
        If Not IsNull(!CodGarante) Then mProps.CodGarante = !CodGarante
        If Not IsNull(!BandGarante) Then mProps.BandGarante = !BandGarante
        If Not IsNull(!DiasPlazo) Then mProps.DiasPlazo = !DiasPlazo
        If Not IsNull(!CodigoUsuario) Then mProps.CodigoUsuario = !CodigoUsuario
        If Not IsNull(!IdProvincia) Then mProps.IdProvincia = !IdProvincia
        If Not IsNull(!CodProvincia) Then mProps.CodProvincia = !CodProvincia
        If Not IsNull(!bandEmpleado) Then mProps.bandEmpleado = !bandEmpleado
        
'        Else
'            mProps.CodProvincia = mProps.objEmpresa.GNOpcion.CodProvinciaPre
'        End If
        If Not IsNull(!IdCanton) Then mProps.IdCanton = !IdCanton
        If Not IsNull(!CodCanton) Then mProps.CodCanton = !CodCanton
'        Else
'            mProps.CodCanton = mProps.objEmpresa.GNOpcion.CodCantonPre
'        End If

        If Not IsNull(!IDParroquia) Then mProps.IDParroquia = !IDParroquia
        If Not IsNull(!CodParroquia) Then mProps.CodParroqui = !CodParroquia
        If Not IsNull(!FechaCreacion) Then mProps.FechaCreacion = !FechaCreacion
        
        If Not IsNull(!FechaLote1) Then mProps.FechaLote1 = !FechaLote1
        If Not IsNull(!FechaLote2) Then mProps.FechaLote2 = !FechaLote2
        If Not IsNull(!BandLote) Then mProps.BandLote = !BandLote
        If Not IsNull(!bandEmpleado) Then mProps.bandEmpleado = !bandEmpleado
        If Not IsNull(!NumSerie) Then mProps.NumSerie = !NumSerie
        If Not IsNull(!NumPunto) Then mProps.NumPunto = !NumPunto
        If Not IsNull(!FechaCaducidad) Then mProps.FechaCaducidad = !FechaCaducidad
        If Not IsNull(!IdDiasCredito) Then mProps.IdDiasCredito = !IdDiasCredito
        If Not IsNull(!CodDiasCredito) Then mProps.CodDiasCredito = !CodDiasCredito
        If Not IsNull(!NumPagos) Then mProps.NumPagos = !NumPagos
        If Not IsNull(!Intervalo) Then mProps.Intervalo = !Intervalo
        If Not IsNull(!PorDescNeto) Then mProps.PorDescNeto = !PorDescNeto
        If Not IsNull(!PorDescPP) Then mProps.PorDescPP = !PorDescPP
        If Not IsNull(!BandRUCValido) Then mProps.BandRUCValido = !BandRUCValido
         If Not IsNull(!IdTurno) Then mProps.IdTurno = !IdTurno
        If Not IsNull(!Codturno) Then mProps.Codturno = !Codturno
        If Not IsNull(!TipoTurno) Then mProps.TipoTurno = !TipoTurno
        If Not IsNull(!BandRelacionado) Then mProps.BandRelacionado = !BandRelacionado
        If Not IsNull(!BandOmitirPendientes) Then mProps.BandOmitirPendientes = !BandOmitirPendientes

        If Not IsNull(!TipoSujeto) Then mProps.TipoSujeto = !TipoSujeto
        If Not IsNull(!Sexo) Then mProps.Sexo = !Sexo
        If Not IsNull(!EstadoCivil) Then mProps.EstadoCivil = !EstadoCivil
        If Not IsNull(!OrigenIngresos) Then mProps.OrigenIngresos = !OrigenIngresos
        If Not IsNull(!BandSolicitudCredito) Then mProps.BandSolicitudCredito = !BandSolicitudCredito
         If Not IsNull(!DescripcionDetalle) Then
            If .Fields("DescripcionDetalle").ActualSize > 0 Then
                mProps.DescripcionDetalle = .Fields("DescripcionDetalle").GetChunk(.Fields("DescripcionDetalle").ActualSize)
            End If
        End If
        If Not IsNull(!NombreConyuge) Then mProps.NombreConyuge = !NombreConyuge
        If Not IsNull(!RucConyuge) Then mProps.RucConyuge = !RucConyuge
        If Not IsNull(!Password) Then mProps.Password = !Password
        If Not IsNull(!CodigoReloj) Then mProps.CodigoRelojP = !CodigoReloj
        mProps.fechagrabado = !fechagrabado
        If Not IsNull(!BandOmitirDinardap) Then mProps.BandOmitirDinardap = !BandOmitirDinardap
         If Not IsNull(!IdCobrador) Then mProps.IdCobrador = !IdCobrador
        If Not IsNull(!CodCobrador) Then mProps.CodCobrador = !CodCobrador
        If Not IsNull(!BandRebate) Then mProps.BandRebate = !BandRebate
        If Not IsNull(!IdPais) Then mProps.IdPais = !IdPais
        If Not IsNull(!CodPais) Then mProps.CodPais = !CodPais
        If Not IsNull(!CordenadaX) Then mProps.CordenadaX = !CordenadaX
        If Not IsNull(!CordenadaY) Then mProps.CordenadaY = !CordenadaY
        If Not IsNull(!posGoogleMaps) Then mProps.posGoogleMaps = !posGoogleMaps
        
        If Not IsNull(!IDCallePri) Then mProps.IDCallePri = !IDCallePri
        If Not IsNull(!CodCallePri) Then mProps.CodCallePri = !CodCallePri
        If Not IsNull(!IdCalleSec) Then mProps.IdCalleSec = !IdCalleSec
        If Not IsNull(!CodCalleSec) Then mProps.CodCalleSec = !CodCalleSec
        
        If Not IsNull(!IdGarante2) Then mProps.IdGarante2 = !IdGarante2
        If Not IsNull(!CodGarante2) Then mProps.CodGarante2 = !CodGarante2
        
        If Not IsNull(!NumCasa) Then mProps.NumCasa = !NumCasa
        If Not IsNull(!bandCompensaIVA) Then mProps.bandCompensaIVA = !bandCompensaIVA
        If Not IsNull(!BandEsConsumidorFinal) Then mProps.BandEsConsumidorFinal = !BandEsConsumidorFinal
        If Not IsNull(!BandIVCodigoIgual) Then mProps.BandIVCodigoIgual = !BandIVCodigoIgual
        
        If Not IsNull(!IdCuentaContableGasto) Then mProps.IdCuentaContableGasto = !IdCuentaContableGasto
        If Not IsNull(!CodCuentaContableGasto) Then mProps.CodCuentaContableGasto = !CodCuentaContableGasto
        If Not IsNull(!Garantia) Then mProps.Garantia = !Garantia
        If Not IsNull(!Garantia) Then mProps.Predio = !Predio
        If Not IsNull(!BandNCPP) Then mProps.BandNCPP = !BandNCPP
        
        'CargaItemsDetalle
        .Close
    End With
    CargaDetalleItem

    If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "CATALI") > 0 Then
        RecuperarTransporte mProps.IdProvcli
    End If
    'Recupera los Contactos
    sql = "SELECT IdContacto FROM PCContacto WHERE IdProvCli=" & mProps.IdProvcli & " ORDER BY Orden"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddContacto
            Contactos(i).Recuperar !IdContacto
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
'        sql = "SELECT IdAgencia FROM PCAgencia WHERE IdProvCli=" & mProps.IdProvCli & " ORDER BY Orden"
'    Set rs = mProps.objEmpresa.OpenRecordset(sql)
'    With rs
'        Do Until .EOF
'            i = AddAgencia
'            Agencia(i).Recuperar !IdAgencia
'            .MoveNext
'        Loop
'        .Close
'    End With
'    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    Dim NumGrupo As Integer, ivg As ivgrupo, campo As String, k As Long, n As Byte
    
    'Verifica si existe Kardex
    sql = "SELECT IdProvCli FROM PCKardex WHERE IdProvCli=" & mProps.IdProvcli
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "PCProvCli.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing

    'Verifica si no esta asignado como Cliente
    sql = "SELECT IdClienteRef FROM Gncomprobante WHERE IdClienteRef=" & mProps.IdProvcli
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "PCProvCli.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing

    'Verifica si no esta asignado como Cliente
    sql = "SELECT IdProveedorRef FROM Gncomprobante WHERE IdProveedorRef=" & mProps.IdProvcli
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "PCProvCli.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing



    'Verifica si no esta asignado como garante
    sql = "SELECT IdGaranteRef FROM Gncomprobante WHERE IdGaranteRef=" & mProps.IdProvcli
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "PCProvCli.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing



    'Verifica si no esta asignado como garante
    sql = "SELECT IdProvcli FROM Pcprovcli WHERE IdGarante=" & mProps.IdProvcli
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "PCProvCli.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing

    'Verifica si no esta asignado en centro costo
    sql = "SELECT Idcliente FROM gncentrocosto WHERE Idcliente=" & mProps.IdProvcli
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "PCProvCli.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing


    'Verifica si no en el historial del item
    sql = "SELECT Idproveedor FROM IVProveedorDetalle WHERE Idproveedor=" & mProps.IdProvcli
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "PCProvCli.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing



    'Elimina los contactos
    sql = "DELETE FROM PCContacto WHERE IdProvCli=" & mProps.IdProvcli
    mProps.objEmpresa.Execute sql, True
    
    
    
        
    'Elimina el registro
    sql = "DELETE FROM PCProvCli WHERE CodProvCli='" & mProps.CodProvcli & "'"
    mProps.objEmpresa.Execute sql, True

    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_PC", "Proveedor/Cliente Eliminado: " & _
                                                         mProps.CodProvcli, "DT"
                                                         
    mbooEsNuevo = True
    mbooModificado = True
    
    ' si esta activado opcion de anexar a ivgrupo
    If Len(mProps.objEmpresa.GNOpcion.ObtenerValor("NumAnexarCatProvxIVG")) > 0 Then
        NumGrupo = mProps.objEmpresa.GNOpcion.ObtenerValor("NumAnexarCatProvxIVG")
    Else
        NumGrupo = 1
    End If
    n = NumGrupo
    If Len(mProps.objEmpresa.GNOpcion.ObtenerValor("AnexarCatProvxIVG")) > 0 Then
        If mProps.objEmpresa.GNOpcion.ObtenerValor("AnexarCatProvxIVG") = "1" Then
            If mProps.BandProveedor Then
                'Elimina el registro
                sql = "DELETE FROM ivgrupo" & NumGrupo & " WHERE Codgrupo" & NumGrupo & "='" & mProps.CodProvcli & "'"
                mProps.objEmpresa.Execute sql, True
            End If
        End If
    End If
    
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PROVCLI
End Sub

Public Function Clone() As PCProvCli
    Dim nuevo As PCProvCli, nc As PCContacto, i As Long
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New PCProvCli
    nuevo.Props = mProps
    
    'Copia los contactos al nuevo objeto
    For i = 1 To mcolContactos.Count
        Set nc = Contactos(i).Clone
        nuevo.AddContactoParaClone nc
    Next i
    
    Set Clone = nuevo
    Set nuevo = Nothing
    Set nc = Nothing
End Function

Friend Sub AddContactoParaClone(obj As PCContacto)
    Set obj.PCProvCli = Me
    mcolContactos.Add obj
End Sub

Public Function ListaContactos() As Variant
    Dim sql As String, rs As Recordset

    sql = "SELECT Cargo,Titulo,Nombre,Telefono1,Telefono2,EMail FROM PCContacto " & _
          " WHERE IdProvCli=" & mProps.IdProvcli & " ORDER BY Orden"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    Set ListaContactos = rs
    Set rs = Nothing
End Function


Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdProvcli = 0
    mbooModificado = True
End Property


'===== Manejo de colección de contactos
Public Function AddContacto() As Long
    Dim nuevo As PCContacto
    
    Set nuevo = New PCContacto
    Set nuevo.PCProvCli = Me
    mcolContactos.Add nuevo
    AddContacto = mcolContactos.Count     'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function

Public Sub RemoveContacto(ix As Variant)
    Dim obj As PCContacto, i As Long

    'Si llama con numero de indice del objeto
    If (VarType(ix) = vbLong) Or (VarType(ix) = vbInteger) Then
        mcolContactos.Remove ix      'Elimina el objeto
        
    'Si llama con la referencia al objeto que quiere eliminar
    Else
        'Busca el objeto dentro de la colección
        For i = mcolContactos.Count To 1 Step -1
            Set obj = mcolContactos.item(i)
            If obj Is ix Then       'Si encuentra
                mcolContactos.Remove i       'Elimina el objeto
                GoTo eliminado
            End If
        Next i
        'Si no encuentra el objeto, genera error
        Err.Raise ERR_INVALIDO, "PCProvCli_RemoveContacto", "No se puede eliminar el contacto."
    End If
eliminado:
    mbooModificado = True
    Set obj = Nothing
End Sub

'Duplica un Contacto para el mismo Proveedor/Cliente
Public Function CopiaContacto(ix As Long) As Long
    Dim nuevo As PCContacto, obj As PCContacto
    
    Set obj = mcolContactos.item(ix)
    Set nuevo = obj.Clone
    mcolContactos.Add nuevo
    
    CopiaContacto = mcolContactos.Count       'Devuelve el indice del nuevo elemento
    Set nuevo = Nothing
    Set obj = Nothing
    mbooModificado = True
End Function

Public Property Get CountContacto() As Long
    CountContacto = mcolContactos.Count
End Property

Public Property Get Contactos(ix As Long) As PCContacto
    Set Contactos = mcolContactos.item(ix)
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
    
'*** MAKOTO 26/mar/01 Se movió a Sub AsignarCuentaPre
'    'Cuenta contable predeterminado
'    If Me.BandCliente Then
'        Me.CodCuentaContable = objEmpresa.GNOpcion.CodCuentaCliPre
'    Else
'        Me.CodCuentaContable = objEmpresa.GNOpcion.CodCuentaProvPre
'    End If
'    mProps.CodCuentaContable2 = Me.CodCuentaContable
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Function CalculaSaldo(FechaHasta As Date) As Currency
    Dim sql As String, rs As Recordset
    
    'sql = "SELECT SUM(Debe)-SUM(Haber) AS Saldo " & _
          "FROM PCKardex INNER JOIN GNComprobante ON PCKardex.TransID=GNComprobante.TransID " & _
          "WHERE Estado<>" & ESTADO_ANULADO & _
          " AND PCKardex.FechaEmision<=" & FechaYMD(FechaHasta, mProps.objEmpresa.TipoDB)
          'PCKardex.FechaDoc  antes  revisar  porque Diego 08/2003
          
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not IsNull(rs!Saldo) Then
        CalculaSaldo = rs!Saldo
    Else
        CalculaSaldo = 0
    End If
End Function


Public Sub SetModificado()
    mbooModificado = True
End Sub


Public Property Get IdProvcli() As Long
    IdProvcli = mProps.IdProvcli
End Property


Public Property Let CodProvcli(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_CodProvCli", MSGERR_INVALIDO
    End If
'    ValidaCodigo value      'Valida si contiene caracteres no validos
    '18-02-04  para controlar que el código no contenga espacio en blanco
    mProps.CodProvcli = Trim(value)
    mbooModificado = True
End Property

Public Property Get CodProvcli() As String
    CodProvcli = Trim$(mProps.CodProvcli)
End Property

Public Property Let Nombre(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_Nombre", MSGERR_INVALIDO
    End If
    
'   If mbooEsNuevo Then
'        ValidaCodigo value      'Valida si contiene caracteres no validos
'    End If
    
    mProps.Nombre = value
    mbooModificado = True
End Property

Public Property Get Nombre() As String
    Nombre = Trim$(mProps.Nombre)
End Property

Public Property Let BandProveedor(ByVal value As Boolean)
    If value <> mProps.BandProveedor Then
        mProps.BandProveedor = value
        mbooModificado = True
        AsignarCodCuentaPre             '*** MAKOTO 26/mar/01
    End If
End Property

Public Property Get BandProveedor() As Boolean
    BandProveedor = mProps.BandProveedor
End Property

Public Property Let BandCliente(ByVal value As Boolean)
    If value <> mProps.BandCliente Then
        mProps.BandCliente = value
        mbooModificado = True
        AsignarCodCuentaPre             '*** MAKOTO 26/mar/01
    End If
End Property

Public Property Get BandCliente() As Boolean
    BandCliente = mProps.BandCliente
End Property

'*** MAKOTO 26/mar/01 Agregado
'Para que asigne bien automáticamente CodCuentaContable cuando crea el objeto
Private Sub AsignarCodCuentaPre()
    'Cuenta contable predeterminado
    If Len(Me.CodCuentaContable) = 0 And Not (Me.Empresa Is Nothing) Then
        If Me.BandCliente Then
            Me.CodCuentaContable = Me.Empresa.GNOpcion.CodCuentaCliPre
        ElseIf Me.BandGarante Then
            Me.CodCuentaContable = Me.Empresa.GNOpcion.CodCuentaCliPre
        Else
            Me.CodCuentaContable = Me.Empresa.GNOpcion.CodCuentaProvPre
        End If
        'jeaa 27/09/04 si en configuracion de cuenta prederminada 1 y 2
        'es 1 verdadero no pone el valor a las dos cuentas
        If (mProps.objEmpresa.GNOpcion.ObtenerValor("CuentaContPred1y2") = "1") Then
            If Len(Me.CodCuentaContable2) = 0 Then
                mProps.CodCuentaContable2 = Me.CodCuentaContable
                mProps.IdCuentaContable2 = mProps.IdCuentaContable
            End If
        End If
        
    End If
    If Len(Me.CodProvincia) = 0 And Not (Me.Empresa Is Nothing) Then
        Me.CodProvincia = mProps.objEmpresa.GNOpcion.CodProvinciaPre
    End If
    
    If Len(Me.CodCanton) = 0 And Not (Me.Empresa Is Nothing) Then
        Me.CodCanton = mProps.objEmpresa.GNOpcion.CodCantonPre
    End If

    
End Sub


Public Property Let IdCuentaContable(ByVal value As Long)
    mProps.IdCuentaContable = value
    mbooModificado = True
End Property

Public Property Get IdCuentaContable() As Long
    IdCuentaContable = mProps.IdCuentaContable
End Property


Public Property Let CodCuentaContable(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCuentaContable Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaContable también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli.CodCuentaContable", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaContable = rs!IdCuenta
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdCuentaContable = 0
        End If
    End If

    mProps.CodCuentaContable = value
    mbooModificado = True
End Property

Public Property Get CodCuentaContable() As String
    CodCuentaContable = Trim$(mProps.CodCuentaContable)
End Property

'*** MAKOTO 12/feb/01 Agregado
Public Property Let IdCuentaContable2(ByVal value As Long)
    mProps.IdCuentaContable2 = value
    mbooModificado = True
End Property

Public Property Get IdCuentaContable2() As Long
    IdCuentaContable2 = mProps.IdCuentaContable2
End Property

'*** MAKOTO 12/feb/01 Agregado
Public Property Let CodCuentaContable2(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCuentaContable2 Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaContable también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli.CodCuentaContable2", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaContable2 = rs!IdCuenta
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdCuentaContable2 = 0
        End If
    End If

    mProps.CodCuentaContable2 = value
    mbooModificado = True
End Property

Public Property Get CodCuentaContable2() As String
    CodCuentaContable2 = Trim$(mProps.CodCuentaContable2)
End Property
Public Property Let Direccion1(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_Direccion1", MSGERR_INVALIDO
    End If
    
    mProps.Direccion1 = value
    mbooModificado = True
End Property

Public Property Get Direccion1() As String
    Direccion1 = Trim$(mProps.Direccion1)
End Property

Public Property Let Direccion2(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_Direccion2", MSGERR_INVALIDO
    End If
    
    mProps.Direccion2 = value
    mbooModificado = True
End Property

Public Property Get Direccion2() As String
    Direccion2 = Trim$(mProps.Direccion2)
End Property

Public Property Let CodPostal(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_CodPostal", MSGERR_INVALIDO
    End If
    
    mProps.CodPostal = value
    mbooModificado = True
End Property

Public Property Get CodPostal() As String
    CodPostal = Trim$(mProps.CodPostal)
End Property

Public Property Let Ciudad(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_Ciudad", MSGERR_INVALIDO
    End If
    
    mProps.Ciudad = value
    mbooModificado = True
End Property

Public Property Get Ciudad() As String
    Ciudad = Trim$(mProps.Ciudad)
End Property

Public Property Let Provincia(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_Provincia", MSGERR_INVALIDO
    End If
    
    mProps.Provincia = value
    mbooModificado = True
End Property

Public Property Get Provincia() As String
    Provincia = Trim$(mProps.Provincia)
End Property

Public Property Let Pais(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_Pais", MSGERR_INVALIDO
    End If
    
    mProps.Pais = value
    mbooModificado = True
End Property

Public Property Get Pais() As String
    Pais = Trim$(mProps.Pais)
End Property



Public Property Let Telefono1(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_Telefono1", MSGERR_INVALIDO
    End If
    
    mProps.Telefono1 = value
    mbooModificado = True
End Property

Public Property Get Telefono1() As String
    Telefono1 = Trim$(mProps.Telefono1)
End Property

Public Property Let Telefono2(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_Telefono2", MSGERR_INVALIDO
    End If
    
    mProps.Telefono2 = value
    mbooModificado = True
End Property

Public Property Get Telefono2() As String
    Telefono2 = Trim$(mProps.Telefono2)
End Property

Public Property Let Telefono3(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_Telefono3", MSGERR_INVALIDO
    End If
    
    mProps.Telefono3 = value
    mbooModificado = True
End Property

Public Property Get Telefono3() As String
    Telefono3 = Trim$(mProps.Telefono3)
End Property



Public Property Let Fax(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_Fax", MSGERR_INVALIDO
    End If
    
    mProps.Fax = value
    mbooModificado = True
End Property

Public Property Get Fax() As String
    Fax = Trim$(mProps.Fax)
End Property

Public Property Let RUC(ByVal value As String)
    
    If Len(value) > 19 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_RUC", MSGERR_INVALIDO
    End If
    
    mProps.RUC = value
    mbooModificado = True
    
End Property
'*** Modificado Oliver
'para el uso de Anexos
Public Property Get RUC() As String
    RUC = mProps.RUC
End Property


Public Property Let EMail(ByVal value As String)
    If Len(value) > 250 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_EMail", MSGERR_INVALIDO
    End If
    
    mProps.EMail = value
    mbooModificado = True
End Property

Public Property Get EMail() As String
    EMail = Trim$(mProps.EMail)
End Property


Public Property Let LimiteCredito(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_LimiteCredito", MSGERR_INVALIDO
    End If
    
    mProps.LimiteCredito = value
    mbooModificado = True
End Property

Public Property Get LimiteCredito() As Currency
    LimiteCredito = mProps.LimiteCredito
End Property

Public Property Let banco(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_Banco", MSGERR_INVALIDO
    End If
    
    mProps.banco = value
    mbooModificado = True
End Property

Public Property Get banco() As String
    banco = Trim$(mProps.banco)
End Property

Public Property Let NumCuenta(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_NumCuenta", MSGERR_INVALIDO
    End If
    
    mProps.NumCuenta = value
    mbooModificado = True
End Property

Public Property Get NumCuenta() As String
    NumCuenta = Trim$(mProps.NumCuenta)
End Property

Public Property Let Swit(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_Swit", MSGERR_INVALIDO
    End If
    
    mProps.Swit = value
    mbooModificado = True
End Property

Public Property Get Swit() As String
    Swit = Trim$(mProps.Swit)
End Property

Public Property Let DirecBanco(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_DirecBanco", MSGERR_INVALIDO
    End If
    
    mProps.DirecBanco = value
    mbooModificado = True
End Property

Public Property Get DirecBanco() As String
    DirecBanco = Trim$(mProps.DirecBanco)
End Property

Public Property Let TelBanco(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_TelBanco", MSGERR_INVALIDO
    End If
    
    mProps.TelBanco = value
    mbooModificado = True
End Property

Public Property Get TelBanco() As String
    TelBanco = Trim$(mProps.TelBanco)
End Property

Public Property Let IdVendedor(ByVal value As Long)
    mProps.IdVendedor = value
    mbooModificado = True
End Property

Public Property Get IdVendedor() As Long
    IdVendedor = mProps.IdVendedor
End Property


Public Property Let CodVendedor(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodVendedor Then
        If Len(value) > 0 Then
            'Actualiza IdVendedor también
            sql = "SELECT IdVendedor FROM FCVendedor WHERE CodVendedor='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodVendedor", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdVendedor = rs!IdVendedor
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdVendedor = 0
        End If
    End If

    mProps.CodVendedor = value
    mbooModificado = True
End Property

Public Property Get CodVendedor() As String
    CodVendedor = Trim$(mProps.CodVendedor)
End Property


Public Property Let IdGrupo1(ByVal value As Long)
    mProps.IdGrupo1 = value
    mbooModificado = True
End Property

Public Property Get IdGrupo1() As Long
    IdGrupo1 = mProps.IdGrupo1
End Property


Public Property Let CodGrupo1(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodGrupo1 Then
        If Len(value) > 0 Then
            'Actualiza IdGrupo1 también
            sql = "SELECT IdGrupo1 FROM PCGrupo1 WHERE CodGrupo1='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodGrupo1", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdGrupo1 = rs!IdGrupo1
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdGrupo1 = 0
        End If
    End If

    mProps.CodGrupo1 = value
    mbooModificado = True
End Property

Public Property Get CodGrupo1() As String
    CodGrupo1 = Trim$(mProps.CodGrupo1)
End Property



Public Property Let IdGrupo2(ByVal value As Long)
    mProps.IdGrupo2 = value
    mbooModificado = True
End Property

Public Property Get IdGrupo2() As Long
    IdGrupo2 = mProps.IdGrupo2
End Property


Public Property Let CodGrupo2(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodGrupo2 Then
        If Len(value) > 0 Then
            'Actualiza IdGrupo2 también
            sql = "SELECT IdGrupo2 FROM PCGrupo2 WHERE CodGrupo2='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodGrupo2", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdGrupo2 = rs!IdGrupo2
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdGrupo2 = 0
        End If
    End If

    mProps.CodGrupo2 = value
    mbooModificado = True
End Property

Public Property Get CodGrupo2() As String
    CodGrupo2 = Trim$(mProps.CodGrupo2)
End Property


Public Property Let IdGrupo3(ByVal value As Long)
    mProps.IdGrupo3 = value
    mbooModificado = True
End Property
Public Property Let IdGrupo4(ByVal value As Long)
    mProps.IdGrupo4 = value
    mbooModificado = True
End Property

Public Property Get IdGrupo3() As Long
    IdGrupo3 = mProps.IdGrupo3
End Property


Public Property Let CodGrupo3(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodGrupo3 Then
        If Len(value) > 0 Then
            'Actualiza IdGrupo3 también
            sql = "SELECT IdGrupo3 FROM PCGrupo3 WHERE CodGrupo3='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodGrupo3", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdGrupo3 = rs!IdGrupo3
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdGrupo3 = 0
        End If
    End If

    mProps.CodGrupo3 = value
    mbooModificado = True
End Property
Public Property Get IdGrupo4() As Long
    IdGrupo4 = mProps.IdGrupo4
End Property

Public Property Let CodGrupo4(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodGrupo4 Then
        If Len(value) > 0 Then
            'Actualiza IdGrupo3 también
            sql = "SELECT IdGrupo4 FROM PCGrupo4 WHERE CodGrupo4='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodGrupo4", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdGrupo4 = rs!IdGrupo4
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdGrupo4 = 0
        End If
    End If
    mProps.CodGrupo4 = value
    mbooModificado = True
End Property

Public Property Get CodGrupo3() As String
    CodGrupo3 = Trim$(mProps.CodGrupo3)
End Property

Public Property Get CodGrupo4() As String
    CodGrupo4 = Trim$(mProps.CodGrupo4)
End Property

Public Property Let Estado(ByVal value As Byte)
    If value > 2 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_Estado", MSGERR_INVALIDO
    End If
    mProps.Estado = value
    mbooModificado = True
End Property
Public Property Get Estado() As Byte
    Estado = mProps.Estado
End Property

Public Property Get fechagrabado() As Date
    fechagrabado = mProps.fechagrabado
End Property
Public Property Get EsNuevo() As Boolean
    EsNuevo = mbooEsNuevo
End Property
Public Property Get Modificado() As Boolean
    Modificado = mbooModificado
End Property

Private Sub Class_Initialize()
    mProps.Estado = 0
    Set mcolContactos = New Collection
    mbooEsNuevo = True
    mbooModificado = False
    mProps.TipoProvCli = "NC"  'jeaa 17/12/2005
    mProps.Pais = "ECUADOR"
    mProps.CodPais = "593"
    BandRUCValido = False
    BandOmitirPendientes = False
    BandOmitirDinardap = False
    Set mcolAgencia = New Collection
'     mProps.RUC = String$(19, " ")          '20 caracteres en total, pero 1 es para tipo documento
End Sub

Private Sub Class_Terminate()
    Set mcolContactos = Nothing
    Set mProps.objEmpresa = Nothing
    Set mcolAgencia = Nothing
End Sub



Public Sub GenerarCodigo(ByVal lon As Integer)
'lon:Longitud de numero. ejm. 4 ==> P0001
    Dim sql As String, num As Long, rs As Recordset, pre As String, cod As String
    Dim lon_pre As Integer, sufijo As String, formato As String
    Dim i As Integer
    '-----
    ' Genera un código en el siguiente formato
    '   Cnnnn (Para clientes)
    '   Pnnnn (Para proveedores)
    '         (nnnn = Número secuencial)
    '-----
    
    'Prefijo de codigo
    If Me.BandCliente Then
        pre = Me.Empresa.GNOpcion.PrefijoCodCli
        formato = Me.Empresa.GNOpcion.FormatoCodCli
        'Si está definido que genere con sufijo, toma prefijo de C.Costo como sufijo
        If Me.Empresa.GNOpcion.GeneraCodCli = 2 Then sufijo = Me.Empresa.GNOpcion.PrefijoCodCC
    ElseIf Me.BandProveedor Then
        pre = Me.Empresa.GNOpcion.PrefijoCodProv
        formato = Me.Empresa.GNOpcion.FormatoCodProv
        'Si está definido que genere con sufijo, toma prefijo de C.Costo como sufijo
        If Me.Empresa.GNOpcion.GeneraCodProv = 2 Then sufijo = Me.Empresa.GNOpcion.PrefijoCodCC
    ElseIf Me.BandGarante Then
        pre = Me.Empresa.GNOpcion.PrefijoCodCli
        formato = Me.Empresa.GNOpcion.FormatoCodCli
        'Si está definido que genere con sufijo, toma prefijo de C.Costo como sufijo
        If Me.Empresa.GNOpcion.GeneraCodCli = 2 Then sufijo = Me.Empresa.GNOpcion.PrefijoCodCC
    End If
    lon_pre = Len(pre)
    
#If DAOLIB Then
    sql = "SELECT Max(Val(Right(abs(CodProvCli),Len(CodProvCli)-" & lon_pre & "))) AS CodMax " & _
          "FROM PCProvCli WHERE CodProvCli Like '" & pre & "*'"
#Else
    sql = "SELECT Max(Right((CodProvCli),Len(CodProvCli)-" & lon_pre & ")) AS CodMax "
    sql = sql & "FROM PCProvCli WHERE CodProvCli Like '" & pre
    For i = 1 To lon_pre
        sql = sql & "[0-9][0-9]" 'AUC 08/06/07 CAMBIADO ANTES "[0-9][0-9]
    Next i
    sql = sql & "%'"
          '"FROM PCProvCli WHERE CodProvCli Like '" & pre & "[0-9][0-9]%'"
    If Me.BandCliente Or Me.BandGarante Then
        sql = sql & " and bandcliente=1"
    Else
        sql = sql & " and bandproveedor=1"
    End If
          
#End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        If Not IsNull(rs!CodMax) Then num = val(rs!CodMax)
    End If
    
    cod = pre & Format(num + 1, formato) & sufijo
    Me.CodProvcli = cod
    rs.Close
    Set rs = Nothing
End Sub

'Agregado Alex: Sept/2002
' Para el control del Aspecto de Anexos
Public Property Get TipoDocumento() As String
    TipoDocumento = mProps.TipoDocumento
End Property

Public Property Let TipoDocumento(ByVal value As String)
    If Me.Empresa.GNOpcion.ObtenerValor("PermiteControlAspectosAnexos") = 1 Then '***Agregado. 10/sep/2003. Angel
        If (Len(value) > 2) Or (Not IsNumeric(value)) Then
            Err.Raise ERR_INVALIDO, "PCProvCli.TipoDocumento", MSGERR_INVALIDO
        End If
    End If
    mProps.TipoDocumento = value
    mbooModificado = True
End Property

Public Property Get TipoComprobante() As String
    TipoComprobante = mProps.TipoComprobante
End Property

Public Property Let TipoComprobante(ByVal value As String)
    If Me.Empresa.GNOpcion.ObtenerValor("PermiteControlAspectosAnexos") = 1 Then '***Agregado. 10/sep/2003. Angel
        If (Len(value) > 5) Then
            Err.Raise ERR_INVALIDO, "PCProvCli.TipoComprobante", MSGERR_INVALIDO
        End If
    End If
    mProps.TipoComprobante = value
    mbooModificado = True
End Property

Public Property Get NumAutSRI() As String
    NumAutSRI = mProps.NumAutSRI
End Property

Public Property Let NumAutSRI(ByVal value As String)
    If Me.Empresa.GNOpcion.ObtenerValor("PermiteControlAspectosAnexos") = 1 Then '***Agregado. 10/sep/2003. Angel
        If (Len(value) > 50) Or (Not IsNumeric(value)) Then
            Err.Raise ERR_INVALIDO, "PCProvCli.NumAutSRI", MSGERR_INVALIDO
        End If
    End If
    mProps.NumAutSRI = value
    mbooModificado = True
End Property

'Agregado Alex: Sept/2002
Public Function VerificaRUC(ByVal RUC As String) As Boolean
    'Solo verifica  si el tipo de documento es Cedula  o RUC
'    If Me.TipoDocumento <> "3" Then
    If Me.codtipoDocumento <> "O" And Me.codtipoDocumento <> "T" Then
        '
        If Me.TipoProvCli <> "CE" Then
            VerificaRUC = VerificaDocumentoNew(RUC) '  TRECER DIGITO = 6
        Else
            VerificaRUC = VerificaDocumento(RUC)
        End If
    Else
        VerificaRUC = True
    End If
End Function

 
'*** Agregado Oliver 30/jun/2003 para verificar control de duplicidad de RUC o CI
'Verifica si se repite el RUC
Public Function VerificaRucRepetido() As Boolean
    Dim rs As Recordset, sql As String
    VerificaRucRepetido = False
    sql = "SELECT  RUC FROM PCProvCli "
    sql = sql & " WHERE IDPROVCLI <> " & Me.IdProvcli & " AND (RUC='" & mProps.RUC & "') AND "
    sql = sql & " ((BandCliente=" & IIf(mProps.BandCliente, 1, 0) & ") OR "
    sql = sql & " (BandProveedor=" & IIf(mProps.BandProveedor, 1, 0) & ")) AND "
    sql = sql & " (IdProvCli<>" & mProps.IdProvcli & ")"
    If (mProps.objEmpresa.GNOpcion.ObtenerValor("DuplicaRUC")) = "0" Then
        sql = sql & " AND (IdGrupo1 = " & mProps.IdGrupo1 & " ) "
        sql = sql & " AND (IdGrupo4 = " & mProps.IdGrupo4 & " ) "
    End If
    
    
    If (mProps.objEmpresa.GNOpcion.ObtenerValor("DuplicaRUCEstab")) = "1" Then

        sql = sql & " AND (numserie = '" & mProps.NumSerie & "' ) "

    End If
    
    sql = sql & " AND BandEsConsumidorFinal <> 1"
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
            
                VerificaRucRepetido = True
'            If (mProps.objEmpresa.GNOpcion.ObtenerValor("DuplicaRUCEstab")) = "0" Then
'
'                mProps.CodProvCli = mProps.RUC + mProps.NumSerie
'                VerificaRucRepetido = False
'            End If
    End If
    rs.Close
End Function

'***Agregado. 05/sep/2003. Angel
Public Property Let NombreAlterno(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_NombreAlterno", MSGERR_INVALIDO
    End If
    
    mProps.NombreAlterno = value
    
    mbooModificado = True
End Property

Public Property Get NombreAlterno() As String
    NombreAlterno = Trim$(mProps.NombreAlterno)
End Property

'***Agregado. 05/sep/2003. Angel
Public Property Let FechaNacimiento(value As Date)
    If Not (IsDate(value)) Then
        Err.Raise ERR_INVALIDO, "PCProvCli_FechaNacimiento", MSGERR_INVALIDO
    End If
    
    mProps.FechaNacimiento = value
    mbooModificado = True
End Property

Public Property Get FechaNacimiento() As Date
    FechaNacimiento = mProps.FechaNacimiento
End Property

'***Agregado. 05/sep/2003. Angel
Public Property Let FechaEntrega(value As Date)
    If Not (IsDate(value)) Then
        Err.Raise ERR_INVALIDO, "PCProvCli_FechaEntrega", MSGERR_INVALIDO
    End If
    
    mProps.FechaEntrega = value
    mbooModificado = True
End Property

Public Property Get FechaEntrega() As Date
    FechaEntrega = mProps.FechaEntrega
End Property

'***Agregado. 05/sep/2003. Angel
Public Property Let FechaExpiracion(value As Date)
    If Not (IsDate(value)) Then
        Err.Raise ERR_INVALIDO, "PCProvCli_FechaExpiracion", MSGERR_INVALIDO
    End If
    
    mProps.FechaExpiracion = value
    mbooModificado = True
End Property

Public Property Get FechaExpiracion() As Date
    FechaExpiracion = mProps.FechaExpiracion
End Property

'***Agregado. 08/sep/2003. Angel
Public Property Let TotalDebe(value As Currency)
mProps.TotalDebe = value
    mbooModificado = True
End Property

'***Agregado. 08/sep/2003. Angel
Public Property Get TotalDebe() As Currency
    TotalDebe = mProps.TotalDebe
End Property

'***Agregado. 08/sep/2003. Angel
Public Property Let TotalHaber(value As Currency)
    mProps.TotalHaber = value
    mbooModificado = True
End Property

'***Agregado. 08/sep/2003. Angel
Public Property Get TotalHaber() As Currency
    TotalHaber = mProps.TotalHaber
End Property
'Agregado AUC 31/08/2005
Public Property Let Observacion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_Observacion", MSGERR_INVALIDO
    End If
    
    mProps.Observacion = value
    mbooModificado = True
End Property
'Agregado AUC 31/08/2005
Public Property Get Observacion() As String
    Observacion = Trim$(mProps.Observacion)
End Property


'jeaa 17/12/2005
Public Property Let TipoProvCli(ByVal value As String)
    mProps.TipoProvCli = value
    mbooModificado = True
End Property

Public Property Get TipoProvCli() As String
    TipoProvCli = mProps.TipoProvCli
End Property


'Agregado jeaa 08/12/2006
Public Property Let ReferidoPor(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_ReferidoPor", MSGERR_INVALIDO
    End If
    
    mProps.ReferidoPor = value
    mbooModificado = True
End Property
'Agregado AUC 31/08/2005
Public Property Get ReferidoPor() As String
    ReferidoPor = Trim$(mProps.ReferidoPor)
End Property


Public Property Get IdTipoDocumento() As Long
    IdTipoDocumento = mProps.IdTipoDocumento
End Property

Public Property Let IdTipoDocumento(ByVal value As Long)
    mProps.IdTipoDocumento = value
    mbooModificado = True
End Property


Public Property Let codtipoDocumento(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.codtipoDocumento Then
        If Len(value) > 0 Then
            'Actualiza IdTipoDocumento también
            sql = "SELECT Id FROM Anexo_TipoDocumento WHERE CodTipoDocumento ='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodTipoDocumento", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdTipoDocumento = rs!Id
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdTipoDocumento = 0
        End If
    End If

    mProps.codtipoDocumento = value
    
    
    mbooModificado = True
End Property

Public Property Get codtipoDocumento() As String
    codtipoDocumento = Trim$(mProps.codtipoDocumento)
End Property

Public Property Let BandEmpresaPublica(ByVal value As Boolean)
    If value <> mProps.BandEmpresaPublica Then
        mProps.BandEmpresaPublica = value
        mbooModificado = True
        AsignarCodCuentaPre             '*** MAKOTO 26/mar/01
    End If
End Property

Public Property Get BandEmpresaPublica() As Boolean
    BandEmpresaPublica = mProps.BandEmpresaPublica
End Property

Public Property Let IdGarante(ByVal value As Long)
    mProps.IdGarante = value
    mbooModificado = True
End Property

Public Property Get IdGarante() As Long
    IdGarante = mProps.IdGarante
End Property


Public Property Let CodGarante(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodGarante Then
        If Len(value) > 0 Then
            'Actualiza IdGarante también
            sql = "SELECT IdProvCli FROM PCProvCli WHERE CodProvCli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodGarante", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdGarante = rs!IdProvcli
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdGarante = 0
        End If
    End If

    mProps.CodGarante = value
    mbooModificado = True
End Property

Public Property Get CodGarante() As String
    CodGarante = Trim$(mProps.CodGarante)
End Property

Public Property Let BandGarante(ByVal value As Boolean)
    If value <> mProps.BandGarante Then
        mProps.BandGarante = value
        mbooModificado = True
        AsignarCodCuentaPre             '*** MAKOTO 26/mar/01
    End If
End Property

Public Property Get BandGarante() As Boolean
    BandGarante = mProps.BandGarante
End Property

Public Sub GrabarEnSucursal(ByVal bd As String)
End Sub

Public Sub GrabarEnSucursales(ByVal bd As String, ByVal servidor As String)
   Dim sql As String, fgrabado As Date, i As Long
    Dim rs As Recordset
    If mbooEsNuevo Then
        'Verifica si no repite el código
        sql = "SELECT CodProvCli FROM " & bd & ".dbo.PCProvCli " & _
              "WHERE (CodProvCli='" & mProps.CodProvcli & "') AND " & _
                    "(IdProvCli<>" & mProps.IdProvcli & ")"
        Set rs = mProps.objEmpresa.OpenRecordsetSUC(sql, bd, servidor)
       If Not rs.EOF Then      'Si ya existe el codigo
            Err.Raise ERR_REPITECODIGO, "PCProvCli", MSGERR_REPITECODIGO
        End If
        rs.Close
    End If
    On Error GoTo ErrTrap
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM " & bd & ".dbo.PCProvCli WHERE 1=0"
                Set rs = mProps.objEmpresa.OpenRecordsetSUC(sql, bd, servidor)
       rs.AddNew
    'Si es modificada
    Else
        'Elimina todos los Contactos
'        sql = "DELETE FROM " & bd & ".dbo.PCContacto WHERE IdProvCli=" & mProps.IdProvCli
'        mProps.objEmpresa.Execute sql, True
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM " & bd & ".dbo.PCProvCli WHERE CodProvCli='" & mProps.CodProvcli & "'"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEditSuc(sql, bd, servidor)
       If rs.RecordCount = "0" Then
            sql = "SELECT * FROM " & bd & ".dbo.PCProvCli WHERE 1=0"
            Set rs = mProps.objEmpresa.OpenRecordsetParaEditSuc(sql, bd, servidor)
           rs.AddNew
        End If
#If DAOLIB Then
        rs.Edit
#End If
    End If
    'Graba el registro en la base
    With rs
        !CodProvcli = mProps.CodProvcli
        !Nombre = mProps.Nombre
        !BandProveedor = mProps.BandProveedor
        !BandCliente = mProps.BandCliente
''''        !IdCuentaContable = mProps.IdCuentaContable
''''        !IdCuentaContable2 = mProps.IdCuentaContable2
        !Direccion1 = mProps.Direccion1
        !Direccion2 = mProps.Direccion2
        !CodPostal = mProps.CodPostal
        !Ciudad = mProps.Ciudad
        !Provincia = mProps.Provincia
        !Pais = mProps.Pais
        !Telefono1 = mProps.Telefono1
        !Telefono2 = mProps.Telefono2
        !Telefono3 = mProps.Telefono3
        !Fax = mProps.Fax
        !RUC = mProps.RUC
        !EMail = mProps.EMail
        If Not !LimiteCredito = mProps.LimiteCredito Then
            Me.Empresa.GrabaAccionModuloEnSucursales bd, servidor, !LimiteCredito, mProps.LimiteCredito, "LimiteCredito", "MOD_CLI", "CL"
       End If
        !LimiteCredito = mProps.LimiteCredito
        !banco = mProps.banco
        !NumCuenta = mProps.NumCuenta
        !Swit = mProps.Swit
        !DirecBanco = mProps.DirecBanco
        !TelBanco = mProps.TelBanco
        !IdVendedor = mProps.IdVendedor
        !IdGrupo1 = mProps.IdGrupo1
        !IdGrupo2 = mProps.IdGrupo2
        !IdGrupo3 = mProps.IdGrupo3
        !IdGrupo4 = mProps.IdGrupo4 'agregado AUC 22/09/2005
        !Estado = mProps.Estado
        !TipoDocumento = mProps.TipoDocumento           'Agregado Alex: Sept/2002
        !TipoComprobante = mProps.TipoComprobante       'Agregado Alex: Sept/2002
        !NumAutSRI = mProps.NumAutSRI                   'Agregado Alex: Sept/2002
        !NombreAlterno = mProps.NombreAlterno           '***Agregado. 05/sep/2003. Angel
        !FechaNacimiento = mProps.FechaNacimiento       '***Agregado. 05/sep/2003. Angel
        !FechaEntrega = mProps.FechaEntrega             '***Agregado. 05/sep/2003. Angel
        !FechaExpiracion = mProps.FechaExpiracion       '***Agregado. 05/sep/2003. Angel
        !TotalDebe = mProps.TotalDebe                   '***Agregado. 08/sep/2003. Angel
        !TotalHaber = mProps.TotalHaber                 '***Agregado. 08/sep/2003. Angel
        !Observacion = mProps.Observacion  'Agregado AUC 31/08/2005
        !TipoProvCli = mProps.TipoProvCli       ' jeaa 17/12/2005
        !ReferidoPor = mProps.ReferidoPor  'Agregado jeaa 12/08/2006
        !IdTipoDocumento = mProps.IdTipoDocumento   'Agregado jeaa 15/05/2007
        !BandEmpresaPublica = mProps.BandEmpresaPublica   'jeaa 16/01/2008 cambio por EQUIDAD TRIBUTARIA
        !IdGarante = mProps.IdGarante
        !BandGarante = mProps.BandGarante
        !IdProvincia = mProps.IdProvincia
        !IdCanton = mProps.IdCanton
        !idparrroquia = mProps.IDParroquia
        !FechaCreacion = mProps.FechaCreacion
        !FechaLote1 = mProps.FechaLote1
        !FechaLote2 = mProps.FechaLote2
        !BandLote = mProps.BandLote
        !bandEmpleado = mProps.bandEmpleado
        !NumPunto = mProps.NumPunto
        !NumSerie = mProps.NumSerie
        !FechaCaducidad = mProps.FechaCaducidad
        !IdDiasCredito = mProps.IdDiasCredito
        !BandRUCValido = mProps.BandRUCValido
        !BandRelacionado = mProps.BandRelacionado
        !BandOmitirPendientes = mProps.BandOmitirPendientes
        !TipoSujeto = mProps.TipoSujeto
        !Sexo = mProps.Sexo
        !EstadoCivil = mProps.EstadoCivil
        !OrigenIngresos = mProps.OrigenIngresos
        !BandOmitirDinardap = BandOmitirDinardap
        !IdGarante2 = mProps.IdGarante2
        !NumCasa = mProps.NumCasa
        !bandCompensaIVA = mProps.bandCompensaIVA
        !BandEsConsumidorFinal = mProps.BandEsConsumidorFinal
        !BandIVCodigoIgual = mProps.BandIVCodigoIgual
        
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        fgrabado = !fechagrabado
''        mProps.IdProvCli = !IdProvCli
        .Close
    End With
    'Graba todos los PCContacto
''''    For i = 1 To CountContacto
''''        Contactos(i).Grabar
''''    Next i
    mProps.fechagrabado = fgrabado
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccionEnSucursales bd, servidor, "NUE_PC", "Proveedor/Cliente Nuevo: " & _
                                                         mProps.CodProvcli, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccionEnSucursales bd, servidor, "MOD_PC", "Proveedor/Cliente Modificado: " & _
                                                         mProps.CodProvcli, "DT"
    End If





    Set rs = Nothing
    Exit Sub
ErrTrap:
#If DAOLIB Then
    gobjWS.Rollback                   '======== Anula la transacción
#Else
    mProps.objEmpresa.Coneccion.RollbackTrans  '======== Anula la transacción
#End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

'AUC agregado diasplazo
Public Property Let DiasPlazo(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_DiasPlazo", MSGERR_INVALIDO
    End If
    mProps.DiasPlazo = value
    mbooModificado = True
End Property

Public Property Get DiasPlazo() As Currency
    DiasPlazo = mProps.DiasPlazo
End Property


Public Property Let CodigoUsuario(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_CodigoUsuario", MSGERR_INVALIDO
    End If
    
    mProps.CodigoUsuario = value
    mbooModificado = True
End Property

Public Property Get CodigoUsuario() As String
    CodigoUsuario = Trim$(mProps.CodigoUsuario)
End Property



Friend Sub RecuperarxUsusario(key As String)
    Dim sql As String, i As Long
    Dim rs As Recordset
    
    sql = "SELECT * FROM vwPCProvCli WHERE CodigoUsuario='" & key & "'"
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IdProvcli = !IdProvcli
        mProps.CodProvcli = !CodProvcli
        If Not IsNull(!Nombre) Then mProps.Nombre = !Nombre
        If Not IsNull(!BandProveedor) Then mProps.BandProveedor = !BandProveedor
        If Not IsNull(!BandCliente) Then mProps.BandCliente = !BandCliente
        If Not IsNull(!IdCuentaContable) Then mProps.IdCuentaContable = !IdCuentaContable
        If Not IsNull(!IdCuentaContable2) Then mProps.IdCuentaContable2 = !IdCuentaContable2
        If Not IsNull(!CodCuentaContable) Then mProps.CodCuentaContable = !CodCuentaContable
        If Not IsNull(!CodCuentaContable2) Then mProps.CodCuentaContable2 = !CodCuentaContable2
        If Not IsNull(!Direccion1) Then mProps.Direccion1 = !Direccion1
        If Not IsNull(!Direccion2) Then mProps.Direccion2 = !Direccion2
        If Not IsNull(!CodPostal) Then mProps.CodPostal = !CodPostal
        If Not IsNull(!Ciudad) Then mProps.Ciudad = !Ciudad
        If Not IsNull(!Provincia) Then mProps.Provincia = !Provincia
        If Not IsNull(!Pais) Then mProps.Pais = !Pais
        If Not IsNull(!Telefono1) Then mProps.Telefono1 = !Telefono1
        If Not IsNull(!Telefono2) Then mProps.Telefono2 = !Telefono2
        If Not IsNull(!Telefono3) Then mProps.Telefono3 = !Telefono3
        If Not IsNull(!Fax) Then mProps.Fax = !Fax
        If Not IsNull(!RUC) Then mProps.RUC = !RUC
        If Not IsNull(!EMail) Then mProps.EMail = !EMail
        If Not IsNull(!LimiteCredito) Then mProps.LimiteCredito = !LimiteCredito
        If Not IsNull(!banco) Then mProps.banco = !banco
        If Not IsNull(!NumCuenta) Then mProps.NumCuenta = !NumCuenta
        If Not IsNull(!Swit) Then mProps.Swit = !Swit
        If Not IsNull(!DirecBanco) Then mProps.DirecBanco = !DirecBanco
        If Not IsNull(!TelBanco) Then mProps.TelBanco = !TelBanco
        If Not IsNull(!IdVendedor) Then mProps.IdVendedor = !IdVendedor
        If Not IsNull(!CodVendedor) Then mProps.CodVendedor = !CodVendedor
        If Not IsNull(!IdGrupo1) Then mProps.IdGrupo1 = !IdGrupo1
        If Not IsNull(!CodGrupo1) Then mProps.CodGrupo1 = !CodGrupo1
        If Not IsNull(!IdGrupo2) Then mProps.IdGrupo2 = !IdGrupo2
        If Not IsNull(!CodGrupo2) Then mProps.CodGrupo2 = !CodGrupo2
        
        If Not IsNull(!IdGrupo3) Then mProps.IdGrupo3 = !IdGrupo3
        If Not IsNull(!CodGrupo3) Then mProps.CodGrupo3 = !CodGrupo3
        
        If Not IsNull(!IdGrupo4) Then mProps.IdGrupo4 = !IdGrupo4 'Agregadoa AUC 22/09/2005
        If Not IsNull(!CodGrupo4) Then mProps.CodGrupo4 = !CodGrupo4
        
        If Not IsNull(!Estado) Then mProps.Estado = !Estado
        If Not IsNull(!TipoDocumento) Then mProps.TipoDocumento = !TipoDocumento          'Agregado Alex Sept/2002
        If Not IsNull(!TipoComprobante) Then mProps.TipoComprobante = !TipoComprobante           'Agregado Alex Sept/2002
        If Not IsNull(!NumAutSRI) Then mProps.NumAutSRI = !NumAutSRI             'Agregado Alex Sept/2002
        If Not IsNull(!NombreAlterno) Then mProps.NombreAlterno = !NombreAlterno          '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!FechaNacimiento) Then mProps.FechaNacimiento = !FechaNacimiento    '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!FechaEntrega) Then mProps.FechaEntrega = !FechaEntrega             '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!FechaExpiracion) Then mProps.FechaExpiracion = !FechaExpiracion    '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!TotalDebe) Then mProps.TotalDebe = !TotalDebe                      '***Agregado. 08/sep/2003. Angel
        If Not IsNull(!TotalHaber) Then mProps.TotalHaber = !TotalHaber                   '***Agregado. 08/sep/2003. Angel
        If Not IsNull(!Observacion) Then mProps.Observacion = !Observacion      'Agregado AUC 31/08/2005
        If Not IsNull(!TipoProvCli) Then mProps.TipoProvCli = Trim$(!TipoProvCli)  ' jeaa 17/12/2005
        If Not IsNull(!ReferidoPor) Then mProps.ReferidoPor = !ReferidoPor      'Agregado jeaa 12/08/2006
        'jeaa 15/05/2007
        If Not IsNull(!IdTipoDocumento) Then mProps.IdTipoDocumento = !IdTipoDocumento 'Agregado jeaa 15/05/2007
        If Not IsNull(!codtipoDocumento) Then mProps.codtipoDocumento = !codtipoDocumento 'Agregado jeaa 15/05/2007
        If Not IsNull(!BandEmpresaPublica) Then mProps.BandEmpresaPublica = !BandEmpresaPublica   'jeaa 16/01/2008 cambio por EQUIDAD TRIBUTARIA
        If Not IsNull(!IdGarante) Then mProps.IdGarante = !IdGarante
        If Not IsNull(!CodGarante) Then mProps.CodGarante = !CodGarante
        If Not IsNull(!BandGarante) Then mProps.BandGarante = !BandGarante
        If Not IsNull(!DiasPlazo) Then mProps.DiasPlazo = !DiasPlazo
        If Not IsNull(!IdProvincia) Then mProps.IdProvincia = !IdProvincia
        If Not IsNull(!IdCanton) Then mProps.IdCanton = !IdCanton
        If Not IsNull(!IDParroquia) Then mProps.IDParroquia = !IDParroquia
        If Not IsNull(!FechaCreacion) Then mProps.FechaCreacion = !FechaCreacion
        If Not IsNull(!CodigoUsuario) Then mProps.CodigoUsuario = !CodigoUsuario
        If Not IsNull(!FechaLote1) Then mProps.CodigoUsuario = !FechaLote1
        If Not IsNull(!FechaLote2) Then mProps.CodigoUsuario = !FechaLote2
        If Not IsNull(!BandLote) Then mProps.CodigoUsuario = !BandLote
        If Not IsNull(!bandEmpleado) Then mProps.bandEmpleado = !bandEmpleado
        If Not IsNull(!NumSerie) Then mProps.NumSerie = !NumSerie
        If Not IsNull(!NumPunto) Then mProps.NumPunto = !NumPunto
        If Not IsNull(!FechaCaducidad) Then mProps.FechaCaducidad = !FechaCaducidad
        If Not IsNull(!IdDiasCredito) Then mProps.IdDiasCredito = !IdDiasCredito
        If Not IsNull(!BandRUCValido) Then mProps.BandRUCValido = !BandRUCValido
        If Not IsNull(!BandRelacionado) Then mProps.BandRelacionado = !BandRelacionado
        If Not IsNull(!BandOmitirPendientes) Then mProps.BandOmitirPendientes = !BandOmitirPendientes
        If Not IsNull(!TipoSujeto) Then mProps.TipoSujeto = !TipoSujeto
        If Not IsNull(!Sexo) Then mProps.Sexo = !Sexo
        If Not IsNull(!EstadoCivil) Then mProps.EstadoCivil = !EstadoCivil
        If Not IsNull(!OrigenIngresos) Then mProps.OrigenIngresos = !OrigenIngresos
        If Not IsNull(!BandOmitirDinardap) Then mProps.BandOmitirDinardap = !BandOmitirDinardap
        If Not IsNull(!IdGarante2) Then mProps.IdGarante2 = !IdGarante2
        If Not IsNull(!CodGarante2) Then mProps.CodGarante2 = !CodGarante2
        If Not IsNull(!NumCasa) Then mProps.NumCasa = !NumCasa
        If Not IsNull(!bandCompensaIVA) Then mProps.bandCompensaIVA = !bandCompensaIVA
        If Not IsNull(!BandEsConsumidorFinal) Then mProps.BandEsConsumidorFinal = !BandEsConsumidorFinal
        If Not IsNull(!BandIVCodigoIgual) Then mProps.BandIVCodigoIgual = !BandIVCodigoIgual
        If Not IsNull(!IdCuentaContableGasto) Then mProps.IdCuentaContableGasto = !IdCuentaContableGasto
        If Not IsNull(!CodCuentaContableGasto) Then mProps.CodCuentaContableGasto = !CodCuentaContableGasto
        
        
        
        
        
        
        mProps.fechagrabado = !fechagrabado
        
        .Close
    End With
    
        
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Property Let IdProvincia(ByVal value As Long)
    mProps.IdProvincia = value
    mbooModificado = True
End Property

Public Property Get IdProvincia() As Long
    IdProvincia = mProps.IdProvincia
End Property

Public Property Let IdCanton(ByVal value As Long)
    mProps.IdCanton = value
    mbooModificado = True
End Property

Public Property Get IdCanton() As Long
    IdCanton = mProps.IdCanton
End Property

Public Property Let IDParroquia(ByVal value As Long)
    mProps.IDParroquia = value
    mbooModificado = True
End Property

Public Property Get IDParroquia() As Long
    IDParroquia = mProps.IDParroquia
End Property

Public Property Let CodProvincia(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodProvincia Then
        If Len(value) > 0 Then
            'Actualiza IdGarante también
            sql = "SELECT IdProvincia FROM PCProvincia WHERE CodProvincia='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodProvincia", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdProvincia = rs!IdProvincia
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdProvincia = 0
        End If
    End If
    mProps.CodProvincia = value
    mbooModificado = True
End Property

Public Property Get CodProvincia() As String
    CodProvincia = Trim$(mProps.CodProvincia)
End Property

Public Property Let CodCanton(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodCanton Then
        If Len(value) > 0 Then
            'Actualiza IdGarante también
            sql = "SELECT Idcanton FROM PCcanton WHERE Codcanton='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_Codcanton", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCanton = rs!IdCanton
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdCanton = 0
        End If
    End If
    mProps.CodCanton = value
    mbooModificado = True
End Property

Public Property Get CodCanton() As String
    CodCanton = Trim$(mProps.CodCanton)
End Property

Public Property Let CodParroquia(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
'    If value <> mProps.CodParroqui Then
        If Len(value) > 0 Then
            'Actualiza IdGarante también
            sql = "SELECT IdParroquia FROM PCParroquia pcp inner join pccanton pcc on pcp.idcanton = pcc.idcanton WHERE codcanton = '" & mProps.CodCanton & "' and CodParroquia='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodParroquia", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IDParroquia = rs!IDParroquia
            End If
            rs.Close
            Set rs = Nothing
        Else
            IDParroquia = 0
        End If
'    End If
    mProps.CodParroqui = value
    mbooModificado = True
End Property

Public Property Get CodParroquia() As String
    CodParroquia = Trim$(mProps.CodParroqui)
End Property

Public Property Let FechaCreacion(value As Date)
    If Not (IsDate(value)) Then
        Err.Raise ERR_INVALIDO, "PCProvCli_Creacion", MSGERR_INVALIDO
    End If
    mProps.FechaCreacion = value
    mbooModificado = True
End Property

Public Property Get FechaCreacion() As Date
    FechaCreacion = mProps.FechaCreacion
End Property

Public Property Let FechaLote1(value As Date)
    If Not (IsDate(value)) Then
        Err.Raise ERR_INVALIDO, "PCProvCli_Lote1", MSGERR_INVALIDO
    End If
    mProps.FechaLote1 = value
    mbooModificado = True
End Property

Public Property Get FechaLote1() As Date
    FechaLote1 = mProps.FechaLote1
End Property

Public Property Let FechaLote2(value As Date)
    If Not (IsDate(value)) Then
        Err.Raise ERR_INVALIDO, "PCProvCli_Lote2", MSGERR_INVALIDO
    End If
    mProps.FechaLote2 = value
    mbooModificado = True
End Property

Public Property Get FechaLote2() As Date
    FechaLote2 = mProps.FechaLote2
End Property

Public Property Let BandLote(ByVal value As Boolean)
    If value <> mProps.BandLote Then
        mProps.BandLote = value
        mbooModificado = True
    End If
End Property

Public Property Get BandLote() As Boolean
    BandLote = mProps.BandLote
End Property

Public Function AddDetalleItem() As Long
    Dim nuevo As IVDetalleItem
    If mcolDetalleItem Is Nothing Then CargaDetalleItem
    Set nuevo = New IVDetalleItem
    Set nuevo.PCProvCli = Me
    nuevo.Cantidad = 1
    mcolDetalleItem.Add nuevo
    AddDetalleItem = mcolDetalleItem.Count
    Set nuevo = Nothing
    mbooModificado = True
End Function

Public Function RecuperaDetalleItem(ix As Long) As IVDetalleItem
    Set RecuperaDetalleItem = mcolDetalleItem.item(ix)
End Function

Public Property Get NumDetalleItem() As Long
    If mcolDetalleItem Is Nothing Then
        NumDetalleItem = 0
    Else
        NumDetalleItem = mcolDetalleItem.Count
    End If
End Property

Private Sub CargaDetalleItem()
    Dim rs As Recordset, sql As String, obj As IVDetalleItem
    Dim i As Long
    Set mcolDetalleItem = New Collection
       sql = "Select piv.idprovcli, IV.IdInventario, Iv.CodInventario, Iv.Descripcion,piv.cantidad, " & _
                " piv.PU,piv.frecuencia, piv.FechaIni, piv.FechaFin, piv.BandProntoPago, piv.Referencia,piv.bandPublicidad, piv.Contrato, piv.plazo  " & _
              "  From VWIvinventarioRECUPERAR Iv " & _
            "INNER JOIN PCProvCliInv PIv "
        sql = sql & " ON PIv.IdInventario = IV.Idinventario"
        sql = sql & " where piv.idprovcli = " & Me.IdProvcli
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddDetalleItem
            Set obj = mcolDetalleItem.item(i)
                  obj.IdProvcli = !IdProvcli
                  obj.IdInventario = !IdInventario
                  obj.CodInventario = !CodInventario
                  obj.Descripcion = !Descripcion
                  obj.Cantidad = !Cantidad
                  obj.PU = !PU
                  If Not IsNull(!Frecuencia) Then obj.Frecuencia = !Frecuencia
                  If Not IsNull(!fechaIni) Then obj.fechaIni = !fechaIni
                  If Not IsNull(!FechaFin) Then obj.FechaFin = !FechaFin
                  If Not IsNull(!BandProntoPago) Then obj.BandProntoPago = !BandProntoPago
                  If Not IsNull(!Referencia) Then obj.Referencia = !Referencia
                  If Not IsNull(!BandPublicidad) Then obj.BandPublicidad = !BandPublicidad
                  If Not IsNull(!Contrato) Then obj.Contrato = !Contrato
                  If Not IsNull(!Plazo) Then obj.Plazo = !Plazo
           .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
        Set obj = Nothing
End Sub

Public Sub GrabarIVDetalle()
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As IVDetalleItem, sql As String, band As Boolean
    Dim NumHijos As Integer, CountNumHijos As Integer
    On Error GoTo CapturaError
    'Elimina  todos  los registros
    sql = "Delete pciv " & _
          "From  PCProvCliInv pciv " & _
          "INNER JOIN IVInventario IV  ON iv.Idinventario = pciv.Idinventario"
          sql = sql & " Where pciv.Idprovcli = " & mProps.IdProvcli
    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    band = True
    If Not mcolDetalleItem Is Nothing Then
        sql = "SELECT * FROM PCProvCliInv WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        NumHijos = mcolDetalleItem.Count
        CountNumHijos = 0
        For i = 1 To mcolDetalleItem.Count
            If CountNumHijos < NumHijos Then
                CountNumHijos = CountNumHijos + 1
                Set obj = mcolDetalleItem.item(i)
                If obj.IdProvcli = 0 Then
                    obj.Recuperar obj.CodInventario, ""
                    If verificacionDuplicado(obj.IdInventario) Then
                        RemoveDetalleItem i
                        i = i - 1
                    End If
                End If
                
                
                If band Then
                    rs1.AddNew
                    rs1!IdProvcli = Me.IdProvcli
                    rs1!IdInventario = Me.Empresa.RecuperaIdInventario(obj.CodInventario)
                    rs1!Cantidad = obj.Cantidad
                    rs1!PU = obj.PU
                    rs1!Frecuencia = obj.Frecuencia
                    rs1!fechagrabado = Now
                    rs1!fechaIni = obj.fechaIni
                    rs1!FechaFin = obj.FechaFin
                    rs1!BandProntoPago = obj.BandProntoPago
                    rs1!Referencia = obj.Referencia
                    rs1!BandPublicidad = obj.BandPublicidad
                    rs1!Contrato = obj.Contrato
                    rs1!Plazo = obj.Plazo
                    rs1.Update
                End If
            End If
            band = True
        Next i
    End If
    Set rs1 = Nothing
    Set obj = Nothing
    Exit Sub
CapturaError:
    MsgBox Err.Description
    Set rs1 = Nothing
    Set obj = Nothing
    Exit Sub
End Sub

Private Function verificacionDuplicado(ByVal IdInventario As String) As Boolean
    Dim rs As Recordset, sql As String
    verificacionDuplicado = False
    'Recupera los detalles de existencia
    sql = "SELECT IdInventario FROM PCProvCliInv where idinventario =" & IdInventario
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        verificacionDuplicado = True
    End If
    Set rs = Nothing
End Function

Public Sub RemoveDetalleItem(ByVal ix As Long, Optional obj As IVDetalleItem)
    If mcolDetalleItem Is Nothing Then CargaDetalleItem
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolDetalleItem.Count
            If mcolDetalleItem.item(ix) Is obj Then
                mcolDetalleItem.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolDetalleItem.Remove ix
    End If
    mbooModificado = True
End Sub

Public Property Let bandEmpleado(ByVal value As Boolean)
    If value <> mProps.bandEmpleado Then
        mProps.bandEmpleado = value
        mbooModificado = True
    End If
End Property

Public Property Get bandEmpleado() As Boolean
    bandEmpleado = mProps.bandEmpleado
End Property

Public Property Get NumSerie() As String
    NumSerie = mProps.NumSerie
End Property

Public Property Let NumSerie(ByVal value As String)
'    If ((Len(value) > 3) Or Len(value) > 0) Or (Not IsNumeric(value)) Then
'        Err.Raise ERR_INVALIDO, "PCProvCli.NumSerie", MSGERR_INVALIDO
'    End If
    mProps.NumSerie = value
    mbooModificado = True
End Property

Public Property Get NumPunto() As String
    NumPunto = mProps.NumPunto
End Property

Public Property Let NumPunto(ByVal value As String)
    If (Len(value) > 3) Or (Not IsNumeric(value)) Then
        Err.Raise ERR_INVALIDO, "PCProvCli.NumPunto", MSGERR_INVALIDO
    End If
    mProps.NumPunto = value
    mbooModificado = True
End Property

Public Property Let FechaCaducidad(value As Date)
    If Not (IsDate(value)) Then
        Err.Raise ERR_INVALIDO, "PCProvCli_FechaCaducidad", MSGERR_INVALIDO
    End If
    mProps.FechaCaducidad = value
    mbooModificado = True
End Property

Public Property Get FechaCaducidad() As Date
    FechaCaducidad = mProps.FechaCaducidad
End Property

Public Property Let IdDiasCredito(ByVal value As Long)
    mProps.IdDiasCredito = value
    mbooModificado = True
End Property

Public Property Get IdDiasCredito() As Long
    IdDiasCredito = mProps.IdDiasCredito
End Property

Public Property Let CodDiasCredito(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodDiasCredito Then
        If Len(value) > 0 Then
            'Actualiza IdGarante también
            sql = "SELECT IdDiasCredito FROM PCDiasCredito WHERE CodDiasCredito='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodDiasCredito", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdDiasCredito = rs!IdDiasCredito
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdDiasCredito = 0
        End If
    End If
    mProps.CodDiasCredito = value
    mbooModificado = True
End Property

Public Property Get CodDiasCredito() As String
    CodDiasCredito = Trim$(mProps.CodDiasCredito)
End Property

Public Sub AnulaDescuentos()
    Dim sql As String, rs As Recordset
    sql = " update ivdescuento set bandvalida=0"
    sql = sql & " from ivdescuento ivd"
    sql = sql & " inner join ivdescuentodetallepc ivdpc"
    sql = sql & " inner join pcprovcli pc"
    sql = sql & " on ivdpc.idprovcli=pc.idprovcli"
    sql = sql & " on ivd.iddescuento=ivdpc.iddescuento"
    sql = sql & " where codprovcli='" & Me.CodProvcli & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
End Sub

Public Property Let IsCopy(ByVal value As Boolean)
    mProps.IsCopy = value
End Property

Public Property Get IsCopy() As Boolean
    IsCopy = Trim$(mProps.IsCopy)
End Property

Public Property Let NumPagos(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_NumPagos", MSGERR_INVALIDO
    End If
    mProps.NumPagos = value
    mbooModificado = True
End Property

Public Property Get NumPagos() As Currency
    NumPagos = mProps.NumPagos
End Property

Public Property Let Intervalo(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_Intervalo", MSGERR_INVALIDO
    End If
    mProps.Intervalo = value
    mbooModificado = True
End Property

Public Property Get Intervalo() As Currency
    Intervalo = mProps.Intervalo
End Property

Public Property Let PorDescNeto(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_PorDescNeto", MSGERR_INVALIDO
    End If
    mProps.PorDescNeto = value
    mbooModificado = True
End Property
Public Property Get PorDescNeto() As Currency
    PorDescNeto = mProps.PorDescNeto
End Property

Public Property Let PorDescPP(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_PorDescPP", MSGERR_INVALIDO
    End If
    mProps.PorDescPP = value
    mbooModificado = True
End Property
Public Property Get PorDescPP() As Currency
    PorDescPP = mProps.PorDescPP
End Property

Friend Sub RecuperarxRUC(key As String, BandCliente As Boolean, BandProveedor As Boolean, bandEmpleado As Boolean)
    Dim sql As String, i As Long
    Dim rs As Recordset
    
    sql = "SELECT * FROM vwPCProvCli WHERE ruc='" & key & "'"
    If BandCliente Then
        sql = sql & " and bandCliente=1"
    Else
        sql = sql & " and bandCliente=0"
    End If
    If BandProveedor Then
        sql = sql & " and bandProveedor=1"
    Else
        sql = sql & " and bandProveedor=0"
    End If
    If bandEmpleado Then
        sql = sql & " and BandEmpleado=1"
    Else
        sql = sql & " and BandEmpleado=0"
    End If
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IdProvcli = !IdProvcli
        mProps.CodProvcli = !CodProvcli
        If Not IsNull(!Nombre) Then mProps.Nombre = !Nombre
        If Not IsNull(!BandProveedor) Then mProps.BandProveedor = !BandProveedor
        If Not IsNull(!BandCliente) Then mProps.BandCliente = !BandCliente
        If Not IsNull(!IdCuentaContable) Then mProps.IdCuentaContable = !IdCuentaContable
        If Not IsNull(!IdCuentaContable2) Then mProps.IdCuentaContable2 = !IdCuentaContable2
        If Not IsNull(!CodCuentaContable) Then mProps.CodCuentaContable = !CodCuentaContable
        If Not IsNull(!CodCuentaContable2) Then mProps.CodCuentaContable2 = !CodCuentaContable2
        If Not IsNull(!Direccion1) Then mProps.Direccion1 = !Direccion1
        If Not IsNull(!Direccion2) Then mProps.Direccion2 = !Direccion2
        If Not IsNull(!CodPostal) Then mProps.CodPostal = !CodPostal
        If Not IsNull(!Ciudad) Then mProps.Ciudad = !Ciudad
        If Not IsNull(!Provincia) Then mProps.Provincia = !Provincia
        If Not IsNull(!Pais) Then mProps.Pais = !Pais
        If Not IsNull(!Telefono1) Then mProps.Telefono1 = !Telefono1
        If Not IsNull(!Telefono2) Then mProps.Telefono2 = !Telefono2
        If Not IsNull(!Telefono3) Then mProps.Telefono3 = !Telefono3
        If Not IsNull(!Fax) Then mProps.Fax = !Fax
        If Not IsNull(!RUC) Then mProps.RUC = !RUC
        If Not IsNull(!EMail) Then mProps.EMail = !EMail
        If Not IsNull(!LimiteCredito) Then mProps.LimiteCredito = !LimiteCredito
        If Not IsNull(!banco) Then mProps.banco = !banco
        If Not IsNull(!NumCuenta) Then mProps.NumCuenta = !NumCuenta
        If Not IsNull(!Swit) Then mProps.Swit = !Swit
        If Not IsNull(!DirecBanco) Then mProps.DirecBanco = !DirecBanco
        If Not IsNull(!TelBanco) Then mProps.TelBanco = !TelBanco
        If Not IsNull(!IdVendedor) Then mProps.IdVendedor = !IdVendedor
        If Not IsNull(!CodVendedor) Then mProps.CodVendedor = !CodVendedor
        If Not IsNull(!IdGrupo1) Then mProps.IdGrupo1 = !IdGrupo1
        If Not IsNull(!CodGrupo1) Then mProps.CodGrupo1 = !CodGrupo1
        If Not IsNull(!IdGrupo2) Then mProps.IdGrupo2 = !IdGrupo2
        If Not IsNull(!CodGrupo2) Then mProps.CodGrupo2 = !CodGrupo2
        
        If Not IsNull(!IdGrupo3) Then mProps.IdGrupo3 = !IdGrupo3
        If Not IsNull(!CodGrupo3) Then mProps.CodGrupo3 = !CodGrupo3
        
        If Not IsNull(!IdGrupo4) Then mProps.IdGrupo4 = !IdGrupo4 'Agregadoa AUC 22/09/2005
        If Not IsNull(!CodGrupo4) Then mProps.CodGrupo4 = !CodGrupo4
        
        If Not IsNull(!Estado) Then mProps.Estado = !Estado
        If Not IsNull(!TipoDocumento) Then mProps.TipoDocumento = !TipoDocumento          'Agregado Alex Sept/2002
        If Not IsNull(!TipoComprobante) Then mProps.TipoComprobante = !TipoComprobante           'Agregado Alex Sept/2002
        If Not IsNull(!NumAutSRI) Then mProps.NumAutSRI = !NumAutSRI             'Agregado Alex Sept/2002
        If Not IsNull(!NombreAlterno) Then mProps.NombreAlterno = !NombreAlterno          '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!FechaNacimiento) Then mProps.FechaNacimiento = !FechaNacimiento    '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!FechaEntrega) Then mProps.FechaEntrega = !FechaEntrega             '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!FechaExpiracion) Then mProps.FechaExpiracion = !FechaExpiracion    '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!TotalDebe) Then mProps.TotalDebe = !TotalDebe                      '***Agregado. 08/sep/2003. Angel
        If Not IsNull(!TotalHaber) Then mProps.TotalHaber = !TotalHaber                   '***Agregado. 08/sep/2003. Angel
        If Not IsNull(!Observacion) Then mProps.Observacion = !Observacion      'Agregado AUC 31/08/2005
        If Not IsNull(!TipoProvCli) Then mProps.TipoProvCli = Trim$(!TipoProvCli)  ' jeaa 17/12/2005
        If Not IsNull(!ReferidoPor) Then mProps.ReferidoPor = !ReferidoPor      'Agregado jeaa 12/08/2006
        'jeaa 15/05/2007
        If Not IsNull(!IdTipoDocumento) Then mProps.IdTipoDocumento = !IdTipoDocumento 'Agregado jeaa 15/05/2007
        If Not IsNull(!codtipoDocumento) Then mProps.codtipoDocumento = !codtipoDocumento 'Agregado jeaa 15/05/2007
        If Not IsNull(!BandEmpresaPublica) Then mProps.BandEmpresaPublica = !BandEmpresaPublica   'jeaa 16/01/2008 cambio por EQUIDAD TRIBUTARIA
        If Not IsNull(!IdGarante) Then mProps.IdGarante = !IdGarante
        If Not IsNull(!CodGarante) Then mProps.CodGarante = !CodGarante
        If Not IsNull(!BandGarante) Then mProps.BandGarante = !BandGarante
        If Not IsNull(!DiasPlazo) Then mProps.DiasPlazo = !DiasPlazo
        If Not IsNull(!IdProvincia) Then mProps.IdProvincia = !IdProvincia
        If Not IsNull(!CodProvincia) Then mProps.CodProvincia = !CodProvincia
        If Not IsNull(!IdCanton) Then mProps.IdCanton = !IdCanton
        If Not IsNull(!CodCanton) Then mProps.CodCanton = !CodCanton
        If Not IsNull(!IDParroquia) Then mProps.IDParroquia = !IDParroquia
        If Not IsNull(!CodParroquia) Then mProps.CodParroqui = !CodParroquia
        If Not IsNull(!FechaCreacion) Then mProps.FechaCreacion = !FechaCreacion
        If Not IsNull(!CodigoUsuario) Then mProps.CodigoUsuario = !CodigoUsuario
        If Not IsNull(!FechaLote1) Then mProps.CodigoUsuario = !FechaLote1
        If Not IsNull(!FechaLote2) Then mProps.CodigoUsuario = !FechaLote2
        If Not IsNull(!BandLote) Then mProps.CodigoUsuario = !BandLote
        If Not IsNull(!bandEmpleado) Then mProps.bandEmpleado = !bandEmpleado
        If Not IsNull(!NumSerie) Then mProps.NumSerie = !NumSerie
        If Not IsNull(!NumPunto) Then mProps.NumPunto = !NumPunto
        If Not IsNull(!FechaCaducidad) Then mProps.FechaCaducidad = !FechaCaducidad
        If Not IsNull(!IdDiasCredito) Then mProps.IdDiasCredito = !IdDiasCredito
        If Not IsNull(!BandRUCValido) Then mProps.BandRUCValido = !BandRUCValido
        If Not IsNull(!BandOmitirPendientes) Then mProps.BandOmitirPendientes = !BandOmitirPendientes
        If Not IsNull(!TipoSujeto) Then mProps.TipoSujeto = !TipoSujeto
        If Not IsNull(!Sexo) Then mProps.Sexo = !Sexo
        If Not IsNull(!EstadoCivil) Then mProps.EstadoCivil = !EstadoCivil
        If Not IsNull(!OrigenIngresos) Then mProps.OrigenIngresos = !OrigenIngresos
        If Not IsNull(!BandOmitirDinardap) Then mProps.BandOmitirDinardap = !BandOmitirDinardap
        If Not IsNull(!BandRelacionado) Then mProps.BandRelacionado = !BandRelacionado
        If Not IsNull(!IdGarante2) Then mProps.IdGarante2 = !IdGarante2
        If Not IsNull(!CodGarante2) Then mProps.CodGarante2 = !CodGarante2
        If Not IsNull(!NumCasa) Then mProps.NumCasa = !NumCasa
        If Not IsNull(!bandCompensaIVA) Then mProps.bandCompensaIVA = !bandCompensaIVA
        If Not IsNull(!BandEsConsumidorFinal) Then mProps.BandEsConsumidorFinal = !BandEsConsumidorFinal
        If Not IsNull(!BandIVCodigoIgual) Then mProps.BandIVCodigoIgual = !BandIVCodigoIgual
        If Not IsNull(!IdCuentaContableGasto) Then mProps.IdCuentaContableGasto = !IdCuentaContableGasto
        If Not IsNull(!CodCuentaContableGasto) Then mProps.CodCuentaContableGasto = !CodCuentaContableGasto
        
        
        
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Property Let BandRUCValido(ByVal value As Boolean)
    If value <> mProps.BandRUCValido Then
        mProps.BandRUCValido = value
        mbooModificado = True
    End If
End Property

Public Property Get BandRUCValido() As Boolean
    BandRUCValido = mProps.BandRUCValido
End Property

Friend Sub RecuperarQuick(key As Variant)
    Dim sql As String, i As Long
    Dim rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM vwPCProvCli WHERE CodProvCli='" & key & "'"
    Else
        sql = "SELECT * FROM vwPCProvCli WHERE IdProvCli=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IdProvcli = !IdProvcli
        mProps.CodProvcli = !CodProvcli
        If Not IsNull(!Nombre) Then mProps.Nombre = !Nombre
        If Not IsNull(!BandProveedor) Then mProps.BandProveedor = !BandProveedor
        If Not IsNull(!BandCliente) Then mProps.BandCliente = !BandCliente
        If Not IsNull(!IdCuentaContable) Then mProps.IdCuentaContable = !IdCuentaContable
        If Not IsNull(!IdCuentaContable2) Then mProps.IdCuentaContable2 = !IdCuentaContable2
        If Not IsNull(!CodCuentaContable) Then mProps.CodCuentaContable = !CodCuentaContable
        If Not IsNull(!CodCuentaContable2) Then mProps.CodCuentaContable2 = !CodCuentaContable2
        If Not IsNull(!Direccion1) Then mProps.Direccion1 = !Direccion1
        If Not IsNull(!Direccion2) Then mProps.Direccion2 = !Direccion2
        If Not IsNull(!CodPostal) Then mProps.CodPostal = !CodPostal
        If Not IsNull(!Ciudad) Then mProps.Ciudad = !Ciudad
        If Not IsNull(!Provincia) Then mProps.Provincia = !Provincia
        If Not IsNull(!Pais) Then
            If Len(!Pais) = 0 Then
                mProps.Pais = "ECUADOR"
            Else
                mProps.Pais = !Pais
            End If
        End If
        If Not IsNull(!Telefono1) Then mProps.Telefono1 = !Telefono1
        If Not IsNull(!Telefono2) Then mProps.Telefono2 = !Telefono2
        If Not IsNull(!Telefono3) Then mProps.Telefono3 = !Telefono3
        If Not IsNull(!Fax) Then mProps.Fax = !Fax
        If Not IsNull(!RUC) Then mProps.RUC = !RUC
        If Not IsNull(!EMail) Then mProps.EMail = !EMail
        If Not IsNull(!LimiteCredito) Then mProps.LimiteCredito = !LimiteCredito
        If Not IsNull(!banco) Then mProps.banco = !banco
        If Not IsNull(!NumCuenta) Then mProps.NumCuenta = !NumCuenta
        If Not IsNull(!Swit) Then mProps.Swit = !Swit
        If Not IsNull(!DirecBanco) Then mProps.DirecBanco = !DirecBanco
        If Not IsNull(!TelBanco) Then mProps.TelBanco = !TelBanco
        If Not IsNull(!IdVendedor) Then mProps.IdVendedor = !IdVendedor
        If Not IsNull(!CodVendedor) Then mProps.CodVendedor = !CodVendedor
        If Not IsNull(!IdGrupo1) Then mProps.IdGrupo1 = !IdGrupo1
        If Not IsNull(!CodGrupo1) Then mProps.CodGrupo1 = !CodGrupo1
        If Not IsNull(!IdGrupo2) Then mProps.IdGrupo2 = !IdGrupo2
        If Not IsNull(!CodGrupo2) Then mProps.CodGrupo2 = !CodGrupo2
        
        If Not IsNull(!IdGrupo3) Then mProps.IdGrupo3 = !IdGrupo3
        If Not IsNull(!CodGrupo3) Then mProps.CodGrupo3 = !CodGrupo3
        
        If Not IsNull(!IdGrupo4) Then mProps.IdGrupo4 = !IdGrupo4 'Agregadoa AUC 22/09/2005
        If Not IsNull(!CodGrupo4) Then mProps.CodGrupo4 = !CodGrupo4
        
        If Not IsNull(!Estado) Then mProps.Estado = !Estado
        If Not IsNull(!TipoDocumento) Then mProps.TipoDocumento = !TipoDocumento          'Agregado Alex Sept/2002
        If Not IsNull(!TipoComprobante) Then mProps.TipoComprobante = !TipoComprobante           'Agregado Alex Sept/2002
        If Not IsNull(!NumAutSRI) Then mProps.NumAutSRI = !NumAutSRI             'Agregado Alex Sept/2002
        If Not IsNull(!NombreAlterno) Then mProps.NombreAlterno = !NombreAlterno          '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!FechaNacimiento) Then mProps.FechaNacimiento = !FechaNacimiento    '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!FechaEntrega) Then mProps.FechaEntrega = !FechaEntrega             '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!FechaExpiracion) Then mProps.FechaExpiracion = !FechaExpiracion    '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!TotalDebe) Then mProps.TotalDebe = !TotalDebe                      '***Agregado. 08/sep/2003. Angel
        If Not IsNull(!TotalHaber) Then mProps.TotalHaber = !TotalHaber                   '***Agregado. 08/sep/2003. Angel
        If Not IsNull(!Observacion) Then mProps.Observacion = !Observacion      'Agregado AUC 31/08/2005
        If Not IsNull(!TipoProvCli) Then mProps.TipoProvCli = Trim$(!TipoProvCli)  ' jeaa 17/12/2005
        If Not IsNull(!ReferidoPor) Then mProps.ReferidoPor = !ReferidoPor      'Agregado jeaa 12/08/2006
        'jeaa 15/05/2007
        If Not IsNull(!IdTipoDocumento) Then mProps.IdTipoDocumento = !IdTipoDocumento 'Agregado jeaa 15/05/2007
        If Not IsNull(!codtipoDocumento) Then mProps.codtipoDocumento = !codtipoDocumento 'Agregado jeaa 15/05/2007
        If Not IsNull(!BandEmpresaPublica) Then mProps.BandEmpresaPublica = !BandEmpresaPublica   'jeaa 16/01/2008 cambio por EQUIDAD TRIBUTARIA
        If Not IsNull(!IdGarante) Then mProps.IdGarante = !IdGarante
        If Not IsNull(!CodGarante) Then mProps.CodGarante = !CodGarante
        If Not IsNull(!BandGarante) Then mProps.BandGarante = !BandGarante
        If Not IsNull(!DiasPlazo) Then mProps.DiasPlazo = !DiasPlazo
        If Not IsNull(!CodigoUsuario) Then mProps.CodigoUsuario = !CodigoUsuario
        If Not IsNull(!IdProvincia) Then mProps.IdProvincia = !IdProvincia
        If Not IsNull(!CodProvincia) Then mProps.CodProvincia = !CodProvincia
        If Not IsNull(!bandEmpleado) Then mProps.bandEmpleado = !bandEmpleado

'        Else
'            mProps.CodProvincia = mProps.objEmpresa.GNOpcion.CodProvinciaPre
'        End If
        If Not IsNull(!IdCanton) Then mProps.IdCanton = !IdCanton
        If Not IsNull(!CodCanton) Then mProps.CodCanton = !CodCanton
'        Else
'            mProps.CodCanton = mProps.objEmpresa.GNOpcion.CodCantonPre
'        End If
        If Not IsNull(!IDParroquia) Then mProps.IDParroquia = !IDParroquia
        If Not IsNull(!CodParroquia) Then mProps.CodParroqui = !CodParroquia
        If Not IsNull(!FechaCreacion) Then mProps.FechaCreacion = !FechaCreacion
        
        If Not IsNull(!FechaLote1) Then mProps.FechaLote1 = !FechaLote1
        If Not IsNull(!FechaLote2) Then mProps.FechaLote2 = !FechaLote2
        If Not IsNull(!BandLote) Then mProps.BandLote = !BandLote
        If Not IsNull(!bandEmpleado) Then mProps.bandEmpleado = !bandEmpleado
        If Not IsNull(!NumSerie) Then mProps.NumSerie = !NumSerie
        If Not IsNull(!NumPunto) Then mProps.NumPunto = !NumPunto
        If Not IsNull(!FechaCaducidad) Then mProps.FechaCaducidad = !FechaCaducidad
        If Not IsNull(!IdDiasCredito) Then mProps.IdDiasCredito = !IdDiasCredito
        If Not IsNull(!CodDiasCredito) Then mProps.CodDiasCredito = !CodDiasCredito
        If Not IsNull(!NumPagos) Then mProps.NumPagos = !NumPagos
        If Not IsNull(!Intervalo) Then mProps.Intervalo = !Intervalo
        If Not IsNull(!PorDescNeto) Then mProps.PorDescNeto = !PorDescNeto
        If Not IsNull(!PorDescPP) Then mProps.PorDescPP = !PorDescPP
        If Not IsNull(!BandRUCValido) Then mProps.BandRUCValido = !BandRUCValido
        If Not IsNull(!BandOmitirPendientes) Then mProps.BandOmitirPendientes = !BandOmitirPendientes
        If Not IsNull(!TipoSujeto) Then mProps.TipoSujeto = !TipoSujeto
        If Not IsNull(!Sexo) Then mProps.Sexo = !Sexo
        If Not IsNull(!EstadoCivil) Then mProps.EstadoCivil = !EstadoCivil
        If Not IsNull(!OrigenIngresos) Then mProps.OrigenIngresos = !OrigenIngresos
        If Not IsNull(!BandOmitirDinardap) Then mProps.BandOmitirDinardap = !BandOmitirDinardap
        If Not IsNull(!BandRelacionado) Then mProps.BandRelacionado = !BandRelacionado
        If Not IsNull(!IdGarante2) Then mProps.IdGarante2 = !IdGarante2
        If Not IsNull(!CodGarante2) Then mProps.CodGarante2 = !CodGarante2
        If Not IsNull(!NumCasa) Then mProps.NumCasa = !NumCasa
        If Not IsNull(!bandCompensaIVA) Then mProps.bandCompensaIVA = !bandCompensaIVA
        If Not IsNull(!BandEsConsumidorFinal) Then mProps.BandEsConsumidorFinal = !BandEsConsumidorFinal
        If Not IsNull(!BandIVCodigoIgual) Then mProps.BandIVCodigoIgual = !BandIVCodigoIgual
        If Not IsNull(!IdCuentaContableGasto) Then mProps.IdCuentaContableGasto = !IdCuentaContableGasto
        If Not IsNull(!CodCuentaContableGasto) Then mProps.CodCuentaContableGasto = !CodCuentaContableGasto
        
        

        
        
        mProps.fechagrabado = !fechagrabado
        
        .Close
    End With
    
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Property Let Codturno(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.Codturno Then
        If Len(value) > 0 Then
            'Actualiza IdGarante también
            sql = "SELECT IdTurno FROM RTurno WHERE Codturno='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodTurno", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdTurno = rs!IdTurno
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdTurno = 0
        End If
    End If
    mProps.Codturno = value
    mbooModificado = True
End Property

Public Property Get Codturno() As String
    Codturno = Trim$(mProps.Codturno)
End Property

Public Property Get IdTurno() As Long
    IdTurno = mProps.IdTurno
End Property

Public Property Let IdTurno(ByVal value As Long)
    mProps.IdTurno = value
    mbooModificado = True
End Property

Public Property Get TipoTurno() As Byte
    TipoTurno = mProps.TipoTurno
End Property

Public Property Let TipoTurno(ByVal value As Byte)
    mProps.TipoTurno = value
    mbooModificado = True
End Property

Friend Sub RecuperarpcprovcliReloj(key As Variant)
    Dim sql As String, i As Long
    Dim rs As Recordset
    If VarType(key) = vbString Then
        sql = "SELECT  pc.idprovcli,pc.codprovcli,pc.nombre,pc.direccion1,pc.telefono1,p.codigoreloj FROM  Empleado pc inner join personal p on p.idEmpleado = pc.idprovcli WHERE Codigoreloj='" & key & "'"
    Else
        sql = "SELECT pc.idprovcli,pc.codprovcli,pc.nombre,pc.direccion1,pc.telefono1,p.codigoreloj FROM  Empleado pc inner join personal p on p.idEmpleado = pc.idprovcli WHERE Codigoreloj=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'Recupera los datos de Usuario
    With rs
        mProps.IdProvcli = !IdProvcli
        mProps.CodProvcli = !CodProvcli
        If Not IsNull(!CodigoReloj) Then mProps.CodigoReloj = !CodigoReloj
        If Not IsNull(!Nombre) Then mProps.Nombre = !Nombre
        If Not IsNull(!Direccion1) Then mProps.Direccion1 = !Direccion1
        If Not IsNull(!Telefono1) Then mProps.Telefono1 = !Telefono1
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub


Public Property Let BandRelacionado(ByVal value As Boolean)
    If value <> mProps.BandRelacionado Then
        mProps.BandRelacionado = value
        mbooModificado = True
    End If
End Property

Public Property Get BandRelacionado() As Boolean
    BandRelacionado = mProps.BandRelacionado
End Property

Public Property Let CodigoReloj(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_CodigoReloj", MSGERR_INVALIDO
    End If
    
    mProps.CodigoReloj = value
    mbooModificado = True
End Property

Public Property Get CodigoReloj() As String
    CodigoReloj = Trim$(mProps.CodigoReloj)
End Property

Public Property Let BandOmitirPendientes(ByVal value As Boolean)
    If value <> mProps.BandOmitirPendientes Then
        mProps.BandOmitirPendientes = value
        mbooModificado = True
    End If
End Property

Public Property Get BandOmitirPendientes() As Boolean
    BandOmitirPendientes = mProps.BandOmitirPendientes
End Property

Public Property Let TipoSujeto(ByVal value As String)
    If Len(value) > 1 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_TipoSujeto", MSGERR_INVALIDO
    End If
    
    If value <> mProps.TipoSujeto Then
        mProps.TipoSujeto = value
        mbooModificado = True
    End If
End Property

Public Property Get TipoSujeto() As String
    TipoSujeto = mProps.TipoSujeto
End Property

Public Property Let Sexo(ByVal value As String)
    If Len(value) > 1 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_Sexo", MSGERR_INVALIDO
    End If
    
    If value <> mProps.Sexo Then
        mProps.Sexo = value
        mbooModificado = True
    End If
End Property

Public Property Get Sexo() As String
    Sexo = mProps.Sexo
End Property

Public Property Let EstadoCivil(ByVal value As String)
    If Len(value) > 1 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_EstadoCivil", MSGERR_INVALIDO
    End If
    
    If value <> mProps.EstadoCivil Then
        mProps.EstadoCivil = value
        mbooModificado = True
    End If
End Property

Public Property Get EstadoCivil() As String
    EstadoCivil = mProps.EstadoCivil
End Property


Public Property Let OrigenIngresos(ByVal value As String)
    If Len(value) > 1 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_OrigenIngresos", MSGERR_INVALIDO
    End If
    
    If value <> mProps.OrigenIngresos Then
        mProps.OrigenIngresos = value
        mbooModificado = True
    End If
End Property

Public Property Get OrigenIngresos() As String
    OrigenIngresos = mProps.OrigenIngresos
End Property

Public Property Let BandSolicitudCredito(ByVal value As Boolean)
    If value <> mProps.BandSolicitudCredito Then
        mProps.BandSolicitudCredito = value
       mbooModificado = True
        AsignarCodCuentaPre             '*** MAKOTO 26/mar/01
    End If
End Property
Public Property Get BandSolicitudCredito() As Boolean
    BandSolicitudCredito = mProps.BandSolicitudCredito
End Property
Public Property Let DescripcionDetalle(ByVal value As String)
    mProps.DescripcionDetalle = value
    mbooModificado = True
End Property
Public Property Get DescripcionDetalle() As String
    DescripcionDetalle = Trim$(mProps.DescripcionDetalle)
End Property
Public Property Let NombreConyuge(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_NombreConyuge", MSGERR_INVALIDO
   End If
    mProps.NombreConyuge = value
   mbooModificado = True
End Property
Public Property Get NombreConyuge() As String
    NombreConyuge = Trim$(mProps.NombreConyuge)
End Property
Public Property Let RucConyuge(ByVal value As String)
    If Len(value) > 19 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_RucConyuge", MSGERR_INVALIDO
   End If
    mProps.RucConyuge = value
   mbooModificado = True
End Property
'para el uso de Anexos
Public Property Get RucConyuge() As String
    RucConyuge = mProps.RucConyuge
End Property
Public Property Let BandSuperGarante(ByVal value As Boolean)
    If value <> mProps.BandSuperGarante Then
        mProps.BandSuperGarante = value
       mbooModificado = True
    End If
End Property
Public Property Get BandSuperGarante() As Boolean
    BandSuperGarante = mProps.BandSuperGarante
End Property
'

Friend Sub RecuperarEmpleado(key As Variant)
    Dim sql As String, i As Long
    Dim rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM vwEmpleado WHERE CodProvCli='" & key & "'"
    Else
        sql = "SELECT * FROM vwEmpleado WHERE IdProvCli=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IdProvcli = !IdProvcli
        mProps.CodProvcli = !CodProvcli
        If Not IsNull(!Nombre) Then mProps.Nombre = !Nombre
        If Not IsNull(!BandProveedor) Then mProps.BandProveedor = !BandProveedor
        If Not IsNull(!BandCliente) Then mProps.BandCliente = !BandCliente
        If Not IsNull(!IdCuentaContable) Then mProps.IdCuentaContable = !IdCuentaContable
        If Not IsNull(!IdCuentaContable2) Then mProps.IdCuentaContable2 = !IdCuentaContable2
        If Not IsNull(!CodCuentaContable) Then mProps.CodCuentaContable = !CodCuentaContable
        If Not IsNull(!CodCuentaContable2) Then mProps.CodCuentaContable2 = !CodCuentaContable2
        If Not IsNull(!Direccion1) Then mProps.Direccion1 = !Direccion1
        If Not IsNull(!Direccion2) Then mProps.Direccion2 = !Direccion2
        If Not IsNull(!CodPostal) Then mProps.CodPostal = !CodPostal
        If Not IsNull(!Ciudad) Then mProps.Ciudad = !Ciudad
        If Not IsNull(!Provincia) Then mProps.Provincia = !Provincia
        If Not IsNull(!Pais) Then
            If Len(!Pais) = 0 Then
                mProps.Pais = "ECUADOR"
            Else
                mProps.Pais = !Pais
            End If
        End If
        If Not IsNull(!Telefono1) Then mProps.Telefono1 = !Telefono1
        If Not IsNull(!Telefono2) Then mProps.Telefono2 = !Telefono2
        If Not IsNull(!Telefono3) Then mProps.Telefono3 = !Telefono3
        If Not IsNull(!Fax) Then mProps.Fax = !Fax
        If Not IsNull(!RUC) Then mProps.RUC = !RUC
        If Not IsNull(!EMail) Then mProps.EMail = !EMail
        If Not IsNull(!LimiteCredito) Then mProps.LimiteCredito = !LimiteCredito
        If Not IsNull(!banco) Then mProps.banco = !banco
        If Not IsNull(!NumCuenta) Then mProps.NumCuenta = !NumCuenta
        If Not IsNull(!Swit) Then mProps.Swit = !Swit
        If Not IsNull(!DirecBanco) Then mProps.DirecBanco = !DirecBanco
        If Not IsNull(!TelBanco) Then mProps.TelBanco = !TelBanco
        If Not IsNull(!IdVendedor) Then mProps.IdVendedor = !IdVendedor
        If Not IsNull(!CodVendedor) Then mProps.CodVendedor = !CodVendedor
        If Not IsNull(!IdGrupo1) Then mProps.IdGrupo1 = !IdGrupo1
        If Not IsNull(!CodGrupo1) Then mProps.CodGrupo1 = !CodGrupo1
        If Not IsNull(!IdGrupo2) Then mProps.IdGrupo2 = !IdGrupo2
        If Not IsNull(!CodGrupo2) Then mProps.CodGrupo2 = !CodGrupo2
        
        If Not IsNull(!IdGrupo3) Then mProps.IdGrupo3 = !IdGrupo3
        If Not IsNull(!CodGrupo3) Then mProps.CodGrupo3 = !CodGrupo3
        
        If Not IsNull(!IdGrupo4) Then mProps.IdGrupo4 = !IdGrupo4 'Agregadoa AUC 22/09/2005
        If Not IsNull(!CodGrupo4) Then mProps.CodGrupo4 = !CodGrupo4
        
        If Not IsNull(!Estado) Then mProps.Estado = !Estado
        If Not IsNull(!TipoDocumento) Then mProps.TipoDocumento = !TipoDocumento          'Agregado Alex Sept/2002
        If Not IsNull(!TipoComprobante) Then mProps.TipoComprobante = !TipoComprobante           'Agregado Alex Sept/2002
        If Not IsNull(!NumAutSRI) Then mProps.NumAutSRI = !NumAutSRI             'Agregado Alex Sept/2002
        If Not IsNull(!NombreAlterno) Then mProps.NombreAlterno = !NombreAlterno          '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!FechaNacimiento) Then mProps.FechaNacimiento = !FechaNacimiento    '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!FechaEntrega) Then mProps.FechaEntrega = !FechaEntrega             '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!FechaExpiracion) Then mProps.FechaExpiracion = !FechaExpiracion    '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!TotalDebe) Then mProps.TotalDebe = !TotalDebe                      '***Agregado. 08/sep/2003. Angel
        If Not IsNull(!TotalHaber) Then mProps.TotalHaber = !TotalHaber                   '***Agregado. 08/sep/2003. Angel
        If Not IsNull(!Observacion) Then mProps.Observacion = !Observacion      'Agregado AUC 31/08/2005
        If Not IsNull(!TipoProvCli) Then mProps.TipoProvCli = Trim$(!TipoProvCli)  ' jeaa 17/12/2005
        If Not IsNull(!ReferidoPor) Then mProps.ReferidoPor = !ReferidoPor      'Agregado jeaa 12/08/2006
        'jeaa 15/05/2007
        If Not IsNull(!IdTipoDocumento) Then mProps.IdTipoDocumento = !IdTipoDocumento 'Agregado jeaa 15/05/2007
        If Not IsNull(!codtipoDocumento) Then mProps.codtipoDocumento = !codtipoDocumento 'Agregado jeaa 15/05/2007
        If Not IsNull(!BandEmpresaPublica) Then mProps.BandEmpresaPublica = !BandEmpresaPublica   'jeaa 16/01/2008 cambio por EQUIDAD TRIBUTARIA
        If Not IsNull(!IdGarante) Then mProps.IdGarante = !IdGarante
        If Not IsNull(!CodGarante) Then mProps.CodGarante = !CodGarante
        If Not IsNull(!BandGarante) Then mProps.BandGarante = !BandGarante
        If Not IsNull(!DiasPlazo) Then mProps.DiasPlazo = !DiasPlazo
        If Not IsNull(!CodigoUsuario) Then mProps.CodigoUsuario = !CodigoUsuario
        If Not IsNull(!IdProvincia) Then mProps.IdProvincia = !IdProvincia
        If Not IsNull(!CodProvincia) Then mProps.CodProvincia = !CodProvincia
        If Not IsNull(!bandEmpleado) Then mProps.bandEmpleado = !bandEmpleado
        
'        Else
'            mProps.CodProvincia = mProps.objEmpresa.GNOpcion.CodProvinciaPre
'        End If
        If Not IsNull(!IdCanton) Then mProps.IdCanton = !IdCanton
        If Not IsNull(!CodCanton) Then mProps.CodCanton = !CodCanton
'        Else
'            mProps.CodCanton = mProps.objEmpresa.GNOpcion.CodCantonPre
'        End If

        If Not IsNull(!IDParroquia) Then mProps.IDParroquia = !IDParroquia
        If Not IsNull(!CodParroquia) Then mProps.CodParroqui = !CodParroquia
        If Not IsNull(!FechaCreacion) Then mProps.FechaCreacion = !FechaCreacion
        
        If Not IsNull(!FechaLote1) Then mProps.FechaLote1 = !FechaLote1
        If Not IsNull(!FechaLote2) Then mProps.FechaLote2 = !FechaLote2
        If Not IsNull(!BandLote) Then mProps.BandLote = !BandLote
        If Not IsNull(!bandEmpleado) Then mProps.bandEmpleado = !bandEmpleado
        If Not IsNull(!NumSerie) Then mProps.NumSerie = !NumSerie
        If Not IsNull(!NumPunto) Then mProps.NumPunto = !NumPunto
        If Not IsNull(!FechaCaducidad) Then mProps.FechaCaducidad = !FechaCaducidad
        If Not IsNull(!IdDiasCredito) Then mProps.IdDiasCredito = !IdDiasCredito
        If Not IsNull(!CodDiasCredito) Then mProps.CodDiasCredito = !CodDiasCredito
        If Not IsNull(!NumPagos) Then mProps.NumPagos = !NumPagos
        If Not IsNull(!Intervalo) Then mProps.Intervalo = !Intervalo
        If Not IsNull(!PorDescNeto) Then mProps.PorDescNeto = !PorDescNeto
        If Not IsNull(!PorDescPP) Then mProps.PorDescPP = !PorDescPP
        If Not IsNull(!BandRUCValido) Then mProps.BandRUCValido = !BandRUCValido
         If Not IsNull(!IdTurno) Then mProps.IdTurno = !IdTurno
        If Not IsNull(!Codturno) Then mProps.Codturno = !Codturno
        If Not IsNull(!TipoTurno) Then mProps.TipoTurno = !TipoTurno
        If Not IsNull(!BandRelacionado) Then mProps.BandRelacionado = !BandRelacionado
        If Not IsNull(!BandOmitirPendientes) Then mProps.BandOmitirPendientes = !BandOmitirPendientes
        
        If Not IsNull(!TipoSujeto) Then mProps.TipoSujeto = !TipoSujeto
        If Not IsNull(!Sexo) Then mProps.Sexo = !Sexo
        If Not IsNull(!EstadoCivil) Then mProps.EstadoCivil = !EstadoCivil
        If Not IsNull(!OrigenIngresos) Then mProps.OrigenIngresos = !OrigenIngresos
        If Not IsNull(!BandSolicitudCredito) Then mProps.BandSolicitudCredito = !BandSolicitudCredito
         If Not IsNull(!DescripcionDetalle) Then
            If .Fields("DescripcionDetalle").ActualSize > 0 Then
                mProps.DescripcionDetalle = .Fields("DescripcionDetalle").GetChunk(.Fields("DescripcionDetalle").ActualSize)
            End If
        End If
        If Not IsNull(!NombreConyuge) Then mProps.NombreConyuge = !NombreConyuge
        If Not IsNull(!RucConyuge) Then mProps.RucConyuge = !RucConyuge
        mProps.fechagrabado = !fechagrabado
        
        .Close
    End With
    CargaDetalleItem

    
    'Recupera los Contactos
    sql = "SELECT IdContacto FROM PCContacto WHERE IdProvCli=" & mProps.IdProvcli & " ORDER BY Orden"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddContacto
            Contactos(i).Recuperar !IdContacto
            .MoveNext
        Loop
        .Close
    End With
        
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Function VerificaCedulaEmpRepetida() As Boolean
    Dim rs As Recordset, sql As String
    VerificaCedulaEmpRepetida = False
    sql = "SELECT  RUC FROM Empleado "
    sql = sql & " WHERE (RUC='" & mProps.RUC & "') AND "
    sql = sql & " (IdProvCli<>" & mProps.IdProvcli & ")"
    If (mProps.objEmpresa.GNOpcion.ObtenerValor("DuplicaRUC")) = "0" Then
        sql = sql & " AND (IdGrupo1 = " & mProps.IdGrupo1 & " ) "
        sql = sql & " AND (IdGrupo4 = " & mProps.IdGrupo4 & " ) "
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        VerificaCedulaEmpRepetida = True
    End If
    rs.Close
End Function

Public Sub GrabarEmpleado()
    Dim sql As String, fgrabado As Date, i As Long
    Dim rs As Recordset
    Dim rsSuc As Recordset, gnSuc As GNSucursal
    Dim NumGrupo As Integer, ivg As ivgrupo, campo As String, k As Long, n As Byte
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
       
    'Verifica si no repite el código
    sql = "SELECT CodProvCli FROM Empleado " & _
          "WHERE (CodProvCli='" & mProps.CodProvcli & "') AND " & _
                "(IdProvCli<>" & mProps.IdProvcli & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "Empleado", MSGERR_REPITECODIGO
    End If
    rs.Close
    On Error GoTo ErrTrap

#If DAOLIB Then
    gobjWS.BeginTrans                           '======= Comienza la transacción
#Else
    mProps.objEmpresa.Coneccion.BeginTrans      '======= Comienza la transacción
#End If
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM Empleado WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Elimina todos los Contactos
        sql = "DELETE FROM PCContacto WHERE IdProvCli=" & mProps.IdProvcli
        mProps.objEmpresa.Execute sql, True
    
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM Empleado WHERE IdProvCli=" & mProps.IdProvcli
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "Empleado_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en la base
    With rs
        !CodProvcli = mProps.CodProvcli
        !Nombre = mProps.Nombre
        !BandProveedor = mProps.BandProveedor
        !BandCliente = mProps.BandCliente
        !IdCuentaContable = mProps.IdCuentaContable
        !IdCuentaContable2 = mProps.IdCuentaContable2
        !Direccion1 = mProps.Direccion1
        !Direccion2 = mProps.Direccion2
        !CodPostal = mProps.CodPostal
        !Ciudad = mProps.Ciudad
        !Provincia = mProps.Provincia
        !Pais = mProps.Pais
        !Telefono1 = mProps.Telefono1
        !Telefono2 = mProps.Telefono2
        !Telefono3 = mProps.Telefono3
        !Fax = mProps.Fax
        
        If Not !RUC = mProps.RUC Then
            Me.Empresa.GrabaAccionModulo !RUC, mProps.RUC, "RUC", "MOD_CLI", "CL"
        End If
        !RUC = mProps.RUC
        !EMail = mProps.EMail
        
        If Not !LimiteCredito = mProps.LimiteCredito Then
            Me.Empresa.GrabaAccionModulo !LimiteCredito, mProps.LimiteCredito, "LimiteCredito", "MOD_CLI", "CL"
        End If
        !LimiteCredito = mProps.LimiteCredito
        !banco = mProps.banco
        !NumCuenta = mProps.NumCuenta
        !Swit = mProps.Swit
        !DirecBanco = mProps.DirecBanco
        !TelBanco = mProps.TelBanco
        !IdVendedor = mProps.IdVendedor
        !IdGrupo1 = mProps.IdGrupo1
        !IdGrupo2 = mProps.IdGrupo2
        !IdGrupo3 = mProps.IdGrupo3
        
        !IdGrupo4 = mProps.IdGrupo4 'agregado AUC 22/09/2005
        
        If Not !Estado = mProps.Estado Then
            Me.Empresa.GrabaAccionModulo !Estado, mProps.Estado, "Estado", "MOD_CLI", "CL"
        End If
        !Estado = mProps.Estado
        !TipoDocumento = mProps.TipoDocumento           'Agregado Alex: Sept/2002
        !TipoComprobante = mProps.TipoComprobante       'Agregado Alex: Sept/2002
        !NumAutSRI = mProps.NumAutSRI                   'Agregado Alex: Sept/2002
        !NombreAlterno = mProps.NombreAlterno           '***Agregado. 05/sep/2003. Angel
        !FechaNacimiento = mProps.FechaNacimiento       '***Agregado. 05/sep/2003. Angel
        !FechaEntrega = mProps.FechaEntrega             '***Agregado. 05/sep/2003. Angel
        !FechaExpiracion = mProps.FechaExpiracion       '***Agregado. 05/sep/2003. Angel
        !TotalDebe = mProps.TotalDebe                   '***Agregado. 08/sep/2003. Angel
        !TotalHaber = mProps.TotalHaber                 '***Agregado. 08/sep/2003. Angel
        !Observacion = mProps.Observacion  'Agregado AUC 31/08/2005
        !TipoProvCli = mProps.TipoProvCli       ' jeaa 17/12/2005
        !ReferidoPor = mProps.ReferidoPor  'Agregado jeaa 12/08/2006
        !IdTipoDocumento = mProps.IdTipoDocumento   'Agregado jeaa 15/05/2007
        !BandEmpresaPublica = mProps.BandEmpresaPublica   'jeaa 16/01/2008 cambio por EQUIDAD TRIBUTARIA
        !IdGarante = mProps.IdGarante
        !BandGarante = mProps.BandGarante
        !bandEmpleado = mProps.bandEmpleado
        If Not !DiasPlazo = mProps.DiasPlazo Then
            Me.Empresa.GrabaAccionModulo !DiasPlazo, mProps.DiasPlazo, "DiasPlazo", "MOD_CLI", "CL"
        End If
        !DiasPlazo = mProps.DiasPlazo
        
        If Not !CodigoUsuario = mProps.CodigoUsuario Then
            Me.Empresa.GrabaAccionModulo !CodigoUsuario, mProps.CodigoUsuario, "CodigoUsuario", "MOD_CLI", "CL"
        End If
        
        !CodigoUsuario = mProps.CodigoUsuario
        !IdProvincia = mProps.IdProvincia
        !IdCanton = mProps.IdCanton
        !IDParroquia = mProps.IDParroquia
        !FechaCreacion = mProps.FechaCreacion
        !FechaLote1 = mProps.FechaLote1
        !FechaLote2 = mProps.FechaLote2
        !BandLote = mProps.BandLote
        !bandEmpleado = mProps.bandEmpleado
        !NumPunto = mProps.NumPunto
        !NumSerie = mProps.NumSerie
        !FechaCaducidad = mProps.FechaCaducidad
        !IdDiasCredito = mProps.IdDiasCredito
        !NumPagos = mProps.NumPagos
        !Intervalo = mProps.Intervalo
        !PorDescNeto = mProps.PorDescNeto
        !PorDescPP = mProps.PorDescPP
        !BandRUCValido = mProps.BandRUCValido
         !IdTurno = mProps.IdTurno
        !TipoTurno = mProps.TipoTurno
        !BandRelacionado = mProps.BandRelacionado
        
        If Not !BandOmitirPendientes = mProps.BandOmitirPendientes Then
            Me.Empresa.GrabaAccionModulo !BandOmitirPendientes, mProps.BandOmitirPendientes, "BandOmitirPendientes", "MOD_CLI", "CL"
        End If
        !BandOmitirPendientes = mProps.BandOmitirPendientes
        
        !TipoSujeto = mProps.TipoSujeto
        !Sexo = mProps.Sexo
        !EstadoCivil = mProps.EstadoCivil
        !OrigenIngresos = mProps.OrigenIngresos
        !BandSuperGarante = mProps.BandSuperGarante
        !BandSolicitudCredito = mProps.BandSolicitudCredito
        .Fields("DescripcionDetalle").AppendChunk mProps.DescripcionDetalle
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        fgrabado = !fechagrabado
        mProps.IdProvcli = !IdProvcli
        .Close
    End With
    
    'Graba todos los PCContacto
'    For i = 1 To CountContacto
'        Contactos(i).Grabar
'    Next i
    If mbooEsNuevo Then
        AgregaPersonal
    End If
    
    
    
#If DAOLIB Then
    gobjWS.CommitTrans                  '======== Completa la transacción
#Else
    mProps.objEmpresa.Coneccion.CommitTrans      '======== Completa la transacción
#End If
    mProps.fechagrabado = fgrabado
'    GrabarIVDetalle 'guarda configuracion para facturacion x lote x cliente
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_EMP", "Empleado Nuevo: " & _
                                                         mProps.CodProvcli, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_EMP", "Empleado Modificado: " & _
                                                         mProps.CodProvcli, "DT"
    End If
    
        'actualiza Inventario en Bases de Datos Sucursales
'    sql = "SELECT * FROM GnSucursal WHERE bandvalida=1"
'    Set rsSuc = mProps.objEmpresa.OpenRecordset(sql)
    
'    For i = 1 To rsSuc.RecordCount
'        Set gnSuc = mProps.objEmpresa.RecuperaGNSucursal(rsSuc.Fields("CodSucursal"))
'        If gnSuc.BandCliente Or gnSuc.BandProveedor Then
'            GrabarEnSucursales gnSuc.BaseDatos, gnSuc.servidor
'        End If
'        rsSuc.MoveNext
'    Next i
    Set gnSuc = Nothing
    Set rsSuc = Nothing

    ' si esta activado opcion de anexar a ivgrupo
'    If Len(mProps.objEmpresa.GNOpcion.ObtenerValor("NumAnexarCatProvxIVG")) > 0 Then
'        NumGrupo = mProps.objEmpresa.GNOpcion.ObtenerValor("NumAnexarCatProvxIVG")
'    Else
'        NumGrupo = 1
'    End If
'    n = NumGrupo
'    If Len(mProps.objEmpresa.GNOpcion.ObtenerValor("AnexarCatProvxIVG")) > 0 Then
'        If mProps.objEmpresa.GNOpcion.ObtenerValor("AnexarCatProvxIVG") = "1" Then
'            If mProps.BandProveedor Then
'                If mbooEsNuevo Then
'                    sql = "SELECT * FROM ivgrupo" & NumGrupo & " WHERE 1=0"
'                    Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
'                    rs.AddNew
'                    Set ivg = mProps.objEmpresa.CreaIVGrupo(n)
'                Else
'                    Set ivg = mProps.objEmpresa.RecuperaIVGrupo(n, mProps.CodProvCli)
'                End If
'                If Not ivg Is Nothing Then
'                With ivg
'                    .CodGrupo = mProps.CodProvCli
'                    .Descripcion = mProps.Nombre
'                    .bandValida = IIf(mProps.Estado = 0, 1, 0)
'                    .Grabar
'                End With
'                End If
'            End If
'        End If
'    End If

    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PROVCLI
    Exit Sub
ErrTrap:
#If DAOLIB Then
    gobjWS.Rollback                   '======== Anula la transacción
#Else
    mProps.objEmpresa.Coneccion.RollbackTrans  '======== Anula la transacción
#End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Public Sub EliminarEmpleado()
    Dim sql As String, rs As Recordset
    Dim NumGrupo As Integer, ivg As ivgrupo, campo As String, k As Long, n As Byte
    'Verifica si existe Kardex
    sql = "SELECT IdProvCli FROM PCKardex WHERE idelemento <> 0 And IdProvCli=" & mProps.IdProvcli
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "Empleado.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing
    'Verifica si no esta asignado como Cliente
    sql = "SELECT IdEmpleadoRef FROM GNOferta WHERE IdEmpleadoRef=" & mProps.IdProvcli
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "PCProvCli.EliminarEmpleado", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing
    sql = "SELECT IdEmpleado FROM Roldetalle WHERE IdEmpleado=" & mProps.IdProvcli
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "PCProvCli.EliminarEmpleado", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing
    
    sql = "SELECT IdEmpleado FROM IRKardex WHERE IdEmpleado=" & mProps.IdProvcli
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "PCProvCli.EliminarEmpleado", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing
    
    sql = "DELETE FROM EmpCurriculum WHERE idEmpleado=" & mProps.IdProvcli
    mProps.objEmpresa.Execute sql, True
    
    sql = "DELETE FROM EmpHistorial WHERE idEmpleado=" & mProps.IdProvcli
    mProps.objEmpresa.Execute sql, True
    
    'Elimina el registro
    sql = "DELETE FROM Personal WHERE idEmpleado=" & mProps.IdProvcli
    mProps.objEmpresa.Execute sql, True
    'Elimina el registro
    sql = "DELETE FROM Empleado WHERE CodProvCli='" & mProps.CodProvcli & "'"
    mProps.objEmpresa.Execute sql, True
        
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
     mProps.objEmpresa.GrabaGNLogAccion "ELI_EMP", "Empleado Eliminado: " & _
                                                         mProps.CodProvcli, "DT"
    mbooEsNuevo = True
    mbooModificado = True
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PROVCLI
End Sub


Public Function RecuperarFichamedicaEmp(ByVal key As Variant) As FichaMedicaEmp
    Dim obj As FichaMedicaEmp
    Set obj = New FichaMedicaEmp
    Set obj.mobjEmpleado = Me
    If Len(key) > 0 Then obj.Recuperar key
    Set RecuperarFichamedicaEmp = obj
    Set obj = Nothing
End Function

Public Function RecuperarFichaMedFamOcu(ByVal key As Variant) As FichamedFamilTrbjOcu
    Dim obj As FichamedFamilTrbjOcu
    Set obj = New FichamedFamilTrbjOcu
    Set obj.mobjEmpleado = Me
    If Len(key) > 0 Then obj.Recuperar key
    Set RecuperarFichaMedFamOcu = obj
    Set obj = Nothing
End Function

Public Function CreaFichaExaDetalle() As FichaMedicaExamenes
    Dim obj As FichaMedicaExamenes
    Set obj = New FichaMedicaExamenes
    Set obj.mobjEmpleado = Me
    Set CreaFichaExaDetalle = obj
    Set obj = Nothing
End Function

Public Function RecuperaFichaMedicaDetalle(key As Variant) As FichaMedicaExamenes
    Dim obj As FichaMedicaExamenes
    Set obj = New FichaMedicaExamenes
    Set obj.mobjEmpleado = Me
    On Error Resume Next
    obj.Recuperar key
    If Err.Number = 0 Then Set RecuperaFichaMedicaDetalle = obj
    Err.Clear
    On Error GoTo 0
    Set obj = Nothing
End Function

Private Function QuitarEspacios(cad As String) As String
    Dim W As Variant, i As Integer, Nombre As String
    If Len(cad) = 0 Then Exit Function
    W = Split(cad, " ")
    Nombre = ""
    For i = 0 To UBound(W)
        If W(i) = "" Then
        Else
            Nombre = Nombre & W(i) & " "
        End If
    Next i
    
    QuitarEspacios = Mid$(Nombre, 1, Len(Nombre) - 1)

End Function

Friend Sub RecuperarxSolCred(key As Variant, IdFuente As Long)
    Dim sql As String, i As Long
    Dim rs As Recordset
    
        sql = "SELECT pcs.id,vw.* FROM vwPCProvCli vw "
        sql = sql & " Inner Join PCSolCreProvCli pcs ON pcs.idprovcli= vw.idprovcli "
        sql = sql & " WHERE pcs.tipoSolicitud='" & key & "'"
        sql = sql & " AND pcs.IdTransFuente=" & IdFuente
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.idSolCredito = !Id
        mProps.IdProvcli = !IdProvcli
        mProps.CodProvcli = !CodProvcli
        If Not IsNull(!Nombre) Then mProps.Nombre = !Nombre
        If Not IsNull(!BandProveedor) Then mProps.BandProveedor = !BandProveedor
        If Not IsNull(!BandCliente) Then mProps.BandCliente = !BandCliente
        If Not IsNull(!IdCuentaContable) Then mProps.IdCuentaContable = !IdCuentaContable
        If Not IsNull(!IdCuentaContable2) Then mProps.IdCuentaContable2 = !IdCuentaContable2
        If Not IsNull(!CodCuentaContable) Then mProps.CodCuentaContable = !CodCuentaContable
        If Not IsNull(!CodCuentaContable2) Then mProps.CodCuentaContable2 = !CodCuentaContable2
        If Not IsNull(!Direccion1) Then mProps.Direccion1 = !Direccion1
        If Not IsNull(!Direccion2) Then mProps.Direccion2 = !Direccion2
        If Not IsNull(!CodPostal) Then mProps.CodPostal = !CodPostal
        If Not IsNull(!Ciudad) Then mProps.Ciudad = !Ciudad
        If Not IsNull(!Provincia) Then mProps.Provincia = !Provincia
        If Not IsNull(!Pais) Then
            If Len(!Pais) = 0 Then
                mProps.Pais = "ECUADOR"
            Else
                mProps.Pais = !Pais
            End If
        End If
        If Not IsNull(!Telefono1) Then mProps.Telefono1 = !Telefono1
        If Not IsNull(!Telefono2) Then mProps.Telefono2 = !Telefono2
        If Not IsNull(!Telefono3) Then mProps.Telefono3 = !Telefono3
        If Not IsNull(!Fax) Then mProps.Fax = !Fax
        If Not IsNull(!RUC) Then mProps.RUC = !RUC
        If Not IsNull(!EMail) Then mProps.EMail = !EMail
        If Not IsNull(!LimiteCredito) Then mProps.LimiteCredito = !LimiteCredito
        If Not IsNull(!banco) Then mProps.banco = !banco
        If Not IsNull(!NumCuenta) Then mProps.NumCuenta = !NumCuenta
        If Not IsNull(!Swit) Then mProps.Swit = !Swit
        If Not IsNull(!DirecBanco) Then mProps.DirecBanco = !DirecBanco
        If Not IsNull(!TelBanco) Then mProps.TelBanco = !TelBanco
        If Not IsNull(!IdVendedor) Then mProps.IdVendedor = !IdVendedor
        If Not IsNull(!CodVendedor) Then mProps.CodVendedor = !CodVendedor
        If Not IsNull(!IdGrupo1) Then mProps.IdGrupo1 = !IdGrupo1
        If Not IsNull(!CodGrupo1) Then mProps.CodGrupo1 = !CodGrupo1
        If Not IsNull(!IdGrupo2) Then mProps.IdGrupo2 = !IdGrupo2
        If Not IsNull(!CodGrupo2) Then mProps.CodGrupo2 = !CodGrupo2
        
        If Not IsNull(!IdGrupo3) Then mProps.IdGrupo3 = !IdGrupo3
        If Not IsNull(!CodGrupo3) Then mProps.CodGrupo3 = !CodGrupo3
        
        If Not IsNull(!IdGrupo4) Then mProps.IdGrupo4 = !IdGrupo4 'Agregadoa AUC 22/09/2005
        If Not IsNull(!CodGrupo4) Then mProps.CodGrupo4 = !CodGrupo4
        
        If Not IsNull(!Estado) Then mProps.Estado = !Estado
        If Not IsNull(!TipoDocumento) Then mProps.TipoDocumento = !TipoDocumento          'Agregado Alex Sept/2002
        If Not IsNull(!TipoComprobante) Then mProps.TipoComprobante = !TipoComprobante           'Agregado Alex Sept/2002
        If Not IsNull(!NumAutSRI) Then mProps.NumAutSRI = !NumAutSRI             'Agregado Alex Sept/2002
        If Not IsNull(!NombreAlterno) Then mProps.NombreAlterno = !NombreAlterno          '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!FechaNacimiento) Then mProps.FechaNacimiento = !FechaNacimiento    '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!FechaEntrega) Then mProps.FechaEntrega = !FechaEntrega             '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!FechaExpiracion) Then mProps.FechaExpiracion = !FechaExpiracion    '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!TotalDebe) Then mProps.TotalDebe = !TotalDebe                      '***Agregado. 08/sep/2003. Angel
        If Not IsNull(!TotalHaber) Then mProps.TotalHaber = !TotalHaber                   '***Agregado. 08/sep/2003. Angel
        If Not IsNull(!Observacion) Then mProps.Observacion = !Observacion      'Agregado AUC 31/08/2005
        If Not IsNull(!TipoProvCli) Then mProps.TipoProvCli = Trim$(!TipoProvCli)  ' jeaa 17/12/2005
        If Not IsNull(!ReferidoPor) Then mProps.ReferidoPor = !ReferidoPor      'Agregado jeaa 12/08/2006
        'jeaa 15/05/2007
        If Not IsNull(!IdTipoDocumento) Then mProps.IdTipoDocumento = !IdTipoDocumento 'Agregado jeaa 15/05/2007
        If Not IsNull(!codtipoDocumento) Then mProps.codtipoDocumento = !codtipoDocumento 'Agregado jeaa 15/05/2007
        If Not IsNull(!BandEmpresaPublica) Then mProps.BandEmpresaPublica = !BandEmpresaPublica   'jeaa 16/01/2008 cambio por EQUIDAD TRIBUTARIA
        If Not IsNull(!IdGarante) Then mProps.IdGarante = !IdGarante
        If Not IsNull(!CodGarante) Then mProps.CodGarante = !CodGarante
        If Not IsNull(!BandGarante) Then mProps.BandGarante = !BandGarante
        If Not IsNull(!DiasPlazo) Then mProps.DiasPlazo = !DiasPlazo
        If Not IsNull(!CodigoUsuario) Then mProps.CodigoUsuario = !CodigoUsuario
        If Not IsNull(!IdProvincia) Then mProps.IdProvincia = !IdProvincia
        If Not IsNull(!CodProvincia) Then mProps.CodProvincia = !CodProvincia
        If Not IsNull(!bandEmpleado) Then mProps.bandEmpleado = !bandEmpleado
        If Not IsNull(!IdCanton) Then mProps.IdCanton = !IdCanton
        If Not IsNull(!CodCanton) Then mProps.CodCanton = !CodCanton
        If Not IsNull(!IDParroquia) Then mProps.IDParroquia = !IDParroquia
        If Not IsNull(!CodParroquia) Then mProps.CodParroqui = !CodParroquia
        If Not IsNull(!FechaCreacion) Then mProps.FechaCreacion = !FechaCreacion
        
        If Not IsNull(!FechaLote1) Then mProps.FechaLote1 = !FechaLote1
        If Not IsNull(!FechaLote2) Then mProps.FechaLote2 = !FechaLote2
        If Not IsNull(!BandLote) Then mProps.BandLote = !BandLote
        If Not IsNull(!bandEmpleado) Then mProps.bandEmpleado = !bandEmpleado
        If Not IsNull(!NumSerie) Then mProps.NumSerie = !NumSerie
        If Not IsNull(!NumPunto) Then mProps.NumPunto = !NumPunto
        If Not IsNull(!FechaCaducidad) Then mProps.FechaCaducidad = !FechaCaducidad
        If Not IsNull(!IdDiasCredito) Then mProps.IdDiasCredito = !IdDiasCredito
        If Not IsNull(!CodDiasCredito) Then mProps.CodDiasCredito = !CodDiasCredito
        If Not IsNull(!NumPagos) Then mProps.NumPagos = !NumPagos
        If Not IsNull(!Intervalo) Then mProps.Intervalo = !Intervalo
        If Not IsNull(!PorDescNeto) Then mProps.PorDescNeto = !PorDescNeto
        If Not IsNull(!PorDescPP) Then mProps.PorDescPP = !PorDescPP
        If Not IsNull(!BandRUCValido) Then mProps.BandRUCValido = !BandRUCValido
         If Not IsNull(!IdTurno) Then mProps.IdTurno = !IdTurno
        If Not IsNull(!Codturno) Then mProps.Codturno = !Codturno
        If Not IsNull(!TipoTurno) Then mProps.TipoTurno = !TipoTurno
        If Not IsNull(!BandRelacionado) Then mProps.BandRelacionado = !BandRelacionado
        If Not IsNull(!BandOmitirPendientes) Then mProps.BandOmitirPendientes = !BandOmitirPendientes
        
        If Not IsNull(!TipoSujeto) Then mProps.TipoSujeto = !TipoSujeto
        If Not IsNull(!Sexo) Then mProps.Sexo = !Sexo
        If Not IsNull(!EstadoCivil) Then mProps.EstadoCivil = !EstadoCivil
        If Not IsNull(!OrigenIngresos) Then mProps.OrigenIngresos = !OrigenIngresos
        If Not IsNull(!BandSolicitudCredito) Then mProps.BandSolicitudCredito = !BandSolicitudCredito
         If Not IsNull(!DescripcionDetalle) Then
            If .Fields("DescripcionDetalle").ActualSize > 0 Then
                mProps.DescripcionDetalle = .Fields("DescripcionDetalle").GetChunk(.Fields("DescripcionDetalle").ActualSize)
            End If
        End If
        If Not IsNull(!NombreConyuge) Then mProps.NombreConyuge = !NombreConyuge
        If Not IsNull(!RucConyuge) Then mProps.RucConyuge = !RucConyuge
        mProps.fechagrabado = !fechagrabado
        
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

    
Public Property Let idSolCredito(ByVal value As Long)
    mProps.idSolCredito = value
    'mbooModificado = True
End Property
Public Property Get idSolCredito() As Long
    idSolCredito = mProps.idSolCredito
End Property
Public Property Get CountAgencia() As Long
    CountAgencia = mcolAgencia.Count
End Property
Public Property Get Agencia(ix As Long) As PCAgencia
    Set Agencia = mcolAgencia.item(ix)
End Property
Public Function AddAgencia() As Long
    Dim nuevo As PCAgencia
    Set nuevo = New PCAgencia
    Set nuevo.PCProvCli = Me
    mcolAgencia.Add nuevo
    AddAgencia = mcolAgencia.Count     'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function
Public Function RemoveAgencia(ix As Variant) As Boolean
    Dim obj As PCAgencia, i As Long
    Dim sql As String
    Dim rs As Recordset
    Set obj = Me.Empresa.RecuperaPCAgencia(ix)
    'Si llama con numero de indice del objeto
'    If (VarType(ix) = vbLong) Or (VarType(ix) = vbInteger) Then
'        mcolAgencia.Remove ix      'Elimina el objeto
'    'Si llama con la referencia al objeto que quiere eliminar
'    Else
        'Busca el objeto dentro de la colección
       ' For i = mcolAgencia.Count To 1 Step -1
        '    Set obj = mcolAgencia.item(i)
         '   If obj Is ix Then       'Si encuentra
          '      If obj.IdAgencia > 0 Then
            If Not obj Is Nothing Then
                    sql = "Select g.codtrans,g.numtrans from gncomprobante g inner join gnoferta gof on gof.transid = g.transid where gof.idagencia = " & ix
                    Set rs = Me.Empresa.OpenRecordset(sql)
                    If rs.RecordCount > 0 Then
                         MsgBox "No se puede eliminar tiene trans relacionadas." & rs!CodTrans & rs!NumTrans
                        RemoveAgencia = False
                        Exit Function
                    End If
                'End If
                'mcolAgencia.Remove i       'Elimina el objeto
                '
                'gobjMain.EmpresaActual.EjecutarSQL sql, 1
                obj.EliminaFisica
                GoTo eliminado
            End If
        'Next i
        'Si no encuentra el objeto, genera error
        Err.Raise ERR_INVALIDO, "PCProvCli_RemoveAgencia", "No se puede eliminar "
    'End If
eliminado:
    RemoveAgencia = True
    Set obj = Nothing
End Function

Private Sub AgregaPersonal()
    Dim sql As String
    sql = "INSERT INTO Personal (IdEmpleado,BandActivo,Sexo,EstadoCivil,numcargas,Salario,TipoSalario,varios1,varios2,varios3,varios4,varios5,varios6,varios7,varios8,pagarSeguro,bandfr,bandpagarhe,contador,FechaIngreso,fechapagoseguro) " & _
              "VALUES(" & mProps.IdProvcli & ",1,1,1,0,0,1,0,0,0,0,0,0,0,0,1,0,0,0,'" & Date & "','" & Date & "')"
        mProps.objEmpresa.Execute sql, True
End Sub

Public Property Let Password(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_Password", MSGERR_INVALIDO
    End If
    
    mProps.Password = value
    mbooModificado = True
End Property

Public Property Get Password() As String
    Password = Trim$(mProps.Password)
End Property
Public Property Let CodigoRelojP(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_CodigoRelojP", MSGERR_INVALIDO
    End If
    
    mProps.CodigoRelojP = value
    mbooModificado = True
End Property

Public Property Get CodigoRelojP() As String
    CodigoRelojP = Trim$(mProps.CodigoRelojP)
End Property

Friend Sub RecuperarxCodRelojProv(key As Variant)
    Dim sql As String, i As Long
    Dim rs As Recordset
        
    sql = "SELECT * FROM vwPCProvCli WHERE CodigoReloj='" & key & "'"
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IdProvcli = !IdProvcli
        mProps.CodProvcli = !CodProvcli
        If Not IsNull(!Nombre) Then mProps.Nombre = !Nombre
        If Not IsNull(!BandProveedor) Then mProps.BandProveedor = !BandProveedor
        If Not IsNull(!BandCliente) Then mProps.BandCliente = !BandCliente
        If Not IsNull(!CodigoReloj) Then mProps.CodigoReloj = !CodigoReloj
        If Not IsNull(!Password) Then mProps.Password = !Password
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Property Get NumItemsDetalle() As Long
    If mcolItemsDetalle Is Nothing Then
        NumItemsDetalle = 0
    Else
        NumItemsDetalle = mcolItemsDetalle.Count
    End If
End Property

Public Function RecuperaItemsDetalle(ix As Integer) As PCItemsDetalle
    Set RecuperaItemsDetalle = mcolItemsDetalle.item(ix)
End Function

Public Sub RemoveItemsDetalle(ByVal ix As Long, Optional obj As PCItemsDetalle)
    If mcolItemsDetalle Is Nothing Then CargaItemsDetalle
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolItemsDetalle.Count
            If mcolItemsDetalle.item(ix) Is obj Then
                mcolItemsDetalle.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolItemsDetalle.Remove ix
    End If
    'mcolDetalleFamilia.Remove ix
    mbooModificado = True
End Sub



Private Sub CargaItemsDetalle()
    Dim rs As Recordset, sql As String, obj As PCItemsDetalle
    Dim i As Long
    Set mcolItemsDetalle = New Collection
    sql = " Select iv.IdInventario, iv.CodInventario, iv.Descripcion,"
    sql = sql & " ivp.Cantidad  , ivp.CostoUnitario as PrecioUnitario, ivp.fecha, ivp.BandPrincipal, ivp.PrecioCompra"
    sql = sql & " From Ivinventario  iv inner join  ivproveedordetalle ivp on ivp.idinventario= iv.idinventario"
    sql = sql & " Where ivp.Idproveedor = " & Me.IdProvcli
           
    sql = sql & " order by ivp.fecha"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddItemsDetalle
            Set obj = mcolItemsDetalle.item(i)
            obj.IdInventario = !IdInventario
            obj.CodInventario = !CodInventario
    
            If Not IsNull(!Descripcion) Then obj.Descripcion = !Descripcion
            If Not IsNull(!Cantidad) Then obj.Cantidad = !Cantidad
            If Not IsNull(!PrecioUnitario) Then obj.PrecioUnitario = !PrecioUnitario
            If Not IsNull(!fecha) Then obj.fecha = !fecha
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

'AUC
Public Property Get CountItems() As Long
    CountItems = mcolItemsDetalle.Count
End Property

Public Property Get Items(ix As Integer) As PCItemsDetalle
    Set Items = mcolItemsDetalle.item(ix)
End Property


Public Function AddItemsDetalle() As Long
    Dim nuevo As PCItemsDetalle
    
    If mcolItemsDetalle Is Nothing Then CargaItemsDetalle
    
    Set nuevo = New PCItemsDetalle
    Set nuevo.PCProvCli = Me
    mcolItemsDetalle.Add nuevo
    AddItemsDetalle = mcolItemsDetalle.Count      'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function


Private Sub GrabarItemsDetalle(ByRef cod As String)
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj  As PCItemsDetalle, sql As String, band As Boolean
    'Elimina  todos  los registros
    If mcolItemsDetalle Is Nothing Then Exit Sub
    sql = "Delete From ivproveedordetalle where Idproveedor =" & RecuperaID(cod)
    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    band = True
    If Not mcolItemsDetalle Is Nothing Then
        sql = "SELECT * FROM  ivproveedordetalle WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        For i = 1 To mcolItemsDetalle.Count
            Set obj = mcolItemsDetalle.item(i)
            If obj.IdInventario = 0 Then
                        obj.Recuperar obj.CodInventario
                        If verificacionDuiplicadoItems(obj.IdInventario) Then
                            RemoveItemsDetalle i
                            band = False
                        End If
            End If
            If band Then
                    rs1.AddNew
                    rs1!idProveedor = RecuperaID(cod)
                    rs1!IdInventario = obj.IdInventario
                    rs1!Cantidad = obj.Cantidad
                    rs1!CostoUnitario = obj.PrecioUnitario
                    rs1!fecha = obj.fecha
                rs1.Update
            End If
            band = True
        Next i
    End If
    Set rs1 = Nothing
    Set obj = Nothing
End Sub

Public Function RecuperaID(ByRef cod As String) As Long
    Dim sql As String, rs As Recordset

        sql = "SELECT idprovcli FROM pcprovcli WHERE Codprovcli='" & cod & "'"
    Set rs = Me.Empresa.OpenRecordset(sql)
    With rs
     If rs.RecordCount = 0 Then
        RecuperaID = 0
      Else
        RecuperaID = !IdProvcli
        Exit Function
        .Close
      End If
    End With
    Set rs = Nothing
End Function

Private Function verificacionDuiplicadoItems(ByVal idItem As String) As Boolean
    Dim rs As Recordset, sql As String
    verificacionDuiplicadoItems = False
    'Recupera los detalles de existencia
    sql = "SELECT IdInventario FROM ivproveedordetalle where idInventario=" & idItem
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        verificacionDuiplicadoItems = True
    End If
    Set rs = Nothing
End Function



Public Property Let BandOmitirDinardap(ByVal value As Boolean)
    If value <> mProps.BandOmitirDinardap Then
        mProps.BandOmitirDinardap = value
        mbooModificado = True
    End If
End Property

Public Property Get BandOmitirDinardap() As Boolean
    BandOmitirDinardap = mProps.BandOmitirDinardap
End Property

Public Sub GrabarTransporte()
    Dim sql As String, fgrabado As Date, i As Long
    Dim rs As Recordset
    Dim rsSuc As Recordset, gnSuc As GNSucursal
    Dim NumGrupo As Integer, ivg As ivgrupo, campo As String, k As Long, n As Byte
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
       
    'Verifica si no repite el código
'    sql = "SELECT CodProvCli FROM PcProvcli " & _
'          "WHERE (CodProvCli='" & mProps.CodProvCli & "') AND " & _
'                "(IdProvCli<>" & mProps.IdProvCli & ")"
'    Set rs = mProps.objEmpresa.OpenRecordset(sql)
'    If Not rs.EOF Then      'Si ya existe el codigo
'        Err.Raise ERR_REPITECODIGO, "Empleado", MSGERR_REPITECODIGO
'    End If
'    rs.Close
    On Error GoTo ErrTrap

#If DAOLIB Then
    gobjWS.BeginTrans                           '======= Comienza la transacción
#Else
    mProps.objEmpresa.Coneccion.BeginTrans      '======= Comienza la transacción
#End If

    sql = "SELECT * FROM PCTransporte WHERE idprovcli= " & mProps.IdProvcli
    Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)


    'Si es nuevo
    If rs.RecordCount = 0 Then
        sql = "SELECT * FROM PCTransporte WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM PCTransporte WHERE IdProvCli=" & mProps.IdProvcli
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
'        If rs!FechaGrabado <> mProps.FechaGrabado Then
'            Err.Raise ERR_MODIFICADO, "Empleado_Grabar", MSGERR_MODIFICADO
'        End If
    End If

    'Graba el registro en la base
    With rs
        !IdProvcli = mProps.IdProvcli 'Es el mismo asignado a PcProvCli
        !idpadre = mProps.idpadre
        !IdMadre = mProps.IdMadre
        !IdOtro = mProps.IdOtro
        .Fields("DireccionAM").AppendChunk mProps.DireccionAM
        .Fields("DireccionPM").AppendChunk mProps.DireccionPM
        !BandMismaDireccion = mProps.BandMismaDireccion
        !NumRepresentante = mProps.NumRepresentante
        !NumEmiteFactura = mProps.NumEmiteFactura
        !IdInventario = mProps.IdInventario
        !IdVehiculoAM = mProps.IdVehiculoAM
        !IdVehiculoPM = mProps.IdVehiculoPM
        !BandTransporte = 1

        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
'        fgrabado = !FechaGrabado
        mProps.IdProvcli = !IdProvcli
        
'        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    
    
#If DAOLIB Then
    gobjWS.CommitTrans                  '======== Completa la transacción
#Else
    mProps.objEmpresa.Coneccion.CommitTrans      '======== Completa la transacción
#End If
'    mProps.FechaGrabado = fgrabado
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_TR", "Empleado Nuevo: " & _
                                                         mProps.CodProvcli, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_TR", "Empleado Modificado: " & _
                                                         mProps.CodProvcli, "DT"
    End If
    
    Set gnSuc = Nothing
    Set rsSuc = Nothing


    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PROVCLI
    Exit Sub
ErrTrap:
#If DAOLIB Then
    gobjWS.Rollback                   '======== Anula la transacción
#Else
    mProps.objEmpresa.Coneccion.RollbackTrans  '======== Anula la transacción
#End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub


  Public Sub RecuperarTransporte(IdProvcli As Long)
    Dim sql As String, i As Long
    Dim rs As Recordset
    
    sql = "SELECT PCT.*, pcp.codprovcli as Codpadre , pcm.codprovcli as CodMadre, pco.codprovcli as CodOtro, "
    sql = sql & " CodInventario, gnvAM.codvehiculo as CodVehiculoAM , gnvPM.codvehiculo as CodVehiculoPM "
    sql = sql & " FROM PCTransporte pct "
    sql = sql & " left JOIN Pcprovcli pcp ON  pct.idPadre = pcp.idprovcli "
    sql = sql & " left JOIN Pcprovcli pcm ON  pct.idMadre = pcm.idprovcli "
    sql = sql & " left JOIN Pcprovcli pco ON  pct.idOtro = pco.idprovcli "
    sql = sql & " left JOIN IvInventario iv ON  pct.idInventario = iv.idinventario "
    sql = sql & " left JOIN GnVehiculo gnvAM ON  pct.idvehiculoAM = gnvAM.idvehiculo "
    sql = sql & " left JOIN GnVehiculo gnvPM ON  pct.idvehiculoPM = gnvpM.idvehiculo "
    sql = sql & " Where  pct.idProvcli = " & IdProvcli
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If rs.EOF Then Exit Sub
        With rs
            If Not IsNull(!idpadre) Then mProps.idpadre = !idpadre
            If Not IsNull(!CodPadre) Then mProps.CodPadre = !CodPadre
            If Not IsNull(!IdMadre) Then mProps.IdMadre = !IdMadre
            If Not IsNull(!CodMadre) Then mProps.CodMadre = !CodMadre
            If Not IsNull(!IdOtro) Then mProps.IdOtro = !IdOtro
            If Not IsNull(!CodOtro) Then mProps.CodOtro = !CodOtro
            If Not IsNull(!DireccionAM) Then
                    If .Fields("DireccionAM").ActualSize > 0 Then
                    mProps.DireccionAM = .Fields("DireccionAM").GetChunk(.Fields("DireccionAM").ActualSize)
                End If
            End If
            If Not IsNull(!DireccionPM) Then
                If .Fields("DireccionPM").ActualSize > 0 Then
                    mProps.DireccionPM = .Fields("DireccionPM").GetChunk(.Fields("DireccionPM").ActualSize)
                End If
            End If
            
            If Not IsNull(!BandMismaDireccion) Then mProps.BandMismaDireccion = !BandMismaDireccion
            If Not IsNull(!NumRepresentante) Then mProps.NumRepresentante = !NumRepresentante
            If Not IsNull(!NumEmiteFactura) Then mProps.NumEmiteFactura = !NumEmiteFactura
            If Not IsNull(!IdInventario) Then mProps.IdInventario = !IdInventario
            If Not IsNull(!CodInventario) Then mProps.CodInventario = !CodInventario
            If Not IsNull(!IdVehiculoAM) Then mProps.IdVehiculoAM = !IdVehiculoAM
            If Not IsNull(!CodVehiculoAM) Then mProps.CodVehiculoAM = !CodVehiculoAM
            If Not IsNull(!IdVehiculoPM) Then mProps.IdVehiculoPM = !IdVehiculoPM
            If Not IsNull(!CodVehiculoPM) Then mProps.CodVehiculoPM = !CodVehiculoPM
            If Not IsNull(!BandTransporte) Then mProps.BandTransporte = !BandTransporte
            If Not IsNull(!fechagrabado) Then mProps.FechaInscripcion = !fechagrabado
            
            
            
    
'            mProps.FechaGrabado = !FechaGrabado
            .Close
        End With
    
        
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Property Let BandMismaDireccion(ByVal value As Boolean)
    If value <> mProps.BandMismaDireccion Then
        mProps.BandMismaDireccion = value
        mbooModificado = True
    End If
End Property

Public Property Get BandMismaDireccion() As Boolean
    BandMismaDireccion = mProps.BandMismaDireccion
End Property

Public Property Let idpadre(ByVal value As Long)
    mProps.idpadre = value
    mbooModificado = True
End Property

Public Property Get idpadre() As Long
    idpadre = mProps.idpadre
End Property


Public Property Let CodPadre(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodPadre Then
        If Len(value) > 0 Then
            'Actualiza IdPadre también
            sql = "SELECT IdProvcli FROM Pcprovcli WHERE CodProvcli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli.CodPadre", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                idpadre = rs!IdProvcli
            End If
            rs.Close
            Set rs = Nothing
        Else
            idpadre = 0
        End If
    End If

    mProps.CodPadre = value
    mbooModificado = True
End Property

Public Property Get CodPadre() As String
    CodPadre = Trim$(mProps.CodPadre)
End Property

Public Property Let IdMadre(ByVal value As Long)
    mProps.IdMadre = value
    mbooModificado = True
End Property

Public Property Get IdMadre() As Long
    IdMadre = mProps.IdMadre
End Property


Public Property Let CodMadre(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodMadre Then
        If Len(value) > 0 Then
            'Actualiza IdMadre también
            sql = "SELECT IdProvcli FROM Pcprovcli WHERE CodProvcli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli.CodMadre", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdMadre = rs!IdProvcli
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdMadre = 0
        End If
    End If

    mProps.CodMadre = value
    mbooModificado = True
End Property

Public Property Get CodMadre() As String
    CodMadre = Trim$(mProps.CodMadre)
End Property

Public Property Let IdOtro(ByVal value As Long)
    mProps.IdOtro = value
    mbooModificado = True
End Property

Public Property Get IdOtro() As Long
    IdOtro = mProps.IdOtro
End Property


Public Property Let CodOtro(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodOtro Then
        If Len(value) > 0 Then
            'Actualiza IdOtro también
            sql = "SELECT IdProvcli FROM Pcprovcli WHERE CodProvcli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli.CodOtro", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdOtro = rs!IdProvcli
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdOtro = 0
        End If
    End If

    mProps.CodOtro = value
    mbooModificado = True
End Property

Public Property Get CodOtro() As String
    CodOtro = Trim$(mProps.CodOtro)
End Property

Public Property Let IdInventario(ByVal value As Long)
    mProps.IdInventario = value
    mbooModificado = True
End Property

Public Property Get IdInventario() As Long
    IdInventario = mProps.IdInventario
End Property


Public Property Let CodInventario(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodInventario Then
        If Len(value) > 0 Then
            'Actualiza IdInventario también
            sql = "SELECT IdInventario FROM IvInventario WHERE CodInventario='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVInventario.CodInventario", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdInventario = rs!IdInventario
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdInventario = 0
        End If
    End If

    mProps.CodInventario = value
    mbooModificado = True
End Property

Public Property Get CodInventario() As String
    CodInventario = Trim$(mProps.CodInventario)
End Property

Public Property Let IdVehiculoAM(ByVal value As Long)
    mProps.IdVehiculoAM = value
    mbooModificado = True
End Property

Public Property Get IdVehiculoAM() As Long
    IdVehiculoAM = mProps.IdVehiculoAM
End Property


Public Property Let CodVehiculoAM(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodVehiculoAM Then
        If Len(value) > 0 Then
            'Actualiza IdVehiculoAM también
            sql = "SELECT IdVehiculo FROM GNVehiculo WHERE CodVehiculo='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNVehiculo.CodVehiculoAM", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdVehiculoAM = rs!IdVehiculo
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdVehiculoAM = 0
        End If
    End If

    mProps.CodVehiculoAM = value
    mbooModificado = True
End Property

Public Property Get CodVehiculoAM() As String
    CodVehiculoAM = Trim$(mProps.CodVehiculoAM)
End Property

Public Property Let IdVehiculoPM(ByVal value As Long)
    mProps.IdVehiculoPM = value
    mbooModificado = True
End Property

Public Property Get IdVehiculoPM() As Long
    IdVehiculoPM = mProps.IdVehiculoPM
End Property


Public Property Let CodVehiculoPM(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodVehiculoPM Then
        If Len(value) > 0 Then
            'Actualiza IdVehiculoPM también
            sql = "SELECT IdVehiculo FROM gnVehiculo WHERE CodVehiculo='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNVehiculo.CodVehiculoPM", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdVehiculoPM = rs!IdVehiculo
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdVehiculoPM = 0
        End If
    End If

    mProps.CodVehiculoPM = value
    mbooModificado = True
End Property

Public Property Get CodVehiculoPM() As String
    CodVehiculoPM = Trim$(mProps.CodVehiculoPM)
End Property

Public Property Let NumRepresentante(ByVal value As Integer)
    mProps.NumRepresentante = value
    mbooModificado = True
End Property
Public Property Get NumRepresentante() As Integer
    NumRepresentante = Trim$(mProps.NumRepresentante)
End Property

Public Property Let NumEmiteFactura(ByVal value As Integer)
    mProps.NumEmiteFactura = value
    mbooModificado = True
End Property
Public Property Get NumEmiteFactura() As Integer
    NumEmiteFactura = Trim$(mProps.NumEmiteFactura)
End Property

Public Property Let DireccionAM(ByVal value As String)
    mProps.DireccionAM = value
    mbooModificado = True
End Property

Public Property Get DireccionAM() As String
    DireccionAM = Trim$(mProps.DireccionAM)
End Property

Public Property Let DireccionPM(ByVal value As String)
    mProps.DireccionPM = value
    mbooModificado = True
End Property

Public Property Get DireccionPM() As String
    DireccionPM = Trim$(mProps.DireccionPM)
End Property

Public Property Let BandTransporte(ByVal value As Boolean)
    If value <> mProps.BandTransporte Then
        mProps.BandTransporte = value
        mbooModificado = True
    End If
End Property

Public Property Get BandTransporte() As Boolean
    BandTransporte = mProps.BandTransporte
End Property

Public Property Let FechaInscripcion(ByVal value As Date)
        mProps.FechaInscripcion = value
        mbooModificado = True
    
End Property

Public Property Get FechaInscripcion() As Date
    FechaInscripcion = mProps.FechaInscripcion
End Property



Public Property Let IdCobrador(ByVal value As Long)
    mProps.IdCobrador = value
    mbooModificado = True
End Property
Public Property Get IdCobrador() As Long
    IdCobrador = mProps.IdCobrador
End Property
Public Property Let CodCobrador(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodCobrador Then
        If Len(value) > 0 Then
            'Actualiza IdVendedor también
            sql = "SELECT IdVendedor FROM FCVendedor WHERE CodVendedor='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodCobrador", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCobrador = rs!IdVendedor
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdCobrador = 0
        End If
    End If
    mProps.CodCobrador = value
    mbooModificado = True
End Property
Public Property Get CodCobrador() As String
    CodCobrador = Trim$(mProps.CodCobrador)
End Property
Friend Sub RecuperarxCobrador(key As String)
    Dim sql As String, i As Long
    Dim rs As Recordset
    sql = "SELECT * FROM vwPCProvCli WHERE CodCobrador='" & key & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'Recupera los datos de Usuario
    With rs
        mProps.IdProvcli = !IdProvcli
        mProps.CodProvcli = !CodProvcli
        If Not IsNull(!Nombre) Then mProps.Nombre = !Nombre
        If Not IsNull(!BandProveedor) Then mProps.BandProveedor = !BandProveedor
        If Not IsNull(!BandCliente) Then mProps.BandCliente = !BandCliente
        If Not IsNull(!IdCuentaContable) Then mProps.IdCuentaContable = !IdCuentaContable
        If Not IsNull(!IdCuentaContable2) Then mProps.IdCuentaContable2 = !IdCuentaContable2
        If Not IsNull(!CodCuentaContable) Then mProps.CodCuentaContable = !CodCuentaContable
        If Not IsNull(!CodCuentaContable2) Then mProps.CodCuentaContable2 = !CodCuentaContable2
        If Not IsNull(!Direccion1) Then mProps.Direccion1 = !Direccion1
        If Not IsNull(!Direccion2) Then mProps.Direccion2 = !Direccion2
        If Not IsNull(!CodPostal) Then mProps.CodPostal = !CodPostal
        If Not IsNull(!Ciudad) Then mProps.Ciudad = !Ciudad
        If Not IsNull(!Provincia) Then mProps.Provincia = !Provincia
        If Not IsNull(!Pais) Then mProps.Pais = !Pais
        If Not IsNull(!Telefono1) Then mProps.Telefono1 = !Telefono1
        If Not IsNull(!Telefono2) Then mProps.Telefono2 = !Telefono2
        If Not IsNull(!Telefono3) Then mProps.Telefono3 = !Telefono3
        If Not IsNull(!Fax) Then mProps.Fax = !Fax
        If Not IsNull(!RUC) Then mProps.RUC = !RUC
        If Not IsNull(!EMail) Then mProps.EMail = !EMail
        If Not IsNull(!LimiteCredito) Then mProps.LimiteCredito = !LimiteCredito
        If Not IsNull(!banco) Then mProps.banco = !banco
        If Not IsNull(!NumCuenta) Then mProps.NumCuenta = !NumCuenta
        If Not IsNull(!Swit) Then mProps.Swit = !Swit
        If Not IsNull(!DirecBanco) Then mProps.DirecBanco = !DirecBanco
        If Not IsNull(!TelBanco) Then mProps.TelBanco = !TelBanco
        If Not IsNull(!IdVendedor) Then mProps.IdVendedor = !IdVendedor
        If Not IsNull(!CodVendedor) Then mProps.CodVendedor = !CodVendedor
        If Not IsNull(!IdGrupo1) Then mProps.IdGrupo1 = !IdGrupo1
        If Not IsNull(!CodGrupo1) Then mProps.CodGrupo1 = !CodGrupo1
        If Not IsNull(!IdGrupo2) Then mProps.IdGrupo2 = !IdGrupo2
        If Not IsNull(!CodGrupo2) Then mProps.CodGrupo2 = !CodGrupo2
        If Not IsNull(!IdGrupo3) Then mProps.IdGrupo3 = !IdGrupo3
        If Not IsNull(!CodGrupo3) Then mProps.CodGrupo3 = !CodGrupo3
        If Not IsNull(!IdGrupo4) Then mProps.IdGrupo4 = !IdGrupo4 'Agregadoa AUC 22/09/2005
        If Not IsNull(!CodGrupo4) Then mProps.CodGrupo4 = !CodGrupo4
        If Not IsNull(!Estado) Then mProps.Estado = !Estado
        If Not IsNull(!TipoDocumento) Then mProps.TipoDocumento = !TipoDocumento          'Agregado Alex Sept/2002
        If Not IsNull(!TipoComprobante) Then mProps.TipoComprobante = !TipoComprobante           'Agregado Alex Sept/2002
        If Not IsNull(!NumAutSRI) Then mProps.NumAutSRI = !NumAutSRI             'Agregado Alex Sept/2002
        If Not IsNull(!NombreAlterno) Then mProps.NombreAlterno = !NombreAlterno          '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!FechaNacimiento) Then mProps.FechaNacimiento = !FechaNacimiento    '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!FechaEntrega) Then mProps.FechaEntrega = !FechaEntrega             '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!FechaExpiracion) Then mProps.FechaExpiracion = !FechaExpiracion    '***Agregado. 05/sep/2003. Angel
        If Not IsNull(!TotalDebe) Then mProps.TotalDebe = !TotalDebe                      '***Agregado. 08/sep/2003. Angel
        If Not IsNull(!TotalHaber) Then mProps.TotalHaber = !TotalHaber                   '***Agregado. 08/sep/2003. Angel
        If Not IsNull(!Observacion) Then mProps.Observacion = !Observacion      'Agregado AUC 31/08/2005
        If Not IsNull(!TipoProvCli) Then mProps.TipoProvCli = Trim$(!TipoProvCli)  ' jeaa 17/12/2005
        If Not IsNull(!ReferidoPor) Then mProps.ReferidoPor = !ReferidoPor      'Agregado jeaa 12/08/2006
        'jeaa 15/05/2007
        If Not IsNull(!IdTipoDocumento) Then mProps.IdTipoDocumento = !IdTipoDocumento 'Agregado jeaa 15/05/2007
        If Not IsNull(!codtipoDocumento) Then mProps.codtipoDocumento = !codtipoDocumento 'Agregado jeaa 15/05/2007
        If Not IsNull(!BandEmpresaPublica) Then mProps.BandEmpresaPublica = !BandEmpresaPublica   'jeaa 16/01/2008 cambio por EQUIDAD TRIBUTARIA
        If Not IsNull(!IdGarante) Then mProps.IdGarante = !IdGarante
        If Not IsNull(!CodGarante) Then mProps.CodGarante = !CodGarante
        If Not IsNull(!BandGarante) Then mProps.BandGarante = !BandGarante
        If Not IsNull(!DiasPlazo) Then mProps.DiasPlazo = !DiasPlazo
        If Not IsNull(!IdProvincia) Then mProps.IdProvincia = !IdProvincia
        If Not IsNull(!IdCanton) Then mProps.IdCanton = !IdCanton
        If Not IsNull(!IDParroquia) Then mProps.IDParroquia = !IDParroquia
        If Not IsNull(!FechaCreacion) Then mProps.FechaCreacion = !FechaCreacion
        If Not IsNull(!CodigoUsuario) Then mProps.CodigoUsuario = !CodigoUsuario
        If Not IsNull(!FechaLote1) Then mProps.CodigoUsuario = !FechaLote1
        If Not IsNull(!FechaLote2) Then mProps.CodigoUsuario = !FechaLote2
        If Not IsNull(!BandLote) Then mProps.CodigoUsuario = !BandLote
        If Not IsNull(!bandEmpleado) Then mProps.bandEmpleado = !bandEmpleado
        If Not IsNull(!NumSerie) Then mProps.NumSerie = !NumSerie
        If Not IsNull(!NumPunto) Then mProps.NumPunto = !NumPunto
        If Not IsNull(!FechaCaducidad) Then mProps.FechaCaducidad = !FechaCaducidad
        If Not IsNull(!IdDiasCredito) Then mProps.IdDiasCredito = !IdDiasCredito
        If Not IsNull(!BandRUCValido) Then mProps.BandRUCValido = !BandRUCValido
        If Not IsNull(!BandRelacionado) Then mProps.BandRelacionado = !BandRelacionado
        If Not IsNull(!BandOmitirPendientes) Then mProps.BandOmitirPendientes = !BandOmitirPendientes
        If Not IsNull(!TipoSujeto) Then mProps.TipoSujeto = !TipoSujeto
        If Not IsNull(!Sexo) Then mProps.Sexo = !Sexo
        If Not IsNull(!EstadoCivil) Then mProps.EstadoCivil = !EstadoCivil
        If Not IsNull(!OrigenIngresos) Then mProps.OrigenIngresos = !OrigenIngresos
        If Not IsNull(!BandOmitirDinardap) Then mProps.BandOmitirDinardap = !BandOmitirDinardap
        If Not IsNull(!IdCobrador) Then mProps.IdCobrador = !IdCobrador
        If Not IsNull(!CodCobrador) Then mProps.CodCobrador = !CodCobrador
        If Not IsNull(!IdPais) Then mProps.IdPais = !IdPais
        If Not IsNull(!CodPais) Then mProps.CodPais = !CodPais
        If Not IsNull(!IDCallePri) Then mProps.IDCallePri = !IDCallePri
        If Not IsNull(!CodCallePri) Then mProps.CodCallePri = !CodCallePri
        If Not IsNull(!IdCalleSec) Then mProps.IdCalleSec = !IdCalleSec
        If Not IsNull(!CodCalleSec) Then mProps.CodCalleSec = !CodCalleSec
        If Not IsNull(!IdCuentaContableGasto) Then mProps.IdCuentaContableGasto = !IdCuentaContableGasto
        If Not IsNull(!CodCuentaContableGasto) Then mProps.CodCuentaContableGasto = !CodCuentaContableGasto
        
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Property Let BandRebate(ByVal value As Boolean)
    If value <> mProps.BandRebate Then
        mProps.BandRebate = value
        AsignarCodCuentaPre             '*** MAKOTO 26/mar/01
    End If
End Property

Public Property Get BandRebate() As Boolean
    BandRebate = mProps.BandRebate
End Property

Public Property Let CodPais(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodProvincia Then
        If Len(value) > 0 Then
            sql = "SELECT IdPais FROM PCPais WHERE CodPais='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodPais", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdPais = rs!IdPais
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdPais = 0
        End If
    End If
    mProps.CodPais = value
    mbooModificado = True
End Property

Public Property Get CodPais() As String
    CodPais = Trim$(mProps.CodPais)
End Property

Public Property Let IdPais(ByVal value As Long)
    mProps.IdPais = value
    mbooModificado = True
End Property

Public Property Get IdPais() As Long
    IdPais = mProps.IdPais
End Property

Public Property Let CordenadaX(ByVal value As Long)
    mProps.CordenadaX = value
    mbooModificado = True
End Property

Public Property Get CordenadaX() As Long
    CordenadaX = mProps.CordenadaX
End Property

Public Property Let CordenadaY(ByVal value As Long)
    mProps.CordenadaY = value
    mbooModificado = True
End Property

Public Property Get posGoogleMaps() As String
    posGoogleMaps = mProps.posGoogleMaps
End Property

Public Property Let posGoogleMaps(ByVal value As String)
    mProps.posGoogleMaps = value
    mbooModificado = True
End Property



Public Property Get CordenadaY() As Long
    CordenadaY = mProps.CordenadaY
End Property


Public Property Let CodCallePri(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodCallePri Then
        If Len(value) > 0 Then
            sql = "SELECT IdCalle FROM PCCalle WHERE CodCalle='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodCallePri", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IDCallePri = rs!IdCalle
            End If
            rs.Close
            Set rs = Nothing
        Else
            IDCallePri = 0
        End If
    End If
    mProps.CodCallePri = value
    mbooModificado = True
End Property

Public Property Get CodCallePri() As String
    CodCallePri = Trim$(mProps.CodCallePri)
End Property

Public Property Let IDCallePri(ByVal value As Long)
    mProps.IDCallePri = value
    mbooModificado = True
End Property

Public Property Get IDCallePri() As Long
    IDCallePri = mProps.IDCallePri
End Property

Public Property Let CodCalleSec(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodCalleSec Then
        If Len(value) > 0 Then
            
            sql = "SELECT IdCalle FROM PCCalle WHERE CodCalle='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodCalleSec", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCalleSec = rs!IdCalle
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdCalleSec = 0
        End If
    End If
    mProps.CodCalleSec = value
    mbooModificado = True
End Property

Public Property Get CodCalleSec() As String
    CodCalleSec = Trim$(mProps.CodCalleSec)
End Property

Public Property Let IdCalleSec(ByVal value As Long)
    mProps.IdCalleSec = value
    mbooModificado = True
End Property

Public Property Get IdCalleSec() As Long
    IdCalleSec = mProps.IdCalleSec
End Property


Public Property Let IdGarante2(ByVal value As Long)
    mProps.IdGarante2 = value
    mbooModificado = True
End Property

Public Property Get IdGarante2() As Long
    IdGarante2 = mProps.IdGarante2
End Property


Public Property Let CodGarante2(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodGarante2 Then
        If Len(value) > 0 Then
            'Actualiza IdGarante2 también
            sql = "SELECT IdProvCli FROM PCProvCli WHERE CodProvCli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodGarante2", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdGarante2 = rs!IdProvcli
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdGarante2 = 0
        End If
    End If

    mProps.CodGarante2 = value
    mbooModificado = True
End Property

Public Property Get CodGarante2() As String
    CodGarante2 = Trim$(mProps.CodGarante2)
End Property

Public Property Get NumCasa() As String
    NumCasa = mProps.NumCasa
End Property

Public Property Let NumCasa(ByVal value As String)
    mProps.NumCasa = value
    mbooModificado = True
End Property

Public Property Let bandCompensaIVA(ByVal value As Boolean)
    If value <> mProps.bandCompensaIVA Then
        mProps.bandCompensaIVA = value
        AsignarCodCuentaPre             '*** MAKOTO 26/mar/01
    End If
End Property

Public Property Get bandCompensaIVA() As Boolean
    bandCompensaIVA = mProps.bandCompensaIVA
End Property

Public Function RemoveDest(ix As Long) As Boolean
    Dim obj As PCAsignado
    Dim sql As String
    Dim rs As Recordset
    Dim msg
    '----------------
    sql = "Select g.codtrans,g.numtrans from gncomprobante g inner join gnoferta gof on gof.transid = g.transid where gof.iddESTINATARIO = " & ix
    Set rs = Me.Empresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        msg = MsgBox("No se puede eliminar tiene trans relacionadas." & rs!CodTrans & rs!NumTrans & Chr(13) & "Desea Eliminar de todas maneras ", vbYesNo)
        If msg = vbNo Then
            RemoveDest = False
            Exit Function
        End If
    End If
    Set rs = Nothing
    '-----------------
    
    Set obj = Me.Empresa.RecuperaPCAsigna(ix)
        If Not obj Is Nothing Then
            obj.EliminaFisica
            GoTo eliminado
        End If
        'Si no encuentra el objeto, genera error
        Err.Raise ERR_INVALIDO, "PCProvCli_RemovePCAsigna", "No se puede eliminar "
    
eliminado:
    RemoveDest = True
    Set obj = Nothing
End Function

Public Function RecuperafechaGrabado(ByVal Id As Long) As Date
    Dim sql As String
    Dim rs As Recordset
    
    sql = "Select fechagrabado from pctransporte where idprovcli= " & Id
    Set rs = Me.Empresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        RecuperafechaGrabado = rs!fechagrabado
    End If
End Function

Public Function RecuperaNumeroItemsFacturarxLoteCC(ByVal Id As Long, CodContrato As String) As Integer
    Dim sql As String
    Dim rs As Recordset
    
    RecuperaNumeroItemsFacturarxLoteCC = 0
    sql = "Select count( IV.IdInventario) as cant From PCProvCliInv Iv "
    sql = sql & " INNER JOIN PCProvCliInv PIv  ON PIv.IdInventario = IV.Idinventario"
    sql = sql & " where piv.idprovcli = " & Id
    sql = sql & " and Iv.referencia = '" & CodContrato & "'"
    Set rs = Me.Empresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        RecuperaNumeroItemsFacturarxLoteCC = rs!cant
    End If
    Set rs = Nothing
End Function

Public Property Let BandEsConsumidorFinal(ByVal value As Boolean)
    If value <> mProps.BandEsConsumidorFinal Then
        mProps.BandEsConsumidorFinal = value
        mbooModificado = True
    End If
End Property

Public Property Get BandEsConsumidorFinal() As Boolean
    BandEsConsumidorFinal = mProps.BandEsConsumidorFinal
End Property

Public Function VerificaCIyRUCRepetido(RUC As String, ByRef rucRepetido As String, ByRef Nombre As String) As Boolean
    Dim rs As Recordset, sql As String
    VerificaCIyRUCRepetido = False
    sql = "SELECT  RUC , nombre  FROM PCProvCli "
    sql = sql & " WHERE IDPROVCLI <> " & Me.IdProvcli & " AND (RUC like '%" & Mid$(RUC, 1, 10) & "%') AND "
    sql = sql & " BandCliente=1"
    sql = sql & " and (IdProvCli<>" & mProps.IdProvcli & ")"
   
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
                rucRepetido = rs.Fields(0)
                Nombre = rs.Fields(1)
                VerificaCIyRUCRepetido = True
    End If
    rs.Close
End Function

Public Property Let BandIVCodigoIgual(ByVal value As Boolean)
    If value <> mProps.BandIVCodigoIgual Then
        mProps.BandIVCodigoIgual = value
        mbooModificado = True
    End If
End Property

Public Property Get BandIVCodigoIgual() As Boolean
    BandIVCodigoIgual = mProps.BandIVCodigoIgual
End Property

Public Property Let IdCuentaContableGasto(ByVal value As Long)
    mProps.IdCuentaContableGasto = value
    mbooModificado = True
End Property

Public Property Get IdCuentaContableGasto() As Long
    IdCuentaContableGasto = mProps.IdCuentaContableGasto
End Property


Public Property Let CodCuentaContableGasto(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCuentaContableGasto Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaContableGasto también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli.CodCuentaContableGasto", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaContableGasto = rs!IdCuenta
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdCuentaContableGasto = 0
        End If
    End If

    mProps.CodCuentaContableGasto = value
    mbooModificado = True
End Property

Public Property Get CodCuentaContableGasto() As String
    CodCuentaContableGasto = Trim$(mProps.CodCuentaContableGasto)
End Property

Public Property Get Garantia() As String
    Garantia = mProps.Garantia
End Property

Public Property Let Garantia(ByVal value As String)
    mProps.Garantia = value
    mbooModificado = True
End Property

Public Property Get Predio() As String
    Predio = mProps.Predio
End Property

Public Property Let Predio(ByVal value As String)
    mProps.Predio = value
    mbooModificado = True
End Property

Public Property Get BandNCPP() As Boolean
    BandNCPP = mProps.BandNCPP
End Property

Public Property Let BandNCPP(ByVal value As Boolean)
    mProps.BandNCPP = value
    mbooModificado = True
End Property

