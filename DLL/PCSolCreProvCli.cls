VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCSolCreProvCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type T_PROPS
    Id As Long
    IdProvCli As Long
    CodProvCli As String
    CodGar As String
    IdGar As Long
    CodEstadoCivil As String
    CodInstruccion As String
    BandPatriFam As Boolean
    BandSepBienes As Boolean
    TipoVivienda As String
    AnioResidencia As String
    NombreArrend As String
    TelfArrend As String
    'SOLICITANTE
    ActividadSoli As String
    NombreTrabajoSoli As String
    DirTrabajoSoli As String
    AntiguedadSoli As String
    TelfActividadSol As String
    TipoLocalSoli As String
    NombreArrendSoli As String
    TelfArrendSoli As String
    IngresosSoli As Currency
    OtrosIngresosSoli As Currency
    GastosFamSoli As Currency
    GastosArriendoSoli As Currency
    OtrosGastosSoli As Currency
    'CONYUGUE
    
    NombreConyu As String
    CIConyu As String
    FechaNacimientoConyu As Date
    NumcargasConyu As String
    ActividadConyu As String
    NombreTrabajoConyu As String
    DirTrabajoConyu  As String
    AntiguedadConyu As String
    TelfActividadCounyu As String
    TipoLocalConyu As String
    NombreArrendConyu As String
    TelfArrendConyu As String
    IngresosConyu As Currency
    ValorVehiculo1Conyu As Currency
    OtrosIngresosConyu As Currency
    GastosFamConyu As Currency
    GastosArriendoConyu As Currency
    OtrosGastosConyu As Currency
    'INMUEBLES
    Inmueble1 As String
    Ubicacion1 As String
    ValorInmueble1 As Currency
    BandHipoteca1 As Boolean
    Inmueble2 As String
    Ubicacion2 As String
    ValorInmueble2 As Currency
    BandHipoteca2 As Boolean
    Inmueble3 As String
    Ubicacion3 As String
    ValorInmueble3 As Currency
    BandHipoteca3 As Boolean
    'VEHICULOS
    Vehiculo1 As String
    MarcaPlaca1 As String
    ValorVehiculo1 As Currency
    BandPrendado1 As Boolean
    Vehiculo2 As String
    MarcaPlaca2 As String
    ValorVehiculo2 As Currency
    BandPrendado2 As Boolean
    Vehiculo3 As String
    MarcaPlaca3 As String
    ValorVehiculo3 As Currency
    BandPrendado3 As Boolean
    'PASIVO
    MontoDeudaBancaria As Currency
    SaldoDeudaBancaria As Currency
    MensualDeudaBancaria As Currency
    MontoHipoteca As Currency
    SaldoHipoteca As Currency
    MensualHipoteca As Currency
    MontoDeudaOtarsCuentas As Currency
    SaldoDeudaOtarsCuentas As Currency
    MensualDeudaOtarsCuentas As Currency
    
    'REF FAMILIAR
    NombreFami As String
    RelacionFami As String
    TelfFami As String
    CiudadFami As String
    
    NombreFami2 As String
    RelacionFami2 As String
    TelfFami2 As String
    CiudadFami2 As String
    
    'REF COMERCIAL
    NombreComer1 As String
    DirComer1 As String
    TelfComer1 As String
    NombreComer2 As String
    DirComer2 As String
    TelfComer2 As String
    NombreComer3 As String
    DirComer3 As String
    TelfComer3 As String
    'REF PERSONAL
    NombrePerso1 As String
    DirPerso1 As String
    TelfPerso1 As String
    NombrePerso2 As String
    DirPerso2 As String
    TelfPerso2 As String
    NombrePerso3 As String
    DirPerso3 As String
    TelfPerso3 As String
    'REF BANCARIA
    NombreBanca1 As String
    NumCtaBanca1 As String
    CodTipoBanca1 As String
    NombreBanca2 As String
    NumCtaBanca2 As String
    CodTipoBanca2 As String
    NombreBanca3 As String
    NumCtaBanca3 As String
    CodTipoBanca3 As String
    'motivo solicitud
    IdInventario As Long
    CodInventario As String
    BandAprobado As Boolean
    FechaSolicitud As Date
    BandCliente As Boolean
    MontoCredito As Currency
    IdTransFuente As Long
    TransFuente As String
    TipoSolicitud As Byte
    IdEmpleado As Long
    CodEmpleado As String
    FechaGrabado As Date
    CodUsuarioAutoriza As String
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mProps As T_PROPS
Private mcolContactos As Collection
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Public Sub Grabar()
    Dim sql As String, fgrabado As Date, i As Long
    Dim rs As Recordset
    Dim rsSuc As Recordset, gnSuc As GNSucursal
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    On Error GoTo ErrTrap
#If DAOLIB Then
    gobjWS.BeginTrans                           '======= Comienza la transacción
#Else
    mProps.objEmpresa.Coneccion.BeginTrans      '======= Comienza la transacción
#End If
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM PCSolCreProvCli WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
    
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM PCSolCreProvCli WHERE Id=" & mProps.Id
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "PCSolCreProvCli_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en la base
    With rs
        !IdProvCli = mProps.IdProvCli
        !CodEstadoCivil = mProps.CodEstadoCivil
        !CodInstruccion = mProps.CodInstruccion
        !BandPatriFam = mProps.BandPatriFam
        !BandSepBienes = mProps.BandSepBienes
        !TipoVivienda = mProps.TipoVivienda
        !AnioResidencia = mProps.AnioResidencia
        !NombreArrend = mProps.NombreArrend
        !TelfArrend = mProps.TelfArrend
        'datos soli
        !ActividadSoli = mProps.ActividadSoli
        !NombreTrabajoSoli = mProps.NombreTrabajoSoli
        !DirTrabajoSoli = mProps.DirTrabajoSoli
        !AntiguedadSoli = mProps.AntiguedadSoli
        !TelfActividadSol = mProps.TelfActividadSol
        !TelfArrend = mProps.TelfArrend
        !TipoLocalSoli = mProps.TipoLocalSoli
        !NombreArrendSoli = mProps.NombreArrendSoli
        !TelfArrendSoli = mProps.TelfArrendSoli
        !IngresosSoli = mProps.IngresosSoli
        !OtrosIngresosSoli = mProps.OtrosIngresosSoli
        !GastosFamSoli = mProps.GastosFamSoli
        !GastosArriendoSoli = mProps.GastosArriendoSoli
        !OtrosGastosSoli = mProps.OtrosGastosSoli
        'conyu
        !NombreConyu = mProps.NombreConyu
        !CIConyu = mProps.CIConyu
        !FechaNacimientoConyu = mProps.FechaNacimientoConyu
        !NumcargasConyu = mProps.NumcargasConyu
        !ActividadConyu = mProps.ActividadConyu
        !NombreTrabajoConyu = mProps.NombreTrabajoConyu
        !DirTrabajoConyu = mProps.DirTrabajoConyu
        !AntiguedadConyu = mProps.AntiguedadConyu
        !TelfActividadCounyu = mProps.TelfActividadCounyu
        !TipoLocalConyu = mProps.TipoLocalConyu
        !NombreArrendConyu = mProps.NombreArrendConyu
        !TelfArrendConyu = mProps.TelfArrendConyu
        
        !IngresosConyu = mProps.IngresosConyu
        !OtrosIngresosConyu = mProps.OtrosIngresosConyu
        !GastosFamConyu = mProps.GastosFamConyu
        !GastosArriendoConyu = mProps.GastosArriendoConyu
        !OtrosGastosConyu = mProps.OtrosGastosConyu
        'inmuebles
        !Inmueble1 = mProps.Inmueble1
        !Ubicacion1 = mProps.Ubicacion1
        !ValorInmueble1 = mProps.ValorInmueble1
        !BandHipoteca1 = mProps.BandHipoteca1
        !Inmueble2 = mProps.Inmueble2
        !Ubicacion2 = mProps.Ubicacion2
        !ValorInmueble2 = mProps.ValorInmueble2
        !BandHipoteca2 = mProps.BandHipoteca2
        !Inmueble3 = mProps.Inmueble3
        !Ubicacion3 = mProps.Ubicacion3
        !ValorInmueble3 = mProps.ValorInmueble3
        !BandHipoteca3 = mProps.BandHipoteca3
        'vehiculos
        !Vehiculo1 = mProps.Vehiculo1
        !MarcaPlaca1 = mProps.MarcaPlaca1
        !ValorVehiculo1 = mProps.ValorVehiculo1
        !BandPrendado1 = mProps.BandPrendado1
        !Vehiculo2 = mProps.Vehiculo2
        !MarcaPlaca2 = mProps.MarcaPlaca2
        !ValorVehiculo2 = mProps.ValorVehiculo2
        !BandPrendado2 = mProps.BandPrendado2
        !Vehiculo3 = mProps.Vehiculo3
        !MarcaPlaca3 = mProps.MarcaPlaca3
        !ValorVehiculo3 = mProps.ValorVehiculo3
        !BandPrendado3 = mProps.BandPrendado3
        'pasivos
        !MontoDeudaBancaria = mProps.MontoDeudaBancaria
        !SaldoDeudaBancaria = mProps.SaldoDeudaBancaria
        !MensualDeudaBancaria = mProps.MensualDeudaBancaria
        !MontoHipoteca = mProps.MontoHipoteca
        !SaldoHipoteca = mProps.SaldoHipoteca
        !MensualHipoteca = mProps.MensualHipoteca
        !MontoDeudaOtarsCuentas = mProps.MontoDeudaOtarsCuentas
        !SaldoDeudaOtarsCuentas = mProps.SaldoDeudaOtarsCuentas
        !MensualDeudaOtarsCuentas = mProps.MensualDeudaOtarsCuentas
        'ref familiar
        !NombreFami = mProps.NombreFami
        !RelacionFami = mProps.RelacionFami
        !CiudadFami = mProps.CiudadFami
        !TelfFami = mProps.TelfFami
        !NombreFami2 = mProps.NombreFami2
        !RelacionFami2 = mProps.RelacionFami2
        !CiudadFami2 = mProps.CiudadFami2
        !TelfFami2 = mProps.TelfFami2
        
        'ref comercial
        !NombreComer1 = mProps.NombreComer1
        !DirComer1 = mProps.DirComer1
        !TelfComer1 = mProps.TelfComer1
        !NombreComer2 = mProps.NombreComer2
        !DirComer2 = mProps.DirComer2
        !TelfComer2 = mProps.TelfComer2
        !NombreComer3 = mProps.NombreComer3
        !DirComer3 = mProps.DirComer3
        !TelfComer3 = mProps.TelfComer3
        'ref Personal
        !NombrePerso1 = mProps.NombrePerso1
        !DirPerso1 = mProps.DirPerso1
        !TelfPerso1 = mProps.TelfPerso1
        !NombrePerso2 = mProps.NombrePerso2
        !DirPerso2 = mProps.DirPerso2
        !TelfPerso2 = mProps.TelfPerso2
        !NombrePerso3 = mProps.NombrePerso3
        !DirPerso3 = mProps.DirPerso3
        !TelfPerso3 = mProps.TelfPerso3
        'ref bancaria
        !NombreBanca1 = mProps.NombreBanca1
        !NumCtaBanca1 = mProps.NumCtaBanca1
        !CodTipoBanca1 = mProps.CodTipoBanca1
        !NombreBanca2 = mProps.NombreBanca2
        !NumCtaBanca2 = mProps.NumCtaBanca2
        !CodTipoBanca2 = mProps.CodTipoBanca2
        !NombreBanca3 = mProps.NombreBanca3
        !NumCtaBanca3 = mProps.NumCtaBanca3
        !CodTipoBanca3 = mProps.CodTipoBanca3
        'motivo
        !IdInventario = mProps.IdInventario
        !BandAprobado = mProps.BandAprobado
        !FechaSolicitud = mProps.FechaSolicitud
        !BandCliente = mProps.BandCliente
        !MontoCredito = mProps.MontoCredito
        !IdTransFuente = mProps.IdTransFuente
        !TipoSolicitud = mProps.TipoSolicitud
        !IdGar = mProps.IdGar
        !IdEmpleado = mProps.IdEmpleado
        !CodUsuarioAutoriza = mProps.CodUsuarioAutoriza
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.Id = !Id             'Guarda el Id(Autonumerico)
        fgrabado = !FechaGrabado
        mProps.IdProvCli = !IdProvCli
        .Close
    End With
    
#If DAOLIB Then
    gobjWS.CommitTrans                  '======== Completa la transacción
#Else
    mProps.objEmpresa.Coneccion.CommitTrans      '======== Completa la transacción
#End If
    mProps.FechaGrabado = fgrabado
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_PC", "Solicitud Credito Prov/Cli Nuevo: " & _
                                                         mProps.CodProvCli, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_PC", "Solicitud Credito Prov/Cli Modificado: " & _
                                                         mProps.CodProvCli, "DT"
    End If
    

    mProps.objEmpresa.RaiseRefreshEvent REFRESH_SOLPROVCLI
    

    mbooEsNuevo = False
    mbooModificado = False
    
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
        sql = "SELECT * FROM vwPCSolCreProvCli WHERE id=" & key
    Else
        sql = "SELECT * FROM vwPCSolCreProvCli WHERE id=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.Id = !Id
        If Not IsNull(!IdProvCli) Then mProps.IdProvCli = !IdProvCli
        If Not IsNull(!CodProvCli) Then mProps.CodProvCli = !CodProvCli
        If Not IsNull(!CodEstadoCivil) Then mProps.CodEstadoCivil = !CodEstadoCivil
        If Not IsNull(!CodInstruccion) Then mProps.CodInstruccion = !CodInstruccion
        If Not IsNull(!BandPatriFam) Then mProps.BandPatriFam = !BandPatriFam
        If Not IsNull(!BandSepBienes) Then mProps.BandSepBienes = !BandSepBienes
        If Not IsNull(!TipoVivienda) Then mProps.TipoVivienda = !TipoVivienda
        If Not IsNull(!AnioResidencia) Then mProps.AnioResidencia = !AnioResidencia
        If Not IsNull(!NombreArrend) Then mProps.NombreArrend = !NombreArrend
        If Not IsNull(!TelfArrend) Then mProps.TelfArrend = !TelfArrend
        'solicitante
        If Not IsNull(!ActividadSoli) Then mProps.ActividadSoli = !ActividadSoli
        If Not IsNull(!NombreTrabajoSoli) Then mProps.NombreTrabajoSoli = !NombreTrabajoSoli
        If Not IsNull(!DirTrabajoSoli) Then mProps.DirTrabajoSoli = !DirTrabajoSoli
        If Not IsNull(!AntiguedadSoli) Then mProps.AntiguedadSoli = !AntiguedadSoli
        If Not IsNull(!TelfActividadSol) Then mProps.TelfActividadSol = !TelfActividadSol
        If Not IsNull(!TipoLocalSoli) Then mProps.TipoLocalSoli = !TipoLocalSoli
        If Not IsNull(!NombreArrendSoli) Then mProps.NombreArrendSoli = !NombreArrendSoli
        If Not IsNull(!TelfArrendSoli) Then mProps.TelfArrendSoli = !TelfArrendSoli
        If Not IsNull(!IngresosSoli) Then mProps.IngresosSoli = !IngresosSoli
        If Not IsNull(!OtrosIngresosSoli) Then mProps.OtrosIngresosSoli = !OtrosIngresosSoli
        If Not IsNull(!GastosFamSoli) Then mProps.GastosFamSoli = !GastosFamSoli
        If Not IsNull(!GastosArriendoSoli) Then mProps.GastosArriendoSoli = !GastosArriendoSoli
        If Not IsNull(!OtrosGastosSoli) Then mProps.OtrosGastosSoli = !OtrosGastosSoli
        'conyuge
        If Not IsNull(!NombreConyu) Then mProps.NombreConyu = !NombreConyu
        If Not IsNull(!CIConyu) Then mProps.CIConyu = !CIConyu
        If Not IsNull(!FechaNacimientoConyu) Then mProps.FechaNacimientoConyu = !FechaNacimientoConyu
        If Not IsNull(!NumcargasConyu) Then mProps.NumcargasConyu = !NumcargasConyu
        If Not IsNull(!ActividadConyu) Then mProps.ActividadConyu = !ActividadConyu
        If Not IsNull(!NombreTrabajoConyu) Then mProps.NombreTrabajoConyu = !NombreTrabajoConyu
        If Not IsNull(!DirTrabajoConyu) Then mProps.DirTrabajoConyu = !DirTrabajoConyu
        If Not IsNull(!AntiguedadConyu) Then mProps.AntiguedadConyu = !AntiguedadConyu
        If Not IsNull(!TelfActividadCounyu) Then mProps.TelfActividadCounyu = !TelfActividadCounyu
        If Not IsNull(!TipoLocalConyu) Then mProps.TipoLocalConyu = !TipoLocalConyu
        If Not IsNull(!NombreArrendConyu) Then mProps.NombreArrendConyu = !NombreArrendConyu
        If Not IsNull(!TelfArrendConyu) Then mProps.TelfArrendConyu = !TelfArrendConyu
        If Not IsNull(!IngresosConyu) Then mProps.IngresosConyu = !IngresosConyu
        If Not IsNull(!OtrosIngresosConyu) Then mProps.OtrosIngresosConyu = !OtrosIngresosConyu
        If Not IsNull(!GastosFamConyu) Then mProps.GastosFamConyu = !GastosFamConyu
        If Not IsNull(!GastosArriendoConyu) Then mProps.GastosArriendoConyu = !GastosArriendoConyu
        If Not IsNull(!OtrosGastosConyu) Then mProps.OtrosGastosConyu = !OtrosGastosConyu
        'inmuebles
        If Not IsNull(!Inmueble1) Then mProps.Inmueble1 = !Inmueble1
        If Not IsNull(!Ubicacion1) Then mProps.Ubicacion1 = !Ubicacion1
        If Not IsNull(!ValorInmueble1) Then mProps.ValorInmueble1 = !ValorInmueble1
        If Not IsNull(!BandHipoteca1) Then mProps.BandHipoteca1 = !BandHipoteca1
        If Not IsNull(!Inmueble2) Then mProps.Inmueble2 = !Inmueble2
        If Not IsNull(!Ubicacion2) Then mProps.Ubicacion2 = !Ubicacion2
        If Not IsNull(!ValorInmueble2) Then mProps.ValorInmueble2 = !ValorInmueble2
        If Not IsNull(!BandHipoteca2) Then mProps.BandHipoteca2 = !BandHipoteca2
        If Not IsNull(!Inmueble3) Then mProps.Inmueble3 = !Inmueble3
        If Not IsNull(!Ubicacion3) Then mProps.Ubicacion3 = !Ubicacion3
        If Not IsNull(!ValorInmueble3) Then mProps.ValorInmueble3 = !ValorInmueble3
        If Not IsNull(!BandHipoteca3) Then mProps.BandHipoteca3 = !BandHipoteca3
        
        'vehiculo
        If Not IsNull(!Vehiculo1) Then mProps.Vehiculo1 = !Vehiculo1
        If Not IsNull(!MarcaPlaca1) Then mProps.MarcaPlaca1 = !MarcaPlaca1
        If Not IsNull(!ValorVehiculo1) Then mProps.ValorVehiculo1 = !ValorVehiculo1
        If Not IsNull(!BandPrendado1) Then mProps.BandPrendado1 = !BandPrendado1
        If Not IsNull(!Vehiculo2) Then mProps.Vehiculo2 = !Vehiculo2
        If Not IsNull(!MarcaPlaca2) Then mProps.MarcaPlaca2 = !MarcaPlaca2
        If Not IsNull(!ValorVehiculo2) Then mProps.ValorVehiculo2 = !ValorVehiculo2
        If Not IsNull(!BandPrendado2) Then mProps.BandPrendado2 = !BandPrendado2
        If Not IsNull(!Vehiculo3) Then mProps.Vehiculo3 = !Vehiculo3
        If Not IsNull(!MarcaPlaca3) Then mProps.MarcaPlaca3 = !MarcaPlaca3
        If Not IsNull(!ValorVehiculo3) Then mProps.ValorVehiculo3 = !ValorVehiculo3
        If Not IsNull(!BandPrendado3) Then mProps.BandPrendado3 = !BandPrendado3
        'pasivo
        If Not IsNull(!MontoDeudaBancaria) Then mProps.MontoDeudaBancaria = !MontoDeudaBancaria
        If Not IsNull(!SaldoDeudaBancaria) Then mProps.SaldoDeudaBancaria = !SaldoDeudaBancaria
        If Not IsNull(!MensualDeudaBancaria) Then mProps.MensualDeudaBancaria = !MensualDeudaBancaria
        If Not IsNull(!MontoHipoteca) Then mProps.MontoHipoteca = !MontoHipoteca
        If Not IsNull(!SaldoHipoteca) Then mProps.SaldoHipoteca = !SaldoHipoteca
        If Not IsNull(!MensualHipoteca) Then mProps.MensualHipoteca = !MensualHipoteca
        If Not IsNull(!MontoDeudaOtarsCuentas) Then mProps.MontoDeudaOtarsCuentas = !MontoDeudaOtarsCuentas
        If Not IsNull(!SaldoDeudaOtarsCuentas) Then mProps.SaldoDeudaOtarsCuentas = !SaldoDeudaOtarsCuentas
        If Not IsNull(!MensualDeudaOtarsCuentas) Then mProps.MensualDeudaOtarsCuentas = !MensualDeudaOtarsCuentas
        
        'ref familiar
        If Not IsNull(!NombreFami) Then mProps.NombreFami = !NombreFami
        If Not IsNull(!RelacionFami) Then mProps.RelacionFami = !RelacionFami
        If Not IsNull(!TelfFami) Then mProps.TelfFami = !TelfFami
        If Not IsNull(!CiudadFami) Then mProps.CiudadFami = !CiudadFami
        
        If Not IsNull(!NombreFami2) Then mProps.NombreFami2 = !NombreFami2
        If Not IsNull(!RelacionFami2) Then mProps.RelacionFami2 = !RelacionFami2
        If Not IsNull(!TelfFami2) Then mProps.TelfFami2 = !TelfFami2
        If Not IsNull(!CiudadFami2) Then mProps.CiudadFami2 = !CiudadFami2
        
        
        'ref comer
        If Not IsNull(!NombreComer1) Then mProps.NombreComer1 = !NombreComer1
        If Not IsNull(!DirComer1) Then mProps.DirComer1 = !DirComer1
        If Not IsNull(!TelfComer1) Then mProps.TelfComer1 = !TelfComer1
        If Not IsNull(!NombreComer2) Then mProps.NombreComer2 = !NombreComer2
        If Not IsNull(!DirComer2) Then mProps.DirComer2 = !DirComer2
        If Not IsNull(!TelfComer2) Then mProps.TelfComer2 = !TelfComer2
        If Not IsNull(!NombreComer3) Then mProps.NombreComer3 = !NombreComer3
        If Not IsNull(!DirComer3) Then mProps.DirComer3 = !DirComer3
        If Not IsNull(!TelfComer3) Then mProps.TelfComer3 = !TelfComer3
        'ref.personal
        If Not IsNull(!NombrePerso1) Then mProps.NombrePerso1 = !NombrePerso1
        If Not IsNull(!DirPerso1) Then mProps.DirPerso1 = !DirPerso1
        If Not IsNull(!TelfPerso1) Then mProps.TelfPerso1 = !TelfPerso1
        If Not IsNull(!NombrePerso2) Then mProps.NombrePerso2 = !NombrePerso2
        If Not IsNull(!DirPerso2) Then mProps.DirPerso2 = !DirPerso2
        If Not IsNull(!TelfPerso2) Then mProps.TelfPerso2 = !TelfPerso2
        If Not IsNull(!NombrePerso3) Then mProps.NombrePerso3 = !NombrePerso3
        If Not IsNull(!DirPerso3) Then mProps.DirPerso3 = !DirPerso3
        If Not IsNull(!TelfPerso3) Then mProps.TelfPerso3 = !TelfPerso3
        'ref bancaria
        If Not IsNull(!NombreBanca1) Then mProps.NombreBanca1 = !NombreBanca1
        If Not IsNull(!NumCtaBanca1) Then mProps.NumCtaBanca1 = !NumCtaBanca1
        If Not IsNull(!CodTipoBanca1) Then mProps.CodTipoBanca1 = !CodTipoBanca1
        If Not IsNull(!NombreBanca2) Then mProps.NombreBanca2 = !NombreBanca2
        If Not IsNull(!NumCtaBanca2) Then mProps.NumCtaBanca2 = !NumCtaBanca2
        If Not IsNull(!CodTipoBanca2) Then mProps.CodTipoBanca2 = !CodTipoBanca2
        If Not IsNull(!NombreBanca3) Then mProps.NombreBanca3 = !NombreBanca3
        If Not IsNull(!NumCtaBanca3) Then mProps.NumCtaBanca3 = !NumCtaBanca3
        If Not IsNull(!CodTipoBanca3) Then mProps.CodTipoBanca3 = !CodTipoBanca3

        If Not IsNull(!IdInventario) Then mProps.IdInventario = !IdInventario
        If Not IsNull(!CodInventario) Then mProps.CodInventario = !CodInventario
        If Not IsNull(!BandAprobado) Then mProps.BandAprobado = !BandAprobado
        If Not IsNull(!FechaSolicitud) Then mProps.FechaSolicitud = !FechaSolicitud
        If Not IsNull(!BandCliente) Then mProps.BandCliente = !BandCliente
        If Not IsNull(!MontoCredito) Then mProps.MontoCredito = !MontoCredito
        If Not IsNull(!IdTransFuente) Then mProps.IdTransFuente = !IdTransFuente
        If Not IsNull(!TransFuente) Then mProps.TransFuente = !TransFuente
        If Not IsNull(!TipoSolicitud) Then mProps.TipoSolicitud = !TipoSolicitud
        If Not IsNull(!CodUsuarioAutoriza) Then mProps.CodUsuarioAutoriza = !CodUsuarioAutoriza
        
        mProps.FechaGrabado = !FechaGrabado
        
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    'Verifica si existe Kardex
    sql = "select g.Idtransfuente FROM PCsolcreprovcli  p  inner join gncomprobante g on g.transid = p.idtransfuente  WHERE Id = " & mProps.Id
    sql = sql & " and g.idtransfuente <> 0"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "PCSolCreProvCli.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing

    'Verifica si no esta asignado como garante
      
    'Elimina el registro
    sql = "DELETE FROM PCSolCreProvCli WHERE IdProvCli = " & mProps.IdProvCli
    mProps.objEmpresa.Execute sql, True

    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_PC", "Solicitud Eliminado: " & _
                                                         mProps.CodProvCli, "DT"
                                                         
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_SOLPROVCLI
End Sub

Public Function Clone() As PCSolCreProvCli
    Dim nuevo As PCSolCreProvCli, nc As PCContacto, i As Long
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New PCSolCreProvCli
    nuevo.Props = mProps
    
    
    Set Clone = nuevo
    Set nuevo = Nothing
    Set nc = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.Id = 0
    mProps.IdProvCli = 0
    mbooModificado = True
    
End Property


Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Sub SetModificado()
    mbooModificado = True
End Sub

Public Property Get Id() As Long
    Id = mProps.Id
End Property

Public Property Get IdProvCli() As Long
    IdProvCli = mProps.IdProvCli
End Property



Public Property Let CodProvCli(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodProvCli Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaContable también
            sql = "SELECT IdProvCli FROM PcProvCli WHERE CodProvcli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli.CodProvCli", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdProvCli = rs!IdProvCli
            End If
            rs.Close
            Set rs = Nothing
        Else
            mProps.IdProvCli = 0
        End If
    End If
    
    
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_CodProvCli", MSGERR_INVALIDO
    End If
    mProps.CodProvCli = Trim(value)
    mbooModificado = True
End Property


Public Property Get CodProvCli() As String
    CodProvCli = Trim$(mProps.CodProvCli)
End Property

Public Property Let CodEstadoCivil(ByVal value As String)
    If Len(value) > 1 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_CodEstadoCivil", MSGERR_INVALIDO
    End If
    mProps.CodEstadoCivil = Trim(value)
    mbooModificado = True
End Property

Public Property Get CodEstadoCivil() As String
    CodEstadoCivil = Trim$(mProps.CodEstadoCivil)
End Property

Public Property Let CodInstruccion(ByVal value As String)
    If Len(value) > 1 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_CodInstruccion", MSGERR_INVALIDO
    End If
    mProps.CodInstruccion = Trim(value)
    mbooModificado = True
End Property

Public Property Get CodInstruccion() As String
    CodInstruccion = Trim$(mProps.CodInstruccion)
End Property

Public Property Let BandPatriFam(ByVal value As Boolean)
    If value <> mProps.BandPatriFam Then
        mProps.BandPatriFam = value
        mbooModificado = True
    End If
End Property

Public Property Get BandPatriFam() As Boolean
    BandPatriFam = mProps.BandPatriFam
End Property

Public Property Let BandSepBienes(ByVal value As Boolean)
    If value <> mProps.BandSepBienes Then
        mProps.BandSepBienes = value
        mbooModificado = True
    End If
End Property

Public Property Get BandSepBienes() As Boolean
    BandSepBienes = mProps.BandSepBienes
End Property

Public Property Let TipoVivienda(ByVal value As String)
    If Len(value) > 1 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_TipoVivienda", MSGERR_INVALIDO
    End If
    
    mProps.TipoVivienda = value
    mbooModificado = True
End Property

Public Property Get TipoVivienda() As String
    TipoVivienda = Trim$(mProps.TipoVivienda)
End Property

Public Property Let AnioResidencia(ByVal value As String)
    If Len(value) > 2 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_AnioResidencia", MSGERR_INVALIDO
    End If
    mProps.AnioResidencia = value
    mbooModificado = True
End Property

Public Property Get AnioResidencia() As String
    AnioResidencia = Trim$(mProps.AnioResidencia)
End Property

Public Property Let NombreArrend(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_NombreArrend", MSGERR_INVALIDO
    End If
    mProps.NombreArrend = value
    mbooModificado = True
End Property

Public Property Get NombreArrend() As String
    NombreArrend = Trim$(mProps.NombreArrend)
End Property

Public Property Let TelfArrend(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_TelfArrend", MSGERR_INVALIDO
    End If
    mProps.TelfArrend = value
    mbooModificado = True
End Property

Public Property Get TelfArrend() As String
    TelfArrend = Trim$(mProps.TelfArrend)
End Property

'solicitante
Public Property Let NombreTrabajoSoli(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_NombreTrabajoSoli", MSGERR_INVALIDO
    End If
    
    mProps.NombreTrabajoSoli = value
    mbooModificado = True
End Property

Public Property Get NombreTrabajoSoli() As String
    NombreTrabajoSoli = Trim$(mProps.NombreTrabajoSoli)
End Property

Public Property Let DirTrabajoSoli(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_DirTrabajoSoli", MSGERR_INVALIDO
    End If
    mProps.DirTrabajoSoli = value
    mbooModificado = True
End Property

Public Property Get DirTrabajoSoli() As String
    DirTrabajoSoli = Trim$(mProps.DirTrabajoSoli)
End Property

Public Property Let AntiguedadSoli(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_AntiguedadSoli", MSGERR_INVALIDO
    End If
    mProps.AntiguedadSoli = value
    mbooModificado = True
End Property

Public Property Get AntiguedadSoli() As String
    AntiguedadSoli = Trim$(mProps.AntiguedadSoli)
End Property

Public Property Let TelfActividadSol(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_TelfActividadSol", MSGERR_INVALIDO
    End If
    
    mProps.TelfActividadSol = value
    mbooModificado = True
End Property

Public Property Get TelfActividadSol() As String
    TelfActividadSol = Trim$(mProps.TelfActividadSol)
End Property


Public Property Let TipoLocalSoli(ByVal value As String)
    If Len(value) > 1 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_TipoLocalSoli", MSGERR_INVALIDO
    End If
    
    mProps.TipoLocalSoli = value
    mbooModificado = True
End Property

Public Property Get TipoLocalSoli() As String
    TipoLocalSoli = Trim$(mProps.TipoLocalSoli)
End Property

Public Property Let NombreArrendSoli(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_NombreArrendSoli", MSGERR_INVALIDO
    End If
    
    mProps.NombreArrendSoli = value
    mbooModificado = True
End Property

Public Property Get NombreArrendSoli() As String
    NombreArrendSoli = Trim$(mProps.NombreArrendSoli)
End Property

Public Property Let TelfArrendSoli(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_TelfArrendSoli", MSGERR_INVALIDO
    End If
    
    mProps.TelfArrendSoli = value
    mbooModificado = True
End Property

Public Property Get TelfArrendSoli() As String
    TelfArrendSoli = Trim$(mProps.TelfArrendSoli)
End Property

Public Property Get IngresosSoli() As Currency
    IngresosSoli = mProps.IngresosSoli
End Property

Public Property Let IngresosSoli(value As Currency)
    mProps.IngresosSoli = value
    mbooModificado = True
End Property

Public Property Get OtrosIngresosSoli() As Currency
    OtrosIngresosSoli = mProps.OtrosIngresosSoli
End Property

Public Property Let OtrosIngresosSoli(value As Currency)
    mProps.OtrosIngresosSoli = value
    mbooModificado = True
End Property
'
Public Property Get GastosFamSoli() As Currency
    GastosFamSoli = mProps.GastosFamSoli
End Property

Public Property Let GastosFamSoli(value As Currency)
    mProps.GastosFamSoli = value
    mbooModificado = True
End Property
'
Public Property Get GastosArriendoSoli() As Currency
    GastosArriendoSoli = mProps.GastosArriendoSoli
End Property

Public Property Let GastosArriendoSoli(value As Currency)
    mProps.GastosArriendoSoli = value
    mbooModificado = True
End Property

Public Property Get OtrosGastosSoli() As Currency
    OtrosGastosSoli = mProps.OtrosGastosSoli
End Property

Public Property Let OtrosGastosSoli(value As Currency)
    mProps.OtrosGastosSoli = value
    mbooModificado = True
End Property

'conyuge

Public Property Let NombreConyu(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_NombreConyu", MSGERR_INVALIDO
    End If
    mProps.NombreConyu = value
    mbooModificado = True
End Property

Public Property Get NombreConyu() As String
    NombreConyu = mProps.NombreConyu
End Property

Public Property Let CIConyu(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_CIConyu", MSGERR_INVALIDO
    End If
    
    mProps.CIConyu = value
    mbooModificado = True
End Property

Public Property Get CIConyu() As String
    CIConyu = Trim$(mProps.CIConyu)
End Property

Public Property Let NombreTrabajoConyu(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_NombreTrabajoConyu", MSGERR_INVALIDO
    End If
    
    mProps.NombreTrabajoConyu = value
    mbooModificado = True
End Property

Public Property Get NombreTrabajoConyu() As String
    NombreTrabajoConyu = Trim$(mProps.NombreTrabajoConyu)
End Property

Public Property Let DirTrabajoConyu(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_DirTrabajoConyu", MSGERR_INVALIDO
    End If
    mProps.DirTrabajoConyu = value
    mbooModificado = True
End Property

Public Property Get DirTrabajoConyu() As String
    DirTrabajoConyu = Trim$(mProps.DirTrabajoConyu)
End Property

Public Property Let AntiguedadConyu(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_AntiguedadConyu", MSGERR_INVALIDO
    End If
    mProps.AntiguedadConyu = value
    mbooModificado = True
End Property

Public Property Get AntiguedadConyu() As String
    AntiguedadConyu = Trim$(mProps.AntiguedadConyu)
End Property

Public Property Let TelfActividadCounyu(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_TelfActividadCounyu", MSGERR_INVALIDO
    End If
    mProps.TelfActividadCounyu = value
    mbooModificado = True
End Property

Public Property Get TelfActividadCounyu() As String
    TelfActividadCounyu = Trim$(mProps.TelfActividadCounyu)
End Property

Public Property Let TipoLocalConyu(ByVal value As String)
    If Len(value) > 1 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_TipoLocalConyu", MSGERR_INVALIDO
    End If
    mProps.TipoLocalConyu = value
    mbooModificado = True
End Property

Public Property Get TipoLocalConyu() As String
    TipoLocalConyu = Trim$(mProps.TipoLocalConyu)
End Property


Public Property Let NombreArrendConyu(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_NombreArrendConyu", MSGERR_INVALIDO
    End If
    mProps.NombreArrendConyu = value
    mbooModificado = True
End Property

Public Property Get NombreArrendConyu() As String
    NombreArrendConyu = Trim$(mProps.NombreArrendConyu)
End Property

Public Property Let TelfArrendConyu(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_TelfArrendConyu", MSGERR_INVALIDO
    End If
    mProps.TelfArrendConyu = value
    mbooModificado = True
End Property

Public Property Get TelfArrendConyu() As String
    TelfArrendConyu = Trim$(mProps.TelfArrendConyu)
End Property

Public Property Get IngresosConyu() As Currency
    IngresosConyu = mProps.IngresosConyu
End Property

Public Property Let IngresosConyu(value As Currency)
    mProps.IngresosConyu = value
    mbooModificado = True
End Property

Public Property Get OtrosIngresosConyu() As Currency
    OtrosIngresosConyu = mProps.OtrosIngresosConyu
End Property

Public Property Let OtrosIngresosConyu(value As Currency)
    mProps.OtrosIngresosConyu = value
    mbooModificado = True
End Property
'
Public Property Get GastosFamConyu() As Currency
    GastosFamConyu = mProps.GastosFamConyu
End Property

Public Property Let GastosFamConyu(value As Currency)
    mProps.GastosFamConyu = value
    mbooModificado = True
End Property
'
Public Property Get GastosArriendoConyu() As Currency
    GastosArriendoConyu = mProps.GastosArriendoConyu
End Property

Public Property Let GastosArriendoConyu(value As Currency)
    mProps.GastosArriendoConyu = value
    mbooModificado = True
End Property

Public Property Get OtrosGastosConyu() As Currency
    OtrosGastosConyu = mProps.OtrosGastosConyu
End Property

Public Property Let OtrosGastosConyu(value As Currency)
    mProps.OtrosGastosConyu = value
    mbooModificado = True
End Property

'inmueble
Public Property Let Inmueble1(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_Inmueble1", MSGERR_INVALIDO
    End If
    mProps.Inmueble1 = value
    mbooModificado = True
End Property

Public Property Get Inmueble1() As String
    Inmueble1 = Trim$(mProps.Inmueble1)
End Property

Public Property Let Ubicacion1(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_Ubicacion1", MSGERR_INVALIDO
    End If
    mProps.Ubicacion1 = value
    mbooModificado = True
End Property


Public Property Get Ubicacion1() As String
    Ubicacion1 = Trim$(mProps.Ubicacion1)
End Property

Public Property Get ValorInmueble1() As Currency
    ValorInmueble1 = mProps.ValorInmueble1
End Property

Public Property Let ValorInmueble1(value As Currency)
    mProps.ValorInmueble1 = value
    mbooModificado = True
End Property

Public Property Let BandHipoteca1(ByVal value As Boolean)
    If value <> mProps.BandHipoteca1 Then
        mProps.BandHipoteca1 = value
        mbooModificado = True
    End If
End Property

Public Property Get BandHipoteca1() As Boolean
    BandHipoteca1 = mProps.BandHipoteca1
End Property

Public Property Let Inmueble2(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_Inmueble2", MSGERR_INVALIDO
    End If
    mProps.Inmueble2 = value
    mbooModificado = True
End Property

Public Property Get Inmueble2() As String
    Inmueble2 = Trim$(mProps.Inmueble2)
End Property

Public Property Let Ubicacion2(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_Ubicacion2", MSGERR_INVALIDO
    End If
    mProps.Ubicacion2 = value
    mbooModificado = True
End Property


Public Property Get Ubicacion2() As String
    Ubicacion2 = Trim$(mProps.Ubicacion2)
End Property

Public Property Get ValorInmueble2() As Currency
    ValorInmueble2 = mProps.ValorInmueble2
End Property

Public Property Let ValorInmueble2(value As Currency)
    mProps.ValorInmueble2 = value
    mbooModificado = True
End Property

Public Property Let BandHipoteca2(ByVal value As Boolean)
    If value <> mProps.BandHipoteca2 Then
        mProps.BandHipoteca2 = value
        mbooModificado = True
    End If
End Property

Public Property Get BandHipoteca2() As Boolean
    BandHipoteca2 = mProps.BandHipoteca2
End Property

Public Property Let Inmueble3(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_Inmueble3", MSGERR_INVALIDO
    End If
    mProps.Inmueble3 = value
    mbooModificado = True
End Property

Public Property Get Inmueble3() As String
    Inmueble3 = Trim$(mProps.Inmueble3)
End Property

Public Property Let Ubicacion3(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_Ubicacion3", MSGERR_INVALIDO
    End If
    mProps.Ubicacion3 = value
    mbooModificado = True
End Property


Public Property Get Ubicacion3() As String
    Ubicacion3 = Trim$(mProps.Ubicacion3)
End Property

Public Property Get ValorInmueble3() As Currency
    ValorInmueble3 = mProps.ValorInmueble3
End Property

Public Property Let ValorInmueble3(value As Currency)
    mProps.ValorInmueble3 = value
    mbooModificado = True
End Property

Public Property Let BandHipoteca3(ByVal value As Boolean)
    If value <> mProps.BandHipoteca3 Then
        mProps.BandHipoteca3 = value
        mbooModificado = True
    End If
End Property

Public Property Get BandHipoteca3() As Boolean
    BandHipoteca3 = mProps.BandHipoteca3
End Property

'vehiculo
Public Property Let Vehiculo1(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_Vehiculo1", MSGERR_INVALIDO
    End If
    mProps.Vehiculo1 = value
    mbooModificado = True
End Property

Public Property Get Vehiculo1() As String
    Vehiculo1 = Trim$(mProps.Vehiculo1)
End Property

Public Property Let MarcaPlaca1(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_MarcaPlaca1", MSGERR_INVALIDO
    End If
    mProps.MarcaPlaca1 = value
    mbooModificado = True
End Property

Public Property Get MarcaPlaca1() As String
    MarcaPlaca1 = mProps.MarcaPlaca1
End Property

Public Property Get ValorVehiculo1() As Currency
    ValorVehiculo1 = mProps.ValorVehiculo1
End Property

Public Property Let ValorVehiculo1(value As Currency)
    mProps.ValorVehiculo1 = value
    mbooModificado = True
End Property

Public Property Let BandPrendado1(ByVal value As Boolean)
    If value <> mProps.BandPrendado1 Then
        mProps.BandPrendado1 = value
        mbooModificado = True
    End If
End Property

Public Property Get BandPrendado1() As Boolean
    BandPrendado1 = mProps.BandPrendado1
End Property
'---------
Public Property Let Vehiculo2(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_Vehiculo2", MSGERR_INVALIDO
    End If
    mProps.Vehiculo2 = value
    mbooModificado = True
End Property

Public Property Get Vehiculo2() As String
    Vehiculo2 = Trim$(mProps.Vehiculo2)
End Property

Public Property Let MarcaPlaca2(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_MarcaPlaca2", MSGERR_INVALIDO
    End If
    mProps.MarcaPlaca2 = value
    mbooModificado = True
End Property

Public Property Get MarcaPlaca2() As String
    MarcaPlaca2 = mProps.MarcaPlaca2
End Property

Public Property Get ValorVehiculo2() As Currency
    ValorVehiculo2 = mProps.ValorVehiculo2
End Property

Public Property Let ValorVehiculo2(value As Currency)
    mProps.ValorVehiculo2 = value
    mbooModificado = True
End Property

Public Property Let BandPrendado2(ByVal value As Boolean)
    If value <> mProps.BandPrendado2 Then
        mProps.BandPrendado2 = value
        mbooModificado = True
    End If
End Property

Public Property Get BandPrendado2() As Boolean
    BandPrendado2 = mProps.BandPrendado2
End Property

Public Property Let Vehiculo3(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_Vehiculo3", MSGERR_INVALIDO
    End If
    mProps.Vehiculo3 = value
    mbooModificado = True
End Property

Public Property Get Vehiculo3() As String
    Vehiculo3 = Trim$(mProps.Vehiculo3)
End Property

Public Property Let MarcaPlaca3(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_MarcaPlaca3", MSGERR_INVALIDO
    End If
    mProps.MarcaPlaca3 = value
    mbooModificado = True
End Property

Public Property Get MarcaPlaca3() As String
    MarcaPlaca3 = mProps.MarcaPlaca3
End Property

Public Property Get ValorVehiculo3() As Currency
    ValorVehiculo3 = mProps.ValorVehiculo3
End Property

Public Property Let ValorVehiculo3(value As Currency)
    mProps.ValorVehiculo3 = value
    mbooModificado = True
End Property

Public Property Let BandPrendado3(ByVal value As Boolean)
    If value <> mProps.BandPrendado3 Then
        mProps.BandPrendado3 = value
        mbooModificado = True
    End If
End Property

Public Property Get BandPrendado3() As Boolean
    BandPrendado3 = mProps.BandPrendado3
End Property
'pasivo
Public Property Get MontoDeudaBancaria() As Currency
    MontoDeudaBancaria = mProps.MontoDeudaBancaria
End Property

Public Property Let MontoDeudaBancaria(value As Currency)
    mProps.MontoDeudaBancaria = value
    mbooModificado = True
End Property

Public Property Get SaldoDeudaBancaria() As Currency
    SaldoDeudaBancaria = mProps.SaldoDeudaBancaria
End Property

Public Property Let SaldoDeudaBancaria(value As Currency)
    mProps.SaldoDeudaBancaria = value
    mbooModificado = True
End Property

Public Property Get MensualDeudaBancaria() As Currency
    MensualDeudaBancaria = mProps.MensualDeudaBancaria
End Property

Public Property Let MensualDeudaBancaria(value As Currency)
    mProps.MensualDeudaBancaria = value
    mbooModificado = True
End Property

Public Property Get MontoHipoteca() As Currency
    MontoHipoteca = mProps.MontoHipoteca
End Property

Public Property Let MontoHipoteca(value As Currency)
    mProps.MontoHipoteca = value
    mbooModificado = True
End Property

Public Property Get SaldoHipoteca() As Currency
    SaldoHipoteca = mProps.SaldoHipoteca
End Property

Public Property Let SaldoHipoteca(value As Currency)
    mProps.SaldoHipoteca = value
    mbooModificado = True
End Property

Public Property Get MensualHipoteca() As Currency
    MensualHipoteca = mProps.MensualHipoteca
End Property

Public Property Let MensualHipoteca(value As Currency)
    mProps.MensualHipoteca = value
    mbooModificado = True
End Property

Public Property Get MontoDeudaOtarsCuentas() As Currency
    MontoDeudaOtarsCuentas = mProps.MontoDeudaOtarsCuentas
End Property

Public Property Let MontoDeudaOtarsCuentas(value As Currency)
    mProps.MontoDeudaOtarsCuentas = value
    mbooModificado = True
End Property

Public Property Get SaldoDeudaOtarsCuentas() As Currency
    SaldoDeudaOtarsCuentas = mProps.SaldoDeudaOtarsCuentas
End Property

Public Property Let SaldoDeudaOtarsCuentas(value As Currency)
    mProps.SaldoDeudaOtarsCuentas = value
    mbooModificado = True
End Property

Public Property Get MensualDeudaOtarsCuentas() As Currency
    MensualDeudaOtarsCuentas = mProps.MensualDeudaOtarsCuentas
End Property

Public Property Let MensualDeudaOtarsCuentas(value As Currency)
    mProps.MensualDeudaOtarsCuentas = value
    mbooModificado = True
End Property

'ref familias
Public Property Get NombreFami() As String
    NombreFami = mProps.NombreFami
End Property

Public Property Let NombreFami(ByVal value As String)
    If (Len(value) > 80) Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli.NombreFami", MSGERR_INVALIDO
    End If
    mProps.NombreFami = value
    mbooModificado = True
End Property

Public Property Get RelacionFami() As String
    RelacionFami = mProps.RelacionFami
End Property

Public Property Let RelacionFami(ByVal value As String)
    If Len(value) > 40 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli.RelacionFami", MSGERR_INVALIDO
    End If
    mProps.RelacionFami = value
    mbooModificado = True
End Property

Public Property Let TelfFami(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_TelfFami", MSGERR_INVALIDO
    End If
    mProps.TelfFami = value
    mbooModificado = True
End Property

Public Property Get TelfFami() As String
    TelfFami = Trim$(mProps.TelfFami)
End Property

Public Property Let CiudadFami(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_CiudadFami", MSGERR_INVALIDO
    End If
    mProps.CiudadFami = value
    mbooModificado = True
End Property

Public Property Get CiudadFami() As String
    CiudadFami = Trim$(mProps.CiudadFami)
End Property

'ref per 2

Public Property Get NombreFami2() As String
    NombreFami2 = mProps.NombreFami2
End Property

Public Property Let NombreFami2(ByVal value As String)
    If (Len(value) > 80) Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli.NombreFami2", MSGERR_INVALIDO
    End If
    mProps.NombreFami2 = value
    mbooModificado = True
End Property

Public Property Get RelacionFami2() As String
    RelacionFami2 = mProps.RelacionFami2
End Property

Public Property Let RelacionFami2(ByVal value As String)
    If Len(value) > 40 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli.RelacionFami2", MSGERR_INVALIDO
    End If
    mProps.RelacionFami2 = value
    mbooModificado = True
End Property

Public Property Let TelfFami2(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_TelfFami2", MSGERR_INVALIDO
    End If
    mProps.TelfFami2 = value
    mbooModificado = True
End Property

Public Property Get TelfFami2() As String
    TelfFami2 = Trim$(mProps.TelfFami2)
End Property

Public Property Let CiudadFami2(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_CiudadFami2", MSGERR_INVALIDO
    End If
    mProps.CiudadFami2 = value
    mbooModificado = True
End Property

Public Property Get CiudadFami2() As String
    CiudadFami2 = Trim$(mProps.CiudadFami2)
End Property



'ref comercial
Public Property Let NombreComer1(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_NombreComer1", MSGERR_INVALIDO
    End If
    
    mProps.NombreComer1 = value
    mbooModificado = True
End Property

Public Property Get NombreComer1() As String
    NombreComer1 = Trim$(mProps.NombreComer1)
End Property

Public Property Let DirComer1(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_DirComer1", MSGERR_INVALIDO
    End If
    mProps.DirComer1 = value
    mbooModificado = True
End Property

Public Property Get DirComer1() As String
    DirComer1 = mProps.DirComer1
End Property

Public Property Let TelfComer1(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_TelfComer1", MSGERR_INVALIDO
    End If
    mProps.TelfComer1 = value
    mbooModificado = True
End Property

Public Property Get TelfComer1() As String
    TelfComer1 = Trim$(mProps.TelfComer1)
End Property

Public Property Let NombreComer2(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_NombreComer2", MSGERR_INVALIDO
    End If
    
    mProps.NombreComer2 = value
    mbooModificado = True
End Property

Public Property Get NombreComer2() As String
    NombreComer2 = Trim$(mProps.NombreComer2)
End Property

Public Property Let DirComer2(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_DirComer2", MSGERR_INVALIDO
    End If
    mProps.DirComer2 = value
    mbooModificado = True
End Property

Public Property Get DirComer2() As String
    DirComer2 = mProps.DirComer2
End Property

Public Property Let TelfComer2(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_TelfComer2", MSGERR_INVALIDO
    End If
    mProps.TelfComer2 = value
    mbooModificado = True
End Property

Public Property Get TelfComer2() As String
    TelfComer2 = Trim$(mProps.TelfComer2)
End Property

Public Property Let NombreComer3(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_NombreComer3", MSGERR_INVALIDO
    End If
    
    mProps.NombreComer3 = value
    mbooModificado = True
End Property

Public Property Get NombreComer3() As String
    NombreComer3 = Trim$(mProps.NombreComer3)
End Property

Public Property Let DirComer3(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_DirComer3", MSGERR_INVALIDO
    End If
    mProps.DirComer3 = value
    mbooModificado = True
End Property

Public Property Get DirComer3() As String
    DirComer3 = mProps.DirComer3
End Property

Public Property Let TelfComer3(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_TelfComer3", MSGERR_INVALIDO
    End If
    mProps.TelfComer3 = value
    mbooModificado = True
End Property

Public Property Get TelfComer3() As String
    TelfComer3 = Trim$(mProps.TelfComer3)
End Property

'ref personal
Public Property Let NombrePerso1(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_NombrePerso1", MSGERR_INVALIDO
    End If
    
    mProps.NombrePerso1 = value
    mbooModificado = True
End Property

Public Property Get NombrePerso1() As String
    NombrePerso1 = Trim$(mProps.NombrePerso1)
End Property

Public Property Let DirPerso1(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_DirPerso1", MSGERR_INVALIDO
    End If
    mProps.DirPerso1 = value
    mbooModificado = True
End Property

Public Property Get DirPerso1() As String
    DirPerso1 = Trim$(mProps.DirPerso1)
End Property

Public Property Let TelfPerso1(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_TelfPerso1", MSGERR_INVALIDO
    End If
    mProps.TelfPerso1 = value
    mbooModificado = True
End Property

Public Property Get TelfPerso1() As String
    TelfPerso1 = Trim$(mProps.TelfPerso1)
End Property

Public Property Let NombrePerso2(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_NombrePerso2", MSGERR_INVALIDO
    End If
    
    mProps.NombrePerso2 = value
    mbooModificado = True
End Property

Public Property Get NombrePerso2() As String
    NombrePerso2 = Trim$(mProps.NombrePerso2)
End Property

Public Property Let DirPerso2(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_DirPerso2", MSGERR_INVALIDO
    End If
    mProps.DirPerso2 = value
    mbooModificado = True
End Property

Public Property Get DirPerso2() As String
    DirPerso2 = Trim$(mProps.DirPerso2)
End Property

Public Property Let TelfPerso2(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_TelfPerso2", MSGERR_INVALIDO
    End If
    mProps.TelfPerso2 = value
    mbooModificado = True
End Property

Public Property Get TelfPerso2() As String
    TelfPerso2 = Trim$(mProps.TelfPerso2)
End Property

Public Property Let NombrePerso3(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_NombrePerso3", MSGERR_INVALIDO
    End If
    
    mProps.NombrePerso3 = value
    mbooModificado = True
End Property

Public Property Get NombrePerso3() As String
    NombrePerso3 = Trim$(mProps.NombrePerso3)
End Property

Public Property Let DirPerso3(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_DirPerso3", MSGERR_INVALIDO
    End If
    mProps.DirPerso3 = value
    mbooModificado = True
End Property

Public Property Get DirPerso3() As String
    DirPerso3 = Trim$(mProps.DirPerso3)
End Property

Public Property Let TelfPerso3(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_TelfPerso3", MSGERR_INVALIDO
    End If
    mProps.TelfPerso3 = value
    mbooModificado = True
End Property

Public Property Get TelfPerso3() As String
    TelfPerso3 = Trim$(mProps.TelfPerso3)
End Property

'ref bancarias
Public Property Let NombreBanca1(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_NombreBanca1", MSGERR_INVALIDO
    End If
    mProps.NombreBanca1 = value
    mbooModificado = True
End Property

Public Property Get NombreBanca1() As String
    NombreBanca1 = Trim$(mProps.NombreBanca1)
End Property

Public Property Let NumCtaBanca1(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_NumCtaBanca1", MSGERR_INVALIDO
    End If
    mProps.NumCtaBanca1 = value
    mbooModificado = True
End Property

Public Property Get NumCtaBanca1() As String
    NumCtaBanca1 = Trim$(mProps.NumCtaBanca1)
End Property

Public Property Let CodTipoBanca1(ByVal value As String)
    If Len(value) > 1 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_CodTipoBanca1", MSGERR_INVALIDO
    End If
    mProps.CodTipoBanca1 = value
    mbooModificado = True
End Property

Public Property Get CodTipoBanca1() As String
    CodTipoBanca1 = Trim$(mProps.CodTipoBanca1)
End Property

Public Property Let NombreBanca2(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_NombreBanca2", MSGERR_INVALIDO
    End If
    mProps.NombreBanca2 = value
    mbooModificado = True
End Property

Public Property Get NombreBanca2() As String
    NombreBanca2 = Trim$(mProps.NombreBanca2)
End Property

Public Property Let NumCtaBanca2(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_NumCtaBanca2", MSGERR_INVALIDO
    End If
    mProps.NumCtaBanca2 = value
    mbooModificado = True
End Property

Public Property Get NumCtaBanca2() As String
    NumCtaBanca2 = Trim$(mProps.NumCtaBanca2)
End Property

Public Property Let CodTipoBanca2(ByVal value As String)
    If Len(value) > 2 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_CodTipoBanca2", MSGERR_INVALIDO
    End If
    mProps.CodTipoBanca2 = value
    mbooModificado = True
End Property

Public Property Get CodTipoBanca2() As String
    CodTipoBanca2 = Trim$(mProps.CodTipoBanca2)
End Property

Public Property Let NombreBanca3(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_NombreBanca3", MSGERR_INVALIDO
    End If
    mProps.NombreBanca3 = value
    mbooModificado = True
End Property

Public Property Get NombreBanca3() As String
    NombreBanca3 = Trim$(mProps.NombreBanca3)
End Property

Public Property Let NumCtaBanca3(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_NumCtaBanca3", MSGERR_INVALIDO
    End If
    mProps.NumCtaBanca3 = value
    mbooModificado = True
End Property

Public Property Get NumCtaBanca3() As String
    NumCtaBanca3 = Trim$(mProps.NumCtaBanca3)
End Property

Public Property Let CodTipoBanca3(ByVal value As String)
    If Len(value) > 3 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_CodTipoBanca3", MSGERR_INVALIDO
    End If
    mProps.CodTipoBanca3 = value
    mbooModificado = True
End Property

Public Property Get CodTipoBanca3() As String
    CodTipoBanca3 = Trim$(mProps.CodTipoBanca3)
End Property

Public Property Get FechaGrabado() As Date
    FechaGrabado = mProps.FechaGrabado
End Property

Public Property Get EsNuevo() As Boolean
    EsNuevo = mbooEsNuevo
End Property

Public Property Get Modificado() As Boolean
    Modificado = mbooModificado
End Property

Private Sub Class_Initialize()
    mbooEsNuevo = True
    mbooModificado = False
    mProps.BandCliente = True
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Sub GrabarEnSucursal(ByVal bd As String, ByVal servidor As String)
    Dim sql As String, fgrabado As Date, i As Long
    Dim rs As Recordset
    
    
    If mbooEsNuevo Then
        'Verifica si no repite el código
        sql = "SELECT CodProvCli FROM " & bd & ".dbo.PCSolCreProvCli " & _
              "WHERE (CodProvCli='" & mProps.CodProvCli & "') AND " & _
                    "(IdProvCli<>" & mProps.IdProvCli & ")"
        
        Set rs = mProps.objEmpresa.OpenRecordsetSUC(sql, bd, servidor)
        If Not rs.EOF Then      'Si ya existe el codigo
            Err.Raise ERR_REPITECODIGO, "PCSolCreProvCli", MSGERR_REPITECODIGO
        End If
        rs.Close
    End If
    
    On Error GoTo ErrTrap

''#If DAOLIB Then
''    gobjWS.BeginTrans                           '======= Comienza la transacción
''#Else
''    mProps.objEmpresa.Coneccion.BeginTrans      '======= Comienza la transacción
''#End If
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM " & bd & ".dbo.PCSolCreProvCli WHERE 1=0"
                Set rs = mProps.objEmpresa.OpenRecordsetSUC(sql, bd, servidor)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM " & bd & ".dbo.PCSolCreProvCli WHERE CodProvCli='" & mProps.CodProvCli & "'"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEditSuc(sql, bd, servidor)
        If rs.RecordCount = "0" Then
            sql = "SELECT * FROM " & bd & ".dbo.PCSolCreProvCli WHERE 1=0"
            Set rs = mProps.objEmpresa.OpenRecordsetParaEditSuc(sql, bd, servidor)
            rs.AddNew
        
        End If
        
        
#If DAOLIB Then
        rs.Edit
#End If
        
    End If

    'Graba el registro en la base
    With rs
        !CodProvCli = mProps.CodProvCli
        !NombreArrend = mProps.NombreArrend
        !BandPatriFam = mProps.BandPatriFam
        !BandSepBienes = mProps.BandSepBienes
        !NombreTrabajoSoli = mProps.NombreTrabajoSoli
        !DirTrabajoSoli = mProps.DirTrabajoSoli
        !TipoVivienda = mProps.TipoVivienda
        !AnioResidencia = mProps.AnioResidencia
        !AntiguedadSoli = mProps.AntiguedadSoli
        !TelfActividadSol = mProps.TelfActividadSol
        !TelfArrend = mProps.TelfArrend
        !TipoLocalSoli = mProps.TipoLocalSoli
        !NombreArrendSoli = mProps.NombreArrendSoli
        !TelfArrendSoli = mProps.TelfArrendSoli
        !NombreConyu = mProps.NombreConyu
        !CIConyu = mProps.CIConyu
        
        !IngresosSoli = mProps.IngresosSoli
        !NombreTrabajoConyu = mProps.NombreTrabajoConyu
        !DirTrabajoConyu = mProps.DirTrabajoConyu
        !AntiguedadConyu = mProps.AntiguedadConyu
        !TelfActividadCounyu = mProps.TelfActividadCounyu
        !TipoLocalConyu = mProps.TipoLocalConyu
        !MarcaPlaca1 = mProps.MarcaPlaca1           'Agregado Alex: Sept/2002
        !NombreFami = mProps.NombreFami       'Agregado Alex: Sept/2002
        !RelacionFami = mProps.RelacionFami                   'Agregado Alex: Sept/2002
        !TelfFami = mProps.TelfFami           '***Agregado. 05/sep/2003. Angel
        !ValorInmueble1 = mProps.ValorInmueble1                   '***Agregado. 08/sep/2003. Angel
        !ValorVehiculo1 = mProps.ValorVehiculo1                 '***Agregado. 08/sep/2003. Angel
        !NombreComer1 = mProps.NombreComer1  'Agregado AUC 31/08/2005
        !DirComer1 = mProps.DirComer1       ' jeaa 17/12/2005
        !TelfComer1 = mProps.TelfComer1  'Agregado jeaa 12/08/2006
        !BandHipoteca1 = mProps.BandHipoteca1   'jeaa 16/01/2008 cambio por EQUIDAD TRIBUTARIA
        !BandPrendado1 = mProps.BandPrendado1
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        fgrabado = !FechaGrabado
''        mProps.IdProvCli = !IdProvCli
        .Close
    End With
    
    'Graba todos los PCContacto
''''    For i = 1 To CountContacto
''''        Contactos(i).Grabar
''''    Next i
    
    mProps.FechaGrabado = fgrabado
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    
    
    
    
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

Public Property Let FechaNacimientoConyu(value As Date)
    If Not (IsDate(value)) Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_FechaNacimientoConyu", MSGERR_INVALIDO
    End If
    mProps.FechaNacimientoConyu = value
    mbooModificado = True
End Property

Public Property Get FechaNacimientoConyu() As Date
    FechaNacimientoConyu = mProps.FechaNacimientoConyu
End Property

Public Property Let ActividadSoli(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_ActividadSoli", MSGERR_INVALIDO
    End If
    mProps.ActividadSoli = value
    mbooModificado = True
End Property

Public Property Get ActividadSoli() As String
    ActividadSoli = Trim$(mProps.ActividadSoli)
End Property

Public Property Let ActividadConyu(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_ActividadConyu", MSGERR_INVALIDO
    End If
    mProps.ActividadConyu = value
    mbooModificado = True
End Property

Public Property Get ActividadConyu() As String
    ActividadConyu = Trim$(mProps.ActividadConyu)
End Property

Public Property Let NumcargasConyu(ByVal value As String)
    If Len(value) > 2 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_NumcargasConyu", MSGERR_INVALIDO
    End If
    mProps.NumcargasConyu = value
    mbooModificado = True
End Property

Public Property Get NumcargasConyu() As String
    NumcargasConyu = Trim$(mProps.NumcargasConyu)
End Property

Friend Property Let IdInventario(ByVal value As Long)
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
            sql = "SELECT IdInventario " & _
                  "FROM IVInventario WHERE CodInventario='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCSolCreProvCli_CodInventario", MSGERR_NOHAYCODIGO
            Else
                mProps.IdInventario = rs!IdInventario
            End If
            rs.Close
        Else
            mProps.IdInventario = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodInventario = value
    mbooModificado = True
End Property

Public Property Get CodInventario() As String
    CodInventario = Trim$(mProps.CodInventario)
End Property


Public Property Let BandAprobado(ByVal value As Boolean)
    If value <> mProps.BandAprobado Then
        mProps.BandAprobado = value
        mbooModificado = True
    End If
End Property

Public Property Get BandAprobado() As Boolean
    BandAprobado = mProps.BandAprobado
End Property

Public Property Let FechaSolicitud(value As Date)
    If Not (IsDate(value)) Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_FechaSolicitud", MSGERR_INVALIDO
    End If
    mProps.FechaSolicitud = value
    mbooModificado = True
End Property

Public Property Get FechaSolicitud() As Date
    FechaSolicitud = mProps.FechaSolicitud
End Property

Public Property Let BandCliente(ByVal value As Boolean)
    If value <> mProps.BandCliente Then
        mProps.BandCliente = value
        mbooModificado = True
    End If
End Property

Public Property Get BandCliente() As Boolean
    BandCliente = mProps.BandCliente
End Property

Public Property Get MontoCredito() As Currency
    MontoCredito = mProps.MontoCredito
End Property

Public Property Let MontoCredito(value As Currency)
    mProps.MontoCredito = value
    mbooModificado = True
End Property

Public Property Let IdTransFuente(ByVal value As Long)
    mProps.IdTransFuente = value
    mbooModificado = True
End Property

Public Property Get IdTransFuente() As Long
    IdTransFuente = mProps.IdTransFuente
End Property

Public Property Let TransFuente(ByVal value As String)
    Dim sql As String, rs As Recordset
    Dim v As Variant
    If Len(value) <> 0 Then
        v = Split(value, "-")
        'Cuando cambia el valor
        If value <> mProps.TransFuente Then
            If Len(value) > 0 Then
                'Actualiza IdTrans también
                sql = "SELECT TransId " & _
                      "FROM GnComprobante WHERE CodTrans='" & v(0) & "' and numtrans=" & v(1)
                Set rs = mProps.objEmpresa.OpenRecordset(sql)
                If rs.EOF Then
                    Err.Raise ERR_NOHAYCODIGO, "PCSolCreProvCli_CodTrans", MSGERR_NOHAYCODIGO
                Else
                    mProps.IdTransFuente = rs!transid
                End If
                rs.Close
            Else
                mProps.IdTransFuente = 0
            End If
        End If
    
        Set rs = Nothing
        mProps.TransFuente = value
        mbooModificado = True
    Else
        mProps.IdTransFuente = 0
    End If
End Property

Public Property Get TransFuente() As String
    TransFuente = Trim$(mProps.TransFuente)
End Property

Public Property Let TipoSolicitud(ByVal value As Byte)
    mProps.TipoSolicitud = value
    mbooModificado = True
End Property

Public Property Get TipoSolicitud() As Byte
    TipoSolicitud = mProps.TipoSolicitud
End Property

Public Function EstaAprobada(ByVal tiposol As Integer) As Boolean
Dim sql As String
Dim rs As Recordset
    sql = "select * from  vwPCSolCreProvCli where  IdTransFuente = " & Me.IdTransFuente
    sql = sql & " and bandaprobado = 1 AND tiposolicitud =  " & tiposol
     sql = sql & " and id <>" & Me.Id
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
            Err.Raise ERR_INVALIDO, "PCSolCreProvCli_EstaAprobada", "Ya Existe una solicitud anterior" & Chr(13) & "Transaccion:" & rs!TransFuente & Chr(13) & " Cliente: " & rs!Nombre
            EstaAprobada = True
    End If
End Function

Friend Sub RecuperarxTransFuente(key As Long, bandCli As Boolean)
    Dim sql As String, i As Long
    Dim rs As Recordset
    
        sql = "SELECT * FROM vwPCSolCreProvCli WHERE idtransfuente=" & key
        If bandCli Then
            sql = sql & " And TipoSolicitud= 0"
        Else
            sql = sql & " And TipoSolicitud= 1"
        End If
            
    Set rs = mProps.objEmpresa.OpenRecordset(sql)

    'Recupera los datos de Usuario
    With rs
        mProps.Id = !Id
        If Not IsNull(!IdProvCli) Then mProps.IdProvCli = !IdProvCli
        If Not IsNull(!CodProvCli) Then mProps.CodProvCli = !CodProvCli
        If Not IsNull(!CodEstadoCivil) Then mProps.CodEstadoCivil = !CodEstadoCivil
        If Not IsNull(!CodInstruccion) Then mProps.CodInstruccion = !CodInstruccion
        If Not IsNull(!BandPatriFam) Then mProps.BandPatriFam = !BandPatriFam
        If Not IsNull(!BandSepBienes) Then mProps.BandSepBienes = !BandSepBienes
        If Not IsNull(!TipoVivienda) Then mProps.TipoVivienda = !TipoVivienda
        If Not IsNull(!AnioResidencia) Then mProps.AnioResidencia = !AnioResidencia
        If Not IsNull(!NombreArrend) Then mProps.NombreArrend = !NombreArrend
        If Not IsNull(!TelfArrend) Then mProps.TelfArrend = !TelfArrend
        'solicitante
        If Not IsNull(!ActividadSoli) Then mProps.ActividadSoli = !ActividadSoli
        If Not IsNull(!NombreTrabajoSoli) Then mProps.NombreTrabajoSoli = !NombreTrabajoSoli
        If Not IsNull(!DirTrabajoSoli) Then mProps.DirTrabajoSoli = !DirTrabajoSoli
        If Not IsNull(!AntiguedadSoli) Then mProps.AntiguedadSoli = !AntiguedadSoli
        If Not IsNull(!TelfActividadSol) Then mProps.TelfActividadSol = !TelfActividadSol
        If Not IsNull(!TipoLocalSoli) Then mProps.TipoLocalSoli = !TipoLocalSoli
        If Not IsNull(!NombreArrendSoli) Then mProps.NombreArrendSoli = !NombreArrendSoli
        If Not IsNull(!TelfArrendSoli) Then mProps.TelfArrendSoli = !TelfArrendSoli
        If Not IsNull(!IngresosSoli) Then mProps.IngresosSoli = !IngresosSoli
        If Not IsNull(!OtrosIngresosSoli) Then mProps.OtrosIngresosSoli = !OtrosIngresosSoli
        If Not IsNull(!GastosFamSoli) Then mProps.GastosFamSoli = !GastosFamSoli
        If Not IsNull(!GastosArriendoSoli) Then mProps.GastosArriendoSoli = !GastosArriendoSoli
        If Not IsNull(!OtrosGastosSoli) Then mProps.OtrosGastosSoli = !OtrosGastosSoli
        'conyuge
        If Not IsNull(!NombreConyu) Then mProps.NombreConyu = !NombreConyu
        If Not IsNull(!CIConyu) Then mProps.CIConyu = !CIConyu
        If Not IsNull(!FechaNacimientoConyu) Then mProps.FechaNacimientoConyu = !FechaNacimientoConyu
        If Not IsNull(!NumcargasConyu) Then mProps.NumcargasConyu = !NumcargasConyu
        If Not IsNull(!ActividadConyu) Then mProps.ActividadConyu = !ActividadConyu
        If Not IsNull(!NombreTrabajoConyu) Then mProps.NombreTrabajoConyu = !NombreTrabajoConyu
        If Not IsNull(!DirTrabajoConyu) Then mProps.DirTrabajoConyu = !DirTrabajoConyu
        If Not IsNull(!AntiguedadConyu) Then mProps.AntiguedadConyu = !AntiguedadConyu
        If Not IsNull(!TelfActividadCounyu) Then mProps.TelfActividadCounyu = !TelfActividadCounyu
        If Not IsNull(!TipoLocalConyu) Then mProps.TipoLocalConyu = !TipoLocalConyu
        If Not IsNull(!NombreArrendConyu) Then mProps.NombreArrendConyu = !NombreArrendConyu
        If Not IsNull(!TelfArrendConyu) Then mProps.TelfArrendConyu = !TelfArrendConyu
        If Not IsNull(!IngresosConyu) Then mProps.IngresosConyu = !IngresosConyu
        If Not IsNull(!OtrosIngresosConyu) Then mProps.OtrosIngresosConyu = !OtrosIngresosConyu
        If Not IsNull(!GastosFamConyu) Then mProps.GastosFamConyu = !GastosFamConyu
        If Not IsNull(!GastosArriendoConyu) Then mProps.GastosArriendoConyu = !GastosArriendoConyu
        If Not IsNull(!OtrosGastosConyu) Then mProps.OtrosGastosConyu = !OtrosGastosConyu
        'inmuebles
        If Not IsNull(!Inmueble1) Then mProps.Inmueble1 = !Inmueble1
        If Not IsNull(!Ubicacion1) Then mProps.Ubicacion1 = !Ubicacion1
        If Not IsNull(!ValorInmueble1) Then mProps.ValorInmueble1 = !ValorInmueble1
        If Not IsNull(!BandHipoteca1) Then mProps.BandHipoteca1 = !BandHipoteca1
        If Not IsNull(!Inmueble2) Then mProps.Inmueble2 = !Inmueble2
        If Not IsNull(!Ubicacion2) Then mProps.Ubicacion2 = !Ubicacion2
        If Not IsNull(!ValorInmueble2) Then mProps.ValorInmueble2 = !ValorInmueble2
        If Not IsNull(!BandHipoteca2) Then mProps.BandHipoteca2 = !BandHipoteca2
        If Not IsNull(!Inmueble3) Then mProps.Inmueble3 = !Inmueble3
        If Not IsNull(!Ubicacion3) Then mProps.Ubicacion3 = !Ubicacion3
        If Not IsNull(!ValorInmueble3) Then mProps.ValorInmueble3 = !ValorInmueble3
        If Not IsNull(!BandHipoteca3) Then mProps.BandHipoteca3 = !BandHipoteca3
        
        'vehiculo
        If Not IsNull(!Vehiculo1) Then mProps.Vehiculo1 = !Vehiculo1
        If Not IsNull(!MarcaPlaca1) Then mProps.MarcaPlaca1 = !MarcaPlaca1
        If Not IsNull(!ValorVehiculo1) Then mProps.ValorVehiculo1 = !ValorVehiculo1
        If Not IsNull(!BandPrendado1) Then mProps.BandPrendado1 = !BandPrendado1
        If Not IsNull(!Vehiculo2) Then mProps.Vehiculo2 = !Vehiculo2
        If Not IsNull(!MarcaPlaca2) Then mProps.MarcaPlaca2 = !MarcaPlaca2
        If Not IsNull(!ValorVehiculo2) Then mProps.ValorVehiculo2 = !ValorVehiculo2
        If Not IsNull(!BandPrendado2) Then mProps.BandPrendado2 = !BandPrendado2
        If Not IsNull(!Vehiculo3) Then mProps.Vehiculo3 = !Vehiculo3
        If Not IsNull(!MarcaPlaca3) Then mProps.MarcaPlaca3 = !MarcaPlaca3
        If Not IsNull(!ValorVehiculo3) Then mProps.ValorVehiculo3 = !ValorVehiculo3
        If Not IsNull(!BandPrendado3) Then mProps.BandPrendado3 = !BandPrendado3
        'pasivo
        If Not IsNull(!MontoDeudaBancaria) Then mProps.MontoDeudaBancaria = !MontoDeudaBancaria
        If Not IsNull(!SaldoDeudaBancaria) Then mProps.SaldoDeudaBancaria = !SaldoDeudaBancaria
        If Not IsNull(!MensualDeudaBancaria) Then mProps.MensualDeudaBancaria = !MensualDeudaBancaria
        If Not IsNull(!MontoHipoteca) Then mProps.MontoHipoteca = !MontoHipoteca
        If Not IsNull(!SaldoHipoteca) Then mProps.SaldoHipoteca = !SaldoHipoteca
        If Not IsNull(!MensualHipoteca) Then mProps.MensualHipoteca = !MensualHipoteca
        If Not IsNull(!MontoDeudaOtarsCuentas) Then mProps.MontoDeudaOtarsCuentas = !MontoDeudaOtarsCuentas
        If Not IsNull(!SaldoDeudaOtarsCuentas) Then mProps.SaldoDeudaOtarsCuentas = !SaldoDeudaOtarsCuentas
        If Not IsNull(!MensualDeudaOtarsCuentas) Then mProps.MensualDeudaOtarsCuentas = !MensualDeudaOtarsCuentas
        
        'ref familiar
        If Not IsNull(!NombreFami) Then mProps.NombreFami = !NombreFami
        If Not IsNull(!RelacionFami) Then mProps.RelacionFami = !RelacionFami
        If Not IsNull(!TelfFami) Then mProps.TelfFami = !TelfFami
        If Not IsNull(!CiudadFami) Then mProps.CiudadFami = !CiudadFami
        
        If Not IsNull(!NombreFami2) Then mProps.NombreFami2 = !NombreFami2
        If Not IsNull(!RelacionFami2) Then mProps.RelacionFami2 = !RelacionFami2
        If Not IsNull(!TelfFami2) Then mProps.TelfFami2 = !TelfFami2
        If Not IsNull(!CiudadFami2) Then mProps.CiudadFami2 = !CiudadFami2
        
        
        'ref comer
        If Not IsNull(!NombreComer1) Then mProps.NombreComer1 = !NombreComer1
        If Not IsNull(!DirComer1) Then mProps.DirComer1 = !DirComer1
        If Not IsNull(!TelfComer1) Then mProps.TelfComer1 = !TelfComer1
        If Not IsNull(!NombreComer2) Then mProps.NombreComer2 = !NombreComer2
        If Not IsNull(!DirComer2) Then mProps.DirComer2 = !DirComer2
        If Not IsNull(!TelfComer2) Then mProps.TelfComer2 = !TelfComer2
        If Not IsNull(!NombreComer3) Then mProps.NombreComer3 = !NombreComer3
        If Not IsNull(!DirComer3) Then mProps.DirComer3 = !DirComer3
        If Not IsNull(!TelfComer3) Then mProps.TelfComer3 = !TelfComer3
        'ref.personal
        If Not IsNull(!NombrePerso1) Then mProps.NombrePerso1 = !NombrePerso1
        If Not IsNull(!DirPerso1) Then mProps.DirPerso1 = !DirPerso1
        If Not IsNull(!TelfPerso1) Then mProps.TelfPerso1 = !TelfPerso1
        If Not IsNull(!NombrePerso2) Then mProps.NombrePerso2 = !NombrePerso2
        If Not IsNull(!DirPerso2) Then mProps.DirPerso2 = !DirPerso2
        If Not IsNull(!TelfPerso2) Then mProps.TelfPerso2 = !TelfPerso2
        If Not IsNull(!NombrePerso3) Then mProps.NombrePerso3 = !NombrePerso3
        If Not IsNull(!DirPerso3) Then mProps.DirPerso3 = !DirPerso3
        If Not IsNull(!TelfPerso3) Then mProps.TelfPerso3 = !TelfPerso3
        'ref bancaria
        If Not IsNull(!NombreBanca1) Then mProps.NombreBanca1 = !NombreBanca1
        If Not IsNull(!NumCtaBanca1) Then mProps.NumCtaBanca1 = !NumCtaBanca1
        If Not IsNull(!CodTipoBanca1) Then mProps.CodTipoBanca1 = !CodTipoBanca1
        If Not IsNull(!NombreBanca2) Then mProps.NombreBanca2 = !NombreBanca2
        If Not IsNull(!NumCtaBanca2) Then mProps.NumCtaBanca2 = !NumCtaBanca2
        If Not IsNull(!CodTipoBanca2) Then mProps.CodTipoBanca2 = !CodTipoBanca2
        If Not IsNull(!NombreBanca3) Then mProps.NombreBanca3 = !NombreBanca3
        If Not IsNull(!NumCtaBanca3) Then mProps.NumCtaBanca3 = !NumCtaBanca3
        If Not IsNull(!CodTipoBanca3) Then mProps.CodTipoBanca3 = !CodTipoBanca3

        If Not IsNull(!IdInventario) Then mProps.IdInventario = !IdInventario
        If Not IsNull(!CodInventario) Then mProps.CodInventario = !CodInventario
        If Not IsNull(!BandAprobado) Then mProps.BandAprobado = !BandAprobado
        If Not IsNull(!FechaSolicitud) Then mProps.FechaSolicitud = !FechaSolicitud
        If Not IsNull(!BandCliente) Then mProps.BandCliente = !BandCliente
        If Not IsNull(!MontoCredito) Then mProps.MontoCredito = !MontoCredito
        If Not IsNull(!IdTransFuente) Then mProps.IdTransFuente = !IdTransFuente
        If Not IsNull(!TransFuente) Then mProps.TransFuente = !TransFuente
        If Not IsNull(!TipoSolicitud) Then mProps.TipoSolicitud = !TipoSolicitud
        
        mProps.FechaGrabado = !FechaGrabado
        
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub
Friend Sub RecuperarCliGar(key As Variant)
    Dim sql As String, i As Long
    Dim rs As Recordset
    
    sql = "Select pcs.*,(gc.CodTrans+'-'+CONVERT(VARCHAR,gc.NumTrans)) AS TransFuente,pc.codprovcli,pc2.codprovcli as codgar,e.codProvcli as codEmpleado "
    sql = sql & " FROM PCSolCreProvCli pcs inner join PCPROVCLI PC on pc.idprovcli = pcs.idprovcli left Join pcprovcli pc2 on  "
    sql = sql & " pc2.idprovcli = pcs.idgar left join gncomprobante gc  on gc.transid=pcs.IdTransFuente"
    sql = sql & " Left Join Empleado e on e.idprovcli = pcs.idEmpleado"
     sql = sql & " Where pcs.Id = " & key
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.Id = !Id
        If Not IsNull(!IdProvCli) Then mProps.IdProvCli = !IdProvCli
        If Not IsNull(!CodProvCli) Then mProps.CodProvCli = !CodProvCli
        If Not IsNull(!IdGar) Then mProps.IdGar = !IdGar
        If Not IsNull(!CodGar) Then mProps.CodGar = !CodGar
        
        If Not IsNull(!CodEstadoCivil) Then mProps.CodEstadoCivil = !CodEstadoCivil
        If Not IsNull(!CodInstruccion) Then mProps.CodInstruccion = !CodInstruccion
        If Not IsNull(!BandPatriFam) Then mProps.BandPatriFam = !BandPatriFam
        If Not IsNull(!BandSepBienes) Then mProps.BandSepBienes = !BandSepBienes
        If Not IsNull(!TipoVivienda) Then mProps.TipoVivienda = !TipoVivienda
        If Not IsNull(!AnioResidencia) Then mProps.AnioResidencia = !AnioResidencia
        If Not IsNull(!NombreArrend) Then mProps.NombreArrend = !NombreArrend
        If Not IsNull(!TelfArrend) Then mProps.TelfArrend = !TelfArrend
        'solicitante
        If Not IsNull(!ActividadSoli) Then mProps.ActividadSoli = !ActividadSoli
        If Not IsNull(!NombreTrabajoSoli) Then mProps.NombreTrabajoSoli = !NombreTrabajoSoli
        If Not IsNull(!DirTrabajoSoli) Then mProps.DirTrabajoSoli = !DirTrabajoSoli
        If Not IsNull(!AntiguedadSoli) Then mProps.AntiguedadSoli = !AntiguedadSoli
        If Not IsNull(!TelfActividadSol) Then mProps.TelfActividadSol = !TelfActividadSol
        If Not IsNull(!TipoLocalSoli) Then mProps.TipoLocalSoli = !TipoLocalSoli
        If Not IsNull(!NombreArrendSoli) Then mProps.NombreArrendSoli = !NombreArrendSoli
        If Not IsNull(!TelfArrendSoli) Then mProps.TelfArrendSoli = !TelfArrendSoli
        If Not IsNull(!IngresosSoli) Then mProps.IngresosSoli = !IngresosSoli
        If Not IsNull(!OtrosIngresosSoli) Then mProps.OtrosIngresosSoli = !OtrosIngresosSoli
        If Not IsNull(!GastosFamSoli) Then mProps.GastosFamSoli = !GastosFamSoli
        If Not IsNull(!GastosArriendoSoli) Then mProps.GastosArriendoSoli = !GastosArriendoSoli
        If Not IsNull(!OtrosGastosSoli) Then mProps.OtrosGastosSoli = !OtrosGastosSoli
        'conyuge
        If Not IsNull(!NombreConyu) Then mProps.NombreConyu = !NombreConyu
        If Not IsNull(!CIConyu) Then mProps.CIConyu = !CIConyu
        If Not IsNull(!FechaNacimientoConyu) Then mProps.FechaNacimientoConyu = !FechaNacimientoConyu
        If Not IsNull(!NumcargasConyu) Then mProps.NumcargasConyu = !NumcargasConyu
        If Not IsNull(!ActividadConyu) Then mProps.ActividadConyu = !ActividadConyu
        If Not IsNull(!NombreTrabajoConyu) Then mProps.NombreTrabajoConyu = !NombreTrabajoConyu
        If Not IsNull(!DirTrabajoConyu) Then mProps.DirTrabajoConyu = !DirTrabajoConyu
        If Not IsNull(!AntiguedadConyu) Then mProps.AntiguedadConyu = !AntiguedadConyu
        If Not IsNull(!TelfActividadCounyu) Then mProps.TelfActividadCounyu = !TelfActividadCounyu
        If Not IsNull(!TipoLocalConyu) Then mProps.TipoLocalConyu = !TipoLocalConyu
        If Not IsNull(!NombreArrendConyu) Then mProps.NombreArrendConyu = !NombreArrendConyu
        If Not IsNull(!TelfArrendConyu) Then mProps.TelfArrendConyu = !TelfArrendConyu
        If Not IsNull(!IngresosConyu) Then mProps.IngresosConyu = !IngresosConyu
        If Not IsNull(!OtrosIngresosConyu) Then mProps.OtrosIngresosConyu = !OtrosIngresosConyu
        If Not IsNull(!GastosFamConyu) Then mProps.GastosFamConyu = !GastosFamConyu
        If Not IsNull(!GastosArriendoConyu) Then mProps.GastosArriendoConyu = !GastosArriendoConyu
        If Not IsNull(!OtrosGastosConyu) Then mProps.OtrosGastosConyu = !OtrosGastosConyu
        'inmuebles
        If Not IsNull(!Inmueble1) Then mProps.Inmueble1 = !Inmueble1
        If Not IsNull(!Ubicacion1) Then mProps.Ubicacion1 = !Ubicacion1
        If Not IsNull(!ValorInmueble1) Then mProps.ValorInmueble1 = !ValorInmueble1
        If Not IsNull(!BandHipoteca1) Then mProps.BandHipoteca1 = !BandHipoteca1
        If Not IsNull(!Inmueble2) Then mProps.Inmueble2 = !Inmueble2
        If Not IsNull(!Ubicacion2) Then mProps.Ubicacion2 = !Ubicacion2
        If Not IsNull(!ValorInmueble2) Then mProps.ValorInmueble2 = !ValorInmueble2
        If Not IsNull(!BandHipoteca2) Then mProps.BandHipoteca2 = !BandHipoteca2
        If Not IsNull(!Inmueble3) Then mProps.Inmueble3 = !Inmueble3
        If Not IsNull(!Ubicacion3) Then mProps.Ubicacion3 = !Ubicacion3
        If Not IsNull(!ValorInmueble3) Then mProps.ValorInmueble3 = !ValorInmueble3
        If Not IsNull(!BandHipoteca3) Then mProps.BandHipoteca3 = !BandHipoteca3
        
        'vehiculo
        If Not IsNull(!Vehiculo1) Then mProps.Vehiculo1 = !Vehiculo1
        If Not IsNull(!MarcaPlaca1) Then mProps.MarcaPlaca1 = !MarcaPlaca1
        If Not IsNull(!ValorVehiculo1) Then mProps.ValorVehiculo1 = !ValorVehiculo1
        If Not IsNull(!BandPrendado1) Then mProps.BandPrendado1 = !BandPrendado1
        If Not IsNull(!Vehiculo2) Then mProps.Vehiculo2 = !Vehiculo2
        If Not IsNull(!MarcaPlaca2) Then mProps.MarcaPlaca2 = !MarcaPlaca2
        If Not IsNull(!ValorVehiculo2) Then mProps.ValorVehiculo2 = !ValorVehiculo2
        If Not IsNull(!BandPrendado2) Then mProps.BandPrendado2 = !BandPrendado2
        If Not IsNull(!Vehiculo3) Then mProps.Vehiculo3 = !Vehiculo3
        If Not IsNull(!MarcaPlaca3) Then mProps.MarcaPlaca3 = !MarcaPlaca3
        If Not IsNull(!ValorVehiculo3) Then mProps.ValorVehiculo3 = !ValorVehiculo3
        If Not IsNull(!BandPrendado3) Then mProps.BandPrendado3 = !BandPrendado3
        'pasivo
        If Not IsNull(!MontoDeudaBancaria) Then mProps.MontoDeudaBancaria = !MontoDeudaBancaria
        If Not IsNull(!SaldoDeudaBancaria) Then mProps.SaldoDeudaBancaria = !SaldoDeudaBancaria
        If Not IsNull(!MensualDeudaBancaria) Then mProps.MensualDeudaBancaria = !MensualDeudaBancaria
        If Not IsNull(!MontoHipoteca) Then mProps.MontoHipoteca = !MontoHipoteca
        If Not IsNull(!SaldoHipoteca) Then mProps.SaldoHipoteca = !SaldoHipoteca
        If Not IsNull(!MensualHipoteca) Then mProps.MensualHipoteca = !MensualHipoteca
        If Not IsNull(!MontoDeudaOtarsCuentas) Then mProps.MontoDeudaOtarsCuentas = !MontoDeudaOtarsCuentas
        If Not IsNull(!SaldoDeudaOtarsCuentas) Then mProps.SaldoDeudaOtarsCuentas = !SaldoDeudaOtarsCuentas
        If Not IsNull(!MensualDeudaOtarsCuentas) Then mProps.MensualDeudaOtarsCuentas = !MensualDeudaOtarsCuentas
        
        'ref familiar
        If Not IsNull(!NombreFami) Then mProps.NombreFami = !NombreFami
        If Not IsNull(!RelacionFami) Then mProps.RelacionFami = !RelacionFami
        If Not IsNull(!TelfFami) Then mProps.TelfFami = !TelfFami
        If Not IsNull(!CiudadFami) Then mProps.CiudadFami = !CiudadFami
        
        If Not IsNull(!NombreFami2) Then mProps.NombreFami2 = !NombreFami2
        If Not IsNull(!RelacionFami2) Then mProps.RelacionFami2 = !RelacionFami2
        If Not IsNull(!TelfFami2) Then mProps.TelfFami2 = !TelfFami2
        If Not IsNull(!CiudadFami2) Then mProps.CiudadFami2 = !CiudadFami2
        
        
        'ref comer
        If Not IsNull(!NombreComer1) Then mProps.NombreComer1 = !NombreComer1
        If Not IsNull(!DirComer1) Then mProps.DirComer1 = !DirComer1
        If Not IsNull(!TelfComer1) Then mProps.TelfComer1 = !TelfComer1
        If Not IsNull(!NombreComer2) Then mProps.NombreComer2 = !NombreComer2
        If Not IsNull(!DirComer2) Then mProps.DirComer2 = !DirComer2
        If Not IsNull(!TelfComer2) Then mProps.TelfComer2 = !TelfComer2
        If Not IsNull(!NombreComer3) Then mProps.NombreComer3 = !NombreComer3
        If Not IsNull(!DirComer3) Then mProps.DirComer3 = !DirComer3
        If Not IsNull(!TelfComer3) Then mProps.TelfComer3 = !TelfComer3
        'ref.personal
        If Not IsNull(!NombrePerso1) Then mProps.NombrePerso1 = !NombrePerso1
        If Not IsNull(!DirPerso1) Then mProps.DirPerso1 = !DirPerso1
        If Not IsNull(!TelfPerso1) Then mProps.TelfPerso1 = !TelfPerso1
        If Not IsNull(!NombrePerso2) Then mProps.NombrePerso2 = !NombrePerso2
        If Not IsNull(!DirPerso2) Then mProps.DirPerso2 = !DirPerso2
        If Not IsNull(!TelfPerso2) Then mProps.TelfPerso2 = !TelfPerso2
        If Not IsNull(!NombrePerso3) Then mProps.NombrePerso3 = !NombrePerso3
        If Not IsNull(!DirPerso3) Then mProps.DirPerso3 = !DirPerso3
        If Not IsNull(!TelfPerso3) Then mProps.TelfPerso3 = !TelfPerso3
        'ref bancaria
        If Not IsNull(!NombreBanca1) Then mProps.NombreBanca1 = !NombreBanca1
        If Not IsNull(!NumCtaBanca1) Then mProps.NumCtaBanca1 = !NumCtaBanca1
        If Not IsNull(!CodTipoBanca1) Then mProps.CodTipoBanca1 = !CodTipoBanca1
        If Not IsNull(!NombreBanca2) Then mProps.NombreBanca2 = !NombreBanca2
        If Not IsNull(!NumCtaBanca2) Then mProps.NumCtaBanca2 = !NumCtaBanca2
        If Not IsNull(!CodTipoBanca2) Then mProps.CodTipoBanca2 = !CodTipoBanca2
        If Not IsNull(!NombreBanca3) Then mProps.NombreBanca3 = !NombreBanca3
        If Not IsNull(!NumCtaBanca3) Then mProps.NumCtaBanca3 = !NumCtaBanca3
        If Not IsNull(!CodTipoBanca3) Then mProps.CodTipoBanca3 = !CodTipoBanca3

        'If Not IsNull(!IdInventario) Then mProps.IdInventario = !IdInventario
        'If Not IsNull(!CodInventario) Then mProps.CodInventario = !CodInventario
        If Not IsNull(!BandAprobado) Then mProps.BandAprobado = !BandAprobado
        If Not IsNull(!FechaSolicitud) Then mProps.FechaSolicitud = !FechaSolicitud
        If Not IsNull(!BandCliente) Then mProps.BandCliente = !BandCliente
        If Not IsNull(!MontoCredito) Then mProps.MontoCredito = !MontoCredito
        If Not IsNull(!IdTransFuente) Then mProps.IdTransFuente = !IdTransFuente
        If Not IsNull(!TransFuente) Then mProps.TransFuente = !TransFuente
        If Not IsNull(!TipoSolicitud) Then mProps.TipoSolicitud = !TipoSolicitud
        If Not IsNull(!IdEmpleado) Then mProps.IdEmpleado = !IdEmpleado
        If Not IsNull(!CodEmpleado) Then mProps.CodEmpleado = !CodEmpleado
        If Not IsNull(!CodUsuarioAutoriza) Then mProps.CodUsuarioAutoriza = !CodUsuarioAutoriza
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Property Get IdGar() As Long
    IdGar = mProps.IdGar
End Property



Public Property Let CodGar(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodGar Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaContable también
            sql = "SELECT IdProvCli FROM PcProvCli WHERE CodProvcli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli.CodProvCli", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdGar = rs!IdProvCli
            End If
            rs.Close
            Set rs = Nothing
        Else
            mProps.IdGar = 0
        End If
    End If
    mProps.CodGar = Trim(value)
    mbooModificado = True
End Property


Public Property Get CodGar() As String
    CodGar = Trim$(mProps.CodGar)
End Property

Public Property Let CodEmpleado(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodEmpleado Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaContable también
            sql = "SELECT IdProvCli FROM Empleado WHERE CodProvcli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Empleado.CodProvCli", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdEmpleado = rs!IdProvCli
            End If
            rs.Close
            Set rs = Nothing
        Else
            mProps.IdEmpleado = 0
        End If
    End If
    mProps.CodEmpleado = Trim(value)
    mbooModificado = True
End Property


Public Property Get CodEmpleado() As String
    CodEmpleado = Trim$(mProps.CodEmpleado)
End Property

Public Property Let CodUsuarioAutoriza(value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli.CodUsuarioAutoriza", MSGERR_INVALIDO
    End If

    mProps.CodUsuarioAutoriza = value
    mbooModificado = True
End Property

Public Property Get CodUsuarioAutoriza() As String
    CodUsuarioAutoriza = Trim$(mProps.CodUsuarioAutoriza)
End Property
