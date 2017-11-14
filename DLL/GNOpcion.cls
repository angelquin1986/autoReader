VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GNOpcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Empresa"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type T_PROPS
    NombreEmpresa As String
    Direccion1 As String
    Direccion2 As String
    Telefono1 As String
    Telefono2 As String
    Telefono3 As String
    Fax1 As String
    Fax2 As String
    RUC As String
    EMail As String
    LogoTipo As String          'Pendiente el tipo de datos
    FechaInicio As Date
    FechaFinal As Date
    CodAsientoSI As Long
    CodAsientoFinal As Long
    IdCuentaResultado As Long
    CodCuentaResultado As String
    EtiqGrupo(1 To 6) As String
    EtiqPCGrupo(1 To 4) As String 'AUC 21/09/2005 cambiado antes 2
    FechaLimiteDesde As Date
    FechaLimiteHasta As Date
    CodMoneda(1 To 4) As String
    NombreMoneda(1 To 4) As String
    FormatoMoneda(1 To 4) As String
    FormatoCantidad As String
    FormatoFecha As String
    Costeo As Byte              '0/1/2/3 PROM,FIFO,LIFO,ULTIMO
    
    GeneraCodItem As Byte
    GeneraCodProv As Byte
    GeneraCodCli As Byte
    GeneraCodCC As Byte
    GeneraCodEmp As Byte
    GeneraCodEmpLosep As Byte
    
    PrefijoCodItem As String
    PrefijoCodProv As String
    PrefijoCodCli As String
    PrefijoCodEmp As String
    PrefijoCodEmpLosep As String
    
    PrefijoCodCC As String
    
    IdCuentaProvPre As Long
    CodCuentaProvPre As String
    IdCuentaCliPre As Long
    CodCuentaCliPre As String
    
    PermiteItemSinCuenta As Boolean
    OcultaContable As Boolean
    colOpciones As Collection
    fechagrabado As Date
    Los100DocCobroPago As Boolean       '  ***  Agregado Oliver 2/08/2001   ***         '
                                            '  para sacar los Primero 100 registros en      '
                                            '  DocCobroPago en caso de haber muchisimos y   '
                                            '  ocurrir error de tiempo de espera            '
    TipoEmpresa As String    ' jeaa 17/12/2005
    PorcentajeIVA As Currency          ' jeaa 08/03/2006 cambio para SRI
    PorcentajeIVAAnt As Currency          ' jeaa 08/03/2006 cambio para SRI
    FechaIVA As Date
    FormatoCodItem As String            'jeaa 27/07/2007
    FormatoCodProv As String
    FormatoCodCli As String
    FormatoCodEmp As String
    FormatoCodEmpLosep As String
    FormatoCodCC As String
    EtiqEst(1 To 2) As String 'AUC 21/10/2007
    NumAutorizacion_AutoImp As String  'jeaa 22/09/2008 requisito SRI
    FechaCaducidad_AutoImp As Date 'jeaa 22/09/2008 requisito SRI
    NumAutorizacion_AutoImpOld As String  'jeaa 22/09/2008 requisito SRI
    FechaCaducidad_AutoImpOld As Date 'jeaa 22/09/2008 requisito SRI
    GeneraCodAFItem As Byte         'jeaa 12/09/2008
    PrefijoCodAFItem As String      'jeaa 12/09/2008
    FormatoCodAFItem As String      'jeaa 12/09/2008
    EtiqAFGrupo(1 To 5) As String   'jeaa 12/09/2008
    RazonSocial  As String  'jeaa 15/01/2009 SRI
    EtiqGNCGrupo(1 To 4) As String   'jeaa 21/03/2009
    'jeaa 24/03/2009
    IdCuentaActivoPre As Long
    CodCuentaActivoPre As String
    IdCuentaCostoPre As Long
    CodCuentaCostoPre As String
    IdCuentaVentaPre As Long
    CodCuentaVentaPre As String
    EtiqGNVGrupo(1 To 4) As String   'jeaa 30/03/2009
    IdCuentaSobranteCajaPre As Long
    CodCuentaSobranteCajaPre As String
    FechaAutorizacion_AutoImp As Date 'jeaa 24/11/2010 requisito SRI
    FechaAutorizacion_AutoImpOld As Date 'jeaa 24/11/2010 requisito SRI
    MontoMaximoCF As Currency          ' jeaa 08/03/2006 cambio para SRI
    RutaXMLValido As String
    RutaXMLAnulado As String
    Resolucion   As String
    TipoEmpresaSRI As String    ' jeaa 06/04/2011
    AutoImpresor As Boolean
    UtilidadMarkDown As Boolean
    IdProvinciaPre As Long
    CodProvinciaPre As String
    IdCantonPre As Long
    CodCantonPre As String
    EtiqPCGrupoP(1 To 4) As String 'AUC 21/09/2005 cambiado antes 2
    EtiqPCGrupoC(1 To 4) As String 'AUC 21/09/2005 cambiado antes 2
    EtiqPCGrupoG(1 To 4) As String 'AUC 21/09/2005 cambiado antes 2
    EtiqPCGrupoE(1 To 4) As String 'AUC 21/09/2005 cambiado antes 2
    EtiqPCGrupoL(1 To 4) As String 'AUC 21/09/2005 cambiado antes 2
    GeneraCodAFItemALT As Byte         'jeaa 12/09/2008
    PrefijoCodAFItemALT As String      'jeaa 12/09/2008
    FormatoCodAFItemALT As String      'jeaa 12/09/2008
    VersionSQL As String
    NumeroLicencias As Integer
    CodigoSC As String
     IdCuentaDiferidaPre As Long
    CodCuentaDiferidaPre As String
    BandDINARDAP As Boolean
    BandCHP As Boolean
    OrdenBodegas As String
    BandOrdenItemsCodigo As Boolean
    BandOrdenGrupoItemsCodigo As Boolean
    BandOrdenPCCodigo As Boolean
    BandOrdenPCGrupoCodigo As Boolean
    BandFactElect As Boolean
'/* GRV 11/AGO/2014 - INICIO */
    ComprobantesGenerados As String
    ComprobantesFirmados As String
    ComprobantesAutorizados As String
    ComprobantesNoAutorizados As String
    ComprobantesEnviados As String
    ComprobantesSubidos As String
    BandPortal As Boolean
    
    UbicacionArchivoToken As String
    ContrasenaToken As String
    TipoAmbiente As String
    ServidorCorreo  As String
    CuentaCorreo As String
    PasswordCorreo As String
    MensajeCorreo As String
   
    PuertoCorreo As String
    CopiaCorreo As String
    NombreUsuario As String
    BandConexionSegura As Boolean
    BandCorreoAutomatico As Boolean
    BandCorreoxEquipo As Boolean
    PenalizaDesctoxItem As String

    '/* GRV 11/AGO/2014 - FIN */
    BandRebateIVG As Boolean
    NumIVGRebate As Integer
    BandOmitirCambioIVA As Boolean
    BandCostoUCParaCosto As Boolean
    
    BandConfigLira As Boolean
    
    BandPorDescGeneral As Boolean
    PorDescGeneral As Currency
    BandDbleClickNOMod As Boolean
    'balgran
    IVKTipoDatoDouble As Boolean
    MensajeCorreoCartera As String
    RutaRepCarteraGen As String
    MensajePDFCartera As String
End Type

Private mProps As T_PROPS
Private mbooModificado As Boolean
Private mbooEsNuevo As Boolean
Private mobjEmpresa As Empresa      'Referencia al objeto padre


Public Sub Grabar()
    Dim sql As String, i As Integer
    Dim rs As Recordset
    Dim NumAutOld As String, FechaCadOld As Date, FechaAutold As Date
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub

    sql = "SELECT * FROM GNOpcion"
    Set rs = mobjEmpresa.OpenRecordsetParaEdit(sql)
    
    If mbooEsNuevo Then
        rs.AddNew
    Else
#If DAOLIB Then
        rs.Edit
#End If
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "GNOpcion.Grabar", MSGERR_MODIFICADO
        End If
    End If


    If Not VerificaDocumento(mProps.RUC) Then
        Err.Raise ERR_RUCINCORRECTO, "GNOpcion.Grabar", MSGERR_RUCINCORRECTO
        Exit Sub
    End If

    'Graba en la base
    With rs
        If Not !NombreEmpresa = mProps.NombreEmpresa Then
            mobjEmpresa.GrabaAccionModulo !NombreEmpresa, mProps.NombreEmpresa, "NombreEmpresa", "MOD_INFEMP", "GN"
        End If
        !NombreEmpresa = mProps.NombreEmpresa
        If Not !Direccion1 = mProps.Direccion1 Then
            mobjEmpresa.GrabaAccionModulo !Direccion1, mProps.Direccion1, "Direccion1", "MOD_INFEMP", "GN"
        End If
        !Direccion1 = mProps.Direccion1
        If Not !Direccion2 = mProps.Direccion2 Then
            mobjEmpresa.GrabaAccionModulo !Direccion2, mProps.Direccion2, "Direccion2", "MOD_INFEMP", "GN"
        End If
        !Direccion2 = mProps.Direccion2
        If Not !Telefono1 = mProps.Telefono1 Then
            mobjEmpresa.GrabaAccionModulo !Telefono1, mProps.Telefono1, "Telefono1", "MOD_INFEMP", "GN"
        End If
        !Telefono1 = mProps.Telefono1
        If Not !Telefono2 = mProps.Telefono2 Then
            mobjEmpresa.GrabaAccionModulo !Telefono2, mProps.Telefono2, "Telefono2", "MOD_INFEMP", "GN"
        End If
        !Telefono2 = mProps.Telefono2
        If Not !Telefono3 = mProps.Telefono3 Then
            mobjEmpresa.GrabaAccionModulo !Telefono3, mProps.Telefono3, "Telefono3", "MOD_INFEMP", "GN"
        End If
        !Telefono3 = mProps.Telefono3
        If Not !Fax1 = mProps.Fax1 Then
            mobjEmpresa.GrabaAccionModulo !Fax1, mProps.Fax1, "Fax1", "MOD_INFEMP", "GN"
        End If
        !Fax1 = mProps.Fax1
        If Not !Fax2 = mProps.Fax2 Then
            mobjEmpresa.GrabaAccionModulo !Fax2, mProps.Fax2, "Fax2", "MOD_INFEMP", "GN"
        End If
        !Fax2 = mProps.Fax2
        If Not !RUC = mProps.RUC Then
            mobjEmpresa.GrabaAccionModulo !RUC, mProps.RUC, "RUC", "MOD_INFEMP", "GN"
        End If
        !RUC = mProps.RUC
        If Not !EMail = mProps.EMail Then
            mobjEmpresa.GrabaAccionModulo !EMail, mProps.EMail, "EMail", "MOD_INFEMP", "GN"
        End If
        If Not !RazonSocial = mProps.RazonSocial Then
            mobjEmpresa.GrabaAccionModulo !RazonSocial, mProps.RazonSocial, "RazonSocial", "MOD_INFEMP", "GN"
        End If
        
        !RazonSocial = mProps.RazonSocial
        
        !EMail = mProps.EMail
'        !LogoTipo = mProps.LogoTipo         'Pendiente
        !FechaInicio = mProps.FechaInicio
        !FechaFinal = mProps.FechaFinal
        !CodAsientoSI = mProps.CodAsientoSI
        !CodAsientoFinal = mProps.CodAsientoFinal
        !IdCuentaResultado = mProps.IdCuentaResultado
        For i = 1 To IVGRUPO_MAX
            .Fields("EtiquetaGrupo" & i) = mProps.EtiqGrupo(i)
        Next i
        For i = 1 To PCGRUPO_MAX
            .Fields("EtiquetaPCGrupo" & i) = mProps.EtiqPCGrupo(i)
        Next i
        If Not !FechaLimiteDesde = mProps.FechaLimiteDesde Then
            mobjEmpresa.GrabaAccionModulo !FechaLimiteDesde, mProps.FechaLimiteDesde, "FechaLimiteDesde", "MOD_INFEMP", "GN"
        End If
        !FechaLimiteDesde = mProps.FechaLimiteDesde
        
        If Not !FechaLimiteHasta = mProps.FechaLimiteHasta Then
            mobjEmpresa.GrabaAccionModulo !FechaLimiteHasta, mProps.FechaLimiteHasta, "FechaLimiteHasta", "MOD_INFEMP", "GN"
        End If
        !FechaLimiteHasta = mProps.FechaLimiteHasta
        For i = 1 To 4
            .Fields("CodMoneda" & i) = mProps.CodMoneda(i)
            .Fields("NombreMoneda" & i) = mProps.NombreMoneda(i)
            .Fields("FormatoMoneda" & i) = mProps.FormatoMoneda(i)
        Next i
        !FormatoCantidad = mProps.FormatoCantidad
        !FormatoFecha = mProps.FormatoFecha
        !Costeo = mProps.Costeo
        
        !GeneraCodItem = mProps.GeneraCodItem
        !GeneraCodProv = mProps.GeneraCodProv
        !GeneraCodCli = mProps.GeneraCodCli
        !GeneraCodEmp = mProps.GeneraCodEmp
        !GeneraCodCC = mProps.GeneraCodCC
        !PrefijoCodItem = mProps.PrefijoCodItem
        !PrefijoCodProv = mProps.PrefijoCodProv
        !PrefijoCodCli = mProps.PrefijoCodCli
        !PrefijoCodEmp = mProps.PrefijoCodEmp
        !PrefijoCodCC = mProps.PrefijoCodCC
        !IdCuentaProvPre = mProps.IdCuentaProvPre
        !IdCuentaCliPre = mProps.IdCuentaCliPre
        
        
        !PermiteItemSinCuenta = mProps.PermiteItemSinCuenta
        !Los100DocCobroPago = mProps.Los100DocCobroPago
        !OcultaContable = mProps.OcultaContable
        
        If Not !TipoEmpresaSRI = mProps.TipoEmpresaSRI Then
            mobjEmpresa.GrabaAccionModulo !TipoEmpresaSRI, mProps.TipoEmpresaSRI, "TipoEmpresa", "MOD_INFEMP", "GN"
        ElseIf IsNull(!TipoEmpresaSRI) Then
            mobjEmpresa.GrabaAccionModulo "VACIO", mProps.TipoEmpresaSRI, "TipoEmpresa", "MOD_INFEMP", "GN"
        End If
        !TipoEmpresaSRI = mProps.TipoEmpresaSRI       ' jeaa 17/12/2005
        !TipoEmpresa = mProps.TipoEmpresa       ' jeaa 17/12/2005
        
        
        If Not !PorcentajeIVA = mProps.PorcentajeIVA Then
            mobjEmpresa.GrabaAccionModulo !PorcentajeIVA, mProps.PorcentajeIVA, "PorcentajeIVA", "MOD_INFEMP", "GN"
            ActulizaIVAItems mProps.PorcentajeIVA
            ActulizaIVAItemsAnt mProps.PorcentajeIVAAnt
            ActulizaIVAPredeterminadoTrans mProps.PorcentajeIVA
        End If
        !PorcentajeIVA = mProps.PorcentajeIVA       ' jeaa 08/03/2006
        !PorcentajeIVAAnt = mProps.PorcentajeIVAAnt       ' jeaa 08/03/2006
        
        If Not !FechaIVA = mProps.FechaIVA Then
            mobjEmpresa.GrabaAccionModulo !FechaIVA, mProps.FechaIVA, "Fecha IVA", "MOD_INFEMP", "GN"
            ActulizaIVAPredeterminadoTrans mProps.PorcentajeIVA
        End If
        !FechaIVA = mProps.FechaIVA       ' jeaa 08/03/2006
        
        !FormatoCodItem = mProps.FormatoCodItem
        !FormatoCodProv = mProps.FormatoCodProv
        !FormatoCodCli = mProps.FormatoCodCli
        !FormatoCodEmp = mProps.FormatoCodEmp
        !FormatoCodCC = mProps.FormatoCodCC
         'AUC 29/10/07
        For i = 1 To 2
                .Fields("EtiquetaEstado" & i) = mProps.EtiqEst(i)
        Next
        !fechagrabado = Now
        NumAutOld = IIf(IsNull(!NumAutorizacion_AutoImp), "9999999999", !NumAutorizacion_AutoImp)
        FechaCadOld = IIf(IsNull(!FechaCaducidad_AutoImp), Date, !FechaCaducidad_AutoImp)
        FechaAutold = IIf(IsNull(!FechaAutorizacion_AutoImp), Date, !FechaAutorizacion_AutoImp)
        
        'jeaa 22/09/2008
        If Not !NumAutorizacion_AutoImp = mProps.NumAutorizacion_AutoImp Then
            mobjEmpresa.GrabaAccionModulo !NumAutorizacion_AutoImp, mProps.NumAutorizacion_AutoImp, "NumAutorizacion_AutoImp", "MOD_INFEMP", "GN"
            !NumAutorizacion_AutoImpOld = NumAutOld
        End If
        !NumAutorizacion_AutoImp = mProps.NumAutorizacion_AutoImp
        If Not !FechaCaducidad_AutoImp = mProps.FechaCaducidad_AutoImp Then
            mobjEmpresa.GrabaAccionModulo !FechaCaducidad_AutoImp, mProps.FechaCaducidad_AutoImp, "FechaCaducidad_AutoImp", "MOD_INFEMP", "GN"
            !FechaCaducidad_AutoImpOld = FechaCadOld
        End If
        !FechaCaducidad_AutoImp = mProps.FechaCaducidad_AutoImp
        
        If Not !FechaAutorizacion_AutoImp = mProps.FechaAutorizacion_AutoImp Then
            mobjEmpresa.GrabaAccionModulo !FechaAutorizacion_AutoImp, mProps.FechaAutorizacion_AutoImp, "FechaAutorizacion_AutoImp", "MOD_INFEMP", "GN"
            !FechaAutorizacion_AutoImpOld = FechaCadOld
        End If
        !FechaAutorizacion_AutoImp = mProps.FechaAutorizacion_AutoImp
        
        
        'jeaa 12/09/2008
        For i = 1 To AFGRUPO_MAX
                .Fields("EtiquetaAFGrupo" & i) = mProps.EtiqAFGrupo(i)
        Next i
        !GeneraCodAFItem = mProps.GeneraCodAFItem
        !PrefijoCodAFItem = mProps.PrefijoCodAFItem
        !FormatoCodAFItem = mProps.FormatoCodAFItem
        'jeaa 21/03/2009
        For i = 1 To GNCGRUPO_MAX
                .Fields("EtiquetaGNCGrupo" & i) = mProps.EtiqGNCGrupo(i)
        Next i
        !IdCuentaActivoPre = mProps.IdCuentaActivoPre
        !IdCuentaCostoPre = mProps.IdCuentaCostoPre
        !IdCuentaVentaPre = mProps.IdCuentaVentaPre
        'jeaa 30/03/2009
        For i = 1 To GNVGRUPO_MAX
                .Fields("EtiquetaGNVGrupo" & i) = mProps.EtiqGNVGrupo(i)
        Next i
        !IdCuentaSobranteCajaPre = mProps.IdCuentaSobranteCajaPre
        
        If Not !MontoMaximoCF = mProps.MontoMaximoCF Then
            mobjEmpresa.GrabaAccionModulo !MontoMaximoCF, mProps.MontoMaximoCF, "Monto Maximo Cons. Final", "MOD_INFEMP", "GN"
        End If
        !MontoMaximoCF = mProps.MontoMaximoCF
        
        If Not !RutaXMLValido = mProps.RutaXMLValido Then
            mobjEmpresa.GrabaAccionModulo !RutaXMLValido, mProps.RutaXMLValido, "RutaXMLValido", "MOD_INFEMP", "GN"
        End If
        !RutaXMLValido = mProps.RutaXMLValido
        If Not !RutaXMLAnulado = mProps.RutaXMLAnulado Then
            mobjEmpresa.GrabaAccionModulo !RutaXMLAnulado, mProps.RutaXMLAnulado, "RutaXMLAnulado", "MOD_INFEMP", "GN"
        End If
        !RutaXMLAnulado = mProps.RutaXMLAnulado
        
        If Not !Resolucion = mProps.Resolucion Then
            mobjEmpresa.GrabaAccionModulo !Resolucion, mProps.Resolucion, "Resolucion", "MOD_INFEMP", "GN"
        End If
        !Resolucion = mProps.Resolucion
        
        If Not !CodigoSC = mProps.CodigoSC Then
            mobjEmpresa.GrabaAccionModulo !CodigoSC, mProps.CodigoSC, "CodigoSC", "MOD_INFEMP", "GN"
        End If
        !CodigoSC = mProps.CodigoSC
        
        
        
        
        !AutoImpresor = mProps.AutoImpresor
        !UtilidadMarkDown = mProps.UtilidadMarkDown
        !IdProvinciaPre = mProps.IdProvinciaPre
        !IdCantonPre = mProps.IdCantonPre
        For i = 1 To PCGRUPO_MAX
            .Fields("EtiquetaPCGrupoP" & i) = mProps.EtiqPCGrupoP(i)
        Next i
        For i = 1 To PCGRUPO_MAX
            .Fields("EtiquetaPCGrupoC" & i) = mProps.EtiqPCGrupoC(i)
        Next i
        For i = 1 To PCGRUPO_MAX
            .Fields("EtiquetaPCGrupoG" & i) = mProps.EtiqPCGrupoG(i)
        Next i
        
        For i = 1 To PCGRUPO_MAX
            .Fields("EtiquetaPCGrupoE" & i) = mProps.EtiqPCGrupoE(i)
        Next i
        
        For i = 1 To PCGRUPO_MAX
            .Fields("EtiquetaPCGrupoL" & i) = mProps.EtiqPCGrupoL(i)
        Next i
        !GeneraCodAFItemALT = mProps.GeneraCodAFItemALT
        !PrefijoCodAFItemALT = mProps.PrefijoCodAFItemALT
        !FormatoCodAFItemALT = mProps.FormatoCodAFItemALT
        !VersionSQL = mProps.VersionSQL
        !xx = mProps.NumeroLicencias
        !BandDINARDAP = mProps.BandDINARDAP
        !BandCHP = mProps.BandCHP
        !OrdenBodegas = mProps.OrdenBodegas
        
        
        !BandOrdenItemsCodigo = mProps.BandOrdenItemsCodigo
        !BandOrdenGrupoItemsCodigo = mProps.BandOrdenGrupoItemsCodigo
        !BandOrdenPCCodigo = mProps.BandOrdenPCCodigo
        !BandOrdenPCGrupoCodigo = mProps.BandOrdenPCGrupoCodigo
        !BandFactElect = mProps.BandFactElect

        !ComprobantesGenerados = mProps.ComprobantesGenerados
        If Not !ComprobantesGenerados = mProps.ComprobantesGenerados Then
            mobjEmpresa.GrabaAccionModulo !ComprobantesGenerados, mProps.ComprobantesGenerados, "ComprobantesGenerados", "MOD_INFEMP", "GN"
        End If
        !ComprobantesAutorizados = mProps.ComprobantesAutorizados
        If Not !ComprobantesAutorizados = mProps.ComprobantesAutorizados Then
            mobjEmpresa.GrabaAccionModulo !ComprobantesAutorizados, mProps.ComprobantesAutorizados, "ComprobantesAutorizados", "MOD_INFEMP", "GN"
        End If
        !ComprobantesFirmados = mProps.ComprobantesFirmados
        If Not !ComprobantesFirmados = mProps.ComprobantesFirmados Then
            mobjEmpresa.GrabaAccionModulo !ComprobantesFirmados, mProps.ComprobantesFirmados, "ComprobantesFirmados", "MOD_INFEMP", "GN"
        End If
        
        !ComprobantesEnviados = mProps.ComprobantesEnviados
        If Not !ComprobantesEnviados = mProps.ComprobantesEnviados Then
            mobjEmpresa.GrabaAccionModulo !ComprobantesEnviados, mProps.ComprobantesEnviados, "ComprobantesEnviados", "MOD_INFEMP", "GN"
        End If
        
        
        !ComprobantesNoAutorizados = mProps.ComprobantesNoAutorizados
        If Not !ComprobantesNoAutorizados = mProps.ComprobantesNoAutorizados Then
            mobjEmpresa.GrabaAccionModulo !ComprobantesNoAutorizados, mProps.ComprobantesNoAutorizados, "ComprobantesNoAutorizados", "MOD_INFEMP", "GN"
        End If
        
        
        !UbicacionArchivoToken = mProps.UbicacionArchivoToken
        !ContrasenaToken = mProps.ContrasenaToken
        !TipoAmbiente = mProps.TipoAmbiente
        
        !ServidorCorreo = mProps.ServidorCorreo
        !CuentaCorreo = mProps.CuentaCorreo
        !PasswordCorreo = mProps.PasswordCorreo
        
        !MensajeCorreo = mProps.MensajeCorreo
        If Not !MensajeCorreo = mProps.MensajeCorreo Then
            mobjEmpresa.GrabaAccionModulo !MensajeCorreo, mProps.MensajeCorreo, "MensajeCorreo", "MOD_INFEMP", "GN"
        End If
        
        !MensajeCorreoCartera = mProps.MensajeCorreoCartera
        If Not !MensajeCorreoCartera = mProps.MensajeCorreoCartera Then
            mobjEmpresa.GrabaAccionModulo !MensajeCorreoCartera, mProps.MensajeCorreoCartera, "MensajeCorreoCartera", "MOD_INFEMP", "GN"
        End If
        
          !MensajePDFCartera = mProps.MensajePDFCartera
        If Not !MensajePDFCartera = mProps.MensajePDFCartera Then
            mobjEmpresa.GrabaAccionModulo !MensajePDFCartera, mProps.MensajePDFCartera, "MensajePDFCartera", "MOD_INFEMP", "GN"
        End If
        
        !RutaRepCarteraGen = mProps.RutaRepCarteraGen
        If Not !RutaRepCarteraGen = mProps.RutaRepCarteraGen Then
            mobjEmpresa.GrabaAccionModulo !RutaRepCarteraGen, mProps.RutaRepCarteraGen, "RutaRepCarteraGen", "MOD_INFEMP", "GN"
        End If
        


        !ComprobantesEnviados = mProps.ComprobantesEnviados
        If Not !ComprobantesEnviados = mProps.ComprobantesEnviados Then
            mobjEmpresa.GrabaAccionModulo !ComprobantesEnviados, mProps.ComprobantesEnviados, "ComprobantesEnviados", "MOD_INFEMP", "GN"
        End If

        !PuertoCorreo = mProps.PuertoCorreo
        If Not !PuertoCorreo = mProps.PuertoCorreo Then
            mobjEmpresa.GrabaAccionModulo !PuertoCorreo, mProps.PuertoCorreo, "PuertoCorreo", "MOD_INFEMP", "GN"
        End If

        !CopiaCorreo = mProps.CopiaCorreo
        If Not !CopiaCorreo = mProps.CopiaCorreo Then
            mobjEmpresa.GrabaAccionModulo !CopiaCorreo, mProps.CopiaCorreo, "CopiaCorreo", "MOD_INFEMP", "GN"
        End If
        
        
        !NombreUsuario = mProps.NombreUsuario
        If Not !NombreUsuario = mProps.NombreUsuario Then
            mobjEmpresa.GrabaAccionModulo !NombreUsuario, mProps.NombreUsuario, "NombreUsuario", "MOD_INFEMP", "GN"
        End If
        !ConexionSegura = mProps.BandConexionSegura
        If Not !ConexionSegura = mProps.BandConexionSegura Then
            mobjEmpresa.GrabaAccionModulo !ConexionSegura, mProps.BandConexionSegura, "ConexionSegura", "MOD_INFEMP", "GN"
        End If

        !ComprobantesSubidos = mProps.ComprobantesSubidos
        If Not !ComprobantesSubidos = mProps.ComprobantesSubidos Then
            mobjEmpresa.GrabaAccionModulo !ComprobantesSubidos, mProps.ComprobantesSubidos, "ComprobantesSubidos", "MOD_INFEMP", "GN"
        End If
        
        !BandCorreoAutomatico = mProps.BandCorreoAutomatico
        If Not !BandCorreoAutomatico = mProps.BandCorreoAutomatico Then
            mobjEmpresa.GrabaAccionModulo !BandCorreoAutomatico, mProps.BandCorreoAutomatico, "BandCorreoAutomatico", "MOD_INFEMP", "GN"
        End If
        
        
        !BandCorreoxEquipo = mProps.BandCorreoxEquipo
        If Not !BandCorreoxEquipo = mProps.BandCorreoxEquipo Then
            mobjEmpresa.GrabaAccionModulo !BandCorreoxEquipo, mProps.BandCorreoxEquipo, "BandCorreoxEquipo", "MOD_INFEMP", "GN"
        End If
        
        
        !BandPortal = mProps.BandPortal
        !PenalizaDesctoxItem = mProps.PenalizaDesctoxItem
        
        !BandRebateIVG = mProps.BandRebateIVG
        !NumIVGRebate = mProps.NumIVGRebate
        !BandOmitirCambioIVA = mProps.BandOmitirCambioIVA
        !BandCostoUCParaCosto = mProps.BandCostoUCParaCosto
        !IVKTipoDatoDouble = mProps.IVKTipoDatoDouble
        !BandConfigLira = mProps.BandConfigLira
        !BandPorDescGeneral = mProps.BandPorDescGeneral
        !PorDescGeneral = mProps.PorDescGeneral
        !BandDbleClickNOMod = mProps.BandDbleClickNOMod
        If Not IsNull(!MensajeCorreoCartera) Then mProps.MensajeCorreoCartera = !MensajeCorreoCartera
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    'Graba en la tabla GNOpcion2    '*** MAKOTO 27/nov/00 Agregado
    GrabarGNOpcion2
    
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mobjEmpresa.RaiseRefreshEvent REFRESH_OPCION
End Sub

Public Sub Recuperar()
    Dim sql As String, i As Long
    Dim rs As Recordset
    
    sql = "SELECT * FROM vwOpcion"
    Set rs = mobjEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        If Not IsNull(!NombreEmpresa) Then mProps.NombreEmpresa = !NombreEmpresa
        If Not IsNull(!Direccion1) Then mProps.Direccion1 = !Direccion1
        If Not IsNull(!Direccion2) Then mProps.Direccion2 = !Direccion2
        If Not IsNull(!Telefono1) Then mProps.Telefono1 = !Telefono1
        If Not IsNull(!Telefono2) Then mProps.Telefono2 = !Telefono2
        If Not IsNull(!Telefono3) Then mProps.Telefono3 = !Telefono3
        If Not IsNull(!Fax1) Then mProps.Fax1 = !Fax1
        If Not IsNull(!Fax2) Then mProps.Fax2 = !Fax2
        If Not IsNull(!RUC) Then mProps.RUC = !RUC
        If Not IsNull(!EMail) Then mProps.EMail = !EMail
        If Not IsNull(!FechaInicio) Then mProps.FechaInicio = !FechaInicio
        If Not IsNull(!FechaFinal) Then mProps.FechaFinal = !FechaFinal
        If Not IsNull(!CodAsientoSI) Then mProps.CodAsientoSI = !CodAsientoSI
        If Not IsNull(!CodAsientoFinal) Then mProps.CodAsientoFinal = !CodAsientoFinal
        If Not IsNull(!IdCuentaResultado) Then mProps.IdCuentaResultado = !IdCuentaResultado
        If Not IsNull(!CodCuentaResultado) Then mProps.CodCuentaResultado = !CodCuentaResultado
        For i = 1 To IVGRUPO_MAX
            If Not IsNull(.Fields("EtiquetaGrupo" & i)) Then mProps.EtiqGrupo(i) = .Fields("EtiquetaGrupo" & i)
        Next i
        For i = 1 To PCGRUPO_MAX
            If Not IsNull(.Fields("EtiquetaPCGrupo" & i)) Then mProps.EtiqPCGrupo(i) = .Fields("EtiquetaPCGrupo" & i)
        Next i
        If Not IsNull(!FechaLimiteDesde) Then mProps.FechaLimiteDesde = !FechaLimiteDesde
        If Not IsNull(!FechaLimiteHasta) Then mProps.FechaLimiteHasta = !FechaLimiteHasta
        For i = 1 To 4
            If Not IsNull(.Fields("CodMoneda" & i)) Then mProps.CodMoneda(i) = .Fields("CodMoneda" & i)
            If Not IsNull(.Fields("NombreMoneda" & i)) Then mProps.NombreMoneda(i) = .Fields("NombreMoneda" & i)
            If Not IsNull(.Fields("FormatoMoneda" & i)) Then mProps.FormatoMoneda(i) = .Fields("FormatoMoneda" & i)
        Next i
        If Not IsNull(!FormatoCantidad) Then mProps.FormatoCantidad = !FormatoCantidad
        If Not IsNull(!FormatoFecha) Then mProps.FormatoFecha = !FormatoFecha
        If Not IsNull(!Costeo) Then mProps.Costeo = !Costeo
        
        If Not IsNull(!GeneraCodItem) Then mProps.GeneraCodItem = !GeneraCodItem
        If Not IsNull(!GeneraCodProv) Then mProps.GeneraCodProv = !GeneraCodProv
        If Not IsNull(!GeneraCodCli) Then mProps.GeneraCodCli = !GeneraCodCli
        If Not IsNull(!GeneraCodEmp) Then mProps.GeneraCodEmp = !GeneraCodEmp
        If Not IsNull(!GeneraCodCC) Then mProps.GeneraCodCC = !GeneraCodCC
        If Not IsNull(!PrefijoCodItem) Then mProps.PrefijoCodItem = !PrefijoCodItem
        If Not IsNull(!PrefijoCodProv) Then mProps.PrefijoCodProv = !PrefijoCodProv
        If Not IsNull(!PrefijoCodCli) Then mProps.PrefijoCodCli = !PrefijoCodCli
        If Not IsNull(!PrefijoCodEmp) Then mProps.PrefijoCodEmp = !PrefijoCodEmp
        If Not IsNull(!PrefijoCodCC) Then mProps.PrefijoCodCC = !PrefijoCodCC
        If Not IsNull(!IdCuentaProvPre) Then mProps.IdCuentaProvPre = !IdCuentaProvPre
        If Not IsNull(!CodCuentaProvPre) Then mProps.CodCuentaProvPre = !CodCuentaProvPre
        If Not IsNull(!IdCuentaCliPre) Then mProps.IdCuentaCliPre = !IdCuentaCliPre
        If Not IsNull(!CodCuentaCliPre) Then mProps.CodCuentaCliPre = !CodCuentaCliPre
        If Not IsNull(!PermiteItemSinCuenta) Then mProps.PermiteItemSinCuenta = !PermiteItemSinCuenta
        If Not IsNull(!Los100DocCobroPago) Then mProps.Los100DocCobroPago = !Los100DocCobroPago
        
        
        If Not IsNull(!OcultaContable) Then mProps.OcultaContable = !OcultaContable
        If Not IsNull(!TipoEmpresa) Then mProps.TipoEmpresa = !TipoEmpresa  ' jeaa 17/12/2005
        If Not IsNull(!TipoEmpresaSRI) Then mProps.TipoEmpresaSRI = !TipoEmpresaSRI  ' jeaa 17/12/2005
        If Not IsNull(!PorcentajeIVA) Then mProps.PorcentajeIVA = Trim$(!PorcentajeIVA)  ' jeaa 08/03/2006
        If Not IsNull(!PorcentajeIVAAnt) Then mProps.PorcentajeIVAAnt = Trim$(!PorcentajeIVAAnt)  ' jeaa 08/03/2006
        If Not IsNull(!FechaIVA) Then mProps.FechaIVA = Trim$(!FechaIVA)  ' jeaa 08/03/2006
        If Not IsNull(!FormatoCodItem) Then mProps.FormatoCodItem = !FormatoCodItem
        If Not IsNull(!FormatoCodProv) Then mProps.FormatoCodProv = !FormatoCodProv
        If Not IsNull(!FormatoCodCli) Then mProps.FormatoCodCli = !FormatoCodCli
        If Not IsNull(!FormatoCodEmp) Then mProps.FormatoCodEmp = !FormatoCodEmp
        If Not IsNull(!FormatoCodCC) Then mProps.FormatoCodCC = !FormatoCodCC
         'AUC 29/10/07
        For i = 1 To 2
            If Not IsNull(.Fields("EtiquetaEstado" & i)) Then mProps.EtiqEst(i) = .Fields("EtiquetaEstado" & i)
        Next
        mProps.fechagrabado = !fechagrabado
        'jeaa 22/09/2008
        If Not IsNull(!NumAutorizacion_AutoImp) Then mProps.NumAutorizacion_AutoImp = !NumAutorizacion_AutoImp
        If Not IsNull(!FechaCaducidad_AutoImp) Then mProps.FechaCaducidad_AutoImp = !FechaCaducidad_AutoImp
        If Not IsNull(!FechaAutorizacion_AutoImp) Then mProps.FechaAutorizacion_AutoImp = !FechaAutorizacion_AutoImp
        If Not IsNull(!NumAutorizacion_AutoImpOld) Then mProps.NumAutorizacion_AutoImpOld = !NumAutorizacion_AutoImpOld
        If Not IsNull(!FechaCaducidad_AutoImpOld) Then mProps.FechaCaducidad_AutoImpOld = !FechaCaducidad_AutoImpOld
        If Not IsNull(!FechaAutorizacion_AutoImpOld) Then mProps.FechaAutorizacion_AutoImpOld = !FechaAutorizacion_AutoImpOld
        'jeaa 09/12/2008
        For i = 1 To AFGRUPO_MAX
            If Not IsNull(.Fields("EtiquetaAFGrupo" & i)) Then mProps.EtiqAFGrupo(i) = .Fields("EtiquetaAFGrupo" & i)
        Next i
        If Not IsNull(!GeneraCodAFItem) Then mProps.GeneraCodAFItem = !GeneraCodAFItem
        If Not IsNull(!PrefijoCodAFItem) Then mProps.PrefijoCodAFItem = !PrefijoCodAFItem
        If Not IsNull(!GeneraCodAFItem) Then mProps.GeneraCodAFItem = !GeneraCodAFItem
        If Not IsNull(!RazonSocial) Then mProps.RazonSocial = !RazonSocial
        'jeaa 21/03/2009
        For i = 1 To GNCGRUPO_MAX
            If Not IsNull(.Fields("Etiquetagncgrupo" & i)) Then mProps.EtiqGNCGrupo(i) = .Fields("EtiquetagncGrupo" & i)
        Next i
        If Not IsNull(!IdCuentaActivoPre) Then mProps.IdCuentaActivoPre = !IdCuentaActivoPre
        If Not IsNull(!CodCuentaActivoPre) Then mProps.CodCuentaActivoPre = !CodCuentaActivoPre
        If Not IsNull(!IdCuentaCostoPre) Then mProps.IdCuentaCostoPre = !IdCuentaCostoPre
        If Not IsNull(!CodCuentaCostoPre) Then mProps.CodCuentaCostoPre = !CodCuentaCostoPre
        If Not IsNull(!IdCuentaVentaPre) Then mProps.IdCuentaVentaPre = !IdCuentaVentaPre
        If Not IsNull(!CodCuentaVentaPre) Then mProps.CodCuentaVentaPre = !CodCuentaVentaPre
        'jeaa 30/03/2009
        For i = 1 To GNVGRUPO_MAX
            If Not IsNull(.Fields("Etiquetagnvgrupo" & i)) Then mProps.EtiqGNVGrupo(i) = .Fields("EtiquetagnvGrupo" & i)
        Next i
        If Not IsNull(!IdCuentaSobranteCajaPre) Then mProps.IdCuentaSobranteCajaPre = !IdCuentaSobranteCajaPre
        If Not IsNull(!CodCuentaSobranteCajaPre) Then mProps.CodCuentaSobranteCajaPre = !CodCuentaSobranteCajaPre
        If Not IsNull(!MontoMaximoCF) Then mProps.MontoMaximoCF = Trim$(!MontoMaximoCF)  ' jeaa 25/02/2011
        If Not IsNull(!RutaXMLValido) Then mProps.RutaXMLValido = Trim$(!RutaXMLValido)  ' jeaa 25/02/2011
        If Not IsNull(!RutaXMLAnulado) Then mProps.RutaXMLAnulado = Trim$(!RutaXMLAnulado)  ' jeaa 25/02/2011
        If Not IsNull(!Resolucion) Then mProps.Resolucion = Trim$(!Resolucion)  ' jeaa 25/02/2011
        If Not IsNull(!AutoImpresor) Then mProps.AutoImpresor = !AutoImpresor
        If Not IsNull(!UtilidadMarkDown) Then mProps.UtilidadMarkDown = !UtilidadMarkDown
        If Not IsNull(!IdProvinciaPre) Then mProps.IdProvinciaPre = !IdProvinciaPre
        If Not IsNull(!CodProvinciaPre) Then mProps.CodProvinciaPre = !CodProvinciaPre
        If Not IsNull(!IdCantonPre) Then mProps.IdCantonPre = !IdCantonPre
        If Not IsNull(!CodCantonPre) Then mProps.CodCantonPre = !CodCantonPre
        For i = 1 To PCGRUPO_MAX
            If Not IsNull(.Fields("EtiquetaPCGrupoP" & i)) Then mProps.EtiqPCGrupoP(i) = .Fields("EtiquetaPCGrupoP" & i)
        Next i
        For i = 1 To PCGRUPO_MAX
            If Not IsNull(.Fields("EtiquetaPCGrupoC" & i)) Then mProps.EtiqPCGrupoC(i) = .Fields("EtiquetaPCGrupoC" & i)
        Next i
        For i = 1 To PCGRUPO_MAX
            If Not IsNull(.Fields("EtiquetaPCGrupoG" & i)) Then mProps.EtiqPCGrupoG(i) = .Fields("EtiquetaPCGrupoG" & i)
        Next i
        For i = 1 To PCGRUPO_MAX
            If Not IsNull(.Fields("EtiquetaPCGrupoE" & i)) Then mProps.EtiqPCGrupoE(i) = .Fields("EtiquetaPCGrupoE" & i)
        Next i

        For i = 1 To PCGRUPO_MAX
            If Not IsNull(.Fields("EtiquetaPCGrupoL" & i)) Then mProps.EtiqPCGrupoL(i) = .Fields("EtiquetaPCGrupoL" & i)
        Next i

        If Not IsNull(!GeneraCodAFItemALT) Then mProps.GeneraCodAFItemALT = !GeneraCodAFItemALT
        If Not IsNull(!PrefijoCodAFItemALT) Then mProps.PrefijoCodAFItemALT = !PrefijoCodAFItemALT
        If Not IsNull(!FormatoCodAFItemALT) Then mProps.FormatoCodAFItemALT = !FormatoCodAFItemALT
        If Not IsNull(!VersionSQL) Then mProps.VersionSQL = !VersionSQL
        If Not IsNull(!xx) Then mProps.NumeroLicencias = !xx Else mProps.NumeroLicencias = "0"
        If Not IsNull(!CodigoSC) Then mProps.CodigoSC = Trim$(!CodigoSC)
        If Not IsNull(!BandDINARDAP) Then mProps.BandDINARDAP = Trim$(!BandDINARDAP)
        If Not IsNull(!BandCHP) Then mProps.BandCHP = Trim$(!BandCHP)
        If Not IsNull(!OrdenBodegas) Then mProps.OrdenBodegas = Trim$(!OrdenBodegas)
        
        
        
        If Not IsNull(!BandOrdenItemsCodigo) Then mProps.BandOrdenItemsCodigo = !BandOrdenItemsCodigo
        If Not IsNull(!BandOrdenGrupoItemsCodigo) Then mProps.BandOrdenGrupoItemsCodigo = !BandOrdenGrupoItemsCodigo
        If Not IsNull(!BandOrdenPCCodigo) Then mProps.BandOrdenPCCodigo = !BandOrdenPCCodigo
        If Not IsNull(!BandOrdenPCGrupoCodigo) Then mProps.BandOrdenPCGrupoCodigo = !BandOrdenPCGrupoCodigo

        If Not IsNull(!BandFactElect) Then mProps.BandFactElect = Trim$(!BandFactElect)
        If Not IsNull(!ComprobantesGenerados) Then mProps.ComprobantesGenerados = !ComprobantesGenerados
        If Not IsNull(!ComprobantesFirmados) Then mProps.ComprobantesFirmados = !ComprobantesFirmados
        If Not IsNull(!ComprobantesAutorizados) Then mProps.ComprobantesAutorizados = !ComprobantesAutorizados
        If Not IsNull(!ComprobantesNoAutorizados) Then mProps.ComprobantesNoAutorizados = !ComprobantesNoAutorizados
        If Not IsNull(!ComprobantesEnviados) Then mProps.ComprobantesEnviados = !ComprobantesEnviados
        
        If Not IsNull(!UbicacionArchivoToken) Then mProps.UbicacionArchivoToken = !UbicacionArchivoToken
        If Not IsNull(!ContrasenaToken) Then mProps.ContrasenaToken = !ContrasenaToken
        If Not IsNull(!TipoAmbiente) Then mProps.TipoAmbiente = !TipoAmbiente
        
        If Not IsNull(!ServidorCorreo) Then mProps.ServidorCorreo = !ServidorCorreo
        If Not IsNull(!CuentaCorreo) Then mProps.CuentaCorreo = !CuentaCorreo
        If Not IsNull(!PasswordCorreo) Then mProps.PasswordCorreo = !PasswordCorreo
        If Not IsNull(!MensajeCorreo) Then mProps.MensajeCorreo = !MensajeCorreo
        
        If Not IsNull(!PuertoCorreo) Then mProps.PuertoCorreo = !PuertoCorreo
        If Not IsNull(!CopiaCorreo) Then mProps.CopiaCorreo = !CopiaCorreo
        If Not IsNull(!NombreUsuario) Then mProps.NombreUsuario = !NombreUsuario
        If Not IsNull(!ConexionSegura) Then mProps.BandConexionSegura = !ConexionSegura
        
        If Not IsNull(!ComprobantesSubidos) Then mProps.ComprobantesSubidos = !ComprobantesSubidos
        If Not IsNull(!BandPortal) Then mProps.BandPortal = Trim$(!BandPortal)
        If Not IsNull(!BandCorreoAutomatico) Then mProps.BandCorreoAutomatico = !BandCorreoAutomatico
        If Not IsNull(!BandCorreoxEquipo) Then mProps.BandCorreoxEquipo = !BandCorreoxEquipo
        If Not IsNull(!PenalizaDesctoxItem) Then mProps.PenalizaDesctoxItem = Trim$(!PenalizaDesctoxItem)
        
        If Not IsNull(!BandRebateIVG) Then mProps.BandRebateIVG = Trim$(!BandRebateIVG)
        If Not IsNull(!NumIVGRebate) Then mProps.NumIVGRebate = Trim$(!NumIVGRebate)
        If Not IsNull(!BandOmitirCambioIVA) Then mProps.BandOmitirCambioIVA = Trim$(!BandOmitirCambioIVA)
        If Not IsNull(!BandCostoUCParaCosto) Then mProps.BandCostoUCParaCosto = !BandCostoUCParaCosto
        If Not IsNull(!IVKTipoDatoDouble) Then mProps.IVKTipoDatoDouble = !IVKTipoDatoDouble
        If Not IsNull(!BandConfigLira) Then mProps.BandConfigLira = !BandConfigLira
        If Not IsNull(!BandPorDescGeneral) Then mProps.BandPorDescGeneral = !BandPorDescGeneral
        If Not IsNull(!PorDescGeneral) Then mProps.PorDescGeneral = !PorDescGeneral
        If Not IsNull(!BandDbleClickNOMod) Then mProps.BandDbleClickNOMod = !BandDbleClickNOMod
        If Not IsNull(!MensajeCorreoCartera) Then mProps.MensajeCorreoCartera = !MensajeCorreoCartera
        If Not IsNull(!RutaRepCarteraGen) Then mProps.RutaRepCarteraGen = !RutaRepCarteraGen
        If Not IsNull(!MensajePDFCartera) Then mProps.MensajePDFCartera = !MensajePDFCartera
        
        
        
        

        
        .Close
    End With
    
    '*** MAKOTO 27/nov/00 Agregado
    'Recupera opciones adicionales desde la tabla GNOpcion2
    RecuperarGNOpcion2
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

'*** MAKOTO 27/nov/00 Agregado
Public Sub GrabarGNOpcion2()
    Dim sql As String, rs As Recordset, k As String, v As String
    Dim i As Long, p As Integer
    
    With mProps.colOpciones
        For i = 1 To .Count
            k = ""
            v = .item(i)
            p = InStr(v, ";")
            
            If p > 0 Then               'Separa el código y valor
                k = Left$(v, p - 1)
                v = Right$(v, Len(v) - p)
            End If
            
            'Verifica si existe o no            '*** MAKOTO 30/ene/01 Mod.
            sql = "SELECT * FROM GNOpcion2 WHERE Codigo='" & k & "'"
            Set rs = mobjEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                'Si es que no existe, inserta uno nuevo
                sql = "INSERT GNOpcion2 (Codigo, Valor) VALUES ('" & k & "', '" & Mid$(v, 1, 255) & "')"
            Else
                'Si es que ya existe, lo modifica
                sql = "UPDATE GNOpcion2 SET Valor='" & Mid$(v, 1, 255) & "' " & _
                      "WHERE Codigo='" & k & "' AND Valor<>'" & Mid$(v, 1, 255) & "'"
            End If
            mobjEmpresa.Execute sql, True
            
            rs.Close
            Set rs = Nothing
        Next i
    End With
End Sub

'*** MAKOTO 27/nov/00 Agregado
Private Sub RecuperarGNOpcion2()
    Dim sql As String, rs As Recordset
    Dim v As String, k As String
    
    'Una vez borra todo
    Set mProps.colOpciones = Nothing
    Set mProps.colOpciones = New Collection
    
    sql = "SELECT Codigo, Valor FROM GNOpcion2 ORDER BY Codigo"
    Set rs = mobjEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            k = CStr(rs.Fields("Codigo"))
            v = k & ";" & CStr(rs.Fields("Valor"))  '   (Guarda la clave más el valor)
            mProps.colOpciones.Add item:=v, key:=k
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
End Sub

'*** MAKOTO 27/nov/00 Agregado
Public Function ObtenerValor(ByVal cod As String) As String
    Dim s As String
    
    On Error Resume Next
    s = mProps.colOpciones.item(cod)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        ObtenerValor = ""
    Else
        If InStr(s, ";") > 0 Then
            s = Right$(s, Len(s) - InStr(s, ";"))   'Devuelve solo la parte de valor
        End If
        ObtenerValor = s
    End If
End Function

'*** MAKOTO 27/nov/00 Agregado
Public Sub AsignarValor(ByVal cod As String, ByVal Valor As String)
    With mProps.colOpciones
        On Error Resume Next        'Ignorar error al llamar .Remove
        .Remove cod                     'Primero elimina
        Err.Clear                   'Limpia error si ha ocurrido
        On Error GoTo 0             'Anular la ignoración de error
        Valor = cod & ";" & Valor       '   (Guarda la clave más el valor)
        .Add item:=Valor, key:=cod      ' y luego agrega de nuevo
    End With
    mbooModificado = True
End Sub


Friend Property Set Empresa(objEmpresa As Empresa)
    Set mobjEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mobjEmpresa
End Property

Public Property Let NombreEmpresa(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.NombreEmpresa", MSGERR_INVALIDO
    ElseIf Len(value) = 0 Then
        value = "(Nombre de empresa)"
    End If
    
    mProps.NombreEmpresa = value
    mbooModificado = True
End Property

Public Property Get NombreEmpresa() As String
    NombreEmpresa = Trim$(mProps.NombreEmpresa)
End Property

Public Property Let Direccion1(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.Direccion1", MSGERR_INVALIDO
    End If
    
    mProps.Direccion1 = value
    mbooModificado = True
End Property

Public Property Get Direccion1() As String
    Direccion1 = Trim$(mProps.Direccion1)
End Property

Public Property Let Direccion2(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.Direccion2", MSGERR_INVALIDO
    End If
    
    mProps.Direccion2 = value
    mbooModificado = True
End Property

Public Property Get Direccion2() As String
    Direccion2 = Trim$(mProps.Direccion2)
End Property

Public Property Let Telefono1(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.Telefono1", MSGERR_INVALIDO
    End If
    
    mProps.Telefono1 = value
    mbooModificado = True
End Property

Public Property Get Telefono1() As String
    Telefono1 = Trim$(mProps.Telefono1)
End Property

Public Property Let Telefono2(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.Telefono2", MSGERR_INVALIDO
    End If
    
    mProps.Telefono2 = value
    mbooModificado = True
End Property

Public Property Get Telefono2() As String
    Telefono2 = Trim$(mProps.Telefono2)
End Property

Public Property Let Telefono3(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.Telefono3", MSGERR_INVALIDO
    End If
    
    mProps.Telefono3 = value
    mbooModificado = True
End Property

Public Property Get Telefono3() As String
    Telefono3 = Trim$(mProps.Telefono3)
End Property

Public Property Let Fax1(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.Fax1", MSGERR_INVALIDO
    End If
    
    mProps.Fax1 = value
    mbooModificado = True
End Property

Public Property Get Fax1() As String
    Fax1 = Trim$(mProps.Fax1)
End Property

Public Property Let Fax2(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.Fax2", MSGERR_INVALIDO
    End If
    
    mProps.Fax2 = value
    mbooModificado = True
End Property

Public Property Get Fax2() As String
    Fax2 = Trim$(mProps.Fax2)
End Property

Public Property Let RUC(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.RUC", MSGERR_INVALIDO
    End If
    
    mProps.RUC = value
    mbooModificado = True
End Property

Public Property Get RUC() As String
    RUC = Trim$(mProps.RUC)
End Property

Public Property Let EMail(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.EMail", MSGERR_INVALIDO
    End If
    
    mProps.EMail = value
    mbooModificado = True
End Property

Public Property Get EMail() As String
    EMail = Trim$(mProps.EMail)
End Property


'Manejo de LogoTipo está Pendiente
Public Property Let LogoTipo(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.LogoTipo", MSGERR_INVALIDO
    End If
    
    mProps.LogoTipo = value
    mbooModificado = True
End Property

Public Property Get LogoTipo() As String
    LogoTipo = Trim$(mProps.LogoTipo)
End Property

Public Property Let FechaInicio(ByVal value As Date)
    mProps.FechaInicio = value
    mbooModificado = True
End Property

Public Property Get FechaInicio() As Date
    FechaInicio = mProps.FechaInicio
End Property

Public Property Let FechaFinal(ByVal value As Date)
    mProps.FechaFinal = value
    mbooModificado = True
End Property

Public Property Get FechaFinal() As Date
    FechaFinal = mProps.FechaFinal
End Property

Public Property Let CodAsientoSI(ByVal value As Long)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.CodAsientoSI", MSGERR_INVALIDO
    End If
    
    mProps.CodAsientoSI = value
    mbooModificado = True
End Property

Public Property Get CodAsientoSI() As Long
    CodAsientoSI = mProps.CodAsientoSI
End Property

Public Property Let CodAsientoFinal(ByVal value As Long)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.CodAsientoFinal", MSGERR_INVALIDO
    End If
    
    mProps.CodAsientoFinal = value
    mbooModificado = True
End Property

Public Property Get CodAsientoFinal() As Long
    CodAsientoFinal = mProps.CodAsientoFinal
End Property

Public Property Let IdCuentaResultado(ByVal value As Long)
    mProps.IdCuentaResultado = value
    mbooModificado = True
End Property

Public Property Get IdCuentaResultado() As Long
    IdCuentaResultado = mProps.IdCuentaResultado
End Property


Public Property Let CodCuentaResultado(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodCuentaResultado Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaResultado también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mobjEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNOpcion.CodCuentaResultado", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaResultado = rs!IdCuenta
            End If
            rs.Close
        Else
            IdCuentaResultado = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodCuentaResultado = value
    mbooModificado = True
End Property

Public Property Get CodCuentaResultado() As String
    CodCuentaResultado = Trim$(mProps.CodCuentaResultado)
End Property


Public Property Let IdCuentaProvPre(ByVal value As Long)
    mProps.IdCuentaProvPre = value
    mbooModificado = True
End Property

Public Property Get IdCuentaProvPre() As Long
    IdCuentaProvPre = mProps.IdCuentaProvPre
End Property


Public Property Let CodCuentaProvPre(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodCuentaProvPre Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaProvPre también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mobjEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNOpcion.CodCuentaProvPre", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaProvPre = rs!IdCuenta
            End If
            rs.Close
        Else
            IdCuentaProvPre = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodCuentaProvPre = value
    mbooModificado = True
End Property

Public Property Get CodCuentaProvPre() As String
    CodCuentaProvPre = Trim$(mProps.CodCuentaProvPre)
End Property

Public Property Let IdCuentaCliPre(ByVal value As Long)
    mProps.IdCuentaCliPre = value
    mbooModificado = True
End Property

Public Property Get IdCuentaCliPre() As Long
    IdCuentaCliPre = mProps.IdCuentaCliPre
End Property


Public Property Let CodCuentaCliPre(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodCuentaCliPre Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaCliPre también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mobjEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNOpcion.CodCuentaCliPre", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaCliPre = rs!IdCuenta
            End If
            rs.Close
        Else
            IdCuentaCliPre = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodCuentaCliPre = value
    mbooModificado = True
End Property

Public Property Get CodCuentaCliPre() As String
    CodCuentaCliPre = Trim$(mProps.CodCuentaCliPre)
End Property


Public Property Let EtiqGrupo(NumGrupo As Integer, ByVal value As String)
'----- NumGrupo = Entre 1 a 5

    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.EtiqGrupo", MSGERR_INVALIDO
    ElseIf Len(value) = 0 Then
        value = "Grupo" & NumGrupo
    End If
    
    mProps.EtiqGrupo(NumGrupo) = value
    mbooModificado = True
End Property

Public Property Get EtiqGrupo(NumGrupo As Integer) As String
'----- NumGrupo = Entre 1 a 5
    EtiqGrupo = Trim$(mProps.EtiqGrupo(NumGrupo))
End Property


Public Property Let EtiqPCGrupo(NumGrupo As Integer, ByVal value As String)
'----- NumGrupo = Entre 1 a 3
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.EtiqPCGrupo", MSGERR_INVALIDO
    ElseIf Len(value) = 0 Then
        value = "Grupo" & NumGrupo
    End If
    
    mProps.EtiqPCGrupo(NumGrupo) = value
    mbooModificado = True
End Property

Public Property Get EtiqPCGrupo(NumGrupo As Integer) As String
'----- NumGrupo = Entre 1 a 3
    EtiqPCGrupo = Trim$(mProps.EtiqPCGrupo(NumGrupo))
End Property

Public Property Let FechaLimiteDesde(ByVal value As Date)
    mProps.FechaLimiteDesde = value
    mbooModificado = True
End Property

Public Property Get FechaLimiteDesde() As Date
    FechaLimiteDesde = mProps.FechaLimiteDesde
End Property

Public Property Let FechaLimiteHasta(ByVal value As Date)
    mProps.FechaLimiteHasta = value
    mbooModificado = True
End Property

Public Property Get FechaLimiteHasta() As Date
    FechaLimiteHasta = mProps.FechaLimiteHasta
End Property



Public Property Let CodMoneda(ByVal index As Integer, ByVal value As String)
    If Len(value) > 5 Or index > 4 Or IsNumeric(value) Then
        Err.Raise ERR_INVALIDO, "GNOpcion.CodMoneda", MSGERR_INVALIDO
    End If
    
    mProps.CodMoneda(index) = value
    mbooModificado = True
End Property

Public Property Get CodMoneda(ByVal index As Integer) As String
    CodMoneda = Trim$(mProps.CodMoneda(index))
End Property

'*** MAKOTO 31/ene/01 Mod. a Public
'Friend Property Get IndiceMoneda(ByVal CodMoneda As String) As Integer
Public Property Get IndiceMoneda(ByVal CodMoneda As String) As Integer
    Dim i As Integer
    
    'Busca el código de moneda
    For i = 1 To 4
        If mProps.CodMoneda(i) = CodMoneda Then Exit For
    Next i
    If i > 4 Then
'        Err.Raise ERR_INVALIDO, "GNOpcion.FormatoMoneda", MSGERR_INVALIDO
        i = -1
    End If
    
    IndiceMoneda = i
End Property


Public Property Let NombreMoneda(ByVal key As Variant, ByVal value As String)
    Dim i As Integer
    If Len(value) > 20 Or Len(CStr(key)) = 0 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.NombreMoneda", MSGERR_INVALIDO
    End If
    
    If IsNumeric(key) Then
        i = key
    Else
        i = IndiceMoneda(key)       'Busca el código de moneda
    End If
    
    If i >= LBound(mProps.NombreMoneda) And _
       i <= UBound(mProps.NombreMoneda) Then
        mProps.NombreMoneda(i) = value
    End If
    mbooModificado = True
End Property

Public Property Get NombreMoneda(ByVal key As Variant) As String
    Dim i As Integer
    
    i = -1
    If IsNumeric(key) Then
        i = key
    Else
        i = IndiceMoneda(key)                             'Busca el código de moneda
    End If
    
    If i >= LBound(mProps.NombreMoneda) And _
       i <= UBound(mProps.NombreMoneda) Then
        NombreMoneda = mProps.NombreMoneda(i)
    End If
End Property

Public Property Let FormatoMoneda(ByVal key As Variant, ByVal value As String)
    Dim i As Integer
    If Len(value) > 20 Or Len(CStr(key)) = 0 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.FormatoMoneda", MSGERR_INVALIDO
    ElseIf Len(value) = 0 Then
        value = "#,0.00"       'Valor predeterminado
    End If
    
    If IsNumeric(key) Then
        i = key
    Else
        i = IndiceMoneda(key)       'Busca el código de moneda
    End If
    
    If i >= LBound(mProps.FormatoMoneda) And _
       i <= UBound(mProps.FormatoMoneda) Then
        mProps.FormatoMoneda(i) = value
    End If
    mbooModificado = True
End Property

Public Property Get FormatoMoneda(ByVal key As Variant) As String
    Dim i As Integer
    
    i = -1
    If IsNumeric(key) Then
        i = key
    Else
        i = IndiceMoneda(key)       'Busca el código de moneda
    End If
    
    If i >= LBound(mProps.FormatoMoneda) And _
       i <= UBound(mProps.FormatoMoneda) Then
        FormatoMoneda = mProps.FormatoMoneda(i)
    End If
End Property




Public Property Let FormatoCantidad(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.FormatoCantidad", MSGERR_INVALIDO
    ElseIf Len(value) = 0 Then
        value = "#,0"       'Valor predeterminado
    End If
    
    mProps.FormatoCantidad = value
    mbooModificado = True
End Property

Public Property Get FormatoCantidad() As String
    FormatoCantidad = Trim$(mProps.FormatoCantidad)
End Property

Public Property Let FormatoFecha(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.FormatoFecha", MSGERR_INVALIDO
    ElseIf Len(value) = 0 Then
        value = "aaaa/MM/dd"       'Valor predeterminado
    End If
    
    mProps.FormatoFecha = value
    mbooModificado = True
End Property

Public Property Get FormatoFecha() As String
    FormatoFecha = Trim$(mProps.FormatoFecha)
End Property

Public Property Let Costeo(ByVal value As Byte)
    If value > 3 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.Costeo", MSGERR_INVALIDO
    End If
    mProps.Costeo = value
    mbooModificado = True
End Property

Public Property Get Costeo() As Byte
    Costeo = mProps.Costeo
End Property


Public Property Let GeneraCodItem(ByVal value As Byte)
    If value > 3 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.GeneraCodItem", MSGERR_INVALIDO
    End If
    mProps.GeneraCodItem = value
    mbooModificado = True
End Property

Public Property Get GeneraCodItem() As Byte
    GeneraCodItem = mProps.GeneraCodItem
End Property

Public Property Let GeneraCodProv(ByVal value As Byte)
    If value > 4 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.GeneraCodProv", MSGERR_INVALIDO
    End If
    mProps.GeneraCodProv = value
    mbooModificado = True
End Property

Public Property Get GeneraCodProv() As Byte
    GeneraCodProv = mProps.GeneraCodProv
End Property

Public Property Let GeneraCodCli(ByVal value As Byte)
    If value > 5 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.GeneraCodCli", MSGERR_INVALIDO
    End If
    mProps.GeneraCodCli = value
    mbooModificado = True
End Property

Public Property Get GeneraCodCli() As Byte
    GeneraCodCli = mProps.GeneraCodCli
End Property

Public Property Let GeneraCodCC(ByVal value As Byte)
    If value > 3 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.GeneraCodCC", MSGERR_INVALIDO
    End If
    mProps.GeneraCodCC = value
    mbooModificado = True
End Property

Public Property Get GeneraCodCC() As Byte
    GeneraCodCC = mProps.GeneraCodCC
End Property



Public Property Let PrefijoCodItem(ByVal value As String)
    If Len(value) > 3 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.PrefijoCodItem", MSGERR_INVALIDO
    End If
    mProps.PrefijoCodItem = value
    mbooModificado = True
End Property

Public Property Get PrefijoCodItem() As String
    PrefijoCodItem = Trim$(mProps.PrefijoCodItem)
End Property

Public Property Let PrefijoCodProv(ByVal value As String)
    If Len(value) > 3 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.PrefijoCodProv", MSGERR_INVALIDO
    End If
    mProps.PrefijoCodProv = value
    mbooModificado = True
End Property

Public Property Get PrefijoCodProv() As String
    PrefijoCodProv = Trim$(mProps.PrefijoCodProv)
End Property

Public Property Let PrefijoCodCli(ByVal value As String)
    If Len(value) > 3 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.PrefijoCodCli", MSGERR_INVALIDO
    End If
    mProps.PrefijoCodCli = value
    mbooModificado = True
End Property

Public Property Get PrefijoCodCli() As String
    PrefijoCodCli = Trim$(mProps.PrefijoCodCli)
End Property

Public Property Let PrefijoCodCC(ByVal value As String)
    If Len(value) > 3 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.PrefijoCodCC", MSGERR_INVALIDO
    End If
    mProps.PrefijoCodCC = value
    mbooModificado = True
End Property

Public Property Get PrefijoCodCC() As String
    PrefijoCodCC = Trim$(mProps.PrefijoCodCC)
End Property

Public Property Let PermiteItemSinCuenta(ByVal value As Boolean)
    mProps.PermiteItemSinCuenta = value
    mbooModificado = True
End Property

Public Property Get PermiteItemSinCuenta() As Boolean
    PermiteItemSinCuenta = Trim$(mProps.PermiteItemSinCuenta)
End Property

Public Property Let OcultaContable(ByVal value As Boolean)
    mProps.OcultaContable = value
    mbooModificado = True
End Property

Public Property Get OcultaContable() As Boolean
    OcultaContable = Trim$(mProps.OcultaContable)
End Property

Public Property Let Los100DocCobroPago(ByVal value As Boolean)
    mProps.Los100DocCobroPago = value
    mbooModificado = True
End Property

Public Property Get Los100DocCobroPago() As Boolean
    Los100DocCobroPago = Trim$(mProps.Los100DocCobroPago)
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
    Dim i As Integer
    mbooEsNuevo = True
    mbooModificado = False
    
    'Valores predeterminados
    With mProps
        .NombreEmpresa = "(Nombre de empresa)"
        For i = 1 To IVGRUPO_MAX
            .EtiqGrupo(i) = "Grupo" & i
        Next i
        For i = 1 To PCGRUPO_MAX
            .EtiqPCGrupo(i) = "Grupo" & i
        Next i
        .FechaLimiteDesde = DateSerial(1990, 1, 1)
        .FechaLimiteHasta = DateSerial(2100, 12, 31)
        For i = 1 To 4
            .FormatoMoneda(i) = "#,0.00"
        Next i
        .CodMoneda(1) = MONEDA_PRE        'Moneda predeterminada
        .CodMoneda(2) = MONEDA_SEC        'Moneda secundaria
        .NombreMoneda(1) = "Sucres"
        .NombreMoneda(2) = "Dolares"
        .FormatoCantidad = "#,0"
        .FormatoFecha = "yyyy/MM/dd"
        .Costeo = COSTO_PROM
        
        .GeneraCodItem = 0          'Manual
        .GeneraCodProv = 1          'Auto generacion secuencial
        .GeneraCodCli = 1           'Auto generacion secuencial
        .GeneraCodCC = 2            'Auto generacion con formato
        .PrefijoCodItem = "CUE"
        .PrefijoCodProv = "P"
        .PrefijoCodCli = "C"
        .PrefijoCodCC = "CUE"
        .PermiteItemSinCuenta = False
        .OcultaContable = False
        .TipoEmpresa = "SC"
        .TipoEmpresaSRI = "SC"
        'jeaa 09/12/2008
        For i = 1 To AFGRUPO_MAX
            .EtiqAFGrupo(i) = "Grupo" & i
        Next i
        .GeneraCodAFItem = 0          'Manual
        .PrefijoCodAFItem = "AF"
        .RazonSocial = .NombreEmpresa
        'jeaa 21/03/2009
        For i = 1 To GNCGRUPO_MAX
            .EtiqGNCGrupo(i) = "Grupo" & i
        Next i
        'jeaa 30/03/2009
        For i = 1 To GNVGRUPO_MAX
            .EtiqGNVGrupo(i) = "Grupo" & i
        Next i
        
        For i = 1 To PCGRUPO_MAX
            .EtiqPCGrupoP(i) = "Grupo" & i
        Next i
        For i = 1 To PCGRUPO_MAX
            .EtiqPCGrupoC(i) = "Grupo" & i
        Next i
        For i = 1 To PCGRUPO_MAX
            .EtiqPCGrupoG(i) = "Grupo" & i
        Next i
        For i = 1 To PCGRUPO_MAX
            .EtiqPCGrupoE(i) = "Grupo" & i
        Next i
        
        For i = 1 To PCGRUPO_MAX
            .EtiqPCGrupoL(i) = "Grupo" & i
        Next i
        
        .GeneraCodAFItemALT = 0          'Manual
        .PrefijoCodAFItemALT = "AF"
        
        
        'Crea una instancia de la colección         '*** MAKOTO 27/nov/00 Agregado
        Set .colOpciones = New Collection
    End With
End Sub

Private Sub Class_Terminate()
    Set mProps.colOpciones = Nothing        '*** MAKOTO 27/nov/00 Agregado
    Set mobjEmpresa = Nothing
End Sub

Public Property Let TipoEmpresa(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.TipoEmpresa", MSGERR_INVALIDO
    End If
    mProps.TipoEmpresa = value
    mbooModificado = True
End Property

Public Property Get TipoEmpresa() As String
    TipoEmpresa = mProps.TipoEmpresa
End Property

Public Property Let TipoEmpresaSRI(ByVal value As String)
    If Len(value) > 4 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.TipoEmpresaSRI", MSGERR_INVALIDO
    End If
    mProps.TipoEmpresaSRI = value
    mbooModificado = True
End Property

Public Property Get TipoEmpresaSRI() As String
    TipoEmpresaSRI = mProps.TipoEmpresaSRI
End Property


Public Property Let PorcentajeIVA(ByVal value As Currency)
    If value < 0 Or value > 1 Then
        Err.Raise ERR_INVALIDO, "IVInventario_PorcentajeIVA", MSGERR_INVALIDO
    End If
    
    If mProps.PorcentajeIVA <> value Then
        mProps.PorcentajeIVA = value
        mbooModificado = True
    End If
End Property

Public Property Get PorcentajeIVA() As Currency
    PorcentajeIVA = mProps.PorcentajeIVA
End Property

Public Property Let PorcentajeIVAAnt(ByVal value As Currency)
    If value < 0 Or value > 1 Then
        Err.Raise ERR_INVALIDO, "IVInventario_PorcentajeIVAAnt", MSGERR_INVALIDO
    End If
    
    If mProps.PorcentajeIVAAnt <> value Then
        mProps.PorcentajeIVAAnt = value
        mbooModificado = True
    End If
End Property

Public Property Get PorcentajeIVAAnt() As Currency
    PorcentajeIVAAnt = mProps.PorcentajeIVAAnt
End Property

Public Property Let FechaIVA(ByVal value As Date)
    mProps.FechaIVA = value
    mbooModificado = True
End Property

Public Property Get FechaIVA() As Date
    FechaIVA = mProps.FechaIVA
End Property



'AUC valida si existen los valores
' se realiza solo la primera vez cuando no se tiene cargado nada aun
Public Function ValidarValor(ByVal cod As String) As Boolean
Dim rs As Recordset, sql As String
sql = "select * from gnopcion2 where codigo = '" & cod & "'"
Set rs = Me.Empresa.OpenRecordset(sql)
    If rs.RecordCount = 0 Then
        ValidarValor = False
    Else
        ValidarValor = True
    End If
Set rs = Nothing
End Function

Public Property Let FormatoCodItem(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.FormatoCodItem", MSGERR_INVALIDO
    End If
    mProps.FormatoCodItem = value
    mbooModificado = True
End Property

Public Property Get FormatoCodItem() As String
    FormatoCodItem = Trim$(mProps.FormatoCodItem)
End Property

Public Property Let FormatoCodProv(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.FormatoCodProv", MSGERR_INVALIDO
    End If
    mProps.FormatoCodProv = value
    mbooModificado = True
End Property

Public Property Get FormatoCodProv() As String
    FormatoCodProv = Trim$(mProps.FormatoCodProv)
End Property

Public Property Let FormatoCodCli(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.FormatoCodCli", MSGERR_INVALIDO
    End If
    mProps.FormatoCodCli = value
    mbooModificado = True
End Property

Public Property Get FormatoCodCli() As String
    FormatoCodCli = Trim$(mProps.FormatoCodCli)
End Property

Public Property Let FormatoCodCC(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.FormatoCodCC", MSGERR_INVALIDO
    End If
    mProps.FormatoCodCC = value
    mbooModificado = True
End Property

Public Property Get FormatoCodCC() As String
    FormatoCodCC = Trim$(mProps.FormatoCodCC)
End Property

'AUC 29/10/07
Public Property Get EtiqEst(num As Integer) As String
    EtiqEst = Trim$(mProps.EtiqEst(num))
End Property
Public Property Let EtiqEst(num As Integer, ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.Est", MSGERR_INVALIDO
    ElseIf Len(value) = 0 Then
        value = "Est" & num
    End If
    mProps.EtiqEst(num) = value
    mbooModificado = True
End Property

Public Property Let NumAutorizacion_AutoImp(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.NumAutorizacion_AutoImp", MSGERR_INVALIDO
    ElseIf Len(value) = 0 Then
        value = "(Número Autorización Auto Impresores )"
    End If
    
    mProps.NumAutorizacion_AutoImp = value
    mbooModificado = True
End Property

Public Property Get NumAutorizacion_AutoImp() As String
    NumAutorizacion_AutoImp = Trim$(mProps.NumAutorizacion_AutoImp)
End Property

Public Property Let FechaCaducidad_AutoImp(ByVal value As Date)
    mProps.FechaCaducidad_AutoImp = value
    mbooModificado = True
End Property

Public Property Get FechaCaducidad_AutoImp() As Date
    FechaCaducidad_AutoImp = mProps.FechaCaducidad_AutoImp
End Property


Public Property Let FechaAutorizacion_AutoImp(ByVal value As Date)
    mProps.FechaAutorizacion_AutoImp = value
    mbooModificado = True
End Property

Public Property Get FechaAutorizacion_AutoImp() As Date
    FechaAutorizacion_AutoImp = mProps.FechaAutorizacion_AutoImp
End Property



Public Property Let NumAutorizacion_AutoImpOld(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.NumAutorizacion_AutoImpOld", MSGERR_INVALIDO
    ElseIf Len(value) = 0 Then
        value = "(Número Autorización Auto Impresores )"
    End If
    
    mProps.NumAutorizacion_AutoImpOld = value
    mbooModificado = True
End Property

Public Property Get NumAutorizacion_AutoImpOld() As String
    NumAutorizacion_AutoImpOld = Trim$(mProps.NumAutorizacion_AutoImpOld)
End Property

Public Property Let FechaCaducidad_AutoImpOld(ByVal value As Date)
    mProps.FechaCaducidad_AutoImpOld = value
    mbooModificado = True
End Property

Public Property Get FechaCaducidad_AutoImpOld() As Date
    FechaCaducidad_AutoImpOld = mProps.FechaCaducidad_AutoImpOld
End Property


Public Property Let FechaAutorizacion_AutoImpOld(ByVal value As Date)
    mProps.FechaAutorizacion_AutoImpOld = value
    mbooModificado = True
End Property

Public Property Get FechaAutorizacion_AutoImpOld() As Date
    FechaAutorizacion_AutoImpOld = mProps.FechaAutorizacion_AutoImpOld
End Property


Public Property Let EtiqAFGrupo(NumGrupo As Integer, ByVal value As String)
'----- NumGrupo = Entre 1 a 5
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.EtiqAFGrupo", MSGERR_INVALIDO
    ElseIf Len(value) = 0 Then
        value = "Grupo" & NumGrupo
    End If
    mProps.EtiqAFGrupo(NumGrupo) = value
    mbooModificado = True
End Property
Public Property Get EtiqAFGrupo(NumGrupo As Integer) As String
'----- NumGrupo = Entre 1 a 5
    EtiqAFGrupo = Trim$(mProps.EtiqAFGrupo(NumGrupo))
End Property

Public Property Let GeneraCodAFItem(ByVal value As Byte)
    If value > 4 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.GeneraCodAFItem", MSGERR_INVALIDO
    End If
    mProps.GeneraCodAFItem = value
    mbooModificado = True
End Property

Public Property Get GeneraCodAFItem() As Byte
    GeneraCodAFItem = mProps.GeneraCodAFItem
End Property

Public Property Let PrefijoCodAFItem(ByVal value As String)
    If Len(value) > 3 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.PrefijoCodAFItem", MSGERR_INVALIDO
    End If
    mProps.PrefijoCodAFItem = value
    mbooModificado = True
End Property

Public Property Get PrefijoCodAFItem() As String
    PrefijoCodAFItem = Trim$(mProps.PrefijoCodAFItem)
End Property

Public Property Let FormatoCodAFItem(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.FormatoCodAFItem", MSGERR_INVALIDO
    End If
    mProps.FormatoCodAFItem = value
    mbooModificado = True
End Property

Public Property Get FormatoCodAFItem() As String
    FormatoCodAFItem = Trim$(mProps.FormatoCodAFItem)
End Property

Public Property Let RazonSocial(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.RazonSocial", MSGERR_INVALIDO
    ElseIf Len(value) = 0 Then
        value = "(Razon Social)"
    End If
    
    mProps.RazonSocial = value
    mbooModificado = True
End Property

Public Property Get RazonSocial() As String
    RazonSocial = Trim$(mProps.RazonSocial)
End Property

Private Sub ActulizaIVAItems(ByVal IVA As Double)
    Dim sql As String, rs As Recordset
    sql = "update IVinventario set porcentajeIVA=" & IVA
    sql = sql & " where bandIVA=1"
    Set rs = mobjEmpresa.OpenRecordset(sql)
    
    sql = "update AFinventario set porcentajeIVA=" & IVA
    sql = sql & " where bandIVA=1"
    Set rs = mobjEmpresa.OpenRecordset(sql)
    
    
End Sub

Private Sub ActulizaIVAItemsAnt(ByVal IVA As Double)
    Dim sql As String, rs As Recordset
    sql = "update IVinventario set porcentajeIVAAnt =" & IVA
    sql = sql & " where bandIVA=1"
    Set rs = mobjEmpresa.OpenRecordset(sql)
    
    sql = "update afinventario set porcentajeIVAAnt =" & IVA
    sql = sql & " where bandIVA=1"
    Set rs = mobjEmpresa.OpenRecordset(sql)

End Sub


Public Property Let EtiqGNCGrupo(NumGrupo As Integer, ByVal value As String)
'----- NumGrupo = Entre 1 a 4
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.EtiqGNCGrupo", MSGERR_INVALIDO
    ElseIf Len(value) = 0 Then
        value = "Grupo" & NumGrupo
    End If
    mProps.EtiqGNCGrupo(NumGrupo) = value
    mbooModificado = True
End Property

Public Property Get EtiqGNCGrupo(NumGrupo As Integer) As String
'----- NumGrupo = Entre 1 a 4
    EtiqGNCGrupo = Trim$(mProps.EtiqGNCGrupo(NumGrupo))
End Property

Public Property Let IdCuentaActivoPre(ByVal value As Long)
    mProps.IdCuentaActivoPre = value
    mbooModificado = True
End Property

Public Property Get IdCuentaActivoPre() As Long
    IdCuentaActivoPre = mProps.IdCuentaActivoPre
End Property

Public Property Let IdCuentaCostoPre(ByVal value As Long)
    mProps.IdCuentaCostoPre = value
    mbooModificado = True
End Property

Public Property Get IdCuentaCostoPre() As Long
    IdCuentaCostoPre = mProps.IdCuentaCostoPre
End Property

Public Property Let IdCuentaVentaPre(ByVal value As Long)
    mProps.IdCuentaVentaPre = value
    mbooModificado = True
End Property

Public Property Get IdCuentaVentaPre() As Long
    IdCuentaVentaPre = mProps.IdCuentaVentaPre
End Property

Public Property Let CodCuentaActivoPre(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodCuentaActivoPre Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivoPre también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mobjEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNOpcion.CodCuentaActivoPre", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaActivoPre = rs!IdCuenta
            End If
            rs.Close
        Else
            IdCuentaActivoPre = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodCuentaActivoPre = value
    mbooModificado = True
End Property

Public Property Get CodCuentaActivoPre() As String
    CodCuentaActivoPre = Trim$(mProps.CodCuentaActivoPre)
End Property

Public Property Let CodCuentaCostoPre(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodCuentaCostoPre Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaCostoPre también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mobjEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNOpcion.CodCuentaCostoPre", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaCostoPre = rs!IdCuenta
            End If
            rs.Close
        Else
            IdCuentaCostoPre = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodCuentaCostoPre = value
    mbooModificado = True
End Property

Public Property Get CodCuentaCostoPre() As String
    CodCuentaCostoPre = Trim$(mProps.CodCuentaCostoPre)
End Property

Public Property Let CodCuentaVentaPre(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodCuentaVentaPre Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaVentaPre también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mobjEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNOpcion.CodCuentaVentaPre", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaVentaPre = rs!IdCuenta
            End If
            rs.Close
        Else
            IdCuentaVentaPre = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodCuentaVentaPre = value
    mbooModificado = True
End Property

Public Property Get CodCuentaVentaPre() As String
    CodCuentaVentaPre = Trim$(mProps.CodCuentaVentaPre)
End Property

Public Property Let EtiqGNVGrupo(NumGrupo As Integer, ByVal value As String)
'----- NumGrupo = Entre 1 a 4
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.EtiqGNVGrupo", MSGERR_INVALIDO
    ElseIf Len(value) = 0 Then
        value = "Grupo" & NumGrupo
    End If
    mProps.EtiqGNVGrupo(NumGrupo) = value
    mbooModificado = True
End Property

Public Property Get EtiqGNVGrupo(NumGrupo As Integer) As String
'----- NumGrupo = Entre 1 a 4
    EtiqGNVGrupo = Trim$(mProps.EtiqGNVGrupo(NumGrupo))
End Property

Private Sub ActulizaIVAPredeterminadoTrans(ByVal IVA As Double)
    Dim sql As String, rs As Recordset
    sql = " Update GNTransRecargo"
    sql = sql & " set porcentajepre=" & IVA
    sql = sql & " from gntransrecargo gntr"
    sql = sql & " inner join ivrecargo ivr on gntr.idrecargo=ivr.idrecargo"
    sql = sql & " where codrecargo='IVA' and porcentajepre<>0"
    Set rs = mobjEmpresa.OpenRecordset(sql)
End Sub

Public Property Let IdCuentaSobranteCajaPre(ByVal value As Long)
    mProps.IdCuentaSobranteCajaPre = value
    mbooModificado = True
End Property

Public Property Get IdCuentaSobranteCajaPre() As Long
    IdCuentaSobranteCajaPre = mProps.IdCuentaSobranteCajaPre
End Property


Public Property Let CodCuentaSobranteCajaPre(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodCuentaSobranteCajaPre Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaSobranteCajaPre también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mobjEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNOpcion.CodCuentaSobranteCajaPre", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaSobranteCajaPre = rs!IdCuenta
            End If
            rs.Close
        Else
            IdCuentaSobranteCajaPre = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodCuentaSobranteCajaPre = value
    mbooModificado = True
End Property

Public Property Get CodCuentaSobranteCajaPre() As String
    CodCuentaSobranteCajaPre = Trim$(mProps.CodCuentaSobranteCajaPre)
End Property


Public Property Let MontoMaximoCF(ByVal value As Currency)
    If mProps.MontoMaximoCF <> value Then
        mProps.MontoMaximoCF = value
        mbooModificado = True
    End If
End Property

Public Property Get MontoMaximoCF() As Currency
    MontoMaximoCF = mProps.MontoMaximoCF
End Property

Public Property Let RutaXMLValido(ByVal value As String)
    If mProps.RutaXMLValido <> value Then
        mProps.RutaXMLValido = value
        mbooModificado = True
    End If
End Property

Public Property Get RutaXMLValido() As String
    RutaXMLValido = mProps.RutaXMLValido
End Property

Public Property Let RutaXMLAnulado(ByVal value As String)
    If mProps.RutaXMLAnulado <> value Then
        mProps.RutaXMLAnulado = value
        mbooModificado = True
    End If
End Property

Public Property Get RutaXMLAnulado() As String
    RutaXMLAnulado = mProps.RutaXMLAnulado
End Property


Public Property Let Resolucion(ByVal value As String)
    If mProps.Resolucion <> value Then
        mProps.Resolucion = value
        mbooModificado = True
    End If
End Property

Public Property Get Resolucion() As String
    Resolucion = mProps.Resolucion
End Property

Public Property Let AutoImpresor(ByVal value As Boolean)
    mProps.AutoImpresor = value
    mbooModificado = True
End Property

Public Property Get AutoImpresor() As Boolean
    AutoImpresor = Trim$(mProps.AutoImpresor)
End Property

Public Property Let UtilidadMarkDown(ByVal value As Boolean)
    mProps.UtilidadMarkDown = value
    mbooModificado = True
End Property

Public Property Get UtilidadMarkDown() As Boolean
    UtilidadMarkDown = Trim$(mProps.UtilidadMarkDown)
End Property


Public Property Let IdProvinciaPre(ByVal value As Long)
    mProps.IdProvinciaPre = value
    mbooModificado = True
End Property

Public Property Get IdProvinciaPre() As Long
    IdProvinciaPre = mProps.IdProvinciaPre
End Property


Public Property Let CodProvinciaPre(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodProvinciaPre Then
        If Len(value) > 0 Then
            'Actualiza IdProvinciaPre también
            sql = "SELECT IdProvincia FROM PCProvincia WHERE CodProvincia='" & value & "'"
            Set rs = mobjEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNOpcion.CodProvinciaPre", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdProvinciaPre = rs!IdProvincia
            End If
            rs.Close
        Else
            IdProvinciaPre = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodProvinciaPre = value
    mbooModificado = True
End Property

Public Property Get CodProvinciaPre() As String
    CodProvinciaPre = Trim$(mProps.CodProvinciaPre)
End Property

Public Property Let IdCantonPre(ByVal value As Long)
    mProps.IdCantonPre = value
    mbooModificado = True
End Property

Public Property Get IdCantonPre() As Long
    IdCantonPre = mProps.IdCantonPre
End Property


Public Property Let CodCantonPre(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodCantonPre Then
        If Len(value) > 0 Then
            'Actualiza IdCantonPre también
            sql = "SELECT IdCanton FROM PCCanton WHERE CodCanton='" & value & "'"
            Set rs = mobjEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNOpcion.CodCantonPre", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCantonPre = rs!IdCanton
            End If
            rs.Close
        Else
            IdCantonPre = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodCantonPre = value
    mbooModificado = True
End Property

Public Property Get CodCantonPre() As String
    CodCantonPre = Trim$(mProps.CodCantonPre)
End Property

Public Property Let EtiqPCGrupoC(NumGrupo As Integer, ByVal value As String)
'----- NumGrupo = Entre 1 a 4
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.EtiqPCGrupoC", MSGERR_INVALIDO
    ElseIf Len(value) = 0 Then
        value = "Grupo" & NumGrupo
    End If
    mProps.EtiqPCGrupoC(NumGrupo) = value
    mbooModificado = True
End Property

Public Property Get EtiqPCGrupoC(NumGrupo As Integer) As String
'----- NumGrupo = Entre 1 a 4
    EtiqPCGrupoC = Trim$(mProps.EtiqPCGrupoC(NumGrupo))
End Property

Public Property Let EtiqPCGrupoP(NumGrupo As Integer, ByVal value As String)
'----- NumGrupo = Entre 1 a 4
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.EtiqPCGrupoP", MSGERR_INVALIDO
    ElseIf Len(value) = 0 Then
        value = "Grupo" & NumGrupo
    End If
    mProps.EtiqPCGrupoP(NumGrupo) = value
    mbooModificado = True
End Property

Public Property Get EtiqPCGrupoP(NumGrupo As Integer) As String
'----- NumGrupo = Entre 1 a 4
    EtiqPCGrupoP = Trim$(mProps.EtiqPCGrupoP(NumGrupo))
End Property

Public Property Let EtiqPCGrupoE(NumGrupo As Integer, ByVal value As String)
'----- NumGrupo = Entre 1 a 4
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.EtiqPCGrupoE", MSGERR_INVALIDO
    ElseIf Len(value) = 0 Then
        value = "Grupo" & NumGrupo
    End If
    mProps.EtiqPCGrupoE(NumGrupo) = value
    mbooModificado = True
End Property

Public Property Get EtiqPCGrupoE(NumGrupo As Integer) As String
'----- NumGrupo = Entre 1 a 4
    EtiqPCGrupoE = Trim$(mProps.EtiqPCGrupoE(NumGrupo))
End Property

Public Property Let EtiqPCGrupoG(NumGrupo As Integer, ByVal value As String)
'----- NumGrupo = Entre 1 a 4
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.EtiqPCGrupoG", MSGERR_INVALIDO
    ElseIf Len(value) = 0 Then
        value = "Grupo" & NumGrupo
    End If
    mProps.EtiqPCGrupoG(NumGrupo) = value
    mbooModificado = True
End Property

Public Property Get EtiqPCGrupoG(NumGrupo As Integer) As String
'----- NumGrupo = Entre 1 a 4
    EtiqPCGrupoG = Trim$(mProps.EtiqPCGrupoG(NumGrupo))
End Property

Public Property Let FormatoCodEmp(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.FormatoCodEmp", MSGERR_INVALIDO
    End If
    mProps.FormatoCodEmp = value
    mbooModificado = True
End Property
Public Property Get FormatoCodEmp() As String
    FormatoCodEmp = Trim$(mProps.FormatoCodEmp)
End Property
Public Property Let GeneraCodEmp(ByVal value As Byte)
    If value > 4 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.GeneraCodEmp", MSGERR_INVALIDO
    End If
    mProps.GeneraCodEmp = value
    mbooModificado = True
End Property
Public Property Get GeneraCodEmp() As Byte
    GeneraCodEmp = mProps.GeneraCodEmp
End Property
Public Property Let PrefijoCodEmp(ByVal value As String)
    If Len(value) > 3 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.PrefijoCodEmp", MSGERR_INVALIDO
    End If
    mProps.PrefijoCodEmp = value
    mbooModificado = True
End Property
Public Property Get PrefijoCodEmp() As String
    PrefijoCodEmp = Trim$(mProps.PrefijoCodEmp)
End Property


Public Property Let GeneraCodAFItemALT(ByVal value As Byte)
    If value > 4 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.GeneraCodAFItemALT", MSGERR_INVALIDO
    End If
    mProps.GeneraCodAFItemALT = value
    mbooModificado = True
End Property

Public Property Get GeneraCodAFItemALT() As Byte
    GeneraCodAFItemALT = mProps.GeneraCodAFItemALT
End Property

Public Property Let PrefijoCodAFItemALT(ByVal value As String)
    If Len(value) > 3 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.PrefijoCodAFItemALT", MSGERR_INVALIDO
    End If
    mProps.PrefijoCodAFItemALT = value
    mbooModificado = True
End Property

Public Property Get PrefijoCodAFItemALT() As String
    PrefijoCodAFItemALT = Trim$(mProps.PrefijoCodAFItemALT)
End Property

Public Property Let FormatoCodAFItemALT(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.FormatoCodAFItemALT", MSGERR_INVALIDO
    End If
    mProps.FormatoCodAFItemALT = value
    mbooModificado = True
End Property

Public Property Get FormatoCodAFItemALT() As String
    FormatoCodAFItemALT = Trim$(mProps.FormatoCodAFItemALT)
End Property

'Empleado losep
Public Property Let FormatoCodEmpLosep(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.FormatoCodEmpLosep", MSGERR_INVALIDO
    End If
    mProps.FormatoCodEmpLosep = value
    mbooModificado = True
End Property
Public Property Get FormatoCodEmpLosep() As String
    FormatoCodEmpLosep = Trim$(mProps.FormatoCodEmpLosep)
End Property
Public Property Let GeneraCodEmpLosep(ByVal value As Byte)
    If value > 4 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.GeneraCodEmpLosep", MSGERR_INVALIDO
    End If
    mProps.GeneraCodEmpLosep = value
    mbooModificado = True
End Property
Public Property Get GeneraCodEmpLosep() As Byte
    GeneraCodEmpLosep = mProps.GeneraCodEmpLosep
End Property
Public Property Let PrefijoCodEmpLosep(ByVal value As String)
    If Len(value) > 3 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.PrefijoCodEmpLosep", MSGERR_INVALIDO
    End If
    mProps.PrefijoCodEmpLosep = value
    mbooModificado = True
End Property
Public Property Get PrefijoCodEmpLosep() As String
    PrefijoCodEmpLosep = Trim$(mProps.PrefijoCodEmpLosep)
End Property

Public Property Let EtiqPCGrupoL(NumGrupo As Integer, ByVal value As String)
'----- NumGrupo = Entre 1 a 4
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.EtiqPCGrupoL", MSGERR_INVALIDO
    ElseIf Len(value) = 0 Then
        value = "Grupo" & NumGrupo
    End If
    mProps.EtiqPCGrupoL(NumGrupo) = value
    mbooModificado = True
End Property

Public Property Get EtiqPCGrupoL(NumGrupo As Integer) As String
'----- NumGrupo = Entre 1 a 4
    EtiqPCGrupoL = Trim$(mProps.EtiqPCGrupoL(NumGrupo))
End Property

Public Property Let VersionSQL(ByVal value As String)
    mProps.VersionSQL = value
    mbooModificado = True
End Property

Public Property Get VersionSQL() As String
    
    Dim sql As String, rs As Recordset
    
    sql = "SELECT substring(@@version,1,25)  "
    Set rs = mobjEmpresa.OpenRecordsetParaEdit(sql)
    
    VersionSQL = rs.Fields(0)
    
End Property

Public Property Let NumeroLicencias(ByVal value As Integer)
    mProps.NumeroLicencias = value
    mbooModificado = True
End Property

Public Property Get NumeroLicencias() As Integer
    
    Dim sql As String, rs As Recordset
    
    sql = "SELECT isnull(xx,0) from gnopcion  "
    Set rs = mobjEmpresa.OpenRecordsetParaEdit(sql)
    
    NumeroLicencias = rs.Fields(0)
    
End Property


Public Property Let CodigoSC(ByVal value As String)
    If mProps.CodigoSC <> value Then
        mProps.CodigoSC = value
        mbooModificado = True
    End If
End Property

Public Property Get CodigoSC() As String
    CodigoSC = mProps.CodigoSC
End Property

Public Property Let IdCuentaDiferidaPre(ByVal value As Long)
    mProps.IdCuentaDiferidaPre = value
    mbooModificado = True
End Property

Public Property Get IdCuentaDiferidaPre() As Long
    IdCuentaDiferidaPre = mProps.IdCuentaDiferidaPre
End Property

Public Property Let CodCuentaDiferidaPre(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodCuentaDiferidaPre Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaDiferidaPre también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mobjEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNOpcion.CodCuentaDiferidaPre", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaDiferidaPre = rs!IdCuenta
            End If
            rs.Close
        Else
            IdCuentaDiferidaPre = 0
       End If
    End If
    Set rs = Nothing
    mProps.CodCuentaDiferidaPre = value
    mbooModificado = True
End Property

Public Property Get CodCuentaDiferidaPre() As String
    CodCuentaDiferidaPre = Trim$(mProps.CodCuentaDiferidaPre)
End Property


Public Property Let BandDINARDAP(ByVal value As Boolean)
        mProps.BandDINARDAP = value
        mbooModificado = True
End Property

Public Property Get BandDINARDAP() As Boolean
    BandDINARDAP = mProps.BandDINARDAP
End Property

Public Property Let BandCHP(ByVal value As Boolean)
        mProps.BandCHP = value
        mbooModificado = True
End Property

Public Property Get BandCHP() As Boolean
    BandCHP = mProps.BandCHP
End Property

Public Property Get OrdenBodegas() As String
    OrdenBodegas = Trim$(mProps.OrdenBodegas)
End Property

Public Property Let OrdenBodegas(ByVal value As String)
    If Len(value) > 250 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.OrdenBodegas", MSGERR_INVALIDO
    End If
    mProps.OrdenBodegas = value
    mbooModificado = True
End Property


Public Property Let BandOrdenItemsCodigo(ByVal value As Boolean)
        mProps.BandOrdenItemsCodigo = value
        mbooModificado = True
End Property

Public Property Get BandOrdenItemsCodigo() As Boolean
    BandOrdenItemsCodigo = mProps.BandOrdenItemsCodigo
End Property

Public Property Let BandOrdenGrupoItemsCodigo(ByVal value As Boolean)
        mProps.BandOrdenGrupoItemsCodigo = value
        mbooModificado = True
End Property

Public Property Get BandOrdenGrupoItemsCodigo() As Boolean
    BandOrdenGrupoItemsCodigo = mProps.BandOrdenGrupoItemsCodigo
End Property

Public Property Let BandOrdenPCCodigo(ByVal value As Boolean)
        mProps.BandOrdenPCCodigo = value
        mbooModificado = True
End Property

Public Property Get BandOrdenPCCodigo() As Boolean
    BandOrdenPCCodigo = mProps.BandOrdenPCCodigo
End Property

Public Property Let BandOrdenPCGrupoCodigo(ByVal value As Boolean)
        mProps.BandOrdenPCGrupoCodigo = value
        mbooModificado = True
End Property

Public Property Get BandOrdenPCGrupoCodigo() As Boolean
    BandOrdenPCGrupoCodigo = mProps.BandOrdenPCGrupoCodigo
End Property

Public Property Let BandFactElect(ByVal value As Boolean)
        mProps.BandFactElect = value
        mbooModificado = True
End Property

Public Property Get BandFactElect() As Boolean
    BandFactElect = mProps.BandFactElect
End Property

'/* GRV 11/AGO/2014 - INICIO */
Public Property Let ComprobantesGenerados(ByVal value As String)
    mProps.ComprobantesGenerados = value
    mbooModificado = True
End Property

Public Property Get ComprobantesGenerados() As String
    ComprobantesGenerados = Trim$(mProps.ComprobantesGenerados)
End Property

Public Property Let ComprobantesFirmados(ByVal value As String)
    mProps.ComprobantesFirmados = value
    mbooModificado = True
End Property

Public Property Get ComprobantesFirmados() As String
    ComprobantesFirmados = Trim$(mProps.ComprobantesFirmados)
End Property

Public Property Let ComprobantesAutorizados(ByVal value As String)
    mProps.ComprobantesAutorizados = value
    mbooModificado = True
End Property

Public Property Get ComprobantesAutorizados() As String
    ComprobantesAutorizados = Trim$(mProps.ComprobantesAutorizados)
End Property

Public Property Let ComprobantesNoAutorizados(ByVal value As String)
    mProps.ComprobantesNoAutorizados = value
    mbooModificado = True
End Property

Public Property Get ComprobantesNoAutorizados() As String
    ComprobantesNoAutorizados = Trim$(mProps.ComprobantesNoAutorizados)
End Property

Public Property Let UbicacionArchivoToken(ByVal value As String)
    mProps.UbicacionArchivoToken = value
    mbooModificado = True
End Property

Public Property Get UbicacionArchivoToken() As String
    UbicacionArchivoToken = Trim$(mProps.UbicacionArchivoToken)
End Property

Public Property Let ContrasenaToken(ByVal value As String)
    mProps.ContrasenaToken = value
    mbooModificado = True
End Property

Public Property Get ContrasenaToken() As String
    ContrasenaToken = Trim$(mProps.ContrasenaToken)
End Property

Public Property Let TipoAmbiente(ByVal value As String)
    mProps.TipoAmbiente = value
    mbooModificado = True
End Property

Public Property Get TipoAmbiente() As String
    TipoAmbiente = Trim$(mProps.TipoAmbiente)
End Property

'/* GRV 11/AGO/2014 - FIN */

Public Property Let ServidorCorreo(ByVal value As String)
    mProps.ServidorCorreo = value
    mbooModificado = True
End Property

Public Property Get ServidorCorreo() As String
    ServidorCorreo = Trim$(mProps.ServidorCorreo)
End Property

Public Property Let CuentaCorreo(ByVal value As String)
    mProps.CuentaCorreo = value
    mbooModificado = True
End Property

Public Property Get CuentaCorreo() As String
    CuentaCorreo = Trim$(mProps.CuentaCorreo)
End Property

Public Property Let PasswordCorreo(ByVal value As String)
    mProps.PasswordCorreo = value
    mbooModificado = True
End Property

Public Property Get PasswordCorreo() As String
    PasswordCorreo = Trim$(mProps.PasswordCorreo)
End Property

Public Property Let MensajeCorreo(ByVal value As String)
    mProps.MensajeCorreo = value
    mbooModificado = True
End Property

Public Property Get MensajeCorreo() As String
    MensajeCorreo = Trim$(mProps.MensajeCorreo)
End Property

Public Property Let ComprobantesEnviados(ByVal value As String)
    mProps.ComprobantesEnviados = value
    mbooModificado = True
End Property

Public Property Get ComprobantesEnviados() As String
    ComprobantesEnviados = Trim$(mProps.ComprobantesEnviados)
End Property

Public Property Let PuertoCorreo(ByVal value As String)
    mProps.PuertoCorreo = value
    mbooModificado = True
End Property

Public Property Get PuertoCorreo() As String
    PuertoCorreo = Trim$(mProps.PuertoCorreo)
End Property

Public Property Let CopiaCorreo(ByVal value As String)
    mProps.CopiaCorreo = value
    mbooModificado = True
End Property

Public Property Get CopiaCorreo() As String
    CopiaCorreo = Trim$(mProps.CopiaCorreo)
End Property

Public Property Let NombreUsuario(ByVal value As String)
    mProps.NombreUsuario = value
    mbooModificado = True
End Property

Public Property Get NombreUsuario() As String
    NombreUsuario = Trim$(mProps.NombreUsuario)
End Property


Public Property Let BandConexionSegura(ByVal value As Boolean)
        mProps.BandConexionSegura = value
        mbooModificado = True
End Property

Public Property Get BandConexionSegura() As Boolean
    BandConexionSegura = mProps.BandConexionSegura
End Property

Public Property Let ComprobantesSubidos(ByVal value As String)
    mProps.ComprobantesSubidos = value
    mbooModificado = True
End Property

Public Property Get ComprobantesSubidos() As String
    ComprobantesSubidos = Trim$(mProps.ComprobantesSubidos)
End Property

Public Property Let BandPortal(ByVal value As Boolean)
        mProps.BandPortal = value
        mbooModificado = True
End Property

Public Property Get BandPortal() As Boolean
    BandPortal = mProps.BandPortal
End Property


Public Property Let BandCorreoAutomatico(ByVal value As Boolean)
        mProps.BandCorreoAutomatico = value
        mbooModificado = True
End Property

Public Property Get BandCorreoAutomatico() As Boolean
    BandCorreoAutomatico = mProps.BandCorreoAutomatico
End Property

Public Property Let BandCorreoxEquipo(ByVal value As Boolean)
        mProps.BandCorreoxEquipo = value
        mbooModificado = True
End Property

Public Property Get BandCorreoxEquipo() As Boolean
    BandCorreoxEquipo = mProps.BandCorreoxEquipo
End Property


Public Property Get PenalizaDesctoxItem() As String
    PenalizaDesctoxItem = Trim$(mProps.PenalizaDesctoxItem)
End Property

Public Property Let PenalizaDesctoxItem(ByVal value As String)
    If Len(value) > 100 Then
        Err.Raise ERR_INVALIDO, "GNOpcion.PenalizaDesctoxItem", MSGERR_INVALIDO
    End If
    mProps.PenalizaDesctoxItem = value
    mbooModificado = True
End Property

Public Property Let BandRebateIVG(ByVal value As Boolean)
        mProps.BandRebateIVG = value
        mbooModificado = True
End Property

Public Property Get BandRebateIVG() As Boolean
    BandRebateIVG = mProps.BandRebateIVG
End Property

Public Property Let NumIVGRebate(ByVal value As Integer)
    mProps.NumIVGRebate = value
    mbooModificado = True
End Property

Public Property Get NumIVGRebate() As Integer
    NumIVGRebate = mProps.NumIVGRebate
End Property


Public Property Let BandOmitirCambioIVA(ByVal value As Boolean)
        mProps.BandOmitirCambioIVA = value
        mbooModificado = True
End Property

Public Property Get BandOmitirCambioIVA() As Boolean
    BandOmitirCambioIVA = mProps.BandOmitirCambioIVA
End Property

Public Property Let BandCostoUCParaCosto(ByVal value As Boolean)
        mProps.BandCostoUCParaCosto = value
        mbooModificado = True
End Property

Public Property Get BandCostoUCParaCosto() As Boolean
    BandCostoUCParaCosto = mProps.BandCostoUCParaCosto
End Property

Public Property Let IVKTipoDatoDouble(ByVal value As Boolean)
        mProps.IVKTipoDatoDouble = value
        mbooModificado = True
End Property

Public Property Get IVKTipoDatoDouble() As Boolean
    IVKTipoDatoDouble = mProps.IVKTipoDatoDouble
End Property



Public Property Let BandConfigLira(ByVal value As Boolean)
        mProps.BandConfigLira = value
        mbooModificado = True
End Property

Public Property Get BandConfigLira() As Boolean
    BandConfigLira = mProps.BandConfigLira
End Property


Public Sub GrabarSoloGnOpcion2()
    Dim sql As String, i As Integer
    Dim rs As Recordset
    Dim NumAutOld As String, FechaCadOld As Date, FechaAutold As Date
    
    'Si no está modificado no hace nada
    GrabarGNOpcion2

End Sub



Public Property Let PorDescGeneral(ByVal value As Currency)
        mProps.PorDescGeneral = value
        mbooModificado = True
End Property

Public Property Get PorDescGeneral() As Currency
    PorDescGeneral = mProps.PorDescGeneral
End Property

Public Property Let BandPorDescGeneral(ByVal value As Boolean)
        mProps.BandPorDescGeneral = value
        mbooModificado = True
End Property

Public Property Get BandPorDescGeneral() As Boolean
    BandPorDescGeneral = mProps.BandPorDescGeneral
End Property



Public Property Let BandDbleClickNOMod(ByVal value As Currency)
        mProps.BandDbleClickNOMod = value
        mbooModificado = True
End Property

Public Property Get BandDbleClickNOMod() As Currency
    BandDbleClickNOMod = mProps.BandDbleClickNOMod
End Property

Public Property Let RutaRepCarteraGen(ByVal value As String)
    mProps.RutaRepCarteraGen = value
    mbooModificado = True
End Property

Public Property Get RutaRepCarteraGen() As String
    RutaRepCarteraGen = Trim$(mProps.RutaRepCarteraGen)
End Property

Public Property Let MensajeCorreoCartera(ByVal value As String)
    mProps.MensajeCorreoCartera = value
    mbooModificado = True
End Property

Public Property Get MensajeCorreoCartera() As String
    MensajeCorreoCartera = Trim$(mProps.MensajeCorreoCartera)
End Property

Public Property Let MensajePDFCartera(ByVal value As String)
    mProps.MensajePDFCartera = value
    mbooModificado = True
End Property

Public Property Get MensajePDFCartera() As String
    MensajePDFCartera = Trim$(mProps.MensajePDFCartera)
End Property

