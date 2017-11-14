VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GNTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Empresa"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type T_PROPS
    CodTrans As String
    NombreTrans As String
    Descripcion As String
    Modulo As String
    CodPantalla As String
    CodTransFuente As String
    NumTransSiguiente As Long
    FechaAuto As Boolean
    HoraAuto As Boolean
    HoraVisible As Boolean
    ArchivoReporte As String
    IdResponsablePre As Long
    CodResponsablePre As String
    NecesitaAprobacion As Boolean
    RequiereCentroCosto As Boolean
    RequiereFuente As Boolean
    ImprimeComprobante As String * 1
    ImprimeSoloAprobado As Boolean
    NumTransModificable As Boolean
    MuestraAsiento As Boolean
    Opcion As String
    BandEliminable As Boolean
    BandValida As Boolean
    BandVisible As Boolean
    IdBodegaPre As Long
    CodBodegaPre As String
    IdProveedorPre As Long
    CodProveedorPre As String
    IdClientePre As Long
    CodClientePre As String
    CantidadPre As Currency
    IdFormaPre As Long
    CodFormaPre As String
    MonedaPre As String
    ColVisible As Long
    ColEditable As Long
    ColSeleccionable As Long
    AfectaCantidad As Boolean
    AfectaSaldoPC  As Boolean
    BandImportacion As Long
    ImportaSoloAprobado As Boolean
    VerificaLimiteCon As String           '*** MAKOTO 14/oct/00 Agregado
    IVPrecioxCliente As Boolean           '***ALEX 18/mar/2003 Agregado
    IVGrupoClientesParaPrecio  As Integer '***ALEX 20/mar/2003 Agregado
    fechagrabado As Date
    CotizacionAuto As Boolean       'jeaa 20/05/2005 Agregado
    CodTransHijoCI As String        'jeaa 26/07/2005 Agregado
    CodTransHijoRUC As String        'jeaa 26/07/2005 Agregado
    CodTransHijoOtro As String        'jeaa 26/07/2005 Agregado
    CodTransHijoPasp As String        'jeaa 31/01/2006 Agregado
    CodLineaDevolucion As String        'jeaa 16/11/2005 Agregado
    NumSerieEstablecimiento As String       'jeaa 04/04/2006 Agregado
    NumSeriePunto As String                         'jeaa 04/04/2006 Agregado
    NumAutorizacion As String                       'jeaa 04/04/2006 Agregado
    FechaCaducidad As Date                        'jeaa 04/04/2006 Agregado
    FechaDevolucion As Boolean    'AUC 13/03/06 fecha de devolucion
    ivDevolxSubitems As Boolean   'AUC 13/03/06 para escojer si devolvemos por items o subitems
    ivItemMenor As Boolean        'AUC 14/03/06 para mostrar la existencia por el item menor
    AnexoCodTipoTrans As String
    AnexoCodTipoComp As String
    LimiteDias As Currency        'AUC 05/05/06 para mostrar dias para alquiler
    ItemsImportantes As Boolean        'AUC para controlar por importancia de items en una preparacion
    NumTransUltimo As String  ' jeaa 15/11/2006 se guardará desde que numero inicio esta autorizacion
    NumAutorizacionOld As String                       'jeaa 15/11/2006 Agregado
    FechaCaducidadOld As Date                        'jeaa 15/11/2006 Agregado
    NumTransInicio As String
    NumTransInicioOld As String
    CodTransImp As String        'jeaa 26/07/2005 Agregado
    AnexoCodTipoRetencion As String
    ImportacionUnicaTrans As Boolean
    TransDespuesDevolucion As String        'AUC 05/05/06 para mostrar dias para alquiler (Hormipiosos igual que VerificaLimiteCon pero una lista)
    TransHistorial As String        'AUC 05/05/06 para mostrar dias para alquiler (Hormipiosos igual que VerificaLimiteCon pero una lista)
    FormasLimteCredito As String   'jeaa 20 de julio del 2007 para formas de cobro en limite de credito
    NumTransUltimoPreimpreso As String   'jeaa 25/09/2007
    MargenMinimoUtil As Currency 'AUC 26/11/07 Guarda el Minimo porcentaje de utilidad
    IdGastoPre As Long          'jeaa 07/02/2008
    CodGastoPre As String      'jeaa 07/02/2008
    IdBodegaDestPre As Long         'jeaa 26/03/2008
    CodBodegaDestPre As String      'jeaa 26/03/2008
    TipoTrans As String
    IDSucursal As Long                'jeaa 05/02/2009
    CodSucursal As String           'jeaa 05/02/2009
    NumIvGrupoFiltro As Long    'jeaa 12/06/2009
    IdGruposFiltro As String        'jeaa 12/06/2009
    NumIvGrupoTrabajo As Long    'jeaa 30/07/2009 isollanta
    IdGruposTrabajo As String        'jeaa 30/06/2009
    NumIvGrupoTamaño As Long    'jeaa 30/07/2009 isollanta
    IdGruposTamaño As String        'jeaa 30/06/2009
    NumIvGrupoMarca As Long    'jeaa 30/07/2009 isollanta
    IdGruposMarca As String        'jeaa 30/06/2009
    RutaLibImp As String       'jeaa 27/11/2009
    BandImpGlobal As Boolean       'jeaa 27/11/2009
    objEmpresa As Empresa
    CodBodegaDesPre As String 'AUC para la bodega destino en los alq/Res
    IdBodegaDesPre As Long
    ImportacionSoloSaldosInv As Boolean 'Importacion solo saldos inventario
    HistPorTrans As Boolean
    HistPorItem As Boolean
    IVTransDes As String 'AUC genera transaccion de desperdicio
    FechaAutorizacion As Date                        'jeaa 04/04/2006 Agregado
    FechaAutorizacionOld As Date                        'jeaa 15/11/2006 Agregado
    AsientoTrans As String          ' jeaa 28/02/2011 autoimpresores
    AsientoTransAnulada As String
    BandNuevoAutoImpresor As Boolean
    AnexoSustento As String
    IdInventarioPre As Long 'AUC 31/12/2011
    CodInventarioPre As String
    Opcion2 As String
    AsientoGastoxZona As Boolean
    OmitirGastoxZona As Boolean
    CentroCostoDetalleVisible  As Boolean
    IdBodegaFiltro As String
    IdPrecioFiltro As String
    NombreFuente  As String
    TamanioFuente  As Integer
    AfectaCantidadDocumento As Boolean
   FormaPagoSRI As String
     ColVisibleNew As Long 'rocar
    ColEditableNew As Long
    ColSeleccionableNew As Long
    TSTransAutoDescMora As String
    TipoInteres As Boolean 'Int Simple/Compuesto 'AUC 08/02/2011
    CierreTodosU As Boolean
    ImpTransSolCreApr As Boolean
    ArchivoReporteRIDE As String
    BandCompElectonico As Boolean
    IdEmpleadoPre As Long
    CodEmpleadoPre As String
    NumCopiasImp As Integer
    AsientoxPCG As Boolean
    ConfigBotones As String
    RutaXMLOrigen As String
    RutaXMLDestino As String
    NumRefSig As Long

End Type

Private mProps As T_PROPS
Private mcolAsientos As Collection
Private mcolRecargos As Collection
Private mcolAsientosPresup As Collection
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Public Sub Grabar()
    Dim sql As String, i As Long, fgrabado As Date
    Dim rs As Recordset
    Dim gnSuc As GNSucursal
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Comienza la transaccion
#If DAOLIB Then
    gobjWS.BeginTrans
#Else
    mProps.objEmpresa.Coneccion.BeginTrans
#End If
    On Error GoTo ErrTrap
    
        'Verifica si no repite el numseriepunto
        If Me.IVAutoImpresor Then
            sql = "SELECT CodTrans FROM GNTrans gt left join gnsucursal gs on gt.idsucursal= gs.idsucursal"
            sql = sql & "  WHERE TipoTrans='" & mProps.TipoTrans & "' and NumSeriePunto='" & mProps.NumSeriePunto & "'"
            sql = sql & " and codsucursal='" & mProps.CodSucursal & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            
            If Not rs.EOF Then      'Si ya existe el codigo
                If rs.Fields("CODTRANS") <> Me.CodTrans Then
                    Err.Raise ERR_REPITENNUMPUNTO, "GNTrans", MSGERR_REPITENUMPUNTO
                End If
            End If
            rs.Close
            
            If Len(mProps.objEmpresa.GNOpcion.ObtenerValor("TramitesPosiblesSRI")) > 0 Then
                If mProps.objEmpresa.GNOpcion.ObtenerValor("TramitesPosiblesSRI") <> "A" Then
                    If Len(mProps.objEmpresa.GNOpcion.ObtenerValor("RealizarReporteRangos")) > 0 Then
                        If mProps.objEmpresa.GNOpcion.ObtenerValor("RealizarReporteRangos") = "1" Then
                            If mProps.objEmpresa.GNOpcion.ObtenerValor("TramitesPosiblesSRI") <> "E" And Me.EsNuevo Then
                                Err.Raise ERR_REPORTERAGO, "GNTrans", MSGERR_REPORTERANGO
                            End If
                        End If
                    End If
                End If
            End If
       End If
    
    
    If mbooEsNuevo Then
        'Verifica si no repite el código
        sql = "SELECT CodTrans FROM GNTrans WHERE CodTrans='" & mProps.CodTrans & "'"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then      'Si ya existe el codigo
            Err.Raise ERR_REPITECODIGO, "GNTrans", MSGERR_REPITECODIGO
        End If
        rs.Close
        
        
        sql = "SELECT * FROM GNTrans WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
        
        If Me.IVAutoImpresor Then
            Me.BandNuevoAutoImpresor = True
        End If
    Else
        'Elimina todos los registros relacionados en GNTransAsiento
        sql = "DELETE FROM GNTransAsiento WHERE CodTrans='" & mProps.CodTrans & "'"
        mProps.objEmpresa.Execute sql, True
    
        'Elimina todos los registros relacionados en GNTransRecargo
        sql = "DELETE FROM GNTransRecargo WHERE CodTrans='" & mProps.CodTrans & "'"
        mProps.objEmpresa.Execute sql, True
    
        If Len(mProps.objEmpresa.GNOpcion.ObtenerValor("Presupuesto")) > 0 Then
            If mProps.objEmpresa.GNOpcion.ObtenerValor("Presupuesto") = "1" Then
                'Elimina todos los registros relacionados en GNTransAsientoPresup
                sql = "DELETE FROM GNTransAsientoPresup WHERE CodTrans='" & mProps.CodTrans & "'"
                mProps.objEmpresa.Execute sql, True
            End If
        End If
        
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM GNTrans WHERE CodTrans='" & mProps.CodTrans & "'"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "GNTrans_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el GNTrans
    With rs
        !CodTrans = mProps.CodTrans
        !NombreTrans = mProps.NombreTrans
        !Descripcion = mProps.Descripcion
        !Modulo = mProps.Modulo
        !CodPantalla = mProps.CodPantalla
        !CodTransFuente = mProps.CodTransFuente
        'jeaa 08/03/2006
        If Not !NumTransSiguiente = mProps.NumTransSiguiente Then
            mProps.objEmpresa.GrabaAccionModulo !NumTransSiguiente, mProps.NumTransSiguiente, "NumTransSiguiente", "MOD_TRANS", "DT"
        End If
        !NumTransSiguiente = mProps.NumTransSiguiente
        !FechaAuto = mProps.FechaAuto
        !HoraAuto = mProps.HoraAuto
        !HoraVisible = mProps.HoraVisible
        !ArchivoReporte = mProps.ArchivoReporte
        !IdResponsablePre = mProps.IdResponsablePre
        !NecesitaAprobacion = mProps.NecesitaAprobacion
        !RequiereCentroCosto = mProps.RequiereCentroCosto
        !RequiereFuente = mProps.RequiereFuente
        !ImprimeComprobante = mProps.ImprimeComprobante
        !ImprimeSoloAprobado = mProps.ImprimeSoloAprobado
        !NumTransModificable = mProps.NumTransModificable
        !MuestraAsiento = mProps.MuestraAsiento
        !Opcion = Trim$(mProps.Opcion)
        !BandEliminable = mProps.BandEliminable
        !BandValida = mProps.BandValida
        !BandVisible = mProps.BandVisible
        !IdBodegaPre = mProps.IdBodegaPre
        !IdProveedorPre = mProps.IdProveedorPre
        !IdClientePre = mProps.IdClientePre
        !CantidadPre = mProps.CantidadPre
        !IdFormaPre = mProps.IdFormaPre
        !MonedaPre = mProps.MonedaPre
        !ColVisible = mProps.ColVisible
        !ColEditable = mProps.ColEditable
        !ColSeleccionable = mProps.ColSeleccionable
        !AfectaCantidad = mProps.AfectaCantidad
        !AfectaSaldoPC = mProps.AfectaSaldoPC
        !BandImportacion = mProps.BandImportacion
        !ImportaSoloAprobado = mProps.ImportaSoloAprobado
        !VerificaLimiteCon = mProps.VerificaLimiteCon       '*** MAKOTO 14/oct/00 Agregado
        !CotizacionAuto = mProps.CotizacionAuto                 'jeaa 20/06/2005
        !CodTransHijoCI = mProps.CodTransHijoCI             'jeaa 25/07/2005
        !CodTransHijoRUC = mProps.CodTransHijoRUC             'jeaa 25/07/2005
        !CodTransHijoPasp = mProps.CodTransHijoPasp                 'jeaa 31/01/2006
        !CodTransHijoOtro = mProps.CodTransHijoOtro             'jeaa 25/07/2005
        !CodLineaDevolucion = mProps.CodLineaDevolucion               'jeaa 15/11/2005
        If Not !NumSerieEstablecimiento = mProps.NumSerieEstablecimiento Then
            mProps.objEmpresa.GrabaAccionModulo !NumSerieEstablecimiento & "-" & !NumSeriePunto, mProps.NumSerieEstablecimiento & "-" & !NumSeriePunto, "NumSerieEstablecimiento", "MOD_TRANS", "DT"
        End If
        !NumSerieEstablecimiento = mProps.NumSerieEstablecimiento   'jeaa 04/04/2006
        If Not !NumSeriePunto = mProps.NumSeriePunto Then
            mProps.objEmpresa.GrabaAccionModulo !NumSerieEstablecimiento & "-" & !NumSeriePunto, !NumSerieEstablecimiento & "-" & mProps.NumSeriePunto, "NumSeriePunto", "MOD_TRANS", "DT"
        End If
        !NumSeriePunto = mProps.NumSeriePunto  'jeaa 04/04/2006
        If Not !NumAutorizacion = mProps.NumAutorizacion Then
            mProps.objEmpresa.GrabaAccionModulo !NumAutorizacion, mProps.NumAutorizacion, "NumAutorizacion", "MOD_TRANS", "DT"
        End If
        !NumAutorizacion = mProps.NumAutorizacion   'jeaa 04/04/2006
         If Not !FechaCaducidad = mProps.FechaCaducidad Then
            mProps.objEmpresa.GrabaAccionModulo !FechaCaducidad, mProps.FechaCaducidad, "FechaCaducidad", "MOD_TRANS", "DT"
        End If
        !FechaCaducidad = mProps.FechaCaducidad   'jeaa 04/04/2006
        
         If Not !FechaAutorizacion = mProps.FechaAutorizacion Then
            mProps.objEmpresa.GrabaAccionModulo !FechaAutorizacion, mProps.FechaAutorizacion, "FechaAutorizacion", "MOD_TRANS", "DT"
        End If
        !FechaAutorizacion = mProps.FechaAutorizacion   'jeaa 04/04/2006
        
        !LimiteDias = mProps.LimiteDias                     'AUC 05/05/06
        !TransDespuesDevolucion = mProps.TransDespuesDevolucion
        !TransHistorial = mProps.TransHistorial
        'jeaa 10/08/2006
        !AnexoCodTipoTrans = mProps.AnexoCodTipoTrans
        !AnexoCodTipoComp = mProps.AnexoCodTipoComp
        !FechaDevolucion = mProps.FechaDevolucion
        
        !NumTransUltimo = mProps.NumTransSiguiente - 1 'jeaa 15/11/2006
        !NumAutorizacionOld = mProps.NumAutorizacionOld   'jeaa 15/11/2006
        !FechaCaducidadOld = mProps.FechaCaducidadOld   'jeaa 15/11/2006
        !FechaAutorizacionOld = mProps.FechaAutorizacionOld   'jeaa 15/11/2006
        !NumTransInicio = mProps.NumTransInicio
        !NumTransInicioOld = mProps.NumTransInicioOld
        !CodTransImp = mProps.CodTransImp                    'jeaa 07/12/2005
        !AnexoCodTipoRetencion = mProps.AnexoCodTipoRetencion 'jeaa 18/12/2006
        !ImportacionUnicaTrans = mProps.ImportacionUnicaTrans 'jeaa 01/06/2007
        !TransDespuesDevolucion = mProps.TransDespuesDevolucion 'jeaa 20/06/2006
        !TransHistorial = mProps.TransHistorial 'jeaa 20/06/2006
        !FormasLimteCredito = mProps.FormasLimteCredito 'jeaa 20/07/2006
        !NumTransUltimoPreimpreso = mProps.NumTransUltimoPreimpreso  'jeaa 25/09/2007
        !MargenMinimoUtil = mProps.MargenMinimoUtil 'AUC 26/11/07
        !IdGastoPre = mProps.IdGastoPre  ' jeaa 07/01/2008
        If Me.CodPantalla = "IVRES_ALQ" Or Me.CodPantalla = "IVBQD2B" Or Me.CodPantalla = "IVDEV" Or Me.CodPantalla = "IVBQD2BE" Or Me.CodPantalla = "IVDEVE" Or Me.CodPantalla = "IVCAMIEREC" Then
            !IdBodegaDestPre = mProps.IdBodegaDesPre
        Else
            !IdBodegaDestPre = mProps.IdBodegaDestPre  'jeaa 26/03/2008
        End If
        
        If Not !TipoTrans = mProps.TipoTrans Then
            mProps.objEmpresa.GrabaAccionModulo !NumSerieEstablecimiento & "-" & !NumSeriePunto & " Tipo Comprobante: " & !TipoTrans, !NumSerieEstablecimiento & "-" & mProps.NumSeriePunto & " Tipo Comprobante: " & mProps.TipoTrans, "Tipo Comprobante", "MOD_TRANS", "DT"
        End If
        !TipoTrans = mProps.TipoTrans
         If Not !IDSucursal = mProps.IDSucursal Then
            Set gnSuc = mProps.objEmpresa.RecuperaGNSucursal(mProps.IDSucursal)
            If Not gnSuc Is Nothing Then
                mProps.objEmpresa.GrabaAccionModulo gnSuc.CodSucursal, mProps.CodSucursal, "Codigo Sucursal", "MOD_TRANS", "DT"
            Else
                mProps.objEmpresa.GrabaAccionModulo "Vacio", mProps.CodSucursal, "Codigo Sucursal", "MOD_TRANS", "DT"
            End If
            Set gnSuc = Nothing
        End If
        !IDSucursal = mProps.IDSucursal
        !NumIvGrupoFiltro = mProps.NumIvGrupoFiltro  ' jeaa 12/06/2009
        !IdGruposFiltro = mProps.IdGruposFiltro  ' jeaa 12/06/2009
        !NumIvGrupoTrabajo = mProps.NumIvGrupoTrabajo  ' jeaa 30/07/2009
        !IdGruposTrabajo = mProps.IdGruposTrabajo  ' jeaa 30/07/2009
        !NumIvGrupoTamanio = mProps.NumIvGrupoTamaño  ' jeaa 30/07/2009
        !IdGruposTamanio = mProps.IdGruposTamaño  ' jeaa 30/07/2009
        !NumIvGrupoMarca = mProps.NumIvGrupoMarca  ' jeaa 30/07/2009
        !IdGruposMarca = mProps.IdGruposMarca  ' jeaa 30/07/2009
        !RutaLibImp = mProps.RutaLibImp
        !BandImpGlobal = mProps.BandImpGlobal
        !BandNuevoAutoImpresor = mProps.BandNuevoAutoImpresor
        !AnexoSustento = mProps.AnexoSustento
        !IdInventarioPre = mProps.IdInventarioPre  'AUC 31/12/2011
        !Opcion2 = Trim$(mProps.Opcion2)
        !AsientoGastoxZona = mProps.AsientoGastoxZona
        !OmitirGastoxZona = mProps.OmitirGastoxZona
        !IdBodegaFiltro = mProps.IdBodegaFiltro  ' jeaa 12/06/2009
        !IdPrecioFiltro = mProps.IdPrecioFiltro  ' jeaa 12/06/2009
        !NombreFuente = mProps.NombreFuente  ' jeaa 12/06/2009
        !TamanioFuente = mProps.TamanioFuente  ' jeaa 12/06/2009
        !AfectaCantidadDocumento = mProps.AfectaCantidadDocumento
        !FormaPagoSRI = mProps.FormaPagoSRI
        !ColVisibleNew = mProps.ColVisibleNew
        !ColEditableNew = mProps.ColEditableNew
        !ColSeleccionableNew = mProps.ColSeleccionableNew
        !ArchivoReporteRIDE = mProps.ArchivoReporteRIDE
        !IdEmpleadoPre = mProps.IdEmpleadoPre
        !NumCopiasImp = mProps.NumCopiasImp
        !AsientoxPCG = mProps.AsientoxPCG
        !ConfigBotones = mProps.ConfigBotones
        !RutaXMLOrigen = mProps.RutaXMLOrigen
        !RutaXMLDestino = mProps.RutaXMLDestino
        !NumRefSig = mProps.NumRefSig

        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        fgrabado = !fechagrabado
        .Close
    End With
    
    
    'Graba todos los Asientos modelos
    For i = 1 To CountAsiento
        Asientos(i).Grabar
    Next i
    'Graba todos los Recargos
    For i = 1 To CountRecargo
        Recargos(i).Grabar
    Next i
    
    If Len(mProps.objEmpresa.GNOpcion.ObtenerValor("Presupuesto")) > 0 Then
        If mProps.objEmpresa.GNOpcion.ObtenerValor("Presupuesto") = "1" Then
            For i = 1 To CountAsientoPresup
                AsientosPresup(i).Grabar
            Next i
        End If
    End If
    
    
    'Concluye la transacción
#If DAOLIB Then
    gobjWS.CommitTrans
#Else
    mProps.objEmpresa.Coneccion.CommitTrans
#End If
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_TRANS", "Transacción Nueva: " & _
                                                         mProps.CodTrans, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_TRANS", "Transacción Modificada: " & _
                                                         mProps.CodTrans, "DT"
    End If
        
    If mbooEsNuevo Then
        mProps.objEmpresa.GNOpcion.AsignarValor "RealizarReporteRangos", "1"
        If mProps.objEmpresa.GNOpcion.ObtenerValor("TramitesPosiblesSRI") <> "A" Then
            mProps.objEmpresa.GNOpcion.AsignarValor "TramitesPosiblesSRI", "E"
        End If
        mProps.objEmpresa.GNOpcion.AsignarValor "ValidacionAutoimpresores", "0"
        mProps.objEmpresa.GNOpcion.Grabar
    End If
        
    mProps.fechagrabado = fgrabado
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_TRANS
    Exit Sub
    
ErrTrap:
    'Cancela la transacción
#If DAOLIB Then
    gobjWS.Rollback
#Else
    mProps.objEmpresa.Coneccion.RollbackTrans
#End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Friend Sub Recuperar()
    Dim sql As String, i As Long
    Dim rs As Recordset
    
    sql = "SELECT * FROM vwGNTrans WHERE CodTrans='" & CodTrans & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de GNTrans
    With rs
        If Not IsNull(!NombreTrans) Then mProps.NombreTrans = !NombreTrans
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!Modulo) Then mProps.Modulo = !Modulo
        If Not IsNull(!CodPantalla) Then mProps.CodPantalla = !CodPantalla
        If Not IsNull(!CodTransFuente) Then mProps.CodTransFuente = !CodTransFuente
        If Not IsNull(!NumTransSiguiente) Then mProps.NumTransSiguiente = !NumTransSiguiente
        If Not IsNull(!FechaAuto) Then mProps.FechaAuto = !FechaAuto
        If Not IsNull(!HoraAuto) Then mProps.HoraAuto = !HoraAuto
        If Not IsNull(!HoraVisible) Then mProps.HoraVisible = !HoraVisible
        If Not IsNull(!ArchivoReporte) Then mProps.ArchivoReporte = !ArchivoReporte
        If Not IsNull(!IdResponsablePre) Then mProps.IdResponsablePre = !IdResponsablePre
        If Not IsNull(!CodResponsablePre) Then mProps.CodResponsablePre = !CodResponsablePre
        If Not IsNull(!NecesitaAprobacion) Then mProps.NecesitaAprobacion = !NecesitaAprobacion
        If Not IsNull(!RequiereCentroCosto) Then mProps.RequiereCentroCosto = !RequiereCentroCosto
        If Not IsNull(!RequiereFuente) Then mProps.RequiereFuente = !RequiereFuente
        If Not IsNull(!ImprimeComprobante) Then mProps.ImprimeComprobante = !ImprimeComprobante
        If Not IsNull(!ImprimeSoloAprobado) Then mProps.ImprimeSoloAprobado = !ImprimeSoloAprobado
        If Not IsNull(!NumTransModificable) Then mProps.NumTransModificable = !NumTransModificable
        If Not IsNull(!MuestraAsiento) Then mProps.MuestraAsiento = !MuestraAsiento
        If Not IsNull(!Opcion) Then mProps.Opcion = !Opcion
        'If Len(mProps.Opcion) < 120 Then mProps.Opcion = mProps.Opcion & String$(120 - Len(mProps.Opcion), " ")   'Rellana con espacio
        'If Len(mProps.Opcion) < 200 Then mProps.Opcion = mProps.Opcion & String$(200 - Len(mProps.Opcion), " ")   'Rellana con espacio
        If Len(mProps.Opcion) < 256 Then mProps.Opcion = mProps.Opcion & String$(256 - Len(mProps.Opcion), " ")   'Rellana con espacio
        If Not IsNull(!BandEliminable) Then mProps.BandEliminable = !BandEliminable
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!BandVisible) Then mProps.BandVisible = !BandVisible
        If Not IsNull(!IdBodegaPre) Then mProps.IdBodegaPre = !IdBodegaPre
        If Not IsNull(!CodBodegaPre) Then mProps.CodBodegaPre = !CodBodegaPre
        If Not IsNull(!IdProveedorPre) Then mProps.IdProveedorPre = !IdProveedorPre
        If Not IsNull(!CodProveedorPre) Then mProps.CodProveedorPre = !CodProveedorPre
        If Not IsNull(!IdClientePre) Then mProps.IdClientePre = !IdClientePre
        If Not IsNull(!CodClientePre) Then mProps.CodClientePre = !CodClientePre
        If Not IsNull(!CantidadPre) Then mProps.CantidadPre = !CantidadPre
        If Not IsNull(!IdFormaPre) Then mProps.IdFormaPre = !IdFormaPre
        If Not IsNull(!CodFormaPre) Then mProps.CodFormaPre = !CodFormaPre
        If Not IsNull(!MonedaPre) Then mProps.MonedaPre = !MonedaPre
        If Not IsNull(!ColVisible) Then mProps.ColVisible = !ColVisible
        If Not IsNull(!ColEditable) Then mProps.ColEditable = !ColEditable
        If Not IsNull(!ColSeleccionable) Then mProps.ColSeleccionable = !ColSeleccionable
        If Not IsNull(!AfectaCantidad) Then mProps.AfectaCantidad = !AfectaCantidad
        If Not IsNull(!AfectaSaldoPC) Then mProps.AfectaSaldoPC = !AfectaSaldoPC
        If Not IsNull(!BandImportacion) Then mProps.BandImportacion = !BandImportacion
        If Not IsNull(!ImportaSoloAprobado) Then mProps.ImportaSoloAprobado = !ImportaSoloAprobado
        If Not IsNull(!VerificaLimiteCon) Then mProps.VerificaLimiteCon = !VerificaLimiteCon
        If Not IsNull(!CotizacionAuto) Then mProps.CotizacionAuto = !CotizacionAuto
        If Not IsNull(!CodTransHijoCI) Then mProps.CodTransHijoCI = !CodTransHijoCI               'jeaa 25/07/2005
        If Not IsNull(!CodTransHijoRUC) Then mProps.CodTransHijoRUC = !CodTransHijoRUC      'jeaa 25/07/2005
        If Not IsNull(!CodTransHijoOtro) Then mProps.CodTransHijoOtro = !CodTransHijoOtro      'jeaa 25/07/2005
        If Not IsNull(!CodTransHijoPasp) Then mProps.CodTransHijoPasp = !CodTransHijoPasp      'jeaa 31/01/2006
        If Not IsNull(!CodLineaDevolucion) Then mProps.CodLineaDevolucion = !CodLineaDevolucion                        'jeaa 15/11/2005
        If Not IsNull(!NumSerieEstablecimiento) Then mProps.NumSerieEstablecimiento = !NumSerieEstablecimiento                           'jeaa 04/04/2006
        If Not IsNull(!NumSeriePunto) Then mProps.NumSeriePunto = !NumSeriePunto                           'jeaa 04/04/2006
        If Not IsNull(!NumAutorizacion) Then mProps.NumAutorizacion = !NumAutorizacion                           'jeaa 04/04/2006
        If Not IsNull(!FechaCaducidad) Then mProps.FechaCaducidad = !FechaCaducidad                           'jeaa 04/04/2006
        If Not IsNull(!FechaAutorizacion) Then mProps.FechaAutorizacion = !FechaAutorizacion                           'jeaa 04/04/2006
        If Not IsNull(!AnexoCodTipoTrans) Then mProps.AnexoCodTipoTrans = !AnexoCodTipoTrans                           'jeaa 10/08/2006
        If Not IsNull(!AnexoCodTipoComp) Then mProps.AnexoCodTipoComp = !AnexoCodTipoComp                           'jeaa 10/08/2006
        If Not IsNull(!FechaDevolucion) Then mProps.FechaDevolucion = !FechaDevolucion  'AUC 13/03/06
        If Not IsNull(!LimiteDias) Then mProps.LimiteDias = !LimiteDias   'AUC 05/05/06
        If Not IsNull(!TransDespuesDevolucion) Then mProps.TransDespuesDevolucion = !TransDespuesDevolucion    'AUC 05/05/06
        If Not IsNull(!TransHistorial) Then mProps.TransHistorial = !TransHistorial     'AUC 05/05/06
        If Not IsNull(!NumTransUltimo) Then
            mProps.NumTransUltimo = !NumTransUltimo   'jeaa 15/11/2006
        Else
            mProps.NumTransUltimo = "0"   'jeaa 15/11/2006
        End If
        If Not IsNull(!NumAutorizacionOld) Then
            mProps.NumAutorizacionOld = !NumAutorizacionOld                           'jeaa 15/11/2006
        Else
            mProps.NumAutorizacionOld = !NumAutorizacion                           'jeaa 15/11/2006
        End If
        If Not IsNull(!FechaCaducidadOld) Then
            mProps.FechaCaducidadOld = !FechaCaducidadOld                           'jeaa 15/11/2006
        Else
            mProps.FechaCaducidadOld = !FechaCaducidad                           'jeaa 15/11/2006
        End If
        
        If Not IsNull(!FechaAutorizacionOld) And Not IsNull(!FechaAutorizacion) Then
            mProps.FechaAutorizacionOld = !FechaAutorizacionOld                           'jeaa 15/11/2006
        Else
            If Not IsNull(!FechaAutorizacion) Then
                mProps.FechaAutorizacionOld = !FechaAutorizacion                           'jeaa 15/11/2006
            Else
                mProps.FechaAutorizacionOld = Date
            End If
        End If
        
        If Not IsNull(!AnexoCodTipoTrans) Then mProps.AnexoCodTipoTrans = !AnexoCodTipoTrans                           'jeaa 10/08/2006
        If Not IsNull(!AnexoCodTipoComp) Then mProps.AnexoCodTipoComp = !AnexoCodTipoComp                           'jeaa 10/08/2006
        If Not IsNull(!NumTransInicio) Then mProps.NumTransInicio = !NumTransInicio
        If Not IsNull(!NumTransInicioOld) Then
            mProps.NumTransInicioOld = !NumTransInicioOld
        Else
            mProps.NumTransInicioOld = "0"
        End If
        If Not IsNull(!CodTransImp) Then mProps.CodTransImp = !CodTransImp               'jeaa 07/12/2006
        If Not IsNull(!AnexoCodTipoRetencion) Then mProps.AnexoCodTipoRetencion = !AnexoCodTipoRetencion                           'jeaa 10/08/2006
        If Not IsNull(!ImportacionUnicaTrans) Then mProps.ImportacionUnicaTrans = !ImportacionUnicaTrans 'jeaa 01/06/2007
        If Not IsNull(!TransDespuesDevolucion) Then mProps.TransDespuesDevolucion = !TransDespuesDevolucion    'jeaa 01/06/2007
        If Not IsNull(!TransHistorial) Then mProps.TransHistorial = !TransHistorial    'jeaa 01/06/2007
        If Not IsNull(!FormasLimteCredito) Then mProps.FormasLimteCredito = !FormasLimteCredito    'jeaa 20/07/2007
        If Not IsNull(!NumTransUltimoPreimpreso) Then mProps.NumTransUltimoPreimpreso = !NumTransUltimoPreimpreso 'jeaa 25/09/2007
        If Not IsNull(!MargenMinimoUtil) Then mProps.MargenMinimoUtil = !MargenMinimoUtil 'AUC 26/11/07
        If Not IsNull(!IdGastoPre) Then mProps.IdGastoPre = !IdGastoPre 'jeaa 07/01/2008
        If Not IsNull(!CodGastoPre) Then mProps.CodGastoPre = !CodGastoPre 'jeaa 07/01/2008
        If Not IsNull(!IdBodegaDestPre) Then mProps.IdBodegaDestPre = !IdBodegaDestPre  'jeaa 26/03/2008
        If Not IsNull(!CodBodegaDestPre) Then mProps.CodBodegaDestPre = !CodBodegaDestPre  'jeaa 26/03/2008
        If Not IsNull(!TipoTrans) Then mProps.TipoTrans = !TipoTrans
        If Not IsNull(!IDSucursal) Then mProps.IDSucursal = !IDSucursal
        If Not IsNull(!NumIvGrupoFiltro) Then mProps.NumIvGrupoFiltro = !NumIvGrupoFiltro       'JEAA 12/06/2009
        If Not IsNull(!IdGruposFiltro) Then mProps.IdGruposFiltro = !IdGruposFiltro                     'JEAA 12/06/2009
        If Not IsNull(!NumIvGrupoTrabajo) Then mProps.NumIvGrupoTrabajo = !NumIvGrupoTrabajo       'JEAA 12/06/2009
        If Not IsNull(!IdGruposTrabajo) Then mProps.IdGruposTrabajo = !IdGruposTrabajo                     'JEAA 12/06/2009
        If Not IsNull(!NumIvGrupoTamanio) Then mProps.NumIvGrupoTamaño = !NumIvGrupoTamanio       'JEAA 12/06/2009
        If Not IsNull(!IdGruposTamanio) Then mProps.IdGruposTamaño = !IdGruposTamanio                     'JEAA 12/06/2009
        If Not IsNull(!NumIvGrupoMarca) Then mProps.NumIvGrupoMarca = !NumIvGrupoMarca       'JEAA 12/06/2009
        If Not IsNull(!IdGruposMarca) Then mProps.IdGruposMarca = !IdGruposMarca                     'JEAA 12/06/2009
        If Not IsNull(!RutaLibImp) Then mProps.RutaLibImp = !RutaLibImp
        If Not IsNull(!BandImpGlobal) Then mProps.BandImpGlobal = !BandImpGlobal
        If Not IsNull(!IdBodegaDestPre) Then mProps.IdBodegaDesPre = !IdBodegaDestPre
        If Not IsNull(!CodBodegaDestPre) Then mProps.CodBodegaDesPre = !CodBodegaDestPre
        If Not IsNull(!BandNuevoAutoImpresor) Then mProps.BandNuevoAutoImpresor = !BandNuevoAutoImpresor
        If Not IsNull(!AnexoSustento) Then mProps.AnexoSustento = !AnexoSustento
        If Not IsNull(!IdInventarioPre) Then mProps.IdInventarioPre = !IdInventarioPre
        If Not IsNull(!CodInventarioPre) Then mProps.CodInventarioPre = !CodInventarioPre
        If Not IsNull(!Opcion2) Then mProps.Opcion2 = !Opcion2
        If Len(mProps.Opcion2) < 256 Then mProps.Opcion2 = mProps.Opcion2 & String$(256 - Len(mProps.Opcion2), " ")   'Rellana con espacio
        If Not IsNull(!AsientoGastoxZona) Then mProps.AsientoGastoxZona = !AsientoGastoxZona
        If Not IsNull(!OmitirGastoxZona) Then mProps.OmitirGastoxZona = !OmitirGastoxZona
        If Not IsNull(!IdBodegaFiltro) Then mProps.IdBodegaFiltro = !IdBodegaFiltro                     'JEAA 12/06/2009
        If Not IsNull(!IdPrecioFiltro) Then mProps.IdPrecioFiltro = !IdPrecioFiltro                     'JEAA 12/06/2009
        If Not IsNull(!NombreFuente) Then mProps.NombreFuente = !NombreFuente
        If Not IsNull(!TamanioFuente) Then mProps.TamanioFuente = !TamanioFuente
        If Not IsNull(!AfectaCantidadDocumento) Then mProps.AfectaCantidadDocumento = !AfectaCantidadDocumento
        If Not IsNull(!FormaPagoSRI) Then mProps.FormaPagoSRI = !FormaPagoSRI
        If Not IsNull(!ColVisibleNew) Then mProps.ColVisibleNew = !ColVisibleNew
        If Not IsNull(!ColEditableNew) Then mProps.ColEditableNew = !ColEditableNew
        If Not IsNull(!ColSeleccionableNew) Then mProps.ColSeleccionableNew = !ColSeleccionableNew
        If Not IsNull(!ArchivoReporteRIDE) Then mProps.ArchivoReporteRIDE = !ArchivoReporteRIDE
        If Not IsNull(!IdEmpleadoPre) Then mProps.IdEmpleadoPre = !IdEmpleadoPre
        If Not IsNull(!CodEmpleadoPre) Then mProps.CodEmpleadoPre = !CodEmpleadoPre
        If Not IsNull(!NumCopiasImp) Then mProps.NumCopiasImp = !NumCopiasImp
        If Not IsNull(!AsientoxPCG) Then mProps.AsientoxPCG = !AsientoxPCG
        If Not IsNull(!ConfigBotones) Then mProps.ConfigBotones = !ConfigBotones
        If Not IsNull(!RutaXMLOrigen) Then mProps.RutaXMLOrigen = !RutaXMLOrigen
        If Not IsNull(!RutaXMLDestino) Then mProps.RutaXMLDestino = !RutaXMLDestino
        If Not IsNull(!NumRefSig) Then mProps.NumRefSig = !NumRefSig

        
        
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    'Recupera los Asientos
    sql = "SELECT Id FROM GNTransAsiento WHERE CodTrans='" & CodTrans & "' ORDER BY Orden"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddAsiento
            Asientos(i).Recuperar !Id
            .MoveNext
        Loop
        .Close
    End With
    
    'Recupera los Recargos
    sql = "SELECT Id FROM GNTransRecargo WHERE CodTrans='" & CodTrans & "' ORDER BY Orden"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            If !Id > 0 Then
                i = AddRecargo
                Recargos(i).Recuperar !Id
            End If
            .MoveNext
        Loop
        .Close
    End With
        
    If Len(mProps.objEmpresa.GNOpcion.ObtenerValor("Presupuesto")) > 0 Then
        If mProps.objEmpresa.GNOpcion.ObtenerValor("Presupuesto") = 1 Then
            'Recupera los Asientos Presupuesto
            sql = "SELECT Id FROM GNTransAsientoPresup WHERE CodTrans='" & CodTrans & "' ORDER BY Orden"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            With rs
                Do Until .EOF
                    i = AddAsientoPresup
                    AsientosPresup(i).Recuperar !Id
                    .MoveNext
                Loop
                .Close
            End With
        End If
    End If
        
        
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset

    'Verifica si existe Transacciones
    sql = "SELECT TransID FROM GNComprobante WHERE CodTrans='" & mProps.CodTrans & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "GNTrans.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing

    'Si no es eliminable
    If Not mProps.BandEliminable Then
        Err.Raise ERR_NOELIMINABLE, "GNTrans_Eliminar", MSGERR_NOELIMINABLE
    End If

    'Comienza la transacción
#If DAOLIB Then
    gobjWS.BeginTrans
#Else
    mProps.objEmpresa.Coneccion.BeginTrans
#End If
    On Error GoTo ErrTrap

    'Elimina los Asientos
    sql = "DELETE FROM GNTransAsiento WHERE CodTrans='" & CodTrans & "'"
    mProps.objEmpresa.Execute sql, True
    
    'Elimina los Recargos
    sql = "DELETE FROM GNTransRecargo WHERE CodTrans='" & CodTrans & "'"
    mProps.objEmpresa.Execute sql, True
        
    'Elimina el GNTrans
    sql = "DELETE FROM GNTrans WHERE CodTrans='" & CodTrans & "'"
    mProps.objEmpresa.Execute sql, True

    'Elimina los Asientos Presupuesto
    sql = "DELETE FROM GNTransAsientoPresup WHERE CodTrans='" & CodTrans & "'"
    mProps.objEmpresa.Execute sql, True


    'Concluye la transacción
#If DAOLIB Then
    gobjWS.CommitTrans
#Else
    mProps.objEmpresa.Coneccion.CommitTrans
#End If
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_TRANS", "Transacción Eliminada: " & _
                                                         mProps.CodTrans, "DT"
    
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_TRANS
    Exit Sub

ErrTrap:
    'Cancela la transacción
#If DAOLIB Then
    gobjWS.Rollback
#Else
    mProps.objEmpresa.Coneccion.RollbackTrans
#End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Public Function Clone() As GNTrans
    Dim nuevo As GNTrans, i As Long, na As GNTransAsiento, nr As GNTransRecargo
    
    'Copia las propiedades de GNTrans al nuevo objeto
    Set nuevo = New GNTrans
    nuevo.Props = mProps
    nuevo.BandEliminable = True
    nuevo.NumTransSiguiente = 1     '*** MAKOTO 25/ene/01 Agregado.
    nuevo.NumTransUltimo = 1     '*** jeaa /15/11/2006
    nuevo.NumTransInicio = 1     '*** jeaa /15/11/2006
    nuevo.NumTransUltimoPreimpreso = 0  'jeaa 25/09/2007
    nuevo.NumCopiasImp = 1
                                    
                                    '*** Porque no tiene sentido copiar éste número
    
    'Copia los Asientos al nuevo objeto
    For i = 1 To mcolAsientos.Count
        Set na = Asientos(i).Clone
        nuevo.AddAsientoParaClone na
    Next i
    
    'Copia los Recargos al nuevo objeto
    For i = 1 To mcolRecargos.Count
        Set nr = Recargos(i).Clone
        nuevo.AddRecargoParaClone nr
    Next i
    
    Set Clone = nuevo
    Set nuevo = Nothing
    Set na = Nothing
    Set nr = Nothing
End Function

Friend Sub AddAsientoParaClone(obj As GNTransAsiento)
    Set obj.GNTrans = Me
    mcolAsientos.Add obj
End Sub

Friend Sub AddRecargoParaClone(obj As GNTransRecargo)
    Set obj.GNTrans = Me
    mcolRecargos.Add obj
End Sub



Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property








Public Function AddAsiento() As Long
    Dim nuevo As GNTransAsiento
    
    Set nuevo = New GNTransAsiento
    Set nuevo.GNTrans = Me
    mcolAsientos.Add nuevo
    AddAsiento = mcolAsientos.Count    'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function

Public Sub RemoveAsiento(ix As Variant)
    Dim obj As GNTransAsiento, i As Long

    'Si llama con numero de indice del objeto
    If (VarType(ix) = vbLong) Or (VarType(ix) = vbInteger) Then
        mcolAsientos.Remove ix      'Elimina el objeto
        
    'Si llama con la referencia al objeto que quiere eliminar
    Else
        'Busca el objeto dentro de la colección
        For i = mcolAsientos.Count To 1 Step -1
            Set obj = mcolAsientos.item(i)
            If obj Is ix Then       'Si encuentra
                mcolAsientos.Remove i       'Elimina el objeto
                GoTo eliminado
            End If
        Next i
        'Si no encuentra el objeto, genera error
        Err.Raise ERR_INVALIDO, "GNTrans_RemoveAsiento", "No se puede eliminar el objeto."
    End If
eliminado:
    mbooModificado = True
    Set obj = Nothing
End Sub

Public Property Get CountAsiento() As Long
    CountAsiento = mcolAsientos.Count
End Property

Public Property Get Asientos(ix) As GNTransAsiento
    Set Asientos = mcolAsientos.item(ix)
End Property

Public Function AddRecargo() As Long
    Dim nuevo As GNTransRecargo
    
    Set nuevo = New GNTransRecargo
    Set nuevo.GNTrans = Me
    mcolRecargos.Add nuevo
    AddRecargo = mcolRecargos.Count    'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function

Public Sub RemoveRecargo(ix As Variant)
    Dim obj As GNTransRecargo, i As Long

    'Si llama con numero de indice del objeto
    If (VarType(ix) = vbLong) Or (VarType(ix) = vbInteger) Then
        mcolRecargos.Remove ix      'Elimina el objeto
        
    'Si llama con la referencia al objeto que quiere eliminar
    Else
        'Busca el objeto dentro de la colección
        For i = mcolRecargos.Count To 1 Step -1
            Set obj = mcolRecargos.item(i)
            If obj Is ix Then       'Si encuentra
                mcolRecargos.Remove i       'Elimina el objeto
                GoTo eliminado
            End If
        Next i
        'Si no encuentra el objeto, genera error
        Err.Raise ERR_INVALIDO, "GNTransRecargo_RemoveRecargo", "No se puede eliminar el Recargo."
    End If
eliminado:
    mbooModificado = True
    Set obj = Nothing
End Sub

Public Property Get CountRecargo() As Long
    CountRecargo = mcolRecargos.Count
End Property

Public Property Get Recargos(ix As Long) As GNTransRecargo
    Set Recargos = mcolRecargos.item(ix)
End Property

Public Property Get BuscaRecargo(CodRecargo As String) As GNTransRecargo
    Dim obj As GNTransRecargo
    
    For Each obj In mcolRecargos
        If obj.CodRecargo = CodRecargo Then
            Set BuscaRecargo = obj
            Exit For
        End If
    Next obj
    Set obj = Nothing
End Property


Public Function ListaGNTransAsiento() As Variant
    Dim sql As String, rs As Recordset
    
    sql = "SELECT IdCuenta,CodCuenta,Descripcion,Formula,BandDeudor,Comentario " & _
          "FROM GNTransAsiento " & _
          "WHERE CodTrans='" & mProps.CodTrans & "' " & _
          "ORDER BY Orden"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    Set ListaGNTransAsiento = rs
    Set rs = Nothing
End Function

Public Function ListaGNTransRecargo() As Variant
    Dim sql As String, rs As Recordset
    
    sql = "SELECT IdRecargo,CodRecargo,ValorPre,PorcentajePre,BandModificable,BandIVAItem " & _
          "FROM GNTransRecargo " & _
          "WHERE CodTrans='" & mProps.CodTrans & "' " & _
          "ORDER BY Orden"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    Set ListaGNTransRecargo = rs
    Set rs = Nothing
End Function


Public Function ListaGNTransRecargoParaFlex() As String
    Dim sql As String, s As String
    Dim rs As Recordset
    
    sql = "SELECT CodRecargo,Descripcion FROM IVRecargo ORDER BY CodRecargo"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)

#If DAOLIB Then
    s = "|" & GetStringDAO(rs, 5 + 50)      '*** MAKOTO 25/ene/01 Mod.
#Else
    If Not rs.EOF Then s = "|" & rs.GetString(adClipString, , vbTab, "|")
#End If
    ListaGNTransRecargoParaFlex = s
    
    rs.Close
    Set rs = Nothing
End Function


'======================= Opciones para módulo IV
Public Property Let IVTipoTrans(ByVal value As String)
        
    If value <> "I" And value <> "E" And value <> "T" And value <> "C" And value <> "A" And value <> "P" And value <> "R" And value <> "X" And value <> "Y" And value <> "R" Then
        Err.Raise ERR_INVALIDO, "GNTrans_IVTipoTrans", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion, 1, 1) = value
    mbooModificado = True
End Property

Public Property Get IVTipoTrans() As String
    IVTipoTrans = Mid$(Opcion, 1, 1)
End Property

Public Property Let IVPideProveedor(ByVal value As Boolean)
    Mid$(mProps.Opcion, 2, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVPideProveedor() As Boolean
    IVPideProveedor = (Mid$(Opcion, 2, 1) = "S")
End Property

Public Property Let IVPideCliente(ByVal value As Boolean)
    Mid$(mProps.Opcion, 3, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVPideCliente() As Boolean
    IVPideCliente = (Mid$(Opcion, 3, 1) = "S")
End Property

Public Property Let IVPideNumDoc(ByVal value As Boolean)
    Mid$(mProps.Opcion, 4, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVPideNumDoc() As Boolean
    IVPideNumDoc = (Mid$(Opcion, 4, 1) = "S")
End Property

Public Property Let IVPideVendedor(ByVal value As Boolean)
    Mid$(mProps.Opcion, 5, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVPideVendedor() As Boolean
    IVPideVendedor = (Mid$(Opcion, 5, 1) = "S")
End Property

Public Property Let IVPorPagar(ByVal value As Boolean)
    Mid$(mProps.Opcion, 6, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVPorPagar() As Boolean
    IVPorPagar = (Mid$(Opcion, 6, 1) = "S")
End Property

Public Property Let IVRecargoEnCosto(ByVal value As Boolean)
    Mid$(mProps.Opcion, 7, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVRecargoEnCosto() As Boolean
    IVRecargoEnCosto = (Mid$(Opcion, 7, 1) = "S")
End Property

Public Property Let IVVisibleF6(ByVal value As Boolean)
    Mid$(mProps.Opcion, 8, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVVisibleF6() As Boolean
    IVVisibleF6 = (Mid$(Opcion, 8, 1) = "S")
End Property

Public Property Let IVVisibleF7(ByVal value As Boolean)
    Mid$(mProps.Opcion, 9, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVVisibleF7() As Boolean
    IVVisibleF7 = (Mid$(Opcion, 9, 1) = "S")
End Property

Public Property Let IVVisibleF8(ByVal value As Boolean)
    Mid$(mProps.Opcion, 10, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVVisibleF8() As Boolean
    IVVisibleF8 = (Mid$(Opcion, 10, 1) = "S")
End Property

Public Property Let IVVerificaExist(ByVal value As Boolean)
    Mid$(mProps.Opcion, 11, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property
Public Property Get IVVerificaExist() As Boolean
    IVVerificaExist = (Mid$(Opcion, 11, 1) = "S")
End Property

Public Property Let IVProvCliPorFila(ByVal value As Boolean)
    Mid$(mProps.Opcion, 12, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVProvCliPorFila() As Boolean
    IVProvCliPorFila = (Mid$(Opcion, 12, 1) = "S")
End Property

Public Property Let IVVerificaLimite(ByVal value As Boolean)
    Mid$(mProps.Opcion, 13, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVVerificaLimite() As Boolean
    IVVerificaLimite = (Mid$(Opcion, 13, 1) = "S")
End Property

Public Property Let IVVerificaLimiteCon(ByVal value As String)
    If Len(value) > 5 Then
        Err.Raise ERR_INVALIDO, "GNTrans.IVVerificaLimiteCon", MSGERR_INVALIDO
    End If

    mProps.VerificaLimiteCon = value
    mbooModificado = True
End Property

Public Property Get IVVerificaLimiteCon() As String
    IVVerificaLimiteCon = Trim$(mProps.VerificaLimiteCon)
End Property

Public Property Let IVVerificaLimiteNoGrabar(ByVal value As Boolean)
    Mid$(mProps.Opcion, 19, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVVerificaLimiteNoGrabar() As Boolean
    IVVerificaLimiteNoGrabar = (Mid$(Opcion, 19, 1) = "S")
End Property


Public Property Let IVNumFilaMax(ByVal value As Integer)
    If value > 999 Then
        Err.Raise ERR_INVALIDO, "GNTrans.IVNumFilaMax", MSGERR_INVALIDO
    End If
    
    'Guarda llenando con espacio si tiene menos de 3 caracteres
    '--> SIEMPRE ocupa 3 espacios (Posición 20 - 22)
    Mid$(mProps.Opcion, 20, 3) = Format(value, "000")
    mbooModificado = True
End Property

Public Property Get IVNumFilaMax() As Integer
    IVNumFilaMax = Val(Mid$(Opcion, 20, 3))
End Property


'*** MAKOTO 25/ene/01 Agregado
Public Property Let IVPermitirSignoNegativo(ByVal value As Boolean)
    Mid$(mProps.Opcion, 23, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'*** MAKOTO 25/ene/01 Agregado
Public Property Get IVPermitirSignoNegativo() As Boolean
    IVPermitirSignoNegativo = (Mid$(Opcion, 23, 1) = "S")
End Property

'*** MAKOTO 07/feb/01 Agregado
Public Property Let IVNoPermitirCostoCero(ByVal value As Boolean)
    Mid$(mProps.Opcion, 24, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'*** MAKOTO 25/ene/01 Agregado
Public Property Get IVNoPermitirCostoCero() As Boolean
    IVNoPermitirCostoCero = (Mid$(Opcion, 24, 1) = "S")
End Property
'*** DIEGO 28/may/01 Agregado
Public Property Let IVNoVerificaTotalCuadrado(ByVal value As Boolean)
    Mid$(mProps.Opcion, 25, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property


Public Property Get IVNoVerificaTotalCuadrado() As Boolean
    IVNoVerificaTotalCuadrado = (Mid$(Opcion, 25, 1) = "S")
End Property


'*** DIEGO 06/jun/01 Agregado
Public Property Let IVVerificaItemsFuente(ByVal value As Boolean)
    Mid$(mProps.Opcion, 26, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property


Public Property Get IVVerificaItemsFuente() As Boolean
    IVVerificaItemsFuente = (Mid$(Opcion, 26, 1) = "S")
End Property

'*** Oliver 07/feb/2002
'*** Agregado
Public Property Let IVPrecioPre(ByVal value As Integer)
    If value > 7 Or value < 0 Then 'GRV Precio 7 - 16MAY2014
        Err.Raise ERR_INVALIDO, "GNTrans.IVPrecioPre", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion, 27, 1) = value
    mbooModificado = True
End Property

Public Property Get IVPrecioPre() As Integer
    IVPrecioPre = Val(Mid$(Opcion, 27, 1))
'    If IVPrecioPre = 0 Then IVPrecioPre = 1   'cuando utilize y no este
                                              'configurado debe ser  1 por defecto
End Property

'*** Agregado Alex Julio/2002
Public Property Let IVNumDocsPendientes(ByVal value As Integer)
    If value > 99 Or value < 0 Then
        Err.Raise ERR_INVALIDO, "GNTrans.IVNumDocsPendientes", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion, 28, 2) = Format(value, "00")
    mbooModificado = True
End Property

Public Property Get IVNumDocsPendientes() As Integer
    IVNumDocsPendientes = Val(Mid$(Opcion, 28, 2))
End Property

Public Property Let IVNOPermitirDocsPend(ByVal value As Boolean)
    Mid$(mProps.Opcion, 30, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVNOPermitirDocsPend() As Boolean
    IVNOPermitirDocsPend = (Mid$(Opcion, 30, 1) = "S")
End Property

'Agregado Alex Sept/2002
Public Property Get IVVisibleAnexos() As Boolean
    IVVisibleAnexos = (Mid$(Opcion, 31, 1) = "S")
End Property

Public Property Let IVVisibleAnexos(ByVal value As Boolean)
    Mid$(mProps.Opcion, 31, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property
'Agregado Alex Marzo/2003
Public Property Get IVPrecioxCliente() As Boolean
    IVPrecioxCliente = (Mid$(Opcion, 32, 1) = "S")
End Property
Public Property Let IVPrecioxCliente(ByVal value As Boolean)
    Mid$(mProps.Opcion, 32, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVGrupoClientesParaPrecio() As Integer
    IVGrupoClientesParaPrecio = Val(Mid$(Opcion, 33, 1))
End Property

Public Property Let IVGrupoClientesParaPrecio(ByVal value As Integer)
    If value > 9 Or value < 0 Then      '0: ningún grupo; PCGrupo 1, 2, 3 ...
        Err.Raise ERR_INVALIDO, "GNTrans.IVGrupoClientesParaPrecio", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion, 33, 1) = value
    mbooModificado = True
End Property

'*** Angel 20/mar/03
Public Property Let IVAlertarExisMaxMin(ByVal value As Boolean)
    Mid$(mProps.Opcion, 34, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property
'*** Angel 20/mar/03
Public Property Get IVAlertarExisMaxMin() As Boolean
    IVAlertarExisMaxMin = (Mid$(Opcion, 34, 1) = "S")
End Property

'*** Oliver 15/07/2003
Public Property Let IVControlPrecio0(ByVal value As Boolean)
    Mid$(mProps.Opcion, 35, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'*** Oliver 15/07/2003
Public Property Get IVControlPrecio0() As Boolean
    IVControlPrecio0 = (Mid$(Opcion, 35, 1) = "S")
End Property

'*** Diego 01/08/2003
Public Property Let IVVerificaLimiteCredito(ByVal value As Boolean)
    Mid$(mProps.Opcion, 36, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'*** Diego 01/08/2003
Public Property Get IVVerificaLimiteCredito() As Boolean
    IVVerificaLimiteCredito = (Mid$(Opcion, 36, 1) = "S")
End Property


'*** Diego 01/08/2003
Public Property Let IVNumDiasVencidos(ByVal value As Integer)
    If value > 99 Or value < 0 Then
        Err.Raise ERR_INVALIDO, "GNTrans.IVNumDiasVencidos", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion, 37, 2) = Format(value, "00")
    mbooModificado = True
End Property

Public Property Get IVNumDiasVencidos() As Integer
    IVNumDiasVencidos = Val(Mid$(Opcion, 37, 2))
End Property

Public Property Let IVVerificaDiasVencidos(ByVal value As Boolean)
    Mid$(mProps.Opcion, 39, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'*** Diego 01/08/2003
Public Property Get IVVerificaDiasVencidos() As Boolean
    IVVerificaDiasVencidos = (Mid$(Opcion, 39, 1) = "S")
End Property

Public Property Get IVPendientesAccion() As Integer
    IVPendientesAccion = Val(Mid$(mProps.Opcion, 40, 1))
End Property

Public Property Let IVPendientesAccion(ByVal value As Integer)
    If value > 9 Or value < 0 Then      '0: no controla 1 mensaje de advertencia, 2 no permite grabar
        Err.Raise ERR_INVALIDO, "GNTrans.IVGrupoClientesParaPrecio", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion, 40, 1) = value
    mbooModificado = True
End Property


Public Property Let IVSoloDocsVencidos(ByVal value As Boolean)
    Mid$(mProps.Opcion, 41, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVSoloDocsVencidos() As Boolean
    IVSoloDocsVencidos = (Mid$(Opcion, 41, 1) = "S")
End Property

'***Diego 10/09/2003
Public Property Let IVMostrarVuelto(ByVal value As Boolean)
    Mid$(mProps.Opcion, 42, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVMostrarVuelto() As Boolean
    IVMostrarVuelto = (Mid$(Opcion, 42, 1) = "S")
End Property

Public Property Let IVCobroContado(ByVal value As String)
    If Len(value) > 5 Then
        Err.Raise ERR_INVALIDO, "GNTrans.IVCobroContado", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion, 43, 5) = String(5, " ")
    Mid$(mProps.Opcion, 43, 5) = value
    mbooModificado = True
End Property

Public Property Get IVCobroContado() As String
    IVCobroContado = Trim$(Mid$(mProps.Opcion, 43, 5))
End Property

Public Property Let IVVerificaCobroConsFinal(ByVal value As Boolean)
    Mid$(mProps.Opcion, 48, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVVerificaCobroConsFinal() As Boolean
    IVVerificaCobroConsFinal = (Mid$(Opcion, 48, 1) = "S")
End Property

'***Diego 20/10/2003 Descuento  por PCGrupo
Public Property Let IVDescXPCGrupo(ByVal value As Boolean)
    Mid$(mProps.Opcion, 49, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVDescXPCGrupo() As Boolean
    IVDescXPCGrupo = (Mid$(Opcion, 49, 1) = "S")
End Property

'***Angel 22/oct/2003 Datos adicionales de Proveedor/Cliente
Public Property Let IVVisualizaDatosPC(ByVal value As Boolean)
    Mid$(mProps.Opcion, 50, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVVisualizaDatosPC() As Boolean
    IVVisualizaDatosPC = (Mid$(Opcion, 50, 1) = "S")
End Property

'***Esteban 27/mayo/2004 si muestra existencia de todas las bodegas
Public Property Let IVVisualizaTodaBodega(ByVal value As Boolean)
    Mid$(mProps.Opcion, 51, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVVisualizaTodaBodega() As Boolean
    IVVisualizaTodaBodega = (Mid$(Opcion, 51, 1) = "S")
End Property

Public Property Let IVVisualizaExistencia0(ByVal value As Boolean)
    Mid$(mProps.Opcion, 52, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVVisualizaExistencia0() As Boolean
    IVVisualizaExistencia0 = (Mid$(Opcion, 52, 1) = "S")
End Property

'***Esteban 07/septiembre/2004 avance de linea en IVPVT
Public Property Let IVAvanceLinea(ByVal value As Boolean)
    Mid$(mProps.Opcion, 53, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVAvanceLinea() As Boolean
    IVAvanceLinea = (Mid$(Opcion, 53, 1) = "S")
End Property

'***Esteban 07/12/2004 PERMITE controlar documentos pendientes en por cobrar/pagar
Public Property Let ControlarPCenDocs(ByVal value As Boolean)
    Mid$(mProps.Opcion, 54, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get ControlarPCenDocs() As Boolean
    ControlarPCenDocs = (Mid$(Opcion, 54, 1) = "S")
End Property
'jeaa 17/01/05 para controlar limites en devoluciones
Public Property Get IVControlaLimite() As Boolean
    IVControlaLimite = (Mid$(Opcion, 55, 1) = "S")
End Property

Public Property Let IVControlaLimite(ByVal value As Boolean)
    Mid$(mProps.Opcion, 55, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'jeaa 20/01/05 para controlar limites en devoluciones
Public Property Get IVItemsporProveedor() As Boolean
    IVItemsporProveedor = (Mid$(Opcion, 56, 1) = "S")
End Property

Public Property Let IVItemsporProveedor(ByVal value As Boolean)
    Mid$(mProps.Opcion, 56, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property
'jeaa 20/01/05 para guardar motivo
Public Property Let IVPideMotivo(ByVal value As Boolean)
    Mid$(mProps.Opcion, 57, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVPideMotivo() As Boolean
    IVPideMotivo = (Mid$(Opcion, 57, 1) = "S")
End Property

'jeaa 09/05/05 para verificacion de fechas
Public Property Let VerificaFechas(ByVal value As Boolean)
    Mid$(mProps.Opcion, 58, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get VerificaFechas() As Boolean
    VerificaFechas = (Mid$(Opcion, 58, 1) = "S")
End Property

Public Property Let IVValNumDoc(ByVal value As Boolean)
    Mid$(mProps.Opcion, 59, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVValNumDoc() As Boolean
    IVValNumDoc = (Mid$(Opcion, 59, 1) = "S")
End Property

Public Property Let IVCambioTrans(ByVal value As Boolean)
    Mid$(mProps.Opcion, 73, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVCambioTrans() As Boolean
    IVCambioTrans = (Mid$(Opcion, 73, 1) = "S")
End Property


Public Property Let IVValTipoDoc(ByVal value As Boolean)
    Mid$(mProps.Opcion, 74, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVValTipoDoc() As Boolean
    IVValTipoDoc = (Mid$(Opcion, 74, 1) = "S")
End Property

Public Property Let IVTipoDocTrans(ByVal value As String)
    Mid$(mProps.Opcion, 75, 1) = value
    mbooModificado = True
End Property

Public Property Get IVTipoDocTrans() As String
    IVTipoDocTrans = Mid$(Opcion, 75, 1)
End Property

Public Property Let IVOmitirPorcentajeDesct(ByVal value As Boolean)
    Mid$(mProps.Opcion, 76, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVOmitirPorcentajeDesct() As Boolean
    IVOmitirPorcentajeDesct = (Mid$(Opcion, 76, 1) = "S")
End Property

'jeaa 12/10/2005
Public Property Let IVTransRTP(ByVal value As String)
    If Len(value) > 5 Then Err.Raise ERR_INVALIDO, "GNTrans.IVTransRTP", MSGERR_INVALIDO
    Mid$(mProps.Opcion, 77, 5) = String(5, " ")
    Mid$(mProps.Opcion, 77, 5) = value
    mbooModificado = True
End Property

Public Property Get IVTransRTP() As String
    IVTransRTP = Trim(Mid$(Opcion, 77, 5))
End Property

'AUC 19/10/2005
Public Property Let IVControlaCreditos(ByVal value As Boolean)
 Mid$(mProps.Opcion, 82, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVControlaCreditos() As Boolean
    IVControlaCreditos = (Mid$(Opcion, 82, 1) = "S")
 End Property
'jeaa 15/11/2005
Public Property Let IVItemDevolucion(ByVal value As Boolean)
    Mid$(mProps.Opcion, 83, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property
Public Property Get IVItemDevolucion() As Boolean
    IVItemDevolucion = (Mid$(Opcion, 83, 1) = "S")
End Property

'jeaa 15/11/2005
Public Property Let IVControlaTotalDev(ByVal value As Boolean)
    Mid$(mProps.Opcion, 84, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVControlaTotalDev() As Boolean
    IVControlaTotalDev = (Mid$(Opcion, 84, 1) = "S")
End Property

Public Property Let ControlaCCxCli(ByVal value As Boolean) 'AUC 17/17/2005
        Mid$(mProps.Opcion, 85, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property
Public Property Get ControlaCCxCli() As Boolean 'AUC 17/17/2005
    ControlaCCxCli = (Mid$(Opcion, 85, 1) = "S")
End Property

Public Property Let RequierePasswordVen(ByVal value As Boolean) 'AUC 17/17/2005
        Mid$(mProps.Opcion, 86, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property
Public Property Get RequierePasswordVen() As Boolean 'AUC 17/17/2005
    RequierePasswordVen = (Mid$(Opcion, 86, 1) = "S")
End Property

'------ Auc 14/12/05
Public Property Get TotalizaItemsRepetidos() As Boolean
    TotalizaItemsRepetidos = (Mid$(Opcion, 87, 1) = "S")
End Property
Public Property Let TotalizaItemsRepetidos(ByVal value As Boolean)
    Mid$(mProps.Opcion, 87, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property
'jeaa 21/12/2005
Public Property Let IVRetencionAutomatica(ByVal value As Boolean)
        Mid$(mProps.Opcion, 88, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property
Public Property Get IVRetencionAutomatica() As Boolean 'jeaa 21/12/2005
    IVRetencionAutomatica = (Mid$(Opcion, 88, 1) = "S")
End Property

'jeaa 27/01/2006
Public Property Let IVActualizaItemProv(ByVal value As Boolean)
        Mid$(mProps.Opcion, 89, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property
Public Property Get IVActualizaItemProv() As Boolean 'jeaa 21/12/2005
    IVActualizaItemProv = (Mid$(Opcion, 89, 1) = "S")
End Property

'03/06 agregado para la casa del disfraz ver si el num 89 no existe aun
Public Property Let MostrarSubItem(ByVal value As Boolean)
    Mid$(mProps.Opcion, 90, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get MostrarSubItem() As Boolean
    MostrarSubItem = (Mid$(Opcion, 90, 1) = "S")
End Property
'AUC agregado

Public Property Let FechaDevolucion(ByVal value As Boolean)
    mProps.FechaDevolucion = value
    mbooModificado = True
End Property

Public Property Get FechaDevolucion() As Boolean
    FechaDevolucion = mProps.FechaDevolucion
End Property

Public Property Let ivDevolxSubitems(ByVal value As Boolean)
    Mid$(mProps.Opcion, 91, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get ivDevolxSubitems() As Boolean
    ivDevolxSubitems = (Mid$(Opcion, 91, 1) = "S")
End Property

Public Property Let ivItemMenor(ByVal value As Boolean)
    Mid$(mProps.Opcion, 92, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get ivItemMenor() As Boolean
    ivItemMenor = (Mid$(Opcion, 92, 1) = "S")
End Property
'AUC 23/03/06
Public Property Let GenerarKardexSubItem(ByVal value As Boolean)
'    Mid$(mProps.Opcion, 93, 1) = IIf(value, "S", "N")
'   mbooModificado = True
End Property

Public Property Get GenerarKardexSubItem() As Boolean
'    GenerarKardexSubItem = (Mid$(Opcion, 93, 1) = "S")
End Property

Public Property Let MostrarSubItemsImp(ByVal value As Boolean)
    Mid$(mProps.Opcion, 93, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property
Public Property Get MostrarSubItemsImp() As Boolean
    MostrarSubItemsImp = (Mid$(Opcion, 93, 1) = "S")
End Property


'jeaa 17/05/2006
Public Property Let IVControlaLimitePorFactura(ByVal value As Boolean)
        Mid$(mProps.Opcion, 94, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property
Public Property Get IVControlaLimitePorFactura() As Boolean
    IVControlaLimitePorFactura = (Mid$(Opcion, 94, 1) = "S")
End Property

'AUC para controlar que el num documento referencia sea igual al numero de transaccion
Public Property Let NumRefIgualNumTrans(ByVal value As Boolean)
        Mid$(mProps.Opcion, 95, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property

Public Property Get NumRefIgualNumTrans() As Boolean
    NumRefIgualNumTrans = (Mid$(Opcion, 95, 1) = "S")
End Property


'----------
'AUC 06/04/06
Public Property Let IvPideMoneda(ByVal value As Boolean)
    Mid$(mProps.Opcion, 96, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IvPideMoneda() As Boolean
    IvPideMoneda = (Mid$(Opcion, 96, 1) = "S")
End Property
'AUC 10/04/06
Public Property Let HistorialCliente(ByVal value As Boolean)
    Mid$(mProps.Opcion, 97, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property
Public Property Get HistorialCliente() As Boolean
    HistorialCliente = (Mid$(Opcion, 97, 1) = "S")
End Property

'AUC 10/04/06 para controlar por importancia de items
Public Property Let ItemsImportantes(ByVal value As Boolean)
    Mid$(mProps.Opcion, 98, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property
Public Property Get ItemsImportantes() As Boolean
    ItemsImportantes = (Mid$(Opcion, 98, 1) = "S")
End Property

'jeaa 21/09/2006
Public Property Let IVGeneraDesperdicio(ByVal value As Boolean)
    Mid$(mProps.Opcion, 99, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property
Public Property Get IVGeneraDesperdicio() As Boolean
    IVGeneraDesperdicio = (Mid$(Opcion, 99, 1) = "S")
End Property

'AUC 05/10/06 para Ubicacion del cursor en pantalla busqueda
Public Property Let UbicaCursor(ByVal value As String)
    If value <> "C" And value <> "D" And value <> "S" And value <> "A" Then
        Err.Raise ERR_INVALIDO, "GNTrans_UbicaCursor", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion, 100, 1) = value
    mbooModificado = True
End Property

Public Property Get UbicaCursor() As String
    UbicaCursor = Mid$(Opcion, 100, 1)
End Property

'jeaa 20/10/2006
Public Property Let IVPantallaComprometido(ByVal value As Boolean)
        Mid$(mProps.Opcion, 101, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property
Public Property Get IVPantallaComprometido() As Boolean
    IVPantallaComprometido = (Mid$(Opcion, 101, 1) = "S")
End Property

Public Property Let IVAfectaComprometido(ByVal value As Boolean)
        Mid$(mProps.Opcion, 102, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property
Public Property Get IVAfectaComprometido() As Boolean
    IVAfectaComprometido = (Mid$(Opcion, 102, 1) = "S")
End Property

Public Property Let RecuperaRECCC(ByVal value As Boolean)
        Mid$(mProps.Opcion, 103, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property
Public Property Get RecuperaRECCC() As Boolean
    RecuperaRECCC = (Mid$(Opcion, 103, 1) = "S")
End Property
'control sola para Hormimipios 31/05/2007
Public Property Let IVControlaSoloCantidad(ByVal value As Boolean)
        Mid$(mProps.Opcion, 104, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property
Public Property Get IVControlaSoloCantidad() As Boolean
    IVControlaSoloCantidad = Mid$(Opcion, 104, 1) = "S"
    End Property
'control sola para Hormimipios 31/05/2007
Public Property Let IVControlaSoloFamilia(ByVal value As Boolean)
        Mid$(mProps.Opcion, 105, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property
Public Property Get IVControlaSoloFamilia() As Boolean
    IVControlaSoloFamilia = Mid$(Opcion, 105, 1) = "S"
End Property

'control sola para Hormimipios 11/06/2007
Public Property Let ControlCantidadxFamilia(ByVal value As Boolean)
        Mid$(mProps.Opcion, 106, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property
Public Property Get ControlCantidadxFamilia() As Boolean
    ControlCantidadxFamilia = Mid$(Opcion, 106, 1) = "S"
End Property

'control sola para Hormimipios 11/06/2007
Public Property Let ControlCantidadxGrupo(ByVal value As Boolean)
        Mid$(mProps.Opcion, 107, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property
Public Property Get ControlCantidadxGrupo() As Boolean
    ControlCantidadxGrupo = Mid$(Opcion, 107, 1) = "S"
End Property

'control sola para Hormimipios 11/06/2007
Public Property Let GrupoControlaMEMO(ByVal value As String)
        Mid$(mProps.Opcion, 108, 1) = value
        mbooModificado = True
End Property
Public Property Get GrupoControlaMEMO() As String
    GrupoControlaMEMO = Mid$(Opcion, 108, 1)
End Property

'control sola para Hormimipios 11/06/2007
Public Property Let ListaTranVerificaLimite(ByVal value As Boolean)
        Mid$(mProps.Opcion, 109, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property
Public Property Get ListaTranVerificaLimite() As Boolean
    ListaTranVerificaLimite = Mid$(Opcion, 109, 1) = "S"
End Property

Public Property Let IVActualizaPrecios(ByVal value As Boolean)
        Mid$(mProps.Opcion, 110, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property

Public Property Get IVActualizaPrecios() As Boolean
    IVActualizaPrecios = Mid$(Opcion, 110, 1) = "S"
End Property

Public Property Let IVactualizaTRansFuente(ByVal value As Boolean)
        Mid$(mProps.Opcion, 111, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property
Public Property Get IVactualizaTRansFuente() As Boolean
    IVactualizaTRansFuente = Mid$(Opcion, 111, 1) = "S"
End Property

Public Property Let IvCambioColorGris(ByVal value As Boolean)
        Mid$(mProps.Opcion, 112, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property

Public Property Get IvCambioColorGris() As Boolean
    IvCambioColorGris = Mid$(Opcion, 112, 1) = "S"
End Property

Public Property Let IvFiltrarFormaPendientes(ByVal value As Boolean)
        Mid$(mProps.Opcion, 113, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property

Public Property Get IvFiltrarFormaPendientes() As Boolean
    IvFiltrarFormaPendientes = Mid$(Opcion, 113, 1) = "S"
End Property




'======================= Opciones para módulo PC
'*** DIEGO 10/abr/01 Modificado

Public Property Let PCPorCobrar(ByVal value As String)
    If value <> "C" And value <> "P" And value <> "T" Then
        Err.Raise ERR_INVALIDO, "GNTrans_PCPorCobrar", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion, 1, 1) = value
    mbooModificado = True
End Property

Public Property Get PCPorCobrar() As String
    PCPorCobrar = Mid$(Opcion, 1, 1)
End Property

Public Property Let PCModoProveedor(ByVal value As Boolean)
    Mid$(mProps.Opcion, 2, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get PCModoProveedor() As Boolean
    PCModoProveedor = (Mid$(Opcion, 2, 1) = "S")
End Property



'======================= Opciones para módulo TS
Public Property Let TSDebeVisible(ByVal value As Boolean)
    Mid$(mProps.Opcion, 1, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSDebeVisible() As Boolean
    TSDebeVisible = (Mid$(Opcion, 1, 1) = "S")
End Property

Public Property Let TSHaberVisible(ByVal value As Boolean)
    Mid$(mProps.Opcion, 2, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSHaberVisible() As Boolean
    TSHaberVisible = (Mid$(Opcion, 2, 1) = "S")
End Property

Public Property Let TSFechaVisible(ByVal value As Boolean)
    Mid$(mProps.Opcion, 3, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSFechaVisible() As Boolean
    TSFechaVisible = (Mid$(Opcion, 3, 1) = "S")
End Property

Public Property Let TSBancoVisible(ByVal value As Boolean)
    Mid$(mProps.Opcion, 4, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSBancoVisible() As Boolean
    TSBancoVisible = (Mid$(Opcion, 4, 1) = "S")
End Property


Public Property Let TSCobro(ByVal value As Boolean)
    Mid$(mProps.Opcion, 5, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSCobro() As Boolean
    TSCobro = (Mid$(Opcion, 5, 1) = "S")
End Property

'============== Agregado Oliver para verificar que el numero de documento no se repita

Public Property Let TSVerificaDocRepetido(ByVal value As Boolean)
    Mid$(mProps.Opcion, 6, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSVerificaDocRepetido() As Boolean
    TSVerificaDocRepetido = (Mid$(Opcion, 6, 1) = "S")
End Property

Public Property Let TSVerificaTotalCuadrado(ByVal value As Boolean)
    Mid$(mProps.Opcion, 7, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property


Public Property Get TSVerificaTotalCuadrado() As Boolean
    TSVerificaTotalCuadrado = (Mid$(Opcion, 7, 1) = "S")
End Property

'============== Agregado Esteban para verificar que hojas se quieren visualizar

Public Property Let TSVisibleF5(ByVal value As Boolean)
    Mid$(mProps.Opcion, 8, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSVisibleF5() As Boolean
    TSVisibleF5 = (Mid$(Opcion, 8, 1) = "S")
End Property

Public Property Let TSVisibleF6(ByVal value As Boolean)
    Mid$(mProps.Opcion, 9, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSVisibleF6() As Boolean
    TSVisibleF6 = (Mid$(Opcion, 9, 1) = "S")
End Property

Public Property Let TSVisibleF7(ByVal value As Boolean)
    Mid$(mProps.Opcion, 10, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSVisibleF7() As Boolean
    TSVisibleF7 = (Mid$(Opcion, 10, 1) = "S")
End Property

Public Property Let TSVisibleF8(ByVal value As Boolean)
    Mid$(mProps.Opcion, 11, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSVisibleF8() As Boolean
    TSVisibleF8 = (Mid$(Opcion, 11, 1) = "S")
End Property

Public Property Let TSVisibleAsiento(ByVal value As Boolean)
    Mid$(mProps.Opcion, 12, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSVisibleAsiento() As Boolean
    TSVisibleAsiento = (Mid$(Opcion, 12, 1) = "S")
End Property

Public Property Let TSBancoPred(ByVal value As String)
    If Len(value) > 20 Then Err.Raise ERR_INVALIDO, "GNTrans.TSBancoPred", MSGERR_INVALIDO
    Mid$(mProps.Opcion, 13, 20) = String(20, " ")
    Mid$(mProps.Opcion, 13, 20) = value
    mbooModificado = True
End Property

Public Property Get TSBancoPred() As String
    TSBancoPred = Trim(Mid$(Opcion, 13, 20))
End Property

Public Property Let TSDocBancarioPred(ByVal value As String)
    If Len(value) > 5 Then Err.Raise ERR_INVALIDO, "GNTrans.TSDocBancarioPred", MSGERR_INVALIDO
    Mid$(mProps.Opcion, 33, 5) = String(5, " ")
    Mid$(mProps.Opcion, 33, 5) = value
    mbooModificado = True
End Property

Public Property Get TSDocBancarioPred() As String
    TSDocBancarioPred = Trim(Mid$(Opcion, 33, 5))
End Property

Public Property Let TSProvCliPorFila(ByVal value As Boolean)
    Mid$(mProps.Opcion, 38, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSProvCliPorFila() As Boolean
    TSProvCliPorFila = (Mid$(Opcion, 38, 1) = "S")
End Property
'jeaa 25/04/2005
Public Property Let TSPideCobrador(ByVal value As Boolean)
    Mid$(mProps.Opcion, 39, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSPideCobrador() As Boolean
    TSPideCobrador = (Mid$(Opcion, 39, 1) = "S")
End Property

'Agregado Alex Sept/2002
Public Property Get TSVisibleAnexos() As Boolean
    TSVisibleAnexos = (Mid$(Opcion, 40, 1) = "S")
End Property

Public Property Let TSVisibleAnexos(ByVal value As Boolean)
    Mid$(mProps.Opcion, 40, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property
'jeaa 27/05/2007
Public Property Get TSCobradorPorFila() As Boolean
    TSCobradorPorFila = (Mid$(Opcion, 41, 1) = "S")
End Property

Public Property Let TSCobradorPorFila(ByVal value As Boolean)
    Mid$(mProps.Opcion, 41, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'jeaa 30/06/2007
Public Property Get TSDatosPrestamo() As Boolean
    TSDatosPrestamo = (Mid$(Opcion, 42, 1) = "S")
End Property

Public Property Let TSDatosPrestamo(ByVal value As Boolean)
    Mid$(mProps.Opcion, 42, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSDatosVacaciones() As Boolean
    TSDatosVacaciones = (Mid$(Opcion, 43, 1) = "S")
End Property

Public Property Let TSDatosVacaciones(ByVal value As Boolean)
    Mid$(mProps.Opcion, 43, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSGeneraArchivo() As Boolean
    TSGeneraArchivo = (Mid$(Opcion, 44, 1) = "S")
End Property

Public Property Let TSGeneraArchivo(ByVal value As Boolean)
    Mid$(mProps.Opcion, 44, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSGrupoVacaciones() As String
    TSGrupoVacaciones = Trim((Mid$(Opcion, 45, 10)))
End Property

Public Property Let TSGrupoVacaciones(ByVal value As String)
If Len(value) > 20 Then Err.Raise ERR_INVALIDO, "GNTrans.TSGrupoVacaciones", MSGERR_INVALIDO
    Mid$(mProps.Opcion, 45, 10) = String(10, " ")
    Mid$(mProps.Opcion, 45, 10) = value
    mbooModificado = True
End Property

Public Property Get TSIntervaloDias() As String
    TSIntervaloDias = Trim((Mid$(Opcion, 55, 3)))
End Property

Public Property Let TSIntervaloDias(ByVal value As String)
If Len(value) > 3 Then Err.Raise ERR_INVALIDO, "GNTrans.TSIntervaloDias", MSGERR_INVALIDO
    Mid$(mProps.Opcion, 55, 3) = String(10, " ")
    Mid$(mProps.Opcion, 55, 3) = value
    mbooModificado = True
End Property

Public Property Get TSExcluirVacaciones() As Boolean
    TSExcluirVacaciones = (Mid$(Opcion, 58, 1) = "S")
End Property

Public Property Let TSExcluirVacaciones(ByVal value As Boolean)
    Mid$(mProps.Opcion, 58, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Let TSVerificaLimiteCredito(ByVal value As Boolean)
    Mid$(mProps.Opcion, 59, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSVerificaLimiteCredito() As Boolean
    TSVerificaLimiteCredito = (Mid$(Opcion, 59, 1) = "S")
End Property

Public Property Get TSPendientesAccion() As Integer
    TSPendientesAccion = Val(Mid$(mProps.Opcion, 73, 1))
End Property

Public Property Let TSPendientesAccion(ByVal value As Integer)
    If value > 9 Or value < 0 Then      '0: no controla 1 mensaje de advertencia, 2 no permite grabar
        Err.Raise ERR_INVALIDO, "GNTrans.TSGrupoClientesParaPrecio", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion, 73, 1) = value
    mbooModificado = True
End Property

Public Property Let TSCobroAuto(ByVal value As Boolean)
    Mid$(mProps.Opcion, 74, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSCobroAuto() As Boolean
    TSCobroAuto = (Mid$(Opcion, 74, 1) = "S")
End Property

Public Property Let TSRetCobroPAgoAuto(ByVal value As Boolean)
    Mid$(mProps.Opcion, 75, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSRetCobroPAgoAuto() As Boolean
    TSRetCobroPAgoAuto = (Mid$(Opcion, 75, 1) = "S")
End Property

Public Property Let TSCobrosOrdenCrono(ByVal value As Boolean)
    Mid$(mProps.Opcion, 76, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSCobrosOrdenCrono() As Boolean
    TSCobrosOrdenCrono = (Mid$(Opcion, 76, 1) = "S")
End Property

Public Property Let TSGeneraTransIntMora(ByVal value As Boolean)
    Mid$(mProps.Opcion, 77, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSGeneraTransIntMora() As Boolean
    TSGeneraTransIntMora = (Mid$(Opcion, 77, 1) = "S")
End Property

'jeaa 26/03/2009
Public Property Let TSTransAutoIntMora(ByVal value As String)
    If Len(value) > 5 Then Err.Raise ERR_INVALIDO, "GNTrans.TSTransAutoIntMora", MSGERR_INVALIDO
    Mid$(mProps.Opcion, 78, 5) = String(5, " ")
    Mid$(mProps.Opcion, 78, 5) = value
    mbooModificado = True
End Property

Public Property Get TSTransAutoIntMora() As String
    TSTransAutoIntMora = Trim(Mid$(Opcion, 78, 5))
End Property

Public Property Let TSItemAutoIntMora(ByVal value As String)
    If Len(value) > 5 Then Err.Raise ERR_INVALIDO, "GNTrans.TSTransAutoIntMora", MSGERR_INVALIDO
    Mid$(mProps.Opcion, 83, 5) = String(5, " ")
    Mid$(mProps.Opcion, 83, 5) = value
    mbooModificado = True
End Property

Public Property Get TSItemAutoIntMora() As String
    TSItemAutoIntMora = Trim(Mid$(Opcion, 83, 5))
End Property

'jeaa 20/04/2009
Public Property Let TSControlPorSucursal(ByVal value As Boolean)
    Mid$(mProps.Opcion, 88, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSControlPorSucursal() As Boolean
    TSControlPorSucursal = (Mid$(Opcion, 88, 1) = "S")
End Property

Public Property Let TSDatosAdicionales(ByVal value As Boolean)
    Mid$(mProps.Opcion, 89, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSDatosAdicionales() As Boolean
    TSDatosAdicionales = (Mid$(Opcion, 89, 1) = "S")
End Property

Public Property Let TSGeneraTransNDAnti(ByVal value As Boolean)
    Mid$(mProps.Opcion, 90, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSGeneraTransNDAnti() As Boolean
    TSGeneraTransNDAnti = (Mid$(Opcion, 90, 1) = "S")
End Property

Public Property Let TSCobrosporCuota(ByVal value As Boolean)
    Mid$(mProps.Opcion, 91, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSCobrosporCuota() As Boolean
    TSCobrosporCuota = (Mid$(Opcion, 91, 1) = "S")
End Property

Public Property Let TSPagoAuto(ByVal value As Boolean)
    Mid$(mProps.Opcion, 92, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSPagoAuto() As Boolean
    TSPagoAuto = (Mid$(Opcion, 92, 1) = "S")
End Property

Public Property Let TSDatosRetencion(ByVal value As Boolean)
    Mid$(mProps.Opcion, 93, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSDatosRetencion() As Boolean
    TSDatosRetencion = (Mid$(Opcion, 93, 1) = "S")
End Property

Public Property Let TSDatosAdicionalesCHR(ByVal value As Boolean)
    Mid$(mProps.Opcion, 94, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSDatosAdicionalesCHR() As Boolean
    TSDatosAdicionalesCHR = (Mid$(Opcion, 94, 1) = "S")
End Property

Public Property Let TSDatosRetencionxFila(ByVal value As Boolean)
    Mid$(mProps.Opcion, 95, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSDatosRetencionxFila() As Boolean
    TSDatosRetencionxFila = (Mid$(Opcion, 95, 1) = "S")
End Property

Public Property Let TSVisibleSucursalSolicita(ByVal value As Boolean)
    Mid$(mProps.Opcion, 96, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSVisibleSucursalSolicita() As Boolean
    TSVisibleSucursalSolicita = (Mid$(Opcion, 96, 1) = "S")
End Property

Public Property Let TSVerificaNuevoPlazo(ByVal value As Boolean)
    Mid$(mProps.Opcion, 97, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSVerificaNuevoPlazo() As Boolean
    TSVerificaNuevoPlazo = (Mid$(Opcion, 97, 1) = "S")
End Property

Public Property Let TSClaveNuevoPlazo(ByVal value As Boolean)
    Mid$(mProps.Opcion, 98, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSClaveNuevoPlazo() As Boolean
    TSClaveNuevoPlazo = (Mid$(Opcion, 98, 1) = "S")
End Property

Public Property Let TSVisualizaPagoCHP(ByVal value As Boolean)
    Mid$(mProps.Opcion, 99, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSVisualizaPagoCHP() As Boolean
    TSVisualizaPagoCHP = (Mid$(Opcion, 99, 1) = "S")
End Property

Public Property Let TSVisializaTransFuente(ByVal value As Boolean)
    Mid$(mProps.Opcion, 100, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSVisializaTransFuente() As Boolean
    TSVisializaTransFuente = (Mid$(Opcion, 100, 1) = "S")
End Property


Public Property Let TSNumDocsPendientes(ByVal value As Integer)
    If value > 99 Or value < 0 Then
        Err.Raise ERR_INVALIDO, "GNTrans.TSNumDocsPendientes", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion, 101, 2) = Format(value, "00")
    mbooModificado = True
End Property

Public Property Get TSNumDocsPendientes() As Integer
    TSNumDocsPendientes = Val(Mid$(Opcion, 101, 2))
End Property

Public Property Let TSNOPermitirDocsPend(ByVal value As Boolean)
    Mid$(mProps.Opcion, 103, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSNOPermitirDocsPend() As Boolean
    TSNOPermitirDocsPend = (Mid$(Opcion, 103, 1) = "S")
End Property

Public Property Let TSNumDiasVencidos(ByVal value As Integer)
    If value > 99 Or value < 0 Then
        Err.Raise ERR_INVALIDO, "GNTrans.TSNumDiasVencidos", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion, 104, 2) = Format(value, "00")
    mbooModificado = True
End Property

Public Property Get TSNumDiasVencidos() As Integer
    TSNumDiasVencidos = Val(Mid$(Opcion, 104, 2))
End Property

Public Property Let TSVerificaDiasVencidos(ByVal value As Boolean)
    Mid$(mProps.Opcion, 106, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'*** Diego 01/08/2003
Public Property Get TSVerificaDiasVencidos() As Boolean
    TSVerificaDiasVencidos = (Mid$(Opcion, 106, 1) = "S")
End Property

Public Property Let TSSoloDocsVencidos(ByVal value As Boolean)
    Mid$(mProps.Opcion, 107, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSSoloDocsVencidos() As Boolean
    TSSoloDocsVencidos = (Mid$(Opcion, 107, 1) = "S")
End Property


Public Property Let TSClaveOmitirDocPen(ByVal value As Boolean)
    Mid$(mProps.Opcion, 108, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSClaveOmitirDocPen() As Boolean
    TSClaveOmitirDocPen = (Mid$(Opcion, 108, 1) = "S")
End Property

Public Property Let TSunirCobrosFacturas(ByVal value As Boolean)
    Mid$(mProps.Opcion, 109, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSunirCobrosFacturas() As Boolean
    TSunirCobrosFacturas = (Mid$(Opcion, 109, 1) = "S")
End Property

Public Property Let TSNoGrabaReteMesAnte(ByVal value As Boolean)
   Mid$(mProps.Opcion, 110, 1) = IIf(value, "S", "N")
   mbooModificado = True
End Property

Public Property Get TSNoGrabaReteMesAnte() As Boolean
    TSNoGrabaReteMesAnte = (Mid$(Opcion, 110, 1) = "S")
End Property

Public Property Let TSPideEscopia(ByVal value As Boolean)
    Mid$(mProps.Opcion, 111, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSPideEscopia() As Boolean
    TSPideEscopia = (Mid$(Opcion, 111, 1) = "S")
End Property


Public Property Let TSUtilizaFechaGrabado(ByVal value As Boolean)
   Mid$(mProps.Opcion2, 112, 1) = IIf(value, "S", "N")
   mbooModificado = True
End Property

Public Property Get TSUtilizaFechaGrabado() As Boolean
    TSUtilizaFechaGrabado = (Mid$(Opcion2, 112, 1) = "S")
End Property

Public Property Get TSTransCierre() As String
   TSTransCierre = Trim((Mid$(Opcion2, 113, 5)))
End Property

Public Property Let TSTransCierre(ByVal value As String)
If Len(value) > 5 Then Err.Raise ERR_INVALIDO, "GNTrans.TSTransCierre", MSGERR_INVALIDO
    Mid$(mProps.Opcion2, 113, 5) = String(5, " ")
    Mid$(mProps.Opcion2, 113, 5) = value
    mbooModificado = True
End Property

Public Property Get TSCierreporVende() As Boolean
    TSCierreporVende = (Mid$(Opcion, 112, 1) = "S")
End Property

Public Property Let TSCierreporVende(ByVal value As Boolean)
    Mid$(mProps.Opcion, 112, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

















'======================= Opciones comunes para todos los módulos
Public Property Let AprobacionUnica(ByVal value As Boolean)
    Mid$(mProps.Opcion, 60, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get AprobacionUnica() As Boolean
    AprobacionUnica = (Mid$(Opcion, 60, 1) = "S")
End Property

'*** MAKOTO 03/dic/00 Agregado
Public Property Let CentroCostoVisible(ByVal value As Boolean)
    Mid$(mProps.Opcion, 61, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'*** MAKOTO 03/dic/00 Agregado
Public Property Get CentroCostoVisible() As Boolean
    CentroCostoVisible = (Mid$(Opcion, 61, 1) = "S")
End Property

'*** MAKOTO 07/dic/00 Agregado
Public Property Let PermitirFechaFutura(ByVal value As Boolean)
    Mid$(mProps.Opcion, 62, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'*** MAKOTO 07/dic/00 Agregado
Public Property Get PermitirFechaFutura() As Boolean
    PermitirFechaFutura = (Mid$(Opcion, 62, 1) = "S")
End Property

'*** MAKOTO 06/mar/01 Agregado
Public Property Let ProveedorVisible(ByVal value As Boolean)
    Mid$(mProps.Opcion, 63, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'*** MAKOTO 06/mar/01 Agregado
Public Property Get ProveedorVisible() As Boolean
    ProveedorVisible = (Mid$(Opcion, 63, 1) = "S")
End Property

'*** MAKOTO 06/mar/01 Agregado
Public Property Let ClienteVisible(ByVal value As Boolean)
    Mid$(mProps.Opcion, 64, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'*** MAKOTO 06/mar/01 Agregado
Public Property Get ClienteVisible() As Boolean
    ClienteVisible = (Mid$(Opcion, 64, 1) = "S")
End Property

'*** ALEX dic/04/02
Public Property Let TransAsiento(ByVal value As String)
    If Len(value) > 5 Then Err.Raise ERR_INVALIDO, "GNTrans.TransAsiento", MSGERR_INVALIDO
    Mid$(mProps.Opcion, 65, 5) = String(5, " ")
    Mid$(mProps.Opcion, 65, 5) = value
    mbooModificado = True
End Property

Public Property Get TransAsiento() As String
    TransAsiento = Trim(Mid$(mProps.Opcion, 65, 5))
End Property

'*** oliver 27/ene/03 Agregado
Public Property Let AprobacionAutomatica(ByVal value As Boolean)
    Mid$(mProps.Opcion, 70, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'*** oliver 27/ene/03 Agregado
Public Property Get AprobacionAutomatica() As Boolean
    AprobacionAutomatica = (Mid$(Opcion, 70, 1) = "S")
End Property

'*** Oliver 28/ene/03
Public Property Let F3Deshabilitado(ByVal value As Boolean)
    Mid$(mProps.Opcion, 71, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property
'*** Oliver 28/ene/03
Public Property Get F3Deshabilitado() As Boolean
    F3Deshabilitado = (Mid$(Opcion, 71, 1) = "S")
End Property

'***ANGEL 26/jul/2004 Agregado
Public Property Let VerTodosDocCobrarPagar(ByVal value As Boolean)
    Mid$(mProps.Opcion, 72, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'***ANGEL 26/jul/2004 Agregado
Public Property Get VerTodosDocCobrarPagar() As Boolean
    VerTodosDocCobrarPagar = (Mid$(Opcion, 72, 1) <> "N") ' modificado Oliver antes estaba (Opcion, 72, 1) = "S" pero por compatibilidad para que cuando es la primera ves que revisa esta configuracion tenga la configuracion de VerTodoslosDocPendientes=TRUE
End Property

'======================= Propiedades generales de transacción
Public Property Let CodTrans(ByVal value As String)
    If Not mbooEsNuevo Then
        Err.Raise ERR_NOMODIFICABLE, "GNTrans_CodTrans", MSGERR_NOMODIFICABLE
    End If
    
    If Len(value) > 5 Then
        Err.Raise ERR_INVALIDO, "GNTrans_CodTrans", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodTrans = UCase(value)
    mbooModificado = True
End Property

Public Property Get CodTrans() As String
    CodTrans = Trim$(mProps.CodTrans)
End Property

Public Property Let NombreTrans(ByVal value As String)
    If Len(value) > 40 Then
        Err.Raise ERR_INVALIDO, "GNTrans_NombreTrans", MSGERR_INVALIDO
    End If
    
    mProps.NombreTrans = value
    mbooModificado = True
End Property

Public Property Get NombreTrans() As String
    NombreTrans = Trim$(mProps.NombreTrans)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 120 Then
        Err.Raise ERR_INVALIDO, "GNTrans_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let Modulo(ByVal value As String)
    If Len(value) > 2 Then
        Err.Raise ERR_INVALIDO, "GNTrans_Modulo", MSGERR_INVALIDO
    End If
    
    mProps.Modulo = value
    mbooModificado = True
End Property

Public Property Get Modulo() As String
    Modulo = Trim$(mProps.Modulo)
End Property

Public Property Let CodPantalla(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "GNTrans_CodPantalla", MSGERR_INVALIDO
    End If
    
    mProps.CodPantalla = value
    mbooModificado = True
End Property

Public Property Get CodPantalla() As String
    CodPantalla = Trim$(mProps.CodPantalla)
End Property

Public Property Let CodTransFuente(ByVal value As String)
    If Len(value) > 120 Then 'AUC modificado antes cadena de 80
        Err.Raise ERR_INVALIDO, "GNTrans_CodTransFuente", MSGERR_INVALIDO
    End If
    
    mProps.CodTransFuente = value
    mbooModificado = True
    
End Property

Public Property Get CodTransFuente() As String
    CodTransFuente = Trim$(mProps.CodTransFuente)
End Property


Public Property Let NumTransSiguiente(ByVal value As Long)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "GNTrans_NumTransSiguiente", MSGERR_INVALIDO
    End If
    
    mProps.NumTransSiguiente = value
    mbooModificado = True
End Property

Public Property Get NumTransSiguiente() As Long
    NumTransSiguiente = mProps.NumTransSiguiente
End Property

Public Property Let FechaAuto(ByVal value As Boolean)
    mProps.FechaAuto = value
    mbooModificado = True
End Property

Public Property Get FechaAuto() As Boolean
    FechaAuto = mProps.FechaAuto
End Property

Public Property Let HoraAuto(ByVal value As Boolean)
    mProps.HoraAuto = value
    mbooModificado = True
End Property

Public Property Get HoraAuto() As Boolean
    HoraAuto = mProps.HoraAuto
End Property

Public Property Let HoraVisible(ByVal value As Boolean)
    mProps.HoraVisible = value
    mbooModificado = True
End Property

Public Property Get HoraVisible() As Boolean
    HoraVisible = mProps.HoraVisible
End Property



Public Property Let ArchivoReporte(ByVal value As String)
    If Len(value) > 200 Then
        Err.Raise ERR_INVALIDO, "GNTrans_ArchivoReporte", MSGERR_INVALIDO
    End If
    
    mProps.ArchivoReporte = value
    mbooModificado = True
End Property

Public Property Get ArchivoReporte() As String
    ArchivoReporte = Trim$(mProps.ArchivoReporte)
End Property

Public Property Let IdResponsablePre(ByVal value As Long)
    mProps.IdResponsablePre = value
    mbooModificado = True
End Property

Public Property Get IdResponsablePre() As Long
    IdResponsablePre = mProps.IdResponsablePre
End Property


Public Property Let CodResponsablePre(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodResponsablePre Then
        If Len(value) > 0 Then
            'Actualiza IdResponsable también
            sql = "SELECT IdResponsable FROM GNResponsable WHERE CodResponsable='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNTrans_CodResponsablePre", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdResponsablePre = rs!IdResponsable
            End If
            rs.Close
        Else
            IdResponsablePre = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodResponsablePre = value
    mbooModificado = True
End Property

Public Property Get CodResponsablePre() As String
    Dim sql As String
    Dim rs As Recordset

    'Cuando pide por primera vez
    If (Len(mProps.CodResponsablePre) = 0) And (mProps.IdResponsablePre <> 0) Then
        'Obtiene CodResponsable
        sql = "SELECT CodResponsable FROM GNResponsable WHERE IdResponsable=" & mProps.IdResponsablePre
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
            '*** No generamos error para que no se suspenda proceso en UI
'            Err.Raise ERR_NOHAYCODIGO, "GNTrans_IdResponsablePre", MSGERR_NOHAYCODIGO
'            Exit Property
        Else
            mProps.CodResponsablePre = rs!CodResponsable
        End If
        rs.Close
    End If

    Set rs = Nothing
    CodResponsablePre = Trim$(mProps.CodResponsablePre)
End Property


Public Property Let NecesitaAprobacion(ByVal value As Boolean)
    mProps.NecesitaAprobacion = value
    mbooModificado = True
End Property

Public Property Get NecesitaAprobacion() As Boolean
    NecesitaAprobacion = mProps.NecesitaAprobacion
End Property

Public Property Let RequiereCentroCosto(ByVal value As Boolean)
    mProps.RequiereCentroCosto = value
    mbooModificado = True
End Property

Public Property Get RequiereCentroCosto() As Boolean
    RequiereCentroCosto = mProps.RequiereCentroCosto
End Property

Public Property Let RequiereFuente(ByVal value As Boolean)
    mProps.RequiereFuente = value
    mbooModificado = True
End Property

Public Property Get RequiereFuente() As Boolean
    RequiereFuente = mProps.RequiereFuente
End Property




Public Property Let ImprimeComprobante(ByVal value As String)
    If value <> "S" And value <> "P" And value <> "N" Then
        Err.Raise ERR_INVALIDO, "GNTrans_ImprimeComprobante", MSGERR_INVALIDO
    End If
    
    mProps.ImprimeComprobante = value
    mbooModificado = True
End Property

Public Property Get ImprimeComprobante() As String
    ImprimeComprobante = Trim$(mProps.ImprimeComprobante)
End Property

Public Property Let ImprimeSoloAprobado(ByVal value As Boolean)
    mProps.ImprimeSoloAprobado = value
    mbooModificado = True
End Property

Public Property Get ImprimeSoloAprobado() As Boolean
    ImprimeSoloAprobado = mProps.ImprimeSoloAprobado
End Property

Public Property Let NumTransModificable(ByVal value As Boolean)
    mProps.NumTransModificable = value
    mbooModificado = True
End Property

Public Property Get NumTransModificable() As Boolean
    NumTransModificable = mProps.NumTransModificable
End Property

Public Property Let MuestraAsiento(ByVal value As Boolean)
    mProps.MuestraAsiento = value
    mbooModificado = True
End Property

Public Property Get MuestraAsiento() As Boolean
    MuestraAsiento = mProps.MuestraAsiento
End Property

Public Property Let Opcion(ByVal value As String)
    If Len(value) > 120 Then
        Err.Raise ERR_INVALIDO, "GNTrans_Opcion", MSGERR_INVALIDO
    End If
    
    mProps.Opcion = value
    mbooModificado = True
End Property

Public Property Get Opcion() As String
    Opcion = Trim$(mProps.Opcion)
End Property

Public Property Let BandEliminable(ByVal value As Boolean)
    mProps.BandEliminable = value
    mbooModificado = True
End Property

Public Property Get BandEliminable() As Boolean
    BandEliminable = mProps.BandEliminable
End Property

Public Property Let BandValida(ByVal value As Boolean)
    mProps.BandValida = value
    mbooModificado = True
End Property

Public Property Get BandValida() As Boolean
    BandValida = mProps.BandValida
End Property

Public Property Let BandVisible(ByVal value As Boolean)
    mProps.BandVisible = value
    mbooModificado = True
End Property

Public Property Get BandVisible() As Boolean
    BandVisible = mProps.BandVisible
End Property

Public Property Let IdBodegaPre(ByVal value As Long)
    mProps.IdBodegaPre = value
    mbooModificado = True
End Property

Public Property Get IdBodegaPre() As Long
    IdBodegaPre = mProps.IdBodegaPre
End Property


Public Property Let CodBodegaPre(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodBodegaPre Then
        If Len(value) > 0 Then
            'Actualiza IdBodega también
            sql = "SELECT IdBodega FROM "
            If Me.Modulo = "IV" Then
                sql = sql & " IVBodega"
            Else
                    sql = sql & " AFBodega"
            End If
            
            sql = sql & " WHERE CodBodega='" & value & "'"
            
            If Me.CodPantalla = "AFBDBQD" Then
                sql = "SELECT IdProvcli FROM PCPROVCLI"
                sql = sql & " WHERE Codprovcli='" & value & "'"
            End If
            
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNTrans_CodBodegaPre", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                If Me.CodPantalla = "AFBDBQD" Then
                    IdBodegaPre = rs!IdProvcli
                Else
                    IdBodegaPre = rs!IdBodega
                End If
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdBodegaPre = 0
        End If
    End If

    mProps.CodBodegaPre = value
    mbooModificado = True
End Property

Public Property Get CodBodegaPre() As String
    CodBodegaPre = Trim$(mProps.CodBodegaPre)
End Property



Public Property Let IdProveedorPre(ByVal value As Long)
    mProps.IdProveedorPre = value
    mbooModificado = True
End Property

Public Property Get IdProveedorPre() As Long
    IdProveedorPre = mProps.IdProveedorPre
End Property


Public Property Let CodProveedorPre(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodProveedorPre Then
        If Len(value) > 0 Then
            'Actualiza IdProvCli también
            sql = "SELECT IdProvCli FROM PCProvCli WHERE CodProvCli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNTrans_CodProveedorPre", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdProveedorPre = rs!IdProvcli
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdProveedorPre = 0
        End If
    End If

    mProps.CodProveedorPre = value
    mbooModificado = True
End Property

Public Property Get CodProveedorPre() As String
    CodProveedorPre = Trim$(mProps.CodProveedorPre)
End Property


Public Property Let IdClientePre(ByVal value As Long)
    mProps.IdClientePre = value
    mbooModificado = True
End Property

Public Property Get IdClientePre() As Long
    IdClientePre = mProps.IdClientePre
End Property


Public Property Let CodClientePre(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodClientePre Then
        If Len(value) > 0 Then
            'Actualiza IdProvCli también
            sql = "SELECT IdProvCli FROM PCProvCli WHERE CodProvCli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNTrans_CodClientePre", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdClientePre = rs!IdProvcli
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdClientePre = 0
        End If
    End If

    mProps.CodClientePre = value
    mbooModificado = True
End Property

Public Property Get CodClientePre() As String
    CodClientePre = Trim$(mProps.CodClientePre)
End Property



Public Property Let CantidadPre(ByVal value As Currency)
    mProps.CantidadPre = value
    mbooModificado = True
End Property

Public Property Get CantidadPre() As Currency
    CantidadPre = mProps.CantidadPre
End Property

Public Property Let IdFormaPre(ByVal value As Long)
    mProps.IdFormaPre = value
    mbooModificado = True
End Property

Public Property Get IdFormaPre() As Long
    IdFormaPre = mProps.IdFormaPre
End Property


Public Property Let CodFormaPre(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodFormaPre Then
        If Len(value) > 0 Then
            'Actualiza IdForma también
            sql = "SELECT IdForma FROM TSFormaCobroPago WHERE CodForma='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNTrans_CodFormaPre", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdFormaPre = rs!IDForma
            End If
            rs.Close
        Else
            IdFormaPre = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodFormaPre = value
    mbooModificado = True
End Property

Public Property Get CodFormaPre() As String
    CodFormaPre = Trim$(mProps.CodFormaPre)
End Property




Public Property Let MonedaPre(ByVal value As String)
    If Len(value) > 5 Then
        Err.Raise ERR_INVALIDO, "GNTrans_MonedaPre", MSGERR_INVALIDO
    End If
    
    mProps.MonedaPre = value
    mbooModificado = True
End Property

Public Property Get MonedaPre() As String
    MonedaPre = Trim$(mProps.MonedaPre)
End Property

Public Property Let ColVisible(ByVal value As Long)
    mProps.ColVisible = value
    mbooModificado = True
End Property

Public Property Get ColVisible() As Long
    ColVisible = mProps.ColVisible
End Property

Public Property Let ColEditable(ByVal value As Long)
    mProps.ColEditable = value
    mbooModificado = True
End Property

Public Property Get ColEditable() As Long
    ColEditable = mProps.ColEditable
End Property

Public Property Let ColSeleccionable(ByVal value As Long)
    mProps.ColSeleccionable = value
    mbooModificado = True
End Property

Public Property Get ColSeleccionable() As Long
    ColSeleccionable = mProps.ColSeleccionable
End Property

Public Property Let AfectaCantidad(ByVal value As Boolean)
    mProps.AfectaCantidad = value
    mbooModificado = True
End Property

Public Property Get AfectaCantidad() As Boolean
    AfectaCantidad = mProps.AfectaCantidad
End Property

Public Property Let AfectaSaldoPC(ByVal value As Boolean)
    mProps.AfectaSaldoPC = value
    mbooModificado = True
End Property

Public Property Get AfectaSaldoPC() As Boolean
    AfectaSaldoPC = mProps.AfectaSaldoPC
End Property




Public Property Let ImportaPCK(ByVal value As Boolean)
    If value Then
        mProps.BandImportacion = mProps.BandImportacion Or &H1
    Else
        mProps.BandImportacion = mProps.BandImportacion And &HFE
    End If
    mbooModificado = True
End Property

Public Property Get ImportaPCK() As Boolean
    ImportaPCK = CBool(mProps.BandImportacion And &H1)
End Property

Public Property Let ImportaIVKR(ByVal value As Boolean)
    If value Then
        mProps.BandImportacion = mProps.BandImportacion Or &H2
    Else
        mProps.BandImportacion = mProps.BandImportacion And &HFD
    End If
    mbooModificado = True
End Property

Public Property Get ImportaIVKR() As Boolean
    ImportaIVKR = CBool(mProps.BandImportacion And &H2)
End Property

Public Property Let ImportaIVK(ByVal value As Boolean)
    If value Then
        ImportaIVKReal = False          'Hay que desactivar ImportaIVKReal
        mProps.BandImportacion = mProps.BandImportacion Or &H4
    Else
        mProps.BandImportacion = mProps.BandImportacion And &HFB
    End If
    mbooModificado = True
End Property

Public Property Get ImportaIVK() As Boolean
    ImportaIVK = CBool(mProps.BandImportacion And &H4)
End Property


Public Property Let ImportaTSK(ByVal value As Boolean)
    If value Then
        mProps.BandImportacion = mProps.BandImportacion Or &H8
    Else
        mProps.BandImportacion = mProps.BandImportacion And &HF7
    End If
    mbooModificado = True
End Property

Public Property Get ImportaTSK() As Boolean
    ImportaTSK = CBool(mProps.BandImportacion And &H8)
End Property

Public Property Let ImportaAFK(ByVal value As Boolean)
    If value Then
        mProps.BandImportacion = mProps.BandImportacion Or &H10
    Else
        mProps.BandImportacion = mProps.BandImportacion And &HEF
    End If
    mbooModificado = True
End Property

Public Property Get ImportaAFK() As Boolean
    ImportaAFK = CBool(mProps.BandImportacion And &H10)
End Property

Public Property Let ImportaIVKReal(ByVal value As Boolean)
    If value Then
        ImportaIVK = False          'Hay que desactivar ImportaIVK
        mProps.BandImportacion = mProps.BandImportacion Or &H20
    Else
        mProps.BandImportacion = mProps.BandImportacion And &HDF
    End If
    mbooModificado = True
End Property

Public Property Get ImportaIVKReal() As Boolean
    ImportaIVKReal = CBool(mProps.BandImportacion And &H20)
End Property

'*** MAKOTO 18/nov/00 Agregado para importación de asiento contable
Public Property Let ImportaCTD(ByVal value As Boolean)
    If value Then
        mProps.BandImportacion = mProps.BandImportacion Or &H40
    Else
        mProps.BandImportacion = mProps.BandImportacion And &HBF
    End If
    mbooModificado = True
End Property

Public Property Get ImportaCTD() As Boolean
    ImportaCTD = CBool(mProps.BandImportacion And &H40)
End Property

'*** MAKOTO 07/feb/01 Agregado
Public Property Let ImportaTSKRet(ByVal value As Boolean)
    If value Then
        mProps.BandImportacion = mProps.BandImportacion Or &H80
    Else
        mProps.BandImportacion = mProps.BandImportacion And &H7F
    End If
    mbooModificado = True
End Property

Public Property Get ImportaTSKRet() As Boolean
    ImportaTSKRet = CBool(mProps.BandImportacion And &H80)
End Property

Public Property Let ImportaSoloAprobado(ByVal value As Boolean)
    mProps.ImportaSoloAprobado = value
    mbooModificado = True
End Property

Public Property Get ImportaSoloAprobado() As Boolean
    ImportaSoloAprobado = mProps.ImportaSoloAprobado
End Property

Public Property Get fechagrabado() As Date
    fechagrabado = mProps.fechagrabado
End Property

Public Property Get EsNuevo() As Boolean
    EsNuevo = mbooEsNuevo
End Property

Friend Sub SetModificado()
    mbooModificado = True
End Sub

Public Property Get Modificado() As Boolean
    Modificado = mbooModificado
End Property

Public Function ListaTransSiguiente() As Variant
    Dim sql As String, v As Variant, c As String, rs As Recordset
    
#If DAOLIB Then
    c = "*"         'Comodin
#Else
    c = "%"         'Comodin
#End If

    sql = "SELECT CodTrans,Descripcion FROM GNTrans WHERE " & _
            "CodTransFuente LIKE '" & c & "(" & mProps.CodTrans & ")" & c & "' ORDER BY CodTrans"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    ListaTransSiguiente = MiGetRows(rs)
    rs.Close
    Set rs = Nothing
End Function


Private Sub Class_Initialize()
    Set mcolAsientos = New Collection
    Set mcolRecargos = New Collection
    Set mcolAsientosPresup = New Collection

    
    'Valores predeterminados
    With mProps
        .FechaAuto = False
        .HoraAuto = False
        .HoraVisible = False
        .AfectaCantidad = True
        .ImprimeComprobante = "P"
        .BandEliminable = True
        .BandValida = True
        .BandVisible = True
        .NumTransModificable = False
        .MuestraAsiento = True
        .NecesitaAprobacion = False
        .RequiereCentroCosto = False
'        .MonedaPre = "SUCRE"
        .MonedaPre = "USD"              '*** MAKOTO 02/oct/00 Modificado
        .Opcion = String$(120, " ")
        .NumTransSiguiente = 1
        .AfectaCantidad = True
        .AfectaSaldoPC = True
        .BandImportacion = -1
        .NumTransUltimo = 1  'jeaa 15/11/2006
        .NumTransInicio = 1
        .NumTransUltimoPreimpreso = 0 'jeaa 25/09/2007
        .AsientoGastoxZona = False
        .OmitirGastoxZona = False
        .AfectaCantidadDocumento = False
        .NumCopiasImp = 1
    End With
    
    Me.IVPermitirSignoNegativo = False      '*** MAKOTO 25/ene/01 Agregado
    
    
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mcolAsientos = Nothing
    Set mcolRecargos = Nothing
    Set mProps.objEmpresa = Nothing
End Sub

Public Property Let CotizacionAuto(ByVal value As Boolean)
    mProps.CotizacionAuto = value
    mbooModificado = True
End Property

Public Property Get CotizacionAuto() As Boolean
    CotizacionAuto = mProps.CotizacionAuto
End Property

Public Property Let CodTransHijoCI(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNTrans_CodTransHijoCI", MSGERR_INVALIDO
    End If
    
    mProps.CodTransHijoCI = value
    mbooModificado = True
End Property

Public Property Get CodTransHijoCI() As String
    CodTransHijoCI = Trim$(mProps.CodTransHijoCI)
End Property

Public Property Let CodTransHijoRUC(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNTrans_CodTransHijoRUC", MSGERR_INVALIDO
    End If
    
    mProps.CodTransHijoRUC = value
    mbooModificado = True
End Property

Public Property Get CodTransHijoRUC() As String
    CodTransHijoRUC = Trim$(mProps.CodTransHijoRUC)
End Property

Public Property Let CodTransHijoOtro(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNTrans_CodTransHijoOtro", MSGERR_INVALIDO
    End If
    
    mProps.CodTransHijoOtro = value
    mbooModificado = True
End Property

Public Property Get CodTransHijoOtro() As String
    CodTransHijoOtro = Trim$(mProps.CodTransHijoOtro)
End Property

'jeaa 15/11/2005
Public Property Let CodLineaDevolucion(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNTrans_CodLineaDevolucion", MSGERR_INVALIDO
    End If
    mProps.CodLineaDevolucion = value
    mbooModificado = True
End Property
Public Property Get CodLineaDevolucion() As String
    CodLineaDevolucion = Trim$(mProps.CodLineaDevolucion)
End Property
 
 'jeaa 04/04/2006
Public Property Let NumSerieEstablecimiento(ByVal value As String)
    If Len(value) > 3 Then
        Err.Raise ERR_INVALIDO, "GNTrans_NumSerieEstablecimiento ", MSGERR_INVALIDO
    End If
    
    mProps.NumSerieEstablecimiento = value
    mbooModificado = True
End Property

Public Property Get NumSerieEstablecimiento() As String
    NumSerieEstablecimiento = Trim$(mProps.NumSerieEstablecimiento)
End Property

Public Property Let NumSeriePunto(ByVal value As String)
    If Len(value) > 3 Then
        Err.Raise ERR_INVALIDO, "GNTrans_NumSeriePunto ", MSGERR_INVALIDO
    End If
    
    mProps.NumSeriePunto = value
    mbooModificado = True
End Property

Public Property Get NumSeriePunto() As String
    NumSeriePunto = Trim$(mProps.NumSeriePunto)
End Property

Public Property Let NumAutorizacion(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "GNTrans_NumAutorizacion ", MSGERR_INVALIDO
    End If
    
    mProps.NumAutorizacion = value
    mbooModificado = True
End Property

Public Property Get NumAutorizacion() As String
    NumAutorizacion = Trim$(mProps.NumAutorizacion)
End Property

Public Property Let FechaCaducidad(ByVal value As Date)
    mProps.FechaCaducidad = value
    mbooModificado = True
End Property

Public Property Get FechaCaducidad() As Date
    FechaCaducidad = Trim$(mProps.FechaCaducidad)
End Property

Public Property Let FechaAutorizacion(ByVal value As Date)
    mProps.FechaAutorizacion = value
    mbooModificado = True
End Property

Public Property Get FechaAutorizacion() As Date
    FechaAutorizacion = Trim$(mProps.FechaAutorizacion)
End Property

Public Property Let CodTransHijoPasp(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNTrans_CodTransHijoPasp", MSGERR_INVALIDO
    End If
    mProps.CodTransHijoPasp = value
    mbooModificado = True
End Property
Public Property Get CodTransHijoPasp() As String
    CodTransHijoPasp = Trim$(mProps.CodTransHijoPasp)
End Property

Public Property Let AnexoCodTipoComp(ByVal value As String)
    If Len(value) > 2 Then
        Err.Raise ERR_INVALIDO, "GNTrans_AnexoCodTipoComp ", MSGERR_INVALIDO
    End If
    
    mProps.AnexoCodTipoComp = value
    mbooModificado = True
End Property

Public Property Get AnexoCodTipoComp() As String
    AnexoCodTipoComp = Trim$(mProps.AnexoCodTipoComp)
End Property

Public Property Let AnexoCodTipoTrans(ByVal value As String)
    If Len(value) > 2 Then
        Err.Raise ERR_INVALIDO, "GNTrans_AnexoCodTipoTrans ", MSGERR_INVALIDO
    End If
    
    mProps.AnexoCodTipoTrans = value
    mbooModificado = True
End Property

Public Property Get AnexoCodTipoTrans() As String
    AnexoCodTipoTrans = Trim$(mProps.AnexoCodTipoTrans)
End Property


Public Property Let NumTransUltimo(ByVal value As String)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "GNTrans_NumTransUltimo", MSGERR_INVALIDO
    End If
    mProps.NumTransUltimo = value
    mbooModificado = True
End Property

Public Property Get NumTransUltimo() As String
    NumTransUltimo = mProps.NumTransUltimo
End Property

Public Property Let FechaCaducidadOld(ByVal value As Date)
    mProps.FechaCaducidadOld = value
    mbooModificado = True
End Property

Public Property Get FechaCaducidadOld() As Date
    FechaCaducidadOld = Trim$(mProps.FechaCaducidadOld)
End Property

Public Property Let FechaAutorizacionOld(ByVal value As Date)
    mProps.FechaAutorizacionOld = value
    mbooModificado = True
End Property

Public Property Get FechaAutorizacionOld() As Date
    FechaAutorizacionOld = Trim$(mProps.FechaAutorizacionOld)
End Property


Public Property Let NumAutorizacionOld(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "GNTrans_NumAutorizacionOld ", MSGERR_INVALIDO
    End If
    mProps.NumAutorizacionOld = value
    mbooModificado = True
End Property
Public Property Get NumAutorizacionOld() As String
    NumAutorizacionOld = Trim$(mProps.NumAutorizacionOld)
End Property

Public Property Let NumTransInicio(ByVal value As String)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "GNTrans_NumTransInicio", MSGERR_INVALIDO
    End If
    mProps.NumTransInicio = value
    mbooModificado = True
End Property
Public Property Get NumTransInicio() As String
    NumTransInicio = mProps.NumTransInicio
End Property
Public Property Let NumTransInicioOld(ByVal value As String)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "GNTrans_NumTransInicioOld", MSGERR_INVALIDO
    End If
    mProps.NumTransInicioOld = value
    mbooModificado = True
End Property
Public Property Get NumTransInicioOld() As String
    NumTransInicioOld = mProps.NumTransInicioOld
End Property

Public Property Let CodTransImp(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNTrans_CodTransImp", MSGERR_INVALIDO
    End If
    
    mProps.CodTransImp = value
    mbooModificado = True
End Property

Public Property Get CodTransImp() As String
    CodTransImp = Trim$(mProps.CodTransImp)
End Property

Public Property Let AnexoCodTipoRetencion(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "GNTrans_AnexoCodTipoRetencion ", MSGERR_INVALIDO
    End If
    
    mProps.AnexoCodTipoRetencion = value
    mbooModificado = True
End Property

Public Property Get AnexoCodTipoRetencion() As String
    AnexoCodTipoRetencion = Trim$(mProps.AnexoCodTipoRetencion)
End Property

Public Property Let ImportacionUnicaTrans(ByVal value As Boolean)
    mProps.ImportacionUnicaTrans = value
    mbooModificado = True
End Property

Public Property Get ImportacionUnicaTrans() As Boolean
    ImportacionUnicaTrans = mProps.ImportacionUnicaTrans
End Property

Public Property Let TransDespuesDevolucion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "GNTrans_TransDespuesDevolucion", MSGERR_INVALIDO
    End If
    mProps.TransDespuesDevolucion = value
    mbooModificado = True
End Property
Public Property Get TransDespuesDevolucion() As String
    TransDespuesDevolucion = Trim$(mProps.TransDespuesDevolucion)
End Property

Public Property Let TransHistorial(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "GNTrans_TransHistorial", MSGERR_INVALIDO
    End If
    mProps.TransHistorial = value
    mbooModificado = True
End Property
Public Property Get TransHistorial() As String
    TransHistorial = Trim$(mProps.TransHistorial)
End Property


Public Property Let FormasLimteCredito(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "GNTrans_FormasLimteCredito", MSGERR_INVALIDO
    End If
    mProps.FormasLimteCredito = value
    mbooModificado = True
End Property
Public Property Get FormasLimteCredito() As String
    FormasLimteCredito = Trim$(mProps.FormasLimteCredito)
End Property

'AUC 30/08/07
Public Property Let IVGeneraTransAuto(ByVal value As Boolean)
        Mid$(mProps.Opcion, 114, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property
Public Property Get IVGeneraTransAuto() As Boolean
    IVGeneraTransAuto = Mid$(Opcion, 114, 1) = "S"
End Property
Public Property Let IVBodegasDestinos(ByVal value As Boolean)
        Mid$(mProps.Opcion, 115, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property
Public Property Get IVBodegasDestinos() As Boolean
    IVBodegasDestinos = Mid$(Opcion, 115, 1) = "S"
End Property
'jeaa 25/09/2007
Public Property Let IVTipoCompraBienes(ByVal value As Boolean)
        Mid$(mProps.Opcion, 116, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property

Public Property Get IVTipoCompraBienes() As Boolean
    IVTipoCompraBienes = Mid$(Opcion, 116, 1) = "S"
End Property

Public Property Let NumTransUltimoPreimpreso(ByVal value As String)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "GNTrans_NumTransUltimoPreimpreso", MSGERR_INVALIDO
    End If
    mProps.NumTransUltimoPreimpreso = value
    mbooModificado = True
End Property
Public Property Get NumTransUltimoPreimpreso() As String
    NumTransUltimoPreimpreso = mProps.NumTransUltimoPreimpreso
End Property

'AUC 27/09/2007
Public Property Let IVVisItemPromocion(ByVal value As Boolean)
        Mid$(mProps.Opcion, 117, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property
Public Property Get IVVisItemPromocion() As Boolean
    IVVisItemPromocion = Mid$(Opcion, 117, 1) = "S"
End Property

'AUC 27/09/2007
Public Property Let IVRetornaNumPrecio(ByVal value As Boolean)
        Mid$(mProps.Opcion, 118, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property
Public Property Get IVRetornaNumPrecio() As Boolean
    IVRetornaNumPrecio = Mid$(Opcion, 118, 1) = "S"
End Property

Public Property Let IVNumLineas(ByVal value As String)
    Mid$(mProps.Opcion, 119, 2) = value
    mbooModificado = True
End Property

Public Property Get IVNumLineas() As String
    IVNumLineas = Mid$(Opcion, 119, 2)
End Property
'AUC 30/11/07
Public Property Let VerFacturados(ByVal value As Boolean)
    Mid$(mProps.Opcion, 121, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get VerFacturados() As Boolean
    VerFacturados = (Mid$(Opcion, 121, 1) <> "N")
End Property

Public Property Let VerComprados(ByVal value As Boolean)
    Mid$(mProps.Opcion, 122, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get VerComprados() As Boolean
    VerComprados = (Mid$(Opcion, 122, 1) <> "N")
End Property

'AUC 27/11/07
Public Property Get IVUtilidades() As Integer
    IVUtilidades = Val(Mid$(mProps.Opcion, 123, 1))
End Property

Public Property Let IVUtilidades(ByVal value As Integer)
    If value > 9 Or value < 0 Then      '0: no controla 1 mensaje de advertencia, 2 no permite grabar
        Err.Raise ERR_INVALIDO, "GNTrans.IVUtilidades", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion, 123, 1) = value
    mbooModificado = True
End Property

Public Property Get IVCostoCalculaUtil() As Integer
    IVCostoCalculaUtil = Val(Mid$(mProps.Opcion, 124, 1))
End Property

Public Property Let IVCostoCalculaUtil(ByVal value As Integer)
    If value > 9 Or value < 0 Then      '0: no controla 1 mensaje de advertencia, 2 no permite grabar
        Err.Raise ERR_INVALIDO, "GNTrans.IVCostoCalculaUtil", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion, 124, 1) = value
    mbooModificado = True
End Property

'jeaa  17/01/2008 sri equidad tributaria
Public Property Let IVOmitirIVA(ByVal value As Boolean)
    Mid$(mProps.Opcion, 125, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVOmitirIVA() As Boolean
    IVOmitirIVA = (Mid$(Opcion, 125, 1) = "S")
End Property

'jeaa  26/03/2008
Public Property Let IVBloquearBodegas(ByVal value As Boolean)
    Mid$(mProps.Opcion, 126, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVBloquearBodegas() As Boolean
    IVBloquearBodegas = (Mid$(Opcion, 126, 1) = "S")
End Property

'jeaa  21/04/2008
Public Property Let IVIngresoAuto(ByVal value As Boolean)
    Mid$(mProps.Opcion, 127, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVIngresoAuto() As Boolean
    IVIngresoAuto = (Mid$(Opcion, 127, 1) = "S")
End Property

'jeaa 21/04/2008
Public Property Let IVTransAuto(ByVal value As String)
    If Len(value) > 5 Then Err.Raise ERR_INVALIDO, "GNTrans.IVTransAuto", MSGERR_INVALIDO
    Mid$(mProps.Opcion, 128, 5) = String(5, " ")
    Mid$(mProps.Opcion, 128, 5) = value
    mbooModificado = True
End Property

Public Property Get IVTransAuto() As String
    IVTransAuto = Trim(Mid$(Opcion, 128, 5))
End Property


Public Property Let IVVisualizaIngresoAuto(ByVal value As Boolean)
    Mid$(mProps.Opcion, 133, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property
Public Property Get IVVisualizaIngresoAuto() As Boolean
    IVVisualizaIngresoAuto = (Mid$(Opcion, 133, 1) = "S")
End Property

Public Property Let SalirDespuesGrabar(ByVal value As Boolean)
        Mid$(mProps.Opcion, 134, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property

Public Property Get SalirDespuesGrabar() As Boolean
    SalirDespuesGrabar = (Mid$(Opcion, 134, 1) = "S")
End Property

Public Property Let AbrirCaja(ByVal value As Boolean)
        Mid$(mProps.Opcion, 135, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property

Public Property Get AbrirCaja() As Boolean
    AbrirCaja = (Mid$(Opcion, 135, 1) = "S")
End Property

Public Property Let TSPago(ByVal value As Boolean)
    Mid$(mProps.Opcion, 136, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSPago() As Boolean
    TSPago = (Mid$(Opcion, 136, 1) = "S")
End Property

Public Property Let TSVisibleF10(ByVal value As Boolean)
    Mid$(mProps.Opcion, 137, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSVisibleF10() As Boolean
    TSVisibleF10 = (Mid$(Opcion, 137, 1) = "S")
End Property

'AUC  21/12/06
Public Property Let IVPideDescripcion(ByVal value As Boolean)
    Mid$(mProps.Opcion, 138, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVPideDescripcion() As Boolean
    IVPideDescripcion = (Mid$(Opcion, 138, 1) = "S")
End Property

Public Property Let IVPideObservacion(ByVal value As Boolean)
    Mid$(mProps.Opcion, 139, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVPideObservacion() As Boolean
    IVPideObservacion = (Mid$(Opcion, 139, 1) = "S")
End Property
'AUC 02/08/07 SoloParaMostrarSubitems Totalizados 'presupuestos construccion'
Public Property Let IVMostrarSubItemsTotal(ByVal value As Boolean)
    Mid$(mProps.Opcion, 140, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVMostrarSubItemsTotal() As Boolean
    IVMostrarSubItemsTotal = (Mid$(Opcion, 140, 1) = "S")
End Property

Public Property Let GrabaHistorial(ByVal value As Boolean)
    Mid$(mProps.Opcion, 141, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get GrabaHistorial() As Boolean
    GrabaHistorial = (Mid$(Opcion, 141, 1) = "S")
End Property

'jeaa 09/06/2008
Public Property Let IVTotItemSignosDif(ByVal value As Boolean)
    Mid$(mProps.Opcion, 142, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVTotItemSignosDif() As Boolean
    IVTotItemSignosDif = (Mid$(Opcion, 142, 1) = "S")
End Property

Public Property Let ImpRespetarSignosCantidad(ByVal value As Boolean)
    Mid$(mProps.Opcion, 143, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get ImpRespetarSignosCantidad() As Boolean
    ImpRespetarSignosCantidad = (Mid$(Opcion, 143, 1) = "S")
End Property
'jeaa 23/06/2008
'configuracion para importar transacciones
Public Property Let IVImportaNoAprobado(ByVal value As Boolean)
    Mid$(mProps.Opcion, 144, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVImportaNoAprobado() As Boolean
    IVImportaNoAprobado = (Mid$(Opcion, 144, 1) = "S")
End Property

Public Property Let IVImportaAprobado(ByVal value As Boolean)
    Mid$(mProps.Opcion, 145, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVImportaAprobado() As Boolean
    IVImportaAprobado = (Mid$(Opcion, 145, 1) = "S")
End Property

Public Property Let IVImportaDespachado(ByVal value As Boolean)
    Mid$(mProps.Opcion, 146, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVImportaDespachado() As Boolean
    IVImportaDespachado = (Mid$(Opcion, 146, 1) = "S")
End Property


Public Property Let IVImportaSemDespachado(ByVal value As Boolean)
    Mid$(mProps.Opcion, 147, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVImportaSemDespachado() As Boolean
    IVImportaSemDespachado = (Mid$(Opcion, 147, 1) = "S")
End Property

Public Property Let NoCambiarEstadoTransFuente(ByVal value As Boolean)
    Mid$(mProps.Opcion, 148, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get NoCambiarEstadoTransFuente() As Boolean
    NoCambiarEstadoTransFuente = (Mid$(Opcion, 148, 1) = "S")
End Property

Public Property Let GuardaRelacionImportar(ByVal value As Boolean)
    Mid$(mProps.Opcion, 149, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get GuardaRelacionImportar() As Boolean
    GuardaRelacionImportar = (Mid$(Opcion, 149, 1) = "S")
End Property

'jeaa 01/09/2008
Public Property Let IVVisualizarNumPagos(ByVal value As Boolean)
    Mid$(mProps.Opcion, 150, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVVisualizarNumPagos() As Boolean
    IVVisualizarNumPagos = (Mid$(Opcion, 150, 1) = "S")
End Property

'jeaa 10/09/2008
Public Property Let IVNoEliminarItemsImportar(ByVal value As Boolean)
    Mid$(mProps.Opcion, 151, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVNoEliminarItemsImportar() As Boolean
    IVNoEliminarItemsImportar = (Mid$(Opcion, 151, 1) = "S")
End Property

'jeaa 22/09/2008
Public Property Let IVAutoImpresor(ByVal value As Boolean)
    Mid$(mProps.Opcion, 152, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVAutoImpresor() As Boolean
    IVAutoImpresor = (Mid$(Opcion, 152, 1) = "S")
End Property

'jeaa 25/09/2008 Distablasa
Public Property Let IVClaveOmitirDocPen(ByVal value As Boolean)
    Mid$(mProps.Opcion, 153, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVClaveOmitirDocPen() As Boolean
    IVClaveOmitirDocPen = (Mid$(Opcion, 153, 1) = "S")
End Property
'control sola para Hormimipios 06/10/2008

Public Property Let IVOmitirControlCatidadTotal(ByVal value As Boolean)
        Mid$(mProps.Opcion, 154, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property
Public Property Get IVOmitirControlCatidadTotal() As Boolean
    IVOmitirControlCatidadTotal = Mid$(Opcion, 154, 1) = "S"
End Property

'***Esteban 10/10/2008 Descuento  por Forma CobroPago
Public Property Let IVDescXFormaCP(ByVal value As Boolean)
    Mid$(mProps.Opcion, 155, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVDescXFormaCP() As Boolean
    IVDescXFormaCP = (Mid$(Opcion, 155, 1) = "S")
End Property

'***Esteban 10/10/2008 Modifica el precio
Public Property Let IVModificaPrecio(ByVal value As Boolean)
    Mid$(mProps.Opcion, 156, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVModificaPrecio() As Boolean
    IVModificaPrecio = (Mid$(Opcion, 156, 1) = "S")
End Property

'***Esteban 27/10/2008 Modifica el precio
Public Property Let ControlVuelos(ByVal value As Boolean)
    Mid$(mProps.Opcion, 157, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get ControlVuelos() As Boolean
    ControlVuelos = (Mid$(Opcion, 157, 1) = "S")
End Property

'***Esteban 05/01/2009 Verifica Fecha UltimaCompra
Public Property Let IVVerificaFechaCompra(ByVal value As Boolean)
    Mid$(mProps.Opcion, 158, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVVerificaFechaCompra() As Boolean
    IVVerificaFechaCompra = (Mid$(Opcion, 158, 1) = "S")
End Property

Public Property Let IVNumDiasVerificaFechaCompra(ByVal value As String)
    If Len(value) > 3 Then Err.Raise ERR_INVALIDO, "GNTrans.IVNumDiasVerificaFechaCompra", MSGERR_INVALIDO
    Mid$(mProps.Opcion, 159, 3) = String(3, " ")
    Mid$(mProps.Opcion, 159, 3) = value
    mbooModificado = True
End Property

Public Property Get IVNumDiasVerificaFechaCompra() As String
    IVNumDiasVerificaFechaCompra = Trim(Mid$(Opcion, 159, 3))
End Property

Public Property Let GuardaMotivoAnulacion(ByVal value As Boolean)
    Mid$(mProps.Opcion, 162, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get GuardaMotivoAnulacion() As Boolean
    GuardaMotivoAnulacion = (Mid$(Opcion, 162, 1) = "S")
End Property

'***Angel 21/03t/2009 Datos adicionales de Centro Costo
Public Property Let IVVisualizaDatosGNC(ByVal value As Boolean)
    Mid$(mProps.Opcion, 163, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVVisualizaDatosGNC() As Boolean
    IVVisualizaDatosGNC = (Mid$(Opcion, 163, 1) = "S")
End Property

Public Property Let IVControlaItemServicio(ByVal value As Boolean)
        Mid$(mProps.Opcion, 164, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property

Public Property Get IVControlaItemServicio() As Boolean
    IVControlaItemServicio = (Mid$(Opcion, 164, 1) = "S")
End Property

Public Property Let TerminarCentroAlGrabar(ByVal value As Boolean)
        Mid$(mProps.Opcion, 165, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property

Public Property Get TerminarCentroAlGrabar() As Boolean
    TerminarCentroAlGrabar = (Mid$(Opcion, 165, 1) = "S")
End Property

'jeaa 09/06/2008
Public Property Let IVActualizaCantidadDespachada(ByVal value As Boolean)
    Mid$(mProps.Opcion, 166, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVActualizaCantidadDespachada() As Boolean
    IVActualizaCantidadDespachada = (Mid$(Opcion, 166, 1) = "S")
End Property

'jeaa 26/05/2009
Public Property Let IVAplicaFinaciamiento(ByVal value As Boolean)
    Mid$(mProps.Opcion, 167, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVAplicaFinaciamiento() As Boolean
    IVAplicaFinaciamiento = (Mid$(Opcion, 167, 1) = "S")
End Property

'jeaa 26/05/2009
Public Property Let ivBloquearCuotas(ByVal value As Boolean)
    Mid$(mProps.Opcion, 168, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get ivBloquearCuotas() As Boolean
    ivBloquearCuotas = (Mid$(Opcion, 168, 1) = "S")
End Property

'jeaa 26/05/2009
Public Property Let IVRedondearCobrosPagos(ByVal value As Boolean)
    Mid$(mProps.Opcion, 169, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVRedondearCobrosPagos() As Boolean
    IVRedondearCobrosPagos = (Mid$(Opcion, 169, 1) = "S")
End Property

'jeaa 26/05/2009
Public Property Let IVNumDecimalesCuotas(ByVal value As String)
    Mid$(mProps.Opcion, 170, 1) = value
    mbooModificado = True
End Property

Public Property Get IVNumDecimalesCuotas() As String
    IVNumDecimalesCuotas = (Mid$(Opcion, 170, 1))
End Property

'jeaa 26/05/2009
Public Property Let IVModificaPagosFinanciamiento(ByVal value As Boolean)
    Mid$(mProps.Opcion, 171) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVModificaPagosFinanciamiento() As Boolean
    IVModificaPagosFinanciamiento = (Mid$(Opcion, 171, 1) = "S")
End Property

'jeaa 26/05/2009
Public Property Let IVVisualizaEntrada(ByVal value As Boolean)
    Mid$(mProps.Opcion, 172, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVVisualizaEntrada() As Boolean
    IVVisualizaEntrada = (Mid$(Opcion, 172, 1) = "S")
End Property

'jeaa 26/05/2009
Public Property Let IVVisualizaSegundaEntrada(ByVal value As Boolean)
    Mid$(mProps.Opcion, 173, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVVisualizaSegundaEntrada() As Boolean
    IVVisualizaSegundaEntrada = (Mid$(Opcion, 173, 1) = "S")
End Property

'jeaa 26/05/2009
Public Property Let IVCambiarTasa(ByVal value As Boolean)
    Mid$(mProps.Opcion, 174, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVCambiarTasa() As Boolean
    IVCambiarTasa = (Mid$(Opcion, 174, 1) = "S")
End Property

'*** jeaa 29/05/2009
Public Property Let IVPideGarante(ByVal value As Boolean)
    Mid$(mProps.Opcion, 175, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'*** jeaa 29/05/2009
Public Property Get IVPideGarante() As Boolean
    IVPideGarante = (Mid$(Opcion, 175, 1) = "S")
End Property

'*** jeaa 10/06/2009
Public Property Let GaranteVisible(ByVal value As Boolean)
    Mid$(mProps.Opcion, 176, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'*** jeaa 10/06/2009
Public Property Get GaranteVisible() As Boolean
    GaranteVisible = (Mid$(Opcion, 176, 1) = "S")
End Property

'*** jeaa 10/06/2009
Public Property Let IVControlaGaranteCreditos(ByVal value As Boolean)
    Mid$(mProps.Opcion, 177, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'*** jeaa 10/06/2009
Public Property Get IVControlaGaranteCreditos() As Boolean
    IVControlaGaranteCreditos = (Mid$(Opcion, 177, 1) = "S")
End Property

'*** jeaa 12/06/2009
Public Property Let IVFiltrarGrupoItem(ByVal value As Boolean)
    Mid$(mProps.Opcion, 178, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'*** jeaa 12/06/2009
Public Property Get IVFiltrarGrupoItem() As Boolean
    IVFiltrarGrupoItem = (Mid$(Opcion, 178, 1) = "S")
End Property

'*** jeaa 12/06/2009
Public Property Let IVVisualizarPesoNeto(ByVal value As Boolean)
    Mid$(mProps.Opcion, 179, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'*** jeaa 12/06/2009
Public Property Get IVVisualizarPesoNeto() As Boolean
    IVVisualizarPesoNeto = (Mid$(Opcion, 179, 1) = "S")
End Property

Public Property Let IVVisualizarPesoBruto(ByVal value As Boolean)
    Mid$(mProps.Opcion, 180, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'*** jeaa 12/06/2009
Public Property Get IVVisualizarPesoBruto() As Boolean
    IVVisualizarPesoBruto = (Mid$(Opcion, 180, 1) = "S")
End Property

Public Property Let IVVisualizarUnidadPeso(ByVal value As Boolean)
    Mid$(mProps.Opcion, 181, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'*** jeaa 12/06/2009
Public Property Get IVVisualizarUnidadPeso() As Boolean
    IVVisualizarUnidadPeso = (Mid$(Opcion, 181, 1) = "S")
End Property

'*** jeaa 21/07/2009
Public Property Let IVDatosAdicionales(ByVal value As Boolean)
    Mid$(mProps.Opcion, 182, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVDatosAdicionales() As Boolean
    IVDatosAdicionales = (Mid$(Opcion, 182, 1) = "S")
End Property

Public Property Let IVAutorizacionDescto(ByVal value As Boolean)
    Mid$(mProps.Opcion, 183, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVAutorizacionDescto() As Boolean
    IVAutorizacionDescto = (Mid$(Opcion, 183, 1) = "S")
End Property


Public Property Let IVVisualizaTrabajo(ByVal value As Boolean)
End Property

Public Property Get IVVisualizaTrabajo() As Boolean
End Property

Public Property Let IVTransCal(ByVal value As Boolean)
    Mid$(mProps.Opcion, 184, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVTransCal() As Boolean
    IVTransCal = (Mid$(Opcion, 184, 1) = "S")
End Property

Public Property Let IVTransProd(ByVal value As Boolean)
    Mid$(mProps.Opcion, 185, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVTransProd() As Boolean
    IVTransProd = (Mid$(Opcion, 185, 1) = "S")
End Property

Public Property Let IVTransIngCar(ByVal value As Boolean)
    Mid$(mProps.Opcion, 186, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVTransIngCar() As Boolean
    IVTransIngCar = (Mid$(Opcion, 186, 1) = "S")
End Property


Public Property Let IVVisualizaEstadoTrabajo(ByVal value As Boolean)
End Property

Public Property Get IVVisualizaEstadoTrabajo() As Boolean
End Property

Public Property Let IVSecuencialItem(ByVal value As Boolean)
End Property

Public Property Get IVSecuencialItem() As Boolean
End Property

Public Property Let IVControlaPrecioxFormaCobro(ByVal value As Boolean)
    Mid$(mProps.Opcion, 187, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVControlaPrecioxFormaCobro() As Boolean
    IVControlaPrecioxFormaCobro = (Mid$(Opcion, 187, 1) = "S")
End Property

Public Property Let IVSoloAdviertePendientes(ByVal value As Boolean)
    Mid$(mProps.Opcion, 188, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVSoloAdviertePendientes() As Boolean
    IVSoloAdviertePendientes = (Mid$(Opcion, 188, 1) = "S")
End Property

Public Property Let IVActualizaCostoUCenItems(ByVal value As Boolean)
    Mid$(mProps.Opcion, 189, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVActualizaCostoUCenItems() As Boolean
    IVActualizaCostoUCenItems = (Mid$(Opcion, 189, 1) = "S")
End Property

Public Property Let IVCambiarCostoPromxCostoUC(ByVal value As Boolean)
    Mid$(mProps.Opcion, 190, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVCambiarCostoPromxCostoUC() As Boolean
    IVCambiarCostoPromxCostoUC = (Mid$(Opcion, 190, 1) = "S")
End Property

'*********** ISOLLANTA
Public Property Let IVTipoIngresoISO(ByVal value As Boolean)
    Mid$(mProps.Opcion, 191, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVTipoIngresoISO() As Boolean
    IVTipoIngresoISO = (Mid$(Opcion, 191, 1) = "S")
End Property

Public Property Let IVBusqTiketISO(ByVal value As Boolean)
    Mid$(mProps.Opcion, 192, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVBusqTiketISO() As Boolean
    IVBusqTiketISO = (Mid$(Opcion, 192, 1) = "S")
End Property

Public Property Let IVMotReclamoISO(ByVal value As Boolean)
    Mid$(mProps.Opcion, 192, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVMotReclamoISO() As Boolean
    IVMotReclamoISO = (Mid$(Opcion, 192, 1) = "S")
End Property

Public Property Let IVPrecioxPCGrupo(ByVal value As Boolean)
    Mid$(mProps.Opcion, 193, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVPrecioxPCGrupo() As Boolean
    IVPrecioxPCGrupo = (Mid$(Opcion, 193, 1) = "S")
End Property

Public Property Let IVPrecioItemPromocion(ByVal value As Boolean)
    Mid$(mProps.Opcion, 194, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVPrecioItemPromocion() As Boolean
    IVPrecioItemPromocion = (Mid$(Opcion, 194, 1) = "S")
End Property

Public Property Get IVVisibleOfertas() As Boolean
    IVVisibleOfertas = (Mid$(Opcion, 195, 1) = "S")
End Property

Public Property Let IVVisibleOfertas(ByVal value As Boolean)
    Mid$(mProps.Opcion, 195, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVVisibleObra() As Boolean
    IVVisibleObra = (Mid$(Opcion, 196, 1) = "S")
End Property

Public Property Let IVVisibleObra(ByVal value As Boolean)
    Mid$(mProps.Opcion, 196, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'Auc Genera Transaccion para despedicio
Public Property Let IVGeneraTransParaDes(ByVal value As Boolean)
    Mid$(mProps.Opcion, 197, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVGeneraTransParaDes() As Boolean
    IVGeneraTransParaDes = (Mid$(Opcion, 197, 1) = "S")
End Property

Public Property Get IVVisibleOfertasDetalle() As Boolean
    IVVisibleOfertasDetalle = (Mid$(Opcion, 198, 1) = "S")
End Property

Public Property Let IVVisibleOfertasDetalle(ByVal value As Boolean)
    Mid$(mProps.Opcion, 198, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVVisibleRetencion() As Boolean
    IVVisibleRetencion = (Mid$(Opcion, 199, 1) = "S")
End Property

Public Property Let IVVisibleRetencion(ByVal value As Boolean)
    Mid$(mProps.Opcion, 199, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'jeaa 09/05/05 para verificacion de fechas
Public Property Let IVDesbloquearFechas(ByVal value As Boolean)
    Mid$(mProps.Opcion, 200, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVDesbloquearFechas() As Boolean
    IVDesbloquearFechas = (Mid$(Opcion, 200, 1) = "S")
End Property

Public Property Let UtilizarBodegaDestino(ByVal value As Boolean)
   Mid$(mProps.Opcion, 201, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get UtilizarBodegaDestino() As Boolean
    UtilizarBodegaDestino = (Mid$(Opcion, 201, 1) = "S")
End Property

Public Property Let ImportacionSoloSaldosInv(ByVal value As Boolean)
   Mid$(mProps.Opcion, 202, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get ImportacionSoloSaldosInv() As Boolean
    ImportacionSoloSaldosInv = (Mid$(Opcion, 202, 1) = "S")
End Property

Public Property Let HistPorTrans(ByVal value As Boolean)
    Mid$(mProps.Opcion, 203, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get HistPorTrans() As Boolean
    HistPorTrans = (Mid$(Opcion, 203, 1) = "S")
End Property

Public Property Let HistPorItem(ByVal value As Boolean)
   Mid$(mProps.Opcion, 204, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get HistPorItem() As Boolean
    HistPorItem = (Mid$(Opcion, 204, 1) = "S")
End Property

Public Property Get TotalizarAlImprimir() As Boolean
    TotalizarAlImprimir = (Mid$(Opcion, 205, 1) = "S")
End Property

Public Property Let TotalizarAlImprimir(ByVal value As Boolean)
   Mid$(mProps.Opcion, 205, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get CambiaFuenteSemi() As Boolean
    CambiaFuenteSemi = (Mid$(Opcion, 206, 1) = "S")
End Property
Public Property Let CambiaFuenteSemi(ByVal value As Boolean)
   Mid$(mProps.Opcion, 206, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'JEAA 13/09/2010
Public Property Let IVVendedorigualUsuario(ByVal value As Boolean)
        Mid$(mProps.Opcion, 207, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property

Public Property Get IVVendedorigualUsuario() As Boolean
    IVVendedorigualUsuario = (Mid$(Opcion, 207, 1) = "S")
End Property
'jeaa 22/09/2010
Public Property Let IVActualizaCostoRefItems(ByVal value As Boolean)
    Mid$(mProps.Opcion, 208, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVActualizaCostoRefItems() As Boolean
    IVActualizaCostoRefItems = (Mid$(Opcion, 208, 1) = "S")
End Property

'JEAA 13/09/2010
Public Property Let IVNoContinuaVendedor(ByVal value As Boolean)
        Mid$(mProps.Opcion, 209, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property

Public Property Get IVNoContinuaVendedor() As Boolean
    IVNoContinuaVendedor = (Mid$(Opcion, 209, 1) = "S")
End Property

'AUC 13/10/2010
Public Property Let IVTransPago(ByVal value As String)
    If Len(value) > 5 Then Err.Raise ERR_INVALIDO, "GNTrans.IVTransPago", MSGERR_INVALIDO
    Mid$(mProps.Opcion, 210, 5) = String(5, " ")
    Mid$(mProps.Opcion, 210, 5) = value
    mbooModificado = True
End Property

Public Property Get IVTransPago() As String
    IVTransPago = Trim(Mid$(Opcion, 210, 5))
End Property

Public Property Let IVGeneraPagoAuto(ByVal value As Boolean)
        Mid$(mProps.Opcion, 215, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property

Public Property Get IVGeneraPagoAuto() As Boolean
    IVGeneraPagoAuto = (Mid$(Opcion, 215, 1) = "S")
End Property

'AUC para transaccion de desperdicio
Public Property Let IVTransDes(ByVal value As String)
    If Len(value) > 5 Then Err.Raise ERR_INVALIDO, "GNTrans.IVTransDes", MSGERR_INVALIDO
    Mid$(mProps.Opcion, 216, 5) = String(5, " ")
    Mid$(mProps.Opcion, 216, 5) = value
    mbooModificado = True
End Property
Public Property Get IVTransDes() As String
    IVTransDes = Trim(Mid$(Opcion, 216, 5))
End Property


Public Property Let IVAplicaReceta(ByVal value As Boolean)
        Mid$(mProps.Opcion, 221, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property

Public Property Get IVAplicaReceta() As Boolean
    IVAplicaReceta = (Mid$(Opcion, 221, 1) = "S")
End Property

Public Property Let IVDetalleConsumoISO(ByVal value As Boolean)
    Mid$(mProps.Opcion, 222, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVDetalleConsumoISO() As Boolean
    IVDetalleConsumoISO = (Mid$(Opcion, 222, 1) = "S")
End Property

Public Property Let IVVerificarMontoCF(ByVal value As Boolean)
    Mid$(mProps.Opcion, 223, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVVerificarMontoCF() As Boolean
    IVVerificarMontoCF = (Mid$(Opcion, 223, 1) = "S")
End Property


Public Property Let IVGenerarXML(ByVal value As Boolean)
    Mid$(mProps.Opcion, 224, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVGenerarXML() As Boolean
    IVGenerarXML = (Mid$(Opcion, 224, 1) = "S")
End Property

Public Property Let AsientoTrans(ByVal value As String)
    If Len(value) > 5 Then Err.Raise ERR_INVALIDO, "GNTrans.TransAsiento", MSGERR_INVALIDO
    Mid$(mProps.Opcion, 225, 5) = String(5, " ")
    Mid$(mProps.Opcion, 225, 5) = value
    mbooModificado = True
End Property

Public Property Get AsientoTrans() As String
    AsientoTrans = Trim(Mid$(mProps.Opcion, 225, 5))
End Property

Public Property Let AsientoTransAnulada(ByVal value As String)
    If Len(value) > 5 Then Err.Raise ERR_INVALIDO, "GNTrans.TransAsiento", MSGERR_INVALIDO
    Mid$(mProps.Opcion, 230, 5) = String(5, " ")
    Mid$(mProps.Opcion, 230, 5) = value
    mbooModificado = True
End Property

Public Property Get AsientoTransAnulada() As String
    AsientoTransAnulada = Trim(Mid$(mProps.Opcion, 230, 5))
End Property

Public Property Let IVVerificaLimiteNoGrabarCosto(ByVal value As Boolean)
    Mid$(mProps.Opcion, 235, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVVerificaLimiteNoGrabarCosto() As Boolean
    IVVerificaLimiteNoGrabarCosto = (Mid$(Opcion, 235, 1) = "S")
End Property

Public Property Let IVReprocesaCosto(ByVal value As Boolean)
    Mid$(mProps.Opcion, 236, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVReprocesaCosto() As Boolean
    IVReprocesaCosto = (Mid$(Opcion, 236, 1) = "S")
End Property

Public Property Let IVImagenItem(ByVal value As Boolean)
   Mid$(mProps.Opcion, 237, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVImagenItem() As Boolean
    IVImagenItem = (Mid$(Opcion, 237, 1) = "S")
End Property

Public Property Let IVVisiblePresupuesto(ByVal value As Boolean)
    Mid$(mProps.Opcion, 238, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVVisiblePresupuesto() As Boolean
    
  If Len(mProps.objEmpresa.GNOpcion.ObtenerValor("Presupuesto")) > 0 Then
            If mProps.objEmpresa.GNOpcion.ObtenerValor("Presupuesto") <> "0" Then
                IVVisiblePresupuesto = (Mid$(Opcion, 238, 1) = "S")
            Else
                IVVisiblePresupuesto = False
            End If
    Else
        IVVisiblePresupuesto = False
    End If

End Property

Public Property Let IVPlazoProvCli(ByVal value As Boolean)
   Mid$(mProps.Opcion, 239, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVPlazoProvCli() As Boolean
    IVPlazoProvCli = (Mid$(Opcion, 239, 1) = "S")
End Property

Public Property Let IVBloquearItemsImportar(ByVal value As Boolean)
   Mid$(mProps.Opcion, 240, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVBloquearItemsImportar() As Boolean
    IVBloquearItemsImportar = (Mid$(Opcion, 240, 1) = "S")
End Property

Public Property Let IVValidarFechaValidez(ByVal value As Boolean)
   Mid$(mProps.Opcion, 241, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVValidarFechaValidez() As Boolean
    IVValidarFechaValidez = (Mid$(Opcion, 241, 1) = "S")
End Property

Public Property Get IVVisibleTransporteDetalle() As Boolean
    IVVisibleTransporteDetalle = (Mid$(Opcion, 242, 1) = "S")
End Property


Public Property Let IVVisibleTransporteDetalle(ByVal value As Boolean)
    Mid$(mProps.Opcion, 242, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get MostraNomAltPC() As Boolean
    MostraNomAltPC = (Mid$(Opcion, 243, 1) = "S")
End Property

Public Property Let MostraNomAltPC(ByVal value As Boolean)
    Mid$(mProps.Opcion, 243, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Let IVAvanceLineaCant(ByVal value As Boolean)
   Mid$(mProps.Opcion, 244, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVAvanceLineaCant() As Boolean
    IVAvanceLineaCant = (Mid$(Opcion, 244, 1) = "S")
End Property

Public Property Get IVVisibleSolicita() As Boolean
    IVVisibleSolicita = (Mid$(Opcion, 245, 1) = "S")
End Property

Public Property Let IVVisibleSolicita(ByVal value As Boolean)
    Mid$(mProps.Opcion, 245, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get VisualizaTransFuente() As Boolean
    VisualizaTransFuente = (Mid$(Opcion, 246, 1) = "S")
End Property

Public Property Let VisualizaTransFuente(ByVal value As Boolean)
    Mid$(mProps.Opcion, 246, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVBotonTerminarPedido() As Boolean
    IVBotonTerminarPedido = (Mid$(Opcion, 247, 1) = "S")
End Property

Public Property Let IVBotonTerminarPedido(ByVal value As Boolean)
    Mid$(mProps.Opcion, 247, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Let IVControlaLimitePorFacturaxLinea(ByVal value As Boolean)
        Mid$(mProps.Opcion, 248, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property

Public Property Get IVControlaLimitePorFacturaxLinea() As Boolean
    IVControlaLimitePorFacturaxLinea = (Mid$(Opcion, 248, 1) = "S")
End Property

Public Property Let IVOmitirIVARISE(ByVal value As Boolean)
    Mid$(mProps.Opcion, 249, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVOmitirIVARISE() As Boolean
    IVOmitirIVARISE = (Mid$(Opcion, 249, 1) = "S")
End Property

Public Property Let IVPideEmpleado(ByVal value As Boolean)
    Mid$(mProps.Opcion, 250, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVPideEmpleado() As Boolean
    IVPideEmpleado = (Mid$(Opcion, 250, 1) = "S")
End Property

Public Property Let EmpleadoVisible(ByVal value As Boolean)
    Mid$(mProps.Opcion, 251, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get EmpleadoVisible() As Boolean
    EmpleadoVisible = (Mid$(Opcion, 251, 1) = "S")
End Property

Public Property Get IVMostratDetalleItem() As Boolean
    IVMostratDetalleItem = (Mid$(Opcion, 252, 1) = "S")
End Property

Public Property Let IVMostratDetalleItem(ByVal value As Boolean)
    Mid$(mProps.Opcion, 252, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVGeneraPagosxPcGrupo() As Boolean
    IVGeneraPagosxPcGrupo = (Mid$(Opcion, 253, 1) = "S")
End Property

Public Property Let IVGeneraPagosxPcGrupo(ByVal value As Boolean)
    Mid$(mProps.Opcion, 253, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Let IVTipoTransCompromete(ByVal value As Boolean)
    Mid$(mProps.Opcion, 254, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVTipoTransCompromete() As Boolean
    IVTipoTransCompromete = (Mid$(Opcion, 254, 1) = "S")
End Property

Public Property Let IVSugerirPorcentajeDesct(ByVal value As Boolean)
    Mid$(mProps.Opcion, 255, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVSugerirPorcentajeDesct() As Boolean
    IVSugerirPorcentajeDesct = (Mid$(Opcion, 255, 1) = "S")
End Property

Public Property Let IVActualizaPrecioLista(ByVal value As Boolean)
    Mid$(mProps.Opcion, 256, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVActualizaPrecioLista() As Boolean
    IVActualizaPrecioLista = (Mid$(Opcion, 256, 1) = "S")
End Property




Public Property Get MargenMinimoUtil() As Currency
    MargenMinimoUtil = mProps.MargenMinimoUtil
End Property

Public Property Let MargenMinimoUtil(ByVal value As Currency)
    mProps.MargenMinimoUtil = value
    mbooModificado = True
End Property

Public Property Let IdGastoPre(ByVal value As Long)
    mProps.IdGastoPre = value
    mbooModificado = True
End Property

Public Property Get IdGastoPre() As Long
    IdGastoPre = mProps.IdGastoPre
End Property


Public Property Let CodGastoPre(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodGastoPre Then
        If Len(value) > 0 Then
            'Actualiza IdGasto también
            sql = "SELECT IdGasto FROM GNGasto WHERE CodGasto='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNTrans_CodGastoPre", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdGastoPre = rs!idGasto
            End If
            rs.Close
        Else
            IdGastoPre = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodGastoPre = value
    mbooModificado = True
End Property

Public Property Get CodGastoPre() As String
    Dim sql As String
    Dim rs As Recordset

    'Cuando pide por primera vez
    If (Len(mProps.CodGastoPre) = 0) And (mProps.IdGastoPre <> 0) Then
        'Obtiene CodGasto
        sql = "SELECT CodGasto FROM GNGasto WHERE IdGasto=" & mProps.IdGastoPre
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
        Else
            mProps.CodGastoPre = rs!CodGasto
        End If
        rs.Close
    End If

    Set rs = Nothing
    CodGastoPre = Trim$(mProps.CodGastoPre)
End Property

Public Property Let IdBodegaDestPre(ByVal value As Long)
    mProps.IdBodegaDestPre = value
    mbooModificado = True
End Property

Public Property Get IdBodegaDestPre() As Long
    IdBodegaDestPre = mProps.IdBodegaDestPre
End Property


Public Property Let CodBodegaDestPre(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodBodegaDestPre Then
        If Len(value) > 0 Then
            'Actualiza IdBodega también
            sql = "SELECT IdBodega FROM IVBodega WHERE CodBodega='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNTrans_CodBodegaDestPre", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdBodegaDestPre = rs!IdBodega
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdBodegaDestPre = 0
        End If
    End If

    mProps.CodBodegaDestPre = value
    mbooModificado = True
End Property

Public Property Get CodBodegaDestPre() As String
    CodBodegaDestPre = Trim$(mProps.CodBodegaDestPre)
End Property


Public Property Let TipoTrans(ByVal value As String)
    If Len(value) > 1 Then
        Err.Raise ERR_INVALIDO, "GNTrans_TipoTrans", MSGERR_INVALIDO
    End If
    
    mProps.TipoTrans = value
    mbooModificado = True
End Property

Public Property Get TipoTrans() As String
    TipoTrans = Trim$(mProps.TipoTrans)
End Property

Public Property Let IDSucursal(ByVal value As Long)
    mProps.IDSucursal = value
    mbooModificado = True
End Property

Public Property Get IDSucursal() As Long
    IDSucursal = mProps.IDSucursal
End Property


Public Property Let CodSucursal(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodSucursal Then
        If Len(value) > 0 Then
            'Actualiza IdResponsable también
            sql = "SELECT IdSucursal FROM GNSucursal WHERE CodSucursal='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNTrans_CodSucursal", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IDSucursal = rs!IDSucursal
            End If
            rs.Close
        Else
            IDSucursal = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodSucursal = value
    mbooModificado = True
End Property

Public Property Get CodSucursal() As String
    Dim sql As String
    Dim rs As Recordset

    'Cuando pide por primera vez
    If (Len(mProps.CodSucursal) = 0) And (mProps.IDSucursal <> 0) Then
        'Obtiene CodResponsable
        sql = "SELECT CodSucursal FROM GNSucursal WHERE IdSucursal=" & mProps.IDSucursal
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
            '*** No generamos error para que no se suspenda proceso en UI
'            Err.Raise ERR_NOHAYCODIGO, "GNTrans_IdSucursal", MSGERR_NOHAYCODIGO
'            Exit Property
        Else
            mProps.CodSucursal = rs!CodSucursal
        End If
        rs.Close
    End If

    Set rs = Nothing
    CodSucursal = Trim$(mProps.CodSucursal)
End Property

Public Property Let LimiteDias(ByVal value As Currency)
    mProps.LimiteDias = value
    mbooModificado = True
End Property

Public Property Get LimiteDias() As Currency
    LimiteDias = mProps.LimiteDias
End Property


Public Property Let NumIvGrupoFiltro(ByVal value As Long)
    mProps.NumIvGrupoFiltro = value
    mbooModificado = True
End Property

Public Property Get NumIvGrupoFiltro() As Long
    NumIvGrupoFiltro = mProps.NumIvGrupoFiltro
End Property

Public Property Let IdGruposFiltro(ByVal value As String)
    mProps.IdGruposFiltro = value
    mbooModificado = True
End Property

Public Property Get IdGruposFiltro() As String
    IdGruposFiltro = mProps.IdGruposFiltro
End Property

Public Property Let NumIvGrupoTrabajo(ByVal value As Long)
    mProps.NumIvGrupoTrabajo = value
    mbooModificado = True
End Property

Public Property Get NumIvGrupoTrabajo() As Long
    NumIvGrupoTrabajo = mProps.NumIvGrupoTrabajo
End Property

Public Property Let IdGruposTrabajo(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNTrans.IdGruposTrabajo", MSGERR_INVALIDO
    End If
    
    mProps.IdGruposTrabajo = value
    mbooModificado = True
End Property

Public Property Get IdGruposTrabajo() As String
    IdGruposTrabajo = mProps.IdGruposTrabajo
End Property
'/**** error*///////
Public Property Let NumIvGrupoTamaño(ByVal value As Long)
End Property

Public Property Get NumIvGrupoTamaño() As Long
End Property

Public Property Let IdGruposTamaño(ByVal value As String)
End Property

Public Property Get IdGruposTamaño() As String
End Property

Public Property Let NumIvGrupoMarca(ByVal value As Long)
    mProps.NumIvGrupoMarca = value
    mbooModificado = True
End Property

Public Property Get NumIvGrupoMarca() As Long
    NumIvGrupoMarca = mProps.NumIvGrupoMarca
End Property

Public Property Let IdGruposMarca(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNTrans.IdGruposMarca", MSGERR_INVALIDO
    End If
    
    mProps.IdGruposMarca = value
    mbooModificado = True
End Property

Public Property Get IdGruposMarca() As String
    IdGruposMarca = mProps.IdGruposMarca
End Property

Public Property Let RutaLibImp(ByVal value As String)
    mProps.RutaLibImp = value
    mbooModificado = True
End Property

Public Property Get RutaLibImp() As String
    RutaLibImp = mProps.RutaLibImp
End Property

Public Property Let BandImpGlobal(ByVal value As Boolean)
    mProps.BandImpGlobal = value
    mbooModificado = True
End Property

Public Property Get BandImpGlobal() As Boolean
    BandImpGlobal = mProps.BandImpGlobal
End Property

Public Property Let CodBodegaDesPre(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodBodegaDesPre Then
        If Len(value) > 0 Then
            'Actualiza IdBodega también
            sql = "SELECT IdBodega FROM IVBodega WHERE CodBodega='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNTrans_CodBodegaDesPre", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdBodegaDesPre = rs!IdBodega
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdBodegaDesPre = 0
        End If
    End If
    mProps.CodBodegaDesPre = value
    mbooModificado = True
End Property

Public Property Get CodBodegaDesPre() As String
    CodBodegaDesPre = Trim$(mProps.CodBodegaDesPre)
End Property

Public Property Let IdBodegaDesPre(ByVal value As Long)
    mProps.IdBodegaDesPre = value
    mbooModificado = True
End Property

Public Property Get IdBodegaDesPre() As Long
    IdBodegaDesPre = mProps.IdBodegaDesPre
End Property

Public Property Let BandNuevoAutoImpresor(ByVal value As Boolean)
    mProps.BandNuevoAutoImpresor = value
    mbooModificado = True
End Property

Public Property Get BandNuevoAutoImpresor() As Boolean
    BandNuevoAutoImpresor = mProps.BandNuevoAutoImpresor
End Property

Public Function AddAsientoPresup() As Long
    Dim nuevo As GNTransAsientoPresup
    Set nuevo = New GNTransAsientoPresup
    Set nuevo.GNTrans = Me
    mcolAsientosPresup.Add nuevo
    AddAsientoPresup = mcolAsientosPresup.Count    'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function

Public Sub RemoveAsientoPresup(ix As Variant)
    Dim obj As GNTransAsientoPresup, i As Long
    'Si llama con numero de indice del objeto
    If (VarType(ix) = vbLong) Or (VarType(ix) = vbInteger) Then
        mcolAsientosPresup.Remove ix      'Elimina el objeto
    'Si llama con la referencia al objeto que quiere eliminar
    Else
        'Busca el objeto dentro de la colección
        For i = mcolAsientosPresup.Count To 1 Step -1
            Set obj = mcolAsientosPresup.item(i)
            If obj Is ix Then       'Si encuentra
                mcolAsientosPresup.Remove i       'Elimina el objeto
                GoTo eliminado
            End If
        Next i
        'Si no encuentra el objeto, genera error
        Err.Raise ERR_INVALIDO, "GNTrans_RemoveAsientoPresup", "No se puede eliminar el objeto."
    End If
eliminado:
    mbooModificado = True
    Set obj = Nothing
End Sub

Public Property Get CountAsientoPresup() As Long
    CountAsientoPresup = mcolAsientosPresup.Count
End Property

Public Property Get AsientosPresup(ix) As GNTransAsientoPresup
    Set AsientosPresup = mcolAsientosPresup.item(ix)
End Property

    
Public Property Let AnexoSustento(ByVal value As String)
    If Len(value) > 2 Then
        Err.Raise ERR_INVALIDO, "GNTrans_AnexoSustento", MSGERR_INVALIDO
    End If
    
    mProps.AnexoSustento = value
    mbooModificado = True
End Property

Public Property Get AnexoSustento() As String
    AnexoSustento = Trim$(mProps.AnexoSustento)
End Property

Public Property Let ImportaPRCTD(ByVal value As Boolean)
    If value Then
        mProps.BandImportacion = mProps.BandImportacion Or &H100
    Else
        mProps.BandImportacion = mProps.BandImportacion And &HBF
    End If
    mbooModificado = True
End Property

Public Property Get ImportaPRCTD() As Boolean
    ImportaPRCTD = CBool(mProps.BandImportacion And &H100)
End Property

Public Property Let ImportaIVKPlista(ByVal value As Boolean)
End Property

Public Property Get ImportaIVKPlista() As Boolean
End Property

Public Property Let IdInventarioPre(ByVal value As Long)
    mProps.IdInventarioPre = value
    mbooModificado = True
End Property

Public Property Get IdInventarioPre() As Long
    IdInventarioPre = mProps.IdInventarioPre
End Property

Public Property Let CodInventarioPre(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodInventarioPre Then
        If Len(value) > 0 Then
            'Actualiza IdProvCli también
            sql = "SELECT IdInventario FROM ivinventario WHERE Codinventario='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNTrans_CodInventarioPre", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdInventarioPre = rs!IdInventario
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdInventarioPre = 0
        End If
    End If
    mProps.CodInventarioPre = value
    mbooModificado = True
End Property

Public Property Get CodInventarioPre() As String
    CodInventarioPre = Trim$(mProps.CodInventarioPre)
End Property

Public Property Let AsientoGastoxZona(ByVal value As Boolean)
    mProps.AsientoGastoxZona = value
    mbooModificado = True
End Property

Public Property Get AsientoGastoxZona() As Boolean
    AsientoGastoxZona = mProps.AsientoGastoxZona
End Property



'*----------------------------  GNOPCION2


Public Property Let Opcion2(ByVal value As String)
    If Len(value) > 120 Then
        Err.Raise ERR_INVALIDO, "GNTrans_Opcion2", MSGERR_INVALIDO
    End If
    mProps.Opcion2 = value
    mbooModificado = True
End Property

Public Property Get Opcion2() As String
    Opcion2 = Trim$(mProps.Opcion2)
End Property

Public Property Let IVVisualizaCodAlterno(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 1, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVVisualizaCodAlterno() As Boolean
    IVVisualizaCodAlterno = (Mid$(Opcion2, 1, 1) = "S")
End Property

Public Property Let IVTotItemSoloSignosDif(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 2, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVTotItemSoloSignosDif() As Boolean
    IVTotItemSoloSignosDif = (Mid$(Opcion2, 2, 1) = "S")
End Property

'AUC LIRA
Public Property Let IVMostrarSoloItemMatriz(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 3, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVMostrarSoloItemMatriz() As Boolean
    IVMostrarSoloItemMatriz = (Mid$(Opcion2, 3, 1) = "S")
End Property

Public Property Get IVVisibleFormaPago() As Boolean
    IVVisibleFormaPago = (Mid$(Opcion2, 4, 1) = "S")
End Property

Public Property Let IVVisibleFormaPago(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 4, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVVisibleSucursalSolicita() As Boolean
    IVVisibleSucursalSolicita = (Mid$(Opcion2, 5, 1) = "S")
End Property

Public Property Let IVVisibleSucursalSolicita(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 5, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Let IVSolicitaCambioPrecio(ByVal value As Boolean)
    Mid$(mProps.Opcion, 6, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVSolicitaCambioPrecio() As Boolean
    IVSolicitaCambioPrecio = (Mid$(Opcion, 6, 1) = "S")
End Property

'AUC UNIDO PARA MEGA
Public Property Get IVSeleccionarFormaCobroInicio() As Boolean
    IVSeleccionarFormaCobroInicio = (Mid$(Opcion2, 7, 1) = "S")
End Property

Public Property Let IVSeleccionarFormaCobroInicio(ByVal value As Boolean)
   Mid$(mProps.Opcion2, 7, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property
'AUC MARZO 2012

Public Property Get IVFiltroGrupoAlFact() As Boolean
    IVFiltroGrupoAlFact = (Mid$(Opcion2, 8, 1) = "S")
End Property

Public Property Let IVFiltroGrupoAlFact(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 8, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVNoEstadoFinalizar() As Boolean
    IVNoEstadoFinalizar = (Mid$(Opcion2, 9, 1) = "S")
End Property

Public Property Let IVNoEstadoFinalizar(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 9, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVCambiarEstadoFinalizar() As Boolean
    IVCambiarEstadoFinalizar = (Mid$(Opcion2, 10, 1) = "S")
End Property

Public Property Let IVCambiarEstadoFinalizar(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 10, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Let OmitirGastoxZona(ByVal value As Boolean)
    mProps.OmitirGastoxZona = value
    mbooModificado = True
End Property

Public Property Get OmitirGastoxZona() As Boolean
    OmitirGastoxZona = mProps.OmitirGastoxZona
End Property

Public Property Get IVClaveEliminarItem() As Boolean
    IVClaveEliminarItem = (Mid$(Opcion2, 11, 1) = "S")
End Property

Public Property Let IVClaveEliminarItem(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 11, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVUtilizarFiltroIvGDiasCred() As Boolean
    IVUtilizarFiltroIvGDiasCred = (Mid$(Opcion2, 12, 1) = "S")
End Property

Public Property Let IVUtilizarFiltroIvGDiasCred(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 12, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property


Public Property Get IVclaveModPagos() As Boolean
    IVclaveModPagos = (Mid$(Opcion2, 13, 1) = "S")
End Property

Public Property Let IVclaveModPagos(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 13, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property
'AUC
Public Property Let IVVerificaDisponible(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 14, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property
Public Property Get IVVerificaDisponible() As Boolean
    IVVerificaDisponible = (Mid$(Opcion2, 14, 1) = "S")
End Property

Public Property Get IVGeneraPagosxCliente() As Boolean
    IVGeneraPagosxCliente = (Mid$(Opcion2, 15, 1) = "S")
End Property

Public Property Let IVGeneraPagosxCliente(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 15, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

'PROCESOS VER SI SE UTILIZA SI BORRAR
Public Property Let VisualizaFechaEntrega(ByVal value As Boolean)
       Mid$(mProps.Opcion2, 16, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property
Public Property Get VisualizaFechaEntrega() As Boolean
    VisualizaFechaEntrega = (Mid$(Opcion2, 16, 1) = "S")
End Property
'AUC para importacion solo de  materiales modulo de produccion
Public Property Let ImportaSoloMateriales(ByVal value As Boolean)
       Mid$(mProps.Opcion2, 17, 1) = IIf(value, "S", "N")
       mbooModificado = True
End Property
Public Property Get ImportaSoloMateriales() As Boolean
    ImportaSoloMateriales = (Mid$(Opcion2, 17, 1) = "S")
End Property

Public Property Let IVValidarBodegaPre(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 18, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVValidarBodegaPre() As Boolean
    IVValidarBodegaPre = (Mid$(Opcion2, 18, 1) = "S")
End Property

Public Property Let CentroCostoDetalleVisible(ByVal value As Boolean)
   Mid$(mProps.Opcion2, 19, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property
Public Property Get CentroCostoDetalleVisible() As Boolean
    CentroCostoDetalleVisible = (Mid$(Opcion2, 19, 1) = "S")
End Property

Public Property Let ImportaPCKCHP(ByVal value As Boolean)
    If value Then
        mProps.BandImportacion = mProps.BandImportacion Or &H1
    Else
        mProps.BandImportacion = mProps.BandImportacion And &HFE
    End If
    mbooModificado = True
End Property

Public Property Get ImportaPCKCHP() As Boolean
    ImportaPCKCHP = CBool(mProps.BandImportacion And &H1)
End Property

Public Property Let IVFechaActual(ByVal value As Boolean)
   Mid$(mProps.Opcion2, 20, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVFechaActual() As Boolean
    IVFechaActual = (Mid$(Opcion2, 20, 1) = "S")
End Property

Public Property Let IVBloquearVendedorCliente(ByVal value As Boolean)
   Mid$(mProps.Opcion2, 21, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVBloquearVendedorCliente() As Boolean
    IVBloquearVendedorCliente = (Mid$(Opcion2, 21, 1) = "S")
End Property

Public Property Let IVOmitirPorcentajeDesctxMora(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 22, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVOmitirPorcentajeDesctxMora() As Boolean
    IVOmitirPorcentajeDesctxMora = (Mid$(Opcion2, 22, 1) = "S")
End Property
'lira
Public Property Let IVModificaPcGrupo(ByVal value As Boolean)
  Mid$(mProps.Opcion2, 23, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property
Public Property Get IVModificaPcGrupo() As Boolean
    IVModificaPcGrupo = (Mid$(Opcion2, 23, 1) = "S")
End Property

Public Property Let IVBusqueda2Bodegas(ByVal value As Boolean)
  Mid$(mProps.Opcion2, 24, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVBusqueda2Bodegas() As Boolean
    IVBusqueda2Bodegas = (Mid$(Opcion2, 24, 1) = "S")
End Property

Public Property Let IdBodegaFiltro(ByVal value As String)
    mProps.IdBodegaFiltro = value
    mbooModificado = True
End Property

Public Property Get IdBodegaFiltro() As String
    IdBodegaFiltro = mProps.IdBodegaFiltro
End Property

Public Property Let IVMostarPantTotalExist(ByVal value As Boolean)
  Mid$(mProps.Opcion2, 25, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVMostarPantTotalExist() As Boolean
    IVMostarPantTotalExist = (Mid$(Opcion2, 25, 1) = "S")
End Property

Public Property Let IVCCIgualTrans(ByVal value As Boolean)
  Mid$(mProps.Opcion2, 26, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVCCIgualTrans() As Boolean
    IVCCIgualTrans = (Mid$(Opcion2, 26, 1) = "S")
End Property


Public Property Let IVCostoAdicionalPrecio(ByVal value As Boolean)
  Mid$(mProps.Opcion2, 27, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVCostoAdicionalPrecio() As Boolean
    IVCostoAdicionalPrecio = (Mid$(Opcion2, 27, 1) = "S")
End Property

Public Property Let IVBusqueda2Precios(ByVal value As Boolean)
  Mid$(mProps.Opcion2, 28, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVBusqueda2Precios() As Boolean
    IVBusqueda2Precios = (Mid$(Opcion2, 28, 1) = "S")
End Property

Public Property Let IdPrecioFiltro(ByVal value As String)
    mProps.IdPrecioFiltro = value
    mbooModificado = True
End Property

Public Property Get IdPrecioFiltro() As String
    IdPrecioFiltro = mProps.IdPrecioFiltro
End Property

Public Property Let IVActualizaCostoUCenItemsMod(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 29, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVActualizaCostoUCenItemsMod() As Boolean
    IVActualizaCostoUCenItemsMod = (Mid$(Opcion2, 29, 1) = "S")
End Property

Public Property Let IVMuestarFotoIVG1(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 30, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVMuestarFotoIVG1() As Boolean
    IVMuestarFotoIVG1 = (Mid$(Opcion2, 30, 1) = "S")
End Property

Public Property Let IVAnchoFila(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 31, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVAnchoFila() As Boolean
    IVAnchoFila = (Mid$(Opcion2, 31, 1) = "S")
End Property

Public Property Let IVValorAnchoFila(ByVal value As Integer)
    If value > 9999 Then
        Err.Raise ERR_INVALIDO, "GNTrans.IVValorAnchoFila ", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion2, 32, 4) = Format(value, "0000")
    mbooModificado = True
End Property

Public Property Get IVValorAnchoFila() As Integer
    IVValorAnchoFila = Val(Mid$(Opcion2, 32, 4))
End Property

Public Property Let IVTipoLetraItems(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 36, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVTipoLetraItems() As Boolean
    IVTipoLetraItems = (Mid$(Opcion2, 36, 1) = "S")
End Property

Public Property Let NombreFuente(ByVal value As String)
    mProps.NombreFuente = value
    mbooModificado = True
End Property

Public Property Get NombreFuente() As String
    NombreFuente = mProps.NombreFuente
End Property

Public Property Let TamanioFuente(ByVal value As Integer)
    mProps.TamanioFuente = value
    mbooModificado = True
End Property

Public Property Get TamanioFuente() As Integer
    TamanioFuente = mProps.TamanioFuente
End Property


Public Property Let IVFiltroTipoItem(ByVal value As Boolean)
 Mid$(mProps.Opcion2, 37, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVFiltroTipoItem() As Boolean
    IVFiltroTipoItem = (Mid$(Opcion2, 37, 1) = "S")
End Property

Public Property Let IVFiltroNumTipoItem(ByVal value As String)
  Mid$(mProps.Opcion2, 38, 1) = value
   mbooModificado = True
End Property

Public Property Get IVFiltroNumTipoItem() As String
    IVFiltroNumTipoItem = Mid$(Opcion2, 38, 1)
End Property


Public Property Let AfectaCantidadDocumento(ByVal value As Boolean)
    mProps.AfectaCantidadDocumento = value
    mbooModificado = True
End Property

Public Property Get AfectaCantidadDocumento() As Boolean
    AfectaCantidadDocumento = mProps.AfectaCantidadDocumento
End Property

Public Property Let FormaPagoSRI(ByVal value As String)
    If Len(value) > 2 Then
        Err.Raise ERR_INVALIDO, "GNTrans_FormaPagoSRI", MSGERR_INVALIDO
    End If
    
    mProps.FormaPagoSRI = value
    mbooModificado = True
End Property

Public Property Get FormaPagoSRI() As String
    FormaPagoSRI = Trim$(mProps.FormaPagoSRI)
End Property

Public Property Let IVObservaCambioPrecio(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 39, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVObservaCambioPrecio() As Boolean
    IVObservaCambioPrecio = (Mid$(Opcion2, 39, 1) = "S")
End Property

Public Property Let IVModificarPrecioIVA(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 40, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVModificarPrecioIVA() As Boolean
    IVModificarPrecioIVA = (Mid$(Opcion2, 40, 1) = "S")
End Property

Public Property Let IVMaxDescuentoPTIVA(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 41, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVMaxDescuentoPTIVA() As Boolean
    IVMaxDescuentoPTIVA = (Mid$(Opcion2, 41, 1) = "S")
End Property


Public Property Let IVValorPorMaxDesc(ByVal value As Currency)
    If value > 100 Then
        Err.Raise ERR_INVALIDO, "GNTrans.IVValorPorMaxDesc  ", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion2, 42, 5) = Format(value, "00.00")
    mbooModificado = True
End Property

Public Property Get IVValorPorMaxDesc() As Currency
    IVValorPorMaxDesc = Val(Mid$(Opcion2, 42, 5))
End Property

Public Property Let IVDatosExportacion(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 47, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVDatosExportacion() As Boolean
    IVDatosExportacion = (Mid$(Opcion2, 47, 1) = "S")
End Property

Public Property Let CerrarFichaAlGrabar(ByVal value As Boolean)
        Mid$(mProps.Opcion2, 48, 1) = IIf(value, "S", "N")
        mbooModificado = True
End Property

Public Property Get CerrarFichaAlGrabar() As Boolean
    CerrarFichaAlGrabar = (Mid$(Opcion2, 48, 1) = "S")
End Property

Public Property Let IVPCGrupoGeneraPagosxCliente(ByVal value As String)
        Mid$(mProps.Opcion2, 49, 1) = value
        mbooModificado = True
End Property

Public Property Get IVPCGrupoGeneraPagosxCliente() As String
    IVPCGrupoGeneraPagosxCliente = Mid$(Opcion2, 49, 1)
End Property

Public Property Let UtilizaPreImpreso(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 50, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get UtilizaPreImpreso() As Boolean
    UtilizaPreImpreso = (Mid$(Opcion2, 50, 1) = "S")
End Property

'----------------
'AUC OPCION PARA CERRAR TODOS LOS USUARIOS
Public Property Let CierreTodosU(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 51, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get CierreTodosU() As Boolean
    CierreTodosU = (Mid$(Opcion2, 51, 1) = "S")
End Property

'importa solo las que tienen solicitud de credito aprobada
Public Property Let ImpTransSolCreApr(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 52, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get ImpTransSolCreApr() As Boolean
    ImpTransSolCreApr = (Mid$(Opcion2, 52, 1) = "S")
End Property

Public Property Let IVTransAutoComision(ByVal value As String)
    If Len(value) > 5 Then Err.Raise ERR_INVALIDO, "GNTrans.IVTransAutoComision", MSGERR_INVALIDO
    Mid$(mProps.Opcion2, 53, 5) = String(5, " ")
    Mid$(mProps.Opcion2, 53, 5) = value
    mbooModificado = True
End Property

Public Property Get IVTransAutoComision() As String
    IVTransAutoComision = Trim(Mid$(Opcion2, 53, 5))
End Property

Public Property Let IVItemComision(ByVal value As String)
    If Len(value) > 5 Then Err.Raise ERR_INVALIDO, "GNTrans.IVItemComision", MSGERR_INVALIDO
    Mid$(mProps.Opcion2, 58, 5) = String(5, " ")
    Mid$(mProps.Opcion2, 58, 5) = value
    mbooModificado = True
End Property

Public Property Get IVItemComision() As String
    IVItemComision = Trim(Mid$(Opcion2, 58, 5))
End Property

Public Property Let IVClienteComision(ByVal value As String)
    If Len(value) > 5 Then Err.Raise ERR_INVALIDO, "GNTrans.IVClienteComision", MSGERR_INVALIDO
    Mid$(mProps.Opcion2, 63, 5) = String(5, " ")
    Mid$(mProps.Opcion2, 63, 5) = value
    mbooModificado = True
End Property

Public Property Get IVClienteComision() As String
    IVClienteComision = Trim(Mid$(Opcion2, 63, 5))
End Property

Public Property Let IVComision(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 64, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVComision() As Boolean
    IVComision = (Mid$(Opcion2, 64, 1) = "S")
End Property
'Trasaccion para salida de consignacion

Public Property Let IVTransConsignacion(ByVal value As String)
    If Len(value) > 5 Then Err.Raise ERR_INVALIDO, "GNTrans.IVTransConsignacion", MSGERR_INVALIDO
    Mid$(mProps.Opcion2, 65, 5) = String(5, " ")
    Mid$(mProps.Opcion2, 65, 5) = value
    mbooModificado = True
End Property

Public Property Get IVTransConsignacion() As String
    IVTransConsignacion = Trim(Mid$(Opcion2, 65, 5))
End Property

Public Property Get IVTransCompra() As String
    IVTransCompra = Trim(Mid$(Opcion2, 70, 5))
End Property

'Trasaccion para salida de consignacion
Public Property Let IVTransCompra(ByVal value As String)
    If Len(value) > 5 Then Err.Raise ERR_INVALIDO, "GNTrans.IVTransConsignacion", MSGERR_INVALIDO
    Mid$(mProps.Opcion2, 70, 5) = String(5, " ")
    Mid$(mProps.Opcion2, 70, 5) = value
    mbooModificado = True
End Property

Public Property Let TSTransAutoDescMora(ByVal value As String)
    If Len(value) > 5 Then Err.Raise ERR_INVALIDO, "GNTrans.TSTransAutoDescMora", MSGERR_INVALIDO
    Mid$(mProps.Opcion2, 75, 5) = String(5, " ")
    Mid$(mProps.Opcion2, 75, 5) = value
    mbooModificado = True
End Property

Public Property Get TSTransAutoDescMora() As String
    TSTransAutoDescMora = Trim(Mid$(Opcion2, 75, 5))
End Property
'AUC 08/01/2011
Public Property Let TipoInteres(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 80, 1) = IIf(value, "S", "N")
   mbooModificado = True
End Property
Public Property Get TipoInteres() As Boolean
    TipoInteres = (Mid$(Opcion2, 80, 1) = "S")
End Property

Public Property Let IVDatosDinardap(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 81, 1) = IIf(value, "S", "N")
   mbooModificado = True
End Property

Public Property Get IVDatosDinardap() As Boolean
    IVDatosDinardap = (Mid$(Opcion2, 81, 1) = "S")
End Property

Public Property Let IVSerCorte(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 82, 1) = IIf(value, "S", "N")
   mbooModificado = True
End Property

Public Property Get IVSerCorte() As Boolean
    IVSerCorte = (Mid$(Opcion2, 82, 1) = "S")
End Property

Public Property Let IVCompraAuto(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 83, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVCompraAuto() As Boolean
    IVCompraAuto = (Mid$(Opcion2, 83, 1) = "S")
End Property

Public Property Let IVMostrarSolCre(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 84, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVMostrarSolCre() As Boolean
    IVMostrarSolCre = (Mid$(Opcion2, 84, 1) = "S")
End Property

Public Property Let IVVehSegDis(ByVal value As String)
        If value <> "N" And value <> "V" And value <> "S" And value <> "D" Then
        Err.Raise ERR_INVALIDO, "GNTrans_IVVehSegDis", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion2, 85, 1) = value
    mbooModificado = True
End Property

Public Property Get IVVehSegDis() As String
    IVVehSegDis = Mid$(Opcion2, 85, 1)
End Property

Public Property Let IVNewOrdenCol(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 86, 1) = IIf(value, "S", "N")
   mbooModificado = True
End Property

Public Property Get IVNewOrdenCol() As Boolean
    IVNewOrdenCol = (Mid$(Opcion2, 86, 1) = "S")
End Property

Public Property Let ColVisibleNew(ByVal value As Long)
    mProps.ColVisibleNew = value
   mbooModificado = True
End Property

Public Property Get ColVisibleNew() As Long
    ColVisibleNew = mProps.ColVisibleNew
End Property

Public Property Let ColEditableNew(ByVal value As Long)
    mProps.ColEditableNew = value
   mbooModificado = True
End Property

Public Property Get ColEditableNew() As Long
    ColEditableNew = mProps.ColEditableNew
End Property

Public Property Let ColSeleccionableNew(ByVal value As Long)
    mProps.ColSeleccionableNew = value
    mbooModificado = True
End Property

Public Property Get ColSeleccionableNew() As Long
    ColSeleccionableNew = mProps.ColSeleccionableNew
End Property

Public Property Let IVMostrarIntevaloDias(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 87, 1) = IIf(value, "S", "N")
   mbooModificado = True
End Property

Public Property Get IVMostrarIntevaloDias() As Boolean
    IVMostrarIntevaloDias = (Mid$(Opcion2, 87, 1) = "S")
End Property

Public Property Let IVDescXCantMinima(ByVal value As Boolean)
'    Mid$(mProps.Opcion2, 88, 1) = IIf(value, "S", "N")
'   mbooModificado = True
End Property

Public Property Get IVDescXCantMinima() As Boolean
'    IVDescXCantMinima = (Mid$(Opcion2, 88, 1) = "S")
End Property

Public Property Let IVGeneraNumSerie(ByVal value As Boolean)
   Mid$(mProps.Opcion2, 88, 1) = IIf(value, "S", "N")
   mbooModificado = True
End Property
Public Property Get IVGeneraNumSerie() As Boolean
    IVGeneraNumSerie = (Mid$(Opcion2, 88, 1) = "S")
End Property

Public Property Let IVValTipoDocTarjeta(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 89, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVValTipoDocTarjeta() As Boolean
    IVValTipoDocTarjeta = (Mid$(Opcion2, 89, 1) = "S")
End Property

Public Property Let IVPrecioTarjetaPre(ByVal value As Integer)
    If value > 7 Or value < 0 Then
        Err.Raise ERR_INVALIDO, "GNTrans.IVPrecioTarjetaPre", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion2, 90, 1) = value
    mbooModificado = True
End Property

Public Property Get IVPrecioTarjetaPre() As Integer
    IVPrecioTarjetaPre = Val(Mid$(Opcion2, 90, 1))
End Property

Public Property Let IVCobroTarjeta(ByVal value As String)
    If Len(value) > 5 Then
        Err.Raise ERR_INVALIDO, "GNTrans.IVCobroTarjeta", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion2, 91, 5) = String(5, " ")
    Mid$(mProps.Opcion2, 91, 5) = value
    mbooModificado = True
End Property

Public Property Get IVCobroTarjeta() As String
    IVCobroTarjeta = Trim$(Mid$(mProps.Opcion2, 91, 5))
End Property

'-------------
Public Property Let IVCuotasSeparado(ByVal value As Boolean)
  Mid$(mProps.Opcion2, 96, 1) = IIf(value, "S", "N")
   mbooModificado = True
End Property

Public Property Get IVCuotasSeparado() As Boolean
    IVCuotasSeparado = (Mid$(Opcion2, 96, 1) = "S")
End Property


Public Property Let TSCobroCapIntPorLinea(ByVal value As Boolean)
   Mid$(mProps.Opcion2, 97, 1) = IIf(value, "S", "N")
   mbooModificado = True
End Property

Public Property Get TSCobroCapIntPorLinea() As Boolean
    TSCobroCapIntPorLinea = (Mid$(Opcion2, 97, 1) = "S")
End Property



Public Property Let IVTransNumSerie(ByVal value As String)
    If Len(value) > 5 Then Err.Raise ERR_INVALIDO, "GNTrans.IVTransNumSerie", MSGERR_INVALIDO
    Mid$(mProps.Opcion2, 98, 5) = String(5, " ")
    Mid$(mProps.Opcion2, 98, 5) = value
    mbooModificado = True
End Property

Public Property Get IVTransNumSerie() As String
    IVTransNumSerie = Trim(Mid$(Opcion2, 98, 5))
End Property

Public Property Let IVLineaProduccion(ByVal value As Boolean)
   Mid$(mProps.Opcion2, 103, 1) = IIf(value, "S", "N")
   mbooModificado = True
End Property

Public Property Get IVLineaProduccion() As Boolean
    IVLineaProduccion = (Mid$(Opcion2, 103, 1) = "S")
End Property
Public Property Let IVDivideTrans(ByVal value As Boolean)
   Mid$(mProps.Opcion2, 104, 1) = IIf(value, "S", "N")
   mbooModificado = True
End Property

Public Property Get IVDivideTrans() As Boolean
    IVDivideTrans = (Mid$(Opcion2, 104, 1) = "S")
End Property

Public Property Let IVValorDivideTrans(ByVal value As Integer)
    If value > 99 Then
        Err.Raise ERR_INVALIDO, "GNTrans.IVValorDivideTrans", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion2, 105, 2) = Format(value, "00")
    mbooModificado = True
End Property

Public Property Get IVValorDivideTrans() As Integer
    IVValorDivideTrans = Val(Mid$(Opcion2, 105, 2))
End Property
Public Property Let IVMostrarPrecioEnc(ByVal value As Boolean)
   Mid$(mProps.Opcion2, 107, 1) = IIf(value, "S", "N")
   mbooModificado = True
End Property

Public Property Get IVMostrarPrecioEnc() As Boolean
    IVMostrarPrecioEnc = (Mid$(Opcion2, 107, 1) = "S")
End Property

Public Property Let IVPrecioEnc(ByVal value As Integer)
    If value > 7 Or value < 0 Then
        Err.Raise ERR_INVALIDO, "GNTrans.IVPrecioEnn", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion2, 108, 1) = value
    mbooModificado = True
End Property

Public Property Get IVPrecioEnc() As Integer
    IVPrecioEnc = Val(Mid$(Opcion2, 108, 1))
End Property

Public Property Let IVAplicaPrecioDist(ByVal value As Boolean)
       Mid$(mProps.Opcion2, 109, 1) = IIf(value, "S", "N")
   mbooModificado = True
End Property

Public Property Get IVAplicaPrecioDist() As Boolean
    IVAplicaPrecioDist = (Mid$(Opcion2, 109, 1) = "S")
End Property

'AUC para solis
Public Property Let CierrePorTrans(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 110, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get CierrePorTrans() As Boolean
    CierrePorTrans = (Mid$(Opcion2, 110, 1) = "S")
End Property


Public Property Let IVAplicaCostoxProv(ByVal value As Boolean)
       Mid$(mProps.Opcion2, 116, 1) = IIf(value, "S", "N")
   mbooModificado = True
End Property

Public Property Get IVAplicaCostoxProv() As Boolean
    IVAplicaCostoxProv = (Mid$(Opcion2, 116, 1) = "S")
End Property

Public Property Let IVActualizaPVIVA(ByVal value As Boolean)
       Mid$(mProps.Opcion2, 117, 1) = IIf(value, "S", "N")
   mbooModificado = True
End Property

Public Property Get IVActualizaPVIVA() As Boolean
    IVActualizaPVIVA = (Mid$(Opcion2, 117, 1) = "S")
End Property

Public Property Let IVTipoCostoVerificaFechaCompra(ByVal value As String)
    Mid$(mProps.Opcion2, 118, 1) = value
    mbooModificado = True
End Property

Public Property Get IVTipoCostoVerificaFechaCompra() As String
    If Mid$(mProps.Opcion2, 118, 1) = " " Then
        IVTipoCostoVerificaFechaCompra = "0"
    Else
        IVTipoCostoVerificaFechaCompra = Trim(Mid$(Opcion2, 118, 1))
    End If
End Property

Public Property Let IVMostrarAgencia(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 119, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVMostrarAgencia() As Boolean
    IVMostrarAgencia = (Mid$(Opcion2, 119, 1) = "S")
End Property

Public Property Let IVMostrarMarcacion(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 120, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVMostrarMarcacion() As Boolean
    IVMostrarMarcacion = (Mid$(Opcion2, 120, 1) = "S")
End Property

Public Property Let IVMostrarDescripcion2(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 121, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVMostrarDescripcion2() As Boolean
    IVMostrarDescripcion2 = (Mid$(Opcion2, 121, 1) = "S")
End Property


Public Property Let ArchivoReporteRIDE(ByVal value As String)
    If Len(value) > 200 Then
        Err.Raise ERR_INVALIDO, "GNTrans_ArchivoReporteRIDE", MSGERR_INVALIDO
    End If
    
    mProps.ArchivoReporteRIDE = value
    mbooModificado = True
End Property

Public Property Get ArchivoReporteRIDE() As String
    ArchivoReporteRIDE = Trim$(mProps.ArchivoReporteRIDE)
End Property

Public Property Let IVComprobanteElectronico(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 122, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVComprobanteElectronico() As Boolean
    IVComprobanteElectronico = (Mid$(Opcion2, 122, 1) = "S")
End Property

Public Property Let GeneraSolCredito(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 123, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property
Public Property Get GeneraSolCredito() As Boolean
    GeneraSolCredito = (Mid$(Opcion2, 123, 1) = "S")
End Property

Public Property Let CodEmpleadoPre(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodEmpleadoPre Then
        If Len(value) > 0 Then
            'Actualiza IdProvCli también
            sql = "SELECT IdProvCli FROM Empleado WHERE CodProvCli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNTrans_CodEmpleadoPre", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdEmpleadoPre = rs!IdProvcli
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdEmpleadoPre = 0
        End If
    End If

    mProps.CodEmpleadoPre = value
    mbooModificado = True
End Property

Public Property Get CodEmpleadoPre() As String
    CodEmpleadoPre = Trim$(mProps.CodEmpleadoPre)
End Property

Public Property Let IdEmpleadoPre(ByVal value As Long)
    mProps.IdEmpleadoPre = value
    mbooModificado = True
End Property

Public Property Get IdEmpleadoPre() As Long
    IdEmpleadoPre = mProps.IdEmpleadoPre
End Property

Public Property Let IVMostrarBotonCliGar(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 124, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property
Public Property Get IVMostrarBotonCliGar() As Boolean
    IVMostrarBotonCliGar = (Mid$(Opcion2, 124, 1) = "S")
End Property

Public Property Let IVMostrarEmpleado(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 125, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property
Public Property Get IVMostrarEmpleado() As Boolean
    IVMostrarEmpleado = (Mid$(Opcion2, 125, 1) = "S")
End Property

Public Property Let IVDatosImportacion(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 126, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVDatosImportacion() As Boolean
    IVDatosImportacion = (Mid$(Opcion2, 126, 1) = "S")
End Property

Public Property Let IVMostrarSaldoDespachar(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 127, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVMostrarSaldoDespachar() As Boolean
    IVMostrarSaldoDespachar = (Mid$(Opcion2, 127, 1) = "S")
End Property

Public Property Let IVCalcularDescPrecioPre(ByVal value As Boolean)
   Mid$(mProps.Opcion2, 128, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVCalcularDescPrecioPre() As Boolean
    IVCalcularDescPrecioPre = (Mid$(Opcion2, 128, 1) = "S")
End Property

Public Property Let IVCostoTotalSubItems(ByVal value As Boolean) 'AUC para rubros (ecuamueble)
   Mid$(mProps.Opcion2, 129, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVCostoTotalSubItems() As Boolean
    IVCostoTotalSubItems = (Mid$(Opcion2, 129, 1) = "S")
End Property

Public Property Let IVGeneraFormatoEstandarXML(ByVal value As Boolean)
   Mid$(mProps.Opcion2, 130, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVGeneraFormatoEstandarXML() As Boolean
    IVGeneraFormatoEstandarXML = (Mid$(Opcion2, 130, 1) = "S")
End Property

Public Property Let IVGeneraenBloqueXML(ByVal value As Boolean)
   Mid$(mProps.Opcion2, 131, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVGeneraenBloqueXML() As Boolean
    IVGeneraenBloqueXML = (Mid$(Opcion2, 131, 1) = "S")
End Property


Public Property Let IVGuiaRemision(ByVal value As Boolean)
   Mid$(mProps.Opcion2, 132, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVGuiaRemision() As Boolean
    IVGuiaRemision = (Mid$(Opcion2, 132, 1) = "S")
End Property

Public Property Let IVNoPermitirItemDuplicado(ByVal value As Boolean)
   Mid$(mProps.Opcion2, 133, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVNoPermitirItemDuplicado() As Boolean
    IVNoPermitirItemDuplicado = (Mid$(Opcion2, 133, 1) = "S")
End Property

Public Property Let IVCambioBodegaImportar(ByVal value As Boolean)
   Mid$(mProps.Opcion2, 134, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVCambioBodegaImportar() As Boolean
    IVCambioBodegaImportar = (Mid$(Opcion2, 134, 1) = "S")
End Property

Public Property Let IVAmbientePruebaFE(ByVal value As Boolean)
   Mid$(mProps.Opcion2, 135, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVAmbientePruebaFE() As Boolean
    IVAmbientePruebaFE = (Mid$(Opcion2, 135, 1) = "S")
End Property

Public Property Let IVClaveSeleccionarPU(ByVal value As Boolean)
   Mid$(mProps.Opcion2, 136, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVClaveSeleccionarPU() As Boolean
    IVClaveSeleccionarPU = (Mid$(Opcion2, 136, 1) = "S")
End Property
'Configura campos adicionales exportacion
Public Property Get IVVisibleAdicionalesExp() As Boolean
    IVVisibleAdicionalesExp = (Mid$(Opcion2, 137, 1) = "S")
End Property

Public Property Let IVVisibleAdicionalesExp(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 137, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property


Public Property Let NumCopiasImp(ByVal value As Integer)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "GNTrans_NumCopiasImp", MSGERR_INVALIDO
    End If
    mProps.NumCopiasImp = value
    mbooModificado = True
End Property

Public Property Get NumCopiasImp() As Integer
    NumCopiasImp = mProps.NumCopiasImp
End Property


Public Property Get IVVisiblePartidaArancel() As Boolean
    IVVisiblePartidaArancel = (Mid$(Opcion2, 138, 1) = "S")
End Property

Public Property Let IVVisiblePartidaArancel(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 138, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVActPlanMant() As Boolean
    IVActPlanMant = (Mid$(Opcion2, 139, 1) = "S")
End Property

Public Property Let IVActPlanMant(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 139, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSFormaCobro() As String 'BELLA LUZ
    TSFormaCobro = Mid$(Opcion2, 140, 5)
End Property

Public Property Let TSFormaCobro(ByVal value As String)
    Mid$(mProps.Opcion2, 140, 5) = value
    mbooModificado = True
End Property

Public Property Get TSVerCobrador() As Boolean 'BELLA LUZ clnd
    TSVerCobrador = (Mid$(Opcion2, 145, 1) = "S")
End Property

Public Property Let TSVerCobrador(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 145, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property
Public Property Get TSFiltroxCobrador() As Boolean 'BELLA LUZ clnd
    TSFiltroxCobrador = (Mid$(Opcion2, 146, 1) = "S")
End Property

Public Property Let TSFiltroxCobrador(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 146, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get TSFiltroxVendedor() As Boolean 'BELLA LUZ cobros
    TSFiltroxVendedor = (Mid$(Opcion2, 147, 1) = "S")
End Property

Public Property Let TSFiltroxVendedor(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 147, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Let FiltroItemxBod(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 148, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get FiltroItemxBod() As Boolean
    FiltroItemxBod = (Mid$(Opcion2, 148, 1) = "S")
End Property
'auc intervalo marcacion para lacto
Public Property Get IVIntMarcaProv() As Integer
    IVIntMarcaProv = Val(Mid$(Opcion2, 149, 2))
End Property

Public Property Let IVIntMarcaProv(ByVal value As Integer)
    If value > 99 Then
        Err.Raise ERR_INVALIDO, "GNTrans.IVIntMarcaProv", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion2, 149, 2) = Format(value, "00")
    mbooModificado = True
End Property

Public Property Let IVGeneraCompBod(ByVal value As Boolean) 'habilita opcion para Cambio de bodega
    Mid$(mProps.Opcion2, 151, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVGeneraCompBod() As Boolean
    IVGeneraCompBod = (Mid$(Opcion2, 151, 1) = "S")
End Property

Public Property Let IVTransCompBod(ByVal value As String)
    If Len(value) > 5 Then Err.Raise ERR_INVALIDO, "GNTrans.IVTransCompBod", MSGERR_INVALIDO
    Mid$(mProps.Opcion2, 152, 5) = String(5, " ")
    Mid$(mProps.Opcion2, 152, 5) = value
    mbooModificado = True
End Property

Public Property Get IVTransCompBod() As String
    IVTransCompBod = Trim((Mid$(Opcion2, 152, 5)))
End Property

Public Property Let IVFPagoxDCredito(ByVal value As Boolean) 'habilita opcion para Cambio de bodega
    Mid$(mProps.Opcion2, 157, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVFPagoxDCredito() As Boolean
    IVFPagoxDCredito = (Mid$(Opcion2, 157, 1) = "S")
End Property

Public Property Get IVPenalizaDesctoGlobal() As Boolean
    IVPenalizaDesctoGlobal = (Mid$(Opcion2, 158, 1) = "S")
End Property

Public Property Let IVPenalizaDesctoGlobal(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 158, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVGeneraCuota() As Boolean
    IVGeneraCuota = (Mid$(Opcion2, 159, 1) = "S")
End Property

Public Property Let IVGeneraCuota(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 159, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property


Public Property Let AsientoxPCG(ByVal value As Boolean)
    mProps.AsientoxPCG = value
   mbooModificado = True
End Property
Public Property Get AsientoxPCG() As Boolean
    AsientoxPCG = mProps.AsientoxPCG
End Property


Public Property Get PcKardex2() As Boolean
    PcKardex2 = (Mid$(Opcion2, 160, 1) = "S")
End Property

Public Property Let PcKardex2(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 160, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Let IVPorPagar2PCK(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 161, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVPorPagar2PCK() As Boolean
    IVPorPagar2PCK = (Mid$(Opcion2, 161, 1) = "S")
End Property

Public Property Let IVVisiblePCK2(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 162, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVVisiblePCK2() As Boolean
    IVVisiblePCK2 = (Mid$(Opcion2, 162, 1) = "S")
End Property

Public Property Get PcDosKardex() As Boolean
    PcDosKardex = (Mid$(Opcion2, 163, 1) = "S")
End Property

Public Property Let PcDosKardex(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 163, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVRebateIVG() As Boolean
    IVRebateIVG = (Mid$(Opcion2, 164, 1) = "S")
End Property

Public Property Let IVRebateIVG(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 164, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVImprimeBoleto() As Boolean
    IVImprimeBoleto = (Mid$(Opcion2, 165, 1) = "S")
End Property

Public Property Let IVImprimeBoleto(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 165, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVDesctoAlGrabar() As Boolean
    IVDesctoAlGrabar = (Mid$(Opcion2, 166, 1) = "S")
End Property

Public Property Let IVDesctoAlGrabar(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 166, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Let IVDetalleCurier(ByVal value As Boolean) 'habilita opcion para Cambio de bodega
   Mid$(mProps.Opcion2, 167, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property
Public Property Get IVDetalleCurier() As Boolean
    IVDetalleCurier = (Mid$(Opcion2, 167, 1) = "S")
End Property

Public Property Let IVMantenerDescYProm(ByVal value As Boolean) 'habilita opcion para Cambio de bodega
   Mid$(mProps.Opcion2, 168, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVMantenerDescYProm() As Boolean
    IVMantenerDescYProm = (Mid$(Opcion2, 168, 1) = "S")
End Property

Public Property Let IVObservacionItemObliga(ByVal value As Boolean) 'habilita opcion para Cambio de bodega
   Mid$(mProps.Opcion2, 169, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVObservacionItemObliga() As Boolean
    IVObservacionItemObliga = (Mid$(Opcion2, 169, 1) = "S")
End Property

Public Property Let IVVisibleReceta(ByVal value As Boolean)
   Mid$(mProps.Opcion2, 170, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property
Public Property Get IVVisibleReceta() As Boolean
    IVVisibleReceta = (Mid$(Opcion2, 170, 1) = "S")
End Property

Public Property Let IVNoCambiaEstadoTrans(ByVal value As Boolean)
   Mid$(mProps.Opcion2, 171, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVNoCambiaEstadoTrans() As Boolean
    IVNoCambiaEstadoTrans = (Mid$(Opcion2, 171, 1) = "S")
End Property

Public Property Let IVBandPorcTC(ByVal value As Boolean)
   Mid$(mProps.Opcion2, 172, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVBandPorcTC() As Boolean
    IVBandPorcTC = (Mid$(Opcion2, 172, 1) = "S")
End Property

Public Property Let IVPorcTC(ByVal value As Integer)
    If value > 99 Then
        Err.Raise ERR_INVALIDO, "GNTrans.IVPorcTC", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion2, 173, 2) = Format(value, "00")
    mbooModificado = True
End Property

Public Property Get IVPorcTC() As Integer
    IVPorcTC = Val(Mid$(Opcion2, 173, 2))
End Property


Public Property Let IVxml6Decimales(ByVal value As Boolean)
   Mid$(mProps.Opcion2, 175, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVxml6Decimales() As Boolean
    IVxml6Decimales = (Mid$(Opcion2, 175, 1) = "S")
End Property

Public Property Let IVTransfAuto(ByVal value As Boolean) 'Transformacion automatica
    Mid$(mProps.Opcion2, 176, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVTransfAuto() As Boolean
    IVTransfAuto = (Mid$(Opcion2, 176, 1) = "S")
End Property

Public Property Let IVCodTransfAuto(ByVal value As String) 'Transformacion automatica
    If Len(value) > 5 Then Err.Raise ERR_INVALIDO, "GNTrans.IVTransAuto", MSGERR_INVALIDO
    Mid$(mProps.Opcion2, 177, 5) = String(5, " ")
    Mid$(mProps.Opcion2, 177, 5) = value
    mbooModificado = True
End Property

Public Property Get IVCodTransfAuto() As String
    IVCodTransfAuto = Trim((Mid$(Opcion2, 177, 5)))
End Property

Public Property Get IVCCXFila() As Boolean
    IVCCXFila = (Mid$(Opcion2, 182, 1) = "S")
End Property

Public Property Let IVCCXFila(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 182, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVAgregaFilaIVGrupo() As Boolean
    IVAgregaFilaIVGrupo = (Mid$(Opcion2, 183, 1) = "S")
End Property

Public Property Let IVAgregaFilaIVGrupo(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 183, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVActualizaDsctoItemxCli() As Boolean
    IVActualizaDsctoItemxCli = (Mid$(Opcion2, 184, 1) = "S")
End Property

Public Property Let IVActualizaDsctoItemxCli(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 184, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Let ConfigBotones(ByVal value As String)
'    If Len(value) > 286 Then
'        Err.Raise ERR_INVALIDO, "GNTrans_Opcion2", MSGERR_INVALIDO
'    End If
    mProps.ConfigBotones = value
    mbooModificado = True
End Property

Public Property Get ConfigBotones() As String
    ConfigBotones = Trim$(mProps.ConfigBotones)
End Property

Public Property Get IVCantBotones() As Integer
    IVCantBotones = IIf(Mid$(Opcion2, 185, 1) = "", 1, Mid$(Opcion2, 185, 1))
End Property

Public Property Let IVCantBotones(ByVal value As Integer)
    Mid$(mProps.Opcion2, 185, 1) = value
    mbooModificado = True
End Property


Public Property Get IVMostrarBotones() As Boolean
    IVMostrarBotones = (Mid$(Opcion2, 186, 1) = "S")
End Property

Public Property Let IVMostrarBotones(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 186, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Let RutaXMLOrigen(ByVal value As String)
    mProps.RutaXMLOrigen = value
    mbooModificado = True
End Property

Public Property Get RutaXMLOrigen() As String
    RutaXMLOrigen = mProps.RutaXMLOrigen
End Property

Public Property Let RutaXMLDestino(ByVal value As String)
    mProps.RutaXMLDestino = value
    mbooModificado = True
End Property

Public Property Get RutaXMLDestino() As String
    RutaXMLDestino = mProps.RutaXMLDestino
End Property

Public Property Get IVMostrarFirmaDig() As Boolean
    IVMostrarFirmaDig = (Mid$(Opcion2, 187, 1) = "S")
End Property

Public Property Let IVMostrarFirmaDig(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 187, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Let IVPreciolibPersonal(ByVal value As String)
    If Len(value) > 5 Then
        Err.Raise ERR_INVALIDO, "GNTrans.IVValorPesoPersonal", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion2, 188, 5) = String(5, " ")
    Mid$(mProps.Opcion2, 188, 5) = value
    mbooModificado = True
End Property

Public Property Get IVPreciolibPersonal() As String
    IVPreciolibPersonal = Trim$(Mid$(mProps.Opcion2, 188, 5))
End Property

Public Property Let IVPreciolibComercial(ByVal value As String)
    If Len(value) > 5 Then
        Err.Raise ERR_INVALIDO, "GNTrans.IVValorPesoPersonal", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion2, 193, 5) = String(5, " ")
    Mid$(mProps.Opcion2, 193, 5) = value
    mbooModificado = True
End Property

Public Property Get IVPreciolibComercial() As String
    IVPreciolibComercial = Trim$(Mid$(mProps.Opcion2, 193, 5))
End Property

Public Property Let NumRefSig(ByVal value As Long)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "GNTrans_NumRefSig", MSGERR_INVALIDO
    End If
    mProps.NumRefSig = value
    mbooModificado = True
End Property

Public Property Get NumRefSig() As Long
    NumRefSig = mProps.NumRefSig
End Property

Public Property Get IVNumRefAuto() As Boolean
    IVNumRefAuto = (Mid$(Opcion2, 198, 1) = "S")
End Property

Public Property Let IVNumRefAuto(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 198, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVBusquedaBodegasSuc() As Boolean
    IVBusquedaBodegasSuc = (Mid$(Opcion2, 199, 1) = "S")
End Property

Public Property Let IVBusquedaBodegasSuc(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 199, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get IVTransfHijoPadre() As Boolean
    IVTransfHijoPadre = (Mid$(Opcion2, 200, 1) = "S")
End Property

Public Property Let IVTransfHijoPadre(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 200, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Let IVValorPeso(ByVal value As String)
    If Len(value) > 5 Then
        Err.Raise ERR_INVALIDO, "GNTrans.IVValorPeso", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion2, 201, 5) = String(5, " ")
    Mid$(mProps.Opcion2, 201, 5) = value
    mbooModificado = True
End Property

Public Property Get IVValorPeso() As String
    IVValorPeso = Trim$(Mid$(mProps.Opcion2, 201, 5))
End Property

Public Property Let IVValorPaq(ByVal value As String)
    If Len(value) > 5 Then
        Err.Raise ERR_INVALIDO, "GNTrans.IVValorPeso", MSGERR_INVALIDO
    End If
    Mid$(mProps.Opcion2, 206, 5) = String(5, " ")
    Mid$(mProps.Opcion2, 206, 5) = value
    mbooModificado = True
End Property

Public Property Get IVValorPaq() As String
    IVValorPaq = Trim$(Mid$(mProps.Opcion2, 206, 5))
End Property

Public Property Get IVVisibleCalendario() As Boolean
    IVVisibleCalendario = (Mid$(Opcion2, 211, 1) = "S")
End Property

Public Property Let IVVisibleCalendario(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 211, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

Public Property Get EnvioAutomaticoEmail() As Boolean
    EnvioAutomaticoEmail = (Mid$(Opcion2, 212, 1) = "S")
End Property

Public Property Let EnvioAutomaticoEmail(ByVal value As Boolean)
    Mid$(mProps.Opcion2, 212, 1) = IIf(value, "S", "N")
    mbooModificado = True
End Property

