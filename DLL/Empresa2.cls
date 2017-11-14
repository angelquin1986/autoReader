VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Empresa2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
Private Const BAL_IDCUENTA = 0
Private Const BAL_CODCUENTA = 1
Private Const BAL_NOMBRECUENTA = 2
Private Const BAL_CODSUCURSAL = 3       'JEAA
Private Const BAL_VALOR = 4
Private Const BAL_VALMODPRESUPUESTO = 9
Private Const BAL_NIVEL = 5
Private Const BAL_IDPADRE = 6
Private Const BAL_TIPO = 7
Private Const BAL_BANDTOTAL = 8

Public Function ConsGNTrans22Dou( _
                    Optional ByVal OrdenAscendente As Boolean) As Variant
    Dim sql As String, cond As String, s As String, rs As Recordset
    sql = "SELECT"
    If gobjCond.TopN > 0 Then
        sql = sql & " TOP " & gobjCond.TopN
    End If
    
    sql = sql & " VW.TransID, FechaTrans, CodAsiento, CodTrans, NumTrans, " & _
            " NumDocRef, Nombre, vw.Descripcion, ROUND((costoRealTotalDou/cantidadDou),8) as CUR, ROUND((costoTotalDou/cantidadDou),8) as CU, CodCentro,Estado FROM " & _
            " (IVGrupo1 RIGHT JOIN IVInventario ON IVGrupo1.IdGrupo1 = IVInventario.IdGrupo1) inner join " & _
            " vwConsGNTrans " & _
            " VW INNER JOIN IVKARDEX IVK ON IVK.TRANSID=VW.TRANSID " & _
            " ON IVInventario.IdInventario = IVK.IdInventario "
    
    'Condición de FechaTrans
    cond = "FechaTrans BETWEEN " & FechaYMD(gobjCond.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & FechaYMD(gobjCond.fecha2, gobjMain.EmpresaActual.TipoDB)
    
    'Condición de CodTrans
    If Len(cond) > 0 Then cond = cond & " AND "
    If Len(gobjCond.CodTrans) > 0 Then      'Si está especificado transaccion
        'Si es una sola transaccion
        If InStr(gobjCond.CodTrans, ",") = 0 Then
            cond = cond & "(CodTrans='" & gobjCond.CodTrans & "')"
        'Si está seleccionado varias trans.
        Else
            cond = cond & "(CodTrans IN (" & gobjCond.CodTrans & "))"
        End If
    Else
        'Si no hay condicion de cod.trans, selecciona todas las trans.
        ' que tenga permiso
        s = gobjMain.GrupoActual.PermisoActual.ListaTrans(True)
        If Len(s) > 0 Then
            cond = cond & "(CodTrans " & s & ")"
        Else
            cond = cond & "(CodTrans = '-------')"    'Para que NO devuelva nada si no tiene ningún permiso  '*** MAKOTO 26/dic/00
        End If
    End If
    
    'Condición de NumTrans
    If (gobjCond.NumTrans1 > 0) And (gobjCond.NumTrans2 > 0) Then
        If Len(cond) > 0 Then cond = cond & " AND "
        cond = cond & "(NumTrans BETWEEN " & gobjCond.NumTrans1 & " AND " & gobjCond.NumTrans2 & ")"
    ElseIf gobjCond.NumTrans1 > 0 Then
        If Len(cond) > 0 Then cond = cond & " AND "
        cond = cond & "(NumTrans>=" & gobjCond.NumTrans1 & ")"
    ElseIf gobjCond.NumTrans2 > 0 Then
        If Len(cond) > 0 Then cond = cond & " AND "
        cond = cond & "(NumTrans<=" & gobjCond.NumTrans2 & ")"
    End If
    
    'Condición de Estado
    If gobjCond.EstadoBool(ESTADO_NOAPROBADO) = False Then
        If Len(cond) > 0 Then cond = cond & " AND "
        cond = cond & "(Estado<>" & ESTADO_NOAPROBADO & ")"
    End If
    If gobjCond.EstadoBool(ESTADO_APROBADO) = False Then
        If Len(cond) > 0 Then cond = cond & " AND "
        cond = cond & "(Estado<>" & ESTADO_APROBADO & ")"
    End If
    If gobjCond.EstadoBool(ESTADO_DESPACHADO) = False Then
        If Len(cond) > 0 Then cond = cond & " AND "
        cond = cond & "(Estado<>" & ESTADO_DESPACHADO & ")"
    End If
    If gobjCond.EstadoBool(ESTADO_ANULADO) = False Then
        If Len(cond) > 0 Then cond = cond & " AND "
        cond = cond & "(Estado<>" & ESTADO_ANULADO & ")"
    End If
    If Not (Len(gobjCond.Grupo1) = 0 And Len(gobjCond.Grupo2) = 0) Then
        cond = cond & " AND (codGrupo" & gobjCond.NumGrupo & " BETWEEN '" & _
                gobjCond.Grupo1 & "' AND '" & gobjCond.Grupo2 & "') "
    End If
    If Not (Len(gobjCond.CodItem1) = 0 And Len(gobjCond.CodItem2) = 0) Then
        cond = cond & "AND (CodInventario BETWEEN '" & gobjCond.CodItem1 & "' AND '" & gobjCond.CodItem2 & "') "
    End If
        
    If Len(cond) > 0 Then sql = sql & " WHERE " & cond
    'Ordern de filas
    If OrdenAscendente Then
        sql = sql & " ORDER BY FechaTrans, HoraTrans, VW.TransID"
    Else
        sql = sql & " ORDER BY FechaTrans DESC, HoraTrans DESC, VW.TransID DESC"
    End If
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    Set ConsGNTrans22Dou = rs
    Set rs = Nothing
    gobjMain.EmpresaActual.GrabaGNLogAccion "CONSULTA", "Consulta General " & gobjCond.fecha1 & "-" & gobjCond.fecha2, "GN"
End Function

Public Function ListaIVRecetasItemPadre(BandRS As Boolean) As Variant
    Dim sql As String, rs As Recordset
    sql = " Select iv.codInventario,iv.descripcion "
    sql = sql & " From  IVInventario iv"
    sql = sql & " Inner join IvReceta ivr on ivr.idinvPadre= iv.idInventario "
    sql = sql & " Group by iv.codInventario,iv.descripcion"
    sql = sql & " order by iv.codInventario desc"

   Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If BandRS Then
        Set ListaIVRecetasItemPadre = rs
    Else
        ListaIVRecetasItemPadre = MiGetRows(rs)
        rs.Close
    End If
    Set rs = Nothing
End Function

Public Function ListaIVRecetas(BandRS As Boolean, Valida As Boolean) As Variant
    Dim sql As String, rs As Recordset
    Dim cond As String
    With gobjMain.objCondicion
        If Len(.CodItem1) > 0 And Len(.CodItem2) > 0 Then
            If Len(cond) = 0 Then cond = "Where "
            cond = cond & "iv.codinventario Between '" & .CodItem1 & "' And '" & .CodItem2 & "' "
        End If
        
        If .fecha1 > 0 And .fecha2 > 0 Then
            If Len(cond) = 0 Then
                cond = "Where "
            Else
                cond = cond & " And "
            End If
            cond = cond & "(ivr.FechaGrabado BETWEEN " & FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ")"
        End If
        
        If Len(.CodReceta) > 0 Then
            If Len(cond) = 0 Then
                cond = "Where "
            Else
                cond = cond & " And "
            End If
            cond = cond & "(ivr.CodReceta Like '%" & .CodReceta & "%')"
        End If
        If Len(.Receta) > 0 Then
            If Len(cond) = 0 Then
                cond = "Where "
            Else
                cond = cond & " And "
            End If
            cond = cond & "(ivr.descripcion Like '%" & .Receta & "%')"
        End If

        sql = "Select ivr.codReceta,ivr.descripcion,iv.codInventario,iv.descripcion,ivr.batch,ivr.fechaGrabado,ivr.bandvalida "
        sql = sql & " from  IVReceta ivr"
        sql = sql & " left join IvInventario iv on iv.idinventario = ivr.idinvPadre "
        sql = sql & cond
        If Valida Then
            sql = sql & "And ivr.bandvalida = 1 "
        End If
        sql = sql & " order by ivr.id desc"
    End With
   Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If BandRS Then
        Set ListaIVRecetas = rs
    Else
        ListaIVRecetas = MiGetRows(rs)
        rs.Close
    End If
    Set rs = Nothing
End Function
Public Function ListaIVRecetasFlex(BandRS As Boolean, Valida As Boolean) As Variant
    Dim sql As String, rs As Recordset
    Dim cond As String
    sql = "Select codReceta,descripcion "
    sql = sql & " from  IVReceta "

    If Valida Then
        sql = sql & "Where bandvalida = 1 "
    End If
    sql = sql & " order by id desc"
   Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If BandRS Then
        Set ListaIVRecetasFlex = rs
    Else
        ListaIVRecetasFlex = MiGetRows(rs)
        rs.Close
    End If
    Set rs = Nothing
End Function


Public Function ConsANVentas2016paraXML() As Variant
    Dim pck As PCKardex, pcd As PCDocAsignado
    Dim objGnComp As GNComprobante, objGnCompDocOriginal As GNComprobante
    Dim sql As String, cond As String, CadenaValores As String
    Dim Recargo As String 'CadenaAgrupa  As String
    Dim v As Variant, max As Integer, i As Integer, j As Integer, X As Integer
    Dim from As String, numReg As Long, f1 As String
    Dim rs As Recordset, Retenidos As Variant, NumRet As Long, v1 As Variant, codRecarBaseCero As String, codRecarBaseNoCero As String
    Dim Servicios As String, Bienes As String, RowRet As Long, m As Integer, rsRecargo As Recordset
    Dim rsforma As Recordset
    Const COL_BASESER = 15
    'segunda parte de la consulta: carga los valores de  Detalles de Anexos
    Const COL_FECHA = 2
    Const COL_CODCREDTRIB = 7
    Const COL_TRANSID = 1
    Const COL_TRANSRET = 5
    Dim condServ As String
    Dim condBien As String
    Dim Vista As String
    Dim pos As Integer, pos1 As Integer
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
   With gobjCond
        'Reporte de un mes a la vez
        f1 = DateSerial(Year(.fecha1), Month(.fecha1), 1)
        cond = " AND GNC.FechaTrans BETWEEN " & FechaYMD(f1, gobjMain.EmpresaActual.TipoDB) & _
               " AND " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB)
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            VerificaExistenciaTabla 2
            VerificaExistenciaTabla 3
            pos = InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "DISFRA")
            pos1 = InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "ENCOFRADO")
            
            If gobjMain.EmpresaActual.GNOpcion.IVKTipoDatoDouble Then
                Vista = "vwConsSUMIVKardexIVADou"
            Else
               If pos > 0 Or pos1 > 0 Then
                    Vista = "vwConsSUMIVKardexIVA2B"
                Else
                    Vista = "vwConsSUMIVKardexIVA"
                End If
            End If

        If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
            sql = "Select substring(gnt.opcion2,122,1) AS BandElect, "
        Else
            sql = " Select  'N' AS BandElect, "
        End If

'               sql = "Select substring(gnt.opcion2,122,1) AS BandElect, "
               sql = sql & " Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  INNER JOIN GnTrans gnt on gnc.codtrans=GnT.codtrans " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
                sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(.Sucursal) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(.CodGrupo) & ")" & _
                  "Group by IvkR.TransID,  substring(gnt.opcion2,122,1)"
            
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
        VerificaExistenciaTablaTemp "tan1"
        If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
            sql = "Select substring(gnt.opcion2,122,1) AS BandElect, "
        Else
            sql = " Select  'N' AS BandElect, "
        End If
        sql = sql & "pc.ruc, SUM(IvKr.Valor) as TotalIVA Into tan1 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  INNER JOIN GnTrans gnt on gnc.codtrans=GnT.codtrans inner join pcprovcli pc on gnc.idclienteref= pc.idprovcli " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
                sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN ('IVA') " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(.CodGrupo) & ")" & _
                  "Group by pc.ruc , substring(gnt.opcion2,122,1)"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
        VerificaExistenciaTablaTemp "tmpICE"
        
        
        sql = "Select  pc.ruc, SUM(IvKr.Valor) as TotalICE , "
        If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
            sql = sql & "  substring(gnt.opcion2,122,1) as CompElect, "
            sql = sql & "  substring(gnt.opcion2,135,1) as Prod "
        Else
            sql = sql & "  'N' as CompElect, "
            sql = sql & "  'N' as Prod "
        End If
        sql = sql & "  Into tmpICE "
        
        sql = sql & " From IvRecargo ivR inner join "
        sql = sql & " IvKardexRecargo ivkR Inner join "
        sql = sql & " GnComprobante gNc  "
        sql = sql & " inner join GnTrans gnt on gnc.codtrans=GnT.codtrans"
        sql = sql & " inner join pcprovcli pc on gnc.idclienteref= pc.idprovcli "
        sql = sql & " On ivkr.TransID = gNc.TransID "
        sql = sql & " On Ivr.IdRecargo = IvkR.IdRecargo "
        sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo = 'ICE' " & cond
        sql = sql & " AND GNC.CodTrans IN (" & PreparaCadena(.CodGrupo) & ")"
        sql = sql & " Group by pc.ruc , substring(gnt.opcion2,122,1) "
        If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
            sql = sql & " , substring(gnt.opcion2,122,1), substring(gnt.opcion2,135,1)"
        End If
                  
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
         sql = "select codrecargo, BandBaseCero from ivrecargo where BandAfectaBaseAnexo=1  "
         Set rsRecargo = gobjMain.EmpresaActual.OpenRecordset(sql)
         
         If Not rsRecargo.EOF Then
            rsRecargo.MoveFirst
         End If
        codRecarBaseCero = ""
        codRecarBaseNoCero = ""
         
         For i = 1 To rsRecargo.RecordCount
            If rsRecargo.RecordCount > 0 Then
               If rsRecargo.Fields(1) Then
                   codRecarBaseCero = codRecarBaseCero & "'" & rsRecargo.Fields(0) & "',"
               Else
                   codRecarBaseNoCero = codRecarBaseNoCero & "'" & rsRecargo.Fields(0) & "',"
               End If
            End If
            rsRecargo.MoveNext
         Next i
        
        If Len(codRecarBaseCero) > 0 Then
            codRecarBaseCero = Mid$(codRecarBaseCero, 1, Len(codRecarBaseCero) - 1)
        End If
            
        If Len(codRecarBaseNoCero) > 0 Then
            codRecarBaseNoCero = Mid$(codRecarBaseNoCero, 1, Len(codRecarBaseNoCero) - 1)
        End If
            
            
        
        
        sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalRecargoBaseCero "
        sql = sql & " Into tmpRecarBaseCero "
        sql = sql & " From IvRecargo ivR inner join "
        sql = sql & " IvKardexRecargo ivkR Inner join "
        sql = sql & " GnComprobante gNc  "
        sql = sql & " On ivkr.TransID = gNc.TransID "
        sql = sql & " On Ivr.IdRecargo = IvkR.IdRecargo "
        sql = sql & " WHERE gNc.Estado <> 3 "
        If Len(codRecarBaseCero) > 0 Then
            sql = sql & " AND ivr.CodRecargo IN (" & codRecarBaseCero & ") " & cond
        End If
        sql = sql & " AND GNC.CodTrans IN (" & PreparaCadena(.CodGrupo) & ")"
        sql = sql & " Group by IvkR.TransID"
        
        VerificaExistenciaTablaTemp "tmpRecarBaseCero"
        gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            
            
            
            
'--datos de indices
            VerificaExistenciaTabla 3
            sql = " Select  DISTINCT GNC.IdClienteRef AS FC_id, GNC_b.IdClienteRef AS rtp_id1, "
            sql = sql & " CASE WHEN GNC_b.IdClienteRef=0 THEN PCK_b.IdProvcli WHEN GNC_b.IdClienteRef<>0 THEN GNC_b.IdClienteRef end as rtp_id "
            sql = sql & " Into tmp3"
            sql = sql & " From    (gncomprobante GNC Inner join pckardex pck_a ON GNC.TransID = pck_a.TransID ) INNER JOIN"
            sql = sql & "  (gncomprobante GNC_b Inner join pckardex pck_B ON GNC_b.TransID = pck_B.TransID )"
            sql = sql & " ON  PCK_A.ID = PCK_B.IDASIGNADO"
            sql = sql & " where GNC.CodTrans IN (" & PreparaCadena(gobjCond.CodGrupo) & ")"
            sql = sql & " AND GNC_b.CodTrans IN (" & PreparaCadena(gobjCond.CodRetencion2) & ")"
            sql = sql & " AND (GNC_b.FechaTrans  BETWEEN " & _
                FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = sql & " and GNC.Estado<>3 and GNC.IdClienteRef<>0 and GNC_b.IdClienteRef<>0 "
            
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            VerificaExistenciaTabla 3
            sql = sql & " select"
            sql = sql & " pc.idprovcli, sum(valor) as ValorIce"
            sql = sql & " from gncomprobante g inner join pcprovcli pc on g.IdClienteRef = pc.IdProvCli"
            sql = sql & " inner join ivkardexrecargo i"
            sql = sql & " inner join IVRecargo ir"
            sql = sql & " on i.idrecargo=i.idrecargo"
            sql = sql & " on g.transid=i.transid"
            sql = sql & " where codrecargo='ice'"
            'sql = sql & " and CodTrans in ('fcm')"
            sql = sql & " AND CodTrans IN (" & PreparaCadena(.CodGrupo) & ")"
            sql = sql & " AND (fechatrans BETWEEN " & _
                FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = sql & " group by pc.idprovcli"
            
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            
            '--datos de la Venta
            sql = " Select pc.ruc, ' ' as Nombre, atd.CodTipoDocumento, pc.BandRelacionado, "
            sql = sql & " AnexoCodTipoComp, count(Gnc.NumTrans) as cantTrans,  "
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN (sum( vw.PrecioRealTotalBase0) *-1 + (isnull(sum(tr.TotalRecargoBaseCero),0)*-1) )END As Valor0,"
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotalBaseIVA) *-1 END AS Valor12,"
            sql = sql & " 0 AS ValorNO12,"
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotal) *-1 END  AS PrecioRealTotal, "
            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
                sql = sql & "  substring(gnt.opcion2,122,1) as CompElect, "
                sql = sql & "  substring(gnt.opcion2,135,1) as Prod "
            Else
                sql = sql & "  'N' as CompElect, "
                sql = sql & "  'N' as Prod "
            End If
            sql = sql & " Into tmp1"
            sql = sql & " from    ("
            sql = sql & " tmp0"
            sql = sql & " Right join gncomprobante Gnc "
            sql = sql & " inner join GnTrans gnt on gnc.codtrans=GnT.codtrans"
            sql = sql & " INNER JOIN PCPROVClI pc "
            sql = sql & " left join anexo_TipoDocumento atd on pc.IdTipoDocumento=atd.Id"
            sql = sql & " ON GNC.IdClienteRef=pc.idprovcli "
            sql = sql & " inner join " & Vista & " vw"
            sql = sql & " ON Gnc.TransID = vw.transid"
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " left join tmpRecarBaseCero tr ON tmp0.TransID = tr.TransID"
            
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(gobjCond.CodGrupo) & ")"
            sql = sql & " and Gnc.BandCierre<>9 "
            sql = sql & " and GNC.Estado<>3 " & cond
            sql = sql & " and pc.ruc<>'9999999999999'"
            sql = sql & " and  (atd.CodTipoDocumento in ( 'R','C') or atd.CodTipoDocumento is null)"
            sql = sql & " group by  pc.ruc,   AnexoCodTipoComp  ,CodTipoDocumento,  pc.BandRelacionado "
            'gnc.PorcentajeIVA ,
            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
                sql = sql & " , substring(gnt.opcion2,122,1), substring(gnt.opcion2,135,1)"
            End If
           
           sql = sql & " union all "
            
            sql = sql & " Select  pc.ruc, ' ' as Nombre, atd.CodTipoDocumento , pc.BandRelacionado, "
            sql = sql & " AnexoCodTipoComp, count(Gnc.NumTrans) as cantTrans,  "
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN (sum( vw.PrecioRealTotalBase0) *-1  + (isnull(sum(tr.TotalRecargoBaseCero),0)*-1) ) END As Valor0,"
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotalBaseIVA) *-1 END AS Valor12,"
            sql = sql & " 0 AS ValorNO12,"
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotal) *-1 END  AS PrecioRealTotal, "
            'sql = sql & " gnc.PorcentajeIVA,            "
            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
                sql = sql & "  substring(gnt.opcion2,122,1) as CompElect, "
                sql = sql & "  substring(gnt.opcion2,135,1) as Prod "
            Else
                sql = sql & "  'N' as CompElect, "
                sql = sql & "  'N' as Prod "
            End If
            sql = sql & " from    ("
            sql = sql & " tmp0"
            sql = sql & " Right join gncomprobante Gnc "
            sql = sql & " inner join GnTrans gnt on gnc.codtrans=GnT.codtrans"
            sql = sql & " INNER JOIN PCPROVClI pc "
            sql = sql & " left join anexo_TipoDocumento atd on pc.IdTipoDocumento=atd.Id"
            sql = sql & " ON GNC.IdClienteRef=pc.idprovcli "
            sql = sql & " inner join " & Vista & " vw"
            sql = sql & " ON Gnc.TransID = vw.transid"
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " left join tmpRecarBaseCero tr ON tmp0.TransID = tr.TransID"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(gobjCond.CodGrupo) & ")"
            sql = sql & " AND Gnc.BandCierre<>9 "
            sql = sql & " and GNC.Estado<>3 " & cond
            sql = sql & " and pc.ruc<>'9999999999999'"
            sql = sql & " and  (atd.CodTipoDocumento not in ( 'R','C') or atd.CodTipoDocumento is null)"
            sql = sql & " group by  pc.ruc ,AnexoCodTipoComp , CodTipoDocumento, pc.BandRelacionado "
            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
                sql = sql & " , substring(gnt.opcion2,122,1), substring(gnt.opcion2,135,1)"
            End If
           
           sql = sql & " union all "
            
            sql = sql & " Select '9999999999999' AS ruc, ' ' as Nombre, 'F' as CodTipoDocumento, 2 as BandRelacionado,"
            sql = sql & " AnexoCodTipoComp, count(Gnc.NumTrans) as cantTrans,  "
            
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN (sum( vw.PrecioRealTotalBase0) *-1 + (isnull(sum(tr.TotalRecargoBaseCero),0)*-1) ) END As Valor0,"
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotalBaseIVA) *-1 END AS Valor12,"
            sql = sql & " 0 AS ValorNO12,"
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotal) *-1 END  AS PrecioRealTotal, "
            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
                sql = sql & "  substring(gnt.opcion2,122,1) as CompElect, "
                sql = sql & "  substring(gnt.opcion2,135,1) as Prod "
            Else
                sql = sql & "  'N' as CompElect, "
                sql = sql & "  'N' as Prod "
            End If
            sql = sql & " from    ("
            sql = sql & " tmp0"
            sql = sql & " Right join gncomprobante Gnc "
            sql = sql & " inner join GnTrans gnt on gnc.codtrans=GnT.codtrans"
            sql = sql & " INNER JOIN PCPROVClI pc "
            sql = sql & " left join anexo_TipoDocumento atd on pc.IdTipoDocumento=atd.Id"
            sql = sql & " ON GNC.IdClienteRef=pc.idprovcli "
            sql = sql & " inner join " & Vista & " vw"

            sql = sql & " ON Gnc.TransID = vw.transid"
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " left join tmpRecarBaseCero tr ON tmp0.TransID = tr.TransID"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(gobjCond.CodGrupo) & ")"
            sql = sql & " AND Gnc.BandCierre<>9 "
            sql = sql & " and GNC.Estado<>3 " & cond
            sql = sql & " and  (atd.CodTipoDocumento not in ( 'R','C','P')"
            sql = sql & " OR pc.ruc='9999999999999')"
            sql = sql & " group by  AnexoCodTipoComp "
            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
                sql = sql & " , substring(gnt.opcion2,122,1), substring(gnt.opcion2,135,1) "
            End If
           
           
           VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            
            sql = " select ruc  into tmp4 from tmp1 where anexocodtipocomp in (5,18) "
            
            VerificaExistenciaTabla 4
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            
            sql = " Select pck.idformaSRI, Count(pck.idformaSRI) as cant "
            sql = sql & " From GnComprobante gNc"
            sql = sql & " inner join pckardex pck"
            sql = sql & " on gnc.transid= pck.transid"
            sql = sql & " Where gnc.Estado <> 3"
            sql = sql & " AND (GNC.fechatrans BETWEEN " & _
                FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = sql & " AND gnc.CodTrans IN (" & PreparaCadena(.CodGrupo) & ")"
            sql = sql & " Group by pck.idformaSRI"
            sql = sql & " having Count(pck.idformaSRI)>0"
            sql = sql & " order by pck.idformaSRI "
            
           Set rsforma = gobjMain.EmpresaActual.OpenRecordset(sql)
            If Not rsforma.EOF Then
                rsforma.MoveFirst
            End If
            
            For m = 0 To rsforma.RecordCount - 1
                If Not IsNull(rsforma.Fields(0)) Then
                    sql = " Select"
                    'sql = sql & " case when idtipodocumento =5 then '9999999999999' else ruc end as ruc,  CodFormaPago "
                    sql = sql & " substring(gnt.opcion2,122,1) AS BandElect,  ruc ,  CodFormaPago "
                    sql = sql & " Into tmpForma" & m
                    sql = sql & " From GnComprobante gNc INNER JOIN GnTrans gnt on gnc.codtrans=GnT.codtrans "
                    sql = sql & " inner join pckardex pck"
                    sql = sql & " left join Anexo_FormaPago a  on pck.idformasri=a.id  "
                    sql = sql & " on gnc.transid= pck.transid"
                    sql = sql & " inner join pcprovcli pc on gnc.idclienteref=pc.idprovcli"
                    sql = sql & " Where gnc.Estado <> 3"
                    sql = sql & " AND (GNC.fechatrans BETWEEN " & _
                        FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                        FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
                    sql = sql & " AND gnc.CodTrans IN (" & PreparaCadena(.CodGrupo) & ")"
                    sql = sql & " and pck.IdFormaSRI=" & rsforma.Fields(0)
                    sql = sql & " Group by pc.ruc,  CodFormaPago, idtipodocumento , substring(gnt.opcion2,122,1) "
                
                    VerificaExistenciaTablaTemp "tmpForma" & m
                    gobjMain.EmpresaActual.EjecutarSQL sql, numReg
                Else
                    MsgBox "Error falta enlaces de Forma de Cobro con el Catalogo del SRI "
                    Exit Function
                End If
                rsforma.MoveNext
            Next m
            
            If rsforma.RecordCount > 0 Then
                rsforma.MoveFirst
            End If
            
            
            sql = " Select"
            sql = sql & " compelect as BandElect, tmp1.ruc, "
            If Not rsforma.EOF Then
                For m = 0 To rsforma.RecordCount - 1
                    sql = sql & "tmpForma" & m & ".CodformaPago as Codforma" & m & ", "
                Next m
            End If
            sql = Mid(sql, 1, Len(sql) - 2)

            
            
            sql = sql & " into tmpTodasFormas"
            sql = sql & " from tmp1 "
            If Not rsforma.EOF Then
                For m = 0 To rsforma.RecordCount - 1
                    sql = sql & " left join  tmpForma" & m & " on tmp1.ruc = tmpForma" & m & ".ruc "
                Next m
            End If
            
            
            
            VerificaExistenciaTablaTemp "tmpTodasFormas"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            
            'sql = " select case when idtipodocumento =5 then '9999999999999' else ruc end as ruc,  "
            sql = " select ruc,  "
            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
                sql = sql & "substring(gnt.opcion2,122,1) AS BandElect, "
            Else
                sql = sql & " 'N' AS BandElect, "
            End If
            
'            substring(gnt.opcion2,122,1) AS BandElect, "
            sql = sql & " AnexoCodTipoComp, case when AnexoCodTipoComp=4 then sum(valor) *-1 else  sum(valor) end as PorcentajeIVA"
            sql = sql & " into tmpIVAS"
            sql = sql & " from gncomprobante g "
            sql = sql & " inner join GnTrans gnt on g.codtrans=GnT.codtrans "
            sql = sql & " inner join pcprovcli pc on g.idclienteref=pc.idprovcli"
            sql = sql & " inner join ivkardexrecargo i"
            sql = sql & " inner join ivrecargo iv"
            sql = sql & " on i.idrecargo=iv.idrecargo"
            sql = sql & " on g.transid=i.transid"
            sql = sql & " Where g.estado <>3"
            sql = sql & " AND (G.fechatrans BETWEEN " & _
                FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = sql & " AND g.CodTrans IN (" & PreparaCadena(.CodGrupo) & ")"
            
            sql = sql & " and codrecargo='IVA'"
            sql = sql & " group by ruc, AnexoCodTipoComp, idtipodocumento , substring(gnt.opcion2,122,1)"
            
            
            
            VerificaExistenciaTablaTemp "tmpIVAS"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            sql = " select "
            sql = sql & " RUC , AnexoCodTipoComp, BandElect, "
            sql = sql & "  sum(PorcentajeIVA) as PorcentajeIVA "
            sql = sql & " into tmpIVA"
            sql = sql & " from tmpIVAS "
            sql = sql & " group by RUC , AnexoCodTipoComp, "
            sql = sql & " AnexoCodTipoComp, BandElect "
            
            
            VerificaExistenciaTablaTemp "tmpIVA"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            ' union de las tres tablas
            sql = "SELECT DISTINCT " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB) & " as Fecha, "
            sql = sql & "  CodTipoDocumento,  '' as i,  (a.RUC)as ruc, BandRelacionado, a.AnexoCodTipoComp, isnull(a.cantTrans,0) as cant,"
'            sql = sql & " isnull(a.Valor0*-1,0) , isnull(a.Valor12*-1,0), isnull(a.ValorNO12*-1,0) "
            sql = sql & " case when a.AnexoCodTipoComp=4 then  isnull(a.Valor0,0) else isnull(a.Valor0*-1,0) end as Valor0,"
            sql = sql & " case when a.AnexoCodTipoComp=4 then  isnull(a.Valor12,0) else isnull(a.Valor12*-1,0) end as Valor12,"
            sql = sql & " case when a.AnexoCodTipoComp=4 then isnull(a.ValorNO12,0) else  isnull(a.ValorNO12*-1,0) end as ValorNO12"
            'sql = sql & " ,round(b.totaliva,2) as totaliva, I.porcentajeiva,  "
            sql = sql & " ,round(b.totaliva,2) as totaliva, I.porcentajeiva,  "
            sql = sql & " isnull(TotalICE,0) as ice,"
            sql = sql & "  case when ( a.CompElect ='S'  and  a.Prod ='N' ) then 'E' else 'F' end as TipoEmi,  "
'            sql = sql & rsforma.RecordCount & " as Numpagos,"
            
            For m = 0 To rsforma.RecordCount - 1
                sql = sql & " Codforma" & m & ",  "
            Next m
            
            For m = rsforma.RecordCount To 9
                sql = sql & "'' as forma" & m & ",  "
            Next m
            
            sql = Mid(sql, 1, Len(sql) - 3)
            
            sql = sql & " into tvta"
            sql = sql & " From tmp1 a  left  join tan1 b on a.ruc = b.ruc and a.CompElect = b.BandElect  "
            If rsforma.RecordCount > 0 Then
                sql = sql & " inner join tmpTodasFormas tf on a.ruc= tf.ruc and a.CompElect = tf.BandElect "
            End If
            sql = sql & " left join tmpICE tice on a.ruc = tice.ruc and a.CompElect= tice.CompElect   and  a.Prod  = tice.Prod and a.CompElect = tice.CompElect "
            sql = sql & " left join tmpIVA i on a.ruc=i.ruc and a.AnexoCodTipoComp=i.AnexoCodTipoComp and a.CompElect = i.BandElect "
            sql = sql & " union all "
            'saca retenciones de este periodo afectadas a facturas de otro periodo
            sql = sql & " Select  DISTINCT " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB) & " as Fecha, "
            sql = sql & " atd.CodTipoDocumento, '' as i, (pc.RUC)as ruc, pc.BandRelacionado, '18' as AnexoCodTipoComp, 0 as cantTrans,"
            sql = sql & "0 as Valor0, 0 as Valor12 , 0 as ValorNO12 ,0 as totaliva, 0 AS  porcentajeiva, "
            sql = sql & " 0 as ice,"
            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
                sql = sql & "  'E' as CompElect, "
            Else
                sql = sql & "  'F' as CompElect, "
            End If
 '           sql = sql & " 0 as Numpagos,"
            For m = 0 To 9
                sql = sql & " '' as Codforma" & m & ",  "
            Next m
            sql = Mid(sql, 1, Len(sql) - 3)
            
            sql = sql & " From    (gncomprobante GNC_b INNER JOIN PCPROVClI pc"
            sql = sql & " left join anexo_TipoDocumento atd on pc.IdTipoDocumento=atd.Id"
            sql = sql & " ON GNC_b.IdClienteRef=pc.idprovcli   )"
            sql = sql & " inner join tskardexret ivr on gnc_b.transid = ivr.transid"
            sql = sql & " where GNC_b.CodTrans IN (" & PreparaCadena(gobjCond.CodRetencion2) & ")"
            sql = sql & " AND (GNC_b.FechaTrans  BETWEEN " & _
                FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = sql & " and GNC_B .Estado<>3 and GNC_b.IdClienteRef<>0 "
            sql = sql & " and pc.ruc not in (select  ruc from tmp4) "
            'sql = sql & " group by atd.CodTipoDocumento"
            sql = sql & " order by a.AnexoCodTipoComp " ', a.Nombre
            
            VerificaExistenciaTablaTemp "tvta"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            sql = " select pc.ruc, '18' as AnexoCodTipoComp, "
            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
                sql = sql & "  'E' as CompElect, "
            Else
                sql = sql & "  'F' as CompElect, "
            End If
            sql = sql & " case when bandiva=1 then sum(base) else 0 end as BaseIVA ,"
            sql = sql & " case when bandiva=1 then sum(debe) else 0 end as ValorRetIVA,"
            sql = sql & " case when bandiva<>1 then sum(base) else 0 end as BaseRenta ,"
            sql = sql & " case when bandiva<>1 then sum(debe) else 0 end as ValorRetRenta"
            sql = sql & " Into treten"
            sql = sql & " from gncomprobante g"
            sql = sql & " inner join pcprovcli pc on g.idclienteref= pc.idprovcli"
            sql = sql & " inner join tskardexret tsk"
            sql = sql & " inner join tsretencion t"
            sql = sql & " on t.IdRetencion=tsk.IdRetencion on g.transid=tsk.transid"
            sql = sql & " Where g.Estado <> 3"
            sql = sql & " and bandcliente=1"
            sql = sql & " and bandiva=1"
                        sql = sql & " AND (G.FechaTrans  BETWEEN " & _
                            FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                            FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            
            sql = sql & " group by pc.ruc, bandIVA"
            sql = sql & " Union All"
            sql = sql & " select pc.ruc, '18' as AnexoCodTipoComp, "
            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
                sql = sql & "  'E' as CompElect, "
            Else
                sql = sql & "  'F' as CompElect, "
            End If
            sql = sql & " case when bandiva=1 then sum(base) else 0 end as BaseIVA ,"
            sql = sql & " case when bandiva=1 then sum(debe) else 0 end as ValorRetIVA,"
            sql = sql & " case when bandiva<>1 then sum(base) else 0 end as BaseRenta ,"
            sql = sql & " case when bandiva<>1 then sum(debe) else 0 end as ValorRetRenta"
            sql = sql & " from gncomprobante g"
            sql = sql & " inner join pcprovcli pc on g.idclienteref= pc.idprovcli"
            sql = sql & " inner join tskardexret tsk"
            sql = sql & " inner join tsretencion t"
            sql = sql & " on t.IdRetencion=tsk.IdRetencion"
            sql = sql & " on g.transid=tsk.transid"
            sql = sql & " Where g.Estado <> 3"
            sql = sql & " and bandcliente=1"
            sql = sql & " and bandiva=0"
                        sql = sql & " AND (G.FechaTrans  BETWEEN " & _
                            FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                            FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            
            sql = sql & " group by pc.ruc, bandIVA"
            
            VerificaExistenciaTablaTemp "treten"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            sql = "select ruc, AnexoCodTipoComp, sum(BaseIVA) as BaseIVA , sum(ValorRetIVA) as ValorRetIVA, sum(BaseRenta) as BaseRenta, sum(ValorRetRenta) as ValorRetRenta, CompElect "
            sql = sql & " Into tret"
            sql = sql & " from treten group by ruc, AnexoCodTipoComp, CompElect "
            
            VerificaExistenciaTablaTemp "tret"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
        sql = " select v.*, BaseIVA, ValorRetIVA, BaseRenta, ValorRetRenta  from tvta v full join  tret r on v.ruc=r.ruc and r.CompElect = v.TipoEmi"
        sql = sql & " WHERE V.AnexoCodTipoComp = (18) "
        sql = sql & " UNION ALL"
        sql = sql & " select v.*, 0 as BaseIVA, 0 as ValorRetIVA, 0 as BaseRenta, 0 as ValorRetRenta  from tvta v full join  tret r on v.ruc=r.ruc and r.CompElect = v.TipoEmi"
        sql = sql & " WHERE V.AnexoCodTipoComp = (5) "
        sql = sql & " UNION ALL"
        sql = sql & " select v.*,  '' AS BaseIVA,''  AS ValorRetIVA, '' AS BaseRenta, '' AS ValorRetRenta  from tvta V "
        sql = sql & " WHERE V.AnexoCodTipoComp=4 "

        
        
            
    End With
    rs.Open sql, gobjMain.EmpresaActual.Coneccion, adOpenStatic, adLockReadOnly
    UltimoRecordset rs
    Set ConsANVentas2016paraXML = rs
    Set rs = Nothing
    gobjMain.EmpresaActual.GrabaGNLogAccion "CONSULTA", "Cons Ventas 2016 para Anexos " & gobjCond.fecha1 & "-" & gobjCond.fecha2, "AN"
    VerificaExistenciaTabla 0
    VerificaExistenciaTabla 1
    VerificaExistenciaTabla 2
    VerificaExistenciaTabla 3
    VerificaExistenciaTabla 4
End Function


Public Function ListaIVRecetasFlexDesc(BandRS As Boolean, Valida As Boolean, desc As String) As Variant
    Dim sql As String, rs As Recordset
    Dim cond As String
    sql = "Select codReceta,descripcion "
    sql = sql & " from  IVReceta "
    sql = sql & "Where descripcion like '%" & desc & "%'"
    If Valida Then
        sql = sql & "And bandvalida = 1 "
    End If
    sql = sql & " order by id desc"
   Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If BandRS Then
        Set ListaIVRecetasFlexDesc = rs
    Else
        ListaIVRecetasFlexDesc = MiGetRows(rs)
        rs.Close
    End If
    Set rs = Nothing
End Function

Public Function ListaTransxDevol(CodTrans As String, CodForma As String, coddoc As String, fdesde As Date, fhasta As Date) As Variant
End Function

Public Function ConsANCompras2016ParaXML() As Variant
    Dim pck As PCKardex, pcd As PCDocAsignado
    Dim objGnComp As GNComprobante, objGnCompDocOriginal As GNComprobante
    Dim sql As String, cond As String, CadenaValores As String
    Dim Recargo As String 'CadenaAgrupa  As String
    Dim v As Variant, max As Integer, i As Integer, j As Integer, X As Integer
    Dim from As String, numReg As Long
    Dim rs As Recordset, Retenidos As Variant, NumRet As Long, v1 As Variant
    Dim Servicios As String, Bienes As String, RowRet As Long
    Dim codRecarBaseCero As String, codRecarBaseNoCero As String, rsRecargo As Recordset
    Const COL_BASESER = 15
    'segunda parte de la consulta: carga los valores de  Detalles de Anexos
    Const COL_FECHA = 2
    Const COL_CODCREDTRIB = 7
    Const COL_TRANSID = 1
    Const COL_TRANSRET = 5
    Dim condServ As String
    Dim condBien As String
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    With gobjCond
            cond = " AND (GNC.FechaTrans  BETWEEN " & _
                    FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                    FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            VerificaExistenciaTabla 2
            VerificaExistenciaTabla 3
            VerificaExistenciaTabla 4
            VerificaExistenciaTabla 5
            VerificaExistenciaTabla 6
            VerificaExistenciaTabla 10
            VerificaExistenciaTabla 11
            VerificaExistenciaTabla 12
            VerificaExistenciaTabla 13
            VerificaExistenciaTabla 14
            VerificaExistenciaTabla 15
            VerificaExistenciaTabla 16
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & " WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(.Servicios) & ") "
            sql = sql & " AND GNC.CodTrans IN (" & PreparaCadena(.CodTrans) & ")"
            sql = sql & " AND ivr.CodRecargo <> 'ICE'"
            sql = sql & "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp10 " & _
                    "From IvRecargo ivR inner join " & _
                        "AFKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(.Servicios) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(.CodTrans) & ")" & _
                  "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            'jeaa 09/01/2007
            'datos ICE
            sql = " Select Ivkr.TransID, "
            sql = sql & " (ivkr.Valor) AS BASEice,"
            sql = sql & " 'ICE' AS CodICE, (ivkr.Valor) As ValorICE"
            sql = sql & " Into tmp6"
            sql = sql & " From IvRecargo ivR"
            sql = sql & " inner join IvKardexRecargo ivkR"
            sql = sql & " Inner join GnComprobante gNc"
            sql = sql & " On ivkr.TransID = gNc.TransID"
            sql = sql & " On Ivr.IdRecargo = IvkR.IdRecargo"
            sql = sql & " Where gnc.Estado <> 3"
            'sql = sql & " AND ivr.CodRecargo IN ('" & gobjMain.EmpresaActual.GNOpcion.ObtenerValor("RecDescParaICE") & "')" & cond
            sql = sql & " AND ivr.CodRecargo = 'ICE'" & cond
            sql = sql & " AND GNC.CodTrans IN (" & PreparaCadena(.CodTrans) & ")"
            sql = sql & " and ivkr.Valor<>0"
            'sql = sql & " Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg






'--datos de indices
            VerificaExistenciaTabla 3
            sql = " Select  DISTINCT GNC.TransID AS cp_id, GNC_b.TransID AS rtp_id"
            sql = sql & " Into tmp3"
            sql = sql & " From    (gncomprobante GNC Inner join pckardex pck_a ON GNC.TransID = pck_a.TransID ) INNER JOIN"
            sql = sql & "  (gncomprobante GNC_b Inner join pckardex pck_B ON GNC_b.TransID = pck_B.TransID )"
            sql = sql & " ON  PCK_A.ID = PCK_B.IDASIGNADO"
            sql = sql & " AND GNC.CodTrans IN (" & PreparaCadena(gobjCond.CodTrans) & ")"
            sql = sql & " AND GNC_b.CodTrans IN (" & PreparaCadena(gobjCond.CodRetencion1) & ")"
            sql = sql & " AND (GNC.FechaTrans  BETWEEN " & _
                FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = sql & " and GNC.Estado<>3 and GNC_b.estado<>3 "
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            sql = " Select  DISTINCT GNC.TransID AS cp_id, GNC_b.TransID AS rtp_id"
            sql = sql & " Into tmp13"
            sql = sql & " From    (gncomprobante GNC Inner join pckardex pck_a ON GNC.TransID = pck_a.TransID ) INNER JOIN"
            sql = sql & "  (gncomprobante GNC_b Inner join pckardex pck_B ON GNC_b.TransID = pck_B.TransID )"
            sql = sql & " ON  PCK_A.ID = PCK_B.IDASIGNADO"
            sql = sql & " AND GNC.CodTrans IN (" & PreparaCadena(gobjCond.CodTrans) & ")"
            sql = sql & " AND GNC_b.CodTrans IN (" & PreparaCadena(gobjCond.CodRetencion1) & ")"
            sql = sql & " AND (GNC.FechaTrans  BETWEEN " & _
                FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = sql & " and GNC.Estado<>3 and GNC_b.estado<>3 "
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
         sql = "select codrecargo, BandBaseCero from ivrecargo where BandAfectaBaseAnexo=1  "
         Set rsRecargo = gobjMain.EmpresaActual.OpenRecordset(sql)
         
         If Not rsRecargo.EOF Then
            rsRecargo.MoveFirst
        End If
         
        codRecarBaseCero = ""
        codRecarBaseNoCero = ""
         
         For i = 1 To rsRecargo.RecordCount
            If rsRecargo.RecordCount > 0 Then
               If rsRecargo.Fields(1) Then
                   codRecarBaseCero = codRecarBaseCero & "'" & rsRecargo.Fields(0) & "',"
               Else
                   codRecarBaseNoCero = codRecarBaseNoCero & "'" & rsRecargo.Fields(0) & "',"
               End If
            End If
            rsRecargo.MoveNext
         Next i
        
        If Len(codRecarBaseCero) > 0 Then
            codRecarBaseCero = Mid$(codRecarBaseCero, 1, Len(codRecarBaseCero) - 1)
        End If
            
        If Len(codRecarBaseNoCero) > 0 Then
            codRecarBaseNoCero = Mid$(codRecarBaseNoCero, 1, Len(codRecarBaseNoCero) - 1)
        End If
            
        'If Len(codRecarBaseCero) > 0 Then
        sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalRecargoBaseCero "
        sql = sql & " Into tmpRecarBaseCero "
        sql = sql & " From IvRecargo ivR inner join "
        sql = sql & " IvKardexRecargo ivkR Inner join "
        sql = sql & " GnComprobante gNc  "
        sql = sql & " On ivkr.TransID = gNc.TransID "
        sql = sql & " On Ivr.IdRecargo = IvkR.IdRecargo "
        sql = sql & " WHERE gNc.Estado <> 3 "
        If Len(codRecarBaseCero) > 0 Then
            sql = sql & " AND ivr.CodRecargo IN (" & codRecarBaseCero & ") " & cond
        Else
            sql = sql & " AND ivr.CodRecargo IN ('xxxx') " & cond
        End If
        sql = sql & " AND GNC.CodTrans IN (" & PreparaCadena(.CodTrans) & ")"
        sql = sql & " Group by IvkR.TransID"
        
        VerificaExistenciaTablaTemp "tmpRecarBaseCero"
        gobjMain.EmpresaActual.EjecutarSQL sql, numReg
        
            
            
            
            '--datos de la compra
            sql = " Select  GNC.TransID, Ane.FechaAnexos, GNC.FechaTrans, CodTipoDocumento, pcprovcli.ruc, "
            sql = sql & " pcprovcli.Nombre, Ane.CodCredTrib, Ane.CodTipoComp, Ane.NumAutSRI,"
            sql = sql & " Ane.NumSerie, Ane.NumSecuencial, GNC.Descripcion, CodFormaPagoSRI, BandPagoExterior, CodPaisSRI , BandDobleTributa, BandPagoSujRet, "
            sql = sql & " Case Ane.BandDevolucion WHEN 0 THEN 'N' ELSE 'S' END AS BandDev,"
            sql = sql & " Gnc.CodTrans+' '+CONVERT(varchar,Gnc.NumTrans) as Trans,"
            sql = sql & " Gnc.CodTrans as Transa, Gnc.NumTrans as NumTransa,"
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) )  end As Valor0, "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " ( (vw.SignoCompra * (vw.CostoTotalBaseIVA + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))))) +  isnull(tmp6.BASEice,0)) end AS Valor12, "
            sql = sql & " Case vw.CostoTotalBaseNOIVA When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBaseNOIVA + (vw.CostoTotalBaseNOIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS ValorNOIVA, "
            sql = sql & " Case vw.CostoTotal When 0 then 0 else "
            sql = sql & " vw.SignoCompra  * (vw.CostoTotal + (vw.CostoTotal * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))))  End AS CostoTotal,  "
            sql = sql & " ane.NumSerieEstablecimiento ,ane.NumSeriePunto ,ane.FechaCaducidad, "
            sql = sql & " tmp6.BASEice, tmp6.CodICE ,tmp6.ValorICE,BANDREToTRO, idproveedorref, GNc.porcentajeiva, pcprovcli.BandRelacionado   "
            sql = sql & " Into tmp1"
            sql = sql & " from    (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " inner join gntrans gnt"
            sql = sql & " inner join GNTransRecargo gntr"
            sql = sql & " inner join ivrecargo ivr"
            sql = sql & " on ivr.idrecargo=gntr.idrecargo"
            sql = sql & " on gnt.codtrans= gntr.codtrans"
            sql = sql & " on gnc.codtrans=gnt.codtrans"
            sql = sql & " left join tmp6 on gnc.transid=tmp6.transid"
            
            'sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
              If gobjMain.EmpresaActual.GNOpcion.IVKTipoDatoDouble Then
                sql = sql & " inner join vwConsSUMIVKardexIVADou vw ON Gnc.TransID = vw.transid "
            Else
                sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            End If
            
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  "
            sql = sql & " left join anexo_TipoDocumento atd on pcprovcli.IdTipoDocumento=atd.Id"
            sql = sql & " on gnc.IdProveedorRef=pcprovcli.idprovcli"
            

            
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(gobjCond.CodTrans) & ")"
            sql = sql & " and GNC.Estado<>3 " & cond
            
            sql = sql & " union all"
            '--datos de la compra activos
            sql = sql & " Select  GNC.TransID, Ane.FechaAnexos, GNC.FechaTrans, CodTipoDocumento, pcprovcli.ruc,"
            sql = sql & " pcprovcli.Nombre, Ane.CodCredTrib, Ane.CodTipoComp, Ane.NumAutSRI,"
            sql = sql & " Ane.NumSerie, Ane.NumSecuencial, GNC.Descripcion, CodFormaPagoSRI, BandPagoExterior, CodPaisSRI , BandDobleTributa, BandPagoSujRet, "
            sql = sql & " Case Ane.BandDevolucion WHEN 0 THEN 'N' ELSE 'S' END AS BandDev,"
            sql = sql & " Gnc.CodTrans+' '+CONVERT(varchar,Gnc.NumTrans) as Trans,"
            sql = sql & " Gnc.CodTrans as Transa, Gnc.NumTrans as NumTransa,"
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0, "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBaseIVA + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS Valor12, "
            sql = sql & " 0 as ValorNOIVA, "
            sql = sql & " Case vw.CostoTotal When 0 then 0 else "
            sql = sql & " vw.SignoCompra  * (vw.CostoTotal + (vw.CostoTotal * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))))  End AS CostoTotal,  "
            sql = sql & " ane.NumSerieEstablecimiento ,ane.NumSeriePunto ,ane.FechaCaducidad, "
            sql = sql & " 0 as BASEice, '0' as CodICE , 0 as ValorICE,BANDREToTRO, idproveedorref, GNc.porcentajeiva, pcprovcli.BandRelacionado  "
            sql = sql & " from    (( tmp10 Right join gncomprobante Gnc "
            sql = sql & " inner join gntrans gnt"
            sql = sql & " inner join GNTransRecargo gntr"
            sql = sql & " inner join ivrecargo ivr"
            sql = sql & " on ivr.idrecargo=gntr.idrecargo"
            sql = sql & " on gnt.codtrans= gntr.codtrans"
            sql = sql & " on gnc.codtrans=gnt.codtrans"
            sql = sql & " inner join vwConsSUMAFKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp10.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join pcprovcli  "
            sql = sql & " left join anexo_TipoDocumento atd on pcprovcli.IdTipoDocumento=atd.Id"
            sql = sql & " on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(gobjCond.CodTrans) & ")"
            sql = sql & " and GNC.Estado<>3 " & cond
            
            
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            '-- datos de las retenciones
            ' union de las tres tablas
            ' union de las tres tablas
            sql = " select  DISTINCT  a.Transid,a.FechaAnexos, a.FechaTrans, a.Transa,  a.NumTransa, "
            sql = sql & "a.CodTipoDocumento, a.idproveedorref, a.RUC, a.Nombre, a.BandRelacionado,"
            sql = sql & " a.NumSerieEstablecimiento , a.NumSeriePunto, "
            sql = sql & " case when len(a.NumSecuencial)>9 then right(a.NumSecuencial,9) else a.NumSecuencial end as NumSecuencial ,  "
            sql = sql & " a.NumAutSRI, a.FechaCaducidad, "
            sql = sql & "a.CodCredTrib, a.CodTipoComp, a.CodFormaPagoSRI, a.BandPagoExterior, a.CodPaisSRI , a.BandDobleTributa, a.BandPagoSujRet, "
            sql = sql & " a.Valor0+ (isnull((tr.TotalRecargoBaseCero),0))  , a.Valor12 AS Valor12 ,a.ValorNOIVA, CodICE, ISNULL(ValorICE,0) AS ValorICE, BANDREToTRO, (a.porcentajeiva) , (isnull((tr.TotalRecargoBaseCero),0))"
            sql = sql & " From tmp1 a  "
            sql = sql & " left join tmpRecarBaseCero tr ON a.TransID = tr.TransID"
            sql = sql & " order by a.CodTipoComp"
    End With
    rs.Open sql, gobjMain.EmpresaActual.Coneccion, adOpenStatic, adLockReadOnly
    UltimoRecordset rs
    Set ConsANCompras2016ParaXML = rs
    Set rs = Nothing
    gobjMain.EmpresaActual.GrabaGNLogAccion "CONSULTA", "Cons Compraspara Anexos 2016 XML" & gobjCond.fecha1 & "-" & gobjCond.fecha2, "AN"
    VerificaExistenciaTabla 0
    VerificaExistenciaTabla 1
    VerificaExistenciaTabla 2
    VerificaExistenciaTabla 3
    VerificaExistenciaTabla 4
End Function


Public Function ConsANVentasxEstablecimiento2016paraXML() As Variant
    Dim pck As PCKardex, pcd As PCDocAsignado
    Dim objGnComp As GNComprobante, objGnCompDocOriginal As GNComprobante
    Dim sql As String, cond As String, CadenaValores As String
    Dim Recargo As String 'CadenaAgrupa  As String
    Dim v As Variant, max As Integer, i As Integer, j As Integer, X As Integer
    Dim from As String, numReg As Long, f1 As String
    Dim rs As Recordset, Retenidos As Variant, NumRet As Long, v1 As Variant
    Dim Servicios As String, Bienes As String, RowRet As Long
    Dim codRecarBaseCero As String, codRecarBaseNoCero As String, rsRecargo As Recordset
    Const COL_BASESER = 15
    'segunda parte de la consulta: carga los valores de  Detalles de Anexos
    Const COL_FECHA = 2
    Const COL_CODCREDTRIB = 7
    Const COL_TRANSID = 1
    Const COL_TRANSRET = 5
    Dim condServ As String
    Dim condBien As String
    Dim Vista As String
    Dim pos As Integer
    Dim pos1 As Integer
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
   With gobjCond
        'Reporte de un mes a la vez
        f1 = DateSerial(Year(.fecha1), Month(.fecha1), 1)
        cond = " AND GNC.FechaTrans BETWEEN " & FechaYMD(f1, gobjMain.EmpresaActual.TipoDB) & _
               " AND " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB)
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            VerificaExistenciaTabla 2
            VerificaExistenciaTabla 3
            pos = InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "DISFRA")
            pos1 = InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "ENCOFRADO")
            If gobjMain.EmpresaActual.GNOpcion.IVKTipoDatoDouble Then
                Vista = "vwConsSUMIVKardexIVADou"
            Else
               If pos > 0 Or pos1 > 0 Then
                    Vista = "vwConsSUMIVKardexIVA2B"
                Else
                    Vista = "vwConsSUMIVKardexIVA"
                End If
            End If

               sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
                sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(.Sucursal) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(.CodGrupo) & ")" & _
                  "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
        VerificaExistenciaTablaTemp "tmpICE"
        sql = "Select "
        sql = sql & " gns.numestablecimiento, AnexoCodTipoComp,"
        sql = sql & " SUM(IvKr.Valor) as TotalICE "
        sql = sql & " Into tmpICE "
        sql = sql & "From IvRecargo ivR inner join "
        sql = sql & " IvKardexRecargo ivkR Inner join "
        sql = sql & " GnComprobante gNc  "
        sql = sql & " inner join GnTrans gnt "
        sql = sql & " inner join gnsucursal gns on gnt.idsucursal=gns.idsucursal"
        sql = sql & " on gnc.codtrans=GnT.codtrans"
        sql = sql & " On ivkr.TransID = gNc.TransID "
        sql = sql & " On Ivr.IdRecargo = IvkR.IdRecargo "
        sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo = 'ICE' " & cond & " AND GNC.CodTrans IN (" & PreparaCadena(.CodGrupo) & ")"
        
        If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
            sql = sql & " and   substring(gnt.opcion2,122,1) ='N'  and  substring(gnt.opcion2,135,1) ='S'"
        End If
        
        sql = sql & " Group by gns.numestablecimiento, AnexoCodTipoComp"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
'        If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
'            sql = sql & " , substring(gnt.opcion2,122,1) ,  substring(gnt.opcion2,135,1) "
'        End If
            
            
         sql = "select codrecargo, BandBaseCero from ivrecargo where BandAfectaBaseAnexo=1  "
         Set rsRecargo = gobjMain.EmpresaActual.OpenRecordset(sql)
         
         rsRecargo.MoveFirst
        codRecarBaseCero = ""
        codRecarBaseNoCero = ""
         
         For i = 1 To rsRecargo.RecordCount
            If rsRecargo.RecordCount > 0 Then
               If rsRecargo.Fields(1) Then
                   codRecarBaseCero = codRecarBaseCero & "'" & rsRecargo.Fields(0) & "',"
               Else
                   codRecarBaseNoCero = codRecarBaseNoCero & "'" & rsRecargo.Fields(0) & "',"
               End If
            End If
            rsRecargo.MoveNext
         Next i
        
        If Len(codRecarBaseCero) > 0 Then
            codRecarBaseCero = Mid$(codRecarBaseCero, 1, Len(codRecarBaseCero) - 1)
        End If
            
        If Len(codRecarBaseNoCero) > 0 Then
            codRecarBaseNoCero = Mid$(codRecarBaseNoCero, 1, Len(codRecarBaseNoCero) - 1)
        End If
            
            
            
            
'--datos de indices
            VerificaExistenciaTabla 3
            sql = " Select  DISTINCT GNC.IdClienteRef AS FC_id, GNC_b.IdClienteRef AS rtp_id1, "
            sql = sql & " CASE WHEN GNC_b.IdClienteRef=0 THEN PCK_b.IdProvcli WHEN GNC_b.IdClienteRef<>0 THEN GNC_b.IdClienteRef end as rtp_id "
            sql = sql & " Into tmp3"
            sql = sql & " From    (gncomprobante GNC Inner join pckardex pck_a ON GNC.TransID = pck_a.TransID ) INNER JOIN"
            sql = sql & "  (gncomprobante GNC_b Inner join pckardex pck_B ON GNC_b.TransID = pck_B.TransID )"
            sql = sql & " ON  PCK_A.ID = PCK_B.IDASIGNADO"
            sql = sql & " where GNC.CodTrans IN (" & PreparaCadena(gobjCond.CodGrupo) & ")"
            sql = sql & " AND GNC_b.CodTrans IN (" & PreparaCadena(gobjCond.CodRetencion2) & ")"
            sql = sql & " AND (GNC_b.FechaTrans  BETWEEN " & _
                FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = sql & " and GNC.Estado<>3 and GNC.IdClienteRef<>0 and GNC_b.IdClienteRef<>0 "
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            '--datos de la Venta
            
            
            sql = " Select gns.numestablecimiento, AnexoCodTipoComp, "
            sql = sql & " count(Gnc.NumTrans) as cantTrans,  "
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN (sum( vw.PrecioRealTotalBase0) *-1 + (isnull(sum(tr.TotalRecargoBaseCero),0)*-1)) END As Valor0,"
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotalBaseIVA) *-1 END AS Valor12,"
            sql = sql & " 0 AS ValorNO12,"
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN (sum( vw.PrecioRealTotal) *-1  + (isnull(sum(tr.TotalRecargoBaseCero),0)*-1) )END  AS PrecioRealTotal"
            
            
            sql = sql & " Into tmp1"
            sql = sql & " from    ("
            sql = sql & " tmp0"
            sql = sql & " Right join gncomprobante Gnc "
            sql = sql & " inner join GnTrans gnt "
            sql = sql & " inner join gnsucursal gns on gnt.idsucursal=gns.idsucursal"
            sql = sql & " on gnc.codtrans=GnT.codtrans"
            
            sql = sql & " inner join " & Vista & " vw"

            sql = sql & " ON Gnc.TransID = vw.transid"
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " left join tmpRecarBaseCero tr ON tmp0.TransID = tr.TransID"
            sql = sql & " where  gns.bandvalida=1 and GNC.CodTrans IN (" & PreparaCadena(gobjCond.CodGrupo) & ")"
            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
                sql = sql & " and   substring(gnt.opcion2,122,1) ='N'  and  substring(gnt.opcion2,135,1) ='S'"
            End If
            sql = sql & " and GNC.Estado<>3  AND Gnc.BandCierre<>9 " & cond
            sql = sql & " group by  gns.numestablecimiento, AnexoCodTipoComp "
            
            VerificaExistenciaTabla 4
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            ' union de las tres tablas
            sql = "SELECT DISTINCT  gns.numestablecimiento, "
            sql = sql & "   a.AnexoCodTipoComp, isnull(a.cantTrans,0),"
            sql = sql & " case when a.AnexoCodTipoComp=4 then isnull(a.Valor0,0) else isnull(a.Valor0*-1,0) end as Valor0, "
            sql = sql & " case when a.AnexoCodTipoComp=4 then isnull(a.Valor12,0) else isnull(a.Valor12*-1,0) end as Valor12, "
            sql = sql & " case when a.AnexoCodTipoComp=4then  isnull(a.ValorNO12,0) else isnull(a.ValorNO12*-1,0) end as ValorNO12, "
            sql = sql & "  isnull(TotalICE,0) as ICE, "
            sql = sql & " case when a.AnexoCodTipoComp=4 then  isnull(a.PrecioRealTotal,0) else isnull(a.PrecioRealTotal*-1,0) end as PrecioRealTotal "
            sql = sql & " From gnsucursal gns left join tmp1 a on gns.numestablecimiento  = a.numestablecimiento "
            sql = sql & " left join tmpICE i on gns.numestablecimiento = i.numestablecimiento and a.AnexoCodTipoComp = i.AnexoCodTipoComp "
            sql = sql & " Where gns.BandValida = 1"
            sql = sql & " order by gns.numestablecimiento, a.AnexoCodTipoComp desc"
            
    End With
    rs.Open sql, gobjMain.EmpresaActual.Coneccion, adOpenStatic, adLockReadOnly
    UltimoRecordset rs
    Set ConsANVentasxEstablecimiento2016paraXML = rs
    Set rs = Nothing
    gobjMain.EmpresaActual.GrabaGNLogAccion "CONSULTA", "Cons Ventas Estab Anexos " & gobjCond.fecha1 & "-" & gobjCond.fecha2, "AN"
    VerificaExistenciaTabla 0
    VerificaExistenciaTabla 1
    VerificaExistenciaTabla 2
    VerificaExistenciaTabla 3
    VerificaExistenciaTabla 4
End Function

Public Function ListaIVRecetasGrid(BandRS As Boolean, Valida As Boolean) As Variant
    Dim sql As String, rs As Recordset
    Dim cond As String
    Dim s As String
    sql = "Select codReceta,descripcion "
    sql = sql & " from  IVReceta "
    If Valida Then
        sql = sql & "Where bandvalida = 1 "
    End If
    sql = sql & " order by id desc"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not rs.EOF Then s = "|" & rs.GetString(adClipString, , vbTab, "|")
    ListaIVRecetasGrid = s
    rs.Close
    Set rs = Nothing
End Function

Public Function ListaIVRecetasGridDesc(BandRS As Boolean, Valida As Boolean, desc As String) As Variant
    Dim sql As String, rs As Recordset
    Dim cond As String
    Dim s As String
    sql = "Select codReceta,descripcion "
    sql = sql & " from  IVReceta "
    sql = sql & "Where descripcion like '%" & desc & "%'"
    If Valida Then
        sql = sql & "And bandvalida = 1 "
    End If
    sql = sql & " order by id desc"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not rs.EOF Then s = "|" & rs.GetString(adClipString, , vbTab, "|")
    ListaIVRecetasGridDesc = s
    rs.Close
    Set rs = Nothing
End Function


Public Function ListaGNTransCompElect( _
                    Modulo As String, _
                    BandInvisibles As Boolean, _
                    BandRS As Boolean) As Variant
    Dim sql As String, cond As String, rs As Recordset
    
    sql = "SELECT CodTrans, NombreTrans , "
    sql = sql & " substring(opcion2,122,1), CASE WHEN LEN(substring(opcion2,122,1))>0 THEN case WHEN substring(opcion2,135,1)='N' then 'PRODUCCION' ELSE 'PRUEBAS' END ELSE '' END "
    sql = sql & " from GNTrans"
    If Len(Modulo) > 0 Then cond = cond & "Modulo='" & Modulo & "' "
'    If Not BandInvisibles Then
'        If Len(cond) > 0 Then cond = cond & " AND "
'        cond = cond & "BandVisible=" & CadenaBool(True, TipoDB)
'    End If
    
    If Len(cond) > 0 Then sql = sql & "WHERE " & cond
    sql = sql & " ORDER BY CodTrans"
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If BandRS Then
        Set ListaGNTransCompElect = rs
    Else
        ListaGNTransCompElect = MiGetRows(rs)
        rs.Close
    End If
    Set rs = Nothing
End Function

Public Function ConsANCompras2016ConsolParaXML(cadEmpresas As String) As Variant
    Dim rs As Recordset, v As Variant, W As Variant, Y As Variant, z As Variant, j As Integer
    Dim CodT As String
    Dim emp As Variant
    Dim i As Long
    Dim codret  As String
    Dim sql As String
    Y = Split(cadEmpresas, ";")
    For j = 0 To UBound(Y)
        z = Split(Y(j), ",")
        v = Split(gobjCond.CodTrans, ";")
        CodT = ""
        emp = z(0)
        For i = 0 To UBound(v)
            W = Split(v(i), ",")
            If W(0) = emp Then
                CodT = CodT & W(1) & ","
            End If
        Next i
        
        v = Split(gobjCond.CodRetencion1, ";")
        codret = ""
        For i = 0 To UBound(v)
            W = Split(v(i), ",")
            If W(0) = emp Then
                codret = codret & W(1) & ","
            End If
        Next i
        
    CodT = Mid$(CodT, 1, Len(CodT) - 1)
    codret = Mid$(codret, 1, Len(codret) - 1)
    
    Set rs = ConsultaComprasConsolidadas2016ParaXML(z(1), gobjCond.Servicios, CodT, codret)
    Next j
    
    sql = ""
    For j = 0 To UBound(Y)
        z = Split(Y(j), ",")
        sql = sql & " select * from tmpats" & z(1)
        If j <> UBound(Y) Then
            sql = sql & " union all"
        End If
'        sql = sql & " select * from tmpats" & z(1)
'        sql = sql & " union all"
'        sql = sql & " select * from tmpats" & z(1)
    Next j
    
    sql = sql & " order by CodTipoComp, FechaAnexos, emp"
    
    rs.Open sql, gobjMain.EmpresaActual.Coneccion, adOpenStatic, adLockReadOnly
    UltimoRecordset rs
    
    
    Set ConsANCompras2016ConsolParaXML = rs
    Set rs = Nothing
End Function


Public Function ConsultaComprasConsolidadas2016ParaXML(ByVal BaseDatos As String, ByVal Servicios As String, ByVal CodTrans As String, ByVal CodRetencion1 As String) As Variant
    Dim pck As PCKardex, pcd As PCDocAsignado
    Dim objGnComp As GNComprobante, objGnCompDocOriginal As GNComprobante
    Dim sql As String, cond As String, CadenaValores As String
    Dim Recargo As String 'CadenaAgrupa  As String
    Dim v As Variant, max As Integer, i As Integer, j As Integer, X As Integer
    Dim from As String, numReg As Long
    Dim rs As Recordset, Retenidos As Variant, NumRet As Long, v1 As Variant
    Dim Servicios1 As String, Bienes As String, RowRet As Long
    Const COL_BASESER = 15
    'segunda parte de la consulta: carga los valores de  Detalles de Anexos
    Const COL_FECHA = 2
    Const COL_CODCREDTRIB = 7
    Const COL_TRANSID = 1
    Const COL_TRANSRET = 5
    Dim condServ As String
    Dim condBien As String
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    With gobjCond
            cond = " AND (GNC.FechaTrans  BETWEEN " & _
                    FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                    FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            VerificaExistenciaTabla 2
            VerificaExistenciaTabla 3
            VerificaExistenciaTabla 4
            VerificaExistenciaTabla 5
            VerificaExistenciaTabla 6
            VerificaExistenciaTabla 10
            VerificaExistenciaTabla 11
            VerificaExistenciaTabla 12
            VerificaExistenciaTabla 13
            VerificaExistenciaTabla 14
            VerificaExistenciaTabla 15
            VerificaExistenciaTabla 16
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From " & BaseDatos & ".dbo.IvRecargo ivR inner join " & _
                          BaseDatos & ".dbo.IvKardexRecargo ivkR Inner join " & _
                             BaseDatos & ".dbo.GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(Servicios) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(CodTrans) & ")" & _
                  "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp10 " & _
                    "From " & BaseDatos & ".dbo.IvRecargo ivR inner join " & _
                         BaseDatos & ".dbo.AFKardexRecargo ivkR Inner join " & _
                             BaseDatos & ".dbo.GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(Servicios) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(CodTrans) & ")" & _
                  "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            'jeaa 09/01/2007
            'datos ICE
            sql = " Select Ivkr.TransID, "
            sql = sql & " SUM(IVK.CostoRealTotal) AS BASEice,"
            sql = sql & " MAX(IVC.CodICE) AS CodICE, SUM(ivkr.Valor) As ValorICE"
            sql = sql & " Into tmp6"
            sql = sql & " From " & BaseDatos & ".dbo.IvRecargo ivR"
            sql = sql & " inner join " & BaseDatos & ".dbo.IvKardexRecargo ivkR"
            sql = sql & " INNER JOIN " & BaseDatos & ".dbo.IVKARDEX IVK"
            sql = sql & " INNER JOIN " & BaseDatos & ".dbo.IVINVENTARIO IV"
            sql = sql & " INNER JOIN " & BaseDatos & ".dbo.IVRECARGOICE IVC"
            sql = sql & " ON IV.IDICE=IVC.IDICE"
            sql = sql & " ON IVK.IDINVENTARIO=IV.IDINVENTARIO"
            sql = sql & " ON IVKR.TRANSID=IVK.TRANSID"
            sql = sql & " Inner join " & BaseDatos & ".dbo.GnComprobante gNc"
            sql = sql & " On ivkr.TransID = gNc.TransID"
            sql = sql & " On Ivr.IdRecargo = IvkR.IdRecargo"
            sql = sql & " Where gnc.Estado <> 3"
            sql = sql & " AND ivr.CodRecargo IN ('" & gobjMain.EmpresaActual.GNOpcion.ObtenerValor("RecDescParaICE") & "')" & cond
            sql = sql & " AND GNC.CodTrans IN (" & PreparaCadena(CodTrans) & ")"
            sql = sql & " Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
'--datos de indices
            VerificaExistenciaTabla 3
            sql = " Select  DISTINCT GNC.TransID AS cp_id, GNC_b.TransID AS rtp_id"
            sql = sql & " Into tmp3"
            sql = sql & " From    (" & BaseDatos & ".dbo.gncomprobante GNC Inner join " & BaseDatos & ".dbo.pckardex pck_a ON GNC.TransID = pck_a.TransID ) INNER JOIN"
            sql = sql & "  (" & BaseDatos & ".dbo.gncomprobante GNC_b Inner join " & BaseDatos & ".dbo.pckardex pck_B ON GNC_b.TransID = pck_B.TransID )"
            sql = sql & " ON  PCK_A.ID = PCK_B.IDASIGNADO"
            sql = sql & " AND GNC.CodTrans IN (" & PreparaCadena(CodTrans) & ")"
            sql = sql & " AND GNC_b.CodTrans IN (" & PreparaCadena(CodRetencion1) & ")"
            sql = sql & " AND (GNC.FechaTrans  BETWEEN " & _
                FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = sql & " and GNC.Estado<>3 and GNC_b.estado<>3 "
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            sql = " Select  DISTINCT GNC.TransID AS cp_id, GNC_b.TransID AS rtp_id"
            sql = sql & " Into tmp13"
            sql = sql & " From    (" & BaseDatos & ".dbo.gncomprobante GNC Inner join " & BaseDatos & ".dbo.pckardex pck_a ON GNC.TransID = pck_a.TransID ) INNER JOIN"
            sql = sql & "  (" & BaseDatos & ".dbo.gncomprobante GNC_b Inner join " & BaseDatos & ".dbo.pckardex pck_B ON GNC_b.TransID = pck_B.TransID )"
            sql = sql & " ON  PCK_A.ID = PCK_B.IDASIGNADO"
            sql = sql & " AND GNC.CodTrans IN (" & PreparaCadena(CodTrans) & ")"
            sql = sql & " AND GNC_b.CodTrans IN (" & PreparaCadena(CodRetencion1) & ")"
            sql = sql & " AND (GNC.FechaTrans  BETWEEN " & _
                FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = sql & " and GNC.Estado<>3 and GNC_b.estado<>3 "
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            '--datos de la compra
            sql = " Select  GNC.TransID, Ane.FechaAnexos, GNC.FechaTrans, CodTipoDocumento, pcprovcli.ruc,"
            sql = sql & " GNC.Nombre, Ane.CodCredTrib, Ane.CodTipoComp, Ane.NumAutSRI,"
            sql = sql & " Ane.NumSerie, Ane.NumSecuencial, GNC.Descripcion, CodFormaPagoSRI, BandPagoExterior, CodPaisSRI , BandDobleTributa, BandPagoSujRet, "
            sql = sql & " Case Ane.BandDevolucion WHEN 0 THEN 'N' ELSE 'S' END AS BandDev,"
            sql = sql & " Gnc.CodTrans+' '+CONVERT(varchar,Gnc.NumTrans) as Trans,"
            sql = sql & " Gnc.CodTrans as Transa, Gnc.NumTrans as NumTransa,"
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0, "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBaseIVA + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS Valor12, "
            sql = sql & " Case vw.CostoTotalBaseNOIVA When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBaseNOIVA + (vw.CostoTotalBaseNOIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS ValorNOIVA, "
            sql = sql & " Case vw.CostoTotal When 0 then 0 else "
            sql = sql & " vw.SignoCompra  * (vw.CostoTotal + (vw.CostoTotal * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))))  End AS CostoTotal,  "
            sql = sql & " ane.NumSerieEstablecimiento ,ane.NumSeriePunto ,ane.FechaCaducidad, "
            sql = sql & " tmp6.BASEice, tmp6.CodICE ,tmp6.ValorICE,BANDREToTRO, idproveedorref, GNc.porcentajeiva, pcprovcli.BandRelacionado, pcprovcli.TipoProvCli "
            sql = sql & " Into tmp1"
            sql = sql & " from    (( tmp0 Right join " & BaseDatos & ".dbo.gncomprobante Gnc "
            sql = sql & " inner join " & BaseDatos & ".dbo.gntrans gnt"
            sql = sql & " inner join " & BaseDatos & ".dbo.GNTransRecargo gntr"
            sql = sql & " inner join " & BaseDatos & ".dbo.ivrecargo ivr"
            sql = sql & " on ivr.idrecargo=gntr.idrecargo"
            sql = sql & " on gnt.codtrans= gntr.codtrans"
            sql = sql & " on gnc.codtrans=gnt.codtrans"
            sql = sql & " left join tmp6 on gnc.transid=tmp6.transid"
            sql = sql & " inner join " & BaseDatos & ".dbo.vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join " & BaseDatos & ".dbo.Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join " & BaseDatos & ".dbo.pcprovcli  pcprovcli "
            sql = sql & " left join " & BaseDatos & ".dbo.anexo_TipoDocumento atd on pcprovcli.IdTipoDocumento=atd.Id"
            sql = sql & " on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(CodTrans) & ")"
            sql = sql & " and GNC.Estado<>3 " & cond
            sql = sql & " union all"
            '--datos de la compra activos
            sql = sql & " Select  GNC.TransID, Ane.FechaAnexos, GNC.FechaTrans, CodTipoDocumento, pcprovcli.ruc,"
            sql = sql & " GNC.Nombre, Ane.CodCredTrib, Ane.CodTipoComp, Ane.NumAutSRI,"
            sql = sql & " Ane.NumSerie, Ane.NumSecuencial, GNC.Descripcion, CodFormaPagoSRI, BandPagoExterior, CodPaisSRI , BandDobleTributa, BandPagoSujRet, "
            sql = sql & " Case Ane.BandDevolucion WHEN 0 THEN 'N' ELSE 'S' END AS BandDev,"
            sql = sql & " Gnc.CodTrans+' '+CONVERT(varchar,Gnc.NumTrans) as Trans,"
            sql = sql & " Gnc.CodTrans as Transa, Gnc.NumTrans as NumTransa,"
            sql = sql & " Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end As Valor0, "
            sql = sql & " Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBaseIVA + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end AS Valor12, "
            sql = sql & " 0 as ValorNOIVA, "
            sql = sql & " Case vw.CostoTotal When 0 then 0 else "
            sql = sql & " vw.SignoCompra  * (vw.CostoTotal + (vw.CostoTotal * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))))  End AS CostoTotal,  "
            sql = sql & " ane.NumSerieEstablecimiento ,ane.NumSeriePunto ,ane.FechaCaducidad, "
            sql = sql & " 0 as BASEice, '0' as CodICE , 0 as ValorICE,BANDREToTRO, idproveedorref, GNc.porcentajeiva, pcprovcli.BandRelacionado, pcprovcli.TipoProvCli"
            sql = sql & " from    (( tmp10 Right join " & BaseDatos & ".dbo.gncomprobante Gnc "
            sql = sql & " inner join " & BaseDatos & ".dbo.gntrans gnt"
            sql = sql & " inner join " & BaseDatos & ".dbo.GNTransRecargo gntr"
            sql = sql & " inner join " & BaseDatos & ".dbo.ivrecargo ivr"
            sql = sql & " on ivr.idrecargo=gntr.idrecargo"
            sql = sql & " on gnt.codtrans= gntr.codtrans"
            sql = sql & " on gnc.codtrans=gnt.codtrans"
            sql = sql & " inner join " & BaseDatos & ".dbo.vwConsSUMAFKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp10.TransID = Gnc.TransID)"
            sql = sql & " inner join " & BaseDatos & ".dbo.Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " right join " & BaseDatos & ".dbo.pcprovcli pcprovcli "
            sql = sql & " left join " & BaseDatos & ".dbo.anexo_TipoDocumento atd on pcprovcli.IdTipoDocumento=atd.Id"
            sql = sql & " on gnc.IdProveedorRef=pcprovcli.idprovcli"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(CodTrans) & ")"
            sql = sql & " and GNC.Estado<>3 " & cond
            
            VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            '-- datos de las retenciones
            ' union de las tres tablas
            ' union de las tres tablas
            sql = " select  DISTINCT  a.Transid,a.FechaAnexos, a.FechaTrans, a.Transa,  a.NumTransa, "
            sql = sql & "a.CodTipoDocumento, a.idproveedorref, a.RUC, a.Nombre, a.BandRelacionado,"
            sql = sql & " a.NumSerieEstablecimiento , a.NumSeriePunto, "
            sql = sql & " case when len(a.NumSecuencial)>9 then right(a.NumSecuencial,9) else a.NumSecuencial end as NumSecuencial ,  "
            sql = sql & " a.NumAutSRI, a.FechaCaducidad, "
            sql = sql & "a.CodCredTrib, a.CodTipoComp, a.CodFormaPagoSRI, a.BandPagoExterior, a.CodPaisSRI , a.BandDobleTributa, a.BandPagoSujRet, "
            sql = sql & " a.Valor0, a.Valor12+ISNULL(ValorICE,0) AS Valor12 ,a.ValorNOIVA, CodICE, ISNULL(ValorICE,0) AS ValorICE, "
            sql = sql & " BANDREToTRO, (a.porcentajeiva) , a.TipoProvCli, '" & BaseDatos & "' as emp, '" & CodRetencion1 & "' as CodRetencion "
            sql = sql & " into tmpATS" & BaseDatos
            sql = sql & " From tmp1 a  "
'            sql = sql & " order by a.CodTipoComp"
            
            VerificaExistenciaTablaTemp "tmpATS" & BaseDatos
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
    End With
'    rs.Open sql, gobjMain.EmpresaActual.Coneccion, adOpenStatic, adLockReadOnly
'    UltimoRecordset rs
    Set ConsultaComprasConsolidadas2016ParaXML = rs
    Set rs = Nothing
    gobjMain.EmpresaActual.GrabaGNLogAccion "CONSULTA", "Cons Compras Consolidada para Anexos 2016 XML" & gobjCond.fecha1 & "-" & gobjCond.fecha2, "AN"
    VerificaExistenciaTabla 0
    VerificaExistenciaTabla 1
    VerificaExistenciaTabla 2
    VerificaExistenciaTabla 3
    VerificaExistenciaTabla 4
End Function


Public Function ConsANRetencionComprasConsolidadas2016ParaXML(ByVal transid As Long, BaseDatos As String, CodRetencion1 As String) As Recordset
    Dim sql As String, rs As Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    With gobjCond
        sql = " Select ts.bandiva, ts.codretencion, CASE WHEN ts.bandiva = 1 then iva.codRetencion else ir.CodRetencion end as codretencion,  "
        sql = sql & "  round(ts.porcentaje*100,2) as porcentaje,"
        sql = sql & " gncCP.codtrans, gncCP.numtrans, pc.ruc,"
        sql = sql & " gnc.codtrans, gnc.numtrans,"
        sql = sql & " gnc.FechaTrans , gnt.NumSerieEstablecimiento, gnt.NumSeriePunto, "
        If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImportaDatosRetencion")) > 0 Then
            If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImportaDatosRetencion") > 0 Then
                sql = sql & " gnc.NumDocRef as NumeroTrans , "
            Else
                sql = sql & " gnc.NumTrans as NumTrans, "
            End If
        Else
            sql = sql & " gnc.NumTrans as NumTrans, "
        End If
        
      
        sql = sql & " case when substring(Opcion2, 122, 1)='S' and tipoemision='1' and TipoAmbiente = '2' then NumeroAutorizacion else gnc.AutorizacionSRI end as NumAutorizacion,"
        sql = sql & " tsk.base, tsk.haber "
        sql = sql & " from " & BaseDatos & ".dbo.Gncomprobante gnc "
        sql = sql & " left join " & BaseDatos & ".dbo.infocomprobantes inf on gnc.transid=inf.transid "
        sql = sql & " inner join " & BaseDatos & ".dbo.gntrans gnt on gnc.codtrans=gnt.codtrans"
        sql = sql & " inner join " & BaseDatos & ".dbo.tskardexret tsk"
        sql = sql & " inner join " & BaseDatos & ".dbo.tsretencion ts"
        sql = sql & " LEFT join " & BaseDatos & ".dbo.Anexos_Retencion_IR IR on ts.IdAnexo= IR.id"
        sql = sql & " LEFT join " & BaseDatos & ".dbo.Anexos_Retencion_IVA IVA on ts.IdAnexo= IVA.id"
        sql = sql & " on tsk.idretencion=ts.idretencion"
        sql = sql & " on gnc.transid=tsk.transid"
        sql = sql & " inner join " & BaseDatos & ".dbo.pckardex pck"
        sql = sql & " inner join " & BaseDatos & ".dbo.pcprovcli pc"
        sql = sql & " on pck.idprovcli= pc.idprovcli"
        sql = sql & " on gnc.transid=pck.transid"
        sql = sql & " inner join " & BaseDatos & ".dbo.pckardex pckCP"
        sql = sql & " inner join " & BaseDatos & ".dbo.Gncomprobante gncCP"
        sql = sql & " on gncCP.transid=pckCP.transid"
        sql = sql & " on pck.idasignado=pckCP.id"
        sql = sql & " Where gnc.Estado <> 3"
        sql = sql & " and gnc.codtrans in (" & PreparaCadena(CodRetencion1) & ")"
'        sql = sql & " and TipoAmbiente = 2 "
        sql = sql & " and gnccp.transid=" & transid
        
        rs.Open sql, gobjMain.EmpresaActual.Coneccion, adOpenStatic, adLockReadOnly
        
        
        
    End With
    UltimoRecordset rs
    Set ConsANRetencionComprasConsolidadas2016ParaXML = rs
    Set rs = Nothing

End Function

Public Function ConsANRetencionComprasConsolidadas2016ParaXMLSinRetencion(ByVal transid As Long, BaseDatos As String) As Recordset
    Dim sql As String, rs As Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    With gobjCond
        sql = " Select "
        sql = sql & " 0 as bandiva ,codtiporetencion ,codtiporetencion, "
        sql = sql & "  0 as porcent, codtrans, numtrans,"
        sql = sql & " p.ruc,  "
        sql = sql & " '' as codtrans, '' as  numtrans,  '' as fechatrans, "
        sql = sql & " '' as NumSerieEstablecimiento, '' as NumSeriePunto, '' as NumTrans, "
        sql = sql & " '' as NumAutorizacion,  sum(i.costorealtotal) as base, 0 as haber "
        sql = sql & " from " & BaseDatos & ".dbo.gncomprobante g inner join " & BaseDatos & ".dbo.anexos a on g.transid=a.transid"
        sql = sql & " inner join " & BaseDatos & ".dbo.ivkardex i "
        sql = sql & " inner join " & BaseDatos & ".dbo.ivinventario ivi on i.idinventario=ivi.idinventario"
        sql = sql & " on g.transid=i.transid"
        sql = sql & " inner join " & BaseDatos & ".dbo.pcprovcli p on g.idproveedorref = p.idprovcli "
        sql = sql & " Where g.Estado <> 3  And g.TransID = " & transid
        'sql = sql & " AND BANDCOMPRASINRETENCION=1 "
        sql = sql & " group by codtiporetencion, codtrans, numtrans, p.ruc, g.fechatrans "
        rs.Open sql, gobjMain.EmpresaActual.Coneccion, adOpenStatic, adLockReadOnly
    End With
    UltimoRecordset rs
    Set ConsANRetencionComprasConsolidadas2016ParaXMLSinRetencion = rs
    Set rs = Nothing

End Function

Public Function ConsANVentas2016ConsolParaXML(cadEmpresas As String) As Variant
    Dim rs As Recordset, v As Variant, W As Variant, Y As Variant, z As Variant, j As Integer
    Dim numReg As Long
    Dim CodT As String
    Dim emp As Variant
    Dim i As Long
    Dim codret As String
    Dim sql As String
    Dim BaseDatos As String
    Y = Split(cadEmpresas, ";")
    For j = 0 To UBound(Y)
        z = Split(Y(j), ",")
        v = Split(gobjCond.CodGrupo, ";")
        CodT = ""
        emp = z(0)
        For i = 0 To UBound(v)
            W = Split(v(i), ",")
            If W(0) = emp Then
                CodT = CodT & W(1) & ","
            End If
        Next i
        
        v = Split(gobjCond.CodRetencion2, ";")
        codret = ""
        For i = 0 To UBound(v)
            W = Split(v(i), ",")
            If W(0) = emp Then
                codret = codret & W(1) & ","
            End If
        Next i
        
    CodT = Mid$(CodT, 1, Len(CodT) - 1)
    codret = Mid$(codret, 1, Len(codret) - 1)
    
    Set rs = ConsultaVentasConsolidadas2016paraXML(z(1), gobjCond.Sucursal, CodT, codret)
    Next j
    
    sql = ""
    For j = 0 To UBound(Y)
        z = Split(Y(j), ",")
        sql = sql & " select * "
        If j = 0 Then
            sql = sql & " into tmp_ConsATSVentas"
        End If
        sql = sql & " from tmpats_vta" & z(1)
        If j <> UBound(Y) Then
            sql = sql & " union all"
        End If
'        sql = sql & " select * from tmpats" & z(1)
'        sql = sql & " union all"
'        sql = sql & " select * from tmpats" & z(1)
    Next j
    
    sql = sql & " order by ruc"
    
        VerificaExistenciaTablaTemp "tmp_ConsATSVentas" & BaseDatos
        gobjMain.EmpresaActual.EjecutarSQL sql, numReg
    
    
    sql = "select "
    sql = sql & " fecha,"
    sql = sql & " codtipoDocumento,"
    sql = sql & " i,"
    sql = sql & " RUC,"
    sql = sql & " BandRelacionado,"
    sql = sql & " AnexoCodTipoComp,"
    sql = sql & " sum(cant),"
    sql = sql & " sum(Valor0),"
    sql = sql & " sum(Valor12),"
    sql = sql & " sum(ValorNO12),"
    sql = sql & " sum(totaliva),"
    sql = sql & " sum(PorcentajeIVA),"
    sql = sql & " sum(ICE),"
    sql = sql & " max(TipoEmi),"
    sql = sql & " max(Codforma0),"
    sql = sql & " max(Codforma1),"
    sql = sql & " max(Codforma2),"
    sql = sql & " max(Codforma3),"
    sql = sql & " max(Codforma4),"
    sql = sql & " max(Codforma5),"
    sql = sql & " max(Codforma6),"
    sql = sql & " max(Codforma7),"
    sql = sql & " max(Codforma8),"
    sql = sql & " max(Codforma9),"
    sql = sql & " max(BaseIVA),"
    sql = sql & " max(ValorRetIVA),"
    sql = sql & " sum(BaseRenta),"
    sql = sql & " sum(ValorRetRenta)"
    sql = sql & " from tmp_ConsATSVentas"
    sql = sql & " group by"
    sql = sql & " fecha,"
    sql = sql & " codtipoDocumento,"
    sql = sql & " i,"
    sql = sql & " RUC,"
    sql = sql & " BandRelacionado,"
    sql = sql & " AnexoCodTipoComp "
    sql = sql & " order by codtipoDocumento, ruc "
    rs.Open sql, gobjMain.EmpresaActual.Coneccion, adOpenStatic, adLockReadOnly
    UltimoRecordset rs
    
    
    Set ConsANVentas2016ConsolParaXML = rs
    Set rs = Nothing
End Function


Public Function ConsultaVentasConsolidadas2016paraXML(ByVal BaseDatos As String, ByVal Sucursal As String, ByVal CodGrupo As String, ByVal CodRetencion2 As String) As Variant
    Dim pck As PCKardex, pcd As PCDocAsignado
    Dim objGnComp As GNComprobante, objGnCompDocOriginal As GNComprobante
    Dim sql As String, cond As String, CadenaValores As String
    Dim Recargo As String 'CadenaAgrupa  As String
    Dim v As Variant, max As Integer, i As Integer, j As Integer, X As Integer
    Dim from As String, numReg As Long, f1 As String
    Dim rs As Recordset, Retenidos As Variant, NumRet As Long, v1 As Variant
    Dim Servicios As String, Bienes As String, RowRet As Long, m As Integer
    Dim rsforma As Recordset
    Const COL_BASESER = 15
    'segunda parte de la consulta: carga los valores de  Detalles de Anexos
    Const COL_FECHA = 2
    Const COL_CODCREDTRIB = 7
    Const COL_TRANSID = 1
    Const COL_TRANSRET = 5
    Dim condServ As String
    Dim condBien As String
    Dim Vista As String
    Dim pos As Integer, pos1 As Integer
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
   With gobjCond
        'Reporte de un mes a la vez
        f1 = DateSerial(Year(.fecha1), Month(.fecha1), 1)
        cond = " AND GNC.FechaTrans BETWEEN " & FechaYMD(f1, gobjMain.EmpresaActual.TipoDB) & _
               " AND " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB)
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            VerificaExistenciaTabla 2
            VerificaExistenciaTabla 3
            pos = InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "DISFRA")
            pos1 = InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "ENCOFRADO")
            
            If pos > 0 Or pos1 > 0 Then
                Vista = BaseDatos & ".dbo.vwConsSUMIVKardexIVA2B"
            Else
                Vista = BaseDatos & ".dbo.vwConsSUMIVKardexIVA"
            End If
               sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From " & BaseDatos & ".dbo.IvRecargo ivR inner join " & _
                          BaseDatos & ".dbo.IvKardexRecargo ivkR Inner join " & _
                             BaseDatos & ".dbo.GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
                sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(Sucursal) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(CodGrupo) & ")" & _
                  "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
        VerificaExistenciaTablaTemp "tan1"
        
            sql = "Select pc.ruc, SUM(IvKr.Valor) as TotalIVA Into tan1 " & _
                    "From " & BaseDatos & ".dbo.IvRecargo ivR inner join " & _
                         BaseDatos & ".dbo.IvKardexRecargo ivkR Inner join " & _
                             BaseDatos & ".dbo.GnComprobante gNc  inner join " & BaseDatos & ".dbo.pcprovcli pc on gnc.idclienteref= pc.idprovcli " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
                sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN ('IVA') " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(CodGrupo) & ")" & _
                  "Group by pc.ruc"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
        VerificaExistenciaTablaTemp "tmpICE"
            sql = "Select pc.ruc, SUM(IvKr.Valor) as TotalICE Into tmpICE " & _
                    "From " & BaseDatos & ".dbo.IvRecargo ivR inner join " & _
                        BaseDatos & ".dbo.IvKardexRecargo ivkR Inner join " & _
                            BaseDatos & ".dbo.GnComprobante gNc  inner join " & BaseDatos & ".dbo.pcprovcli pc on gnc.idclienteref= pc.idprovcli " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
                sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo = 'ICE' " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(CodGrupo) & ")" & _
                  " Group by pc.ruc"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
'--datos de indices
            VerificaExistenciaTabla 3
            sql = " Select  DISTINCT GNC.IdClienteRef AS FC_id, GNC_b.IdClienteRef AS rtp_id1, "
            sql = sql & " CASE WHEN GNC_b.IdClienteRef=0 THEN PCK_b.IdProvcli WHEN GNC_b.IdClienteRef<>0 THEN GNC_b.IdClienteRef end as rtp_id "
            sql = sql & " Into tmp3"
            sql = sql & " From    (" & BaseDatos & ".dbo.gncomprobante GNC Inner join " & BaseDatos & ".dbo.pckardex pck_a ON GNC.TransID = pck_a.TransID ) INNER JOIN"
            sql = sql & "  (" & BaseDatos & ".dbo.gncomprobante GNC_b Inner join " & BaseDatos & ".dbo.pckardex pck_B ON GNC_b.TransID = pck_B.TransID )"
            sql = sql & " ON  PCK_A.ID = PCK_B.IDASIGNADO"
            sql = sql & " where GNC.CodTrans IN (" & PreparaCadena(CodGrupo) & ")"
            sql = sql & " AND GNC_b.CodTrans IN (" & PreparaCadena(CodRetencion2) & ")"
            sql = sql & " AND (GNC_b.FechaTrans  BETWEEN " & _
                FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = sql & " and GNC.Estado<>3 and GNC.IdClienteRef<>0 and GNC_b.IdClienteRef<>0 "
            
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            VerificaExistenciaTabla 3
            sql = sql & " select"
            sql = sql & " pc.idprovcli, sum(valor) as ValorIce"
            sql = sql & " from " & BaseDatos & ".dbo.gncomprobante g inner join " & BaseDatos & ".dbo.pcprovcli pc on g.IdClienteRef = pc.IdProvCli"
            sql = sql & " inner join " & BaseDatos & ".dbo.ivkardexrecargo i"
            sql = sql & " inner join " & BaseDatos & ".dbo.IVRecargo ir"
            sql = sql & " on i.idrecargo=i.idrecargo"
            sql = sql & " on g.transid=i.transid"
            sql = sql & " where codrecargo='ice'"
            'sql = sql & " and CodTrans in ('fcm')"
            sql = sql & " AND CodTrans IN (" & PreparaCadena(CodGrupo) & ")"
            sql = sql & " AND (fechatrans BETWEEN " & _
                FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = sql & " group by pc.idprovcli"
            
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            
            '--datos de la Venta
            sql = " Select pc.ruc, ' ' as Nombre, atd.CodTipoDocumento, pc.BandRelacionado, "
            sql = sql & " AnexoCodTipoComp, count(Gnc.NumTrans) as cantTrans,  "
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotalBase0) *-1 END As Valor0,"
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotalBaseIVA) *-1 END AS Valor12,"
            sql = sql & " 0 AS ValorNO12,"
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotal) *-1 END  AS PrecioRealTotal, "
            'sql = sql & " gnc.PorcentajeIVA, "
            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
                sql = sql & "  substring(gnt.opcion2,122,1) as CompElect, "
                sql = sql & "  substring(gnt.opcion2,135,1) as Prod "
            Else
                sql = sql & "  'N' as CompElect, "
                sql = sql & "  'N' as Prod "
            End If
            sql = sql & " Into tmp1"
            sql = sql & " from    ("
            sql = sql & " tmp0"
            sql = sql & " Right join " & BaseDatos & ".dbo.gncomprobante Gnc "
            sql = sql & " inner join " & BaseDatos & ".dbo.GnTrans gnt on gnc.codtrans=GnT.codtrans"
            sql = sql & " INNER JOIN " & BaseDatos & ".dbo.PCPROVClI pc "
            sql = sql & " left join " & BaseDatos & ".dbo.anexo_TipoDocumento atd on pc.IdTipoDocumento=atd.Id"
            sql = sql & " ON GNC.IdClienteRef=pc.idprovcli "
            sql = sql & " inner join " & Vista & " vw"
            sql = sql & " ON Gnc.TransID = vw.transid"
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(CodGrupo) & ")"
            sql = sql & " and pc.bandempresapublica=0 AND Gnc.BandCierre<>9 "
            sql = sql & " and GNC.Estado<>3 " & cond
            sql = sql & " and pc.ruc<>'9999999999999'"
            sql = sql & " and  (atd.CodTipoDocumento in ( 'R','C') or atd.CodTipoDocumento is null)"
            sql = sql & " group by  pc.ruc,   AnexoCodTipoComp  ,CodTipoDocumento,  pc.BandRelacionado "
            'gnc.PorcentajeIVA ,
            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
                sql = sql & " , substring(gnt.opcion2,122,1), substring(gnt.opcion2,135,1)"
            End If
           
           sql = sql & " union all "
            
            sql = sql & " Select  pc.ruc, ' ' as Nombre, atd.CodTipoDocumento , pc.BandRelacionado, "
            sql = sql & " AnexoCodTipoComp, count(Gnc.NumTrans) as cantTrans,  "
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotalBase0) *-1 END As Valor0,"
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotalBaseIVA) *-1 END AS Valor12,"
            sql = sql & " 0 AS ValorNO12,"
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotal) *-1 END  AS PrecioRealTotal, "
            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
                sql = sql & "  substring(gnt.opcion2,122,1) as CompElect, "
                sql = sql & "  substring(gnt.opcion2,135,1) as Prod "
            Else
                sql = sql & "  'N' as CompElect, "
                sql = sql & "  'N' as Prod "
            End If
            sql = sql & " from    ("
            sql = sql & " tmp0"
            sql = sql & " Right join " & BaseDatos & ".dbo.gncomprobante Gnc "
            sql = sql & " inner join " & BaseDatos & ".dbo.GnTrans gnt on gnc.codtrans=GnT.codtrans"
            sql = sql & " INNER JOIN " & BaseDatos & ".dbo.PCPROVClI pc "
            sql = sql & " left join " & BaseDatos & ".dbo.anexo_TipoDocumento atd on pc.IdTipoDocumento=atd.Id"
            sql = sql & " ON GNC.IdClienteRef=pc.idprovcli "
            sql = sql & " inner join " & Vista & " vw"
            sql = sql & " ON Gnc.TransID = vw.transid"
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(CodGrupo) & ")"
            sql = sql & " and pc.bandempresapublica=1  AND Gnc.BandCierre<>9 "
            sql = sql & " and GNC.Estado<>3 " & cond
            sql = sql & " and pc.ruc<>'9999999999999'"
            sql = sql & " and  (atd.CodTipoDocumento in ( 'R','C') or atd.CodTipoDocumento is null)"
            sql = sql & " group by  pc.ruc ,AnexoCodTipoComp , CodTipoDocumento, pc.BandRelacionado "
            'sql = sql & " gnc.PorcentajeIVA, "
            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
                sql = sql & " , substring(gnt.opcion2,122,1), substring(gnt.opcion2,135,1)"
            End If
           
           sql = sql & " union all "
            
            sql = sql & " Select '9999999999999' AS ruc, ' ' as Nombre, 'F' as CodTipoDocumento, 2 as BandRelacionado,"
            sql = sql & " AnexoCodTipoComp, count(Gnc.NumTrans) as cantTrans,  "
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotalBase0) *-1 END As Valor0,"
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotalBaseIVA) *-1 END AS Valor12,"
            sql = sql & " 0 AS ValorNO12,"
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotal) *-1 END  AS PrecioRealTotal, "
            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
                sql = sql & "  substring(gnt.opcion2,122,1) as CompElect, "
                sql = sql & "  substring(gnt.opcion2,135,1) as Prod "
            Else
                sql = sql & "  'N' as CompElect, "
                sql = sql & "  'N' as Prod "
            End If
            sql = sql & " from    ("
            sql = sql & " tmp0"
            sql = sql & " Right join " & BaseDatos & ".dbo.gncomprobante Gnc "
            sql = sql & " inner join " & BaseDatos & ".dbo.GnTrans gnt on gnc.codtrans=GnT.codtrans"
            sql = sql & " INNER JOIN " & BaseDatos & ".dbo.PCPROVClI pc "
            sql = sql & " left join " & BaseDatos & ".dbo.anexo_TipoDocumento atd on pc.IdTipoDocumento=atd.Id"
            sql = sql & " ON GNC.IdClienteRef=pc.idprovcli "
            sql = sql & " inner join " & Vista & " vw"

            sql = sql & " ON Gnc.TransID = vw.transid"
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(CodGrupo) & ")"
            sql = sql & " and pc.bandempresapublica=0  AND Gnc.BandCierre<>9 "
            sql = sql & " and GNC.Estado<>3 " & cond
            sql = sql & " and  (atd.CodTipoDocumento not in ( 'R','C')"
            sql = sql & " OR pc.ruc='9999999999999')"
            sql = sql & " group by  AnexoCodTipoComp "
            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
                sql = sql & " , substring(gnt.opcion2,122,1), substring(gnt.opcion2,135,1) "
            End If
           
           sql = sql & " union all "
           
            sql = sql & " Select  '9999999999999' AS ruc, ' ' as Nombre, 'F' as CodTipoDocumento ,  2 as BandRelacionado,"
            sql = sql & " AnexoCodTipoComp, count(Gnc.NumTrans) as cantTrans,  "
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotalBase0) *-1 END As Valor0,"
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotalBaseIVA) *-1 END AS Valor12,"
            sql = sql & " 0 AS ValorNO12,"
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotal) *-1 END  AS PrecioRealTotal, "
            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
                sql = sql & "  substring(gnt.opcion2,122,1) as CompElect, "
                sql = sql & "  substring(gnt.opcion2,135,1) as Prod "
            Else
                sql = sql & "  'N' as CompElect, "
                sql = sql & "  'N' as Prod "
            End If
            sql = sql & " from    ("
            sql = sql & " tmp0"
            sql = sql & " Right join " & BaseDatos & ".dbo.gncomprobante Gnc "
            sql = sql & " inner join " & BaseDatos & ".dbo.GnTrans gnt on gnc.codtrans=GnT.codtrans"
            sql = sql & " INNER JOIN " & BaseDatos & ".dbo.PCPROVClI pc "
            sql = sql & " left join " & BaseDatos & ".dbo.anexo_TipoDocumento atd on pc.IdTipoDocumento=atd.Id"
            sql = sql & " ON GNC.IdClienteRef=pc.idprovcli "
            sql = sql & " inner join " & Vista & " vw"
            sql = sql & " ON Gnc.TransID = vw.transid"
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(CodGrupo) & ")"
            sql = sql & " and pc.bandempresapublica=1  AND Gnc.BandCierre<>9 "
            sql = sql & " and GNC.Estado<>3 " & cond
            sql = sql & " and (pc.ruc='9999999999999'"
            sql = sql & " OR  atd.CodTipoDocumento not in ( 'R','C'))"
            sql = sql & " group by  AnexoCodTipoComp "
            'sql = sql & " gnc.PorcentajeIVA,  "
            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
                sql = sql & " , substring(gnt.opcion2,122,1), substring(gnt.opcion2,135,1)"
            End If
           
           VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            
            sql = " select ruc  into tmp4 from tmp1 where anexocodtipocomp in (5,18) "
            
            VerificaExistenciaTabla 4
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            
            sql = " Select pck.idformaSRI, Count(pck.idformaSRI) as cant "
            sql = sql & " From " & BaseDatos & ".dbo.GnComprobante gNc"
            sql = sql & " inner join " & BaseDatos & ".dbo.pckardex pck"
            sql = sql & " on gnc.transid= pck.transid"
            sql = sql & " Where gnc.Estado <> 3"
            sql = sql & " AND (GNC.fechatrans BETWEEN " & _
                FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = sql & " AND gnc.CodTrans IN (" & PreparaCadena(CodGrupo) & ")"
            sql = sql & " Group by pck.idformaSRI"
            sql = sql & " having Count(pck.idformaSRI)>0"
            sql = sql & " order by pck.idformaSRI "
            
           Set rsforma = gobjMain.EmpresaActual.OpenRecordset(sql)
           If rsforma.RecordCount > 0 Then
           rsforma.MoveFirst
            
            For m = 0 To rsforma.RecordCount - 1
                If Not IsNull(rsforma.Fields(0)) Then
                    sql = " Select"
                    sql = sql & " case when idtipodocumento =5 then '9999999999999' else ruc end as ruc,  CodFormaPago "
                    sql = sql & " Into tmpForma" & m
                    sql = sql & " From " & BaseDatos & ".dbo.GnComprobante gNc"
                    sql = sql & " inner join " & BaseDatos & ".dbo.pckardex pck"
                    sql = sql & " left join " & BaseDatos & ".dbo.Anexo_FormaPago a  on pck.idformasri=a.id  "
                    sql = sql & " on gnc.transid= pck.transid"
                    sql = sql & " inner join " & BaseDatos & ".dbo.pcprovcli pc on gnc.idclienteref=pc.idprovcli"
                    sql = sql & " Where gnc.Estado <> 3"
                    sql = sql & " AND (GNC.fechatrans BETWEEN " & _
                        FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                        FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
                    sql = sql & " AND gnc.CodTrans IN (" & PreparaCadena(CodGrupo) & ")"
                    sql = sql & " and pck.IdFormaSRI=" & rsforma.Fields(0)
                    sql = sql & " Group by pc.ruc,  CodFormaPago, idtipodocumento "
                
                    VerificaExistenciaTablaTemp "tmpForma" & m
                    gobjMain.EmpresaActual.EjecutarSQL sql, numReg
                Else
                    MsgBox "Error falta enlaces de Forma de Cobro con el Catalogo del SRI "
                    Exit Function
                End If
                rsforma.MoveNext
            Next m
            End If
            sql = " Select"
            sql = sql & " tmp1.ruc, "
            For m = 0 To rsforma.RecordCount - 1
                sql = sql & "tmpForma" & m & ".CodformaPago as Codforma" & m & ", "
            Next m
            sql = Mid(sql, 1, Len(sql) - 2)
            
            sql = sql & " into tmpTodasFormas"
            sql = sql & " from tmp1 "
            For m = 0 To rsforma.RecordCount - 1
                sql = sql & " left join  tmpForma" & m & " on tmp1.ruc = tmpForma" & m & ".ruc "
            Next m
            
            
            VerificaExistenciaTablaTemp "tmpTodasFormas"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            
            sql = " select case when idtipodocumento =5 then '9999999999999' else ruc end as ruc,  "
            sql = sql & " AnexoCodTipoComp, case when AnexoCodTipoComp=4 then sum(valor) *-1 else  sum(valor) end as PorcentajeIVA"
            sql = sql & " into tmpIVAS"
            sql = sql & " from " & BaseDatos & ".dbo.gncomprobante g "
            sql = sql & " inner join " & BaseDatos & ".dbo.GnTrans gnt on g.codtrans=GnT.codtrans "
            sql = sql & " inner join " & BaseDatos & ".dbo.pcprovcli pc on g.idclienteref=pc.idprovcli"
            sql = sql & " inner join " & BaseDatos & ".dbo.ivkardexrecargo i"
            sql = sql & " inner join " & BaseDatos & ".dbo.ivrecargo iv"
            sql = sql & " on i.idrecargo=iv.idrecargo"
            sql = sql & " on g.transid=i.transid"
            sql = sql & " Where g.estado <>3"
            sql = sql & " AND (G.fechatrans BETWEEN " & _
                FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = sql & " AND g.CodTrans IN (" & PreparaCadena(CodGrupo) & ")"
            
            sql = sql & " and codrecargo='IVA'"
            sql = sql & " group by ruc, AnexoCodTipoComp, idtipodocumento"
            
            
            
            VerificaExistenciaTablaTemp "tmpIVAS"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            sql = " select "
            sql = sql & " RUC , AnexoCodTipoComp, "
            sql = sql & "  sum(PorcentajeIVA) as PorcentajeIVA "
            sql = sql & " into tmpIVA"
            sql = sql & " from tmpIVAS "
            sql = sql & " group by RUC , AnexoCodTipoComp, "
            sql = sql & " AnexoCodTipoComp "
            
            
            VerificaExistenciaTablaTemp "tmpIVA"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            ' union de las tres tablas
            sql = "SELECT DISTINCT " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB) & " as Fecha, "
            sql = sql & "  CodTipoDocumento,  '' as i,  (a.RUC)as ruc, BandRelacionado, a.AnexoCodTipoComp, isnull(a.cantTrans,0) as cant,"
            sql = sql & " case when a.AnexoCodTipoComp=4 then  isnull(a.Valor0,0) else isnull(a.Valor0*-1,0) end as Valor0,"
            sql = sql & " case when a.AnexoCodTipoComp=4 then  isnull(a.Valor12,0) else isnull(a.Valor12*-1,0) end as Valor12,"
            sql = sql & " case when a.AnexoCodTipoComp=4 then isnull(a.ValorNO12,0) else  isnull(a.ValorNO12*-1,0) end as ValorNO12"
            sql = sql & " ,round(b.totaliva,2) as totaliva, I.porcentajeiva,  "
            sql = sql & " isnull(TotalICE,0) as ice,"
            sql = sql & "  case when ( CompElect ='S'  and  Prod ='N' ) then 'E' else 'F' end as TipoEmi,  "
            For m = 0 To rsforma.RecordCount - 1
                sql = sql & " Codforma" & m & ",  "
            Next m
            
            For m = rsforma.RecordCount To 9
                sql = sql & "'' as Codforma" & m & ",  "
            Next m
            
            sql = Mid(sql, 1, Len(sql) - 3)
            
            sql = sql & " into tvta"
            sql = sql & " From tmp1 a  left  join tan1 b on a.ruc = b.ruc "
            sql = sql & " inner join tmpTodasFormas tf on a.ruc= tf.ruc "
            sql = sql & " left join tmpICE tice on a.ruc = tice.ruc "
            sql = sql & " left join tmpIVA i on a.ruc=i.ruc and a.AnexoCodTipoComp=i.AnexoCodTipoComp"
            sql = sql & " union all "
            sql = sql & " Select  DISTINCT " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB) & " as Fecha, "
            sql = sql & " atd.CodTipoDocumento, '' as i, (pc.RUC)as ruc, pc.BandRelacionado, '18' as AnexoCodTipoComp, 0 as cantTrans,"
            sql = sql & "0 as Valor0, 0 as Valor12 , 0 as ValorNO12 ,0 as totaliva, 0 AS  porcentajeiva, "
            sql = sql & " 0 as ice,"
            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
                sql = sql & "  'E' as CompElect, "
            Else
                sql = sql & "  'F' as CompElect, "
            End If
 '           sql = sql & " 0 as Numpagos,"
            For m = 0 To 9
                sql = sql & " '' as Codforma" & m & ",  "
            Next m
            sql = Mid(sql, 1, Len(sql) - 3)
            
            sql = sql & " From    (" & BaseDatos & ".dbo.gncomprobante GNC_b INNER JOIN " & BaseDatos & ".dbo.PCPROVClI pc"
            sql = sql & " left join " & BaseDatos & ".dbo.anexo_TipoDocumento atd on pc.IdTipoDocumento=atd.Id"
            sql = sql & " ON GNC_b.IdClienteRef=pc.idprovcli   )"
            sql = sql & " inner join " & BaseDatos & ".dbo.tskardexret ivr on gnc_b.transid = ivr.transid"
            sql = sql & " where GNC_b.CodTrans IN (" & PreparaCadena(CodRetencion2) & ")"
            sql = sql & " AND (GNC_b.FechaTrans  BETWEEN " & _
                FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = sql & " and GNC_B .Estado<>3 and GNC_b.IdClienteRef<>0 "
            sql = sql & " and pc.ruc not in (select  ruc from tmp4) "
            sql = sql & " order by a.AnexoCodTipoComp " ', a.Nombre
            
            VerificaExistenciaTablaTemp "tvta"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            sql = " select pc.ruc, '18' as AnexoCodTipoComp, "
            sql = sql & " case when bandiva=1 then sum(base) else 0 end as BaseIVA ,"
            sql = sql & " case when bandiva=1 then sum(debe) else 0 end as ValorRetIVA,"
            sql = sql & " case when bandiva<>1 then sum(base) else 0 end as BaseRenta ,"
            sql = sql & " case when bandiva<>1 then sum(debe) else 0 end as ValorRetRenta"
            sql = sql & " Into treten"
            sql = sql & " from " & BaseDatos & ".dbo.gncomprobante g"
            sql = sql & " inner join " & BaseDatos & ".dbo.pcprovcli pc on g.idclienteref= pc.idprovcli"
            sql = sql & " inner join " & BaseDatos & ".dbo.tskardexret tsk"
            sql = sql & " inner join " & BaseDatos & ".dbo.tsretencion t"
            sql = sql & " on t.IdRetencion=tsk.IdRetencion on g.transid=tsk.transid"
            sql = sql & " Where g.Estado <> 3"
            sql = sql & " and bandcliente=1"
            sql = sql & " and bandiva=1"
                        sql = sql & " AND (G.FechaTrans  BETWEEN " & _
                            FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                            FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            
            sql = sql & " group by pc.ruc, bandIVA"
            sql = sql & " Union All"
            sql = sql & " select pc.ruc, '18' as AnexoCodTipoComp, "
            sql = sql & " case when bandiva=1 then sum(base) else 0 end as BaseIVA ,"
            sql = sql & " case when bandiva=1 then sum(debe) else 0 end as ValorRetIVA,"
            sql = sql & " case when bandiva<>1 then sum(base) else 0 end as BaseRenta ,"
            sql = sql & " case when bandiva<>1 then sum(debe) else 0 end as ValorRetRenta"
            sql = sql & " from " & BaseDatos & ".dbo.gncomprobante g"
            sql = sql & " inner join " & BaseDatos & ".dbo.pcprovcli pc on g.idclienteref= pc.idprovcli"
            sql = sql & " inner join " & BaseDatos & ".dbo.tskardexret tsk"
            sql = sql & " inner join " & BaseDatos & ".dbo.tsretencion t"
            sql = sql & " on t.IdRetencion=tsk.IdRetencion"
            sql = sql & " on g.transid=tsk.transid"
            sql = sql & " Where g.Estado <> 3"
            sql = sql & " and bandcliente=1"
            sql = sql & " and bandiva=0"
                        sql = sql & " AND (G.FechaTrans  BETWEEN " & _
                            FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                            FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            
            sql = sql & " group by pc.ruc, bandIVA"
            
            VerificaExistenciaTablaTemp "treten"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            sql = "select ruc, AnexoCodTipoComp, sum(BaseIVA) as BaseIVA , sum(ValorRetIVA) as ValorRetIVA, sum(BaseRenta) as BaseRenta, sum(ValorRetRenta) as ValorRetRenta "
            sql = sql & " Into tret"
            sql = sql & " from treten group by ruc, AnexoCodTipoComp"
            
            VerificaExistenciaTablaTemp "tret"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
        sql = " select v.*, BaseIVA, ValorRetIVA, BaseRenta, ValorRetRenta  "
        sql = sql & " into tmpATS_VTA" & BaseDatos
        sql = sql & " from tvta v full join  tret r on v.ruc=r.ruc"
        sql = sql & " WHERE V.AnexoCodTipoComp=18 "
        sql = sql & " UNION ALL"
        sql = sql & " select v.*, 0 as BaseIVA, 0 as ValorRetIVA, 0 as BaseRenta, 0 as ValorRetRenta  from tvta v full join  tret r on v.ruc=r.ruc"
        sql = sql & " WHERE V.AnexoCodTipoComp = 5"
        sql = sql & " UNION ALL"
        sql = sql & " select v.*,  '' AS BaseIVA,''  AS ValorRetIVA, '' AS BaseRenta, '' AS ValorRetRenta  "
        sql = sql & " from tvta V "
        sql = sql & " WHERE V.AnexoCodTipoComp=4 "

        VerificaExistenciaTablaTemp "tmpATS_VTA" & BaseDatos
        gobjMain.EmpresaActual.EjecutarSQL sql, numReg
     
        
            
    End With
    Set ConsultaVentasConsolidadas2016paraXML = rs
    Set rs = Nothing
    gobjMain.EmpresaActual.GrabaGNLogAccion "CONSULTA", "Cons Ventas 2016 para Anexos " & gobjCond.fecha1 & "-" & gobjCond.fecha2, "AN"
    VerificaExistenciaTabla 0
    VerificaExistenciaTabla 1
    VerificaExistenciaTabla 2
    VerificaExistenciaTabla 3
    VerificaExistenciaTabla 4
End Function

Public Function ConsANVentasConsolidadasxEstablecimiento2016paraXML(cadEmpresas As String) As Variant
    Dim rs As Recordset, v As Variant, W As Variant, Y As Variant, z As Variant, j As Integer
    Dim numReg As Long
    Dim CodT As String
    Dim emp As Variant
    Dim i As Long
    Dim codret As String
    Dim sql As String
    Dim BaseDatos As String
    Y = Split(cadEmpresas, ";")
    For j = 0 To UBound(Y)
        z = Split(Y(j), ",")
        v = Split(gobjCond.CodGrupo, ";")
        CodT = ""
        emp = z(0)
        For i = 0 To UBound(v)
            W = Split(v(i), ",")
            If W(0) = emp Then
                CodT = CodT & W(1) & ","
            End If
        Next i
        
        v = Split(gobjCond.CodRetencion2, ";")
        codret = ""
        For i = 0 To UBound(v)
            W = Split(v(i), ",")
            If W(0) = emp Then
                codret = codret & W(1) & ","
            End If
        Next i
        
    If Len(CodT) > 0 Then
        CodT = Mid$(CodT, 1, Len(CodT) - 1)
    End If

    
    Set rs = ConsultaVentasxEstablecimientoConsolidadas2016paraXML(z(1), gobjCond.Sucursal, CodT, codret)
    Next j
    
    sql = ""
    For j = 0 To UBound(Y)
        z = Split(Y(j), ",")
        sql = sql & " select * "
        sql = sql & " from tmpATS_VTAxESTA" & z(1)
        If j <> UBound(Y) Then
            sql = sql & " union all"
        End If
    Next j
    
    sql = sql & " order by numestablecimiento"
    
    rs.Open sql, gobjMain.EmpresaActual.Coneccion, adOpenStatic, adLockReadOnly
    UltimoRecordset rs
    
    
    Set ConsANVentasConsolidadasxEstablecimiento2016paraXML = rs
    Set rs = Nothing
End Function


Public Function ConsultaVentasxEstablecimientoConsolidadas2016paraXML(ByVal BaseDatos As String, ByVal Sucursal As String, ByVal CodGrupo As String, ByVal CodRetencion2 As String) As Variant
    Dim pck As PCKardex, pcd As PCDocAsignado
    Dim objGnComp As GNComprobante, objGnCompDocOriginal As GNComprobante
    Dim sql As String, cond As String, CadenaValores As String
    Dim Recargo As String 'CadenaAgrupa  As String
    Dim v As Variant, max As Integer, i As Integer, j As Integer, X As Integer
    Dim from As String, numReg As Long, f1 As String
    Dim rs As Recordset, Retenidos As Variant, NumRet As Long, v1 As Variant
    Dim Servicios As String, Bienes As String, RowRet As Long
    Const COL_BASESER = 15
    'segunda parte de la consulta: carga los valores de  Detalles de Anexos
    Const COL_FECHA = 2
    Const COL_CODCREDTRIB = 7
    Const COL_TRANSID = 1
    Const COL_TRANSRET = 5
    Dim condServ As String
    Dim condBien As String
    Dim Vista As String
    Dim pos As Integer
    Dim pos1 As Integer
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
   With gobjCond
        'Reporte de un mes a la vez
        f1 = DateSerial(Year(.fecha1), Month(.fecha1), 1)
        cond = " AND GNC.FechaTrans BETWEEN " & FechaYMD(f1, gobjMain.EmpresaActual.TipoDB) & _
               " AND " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB)
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            VerificaExistenciaTabla 2
            VerificaExistenciaTabla 3
            pos = InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "DISFRA")
            pos1 = InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "ENCOFRADO")
            If pos > 0 Or pos1 > 0 Then
                Vista = "vwConsSUMIVKardexIVA2B"
            Else
                Vista = "vwConsSUMIVKardexIVA"
            End If
               sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From " & BaseDatos & ".dbo.IvRecargo ivR inner join " & _
                        "" & BaseDatos & ".dbo.IvKardexRecargo ivkR Inner join " & _
                            "" & BaseDatos & ".dbo.GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
                sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(Sucursal) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(CodGrupo) & ")" & _
                  "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
        VerificaExistenciaTablaTemp "tmpICE"
        sql = "Select "
        sql = sql & " gns.numestablecimiento, AnexoCodTipoComp,"
        sql = sql & " SUM(IvKr.Valor) as TotalICE "
        sql = sql & " Into tmpICE "
        sql = sql & "From " & BaseDatos & ".dbo.IvRecargo ivR inner join "
        sql = sql & " " & BaseDatos & ".dbo.IvKardexRecargo ivkR Inner join "
        sql = sql & " " & BaseDatos & ".dbo.GnComprobante gNc  "
        sql = sql & " inner join " & BaseDatos & ".dbo.GnTrans gnt "
        sql = sql & " inner join " & BaseDatos & ".dbo.gnsucursal gns on gnt.idsucursal=gns.idsucursal"
        sql = sql & " on gnc.codtrans=GnT.codtrans"
        sql = sql & " On ivkr.TransID = gNc.TransID "
        sql = sql & " On Ivr.IdRecargo = IvkR.IdRecargo "
        sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo = 'ICE' " & cond & " AND GNC.CodTrans IN (" & PreparaCadena(CodGrupo) & ")"
        
            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
                sql = sql & " and   substring(gnt.opcion2,122,1) ='N'  and  substring(gnt.opcion2,135,1) ='S'"
            End If
        
        sql = sql & " Group by gns.numestablecimiento, AnexoCodTipoComp"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
'--datos de indices
            VerificaExistenciaTabla 3
            sql = " Select  DISTINCT GNC.IdClienteRef AS FC_id, GNC_b.IdClienteRef AS rtp_id1, "
            sql = sql & " CASE WHEN GNC_b.IdClienteRef=0 THEN PCK_b.IdProvcli WHEN GNC_b.IdClienteRef<>0 THEN GNC_b.IdClienteRef end as rtp_id "
            sql = sql & " Into tmp3"
            sql = sql & " From    (" & BaseDatos & ".dbo.gncomprobante GNC Inner join " & BaseDatos & ".dbo.pckardex pck_a ON GNC.TransID = pck_a.TransID ) INNER JOIN"
            sql = sql & "  (" & BaseDatos & ".dbo.gncomprobante GNC_b Inner join " & BaseDatos & ".dbo.pckardex pck_B ON GNC_b.TransID = pck_B.TransID )"
            sql = sql & " ON  PCK_A.ID = PCK_B.IDASIGNADO"
            sql = sql & " where GNC.CodTrans IN (" & PreparaCadena(CodGrupo) & ")"
'            sql = sql & " AND GNC_b.CodTrans IN (" & PreparaCadena(CodRetencion2) & ")"
            sql = sql & " AND (GNC_b.FechaTrans  BETWEEN " & _
                FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = sql & " and GNC.Estado<>3 and GNC.IdClienteRef<>0 and GNC_b.IdClienteRef<>0 "
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            '--datos de la Venta
            
            
            sql = " Select gns.numestablecimiento, AnexoCodTipoComp, "
            sql = sql & " count(Gnc.NumTrans) as cantTrans,  "
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotalBase0) *-1 END As Valor0,"
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotalBaseIVA) *-1 END AS Valor12,"
            sql = sql & " 0 AS ValorNO12,"
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotal) *-1 END  AS PrecioRealTotal"
            sql = sql & " Into tmp1"
            sql = sql & " from    ("
            sql = sql & " tmp0"
            sql = sql & " Right join " & BaseDatos & ".dbo.gncomprobante Gnc "
            sql = sql & " inner join " & BaseDatos & ".dbo.GnTrans gnt "
            sql = sql & " inner join " & BaseDatos & ".dbo.gnsucursal gns on gnt.idsucursal=gns.idsucursal"
            sql = sql & " on gnc.codtrans=GnT.codtrans"
            
            sql = sql & " inner join " & BaseDatos & ".dbo." & Vista & " vw"

            sql = sql & " ON Gnc.TransID = vw.transid"
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " where  gns.bandvalida=1 and GNC.CodTrans IN (" & PreparaCadena(CodGrupo) & ")"
            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
                sql = sql & " and   substring(gnt.opcion2,122,1) ='N'  and  substring(gnt.opcion2,135,1) ='S'"
            End If
            sql = sql & " and GNC.Estado<>3  AND Gnc.BandCierre<>9 " & cond
            sql = sql & " group by  gns.numestablecimiento, AnexoCodTipoComp "
            
            VerificaExistenciaTabla 4
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            ' union de las tres tablas
            sql = "SELECT DISTINCT  gns.numestablecimiento, "
            sql = sql & "   a.AnexoCodTipoComp, isnull(a.cantTrans,0) as cantTrans ,"
            sql = sql & " case when a.AnexoCodTipoComp=4 then isnull(a.Valor0,0) else isnull(a.Valor0*-1,0) end as Valor0, "
            sql = sql & " case when a.AnexoCodTipoComp=4 then isnull(a.Valor12,0) else isnull(a.Valor12*-1,0) end as Valor12, "
            sql = sql & " case when a.AnexoCodTipoComp=4then  isnull(a.ValorNO12,0) else isnull(a.ValorNO12*-1,0) end as ValorNO12, "
            sql = sql & "  isnull(TotalICE,0) as ICE, "
            sql = sql & " case when a.AnexoCodTipoComp=4 then  isnull(a.PrecioRealTotal,0) else isnull(a.PrecioRealTotal*-1,0) end as PrecioRealTotal "
            sql = sql & " into tmpATS_VTAxESTA" & BaseDatos
            sql = sql & " From " & BaseDatos & ".dbo.gnsucursal gns left join tmp1 a on gns.numestablecimiento  = a.numestablecimiento "
            sql = sql & " left join tmpICE i on gns.numestablecimiento = i.numestablecimiento and a.AnexoCodTipoComp = i.AnexoCodTipoComp "
            sql = sql & " Where gns.BandValida = 1"
            
            VerificaExistenciaTablaTemp "tmpATS_VTAxESTA" & BaseDatos
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg

            
    End With
    Set ConsultaVentasxEstablecimientoConsolidadas2016paraXML = rs
    Set rs = Nothing
    gobjMain.EmpresaActual.GrabaGNLogAccion "CONSULTA", "Cons Ventas Estab Anexos " & gobjCond.fecha1 & "-" & gobjCond.fecha2, "AN"
    VerificaExistenciaTabla 0
    VerificaExistenciaTabla 1
    VerificaExistenciaTabla 2
    VerificaExistenciaTabla 3
    VerificaExistenciaTabla 4
End Function

Public Function ConsANAnuladosConsolidadas2016paraXML(cadEmpresas As String) As Variant
    Dim rs As Recordset, v As Variant, W As Variant, Y As Variant, z As Variant, j As Integer
    Dim numReg As Long
    Dim CodT As String
    Dim emp As Variant
    Dim i As Long
    Dim codret As String
    Dim sql As String
    Dim BaseDatos As String
    Y = Split(cadEmpresas, ";")
    For j = 0 To UBound(Y)
        z = Split(Y(j), ",")
        v = Split(gobjCond.CodPC2, ";")
        CodT = ""
        emp = z(0)
        For i = 0 To UBound(v)
            W = Split(v(i), ",")
            If W(0) = emp Then
                CodT = CodT & W(1) & ","
            End If
        Next i
        
        
    If Len(CodT) > 0 Then
        CodT = Mid$(CodT, 1, Len(CodT) - 1)
    End If

    
    Set rs = ConsultaAnuladoConsolidado2016paraXML(z(1), gobjCond.Sucursal, CodT)
    Next j
    
    sql = ""
    For j = 0 To UBound(Y)
        z = Split(Y(j), ",")
        sql = sql & " select * "
        sql = sql & " from tmpATS_ANU" & z(1)
        If j <> UBound(Y) Then
            sql = sql & " union all"
        End If
    Next j
    
'    sql = sql & " order by numestablecimiento"
    
    rs.Open sql, gobjMain.EmpresaActual.Coneccion, adOpenStatic, adLockReadOnly
    UltimoRecordset rs
    
    
    Set ConsANAnuladosConsolidadas2016paraXML = rs
    Set rs = Nothing
End Function



Public Function ConsultaAnuladoConsolidado2016paraXML(ByVal BaseDatos As String, ByVal Sucursal As String, ByVal CodPC2 As String) As Variant
    Dim sql As String, cond As String, CadenaValores As String
    Dim rs As Recordset, numReg As Long
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    With gobjCond
            cond = " AND (GNC.FechaTrans  BETWEEN " & _
                    FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                    FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = " select  FechaTrans, gnc.CodTrans, AnexoCodTipoComp,  "
            sql = sql & " gnc.NumSerieEstaSRI, gnc.NumSeriePuntoSRI, "
'            'jeaa 07/07/2006
            sql = sql & " case left(gnc.CodTrans,2)"
            sql = sql & " when 'FC' then "
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImportaDatosFactura")) > 0 Then
                If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImportaDatosFactura") > 0 Then
                    sql = sql & " NumDocRef  "
                Else
                     sql = sql & " NumTrans  "
                 End If
            Else
                sql = sql & " NumTrans  "
            End If
            sql = sql & " when 'FD' then "
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImportaDatosFactura")) > 0 Then
                If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImportaDatosFactura") > 0 Then
                    sql = sql & " NumDocRef  "
                Else
                     sql = sql & " NumTrans  "
                 End If
            Else
                sql = sql & " NumTrans  "
            End If
            sql = sql & " when 'FA' then "
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImportaDatosFactura")) > 0 Then
                If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImportaDatosFactura") > 0 Then
                    sql = sql & " NumDocRef  "
                Else
                     sql = sql & " NumTrans  "
                 End If
            Else
                sql = sql & " NumTrans  "
            End If
            sql = sql & " when 'NV' then "
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImportaDatosFactura")) > 0 Then
                If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImportaDatosFactura") > 0 Then
                    sql = sql & " NumDocRef  "
                Else
                     sql = sql & " NumTrans  "
                 End If
            Else
                sql = sql & " NumTrans  "
            End If
            sql = sql & " when 'VT' then "
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImportaDatosFactura")) > 0 Then
                If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImportaDatosFactura") > 0 Then
                    sql = sql & " NumDocRef  "
                Else
                     sql = sql & " NumTrans  "
                 End If
            Else
                sql = sql & " NumTrans  "
            End If
            sql = sql & " when 'RT' then "
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImportaDatosRetencion")) > 0 Then
                If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImportaDatosRetencion") > 0 Then
                    sql = sql & " NumDocRef "
                Else
                     sql = sql & " NumTrans  "
                End If
            Else
                sql = sql & " NumTrans  "
            End If
            sql = sql & " Else "
            sql = sql & " NumTrans "
            sql = sql & " end as numtrans, "
            'sql = sql & " numtrans, "
            'jeaa 12/09/2007
            If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImportaAutorizacion")) > 0 Then
                If gobjMain.EmpresaActual.GNOpcion.ObtenerValor("ImportaAutorizacion") > 0 Then
                    'sql = sql & " gnc.AutorizacionSRI as NumAutorizacion , "
                    sql = sql & " case when substring(Opcion2, 122, 1)='S' and tipoemision='1' and TipoAmbiente = '2'  then NumeroAutorizacion else  gnc.AutorizacionSRI end as AutorizacionSRI,"

                Else
                    sql = sql & " gnc.AutorizacionSRI as NumAutorizacion, "
                End If
            Else
                sql = sql & " gnc.AutorizacionSRI as NumAutorizacion, "
            End If
            sql = sql & " GNC.FechaTrans as FC"
            sql = sql & " into tmpATS_ANU" & BaseDatos
            sql = sql & " from " & BaseDatos & ".dbo.gncomprobante gnc inner join " & BaseDatos & ".dbo.gntrans gnt on gnc.codtrans=gnt.codtrans "
            sql = sql & " left join " & BaseDatos & ".dbo.InfoComprobantes Inf  on gnc.transid=inf.transID"
           sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(CodPC2) & ")"
            sql = sql & " and Estado = 3 " & cond
            'sql = sql & " order by AnexoCodTipoComp "
            
            VerificaExistenciaTablaTemp "tmpATS_ANU" & BaseDatos
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
    End With
    'rs.Open sql, Me.Coneccion, adOpenStatic, adLockReadOnly
    'UltimoRecordset rs
    Set ConsultaAnuladoConsolidado2016paraXML = rs
    Set rs = Nothing
    gobjMain.EmpresaActual.GrabaGNLogAccion "CONSULTA", "Cons CompAnuladas para Anexos " & gobjCond.fecha1 & "-" & gobjCond.fecha2, "AN"
End Function


Public Function ListaIVRecargoTransaccionparaFlex(ByVal CodTrans As String) As String
    Dim sql As String, rs As Recordset
    Dim s As String
    'ivr.idrecargo,
    sql = "SELECT ivr.codrecargo, ivr.descripcion " & _
            " FROM GNTransRecargo gnt inner join IVRecargo ivr on ivr.idrecargo=gnt.idRecargo " & _
            " WHERE codtrans='" & CodTrans & "' ORDER BY CodRecargo "
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    If Not rs.EOF Then s = "|" & rs.GetString(adClipString, , vbTab, "|")
    
    ListaIVRecargoTransaccionparaFlex = s
    rs.Close
    
    
    Set rs = Nothing
End Function

Public Function ConsultaGastosxCCxFila(ByVal codCC As String) As Variant
    Dim sql As String, rs As Recordset

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    sql = sql & " SELECT NOTA1, costoRealtotal "
    sql = sql & " FROM GNCOMPROBANTE G INNER JOIN IVKARDEX I"
    sql = sql & " ON G.TransID=I.TRANSID"
    sql = sql & " Where G.Estado <> 3"
    sql = sql & " AND  NOTA1 LIKE '%" & codCC & "%'"
    
    rs.Open sql, gobjMain.EmpresaActual.Coneccion, adOpenStatic, adLockReadOnly
    
    Set ConsultaGastosxCCxFila = rs
    Set rs = Nothing


End Function

Public Function ListaTransxDevolucion(CodTrans As String, CodForma As String, coddoc As String, fdesde As Date, fhasta As Date, CodGrupo As String) As Variant
    Dim sql As String, rs As Recordset, numReg As Long


    sql = " select "
    sql = sql & " g.idtransfuente "
    sql = sql & " into tmpdev "
    sql = sql & " from gncomprobante g INNER JOIN GNTRANS GT ON G.codtrans=gt.codtrans "
    sql = sql & " inner join ivkardex i"
    sql = sql & " inner join ivinventario ivi"
    sql = sql & " inner join IVGrupo2 ivg"
    sql = sql & " on ivi.idgrupo2 = ivg.idgrupo2"
    sql = sql & " on i.idinventario=ivi.idinventario"
    sql = sql & " on g.transid = i.transid"
    sql = sql & " where g.estado <> 3 "
    sql = sql & " and AnexoCodTipoComp=4"
    If Len(CodGrupo) > 0 Then
        sql = sql & " and codgrupo2='" & CodGrupo & "'"
    End If

    VerificaExistenciaTablaTemp "tmpdev"
    gobjMain.EmpresaActual.EjecutarSQL sql, numReg


    sql = "select"
    sql = sql & " g.transid , FechaTrans, NumTrans, pc.Nombre, "
    sql = sql & " ivg.codgrupo2, codinventario, ivi.descripcion,  cantidad*-1 , preciorealtotal/cantidad as pu , preciorealtotal *-1 as pt , i.orden, "
    sql = sql & " pc.email, at.Descripcion"
    sql = sql & " from gncomprobante g"
    sql = sql & " inner join ivkardex i"
    sql = sql & " inner join ivinventario ivi"
    sql = sql & " inner join IVGrupo2 ivg"
    sql = sql & " on ivi.idgrupo2 = ivg.idgrupo2"
    sql = sql & " on i.idinventario=ivi.idinventario"
    sql = sql & " on g.transid = i.transid"
    sql = sql & " inner join pcprovcli pc"
    sql = sql & " inner join Anexo_TipoDocumento at on pc.IdTipoDocumento= at.id"
    sql = sql & " on g.IdClienteRef =  pc.IdProvCli"
    sql = sql & " inner join tsformacobropago t on g.idforma=t.idforma"
    sql = sql & " Where g.Estado <> 3"
    sql = sql & " and codtrans = '" & CodTrans & "'"
    If Len(coddoc) > 0 Then
        sql = sql & " and at.CodTipoDocumento='" & coddoc & "'"
    End If
    If Len(CodForma) > 0 Then
        sql = sql & " and t.codforma='" & CodForma & "'"
    End If
    If Len(CodGrupo) > 0 Then
        sql = sql & " and codgrupo2='" & CodGrupo & "'"
    End If
    sql = sql & " AND (G.FechaTrans  BETWEEN " & _
        FechaYMD(fdesde, gobjMain.EmpresaActual.TipoDB) & " AND " & _
        FechaYMD(fhasta, gobjMain.EmpresaActual.TipoDB) & ") "
    
    sql = sql & " and g.transid not in (select idtransfuente from tmpdev) "
    
    sql = sql & " order by numtrans, i.orden"



   Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    Set ListaTransxDevolucion = rs
    Set rs = Nothing



End Function


Public Function ConsTSPorCobrarPagarCHP2016() As Variant
    Dim sql As String, cond As String, numMoneda As Integer, rs As Recordset
    Dim objCond As Condicion
    Dim BandModulo As Boolean
    Set objCond = gobjMain.objCondicion
    'Si no está especificada la moneda, coge la Moneda predeterminada
    If Len(objCond.CodMoneda) = 0 Then objCond.CodMoneda = MONEDA_PRE
    numMoneda = gobjMain.EmpresaActual.GNOpcion.IndiceMoneda(objCond.CodMoneda)
    BandModulo = gobjMain.PermisoModuloEspecial(gobjMain.UsuarioActual.codUsuario, ModuloPagaRol)
''        sql = "spConsTSPorCobrarPagarCHP " & numMoneda & ", " & _
''                objCond.CodMoneda & "," & _
''                CadenaBool(objCond.PorCobrar, TipoDB) & "," & _
''                CadenaBool(objCond.Cliente, TipoDB) & "," & _
''                CadenaBool(objCond.Proveedor, TipoDB) & "," & _
''                CadenaBool(objCond.Empleado, TipoDB) & "," & _
''                FechaYMD(objCond.fecha2, TipoDB) & ",'" & _
''                objCond.CodPC1 & "','" & _
''                objCond.CodPC1 & "','" & _
''                objCond.CodForma & "','" & _
''                objCond.CodForma & "','" & _
''                objCond.Codelemento & "','" & _
''                objCond.CodGrupo & "','" & _
''                objCond.Usuario1 & "','" & _
''                objCond.Usuario1 & "'," & _
''                IIf(Me.GNOpcion.Los100DocCobroPago, 1, 0)
''                If InStr(1, UCase(Me.GNOpcion.NombreEmpresa), "ITAL") Then
''                    sql = sql & ",2"
''                Else
''                    sql = sql & ",1"
''                End If
''                sql = sql & ",'" & objCond.CodBanco1 & "'"
''                sql = sql & ",'" & objCond.CodBanco1 & "'"
    
    
    
        sql = "SELECT"
        sql = sql & " Id,"
        sql = sql & " CodProvCli,"
        sql = sql & " Nombre,"
        sql = sql & " TransID,"
        sql = sql & " CodTrans,"
        sql = sql & " Trans,"
        sql = sql & " CodForma,"
        sql = sql & " NumLetra,"
        sql = sql & " Doc,"
        Select Case numMoneda
        Case 1
            sql = sql & " Valor1"
        Case 2
            sql = sql & " Valor2"
        Case 3
            sql = sql & " Valor3"
        Case 4
            sql = sql & " Valor4"
        End Select
        sql = sql & " AS Valor,"
        Select Case numMoneda
        Case 1
            sql = sql & " VCancelado1"
        Case 2
            sql = sql & " VCancelado2"
        Case 3
            sql = sql & "   VCancelado3"
        Case 4
            sql = sql & "   VCancelado4"
        End Select
        sql = sql & " AS Vcancelado,"

        Select Case numMoneda
        Case 1
            sql = sql & " VCancelado1CHP"
        Case 2
            sql = sql & " VCancelado2CHP"
        Case 3
            sql = sql & " VCancelado3CHP"
        Case 4
            sql = sql & " VCancelado4CHP"
        End Select
        sql = sql & " as VcanceladoCHP,"
        Select Case numMoneda
        Case 1
            sql = sql & "Valor1 - VCancelado1"
        Case 2
            sql = sql & "Valor2 - VCancelado2"
        Case 3
            sql = sql & "Valor3 - VCancelado3"
        Case 4
            sql = sql & "Valor4 - VCancelado4"
        End Select
        sql = sql & " AS Saldo,"
        sql = sql & " CodMoneda,"
        sql = sql & " Cancela,0,0,"
        Select Case numMoneda
        Case 1
            sql = sql & " 1"
        Case 2
            sql = sql & " Cotizacion2"
        Case 3
            sql = sql & " Cotizacion3"
        Case 4
            sql = sql & " Cotizacion4"
        End Select
        sql = sql & " AS Cotizacion,"
        sql = sql & " FechaEmision,"
        sql = sql & " Plazo,"
        sql = sql & " FechaVenci,0,"
        sql = sql & " CodVendedor,"
        sql = sql & " nomVendedor,"
        sql = sql & " fechaultimocobropago,"
        sql = sql & " CodSucursal,"
        sql = sql & " codelemento,"
        sql = sql & " Observacion"
        sql = sql & " from vwConsPCPorCobrarPagar"
        sql = sql & " WHERE  (FechaVenci <= " & FechaYMD(objCond.fecha2, gobjMain.EmpresaActual.TipoDB) & ") AND"
        sql = sql & " (CodMoneda= '" & objCond.CodMoneda & "') AND"
        If Len(objCond.CodPC1) > 0 Then
            sql = sql & " (CodProvCli ='" & objCond.CodPC1 & "') AND"
        End If
        If Len(objCond.CodForma) > 0 Then
            sql = sql & " (CodForma ='" & objCond.CodForma & "'" & ") AND"
        End If
        sql = sql & " (PorCobrar = " & CadenaBool(objCond.PorCobrar, gobjMain.EmpresaActual.TipoDB) & ")"
        sql = sql & " AND ( "
        Select Case numMoneda
        Case 1
            sql = sql & " Valor1 - VCancelado1"
        Case 2
            sql = sql & " Valor2 - VCancelado2"
        Case 3
            sql = sql & " Valor3 - VCancelado3"
        Case 4
            sql = sql & " Valor4 - VCancelado4"
        End Select
        sql = sql & " >0) AND"
        If Len(objCond.CodGrupo) > 0 Then
                If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "ITAL") Then
                    sql = sql & "( (CodGrupo2 = '" & objCond.CodGrupo & "')) AND"
                Else
                    sql = sql & "( (CodGrupo1 = '" & objCond.CodGrupo & "')) AND"
                End If
        
            
        End If
        sql = sql & "( (BandProveedor = " & CadenaBool(objCond.Proveedor, gobjMain.EmpresaActual.TipoDB) & ") "
        sql = sql & " OR (BandCliente = " & CadenaBool(objCond.Cliente, gobjMain.EmpresaActual.TipoDB) & " )"
        sql = sql & " OR (BandEmpleado = " & CadenaBool(objCond.Empleado, gobjMain.EmpresaActual.TipoDB) & " ) )"
        If Len(objCond.Usuario1) > 0 Then
            sql = sql & " and codvendedor ='" & objCond.Usuario1 & "'"
        End If
        If Len(objCond.Sucursal) > 0 Then
            sql = sql & " and CodSucursal ='" & objCond.Sucursal & "'"
        End If
        
        sql = sql & " ORDER BY FechaVenci DESC , CodProvCli"
        
        
        
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
        Set ConsTSPorCobrarPagarCHP2016 = rs
        Set rs = Nothing
        Set objCond = Nothing
        End Function
        

Public Function CadenaBool(v As Boolean, TipoDB As Byte) As String
    If TipoDB = TIPODB_JET Then
        CadenaBool = IIf(v, "True", "False")
    Else
        CadenaBool = IIf(v, "1", "0")
    End If
End Function


Public Function ConsTSPorCobrarPagarCheque2016() As Variant
    Dim sql As String, cond As String, numMoneda As Integer, rs As Recordset
    Dim objCond As Condicion
    
    Set objCond = gobjMain.objCondicion
    
    'Si no está especificada la moneda, coge la Moneda predeterminada
    If Len(objCond.CodMoneda) = 0 Then objCond.CodMoneda = MONEDA_PRE
    numMoneda = gobjMain.EmpresaActual.GNOpcion.IndiceMoneda(objCond.CodMoneda)
    
    If Len(objCond.Sucursal) = 0 Then
'''    sql = "spConsTSPorCobrarPagarCheque " & numMoneda & ", " & _
'''                objCond.CodMoneda & "," & _
'''                CadenaBool(objCond.PorCobrar, TipoDB) & "," & _
'''                CadenaBool(objCond.Cliente, TipoDB) & "," & _
'''                CadenaBool(objCond.Proveedor, TipoDB) & "," & _
'''                CadenaBool(objCond.Empleado, TipoDB) & "," & _
'''                FechaYMD(objCond.fecha2, TipoDB) & ",'" & _
'''                objCond.CodPC1 & "','" & _
'''                objCond.CodPC1 & "','" & _
'''                objCond.CodForma & "','" & _
'''                objCond.CodForma & "','" & _
'''                objCond.Codelemento & "','" & _
'''                objCond.CodGrupo & "','" & _
'''                objCond.Usuario1 & "','" & _
'''                objCond.Usuario1 & "'," & _
'''                IIf(Me.GNOpcion.Los100DocCobroPago, 1, 0)
'''                If InStr(1, UCase(Me.GNOpcion.NombreEmpresa), "ITAL") Then
'''                    sql = sql & ",2"
'''                Else
'''                    sql = sql & ",1"
'''                End If
'''                 sql = sql & ",'" & objCond.CodBanco1 & "'"
                
        sql = "SELECT"
        sql = sql & " Id,"
        sql = sql & " CodProvCli,"
        sql = sql & " Nombre,"
        sql = sql & " TransID,"
        sql = sql & " CodTrans,"
        sql = sql & " Trans,"
        sql = sql & " CodForma,"
        sql = sql & " NumLetra,"
        sql = sql & " Doc,"
        Select Case numMoneda
        Case 1
            sql = sql & " Valor1"
        Case 2
            sql = sql & " Valor2"
        Case 3
            sql = sql & " Valor3"
        Case 4
            sql = sql & " Valor4"
        End Select
        sql = sql & " AS Valor,"
        Select Case numMoneda
        Case 1
            sql = sql & " VCancelado1"
        Case 2
            sql = sql & " VCancelado2"
        Case 3
            sql = sql & "   VCancelado3"
        Case 4
            sql = sql & "   VCancelado4"
        End Select
        sql = sql & " AS Vcancelado,"
        sql = sql & " 0 as VcanceladoCHP,"
                
        Select Case numMoneda
        Case 1
            sql = sql & "Valor1 - VCancelado1"
        Case 2
            sql = sql & "Valor2 - VCancelado2"
        Case 3
            sql = sql & "Valor3 - VCancelado3"
        Case 4
            sql = sql & "Valor4 - VCancelado4"
        End Select
        sql = sql & " AS Saldo,"
        sql = sql & " CodMoneda,"
        sql = sql & " Cancela,0,0,"
        Select Case numMoneda
        Case 1
            sql = sql & " 1"
        Case 2
            sql = sql & " Cotizacion2"
        Case 3
            sql = sql & " Cotizacion3"
        Case 4
            sql = sql & " Cotizacion4"
        End Select
        sql = sql & " AS Cotizacion,"
        sql = sql & " FechaEmision,"
        sql = sql & " Plazo,"
        sql = sql & " FechaVenci,0,"
        sql = sql & " CodVendedor,"
        sql = sql & " nomVendedor,"
        sql = sql & " fechaultimocobropago,"
        sql = sql & " CodSucursal,"
        sql = sql & " codelemento,"
        sql = sql & " Observacion,"
        sql = sql & " codbanco ,numcheque, titularcta, 1 as CantCHP"
        sql = sql & " from vwConsPCPorCobrarPagarDep "
        sql = sql & " WHERE  (FechaVenci <= " & FechaYMD(objCond.fecha2, gobjMain.EmpresaActual.TipoDB) & ") AND"
        sql = sql & " (CodMoneda= '" & objCond.CodMoneda & "') AND"
        If Len(objCond.CodPC1) > 0 Then
            sql = sql & " (CodProvCli ='" & objCond.CodPC1 & "') AND"
        End If
        If Len(objCond.CodForma) > 0 Then
            sql = sql & " (CodForma ='" & objCond.CodForma & "'" & ") AND"
        End If
        If Len(objCond.CodBanco1) > 0 Then
            sql = sql & " (CodBanco ='" & objCond.CodBanco1 & "'" & ") AND"
        End If
        
        
        sql = sql & " (PorCobrar = " & CadenaBool(objCond.PorCobrar, gobjMain.EmpresaActual.TipoDB) & ")"
        sql = sql & " AND ( "
        Select Case numMoneda
        Case 1
            sql = sql & " Valor1 - VCancelado1"
        Case 2
            sql = sql & " Valor2 - VCancelado2"
        Case 3
            sql = sql & " Valor3 - VCancelado3"
        Case 4
            sql = sql & " Valor4 - VCancelado4"
        End Select
        sql = sql & " >0) AND"
        If Len(objCond.CodGrupo) > 0 Then
                If InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "ITAL") Then
                    sql = sql & "( (CodGrupo2 = '" & objCond.CodGrupo & "')) AND"
                Else
                    sql = sql & "( (CodGrupo1 = '" & objCond.CodGrupo & "')) AND"
                End If
        
            
        End If
        sql = sql & "( (BandProveedor = " & CadenaBool(objCond.Proveedor, gobjMain.EmpresaActual.TipoDB) & ") "
        sql = sql & " OR (BandCliente = " & CadenaBool(objCond.Cliente, gobjMain.EmpresaActual.TipoDB) & " )"
        sql = sql & " OR (BandEmpleado = " & CadenaBool(objCond.Empleado, gobjMain.EmpresaActual.TipoDB) & " ) )"
        If Len(objCond.Usuario1) > 0 Then
            sql = sql & " and codvendedor ='" & objCond.Usuario1 & "'"
        End If
        If Len(objCond.Sucursal) > 0 Then
            sql = sql & " and CodSucursal ='" & objCond.Sucursal & "'"
        End If
        
        sql = sql & " ORDER BY FechaVenci DESC , CodProvCli"
                
                
                
    Else               'AUC cambiado para agregar el filtro sucusal
'''            sql = "spConsTSPorCobrarPagarVendedorSucursal " & numMoneda & ", " & _
'''                objCond.CodMoneda & "," & _
'''                CadenaBool(objCond.PorCobrar, gobjMain.EmpresaActualTipoDB) & "," & _
'''                CadenaBool(objCond.Cliente, TipoDB) & "," & _
'''                CadenaBool(objCond.Proveedor, TipoDB) & "," & _
'''                FechaYMD(objCond.fecha2, TipoDB) & ",'" & _
'''                objCond.CodPC1 & "','" & _
'''                objCond.CodPC1 & "','" & _
'''                objCond.CodForma & "','" & _
'''                objCond.CodForma & "','" & _
'''                objCond.CodGrupo & "','" & _
'''                objCond.Usuario1 & "','" & _
'''                objCond.Usuario1 & "'," & _
'''                IIf(Me.GNOpcion.Los100DocCobroPago, 1, 0) & ",'" & _
'''                objCond.Sucursal & "','" & _
'''                objCond.Sucursal & "'"
'''
                
    End If
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    Set ConsTSPorCobrarPagarCheque2016 = rs
    Set rs = Nothing
    Set objCond = Nothing
End Function


Public Function RecuperarCodigoxNombre(key As String, tabla As String) As String
    Dim sql As String, rs As Recordset
    
    RecuperarCodigoxNombre = ""
    
    sql = "SELECT * FROM " & tabla & "  WHERE descripcion='" & key & "'"

    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        RecuperarCodigoxNombre = rs.Fields(1)
    End If
    
    Set rs = Nothing
End Function


Public Function RecuperarContratosErrados() As Variant
    Dim sql As String, rs As Recordset
    
   
    sql = "select Count (idprovcliinv), Contrato, IdInventario, pi.IdProvcli, CodProvcli, pc.nombre , pc.nombre"
    sql = sql & " from PCProvCliInv pi inner join pcprovcli pc on pi.idprovcli=pc.idprovcli"
    sql = sql & " Where 1=1 "
    'sql = sql & " and IdInventario = 96644 Or IdInventario = 96645"
    sql = sql & " group by idinventario, contrato, pi.idprovcli , codprovcli, pc.nombre"
    'sql = sql & " Having Count(idprovcliinv) > 1"
    sql = sql & " order by pc.nombre, contrato"
    
sql = " select gcc.idcentro, codcentro, contrato,  contrato,  pc.nombre  , codinventario"
sql = sql & " from gncentrocosto gcc"
sql = sql & " left join pcprovcli pc"
sql = sql & " left join PCProvCliInv piv"
sql = sql & " left join ivinventario iv"
sql = sql & " on piv.idinventario= iv.idinventario"
sql = sql & " on piv.idprovcli=pc.idprovcli"
sql = sql & " on gcc.idcliente = pc.idprovcli"
'sql = sql & " Where CodCentro <> Contrato"
sql = sql & " order by pc.nombre, codcentro"
    
    

    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    Set RecuperarContratosErrados = rs
    Set rs = Nothing

End Function

Public Function ConsCTBalancePorMesNivelesAcu() As Variant
    Dim sql As String, cond As String, rs As Recordset
    Dim v As Variant, v2() As Variant, i As Long, j As Long, numReg As Long
    Dim bandBalance As Boolean
    Dim mes As Long, MesFin As Integer
    Const BAL_VAL = 7
    bandBalance = True
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    MesFin = DateDiff("m", gobjCond.fecha1, gobjCond.fecha2) + 2
    For mes = 1 To MesFin + 1
       '/*-------------------------------------------------
        ' * Filtrar solo los asientos que est n dentro del rango
        ' --------------------------------------------------*/
        VerificaExistenciaTabla (10 * mes) + 1
        If mes = (MesFin + 1) Then
            sql = " SELECT CTCuenta.CodCuenta, 0 as debe0,"
            sql = sql & " 0 AS Haber0, 0 AS Saldo"
            sql = sql & " Into tmp" & (10 * mes) + 1
            sql = sql & " FROM CTCuenta INNER JOIN (GNComprobante"
            sql = sql & " INNER JOIN CTLibroDetalle ON GNComprobante.CodAsiento = CTLibroDetalle.CodAsiento)"
            sql = sql & " ON CTCuenta.IdCuenta = CTLibroDetalle.IdCuenta"
            sql = sql & " where ((GNComprobante.Estado)<>3  And (GNComprobante.Estado<>0))"
            sql = sql & " AND (GNComprobante.FechaTrans >= " & FechaYMD(gobjCond.fecha1, gobjMain.EmpresaActual.TipoDB) & " and GNComprobante.FechaTrans <= " & FechaYMD(DateAdd("m", mes - 2, gobjCond.fecha1), gobjMain.EmpresaActual.TipoDB) & ")"
        Else
            sql = " SELECT CTCuenta.CodCuenta, Debe/Cotizacion2 Debe0,"
            sql = sql & " Haber/Cotizacion2 AS Haber0, 0 AS Saldo"
            sql = sql & " Into tmp" & (10 * mes) + 1
            sql = sql & " FROM CTCuenta INNER JOIN (GNComprobante"
            sql = sql & " INNER JOIN CTLibroDetalle ON GNComprobante.CodAsiento = CTLibroDetalle.CodAsiento)"
            sql = sql & " ON CTCuenta.IdCuenta = CTLibroDetalle.IdCuenta"
            sql = sql & " where ((GNComprobante.Estado)<>3  And (GNComprobante.Estado<>0))"
            If mes = 1 Then
                sql = sql & " AND (GNComprobante.FechaTrans <" & FechaYMD(gobjCond.fecha1, gobjMain.EmpresaActual.TipoDB) & " )"
            ElseIf mes < (MesFin + 1) Then
                sql = sql & " AND (GNComprobante.FechaTrans >= " & FechaYMD(DateAdd("m", mes - 2, gobjCond.fecha1), gobjMain.EmpresaActual.TipoDB) & " and GNComprobante.FechaTrans < " & FechaYMD(DateAdd("m", mes - 1, gobjCond.fecha1), gobjMain.EmpresaActual.TipoDB) & ")"
            ElseIf mes = (MesFin + 1) Then
             '   sql = sql & " AND (GNComprobante.FechaTrans >= " & FechaYMD(gobjCond.fecha1, gobjMain.EmpresaActual.TipoDB) & " and GNComprobante.FechaTrans <= " & FechaYMD(DateAdd("m", mes - 2, gobjCond.fecha1), gobjMain.EmpresaActual.TipoDB) & ")"
            End If
        End If
            If Len(gobjCond.CodCuenta2) = 0 Then gobjCond.CodCuenta2 = gobjCond.CodCuenta1
            If Len(gobjCond.CodCuenta1) > 0 Then
                sql = sql & " AND CTCuenta.CodCuenta BETWEEN '" & gobjCond.CodCuenta1 & "' AND '" & gobjCond.CodCuenta2 & "'"
            End If
            If Len(gobjCond.Sucursal) > 0 Then
                sql = sql & "   and idLocal like " & gobjCond.Sucursal
            End If
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
        
        '/*-------------------------------------------------
        ' * Calcular la suma por cuenta
        ' --------------------------------------------------*/
        VerificaExistenciaTabla (10 * mes) + 2
        sql = " SELECT"
        sql = sql & " tmp" & (10 * mes) + 1 & ".CodCuenta,"
        sql = sql & " Sum(tmp" & (10 * mes) + 1 & ".Debe0) AS Debe,    Sum(tmp" & (10 * mes) + 1 & ".Haber0) AS Haber"
        sql = sql & " Into tmp" & (10 * mes) + 2
        sql = sql & " From tmp" & (10 * mes) + 1
        sql = sql & " GROUP BY tmp" & (10 * mes) + 1 & ".CodCuenta"
        gobjMain.EmpresaActual.EjecutarSQL sql, numReg
        ' /*-------------------------------------------------
        ' * Devolver en formato de Balance/Estado P.G.
        ' --------------------------------------------------*/
        VerificaExistenciaTabla (10 * mes) + 3
        sql = " SELECT"
        sql = sql & " ct.IdCuenta, ct.CodCuenta, ct.NombreCuenta,"
        sql = sql & " CTL.Nombre, isnull(Debe,0) as Debe , isnull(Haber,0) as Haber, isNull(Debe,0) - isNull(Haber,0) AS Valor, ct.Nivel,"
        sql = sql & " ct.IdCuentaSuma , ct.TipoCuenta, ct.BandTotal"
        sql = sql & " Into tmp" & (10 * mes) + 3
        sql = sql & " FROM CTLocal ctl right join CTcuenta ct"
        sql = sql & " left JOIN tmp" & (10 * mes) + 2 & " ON ct.CodCuenta = tmp" & (10 * mes) + 2 & ".CodCuenta"
        sql = sql & " on ct.idlocal=ctl.IDLocal"
        sql = sql & " ORDER BY ct.CodCuenta"
        gobjMain.EmpresaActual.EjecutarSQL sql, numReg
    Next mes
    mes = 1
    Dim mesAcu As Long
    sql = "select "
    sql = sql & " tmp" & mes & "3.IdCuenta, tmp" & mes & "3.CodCuenta, tmp" & mes & "3.NombreCuenta, tmp" & mes & "3.Nombre,  tmp" & mes & "3.BandTotal, "
    For mes = 1 To MesFin + 1
       sql = sql & " tmp" & mes & "3.debe, tmp" & mes & "3.Haber,"
        sql = sql & " tmp" & mes & "3.Valor"
        For mesAcu = 1 To mes - 1
            sql = sql & "+tmp" & mesAcu & "3.Valor"
        Next
        sql = sql & " as Valor,"
        sql = sql & " tmp" & mes & "3.Nivel,"
        sql = sql & " tmp" & mes & "3.IdCuentaSuma , tmp" & mes & "3.TipoCuenta,"
    Next mes
    'quita ultima coma
    sql = Mid$(sql, 1, Len(sql) - 1)
    sql = sql & " from tmp13  "
    If mes > 1 Then
        For mes = 2 To MesFin + 1
           sql = sql & " inner join tmp" & mes & "3 on tmp" & mes & "3.idcuenta=tmp13.idcuenta "
        Next mes
    End If
    If Not gobjCond.IncluirCero Then
        sql = sql & " where tmp13.BandTotal=" & 0
    End If
    sql = sql & " order by tmp13.codcuenta"
    rs.Open sql, gobjMain.EmpresaActual.Coneccion, adOpenStatic, adLockReadOnly
    UltimoRecordset rs
    If Not rs.EOF Then v = rs.GetRows()

    MayorizaxMeses v, MesFin
    'Pasa los valores de v a v2 cambiando formato
    ' y filtrando registros por TipoCuenta, Nivel y Valor
    Dim BandSigue As Boolean, bandDifCero As Boolean
    BandSigue = True
    bandDifCero = True
'    j = 0
    For i = 0 To UBound(v, 2)
        For mes = 0 To MesFin
           If Not IsNull(v(BAL_VAL + (8 * mes), i)) Then
                BandSigue = True
                Exit For
            Else
                BandSigue = False
            End If
        Next mes
        For mes = 0 To MesFin - 1
           If (v(5 + (6 * mes), i) <> 0 Or v(6 + (6 * mes), i) <> 0 Or v(7 + (6 * mes), i) <> 0) Then
                bandDifCero = True
                Exit For
            Else
                bandDifCero = False
            End If
        Next mes
        If BandSigue Then
            'Solo los que tengan valor diferente de 0
            If bandDifCero Then
                   'Solo hasta el nivel indicado
                   If v(8, i) <= gobjCond.Nivel Then
                        ReDim Preserve v2(70, j)
                        v2(0, j) = v(BAL_TIPO + 3, i)         'TipoCuenta
                        v2(1, j) = v(BAL_CODCUENTA, i)     'CodCuenta
                        v2(2, j) = v(BAL_NOMBRECUENTA, i)  'Cuenta
                        v2(3, j) = v(BAL_CODSUCURSAL, i)         'CODSUCURSAL   JEAA 16/01/04
                        v2(4, j) = IIf(v(4, i), 1, 0)  'Cuenta de total
                        For mes = 0 To MesFin - 1
                           v2(5 + (mes * 4), j) = v(5 + (mes * 6), i)     'Valor
                            v2(6 + (mes * 4), j) = v(6 + (mes * 6), i)     'Valor
                            v2(7 + (mes * 4), j) = v(7 + (mes * 6), i)     'Valor
                            If Not bandBalance Then v2(7 + (mes * 4), j) = v2(7 + (mes * 2), j) * -1   'Invierta el signo cuando es EstadoPG
                            v2(8 + (mes * 4), j) = v(8 + (mes * 6), i)      'Nivel
                        Next mes
                        j = j + 1
                    End If
            End If
        End If
    Next i
    If j Then ConsCTBalancePorMesNivelesAcu = v2
    Set rs = Nothing
    gobjMain.EmpresaActual.GrabaGNLogAccion "CONSULTA", "Consulta Balance x Mes Niveles", "CT"
    For mes = 1 To MesFin + 1
        VerificaExistenciaTabla (10 * mes) + 1
    Next mes
End Function


Private Sub MayorizaxMeses(ByRef v As Variant, Nivel As Integer)
    Dim i As Integer, nmax As Integer, j As Long
    Dim num As Integer
    Const INTER = 6
    nmax = ObtieneMaxconNivel2("CTCuenta", "Nivel", Nivel)

    'Desde el nivel inferior hacia el nivel superior
    For num = 0 To Nivel
       For i = nmax To 0 Step -1
            For j = 0 To UBound(v, 2)
                If (v(8, j) = i) And (v(4, j) = True) Then
                    v(5 + (INTER * num), j) = v(5 + (INTER * num), j) 'debe
                    v(6 + (INTER * num), j) = v(6 + (INTER * num), j) 'haber
                    v(7 + (INTER * num), j) = SumaHijosxMes(v, j, num) 'saldo
                End If
            Next j
        Next i
    Next num
End Sub

Private Function SumaHijosxMes(ByRef v As Variant, ByVal j As Long, ByVal mes As Integer) As Currency
    Dim i As Long, t As Currency
    Dim col_val As Integer
    col_val = 7 + (6 * mes)
    'Busca todos los hijos de la cuenta indicada por j
    ' saca la suma total de los hijos en t
    If col_val <> 6 Then
'        MsgBox "hola"
    End If
    For i = 0 To UBound(v, 2)
        'Si IdCuentaSuma es igual a IdCuenta de j, es hijo de la cuenta j
        If v(9, i) = v(BAL_IDCUENTA, j) Then
            If Not IsNull(v(col_val, i)) Then t = t + v(col_val, i)
        End If
    Next i
    SumaHijosxMes = t
End Function

Public Function ConsCCKardexPorTrans2PCK(ByVal BandProveedor As Boolean, _
                                     ByVal BandCliente As Boolean) As Variant
    Dim sql As String, rs As Recordset
    
#If DAOLIB Then
    Dim qdef As QueryDef
    Set qdef = mdbEmpresa.QueryDefs("spConsPCKardexPorTrans")
    qdef.Parameters("NumMoneda") = gobjCond.numMoneda
    qdef.Parameters("CodProvCli1") = IIf(Len(gobjCond.CodPC1) = 0, "!", gobjCond.CodPC1)
    qdef.Parameters("CodProvCli2") = IIf(Len(gobjCond.CodPC2) = 0, String$(20, "z"), gobjCond.CodPC2)
    qdef.Parameters("FechaDesde") = gobjCond.fecha1
    qdef.Parameters("FechaHasta") = gobjCond.fecha2
    qdef.Parameters("prmBandProveedor") = BandProveedor
    qdef.Parameters("prmBandCliente") = BandCliente
    Set rs = qdef.OpenRecordset(dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient

    If gobjMain.GrupoActual.PermisoActual.ConsRepVen And BandCliente Then
        sql = "spConsPCKardexPorTransVen " & _
                gobjCond.numMoneda & ",'" & _
                gobjCond.CodPC1 & "','" & _
                gobjCond.CodPC2 & "'," & _
                FechaYMD(gobjCond.fecha1, gobjMain.EmpresaActual.TipoDB) & "," & _
                FechaYMD(gobjCond.fecha2, gobjMain.EmpresaActual.TipoDB) & "," & _
                CadenaBool(BandProveedor, gobjMain.EmpresaActual.TipoDB) & "," & _
                CadenaBool(BandCliente, gobjMain.EmpresaActual.TipoDB) & ",'" & _
                gobjMain.UsuarioActual.codUsuario & "'"
    Else
        sql = "spConsCCKardexPorTrans2PCK " & _
                gobjCond.numMoneda & ",'" & _
                gobjCond.CodCentro1 & "','" & _
                gobjCond.CodCentro2 & "'," & _
                FechaYMD(gobjCond.fecha1, gobjMain.EmpresaActual.TipoDB) & "," & _
                FechaYMD(gobjCond.fecha2, gobjMain.EmpresaActual.TipoDB) & "," & _
                CadenaBool(BandProveedor, gobjMain.EmpresaActual.TipoDB) & "," & _
                CadenaBool(BandCliente, gobjMain.EmpresaActual.TipoDB)
    End If
    rs.Open sql, gobjMain.EmpresaActual.Coneccion, adOpenStatic, adLockReadOnly
    UltimoRecordset rs
#End If
    
    Set ConsCCKardexPorTrans2PCK = rs
    Set rs = Nothing
End Function


Public Function ListaGNCompetenciayZona(BandSoloValida As Boolean, BandRS As Boolean, CodZona As String) As Variant
    Dim sql As String, rs As Recordset
    

        sql = "SELECT CodCompetencia, GNCompetencia.Nombre, codzona FROM GNCompetencia left join gnzona on GNCompetencia.idzona = gnzona.idzona "

    sql = sql & "WHERE 1=1 "
    If Len(CodZona) > 0 Then
        sql = sql & "and  gnzona.CodZona='" & CodZona & "'"
    End If
    
    If BandSoloValida Then
        sql = sql & "and  GNCompetencia.BandValida=" & CadenaBool(True, gobjMain.EmpresaActual.TipoDB)
    End If
    
    sql = sql & " ORDER BY codzona, CodCompetencia"

    'Set rs = Me.OpenRecordset(sql)
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    If BandRS Then
        Set ListaGNCompetenciayZona = rs
    Else
        ListaGNCompetenciayZona = MiGetRows(rs)
        rs.Close
    End If
    Set rs = Nothing
End Function

Public Function ObtieneDescuentoxClientexItem(idcli As Long) As String
    Dim sql As String, rs As Recordset
    
    sql = "SELECT coddescuento from ivdescuento ivd inner join ivdescuentodetallepc ivdpc on ivd.iddescuento = ivdpc.iddescuento where idprovcli=" & idcli

    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    If rs.RecordCount > 0 Then
        ObtieneDescuentoxClientexItem = rs.Fields(0)
    Else
        ObtieneDescuentoxClientexItem = ""
    End If
    Set rs = Nothing

End Function

Public Function ListaIVDescuentoDet(BandRS As Boolean, bandCli As Boolean) As Variant
    Dim sql As String, rs As Recordset, NumIVGrupo As Integer, ivgrupo As Integer
    
    ivgrupo = CInt(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("DescuentoNumPCGrupo")) + 1
   
    sql = " Select"
    If bandCli Then
        sql = sql & " ivd.CODdescuento , p.codprovcli, p.nombre, "
        sql = sql & " codinventario, i.descripcion, "
    Else
        sql = sql & " i.iDINVENTARIO , "
        sql = sql & " codinventario, i.descripcion, "
        sql = sql & " p.CodProvcli , p.Nombre, "
    
    End If
    sql = sql & " precio5, descuento,  precio5-(precio5* descuento/100) "
    sql = sql & " from  IVDescuento ivd"
    sql = sql & " left join ivdescuentodetallepc idpc"
    sql = sql & " left join pcprovcli p on idpc.idprovcli=p.idprovcli"
    sql = sql & " left join pcgrupo" & ivgrupo & " pg on idpc.idPCgrupo=pg.idgrupo" & ivgrupo
    sql = sql & " on ivd.iddescuento  =idpc.iddescuento"
    
    sql = sql & " left join ivdescuentodetalleiv idiv"
    sql = sql & " left join ivinventario i on idiv.idinventario=i.idinventario"
    sql = sql & " on ivd.iddescuento  =idiv.iddescuento"
    
    If bandCli Then
        sql = sql & " order by p.nombre, codinventario "
    Else
        sql = sql & " order by  codinventario , p.nombre"
    End If
   
   Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
   
    If BandRS Then
        Set ListaIVDescuentoDet = rs
    Else
        ListaIVDescuentoDet = MiGetRows(rs)
        rs.Close
    End If
    Set rs = Nothing
End Function

Public Function ListaIVDescuentoDetxItem(CodInventario As String, _
                                  CodAlterno As String, _
                                  Descripcion As String, _
                                  CodBodega As String, _
                                  CodGrupo() As String, _
                                  Tipo As String, _
                                  BandSoloValida As Boolean, BandRS As Boolean) As Variant
    Dim sql As String, cond As String, comodin As String
    Dim rs As Recordset, i As Integer
#If DAOLIB Then
    comodin = "*"
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    comodin = "%"
#End If

    sql = " Select"
    sql = sql & " i.iDINVENTARIO , "
    sql = sql & " codinventario, i.descripcion, "
    sql = sql & " p.CodProvcli , p.Nombre, "
    sql = sql & " precio5, descuento,  precio5-(precio5* descuento/100) "
    sql = sql & " from  IVDescuento ivd"
    sql = sql & " left join ivdescuentodetallepc idpc"
    sql = sql & " left join pcprovcli p on idpc.idprovcli=p.idprovcli"
    sql = sql & " left join pcgrupo1 pg on idpc.idPCgrupo=pg.idgrupo1"
    sql = sql & " on ivd.iddescuento  =idpc.iddescuento"
    sql = sql & " left join ivdescuentodetalleiv idiv"
    sql = sql & " left join ivinventario i on idiv.idinventario=i.idinventario"
    sql = sql & " on ivd.iddescuento  =idiv.iddescuento"

    If BandSoloValida Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(i.BandValida=" & CadenaBool(True, gobjMain.EmpresaActual.TipoDB) & ")"
    End If
    If Len(CodInventario) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodInventario LIKE '" & CodInventario & comodin & "') "
    End If
    If Len(CodAlterno) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "((CodAlterno1 LIKE '" & CodAlterno & comodin & "') " & _
                      "OR (CodAlterno2 LIKE '" & CodAlterno & comodin & "')) "
    End If
    If Len(Descripcion) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(i.Descripcion LIKE '" & Descripcion & comodin & "') "
    End If
    If Len(CodBodega) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(CodBodega='" & CodBodega & "') "
    End If
    
    If Len(Tipo) > 0 Then
        If Len(cond) > 0 Then cond = cond & "AND "
        cond = cond & "(Tipo=" & Tipo & ") "
    End If
    
    
    For i = LBound(CodGrupo) To UBound(CodGrupo)
        If Len(CodGrupo(i)) > 0 Then
            If Len(cond) > 0 Then cond = cond & "AND "
'            cond = cond & "(CodGrupo" & i & " LIKE '" & CodGrupo(i) & comodin & "') "
            cond = cond & "(CodGrupo" & i & " = '" & CodGrupo(i) & "') "
        End If
    Next i
    If Len(cond) > 0 Then sql = sql & " WHERE " & cond
    'sql = sql & " ORDER BY CodInventario,CodBodega "
    sql = sql & " order by  codinventario , p.nombre"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If BandRS Then
        Set ListaIVDescuentoDetxItem = rs
    Else
        ListaIVDescuentoDetxItem = MiGetRows(rs)
        rs.Close
   End If
    Set rs = Nothing
End Function


Public Function ListaIVDescuentoDetxCli(BandProveedor As Boolean, _
                                    BandCliente As Boolean, _
                                   codPC As String, _
                                    Nombre As String, _
                                    nombreAlt As String, _
                                    Direccion As String, _
                                    RUC As String, _
                                    CodGrupoDesde As String, _
                                    CodGrupoHasta As String, _
                                    numG As Integer, _
                                    CodProvincia As String, _
                                    CodCanton As String, _
                                    CodVendedor As String, _
                                    BandSoloValida As Boolean, _
                                    BandRS As Boolean) As Variant
    Dim sql As String, cond As String, comodin As String
    Dim rs As Recordset, i As Integer
#If DAOLIB Then
    comodin = "*"
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    comodin = "%"
#End If

    sql = " Select"
    sql = sql & " ivd.CODdescuento , pc.codprovcli, pc.nombre, "
    sql = sql & " codinventario, i.descripcion, "
    sql = sql & " precio5, descuento,  precio5-(precio5* descuento/100) "
    sql = sql & " from  IVDescuento ivd"
    sql = sql & " left join ivdescuentodetallepc idpc"
    sql = sql & " left join pcprovcli pc on idpc.idprovcli=pc.idprovcli"
    sql = sql & " left join pcgrupo1 pg on idpc.idPCgrupo=pg.idgrupo1"
    sql = sql & " on ivd.iddescuento  =idpc.iddescuento"
    
    sql = sql & " left join ivdescuentodetalleiv idiv"
    sql = sql & " left join ivinventario i on idiv.idinventario=i.idinventario"
    sql = sql & " on ivd.iddescuento  =idiv.iddescuento"


    If Not (BandProveedor And BandCliente) Then
        If BandProveedor Then
            sql = sql & " WHERE BandProveedor=" & CadenaBool(True, gobjMain.EmpresaActual.TipoDB)
        Else
            sql = sql & " WHERE BandCliente=" & CadenaBool(True, gobjMain.EmpresaActual.TipoDB)
        End If
    End If
   If Len(codPC) > 0 Then
        cond = cond & " and (CodProvCli LIKE '" & codPC & comodin & "') "
    End If
    If Len(Nombre) > 0 Then
        cond = cond & " and ((pc.Nombre LIKE '" & Nombre & comodin & "') " & _
                      "OR (pc.Nombre LIKE '" & Nombre & comodin & "')) "
   End If
    
    If Len(nombreAlt) > 0 Then
        cond = cond & " and ((pc.NombreAlterno LIKE '" & nombreAlt & comodin & "') " & _
                      "OR (pc.NombreAlterno LIKE '" & nombreAlt & comodin & "')) "
   End If
    
    If Len(Direccion) > 0 Then
        cond = cond & " and (Direccion1 LIKE '" & Direccion & comodin & "') "
    End If
    If Len(RUC) > 0 Then
        cond = cond & " and (ruc LIKE '" & RUC & comodin & "') "
    End If
    
    If Len(CodProvincia) > 0 Then
        cond = cond & " and (pcprov.codprovincia='" & CodProvincia & "') "
    End If

    If Len(CodCanton) > 0 Then
        cond = cond & " and (pccan.codcanton='" & CodCanton & "') "
    End If

    If Len(CodVendedor) > 0 Then
        cond = cond & " and (fc.codvendedor='" & CodVendedor & "') "
    End If
    
    
    If Not (Len(CodGrupoDesde) = 0 And Len(CodGrupoHasta) = 0) Then
        cond = cond & " AND (CodGrupo" & numG & " BETWEEN '" & CodGrupoDesde & _
            "' AND '" & CodGrupoHasta & "')"
    End If
    If Len(cond) > 0 Then sql = sql & cond
    sql = sql & " ORDER BY pc.Nombre ,  codinventario"
   
   Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
   
    If BandRS Then
        Set ListaIVDescuentoDetxCli = rs
   Else
        ListaIVDescuentoDetxCli = MiGetRows(rs)
       rs.Close
    End If
    Set rs = Nothing
End Function


Public Function ListaIVDescuentoDetallexItem(IdInventario As String) As Variant
    Dim sql As String, cond As String, comodin As String
    Dim rs As Recordset, i As Integer

    sql = " Select"
    sql = sql & "  "
   sql = sql & "  idiv.IdDescuentoDetalleIV, "
    sql = sql & " p.CodProvcli , p.Nombre, "
    sql = sql & " precio5, descuento,  precio5-(precio5* descuento/100) "
    sql = sql & " from  IVDescuento ivd"
    sql = sql & " left join ivdescuentodetallepc idpc"
    sql = sql & " left join pcprovcli p on idpc.idprovcli=p.idprovcli"
    sql = sql & " left join pcgrupo1 pg on idpc.idPCgrupo=pg.idgrupo1"
    sql = sql & " on ivd.iddescuento  =idpc.iddescuento"
    sql = sql & " left join ivdescuentodetalleiv idiv"
    sql = sql & " left join ivinventario i on idiv.idinventario=i.idinventario"
    sql = sql & " on ivd.iddescuento  =idiv.iddescuento"
    sql = sql & " WHERE i.iDINVENTARIO =" & IdInventario
    sql = sql & " order by   p.nombre"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    Set ListaIVDescuentoDetallexItem = rs
    Set rs = Nothing
End Function


Public Function CargaConsolidadoBufferMega() As Variant
    'jeaa 07-07-04
    ' PARA SIETE EMPRESAS COMO MAXIMO
    Dim sql As String, cond As String, CadenaPrecio As String, Trans As String
    Dim cadempresa As String, cadBodega As String, totalcero As Boolean, cond1 As String
    Dim v As Variant, max As Integer, TotalC As String, TotalP As String, i As Integer, rs As Recordset, TotalCMAx As String
    Dim CodInv As String, codBod As String, Orden As String, cadTrans As String
    Dim W As Variant, j As Integer, Grupo As String
    Dim vBod As Variant, vNomEmp As Variant, vTrans As Variant
    Dim Cadtitulo As String, k As Integer
    Dim condTrans As String, vaux As Variant
    Dim numReg As Long, numMoneda As String
    Dim cadETrans As Variant, cadEBodega As Variant, basePrincipal As String
    Dim TotalE As String, TotalProm As String, difDias As Integer, Intervalo As Double
    Dim TotalCUC As String, TotalUtil As String, TotalPUtil As String, Fijo As Integer
    Dim NumMeses As Long, nmes As Currency
    With gobjCond
        NumMeses = DateDiff("m", .fecha1, .fecha2)
        difDias = DateDiff("d", .fecha1, .fecha2)
        NumMeses = Round(difDias / 30, 2)
        nmes = Round(difDias / 30, 2)
        If NumMeses = 0 Then NumMeses = 1
       numMoneda = Trim(IIf(.numMoneda > 0, Str(.numMoneda + 1), ""))
        v = Split(.Sucursal, ";")
        vTrans = Split(.CodTrans, ";")
        For i = 0 To UBound(v)
            vNomEmp = Split(v(i), ",")
            Cadtitulo = Cadtitulo & vNomEmp(1) & " - "
        Next i
        'para quitar el ultimo "-"
        Cadtitulo = Mid$(Cadtitulo, 1, Len(Cadtitulo) - 3)
        cond = " "
        TotalC = ", ( "
        TotalCMAx = ", ( "
        TotalP = ", ( "
        TotalProm = ", ( "
        TotalE = ", ( "
        TotalCUC = ", ( "
        TotalUtil = ", ( "
        TotalPUtil = ", ( "
        max = UBound(v, 1)
        ReDim cadETrans(max, 2)
        ReDim cadEBodega(max, 2)
'        cadETrans = llenaCadxEmpresa(vTrans, UBound(v))
        W = Split(v(0), ",")
        basePrincipal = W(1)
        difDias = DateDiff("d", .fecha1, .fecha2)
'        max = 0
'        For k = 0 To max
            sql = "SELECT   "
            sql = sql & "  idinventario, CodInventario, Descripcion, DESCRIPCION2, CODUNIDAD, " & NumMeses & "  as TiempoPromVta, TiempoReposicion, "
            sql = sql & " CodGrupo" & .NumGrupo & " as CodGrupo,"
            sql = sql & "  DescGrupo" & .NumGrupo & "  as DescGrupo,"
            sql = sql & " isnull(sum(Cant),0) * -1  AS TCant,   "

            sql = sql & "  0 AS CantMax "
            sql = sql & "  Into TPed" & k
            W = Split(v(k), ",")
            sql = sql & " from "
            sql = sql & " vwVentasPorBodega "
            cond = " WHERE (FechaTrans BETWEEN " & _
                         FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & _
                        " AND " & FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            cond = cond & " AND  AnexoCodTipoComp in (" & PreparaCadena(.CodBanco2) & ") "
            cond = cond & " AND  AnexoCodTipoTrans in (" & PreparaCadena(.CodBanco1) & ") "
'            If Not (Len(.Grupo1) = 0 And Len(.Grupo2) = 0) Then
'                cond = cond & " AND (codGrupo" & .NumGrupo & " BETWEEN '" & _
'                .Grupo1 & "' AND '" & .Grupo2 & "') "
'            End If
'            If Not (Len(.CodItem1) = 0 And Len(.CodItem2) = 0) Then
'                cond = cond & " AND (CodInventario BETWEEN '" & .CodItem1 & "' AND '" & .CodItem2 & "') "
'            End If
            
            If Not (Len(.Bienes)) = 0 Then
                 cond = cond & Replace(.Bienes, "IVInventario.", "vwVentasPorBodega.", 1)       'Aquí se ha grabado SQL de ítems
             End If
            
            
            If Len(.CodBodega1) > 0 Then
                cond = cond & " AND CodBodega in (" & PreparaCadena(.CodBodega1) & ") "
            End If
            
            Grupo = " group by "
            Grupo = Grupo & "  idinventario, CodInventario, Descripcion,   DESCRIPCION2, CODUNIDAD, "
            Grupo = Grupo & " CodGrupo" & .NumGrupo & " , DescGrupo" & .NumGrupo
            Grupo = Grupo & " , TiempoPromVta , TiempoReposicion"
            Orden = " ORDER BY CodInventario"
            sql = sql & cond & cond1 & Grupo & Orden
            
            VerificaExistenciaTablaTemp "TPed" & k
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
    
            sql = "  SELECT     idinventario,"
            sql = sql & "  isnull(max(Abs(cant)), 0) As CantMax "
            sql = sql & "  Into Tped" & k + 50
            sql = sql & "  from "
            sql = sql & W(1) & ".dbo.vwVentasPorBodega"
            sql = sql & " WHERE (FechaTrans BETWEEN " & _
                         FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & _
                        " AND " & FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            
            sql = sql & " AND  AnexoCodTipoComp in ('18') "
            sql = sql & " AND  AnexoCodTipoTrans in (" & PreparaCadena(.CodBanco1) & ") "
            
            If Not (Len(.Grupo1) = 0 And Len(.Grupo2) = 0) Then
                sql = sql & " AND (codGrupo" & .NumGrupo & " BETWEEN '" & _
                .Grupo1 & "' AND '" & .Grupo2 & "') "
            End If
            If Not (Len(.CodItem1) = 0 And Len(.CodItem2) = 0) Then
                sql = sql & " AND (CodInventario BETWEEN '" & .CodItem1 & "' AND '" & .CodItem2 & "') "
            End If
        
            If Len(.CodBodega1) > 0 Then
                sql = sql & " AND CodBodega in (" & PreparaCadena(.CodBodega1) & ") "
            End If
        
            If Not (Len(.Bienes)) = 0 Then
                 sql = sql & Replace(.Bienes, "IVInventario.", "vwVentasPorBodega.", 1)
             End If
        
        
                sql = sql & " group by   idinventario"
            
            VerificaExistenciaTablaTemp "TPed" & k + 50
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
                            
            
            sql = " select idinventario, sum(exist) as exist, "
            sql = sql & "  Codinventario, "
            sql = sql & "  Descripcion, "
            sql = sql & "  Descripcion2, "
            sql = sql & "  CodUnidad, "
            sql = sql & "  idgrupo1, idgrupo2,idgrupo3,idgrupo4,idgrupo5, "
            sql = sql & "  codgrupo1, codgrupo2,codgrupo3,codgrupo4,codgrupo5, "
            sql = sql & "  DescGrupo1, DescGrupo2,DescGrupo3,DescGrupo4,DescGrupo5 "
            
            
            sql = sql & "  into TPedB"
            sql = sql & " from VWivexistxBodega"
            sql = sql & " where 1=1 "
            sql = sql & " and CodBodega in (" & PreparaCadena(.CodBodega1) & ")"
            sql = sql & " group by idinventario,"
            sql = sql & "  Codinventario, "
            sql = sql & "  Descripcion, "
            sql = sql & "  Descripcion2, "
            sql = sql & "  CodUnidad, "
            sql = sql & "  idgrupo1, idgrupo2,idgrupo3,idgrupo4,idgrupo5, "
            sql = sql & "  codgrupo1, codgrupo2,codgrupo3,codgrupo4,codgrupo5, "
            sql = sql & "  DescGrupo1, DescGrupo2,DescGrupo3,DescGrupo4,DescGrupo5 "
            
            
            
            VerificaExistenciaTablaTemp "TPedB"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            
            sql = "select "
            sql = sql & "  isnull(TPed" & k & ".Codinventario, i.Codinventario) As Codinventario, "
            sql = sql & "  isnull(TPed" & k & ".Descripcion, i.Descripcion) as Descripcion, "
            sql = sql & "  isnull(TPed" & k & ".Descripcion2, i.Descripcion2) as Descripcion2, "
            sql = sql & "  isnull(TPed" & k & ".CodUnidad, i.CodUnidad) as CodUnidad, "
            sql = sql & "TPed" & k & ".TiempoPromVta as TiempoPromVta, "
            sql = sql & "TPed" & k & ".TiempoReposicion as TiempoReposicion, "
            sql = sql & "  isnull(TPed" & k & ".CodGrupo,i.CodGrupo" & .NumGrupo & ") as CodGrupo, "
            sql = sql & "  isnull(TPed" & k & ".DescGrupo, i.DescGrupo" & .NumGrupo & ") as DescGrupo, "
            sql = sql & "  isnull(sum(TPed" & k & ".TCant), 0) as TCant, "
            sql = sql & "  isnull(sum(TPed" & k + 50 & ".CantMax), 0)  as TCAntMAx, "
''            sql = sql & "  isnull(sum(TPed" & k & ".TCant), 0)/" & nmes & "+ TPed" & k & ".TiempoReposicion/30 as Prom, "
''            sql = sql & "  round(isnull(sum(Tped" & k & ".TCant), 0)/( " & difDias & " + TPed" & k & ".TiempoReposicion ),2) as PromDiario, "
            sql = sql & "  isnull(sum(TPed" & k & ".TCant), 0)/" & nmes & " as Prom, "
            sql = sql & "  round(isnull(sum(Tped" & k & ".TCant), 0)/( " & difDias & "  ),2) as PromDiario, "
            
            sql = sql & "  isnull(sum(i.exist), 0) as Existencia "
            sql = sql & "  Into TmpPed0"
            sql = sql & " from "
            sql = sql & " Tped" & k & " left join Tped" & k + 50 & " on Tped" & k & ".idinventario = Tped" & k + 50 & ".idinventario"
            sql = sql & "  full join TPedB  i "
            sql = sql & " on Tped" & k & ".idinventario=i.idinventario"
            sql = sql & " WHERE  1=1 "

'            If Len(.CodBodega1) > 0 Then
'                sql = sql & " i.CodBodega in (" & PreparaCadena(.CodBodega1) & ") "
'            End If


            If Not (Len(.Grupo1) = 0 And Len(.Grupo2) = 0) Then
                sql = sql & "  AND (codGrupo" & .NumGrupo & " BETWEEN '" & _
                .Grupo1 & "' AND '" & .Grupo2 & "') "
            End If
            If Not (Len(.CodItem1) = 0 And Len(.CodItem2) = 0) Then
                sql = sql & "AND (i.CodInventario BETWEEN '" & .CodItem1 & "' AND '" & .CodItem2 & "') "
            End If
            
            If Not (Len(.Bienes)) = 0 Then
                 sql = sql & Replace(.Bienes, "IVInventario.", "i.", 1)
             End If
            
            
            
            sql = sql & "  group by  "
            'i.exist, "
            sql = sql & "  isnull(Tped" & k & ".Codinventario, i.Codinventario) , "
            sql = sql & "  isnull(Tped" & k & ".Descripcion, i.Descripcion) , "
            sql = sql & "  isnull(Tped" & k & ".Descripcion2, i.Descripcion2) , "
            sql = sql & "  isnull(Tped" & k & ".CodUnidad, i.CodUnidad) , "
            sql = sql & "Tped" & k & ".TiempoPromVta, "
            sql = sql & "Tped" & k & ".TiempoReposicion, "
            sql = sql & "  isnull(Tped" & k & ".CodGrupo,i.CodGrupo" & .NumGrupo & ") , "
            sql = sql & "  isnull(Tped" & k & ".DescGrupo, i.DescGrupo" & .NumGrupo & ")  "
                
            VerificaExistenciaTablaTemp "TmpPed0"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
 '       Next k
        
        sql = " select "
        sql = sql & "CodInventario, "
        sql = sql & "descripcion, "
        sql = sql & "descripcion2,"
        sql = sql & "CodUnidad, "

        sql = Mid$(sql, 1, Len(sql) - 2)
        i = 0
        TotalC = TotalC & " isnull(TMP.TCant,0) +"
        TotalCMAx = TotalCMAx & " isnull(TMP.TCantMax,0) +"
         TotalProm = TotalProm & " isnull(TMP.prom,0) +"
         TotalP = TotalP & " isnull(TMP.PromDiario,0) +"
         TotalE = TotalE & " isnull(TMP.Existencia,0) +"
        TotalC = Mid(TotalC, 1, Len(TotalC) - 1)
        TotalCMAx = Mid(TotalCMAx, 1, Len(TotalCMAx) - 1)
        TotalP = Mid(TotalP, 1, Len(TotalP) - 1)
        TotalProm = Mid(TotalProm, 1, Len(TotalProm) - 1)
        TotalE = Mid(TotalE, 1, Len(TotalE) - 1)
        i = 0
        sql = sql & "  ,isnull(TMP.TiempoReposicion,0) AS TiempoRep"
        sql = sql & ", Round(isnull(TMP.prom,0),2)  as TotalPromMen "
        sql = sql & ", Round(isnull(TMP.PromDiario,0),2)  as TotalPromDia "
        sql = sql & ", Round(isnull(TMP.Existencia,0),2)  as TotalExist "
        sql = sql & ", Round(" & Mid$(TotalC, 2, Len(TotalC)) & "),2)  as TotalC "
        sql = sql & ", Round(" & Mid$(TotalCMAx, 2, Len(TotalCMAx)) & "),2)  as Totalcmax "
        sql = sql & ", Round(isnull(TMP.PromDiario,0)  *  isnull(TMP.TiempoReposicion,0),2) AS ExistenciaOptima "

        sql = sql & "  Into TmpPed1"

        sql = sql & " From "
        sql = sql & " TmpPed0 TMP "

        If max > 0 Then
            For i = 0 To max - 1
                If i = 0 Then
                    CodInv = " ON TMP.CodInventario="
                    CodInv = CodInv & "TMP.CodInventario"
                    sql = sql & CodInv
                Else
                    Dim m As Integer
                    CodInv = ""
                    For m = 0 To i
                        If m = 0 Then
                            CodInv = CodInv & " ON TMP.CodInventario="
                            CodInv = CodInv & " TMP.CodInventario"
                        Else
                            CodInv = CodInv & " OR TMP.CodInventario="
                            CodInv = CodInv & "TMP.CodInventario"
                        End If
                    Next m
                    sql = sql & CodInv
                End If
            Next i
        End If
        
        VerificaExistenciaTablaTemp "TmpPed1"
        gobjMain.EmpresaActual.EjecutarSQL sql, numReg
                
        sql = " select distinct "
        sql = sql & " i.idinventario, t.CodInventario, "
        sql = sql & " t.descripcion, t.descripcion2, t.codunidad,  '' as CODBODEGA, TotalExist , Totalcmax, "
        'sql = sql & "  abs( ExistenciaOptima - TotalExist) as saldo, "
        sql = sql & "  case when i.tiemporeposicion<>0 then "
'''        sql = sql & "  abs(  ExistenciaOptima *(" & gobjMain.objCondicion.NumDias1 & "+ i.tiemporeposicion )- TotalExist) else 0 end as saldo, "
        sql = sql & " abs( ISNULL(TotalPromDia,0) *(i.tiemporeposicion+" & .NumDias1 & ")- TotalExist) else 0 end as saldo,"

        
''        sql = sql & "  abs(  ISNULL(TotalPromDia,0)  * " & gobjMain.objCondicion.NumDias1 & ") - TotalExist   " & " as saldo, "
        sql = sql & "  0 as BandPedido,  ISNULL(TotalC,0)as TotalC, ISNULL(TotalPromMen,0)as TotalPromMen , "
        sql = sql & "  ISNULL(TotalPromDia,0) as TotalPromDia, "
        sql = sql & " case when ISNULL(TotalPromDia,0)<>0 then TotalExist /ISNULL(TotalPromDia,0) else TotalExist end , "
        sql = sql & " pc.nombre, "
        sql = sql & "  ISNULL(IP.Cantidad,0), ISNULL(CostoUnitario,0), ip.Fecha, ExistenciaOptima, i.tiemporeposicion, "
        sql = sql & "  case when i.tiemporeposicion<>0 then "
        sql = sql & gobjMain.objCondicion.NumDias1 & "/i.tiemporeposicion else 0 end as x,"
        sql = sql & "  case when i.tiemporeposicion<>0 then "
        sql = sql & " ExistenciaOptima *" & gobjMain.objCondicion.NumDias1 & "/i.tiemporeposicion else 0 end as y"
        sql = sql & " from TmpPed1 t "
        sql = sql & " inner join ivinventario i on t.codinventario=i.codinventario "
        Select Case .Nivel
        Case 0
            sql = sql & " left join IVProveedorDetalle ip"
            sql = sql & " inner join pcprovcli pc on ip.IdProveedor=pc.idprovcli and ip.bandprincipal=1"
            sql = sql & " on i.idinventario=ip.idinventario"
        Case 1
            sql = sql & " left join VWIVProveedorDetalle ip"
            sql = sql & " inner join pcprovcli pc on ip.IdProveedor=pc.idprovcli"
            sql = sql & " on i.idinventario=ip.idinventario"
        End Select
        sql = sql & " WHERE i.bandvalida=1 "
        
        If .tipoItem1 Then
            sql = sql & "  and  (TotalExist  >= ( ExistenciaOptima+ (ExistenciaOptima *0.5))) and (TotalExist <= TotalPromMen)"
            If .IncluirCero Then
                sql = sql & " or  (TotalExist  <  ExistenciaOptima) "
            End If
        Else
            If .IncluirCero Then
'                sql = sql & " and (TotalExist  <   "
'                sql = sql & " ExistenciaOptima *" & gobjMain.objCondicion.NumDias1 & "/i.tiemporeposicion )"
                
                sql = sql & " and TotalPromDia>0 and case when ISNULL(TotalPromDia,0)<>0 then TotalExist /ISNULL(TotalPromDia,0) else TotalExist end <=(i.tiemporeposicion+" & gobjMain.objCondicion.NumDias1 & ")"
                
'            Else
'                sql = sql & " WHERE (TotalExist  >= ( ExistenciaOptima+ (ExistenciaOptima *0.5))) and (TotalExist <= TotalPromMen)"
'                sql = sql & " or  (TotalExist  <  ExistenciaOptima) "
            End If
        End If
        
        i = 0
        
        
       sql = sql & " order by "
        sql = sql & "t.descripcion "
 
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    Set CargaConsolidadoBufferMega = rs
    Set rs = Nothing
       
    End With
End Function


Public Function ObtieneIVCodigoRelacionado(cod As String) As Variant
    Dim sql As String, rs As Recordset
    sql = "select idproveedor, idinventario from ivproveedordetalle where codinventarioprov='" & cod & "'"
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    Set ObtieneIVCodigoRelacionado = rs
    Set rs = Nothing
    
    
End Function

Public Function RecuperarItemxNombre(key As Variant) As String
   Dim sql As String, i As Integer
    Dim rs As Recordset
    Dim cond As String
        cond = " Descripcion ='" & key & "'"
        
        If VarType(key) = vbString Then
            sql = "Select idInventario, CodInventario, descripcion "
            sql = sql & " FROM IvInventario WHERE" & cond
        Else
            sql = "Select idInventario, CodInventario, descripcion"
            sql = sql & " FROM IvInventario WHERE" & cond
        End If
        
        
   Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If rs.EOF Then 'AUC cambiado antes estuvo NOT (RS.EOF)
       Err.Raise ERR_NOHAYCODIGO, "IvFamiliaDetalle", MSGERR_NOHAYCODIGO
    Else
                RecuperarItemxNombre = rs!CodInventario
  End If
End Function

Public Function ListaEmpleadoFlexGrid2() As String
    Dim s As String, sql As String
    Dim rs As Recordset
    sql = "SELECT Nombre,CodProvcli FROM Empleado "
    sql = sql & " Inner join personal on personal.idempleado = empleado.idprovcli"
    'sql = sql & " WHERE BandActivo=" & CadenaBool(True, TipoDB)
    sql = sql & " AND BandGestion=" & CadenaBool(False, gobjMain.EmpresaActual.TipoDB)
    sql = sql & " ORDER BY CodProvcli"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
#If DAOLIB Then
    s = "|" & GetStringDAO(rs, 10 + 40)         '*** MAKOTO 25/ene/01 Mod.
#Else
    If Not rs.EOF Then s = "|" & rs.GetString(adClipString, , vbTab, "|")
#End If
    ListaEmpleadoFlexGrid2 = s
    rs.Close
    Set rs = Nothing
End Function

Public Function ConsANCompras2016ParaTalon() As Variant
    Dim pck As PCKardex, pcd As PCDocAsignado
    Dim objGnComp As GNComprobante, objGnCompDocOriginal As GNComprobante
    Dim sql As String, cond As String, CadenaValores As String
    Dim Recargo As String 'CadenaAgrupa  As String
    Dim v As Variant, max As Integer, i As Integer, j As Integer, X As Integer
    Dim from As String, numReg As Long
    Dim rs As Recordset, Retenidos As Variant, NumRet As Long, v1 As Variant
    Dim Servicios As String, Bienes As String, RowRet As Long
    Dim codRecarBaseCero As String, codRecarBaseNoCero As String, rsRecargo As Recordset
    Const COL_BASESER = 15
    'segunda parte de la consulta: carga los valores de  Detalles de Anexos
    Const COL_FECHA = 2
    Const COL_CODCREDTRIB = 7
    Const COL_TRANSID = 1
    Const COL_TRANSRET = 5
    Dim condServ As String
    Dim condBien As String
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    With gobjCond
            cond = " AND (GNC.FechaTrans  BETWEEN " & _
                    FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                    FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            VerificaExistenciaTabla 2
            VerificaExistenciaTabla 3
            VerificaExistenciaTabla 4
            VerificaExistenciaTabla 5
            VerificaExistenciaTabla 6
            VerificaExistenciaTabla 10
            VerificaExistenciaTabla 11
            VerificaExistenciaTabla 12
            VerificaExistenciaTabla 13
            VerificaExistenciaTabla 14
            VerificaExistenciaTabla 15
            VerificaExistenciaTabla 16
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & " WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(.Servicios) & ") "
            sql = sql & " AND GNC.CodTrans IN (" & PreparaCadena(.CodTrans) & ")"
            sql = sql & " AND ivr.CodRecargo <> 'ICE'"
            sql = sql & "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp10 " & _
                    "From IvRecargo ivR inner join " & _
                        "AFKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
            sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(.Servicios) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(.CodTrans) & ")" & _
                  "Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            'jeaa 09/01/2007
            'datos ICE
            sql = " Select Ivkr.TransID, "
            sql = sql & " (ivkr.Valor) AS BASEice,"
            sql = sql & " 'ICE' AS CodICE, (ivkr.Valor) As ValorICE"
            sql = sql & " Into tmp6"
            sql = sql & " From IvRecargo ivR"
            sql = sql & " inner join IvKardexRecargo ivkR"
            sql = sql & " Inner join GnComprobante gNc"
            sql = sql & " On ivkr.TransID = gNc.TransID"
            sql = sql & " On Ivr.IdRecargo = IvkR.IdRecargo"
            sql = sql & " Where gnc.Estado <> 3"
            'sql = sql & " AND ivr.CodRecargo IN ('" & gobjMain.EmpresaActual.GNOpcion.ObtenerValor("RecDescParaICE") & "')" & cond
            sql = sql & " AND ivr.CodRecargo = 'ICE'" & cond
            sql = sql & " AND GNC.CodTrans IN (" & PreparaCadena(.CodTrans) & ")"
            sql = sql & " and ivkr.Valor<>0"
            'sql = sql & " Group by IvkR.TransID"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg

'--datos de indices
            VerificaExistenciaTabla 3
            sql = " Select  DISTINCT GNC.TransID AS cp_id, GNC_b.TransID AS rtp_id"
            sql = sql & " Into tmp3"
            sql = sql & " From    (gncomprobante GNC Inner join pckardex pck_a ON GNC.TransID = pck_a.TransID ) INNER JOIN"
            sql = sql & "  (gncomprobante GNC_b Inner join pckardex pck_B ON GNC_b.TransID = pck_B.TransID )"
            sql = sql & " ON  PCK_A.ID = PCK_B.IDASIGNADO"
            sql = sql & " AND GNC.CodTrans IN (" & PreparaCadena(gobjCond.CodTrans) & ")"
            sql = sql & " AND GNC_b.CodTrans IN (" & PreparaCadena(gobjCond.CodRetencion1) & ")"
            sql = sql & " AND (GNC.FechaTrans  BETWEEN " & _
                FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = sql & " and GNC.Estado<>3 and GNC_b.estado<>3 "
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            sql = " Select  DISTINCT GNC.TransID AS cp_id, GNC_b.TransID AS rtp_id"
            sql = sql & " Into tmp13"
            sql = sql & " From    (gncomprobante GNC Inner join pckardex pck_a ON GNC.TransID = pck_a.TransID ) INNER JOIN"
            sql = sql & "  (gncomprobante GNC_b Inner join pckardex pck_B ON GNC_b.TransID = pck_B.TransID )"
            sql = sql & " ON  PCK_A.ID = PCK_B.IDASIGNADO"
            sql = sql & " AND GNC.CodTrans IN (" & PreparaCadena(gobjCond.CodTrans) & ")"
            sql = sql & " AND GNC_b.CodTrans IN (" & PreparaCadena(gobjCond.CodRetencion1) & ")"
            sql = sql & " AND (GNC.FechaTrans  BETWEEN " & _
                FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = sql & " and GNC.Estado<>3 and GNC_b.estado<>3 "
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
         sql = "select codrecargo, BandBaseCero from ivrecargo where BandAfectaBaseAnexo=1  "
         Set rsRecargo = gobjMain.EmpresaActual.OpenRecordset(sql)
         
         If Not rsRecargo.EOF Then
            rsRecargo.MoveFirst
        End If
         
        codRecarBaseCero = ""
        codRecarBaseNoCero = ""
         
         For i = 1 To rsRecargo.RecordCount
            If rsRecargo.RecordCount > 0 Then
               If rsRecargo.Fields(1) Then
                   codRecarBaseCero = codRecarBaseCero & "'" & rsRecargo.Fields(0) & "',"
               Else
                   codRecarBaseNoCero = codRecarBaseNoCero & "'" & rsRecargo.Fields(0) & "',"
               End If
            End If
            rsRecargo.MoveNext
         Next i
        
        If Len(codRecarBaseCero) > 0 Then
            codRecarBaseCero = Mid$(codRecarBaseCero, 1, Len(codRecarBaseCero) - 1)
        End If
            
        If Len(codRecarBaseNoCero) > 0 Then
            codRecarBaseNoCero = Mid$(codRecarBaseNoCero, 1, Len(codRecarBaseNoCero) - 1)
        End If
            
        'If Len(codRecarBaseCero) > 0 Then
        sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalRecargoBaseCero "
        sql = sql & " Into tmpRecarBaseCero "
        sql = sql & " From IvRecargo ivR inner join "
        sql = sql & " IvKardexRecargo ivkR Inner join "
        sql = sql & " GnComprobante gNc  "
        sql = sql & " On ivkr.TransID = gNc.TransID "
        sql = sql & " On Ivr.IdRecargo = IvkR.IdRecargo "
        sql = sql & " WHERE gNc.Estado <> 3 "
        If Len(codRecarBaseCero) > 0 Then
            sql = sql & " AND ivr.CodRecargo IN (" & codRecarBaseCero & ") " & cond
        Else
            sql = sql & " AND ivr.CodRecargo IN ('xxxx') " & cond
        End If
        sql = sql & " AND GNC.CodTrans IN (" & PreparaCadena(.CodTrans) & ")"
        sql = sql & " Group by IvkR.TransID"
        
        VerificaExistenciaTablaTemp "tmpRecarBaseCero"
        gobjMain.EmpresaActual.EjecutarSQL sql, numReg
        
            
            
            
            '--datos de la compra
            sql = " Select  'CP' AS TRANS,"
            'sql = sql & " Ane.CodTipoComp,"
            sql = sql & " RIGHT('0' + Ane.CodTipoComp,2) AS CodTipoComp, "
            sql = sql & " COUNT(Ane.CodTipoComp) AS CantTrans, "
            sql = sql & " SUM(Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) )  end) As Valor0, "
            sql = sql & " SUM(Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " ( (vw.SignoCompra * (vw.CostoTotalBaseIVA + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))))) +  isnull(tmp6.BASEice,0)) end ) AS Valor12, "
            sql = sql & " SUM(Case vw.CostoTotalBaseNOIVA When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBaseNOIVA + (vw.CostoTotalBaseNOIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end )AS ValorNOIVA, "
            
            sql = sql & " SUM((Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " ( (vw.SignoCompra * (vw.CostoTotalBaseIVA + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))))) +  isnull(tmp6.BASEice,0)) end) *maxiva ) AS ValorIVA "
            sql = sql & " into tmpTaloCP "
            sql = sql & " from    (( tmp0 Right join gncomprobante Gnc "
            sql = sql & " left join tmp6 on gnc.transid=tmp6.transid"
            
              If gobjMain.EmpresaActual.GNOpcion.IVKTipoDatoDouble Then
                sql = sql & " inner join vwConsSUMIVKardexIVADou vw ON Gnc.TransID = vw.transid "
            Else
                sql = sql & " inner join vwConsSUMIVKardexIVA vw ON Gnc.TransID = vw.transid "
            End If
            
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(gobjCond.CodTrans) & ")"
            sql = sql & " and GNC.Estado<>3 " & cond
            sql = sql & " GROUP BY CODTIPOCOMP"
            
            sql = sql & " union all"
            '--datos de la compra activos
            sql = sql & " Select  'CP' AS TRANS, "
            sql = sql & " RIGHT('0' + Ane.CodTipoComp,2) AS CodTipoComp, "
            sql = sql & " COUNT(Ane.CodTipoComp) AS CantTrans, "
            sql = sql & " SUM(Case vw.CostoTotalBase0 When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBase0 + (vw.CostoTotalBase0 * (cast( isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float))) ) end )As Valor0, "
            sql = sql & " SUM(Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " vw.SignoCompra * (vw.CostoTotalBaseIVA + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))) end ) AS Valor12, "
            sql = sql & " 0 as ValorNOIVA, "
            
            sql = sql & " SUM((Case vw.CostoTotalBaseIVA When 0 then 0 else "
            sql = sql & " ( (vw.SignoCompra * (vw.CostoTotalBaseIVA + (vw.CostoTotalBaseIVA * (cast(isnull(TotalDescuento,0) as float) / cast(vw.CostoTotal as float)))))   ) end) *maxiva ) AS ValorIVA "
            
            sql = sql & " from    (( tmp10 Right join gncomprobante Gnc "
            sql = sql & " inner join vwConsSUMAFKardexIVA vw ON Gnc.TransID = vw.transid "
            sql = sql & " ON tmp10.TransID = Gnc.TransID)"
            sql = sql & " inner join Anexos Ane on Gnc.TransID = Ane.Transid)"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(gobjCond.CodTrans) & ")"
            sql = sql & " and GNC.Estado<>3 " & cond
            sql = sql & " GROUP BY CODTIPOCOMP"
            sql = sql & " order by CodTipoComp"
            
        VerificaExistenciaTablaTemp "tmpTaloCP"
        gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            sql = " Select  "
            sql = sql & " TRANS, "
            sql = sql & " CodTipoComp, "
            sql = sql & " sum(CantTrans) as CantTrans, "
            sql = sql & " sum(Valor0) as Valor0 , "
            sql = sql & " sum(Valor12) as Valor12 , "
            sql = sql & " sum(ValorNOIVA) as Valornoiva, "
            sql = sql & " sum(ValorIVA) as ValorIVA "
            
            sql = sql & " from    tmpTaloCP  "
            
            sql = sql & " group by TRANS, "
            sql = sql & " CodTipoComp "
            
            sql = sql & " order by CodTipoComp"
            
            
            
            
    End With
    rs.Open sql, gobjMain.EmpresaActual.Coneccion, adOpenStatic, adLockReadOnly
    UltimoRecordset rs
    Set ConsANCompras2016ParaTalon = rs
    Set rs = Nothing
    gobjMain.EmpresaActual.GrabaGNLogAccion "CONSULTA", "Cons Compraspara Anexos 2016 XML" & gobjCond.fecha1 & "-" & gobjCond.fecha2, "AN"
    VerificaExistenciaTabla 0
    VerificaExistenciaTabla 1
    VerificaExistenciaTabla 2
    VerificaExistenciaTabla 3
    VerificaExistenciaTabla 4
End Function


Public Function ConsANVentas2016paraTALON() As Variant
    Dim pck As PCKardex, pcd As PCDocAsignado
    Dim objGnComp As GNComprobante, objGnCompDocOriginal As GNComprobante
    Dim sql As String, cond As String, CadenaValores As String
    Dim Recargo As String 'CadenaAgrupa  As String
    Dim v As Variant, max As Integer, i As Integer, j As Integer, X As Integer
    Dim from As String, numReg As Long, f1 As String
    Dim rs As Recordset, Retenidos As Variant, NumRet As Long, v1 As Variant, codRecarBaseCero As String, codRecarBaseNoCero As String
    Dim Servicios As String, Bienes As String, RowRet As Long, m As Integer, rsRecargo As Recordset
    Dim rsforma As Recordset
    Const COL_BASESER = 15
    'segunda parte de la consulta: carga los valores de  Detalles de Anexos
    Const COL_FECHA = 2
    Const COL_CODCREDTRIB = 7
    Const COL_TRANSID = 1
    Const COL_TRANSRET = 5
    Dim condServ As String
    Dim condBien As String
    Dim Vista As String
    Dim pos As Integer, pos1 As Integer
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
   With gobjCond
        'Reporte de un mes a la vez
        f1 = DateSerial(Year(.fecha1), Month(.fecha1), 1)
        cond = " AND GNC.FechaTrans BETWEEN " & FechaYMD(f1, gobjMain.EmpresaActual.TipoDB) & _
               " AND " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB)
            VerificaExistenciaTabla 0
            VerificaExistenciaTabla 1
            VerificaExistenciaTabla 2
            VerificaExistenciaTabla 3
            pos = InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "DISFRA")
            pos1 = InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "ENCOFRADO")
            
            If gobjMain.EmpresaActual.GNOpcion.IVKTipoDatoDouble Then
                Vista = "vwConsSUMIVKardexIVADou"
            Else
               If pos > 0 Or pos1 > 0 Then
                    Vista = "vwConsSUMIVKardexIVA2B"
                Else
                    Vista = "vwConsSUMIVKardexIVA"
                End If
            End If

               sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalDescuento Into tmp0 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
                sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN (" & PreparaCadena(.Sucursal) & ") " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(.CodGrupo) & ")" & _
                  "Group by IvkR.TransID"
            
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
        VerificaExistenciaTablaTemp "tan1"
            sql = "Select pc.ruc, SUM(IvKr.Valor) as TotalIVA Into tan1 " & _
                    "From IvRecargo ivR inner join " & _
                        "IvKardexRecargo ivkR Inner join " & _
                            "GnComprobante gNc  inner join pcprovcli pc on gnc.idclienteref= pc.idprovcli " & _
                        "On ivkr.TransID = gNc.TransID " & _
                    "On Ivr.IdRecargo = IvkR.IdRecargo "
                sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo IN ('IVA') " & cond & _
                    " AND GNC.CodTrans IN (" & PreparaCadena(.CodGrupo) & ")" & _
                  "Group by pc.ruc"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
        VerificaExistenciaTablaTemp "tmpICE"
        sql = "Select pc.ruc, SUM(IvKr.Valor) as TotalICE  "
'''        If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
'''            sql = sql & "  substring(gnt.opcion2,122,1) as CompElect, "
'''            sql = sql & "  substring(gnt.opcion2,135,1) as Prod "
'''        Else
'''            sql = sql & "  'N' as CompElect, "
'''            sql = sql & "  'N' as Prod "
'''        End If
        sql = sql & "  Into tmpICE "
        
        sql = sql & " From IvRecargo ivR inner join "
        sql = sql & " IvKardexRecargo ivkR Inner join "
        sql = sql & " GnComprobante gNc  "
        sql = sql & " inner join GnTrans gnt on gnc.codtrans=GnT.codtrans"
        sql = sql & " inner join pcprovcli pc on gnc.idclienteref= pc.idprovcli "
        sql = sql & " On ivkr.TransID = gNc.TransID "
        sql = sql & " On Ivr.IdRecargo = IvkR.IdRecargo "
        sql = sql & "WHERE gNc.Estado <> 3 AND ivr.CodRecargo = 'ICE' " & cond
        sql = sql & " AND GNC.CodTrans IN (" & PreparaCadena(.CodGrupo) & ")"
        sql = sql & " Group by pc.ruc"
'''        If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
'''            sql = sql & " , substring(gnt.opcion2,122,1), substring(gnt.opcion2,135,1)"
'''        End If
                  
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
         sql = "select codrecargo, BandBaseCero from ivrecargo where BandAfectaBaseAnexo=1  "
         Set rsRecargo = gobjMain.EmpresaActual.OpenRecordset(sql)
         
         If Not rsRecargo.EOF Then
            rsRecargo.MoveFirst
         End If
        codRecarBaseCero = ""
        codRecarBaseNoCero = ""
         
         For i = 1 To rsRecargo.RecordCount
            If rsRecargo.RecordCount > 0 Then
               If rsRecargo.Fields(1) Then
                   codRecarBaseCero = codRecarBaseCero & "'" & rsRecargo.Fields(0) & "',"
               Else
                   codRecarBaseNoCero = codRecarBaseNoCero & "'" & rsRecargo.Fields(0) & "',"
               End If
            End If
            rsRecargo.MoveNext
         Next i
        
        If Len(codRecarBaseCero) > 0 Then
            codRecarBaseCero = Mid$(codRecarBaseCero, 1, Len(codRecarBaseCero) - 1)
        End If
            
        If Len(codRecarBaseNoCero) > 0 Then
            codRecarBaseNoCero = Mid$(codRecarBaseNoCero, 1, Len(codRecarBaseNoCero) - 1)
        End If
            
            
        
        
        sql = "Select Ivkr.TransID, SUM(IvKr.Valor) as TotalRecargoBaseCero "
        sql = sql & " Into tmpRecarBaseCero "
        sql = sql & " From IvRecargo ivR inner join "
        sql = sql & " IvKardexRecargo ivkR Inner join "
        sql = sql & " GnComprobante gNc  "
        sql = sql & " On ivkr.TransID = gNc.TransID "
        sql = sql & " On Ivr.IdRecargo = IvkR.IdRecargo "
        sql = sql & " WHERE gNc.Estado <> 3 "
        If Len(codRecarBaseCero) > 0 Then
            sql = sql & " AND ivr.CodRecargo IN (" & codRecarBaseCero & ") " & cond
        End If
        sql = sql & " AND GNC.CodTrans IN (" & PreparaCadena(.CodGrupo) & ")"
        sql = sql & " Group by IvkR.TransID"
        
        VerificaExistenciaTablaTemp "tmpRecarBaseCero"
        gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            
            
            
            
'--datos de indices
            VerificaExistenciaTabla 3
            sql = " Select  DISTINCT GNC.IdClienteRef AS FC_id, GNC_b.IdClienteRef AS rtp_id1, "
            sql = sql & " CASE WHEN GNC_b.IdClienteRef=0 THEN PCK_b.IdProvcli WHEN GNC_b.IdClienteRef<>0 THEN GNC_b.IdClienteRef end as rtp_id "
            sql = sql & " Into tmp3"
            sql = sql & " From    (gncomprobante GNC Inner join pckardex pck_a ON GNC.TransID = pck_a.TransID ) INNER JOIN"
            sql = sql & "  (gncomprobante GNC_b Inner join pckardex pck_B ON GNC_b.TransID = pck_B.TransID )"
            sql = sql & " ON  PCK_A.ID = PCK_B.IDASIGNADO"
            sql = sql & " where GNC.CodTrans IN (" & PreparaCadena(gobjCond.CodGrupo) & ")"
            sql = sql & " AND GNC_b.CodTrans IN (" & PreparaCadena(gobjCond.CodRetencion2) & ")"
            sql = sql & " AND (GNC_b.FechaTrans  BETWEEN " & _
                FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = sql & " and GNC.Estado<>3 and GNC.IdClienteRef<>0 and GNC_b.IdClienteRef<>0 "
            
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            VerificaExistenciaTabla 3
            sql = sql & " select"
            sql = sql & " pc.idprovcli, sum(valor) as ValorIce"
            sql = sql & " from gncomprobante g inner join pcprovcli pc on g.IdClienteRef = pc.IdProvCli"
            sql = sql & " inner join ivkardexrecargo i"
            sql = sql & " inner join IVRecargo ir"
            sql = sql & " on i.idrecargo=i.idrecargo"
            sql = sql & " on g.transid=i.transid"
            sql = sql & " where codrecargo='ice'"
            'sql = sql & " and CodTrans in ('fcm')"
            sql = sql & " AND CodTrans IN (" & PreparaCadena(.CodGrupo) & ")"
            sql = sql & " AND (fechatrans BETWEEN " & _
                FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = sql & " group by pc.idprovcli"
            
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            
            '--datos de la Venta
            sql = " Select pc.ruc, ' ' as Nombre, atd.CodTipoDocumento, pc.BandRelacionado, "
            sql = sql & " AnexoCodTipoComp, count(Gnc.NumTrans) as cantTrans,  "
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN (sum( vw.PrecioRealTotalBase0) *-1 + (isnull(sum(tr.TotalRecargoBaseCero),0)*-1) )END As Valor0,"
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotalBaseIVA) *-1 END AS Valor12,"
            sql = sql & " 0 AS ValorNO12,"
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotal) *-1 END  AS PrecioRealTotal "
'            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
'                sql = sql & "  substring(gnt.opcion2,122,1) as CompElect, "
'                sql = sql & "  substring(gnt.opcion2,135,1) as Prod "
'            Else
'                sql = sql & "  'N' as CompElect, "
'                sql = sql & "  'N' as Prod "
'            End If
            sql = sql & " Into tmp1"
            sql = sql & " from    ("
            sql = sql & " tmp0"
            sql = sql & " Right join gncomprobante Gnc "
            sql = sql & " inner join GnTrans gnt on gnc.codtrans=GnT.codtrans"
            sql = sql & " INNER JOIN PCPROVClI pc "
            sql = sql & " left join anexo_TipoDocumento atd on pc.IdTipoDocumento=atd.Id"
            sql = sql & " ON GNC.IdClienteRef=pc.idprovcli "
            sql = sql & " inner join " & Vista & " vw"
            sql = sql & " ON Gnc.TransID = vw.transid"
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " left join tmpRecarBaseCero tr ON tmp0.TransID = tr.TransID"
            
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(gobjCond.CodGrupo) & ")"
            sql = sql & " and Gnc.BandCierre<>9 "
            sql = sql & " and GNC.Estado<>3 " & cond
            sql = sql & " and pc.ruc<>'9999999999999'"
            sql = sql & " and  (atd.CodTipoDocumento in ( 'R','C') or atd.CodTipoDocumento is null)"
            sql = sql & " group by  pc.ruc,   AnexoCodTipoComp  ,CodTipoDocumento,  pc.BandRelacionado "
            'gnc.PorcentajeIVA ,
'            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
'                sql = sql & " , substring(gnt.opcion2,122,1), substring(gnt.opcion2,135,1)"
'            End If
           
           sql = sql & " union all "
            
            sql = sql & " Select  pc.ruc, ' ' as Nombre, atd.CodTipoDocumento , pc.BandRelacionado, "
            sql = sql & " AnexoCodTipoComp, count(Gnc.NumTrans) as cantTrans,  "
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN (sum( vw.PrecioRealTotalBase0) *-1  + (isnull(sum(tr.TotalRecargoBaseCero),0)*-1) ) END As Valor0,"
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotalBaseIVA) *-1 END AS Valor12,"
            sql = sql & " 0 AS ValorNO12,"
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotal) *-1 END  AS PrecioRealTotal "
            'sql = sql & " gnc.PorcentajeIVA,            "
'            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
'                sql = sql & "  substring(gnt.opcion2,122,1) as CompElect, "
'                sql = sql & "  substring(gnt.opcion2,135,1) as Prod "
'            Else
'                sql = sql & "  'N' as CompElect, "
'                sql = sql & "  'N' as Prod "
'            End If
            sql = sql & " from    ("
            sql = sql & " tmp0"
            sql = sql & " Right join gncomprobante Gnc "
            sql = sql & " inner join GnTrans gnt on gnc.codtrans=GnT.codtrans"
            sql = sql & " INNER JOIN PCPROVClI pc "
            sql = sql & " left join anexo_TipoDocumento atd on pc.IdTipoDocumento=atd.Id"
            sql = sql & " ON GNC.IdClienteRef=pc.idprovcli "
            sql = sql & " inner join " & Vista & " vw"
            sql = sql & " ON Gnc.TransID = vw.transid"
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " left join tmpRecarBaseCero tr ON tmp0.TransID = tr.TransID"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(gobjCond.CodGrupo) & ")"
            sql = sql & " AND Gnc.BandCierre<>9 "
            sql = sql & " and GNC.Estado<>3 " & cond
            sql = sql & " and pc.ruc<>'9999999999999'"
            sql = sql & " and  (atd.CodTipoDocumento not in ( 'R','C') or atd.CodTipoDocumento is null)"
            sql = sql & " group by  pc.ruc ,AnexoCodTipoComp , CodTipoDocumento, pc.BandRelacionado "
'            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
'                sql = sql & " , substring(gnt.opcion2,122,1), substring(gnt.opcion2,135,1)"
'            End If
           
           sql = sql & " union all "
            
            sql = sql & " Select '9999999999999' AS ruc, ' ' as Nombre, 'F' as CodTipoDocumento, 2 as BandRelacionado,"
            sql = sql & " AnexoCodTipoComp, count(Gnc.NumTrans) as cantTrans,  "
            
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN (sum( vw.PrecioRealTotalBase0) *-1 + (isnull(sum(tr.TotalRecargoBaseCero),0)*-1) ) END As Valor0,"
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotalBaseIVA) *-1 END AS Valor12,"
            sql = sql & " 0 AS ValorNO12,"
            sql = sql & "  CASE WHEN SUM(vw.PrecioRealTotal)<>0 THEN sum( vw.PrecioRealTotal) *-1 END  AS PrecioRealTotal "
'            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
'                sql = sql & "  substring(gnt.opcion2,122,1) as CompElect, "
'                sql = sql & "  substring(gnt.opcion2,135,1) as Prod "
'            Else
'                sql = sql & "  'N' as CompElect, "
'                sql = sql & "  'N' as Prod "
'            End If
            sql = sql & " from    ("
            sql = sql & " tmp0"
            sql = sql & " Right join gncomprobante Gnc "
            sql = sql & " inner join GnTrans gnt on gnc.codtrans=GnT.codtrans"
            sql = sql & " INNER JOIN PCPROVClI pc "
            sql = sql & " left join anexo_TipoDocumento atd on pc.IdTipoDocumento=atd.Id"
            sql = sql & " ON GNC.IdClienteRef=pc.idprovcli "
            sql = sql & " inner join " & Vista & " vw"

            sql = sql & " ON Gnc.TransID = vw.transid"
            sql = sql & " ON tmp0.TransID = Gnc.TransID)"
            sql = sql & " left join tmpRecarBaseCero tr ON tmp0.TransID = tr.TransID"
            sql = sql & " where  GNC.CodTrans IN (" & PreparaCadena(gobjCond.CodGrupo) & ")"
            sql = sql & " AND Gnc.BandCierre<>9 "
            sql = sql & " and GNC.Estado<>3 " & cond
            sql = sql & " and  (atd.CodTipoDocumento not in ( 'R','C','P')"
            sql = sql & " OR pc.ruc='9999999999999')"
            sql = sql & " group by  AnexoCodTipoComp "
'            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
'                sql = sql & " , substring(gnt.opcion2,122,1), substring(gnt.opcion2,135,1) "
'            End If
           
           
           VerificaExistenciaTabla 1
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            
            sql = " select ruc  into tmp4 from tmp1 where anexocodtipocomp in (5,18) "
            
            VerificaExistenciaTabla 4
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            
            sql = " Select pck.idformaSRI, Count(pck.idformaSRI) as cant "
            sql = sql & " From GnComprobante gNc"
            sql = sql & " inner join pckardex pck"
            sql = sql & " on gnc.transid= pck.transid"
            sql = sql & " Where gnc.Estado <> 3"
            sql = sql & " AND (GNC.fechatrans BETWEEN " & _
                FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = sql & " AND gnc.CodTrans IN (" & PreparaCadena(.CodGrupo) & ")"
            sql = sql & " Group by pck.idformaSRI"
            sql = sql & " having Count(pck.idformaSRI)>0"
            sql = sql & " order by pck.idformaSRI "
            
           Set rsforma = gobjMain.EmpresaActual.OpenRecordset(sql)
            If Not rsforma.EOF Then
                rsforma.MoveFirst
            End If
            
            For m = 0 To rsforma.RecordCount - 1
                If Not IsNull(rsforma.Fields(0)) Then
                    sql = " Select"
                    'sql = sql & " case when idtipodocumento =5 then '9999999999999' else ruc end as ruc,  CodFormaPago "
                    sql = sql & " ruc ,  CodFormaPago "
                    sql = sql & " Into tmpForma" & m
                    sql = sql & " From GnComprobante gNc"
                    sql = sql & " inner join pckardex pck"
                    sql = sql & " left join Anexo_FormaPago a  on pck.idformasri=a.id  "
                    sql = sql & " on gnc.transid= pck.transid"
                    sql = sql & " inner join pcprovcli pc on gnc.idclienteref=pc.idprovcli"
                    sql = sql & " Where gnc.Estado <> 3"
                    sql = sql & " AND (GNC.fechatrans BETWEEN " & _
                        FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                        FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
                    sql = sql & " AND gnc.CodTrans IN (" & PreparaCadena(.CodGrupo) & ")"
                    sql = sql & " and pck.IdFormaSRI=" & rsforma.Fields(0)
                    sql = sql & " Group by pc.ruc,  CodFormaPago, idtipodocumento "
                
                    VerificaExistenciaTablaTemp "tmpForma" & m
                    gobjMain.EmpresaActual.EjecutarSQL sql, numReg
                Else
                    MsgBox "Error falta enlaces de Forma de Cobro con el Catalogo del SRI "
                    Exit Function
                End If
                rsforma.MoveNext
            Next m
            
            sql = " Select"
            sql = sql & " tmp1.ruc, "
            For m = 0 To rsforma.RecordCount - 1
                sql = sql & "tmpForma" & m & ".CodformaPago as Codforma" & m & ", "
            Next m
            sql = Mid(sql, 1, Len(sql) - 2)
            
            sql = sql & " into tmpTodasFormas"
            sql = sql & " from tmp1 "
            For m = 0 To rsforma.RecordCount - 1
                sql = sql & " left join  tmpForma" & m & " on tmp1.ruc = tmpForma" & m & ".ruc "
            Next m
            
            
            
            VerificaExistenciaTablaTemp "tmpTodasFormas"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            
            'sql = " select case when idtipodocumento =5 then '9999999999999' else ruc end as ruc,  "
            sql = " select ruc,  "
            sql = sql & " AnexoCodTipoComp, case when AnexoCodTipoComp=4 then sum(valor) *-1 else  sum(valor) end as PorcentajeIVA"
            sql = sql & " into tmpIVAS"
            sql = sql & " from gncomprobante g "
            sql = sql & " inner join GnTrans gnt on g.codtrans=GnT.codtrans "
            sql = sql & " inner join pcprovcli pc on g.idclienteref=pc.idprovcli"
            sql = sql & " inner join ivkardexrecargo i"
            sql = sql & " inner join ivrecargo iv"
            sql = sql & " on i.idrecargo=iv.idrecargo"
            sql = sql & " on g.transid=i.transid"
            sql = sql & " Where g.estado <>3"
            sql = sql & " AND (G.fechatrans BETWEEN " & _
                FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = sql & " AND g.CodTrans IN (" & PreparaCadena(.CodGrupo) & ")"
            
            sql = sql & " and codrecargo='IVA'"
            sql = sql & " group by ruc, AnexoCodTipoComp, idtipodocumento"
            
            
            
            VerificaExistenciaTablaTemp "tmpIVAS"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            sql = " select "
            sql = sql & " RUC , AnexoCodTipoComp, "
            sql = sql & "  sum(PorcentajeIVA) as PorcentajeIVA "
            sql = sql & " into tmpIVA"
            sql = sql & " from tmpIVAS "
            sql = sql & " group by RUC , AnexoCodTipoComp, "
            sql = sql & " AnexoCodTipoComp "
            
            
            VerificaExistenciaTablaTemp "tmpIVA"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            ' union de las tres tablas
            sql = "SELECT DISTINCT " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB) & " as Fecha, "
            sql = sql & "  CodTipoDocumento,  '' as i,  (a.RUC)as ruc, BandRelacionado, a.AnexoCodTipoComp, isnull(a.cantTrans,0) as cant,"
'            sql = sql & " isnull(a.Valor0*-1,0) , isnull(a.Valor12*-1,0), isnull(a.ValorNO12*-1,0) "
            sql = sql & " case when a.AnexoCodTipoComp=4 then  isnull(a.Valor0,0) else isnull(a.Valor0*-1,0) end as Valor0,"
            sql = sql & " case when a.AnexoCodTipoComp=4 then  isnull(a.Valor12,0) else isnull(a.Valor12*-1,0) end as Valor12,"
            sql = sql & " case when a.AnexoCodTipoComp=4 then isnull(a.ValorNO12,0) else  isnull(a.ValorNO12*-1,0) end as ValorNO12"
            'sql = sql & " ,round(b.totaliva,2) as totaliva, I.porcentajeiva,  "
            sql = sql & " ,round(b.totaliva,2) as totaliva, I.porcentajeiva,  "
            sql = sql & " isnull(TotalICE,0) as ice,"
            'sql = sql & "  case when ( a.CompElect ='S'  and  a.Prod ='N' ) then 'E' else 'F' end as TipoEmi,  "
'            sql = sql & rsforma.RecordCount & " as Numpagos,"
            
            For m = 0 To rsforma.RecordCount - 1
                sql = sql & " Codforma" & m & ",  "
            Next m
            
            For m = rsforma.RecordCount To 9
                sql = sql & "'' as forma" & m & ",  "
            Next m
            
            sql = Mid(sql, 1, Len(sql) - 3)
            
            sql = sql & " into tvta"
            sql = sql & " From tmp1 a  left  join tan1 b on a.ruc = b.ruc "
            sql = sql & " inner join tmpTodasFormas tf on a.ruc= tf.ruc "
            sql = sql & " left join tmpICE tice on a.ruc = tice.ruc " 'and a.CompElect= tice.CompElect   and  a.Prod  = tice.Prod "
            sql = sql & " left join tmpIVA i on a.ruc=i.ruc and a.AnexoCodTipoComp=i.AnexoCodTipoComp"
            sql = sql & " union all "
            'saca retenciones de este periodo afectadas a facturas de otro periodo
            sql = sql & " Select  DISTINCT " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB) & " as Fecha, "
            sql = sql & " atd.CodTipoDocumento, '' as i, (pc.RUC)as ruc, pc.BandRelacionado, '18' as AnexoCodTipoComp, 0 as cantTrans,"
            sql = sql & "0 as Valor0, 0 as Valor12 , 0 as ValorNO12 ,0 as totaliva, 0 AS  porcentajeiva, "
            sql = sql & " 0 as ice,"
'            If gobjMain.EmpresaActual.GNOpcion.BandFactElect Then
'                sql = sql & "  'E' as CompElect, "
'            Else
'                sql = sql & "  'F' as CompElect, "
'            End If
 '           sql = sql & " 0 as Numpagos,"
            For m = 0 To 9
                sql = sql & " '' as Codforma" & m & ",  "
            Next m
            sql = Mid(sql, 1, Len(sql) - 3)
            
            sql = sql & " From    (gncomprobante GNC_b INNER JOIN PCPROVClI pc"
            sql = sql & " left join anexo_TipoDocumento atd on pc.IdTipoDocumento=atd.Id"
            sql = sql & " ON GNC_b.IdClienteRef=pc.idprovcli   )"
            sql = sql & " inner join tskardexret ivr on gnc_b.transid = ivr.transid"
            sql = sql & " where GNC_b.CodTrans IN (" & PreparaCadena(gobjCond.CodRetencion2) & ")"
            sql = sql & " AND (GNC_b.FechaTrans  BETWEEN " & _
                FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            sql = sql & " and GNC_B .Estado<>3 and GNC_b.IdClienteRef<>0 "
            sql = sql & " and pc.ruc not in (select  ruc from tmp4) "
            'sql = sql & " group by atd.CodTipoDocumento"
            sql = sql & " order by a.AnexoCodTipoComp " ', a.Nombre
            
            VerificaExistenciaTablaTemp "tvta"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
            sql = " select pc.ruc, '18' as AnexoCodTipoComp, "
            sql = sql & " case when bandiva=1 then sum(base) else 0 end as BaseIVA ,"
            sql = sql & " case when bandiva=1 then sum(debe) else 0 end as ValorRetIVA,"
            sql = sql & " case when bandiva<>1 then sum(base) else 0 end as BaseRenta ,"
            sql = sql & " case when bandiva<>1 then sum(debe) else 0 end as ValorRetRenta"
            sql = sql & " Into treten"
            sql = sql & " from gncomprobante g"
            sql = sql & " inner join pcprovcli pc on g.idclienteref= pc.idprovcli"
            sql = sql & " inner join tskardexret tsk"
            sql = sql & " inner join tsretencion t"
            sql = sql & " on t.IdRetencion=tsk.IdRetencion on g.transid=tsk.transid"
            sql = sql & " Where g.Estado <> 3"
            sql = sql & " and bandcliente=1"
            sql = sql & " and bandiva=1"
                        sql = sql & " AND (G.FechaTrans  BETWEEN " & _
                            FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                            FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            
            sql = sql & " group by pc.ruc, bandIVA"
            sql = sql & " Union All"
            sql = sql & " select pc.ruc, '18' as AnexoCodTipoComp, "
            sql = sql & " case when bandiva=1 then sum(base) else 0 end as BaseIVA ,"
            sql = sql & " case when bandiva=1 then sum(debe) else 0 end as ValorRetIVA,"
            sql = sql & " case when bandiva<>1 then sum(base) else 0 end as BaseRenta ,"
            sql = sql & " case when bandiva<>1 then sum(debe) else 0 end as ValorRetRenta"
            sql = sql & " from gncomprobante g"
            sql = sql & " inner join pcprovcli pc on g.idclienteref= pc.idprovcli"
            sql = sql & " inner join tskardexret tsk"
            sql = sql & " inner join tsretencion t"
            sql = sql & " on t.IdRetencion=tsk.IdRetencion"
            sql = sql & " on g.transid=tsk.transid"
            sql = sql & " Where g.Estado <> 3"
            sql = sql & " and bandcliente=1"
            sql = sql & " and bandiva=0"
                        sql = sql & " AND (G.FechaTrans  BETWEEN " & _
                            FechaYMD(.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & _
                            FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") "
            
            sql = sql & " group by pc.ruc, bandIVA"
            
            VerificaExistenciaTablaTemp "treten"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            sql = "select ruc, AnexoCodTipoComp, sum(BaseIVA) as BaseIVA , sum(ValorRetIVA) as ValorRetIVA, sum(BaseRenta) as BaseRenta, sum(ValorRetRenta) as ValorRetRenta "
            sql = sql & " Into tret"
            sql = sql & " from treten group by ruc, AnexoCodTipoComp"
            
            VerificaExistenciaTablaTemp "tret"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
            
        sql = " select v.*, BaseIVA, ValorRetIVA, BaseRenta, ValorRetRenta  from tvta v full join  tret r on v.ruc=r.ruc"
        sql = sql & " WHERE V.AnexoCodTipoComp = (18) "
        sql = sql & " UNION ALL"
        sql = sql & " select v.*, 0 as BaseIVA, 0 as ValorRetIVA, 0 as BaseRenta, 0 as ValorRetRenta  from tvta v full join  tret r on v.ruc=r.ruc"
        sql = sql & " WHERE V.AnexoCodTipoComp = (5) "
        sql = sql & " UNION ALL"
        
        sql = sql & " select v.*,  '' AS BaseIVA,''  AS ValorRetIVA, '' AS BaseRenta, '' AS ValorRetRenta  from tvta V "
        sql = sql & " WHERE V.AnexoCodTipoComp=4 "

        sql = " select 'FC' AS TRANS, RIGHT('0' + v.ANEXOCODTIPOCOMP,2) AS ANEXOCODTIPOCOMP , SUM(v.cant) AS CANTTRANS,"
        sql = sql & " SUM(VALOR0) AS VALOR0, SUM(VALOR12) AS VALOR12, SUM(VALORNO12) AS VALORNO12, SUM(porcentajeiva) AS TOTALIVA,  SUM(ValorRetIVA) AS ValorRetIVA, SUM(ValorRetRenta ) AS ValorRetRenta"
        sql = sql & " from tvta v full join  tret r on v.ruc=r.ruc WHERE V.AnexoCodTipoComp = (18)"
        sql = sql & " GROUP BY v.ANEXOCODTIPOCOMP"
        sql = sql & " Union All"
        sql = sql & " select 'FC' AS TRANS, RIGHT('0' + v.ANEXOCODTIPOCOMP,2) AS ANEXOCODTIPOCOMP, SUM(v.cant) AS CANTTRANS,"
        sql = sql & " SUM(VALOR0) AS VALOR0, SUM(VALOR12) AS VALOR12, SUM(VALORNO12) AS VALORNO12, SUM(porcentajeiva) AS TOTALIVA,  0 AS ValorRetIVA,   0 AS ValorRetRenta"
        sql = sql & " from tvta v full join  tret r on v.ruc=r.ruc WHERE V.AnexoCodTipoComp = (5)"
        sql = sql & " GROUP BY v.ANEXOCODTIPOCOMP"
        sql = sql & " Union All"
        sql = sql & " select 'FC' AS TRANS, RIGHT('0' + v.ANEXOCODTIPOCOMP,2) AS ANEXOCODTIPOCOMP, SUM(v.cant) AS CANTTRANS,"
        sql = sql & " SUM(VALOR0) AS VALOR0, SUM(VALOR12) AS VALOR12, SUM(VALORNO12) AS VALORNO12, SUM(porcentajeiva) AS TOTALIVA,  0 AS ValorRetIVA,   0 AS ValorRetRenta"
        sql = sql & " from tvta V  WHERE V.AnexoCodTipoComp=4"
        sql = sql & " GROUP BY v.ANEXOCODTIPOCOMP"
        sql = sql & " ORDER BY ANEXOCODTIPOCOMP "
        
        
            
    End With
    rs.Open sql, gobjMain.EmpresaActual.Coneccion, adOpenStatic, adLockReadOnly
    UltimoRecordset rs
    Set ConsANVentas2016paraTALON = rs
    Set rs = Nothing
    gobjMain.EmpresaActual.GrabaGNLogAccion "CONSULTA", "Cons Ventas 2016 para Anexos " & gobjCond.fecha1 & "-" & gobjCond.fecha2, "AN"
    VerificaExistenciaTabla 0
    VerificaExistenciaTabla 1
    VerificaExistenciaTabla 2
    VerificaExistenciaTabla 3
    VerificaExistenciaTabla 4
End Function


Public Function ConsANRetencionCompras2016paraTALON(BandIVA As Boolean) As Variant
    Dim sql As String, rs As Recordset, f1 As String
    Dim numReg As Long
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    With gobjCond
    
        f1 = DateSerial(Year(.fecha1), Month(.fecha1), 1)
    
    
        sql = " Select case when bandiva=0 then 'IR' ELSE 'IVA' END AS Tiporet, "
        sql = sql & " CASE WHEN ts.bandiva = 1 then right('000'+convert(varchar,iva.porcentaje),6) else ir.CodRetencion end as codretencion,"
        'sql = sql & " ir.CodRetencion as codretencion, "
        sql = sql & " count(tsk.base) as cant, sum(tsk.base) as base , sum(tsk.haber) as ValorREt"
        sql = sql & " Into tmprettalon"
        sql = sql & " from Gncomprobante gnc  inner join pckardex p on gnc.transid = p.transid inner join tskardexret tsk  inner join tsretencion ts "
        sql = sql & " LEFT join Anexos_Retencion_IR IR on ts.IdAnexo= IR.id"
        sql = sql & " LEFT join Anexos_Retencion_IVA IVA on ts.IdAnexo= IVA.id"
        sql = sql & " on tsk.idretencion=ts.idretencion on gnc.transid=tsk.transid"
        sql = sql & " Where gnc.Estado <> 3 and gnc.codtrans in (" & PreparaCadena(.CodRetencion1) & ")"
        sql = sql & " AND GNC.FechaTrans BETWEEN " & FechaYMD(f1, gobjMain.EmpresaActual.TipoDB)
        sql = sql & " AND " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB)
        If BandIVA Then
            sql = sql & " and bandiva=1 "
        Else
            sql = sql & " and bandiva=0 "
        End If
        sql = sql & " group by ir.CodRetencion, bandiva, iva.porcentaje "
        If Not BandIVA Then
            sql = sql & " Union All"
            sql = sql & " Select 'IR' AS Tiporet ,codtiporetencion as codretencion,    count(i.costorealtotal) as cant, sum(i.costorealtotal) as base, 0 as ValorREt"
            sql = sql & " from gncomprobante g inner join anexos a on g.transid=a.transid inner join ivkardex i"
            sql = sql & " inner join ivinventario ivi on i.idinventario=ivi.idinventario on g.transid=i.transid"
            sql = sql & " Where G.Estado <> 3"
            sql = sql & " AND G.CodTrans IN (" & PreparaCadena(.CodTrans) & ")"
            sql = sql & " AND G.FechaTrans BETWEEN " & FechaYMD(f1, gobjMain.EmpresaActual.TipoDB)
            sql = sql & " AND " & FechaYMD(DateAdd("m", 1, f1) - 1, gobjMain.EmpresaActual.TipoDB)
            sql = sql & " and len(codtiporetencion)>0"
            sql = sql & " group by codtiporetencion"
        End If
        sql = sql & " order by Tiporet , codretencion"
        
            VerificaExistenciaTablaTemp "tmprettalon"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
        
        
    End With
    
    sql = " select Tiporet , codretencion, sum(cant) as cant, sum(base) as base, sum(ValorREt) as ValorREt"
    sql = sql & " from tmprettalon"
    sql = sql & " group by Tiporet , codretencion"
    
    rs.Open sql, gobjMain.EmpresaActual.Coneccion, adOpenStatic, adLockReadOnly
    UltimoRecordset rs
    Set ConsANRetencionCompras2016paraTALON = rs
    Set rs = Nothing
    gobjMain.EmpresaActual.GrabaGNLogAccion "CONSULTA", "Cons Ventas 2016 para Anexos " & gobjCond.fecha1 & "-" & gobjCond.fecha2, "AN"

End Function

Public Function ConsTSPorCobrarPagarPorVendedor( _
                    ByVal PorCobrar As Boolean, _
                    ByVal FechaVenci As Date, _
                    ByVal CodForma As String, _
                    ByVal CodMoneda As String, _
                    ByVal codUsuario As String) As Variant
    Dim aux_codpc As String, aux_f1 As Date, aux_f2 As Date
    Dim aux_PorCobrar As Boolean, aux_Cliente As Boolean, aux_Proveedor As Boolean
    Dim aux_CodForma As String, aux_CodMoneda As String, aux_codgrupo As String
    Dim aux_incluircero As Boolean, aux_codUsuario As String
   With gobjMain.objCondicion
        'Conservar valores actuales para que no afecte a otras consultas
        aux_f1 = .fecha1
        aux_f2 = .fecha2
        aux_PorCobrar = .PorCobrar
        aux_CodForma = .CodForma
        aux_CodMoneda = .CodMoneda
        aux_codgrupo = .CodGrupo
        aux_codUsuario = .Usuario1

        .PorCobrar = PorCobrar
        .fecha1 = gobjMain.EmpresaActual.GNOpcion.FechaInicio
        .fecha2 = FechaVenci
        .CodForma = CodForma
        .CodMoneda = CodMoneda
        .Usuario1 = codUsuario

        Set ConsTSPorCobrarPagarPorVendedor = ConsTSPorCobrarPagarPorVendedorCierreCaja()
        
        'Recuperar valores reservados
        .fecha1 = aux_f1
        .fecha2 = aux_f2
        .PorCobrar = aux_PorCobrar
        .CodForma = aux_CodForma
        .CodMoneda = aux_CodMoneda
        .Usuario1 = aux_codUsuario
   End With
End Function


Public Function ConsTSPorCobrarPagarPorVendedorCierreCaja() As Variant
    Dim sql As String, cond As String, numMoneda As Integer, rs As Recordset
    Dim objCond As Condicion
    
    Set objCond = gobjMain.objCondicion
    
    'Si no está especificada la moneda, coge la Moneda predeterminada
    If Len(objCond.CodMoneda) = 0 Then objCond.CodMoneda = MONEDA_PRE
    numMoneda = gobjMain.EmpresaActual.GNOpcion.IndiceMoneda(objCond.CodMoneda)
  


        sql = " SELECT   Id,"
        sql = sql & " CodProvCli,"
        sql = sql & " Nombre,"
        sql = sql & " TransID,"
        sql = sql & " CodTrans,"
        sql = sql & " Trans,"
        sql = sql & " CodForma,"
        sql = sql & " NumLetra,"
        sql = sql & " Doc,"
        sql = sql & " Valor" & numMoneda & " AS Valor,"
        sql = sql & " VCancelado" & numMoneda & " AS VCancelado,"
        sql = sql & " Valor" & numMoneda & " - VCancelado" & numMoneda & " AS Saldo,"
        sql = sql & " CodMoneda,"
        sql = sql & " Valor" & numMoneda & " - VCancelado" & numMoneda & " AS CANCELA,"
        If numMoneda = 1 Then
            sql = sql & "  1 AS Cotizacion,"
        Else
            sql = sql & "  Cotizacion" & numMoneda & " AS Cotizacion,"
        End If
        sql = sql & " FechaEmision,"
        sql = sql & " Plazo,"
        sql = sql & " FechaVenci,"
        sql = sql & " codvendedor,"
        sql = sql & " NomVendedor"
        sql = sql & " From vwConsPCPorCobrarPagar"
        sql = sql & " WHERE  (FechaEmision between  " & FechaYMD(objCond.fecha2, gobjMain.EmpresaActual.TipoDB) & " and " & FechaYMD(objCond.fecha2, gobjMain.EmpresaActual.TipoDB) & " ) AND"
        sql = sql & " (CodVendedor='" & objCond.Usuario1 & "') AND"
        sql = sql & " (CodMoneda='" & objCond.CodMoneda & "') AND"
        sql = sql & " (CodForma in (" & objCond.CodForma & ")) AND"
        
        sql = sql & " (PorCobrar = " & CadenaBool(objCond.PorCobrar, gobjMain.EmpresaActual.TipoDB) & ") AND"
        sql = sql & " (Valor" & numMoneda & " - VCancelado" & numMoneda & ") >0 "
        sql = sql & " AND   ( (BandProveedor = 0) OR (BandCliente = 1) )"
        sql = sql & " ORDER BY CodProvCli, FechaVenci"
        



    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    Set ConsTSPorCobrarPagarPorVendedorCierreCaja = rs
    Set rs = Nothing
    Set objCond = Nothing
End Function

Public Function ListaIVInventarioSimplesoloServicios() As Variant
    Dim sql As String, rs As Recordset
    
    sql = "SELECT CodInventario, Descripcion FROM IVInventario where bandservicio=1 ORDER BY CodInventario"
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    ListaIVInventarioSimplesoloServicios = MiGetRows(rs)
    rs.Close
    Set rs = Nothing
End Function

Public Function RecuperaDatosAnexos(CodTrans As String, NumTrans As Currency) As Variant
End Function

Public Function ListaGNSeguro(BandRS As Boolean, Valida As Boolean) As Variant
    Dim sql As String, rs As Recordset
    Dim cond As String
    With gobjMain.objCondicion
        sql = "Select codSeguro,descripcion"
        sql = sql & " from  GNSeguro"
        If Valida Then
            sql = sql & " Where bandvalida = 1 "
        End If
        sql = sql & " order by codseguro"
    End With
   Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If BandRS Then
        Set ListaGNSeguro = rs
    Else
        ListaGNSeguro = MiGetRows(rs)
        rs.Close
    End If
    Set rs = Nothing
End Function

Public Function ConsIVCorrecExistDou() As Variant
    Dim sql As String, cond As String, s As String, rs As Recordset
    Dim ex1 As String, ex2 As String, ex3 As String, ex4 As String
    Dim ex5 As String, ex6 As String
    ex1 = "isnull(round(ive.ExistDou,8),0)"
    ex2 = "isnull(iv.IdInventario, iv2.IdInventario)"
    ex3 = "isnull(iv.CodInventario, iv2.CodInventario)"
    ex4 = "isnull(round(vw.SumaDeCantidad,8),0)"
    ex5 = "isnull(ivb.IdBodega, ivb2.IdBodega)"
    ex6 = "isnull(ivb.CodBodega, ivb2.CodBodega)"
     sql = "SELECT " & ex2 & " AS IdInventario, " & ex3 & " AS CodInventario, " & _
                 ex5 & " AS IdBodega, " & ex6 & " AS CodBodega, " & ex1 & " AS Exist1, " & _
                 ex4 & " AS SumaDeCantidad, (" & ex1 & " - " & ex4 & ") AS Dif " & _
          "FROM IVInventario iv2 RIGHT JOIN " & _
                    "(IVInventario iv RIGHT JOIN " & _
                        "(IVBodega ivb2 RIGHT JOIN " & _
                            "(IVBodega ivb RIGHT JOIN "
                    sql = sql & "(vwConsIVKardexSumaDou vw FULL JOIN IVExist ive "
                    sql = sql & "ON (vw.IdBodega = ive.IdBodega) AND (vw.IdInventario = ive.IdInventario) ) " & _
                            "ON ivb.IdBodega = vw.IdBodega) " & _
                        "ON ivb2.IdBodega = ive.IdBodega) " & _
                    "ON iv.IdInventario=vw.IdInventario) " & _
                "ON iv2.IdInventario = ive.IdInventario " & _
          "WHERE (" & ex1 & " - " & ex4 & ")<>0 " & _
                 "AND " & ex2 & " is not null " & _
          "ORDER BY " & ex3
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    Set ConsIVCorrecExistDou = rs
    Set rs = Nothing
End Function

Public Function ListaGNSeguro2(BandSoloValida As Boolean, BandRS As Boolean) As Variant
    Dim sql As String, rs As Recordset
    sql = "SELECT gns.CodSeguro,gns.Descripcion,ivr.codRecargo,gns.Porc  "
    sql = sql & " FROM GNSeguro gns Left Join IVRecargo ivr on ivr.idrecargo=gns.idrec "
    sql = sql & " ORDER BY CodSeguro "
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If BandRS Then
        Set ListaGNSeguro2 = rs
    Else
        ListaGNSeguro2 = MiGetRows(rs)
        rs.Close
    End If
    Set rs = Nothing
End Function

Public Function ListaIVRecargo(BandRS As Boolean, Valida As Boolean) As Variant
    Dim sql As String, rs As Recordset
    Dim cond As String
    With gobjMain.objCondicion
        sql = "Select codRecargo,descripcion"
        sql = sql & " from ivrecargo"
'        If Valida Then
'            sql = sql & " Where bandvalida = 1 "
'        End If
        sql = sql & " order by codrecargo"
    End With
   Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If BandRS Then
        Set ListaIVRecargo = rs
    Else
        ListaIVRecargo = MiGetRows(rs)
        rs.Close
    End If
    Set rs = Nothing
End Function
Public Function CreaGNSeguro() As GNSeguro
    Dim obj As GNSeguro
    Set obj = New GNSeguro
    Set obj.Empresa2 = Me
    Set CreaGNSeguro = obj
    Set obj = Nothing
End Function

Public Function ListaIVGrupoProduccionFlexNew(NumGrupo As Integer, CodBodega As String) As String
    Dim sql As String, s As String
    Dim rs As Recordset
    sql = "SELECT CodGrupo" & NumGrupo & ",IVGrupo" & NumGrupo & ".Descripcion,ive.exist FROM IVGrupo" & NumGrupo
    sql = sql & " Inner join IVInventario iv Inner join ivexist ive Inner join ivbodega ivb on ivb.idbodega = ive.idbodega on ive.idinventario = iv.idinventario on iv.idinventario = IVGrupo" & NumGrupo & ".iditemprod"
    sql = sql & " WHERE IVGrupo" & NumGrupo & ".BandValida=1 "
    sql = sql & " And BandProduccion=-1"
    sql = sql & " And ive.exist > 0"
    sql = sql & " And ivb.codbodega ='" & CodBodega & "'"
    sql = sql & " ORDER BY CodGrupo" & NumGrupo
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
#If DAOLIB Then
    If Not rs.EOF Then s = "|" & GetStringDAO(rs, 20 + 50)
#Else
    If Not rs.EOF Then s = "|" & rs.GetString(adClipString, , vbTab, "|", "(nulo)")
#End If
    ListaIVGrupoProduccionFlexNew = s
    rs.Close
    Set rs = Nothing
End Function

Public Function ListaGNTrans() As Variant
Dim sql As String
Dim rs As Recordset
    'saco solo los q estan en cola todavia
    sql = "Select gn.Transid,gn.FechaTrans,gn.codtrans,gn.numtrans From GNComprobante gn "
    sql = sql & " Left Join ivkProceso ivpk on ivpk.transid = gn.transid "
    sql = sql & " Where CodTrans = '" & gobjCond.CodTrans & "'"
    sql = sql & " And (FechaTrans between '" & gobjCond.fecha1 & "' And '" & gobjCond.fecha2 & "')"
    sql = sql & " And gn.estado <> 3"
    sql = sql & " And ivpk.estado1 NOT in (select valor from gnestadoprod where menu = 'Listo')"
    sql = sql & " Group by gn.Transid,gn.FechaTrans,gn.codtrans,gn.numtrans,gn.descripcion"
    sql = sql & " Order by fechatrans "
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    Set ListaGNTrans = rs
    Set rs = Nothing
End Function

Public Function ListaColas(ByVal transid As Long) As Variant
Dim sql As String
Dim rs As Recordset
    sql = "Select i.Idkp,iv.codProceso,i.Cantidad,i.Orden,i.Descripcion,i.FechaInicio,i.HoraInicio,i.FechaToma"
    sql = sql & ",i.HoraToma,i.FechaFin,i.HoraFin,i.FechaFinEspera,gne.codestado,i.IdCentroDet,gne1.codestado,UPPER(i.Grupo) as codusuario,i.idkpasignado,i.Secuencia"
    sql = sql & ",i.BandGarantia,i.BandUrgente,i.OrdenUrgente "
    sql = sql & " from ivkproceso i inner join gncomprobante g on g.transid = i.transid"
    sql = sql & " Inner join ivproceso iv on iv.idproceso = i.idproceso"
    sql = sql & " Left join gnestadoprod gne on gne.valor = i.estado"
    sql = sql & " Left join gnestadoprod gne1 on gne1.valor = i.estado1"
    sql = sql & " Where g.Transid = " & transid
    sql = sql & " Order by i.secuencia,iv.codproceso,i.idkp "
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    Set ListaColas = rs
    Set rs = Nothing
End Function

Public Function ListaFormulaRubro(ByVal transid As Long, ByVal bandSoloProc As Boolean) As Variant
Dim sql As String
Dim rs As Recordset
    
    sql = "select iv.codinventario,iv.descripcion,iv2.codinventario,iv2.descripcion"
    sql = sql & ",ivm.cantidad,ivp.codproceso,ivm.Orden from gncomprobante g"
    sql = sql & " Inner Join ivkardex ivk "
    sql = sql & " Inner join ivinventario iv "
    sql = sql & " Inner join ivmateria ivm "
    sql = sql & " Inner join ivinventario iv2 on iv2.idinventario = ivm.idmateria"
    sql = sql & " Left Join ivproceso ivp on ivp.idproceso = ivm.idproceso"
    sql = sql & " on ivm.idinventario = iv.idinventario"
    sql = sql & " on iv.idinventario = ivk.idinventario"
    sql = sql & " on ivk.transid = g.transid"
    sql = sql & " Where g.Transid = " & transid
    If bandSoloProc Then
        sql = sql & " and ivm.idproceso<> 0"
    End If
    sql = sql & " and iv.tipo= 6"
    sql = sql & " order by ivm.orden"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    Set ListaFormulaRubro = rs
    Set rs = Nothing
End Function

Public Function RecuperarAgenciaxCodprovcli(IdProvcli As String) As Boolean
    Dim sql As String, rs As Recordset
    RecuperarAgenciaxCodprovcli = False
    sql = "SELECT * FROM pcagencia  WHERE idprovcli=" & IdProvcli
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        RecuperarAgenciaxCodprovcli = True
    End If
    Set rs = Nothing
End Function

Public Function ListaProcesoFlexGrid(BandRS As Boolean) As Variant
   Dim sql As String, rs As Recordset
    Dim s As String
    sql = "select codproceso,descripcion "
    sql = sql & "from ivproceso "
    sql = sql & "  order by Secuencia "
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
#If DAOLIB Then
    If Not rs.EOF Then s = "|" & GetStringDAO(rs, 20 + 50)
#Else
    If Not rs.EOF Then s = "|" & rs.GetString(adClipString, , vbTab, "|", "(nulo)")
#End If
    ListaProcesoFlexGrid = s
    rs.Close
   Set rs = Nothing
End Function

Public Function ListaGNEstadoProdFlexGrid() As Variant
    Dim sql As String, rs As Recordset
    Dim s As String
    sql = "select codEstado,valor "
    sql = sql & "from GNEstadoProd "
    sql = sql & "Where bandvalida = 1 "
    sql = sql & "  order by valor "
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
#If DAOLIB Then
    If Not rs.EOF Then s = "|" & GetStringDAO(rs, 20 + 50)
#Else
    If Not rs.EOF Then s = "|" & rs.GetString(adClipString, , vbTab, "|", "(nulo)")
#End If
    ListaGNEstadoProdFlexGrid = s
    rs.Close
    Set rs = Nothing
End Function

Public Function ListaGNCentroCostoCCNew(BandSoloValida As Boolean, BandDetallado As Boolean, BandRS As Boolean) As Variant
    Dim sql As String, rs As Recordset
    
    If BandDetallado Then
        '*** MAKOTO 14/feb/01 Mod. aumentado Nombre
        sql = "SELECT DISTINCT GNCentroCosto.CodCentro, GNCentroCosto.Nombre, callpri.descripcion + 'No. '+ numcasa + ' y ' + callpri.descripcion, "
        sql = sql & " GNCentroCosto.Descripcion, FechaInicio,  "
        sql = sql & " GNCentroCosto.IdCentro, referencia2, BANDOCUPADO, BAndCorteMOra , codinventario "
        sql = sql & " FROM GNCentroCosto left  join pccalle callpri on GNCentroCosto.idcallepri =  callpri.idcalle left  "
        sql = sql & " join pccalle callsec on GNCentroCosto.idcallesec =  callsec.idcalle "
        sql = sql & "  left join vwcontratos vw on GNCentroCosto.codcentro=vw.contrato"
    Else
        sql = "SELECT CodCentro, Nombre, Descripcion, BANDOCUPADO FROM GNCentroCosto "
    End If
    
    If BandSoloValida Then sql = sql & " WHERE FechaFinal Is Null "
    sql = sql & " ORDER BY GNCentroCosto.Nombre"

    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If BandRS Then
        Set ListaGNCentroCostoCCNew = rs
    Else
        ListaGNCentroCostoCCNew = MiGetRows(rs)
        rs.Close
    End If
    Set rs = Nothing
End Function

Public Function ConsGNTransOT( _
                    Optional ByVal OrdenAscendente As Boolean) As Variant
    Dim sql As String, cond As String, s As String, rs As Recordset
    '*** MAKOTO 03/oct/00
    sql = "SELECT"
    'gobjCond.TopN = 1
    If gobjCond.TopN > 0 Then
        sql = sql & " TOP " & gobjCond.TopN
    End If
    '*** MAKOTO 05/feb/01 Mod. Aumentado 'NumDocRef' y 'Nombre'
    sql = sql & " vwConsGNTrans.TransID, vwConsGNTrans.FechaTrans, vwConsGNTrans.CodAsiento, vwConsGNTrans.CodTrans, "
    sql = sql & " vwConsGNTrans.NumTrans, "
    sql = sql & " vwConsGNTrans.codvehiculo, "
    sql = sql & " vwConsGNTrans.Nombre, vwConsGNTrans.Descripcion, vwConsGNTrans.CodCentro,"
    sql = sql & " vwConsGNTrans.Estado , "
    sql = sql & " case vwConsGNTrans.Estado when '0' THEN 'NO CONT' when '1' THEN 'APROV' when '2' THEN 'DESPA' when '3' THEN 'ANULA' when '4' THEN 'SEM DESP' END AS estadoL, "
    sql = sql & " vwConsGNTrans.CodPRAsiento,  "
    sql = sql & " ivi.descripcion,  fcv.nombre, ivk.nota "
    sql = sql & " FROM vwConsGNTrans "
    sql = sql & " Inner Join gntrans gnt on gnt.codtrans = vwConsGNTrans.CODTRANS "
    sql = sql & " inner join ivkardex ivk "
    sql = sql & " inner join ivinventario ivi on ivk.idinventario=ivi.idinventario"
    sql = sql & " on vwConsGNTrans.transid = ivk.transid"
    sql = sql & " INNER JOIN GNComprobante gn inner join PCProvCli pc "
    sql = sql & " ON pc.IdProvcli = gn.idClienteRef "
    sql = sql & " ON gn.transid = vwConsGNTrans.transid "
    sql = sql & " left join fcvendedor fcv on fcv.codvendedor= ivk.tiempoentrega "
    'Condición de FechaTrans
    cond = "vwConsGNTrans.FechaTrans BETWEEN " & FechaYMD(gobjCond.fecha1, gobjMain.EmpresaActual.TipoDB) & " AND " & FechaYMD(gobjCond.fecha2, gobjMain.EmpresaActual.TipoDB)
   'Condición de CodTrans
    If Len(cond) > 0 Then cond = cond & " AND "
    If Len(gobjCond.CodTrans) > 0 Then      'Si está especificado transaccion
        'Si es una sola transaccion
        If InStr(gobjCond.CodTrans, ",") = 0 Then
            cond = cond & "(vwConsGNTrans.CodTrans='" & gobjCond.CodTrans & "')"

        'Si está seleccionado varias trans.
       Else
            cond = cond & "(vwConsGNTrans.CodTrans IN (" & gobjCond.CodTrans & "))"
        End If
    Else
        'Si no hay condicion de cod.trans, selecciona todas las trans.
        ' que tenga permiso
        s = gobjMain.GrupoActual.PermisoActual.ListaTrans(True)
        If Len(s) > 0 Then
            cond = cond & "(vwConsGNTrans.CodTrans " & s & ")"
        Else
            cond = cond & "(vwConsGNTrans.CodTrans = '-------')"    'Para que NO devuelva nada si no tiene ningún permiso  '*** MAKOTO 26/dic/00
        End If
    End If
    'Condición de NumTrans
    If (gobjCond.NumTrans1 > 0) And (gobjCond.NumTrans2 > 0) Then
        If Len(cond) > 0 Then cond = cond & " AND "
        cond = cond & "(vwConsGNTrans.NumTrans BETWEEN " & gobjCond.NumTrans1 & " AND " & gobjCond.NumTrans2 & ")"
    ElseIf gobjCond.NumTrans1 > 0 Then
        If Len(cond) > 0 Then cond = cond & " AND "
        cond = cond & "(vwConsGNTrans.NumTrans>=" & gobjCond.NumTrans1 & ")"
    ElseIf gobjCond.NumTrans2 > 0 Then
        If Len(cond) > 0 Then cond = cond & " AND "
        cond = cond & "(vwConsGNTrans.NumTrans<=" & gobjCond.NumTrans2 & ")"
    End If
    'Condición de Estado
    If gobjCond.EstadoBool(ESTADO_NOAPROBADO) = False Then
        If Len(cond) > 0 Then cond = cond & " AND "
        cond = cond & "(vwConsGNTrans.Estado<>" & ESTADO_NOAPROBADO & ")"
    End If
    If gobjCond.EstadoBool(ESTADO_APROBADO) = False Then
        If Len(cond) > 0 Then cond = cond & " AND "
        cond = cond & "(vwConsGNTrans.Estado<>" & ESTADO_APROBADO & ")"
    End If
    If gobjCond.EstadoBool(ESTADO_DESPACHADO) = False Then
        If Len(cond) > 0 Then cond = cond & " AND "
        cond = cond & "(vwConsGNTrans.Estado<>" & ESTADO_DESPACHADO & ")"
    End If
    If gobjCond.EstadoBool(ESTADO_ANULADO) = False Then
        If Len(cond) > 0 Then cond = cond & " AND "
        cond = cond & "(vwConsGNTrans.Estado<>" & ESTADO_ANULADO & ")"
    End If
    '----------------------------------------
    '----- CAO 13/06/2006
    '-----------------------------------------
    If gobjCond.EstadoBool(ESTADO_SEMDESPACHADO) = False Then
        If Len(cond) > 0 Then cond = cond & " AND "
        cond = cond & "(vwConsGNTrans.Estado<>" & ESTADO_SEMDESPACHADO & ")"
   End If
    If Len(gobjCond.CodCentro1) > 0 And Len(gobjCond.CodCentro2) > 0 Then
        If Len(cond) > 0 Then cond = cond & " AND "
        cond = cond & " codCentro between '" & gobjCond.CodCentro1 & "' and '" & gobjCond.CodCentro2 & "'"
    End If
    If Len(gobjCond.Codciudad1) > 0 And Len(gobjCond.Codciudad2) > 0 Then
        If Len(cond) > 0 Then cond = cond & " AND "
        cond = cond & " vwConsGNTrans.codVendedor between '" & gobjCond.Codciudad1 & "' and '" & gobjCond.Codciudad2 & "'"
    End If
    If gobjMain.GrupoActual.PermisoActual.ConsRepVen Then
        If Len(cond) > 0 Then cond = cond & " AND "
        cond = cond & " vw.codvendedor = '" & gobjMain.UsuarioActual.codUsuario & "'  "
    End If
    If Len(gobjCond.CodPC1) > 0 Then ' And Not gobjCond.IncluirCero Then
        '---------------- ojo esto falta ANgel 15/may/2015
        '--------------------------
        If Not gobjCond.BandTodo Then
            If Len(cond) > 0 Then cond = cond & " AND "
            If gobjMain.GrupoActual.PermisoActual.ConsRepVen Then
                cond = cond & " pc.codprovcli = '" & gobjCond.CodPC1 & "' " 'AUC aqui porque tiene filtro de vendedor
        Else
                cond = cond & " pc.codprovcli = '" & gobjCond.CodPC1 & "' " 'AQUI con el filtro de la vista
        End If
       End If
    End If

   If Len(cond) > 0 Then sql = sql & " WHERE " & cond
   'Ordern de filas
    If OrdenAscendente Then
'        sql = sql & " ORDER BY FechaTrans, TransID"         '*** MAKOTO 03/oct/00 Modificado para que salga más actual al primero
        sql = sql & " ORDER BY vwConsGNTrans.FechaTrans, vwConsGNTrans.HoraTrans, vwConsGNTrans.TransID"         '*** MAKOTO 07/dic/00 Modificado para HoraTrans
    Else
'        sql = sql & " ORDER BY FechaTrans DESC, TransID DESC"
        sql = sql & " ORDER BY vwConsGNTrans.FechaTrans DESC, vwConsGNTrans.HoraTrans DESC, vwConsGNTrans.TransID DESC"       '*** MAKOTO 07/dic/00 HoraTrans agregado
    End If
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    Set ConsGNTransOT = rs
    Set rs = Nothing
'    Me.GrabaGNLogAccion "CONSULTA", "Consulta General OT" & gobjCond.fecha1 & "-" & gobjCond.fecha2, "GN"
End Function


Public Function CreaFCVCalendario() As FCVCalendario
    Dim obj As FCVCalendario
    Set obj = New FCVCalendario
    Set obj.Empresa2 = Me
    Set CreaFCVCalendario = obj
    Set obj = Nothing
End Function

Public Function RecuperaFCVCalendario(key As Variant, key1 As Variant) As FCVCalendario
    Dim obj As FCVCalendario
    Set obj = New FCVCalendario
   Set obj.Empresa2 = Me
    On Error Resume Next
    obj.Recuperar key, key1
    If Err.Number = 0 Then Set RecuperaFCVCalendario = obj
   Err.Clear
    On Error GoTo 0
    Set obj = Nothing
End Function

Public Function RecuperaTransxCodTransyCodCC(CodTrans As String, codCC As String) As Variant
    Dim sql As String, cond As String, s As String, rs As Recordset
    RecuperaTransxCodTransyCodCC = 0
    sql = "select transid from gncomprobante g inner join gncentrocosto gcc on g.idcentro = gcc.idcentro"
    sql = sql & " where codtrans='" & CodTrans & "' and codcentro='" & codCC & "'"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    RecuperaTransxCodTransyCodCC = rs.Fields(0)
    
End Function

Public Function CreaGNVCalendario() As GNVCalendario
    Dim obj As GNVCalendario
    Set obj = New GNVCalendario
    Set obj.Empresa2 = Me
    Set CreaGNVCalendario = obj
    Set obj = Nothing
End Function

Public Function RecuperaGNVCalendario(key As Variant, key1 As Variant) As GNVCalendario
    Dim obj As GNVCalendario
    Set obj = New GNVCalendario
   Set obj.Empresa2 = Me
    On Error Resume Next
    obj.Recuperar key, key1
    If Err.Number = 0 Then Set RecuperaGNVCalendario = obj
   Err.Clear
    On Error GoTo 0
    Set obj = Nothing
End Function

Public Function RecuperaCodyNumTrans(id As Long) As Variant
    Dim sql As String, cond As String, s As String, rs As Recordset
    RecuperaCodyNumTrans = 0
    sql = "select codtrans + '-'  + convert(varchar,numtrans) as trans from gncomprobante "
    sql = sql & " where transid=" & id
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    RecuperaCodyNumTrans = rs.Fields(0)
    
End Function

Public Function ListaGNRuta(BandSoloValida As Boolean, BandRS As Boolean) As Variant
    Dim sql As String, rs As Recordset
    sql = "SELECT CodRuta,Descripcion,DiaSemana "
    sql = sql & " FROM GNRuta "
    If BandSoloValida Then
        sql = sql & " Where bandvalida = 1"
    End If
    sql = sql & " ORDER BY CodRuta "
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If BandRS Then
        Set ListaGNRuta = rs
    Else
        ListaGNRuta = MiGetRows(rs)
        rs.Close
    End If
    Set rs = Nothing
End Function

Public Function CreaGNRuta() As GNRuta
    Dim obj As GNRuta
    Set obj = New GNRuta
    Set obj.Empresa2 = Me
    Set CreaGNRuta = obj
    Set obj = Nothing
End Function

Public Function RecuperaGNRuta(key As Variant) As GNRuta
    Dim obj As GNRuta
    Set obj = New GNRuta
   Set obj.Empresa2 = Me
    On Error Resume Next
    obj.Recuperar key
    If Err.Number = 0 Then Set RecuperaGNRuta = obj
   Err.Clear
    On Error GoTo 0
    Set obj = Nothing
End Function

Public Sub CorregirExistenciaDocumRuta()
    Dim rs As Recordset, sql As String, rs2 As Recordset
    Set rs = Me.ConsIVCorrecExistDocumRuta
    If rs.EOF Then Exit Sub
    With rs
        Do Until .EOF
            If Not IsNull(!IdRuta) Then
            'Verifica si existe el registro
            sql = "SELECT Exist FROM    GNExistDocumento " & _
                  " WHERE Transid=" & !transid & _
                  " AND IdRuta=" & !IdRuta
            'rs2.Open sql, gobjMain.EmpresaActual.mcnEmpresa, adOpenStatic, adLockReadOnly
            Set rs2 = gobjMain.EmpresaActual.OpenRecordset(sql)
            'Si no existe, lo inserta
            If rs2.EOF Then
                sql = "INSERT INTO GNExistDocumento (Transid,IdProvcli,Exist,Orden,IdRuta) " & _
                      "VALUES (" & !transid & ",0," & !SumaDeCantidad & ",1," & !IdRuta & ")"
            'Si existe, lo modifica
            Else
                sql = "UPDATE GNExistDocumento SET Exist=" & !SumaDeCantidad & _
                      " WHERE Transid=" & !transid & _
                      " AND IdRuta=" & !IdRuta
            End If
            gobjMain.EmpresaActual.Execute sql, True
            rs2.Close
            End If
            .MoveNext
        Loop
        .Close
    End With
    Set rs2 = Nothing
    Set rs = Nothing
End Sub


Public Function ConsIVCorrecExistDocumRuta() As Variant
    Dim sql As String, cond As String, s As String, rs As Recordset
    Dim ex1 As String, ex2 As String, ex3 As String, ex4 As String
    Dim ex5 As String, ex6 As String
    
'#If DAOLIB Then
'    ex1 = "IIF(IsNull(ive.Exist),0,ive.Exist)"
'    ex2 = "IIF(IsNull(iv.IdInventario), iv2.IdInventario, iv.IdInventario)"
'    ex3 = "IIF(IsNull(iv.CodInventario), iv2.CodInventario, iv.CodInventario)"
'    ex4 = "IIF(IsNull(vw.SumaDeCantidad),0, vw.SumaDeCantidad)"
'    ex5 = "IIF(IsNull(ivb.IdBodega), ivb2.IdBodega, ivb.IdBodega)"
'    ex6 = "IIF(IsNull(ivb.CodBodega), ivb2.CodBodega, ivb.CodBodega)"
'#Else
'    ex1 = "isnull(ive.Exist,0)"
'    ex2 = "isnull(ive.IdInventario, iv2.IdInventario)"
'    ex3 = "isnull(iv.CodInventario, iv2.CodInventario)"
'    ex4 = "isnull(vw.SumaDeCantidad,0)"
'    ex5 = "isnull(ivb.IdBodega, ivb2.IdBodega)"
'    ex6 = "isnull(ivb.CodBodega, ivb2.CodBodega)"
'#End If
'
'    'Muestra la diferencia de existencia y suma de IVKardex si es que la hay
'    sql = "SELECT " & ex2 & " AS IdInventario, " & ex3 & " AS CodInventario, " & _
'                 ex5 & " AS IdBodega, " & ex6 & " AS CodBodega, " & ex1 & " AS Exist1, " & _
'                 ex4 & " AS SumaDeCantidad, (" & ex1 & " - " & ex4 & ") AS Dif " & _
'          "FROM IVInventario iv2 RIGHT JOIN " & _
'                    "(IVInventario iv RIGHT JOIN " & _
'                        "(IVBodega ivb2 RIGHT JOIN " & _
'                            "(IVBodega ivb RIGHT JOIN "
'#If DAOLIB Then                         '*** MAKOTO 25/ene/01 Mod. porque en DAO no sporta FULL JOIN
'                    sql = sql & "(vwConsIVKardexSuma vw Right JOIN IVExist ive "
'#Else
'                    sql = sql & "(vwConsIVKardexSuma vw FULL JOIN IVExist ive "
'#End If
'                    sql = sql & "ON (vw.IdBodega = ive.IdBodega) AND (vw.IdInventario = ive.IdInventario) ) " & _
'                            "ON ivb.IdBodega = vw.IdBodega) " & _
'                        "ON ivb2.IdBodega = ive.IdBodega) " & _
'                    "ON iv.IdInventario=vw.IdInventario) " & _
'                "ON iv2.IdInventario = ive.IdInventario " & _
'          "WHERE (" & ex1 & " - " & ex4 & ")<>0 " & _
'                 "AND " & ex2 & " is not null " & _
'          "ORDER BY " & ex3
    
sql = "SELECT"
sql = sql & " isnull(ive.transId, vw.transIdDoc) AS Transid, "
sql = sql & " isnull(ivb.Id, ivb2.Id) AS IdRuta, "
sql = sql & " isnull(ivb.CodRuta, ivb2.CodRuta) AS CodRuta, "
sql = sql & " isnull(ive.Exist,0) AS Exist1,"
sql = sql & " isnull(vw.SumaDeCantidad,0) AS SumaDeCantidad, "
sql = sql & " (isnull(ive.Exist,0) - isnull(vw.SumaDeCantidad,0)) AS Dif "
sql = sql & " FROM Gncomprobante iv2 RIGHT JOIN (Gncomprobante iv "
sql = sql & " RIGHT JOIN (GnRuta ivb2 RIGHT JOIN (GnRuta ivb "
sql = sql & " RIGHT JOIN (vwConsIVKardexSumaDocumRuta vw"
sql = sql & " FULL JOIN GNExistDocumento ive"
sql = sql & " ON (vw.IdRuta= ive.IdRuta) AND (vw.transidDoc = ive.transid) )"
sql = sql & " ON ivb.Id= vw.IdRuta)"
sql = sql & " ON ivb2.Id= ive.IdRuta)"
sql = sql & " ON iv.transid=vw.transidDoc)"
sql = sql & " ON iv2.transid = ive.transid"
sql = sql & "  Where (isnull(ive.exist, 0) - isnull(vw.SumaDeCantidad, 0)) <> 0"
sql = sql & " AND isnull(ive.transid, vw.transidDoc) is not null  AND vw.SumaDeCantidad>0 "
sql = sql & " "
sql = sql & " "
sql = sql & " "
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    Set ConsIVCorrecExistDocumRuta = rs
    Set rs = Nothing
End Function


Public Function ListaGNRutaFlexGrid(BandSoloValida As Boolean) As Variant
    Dim sql As String, rs As Recordset
    Dim s As String
    sql = "SELECT CodRuta,Descripcion,DiaSemana "
    sql = sql & " FROM GNRuta "
    If BandSoloValida Then
        sql = sql & " Where bandvalida = 1"
    End If
    sql = sql & " ORDER BY CodRuta "
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    #If DAOLIB Then
    If Not rs.EOF Then s = "|" & GetStringDAO(rs, 20 + 50)
#Else
    If Not rs.EOF Then s = "|" & rs.GetString(adClipString, , vbTab, "|", "(nulo)")
#End If
    ListaGNRutaFlexGrid = s
    rs.Close
    Set rs = Nothing
End Function

Public Function ListaIVBodegaH(BandSoloValida As Boolean, BandRS As Boolean) As Variant
    Dim sql As String, rs As Recordset
    
    sql = "SELECT IVBodega.CodBodega,IVBodega.Descripcion, codsucursal, ivbPadre.codbodega as CodBodegaHijo, ivbodega.BandValida "
    sql = sql & " FROM IVBodega left join gnsucursal on ivbodega.idsucursal=gnsucursal.idsucursal"
    sql = sql & " left join IVBodega ivbPadre on ivbodega.idBodegaHijo=ivbPadre.idBodega"
    If BandSoloValida Then
        sql = sql & " WHERE ivbodega.BandValida=" & CadenaBool(True, gobjMain.EmpresaActual.TipoDB)
                                                                            
    End If
    sql = sql & " where ivbodega.BandBodegaHijo=0 "
    sql = sql & " ORDER BY ivbodega.CodBodega"
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If BandRS Then
        Set ListaIVBodegaH = rs
    Else
        ListaIVBodegaH = MiGetRows(rs)
        rs.Close
    End If
    Set rs = Nothing
End Function

Public Function ListaIVBodegaSoloH(BandSoloValida As Boolean, BandRS As Boolean) As Variant
    Dim sql As String, rs As Recordset
    sql = "SELECT CodBodega,Descripcion, codsucursal, ivbodega.BandValida FROM IVBodega left join gnsucursal on ivbodega.idsucursal=gnsucursal.idsucursal"
    sql = sql & " WHERE BandBodegaHijo=" & CadenaBool(True, gobjMain.EmpresaActual.TipoDB)
    If BandSoloValida Then
        sql = sql & " AND ivbodega.BandValida=" & CadenaBool(True, gobjMain.EmpresaActual.TipoDB)
    End If
    sql = sql & " ORDER BY CodBodega"
    
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If BandRS Then
        Set ListaIVBodegaSoloH = rs
    Else
        ListaIVBodegaSoloH = MiGetRows(rs)
        rs.Close
    End If
    Set rs = Nothing
End Function

Public Function ObtieneIdBodegaPadre(CodBodega As String) As String
    Dim sql As String, rs As Recordset
    sql = "SELECT ip.codbodega from ivbodega ip left join ivbodega ih on ip.idbodegahijo=ih.idbodega "
    sql = sql & " where ih.codbodega='" & CodBodega & "'"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        ObtieneIdBodegaPadre = rs.Fields(0)
    Else
        ObtieneIdBodegaPadre = ""
    End If
    Set rs = Nothing
End Function

Public Function ListaGrupo(BandRS As Boolean) As Variant
'BandRS : True  --> Devuelve Recordset de ADO
'      False --> Devuelve en Array de Variant
    Dim sql As String, rs As Recordset
    sql = "SELECT CodGrupo,Descripcion " & _
          "FROM " & gobjMain.NombreDBSQL & ".dbo.Grupo "
'    If BandSoloValida Then sql = sql & "WHERE BandValida=" & CadenaBool(True, TipoDB)
    sql = sql & " ORDER BY CodGrupo"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If BandRS Then
        Set ListaGrupo = rs
    Else
        ListaGrupo = MiGetRows(rs)
        rs.Close
    End If
    Set rs = Nothing
End Function

Public Function RecuperaUltimaSacaGuia(ByVal CodTrans As String) As Long
    Dim sql As String, s As String
    Dim rs As Recordset, rs1 As Recordset
    sql = " select  top 1 Guia "
    sql = sql & " from GNSACAHIST gh Inner join gncomprobante g on g.transid = gh.transid "
    sql = sql & " Where g.estado <> 3 and g.codtrans = '" & CodTrans & "'"
    sql = sql & " order by Id desc "
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        RecuperaUltimaSacaGuia = rs.Fields("Guia")
    Else
        RecuperaUltimaSacaGuia = 0
    End If
   rs.Close
    Set rs = Nothing
End Function

Public Function RecuperaUltimaSacaAerolinea(ByVal CodTrans As String) As Long
    Dim sql As String, s As String
    Dim rs As Recordset, rs1 As Recordset
    sql = " select  top 1 Aerolinea "
    sql = sql & " from GNSACAHIST gh Inner join gncomprobante g on g.transid = gh.transid "
    sql = sql & " Where g.estado <> 3 and g.codtrans = '" & CodTrans & "'"
    sql = sql & " order by Id desc "
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        RecuperaUltimaSacaAerolinea = rs.Fields("Aerolinea")
    Else
        RecuperaUltimaSacaAerolinea = 0
    End If
    rs.Close
    Set rs = Nothing
End Function

Public Function ListaGNTransxPantalla( _
                    CodPantalla As String, _
                    BandInvisibles As Boolean, _
                    BandRS As Boolean) As Variant
    Dim sql As String, cond As String, rs As Recordset
    sql = "SELECT CodTrans, NombreTrans FROM GNTrans "
    If Len(CodPantalla) > 0 Then cond = cond & "Codpantalla='" & CodPantalla & "' "
    If Not BandInvisibles Then
        If Len(cond) > 0 Then cond = cond & " AND "
        cond = cond & "BandVisible=" & CadenaBool(True, gobjMain.TipoDB)
    End If
    If Len(cond) > 0 Then sql = sql & "WHERE " & cond
    sql = sql & " ORDER BY CodTrans"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If BandRS Then
        Set ListaGNTransxPantalla = rs
    Else
        ListaGNTransxPantalla = MiGetRows(rs)
        rs.Close
   End If
    Set rs = Nothing
End Function

Public Function CargaBufferDista() As Variant
    Dim sql As String, cond As String, i As Integer, rs As Recordset, bienesmod As String
    Dim numReg As Long, numMoneda As String
    Dim L As String, n As String, O As String, p As String, Q As String, r As String, s As String, t As String, u As String, v As String, W As String, X As String, Y As String
    Dim Lm3 As String
    With gobjCond
        sql = "  select vw.idinventario, "
        sql = sql & "   sum(totalM3) as prom,"
        sql = sql & "  cast( sum(totalCANT)  as float) as promC"
        sql = sql & " Into TMPvsFin"
        sql = sql & " from vwVentasItemxSemanaxSucursal vw inner join ivinventario  on vw.idinventario=ivinventario.idinventario"
        sql = sql & " inner JOIN GNSucursal gsuc on gsuc.idsucursal = vw.idsucursal "
        sql = sql & " where SemanaAnio between    DatePart(wk, " & FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") - " & .NumEsta
        sql = sql & " and  DatePart(wk, " & FechaYMD(Date, gobjMain.EmpresaActual.TipoDB) & ")"
        If Len(.Sucursal) > 0 Then
            sql = sql & " AND gsuc.codsucursal IN (" & PreparaCadena(.Sucursal) & ")"
        End If
        If Not (Len(.Bienes)) = 0 Then
            sql = sql & .Bienes  'Aquí se ha grabado SQL de ítems
        End If
        sql = sql & "And vw.idbodega in (" & .CodBodega1 & ")"
        'sql = sql & " AND CODINVENTARIO = 'NT1A7.8.09MWE'"
        sql = sql & " group by vw.idinventario"
        VerificaExistenciaTablaTemp "TMPvsFin"
        gobjMain.EmpresaActual.EjecutarSQL sql, numReg
        
        sql = "  select vw.idinventario, "
        sql = sql & "   sum(totalM3) as prom, "
        sql = sql & "   cast(sum(totalCANT) as float) as promC"
        sql = sql & " Into TMPvs8sem"
        sql = sql & " from vwVentasItemxSemanaxSucursal vw inner join ivinventario  on vw.idinventario=ivinventario.idinventario"
        sql = sql & " inner JOIN GNSucursal gsuc on gsuc.idsucursal = vw.idsucursal "
        sql = sql & " where SemanaAnio between    DatePart(wk, " & FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") - " & .NumAuto
        sql = sql & " and  DatePart(wk, " & FechaYMD(Date, gobjMain.EmpresaActual.TipoDB) & ")"
        If Len(.Sucursal) > 0 Then
            sql = sql & " AND gsuc.codsucursal IN (" & PreparaCadena(.Sucursal) & ")"
        End If
        
        If Not (Len(.Bienes)) = 0 Then
            sql = sql & .Bienes  'Aquí se ha grabado SQL de ítems
        End If
         sql = sql & "And vw.idbodega in (" & .CodBodega1 & ")"
         'sql = sql & " AND CODINVENTARIO = 'NT1A7.8.09MWE'"
        sql = sql & " group by vw.idinventario"
        
        VerificaExistenciaTablaTemp "TMPvs8sem"
        gobjMain.EmpresaActual.EjecutarSQL sql, numReg

            cond = ""
            cond = cond & " AND gnt.anexocodtipotrans IN (2)"
            cond = cond & " AND gnt.anexocodtipoComp IN (4,18)"
            cond = cond & " AND gsuc.codsucursal IN (" & PreparaCadena(.Sucursal) & ")"
            
        If Not (Len(.Bienes)) = 0 Then
            cond = cond & .Bienes 'Aquí se ha grabado SQL de ítems
        End If
            '---------------
                VerificaExistenciaTablaTemp "TmpSuc"
                sql = "select ive.idinventario,ive.idgrupo"
                sql = sql & " into TmpSuc"
                sql = sql & " from ivexist ive"
                sql = sql & " Inner join ivbodega ivb on ivb.idbodega = ive.idbodega"
                sql = sql & " Inner join ivinventario on ive.idinventario = ivinventario.idInventario"
                sql = sql & " Where  ivb.idsucursal = (select idsucursal from gnsucursal where codsucursal IN (" & PreparaCadena(.Sucursal) & "))"
                If Not (Len(.Bienes)) = 0 Then
                    sql = sql & .Bienes 'Aquí se ha grabado SQL de ítems
                End If
                sql = sql & "And ivb.idbodega in (" & .CodBodega1 & ")"
'sql = sql & " AND CODINVENTARIO = 'NT1A7.8.09MWE'"
                sql = sql & " group by ive.idinventario,ive.idgrupo"
                gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            '---------------
       
            sql = " select"
            sql = sql & " ivk.idinventario, codinventario, codalterno1,  M3, ivinventario.DESCRIPCION, GSUC.CODSUCURSAL, "
            sql = sql & " SUM(CANTIDAD)*-1 AS CANT,"
            sql = sql & " sum(preciorealtotal)*-1 as pt, "
            sql = sql & " SUM(CANTIDAD*M3)*-1 AS TOTALM3, "
            sql = sql & " SUM(CANTIDAD)*-1 AS TOTALC, "
            'sql = sql & " TiempoPromVta "
            sql = sql & " ivg.IdealDias,datepart(ww,g.fechatrans) as SemanaAnio "
            sql = sql & " Into TMPvsV1"
            sql = sql & " from gncomprobante g inner join gntrans gnt "
            sql = sql & " inner JOIN GNSucursal gsuc on gsuc.idsucursal = GNT.idsucursal "
            sql = sql & " on g.codtrans=gnt.codtrans"
            sql = sql & " inner join ivkardex ivk  "
            sql = sql & " inner join ivinventario   Inner join tmpsuc t Inner join ivgrupo4 ivg on ivg.idgrupo4= t.idgrupo on t.idinventario =ivinventario.idinventario"
            sql = sql & " on ivinventario.idinventario = ivk.idinventario"
            sql = sql & " on g.transid=ivk.transid"
            sql = sql & " Where estado <>3"
            sql = sql & cond
            sql = sql & "And ivk.idbodega in (" & .CodBodega1 & ")"
            'sql = sql & " between    DatePart(wk, " & FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") - " & .NumEsta
            'sql = sql & " and  DatePart(wk, " & FechaYMD(Date, gobjMain.EmpresaActual.TipoDB) & ")"
            'sql = sql & " AND M3<>0"
            sql = sql & " group by ivk.idinventario, codinventario, codalterno1, M3, ivinventario.DESCRIPCION, GSUC.CODSUCURSAL,ivg.idealdias,g.fechatrans "
       
            VerificaExistenciaTablaTemp "TMPvsV1"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
               
            sql = " Select idinventario,codinventario,codalterno1,M3,DESCRIPCION,CODSUCURSAL,SUM(CANT) AS CANT ,SUM(pt) AS PT,SUM(TOTALM3) AS TOTALM3,SUM(TOTALC) AS TOTALC,IdealDias Into TMPvsV from TMPvsV1 "
            sql = sql & " Where Semanaanio "
            sql = sql & " between    DatePart(wk, " & FechaYMD(.fecha2, gobjMain.EmpresaActual.TipoDB) & ") - " & .NumEsta
            sql = sql & " and  DatePart(wk, " & FechaYMD(Date, gobjMain.EmpresaActual.TipoDB) & ")"
            sql = sql & " GROUP BY idinventario,codinventario,codalterno1,M3,DESCRIPCION,CODSUCURSAL,IdealDias"
            
            VerificaExistenciaTablaTemp "TMPvsV"
            gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
        sql = " select  "
        sql = sql & " ivinventario.idinventario, ivinventario.codinventario, ivinventario.codalterno1,ivg.idealdias as TiempoPromVta, "
        sql = sql & " ivinventario.M3, ivinventario.DESCRIPCION, CODSUCURSAL, ivg.codgrupo4 as codgrupo, '1' AS ML, "
        sql = sql & " isnull(CANT,0) as cant,"
        sql = sql & " isnull(pt,0) as pt, "
        sql = sql & " isnull(TOTALM3,0) as TOTALM3,  "
        sql = sql & " isnull(TOTALC,0) as TOTALC "
        sql = sql & " Into TMPvs"
        sql = sql & " from ivinventario Inner join tmpsuc tc Inner Join ivgrupo4 ivg on ivg.idgrupo4= tc.idgrupo on tc.idinventario =ivinventario.idinventario"
        sql = sql & " Inner join TMPvsV t on ivinventario.idinventario = t.idinventario"
        sql = sql & " where 1=1"
        sql = sql & .Bienes
        
        VerificaExistenciaTablaTemp "TMPvs"
        gobjMain.EmpresaActual.EjecutarSQL sql, numReg
                         
        ' cantidad pedida
         'sql = " select g.transid,ivk.idinventario, sum(cantidad) as cantpedido"
         sql = " select ivk.idinventario, sum(cantidad) as cantpedido"
         sql = sql & " Into TMPPedido"
         sql = sql & " from gncomprobante g"
        sql = sql & " inner join ivkardex ivk inner  "
        sql = sql & " join ivinventario  on ivk.idinventario=ivinventario.idinventario"
        sql = sql & " on g.transid = ivk.transid"
        sql = sql & " Where Estado = 0"
        'sql = sql & " and codtrans='" & .CodBanco1 & "'"
        sql = sql & " and codtrans='" & .CodTrans & "'"
        sql = sql & "And ivk.idbodega in (" & .CodBodega1 & ")"
        sql = sql & " group by ivk.idinventario"
        VerificaExistenciaTablaTemp "TMPPedido"
        gobjMain.EmpresaActual.EjecutarSQL sql, numReg
        
'        ' cantidad pedida modificada
'         sql = " select ivk.idinventario, sum(cantidad) as cantmodpedido"
'         sql = sql & " Into TMPModPedido"
'         sql = sql & " from gncomprobante g"
'        sql = sql & " inner join ivkardex ivk inner join ivinventario  on ivk.idinventario=ivinventario.idinventario"
'        sql = sql & " on g.transid = ivk.transid"
'        sql = sql & " Where Estado <> 3"
'        sql = sql & " and codtrans='" & .CodTrans & "M'"
'
'        sql = sql & " group by ivk.idinventario"
'        VerificaExistenciaTablaTemp "TMPModPedido"
'        gobjMain.EmpresaActual.EjecutarSQL sql, numReg

        ' cantidad Transito
         sql = " select g.transid,ivk.idinventario, sum(cantidad) as cantTransito"
         'sql = " select ivk.idinventario, sum(cantidad) as cantTransito"
         sql = sql & " Into TMPTrans"
         sql = sql & " from gncomprobante g"
        sql = sql & " inner join ivkardex ivk inner  "
        sql = sql & " join ivinventario  on ivk.idinventario=ivinventario.idinventario"
        sql = sql & " on g.transid = ivk.transid"
        sql = sql & " Where Estado <> 3"
        sql = sql & " and codtrans='" & .CodBanco1 & "'"
        sql = sql & "And ivk.idbodega in (" & .CodBodega1 & ")"
        sql = sql & " group by g.transid,ivk.idinventario"
       
        VerificaExistenciaTablaTemp "TMPTrans"
        gobjMain.EmpresaActual.EjecutarSQL sql, numReg

        ' cantidad Compra
         sql = " select ivk.idinventario, sum(cantidad) as cantCompra"
         sql = sql & " Into TMPComprada"
         sql = sql & " from gncomprobante g"
        sql = sql & " inner join ivkardex ivk inner join ivinventario  on ivk.idinventario=ivinventario.idinventario "
        sql = sql & " on g.transid = ivk.transid"
        sql = sql & " And  g.idtransfuente in (select transid from tmpTrans)"
        sql = sql & " Where Estado <> 3"
        sql = sql & " and codtrans='" & .CodComp & "'"
        sql = sql & "And ivk.idbodega in (" & .CodBodega1 & ")"
        sql = sql & " group by ivk.idinventario"
        VerificaExistenciaTablaTemp "TMPComprada"
        gobjMain.EmpresaActual.EjecutarSQL sql, numReg
       
       'pedido-comprada
         sql = " select isnull(tp.idinventario, tc.idinventario) as idinventario, "
         'sql = sql & " case when (isnull(cantpedido,0) +isnull(cantmodpedido,0)  - isnull(cantCompra,0) )> 0 then  (isnull(cantpedido,0)+isnull(cantmodpedido,0)  - isnull(cantCompra,0) ) else 0 end "
         sql = sql & " case when (isnull(cantTransito,0)  - isnull(cantCompra,0) )> 0 then  (isnull(cantTransito,0)  - isnull(cantCompra,0) ) else 0 end "
         sql = sql & " as transito"
         sql = sql & " Into TMPTransito"
        sql = sql & " from TMPTrans tp" ' left join tmpmodpedido tm on tp.idinventario = tm.idinventario "
        sql = sql & " full join TMPComprada tc on tp.idinventario = tc.idinventario "
            
        VerificaExistenciaTablaTemp "TMPTransito"
        gobjMain.EmpresaActual.EjecutarSQL sql, numReg
            
'            If Not .IncluirCero Then
                Lm3 = " (isnull(tf.prom ,0)/" & .NumEsta & " * " & .NumDias1 / 100 & "  + isnull(tf8.prom ,0)/" & .NumAuto & " * " & .NumDias2 / 100 & ")"
'            Else
                L = " (isnull(tf.promC ,0) /" & .NumEsta & " * " & .NumDias1 / 100 & "  + isnull(tf8.promC ,0)/" & .NumAuto & "* " & .NumDias2 / 100 & ")"
'            End If
            n = L
            O = " (Sum(ISNULL(exist,0)) * vs.M3) "  'existencia m3
            p = " Sum(ISNULL(exist,0)) " 'existencia unidades
'            Q = " isnull((transito*vs.M3),0) "
'            If Not .IncluirCero Then
'                Q = " isnull((cantpedido*vs.M3),0)  "
'            Else
                Q = " isnull(cantpedido,0)  "
'            End If
            'r = "case when " & n & ">0 then ((" & O & "/" & n & ") *30 )else 0 end "
            r = "case when " & n & ">0 then ((" & p & "/" & n & ") *7 )else 0 end "
'            If Not .IncluirCero Then
'                s = O & "-" & n
'            Else
                s = p & "-" & n
'            End If
            v = " TiempoPromVta "
            u = "((" & v & "* " & n & ")/7)"
            t = " case when ( (" & u & ")-(" & s & "))>0 then case when ((" & u & ")-(" & s & ")-(" & Q & "))>0 then  (" & u & ")-(" & s & ")-(" & Q & ") else 0 end else 0 end "  'proyecccion vtas
            'W = "((" & O & ")+(" & t & ")-(" & n & ")+(" & Q & "))"
            W = "((" & p & ")+(" & t & ")-(" & n & ")+(" & Q & "))"
            
            X = "CASE WHEN (" & n & ")>0 THEN ((" & W & ")/(" & n & ")*7) ELSE 0 END "
            Y = " case when  (ROUND((" & X & "),0)=ROUND((" & v & "),0)) And (Round(" & t & ",0))>0 then 'VERDADERO' else 'FALSO' end "
            
            sql = " select"
            sql = sql & " VS.IDINVENTARIO, vs.codinventario, "
            sql = sql & " vs.Descripcion, vs.CodSucursal, codgrupo, codgrp, "
            If Not .IncluirCero Then
                sql = sql & " TOTALM3, "
            Else
                sql = sql & " TOTALC, "
            End If
            sql = sql & " vs.cant,tf8.promc as cant8, "
'            If Not .IncluirCero Then
'                sql = sql & " case when vs.M3<>0 then " & L & " else 0 end as vtasemproyM3, "
'            Else
                sql = sql & " " & Lm3 & " as vtasemproyM3, "
'            End If
'            If Not .IncluirCero Then
'                sql = sql & " case when vs.M3<>0 then round((" & L & "),0) else 0 end as vtasemproyUni, "
'            Else
                'sql = sql & " round((" & L & "),6)  as vtasemproyUni, "
                sql = sql & L & " as vtasemproyUni, "
'            End If
            sql = sql & O & "  as existM3,"
            sql = sql & p & " as exist, "
            sql = sql & " tped.cantpedido as Pedido, "
            sql = sql & Q & " *m3 as transitoM3,"
            sql = sql & " isnull(transito,0) as transito,"
            sql = sql & r & " as diasInv, "
            sql = sql & "(" & s & ") *m3 as InvFnalM3,"
'            If Not .IncluirCero Then
'                sql = sql & " case when vs.M3<>0 then round((" & s & "/ vs.M3),0) else 0 end as InvFnalUni,"
'            Else
                sql = sql & " round((" & s & "),0) as InvFnalUni,"
'            End If
            
            sql = sql & t & "*m3 as compraidealM3,"
'            If Not .IncluirCero Then
'                sql = sql & " case when vs.M3<>0 then round((" & t & "/ vs.M3),0) else 0 end as compraidealUni,"
'            Else
                sql = sql & " round((" & t & "),0)  as compraidealUni,"
'            End If
'            If Not .IncluirCero Then
'                sql = sql & " case when vs.M3<>0 then round((" & t & "/ vs.M3),0) else 0 end as compraidealUr,"
'            Else
                sql = sql & " round((" & t & "),0) as compraidealUr,"
'            End If
            sql = sql & " 0 as Pedir,0 as Aprobar, 0 as stockpedir, "
            sql = sql & u & " *m3 as INVENTARIOIDEALM3, "
'            If Not .IncluirCero Then
'                sql = sql & " case when vs.M3<>0 then round((" & u & "/ vs.M3),0) else 0 end as INVENTARIOIDEALU,"
'            Else
                sql = sql & " round((" & u & "),0) as INVENTARIOIDEALU,"
'            End If
            sql = sql & v & " TiempoPromVta,"
            sql = sql & W & " *m3  as invluegocpM3, "
'
'            If Not .IncluirCero Then
'                sql = sql & " case when vs.M3<>0 then round((" & W & "/ vs.M3),0) else 0 end as invluegocpMuNI,"
'            Else
                sql = sql & " round((" & W & "),0) as invluegocpMuNI,"
'            End If
            sql = sql & X & "  as RealDias, "
            sql = sql & Y & " as PEDIRSTR  Into tmpfin "

            sql = sql & " from TMPvs vs"
            sql = sql & " left join TMPvsFin tf on vs.idinventario = tf.idinventario"
            sql = sql & " left join TMPvs8sem tf8 on vs.idinventario = tf8.idinventario"
            sql = sql & " left join VWivexistxSucursal vwe on vs.idinventario=vwe.idinventario "
            sql = sql & " left join TMPPedido tped on vs.idinventario=tped.idinventario "
            sql = sql & " left join TMPTransito tt on vs.idinventario=tt.idinventario "

            sql = sql & " where 1=1 "
            sql = sql & " and ( exist<>0 or TOTALM3<>0)"
            If Len(.Sucursal) > 0 Then
                sql = sql & " AND vwe.codsucursal IN (" & PreparaCadena(.Sucursal) & ")"
            End If
            sql = sql & "And vwe.idbodega in (" & .CodBodega1 & ")"
        
        If Not (Len(.Bienes)) = 0 Then
            bienesmod = Replace(.Bienes, "IVInventario.", "vwe.")
            bienesmod = Replace(bienesmod, "vwe.IdGrupo4", "vwe.Idgrupo")
'            bienesmod = Replace(bienesmod, "IVInventario. ", "vwe.")
            sql = sql & bienesmod ' .Bienes  'Aquí se ha grabado SQL de ítems
        End If
        sql = sql & " Group by VS.IDINVENTARIO, vs.codinventario,  vs.Descripcion, vs.CodSucursal,codgrupo,codgrp,  TOTALM3,  vs.cant,vs.M3,tf.prom,tf8.prom,cantpedido,transito,TiempoPromVta,tf8.promc,totalc,tf.promc "
        sql = sql & " ORDER BY vs.TOTALM3 DESC, vs.CodSucursal desc, exist desc"
        
        VerificaExistenciaTablaTemp "TMPFin"
        gobjMain.EmpresaActual.EjecutarSQL sql, numReg
        
        sql = "Select * from tmpfin "
        If Not .IncluirCero Then
            sql = sql & " union all "
            sql = sql & "SELECT "
            sql = sql & "ivinventario.IDINVENTARIO,ivinventario.codinventario,ivinventario.Descripcion,'" & .Sucursal & "' as  CodSucursal,'' as codgrupo,ivg.codgrupo4  as codgrp, "
            sql = sql & "0 as TOTALC,0 as cant,0 as cant8,0 as vtasemproyM3,0 as vtasemproyUni,(sum(ive.exist)*ivinventario.m3) as existM3,"
            sql = sql & " sum(ive.exist) as Exist,'' as Pedido,'' as transitoM3,'' as transito,0 as diasInv,0 as InvFnalM3,"
            sql = sql & " 0 as InvFnalUni,0 as compraidealM3, 0 as compraidealUni,0 as compraidealUr,0 as Pedir,"
            sql = sql & " 0 as Aprobar,10000 as stockpedir,0 as INVENTARIOIDEALM3,0 as INVENTARIOIDEALU"
            sql = sql & " ,ivinventario.TiempoPromVta,0 as invluegocpM3,0 as invluegocpMuNI,0 as RealDias,'FALSO' as PEDIRSTR"
            sql = sql & " FROM ivinventario   Inner Join ivexist ive "
            sql = sql & " Inner Join ivgrupo4 ivg on ivg.idgrupo4 = ive.idgrupo"
            sql = sql & " on ive.idinventario = ivinventario.idinventario"
            sql = sql & " Where ivinventario.idinventario not in (select idinventario from tmpfin)"
            sql = sql & "And ive.idbodega in (" & .CodBodega1 & ")"
                If Not (Len(.Bienes)) = 0 Then
                    sql = sql & .Bienes 'Aquí se ha grabado SQL de ítems
                End If
            sql = sql & " and ivinventario.bandservicio = 0"
            sql = sql & " group by ivinventario.IDINVENTARIO,ivinventario.codinventario,ivinventario.Descripcion,ivg.codgrupo4,ivinventario.m3,ivinventario.TiempoPromVta"
            End If
            sql = sql & "  ORDER BY CODGRP "
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
        Set CargaBufferDista = rs
        Set rs = Nothing
    End With
End Function

Public Function ListaTransxDevolucionDscto(CodTrans As String, CodForma As String, coddoc As String, fdesde As Date, fhasta As Date) As Variant
    Dim sql As String, rs As Recordset, numReg As Long
    sql = "Select  g.transid,g.codtrans,g.numtrans,pc.codprovcli,pc.nombre,iv.codinventario,iv.descripcion,abs(ivk.cantidad) as cantidad,abs(ivk.preciorealtotal)as Pt,ivk.descuentopp*100 ,'' as AplicaNC"
    sql = sql & " From gncomprobante g "
    sql = sql & " inner join ivkardex ivk"
    sql = sql & " inner join ivinventario iv"
    sql = sql & " on iv.idinventario=ivk.idinventario"
    sql = sql & " on g.transid = ivk.transid"
    sql = sql & " Inner Join pcprovcli pc on pc.idprovcli = g.idclienteref"
    sql = sql & " where g.estado <> 3 And pc.bandNCPP = 1"
    sql = sql & " And g.codtrans = '" & CodTrans & "'"
    sql = sql & " And g.fechatrans between '" & fdesde & "' And  '" & fhasta & "'"
    sql = sql & " And g.transid not in (select idtransfuente from gncomprobante where codtrans in ('" & coddoc & "') and estado <> 3)"
    If Len(CodForma) > 0 Then
        sql = sql & " And pc.codprovcli =  '" & CodForma & "'"
    End If
   Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    Set ListaTransxDevolucionDscto = rs
    Set rs = Nothing
End Function

Public Function ObtieneDesctoPCxItemxFechaAcumulaNew(ByVal CodInventario As String, ByVal codPC As String, ByVal codFC As String, ByVal IdDescuento As Long, ByRef BandAcumula As Boolean, ByRef FechaTrans As Date, ByRef HoraTrans As Date) As Double
    Dim sql As String, rs As Recordset, NumIVGrupo As Integer, NumPCGrupo As Integer
    Dim CodDescuento As String
    'CodDescuento = ""
    NumIVGrupo = CInt(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("DescuentoNumIVGrupo")) + 1
    NumPCGrupo = CInt(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("DescuentoNumPCGrupo")) + 1
    sql = " select top 1"
    sql = sql & " ivD.codDescuento as codDesc, ivD.BandAcumula,"
    sql = sql & " case when ivdiv.Descuento-Penal>0 then ivdiv.Descuento-Penal else 0 end as TotalDesc,"
    sql = sql & " ivd.IdDescuento , ivd.BandCaducidad, ivd.FechaDesde, ivd.FechaHasta, ivd.HoraDesde, ivd.HoraHasta "
    sql = sql & " From ivdescuento ivd "
    sql = sql & " Left Join ivdescuentodetallepc ivdpc   "
    sql = sql & " Left Join pcgrupo" & NumPCGrupo & " pcg "
    sql = sql & " Left Join pcprovcli pc on pc.idgrupo" & NumPCGrupo & "  = pcg.idgrupo" & NumPCGrupo & " "
    sql = sql & " on pcg.idGrupo" & NumPCGrupo & " =  ivdpc.idpcgrupo"
    sql = sql & " Left Join pcprovcli pcc on pcc.idprovcli = ivdpc.idprovcli"
    sql = sql & "  on ivdpc.iddescuento = ivd.iddescuento  "
    sql = sql & " Left Join  ivdescuentodetalleiv ivdiv"
    sql = sql & "  Left Join ivgrupo" & NumIVGrupo & " ivg "
    sql = sql & "  Left Join ivinventario iv on iv.idgrupo" & NumIVGrupo & " = ivg.idgrupo" & NumIVGrupo & ""
    sql = sql & "  on ivg.idgrupo" & NumIVGrupo & " =  ivdiv.idivgrupo  "
    sql = sql & "  Left Join ivinventario ivv on ivv.idinventario =  ivdiv.idinventario"
    sql = sql & "  on  ivdiv.iddescuento = ivd.iddescuento"
    sql = sql & "  Left Join ivdescuentodetallefc ivdfc "
    sql = sql & "  Left JOIN tsformacobropago tsf on tsf.idforma = ivdfc.idforma"
    sql = sql & "  on  ivdfc.iddescuento = ivd.iddescuento  "
    sql = sql & " WHERE ivd.idDescuento = '" & IdDescuento & "' "
    sql = sql & " And ( pc.CodProvCli ='" & codPC & "' or pcc.CodProvCli ='" & codPC & "')"
    sql = sql & " And ( iv.codinventario ='" & CodInventario & "' or ivv.CodInventario ='" & CodInventario & "')"
    If Len(codFC) > 0 Then
        sql = sql & " and tsf.codforma='" & codFC & "'"
    End If
    sql = sql & " and ivdiv.cantmin= 0"
    sql = sql & " order by ivd.fechagrabado"
   
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not rs.EOF Then
        If rs.Fields("BandCaducidad") Then
            If FechaTrans >= rs.Fields("FechaDesde") And FechaTrans <= rs.Fields("FechaHasta") Then
                ObtieneDesctoPCxItemxFechaAcumulaNew = rs!TotalDesc
                CodDescuento = rs!codDesc
                BandAcumula = rs!BandAcumula
            Else
                ObtieneDesctoPCxItemxFechaAcumulaNew = 0
                CodDescuento = 0
                BandAcumula = False
            End If
        Else
            ObtieneDesctoPCxItemxFechaAcumulaNew = rs!TotalDesc
            CodDescuento = rs!codDesc
            BandAcumula = rs!BandAcumula
        End If
   End If
    Set rs = Nothing
End Function

Public Function CargaTransaccionesParaGenerarPDF(ByVal cad As String) As Variant
    Dim sql As String, rs As Recordset
    With gobjCond
        sql = "select top 10 (CodTrans+' '+CONVERT(VARCHAR,NumTrans)) AS Trans, FechaTrans,  g.transid,0 as pdf, '' as nombre , pc.idtipodocumento"
        sql = sql & " from GNOferta gof inner join gncomprobante g on g.transid = gof.transid "
        sql = sql & " left join  pcprovcli pc on g.idclienteref = pc.idprovcli or g.idproveedorref = pc.idprovcli"
        sql = sql & " where  BandEnvio =0 AND g.estado <> 3  And g.codtrans in (" & PreparaCadena(cad) & ")"
        sql = sql & " order by g.transid  "
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
        Set CargaTransaccionesParaGenerarPDF = rs
        Set rs = Nothing
    End With
End Function
Public Function ActualizaBanderaGeneraPDF(transid As Long) As Boolean
    Dim sql As String, cond As String, CadenaPrecio As String, Trans As String, rs As Recordset
    Dim Dias As Integer, mes As Currency, numReg As Long
    With gobjCond
        sql = "update Gnoferta set Bandenvio =1 where transid = " & transid
        gobjMain.EmpresaActual.EjecutarSQL sql, numReg
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
        ActualizaBanderaGeneraPDF = True
        Set rs = Nothing
    End With
End Function

Public Function ActualizaBanderaBandEnvioTrans(transid As Long) As Boolean
    Dim sql As String, rs As Recordset
    Dim numReg As Long
    With gobjCond
        sql = "update GnOferta set BandEnvio=1 where transid = " & transid
        gobjMain.EmpresaActual.EjecutarSQL sql, numReg
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
        ActualizaBanderaBandEnvioTrans = True
        Set rs = Nothing
    End With
End Function

Public Function CargaTransaccionesParaEnvioEmail() As Variant
    Dim sql As String, rs As Recordset
    With gobjCond
        sql = "select Codtrans"
        sql = sql & " from GNTrans "
        sql = sql & " where  Bandvalida=1 "
        sql = sql & "  And substring(Opcion2, 212,1)= 'S'"
        sql = sql & " order by codtrans "
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
        Set CargaTransaccionesParaEnvioEmail = rs
        Set rs = Nothing
    End With
End Function

Public Function ListaFCVendedorXComision(BandRS As String, CodComision As String) As Variant
    Dim sql As String, rs As Recordset
    sql = "SELECT Codvendedor,fcv.Nombre FROM fcvendedor fcv  "
    sql = sql & "Inner join FCVTablaComision tc on tc.idtablacomision = fcv.idtablacomision"
    sql = sql & " Where tc.codtablacomision = '" & CodComision & "'  "
    sql = sql & " ORDER BY CodVendedor"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If BandRS Then
        Set ListaFCVendedorXComision = rs
    Else
        ListaFCVendedorXComision = MiGetRows(rs)
        rs.Close
    End If
    Set rs = Nothing
End Function

Public Function ListaTransxDevolucionDsctoNew(CodTrans As String, CodCli As String, coddoc As String, fdesde As Date, fhasta As Date, CodForma As String) As Variant
    Dim sql As String, rs As Recordset, numReg As Long
    'verifico las formas de cobro
    VerificaExistenciaTabla 199
    sql = "Select g.transid Into tmp199 from gncomprobante g Inner Join pckardex pck on pck.transid = g.transid"
    sql = sql & " Inner join pcprovcli pc on pc.idprovcli = g.idclienteref"
    sql = sql & " where g.estado <> 3 And pc.bandNCPP = 1"
    sql = sql & " And g.codtrans = '" & CodTrans & "'"
    sql = sql & " And g.fechatrans between '" & fdesde & "' And  '" & fhasta & "'"
    sql = sql & " And g.transid not in (select idtransfuente from gncomprobante where codtrans in ('" & coddoc & "') and estado <> 3)"
    If Len(CodCli) > 0 Then
        sql = sql & " And pc.codprovcli =  '" & CodCli & "'"
    End If
    gobjMain.EmpresaActual.EjecutarSQL sql, numReg
    
    'VERIFICO SI LA FACTURA YA FUE PAGADA
    sql = "Select top 1 g.transid, pckori.fechavenci ,pck.fechaemision as fechacobro Into tmp200 From pckardex pck Inner join pckardex pckori "
    sql = sql & " Inner join gncomprobante g Inner join pcprovcli pc on pc.idprovcli = g.idclienteref on g.transid = pckori.transid "
    sql = sql & " on pckori.id = pck.idasignado "
    sql = sql & " where g.estado <> 3 And pc.bandNCPP = 1"
    sql = sql & " And g.codtrans = '" & CodTrans & "'"
    sql = sql & " And g.fechatrans between '" & fdesde & "' And  '" & fhasta & "'"
    If Len(CodCli) > 0 Then
        sql = sql & " And pc.codprovcli =  '" & CodCli & "'"
    End If
    sql = sql & " order by pck.fechaemision desc "
    VerificaExistenciaTabla 200
    gobjMain.EmpresaActual.EjecutarSQL sql, numReg
    
    sql = "Select  g.transid,g.codtrans,g.numtrans,pc.codprovcli,pc.nombre,iv.codinventario,iv.descripcion,abs(ivk.cantidad) as cantidad,abs(ivk.preciorealtotal)as Pt,ivk.descuentopp*100 "
    sql = sql & ",t2.fechavenci,t2.fechacobro,datediff(d,t2.fechavenci,t2.fechacobro) as diasCobro"
    sql = sql & ",'' as AplicaNC"
    sql = sql & " From gncomprobante g Inner join tmp199 t on t.transid =g.transid  Left Join tmp200 t2 on t2.transid = g.transid  "
    sql = sql & " inner join ivkardex ivk"
    sql = sql & " inner join ivinventario iv"
    sql = sql & " on iv.idinventario=ivk.idinventario"
    sql = sql & " on g.transid = ivk.transid"
    sql = sql & " Inner Join pcprovcli pc on pc.idprovcli = g.idclienteref"
    sql = sql & " where g.estado <> 3 And pc.bandNCPP = 1"
    sql = sql & " And g.codtrans = '" & CodTrans & "'"
    sql = sql & " And g.fechatrans between '" & fdesde & "' And  '" & fhasta & "'"
    sql = sql & " And g.transid not in (select idtransfuente from gncomprobante where codtrans in ('" & coddoc & "') and estado <> 3)"
    If Len(CodCli) > 0 Then
        sql = sql & " And pc.codprovcli =  '" & CodCli & "'"
    End If
   Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    Set ListaTransxDevolucionDsctoNew = rs
    Set rs = Nothing
End Function

