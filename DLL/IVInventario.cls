VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Empresa"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit
'Auxiliar para guardar datos de ultimo ingreso/egreso
Public Enum EIVTipo
    Normal = 0
    Receta = 1
    Familia = 2
    CambioPresentacion = 3
    Preparacion = 4
    Promocion = 5
    rubro = 6 'AUC 05/07/09 para construcciones antes 5
    Preparacionporce = 7
    Matriz = 8 'LIRA

End Enum

Public Enum SRITipo
    ACTIVO_CORRIENTE = 0
    SERVICIO_Y_GASTO = 1
    ACTIVOS_FIJO = 2
    REEMBOLSO_GASTO = 3
    SIN_DERECHO_CREDTRIB = 4
End Enum



Private Type auxUltimo
    fecha As Date
    Cantidad As Currency
    Costo As Currency
    Precio As Currency
    Recuperado As Boolean
    Comision As Currency
End Type

Private Type T_PROPS
    IdInventario As Long
    CodInventario As String
    CodAlterno1 As String
    CodAlterno2 As String
    Descripcion As String
    DescripcionDetalle As String        'memo
    CodMoneda As String
'    Precio(1 To 6) As Currency 'AUC cambiado antes 4
'    Comision(1 To 6) As Currency 'AUC cambiado antes 4
'    Descuento(1 To 6) As Currency  'AUC cambiado antes 4        '***Agregado  Oliver. 30/jul/2004
'    CantLimite(1 To 6) As Currency 'AUC cambiado antes 4
    Precio(1 To 7) As Currency 'GRV Precio 7 - 14MAY2014
    Comision(1 To 7) As Currency 'GRV Precio 7 - 14MAY2014
    Descuento(1 To 7) As Currency  'GRV Precio 7 - 14MAY2014
    CantLimite(1 To 7) As Currency 'GRV Precio 7 - 14MAY2014
    
    PorcentajeIVA As Currency
    PorcentajeIVAAnt As Currency
    IdCuentaActivo As Long
    IdCuentaCosto As Long
    IdCuentaVenta As Long
    CodCuentaActivo As String
    CodCuentaCosto As String
    CodCuentaVenta As String
    Unidad As String
    IdGrupo(1 To 6) As Long
    CodGrupo(1 To 6) As String
    idProveedor As Long
    CodProveedor As String
    Observacion As String
    ExistenciaMinima As Currency
    ExistenciaMaxima As Currency
    UnidadMinimaCompra As Currency
    UnidadMinimaVenta As Currency
    BandValida As Boolean
    BandServicio As Boolean
    Tipo   As EIVTipo   '  As Byte   'Diego Produccion 03/09/2002
                    'Tipo 0 = Normal
                    '     1 = Receta
                    '     2 = Familia
                    '     3 = CambioPresentacion
    ValorRecargo As Currency         '***Agregado. Angel. 29/jul/2004
    fechagrabado As Date
    BandFraccion As Boolean         '***Agregado. Esteban. 8/abr/2005
    BandArea As Boolean         '***Agregado. Esteban. 16/05/2005
    bandVenta As Boolean         '***Agregado. Esteban. 26/12/2005
    CodUnidad As String 'AUC agregado
    CodUnidadConteo As String 'AUC agregado
    IdUnidad As Long    'AUC 08/12/05 'para unidades
    IdUnidadConteo As Long 'AUC 08/12/05 'para unidades de conteo
    CostoUltimoIngreso As Currency 'AUC 08/12/05 'para unidades de conteo
    BandIVA As Boolean      'jeaa 08/03/2006 cambio para sri
    PorcentajeICE As Currency
    PorDesperdicio As Currency 'AUC 08/12/05 'para unidades de conteo
    objEmpresa As Empresa            'Referencia al objeto padre
    TipoSRI As SRITipo
    IdICE As Long    'jeaa 26/07/2006
    CodICE As String 'AUC agregado
    CantRelUnidad As Long
    CantRelUnidadCont As Long
    Descripcion2 As String      'jeaa 4/06/2008
    PesoNeto As Currency      'jeaa 4/06/2008
    PesoBruto As Currency      'jeaa 4/06/2008
    IdUnidadPeso As Long      'jeaa 4/06/2008
    CodUnidadPeso As String      'jeaa 4/06/2008
    BandConversion As Boolean      'jeaa 4/06/2008
    BandRepGastos As Boolean      'jeaa 18/08/2008
    RutaFoto As String
    BandNoSeFactura As Boolean      'jeaa 23/03/2009
    TiempoPromVta As Currency         'jeaa 24/04/2009
    TiempoReposicion As Currency     'jeaa 24/04/2009
    CostoReferencial As Currency ' JEAA 22/09/10 'para Costo en las compras
    BandModPrecio As Boolean      'jeaa 23/03/2009
    IdARANCEL As Long    'jeaa 26/07/2006
    CodArancel As String 'AUC agregado
    NumItems As Integer 'AUC abril 2012
    BandConversionUni As Boolean
    idproceso As Long
    CodProceso As String
    Utilidad As Currency
    Tipo1   As EIVTipo   '  As Byte
    BandAltaRotacion As Boolean
    CostoAdicionalUltimoIngreso As Currency
    BandNoObjetoIVA As Boolean
    BandGarantia As Boolean
    IdCuentaDiferida As Long
    CodCuentaDiferida As String
    TipoAuto As Integer 'AUC para tipo especial
    BandCambiaDescr As Boolean
    bandSerie As Boolean
    FrecuenciaReposicion As Currency
    Buffer As Currency
    FechaModBuffer As Date
    bandPeso As Boolean
    IdPlan As Long
    CodPlan As String
    FrecuenciaReposicionALM As Currency
    TiempoReposicionALM As Currency
    BandOmitirRebate As Boolean
    Batch As Currency 'batch unidad
    BatchProd As Currency 'batch unidad de produccion
    BandNoFormula As Boolean
    CantReluni As Currency
    CantReluniCont As Currency
    
    IdCuentaDevolucion As Long
    CodCuentaDevolucion As String
    
    CantReuni As Currency
    CantReuniCont As Currency
    M3 As Double
    BandGarantiaC As Boolean
    BandGarantiaV As Boolean
       
End Type


Private mProps As T_PROPS
Private mcolDetalles As Collection      'Coleccion de detalles de existencia
Private mcolDetalleFamilia As Collection      'Coleccion de detalles de familia
Private mcolDetalleProveedor As Collection  'Coleccion de detalles proveedor AUC 22/11/05
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean
Private mUIngreso As auxUltimo      'Ultimo ingreso
Private mUEgreso As auxUltimo       'Ultimo egreso
Private mcolDetalleConversion As Collection      'Coleccion de detalles de familia
Private mcolDetalleFamilia1 As Collection      'Coleccion de detalles de familia
Private mcolGarantia As Collection      'Coleccion de detalles de familia

Public Sub Grabar()
    Dim sql As String, i As Integer
    Dim rs As Recordset, rsDetalle As Recordset
    Dim rsSuc As Recordset, gnSuc As GNSucursal
    
    'Si no está modificado no hace nada
    
    If Not mbooModificado Then Exit Sub
        If Len(mProps.CodUnidad) = 0 Then
            If Len(mProps.objEmpresa.GNOpcion.ObtenerValor("UnidadPredeterminada")) > 0 Then
                mProps.CodUnidad = mProps.objEmpresa.GNOpcion.ObtenerValor("UnidadPredeterminada")
            End If
        End If
    'Verifica si no repite el código
    sql = "SELECT CodInventario FROM IVInventario " & _
          "WHERE (CodInventario='" & mProps.CodInventario & "') AND " & _
                "(IdInventario<>" & mProps.IdInventario & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "IVInventario", MSGERR_REPITECODIGO
    End If
    rs.Close
    'Verifica si no repite el código Alterno 1
    
    If Not mProps.objEmpresa.GNOpcion.ObtenerValor("CodAltDuplicado") = "1" Then
        If Len(mProps.CodAlterno1) > 0 Then
            sql = "SELECT CodAlterno1 FROM IVInventario " & _
                  "WHERE (CodAlterno1='" & mProps.CodAlterno1 & "') AND " & _
                        "(IdInventario<>" & mProps.IdInventario & ")"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If Not rs.EOF Then      'Si ya existe el codigo
                Err.Raise ERR_REPITECODIGOALT, "IVInventario", MSGERR_REPITECODIGOALT
            End If
            rs.Close
        End If
    End If
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM IVInventario WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM IVInventario WHERE IdInventario=" & mProps.IdInventario
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
'        If Len(rs!FechaGrabado) Then
            If rs!fechagrabado <> mProps.fechagrabado Then
'        End If
            Err.Raise ERR_MODIFICADO, "IVInventario_Grabar", MSGERR_MODIFICADO
        End If
    End If
    
    'Comienza la transaccion ----------------------------jeaa 16-03-04
#If DAOLIB Then
    gobjWS.BeginTrans
#Else
    mProps.objEmpresa.Coneccion.BeginTrans
#End If
    On Error GoTo ErrTrap
    'Graba el registro
    With rs
        !CodInventario = mProps.CodInventario
        !CodAlterno1 = mProps.CodAlterno1
        !CodAlterno2 = mProps.CodAlterno2
        !Descripcion = mProps.Descripcion
#If DAOLIB Then
        !DescripcionDetalle = mProps.DescripcionDetalle
#Else
        .Fields("DescripcionDetalle").AppendChunk mProps.DescripcionDetalle
#End If
        !CodMoneda = mProps.CodMoneda
        'For i = 1 To 6 'AUC cambiado antes 4
        For i = 1 To 7 'GRV Precio 7 - 14MAY2014
            .Fields("Precio" & i) = mProps.Precio(i)
            .Fields("Comision" & i) = mProps.Comision(i)
            .Fields("CantLimite" & i) = mProps.CantLimite(i)
            .Fields("Descuento" & i) = mProps.Descuento(i)    '***Agregado. jeaa. 30/jul/2004
        Next i
        !PorcentajeIVA = mProps.PorcentajeIVA
        !PorcentajeIVAAnt = mProps.PorcentajeIVAAnt
        !IdCuentaActivo = mProps.IdCuentaActivo
        !IdCuentaCosto = mProps.IdCuentaCosto
        !IdCuentaVenta = mProps.IdCuentaVenta
        !IdCuentaDiferida = mProps.IdCuentaDiferida
        !IdCuentaDevolucion = mProps.IdCuentaDevolucion
        '!Unidad = mProps.Unidad
        
        For i = 1 To IVGRUPO_MAX
            .Fields("IdGrupo" & i) = mProps.IdGrupo(i)
        Next i
        !idProveedor = mProps.idProveedor
        !Observacion = mProps.Observacion
        !ExistenciaMinima = mProps.ExistenciaMinima
        !ExistenciaMaxima = mProps.ExistenciaMaxima
        !UnidadMinimaCompra = mProps.UnidadMinimaCompra
        !UnidadMinimaVenta = mProps.UnidadMinimaVenta
        !BandValida = mProps.BandValida
        !BandServicio = mProps.BandServicio
        !Tipo = mProps.Tipo         'Diego Prod
        'Guarda el detalle de Items en tabla IVMateria
''        If mProps.Tipo <> Normal Then
''            GrabaFamiliaDetalle
''            'GrabaFamiliaDetalleNew 0
''        End If
        !ValorRecargo = mProps.ValorRecargo  '***Agregado. Angel. 29/jul/2004
        !BandFraccion = mProps.BandFraccion  '***Agregado. Esteban 08/04/2005
        !BandArea = mProps.BandArea  '***Agregado. Esteban 15/09/2005
        !bandVenta = mProps.bandVenta '***Agregado. Esteban 26/12/2005
        !IdUnidad = mProps.IdUnidad 'auc
        !IdUnidadConteo = mProps.IdUnidadConteo 'auc
        !BandIVA = mProps.BandIVA '***Agregado. Esteban 08/03/2006
        !PorcentajeICE = mProps.PorcentajeICE
        !PorDesperdicio = mProps.PorDesperdicio  ' *********** jeaa 01/02/2006
        !TipoSRI = mProps.TipoSRI         'jeaa 19/06/2006
        !IdICE = mProps.IdICE         'jeaa 26/07/2006
        !CantRelUnidad = mProps.CantRelUnidad         'jeaa 27/09/2007
        !CantRelUnidadCont = mProps.CantRelUnidadCont         'jeaa 27/07/2007
        
        !CantReluni = mProps.CantReluni         'jeaa 27/09/2007
        !CantReluniCont = mProps.CantReluniCont         'jeaa 27/07/2007
        
        !CantReuni = mProps.CantReuni         'jeaa 27/09/2007
        !CantReuniCont = mProps.CantReuniCont         'jeaa 27/07/2007
        !Descripcion2 = mProps.Descripcion2             'jeaa 4/06/2008
        !PesoNeto = mProps.PesoNeto         'jeaa 04/06/2008
        !PesoBruto = mProps.PesoBruto         'jeaa 04/06/2008
        !IdUnidadPeso = mProps.IdUnidadPeso 'jeaa 04/06/2008

        !BandConversion = mProps.BandConversion 'jeaa 04/06/2008
        !BandRepGastos = mProps.BandRepGastos 'jeaa 18/08/2008
        !RutaFoto = mProps.RutaFoto 'AUC 04/05/06
        !BandNoSeFactura = mProps.BandNoSeFactura 'jeaa 23/03/2009
        !TiempoPromVta = mProps.TiempoPromVta  '***Agregado. jeaa 24/04/2009
        !TiempoReposicion = mProps.TiempoReposicion '***Agregado. jeaa. 24/04/2009
        !CostoReferencial = mProps.CostoReferencial         'jeaa 22/09/2010
        !BandModPrecio = mProps.BandModPrecio 'jeaa 30/08/2011
        !IdARANCEL = mProps.IdARANCEL         'jeaa 26/07/2006
        !BandConversionUni = mProps.BandConversionUni
        !Tipo1 = mProps.Tipo1
        If Tipo1 <> Tipo And Tipo1 <> 0 Then
            If mProps.Tipo1 <> Normal Then
'                GrabaFamiliaDetalleNew 1
            End If
        End If
        !Utilidad = mProps.Utilidad          'AUC
        !idproceso = mProps.idproceso           'AUC
        !BandAltaRotacion = mProps.BandAltaRotacion
        !BandNoObjetoIVA = mProps.BandNoObjetoIVA
        !BandGarantia = mProps.BandGarantia
        !TipoAuto = mProps.TipoAuto 'AUC para tipo especial caso autos
        !BandCambiaDescr = mProps.BandCambiaDescr
        !bandSerie = mProps.bandSerie
        !FrecuenciaReposicion = mProps.FrecuenciaReposicion '***Agregado. jeaa. 26/01/2015
        !Buffer = mProps.Buffer
        !FechaModBuffer = mProps.FechaModBuffer
        !bandPeso = mProps.bandPeso
        !FrecuenciaReposicionALM = mProps.FrecuenciaReposicionALM
        !TiempoReposicionALM = mProps.TiempoReposicionALM
        !BandOmitirRebate = mProps.BandOmitirRebate
        !Batch = mProps.Batch
        !BatchProd = mProps.BatchProd
        !BandNoFormula = mProps.BandNoFormula
        !M3 = mProps.M3
        !BandGarantiaC = mProps.BandGarantiaC
        !BandGarantiaV = mProps.BandGarantiaV
        
        !fechagrabado = Now
        .Update
        mProps.IdInventario = !IdInventario
        If mProps.Tipo <> Normal Then
            GrabaFamiliaDetalle
            'GrabaFamiliaDetalleNew 0
        End If
                
        GrabarProveedorDetalle (mProps.CodInventario) 'AUC 24/11/05
        If mProps.BandConversionUni Then
            GrabaConversionDetalle
        End If
        GrabarIVGarantia
         'Guarda el detalle de Items en tabla IVMateria
'    AsignarBodegaPre
        
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        rs.Move 0           'Para actualizar
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_ITEM", "Item Nuevo: " & _
                                                         mProps.CodInventario, "DT"
        'jeaa 16-03-04 para grabar en ivexistencia la bodega y cantidad solo cuando es nuevo
        AseguraMovimientoExistencia
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_ITEM", "Item Modificado: " & _
                                                         mProps.CodInventario, "DT"
    End If
    Set rs = Nothing
    
    'actualiza Inventario en Bases de Datos Sucursales
    sql = "SELECT * FROM GnSucursal WHERE bandvalida=1"
    Set rsSuc = mProps.objEmpresa.OpenRecordset(sql)
    
    For i = 1 To rsSuc.RecordCount
        Set gnSuc = mProps.objEmpresa.RecuperaGNSucursal(rsSuc.Fields("CodSucursal"))
        If Not gnSuc Is Nothing Then
            If gnSuc.BandInventario Then
                GrabarEnSucursales gnSuc.BaseDatos, gnSuc.servidor
            End If
        End If
        rsSuc.MoveNext
    Next i
    Set gnSuc = Nothing
    Set rsSuc = Nothing
    
    mbooEsNuevo = False
    mbooModificado = False
    
    'Completa la transacción ---------------------------- jeaa 16-03-04
#If DAOLIB Then
    gobjWS.CommitTrans
#Else
    mProps.objEmpresa.Coneccion.CommitTrans
#End If
   mProps.objEmpresa.RaiseRefreshEvent REFRESH_INVENTARIO
    Exit Sub
ErrTrap:
    'Anula la transacción ---------------------------- jeaa 16-03-04
#If DAOLIB Then
    gobjWS.Rollback
#Else
    mProps.objEmpresa.Coneccion.RollbackTrans
#End If

''    GrabarEnSucursal "guala2008"

    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_INVENTARIO
End Sub

Public Sub GrabaFamiliaDetalle()
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As IVFamiliaDetalle, sql As String, band As Boolean
    Dim NumHijos As Integer, CountNumHijos As Integer
    On Error GoTo CapturaError
    
    'Elimina  todos  los registros
    sql = "Delete Ivm " & _
          "From Ivinventario Iv " & _
          "INNER JOIN IVMateria Ivm ON IV.IdInventario = Ivm.IdMateria "
    'jeaa 28/10/2005
    If mProps.Tipo = CambioPresentacion _
        Or mProps.Tipo = Preparacion _
        Or Tipo = Promocion _
        Or mProps.Tipo = rubro _
        Or Tipo = Preparacionporce Then
          sql = sql & " Where Ivm.IdInventario = " & mProps.IdInventario
   Else
          sql = sql & " Where Ivm.Idmateria = " & mProps.IdInventario
   End If

    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    band = True
    If Not mcolDetalleFamilia Is Nothing Then
        sql = "SELECT * FROM IVMateria WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        NumHijos = mcolDetalleFamilia.Count
        CountNumHijos = 0
        For i = 1 To mcolDetalleFamilia.Count
'            rs1.AddNew /jeaa 28-01-04
            If CountNumHijos < NumHijos Then
                CountNumHijos = CountNumHijos + 1
                Set obj = mcolDetalleFamilia.item(i)
                If obj.IdInventario = 0 Then
                        obj.Recuperar obj.CodInventario
                        If verificacionDuiplicado(obj.IdInventario) Then
                            RemoveDetalleFamilia i
                            i = i - 1
    '                        band = False
                        End If
                End If
                If band Then
                    rs1.AddNew
                    If Me.Tipo <> Familia Then
                        rs1!IdInventario = Me.IdInventario
                        rs1!idMateria = obj.IdInventario
                    Else
                        rs1!IdInventario = obj.IdInventario
                        rs1!idMateria = Me.IdInventario
                    End If
                    rs1!TarifaJornal = obj.TarifaJornal
                    rs1!Rendimiento = obj.Rendimiento
          
                    rs1!Cantidad = obj.Cantidad
                    rs1!BandPrincipal = obj.BandPrincipal
                    rs1!BandModificar = obj.BandModificar 'AUC 25/05/06
                    rs1!Orden = obj.Orden
                    rs1!idproceso = obj.idproceso
    
                    If Me.Tipo = Promocion Then
                        rs1!xCuanto = obj.xCuanto 'AUC 26/09/007
                        rs1!TarifaJornal = obj.Precio  'AUC 26/09/007
                    End If
                    rs1.Update
'                    If Me.Tipo = rubro Then 'graba el costoultimoIngreso
'                        sql = " Update ivinventario  " & _
'                              " Set costoultimoingreso = " & obj.TarifaJornal & _
'                              " Where codinventario = '" & obj.CodInventario & "'"
'                                mProps.objEmpresa.Execute sql, True
'                    End If
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
'AUC 22/11/05
Private Sub GrabarProveedorDetalle(ByRef cod As String)
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj  As IVDetalleProveedor, sql As String, band As Boolean
    'Elimina  todos  los registros
    sql = "Delete From ivproveedordetalle where idinventario =" & RecuperaID(cod)
    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    band = True
    If Not mcolDetalleProveedor Is Nothing Then
        sql = "SELECT * FROM  ivproveedordetalle WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        For i = 1 To mcolDetalleProveedor.Count
'            rs1.AddNew /jeaa 28-01-04
            Set obj = mcolDetalleProveedor.item(i)
            If obj.idProveedor = 0 Then
                        obj.Recuperar obj.CodProveedor
                        If verificacionDuiplicado(obj.idProveedor) Then
                            RemoveDetalleProveedor i
                            band = False
                        End If
            End If
            If band Then
            
                rs1.AddNew
                    rs1!IdInventario = RecuperaID(cod)
                    rs1!idProveedor = obj.idProveedor
                    rs1!Cantidad = obj.Cantidad
                    rs1!CostoUnitario = obj.CostoUnitario
                    rs1!fecha = obj.fecha
                    rs1!BandPrincipal = obj.BandPrincipal
                    rs1!PrecioCompra = obj.PrecioCompra
                    rs1!idInventarioProv = obj.idInventarioProv
                    rs1!CodInventarioProv = obj.CodInventarioProv
                rs1.Update
            End If
            band = True
        Next i
    End If
    Set rs1 = Nothing
    Set obj = Nothing
End Sub
'AUC 25/11/05
Public Function RecuperaID(ByRef cod As String) As Long
    Dim sql As String, rs As Recordset

        sql = "SELECT idinventario FROM ivinventario WHERE Codinventario='" & cod & "'"
    Set rs = Me.Empresa.OpenRecordset(sql)
    With rs
     If rs.RecordCount = 0 Then
        RecuperaID = 0
      Else
        RecuperaID = !IdInventario
        Exit Function
        .Close
      End If
    End With
    Set rs = Nothing
End Function

Friend Sub Recuperar(key As Variant)
    Dim sql As String, i As Integer
    Dim rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM vwIVInventarioRecuperar WHERE CodInventario='" & key & "'"
    Else
        sql = "SELECT * FROM vwIVInventarioRecuperar WHERE IdInventario=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IdInventario = !IdInventario
        mProps.CodInventario = !CodInventario
        If Not IsNull(!CodAlterno1) Then mProps.CodAlterno1 = !CodAlterno1
        If Not IsNull(!CodAlterno2) Then mProps.CodAlterno2 = !CodAlterno2
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
#If DAOLIB Then
        If Not IsNull(!DescripcionDetalle) Then mProps.DescripcionDetalle = !DescripcionDetalle
#Else
        If Not IsNull(!DescripcionDetalle) Then
            If .Fields("DescripcionDetalle").ActualSize > 0 Then
                mProps.DescripcionDetalle = .Fields("DescripcionDetalle").GetChunk(.Fields("DescripcionDetalle").ActualSize)
            End If
        End If
#End If
        If Not IsNull(!CodMoneda) Then mProps.CodMoneda = !CodMoneda
        'For i = 1 To 6 'AUC cambiado antes 4
        For i = 1 To 7 'GRV Precio 7 - 14MAY2014
            If Not IsNull(.Fields("Precio" & i)) Then mProps.Precio(i) = .Fields("Precio" & i)
            If Not IsNull(.Fields("Comision" & i)) Then mProps.Comision(i) = .Fields("Comision" & i)
            If Not IsNull(.Fields("CantLimite" & i)) Then mProps.CantLimite(i) = .Fields("CantLimite" & i)
            If Not IsNull(.Fields("Descuento" & i)) Then mProps.Descuento(i) = .Fields("Descuento" & i)   '***Agregado. jeaa. 29/jul/2004
        Next i
        If Not IsNull(!PorcentajeIVA) Then mProps.PorcentajeIVA = !PorcentajeIVA
        If Not IsNull(!PorcentajeIVAAnt) Then mProps.PorcentajeIVAAnt = !PorcentajeIVAAnt
        If Not IsNull(!IdCuentaActivo) Then mProps.IdCuentaActivo = !IdCuentaActivo
        If Not IsNull(!IdCuentaCosto) Then mProps.IdCuentaCosto = !IdCuentaCosto
        If Not IsNull(!IdCuentaVenta) Then mProps.IdCuentaVenta = !IdCuentaVenta
        If Not IsNull(!IdCuentaDiferida) Then mProps.IdCuentaDiferida = !IdCuentaDiferida
        If Not IsNull(!IdCuentaDevolucion) Then mProps.IdCuentaDevolucion = !IdCuentaDevolucion
        
        If Not IsNull(!CodCuentaActivo) Then mProps.CodCuentaActivo = !CodCuentaActivo
        If Not IsNull(!CodCuentaCosto) Then mProps.CodCuentaCosto = !CodCuentaCosto
        If Not IsNull(!CodCuentaVenta) Then mProps.CodCuentaVenta = !CodCuentaVenta
        If Not IsNull(!CodCuentaDiferida) Then mProps.CodCuentaDiferida = !CodCuentaDiferida
        If Not IsNull(!CodCuentaDevolucion) Then mProps.CodCuentaDevolucion = !CodCuentaDevolucion
        
'        If Not IsNull(!Unidad) Then mProps.Unidad = !Unidad
        
        For i = 1 To IVGRUPO_MAX
            If Not IsNull(.Fields("IdGrupo" & i)) Then mProps.IdGrupo(i) = .Fields("IdGrupo" & i)
            If Not IsNull(.Fields("CodGrupo" & i)) Then mProps.CodGrupo(i) = .Fields("CodGrupo" & i)
        Next i
        
        If Not IsNull(!idProveedor) Then mProps.idProveedor = !idProveedor

        CargaDetalleProveedor  'auc
        If Not IsNull(!CodProvcli) Then mProps.CodProveedor = !CodProvcli
        If Not IsNull(!Observacion) Then mProps.Observacion = !Observacion
        If Not IsNull(!ExistenciaMinima) Then mProps.ExistenciaMinima = !ExistenciaMinima
        If Not IsNull(!ExistenciaMaxima) Then mProps.ExistenciaMaxima = !ExistenciaMaxima
        If Not IsNull(!UnidadMinimaCompra) Then mProps.UnidadMinimaCompra = !UnidadMinimaCompra
        If Not IsNull(!UnidadMinimaVenta) Then mProps.UnidadMinimaVenta = !UnidadMinimaVenta
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!BandServicio) Then mProps.BandServicio = !BandServicio
        If Not IsNull(!Tipo) Then mProps.Tipo = !Tipo
        'DiegoProd  Ojo  aumentar  proceso  para  recuperar  detalle
        If Tipo <> Normal Then
            CargaDetalleFamilia 'Pensandolo bien recuperamos
        End If
        '***Agregado. Angel. 29/jul/2004
        If Not IsNull(!ValorRecargo) Then mProps.ValorRecargo = !ValorRecargo
        '*** Agregado esteban 08/04/2005
        If Not IsNull(!BandFraccion) Then mProps.BandFraccion = !BandFraccion
        '*** Agregado esteban 15/09/2005
        If Not IsNull(!BandArea) Then mProps.BandArea = !BandArea
        '*** Agregado esteban 26/12/2005
        If Not IsNull(!bandVenta) Then mProps.bandVenta = !bandVenta
         If Not IsNull(!CodUnidad) Then mProps.CodUnidad = !CodUnidad
        If Not IsNull(!CodUniConteo) Then mProps.CodUnidadConteo = !CodUniConteo
        If Not IsNull(!IdUnidad) Then mProps.IdUnidad = !IdUnidad
        If Not IsNull(!IdUnidadConteo) Then mProps.IdUnidadConteo = !IdUnidadConteo
        If Not IsNull(!CostoUltimoIngreso) Then mProps.CostoUltimoIngreso = !CostoUltimoIngreso 'AUC 22/12/05
        If Not IsNull(!BandIVA) Then mProps.BandIVA = !BandIVA  'jeaa 08/03/2006
        If Not IsNull(!PorcentajeICE) Then mProps.PorcentajeICE = !PorcentajeICE
        If Not IsNull(!PorDesperdicio) Then mProps.PorDesperdicio = !PorDesperdicio 'jeaa 01/02/2006
        If Not IsNull(!RutaFoto) Then mProps.RutaFoto = !RutaFoto  ' AUC 04/05/06
        If Not IsNull(!TipoSRI) Then mProps.TipoSRI = !TipoSRI
        If Not IsNull(!IdICE) Then mProps.IdICE = !IdICE
        If Not IsNull(!CodICE) Then mProps.CodICE = !CodICE
        If Not IsNull(!CantRelUnidad) Then mProps.CantRelUnidad = !CantRelUnidad 'jeaa 27/0/2007
        If Not IsNull(!CantRelUnidadCont) Then mProps.CantRelUnidadCont = !CantRelUnidadCont 'jeaa 27/0/2007
        
        If Not IsNull(!CantReluni) Then mProps.CantReluni = !CantReluni 'jeaa 27/0/2007
        If Not IsNull(!CantReluniCont) Then mProps.CantReluniCont = !CantReluniCont 'jeaa 27/0/2007
        
        If Not IsNull(!CantReuni) Then mProps.CantReuni = !CantReuni 'jeaa 27/0/2007
        If Not IsNull(!CantReuniCont) Then mProps.CantReuniCont = !CantReuniCont 'jeaa 27/0/2007
        
        If Not IsNull(!Descripcion2) Then mProps.Descripcion2 = !Descripcion2
        If Not IsNull(!PesoNeto) Then mProps.PesoNeto = !PesoNeto
        If Not IsNull(!PesoBruto) Then mProps.PesoBruto = !PesoBruto
        If Not IsNull(!IdUnidadPeso) Then mProps.IdUnidadPeso = !IdUnidadPeso
        If Not IsNull(!CodUniPeso) Then mProps.CodUnidadPeso = !CodUniPeso
        If Not IsNull(!BandConversion) Then mProps.BandConversion = !BandConversion
        If Not IsNull(!BandRepGastos) Then mProps.BandRepGastos = !BandRepGastos
        If Not IsNull(!BandNoSeFactura) Then mProps.BandNoSeFactura = !BandNoSeFactura
        If Not IsNull(!TiempoPromVta) Then mProps.TiempoPromVta = !TiempoPromVta
        If Not IsNull(!TiempoReposicion) Then mProps.TiempoReposicion = !TiempoReposicion
        If Not IsNull(!CostoReferencial) Then mProps.CostoReferencial = !CostoReferencial 'AUC 22/12/05
        If Not IsNull(!BandModPrecio) Then mProps.BandModPrecio = !BandModPrecio
        If Not IsNull(!IdARANCEL) Then mProps.IdARANCEL = !IdARANCEL
        If Not IsNull(!CodArancel) Then mProps.CodArancel = !CodArancel
        If Not IsNull(!Tipo1) Then mProps.Tipo1 = !Tipo1
'        If Tipo1 <> Normal Then
'            CargaDetalleFamilia1
'        End If

        If Not IsNull(!Utilidad) Then mProps.Utilidad = !Utilidad 'Auc REVISAR PARA QUE ES
        If Not IsNull(!idproceso) Then mProps.idproceso = !idproceso
        If Not IsNull(!CodProceso) Then mProps.CodProceso = !CodProceso
        If Not IsNull(!BandConversionUni) Then mProps.BandConversionUni = !BandConversionUni
        If Not IsNull(!BandAltaRotacion) Then mProps.BandAltaRotacion = !BandAltaRotacion
        If Not IsNull(!BandNoObjetoIVA) Then mProps.BandNoObjetoIVA = !BandNoObjetoIVA
        If Not IsNull(!BandGarantia) Then mProps.BandGarantia = !BandGarantia
        If Not IsNull(!BandCambiaDescr) Then mProps.BandCambiaDescr = !BandCambiaDescr
        If Not IsNull(!bandSerie) Then mProps.bandSerie = !bandSerie
        If Not IsNull(!FrecuenciaReposicion) Then mProps.FrecuenciaReposicion = !FrecuenciaReposicion
        If Not IsNull(!Buffer) Then mProps.Buffer = !Buffer
        If Not IsNull(!FechaModBuffer) Then mProps.FechaModBuffer = !FechaModBuffer
        If Not IsNull(!bandPeso) Then mProps.bandPeso = !bandPeso
        If Not IsNull(!FrecuenciaReposicionALM) Then mProps.FrecuenciaReposicionALM = !FrecuenciaReposicionALM
        If Not IsNull(!TiempoReposicionALM) Then mProps.TiempoReposicionALM = !TiempoReposicionALM
        If Not IsNull(!BandOmitirRebate) Then mProps.BandOmitirRebate = !BandOmitirRebate
        If Not IsNull(!Batch) Then mProps.Batch = !Batch
        If Not IsNull(!BatchProd) Then mProps.BatchProd = !BatchProd
        If Not IsNull(!BandNoFormula) Then mProps.BandNoFormula = !BandNoFormula
        If Not IsNull(!M3) Then mProps.M3 = !M3
        If Not IsNull(!BandGarantiaC) Then mProps.BandGarantiaC = !BandGarantiaC
        If Not IsNull(!BandGarantiaV) Then mProps.BandGarantiaV = !BandGarantiaV

                
        mProps.fechagrabado = !fechagrabado
            If mProps.BandConversionUni Then
                CargaDetalleConversion
            End If
            CargaGarantia
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Function AddDetalle() As Long
    Dim nuevo As IVInventarioDetalle
    
    Set nuevo = New IVInventarioDetalle
    Set nuevo.IVInventario = Me
    mcolDetalles.Add nuevo
    AddDetalle = mcolDetalles.Count    'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
End Function


Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    
    'Verifica si ya existe un Kardex usando el mismo IdInventario
    sql = "SELECT Id FROM IVKardex WHERE IdInventario=" & mProps.IdInventario
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "IVInventario_Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close
    'Controlar  o  eliminar  todos los registros
    'de  la tabla  IVMateria
    'DiegoProd
    If Me.Tipo = 2 Then
        sql = "DELETE FROM IVMateria Where IdMateria = " & mProps.IdInventario
    Else
        sql = "DELETE FROM IVMateria Where IdInventario = " & mProps.IdInventario
    End If
    mProps.objEmpresa.Execute sql, True
    
'    sql = "DELETE FROM ivsecuencia Where IdInventario = " & mProps.IdInventario
 '   mProps.objEmpresa.Execute sql, True
    
    
    'Elimina el registro
    sql = "DELETE FROM IVInventario WHERE IdInventario=" & mProps.IdInventario
    mProps.objEmpresa.Execute sql, True

    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_ITEM", "Item Eliminado: " & _
                                                         mProps.CodInventario, "DT"
                                                         
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_INVENTARIO
End Sub

Public Function Clone() As IVInventario
    Dim nuevo As IVInventario, ivm As IVFamiliaDetalle, i As Long
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New IVInventario
    nuevo.Props = mProps
    
    'Copia ivmatria al nuevo objeto
    For i = 1 To mcolDetalleFamilia.Count
        Set ivm = IVFamiliaDetalle(i).Clone
        nuevo.AddDetalleFamiliaClone ivm
    Next i
    
    
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Sub AddDetalleFamiliaClone(obj As IVFamiliaDetalle)
    Set obj.IVInventario = Me
    mcolDetalleFamilia.Add obj
End Sub


Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdInventario = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property
'AUC CAMBIADO A PUBLICO
Public Property Get IdInventario() As Long
    IdInventario = mProps.IdInventario
End Property


Public Property Let CodInventario(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "IVInventario_CodInventario", MSGERR_INVALIDO
    End If
'    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodInventario = value
    AsignarCodCuentasPre
    mbooModificado = True
End Property

Public Property Get CodInventario() As String
    CodInventario = Trim$(mProps.CodInventario)
End Property

Public Property Let CodAlterno1(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "IVInventario_CodAlterno1", MSGERR_INVALIDO
    End If
    
    mProps.CodAlterno1 = value
    mbooModificado = True
End Property

Public Property Get CodAlterno1() As String
    CodAlterno1 = Trim$(mProps.CodAlterno1)
End Property

Public Property Let CodAlterno2(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "IVInventario_CodAlterno2", MSGERR_INVALIDO
    End If
    
    mProps.CodAlterno2 = value
    mbooModificado = True
End Property

Public Property Get CodAlterno2() As String
    CodAlterno2 = Trim$(mProps.CodAlterno2)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 120 Then
        Err.Raise ERR_INVALIDO, "IVInventario_Descripcion", MSGERR_INVALIDO
    End If
    
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, ";") > 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario_Descripcion", _
            "El caracter ';' no es válido para la descripción."
    End If
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, "|") > 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario_Descripcion", _
            "El caracter '|' no es válido para la descripción."
    End If
    
'    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let DescripcionDetalle(ByVal value As String)
    mProps.DescripcionDetalle = value
    mbooModificado = True
End Property

Public Property Get DescripcionDetalle() As String
    DescripcionDetalle = Trim$(mProps.DescripcionDetalle)
End Property


Public Property Let CodMoneda(ByVal value As String)
    If Len(value) > 5 Then
        Err.Raise ERR_INVALIDO, "IVInventario_CodMoneda", MSGERR_INVALIDO
    End If
    
    mProps.CodMoneda = value
    mbooModificado = True
End Property

Public Property Get CodMoneda() As String
    CodMoneda = Trim$(mProps.CodMoneda)
End Property

Private Function CostoProm( _
                    ByVal fecha As Date, _
                    ByVal tid As Long, _
                    ByVal Hora As Date) As Double
    Dim sql As String, num As Integer, campo As String
    Dim rs As Recordset, Costo As Double        '*** MAKOTO 23/nov/00 Modificado a Double
    Dim idItem As Long
    num = mProps.objEmpresa.GNOpcion.IndiceMoneda(Me.CodMoneda)
    campo = "CP" & num
#If DAOLIB Then
    Dim qdef As QueryDef
    Set qdef = mProps.objEmpresa.Database.QueryDefs("spIVCostoProm")
        If Me.Tipo = Preparacion Then 'AUC agregado para sacar el costo de un item de preparacion
        qdef.Parameters("IdItem") = idInv
    Else
        qdef.Parameters("IdItem") = mProps.IdInventario
    End If
    'qdef.Parameters("IdItem") = mProps.IdInventario
    qdef.Parameters("FechaHasta") = fecha
    qdef.Parameters("tid") = tid
    qdef.Parameters("HoraHasta") = Hora         '*** MAKOTO 09/dic/00 Agregado
    Set rs = qdef.OpenRecordset(dbOpenSnapshot, dbReadOnly)
    If Not rs.EOF Then
        If Not IsNull(rs.Fields(campo)) Then Costo = rs.Fields(campo)
    End If
#Else
    '*** MAKOTO 08/dic/00 Modificado
    Debug.Print "ini- " & Time()
    If Me.Empresa.GNOpcion.IVKTipoDatoDouble Then
        sql = "spIVCostoProm2Dou " & _
                mProps.IdInventario & ", " & _
                FechaYMD(fecha, Me.Empresa.TipoDB) & ", " & _
                tid & ", " & _
                num & ", " & _
                HoraHMS(Hora, Me.Empresa.TipoDB)
    Else
        sql = "spIVCostoProm2 " & _
                mProps.IdInventario & ", " & _
                FechaYMD(fecha, Me.Empresa.TipoDB) & ", " & _
                tid & ", " & _
                num & ", " & _
                HoraHMS(Hora, Me.Empresa.TipoDB)
    End If
'                "'" & Format$(Hora, "Hh:Nn:Ss") & "'"  '*** MAKOTO 06/mar/01 Mod.
    Set rs = Me.Empresa.OpenRecordset(sql)
'    Debug.Print "fin-" & Time()
    Costo = rs.Fields("CostoProm")
#End If
    
    rs.Close
    Set rs = Nothing

    'Si no hay costo promedio por que no hay existencia,
    '   devuelve el último costo en vez de promedio
    If gobjMain.EmpresaActual.GNOpcion.IVKTipoDatoDouble Then
        If Costo = 0 Then Costo = CostoUltimoDou(fecha, tid, Hora, False)
    Else
        If Costo = 0 Then Costo = CostoUltimo(fecha, tid, Hora, False)
    End If
    
    CostoProm = Costo
End Function


Private Function CostoFIFO( _
                    ByVal fecha As Date, _
                    ByVal cant As Currency, _
                    ByVal tid As Long, _
                    ByVal Hora As Date) As Double
    Dim sql As String, j As Currency, k As Currency, c As Double
    Dim campo As String, num As Integer
    Dim rs As Recordset
    Dim idItem As Long 'AUC agregado para saber cual id del item tomar

    num = Me.Empresa.GNOpcion.IndiceMoneda(Me.CodMoneda)
    campo = "C" & num
    
    'Obtiene la existencia hasta la fecha
    sql = "SELECT Sum(Cantidad) AS SumaCant " & _
          "FROM GNTrans INNER JOIN (GNComprobante INNER JOIN IVKardex " & _
          "ON GNComprobante.TransID = IVKardex.TransID) ON GNTrans.CodTrans = GNComprobante.CodTrans " & _
          "WHERE (Estado <> 3) And (AfectaCantidad=" & CadenaBool(True, Me.Empresa.TipoDB) & ") " & _
            "AND (IdInventario=" & mProps.IdInventario & ") "
    'Condicion de fecha y tid
    '   Hasta antes de la fecha limite o la misma fecha pero con TransId menor
    sql = sql & "AND ((GNComprobante.FechaTrans<" & FechaYMD(fecha, Me.Empresa.TipoDB, True) & ") " & _
                  "OR (GNComprobante.FechaTrans=" & FechaYMD(fecha, Me.Empresa.TipoDB, True) & _
                  " AND (GNComprobante.TransID<" & tid & ")) )"

    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    j = 0
    If Not IsNull(rs!sumaCant) Then j = rs!sumaCant
    rs.Close
    
    'Si pide más de lo que existe
    If cant > j Then cant = j
    
    
    'Obtiene IVKardex ordenado por la fecha descendente
    sql = "SELECT Cantidad, CostoRealTotal/Cantidad AS C1, " & _
          "CostoRealTotal/Cotizacion2/Cantidad AS C2, " & _
          "CostoRealTotal/Cotizacion3/Cantidad AS C3, " & _
          "CostoRealTotal/Cotizacion4/Cantidad AS C4 " & _
          "FROM GNTrans INNER JOIN (GNComprobante INNER JOIN IVKardex " & _
          "ON GNComprobante.TransID = IVKardex.TransID) ON GNTrans.CodTrans = GNComprobante.CodTrans "
    sql = sql & "WHERE Estado<>3 And AfectaCantidad=" & CadenaBool(True, Me.Empresa.TipoDB) & _
          " AND IdInventario=" & mProps.IdInventario
    'Condicion de fecha y tid
    '   Hasta antes de la fecha limite o la misma fecha pero con TransId menor
    sql = sql & " AND ((GNComprobante.FechaTrans<" & FechaYMD(fecha, Me.Empresa.TipoDB, True) & ") " & _
                  "OR (GNComprobante.FechaTrans=" & FechaYMD(fecha, Me.Empresa.TipoDB, True) & _
                  " AND (GNComprobante.TransID<" & tid & ")) )"
'          " AND FechaTrans<=" & FechaYMD(Fecha, Me.Empresa.TipoDB)
    sql = sql & " ORDER BY FechaTrans DESC"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        If .EOF Then
            c = 0
        Else
            'Buscar el lote existente más antiguo
            Do Until .EOF
                If !Cantidad > 0 Then
                    j = j - !Cantidad
                    If j <= 0 Then Exit Do
                End If
                .MoveNext
            Loop
            
'            If not .EOF Then
            If .EOF Then .MoveLast
            j = 0
            
            If cant = 0 Then
                c = .Fields(campo)
            Else
                Do Until .BOF
                    If !Cantidad > 0 Then
                        If !Cantidad + j >= cant Then
                            k = cant - j
                            c = c + .Fields(campo) * k
                            Exit Do
                        Else
'                            k = !Cantidad + j
                            k = !Cantidad
                            j = j + !Cantidad
                            c = c + .Fields(campo) * k
                        End If
                    End If
                    .MovePrevious
                Loop
                c = c / cant
            End If
        End If
    End With
    rs.Close
    Set rs = Nothing
    CostoFIFO = c
End Function

Private Function CostoLIFO( _
                    ByVal fecha As Date, _
                    ByVal cant As Currency, _
                    ByVal tid As Long, _
                    ByVal Hora As Date) As Double
    Dim sql As String, j As Currency, k As Currency, c As Double
    Dim campo As String, num As Integer
    Dim rs As Recordset
    
    num = Me.Empresa.GNOpcion.IndiceMoneda(Me.CodMoneda)
    campo = "C" & num
    
    'Obtiene IVKardex ordenado por la fecha descendente
    sql = "SELECT Cantidad,CostoRealTotal/Cantidad AS C1, " & _
          "CostoRealTotal/Cotizacion2/Cantidad AS C2, " & _
          "CostoRealTotal/Cotizacion3/Cantidad AS C3, " & _
          "CostoRealTotal/Cotizacion4/Cantidad AS C4 " & _
          "FROM GNTrans INNER JOIN (GNComprobante INNER JOIN IVKardex " & _
          "ON GNComprobante.TransID = IVKardex.TransID) ON GNTrans.CodTrans = GNComprobante.CodTrans "
    sql = sql & "WHERE (Estado <> 3) And (AfectaCantidad=" & CadenaBool(True, Me.Empresa.TipoDB) & ") AND (IdInventario=" & mProps.IdInventario & ")"
    'Condicion de fecha y tid
    '   Hasta antes de la fecha limite o la misma fecha pero con TransId menor
    sql = sql & " AND ((GNComprobante.FechaTrans<" & FechaYMD(fecha, Me.Empresa.TipoDB, True) & ") " & _
                  "OR (GNComprobante.FechaTrans=" & FechaYMD(fecha, Me.Empresa.TipoDB, True) & _
                  " AND (GNComprobante.TransID<" & tid & ")) )"
'          " AND (FechaTrans<=" & FechaYMD(Fecha, Me.Empresa.TipoDB) & ")"
    sql = sql & " ORDER BY FechaTrans DESC"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        If rs.EOF Then
            c = 0
        Else
            If cant = 0 Then
                c = .Fields(campo)
            Else
                Do Until .BOF
                    If !Cantidad > 0 Then
                        If !Cantidad + j >= cant Then
                            c = c + .Fields(campo) * (cant - j)
                            Exit Do
                        Else
                            j = j + !Cantidad
                            c = c + .Fields(campo) * (!Cantidad + j)
                        End If
                    End If
                    .MovePrevious
                Loop
                c = c / cant
            End If
        End If
    End With
    rs.Close
    Set rs = Nothing
    CostoLIFO = c
End Function

Private Function CostoUltimo( _
                    ByVal fecha As Date, _
                    ByVal tid As Long, _
                    ByVal Hora As Date, _
                    ByVal BandSoloIngreso As Boolean) As Double
    Dim sql As String, campo As String, num As Integer
    Dim rs As Recordset, cadena_hora As String
    
    num = Me.Empresa.GNOpcion.IndiceMoneda(Me.CodMoneda)
    campo = "1"
    If num > 1 Then campo = "gc.Cotizacion" & num
    
    'Obtiene IVKardex ordenado por la fecha descendiente para coger último ingreso
    ' Busca solo de FechaTrans<=Fecha, y solo de INGRESOS y no anulados
    sql = "SELECT ivk.CostoRealTotal/ivk.Cantidad/" & campo & " AS CU " & _
          "FROM GNTrans gt INNER JOIN " & _
                    "(GNComprobante gc INNER JOIN IVKardex ivk " & _
                    "ON gc.TransID = ivk.TransID) " & _
                "ON gt.CodTrans = gc.CodTrans "
    sql = sql & "WHERE (gc.Estado <> 3) AND " & _
                    "(gt.AfectaCantidad=" & CadenaBool(True, Me.Empresa.TipoDB) & ") AND " & _
                    "(ivk.IdInventario=" & mProps.IdInventario & ") "
    If BandSoloIngreso Then sql = sql & " AND (ivk.Cantidad>0) "
          
    'Condicion de fecha y tid
    '   Hasta antes de la fecha limite
    '   o la misma fecha pero la hora menor
    '   o la misma fecha y misma hora pero tid menor
'    cadena_hora = Format(Hora, "Hh:Nn:Ss")
'#If DAOLIB Then
'    cadena_hora = "#" & cadena_hora & "#"
'#Else
'    cadena_hora = "'" & cadena_hora & "'"
'#End If
    cadena_hora = HoraHMS(Hora, Me.Empresa.TipoDB)        '*** MAKOTO 06/mar/01 Mod.
    
    sql = sql & " AND ( (gc.FechaTrans<" & FechaYMD(fecha, Me.Empresa.TipoDB) & ") " & _
                    "OR (gc.FechaTrans=" & FechaYMD(fecha, Me.Empresa.TipoDB) & _
                        " AND gc.HoraTrans<" & cadena_hora & ") " & _
                    "OR (gc.FechaTrans=" & FechaYMD(fecha, Me.Empresa.TipoDB) & _
                        " AND gc.HoraTrans=" & cadena_hora & _
                        " AND gc.TransID < " & tid & ") )"
    
    sql = sql & " ORDER BY gc.FechaTrans DESC, gc.HoraTrans DESC, gc.TransID DESC"     '*** MAKOTO 08/dic/00 HoraTrans Agregado
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Coge el primer costo como el costo ultimo
    If Not rs.EOF Then
        If Not IsNull(rs.Fields("CU")) Then
            CostoUltimo = rs.Fields("CU")
        '*** Oliver 10 sep 2004
        'Aseguro de devolver 0 porque no encontro mov. de item
        Else
            CostoUltimo = 0
        End If
    Else
        CostoUltimo = 0
    End If
    rs.Close
    Set rs = Nothing
End Function


'Calcula y devuelve el costo del item
'Fehca: Fecha limite para el calculo.
'cant : Es necesario solo cuando es FIFO o LIFO
'       Indica qué cantidad va a coger para calcular el costo
'tid  : TransID para calcular hasta antes de esa transaccion
'           cuando existe varias trans. en la misma fecha
Public Function Costo( _
                    ByVal fecha As Date, _
                    Optional ByVal cant As Currency, _
                    Optional ByVal tid As Long) As Currency
    Costo = CCur(CostoDouble2(fecha, cant, tid))     '*** MAKOTO 08/dic/00
End Function

'*** MAKOTO 23/nov/00 Agregado
'Calcula y devuelve el costo del item
'Fehca: Fecha limite para el calculo.
'cant : Es necesario solo cuando es FIFO o LIFO
'       Indica qué cantidad va a coger para calcular el costo
'tid  : TransID para calcular hasta antes de esa transaccion
'           cuando existe varias trans. en la misma fecha
Public Function CostoDouble( _
                    ByVal fecha As Date, _
                    Optional ByVal cant As Currency, _
                    Optional ByVal tid As Long) As Double
    '*** MAKOTO 08/dic/00 Modificado
    CostoDouble = CostoDouble2(fecha, cant, tid)
End Function

'*** MAKOTO 08/dic/00 Agregado
'Calcula y devuelve el costo del item
'Fehca: Fecha limite para el calculo.
'cant : Es necesario solo cuando es FIFO o LIFO
'       Indica qué cantidad va a coger para calcular el costo
'tid  : TransID para calcular hasta antes de esa transaccion
'           cuando existe varias trans. en la misma fecha
'Hora : Hora límite para el calculo.
Public Function CostoDouble2( _
                    ByVal fecha As Date, _
                    Optional ByVal cant As Currency, _
                    Optional ByVal tid As Long, _
                    Optional ByVal Hora As Date) As Double
    'Si la id de transaccion no está especificada, cambia la fecha limite
    '  para que incluya todas las trans. de la misma fecha
    If Hora = 0 Then Hora = TimeSerial(23, 59, 59)

    Select Case mProps.objEmpresa.GNOpcion.Costeo
    Case COSTO_PROM
        CostoDouble2 = CostoProm(fecha, tid, Hora)
    Case COSTO_FIFO
        CostoDouble2 = CostoFIFO(fecha, cant, tid, Hora)
    Case COSTO_LIFO
        CostoDouble2 = CostoLIFO(fecha, cant, tid, Hora)
    Case COSTO_ULTIMO
        CostoDouble2 = CostoUltimo(fecha, tid, Hora, True)
    End Select
    
    '*** MAKOTO 31/ene/01 Agregado para redondear al formato de costo
    CostoDouble2 = MiCCur(Format$(CostoDouble2, Me.Empresa.GNOpcion.ObtenerValor( _
                            "FormatoCosto" & Me.Empresa.GNOpcion.IndiceMoneda(Me.CodMoneda))))
    CostoDouble2 = MiCDbl(Format$(CostoDouble2, Me.Empresa.GNOpcion.ObtenerValor( _
                            "FormatoCosto" & Me.Empresa.GNOpcion.IndiceMoneda(Me.CodMoneda))))


                            
End Function

Public Property Let Precio(ByVal ix As Integer, ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario.Precio" & ix, _
            "El precio no puede ser negativo. ('" & mProps.CodInventario & "')"
    End If
    
    '*** MAKOTO 30/nov/00 Agregado
    'Si no tiene permiso para modificar precios, genera error
    If gobjMain.GrupoActual.PermisoActual.CatInventarioPrecioMod = False Then
        Err.Raise ERR_NODERECHO, "IVInventario.Precio" & ix, MSGERR_NODERECHO
    End If
    
    'Si es que realmente modifica el valor
    If mProps.Precio(ix) <> value Then
        mProps.Precio(ix) = value
        mbooModificado = True
    End If
End Property

Public Property Get Precio(ByVal ix As Integer) As Currency
    Precio = mProps.Precio(ix)
End Property

Public Sub RedondearPrecio(ByVal ix As Integer, ByVal digitos As Integer)
    mProps.Precio(ix) = Redondear(mProps.Precio(ix), digitos)
End Sub


Public Property Let Comision(ByVal ix As Integer, ByVal value As Currency)
    mProps.Comision(ix) = value
    mbooModificado = True
End Property

Public Property Get Comision(ByVal ix As Integer) As Currency
    Comision = mProps.Comision(ix)
End Property

Public Property Let CantLimite(ByVal ix As Integer, ByVal value As Currency)
    mProps.CantLimite(ix) = value
    mbooModificado = True
End Property

Public Property Get CantLimite(ByVal ix As Integer) As Currency
    CantLimite = mProps.CantLimite(ix)
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

Public Property Let IdCuentaActivo(ByVal value As Long)
    mProps.IdCuentaActivo = value
    mbooModificado = True
End Property

Public Property Get IdCuentaActivo() As Long
    IdCuentaActivo = mProps.IdCuentaActivo
End Property


Public Property Let CodCuentaActivo(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodCuentaActivo Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVInventario_CodCuentaActivo", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaActivo = rs!IdCuenta
            End If
            rs.Close
        Else
            IdCuentaActivo = 0
        End If
        
        mProps.CodCuentaActivo = value
        mbooModificado = True
    End If

    Set rs = Nothing
End Property

Public Property Get CodCuentaActivo() As String
    CodCuentaActivo = Trim$(mProps.CodCuentaActivo)
End Property

Public Property Let IdCuentaCosto(ByVal value As Long)
    mProps.IdCuentaCosto = value
    mbooModificado = True
End Property


Public Property Get IdCuentaCosto() As Long
    IdCuentaCosto = mProps.IdCuentaCosto
End Property


Public Property Let CodCuentaCosto(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodCuentaCosto Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaCosto también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVInventario_CodCuentaCosto", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaCosto = rs!IdCuenta
            End If
            rs.Close
        Else
            IdCuentaCosto = 0
        End If
        
        mProps.CodCuentaCosto = value
        mbooModificado = True
    End If

    Set rs = Nothing
End Property

Public Property Get CodCuentaCosto() As String
    CodCuentaCosto = Trim$(mProps.CodCuentaCosto)
End Property

Public Property Let IdCuentaVenta(ByVal value As Long)
    mProps.IdCuentaVenta = value
    mbooModificado = True
End Property

Public Property Get IdCuentaVenta() As Long
    IdCuentaVenta = mProps.IdCuentaVenta
End Property


Public Property Let CodCuentaVenta(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCuentaVenta Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaVenta también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVInventario_CodCuentaVenta", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaVenta = rs!IdCuenta
            End If
            rs.Close
        Else
            IdCuentaVenta = 0
        End If
    
        mProps.CodCuentaVenta = value
        mbooModificado = True
    End If

    Set rs = Nothing
End Property

Public Property Get CodCuentaVenta() As String
    CodCuentaVenta = Trim$(mProps.CodCuentaVenta)
End Property

Public Property Let Unidad(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "IVInventario_Unidad", MSGERR_INVALIDO
    End If
    
    mProps.Unidad = value
    mbooModificado = True
End Property

Public Property Get Unidad() As String
    Unidad = Trim$(mProps.Unidad)
End Property

Public Property Let IdGrupo(NumGrupo As Integer, ByVal value As Long)
    mProps.IdGrupo(NumGrupo) = value
    mbooModificado = True
End Property

Public Property Get IdGrupo(NumGrupo As Integer) As Long
    IdGrupo = mProps.IdGrupo(NumGrupo)
End Property


Public Property Let CodGrupo(NumGrupo As Integer, ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodGrupo(NumGrupo) Then
        If Len(value) > 0 Then
            'Actualiza IdGrupo1 también
            sql = "SELECT IdGrupo" & NumGrupo & " FROM IVGrupo" & NumGrupo & " WHERE CodGrupo" & NumGrupo & "='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVInventario_CodGrupo" & NumGrupo, MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdGrupo(NumGrupo) = rs.Fields("IdGrupo" & NumGrupo)
            End If
            rs.Close
        Else
            IdGrupo(NumGrupo) = 0
        End If
    End If
    Set rs = Nothing
    mProps.CodGrupo(NumGrupo) = value
    mbooModificado = True
End Property

Public Property Get CodGrupo(NumGrupo As Integer) As String
    CodGrupo = Trim$(mProps.CodGrupo(NumGrupo))
End Property

Public Property Let idProveedor(ByVal value As Long)
    mProps.idProveedor = value
    mbooModificado = True
End Property

Public Property Get idProveedor() As Long
    idProveedor = mProps.idProveedor
End Property


Public Property Let CodProveedor(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodProveedor Then
        If Len(value) > 0 Then
            'Actualiza IdProveedor también
            sql = "SELECT IdProvCli FROM PCProvCli WHERE CodProvCli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVInventario_CodProveedor", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                idProveedor = rs!IdProvcli
            End If
            rs.Close
        Else
            idProveedor = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodProveedor = value
    mbooModificado = True
End Property

Public Property Get CodProveedor() As String
    CodProveedor = Trim$(mProps.CodProveedor)
End Property



Public Property Let Observacion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "IVInventario_Observacion", MSGERR_INVALIDO
    End If
    
    mProps.Observacion = value
    mbooModificado = True
End Property

Public Property Get Observacion() As String
    Observacion = Trim$(mProps.Observacion)
End Property

Public Property Let ExistenciaMinima(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario_ExistenciaMinima", MSGERR_INVALIDO
    End If
    
    mProps.ExistenciaMinima = value
    mbooModificado = True
End Property

Public Property Get ExistenciaMinima() As Currency
    ExistenciaMinima = mProps.ExistenciaMinima
End Property

Public Property Let ExistenciaMaxima(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario_ExistenciaMaxima", MSGERR_INVALIDO
    End If
    
    mProps.ExistenciaMaxima = value
    mbooModificado = True
End Property

Public Property Get ExistenciaMaxima() As Currency
    ExistenciaMaxima = mProps.ExistenciaMaxima
End Property

Public Property Let UnidadMinimaCompra(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario_UnidadMinimaCompra", MSGERR_INVALIDO
    End If
    
    mProps.UnidadMinimaCompra = value
    mbooModificado = True
End Property

Public Property Get UnidadMinimaCompra() As Currency
    UnidadMinimaCompra = mProps.UnidadMinimaCompra
End Property

Private Function ObtieneUltimo(BandIngreso As Boolean) As Boolean
    Dim sql As String, rs As Recordset, cond1 As String
    Dim campo As String
    On Error GoTo ErrTrap
    
    If BandIngreso Then
        If mUIngreso.Recuperado Then
            ObtieneUltimo = True
            Exit Function
        End If
    Else
        If mUEgreso.Recuperado Then
            ObtieneUltimo = True
            Exit Function
        End If
    End If
    
    cond1 = IIf(BandIngreso, "(Ingreso>0)", "(Egreso>0)")
    
    Select Case Me.Empresa.GNOpcion.IndiceMoneda(Me.CodMoneda)
    Case 1: campo = "Costo AS CU,Precio AS PU"
    Case 2: campo = "Costo/Cotizacion2 AS CU,Precio/Cotizacion2 AS PU"
    Case 3: campo = "Costo/Cotizacion3 AS CU,Precio/Cotizacion3 AS PU"
    Case 4: campo = "Costo/Cotizacion4 AS CU,Precio/Cotizacion4 AS PU"
    End Select
    
    '*** MAKOTO 07/dic/00 Agregado HoraTrans DESC
    sql = "SELECT FechaTrans,Ingreso,Egreso," & campo & " FROM vwConsIVKardex " & _
          "WHERE " & cond1 & " AND (CodInventario='" & Me.CodInventario & "') " & _
          "ORDER BY FechaTrans DESC, HoraTrans DESC, TransID DESC"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        If BandIngreso Then
            With mUIngreso
                .fecha = rs!FechaTrans
                .Cantidad = rs!ingreso
                .Costo = rs!cu
                .Precio = rs!PU
                .Recuperado = True
            End With
        Else
            With mUEgreso
                .fecha = rs!FechaTrans
                .Cantidad = rs!egreso
                .Costo = rs!cu
                .Precio = rs!PU
                .Recuperado = True
            End With
        End If
        ObtieneUltimo = True
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
ErrTrap:
    Set rs = Nothing
    Exit Function
End Function


'jeaa cambiado ObtieneUltimo 31/03/2010
Public Property Get FechaUltimoIngreso() As Date
    If ObtieneUltimoCompra(True) Then FechaUltimoIngreso = mUIngreso.fecha
End Property

Public Property Get CantUltimoIngreso() As Currency
    If ObtieneUltimoCompra(True) Then CantUltimoIngreso = mUIngreso.Cantidad
End Property

'AUC
Public Property Let CostoUltimoIngreso(ByVal value As Currency)
    mProps.CostoUltimoIngreso = value
    mbooModificado = True
End Property
'AUC
Public Property Get CostoUltimoIngreso() As Currency
    'If ObtieneUltimoCompra(True) Then CostoUltimoIngreso = mUIngreso.Costo  ' jeaa anulado 15/05/2013
    CostoUltimoIngreso = mProps.CostoUltimoIngreso
End Property


Public Property Let CostoReferencial(ByVal value As Currency)
    If mProps.CostoReferencial <> value Then
        mProps.CostoReferencial = value
        mbooModificado = True
    End If
End Property


'AUC
Public Property Get CostoReferencial() As Currency
    CostoReferencial = mProps.CostoReferencial
End Property
Public Property Get FechaUltimoEgreso() As Date
    If ObtieneUltimo(False) Then FechaUltimoEgreso = mUEgreso.fecha
End Property

Public Property Get CantUltimoEgreso() As Currency
    If ObtieneUltimo(False) Then CantUltimoEgreso = mUEgreso.Cantidad
End Property

Public Property Get CostoUltimoEgreso() As Currency
    If ObtieneUltimo(False) Then CostoUltimoEgreso = mUEgreso.Costo
End Property

Public Property Get PrecioUltimoEgreso() As Currency
    If ObtieneUltimo(False) Then PrecioUltimoEgreso = mUEgreso.Precio
End Property

Public Property Get NumFamiliaDetalle() As Long
    If mcolDetalleFamilia Is Nothing Then
        NumFamiliaDetalle = 0
    Else
        NumFamiliaDetalle = mcolDetalleFamilia.Count
    End If
    'If ObtieneUltimo(False) Then PrecioUltimoEgreso = mUEgreso.Precio
End Property
'AUC 23/11/05
Public Property Get NumProveedorDetalle() As Long
    If mcolDetalleProveedor Is Nothing Then
        NumProveedorDetalle = 0
    Else
        NumProveedorDetalle = mcolDetalleProveedor.Count
    End If
    'If ObtieneUltimo(False) Then PrecioUltimoEgreso = mUEgreso.Precio
End Property




Public Property Let UnidadMinimaVenta(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario_UnidadMinimaVenta", MSGERR_INVALIDO
    End If
    
    mProps.UnidadMinimaVenta = value
    mbooModificado = True
End Property

Public Property Get UnidadMinimaVenta() As Currency
    UnidadMinimaVenta = mProps.UnidadMinimaVenta
End Property

Public Property Let BandValida(ByVal value As Boolean)
    mProps.BandValida = value
    mbooModificado = True
End Property

Public Property Get BandValida() As Boolean
    BandValida = mProps.BandValida
End Property

Public Property Let BandServicio(ByVal value As Boolean)
    mProps.BandServicio = value
    mbooModificado = True
End Property

Public Property Get BandServicio() As Boolean
    BandServicio = mProps.BandServicio
End Property




Public Property Let Tipo(ByVal value As EIVTipo)
    mProps.Tipo = value
    mbooModificado = True
End Property

Public Property Get Tipo() As EIVTipo
    Tipo = mProps.Tipo
End Property

'***Agregado. Angel. 29/jul/2004
Public Property Let ValorRecargo(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario.ValorRecargo", _
            "El Recargo no puede ser negativo. ('" & mProps.CodInventario & "')"
    End If
    
    'Si es que realmente modifica el valor
    If mProps.ValorRecargo <> value Then
        mProps.ValorRecargo = value
        mbooModificado = True
    End If
End Property

'***Agregado. Angel. 29/jul/2004
Public Property Get ValorRecargo() As Currency
    ValorRecargo = mProps.ValorRecargo
End Property
'agregado esteban 08/04/2005
Public Property Let BandFraccion(ByVal value As Boolean)
    mProps.BandFraccion = value
    mbooModificado = True
End Property

Public Property Get BandFraccion() As Boolean
    BandFraccion = mProps.BandFraccion
End Property

'agregado esteban 15/09/2005
Public Property Let BandArea(ByVal value As Boolean)
    mProps.BandArea = value
    mbooModificado = True
End Property

Public Property Get BandArea() As Boolean
    BandArea = mProps.BandArea
End Property

'agregado esteban 15/09/2005
Public Property Let bandVenta(ByVal value As Boolean)
    mProps.bandVenta = value
    mbooModificado = True
End Property

Public Property Get bandVenta() As Boolean
    bandVenta = mProps.bandVenta
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

Friend Property Let Modificado(ByVal value As Boolean)
    mbooModificado = value
End Property

Public Function Existencia(CodBodega As String) As Currency
    Dim obj As IVInventarioDetalle
    
    If mcolDetalles Is Nothing Then RecuperaDetalle
    If Len(CodBodega) > 0 Then          '*** MAKOTO 15/dic/00
        For Each obj In mcolDetalles
            If obj.CodBodega = CodBodega Then
                Existencia = obj.Cantidad
                Exit For
            End If
        Next obj
    Else
        'Si no está especifada una bodega, devuelve la suma de todas las bodegas    '*** MAKOTO 15/dic/00
        For Each obj In mcolDetalles
            Existencia = Existencia + obj.Cantidad
        Next obj
    End If
    Set obj = Nothing
End Function

Public Function ListaPrecioParaFlex(ByVal gc As GNComprobante) As String
    Dim s As String, r As Double, fmt As String
    
    'fmt = gc.Empresa.GNOpcion.FormatoMoneda(gc.CodMoneda)
    'JEAA 24/01/05
    fmt = gc.Empresa.GNOpcion.ObtenerValor("FormatoPU" & gc.Empresa.GNOpcion.IndiceMoneda(MONEDA_SEC))
    
    '*** MAKOTO 28/sep/00 Agregado para corregir error en caso de '#,#'
    If Right$(fmt, 2) = ",#" Then fmt = fmt & "0"
    
    r = 1#
    
    'Si la transacción está en una moneda distinta de la del item
    If gc.CodMoneda <> mProps.CodMoneda Then
        'Calcula el factor para el calculo de equivalencia
        r = gc.Cotizacion(mProps.CodMoneda) / gc.Cotizacion(gc.CodMoneda)
    End If
    
    s = "|" & Format(Precio(1) * r, fmt) & _
        "|" & Format(Precio(2) * r, fmt) & _
        "|" & Format(Precio(3) * r, fmt) & _
        "|" & Format(Precio(4) * r, fmt) & _
        "|" & Format(Precio(5) * r, fmt) & _
        "|" & Format(Precio(6) * r, fmt) & _
        "|" & Format(Precio(7) * r, fmt) 'GRV Precio 7 - 14MAY2014

        'AUC 27/09/07

    ListaPrecioParaFlex = s
End Function

'Diego Prod
Public Function ListaDetalleFamilia(BandRS As Boolean) As Variant
    Dim sql As String, rs As Recordset
    If Me.Tipo <> Normal Then
        sql = "SELECT IV.CodInventario,IV.Descripcion,  IVM.Cantidad " & _
              "FROM IVINVENTARIO IV  Inner Join IVMateria IVM " & _
              "ON IV.IdInventario = IVM.IdInventario " & _
              "Where iv.IdInventario = " & Me.IdInventario
              
        Set rs = Me.Empresa.OpenRecordset(sql)
        If BandRS Then
            Set ListaDetalleFamilia = rs
        Else
            ListaDetalleFamilia = MiGetRows(rs)
            rs.Close
        End If
        Set rs = Nothing
    End If
End Function

Public Sub GenerarCodigo()
    Dim sql As String, num As Double, cod As String, rs As Recordset
    Dim pre As String, formato As String
    Dim sqlaux As String

    pre = Me.Empresa.GNOpcion.PrefijoCodItem
    formato = Me.Empresa.GNOpcion.FormatoCodItem
    Select Case Me.Empresa.GNOpcion.GeneraCodItem
    Case 0  'Manual
        cod = pre
    Case 1  'Auto secuencial
        cod = Me.CodInventario
#If DAOLIB Then
        sql = "SELECT Max(VAL(CodInventario)) AS CodMax FROM IVInventario " & _
              "WHERE CodInventario LIKE '" & pre & "*'"
#Else
        If InStr(1, UCase(pre), "IV") Then
            Select Case Mid$(pre, 3, 1)
            Case 1
            End Select
                If Len(Me.CodGrupo(Mid$(pre, 3, 1))) = 0 Then
                    MsgBox "Según la configuración para poder generar el código, primero debe seleccionar " & Chr(13) & Me.Empresa.GNOpcion.EtiqGrupo(1)
                    Me.CodInventario = ""
                    Exit Sub
                Else
                    pre = Me.CodGrupo(Mid$(pre, 3, 1))
                End If
      End If
        
        sql = "SELECT Max(str(abs(CodInventario))) AS CodMax FROM IVInventario " & _
              "WHERE CodInventario LIKE '" & pre & "[0-9][0-9]%'"

#End If
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then
            If Not IsNull(rs!CodMax) Then num = val(rs!CodMax)
        End If
        rs.Close
        If num = 0 Then
            num = num + 1
            num = pre + Format(num, formato)
        Else
            num = num + 1
        End If
        cod = Format(num, formato)
    Case 2  'Auto con formato (g1+g2+"."+CodAlterno1)
        cod = mProps.CodGrupo(1) & mProps.CodGrupo(2) & "." & mProps.CodAlterno1
    Case 3 'Auto (Código Generado en Base a Grupos)
        cod = GeneraCodigoEncadenado
    End Select
    Set rs = Nothing
    
    'Verifica la longitud de código
    If Len(cod) > 20 Then
        Err.Raise ERR_INVALIDO, "IVInventario_CodInventario", _
            "No se puede generar debido a que el código se hace muy largo."
    End If
    
    'Otra Verificación se podría colocar aquí para otras reglas de codificación.
    '***
    
    'Si no hay error guarda en la propiedad
    Me.CodInventario = cod
End Sub

Public Property Get CountDetalles() As Long
    If mcolDetalles Is Nothing Then RecuperaDetalle
    
    CountDetalles = mcolDetalles.Count
End Property

Public Property Get IVFamiliaDetalle(ix As Long) As IVFamiliaDetalle
    Set IVFamiliaDetalle = mcolDetalleFamilia.item(ix)
End Property


Private Sub CargaDetalleFamilia()
    Dim rs As Recordset, sql As String, obj As IVFamiliaDetalle
    Dim i As Long
    Set mcolDetalleFamilia = New Collection
    If Me.Tipo = Preparacion Then
       sql = "Select IV.IdInventario, Iv.CodInventario, Iv.Descripcion, " & _
              "Iv.CodUnidad, Ivm.Cantidad, ivm.bandPrincipal, ivm.bandModificar ,iv.costoultimoingreso   From VWIvinventarioRECUPERAR Iv " & _
            "INNER JOIN IVMateria Ivm "
    ElseIf Me.Tipo = CambioPresentacion Then
        sql = "Select IV.IdInventario, Iv.CodInventario, Iv.Descripcion, " & _
              "Iv.CodUnidad, Ivm.Cantidad  From VWIvinventarioRECUPERAR Iv " & _
           "INNER JOIN IVMateria Ivm "
    ElseIf Me.Tipo = rubro Then
           sql = "Select iv.codgrupo6,IV.IdInventario, Iv.CodInventario, Iv.Descripcion, " & _
           "ivm.bandModificar,bandPrincipal,ivm.orden ,ivp.codproceso, " & _
              "ivm.cantidad,iv.codunidad,iv.costoultimoingreso,iv.costoreferencial,ivm.rendimiento From VWIvinventarioRECUPERAR Iv  " & _
            "INNER JOIN IVMateria Ivm  Left Join ivproceso ivp on ivp.idproceso = ivm.idproceso "
    ElseIf Me.Tipo = Promocion Then
        sql = "Select IV.IdInventario, Iv.CodInventario, Iv.Descripcion, " & _
             " Iv.CodUnidad,Ivm.Cantidad, ivm.xcuanto, ivm.tarifajornal   From VWIvinventarioRECUPERAR Iv " & _
            "INNER JOIN IVMateria Ivm "
'    ElseIf Me.Tipo = rubro Then
'       sql = "Select IV.IdInventario, Iv.CodInventario, Iv.Descripcion, " & _
'              "ivm.formula,ivm.bandModificar,ivm.orden From VWIvinventarioRECUPERAR Iv " & _
'            "INNER JOIN IVMateria Ivm "
    ElseIf Me.Tipo = Preparacionporce Then
       sql = "Select IV.IdInventario, Iv.CodInventario, Iv.Descripcion, " & _
              "Iv.CodUnidad, Ivm.Cantidad, ivm.Rendimiento, ivm.bandPrincipal From VWIvinventarioRECUPERAR Iv " & _
            "INNER JOIN IVMateria Ivm "
    Else
        sql = "Select IV.IdInventario, Iv.CodInventario, Iv.Descripcion, " & _
              "Iv.codunidad, Ivm.Cantidad From  VWIvinventarioRECUPERAR IV " & _
              "INNER JOIN IVMateria Ivm "
    End If
    '03/10/2005
    If Me.Tipo = CambioPresentacion _
        Or Me.Tipo = Preparacion _
        Or Me.Tipo = Promocion Then


        sql = sql & "  ON IV.IdInventario = Ivm.Idmateria"
        sql = sql & " Where Ivm.IdInventario = " & mProps.IdInventario
        sql = sql & " order by orden"
    ElseIf Me.Tipo = rubro Then
        sql = sql & "  ON IV.IdInventario = Ivm.Idmateria"
        sql = sql & " Where Ivm.IdInventario = " & mProps.IdInventario
        sql = sql & " order by codgrupo6,orden"
   ElseIf Me.Tipo = Preparacionporce Then
        sql = sql & "  ON IV.IdInventario = Ivm.Idmateria"
        sql = sql & " Where Ivm.IdInventario = " & mProps.IdInventario
        sql = sql & " AND IV.TIPO<>2 "
        sql = sql & " order by orden"
    Else
          sql = sql & " ON IV.IdInventario = Ivm.IdInventario "
            sql = sql & " Where Ivm.IdMateria = " & mProps.IdInventario
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddDetalleFamilia
            Set obj = mcolDetalleFamilia.item(i)
            If Me.Tipo = rubro Then
                If Not IsNull(!CodGrupo6) Then obj.CodGrupo6 = !CodGrupo6
               obj.IdInventario = !IdInventario
                obj.CodInventario = !CodInventario
                obj.Cantidad = !Cantidad
                obj.Descripcion = !Descripcion
                 If Not IsNull(!CodUnidad) Then obj.Unidad = !CodUnidad 'AUC 'cambiado antes unidad
                If Not IsNull(!CostoUltimoIngreso) Then obj.CostoUltimoIngreso = !CostoUltimoIngreso
                If Not IsNull(!CostoReferencial) Then obj.CostoReferencial = !CostoReferencial
                obj.BandModificar = !BandModificar
                obj.BandPrincipal = !BandPrincipal
                If Not IsNull(!CodProceso) Then
                    obj.CodProceso = !CodProceso
                End If
                obj.Orden = !Orden
            Else
                  obj.IdInventario = !IdInventario
                  obj.CodInventario = !CodInventario
                  obj.Descripcion = !Descripcion
                  If Not IsNull(!CodUnidad) Then obj.Unidad = !CodUnidad 'AUC 'cambiado antes unidad
                 obj.Cantidad = !Cantidad
                  If Me.Tipo = Preparacion Then
                      If Len(obj.IVInventario.Empresa.GNOpcion.ObtenerValor("itemPreparacion")) > 0 Then 'AUC toma el p.costo
                          If obj.IVInventario.Empresa.GNOpcion.ObtenerValor("itemPreparacion") = 0 Then 'AUC toma el p.costo
                             If Not IsNull(!CostoUltimoIngreso) Then obj.CostoUltimoIngreso = !CostoUltimoIngreso
                           ElseIf obj.IVInventario.Empresa.GNOpcion.ObtenerValor("itemPreparacion") = 1 Then 'AUC toma el p. venta
                              obj.pVenta = obj.ObtienePrecio(obj.CodInventario)
                           End If
                           obj.BandPrincipal = !BandPrincipal  'AUC 10/05/06 restriccion del item
                           obj.BandModificar = !BandModificar   'AUC 25/05/06
                      Else
                          obj.pVenta = obj.ObtienePrecio(obj.CodInventario)
                      End If
                  ElseIf Me.Tipo = Promocion Then
                          obj.xCuanto = !xCuanto
                          obj.Precio = !TarifaJornal
                  ElseIf Me.Tipo = Preparacionporce Then
                        obj.BandPrincipal = !BandPrincipal  'AUC 10/05/06 restriccion del item
                        obj.Rendimiento = !Rendimiento 'AUC 10/05/06 restriccion del item
                End If
            End If
           .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub


Public Function RecuperaDetalleFamilia(ix As Long) As IVFamiliaDetalle
    Set RecuperaDetalleFamilia = mcolDetalleFamilia.item(ix)
End Function
'AUC 22/11/2005
Public Function RecuperaDetalleProveedor(ix As Integer) As IVDetalleProveedor
    Set RecuperaDetalleProveedor = mcolDetalleProveedor.item(ix)
End Function

Private Sub RecuperaDetalle()
    Dim sql As String, i As Long, obj As IVInventarioDetalle
    Dim rs As Recordset
    Set mcolDetalles = New Collection
    
    'Recupera los detalles de existencia
    sql = "SELECT IVExist.IdInventario, IVInventario.CodInventario, " & _
          "IVExist.IdBodega, IVBodega.CodBodega,  "
    If Me.Empresa.GNOpcion.IVKTipoDatoDouble Then
        sql = sql & " IVExist.ExistDou"
    Else
        sql = sql & " IVExist.Exist"
    End If
          
          sql = sql & " FROM (IVExist INNER JOIN IVInventario ON IVExist.IdInventario=" & _
          "IVInventario.IdInventario) INNER JOIN IVBodega ON IVExist.IdBodega " & _
          "= IVBodega.IdBodega " & _
          "WHERE IVExist.IdInventario=" & mProps.IdInventario & " ORDER BY CodBodega"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddDetalle
            Set obj = mcolDetalles.item(i)
            obj.IdBodega = !IdBodega
            obj.CodBodega = !CodBodega
            If Me.Empresa.GNOpcion.IVKTipoDatoDouble Then
                obj.Cantidad = !existDou
            Else
                obj.Cantidad = !exist
            End If
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
End Sub

Private Sub Class_Initialize()
    'Valores predeterminados
    With mProps
'        .CodMoneda = MONEDA_PRE
        .CodMoneda = "USD"              '*** MAKOTO 02/oct/00 Modificado (PROVISIONALMENTE)
        .PorcentajeIVA = 0
        .PorcentajeIVAAnt = 0
        .ExistenciaMaxima = 0
        .ExistenciaMinima = 0
        .UnidadMinimaCompra = 1
        .UnidadMinimaVenta = 1
        .Unidad = "Unidad"
        .BandValida = True
        .BandServicio = False
        .BandFraccion = False       'Agregado jeaa 08/04/2005
        .BandArea = False       'Agregado jeaa 15/09/2005
        .bandVenta = True       'Agregado jeaa 26/12/2005
        .BandIVA = True       'Agregado jeaa 03/08/2006
        .PorcentajeICE = 0
        .Tipo = 0
        .CantRelUnidad = 1
        .CantRelUnidadCont = 1
        
        .CantReluni = 1
        .CantReluniCont = 1
        
        .CantReuni = 1
        .CantReuniCont = 1
        
        .BandRepGastos = False
        .PorcentajeIVA = gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA
        .PorcentajeIVAAnt = gobjMain.EmpresaActual.GNOpcion.PorcentajeIVAAnt
        .BandNoSeFactura = False
        .BandModPrecio = False
        .BandAltaRotacion = False
        .BandNoObjetoIVA = False
        .BandOmitirRebate = False
        .Buffer = 0
        .M3 = 0
    End With
    Set mcolDetalleProveedor = New Collection
    Set mcolDetalleFamilia = New Collection
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mcolDetalles = Nothing
    Set mProps.objEmpresa = Nothing
End Sub

'Diego Prod
'--------------- Manejo de coleccion de DetalleFamilia
Public Function AddDetalleFamilia() As Long
    Dim nuevo As IVFamiliaDetalle
    
    If mcolDetalleFamilia Is Nothing Then CargaDetalleFamilia
    Set nuevo = New IVFamiliaDetalle
    Set nuevo.IVInventario = Me  'Me  olvide  del padre
    'Proporciona valores predeterminados
    nuevo.Cantidad = 1
    If nuevo.IVInventario.Tipo = rubro Then
        nuevo.TarifaJornal = 1
        nuevo.Rendimiento = 1
    End If
    
    mcolDetalleFamilia.Add nuevo
    AddDetalleFamilia = mcolDetalleFamilia.Count    'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function

Public Sub RemoveDetalleFamilia(ByVal ix As Long, Optional obj As IVFamiliaDetalle)
    If mcolDetalleFamilia Is Nothing Then CargaDetalleFamilia

    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolDetalleFamilia.Count
            If mcolDetalleFamilia.item(ix) Is obj Then
                mcolDetalleFamilia.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolDetalleFamilia.Remove ix
    End If

    'mcolDetalleFamilia.Remove ix
    mbooModificado = True
End Sub
''''Public Sub RemoveDetalleFamilia(ByVal ix As Long, Optional obj As IVFamiliaDetalle)
''''    If mProps.Tipo = 0 Then
''''        If mcolDetalleFamilia Is Nothing Then CargaDetalleFamilia
''''      Si está especificado un objeto
''''        If Not (obj Is Nothing) Then
''''            Busca el indice del objeto
''''            For ix = 1 To mcolDetalleFamilia.Count
''''                If mcolDetalleFamilia.item(ix) Is obj Then
''''                    mcolDetalleFamilia.Remove ix
''''                    Exit For
''''                End If
''''            Next ix
''''            Debug.Assert True     'No se puede dar
''''        Else
''''            mcolDetalleFamilia.Remove ix
''''        End If
''''    Else
''''        If mcolDetalleFamilia1 Is Nothing Then CargaDetalleFamilia1
''''        Si está especificado un objeto
''''        If Not (obj Is Nothing) Then
''''            Busca el indice del objeto
''''            For ix = 1 To mcolDetalleFamilia1.Count
''''                If mcolDetalleFamilia1.item(ix) Is obj Then
''''                    mcolDetalleFamilia1.Remove ix
''''                    Exit For
''''                End If
''''            Next ix
''''            Debug.Assert True     'No se puede dar
''''        Else
''''            mcolDetalleFamilia1.Remove ix
''''        End If
''''    End If
''''    mbooModificado = True
''''End Sub

Public Sub RemoveDetalleProveedor(ByVal ix As Long, Optional obj As IVDetalleProveedor)
    If mcolDetalleProveedor Is Nothing Then CargaDetalleProveedor
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolDetalleProveedor.Count
            If mcolDetalleProveedor.item(ix) Is obj Then
                mcolDetalleProveedor.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolDetalleProveedor.Remove ix
    End If
    'mcolDetalleFamilia.Remove ix
    mbooModificado = True
End Sub

Private Function verificacionDuiplicado(ByVal IdInventario As String) As Boolean
    Dim rs As Recordset, sql As String
    verificacionDuiplicado = False
    'Recupera los detalles de existencia
    sql = "SELECT IdInventario FROM ivmateria where idinventario =" & IdInventario
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        verificacionDuiplicado = True
    End If
    Set rs = Nothing
End Function
'AUC 25/11/05
Private Function verificacionDuiplicadoProv(ByVal idProveedor As String) As Boolean
    Dim rs As Recordset, sql As String
    verificacionDuiplicadoProv = False
    'Recupera los detalles de existencia
    sql = "SELECT Idproveedor FROM ivdetalleProv where idproveedor=" & idProveedor
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        verificacionDuiplicadoProv = True
    End If
    Set rs = Nothing
End Function

'jeaa 16-03-04
'Para grabar el rgistro en ivexist y aparesca en lista de items para factura pese a que su cantidad es 0
Public Sub AseguraMovimientoExistencia()
    Dim CodBodega As String
    Dim sql As String, rs As Recordset, IdBodega As Integer, idItem As Long
    'recupera la bodega predeterminada
    CodBodega = mProps.objEmpresa.GNOpcion.ObtenerValor("BodegaPredeterminada")
    ' si no tiene bodega predeterminda sale sin grabar en ivexist
    If Len(CodBodega) = 0 Then Exit Sub
    sql = "SELECT IdBodega FROM IVBodega where codbodega='" & CodBodega & "'"
    'recupera idBodega
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    IdBodega = rs.Fields("IdBodega")
    'recupera iditem
    sql = "SELECT IdInventario FROM IVinventario where codinventario='" & mProps.CodInventario & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    idItem = rs.Fields("IdInventario")
    sql = "INSERT INTO IVExist (IdInventario,IdBodega,Exist) " & _
              "VALUES(" & idItem & "," & IdBodega & "," & 0 & ")"
    mProps.objEmpresa.OpenRecordset sql
    'jeaa 20/11/2006
    sql = "INSERT INTO IVCompr (IdInventario,IdBodega,Compr) " & _
              "VALUES(" & idItem & "," & IdBodega & "," & 0 & ")"
    mProps.objEmpresa.OpenRecordset sql
    
    rs.Close
    Set rs = Nothing
End Sub

'***13/Abril/2004. Angel Pinduisaca
'***Genera Codigo en base a los grupos seleccionados y segun configuraciòn
'***en información de la empresa
Private Function GeneraCodigoEncadenado() As String
    Dim sql As String, num As Long, cod As String, rs As Recordset
    Dim pre As String, codg As String, v1 As Variant, v2 As Variant
    Dim Orden As String, size As String, mascara As String, i As Integer
    
    
    '***Aqui hay formar la parte de los grupos
    'unir con esteban
    Orden = Me.Empresa.GNOpcion.ObtenerValor("GenerarIVCodigoOrden")
    size = Me.Empresa.GNOpcion.ObtenerValor("GenerarIVCodigotamano")
    mascara = Me.Empresa.GNOpcion.ObtenerValor("GenerarIVCodigoMascara")
    
    codg = ""
    v1 = Split(Orden, ",")
    v2 = Split(size, ",")
    For i = LBound(v1, 1) To UBound(v1, 1)
        codg = codg & Trim$(Mid$(mProps.CodGrupo(v1(i)), 1, v2(i)))
    Next i
    
    pre = Me.Empresa.GNOpcion.PrefijoCodItem
    pre = Trim$(pre) & Trim$(codg)
    cod = Me.CodInventario
#If DAOLIB Then
    sql = "SELECT Max(VAL(CodInventario)) AS CodMax FROM IVInventario " & _
          "WHERE CodInventario LIKE '" & pre & "*'"
#Else
    sql = "SELECT (MAX(SUBSTRING(CodInventario," & Len(pre) + 1 & ",LEN(codinventario)))) As CodMax " & _
          "FROM IVInventario " & _
          "WHERE CodInventario LIKE '" & pre & "[0-9][0-9]%'"
#End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        If Not IsNull(rs!CodMax) Then num = val(rs!CodMax)
    End If
    rs.Close
    num = num + 1
    cod = pre & Format(num, mascara) 'cambiar por plantilla
    
    GeneraCodigoEncadenado = cod
    Set rs = Nothing
End Function





'***  Agregado Oliver para sacar el costo promedio de solo los ingresos dentro de un periodo
'***  calcula el costo promedio de todos los ingresos en una bodega especifica

Public Function CostoPromXPeriodo( _
                    ByVal FechaDesde As Date, _
                    ByVal FechaHasta As Date, _
                    ByVal CodBodega As String) As Double
    Dim sql As String, num As Integer, campo As String
    Dim rs As Recordset, Costo As Double        '*** MAKOTO 23/nov/00 Modificado a Double
    
    num = mProps.objEmpresa.GNOpcion.IndiceMoneda(Me.CodMoneda)
    
    sql = "spIVCostoPromXPeriodo " & _
                mProps.IdInventario & ", " & _
                "'" & CodBodega & "' , " & _
                FechaYMD(FechaDesde, Me.Empresa.TipoDB) & ", " & _
                FechaYMD(FechaHasta, Me.Empresa.TipoDB) & ", " & _
                num
    
    Set rs = Me.Empresa.OpenRecordset(sql)
    Costo = rs.Fields("CostoProm")
    
    
    rs.Close
    Set rs = Nothing

    'Si no hay costo promedio por que no hay existencia,
    '   devuelve el último costo en vez de promedio
    '   Pendiente de REVISAR
    '    If Costo = 0 Then Costo = CostoUltimo(Fecha, tid, Hora, False)

    CostoPromXPeriodo = Costo
End Function

'***************   Oliver  29 jun 2004
Public Property Let Descuento(ByVal ix As Integer, ByVal value As Currency)
    mProps.Descuento(ix) = value
    mbooModificado = True
End Property

Public Property Get Descuento(ByVal ix As Integer) As Currency
    Descuento = mProps.Descuento(ix)
End Property

'jeaa 18/11/2005
Friend Sub RecuperarParaCostos(key As Variant)
    Dim sql As String, i As Integer
    Dim rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM vwIVInventarioRecuperar WHERE CodInventario='" & key & "'"
    Else
        sql = "SELECT * FROM vwIVInventarioRecuperar WHERE IdInventario=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IdInventario = !IdInventario
        mProps.CodInventario = !CodInventario
        If Not IsNull(!CodMoneda) Then mProps.CodMoneda = !CodMoneda
        
'        For i = 1 To 4
'            If Not IsNull(.Fields("Precio" & i)) Then mProps.Precio(i) = .Fields("Precio" & i)
'            If Not IsNull(.Fields("Comision" & i)) Then mProps.Comision(i) = .Fields("Comision" & i)
'            If Not IsNull(.Fields("CantLimite" & i)) Then mProps.CantLimite(i) = .Fields("CantLimite" & i)
'            If Not IsNull(.Fields("Descuento" & i)) Then mProps.Descuento(i) = .Fields("Descuento" & i)   '***Agregado. jeaa. 29/jul/2004
'        Next i
'        If Not IsNull(!Unidad) Then mProps.Unidad = !Unidad
        
       
'        If Not IsNull(!ExistenciaMinima) Then mProps.ExistenciaMinima = !ExistenciaMinima
'        If Not IsNull(!ExistenciaMaxima) Then mProps.ExistenciaMaxima = !ExistenciaMaxima
'        If Not IsNull(!UnidadMinimaCompra) Then mProps.UnidadMinimaCompra = !UnidadMinimaCompra
'        If Not IsNull(!UnidadMinimaCompra) Then mProps.UnidadMinimaVenta = !UnidadMinimaCompra
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!BandServicio) Then mProps.BandServicio = !BandServicio
        If Not IsNull(!Tipo) Then mProps.Tipo = !Tipo
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub
'AUC agrega detalle proveedor
Public Function AddDetalleProveedor() As Long
    Dim nuevo As IVDetalleProveedor
    
    If mcolDetalleProveedor Is Nothing Then CargaDetalleProveedor
    
    Set nuevo = New IVDetalleProveedor
    Set nuevo.IVInventario = Me
    mcolDetalleProveedor.Add nuevo
    AddDetalleProveedor = mcolDetalleProveedor.Count      'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function

'AUC 23/11/2005
Private Sub CargaDetalleProveedor()
    Dim rs As Recordset, sql As String, obj As IVDetalleProveedor
    Dim i As Long
    Set mcolDetalleProveedor = New Collection
    sql = "Select pc.Idprovcli, pc.Codprovcli, pc.nombre,pc.direccion1, pc.telefono1, pc.telefono2, pc.ciudad,  " & _
            " ivp.cantidad, ivp.costounitario, ivp.fecha, ivp.BandPrincipal,ivp.preciocompra,  ivp.idinventarioprov, ivp.CodinventarioProv  From pcprovcli pc " & _
           " inner join  ivproveedordetalle ivp on ivp.idproveedor=pc.idprovcli " & _
           " Where pc.BandProveedor = 1 And ivp.IdInventario = " & Me.IdInventario
    sql = sql & " order by ivp.fecha"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddDetalleProveedor
            Set obj = mcolDetalleProveedor.item(i)
'            obj.IdProvCli = !IdProvCli
            obj.CodProveedor = !CodProvcli 'SI VALE NO BORRAR
    
            If Not IsNull(!Nombre) Then obj.Nombre = !Nombre
           If Not IsNull(!Direccion1) Then obj.Direccion = !Direccion1
            If Not IsNull(!Telefono1) Then obj.Telefono1 = !Telefono1
            If Not IsNull(!Telefono2) Then obj.Telefono2 = !Telefono2
            If Not IsNull(!Ciudad) Then obj.Ciudad = !Ciudad
            If Not IsNull(!Cantidad) Then obj.Cantidad = !Cantidad
            If Not IsNull(!CostoUnitario) Then obj.CostoUnitario = !CostoUnitario
            If Not IsNull(!fecha) Then obj.fecha = !fecha
            If Not IsNull(!BandPrincipal) Then obj.BandPrincipal = !BandPrincipal
            If Not IsNull(!idInventarioProv) Then obj.idInventarioProv = !idInventarioProv
            If Not IsNull(!CodInventarioProv) Then obj.CodInventarioProv = !CodInventarioProv
            obj.PrecioCompra = !PrecioCompra
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
Public Property Get CountProveedores() As Long
    CountProveedores = mcolDetalleProveedor.Count
End Property

Public Property Get proveedores(ix As Integer) As IVDetalleProveedor
    Set proveedores = mcolDetalleProveedor.item(ix)
End Property

Public Property Let CodUnidad(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodUnidad Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdUnidad FROM ivUnidad WHERE CodUnidad='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVUnidad_CodUnidad", MSGERR_NOHAYCODIGO
                Exit Property
            Else  'ESTOY AQUI
                IdUnidad = rs!IdUnidad
            End If
            rs.Close
        Else
            IdUnidad = 0
        End If
        mProps.CodUnidad = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property
Public Property Get CodUnidad() As String
    CodUnidad = Trim$(mProps.CodUnidad)
End Property
Public Property Let CodUnidadConteo(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodUnidadConteo Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdUnidad FROM ivUnidad WHERE CodUnidad='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVUnidad_CodUnidadConteo", MSGERR_NOHAYCODIGO
                Exit Property
            Else  'ESTOY AQUI
                IdUnidadConteo = rs!IdUnidad
            End If
            rs.Close
        Else
            IdUnidadConteo = 0
        End If
        mProps.CodUnidadConteo = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property
Public Property Get CodUnidadConteo() As String
    CodUnidadConteo = Trim$(mProps.CodUnidadConteo)
End Property
'AUC 08/12/05 'para unidades conteo
Public Property Let IdUnidadConteo(ByVal value As Long)
    mProps.IdUnidadConteo = value
    mbooModificado = True
End Property

Public Property Get IdUnidad() As Long
    IdUnidad = mProps.IdUnidad
End Property

Public Property Get IdUnidadConteo() As Long
    IdUnidadConteo = mProps.IdUnidadConteo
End Property

'AUC 08/12/05 'para unidades
Public Property Let IdUnidad(ByVal value As Long)
    mProps.IdUnidad = value
    mbooModificado = True
End Property

'AUC 16/12/05 para mostrar en las recetas
Public Function MostrarUltimoCosto(ByRef cod As String) As Currency
    Dim rs As Recordset, sql As String
    sql = "Select Iv.CostoUltimoingreso  From   vwIvinventariorecuperar iv where iv.codinventario = '" & cod & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
     If rs.RecordCount = 0 Or IsNull(rs!CostoUltimoIngreso) Then
        MostrarUltimoCosto = 0
      Else
        MostrarUltimoCosto = !CostoUltimoIngreso
        Exit Function
        .Close
      End If
    End With
    Set rs = Nothing
End Function
'AUC
Private Function CostoUltimoReceta() As Double
    Dim sql As String, campo As String, num As Integer
    Dim rs As Recordset, cadena_hora As String
    num = Me.Empresa.GNOpcion.IndiceMoneda(Me.CodMoneda)
    campo = "1"
    If num > 1 Then campo = "gc.Cotizacion" & num
    'Obtiene IVKardex ordenado por la fecha descendiente para coger último ingreso
    ' Busca solo de FechaTrans<=Fecha, y solo de INGRESOS y no anulados
    sql = "SELECT ivk.CostoRealTotal/ivk.Cantidad/" & campo & " AS CU " & _
          "FROM GNTrans gt INNER JOIN " & _
                    "(GNComprobante gc INNER JOIN IVKardex ivk " & _
                    "ON gc.TransID = ivk.TransID) " & _
                "ON gt.CodTrans = gc.CodTrans "
    sql = sql & "WHERE (gc.Estado <> 3) AND " & _
                    "(gt.AfectaCantidad=" & CadenaBool(True, Me.Empresa.TipoDB) & ") AND " & _
                    "(ivk.IdInventario=" & mProps.IdInventario & ") "
    'If BandSoloIngreso Then sql = sql & " AND (ivk.Cantidad>0) "
    sql = sql & " ORDER BY gc.FechaTrans DESC, gc.HoraTrans DESC, gc.TransID DESC"     '*** MAKOTO 08/dic/00 HoraTrans Agregado
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'Coge el primer costo como el costo ultimo
    If Not rs.EOF Then
        If Not IsNull(rs.Fields("CU")) Then
            CostoUltimoReceta = rs.Fields("CU")
        '*** Oliver 10 sep 2004
        'Aseguro de devolver 0 porque no encontro mov. de item
        Else
            CostoUltimoReceta = 0
        End If
    Else
        CostoUltimoReceta = 0
    End If
    rs.Close
    Set rs = Nothing
End Function


'agregado esteban 15/09/2005
Public Property Let BandIVA(ByVal value As Boolean)
    mProps.BandIVA = value
    mbooModificado = True
End Property

Public Property Get BandIVA() As Boolean
    BandIVA = mProps.BandIVA
End Property

Public Property Let PorcentajeICE(ByVal value As Currency)
    If value < 0 Or value > 1 Then
        Err.Raise ERR_INVALIDO, "IVInventario_PorcentajeICE", MSGERR_INVALIDO
    End If
    
    If mProps.PorcentajeICE <> value Then
        mProps.PorcentajeICE = value
        mbooModificado = True
    End If
End Property

Public Property Get PorcentajeICE() As Currency
    PorcentajeICE = mProps.PorcentajeICE
End Property

Public Function ExistenciaPreparacion(ByRef idInv As Long) As Currency
    Dim rs As Recordset
    Dim sql As String
    sql = "select exist from ivexist inner join ivinventario iv " & _
           "on iv.idinventario = ivexist.idinventario " & _
           "inner join ivbodega ivb on ivb.idbodega = ivexist.idbodega " & _
           " Where iv.IdInventario = " & idInv
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not IsNull(rs) Then
                ExistenciaPreparacion = rs!exist
                Exit Function
           End If
   Set rs = Nothing
End Function

'/*********** COSTO PARA PREPARACIONES *****************'
'*** MAKOTO 08/dic/00 auc
'Calcula y devuelve el costo del item
'Fehca: Fecha limite para el calculo.
'cant : Es necesario solo cuando es FIFO o LIFO
'       Indica qué cantidad va a coger para calcular el costo
'tid  : TransID para calcular hasta antes de esa transaccion
'           cuando existe varias trans. en la misma fecha
'Hora : Hora límite para el calculo.
Public Function CostoDouble2Prep( _
                    ByVal fecha As Date, _
                    Optional ByVal cant As Currency, _
                    Optional ByVal tid As Long, _
                    Optional ByVal Hora As Date, _
                    Optional ByVal idInv As Long) As Double
    'Si la id de transaccion no está especificada, cambia la fecha limite
    '  para que incluya todas las trans. de la misma fecha
    If Hora = 0 Then Hora = TimeSerial(23, 59, 59)
    Select Case mProps.objEmpresa.GNOpcion.Costeo
    Case COSTO_PROM
        CostoDouble2Prep = CostoPromPrep(fecha, tid, Hora, idInv)
    Case COSTO_FIFO
        CostoDouble2Prep = CostoFIFOPrep(fecha, cant, tid, Hora, idInv)
    Case COSTO_LIFO
        CostoDouble2Prep = CostoLIFOPrep(fecha, cant, tid, Hora, idInv)
    Case COSTO_ULTIMO
        CostoDouble2Prep = CostoUltimoPrep(fecha, tid, Hora, True, idInv)
    End Select
    '*** MAKOTO 31/ene/01 Agregado para redondear al formato de costo
    CostoDouble2Prep = MiCCur(Format$(CostoDouble2Prep, Me.Empresa.GNOpcion.ObtenerValor( _
                            "FormatoCosto" & Me.Empresa.GNOpcion.IndiceMoneda(Me.CodMoneda))))
End Function

Private Function CostoPromPrep( _
                    ByVal fecha As Date, _
                    ByVal tid As Long, _
                    ByVal Hora As Date, _
                    ByVal idInv As Long) As Double
    Dim sql As String, num As Integer, campo As String
    Dim rs As Recordset, Costo As Double        '*** MAKOTO 23/nov/00 Modificado a Double
    Dim idItem As Long
    num = mProps.objEmpresa.GNOpcion.IndiceMoneda(Me.CodMoneda)
    campo = "CP" & num
#If DAOLIB Then
    Dim qdef As QueryDef
    Set qdef = mProps.objEmpresa.Database.QueryDefs("spIVCostoProm")
    If Me.Tipo = Preparacion Then 'AUC agregado para sacar el costo de un item de preparacion
        qdef.Parameters("IdItem") = idInv
    Else
        qdef.Parameters("IdItem") = mProps.IdInventario
    End If
    qdef.Parameters("IdItem") = mProps.IdInventario
    qdef.Parameters("FechaHasta") = fecha
    qdef.Parameters("tid") = tid
    qdef.Parameters("HoraHasta") = Hora         '*** MAKOTO 09/dic/00 Agregado
    Set rs = qdef.OpenRecordset(dbOpenSnapshot, dbReadOnly)
    If Not rs.EOF Then
        If Not IsNull(rs.Fields(campo)) Then Costo = rs.Fields(campo)
    End If
#Else
    '*** MAKOTO 08/dic/00 Modificado
    If Me.Tipo = Preparacion Then 'AUC agregado para sacar el costo de un item de preparacion
           idItem = idInv
    Else
           idItem = mProps.IdInventario
    End If
    sql = "spIVCostoProm2 " & _
                idItem & ", " & _
                FechaYMD(fecha, Me.Empresa.TipoDB) & ", " & _
                tid & ", " & _
                num & ", " & _
                HoraHMS(Hora, Me.Empresa.TipoDB)
'                "'" & Format$(Hora, "Hh:Nn:Ss") & "'"  '*** MAKOTO 06/mar/01 Mod.
    Set rs = Me.Empresa.OpenRecordset(sql)
    Costo = rs.Fields("CostoProm")
#End If
    rs.Close
    Set rs = Nothing
    'Si no hay costo promedio por que no hay existencia,
    '   devuelve el último costo en vez de promedio
    If Costo = 0 Then Costo = CostoUltimoPrep(fecha, tid, Hora, False, idInv)
    CostoPromPrep = Costo
End Function

Private Function CostoFIFOPrep( _
                    ByVal fecha As Date, _
                    ByVal cant As Currency, _
                    ByVal tid As Long, _
                    ByVal Hora As Date, _
                    ByVal idInv As Long) As Double
    Dim sql As String, j As Currency, k As Currency, c As Double
    Dim campo As String, num As Integer
    Dim rs As Recordset
    Dim idItem As Long 'AUC agregado para saber cual id del item tomar
    num = Me.Empresa.GNOpcion.IndiceMoneda(Me.CodMoneda)
    campo = "C" & num
    If Me.Tipo = Preparacion Then 'AUC agregado para sacar el costo de un item de preparacion
           idItem = idInv
    Else
           idItem = mProps.IdInventario
    End If
    'Obtiene la existencia hasta la fecha
    sql = "SELECT Sum(Cantidad) AS SumaCant " & _
          "FROM GNTrans INNER JOIN (GNComprobante INNER JOIN IVKardex " & _
          "ON GNComprobante.TransID = IVKardex.TransID) ON GNTrans.CodTrans = GNComprobante.CodTrans " & _
          "WHERE (Estado <> 3) And (AfectaCantidad=" & CadenaBool(True, Me.Empresa.TipoDB) & ") " & _
          "AND (IdInventario=" & idItem & ") " 'AUC 09/01/06 cambiado antes mprops.idinventario
    'Condicion de fecha y tid
    '   Hasta antes de la fecha limite o la misma fecha pero con TransId menor
    sql = sql & "AND ((GNComprobante.FechaTrans<" & FechaYMD(fecha, Me.Empresa.TipoDB, True) & ") " & _
                  "OR (GNComprobante.FechaTrans=" & FechaYMD(fecha, Me.Empresa.TipoDB, True) & _
                  " AND (GNComprobante.TransID<" & tid & ")) )"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    j = 0
    If Not IsNull(rs!sumaCant) Then j = rs!sumaCant
    rs.Close
    'Si pide más de lo que existe
    If cant > j Then cant = j
    'Obtiene IVKardex ordenado por la fecha descendente
    sql = "SELECT Cantidad, CostoRealTotal/Cantidad AS C1, " & _
          "CostoRealTotal/Cotizacion2/Cantidad AS C2, " & _
          "CostoRealTotal/Cotizacion3/Cantidad AS C3, " & _
          "CostoRealTotal/Cotizacion4/Cantidad AS C4 " & _
          "FROM GNTrans INNER JOIN (GNComprobante INNER JOIN IVKardex " & _
          "ON GNComprobante.TransID = IVKardex.TransID) ON GNTrans.CodTrans = GNComprobante.CodTrans "
    sql = sql & "WHERE Estado<>3 And AfectaCantidad=" & CadenaBool(True, Me.Empresa.TipoDB) & _
            " AND IdInventario=" & idItem
    'Condicion de fecha y tid
    '   Hasta antes de la fecha limite o la misma fecha pero con TransId menor
    sql = sql & " AND ((GNComprobante.FechaTrans<" & FechaYMD(fecha, Me.Empresa.TipoDB, True) & ") " & _
                  "OR (GNComprobante.FechaTrans=" & FechaYMD(fecha, Me.Empresa.TipoDB, True) & _
                  " AND (GNComprobante.TransID<" & tid & ")) )"
'          " AND FechaTrans<=" & FechaYMD(Fecha, Me.Empresa.TipoDB)
    sql = sql & " ORDER BY FechaTrans DESC"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        If .EOF Then
            c = 0
        Else
            'Buscar el lote existente más antiguo
            Do Until .EOF
                If !Cantidad > 0 Then
                    j = j - !Cantidad
                    If j <= 0 Then Exit Do
                End If
                .MoveNext
            Loop
'            If not .EOF Then
            If .EOF Then .MoveLast
            j = 0
            If cant = 0 Then
                c = .Fields(campo)
            Else
                Do Until .BOF
                    If !Cantidad > 0 Then
                        If !Cantidad + j >= cant Then
                            k = cant - j
                            c = c + .Fields(campo) * k
                            Exit Do
                        Else
'                            k = !Cantidad + j
                            k = !Cantidad
                            j = j + !Cantidad
                            c = c + .Fields(campo) * k
                        End If
                    End If
                    .MovePrevious
                Loop
                c = c / cant
            End If
        End If
    End With
    rs.Close
    Set rs = Nothing
    CostoFIFOPrep = c
End Function

Private Function CostoLIFOPrep( _
                    ByVal fecha As Date, _
                    ByVal cant As Currency, _
                    ByVal tid As Long, _
                    ByVal Hora As Date, _
                    ByVal idInv As Long) As Double
    Dim sql As String, j As Currency, k As Currency, c As Double
    Dim campo As String, num As Integer
    Dim rs As Recordset
        Dim idItem As Long
    num = Me.Empresa.GNOpcion.IndiceMoneda(Me.CodMoneda)
    campo = "C" & num
    If Me.Tipo = Preparacion Then 'AUC agregado para sacar el costo de un item de preparacion
           idItem = idInv
    Else
           idItem = mProps.IdInventario
    End If
    'Obtiene IVKardex ordenado por la fecha descendente
    sql = "SELECT Cantidad,CostoRealTotal/Cantidad AS C1, " & _
          "CostoRealTotal/Cotizacion2/Cantidad AS C2, " & _
          "CostoRealTotal/Cotizacion3/Cantidad AS C3, " & _
          "CostoRealTotal/Cotizacion4/Cantidad AS C4 " & _
          "FROM GNTrans INNER JOIN (GNComprobante INNER JOIN IVKardex " & _
          "ON GNComprobante.TransID = IVKardex.TransID) ON GNTrans.CodTrans = GNComprobante.CodTrans "
        sql = sql & "WHERE (Estado <> 3) And (AfectaCantidad=" & CadenaBool(True, Me.Empresa.TipoDB) & ") AND (IdInventario=" & idItem & ")"
    'Condicion de fecha y tid
    '   Hasta antes de la fecha limite o la misma fecha pero con TransId menor
    sql = sql & " AND ((GNComprobante.FechaTrans<" & FechaYMD(fecha, Me.Empresa.TipoDB, True) & ") " & _
                  "OR (GNComprobante.FechaTrans=" & FechaYMD(fecha, Me.Empresa.TipoDB, True) & _
                  " AND (GNComprobante.TransID<" & tid & ")) )"
    sql = sql & " ORDER BY FechaTrans DESC"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        If rs.EOF Then
            c = 0
        Else
            If cant = 0 Then
                c = .Fields(campo)
            Else
                Do Until .BOF
                    If !Cantidad > 0 Then
                        If !Cantidad + j >= cant Then
                            c = c + .Fields(campo) * (cant - j)
                            Exit Do
                        Else
                            j = j + !Cantidad
                            c = c + .Fields(campo) * (!Cantidad + j)
                        End If
                    End If
                    .MovePrevious
                Loop
                c = c / cant
            End If
        End If
    End With
    rs.Close
    Set rs = Nothing
    CostoLIFOPrep = c
End Function

Private Function CostoUltimoPrep( _
                    ByVal fecha As Date, _
                    ByVal tid As Long, _
                    ByVal Hora As Date, _
                    ByVal BandSoloIngreso As Boolean, _
                    ByVal idInv As Long) As Double
    Dim sql As String, campo As String, num As Integer
    Dim rs As Recordset, cadena_hora As String
    Dim idItem As Long
    num = Me.Empresa.GNOpcion.IndiceMoneda(Me.CodMoneda)
    campo = "1"
    If num > 1 Then campo = "gc.Cotizacion" & num
        If Me.Tipo = Preparacion Then 'AUC agregado para sacar el costo de un item de preparacion
           idItem = idInv
    Else
           idItem = mProps.IdInventario
    End If
    'Obtiene IVKardex ordenado por la fecha descendiente para coger último ingreso
    ' Busca solo de FechaTrans<=Fecha, y solo de INGRESOS y no anulados
    sql = "SELECT ivk.CostoRealTotal/ivk.Cantidad/" & campo & " AS CU " & _
          "FROM GNTrans gt INNER JOIN " & _
                    "(GNComprobante gc INNER JOIN IVKardex ivk " & _
                    "ON gc.TransID = ivk.TransID) " & _
                "ON gt.CodTrans = gc.CodTrans "
    sql = sql & "WHERE (gc.Estado <> 3) AND " & _
                    "(gt.AfectaCantidad=" & CadenaBool(True, Me.Empresa.TipoDB) & ") AND " & _
                    "(ivk.IdInventario=" & idItem & ") "
    If BandSoloIngreso Then sql = sql & " AND (ivk.Cantidad>0) "
    'Condicion de fecha y tid
    '   Hasta antes de la fecha limite
    '   o la misma fecha pero la hora menor
    '   o la misma fecha y misma hora pero tid menor
'    cadena_hora = Format(Hora, "Hh:Nn:Ss")
'#If DAOLIB Then
'    cadena_hora = "#" & cadena_hora & "#"
'#Else
'    cadena_hora = "'" & cadena_hora & "'"
'#End If
    cadena_hora = HoraHMS(Hora, Me.Empresa.TipoDB)        '*** MAKOTO 06/mar/01 Mod.
    sql = sql & " AND ( (gc.FechaTrans<" & FechaYMD(fecha, Me.Empresa.TipoDB) & ") " & _
                    "OR (gc.FechaTrans=" & FechaYMD(fecha, Me.Empresa.TipoDB) & _
                        " AND gc.HoraTrans<" & cadena_hora & ") " & _
                    "OR (gc.FechaTrans=" & FechaYMD(fecha, Me.Empresa.TipoDB) & _
                        " AND gc.HoraTrans=" & cadena_hora & _
                        " AND gc.TransID < " & tid & ") )"
    sql = sql & " ORDER BY gc.FechaTrans DESC, gc.HoraTrans DESC, gc.TransID DESC"     '*** MAKOTO 08/dic/00 HoraTrans Agregado
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'Coge el primer costo como el costo ultimo
    If Not rs.EOF Then
        If Not IsNull(rs.Fields("CU")) Then
            CostoUltimoPrep = rs.Fields("CU")
        '*** Oliver 10 sep 2004
        'Aseguro de devolver 0 porque no encontro mov. de item
        Else
            CostoUltimoPrep = 0
        End If
    Else
        CostoUltimoPrep = 0
    End If
    rs.Close
    Set rs = Nothing
End Function

Public Property Let PorDesperdicio(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario.PorDesperdicio", _
            "El PorDesperdicio no puede ser negativo. ('" & mProps.CodInventario & "')"
    End If
    'Si es que realmente modifica el valor
    If mProps.PorDesperdicio <> value Then
        mProps.PorDesperdicio = value
       mbooModificado = True
    End If
End Property

Public Property Get PorDesperdicio() As Currency
    PorDesperdicio = mProps.PorDesperdicio
End Property

Public Property Let TipoSRI(ByVal value As SRITipo)
    mProps.TipoSRI = value
    mbooModificado = True
End Property

Public Property Get TipoSRI() As SRITipo
    TipoSRI = mProps.TipoSRI
End Property

'AUC 08/12/05 'para unidades conteo
Public Property Let IdICE(ByVal value As Long)
    mProps.IdICE = value
    mbooModificado = True
End Property
Public Property Get IdICE() As Long
    IdICE = mProps.IdICE
End Property

Public Property Let CodICE(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodICE Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdICE, porcentaje FROM IVRecargoICE WHERE CodICE='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVICE_CodICE", MSGERR_NOHAYCODIGO
                Exit Property
            Else  'ESTOY AQUI
                IdICE = rs!IdICE
                PorcentajeICE = rs!Porcentaje
            End If
            rs.Close
        Else
            IdICE = 0
        End If
        mProps.CodICE = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property

Public Property Get CodICE() As String
    CodICE = Trim$(mProps.CodICE)
End Property

Public Property Let CantRelUnidad(ByVal value As Long)
    mProps.CantRelUnidad = value
    mbooModificado = True
End Property

Public Property Get CantRelUnidad() As Long
    CantRelUnidad = mProps.CantRelUnidad
End Property

Public Property Let CantRelUnidadCont(ByVal value As Long)
    mProps.CantRelUnidadCont = value
    mbooModificado = True
End Property

Public Property Get CantRelUnidadCont() As Long
    CantRelUnidadCont = mProps.CantRelUnidadCont
End Property

Public Function RecuperaExistenciaxBodega() As Recordset
    Dim sql As String, rs As Recordset
    'sql = "select * from ivinventario inner join ivexist on ivinventario.idinventario= ivinventario.idinventario where idinventario=' " & mobjItem.CodInventario
    sql = "select *, codgrupo4 as codgrupo,  ivbodega.descripcion as descbodega, gns.codsucursal,ivg.idealdias from ivexist "
    sql = sql & " inner join ivbodega "
    sql = sql & " left join gnsucursal gns on ivbodega.idsucursal = gns.idsucursal"
    sql = sql & " on  ivexist.idbodega=ivbodega.idbodega "
    sql = sql & " left join ivgrupo4 ivg on ivexist.idgrupo=ivg.idgrupo4"
    sql = sql & " where idinventario= " & Me.IdInventario
    sql = sql & " AND IVBODEGA.BANDVALIDA=1"
    sql = sql & " order by GNS.IDSUCURSAL, codbodega "
    
    Set rs = Me.Empresa.OpenRecordset(sql)
    With rs
     If rs.RecordCount <> 0 Then
        Set RecuperaExistenciaxBodega = rs
        Exit Function
        .Close
      End If
    End With
    Set rs = Nothing
End Function

Public Sub GrabaExistenciaxBodega(ByVal IdInventario As Integer, ByVal IdBodega As Integer, ByVal min As Integer, ByVal max As Integer)
End Sub


Public Sub GrabaExistenciaxBodegaS(ByVal IdInventario As Long, ByVal IdBodega As Integer, ByVal min As Integer, ByVal max As Integer)
    Dim sql As String, rs As Recordset

    sql = "update ivexist "
    sql = sql & " set existmin=" & min & ", existmax=" & max
    sql = sql & " from ivexist"
    sql = sql & " where idinventario= " & IdInventario
    sql = sql & " and idbodega=" & IdBodega
    Set rs = Me.Empresa.OpenRecordset(sql)
    Set rs = Nothing
End Sub

Public Property Let Descripcion2(ByVal value As String)
    If Len(value) > 120 Then
        Err.Raise ERR_INVALIDO, "IVInventario_Descripcion2", MSGERR_INVALIDO
    End If
    
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, ";") > 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario_Descripcion2", _
            "El caracter ';' no es válido para la descripción."
    End If
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, "|") > 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario_Descripcion2", _
            "El caracter '|' no es válido para la descripción."
    End If
    
    mProps.Descripcion2 = value
    mbooModificado = True
End Property

Public Property Get Descripcion2() As String
    Descripcion2 = Trim$(mProps.Descripcion2)
End Property

Public Property Let PesoNeto(ByVal value As Currency)
    If mProps.PesoNeto <> value Then
        mProps.PesoNeto = value
        mbooModificado = True
    End If
End Property

Public Property Get PesoNeto() As Currency
    PesoNeto = mProps.PesoNeto
End Property

Public Property Let PesoBruto(ByVal value As Currency)
    If mProps.PesoBruto <> value Then
        mProps.PesoBruto = value
        mbooModificado = True
    End If
End Property

Public Property Get PesoBruto() As Currency
    PesoBruto = mProps.PesoBruto
End Property

Public Property Let IdUnidadPeso(ByVal value As Long)
    mProps.IdUnidadPeso = value
    mbooModificado = True
End Property

Public Property Get IdUnidadPeso() As Long
    IdUnidadPeso = mProps.IdUnidadPeso
End Property

Public Property Let CodUnidadPeso(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodUnidadPeso Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdUnidad, CodUnidad FROM IVUnidad WHERE CodUnidad='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVUnidad_CodUnidad", MSGERR_NOHAYCODIGO
                Exit Property
            Else  'ESTOY AQUI
                IdUnidadPeso = rs!IdUnidad
            End If
            rs.Close
        Else
            IdUnidadPeso = 0
        End If
        mProps.CodUnidadPeso = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property

Public Property Get CodUnidadPeso() As String
    CodUnidadPeso = Trim$(mProps.CodUnidadPeso)
End Property

Public Property Let BandConversion(ByVal value As Boolean)
    mProps.BandConversion = value
    mbooModificado = True
End Property

Public Property Get BandConversion() As Boolean
    BandConversion = mProps.BandConversion
End Property

Public Sub GrabaFamiliaDetalleCopia(idInv As Long, idMateria As Long, _
                                                            cant As Currency, bandPri As BookmarkEnum, _
                                                            BandMod As Boolean, cuanto As Currency, _
                                                            tarifa As Currency, fila As Long)
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As IVFamiliaDetalle, sql As String, band As Boolean
    'Elimina  todos  los registros
    sql = "Delete Ivm " & _
          "From Ivinventario Iv " & _
          "INNER JOIN IVMateria Ivm ON IV.IdInventario = Ivm.IdMateria "
    'jeaa 28/10/2005
    If mProps.Tipo = CambioPresentacion Or mProps.Tipo = Preparacion Or Tipo = Promocion Then
          sql = sql & " Where Ivm.IdInventario = " & mProps.IdInventario
   Else
          sql = sql & " Where Ivm.Idmateria = " & mProps.IdInventario
   End If

    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    band = True
    If Not mcolDetalleFamilia Is Nothing Then
        sql = "SELECT * FROM IVMateria WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        i = 1
        Set obj = mcolDetalleFamilia.item(i)
        If band Then
            rs1.AddNew
            If Me.Tipo <> Familia Then
                rs1!IdInventario = idInv
                rs1!idMateria = idMateria
            Else
                rs1!IdInventario = idMateria
                rs1!idMateria = idInv
            End If
            rs1!Cantidad = cant
            rs1!BandPrincipal = bandPri
            rs1!BandModificar = BandMod

            If Me.Tipo = Promocion Then
                rs1!xCuanto = cuanto
                rs1!TarifaJornal = tarifa
            End If
            rs1.Update
        End If
        band = True
    End If
    Set rs1 = Nothing
    Set obj = Nothing
End Sub

Public Function RecuperaIdInventario(cod As String) As Long
    Dim sql As String, rs As Recordset

    sql = " select idinventario from ivinventario where codinventario='" & cod & "'"
    Set rs = Me.Empresa.OpenRecordset(sql)
    With rs
     If rs.RecordCount <> 0 Then
        RecuperaIdInventario = rs.Fields("idinventario")
        Exit Function
        .Close
      End If
    End With
End Function

Public Property Let BandRepGastos(ByVal value As Boolean)
    mProps.BandRepGastos = value
    mbooModificado = True
End Property

Public Property Get BandRepGastos() As Boolean
    BandRepGastos = mProps.BandRepGastos
End Property

'AUC para insertar campo RutaFoto
Public Property Let RutaFoto(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "IVInventario_RutaFoto", MSGERR_INVALIDO & "Direccion muy larga"
    End If
    mProps.RutaFoto = value
    mbooModificado = True
End Property

Public Property Get RutaFoto() As String
    RutaFoto = Trim$(mProps.RutaFoto)
End Property

'recupera el numero de piezas del traje para la casa del disfraz
Public Function RecuperaNumPiezas(cod As String, CodTrans As String, NumTrans As Long) As Long
Dim sql As String
Dim rs As Recordset
sql = "select (count(idPadre)-1) as numP from ivkardex ivk  inner join ivinventario iv on ivk.idpadre =  iv.idinventario " & _
      "inner join gncomprobante gn on gn.transid = ivk.transid where gn.codtrans = '" & CodTrans & "'   AND  iv.codinventario = '" & cod & "' " & _
      " AND GN.NUMTRANS = " & NumTrans
 Set rs = Me.Empresa.OpenRecordset(sql)
    With rs
     If rs.RecordCount = 0 Then
        RecuperaNumPiezas = 0
      Else
         RecuperaNumPiezas = !nump
        Exit Function
        .Close
      End If
    End With
    Set rs = Nothing
End Function

Public Function CostoUltimoIngresoRubro(item As IVInventario) As Currency
Dim j As Long, item1 As IVInventario
Dim codsub As String
Dim Cantidad As Currency
Dim Rendimiento As Currency
Dim CostoUltimo As Currency
    If item.NumFamiliaDetalle > 0 Then
        For j = 1 To item.NumFamiliaDetalle 'ciclo para detalles de items
                codsub = item.RecuperaDetalleFamilia(j).CodInventario
                Cantidad = item.RecuperaDetalleFamilia(j).Cantidad
                Rendimiento = item.RecuperaDetalleFamilia(j).Rendimiento
                CostoUltimo = item.RecuperaDetalleFamilia(j).CostoUltimoIngreso
                CostoUltimoIngresoRubro = CostoUltimoIngresoRubro + (Cantidad * CostoUltimo * Rendimiento)
                Set item1 = Me.Empresa.RecuperaIVInventario(codsub)
                If item1.Tipo = rubro Then
                    CostoUltimoIngresoRubro item1 'Proceso recursivo
                End If
        Next j
    End If
End Function

Public Function CostoUltimoIngresoRubro1(item As IVInventario, idpadre As Long, cantPadre As Currency) As Currency
Dim rs As Recordset
Dim sql As String
Dim Cantidad As Currency
Dim Rendimiento As Currency
Dim CostoUltimo As Currency
On Error GoTo CapturaError
    sql = "SELECT IVM.CANTIDAD,IVM.RENDIMIENTO,IV.COSTOULTIMOINGRESO FROM IVMATERIA IVM " & _
          " INNER JOIN IVINVENTARIO IV ON IV.IDINVENTARIO = IVM.IDMATERIA  " & _
          " WHERE IVM.IDMATERIA =  " & item.IdInventario & " AND IVM.IDINVENTARIO = " & idpadre
    Set rs = Me.Empresa.OpenRecordset(sql)
    Do While Not rs.EOF
            Cantidad = rs!Cantidad
            Rendimiento = rs!Rendimiento
            CostoUltimo = rs!CostoUltimoIngreso
            rs.MoveNext
    Loop
    CostoUltimoIngresoRubro1 = (Cantidad * CostoUltimo * Rendimiento) * -1
   Set rs = Nothing
   Exit Function
CapturaError:
        MsgBox Err.Description
        Set rs = Nothing
        Exit Function
End Function

Public Property Let BandNoSeFactura(ByVal value As Boolean)
    mProps.BandNoSeFactura = value
    mbooModificado = True
End Property

Public Property Get BandNoSeFactura() As Boolean
    BandNoSeFactura = mProps.BandNoSeFactura
End Property


'*** MAKOTO 26/mar/01 Agregado
'Para que asigne bien automáticamente CodCuentaContable cuando crea el objeto
Private Sub AsignarCodCuentasPre()
    'Cuenta contable predeterminado
    If Len(Me.CodCuentaActivo) = 0 And Not (Me.Empresa Is Nothing) Then
            Me.CodCuentaActivo = Me.Empresa.GNOpcion.CodCuentaActivoPre
    End If
    If Len(Me.CodCuentaCosto) = 0 And Not (Me.Empresa Is Nothing) Then
            Me.CodCuentaCosto = Me.Empresa.GNOpcion.CodCuentaCostoPre
    End If
    If Len(Me.CodCuentaVenta) = 0 And Not (Me.Empresa Is Nothing) Then
            Me.CodCuentaVenta = Me.Empresa.GNOpcion.CodCuentaVentaPre
    End If
    If Len(Me.CodCuentaDiferida) = 0 And Not (Me.Empresa Is Nothing) Then
            Me.CodCuentaDiferida = Me.Empresa.GNOpcion.CodCuentaDiferidaPre
    End If
End Sub

'jeaa 24/04/2009
Public Property Let TiempoPromVta(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario.TiempoPromVta", _
            "El Recargo no puede ser negativo. ('" & mProps.CodInventario & "')"
    End If
    
    If mProps.TiempoPromVta <> value Then
        mProps.TiempoPromVta = value
        mbooModificado = True
    End If
End Property

Public Property Get TiempoPromVta() As Currency
    TiempoPromVta = mProps.TiempoPromVta
End Property

Public Property Let TiempoReposicion(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario.TiempoReposicion", _
            "El Tiempo Reposicion no puede ser negativo. ('" & mProps.CodInventario & "')"
    End If

    If mProps.TiempoReposicion <> value Then
        mProps.TiempoReposicion = value
        mbooModificado = True
    End If
End Property

Public Property Get TiempoReposicion() As Currency
    TiempoReposicion = mProps.TiempoReposicion
End Property


Public Function ListaPrecioxPCGrupoParaFlex(ByVal gc As GNComprobante, precios As String) As String
    Dim s As String, r As Double, fmt As String
    
    'fmt = gc.Empresa.GNOpcion.FormatoMoneda(gc.CodMoneda)
    'JEAA 24/01/05
    fmt = gc.Empresa.GNOpcion.ObtenerValor("FormatoPU" & gc.Empresa.GNOpcion.IndiceMoneda(MONEDA_SEC))
    
    '*** MAKOTO 28/sep/00 Agregado para corregir error en caso de '#,#'
    If Right$(fmt, 2) = ",#" Then fmt = fmt & "0"
    
    r = 1#
    
    'Si la transacción está en una moneda distinta de la del item
    If gc.CodMoneda <> mProps.CodMoneda Then
        'Calcula el factor para el calculo de equivalencia
        r = gc.Cotizacion(mProps.CodMoneda) / gc.Cotizacion(gc.CodMoneda)
    End If
    If Mid$(precios, 1, 1) = 1 Then
        s = "|" & Format(Precio(1) * r, fmt)
'    Else
'        s = s & Format(0, fmt)
    End If
    If Mid$(precios, 2, 1) = 1 Then
        s = s & "|" & Format(Precio(2) * r, fmt)
'     Else
'        s = s & "|" & Format(0, fmt)
    End If
   
    If Mid$(precios, 3, 1) = 1 Then
        s = s & "|" & Format(Precio(3) * r, fmt)
'    Else
'        s = s & "|" & Format(0, fmt)
    End If
        
    If Mid$(precios, 4, 1) = 1 Then
        s = s & "|" & Format(Precio(4) * r, fmt)
'    Else
'        s = s & "|" & Format(0, fmt)
    End If
        
    If Mid$(precios, 5, 1) = 1 Then
        s = s & "|" & Format(Precio(5) * r, fmt)
'    Else
'        s = s & "|" & Format(0, fmt)
    End If
        
    If Mid$(precios, 6, 1) = 1 Then 'AUC precio 6
        s = s & "|" & Format(Precio(6) * r, fmt)
    End If

    If Len(precios) > 6 Then
         If Mid$(precios, 7, 1) = 1 Then 'GRV Precio 7 - 14MAY2014
            s = s & "|" & Format(Precio(7) * r, fmt)
        End If
    End If
    
    ListaPrecioxPCGrupoParaFlex = s
End Function



Public Function RecuperaCadenaPreciosxGrupo(ByVal NumGrupo As Integer, IdGrupo As Long) As String
Dim rs As Recordset
Dim sql As String
On Error GoTo CapturaError
    sql = "SELECT PreciosDisponibles from pcgrupo" & NumGrupo
    sql = sql & " where idgrupo" & NumGrupo & "=" & IdGrupo
    Set rs = Me.Empresa.OpenRecordset(sql)
    
    With rs
     If rs.RecordCount = 0 Then
        RecuperaCadenaPreciosxGrupo = "00000"
      Else
         RecuperaCadenaPreciosxGrupo = !PreciosDisponibles
      End If
    End With
    Set rs = Nothing
    
   Exit Function
CapturaError:
        MsgBox Err.Description
        Set rs = Nothing
        Exit Function
End Function


Public Function ObtieneNumeroPrecioPadre(codItem As String, gc As GNComprobante) As Integer
    Dim iv As IVInventario, i As Integer, Precio As Currency, j As Long
    
    
    For j = 1 To gc.CountIVKardex
        Set iv = gc.Empresa.RecuperaIVInventario(codItem)
         If Not (iv Is Nothing) Then
            If gc.IVKardex(j).CodInventario = codItem Then
                For i = 1 To 5
                     If gc.IVKardex(j).Precio = iv.Precio(i) Then
                         ObtieneNumeroPrecioPadre = i
                         Exit For
                     End If
                 Next i
            End If
        End If
    Next j
        Set iv = Nothing
End Function




Friend Sub RecuperarALt(key As Variant)
    Dim sql As String, i As Integer
    Dim rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM vwIVInventarioRecuperar WHERE CodAlterno1 ='" & key & "'"
    Else
        sql = "SELECT * FROM vwIVInventarioRecuperar WHERE IdInventario=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IdInventario = !IdInventario
        mProps.CodInventario = !CodInventario
        If Not IsNull(!CodAlterno1) Then mProps.CodAlterno1 = !CodAlterno1
        If Not IsNull(!CodAlterno2) Then mProps.CodAlterno2 = !CodAlterno2
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
#If DAOLIB Then
        If Not IsNull(!DescripcionDetalle) Then mProps.DescripcionDetalle = !DescripcionDetalle
#Else
        If Not IsNull(!DescripcionDetalle) Then
            If .Fields("DescripcionDetalle").ActualSize > 0 Then
                mProps.DescripcionDetalle = .Fields("DescripcionDetalle").GetChunk(.Fields("DescripcionDetalle").ActualSize)
            End If
        End If
#End If
        If Not IsNull(!CodMoneda) Then mProps.CodMoneda = !CodMoneda
        
        'For i = 1 To 6 'AUC cambiado antes 4
        For i = 1 To 7 'GRV Precio 7 - 14MAY2014
            If Not IsNull(.Fields("Precio" & i)) Then mProps.Precio(i) = .Fields("Precio" & i)
            If Not IsNull(.Fields("Comision" & i)) Then mProps.Comision(i) = .Fields("Comision" & i)
            If Not IsNull(.Fields("CantLimite" & i)) Then mProps.CantLimite(i) = .Fields("CantLimite" & i)
            If Not IsNull(.Fields("Descuento" & i)) Then mProps.Descuento(i) = .Fields("Descuento" & i)   '***Agregado. jeaa. 29/jul/2004
        Next i
        If Not IsNull(!PorcentajeIVA) Then mProps.PorcentajeIVA = !PorcentajeIVA
        If Not IsNull(!PorcentajeIVAAnt) Then mProps.PorcentajeIVAAnt = !PorcentajeIVAAnt
        If Not IsNull(!IdCuentaActivo) Then mProps.IdCuentaActivo = !IdCuentaActivo
        If Not IsNull(!IdCuentaCosto) Then mProps.IdCuentaCosto = !IdCuentaCosto
        If Not IsNull(!IdCuentaVenta) Then mProps.IdCuentaVenta = !IdCuentaVenta
        If Not IsNull(!IdCuentaDiferida) Then mProps.IdCuentaDiferida = !IdCuentaDiferida
        If Not IsNull(!IdCuentaDevolucion) Then mProps.IdCuentaDevolucion = !IdCuentaDevolucion
       
        
        
        If Not IsNull(!CodCuentaActivo) Then mProps.CodCuentaActivo = !CodCuentaActivo
        If Not IsNull(!CodCuentaCosto) Then mProps.CodCuentaCosto = !CodCuentaCosto
        If Not IsNull(!CodCuentaVenta) Then mProps.CodCuentaVenta = !CodCuentaVenta
        If Not IsNull(!CodCuentaDiferida) Then mProps.CodCuentaDiferida = !CodCuentaDiferida
        If Not IsNull(!CodCuentaDevolucion) Then mProps.CodCuentaDevolucion = !CodCuentaDevolucion
'        If Not IsNull(!Unidad) Then mProps.Unidad = !Unidad
        
        For i = 1 To IVGRUPO_MAX
            If Not IsNull(.Fields("IdGrupo" & i)) Then mProps.IdGrupo(i) = .Fields("IdGrupo" & i)
            If Not IsNull(.Fields("CodGrupo" & i)) Then mProps.CodGrupo(i) = .Fields("CodGrupo" & i)
        Next i
        
        If Not IsNull(!idProveedor) Then mProps.idProveedor = !idProveedor

        CargaDetalleProveedor  'auc
        If Not IsNull(!CodProvcli) Then mProps.CodProveedor = !CodProvcli
        If Not IsNull(!Observacion) Then mProps.Observacion = !Observacion
        If Not IsNull(!ExistenciaMinima) Then mProps.ExistenciaMinima = !ExistenciaMinima
        If Not IsNull(!ExistenciaMaxima) Then mProps.ExistenciaMaxima = !ExistenciaMaxima
        If Not IsNull(!UnidadMinimaCompra) Then mProps.UnidadMinimaCompra = !UnidadMinimaCompra
        If Not IsNull(!UnidadMinimaCompra) Then mProps.UnidadMinimaVenta = !UnidadMinimaCompra
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!BandServicio) Then mProps.BandServicio = !BandServicio
        If Not IsNull(!Tipo) Then mProps.Tipo = !Tipo
        'DiegoProd  Ojo  aumentar  proceso  para  recuperar  detalle
        If Tipo <> Normal Then
            CargaDetalleFamilia 'Pensandolo bien recuperamos
        End If
        '***Agregado. Angel. 29/jul/2004
        If Not IsNull(!ValorRecargo) Then mProps.ValorRecargo = !ValorRecargo
        '*** Agregado esteban 08/04/2005
        If Not IsNull(!BandFraccion) Then mProps.BandFraccion = !BandFraccion
        '*** Agregado esteban 15/09/2005
        If Not IsNull(!BandArea) Then mProps.BandArea = !BandArea
        '*** Agregado esteban 26/12/2005
        If Not IsNull(!bandVenta) Then mProps.bandVenta = !bandVenta
         If Not IsNull(!CodUnidad) Then mProps.CodUnidad = !CodUnidad
        If Not IsNull(!CodUniConteo) Then mProps.CodUnidadConteo = !CodUniConteo
        If Not IsNull(!IdUnidad) Then mProps.IdUnidad = !IdUnidad
        If Not IsNull(!IdUnidadConteo) Then mProps.IdUnidadConteo = !IdUnidadConteo
        If Not IsNull(!CostoUltimoIngreso) Then mProps.CostoUltimoIngreso = !CostoUltimoIngreso 'AUC 22/12/05
        If Not IsNull(!BandIVA) Then mProps.BandIVA = !BandIVA  'jeaa 08/03/2006
        If Not IsNull(!PorcentajeICE) Then mProps.PorcentajeICE = !PorcentajeICE
        If Not IsNull(!PorDesperdicio) Then mProps.PorDesperdicio = !PorDesperdicio 'jeaa 01/02/2006
        If Not IsNull(!RutaFoto) Then mProps.RutaFoto = !RutaFoto  ' AUC 04/05/06
        If Not IsNull(!TipoSRI) Then mProps.TipoSRI = !TipoSRI
        If Not IsNull(!IdICE) Then mProps.IdICE = !IdICE
        If Not IsNull(!CodICE) Then mProps.CodICE = !CodICE
        If Not IsNull(!CantRelUnidad) Then mProps.CantRelUnidad = !CantRelUnidad 'jeaa 27/0/2007
        If Not IsNull(!CantRelUnidadCont) Then mProps.CantRelUnidadCont = !CantRelUnidadCont 'jeaa 27/0/2007
        
        If Not IsNull(!CantReluni) Then mProps.CantReluni = !CantReluni 'jeaa 27/0/2007
        If Not IsNull(!CantReluniCont) Then mProps.CantReluniCont = !CantReluniCont 'jeaa 27/0/2007
        
        If Not IsNull(!CantReuni) Then mProps.CantReuni = !CantReuni 'jeaa 27/0/2007
        If Not IsNull(!CantReuniCont) Then mProps.CantReuniCont = !CantReuniCont 'jeaa 27/0/2007
        
        
        If Not IsNull(!Descripcion2) Then mProps.Descripcion2 = !Descripcion2
        If Not IsNull(!PesoNeto) Then mProps.PesoNeto = !PesoNeto
        If Not IsNull(!PesoBruto) Then mProps.PesoBruto = !PesoBruto
        If Not IsNull(!IdUnidadPeso) Then mProps.IdUnidadPeso = !IdUnidadPeso
        If Not IsNull(!CodUniPeso) Then mProps.CodUnidadPeso = !CodUniPeso
        If Not IsNull(!BandConversion) Then mProps.BandConversion = !BandConversion
        If Not IsNull(!BandRepGastos) Then mProps.BandRepGastos = !BandRepGastos
        If Not IsNull(!BandNoSeFactura) Then mProps.BandNoSeFactura = !BandNoSeFactura
        If Not IsNull(!TiempoPromVta) Then mProps.TiempoPromVta = !TiempoPromVta
        If Not IsNull(!TiempoReposicion) Then mProps.TiempoReposicion = !TiempoReposicion
        If Not IsNull(!CostoReferencial) Then mProps.CostoReferencial = !CostoReferencial 'AUC 22/12/05
        If Not IsNull(!BandModPrecio) Then mProps.BandModPrecio = !BandModPrecio
        If Not IsNull(!BandAltaRotacion) Then mProps.BandAltaRotacion = !BandAltaRotacion
        If Not IsNull(!BandNoObjetoIVA) Then mProps.BandNoObjetoIVA = !BandNoObjetoIVA
        If Not IsNull(!FrecuenciaReposicion) Then mProps.FrecuenciaReposicion = !FrecuenciaReposicion
        If Not IsNull(!Buffer) Then mProps.Buffer = !Buffer
        If Not IsNull(!FechaModBuffer) Then mProps.FechaModBuffer = !FechaModBuffer
        If Not IsNull(!FrecuenciaReposicionALM) Then mProps.FrecuenciaReposicionALM = !FrecuenciaReposicionALM
        If Not IsNull(!TiempoReposicionALM) Then mProps.TiempoReposicionALM = !TiempoReposicionALM
        If Not IsNull(!M3) Then mProps.M3 = !M3
        
        
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub GrabarEnSucursal(ByVal bd As String)
End Sub

Public Sub GrabaFamiliaDetalleEnSucursal(ByVal bd As String)
End Sub

Private Function verificacionDuiplicadoEnSucursal(ByVal IdInventario As String, ByVal bd As String) As Boolean
End Function

Public Sub RemoveDetalleFamiliaEnSucursal(ByVal bd As String, ByVal ix As Long, Optional obj As IVFamiliaDetalle)
End Sub

Private Sub CargaDetalleFamiliaEnSucursal(ByVal bd As String)
End Sub

Public Sub AseguraMovimientoExistenciaEnSucursal(ByVal bd As String)
End Sub

Public Sub GrabarEnSucursales(ByVal bd As String, ByVal servidor As String)
   Dim sql As String, i As Integer
    Dim rs As Recordset, rsDetalle As Recordset
    'Si no está modificado no hace nada
    If mbooEsNuevo Then
    '    'Verifica si no repite el código
        sql = "SELECT CodInventario FROM " & bd & ".dbo.IVInventario " & _
              "WHERE CodInventario='" & mProps.CodInventario & "' "
        Set rs = mProps.objEmpresa.OpenRecordsetSUC(sql, bd, servidor)
       If Not rs.EOF Then      'Si ya existe el codigo
            Err.Raise ERR_REPITECODIGO, "IVInventario", MSGERR_REPITECODIGO
        End If
        rs.Close
    End If
    'Verifica si no repite el código Alterno 1
    If Not mProps.objEmpresa.GNOpcion.ObtenerValor("CodAltDuplicado") = "1" Then
        If Len(mProps.CodAlterno1) > 0 Then
            sql = "SELECT CodAlterno1 FROM " & bd & ".dbo. IVInventario " & _
                  "WHERE (CodAlterno1='" & mProps.CodAlterno1 & "') AND " & _
                        "(IdInventario<>" & mProps.IdInventario & ")"
            Set rs = mProps.objEmpresa.OpenRecordsetSUC(sql, bd, servidor)
           If Not rs.EOF Then      'Si ya existe el codigo
                Err.Raise ERR_REPITECODIGOALT, "IVInventario", MSGERR_REPITECODIGOALT
            End If
            rs.Close
        End If
    End If
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM " & bd & ".dbo. IVInventario WHERE 1=0"
        'Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        Set rs = mProps.objEmpresa.OpenRecordsetParaEditSuc(sql, bd, servidor)
        rs.AddNew
   'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM " & bd & ".dbo. IVInventario WHERE CodInventario='" & mProps.CodInventario & "'"
        'Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        Set rs = mProps.objEmpresa.OpenRecordsetParaEditSuc(sql, bd, servidor)
        If rs.RecordCount = "0" Then
            sql = "SELECT * FROM " & bd & ".dbo. IVInventario WHERE 1=0"
            'Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
            Set rs = mProps.objEmpresa.OpenRecordsetParaEditSuc(sql, bd, servidor)
           rs.AddNew
        End If
#If DAOLIB Then
        rs.Edit
#End If
    End If
    'Comienza la transaccion ----------------------------jeaa 16-03-04
    On Error GoTo ErrTrap
    'Graba el registro
    With rs
        !CodInventario = mProps.CodInventario
        !CodAlterno1 = mProps.CodAlterno1
        !CodAlterno2 = mProps.CodAlterno2
        !Descripcion = mProps.Descripcion
#If DAOLIB Then
        !DescripcionDetalle = mProps.DescripcionDetalle
#Else
        .Fields("DescripcionDetalle").AppendChunk mProps.DescripcionDetalle
#End If
        !CodMoneda = mProps.CodMoneda
        'For i = 1 To 6 'AUC cambiado antes 4
        For i = 1 To 7 'GRV Precio 7 - 14MAY2014
            .Fields("Precio" & i) = mProps.Precio(i)
            .Fields("Comision" & i) = mProps.Comision(i)
            .Fields("CantLimite" & i) = mProps.CantLimite(i)
            .Fields("Descuento" & i) = mProps.Descuento(i)    '***Agregado. jeaa. 30/jul/2004
        Next i
        !PorcentajeIVA = mProps.PorcentajeIVA
        !IdCuentaActivo = mProps.IdCuentaActivo
        !IdCuentaCosto = mProps.IdCuentaCosto
        !IdCuentaVenta = mProps.IdCuentaVenta
        '!Unidad = mProps.Unidad
        !IdCuentaDiferida = mProps.IdCuentaDiferida
        !IdCuentaDevolucion = mProps.IdCuentaDevolucion
        For i = 1 To IVGRUPO_MAX
            .Fields("IdGrupo" & i) = mProps.IdGrupo(i)
        Next i
        !idProveedor = mProps.idProveedor
        !Observacion = mProps.Observacion
        !ExistenciaMinima = mProps.ExistenciaMinima
        !ExistenciaMaxima = mProps.ExistenciaMaxima
        !UnidadMinimaCompra = mProps.UnidadMinimaCompra
        !UnidadMinimaVenta = mProps.UnidadMinimaCompra
        !BandValida = mProps.BandValida
        !BandServicio = mProps.BandServicio
        !Tipo = mProps.Tipo         'Diego Prod
        'Guarda el detalle de Items en tabla IVMateria
        If mProps.Tipo <> Normal Then
            GrabaFamiliaDetalleEnSucursales bd, servidor
       End If
        !ValorRecargo = mProps.ValorRecargo  '***Agregado. Angel. 29/jul/2004
        !BandFraccion = mProps.BandFraccion  '***Agregado. Esteban 08/04/2005
        !BandArea = mProps.BandArea  '***Agregado. Esteban 15/09/2005
        !bandVenta = mProps.bandVenta '***Agregado. Esteban 26/12/2005
        !IdUnidad = mProps.IdUnidad 'auc
        !IdUnidadConteo = mProps.IdUnidadConteo 'auc
        !BandIVA = mProps.BandIVA '***Agregado. Esteban 08/03/2006
        !PorcentajeICE = mProps.PorcentajeICE
        !PorDesperdicio = mProps.PorDesperdicio  ' *********** jeaa 01/02/2006
        !TipoSRI = mProps.TipoSRI         'jeaa 19/06/2006
        !IdICE = mProps.IdICE         'jeaa 26/07/2006
        !CantRelUnidad = mProps.CantRelUnidad         'jeaa 27/09/2007
        !CantRelUnidadCont = mProps.CantRelUnidadCont         'jeaa 27/07/2007
        
        !CantReluni = mProps.CantReluni         'jeaa 27/09/2007
        !CantReluniCont = mProps.CantReluniCont         'jeaa 27/07/2007
        
        !CantReuni = mProps.CantReuni         'jeaa 27/09/2007
        !CantReuniCont = mProps.CantReuniCont         'jeaa 27/07/2007
        
        
        !Descripcion2 = mProps.Descripcion2             'jeaa 4/06/2008
        !PesoNeto = mProps.PesoNeto         'jeaa 04/06/2008
        !PesoBruto = mProps.PesoBruto         'jeaa 04/06/2008
        !IdUnidadPeso = mProps.IdUnidadPeso 'jeaa 04/06/2008
        !BandConversion = mProps.BandConversion 'jeaa 04/06/2008
        !BandRepGastos = mProps.BandRepGastos 'jeaa 18/08/2008
        !RutaFoto = mProps.RutaFoto 'AUC 04/05/06
        !BandNoSeFactura = mProps.BandNoSeFactura 'jeaa 23/03/2009
        !TiempoPromVta = mProps.TiempoPromVta  '***Agregado. jeaa 24/04/2009
        !TiempoReposicion = mProps.TiempoReposicion '***Agregado. jeaa. 24/04/2009
        !BandModPrecio = mProps.BandModPrecio 'jeaa 18/08/2008
        !BandAltaRotacion = mProps.BandAltaRotacion
        !BandNoObjetoIVA = mProps.BandNoObjetoIVA
        !FrecuenciaReposicion = mProps.FrecuenciaReposicion
        !Buffer = mProps.Buffer
        !FechaModBuffer = mProps.FechaModBuffer
        !FrecuenciaReposicionALM = mProps.FrecuenciaReposicionALM
        !TiempoReposicionALM = mProps.TiempoReposicionALM
        M3 = mProps.M3
        
        !fechagrabado = Now
        .Update
       ''' GrabarProveedorDetalleEnSucursal bd, mProps.CodInventario
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        rs.Move 0           'Para actualizar
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccionEnSucursales bd, servidor, "NUE_ITEM", "Item Nuevo: " & _
                                                        mProps.CodInventario, "DT"
        'jeaa 16-03-04 para grabar en ivexistencia la bodega y cantidad solo cuando es nuevo
        AseguraMovimientoExistenciaEnSucursales bd, servidor
   Else
        mProps.objEmpresa.GrabaGNLogAccionEnSucursales bd, servidor, "MOD_ITEM", "Item Modificado: " & _
                                                        mProps.CodInventario, "DT"
    End If
    Set rs = Nothing
''    mbooEsNuevo = False
''    mbooModificado = False
    'Completa la transacción ---------------------------- jeaa 16-03-04
#If DAOLIB Then
'    gobjWS.CommitTrans
#Else
'    mProps.objEmpresa.Coneccion.CommitTrans
#End If
'   mProps.objEmpresa.RaiseRefreshEvent REFRESH_INVENTARIO
    Exit Sub
ErrTrap:
    'Anula la transacción ---------------------------- jeaa 16-03-04
#If DAOLIB Then
    gobjWS.Rollback
#Else
'    mProps.objEmpresa.Coneccion.RollbackTrans
#End If
    'Genera evento Refresh del objeto Empresa
'    mProps.objEmpresa.RaiseRefreshEvent REFRESH_INVENTARIO
End Sub

Public Sub GrabaFamiliaDetalleEnSucursales(ByVal bd As String, ByVal servidor As String)
   Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As IVFamiliaDetalle, sql As String, band As Boolean
    Dim NumHijos As Integer, CountNumHijos As Integer
    Dim rsIvMat As Recordset, rsIvInv As Recordset
    On Error GoTo CapturaError
    sql = "SELECT IdInventario FROM " & bd & ".dbo.IVInventario WHERE CodInventario='" & mProps.CodInventario & "'"
    Set rsIvInv = mProps.objEmpresa.OpenRecordsetParaEditSuc(sql, bd, servidor)
   'Elimina  todos  los registros
    sql = "Delete Ivm " & _
          " FROM " & bd & ".dbo.Ivinventario Iv " & _
          "INNER JOIN " & bd & ".dbo.IVMateria Ivm ON IV.IdInventario = Ivm.IdMateria "
    'jeaa 28/10/2005
    If mProps.Tipo = CambioPresentacion _
        Or mProps.Tipo = Preparacion _
        Or Tipo = Promocion _
        Or mProps.Tipo = rubro _
        Or Tipo = Preparacionporce Then
          sql = sql & " Where Ivm.IdInventario = " & rsIvInv.Fields("IdInventario")
   Else
          sql = sql & " Where Ivm.Idmateria = " & rsIvInv.Fields("IdInventario")
   End If
    mProps.objEmpresa.EjecutarSQLSuc sql, bd, servidor, numReg
   'guarda  todos  los registros
    band = True
    If Not mcolDetalleFamilia Is Nothing Then
        sql = "SELECT * FROM " & bd & ".dbo.IVMateria WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEditSuc(sql, bd, servidor)
       NumHijos = mcolDetalleFamilia.Count
        CountNumHijos = 0
        For i = 1 To mcolDetalleFamilia.Count
'            rs1.AddNew /jeaa 28-01-04
            If CountNumHijos < NumHijos Then
                CountNumHijos = CountNumHijos + 1
                Set obj = mcolDetalleFamilia.item(i)
                If obj.IdInventario = 0 Then
                        obj.RecuperarEnSucursales obj.CodInventario, bd, servidor
                        If verificacionDuiplicadoEnSucursales(obj.IdInventario, bd, servidor) Then
                            RemoveDetalleFamiliaEnSucursales bd, servidor, i
                           i = i - 1

                        End If
                End If
                If band Then
                    rs1.AddNew
                    sql = "SELECT IdInventario FROM " & bd & ".dbo.IVInventario WHERE CodInventario='" & Me.CodInventario & "'"
                    Set rsIvInv = mProps.objEmpresa.OpenRecordsetParaEditSuc(sql, bd, servidor)
                   sql = "SELECT IdInventario FROM " & bd & ".dbo.IVInventario WHERE CodInventario='" & obj.CodInventario & "'"
                    Set rsIvMat = mProps.objEmpresa.OpenRecordsetParaEditSuc(sql, bd, servidor)
                   If Me.Tipo <> Familia Then
                        rs1!IdInventario = rsIvInv.Fields("IdInventario")
                        rs1!idMateria = rsIvMat.Fields("IdInventario")
                    Else
                        rs1!IdInventario = rsIvMat.Fields("IdInventario")
                        rs1!idMateria = rsIvInv.Fields("IdInventario")
                    End If
                    rs1!TarifaJornal = obj.TarifaJornal
                    rs1!Rendimiento = obj.Rendimiento
                    rs1!Cantidad = obj.Cantidad
                    rs1!BandPrincipal = obj.BandPrincipal
                    rs1!BandModificar = obj.BandModificar 'AUC 25/05/06
                    rs1!Orden = obj.Orden
                    If Me.Tipo = Promocion Then
                        rs1!xCuanto = obj.xCuanto 'AUC 26/09/007
                        rs1!TarifaJornal = obj.Precio  'AUC 26/09/007
                    End If
                    rs1.Update
                    If Me.Tipo = rubro Then 'graba el costoultimoIngreso
                        sql = " Update " & bd & ".dbo.ivinventario  " & _
                              " Set costoultimoingreso = " & obj.TarifaJornal & _
                              " Where codinventario = '" & obj.CodInventario & "'"
                                mProps.objEmpresa.ExecuteSuc sql, bd, servidor, True
                   End If
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

Private Function verificacionDuiplicadoEnSucursales(ByVal IdInventario As String, ByVal bd As String, ByVal servidor As String) As Boolean
   Dim rs As Recordset, sql As String
    verificacionDuiplicadoEnSucursales = False
    'Recupera los detalles de existencia
    sql = "SELECT IdInventario FROM " & bd & ".dbo.ivmateria where idinventario =" & IdInventario
    Set rs = mProps.objEmpresa.OpenRecordsetSUC(sql, bd, servidor)
   If rs.RecordCount > 0 Then
        verificacionDuiplicadoEnSucursales = True
    End If
    Set rs = Nothing
End Function

Public Sub RemoveDetalleFamiliaEnSucursales(ByVal bd As String, ByVal servidor As String, ByVal ix As Long, Optional obj As IVFamiliaDetalle)
    If mcolDetalleFamilia Is Nothing Then CargaDetalleFamiliaEnSucursales bd, servidor
   'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolDetalleFamilia.Count
            If mcolDetalleFamilia.item(ix) Is obj Then
                mcolDetalleFamilia.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolDetalleFamilia.Remove ix
    End If
    'mcolDetalleFamilia.Remove ix
    mbooModificado = True
End Sub

Private Sub CargaDetalleFamiliaEnSucursales(ByVal bd As String, ByVal servidor As String)
   Dim rs As Recordset, sql As String, obj As IVFamiliaDetalle
    Dim i As Long
    Set mcolDetalleFamilia = New Collection
    If Me.Tipo = Preparacion Then
       sql = "Select IV.IdInventario, Iv.CodInventario, Iv.Descripcion, " & _
              "Iv.CodUnidad, Ivm.Cantidad, ivm.bandPrincipal, ivm.bandModificar ,iv.costoultimoingreso  FROM " & bd & ".dbo.VWIvinventarioRECUPERAR Iv " & _
            "INNER JOIN " & bd & ".dbo.IVMateria Ivm "
    ElseIf Me.Tipo = CambioPresentacion Then
        sql = "Select IV.IdInventario, Iv.CodInventario, Iv.Descripcion, " & _
              "Iv.CodUnidad, Ivm.Cantidad  FROM " & bd & ".dbo.VWIvinventarioRECUPERAR Iv " & _
           "INNER JOIN " & bd & ".dbo.IVMateria Ivm "
    ElseIf Me.Tipo = Promocion Then
        sql = "Select IV.IdInventario, Iv.CodInventario, Iv.Descripcion, " & _
             " Iv.CodUnidad,Ivm.Cantidad, ivm.xcuanto, ivm.tarifajornal   FROM " & bd & ".dbo.VWIvinventarioRECUPERAR Iv " & _
            "INNER JOIN " & bd & ".dbo.IVMateria Ivm "
    ElseIf Me.Tipo = rubro Then
       sql = "Select IV.IdInventario, Iv.CodInventario, Iv.Descripcion, " & _
              "ivm.formula,ivm.bandModificar,ivm.orden FROM " & bd & ".dbo.VWIvinventarioRECUPERAR Iv " & _
            "INNER JOIN " & bd & ".dbo.IVMateria Ivm "
    ElseIf Me.Tipo = Preparacionporce Then
       sql = "Select IV.IdInventario, Iv.CodInventario, Iv.Descripcion, " & _
              "Iv.CodUnidad, Ivm.Cantidad, ivm.Rendimiento, ivm.bandPrincipal FROM " & bd & ".dbo.VWIvinventarioRECUPERAR Iv " & _
            "INNER JOIN " & bd & ".dbo.IVMateria Ivm "
    Else
        sql = "Select IV.IdInventario, Iv.CodInventario, Iv.Descripcion, " & _
              "Iv.codunidad, Ivm.Cantidad FROM " & bd & ".dbo.VWIvinventarioRECUPERAR IV " & _
              "INNER JOIN " & bd & ".dbo.IVMateria Ivm "
    End If
    '03/10/2005
    If Me.Tipo = CambioPresentacion _
        Or Me.Tipo = Preparacion _
        Or Me.Tipo = Promocion _
        Or Me.Tipo = rubro Then
        sql = sql & "  ON IV.IdInventario = Ivm.Idmateria"
        sql = sql & " Where Ivm.IdInventario = " & mProps.IdInventario
        sql = sql & " order by orden"
    ElseIf Me.Tipo = Preparacionporce Then
        sql = sql & "  ON IV.IdInventario = Ivm.Idmateria"
        sql = sql & " Where Ivm.IdInventario = " & mProps.IdInventario
        sql = sql & " AND IV.TIPO<>2 "
        sql = sql & " order by orden"
    Else
          sql = sql & " ON IV.IdInventario = Ivm.IdInventario "
            sql = sql & " Where Ivm.IdMateria = " & mProps.IdInventario
    End If
    Set rs = mProps.objEmpresa.OpenRecordsetSUC(sql, bd, servidor)
   With rs
        Do Until .EOF
            i = AddDetalleFamilia
            Set obj = mcolDetalleFamilia.item(i)
            If Me.Tipo = rubro Then
               obj.IdInventario = !IdInventario
                obj.CodInventario = !CodInventario
                obj.Descripcion = !Descripcion
                If Not IsNull(!Formula) Then obj.Formula = !Formula
                obj.BandModificar = !BandModificar
                obj.Orden = !Orden
            Else
                  obj.IdInventario = !IdInventario
                  obj.CodInventario = !CodInventario
                  obj.Descripcion = !Descripcion
                  If Not IsNull(!CodUnidad) Then obj.Unidad = !CodUnidad 'AUC 'cambiado antes unidad
                 obj.Cantidad = !Cantidad
                  If Me.Tipo = Preparacion Then
                      If Len(obj.IVInventario.Empresa.GNOpcion.ObtenerValor("itemPreparacion")) > 0 Then 'AUC toma el p.costo
                          If obj.IVInventario.Empresa.GNOpcion.ObtenerValor("itemPreparacion") = 0 Then 'AUC toma el p.costo
                             If Not IsNull(!CostoUltimoIngreso) Then obj.CostoUltimoIngreso = !CostoUltimoIngreso
                           ElseIf obj.IVInventario.Empresa.GNOpcion.ObtenerValor("itemPreparacion") = 1 Then 'AUC toma el p. venta
                              obj.pVenta = obj.ObtienePrecio(obj.CodInventario)
                           End If
                           obj.BandPrincipal = !BandPrincipal  'AUC 10/05/06 restriccion del item
                           obj.BandModificar = !BandModificar   'AUC 25/05/06
                      Else
                          obj.pVenta = obj.ObtienePrecio(obj.CodInventario)
                      End If
                  ElseIf Me.Tipo = Promocion Then
                          obj.xCuanto = !xCuanto
                          obj.Precio = !TarifaJornal
                  ElseIf Me.Tipo = Preparacionporce Then
                        obj.BandPrincipal = !BandPrincipal  'AUC 10/05/06 restriccion del item
                        obj.Rendimiento = !Rendimiento 'AUC 10/05/06 restriccion del item
                End If
            End If
           .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub AseguraMovimientoExistenciaEnSucursales(ByVal bd As String, ByVal servidor As String)
   Dim CodBodega As String
    Dim sql As String, rs As Recordset, IdBodega As Integer, idItem As Long
    Dim rsbod As Recordset
    sql = "SELECT valor FROM " & bd & ".dbo.Gnopcion2 where codigo='BodegaPredeterminada'"
    Set rsbod = mProps.objEmpresa.OpenRecordsetSUC(sql, bd, servidor)
   'recupera la bodega predeterminada
    CodBodega = rsbod.Fields("valor")
    ' si no tiene bodega predeterminda sale sin grabar en ivexist
    If Len(CodBodega) = 0 Then Exit Sub
    sql = "SELECT IdBodega FROM " & bd & ".dbo.IVBodega where codbodega='" & CodBodega & "'"
    'recupera idBodega
    Set rs = mProps.objEmpresa.OpenRecordsetSUC(sql, bd, servidor)
   IdBodega = rs.Fields("IdBodega")
    'recupera iditem
    sql = "SELECT IdInventario FROM " & bd & ".dbo.IVinventario where codinventario='" & mProps.CodInventario & "'"
    Set rs = mProps.objEmpresa.OpenRecordsetSUC(sql, bd, servidor)
   idItem = rs.Fields("IdInventario")
    sql = "INSERT INTO " & bd & ".dbo.IVExist (IdInventario,IdBodega,Exist) " & _
              "VALUES(" & idItem & "," & IdBodega & "," & 0 & ")"
    mProps.objEmpresa.OpenRecordsetSUC sql, bd, servidor
   'jeaa 20/11/2006
    sql = "INSERT INTO " & bd & ".dbo.IVCompr (IdInventario,IdBodega,Compr) " & _
              "VALUES(" & idItem & "," & IdBodega & "," & 0 & ")"
    mProps.objEmpresa.OpenRecordsetSUC sql, bd, servidor
   rs.Close
    Set rs = Nothing
End Sub

Private Function ObtieneUltimoCompra(BandIngreso As Boolean) As Boolean
    Dim sql As String, rs As Recordset, cond1 As String
    Dim campo As String
    On Error GoTo ErrTrap
    
    If BandIngreso Then
        If mUIngreso.Recuperado Then
            ObtieneUltimoCompra = True
            Exit Function
        End If
    Else
        If mUEgreso.Recuperado Then
            ObtieneUltimoCompra = True
            Exit Function
        End If
    End If
    
    cond1 = IIf(BandIngreso, "(Ingreso>0)", "(Egreso>0)")
    
    Select Case Me.Empresa.GNOpcion.IndiceMoneda(Me.CodMoneda)
    Case 1: campo = "Costo AS CU,Precio AS PU"
    Case 2: campo = "Costo/Cotizacion2 AS CU,Precio/Cotizacion2 AS PU"
    Case 3: campo = "Costo/Cotizacion3 AS CU,Precio/Cotizacion3 AS PU"
    Case 4: campo = "Costo/Cotizacion4 AS CU,Precio/Cotizacion4 AS PU"
    End Select
    
    '*** MAKOTO 07/dic/00 Agregado HoraTrans DESC
    sql = "SELECT FechaTrans,Ingreso,Egreso," & campo & ", comision FROM vwConsIVKardex " & _
            " inner join gntrans on vwConsIVKardex .codtrans=gntrans.codtrans " & _
          "WHERE " & cond1 & " AND (CodInventario='" & Me.CodInventario & "') " & _
          " and substring(opcion,189,1)='S' " & _
          "ORDER BY FechaTrans DESC, HoraTrans DESC, TransID DESC"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        If BandIngreso Then
            With mUIngreso
                .fecha = rs!FechaTrans
                .Cantidad = rs!ingreso
                .Costo = rs!cu
                .Precio = rs!PU
                If Not IsNull(rs!Comision) Then .Comision = rs!Comision / rs!ingreso
                .Recuperado = True
            End With
        Else
            With mUEgreso
                .fecha = rs!FechaTrans
                .Cantidad = rs!egreso
                .Costo = rs!cu
                .Precio = rs!PU
                .Recuperado = True
            End With
        End If
        ObtieneUltimoCompra = True
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
ErrTrap:
    Set rs = Nothing
    Exit Function
End Function


Friend Sub RecuperarQuick(key As Variant)
    Dim sql As String, i As Integer
    Dim rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM vwIVInventarioRecuperar WHERE CodInventario='" & key & "'"
    Else
        sql = "SELECT * FROM vwIVInventarioRecuperar WHERE IdInventario=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IdInventario = !IdInventario
        mProps.CodInventario = !CodInventario
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        'For i = 1 To 6 'AUC cambiado antes 4
        For i = 1 To 7 'GRV Precio 7 - 14MAY2014
            If Not IsNull(.Fields("Precio" & i)) Then mProps.Precio(i) = .Fields("Precio" & i)
            If Not IsNull(.Fields("Comision" & i)) Then mProps.Comision(i) = .Fields("Comision" & i)
            If Not IsNull(.Fields("CantLimite" & i)) Then mProps.CantLimite(i) = .Fields("CantLimite" & i)
            If Not IsNull(.Fields("Descuento" & i)) Then mProps.Descuento(i) = .Fields("Descuento" & i)   '***Agregado. jeaa. 29/jul/2004
        Next i
        For i = 1 To IVGRUPO_MAX
            If Not IsNull(.Fields("IdGrupo" & i)) Then mProps.IdGrupo(i) = .Fields("IdGrupo" & i)
            If Not IsNull(.Fields("CodGrupo" & i)) Then mProps.CodGrupo(i) = .Fields("CodGrupo" & i)
        Next i
        
        If Not IsNull(!PorcentajeIVA) Then mProps.PorcentajeIVA = !PorcentajeIVA
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!BandServicio) Then mProps.BandServicio = !BandServicio
        If Not IsNull(!CostoUltimoIngreso) Then mProps.CostoUltimoIngreso = !CostoUltimoIngreso 'AUC 22/12/05
        If Not IsNull(!BandIVA) Then mProps.BandIVA = !BandIVA  'jeaa 08/03/2006
        If Not IsNull(!Tipo) Then mProps.Tipo = !Tipo  'jeaa 08/03/2006
        If Not IsNull(!CostoReferencial) Then mProps.CostoReferencial = !CostoReferencial 'AUC 22/12/05
        If Not IsNull(!PorDesperdicio) Then mProps.PorDesperdicio = !PorDesperdicio
        If Not IsNull(!RutaFoto) Then mProps.RutaFoto = !RutaFoto
        If Not IsNull(!CodArancel) Then mProps.CodArancel = !CodArancel
        If Not IsNull(!BandNoObjetoIVA) Then mProps.BandNoObjetoIVA = !BandNoObjetoIVA
        If Not IsNull(!BandGarantia) Then mProps.BandGarantia = !BandGarantia
        If Not IsNull(!BandCambiaDescr) Then mProps.BandCambiaDescr = !BandCambiaDescr
        If Not IsNull(!TipoAuto) Then mProps.TipoAuto = !TipoAuto
        If Not IsNull(!bandSerie) Then mProps.bandSerie = !bandSerie
        If Not IsNull(!bandPeso) Then mProps.bandPeso = !bandPeso
        If Not IsNull(!BandRepGastos) Then mProps.BandRepGastos = !BandRepGastos
        If Not IsNull(!BandNoSeFactura) Then mProps.BandNoSeFactura = !BandNoSeFactura
        If Not IsNull(!IdICE) Then mProps.IdICE = !IdICE
        
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Function RecuperaExistenciaFecha(ByVal CodInventario As String, CodBodega As String, FechaCorte As Date) As Currency
    Dim sql As String
    Dim rs As Recordset
    If gobjMain.EmpresaActual.GNOpcion.IVKTipoDatoDouble Then
        sql = "select isnull(sum(ivk.cantidadDou),0) as cant"
    Else
        sql = "select isnull(sum(ivk.cantidad),0) as cant"
    End If
    sql = sql & " from gntrans gnt"
    sql = sql & " inner join gncomprobante g"
    sql = sql & " inner join ivkardex ivk inner join ivbodega ivb on ivb.idbodega = ivk.idbodega"
    sql = sql & " inner join ivinventario ivi"
    sql = sql & " on ivk.idinventario = ivi.idinventario"
    sql = sql & " on g.transid=ivk.transid"
    sql = sql & " on gnt.codtrans = g.codtrans"
    sql = sql & " Where g.Estado <> 3"
    sql = sql & " and GNt.AfectaCantidad=1"
    sql = sql & " and g.fechatrans <=" & FechaYMD(FechaCorte, gobjMain.EmpresaActual.TipoDB)
    sql = sql & " and codinventario='" & CodInventario & "'"
    sql = sql & " and ivb.codbodega='" & CodBodega & "'"
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    If rs.RecordCount > 0 Then
        RecuperaExistenciaFecha = rs.Fields("cant")
    Else
        RecuperaExistenciaFecha = 0
    End If
    Set rs = Nothing
End Function


Public Function ObtieneExistenciaItem(ByVal CodInventario As String) As Currency
    Dim sql As String, rs As Recordset, cond1 As String
    Dim campo As String
    On Error GoTo ErrTrap
    ObtieneExistenciaItem = 0
    
    '*** MAKOTO 07/dic/00 Agregado HoraTrans DESC
    sql = "SELECT isnull(exist,0) as exist  FROM vwIvexist "
    sql = sql & "WHERE CodInventario= '" & CodInventario & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        ObtieneExistenciaItem = rs.Fields("exist")
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
ErrTrap:
    Set rs = Nothing
    Exit Function
End Function


Public Function CostoPromedioIngresos( _
                    ByVal fecha As Date, _
                    Optional ByVal cant As Currency, _
                    Optional ByVal tid As Long) As Double
    '*** MAKOTO 08/dic/00 Modificado
    CostoPromedioIngresos = CostoPromedioIng(fecha, cant, tid)
End Function


Public Function CostoPromedioIng( _
                    ByVal fecha As Date, _
                    Optional ByVal cant As Currency, _
                    Optional ByVal tid As Long, _
                    Optional ByVal Hora As Date) As Double
    'Si la id de transaccion no está especificada, cambia la fecha limite
    '  para que incluya todas las trans. de la misma fecha
    If Hora = 0 Then Hora = TimeSerial(23, 59, 59)

        CostoPromedioIng = CostoPromIng(fecha, tid, Hora)
    
    '*** MAKOTO 31/ene/01 Agregado para redondear al formato de costo
    CostoPromedioIng = MiCCur(Format$(CostoPromedioIng, Me.Empresa.GNOpcion.ObtenerValor( _
                            "FormatoCosto" & Me.Empresa.GNOpcion.IndiceMoneda(Me.CodMoneda))))
    CostoPromedioIng = MiCDbl(Format$(CostoPromedioIng, Me.Empresa.GNOpcion.ObtenerValor( _
                            "FormatoCosto" & Me.Empresa.GNOpcion.IndiceMoneda(Me.CodMoneda))))


                            
End Function


Private Function CostoPromIng( _
                    ByVal fecha As Date, _
                    ByVal tid As Long, _
                    ByVal Hora As Date) As Double
    Dim sql As String, num As Integer, campo As String
    Dim rs As Recordset, Costo As Double        '*** MAKOTO 23/nov/00 Modificado a Double
    Dim idItem As Long
    num = mProps.objEmpresa.GNOpcion.IndiceMoneda(Me.CodMoneda)
    campo = "CP" & num
#If DAOLIB Then
    Dim qdef As QueryDef
    Set qdef = mProps.objEmpresa.Database.QueryDefs("spIVCostoProm")
        If Me.Tipo = Preparacion Then 'AUC agregado para sacar el costo de un item de preparacion
        qdef.Parameters("IdItem") = idInv
    Else
        qdef.Parameters("IdItem") = mProps.IdInventario
    End If
    'qdef.Parameters("IdItem") = mProps.IdInventario
    qdef.Parameters("FechaHasta") = fecha
    qdef.Parameters("tid") = tid
    qdef.Parameters("HoraHasta") = Hora         '*** MAKOTO 09/dic/00 Agregado
    Set rs = qdef.OpenRecordset(dbOpenSnapshot, dbReadOnly)
    If Not rs.EOF Then
        If Not IsNull(rs.Fields(campo)) Then Costo = rs.Fields(campo)
    End If
#Else
    '*** MAKOTO 08/dic/00 Modificado
    sql = "spIVCostoPromIng2 " & _
                mProps.IdInventario & ", " & _
                FechaYMD("01/11/2010", Me.Empresa.TipoDB) & ", " & _
                FechaYMD(fecha, Me.Empresa.TipoDB) & ", " & _
                tid & ", " & _
                num & ", " & _
                HoraHMS(Hora, Me.Empresa.TipoDB)
    Set rs = Me.Empresa.OpenRecordset(sql)
    Costo = rs.Fields("CostoProm")
#End If
    
    rs.Close
    Set rs = Nothing

    'Si no hay costo promedio por que no hay existencia,
    '   devuelve el último costo en vez de promedio
    If Costo = 0 Then Costo = CostoUltimo(fecha, tid, Hora, False)
    
    CostoPromIng = Costo
End Function


Public Property Let BandModPrecio(ByVal value As Boolean)
    mProps.BandModPrecio = value
    mbooModificado = True
End Property

Public Property Get BandModPrecio() As Boolean
    BandModPrecio = mProps.BandModPrecio
End Property

Public Property Let CodArancel(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodArancel Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdARANCEL, codArancel FROM IVRecargoARANCEL WHERE CodARANCEL='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVARANCEL_CodARANCEL", MSGERR_NOHAYCODIGO
                Exit Property
            Else  'ESTOY AQUI
                IdARANCEL = rs!IdARANCEL
            End If
            rs.Close
         End If
        Else
            IdARANCEL = 0
        End If
        mProps.CodArancel = value
        mbooModificado = True
    
    Set rs = Nothing
End Property

Public Property Get CodArancel() As String
    CodArancel = Trim$(mProps.CodArancel)
End Property

Public Property Let IdARANCEL(ByVal value As Long)
    mProps.IdARANCEL = value
    mbooModificado = True
End Property

Public Property Get IdARANCEL() As Long
    IdARANCEL = mProps.IdARANCEL
End Property

Friend Sub RecuperarALt2(key As Variant)
    Dim sql As String, i As Integer
    Dim rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM vwIVInventarioRecuperar WHERE CodAlterno2 ='" & key & "'"
    Else
        sql = "SELECT * FROM vwIVInventarioRecuperar WHERE IdInventario=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IdInventario = !IdInventario
        mProps.CodInventario = !CodInventario
        If Not IsNull(!CodAlterno1) Then mProps.CodAlterno1 = !CodAlterno1
        If Not IsNull(!CodAlterno2) Then mProps.CodAlterno2 = !CodAlterno2
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
#If DAOLIB Then
        If Not IsNull(!DescripcionDetalle) Then mProps.DescripcionDetalle = !DescripcionDetalle
#Else
        If Not IsNull(!DescripcionDetalle) Then
            If .Fields("DescripcionDetalle").ActualSize > 0 Then
                mProps.DescripcionDetalle = .Fields("DescripcionDetalle").GetChunk(.Fields("DescripcionDetalle").ActualSize)
            End If
        End If
#End If
        If Not IsNull(!CodMoneda) Then mProps.CodMoneda = !CodMoneda
        
        'For i = 1 To 6 'AUC cambiado antes 4
         For i = 1 To 7 'GRV Precio 7 - 14MAY2014
            If Not IsNull(.Fields("Precio" & i)) Then mProps.Precio(i) = .Fields("Precio" & i)
            If Not IsNull(.Fields("Comision" & i)) Then mProps.Comision(i) = .Fields("Comision" & i)
            If Not IsNull(.Fields("CantLimite" & i)) Then mProps.CantLimite(i) = .Fields("CantLimite" & i)
            If Not IsNull(.Fields("Descuento" & i)) Then mProps.Descuento(i) = .Fields("Descuento" & i)   '***Agregado. jeaa. 29/jul/2004
        Next i
        If Not IsNull(!PorcentajeIVA) Then mProps.PorcentajeIVA = !PorcentajeIVA
        If Not IsNull(!IdCuentaActivo) Then mProps.IdCuentaActivo = !IdCuentaActivo
        If Not IsNull(!IdCuentaCosto) Then mProps.IdCuentaCosto = !IdCuentaCosto
        If Not IsNull(!IdCuentaVenta) Then mProps.IdCuentaVenta = !IdCuentaVenta
        If Not IsNull(!IdCuentaDiferida) Then mProps.IdCuentaDiferida = !IdCuentaDiferida
        If Not IsNull(!IdCuentaDevolucion) Then mProps.IdCuentaDevolucion = !IdCuentaDevolucion
        
        
        If Not IsNull(!CodCuentaActivo) Then mProps.CodCuentaActivo = !CodCuentaActivo
        If Not IsNull(!CodCuentaCosto) Then mProps.CodCuentaCosto = !CodCuentaCosto
        If Not IsNull(!CodCuentaVenta) Then mProps.CodCuentaVenta = !CodCuentaVenta
        If Not IsNull(!CodCuentaDiferida) Then mProps.CodCuentaDiferida = !CodCuentaDiferida
'        If Not IsNull(!Unidad) Then mProps.Unidad = !Unidad

        If Not IsNull(!CodCuentaDiferida) Then mProps.CodCuentaDiferida = !CodCuentaDiferida
        If Not IsNull(!CodCuentaDevolucion) Then mProps.CodCuentaDevolucion = !CodCuentaDevolucion

        For i = 1 To IVGRUPO_MAX
            If Not IsNull(.Fields("IdGrupo" & i)) Then mProps.IdGrupo(i) = .Fields("IdGrupo" & i)
            If Not IsNull(.Fields("CodGrupo" & i)) Then mProps.CodGrupo(i) = .Fields("CodGrupo" & i)
        Next i
        
        If Not IsNull(!idProveedor) Then mProps.idProveedor = !idProveedor

        CargaDetalleProveedor  'auc
        If Not IsNull(!CodProvcli) Then mProps.CodProveedor = !CodProvcli
        If Not IsNull(!Observacion) Then mProps.Observacion = !Observacion
        If Not IsNull(!ExistenciaMinima) Then mProps.ExistenciaMinima = !ExistenciaMinima
        If Not IsNull(!ExistenciaMaxima) Then mProps.ExistenciaMaxima = !ExistenciaMaxima
        If Not IsNull(!UnidadMinimaCompra) Then mProps.UnidadMinimaCompra = !UnidadMinimaCompra
        If Not IsNull(!UnidadMinimaCompra) Then mProps.UnidadMinimaVenta = !UnidadMinimaCompra
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!BandServicio) Then mProps.BandServicio = !BandServicio
        If Not IsNull(!Tipo) Then mProps.Tipo = !Tipo
        'DiegoProd  Ojo  aumentar  proceso  para  recuperar  detalle
        If Tipo <> Normal Then
            CargaDetalleFamilia 'Pensandolo bien recuperamos
        End If
        '***Agregado. Angel. 29/jul/2004
        If Not IsNull(!ValorRecargo) Then mProps.ValorRecargo = !ValorRecargo
        '*** Agregado esteban 08/04/2005
        If Not IsNull(!BandFraccion) Then mProps.BandFraccion = !BandFraccion
        '*** Agregado esteban 15/09/2005
        If Not IsNull(!BandArea) Then mProps.BandArea = !BandArea
        '*** Agregado esteban 26/12/2005
        If Not IsNull(!bandVenta) Then mProps.bandVenta = !bandVenta
         If Not IsNull(!CodUnidad) Then mProps.CodUnidad = !CodUnidad
        If Not IsNull(!CodUniConteo) Then mProps.CodUnidadConteo = !CodUniConteo
        If Not IsNull(!IdUnidad) Then mProps.IdUnidad = !IdUnidad
        If Not IsNull(!IdUnidadConteo) Then mProps.IdUnidadConteo = !IdUnidadConteo
        If Not IsNull(!CostoUltimoIngreso) Then mProps.CostoUltimoIngreso = !CostoUltimoIngreso 'AUC 22/12/05
        If Not IsNull(!BandIVA) Then mProps.BandIVA = !BandIVA  'jeaa 08/03/2006
        If Not IsNull(!PorcentajeICE) Then mProps.PorcentajeICE = !PorcentajeICE
        If Not IsNull(!PorDesperdicio) Then mProps.PorDesperdicio = !PorDesperdicio 'jeaa 01/02/2006
        If Not IsNull(!RutaFoto) Then mProps.RutaFoto = !RutaFoto  ' AUC 04/05/06
        If Not IsNull(!TipoSRI) Then mProps.TipoSRI = !TipoSRI
        If Not IsNull(!IdICE) Then mProps.IdICE = !IdICE
        If Not IsNull(!CodICE) Then mProps.CodICE = !CodICE
        If Not IsNull(!CantRelUnidad) Then mProps.CantRelUnidad = !CantRelUnidad 'jeaa 27/0/2007
        If Not IsNull(!CantRelUnidadCont) Then mProps.CantRelUnidadCont = !CantRelUnidadCont 'jeaa 27/0/2007
        
        If Not IsNull(!CantReluni) Then mProps.CantReluni = !CantReluni 'jeaa 27/0/2007
        If Not IsNull(!CantReluniCont) Then mProps.CantReluniCont = !CantReluniCont 'jeaa 27/0/2007
        
        If Not IsNull(!CantReuni) Then mProps.CantReuni = !CantReuni 'jeaa 27/0/2007
        If Not IsNull(!CantReuniCont) Then mProps.CantReuniCont = !CantReuniCont 'jeaa 27/0/2007
        
        
        
        If Not IsNull(!Descripcion2) Then mProps.Descripcion2 = !Descripcion2
        If Not IsNull(!PesoNeto) Then mProps.PesoNeto = !PesoNeto
        If Not IsNull(!PesoBruto) Then mProps.PesoBruto = !PesoBruto
        If Not IsNull(!IdUnidadPeso) Then mProps.IdUnidadPeso = !IdUnidadPeso
        If Not IsNull(!CodUniPeso) Then mProps.CodUnidadPeso = !CodUniPeso
        If Not IsNull(!BandConversion) Then mProps.BandConversion = !BandConversion
        If Not IsNull(!BandRepGastos) Then mProps.BandRepGastos = !BandRepGastos
        If Not IsNull(!BandNoSeFactura) Then mProps.BandNoSeFactura = !BandNoSeFactura
        If Not IsNull(!TiempoPromVta) Then mProps.TiempoPromVta = !TiempoPromVta
        If Not IsNull(!TiempoReposicion) Then mProps.TiempoReposicion = !TiempoReposicion
        If Not IsNull(!CostoReferencial) Then mProps.CostoReferencial = !CostoReferencial 'AUC 22/12/05
        If Not IsNull(!BandModPrecio) Then mProps.BandModPrecio = !BandModPrecio
        If Not IsNull(!BandAltaRotacion) Then mProps.BandAltaRotacion = !BandAltaRotacion
        If Not IsNull(!BandNoObjetoIVA) Then mProps.BandNoObjetoIVA = !BandNoObjetoIVA
        If Not IsNull(!FrecuenciaReposicion) Then mProps.FrecuenciaReposicion = !FrecuenciaReposicion
        If Not IsNull(!Buffer) Then mProps.Buffer = !Buffer
        If Not IsNull(!FechaModBuffer) Then mProps.FechaModBuffer = !FechaModBuffer
        If Not IsNull(!FrecuenciaReposicionALM) Then mProps.FrecuenciaReposicionALM = !FrecuenciaReposicionALM
        If Not IsNull(!TiempoReposicionALM) Then mProps.TiempoReposicionALM = !TiempoReposicionALM
        If Not IsNull(!M3) Then mProps.M3 = !M3
        
        
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Property Let NumItems(ByVal value As Integer)
    If mProps.NumItems <> value Then
        mProps.NumItems = value
    End If
End Property

Public Property Get NumItems() As Integer
    NumItems = mProps.NumItems
End Property

Public Property Let EsNuevo(ByVal value As Boolean)
    mbooEsNuevo = value
End Property

Public Property Let BandConversionUni(ByVal value As Boolean)
    mProps.BandConversionUni = value
    mbooModificado = True
End Property

Public Property Get BandConversionUni() As Boolean
    BandConversionUni = mProps.BandConversionUni
End Property

Public Function RecuperaDetalleConversion(ix As Long) As IVConversion
    Set RecuperaDetalleConversion = mcolDetalleConversion.item(ix)
End Function
Public Property Get NumConversionDetalle() As Long
    If mcolDetalleConversion Is Nothing Then
        NumConversionDetalle = 0
    Else
        NumConversionDetalle = mcolDetalleConversion.Count
    End If
End Property

Public Function AddDetalleConversion() As Long
    Dim nuevo As IVConversion
    
    If mcolDetalleConversion Is Nothing Then CargaDetalleConversion
    Set nuevo = New IVConversion
    Set nuevo.IVInventario = Me  'Me  olvide  del padre
    nuevo.Cantidad = 1
    mcolDetalleConversion.Add nuevo
    AddDetalleConversion = mcolDetalleConversion.Count
    Set nuevo = Nothing
    mbooModificado = True
End Function

Private Sub CargaDetalleConversion()
    Dim rs As Recordset, sql As String, obj As IVConversion
    Dim i As Long
    Set mcolDetalleConversion = New Collection
    
        sql = "Select IVC.IdUnidadOri, u.Codunidad as unidadOri, ivc.IdUnidadDes,u1.codunidad as unidadDes,Ivc.Cantidad,ivc.precio,ivc.CantidadDou " & _
              "From  IVConversion ivc " & _
              "INNER JOIN ivunidad u on u.idunidad = ivc.idunidadori " & _
              " INNER JOIN ivunidad u1 on u1.idunidad = ivc.idunidaddes  "
        sql = sql & " Where Ivc.IdInventario = " & mProps.IdInventario
        sql = sql & " order by orden"

    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddDetalleConversion
            Set obj = mcolDetalleConversion.item(i)
            obj.IdUnidadOri = !IdUnidadOri
            obj.CodUnidadOri = !unidadori
            obj.IdUnidadDes = !IdUnidadDes
            obj.CodUnidadDes = !Unidaddes
            obj.Cantidad = !Cantidad
            obj.Precio = !Precio
            obj.CantidadDou = !CantidadDou
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub RemoveDetalleConversion(ByVal ix As Long, Optional obj As IVConversion)
    If mcolDetalleConversion Is Nothing Then CargaDetalleConversion
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolDetalleConversion.Count
            If mcolDetalleConversion.item(ix) Is obj Then
                mcolDetalleConversion.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolDetalleConversion.Remove ix
    End If
    
    mbooModificado = True
End Sub


Public Sub GrabaConversionDetalle()
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As IVConversion, sql As String, band As Boolean
    Dim NumHijos As Integer, CountNumHijos As Integer
    On Error GoTo CapturaError
    
    'Elimina  todos  los registros
    sql = "Delete Ivc " & _
          "From Ivinventario Iv " & _
          "INNER JOIN IVConversion Ivc ON IV.IdInventario = Ivc.Idinventario "
          sql = sql & " Where Ivc.Idinventario = " & mProps.IdInventario

    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    band = True
    If Not mcolDetalleConversion Is Nothing Then
        sql = "SELECT * FROM IVconversion WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        NumHijos = mcolDetalleConversion.Count
        CountNumHijos = 0
        For i = 1 To mcolDetalleConversion.Count
            If CountNumHijos < NumHijos Then
                CountNumHijos = CountNumHijos + 1
                Set obj = mcolDetalleConversion.item(i)
                'If obj.IdInventario = 0 Then
                        'obj.Recuperar obj.CodInventario
                        'If verificacionDuiplicado(obj.IdInventario) Then
                         '   RemoveDetalleFamilia i
                          '  i = i - 1
    '                        band = False
                       ' End If
                'End If
                If band Then
                    rs1.AddNew
                    rs1!IdInventario = Me.IdInventario
                    rs1!IdUnidadOri = obj.IdUnidadOri
                    rs1!IdUnidadDes = obj.IdUnidadDes
                    rs1!Cantidad = obj.Cantidad
                    rs1!CantidadDou = obj.CantidadDou
                    rs1!Precio = obj.Precio
                    rs1!Orden = obj.Orden
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

Public Function ListaDetalleConversion() As String
    Dim sql As String, rs As Recordset
    Dim s As String
      
              
        sql = "SELECT ivc.orden, u.codunidad,"
        If gobjMain.EmpresaActual.GNOpcion.IVKTipoDatoDouble Then
            sql = sql & " ivc.cantidadDou,"
        Else
            sql = sql & " ivc.cantidad,"
        End If
        sql = sql & "u1.codunidad FROM ivconversion ivc inner join ivinventario iv"
        sql = sql & " on iv.idinventario = ivc.idinventario"
        sql = sql & " inner join ivunidad u on u.idunidad = ivc.idunidadori"
        sql = sql & " inner join ivunidad u1 on u1.idunidad = ivc.idunidaddes"
        sql = sql & " Where iv.codInventario = '" & Me.CodInventario & "'"
        
        Set rs = Me.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then s = "|" & rs.GetString(adClipString, , vbTab, "|", "(nulo)")
        Set rs = Nothing
        ListaDetalleConversion = s
End Function

Public Function Conversion(ByVal CodUnidad As String, ByRef UnidadDest As String) As Currency
Dim sql As String
Dim rs As Recordset
On Error GoTo CapturaError
        If gobjMain.EmpresaActual.GNOpcion.IVKTipoDatoDouble Then
            sql = "SELECT ivc.cantidadDou,U2.codUnidad  "
        Else
            sql = "SELECT ivc.cantidad,U2.codUnidad  "
        End If
        sql = "SELECT ivc.cantidad,U2.codUnidad  "
        sql = sql & "FROM IVConversion ivc inner join ivunidad u  " & _
              "ON u.Idunidad = IVC.IdUnidadori  " & _
              "INNER JOIN ivunidad u2 on u2.idunidad= ivc.idunidaddes " & _
              "inner join ivinventario iv on iv.idinventario = ivc.idinventario " & _
              "Where iv.codInventario = '" & Me.CodInventario & "'"
        'sql = sql & " AND u.codunidad = '" & CodUnidad & "'"
        sql = sql & " AND ivc.orden = '" & CodUnidad & "'"
        Set rs = Me.Empresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_INVALIDO, "IVConversion_CodUnidad", MSGERR_INVALIDO
            Exit Function
        End If
        If rs.RecordCount > 1 Then
            Err.Raise ERR_INVALIDO, "IVConversion_Existe mas de una conversion con la misma unidad", MSGERR_INVALIDO
            Exit Function
        End If
        
        Do While Not rs.EOF
            If gobjMain.EmpresaActual.GNOpcion.IVKTipoDatoDouble Then
                Conversion = rs!CantidadDou
            Else
                Conversion = rs!Cantidad
            End If
            UnidadDest = rs!CodUnidad
            rs.MoveNext
        Loop
        Exit Function
CapturaError:
    MsgBox Err.Description
    Set rs = Nothing
    Exit Function
End Function


Public Sub GrabaFamiliaDetalleNew(ByVal Tipo As Integer)
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As IVFamiliaDetalle, sql As String, band As Boolean
    Dim NumHijos As Integer, CountNumHijos As Integer
    On Error GoTo CapturaError
    'Elimina  todos  los registros
    sql = "Delete Ivm " & _
         "From Ivinventario Iv " & _
          "INNER JOIN IVMateria Ivm ON IV.IdInventario = Ivm.IdMateria "
    'jeaa 28/10/2005
    If mProps.Tipo = CambioPresentacion _
        Or mProps.Tipo = Preparacion _
        Or Tipo = Promocion _
        Or mProps.Tipo = rubro _
        Or Tipo = Preparacionporce Then
          sql = sql & " Where Ivm.IdInventario = " & mProps.IdInventario
  Else
          sql = sql & " Where Ivm.Idmateria = " & mProps.IdInventario
  End If
   sql = sql & " AND Ivm.Tipo= " & Tipo
    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    If Tipo = 0 Then
        band = True
        If Not mcolDetalleFamilia Is Nothing Then
            sql = "SELECT * FROM IVMateria WHERE 1=0 and tipo=0"
            Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
            NumHijos = mcolDetalleFamilia.Count
            CountNumHijos = 0
            For i = 1 To mcolDetalleFamilia.Count
    '            rs1.AddNew /jeaa 28-01-04
                If CountNumHijos < NumHijos Then
                    CountNumHijos = CountNumHijos + 1
                    Set obj = mcolDetalleFamilia.item(i)
                    If obj.IdInventario = 0 Then
                            obj.Recuperar obj.CodInventario
                            If verificacionDuiplicado(obj.IdInventario) Then
                                RemoveDetalleFamilia i
                                i = i - 1
        '                        band = False
                           End If
                    End If
                    If band Then
                        rs1.AddNew
                        If Me.Tipo <> Familia Then
                            rs1!IdInventario = Me.IdInventario
                            rs1!idMateria = obj.IdInventario
                        Else
                            rs1!IdInventario = obj.IdInventario
                            rs1!idMateria = Me.IdInventario
                        End If
                        rs1!TarifaJornal = obj.TarifaJornal
                        rs1!Rendimiento = obj.Rendimiento
                        rs1!Cantidad = obj.Cantidad
                        rs1!BandPrincipal = obj.BandPrincipal
                        rs1!BandModificar = obj.BandModificar 'AUC 25/05/06
                        rs1!Orden = obj.Orden
                        If Me.Tipo = Promocion Then
                            rs1!xCuanto = obj.xCuanto 'AUC 26/09/007
                            rs1!TarifaJornal = obj.Precio  'AUC 26/09/007
                        End If
                        rs1.Update
                        If Me.Tipo = rubro Then 'graba el costoultimoIngreso
                            sql = " Update ivinventario  " & _
                                  " Set costoultimoingreso = " & obj.TarifaJornal & _
                                  " Where codinventario = '" & obj.CodInventario & "'"
                                    mProps.objEmpresa.Execute sql, True
                        End If
                   End If
                End If
                band = True
            Next i
        End If
    Else
        band = True
        If Not mcolDetalleFamilia1 Is Nothing Then
            sql = "SELECT * FROM IVMateria WHERE 1=0 and tipo=1"
            Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
            NumHijos = mcolDetalleFamilia1.Count
            CountNumHijos = 0
            For i = 1 To mcolDetalleFamilia1.Count
    '            rs1.AddNew /jeaa 28-01-04
                If CountNumHijos < NumHijos Then
                    CountNumHijos = CountNumHijos + 1
                    Set obj = mcolDetalleFamilia.item(i)
                    If obj.IdInventario = 0 Then
                            obj.Recuperar obj.CodInventario
                            If verificacionDuiplicado(obj.IdInventario) Then
                                RemoveDetalleFamilia i
                                i = i - 1
        '                        band = False
                           End If
                    End If
                    If band Then
                        rs1.AddNew
                        If Me.Tipo <> Familia Then
                            rs1!IdInventario = Me.IdInventario
                            rs1!idMateria = obj.IdInventario
                        Else
                            rs1!IdInventario = obj.IdInventario
                            rs1!idMateria = Me.IdInventario
                        End If
                        rs1!TarifaJornal = obj.TarifaJornal
                        rs1!Rendimiento = obj.Rendimiento
                        rs1!Cantidad = obj.Cantidad
                        rs1!BandPrincipal = obj.BandPrincipal
                        rs1!BandModificar = obj.BandModificar 'AUC 25/05/06
                        rs1!Orden = obj.Orden
                        If Me.Tipo = Promocion Then
                            rs1!xCuanto = obj.xCuanto 'AUC 26/09/007
                            rs1!TarifaJornal = obj.Precio  'AUC 26/09/007
                        End If
                        rs1.Update
                        If Me.Tipo = rubro Then 'graba el costoultimoIngreso
                            sql = " Update ivinventario  " & _
                                  " Set costoultimoingreso = " & obj.TarifaJornal & _
                                  " Where codinventario = '" & obj.CodInventario & "'"
                                    mProps.objEmpresa.Execute sql, True
                        End If
                    End If
                End If
                band = True
            Next i
        End If
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

Private Sub CargaDetalleFamilia1()
    Dim rs As Recordset, sql As String, obj As IVFamiliaDetalle
    Dim i As Long
    Set mcolDetalleFamilia1 = New Collection
    If Me.Tipo1 = Preparacion Then
       sql = "Select IV.IdInventario, Iv.CodInventario, Iv.Descripcion, " & _
              "Iv.CodUnidad, Ivm.Cantidad, ivm.bandPrincipal, ivm.bandModificar ,iv.costoultimoingreso   From VWIvinventarioRECUPERAR Iv " & _
            "INNER JOIN IVMateria Ivm "
    ElseIf Me.Tipo1 = CambioPresentacion Then
        sql = "Select IV.IdInventario, Iv.CodInventario, Iv.Descripcion, " & _
              "Iv.CodUnidad, Ivm.Cantidad  From VWIvinventarioRECUPERAR Iv " & _
           "INNER JOIN IVMateria Ivm "
    ElseIf Me.Tipo1 = Promocion Then
        sql = "Select IV.IdInventario, Iv.CodInventario, Iv.Descripcion, " & _
             " Iv.CodUnidad,Ivm.Cantidad, ivm.xcuanto, ivm.tarifajornal   From VWIvinventarioRECUPERAR Iv " & _
            "INNER JOIN IVMateria Ivm "
    ElseIf Me.Tipo1 = rubro Then
       sql = "Select IV.IdInventario, Iv.CodInventario, Iv.Descripcion, " & _
              "ivm.formula,ivm.bandModificar,ivm.orden From VWIvinventarioRECUPERAR Iv " & _
            "INNER JOIN IVMateria Ivm "
    ElseIf Me.Tipo1 = Preparacionporce Then
       sql = "Select IV.IdInventario, Iv.CodInventario, Iv.Descripcion, " & _
              "Iv.CodUnidad, Ivm.Cantidad, ivm.Rendimiento, ivm.bandPrincipal From VWIvinventarioRECUPERAR Iv " & _
            "INNER JOIN IVMateria Ivm "
    Else
        sql = "Select IV.IdInventario, Iv.CodInventario, Iv.Descripcion, " & _
              "Iv.codunidad, Ivm.Cantidad From  VWIvinventarioRECUPERAR IV " & _
              "INNER JOIN IVMateria Ivm "
    End If
    '03/10/2005
    If Me.Tipo1 = CambioPresentacion _
        Or Me.Tipo1 = Preparacion _
        Or Me.Tipo1 = Promocion _
        Or Me.Tipo1 = rubro Then
        sql = sql & "  ON IV.IdInventario = Ivm.Idmateria"
        sql = sql & " Where ivm.tipo=1 and Ivm.IdInventario = " & mProps.IdInventario
        sql = sql & " order by orden"
    ElseIf Me.Tipo1 = Preparacionporce Then
        sql = sql & "  ON IV.IdInventario = Ivm.Idmateria"
        sql = sql & " Where ivm.tipo=1 and Ivm.IdInventario = " & mProps.IdInventario
        sql = sql & " AND IV.TIPO<>2 "
        sql = sql & " order by orden"
    Else
          sql = sql & " ON IV.IdInventario = Ivm.IdInventario "
            sql = sql & " Where ivm.tipo=1 and Ivm.IdMateria = " & mProps.IdInventario
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddDetalleFamilia
            Set obj = mcolDetalleFamilia.item(i)
            If Me.Tipo1 = rubro Then
               obj.IdInventario = !IdInventario
                obj.CodInventario = !CodInventario
                obj.Descripcion = !Descripcion
                If Not IsNull(!Formula) Then obj.Formula = !Formula
                obj.BandModificar = !BandModificar
                obj.Orden = !Orden
            Else
                  obj.IdInventario = !IdInventario
                  obj.CodInventario = !CodInventario
                  obj.Descripcion = !Descripcion
                  If Not IsNull(!CodUnidad) Then obj.Unidad = !CodUnidad 'AUC 'cambiado antes unidad
                 obj.Cantidad = !Cantidad
                  If Me.Tipo = Preparacion Then
                      If Len(obj.IVInventario.Empresa.GNOpcion.ObtenerValor("itemPreparacion")) > 0 Then 'AUC toma el p.costo
                          If obj.IVInventario.Empresa.GNOpcion.ObtenerValor("itemPreparacion") = 0 Then 'AUC toma el p.costo
                             If Not IsNull(!CostoUltimoIngreso) Then obj.CostoUltimoIngreso = !CostoUltimoIngreso
                           ElseIf obj.IVInventario.Empresa.GNOpcion.ObtenerValor("itemPreparacion") = 1 Then 'AUC toma el p. venta
                              obj.pVenta = obj.ObtienePrecio(obj.CodInventario)
                           End If
                           obj.BandPrincipal = !BandPrincipal  'AUC 10/05/06 restriccion del item
                           obj.BandModificar = !BandModificar   'AUC 25/05/06
                      Else
                          obj.pVenta = obj.ObtienePrecio(obj.CodInventario)
                      End If
                  ElseIf Me.Tipo = Promocion Then
                          obj.xCuanto = !xCuanto
                          obj.Precio = !TarifaJornal
                  ElseIf Me.Tipo = Preparacionporce Then
                        obj.BandPrincipal = !BandPrincipal  'AUC 10/05/06 restriccion del item
                        obj.Rendimiento = !Rendimiento 'AUC 10/05/06 restriccion del item
                End If
            End If
           .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
   mbooEsNuevo = False
    mbooModificado = False
End Sub
Public Sub RemoveDetalleFamiliaNew(ByVal ix As Long, Optional obj As IVFamiliaDetalle, Optional Tipo As Integer)
    If Tipo = 0 Then
        If mcolDetalleFamilia Is Nothing Then CargaDetalleFamilia
        'Si está especificado un objeto
        If Not (obj Is Nothing) Then
            'Busca el indice del objeto
            For ix = 1 To mcolDetalleFamilia.Count
                If mcolDetalleFamilia.item(ix) Is obj Then
                    mcolDetalleFamilia.Remove ix
                    Exit For
                End If
            Next ix
            Debug.Assert True     'No se puede dar
        Else
            mcolDetalleFamilia.Remove ix
        End If
    Else
        If mcolDetalleFamilia1 Is Nothing Then CargaDetalleFamilia1
        'Si está especificado un objeto
        If Not (obj Is Nothing) Then
            'Busca el indice del objeto
            For ix = 1 To mcolDetalleFamilia1.Count
                If mcolDetalleFamilia1.item(ix) Is obj Then
                    mcolDetalleFamilia1.Remove ix
                    Exit For
                End If
            Next ix
            Debug.Assert True     'No se puede dar
        Else
            mcolDetalleFamilia1.Remove ix
        End If
    End If
    mbooModificado = True
End Sub
Public Property Let Tipo1(ByVal value As EIVTipo)
    mProps.Tipo1 = value
    mbooModificado = True
End Property
Public Property Get Tipo1() As EIVTipo
    Tipo1 = mProps.Tipo1
End Property
Public Property Get NumFamiliaDetalle1() As Long
    If mcolDetalleFamilia1 Is Nothing Then
        NumFamiliaDetalle1 = 0
    Else
        NumFamiliaDetalle1 = mcolDetalleFamilia1.Count
    End If
End Property
Public Function RecuperaDetalleFamilia1(ix As Long) As IVFamiliaDetalle
    Set RecuperaDetalleFamilia1 = mcolDetalleFamilia1.item(ix)
End Function

'Public Sub AsignarBodegaPre()
'Dim Bodega As String, SQL As String
'Dim idbodega  As Long, RS As Recordset
'If Len(gobjMain.EmpresaActual.GNOpcion.ObtenerValor("BodegaPredeterminada")) > 0 Then
'    Bodega = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("BodegaPredeterminada")
'Else
'    Err.Raise Err.Number, "IVInventario", "No hay una bodega prederminada"
'End If
'SQL = "SELECT IdBodega FROM IVBodega WHERE codbodega='" & Bodega & "'"
'
'Set RS = gobjMain.EmpresaActual.OpenRecordset(SQL)
'    'Si no existe el registro correspondiente, lo inserta
'    If Not RS.EOF Then
'        SQL = "INSERT INTO IVExist (IdInventario,IdBodega,Exist) " & _
'              "VALUES(" & Me.RecuperaID(mProps.CodInventario) & "," & RS!idbodega & "," & 0 & ")"
'        gobjMain.EmpresaActual.Execute SQL, True
'    End If
'    RS.Close
'    Set RS = Nothing
'End Sub
'

Private Sub GrabaRubroDetalle(ByVal CodInventario As String) 'no se si se utiliza en algun lado revisar sino borrar
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As IVFamiliaDetalle, sql As String, band As Boolean
    'Dim iv As IVInventario
    'Set iv = Me.Empresa.RecuperaIvInventario(CodInventario)
    On Error GoTo CapturaError
    'Elimina  todos  los registros
    sql = "Delete Ivm " & _
          "From ivmateria Ivm " & _
          "where ivm.idinventario = " & RecuperaID(CodInventario)
    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    band = True
    sql = "SELECT * FROM IVMateria WHERE 1=0"
    Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
    If Not mcolDetalleFamilia Is Nothing Then
        For i = 1 To mcolDetalleFamilia.Count
'            rs1.AddNew /jeaa 28-01-04
            Set obj = mcolDetalleFamilia.item(i)
            If obj.IdInventario = 0 Then
                        obj.Recuperar obj.CodInventario
                        If verificacionDuiplicado(obj.IdInventario) Then
                            RemoveDetalleFamilia i
                            band = False
                        End If
            End If
            If band Then
                rs1.AddNew
                If Me.Tipo = rubro Then
                    rs1!IdInventario = RecuperaID(CodInventario) 'mProps.IdInventario
                    rs1!idMateria = obj.IdInventario
                Else
                    rs1!IdInventario = obj.IdInventario
                    rs1!idMateria = Me.IdInventario
                End If
                rs1!TarifaJornal = obj.TarifaJornal
              '  rs1!costoUltimoIngeso = obj.CostoUltimoIngreso
                rs1!Rendimiento = obj.Rendimiento
                rs1!Cantidad = obj.Cantidad
                rs1!BandPrincipal = obj.BandPrincipal
                rs1!BandModificar = obj.BandModificar 'AUC 25/05/06
                'rs1!Formula = obj.Formula  'AUC 25/05/06
                rs1!Orden = obj.Orden
                rs1.Update
'                If Me.tipo = rubro Then 'graba el costoultimoIngreso
'                    sql = " Update ivinventario  " & _
'                          " Set costoultimoingreso = " & obj.CostoUltimoIngreso & _
'                          " Where codinventario = '" & obj.CodInventario & "'"
'                            mProps.objEmpresa.Execute sql, True
'                End If
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
Public Property Let Utilidad(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario.Utilidad", _
            "la utilidad no puede ser negativo. ('" & mProps.CodInventario & "')"
    End If
    'Si es que realmente modifica el valor
    If mProps.Utilidad <> value Then
        mProps.Utilidad = value
      mbooModificado = True
    End If
End Property
Public Property Get Utilidad() As Currency
    Utilidad = mProps.Utilidad
End Property
Public Property Get idproceso() As Long
    idproceso = mProps.idproceso
End Property
Public Property Let idproceso(ByVal value As Long)
    mProps.idproceso = value
    mbooModificado = True
End Property
Public Property Let CodProceso(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodProceso Then
        If Len(value) > 0 Then
            'Actualiza IdGrupo1 también
            sql = "SELECT IdProceso FROM IVProceso  WHERE CodProceso ='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVInventario_CodProceso", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                idproceso = rs!idproceso
            End If
            rs.Close
        Else
            idproceso = 0
       End If
    End If
    Set rs = Nothing
    mProps.CodProceso = value
    mbooModificado = True
End Property
Public Property Get CodProceso() As String
    CodProceso = Trim$(mProps.CodProceso)
End Property
Private Sub CargaRubroDetalle() 'no se si se utiliza en algun lado revisar sino borrar
    Dim rs As Recordset, sql As String, obj As IVFamiliaDetalle
    Dim i As Long
    Set mcolDetalleFamilia = New Collection
       sql = "Select IV.IdInventario, Iv.CodInventario, Iv.Descripcion, " & _
              "ivm.cantidad,iv.codunidad,iv.costoultimoingreso,iv.costoreferencial,ivm.rendimiento,ivm.bandModificar,bandPrincipal,ivm.orden From VWIvinventarioRECUPERAR Iv " & _
            "INNER JOIN IVMateria Ivm "
       sql = sql & "  ON IV.IdInventario = Ivm.Idmateria"
       sql = sql & " Where Ivm.IdInventario = " & mProps.IdInventario
       sql = sql & " order by orden"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddDetalleFamilia
            Set obj = mcolDetalleFamilia.item(i)
                obj.IdInventario = !IdInventario
                obj.CodInventario = !CodInventario
                obj.Descripcion = !Descripcion
                obj.Cantidad = !Cantidad
                If Not IsNull(!CodUnidad) Then obj.Unidad = !CodUnidad 'AUC 'cambiado antes unidad
                If Not IsNull(!CostoUltimoIngreso) Then obj.CostoUltimoIngreso = !CostoUltimoIngreso
                If Not IsNull(!CostoReferencial) Then obj.CostoReferencial = !CostoReferencial
                'If Not IsNull(!Formula) Then obj.Formula = !Formula
                obj.BandModificar = !BandModificar
                obj.BandPrincipal = !BandPrincipal
                obj.Orden = !Orden
           .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
   mbooEsNuevo = False
    mbooModificado = False
End Sub

'saca la existencia del item importante
Public Function ExistenciaPorItemImportante(ByVal IdBodega As Long, ByVal codItem As String) As Currency
    Dim rs As Recordset
    Dim sql As String
    VerificaExistenciaTabla 99 'tmp99 sera utilizada para guardar la primer parte del sql
    sql = "SELECT   exist,ivm.idinventario,ivm.bandprincipal,ivm.idmateria,iv.tipo " & _
           "Into tmp99 " & _
           "FROM IvInventario iv  LEFT Join ivexist  INNER JOIN  ivbodega  on ivexist.IdBodega = ivbodega.IDBODEGA " & _
           "ON Iv.IdInventario = ivexist.IdInventario  left join ivmateria ivm on ivm.idinventario = iv.idinventario " & _
           "Where iv.tipo = 4 and ivm.bandprincipal = 1 " & _
           "AND iv.codinventario =  '" & codItem & "'" & _
           "AND ivexist.IdBodega =" & IdBodega
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    ''VerificaExistenciaTabla 100
    sql = "Select  min(ive.exist) as exist " & _
         " from tmp99 t inner join ivexist ive on ive.idinventario = t.idmateria  WHERE  ive.idbodega = " & IdBodega
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'sql = " SELECT   min(exist) as exist from tmp100 "
    'Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If Not IsNull(rs!exist) Then
        ExistenciaPorItemImportante = rs!exist
        Exit Function
    End If
   Set rs = Nothing
End Function



Public Property Let BandAltaRotacion(ByVal value As Boolean)
    mProps.BandAltaRotacion = value
    mbooModificado = True
End Property

Public Property Get BandAltaRotacion() As Boolean
    BandAltaRotacion = mProps.BandAltaRotacion
End Property

Public Property Let CostoAdicionalUltimoIngreso(ByVal value As Currency)
    mProps.CostoAdicionalUltimoIngreso = value
    mbooModificado = True
End Property

Public Property Get CostoAdicionalUltimoIngreso() As Currency
    If ObtieneUltimoCompra(True) Then CostoAdicionalUltimoIngreso = mUIngreso.Comision
    'CostoUltimoIngreso = mProps.CostoUltimoIngreso
End Property


Public Property Let BandNoObjetoIVA(ByVal value As Boolean)
    mProps.BandNoObjetoIVA = value
    mbooModificado = True
End Property

Public Property Get BandNoObjetoIVA() As Boolean
    BandNoObjetoIVA = mProps.BandNoObjetoIVA
End Property

Public Function VerificaSiEstaPromoDesc(ByRef Id As Long, ByRef cad As String) As Boolean
   Dim sql As String, i As Integer
    Dim rs As Recordset
    VerificaSiEstaPromoDesc = False
    If Not mbooEsNuevo Then
    '    'Verifica si no esta en IVMateria
        sql = "SELECT idInventario FROM IVmateria " & _
              "WHERE IdInventario=" & Id & " "
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
       If Not rs.EOF Then      'Si ya existe el codigo
            VerificaSiEstaPromoDesc = True
            cad = " está relacionado de la tabla de IVMateria"
            Exit Function
        End If
        
    '    'Verifica si no esta en IVDescuentoDetalleIV
        sql = "SELECT idInventario FROM IVDescuentoDetalleIV " & _
              "WHERE IdInventario=" & Id & " "
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
       If Not rs.EOF Then      'Si ya existe el codigo
            VerificaSiEstaPromoDesc = True
            cad = " está relacionado en Descuentos"
            Exit Function
        End If
        
    '    'Verifica si no esta en IvCondPromocionDetalleIVG
        sql = "SELECT idInventario FROM IVCondPromocionDetalle " & _
              "WHERE IdInventario=" & Id & " "
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
       If Not rs.EOF Then      'Si ya existe el codigo
            VerificaSiEstaPromoDesc = True
            cad = " está relacionado en Promociones"
            Exit Function
        End If
        
    '    'Verifica si no esta TIENE EXISTENCIA
        sql = "SELECT idInventario FROM IVExist " & _
              "WHERE exist <>0 and IdInventario=" & Id & " "
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
       If Not rs.EOF Then      'Si ya existe el codigo
            VerificaSiEstaPromoDesc = True
            cad = " está con existencia "
            Exit Function
        End If
        
        
    End If
End Function

Public Property Get BandGarantia() As Boolean
    BandGarantia = mProps.BandGarantia
End Property

Public Property Let BandGarantia(ByVal value As Boolean)
    mProps.BandGarantia = value
    mbooModificado = True
End Property

Public Property Let IdCuentaDiferida(ByVal value As Long)
    mProps.IdCuentaDiferida = value
    mbooModificado = True
End Property

Public Property Get IdCuentaDiferida() As Long
    IdCuentaDiferida = mProps.IdCuentaDiferida
End Property

Public Property Let CodCuentaDiferida(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodCuentaDiferida Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaDiferida también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVInDiferidario_CodCuentaDiferida", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaDiferida = rs!IdCuenta
            End If
            rs.Close
        Else
            IdCuentaDiferida = 0
        End If
        mProps.CodCuentaDiferida = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property

Public Property Get CodCuentaDiferida() As String
    CodCuentaDiferida = Trim$(mProps.CodCuentaDiferida)
End Property

Public Property Let TipoAuto(ByVal value As Integer)
   mProps.TipoAuto = value
    mbooModificado = True
End Property

Public Property Get TipoAuto() As Integer 'antes estaba tipo1 para no duplicar
   TipoAuto = mProps.TipoAuto
End Property

Public Property Get BandCambiaDescr() As Boolean
    BandCambiaDescr = mProps.BandCambiaDescr
End Property

Public Property Let BandCambiaDescr(ByVal value As Boolean)
    mProps.BandCambiaDescr = value
    mbooModificado = True
End Property

Friend Sub RecuperarQuickId(key As Variant)
    Dim sql As String, i As Integer
    Dim rs As Recordset
    
        sql = "SELECT * FROM vwIVInventarioRecuperar WHERE IdInventario=" & key
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IdInventario = !IdInventario
        mProps.CodInventario = !CodInventario
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        'For i = 1 To 6 'AUC cambiado antes 4
        For i = 1 To 7 'GRV Precio 7 - 04MAY2014
            If Not IsNull(.Fields("Precio" & i)) Then mProps.Precio(i) = .Fields("Precio" & i)
            If Not IsNull(.Fields("Comision" & i)) Then mProps.Comision(i) = .Fields("Comision" & i)
            If Not IsNull(.Fields("CantLimite" & i)) Then mProps.CantLimite(i) = .Fields("CantLimite" & i)
            If Not IsNull(.Fields("Descuento" & i)) Then mProps.Descuento(i) = .Fields("Descuento" & i)   '***Agregado. jeaa. 29/jul/2004
        Next i
        For i = 1 To IVGRUPO_MAX
            If Not IsNull(.Fields("IdGrupo" & i)) Then mProps.IdGrupo(i) = .Fields("IdGrupo" & i)
            If Not IsNull(.Fields("CodGrupo" & i)) Then mProps.CodGrupo(i) = .Fields("CodGrupo" & i)
        Next i
        
        If Not IsNull(!PorcentajeIVA) Then mProps.PorcentajeIVA = !PorcentajeIVA
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!BandServicio) Then mProps.BandServicio = !BandServicio
        If Not IsNull(!CostoUltimoIngreso) Then mProps.CostoUltimoIngreso = !CostoUltimoIngreso 'AUC 22/12/05
        If Not IsNull(!BandIVA) Then mProps.BandIVA = !BandIVA  'jeaa 08/03/2006
        If Not IsNull(!Tipo) Then mProps.Tipo = !Tipo  'jeaa 08/03/2006
        If Not IsNull(!CostoReferencial) Then mProps.CostoReferencial = !CostoReferencial 'AUC 22/12/05
        If Not IsNull(!PorDesperdicio) Then mProps.PorDesperdicio = !PorDesperdicio
        If Not IsNull(!RutaFoto) Then mProps.RutaFoto = !RutaFoto
        If Not IsNull(!CodArancel) Then mProps.CodArancel = !CodArancel
        If Not IsNull(!BandNoObjetoIVA) Then mProps.BandNoObjetoIVA = !BandNoObjetoIVA
        If Not IsNull(!BandGarantia) Then mProps.BandGarantia = !BandGarantia
        If Not IsNull(!BandCambiaDescr) Then mProps.BandCambiaDescr = !BandCambiaDescr
        If Not IsNull(!TipoAuto) Then mProps.TipoAuto = !TipoAuto
        If Not IsNull(!bandSerie) Then mProps.bandSerie = !bandSerie
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Property Get bandSerie() As Boolean
    bandSerie = mProps.bandSerie
End Property

Public Property Let bandSerie(ByVal value As Boolean)
    mProps.bandSerie = value
    mbooModificado = True
End Property

Public Sub GrabarPrecioMatriz()
    Dim sql As String
    Dim rs As Recordset
    If Not mbooModificado Then Exit Sub
    sql = "Update dbo.IVInventario " & _
          "SET precio1 = " & mProps.Precio(1) & ", " & _
              "precio2 = " & mProps.Precio(2) & ", " & _
              "precio3 = " & mProps.Precio(3) & ", " & _
              "precio4 = " & mProps.Precio(4) & ", " & _
              "precio5 = " & mProps.Precio(5) & ", " & _
              "precio6 = " & mProps.Precio(6) & ", " & _
              "precio7 = " & mProps.Precio(7) & ", " & _
              "FechaGrabado = '" & Now & "' " & _
          "WHERE CodInventario like '" & mProps.CodInventario & "%'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    Set rs = Nothing
    'rs.Close
    mProps.objEmpresa.GrabaGNLogAccion "MOD_ITEM", "Item Modificado: " & _
                                                         mProps.CodInventario & " Y todos los hijos", "DT"
End Sub


Public Function RecuperaPrecioCompraProv(ByVal item As IVInventario, ByVal CodProveedor As String) As Currency
    Dim PrecioCompra As Currency
    Dim i As Integer
    Dim obj As IVDetalleProveedor
    For i = 1 To item.NumProveedorDetalle
        Set obj = item.RecuperaDetalleProveedor(i)
        If obj.CodProveedor = CodProveedor Then
            PrecioCompra = obj.PrecioCompra
        End If
    Next
    RecuperaPrecioCompraProv = PrecioCompra
End Function


Public Function CostoxProveedor( _
                    ByVal CodProvcli As String, _
                    ByVal codItem As String) As Double

    Dim sql As String, num As Integer, campo As String
    Dim rs As Recordset, Costo As Double        '*** MAKOTO 23/nov/00 Modificado a Double
    Dim idItem As Long
    
    Costo = 0

    sql = " select PrecioCompra from ivproveedordetalle ivp"
    sql = sql & " inner join pcprovcli pc on ivp.idproveedor=pc.idprovcli"
    sql = sql & " inner join ivinventario iv on ivp.idinventario=iv.idinventario"
    sql = sql & " where codprovcli='" & CodProvcli & "'"
    sql = sql & " and codinventario='" & codItem & "'"
    
    Set rs = Me.Empresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        Costo = rs.Fields("PrecioCompra")
    End If
    rs.Close
    Set rs = Nothing

    'Si no hay costo promedio por que no hay existencia,
    '   devuelve el último costo en vez de promedio
    'If Costo = 0 Then Costo = CostoUltimo(fecha, tid, Hora, False)
    
    CostoxProveedor = Costo


End Function



Public Property Let FrecuenciaReposicion(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario.FrecuenciaReposicion", _
            "El FrecuenciaReposicion no puede ser negativo. ('" & mProps.CodInventario & "')"
    End If

    If mProps.FrecuenciaReposicion <> value Then
        mProps.FrecuenciaReposicion = value
        mbooModificado = True
    End If
End Property

Public Property Get FrecuenciaReposicion() As Currency
    FrecuenciaReposicion = mProps.FrecuenciaReposicion
End Property


Public Property Let Buffer(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario.Buffer", _
            "El Buffer no puede ser negativo. ('" & mProps.CodInventario & "')"
    End If

    If mProps.Buffer <> value Then
        mProps.Buffer = value
        mbooModificado = True
    End If
End Property

Public Property Get Buffer() As Currency
    Buffer = mProps.Buffer
End Property

Public Property Let FechaModBuffer(value As Date)
    If Not (IsDate(value)) Then
        Err.Raise ERR_INVALIDO, "PCProvCli_FechaModBuffer", MSGERR_INVALIDO
    End If
    
    mProps.FechaModBuffer = value
    mbooModificado = True
End Property

Public Property Get FechaModBuffer() As Date
    FechaModBuffer = mProps.FechaModBuffer
End Property

Public Property Get bandPeso() As Boolean
    bandPeso = mProps.bandPeso
End Property

Public Property Let bandPeso(ByVal value As Boolean)
    mProps.bandPeso = value
    mbooModificado = True
End Property

Public Property Get IdPlan() As Long ' AUC revisar si se va ha necesitar sino borrar
    IdPlan = mProps.IdPlan
End Property

Public Property Get CodPlan() As String
    CodPlan = mProps.CodPlan
End Property

Public Property Let CodPlan(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodPlan Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaDiferida también
            sql = "SELECT IdPlan FROM ivplan WHERE codplan='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVInDiferidario_CodCuentaDiferida", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdPlan = rs!IdPlan
            End If
            rs.Close
        Else
            IdPlan = 0
        End If
        mProps.CodPlan = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property

Public Property Let IdPlan(ByVal value As Long)
    mProps.IdPlan = value
   mbooModificado = True
End Property

Public Property Let FrecuenciaReposicionALM(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario.FrecuenciaReposicionALM", _
            "El Frecuencia Reposicion ALM no puede ser negativo. ('" & mProps.CodInventario & "')"
    End If

    If mProps.FrecuenciaReposicionALM <> value Then
        mProps.FrecuenciaReposicionALM = value
        mbooModificado = True
    End If
End Property

Public Property Get FrecuenciaReposicionALM() As Currency
    FrecuenciaReposicionALM = mProps.FrecuenciaReposicionALM
End Property


Public Property Let TiempoReposicionALM(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario.TiempoReposicionALM", _
            "El Tiempo reposicion almacen no puede ser negativo. ('" & mProps.CodInventario & "')"
    End If

    If mProps.TiempoReposicionALM <> value Then
        mProps.TiempoReposicionALM = value
        mbooModificado = True
    End If
End Property

Public Property Get TiempoReposicionALM() As Currency
    TiempoReposicionALM = mProps.TiempoReposicionALM
End Property


Public Sub GrabaExistenciaxBodegaNew(ByVal IdInventario As Long, ByVal IdBodega As Integer, ByVal min As Integer, ByVal max As Integer, TiempoRep As Currency, FrecuRep As Currency, Buffer As Currency)
    Dim sql As String, rs As Recordset

    sql = "update ivexist "
    sql = sql & " set existmin=" & min & ", existmax=" & max
    sql = sql & ", TiempoRep=" & TiempoRep & ", FrecuRep=" & FrecuRep
    sql = sql & ", Buffer=" & Buffer
    sql = sql & " from ivexist"
    sql = sql & " where idinventario= " & IdInventario
    sql = sql & " and idbodega=" & IdBodega
    Set rs = Me.Empresa.OpenRecordset(sql)
    Set rs = Nothing
End Sub


Public Property Let BandOmitirRebate(ByVal value As Boolean)
    mProps.BandOmitirRebate = value
    mbooModificado = True
End Property

Public Property Get BandOmitirRebate() As Boolean
    BandOmitirRebate = mProps.BandOmitirRebate
End Property

Public Function ExistenciaReceta(ByRef idInv As Long, ByVal CodBodega As String) As Currency
    Dim rs As Recordset
    Dim sql As String
    sql = "select exist from ivexist inner join ivinventario iv " & _
           "on iv.idinventario = ivexist.idinventario " & _
           "inner join ivbodega ivb on ivb.idbodega = ivexist.idbodega " & _
           " Where iv.IdInventario = " & idInv & _
           " And ivb.codbodega = '" & CodBodega & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not IsNull(rs) Then
        ExistenciaReceta = rs!exist
            Exit Function
        End If
   Set rs = Nothing
End Function
Private Function CostoUltimoDou( _
                    ByVal fecha As Date, _
                    ByVal tid As Long, _
                    ByVal Hora As Date, _
                    ByVal BandSoloIngreso As Boolean) As Double
    Dim sql As String, campo As String, num As Integer
    Dim rs As Recordset, cadena_hora As String
    num = Me.Empresa.GNOpcion.IndiceMoneda(Me.CodMoneda)
    campo = "1"
    If num > 1 Then campo = "gc.Cotizacion" & num
    'Obtiene IVKardex ordenado por la fecha descendiente para coger último ingreso
    ' Busca solo de FechaTrans<=Fecha, y solo de INGRESOS y no anulados
    sql = "SELECT ivk.CostoRealTotalDou/ivk.CantidadDou/" & campo & " AS CU " & _
          "FROM GNTrans gt INNER JOIN " & _
                    "(GNComprobante gc INNER JOIN IVKardex ivk " & _
                    "ON gc.TransID = ivk.TransID) " & _
                "ON gt.CodTrans = gc.CodTrans "
    sql = sql & "WHERE (gc.Estado <> 3) AND " & _
                    "(gt.AfectaCantidad=" & CadenaBool(True, Me.Empresa.TipoDB) & ") AND " & _
                    "(ivk.IdInventario=" & mProps.IdInventario & ") "
    If BandSoloIngreso Then sql = sql & " AND (ivk.CantidadDou>0) "
    'Condicion de fecha y tid
    '   Hasta antes de la fecha limite
    '   o la misma fecha pero la hora menor
    '   o la misma fecha y misma hora pero tid menor
'    cadena_hora = Format(Hora, "Hh:Nn:Ss")
    cadena_hora = HoraHMS(Hora, Me.Empresa.TipoDB)
    sql = sql & " AND ( (gc.FechaTrans<" & FechaYMD(fecha, Me.Empresa.TipoDB) & ") " & _
                    "OR (gc.FechaTrans=" & FechaYMD(fecha, Me.Empresa.TipoDB) & _
                        " AND gc.HoraTrans<" & cadena_hora & ") " & _
                    "OR (gc.FechaTrans=" & FechaYMD(fecha, Me.Empresa.TipoDB) & _
                        " AND gc.HoraTrans=" & cadena_hora & _
                        " AND gc.TransID < " & tid & ") )"
    sql = sql & " ORDER BY gc.FechaTrans DESC, gc.HoraTrans DESC, gc.TransID DESC"     '*** MAKOTO 08/dic/00 HoraTrans Agregado
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'Coge el primer costo como el costo ultimo
    If Not rs.EOF Then
        If Not IsNull(rs.Fields("CU")) Then
            CostoUltimoDou = rs.Fields("CU")
        'Aseguro de devolver 0 porque no encontro mov. de item
        Else
            CostoUltimoDou = 0
        End If
    Else
        CostoUltimoDou = 0
    End If
    rs.Close
    Set rs = Nothing
End Function

Public Property Get Batch() As Currency
    Batch = mProps.Batch
End Property

Public Property Let Batch(ByVal value As Currency)
    mProps.Batch = value
    mbooModificado = True
End Property

Public Property Get BatchProd() As Currency
    BatchProd = mProps.BatchProd
End Property

Public Property Let BatchProd(ByVal value As Currency)
    mProps.BatchProd = value
    mbooModificado = True
End Property

Public Property Get BandNoFormula() As Boolean
    BandNoFormula = mProps.BandNoFormula
End Property

Public Property Let BandNoFormula(ByVal value As Boolean)
    mProps.BandNoFormula = value
    mbooModificado = True
End Property

Public Function ExistenciaRecetaDou(ByRef idInv As Long, ByVal CodBodega As String) As Double
    Dim rs As Recordset
    Dim sql As String
    sql = "select existDou from ivexist inner join ivinventario iv " & _
           "on iv.idinventario = ivexist.idinventario " & _
           "inner join ivbodega ivb on ivb.idbodega = ivexist.idbodega " & _
           " Where iv.IdInventario = " & idInv & _
           " And ivb.codbodega = '" & CodBodega & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        If Not IsNull(rs) Then
            ExistenciaRecetaDou = rs!existDou
                Exit Function
            End If
        Set rs = Nothing
    End If
End Function

Public Function ListaDetalleConversion2() As String
    Dim sql As String, rs As Recordset
    Dim s As String
        sql = "SELECT u.codunidad,ivc.cantidadDou,u1.codunidad "
        sql = sql & " FROM ivconversion ivc inner join ivinventario iv"
        sql = sql & " on iv.idinventario = ivc.idinventario"
        sql = sql & " inner join ivunidad u on u.idunidad = ivc.idunidadori"
        sql = sql & " inner join ivunidad u1 on u1.idunidad = ivc.idunidaddes"
        sql = sql & " Where iv.codInventario = '" & Me.CodInventario & "'"
        Set rs = Me.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then s = "|" & rs.GetString(adClipString, , vbTab, "|", "(nulo)")
        Set rs = Nothing
        ListaDetalleConversion2 = s
End Function


Public Property Let CantReluni(ByVal value As Currency)
    mProps.CantReluni = value
    mbooModificado = True
End Property
Public Property Get CantReluni() As Currency
    CantReluni = mProps.CantReluni
End Property

Public Property Let CantReluniCont(ByVal value As Currency)
    mProps.CantReluniCont = value
    mbooModificado = True
End Property

Public Property Get CantReluniCont() As Currency
    CantReluniCont = mProps.CantReluniCont
End Property

Public Property Let IdCuentaDevolucion(ByVal value As Long)
    mProps.IdCuentaDevolucion = value
    mbooModificado = True
End Property

Public Property Get IdCuentaDevolucion() As Long
    IdCuentaDevolucion = mProps.IdCuentaDevolucion
End Property

Public Property Let CodCuentaDevolucion(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodCuentaDevolucion Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaDevolucion también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVInventario_CodCuentaDevolucion", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaDevolucion = rs!IdCuenta
            End If
            rs.Close
        Else
            IdCuentaDevolucion = 0
        End If
        mProps.CodCuentaDevolucion = value
        mbooModificado = True
    End If
    Set rs = Nothing
End Property

Public Property Get CodCuentaDevolucion() As String
    CodCuentaDevolucion = Trim$(mProps.CodCuentaDevolucion)
End Property

Public Property Let CantReuni(ByVal value As Currency)
    mProps.CantReuni = value
    mbooModificado = True
End Property

Public Property Get CantReuni() As Currency
    CantReuni = mProps.CantReuni
End Property

Public Property Let CantReuniCont(ByVal value As Currency)
    mProps.CantReuniCont = value
    mbooModificado = True
End Property

Public Property Get CantReuniCont() As Currency
    CantReuniCont = mProps.CantReuniCont
End Property

Public Property Let M3(ByVal value As Double)
    If mProps.M3 <> value Then
        mProps.M3 = value
        mbooModificado = True
    End If
End Property

Public Property Get M3() As Double
    M3 = mProps.M3
End Property

Public Property Get BandGarantiaC() As Boolean
    BandGarantiaC = mProps.BandGarantiaC
End Property

Public Property Let BandGarantiaC(ByVal value As Boolean)
    mProps.BandGarantiaC = value
    mbooModificado = True
End Property

Public Property Get BandGarantiaV() As Boolean
    BandGarantiaV = mProps.BandGarantiaV
End Property

Public Property Let BandGarantiaV(ByVal value As Boolean)
    mProps.BandGarantiaV = value
    mbooModificado = True
End Property

Public Property Get NumGarantia() As Long
    If mcolGarantia Is Nothing Then
        NumGarantia = 0
    Else
        NumGarantia = mcolGarantia.Count
    End If
End Property

Public Function RecuperaGarantia(ix As Integer) As IVGarantia
    Set RecuperaGarantia = mcolGarantia.item(ix)
End Function

Private Sub CargaGarantia()
    Dim rs As Recordset, sql As String, obj As IVGarantia
    Dim i As Long
    Set mcolGarantia = New Collection
    sql = " Select * from IVGarantia  Where IdInventario = " & Me.IdInventario
    sql = sql & " order by id"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddGarantia
            Set obj = mcolGarantia.item(i)
'            obj.IdProvCli = !IdProvCli
            If Not IsNull(!FechaDesde) Then obj.FechaDesde = !FechaDesde
            If Not IsNull(!FechaHasta) Then obj.FechaHasta = !FechaHasta
            If Not IsNull(!Valor) Then obj.Valor = !Valor
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub


Public Function AddGarantia() As Long
    Dim nuevo As IVGarantia
    If mcolGarantia Is Nothing Then CargaGarantia
    Set nuevo = New IVGarantia
    Set nuevo.IVInventario = Me
    mcolGarantia.Add nuevo
    AddGarantia = mcolGarantia.Count     'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function

Public Sub RemoveGarantia(ByVal ix As Long, Optional obj As IVGarantia)
    If mcolGarantia Is Nothing Then CargaGarantia
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolGarantia.Count
            If mcolGarantia.item(ix) Is obj Then
                mcolGarantia.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolGarantia.Remove ix
    End If
    'mcolDetalleFamilia.Remove ix
    mbooModificado = True
End Sub

Private Sub GrabarIVGarantia()
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj  As IVGarantia, sql As String, band As Boolean
    'Elimina  todos  los registros
    sql = "Delete From ivGarantia where idinventario =" & Me.IdInventario
    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    band = True
    If Not mcolGarantia Is Nothing Then
        sql = "SELECT * FROM  ivGarantia WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        For i = 1 To mcolGarantia.Count
            Set obj = mcolGarantia.item(i)
'            If obj.IdInventario = 0 Then
''                        obj.Recuperar obj.CodProveedor
''                        If verificacionDuiplicado(obj.idProveedor) Then
''                            RemoveDetalleProveedor i
''                            band = False
''                        End If
'            End If
            If band Then
                rs1.AddNew
                    rs1!IdInventario = Me.IdInventario
                    rs1!FechaDesde = obj.FechaDesde
                    rs1!FechaHasta = obj.FechaHasta
                    rs1!Valor = obj.Valor
                rs1.Update
            End If
            band = True
        Next i
    End If
    Set rs1 = Nothing
    Set obj = Nothing
End Sub

