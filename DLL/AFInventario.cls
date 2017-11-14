VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AFInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Empresa"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Auxiliar para guardar datos de ultimo ingreso/egreso


Private Type auxUltimo
    fecha As Date
    Cantidad As Currency
    Costo As Currency
    Precio As Currency
    Recuperado As Boolean
End Type

Private Type T_PROPS
    IdInventario As Long
    CodInventario As String
    CodAlterno1 As String
    CodAlterno2 As String
    Descripcion As String
    Descripcion2 As String      'jeaa 4/06/2008
    NumSerie As String
    Marca As String
    NumFactura As String
    VidaUtil As String
    DepAnterior As String
    FechaCompra As Date
    Estado As Byte
    tipodepre As Byte
    
    DescripcionDetalle As String        'memo
    CodMoneda As String
    Precio(1 To 5) As Currency 'AUC cambiado antes 4
    PorcentajeIVA As Currency
    PorcentajeIVAAnt As Currency
    IdCuentaActivo As Long
    IdCuentaCosto As Long
    IdCuentaVenta As Long
    IdCuentaDepreGasto As Long
    IdCuentaDepreAcumulada As Long
    IdCuentaRevaloriza As Long
    IdCuentaDepRevaloriza As Long
    
    
    CodCuentaActivo As String
    CodCuentaCosto As String
    CodCuentaVenta As String
    CodCuentaDepreGasto As String
    CodCuentaDepreAcumulada As String
    CodCuentaRevaloriza As String
    CodCuentaDepRevaloriza As String
    
    
    IdGrupo(1 To 5) As Long
    CodGrupo(1 To 5) As String
    idProveedor As Long
    CodProveedor As String
    Observacion As String
    ExistenciaMinima As Currency
    ExistenciaMaxima As Currency
    UnidadMinimaCompra As Currency
    UnidadMinimaVenta As Currency
    BandValida As Boolean
    BandServicio As Boolean
    fechagrabado As Date
    bandVenta As Boolean         '***Agregado. Esteban. 26/12/2005
    CodUnidad As String 'AUC agregado
    CodUnidadConteo As String 'AUC agregado
    IdUnidad As Long    'AUC 08/12/05 'para unidades
    IdUnidadConteo As Long 'AUC 08/12/05 'para unidades de conteo
    CostoUltimoIngreso As Currency 'AUC 08/12/05 'para unidades de conteo
    CostoResidual As Currency 'jeaa 18/08/2008
    BandIVA As Boolean      'jeaa 08/03/2006 cambio para sri
    objEmpresa As Empresa            'Referencia al objeto padre
    CantRelUnidad As Long
    CantRelUnidadCont As Long
    PesoNeto As Currency      'jeaa 4/06/2008
    PesoBruto As Currency      'jeaa 4/06/2008
    IdUnidadPeso As Long      'jeaa 4/06/2008
    CodUnidadPeso As String      'jeaa 4/06/2008
    BandConversion As Boolean      'jeaa 4/06/2008
    BandRepGastos As Boolean      'jeaa 18/08/2008
    FechaIniDepre As Date
    NumeroDepre As Integer
    NumDepAnterior As Integer
    NumVidaUtil As Integer
    ValorDepreciar As Currency
    IdEmpleado As Long
    CodEmpleado As String
    Modelo As String
    RutaFoto As String
    ValorRevalor As Currency
    VidaUtilRevalor As Integer
    FechaRevalor As Date
    BandRevalor As Boolean
    FechaIniRevalor As Date
    NumRevalor As Byte
    NumRevalorNeg As Byte
    IdUbicacion As Long
    CodUbicacion As String
End Type


Private mProps As T_PROPS
Private mcolAFDetalles As Collection      'Coleccion de detalles de existencia
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean
Private mUIngreso As auxUltimo      'Ultimo ingreso
Private mUEgreso As auxUltimo       'Ultimo egreso
Private mcolDetalleSeguroFamilia As Collection      'Coleccion de detalles de Segurofamilia
Private mcolAFDetallesC As Collection      'Coleccion de detalles de existencia

Public Sub Grabar()
    Dim sql As String, i As Integer
    Dim rs As Recordset, rsDetalle As Recordset
    
    'Si no está modificado no hace nada
    
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
    sql = "SELECT CodInventario FROM AFInventario " & _
          "WHERE (CodInventario='" & mProps.CodInventario & "') AND " & _
                "(IdInventario<>" & mProps.IdInventario & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "AFInventario", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM AFInventario WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM AFInventario WHERE IdInventario=" & mProps.IdInventario
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
'        If Len(rs!FechaGrabado) Then
            If rs!fechagrabado <> mProps.fechagrabado Then
'        End If
            Err.Raise ERR_MODIFICADO, "AFInventario_Grabar", MSGERR_MODIFICADO
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
        For i = 1 To 5 'AUC cambiado antes 4
            .Fields("Precio" & i) = mProps.Precio(i)
        Next i
        !PorcentajeIVA = mProps.PorcentajeIVA
        !PorcentajeIVAAnt = mProps.PorcentajeIVAAnt
        !IdCuentaActivo = mProps.IdCuentaActivo
        !IdCuentaCosto = mProps.IdCuentaCosto
        !IdCuentaVenta = mProps.IdCuentaVenta
        !IdCuentaDepreGasto = mProps.IdCuentaDepreGasto
        !IdCuentaDepreAcumulada = mProps.IdCuentaDepreAcumulada
        !IdCuentaRevaloriza = mProps.IdCuentaRevaloriza
        !IdCuentaDepRevaloriza = mProps.IdCuentaDepRevaloriza
        
                
        For i = 1 To AFGRUPO_MAX
            .Fields("IdGrupo" & i) = mProps.IdGrupo(i)
        Next i
        !idProveedor = mProps.idProveedor
        !Marca = mProps.Marca
        !NumeroSerie = mProps.NumSerie
        !NumeroComprobanteCompra = mProps.NumFactura
        !VidaUtil = mProps.VidaUtil
        !DepAnterior = mProps.DepAnterior
        !Estado = mProps.Estado
        !tipodepre = mProps.tipodepre
        !FechaCompra = mProps.FechaCompra
        !Observacion = mProps.Observacion
        !ExistenciaMinima = mProps.ExistenciaMinima
        !ExistenciaMaxima = mProps.ExistenciaMaxima
        !UnidadMinimaCompra = mProps.UnidadMinimaCompra
        !UnidadMinimaVenta = mProps.UnidadMinimaCompra
        !BandValida = mProps.BandValida
        !BandServicio = mProps.BandServicio
        !IdUnidad = mProps.IdUnidad 'auc
        !IdUnidadConteo = mProps.IdUnidadConteo 'auc
        !BandIVA = mProps.BandIVA '***Agregado. Esteban 08/03/2006
        !CantRelUnidad = mProps.CantRelUnidad         'jeaa 27/09/2007
        !CantRelUnidadCont = mProps.CantRelUnidadCont         'jeaa 27/07/2007
        !Descripcion2 = mProps.Descripcion2             'jeaa 4/06/2008
        !PesoNeto = mProps.PesoNeto         'jeaa 04/06/2008
        !PesoBruto = mProps.PesoBruto         'jeaa 04/06/2008
        !IdUnidadPeso = mProps.IdUnidadPeso 'jeaa 04/06/2008
        !CostoUltimoIngreso = mProps.CostoUltimoIngreso 'jeaa 04/06/2008
        !CostoResidual = mProps.CostoResidual 'jeaa 04/06/2008

        !BandConversion = mProps.BandConversion 'jeaa 04/06/2008
        !BandRepGastos = mProps.BandRepGastos 'jeaa 18/08/2008
        
        !NumeroSerie = mProps.NumSerie
        !Estado = mProps.Estado
        !tipodepre = mProps.tipodepre
        !Formuladepre = mProps.tipodepre
        !FechaIniDepre = mProps.FechaIniDepre
        !NumeroDepre = mProps.NumeroDepre
        
        !NumDepAnterior = mProps.NumDepAnterior
        !NumVidaUtil = mProps.NumVidaUtil
        !ValorDepreciar = mProps.ValorDepreciar
        !IdEmpleado = mProps.IdEmpleado
        !Modelo = mProps.Modelo
        !RutaFoto = mProps.RutaFoto
        !ValorRevalor = mProps.ValorRevalor
        !VidaUtilRevalor = mProps.VidaUtilRevalor
        !FechaRevalor = mProps.FechaRevalor
        !BandRevalor = mProps.BandRevalor
        !FechaIniRevalor = mProps.FechaIniRevalor
        !NumRevalor = mProps.NumRevalor
        !NumRevalorNeg = mProps.NumRevalorNeg
        !IdUbicacion = mProps.IdUbicacion
        
        
        !fechagrabado = Now
        
            
            .Update
''''        GrabarProveedorDetalle (mProps.CodInventario) 'AUC 24/11/05
            GrabaSeguroFamiliaDetalle
            
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        rs.Move 0           'Para actualizar
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_AF", "Item AF: " & _
                                                         mProps.CodInventario, "DT"
        'jeaa 16-03-04 para grabar en ivexistencia la bodega y cantidad solo cuando es nuevo
        AseguraMovimientoExistencia
        AseguraMovimientoExistenciaCustodio
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_AF", "AF Modificado: " & _
                                                         mProps.CodInventario, "DT"
    End If
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    'Completa la transacción ---------------------------- jeaa 16-03-04
#If DAOLIB Then
    gobjWS.CommitTrans
#Else
    mProps.objEmpresa.Coneccion.CommitTrans
#End If
   mProps.objEmpresa.RaiseRefreshEvent REFRESH_AFINVENTARIO
    Exit Sub
ErrTrap:
    'Anula la transacción ---------------------------- jeaa 16-03-04
#If DAOLIB Then
    gobjWS.Rollback
#Else
    mProps.objEmpresa.Coneccion.RollbackTrans
#End If
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_AFINVENTARIO
End Sub

'AUC 25/11/05
Public Function RecuperaID(ByRef cod As String) As Long
    Dim sql As String, rs As Recordset

        sql = "SELECT idinventario FROM AFInventario WHERE Codinventario='" & cod & "'"
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
        sql = "SELECT * FROM vwAFInventarioRecuperar WHERE CodInventario='" & key & "'"
    Else
        sql = "SELECT * FROM vwAFInventarioRecuperar WHERE IdInventario=" & key
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
        
        For i = 1 To 5 'AUC cambiado antes 4
            If Not IsNull(.Fields("Precio" & i)) Then mProps.Precio(i) = .Fields("Precio" & i)
        Next i
        If Not IsNull(!PorcentajeIVA) Then mProps.PorcentajeIVA = !PorcentajeIVA
        If Not IsNull(!PorcentajeIVAAnt) Then mProps.PorcentajeIVAAnt = !PorcentajeIVAAnt
        If Not IsNull(!IdCuentaActivo) Then mProps.IdCuentaActivo = !IdCuentaActivo
        If Not IsNull(!IdCuentaCosto) Then mProps.IdCuentaCosto = !IdCuentaCosto
        If Not IsNull(!IdCuentaVenta) Then mProps.IdCuentaVenta = !IdCuentaVenta
        If Not IsNull(!IdCuentaDepreGasto) Then mProps.IdCuentaDepreGasto = !IdCuentaDepreGasto
        If Not IsNull(!IdCuentaDepreAcumulada) Then mProps.IdCuentaDepreAcumulada = !IdCuentaDepreAcumulada
        If Not IsNull(!IdCuentaRevaloriza) Then mProps.IdCuentaRevaloriza = !IdCuentaRevaloriza
        If Not IsNull(!IdCuentaDepRevaloriza) Then mProps.IdCuentaDepRevaloriza = !IdCuentaDepRevaloriza
        If Not IsNull(!CodCuentaActivo) Then mProps.CodCuentaActivo = !CodCuentaActivo
        If Not IsNull(!CodCuentaCosto) Then mProps.CodCuentaCosto = !CodCuentaCosto
        If Not IsNull(!CodCuentaVenta) Then mProps.CodCuentaVenta = !CodCuentaVenta
        If Not IsNull(!CodCuentaDepreGasto) Then mProps.CodCuentaDepreGasto = !CodCuentaDepreGasto
        If Not IsNull(!CodCuentaDepreAcumulada) Then mProps.CodCuentaDepreAcumulada = !CodCuentaDepreAcumulada
        If Not IsNull(!CodCuentaRevaloriza) Then mProps.CodCuentaRevaloriza = !CodCuentaRevaloriza
        If Not IsNull(!CodCuentaDepRevaloriza) Then mProps.CodCuentaDepRevaloriza = !CodCuentaDepRevaloriza
        
        For i = 1 To AFGRUPO_MAX
            If Not IsNull(.Fields("IdGrupo" & i)) Then mProps.IdGrupo(i) = .Fields("IdGrupo" & i)
            If Not IsNull(.Fields("CodGrupo" & i)) Then mProps.CodGrupo(i) = .Fields("CodGrupo" & i)
        Next i
        
        If Not IsNull(!idProveedor) Then mProps.idProveedor = !idProveedor
        If Not IsNull(!Estado) Then mProps.Estado = !Estado
        If Not IsNull(!tipodepre) Then mProps.tipodepre = !tipodepre
        If Not IsNull(!Marca) Then mProps.Marca = !Marca
        If Not IsNull(!NumeroSerie) Then mProps.NumSerie = !NumeroSerie
        If Not IsNull(!NumeroComprobanteCompra) Then mProps.NumFactura = !NumeroComprobanteCompra
        If Not IsNull(!VidaUtil) Then mProps.VidaUtil = !VidaUtil
        If Not IsNull(!DepAnterior) Then mProps.DepAnterior = !DepAnterior
        If Not IsNull(!FechaCompra) Then mProps.FechaCompra = !FechaCompra
        If Not IsNull(!CodProvCli) Then mProps.CodProveedor = !CodProvCli
        If Not IsNull(!Observacion) Then mProps.Observacion = !Observacion
        If Not IsNull(!ExistenciaMinima) Then mProps.ExistenciaMinima = !ExistenciaMinima
        If Not IsNull(!ExistenciaMaxima) Then mProps.ExistenciaMaxima = !ExistenciaMaxima
        If Not IsNull(!UnidadMinimaCompra) Then mProps.UnidadMinimaCompra = !UnidadMinimaCompra
        If Not IsNull(!UnidadMinimaCompra) Then mProps.UnidadMinimaVenta = !UnidadMinimaCompra
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!BandServicio) Then mProps.BandServicio = !BandServicio
        If Not IsNull(!bandVenta) Then mProps.bandVenta = !bandVenta
         If Not IsNull(!CodUnidad) Then mProps.CodUnidad = !CodUnidad
        If Not IsNull(!CodUniConteo) Then mProps.CodUnidadConteo = !CodUniConteo
        If Not IsNull(!IdUnidad) Then mProps.IdUnidad = !IdUnidad
        If Not IsNull(!IdUnidadConteo) Then mProps.IdUnidadConteo = !IdUnidadConteo
        If Not IsNull(!CostoUltimoIngreso) Then mProps.CostoUltimoIngreso = !CostoUltimoIngreso 'AUC 22/12/05
        If Not IsNull(!CostoResidual) Then mProps.CostoResidual = !CostoResidual
        If Not IsNull(!BandIVA) Then mProps.BandIVA = !BandIVA  'jeaa 08/03/2006
        If Not IsNull(!CantRelUnidad) Then mProps.CantRelUnidad = !CantRelUnidad 'jeaa 27/0/2007
        If Not IsNull(!CantRelUnidadCont) Then mProps.CantRelUnidadCont = !CantRelUnidadCont 'jeaa 27/0/2007
        If Not IsNull(!Descripcion2) Then mProps.Descripcion2 = !Descripcion2
        If Not IsNull(!PesoNeto) Then mProps.PesoNeto = !PesoNeto
        If Not IsNull(!PesoBruto) Then mProps.PesoBruto = !PesoBruto
        If Not IsNull(!IdUnidadPeso) Then mProps.IdUnidadPeso = !IdUnidadPeso
        If Not IsNull(!CodUniPeso) Then mProps.CodUnidadPeso = !CodUniPeso
        If Not IsNull(!BandConversion) Then mProps.BandConversion = !BandConversion
        If Not IsNull(!BandRepGastos) Then mProps.BandRepGastos = !BandRepGastos
        If Not IsNull(!FechaIniDepre) Then mProps.FechaIniDepre = !FechaIniDepre
        If Not IsNull(!NumeroDepre) Then mProps.NumeroDepre = !NumeroDepre
        
        If Not IsNull(!NumDepAnterior) Then mProps.NumDepAnterior = !NumDepAnterior
        If Not IsNull(!NumVidaUtil) Then mProps.NumVidaUtil = !NumVidaUtil
        If Not IsNull(!ValorDepreciar) Then mProps.ValorDepreciar = !ValorDepreciar
        
        If Not IsNull(!IdEmpleado) Then mProps.IdEmpleado = !IdEmpleado
        If Not IsNull(!CodEmpleado) Then mProps.CodEmpleado = !CodEmpleado
        If Not IsNull(!Modelo) Then mProps.Modelo = !Modelo
        If Not IsNull(!RutaFoto) Then mProps.RutaFoto = !RutaFoto
        
        If Not IsNull(!ValorRevalor) Then mProps.ValorRevalor = !ValorRevalor
        If Not IsNull(!VidaUtilRevalor) Then mProps.VidaUtilRevalor = !VidaUtilRevalor
        If Not IsNull(!FechaRevalor) Then mProps.FechaRevalor = !FechaRevalor
        If Not IsNull(!BandRevalor) Then mProps.BandRevalor = !BandRevalor
        If Not IsNull(!FechaIniRevalor) Then mProps.FechaIniRevalor = !FechaIniRevalor
        If Not IsNull(!NumRevalor) Then mProps.NumRevalor = !NumRevalor
        If Not IsNull(!NumRevalorNeg) Then mProps.NumRevalorNeg = !NumRevalorNeg
        
        If Not IsNull(!IdUbicacion) Then mProps.IdUbicacion = !IdUbicacion
        If Not IsNull(!CodUbicacion) Then mProps.CodUbicacion = !CodUbicacion
        
        
        mProps.fechagrabado = !fechagrabado
        
        CargaDetalleSeguroFamilia
        
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub


Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    
    'Verifica si ya existe un Kardex usando el mismo IdInventario
    sql = "SELECT Id FROM AFKardex WHERE IdInventario=" & mProps.IdInventario
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "AFInventario_Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close
    
 
    
    
    'Elimina el registro
    sql = "DELETE FROM AFInventario WHERE IdInventario=" & mProps.IdInventario
    mProps.objEmpresa.Execute sql, True

    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_ITEM", "Item Eliminado: " & _
                                                         mProps.CodInventario, "DT"
                                                         
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_AFINVENTARIO
End Sub

Public Function Clone() As AFInventario
    Dim nuevo As AFInventario
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New AFInventario
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

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

Friend Property Get IdInventario() As Long
    IdInventario = mProps.IdInventario
End Property


Public Property Let CodInventario(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "AFInventario_CodInventario", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodInventario = value
    mbooModificado = True
End Property

Public Property Get CodInventario() As String
    CodInventario = Trim$(mProps.CodInventario)
End Property

Public Property Let CodAlterno1(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "AFInventario_CodAlterno1", MSGERR_INVALIDO
    End If
    
    mProps.CodAlterno1 = value
    mbooModificado = True
End Property

Public Property Get CodAlterno1() As String
    CodAlterno1 = Trim$(mProps.CodAlterno1)
End Property

Public Property Let CodAlterno2(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "AFInventario_CodAlterno2", MSGERR_INVALIDO
    End If
    
    mProps.CodAlterno2 = value
    mbooModificado = True
End Property

Public Property Get CodAlterno2() As String
    CodAlterno2 = Trim$(mProps.CodAlterno2)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "AFInventario_Descripcion", MSGERR_INVALIDO
    End If
    
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, ";") > 0 Then
        Err.Raise ERR_INVALIDO, "AFInventario_Descripcion", _
            "El caracter ';' no es válido para la descripción."
    End If
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, "|") > 0 Then
        Err.Raise ERR_INVALIDO, "AFInventario_Descripcion", _
            "El caracter '|' no es válido para la descripción."
    End If
    
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
        Err.Raise ERR_INVALIDO, "AFInventario_CodMoneda", MSGERR_INVALIDO
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
    sql = "spIVCostoProm2 " & _
                mProps.IdInventario & ", " & _
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
    If Costo = 0 Then Costo = CostoUltimo(fecha, tid, Hora, False)
    
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
                            
                            
End Function

Public Property Let Precio(ByVal ix As Integer, ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "AFInventario.Precio" & ix, _
            "El precio no puede ser negativo. ('" & mProps.CodInventario & "')"
    End If
    
    '*** MAKOTO 30/nov/00 Agregado
    'Si no tiene permiso para modificar precios, genera error
    If gobjMain.GrupoActual.PermisoActual.CatInventarioPrecioMod = False Then
        Err.Raise ERR_NODERECHO, "AFInventario.Precio" & ix, MSGERR_NODERECHO
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



Public Property Let PorcentajeIVA(ByVal value As Currency)
    If value < 0 Or value > 1 Then
        Err.Raise ERR_INVALIDO, "AFInventario_PorcentajeIVA", MSGERR_INVALIDO
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
        Err.Raise ERR_INVALIDO, "AFInventario_PorcentajeIVAAnt", MSGERR_INVALIDO
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
                Err.Raise ERR_NOHAYCODIGO, "AFInventario_CodCuentaActivo", MSGERR_NOHAYCODIGO
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
                Err.Raise ERR_NOHAYCODIGO, "AFInventario_CodCuentaCosto", MSGERR_NOHAYCODIGO
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
                Err.Raise ERR_NOHAYCODIGO, "AFInventario_CodCuentaVenta", MSGERR_NOHAYCODIGO
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
            sql = "SELECT IdGrupo" & NumGrupo & " FROM AFGrupo" & NumGrupo & " WHERE CodGrupo" & NumGrupo & "='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "AFInventario_CodGrupo" & NumGrupo, MSGERR_NOHAYCODIGO
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
                Err.Raise ERR_NOHAYCODIGO, "AFInventario_CodProveedor", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                idProveedor = rs!IdProvCli
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
        Err.Raise ERR_INVALIDO, "AFInventario_Observacion", MSGERR_INVALIDO
    End If
    
    mProps.Observacion = value
    mbooModificado = True
End Property

Public Property Get Observacion() As String
    Observacion = Trim$(mProps.Observacion)
End Property

Public Property Let ExistenciaMinima(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "AFInventario_ExistenciaMinima", MSGERR_INVALIDO
    End If
    
    mProps.ExistenciaMinima = value
    mbooModificado = True
End Property

Public Property Get ExistenciaMinima() As Currency
    ExistenciaMinima = mProps.ExistenciaMinima
End Property

Public Property Let ExistenciaMaxima(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "AFInventario_ExistenciaMaxima", MSGERR_INVALIDO
    End If
    
    mProps.ExistenciaMaxima = value
    mbooModificado = True
End Property

Public Property Get ExistenciaMaxima() As Currency
    ExistenciaMaxima = mProps.ExistenciaMaxima
End Property

Public Property Let UnidadMinimaCompra(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "AFInventario_UnidadMinimaCompra", MSGERR_INVALIDO
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

Public Property Get FechaUltimoIngreso() As Date
    If ObtieneUltimo(True) Then FechaUltimoIngreso = mUIngreso.fecha
End Property

Public Property Get CantUltimoIngreso() As Currency
    If ObtieneUltimo(True) Then CantUltimoIngreso = mUIngreso.Cantidad
End Property

'AUC
Public Property Let CostoUltimoIngreso(ByVal value As Currency)
    mProps.CostoUltimoIngreso = value
    mbooModificado = True
End Property
'AUC
Public Property Get CostoUltimoIngreso() As Currency
    CostoUltimoIngreso = mProps.CostoUltimoIngreso
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
    If mcolDetalleSeguroFamilia Is Nothing Then
        NumFamiliaDetalle = 0
    Else
        NumFamiliaDetalle = mcolDetalleSeguroFamilia.Count
    End If
    'If ObtieneUltimo(False) Then PrecioUltimoEgreso = mUEgreso.Precio
End Property




Public Property Let UnidadMinimaVenta(ByVal value As Currency)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "AFInventario_UnidadMinimaVenta", MSGERR_INVALIDO
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
    Dim obj As AFInventarioDetalle
    
    If mcolAFDetalles Is Nothing Then RecuperaDetalle
    
    If Len(CodBodega) > 0 Then          '*** MAKOTO 15/dic/00
        For Each obj In mcolAFDetalles
            If obj.CodBodega = CodBodega Then
                Existencia = obj.Cantidad
                Exit For
            End If
        Next obj
    Else
        'Si no está especifada una bodega, devuelve la suma de todas las bodegas    '*** MAKOTO 15/dic/00
        For Each obj In mcolAFDetalles
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
        "|" & Format(Precio(5) * r, fmt)  'AUC 27/09/07

    ListaPrecioParaFlex = s
End Function


Public Sub GenerarCodigo()
    Dim sql As String, num As Double, cod As String, rs As Recordset
    Dim pre As String, formato As String

    pre = Me.Empresa.GNOpcion.PrefijoCodAFItem
    formato = Me.Empresa.GNOpcion.FormatoCodAFItem
    Select Case Me.Empresa.GNOpcion.GeneraCodAFItem
    Case 0  'Manual
        cod = pre
    Case 1  'Auto secuencial
        cod = Me.CodInventario
#If DAOLIB Then
        sql = "SELECT Max(VAL(CodInventario)) AS CodMax FROM AFInventario " & _
              "WHERE CodInventario LIKE '" & pre & "*'"
#Else
        sql = "SELECT Max(str(abs(CodInventario))) AS CodMax FROM AFInventario " & _
              "WHERE CodInventario LIKE '" & pre & "[0-9][0-9]%'"
#End If
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then
            If Not IsNull(rs!CodMax) Then num = Val(rs!CodMax)
        End If
        rs.Close
        num = num + 1
        cod = Format(num, formato)
    Case 2  'Auto con formato (g1+g2+"."+CodAlterno1)
        cod = mProps.CodGrupo(1) & mProps.CodGrupo(2) & "." & mProps.CodAlterno1
    Case 3 'Auto (Código Generado en Base a Grupos)
        cod = GeneraCodigoEncadenado
    Case 4 'Auto (Código Generado en Base a Grupos y numero continuo)
        cod = GeneraCodigoEncadenadoContinuo
    
    End Select
    
    'Verifica la longitud de código
    If Len(cod) > 20 Then
        Err.Raise ERR_INVALIDO, "AFInventario_CodInventario", _
            "No se puede generar debido a que el código se hace muy largo."
    End If
    
    'Otra Verificación se podría colocar aquí para otras reglas de codificación.
    '***
    
    'Si no hay error guarda en la propiedad
    Me.CodInventario = cod
    
    
    '------------------- CODIGO ALTERNO
    
    cod = ""
    pre = Me.Empresa.GNOpcion.PrefijoCodAFItemALT
    formato = Me.Empresa.GNOpcion.FormatoCodAFItemALT
    Select Case Me.Empresa.GNOpcion.GeneraCodAFItemALT
    Case 0  'Manual
        cod = pre
    Case 1  'Auto secuencial
        cod = Me.CodAlterno1
#If DAOLIB Then
        sql = "SELECT Max(VAL(CodAlterno1)) AS CodMax FROM AFInventario " & _
              "WHERE CodAlterno1 LIKE '" & pre & "*'"
#Else
        sql = "SELECT Max(str(abs(CodAlterno1))) AS CodMax FROM AFInventario " & _
              "WHERE CodAlterno1 LIKE '" & pre & "[0-9][0-9]%'"
#End If
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then
            If Not IsNull(rs!CodMax) Then num = Val(rs!CodMax)
        End If
        rs.Close
        num = num + 1
        cod = Format(num, formato)
    Case 2  'Auto con formato (g1+g2+"."+CodAlterno1)
        cod = mProps.CodGrupo(1) & mProps.CodGrupo(2) & "." & mProps.CodAlterno1
    Case 3 'Auto (Código Generado en Base a Grupos)
        cod = GeneraCodigoEncadenado
    Case 4 'Auto (Código Generado en Base a Grupos y numero continuo)
        cod = GeneraCodigoEncadenadoContinuo
    
    End Select
    Set rs = Nothing
    
    'Verifica la longitud de código
    If Len(cod) > 20 Then
        Err.Raise ERR_INVALIDO, "AFInventario_CodInventario", _
            "No se puede generar debido a que el código se hace muy largo."
    End If
    
    Me.CodAlterno1 = cod
    
    
    
End Sub



Private Sub Class_Initialize()
    'Valores predeterminados
    With mProps
        .CodMoneda = "USD"              '*** MAKOTO 02/oct/00 Modificado (PROVISIONALMENTE)
        .PorcentajeIVA = 0
        .ExistenciaMaxima = 0
        .ExistenciaMinima = 0
        .UnidadMinimaCompra = 1
        .UnidadMinimaVenta = 1
        .BandValida = True
        .BandServicio = False
        .bandVenta = True       'Agregado jeaa 26/12/2005
        .BandIVA = True       'Agregado jeaa 03/08/2006
        .CantRelUnidad = 1
        .CantRelUnidadCont = 1
        .BandRepGastos = False
        .VidaUtil = 1
        .DepAnterior = 0
        .Estado = 0
'        .PorcentajeIVA = Me.Empresa.GNOpcion.PorcentajeIVA
        .PorcentajeIVA = gobjMain.EmpresaActual.GNOpcion.PorcentajeIVA
        .PorcentajeIVAAnt = gobjMain.EmpresaActual.GNOpcion.PorcentajeIVAAnt
        
        .tipodepre = "2"
        .NumeroDepre = 0
        .NumDepAnterior = 0
        .NumVidaUtil = 0
        .ValorDepreciar = 0
        .BandRevalor = False
        .ValorRevalor = 0
    End With
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mcolAFDetalles = Nothing
    Set mProps.objEmpresa = Nothing
End Sub


'jeaa 16-03-04
'Para grabar el rgistro en AFexist y aparesca en lista de items para factura pese a que su cantidad es 0
Public Sub AseguraMovimientoExistencia()
    Dim CodBodega As String
    Dim sql As String, rs As Recordset, IdBodega As Integer, idItem As Long
    'recupera la bodega predeterminada
    CodBodega = mProps.objEmpresa.GNOpcion.ObtenerValor("BodegaAFPredeterminada")
    ' si no tiene bodega predeterminda sale sin grabar en ivexist
    If Len(CodBodega) = 0 Then Exit Sub
    sql = "SELECT IdBodega FROM AFBodega where codbodega='" & CodBodega & "'"
    'recupera idBodega
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        IdBodega = rs.Fields("IdBodega")
    End If
    'recupera iditem
    sql = "SELECT IdInventario FROM AFInventario where codinventario='" & mProps.CodInventario & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    idItem = rs.Fields("IdInventario")
    sql = "INSERT INTO AFExist (IdInventario,IdBodega,Exist) " & _
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
    Orden = Me.Empresa.GNOpcion.ObtenerValor("GenerarAFCodigoOrden")
    size = Me.Empresa.GNOpcion.ObtenerValor("GenerarAFCodigotamano")
    mascara = Me.Empresa.GNOpcion.ObtenerValor("GenerarAFCodigoMascara")
    
    codg = ""
    v1 = Split(Orden, ",")
    v2 = Split(size, ",")
        For i = LBound(v1, 1) To UBound(v1, 1)
        codg = codg & Trim$(Mid$(mProps.CodGrupo(v1(i)), 1, v2(i)))
    Next i
    
    pre = Me.Empresa.GNOpcion.PrefijoCodAFItem
    pre = Trim$(pre) & Trim$(codg)
    cod = Me.CodInventario
#If DAOLIB Then
    sql = "SELECT Max(VAL(CodInventario)) AS CodMax FROM AFInventario " & _
          "WHERE CodInventario LIKE '" & pre & "*'"
#Else
    sql = "SELECT (MAX(SUBSTRING(CodInventario," & Len(pre) + 1 & ",LEN(codinventario)))) As CodMax " & _
          "FROM AFInventario " & _
          "WHERE CodInventario LIKE '" & pre & "[0-9][0-9]%'"
#End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        If Not IsNull(rs!CodMax) Then num = Val(rs!CodMax)
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


Public Property Let CodUnidad(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodUnidad Then
        If Len(value) > 0 Then
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


'agregado esteban 15/09/2005
Public Property Let BandIVA(ByVal value As Boolean)
    mProps.BandIVA = value
    mbooModificado = True
End Property

Public Property Get BandIVA() As Boolean
    BandIVA = mProps.BandIVA
End Property

'AUC 08/12/05 'para unidades conteo
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

Public Function RecuperaAFExistenciaxBodega() As Recordset
    Dim sql As String, rs As Recordset

    'sql = "select * from AFInventario inner join ivexist on AFInventario.idinventario= AFInventario.idinventario where idinventario=' " & mobjItem.CodInventario
    sql = "select * from AFexist inner join afbodega on  AFexist.idbodega=AFbodega.idbodega where idinventario= " & Me.IdInventario
    Set rs = Me.Empresa.OpenRecordset(sql)
    With rs
     If rs.RecordCount <> 0 Then
        Set RecuperaAFExistenciaxBodega = rs
        Exit Function
        .Close
      End If
    End With
    Set rs = Nothing
End Function

Public Sub GrabaExistenciaxBodega(ByVal IdInventario As Integer, ByVal IdBodega As Integer, ByVal min As Integer, ByVal max As Integer)
End Sub


Public Sub GrabaAFExistenciaxBodegaS(ByVal IdInventario As Long, ByVal IdBodega As Integer, ByVal min As Integer, ByVal max As Integer)
    Dim sql As String, rs As Recordset

    sql = "update afexist "
    sql = sql & " set existmin=" & min & ", existmax=" & max
    sql = sql & " from afexist"
        sql = sql & " where idinventario= " & IdInventario
    sql = sql & " and idbodega=" & IdBodega
    Set rs = Me.Empresa.OpenRecordset(sql)
    Set rs = Nothing
End Sub

Public Property Let Descripcion2(ByVal value As String)
    If Len(value) > 120 Then
        Err.Raise ERR_INVALIDO, "AFInventario_Descripcion2", MSGERR_INVALIDO
    End If
    
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, ";") > 0 Then
        Err.Raise ERR_INVALIDO, "AFInventario_Descripcion2", _
            "El caracter ';' no es válido para la descripción."
    End If
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, "|") > 0 Then
        Err.Raise ERR_INVALIDO, "AFInventario_Descripcion2", _
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


Public Function RecuperaIdInventario(cod As String) As Long
    Dim sql As String, rs As Recordset

    sql = " select idinventario from AFInventario where codinventario='" & cod & "'"
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

Public Property Let NumSerie(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "AFActivoFijo_NumSerie", MSGERR_INVALIDO
    End If
    
    mProps.NumSerie = value
    mbooModificado = True
End Property

Public Property Get NumSerie() As String
    NumSerie = Trim$(mProps.NumSerie)
End Property

Public Property Let Marca(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "AFActivoFijo_Marca", MSGERR_INVALIDO
    End If
    
    mProps.Marca = value
    mbooModificado = True
End Property

Public Property Get Marca() As String
    Marca = Trim$(mProps.Marca)
End Property

Public Property Let NumFactura(ByVal value As String)
    If Len(value) > 13 Then
        Err.Raise ERR_INVALIDO, "AFActivoFijo_NumFactura", MSGERR_INVALIDO
    End If
    
    mProps.NumFactura = value
    mbooModificado = True
End Property

Public Property Get NumFactura() As String
    NumFactura = Trim$(mProps.NumFactura)
End Property

Public Property Let VidaUtil(ByVal value As String)
    If Len(value) < 0 Then
        Err.Raise ERR_INVALIDO, "AFActivoFijo_VidaUtil", MSGERR_INVALIDO
    End If
    
    mProps.VidaUtil = value
    mbooModificado = True
End Property

Public Property Get VidaUtil() As String
    VidaUtil = mProps.VidaUtil
End Property

Public Property Let FechaCompra(ByVal value As Date)
    mProps.FechaCompra = value
    mbooModificado = True
End Property

Public Property Get FechaCompra() As Date
    FechaCompra = mProps.FechaCompra
End Property

Public Property Let Estado(ByVal value As Byte)
    If value < 0 Or value > 3 Then
        Err.Raise ERR_INVALIDO, "AFActivoFijo_Estado", MSGERR_INVALIDO
    End If
    
    mProps.Estado = value
    mbooModificado = True
End Property

Public Property Get Estado() As Byte
    Estado = mProps.Estado
End Property

Public Property Let IdCuentaDepreGasto(ByVal value As Long)
    mProps.IdCuentaDepreGasto = value
    mbooModificado = True
End Property

Public Property Get IdCuentaDepreGasto() As Long
    IdCuentaDepreGasto = mProps.IdCuentaDepreGasto
End Property

            
Public Property Let CodCuentaDepreGasto(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodCuentaDepreGasto Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "AFInventario_CodCuentaDepreGasto", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaDepreGasto = rs!IdCuenta
            End If
            rs.Close
        Else
            IdCuentaDepreGasto = 0
        End If
        
        mProps.CodCuentaDepreGasto = value
        mbooModificado = True
    End If

    Set rs = Nothing
End Property

Public Property Get CodCuentaDepreGasto() As String
    CodCuentaDepreGasto = Trim$(mProps.CodCuentaDepreGasto)
End Property
            
            
            
Public Property Let IdCuentaDepreAcumulada(ByVal value As Long)
    mProps.IdCuentaDepreAcumulada = value
    mbooModificado = True
End Property

Public Property Get IdCuentaDepreAcumulada() As Long
    IdCuentaDepreAcumulada = mProps.IdCuentaDepreAcumulada
End Property


Public Property Let CodCuentaDepreAcumulada(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodCuentaDepreAcumulada Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "AFInventario_CodCuentaDepreAcumulada", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaDepreAcumulada = rs!IdCuenta
            End If
            rs.Close
        Else
            IdCuentaDepreAcumulada = 0
        End If
        
        mProps.CodCuentaDepreAcumulada = value
        mbooModificado = True
    End If

    Set rs = Nothing
End Property

Public Property Get CodCuentaDepreAcumulada() As String
    CodCuentaDepreAcumulada = Trim$(mProps.CodCuentaDepreAcumulada)
End Property


Public Property Let IdCuentaRevaloriza(ByVal value As Long)
    mProps.IdCuentaRevaloriza = value
    mbooModificado = True
End Property

Public Property Get IdCuentaRevaloriza() As Long
    IdCuentaRevaloriza = mProps.IdCuentaRevaloriza
End Property

Public Property Let CodCuentaRevaloriza(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodCuentaRevaloriza Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "AFInventario_CodCuentaRevaloriza", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaRevaloriza = rs!IdCuenta
            End If
            rs.Close
        Else
            IdCuentaRevaloriza = 0
        End If
        
        mProps.CodCuentaRevaloriza = value
        mbooModificado = True
    End If

    Set rs = Nothing
End Property

Public Property Get CodCuentaRevaloriza() As String
    CodCuentaRevaloriza = Trim$(mProps.CodCuentaRevaloriza)
End Property


Public Property Let IdCuentaDepRevaloriza(ByVal value As Long)
    mProps.IdCuentaDepRevaloriza = value
    mbooModificado = True
End Property

Public Property Get IdCuentaDepRevaloriza() As Long
    IdCuentaDepRevaloriza = mProps.IdCuentaDepRevaloriza
End Property

Public Property Let CodCuentaDepRevaloriza(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodCuentaDepRevaloriza Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "AFInventario_CodCuentaDepRevaloriza ", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaDepRevaloriza = rs!IdCuenta
            End If
            rs.Close
        Else
            IdCuentaDepRevaloriza = 0
        End If
        
        mProps.CodCuentaDepRevaloriza = value
        mbooModificado = True
    End If

    Set rs = Nothing
End Property

Public Property Get CodCuentaDepRevaloriza() As String
    CodCuentaDepRevaloriza = Trim$(mProps.CodCuentaDepRevaloriza)
End Property


'Diego Prod
'--------------- Manejo de coleccion de DetalleSeguroFamilia
Public Function AddDetalleSeguroFamilia() As Long
    Dim nuevo As AFSeguroFamilia
    
    If mcolDetalleSeguroFamilia Is Nothing Then CargaDetalleSeguroFamilia
    Set nuevo = New AFSeguroFamilia
    Set nuevo.AFInventario = Me  'Me  olvide  del padre
    'Proporciona valores predeterminados
    nuevo.Cantidad = 1
    nuevo.FechaInicio = Date
    nuevo.FechaFin = Date
    mcolDetalleSeguroFamilia.Add nuevo
    AddDetalleSeguroFamilia = mcolDetalleSeguroFamilia.Count    'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    
    mbooModificado = True
    
End Function

Public Sub RemoveDetalleSeguroFamilia(ByVal ix As Long, Optional obj As AFSeguroFamilia)
    If mcolDetalleSeguroFamilia Is Nothing Then CargaDetalleSeguroFamilia
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolDetalleSeguroFamilia.Count
            If mcolDetalleSeguroFamilia.item(ix) Is obj Then
                mcolDetalleSeguroFamilia.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolDetalleSeguroFamilia.Remove ix
    End If
    
    'mcolDetalleFamilia.Remove ix
    mbooModificado = True
End Sub

Private Sub CargaDetalleSeguroFamilia()
    Dim rs As Recordset, sql As String, obj As AFSeguroFamilia
    Dim i As Long
    Set mcolDetalleSeguroFamilia = New Collection


    sql = " select aft.IdTipoSeguro, aft.CodTipoSeguro, aft.Descripcion, "
    sql = sql & " pc.IdProvCli, pc.CodProvCli, pc.Nombre, "
    sql = sql & " Referencia , Cantidad, ValorAsegurado, "
    sql = sql & " FechaInicio, FechaFin "
    sql = sql & " from AFSeguroDetalle afs"
    sql = sql & " inner join afTipoSeguro aft on afs.idtiposeguro=aft.idtiposeguro"
    sql = sql & " inner join PcProvCli pc on afs.idproveedor=pc.idProvcli"
    sql = sql & " Where afs.IdInventario = " & mProps.IdInventario
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddDetalleSeguroFamilia
            Set obj = mcolDetalleSeguroFamilia.item(i)
            obj.IdTipoSeguro = !IdTipoSeguro
            obj.CodTipoSeguro = !CodTipoSeguro
            obj.DescTipoSeguro = !Descripcion
            obj.idProveedor = !IdProvCli
            obj.CodProveedor = !CodProvCli
            obj.NombreProveedor = !Nombre
            obj.Referencia = !Referencia
            obj.Cantidad = !Cantidad
            obj.ValorAsegurado = !ValorAsegurado
            obj.FechaInicio = !FechaInicio
            obj.FechaFin = !FechaFin
           .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Function RecuperaDetalleSeguroFamilia(ix As Long) As AFSeguroFamilia
    Set RecuperaDetalleSeguroFamilia = mcolDetalleSeguroFamilia.item(ix)
End Function


Public Sub GrabaSeguroFamiliaDetalle()
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As AFSeguroFamilia, sql As String, band As Boolean
    'Elimina  todos  los registros
    sql = "Delete AfS "
    sql = sql & " From AFinventario AF "
    sql = sql & "INNER JOIN AFSeguroDetalle AfS ON AF.IdInventario = AFS.IdInventario "
    sql = sql & " Where afs.IdInventario = " & mProps.IdInventario

    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    band = True
    If Not mcolDetalleSeguroFamilia Is Nothing Then
        sql = "SELECT * FROM AFSeguroDetalle WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        For i = 1 To mcolDetalleSeguroFamilia.Count
            Set obj = mcolDetalleSeguroFamilia.item(i)
            If band Then
                rs1.AddNew
                    rs1!IdInventario = Me.IdInventario
                    rs1!IdTipoSeguro = obj.IdTipoSeguro
                    rs1!idProveedor = obj.idProveedor
                    rs1!Referencia = obj.Referencia
                    rs1!Cantidad = obj.Cantidad
                    rs1!ValorAsegurado = obj.ValorAsegurado
                    rs1!FechaInicio = obj.FechaInicio
                    rs1!FechaFin = obj.FechaFin
                    
                rs1.Update
            End If
            band = True
        Next i
    End If
    Set rs1 = Nothing
    Set obj = Nothing
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

Private Sub RecuperaDetalle()
    Dim sql As String, i As Long, obj As AFInventarioDetalle
    Dim rs As Recordset

    Set mcolAFDetalles = New Collection
    
    'Recupera los detalles de existencia
    sql = "SELECT AFExist.IdInventario, AFInventario.CodInventario, " & _
          "AFExist.IdBodega, AFBodega.CodBodega, AFExist.Exist " & _
          "FROM (AFExist INNER JOIN AFInventario ON AFExist.IdInventario=" & _
          "AFInventario.IdInventario) INNER JOIN AFBodega ON AFExist.IdBodega " & _
          "= AFBodega.IdBodega " & _
          "WHERE AFExist.IdInventario=" & mProps.IdInventario & " ORDER BY CodBodega"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddAFDetalle
            Set obj = mcolAFDetalles.item(i)
            obj.IdBodega = !IdBodega
            obj.CodBodega = !CodBodega
            obj.Cantidad = !exist
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
End Sub


Public Function AddAFDetalle() As Long
    Dim nuevo As AFInventarioDetalle
    
    Set nuevo = New AFInventarioDetalle
    Set nuevo.AFInventario = Me
    mcolAFDetalles.Add nuevo
    AddAFDetalle = mcolAFDetalles.Count    'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
End Function

Private Function GeneraCodigoEncadenadoContinuo() As String
    Dim sql As String, num As Long, cod As String, rs As Recordset
    Dim pre As String, codg As String, v1 As Variant, v2 As Variant
    Dim Orden As String, size As String, mascara As String, i As Integer
    
    
    '***Aqui hay formar la parte de los grupos
    'unir con esteban
    Orden = Me.Empresa.GNOpcion.ObtenerValor("GenerarAFCodigoOrden")
    size = Me.Empresa.GNOpcion.ObtenerValor("GenerarAFCodigotamano")
    mascara = Me.Empresa.GNOpcion.ObtenerValor("GenerarAFCodigoMascara")
    
    codg = ""
    v1 = Split(Orden, ",")
    v2 = Split(size, ",")
        For i = LBound(v1, 1) To UBound(v1, 1)
        codg = codg & Trim$(Mid$(mProps.CodGrupo(v1(i)), 1, v2(i)))
    Next i
    
    pre = Me.Empresa.GNOpcion.PrefijoCodAFItem
    pre = Trim$(pre) & Trim$(codg)
    cod = Me.CodInventario
#If DAOLIB Then
    sql = "SELECT Max(VAL(CodInventario)) AS CodMax FROM AFInventario " & _
          "WHERE CodInventario LIKE '" & pre & "*'"
#Else
    sql = "SELECT (MAX(SUBSTRING(CodInventario," & Len(pre) + 1 & ",LEN(codinventario)))) As CodMax " & _
          "FROM AFInventario " & _
          "WHERE CodInventario LIKE '" & "[0-9][0-9]%'"
#End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        If Not IsNull(rs!CodMax) Then num = Val(rs!CodMax)
    End If
    rs.Close
    num = num + 1
    cod = pre & Format(num, mascara) 'cambiar por plantilla
    
    GeneraCodigoEncadenadoContinuo = cod
    Set rs = Nothing
End Function


Public Property Let DepAnterior(ByVal value As String)
    If Len(value) < 0 Then
        Err.Raise ERR_INVALIDO, "AFActivoFijo_DepAnterior", MSGERR_INVALIDO
    End If
    
    mProps.DepAnterior = value
    mbooModificado = True
End Property

Public Property Get DepAnterior() As String
    DepAnterior = mProps.DepAnterior
End Property

Public Property Let tipodepre(ByVal value As Byte)
    If value < 0 Or value > 2 Then
        Err.Raise ERR_INVALIDO, "AFActivoFijo_TipoDepre", MSGERR_INVALIDO
    End If
    
    mProps.tipodepre = value
    mbooModificado = True
End Property

Public Property Get tipodepre() As Byte
    tipodepre = mProps.tipodepre
End Property

Public Property Let CostoResidual(ByVal value As Currency)
    mProps.CostoResidual = value
    mbooModificado = True
End Property
'AUC
Public Property Get CostoResidual() As Currency
    CostoResidual = mProps.CostoResidual
End Property

Public Property Let FechaIniDepre(ByVal value As Date)
    mProps.FechaIniDepre = value
    mbooModificado = True
End Property

Public Property Get FechaIniDepre() As Date
    FechaIniDepre = mProps.FechaIniDepre
End Property

Public Property Let NumeroDepre(ByVal value As Integer)
    mProps.NumeroDepre = value
    mbooModificado = True
End Property

Public Property Get NumeroDepre() As Integer
    NumeroDepre = mProps.NumeroDepre
End Property

Public Property Let NumDepAnterior(ByVal value As Integer)
    mProps.NumDepAnterior = value
    mbooModificado = True
End Property

Public Property Get NumDepAnterior() As Integer
    NumDepAnterior = mProps.NumDepAnterior
End Property

Public Property Let NumVidaUtil(ByVal value As Integer)
    mProps.NumVidaUtil = value
    mbooModificado = True
End Property

Public Property Get NumVidaUtil() As Integer
    NumVidaUtil = mProps.NumVidaUtil
End Property

Public Property Let ValorDepreciar(ByVal value As Currency)
    mProps.ValorDepreciar = value
    mbooModificado = True
End Property

Public Property Get ValorDepreciar() As Currency
    ValorDepreciar = mProps.ValorDepreciar
End Property



Public Property Let IdEmpleado(ByVal value As Long)
    mProps.IdEmpleado = value
    mbooModificado = True
End Property

Public Property Get IdEmpleado() As Long
    IdEmpleado = mProps.IdEmpleado
End Property


Public Property Let CodEmpleado(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodEmpleado Then
        If Len(value) > 0 Then
            'Actualiza IdEmpleado también
            sql = "SELECT IdProvCli FROM Empleado WHERE CodProvCli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "AFInventario_CodEmpleado", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdEmpleado = rs!IdProvCli
            End If
            rs.Close
        Else
            IdEmpleado = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodEmpleado = value
    mbooModificado = True
End Property

Public Property Get CodEmpleado() As String
    CodEmpleado = Trim$(mProps.CodEmpleado)
End Property

Public Property Let Modelo(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "AFActivoFijo_Modelo", MSGERR_INVALIDO
    End If
    
    mProps.Modelo = value
    mbooModificado = True
End Property

Public Property Get Modelo() As String
    Modelo = Trim$(mProps.Modelo)
End Property

Public Property Let RutaFoto(ByVal value As String)
    If Len(value) > 200 Then
        Err.Raise ERR_INVALIDO, "AFActivoFijo_RutaFoto", MSGERR_INVALIDO
    End If
    
    mProps.RutaFoto = value
    mbooModificado = True
End Property

Public Property Get RutaFoto() As String
    RutaFoto = Trim$(mProps.RutaFoto)
End Property


Private Sub RecuperaDetalleCustodio()
    Dim sql As String, i As Long, obj As AFInventarioDetalleC
    Dim rs As Recordset

    Set mcolAFDetallesC = New Collection
    
    'Recupera los detalles de existencia
    sql = "SELECT AFExistCustodio.IdInventario, AFInventario.CodInventario, " & _
          "AFExistCustodio.IdProvCli, Pcprovcli.Codprovcli as CodEmpleado, AFExistCustodio.Exist " & _
          "FROM (AFExistCustodio INNER JOIN AFInventario ON AFExistCustodio.IdInventario=" & _
          "AFInventario.IdInventario) INNER JOIN empleado PCProvCli ON AFExistCustodio.IdProvcli " & _
          "= PCProvcli.IdProvCli " & _
          "WHERE AFExistCustodio.IdInventario=" & mProps.IdInventario & " ORDER BY CodEmpleado"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddAFDetalleC
            Set obj = mcolAFDetallesC.item(i)
            obj.IdProvCli = !IdProvCli
            obj.CodEmpleado = !CodEmpleado
            obj.Cantidad = !exist
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
End Sub


Public Function AddAFDetalleC() As Long
    Dim nuevo As AFInventarioDetalleC
    
    Set nuevo = New AFInventarioDetalleC
    Set nuevo.AFInventario = Me
    mcolAFDetallesC.Add nuevo
    AddAFDetalleC = mcolAFDetallesC.Count    'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
End Function



Public Function ExistenciaCustodio(CodEmpleado As String) As Currency
    Dim obj As AFInventarioDetalleC
    
    If mcolAFDetallesC Is Nothing Then RecuperaDetalleCustodio
    
    If Len(CodEmpleado) > 0 Then          '*** MAKOTO 15/dic/00
        For Each obj In mcolAFDetallesC
            If obj.CodEmpleado = CodEmpleado Then
                ExistenciaCustodio = obj.Cantidad
                Exit For
            End If
        Next obj
    Else
        'Si no está especifada una bodega, devuelve la suma de todas las bodegas    '*** MAKOTO 15/dic/00
        For Each obj In mcolAFDetallesC
            ExistenciaCustodio = ExistenciaCustodio + obj.Cantidad
        Next obj
    End If
    Set obj = Nothing
End Function

Public Property Let ValorRevalor(ByVal value As Currency)
    mProps.ValorRevalor = value
    mbooModificado = True
End Property

Public Property Get ValorRevalor() As Currency
    ValorRevalor = mProps.ValorRevalor
End Property

Public Property Let VidaUtilRevalor(ByVal value As Integer)
    mProps.VidaUtilRevalor = value
    mbooModificado = True
End Property

Public Property Get VidaUtilRevalor() As Integer
    VidaUtilRevalor = mProps.VidaUtilRevalor
End Property

Public Property Let FechaRevalor(ByVal value As Date)
    mProps.FechaRevalor = value
    mbooModificado = True
End Property

Public Property Get FechaRevalor() As Date
    FechaRevalor = mProps.FechaRevalor
End Property

Public Property Let BandRevalor(ByVal value As Boolean)
    mProps.BandRevalor = value
    mbooModificado = True
End Property

Public Property Get BandRevalor() As Boolean
    BandRevalor = mProps.BandRevalor
End Property

Public Sub ActualizaDatosRevalorizacion(ByVal CodInv As String, Valor As Double, _
    cant As Integer, fecha As Date, fechaIni As Date, transid As Long, CodTrans As String, _
    neg As Boolean)
    Dim sql As String, rs As Recordset
    Dim numero As Integer
    
    numero = ObtieneNumeroRevalorizacion(transid, CodTrans, neg)

    sql = "Update AFInventario set"
    sql = sql & " ValorRevalor=" & Valor & ","
    sql = sql & " vidaUtilRevalor=" & cant & ","
    sql = sql & " fechaRevalor='" & fecha & "',"
    sql = sql & " fechaIniRevalor='" & fechaIni & "',"
    sql = sql & " BandRevalor=1,"
    If neg Then
        sql = sql & " NumRevalor=" & numero
    Else
        sql = sql & " NumRevalorNeg=" & numero
    End If
    sql = sql & " where codInventario='" & CodInventario & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    
End Sub

Public Property Let FechaIniRevalor(ByVal value As Date)
    mProps.FechaIniRevalor = value
    mbooModificado = True
End Property

Public Property Get FechaIniRevalor() As Date
    FechaIniRevalor = mProps.FechaIniRevalor
End Property

Public Property Let NumRevalor(ByVal value As Date)
    mProps.NumRevalor = value
    mbooModificado = True
End Property

Public Property Get NumRevalor() As Date
    NumRevalor = mProps.NumRevalor
End Property

Public Property Let NumRevalorNeg(ByVal value As Date)
    mProps.NumRevalorNeg = value
    mbooModificado = True
End Property

Public Property Get NumRevalorNeg() As Date
    NumRevalorNeg = mProps.NumRevalorNeg
End Property


Public Function ObtieneNumeroRevalorizacion(transid As Long, CodTrans As String, _
    neg As Boolean) As Integer
    Dim sql As String, rs As Recordset
    Dim numero As Integer

    sql = "select count(codtrans) as cant from gncomprobante g"
    sql = sql & " inner join afkardex a "
    sql = sql & " inner join afinventario af "
    sql = sql & " on af.idinventario=a.idinventario"
    sql = sql & " on g.transid=a.transid "
    sql = sql & " where g.estado<>3 and codtrans in ('" & CodTrans & "','" & CodTrans & "M')"
    sql = sql & " and g.transid<>" & transid
    sql = sql & " and codinventario='" & Me.CodInventario & "'"
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    ObtieneNumeroRevalorizacion = rs.Fields("cant") + 1

End Function

Public Sub AseguraMovimientoExistenciaCustodio()
    Dim CodBodega As String
    Dim sql As String, rs As Recordset, IdBodega As Long, idItem As Long
    'recupera la bodega predeterminada
    CodBodega = mProps.objEmpresa.GNOpcion.ObtenerValor("CustodioAFPredeterminada")
    ' si no tiene bodega predeterminda sale sin grabar en ivexist
    If Len(CodBodega) = 0 Then Exit Sub
    sql = "SELECT IdProvCli FROM empleado where codProvcli='" & CodBodega & "'"
    'recupera idBodega
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    IdBodega = rs.Fields("IdProvcli")
    'recupera iditem
    sql = "SELECT IdInventario FROM AFinventario where codinventario='" & mProps.CodInventario & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    idItem = rs.Fields("IdInventario")
    sql = "INSERT INTO AFExistCustodio (IdInventario,IdProvCli,Exist) " & _
              "VALUES(" & idItem & "," & IdBodega & "," & 0 & ")"
    mProps.objEmpresa.OpenRecordset sql
    
    rs.Close
    Set rs = Nothing
End Sub


Friend Sub RecuperarQuick(key As Variant)
    Dim sql As String, i As Integer
    Dim rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM vwAFInventarioRecuperar WHERE CodInventario='" & key & "'"
    Else
        sql = "SELECT * FROM vwAFInventarioRecuperar WHERE IdInventario=" & key
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
'        If Not IsNull(!DescripcionDetalle) Then mProps.DescripcionDetalle = !DescripcionDetalle
#Else
'        If Not IsNull(!DescripcionDetalle) Then
'            If .Fields("DescripcionDetalle").ActualSize > 0 Then
'                mProps.DescripcionDetalle = .Fields("DescripcionDetalle").GetChunk(.Fields("DescripcionDetalle").ActualSize)
'            End If
'        End If
#End If
        If Not IsNull(!CodMoneda) Then mProps.CodMoneda = !CodMoneda
        
'        If Not IsNull(!PorcentajeIVA) Then mProps.PorcentajeIVA = !PorcentajeIVA
        If Not IsNull(!IdCuentaActivo) Then mProps.IdCuentaActivo = !IdCuentaActivo
        If Not IsNull(!IdCuentaCosto) Then mProps.IdCuentaCosto = !IdCuentaCosto
        If Not IsNull(!IdCuentaVenta) Then mProps.IdCuentaVenta = !IdCuentaVenta
        If Not IsNull(!IdCuentaDepreGasto) Then mProps.IdCuentaDepreGasto = !IdCuentaDepreGasto
        If Not IsNull(!IdCuentaDepreAcumulada) Then mProps.IdCuentaDepreAcumulada = !IdCuentaDepreAcumulada
        If Not IsNull(!IdCuentaRevaloriza) Then mProps.IdCuentaRevaloriza = !IdCuentaRevaloriza
        If Not IsNull(!IdCuentaDepRevaloriza) Then mProps.IdCuentaDepRevaloriza = !IdCuentaDepRevaloriza
        If Not IsNull(!CodCuentaActivo) Then mProps.CodCuentaActivo = !CodCuentaActivo
        If Not IsNull(!CodCuentaCosto) Then mProps.CodCuentaCosto = !CodCuentaCosto
        If Not IsNull(!CodCuentaVenta) Then mProps.CodCuentaVenta = !CodCuentaVenta
        If Not IsNull(!CodCuentaDepreGasto) Then mProps.CodCuentaDepreGasto = !CodCuentaDepreGasto
        If Not IsNull(!CodCuentaDepreAcumulada) Then mProps.CodCuentaDepreAcumulada = !CodCuentaDepreAcumulada
        If Not IsNull(!CodCuentaRevaloriza) Then mProps.CodCuentaRevaloriza = !CodCuentaRevaloriza
        If Not IsNull(!CodCuentaDepRevaloriza) Then mProps.CodCuentaDepRevaloriza = !CodCuentaDepRevaloriza
        
        For i = 1 To AFGRUPO_MAX
            If Not IsNull(.Fields("IdGrupo" & i)) Then mProps.IdGrupo(i) = .Fields("IdGrupo" & i)
            If Not IsNull(.Fields("CodGrupo" & i)) Then mProps.CodGrupo(i) = .Fields("CodGrupo" & i)
        Next i
        
        
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Property Let IdUbicacion(ByVal value As Long)
    mProps.IdUbicacion = value
    mbooModificado = True
End Property

Public Property Get IdUbicacion() As Long
    IdUbicacion = mProps.IdUbicacion
End Property


Public Property Let CodUbicacion(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodUbicacion Then
        If Len(value) > 0 Then
            'Actualiza IdUbicacion también
            sql = "SELECT IdUbicacion FROM AFUbicacion WHERE CodUbicacion='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "AFInventario_CodUbicacion", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdUbicacion = rs!IdUbicacion
            End If
            rs.Close
        Else
            IdUbicacion = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodUbicacion = value
    mbooModificado = True
End Property

Public Property Get CodUbicacion() As String
    CodUbicacion = Trim$(mProps.CodUbicacion)
End Property

