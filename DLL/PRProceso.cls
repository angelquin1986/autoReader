VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PRProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Auxiliar para guardar datos de ultimo ingreso/egreso

Private Type auxUltimo
    Fecha As Date
    Cantidad As Currency
    Costo As Currency
    Precio As Currency
    Recuperado As Boolean
End Type

Private Type T_PROPS
    IdProceso As Long
    CodProceso As String
    Descripcion As String
    Observacion As String
    BandValida As Boolean
    FechaGrabado As Date
    IdUnidad As Long
    CodUnidad As String
    objEmpresa As Empresa            'Referencia al objeto padre
End Type


Private mProps As T_PROPS
Private mcolDetalles As Collection      'Coleccion de detalles de existencia
Private mcolDetalleFamilia As Collection      'Coleccion de detalles de familia
Private mcolDetalleMaterial As Collection      'Coleccion de detalles de Material
Private mcolDetalleMaquinaria As Collection      'Coleccion de detalles de Maquinaria
Private mcolDetalleManoObra As Collection      'Coleccion de detalles de ManoObra

Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean
Private mUIngreso As auxUltimo      'Ultimo ingreso
Private mUEgreso As auxUltimo       'Ultimo egreso

Public Sub Grabar()
    Dim sql As String, i As Integer
    Dim rs As Recordset, rsDetalle As Recordset
    
    'Si no está modificado no hace nada
    
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
    sql = "SELECT CodProceso FROM PRProceso " & _
          "WHERE (CodProceso='" & mProps.CodProceso & "') AND " & _
                "(IdProceso<>" & mProps.IdProceso & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "PRProceso", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM PRProceso WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM PRProceso WHERE IdProceso=" & mProps.IdProceso
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
'        If Len(rs!FechaGrabado) Then
            If rs!FechaGrabado <> mProps.FechaGrabado Then
'        End If
            Err.Raise ERR_MODIFICADO, "PRProceso_Grabar", MSGERR_MODIFICADO
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
        !CodProceso = mProps.CodProceso
        !Descripcion = mProps.Descripcion
        !Observacion = mProps.Observacion
'        GrabaFamiliaDetalle
        !BandValida = mProps.BandValida
        !IdUnidad = mProps.IdUnidad 'auc
        !FechaGrabado = Now
        .Update
        
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        rs.Move 0           'Para actualizar
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_ITEM", "Item Nuevo: " & _
                                                         mProps.CodProceso, "DT"
        'jeaa 16-03-04 para grabar en ivexistencia la bodega y cantidad solo cuando es nuevo
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_ITEM", "Item Modificado: " & _
                                                         mProps.CodProceso, "DT"
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
   mProps.objEmpresa.RaiseRefreshEvent REFRESH_PRPROCESO
    Exit Sub
ErrTrap:
    'Anula la transacción ---------------------------- jeaa 16-03-04
#If DAOLIB Then
    gobjWS.Rollback
#Else
    mProps.objEmpresa.Coneccion.RollbackTrans
#End If
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PRPROCESO
End Sub

'Private Sub GrabaFamiliaDetalle()
'    Dim rs1 As Recordset, numReg As Long, i As Long
'    Dim obj As IVFamiliaDetalle, sql As String, band As Boolean
'    'Elimina  todos  los registros
'    sql = "Delete Ivm " & _
'          "From PRProceso Iv " & _
'          "INNER JOIN IVMateria Ivm ON IV.IdProceso = Ivm.IdMateria "
'    'jeaa 28/10/2005
'   'If mProps.tipo = CambioPresentacion Or mProps.tipo = Preparacion Or tipo = Promocion Then
'          sql = sql & " Where Ivm.IdProceso = " & mProps.IdProceso
'   'Else
'    '      sql = sql & " Where Ivm.Idmateria = " & mProps.IdProceso
'   'End If
'
'    mProps.objEmpresa.EjecutarSQL sql, numReg
'    'guarda  todos  los registros
'    band = True
'    If Not mcolDetalleFamilia Is Nothing Then
'        sql = "SELECT * FROM IVMateria WHERE 1=0"
'        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
'        For i = 1 To mcolDetalleFamilia.Count
''            rs1.AddNew /jeaa 28-01-04
'            Set obj = mcolDetalleFamilia.item(i)
''            If obj.IdProceso = 0 Then
'                        obj.Recuperar obj.CodProceso
'                        If verificacionDuiplicado(obj.IdProceso) Then
'                            RemoveDetalleFamilia i
'                            band = False
'                        End If
' '           End If
'            If band Then
'                rs1.AddNew
'                If Me.tipo <> Familia Then
'                    rs1!IdProceso = Me.IdProceso
'                    rs1!IdMateria = obj.IdProceso
'                Else
'                    rs1!IdProceso = obj.IdProceso
'                    rs1!IdMateria = Me.IdProceso
'                End If
'                rs1!Cantidad = obj.Cantidad
'                If Me.tipo = Promocion Then
'                    rs1!xCuanto = obj.xCuanto 'AUC 26/09/007
'                    rs1!tarifajornal = obj.Precio  'AUC 26/09/007
'                End If
'                rs1.Update
'            End If
'            band = True
'        Next i
'    End If
'    Set rs1 = Nothing
'    Set obj = Nothing
'End Sub
'AUC 25/11/05
Public Function RecuperaID(ByRef cod As String) As Long
    Dim sql As String, rs As Recordset

        sql = "SELECT idProceso FROM PRProceso WHERE CodProceso='" & cod & "'"
    Set rs = Me.Empresa.OpenRecordset(sql)
    With rs
     If rs.RecordCount = 0 Then
        RecuperaID = 0
      Else
        RecuperaID = !IdProceso
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
        sql = "SELECT * FROM vwIVProcesoRecuperar WHERE CodProceso='" & key & "'"
    Else
        sql = "SELECT * FROM vwIVProcesoRecuperar WHERE IdProceso=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IdProceso = !IdProceso
        mProps.CodProceso = !CodProceso
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!Observacion) Then mProps.Observacion = !Observacion
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        CargaDetalleFamilia 'Pensandolo bien recuperamos
        If Not IsNull(!IdUnidad) Then mProps.IdUnidad = !IdUnidad
        If Not IsNull(!CodUnidad) Then mProps.CodUnidad = !CodUnidad
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Function AddDetalle() As Long
    Dim nuevo As PRProcesoDetalle
    
    Set nuevo = New PRProcesoDetalle
    Set nuevo.PRProceso = Me
    mcolDetalles.Add nuevo
    AddDetalle = mcolDetalles.Count    'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
End Function


Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    
    'Verifica si ya existe un Kardex usando el mismo IdProceso
    sql = "SELECT Id FROM IVKardex WHERE IdProceso=" & mProps.IdProceso
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "PRProceso_Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close
    'Controlar  o  eliminar  todos los registros
    'de  la tabla  IVMateria
    'DiegoProd
    sql = "DELETE FROM IVMateria Where IdMateria = " & mProps.IdProceso
    mProps.objEmpresa.Execute sql, True
    
    'Elimina el registro
    sql = "DELETE FROM PRProceso WHERE IdProceso=" & mProps.IdProceso
    mProps.objEmpresa.Execute sql, True

    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_ITEM", "Item Eliminado: " & _
                                                         mProps.CodProceso, "DT"
                                                         
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PRPROCESO
End Sub

Public Function Clone() As PRProceso
    Dim nuevo As PRProceso
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New PRProceso
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdProceso = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Friend Property Get IdProceso() As Long
    IdProceso = mProps.IdProceso
End Property


Public Property Let CodProceso(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PRProceso_CodProceso", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodProceso = value
    mbooModificado = True
End Property

Public Property Get CodProceso() As String
    CodProceso = Trim$(mProps.CodProceso)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PRProceso_Descripcion", MSGERR_INVALIDO
    End If
    
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, ";") > 0 Then
        Err.Raise ERR_INVALIDO, "PRProceso_Descripcion", _
            "El caracter ';' no es válido para la descripción."
    End If
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, "|") > 0 Then
        Err.Raise ERR_INVALIDO, "PRProceso_Descripcion", _
            "El caracter '|' no es válido para la descripción."
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property


Public Property Let Observacion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PRProceso_Observacion", MSGERR_INVALIDO
    End If
    
    mProps.Observacion = value
    mbooModificado = True
End Property

Public Property Get Observacion() As String
    Observacion = Trim$(mProps.Observacion)
End Property


Public Property Get NumFamiliaDetalle() As Long
    If mcolDetalleFamilia Is Nothing Then
        NumFamiliaDetalle = 0
    Else
        NumFamiliaDetalle = mcolDetalleFamilia.Count
    End If
    'If ObtieneUltimo(False) Then PrecioUltimoEgreso = mUEgreso.Precio
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

Friend Property Let Modificado(ByVal value As Boolean)
    mbooModificado = value
End Property


Public Property Get CountDetalles() As Long
    If mcolDetalles Is Nothing Then RecuperaDetalle
    
    CountDetalles = mcolDetalles.Count
End Property

Private Sub CargaDetalleFamilia()
    Dim rs As Recordset, sql As String, obj As IVFamiliaDetalle
    Dim i As Long
    Set mcolDetalleFamilia = New Collection
    'If Me.tipo = Preparacion Then
       sql = "Select IV.IdProceso, Iv.CodProceso, Iv.Descripcion, " & _
              "Iv.CodUnidad, Ivm.Cantidad, iv.costoultimoingreso   From VWIvProcesoRECUPERAR Iv " & _
            "INNER JOIN IVMateria Ivm "
'    ElseIf Me.tipo = CambioPresentacion Then
'        sql = "Select IV.IdProceso, Iv.CodProceso, Iv.Descripcion, " & _
'              "Iv.CodUnidad, Ivm.Cantidad  From VWIvProcesoRECUPERAR Iv " & _
'           "INNER JOIN IVMateria Ivm "
'    ElseIf Me.tipo = Promocion Then
'        sql = "Select IV.IdProceso, Iv.CodProceso, Iv.Descripcion, " & _
'             " Iv.CodUnidad,Ivm.Cantidad, ivm.xcuanto, ivm.tarifajornal   From VWIvProcesoRECUPERAR Iv " & _
'            "INNER JOIN IVMateria Ivm "
'    Else
'        sql = "Select IV.IdProceso, Iv.CodProceso, Iv.Descripcion, " & _
'              "Iv.codunidad, Ivm.Cantidad From  VWIvProcesoRECUPERAR IV " & _
'              "INNER JOIN IVMateria Ivm "
'
'    End If
'    03/10/2005
 '   If Me.tipo = CambioPresentacion Or Me.tipo = Preparacion Or Me.tipo = Promocion Then
       sql = sql & "  ON IV.IdProceso = Ivm.Idmateria"
        sql = sql & " Where Ivm.IdProceso = " & mProps.IdProceso
'    Else
'          sql = sql & " ON IV.IdProceso = Ivm.IdProceso "
'            sql = sql & " Where Ivm.IdMateria = " & mProps.IdProceso
'    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddDetalleFamilia
            Set obj = mcolDetalleFamilia.item(i)
'            obj.IdProceso = !IdProceso
'            obj.CodProceso = !CodProceso
'            obj.Descripcion = !Descripcion
'            If Not IsNull(!CodUnidad) Then obj.Unidad = !CodUnidad 'AUC 'cambiado antes unidad
'           obj.Cantidad = !Cantidad
'            If Me.tipo = Preparacion Then
'                If obj.PRProceso.Empresa.GNOpcion.ObtenerValor("itemPreparacion") = 0 Then 'AUC toma el p.costo
'                   If Not IsNull(!CostoUltimoIngreso) Then obj.CostoUltimoIngreso = !CostoUltimoIngreso
'                 ElseIf obj.PRProceso.Empresa.GNOpcion.ObtenerValor("itemPreparacion") = 1 Then 'AUC toma el p. venta
'                    obj.pVenta = obj.ObtienePrecio(obj.CodProceso)
'                 End If
'            ElseIf Me.tipo = Promocion Then
'                    obj.xCuanto = !xCuanto
'                    obj.Precio = !tarifajornal
'          End If
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

Private Sub RecuperaDetalle()
    Dim sql As String, i As Long, obj As PRProcesoDetalle
    Dim rs As Recordset

    Set mcolDetalles = New Collection
    
    'Recupera los detalles de existencia
    sql = "SELECT IVExist.IdProceso, PRProceso.CodProceso, " & _
          "IVExist.IdBodega, IVBodega.CodBodega, IVExist.Exist " & _
          "FROM (IVExist INNER JOIN PRProceso ON IVExist.IdProceso=" & _
          "PRProceso.IdProceso) INNER JOIN IVBodega ON IVExist.IdBodega " & _
          "= IVBodega.IdBodega " & _
          "WHERE IVExist.IdProceso=" & mProps.IdProceso & " ORDER BY CodBodega"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddDetalle
            Set obj = mcolDetalles.item(i)
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


Private Sub Class_Initialize()
    'Valores predeterminados
    With mProps
'        .CodMoneda = MONEDA_PRE
        .BandValida = True
    End With
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
    Dim nuevo As PRProcesoDetalle
    
    If mcolDetalleFamilia Is Nothing Then CargaDetalleFamilia
    Set nuevo = New PRProcesoDetalle
    Set nuevo.PRProceso = Me  'Me  olvide  del padre
    'Proporciona valores predeterminados
    nuevo.Cantidad = 1
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

'AUC 25/11/05
Private Function verificacionDuiplicadoProv(ByVal IdProveedor As String) As Boolean
    Dim rs As Recordset, sql As String
    verificacionDuiplicadoProv = False
    'Recupera los detalles de existencia
    sql = "SELECT Idproveedor FROM ivdetalleProv where idproveedor=" & IdProveedor
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        verificacionDuiplicadoProv = True
    End If
    Set rs = Nothing
End Function


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
Public Property Get IdUnidad() As Long
    IdUnidad = mProps.IdUnidad
End Property

Public Property Let IdUnidad(ByVal value As Long)
    mProps.IdUnidad = value
    mbooModificado = True
End Property

Public Function ExistenciaPreparacion(ByRef idInv As Long) As Currency
    Dim rs As Recordset
    Dim sql As String
    sql = "select exist from ivexist inner join PRProceso iv " & _
           "on iv.idProceso = ivexist.idProceso " & _
           "inner join ivbodega ivb on ivb.idbodega = ivexist.idbodega " & _
           " Where iv.IdProceso = " & idInv
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not IsNull(rs) Then
                ExistenciaPreparacion = rs!exist
                Exit Function
           End If
   Set rs = Nothing
End Function

Public Property Let BandValida(ByVal value As Boolean)
    If value <> mProps.BandValida Then
        mProps.BandValida = value
        mbooModificado = True
    End If
End Property

Public Property Get BandValida() As Boolean
    BandValida = mProps.BandValida
End Property

