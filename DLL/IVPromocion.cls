VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVPromocion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***jeaa  11/10/2010
Private Type T_PROPS
    IdPromocion As Long
    CodPromocion As String
    Descripcion As String
    BandValida As Boolean
    BandCaducidad As Boolean
    
    FechaDesde As Date
    FechaHasta As Date
    HoraDesde As Date
    HoraHasta As Date
    NumIVGrupoD  As Byte
    NumIVGrupo  As Byte
    fechagrabado As Date
    
    BandTotal(1 To 5) As Boolean
    NumTotal(1 To 5) As Byte
    ValorTotal(1 To 5) As Currency
    
    IDSucursal As Long                'jeaa 05/02/2009
    CodSucursal As String           'jeaa 05/02/2009
    
    CodSucursales As String                'jeaa 20/01/2011
    
    NumPCGrupoD  As Byte
    NumPCGrupo  As Byte
    
    BandSoloUno As Boolean
    
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mcolCondPromocionDetalle As Collection  '*** jeaa
Private mcolCondPromocionDetalleIVG As Collection  '*** jeaa
Private mcolCondPromocionDetalleP As Collection  '*** jeaa
Private mcolCondPromocionItemDup As Collection  '*** jeaa
Private mcolCondPromocionDetallePCG As Collection  '*** jeaa

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Public Sub Grabar()
    Dim sql As String, rs As Recordset, i As Integer
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    
    'Verifica si no repite el código
    sql = "SELECT codPromocion FROM IVPromocion "
    sql = sql & " WHERE codPromocion='" & mProps.CodPromocion & "'"
    sql = sql & " and idPromocion<> " & mProps.IdPromocion

    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "IVPromocion", MSGERR_REPITECODIGO
    End If
    rs.Close
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM IVPromocion WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM IVPromocion WHERE IdPromocion=" & mProps.IdPromocion
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)

        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "IVPromocion_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro
    With rs
        'jeaa 05/10/2010
        !CodPromocion = mProps.CodPromocion
        !Descripcion = mProps.Descripcion
        !BandValida = mProps.BandValida
        !BandCaducidad = mProps.BandCaducidad
        !FechaDesde = mProps.FechaDesde
        !FechaHasta = mProps.FechaHasta
        !HoraDesde = mProps.HoraDesde
        !HoraHasta = mProps.HoraHasta
        !NumIVGrupoD = mProps.NumIVGrupoD
        !NumIVGrupo = mProps.NumIVGrupo
        !NumPCGrupoD = mProps.NumPCGrupoD
        !NumPCGrupo = mProps.NumPCGrupo
        
        For i = 1 To 5 'AUC cambiado antes 4
            .Fields("BandTotal" & i) = mProps.BandTotal(i)
            .Fields("NumTotal" & i) = mProps.NumTotal(i)
            .Fields("ValorTotal" & i) = mProps.ValorTotal(i)
        Next i
        !IDSucursal = mProps.IDSucursal
        !CodSucursales = mProps.CodSucursales
        !BandSoloUno = mProps.BandSoloUno
        !fechagrabado = Now
        .Update
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IdPromocion = !IdPromocion
        .Close
    End With
    Set rs = Nothing
    GrabaCondPromocionDetalle
    GrabaCondPromocionDetalleIVG
    GrabaCondPromocionDetalleP
    GrabaCondPromocionItemDup
    GrabaCondPromocionDetallePCG
    
    
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_PRO", "Descto Nuevo: " & _
                                                         mProps.CodPromocion, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_PRO", "Descto Modificado: " & _
                                                         mProps.CodPromocion, "DT"
    End If
    
    mbooEsNuevo = False
    mbooModificado = False
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVPROMOCION
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, i As Integer
    Dim rs As Recordset
    Dim v As Variant

    v = Split(key, ",")
        If VarType(key) = vbString Then
            sql = "SELECT *"
            sql = sql & " FROM IVPromocion "
            sql = sql & " WHERE CodPromocion='" & key & "'"
        Else
            sql = "SELECT *"
            sql = sql & " FROM IVPromocion "
            sql = sql & " WHERE IdPromocion=" & key
        End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'Recupera los datos de Usuario
    With rs
        mProps.IdPromocion = !IdPromocion
        If Not IsNull(!CodPromocion) Then mProps.CodPromocion = !CodPromocion
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!BandCaducidad) Then mProps.BandCaducidad = !BandCaducidad
        If Not IsNull(!NumIVGrupoD) Then mProps.NumIVGrupoD = !NumIVGrupoD
        If Not IsNull(!NumIVGrupo) Then mProps.NumIVGrupo = !NumIVGrupo
        If Not IsNull(!NumPCGrupoD) Then mProps.NumPCGrupoD = !NumPCGrupoD
        If Not IsNull(!NumPCGrupo) Then mProps.NumPCGrupo = !NumPCGrupo
        If Not IsNull(!BandSoloUno) Then mProps.BandSoloUno = !BandSoloUno
        
        
        
        mProps.FechaDesde = !FechaDesde
        mProps.FechaHasta = !FechaHasta
        mProps.HoraDesde = !HoraDesde
        mProps.HoraHasta = !HoraHasta
        mProps.fechagrabado = !fechagrabado
        
        For i = 1 To 5
            If Not IsNull(.Fields("BandTotal" & i)) Then mProps.BandTotal(i) = .Fields("BandTotal" & i)
            If Not IsNull(.Fields("NumTotal" & i)) Then mProps.NumTotal(i) = .Fields("NumTotal" & i)
            If Not IsNull(.Fields("ValorTotal" & i)) Then mProps.ValorTotal(i) = .Fields("ValorTotal" & i)
        Next i
        If Not IsNull(!IDSucursal) Then mProps.IDSucursal = !IDSucursal
        If Not IsNull(!CodSucursales) Then mProps.CodSucursales = !CodSucursales
        .Close
    End With
    'Recupera  el detalle
        CargaCondPromoDetalle
        RecuperaCondPromoconDetalleIVG
        CargaCondPromoDetalleP
        CargaCondPromoItemDup
        RecuperaCondPromoconDetallePCG
        
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
  

    sql = "DELETE FROM IVCondPromocionDetalle Where IdPromocion = " & mProps.IdPromocion
    mProps.objEmpresa.Execute sql, True

    sql = "DELETE FROM IVCondPromocionDetalleIVG Where IdPromocion = " & mProps.IdPromocion
    mProps.objEmpresa.Execute sql, True

    sql = "DELETE FROM IVCondPromocionDetalleP Where IdPromocion = " & mProps.IdPromocion
    mProps.objEmpresa.Execute sql, True

    sql = "DELETE FROM IVCondPromocionItemDup Where IdPromocion = " & mProps.IdPromocion
    mProps.objEmpresa.Execute sql, True
  
    sql = "DELETE FROM IVCondPromocionDetallePCG Where IdPromocion = " & mProps.IdPromocion
    mProps.objEmpresa.Execute sql, True
  
  
    'Elimina el registro
    sql = "DELETE FROM IVPromocion WHERE IdPromocion=" & mProps.IdPromocion
    mProps.objEmpresa.Execute sql, True

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVPROMOCION
End Sub

Public Function Clone() As IVPromocion
    Dim nuevo As IVPromocion
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New IVPromocion
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdPromocion = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Get IdPromocion() As Long
    IdPromocion = mProps.IdPromocion
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


Private Sub Class_Initialize()
    'Valores predeterminados
    mbooEsNuevo = True
    mbooModificado = False
    mProps.BandValida = True
    mProps.BandCaducidad = False
    mProps.BandSoloUno = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Property Let CodPromocion(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "IVPromocion_CodPromocion", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodPromocion = value
    mbooModificado = True
End Property

Public Property Get CodPromocion() As String
     CodPromocion = Trim$(mProps.CodPromocion)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 254 Then
        Err.Raise ERR_INVALIDO, "IVPromocion_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let BandValida(ByVal value As Boolean)
    mProps.BandValida = value
    mbooModificado = True
End Property

Public Property Get BandValida() As Boolean
    BandValida = mProps.BandValida
End Property

Public Property Let BandCaducidad(ByVal value As Boolean)
    mProps.BandCaducidad = value
    mbooModificado = True
End Property

Public Property Get BandCaducidad() As Boolean
    BandCaducidad = mProps.BandCaducidad
End Property

Public Property Let FechaDesde(value As Date)
        mProps.FechaDesde = value
        mbooModificado = True
End Property

Public Property Get FechaDesde() As Date
    FechaDesde = mProps.FechaDesde
End Property

Public Property Let FechaHasta(value As Date)
        mProps.FechaHasta = value
        mbooModificado = True
End Property

Public Property Get FechaHasta() As Date
    FechaHasta = mProps.FechaHasta
End Property


Public Property Let HoraDesde(value As Date)
        mProps.HoraDesde = value
        mbooModificado = True
End Property

Public Property Get HoraDesde() As Date
    HoraDesde = mProps.HoraDesde
End Property

Public Property Let HoraHasta(value As Date)
        mProps.HoraHasta = value
        mbooModificado = True
End Property

Public Property Get HoraHasta() As Date
    HoraHasta = mProps.HoraHasta
End Property


Public Property Let NumIVGrupoD(ByVal value As Byte)
    mProps.NumIVGrupoD = value
    mbooModificado = True
End Property

Public Property Get NumIVGrupoD() As Byte
    NumIVGrupoD = mProps.NumIVGrupoD
End Property

Public Property Let NumIVGrupo(ByVal value As Byte)
    mProps.NumIVGrupo = value
    mbooModificado = True
End Property

Public Property Get NumIVGrupo() As Byte
    NumIVGrupo = mProps.NumIVGrupo
End Property

'************** coleccion de promo x item
Private Sub BorrarCondPromocionDetalle()
    'Borra  toda  la informacion de la coleccion
     If mcolCondPromocionDetalle Is Nothing Then Set mcolCondPromocionDetalle = New Collection
    With mcolCondPromocionDetalle
        Do While .Count > 0
            .Remove .Count
        Loop
    End With
    mbooModificado = True
End Sub

Public Property Get IvCondPromocionDetalle(ix As Long) As IvCondPromocionDetalle
    If mcolCondPromocionDetalle Is Nothing Then RecuperaCondPromocionDetalle (ix)  'posilbe recursividad
    Set mcolCondPromocionDetalle = mcolCondPromocionDetalle.item(ix)
End Property

Public Property Get NumCondPromocionDetalle() As Long
    If mcolCondPromocionDetalle Is Nothing Then
        NumCondPromocionDetalle = 0
    Else
        NumCondPromocionDetalle = mcolCondPromocionDetalle.Count
    End If
End Property

Public Function AddCondPromocionDetalle() As Long
    Dim nuevo As IvCondPromocionDetalle
    
    If mcolCondPromocionDetalle Is Nothing Then CargaCondPromoDetalle
    Set nuevo = New IvCondPromocionDetalle
    Set nuevo.IVPromocion = Me  'Me  olvide  del padre
    'Proporciona valores predeterminados
    nuevo.xCuanto = 0
    nuevo.xCuantoPeso = 0
    nuevo.xCuantoUSD = 0
    
    mcolCondPromocionDetalle.Add nuevo
    AddCondPromocionDetalle = mcolCondPromocionDetalle.Count    'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    
    mbooModificado = True
    
End Function




Private Sub CargaCondPromoDetalle()
    Dim rs As Recordset, sql As String, obj As IvCondPromocionDetalle
    Dim i As Long
    Set mcolCondPromocionDetalle = New Collection
        sql = " select ivcp.*,  IV.IdInventario, Iv.CodInventario, Iv.Descripcion, ivu.codunidad "
        sql = sql & " from IVCondPromocionDetalle ivcp"
        sql = sql & " INNER JOIN IVInventario Iv "
        sql = sql & " left JOIN IVUnidad Ivu "
        sql = sql & " ON IV.Idunidad = ivu.Idunidad "
        sql = sql & " ON IV.IdInventario = ivcp.IdInventario "
        sql = sql & " Where Ivcp.IdPromocion = " & mProps.IdPromocion
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddCondPromocionDetalle
            Set obj = mcolCondPromocionDetalle.item(i)
            obj.CodInventario = !CodInventario
            obj.Descripcion = !Descripcion
            obj.Orden = !Orden
            If Not IsNull(!xCuanto) Then obj.xCuanto = !xCuanto

            If Not IsNull(!xCuantoUSD) Then obj.xCuantoUSD = !xCuantoUSD
           .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
    mbooModificado = False
End Sub


Public Function RecuperaCondPromocionDetalle(ix As Long) As IvCondPromocionDetalle
    Set RecuperaCondPromocionDetalle = mcolCondPromocionDetalle.item(ix)
End Function


Public Sub RemoveCondPromocionDetalle(ByVal ix As Long, Optional obj As IvCondPromocionDetalle)
    If mcolCondPromocionDetalle Is Nothing Then CargaCondPromocionDetalle
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolCondPromocionDetalle.Count
            If mcolCondPromocionDetalle.item(ix) Is obj Then
                mcolCondPromocionDetalle.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolCondPromocionDetalle.Remove ix
    End If
    
    'mcolDetalleFamilia.Remove ix
    mbooModificado = True
End Sub

Private Sub CargaCondPromocionDetalle()
    Dim rs As Recordset, sql As String, obj As IvCondPromocionDetalle
    Dim i As Long
    Set mcolCondPromocionDetalle = New Collection
        sql = "Select IV.IdInventario, Iv.CodInventario, Iv.Descripcion, " & _
             " Iv.CodUnidad,Ivm.Cantidad, ivm.xcuanto, ivm.tarifajornal   From VWIvinventarioRECUPERAR Iv " & _
            "INNER JOIN IVMateria Ivm "
        sql = sql & "  ON IV.IdInventario = Ivm.Idmateria"
        sql = sql & " Where Ivm.IdInventario = " & mProps.IdPromocion
        sql = sql & " order by orden"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddCondPromocionDetalle
            Set obj = mcolCondPromocionDetalle.item(i)
                  obj.IdPromocion = !IdInventario
                  obj.CodInventario = !CodInventario
                  obj.Descripcion = !Descripcion
                  obj.xCuanto = !xCuanto
                  obj.xCuantoPeso = !xCuantoPeso
                  obj.xCuantoUSD = !xCuantoUSD

           .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub


Public Sub GrabaCondPromocionDetalle()
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As IvCondPromocionDetalle, sql As String, band As Boolean
    Dim NumHijos As Integer, CountNumHijos As Integer
    On Error GoTo CapturaError
    
    'Elimina  todos  los registros
    sql = "Delete "
    sql = sql & " From IVCondPromocionDetalle "
    sql = sql & " Where IdPromocion = " & mProps.IdPromocion
   
    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    band = True
    If Not mcolCondPromocionDetalle Is Nothing Then
        sql = "SELECT * FROM IVCondPromocionDetalle WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        NumHijos = mcolCondPromocionDetalle.Count
        CountNumHijos = 0
        For i = 1 To mcolCondPromocionDetalle.Count
            If CountNumHijos < NumHijos Then
                CountNumHijos = CountNumHijos + 1
                Set obj = mcolCondPromocionDetalle.item(i)
'                If obj.IdPromocion = 0 Then
'                        obj.Recuperar obj.CodInventario
'                        If VerificacionItemDuiplicado(obj.Idinventario) Then
'                            RemoveCondPromocionDetalle i
'                            i = i - 1
'                        End If
'                End If
                If band Then
                    rs1.AddNew
                    rs1!IdPromocion = Me.IdPromocion
                    rs1!IdInventario = obj.IdInventario
                    rs1!Orden = obj.Orden
                    rs1!xCuanto = obj.xCuanto
                    rs1!xCuantoUSD = obj.xCuantoUSD
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

Private Function VerificacionItemDuiplicado(ByVal IdInventario As String) As Boolean
    Dim rs As Recordset, sql As String
    VerificacionItemDuiplicado = False
    'Recupera los detalles de existencia
    sql = "SELECT IdInventario FROM IVCondPromocionDetalle where idinventario =" & IdInventario
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        VerificacionItemDuiplicado = True
    End If
    Set rs = Nothing
End Function

'************** coleccion de promo x IVGrupo

Public Property Get NumCondPromocionDetalleIVG() As Long
    If mcolCondPromocionDetalleIVG Is Nothing Then
        NumCondPromocionDetalleIVG = 0
    Else
        NumCondPromocionDetalleIVG = mcolCondPromocionDetalleIVG.Count
    End If
End Property

Public Function AddCondPromocionDetalleIVG() As Long
    Dim nuevo As IvCondPromocionDetalleIVG
    
    If mcolCondPromocionDetalleIVG Is Nothing Then RecuperaCondPromoconDetalleIVG
    Set nuevo = New IvCondPromocionDetalleIVG
    Set nuevo.IVPromocion = Me  'Me  olvide  del padre
    'Proporciona valores predeterminados
    nuevo.xCuanto = 0
    nuevo.xCuantoPeso = 0
    nuevo.xCuantoUSD = 0
    
    mcolCondPromocionDetalleIVG.Add nuevo
    AddCondPromocionDetalleIVG = mcolCondPromocionDetalleIVG.Count    'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    
    mbooModificado = True
    
End Function




Private Sub CargaCondPromoDetalleIVG(ByVal NumGrupo As Integer)
    Dim rs As Recordset, sql As String, obj As IvCondPromocionDetalleIVG
    Dim i As Long
    Set mcolCondPromocionDetalleIVG = New Collection
        sql = " select ivcp.*,  IVG.IdGrupo" & NumGrupo & " as IdGrupo, "
        sql = sql & " IvG.Codgrupo" & NumGrupo & " as CodGrupo, Ivg.Descripcion "
        sql = sql & " from IVCondPromocionDetalleIVG ivcp"
        sql = sql & " INNER JOIN IVGrupo" & NumGrupo & " IVG"
        sql = sql & " ON IVG.IdGrupo" & NumGrupo & " = ivcp.IdIVGrupo "
        sql = sql & " Where Ivcp.IdPromocion = " & mProps.IdPromocion
    
        sql = " select * "
        sql = sql & " from IVCondPromocionDetalleIVG ivcp"
        sql = sql & " Where Ivcp.IdPromocion = " & mProps.IdPromocion
    
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddCondPromocionDetalleIVG
            Set obj = mcolCondPromocionDetalleIVG.item(i)
            obj.CodIVGrupo = !CodGrupo
'            obj.DescripcionIVGrupo = !Descripcion
            obj.Orden = !Orden
            If Not IsNull(!xCuanto) Then obj.xCuanto = !xCuanto
            If Not IsNull(!xCuantoPeso) Then obj.xCuanto = !xCuanto
            If Not IsNull(!xCuantoUSD) Then obj.xCuantoUSD = !xCuantoUSD
           .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
    mbooModificado = False
End Sub




Public Sub RemoveCondPromocionDetalleIVG(ByVal ix As Long, Optional obj As IvCondPromocionDetalleIVG)
    If mcolCondPromocionDetalleIVG Is Nothing Then CargaCondPromocionDetalleIVG (mProps.NumIVGrupoD)
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolCondPromocionDetalleIVG.Count
            If mcolCondPromocionDetalleIVG.item(ix) Is obj Then
                mcolCondPromocionDetalleIVG.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolCondPromocionDetalleIVG.Remove ix
    End If
    
    'mcolDetalleFamilia.Remove ix
    mbooModificado = True
End Sub

Private Sub CargaCondPromocionDetalleIVG(ByVal NumGrupo As Integer)
    Dim rs As Recordset, sql As String, obj As IvCondPromocionDetalleIVG
    Dim i As Long
    Set mcolCondPromocionDetalleIVG = New Collection
    
        sql = " select ivcp.*,  IVG.IdGrupo" & NumGrupo & " as IdGrupo, "
        sql = sql & " IvG.Codgrupo" & NumGrupo & " as CodGrupo, Iv.Descripcion "
        sql = sql & " from IVCondPromocionDetalleIVG ivcp"
        sql = sql & " INNER JOIN IVGrupo" & NumGrupo & " IVG"
        sql = sql & " ON IVG.IdGrupo" & NumGrupo & " = ivcp.IdIVGrupo "
        sql = sql & " Where Ivcp.IdPromocion = " & mProps.IdPromocion
        sql = sql & " order by orden"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddCondPromocionDetalleIVG
            Set obj = mcolCondPromocionDetalleIVG.item(i)
                  obj.IdPromocion = !IdInventario
                  obj.CodIVGrupo = !CodGrupo
'                  obj.DescripcionIVGrupo = !Descripcion
                  obj.xCuanto = !xCuanto
                  obj.xCuantoPeso = !xCuantoPeso
                  obj.xCuantoUSD = !xCuantoUSD
           .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub


Public Sub GrabaCondPromocionDetalleIVG()
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As IvCondPromocionDetalleIVG, sql As String, band As Boolean
    Dim NumHijos As Integer, CountNumHijos As Integer
    On Error GoTo CapturaError
    
    'Elimina  todos  los registros
    sql = "Delete "
    sql = sql & " From IVCondPromocionDetalleIVG "
    sql = sql & " Where IdPromocion = " & mProps.IdPromocion
   
    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    band = True
    If Not mcolCondPromocionDetalleIVG Is Nothing Then
        sql = "SELECT * FROM IVCondPromocionDetalleIVG WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        NumHijos = mcolCondPromocionDetalleIVG.Count
        CountNumHijos = 0
        For i = 1 To mcolCondPromocionDetalleIVG.Count
            If CountNumHijos < NumHijos Then
                CountNumHijos = CountNumHijos + 1
                Set obj = mcolCondPromocionDetalleIVG.item(i)
'                If obj.IdPromocion = 0 Then
'                        obj.Recuperar obj.CodInventario
'                        If VerificacionItemDuiplicado(obj.Idinventario) Then
'                            RemoveCondPromocionDetalleIVG i
'                            i = i - 1
'                        End If
'                End If
                If band Then
                    rs1.AddNew
                    rs1!IdPromocion = Me.IdPromocion
                    rs1!IDIVGrupo = obj.IDIVGrupo
                    rs1!Orden = obj.Orden
                    rs1!xCuanto = obj.xCuanto
                    rs1!xCuantoPeso = obj.xCuantoPeso
                    rs1!xCuantoUSD = obj.xCuantoUSD
                    rs1!NumIVGrupo = obj.NumIVGrupo
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

Private Function VerificacionItemDuiplicadoIVG(ByVal IdInventario As String) As Boolean
    Dim rs As Recordset, sql As String
    VerificacionItemDuiplicadoIVG = False
    'Recupera los detalles de existencia
    sql = "SELECT IdInventario FROM IVCondPromocionDetalleIVG where idinventario =" & IdInventario
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        VerificacionItemDuiplicadoIVG = True
    End If
    Set rs = Nothing
End Function



Public Property Let ValorTotal(ByVal ix As Integer, ByVal value As Currency)
   
    'Si es que realmente modifica el valor
    If mProps.ValorTotal(ix + 1) <> value Then
        mProps.ValorTotal(ix + 1) = value
        mbooModificado = True
    End If
End Property

Public Property Get ValorTotal(ByVal ix As Integer) As Currency
    ValorTotal = mProps.ValorTotal(ix)
End Property


Public Property Let BandTotal(ByVal ix As Integer, ByVal value As Boolean)
    'Si es que realmente modifica el valor
    If mProps.BandTotal(ix + 1) <> value Then
        mProps.BandTotal(ix + 1) = value
        mbooModificado = True
    End If
End Property

Public Property Get BandTotal(ByVal ix As Integer) As Boolean
    BandTotal = mProps.BandTotal(ix)
End Property


Public Property Let NumTotal(ByVal ix As Integer, ByVal value As Byte)
    
    'Si es que realmente modifica el valor
    If mProps.NumTotal(ix) <> value Then
        mProps.NumTotal(ix) = value
        mbooModificado = True
    End If
End Property

Public Property Get NumTotal(ByVal ix As Integer) As Byte
    NumTotal = mProps.NumTotal(ix)
End Property




Private Sub RecuperaCondPromoconDetalleIVG()
    Dim sql As String, rs As Recordset, i As Long, mod_antes As Boolean
    Dim nuevo As IvCondPromocionDetalleIVG
    mod_antes = mbooModificado      'Guarda el estado de Modificado
    If mcolCondPromocionDetalleIVG Is Nothing Then Set mcolCondPromocionDetalleIVG = New Collection
    
    BorrarCondPromoconDetalleIVG
    
    'Recupera los IVKardex
    sql = "SELECT * FROM IVCondPromocionDetalleIVG "
        sql = sql & " WHERE Idpromocion=" & mProps.IdPromocion & " ORDER BY Orden"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            Set nuevo = New IvCondPromocionDetalleIVG
            mcolCondPromocionDetalleIVG.Add nuevo
            i = mcolCondPromocionDetalleIVG.Count
            Set IvCondPromocionDetalleIVG(i).IVPromocion = Me
            IvCondPromocionDetalleIVG(i).Recuperar rs
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    mbooModificado = mod_antes      'Para que no cambie el estado de Modificado
End Sub

Private Sub BorrarCondPromoconDetalleIVG()
    'Borra  toda  la informacion de la coleccion
     If mcolCondPromocionDetalleIVG Is Nothing Then Set mcolCondPromocionDetalleIVG = New Collection
    With mcolCondPromocionDetalleIVG
        Do While .Count > 0
            .Remove .Count
        Loop
    End With
    mbooModificado = True
End Sub

Public Property Get IvCondPromocionDetalleIVG(ix As Long) As IvCondPromocionDetalleIVG
    If mcolCondPromocionDetalleIVG Is Nothing Then RecuperaCondPromoconDetalleIVG   'posilbe recursividad
    Set IvCondPromocionDetalleIVG = mcolCondPromocionDetalleIVG.item(ix)
End Property
 
 '*********** item de promocion a ser entragados
 
 Private Sub BorrarCondPromocionDetallePP()
    'Borra  toda  la informacion de la coleccion
     If mcolCondPromocionDetalleP Is Nothing Then Set mcolCondPromocionDetalleP = New Collection
    With mcolCondPromocionDetalleP
        Do While .Count > 0
            .Remove .Count
        Loop
    End With
    mbooModificado = True
End Sub

Public Property Get IvCondPromocionDetalleP(ix As Long) As IvCondPromocionDetalleP
    If mcolCondPromocionDetalleP Is Nothing Then RecuperaCondPromocionDetalleP (ix)  'posilbe recursividad
    Set mcolCondPromocionDetalleP = mcolCondPromocionDetalleP.item(ix)
End Property

Public Property Get NumCondPromocionDetalleP() As Long
    If mcolCondPromocionDetalleP Is Nothing Then
        NumCondPromocionDetalleP = 0
    Else
        NumCondPromocionDetalleP = mcolCondPromocionDetalleP.Count
    End If
End Property

Public Function AddCondPromocionDetalleP() As Long
    Dim nuevo As IvCondPromocionDetalleP
    
    If mcolCondPromocionDetalleP Is Nothing Then CargaCondPromoDetalleP
    Set nuevo = New IvCondPromocionDetalleP
    Set nuevo.IVPromocion = Me  'Me  olvide  del padre
    'Proporciona valores predeterminados
    nuevo.cant = 0
    nuevo.PrecioU = 0
    
    mcolCondPromocionDetalleP.Add nuevo
    AddCondPromocionDetalleP = mcolCondPromocionDetalleP.Count    'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    
    mbooModificado = True
    
End Function

Private Sub CargaCondPromoDetalleP()
    Dim rs As Recordset, sql As String, obj As IvCondPromocionDetalleP
    Dim i As Long
    Set mcolCondPromocionDetalleP = New Collection
        sql = " select ivcp.*,  IV.IdInventario, Iv.CodInventario, Iv.Descripcion, ivu.codunidad "
        sql = sql & " from IVCondPromocionDetalleP ivcp"
        sql = sql & " INNER JOIN IVInventario Iv "
        sql = sql & " left JOIN IVUnidad Ivu "
        sql = sql & " ON IV.Idunidad = ivu.Idunidad "
        sql = sql & " ON IV.IdInventario = ivcp.IdInventario "
        sql = sql & " Where Ivcp.IdPromocion = " & mProps.IdPromocion
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddCondPromocionDetalleP
            Set obj = mcolCondPromocionDetalleP.item(i)
            obj.CodInventario = !CodInventario
            obj.Descripcion = !Descripcion
            obj.Orden = !Orden
            If Not IsNull(!cant) Then obj.cant = !cant
            If Not IsNull(!PrecioU) Then obj.PrecioU = !PrecioU
           .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
    mbooModificado = False
End Sub


Public Function RecuperaCondPromocionDetalleP(ix As Long) As IvCondPromocionDetalleP
    Set RecuperaCondPromocionDetalleP = mcolCondPromocionDetalleP.item(ix)
End Function


Public Sub RemoveCondPromocionDetalleP(ByVal ix As Long, Optional obj As IvCondPromocionDetalleP)
    If mcolCondPromocionDetalleP Is Nothing Then CargaCondPromocionDetalleP
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolCondPromocionDetalleP.Count
            If mcolCondPromocionDetalleP.item(ix) Is obj Then
                mcolCondPromocionDetalleP.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolCondPromocionDetalleP.Remove ix
    End If
    
    'mcolDetalleFamilia.Remove ix
    mbooModificado = True
End Sub

Private Sub CargaCondPromocionDetalleP()
    Dim rs As Recordset, sql As String, obj As IvCondPromocionDetalleP
    Dim i As Long
    Set mcolCondPromocionDetalleP = New Collection
        sql = "Select IV.IdInventario, Iv.CodInventario, Iv.Descripcion, " & _
             " Iv.CodUnidad,Ivm.Cantidad, ivm.xcuanto, ivm.tarifajornal   From VWIvinventarioRECUPERAR Iv " & _
            "INNER JOIN IVMateria Ivm "
        sql = sql & "  ON IV.IdInventario = Ivm.Idmateria"
        sql = sql & " Where Ivm.IdInventario = " & mProps.IdPromocion
        sql = sql & " order by orden"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddCondPromocionDetalleP
            Set obj = mcolCondPromocionDetalleP.item(i)
                  obj.IdPromocion = !IdInventario
                  obj.CodInventario = !CodInventario
                  obj.Descripcion = !Descripcion
                  obj.cant = !cant
                  obj.PrecioU = !PrecioU

           .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub


Public Sub GrabaCondPromocionDetalleP()
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As IvCondPromocionDetalleP, sql As String, band As Boolean
    Dim NumHijos As Integer, CountNumHijos As Integer
    On Error GoTo CapturaError
    
    'Elimina  todos  los registros
    sql = "Delete "
    sql = sql & " From IVCondPromocionDetalleP "
    sql = sql & " Where IdPromocion = " & mProps.IdPromocion
   
    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    band = True
    If Not mcolCondPromocionDetalleP Is Nothing Then
        sql = "SELECT * FROM IVCondPromocionDetalleP WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        NumHijos = mcolCondPromocionDetalleP.Count
        CountNumHijos = 0
        For i = 1 To mcolCondPromocionDetalleP.Count
            If CountNumHijos < NumHijos Then
                CountNumHijos = CountNumHijos + 1
                Set obj = mcolCondPromocionDetalleP.item(i)
'                If obj.IdPromocion = 0 Then
'                        obj.Recuperar obj.CodInventario
'                        If VerificacionItemDuiplicado(obj.Idinventario) Then
'                            RemoveCondPromocionDetalleP i
'                            i = i - 1
'                        End If
'                End If
                If band Then
                    rs1.AddNew
                    rs1!IdPromocion = Me.IdPromocion
                    rs1!IdInventario = obj.IdInventario
                    rs1!Orden = obj.Orden
                    rs1!cant = obj.cant
                    rs1!PrecioU = obj.PrecioU
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

Private Function VerificacionItemDuiplicadoP(ByVal IdInventario As String) As Boolean
    Dim rs As Recordset, sql As String
    VerificacionItemDuiplicadoP = False
    'Recupera los detalles de existencia
    sql = "SELECT IdInventario FROM IVCondPromocionDetalleP where idinventario =" & IdInventario
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        VerificacionItemDuiplicadoP = True
    End If
    Set rs = Nothing
End Function



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

Public Property Let CodSucursales(ByVal value As String)
    If Len(value) > 200 Then
        Err.Raise ERR_INVALIDO, "IVPromocion_CodSucursales", MSGERR_INVALIDO
    End If
    
    mProps.CodSucursales = value
    mbooModificado = True
End Property

Public Property Get CodSucursales() As String
    CodSucursales = Trim$(mProps.CodSucursales)
End Property

'---------------- ITEM QUE DUPLICA
Public Function AddCondPromocionItemDup() As Long
    Dim nuevo As IvCondPromocionItemDup
    
    If mcolCondPromocionItemDup Is Nothing Then CargaCondPromoItemDup
    Set nuevo = New IvCondPromocionItemDup
    Set nuevo.IVPromocion = Me  'Me  olvide  del padre
    'Proporciona valores predeterminados
    nuevo.xCuanto = 0
    
    mcolCondPromocionItemDup.Add nuevo
    AddCondPromocionItemDup = mcolCondPromocionItemDup.Count    'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    
    mbooModificado = True
    
End Function


Private Sub CargaCondPromoItemDup()
    Dim rs As Recordset, sql As String, obj As IvCondPromocionItemDup
    Dim i As Long
    Set mcolCondPromocionItemDup = New Collection
        sql = " select ivcp.*,  IV.IdInventario, Iv.CodInventario, Iv.Descripcion, ivu.codunidad "
        sql = sql & " from IVCondPromocionItemDup ivcp"
        sql = sql & " INNER JOIN IVInventario Iv "
        sql = sql & " left JOIN IVUnidad Ivu "
        sql = sql & " ON IV.Idunidad = ivu.Idunidad "
        sql = sql & " ON IV.IdInventario = ivcp.IdInventario "
        sql = sql & " Where Ivcp.IdPromocion = " & mProps.IdPromocion
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddCondPromocionItemDup
            Set obj = mcolCondPromocionItemDup.item(i)
            obj.CodInventario = !CodInventario
            obj.Descripcion = !Descripcion
            obj.Orden = !Orden
            If Not IsNull(!xCuanto) Then obj.xCuanto = !xCuanto

           .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
    mbooModificado = False
End Sub

Public Function RecuperaCondPromocionItemDup(ix As Long) As IvCondPromocionItemDup
    Set RecuperaCondPromocionItemDup = mcolCondPromocionItemDup.item(ix)
End Function


Public Sub RemoveCondPromocionItemDup(ByVal ix As Long, Optional obj As IvCondPromocionItemDup)
    If mcolCondPromocionItemDup Is Nothing Then CargaCondPromoItemDup
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolCondPromocionItemDup.Count
            If mcolCondPromocionItemDup.item(ix) Is obj Then
                mcolCondPromocionItemDup.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolCondPromocionItemDup.Remove ix
    End If
    
    'mcolItemDupFamilia.Remove ix
    mbooModificado = True
End Sub

Public Sub GrabaCondPromocionItemDup()
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As IvCondPromocionItemDup, sql As String, band As Boolean
    Dim NumHijos As Integer, CountNumHijos As Integer
    On Error GoTo CapturaError
    
    'Elimina  todos  los registros
    sql = "Delete "
    sql = sql & " From IVCondPromocionItemDup "
    sql = sql & " Where IdPromocion = " & mProps.IdPromocion
   
    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    band = True
    If Not mcolCondPromocionItemDup Is Nothing Then
        sql = "SELECT * FROM IVCondPromocionItemDup WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        NumHijos = mcolCondPromocionItemDup.Count
        CountNumHijos = 0
        For i = 1 To mcolCondPromocionItemDup.Count
            If CountNumHijos < NumHijos Then
                CountNumHijos = CountNumHijos + 1
                Set obj = mcolCondPromocionItemDup.item(i)
                If band Then
                    rs1.AddNew
                    rs1!IdPromocion = Me.IdPromocion
                    rs1!IdInventario = obj.IdInventario
                    rs1!Orden = obj.Orden
                    rs1!xCuanto = obj.xCuanto
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

Public Property Get NumCondPromocionItemDup() As Long
    If mcolCondPromocionItemDup Is Nothing Then
        NumCondPromocionItemDup = 0
    Else
        NumCondPromocionItemDup = mcolCondPromocionItemDup.Count
    End If
End Property

'-------------- DETALLE PCGRUPO

Public Property Get NumCondPromocionDetallePCG() As Long
    If mcolCondPromocionDetallePCG Is Nothing Then
        NumCondPromocionDetallePCG = 0
    Else
        NumCondPromocionDetallePCG = mcolCondPromocionDetallePCG.Count
    End If
End Property


Public Function AddCondPromocionDetallePCG() As Long
    Dim nuevo As IvCondPromocionDetallePCG
    
    If mcolCondPromocionDetallePCG Is Nothing Then RecuperaCondPromoconDetallePCG
    Set nuevo = New IvCondPromocionDetallePCG
    Set nuevo.IVPromocion = Me  'Me  olvide  del padre
    'Proporciona valores predeterminados
    
    mcolCondPromocionDetallePCG.Add nuevo
    AddCondPromocionDetallePCG = mcolCondPromocionDetallePCG.Count    'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    
    mbooModificado = True
    
End Function


Private Sub CargaCondPromoDetallePCG(ByVal NumGrupo As Integer)
    Dim rs As Recordset, sql As String, obj As IvCondPromocionDetallePCG
    Dim i As Long
    Set mcolCondPromocionDetallePCG = New Collection
        sql = " select ivcp.*,  PCG.IdGrupo" & NumGrupo & " as IdGrupo, "
        sql = sql & " PCG.Codgrupo" & NumGrupo & " as CodGrupo, PCG.Descripcion "
        sql = sql & " from IVCondPromocionDetallePCG ivcp"
        sql = sql & " INNER JOIN PCGrupo" & NumGrupo & " PCG"
        sql = sql & " ON PCG.IdGrupo" & NumGrupo & " = ivcp.IdPCGrupo "
        sql = sql & " Where Ivcp.IdPromocion = " & mProps.IdPromocion
    
        sql = " select * "
        sql = sql & " from IVCondPromocionDetallePCG ivcp"
        sql = sql & " Where Ivcp.IdPromocion = " & mProps.IdPromocion
    
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddCondPromocionDetallePCG
            Set obj = mcolCondPromocionDetallePCG.item(i)
            obj.CodPCGrupo = !CodGrupo
            obj.Orden = !Orden
           .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
    mbooModificado = False
End Sub


Public Sub RemoveCondPromocionDetallePCG(ByVal ix As Long, Optional obj As IvCondPromocionDetallePCG)
    If mcolCondPromocionDetallePCG Is Nothing Then CargaCondPromocionDetallePCG (mProps.NumPCGrupoD)
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolCondPromocionDetallePCG.Count
            If mcolCondPromocionDetallePCG.item(ix) Is obj Then
                mcolCondPromocionDetallePCG.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolCondPromocionDetallePCG.Remove ix
    End If
    
    'mcolDetalleFamilia.Remove ix
    mbooModificado = True
End Sub

Public Sub GrabaCondPromocionDetallePCG()
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As IvCondPromocionDetallePCG, sql As String, band As Boolean
    Dim NumHijos As Integer, CountNumHijos As Integer
    On Error GoTo CapturaError
    
    'Elimina  todos  los registros
    sql = "Delete "
    sql = sql & " From IVCondPromocionDetallePCG "
    sql = sql & " Where IdPromocion = " & mProps.IdPromocion
   
    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    band = True
    If Not mcolCondPromocionDetallePCG Is Nothing Then
        sql = "SELECT * FROM IVCondPromocionDetallePCG WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        NumHijos = mcolCondPromocionDetallePCG.Count
        CountNumHijos = 0
        For i = 1 To mcolCondPromocionDetallePCG.Count
            If CountNumHijos < NumHijos Then
                CountNumHijos = CountNumHijos + 1
                Set obj = mcolCondPromocionDetallePCG.item(i)
                If band Then
                    rs1.AddNew
                    rs1!IdPromocion = Me.IdPromocion
                    rs1!idPcGrupo = obj.idPcGrupo
                    rs1!Orden = obj.Orden
                    rs1!NumPCGrupo = obj.NumPCGrupo
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

Private Sub RecuperaCondPromoconDetallePCG()
    Dim sql As String, rs As Recordset, i As Long, mod_antes As Boolean
    Dim nuevo As IvCondPromocionDetallePCG
    mod_antes = mbooModificado      'Guarda el estado de Modificado
    If mcolCondPromocionDetallePCG Is Nothing Then Set mcolCondPromocionDetallePCG = New Collection
    
    BorrarCondPromoconDetallePCG
    
    'Recupera los IVKardex
    sql = "SELECT * FROM IVCondPromocionDetallePCG "
        sql = sql & " WHERE Idpromocion=" & mProps.IdPromocion & " ORDER BY Orden"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            Set nuevo = New IvCondPromocionDetallePCG
            mcolCondPromocionDetallePCG.Add nuevo
            i = mcolCondPromocionDetallePCG.Count
            Set IvCondPromocionDetallePCG(i).IVPromocion = Me
            IvCondPromocionDetallePCG(i).Recuperar rs
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    mbooModificado = mod_antes      'Para que no cambie el estado de Modificado
End Sub

Private Sub BorrarCondPromoconDetallePCG()
    'Borra  toda  la informacion de la coleccion
     If mcolCondPromocionDetallePCG Is Nothing Then Set mcolCondPromocionDetallePCG = New Collection
    With mcolCondPromocionDetallePCG
        Do While .Count > 0
            .Remove .Count
        Loop
    End With
    mbooModificado = True
End Sub

Public Property Get IvCondPromocionDetallePCG(ix As Long) As IvCondPromocionDetallePCG
    If mcolCondPromocionDetallePCG Is Nothing Then RecuperaCondPromoconDetallePCG   'posilbe recursividad
    Set IvCondPromocionDetallePCG = mcolCondPromocionDetallePCG.item(ix)
End Property

Public Property Let NumPCGrupoD(ByVal value As Byte)
    mProps.NumPCGrupoD = value
    mbooModificado = True
End Property

Public Property Get NumPCGrupoD() As Byte
    NumPCGrupoD = mProps.NumPCGrupoD
End Property

Public Property Let NumPCGrupo(ByVal value As Byte)
    mProps.NumPCGrupo = value
    mbooModificado = True
End Property

Public Property Get NumPCGrupo() As Byte
    NumPCGrupo = mProps.NumPCGrupo
End Property

Private Sub CargaCondPromocionDetallePCG(ByVal NumGrupo As Integer)
    Dim rs As Recordset, sql As String, obj As IvCondPromocionDetallePCG
    Dim i As Long
    Set mcolCondPromocionDetallePCG = New Collection
    
        sql = " select ivcp.*,  PCG.IdGrupo" & NumGrupo & " as IdGrupo, "
        sql = sql & " PCG.Codgrupo" & NumGrupo & " as CodGrupo, Iv.Descripcion "
        sql = sql & " from IVCondPromocionDetallePCG ivcp"
        sql = sql & " INNER JOIN PCGrupo" & NumGrupo & " PCG"
        sql = sql & " ON PCG.IdGrupo" & NumGrupo & " = ivcp.IdPCGrupo "
        sql = sql & " Where Ivcp.IdPromocion = " & mProps.IdPromocion
        sql = sql & " order by orden"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddCondPromocionDetallePCG
            Set obj = mcolCondPromocionDetallePCG.item(i)
                  obj.IdPromocion = !IdInventario
                  obj.CodPCGrupo = !CodGrupo
           .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub


Public Property Let BandSoloUno(ByVal value As Boolean)
    mProps.BandSoloUno = value
    mbooModificado = True
End Property

Public Property Get BandSoloUno() As Boolean
    BandSoloUno = mProps.BandSoloUno
End Property

