VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVReceta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdReceta As Long
    CodReceta As String
    Descripcion As String
    fecha As Date
    BandValida As Boolean
    IdInventario As Long
    CodInventario As String
    Batch As Currency
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean
Private mcolIVRecetaDetalle As Collection

Public Sub Grabar()
    Dim sql As String, rs As Recordset
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
    sql = "SELECT codReceta FROM IVReceta "
    sql = sql & " WHERE codReceta='" & mProps.CodReceta & "'"
    sql = sql & " and Id<> " & mProps.IdReceta
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "IVReceta", MSGERR_REPITECODIGO
    End If
    rs.Close
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM IVReceta WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM IVReceta WHERE Id=" & mProps.IdReceta
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)

        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fecha Then
            Err.Raise ERR_MODIFICADO, "IVReceta_Grabar", MSGERR_MODIFICADO
        End If
    End If


    'Graba el registro
    With rs
        !CodReceta = mProps.CodReceta
        !Descripcion = mProps.Descripcion
        !BandValida = mProps.BandValida
        !fechagrabado = mProps.fecha
        !IdInvPadre = mProps.IdInventario
        !Batch = mProps.Batch
        .Update
        .Move 0             'Para actualizar
        
        mProps.IdReceta = !id
        .Close
    End With
    Set rs = Nothing
    GrabaRecetaDetalle
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_IVREC", "Nueva Receta: " & _
                                                         mProps.CodReceta, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_IVRES", "Modi Receta: " & _
                                                         mProps.CodReceta, "DT"
    End If
    
    mbooEsNuevo = False
    mbooModificado = False
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVRECETA
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String
    Dim rs As Recordset
    Dim v As Variant

   ' v = Split(key, ",")
        If VarType(key) = vbString Then
            sql = "SELECT ivr.*,iv.codInventario"
            sql = sql & " FROM IVReceta ivr Left join Ivinventario iv on iv.idinventario = ivr.idinvPadre"
            sql = sql & " WHERE ivr.CodReceta='" & key & "'"
        Else
           sql = "SELECT ivr.*,iv.codInventario "
            sql = sql & " FROM IVReceta ivr Left join Ivinventario iv on iv.idinventario = ivr.idinvPadre "
            sql = sql & " WHERE Id=" & key
        End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'Recupera los datos de Usuario
    With rs
        mProps.IdReceta = !id
        If Not IsNull(!CodReceta) Then mProps.CodReceta = !CodReceta
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!fechagrabado) Then mProps.fecha = !fechagrabado
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!IdInvPadre) Then mProps.IdInventario = !IdInvPadre
        If Not IsNull(!CodInventario) Then mProps.CodInventario = !CodInventario
        If Not IsNull(!Batch) Then mProps.Batch = !Batch
        
        .Close
    End With
    'Recupera  el detalle
    RecuperaRecetaDetalle
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar(ByRef s As String)
    Dim sql As String, rs As Recordset
  
    'Verifica si existe items registrados
    sql = "select TOP 10 g.codtrans,g.numtrans FROM ivkardex ivk Inner Join GnComprobante g on g.transid = ivk.transid  WHERE ivk.IdReceta=" & mProps.IdReceta
    sql = sql & " Group by g.codtrans,g.numtrans Order by g.codtrans,g.numtrans"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    Do While Not rs.EOF
        s = s & rs!CodTrans & "" & rs!NumTrans & Chr(13)
        rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
        Err.Raise ERR_NOELIMINABLE, "Tiene registro relacionados " & s, MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing
    sql = "DELETE FROM IVRecetaDetalle WHERE IdReceta=" & mProps.IdReceta
    mProps.objEmpresa.Execute sql, True
    
    sql = "DELETE FROM IVReceta WHERE Id=" & mProps.IdReceta
    mProps.objEmpresa.Execute sql, True
  
    

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVRECETA
End Sub

Public Function Clone() As IVReceta
    Dim nuevo As IVReceta
    'Copia las propiedades al nuevo objeto
    Set nuevo = New IVReceta
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

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

Public Property Get IdReceta() As Long
    IdReceta = mProps.IdReceta
End Property

Public Property Get fecha() As Date
    fecha = mProps.fecha
End Property

Public Property Get EsNuevo() As Boolean
    EsNuevo = mbooEsNuevo
End Property

Public Property Get Modificado() As Boolean
    Modificado = mbooModificado
End Property

Public Property Let Modificado(ByVal value As Boolean)
    mbooModificado = value
End Property

Private Sub Class_Initialize()
    'Valores predeterminados
    mbooEsNuevo = True
    mbooModificado = False
    mProps.BandValida = True
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Property Let CodReceta(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "IVDescuento_CodDescuento", MSGERR_INVALIDO
    End If
    ValidaCodigo value      '
    mProps.CodReceta = value
    mbooModificado = True
End Property

Public Property Get CodReceta() As String
    CodReceta = Trim$(mProps.CodReceta)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 254 Then
        Err.Raise ERR_INVALIDO, "IVDescuento_Descripcion", MSGERR_INVALIDO
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

Public Property Let fecha(value As Date)
        mProps.fecha = value
        mbooModificado = True
End Property


Private Sub GrabaRecetaDetalle()
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As IVRecetaDetalle, sql As String
    'Elimina  todos  los registros
    sql = "Delete IVRD " & _
          "From IVReceta IVR " & _
          "INNER JOIN IVRecetaDetalle IVRD ON IVRD.IdReceta = IVRD.IdReceta " & _
          "Where IVRD.IdReceta= " & mProps.IdReceta
    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    If Not mcolIVRecetaDetalle Is Nothing Then
        sql = "SELECT * FROM IVRecetaDetalle WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        For i = 1 To mcolIVRecetaDetalle.Count
            rs1.AddNew
            Set obj = mcolIVRecetaDetalle.item(i)
            rs1!IdReceta = mProps.IdReceta
            rs1!IdInventario = obj.IdInventario
            rs1!Cantidad = obj.Cantidad
'            rs1!Costo = obj.Costo
            rs1!Orden = obj.Orden
            rs1.Update
        Next i
    End If
    Set rs1 = Nothing
    Set obj = Nothing
End Sub

Public Property Get IVRecetaDetalle(ix As Long) As IVRecetaDetalle
    If mcolIVRecetaDetalle Is Nothing Then RecuperaRecetaDetalle
    Set IVRecetaDetalle = mcolIVRecetaDetalle.item(ix)
End Property

Public Property Get NumIVRecetaDetalle() As Long
    If mcolIVRecetaDetalle Is Nothing Then
        NumIVRecetaDetalle = 0
    Else
        NumIVRecetaDetalle = mcolIVRecetaDetalle.Count
    End If
End Property
Public Function AddRecetaDetalle() As Long
    Dim nuevo As IVRecetaDetalle
    If mcolIVRecetaDetalle Is Nothing Then RecuperaRecetaDetalle
    Set nuevo = New IVRecetaDetalle
    Set nuevo.IVReceta = Me     'Me  olvide  del padre
    'Proporciona valores predeterminados
    nuevo.IdReceta = Me.IdReceta
    mcolIVRecetaDetalle.Add nuevo
    AddRecetaDetalle = mcolIVRecetaDetalle.Count   'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function

Private Sub RecuperaRecetaDetalle()
    Dim sql As String, rs As Recordset, i As Long, mod_antes As Boolean
    Dim nuevo As IVRecetaDetalle
    mod_antes = mbooModificado      'Guarda el estado de Modificado
    If mcolIVRecetaDetalle Is Nothing Then Set mcolIVRecetaDetalle = New Collection
    BorrarIVRecetaDetalle
    'Recupera los IVKardex
    sql = "SELECT ivr.*,iv.codinventario,iv.descripcion,ivu.codunidad FROM IVRecetaDetalle ivr "
    sql = sql & " Left Join IVInventario iv Left Join IVUnidad ivu on ivu.idunidad = iv.idunidad on iv.idinventario = ivr.idinventario"
    sql = sql & " Where IdRECETA=" & mProps.IdReceta & " ORDER BY Orden"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            Set nuevo = New IVRecetaDetalle
            mcolIVRecetaDetalle.Add nuevo
            i = mcolIVRecetaDetalle.Count
            Set IVRecetaDetalle(i).IVReceta = Me
            IVRecetaDetalle(i).Recuperar rs
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    mbooModificado = mod_antes
End Sub

Private Sub BorrarIVRecetaDetalle()
    'Borra  toda  la informacion de la coleccion
     If mcolIVRecetaDetalle Is Nothing Then Set mcolIVRecetaDetalle = New Collection
    With mcolIVRecetaDetalle
        Do While .Count > 0
            .Remove .Count
        Loop
    End With
    mbooModificado = True
End Sub
Public Property Let CodInventario(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodInventario Then
        If Len(value) > 0 Then
            'Actualiza IdProveedor también
            sql = "SELECT IdInventario FROM IVinventario WHERE CodInventario='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVInventario_CodInventario", MSGERR_NOHAYCODIGO
                Exit Property
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

Public Property Let Batch(ByVal value As Currency)
    mProps.Batch = value
    mbooModificado = True
End Property

Public Property Get Batch() As Currency
    Batch = mProps.Batch
End Property

Public Sub RemoveRecetaDetalle(ByVal ix As Long, Optional obj As IVRecetaDetalle)
    If mcolIVRecetaDetalle Is Nothing Then RecuperaRecetaDetalle
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolIVRecetaDetalle.Count
            If mcolIVRecetaDetalle.item(ix) Is obj Then
                mcolIVRecetaDetalle.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolIVRecetaDetalle.Remove ix
        End If
    mbooModificado = True
End Sub
