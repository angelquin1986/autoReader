VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVComision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***jeaa  30/12/2004
Private Type T_PROPS
    IdComision As Long
    CodComision As String
    Descripcion As String
    BandValida As Boolean
    BandCaducidad As Boolean
    FechaDesde As Date
    FechaHasta As Date
    HoraDesde As Date
    HoraHasta As Date
    fechagrabado As Date
    BandAcumula As Boolean
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mcolComisionDetalle As Collection  '*** jeaa

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Public Sub Grabar()
    Dim sql As String, rs As Recordset
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    
    'Verifica si no repite el código
    sql = "SELECT codComision FROM IVComision "
    sql = sql & " WHERE codComision='" & mProps.CodComision & "'"
    sql = sql & " and idComision<> " & mProps.IdComision

    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "IVComision", MSGERR_REPITECODIGO
    End If
    rs.Close
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM IVComision WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM IVComision WHERE IdComision=" & mProps.IdComision
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)

        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "IVComision_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro
    With rs
        'jeaa 05/10/2010
        !CodComision = mProps.CodComision
        !Descripcion = mProps.Descripcion
        !BandValida = mProps.BandValida
        !BandCaducidad = mProps.BandCaducidad
        !FechaDesde = mProps.FechaDesde
        !FechaHasta = mProps.FechaHasta
        !HoraDesde = mProps.HoraDesde
        !HoraHasta = mProps.HoraHasta
        !BandAcumula = mProps.BandAcumula
        !fechagrabado = Now
        .Update
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IdComision = !IdComision
        .Close
    End With
    Set rs = Nothing
    GrabaComisionDetalle
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_COM", "Descto Nuevo: " & _
                                                         mProps.CodComision, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_COM", "Descto Modificado: " & _
                                                         mProps.CodComision, "DT"
    End If
    
    mbooEsNuevo = False
    mbooModificado = False
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVCOMIXVEN
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVGCOMIXVEN
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String
    Dim rs As Recordset
    Dim v As Variant

    v = Split(key, ",")
        If VarType(key) = vbString Then
            sql = "SELECT *"
            sql = sql & " FROM IVComision "
            sql = sql & " WHERE CodComision='" & key & "'"
        Else
            sql = "SELECT *"
            sql = sql & " FROM IVComision "
            sql = sql & " WHERE IdComision=" & key
        End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'Recupera los datos de Usuario
    With rs
        mProps.IdComision = !IdComision
        If Not IsNull(!CodComision) Then mProps.CodComision = !CodComision
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!BandCaducidad) Then mProps.BandCaducidad = !BandCaducidad
        mProps.FechaDesde = !FechaDesde
        mProps.FechaHasta = !FechaHasta
        mProps.HoraDesde = !HoraDesde
        mProps.HoraHasta = !HoraHasta
        If Not IsNull(!BandAcumula) Then mProps.BandAcumula = !BandAcumula
        mProps.fechagrabado = !fechagrabado
        
        .Close
    End With
    'Recupera  el detalle
    RecuperaDetalle
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
  

    sql = "DELETE FROM IVComisionDetalle WHERE IdComision=" & mProps.IdComision
    mProps.objEmpresa.Execute sql, True
  
  
  
  
    'Elimina el registro
    sql = "DELETE FROM IVComision WHERE IdComision=" & mProps.IdComision
    mProps.objEmpresa.Execute sql, True

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVCOMIXVEN
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVGCOMIXVEN
End Sub

Public Function Clone() As IVComision
    Dim nuevo As IVComision
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New IVComision
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdComision = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Get IdComision() As Long
    IdComision = mProps.IdComision
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
    mProps.BandAcumula = False

End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Property Let CodComision(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "IVComision_CodComision", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodComision = value
    mbooModificado = True
End Property

Public Property Get CodComision() As String
    CodComision = Trim$(mProps.CodComision)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 254 Then
        Err.Raise ERR_INVALIDO, "IVComision_Descripcion", MSGERR_INVALIDO
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



'*********** coleccion de Grupo
Private Sub GrabaComisionDetalle()
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As IVComisionDetalle, sql As String
    'Elimina  todos  los registros
    sql = "Delete IVRD " & _
          "From IVComision IVR " & _
          "INNER JOIN IVComisionDetalle IVRD ON IVR.IdComision = IVRD.IdComision " & _
          "Where IVRD.IdComision = " & mProps.IdComision
    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    If Not mcolComisionDetalle Is Nothing Then
        sql = "SELECT * FROM IVComisionDetalle WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        For i = 1 To mcolComisionDetalle.Count
            rs1.AddNew
            Set obj = mcolComisionDetalle.item(i)
            rs1!IdComision = mProps.IdComision
            rs1!IdInventario = obj.IdInventario
            rs1!IdGrupo = obj.IdGrupo
            rs1!IdVendedor = obj.IdVendedor
            rs1!Comision = obj.Comision
            rs1!Orden = obj.Orden
            rs1.Update
        Next i
    End If
    Set rs1 = Nothing
    Set obj = Nothing
End Sub



Private Sub BorrarDetalle()
    'Borra  toda  la informacion de la coleccion
     If mcolComisionDetalle Is Nothing Then Set mcolComisionDetalle = New Collection
    With mcolComisionDetalle
        Do While .Count > 0
            .Remove .Count
        Loop
    End With
    mbooModificado = True
End Sub



Private Sub RecuperaDetalle()
    Dim sql As String, rs As Recordset, i As Long, mod_antes As Boolean, NumIVGrupo As Integer
    Dim nuevo As IVComisionDetalle
    mod_antes = mbooModificado      'Guarda el estado de Modificado
    If mcolComisionDetalle Is Nothing Then Set mcolComisionDetalle = New Collection
    
        If Len(mProps.objEmpresa.GNOpcion.ObtenerValor("ComisionNumIVGrupo")) > 0 Then
            NumIVGrupo = mProps.objEmpresa.GNOpcion.ObtenerValor("ComisionNumIVGrupo") + 1
        Else
            NumIVGrupo = 1
        End If

    
    
    BorrarDetalle
    
    'Recupera los IVKardex
    sql = "SELECT IVComisionDetalle.* , ivinventario.codinventario, codgrupo" & NumIVGrupo & " as codgrupo FROM IVComisionDetalle "
    sql = sql & " left join Ivinventario on IVComisionDetalle.idinventario= ivinventario.idinventario "
    sql = sql & " left join IvGrupo" & NumIVGrupo & " ivg on IVComisionDetalle.idgrupo= ivg.idgrupo" & NumIVGrupo & " "
    sql = sql & " WHERE IdComision=" & mProps.IdComision & " ORDER BY Orden"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            Set nuevo = New IVComisionDetalle
            mcolComisionDetalle.Add nuevo
            i = mcolComisionDetalle.Count
            Set IVComisionDetalle(i).IVComision = Me
            IVComisionDetalle(i).Recuperar rs
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    mbooModificado = mod_antes      'Para que no cambie el estado de Modificado
End Sub

Public Property Get IVComisionDetalle(ix As Long) As IVComisionDetalle
    If mcolComisionDetalle Is Nothing Then RecuperaDetalle   'posilbe recursividad
    Set IVComisionDetalle = mcolComisionDetalle.item(ix)
End Property

Public Property Get NumComisionDetalle() As Long
    If mcolComisionDetalle Is Nothing Then
        NumComisionDetalle = 0
    Else
        NumComisionDetalle = mcolComisionDetalle.Count
    End If
End Property


'--------------- Manejo de coleccion de ComisionDetallePC
Public Function AddDetalleComision() As Long
    Dim nuevo As IVComisionDetalle
    
    If mcolComisionDetalle Is Nothing Then RecuperaDetalle
    Set nuevo = New IVComisionDetalle
    Set nuevo.IVComision = Me    'Me  olvide  del padre
    'Proporciona valores predeterminados
    nuevo.IdComision = Me.IdComision
    mcolComisionDetalle.Add nuevo
    AddDetalleComision = mcolComisionDetalle.Count   'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function

Public Sub RemoveComisionDetalle(ByVal ix As Long, Optional obj As IVComisionDetalle)
    If mcolComisionDetalle Is Nothing Then RecuperaDetalle
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolComisionDetalle.Count
            If mcolComisionDetalle.item(ix) Is obj Then
                mcolComisionDetalle.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolComisionDetalle.Remove ix
    End If
    
    mbooModificado = True
End Sub

Public Property Let BandAcumula(ByVal value As Boolean)
    mProps.BandAcumula = value
    mbooModificado = True
End Property

Public Property Get BandAcumula() As Boolean
    BandAcumula = mProps.BandAcumula
End Property



