VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVUnidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'AUC 02/12/2005
Option Explicit
Private Type T_PROPS
    IdUnidad As Long
    CodUnidad As String
    Descripcion As String
    BandValida As Boolean
        
    fechagrabado As Date
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean


Public Sub Grabar()
    Dim sql As String, rs As Recordset
    Dim rsSuc As Recordset, gnSuc As GNSucursal, i As Integer
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
    sql = "SELECT CodUnidad FROM IVUnidad " & _
          " WHERE (CodUnidad)='" & mProps.CodUnidad & "' AND " & _
                 "(IdUnidad <>" & mProps.IdUnidad & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "IVUnidad_Grabar", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM IVUnidad WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM IVUnidad WHERE IdUnidad =" & mProps.IdUnidad
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "IVUnidad_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en Pos
    With rs
        !CodUnidad = mProps.CodUnidad
        !Descripcion = mProps.Descripcion
        !BandValida = mProps.BandValida
        
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IdUnidad = !IdUnidad
        .Close
    End With
    
    'actualiza Inventario en Bases de Datos Sucursales
    sql = "SELECT * FROM GnSucursal WHERE bandvalida=1"
    Set rsSuc = mProps.objEmpresa.OpenRecordset(sql)
    
    For i = 1 To rsSuc.RecordCount
        Set gnSuc = mProps.objEmpresa.RecuperaGNSucursal(rsSuc.Fields("CodSucursal"))
        If gnSuc.BandInventario Then
            GrabarEnSucursales gnSuc.BaseDatos, gnSuc.servidor
        End If
        rsSuc.MoveNext
    Next i
    Set gnSuc = Nothing
    Set rsSuc = Nothing
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_UNI", "Unidad Nueva: " & _
                                                         mProps.CodUnidad, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_UNI", "Unidad Modificado: " & _
                                                         mProps.CodUnidad, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVUNIDAD
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM IVUnidad WHERE CodUnidad='" & key & "'"
    Else
        sql = "SELECT * FROM IVUnidad WHERE IdUnidad=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Unidad
    With rs
        mProps.IdUnidad = !IdUnidad
        mProps.CodUnidad = !CodUnidad
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String
    Dim rs As Recordset
    
    'Verifica si ya existe un item usando el mismo CodUnidad
    sql = "SELECT CodInventario FROM IvInventario" & _
          " WHERE IdUnidad=" & mProps.IdUnidad & " or idUnidadConteo=" & mProps.IdUnidad
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "Unidad_Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close

    'Elimina el registro
    sql = "DELETE FROM IVUnidad WHERE IdUnidad=" & mProps.IdUnidad
    
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_UNIDAD", "Unidad Eliminado: " & _
                                                         mProps.CodUnidad, "DT"
                                                         
    Set rs = Nothing
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVUNIDAD
End Sub

Friend Function Clone() As IVUnidad
    Dim nuevo As IVUnidad
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New IVUnidad
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdUnidad = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Let CodUnidad(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "IVUnidad_CodUnidad", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodUnidad = value
    mbooModificado = True
End Property

Public Property Get CodUnidad() As String
    CodUnidad = Trim$(mProps.CodUnidad)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "Unidad_Descripcion", MSGERR_INVALIDO
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
    mbooEsNuevo = True
    mbooModificado = False
    
    'Valores iniciales
    mProps.BandValida = True
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Sub GrabarEnSucursal(ByVal bd As String)
End Sub

Public Sub GrabarEnSucursales(ByVal bd As String, ByVal servidor As String)
   Dim sql As String, rs As Recordset
    
    sql = "SELECT CodUnidad FROM " & bd & ".dbo.IVUnidad " & _
          " WHERE (CodUnidad)='" & mProps.CodUnidad & "' "
    Set rs = mProps.objEmpresa.OpenRecordsetSUC(sql, bd, servidor)
   If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "IVUnidad_Grabar", MSGERR_REPITECODIGO
    End If
    rs.Close
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM " & bd & ".dbo.IVUnidad WHERE 1=0"

        Set rs = mProps.objEmpresa.OpenRecordsetParaEditSuc(sql, bd, servidor)
        rs.AddNew
   'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM " & bd & ".dbo.IVUnidad WHERE IdUnidad =" & mProps.IdUnidad
        Set rs = mProps.objEmpresa.OpenRecordsetParaEditSuc(sql, bd, servidor)
#If DAOLIB Then
        rs.Edit
#End If
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
    End If
    'Graba el registro en Pos
    With rs
        !CodUnidad = mProps.CodUnidad
        !Descripcion = mProps.Descripcion
        !BandValida = mProps.BandValida
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IdUnidad = !IdUnidad
        .Close
    End With
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccionEnSucursales bd, servidor, "NUE_UNI", "Unidad Nueva: " & _
                                                        mProps.CodUnidad, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccionEnSucursales bd, servidor, "MOD_UNI", "Unidad Modificado: " & _
                                                        mProps.CodUnidad, "DT"
    End If
    Set rs = Nothing
End Sub
    
