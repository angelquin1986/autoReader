VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVTipoCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'jeaa 23/12/2005
Option Explicit
Private Type T_PROPS
    IdTipoCompra As Long
    CodTipoCompra As String
    Descripcion As String
    BandValida As Boolean
        
    FechaGrabado As Date
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean


Public Sub Grabar()
    Dim sql As String, rs As Recordset
    
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
    sql = "SELECT CodTipoCompra FROM IVTipoCompra " & _
          " WHERE (CodTipoCompra)='" & mProps.CodTipoCompra & "' AND " & _
                 "(IdTipoCompra <>" & mProps.IdTipoCompra & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "IVTipoCompra_Grabar", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM IVTipoCompra WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM IVTipoCompra WHERE IdTipoCompra =" & mProps.IdTipoCompra
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "IVTipoCompra_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en Pos
    With rs
        !CodTipoCompra = mProps.CodTipoCompra
        !Descripcion = mProps.Descripcion
        !BandValida = mProps.BandValida
        
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.FechaGrabado = !FechaGrabado
        mProps.IdTipoCompra = !IdTipoCompra
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_TC", "TipoCompra Nuevo: " & _
                                                         mProps.CodTipoCompra, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_TC", "TipoCompra Modificado: " & _
                                                         mProps.CodTipoCompra, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_TIPOCOMPRA
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM IVTipoCompra WHERE CodTipoCompra='" & key & "'"
    Else
        sql = "SELECT * FROM IVTipoCompra WHERE IdTipoCompra=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de TipoCompra
    With rs
        mProps.IdTipoCompra = !IdTipoCompra
        mProps.CodTipoCompra = !CodTipoCompra
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String
    Dim rs As Recordset
    
    'Verifica si ya existe un item usando el mismo CodTipoCompra
'    sql = "SELECT CodTrans FROM GnComprobante" & _
'          " WHERE IdTipoCompra=" & mProps.IdTipoCompra
'    Set rs = mProps.objEmpresa.OpenRecordset(sql)
'    If Not rs.EOF Then
'        Err.Raise ERR_NOELIMINABLE, "TipoCompra_Eliminar", MSGERR_NOELIMINABLE
'    End If
'    rs.Close

    'Elimina el registro
    sql = "DELETE FROM IVTipoCompra WHERE IdTipoCompra=" & mProps.IdTipoCompra
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_TC", "TipoCompra Eliminado: " & _
                                                         mProps.CodTipoCompra, "DT"
                                                         
    Set rs = Nothing
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_TIPOCOMPRA
End Sub

Friend Function Clone() As IVTipoCompra
    Dim nuevo As IVTipoCompra
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New IVTipoCompra
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdTipoCompra = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Let CodTipoCompra(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "TipoCompra_CodTipoCompra", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodTipoCompra = value
    mbooModificado = True
End Property

Public Property Get CodTipoCompra() As String
    CodTipoCompra = Trim$(mProps.CodTipoCompra)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "TipoCompra_Descripcion", MSGERR_INVALIDO
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




Public Property Get FechaGrabado() As Date
    FechaGrabado = mProps.FechaGrabado
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




    


