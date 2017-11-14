VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Empresa"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type T_PROPS
    NumGrupo As Byte            '1,2,3
    IdGrupo As Long
    CodGrupo As String
    Descripcion As String
    BandValida As Boolean
    BandProduccion As Boolean
    sItems As String
    iditemprod As Long
    Coditemprod As String
    fechagrabado As Date
    Rebate As Currency
    IdealDias As Currency
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
    sql = "SELECT " & tabla("Cod") & " FROM " & tabla("IV") & _
          " WHERE (" & tabla("Cod") & "='" & mProps.CodGrupo & "') AND " & _
                 "(" & tabla("Id") & "<>" & mProps.IdGrupo & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "IVGrupo_Grabar", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM " & tabla("IV") & " WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM " & tabla("IV") & " WHERE " & tabla("Id") & "=" & mProps.IdGrupo
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "IVGrupo_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en Pos
    With rs
        .Fields(tabla("Cod")) = mProps.CodGrupo
        !Descripcion = mProps.Descripcion
        !BandValida = mProps.BandValida
        !BandProduccion = mProps.BandProduccion
        !iditemprod = mProps.iditemprod
        !Rebate = mProps.Rebate
        !IdealDias = mProps.IdealDias
        '!sItems = mProps.sItems
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IdGrupo = .Fields(tabla("Id"))
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
        mProps.objEmpresa.GrabaGNLogAccion "NUE_GRU", "Grupo Nuevo: " & _
                                                         mProps.CodGrupo, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_GRU", "Grupo Modificado: " & _
                                                         mProps.CodGrupo, "DT"
    End If
    
    
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    If tabla("Cod") = "CodGrupo6" Then
        mProps.objEmpresa.RaiseRefreshEvent REFRESH_GRUPO6
    Else
        mProps.objEmpresa.RaiseRefreshEvent REFRESH_GRUPO1 + NumGrupo - 1
    End If
End Sub

Private Function tabla(prefix As String) As String
    tabla = prefix & "Grupo" & NumGrupo
End Function

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT ivg.*,iv.codinventario FROM " & tabla("IV") & " ivg Left Join Ivinventario iv on iv.idinventario= ivg.iditemprod WHERE ivg." & tabla("Cod") & "='" & key & "'"
    Else
        sql = "SELECT ivg.*,iv.codinventario FROM " & tabla("IV") & " ivg Left Join Ivinventario iv on iv.idinventario= ivg.iditemprod WHERE ivg." & tabla("Id") & "=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de grupo
    With rs
        mProps.IdGrupo = .Fields(tabla("Id"))
        mProps.CodGrupo = .Fields(tabla("Cod"))
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!BandProduccion) Then mProps.BandProduccion = !BandProduccion
        'If Not IsNull(!sItems) Then mProps.sItems = !sItems
        If Not IsNull(!iditemprod) Then mProps.iditemprod = !iditemprod
        If Not IsNull(!CodInventario) Then mProps.Coditemprod = !CodInventario
        If Not IsNull(!Rebate) Then mProps.Rebate = !Rebate
        If Not IsNull(!IdealDias) Then mProps.IdealDias = !IdealDias
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
    
    'Verifica si ya existe un item usando el mismo CodGrupo
    sql = "SELECT CodInventario FROM IVInventario" & _
          " WHERE " & tabla("Id") & "=" & mProps.IdGrupo
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "IVGrupo_Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close

    'Elimina el registro
    sql = "DELETE FROM " & tabla("IV") & " WHERE " & tabla("Id") & "=" & mProps.IdGrupo
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_GRU", "Grupo Eliminado: " & _
                                                         mProps.CodGrupo, "DT"
                                                         
    Set rs = Nothing
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_GRUPO1 + NumGrupo - 1
End Sub

Friend Function Clone() As ivgrupo
    Dim nuevo As ivgrupo
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New ivgrupo
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdGrupo = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property




Friend Property Let NumGrupo(ByVal value As Byte)
    mProps.NumGrupo = value
    mbooModificado = True
End Property

Public Property Get NumGrupo() As Byte
    NumGrupo = mProps.NumGrupo
End Property


Public Property Let CodGrupo(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "IVGrupo_CodGrupo", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodGrupo = value
    mbooModificado = True
End Property

Public Property Get CodGrupo() As String
    CodGrupo = Trim$(mProps.CodGrupo)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "IVGrupo_Descripcion", MSGERR_INVALIDO
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


    'Verifica si no repite el código
    If mbooEsNuevo Then
        sql = "SELECT " & tabla("Cod") & " FROM " & bd & ".dbo." & tabla("IV") & _
              " WHERE (" & tabla("Cod") & "='" & mProps.CodGrupo & "')  " & _
                     ""
        Set rs = mProps.objEmpresa.OpenRecordsetSUC(sql, bd, servidor)
       If Not rs.EOF Then      'Si ya existe el codigo
            Err.Raise ERR_REPITECODIGO, "IVGrupo_Grabar", MSGERR_REPITECODIGO
        End If
        rs.Close
    End If
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM " & bd & ".dbo." & tabla("IV") & " WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEditSuc(sql, bd, servidor)
       rs.AddNew
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM " & bd & ".dbo." & tabla("IV") & " WHERE " & tabla("Id") & "=" & mProps.IdGrupo
        Set rs = mProps.objEmpresa.OpenRecordsetParaEditSuc(sql, bd, servidor)
       If rs.RecordCount = "0" Then
            sql = "SELECT * FROM " & bd & ".dbo." & tabla("IV") & " WHERE 1=0"
            Set rs = mProps.objEmpresa.OpenRecordsetParaEditSuc(sql, bd, servidor)
           rs.AddNew
        End If
#If DAOLIB Then
        rs.Edit
#End If

    End If
    'Graba el registro en Pos
    With rs
        .Fields(tabla("Cod")) = mProps.CodGrupo
        !Descripcion = mProps.Descripcion
        !BandValida = mProps.BandValida
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IdGrupo = .Fields(tabla("Id"))
        .Close
    End With
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccionEnSucursales bd, servidor, "NUE_GRU", "Grupo Nuevo: " & _
                                                        mProps.CodGrupo, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccionEnSucursales bd, servidor, "MOD_GRU", "Grupo Modificado: " & _
                                                        mProps.CodGrupo, "DT"
    End If
    Set rs = Nothing
End Sub

Public Property Get IdGrupo() As String
    IdGrupo = Trim$(mProps.IdGrupo)
End Property

Public Property Let BandProduccion(ByVal value As Boolean)
    mProps.BandProduccion = value
    mbooModificado = True
End Property

Public Property Get BandProduccion() As Boolean
    BandProduccion = mProps.BandProduccion
End Property

Public Property Let sItems(ByVal value As String)
'    mProps.sItems = value
 '   mbooModificado = True
End Property

Public Property Get sItems() As String
  '  sItems = mProps.sItems
End Property

Public Property Let Coditemprod(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.Coditemprod Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaActivo también
            sql = "SELECT IdInventario FROM IVInventario WHERE CodInventario='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVGrupo_CodItemProd", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                iditemprod = rs!IdInventario
            End If
            rs.Close
        Else
            iditemprod = 0
        End If
        
        mProps.Coditemprod = value
        mbooModificado = True
    End If

    Set rs = Nothing
End Property

Public Property Get Coditemprod() As String
    Coditemprod = Trim$(mProps.Coditemprod)
End Property

Public Property Let iditemprod(ByVal value As String)
    mProps.iditemprod = value
    mbooModificado = True
End Property

Public Property Get iditemprod() As String
    iditemprod = mProps.iditemprod
End Property

Public Property Let Rebate(ByVal value As Currency)
    mProps.Rebate = value
    mbooModificado = True
End Property

Public Property Get Rebate() As Currency
    Rebate = mProps.Rebate
End Property

Public Property Let IdealDias(ByVal value As Currency)
    mProps.IdealDias = value
    mbooModificado = True
End Property

Public Property Get IdealDias() As Currency
    IdealDias = mProps.IdealDias
End Property

