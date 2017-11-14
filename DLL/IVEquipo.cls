VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVEquipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Type T_PROPS
    IdEquipo As Long
    CodEquipo As String
    Descripcion As String
    IDForma As Long
    CodForma As String
    BandValida As Boolean
    IdCliente As Long
    CodCliente As String
    
    fechagrabado As Date
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
    sql = "SELECT CodEquipo FROM IVEquipo " & _
          "WHERE (CodEquipo='" & mProps.CodEquipo & "') AND " & _
                "(IdEquipo<>" & mProps.IdEquipo & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "IVEquipo", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM IVEquipo WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM IVEquipo WHERE IdEquipo=" & mProps.IdEquipo
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "IVEquipo_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en Pos
    With rs
        !CodEquipo = mProps.CodEquipo
        !Descripcion = mProps.Descripcion
        '!IdForma = mProps.IdForma
'        !CodForma = mProps.CodForma
        !BandValida = mProps.BandValida
        !IdCliente = mProps.IdCliente
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IdEquipo = !IdEquipo
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_EQUI", "Equipo Nuevo: " & _
                                                         mProps.CodEquipo, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_EQUI", "Equipo Mod: " & _
                                                         mProps.CodEquipo, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVEQUIPO
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT IVEquipo.* , CodProvcli as CodCliente "
        sql = sql & " FROM IVEquipo "
        sql = sql & " LEFT JOIN PCProvcli "
        sql = sql & "  ON IVEquipo.IDcliente = Pcprovcli.IDProvcli"
        sql = sql & "  WHERE CodEquipo='" & key & "'"
    Else
        sql = "SELECT IVEquipo.* , CodProvcli as CodCliente "
        sql = sql & " FROM IVEquipo "
        sql = sql & " LEFT JOIN PCProvcli "
        sql = sql & "  ON IVEquipo.IDcliente = Pcprovcli.IDProvcli"
        sql = sql & "  WHERE IdEquipo=" & key
        
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de FormaCopro
    With rs
        mProps.IdEquipo = !IdEquipo
        mProps.CodEquipo = !CodEquipo
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
'        If Not IsNull(!CodForma) Then mProps.CodForma = !CodForma
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!IdCliente) Then mProps.IdCliente = !IdCliente
        If Not IsNull(!CodCliente) Then mProps.CodCliente = !CodCliente
       
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset

    'Verifica si existe Kardex del Equipo
    sql = "SELECT IdEquipo FROM PCKardex WHERE IdEquipo=" & mProps.IdEquipo
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "IVEquipo.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing

    'Elimina el Equipo
    sql = "DELETE FROM IVEquipo WHERE CodEquipo='" & mProps.CodEquipo & "'"
    mProps.objEmpresa.Execute sql, True

    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_EQUI", "Equipo Eliminado: " & _
                                                         mProps.CodEquipo, "DT"
                                                         
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVEQUIPO
End Sub

Public Function Clone() As IVEquipo
    Dim nuevo As IVEquipo
    
    Set nuevo = New IVEquipo
    'Copia las propiedades al nuevo objeto
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdEquipo = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property


Public Property Let CodEquipo(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "IVEquipo_CodEquipo", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodEquipo = value
    mbooModificado = True
End Property

Public Property Get CodEquipo() As String
    CodEquipo = Trim$(mProps.CodEquipo)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "IVEquipo_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let IDForma(ByVal value As Long)
    mProps.IDForma = value
    mbooModificado = True
End Property

Public Property Get IDForma() As Long
    IDForma = mProps.IDForma
End Property


Public Property Let CodForma(ByVal value As String)
''    Dim sql As String, rs As Recordset
''
''    'Cuando cambia el valor
''    If value <> mProps.CodForma Then
''        'Actualiza IdForma también
''        sql = "SELECT IdForma FROM TsFormaCobroPago WHERE CodForma='" & value & "'"
''        Set rs = mProps.objEmpresa.OpenRecordset(sql)
''        If rs.EOF Then
''            Err.Raise ERR_NOHAYCODIGO, "IVEquipo_CodForma", MSGERR_NOHAYCODIGO
''            Exit Property
''        Else
''            IdForma = rs!IdForma
''        End If
''        rs.Close
''    End If
''
''    Set rs = Nothing
''    mProps.CodForma = value
''    mbooModificado = True
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "IVEquipo_Codforma", MSGERR_INVALIDO
    End If
    
    mProps.CodForma = value
    mbooModificado = True

End Property

Public Property Get CodForma() As String
    CodForma = Trim$(mProps.CodForma)
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
    With mProps
        .BandValida = True
    End With
    
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Property Let IdCliente(ByVal value As Long)
    mProps.IdCliente = value
    mbooModificado = True
End Property

Public Property Get IdCliente() As Long
    IdCliente = mProps.IdCliente
End Property


Public Property Let CodCliente(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCliente Then
        'Actualiza IdCliente también
        sql = "SELECT IdProvCli FROM Pcprovcli WHERE CodProvCli='" & value & "'"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "IVBanco_CodCliente", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            IdCliente = rs!IdProvCli
        End If
        rs.Close
    End If

    Set rs = Nothing
    mProps.CodCliente = value
    mbooModificado = True
End Property

Public Property Get CodCliente() As String
    CodCliente = Trim$(mProps.CodCliente)
End Property




