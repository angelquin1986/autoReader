VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCCanton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Type T_PROPS
    IdCanton As Long
    CodCanton As String
    Descripcion As String
    IdProvincia As Long
    CodProvincia As String
    bandValida As Boolean
    CodCantonSC As String
    
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
    sql = "SELECT CodCanton FROM PCCanton " & _
          "WHERE (CodCanton='" & mProps.CodCanton & "') AND " & _
                "(IdCanton<>" & mProps.IdCanton & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "PCCanton", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM PCCanton WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM PCCanton WHERE IdCanton=" & mProps.IdCanton
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "PCCanton_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en Pos
    With rs
        !CodCanton = mProps.CodCanton
        !Descripcion = mProps.Descripcion
        !IdProvincia = mProps.IdProvincia
        !bandValida = mProps.bandValida
        !CodCantonSC = mProps.CodCantonSC
        
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.FechaGrabado = !FechaGrabado
        mProps.IdCanton = !IdCanton
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_PCCANT", "Canton Nuevo: " & _
                                                         mProps.CodCanton, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_PCCANT", "Canton Modificado: " & _
                                                         mProps.CodCanton, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PCCANTON
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT PCC.*, pcp.codprovincia "
        sql = sql & " FROM PCCanton pcc LEFT JOIN pcprovincia pcp on pcc.idprovincia = pcp.idprovincia"
        sql = sql & "  WHERE CodCanton='" & key & "'"
    Else
        sql = "SELECT PCC.*, pcp.codprovincia "
        sql = sql & " FROM PCCanton pcc LEFT JOIN pcprovincia pcp on pcc.idprovincia = pcp.idprovincia"
        sql = sql & "  WHERE IdCanton=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
     If rs.RecordCount = 0 Then 'AUC si no existe por codigo busca a ver si emcuenta por descripcion '
        Set rs = Nothing
        sql = "SELECT PCC.*, pcp.codprovincia " 'para importar roles
        sql = sql & " FROM PCCanton pcc LEFT JOIN pcprovincia pcp on pcc.idprovincia = pcp.idprovincia"
        sql = sql & "  WHERE pcc.Descripcion='" & key & "'"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
    End If

    'Recupera los datos de ProvinciaCopro
    With rs
        mProps.IdCanton = !IdCanton
        mProps.CodCanton = !CodCanton
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!IdProvincia) Then mProps.IdProvincia = !IdProvincia
        If Not IsNull(!CodProvincia) Then mProps.CodProvincia = !CodProvincia
        If Not IsNull(!bandValida) Then mProps.bandValida = !bandValida
        If Not IsNull(!CodCantonSC) Then mProps.CodCantonSC = !CodCantonSC
       
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset

    'Verifica si existe cliente
    sql = "SELECT IdCanton FROM Pcprovcli WHERE IdCanton=" & mProps.IdCanton
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "PCCanton.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing


    'Verifica si existe Parroquia del Canton
    sql = "SELECT IdCanton FROM PCParroquia WHERE IdCanton=" & mProps.IdCanton
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "PCCanton.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing

    'Elimina el Canton
    sql = "DELETE FROM PCCanton WHERE CodCanton='" & mProps.CodCanton & "'"
    mProps.objEmpresa.Execute sql, True

    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_PCCANT", "Canton Eliminado: " & _
                                                         mProps.CodCanton, "DT"
                                                         
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PCCANTON
End Sub

Public Function Clone() As PCCanton
    Dim nuevo As PCCanton
    
    Set nuevo = New PCCanton
    'Copia las propiedades al nuevo objeto
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdCanton = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property


Public Property Let CodCanton(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCCanton_CodCanton", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodCanton = value
    mbooModificado = True
End Property

Public Property Get CodCanton() As String
    CodCanton = Trim$(mProps.CodCanton)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "PCCanton_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let IdProvincia(ByVal value As Long)
    mProps.IdProvincia = value
    mbooModificado = True
End Property

Public Property Get IdProvincia() As Long
    IdProvincia = mProps.IdProvincia
End Property


Public Property Let CodProvincia(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodProvincia Then
        'Actualiza IdProvincia también
        sql = "SELECT IdProvincia FROM Pcprovincia WHERE CodProvincia='" & value & "'"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "PCCanton_CodProvincia", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            IdProvincia = rs!IdProvincia
        End If
        rs.Close
    End If

    Set rs = Nothing
    mProps.CodProvincia = value
    mbooModificado = True
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "PCCanton_CodProvincia", MSGERR_INVALIDO
    End If
    
    mProps.CodProvincia = value
    mbooModificado = True

End Property

Public Property Get CodProvincia() As String
    CodProvincia = Trim$(mProps.CodProvincia)
End Property


Public Property Let bandValida(ByVal value As Boolean)
    mProps.bandValida = value
    mbooModificado = True
End Property

Public Property Get bandValida() As Boolean
    bandValida = mProps.bandValida
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
    With mProps
        .bandValida = True
    End With
    
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Property Let IdCanton(ByVal value As Long)
    mProps.IdCanton = value
    mbooModificado = True
End Property
Public Property Get IdCanton() As Long
    IdCanton = mProps.IdCanton
End Property


Public Property Let CodCantonSC(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCCanton_CodCantonSC", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodCantonSC = value
    mbooModificado = True
End Property

Public Property Get CodCantonSC() As String
    CodCantonSC = Trim$(mProps.CodCantonSC)
End Property

