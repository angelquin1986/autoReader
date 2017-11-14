VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCParroquia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IDParroquia As Long
    CodParroquia As String
    Descripcion As String
    IdCanton As Long
    CodCanton As String
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
'    sql = "SELECT CodParroquia FROM PCParroquia " & _
'          "WHERE (CodParroquia='" & mProps.CodParroquia & "') AND " & _
'                "(IdParroquia<>" & mProps.IDParroquia & ")"
'    Set rs = mProps.objEmpresa.OpenRecordset(sql)
'    If Not rs.EOF Then      'Si ya existe el codigo
'        Err.Raise ERR_REPITECODIGO, "PCParroquia", MSGERR_REPITECODIGO
'    End If
'    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM PCParroquia WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM PCParroquia WHERE IdParroquia=" & mProps.IDParroquia
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "PCParroquia_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en Pos
    With rs
        !CodParroquia = mProps.CodParroquia
        !Descripcion = mProps.Descripcion
        !IdCanton = mProps.IdCanton
        '!CodCanton = mProps.CodCanton
        !BandValida = mProps.BandValida
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.FechaGrabado = !FechaGrabado
        mProps.IDParroquia = !IDParroquia
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_PCPARR", "Parroquia Nuevo: " & _
                                                         mProps.CodParroquia, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_PCPARR", "Parroquia Modificado: " & _
                                                         mProps.CodParroquia, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PCPARROQUIA
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset, W As Variant
    
    W = Split(key, ",")
    
    If VarType(key) = vbString Then
        sql = "SELECT PCP.*, pcc.codCanton "
        sql = sql & " FROM PCParroquia pcp LEFT JOIN PCCanton pcc on pcc.idcanton = pcp.idcanton"
        sql = sql & " LEFT JOIN PCPROVINCIA prov on pcc.idprovincia = prov.idprovincia"
        sql = sql & "  WHERE CodParroquia='" & W(0) & "' AND CODCANTONSC='" & W(1) & "' AND CODPROVINCIASC='" & W(2) & "'"
    Else
        sql = "SELECT PCP.*, pcc.codCanton "
        sql = sql & " FROM PCParroquia pcp LEFT JOIN PCCanton pcc on pcc.idcanton = pcp.idcanton"
        sql = sql & "  WHERE IdParroquia=" & key
    End If

    
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de CantonCopro
    With rs
        mProps.IDParroquia = !IDParroquia
        mProps.CodParroquia = !CodParroquia
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!IdCanton) Then mProps.IdCanton = !IdCanton
        If Not IsNull(!CodCanton) Then mProps.CodCanton = !CodCanton
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
       
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset

    'Verifica si existe Kardex del Parroquia
    sql = "SELECT IdParroquia FROM Pcprovcli WHERE IdParroquia=" & mProps.IDParroquia
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "PCParroquia.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing

    'Elimina el Parroquia
    sql = "DELETE FROM PCParroquia WHERE CodParroquia='" & mProps.CodParroquia & "'"
    mProps.objEmpresa.Execute sql, True

    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_PCPARR", "Parroquia Eliminado: " & _
                                                         mProps.CodParroquia, "DT"
                                                         
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PCPARROQUIA
End Sub

Public Function Clone() As PCParroquia
    Dim nuevo As PCParroquia
    
    Set nuevo = New PCParroquia
    'Copia las propiedades al nuevo objeto
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IDParroquia = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property


Public Property Let CodParroquia(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCParroquia_CodParroquia", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodParroquia = value
    mbooModificado = True
End Property

Public Property Get CodParroquia() As String
    CodParroquia = Trim$(mProps.CodParroquia)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "PCParroquia_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let IdCanton(ByVal value As Long)
    mProps.IdCanton = value
    mbooModificado = True
End Property

Public Property Get IdCanton() As Long
    IdCanton = mProps.IdCanton
End Property


Public Property Let CodCanton(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodCanton Then
        'Actualiza IdCanton también
        sql = "SELECT IdCanton FROM PCCanton WHERE CodCanton='" & value & "'"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "PCParroquia_CodCanton", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            IdCanton = rs!IdCanton
        End If
        rs.Close
    
    End If

    Set rs = Nothing
    mProps.CodCanton = value
    mbooModificado = True
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "PCParroquia_CodCanton", MSGERR_INVALIDO
    End If
    
    mProps.CodCanton = value
    
    mbooModificado = True

End Property

Public Property Get CodCanton() As String
    CodCanton = Trim$(mProps.CodCanton)
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
    With mProps
        .BandValida = True
    End With
    
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub





