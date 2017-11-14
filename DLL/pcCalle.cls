VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCCalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdCalle As Long
    CodCalle As String
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
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    'Verifica si no repite el código
'    sql = "SELECT CodCalle FROM PCCalle " & _
'          "WHERE (CodCalle='" & mProps.CodCalle & "') AND " & _
'                "(IdCalle<>" & mProps.IDCalle & ")"
'    Set rs = mProps.objEmpresa.OpenRecordset(sql)
'    If Not rs.EOF Then      'Si ya existe el codigo
'        Err.Raise ERR_REPITECODIGO, "PCCalle", MSGERR_REPITECODIGO
'    End If
'    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM PCCalle WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM PCCalle WHERE IdCalle=" & mProps.IdCalle
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "PCCalle_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en Pos
    With rs
        !CodCalle = mProps.CodCalle
        !Descripcion = mProps.Descripcion
        !BandValida = mProps.BandValida
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IdCalle = !IdCalle
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_PCCALL", "Calle Nuevo: " & _
                                                         mProps.CodCalle, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_PCCALL", "Calle Modificado: " & _
                                                         mProps.CodCalle, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PCCALLE
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset, W As Variant
    
    W = Split(key, ",")
    
    If VarType(key) = vbString Then
        sql = "SELECT * "
        sql = sql & " FROM PCCalle "
        'sql = sql & "  WHERE CodCalle='" & W(0) & "'"
        sql = sql & "  WHERE CodCalle='" & key & "'"
    Else
        sql = "SELECT *"
        sql = sql & " FROM PCCalle "
        sql = sql & "  WHERE IdCalle=" & key
    End If

    
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de CantonCopro
    With rs
        mProps.IdCalle = !IdCalle
        mProps.CodCalle = !CodCalle
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
       
'        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset

    'Verifica si existe Kardex del Calle
    sql = "SELECT IdCallePri FROM Pcprovcli WHERE IdCallePri=" & mProps.IdCalle & " or IdCalleSec =" & mProps.IdCalle
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "PCCalle.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing

    'Elimina el Calle
    sql = "DELETE FROM PCCalle WHERE CodCalle='" & mProps.CodCalle & "'"
    mProps.objEmpresa.Execute sql, True

    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_PCCALL", "Calle Eliminado: " & _
                                                         mProps.CodCalle, "DT"
                                                         
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PCCALLE
End Sub

Public Function Clone() As PCCalle
    Dim nuevo As PCCalle
    
    Set nuevo = New PCCalle
    'Copia las propiedades al nuevo objeto
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdCalle = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property


Public Property Let CodCalle(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCCalle_CodCalle", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodCalle = value
    mbooModificado = True
End Property

Public Property Get CodCalle() As String
    CodCalle = Trim$(mProps.CodCalle)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 100 Then
        Err.Raise ERR_INVALIDO, "PCCalle_Descripcion", MSGERR_INVALIDO
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
    With mProps
        .BandValida = True
    End With
    
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Property Get IdCalle() As Long
    IdCalle = mProps.IdCalle
End Property

