VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AFGrupo"
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
    sql = "SELECT " & tabla("Cod") & " FROM " & tabla("AF") & _
          " WHERE (" & tabla("Cod") & "='" & mProps.CodGrupo & "') AND " & _
                 "(" & tabla("Id") & "<>" & mProps.IdGrupo & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "AFGrupo_Grabar", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM " & tabla("AF") & " WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM " & tabla("AF") & " WHERE " & tabla("Id") & "=" & mProps.IdGrupo
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "AFGrupo_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en Pos
    With rs
        .Fields(tabla("Cod")) = mProps.CodGrupo
        !Descripcion = mProps.Descripcion
        !BandValida = mProps.BandValida
        
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.FechaGrabado = !FechaGrabado
        mProps.IdGrupo = .Fields(tabla("Id"))
        .Close
    End With
    
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
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_AFGRUPO1 + NumGrupo - 1
End Sub

Private Function tabla(prefix As String) As String
    tabla = prefix & "Grupo" & NumGrupo
End Function

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM " & tabla("AF") & " WHERE " & tabla("Cod") & "='" & key & "'"
    Else
        sql = "SELECT * FROM " & tabla("AF") & " WHERE " & tabla("Id") & "=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de grupo
    With rs
        mProps.IdGrupo = .Fields(tabla("Id"))
        mProps.CodGrupo = .Fields(tabla("Cod"))
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
    
    'Verifica si ya existe un item usando el mismo CodGrupo
    sql = "SELECT CodInventario FROM AFInventario" & _
          " WHERE " & tabla("Id") & "=" & mProps.IdGrupo
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "AFGrupo_Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close

    'Elimina el registro
    sql = "DELETE FROM " & tabla("AF") & " WHERE " & tabla("Id") & "=" & mProps.IdGrupo
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_GRU", "Grupo Eliminado: " & _
                                                         mProps.CodGrupo, "DT"
                                                         
    Set rs = Nothing
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_AFGRUPO1 + NumGrupo - 1
End Sub

Friend Function Clone() As AFGrupo
    Dim nuevo As AFGrupo
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New AFGrupo
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
        Err.Raise ERR_INVALIDO, "AFGrupo_CodGrupo", MSGERR_INVALIDO
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
        Err.Raise ERR_INVALIDO, "AFGrupo_Descripcion", MSGERR_INVALIDO
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



