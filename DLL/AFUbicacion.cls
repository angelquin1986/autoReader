VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AFUbicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit




Private Type T_Props
    IdUbicacion As Long
    CodUbicacion As String
    Descripcion As String
    bandValida As Boolean
        
    FechaGrabado As Date
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mProps As T_Props
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean


Public Sub grabar()
    Dim sql As String, rs As Recordset
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
    sql = "SELECT CodUbicacion FROM AFUbicacion " & _
          "WHERE (CodUbicacion='" & mProps.CodUbicacion & "') AND " & _
                "(IdUbicacion<>" & mProps.IdUbicacion & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "AFUbicacion", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM AFUbicacion WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM AFUbicacion WHERE IdUbicacion=" & mProps.IdUbicacion
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "AFUbicacion_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en Pos
    With rs
        !CodUbicacion = mProps.CodUbicacion
        !Descripcion = mProps.Descripcion
        !bandValida = mProps.bandValida
        
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.FechaGrabado = !FechaGrabado
        mProps.IdUbicacion = !IdUbicacion
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_BOD", "Ubicacion Nueva: " & _
                                                         mProps.CodUbicacion, "AF"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_BOD", "Ubicacion Modificada: " & _
                                                         mProps.CodUbicacion, "AF"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_AFUBICACION
End Sub

Friend Sub recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM AFUbicacion WHERE CodUbicacion='" & key & "'"
    Else
        sql = "SELECT * FROM AFUbicacion WHERE IdUbicacion=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IdUbicacion = !IdUbicacion
        mProps.CodUbicacion = !CodUbicacion
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!bandValida) Then mProps.bandValida = !bandValida
        
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset

    'Verifica si existe items en esta Ubicacion
    sql = "SELECT IdUbicacion FROM AFInventario WHERE IdUbicacion=" & mProps.IdUbicacion
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "AFUbicacion.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing

    'Elimina la empresa
    sql = "DELETE FROM AFUbicacion WHERE IdUbicacion=" & mProps.IdUbicacion
    mProps.objEmpresa.Execute sql, True

    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_BOD", "Ubicacion Eliminada: " & _
                                                         mProps.CodUbicacion, "AF"
                                                         
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_AFUBICACION
End Sub

Friend Function Clone() As AFUbicacion
    Dim nuevo As AFUbicacion
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New AFUbicacion
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(Value As T_Props)         'Para el método Clone
    mProps = Value
    mProps.IdUbicacion = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property


Public Property Let CodUbicacion(ByVal Value As String)
    If Len(Value) > 10 Then
        Err.Raise ERR_INVALIDO, "AFUbicacion_CodUbicacion", MSGERR_INVALIDO
    End If
    ValidaCodigo Value      'Valida si contiene caracteres no validos
    
    mProps.CodUbicacion = Value
    mbooModificado = True
End Property

Public Property Get CodUbicacion() As String
    CodUbicacion = Trim$(mProps.CodUbicacion)
End Property

Public Property Let Descripcion(ByVal Value As String)
    If Len(Value) > 80 Then
        Err.Raise ERR_INVALIDO, "AFUbicacion_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = Value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let bandValida(ByVal Value As Boolean)
    mProps.bandValida = Value
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
    mbooEsNuevo = True
    mbooModificado = False
    
    'Valores iniciales
    mProps.bandValida = True
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

