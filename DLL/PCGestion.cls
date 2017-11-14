VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCGestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    Id As Long
    CodGestion As String
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
'    sql = "SELECT idpcggasto,idpcgrupo FROM pcggasto"
'    sql = sql & " WHERE idpcggasto<> " & mProps.IdPCGGasto
'    Set rs = mProps.objEmpresa.OpenRecordset(sql)
'    If Not rs.EOF Then      'Si ya existe el codigo
'        Err.Raise ERR_REPITECODIGO, "PCGGasto", MSGERR_REPITECODIGO
'    End If
'    rs.Close
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM pcGestion WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM pcGestion WHERE Id=" & mProps.Id
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)

        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "pcGestion_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro
    With rs
        !CodGestion = mProps.CodGestion
        !Descripcion = mProps.Descripcion
        !BandValida = mProps.BandValida
        !FechaGrabado = Now
        .Update
        .Move 0             'Para actualizar
        mProps.Id = !Id
        .Close
    End With
    Set rs = Nothing
 
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_DES", "PCGestion Nuevo: " & _
                                                         mProps.CodGestion, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_DES", "PCGestion Modificado: " & _
                                                         mProps.CodGestion, "DT"
    End If
    mbooEsNuevo = False
    mbooModificado = False
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PCGESTION
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String
    Dim rs As Recordset
    Dim v As Variant

    
        If VarType(key) = vbString Then
            sql = "SELECT *"
            sql = sql & " FROM PCGestion "
            sql = sql & " WHERE codGestion='" & key & "'"
        Else
            sql = "SELECT *"
            sql = sql & " FROM PCGestion "
            sql = sql & " WHERE id=" & key
        End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'Recupera los datos de Usuario
    With rs
        mProps.Id = !Id
        mProps.CodGestion = !CodGestion
        mProps.Descripcion = !Descripcion
        mProps.BandValida = !BandValida
        mProps.FechaGrabado = !FechaGrabado
        
        .Close
    End With
       
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
  
    sql = "DELETE FROM pcgestion WHERE Id=" & mProps.Id
    mProps.objEmpresa.Execute sql, True
  
    'Elimina el registro
    sql = "DELETE FROM pcgestion WHERE Id=" & mProps.Id
    mProps.objEmpresa.Execute sql, True
    
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PCGESTION
End Sub

Public Function Clone() As PCGestion
    Dim nuevo As PCGestion
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New PCGestion
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.Id = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Get Id() As Long
    Id = mProps.Id
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

Friend Property Let Modificado(ByVal value As Boolean)
    mbooModificado = value
End Property


Private Sub Class_Initialize()
    'Valores predeterminados
    mbooEsNuevo = True
    mbooModificado = False
    
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Property Let CodGestion(ByVal value As String)
    mProps.CodGestion = value
    mbooModificado = True
End Property

Public Property Get CodGestion() As String
    CodGestion = mProps.CodGestion
End Property

Public Property Let BandValida(ByVal value As Boolean)
    mProps.BandValida = value
    mbooModificado = True
End Property

Public Property Get BandValida() As Boolean
    BandValida = mProps.BandValida
End Property

Public Property Let Descripcion(ByVal value As String)
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = mProps.Descripcion
End Property

