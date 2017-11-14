VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GNTransporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Private Type T_PROPS
    IdTransporte As Long
    CodTransporte As String
    Descripcion As String
    FechaGrabado As Date
    Observacion As String
    BandValida As Boolean
    Observacion1 As String
    Observacion2 As String
    
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean


Public Sub Grabar()
    Dim sql As String, rs As Recordset, i As Integer
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
    sql = "SELECT CodTransporte FROM GNTransporte " & _
          "WHERE (CodTransporte='" & mProps.CodTransporte & "') AND " & _
          "(IdTransporte<>" & mProps.IdTransporte & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "GNTransporte", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM GNTransporte WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM GNTransporte WHERE IdTransporte=" & mProps.IdTransporte
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "GNTransporte_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en GNTransporte
    With rs
        !CodTransporte = UCase(mProps.CodTransporte)
        !Descripcion = mProps.Descripcion
        !Observacion = mProps.Observacion
        !BandValida = mProps.BandValida
        !Observacion1 = mProps.Observacion1
        !Observacion2 = mProps.Observacion2
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.FechaGrabado = !FechaGrabado
        mProps.IdTransporte = !IdTransporte
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_VEH", "Transporte  Nuevo: " & _
                                                         mProps.CodTransporte, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_VEH", "Transporte Modificado: " & _
                                                         mProps.CodTransporte, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_GNTRANSPORTE
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset, i As Integer
    
    If VarType(key) = vbString Then
        sql = "SELECT "
        sql = sql & " IdTransporte, CodTransporte, Descripcion , Observacion, BandValida, FechaGrabado,observacion1,observacion2 "
        sql = sql & " FROM GNTransporte "
        sql = sql & " WHERE CodTransporte='" & key & "'"
    Else
        sql = "SELECT "
        sql = sql & " IdTransporte, CodTransporte, Descripcion , Observacion, BandValida, FechaGrabado,observacion1,observacion2 "
        sql = sql & " FROM GNTransporte "
        sql = sql & " WHERE IdTransporte=" & key
        
    End If

    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IdTransporte = !IdTransporte
        mProps.CodTransporte = !CodTransporte
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!Observacion) Then mProps.Observacion = !Observacion
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!Observacion1) Then mProps.Observacion1 = !Observacion1
        If Not IsNull(!Observacion2) Then mProps.Observacion2 = !Observacion2
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    
    'Verifica si existe un comprobante que usa el codigo de Transporte de costo
'    sql = "SELECT IDTransporte FROM GNOferta WHERE IdTransporte=" & mProps.IdTransporte
'    Set rs = mProps.objEmpresa.OpenRecordset(sql)
'    If Not rs.EOF Then
'        rs.Close
'        Set rs = Nothing
'        Err.Raise ERR_NOELIMINABLE, "GNTransporte.Eliminar", MSGERR_NOELIMINABLE
'    End If
'    rs.Close
'    Set rs = Nothing
    
    'Elimina el Transporte de costo
    sql = "DELETE FROM GNTransporte WHERE IdTransporte=" & mProps.IdTransporte
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_TRA", "Transporte  Eliminado: " & _
                                                         mProps.CodTransporte, "DT"

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_GNTRANSPORTE
End Sub

Friend Function Clone() As GNTransporte
    Dim nuevo As GNTransporte
    
    Set nuevo = New GNTransporte
    'Copia las propiedades al nuevo objeto
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdTransporte = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property





Public Property Let CodTransporte(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNTransporte_CodTransporte", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodTransporte = value
    mbooModificado = True
End Property

Public Property Get CodTransporte() As String
    CodTransporte = Trim$(mProps.CodTransporte)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "GNTransporte_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
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
    mProps.BandValida = True
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub


Public Property Get IdTransporte() As Long
    IdTransporte = mProps.IdTransporte
End Property

Public Property Let Observacion(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "GNTransporte_Dimesion", MSGERR_INVALIDO
    End If
    
    mProps.Observacion = value
    mbooModificado = True
End Property

Public Property Get Observacion() As String
    Observacion = Trim$(mProps.Observacion)
End Property

Public Property Let BandValida(ByVal value As Boolean)
    mProps.BandValida = value
    mbooModificado = True
End Property

Public Property Get BandValida() As Boolean
    BandValida = mProps.BandValida
End Property


Public Property Let Observacion1(ByVal value As String)
    If Len(value) > 20 Then 'AUC aqui voy a guardar el telefono del chofer para ricuartesa
        Err.Raise ERR_INVALIDO, "GNTransporte_observacion1", MSGERR_INVALIDO
    End If
    
    mProps.Observacion1 = value
    mbooModificado = True
End Property

Public Property Get Observacion1() As String
    Observacion1 = Trim$(mProps.Observacion1)
End Property


Public Property Let Observacion2(ByVal value As String)
    If Len(value) > 20 Then 'AUC aqui voy a guardar el num licencia para ricuartesa
        Err.Raise ERR_INVALIDO, "GNTransporte_observacion2", MSGERR_INVALIDO
    End If
    
    mProps.Observacion2 = value
    mbooModificado = True
End Property

Public Property Get Observacion2() As String
    Observacion2 = Trim$(mProps.Observacion2)
End Property
