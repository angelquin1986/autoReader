VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Proceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type T_PROPS
    idproceso As Long
    CodProceso As String
    Descripcion As String
    BandValida As Boolean
    fechagrabado As Date
    codUsuario As String
    secuencia As Integer
    CodProcesoV As String
    CodTrans As String
    EjecutaCola As Integer
    OrdenesMostrar As Integer
    Grupo As String
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
    sql = "SELECT CodProceso FROM IVProceso " & _
          " WHERE (CodProceso)='" & mProps.CodProceso & "' AND " & _
                 "(Idproceso <>" & mProps.idproceso & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "IVProceso_Grabar", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM IVproceso WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM IVProceso WHERE Idproceso =" & mProps.idproceso
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "IVproceso_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en Pos
    With rs
        !CodProceso = mProps.CodProceso
        !Descripcion = mProps.Descripcion
        !BandValida = mProps.BandValida
        !codUsuario = mProps.codUsuario
        !secuencia = mProps.secuencia
        !CodProcesoV = mProps.CodProcesoV
        !CodTrans = mProps.CodTrans
        !EjecutaCola = mProps.EjecutaCola
        !OrdenesMostrar = mProps.OrdenesMostrar
        !Grupo = mProps.Grupo
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.idproceso = !idproceso
        .Close
    End With
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_UNI", "IVProceso Nuevo: " & _
                                                         mProps.CodProceso, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_UNI", "IVProceso Modificado: " & _
                                                         mProps.CodProceso, "DT"
    End If
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVPROCESO
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    If VarType(key) = vbString Then
        sql = "SELECT * FROM IVproceso WHERE Codproceso='" & key & "'"
    Else
        sql = "SELECT * FROM IVproceso WHERE IdProceso=" & key
    End If
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    'Recupera los datos de Motivo
    With rs
        mProps.idproceso = !idproceso
        mProps.CodProceso = !CodProceso
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!secuencia) Then mProps.secuencia = !secuencia
        If Not IsNull(!codUsuario) Then mProps.codUsuario = !codUsuario
        If Not IsNull(!CodProcesoV) Then mProps.CodProcesoV = !CodProcesoV
        If Not IsNull(!CodTrans) Then mProps.CodTrans = !CodTrans
        If Not IsNull(!EjecutaCola) Then mProps.EjecutaCola = !EjecutaCola
        If Not IsNull(!OrdenesMostrar) Then mProps.OrdenesMostrar = !OrdenesMostrar
        If Not IsNull(!Grupo) Then mProps.Grupo = !Grupo
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
    
    'Verifica si ya existe un item usando el mismo CodMotivo
    'VER DESPUES AUC
'    sql = "SELECT CodTrans FROM GnComprobante" & _
'          " WHERE IdMotivo=" & mProps.IdMotivo
'    Set rs = mProps.objEmpresa.OpenRecordset(sql)
'    If Not rs.EOF Then
'        Err.Raise ERR_NOELIMINABLE, "Motivo_Eliminar", MSGERR_NOELIMINABLE
'    End If
'    rs.Close

    'Elimina el registro
    sql = "DELETE FROM IVproceso WHERE Idproceso=" & mProps.idproceso
    
    mProps.objEmpresa.Execute sql, True
    
    mProps.objEmpresa.GrabaGNLogAccion "ELI_IVPROCESO", "IVProceso Eliminado: " & _
                                                         mProps.CodProceso, "DT"
                                                         
    Set rs = Nothing
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVPROCESO
End Sub

Friend Function Clone() As Proceso
    Dim nuevo As Proceso
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New Proceso
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.idproceso = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Let CodProceso(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "proceso_Codproceso", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodProceso = value
    mbooModificado = True
End Property

Public Property Get CodProceso() As String
    CodProceso = Trim$(mProps.CodProceso)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "Unidad_Descripcion", MSGERR_INVALIDO
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

Public Property Let secuencia(ByVal value As Integer)
    mProps.secuencia = value
    mbooModificado = True
End Property

Public Property Get secuencia() As Integer
    secuencia = mProps.secuencia
End Property

Public Property Let codUsuario(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "Procesos", MSGERR_INVALIDO
    End If
    mProps.codUsuario = value
    mbooModificado = True
End Property

Public Property Get codUsuario() As String
    codUsuario = mProps.codUsuario
End Property

Public Property Let CodProcesoV(ByVal value As String)
'    Dim sql As String, rs As Recordset
'    If value <> mProps.CodProceso1 Then
'        If Len(value) > 0 Then
'            sql = "SELECT IdProceso FROM ivproceso WHERE CodProceso='" & value & "'"
'            Set rs = mProps.objEmpresa.OpenRecordset(sql)
'            If rs.EOF Then
'                Err.Raise ERR_NOHAYCODIGO, "IVProceso_CodProceso1", MSGERR_NOHAYCODIGO
'                Exit Property
'            Else
'                IdProceso1 = rs!IdProceso
'            End If
'            rs.Close
'        Else
'            IdProceso1 = 0
'        End If
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "Procesos", MSGERR_INVALIDO
    End If
    mProps.CodProcesoV = value
    mbooModificado = True
'    End If
'    Set rs = Nothing
End Property
    

Public Property Get CodProcesoV() As String
'Dim sql As String
'Dim rs As Recordset
'            sql = "SELECT codProceso FROM ivproceso WHERE idProceso=" & mProps.IdProceso1
'            Set rs = mProps.objEmpresa.OpenRecordset(sql)
'            If rs.EOF Then
'                CodProceso1 = ""
'            Else
                CodProcesoV = mProps.CodProcesoV
'            End If
'            rs.Close
    
End Property
Public Property Get IdProceso1() As Long
'    IdProceso1 = mProps.IdProceso1
End Property

Public Property Let IdProceso1(ByVal value As Long)
'    mProps.IdProceso1 = value
'    mbooModificado = True
End Property
Public Property Let CodTrans(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "Procesos_codtrans", MSGERR_INVALIDO
    End If
    mProps.CodTrans = value
    mbooModificado = True
End Property

Public Property Get CodTrans() As String
    CodTrans = mProps.CodTrans
End Property
Friend Sub recuperarxUsuario(key As String)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT * FROM IVproceso WHERE "
    'sql = sql & " CHARINDEX('" & key & "',codusuario) > 0 "
    sql = sql & " CHARINDEX('" & key & "',Grupo) > 0 "
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    With rs
        mProps.idproceso = !idproceso
        mProps.CodProceso = !CodProceso
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!secuencia) Then mProps.secuencia = !secuencia
        If Not IsNull(!codUsuario) Then mProps.codUsuario = !codUsuario
        If Not IsNull(!CodProcesoV) Then mProps.CodProcesoV = !CodProcesoV
        If Not IsNull(!CodTrans) Then mProps.CodTrans = !CodTrans
        If Not IsNull(!EjecutaCola) Then mProps.EjecutaCola = !EjecutaCola
        If Not IsNull(!Grupo) Then mProps.Grupo = !Grupo
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Property Let EjecutaCola(ByVal value As Integer)
    mProps.EjecutaCola = value
    mbooModificado = True
End Property

Public Property Get EjecutaCola() As Integer
    EjecutaCola = mProps.EjecutaCola
End Property

Public Property Let OrdenesMostrar(ByVal value As Integer)
    mProps.OrdenesMostrar = value
    mbooModificado = True
End Property

Public Property Get OrdenesMostrar() As Integer
    OrdenesMostrar = mProps.OrdenesMostrar
End Property

Public Property Let Grupo(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Procesos_grupo", MSGERR_INVALIDO
    End If
    mProps.Grupo = value
    mbooModificado = True
End Property

Public Property Get Grupo() As String
    Grupo = mProps.Grupo
End Property
