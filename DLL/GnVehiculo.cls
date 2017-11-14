VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GnVehiculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Private Type T_PROPS
    IdVehiculo As Long
    CodVehiculo As String
    Descripcion As String
    IdCliente As Long
    CodCliente As String
    IdGrupo(1 To 4) As Long
    CodGrupo(1 To 4) As String
    FechaGrabado As Date
    
    Dimension As String
    BandValida As Boolean
    BandLibre As Boolean
    RutaFoto As String
    PlanMant As String
    
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
    sql = "SELECT CodVehiculo FROM GNVehiculo " & _
          "WHERE (CodVehiculo='" & mProps.CodVehiculo & "') AND " & _
          "(IdVehiculo<>" & mProps.IdVehiculo & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "GNVehiculo", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM GNVehiculo WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM GNVehiculo WHERE IdVehiculo=" & mProps.IdVehiculo
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "GNVehiculo_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en GNVehiculo
    With rs
        !CodVehiculo = UCase(mProps.CodVehiculo)
        !Descripcion = mProps.Descripcion
        !IdCliente = mProps.IdCliente
        
        For i = 1 To GNVGRUPO_MAX
            .Fields("IdGrupo" & i) = mProps.IdGrupo(i)
        Next i
        
        !Dimension = mProps.Dimension
        !BandValida = mProps.BandValida
        !BandLibre = mProps.BandLibre
        !RutaFoto = mProps.RutaFoto
        !PlanMant = mProps.PlanMant
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.FechaGrabado = !FechaGrabado
        mProps.IdVehiculo = !IdVehiculo
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_VEH", "Vehiculo  Nuevo: " & _
                                                         mProps.CodVehiculo, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_VEH", "Vehiculo Modificado: " & _
                                                         mProps.CodVehiculo, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_GNVEHICULO
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset, i As Integer
    
    If VarType(key) = vbString Then
        sql = "SELECT "
        sql = sql & " IdVehiculo, CodVehiculo, gnc.Descripcion , Dimension, GNC.BandValida, BandLibre, RutaFoto, "
        sql = sql & " gnc.FechaGrabado, "
        sql = sql & " IdCliente, "
        sql = sql & " gnc.IdGrupo1, g1.codgrupo1,"
        sql = sql & " gnc.IdGrupo2, g2.codgrupo2,"
        sql = sql & " gnc.IdGrupo3, g3.codgrupo3,"
        sql = sql & " gnc.IdGrupo4, g4.codgrupo4, "
        sql = sql & " PlanMant"
        sql = sql & " FROM GNVehiculo gnc"
        sql = sql & " left join GNVGRUPO1 g1 on gnc.idgrupo1=g1.idgrupo1"
        sql = sql & " left join GNVGRUPO2 g2 on gnc.idgrupo2=g2.idgrupo2"
        sql = sql & " left join GNVGRUPO3 g3 on gnc.idgrupo3=g3.idgrupo3"
        sql = sql & " left join GNVGRUPO4 g4 on gnc.idgrupo4=g4.idgrupo4"
'        sql = sql & " left join IVPlan on ivplan.idplan=gnc.idplan"
        sql = sql & " WHERE CodVehiculo='" & key & "'"
    Else
        sql = "SELECT "
        sql = sql & " IdVehiculo, CodVehiculo, gnc.Descripcion, Dimension, GNC.BandValida, BandLibre, RutaFoto,"
        sql = sql & " gnc.FechaGrabado,"
        sql = sql & " IdCliente, "
        sql = sql & " gnc.IdGrupo1, g1.codgrupo1,G1.descripcion,"
        sql = sql & " gnc.IdGrupo2, g2.codgrupo2,G2.descripcion,"
        sql = sql & " gnc.IdGrupo3, g3.codgrupo3,G3.descripcion,"
        sql = sql & " gnc.IdGrupo4, g4.codgrupo4,G4.descripcion"
        sql = sql & " PlanMant"
        sql = sql & " FROM GNVehiculo gnc"
        sql = sql & " left join GNVGRUPO1 g1 on gnc.idgrupo1=g1.idgrupo1"
        sql = sql & " left join GNVGRUPO2 g2 on gnc.idgrupo2=g2.idgrupo2"
        sql = sql & " left join GNVGRUPO3 g3 on gnc.idgrupo3=g3.idgrupo3"
        sql = sql & " left join GNVGRUPO4 g4 on gnc.idgrupo4=g4.idgrupo4"
'        sql = sql & " left join IVPlan on ivplan.idplan=gnc.idplan"
        sql = sql & " WHERE IdVehiculo=" & key
        
    End If

    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IdVehiculo = !IdVehiculo
        mProps.CodVehiculo = !CodVehiculo
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!IdCliente) Then IdCliente = !IdCliente
        For i = 1 To GNVGRUPO_MAX
            If Not IsNull(.Fields("IdGrupo" & i)) Then mProps.IdGrupo(i) = .Fields("IdGrupo" & i)
            If Not IsNull(.Fields("CodGrupo" & i)) Then mProps.CodGrupo(i) = .Fields("CodGrupo" & i)
        Next i
        
        If Not IsNull(!Dimension) Then mProps.Dimension = !Dimension
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!BandLibre) Then mProps.BandLibre = !BandLibre
        If Not IsNull(!RutaFoto) Then mProps.RutaFoto = !RutaFoto  ' AUC 04/05/06
'        If Not IsNull(!IdPlan) Then mProps.IdPlan = !IdPlan
        If Not IsNull(!PlanMant) Then mProps.PlanMant = !PlanMant
        

        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    
    'Verifica si existe un comprobante que usa el codigo de Vehiculo de costo
    sql = "SELECT IDVehiculo FROM GNCentrocosto WHERE IdVehiculo=" & mProps.IdVehiculo
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        rs.Close
        Set rs = Nothing
        Err.Raise ERR_NOELIMINABLE, "GNVehiculo.Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close
    Set rs = Nothing
    
    'Elimina el Vehiculo de costo
    sql = "DELETE FROM GNVehiculo WHERE IdVehiculo=" & mProps.IdVehiculo
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_VEH", "Vehiculo  Eliminado: " & _
                                                         mProps.CodVehiculo, "DT"

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_GNVEHICULO
End Sub

Friend Function Clone() As GnVehiculo
    Dim nuevo As GnVehiculo
    
    Set nuevo = New GnVehiculo
    'Copia las propiedades al nuevo objeto
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdVehiculo = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property





Public Property Let CodVehiculo(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNVehiculo_CodVehiculo", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodVehiculo = value
    mbooModificado = True
End Property

Public Property Get CodVehiculo() As String
    CodVehiculo = Trim$(mProps.CodVehiculo)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "GNVehiculo_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

'*** jeaa 30/mar/09 Agregado
Public Property Let IdCliente(ByVal value As Long)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.IdCliente Then
        'Actualiza el Codigo también
        mProps.CodCliente = ""
        sql = "SELECT CodProvCli FROM PCProvCli WHERE IdProvCli=" & value
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then mProps.CodCliente = rs!CodProvCli
        rs.Close
        Set rs = Nothing
    End If
    
    mProps.IdCliente = value
    mbooModificado = True
End Property

'*** jeaa 30/mar/09 Agregado
Public Property Get IdCliente() As Long
    IdCliente = mProps.IdCliente
End Property

Public Property Let CodCliente(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCliente Then
        If Len(value) > 0 Then
            'Actualiza IdCliente también
            sql = "SELECT IdProvCli FROM PCProvCli WHERE CodProvCli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNVehiculo.CodCliente", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdCliente = rs!IdProvCli
            End If
            rs.Close
        Else
            mProps.IdCliente = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodCliente = value
    mbooModificado = True
End Property

Public Property Get CodCliente() As String
    CodCliente = Trim$(mProps.CodCliente)
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
    mProps.BandLibre = True
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Property Let IdGrupo(NumGrupo As Integer, ByVal value As Long)
    mProps.IdGrupo(NumGrupo) = value
    mbooModificado = True
End Property

Public Property Get IdGrupo(NumGrupo As Integer) As Long
    IdGrupo = mProps.IdGrupo(NumGrupo)
End Property


Public Property Let CodGrupo(NumGrupo As Integer, ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodGrupo(NumGrupo) Then
        If Len(value) > 0 Then
            'Actualiza IdGrupo1 también
            sql = "SELECT IdGrupo" & NumGrupo & " FROM GNVGRUPO" & NumGrupo & " WHERE CodGrupo" & NumGrupo & "='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNVGRUPO_CodGrupo" & NumGrupo, MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdGrupo(NumGrupo) = rs.Fields("IdGrupo" & NumGrupo)
            End If
            rs.Close
        Else
            IdGrupo(NumGrupo) = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodGrupo(NumGrupo) = value
    mbooModificado = True
End Property

Public Property Get CodGrupo(NumGrupo As Integer) As String
    CodGrupo = Trim$(mProps.CodGrupo(NumGrupo))
End Property



Public Property Get IdVehiculo() As Long
    IdVehiculo = mProps.IdVehiculo
End Property

Public Property Let Dimension(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "GNVehiculo_Dimesion", MSGERR_INVALIDO
    End If
    
    mProps.Dimension = value
    mbooModificado = True
End Property

Public Property Get Dimension() As String
    Dimension = Trim$(mProps.Dimension)
End Property

Public Property Let BandValida(ByVal value As Boolean)
    mProps.BandValida = value
    mbooModificado = True
End Property

Public Property Get BandValida() As Boolean
    BandValida = mProps.BandValida
End Property

Public Property Let BandLibre(ByVal value As Boolean)
    mProps.BandLibre = value
    mbooModificado = True
End Property

Public Property Get BandLibre() As Boolean
    BandLibre = mProps.BandLibre
End Property

Public Property Let RutaFoto(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "IVInventario_RutaFoto", MSGERR_INVALIDO & "Direccion muy larga"
    End If
    mProps.RutaFoto = value
    mbooModificado = True
End Property

Public Property Get RutaFoto() As String
    RutaFoto = Trim$(mProps.RutaFoto)
End Property

'----
Public Property Get PlanMant() As String
    PlanMant = mProps.PlanMant
End Property

Public Property Let PlanMant(ByVal value As String)
    mProps.PlanMant = value
    mbooModificado = True
End Property


