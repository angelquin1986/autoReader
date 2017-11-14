VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GNProyecto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Private Type T_PROPS
    IdProyecto As Long
    CodProyecto As String
    IdVendedor As Long
    CodVendedor As String
    Propietario As String
    Descripcion As String
    IdProvincia As Long
    CodProvincia As String
    IdCanton As Long
    CodCanton As String
    
    Nombre1 As String
    Cargo1 As String
    Telefono1 As String
    Obser1 As String
    
    Nombre2 As String
    Cargo2 As String
    Telefono2 As String
    Obser2 As String
    
    
    Nombre3 As String
    Cargo3 As String
    Telefono3 As String
    Obser3 As String
    
    
    
    Observacion As String
    FechaInicio As Date
    FechaFinal As Variant
    
    IdInventario As Long
    CodInventario As String
    
    
    Cantidad As Currency
    Estado As Integer
    
    FechaProyeccion As Date
    EstadoNegociacion As Integer
    
    Num1 As String
    Num2 As String
    Observacion2 As String
    IdCompetencia As Long
    CodCompetencia As String
    
    
    
    FechaGrabado As Date
    objEmpresa As Empresa       'Referencia al objeto padre
    
    
End Type

Private mcolDetalleProyecto As Collection      'Coleccion de detalles de familia

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean


Public Sub Grabar()
    Dim sql As String, rs As Recordset, i As Integer
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
    If mbooEsNuevo Then
        sql = "SELECT CodProyecto FROM GNProyecto " & _
              "WHERE (CodProyecto='" & mProps.CodProyecto & "') AND " & _
              "(IdProyecto<>" & mProps.IdProyecto & ")"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then      'Si ya existe el codigo
            Err.Raise ERR_REPITECODIGO, "GNProyecto", MSGERR_REPITECODIGO
        End If
        rs.Close
    End If
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM GNProyecto WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM GNProyecto WHERE IdProyecto=" & mProps.IdProyecto
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "GNProyecto_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en GNProyecto
    With rs
        !CodProyecto = mProps.CodProyecto
        !IdVendedor = mProps.IdVendedor
        !Propietario = mProps.Propietario
        !Descripcion = mProps.Descripcion
        !IdProvincia = mProps.IdProvincia
        !IdCanton = mProps.IdCanton
        

        !Nombre1 = mProps.Nombre1
        !Cargo1 = mProps.Cargo1
        !Telefono1 = mProps.Telefono1
        !Obser1 = mProps.Obser1
        
        !Nombre2 = mProps.Nombre2
        !Cargo2 = mProps.Cargo2
        !Telefono2 = mProps.Telefono2
        !Obser2 = mProps.Obser2
        
        !Nombre3 = mProps.Nombre3
        !Cargo3 = mProps.Cargo3
        !Telefono3 = mProps.Telefono3
        !Obser3 = mProps.Obser3
        
        !FechaInicio = mProps.FechaInicio
        If Not IsNull(mProps.FechaFinal) Then
            !FechaFinal = mProps.FechaFinal
        Else
            !FechaFinal = Null
        End If
        
        '*** MAKOTO 05/mar/01 Agregado

        !Observacion = mProps.Observacion
        !IdInventario = mProps.IdInventario
        !Cantidad = mProps.Cantidad
        !Estado = mProps.Estado
        !FechaProyeccion = mProps.FechaProyeccion
        !EstadoNegociacion = mProps.EstadoNegociacion
        !Observacion2 = mProps.Observacion2
        !Num2 = mProps.Num1
        !Num2 = mProps.Num2
        !IdCompetencia = mProps.IdCompetencia

'        GrabaProyectoDetalle
        
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.FechaGrabado = !FechaGrabado
        mProps.IdProyecto = !IdProyecto
        GrabaProyectoDetalle
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_PR", "Proyecto de Costo Nuevo: " & _
                                                         mProps.CodProyecto, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_PR", "Proyecto de Costo Modificado: " & _
                                                         mProps.CodProyecto, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_GNPROYECTO
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset, i As Integer
    
    If VarType(key) = vbString Then
        sql = "SELECT "
        sql = sql & " IdProyecto, CodProyecto, gnp.Descripcion , FechaInicio, gnp.Observacion, "
        sql = sql & " FechaFinal, gnp.FechaGrabado, Propietario, IdVendedor,"
        sql = sql & " gnc.IdProvincia, Nombre1,Nombre2,Nombre3, Cargo1,Cargo2,Cargo3,Telefono1,Telefono2,Telefono3, Obser1,Obser2, Obser3, "
        sql = sql & " gnc.idCanton, CodCanton, gnp.idinventario, ivi.CodInventario, Cantidad,FechaProyeccion, Estado, "
        sql = sql & " EstadoNegociacion, gnp.Observacion, gnp.num1, gnp.num2, gnp.Observacion2, gnp.idcompetencia, gncmp.CodCompetencia "
        sql = sql & " FROM GNProyecto gnp left join ivinventario ivi on gnp.idinventario= ivi.idinventario "
        sql = sql & " left join  PCCanton gnc on gnc.idCanton=gnp.idCanton"
        sql = sql & " left join  gnCompetencia gncmp on gnp.idCompetencia=gncmp.idCompetencia "
        sql = sql & " WHERE CodProyecto='" & key & "'"
    Else
        sql = "SELECT "
        sql = sql & " IdProyecto, CodProyecto, gnp.Descripcion, FechaInicio,"
        sql = sql & " FechaFinal, gnp.FechaGrabado, Propietario, IdVendedor,"
        sql = sql & " IdProvincia, Observacion, gnc.IdInventario, Cantidad, estado "
        sql = sql & " FROM GNProyecto gnc"
        sql = sql & " WHERE IdProyecto=" & key
        
    End If

    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario

    With rs

        mProps.IdProyecto = !IdProyecto
        mProps.CodProyecto = !CodProyecto
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!Propietario) Then mProps.Propietario = !Propietario     '*** MAKOTO 14/feb/01 Agregado
        If Not IsNull(!FechaInicio) Then mProps.FechaInicio = !FechaInicio

        mProps.FechaFinal = !FechaFinal         'Null tambien es valido
        
        '*** MAKOTO 05/mar/01 Agregado
        If Not IsNull(!IdVendedor) Then IdVendedor = !IdVendedor
        If Not IsNull(!IdProvincia) Then IdProvincia = !IdProvincia
        If Not IsNull(!Observacion) Then mProps.Observacion = !Observacion
        
        
        If Not IsNull(!Nombre1) Then mProps.Nombre1 = !Nombre1
        If Not IsNull(!Cargo1) Then mProps.Cargo1 = !Cargo1
        If Not IsNull(!Telefono1) Then mProps.Telefono1 = !Telefono1
        If Not IsNull(!Obser1) Then mProps.Obser1 = !Obser1
        

        If Not IsNull(!Nombre2) Then mProps.Nombre2 = !Nombre2
        If Not IsNull(!Cargo2) Then mProps.Cargo2 = !Cargo2
        If Not IsNull(!Telefono2) Then mProps.Telefono2 = !Telefono2
        If Not IsNull(!Obser2) Then mProps.Obser2 = !Obser2

        If Not IsNull(!Nombre3) Then mProps.Nombre3 = !Nombre3
        If Not IsNull(!Cargo3) Then mProps.Cargo3 = !Cargo3
        If Not IsNull(!Telefono3) Then mProps.Telefono3 = !Telefono3
        If Not IsNull(!Obser3) Then mProps.Obser3 = !Obser3


        If Not IsNull(!IdCanton) Then IdCanton = !IdCanton
        If Not IsNull(!IdInventario) Then IdInventario = .Fields("IdInventario")
        If Not IsNull(!Cantidad) Then Cantidad = .Fields("Cantidad")
        If Not IsNull(.Fields("Estado")) Then mProps.Estado = .Fields("Estado")
        If Not IsNull(!FechaProyeccion) Then mProps.FechaProyeccion = !FechaProyeccion
        If Not IsNull(.Fields("EstadoNegociacion")) Then mProps.EstadoNegociacion = .Fields("EstadoNegociacion")
        If Not IsNull(!Observacion2) Then mProps.Observacion2 = !Observacion2
        If Not IsNull(!Num1) Then mProps.Num1 = !Num1
        If Not IsNull(!Num2) Then mProps.Num2 = !Num2
        If Not IsNull(!IdCompetencia) Then IdCompetencia = .Fields("Idcompetencia")
        If Not IsNull(!CodCompetencia) Then CodCompetencia = .Fields("CodCompetencia")
                
        CargaDetalleProyecto
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    
    'Verifica si existe un comprobante que usa el codigo de Proyecto de costo
    sql = "SELECT TransID FROM GNComprobante WHERE IdProyecto=" & mProps.IdProyecto
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        rs.Close
        Set rs = Nothing
        Err.Raise ERR_NOELIMINABLE, "GNProyecto.Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close
    Set rs = Nothing
    
    'Elimina el Proyecto de costo
    sql = "DELETE FROM GNProyecto WHERE IdProyecto=" & mProps.IdProyecto
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_CC", "Proyecto de Costo Eliminado: " & _
                                                         mProps.CodProyecto, "DT"

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_GNPROYECTO
End Sub

Friend Function Clone() As GNProyecto
    Dim nuevo As GNProyecto
    
    Set nuevo = New GNProyecto
    'Copia las propiedades al nuevo objeto
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdProyecto = 0
    mProps.Estado = 0
    mProps.EstadoNegociacion = 0
    mProps.FechaFinal = Null
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property


Public Property Let CodProyecto(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNProyecto_CodProyecto", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodProyecto = value
    mbooModificado = True
End Property

Public Property Get CodProyecto() As String
    CodProyecto = Trim$(mProps.CodProyecto)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "GNProyecto_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

'*** MAKOTO 14/feb/01 Agregado
Public Property Let Propietario(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "GNProyecto.Propietario", MSGERR_INVALIDO
    End If
    
    mProps.Propietario = value
    mbooModificado = True
End Property

'*** MAKOTO 14/feb/01 Agregado
Public Property Get Propietario() As String
    Propietario = Trim$(mProps.Propietario)
End Property

'*** MAKOTO 05/mar/01 Agregado
Public Property Let IdVendedor(ByVal value As Long)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.IdVendedor Then
        'Actualiza el Codigo también
        mProps.CodVendedor = ""
        sql = "SELECT CodVendedor FROM FcVendedor WHERE IdVendedor=" & value
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then mProps.CodVendedor = rs!CodVendedor
        rs.Close
        Set rs = Nothing
    End If
    
    mProps.IdVendedor = value
    mbooModificado = True
End Property

'*** MAKOTO 05/mar/01 Agregado
Public Property Get IdVendedor() As Long
    IdVendedor = mProps.IdVendedor
End Property

Public Property Let CodVendedor(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodVendedor Then
        If Len(value) > 0 Then
            'Actualiza IdVendedor también
            sql = "SELECT IdVendedor FROM FcVendedor WHERE CodVendedor='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNProyecto.CodVendedor", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdVendedor = rs!IdVendedor
            End If
            rs.Close
        Else
            mProps.IdVendedor = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodVendedor = value
    mbooModificado = True
End Property

Public Property Get CodVendedor() As String
    CodVendedor = Trim$(mProps.CodVendedor)
End Property

'*** MAKOTO 05/mar/01 Agregado
Public Property Let IdProvincia(ByVal value As Long)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.IdProvincia Then
        'Actualiza el Codigo también
        mProps.CodProvincia = ""
        sql = "SELECT CodProvincia FROM PCProvincia WHERE IdProvincia=" & value
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then mProps.CodProvincia = rs!CodProvincia
        rs.Close
        Set rs = Nothing
    End If
    
    mProps.IdProvincia = value
    mbooModificado = True
End Property

'*** MAKOTO 05/mar/01 Agregado
Public Property Get IdProvincia() As Long
    IdProvincia = mProps.IdProvincia
End Property

Public Property Let CodProvincia(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodProvincia Then
        If Len(value) > 0 Then
            'Actualiza IdProvincia también
            sql = "SELECT IdProvincia FROM PCProvincia WHERE CodProvincia='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNProyecto.CodProvincia", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdProvincia = rs!IdProvincia
            End If
            rs.Close
        Else
            mProps.IdProvincia = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodProvincia = value
    mbooModificado = True
End Property

Public Property Get CodProvincia() As String
    CodProvincia = Trim$(mProps.CodProvincia)
End Property



Public Property Let FechaInicio(ByVal value As Date)
    mProps.FechaInicio = value
    mbooModificado = True
End Property

Public Property Get FechaInicio() As Date
    FechaInicio = mProps.FechaInicio
End Property

Public Property Let FechaFinal(ByVal value As Variant)
    If Not IsNull(value) Then
        If Not IsDate(value) Then
            Err.Raise ERR_INVALIDO, "GNProyecto_FechaFinal", MSGERR_INVALIDO
        Else
            mProps.FechaFinal = CDate(value)
        End If
    Else
        mProps.FechaFinal = Null
    End If
    
    mbooModificado = True
End Property

Public Property Get FechaFinal() As Variant
    FechaFinal = mProps.FechaFinal
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

    'Valores predeterminados
    With mProps
        .FechaInicio = Date
        .FechaProyeccion = Date
        .Estado = 0
        .EstadoNegociacion = 0
        .FechaFinal = Null
    End With
    
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub


Public Sub GeneraCodProyecto()
    Dim sql As String, pre As String, num As Integer, rs As Recordset
    Dim lon_pre As Integer, i As Integer, v As Variant
    
    '-----
    ' Genera el código de Proyecto en el siguiente formato
    '   XXXYYYYMMDD-nnn
    '     XXX = PrefijoCodigo de GNOpcion (Longitud puede variar)
    '     YYYYMMDD = Fecha actual
    '     nnn = Número secuencial dentro de la fecha
    '-----
    
    'Busca el valor maximo de los últimos 3 dígitos de CodProyecto de la misma fecha
    pre = mProps.objEmpresa.GNOpcion.PrefijoCodCC
    lon_pre = Len(pre)
    
    Select Case Me.Empresa.GNOpcion.GeneraCodCC
    Case 0  'Manual
    Case 1  'Secuencial
        'Falta
#If DAOLIB Then
        sql = "SELECT Max(Right([CodProyecto],3)) AS CodMax " & _
              "FROM GNProyecto " & _
              "WHERE (Left$([CodProyecto]," & (Len(pre) + 8) & ")='" & _
                        pre & Format(Date, "yyyymmdd") & "')"
#Else
        sql = "SELECT case  " & _
                     "when isnumeric(substring(CodProyecto," & (lon_pre + 1) & ",len(codProyecto)))=1  " & _
                          "then substring(CodProyecto," & (lon_pre + 1) & ",len(codProyecto)) " & _
                     "when isnumeric(substring(CodProyecto," & (lon_pre + 1) & ",len(codProyecto)))=1  " & _
                          "then 0 " & _
                     "end AS Cod  INTO tmp0 " & _
              "FROM GNProyecto " & _
              "WHERE (Substring(CodProyecto,1," & Len(pre) & ")='" & _
                        pre & "')  group by CodProyecto"
#End If
        VerificaExistenciaTabla 0
        gobjMain.EmpresaActual.EjecutarSQL sql, 0
        
        sql = "SELECT MAX(Cod) as CodMax FROM tmp0"
        
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        
        If Not rs.EOF Then
            If Not IsNull(rs!CodMax) Then num = rs!CodMax
        End If
        
        CodProyecto = pre & Format(num + 1, mProps.objEmpresa.GNOpcion.FormatoCodCC)
        rs.Close
        VerificaExistenciaTabla 0
    Case 2  'Con formato
#If DAOLIB Then
        sql = "SELECT Max(Right([CodProyecto],3)) AS CodMax " & _
              "FROM GNProyecto " & _
              "WHERE (Left$([CodProyecto]," & (Len(pre) + 8) & ")='" & _
                        pre & Format(Date, "yyyymmdd") & "')"
#Else
        sql = "SELECT Max(Right(CodProyecto,3)) AS CodMax " & _
              "FROM GNProyecto " & _
              "WHERE (Substring(CodProyecto,1," & (Len(pre) + 8) & ")='" & _
                        pre & Format(Date, "yyyymmdd") & "')"
#End If
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then
            If Not IsNull(rs!CodMax) Then num = Val(rs!CodMax)
        End If
        
        CodProyecto = pre & Format(Date, "yyyymmdd") & "-" & Format(num + 1, mProps.objEmpresa.GNOpcion.FormatoCodCC)
        rs.Close
Case 3  'Con el Codigo de la Sucursal jeaa 30/03/2008
    If Len(GetSetting(APPNAME, SECTION, "CodSucursal_Activa_" & mProps.objEmpresa.CodEmpresa, "")) > 0 Then
        pre = GetSetting(APPNAME, SECTION, "CodSucursal_Activa_" & mProps.objEmpresa.CodEmpresa, "")
    Else
        MsgBox "No existe Sucursal Activa para este equipo" & Chr(13) & "Revice la configuracion en Información de la Empresa"
        Exit Sub
    End If
    lon_pre = Len(pre) + 1
#If DAOLIB Then
        sql = "SELECT Max(CodProyecto)) AS CodMax " & _
              "FROM GNProyecto " & _
              "WHERE (Left$([CodProyecto]," & (Len(pre) + 8) & ")='" & _
                        pre & Format(Date, "yyyymmdd") & "')"
#Else
        sql = "SELECT case  " & _
                     "when isnumeric(substring(CodProyecto," & (lon_pre + 1) & ",len(codProyecto)))=1  " & _
                          "then substring(CodProyecto," & (lon_pre + 1) & ",len(codProyecto)) " & _
                     "when isnumeric(substring(CodProyecto," & (lon_pre + 1) & ",len(codProyecto)))=1  " & _
                          "then 0 " & _
                     "end AS Cod  INTO tmp0 " & _
              "FROM GNProyecto " & _
              "WHERE (Substring(CodProyecto,1," & Len(pre) & ")='" & _
                        pre & "')  group by CodProyecto"
#End If

        VerificaExistenciaTabla 0
        gobjMain.EmpresaActual.EjecutarSQL sql, 0
        
        sql = "SELECT MAX(Cod) as CodMax FROM tmp0"
        
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        
        If Not rs.EOF Then
            If Not IsNull(rs!CodMax) Then num = rs!CodMax
        End If


        CodProyecto = pre & "-" & Format(num + 1, mProps.objEmpresa.GNOpcion.FormatoCodCC)
        rs.Close
    End Select
    Set rs = Nothing
End Sub

Public Property Let Observacion(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "GNProyecto_Observacion", MSGERR_INVALIDO
    End If
    
    mProps.Observacion = value
    mbooModificado = True
End Property

Public Property Get Observacion() As String
    Observacion = Trim$(mProps.Observacion)
End Property

Public Property Let Cantidad(ByVal value As Currency)
    mProps.Cantidad = value
    mbooModificado = True
End Property

Public Property Get Cantidad() As Currency
    Cantidad = mProps.Cantidad
End Property




Public Property Let IdCanton(ByVal value As Long)
    Dim sql As String, rs As Recordset
  
    'Cuando cambia el valor
    If value <> mProps.IdCanton Then
        'Actualiza el Codigo también
        mProps.CodCanton = ""
        sql = "SELECT CodCanton FROM PCCanton WHERE IdCanton=" & value
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then mProps.CodCanton = rs!CodCanton
        rs.Close
        Set rs = Nothing
    End If
    
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
        If Len(value) > 0 Then
            'Actualiza IdCanton1 también
            sql = "SELECT IdCanton FROM PCCanton WHERE CodCanton" & "='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNCCanton_CodCanton", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCanton = rs.Fields("IdCanton")
            End If
            rs.Close
        Else
            IdCanton = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodCanton = value
    mbooModificado = True
End Property

Public Property Get CodCanton() As String
    CodCanton = Trim$(mProps.CodCanton)
End Property




Public Property Let IdInventario(ByVal value As Long)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.IdInventario Then
        'Actualiza el Codigo también
        mProps.CodInventario = ""
        sql = "SELECT CodInventario FROM IVInventario WHERE IdInventario=" & value
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then mProps.CodInventario = rs!CodInventario
        rs.Close
        Set rs = Nothing
    End If
    
    mProps.IdInventario = value
    mbooModificado = True
End Property

'*** MAKOTO 05/mar/01 Agregado
Public Property Get IdInventario() As Long
    IdInventario = mProps.IdInventario
End Property

Public Property Let CodInventario(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodInventario Then
        If Len(value) > 0 Then
            'Actualiza IdInventario también
            sql = "SELECT IdInventario FROM IVInventario WHERE CodInventario='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNProyecto.CodInventario", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdInventario = rs!IdInventario
            End If
            rs.Close
        Else
            mProps.IdInventario = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodInventario = value
    mbooModificado = True
End Property

Public Property Get CodInventario() As String
    CodInventario = Trim$(mProps.CodInventario)
End Property


Public Function GeneraCodigoxVendedor(ByVal codVende As String) As String
    Dim sql As String, num As Long, cod As String, rs As Recordset
    Dim pre As String, codg As String, v1 As Variant, v2 As Variant
    Dim Orden As String, size As String, mascara As String, i As Integer
    
    
    pre = Trim$(codVende)
    cod = Me.CodProyecto
    
    sql = "SELECT (MAX(SUBSTRING(CodProyecto," & Len(pre) + 1 & ",LEN(codProyecto)))) As CodMax " & _
          "FROM GNProyecto " & _
          "WHERE CodProyecto LIKE '" & pre & "[0-9][0-9]%'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        If Not IsNull(rs!CodMax) Then num = Val(rs!CodMax)
    End If
    rs.Close
    num = num + 1
    cod = pre & Format(num, "0000") 'cambiar por plantilla
    
    GeneraCodigoxVendedor = cod
    Set rs = Nothing
End Function



Public Property Let Estado(ByVal value As Integer)
    mProps.Estado = value
    mbooModificado = True
End Property

Public Property Get Estado() As Integer
    Estado = mProps.Estado
End Property

Public Property Let Nombre1(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "GNProyecto_Nombre1", MSGERR_INVALIDO
    End If
    
    mProps.Nombre1 = value
    mbooModificado = True
End Property

Public Property Get Nombre1() As String
    Nombre1 = Trim$(mProps.Nombre1)
End Property

Public Property Let Nombre2(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "GNProyecto_Nombre2", MSGERR_INVALIDO
    End If
    
    mProps.Nombre2 = value
    mbooModificado = True
End Property

Public Property Get Nombre2() As String
    Nombre2 = Trim$(mProps.Nombre2)
End Property

Public Property Let Nombre3(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "GNProyecto_Nombre3", MSGERR_INVALIDO
    End If
    
    mProps.Nombre3 = value
    mbooModificado = True
End Property

Public Property Get Nombre3() As String
    Nombre3 = Trim$(mProps.Nombre3)
End Property

Public Property Let Cargo1(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "GNProyecto_Cargo1", MSGERR_INVALIDO
    End If
    
    mProps.Cargo1 = value
    mbooModificado = True
End Property

Public Property Get Cargo1() As String
    Cargo1 = Trim$(mProps.Cargo1)
End Property

Public Property Let Cargo2(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "GNProyecto_Cargo2", MSGERR_INVALIDO
    End If
    
    mProps.Cargo2 = value
    mbooModificado = True
End Property

Public Property Get Cargo2() As String
    Cargo2 = Trim$(mProps.Cargo2)
End Property

Public Property Let Cargo3(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "GNProyecto_Cargo3", MSGERR_INVALIDO
    End If
    
    mProps.Cargo3 = value
    mbooModificado = True
End Property

Public Property Get Cargo3() As String
    Cargo3 = Trim$(mProps.Cargo3)
End Property

Public Property Let Telefono1(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "GNProyecto_Telefono1", MSGERR_INVALIDO
    End If
    
    mProps.Telefono1 = value
    mbooModificado = True
End Property

Public Property Get Telefono1() As String
    Telefono1 = Trim$(mProps.Telefono1)
End Property

Public Property Let Telefono2(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "GNProyecto_Telefono2", MSGERR_INVALIDO
    End If
    
    mProps.Telefono2 = value
    mbooModificado = True
End Property

Public Property Get Telefono2() As String
    Telefono2 = Trim$(mProps.Telefono2)
End Property

Public Property Let Telefono3(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "GNProyecto_Telefono3", MSGERR_INVALIDO
    End If
    
    mProps.Telefono3 = value
    mbooModificado = True
End Property

Public Property Get Telefono3() As String
    Telefono3 = Trim$(mProps.Telefono3)
End Property

Public Property Let Obser1(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "GNProyecto_Obser1", MSGERR_INVALIDO
    End If
    
    mProps.Obser1 = value
    mbooModificado = True
End Property

Public Property Get Obser1() As String
    Obser1 = Trim$(mProps.Obser1)
End Property

Public Property Let Obser2(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "GNProyecto_Obser2", MSGERR_INVALIDO
    End If
    
    mProps.Obser2 = value
    mbooModificado = True
End Property

Public Property Get Obser2() As String
    Obser2 = Trim$(mProps.Obser2)
End Property

Public Property Let Obser3(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "GNProyecto_Obser3", MSGERR_INVALIDO
    End If
    
    mProps.Obser3 = value
    mbooModificado = True
End Property

Public Property Get Obser3() As String
    Obser3 = Trim$(mProps.Obser3)
End Property

Public Property Let FechaProyeccion(ByVal value As Date)
    mProps.FechaProyeccion = value
    mbooModificado = True
End Property

Public Property Get FechaProyeccion() As Date
    FechaProyeccion = mProps.FechaProyeccion
End Property


Public Property Let EstadoNegociacion(ByVal value As Integer)
    mProps.EstadoNegociacion = value
    mbooModificado = True
End Property

Public Property Get EstadoNegociacion() As Integer
    EstadoNegociacion = mProps.EstadoNegociacion
End Property

Public Sub GrabaProyectoDetalle()
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As GnProyectoDetalle, sql As String, band As Boolean
    Dim NumHijos As Integer, CountNumHijos As Integer
    On Error GoTo CapturaError
    
    'Elimina  todos  los registros
    sql = "Delete Ivm " & _
          "From Ivinventario Iv " & _
          "INNER JOIN GNProyectoDetalle Ivm ON IV.IdInventario = Ivm.IdInventario "
        sql = sql & " Where Ivm.IdProyecto = " & mProps.IdProyecto

    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    band = True
    If Not mcolDetalleProyecto Is Nothing Then
        sql = "SELECT * FROM GNProyectoDetalle WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        NumHijos = mcolDetalleProyecto.Count
        CountNumHijos = 0
        For i = 1 To mcolDetalleProyecto.Count
            If CountNumHijos < NumHijos Then
                CountNumHijos = CountNumHijos + 1
                Set obj = mcolDetalleProyecto.item(i)
                If obj.IdInventario = 0 Then
                        obj.Recuperar obj.CodInventario
                        If verificacionDuiplicado(obj.IdInventario) Then
                            RemoveDetalleProyecto i
                            i = i - 1
                        End If
                End If
                If band Then
                    rs1.AddNew
                    rs1!IdInventario = obj.IdInventario
                    rs1!IdProyecto = Me.IdProyecto
                    rs1!Cantidad = obj.Cantidad
                    rs1!Orden = obj.Orden
                    rs1.Update
                End If
            End If
            band = True

        Next i
    End If
    Set rs1 = Nothing
    Set obj = Nothing
    Exit Sub
CapturaError:
    MsgBox Err.Description
    Set rs1 = Nothing
    Set obj = Nothing
    Exit Sub
End Sub


Public Function RecuperaDetalleProyecto(ix As Long) As GnProyectoDetalle
    Set RecuperaDetalleProyecto = mcolDetalleProyecto.item(ix)
End Function

Public Function AddDetalleProyecto() As Long
    Dim nuevo As GnProyectoDetalle
    
    If mcolDetalleProyecto Is Nothing Then CargaDetalleProyecto
    Set nuevo = New GnProyectoDetalle
    Set nuevo.GNProyecto = Me  'Me  olvide  del padre
    'Proporciona valores predeterminados
    nuevo.Cantidad = 1
    
    mcolDetalleProyecto.Add nuevo
    AddDetalleProyecto = mcolDetalleProyecto.Count    'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    
    mbooModificado = True
    
End Function

Public Sub RemoveDetalleProyecto(ByVal ix As Long, Optional obj As GnProyectoDetalle)
    If mcolDetalleProyecto Is Nothing Then CargaDetalleProyecto
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolDetalleProyecto.Count
            If mcolDetalleProyecto.item(ix) Is obj Then
                mcolDetalleProyecto.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolDetalleProyecto.Remove ix
    End If
    
    'mcolDetalleProyecto.Remove ix
    mbooModificado = True
End Sub

Private Sub CargaDetalleProyecto()
    Dim rs As Recordset, sql As String, obj As GnProyectoDetalle
    Dim i As Long
    Set mcolDetalleProyecto = New Collection
        sql = "Select IV.IdInventario, Iv.CodInventario, Iv.Descripcion, " & _
              " Ivm.Cantidad From  VWIvinventarioRECUPERAR IV " & _
              "INNER JOIN GNProyectoDetalle Ivm "
          sql = sql & " ON IV.IdInventario = Ivm.IdInventario "
            sql = sql & " Where Ivm.IdProyecto = " & mProps.IdProyecto

    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddDetalleProyecto
            Set obj = mcolDetalleProyecto.item(i)
                  obj.IdInventario = !IdInventario
                  obj.CodInventario = !CodInventario
                  obj.Descripcion = !Descripcion
                 obj.Cantidad = !Cantidad
            
           .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
    'mbooEsNuevo = False
    'mbooModificado = False
End Sub


Private Function verificacionDuiplicado(ByVal IdInventario As String) As Boolean
    Dim rs As Recordset, sql As String
    verificacionDuiplicado = False
    'Recupera los detalles de existencia
    sql = "SELECT IdInventario FROM GNProyectoDetalle where idinventario =" & IdInventario
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        verificacionDuiplicado = True
    End If
    Set rs = Nothing
End Function

Friend Property Get IdProyecto() As Long
    IdProyecto = mProps.IdProyecto
End Property

Public Property Get NumProyectoDetalle() As Long
    If mcolDetalleProyecto Is Nothing Then
        NumProyectoDetalle = 0
    Else
        NumProyectoDetalle = mcolDetalleProyecto.Count
    End If
    
End Property

Public Property Let Observacion2(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "GNProyecto_Observacion", MSGERR_INVALIDO
    End If
    
    mProps.Observacion2 = value
    mbooModificado = True
End Property

Public Property Get Observacion2() As String
    Observacion2 = Trim$(mProps.Observacion2)
End Property

Public Property Let Num1(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNProyecto_Observacion", MSGERR_INVALIDO
    End If
    
    mProps.Num1 = value
    mbooModificado = True
End Property

Public Property Get Num1() As String
    Num1 = Trim$(mProps.Num1)
End Property

Public Property Let Num2(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNProyecto_Observacion", MSGERR_INVALIDO
    End If
    
    mProps.Num2 = value
    mbooModificado = True
End Property

Public Property Get Num2() As String
    Num2 = Trim$(mProps.Num2)
End Property

Public Property Let IdCompetencia(ByVal value As Long)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.IdCompetencia Then
        'Actualiza el Codigo también
        mProps.CodCompetencia = ""
        sql = "SELECT CodCompetencia FROM gnCompetencia WHERE IdCompetencia=" & value
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then mProps.CodCompetencia = rs!CodCompetencia
        rs.Close
        Set rs = Nothing
    End If
    
    mProps.IdCompetencia = value
    mbooModificado = True
End Property

'*** MAKOTO 05/mar/01 Agregado
Public Property Get IdCompetencia() As Long
    IdCompetencia = mProps.IdCompetencia
End Property

Public Property Let CodCompetencia(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCompetencia Then
        If Len(value) > 0 Then
            'Actualiza IdCompetencia también
            sql = "SELECT IdCompetencia FROM gnCompetencia WHERE CodCompetencia='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNProyecto.CodCompetencia", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdCompetencia = rs!IdCompetencia
            End If
            rs.Close
        Else
            mProps.IdCompetencia = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodCompetencia = value
    mbooModificado = True
End Property

Public Property Get CodCompetencia() As String
    CodCompetencia = Trim$(mProps.CodCompetencia)
End Property

