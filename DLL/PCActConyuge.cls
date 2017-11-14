VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCActConyuge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type T_PROPS
    Id As Long
    IdProvCli As Long
    CodProvCli As String
    'SOLICITANTE
    ActividadSoli As String
    NombreTrabajoSoli As String
    DirTrabajoSoli As String
    AntiguedadSoli As String
    TelfActividadSol As String
    TipoLocalSoli As String
    NombreArrendSoli As String
    TelfArrendSoli As String
    IngresosSoli As Currency
    OtrosIngresosSoli As Currency
    GastosFamSoli As Currency
    GastosArriendoSoli As Currency
    OtrosGastosSoli As Currency
    'CONYUGUE
    
    NombreConyu As String
    CIConyu As String
    FechaNacimientoConyu As Date
    NumcargasConyu As Integer
    ActividadConyu As String
    NombreTrabajoConyu As String
    DirTrabajoConyu  As String
    AntiguedadConyu As String
    TelfActividadCounyu As String
    TipoLocalConyu As String
    NombreArrendConyu As String
    TelfArrendConyu As String
    IngresosConyu As Currency
    ValorVehiculo1Conyu As Currency
    OtrosIngresosConyu As Currency
    GastosFamConyu As Currency
    GastosArriendoConyu As Currency
    OtrosGastosConyu As Currency
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mProps As T_PROPS
Private mcolContactos As Collection
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Public Sub Grabar()
    Dim sql As String, fgrabado As Date, i As Long
    Dim rs As Recordset
    Dim rsSuc As Recordset, gnSuc As GNSucursal
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    On Error GoTo ErrTrap
#If DAOLIB Then
    gobjWS.BeginTrans                           '======= Comienza la transacción
#Else
    mProps.objEmpresa.Coneccion.BeginTrans      '======= Comienza la transacción
#End If
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM PCActConyuge WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
    
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM PCActConyuge WHERE idProvCli=" & mProps.IdProvCli
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
'        If rs!FechaGrabado <> mProps.FechaGrabado Then
'            Err.Raise ERR_MODIFICADO, "PCActConyuge_Grabar", MSGERR_MODIFICADO
'        End If
    End If

    'Graba el registro en la base
    With rs
        !IdProvCli = mProps.IdProvCli
        'datos soli
        !Actividad = mProps.ActividadSoli
        !NombreTrabajo = mProps.NombreTrabajoSoli
        !DirTrabajo = mProps.DirTrabajoSoli
        !Antiguedad = mProps.AntiguedadSoli
        !TelfActividad = mProps.TelfActividadSol
        '!TelfArrend = mProps.TelfArrend
        !TipoLocal = mProps.TipoLocalSoli
        !NombreArrend = mProps.NombreArrendSoli
        !TelfArrend = mProps.TelfArrendSoli
        !Ingresos = mProps.IngresosSoli
        !OtrosIngresos = mProps.OtrosIngresosSoli
        !GastosFam = mProps.GastosFamSoli
        !GastosArriendo = mProps.GastosArriendoSoli
        !OtrosGastos = mProps.OtrosGastosSoli
        'conyu
        !NombreConyu = mProps.NombreConyu
        !CIConyu = mProps.CIConyu
        !FechaNacimientoConyu = mProps.FechaNacimientoConyu
        !NumcargasConyu = mProps.NumcargasConyu
        !ActividadConyu = mProps.ActividadConyu
        !NombreTrabajoConyu = mProps.NombreTrabajoConyu
        !DirTrabajoConyu = mProps.DirTrabajoConyu
        !AntiguedadConyu = mProps.AntiguedadConyu
        !TelfActividadCounyu = mProps.TelfActividadCounyu
        !TipoLocalConyu = mProps.TipoLocalConyu
        !NombreArrendConyu = mProps.NombreArrendConyu
        !TelfArrendConyu = mProps.TelfArrendConyu
        !IngresosConyu = mProps.IngresosConyu
        !OtrosIngresosConyu = mProps.OtrosIngresosConyu
        !GastosFamConyu = mProps.GastosFamConyu
        !GastosArriendoConyu = mProps.GastosArriendoConyu
        !OtrosGastosConyu = mProps.OtrosGastosConyu
        
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
'        mProps.Id = !Id             'Guarda el Id(Autonumerico)

        mProps.IdProvCli = !IdProvCli
        .Close
    End With
    
#If DAOLIB Then
    gobjWS.CommitTrans                  '======== Completa la transacción
#Else
    mProps.objEmpresa.Coneccion.CommitTrans      '======== Completa la transacción
#End If
    'mProps.FechaGrabado = fgrabado
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
'    If mbooEsNuevo Then
'        mProps.objEmpresa.GrabaGNLogAccion "NUE_PCAc", "Solicitud Credito Prov/Cli Nuevo: " & _
'                                                         mProps.CodProvCli, "DT"
'    Else
'        mProps.objEmpresa.GrabaGNLogAccion "MOD_PC", "Solicitud Credito Prov/Cli Modificado: " & _
'                                                         mProps.CodProvCli, "DT"
'    End If
    

'    mProps.objEmpresa.RaiseRefreshEvent REFRESH_SOLPROVCLI
    

    mbooEsNuevo = False
    mbooModificado = False
    
    Exit Sub
ErrTrap:
#If DAOLIB Then
    gobjWS.Rollback                   '======== Anula la transacción
#Else
    mProps.objEmpresa.Coneccion.RollbackTrans  '======== Anula la transacción
#End If
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Friend Sub Recuperar(key As Long)
    Dim sql As String, i As Long
    Dim rs As Recordset
    
    sql = "SELECT * FROM PCActConyuge WHERE idprovcli=" & key
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    
    'Recupera los datos de Usuario
    With rs
        If rs.RecordCount > 0 Then
            If Not IsNull(!IdProvCli) Then mProps.IdProvCli = !IdProvCli
            'If Not IsNull(!CodProvCli) Then mProps.CodProvCli = !CodProvCli
            'solicitante
            If Not IsNull(!Actividad) Then mProps.ActividadSoli = !Actividad
            If Not IsNull(!NombreTrabajo) Then mProps.NombreTrabajoSoli = !NombreTrabajo
            If Not IsNull(!DirTrabajo) Then mProps.DirTrabajoSoli = !DirTrabajo
            If Not IsNull(!Antiguedad) Then mProps.AntiguedadSoli = !Antiguedad
            If Not IsNull(!TelfActividad) Then mProps.TelfActividadSol = !TelfActividad
            If Not IsNull(!TipoLocal) Then mProps.TipoLocalSoli = !TipoLocal
            If Not IsNull(!NombreArrend) Then mProps.NombreArrendSoli = !NombreArrend
            If Not IsNull(!TelfArrend) Then mProps.TelfArrendSoli = !TelfArrend
            If Not IsNull(!Ingresos) Then mProps.IngresosSoli = !Ingresos
            If Not IsNull(!OtrosIngresos) Then mProps.OtrosIngresosSoli = !OtrosIngresos
            If Not IsNull(!GastosFam) Then mProps.GastosFamSoli = !GastosFam
            If Not IsNull(!GastosArriendo) Then mProps.GastosArriendoSoli = !GastosArriendo
            If Not IsNull(!OtrosGastos) Then mProps.OtrosGastosSoli = !OtrosGastos
            'conyuge
            If Not IsNull(!NombreConyu) Then mProps.NombreConyu = !NombreConyu
            If Not IsNull(!CIConyu) Then mProps.CIConyu = !CIConyu
            If Not IsNull(!FechaNacimientoConyu) Then mProps.FechaNacimientoConyu = !FechaNacimientoConyu
            If Not IsNull(!NumcargasConyu) Then mProps.NumcargasConyu = !NumcargasConyu
            If Not IsNull(!ActividadConyu) Then mProps.ActividadConyu = !ActividadConyu
            If Not IsNull(!NombreTrabajoConyu) Then mProps.NombreTrabajoConyu = !NombreTrabajoConyu
            If Not IsNull(!DirTrabajoConyu) Then mProps.DirTrabajoConyu = !DirTrabajoConyu
            If Not IsNull(!AntiguedadConyu) Then mProps.AntiguedadConyu = !AntiguedadConyu
            If Not IsNull(!TelfActividadCounyu) Then mProps.TelfActividadCounyu = !TelfActividadCounyu
            If Not IsNull(!TipoLocalConyu) Then mProps.TipoLocalConyu = !TipoLocalConyu
            If Not IsNull(!NombreArrendConyu) Then mProps.NombreArrendConyu = !NombreArrendConyu
            If Not IsNull(!TelfArrendConyu) Then mProps.TelfArrendConyu = !TelfArrendConyu
            If Not IsNull(!IngresosConyu) Then mProps.IngresosConyu = !IngresosConyu
            If Not IsNull(!OtrosIngresosConyu) Then mProps.OtrosIngresosConyu = !OtrosIngresosConyu
            If Not IsNull(!GastosFamConyu) Then mProps.GastosFamConyu = !GastosFamConyu
            If Not IsNull(!GastosArriendoConyu) Then mProps.GastosArriendoConyu = !GastosArriendoConyu
            If Not IsNull(!OtrosGastosConyu) Then mProps.OtrosGastosConyu = !OtrosGastosConyu
            .Close
        End If
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    'Verifica si existe Kardex
    sql = "select g.Idtransfuente FROM PCActConyuge p  inner join gncomprobante g on g.transid = p.idtransfuente  WHERE Id = " & mProps.Id
    sql = sql & " and g.idtransfuente <> 0"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "PCSolCreProvCli.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing

    'Verifica si no esta asignado como garante
      
    'Elimina el registro
    'sql = "DELETE FROM PCSolCreProvCli WHERE IdProvCli = " & mProps.IdProvCli
    'mProps.objEmpresa.Execute sql, True

    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    'mProps.objEmpresa.GrabaGNLogAccion "ELI_PC", "Proveedor/Cliente Eliminado: " & _
                                                         mProps.CodProvCli, "DT"
                                                         
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PROVCLI
End Sub

'Public Function Clone() As PCSolCreProvCli
'    Dim nuevo As PCSolCreProvCli, nc As PCContacto, i As Long
'
'    'Copia las propiedades al nuevo objeto
'    Set nuevo = New PCSolCreProvCli
'    nuevo.Props = mProps
'
'
'    Set Clone = nuevo
'    Set nuevo = Nothing
'    Set nc = Nothing
'End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.Id = 0
    mProps.IdProvCli = 0
    mbooModificado = True
    
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Sub SetModificado()
    mbooModificado = True
End Sub

Public Property Get Id() As Long
    Id = mProps.Id
End Property

Public Property Get IdProvCli() As Long
    IdProvCli = mProps.IdProvCli
End Property

Public Property Let IdProvCli(ByVal value As Long)
    mProps.IdProvCli = Trim(value)
    mbooModificado = True
End Property


Public Property Let CodProvCli(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodProvCli Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaContable también
            sql = "SELECT IdProvCli FROM PcProvCli WHERE CodProvcli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli.CodProvCli", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdProvCli = rs!IdProvCli
            End If
            rs.Close
            Set rs = Nothing
        Else
            mProps.IdProvCli = 0
        End If
    End If
    
    
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_CodProvCli", MSGERR_INVALIDO
    End If
    mProps.CodProvCli = Trim(value)
    mbooModificado = True
End Property


Public Property Get CodProvCli() As String
    CodProvCli = Trim$(mProps.CodProvCli)
End Property



'solicitante
Public Property Let NombreTrabajoSoli(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_NombreTrabajoSoli", MSGERR_INVALIDO
    End If
    
    mProps.NombreTrabajoSoli = value
    mbooModificado = True
End Property

Public Property Get NombreTrabajoSoli() As String
    NombreTrabajoSoli = Trim$(mProps.NombreTrabajoSoli)
End Property

Public Property Let DirTrabajoSoli(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_DirTrabajoSoli", MSGERR_INVALIDO
    End If
    mProps.DirTrabajoSoli = value
    mbooModificado = True
End Property

Public Property Get DirTrabajoSoli() As String
    DirTrabajoSoli = Trim$(mProps.DirTrabajoSoli)
End Property

Public Property Let AntiguedadSoli(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_AntiguedadSoli", MSGERR_INVALIDO
    End If
    mProps.AntiguedadSoli = value
    mbooModificado = True
End Property

Public Property Get AntiguedadSoli() As String
    AntiguedadSoli = Trim$(mProps.AntiguedadSoli)
End Property

Public Property Let TelfActividadSol(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_TelfActividadSol", MSGERR_INVALIDO
    End If
    
    mProps.TelfActividadSol = value
    mbooModificado = True
End Property

Public Property Get TelfActividadSol() As String
    TelfActividadSol = Trim$(mProps.TelfActividadSol)
End Property


Public Property Let TipoLocalSoli(ByVal value As String)
    If Len(value) > 1 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_TipoLocalSoli", MSGERR_INVALIDO
    End If
    
    mProps.TipoLocalSoli = value
    mbooModificado = True
End Property

Public Property Get TipoLocalSoli() As String
    TipoLocalSoli = Trim$(mProps.TipoLocalSoli)
End Property

Public Property Let NombreArrendSoli(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_NombreArrendSoli", MSGERR_INVALIDO
    End If
    
    mProps.NombreArrendSoli = value
    mbooModificado = True
End Property

Public Property Get NombreArrendSoli() As String
    NombreArrendSoli = Trim$(mProps.NombreArrendSoli)
End Property

Public Property Let TelfArrendSoli(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_TelfArrendSoli", MSGERR_INVALIDO
    End If
    
    mProps.TelfArrendSoli = value
    mbooModificado = True
End Property

Public Property Get TelfArrendSoli() As String
    TelfArrendSoli = Trim$(mProps.TelfArrendSoli)
End Property

Public Property Get IngresosSoli() As Currency
    IngresosSoli = mProps.IngresosSoli
End Property

Public Property Let IngresosSoli(value As Currency)
    mProps.IngresosSoli = value
    mbooModificado = True
End Property

Public Property Get OtrosIngresosSoli() As Currency
    OtrosIngresosSoli = mProps.OtrosIngresosSoli
End Property

Public Property Let OtrosIngresosSoli(value As Currency)
    mProps.OtrosIngresosSoli = value
    mbooModificado = True
End Property
'
Public Property Get GastosFamSoli() As Currency
    GastosFamSoli = mProps.GastosFamSoli
End Property

Public Property Let GastosFamSoli(value As Currency)
    mProps.GastosFamSoli = value
    mbooModificado = True
End Property
'
Public Property Get GastosArriendoSoli() As Currency
    GastosArriendoSoli = mProps.GastosArriendoSoli
End Property

Public Property Let GastosArriendoSoli(value As Currency)
    mProps.GastosArriendoSoli = value
    mbooModificado = True
End Property

Public Property Get OtrosGastosSoli() As Currency
    OtrosGastosSoli = mProps.OtrosGastosSoli
End Property

Public Property Let OtrosGastosSoli(value As Currency)
    mProps.OtrosGastosSoli = value
    mbooModificado = True
End Property

'conyuge

Public Property Let NombreConyu(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_NombreConyu", MSGERR_INVALIDO
    End If
    mProps.NombreConyu = value
    mbooModificado = True
End Property

Public Property Get NombreConyu() As String
    NombreConyu = mProps.NombreConyu
End Property

Public Property Let CIConyu(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_CIConyu", MSGERR_INVALIDO
    End If
    
    mProps.CIConyu = value
    mbooModificado = True
End Property

Public Property Get CIConyu() As String
    CIConyu = Trim$(mProps.CIConyu)
End Property

Public Property Let NombreTrabajoConyu(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_NombreTrabajoConyu", MSGERR_INVALIDO
    End If
    
    mProps.NombreTrabajoConyu = value
    mbooModificado = True
End Property

Public Property Get NombreTrabajoConyu() As String
    NombreTrabajoConyu = Trim$(mProps.NombreTrabajoConyu)
End Property

Public Property Let DirTrabajoConyu(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_DirTrabajoConyu", MSGERR_INVALIDO
    End If
    mProps.DirTrabajoConyu = value
    mbooModificado = True
End Property

Public Property Get DirTrabajoConyu() As String
    DirTrabajoConyu = Trim$(mProps.DirTrabajoConyu)
End Property

Public Property Let AntiguedadConyu(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_AntiguedadConyu", MSGERR_INVALIDO
    End If
    mProps.AntiguedadConyu = value
    mbooModificado = True
End Property

Public Property Get AntiguedadConyu() As String
    AntiguedadConyu = Trim$(mProps.AntiguedadConyu)
End Property

Public Property Let TelfActividadCounyu(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_TelfActividadCounyu", MSGERR_INVALIDO
    End If
    mProps.TelfActividadCounyu = value
    mbooModificado = True
End Property

Public Property Get TelfActividadCounyu() As String
    TelfActividadCounyu = Trim$(mProps.TelfActividadCounyu)
End Property

Public Property Let TipoLocalConyu(ByVal value As String)
    If Len(value) > 1 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_TipoLocalConyu", MSGERR_INVALIDO
    End If
    mProps.TipoLocalConyu = value
    mbooModificado = True
End Property

Public Property Get TipoLocalConyu() As String
    TipoLocalConyu = Trim$(mProps.TipoLocalConyu)
End Property


Public Property Let NombreArrendConyu(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_NombreArrendConyu", MSGERR_INVALIDO
    End If
    mProps.NombreArrendConyu = value
    mbooModificado = True
End Property

Public Property Get NombreArrendConyu() As String
    NombreArrendConyu = Trim$(mProps.NombreArrendConyu)
End Property

Public Property Let TelfArrendConyu(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_TelfArrendConyu", MSGERR_INVALIDO
    End If
    mProps.TelfArrendConyu = value
    mbooModificado = True
End Property

Public Property Get TelfArrendConyu() As String
    TelfArrendConyu = Trim$(mProps.TelfArrendConyu)
End Property

Public Property Get IngresosConyu() As Currency
    IngresosConyu = mProps.IngresosConyu
End Property

Public Property Let IngresosConyu(value As Currency)
    mProps.IngresosConyu = value
    mbooModificado = True
End Property

Public Property Get OtrosIngresosConyu() As Currency
    OtrosIngresosConyu = mProps.OtrosIngresosConyu
End Property

Public Property Let OtrosIngresosConyu(value As Currency)
    mProps.OtrosIngresosConyu = value
    mbooModificado = True
End Property
'
Public Property Get GastosFamConyu() As Currency
    GastosFamConyu = mProps.GastosFamConyu
End Property

Public Property Let GastosFamConyu(value As Currency)
    mProps.GastosFamConyu = value
    mbooModificado = True
End Property
'
Public Property Get GastosArriendoConyu() As Currency
    GastosArriendoConyu = mProps.GastosArriendoConyu
End Property

Public Property Let GastosArriendoConyu(value As Currency)
    mProps.GastosArriendoConyu = value
    mbooModificado = True
End Property

Public Property Get OtrosGastosConyu() As Currency
    OtrosGastosConyu = mProps.OtrosGastosConyu
End Property

Public Property Let OtrosGastosConyu(value As Currency)
    mProps.OtrosGastosConyu = value
    mbooModificado = True
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
    
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Sub GrabarEnSucursal(ByVal bd As String, ByVal servidor As String)
'    Dim sql As String, fgrabado As Date, i As Long
'    Dim rs As Recordset
'
'
'    If mbooEsNuevo Then
'        'Verifica si no repite el código
'        sql = "SELECT CodProvCli FROM " & bd & ".dbo.PCActConyuge " & _
'              "WHERE (CodProvCli='" & mProps.CodProvCli & "') AND " & _
'                    "(IdProvCli<>" & mProps.IdProvCli & ")"
'
'        Set rs = mProps.objEmpresa.OpenRecordsetSUC(sql, bd, servidor)
'        If Not rs.EOF Then      'Si ya existe el codigo
'            Err.Raise ERR_REPITECODIGO, "PCSolCreProvCli", MSGERR_REPITECODIGO
'        End If
'        rs.Close
'    End If
'
'    On Error GoTo ErrTrap
'
'''#If DAOLIB Then
'''    gobjWS.BeginTrans                           '======= Comienza la transacción
'''#Else
'''    mProps.objEmpresa.Coneccion.BeginTrans      '======= Comienza la transacción
'''#End If
'    'Si es nuevo
'    If mbooEsNuevo Then
'        sql = "SELECT * FROM " & bd & ".dbo.PCSolCreProvCli WHERE 1=0"
'                Set rs = mProps.objEmpresa.OpenRecordsetSUC(sql, bd, servidor)
'        rs.AddNew
'
'    'Si es modificada
'    Else
'        'Recupera el registro correspondiente para modificar
'        sql = "SELECT * FROM " & bd & ".dbo.PCSolCreProvCli WHERE CodProvCli='" & mProps.CodProvCli & "'"
'        Set rs = mProps.objEmpresa.OpenRecordsetParaEditSuc(sql, bd, servidor)
'        If rs.RecordCount = "0" Then
'            sql = "SELECT * FROM " & bd & ".dbo.PCSolCreProvCli WHERE 1=0"
'            Set rs = mProps.objEmpresa.OpenRecordsetParaEditSuc(sql, bd, servidor)
'            rs.AddNew
'
'        End If
'
'
'#If DAOLIB Then
'        rs.Edit
'#End If
'
'    End If
'
'    'Graba el registro en la base
'    With rs
'        !CodProvCli = mProps.CodProvCli
'        !NombreArrend = mProps.NombreArrend
'        !BandPatriFam = mProps.BandPatriFam
'        !BandSepBienes = mProps.BandSepBienes
'        !NombreTrabajoSoli = mProps.NombreTrabajoSoli
'        !DirTrabajoSoli = mProps.DirTrabajoSoli
'        !TipoVivienda = mProps.TipoVivienda
'        !AnioResidencia = mProps.AnioResidencia
'        !AntiguedadSoli = mProps.AntiguedadSoli
'        !TelfActividadSol = mProps.TelfActividadSol
'        !TelfArrend = mProps.TelfArrend
'        !TipoLocalSoli = mProps.TipoLocalSoli
'        !NombreArrendSoli = mProps.NombreArrendSoli
'        !TelfArrendSoli = mProps.TelfArrendSoli
'        !NombreConyu = mProps.NombreConyu
'        !CIConyu = mProps.CIConyu
'
'        !IngresosSoli = mProps.IngresosSoli
'        !NombreTrabajoConyu = mProps.NombreTrabajoConyu
'        !DirTrabajoConyu = mProps.DirTrabajoConyu
'        !AntiguedadConyu = mProps.AntiguedadConyu
'        !TelfActividadCounyu = mProps.TelfActividadCounyu
'        !TipoLocalConyu = mProps.TipoLocalConyu
'        !MarcaPlaca1 = mProps.MarcaPlaca1           'Agregado Alex: Sept/2002
'        !NombreFami = mProps.NombreFami       'Agregado Alex: Sept/2002
'        !RelacionFami = mProps.RelacionFami                   'Agregado Alex: Sept/2002
'        !TelfFami = mProps.TelfFami           '***Agregado. 05/sep/2003. Angel
'        !ValorInmueble1 = mProps.ValorInmueble1                   '***Agregado. 08/sep/2003. Angel
'        !ValorVehiculo1 = mProps.ValorVehiculo1                 '***Agregado. 08/sep/2003. Angel
'        !NombreComer1 = mProps.NombreComer1  'Agregado AUC 31/08/2005
'        !DirComer1 = mProps.DirComer1       ' jeaa 17/12/2005
'        !TelfComer1 = mProps.TelfComer1  'Agregado jeaa 12/08/2006
'        !BandHipoteca1 = mProps.BandHipoteca1   'jeaa 16/01/2008 cambio por EQUIDAD TRIBUTARIA
'        !BandPrendado1 = mProps.BandPrendado1
'        !FechaGrabado = Now
'        .Update
'#If DAOLIB Then
'        rs.Bookmark = rs.LastModified
'#End If
'        .Move 0             'Para actualizar
'        fgrabado = !FechaGrabado
'''        mProps.IdProvCli = !IdProvCli
'        .Close
'    End With
'
'    'Graba todos los PCContacto
'''''    For i = 1 To CountContacto
'''''        Contactos(i).Grabar
'''''    Next i
'
'    mProps.FechaGrabado = fgrabado
'
'    'Graba un registro para auditoria   *** ANGEL 20/mar/02
'
'
'
'
'    Set rs = Nothing
'    Exit Sub
'ErrTrap:
'#If DAOLIB Then
'    gobjWS.Rollback                   '======== Anula la transacción
'#Else
'    mProps.objEmpresa.Coneccion.RollbackTrans  '======== Anula la transacción
'#End If
'    Err.Raise Err.Number, Err.Source, Err.Description
'    Exit Sub
End Sub

Public Property Let FechaNacimientoConyu(value As Date)
    If Not (IsDate(value)) Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_FechaNacimientoConyu", MSGERR_INVALIDO
    End If
    mProps.FechaNacimientoConyu = value
    mbooModificado = True
End Property

Public Property Get FechaNacimientoConyu() As Date
    FechaNacimientoConyu = mProps.FechaNacimientoConyu
End Property

Public Property Let ActividadSoli(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_ActividadSoli", MSGERR_INVALIDO
    End If
    mProps.ActividadSoli = value
    mbooModificado = True
End Property

Public Property Get ActividadSoli() As String
    ActividadSoli = Trim$(mProps.ActividadSoli)
End Property

Public Property Let ActividadConyu(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PCSolCreProvCli_ActividadConyu", MSGERR_INVALIDO
    End If
    mProps.ActividadConyu = value
    mbooModificado = True
End Property

Public Property Get ActividadConyu() As String
    ActividadConyu = Trim$(mProps.ActividadConyu)
End Property

Public Property Let NumcargasConyu(ByVal value As Long)
    mProps.NumcargasConyu = value
    mbooModificado = True
End Property

Public Property Get NumcargasConyu() As Long
    NumcargasConyu = Trim$(mProps.NumcargasConyu)
End Property
