VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Grupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Private Type T_PROPS
    CodGrupo As String
    Descripcion As String
    BandCaja As Boolean
    BandVendedor As Boolean 'AUC 26/11/2005
    BandBodeguero As Boolean 'jeaa 09/08/2006
    FechaGrabado As Date
    BandPrincipal As Boolean
    BandModificar As Boolean
    Opcion As String
End Type

Private mProps As T_PROPS
Private mcolPermisos As Collection
Private mobjPermisoActual As Permiso

Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Event Procesando(ByVal msg As String)           '*** MAKOTO 27/dic/00 Agregado

Public Sub Grabar()
    Dim sql As String, i As Long, fgrabado As Date
    Dim rs As Recordset
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    On Error GoTo ErrTrap
    
#If DAOLIB Then
    gobjWS.BeginTrans                   '======= Comienza la transacción
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    gobjMain.Coneccion.BeginTrans       '======= Comienza la transacción
#End If

    '*** MAKOTO 27/dic/00 Agregado
    'Avisa al programa UI
    RaiseEvent Procesando("Está grabándo datos de grupo...")

    If mbooEsNuevo Then
        'Verifica si no repite el código
        sql = "SELECT CodGrupo FROM Grupo WHERE CodGrupo='" & CodGrupo & "'"
#If DAOLIB Then
        Set rs = gobjMain.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
        rs.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly
#End If
        If Not rs.EOF Then      'Si ya existe el codigo
            Err.Raise ERR_REPITECODIGO, "Grupo", MSGERR_REPITECODIGO
        End If
        rs.Close
        
        sql = "SELECT * FROM Grupo WHERE 1=0"
#If DAOLIB Then
        Set rs = gobjMain.Database.OpenRecordset(sql)
#Else
        rs.Open sql, gobjMain.Coneccion, adOpenKeyset, adLockPessimistic
#End If
        rs.AddNew
        
    'Cuando es Modificación
    Else
        'Elimina todos los permisos que pertenece en el grupo
        sql = "DELETE FROM Permiso WHERE CodGrupo='" & CodGrupo & "'"
#If DAOLIB Then
        gobjMain.Database.Execute sql, dbFailOnError
#Else
        gobjMain.Coneccion.Execute sql
#End If
    
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM Grupo WHERE CodGrupo='" & CodGrupo & "'"
#If DAOLIB Then
        Set rs = gobjMain.Database.OpenRecordset(sql)
        rs.Edit
#Else
        rs.Open sql, gobjMain.Coneccion, adOpenKeyset, adLockPessimistic
#End If
        
        'Verifica la FechaGrabado para asegurar que otros Grupos no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "Grupo_Grabar", MSGERR_MODIFICADO
        End If
    End If


    'Graba el Grupo
    With rs
        !CodGrupo = CodGrupo
        !Descripcion = Descripcion
        !Opcion = Opcion
        !BandCaja = BandCaja
        !BandVendedor = BandVendedor 'AUC 26/11/2005
        !BandBodeguero = BandBodeguero 'jeaa 09/08/2006
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0
        fgrabado = !FechaGrabado
        .Close
    End With
    
    
    'Graba todos los permisos
    For i = 1 To CountPermiso
        Permisos(i).Grabar
    Next i
    
#If DAOLIB Then
    gobjWS.CommitTrans                  '======== Completa la transacción
#Else
    gobjMain.Coneccion.CommitTrans      '======== Completa la transacción
#End If
    mProps.FechaGrabado = fgrabado
    mbooEsNuevo = False
    mbooModificado = False
    
    '*** MAKOTO 27/dic/00 Agregado
    RaiseEvent Procesando("")       'Avisa al programa UI que terminó el proceso
    Exit Sub
ErrTrap:
#If DAOLIB Then
    gobjWS.Rollback                   '======== Anula la transacción
#Else
    gobjMain.Coneccion.RollbackTrans  '======== Anula la transacción
#End If
    '*** MAKOTO 27/dic/00 Agregado
    RaiseEvent Procesando("")       'Avisa al programa UI que terminó el proceso
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

'*** MAKOTO 27/dic/00 Agregado para llamar de los objetos hijos
Friend Sub GenerarProcesando(ByVal msg As String)
    RaiseEvent Procesando(msg)
End Sub

Friend Sub Recuperar()
    Dim sql As String, i As Long
    Dim rs As Recordset
    
    sql = "SELECT * FROM Grupo WHERE CodGrupo='" & CodGrupo & "'"
#If DAOLIB Then
    Set rs = gobjMain.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly
#End If
    'Recupera los datos de Grupo
    With rs
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!BandCaja) Then mProps.BandCaja = !BandCaja  'jeaa 04/08/2005
        If Not IsNull(!BandCaja) Then mProps.BandVendedor = !BandVendedor 'AUC 26/11/05
        If Not IsNull(!BandBodeguero) Then mProps.BandBodeguero = !BandBodeguero 'jeaa 09/08/2006
        If Not IsNull(!Opcion) Then mProps.Opcion = !Opcion
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
'    'Recupera los permisos
'    sql = "SELECT CodEmpresa FROM Permiso WHERE CodGrupo='" & CodGrupo & _
'                "' ORDER BY CodEmpresa"
'#If DAOLIB Then
'    Set rs = gobjMain.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
'#Else
'    rs.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly
'#End If
'    With rs
'        Do Until .EOF
'            i = AddPermiso
'            Permisos(i).CodEmpresa = !CodEmpresa
'            Permisos(i).Recuperar
'            .MoveNext
'        Loop
'        .Close
'    End With
        
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String
        Dim rs As Recordset

#If DAOLIB Then
    gobjWS.BeginTrans                   '======= Comienza la transacción
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    gobjMain.Coneccion.BeginTrans       '======= Comienza la transacción
#End If


'#If DAOLIB Then
'    gobjWS.BeginTrans
'#Else
'    gobjMain.EmpresaActual.Coneccion.BeginTrans
'#End If

On Error GoTo ErrTrap
    'Si no tiene derecho de supervisor, no puede hacer
    If Not gobjUsuarioActual.BandSupervisor Then
        Err.Raise ERR_NODERECHO, "Grupo_Eliminar", MSGERR_NODERECHO
    End If
    
    'Elimina los permisos
    sql = "DELETE FROM Permiso WHERE CodGrupo='" & CodGrupo & "'"
#If DAOLIB Then
    gobjMain.Database.Execute sql, dbFailOnError
#Else
    gobjMain.Coneccion.Execute sql
#End If
        
    'Elimina el Grupo
    sql = "DELETE FROM Grupo WHERE CodGrupo='" & CodGrupo & "'"
#If DAOLIB Then
    gobjMain.Database.Execute sql, dbFailOnError
#Else
    gobjMain.Coneccion.Execute sql
#End If

    mbooEsNuevo = True
    mbooModificado = True

#If DAOLIB Then
    gobjWS.CommitTrans                  '======== Completa la transacción
#Else
    gobjMain.Coneccion.CommitTrans      '======== Completa la transacción
#End If
Exit Sub

'#If DAOLIB Then
'    gobjWS.CommitTrans
'#Else
'    gobjMain.EmpresaActual.Coneccion.CommitTrans
'#End If
    
ErrTrap:
'#If DAOLIB Then
'    gobjWS.Rollback
'#Else
'    gobjMain.EmpresaActual.Coneccion.RollbackTrans
'#End If
#If DAOLIB Then
    gobjWS.CommitTrans                  '======== Completa la transacción
#Else
    gobjMain.Coneccion.RollbackTrans '======== Completa la transacción
#End If
    
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Public Function Clone() As Grupo
    Dim nuevo As Grupo, i As Long, np As Permiso
    
    If mcolPermisos Is Nothing Then RecuperaPermisos
    
    'Copia las propiedades de Grupo al nuevo objeto
    Set nuevo = New Grupo
    nuevo.Props = mProps
    
    'Copia los permisos al nuevo objeto
    For i = 1 To mcolPermisos.Count
        Set np = Permisos(i).Clone
        nuevo.AddPermisoParaClone np
    Next i
    
    Set Clone = nuevo
    Set nuevo = Nothing
    Set np = Nothing
End Function

Friend Sub AddPermisoParaClone(objPermiso As Permiso)
'    If mcolPermisos Is Nothing Then RecuperaPermisos
    If mcolPermisos Is Nothing Then Set mcolPermisos = New Collection
    
    Set objPermiso.Grupo = Me
    mcolPermisos.Add objPermiso
End Sub

Public Function ListaPermisos() As Variant
    Dim sql As String, rs As Recordset

    sql = "SELECT Id,CodEmpresa FROM Permiso WHERE CodGrupo='" & mProps.CodGrupo & "' ORDER BY CodEmpresa"

#If DAOLIB Then
    Set rs = gobjMain.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly
#End If
    Set ListaPermisos = rs
    Set rs = Nothing
End Function



Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mbooModificado = True
End Property

Public Function AddPermiso() As Long
    Dim nuevo As Permiso
    
    If mcolPermisos Is Nothing Then RecuperaPermisos
    
    Set nuevo = New Permiso
    Set nuevo.Grupo = Me
    mcolPermisos.Add nuevo
    AddPermiso = mcolPermisos.Count     'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function

Public Sub RemovePermiso(ix As Long)
    If mcolPermisos Is Nothing Then RecuperaPermisos
    
    mcolPermisos.Remove ix
    mbooModificado = True
End Sub

'Duplica un permiso para el mismo Grupo
Public Function CopiaPermiso(ix As Long) As Long
    Dim nuevo As Permiso, obj As Permiso
    
    If mcolPermisos Is Nothing Then RecuperaPermisos
    
    Set obj = mcolPermisos.item(ix)
    Set nuevo = obj.Clone
    mcolPermisos.Add nuevo
    
    
    CopiaPermiso = mcolPermisos.Count       'Devuelve el indice del nuevo elemento
    Set nuevo = Nothing
    Set obj = Nothing
    mbooModificado = True
End Function

Public Property Get CountPermiso() As Long
    If mcolPermisos Is Nothing Then RecuperaPermisos
    CountPermiso = mcolPermisos.Count
End Property

Public Property Get Permisos(ix As Long) As Permiso
    If mcolPermisos Is Nothing Then RecuperaPermisos
    Set Permisos = mcolPermisos.item(ix)
End Property

Public Property Get PermisoActual() As Permiso
    Dim mod_antes As Boolean, codemp As String
    
    mod_antes = mbooModificado      'Guarda el estado de Modificado
    codemp = gobjMain.EmpresaActual.CodEmpresa
    
    'Si es primera vez, crea una instancia del Permiso
    If mobjPermisoActual Is Nothing Then
        Set mobjPermisoActual = New Permiso
        Set mobjPermisoActual.Grupo = Me
    End If
    
    'Cuando cambia de empresa o primera vez, recupera los datos del objeto
    If mobjPermisoActual.CodEmpresa <> codemp Then
        mobjPermisoActual.CodEmpresa = codemp
        mobjPermisoActual.Recuperar
    End If
    Set PermisoActual = mobjPermisoActual
    
    mbooModificado = mod_antes      'Para que no cambie el estado de Modificado
    
    
'*** Antiguo: Se demora mucho porque recupera permisos para todas las empresas
'    If mcolPermisos Is Nothing Then RecuperaPermisos
'
'    Set ea = gobjMain.EmpresaActual
'    For Each p In mcolPermisos
'        If UCase$(p.CodEmpresa) = UCase$(ea.CodEmpresa) Then
'            Set PermisoActual = p
'            Exit For
'        End If
'    Next p
'    Set p = Nothing
'    Set ea = Nothing
End Property

Private Sub RecuperaPermisos()
    Dim sql As String, rs As Recordset, i As Long, mod_antes As Boolean
        
    mod_antes = mbooModificado      'Guarda el estado de Modificado
    If mcolPermisos Is Nothing Then Set mcolPermisos = New Collection
    
    'Recupera los permisos
    LimpiaColeccion mcolPermisos
    sql = "SELECT * FROM Permiso WHERE CodGrupo='" & CodGrupo & "' ORDER BY CodEmpresa"
#If DAOLIB Then
    Set rs = gobjMain.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly
#End If
    With rs
        Do Until .EOF
            i = AddPermiso
            Permisos(i).CodEmpresa = !CodEmpresa
            Permisos(i).RecuperarRS rs
            .MoveNext
        Loop
        .Close
    End With
    
    mbooModificado = mod_antes      'Para que no cambie el estado de Modificado
End Sub







Public Property Let CodGrupo(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "Grupo_CodGrupo", MSGERR_INVALIDO
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
        Err.Raise ERR_INVALIDO, "Grupo_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let Opcion(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "Grupo_Opcion", MSGERR_INVALIDO
    End If
    
    mProps.Opcion = value
    mbooModificado = True
End Property

Public Property Get Opcion() As String
    Opcion = Trim$(mProps.Opcion)
End Property


Public Property Get FechaGrabado() As Date
    FechaGrabado = mProps.FechaGrabado
End Property

Public Property Get EsNuevo() As Boolean
    EsNuevo = mbooEsNuevo
End Property

Friend Sub SetModificado()
    mbooModificado = True
End Sub

Public Property Get Modificado() As Boolean
    Modificado = mbooModificado
End Property
Private Sub Class_Initialize()
'    Set mcolPermisos = New Collection          '** Hace en RecuperPermisos
    
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mobjPermisoActual = Nothing
    Set mcolPermisos = Nothing
End Sub


Public Property Let BandCaja(ByVal value As Boolean)
    mProps.BandCaja = value
    mbooModificado = True
End Property

Public Property Get BandCaja() As Boolean
    BandCaja = mProps.BandCaja
End Property
'AUC 26/11/2005
Public Property Let BandVendedor(ByVal value As Boolean)
    mProps.BandVendedor = value
    mbooModificado = True
End Property
'AUC 26/11/2005
Public Property Get BandVendedor() As Boolean
    BandVendedor = mProps.BandVendedor
End Property

'jeaa 09/08/2006
Public Property Let BandBodeguero(ByVal value As Boolean)
    mProps.BandBodeguero = value
    mbooModificado = True
End Property
'jeaa 09/08/2006
Public Property Get BandBodeguero() As Boolean
    BandBodeguero = mProps.BandBodeguero
End Property

' 09/05/06  Agregado para la restriccion en la formula
Public Property Let BandPrincipal(ByVal value As Boolean)
    mProps.BandPrincipal = value
    'mProps.objIVInventario.Modificado = True
End Property

Public Property Get BandPrincipal() As Boolean
    BandPrincipal = mProps.BandPrincipal
End Property

' 09/05/06  Agregado para la restriccion en la formula
Public Property Let BandModificar(ByVal value As Boolean)
    mProps.BandModificar = value
    'mProps.objIVInventario.Modificado = True
End Property

Public Property Get BandModificar() As Boolean
    BandModificar = mProps.BandModificar
End Property


