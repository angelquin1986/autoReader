VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Motivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'jeaa 25/04/2005
Option Explicit

Private Type T_PROPS
    CodMotivo As String
    Descripcion As String
    FechaGrabado As Date
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
    
    On Error GoTo errtrap
    
#If DAOLIB Then
    gobjWS.BeginTrans                   '======= Comienza la transacción
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    gobjMain.Coneccion.BeginTrans       '======= Comienza la transacción
#End If

    '*** MAKOTO 27/dic/00 Agregado
    'Avisa al programa UI
    RaiseEvent Procesando("Está grabándo datos de Motivos...")

    If mbooEsNuevo Then
        'Verifica si no repite el código
        sql = "SELECT CodMotivo FROM Motivo WHERE CodMotivo='" & CodMotivo & "'"
#If DAOLIB Then
        Set rs = gobjMain.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
        rs.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly
#End If
        If Not rs.EOF Then      'Si ya existe el codigo
            Err.Raise ERR_REPITECODIGO, "Motivo", MSGERR_REPITECODIGO
        End If
        rs.Close
        
        sql = "SELECT * FROM Motivo WHERE 1=0"
#If DAOLIB Then
        Set rs = gobjMain.Database.OpenRecordset(sql)
#Else
        rs.Open sql, gobjMain.Coneccion, adOpenKeyset, adLockPessimistic
#End If
        rs.AddNew
        
    'Cuando es Modificación
    Else
        'Elimina todos los permisos que pertenece en el Motivo
        sql = "DELETE FROM Permiso WHERE CodMotivo='" & CodMotivo & "'"
#If DAOLIB Then
        gobjMain.Database.Execute sql, dbFailOnError
#Else
        gobjMain.Coneccion.Execute sql
#End If
    
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM Motivo WHERE CodMotivo='" & CodMotivo & "'"
#If DAOLIB Then
        Set rs = gobjMain.Database.OpenRecordset(sql)
        rs.Edit
#Else
        rs.Open sql, gobjMain.Coneccion, adOpenKeyset, adLockPessimistic
#End If
        
        'Verifica la FechaGrabado para asegurar que otros Motivos no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "Motivo_Grabar", MSGERR_MODIFICADO
        End If
    End If


    'Graba el Motivo
    With rs
        !CodMotivo = CodMotivo
        !Descripcion = Descripcion
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
errtrap:
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
    
    sql = "SELECT * FROM Motivo WHERE CodMotivo='" & CodMotivo & "'"
#If DAOLIB Then
    Set rs = gobjMain.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly
#End If
    'Recupera los datos de Motivo
    With rs
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String

    'Si no tiene derecho de supervisor, no puede hacer
    If Not gobjUsuarioActual.BandSupervisor Then
        Err.Raise ERR_NODERECHO, "Motivo_Eliminar", MSGERR_NODERECHO
    End If
    
    'Elimina los permisos
    sql = "DELETE FROM Permiso WHERE CodMotivo='" & CodMotivo & "'"
#If DAOLIB Then
    gobjMain.Database.Execute sql, dbFailOnError
#Else
    gobjMain.Coneccion.Execute sql
#End If
        
    'Elimina el Motivo
    sql = "DELETE FROM Motivo WHERE CodMotivo='" & CodMotivo & "'"
#If DAOLIB Then
    gobjMain.Database.Execute sql, dbFailOnError
#Else
    gobjMain.Coneccion.Execute sql
#End If

    mbooEsNuevo = True
    mbooModificado = True
End Sub

Public Function Clone() As Motivo
    Dim nuevo As Motivo, i As Long, np As Permiso
    
    If mcolPermisos Is Nothing Then RecuperaPermisos
    
    'Copia las propiedades de Motivo al nuevo objeto
    Set nuevo = New Motivo
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
    
    Set objPermiso.Motivo = Me
    mcolPermisos.Add objPermiso
End Sub

Public Function ListaPermisos() As Variant
    Dim sql As String, rs As Recordset

    sql = "SELECT Id,CodEmpresa FROM Permiso WHERE CodMotivo='" & mProps.CodMotivo & "' ORDER BY CodEmpresa"

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
    Set nuevo.Motivo = Me
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

'Duplica un permiso para el mismo Motivo
Public Function CopiaPermiso(ix As Long) As Long
    Dim nuevo As Permiso, obj As Permiso
    
    If mcolPermisos Is Nothing Then RecuperaPermisos
    
    Set obj = mcolPermisos.Item(ix)
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
    Set Permisos = mcolPermisos.Item(ix)
End Property

Public Property Get PermisoActual() As Permiso
    Dim mod_antes As Boolean, codemp As String
    
    mod_antes = mbooModificado      'Guarda el estado de Modificado
    codemp = gobjMain.EmpresaActual.CodEmpresa
    
    'Si es primera vez, crea una instancia del Permiso
    If mobjPermisoActual Is Nothing Then
        Set mobjPermisoActual = New Permiso
        Set mobjPermisoActual.Motivo = Me
    End If
    
    'Cuando cambia de empresa o primera vez, recupera los datos del objeto
    If mobjPermisoActual.CodEmpresa <> codemp Then
        mobjPermisoActual.CodEmpresa = codemp
        mobjPermisoActual.Recuperar
    End If
    Set PermisoActual = mobjPermisoActual
    
    mbooModificado = mod_antes      'Para que no cambie el estado de Modificado
End Property

Private Sub RecuperaPermisos()
    Dim sql As String, rs As Recordset, i As Long, mod_antes As Boolean
        
    mod_antes = mbooModificado      'Guarda el estado de Modificado
    If mcolPermisos Is Nothing Then Set mcolPermisos = New Collection
    
    'Recupera los permisos
    LimpiaColeccion mcolPermisos
    sql = "SELECT * FROM Permiso WHERE CodMotivo='" & CodMotivo & "' ORDER BY CodEmpresa"
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

Public Property Let CodMotivo(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "Motivo_CodMotivo", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodMotivo = value
    mbooModificado = True
End Property

Public Property Get CodMotivo() As String
    CodMotivo = Trim$(mProps.CodMotivo)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "Motivo_Descripcion", MSGERR_INVALIDO
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
