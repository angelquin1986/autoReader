VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Operador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type T_PROPS
    NombreOperador As String
    NombreCompleto As String
    Clave As String
    BandSupervisor As Boolean
    FechaGrabado As Date
End Type

Private mProps As T_PROPS
Private mcolPermisos As Collection
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean


Public Sub Grabar()
    Dim sql As String, i As Long
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
#If DAOLIB Then
    Dim rs As DAO.Recordset
#Else
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
#End If
    If mbooEsNuevo Then
        'Verifica si no repite el código
        sql = "SELECT NombreOperador FROM Operador WHERE NombreOperador='" & NombreOperador & "'"
#If DAOLIB Then
        Set rs = gobjMain.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
        rs.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly
#End If
        If Not rs.EOF Then      'Si ya existe el codigo
            Err.Raise ERR_REPITECODIGO, "Operador", MSGERR_REPITECODIGO
        End If
        rs.Close
        
        sql = "SELECT * FROM Operador"
#If DAOLIB Then
        Set rs = gobjMain.Database.OpenRecordset(sql)
#Else
        rs.Open sql, gobjMain.Coneccion, adOpenDynamic, adLockPessimistic
#End If
        rs.AddNew
    Else
        'Elimina todos los permisos
        sql = "DELETE FROM Permiso WHERE NombreOperador='" & NombreOperador & "'"
#If DAOLIB Then
        gobjMain.Database.Execute sql, dbFailOnError
#Else
        gobjMain.Coneccion.Execute sql
#End If
    
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM Operador WHERE NombreOperador='" & NombreOperador & "'"
#If DAOLIB Then
        Set rs = gobjMain.Database.OpenRecordset(sql)
        rs.Edit
#Else
        rs.Open sql, gobjMain.Coneccion, adOpenDynamic, adLockPessimistic
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "Operador_Grabar", MSGERR_MODIFICADO
        End If
    End If


    'Graba el Operador
    With rs
        !NombreOperador = NombreOperador
        !NombreCompleto = NombreCompleto
        !Clave = Clave
        !BandSupervisor = BandSupervisor
        
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    
    'Graba todos los permisos
    For i = 1 To CountPermiso
        Permisos(i).Grabar
    Next i
    
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Friend Sub Recuperar()
    Dim sql As String, i As Long
    
    sql = "SELECT * FROM Operador WHERE NombreOperador='" & NombreOperador & "'"
#If DAOLIB Then
    Dim rs As DAO.Recordset
    Set rs = gobjMain.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly
#End If
    'Recupera los datos de Operador
    With rs
        If Not IsNull(!NombreCompleto) Then mProps.NombreCompleto = !NombreCompleto
        If Not IsNull(!Clave) Then mProps.Clave = !Clave
        If Not IsNull(!BandSupervisor) Then mProps.BandSupervisor = !BandSupervisor
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    'Recupera los permisos
    sql = "SELECT CodEmpresa FROM Permiso WHERE NombreOperador='" & NombreOperador & "'"
#If DAOLIB Then
    Set rs = gobjMain.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    rs.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly
#End If
    With rs
        Do Until .EOF
            i = AddPermiso
            Permisos(i).CodEmpresa = !CodEmpresa
            Permisos(i).Recuperar
            .MoveNext
        Loop
        .Close
    End With
        
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String

    'Si no tiene derecho de supervisor, no puede hacer
    If Not gobjOperadorActual.BandSupervisor Then
        Err.Raise ERR_NODERECHO, "Operador_Eliminar", MSGERR_NODERECHO
    End If
    
    'Elimina los permisos
    sql = "DELETE FROM Permiso WHERE NombreOperador='" & NombreOperador & "'"
#If DAOLIB Then
    gobjMain.Database.Execute sql, dbFailOnError
#Else
    gobjMain.Coneccion.Execute sql
#End If
        
    'Elimina el Operador
    sql = "DELETE FROM Operador WHERE NombreOperador='" & NombreOperador & "'"
#If DAOLIB Then
    gobjMain.Database.Execute sql, dbFailOnError
#Else
    gobjMain.Coneccion.Execute sql
#End If

    mbooEsNuevo = True
    mbooModificado = True
End Sub

Friend Function Clone() As Operador
    Dim nuevo As Operador, i As Long, np As Permiso
    
    'Copia las propiedades de Operador al nuevo objeto
    Set nuevo = New Operador
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
    Set objPermiso.Operador = Me
    mcolPermisos.Add objPermiso
End Sub



Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
End Property

Public Function AddPermiso() As Long
    Dim nuevo As Permiso
    
    Set nuevo = New Permiso
    Set nuevo.Operador = Me
    mcolPermisos.Add nuevo
    AddPermiso = mcolPermisos.Count     'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function

Public Sub RemovePermiso(ix As Long)
    mcolPermisos.Remove ix
    mbooModificado = True
End Sub

'Duplica un permiso para el mismo operador
Public Function CopiaPermiso(ix As Long) As Long
    Dim nuevo As Permiso, obj As Permiso
    
    Set obj = mcolPermisos.Item(ix)
    Set nuevo = obj.Clone
    mcolPermisos.Add nuevo
    
    CopiaPermiso = mcolPermisos.Count       'Devuelve el indice del nuevo elemento
    Set nuevo = Nothing
    Set obj = Nothing
    mbooModificado = True
End Function

Public Property Get CountPermiso() As Long
    CountPermiso = mcolPermisos.Count
End Property

Public Property Get Permisos(ix As Long) As Permiso
    Set Permisos = mcolPermisos.Item(ix)
End Property

Public Property Get PermisoActual() As Permiso
    Dim p As Permiso
    
    For Each p In mcolPermisos
        If UCase(p.CodEmpresa) = UCase(gobjMain.EmpresaActual.CodEmpresa) Then
            Set PermisoActual = p
        End If
    Next p
    Set p = Nothing
End Property



Public Property Let NombreOperador(value As String)
    If Not mbooEsNuevo Then
        Err.Raise ERR_NOMODIFICABLE, "Operador_NombreOperador", MSGERR_NOMODIFICABLE
    End If
    
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "Operador_NombreOperador", MSGERR_INVALIDO
    End If
    
    mProps.NombreOperador = value
    mbooModificado = True
End Property

Public Property Get NombreOperador() As String
    NombreOperador = mProps.NombreOperador
End Property

Public Property Let NombreCompleto(value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "Operador_NombreCompleto", MSGERR_INVALIDO
    End If
    
    mProps.NombreCompleto = value
    mbooModificado = True
End Property

Public Property Get NombreCompleto() As String
    NombreCompleto = mProps.NombreCompleto
End Property

Public Property Let Clave(value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "Operador_Clave", MSGERR_INVALIDO
    End If
    
    mProps.Clave = value
    mbooModificado = True
End Property

Public Property Get Clave() As String
    Clave = mProps.Clave
End Property

Public Property Let BandSupervisor(value As Boolean)
    mProps.BandSupervisor = value
    mbooModificado = True
End Property

Public Property Get BandSupervisor() As Boolean
    BandSupervisor = mProps.BandSupervisor
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
    Set mcolPermisos = New Collection
    
    With mProps
        .BandSupervisor = False
    End With
    
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mcolPermisos = Nothing
End Sub
