VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Equipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    codEquipo As String
    UsuarioRed As String
    NombreEquipo As String
    Clave As String
    bandValida As Boolean
    FechaGrabado As Date
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean


Public Sub Grabar()
    Dim sql As String, i As Long
    Dim rs As Recordset
    Dim emp  As Sii4A32.Empresa
    Dim UsuarioRed As String, NombreEquipo As String, modulos As String, bandValida As Boolean, BandSupervisor As Boolean, Grupo As String
    Dim UsuarioRed_n As String, NombreEquipo_n As String, codequipo_n As String, BandValida_n As Boolean, BandSupervisor_n As Boolean, grupo_n As String, clave_n As String
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
#If DAOLIB = 0 Then
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
#End If
    If mbooEsNuevo Then
        'Verifica si no repite el código
        sql = "SELECT CodEquipo FROM Equipo WHERE CodEquipo='" & CifrarTexto(codEquipo, "Equipo") & "'"
#If DAOLIB Then
        Set rs = gobjMain.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
        rs.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly
#End If
        If Not rs.EOF Then      'Si ya existe el codigo
            Err.Raise ERR_REPITECODIGO, "Equipo", MSGERR_REPITECODIGO
        End If
        rs.Close
        
        sql = "SELECT * FROM Equipo WHERE 1=0"
#If DAOLIB Then
        Set rs = gobjMain.Database.OpenRecordset(sql)
#Else
        rs.Open sql, gobjMain.Coneccion, adOpenKeyset, adLockPessimistic
#End If
        rs.AddNew
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM Equipo WHERE CodEquipo='" & CifrarTexto(codEquipo, "Equipo") & "'"
#If DAOLIB Then
        Set rs = gobjMain.Database.OpenRecordset(sql)
        rs.Edit
#Else
        rs.Open sql, gobjMain.Coneccion, adOpenKeyset, adLockPessimistic
#End If
        
        'Verifica la FechaGrabado para asegurar que otros Equipo no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "Equipo_Grabar", MSGERR_MODIFICADO
        End If
    End If


    'Graba el Equipo
    With rs
        UsuarioRed = mProps.UsuarioRed
        bandValida = mProps.bandValida
        NombreEquipo = mProps.NombreEquipo
        
        UsuarioRed_n = !UsuarioRed
        BandValida_n = !bandValida
        NombreEquipo_n = !NombreEquipo
        
        
        
        !codEquipo = CifrarTexto(codEquipo, "Equipo")    'Cifra la clave
        !UsuarioRed = CifrarTexto(UsuarioRed, "Equipo")    'Cifra la clave
        !NombreEquipo = CifrarTexto(NombreEquipo, "Equipo")    'Cifra la clave
        !Clave = CifrarTexto(Clave, "Equipo")    'Cifra la clave
        !bandValida = bandValida  ' jeaa 23/09/2008
        
        
        !FechaGrabado = Now
        .Update
        
        If mbooEsNuevo Then
'            emp.GrabaGNLogAccion "NUE_EQU", "Equipo Nuevo: " & _
                                 codEquipo, "GN"
        Else

'                If Not NombreEquipo_n = NombreEquipo Then
'                    emp.GrabaAccionModulo NombreEquipo_n, NombreEquipo, "Codigo " & codEquipo & " Campo: " & "Nombre Equipo", "MOD_EQU", "GN"
'                End If
'
'                If Not UsuarioRed_n = UsuarioRed Then
'                    emp.GrabaAccionModulo UsuarioRed_n, NombreEquipo, "Codigo " & codEquipo & " Campo: " & "Nombre Equipo", "MOD_EQU", "GN"
'                End If
'                If Not codequipo_n = codEquipo Then
'                    emp.GrabaAccionModulo codequipo_n, codEquipo, "Codigo " & codEquipo & " Campo: " & "CodEquipo Equipo", "MOD_EQU", "GN"
'                End If


        End If

        
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Friend Sub Recuperar()
    Dim sql As String, rs As Recordset, rsaux As Recordset
    
    sql = "SELECT * FROM Equipo WHERE CodEquipo='" & CifrarTexto(codEquipo, "Equipo") & "'"
#If DAOLIB Then
    Set rs = gobjMain.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly

#End If
    'Recupera los datos de Equipo
    With rs
        If Not IsNull(!UsuarioRed) Then mProps.UsuarioRed = !UsuarioRed
        mProps.UsuarioRed = DecifrarTexto(mProps.UsuarioRed, "Equipo")
        
        If Not IsNull(!Clave) Then mProps.Clave = !Clave
        mProps.Clave = DecifrarTexto(mProps.Clave, "Equipo")
        
        If Not IsNull(!NombreEquipo) Then mProps.NombreEquipo = !NombreEquipo
        mProps.NombreEquipo = DecifrarTexto(mProps.NombreEquipo, "Equipo")
        
        If Not IsNull(!codEquipo) Then mProps.codEquipo = !codEquipo
        mProps.codEquipo = DecifrarTexto(mProps.codEquipo, "Equipo")

        
        If Not IsNull(!bandValida) Then mProps.bandValida = !bandValida  'jeaa 23/09/2008
        mProps.FechaGrabado = !FechaGrabado
        .Close
        

    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String
'anulado no se puede eliminar Equipo
    'Si no tiene derecho de supervisor, no puede hacer
    If Not gobjUsuarioActual.BandSupervisor Then
        Err.Raise ERR_NODERECHO, "Equipo_Eliminar", MSGERR_NODERECHO
    End If

    'Si es 'Supervisor' no es eliminable

    'Elimina el Equipo
    sql = "DELETE FROM Equipo WHERE CodEquipo='" & CifrarTexto(codEquipo, "Equipo") & "'"
#If DAOLIB Then
    gobjMain.Database.Execute sql, dbFailOnError
#Else
'***Agregado. 09/abr/02. Angel
'***Registra en todas las empresas que tenga acceso dicho Equipo,
'***para control de auditoría
    Dim emp As Sii4A32.Empresa, rs As Recordset, sql2 As String

    sql2 = "SELECT Permiso.CodEmpresa " & _
          "FROM Usuario INNER JOIN(" & _
                    "Grupo INNER JOIN Permiso " & _
                    "ON Grupo.CodGrupo = Permiso.CodGrupo) " & _
                "ON Usuario.CodGrupo = Grupo.CodGrupo "
    sql2 = sql2 & "WHERE Usuario.CodUsuario = '" & gobjMain.UsuarioActual.codUsuario & "'" & _
                " ORDER BY Permiso.CodEmpresa"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql2, gobjMain.Coneccion, adOpenStatic, adLockReadOnly
    With rs
        If Not (.BOF And .EOF) Then
            Do Until .EOF
                Set emp = gobjMain.RecuperaEmpresa(!CodEmpresa)
                If Not (emp Is Nothing) Then
                    emp.Abrir
                    
                    If emp.VerificaAuditoriaEquipo(UsuarioRed) Then
                        MsgBox "Este Equipo no puede ser eliminado, ya tiene registrados movimientos"
                        Set emp = Nothing
                        Exit Sub
                    End If
                    emp.GrabaGNLogAccion "EQUIPO", "Usuario Eliminado: " & _
                                             codEquipo, "GN"
                    emp.Cerrar
                End If
                Set emp = Nothing
                .MoveNext
            Loop
        End If
    End With
'***************
'***************

    gobjMain.Coneccion.Execute sql
#End If

    mbooEsNuevo = True
    mbooModificado = True
End Sub

Public Function Clone() As Equipo
    Dim nuevo As Equipo
    
    'Copia las propiedades de Equipo al nuevo objeto
    Set nuevo = New Equipo
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mbooModificado = True
End Property



Public Property Let codEquipo(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "Equipo_CodEquipo", MSGERR_INVALIDO
    End If
    
    mProps.codEquipo = value
    mbooModificado = True
End Property

Public Property Get codEquipo() As String
    codEquipo = Trim$(mProps.codEquipo)
End Property

Public Property Let UsuarioRed(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "Equipo_UsuarioRed", MSGERR_INVALIDO
    End If
    
    mProps.UsuarioRed = value
    mbooModificado = True
End Property

Public Property Get UsuarioRed() As String
    UsuarioRed = Trim$(mProps.UsuarioRed)
End Property

Public Property Let Clave(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "Equipo_Clave", MSGERR_INVALIDO
    End If
    
    mProps.Clave = value
    mbooModificado = True
End Property

Public Property Get Clave() As String
    Clave = Trim$(mProps.Clave)
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
    With mProps
        .bandValida = True
    End With
    
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Public Property Let bandValida(ByVal value As Boolean)
    mProps.bandValida = value
    mbooModificado = True
End Property

Public Property Get bandValida() As Boolean
    bandValida = mProps.bandValida
End Property



Public Property Let NombreEquipo(ByVal value As String)
    If Len(value) > 200 Then
        Err.Raise ERR_INVALIDO, "NombreEquipo_CodEquipo", MSGERR_INVALIDO
    End If
    
    mProps.NombreEquipo = value
    mbooModificado = True
End Property

Public Property Get NombreEquipo() As String
    NombreEquipo = Trim$(mProps.NombreEquipo)
End Property

