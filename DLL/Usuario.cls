VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Usuario"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type T_PROPS
    CodUsuario As String
    CodGrupo As String
    NombreUsuario As String
    Clave As String
    BandSupervisor As Boolean
    FechaGrabado As Date
    modulos  As String 'AUC 31/05/07 Para permisos de que ha modulos debe entrar
    BandValida As Boolean
    BandNoImpRep As Boolean
    BandCambiaClave As Boolean
    BandPresupuesto As Boolean
    BandSeccionItem As Boolean
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean


Public Sub Grabar()
    Dim sql As String, i As Long
    Dim rs As Recordset
    Dim emp  As Sii4A32.Empresa
    Dim CodGrupo As String, NombreUsuario As String, modulos As String, BandValida As Boolean, BandSupervisor As Boolean, Grupo As String
    Dim codgrupo_n As String, NombreUsuario_n As String, modulos_n As String, BandValida_n As Boolean, BandSupervisor_n As Boolean, grupo_n As String, clave_n As String
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
#If DAOLIB = 0 Then
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
#End If
    If mbooEsNuevo Then
        'Verifica si no repite el código
        sql = "SELECT CodUsuario FROM Usuario WHERE CodUsuario='" & CodUsuario & "'"
#If DAOLIB Then
        Set rs = gobjMain.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
        rs.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly
#End If
        If Not rs.EOF Then      'Si ya existe el codigo
            Err.Raise ERR_REPITECODIGO, "Usuario", MSGERR_REPITECODIGO
        End If
        rs.Close
        
        sql = "SELECT * FROM Usuario WHERE 1=0"
#If DAOLIB Then
        Set rs = gobjMain.Database.OpenRecordset(sql)
#Else
        rs.Open sql, gobjMain.Coneccion, adOpenKeyset, adLockPessimistic
#End If
        rs.AddNew
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM Usuario WHERE CodUsuario='" & CodUsuario & "'"
#If DAOLIB Then
        Set rs = gobjMain.Database.OpenRecordset(sql)
        rs.Edit
#Else
        rs.Open sql, gobjMain.Coneccion, adOpenKeyset, adLockPessimistic
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "Usuario_Grabar", MSGERR_MODIFICADO
        End If
    End If


    'Graba el Usuario
    With rs
        
        

        CodGrupo = mProps.CodGrupo
        NombreUsuario = mProps.NombreUsuario
        BandSupervisor = mProps.BandSupervisor
        modulos = mProps.modulos
        BandValida = mProps.BandValida
        
        codgrupo_n = !CodGrupo
        NombreUsuario_n = !NombreUsuario
        BandSupervisor_n = !BandSupervisor
        modulos_n = !modulos
        BandValida_n = !BandValida
        
        clave_n = DecifrarTexto((!Clave), "Usuario")
        
        !CodUsuario = CodUsuario
        !CodGrupo = CodGrupo
        
        If !NombreUsuario <> NombreUsuario Then
            !BandCambiaClave = True
        Else
            !BandCambiaClave = mProps.BandCambiaClave
        End If
        
        If !CodGrupo = CodGrupo Then
            !BandCambiaClave = True
        Else
            If Not !BandCambiaClave Then !BandCambiaClave = mProps.BandCambiaClave
        End If
        
        
        If DecifrarTexto((!Clave), "Usuario") <> Clave Then
            !BandCambiaClave = True
        Else
            If Not !BandCambiaClave Then !BandCambiaClave = mProps.BandCambiaClave
        End If
        
        !NombreUsuario = NombreUsuario
         !Clave = CifrarTexto(Clave, "Usuario")    'Cifra la clave
      ' !Clave = CifrarTexto(LCase$(Clave), "Usuario")    'Cifra la clave
        !BandSupervisor = BandSupervisor
        !modulos = modulos  'AUC 31/05/07
        !BandValida = BandValida  ' jeaa 23/09/2008
        !BandNoImpRep = mProps.BandNoImpRep
        !BandPresupuesto = mProps.BandPresupuesto
        !BandSeccionItem = mProps.BandSeccionItem
        
        
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
'***Agregado. 09/abr/02. Angel
'***Registra en todas las empresas que tenga acceso dicho usuario,
'***para control de auditoría
    
    
    sql = "SELECT Permiso.CodEmpresa " & _
          "FROM Usuario INNER JOIN(" & _
                    "Grupo INNER JOIN Permiso " & _
                    "ON Grupo.CodGrupo = Permiso.CodGrupo) " & _
                "ON Usuario.CodGrupo = Grupo.CodGrupo "
    sql = sql & "WHERE Usuario.CodUsuario = '" & CodUsuario & "'" & _
                " ORDER BY Permiso.CodEmpresa"
    rs.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly
    With rs
        If Not (.BOF And .EOF) Then
            Do Until .EOF
                Set emp = gobjMain.RecuperaEmpresa(!CodEmpresa)
                If Not (emp Is Nothing) Then
                    emp.Abrir
                    If mbooEsNuevo Then
                        emp.GrabaGNLogAccion "NUE_USR", "Usuario Nuevo: " & _
                                             CodUsuario, "GN"
                    Else
                        'emp.GrabaGNLogAccion "USUARIO", "Usuario Modificado: " & _
                        CodUsuario, "GN"
                         If Not codgrupo_n = CodGrupo Then
                            'emp.GrabaAccionModulo codgrupo_n, CodGrupo, "Codigo " & CodUsuario & " Campo: " & "Codigo Grupo", "MOD_USR", "GN"
                        End If
                        If Not NombreUsuario_n = NombreUsuario Then
                            emp.GrabaAccionModulo NombreUsuario_n, NombreUsuario, "Codigo " & CodUsuario & " Campo: " & "Nombre Usuario", "MOD_USR", "GN"
                        End If
                        If Not BandSupervisor_n = BandSupervisor Then
                            emp.GrabaAccionModulo BandSupervisor_n, BandSupervisor, "Codigo " & CodUsuario & " Campo: " & "Band Supervisor", "MOD_USR", "GN"
                        End If
                        If Not modulos_n = modulos Then
                            emp.GrabaAccionModulo ModulosSeleccionadosAuditoria(modulos_n), ModulosSeleccionadosAuditoria(modulos), "Codigo " & CodUsuario & " Campo: " & "Modulos", "MOD_USR", "GN"
                        End If
                        If Not BandValida_n = BandValida Then
                            emp.GrabaAccionModulo BandValida_n, BandValida, "Codigo " & CodUsuario & " Campo: " & "Band Activo", "MOD_USR", "GN"
                        End If
                        If Not clave_n = Clave Then
                            emp.GrabaAccionModulo "***", "***", "Codigo " & CodUsuario & " Campo: " & "Clave", "MOD_USR", "GN"
                        End If
                        
                        
                    End If
                    emp.Cerrar
                End If
                Set emp = Nothing
                .MoveNext
            Loop
        End If
    End With
'***************
'***************
    
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Friend Sub Recuperar()
    Dim sql As String, rs As Recordset, rsaux As Recordset
    
    sql = "SELECT * FROM Usuario WHERE CodUsuario='" & CodUsuario & "'"
#If DAOLIB Then
    Set rs = gobjMain.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly

#End If
    'Recupera los datos de Usuario
    With rs
        If Not IsNull(!CodGrupo) Then mProps.CodGrupo = !CodGrupo
        If Not IsNull(!NombreUsuario) Then mProps.NombreUsuario = !NombreUsuario
        
        If Not IsNull(!Clave) Then mProps.Clave = !Clave
        'Decifra la clave       '*** MAKOTO 14/sep/00 Agregado
        mProps.Clave = DecifrarTexto(mProps.Clave, "Usuario")
        
        If Not IsNull(!BandSupervisor) Then mProps.BandSupervisor = !BandSupervisor
        If Not IsNull(!modulos) Then mProps.modulos = !modulos 'AUC 31/05/07
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida  'jeaa 23/09/2008
        If Not IsNull(!BandNoImpRep) Then mProps.BandNoImpRep = !BandNoImpRep 'No imprimir reportes
        If Not IsNull(!BandCambiaClave) Then mProps.BandCambiaClave = !BandCambiaClave 'Cambia de Clave la primera vez que ingresa al sistema
        If Not IsNull(!BandPresupuesto) Then mProps.BandPresupuesto = !BandPresupuesto
        If Not IsNull(!BandSeccionItem) Then mProps.BandSeccionItem = !BandSeccionItem
      
        mProps.FechaGrabado = !FechaGrabado
        .Close
        

    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String
'anulado no se puede eliminar usuario
    'Si no tiene derecho de supervisor, no puede hacer
    If Not gobjUsuarioActual.BandSupervisor Then
        Err.Raise ERR_NODERECHO, "Usuario_Eliminar", MSGERR_NODERECHO
    End If

    'Si es 'Supervisor' no es eliminable
    If LCase(mProps.CodUsuario) = "supervisor" Then
        Err.Raise ERR_NOELIMINABLE, "Usuario_Eliminar", MSGERR_NOELIMINABLE
    End If

    'Elimina el Usuario
    sql = "DELETE FROM Usuario WHERE CodUsuario='" & CodUsuario & "'"
#If DAOLIB Then
    gobjMain.Database.Execute sql, dbFailOnError
#Else
'***Agregado. 09/abr/02. Angel
'***Registra en todas las empresas que tenga acceso dicho usuario,
'***para control de auditoría
    Dim emp As Sii4A32.Empresa, rs As Recordset, sql2 As String

    sql2 = "SELECT Permiso.CodEmpresa " & _
          "FROM Usuario INNER JOIN(" & _
                    "Grupo INNER JOIN Permiso " & _
                    "ON Grupo.CodGrupo = Permiso.CodGrupo) " & _
                "ON Usuario.CodGrupo = Grupo.CodGrupo "
    sql2 = sql2 & "WHERE Usuario.CodUsuario = '" & CodUsuario & "'" & _
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
                    If emp.VerificaAuditoriaUsuario(CodUsuario) Then
                        MsgBox "Este usuario no puede ser eliminado, ya tiene registrados movimientos"
                        Set emp = Nothing
                        Exit Sub
                    End If
                    emp.GrabaGNLogAccion "USUARIO", "Usuario Eliminado: " & _
                                             CodUsuario, "GN"
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

Public Function Clone() As Usuario
    Dim nuevo As Usuario
    
    'Copia las propiedades de Usuario al nuevo objeto
    Set nuevo = New Usuario
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mbooModificado = True
End Property



Public Property Let CodUsuario(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "Usuario_CodUsuario", MSGERR_INVALIDO
    End If
    
    mProps.CodUsuario = value
    mbooModificado = True
End Property

Public Property Get CodUsuario() As String
    CodUsuario = Trim$(mProps.CodUsuario)
End Property

Public Property Let CodGrupo(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "Usuario_CodGrupo", MSGERR_INVALIDO
    End If
    
    mProps.CodGrupo = value
    mbooModificado = True
End Property

Public Property Get CodGrupo() As String
    CodGrupo = Trim$(mProps.CodGrupo)
End Property

Public Property Let NombreUsuario(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "Usuario_NombreUsuario", MSGERR_INVALIDO
    End If
    
    mProps.NombreUsuario = value
    mbooModificado = True
End Property

Public Property Get NombreUsuario() As String
    NombreUsuario = Trim$(mProps.NombreUsuario)
End Property

Public Property Let Clave(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "Usuario_Clave", MSGERR_INVALIDO
    End If
    
    mProps.Clave = value
    mbooModificado = True
End Property

Public Property Get Clave() As String
    Clave = Trim$(mProps.Clave)
End Property

Public Property Let BandSupervisor(ByVal value As Boolean)
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

Public Property Get Modificado() As Boolean
    Modificado = mbooModificado
End Property
Private Sub Class_Initialize()
    With mProps
        .BandSupervisor = False
        .BandValida = True
        .BandCambiaClave = False
    End With
    
    mbooEsNuevo = True
    mbooModificado = False
End Sub
'AUC 31/05/07
Public Property Let modulos(ByVal value As String)
    If Len(value) > 250 Then
        Err.Raise ERR_INVALIDO, "Usuario_Modulos", MSGERR_INVALIDO
    End If
    mProps.modulos = value
    mbooModificado = True
End Property

Public Property Get modulos() As String
    modulos = Trim$(mProps.modulos)
End Property

Public Property Let BandValida(ByVal value As Boolean)
    mProps.BandValida = value
    mbooModificado = True
End Property

Public Property Get BandValida() As Boolean
    BandValida = mProps.BandValida
End Property

Public Sub GrabarAuditoriaCambioGrupo(ByVal CodGrupoAnterior As String)
End Sub

Public Sub GrabarAuditoriaporCambioGrupo(ByVal CodGrupoAnterior As String, ByVal CodGrupoNuevo As String)
    Dim sql As String, i As Long
    Dim rs As Recordset
    Dim emp  As Sii4A32.Empresa
    Dim CodGrupo As String, NombreUsuario As String, modulos As String, BandValida As Boolean, BandSupervisor As Boolean, Grupo As String
    Dim codgrupo_n As String, NombreUsuario_n As String, modulos_n As String, BandValida_n As Boolean, BandSupervisor_n As Boolean, grupo_n As String
    'Si no está modificado no hace nada
'    If Not mbooModificado Then Exit Sub
    
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    
    
    sql = "SELECT Permiso.CodEmpresa " & _
          "FROM Usuario INNER JOIN(" & _
                    "Grupo INNER JOIN Permiso " & _
                    "ON Grupo.CodGrupo = Permiso.CodGrupo) " & _
                "ON Usuario.CodGrupo = Grupo.CodGrupo "
    sql = sql & "WHERE Usuario.CodUsuario = '" & CodUsuario & "'" & _
                " ORDER BY Permiso.CodEmpresa"
    rs.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly
    With rs
        If Not (.BOF And .EOF) Then
            Do Until .EOF
                Set emp = gobjMain.RecuperaEmpresa(!CodEmpresa)
                If Not (emp Is Nothing) Then
                    emp.Abrir
                    If mbooEsNuevo Then
                        emp.GrabaGNLogAccion "NUE_USR", "Usuario Nuevo: " & _
                                             CodUsuario, "GN"
                    Else
                         If CodGrupoAnterior <> CodGrupoNuevo Then
                            emp.GrabaAccionModulo CodGrupoAnterior, CodGrupoNuevo, "Codigo " & CodUsuario & " Campo: " & "Codigo Grupo", "MOD_USR", "GN"
                        End If
                        
                        
                    End If
                    emp.Cerrar
                End If
                Set emp = Nothing
                .MoveNext
            Loop
        End If
    End With
End Sub

Public Property Let BandNoImpRep(ByVal value As Boolean)
    mProps.BandNoImpRep = value
    mbooModificado = True
End Property

Public Property Get BandNoImpRep() As Boolean
    BandNoImpRep = mProps.BandNoImpRep
End Property

Public Property Let BandCambiaClave(ByVal value As Boolean)
    mProps.BandCambiaClave = value
    mbooModificado = True
End Property

Public Property Get BandCambiaClave() As Boolean
    BandCambiaClave = mProps.BandCambiaClave
End Property


Private Function ModulosSeleccionadosAuditoria(cad As String) As String
 Dim s As String, i As Long, v As Variant
    On Error GoTo ErrTrap
    'If mbooIniciando Then Exit Sub
    v = Split(cad, ";")
        For i = 0 To UBound(v)


                s = s & Left(Mid$(v(i), 2, Len(v(i)) - 2), 4) & ","
            
        Next i
        s = Mid$(s, 1, Len(s) - 1)
    ModulosSeleccionadosAuditoria = s
    Exit Function
ErrTrap:
    DispErr
    Exit Function
End Function

Public Sub GrabarCambioClave()
    Dim sql As String, i As Long
    Dim rs As Recordset
    Dim emp  As Sii4A32.Empresa
    Dim CodGrupo As String, NombreUsuario As String, modulos As String, BandValida As Boolean, BandSupervisor As Boolean, Grupo As String
    Dim codgrupo_n As String, NombreUsuario_n As String, modulos_n As String, BandValida_n As Boolean, BandSupervisor_n As Boolean, grupo_n As String, clave_n As String
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
#If DAOLIB = 0 Then
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
#End If
    If mbooEsNuevo Then
        'Verifica si no repite el código
        sql = "SELECT CodUsuario FROM Usuario WHERE CodUsuario='" & CodUsuario & "'"
#If DAOLIB Then
        Set rs = gobjMain.Database.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
        rs.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly
#End If
        If Not rs.EOF Then      'Si ya existe el codigo
            Err.Raise ERR_REPITECODIGO, "Usuario", MSGERR_REPITECODIGO
        End If
        rs.Close
        
        sql = "SELECT * FROM Usuario WHERE 1=0"
#If DAOLIB Then
        Set rs = gobjMain.Database.OpenRecordset(sql)
#Else
        rs.Open sql, gobjMain.Coneccion, adOpenKeyset, adLockPessimistic
#End If
        rs.AddNew
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM Usuario WHERE CodUsuario='" & CodUsuario & "'"
#If DAOLIB Then
        Set rs = gobjMain.Database.OpenRecordset(sql)
        rs.Edit
#Else
        rs.Open sql, gobjMain.Coneccion, adOpenKeyset, adLockPessimistic
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "Usuario_Grabar", MSGERR_MODIFICADO
        End If
    End If


    'Graba el Usuario
    With rs
        
        clave_n = DecifrarTexto((!Clave), "Usuario")
        
'        !Clave = CifrarTexto(LCase$(Clave), "Usuario")    'Cifra la clave
        !Clave = CifrarTexto(Clave, "Usuario")    'Cifra la clave
        !BandCambiaClave = 0
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    
        sql = "SELECT Permiso.CodEmpresa " & _
          "FROM Usuario INNER JOIN(" & _
                    "Grupo INNER JOIN Permiso " & _
                    "ON Grupo.CodGrupo = Permiso.CodGrupo) " & _
                "ON Usuario.CodGrupo = Grupo.CodGrupo "
    sql = sql & "WHERE Usuario.CodUsuario = '" & CodUsuario & "'" & _
                " ORDER BY Permiso.CodEmpresa"
    rs.Open sql, gobjMain.Coneccion, adOpenStatic, adLockReadOnly
    With rs
        If Not (.BOF And .EOF) Then
            Do Until .EOF
                Set emp = gobjMain.RecuperaEmpresa(!CodEmpresa)
                If Not (emp Is Nothing) Then
                    emp.Abrir
                    If mbooEsNuevo Then
                        emp.GrabaGNLogAccion "NUE_USR", "Usuario Nuevo: " & _
                                             CodUsuario, "GN"
                    Else
'                        If Not clave_n = Clave Then
                            emp.GrabaAccionModulo "***", "***", "Codigo " & CodUsuario & " Campo: " & "Clave", "MOD_USR", "GN"
'                        End If
                        
                        
                    End If
                    emp.Cerrar
                End If
                Set emp = Nothing
                .MoveNext
            Loop
        End If
    End With
    
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Property Let BandPresupuesto(ByVal value As Boolean)
    mProps.BandPresupuesto = value
    mbooModificado = True
End Property

Public Property Get BandPresupuesto() As Boolean
    BandPresupuesto = mProps.BandPresupuesto
End Property

Public Property Let BandSeccionItem(ByVal value As Boolean)
    mProps.BandSeccionItem = value
    mbooModificado = True
End Property

Public Property Get BandSeccionItem() As Boolean
    BandSeccionItem = mProps.BandSeccionItem
End Property

