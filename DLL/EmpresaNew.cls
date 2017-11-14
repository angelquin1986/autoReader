VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmpresaNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Private Type T_PROPS
    CodEmpresa As String
    Descripcion As String
    TipoDB As Byte
    Ruta As String
    NombreDB As String
    Server As String
    Device As String
    UserName As String
    Password As String
    LimiteParaCompactar As Integer      'Límite de # de LOGIN para auto compactación
    fechagrabado As Date
    objSiiMain As SiiMain
    objGNOpcion As GNOpcion
End Type



Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean
Private mbooCompact As Boolean
Public mbooBandPorCobrar As Boolean

Private mcnEmpresa As ADODB.Connection

Friend Property Set SiiMain(value As SiiMain)
    Set mProps.objSiiMain = value
End Property

Friend Property Get SiiMain() As SiiMain
    Set SiiMain = mProps.objSiiMain
End Property

Public Property Let CodEmpresa(ByVal value As String)
    If Not mbooEsNuevo Then
        Err.Raise ERR_NOMODIFICABLE, "Empresa_CodEmpresa", MSGERR_NOMODIFICABLE
    End If
    
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "Empresa_CodEmpresa", MSGERR_INVALIDO
    End If
    
    mProps.CodEmpresa = value
    mbooModificado = True
End Property

Public Property Get CodEmpresa() As String
    CodEmpresa = Trim$(mProps.CodEmpresa)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "Empresa_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let TipoDB(ByVal value As Byte)
    If value <> TIPODB_JET And value <> TIPODB_SQL Then
        Err.Raise ERR_INVALIDO, "Empresa_TipoDB", MSGERR_INVALIDO
    End If
    
    mProps.TipoDB = value
    mbooModificado = True
End Property

Public Property Get TipoDB() As Byte
    TipoDB = mProps.TipoDB
End Property

Public Property Let Ruta(ByVal value As String)
    If Right$(value, 1) <> "\" And Len(value) > 0 Then value = value & "\"
    mProps.Ruta = value
    mbooModificado = True
End Property

Public Property Get Ruta() As String
    Ruta = Trim$(mProps.Ruta)
End Property

Public Property Let NombreDB(ByVal value As String)
    mProps.NombreDB = QuitaExtension(value)
    mbooModificado = True
End Property

Public Property Get NombreDB() As String
    NombreDB = Trim$(mProps.NombreDB)
End Property

Public Property Let Server(ByVal value As String)
    mProps.Server = value
    mbooModificado = True
End Property

Public Property Get Server() As String
    Server = Trim$(mProps.Server)
End Property

Public Property Let Device(ByVal value As String)
    mProps.Device = value
    mbooModificado = True
End Property

Public Property Get Device() As String
    Device = Trim$(mProps.Device)
End Property

Public Property Let UserName(ByVal value As String)
    mProps.UserName = value
    mbooModificado = True
End Property

Public Property Get UserName() As String
    UserName = Trim$(mProps.UserName)
End Property

Public Property Let Password(ByVal value As String)
    mProps.Password = value
    mbooModificado = True
End Property

Public Property Get Password() As String
    Password = Trim$(mProps.Password)
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



Public Sub EjecutarSQL( _
                ByVal sql As String, _
                ByRef numReg As Long)
    Dim f As Integer, s As String

    Select Case UCase(Left$(sql, 5))
    'Caso A. Si comienza con 'FILE:'
    Case "FILE:"
        sql = Right$(sql, Len(sql) - 5)         'Quita 'FILE:' y saca nombre de archivo a abrir

        'Lee el contenido del archivo a la variable s
        f = FreeFile
        Open sql For Binary As #f
        s = Input(LOF(f), f)
        Close #f
    
    'Caso B. Si comienza con 'PROC:'
    Case "PROC:"
        sql = Right$(sql, Len(sql) - 5)         'Quita 'PROC:' y saca nombre de procedimiento para ejecutarlo
        
        'Ejecuta todos los SQLs
        'EjecutarSiiEsquema sql, numReg
        s = ""
    
    'Caso C. Si envia directamente SQL
    Case Else
        s = sql
    End Select

    'En caso A. y C. aquí se ejecuta el SQL
    If Len(s) > 0 Then
        Me.Execute s, True, numReg
    End If
End Sub


Public Function OpenRecordset(ByVal sql As String) As Recordset
#If DAOLIB Then
    Set OpenRecordset = mdbEmpresa.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Dim rs As Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    'Modificado 17/07/2001 OLIVER
    'Pruebas de compilacion
    mcnEmpresa.CommandTimeout = GetSetting(APPNAME, SECTION, "TimeOut", 0)
        
    rs.Open sql, mcnEmpresa, adOpenStatic, adLockReadOnly
    
         Set OpenRecordset = rs
    Set rs = Nothing
#End If
End Function

'Encapsula la diferencia de OpenRecordset(DAO) y rs.Open(ADO)
'Esta sirve para hacer Edit/AddNew después
Public Function OpenRecordsetParaEdit(ByVal sql As String) As Recordset
#If DAOLIB Then
    Set OpenRecordsetParaEdit = mdbEmpresa.OpenRecordset(sql)
#Else
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.Open sql, mcnEmpresa, adOpenKeyset, adLockPessimistic
    Set OpenRecordsetParaEdit = rs
    Set rs = Nothing
#End If

End Function

'Encapsula la diferencia del metodo Execute de DAO y ADO
Public Sub Execute( _
                ByVal sql As String, _
                Optional ByVal FailOnError As Boolean, _
                Optional ByRef numReg As Long) 'AUC cambiado public por friend
#If DAOLIB Then
    If FailOnError Then
        mdbEmpresa.Execute sql, dbFailOnError
    Else
        mdbEmpresa.Execute sql
    End If
#Else
    mcnEmpresa.Execute sql, numReg
#End If
End Sub

Public Function NombrePC() As String
   'Declara variables locales
   Dim sEquipo As String * 255
   Dim sName As String
   Dim p As Integer
   'Obtiene nombre pc
   GetComputerName sEquipo, 255
   p = InStr(1, sEquipo, Chr(0), vbTextCompare)
   If p > 0 Then
      sName = Left(sEquipo, p - 1)
   End If
   'Retorna resultado
   NombrePC = UCase(Mid$(Trim$(sName), 1, 20))
End Function

Public Function NombreUsuarioPC() As String
   'Declara variables locales
   Dim sUsuario As String * 255
   Dim sName As String
   Dim p As Integer
   'Obtiene nombre pc
   'GetUserName(sName, 256)
   GetUserName sUsuario, 255
   p = InStr(1, sUsuario, Chr(0), vbTextCompare)
   If p > 0 Then
      sName = Left(sUsuario, p - 1)
   End If
   'Retorna resultado
   NombreUsuarioPC = UCase(Mid$(Trim$(sName), 1, 20))
End Function


Public Sub GrabaGNLogAccion(ByVal CodAccion As String, _
                            ByVal Descripcion As String, _
                            ByVal Modulo)
    Dim sql As String, codu As String, n As Long, rs As Recordset
    
    On Error Resume Next
    
    codu = "(Anonimo)"
    
    If Not (gobjUsuarioActual Is Nothing) Then
        codu = gobjUsuarioActual.codUsuario
    End If
    'jeaa 05/04/2005 antes 40
    Descripcion = Left(Descripcion, 200)
    'jeaa 23/11/2009 se aumenta nombreEquipo
    sql = "INSERT INTO GNLogAccion (CodUsuario,CodAccion,Descripcion,Modulo,Equipo,UsuarioRed) " & _
          " VALUES ('" & UCase(codu) & "','" & _
                         UCase(CodAccion) & "','" & _
                         Descripcion & "','" & _
                         UCase(Modulo) & "','" & NombrePC & "'" & _
                          ",'" & NombreUsuarioPC & "')"
    
    Me.Execute sql, True
    
    'Elimina registros antiguos
    If CodAccion = "LOGOUT" Then
        If Not (gobjMain Is Nothing) Then
            '*** MAKOTO 09/feb/01 Mod. , 11/abr/02 Angel mod. de 2000 a 3000
            'Elimina registros más antiguos que 60 días
            'ó si existe más de 3000 registros, hasta que quede solo 3000
            sql = "SELECT Count(*) AS Cnt FROM GNLogAccion"
            Set rs = Me.OpenRecordset(sql)
            If Not rs.EOF Then n = rs.Fields("Cnt") - 3000
            rs.Close
            Set rs = Nothing
            If n > 0 Then
                    '5000 regs
                    '19-01-05 jeaa pasa a la tabla RegAuditoria
                    sql = "Insert RegAuditoria select FechaHora, CodUsuario, CodAccion, Descripcion, " & _
                    " Modulo, equipo, UsuarioRed from GNlogaccion WHERE GNLOGACCION.ID < (SELECT MAX (A.ID)  FROM " & _
                    " (SELECT TOP " & n & " ID FROM GNLOGACCION ORDER BY FECHAHORA) AS A ) "
                    Me.OpenRecordset sql
                   
                    
                   sql = "DELETE GNLOGACCION " & _
                            " WHERE GNLOGACCION.ID < (SELECT MAX (A.ID)  FROM " & _
                            "(SELECT TOP " & n & " ID FROM GNLOGACCION ORDER BY FECHAHORA) AS A) "
    
                Else
                    'Solo 185 días  6 meses
                    '19-01-05 jeaa pasa a la tabla RegAuditoria
                    sql = "Insert RegAuditoria select FechaHora, CodUsuario, CodAccion, Descripcion, " & _
                    " Modulo, Equipo, UsuarioRed from GNlogaccion WHERE FechaHora<" & FechaYMD(Date - 185, mProps.TipoDB)
                    Me.OpenRecordset sql
                    'borra logaccion de la base de datos
                   sql = "DELETE FROM GNLogAccion " & _
                          "WHERE FechaHora<" & FechaYMD(Date - 185, mProps.TipoDB)
                End If
                Me.Execute sql, True
        End If
    End If
End Sub



