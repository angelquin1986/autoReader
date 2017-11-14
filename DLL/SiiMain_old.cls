VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SiiMain"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Valores predeterminados            '*** MAKOTO 08/sep/00 Agregado
Private Const RUTA_PRE = "C:\Archivos de programa\Sii4\"
Private Const NOMBREDB_PRE = "SiiMain"
Private Const RUTATEMP_PRE = "c:\Windows\temp"

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long


'Private Const FECHA_NOCADUCA = "NUNCA140900"    'Para verificación de caducidad
'Private Const FECHA_NOCADUCA = "MVMDB251011"    'Para verificación de caducidad
Private Const FECHA_NOCADUCA = "OXOFD473233"    'Para verificación de caducidad

Private Type T_PROPS
    TipoDB As Byte
    Ruta As String
    NombreDB As String
    NombreDBSQL As String
    Server As String
    Device As String
    UserName As String
    UserNameSQL As String
    Password As String
    PasswordSQL As String
    FormatoFechaSQL As String       'Formato de fecha en SQLServer
    TeclaParaCerrar As String
    RutaTempDB As String
    PosID As String
    booDEMO As Boolean
    CPicture As String ' AUC agregado para guardar la direccion de las fotos de IV 05/04/06
    Modulo As String 'AUC 31/05/07
    ModuloCargado As String
End Type

Private mProps As T_PROPS
#If DAOLIB Then
    Private mdbMain As DAO.Database
#Else
    Private mcnMain As ADODB.Connection
    Private mcnMainAux As ADODB.Connection
    Private mcnMainReport As ADODB.Connection
#End If

    Private numFile As Integer
    Private file As String


Private Sub Class_Initialize()
    'Valores predeterminados
    With mProps
        .TipoDB = TIPODB_JET
        .Ruta = RUTA_PRE
        .NombreDB = NOMBREDB_PRE
        .NombreDBSQL = NOMBREDB_PRE
        .Server = ""
        .Device = ""
        .UserName = "admin"
        .Password = "admin111"
        .UserNameSQL = "todos"
        .PasswordSQL = "111"
        .FormatoFechaSQL = "dd/MM/yyyy"
        .TeclaParaCerrar = "ESC"
        .RutaTempDB = "c:\Windows\temp"
        .PosID = "C01"
        
    End With
End Sub

Public Sub Inicializar()
    Dim band As Boolean, objSeg As Sii4Seg.clsSeg, strInicializado As String

        file = "c:\ia\TramaTrans\1.1.0 Inicializar" & "-" & Format(Date, "ddmmmyyyy") & "-" & Format(Time, "hhmmss") & ".txt"
        numFile = FreeFile
        Open file For Output Access Write As #numFile
        Print #numFile, "Inicializar"
        Close numFile


    'Verifica la clave de licencia
    mProps.booDEMO = False
    Set objSeg = New Sii4Seg.clsSeg
    
        file = "c:\ia\TramaTrans\1.1.0.1 Inicializar" & "-" & Format(Date, "ddmmmyyyy") & "-" & Format(Time, "hhmmss") & ".txt"
        numFile = FreeFile
        Open file For Output Access Write As #numFile
        Print #numFile, "Inicializar"
        Close numFile
    
'''''        If Not objSeg.VerificaClaveLicencia Then
'''''            MsgBox "El programa se ejecutará en modo de demostración hasta que ingrese la clave de licencia." & vbCr & vbCr & _
'''''                   "Mientras tanto usted tendrá los siguientes limitaciones:" & vbCr & vbCr & _
'''''                   "    - Puede registrar sólo hasta " & MAXTRANS_DEMO & " transacciones en total.", vbInformation
'''''            mProps.booDEMO = True     'Indica que está en modo DEMO
'''''        End If

    'Asigna si mismo a la variable global
    
    
        file = "c:\ia\TramaTrans\1.1.1 Inicializar" & "-" & Format(Date, "ddmmmyyyy") & "-" & Format(Time, "hhmmss") & ".txt"
        numFile = FreeFile
        Open file For Output Access Write As #numFile
        Print #numFile, "Inicializar"
        Close numFile
    
    
        Set gobjMain = Me
        
        file = "c:\ia\TramaTrans\1.1.2 ME" & "-" & Format(Date, "ddmmmyyyy") & "-" & Format(Time, "hhmmss") & ".txt"
        numFile = FreeFile
        Open file For Output Access Write As #numFile
        Print #numFile, "Me"
        Close numFile
    
    'Crea una instancia del objeto Condicion
    Set gobjCond = New Condicion

        file = "c:\ia\TramaTrans\1.1.3 Condicion" & "-" & Format(Date, "ddmmmyyyy") & "-" & Format(Time, "hhmmss") & ".txt"
        numFile = FreeFile
        Open file For Output Access Write As #numFile
        Print #numFile, "Condicion"
        Close numFile

    'Crea una instancia de coneccion a la base
#If DAOLIB Then
#Else
    
    
    Set mcnMain = New ADODB.Connection
    Set mcnMainAux = New ADODB.Connection
    Set mcnMainReport = New ADODB.Connection
#End If
    
    'Recupera las informaciones de base principal desde el registro de sistema
    With mProps
        .TipoDB = GetSetting(APPNAME, SECTION, "TipoDB", TIPODB_JET)
        .Ruta = GetSetting(APPNAME, SECTION, "Ruta", RUTA_PRE)
        .NombreDB = GetSetting(APPNAME, SECTION, "NombreDB", NOMBREDB_PRE)
        .NombreDBSQL = GetSetting(APPNAME, SECTION, "NombreDBSQL", NOMBREDB_PRE)
        .Server = GetSetting(APPNAME, SECTION, "Server", "")
        .Device = GetSetting(APPNAME, SECTION, "Device", "")
        .UserName = GetSetting(APPNAME, SECTION, "UserName", "admin")
        .UserNameSQL = GetSetting(APPNAME, SECTION, "UserNameSQL", "todos")
        .Password = GetSetting(APPNAME, SECTION, "Password", "")
        .Password = DecifrarTexto(.Password)       'Decifrar texto *** MAKOTO 14/sep/00
        .PasswordSQL = GetSetting(APPNAME, SECTION, "PasswordSQL", "")
        .PasswordSQL = DecifrarTexto(.PasswordSQL) 'Decifrar texto *** MAKOTO 14/sep/00
        .FormatoFechaSQL = GetSetting(APPNAME, SECTION, "FormatoFechaSQL", "dd/MM/yyyy")
        .TeclaParaCerrar = GetSetting(APPNAME, SECTION, "TeclaParaCerrar", "ESC")
        .RutaTempDB = GetSetting(APPNAME, SECTION, "RutaTempDB", RUTATEMP_PRE)
        .PosID = GetSetting(APPNAME, SECTION, "PosID", "C01")
        
        'La clave 'INICIO' sirve para reconocer si es primera vez que ejecuta
        strInicializado = GetSetting(APPNAME, SECTION, "INICIO", "")
    End With
    
        file = "c:\ia\TramaTrans\1.1.4 mprop" & "-" & Format(Date, "ddmmmyyyy") & "-" & Format(Time, "hhmmss") & ".txt"
        numFile = FreeFile
        Open file For Output Access Write As #numFile
        Print #numFile, "mprop"
        Close numFile
    
    'Verifica si todas las informaciones fueron recuperadas.
    band = True
    If Len(mProps.RutaTempDB) = 0 Then band = False
    If Len(mProps.PosID) = 0 Then band = False
    If Len(strInicializado) = 0 Then band = False
    If Not band Then
        Err.Raise ERR_NOREGINFO, "SiiMain.Inicializar"
    End If
    
    'Abre la base principal
        file = "c:\ia\TramaTrans\1.1.5 Abrir" & "-" & Format(Date, "ddmmmyyyy") & "-" & Format(Time, "hhmmss") & ".txt"
        numFile = FreeFile
        Open file For Output Access Write As #numFile
        Print #numFile, "Abrir"
        Close numFile

    Abrir
End Sub


Public Property Let TipoDB(ByVal value As Byte)
    If value > 1 Then Err.Raise ERR_INVALIDO, "SiiMain_TipoDB", MSGERR_INVALIDO
    mProps.TipoDB = value
End Property

Public Property Get TipoDB() As Byte
    TipoDB = mProps.TipoDB
End Property

Public Property Let Ruta(ByVal value As String)
    If (Len(value) = 0) And (mProps.TipoDB = TIPODB_JET) Then
        Err.Raise ERR_INVALIDO, "SiiMain_Ruta", MSGERR_INVALIDO
    End If
    
    If Right$(value, 1) <> "\" And Len(value) > 0 Then value = value & "\"
    mProps.Ruta = value
End Property

Public Property Get Ruta() As String
    Ruta = Trim$(mProps.Ruta)
End Property

Public Property Let NombreDB(ByVal value As String)
    'Si está con la extensión mdb, quita la extensión
    If LCase(Right$(value, 4)) = ".mdb" Then
        value = Left$(value, Len(value) - 4)
    End If
    mProps.NombreDB = value
End Property

Public Property Get NombreDB() As String
    NombreDB = Trim$(mProps.NombreDB)
End Property


Public Property Let NombreDBSQL(ByVal value As String)
    mProps.NombreDBSQL = value
End Property

Public Property Get NombreDBSQL() As String
    NombreDBSQL = Trim$(mProps.NombreDBSQL)
End Property


Public Property Let Server(ByVal value As String)
    mProps.Server = value
End Property

Public Property Get Server() As String
    Server = Trim$(mProps.Server)
End Property

Public Property Let Device(ByVal value As String)
    mProps.Device = value
End Property

Public Property Get Device() As String
    Device = Trim$(mProps.Device)
End Property

Public Property Let UserName(ByVal value As String)
    mProps.UserName = value
End Property

Public Property Get UserName() As String
    UserName = Trim$(mProps.UserName)
End Property

Public Property Let UserNameSQL(ByVal value As String)
    mProps.UserNameSQL = value
End Property

Public Property Get UserNameSQL() As String
    UserNameSQL = Trim$(mProps.UserNameSQL)
End Property

Public Property Let Password(ByVal value As String)
    mProps.Password = value
End Property

Public Property Get Password() As String
    Password = Trim$(mProps.Password)
End Property

Public Property Let PasswordSQL(ByVal value As String)
    mProps.PasswordSQL = value
End Property

Public Property Get PasswordSQL() As String
    PasswordSQL = Trim$(mProps.PasswordSQL)
End Property

Public Property Let FormatoFechaSQL(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "SiiMain.FormatoFechaSQL", MSGERR_INVALIDO
    End If
    mProps.FormatoFechaSQL = value
End Property

Public Property Get FormatoFechaSQL() As String
    FormatoFechaSQL = Trim$(mProps.FormatoFechaSQL)
End Property

Public Property Let TeclaParaCerrar(ByVal value As String)
    mProps.TeclaParaCerrar = value
End Property

Public Property Get TeclaParaCerrar() As String
    TeclaParaCerrar = Trim$(mProps.TeclaParaCerrar)
End Property

Public Property Let RutaTempDB(ByVal value As String)
    If Right$(value, 1) <> "\" And Len(value) > 0 Then value = value & "\"
    mProps.RutaTempDB = value
End Property

Public Property Get RutaTempDB() As String
    RutaTempDB = Trim$(mProps.RutaTempDB)
End Property

Public Property Let PosID(ByVal value As String)
    mProps.PosID = value
End Property

Public Property Get PosID() As String
    PosID = Trim$(mProps.PosID)
End Property

Public Property Get UsuarioActual() As USUARIO
    Set UsuarioActual = gobjUsuarioActual
End Property

Public Property Get GrupoActual() As Grupo
    Set GrupoActual = gobjGrupoActual
End Property

Public Property Get EmpresaActual() As Empresa
    Set EmpresaActual = gobjEmpresaActual
End Property



Public Property Get objCondicion() As Condicion
    Set objCondicion = gobjCond
End Property

#If DAOLIB Then
Friend Property Get Database() As DAO.Database
    Set Database = mdbMain
End Property
#Else
Friend Property Get Coneccion() As ADODB.Connection
    Set Coneccion = mcnMain
    'Set Coneccion = mcnMainAux
End Property
#End If

Friend Property Get ConeccionAux() As ADODB.Connection
    'Set Coneccion = mcnMain
    Set ConeccionAux = mcnMainAux
End Property

Friend Property Get ConeccionReport() As ADODB.Connection
    Set ConeccionReport = mcnMainReport
End Property


Public Property Get ModoDemo() As Boolean
    ModoDemo = mProps.booDEMO
End Property


'Devuelve nombre de ultimo Usuario
Public Property Get UsuarioAnterior() As String
    Dim s As String
    
    s = GetSetting(APPNAME, SECTION, "Usuario_Anterior", "")
    UsuarioAnterior = Trim$(s)
End Property

'Devuelve codigo de la ultima empresa
Public Property Get EmpresaAnterior() As String
    Dim s As String
    
    s = GetSetting(APPNAME, SECTION, "Empresa_Anterior", "")
    EmpresaAnterior = Trim$(s)
End Property


Private Sub Abrir()
    Dim s As String, sr As String
    
    'Si está abierta cierra
    Cerrar
    
    
#If DAOLIB Then
    '*** MAKOTO 18/sep/00 Modificado
'    Set mdbMain = DBEngine.Workspaces(0).OpenDatabase(mProps.Ruta & mProps.NombreDB & ".mdb")
    Set mdbMain = DBEngine.Workspaces(0).OpenDatabase( _
                    mProps.Ruta & mProps.NombreDB & ".mdb", _
                    False, False, _
                    ";UID=" & mProps.UserName & _
                    ";pwd=" & mProps.Password)
#Else
    'Prepara la cadena de coneccion
    If mProps.TipoDB = TIPODB_JET Then
        s = CADENA_CONECCION_JET & mProps.Ruta & mProps.NombreDB & ".mdb"
        
        'Establece la conección con la base principal
        mcnMain.Open s, mProps.UserName, mProps.Password
    Else
'        s = CADENA_CONECCION_SQL & _
'            "Initial Catalog=" & mProps.NombreDB & ";Origen de datos=" & mProps.Server
        s = CADENA_CONECCION_SQL & _
            "Initial Catalog=" & mProps.NombreDBSQL & ";Data Source=" & mProps.Server
        
        'Establece la conección con la base principal en SQLServer
        
        mcnMain.Open s, mProps.UserNameSQL, mProps.PasswordSQL
        
        file = "c:\ia\TramaTrans\1.1.7 Inicializar" & "-" & Format(Date, "ddmmmyyyy") & "-" & Format(Time, "hhmmss") & ".txt"
        numFile = FreeFile
        Open file For Output Access Write As #numFile
        Print #numFile, "sql Abrir"
        Close numFile
        
        
        ' abre coneccion reporte
        
        
'''        CerrarReport
'''        sr = CADENA_CONECCION_SQL & _
'''            "Initial Catalog=ConfigReporte4sql;Data Source=" & mProps.Server
'''
'''
'''        mcnMainReport.Open sr, mProps.UserNameSQL, mProps.PasswordSQL
        
    End If
#End If
End Sub

'Verifica la caducidad y genera error si está caducado
Private Sub VerificarCaducidad()
    Dim fcad As Date, res As Integer, msg As String
    
    fcad = RecuperarFechaCaducidad(res)
    
    'Si está caducado o está invalido
    If ((res = 0) And (fcad < Date)) Or _
        (res = -1) Then
        'No permite entrar al sistema
        Err.Raise ERR_CADUCADO, "SiiMain.FechaCaducidad", MSGERR_CADUCADO
        
    ElseIf res = 0 Then
        'Desde 10 días antes saca mensaje saltando un día
        'Desde 4 días antes saca mansaje todos los días
        ' 10, 8, 6, 4, 3, 2, 1, 0
        If ((fcad <= Date + 10) And ((fcad - Date) Mod 2 = 0)) Or _
            (fcad - Date < 5) Then
            'Sacamos mensaje de advertencia
            msg = "Faltan " & fcad - Date & " día(s) para la fecha de caducidad del sistema : " & vbCr & vbCr & _
                  Format(fcad, "Long Date") & vbCr & vbCr & _
                  "Por favor comuníquese con su proveedor para continuar trabajando."
            MsgBox msg, vbInformation
        End If
    End If
End Sub

'Devuelve fecha de caducidad si está definido que caduca
'Resultado: 1=Nunca, 0=Caduca en la fecha devualta, -1=Error
Public Function RecuperarFechaCaducidad(ByRef resultado As Integer) As Date
    Dim rs As Recordset, sql As String
    Dim Y As Integer, m As Integer, d As Integer
    Dim s As String, fcad As Date
    
    resultado = -1
    
    sql = "SELECT * FROM Version WHERE Version = '0000'"
#If DAOLIB Then
    Set rs = mdbMain.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnMain, adOpenStatic, adLockReadOnly
#End If
    If Not rs.EOF Then
        s = rs.Fields("Comentario")
        
        s = DecifrarTexto(s)
        
        If s = FECHA_NOCADUCA Then
            resultado = 1
        Else
            Y = Val(Mid$(s, 1, 4))
            m = Val(Mid$(s, 6, 2))
            d = Val(Mid$(s, 9, 2))
            If Y > 1998 And Y < 3000 And m > 0 And d > 0 Then
                fcad = DateSerial(Y, m, d)
                resultado = 0
            End If
        End If
    End If
    rs.Close
    
    RecuperarFechaCaducidad = fcad
    
    Set rs = Nothing
End Function

'Graba la fecha de caducidad
'Tiene que saber la clave secreta para llamar éste método
Public Sub GrabarFechaCaducidad(ByVal f As Date, ByVal nunca As Boolean, ByVal key As String)
    Dim s As String, rs As Recordset, sql As String
    
    'Si no sabe la clave del método no permite
    If key <> KEY_CADUCIDAD Then
        Err.Raise ERR_NODERECHO, "SiiMain.FechaCaducidad", MSGERR_NODERECHO
    End If
    
    If nunca Then
        s = FECHA_NOCADUCA
    Else
        s = Format(Year(f), "0000") & "/" & _
            Format(Month(f), "00") & "/" & _
            Format(Day(f), "00")
    End If
    s = CifrarTexto(s)    'Cifra la cadena
    
    sql = "SELECT * FROM Version WHERE Version = '0000'"
#If DAOLIB Then
    Set rs = mdbMain.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnMain, adOpenStatic, adLockReadOnly
#End If
    If rs.EOF Then  'Si no existe, inserta
        sql = "INSERT INTO Version (Version,Fecha,Nombre,Comentario) " & _
              "VALUES ('0000'," & FechaYMD(Date, mProps.TipoDB) & ",'SiiMain','" & s & "')"
    Else                'Si existe, modifica
        sql = "UPDATE Version SET Fecha=" & FechaYMD(Date, mProps.TipoDB) & _
                ", Comentario='" & s & "' " & _
              "WHERE Version = '0000'"
    End If
    rs.Close
    
#If DAOLIB Then
    mdbMain.Execute sql, dbFailOnError
#Else
    mcnMain.Execute sql
#End If
    Set rs = Nothing
End Sub



Public Sub Cerrar()
#If DAOLIB Then
    If Not (mdbMain Is Nothing) Then
        mdbMain.Close
        Set mdbMain = Nothing
    End If
#Else
    'Cierra la conección con la base principal
    If Not (mcnMain Is Nothing) Then
        If mcnMain.State <> adStateClosed Then
            mcnMain.Close
            Set mcnMain = Nothing
        End If
    End If
#End If
    
    'Descarga empresa actual y Usuario actual
    Set gobjEmpresaActual = Nothing
    Set gobjUsuarioActual = Nothing
    Set gobjGrupoActual = Nothing
    
End Sub

Public Sub Grabar()
    If Not (gobjUsuarioActual Is Nothing) Then
        'Si no tiene derecho de supervisor, no puede hacer
        If Not gobjUsuarioActual.BandSupervisor Then
            Err.Raise ERR_NODERECHO, "Empresa_Eliminar", MSGERR_NODERECHO
        End If
    End If
    
    'Graba las informaciones de base principal al registro de sistema
    SaveSetting APPNAME, SECTION, "TipoDB", mProps.TipoDB
    SaveSetting APPNAME, SECTION, "Ruta", mProps.Ruta
    SaveSetting APPNAME, SECTION, "NombreDB", mProps.NombreDB
    SaveSetting APPNAME, SECTION, "NombreDBSQL", mProps.NombreDBSQL
    SaveSetting APPNAME, SECTION, "Server", mProps.Server
    SaveSetting APPNAME, SECTION, "Device", mProps.Device
    SaveSetting APPNAME, SECTION, "UserName", mProps.UserName
    SaveSetting APPNAME, SECTION, "UserNameSQL", mProps.UserNameSQL
    SaveSetting APPNAME, SECTION, "Password", CifrarTexto(mProps.Password)    'Decifrar texto *** MAKOTO 14/sep/00
    SaveSetting APPNAME, SECTION, "PasswordSQL", CifrarTexto(mProps.PasswordSQL)  'Decifrar texto *** MAKOTO 14/sep/00
    SaveSetting APPNAME, SECTION, "FormatoFechaSQL", mProps.FormatoFechaSQL
    SaveSetting APPNAME, SECTION, "TeclaParaCerrar", mProps.TeclaParaCerrar
    SaveSetting APPNAME, SECTION, "RutaTempDB", mProps.RutaTempDB
    SaveSetting APPNAME, SECTION, "PosID", mProps.PosID
    
    'Esto nos sirve para reconocer si es primera ejecución o no.    '*** MAKOTO 08/sep/00
    SaveSetting APPNAME, SECTION, "INICIO", FechaYMD(Date, 0)  'Valor puede ser cualquier cadena
End Sub

Public Function Login(codUsuario As String, Clave As String) As USUARIO
    Dim sql As String, objUsuario As USUARIO, k As String
    Dim rs As Recordset
    
    'Verifica la caducidad y genera error si está caducado  '*** MAKOTO 14/sep/00
''''    VerificarCaducidad
''''    VerificarServidor
    'Obtiene cadena cifrada para comparar con lo que está en la BD  '*** MAKOTO 14/sep/00
    'k = CifrarTexto(Clave, "Usuario")
        file = "c:\ia\TramaTrans\1.4.1 Usuario" & "-" & Format(Date, "ddmmmyyyy") & "-" & Format(Time, "hhmmss") & ".txt"
        numFile = FreeFile
        Open file For Output Access Write As #numFile
        Print #numFile, "Usuario"
        Close numFile
    
        k = CifrarTexto(LCase$(Clave), "Usuario")
    
        file = "c:\ia\TramaTrans\1.4.2 CifrarTexto" & "-" & Format(Date, "ddmmmyyyy") & "-" & Format(Time, "hhmmss") & ".txt"
        numFile = FreeFile
        Open file For Output Access Write As #numFile
        Print #numFile, "CifrarTexto"
        Close numFile
    
    
    sql = "SELECT Clave FROM Usuario " & _
          "WHERE CodUsuario='" & codUsuario & "' AND Clave='" & k & "'"
#If DAOLIB Then
    Set rs = mdbMain.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnMain, adOpenStatic, adLockReadOnly
#End If
    If rs.EOF Then
        Set Login = Nothing
    Else
        '*** MAKOTO 18/jun/2000 No se la razón pero sin esto no funciona bien
        
        file = "c:\ia\TramaTrans\1.4.3 USuario ME" & "-" & Format(Date, "ddmmmyyyy") & "-" & Format(Time, "hhmmss") & ".txt"
        numFile = FreeFile
        Open file For Output Access Write As #numFile
        Print #numFile, "CifrarTexto"
        Close numFile
        
        
        Set gobjMain = Me       'Asigna si mismo a la variable global
        Set objUsuario = New USUARIO
        objUsuario.codUsuario = codUsuario
        
        file = "c:\ia\TramaTrans\1.4.4 Recuperar Usuario" & "-" & Format(Date, "ddmmmyyyy") & "-" & Format(Time, "hhmmss") & ".txt"
        numFile = FreeFile
        Open file For Output Access Write As #numFile
        Print #numFile, "CifrarTexto"
        Close numFile
        
        
        objUsuario.Recuperar
        
        file = "c:\ia\TramaTrans\1.4.5 YA Recuperar Usuario" & "-" & Format(Date, "ddmmmyyyy") & "-" & Format(Time, "hhmmss") & ".txt"
        numFile = FreeFile
        Open file For Output Access Write As #numFile
        Print #numFile, "CifrarTexto"
        Close numFile
        
        
        Set Login = objUsuario                 'Devuelve referencia al Usuario
        Set gobjUsuarioActual = objUsuario    'Establece como Usuario actual en todo DLL
        Set gobjGrupoActual = gobjMain.RecuperaGrupo(objUsuario.CodGrupo)
        Set objUsuario = Nothing
        
        'Graba nombre de Usuario para recuperar cuando inicia de nuevo el programa
        SaveSetting APPNAME, SECTION, "Usuario_Anterior", codUsuario
    End If
    Set rs = Nothing
End Function

Public Function ListaEmpresas(bandArray As Boolean, bandTodas As Boolean) As Variant
    Dim sql As String, rs As Recordset
    
    sql = "SELECT CodEmpresa,Descripcion,TipoDB,Ruta,Server,Device,NombreDB "
    
    If Not bandTodas Then
        sql = sql & " FROM vwMain "
        If Len(gobjUsuarioActual.codUsuario) > 0 Then
            sql = sql & "WHERE CodUsuario='" & gobjUsuarioActual.codUsuario & "' "
        Else
            sql = sql & "ORDER BY CodUsuario"
        End If
    Else
        sql = sql & " FROM Empresa ORDER BY CodEmpresa"
    End If
    
#If DAOLIB Then
    Set rs = mdbMain.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnMain, adOpenStatic, adLockReadOnly
#End If
    
    If bandArray Then
#If DAOLIB Then
        If Not rs.EOF Then
            rs.MoveLast
            rs.MoveFirst
        End If
        ListaEmpresas = rs.GetRows(rs.RecordCount)
#Else
       If Not rs.EOF Then ListaEmpresas = rs.GetRows    'Diego 20/02/2004
#End If
    Else
        Set ListaEmpresas = rs
    End If
    Set rs = Nothing
End Function

Public Function ListaUsuarios() As Variant
    Dim sql As String, rs As Recordset
    
    sql = "SELECT CodUsuario,NombreUsuario,CodGrupo,BandSupervisor, BandValida as Activo " & _
          "FROM Usuario ORDER BY CodUsuario"
#If DAOLIB Then
    Set rs = mdbMain.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnMain, adOpenStatic, adLockReadOnly
#End If
    Set ListaUsuarios = rs
    Set rs = Nothing
End Function

Public Function ListaUsuariosparaCombo() As Variant
    Dim sql As String
    Dim rs As Recordset

    
    sql = "SELECT CodUsuario,NombreUsuario,CodGrupo,BandSupervisor " & _
          "FROM Usuario ORDER BY CodUsuario"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnMain, adOpenStatic, adLockReadOnly
    
    ListaUsuariosparaCombo = MiGetRows(rs)
    rs.Close
    Set rs = Nothing
End Function

Public Function ListaGrupos(bandArray As Boolean) As Variant
    Dim sql As String, rs As Recordset
    
    sql = "SELECT CodGrupo,Descripcion, BandCaja, BandVendedor, BandBodeguero FROM Grupo ORDER BY CodGrupo"
#If DAOLIB Then
    Set rs = mdbMain.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnMain, adOpenStatic, adLockReadOnly
#End If

    If bandArray Then
#If DAOLIB Then
        If Not rs.EOF Then
            rs.MoveLast
            rs.MoveFirst
        End If
        ListaGrupos = rs.GetRows(rs.RecordCount)
#Else
        ListaGrupos = rs.GetRows
#End If
    Else
        Set ListaGrupos = rs
    End If
    Set rs = Nothing
End Function

Public Function ListaPOS() As Variant
    Dim sql As String, rs As Recordset
    
    sql = "SELECT PosID,Descripcion FROM Pos ORDER BY PosID"
#If DAOLIB Then
    Set rs = mdbMain.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnMain, adOpenStatic, adLockReadOnly
#End If
    Set ListaPOS = rs
    Set rs = Nothing
End Function



Public Function ListaGNModulo() As Variant
    Dim v As Variant, temp(1, 8) As String
    
    temp(0, 0) = "GN"
    temp(1, 0) = "General"
    temp(0, 1) = "CT"
    temp(1, 1) = "Contabilidad"
    temp(0, 2) = "TS"
    temp(1, 2) = "Tesorería"
    temp(0, 3) = "IV"
    temp(1, 3) = "Inventario"
    temp(0, 4) = "FC"
    temp(1, 4) = "Facturación"
    temp(0, 5) = "PV"
    temp(1, 5) = "Proveedor"
    temp(0, 6) = "CL"
    temp(1, 6) = "Cliente"
    temp(0, 7) = "AF"
    temp(1, 7) = "Activos fijos"
    temp(0, 8) = "RL"
    temp(1, 8) = "Rol"
    v = temp
    ListaGNModulo = v
End Function

'*** MAKOTO 23/ene/01 Agregado
'Listar códigos de fórmula disponible para la generación de asiento
Public Function ListaGNCodFormula() As String
    Dim s As String
    
    s = "|TSBANCO" & vbTab & "Valor de (Debe-Haber) de bancos"
    s = s & "|TSBANCODEBE" & vbTab & "Valor de Debe de bancos"
    s = s & "|TSBANCOHABER" & vbTab & "Valor de Haber de bancos"
    s = s & "|TSRETACTIVO" & vbTab & "Valor de retención agrupado por cta de Activo"             '*** MAKOTO 08/feb/01 Agregado
    s = s & "|TSRETPASIVO" & vbTab & "Valor de retención agrupado por cta de Pasivo"             '*** MAKOTO 08/feb/01 Agregado
    s = s & "|PCTOTAL" & vbTab & "Total de (Debe + Haber) de docs por cobrar/pagar"
    s = s & "|PCTOTAL2" & vbTab & "Total de (Debe + Haber) de docs por cobrar/pagar (Cuenta 2)"
    s = s & "|PCTOTALDEBE" & vbTab & "Total de Debe de docs. por cobrar/pagar"
    s = s & "|PCTOTALDEBE2" & vbTab & "Total de Debe de docs. por cobrar/pagar (Cuenta 2)"
    s = s & "|PCTOTALHABER" & vbTab & "Total de Haber de docs. por cobrar/pagar"
    s = s & "|PCTOTALHABER2" & vbTab & "Total de Haber de docs. por cobrar/pagar (Cuenta 2)"
    s = s & "|PCPAGO" & vbTab & "Total de pagos (Debe) a cotización del doc. original"
    s = s & "|PCPAGO2" & vbTab & "Total de pagos (Debe) a cotización del doc. original (Cuenta 2)"
    s = s & "|PCCOBRO" & vbTab & "Total de cobros (Haber) a cotizacion del doc. original"
    s = s & "|PCCOBRO2" & vbTab & "Total de cobros (Haber) a cotizacion del doc. original (Cuenta 2)"
    s = s & "|PCXCOBRAR" & vbTab & "Total de docs.x cobrar agrupado por cuenta 2"
    s = s & "|PCXPAGAR" & vbTab & "Total de docs.x pagar agrupado por cuenta 2"
    s = s & "|IVACTIVO" & vbTab & "Total de ABS(CostoRealTotal) de items. Agrupado por cuenta de activo"
    s = s & "|IVACTIVO+" & vbTab & "  Igual que IVACTIVO. Pero sólo de items con signo positivo"
    s = s & "|IVACTIVO-" & vbTab & "  Igual que IVACTIVO. Pero sólo de items con signo negativo"
    s = s & "|IVCOSTO" & vbTab & "Total de ABS(CostoRealTotal) de items. Agrupado por cuenta de costo"
    s = s & "|IVCOSTO+" & vbTab & "  Igual que IVCOSTO. Pero sólo de items con signo positivo"
    s = s & "|IVCOSTO-" & vbTab & "  Igual que IVCOSTO. Pero sólo de items con signo negativo"
    s = s & "|IVVENTA" & vbTab & "Total de ABS(PrecioRealTotal) de items. Agrupado por cuenta de venta"
    s = s & "|IVVENTA2" & vbTab & "Total de ABS(PrecioTotal) de items. Agrupado por cuenta de venta"
    s = s & "|IVVENTA+" & vbTab & "  Igual que IVVENTA. Pero sólo de items con signo positivo"
    s = s & "|IVVENTA-" & vbTab & "  Igual que IVVENTA. Pero sólo de items con signo negativo"
    s = s & "|IVVENTA*" & vbTab & "  Igual que IVVENTA. Pero suma con signos"
    s = s & "|IVCOSTOTOTAL" & vbTab & "Total de ABS(CostoRealTotal) sin diferenciar cuentas"
    s = s & "|IVCOSTOTOTAL+" & vbTab & "  Igual que IVCOSTOTOTAL. Pero sólo de items con signo positivo"
    s = s & "|IVCOSTOTOTAL-" & vbTab & "  Igual que IVCOSTOTOTAL. Pero sólo de items con signo negativo"
    s = s & "|IVPRECIOTOTAL" & vbTab & "Total de ABS(PrecioTotal) sin diferenciar cuentas"
    s = s & "|IVPRECIOTOTAL+" & vbTab & "  Igual que IVPRECIOTOTAL. Pero sólo de items con signo positivo"
    s = s & "|IVPRECIOTOTAL-" & vbTab & "  Igual que IVPRECIOTOTAL. Pero sólo de items con signo negativo"
    s = s & "|IVPRECIOTOTAL*" & vbTab & "  Igual que IVPRECIOTOTAL. Pero suma con signos."
    s = s & "|IVCOSTOTOTAL2" & vbTab & "Total de ABS(CostoRealTotal) + total de recargos/descuentos"
    s = s & "|IVPRECIOTOTAL2" & vbTab & "Total de ABS(PrecioTotal) + total de recargos/descuentos"
    s = s & "|IVDESCITEMTOTAL" & vbTab & "Suma Total de DescuentosxItem (Asignar Cta. Contable)" '***Agregado. 02/ago/2004. Angel
    s = s & "|PRODMATERIAPRIMA" & vbTab & "Total de Materia prima utilizada para Ingreso de Producto Terminado, solo funciona en el SiiProdA" '***Agregado Oliver 09/08/2004 formulas para mod. de produccion.
    s = s & "|PRODMANOOBRA" & vbTab & "Total de Mano de Obra utilizada para Ingreso de Producto Terminado, solo funciona en el SiiProdA"      '***Agregado Oliver 09/08/2004 formulas para mod. de produccion.
    s = s & "|PRODGASTOFABRICACION" & vbTab & "Total de Gastos de Fabricacion utilizada para Ingreso de Producto Terminado, solo funciona en el SiiProdA " '***Agregado Oliver 09/08/2004 formulas para mod. de produccion.
    'ACTIVO FIJO
    s = s & "|AFACTIVO" & vbTab & "Total de ABS(CostoRealTotal) de Activo Fijo. Agrupado por cuenta de activo"
    s = s & "|AFACTIVO+" & vbTab & "  Igual que IVACTIVO. Pero sólo de Activo Fijo con signo positivo"
    s = s & "|AFACTIVO-" & vbTab & "  Igual que IVACTIVO. Pero sólo de Activo Fijo con signo negativo"
    s = s & "|AFDEPREGASTO" & vbTab & "  Total Depre Gasto Activo Fijo. Pero sólo de Activo Fijo "
    s = s & "|AFDEPREACUM" & vbTab & "  Total Depre Acum Activo Fijo. Pero sólo de Activo Fijo "
    s = s & "|AFDEPREGASTOREV" & vbTab & "  Total Depre Gasto Activo Fijo. Pero sólo de Activo Fijo Reval "
    s = s & "|AFDEPREACUMREV" & vbTab & "  Total Depre Acum Activo Fijo. Pero sólo de Activo Fijo Reval"
    'AUC roles
    s = s & "|ROLDEBE" & vbTab & "  Total Debe rol "
    s = s & "|ROLHABER" & vbTab & "  Total Haber rol "
    s = s & "|PCPAGOXIII" & vbTab & "  Total rol Decimo Tercero"
    s = s & "|PCPAGOXIV" & vbTab & "  Total rol Decimo Cuarto"
    s = s & "|PCPAGOVACACION" & vbTab & "  Total rol Vacaciones"
    s = s & "|PCPAGONETOROL" & vbTab & "  Total rol a Debe"
    s = s & "|PCPAGONETOROLH" & vbTab & "  Total rol a Haber"
  'INGRESOS DIFERIDOS
    s = s & "|IVVENTAD" & vbTab & "Total de ABS(PrecioRealTotal) de items. Agrupado por cuenta de ingreso diferido"
    s = s & "|PCCOBROD" & vbTab & "Total de cobros Diferidos. Agrupado por forma de cobro "
    s = s & "|PCVENTAD" & vbTab & "Total de cobros Diferidos(ventas). Agrupado por forma de cobro "
    s = s & "|PCXCOBRARD" & vbTab & "Total de docs.x cobrar Diferidos agrupado por cuenta 2"
    s = s & "|PCXPAGARD" & vbTab & "Total de docs.x pagar Diferidos agrupado por cuenta 2"
    'SEPARA CAPITALINTERES
    s = s & "|PCCOBROCAP" & vbTab & "Total de cobros Diferidos. Agrupado por forma de cobro Solo Capital"
    s = s & "|PCCOBROINT" & vbTab & "Total PC de Haber de docs. Solo Interes Solo Interes"
    s = s & "|PCPAGOINT" & vbTab & "Total PC de Debe de docs. Solo Interes Solo Interes"
    
    s = s & "|PCTOTALDEBECHP" & vbTab & "Total PC de Debe de docs. Solo CHP"
    s = s & "|PCTOTALHABERCHP" & vbTab & "Total PC de Haber de docs. Solo CHP"
    s = s & "|IVAJUSTEDIF" & vbTab & "Ajusta descuadre cta Activo de inventarios (Diferentes ctas ejm transformacion)"
    ListaGNCodFormula = s
End Function


Public Function CreaEmpresa() As Empresa
    Dim nuevo As Empresa
    
    'Si no tiene derecho de supervisor, no puede hacer
    If Not gobjUsuarioActual.BandSupervisor Then
        Err.Raise ERR_NODERECHO, "SiiMain_CreaEmpresa", MSGERR_NODERECHO
    End If
    
    Set nuevo = New Empresa
    Set nuevo.SiiMain = Me
    Set CreaEmpresa = nuevo
    Set nuevo = Nothing
End Function

Public Function RecuperaEmpresa(CodEmpresa As String) As Empresa
    Dim nuevo As Empresa, sql As String
    Dim rs As Recordset
    
    'Debe estar hecho Login para recuperar una empresa
    If gobjMain.UsuarioActual Is Nothing Then
        Err.Raise ERR_NOUsuario, "SiiMain_RecuperaEmpresa", MSGERR_NOUsuario
    End If
    
    'Verifica si el Usuario actual tiene derecho o no
    sql = "SELECT CodEmpresa FROM Permiso " & _
          "WHERE CodGrupo='" & gobjMain.GrupoActual.CodGrupo & "' " & _
          "AND CodEmpresa='" & CodEmpresa & "'"
#If DAOLIB Then
    Set rs = mdbMain.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnMain, adOpenStatic, adLockReadOnly
#End If
    If rs.EOF _
        And (Not gobjUsuarioActual.BandSupervisor) Then     '*** MAKOTO 13/dic/00 Modificado
                                                            'Para que Supervisores pueda recuperar cualquier empresa
        Err.Raise ERR_NODERECHO, "SiiMain_RecuperaEmpresa", MSGERR_NODERECHO
    End If
    rs.Close
    Set rs = Nothing
    
    Set nuevo = New Empresa
    Set nuevo.SiiMain = Me
    nuevo.CodEmpresa = CodEmpresa
    nuevo.Recuperar
    Set RecuperaEmpresa = nuevo
    Set nuevo = Nothing
End Function

Public Function CreaUsuario() As USUARIO
    Dim nuevo As USUARIO
    
    'Si no tiene derecho de supervisor, no puede hacer
    If Not gobjUsuarioActual.BandSupervisor Then
        Err.Raise ERR_NODERECHO, "SiiMain_CreaUsuario", MSGERR_NODERECHO
    End If
    
    Set nuevo = New USUARIO
    Set CreaUsuario = nuevo
    Set nuevo = Nothing
End Function

Public Function RecuperaUsuario(codUsuario As String) As USUARIO
    Dim nuevo As USUARIO
    
    Set nuevo = New USUARIO
    nuevo.codUsuario = codUsuario
    nuevo.Recuperar
    Set RecuperaUsuario = nuevo
    Set nuevo = Nothing
End Function

Public Function CreaGrupo() As Grupo
    Dim nuevo As Grupo
    
    'Si no tiene derecho de supervisor, no puede hacer
    If Not gobjUsuarioActual.BandSupervisor Then
        Err.Raise ERR_NODERECHO, "SiiMain_CreaGrupo", MSGERR_NODERECHO
    End If
    
    Set nuevo = New Grupo
    Set CreaGrupo = nuevo
    Set nuevo = Nothing
End Function

Public Function RecuperaGrupo(CodGrupo As String) As Grupo
    Dim nuevo As Grupo
    
    Set nuevo = New Grupo
    nuevo.CodGrupo = CodGrupo
    nuevo.Recuperar
    Set RecuperaGrupo = nuevo
    Set nuevo = Nothing
End Function



Public Function CreaPos() As pos
    Dim nuevo As pos
    
    'Si no tiene derecho de supervisor, no puede hacer
    If Not gobjUsuarioActual.BandSupervisor Then
        Err.Raise ERR_NODERECHO, "SiiMain_CreaPos", MSGERR_NODERECHO
    End If
    
    Set nuevo = New pos
    Set CreaPos = nuevo
    Set nuevo = Nothing
End Function

Public Function RecuperaPos(PosID As String) As pos
    Dim nuevo As pos
    
    Set nuevo = New pos
    nuevo.PosID = PosID
    nuevo.Recuperar
    Set RecuperaPos = nuevo
    Set nuevo = Nothing
End Function


Private Sub Class_Terminate()
    'Primero cierra la empresa actual
    If Not gobjEmpresaActual Is Nothing Then
        gobjEmpresaActual.Cerrar
    End If
    
    Cerrar
#If DAOLIB Then
    Set mdbMain = Nothing
#Else
    Set mcnMain = Nothing
#End If
    Set gobjCond = Nothing
    Set gobjEmpresaActual = Nothing
    Set gobjUsuarioActual = Nothing
    Set gobjGrupoActual = Nothing
    Set gobjMain = Nothing
End Sub

Public Function RecuperaTransconModificar(id As Long) As Recordset
    Dim sql As String, rs As Recordset
    sql = "SELECT CodTRans FROM PermisoTrans " & _
          "WHERE IdPermiso=" & id & _
             " AND Modificar=1"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnMain, adOpenStatic, adLockReadOnly
    Set RecuperaTransconModificar = rs
    Set rs = Nothing
End Function

'AUC para controlar la clave del usuario cuando graba la transaccion
Public Function LoginVen(codUsuario As String, Clave As String) As Boolean
    Dim sql As String, objVen As FCVendedor, k As String
    Dim rs As Recordset
    
        k = CifrarTexto(LCase$(Clave), "Usuario")
    
    sql = "SELECT Clave FROM Usuario " & _
          "WHERE CodUsuario='" & codUsuario & "' AND Clave='" & k & "'"

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnMain, adOpenStatic, adLockReadOnly

    
'    Set rs = Me.EmpresaActual.OpenRecordset(sql)
    If rs.EOF Then
        LoginVen = False
    Else
        '*** MAKOTO 18/jun/2000 No se la razón pero sin esto no funciona bien
        LoginVen = True 'Devuelve referencia al Usuario
    End If
    Set rs = Nothing
End Function


Public Function LoginSupervisor(codUsuario As String, Clave As String) As Boolean
    Dim sql As String, objUsuario As USUARIO, k As String
    Dim rs As Recordset
    
    LoginSupervisor = False
    'Verifica la caducidad y genera error si está caducado  '*** MAKOTO 14/sep/00'
    'VerificarCaducidad
    
    'Obtiene cadena cifrada para comparar con lo que está en la BD  '*** MAKOTO 14/sep/00
'    k = CifrarTexto(Clave, "Usuario")          '*** MAKOTO 21/sep/00
    k = CifrarTexto(LCase$(Clave), "Usuario")
    
    sql = "SELECT Clave,bandSupervisor FROM Usuario " & _
          "WHERE CodUsuario='" & codUsuario & "' AND Clave='" & k & "'"
#If DAOLIB Then
    Set rs = mdbMain.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnMain, adOpenStatic, adLockReadOnly
#End If

    
    If Not rs.EOF Then
        If rs.Fields("bandSupervisor") = True Then
            LoginSupervisor = True
        End If
        'Graba nombre de Usuario para recuperar cuando inicia de nuevo el programa
        SaveSetting APPNAME, SECTION, "Supervisor_Anterior", codUsuario
    End If
    Set rs = Nothing
End Function


'AUC 30/01/07 para saber si el usuario activo es supervisor
Public Function EsSupervisor(ByRef codUsuario As String) As Boolean
Dim rs As Recordset
Dim sql As String
    sql = "SELECT bandsupervisor FROM Usuario " & _
          "WHERE CodUsuario='" & codUsuario & "' "
   Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnMain, adOpenStatic, adLockReadOnly
    If rs!BandSupervisor = 0 Then
        EsSupervisor = False
    Else
        EsSupervisor = True
   End If
    Set rs = Nothing
End Function


'AUC 31/05/07
Public Function ListaModulos(BandSoloValida As Boolean, BandRS As Boolean) As Variant
'BandRS : True  --> Devuelve Recordset de ADO
'      False --> Devuelve en Array de Variant
    Dim sql As String, rs As Recordset
    sql = "SELECT CodModulo,Descripcion " & _
          "FROM " & gobjMain.NombreDBSQL & ".dbo.Modulo "
    If BandSoloValida Then sql = sql & "WHERE BandValida=" & CadenaBool(True, TipoDB)
    sql = sql & " ORDER BY CodModulo"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
   
    rs.Open sql, mcnMain, adOpenStatic, adLockReadOnly
    If BandRS Then
        Set ListaModulos = rs
    Else
        ListaModulos = MiGetRows(rs)
        rs.Close
    End If
    Set rs = Nothing
End Function

'AUC 31/05/07 verfica si tiene permiso para entrar en este modulo
Public Function PermisoModulo(codUsuario As String, Clave As String, Modulo As String) As Boolean
    Dim sql As String, k As String
    Dim rs As Recordset, Desde As Integer
    On Error GoTo captura
        'k = CifrarTexto(Clave, "Usuario")
        k = CifrarTexto(LCase$(Clave), "Usuario")
    sql = "SELECT modulos FROM Usuario " & _
          "WHERE CodUsuario='" & codUsuario & "' AND Clave='" & k & "'"
#If DAOLIB Then
    Set rs = mdbMain.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnMain, adOpenStatic, adLockReadOnly
#End If
    If rs.EOF Then
        Set rs = Nothing
    Else
        If Not IsNull(rs!modulos) Then
         Desde = Val(InStr(rs!modulos, Modulo))
         
         If Mid(rs!modulos, Desde, Len(Modulo)) = Modulo Then
            PermisoModulo = True
         End If
        End If
    End If
    Set rs = Nothing
    Exit Function
captura:
    PermisoModulo = False
    Exit Function
End Function

Public Function RegresaLogin(codUsuario As String) As USUARIO
    Dim sql As String, objUsuario As USUARIO, k As String
    Dim rs As Recordset
        Set gobjMain = Me       'Asigna si mismo a la variable global
        Set objUsuario = New USUARIO
        objUsuario.codUsuario = codUsuario
        objUsuario.Recuperar
        Set gobjUsuarioActual = objUsuario    'Establece como Usuario actual en todo DLL
        Set gobjGrupoActual = gobjMain.RecuperaGrupo(objUsuario.CodGrupo)
        Set objUsuario = Nothing
    Set rs = Nothing
End Function

Public Function ListaEmpresasMain(bandArray As Boolean, bandTodas As Boolean) As Variant
    Dim sql As String, rs As Recordset
    
    sql = "SELECT Name FROM master.dbo.sysdatabases "
    sql = sql & "where name like '%Main%'ORDER BY Name"
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnMain, adOpenStatic, adLockReadOnly
    
    If bandArray Then
#If DAOLIB Then
        If Not rs.EOF Then
            rs.MoveLast
            rs.MoveFirst
        End If
        ListaEmpresasMain = rs.GetRows(rs.RecordCount)
#Else
       If Not rs.EOF Then ListaEmpresasMain = rs.GetRows    'Diego 20/02/2004
#End If
    Else
        Set ListaEmpresasMain = rs
    End If
    Set rs = Nothing
End Function

Public Function ObtenerUltimaVersionMain(ByVal NombreMain As String) As Long
    
    Dim sql As String, rs As Recordset
    On Error GoTo captura
    sql = "SELECT isnull(Max(Version),0) AS MaxVer "
    sql = sql & "FROM " & NombreMain & ".dbo.VersionMain  "
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnMain, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then
        ObtenerUltimaVersionMain = rs.Fields("MaxVer")
    Else
        ObtenerUltimaVersionMain = 0
    End If
    
    Set rs = Nothing
    Exit Function
captura:
    ObtenerUltimaVersionMain = 0
    Exit Function
    
End Function

Public Sub EjecutarSQLMain( _
                emp As String, ByVal sql As String, _
                ByRef numReg As Long)
    Dim f As Integer, s As String, rs As Recordset

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
'        EjecutarSiiEsquema sql, numReg
        s = ""
    
    'Caso C. Si envia directamente SQL
    Case Else
        s = sql
    End Select

    'En caso A. y C. aquí se ejecuta el SQL
    If Len(s) > 0 Then
        Me.ExecuteMain emp, s, True, numReg
        'rs.Open s, mcnMain, adOpenKeyset, adLockPessimistic
    End If
End Sub


Public Sub ExecuteMain( _
                emp As String, ByVal sql As String, _
                Optional ByVal FailOnError As Boolean, _
                Optional ByRef numReg As Long)  'AUC cambiado public por friend
                Dim s As String
#If DAOLIB Then
    If FailOnError Then
        mdbEmpresa.Execute sql, dbFailOnError
    Else
        mdbEmpresa.Execute sql
    End If
#Else
        s = CADENA_CONECCION_SQL & _
            "Initial Catalog=" & emp & ";Data Source=" & mProps.Server
        
        'Establece la conección con la base principal en SQLServer
'        mcnMainAux.Close
        
    If Not (mcnMainAux Is Nothing) Then
        If mcnMainAux.State <> adStateClosed Then
            mcnMainAux.Close

        End If
    End If
        
        
        
        mcnMainAux.Open s, mProps.UserNameSQL, mProps.PasswordSQL

    mcnMainAux.Execute sql, numReg

#End If
End Sub



Public Function PermisoModuloEspecial(codUsuario As String, Modulo As String) As Boolean
    Dim sql As String, k As String
    Dim rs As Recordset, Desde As Integer
    On Error GoTo captura
        
    sql = "SELECT modulos FROM Usuario " & _
          "WHERE CodUsuario='" & codUsuario & "'"
#If DAOLIB Then
    Set rs = mdbMain.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnMain, adOpenStatic, adLockReadOnly
#End If
    If rs.EOF Then
        Set rs = Nothing
    Else
        If Not IsNull(rs!modulos) Then
         Desde = Val(InStr(rs!modulos, Modulo))
         
         If Mid(rs!modulos, Desde, Len(Modulo)) = Modulo Then
            PermisoModuloEspecial = True
         End If
        End If
    End If
    Set rs = Nothing
    Exit Function
captura:
    PermisoModuloEspecial = False
    Exit Function
End Function

Public Function ListaGNCodFormulaPresupuesto() As String
    Dim s As String
    s = ""
    s = s & "|TOTALDEBE" & vbTab & "Total de Debe de docs. Presupuesto"
    s = s & "|TOTALHABER" & vbTab & "Total de Haber de docs. Presupuesto"
    s = s & "|PCTOTALDEBE" & vbTab & "Total PC de Debe de docs. Presupuesto"
    s = s & "|PCTOTALHABER" & vbTab & "Total PC de Haber de docs. Presupuesto"
    s = s & "|CTTOTALDEBE" & vbTab & "Total Contable de Debe de docs. Presupuesto"
    s = s & "|CTTOTALHABER" & vbTab & "Total Contable de Haber de docs. Presupuesto"
    s = s & "|IVVENTA" & vbTab & "Total de ABS(PrecioRealTotal) de items. Agrupado por cuenta de venta"
    s = s & "|ROLDEBE" & vbTab & "Total elemento rol para debe Presupuesto"
    s = s & "|ROLHABER" & vbTab & "Total elemento rol para haber Presupuesto"
    ListaGNCodFormulaPresupuesto = s
End Function

Public Sub CerrarReport()
    'Cierra la conección con la base principal
    If Not (mcnMainReport Is Nothing) Then
        If mcnMainReport.State <> adStateClosed Then
            mcnMainReport.Close
            Set mcnMainReport = Nothing
        End If
    End If
    
End Sub

Public Function RecuperaEmpresaColas(CodEmpresa As String) As Empresa
    Dim nuevo As Empresa, sql As String
    Dim rs As Recordset
    
        
    'Verifica si el Usuario actual tiene derecho o no
    sql = "SELECT CodEmpresa FROM Permiso " & _
          "WHERE CodGrupo='" & gobjMain.UsuarioAnterior & "' " & _
          "AND CodEmpresa='" & CodEmpresa & "'"
#If DAOLIB Then
    Set rs = mdbMain.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnMain, adOpenStatic, adLockReadOnly
#End If
'    If rs.EOF _
'        And (Not gobjUsuarioActual.BandSupervisor) Then     '*** MAKOTO 13/dic/00 Modificado
'                                                            'Para que Supervisores pueda recuperar cualquier empresa
'        Err.Raise ERR_NODERECHO, "SiiMain_RecuperaEmpresa", MSGERR_NODERECHO
'    End If
    rs.Close
    Set rs = Nothing
    
    Set nuevo = New Empresa
    Set nuevo.SiiMain = Me
    nuevo.CodEmpresa = CodEmpresa
    nuevo.Recuperar
    Set RecuperaEmpresaColas = nuevo
    Set nuevo = Nothing
End Function

Public Function CreaEquipo() As Equipo
    Dim nuevo As Equipo
    
    'Si no tiene derecho de supervisor, no puede hacer
'    If Not gobjUsuarioActual.BandSupervisor Then
'        Err.Raise ERR_NODERECHO, "SiiMain_CreaEquipo", MSGERR_NODERECHO
'    End If
    
    Set nuevo = New Equipo
    Set CreaEquipo = nuevo
    Set nuevo = Nothing
End Function

Public Function RecuperaEquipo(CodEquipo As String) As Equipo
    Dim nuevo As Equipo
    
    Set nuevo = New Equipo
    nuevo.CodEquipo = CodEquipo
    nuevo.Recuperar
    Set RecuperaEquipo = nuevo
    Set nuevo = Nothing
End Function


Public Function ListaEquipos() As Variant
    Dim sql As String, rs As Recordset
    
    sql = "SELECT  CLAVE, Nombreequipo ,UsuarioRed, BandValida as Activo " & _
          "FROM Equipo ORDER BY CodEquipo"
#If DAOLIB Then
    Set rs = mdbMain.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnMain, adOpenStatic, adLockReadOnly
#End If
    Set ListaEquipos = rs
    Set rs = Nothing
End Function

Public Function DecifrarTextoLista(ByVal texto As String) As String
    DecifrarTextoLista = DecifrarTexto(texto, "Equipo")
End Function

Private Sub VerificarServidor()
    Dim fcad As Date, res As Integer, msg As String
    Dim Server As String, nom As String
    
    Server = RecuperarNombreServidor(res)
    
    'Si está caducado o está invalido
    If ((res = 0) And (Server <> mProps.NombreDBSQL)) Or _
        (res = -1) Then
        'No permite entrar al sistema
        Err.Raise ERR_BASEINCORRECTA, "SiiMain.NombreServidor", MSGERR_BASEINCORRECTA
        
    ElseIf res = 0 Then
        'Desde 10 días antes saca mensaje saltando un día
        'Desde 4 días antes saca mansaje todos los días
        ' 10, 8, 6, 4, 3, 2, 1, 0
        If ((fcad <= Date + 10) And ((fcad - Date) Mod 2 = 0)) Or _
            (fcad - Date < 5) Then
            'Sacamos mensaje de advertencia
            msg = "Faltan " & fcad - Date & " día(s) para la fecha de caducidad del sistema : " & vbCr & vbCr & _
                  Format(fcad, "Long Date") & vbCr & vbCr & _
                  "Por favor comuníquese con su proveedor para continuar trabajando."
            MsgBox msg, vbInformation
        End If
    End If
End Sub


'Devuelve fecha de caducidad si está definido que caduca
'Resultado: 1=Nunca, 0=Caduca en la fecha devualta, -1=Error
Public Function RecuperarNombreServidor(ByRef resultado As Integer) As String
    Dim rs As Recordset, sql As String
    Dim Y As Integer, m As Integer, d As Integer
    Dim s As String, fcad As Date, servi As String
    
    resultado = -1
    
    sql = "SELECT * FROM Version WHERE Version = '0000'"
#If DAOLIB Then
    Set rs = mdbMain.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnMain, adOpenStatic, adLockReadOnly
#End If
    If Not rs.EOF Then
        s = rs.Fields("Comentarios")
        If Len(s) = 0 Then
            s = mProps.Server
            GrabarNombreServidor UCase(mProps.Server)
        End If
        s = UCase(DecifrarTexto(s))
        
        If s = UCase(mProps.Server) Then
            resultado = 1
        Else
'            Y = Val(Mid$(s, 1, 4))
'            m = Val(Mid$(s, 6, 2))
'            d = Val(Mid$(s, 9, 2))
'            If Y > 1998 And Y < 3000 And m > 0 And d > 0 Then
'                fcad = DateSerial(Y, m, d)
             servi = s
                resultado = 0
'            End If
        End If
    End If
    rs.Close
    RecuperarNombreServidor = servi
    Set rs = Nothing
End Function

Public Sub GrabarNombreServidor(ByVal Server As String)
    Dim s As String, rs As Recordset, sql As String
    
    'Si no sabe la clave del método no permite
'    If key <> KEY_CADUCIDAD Then
'        Err.Raise ERR_NODERECHO, "SiiMain.FechaCaducidad", MSGERR_NODERECHO
'    End If
    
'    If nunca Then
'        s = FECHA_NOCADUCA
'    Else
'        s = Format(Year(f), "0000") & "/" & _
'            Format(Month(f), "00") & "/" & _
'            Format(Day(f), "00")
'    End If
    s = CifrarTexto(Server)    'Cifra la cadena
    
    sql = "SELECT * FROM Version WHERE Version = '0000'"
#If DAOLIB Then
    Set rs = mdbMain.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnMain, adOpenStatic, adLockReadOnly
#End If
    If rs.EOF Then  'Si no existe, inserta
        sql = "INSERT INTO Version (Version,Fecha,Nombre,Comentarios) " & _
              "VALUES ('0000'," & FechaYMD(Date, mProps.TipoDB) & ",'SiiMain','" & s & "')"
    Else                'Si existe, modifica
        sql = "UPDATE Version SET Fecha=" & FechaYMD(Date, mProps.TipoDB) & _
                ", Comentarios='" & s & "' " & _
              "WHERE Version = '0000'"
        End If
    rs.Close
    
#If DAOLIB Then
    mdbMain.Execute sql, dbFailOnError
#Else
    mcnMain.Execute sql
#End If
    Set rs = Nothing
End Sub


'Devuelve codigo de la empresaActual
Public Property Get EmpresaDefault() As String
    Dim s As String
    
    s = GetSetting(APPNAME, SECTION, "EmpresaDefault", "")
    EmpresaDefault = Trim$(s)
End Property


Public Function LoginEquipo(Clave As String) As Boolean
    Dim sql As String, objUsuario As USUARIO, k As String, mobjEquipo As Equipo, usred As String
    Dim mobjSeg As Sii4Seg.clsSeg, UsuarioRed As String
    Dim rs As Recordset
    
    Set mobjSeg = CreateObject("Sii4Seg.clsSeg")
    
    UsuarioRed = Replace(NombreUsuarioPC, " ", "")
    
    k = CifrarTexto((Clave), "Equipo")
    sql = "SELECT Clave FROM equipo "
    sql = sql & "WHERE bandvalida=1 and Clave='" & k & "'"
    usred = CifrarTexto((UsuarioRed), "Equipo")
    sql = sql & " and UsuarioRed ='" & usred & "'"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, mcnMain, adOpenStatic, adLockReadOnly

    If rs.EOF Then
        LoginEquipo = False
        
        Set mobjEquipo = gobjMain.CreaEquipo
    
    
        mobjEquipo.CodEquipo = Mid$(mobjSeg.PreparaCadenaOriginal + ";" + UsuarioRed, 1, 50)
        mobjEquipo.NombreEquipo = NombrePC
        mobjEquipo.UsuarioRed = UsuarioRed
        mobjEquipo.BandValida = 1
        mobjEquipo.Clave = mobjSeg.PreparaCadenaOriginal
        mobjEquipo.Grabar
        
        Set mobjEquipo = Nothing
        
    Else
        LoginEquipo = True
    End If
    Set rs = Nothing
End Function


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

Public Property Let ModuloCargado(ByVal value As String)
    mProps.ModuloCargado = value
End Property

Public Property Get ModuloCargado() As String
    ModuloCargado = mProps.ModuloCargado
End Property

