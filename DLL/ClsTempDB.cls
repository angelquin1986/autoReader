VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TempDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Clase para el manejo de base de datos temporal que sirve para mayorización y reporte de balance."
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private mvarPath As String

#If DAOLIB Then
    Private mvarDatabase As DAO.Database
#Else
    Private mvarConeccion As ADODB.Connection
#End If



Public Function Abrir(objMain As SiiMain) As Variant
    Dim n As Integer, nom As String

#If DAOLIB Then
    'Si ya está abierta
    If Not mvarDatabase Is Nothing Then
        Err.Raise ERR_TEMPDB, "TempDB_Abrir", MSGERR_TEMPDB
    End If
#Else
    'Si ya está abierta
    If mvarConeccion.State <> adStateClosed Then
        Err.Raise ERR_TEMPDB, "TempDB_Abrir", MSGERR_TEMPDB
    End If
#End If
    
    mvarPath = objMain.RutaTempDB
    If mvarPath = "" Then
        mvarPath = App.Path
        If Right$(mvarPath, 1) <> "\" Then mvarPath = mvarPath & "\"
    End If

    'Genera un nombre que NO esté utilizado todavía
    Do
        n = n + 1
        nom = NOMBREDB_TEMP & Format(n, "000") & ".mdb"
    Loop While ExisteArchivo(mvarPath & nom)
    
    mvarPath = mvarPath & nom

    'Crea el nuevo archivo MDB copiando de Temp_Ori.mdb
    FileCopy objMain.Ruta & "Temp_Ori.mdb", mvarPath
    
    'Abre la base temporal con ADO
#If DAOLIB Then
    Set mvarDatabase = gobjWS.OpenDatabase(mvarPath)
    Set Abrir = mvarDatabase
#Else
    mvarConeccion.Open CADENA_CONECCION_JET & mvarPath
    Set Abrir = mvarConeccion
#End If
End Function

Public Sub Cerrar()
#If DAOLIB Then
    On Error Resume Next
    mvarDatabase.Close
    Err.Clear
    On Error GoTo 0
#Else
    'Cierra la base
    If mvarConeccion.State <> adStateClosed Then
        mvarConeccion.Close
    End If
#End If
    
    On Error Resume Next        'Para evitar error de Kill
    'Elimina el archivo MDB
    Kill Path
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub Class_Initialize()
#If DAOLIB Then
#Else
    Set mvarConeccion = New ADODB.Connection
#End If
End Sub

Private Sub Class_Terminate()
#If DAOLIB Then
    If Not mvarDatabase Is Nothing Then Cerrar
    Set mvarDatabase = Nothing
#Else
    'Cierra la base
    If mvarConeccion.State <> adStateClosed Then
        Cerrar
    End If
    Set mvarConeccion = Nothing
#End If
End Sub

Private Function ExisteArchivo(Archi As String) As Boolean
    Dim n As Integer
    On Error GoTo errtrap   'Necesario siempre
    
    n = FreeFile
    Open Archi For Input As #n
    Close n
    ExisteArchivo = True
    Exit Function
    
errtrap:
    Exit Function
End Function


Public Property Get Path() As String
    Path = Trim$(mvarPath)
End Property

#If DAOLIB Then
Public Property Get Database() As DAO.Database
    Set Database = mvarDatabase
End Property
#Else
Public Property Get Coneccion() As ADODB.Connection
    Set Coneccion = mvarConeccion
End Property
#End If

