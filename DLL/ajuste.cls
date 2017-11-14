VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Ajuste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Ajusta la hora de la máquina a la del servidor
Public Sub AjustarAServidor( _
                    ByVal servidor As String, _
                    ByVal uid As String, _
                    ByVal pwd As String)
    Dim con As Connection
    
    Set con = Conectar(servidor, uid, pwd)
    AjustarAServidorCon con
    Set con = Nothing
End Sub

Public Sub AjustarAServidorCon( _
                    ByVal con As Connection)
    Dim hservidor As Date
    
    hservidor = HoraDelServidorCon(con)
    Date = hservidor
    Time = hservidor
End Sub

'Devuelve diferencia de tiempo con el servidor en segundos.
'Utilizo el tipo Currency para que no haya desbordamiento.
Public Function DiferenciaConServidor( _
                    ByVal servidor As String, _
                    ByVal uid As String, _
                    ByVal pwd As String) As Currency
    Dim con As Connection
    
    Set con = Conectar(servidor, uid, pwd)
    DiferenciaConServidor = DiferenciaConServidorCon(con)
    Set con = Nothing
End Function

'Devuelve diferencia de tiempo con el servidor conectado en segundos.
'Utilizo el tipo Currency para que no haya desbordamiento.
Public Function DiferenciaConServidorCon( _
                    ByVal con As Connection) As Currency
    Dim hservidor As Date, hlocal As Date

    hlocal = Now
    hservidor = HoraDelServidorCon(con)
    DiferenciaConServidorCon = DateDiff("s", hservidor, hlocal)
End Function

'Devuelve la hora del servidor especificado
Public Function HoraDelServidor( _
                    ByVal servidor As String, _
                    ByVal uid As String, _
                    ByVal pwd As String) As Date
    Dim con As Connection
    
    Set con = Conectar(servidor, uid, pwd)
    HoraDelServidor = HoraDelServidorCon(con)
    Set con = Nothing
End Function

'Devuelve la hora del servidor conectado
Public Function HoraDelServidorCon( _
            ByVal con As Connection) As Date
    Dim sql As String, rs As Recordset
    
    If (con Is Nothing) Or IsEmpty(con) Then
        Err.Raise vbObjectError + 531, App.EXEName, _
                    "No puede estableser conección con servidor."
    Else
        sql = "SELECT getdate() AS HoraDelServidor"
        Set rs = New Recordset
        rs.Open sql, con, adOpenForwardOnly, adLockReadOnly
        If Not rs.EOF Then
            HoraDelServidorCon = rs.Fields("HoraDelServidor")
        End If
    End If
End Function

'Se conecta al servidor
Private Function Conectar( _
                    ByVal servidor As String, _
                    ByVal uid As String, _
                    ByVal pwd As String) As Connection
    Dim s As String, con As Connection
    
    s = "Provider=SQLOLEDB.1;" & _
        "Persist Security Info=False;" & _
        "User ID=" & uid & ";" & _
        "Password=" & pwd & ";" & _
        "Data Source=" & servidor
    
   
    
    Set con = New Connection
    con.Open s
    
 
    Set Conectar = con
    Set con = Nothing
End Function
