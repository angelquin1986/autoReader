Attribute VB_Name = "modSiiComun2"
Option Explicit
'Modulo común para los proyectos que usan directamente DAO/ADO
'***Para incluir éste modulo, es necesario que está declarada una variable
'    global que se llama gobjMain de tipo SiiMain.
'GetRows común para DAO y ADO
' obj tiene que ser de tipo Recordset
Public Function MiGetRows(ByVal obj As Object) As Variant
    Dim v As Variant
        If Not obj.EOF Then
#If DAOLIB Then
        obj.MoveLast
        obj.MoveFirst
        v = obj.GetRows(obj.RecordCount)
#Else
        v = obj.GetRows
#End If
    End If
    MiGetRows = v
End Function
'Para que haga la conversión de tipo Date a String sin depender de Format()
'que no funciona bien cuando Configuración Regional de Windows está mal
Public Function FechaYMD( _
                    ByVal f As Date, _
                    ByVal TipoDB As Byte, _
                    Optional ByVal Hora As Boolean) As String
    If TipoDB = TIPODB_JET Then
        If Hora Then
            FechaYMD = "#" & Year(f) & "/" & Month(f) & "/" & Day(f) & " " & _
                        Hour(f) & ":" & Minute(f) & ":" & Second(f) & "#"
        Else
            FechaYMD = "#" & Year(f) & "/" & Month(f) & "/" & Day(f) & "#"
        End If
    Else
        If Hora Then
            FechaYMD = "'" & Format(f, gobjMain.FormatoFechaSQL) & " " & _
                        Hour(f) & ":" & Minute(f) & ":" & Second(f) & "'"
        Else
            FechaYMD = "'" & Format(f, gobjMain.FormatoFechaSQL) & "'"
        End If
    End If
End Function

'*** MAKOTO 06/mar/01 Agregado
'Para que haga la conversión de tipo Date a String sin depender de Format()
'que no funciona bien cuando Configuración Regional de Windows está mal
Public Function HoraHMS( _
                    ByVal h As Date, _
                    ByVal TipoDB As Byte, _
                    Optional ByVal SinComilla As Boolean, _
                    Optional ByVal ConFechaIni As Boolean) As String
    HoraHMS = Hour(h) & ":" & Minute(h) & ":" & Second(h)
    If ConFechaIni Then
        If TipoDB = TIPODB_JET Then
            HoraHMS = "1899/12/30" & HoraHMS
        Else
            HoraHMS = "1900/01/01 " & HoraHMS
        End If
    End If
    
    If Not SinComilla Then
        If TipoDB = TIPODB_JET Then
            HoraHMS = "#" & HoraHMS & "#"
        Else
            HoraHMS = "'" & HoraHMS & "'"
        End If
    End If
End Function

Public Function CadenaBool(v As Boolean, TipoDB As Byte) As String
    If TipoDB = TIPODB_JET Then
        CadenaBool = IIf(v, "True", "False")
    Else
        CadenaBool = IIf(v, "1", "0")
    End If
End Function

#If DAOLIB = 0 Then   'Solo en ADO        --------

'Ubica al ultimo Recordset que contiene un objeto Recordset ADO
Public Sub UltimoRecordset(ByRef rs As Object)
    Dim rsTmp As Object
    Do
        Set rsTmp = rs.NextRecordset
        If rsTmp Is Nothing Then Exit Do
        Set rs = rsTmp
    Loop
End Sub

#End If                 '                   -----------


#If DAOLIB = 0 Then   'Solo en ADO        --------

Public Sub VerificaExistenciaTabla(i As Integer)
    Dim rs As Recordset
    Dim sql As String
    'verifica  si la  tabla no esta  creada
    sql = "SELECT * FROM sysobjects WHERE NAME =  'tmp" & i & "'"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not (rs.EOF And rs.BOF) Then
        'elimina la tabla
        gobjMain.EmpresaActual.EjecutarSQL "drop table Tmp" & i, 0
    End If
End Sub

#End If                 '                   -----------



#If DAOLIB = 0 Then   'Solo en ADO        --------

Public Sub VerificaExistenciaTablaenBD(i As Integer, Nombre As Variant)
    Dim rs As Recordset
    Dim sql As String
    'verifica  si la  tabla no esta  creada
    sql = "SELECT * FROM " & Nombre & ".dbo.sysobjects WHERE NAME =  'tmp" & i & "'"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not (rs.EOF And rs.BOF) Then
        'elimina la tabla
        gobjMain.EmpresaActual.EjecutarSQL "drop table " & Nombre & ".dbo.tmp" & i, 0
    End If
End Sub

#End If                 '                   -----------

'JEAA 17/11/2005
Public Sub VerificaExistenciaTablaTemporal(i As Integer)
    Dim rs As Recordset
    Dim sql As String
    'verifica  si la  tabla no esta  creada
    sql = "SELECT * FROM sysobjects WHERE NAME =  't" & i & "'"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not (rs.EOF And rs.BOF) Then
        'elimina la tabla
        gobjMain.EmpresaActual.EjecutarSQL "drop table T" & i, 0
    End If
End Sub


'JEAA 12/06/2006
Public Sub VerificaExistenciaTablaTemp(tabla As String)
    Dim rs As Recordset
    Dim sql As String
    'verifica  si la  tabla no esta  creada
    sql = "SELECT * FROM sysobjects WHERE NAME =  '" & tabla & "'"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not (rs.EOF And rs.BOF) Then
        'elimina la tabla
        gobjMain.EmpresaActual.EjecutarSQL "drop table " & tabla, 0
    End If
End Sub

'Public Function VerificaExistenciaUsuarioenGrupo(cod As String) As Boolean
'    Dim rs As Recordset
'    Dim sql As String
'    VerificaExistenciaUsuarioenGrupo = False
'    'veridfica si existe asignado el grupo a algun usuario
'    sql = "select CodGrupo from usuario where codgrupo='" & cod & "'"
'
'    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
'    If Not rs.EOF Then
'        VerificaExistenciaUsuarioenGrupo = True
'    End If
'    Set rs = Nothing
'End Function

Public Sub VerificaExistenciaTablaTenBD(i As Integer, Nombre As Variant)
    Dim rs As Recordset
    Dim sql As String
    'verifica  si la  tabla no esta  creada
    sql = "SELECT * FROM " & Nombre & ".dbo.sysobjects WHERE NAME =  't" & i & "'"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not (rs.EOF And rs.BOF) Then
        'elimina la tabla
        gobjMain.EmpresaActual.EjecutarSQL "drop table " & Nombre & ".dbo.t" & i, 0
    End If
End Sub

'*** MAKOTO 06/mar/01 Agregado
'Para que haga la conversión de tipo Date a String sin depender de Format()
'que no funciona bien cuando Configuración Regional de Windows está mal
Public Function HoraHMSNew( _
                    ByVal h As Date, _
                    ByVal TipoDB As Byte, _
                    Optional ByVal SinComilla As Boolean, _
                    Optional ByVal ConFechaIni As Boolean) As String
    
    HoraHMSNew = Hour(h) & ":" & Minute(h) & ":" & Second(h)
    
    If ConFechaIni Then
            HoraHMSNew = "30/12/1899 " & HoraHMSNew
    End If
    
    If Not SinComilla Then
        If TipoDB = TIPODB_JET Then
            HoraHMSNew = "#" & HoraHMSNew & "#"
        Else
            HoraHMSNew = "'" & HoraHMSNew & "'"
        End If
    End If
End Function

Public Sub VerificaExistenciaVista(Vista As String)
    Dim rs As Recordset
    Dim sql As String
    'verifica  si la  tabla no esta  creada
    sql = "SELECT * FROM sysobjects WHERE NAME =  '" & Vista & "'"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not (rs.EOF And rs.BOF) Then
        'elimina la tabla
        gobjMain.EmpresaActual.EjecutarSQL "drop VIEW " & Vista, 0
    End If
End Sub

