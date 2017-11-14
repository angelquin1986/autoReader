Attribute VB_Name = "modMain"
Option Explicit
'Constantes para todo DLL

Public Const CADENA_CONECCION_JET = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source="        'Se concatena la ruta y nombre del archivo MDB
Public Const CADENA_CONECCION_SQL = "Provider=SQLOLEDB.1;Persist Security Info=False;"   'User ID=sa;Initial Catalog=sysIA;Data Source=iadc01"


Public Const NOMBREDB_ORIGEN_JET = "Origin.mdb"  'Nombre de archivo de origen para crear nueva empresa
Public Const ARCHIVO_SQL = "CreaTablas.sql"     'Nombre de archivo para crear empresa en SQLSvr
Public Const NOMBREDB_TEMP = "_Tmp#"   'Nombre de archivo de base temporal a este se agrega numero unico de instancia

'Variable global dentro de DLL
Public gobjMain As SiiMain
Public gobjUsuarioActual As USUARIO   'Usuario actual (Usuario que hizo Login)
Public gobjGrupoActual As Grupo       'Grupo en el que pertenece el usuario actual
Public gobjEmpresaActual As Empresa   'Empresa actual
Public gobjCond As Condicion

#If DAOLIB Then
Public gobjWS As Workspace
#End If

Public Sub Main()
#If DAOLIB Then
    'Guarda la referencia al espacio de trabajo
    Set gobjWS = DBEngine.Workspaces(0)
#End If
End Sub

#If DAOLIB Then
Public Function GetStringDAO( _
                    rs As Recordset, _
                    tam_fila As Long) As String
    Dim s As String, i As Integer, lon As Integer, numr As Long
    Dim pos As Long, lf As Integer, f As DAO.Field
    
    With rs
        'Si no tiene nada no hace nada
        If .EOF Then
            GetStringDAO = " "      '*** MAKOTO 07/ago/2000 Modificado
            Exit Function
        End If

        'Obtiene longitud de cada registro
        ' y calcula número de caracteres para almacenar una fila
        lon = tam_fila + 1 + .Fields.Count
'        For Each f In .Fields
'            lon = lon + f.Size + 1     '+1 es para vbTab
'        Next f
'        lon = lon + 1           '+1 es para el signo "|" que separa las filas

        'Carga todos los registros a la memoria para saber numero de registros
        .MoveLast
        numr = .RecordCount
        .MoveFirst

        'Prepara el espacio para todos los registros en la memoria
        s = Space$(numr * lon)

        pos = 1
        Do Until .EOF
            If pos > 1 Then
                Mid$(s, pos, 1) = "|"
                pos = pos + 1
            End If

            For i = 0 To .Fields.Count - 1
                lf = Len(.Fields(i).value)
                Mid$(s, pos, lf) = .Fields(i).value
                pos = pos + lf
                
                If i < .Fields.Count - 1 Then
                    Mid$(s, pos, 1) = vbTab
                    pos = pos + 1
                End If
            Next i
            .MoveNext
        Loop
    End With

    GetStringDAO = RTrim$(s)
End Function
#End If

'Obtiene el valor maximo de un campo de una tabla
#If DAOLIB Then
Public Function ObtieneMax(db As Database, tabla As String, campo As String) As Variant
#Else
Public Function ObtieneMax(cn As Connection, tabla As String, campo As String) As Variant
#End If
    Dim sql As String, rs As Recordset
    
    sql = "SELECT Max(" & campo & ") AS MaxValue FROM " & tabla
    
#If DAOLIB Then
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, cn, adOpenStatic, adLockReadOnly
#End If
    With rs
        If Not .EOF Then
            If Not IsNull(!MaxValue) Then
                ObtieneMax = !MaxValue
            End If
        End If
        .Close
    End With
    Set rs = Nothing
End Function

Public Sub ValidaCodigo(ByVal cod As String)
    Dim s As String, i As Integer, pos As Integer, cad As String
    
'    s = "*_?=&{}!#$',[]|()\;:%" & Chr$(34) '& Chr$(32)        'Caracteres no validos para códigos
    
    For i = 1 To Len(s)
        If InStr(cod, Mid$(s, i, 1)) > 0 Then
'                Select Case i
'                    Case 5
'                        pos = InStr(1, cod, Mid$(s, i, 1), vbBinaryCompare)
'                        cad = Mid(cod, 1, pos - 1) & "Y" & Mid(cod, pos + 1, Len(cod))
'                    End Select
                
'            Err.Raise ERR_INVALIDO, "ValidaCodigo", _
                "La letra '" & Mid$(s, i, 1) & "' NO es válida para Códigos/Descripciones/Nombres"
                MsgBox "La letra '" & Mid$(s, i, 1) & "' NO es válida para Códigos/Descripciones/Nombres"
        End If
    Next i
        
End Sub

'Cifra cualquier texto
Public Function CifrarTexto(ByVal s As String, Optional ByVal key As String) As String
    Dim obj As clsCifrar
    
    If Len(key) = 0 Then key = KEY_CADUCIDAD
    Set obj = New clsCifrar
    CifrarTexto = obj.Cifrar(s, key)
    
    Set obj = Nothing
End Function

'Decifra texto cifrado con CifrarTexto()
Public Function DecifrarTexto(ByVal s As String, Optional ByVal key As String) As String
    Dim obj As clsCifrar
    
    If Len(key) = 0 Then key = KEY_CADUCIDAD
    Set obj = New clsCifrar
    DecifrarTexto = obj.Decifrar(s, key)
    
    Set obj = Nothing
End Function

'Agregado Alex: Sept/2002
Public Function VerificaDocumento(ByVal RUC As String) As Boolean
    'verifica  cedula o ruc
    'tipo:  1 = RUC
    '       2 = Cedula
    Const NATURAL = 1
    Const JURIDICA = 2
    Const PUBLICA = 3
    
    Dim i As Integer, m As Integer, Tipo As Byte
    Dim mul As Integer, suma As Integer, v As Integer, ultimo As Integer
    Dim numOriginal As Integer
    Dim Vector(1 To 9) As Integer
    'jeaa 24/09/04 en caso de que sea consumidor final
    ' es un caso especial para el RUC de consumidor final
    If RUC = "9999999999999" Then
        VerificaDocumento = True
        Exit Function
    End If
    If Not (Len(RUC) = 13 Or Len(RUC) = 10) Then
      VerificaDocumento = False
      Exit Function
    End If
    'Compara  si todos  son  repetidos
    Dim bandDiferente  As Boolean
    bandDiferente = False
    For i = 1 To Len(RUC)
        If Left(RUC, 1) <> Mid(RUC, i, 1) Then
            bandDiferente = True
            Exit For
        End If

    Next i
    If bandDiferente = False Then
        VerificaDocumento = False
        Exit Function
    End If
    
    
    If Len(RUC) = 13 Then    'RUC
        If Right(RUC, 3) <> "001" Then   'revisar  las condiciones en documento
            VerificaDocumento = False
            Exit Function
        End If
    End If
    
    
    RUC = Left(RUC, 10)  'solo saca los 10 primeros caracteres
    If Mid(RUC, 3, 1) < 6 Then
        Tipo = NATURAL
    ElseIf Mid(RUC, 3, 1) = 9 Then
        Tipo = JURIDICA
    ElseIf Mid(RUC, 3, 1) = 6 Then
        Tipo = PUBLICA
    End If
    Select Case Tipo
    Case NATURAL
        mul = 2
        For i = 1 To Len(RUC) - 1
            m = Val(Mid$(RUC, i, 1))
            v = m * mul
            If v >= 10 Then v = v - 9
            suma = suma + v
            If mul = 2 Then mul = 1 Else mul = 2
        Next i
        If (suma Mod 10) = 0 Then
          ultimo = 0
        Else
          ultimo = 10 - (suma Mod 10)
        End If
        numOriginal = Val(Right$(RUC, 1))
    Case JURIDICA
        Vector(1) = 4: Vector(2) = 3: Vector(3) = 2: Vector(4) = 7    '0    ,3,2,7,6,5,4,3,2)
        Vector(5) = 6: Vector(6) = 5: Vector(7) = 4: Vector(8) = 3: Vector(9) = 2
        For i = 1 To Len(RUC) - 1
            m = Val(Mid$(RUC, i, 1))
            v = m * Vector(i)
            suma = suma + v
        Next i
        If (suma Mod 11) = 0 Then
            ultimo = 0
            
        Else
            ultimo = 11 - (suma Mod 11)
        End If
        numOriginal = Val(Right$(RUC, 1))
    Case PUBLICA
        Vector(1) = 3: Vector(2) = 2: Vector(3) = 7: Vector(4) = 6    '0    ,3,2,7,6,5,4,3,2)
        Vector(5) = 5: Vector(6) = 4: Vector(7) = 3: Vector(8) = 2
        For i = 1 To Len(RUC) - 2   ' solo   8 caracteres
            m = Val(Mid$(RUC, i, 1))
            v = m * Vector(i)
            suma = suma + v
        Next i
        If (suma Mod 11) = 0 Then
            ultimo = 0
        Else
            ultimo = 11 - (suma Mod 11)
        End If
        numOriginal = Val(Mid$(RUC, 9, 1))
    Case Else
        VerificaDocumento = False
        Exit Function
    End Select
    
    If ultimo = numOriginal Then
        VerificaDocumento = True
    Else
        VerificaDocumento = False
    End If
End Function

Public Function PreparaCadena(ByVal cadena As String) As String
'Funcion que concatena apostrofes en una cadena separada por comas
Dim v As Variant, max As Integer, i As Integer
Dim Respuesta As String
    If cadena = "" Then
        PreparaCadena = "''"
        Exit Function
    End If
    v = Split(cadena, ",")
    max = UBound(v, 1)
    For i = 0 To max
        Respuesta = Respuesta & "'" & v(i) & "'" & ","
    Next i
    Respuesta = Left(Respuesta, Len(Respuesta) - 1) 'Quita la útima coma
    PreparaCadena = Respuesta
    
End Function

'Obtiene el valor maximo de un campo de una tabla
#If DAOLIB Then
Public Function ObtieneMaxconNivel(db As Database, tabla As String, campo As String) As Variant

End Function
#Else
Public Function ObtieneMaxconNivel(cn As Connection, tabla As String, campo As String, Nivel As Integer) As Variant
#End If
    Dim sql As String, rs As Recordset
    
    sql = "SELECT Max(" & campo & ") AS MaxValue FROM " & tabla
    
#If DAOLIB Then
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, cn, adOpenStatic, adLockReadOnly
#End If
    With rs
        If Not .EOF Then
            If Not IsNull(!MaxValue) Then
                ObtieneMaxconNivel = !MaxValue
            End If
        End If
        .Close
    End With
    Set rs = Nothing
End Function

Public Function ObtieneMaxconNivel2(tabla As String, campo As String, Nivel As Integer) As Variant
    Dim sql As String, rs As Recordset
    sql = "SELECT Max(" & campo & ") AS MaxValue FROM " & tabla
    Set rs = New ADODB.Recordset
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    With rs
        If Not .EOF Then
            If Not IsNull(!MaxValue) Then
                ObtieneMaxconNivel2 = !MaxValue
            End If
        End If
        .Close
    End With
    Set rs = Nothing
End Function

Public Function VerificaDocumentoNew(ByVal RUC As String) As Boolean
    'verifica  cedula o ruc
    'tipo:  1 = RUC
    '       2 = Cedula
    Const NATURAL = 1
    Const JURIDICA = 2
    Const PUBLICA = 3
    
    Dim i As Integer, m As Integer, Tipo As Byte
    Dim mul As Integer, suma As Integer, v As Integer, ultimo As Integer
    Dim numOriginal As Integer
    Dim Vector(1 To 9) As Integer
    'jeaa 24/09/04 en caso de que sea consumidor final
    ' es un caso especial para el RUC de consumidor final
    If RUC = "9999999999999" Then
        VerificaDocumentoNew = True
        Exit Function
    End If
    If Not (Len(RUC) = 13 Or Len(RUC) = 10) Then
      VerificaDocumentoNew = False
      Exit Function
    End If
    'Compara  si todos  son  repetidos
    Dim bandDiferente  As Boolean
    bandDiferente = False
    For i = 1 To Len(RUC)
        If Left(RUC, 1) <> Mid(RUC, i, 1) Then
            bandDiferente = True
            Exit For
        End If

    Next i
    If bandDiferente = False Then
        VerificaDocumentoNew = False
        Exit Function
    End If
    
    
    If Len(RUC) = 13 Then    'RUC
        If Right(RUC, 3) <> "001" Then   'revisar  las condiciones en documento
            VerificaDocumentoNew = False
            Exit Function
        End If
    End If
    
    
    RUC = Left(RUC, 10)  'solo saca los 10 primeros caracteres
    If Mid(RUC, 3, 1) < 7 Then
        Tipo = NATURAL
    ElseIf Mid(RUC, 3, 1) = 9 Then
        Tipo = JURIDICA
    ElseIf Mid(RUC, 3, 1) = 6 Then
        Tipo = PUBLICA
    End If
    Select Case Tipo
    Case NATURAL
        mul = 2
        For i = 1 To Len(RUC) - 1
            m = Val(Mid$(RUC, i, 1))
            v = m * mul
            If v >= 10 Then v = v - 9
            suma = suma + v
            If mul = 2 Then mul = 1 Else mul = 2
        Next i
        If (suma Mod 10) = 0 Then
          ultimo = 0
        Else
          ultimo = 10 - (suma Mod 10)
        End If
        numOriginal = Val(Right$(RUC, 1))
    Case JURIDICA
        Vector(1) = 4: Vector(2) = 3: Vector(3) = 2: Vector(4) = 7    '0    ,3,2,7,6,5,4,3,2)
        Vector(5) = 6: Vector(6) = 5: Vector(7) = 4: Vector(8) = 3: Vector(9) = 2
        For i = 1 To Len(RUC) - 1
            m = Val(Mid$(RUC, i, 1))
            v = m * Vector(i)
            suma = suma + v
        Next i
        If (suma Mod 11) = 0 Then
            ultimo = 0
            
        Else
            ultimo = 11 - (suma Mod 11)
        End If
        numOriginal = Val(Right$(RUC, 1))
    Case PUBLICA
        Vector(1) = 3: Vector(2) = 2: Vector(3) = 7: Vector(4) = 6    '0    ,3,2,7,6,5,4,3,2)
        Vector(5) = 5: Vector(6) = 4: Vector(7) = 3: Vector(8) = 2
        For i = 1 To Len(RUC) - 2   ' solo   8 caracteres
            m = Val(Mid$(RUC, i, 1))
            v = m * Vector(i)
            suma = suma + v
        Next i
        If (suma Mod 11) = 0 Then
            ultimo = 0
        Else
            ultimo = 11 - (suma Mod 11)
        End If
        numOriginal = Val(Mid$(RUC, 9, 1))
    Case Else
        VerificaDocumentoNew = False
        Exit Function
    End Select
    
    If ultimo = numOriginal Then
        VerificaDocumentoNew = True
    Else
        VerificaDocumentoNew = False
    End If
End Function

