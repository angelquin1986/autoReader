VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Reglas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROP
    idregla As Long
    Regla As Long
    subRegla As Long
    Condicion As String
    Variable As String
    Operador As String
    Valor As String
    objEmpresa As Empresa
End Type
Private mProps As T_PROP
Private mcolReglas As Collection

Public Sub Grabar(ByVal EsNuevo As Boolean)
    Dim rs As Recordset, sql As String
    'Si no está modificado no hace nada
    'If Not mbooModificado Then Exit Sub
    If EsNuevo Then
        sql = "SELECT * FROM Reglas WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM Reglas WHERE IdRegla =" & mProps.idregla
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
'        If rs!FechaGrabado <> mProps.FechaGrabado Then
'            Err.Raise ERR_MODIFICADO, "PCHistorial_Grabar", MSGERR_MODIFICADO
'        End If
    End If
    With rs
       ' !idRegla = mProps.idRegla
        !Regla = mProps.Regla
        !subRegla = mProps.subRegla
        !Condicion = mProps.Condicion
        !Variable = mProps.Variable
        !Operador = mProps.Operador
        !Valor = mProps.Valor
        .Update
#If DAOLIB Then
        .Bookmark = .LastModified
#End If
        rs.Move 0
        mProps.idregla = !idregla
        .Close
    End With
    Set rs = Nothing
End Sub

Friend Sub Recuperar(id As Long)
    Dim sql As String, rs As Recordset
  
    sql = "SELECT * FROM Reglas WHERE idregla=" & id
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        mProps.idregla = !idregla
        If Not IsNull(!Regla) Then mProps.Regla = !Regla
        If Not IsNull(!subRegla) Then mProps.subRegla = !subRegla
        If Not IsNull(!Condicion) Then mProps.Condicion = !Condicion
        If Not IsNull(!Variable) Then mProps.Variable = !Variable
        If Not IsNull(!Operador) Then mProps.Operador = !Operador
        If Not IsNull(!Valor) Then mProps.Valor = !Valor
        .Close
    End With
    Set rs = Nothing
End Sub

Friend Function Clone() As Reglas
    Dim nuevo As Reglas
    
    Set nuevo = New Reglas
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROP)          'Para el método Clone
    mProps = value
    'mProps.objPC.SetModificado
End Property
Public Property Get idregla() As Long
    idregla = mProps.idregla
End Property
Public Property Let Regla(ByVal value As Long)
    mProps.Regla = value
    'mProps.objPC.SetModificado
End Property
Public Property Get Regla() As Long
    Regla = mProps.Regla
End Property
Public Property Let subRegla(ByVal value As Long)
    mProps.subRegla = value
    'mProps.objPC.SetModificado
End Property

Public Property Get subRegla() As Long
    subRegla = mProps.subRegla
End Property


Public Property Let Condicion(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "Reglas_variable", MSGERR_INVALIDO
    End If
    mProps.Condicion = value
    'mProps.objPC.SetModificado
End Property

Public Property Get Condicion() As String
    Condicion = Trim$(mProps.Condicion)
End Property


Public Property Let Variable(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Reglas_variable", MSGERR_INVALIDO
    End If
    mProps.Variable = value
    'mProps.objPC.SetModificado
End Property

Public Property Get Variable() As String
    Variable = Trim$(mProps.Variable)
End Property

Public Property Let Operador(ByVal value As String)
    If Len(value) > 2 Then
        Err.Raise ERR_INVALIDO, "Reglas_Operador", MSGERR_INVALIDO
    End If
    mProps.Operador = value
    'mProps.objPC.SetModificado
End Property

Public Property Get Operador() As String
    Operador = Trim$(mProps.Operador)
End Property
Public Property Let Valor(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "Reglas_Valor", MSGERR_INVALIDO
    End If
    mProps.Valor = value
    'mProps.objPC.SetModificado
End Property

Public Property Get Valor() As String
    Valor = Trim$(mProps.Valor)
End Property
'Private Sub Class_Terminate()
'    Set mProps.objEmpresa = Nothing
'End Sub

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Function AddRegla() As Long
    Dim nuevo As Reglas
    Set nuevo = New Reglas
    'Set nuevo.Regla = Me
    mcolReglas.Add nuevo
    AddRegla = mcolReglas.Count     'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    'mbooModificado = True
End Function

Public Property Get CountReglas() As Long

    If mcolReglas Is Nothing Then RecuperaReglas Me.idregla
    MsgBox mcolReglas.Count
          CountReglas = mcolReglas.Count
MsgBox mcolReglas.Count
End Property
''AUC 30/01/08
Public Sub RecuperaReglas(ByVal id As Long)
    Dim sql As String, rs As Recordset, i As Long, mod_antes As Boolean

    'mod_antes = mbooModificado      'Guarda el estado de Modificado
    If mcolReglas Is Nothing Then Set mcolReglas = New Collection

    'BorrarIVKardex

    'Recupera los IVKardex
    'sql = "SELECT * FROM Reglas WHERE TransId=" & mProps.TransID & " ORDER BY Orden"
    sql = "SELECT * FROM Reglas where idregla = " & id & " order by regla"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            If !idregla > 0 Then                 'Para asegurar el Bug de ADO
                i = AddRegla
                mReglas(i).Recuperar id

            End If
                .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
'    mbooModificado = mod_antes      'Para que no cambie el estado de Modificado
End Sub

Public Property Get mReglas(ix As Long) As Reglas
    If mcolReglas Is Nothing Then RecuperaReglas (ix)
    Set mReglas = mcolReglas.Item(ix)
End Property

Public Sub RemoveReglas(idregla As Long)
Dim sql As String
If idregla > 0 Then
          sql = "DELETE from reglas where idregla = " & idregla
End If
          gobjMain.EmpresaActual.Execute sql, True, 1
End Sub
Public Function Clave() As String
Dim sql As String
Dim rs As Recordset
On Error GoTo CapturaError
    sql = "select count(idregla) as cuantos from reglas"
    Set rs = Me.Empresa.OpenRecordset(sql)
    Clave = rs!Cuantos
    Set rs = Nothing
    Exit Function
CapturaError:
    MsgBox Err.Description
    Set rs = Nothing
    Exit Function
End Function

