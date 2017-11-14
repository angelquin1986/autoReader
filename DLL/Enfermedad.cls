VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Enfermedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    idEnfermedad As Long
    CodEnfermedad As String
    Descripcion As String
   objEmpresa As Empresa
End Type
Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean
Event Refresh(Tipo As Integer)

Public Property Let idEnfermedad(ByVal value As Long)
    mProps.idEnfermedad = value
    mbooModificado = True
End Property

Public Property Get idEnfermedad() As Long
    idEnfermedad = mProps.idEnfermedad
End Property

Friend Property Set objEmpresa(ByVal obj As Empresa)
    Set mProps.objEmpresa = obj
End Property

Public Property Get objEmpresa() As Empresa
    Set objEmpresa = mProps.objEmpresa
End Property

Public Property Let CodEnfermedad(ByVal value As String)
    mProps.CodEnfermedad = value
    mbooModificado = True
End Property

Public Property Get CodEnfermedad() As String
    CodEnfermedad = mProps.CodEnfermedad
End Property

Public Property Let Descripcion(ByVal value As String)
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = mProps.Descripcion
End Property

Public Property Get EsNuevo() As Boolean
    EsNuevo = mbooEsNuevo
End Property

Public Property Get Modificado() As Boolean
    Modificado = mbooModificado
End Property

Private Sub Class_Initialize()
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Sub Recuperar(ByVal key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM FichaEnfermedad WHERE CodEnfermedad='" & key & "'"
        'sql = "SELECT * FROM FichaEnfermedad WHERE CodCuenta='" & key & "'"
    Else
        sql = "SELECT * FROM FichaEnfermedad WHERE IdEnfermedad=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do While Not .EOF
            mProps.idEnfermedad = !idEnfermedad
            'mProps.idEnfermedad = !IdCuenta
            If Not IsNull(!CodEnfermedad) Then mProps.CodEnfermedad = !CodEnfermedad
            'If Not IsNull(!CodCuenta) Then mProps.codEnfermedad = !CodCuenta
            If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
            'If Not IsNull(!NombreCuenta) Then mProps.Descripcion = !NombreCuenta
            
            rs.MoveNext
        Loop
            .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Recuperar01(ByVal key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM FichaEnfermedad WHERE CodCuenta='" & key & "'"
    Else
        sql = "SELECT * FROM FichaEnfermedad WHERE IdEnfermedad=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do While Not .EOF
            mProps.idEnfermedad = !IdCuenta
            If Not IsNull(!CodCuenta) Then mProps.CodEnfermedad = !CodCuenta
            If Not IsNull(!NombreCuenta) Then mProps.Descripcion = !NombreCuenta
            rs.MoveNext
        Loop
            .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Grabar(ByVal idficha As Long)
    Dim sql As String, rs As Recordset
    On Error GoTo Etiq
            
'    If EsNuevo Then
'        mProps.HistoriaClinica = AsignaNumFicha
'    End If
        sql = "SELECT * FROM FichaEnfermedad WHERE idEnfermedad= " & idEnfermedad
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        If rs.RecordCount = 0 Then
            Set rs = Nothing
            sql = "SELECT * FROM Ficha WHERE 1=0"
            Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
            rs.AddNew
        Else
        End If
    
    With rs
            
            !CodEnfermedad = mProps.CodEnfermedad
            !Descripcion = mProps.Descripcion
        
            .Update
            .Move 0
            mProps.idEnfermedad = !idEnfermedad
        .Close
    End With


    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    'Me.objEmpresa.RaiseRefreshEvent REFRESH_FICHA
    Exit Sub
Etiq:
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub


Public Function Eliminar(cod As Long) As Boolean
'    Dim sql As String
'    Dim rs As Recordset
'    On Error GoTo CapturaError
'    'revisamos que no tenga trans relacionada
'    sql = "Select gn.codtrans,gn.numtrans,gn.idcentrohijo from gncomprobante gn "
'    sql = sql & " Inner Join gncentrocostodetalle ccd "
'    sql = sql & "inner join FichaDetalle fd Inner join Ficha f on f.idficha = fd.idficha on fd.id = ccd.idfichadet"
'    sql = sql & " on ccd.idcentrodet = gn.idcentrohijo "
'    sql = sql & " where f.idficha = " & cod
'    Set rs = Me.objEmpresa.OpenRecordset(sql)
'    If rs.RecordCount > 0 Then
'        MsgBox " No puede borrar la ficha tiene registros relacionados" & Chr(13) & rs!CodTrans & rs!NumTrans, vbCritical
'        Eliminar = False
'        Exit Function
'    End If
'
'    'cambia el estado de la habitación
'    sql = "update cc set cc.bandocupado = 0"
'    sql = sql & " from gncentrocosto cc  "
'    sql = sql & " INNER JOIN FichaDetalle fd"
'    sql = sql & " inner join ficha f on f.idficha= fd.idficha"
'    sql = sql & "  on fd.idcentro = cc.idcentro"
'    sql = sql & " Where f.idficha = " & cod
'    objEmpresa.EjecutarSQL sql, 1
'
'    sql = "delete from Ficha where IdFicha= " & cod
'    objEmpresa.EjecutarSQL sql, 1
'    Eliminar = True
'    Me.objEmpresa.RaiseRefreshEvent REFRESH_FICHA
'    Set rs = Nothing
'    Exit Function
'CapturaError:
'    MsgBox Err.Description
'    Set rs = Nothing
'    Eliminar = False
'    Exit Function
End Function



