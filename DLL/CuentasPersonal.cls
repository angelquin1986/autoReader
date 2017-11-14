VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CuentasPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdEmpleado As Long
    idElemento As Long
    IdCuenta As Long
    FechaGrabado As Date
    objPersonal As Personal
End Type

Private mProps As T_PROPS
Private mbooModificado As Boolean
Private mbooEsNuevo As Boolean

Friend Property Set Personal(ByVal obj As Personal)
    Set mProps.objPersonal = obj
End Property

Public Property Get Personal() As Personal
    Set Personal = mProps.objPersonal
End Property

Public Property Let IdEmpleado(ByVal value As Long)
    mProps.IdEmpleado = value
    mbooModificado = True
End Property

Public Property Get IdEmpleado() As Long
    IdEmpleado = mProps.IdEmpleado
End Property

Public Property Let idElemento(ByVal value As Long)
    mProps.idElemento = value
    mbooModificado = True
End Property

Public Property Get idElemento() As Long
    idElemento = mProps.idElemento
End Property

Public Property Let IdCuenta(ByVal value As Long)
    If Not (IsNumeric(value)) Then Err.Raise ERR_INVALIDO, "CuentasPersonal.IdCuenta", MSGERR_INVALIDO
    mProps.IdCuenta = value
    mbooModificado = True
End Property

Public Property Get IdCuenta() As Long
    IdCuenta = mProps.IdCuenta
End Property

Public Property Let FechaGrabado(ByVal value As Date)
    If Not (IsDate(value)) Then Err.Raise ERR_INVALIDO, "CuentasPersonal.FechaGrabado", MSGERR_INVALIDO
    mProps.FechaGrabado = value
    mbooModificado = True
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
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objPersonal = Nothing
End Sub

'*** METODOS
Public Sub Grabar()
    Dim sql As String, rs As Recordset
    On Error GoTo Etiq
    
        'Me.Empresa.Coneccion.BeginTrans
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient 'Evitar el bug de ADO
    
    If mbooEsNuevo Then
'            sql = "Select * FROM CuentaPersonal WHERE idEmpleado = " & Me.IdEmpleado & _
                                                       " And idElemento = " & Me.idElemento
'            sql = " Delete CuentaPersonal WHERE idEmpleado = " & Me.IdEmpleado & " And idElemento = " & Me.idElemento
 '            Me.Personal.Empresa.EjecutarSQL sql, 1
'            If Not (rs.EOF) Then
'                Err.Raise ERR_YAEXISTECODIGO, "CuentaPersonal.Grabar", MSGERR_YAEXISTECODIGO
'            Else
'                rs.Close
'            End If
            
            
            sql = "SELECT * FROM CuentaPersonal WHERE 1=0"
            Set rs = Me.Personal.Empresa.OpenRecordsetParaEdit(sql)
    
        rs.AddNew
        rs!IdEmpleado = Me.IdEmpleado
    Else
        sql = "Select * FROM CuentaPersonal WHERE idEmpleado = '" & Me.IdEmpleado & "'"
        
            Set rs = Me.Personal.Empresa.OpenRecordsetParaEdit(sql)
            'ADO no tiene rs.edit
        
    End If
    With rs
        !idElemento = Me.idElemento
        !IdCuenta = Me.IdCuenta
        !FechaGrabado = Now
        .Update
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    Exit Sub
    
Etiq:
    
    Err.Raise Err.Number, Err.Source, Err.Description
    Exit Sub
End Sub

Public Sub Eliminar()
'    Dim sql As String, rs As Recordset
'
'    #If DAOLIB Then
'        sql = "DELETE * FROM CuentaPersonal WHERE CodEmpleado = '" & Me.CodEmpleado & "'" & _
'                                                   " And IdCuenta = " & Me.IdCuenta
'        Me.Empresa.Database.Execute sql, dbFailOnError
'    #Else
'        sql = "DELETE FROM CuentaPersonal WHERE CodEmpleado = '" & Me.CodEmpleado & "'" & _
'                                                   " And IdCuenta = " & Me.IdCuenta
'        Set rs = New ADODB.Recordset
'        rs.CursorLocation = adUseClient
'        Me.Personal.Empresa.Coneccion.Execute sql
'    #End If
'    mbooEsNuevo = True
'    mbooModificado = True
'    'Me.Empresa.RaiseRefresh REFRESH_CARGO
End Sub






