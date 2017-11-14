VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CuentasDepartamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    idPcGrupo As Long
    idElemento As Long
    IdCuenta As Long
    CodCuenta As String
    FechaGrabado As Date
    objPCG As PcGrupo
End Type

Private mProps As T_PROPS
Private mbooModificado As Boolean
Private mbooEsNuevo As Boolean

Friend Property Set objPCG(ByVal obj As PcGrupo)
    Set mProps.objPCG = obj
End Property

Public Property Get objPCG() As PcGrupo
    Set Dep = mProps.objPCG
End Property
    
Public Property Let idPcGrupo(ByVal value As Long)
    mProps.idPcGrupo = value
    mbooModificado = True
End Property

Public Property Get idPcGrupo() As Long
    idPcGrupo = mProps.idPcGrupo
End Property

Public Property Let idElemento(ByVal value As Long)
    mProps.idElemento = value
    mbooModificado = True
End Property

Public Property Get idElemento() As Long
    idElemento = mProps.idElemento
End Property

Public Property Let IdCuenta(ByVal value As Long)
    mProps.IdCuenta = value
    mbooModificado = True
End Property

Public Property Get IdCuenta() As Long
    IdCuenta = mProps.IdCuenta
End Property


Public Property Let FechaGrabado(ByVal value As Date)
    If Not (IsDate(value)) Then Err.Raise ERR_INVALIDO, "CuentasPCGrupo.FechaGrabado", MSGERR_INVALIDO
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
    'Set mProps.objDep = Nothing
End Sub

'*** METODOS
Public Sub Grabar()
'    Dim sql As String, rs As Recordset
'    On Error GoTo Etiq
'        'Me.Empresa.Coneccion.BeginTrans
'        Set rs = New ADODB.Recordset
'        rs.CursorLocation = adUseClient 'Evitar el bug de ADO
'
'    If mbooEsNuevo Then
'            sql = "Select * FROM CuentaPcGrupo WHERE idPcGrupo = '" & Me.idPcGrupo & "'" & _
'                                                       " And idElemento = '" & Me.idElemento & "'"
'            rs.Open sql, Me.Dep.Empresa.Coneccion, adOpenKeyset, adLockPessimistic
'            If Not (rs.EOF) Then
'                'Err.Raise ERR_YAEXISTECODIGO, "CuentasDepartamento.Grabar", MSGERR_YAEXISTECODIGO
'            Else
'                rs.Close
'            End If
'            sql = "SELECT * FROM CuentaPcGrupo WHERE 1=0"
'            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
'
'        rs.AddNew
'        rs!idPcGrupo = Me.idPcGrupo
'    Else
'        sql = "Select * FROM CuentaPcGrupo WHERE idpcgrupo = '" & Me.Idpcidgrupo & "'"
'            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
'    End If
'    With rs
'        !idElemento = Me.idElemento
'        !IdCuenta = Me.IdCuenta
'        !FechaGrabado = Now
'        .Update
'        .Close
'    End With
'    Set rs = Nothing
'    mbooEsNuevo = False
'    mbooModificado = False
'    Exit Sub
'Etiq:
'    Err.Raise Err.Number, Err.Source, Err.Description
'    Exit Sub
End Sub


Friend Property Set Dep(ByVal obj As PcGrupo)
    'Set mProps.objDep = obj
End Property

Public Property Get Dep() As PcGrupo
  '  Set Dep = mProps.objDep
End Property


Public Property Let IdDepartamento(ByVal value As Long)
'    mProps.IdDepartamento = value
'    mbooModificado = True
End Property

Public Property Get IdDepartamento() As Long
'    IdDepartamento = mProps.IdDepartamento
End Property

Public Property Let codDepartamento(ByVal value As String)
'    Dim sql As String, rs As Recordset
'    Dim idDep As Long
'    idDep = gobjMain.EmpresaActual.GNOpcion.ObtenerValor("Departamento") + 1
'
'    'Cuando cambia el valor
'    If value <> mProps.codDepartamento Then
'        If Len(value) > 0 Then
'            'Actualiza IdCuentaVenta también
'            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
'            Set rs = mProps.objDep.Empresa.OpenRecordset(sql)
'            If rs.EOF Then
'                Err.Raise ERR_NOHAYCODIGO, "IVInventario_CodCuentaVenta", MSGERR_NOHAYCODIGO
'                Exit Property
'            Else
'                IdCuenta = rs!IdCuenta
'            End If
'            rs.Close
'        Else
'            IdCuenta = 0
'        End If
'
'        mProps.CodCuenta = value
'        mbooModificado = True
'    End If
'
'    Set rs = Nothing
End Property


Public Property Let CodCuenta(ByVal value As String)
'    Dim sql As String, rs As Recordset
'
'    'Cuando cambia el valor
'    If value <> mProps.CodCuenta Then
'        If Len(value) > 0 Then
'            'Actualiza IdCuentaVenta también
'            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
'            Set rs = mProps.objDep.Empresa.OpenRecordset(sql)
'            If rs.EOF Then
'                Err.Raise ERR_NOHAYCODIGO, "CuentaDepartamento_CodCuenta", MSGERR_NOHAYCODIGO
'                Exit Property
'            Else
'                IdCuenta = rs!IdCuenta
'            End If
'            rs.Close
'        Else
'            IdCuenta = 0
'        End If
'
'        mProps.CodCuenta = value
'        mbooModificado = True
'    End If
'
'    Set rs = Nothing
End Property



