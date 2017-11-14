VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCHistorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROP
    objEmpresa As Empresa
    IdHistorial As Long             'Id de Historial
    IdProvCli As Long
    TransID As Long
    FechaTrans As Date
    FechaGrabado As Date
    Estado As String
    Trans As String
    Descripcion As String
    Valor As Currency
End Type

Private mProps As T_PROP

Public Sub Grabar(ByVal EsNuevo As Boolean)
    Dim rs As Recordset, sql As String
    'Si no está modificado no hace nada
    'If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
'    sql = "SELECT transid FROM PCHistorial " & _
'          " WHERE (transid)='" & mProps.TransID & "' AND " & _
'                 "(Idhistorial <>" & mProps.IdHistorial & ")"
'    Set rs = mProps.objEmpresa.OpenRecordset(sql)
'    If Not rs.EOF Then      'Si ya existe el codigo
'        Err.Raise ERR_REPITECODIGO, "PCHistorial_Grabar", MSGERR_REPITECODIGO
'    End If
'    rs.Close
    'Si es nuevo
    If EsNuevo Then
        sql = "SELECT * FROM PCHistorial WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM PCHistorial WHERE IdHistorial =" & mProps.IdHistorial
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
        !IdProvCli = mProps.IdProvCli
        !TransID = mProps.TransID
        !FechaTrans = mProps.FechaTrans
        !Estado = mProps.Estado
        !Trans = mProps.Trans
        !Descripcion = mProps.Descripcion
        !Valor = mProps.Valor
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        .Bookmark = .LastModified
#End If
        rs.Move 0
        mProps.IdHistorial = !IdHistorial
        .Close
    End With
    Set rs = Nothing
End Sub

Friend Sub Recuperar(ByVal Trans As String)
    Dim sql As String, rs As Recordset
    
    'Trans = CodTrans & " " & NumTrans
    sql = "SELECT * FROM PCHistorial WHERE trans ='" & Trans & "' "
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        mProps.IdHistorial = !IdHistorial
        If Not IsNull(!IdProvCli) Then mProps.IdProvCli = !IdProvCli
        If Not IsNull(!TransID) Then mProps.TransID = !TransID
        If Not IsNull(!FechaTrans) Then mProps.FechaTrans = !FechaTrans
        If Not IsNull(!Estado) Then mProps.Estado = !Estado
        If Not IsNull(!Trans) Then mProps.Trans = !Trans
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!Valor) Then mProps.Valor = !Valor
        If Not IsNull(!FechaGrabado) Then mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    Set rs = Nothing
End Sub

Friend Function Clone() As PCHistorial
    Dim nuevo As PCHistorial
    
    Set nuevo = New PCHistorial
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROP)          'Para el método Clone
    mProps = value
    'mProps.objPC.SetModificado
End Property

'Friend Property Set PCProvCli(value As PCProvCli)     'Para el método AddPCContacto de Grupo
'    Set mProps.objPC = value
'End Property
'
'Public Property Get PCProvCli() As PCProvCli
'    Set PCProvCli = mProps.objPC
'End Property

Public Property Get IdHistorial() As Long
    IdHistorial = mProps.IdHistorial
End Property
Public Property Let TransID(ByVal value As Long)
    mProps.TransID = value
    'mProps.objPC.SetModificado
End Property

Public Property Get TransID() As Long
    TransID = mProps.TransID
End Property

Public Property Let FechaTrans(value As Date)
    'Solo cuando cambia el valor
    If mProps.FechaTrans <> value Then
        mProps.FechaTrans = value
        '        mProps.objPC.SetModificado
    End If
End Property
Public Property Get FechaTrans() As Date
    FechaTrans = mProps.FechaTrans
End Property

Public Property Let Estado(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCContacto_Estado", MSGERR_INVALIDO
    End If
    mProps.Estado = value
    'mProps.objPC.SetModificado
End Property

Public Property Get Estado() As String
    Estado = Trim$(mProps.Estado)
End Property


Public Property Let Trans(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCHistorial_Trans", MSGERR_INVALIDO
    End If
    mProps.Trans = value
    'mProps.objPC.SetModificado
End Property

Public Property Get Trans() As String
    Trans = Trim$(mProps.Trans)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "PCHistorial_Descripcion", MSGERR_INVALIDO
    End If
    mProps.Descripcion = value
    'mProps.objPC.SetModificado
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let Valor(ByVal value As Currency)
    mProps.Valor = value
End Property
Public Property Get Valor() As Currency
    Valor = mProps.Valor
End Property
Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub
Public Property Get FechaGrabado() As Date
    FechaGrabado = mProps.FechaGrabado
End Property
Public Property Let FechaGrabado(ByVal value As Date)
    mProps.FechaGrabado = value
    'mProps.objPC.SetModificado
End Property

Public Property Let IdProvCli(ByVal value As Long)
    mProps.IdProvCli = value
    'mProps.objPC.SetModificado
End Property

Public Property Get IdProvCli() As Long
    IdProvCli = mProps.IdProvCli
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property
