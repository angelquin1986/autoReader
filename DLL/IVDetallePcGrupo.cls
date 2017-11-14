VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVDetallePcGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type T_PROPS
    IdGrupo1 As Long
    IdGrupo2 As Long
    IdGrupo3 As Long
    IdGrupo4 As Long
    IdInventario As Long
    CodInventario As String
    Descripcion As String
    FechaGrabado As Date
    Cantidad As Currency
    PU As Currency
    objPCGrupo As PcGrupo
End Type
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean
Private mProps As T_PROPS

Friend Property Set PcGrupo(obj As PcGrupo)
    Set mProps.objPCGrupo = obj
End Property

Public Property Get PcGrupo() As PcGrupo
    Set PcGrupo = mProps.objPCGrupo
End Property

Public Property Let IdGrupo1(ByVal value As Long)
    mProps.IdGrupo1 = value
End Property

Public Property Get IdGrupo1() As Long
    IdGrupo1 = mProps.IdGrupo1
End Property
Public Property Let IdGrupo2(ByVal value As Long)
    mProps.IdGrupo2 = value
End Property

Public Property Get IdGrupo2() As Long
    IdGrupo2 = mProps.IdGrupo2
End Property
Public Property Let IdGrupo3(ByVal value As Long)
    mProps.IdGrupo3 = value
End Property

Public Property Get IdGrupo3() As Long
    IdGrupo3 = mProps.IdGrupo3
End Property
Public Property Let IdGrupo4(ByVal value As Long)
    mProps.IdGrupo4 = value
End Property

Public Property Get IdGrupo4() As Long
    IdGrupo4 = mProps.IdGrupo4
End Property

Public Property Let IdInventario(ByVal value As Long)
    mProps.IdInventario = value
End Property

Public Property Get IdInventario() As Long
    IdInventario = mProps.IdInventario
End Property


Public Property Get CodInventario() As String
    CodInventario = Trim$(mProps.CodInventario)
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Private Sub Class_Terminate()
    Set mProps.objPCGrupo = Nothing
End Sub

 'key puede ser IdInventario  o CodInventario
Public Sub Recuperar(key As Variant)
    Dim sql As String, i As Integer
    Dim rs As Recordset
    
    sql = "Select iv.codinventario,iv.descripcion FROM vwivinventariorecuperar  iv  WHERE codinventario='" & key & "'"
    
    Set rs = Me.PcGrupo.Empresa.OpenRecordset(sql)
    If rs.EOF Then
        Err.Raise ERR_NOHAYCODIGO, "PCGrupoInventario", MSGERR_NOHAYCODIGO
    Else
        mProps.CodInventario = rs!CodInventario
        mProps.Descripcion = rs!Descripcion
        
    End If
End Sub
Public Property Get FechaGrabado() As Date
    FechaGrabado = mProps.FechaGrabado
End Property

Public Property Let FechaGrabado(ByVal value As Date)
    mProps.FechaGrabado = value
    'mProps.objPCGrupo.Modificado = True
End Property

Public Property Let CodInventario(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "IVDetallePcGrupo_CodInventario", MSGERR_INVALIDO
    End If
'    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodInventario = value
    'mProps.objPCGrupo.Modificado = True
    
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "IVDetallePcGrupo_Descripcion", MSGERR_INVALIDO
    End If
    mProps.Descripcion = value
    'mbooModificado = True
End Property
Public Property Let Cantidad(ByVal value As Currency)
    mProps.Cantidad = value
    mProps.objPCGrupo.Modificado = True    'trabaja con  la bandera modificado del papa
End Property

Public Property Get Cantidad() As Currency
    Cantidad = mProps.Cantidad
End Property
Public Property Let PU(ByVal value As Currency)
    mProps.PU = value
    mProps.objPCGrupo.Modificado = True    'trabaja con  la bandera modificado del papa
End Property

Public Property Get PU() As Currency
    PU = mProps.PU
End Property



