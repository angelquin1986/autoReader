VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVGarantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type T_PROPS
    IdInventario As Long
    fechadesde As Date
    fechaHasta As Date
    Valor As Currency
    objIVInventario As IVInventario
End Type
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean
Private mProps As T_PROPS

Friend Property Set IVInventario(obj As IVInventario)
    Set mProps.objIVInventario = obj
End Property

Public Property Get IVInventario() As IVInventario
    Set IVInventario = mProps.objIVInventario
End Property


Public Property Let IdInventario(ByVal value As Long)
    mProps.IdInventario = value
End Property

Public Property Get IdInventario() As Long
    IdInventario = mProps.IdInventario
End Property


Private Sub Class_Terminate()
    Set mProps.objIVInventario = Nothing
End Sub
 
 'key puede ser IdInventario  o CodInventario
Public Sub Recuperar(key As Variant)
    Dim sql As String, i As Integer
    Dim rs As Recordset
    'select IdInventario,  CodInventario, Descripcion, Unidad   from IVInventario
    If VarType(key) = vbString Then
            sql = "SELECT * "
            sql = sql & " FROM IVGarantia"
            sql = sql & " WHERE IdInventario='" & key & "'"
    Else
        sql = "SELECT * "
            sql = sql & " FROM IVGarantia"
            sql = sql & " WHERE IdInventario=" & key
  End If
    Set rs = Me.IVInventario.Empresa.OpenRecordset(sql)
    If rs.EOF Then
        Err.Raise ERR_NOHAYCODIGO, "IvGarantia", MSGERR_NOHAYCODIGO
    Else
        mProps.IdInventario = rs!IdInventario
        mProps.fechadesde = rs!fechadesde
        mProps.fechaHasta = rs!fechaHasta
        mProps.Valor = rs!Valor
  End If
End Sub

Public Property Get fechadesde() As Date
    fechadesde = mProps.fechadesde
End Property

Public Property Let fechadesde(ByVal value As Date)
    mProps.fechadesde = value
    mProps.objIVInventario.Modificado = True
End Property

Public Property Get fechaHasta() As Date
    fechaHasta = mProps.fechaHasta
End Property

Public Property Let fechaHasta(ByVal value As Date)
    mProps.fechaHasta = value
    mProps.objIVInventario.Modificado = True
End Property

Public Property Let Valor(ByVal value As Currency)
    mProps.Valor = value
    mProps.objIVInventario.Modificado = True
End Property

Public Property Get Valor() As Currency
    Valor = mProps.Valor
End Property

