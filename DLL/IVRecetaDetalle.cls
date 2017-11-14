VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVRecetaDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type T_PROPS
    IdReceta As Long
    IdInventario As Long
    CodInventario As String
    DescInventario As String
    Cantidad As Double
    'Costo As Double
    Unidad As String
    Orden As Integer
    objReceta As IVReceta
End Type

Private mProps As T_PROPS

'Establece relacion  con el PADRE
Friend Property Set IVReceta(objReceta As IVReceta)
    Set mProps.objReceta = objReceta
End Property

Public Property Get IVReceta() As IVReceta
    Set IVReceta = mProps.objReceta
End Property

Friend Sub Recuperar(rs As Recordset)
    Dim sql As String, rs1  As Recordset
    With rs
        If Not .EOF Then
            mProps.IdReceta = !IdReceta
            'mProps.IdDescuentoDetalleIV = !IdDescuentoDetalleIV
            If Not IsNull(!IdInventario) Then mProps.IdInventario = !IdInventario
            If Not IsNull(!Cantidad) Then mProps.Cantidad = !Cantidad
'            If Not IsNull(!Costo) Then mProps.Costo = !Costo
            If Not IsNull(!CodInventario) Then mProps.CodInventario = !CodInventario
            If Not IsNull(!Descripcion) Then mProps.DescInventario = !Descripcion
            If Not IsNull(!CodUnidad) Then mProps.Unidad = !CodUnidad
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
                        
        End If
    End With
End Sub

Public Function Clone() As IVRecetaDetalle
    Dim nuevo As IVRecetaDetalle

    'Copia las propiedades al nuevo objeto
    Set nuevo = New IVRecetaDetalle
    nuevo.Props = mProps

    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdReceta = 0
    'mbooModificado = True
End Property


Public Property Get IdReceta() As Long
    IdReceta = mProps.IdReceta
End Property

Public Property Let IdReceta(ByVal value As Long)
    mProps.IdReceta = value
End Property

Public Property Get Orden() As Long
    Orden = mProps.Orden
End Property

Public Property Let Orden(ByVal value As Long)
    mProps.Orden = value
End Property

Private Sub Class_Terminate()
    Set mProps.objReceta = Nothing
End Sub

Public Property Get CodInventario() As String
    CodInventario = mProps.CodInventario
End Property

Public Property Let CodInventario(value As String)
    Dim rs As Recordset, sql As String
    mProps.CodInventario = value
    If value = "" Then
        mProps.IdInventario = 0
    Else
        'actualiza  el IDGrupo
        sql = "Select IdInventario From IvInventario"
        sql = sql & " Where CodInventario = '" & value & "'"
        Set rs = mProps.objReceta.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then
            mProps.IdInventario = rs.Fields(0).value
        End If
        rs.Close
        Set rs = Nothing
    End If
    mProps.objReceta.Modificado = True
End Property

Public Property Get IdInventario() As Long
    IdInventario = mProps.IdInventario
End Property

Public Property Let IdInventario(ByVal value As Long)
    mProps.IdInventario = value
    mProps.objReceta.Modificado = True
End Property

'Public Property Let Costo(ByVal value As Double)
'    mProps.Costo = value
'    mProps.objReceta.Modificado = True
'End Property
'
'Public Property Get Costo() As Double
'    Costo = mProps.Costo
'End Property

Public Property Let Cantidad(ByVal value As Double)
    mProps.Cantidad = value
    mProps.objReceta.Modificado = True
End Property

Public Property Get Cantidad() As Double
    Cantidad = mProps.Cantidad
End Property

Public Property Get DescInventario() As String
    DescInventario = mProps.DescInventario
End Property

Public Property Let DescInventario(ByVal value As String)
    mProps.DescInventario = value
End Property

Public Property Get Unidad() As String
    Unidad = mProps.Unidad
End Property

Public Property Let Unidad(ByVal value As String)
    mProps.Unidad = value
End Property
