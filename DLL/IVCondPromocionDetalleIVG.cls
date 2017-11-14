VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVCondPromocionDetalleIVG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***Diego 15/09/2003
Private Type T_PROPS
    IdPromocion As Long
    IdCondPromocionDetalle As Long

    NumIVGrupo  As Byte
    IDIVGrupo  As Integer
    CodIVGrupo As String      'revisar  si nos sirve
    DescripcionIVGrupo As String
    Idinventario As Integer
    
    xCuanto As Currency
    xCuantoPeso As Currency
    xCuantoUSD As Currency

    Orden As Long
    objIVpromocion As IVpromocion  'Referencia al Objeto Padre
End Type

Private mProps As T_PROPS
'Private mbooEsNuevo As Boolean
'Private mbooModificado As Boolean


'Establece relacion  con el PADRE
Friend Property Set IVpromocion(objIVpromocion As IVpromocion)
    Set mProps.objIVpromocion = objIVpromocion

End Property

Public Property Get IVpromocion() As IVpromocion
    Set IVpromocion = mProps.objIVpromocion
End Property

Friend Sub Recuperar(rs As Recordset)
    Dim sql As String, rs1  As Recordset
    With rs
        If Not .EOF Then
            mProps.IdPromocion = !IdPromocion
            mProps.IdCondPromocionDetalle = !IdCondPromocionDetalle
            If Not IsNull(!NumIVGrupo) Then mProps.NumIVGrupo = !NumIVGrupo
            If Not IsNull(!IDIVGrupo) Then mProps.IDIVGrupo = !IDIVGrupo
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            If Not IsNull(!xCuanto) Then mProps.xCuanto = !xCuanto
            If Not IsNull(!xCuantoPeso) Then mProps.xCuantoPeso = !xCuantoPeso
            If Not IsNull(!xCuantoUSD) Then mProps.xCuantoUSD = !xCuantoUSD

        End If
    End With
End Sub

Public Function Clone() As IvCondPromocionDetalleIVG
    Dim nuevo As IvCondPromocionDetalleIVG

    'Copia las propiedades al nuevo objeto
    Set nuevo = New IvCondPromocionDetalleIVG
    nuevo.Props = mProps

    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdPromocion = 0
    'mbooModificado = True
End Property


Public Property Get IdPromocion() As Long
    IdPromocion = mProps.IdPromocion
End Property

Public Property Let IdPromocion(ByVal value As Long)
    mProps.IdPromocion = value
End Property

Public Property Get Orden() As Long
    Orden = mProps.Orden
End Property

Public Property Let Orden(ByVal value As Long)
    mProps.Orden = value
End Property



Private Sub Class_Terminate()
    Set mProps.objIVpromocion = Nothing
End Sub


Public Property Let NumIVGrupo(ByVal value As Byte)
    mProps.NumIVGrupo = value
    mProps.objIVpromocion.Modificado = True
End Property

Public Property Get NumIVGrupo() As Byte
    NumIVGrupo = mProps.NumIVGrupo
End Property



Public Property Get CodIVGrupo() As String
    CodIVGrupo = mProps.CodIVGrupo
End Property


Public Property Let CodIVGrupo(value As String)
    Dim rs As Recordset, sql As String
    mProps.CodIVGrupo = value
    If value = "" Then
        mProps.IDIVGrupo = 0
    Else
        'actualiza  el IDGrupo
        sql = "Select IdGrupo" & mProps.NumIVGrupo & " From IVGrupo" & mProps.NumIVGrupo & _
              " Where CodGrupo" & mProps.NumIVGrupo & " = '" & value & "'"
        Set rs = mProps.objIVpromocion.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then
            mProps.IDIVGrupo = rs.Fields(0).value
        End If
        rs.Close
        Set rs = Nothing
    End If
    mProps.objIVpromocion.Modificado = True
End Property

Public Property Get IDIVGrupo() As Long
    IDIVGrupo = mProps.IDIVGrupo
End Property


Public Property Let xCuanto(ByVal value As Currency)
    mProps.xCuanto = value
    mProps.objIVpromocion.Modificado = True
End Property

Public Property Get xCuanto() As Currency
    xCuanto = mProps.xCuanto
End Property

Public Property Let xCuantoPeso(ByVal value As Currency)
    mProps.xCuantoPeso = value
    mProps.objIVpromocion.Modificado = True
End Property

Public Property Get xCuantoPeso() As Currency
    xCuantoPeso = mProps.xCuantoPeso
End Property

Public Property Let xCuantoUSD(ByVal value As Currency)
    mProps.xCuantoUSD = value
    mProps.objIVpromocion.Modificado = True
End Property

Public Property Get xCuantoUSD() As Currency
    xCuantoUSD = mProps.xCuantoUSD
End Property



