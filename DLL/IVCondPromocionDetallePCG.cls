VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVCondPromocionDetallePCG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***Diego 15/09/2003
Private Type T_PROPS
    IdPromocion As Long
    IdCondPromocionDetalle As Long

    NumPCGrupo  As Byte
    IDPCGrupo  As Integer
    CodPCGrupo As String      'revisar  si nos sirve
    DescripcionIVGrupo As String
    IdProvCli As Integer
    
    Orden As Long
    objIVPromocion As IVPromocion  'Referencia al Objeto Padre
End Type

Private mProps As T_PROPS
'Private mbooEsNuevo As Boolean
'Private mbooModificado As Boolean


'Establece relacion  con el PADRE
Friend Property Set IVPromocion(objIVPromocion As IVPromocion)
    Set mProps.objIVPromocion = objIVPromocion

End Property

Public Property Get IVPromocion() As IVPromocion
    Set IVPromocion = mProps.objIVPromocion
End Property

Friend Sub Recuperar(rs As Recordset)
    Dim sql As String, rs1  As Recordset
    With rs
        If Not .EOF Then
            mProps.IdPromocion = !IdPromocion
            mProps.IdCondPromocionDetalle = !IdCondPromocionDetalle
            If Not IsNull(!NumPCGrupo) Then mProps.NumPCGrupo = !NumPCGrupo
            If Not IsNull(!IDPCGrupo) Then mProps.IDPCGrupo = !IDPCGrupo
            If Not IsNull(!Orden) Then mProps.Orden = !Orden

        End If
    End With
End Sub

Public Function Clone() As IvCondPromocionDetallePCG
    Dim nuevo As IvCondPromocionDetallePCG

    'Copia las propiedades al nuevo objeto
    Set nuevo = New IvCondPromocionDetallePCG
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
    Set mProps.objIVPromocion = Nothing
End Sub


Public Property Let NumPCGrupo(ByVal value As Byte)
    mProps.NumPCGrupo = value
    mProps.objIVPromocion.Modificado = True
End Property

Public Property Get NumPCGrupo() As Byte
    NumPCGrupo = mProps.NumPCGrupo
End Property



Public Property Get CodPCGrupo() As String
    CodPCGrupo = mProps.CodPCGrupo
End Property


Public Property Let CodPCGrupo(value As String)
    Dim rs As Recordset, sql As String
    mProps.CodPCGrupo = value
    If value = "" Then
        mProps.IDPCGrupo = 0
    Else
        'actualiza  el IDGrupo
        sql = "Select IdGrupo" & mProps.NumPCGrupo & " From PCGrupo" & mProps.NumPCGrupo & _
              " Where CodGrupo" & mProps.NumPCGrupo & " = '" & value & "'"
        Set rs = mProps.objIVPromocion.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then
            mProps.IDPCGrupo = rs.Fields(0).value
        End If
        rs.Close
        Set rs = Nothing
    End If
    mProps.objIVPromocion.Modificado = True
End Property

Public Property Get IDPCGrupo() As Long
    IDPCGrupo = mProps.IDPCGrupo
End Property


