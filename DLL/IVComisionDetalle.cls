VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVComisionDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***Diego 15/09/2003
Private Type T_PROPS
    IdComision As Long
    IdComisionDetalle As Long
    Comision As Currency

    IdInventario  As Long
    CodInventario As String      'revisar  si nos sirve

    IdVendedor  As Long
    CodVendedor As String      'revisar  si nos sirve


    IdGrupo  As Long
    CodGrupo As String      'revisar  si nos sirve


    Orden As Long
    objIVComision As IVComision   'Referencia al Objeto Padre
End Type

Private mProps As T_PROPS


'Establece relacion  con el PADRE
Friend Property Set IVComision(objIVComision As IVComision)
    Set mProps.objIVComision = objIVComision
End Property

Public Property Get IVComision() As IVComision
    Set IVComision = mProps.objIVComision
End Property

Friend Sub Recuperar(rs As Recordset)
    Dim sql As String, rs1  As Recordset
    With rs
        If Not .EOF Then
            mProps.IdComision = !IdComision
            mProps.IdComisionDetalle = !IdComisionDetalle
            If Not IsNull(!IdInventario) Then mProps.IdInventario = !IdInventario
            If Not IsNull(!CodInventario) Then mProps.CodInventario = !CodInventario
            If Not IsNull(!IdVendedor) Then mProps.IdVendedor = !IdVendedor
            If Not IsNull(!IdGrupo) Then mProps.IdGrupo = !IdGrupo
            If Not IsNull(!CodGrupo) Then mProps.CodGrupo = !CodGrupo
            mProps.Comision = !Comision
            
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            'recupera  el codigo  de Vendedor
            On Error Resume Next
            
                sql = "Select CodVendedor From FCVendedor"
                sql = sql & " Where IDVendedor = " & mProps.IdVendedor
                Set rs1 = mProps.objIVComision.Empresa.OpenRecordset(sql)
                If Not rs1.EOF Then
                    mProps.CodVendedor = rs1.Fields(0).value
                End If
                rs1.Close
                Set rs1 = Nothing


        End If
    End With
End Sub

Public Function Clone() As IVComisionDetalle
    Dim nuevo As IVComisionDetalle

    'Copia las propiedades al nuevo objeto
    Set nuevo = New IVComisionDetalle
    nuevo.Props = mProps

    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdComision = 0
End Property


Public Property Get IdComision() As Long
    IdComision = mProps.IdComision
End Property

Public Property Let IdComision(ByVal value As Long)
    mProps.IdComision = value
End Property

Public Property Get Orden() As Long
    Orden = mProps.Orden
End Property

Public Property Let Orden(ByVal value As Long)
    mProps.Orden = value
End Property

Public Property Let Comision(ByVal value As Currency)
    mProps.Comision = value
    mProps.objIVComision.Modificado = True
End Property

Public Property Get Comision() As Currency
    Comision = mProps.Comision
End Property


Public Property Get CodVendedor() As String
    CodVendedor = mProps.CodVendedor
End Property


Public Property Let CodVendedor(value As String)
    Dim rs As Recordset, sql As String
    mProps.CodVendedor = value
    If value = "" Then
        mProps.IdVendedor = 0
    Else
        'actualiza  el IDGrupo
        sql = "Select IdVendedor  From FCVendedor "
        sql = sql & " Where CodVendedor  = '" & value & "'"
        Set rs = mProps.objIVComision.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then
            mProps.IdVendedor = rs.Fields(0).value
        End If
        rs.Close
        Set rs = Nothing
    End If
    mProps.objIVComision.Modificado = True
End Property


Public Property Get IdVendedor() As Long
    IdVendedor = mProps.IdVendedor
End Property



Private Sub Class_Terminate()
    Set mProps.objIVComision = Nothing
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
        Set rs = mProps.objIVComision.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then
            mProps.IdInventario = rs.Fields(0).value
        End If
        rs.Close
        Set rs = Nothing
    End If
    mProps.objIVComision.Modificado = True
End Property

Public Property Get IdInventario() As Long
    IdInventario = mProps.IdInventario
End Property



Public Property Get CodGrupo() As String
    CodGrupo = mProps.CodGrupo
End Property

Public Property Let CodGrupo(value As String)
    Dim rs As Recordset, sql As String, NumIVGrupo As Integer
    mProps.CodGrupo = value
    If value = "" Then
        mProps.IdGrupo = 0
    Else
            If Len(mProps.objIVComision.Empresa.GNOpcion.ObtenerValor("ComisionNumIVGrupo")) > 0 Then
                NumIVGrupo = mProps.objIVComision.Empresa.GNOpcion.ObtenerValor("ComisionNumIVGrupo") + 1
            Else
                NumIVGrupo = 1
            End If
    
        'actualiza  el IDGrupo
        sql = "Select IdGrupo" & NumIVGrupo & " From IvGrupo" & NumIVGrupo
        sql = sql & " Where CodGrupo" & NumIVGrupo & " = '" & value & "'"
        Set rs = mProps.objIVComision.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then
            mProps.IdGrupo = rs.Fields(0).value
        End If
        rs.Close
        Set rs = Nothing
    End If
    mProps.objIVComision.Modificado = True
End Property

Public Property Get IdGrupo() As Long
    IdGrupo = mProps.IdGrupo
End Property


