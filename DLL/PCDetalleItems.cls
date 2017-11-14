VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCDetalleItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type T_PROPS
    idProveedor As Long
    IdInventario As Long
    CodProveedor As String
    CodInventario As String
    Descripcion As String
'    Direccion As String
'    Telefono1 As String
'    Telefono2 As String
'    Ciudad As String
    Cantidad As Currency
    PrecioUnitario As Currency
    fecha As Date
 '   BandPrincipal As Boolean
'    PrecioCompra As Currency
    objPcProvCli As pcprovcli
End Type
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean
Private mProps As T_PROPS

Friend Property Set Pcpprovcli(obj As pcprovcli)
    Set mProps.objPcProvCli = obj
End Property

Public Property Get pcprovcli() As pcprovcli
    Set pcprovcli = mProps.objPcProvCli
End Property

Public Property Let idProveedor(ByVal value As Long)
    mProps.idProveedor = value
End Property

Public Property Get idProveedor() As Long
    idProveedor = mProps.idProveedor
End Property

Public Property Let IdInventario(ByVal value As Long)
    mProps.IdInventario = value
End Property

Public Property Get IdInventario() As Long
    IdInventario = mProps.IdInventario
End Property

Public Property Let CodProveedor(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_CodProvCli", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodProveedor = value
'    mProps.objPcProvCli.Modificado = True
    'mbooModificado = True
End Property

Public Property Get CodProveedor() As String
    CodProveedor = Trim$(mProps.CodProveedor)
End Property

Public Property Let CodInventario(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "IVInvent_CodInventario", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodInventario = value
'    mProps.objPcProvCli.Modificado = True
    'mbooModificado = True
End Property

Public Property Get CodInventario() As String
    CodInventario = Trim$(mProps.CodInventario)
End Property

Private Sub Class_Terminate()
    Set mProps.objPcProvCli = Nothing
End Sub


Public Sub Recuperar(key As Variant)
    Dim sql As String, i As Integer
    Dim rs As Recordset
    'select IdInventario,  CodInventario, Descripcion, Unidad   from IVInventario
    If VarType(key) = vbString Then
            sql = "SELECT pc.Idprovcli, pc.Codprovcli  FROM pcprovcli pc "
            sql = sql & " WHERE Codprovcli='" & key & "'"
    Else
        sql = "SELECT pc.Idprovcli, pc.Codprovcli,pc.DescripcionFROM pcprovcli pc "
        sql = sql & " WHERE IDprovcli=" & key
  End If
    Set rs = Me.pcprovcli.Empresa.OpenRecordset(sql)
    If rs.EOF Then
        Err.Raise ERR_NOHAYCODIGO, "IvproveedorDetalle", MSGERR_NOHAYCODIGO
    Else
        mProps.idProveedor = rs!IdProvCli
        mProps.CodProveedor = rs!CodProvCli
        mProps.Descripcion = rs!Descripcion
  End If
End Sub

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "IVDetalleProveedor_Descripcion", MSGERR_INVALIDO
    End If
    mProps.Descripcion = value
'    mProps.objPcProvCli.Modificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Get fecha() As Date
    fecha = mProps.fecha
End Property

Public Property Let fecha(ByVal value As Date)
    mProps.fecha = value
'    mProps.objPcProvCli.Modificado = True
End Property

Public Property Get PrecioUnitario() As Currency
    PrecioUnitario = mProps.PrecioUnitario
End Property

Public Property Let PrecioUnitario(ByVal value As Currency)
    mProps.PrecioUnitario = value
'    mProps.objPcProvCli.Modificado = True
End Property

Public Property Let Cantidad(ByVal value As Currency)
    mProps.Cantidad = value
'    mProps.objPcProvCli.Modificado = True
End Property

Public Property Get Cantidad() As Currency
    Cantidad = mProps.Cantidad
End Property


Public Sub RecuperarxDescripcion(key As Variant)
    Dim sql As String, i As Integer
    Dim rs As Recordset
        sql = "Select pc.Idprovcli, pc.Codprovcli,pc.Descripcion, pc.direccion1,pc.telefono1,pc.telefono2,pcc.Codcanton FROM pcprovcli pc "
        sql = sql & " LEFT JOIN PCCanton pcc on pcc.idcanton = pc.idcanton "
        sql = sql & " WHERE pc.Descripcion='" & key & "'"
    Set rs = Me.pcprovcli.Empresa.OpenRecordset(sql)
    If rs.EOF Then
        Err.Raise ERR_NOHAYCODIGO, "IvproveedorDetalle", MSGERR_NOHAYCODIGO
    Else
        mProps.idProveedor = rs!IdProvCli
        mProps.CodProveedor = rs!CodProvCli
        mProps.Descripcion = rs!Descripcion
'        mProps.Direccion = rs!Direccion1
 '       mProps.Telefono1 = rs!Telefono1
  '      mProps.Telefono2 = rs!Telefono2
   '     If Not IsNull(rs!CodCanton) Then mProps.Ciudad = rs!CodCanton
    End If
End Sub

'Public Property Let PrecioCompra(ByVal value As Currency)
'    mProps.PrecioCompra = value
'    mProps.objPcProvCli.Modificado = True
'End Property
'
'Public Property Get PrecioCompra() As Currency
'    PrecioCompra = mProps.PrecioCompra
'End Property
'
