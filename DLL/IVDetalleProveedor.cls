VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVDetalleProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type T_PROPS
    idProveedor As Long
    IdInventario As Long
    CodProveedor As String
    CodInventario As String
    Nombre As String
    Direccion As String
    Telefono1 As String
    Telefono2 As String
    Ciudad As String
    Cantidad As Currency
    CostoUnitario As Currency
    fecha As Date
    BandPrincipal As Boolean
    PrecioCompra As Currency
    idInventarioProv As Long
    CodInventarioProv As String
    
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
    mProps.objIVInventario.Modificado = True
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
    mProps.objIVInventario.Modificado = True
    'mbooModificado = True
End Property

Public Property Get CodInventario() As String
    CodInventario = Trim$(mProps.CodInventario)
End Property

Private Sub Class_Terminate()
    Set mProps.objIVInventario = Nothing
End Sub

Public Property Get Direccion() As String 'Propiedad  solo de lectura
    Direccion = Trim$(mProps.Direccion)
End Property

Public Property Let Direccion(ByVal value As String)
    If Len(value) > 120 Then
        Err.Raise ERR_INVALIDO, "IVDetalleProveedor_Direccion", MSGERR_INVALIDO
    End If
     mProps.Direccion = value
    mProps.objIVInventario.Modificado = True
End Property

Public Property Let Telefono1(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCContacto_Telefono1", MSGERR_INVALIDO
    End If
    mProps.Telefono1 = value
    mProps.objIVInventario.Modificado = True
End Property

Public Property Get Telefono1() As String
    Telefono1 = Trim$(mProps.Telefono1)
End Property

Public Property Let Telefono2(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCContacto_Telefono2", MSGERR_INVALIDO
    End If
    mProps.Telefono2 = value
    mProps.objIVInventario.Modificado = True
End Property

Public Property Get Telefono2() As String
    Telefono2 = Trim$(mProps.Telefono2)
End Property

Public Property Let Ciudad(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_Ciudad", MSGERR_INVALIDO
    End If
    
    mProps.Ciudad = value
    mProps.objIVInventario.Modificado = True
End Property

Public Property Get Ciudad() As String
    Ciudad = Trim$(mProps.Ciudad)
End Property
 
 'key puede ser IdInventario  o CodInventario
Public Sub Recuperar(key As Variant)
    Dim sql As String, i As Integer
    Dim rs As Recordset
    'select IdInventario,  CodInventario, Descripcion, Unidad   from IVInventario
    If VarType(key) = vbString Then
            sql = "SELECT pc.Idprovcli, pc.Codprovcli,pc.Nombre, pc.direccion1,pc.telefono1,pc.telefono2,pcc.Codcanton "
            sql = sql & " FROM pcprovcli pc "
            sql = sql & " LEFT JOIN PCCanton pcc on pcc.idcanton = pc.idcanton "
            sql = sql & " WHERE Codprovcli='" & key & "'"
    Else
        sql = "SELECT pc.Idprovcli, pc.Codprovcli,pc.Nombre, pc.direccion1,pc.telefono1,pc.telefono2,pcc.Codcanton "
        sql = sql & " FROM pcprovcli pc "
        sql = sql & " LEFT JOIN PCCanton pcc on pcc.idcanton = pc.idcanton "
        sql = sql & " WHERE IDprovcli=" & key
  End If
    Set rs = Me.IVInventario.Empresa.OpenRecordset(sql)
    If rs.EOF Then
        Err.Raise ERR_NOHAYCODIGO, "IvproveedorDetalle", MSGERR_NOHAYCODIGO
    Else
        mProps.idProveedor = rs!IdProvcli
        mProps.CodProveedor = rs!CodProvcli
        mProps.Nombre = rs!Nombre
        mProps.Direccion = rs!Direccion1
        mProps.Telefono1 = rs!Telefono1
        mProps.Telefono2 = rs!Telefono2
        If Not IsNull(rs!CodCanton) Then mProps.Ciudad = rs!CodCanton
  End If
End Sub

Public Property Let Nombre(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "IVDetalleProveedor_Nombre", MSGERR_INVALIDO
    End If
    mProps.Nombre = value
    mProps.objIVInventario.Modificado = True
End Property

Public Property Get Nombre() As String
    Nombre = Trim$(mProps.Nombre)
End Property

Public Property Get fecha() As Date
    fecha = mProps.fecha
End Property

Public Property Let fecha(ByVal value As Date)
    mProps.fecha = value
    mProps.objIVInventario.Modificado = True
End Property

Public Property Get CostoUnitario() As Currency
    CostoUnitario = mProps.CostoUnitario
End Property

Public Property Let CostoUnitario(ByVal value As Currency)
    mProps.CostoUnitario = value
    mProps.objIVInventario.Modificado = True
End Property

Public Property Let Cantidad(ByVal value As Currency)
    mProps.Cantidad = value
    mProps.objIVInventario.Modificado = True
End Property

Public Property Get Cantidad() As Currency
    Cantidad = mProps.Cantidad
End Property

Public Property Let BandPrincipal(ByVal value As Boolean)
    mProps.BandPrincipal = value
    mProps.objIVInventario.Modificado = True
End Property

Public Property Get BandPrincipal() As Boolean
    BandPrincipal = mProps.BandPrincipal
End Property

Public Sub RecuperarxNombre(key As Variant)
    Dim sql As String, i As Integer
    Dim rs As Recordset
        sql = "Select pc.Idprovcli, pc.Codprovcli,pc.Nombre, pc.direccion1,pc.telefono1,pc.telefono2,pcc.Codcanton "
        sql = sql & " FROM pcprovcli pc "
        sql = sql & " LEFT JOIN PCCanton pcc on pcc.idcanton = pc.idcanton "
        sql = sql & " WHERE pc.nombre='" & key & "'"
    Set rs = Me.IVInventario.Empresa.OpenRecordset(sql)
    If rs.EOF Then
        Err.Raise ERR_NOHAYCODIGO, "IvproveedorDetalle", MSGERR_NOHAYCODIGO
    Else
        mProps.idProveedor = rs!IdProvcli
        mProps.CodProveedor = rs!CodProvcli
        mProps.Nombre = rs!Nombre
        mProps.Direccion = rs!Direccion1
        mProps.Telefono1 = rs!Telefono1
        mProps.Telefono2 = rs!Telefono2
        If Not IsNull(rs!CodCanton) Then mProps.Ciudad = rs!CodCanton
    End If
End Sub

Public Property Let PrecioCompra(ByVal value As Currency)
    mProps.PrecioCompra = value
    mProps.objIVInventario.Modificado = True
End Property

Public Property Get PrecioCompra() As Currency
    PrecioCompra = mProps.PrecioCompra
End Property

Public Property Let CodInventarioProv(ByVal value As String)

    Dim sql As String, rs As Recordset
    
    If Len(value) > 40 Then
        Err.Raise ERR_INVALIDO, "PCProvCli_CodProvCli", MSGERR_INVALIDO
    End If
    
    
    
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodInventarioProv = value
    mProps.objIVInventario.Modificado = True
    'mbooModificado = True
End Property

Public Property Get CodInventarioProv() As String
    CodInventarioProv = Trim$(mProps.CodInventarioProv)
End Property

Public Property Let idInventarioProv(ByVal value As Long)
    mProps.idInventarioProv = value
End Property

Public Property Get idInventarioProv() As Long
    idInventarioProv = mProps.idInventarioProv
End Property

