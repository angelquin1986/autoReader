VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GNSucursal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IDSucursal As Long
    CodSucursal As String
    Nombre As String
    BandValida As Boolean
    NumEstablecimiento As String
    Direccion As String
    Telefono As String
    BaseDatos As String
    BandCliente As Boolean
    BandProveedor As Boolean
    BandInventario As Boolean
    BandMatriz As Boolean
    servidor As String
    NumPuntos As Integer
    fechagrabado As Date
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean


Public Sub Grabar()
    Dim sql As String, rs As Recordset
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código                '*** MAKOTO 17/oct/00
    sql = "SELECT CodSucursal FROM GNSucursal " & _
          "WHERE (CodSucursal='" & mProps.CodSucursal & "') AND " & _
                "(IdSucursal<>" & mProps.IDSucursal & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "GNSucursal", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    
''Verifica si no repite el NumEstablecimiento
    If mbooEsNuevo Then
        sql = "SELECT CodSucursal FROM GNSucursal " & _
              "WHERE (NumEstablecimiento='" & mProps.NumEstablecimiento & "') "
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then      'Si ya existe el NumEstablecimiento
            If rs.Fields("CodSucursal") <> Me.CodSucursal Then
                Err.Raise ERR_REPITENNUMESTA, "GNSucursal", MSGERR_REPITENUMESTA
            End If
        End If
        rs.Close
    End If
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM GNSucursal WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM GNSucursal WHERE IDSucursal=" & mProps.IDSucursal
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "GNSucursal.Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro
    With rs
        If Not !CodSucursal = mProps.CodSucursal Then
            mProps.objEmpresa.GrabaAccionModulo !CodSucursal, mProps.CodSucursal, "Codigo Sucursal", "MOD_SUC", "GN"
        End If
        !CodSucursal = mProps.CodSucursal
        If Not !Nombre = mProps.Nombre Then
            mProps.objEmpresa.GrabaAccionModulo !Nombre, mProps.Nombre, "Nombre Sucursal", "MOD_SUC", "GN"
        End If
        !Nombre = mProps.Nombre
        If Not !NumEstablecimiento = mProps.NumEstablecimiento Then
            mProps.objEmpresa.GrabaAccionModulo !NumEstablecimiento, mProps.NumEstablecimiento, "Numero Establecimiento", "MOD_INFSUC", "GN"
        End If
        !NumEstablecimiento = mProps.NumEstablecimiento
        If Not !Direccion = mProps.Direccion Then
            mProps.objEmpresa.GrabaAccionModulo !Direccion, mProps.Direccion, "Direccion Sucursal", "MOD_SUC", "GN"
        End If
        !Direccion = mProps.Direccion
        If Not !Telefono = mProps.Telefono Then
            mProps.objEmpresa.GrabaAccionModulo !Telefono, mProps.Telefono, "Telefono Sucursal", "MOD_SUC", "GN"
        End If
        !Telefono = mProps.Telefono
        !BandValida = mProps.BandValida
        !BaseDatos = mProps.BaseDatos
        !BandCliente = False 'mProps.BandCliente
        !BandProveedor = False 'mProps.BandProveedor
        !BandInventario = False 'mProps.BandInventario
        If Not !BandMatriz = mProps.BandMatriz Then
            mProps.objEmpresa.GrabaAccionModulo !BandMatriz, mProps.BandMatriz, "BanderaMatriz", "MOD_SUC", "GN"
        End If
        !BandMatriz = mProps.BandMatriz
        !servidor = mProps.servidor
        If Not !NumPuntos = mProps.NumPuntos Then
            mProps.objEmpresa.GrabaAccionModulo !NumPuntos, mProps.NumPuntos, "Numero Puntos", "MOD_SUC", "GN"
        End If
        !NumPuntos = mProps.NumPuntos
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IDSucursal = !IDSucursal
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_SUC", "Sucursal Nuevo: " & _
                                                         mProps.CodSucursal, "GN"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_SUC", "Sucursal Modificado: " & _
                                                         mProps.CodSucursal, "GN"
    End If
    'actualiza Datos en transacciones de esta sucursal
    mProps.objEmpresa.ActualizaNumEstaBlecimientoDatosSucursal mProps.CodSucursal, mProps.NumEstablecimiento
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_GNSUCURSAL
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM GNSucursal  WHERE CodSucursal='" & key & "'"
    Else
        sql = "SELECT * FROM GNSucursal WHERE IdSucursal=" & key
    End If

    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'if rs.RecordCount= 0 end
    'rs.MoveFirst   'para que se coloque al inicio jeaa
    'Recupera los datos
    With rs
        mProps.IDSucursal = !IDSucursal
        mProps.CodSucursal = !CodSucursal
        If Not IsNull(!Nombre) Then mProps.Nombre = !Nombre
        If Not IsNull(!NumEstablecimiento) Then mProps.NumEstablecimiento = !NumEstablecimiento
        If Not IsNull(!Direccion) Then mProps.Direccion = !Direccion
        If Not IsNull(!Telefono) Then mProps.Telefono = !Telefono
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!BaseDatos) Then mProps.BaseDatos = !BaseDatos
        If Not IsNull(!BandCliente) Then mProps.BandCliente = !BandCliente
        If Not IsNull(!BandProveedor) Then mProps.BandProveedor = !BandProveedor
        If Not IsNull(!BandInventario) Then mProps.BandInventario = !BandInventario
        If Not IsNull(!servidor) Then mProps.servidor = !servidor
        If Not IsNull(!NumPuntos) Then mProps.NumPuntos = !NumPuntos
        If Not IsNull(!BandMatriz) Then mProps.BandMatriz = !BandMatriz
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    
    'Verifica si ya existe una transacción usando el mismo
    sql = "SELECT CodTrans FROM GNTrans  WHERE IdSucursal=" & mProps.IDSucursal
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "GNSucursal.Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close

    'Elimina el registro
    sql = "DELETE FROM GNSucursal WHERE IdSucursal=" & mProps.IDSucursal
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_SUC", "Sucursal Eliminado: " & _
                                                         mProps.CodSucursal, "GN"

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_GNSUCURSAL
End Sub

Friend Function Clone() As GNSucursal
    Dim nuevo As GNSucursal

    'Copia las propiedades al nuevo objeto
    Set nuevo = New GNSucursal
    nuevo.Props = mProps

    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IDSucursal = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property





Public Property Let CodSucursal(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "GNSucursal_CodSucursal", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodSucursal = value
    mbooModificado = True
End Property

Public Property Get CodSucursal() As String
    CodSucursal = Trim$(mProps.CodSucursal)
End Property

Public Property Let Nombre(ByVal value As String)
    If Len(value) > 40 Then
        Err.Raise ERR_INVALIDO, "CodSucursal_Nombre", MSGERR_INVALIDO
    End If
    
    mProps.Nombre = value
    mbooModificado = True
End Property

Public Property Get Nombre() As String
    Nombre = Trim$(mProps.Nombre)
End Property

Public Property Let BandValida(ByVal value As Boolean)
    mProps.BandValida = value
    mbooModificado = True
End Property

Public Property Get BandValida() As Boolean
    BandValida = mProps.BandValida
End Property




Public Property Get fechagrabado() As Date
    fechagrabado = mProps.fechagrabado
End Property

Public Property Get EsNuevo() As Boolean
    EsNuevo = mbooEsNuevo
End Property

Public Property Get Modificado() As Boolean
    Modificado = mbooModificado
End Property


Private Sub Class_Initialize()
    'Valores iniciales
    mProps.BandValida = True
    
    mbooEsNuevo = True
    mbooModificado = False
    
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub



Public Property Let NumEstablecimiento(ByVal value As String)
    If Len(value) > 3 Then
        Err.Raise ERR_INVALIDO, "CodSucursal_NumEstablecimiento", MSGERR_INVALIDO
    End If
    
    mProps.NumEstablecimiento = value
    mbooModificado = True
End Property

Public Property Get NumEstablecimiento() As String
    NumEstablecimiento = Trim$(mProps.NumEstablecimiento)
End Property

Public Property Let Direccion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "CodSucursal_Direccion", MSGERR_INVALIDO
    End If
    
    mProps.Direccion = value
    mbooModificado = True
End Property

Public Property Get Direccion() As String
    Direccion = Trim$(mProps.Direccion)
End Property

Public Property Let Telefono(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "CodSucursal_Telefono", MSGERR_INVALIDO
    End If
    
    mProps.Telefono = value
    mbooModificado = True
End Property

Public Property Get Telefono() As String
    Telefono = Trim$(mProps.Telefono)
End Property



Public Property Let BaseDatos(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "CodSucursal_BaseDatos", MSGERR_INVALIDO
    End If
    
    mProps.BaseDatos = value
    mbooModificado = True
End Property

Public Property Get BaseDatos() As String
    BaseDatos = Trim$(mProps.BaseDatos)
End Property


Public Property Let BandCliente(ByVal value As Boolean)
    mProps.BandCliente = value
    mbooModificado = True
End Property

Public Property Get BandCliente() As Boolean
    BandCliente = mProps.BandCliente
End Property

Public Property Let BandProveedor(ByVal value As Boolean)
    mProps.BandProveedor = value
    mbooModificado = True
End Property

Public Property Get BandProveedor() As Boolean
    BandProveedor = mProps.BandProveedor
End Property

Public Property Let BandInventario(ByVal value As Boolean)
    mProps.BandInventario = value
    mbooModificado = True
End Property

Public Property Get BandInventario() As Boolean
    BandInventario = mProps.BandInventario
End Property

Public Property Let servidor(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "CodSucursal_Servidor", MSGERR_INVALIDO
    End If
    
    mProps.servidor = value
    mbooModificado = True
End Property

Public Property Get servidor() As String
    servidor = Trim$(mProps.servidor)
End Property

Public Property Let NumPuntos(ByVal value As Integer)
    mProps.NumPuntos = value
    mbooModificado = True
End Property

Public Property Get NumPuntos() As Integer
    NumPuntos = Trim$(mProps.NumPuntos)
End Property

Public Property Let BandMatriz(ByVal value As Boolean)
    mProps.BandMatriz = value
    mbooModificado = True
End Property

Public Property Get BandMatriz() As Boolean
    BandMatriz = mProps.BandMatriz
End Property

Public Function VerificaMatriz() As Boolean
    Dim sql As String, rs As Recordset
    VerificaMatriz = True
    sql = "SELECT CodSucursal FROM GNSucursal " & _
          "WHERE bandmatriz=1"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then    'Si ya existe el codigo
        VerificaMatriz = False
    End If
    rs.Close
End Function
