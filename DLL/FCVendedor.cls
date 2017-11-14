VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FCVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Empresa"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Type T_PROPS
    IdVendedor As Long
    CodVendedor As String
    Nombre As String
    BandValida As Boolean
    BandVendedor As Boolean     'jeaa 27/05/2007
    BandCobrador As Boolean     'jeaa 27/05/2007
    IdCuentaFaltante As Long
    IdCuentaSobrante As Long
    CodCuentaFaltante As String
    CodCuentaSobrante As String        '*** MAKOTO 12/feb/01 Agregado
    fechagrabado As Date
    Telefono As String
    EMail As String
    codUsuario As String
    IdCotizador As Long
    CodCotizador As String
    TipoTabla As Integer
    Cargo As String
    LineasTablet As String
    VendedoresTablet As String
    BandTodoIVG As Boolean
    BandTodoVende As Boolean
    OrdenBodegas As String
    RutasTablet As String
    IdTablaComision As Long
    CodTablaComision As String
    IdCobrador As Long
    CodCobrador As String
    IVG2Tablet As String
    BandTodoIVG2 As Boolean
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
    sql = "SELECT CodVendedor FROM FCVendedor " & _
          "WHERE (CodVendedor='" & mProps.CodVendedor & "') AND " & _
                "(IdVendedor<>" & mProps.IdVendedor & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "FCVendedor", MSGERR_REPITECODIGO
    End If
    rs.Close
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM FCVendedor WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM FCVendedor WHERE IdVendedor=" & mProps.IdVendedor
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "FCVendedor.Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro
    With rs
        !CodVendedor = mProps.CodVendedor
        !Nombre = mProps.Nombre
        !BandValida = mProps.BandValida
        !BandVendedor = mProps.BandVendedor
        !BandCobrador = mProps.BandCobrador
        !IdCuentaFaltante = mProps.IdCuentaFaltante
        !IdCuentaSobrante = mProps.IdCuentaSobrante
        !Telefono = mProps.Telefono
        !EMail = mProps.EMail
        !codUsuario = mProps.codUsuario
        !IdCotizador = mProps.IdCotizador
        !TipoTabla = mProps.TipoTabla
        !Cargo = mProps.Cargo
        !LineasTablet = mProps.LineasTablet
        !VendedoresTablet = mProps.VendedoresTablet
        !BandTodoIVG = mProps.BandTodoIVG
        !BandTodoVende = mProps.BandTodoVende
        !OrdenBodegas = mProps.OrdenBodegas
        !RutasTablet = mProps.RutasTablet
        !IdTablaComision = mProps.IdTablaComision
        !IVG2Tablet = mProps.IVG2Tablet
        !BandTodoIVG2 = mProps.BandTodoIVG2
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IdVendedor = !IdVendedor
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_VEND", "Vendedor NuevO: " & _
                                                         mProps.CodVendedor, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_VEND", "Vendedor Modificado: " & _
                                                         mProps.CodVendedor, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    If Me.BandVendedor Then
        mProps.objEmpresa.RaiseRefreshEvent REFRESH_VENDEDOR
    Else
        mProps.objEmpresa.RaiseRefreshEvent REFRESH_COBRADOR
    End If
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT FCVendedor.* , ctf.codcuenta as CodCuentaFaltante, cts.codcuenta as CodCuentaSobrante, "
        sql = sql & " FCVendedor.Telefono , FCVendedor.email, FCVendedor.CodUsuario,fcc.idvendedor as idcotizador, "
        sql = sql & " fcc.codvendedor as codcotizador, FCVendedor.TipoTabla, FCVendedor.Cargo, "
        sql = sql & " fcvtc.codTablaComision "
        sql = sql & " FROM FCVendedor "
        sql = sql & " left join ctcuenta ctf on FCVendedor.idcuentafaltante=ctf.idcuenta"
        sql = sql & " left join ctcuenta cts on FCVendedor.idcuentaSobrante=cts.idcuenta"
        sql = sql & " left join fcvendedor fcc on FCVendedor.idcotizador=fcc.idVendedor"
        sql = sql & " left join FCVTablaComision fcvtc on FCVendedor.idtablacomision=fcvtc.idtablacomision"
        sql = sql & " WHERE FCVendedor.CodVendedor='" & key & "'"
    Else
        sql = "SELECT FCVendedor.*, ctf.codcuenta as CodCuentaFaltante, cts.codcuenta as CodCuentaSobrante,   "
        sql = sql & " FCVendedor.Telefono , FCVendedor.email, FCVendedor.CodUsuario,fcc.idvendedor as idcotizador , "
        sql = sql & " fcc.codvendedor as codcotizador , FCVendedor.TipoTabla, FCVendedor.Cargo, "
        sql = sql & " fcvtc.codTablaComision "
        sql = sql & " FROM FCVendedor "
        sql = sql & " left join ctcuenta ctf on FCVendedor.idcuentafaltante=ctf.idcuenta"
        sql = sql & " left join ctcuenta cts on FCVendedor.idcuentaSobrante=cts.idcuenta"
        sql = sql & " left join fcvendedor fcc on FCVendedor.idcotizador=fcc.idVendedor"
        sql = sql & " left join FCVTablaComision fcvtc on FCVendedor.idtablacomision=fcvtc.idtablacomision"
        sql = sql & " WHERE FCVendedor.IdVendedor=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'Recupera los datos
    With rs
        mProps.IdVendedor = !IdVendedor
        mProps.CodVendedor = !CodVendedor
        If Not IsNull(!Nombre) Then mProps.Nombre = !Nombre
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!BandVendedor) Then mProps.BandVendedor = !BandVendedor
        If Not IsNull(!BandCobrador) Then mProps.BandCobrador = !BandCobrador
        If Not IsNull(!IdCuentaFaltante) Then mProps.IdCuentaFaltante = !IdCuentaFaltante
        If Not IsNull(!IdCuentaSobrante) Then mProps.IdCuentaSobrante = !IdCuentaSobrante
        If Not IsNull(!CodCuentaFaltante) Then mProps.CodCuentaFaltante = !CodCuentaFaltante
        If Not IsNull(!CodCuentaSobrante) Then mProps.CodCuentaSobrante = !CodCuentaSobrante
        If Not IsNull(!Telefono) Then mProps.Telefono = !Telefono
        If Not IsNull(!EMail) Then mProps.EMail = !EMail
        If Not IsNull(!codUsuario) Then mProps.codUsuario = !codUsuario
        
        If Not IsNull(!IdCotizador) Then mProps.IdCotizador = !IdCotizador
        If Not IsNull(!CodCotizador) Then mProps.CodCotizador = !CodCotizador
        
        If Not IsNull(!TipoTabla) Then mProps.TipoTabla = !TipoTabla
        
        If Not IsNull(!Cargo) Then mProps.Cargo = !Cargo
        If Not IsNull(!LineasTablet) Then mProps.LineasTablet = !LineasTablet
        If Not IsNull(!VendedoresTablet) Then mProps.VendedoresTablet = !VendedoresTablet
        If Not IsNull(!BandTodoIVG) Then mProps.BandTodoIVG = !BandTodoIVG
        If Not IsNull(!BandTodoVende) Then mProps.BandTodoVende = !BandTodoVende
        If Not IsNull(!OrdenBodegas) Then mProps.OrdenBodegas = !OrdenBodegas
        If Not IsNull(!RutasTablet) Then mProps.RutasTablet = !RutasTablet
        
        If Not IsNull(!IdTablaComision) Then mProps.IdTablaComision = !IdTablaComision
        If Not IsNull(!CodTablaComision) Then mProps.CodTablaComision = !CodTablaComision
        If Not IsNull(!IVG2Tablet) Then mProps.IVG2Tablet = !IVG2Tablet
        If Not IsNull(!BandTodoIVG2) Then mProps.BandTodoIVG2 = !BandTodoIVG2
        
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
    sql = "SELECT TransID FROM GNComprobante  WHERE IdVendedor=" & mProps.IdVendedor
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "FCVendedor.Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close

    'Elimina el registro
    sql = "DELETE FROM FCVendedor WHERE IdVendedor=" & mProps.IdVendedor
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_VEND", "Vendedor Eliminado: " & _
                                                         mProps.CodVendedor, "DT"

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_VENDEDOR
End Sub

Friend Function Clone() As FCVendedor
    Dim nuevo As FCVendedor
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New FCVendedor
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdVendedor = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Let CodVendedor(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "FCVendedor_CodVendedor", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodVendedor = value
    mbooModificado = True
End Property

Public Property Get CodVendedor() As String
    CodVendedor = Trim$(mProps.CodVendedor)
End Property

Public Property Let Nombre(ByVal value As String)
    If Len(value) > 40 Then
        Err.Raise ERR_INVALIDO, "FCVendedor_Nombre", MSGERR_INVALIDO
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

Public Property Let BandVendedor(ByVal value As Boolean)
    mProps.BandVendedor = value
    mbooModificado = True
End Property

Public Property Get BandVendedor() As Boolean
    BandVendedor = mProps.BandVendedor
End Property

Public Property Let BandCobrador(ByVal value As Boolean)
    mProps.BandCobrador = value
    mbooModificado = True
End Property

Public Property Get BandCobrador() As Boolean
    BandCobrador = mProps.BandCobrador
End Property

Public Property Let IdCuentaFaltante(ByVal value As Long)
    mProps.IdCuentaFaltante = value
    mbooModificado = True
End Property

Public Property Get IdCuentaFaltante() As Long
    IdCuentaFaltante = mProps.IdCuentaFaltante
End Property


Public Property Let CodCuentaFaltante(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodCuentaFaltante Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaFaltante también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "FcVendedor.CodCuentaFaltante", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaFaltante = rs!IdCuenta
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdCuentaFaltante = 0
        End If
    End If

    mProps.CodCuentaFaltante = value
    mbooModificado = True
End Property

Public Property Get CodCuentaFaltante() As String
    CodCuentaFaltante = Trim$(mProps.CodCuentaFaltante)
End Property

'*** MAKOTO 12/feb/01 Agregado
Public Property Let IdCuentaSobrante(ByVal value As Long)
    mProps.IdCuentaSobrante = value
    mbooModificado = True
End Property

Public Property Get IdCuentaSobrante() As Long
    IdCuentaSobrante = mProps.IdCuentaSobrante
End Property

'*** MAKOTO 12/feb/01 Agregado
Public Property Let CodCuentaSobrante(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodCuentaSobrante Then
        If Len(value) > 0 Then
            'Actualiza IdCuentaSobrante también
            sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "FcVendedor.CodCuentaSobrante2", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCuentaSobrante = rs!IdCuenta
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdCuentaSobrante = 0
        End If
    End If

    mProps.CodCuentaSobrante = value
    mbooModificado = True
End Property

Public Property Get CodCuentaSobrante() As String
    CodCuentaSobrante = Trim$(mProps.CodCuentaSobrante)
End Property

Public Property Let Telefono(ByVal value As String)
    If Len(value) > 40 Then
        Err.Raise ERR_INVALIDO, "FCVendedor_Telefono", MSGERR_INVALIDO
    End If
    
    mProps.Telefono = value
    mbooModificado = True
End Property

Public Property Get Telefono() As String
    Telefono = Trim$(mProps.Telefono)
End Property

Public Property Let EMail(ByVal value As String)
    If Len(value) > 40 Then
        Err.Raise ERR_INVALIDO, "FCVendedor_email", MSGERR_INVALIDO
    End If
    
    mProps.EMail = value
    mbooModificado = True
End Property

Public Property Get EMail() As String
    EMail = Trim$(mProps.EMail)
End Property

Public Property Let codUsuario(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "FCVendedor_CodUsuario", MSGERR_INVALIDO
    End If
    
    mProps.codUsuario = value
    mbooModificado = True
End Property

Public Property Get codUsuario() As String
    codUsuario = Trim$(mProps.codUsuario)
End Property

Public Property Let IdCotizador(ByVal value As Long)
    mProps.IdCotizador = value
    mbooModificado = True
End Property

Public Property Get IdCotizador() As Long
    IdCotizador = mProps.IdCotizador
End Property


Public Property Let CodCotizador(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodCotizador Then
        If Len(value) > 0 Then
            'Actualiza IdCotizador también
            sql = "SELECT IdVendedor FROM Fcvendedor WHERE CodVendedor='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "FcVendedor.CodCotizador", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCotizador = rs!IdVendedor
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdCotizador = 0
        End If
    End If

    mProps.CodCotizador = value
    mbooModificado = True
End Property

Public Property Get CodCotizador() As String
    CodCotizador = Trim$(mProps.CodCotizador)
End Property

Public Property Let TipoTabla(ByVal value As Integer)
    
    mProps.TipoTabla = value
    mbooModificado = True
End Property

Public Property Get TipoTabla() As Integer
    TipoTabla = Trim$(mProps.TipoTabla)
End Property

Public Property Let Cargo(ByVal value As String)
    If Len(value) > 40 Then
        Err.Raise ERR_INVALIDO, "FCVendedor_Cargo", MSGERR_INVALIDO
    End If
    mProps.Cargo = value
    mbooModificado = True
End Property

Public Property Get Cargo() As String
    Cargo = Trim$(mProps.Cargo)
End Property

Public Property Let LineasTablet(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "FCVendedor_LineasTablet", MSGERR_INVALIDO
    End If
    mProps.LineasTablet = value
    mbooModificado = True
End Property

Public Property Get LineasTablet() As String
    LineasTablet = Trim$(mProps.LineasTablet)
End Property

Public Property Let VendedoresTablet(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "FCVendedor_VendedoresTablet", MSGERR_INVALIDO
    End If
    mProps.VendedoresTablet = value
    mbooModificado = True
End Property

Public Property Get VendedoresTablet() As String
    VendedoresTablet = Trim$(mProps.VendedoresTablet)
End Property


Public Property Let BandTodoIVG(ByVal value As Boolean)
    mProps.BandTodoIVG = value
    mbooModificado = True
End Property


Public Property Get BandTodoIVG() As Boolean
    BandTodoIVG = mProps.BandTodoIVG
End Property

Public Property Let BandTodoVende(ByVal value As Boolean)
    mProps.BandTodoVende = value
    mbooModificado = True
End Property


Public Property Get BandTodoVende() As Boolean
    BandTodoVende = mProps.BandTodoVende
End Property

Public Property Get OrdenBodegas() As String
    OrdenBodegas = Trim$(mProps.OrdenBodegas)
End Property

Public Property Let OrdenBodegas(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "FCVendedor.OrdenBodegas", MSGERR_INVALIDO
    End If
    mProps.OrdenBodegas = value
    mbooModificado = True
End Property



Public Property Let RutasTablet(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "FCVendedor_RutasTablet", MSGERR_INVALIDO
    End If
    mProps.RutasTablet = value
    mbooModificado = True
End Property

Public Property Get RutasTablet() As String
    RutasTablet = Trim$(mProps.RutasTablet)
End Property

Public Property Let IdTablaComision(ByVal value As Long)
    mProps.IdTablaComision = value
    mbooModificado = True
End Property

Public Property Get IdTablaComision() As Long
    IdTablaComision = mProps.IdTablaComision
End Property


Public Property Let CodTablaComision(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodTablaComision Then
        If Len(value) > 0 Then
            'Actualiza IdTablaComision también
            sql = "SELECT IdTablaComision FROM FCVTablaComision WHERE CodTablaComision='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "FCVTablaComision.CodTablaComision", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdTablaComision = rs!IdTablaComision
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdTablaComision = 0
        End If
    End If

    mProps.CodTablaComision = value
    mbooModificado = True
End Property

Public Property Get CodTablaComision() As String
    CodTablaComision = Trim$(mProps.CodTablaComision)
End Property

Public Property Let IVG2Tablet(ByVal value As String)
    If Len(value) > 1200 Then
        Err.Raise ERR_INVALIDO, "FCVendedor_IVG2Tablet", MSGERR_INVALIDO
    End If
    mProps.IVG2Tablet = value
    mbooModificado = True
End Property

Public Property Get IVG2Tablet() As String
    IVG2Tablet = Trim$(mProps.IVG2Tablet)
End Property
Public Property Let BandTodoIVG2(ByVal value As Boolean)
    mProps.BandTodoIVG2 = value
    mbooModificado = True
End Property

Public Property Get BandTodoIVG2() As Boolean
    BandTodoIVG2 = mProps.BandTodoIVG2
End Property
