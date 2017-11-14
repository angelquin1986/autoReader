VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GNCentroCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    idCentro As Long
    CodCentro As String
    Descripcion As String
    Nombre As String            '*** MAKOTO 14/feb/01 Agregado
    FechaInicio As Date
    FechaFinal As Variant       'Variant para almacenar Null
    
    '*** MAKOTO 05/mar/01 Agregado
    idProveedor As Long
    CodProveedor As String
    IdCliente As Long
    CodCliente As String
    Observacion As String   'jeaa 21/01/2009
    Valor1 As Currency
    Valor2 As Currency
    Valor3 As Currency
    IdGrupo(1 To 5) As Long
    CodGrupo(1 To 5) As String
    
    IdVehiculo As Long
    CodVehiculo As String
    
    FechaEntrega As Date
    BandOcupado As Boolean 'AUC bandera para fichas medicas
    IdTransporte As Long
    CodTransporte As String
    IdTipoHab As Long
    Codhab As String
    fechagrabado As Date
    IdInventario As Long
    CodInventario As String
    
    'gualaTV
    IdGar As Long
    CodGar As String
    Referencia As String
    IdGAR2 As Long
    CodGAR2 As String
    IDCallePri As Long
    CodCallePri As String
    IdCalleSec As Long
    CodCalleSec As String
    NumCasa As String
    posGoogleMaps As String
    Referencia2 As String
    IdPais As Long
    CodPais As String
    IdProvincia As Long
    CodProvincia As String
    IdCanton As Long
    CodCanton As String
    IDParroquia As Long
    CodParroqui As String
    Telefono2 As String
    Propietario As String
    BandCasaPropia As String
    IdZona As Long
    CodZona As String
    IdSubZona As Long
    CodSubZona As String
    IdVendedor As Long
    CodVendedor As String
    NumTV As Integer
    Referencia3 As String
    BandCorteMora As Boolean
    
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean


Public Sub Grabar()
    Dim sql As String, rs As Recordset, i As Integer
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
    sql = "SELECT CodCentro FROM GNCentroCosto " & _
          "WHERE (CodCentro='" & mProps.CodCentro & "') AND " & _
          "(IdCentro<>" & mProps.idCentro & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "GNCentroCosto", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM GNCentroCosto WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM GNCentroCosto WHERE IdCentro=" & mProps.idCentro
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "GNCentroCosto_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en GNCentroCosto
    With rs
        !CodCentro = mProps.CodCentro
        !Descripcion = mProps.Descripcion
        !Nombre = mProps.Nombre             '*** MAKOTO 14/feb/01 Agregado
        !FechaInicio = mProps.FechaInicio
        

        
        If Not IsNull(mProps.FechaFinal) Then
            If IsNull(!FechaFinal) Then
                Me.Empresa.GrabaAccionModulo "NULL", mProps.FechaFinal, "FechaFinal", "MOD_CC", "CL"
            End If
        Else
            If Not IsNull(!FechaFinal) Then
                Me.Empresa.GrabaAccionModulo !FechaFinal, "NULL", "FechaFinal", "MOD_CC", "CL"
            End If
        End If

        
        
        If Not IsNull(mProps.FechaFinal) Then
            !FechaFinal = mProps.FechaFinal
        Else
            !FechaFinal = Null
        End If
        
        
        
        '*** MAKOTO 05/mar/01 Agregado
        !idProveedor = mProps.idProveedor
        !IdCliente = mProps.IdCliente
        !Observacion = mProps.Observacion
        !Valor1 = mProps.Valor1
        !Valor2 = mProps.Valor2
        !Valor3 = mProps.Valor3
        !FechaEntrega = mProps.FechaEntrega
        
        For i = 1 To GNCGRUPO_MAX
            .Fields("IdGrupo" & i) = mProps.IdGrupo(i)
        Next i
        .Fields("IdVehiculo") = mProps.IdVehiculo
        .Fields("BandOcupado") = mProps.BandOcupado
        .Fields("idtransporte") = mProps.IdTransporte
        .Fields("idTipoHab") = mProps.IdTipoHab
        .Fields("Idinventario") = mProps.IdInventario
        '.Fields("codinventario") = mProps.CodInventario
        
        !IdGar = mProps.IdGar
        !IdGAR2 = mProps.IdGAR2
        !Referencia = mProps.Referencia
        !IDCallePri = mProps.IDCallePri
        !IdCalleSec = mProps.IdCalleSec
        !NumCasa = mProps.NumCasa
        !posGoogleMaps = mProps.posGoogleMaps
        !Referencia2 = mProps.Referencia2
        !IdPais = mProps.IdPais
        !IdProvincia = mProps.IdProvincia
        !IdCanton = mProps.IdCanton
        !IDParroquia = mProps.IDParroquia
        !Telefono2 = mProps.Telefono2
        !Propietario = mProps.Propietario
        !BandCasaPropia = mProps.BandCasaPropia
        !IdZona = mProps.IdZona
        !IdSubZona = mProps.IdSubZona
        !IdVendedor = mProps.IdVendedor
        !NumTV = mProps.NumTV
        !Referencia3 = mProps.Referencia3
        !BandCorteMora = mProps.BandCorteMora
        
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.idCentro = !idCentro
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_CC", "Centro de Costo Nuevo: " & _
                                                         mProps.CodCentro, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_CC", "Centro de Costo Modificado: " & _
                                                         mProps.CodCentro, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_CENTROCOSTO
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset, i As Integer
    
    If VarType(key) = vbString Then
        'sql = "SELECT * FROM GNCentroCosto WHERE CodCentro='" & key & "'"
        sql = "SELECT gnc.*, "
        'sql = sql & " IdCentro, CodCentro, gnc.Descripcion , FechaInicio,"
        'sql = sql & " FechaFinal, gnc.FechaGrabado, gnc.Nombre, gnc.IdProveedor,"
        'sql = sql & " gnc.IdCliente, gnc.Observacion, "
        sql = sql & " Valor1 , Valor2, Valor3, CodVehiculo, "
        'sql = sql & " gnc.idvehiculo, CodVehiculo,gnc.FechaEntrega,gnc.nombre ,gnc.bandocupado"
        sql = sql & " gnt.idtransporte,gnt.codtransporte"
        sql = sql & " ,th.idTipo,th.Codigo,iv.idinventario,iv.codinventario, codvendedor  "
        sql = sql & ", PC.CODPROVCLI AS  CodProvCli , PCG.CODPROVCLI AS  CodGar,  PCG2.CODPROVCLI AS  CodGar2,"
        sql = sql & " pccallePri.codCalle as CodCallePri, pccalleSec.codCalle as CodCalleSec, "
        sql = sql & " pcp.codpais, pcpro.codProvincia, pccan.codcanton, pcparr.codparroquia, gnz.Codzona, gncomp.codcompetencia as CodSubZona  "
        
        
        sql = sql & " FROM GNCentroCosto gnc"
        sql = sql & " left join  GNVehiculo gnv on gnc.idvehiculo=gnv.idvehiculo"
        sql = sql & " left join  GNTransporte gnt on gnt.idtransporte=gnc.idTransporte"
        sql = sql & " left join  TipoHabitacion TH left join  Ivinventario iv on iv.idinventario=th.idinventario on tH.idTIPO=gnc.idTipohab"
        
        sql = sql & " left JOIN PCPROVCLI PC ON GnC.IDcliente = PC.IDPROVCLI"
        sql = sql & " left JOIN PCPROVCLI PCG ON GnC.IDGar = PCG.IDPROVCLI"
        sql = sql & " left JOIN PCPROVCLI PCG2 ON GnC.IDGar2 = PCG2.IDPROVCLI"
        sql = sql & " left join pcCalle pcCallePri on gnc.IdCallePri=pcCallePri.IdCAlle"
        sql = sql & " left join pcCalle pcCalleSec on gnc.IdCalleSec=pcCalleSec.IdCAlle"
        sql = sql & " left join pcPais pcp on gnc.Idpais=pcp.IdPais"
        sql = sql & " left join pcProvincia pcpro on gnc.Idprovincia=pcpro.IdProvincia"
        sql = sql & " left join pcCanton pccan on gnc.Idcanton=pccan.IdCanton"
        sql = sql & " left join pcParroquia pcparr on gnc.IdParroquia=pcparr.IdParroquia"
        sql = sql & " left join GNZona gnz on gnc.IdZona=gnz.IdZona"
        sql = sql & " left join GNCompetencia gncomp on gnc.IdSubZona=gncomp.IdCompetencia "
        sql = sql & " left join FcVendedor fc on gnc.IdVendedor=fc.IdVendedor"
        
        sql = sql & " WHERE CodCentro='" & key & "'"
        
    Else
        'sql = "SELECT * FROM GNCentroCosto WHERE IdCentro=" & key
        sql = "SELECT "
        sql = sql & " IdCentro, CodCentro, gnc.Descripcion, FechaInicio,"
        sql = sql & " FechaFinal, gnc.FechaGrabado, gnc.Nombre, gnc.IdProveedor,"
        sql = sql & " gnc.IdCliente, gnc.Observacion, valor1, valor2, valor3,"
        sql = sql & " gnc.IdGrupo1, g1.codgrupo1,G1.descripcion,"
        sql = sql & " gnc.IdGrupo2, g2.codgrupo2,G2.descripcion,"
        sql = sql & " gnc.IdGrupo3, g3.codgrupo3,G3.descripcion,"
        sql = sql & " gnc.IdGrupo4, g4.codgrupo4,G4.descripcion, "
        sql = sql & " gnc.idvehiculo, CodVehiculo,FechaEntrega "
        sql = sql & " ,gnc.bandocupado,gnt.idtransporte,gnt.codtransporte"
        sql = sql & " ,th.idTipo,th.Codigo,iv.idinventario,iv.codinventario"
        sql = sql & ", PC.CODPROVCLI AS  CodProvCli , PCG.CODPROVCLI AS  CodGar,  PCG2.CODPROVCLI AS  CodGar2,"
        sql = sql & " pccallePri.codCalle as CodCallePri, pccalleSec.codCalle as CodCalleSec, "
        sql = sql & " pcp.codpais, pcpro.codProvincia, pccan.codcanton, pcparr.codparroquia, gnz.Codzona "
        
        
        
        sql = sql & " FROM GNCentroCosto gnc"
        sql = sql & " left join GNCGrupo1 g1 on gnc.idgrupo1=g1.idgrupo1"
        sql = sql & " left join GNCGrupo2 g2 on gnc.idgrupo2=g2.idgrupo2"
        sql = sql & " left join GNCGrupo3 g3 on gnc.idgrupo3=g3.idgrupo3"
        sql = sql & " left join GNCGrupo4 g4 on gnc.idgrupo4=g4.idgrupo4"
        sql = sql & " left join  GNVehiculo gnv on gnc.idvehiculo=gnv.idvehiculo"
        sql = sql & " left join  GNTransporte gnt on gnt.idtransporte=gnc.idTransporte"
        sql = sql & " left join  TipoHabitacion TH left join  Ivinventario iv on iv.idinventario=th.idinventario on tH.idTIPO=gnc.idTipohab"
        
        sql = sql & " left JOIN PCPROVCLI PC ON GC.IDPROVCLI = PC.IDPROVCLI"
        sql = sql & " left JOIN PCPROVCLI PCG ON GC.IDGar = PCG.IDPROVCLI"
        sql = sql & " left JOIN PCPROVCLI PCG2 ON GC.IDGar2 = PCG2.IDPROVCLI"
        sql = sql & " left join pcCalle pcCallePri on gc.IdCallePri=pcCallePri.IdCAlle"
        sql = sql & " left join pcCalle pcCalleSec on gc.IdCalleSec=pcCalleSec.IdCAlle"
        sql = sql & " left join pcPais pcp on gc.Idpais=pcp.IdPais"
        sql = sql & " left join pcProvincia pcpro on gc.Idprovincia=pcpro.IdProvincia"
        sql = sql & " left join pcCanton pccan on gc.Idcanton=pccan.IdCanton"
        sql = sql & " left join pcParroquia pcparr on gc.IdParroquia=pcparr.IdParroquia"
        sql = sql & " left join GNZona gnz on gc.IdZona=gnz.IdZona"
        
        
        sql = sql & " WHERE IdCentro=" & key
        
    End If

    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.idCentro = !idCentro
        mProps.CodCentro = !CodCentro
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!Nombre) Then mProps.Nombre = !Nombre     '*** MAKOTO 14/feb/01 Agregado
        If Not IsNull(!FechaInicio) Then mProps.FechaInicio = !FechaInicio
        mProps.FechaFinal = !FechaFinal         'Null tambien es valido
        
        '*** MAKOTO 05/mar/01 Agregado
        If Not IsNull(!idProveedor) Then idProveedor = !idProveedor
        If Not IsNull(!IdCliente) Then IdCliente = !IdCliente
        If Not IsNull(!Observacion) Then mProps.Observacion = !Observacion
        If Not IsNull(!Valor1) Then mProps.Valor1 = !Valor1
        If Not IsNull(!Valor2) Then mProps.Valor2 = !Valor2
        If Not IsNull(!Valor3) Then mProps.Valor3 = !Valor3

'        For i = 1 To GNCGRUPO_MAX
'            If Not IsNull(.Fields("IdGrupo" & i)) Then mProps.IdGrupo(i) = .Fields("IdGrupo" & i)
'            If Not IsNull(.Fields("CodGrupo" & i)) Then mProps.CodGrupo(i) = .Fields("CodGrupo" & i)
'        Next i
        
        If Not IsNull(.Fields("IdVehiculo")) Then mProps.IdVehiculo = .Fields("IdVehiculo")
        If Not IsNull(.Fields("CodVehiculo")) Then mProps.CodVehiculo = .Fields("CodVehiculo")
        If Not IsNull(.Fields("fechaentrega")) Then mProps.FechaEntrega = .Fields("fechaentrega")
        If Not IsNull(.Fields("BandOcupado")) Then mProps.BandOcupado = .Fields("BandOcupado")
        If Not IsNull(.Fields("IdTransporte")) Then mProps.IdTransporte = .Fields("IdTransporte")
        If Not IsNull(.Fields("CodTransporte")) Then mProps.CodTransporte = .Fields("CodTransporte")
        If Not IsNull(.Fields("idtipo")) Then mProps.IdTipoHab = .Fields("idtipo")
        If Not IsNull(.Fields("Codigo")) Then mProps.Codhab = .Fields("Codigo")
        If Not IsNull(.Fields("idInventario")) Then mProps.IdInventario = .Fields("idInventario")
        If Not IsNull(.Fields("codInventario")) Then mProps.CodInventario = .Fields("codInventario")
        
        If Not IsNull(!IdGar) Then IdGar = !IdGar
        If Not IsNull(!CodGar) Then CodGar = !CodGar
        If Not IsNull(!Referencia) Then mProps.Referencia = !Referencia
        If Not IsNull(!IdGAR2) Then IdGAR2 = !IdGAR2
        If Not IsNull(!CodGAR2) Then CodGAR2 = !CodGAR2
        If Not IsNull(!IDCallePri) Then mProps.IDCallePri = !IDCallePri
        If Not IsNull(!CodCallePri) Then mProps.CodCallePri = !CodCallePri
        If Not IsNull(!IdCalleSec) Then mProps.IdCalleSec = !IdCalleSec
        If Not IsNull(!CodCalleSec) Then mProps.CodCalleSec = !CodCalleSec
        If Not IsNull(!NumCasa) Then mProps.NumCasa = !NumCasa
        If Not IsNull(!posGoogleMaps) Then mProps.posGoogleMaps = !posGoogleMaps
        If Not IsNull(!Referencia2) Then mProps.Referencia2 = !Referencia2
        If Not IsNull(!IdPais) Then mProps.IdPais = !IdPais
        If Not IsNull(!CodPais) Then mProps.CodPais = !CodPais
        If Not IsNull(!IdProvincia) Then mProps.IdProvincia = !IdProvincia
        If Not IsNull(!CodProvincia) Then mProps.CodProvincia = !CodProvincia
        If Not IsNull(!IdCanton) Then mProps.IdCanton = !IdCanton
        If Not IsNull(!CodCanton) Then mProps.CodCanton = !CodCanton
        If Not IsNull(!IDParroquia) Then mProps.IDParroquia = !IDParroquia
        If Not IsNull(!CodParroquia) Then mProps.CodParroqui = !CodParroquia
        If Not IsNull(!Telefono2) Then mProps.Telefono2 = !Telefono2
        If Not IsNull(!Propietario) Then mProps.Propietario = !Propietario
        If Not IsNull(!BandCasaPropia) Then mProps.BandCasaPropia = !BandCasaPropia
        If Not IsNull(!IdZona) Then IdZona = !IdZona
        If Not IsNull(!CodZona) Then CodZona = !CodZona
        If Not IsNull(!IdSubZona) Then IdSubZona = !IdSubZona
        If Not IsNull(!CodSubZona) Then CodSubZona = !CodSubZona
        If Not IsNull(!IdVendedor) Then IdVendedor = !IdVendedor
        If Not IsNull(!CodVendedor) Then CodVendedor = !CodVendedor
        If Not IsNull(!NumTV) Then NumTV = !NumTV
        If Not IsNull(!Referencia3) Then mProps.Referencia3 = !Referencia3
        If Not IsNull(!BandCorteMora) Then mProps.BandCorteMora = !BandCorteMora
        
        
        
        
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    
    'Verifica si existe un comprobante que usa el codigo de centro de costo
    sql = "SELECT TransID FROM GNComprobante WHERE IdCentro=" & mProps.idCentro
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        rs.Close
        Set rs = Nothing
        Err.Raise ERR_NOELIMINABLE, "GNCentroCosto.Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close
    Set rs = Nothing
    
    'Elimina el Centro de costo
    sql = "DELETE FROM GNCentroCosto WHERE IdCentro=" & mProps.idCentro
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_CC", "Centro de Costo Eliminado: " & _
                                                         mProps.CodCentro, "DT"

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_CENTROCOSTO
End Sub

Friend Function Clone() As GNCentroCosto
    Dim nuevo As GNCentroCosto
    
    Set nuevo = New GNCentroCosto
    'Copia las propiedades al nuevo objeto
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.idCentro = 0
    mProps.FechaFinal = Null
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Let CodCentro(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNCentroCosto_CodCentro", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodCentro = value
    mbooModificado = True
End Property

Public Property Get CodCentro() As String
    CodCentro = Trim$(mProps.CodCentro)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "GNCentroCosto_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

'*** MAKOTO 14/feb/01 Agregado
Public Property Let Nombre(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "GNCentroCosto.Nombre", MSGERR_INVALIDO
    End If
    
    mProps.Nombre = value
    mbooModificado = True
End Property

'*** MAKOTO 14/feb/01 Agregado
Public Property Get Nombre() As String
    Nombre = Trim$(mProps.Nombre)
End Property

'*** MAKOTO 05/mar/01 Agregado
Public Property Let idProveedor(ByVal value As Long)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.idProveedor Then
        'Actualiza el Codigo también
        mProps.CodProveedor = ""
        sql = "SELECT CodProvCli FROM PCProvCli WHERE IdProvCli=" & value
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then mProps.CodProveedor = rs!CodProvcli
        rs.Close
        Set rs = Nothing
    End If
    
    mProps.idProveedor = value
    mbooModificado = True
End Property

'*** MAKOTO 05/mar/01 Agregado
Public Property Get idProveedor() As Long
    idProveedor = mProps.idProveedor
End Property

Public Property Let CodProveedor(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodProveedor Then
        If Len(value) > 0 Then
            'Actualiza IdProveedor también
            sql = "SELECT IdProvCli FROM PCProvCli WHERE CodProvCli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNCentroCosto.CodProveedor", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.idProveedor = rs!IdProvcli
            End If
            rs.Close
        Else
            mProps.idProveedor = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodProveedor = value
    mbooModificado = True
End Property

Public Property Get CodProveedor() As String
    CodProveedor = Trim$(mProps.CodProveedor)
End Property

'*** MAKOTO 05/mar/01 Agregado
Public Property Let IdCliente(ByVal value As Long)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.IdCliente Then
        'Actualiza el Codigo también
        mProps.CodCliente = ""
        sql = "SELECT CodProvCli FROM PCProvCli WHERE IdProvCli=" & value
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then mProps.CodCliente = rs!CodProvcli
        rs.Close
        Set rs = Nothing
    End If
    
    mProps.IdCliente = value
    mbooModificado = True
End Property

'*** MAKOTO 05/mar/01 Agregado
Public Property Get IdCliente() As Long
    IdCliente = mProps.IdCliente
End Property

Public Property Let CodCliente(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCliente Then
        If Len(value) > 0 Then
            'Actualiza IdCliente también
            sql = "SELECT IdProvCli FROM PCProvCli WHERE CodProvCli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNCentroCosto.CodCliente", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdCliente = rs!IdProvcli
            End If
            rs.Close
        Else
            mProps.IdCliente = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodCliente = value
    mbooModificado = True
End Property

Public Property Get CodCliente() As String
    CodCliente = Trim$(mProps.CodCliente)
End Property



Public Property Let FechaInicio(ByVal value As Date)
    mProps.FechaInicio = value
    mbooModificado = True
End Property

Public Property Get FechaInicio() As Date
    FechaInicio = mProps.FechaInicio
End Property

Public Property Let FechaFinal(ByVal value As Variant)
    If Not IsNull(value) Then
        If Not IsDate(value) Then
            Err.Raise ERR_INVALIDO, "GNCentroCosto_FechaFinal", MSGERR_INVALIDO
        Else
            mProps.FechaFinal = CDate(value)
        End If
    Else
        mProps.FechaFinal = Null
    End If
    
    mbooModificado = True
End Property

Public Property Get FechaFinal() As Variant
    FechaFinal = mProps.FechaFinal
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

    'Valores predeterminados
    With mProps
        .FechaInicio = Date
        .FechaFinal = Null
        .FechaEntrega = Date
        .NumTV = 1
    End With
    
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub


Public Sub GeneraCodCentro()
    Dim sql As String, pre As String, num As Double, rs As Recordset
    Dim lon_pre As Integer, i As Integer, v As Variant
    
    '-----
    ' Genera el código de centro en el siguiente formato
    '   XXXYYYYMMDD-nnn
    '     XXX = PrefijoCodigo de GNOpcion (Longitud puede variar)
    '     YYYYMMDD = Fecha actual
    '     nnn = Número secuencial dentro de la fecha
    '-----
    
    'Busca el valor maximo de los últimos 3 dígitos de CodCentro de la misma fecha
    pre = mProps.objEmpresa.GNOpcion.PrefijoCodCC
    lon_pre = Len(pre)
    
    Select Case Me.Empresa.GNOpcion.GeneraCodCC
    Case 0  'Manual
    Case 1  'Secuencial
        'Falta
#If DAOLIB Then
        sql = "SELECT Max(Right([CodCentro],3)) AS CodMax " & _
              "FROM GNCentroCosto " & _
              "WHERE (Left$([CodCentro]," & (Len(pre) + 8) & ")='" & _
                        pre & Format(Date, "yyyymmdd") & "')"
#Else
        sql = "SELECT case  " & _
                     "when isnumeric(substring(CodCentro," & (lon_pre + 1) & ",len(codcentro)))=1  " & _
                          "then substring(CodCentro," & (lon_pre + 1) & ",len(codcentro)) " & _
                     "when isnumeric(substring(CodCentro," & (lon_pre + 1) & ",len(codcentro)))=1  " & _
                          "then 0 " & _
                     "end AS Cod  INTO tmp0 " & _
              "FROM GNCentroCosto " & _
              "WHERE (Substring(CodCentro,1," & Len(pre) & ")='" & _
                        pre & "')  group by CodCentro"
#End If
        VerificaExistenciaTabla 0
        gobjMain.EmpresaActual.EjecutarSQL sql, 0
        
        sql = "SELECT MAX(Cod) as CodMax FROM tmp0"
        
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        
        If Not rs.EOF Then
            If Not IsNull(rs!CodMax) Then num = rs!CodMax
        End If
        
        CodCentro = pre & Format(num + 1, mProps.objEmpresa.GNOpcion.FormatoCodCC)
        rs.Close
        VerificaExistenciaTabla 0
    Case 2  'Con formato
#If DAOLIB Then
        sql = "SELECT Max(Right([CodCentro],3)) AS CodMax " & _
              "FROM GNCentroCosto " & _
              "WHERE (Left$([CodCentro]," & (Len(pre) + 8) & ")='" & _
                        pre & Format(Date, "yyyymmdd") & "')"
#Else
        sql = "SELECT Max(Right(CodCentro,3)) AS CodMax " & _
              "FROM GNCentroCosto " & _
              "WHERE (Substring(CodCentro,1," & (Len(pre) + 8) & ")='" & _
                        pre & Format(Date, "yyyymmdd") & "')"
#End If
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then
            If Not IsNull(rs!CodMax) Then num = Val(rs!CodMax)
        End If
        
        CodCentro = pre & Format(Date, "yyyymmdd") & "-" & Format(num + 1, mProps.objEmpresa.GNOpcion.FormatoCodCC)
        rs.Close
Case 3  'Con el Codigo de la Sucursal jeaa 30/03/2008
    If Len(GetSetting(APPNAME, SECTION, "CodSucursal_Activa_" & mProps.objEmpresa.CodEmpresa, "")) > 0 Then
        pre = GetSetting(APPNAME, SECTION, "CodSucursal_Activa_" & mProps.objEmpresa.CodEmpresa, "")
    Else
        MsgBox "No existe Sucursal Activa para este equipo" & Chr(13) & "Revice la configuracion en Información de la Empresa"
        Exit Sub
    End If
    lon_pre = Len(pre) + 1
#If DAOLIB Then
        sql = "SELECT Max(CodCentro)) AS CodMax " & _
              "FROM GNCentroCosto " & _
              "WHERE (Left$([CodCentro]," & (Len(pre) + 8) & ")='" & _
                        pre & Format(Date, "yyyymmdd") & "')"
#Else
        sql = "SELECT case  " & _
                     "when isnumeric(substring(CodCentro," & (lon_pre + 1) & ",len(codcentro)))=1  " & _
                          "then substring(CodCentro," & (lon_pre + 1) & ",len(codcentro)) " & _
                     "when isnumeric(substring(CodCentro," & (lon_pre + 1) & ",len(codcentro)))=1  " & _
                          "then 0 " & _
                     "end AS Cod  INTO tmp0 " & _
              "FROM GNCentroCosto " & _
              "WHERE (Substring(CodCentro,1," & Len(pre) & ")='" & _
                        pre & "')  group by CodCentro"
#End If

        VerificaExistenciaTabla 0
        gobjMain.EmpresaActual.EjecutarSQL sql, 0
        
        sql = "SELECT MAX(Cod) as CodMax FROM tmp0"
        
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        
        If Not rs.EOF Then
            If Not IsNull(rs!CodMax) Then num = rs!CodMax
        End If


        CodCentro = pre & "-" & Format(num + 1, mProps.objEmpresa.GNOpcion.FormatoCodCC)
        rs.Close
    End Select
    Set rs = Nothing
End Sub

Public Property Let Observacion(ByVal value As String)
    If Len(value) > 100 Then
        Err.Raise ERR_INVALIDO, "GNCentroCosto_Observacion", MSGERR_INVALIDO
    End If
    
    mProps.Observacion = value
    mbooModificado = True
End Property

Public Property Get Observacion() As String
    Observacion = Trim$(mProps.Observacion)
End Property

Public Property Let Valor1(ByVal value As Currency)
    mProps.Valor1 = value
    mbooModificado = True
End Property

Public Property Get Valor1() As Currency
    Valor1 = mProps.Valor1
End Property

Public Property Let Valor2(ByVal value As Currency)
    mProps.Valor2 = value
    mbooModificado = True
End Property

Public Property Get Valor2() As Currency
    Valor2 = mProps.Valor2
End Property

Public Property Let Valor3(ByVal value As Currency)
    mProps.Valor3 = value
    mbooModificado = True
End Property

Public Property Get Valor3() As Currency
    Valor3 = mProps.Valor3
End Property

Public Property Let IdGrupo(NumGrupo As Integer, ByVal value As Long)
    mProps.IdGrupo(NumGrupo) = value
    mbooModificado = True
End Property

Public Property Get IdGrupo(NumGrupo As Integer) As Long
    IdGrupo = mProps.IdGrupo(NumGrupo)
End Property


Public Property Let CodGrupo(NumGrupo As Integer, ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodGrupo(NumGrupo) Then
        If Len(value) > 0 Then
            'Actualiza IdGrupo1 también
            sql = "SELECT IdGrupo" & NumGrupo & " FROM GNCGrupo" & NumGrupo & " WHERE CodGrupo" & NumGrupo & "='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNCgRUPO_CodGrupo" & NumGrupo, MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdGrupo(NumGrupo) = rs.Fields("IdGrupo" & NumGrupo)
            End If
            rs.Close
        Else
            IdGrupo(NumGrupo) = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodGrupo(NumGrupo) = value
    mbooModificado = True
End Property

Public Property Get CodGrupo(NumGrupo As Integer) As String
    CodGrupo = Trim$(mProps.CodGrupo(NumGrupo))
End Property




Public Property Let IdVehiculo(ByVal value As Long)
    mProps.IdVehiculo = value
    mbooModificado = True
End Property

Public Property Get IdVehiculo() As Long
    IdVehiculo = mProps.IdVehiculo
End Property


Public Property Let CodVehiculo(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodVehiculo Then
        If Len(value) > 0 Then
            'Actualiza IdVehiculo1 también
            sql = "SELECT IdVehiculo FROM GNVehiculo WHERE CodVehiculo" & "='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNCVehiculo_CodVehiculo", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdVehiculo = rs.Fields("IdVehiculo")
            End If
            rs.Close
        Else
            IdVehiculo = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodVehiculo = value
    mbooModificado = True
End Property

Public Property Get CodVehiculo() As String
    CodVehiculo = Trim$(mProps.CodVehiculo)
End Property

Public Property Let FechaEntrega(ByVal value As Date)
    mProps.FechaEntrega = value
    mbooModificado = True
End Property

Public Property Get FechaEntrega() As Date
    FechaEntrega = mProps.FechaEntrega
End Property

Public Function ListaIVKardexTipo(ByVal Tipo As Integer) As Variant
    Dim sql As String, rs As Recordset
    Dim TransOP As String
        If Len(mProps.objEmpresa.GNOpcion.ObtenerValor("ProduccionTransOP")) > 0 Then
            TransOP = mProps.objEmpresa.GNOpcion.ObtenerValor("ProduccionTransOP")
        Else
            Err.Raise MSGERR_INVALIDO, "Empresa_TransOrdenProduccion", "No esta Configurado para tipo trans OP "
        End If
    
        sql = "spConsIVKardexTipoCCosto '" & TransOP & "','" & _
                                  Me.CodCentro & "', " & _
                                  Tipo
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
    Set ListaIVKardexTipo = rs
    Set rs = Nothing
End Function

Public Property Let BandOcupado(ByVal value As Boolean)
    mProps.BandOcupado = value
    mbooModificado = True
End Property

Public Property Get BandOcupado() As Boolean
    BandOcupado = mProps.BandOcupado
End Property

'---------
Public Property Let IdTransporte(ByVal value As Long)
    mProps.IdTransporte = value
    mbooModificado = True
End Property

Public Property Get IdTransporte() As Long
    IdTransporte = mProps.IdTransporte
End Property

Public Property Let CodTransporte(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodTransporte Then
        If Len(value) > 0 Then
            'Actualiza IdVehiculo1 también
            sql = "SELECT IdTransporte FROM GNTransporte WHERE CodTransporte" & "='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNTransporte_CodTransporte", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdTransporte = rs.Fields("IdTransporte")
            End If
            rs.Close
        Else
            IdTransporte = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodTransporte = value
    mbooModificado = True
End Property

Public Property Get CodTransporte() As String
    CodTransporte = Trim$(mProps.CodTransporte)
End Property

Friend Sub RecuperarxVeh(key As Variant)
    Dim sql As String, rs As Recordset, i As Integer
    
    If VarType(key) = vbString Then
        'sql = "SELECT * FROM GNCentroCosto WHERE CodCentro='" & key & "'"
        sql = "SELECT "
        sql = sql & " IdCentro, CodCentro, gnc.Descripcion , FechaInicio,"
        sql = sql & " FechaFinal, gnc.FechaGrabado, Nombre, IdProveedor,"
        sql = sql & " gnc.IdCliente, gnc.Observacion, valor1, valor2, valor3, "
        sql = sql & " gnc.idvehiculo, CodVehiculo,FechaEntrega,gnc.nombre"
        sql = sql & " ,gnc.bandocupado,gnt.idtransporte,gnt.codtransporte"
        sql = sql & " FROM GNCentroCosto gnc"
        sql = sql & " left join  GNVehiculo gnv on gnc.idvehiculo=gnv.idvehiculo"
        sql = sql & " left join  GNTransporte gnt on gnt.idtransporte=gnc.idTransporte"
        sql = sql & " WHERE CodVehiculo='" & key & "'"
        
    Else
        'sql = "SELECT * FROM GNCentroCosto WHERE IdCentro=" & key
        sql = "SELECT "
        sql = sql & " IdCentro, CodCentro, gnc.Descripcion, FechaInicio,"
        sql = sql & " FechaFinal, gnc.FechaGrabado, Nombre, IdProveedor,"
        sql = sql & " gnc.IdCliente, gnc.Observacion, valor1, valor2, valor3,"
        sql = sql & " gnc.IdGrupo1, g1.codgrupo1,G1.descripcion,"
        sql = sql & " gnc.IdGrupo2, g2.codgrupo2,G2.descripcion,"
        sql = sql & " gnc.IdGrupo3, g3.codgrupo3,G3.descripcion,"
        sql = sql & " gnc.IdGrupo4, g4.codgrupo4,G4.descripcion, "
        sql = sql & " gnc.idvehiculo, CodVehiculo,FechaEntrega "
        sql = sql & " ,gnc.bandocupado,gnt.idtransporte,gnt.codtransporte"
        sql = sql & " FROM GNCentroCosto gnc"
        sql = sql & " left join GNCGrupo1 g1 on gnc.idgrupo1=g1.idgrupo1"
        sql = sql & " left join GNCGrupo2 g2 on gnc.idgrupo2=g2.idgrupo2"
        sql = sql & " left join GNCGrupo3 g3 on gnc.idgrupo3=g3.idgrupo3"
        sql = sql & " left join GNCGrupo4 g4 on gnc.idgrupo4=g4.idgrupo4"
        sql = sql & " left join  GNVehiculo gnv on gnc.idvehiculo=gnv.idvehiculo"
        sql = sql & " left join  GNTransporte gnt on gnt.idtransporte=gnc.idTransporte"
        sql = sql & " WHERE IdVehiculo=" & key
        
    End If

    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.idCentro = !idCentro
        mProps.CodCentro = !CodCentro
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!Nombre) Then mProps.Nombre = !Nombre     '*** MAKOTO 14/feb/01 Agregado
        If Not IsNull(!FechaInicio) Then mProps.FechaInicio = !FechaInicio
        mProps.FechaFinal = !FechaFinal         'Null tambien es valido
        
        '*** MAKOTO 05/mar/01 Agregado
        If Not IsNull(!idProveedor) Then idProveedor = !idProveedor
        If Not IsNull(!IdCliente) Then IdCliente = !IdCliente
        If Not IsNull(!Observacion) Then mProps.Observacion = !Observacion
        If Not IsNull(!Valor1) Then mProps.Valor1 = !Valor1
        If Not IsNull(!Valor2) Then mProps.Valor2 = !Valor2
        If Not IsNull(!Valor3) Then mProps.Valor3 = !Valor3

'        For i = 1 To GNCGRUPO_MAX
'            If Not IsNull(.Fields("IdGrupo" & i)) Then mProps.IdGrupo(i) = .Fields("IdGrupo" & i)
'            If Not IsNull(.Fields("CodGrupo" & i)) Then mProps.CodGrupo(i) = .Fields("CodGrupo" & i)
'        Next i
        
        If Not IsNull(.Fields("IdVehiculo")) Then mProps.IdVehiculo = .Fields("IdVehiculo")
        If Not IsNull(.Fields("CodVehiculo")) Then mProps.CodVehiculo = .Fields("CodVehiculo")
        If Not IsNull(.Fields("fechaentrega")) Then mProps.FechaEntrega = .Fields("fechaentrega")
        If Not IsNull(.Fields("BandOcupado")) Then mProps.BandOcupado = .Fields("BandOcupado")
        If Not IsNull(.Fields("IdTransporte")) Then mProps.IdTransporte = .Fields("IdTransporte")
        If Not IsNull(.Fields("CodTransporte")) Then mProps.CodTransporte = .Fields("CodTransporte")
        
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Property Let Codhab(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.Codhab Then
        If Len(value) > 0 Then
            'Actualiza IdCliente también
            sql = "SELECT Idtipo FROM Tipohabitacion WHERE Codigo='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNCentroCosto.CodCliente", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdTipoHab = rs!IdTipo
            End If
            rs.Close
        Else
            mProps.IdTipoHab = 0
        End If
    End If

    Set rs = Nothing
    mProps.Codhab = value
    mbooModificado = True
End Property

Public Property Get Codhab() As String
    Codhab = Trim$(mProps.Codhab)
End Property


Public Property Let IdTipoHab(ByVal value As Long)
    mProps.IdTipoHab = value
    mbooModificado = True
End Property

Public Property Get IdTipoHab() As Long
    IdTipoHab = mProps.IdTipoHab
End Property


Public Property Get CodInventario() As String
    CodInventario = Trim$(mProps.CodInventario)
End Property

Public Property Let CodInventario(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodInventario Then
        If Len(value) > 0 Then
            'Actualiza IdVendedor también
            sql = "SELECT IdInventario FROM IVInventario WHERE CodInventario='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNCentroCosto.CodInventario", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdInventario = rs!IdInventario
            End If
            rs.Close
        Else
            mProps.IdInventario = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodInventario = value
    mbooModificado = True
End Property



Public Property Get IdInventario() As Long
    IdInventario = mProps.IdInventario
End Property

Public Property Let IdInventario(ByVal value As Long)
    mProps.IdInventario = value
    mbooModificado = True
End Property


Public Function verificaCentroCosto(ByVal CodCentro As String) As Boolean
Dim sql As String, rs As Recordset, i As Integer
    verificaCentroCosto = False
    sql = "SELECT * FROM GNCentroCosto WHERE CodCentro='" & CodCentro & "'"
    Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
    If rs.RecordCount > 0 Then
        verificaCentroCosto = True
    End If
End Function

Public Property Let IdGar(ByVal value As Long)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.IdGar Then
        'Actualiza el Codigo también
        mProps.CodGar = ""
        sql = "SELECT CodProvCli FROM pcProvCli WHERE IdProvCli=" & value
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then mProps.CodGar = rs!CodProvcli
        rs.Close
        Set rs = Nothing
    End If
    
    mProps.IdGar = value
    mbooModificado = True
End Property

'*** MAKOTO 05/mar/01 Agregado
Public Property Get IdGar() As Long
    IdGar = mProps.IdGar
End Property

Public Property Let CodGar(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodGar Then
        If Len(value) > 0 Then
            'Actualiza IdGar también
            sql = "SELECT IdProvcli FROM pcProvCli WHERE CodProvCli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNCentroCosto.CodGar", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdGar = rs!IdProvcli
            End If
            rs.Close
        Else
            mProps.IdGar = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodGar = value
    mbooModificado = True
End Property

Public Property Get CodGar() As String
    CodGar = Trim$(mProps.CodGar)
End Property

Public Property Let Referencia(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNCentroCosto_Referencia", MSGERR_INVALIDO
    End If
    
    mProps.Referencia = value
    mbooModificado = True
End Property

Public Property Get Referencia() As String
    Referencia = Trim$(mProps.Referencia)
End Property

Public Property Let IdGAR2(ByVal value As Long)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.IdGAR2 Then
        mProps.CodGAR2 = ""
        sql = "SELECT CodProvCli FROM pcProvCli WHERE IdProvCli=" & value
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then mProps.CodGAR2 = rs!CodProvcli
        rs.Close
        Set rs = Nothing
    End If
    
    mProps.IdGAR2 = value
    mbooModificado = True
End Property

'*** MAKOTO 05/mar/01 Agregado
Public Property Get IdGAR2() As Long
    IdGAR2 = mProps.IdGAR2
End Property

Public Property Let CodGAR2(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodGAR2 Then
        If Len(value) > 0 Then
            'Actualiza IdGAR2 también
            sql = "SELECT IdProvcli FROM pcProvCli WHERE CodProvCli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNCentroCosto.CodGAR2", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdGAR2 = rs!IdProvcli
            End If
            rs.Close
        Else
            mProps.IdGAR2 = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodGAR2 = value
    mbooModificado = True
End Property

Public Property Get CodGAR2() As String
    CodGAR2 = Trim$(mProps.CodGAR2)
End Property

Public Property Let CodCallePri(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodCallePri Then
        If Len(value) > 0 Then
            sql = "SELECT IdCalle FROM PCCalle WHERE CodCalle='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodCallePri", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IDCallePri = rs!IdCalle
            End If
            rs.Close
            Set rs = Nothing
        Else
            IDCallePri = 0
        End If
    End If
    mProps.CodCallePri = value
    mbooModificado = True
End Property

Public Property Get CodCallePri() As String
    CodCallePri = Trim$(mProps.CodCallePri)
End Property

Public Property Let IDCallePri(ByVal value As Long)
    mProps.IDCallePri = value
    mbooModificado = True
End Property

Public Property Get IDCallePri() As Long
    IDCallePri = mProps.IDCallePri
End Property

Public Property Let CodCalleSec(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodCalleSec Then
        If Len(value) > 0 Then
            
            sql = "SELECT IdCalle FROM PCCalle WHERE CodCalle='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodCalleSec", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCalleSec = rs!IdCalle
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdCalleSec = 0
        End If
    End If
    mProps.CodCalleSec = value
    mbooModificado = True
End Property

Public Property Get CodCalleSec() As String
    CodCalleSec = Trim$(mProps.CodCalleSec)
End Property

Public Property Let IdCalleSec(ByVal value As Long)
    mProps.IdCalleSec = value
    mbooModificado = True
End Property

Public Property Get IdCalleSec() As Long
    IdCalleSec = mProps.IdCalleSec
End Property

Public Property Get NumCasa() As String
    NumCasa = mProps.NumCasa
End Property

Public Property Let NumCasa(ByVal value As String)
    mProps.NumCasa = value
    mbooModificado = True
End Property

Public Property Get posGoogleMaps() As String
    posGoogleMaps = mProps.posGoogleMaps
End Property

Public Property Let posGoogleMaps(ByVal value As String)
    mProps.posGoogleMaps = value
    mbooModificado = True
End Property

Public Property Let Referencia2(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNCentroCosto_Referencia2", MSGERR_INVALIDO
    End If
    
    mProps.Referencia2 = value
    mbooModificado = True
End Property

Public Property Get Referencia2() As String
    Referencia2 = Trim$(mProps.Referencia2)
End Property

Public Property Let IdPais(ByVal value As Long)
    mProps.IdPais = value
    mbooModificado = True
End Property

Public Property Get IdPais() As Long
    IdPais = mProps.IdPais
End Property


Public Property Let IdProvincia(ByVal value As Long)
    mProps.IdProvincia = value
    mbooModificado = True
End Property

Public Property Get IdProvincia() As Long
    IdProvincia = mProps.IdProvincia
End Property

Public Property Let IdCanton(ByVal value As Long)
    mProps.IdCanton = value
    mbooModificado = True
End Property

Public Property Get IdCanton() As Long
    IdCanton = mProps.IdCanton
End Property

Public Property Let IDParroquia(ByVal value As Long)
    mProps.IDParroquia = value
    mbooModificado = True
End Property

Public Property Get IDParroquia() As Long
    IDParroquia = mProps.IDParroquia
End Property

Public Property Let CodPais(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodProvincia Then
        If Len(value) > 0 Then
            sql = "SELECT IdPais FROM PCPais WHERE CodPais='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodPais", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdPais = rs!IdPais
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdPais = 0
        End If
    End If
    mProps.CodPais = value
    mbooModificado = True
End Property

Public Property Get CodPais() As String
    CodPais = Trim$(mProps.CodPais)
End Property

Public Property Let CodProvincia(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodProvincia Then
        If Len(value) > 0 Then
            'Actualiza IdGarante también
            sql = "SELECT IdProvincia FROM PCProvincia WHERE CodProvincia='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNCentroCosto_CodProvincia", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdProvincia = rs!IdProvincia
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdProvincia = 0
        End If
    End If
    mProps.CodProvincia = value
    mbooModificado = True
End Property

Public Property Get CodProvincia() As String
    CodProvincia = Trim$(mProps.CodProvincia)
End Property

Public Property Let CodCanton(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodCanton Then
        If Len(value) > 0 Then
            'Actualiza IdGarante también
            sql = "SELECT Idcanton FROM PCcanton WHERE Codcanton='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNCentroCosto_Codcanton", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdCanton = rs!IdCanton
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdCanton = 0
        End If
    End If
    mProps.CodCanton = value
    mbooModificado = True
End Property

Public Property Get CodCanton() As String
    CodCanton = Trim$(mProps.CodCanton)
End Property

Public Property Let CodParroquia(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
'    If value <> mProps.CodParroqui Then
        If Len(value) > 0 Then
            'Actualiza IdGarante también
            sql = "SELECT IdParroquia FROM PCParroquia pcp inner join pccanton pcc on pcp.idcanton = pcc.idcanton WHERE codcanton = '" & mProps.CodCanton & "' and CodParroquia='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNCentroCosto_CodParroquia", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IDParroquia = rs!IDParroquia
            End If
            rs.Close
            Set rs = Nothing
        Else
            IDParroquia = 0
        End If
'    End If
    mProps.CodParroqui = value
    mbooModificado = True
End Property

Public Property Get CodParroquia() As String
    CodParroquia = Trim$(mProps.CodParroqui)
End Property

Public Property Let Telefono2(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNCentroCosto_telefono2", MSGERR_INVALIDO
    End If
    
    mProps.Telefono2 = value
    mbooModificado = True
End Property

Public Property Get Telefono2() As String
    Telefono2 = Trim$(mProps.Telefono2)
End Property

Public Property Let Propietario(ByVal value As String)
    If Len(value) > 200 Then
        Err.Raise ERR_INVALIDO, "GNCentroCosto_Propietario", MSGERR_INVALIDO
    End If
    
    mProps.Propietario = value
    mbooModificado = True
End Property

Public Property Get Propietario() As String
    Propietario = Trim$(mProps.Propietario)
End Property

Public Property Let BandCasaPropia(ByVal value As String)
    'If Len(value) > 1 Then
        mProps.BandCasaPropia = value
        mbooModificado = True
    'End If
End Property

Public Property Get BandCasaPropia() As String
    BandCasaPropia = mProps.BandCasaPropia
End Property

Public Property Let IdZona(ByVal value As Long)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.IdZona Then
        'Actualiza el Codigo también
        mProps.CodZona = ""
        sql = "SELECT CodZona FROM gnZona WHERE IdZona=" & value
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then mProps.CodZona = rs!CodZona
        rs.Close
        Set rs = Nothing
    End If
    
    mProps.IdZona = value
    mbooModificado = True
End Property

'*** MAKOTO 05/mar/01 Agregado
Public Property Get IdZona() As Long
    IdZona = mProps.IdZona
End Property

Public Property Let CodZona(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodZona Then
        If Len(value) > 0 Then
            'Actualiza IdZona también
            sql = "SELECT IdZona FROM gnZona WHERE Codzona='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNCentroCosto.CodZona", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdZona = rs!IdZona
            End If
            rs.Close
        Else
            mProps.IdZona = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodZona = value
    mbooModificado = True
End Property

Public Property Get CodZona() As String
    CodZona = Trim$(mProps.CodZona)
End Property

Public Property Let IdSubZona(ByVal value As Long)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.IdSubZona Then
        'Actualiza el Codigo también
        mProps.CodSubZona = ""
        sql = "SELECT CodCompetencia FROM gnCompetencia WHERE IdCompetencia=" & value
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then mProps.CodSubZona = rs!CodCompetencia
        rs.Close
        Set rs = Nothing
    End If
    
    mProps.IdSubZona = value
    mbooModificado = True
End Property

'*** MAKOTO 05/mar/01 Agregado
Public Property Get IdSubZona() As Long
    IdSubZona = mProps.IdSubZona
End Property

Public Property Let CodSubZona(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodSubZona Then
        If Len(value) > 0 Then
            'Actualiza IdSubZona también
            'sql = "SELECT IdSubZona FROM gnSubZona WHERE CodSubZona='" & value & "'"
            sql = "SELECT Idcompetencia FROM gncompetencia WHERE Codcompetencia='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNCentroCosto.CodSubZona", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdSubZona = rs!IdCompetencia
            End If
            rs.Close
        Else
            mProps.IdSubZona = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodSubZona = value
    mbooModificado = True
End Property

Public Property Get CodSubZona() As String
    CodSubZona = Trim$(mProps.CodSubZona)
End Property


Public Property Let IdVendedor(ByVal value As Long)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.IdVendedor Then
        'Actualiza el Codigo también
        mProps.CodVendedor = ""
        sql = "SELECT CodVendedor FROM fcVendedor WHERE IdVendedor=" & value
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then mProps.CodVendedor = rs!CodVendedor
        rs.Close
        Set rs = Nothing
    End If
    
    mProps.IdVendedor = value
    mbooModificado = True
End Property

'*** MAKOTO 05/mar/01 Agregado
Public Property Get IdVendedor() As Long
    IdVendedor = mProps.IdVendedor
End Property

Public Property Let CodVendedor(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodVendedor Then
        If Len(value) > 0 Then
            'Actualiza IdVendedor también
            sql = "SELECT IdVendedor FROM fcVendedor WHERE CodVendedor='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNCentroCosto.CodVendedor", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdVendedor = rs!IdVendedor
            End If
            rs.Close
        Else
            mProps.IdVendedor = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodVendedor = value
    mbooModificado = True
End Property

Public Property Get CodVendedor() As String
    CodVendedor = Trim$(mProps.CodVendedor)
End Property

Public Sub GrabarHistorial(ByVal bandsusp As Boolean, fechaSuspend As Date, motivo As Integer, detalle As String, USUARIO As String)
    Dim sql  As String


    sql = "INSERT INTO GNHistorialContratos(IDCentro,Motivo,Detalle,BandSuspendido,Usuario,fecha) " & _
          " VALUES (" & mProps.idCentro & "," & motivo & ",'" & _
                         UCase(detalle) & "'," & _
                         IIf(bandsusp, 1, 0) & _
                          ",'" & USUARIO & "','" & fechaSuspend & "')"
    
    mProps.objEmpresa.Execute sql, True

End Sub

Public Function RecuperaHistorial() As Variant
    Dim sql As String, rs As Recordset
    
    sql = "select FechaGrabado ,"
    sql = sql & " case Motivo when 0 then 'COSTO' WHEN 1 THEN 'CALIDAD' WHEN 3 THEN 'ATENCION' ELSE 'OTROS' END AS MOTIVO, "
    sql = sql & " BandSuspendido,Usuario, fecha "
    sql = sql & " From GNHistorialContratos"
    sql = sql & " Where IDCentro = " & mProps.idCentro
    sql = sql & " order by FechaGrabado"
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)

    Set RecuperaHistorial = rs
End Function


Public Property Let NumTV(ByVal value As Integer)
    mProps.NumTV = value
    mbooModificado = True
End Property

Public Property Get NumTV() As Integer
    NumTV = mProps.NumTV
End Property

Public Property Let Referencia3(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNCentroCosto_Referencia3", MSGERR_INVALIDO
    End If
    
    mProps.Referencia3 = value
    mbooModificado = True
End Property

Public Property Get Referencia3() As String
    Referencia3 = Trim$(mProps.Referencia3)
End Property

Public Function verificaReferencia2(ByVal ref2 As String, ByRef cod As String) As Boolean
Dim sql As String, rs As Recordset, i As Integer
    verificaReferencia2 = False
    sql = "SELECT codcentro FROM GNCentroCosto WHERE referencia2='" & ref2 & "'"
    sql = sql & "and codcentro <> '" & Me.CodCentro & "'"
    Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
    If rs.RecordCount > 0 Then
        cod = rs.Fields("codcentro")
        verificaReferencia2 = True
    End If
End Function

Public Property Let BandCorteMora(ByVal value As Boolean)
    mProps.BandCorteMora = value
    mbooModificado = True
End Property

Public Property Get BandCorteMora() As Boolean
    BandCorteMora = mProps.BandCorteMora
End Property

