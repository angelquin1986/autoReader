VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GNContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Private Type T_PROPS
    IdContrato As Long
    CodContrato As String
    IdProvcli As Long
    CodProvcli As String
    Descripcion As String
    IdGarantia As Long
    CodGarantia As String
    Observacion As String
    FechaInicio As Date
    FechaFinal As Date
    BandValida As Boolean
    
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
    
    
    
    fechagrabado As Date
    objEmpresa As Empresa       'Referencia al objeto padre
    
    
End Type

Private mcolDetalleContrato As Collection      'Coleccion de detalles de familia

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean


Public Sub Grabar()
    Dim sql As String, rs As Recordset, i As Integer
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
    If mbooEsNuevo Then
        sql = "SELECT CodContrato FROM GNContrato " & _
              "WHERE (CodContrato='" & mProps.CodContrato & "') AND " & _
              "(IdContrato<>" & mProps.IdContrato & ")"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then      'Si ya existe el codigo
            Err.Raise ERR_REPITECODIGO, "GNContrato", MSGERR_REPITECODIGO
        End If
        rs.Close
    End If
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM GNContrato WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM GNContrato WHERE IdContrato=" & mProps.IdContrato
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "GNContrato_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en GNContrato
    With rs
        !CodContrato = mProps.CodContrato
        !IdProvcli = mProps.IdProvcli
        !Descripcion = mProps.Descripcion
        !IdGarantia = mProps.IdGarantia
        !FechaInicio = mProps.FechaInicio
        !FechaFinal = mProps.FechaFinal
        !Observacion = mProps.Observacion
        !BandValida = mProps.BandValida
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
        
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IdContrato = !IdContrato
        GrabaContratoDetalle
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_PR", "Contrato Nuevo: " & _
                                                         mProps.CodContrato, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_PR", "Contrato Modificado: " & _
                                                         mProps.CodContrato, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_GNCONTRATO
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset, i As Integer
    
    If VarType(key) = vbString Then
        sql = "SELECT GC.*, PC.CODPROVCLI AS  CodProvCli , PCG.CODPROVCLI AS  CodGar,  PCG2.CODPROVCLI AS  CodGar2,"
        sql = sql & " pccallePri.codCalle as CodCallePri, pccalleSec.codCalle as CodCalleSec, "
        sql = sql & " pcp.codpais, pcpro.codProvincia, pccan.codcanton, pcparr.codparroquia, gnz.Codzona "
        sql = sql & " from gncontrato GC "
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
        
        
        
'''        sql = sql & " IdContrato, CodContrato, gnp.Descripcion , FechaInicio, gnp.Observacion, "
'''        sql = sql & " FechaFinal, gnp.FechaGrabado, Propietario, IdProvCli,"
'''        sql = sql & " gnc.IdGarantia, Nombre1,Nombre2,Nombre3, Cargo1,Cargo2,Cargo3,Telefono1,Telefono2,Telefono3, Obser1,Obser2, Obser3, "
'''        sql = sql & " gnc.idCanton, CodCanton, gnp.idinventario, ivi.CodInventario, Cantidad,FechaProyeccion, Estado, "
'''        sql = sql & " EstadoNegociacion, gnp.Observacion, gnp.num1, gnp.num2, gnp.Observacion2, gnp.idcompetencia, gncmp.CodCompetencia "
'''        sql = sql & " FROM GNContrato gnp left join ivinventario ivi on gnp.idinventario= ivi.idinventario "
'''        sql = sql & " left join  PCCanton gnc on gnc.idCanton=gnp.idCanton"
'''        sql = sql & " left join  gnCompetencia gncmp on gnp.idCompetencia=gncmp.idCompetencia "
        sql = sql & " WHERE CodContrato='" & key & "'"
    Else
        sql = "SELECT * from gncontrato"
'        sql = sql & " IdContrato, CodContrato, gnp.Descripcion, FechaInicio,"
'        sql = sql & " FechaFinal, gnp.FechaGrabado, Propietario, IdProvCli,"
'        sql = sql & " IdGarantia, Observacion, gnc.IdInventario, Cantidad, estado "
'        sql = sql & " FROM GNContrato gnc"
        sql = sql & " WHERE IdContrato=" & key
        
    End If

    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario

    With rs

        mProps.IdContrato = !IdContrato
        mProps.CodContrato = !CodContrato
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        mProps.FechaInicio = !FechaInicio
        mProps.FechaFinal = !FechaFinal         'Null tambien es valido
        
        '*** MAKOTO 05/mar/01 Agregado
        If Not IsNull(!IdProvcli) Then IdProvcli = !IdProvcli
        If Not IsNull(!CodProvcli) Then CodProvcli = !CodProvcli
        
        If Not IsNull(!IdGar) Then IdGar = !IdGar
        If Not IsNull(!CodGar) Then CodGar = !CodGar
        If Not IsNull(!IdGarantia) Then IdGarantia = !IdGarantia
        If Not IsNull(!Observacion) Then mProps.Observacion = !Observacion

        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
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
                
                
                
       CargaDetalleContrato
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    Dim gcc As GNCentroCosto
    
        'Verifica si existe un comprobante que usa el codigo de Contrato de costo
    sql = "SELECT TransID FROM GNComprobante g inner join gncentrocosto gcc on g.idcentro= gcc.idcentro "
    sql = sql & " WHERE gcc.codcentro='" & mProps.CodContrato & "'"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        rs.Close
        Set rs = Nothing
        Err.Raise ERR_NOELIMINABLE, "GNContrato.Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close
    Set rs = Nothing
    
    'Elimina el Contrato de costo
    sql = "DELETE FROM GNContrato WHERE IdContrato=" & mProps.IdContrato
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_CC", "Contrato Eliminado: " & _
                                                         mProps.CodContrato, "DT"

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_GNCONTRATO
End Sub

Friend Function Clone() As GNContrato
    Dim nuevo As GNContrato
    
    Set nuevo = New GNContrato
    'Copia las propiedades al nuevo objeto
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdContrato = 0
    mProps.BandValida = True
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property


Public Property Let CodContrato(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNContrato_CodContrato", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodContrato = value
    mbooModificado = True
End Property

Public Property Get CodContrato() As String
    CodContrato = Trim$(mProps.CodContrato)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "GNContrato_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

'*** MAKOTO 05/mar/01 Agregado
Public Property Let IdProvcli(ByVal value As Long)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.IdProvcli Then
        'Actualiza el Codigo también
        mProps.CodProvcli = ""
        sql = "SELECT CodProvCli FROM pcProvCli WHERE IdProvCli=" & value
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then mProps.CodProvcli = rs!CodProvcli
        rs.Close
        Set rs = Nothing
    End If
    
    mProps.IdProvcli = value
    mbooModificado = True
End Property

'*** MAKOTO 05/mar/01 Agregado
Public Property Get IdProvcli() As Long
    IdProvcli = mProps.IdProvcli
End Property

Public Property Let CodProvcli(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodProvcli Then
        If Len(value) > 0 Then
            'Actualiza IdProvCli también
            sql = "SELECT IdProvCli FROM pcProvCli WHERE CodProvCli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNContrato.CodProvCli", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdProvcli = rs!IdProvcli
            End If
            rs.Close
        Else
            mProps.IdProvcli = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodProvcli = value
    mbooModificado = True
End Property

Public Property Get CodProvcli() As String
    CodProvcli = Trim$(mProps.CodProvcli)
End Property

'*** MAKOTO 05/mar/01 Agregado
Public Property Let IdGarantia(ByVal value As Long)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.IdGarantia Then
        'Actualiza el Codigo también
        mProps.CodGarantia = ""
        sql = "SELECT Codcompetencia FROM gncompetencia WHERE Idcompetencia=" & value
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then mProps.CodGarantia = rs!CodCompetencia
        rs.Close
        Set rs = Nothing
    End If
    
    mProps.IdGarantia = value
    mbooModificado = True
End Property

'*** MAKOTO 05/mar/01 Agregado
Public Property Get IdGarantia() As Long
    IdGarantia = mProps.IdGarantia
End Property

Public Property Let CodGarantia(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodGarantia Then
        If Len(value) > 0 Then
            'Actualiza IdGarantia también
            sql = "SELECT Idcompetencia FROM gncompetencia WHERE Codcompetencia='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNContrato.CodGarantia", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdGarantia = rs!IdCompetencia
            End If
            rs.Close
        Else
            mProps.IdGarantia = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodGarantia = value
    mbooModificado = True
End Property

Public Property Get CodGarantia() As String
    CodGarantia = Trim$(mProps.CodGarantia)
End Property

Public Property Let FechaInicio(ByVal value As Date)
    mProps.FechaInicio = value
    mbooModificado = True
End Property

Public Property Get FechaInicio() As Date
    FechaInicio = mProps.FechaInicio
End Property

Public Property Let FechaFinal(ByVal value As Date)
    mProps.FechaFinal = value
    mbooModificado = True
End Property

Public Property Get FechaFinal() As Date
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
        .FechaFinal = DateAdd("yyyy", 1, Date)
        .BandValida = True
    End With
    
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub


Public Property Let Observacion(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "GNContrato_Observacion", MSGERR_INVALIDO
    End If
    
    mProps.Observacion = value
    mbooModificado = True
End Property

Public Property Get Observacion() As String
    Observacion = Trim$(mProps.Observacion)
End Property

Public Sub GrabaContratoDetalle()
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As GnContratoDetalle, sql As String, band As Boolean
    Dim NumHijos As Integer, CountNumHijos As Integer
    On Error GoTo CapturaError
    
    'Elimina  todos  los registros
    sql = "Delete gnc " & _
          "From gnVehiculo Iv " & _
          "INNER JOIN GNContratoDetalle gnc ON IV.Idvehiculo = gnc.Idvehiculo "
        sql = sql & " Where gnc.IdContrato = " & mProps.IdContrato

    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    band = True
    If Not mcolDetalleContrato Is Nothing Then
        sql = "SELECT * FROM GNContratoDetalle WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        NumHijos = mcolDetalleContrato.Count
        CountNumHijos = 0
        For i = 1 To mcolDetalleContrato.Count
            If CountNumHijos < NumHijos Then
                CountNumHijos = CountNumHijos + 1
                Set obj = mcolDetalleContrato.item(i)
                If obj.IdVehiculo = 0 Then
                        obj.Recuperar obj.CodVehiculo
                        If verificacionDuiplicado(obj.IdVehiculo) Then
                            RemoveDetalleContrato i
                            i = i - 1
                        End If
                End If
                If band Then
                    rs1.AddNew
                    rs1!IdVehiculo = obj.IdVehiculo
                    rs1!IdContrato = Me.IdContrato
                    rs1!Orden = obj.Orden
                    rs1!Cantidad = obj.Cantidad
                    rs1.Update
                End If
            End If
            band = True

        Next i
    End If
    Set rs1 = Nothing
    Set obj = Nothing
    Exit Sub
CapturaError:
    MsgBox Err.Description
    Set rs1 = Nothing
    Set obj = Nothing
    Exit Sub
End Sub


Public Function RecuperaDetalleContrato(ix As Long) As GnContratoDetalle
    Set RecuperaDetalleContrato = mcolDetalleContrato.item(ix)
End Function

Public Function AddDetalleContrato() As Long
    Dim nuevo As GnContratoDetalle
    
    If mcolDetalleContrato Is Nothing Then CargaDetalleContrato
    Set nuevo = New GnContratoDetalle
    Set nuevo.GNContrato = Me  'Me  olvide  del padre
    'Proporciona valores predeterminados
    nuevo.Cantidad = 1
    
    mcolDetalleContrato.Add nuevo
    AddDetalleContrato = mcolDetalleContrato.Count    'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    
    mbooModificado = True
    
End Function

Public Sub RemoveDetalleContrato(ByVal ix As Long, Optional obj As GnContratoDetalle)
    If mcolDetalleContrato Is Nothing Then CargaDetalleContrato
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolDetalleContrato.Count
            If mcolDetalleContrato.item(ix) Is obj Then
                mcolDetalleContrato.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolDetalleContrato.Remove ix
    End If
    
    'mcolDetalleContrato.Remove ix
    mbooModificado = True
End Sub

Private Sub CargaDetalleContrato()
    Dim rs As Recordset, sql As String, obj As GnContratoDetalle
    Dim i As Long
    Set mcolDetalleContrato = New Collection
        sql = "Select gnV.Idvehiculo, gnv.Codvehiculo, gnv.Descripcion, gncontr.cantidad, gncontr.Orden " & _
              " From  GnVehiculo gnV " & _
              "INNER JOIN GNContratoDetalle gncontr  "
          sql = sql & " ON gnv.Idvehiculo = gncontr.Idvehiculo "
            sql = sql & " Where gncontr.IdContrato = " & mProps.IdContrato
            sql = sql & " order by gncontr.Orden "

    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            i = AddDetalleContrato
            Set obj = mcolDetalleContrato.item(i)
                  obj.IdVehiculo = !IdVehiculo
                  obj.CodVehiculo = !CodVehiculo
                  obj.Descripcion = !Descripcion
                 obj.Cantidad = !Cantidad
                 obj.Orden = !Orden
            
           .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    Set obj = Nothing
    'mbooEsNuevo = False
    'mbooModificado = False
End Sub


Private Function verificacionDuiplicado(ByVal IdInventario As String) As Boolean
    Dim rs As Recordset, sql As String
    verificacionDuiplicado = False
    'Recupera los detalles de existencia
    sql = "SELECT IdInventario FROM GNContratoDetalle where idinventario =" & IdInventario
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If rs.RecordCount > 0 Then
        verificacionDuiplicado = True
    End If
    Set rs = Nothing
End Function

Friend Property Get IdContrato() As Long
    IdContrato = mProps.IdContrato
End Property

Public Property Get NumContratoDetalle() As Long
    If mcolDetalleContrato Is Nothing Then
        NumContratoDetalle = 0
    Else
        NumContratoDetalle = mcolDetalleContrato.Count
    End If
    
End Property

Public Property Let BandValida(ByVal value As Boolean)
    mProps.BandValida = value
    mbooModificado = True
End Property

Public Property Get BandValida() As Boolean
    BandValida = mProps.BandValida
End Property

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
                Err.Raise ERR_NOHAYCODIGO, "GNContrato.CodGar", MSGERR_NOHAYCODIGO
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
        Err.Raise ERR_INVALIDO, "GNContrato_Referencia", MSGERR_INVALIDO
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
                Err.Raise ERR_NOHAYCODIGO, "GNContrato.CodGAR2", MSGERR_NOHAYCODIGO
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
        Err.Raise ERR_INVALIDO, "GNContrato_Referencia2", MSGERR_INVALIDO
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
                Err.Raise ERR_NOHAYCODIGO, "GNContrato_CodProvincia", MSGERR_NOHAYCODIGO
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
                Err.Raise ERR_NOHAYCODIGO, "GNContrato_Codcanton", MSGERR_NOHAYCODIGO
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
                Err.Raise ERR_NOHAYCODIGO, "GNContrato_CodParroquia", MSGERR_NOHAYCODIGO
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
        Err.Raise ERR_INVALIDO, "GNContrato_telefono2", MSGERR_INVALIDO
    End If
    
    mProps.Telefono2 = value
    mbooModificado = True
End Property

Public Property Get Telefono2() As String
    Telefono2 = Trim$(mProps.Telefono2)
End Property

Public Property Let Propietario(ByVal value As String)
    If Len(value) > 200 Then
        Err.Raise ERR_INVALIDO, "GNContrato_Propietario", MSGERR_INVALIDO
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
                Err.Raise ERR_NOHAYCODIGO, "GNContrato.CodZona", MSGERR_NOHAYCODIGO
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

