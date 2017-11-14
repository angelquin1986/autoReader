VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GNReservaHotel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    idReserva As Long
    idCentro As Long
    transid As Long
    CodCentro As String
    Huesped As String
    FechaReserva As Date
    FechaInicio As Date
    FechaFinal As Date       'Variant para almacenar Null
    IdCliente As Long
    CodCliente As String
    Precio As Currency
    NumOcupantes As Integer
    idEstado As Long
    FechaChecKOut As Date
    BandValida As Boolean
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean


Public Sub Grabar(transid)
    Dim sql As String, rs As Recordset, i As Integer
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
'    'Verifica si no repite el código
'    sql = "SELECT CodCentro FROM GNCentroCosto " & _
'          "WHERE (CodCentro='" & mProps.CodCentro & "') AND " & _
'          "(IdCentro<>" & mProps.idCentro & ")"
'    Set rs = mProps.objEmpresa.OpenRecordset(sql)
'    If Not rs.EOF Then      'Si ya existe el codigo
'        Err.Raise ERR_REPITECODIGO, "GNCentroCosto", MSGERR_REPITECODIGO
'    End If
'    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM GNReserva WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM GNReserva WHERE IdReserva=" & mProps.idReserva
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
'        If rs!fechagrabado <> mProps.fechagrabado Then
'            Err.Raise ERR_MODIFICADO, "GNCentroCosto_Grabar", MSGERR_MODIFICADO
'        End If
    End If

    'Graba el registro en GNCentroCosto
    With rs
        !idCentro = mProps.idCentro
        !IdCliente = mProps.IdCliente
        !transid = transid
        !fecha = mProps.FechaReserva
        !FechaInicio = mProps.FechaInicio
        !FechaFinal = mProps.FechaFinal
        !Precio = mProps.Precio
        !NumOcupantes = mProps.NumOcupantes
        !Huesped = mProps.Huesped
        !idEstado = mProps.idEstado
        !FechaChecKOut = mProps.FechaChecKOut
        !BandOcupado = mProps.BandValida
'        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        'mProps.fechagrabado = !fechagrabado
        mProps.idReserva = !idReserva
        .Close
    End With
    If mbooEsNuevo Then
        GrabarFecha
    End If
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_RESERVA", "Reserva Hab: " & _
                                                         mProps.CodCentro, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_RESERVA", "Reserva Hab Modificado: " & _
                                                         mProps.CodCentro, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    'mProps.objEmpresa.RaiseRefreshEvent REFRESH_CENTROCOSTO
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset, i As Integer
    
        sql = "SELECT gr.* "
        sql = sql & " FROM GNReserva gr left join  GNCentroCosto gnc on gnc.idCentro=gr.idcentro "
        sql = sql & " left join  pcprovcli pc on pc.idProvcli=gr.idcliente"
        sql = sql & " WHERE transid=" & key

    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.idReserva = !idReserva
        mProps.idCentro = !idCentro
'        mProps.CodCentro = !CodCentro
        If Not IsNull(!fecha) Then mProps.FechaReserva = !fecha
        If Not IsNull(!FechaInicio) Then mProps.FechaInicio = !FechaInicio
        If Not IsNull(!FechaFinal) Then mProps.FechaFinal = !FechaFinal
        If Not IsNull(!IdCliente) Then IdCliente = !IdCliente
        If Not IsNull(!Precio) Then mProps.Precio = !Precio
        If Not IsNull(!NumOcupantes) Then mProps.NumOcupantes = !NumOcupantes
        If Not IsNull(!Huesped) Then mProps.Huesped = !Huesped
        If Not IsNull(!transid) Then mProps.transid = !transid
        If Not IsNull(!idEstado) Then mProps.idEstado = !idEstado
        If Not IsNull(!FechaChecKOut) Then mProps.FechaChecKOut = !FechaChecKOut
        If Not IsNull(!BandOcupado) Then mProps.BandValida = !BandOcupado
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


Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.idReserva = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property


Public Property Get CodCentro() As String
    CodCentro = Trim$(mProps.CodCentro)
End Property

Public Property Let CodCentro(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCentro Then
        If Len(value) > 0 Then
            'Actualiza IdProveedor también
            sql = "SELECT IdCentro FROM GNCentro WHERE CodCentro='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNReserva.CodCentro", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.idCentro = rs!idCentro
            End If
            rs.Close
        Else
            mProps.idCentro = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodCentro = value
    mbooModificado = True
End Property
Public Property Let idCentro(ByVal value As Long)
    Dim sql As String, rs As Recordset
    mProps.idCentro = value
    mbooModificado = True
End Property


Public Property Get idCentro() As Long
    idCentro = mProps.idCentro
End Property

Public Property Let IdCliente(ByVal value As Long)
    Dim sql As String, rs As Recordset
    mProps.IdCliente = value
    mbooModificado = True
End Property


Public Property Get IdCliente() As Long
    IdCliente = mProps.IdCliente
End Property

'Public Property Let CodCliente(ByVal value As String)
'    Dim sql As String, rs As Recordset
'
'    'Cuando cambia el valor
'    If value <> mProps.CodCliente Then
'        If Len(value) > 0 Then
'            'Actualiza IdCliente también
'            sql = "SELECT IdProvCli FROM PCProvCli WHERE CodProvCli='" & value & "'"
'            Set rs = mProps.objEmpresa.OpenRecordset(sql)
'            If rs.EOF Then
'                Err.Raise ERR_NOHAYCODIGO, "GNCentroCosto.CodCliente", MSGERR_NOHAYCODIGO
'                Exit Property
'            Else
'                mProps.IdCliente = rs!IdProvCli
'            End If
'            rs.Close
'        Else
'            mProps.IdCliente = 0
'        End If
'    End If
'
'    Set rs = Nothing
'    mProps.CodCliente = value
'    mbooModificado = True
'End Property
'
'Public Property Get CodCliente() As String
'    CodCliente = Trim$(mProps.CodCliente)
'End Property

Public Property Let FechaReserva(ByVal value As Date)
    mProps.FechaReserva = value
    mbooModificado = True
End Property

Public Property Get FechaReserva() As Date
    FechaReserva = mProps.FechaReserva
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

Public Property Get EsNuevo() As Boolean
    EsNuevo = mbooEsNuevo
End Property

Public Property Get Modificado() As Boolean
    Modificado = mbooModificado
End Property


Private Sub Class_Initialize()

    'Valores predeterminados
    With mProps
'        .FechaInicio = Date
 '       .FechaFinal = Null
  '      .FechaEntrega = Date
    End With
    
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Property Let Precio(ByVal value As Currency)
    mProps.Precio = value
    mbooModificado = True
End Property

Public Property Get Precio() As Currency
    Precio = mProps.Precio
End Property

Public Property Let NumOcupantes(ByVal value As Integer)
    mProps.NumOcupantes = value
    mbooModificado = True
End Property

Public Property Get NumOcupantes() As Integer
    NumOcupantes = mProps.NumOcupantes
End Property

Public Property Get CodCliente() As String
    CodCliente = Trim$(mProps.CodCliente)
End Property

Public Property Let CodCliente(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodCliente Then
        If Len(value) > 0 Then
            'Actualiza IdProveedor también
            sql = "SELECT IdProvCli FROM PCProvCli WHERE CodProvcli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNReserva.CodCliente", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdCliente = rs!IdProvCli
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

Private Sub GrabarFecha()
Dim sql As String
Dim i As Long
Dim diaI As Integer
Dim diaF As Integer
Dim diaFin As Integer
Dim mes As Integer
Dim mesh  As Integer
Dim año As Integer
Dim rs As Recordset

'---------------------------
                        ' mes = Str(DatePart("m", dtpFechaDesde.value))
                        'mesh = Str(DatePart("m", dtpFechaHasta.value))
                        'diaInicial = Str(DatePart("d", dtpFechaDesde.value))
                        'diaFinal = Str(DatePart("d", dtpFechaHasta.value))
                        'año = DatePart("yyyy", dtpFechaDesde.value)
                        diaI = DatePart("d", mProps.FechaInicio)
                        diaF = DatePart("d", mProps.FechaFinal)
                        mes = DatePart("m", mProps.FechaInicio)
                        mesh = DatePart("m", mProps.FechaFinal)
                        año = DatePart("yyyy", mProps.FechaInicio)
                        
                        If mes <> mesh Then
                            diaFin = DatePart("d", DateAdd("d", -1, CDate("01/" & mesh & "/" & DatePart("yyyy", mProps.FechaInicio))))
                        Else
                            diaFin = DatePart("d", mProps.FechaFinal)
                        End If
                        '------
'                        For i = diaInicial To diaFin
'                            ix = ix + 1
'                            AuxDesct = gobjMain.EmpresaActual.ObtieneDesctoPCxIVxDia(iv.CodGrupo(ivgrupo), fcbCli.KeyText, obj.GNTrans.CodFormaPre, codDesc, CDate(i & "/ " & mes & "/ " & año), obj.HoraTrans)
'                            Precio = Precio + iv.Precio(NumPrecio) - ((AuxDesct / 100) * iv.Precio(NumPrecio))
'                            grd.AddItem ix & vbTab & CDate(i & "/ " & mes & "/ " & año) & vbTab & iv.Precio(NumPrecio) & vbTab & AuxDesct & vbTab & iv.Precio(NumPrecio) - ((AuxDesct / 100) * iv.Precio(NumPrecio))
'                        Next
'                        If mes <> mesh Then
'                            For i = 1 To diaFinal
'                                AuxDesct = gobjMain.EmpresaActual.ObtieneDesctoPCxIVxDia(iv.CodGrupo(ivgrupo), fcbCli.KeyText, obj.GNTrans.CodFormaPre, codDesc, CDate(i & "/ " & mesh & "/ " & año), obj.HoraTrans)
'                                Precio = Precio + iv.Precio(NumPrecio) - ((AuxDesct / 100) * iv.Precio(NumPrecio))
'                                grd.AddItem ix & vbTab & CDate(i & "/ " & mesh & "/ " & año) & vbTab & iv.Precio(NumPrecio) & vbTab & AuxDesct & vbTab & iv.Precio(NumPrecio) - ((AuxDesct / 100) * iv.Precio(NumPrecio))
'                            Next
'                        End If
'--------------------------
    sql = "SELECT * FROM GNReservaFechas WHERE 1=0"
    Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
    
    With rs
        For i = diaI To diaFin
            .AddNew
            !idReserva = mProps.idReserva
            !fecha = CDate(i & "/" & mes & "/" & año)
            .Update
        Next
        If mes <> mesh Then
            For i = 1 To diaF
                .AddNew
                !idReserva = mProps.idReserva
                !fecha = CDate(i & "/" & mesh & "/" & año)
                .Update
            Next
        End If
                
    #If DAOLIB Then
        rs.Bookmark = rs.LastModified
    #End If
    .Move 0             'Para actualizar
         
        .Close
    End With
    Set rs = Nothing
End Sub

Public Property Let Huesped(ByVal value As String)
    mProps.Huesped = value
    mbooModificado = True
End Property

Public Property Get Huesped() As String
    Huesped = mProps.Huesped
End Property

Public Property Let transid(ByVal value As Long)
    mProps.transid = value
    mbooModificado = True
End Property

Public Property Get transid() As Long
    transid = mProps.transid
End Property
Public Property Let idEstado(ByVal value As Long)
    mProps.idEstado = value
    mbooModificado = True
End Property

Public Property Get idEstado() As Long
    idEstado = mProps.idEstado
End Property

Public Property Let FechaChecKOut(ByVal value As Date)
    mProps.FechaChecKOut = value
    mbooModificado = True
End Property

Public Property Get FechaChecKOut() As Date
    FechaChecKOut = mProps.FechaChecKOut
End Property

Public Property Let BandValida(ByVal value As Boolean)
    mProps.BandValida = value
    mbooModificado = True
End Property

Public Property Get BandValida() As Boolean
    BandValida = mProps.BandValida
End Property

