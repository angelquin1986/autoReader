VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVKProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"GNComprobante"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Type T_PROP
    IDkp As Long
    idproceso As Long
    CodProceso As String
    DescProceso As String
    IdResponsable As Long
    CodResponsable As String
    HorasPre As Currency
    HorasCon As Currency
    Orden As Integer
    Descripcion As String
    FechaInicio As Date
    FechaFinal As Date
    FechaToma As Date
    Estado As Integer
    idCentroDet As Long
    CodCentroDet As String
    NumCentro As Long
    objGNComprobante As GNComprobante 'Referencia al objeto padre
    transid As Long
    HoraToma As Date
    HoraFinal As Date
    Cantidad  As Currency
    Estado1 As Integer
    IdProceso1 As Integer
    CodProceso1 As String
    HoraInicio As Date
    secuencia As Long
    codUsuario As String
    idkpAsignado As Long
    BandGarantia As Long
    FechaFinEspera As Date
    BandUrgente As Boolean
    OrdenUrgente As Long
    CodEstado As String
    CodEstado1 As String
    grupo As String
    objEmpresa As Empresa
End Type
Private mProps As T_PROP

Public Sub VerificaDatos()
Dim item As IVInventario
    If mProps.idproceso = 0 Then
        Err.Raise ERR_INVALIDO, "IVKProceso_VerificaDatos", _
            "Debe seleccionar un proceso."
    End If
Set item = Nothing
End Sub

Friend Sub Grabar()
    Dim sql As String, rs As Recordset
    VerificaDatos
    sql = "SELECT * FROM IVKProceso WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !transid = mProps.objGNComprobante.transid
        !idproceso = mProps.idproceso
        !IdResponsable = mProps.IdResponsable     'AUC 15/09/06
        !Cantidad = mProps.Cantidad  'mProps.HorasPre
        !Orden = mProps.Orden
        !Descripcion = mProps.Descripcion
        !Estado = mProps.Estado
        !Estado1 = mProps.Estado1
        !idCentroDet = mProps.NumCentro
        !secuencia = mProps.secuencia
        !codUsuario = mProps.codUsuario
        !grupo = mProps.grupo
        !BandGarantia = Abs(mProps.BandGarantia)
        'ACTIVADO 09/05/2017
        If mProps.FechaInicio <> "00:00:00" Then !FechaInicio = mProps.FechaInicio
        If mProps.FechaFinal <> "00:00:00" Then !FechaFin = mProps.FechaFinal
        If mProps.FechaToma <> "00:00:00" Then !FechaToma = mProps.FechaToma
        If mProps.HoraInicio <> "00:00:00" Then !HoraInicio = mProps.HoraInicio
        If mProps.HoraToma <> "00:00:00" Then !HoraToma = mProps.HoraToma
        If mProps.HoraFinal <> "00:00:00" Then !HoraFIN = mProps.HoraFinal
        'If mProps.HoraFinal <> "00:00:00" Then !HoraFIN = mProps.HoraFinal
        !idkpAsignado = mProps.idkpAsignado
        If mProps.FechaFinEspera <> "00:00:00" Then !FechaFinEspera = mProps.FechaFinEspera
        !BandUrgente = Abs(mProps.BandUrgente)
        !OrdenUrgente = mProps.OrdenUrgente
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.IDkp = !IDkp             'Guarda el idkp(Autonumerico)
        .Close
        
    End With
    
    ActualizaPrimerProceso
    
    Set rs = Nothing
'        mProps.objEmpresa.RaiseRefreshEvent REFRESH_DETALLEPROCESO
End Sub

Friend Sub RecuperarProcesos(transid As Long, codPro As String)

End Sub
Friend Sub Recuperar(rs As Recordset)
Dim pos1 As Integer
    With rs
        If Not .EOF Then
'            mProps.IDkp = !IDkp
'            If Not IsNull(!idproceso) Then mProps.idproceso = !idproceso
'            If Not IsNull(!CodProceso) Then mProps.CodProceso = !CodProceso
'
'            If Not IsNull(!IdResponsable) Then mProps.IdResponsable = !IdResponsable
'
'            If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
'            If Not IsNull(!Estado) Then mProps.Estado = !Estado
'            If Not IsNull(!Estado1) Then mProps.Estado1 = !Estado1
'            If Not IsNull(!codUsuario) Then mProps.codUsuario = !codUsuario
'            If Not IsNull(!secuencia) Then mProps.secuencia = !secuencia
'            If Not IsNull(!BandGarantia) Then mProps.BandGarantia = !BandGarantia
'            -------------
            mProps.IDkp = !IDkp
            If Not IsNull(!idproceso) Then mProps.idproceso = !idproceso:
            If Not IsNull(!CodProceso) Then mProps.CodProceso = !CodProceso
            If Not IsNull(!DescProceso) Then mProps.DescProceso = !DescProceso
            If Not IsNull(!codUsuario) Then mProps.CodResponsable = !codUsuario
            If Not IsNull(!grupo) Then mProps.grupo = !grupo
            
            pos1 = InStr(1, UCase(gobjMain.EmpresaActual.GNOpcion.NombreEmpresa), "MAHAV")
            If pos1 > 0 Then
                'recupera para maderamica
                If Not IsNull(!Cantidad) Then mProps.HorasPre = !Cantidad
                If Not IsNull(!codUsuario) Then mProps.CodResponsable = !codUsuario
'                If Not IsNull(!Grupo) Then mProps.CodResponsable = !Grupo
            Else
                'recupera para ecuamueble
                If Not IsNull(!Cantidad) Then mProps.Cantidad = !Cantidad
                If Not IsNull(!codUsuario) Then mProps.codUsuario = !codUsuario
                If Not IsNull(!grupo) Then mProps.grupo = !grupo
            End If
            If Not IsNull(!transid) Then mProps.transid = !transid
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
            If Not IsNull(!FechaInicio) Then mProps.FechaInicio = !FechaInicio
            If Not IsNull(!HoraInicio) Then mProps.HoraInicio = !HoraInicio
            If Not IsNull(!FechaToma) Then mProps.FechaToma = !FechaToma
            If Not IsNull(!HoraToma) Then mProps.HoraToma = !HoraToma
            If Not IsNull(!FechaFin) Then mProps.FechaFinal = !FechaFin
            If Not IsNull(!HoraFIN) Then mProps.HoraFinal = !HoraFIN
            If Not IsNull(!Estado) Then mProps.Estado = !Estado
            If Not IsNull(!Estado1) Then mProps.Estado1 = !Estado1
            If Not IsNull(!secuencia) Then mProps.secuencia = !secuencia
            If Not IsNull(!BandGarantia) Then mProps.BandGarantia = !BandGarantia
            If Not IsNull(!BandUrgente) Then mProps.BandUrgente = !BandUrgente
            If Not IsNull(!OrdenUrgente) Then mProps.OrdenUrgente = !OrdenUrgente
        End If
    End With
End Sub

Friend Function Clone() As IVKProceso
    Dim nuevo As IVKProceso
    
    Set nuevo = New IVKProceso
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROP)          'Para el método Clone
    mProps = value
    mProps.IDkp = 0
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Set GNComprobante(value As GNComprobante)
    Set mProps.objGNComprobante = value
End Property

Public Property Get GNComprobante() As GNComprobante
    Set GNComprobante = mProps.objGNComprobante
End Property
Friend Property Set objEmpresa(value As Empresa)
    Set mProps.objEmpresa = value
End Property

Public Property Get objEmpresa() As Empresa
    Set objEmpresa = mProps.objEmpresa
End Property
'AUC cambiado para el reproceso de recetas 19/10/06
Public Property Get IDkp() As Long
    IDkp = mProps.IDkp
End Property


Public Property Get idproceso() As Long
    idproceso = mProps.idproceso
End Property

Public Property Let idproceso(ByVal value As Long)
    mProps.idproceso = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Let CodProceso(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodProceso Then
        If Len(value) > 0 Then
            'Actualiza IdInventario también
            sql = "SELECT IdProceso " & _
                " FROM ivProceso WHERE Codproceso='" & value & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVKProceso_CodProceso", MSGERR_NOHAYCODIGO
            Else
                mProps.idproceso = rs!idproceso
            End If
            rs.Close
        Else
            mProps.idproceso = 0
        End If
    End If
    Set rs = Nothing
    mProps.CodProceso = value
   ' mProps.objGNComprobante.SetModificado
End Property
Public Property Let HorasPre(ByVal value As Currency)
    mProps.HorasPre = value
'    mProps.objGNComprobante.SetModificado
End Property

Public Property Get HorasPre() As Currency
    HorasPre = mProps.HorasPre
End Property
Public Property Let HorasCon(ByVal value As Currency)
    mProps.HorasCon = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get HorasCon() As Currency
    HorasCon = mProps.HorasCon
End Property

Public Property Let Orden(ByVal value As Integer)
    mProps.Orden = value
'    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property

Public Property Let Descripcion(ByVal value As String)
    value = RTrim$(Mid$(value, 1, 512))
    If Len(value) > 512 Then
        Err.Raise ERR_INVALIDO, "IVKProceso_Descripcion", MSGERR_INVALIDO
    End If
    mProps.Descripcion = value
'    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Descripcion() As String
    Descripcion = RTrim$(mProps.Descripcion)
End Property
Public Property Let DescProceso(ByVal value As String)
    mProps.DescProceso = value
End Property

Public Property Get DescProceso() As String
    DescProceso = RTrim$(mProps.DescProceso)
End Property

Private Sub Class_Terminate()
    Set mProps.objGNComprobante = Nothing
End Sub


Public Property Let CodResponsable(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodResponsable Then
        If Len(value) > 0 Then
            'Actualiza IdInventario también
            sql = "SELECT Idresponsable " & _
                " FROM Gnresponsable WHERE CodResponsable='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVKardex_CodResponsable", MSGERR_NOHAYCODIGO
            Else
                mProps.IdResponsable = rs!IdResponsable
                
            End If
            rs.Close
        Else
            mProps.IdResponsable = 0
        End If
    End If
    Set rs = Nothing
    mProps.CodResponsable = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Let FechaInicio(value As Date)
    
    If mProps.FechaInicio <> value Then
        mProps.FechaInicio = value
        
    End If
End Property

Public Property Get FechaInicio() As Date
    FechaInicio = mProps.FechaInicio
End Property
Public Property Let FechaFinal(value As Date)
    'If mbooSoloVer Then Err.Raise ERR_SOLOVER, "IVKProceso.FechaFinal", MSGERR_SOLOVER
    If mProps.FechaFinal <> value Then
        mProps.FechaFinal = value
     '   mbooModificado = True
    End If
End Property

Public Property Get FechaFinal() As Date
    FechaFinal = mProps.FechaFinal
End Property
Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property
Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property
Public Property Get CodProceso() As String
    CodProceso = mProps.CodProceso
End Property
Public Property Get CodResponsable() As String
    CodResponsable = mProps.CodResponsable
End Property

Public Property Let Estado(ByVal value As Integer)
        mProps.Estado = value
End Property

Public Property Get Estado() As Integer
    Estado = mProps.Estado
End Property

Public Sub cambiaEstado()
Dim sql As String
            sql = "Update IVKProceso " & _
                  "set estado = " & Me.Estado & _
                  "Where idkp= " & Me.IDkp
        Me.Empresa.EjecutarSQL sql, 1
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_DETALLEPROCESO
End Sub

Public Property Let idCentroDet(ByVal value As Long)
    mProps.idCentroDet = value
'    mProps.objGNComprobante.SetModificado
End Property
Public Property Get idCentroDet() As Long
    idCentroDet = mProps.idCentroDet
End Property

Public Property Let CodCentroDet(ByVal value As String)
    Dim sql As String
    mProps.CodCentroDet = value
    mProps.objGNComprobante.SetModificado
End Property
Public Property Get CodCentroDet() As String
    CodCentroDet = mProps.CodCentroDet
End Property

Private Sub GrabarCentroCosto()
Dim sql As String
Dim rs As Recordset
    sql = "SELECT * FROM gncentrocostodetalle WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !CodCentroDet = mProps.CodCentroDet
        !DescripcionDet = mProps.CodCentroDet
        !Orden = mProps.Orden 'el mismo orden de ivkardex
        !idCentro = mProps.objGNComprobante.idCentro
        !NumCentro = mProps.NumCentro
        .Update
    #If DAOLIB Then
        rs.Bookmark = rs.LastModified
    #End If
    .Move 0             'Para actualizar
         mProps.idCentroDet = !idCentroDet              'Guarda el Id(Autonumerico)
        .Close
    End With
    Set rs = Nothing
End Sub

Friend Sub RecuperarProcesosM(ByVal transid As Long, ByVal CodCentroDet As String, ByVal codPro As String)
    Dim sql As String, rs As Recordset
On Error GoTo CapturaError
    sql = "SELECT kp.idkp,p.codproceso,p.descripcion as DescProceso ,r.codresponsable,kp.idproceso,"
    sql = sql & " kp.idresponsable,kp.orden,kp.descripcion,kp.fechainicio,kp.fechafin,kp.cantidad, "
    sql = sql & " SUM(isnull(datediff(hour,dp.horainicio,dp.horafinal),0)) as hora,kp.Estado "
    sql = sql & " FROM IVKProceso kp "
    sql = sql & " INNER JOIN GNCentroCostoDetalle gcc ON gcc.idcentrodet = kp.idcentrodet   "
    sql = sql & " INNER JOIN IVProceso P ON p.idproceso = kp.idproceso"
    sql = sql & " LEFT JOIN gnresponsable r ON R.idresponsable = kp.idresponsable "
    sql = sql & " LEFT JOIN DetalleProceso dp on dp.idkp = kp.idkp "
    sql = sql & " Where  p.descripcion = '" & codPro & "'"
    sql = sql & " AND gcc.codcentrodet = '" & CodCentroDet & "'"
    
    sql = sql & " group by kp.idkp,p.codproceso,p.descripcion,r.codresponsable,kp.idproceso,"
    sql = sql & " kp.IdResponsable , kp.Cantidad, kp.Orden, kp.Descripcion, kp.FechaInicio, kp.fechafin, kp.Cantidad,kp.Estado"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
            mProps.IDkp = !IDkp
            If Not IsNull(!idproceso) Then mProps.idproceso = !idproceso:
            If Not IsNull(!CodProceso) Then mProps.CodProceso = !CodProceso
            If Not IsNull(!DescProceso) Then mProps.DescProceso = !DescProceso
            If Not IsNull(!IdResponsable) Then mProps.IdResponsable = !IdResponsable
            If Not IsNull(!CodResponsable) Then mProps.CodResponsable = !CodResponsable
            If Not IsNull(!Cantidad) Then mProps.HorasPre = !Cantidad
            If Not IsNull(!Hora) Then mProps.HorasCon = !Hora
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
            If Not IsNull(!FechaInicio) Then mProps.FechaInicio = !FechaInicio
            If Not IsNull(!FechaFin) Then mProps.FechaFinal = !FechaFin
            If Not IsNull(!Estado) Then mProps.Estado = !Estado

        .Close
    End With
    Set rs = Nothing
    Exit Sub
CapturaError:
    MsgBox Err.Description
    rs.Close
    Set rs = Nothing
    Exit Sub
End Sub

Public Property Get NumCentro() As Long
    NumCentro = mProps.NumCentro
End Property

Public Property Let NumCentro(ByVal value As Long)
    mProps.NumCentro = value
    mProps.objGNComprobante.SetModificado
End Property


Friend Sub RecuperarColas(ByVal Id As Long)
    Dim sql As String, rs As Recordset
On Error GoTo CapturaError
    sql = "SELECT p.codproceso,p.descripcion as DescProceso ,p.codusuario,p.grupo,kp.* "
    sql = sql & " FROM IVKProceso kp "
    sql = sql & " INNER JOIN IVProceso P ON p.idproceso = kp.idproceso"
    sql = sql & " Where  kp.idkp = " & Id
       
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    With rs
            mProps.IDkp = !IDkp
            If Not IsNull(!idproceso) Then mProps.idproceso = !idproceso
            If Not IsNull(!CodProceso) Then mProps.CodProceso = !CodProceso
            If Not IsNull(!DescProceso) Then mProps.DescProceso = !DescProceso
            If Not IsNull(!codUsuario) Then mProps.CodResponsable = !codUsuario
            If Not IsNull(!grupo) Then mProps.grupo = !grupo
            If Not IsNull(!Cantidad) Then mProps.Cantidad = !Cantidad
            If Not IsNull(!transid) Then mProps.transid = !transid
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
            If Not IsNull(!FechaInicio) Then mProps.FechaInicio = !FechaInicio
            If Not IsNull(!HoraInicio) Then mProps.HoraInicio = !HoraInicio 'activado
            If Not IsNull(!FechaToma) Then mProps.FechaToma = !FechaToma
            If Not IsNull(!HoraToma) Then mProps.HoraToma = !HoraToma
            If Not IsNull(!FechaFin) Then mProps.FechaFinal = !FechaFin
            If Not IsNull(!FechaFinEspera) Then mProps.FechaFinEspera = !FechaFinEspera
            If Not IsNull(!HoraFIN) Then mProps.HoraFinal = !HoraFIN
            If Not IsNull(!idCentroDet) Then mProps.idCentroDet = !idCentroDet
            If Not IsNull(!idkpAsignado) Then mProps.idkpAsignado = !idkpAsignado
            If Not IsNull(!Estado) Then mProps.Estado = !Estado
            If Not IsNull(!Estado1) Then mProps.Estado1 = !Estado1
            If Not IsNull(!secuencia) Then mProps.secuencia = !secuencia
            If Not IsNull(!BandGarantia) Then mProps.BandGarantia = !BandGarantia
            If Not IsNull(!BandUrgente) Then mProps.BandUrgente = !BandUrgente
            If Not IsNull(!OrdenUrgente) Then mProps.OrdenUrgente = !OrdenUrgente
        .Close
    End With
    Set rs = Nothing
    Exit Sub
CapturaError:
    'rs.Close
    Set rs = Nothing
    MsgBox Err.Description
    Exit Sub
End Sub

Public Sub GrabarCola()
    Dim sql As String, rs As Recordset
    VerificaDatos
    sql = "SELECT * FROM IVKProceso WHERE idkp = " & Me.IDkp
    Set rs = Me.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        !transid = mProps.transid
        !idproceso = mProps.idproceso
'        !IdResponsable = mProps.IdResponsable     'AUC 15/09/06
'        !Cantidad = mProps.HorasPre
'        !Orden = mProps.Orden
        !Descripcion = mProps.Descripcion
        !FechaToma = mProps.FechaToma
'        !FechaFin = mProps.FechaFinal
'        !Estado = mProps.Estado
'            'GrabarCentroCosto 'AUC esto no va aqui
'        !idCentroDet = mProps.NumCentro
        .Update
        
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.IDkp = !IDkp             'Guarda el idkp(Autonumerico)
        .Close
        
    End With
    Set rs = Nothing
'        mProps.objEmpresa.RaiseRefreshEvent REFRESH_DETALLEPROCESO
End Sub
Private Sub ActualizaPrimerProceso()
Dim sql As String
Dim rs As Recordset
Dim v As Long, IdProc As Long
Dim v1 As Long
Dim codUsuario As String
Dim grupo As String
On Error GoTo CapturaError
        sql = "Select idproceso from ivproceso  "
        sql = sql & "Where secuencia = 0 "
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
        Do While Not rs.EOF
            IdProc = rs!idproceso
            'codUsuario = rs!cod
            rs.MoveNext
        Loop
    Set rs = Nothing
    
    If IdProc = mProps.idproceso Then 'PARA EL PRIMER PROCESO
        'secuencia = 0 siempre tiene que se el primer proceso
        sql = "select menu,valor from gnestadoprod where menu IN ('Listo','Ejecucion')"
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
        Do While Not rs.EOF
            Select Case rs!Menu
                Case "Listo"
                    v = rs!Valor
                Case "Ejecucion"
                    v1 = rs!Valor
            End Select
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
            sql = " Update ivkproceso set estado =" & v1
            sql = sql & " ,estado1 =" & v
            sql = sql & " ,FechaInicio = '" & Date & "'"
            sql = sql & " ,FechaFin = '" & Date & "'"
            sql = sql & " ,fechaToma = '" & Date & "'"
            sql = sql & " ,HoraInicio = '" & Time & "'"
            sql = sql & " ,HoraToma = '" & Time & "'"
            sql = sql & " ,HoraFin = '" & Time & "'"
            sql = sql & " ,descripcion = 'Ejecutado Automaticamente'"
            sql = sql & " where idkp =  " & mProps.IDkp
            gobjMain.EmpresaActual.EjecutarSQL sql, 1
    End If
    
    'PARA EL SEGUNDO PROCESO DEBE INGRESAR DIRECTAMENTE LA HORA
    If mProps.secuencia = 1 Then
        sql = "Select idproceso,codUsuario,grupo from ivproceso "
        'sql = sql & "Where secuencia = 1 "
        sql = sql & "Where codproceso  = '" & mProps.CodProceso & "' "
        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
        Do While Not rs.EOF
            IdProc = rs!idproceso
            codUsuario = rs!codUsuario
            grupo = rs!grupo

            rs.MoveNext
        Loop
        Set rs = Nothing
    End If
    If IdProc = mProps.idproceso Then 'PARA EL PRIMER PROCESO
'        'secuencia = 0 siempre tiene que se el primer proceso
'        sql = "select valor from gnestadoprod where menu = 'Listo'"
'        Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
'        Do While Not rs.EOF
'            v = rs!Valor
'            rs.MoveNext
'        Loop
'        Set rs = Nothing
        
        
        sql = " Update ivkproceso" ' set estado = " & v
        sql = sql & " SET FechaInicio = '" & Date & "'"
        sql = sql & ",HoraInicio = '" & Time & "'"
     '   sql = sql & ",codUsuario = '" & codUsuario & "'"
        
        If mProps.BandGarantia Then
            sql = sql & " ,FechaToma = '" & Date & "'"
            sql = sql & ",HoraToma = '" & Time & "'"
        End If
        
        sql = sql & " where idkp =  " & mProps.IDkp
        gobjMain.EmpresaActual.EjecutarSQL sql, 1
    End If
    Exit Sub
CapturaError:
    MsgBox Err.Description
    rs.Close
    Set rs = Nothing
    Exit Sub
End Sub
Public Sub grabarNuevo()
End Sub

Public Sub GrabarNew(ByVal bandNew As Boolean)
Dim gne As GNEstadoProd
    Dim sql As String, rs As Recordset
    If bandNew Then
        sql = "SELECT * FROM IVKProceso WHERE 1=0 "
        Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
    Else
        sql = "SELECT * FROM IVKProceso WHERE idkp = " & Me.IDkp
        Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
    End If
    Set gne = gobjMain.EmpresaActual.RecuperaGNEstadoProd(mProps.Estado)
    With rs
        If bandNew Then rs.AddNew
        !transid = mProps.transid
        !idproceso = mProps.idproceso
        
        !Cantidad = mProps.Cantidad
        !Orden = mProps.Orden
        !Descripcion = mProps.Descripcion
        If gne.Menu = "Revision" Or gne.Menu = "Espera" Then
            !FechaInicio = mProps.FechaInicio
            !HoraInicio = mProps.HoraInicio
        Else
            !FechaToma = mProps.FechaToma
            !HoraToma = mProps.HoraToma
        End If
        !Estado = mProps.Estado
        !Estado1 = mProps.Estado1 'segundo estado
        If mProps.Estado1 <> 0 And gne.Menu <> "Espera" Then
            !FechaFin = mProps.FechaFinal
            !HoraFIN = mProps.HoraFinal
        ElseIf gne.Menu = "Espera" Then
            !FechaFinEspera = mProps.FechaFinEspera
'            !horaFin = mProps.HoraFinal
        End If
        !codUsuario = mProps.codUsuario
        !grupo = mProps.grupo
        If bandNew Then !idkpAsignado = mProps.idkpAsignado
        !secuencia = mProps.secuencia
        !BandGarantia = mProps.BandGarantia
        !BandUrgente = Abs(mProps.BandUrgente)
        !OrdenUrgente = mProps.OrdenUrgente
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.IDkp = !IDkp             'Guarda el idkp(Autonumerico)
        .Close
    End With
    Set rs = Nothing
    Set gne = Nothing
'        mProps.objEmpresa.RaiseRefreshEvent REFRESH_DETALLEPROCESO
End Sub

Public Property Let FechaToma(value As Date)
    If mProps.FechaToma <> value Then
        mProps.FechaToma = value
        'mbooModificado = True
    End If
End Property

Public Property Get FechaToma() As Date
    FechaToma = mProps.FechaToma
End Property

Public Property Get transid() As Long
    transid = mProps.transid
End Property

Public Property Let transid(ByVal value As Long)
    mProps.transid = value
End Property
Public Property Let HoraToma(value As Date)
    If mProps.HoraToma <> value Then
        mProps.HoraToma = value
    End If
End Property

Public Property Get HoraToma() As Date
    HoraToma = mProps.HoraToma
End Property

Public Property Get Cantidad() As Long
    Cantidad = mProps.Cantidad
End Property

Public Property Let Cantidad(ByVal value As Long)
    mProps.Cantidad = value
End Property

Public Property Let Estado1(ByVal value As Integer)
        mProps.Estado1 = value
End Property

Public Property Get Estado1() As Integer
    Estado1 = mProps.Estado1
End Property

Public Property Let HoraFinal(value As Date)
    If mProps.HoraFinal <> value Then
        mProps.HoraFinal = value
    End If
End Property

Public Property Get HoraFinal() As Date
    HoraFinal = mProps.HoraFinal
End Property

Public Property Get IdProceso1() As Long
    IdProceso1 = mProps.IdProceso1
End Property

Public Property Let IdProceso1(ByVal value As Long)
    mProps.IdProceso1 = value
End Property


Public Property Let CodProceso1(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodProceso1 Then
        If Len(value) > 0 Then
            'Actualiza IdInventario también
            sql = "SELECT IdProceso " & _
                " FROM ivProceso WHERE Codproceso='" & value & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVKProceso_CodProceso1", MSGERR_NOHAYCODIGO
            Else
                mProps.IdProceso1 = rs!idproceso
            End If
            rs.Close
        Else
            mProps.IdProceso1 = 0
        End If
    End If
    Set rs = Nothing
    mProps.CodProceso1 = value
End Property

Public Property Get CodProceso1() As String
    CodProceso1 = mProps.CodProceso1
End Property
Public Property Let HoraInicio(value As Date)
    If mProps.HoraInicio <> value Then
        mProps.HoraInicio = value
    End If
End Property

Public Property Get HoraInicio() As Date
    HoraInicio = mProps.HoraInicio
End Property

Public Property Get secuencia() As Long
    secuencia = mProps.secuencia
End Property

Public Property Let secuencia(ByVal value As Long)
    mProps.secuencia = value
End Property

Public Property Let codUsuario(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "IVKProceso", MSGERR_INVALIDO
    End If
    mProps.codUsuario = value
'    mbooModificado = True
End Property

Public Property Get codUsuario() As String
    codUsuario = mProps.codUsuario
End Property

Public Property Get idkpAsignado() As Long
    idkpAsignado = mProps.idkpAsignado
End Property

Public Property Let idkpAsignado(ByVal value As Long)
    mProps.idkpAsignado = value
End Property
Public Property Get BandGarantia() As Long
    BandGarantia = mProps.BandGarantia
End Property

Public Property Let BandGarantia(ByVal value As Long)
    mProps.BandGarantia = value
End Property

'para ecuamueble
Public Sub grabarRuta()
    Dim sql As String, rs As Recordset
    VerificaDatos
    
    sql = "SELECT * FROM IVKProceso WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !transid = mProps.objGNComprobante.transid
        !idproceso = mProps.idproceso
        !IdResponsable = mProps.IdResponsable     'AUC 15/09/06
        !Cantidad = mProps.Cantidad
        !Orden = mProps.Orden
        !Descripcion = mProps.Descripcion
        If mProps.FechaInicio <> "00:00:00" Then !FechaInicio = mProps.FechaInicio
        If mProps.HoraInicio <> "00:00:00" Then !HoraInicio = mProps.HoraInicio
        If mProps.FechaFinal <> "00:00:00" Then !FechaFin = mProps.FechaFinal
        If mProps.HoraFinal <> "00:00:00" Then !HoraFIN = mProps.HoraFinal
        If mProps.FechaToma <> "00:00:00" Then !FechaToma = mProps.FechaToma
        If mProps.HoraToma <> "00:00:00" Then !HoraToma = mProps.HoraToma
        !Estado = mProps.Estado
        !Estado1 = mProps.Estado1
        !idCentroDet = mProps.NumCentro
        !secuencia = mProps.secuencia
        !codUsuario = mProps.codUsuario
        !grupo = mProps.grupo
        !BandGarantia = mProps.BandGarantia
        .Update
        
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.IDkp = !IDkp             'Guarda el idkp(Autonumerico)
        .Close
    End With
'    ActualizaPrimerProceso
    Set rs = Nothing
End Sub

Public Property Let FechaFinEspera(ByVal value As Date)
        mProps.FechaFinEspera = value
End Property

Public Property Get FechaFinEspera() As Date
    FechaFinEspera = mProps.FechaFinEspera
End Property

Public Property Let BandUrgente(ByVal value As Boolean)
        mProps.BandUrgente = value
End Property

Public Property Get BandUrgente() As Boolean
    BandUrgente = mProps.BandUrgente
End Property

Public Property Let OrdenUrgente(ByVal value As Long)
        mProps.OrdenUrgente = value
End Property

Public Property Get OrdenUrgente() As Long
    OrdenUrgente = mProps.OrdenUrgente
End Property

Public Property Let CodEstado(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodEstado Then
        If Len(value) > 0 Then
            'Actualiza IdInventario también
            sql = "SELECT valor " & _
                " FROM gnestadoprod WHERE CodEstado='" & value & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVKProceso_CodEstado1", MSGERR_NOHAYCODIGO
            Else
                mProps.Estado = rs!Valor
            End If
            rs.Close
        Else
            mProps.Estado = 0
        End If
    End If
    Set rs = Nothing
    mProps.CodEstado = value
End Property

Public Property Get CodEstado() As String
    CodEstado = mProps.CodEstado
End Property

Public Property Let CodEstado1(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodEstado1 Then
        If Len(value) > 0 Then
            'Actualiza IdInventario también
            sql = "SELECT valor " & _
                " FROM gnestadoprod WHERE CodEstado='" & value & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVKProceso_CodEstado1", MSGERR_NOHAYCODIGO
            Else
                mProps.Estado1 = rs!Valor
            End If
            rs.Close
        Else
            mProps.Estado1 = 0
        End If
    End If
    Set rs = Nothing
    mProps.CodEstado1 = value
End Property

Public Property Get CodEstado1() As String
    CodEstado1 = mProps.CodEstado1
End Property


Public Property Let grupo(ByVal value As String)
    If Len(value) > 40 Then
        Err.Raise ERR_INVALIDO, "IVKProceso", MSGERR_INVALIDO
    End If
    mProps.grupo = value
'    mbooModificado = True
End Property

Public Property Get grupo() As String
    grupo = mProps.grupo
End Property



