VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GNCall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
    Option Explicit

Private Type T_PROPS
    Id As Long
    transid As Long
    idpckardex1 As Long
    idpckardex2 As Long
    IdCliente As Long 'pendiente
    fecha As Date
    Hora As Date
    Contesta As String
    Relacion As String
    Observa As String
    FechaVisita As Date
    HoraProxGestion As Date
    IdCobrador As Long
    CodCobrador As String
    Orden As Integer
    idGestion As Long
    CodGestion As String
    codUsuario As String
    CodUsuarioMod As String
    FechaMod As Date
    objGNComprobante As GNComprobante 'cambiado ya no utilizo
    objEmpresa As Empresa
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    sql = "SELECT codtrans,numtrans FROM GNComprobante g"
    sql = sql & " Inner Join ivkardex ivk on ivk.transid = g.transid"
    sql = sql & " WHERE idTransFuente=" & mProps.transid
    sql = sql & " And ivk.idpadre=" & mProps.idpckardex1
    sql = sql & " And ivk.idpadreSub=" & mProps.idpckardex2
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "GNCall_Eliminar ", MSGERR_NOELIMINABLE & " Trans Rel. " & rs!CodTrans & " " & rs!NumTrans
    End If
    rs.Close
    'Elimina el registro
    sql = "DELETE FROM GNCall WHERE Id=" & mProps.Id
    mProps.objEmpresa.Execute sql, True
    mbooEsNuevo = True
    mbooModificado = True
End Sub

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.Id = 0
    mbooModificado = True
End Property

Friend Property Set GNComprobante(value As GNComprobante)
   ' Set mProps.objGNComprobante = value
End Property

Public Property Get GNComprobante() As GNComprobante
    'Set GNComprobante = mProps.objGNComprobante
End Property
Friend Property Set objEmpresa(value As Empresa)
    Set mProps.objEmpresa = value
End Property

Public Property Get objEmpresa() As Empresa
    Set objEmpresa = mProps.objEmpresa
End Property


Public Property Get idpckardex1() As Long
    idpckardex1 = mProps.idpckardex1
End Property

Public Property Let idpckardex1(ByVal value As Long)
    mProps.idpckardex1 = value
    mbooModificado = True
End Property
Public Property Get idpckardex2() As Long
    idpckardex2 = mProps.idpckardex2
End Property

Public Property Let idpckardex2(ByVal value As Long)
    mProps.idpckardex2 = value
    mbooModificado = True
End Property

Public Property Get EsNuevo() As Boolean
    EsNuevo = mbooEsNuevo
End Property

Public Property Get Modificado() As Boolean
    Modificado = mbooModificado
End Property

Friend Property Let Modificado(ByVal value As Boolean)
    mbooModificado = value
End Property

Private Sub Class_Initialize()
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objGNComprobante = Nothing
End Sub

Public Property Get transid() As Long
    transid = mProps.transid
End Property
Public Property Let transid(ByVal value As Long)
    mProps.transid = value
    mbooModificado = True
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property
Public Property Let Orden(ByVal value As Integer)
    mProps.Orden = value
    mbooModificado = True
End Property

Public Sub Grabar()
    Dim sql As String, i As Integer
    Dim rs As Recordset, rsDetalle As Recordset

    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    On Error GoTo ErrTrap
    
    
    If mbooEsNuevo Then
''        mProps.codUsuario = gobjUsuarioActual.codUsuario
''        mProps.CodUsuarioMod = gobjUsuarioActual.codUsuario
''        mProps.FechaMod = Now
        'esto va en el formulario
        
        sql = "SELECT * FROM GNCall WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    'Si es modificada
    Else
        Me.CodUsuarioMod = gobjUsuarioActual.codUsuario
        Me.FechaMod = Now
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM GNCall WHERE Id=" & mProps.Id
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
    End If
    
    'Graba el registro
    With rs
        
        !transid = mProps.transid
        !idpckardex1 = mProps.idpckardex1
        !idpckardex2 = mProps.idpckardex2
        !fecha = mProps.fecha
        !Hora = mProps.Hora
        !Contesta = mProps.Contesta
        !Relacion = mProps.Relacion
        !Observacion = mProps.Observa
        !FechaVisita = mProps.FechaVisita
        !HoraProxGestion = mProps.HoraProxGestion
        !IdCobrador = mProps.IdCobrador
        !idGestion = mProps.idGestion
        !codUsuario = mProps.codUsuario
        !CodUsuarioMod = mProps.CodUsuarioMod
        !FechaMod = mProps.FechaMod
       
        !Orden = mProps.Orden
        .Update
        rs.Move 0           'Para actualizar
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    Exit Sub
ErrTrap:
 MsgBox Err.Description
End Sub

Public Sub VerificaDatos()
'    Dim cant As Currency
'    If mProps.idpckardex = 0 Then
'        Err.Raise ERR_INVALIDO, "GnCall_VerificaDatos", _
'            "La id no puede ser 0. (" & mProps.idpckardex & ")"
'    End If
End Sub

Public Property Get fecha() As Date
    fecha = mProps.fecha
End Property

Public Property Let fecha(ByVal value As Date)
    mProps.fecha = value
    mbooModificado = True
End Property
Public Property Get Hora() As Date
    Hora = mProps.Hora
End Property

Public Property Let Hora(ByVal value As Date)
    mProps.Hora = value
    mbooModificado = True
End Property

Public Property Get Contesta() As String
    Contesta = mProps.Contesta
End Property

Public Property Let Contesta(ByVal value As String)
    mProps.Contesta = value
    mbooModificado = True
End Property
Public Property Get Relacion() As String
    Relacion = mProps.Relacion
End Property

Public Property Let Relacion(ByVal value As String)
    mProps.Relacion = value
    mbooModificado = True
End Property
Public Property Get Observa() As String
    Observa = mProps.Observa
End Property

Public Property Let Observa(ByVal value As String)
    If Len(value) > 255 Then
        Err.Raise ERR_INVALIDO, "GNCALL_Observacion", MSGERR_INVALIDO
    End If
    mProps.Observa = value
    mbooModificado = True
End Property
Public Property Get FechaVisita() As Date
    FechaVisita = mProps.FechaVisita
End Property

Public Property Let FechaVisita(ByVal value As Date)
    mProps.FechaVisita = value
    mbooModificado = True
End Property

Public Property Let CodCobrador(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodCobrador Then
        If Len(value) > 0 Then
            'Actualiza IdVendedor también
            sql = "SELECT IdVendedor FROM FCVendedor WHERE CodVendedor='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNCall.CodVendedor", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdCobrador = rs!IdVendedor
            End If
            rs.Close
            Set rs = Nothing
        Else
            mProps.IdCobrador = 0
        End If
        mProps.CodCobrador = value
        mbooModificado = True
    End If
End Property

Public Property Get CodCobrador() As String
    CodCobrador = Trim$(mProps.CodCobrador)
End Property

Friend Sub Recuperar(key As Long)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT g.*, codVendedor,codgestion FROM GNCall g "
    sql = sql & "LEFT JOIN fcVendedor fcv on fcv.idvendedor=g.idCobrador "
    sql = sql & "LEFT JOIN pcgestion pcg  on pcg.id=g.idgestion "
    sql = sql & "Where g.Id = " & key
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.Id = !Id
        mProps.transid = !transid
        mProps.idpckardex1 = !idpckardex1
        mProps.idpckardex2 = !idpckardex2
        If Not IsNull(!fecha) Then mProps.fecha = !fecha
        If Not IsNull(!Hora) Then mProps.Hora = !Hora
        If Not IsNull(!Contesta) Then mProps.Contesta = !Contesta
        If Not IsNull(!Relacion) Then mProps.Relacion = !Relacion
        If Not IsNull(!Observacion) Then mProps.Observa = !Observacion
        If Not IsNull(!FechaVisita) Then mProps.FechaVisita = !FechaVisita
        If Not IsNull(!HoraProxGestion) Then mProps.HoraProxGestion = !HoraProxGestion
        If Not IsNull(!IdCobrador) Then mProps.IdCobrador = !IdCobrador
        If Not IsNull(!CodVendedor) Then mProps.CodCobrador = !CodVendedor
        mProps.idGestion = !idGestion
        If Not IsNull(!CodGestion) Then mProps.CodGestion = !CodGestion
        If Not IsNull(!codUsuario) Then mProps.codUsuario = !codUsuario
        If Not IsNull(!CodUsuarioMod) Then mProps.CodUsuarioMod = !CodUsuarioMod
        If Not IsNull(!FechaMod) Then mProps.FechaMod = !FechaMod
        
'        If Not IsNull(!Orden) Then mProps.Orden = !Orden
        'mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub
Public Property Let CodGestion(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodCobrador Then
        If Len(value) > 0 Then
            'Actualiza IdVendedor también
            sql = "SELECT Id FROM PCGestion WHERE CodGestion='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNCall.CodGestion", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.idGestion = rs!Id
            End If
            rs.Close
            Set rs = Nothing
        Else
            mProps.idGestion = 0
        End If
        mProps.CodGestion = value
        mbooModificado = True
    End If
End Property

Public Property Get CodGestion() As String
    CodGestion = Trim$(mProps.CodGestion)
End Property
Public Property Let CodUsuariod(value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "GNCall.CodUsuario", MSGERR_INVALIDO
    End If

    mProps.codUsuario = value
    mbooModificado = True
End Property

Public Property Get codUsuario() As String
    codUsuario = Trim$(mProps.codUsuario)
End Property


Public Property Let CodUsuarioMod(value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "GNCall.CodUsuarioMod", MSGERR_INVALIDO
    End If

    mProps.CodUsuarioMod = value
    mbooModificado = True
End Property

Public Property Get CodUsuarioMod() As String
    CodUsuarioMod = Trim$(mProps.CodUsuarioMod)
End Property

Public Property Let FechaMod(value As Date)
    'Cuando SoloVer está establecido a True, genera error
   ' If mbooSoloVer Then Err.Raise ERR_SOLOVER, "GNCall.FechaMod", MSGERR_SOLOVER
    'Solo cuando cambia el valor
    If mProps.FechaMod <> value Then
        mProps.FechaMod = value
        mbooModificado = True
    End If
End Property

Public Property Get FechaMod() As Date
    FechaMod = mProps.FechaMod
End Property


Public Property Let HoraProxGestion(value As Date)
    'Cuando SoloVer está establecido a True, genera error
    'If mbooSoloVer Then Err.Raise ERR_SOLOVER, "GNCall.HoraProxGestion", MSGERR_SOLOVER
    
    'Solo cuando cambia el valor
    If mProps.HoraProxGestion <> value Then
        mProps.HoraProxGestion = value
        mbooModificado = True
    End If
End Property

Public Property Get HoraProxGestion() As Date
    HoraProxGestion = mProps.HoraProxGestion
End Property
Public Property Let IdCobrador(ByVal value As Long)
    mProps.IdCobrador = value
    mbooModificado = True
End Property

Public Property Get IdCobrador() As Long
    IdCobrador = mProps.IdCobrador
End Property

Public Function GrabarDesdeTablet() As Double
    Dim sql As String, i As Integer
    Dim rs As Recordset, rsDetalle As Recordset
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Function
    On Error GoTo ErrTrap
    If mbooEsNuevo Then
        sql = "SELECT * FROM GNCall WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    'Si es modificada
    Else
        Me.CodUsuarioMod = gobjUsuarioActual.codUsuario
        Me.FechaMod = Now
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM GNCall WHERE Id=" & mProps.Id
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
    End If
    'Graba el registro
    With rs
        !transid = mProps.transid
        !idpckardex1 = mProps.idpckardex1
        !idpckardex2 = mProps.idpckardex2
        !fecha = mProps.fecha
        !Hora = mProps.Hora
        !Contesta = mProps.Contesta
        !Relacion = mProps.Relacion
        !Observacion = mProps.Observa
        !FechaVisita = mProps.FechaVisita
        !HoraProxGestion = mProps.HoraProxGestion
        !IdCobrador = mProps.IdCobrador
        !idGestion = mProps.idGestion
        !codUsuario = mProps.codUsuario
        !CodUsuarioMod = mProps.CodUsuarioMod
        !FechaMod = mProps.FechaMod
        !Orden = mProps.Orden
        .Update
        rs.Move 0           'Para actualizar
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    GrabarDesdeTablet = True
    Exit Function
ErrTrap:
 MsgBox Err.Description
End Function
