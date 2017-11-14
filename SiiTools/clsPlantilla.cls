VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlantilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type T_PROPS
    Coneccion As ADODB.Connection
    IdPlantilla As Long
    CodPlantilla As String
    Descripcion As String
    Tipo As Byte
    PrefijoNombreArchivo As String
    BandRangoFechaHora As Boolean
    BandTipoFecha As Boolean
    BandIgnorarContabilidad As Boolean
    BandIgnorarDocAsignado As Boolean
    BandActualizaCatalogos As Boolean
    BandFiltrarxBodega As Boolean
    ListaBodegas As String
    ListaCatalogos As String
    ListaTransacciones As String
    BandActualizarCosto As Boolean
    BandGuardarResultado As Boolean
    BandValida As Boolean
    FechaDesde As Date
    FechaHasta As Date
    HoraDesde As Date
    HoraHasta As Date
    FechaGrabado As Date
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Public Function Clone() As clsPlantilla
    Dim nuevo As clsPlantilla
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New clsPlantilla
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdPlantilla = 0
    mbooModificado = True
End Property

Public Property Let Coneccion(cn As ADODB.Connection)
    Set mProps.Coneccion = cn
End Property

Public Property Get Coneccion() As ADODB.Connection
    Set Coneccion = mProps.Coneccion
End Property

Public Property Get IdPlantilla() As Long
    IdPlantilla = mProps.IdPlantilla
End Property

Public Property Let CodPlantilla(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "Plantilla.CodPlantilla", MSGERR_INVALIDO
    End If
    mProps.CodPlantilla = value
    mbooModificado = True
End Property

Public Property Get CodPlantilla() As String
    CodPlantilla = mProps.CodPlantilla
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "Plantilla.Descripcion", MSGERR_INVALIDO
    End If
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = mProps.Descripcion
End Property

Public Property Let Tipo(ByVal value As Integer)
    If (value < 0) Or (value > 1) Then
        Err.Raise ERR_INVALIDO, "Plantilla.Tipo", MSGERR_INVALIDO
    End If
    mProps.Tipo = value
    mbooModificado = True
End Property

Public Property Get Tipo() As Integer
    Tipo = mProps.Tipo
End Property

Public Property Let PrefijoNombreArchivo(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "Plantilla.PrefijoNombreArchivo", MSGERR_INVALIDO
    End If
    mProps.PrefijoNombreArchivo = value
    mbooModificado = True
End Property

Public Property Get PrefijoNombreArchivo() As String
    PrefijoNombreArchivo = mProps.PrefijoNombreArchivo
End Property

Public Property Let BandRangoFechaHora(ByVal value As Boolean)
    mProps.BandRangoFechaHora = value
    mbooModificado = True
End Property

Public Property Get BandRangoFechaHora() As Boolean
    BandRangoFechaHora = mProps.BandRangoFechaHora
End Property

Public Property Let BandTipoFecha(ByVal value As Boolean)
    mProps.BandTipoFecha = value
    mbooModificado = True
End Property

Public Property Get BandTipoFecha() As Boolean
    BandTipoFecha = mProps.BandTipoFecha
End Property

Public Property Let BandIgnorarContabilidad(ByVal value As Boolean)
    mProps.BandIgnorarContabilidad = value
    mbooModificado = True
End Property

Public Property Get BandIgnorarContabilidad() As Boolean
    BandIgnorarContabilidad = mProps.BandIgnorarContabilidad
End Property

Public Property Let BandIgnorarDocAsignado(ByVal value As Boolean)
    mProps.BandIgnorarDocAsignado = value
    mbooModificado = True
End Property

Public Property Get BandIgnorarDocAsignado() As Boolean
    BandIgnorarDocAsignado = mProps.BandIgnorarDocAsignado
End Property

Public Property Let BandActualizaCatalogos(ByVal value As Boolean)
    mProps.BandActualizaCatalogos = value
    mbooModificado = True
End Property

Public Property Get BandActualizaCatalogos() As Boolean
    BandActualizaCatalogos = mProps.BandActualizaCatalogos
End Property

Public Property Let BandFiltrarxBodega(ByVal value As Boolean)
    mProps.BandFiltrarxBodega = value
    mbooModificado = True
End Property

Public Property Get BandFiltrarxBodega() As Boolean
    BandFiltrarxBodega = mProps.BandFiltrarxBodega
End Property

Public Property Let ListaBodegas(ByVal value As String)
    mProps.ListaBodegas = value
    mbooModificado = True
End Property

Public Property Get ListaBodegas() As String
    ListaBodegas = mProps.ListaBodegas
End Property

Public Property Let ListaCatalogos(ByVal value As String)
    mProps.ListaCatalogos = value
    mbooModificado = True
End Property

Public Property Get ListaCatalogos() As String
    ListaCatalogos = mProps.ListaCatalogos
End Property

Public Property Let ListaTransacciones(ByVal value As String)
    mProps.ListaTransacciones = value
    mbooModificado = True
End Property

Public Property Get ListaTransacciones() As String
    ListaTransacciones = mProps.ListaTransacciones
End Property

Public Property Let BandActualizarCosto(ByVal value As Boolean)
    mProps.BandActualizarCosto = value
    mbooModificado = True
End Property

Public Property Get BandActualizarCosto() As Boolean
    BandActualizarCosto = mProps.BandActualizarCosto
End Property

Public Property Let BandGuardarResultado(ByVal value As Boolean)
    mProps.BandGuardarResultado = value
    mbooModificado = True
End Property

Public Property Get BandGuardarResultado() As Boolean
    BandGuardarResultado = mProps.BandGuardarResultado
End Property

Public Property Let BandValida(ByVal value As Boolean)
    mProps.BandValida = value
    mbooModificado = True
End Property

Public Property Get BandValida() As Boolean
    BandValida = mProps.BandValida
End Property

Public Property Let FechaDesde(value As Date)
    mProps.FechaDesde = value
    mbooModificado = True
End Property

Public Property Get FechaDesde() As Date
    FechaDesde = mProps.FechaDesde
End Property

Public Property Let FechaHasta(value As Date)
    mProps.FechaHasta = value
    mbooModificado = True
End Property

Public Property Get FechaHasta() As Date
    FechaHasta = mProps.FechaHasta
End Property

Public Property Let HoraDesde(value As Date)
    mProps.HoraDesde = value
    mbooModificado = True
End Property

Public Property Get HoraDesde() As Date
    HoraDesde = mProps.HoraDesde
End Property

Public Property Let HoraHasta(value As Date)
    mProps.HoraHasta = value
    mbooModificado = True
End Property

Public Property Get HoraHasta() As Date
    HoraHasta = mProps.HoraHasta
End Property

Public Property Get FechaGrabado() As Date
    FechaGrabado = mProps.FechaGrabado
End Property

Public Property Get EsNuevo() As Boolean
    EsNuevo = mbooEsNuevo
End Property

Public Property Get Modificado() As Boolean
    Modificado = mbooModificado
End Property

Public Property Get BDDestino() As String
    If Len(mProps.PrefijoNombreArchivo) > 0 Then
        BDDestino = mProps.PrefijoNombreArchivo & "_" & _
                     Format(Date, "dd-mm-yyyy") & "_" & Format(Time, "HH-mm-ss") & ".mdb"
    Else
        BDDestino = gobjMain.EmpresaActual.CodEmpresa & "_" & _
                    Format(Date, "dd-mm-yyyy") & ".mdb"
    End If
End Property

Public Function Grabar() As Boolean
    Dim sql As String, fgrabado As Date, i As Long
    Dim rs As Recordset
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Function
    
       
    'Verifica si no repite el código
    sql = "SELECT CodPlantilla FROM Plantilla_EI " & _
          "WHERE (CodPlantilla='" & mProps.CodPlantilla & "') AND " & _
                "(IdPlantilla<>" & mProps.IdPlantilla & ")"
    Set rs = AbrirTablaLectura(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "Plantilla", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    On Error GoTo ErrTrap

    Me.Coneccion.BeginTrans   '======= Comienza la transacción
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM Plantilla_EI WHERE 1=0"
        Set rs = AbrirTablaEscritura(sql)
        rs.AddNew
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM Plantilla_EI WHERE IdPlantilla=" & mProps.IdPlantilla
        Set rs = AbrirTablaEscritura(sql)
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "Plantilla_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en la base
    With rs
        !CodPlantilla = mProps.CodPlantilla
        !Descripcion = mProps.Descripcion
        !Tipo = mProps.Tipo
        !PrefijoNombreArchivo = mProps.PrefijoNombreArchivo
        !BandRangoFechaHora = mProps.BandRangoFechaHora
        !BandTipoFecha = mProps.BandTipoFecha
        !BandIgnorarContabilidad = mProps.BandIgnorarContabilidad
        !BandIgnorarDocAsignado = mProps.BandIgnorarDocAsignado
        !BandActualizaCatalogos = mProps.BandActualizaCatalogos
        !BandFiltrarxBodega = mProps.BandFiltrarxBodega
        !ListaBodegas = mProps.ListaBodegas
        !ListaCatalogos = mProps.ListaCatalogos
        !ListaTransacciones = mProps.ListaTransacciones
        !BandActualizarCosto = mProps.BandActualizarCosto
        !BandGuardarResultado = mProps.BandGuardarResultado
        !BandValida = mProps.BandValida
        !FechaDesde = IIf((mProps.FechaDesde = 0), Now, mProps.FechaDesde)
        !FechaHasta = IIf((mProps.FechaHasta = 0), Now, mProps.FechaHasta)
        !HoraDesde = IIf((mProps.HoraDesde = 0), Now, mProps.HoraDesde)
        !HoraHasta = IIf((mProps.HoraHasta = 0), Now, mProps.HoraHasta)
        !FechaGrabado = Now
        .Update
        .Move 0             'Para actualizar
        fgrabado = !FechaGrabado
        mProps.IdPlantilla = !IdPlantilla
        .Close
    End With
    
    Me.Coneccion.CommitTrans       '======== Completa la transacción
    mProps.FechaGrabado = fgrabado
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    Grabar = True
    Exit Function
    
ErrTrap:
    Me.Coneccion.RollbackTrans  '======== Anula la transacción
    Err.Raise Err.Number, Err.Source, Err.Description
    Grabar = False
    Exit Function
End Function

Public Function Recuperar(ByVal Key As Variant) As Boolean
    Dim sql As String, rs As Recordset, s As String, valor As String
    On Error GoTo ErrTrap
            
    If VarType(Key) = vbString Then
        sql = "SELECT * FROM Plantilla_EI WHERE CodPlantilla='" & Key & "'"
    Else
        sql = "SELECT * FROM Plantilla_EI WHERE IdPlantilla=" & Key
    End If
    
    Set rs = AbrirTablaLectura(sql)
    If Not (rs.BOF And rs.EOF) Then
        With mProps
            .IdPlantilla = rs!IdPlantilla
            .CodPlantilla = rs!CodPlantilla
            .Descripcion = rs!Descripcion
            .Tipo = rs!Tipo
            .PrefijoNombreArchivo = IIf(IsNull(rs!PrefijoNombreArchivo), "", rs!PrefijoNombreArchivo)
            .BandRangoFechaHora = rs!BandRangoFechaHora
            .BandTipoFecha = rs!BandTipoFecha
            .BandIgnorarContabilidad = rs!BandIgnorarContabilidad
            .BandIgnorarDocAsignado = rs!BandIgnorarDocAsignado
            .BandActualizaCatalogos = rs!BandActualizaCatalogos
            .BandFiltrarxBodega = rs!BandFiltrarxBodega
            .ListaBodegas = IIf(IsNull(rs!ListaBodegas), "", rs!ListaBodegas)
            .ListaCatalogos = IIf(IsNull(rs!ListaCatalogos), "", rs!ListaCatalogos)
            
            If Not IsNull(rs.Fields("ListaTransacciones")) Then
                If Len(rs.Fields("ListaTransacciones")) > 0 Then
                    'Obtiene todo el contenido del campo 'valor'
                    ' Hay que usar el mátodo GetChunk porque es de tipo 'memo'
                    valor = ""
                    Do
                        s = rs.Fields("ListaTransacciones").GetChunk(1024)
                        valor = valor & s
                    Loop Until Len(s) < 1024
                End If
            Else
                valor = ""
            End If
            .ListaTransacciones = valor
            
            .BandActualizarCosto = rs!BandActualizarCosto
            .BandGuardarResultado = rs!BandGuardarResultado
            .BandValida = rs!BandValida
            .FechaDesde = IIf(IsNull(rs!FechaDesde), Date, rs!FechaDesde)
            .FechaHasta = IIf(IsNull(rs!FechaHasta), Date, rs!FechaHasta)
            .HoraDesde = IIf(IsNull(rs!HoraDesde), Date, rs!HoraDesde)
            .HoraHasta = IIf(IsNull(rs!HoraHasta), Date, rs!HoraHasta)
            .FechaGrabado = rs!FechaGrabado
        End With
        Recuperar = True
    Else
        With mProps
            .IdPlantilla = 0
            .CodPlantilla = ""
            .Descripcion = ""
            .Tipo = 0
            .PrefijoNombreArchivo = ""
            .BandRangoFechaHora = False
            .BandTipoFecha = False
            .BandIgnorarContabilidad = False
            .BandIgnorarDocAsignado = False
            .BandActualizaCatalogos = False
            .BandFiltrarxBodega = False
            .ListaBodegas = ""
            .ListaCatalogos = ""
            .ListaTransacciones = ""
            .BandActualizarCosto = False
            .BandGuardarResultado = False
            .BandValida = False
            .FechaDesde = Date
            .FechaHasta = Date
            .HoraDesde = Date
            .HoraHasta = Date
        End With
        Recuperar = False
    End If
    rs.Close
    mbooEsNuevo = False
    mbooModificado = False
    Set rs = Nothing
    Exit Function
    
ErrTrap:
    Err.Raise ERRNUM, "Plantilla.Recuperar", "No se puedo abrir la plantilla seleccionada"
End Function

Private Function AbrirTablaLectura(ByVal sql As String) As Recordset
    Dim rs As Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open sql, Me.Coneccion, adOpenStatic, adLockReadOnly
    Set AbrirTablaLectura = rs
    Set rs = Nothing
End Function

Private Function AbrirTablaEscritura(ByVal sql As String) As Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = IIf(GetSetting(APPNAME, App.Title, "ModoBDOpen", "Servidor") = "Servidor", adUseServer, adUseClient)
    rs.Open sql, Me.Coneccion, adOpenKeyset, adLockPessimistic
    Set AbrirTablaEscritura = rs
    Set rs = Nothing
End Function

Private Sub Class_Initialize()
    mbooEsNuevo = True
    mbooModificado = False
End Sub


