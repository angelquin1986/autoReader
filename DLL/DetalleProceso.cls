VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DetalleProceso"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"GNComprobante"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type T_PROP
    id As Long
    IDkp As Long
    CodProceso As String
    fecha As Date
    HoraInicio As Date
    HoraFinal As Date
    HoraGracia As Date
    NumHoras As Single
    CodCentro As String
    CodCentroHijo As String
    DesCentroHijo As String
    Cliente As String
    CodResponsable As String
    Nota As String
    objEmpresa As Empresa
End Type
Private mProps As T_PROP
Public Sub grabar()
    Dim sql As String, rs As Recordset
        sql = "SELECT * FROM DetalleProceso WHERE 1=0"
    Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !IDkp = Me.IDkp
        !fecha = mProps.fecha
        !HoraInicio = mProps.HoraInicio
        !HoraFinal = mProps.HoraFinal
        '!NumHoras = NumHoras 'CAMBIO ULTIMO PARA MAKAHAIV
        !NumHoras = mProps.NumHoras
        !HoraGracia = mProps.HoraGracia
        !Nota = mProps.Nota
        .Update
        
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.id = !id             'Guarda el Id(Autonumerico)
        .Close
    End With
    Set rs = Nothing
'    mProps.objEmpresa.RaiseRefreshEvent REFRESH_DETALLEPROCESO
End Sub

Friend Sub recuperar(rs As Recordset)
    With rs
        If Not .EOF Then
            mProps.id = !id
            If Not IsNull(!IDkp) Then mProps.IDkp = !IDkp
            If Not IsNull(!fecha) Then mProps.fecha = !fecha
            If Not IsNull(!HoraInicio) Then mProps.HoraInicio = !HoraInicio
            If Not IsNull(!HoraFinal) Then mProps.HoraFinal = !HoraFinal
            If Not IsNull(!NumHoras) Then mProps.HoraFinal = !HoraFinal
            If Not IsNull(!HoraGracia) Then mProps.HoraGracia = !HoraGracia
            If Not IsNull(!Nota) Then mProps.Nota = !Nota
        
        End If
    End With
End Sub

Friend Function Clone() As DetalleProceso
    Dim nuevo As DetalleProceso
    
    Set nuevo = New DetalleProceso
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(Value As T_PROP)          'Para el método Clone
    mProps = Value
    mProps.id = 0
    'mProps.objIVKProceso.GNComprobante.SetModificado
End Property

'Friend Property Set GNComprobante(value As GNComprobante)
'    Set mProps.objGNComprobante = value
'End Property
'
'Public Property Get GNComprobante() As GNComprobante
'    Set GNComprobante = mProps.objGNComprobante
'End Property
'Friend Property Set IVKProceso(value As IVKProceso)
'    Set mProps.IVKProceso = value
'End Property
'
'Public Property Get IVKProceso() As IVKProceso
'    Set IVKProceso = mProps.objIVKProceso
'End Property
Friend Property Set objEmpresa(Value As Empresa)
    Set mProps.objEmpresa = Value
End Property
Public Property Get objEmpresa() As Empresa
    Set objEmpresa = mProps.objEmpresa
End Property
'AUC cambiado para el reproceso de recetas 19/10/06
Public Property Get id() As Long
    id = mProps.id
End Property
Public Property Get IDkp() As Long
    IDkp = mProps.IDkp
End Property

Public Property Let IDkp(ByVal Value As Long)
    mProps.IDkp = Value
End Property
Public Property Let fecha(Value As Date)
    If mProps.fecha <> Value Then
        mProps.fecha = Value
    End If
End Property

Public Property Get fecha() As Date
    fecha = mProps.fecha
End Property
Public Property Let HoraInicio(Value As Date)
    If mProps.HoraInicio <> Value Then
        mProps.HoraInicio = Value
    End If
End Property

Public Property Get HoraInicio() As Date
    HoraInicio = mProps.HoraInicio
End Property
Public Property Let HoraFinal(Value As Date)
    If mProps.HoraFinal <> Value Then
        mProps.HoraFinal = Value
    End If
End Property

Public Property Get HoraFinal() As Date
    HoraFinal = mProps.HoraFinal
End Property

Private Sub Class_Initialize()
  Set objEmpresa = gobjMain.EmpresaActual
    HoraInicio = Time
    HoraFinal = Time
End Sub
Public Property Get NumHoras() As Single
    NumHoras = CalculaNumHora
End Property

Public Property Let NumHoras(ByVal Value As Single)
    mProps.NumHoras = Value
End Property

Private Function CalculaNumHora() As Single
    Dim n As Single
    
    'Calcula la diferencia en horas
    n = DateDiff("n", HoraInicio, (HoraFinal - HoraGracia)) / 60
    CalculaNumHora = n
End Function

Public Property Let HoraGracia(Value As Date)
    If mProps.HoraGracia <> Value Then
        mProps.HoraGracia = Value
    End If
End Property

Public Property Get HoraGracia() As Date
    HoraGracia = mProps.HoraGracia
End Property
Friend Sub RecuperarTiempos(id As Long)
Dim sql As String, rs As Recordset
On Error GoTo CapturaError
    sql = "SELECT dp.*,p.codproceso,p.descripcion as DescProceso,kp.idproceso, kp.descripcion,pc.nombre,"
    sql = sql & " gc.codcentro,gcd.codcentrodet,gcd.descripciondet,gnr.codresponsable,dp.nota "
    sql = sql & "  FROM DetalleProceso dp INNER JOIN  IVKProceso kp  left JOIN GNResponsable gnr on gnr.idresponsable= kp.idresponsable  "
    sql = sql & " INNER JOIN IVProceso P ON p.idproceso = kp.idproceso "
    sql = sql & " INNER JOIN gncentrocostoDetalle gcd "
    sql = sql & " inner join gncentrocosto gc "
    sql = sql & " inner join pcprovcli pc on pc.idprovcli = gc.idcliente"
    sql = sql & " on gc.idcentro = gcd.idcentro"
    sql = sql & " on gcd.idcentrodet = kp.idcentrodet"
    sql = sql & " ON DP.IDKP = KP.IDKP "
    sql = sql & " Where  dp.id = " & id
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
            mProps.id = !id
            mProps.IDkp = !IDkp
            If Not IsNull(!fecha) Then mProps.fecha = !fecha
            If Not IsNull(!HoraInicio) Then mProps.HoraInicio = !HoraInicio
            If Not IsNull(!HoraFinal) Then mProps.HoraFinal = !HoraFinal
            If Not IsNull(!NumHoras) Then mProps.NumHoras = !NumHoras
            If Not IsNull(!HoraGracia) Then mProps.HoraGracia = !HoraGracia
            If Not IsNull(!CodProceso) Then mProps.CodProceso = !CodProceso
            If Not IsNull(!CodCentro) Then mProps.CodCentro = !CodCentro
            If Not IsNull(!CodCentroDet) Then mProps.CodCentroHijo = !CodCentroDet
            If Not IsNull(!DescripcionDet) Then mProps.DesCentroHijo = !DescripcionDet
            If Not IsNull(!Nombre) Then mProps.Cliente = !Nombre
            If Not IsNull(!CodResponsable) Then mProps.CodResponsable = !CodResponsable
            If Not IsNull(!Nota) Then mProps.Nota = !Nota
            
            
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
Public Property Get CodProceso() As String
    CodProceso = mProps.CodProceso
End Property
Public Property Get Cliente() As String
    Cliente = mProps.Cliente
End Property

Public Property Get CodCentroHijo() As String
    CodCentroHijo = mProps.CodCentroHijo
End Property

Public Property Get CodCentro() As String
    CodCentro = mProps.CodCentro
End Property

Public Property Get DesCentroHijo() As String
    DesCentroHijo = mProps.DesCentroHijo
End Property
Public Property Get CodResponsable() As String
    CodResponsable = mProps.CodResponsable
End Property

Public Sub GrabarDetalle()
    Dim sql As String, rs As Recordset
        sql = "SELECT * FROM DetalleProceso WHERE id = " & mProps.id

    Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
    With rs
        

        
        !IDkp = Me.IDkp
        !fecha = mProps.fecha
        !HoraInicio = mProps.HoraInicio
        !HoraFinal = mProps.HoraFinal
        !NumHoras = mProps.NumHoras
        !HoraGracia = mProps.HoraGracia
        .Update
        
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.id = !id             'Guarda el Id(Autonumerico)
        .Close
    End With
    Set rs = Nothing
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PROTIEMPOS
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    
    'Elimina el Contrato de costo
    sql = "DELETE FROM DetalleProceso  WHERE Id=" & mProps.id
    mProps.objEmpresa.Execute sql, True
    
    mProps.objEmpresa.GrabaGNLogAccion "ELI_DETProceso", "Proceso Detalle Eliminado: " & _
                                                         mProps.CodCentroHijo, "DT"
   ' mbooEsNuevo = True
    'mbooModificado = True
    
     'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PROTIEMPOS
End Sub

Public Property Get Nota() As String
    Nota = mProps.Nota
End Property
Public Property Let Nota(ByVal Value As String)
    If Len(Value) > 120 Then
        Err.Raise ERR_INVALIDO, "ivkprocesodetalle.Nota", MSGERR_INVALIDO
    End If
    mProps.Nota = Value
'    mbooModificado = True
End Property
