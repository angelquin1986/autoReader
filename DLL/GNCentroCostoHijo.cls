VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GNCentroCostoHijo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    idCentroDet As Long
    CodCentroDet As String
    idCentro As Long
    Descripcion As String
    Nombre As String
    FechaFinal As Variant
    transid As Long
    FechaGrabado As Date
    CodTrans As String
    NumTrans As Long
    CodCentro As String
    idFichaDet As Long
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean


Public Sub Grabar(ByVal CodTrans As String, ByVal NumTrans As Long)
    Dim sql As String, rs As Recordset, i As Integer
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
    If InStr(1, UCase(mProps.objEmpresa.GNOpcion.NombreEmpresa), "BOLIVAR") > 0 Then
        sql = "SELECT CodCentroDet FROM GNCentroCostoDetalle " & _
          "WHERE (CodCentroDet='" & mProps.CodCentroDet & "') AND " & _
          "(IdFichadet=" & mProps.idFichaDet & ")"
    Else
        sql = "SELECT CodCentroDet FROM GNCentroCostoDetalle " & _
          "WHERE (CodCentroDet='" & mProps.CodCentroDet & "') AND " & _
          "(IdCentrodet<>" & mProps.idCentroDet & ")"
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "GNCentroCostoDetalle", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM GNCentroCostoDetalle WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM GNCentroCostoDetalle WHERE IdCentroDET=" & mProps.idCentroDet
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "GNCentroCostoDetalle_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en GNCentroCosto
    With rs
        !CodCentroDet = mProps.CodCentroDet
        !DescripcionDet = mProps.Descripcion
        !idCentro = mProps.idCentro
        !Orden = mProps.transid
        !idFichaDet = mProps.idFichaDet
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.FechaGrabado = !FechaGrabado
        mProps.idCentroDet = !idCentroDet
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_CCDET", "Centro de Costo Det Nuevo: " & _
                                                         mProps.CodCentroDet, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_CCDET", "Centro de Costo Det Modificado: " & _
                                                         mProps.CodCentroDet, "DT"
    End If
    
    sql = "UPDATE gncomprobante "
    sql = sql & "SET  IDCentroHijo = " & mProps.idCentroDet
    sql = sql & " WHERE CodTrans = '" & CodTrans & "'"
    sql = sql & " AND numTrans = " & NumTrans
    mProps.objEmpresa.EjecutarSQL sql, 1
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_CENTROCOSTOHIJO
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset, i As Integer
    Dim Trans As String
    
    If Len(mProps.objEmpresa.GNOpcion.ObtenerValor("ProduccionTransOP")) > 0 Then
        Trans = mProps.objEmpresa.GNOpcion.ObtenerValor("ProduccionTransOP")
    End If
    
    
    If VarType(key) = vbString Then
        sql = "SELECT GCCD.*,GCC.CODCENTRO,gn.codtrans,gn.numtrans FROM GNCentroCostoDetalle gccd "
        sql = sql & " INNER  JOIN GNCentroCosto gcc ON gcc.idcentro = gccd.idcentro "
        sql = sql & " INNER  JOIN GNComprobante gn ON gn.idcentrohijo = gccd.idcentrodet "
        
        sql = sql & " WHERE CodCentroDet='" & key & "'"
    Else
        sql = "SELECT * FROM GNCentroCostoDetalle WHERE IdCentro=" & key
    End If
    If Len(Trans) > 0 Then
        sql = sql & " And gn.codtrans = '" & Trans & "'"
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'Recupera los datos de Usuario
    With rs
        mProps.idCentroDet = !idCentroDet
        mProps.idCentro = !idCentro 'PADRE
        mProps.CodCentroDet = !CodCentroDet
        If Not IsNull(!DescripcionDet) Then mProps.Descripcion = !DescripcionDet
        If Not IsNull(!Orden) Then mProps.transid = !Orden
        If Not IsNull(!CodCentro) Then mProps.CodCentro = !CodCentro
        If Not IsNull(!CodTrans) Then mProps.CodTrans = !CodTrans
        If Not IsNull(!NumTrans) Then mProps.NumTrans = !NumTrans
        mProps.idFichaDet = !idFichaDet
        
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    
    'Verifica si existe un comprobante que usa el codigo de centro de costo
    sql = "SELECT TransID FROM GNComprobante WHERE IdCentroHijo=" & mProps.idCentroDet
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        rs.Close
        Set rs = Nothing
        Err.Raise ERR_NOELIMINABLE, "GNCentroCostoDetalle.Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close
    Set rs = Nothing
    
    'Elimina el Centro de costo
    sql = "DELETE FROM GNCentroCostoDetalle WHERE IdCentrodet=" & mProps.idCentroDet
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_CCDet", "Centro de Costo Detalle Eliminado: " & _
                                                         mProps.CodCentroDet, "DT"

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_CENTROCOSTOHIJO
End Sub

Friend Function Clone() As GNCentroCostoHijo
    Dim nuevo As GNCentroCostoHijo
    
    Set nuevo = New GNCentroCostoHijo
    'Copia las propiedades al nuevo objeto
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.idCentroDet = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property





Public Property Let CodCentroDet(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNCentroCosto_CodCentrodet", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodCentroDet = value
    mbooModificado = True
End Property

Public Property Get CodCentroDet() As String
    CodCentroDet = Trim$(mProps.CodCentroDet)
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

Public Property Let transid(ByVal value As Long)
    Dim sql As String, rs As Recordset
    
    
    
    mProps.transid = value
    mbooModificado = True
End Property


Public Property Get transid() As Long
    transid = mProps.transid
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


Private Sub Class_Initialize()

    'Valores predeterminados
'    With mProps
'        .FechaInicio = Date
'        .FechaFinal = Null
'    End With
    
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub


Public Sub GeneraCodCentro()
    Dim sql As String, pre As String, num As Integer, rs As Recordset
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
Public Property Let CodTrans(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodTrans Then
        If Len(value) > 0 Then
            'Actualiza IdCliente también
            sql = "SELECT Transid FROM GNComprobante WHERE CodTrans='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNTransaccion.CodTrans", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.transid = rs!transid
            End If
            rs.Close
        Else
            mProps.transid = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodTrans = value
    mbooModificado = True
End Property

Public Property Get CodTrans() As String
    CodTrans = Trim$(mProps.CodTrans)
End Property
Public Property Get NumTrans() As Long
    NumTrans = Trim$(mProps.NumTrans)
End Property

Public Property Let CodCentro(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodCentro Then
        If Len(value) > 0 Then
            'Actualiza IdCliente también
            sql = "SELECT idcentro FROM GNCentroCosto WHERE CodCentro='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNCentro.Codcentro", MSGERR_NOHAYCODIGO
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

Public Property Get CodCentro() As String
    CodCentro = Trim$(mProps.CodCentro)
End Property

Public Property Get idCentroDet() As Long
    idCentroDet = mProps.idCentroDet
End Property


Public Property Get idFichaDet() As Long
    idFichaDet = mProps.idFichaDet
End Property

Public Property Let idFichaDet(ByVal value As Long)
    mProps.idFichaDet = value
    mbooModificado = True
End Property

Public Function RecuperaFechaFicha() As Date
Dim sql As String
Dim fecha As Date
Dim rs As Recordset
On Error GoTo CapturaError
    sql = "Select fechaIngreso from FichaDetalle fd "
    sql = sql & "Inner Join GNCentroCostodetalle ccd "
    sql = sql & " ON ccd.idfichadet=fd.id"
    sql = sql & " Where ccd.idcentrodet= " & Me.idCentroDet
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    If Not rs.EOF Then
        fecha = rs!FechaIngreso
    
    End If
    RecuperaFechaFicha = fecha
    Set rs = Nothing
    Exit Function
CapturaError:
    MsgBox Err.Description
    Set rs = Nothing
    Exit Function
End Function

Friend Sub recuperarxId(key As Long)
    Dim sql As String, rs As Recordset, i As Integer
    
    sql = "SELECT * FROM GNCentroCostoDetalle WHERE IdCentroDet=" & key
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'Recupera los datos de Usuario
    With rs
        mProps.idCentroDet = !idCentroDet
        mProps.idCentro = !idCentro 'PADRE
        mProps.CodCentroDet = !CodCentroDet
        If Not IsNull(!DescripcionDet) Then mProps.Descripcion = !DescripcionDet
        If Not IsNull(!Orden) Then mProps.transid = !Orden
     '   If Not IsNull(!CodCentro) Then mProps.CodCentro = !CodCentro
       ' I'f Not IsNull(!CodTrans) Then mProps.CodTrans = !CodTrans
       ' If Not IsNull(!NumTrans) Then mProps.NumTrans = !NumTrans
        mProps.idFichaDet = !idFichaDet
        
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub












