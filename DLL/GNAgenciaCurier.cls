VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GNAgenciaCurier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdAgeCurier As Long
    CodAgeCurier As String
    Descripcion As String
    Direccion As String
    Telefono As String
    CodPostal As String
    IdPais As Long
    CodPais As String
    IdProvincia As Long
    CodProvincia As String
    IdCanton As Long
    CodCanton As String
    IDParroquia As Long
    CodParroquia As String
    BandValida As Boolean
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
    sql = "SELECT CodAgeCurier FROM GNAgenciaCurier " & _
          "WHERE (CodAgeCurier='" & mProps.CodAgeCurier & "') AND " & _
                "(IdAgeCurier<>" & mProps.IdAgeCurier & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "GNAgenciaCurier", MSGERR_REPITECODIGO
    End If
    rs.Close
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM GNAgenciaCurier WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM GNAgenciaCurier WHERE IdAgeCurier=" & mProps.IdAgeCurier
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "PCPais.Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro
    With rs
        If Not !CodAgeCurier = mProps.CodAgeCurier Then
            mProps.objEmpresa.GrabaAccionModulo !CodAgeCurier, mProps.CodAgeCurier, "Codigo Agencia", "MOD_SUC", "GN"
        End If
        !CodAgeCurier = mProps.CodAgeCurier
        If Not !Descripcion = mProps.Descripcion Then
            mProps.objEmpresa.GrabaAccionModulo !Descripcion, mProps.Descripcion, "Descripcion Agencia", "MOD_SUC", "GN"
        End If
        !Descripcion = mProps.Descripcion
        !Direccion = mProps.Direccion
        !Telefono = mProps.Telefono
        !CodPostal = mProps.CodPostal
        !IdPais = mProps.IdPais
        !IdProvincia = mProps.IdProvincia
        !IdCanton = mProps.IdCanton
        !IDParroquia = mProps.IDParroquia
        !BandValida = mProps.BandValida
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IdAgeCurier = !IdAgeCurier
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_GNAgenciaCurier", "Agencia Nuevo: " & _
                                                         mProps.CodAgeCurier, "GN"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_GNAgenciaCurier", "Agencia Modificado: " & _
                                                         mProps.CodAgeCurier, "GN"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_GNAGENCIACURIER
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT gna.*,pcpais.codpais,pcprov.codprovincia,pcc.codcanton,pcparr.codparroquia FROM GNAgenciaCurier gna"
        sql = sql & " Left Join pcpais on pcpais.idpais = gna.idpais"
        sql = sql & " Left Join pcprovincia pcprov on pcprov.idprovincia = gna.idprovincia"
        sql = sql & " Left Join pccanton pcc on pcc.idcanton = gna.idcanton"
        sql = sql & " Left Join pcparroquia pcparr on pcparr.idparroquia = gna.idParroquia"
        sql = sql & " WHERE CodAgeCurier='" & key & "'"
    Else
        sql = "SELECT gna.*,pcpais.codpais,pcprov.codprovincia,pcc.codcanton,pcparr.codparroquia FROM GNAgenciaCurier gna"
        sql = sql & " Left Join pcpais on pcpais.idpais = gna.idpais"
        sql = sql & " Left Join pcprovincia pcprov on pcprov.idprovincia = gna.idprovincia"
        sql = sql & " Left Join pccanton pcc on pcc.idcanton = gna.idcanton"
        sql = sql & " Left Join pcparroquia pcparr on pcparr.idparroquia = gna.idParroquia"
        sql = sql & " WHERE IdAgeCurier=" & key
    End If

    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    rs.MoveFirst   'para que se coloque al inicio jeaa
    'Recupera los datos
    With rs
        mProps.IdAgeCurier = !IdAgeCurier
        If Not IsNull(!CodAgeCurier) Then mProps.CodAgeCurier = !CodAgeCurier
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!Direccion) Then mProps.Direccion = !Direccion
        If Not IsNull(!Telefono) Then mProps.Telefono = !Telefono
        If Not IsNull(!CodPostal) Then mProps.CodPostal = !CodPostal
        If Not IsNull(!IdPais) Then mProps.IdPais = !IdPais
        If Not IsNull(!CodPais) Then mProps.CodPais = !CodPais
        If Not IsNull(!IdProvincia) Then mProps.IdProvincia = !IdProvincia
        If Not IsNull(!CodProvincia) Then mProps.CodProvincia = !CodProvincia
        If Not IsNull(!IdCanton) Then mProps.IdCanton = !IdCanton
        If Not IsNull(!CodCanton) Then mProps.CodCanton = !CodCanton
        If Not IsNull(!IDParroquia) Then mProps.IDParroquia = !IDParroquia
        If Not IsNull(!CodParroquia) Then mProps.CodParroquia = !CodParroquia
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        
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
    sql = "SELECT gof.IdAgeCurier,g.codtrans,g.numtrans FROM GNOferta gof Inner join gncomprobante g on g.transid = gof.transid WHERE gof.IdAgeCurier=" & mProps.IdAgeCurier
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "GNAgenciaCurier.Eliminar", MSGERR_NOELIMINABLE & " Esta relacionado con: " & rs!CodTrans & rs!NumTrans
    End If
    rs.Close

    'Elimina el registro
    sql = "DELETE FROM GNAgenciaCurier WHERE IdAgeCurier=" & mProps.IdAgeCurier
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_GNAgenciaCurier", "Agencia Eliminado: " & _
                                                         mProps.CodAgeCurier, "GN"

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_GNAGENCIACURIER
End Sub

Friend Function Clone() As GNAgenciaCurier
    Dim nuevo As GNAgenciaCurier

    'Copia las propiedades al nuevo objeto
    Set nuevo = New GNAgenciaCurier
    nuevo.Props = mProps

    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdAgeCurier = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Let CodAgeCurier(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNAgenciaCurier_CodPais", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodAgeCurier = value
    mbooModificado = True
End Property

Public Property Get CodAgeCurier() As String
    CodAgeCurier = Trim$(mProps.CodAgeCurier)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "Pais_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
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

Public Property Let CodProvincia(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodProvincia Then
        If Len(value) > 0 Then
            'Actualiza IdGarante también
            sql = "SELECT IdProvincia FROM PCProvincia WHERE CodProvincia='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodProvincia", MSGERR_NOHAYCODIGO
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
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_Codcanton", MSGERR_NOHAYCODIGO
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
                Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodParroquia", MSGERR_NOHAYCODIGO
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
    mProps.CodParroquia = value
    mbooModificado = True
End Property

Public Property Get CodParroquia() As String
    CodParroquia = Trim$(mProps.CodParroquia)
End Property



Public Property Let CodPais(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodProvincia Then
        If Len(value) > 0 Then
            'Actualiza IdGarante también
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

Public Property Let IdPais(ByVal value As Long)
    mProps.IdPais = value
    mbooModificado = True
End Property

Public Property Get IdPais() As Long
    IdPais = mProps.IdPais
End Property

Public Property Let CodPostal(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNAgencia_CodPostal", MSGERR_INVALIDO
    End If
    mProps.CodPostal = value
    mbooModificado = True
End Property

Public Property Get CodPostal() As String
    CodPostal = Trim$(mProps.CodPostal)
End Property

Public Property Let Direccion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "GNAgencia_Direccion", MSGERR_INVALIDO
    End If
    mProps.Direccion = value
    mbooModificado = True
End Property

Public Property Get Direccion() As String
    Direccion = Trim$(mProps.Direccion)
End Property

Public Property Let Telefono(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNAgencia_Telefono", MSGERR_INVALIDO
    End If
    mProps.Telefono = value
    mbooModificado = True
End Property

Public Property Get Telefono() As String
    Telefono = Trim$(mProps.Telefono)
End Property

