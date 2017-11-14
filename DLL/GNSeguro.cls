VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GNSeguro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    Id As Long
    CodSeguro As String
    Descripcion As String
    idrec As Long
    CodRec As String
    Porc As Currency
    BandValida As Boolean
    fechagrabado As Date
    objEmpresa As Empresa       'Referencia al objeto padre
    objEmpresa2 As Empresa2       'Referencia al objeto padre
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Public Sub Grabar()
    Dim sql As String, rs As Recordset
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    'Verifica si no repite el código
    sql = "SELECT CodSeguro FROM GNSeguro " & _
          "WHERE (CodSeguro='" & mProps.CodSeguro & "') AND " & _
                "(Id<>" & mProps.Id & ")"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "GNSeguro", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM GNSeguro WHERE 1=0"
        Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
        rs.AddNew
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM GNSeguro WHERE Id=" & mProps.Id
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "GNSeguro_Grabar", MSGERR_MODIFICADO
        End If
    End If
    'Graba el registro en Pos
    With rs
        !CodSeguro = mProps.CodSeguro
        !Descripcion = mProps.Descripcion
        !idrec = mProps.idrec
        !Porc = mProps.Porc
        !BandValida = mProps.BandValida
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.idrec = !idrec
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        gobjMain.EmpresaActual.GrabaGNLogAccion "NUE_SEG", "Seguro Nuevo: " & _
                                                         mProps.CodSeguro, "DT"
    Else
       gobjMain.EmpresaActual.GrabaGNLogAccion "MOD_SEG", "Seguro Modificado: " & _
                                                         mProps.CodSeguro, "DT"
    End If
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    'Genera evento Refresh del objeto Empresa
    gobjMain.EmpresaActual.RaiseRefreshEvent REFRESH_GNSEGURO
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    If VarType(key) = vbString Then
        sql = "SELECT gnseguro.*,ivr.codrecargo FROM GNSeguro"
        sql = sql & " Left Join IVRecargo ivr on ivr.idrecargo = gnseguro.idRec"
        sql = sql & " WHERE CodSeguro='" & key & "'"
    Else
        sql = "SELECT GNSeguro.*,ivr.codrecargo FROM GNSeguro "
        sql = sql & " Left Join IVRecargo ivr on ivr.idrecargo = gnseguro.idRec"
        sql = sql & " Where IdSeguro = " & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.Id = !Id
        mProps.CodSeguro = !CodSeguro
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!idrec) Then mProps.idrec = !idrec
        If Not IsNull(!CodRecargo) Then mProps.CodRec = !CodRecargo
        If Not IsNull(!Porc) Then mProps.Porc = !Porc
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!fechagrabado) Then mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    
    'Verifica si ya existe una transaccion usando el mismo
    sql = "SELECT IdSeguro FROM GNoferta WHERE IdSeguro=" & mProps.Id
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "GNseguro.Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close

    'Elimina el registro
    sql = "DELETE FROM GNSeguro WHERE Id=" & mProps.Id
    mProps.objEmpresa.Execute sql, True
    
    mProps.objEmpresa.GrabaGNLogAccion "ELI_SEG", "Seguro Eliminado: " & _
                                                         mProps.CodSeguro, "DT"
    mbooEsNuevo = True
    mbooModificado = True
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_GNSEGURO
End Sub

Friend Function Clone() As GNSeguro
    Dim nuevo As GNSeguro
    'Copia las propiedades al nuevo objeto
    Set nuevo = New GNSeguro
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.Id = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Friend Property Set Empresa2(objEmpresa2 As Empresa2)
    Set mProps.objEmpresa2 = objEmpresa2
End Property

Public Property Get Empresa2() As Empresa2
    Set Empresa2 = mProps.objEmpresa2
End Property

Public Property Let CodSeguro(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "GNSeguro_CodSeguro", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    mProps.CodSeguro = value
    mbooModificado = True
End Property

Public Property Get CodSeguro() As String
    CodSeguro = Trim$(mProps.CodSeguro)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "GNSeguro_Descripcion", MSGERR_INVALIDO
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

Public Property Let idrec(ByVal value As Long)
    mProps.idrec = value
    mbooModificado = True
End Property

Public Property Get idrec() As Long
    idrec = mProps.idrec
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
    With mProps
        .BandValida = True
    End With
    
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub
Public Property Get CodRec() As String
    CodRec = Trim$(mProps.CodRec)
End Property

Public Property Let CodRec(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodRec Then
        If Len(value) > 0 Then
            'Actualiza IdProvCli también
            sql = "SELECT IdRecargo FROM ivrecargo WHERE Codrecargo='" & value & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNseguro.Codrecargo", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.idrec = rs!IdRecargo
            End If
            rs.Close
        Else
            mProps.idrec = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodRec = value
    mbooModificado = True
End Property

Public Property Let Porc(ByVal value As Currency)
    mProps.Porc = value
    mbooModificado = True
End Property

Public Property Get Porc() As Currency
    Porc = mProps.Porc
End Property
