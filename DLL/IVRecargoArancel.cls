VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVRecargoArancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'jeaa 25/07/2006
Option Explicit
Private Type T_PROPS
    IdArancel As Long
    CodArancel As String
    Descripcion As String
    BandValida As Boolean
    Porcentaje As Currency
    RecarPorcentaje As Currency
    ValorFijo As Currency
    Peso As Currency
    Partida As String
    MontoMaximo As Currency
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
    
    'Verifica si no repite el código
    sql = "SELECT CodARANCEL FROM IVRecargoARANCEL " & _
          " WHERE (CodARANCEL)='" & mProps.CodArancel & "' AND " & _
                 "(IdARANCEL <>" & mProps.IdArancel & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "IVARANCEL_Grabar", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM IVRecargoARANCEL WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM IVRecargoARANCEL WHERE IdARANCEL =" & mProps.IdArancel
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "IVARANCEL_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en Pos
    With rs
        !CodArancel = mProps.CodArancel
        !Descripcion = mProps.Descripcion
        !BandValida = mProps.BandValida
        
        If Not !Porcentaje = mProps.Porcentaje Then
            mProps.objEmpresa.GrabaAccionModulo !Porcentaje, mProps.Porcentaje, "Porcentajeara", "MOD_ARA", "GN"
        End If
       
        
        !Porcentaje = mProps.Porcentaje
        
        If Not !RecarPorcentaje = mProps.RecarPorcentaje Then
            mProps.objEmpresa.GrabaAccionModulo !RecarPorcentaje, mProps.RecarPorcentaje, "RecarPorcentaje", "MOD_ARA", "GN"
        End If
        
        !RecarPorcentaje = mProps.RecarPorcentaje
        !ValorFijo = mProps.ValorFijo
        !MontoMaximo = mProps.MontoMaximo
'        !Peso = mProps.Peso
'        !Partida = mProps.Partida
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IdArancel = !IdArancel
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_ARANCEL", "ARANCEL Nuevo: " & _
                                                         mProps.CodArancel, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_ARANCEL", "ARANCEL Modificado: " & _
                                                         mProps.CodArancel, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVRECARGOARANCEL
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM IVRecargoARANCEL WHERE CodARANCEL='" & key & "'"
    Else
        sql = "SELECT * FROM IVRecargoARANCEL WHERE IdARANCEL=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de ARANCEL
    With rs
        mProps.IdArancel = !IdArancel
        mProps.CodArancel = !CodArancel
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        mProps.Porcentaje = !Porcentaje
        mProps.RecarPorcentaje = !RecarPorcentaje
        mProps.ValorFijo = !ValorFijo
        mProps.MontoMaximo = !MontoMaximo
 '       mProps.Peso = !Peso
'        mProps.Partida = !Partida
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String
    Dim rs As Recordset
    
    'Verifica si ya existe un item usando el mismo CodARANCEL
    sql = "SELECT CodInventario FROM IvInventario " & _
          " WHERE IdARANCEL=" & mProps.IdArancel
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "ARANCEL_Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close

    'Elimina el registro
    sql = "DELETE FROM IVRecargoARANCEL WHERE IdARANCEL=" & mProps.IdArancel
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_ARANCEL", "ARANCEL Eliminado: " & _
                                                         mProps.CodArancel, "DT"
                                                         
    Set rs = Nothing
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVRECARGOARANCEL
End Sub

Friend Function Clone() As IVRecargoArancel
    Dim nuevo As IVRecargoArancel
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New IVRecargoArancel
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdArancel = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Let CodArancel(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "ARANCEL_CodARANCEL", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodArancel = value
    mbooModificado = True
End Property

Public Property Get CodArancel() As String
    CodArancel = Trim$(mProps.CodArancel)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "ARANCEL_Descripcion", MSGERR_INVALIDO
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
    mbooEsNuevo = True
    mbooModificado = False
    
    'Valores iniciales
    mProps.BandValida = True
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Property Let Porcentaje(ByVal value As Currency)
    If value < 0 Or value > 1 Then
        Err.Raise ERR_INVALIDO, "IVRecargoARANCEL_Porcentaje", MSGERR_INVALIDO
    End If
    
    If mProps.Porcentaje <> value Then
        mProps.Porcentaje = value
        mbooModificado = True
    End If
End Property

Public Property Get Porcentaje() As Currency
    Porcentaje = mProps.Porcentaje
End Property

Public Property Let RecarPorcentaje(ByVal value As Currency)
    If value < 0 Or value > 1 Then
        Err.Raise ERR_INVALIDO, "IVRecargoARANCEL_RecarPorcentaje", MSGERR_INVALIDO
    End If
    
    If mProps.RecarPorcentaje <> value Then
        mProps.RecarPorcentaje = value
        mbooModificado = True
    End If
End Property

Public Property Get RecarPorcentaje() As Currency
    RecarPorcentaje = mProps.RecarPorcentaje
End Property

Public Property Let ValorFijo(ByVal value As Currency)
'    If value < 0 Or value > 1 Then
'        Err.Raise ERR_INVALIDO, "IVRecargoARANCEL_ValorFijo", MSGERR_INVALIDO
'    End If
    
    If mProps.ValorFijo <> value Then
        mProps.ValorFijo = value
        mbooModificado = True
    End If
End Property

Public Property Get ValorFijo() As Currency
    ValorFijo = mProps.ValorFijo
End Property


Public Property Let Peso(ByVal value As Currency)
'    If mProps.Peso <> value Then
'        mProps.Peso = value
'        mbooModificado = True
'    End If
End Property

Public Property Get Peso() As Currency
'    Peso = mProps.Peso
End Property


Public Property Let MontoMaximo(ByVal value As Currency)
    If mProps.MontoMaximo <> value Then
        mProps.MontoMaximo = value
        mbooModificado = True
    End If
End Property

Public Property Get MontoMaximo() As Currency
    MontoMaximo = mProps.MontoMaximo
End Property


Public Property Let Partida(ByVal value As String)
'    If Len(value) > 50 Then
'        Err.Raise ERR_INVALIDO, "ARANCEL_Partida", MSGERR_INVALIDO
'    End If
'    mProps.Partida = value
'    mbooModificado = True
End Property

Public Property Get Partida() As String
'    Partida = Trim$(mProps.Partida)
End Property

