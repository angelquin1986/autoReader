VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVRecargoICE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'jeaa 25/07/2006
Option Explicit
Private Type T_PROPS
    IdICE As Long
    CodICE As String
    Descripcion As String
    BandValida As Boolean
    Porcentaje As Currency
    
    IdICESRI As Long
    CodICESRI As String
    
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
    sql = "SELECT CodICE FROM IVRecargoICE " & _
          " WHERE (CodICE)='" & mProps.CodICE & "' AND " & _
                 "(IdICE <>" & mProps.IdICE & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "IVICE_Grabar", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM IVRecargoICE WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM IVRecargoICE WHERE IdICE =" & mProps.IdICE
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "IVICE_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en Pos
    With rs
        !CodICE = mProps.CodICE
        !Descripcion = mProps.Descripcion
        !BandValida = mProps.BandValida
        
        If Not !Porcentaje = mProps.Porcentaje Then
            mProps.objEmpresa.GrabaAccionModulo !Porcentaje, mProps.Porcentaje, "PorcentajeICE", "MOD_ICE", "GN"
            ActulizaICEItems mProps.Porcentaje, !IdICE
        End If
       
        
        !Porcentaje = mProps.Porcentaje
        !IdICESRI = mProps.IdICESRI
        
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IdICE = !IdICE
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_ICE", "ICE Nuevo: " & _
                                                         mProps.CodICE, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_ICE", "ICE Modificado: " & _
                                                         mProps.CodICE, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVRECARGOICE
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT i.*, a.codice as CodICESRI  FROM IVRecargoICE i left join Anexo_ICE a on i.idicesri = a.id WHERE i.CodICE='" & key & "'"
    Else
        sql = "SELECT i.*, a.codice as CodICESRI  FROM IVRecargoICE i left join Anexo_ICE a on i.idicesri = a.id WHERE IdICE=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de ICE
    With rs
        mProps.IdICE = !IdICE
        mProps.CodICE = !CodICE
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        mProps.Porcentaje = !Porcentaje
        If Not IsNull(!IdICESRI) Then mProps.IdICESRI = !IdICESRI
        If Not IsNull(!CodICESRI) Then mProps.CodICESRI = !CodICESRI
        
        
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
    
    'Verifica si ya existe un item usando el mismo CodICE
    sql = "SELECT CodInventario FROM IvInventario" & _
          " WHERE IdICE=" & mProps.IdICE
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "ICE_Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close

    'Elimina el registro
    sql = "DELETE FROM IVRecargoICE WHERE IdICE=" & mProps.IdICE
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_ICE", "ICE Eliminado: " & _
                                                         mProps.CodICE, "DT"
                                                         
    Set rs = Nothing
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVRECARGOICE
End Sub

Friend Function Clone() As IVRecargoICE
    Dim nuevo As IVRecargoICE
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New IVRecargoICE
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdICE = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Let CodICE(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "ICE_CodICE", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodICE = value
    mbooModificado = True
End Property

Public Property Get CodICE() As String
    CodICE = Trim$(mProps.CodICE)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "ICE_Descripcion", MSGERR_INVALIDO
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
    If value < 0 Or value > 10 Then
        Err.Raise ERR_INVALIDO, "IVRecargoICE_Porcentaje", MSGERR_INVALIDO
    End If
    
    If mProps.Porcentaje <> value Then
        mProps.Porcentaje = value
        mbooModificado = True
    End If
End Property

Public Property Get Porcentaje() As Currency
    Porcentaje = mProps.Porcentaje
End Property


Public Property Let IdICESRI(ByVal value As Long)
    mProps.IdICESRI = value
    mbooModificado = True
End Property

Public Property Get IdICESRI() As Long
    IdICESRI = mProps.IdICESRI
End Property


Public Property Let CodICESRI(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodICESRI Then
        If Len(value) > 0 Then
            'Actualiza IdICESRI también
            sql = "SELECT Id FROM Anexo_ICE WHERE CodICE='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Anexo_CodICESRI", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdICESRI = rs!id
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdICESRI = 0
        End If
        
        mProps.CodICESRI = value
        mbooModificado = True
    End If
End Property

Public Property Get CodICESRI() As String
    CodICESRI = Trim$(mProps.CodICESRI)
End Property



Private Sub ActulizaICEItems(ByVal ICE As Double, IdICE As Long)
    Dim sql As String, rs As Recordset
    sql = "update IVinventario set porcentajeICE=" & ICE
    sql = sql & " where idice=" & IdICE
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
  
End Sub

