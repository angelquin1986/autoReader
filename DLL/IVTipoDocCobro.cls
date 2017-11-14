VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVTipoDocCobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
    Option Explicit
    
    
    Private Type T_PROPS
        IDTipoDocCobro As Long
        CodTipoDocCobro As String
        Descripcion As String
        IdForma As Long
        CodForma As String
        BandValida As Boolean
        IdBanco As Long
        CodBanco As String
        
        FechaGrabado As Date
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
        sql = "SELECT CodTipoDocCobro FROM IVTipoDocCobro " & _
              "WHERE (CodTipoDocCobro='" & mProps.CodTipoDocCobro & "') AND " & _
                    "(IdTipoDocCobro<>" & mProps.IDTipoDocCobro & ")"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then      'Si ya existe el codigo
            Err.Raise ERR_REPITECODIGO, "IVTipoDocCobro", MSGERR_REPITECODIGO
        End If
        rs.Close
        
        'Si es nuevo
        If mbooEsNuevo Then
            sql = "SELECT * FROM IVTipoDocCobro WHERE 1=0"
            Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
            rs.AddNew
        
        'Si es modificada
        Else
            'Recupera el registro correspondiente para modificar
            sql = "SELECT * FROM IVTipoDocCobro WHERE IdTipoDocCobro=" & mProps.IDTipoDocCobro
            Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
    #If DAOLIB Then
            rs.Edit
    #End If
            
            'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
            If rs!FechaGrabado <> mProps.FechaGrabado Then
                Err.Raise ERR_MODIFICADO, "IVTipoDocCobro_Grabar", MSGERR_MODIFICADO
            End If
        End If
    
        'Graba el registro en Pos
        With rs
            !CodTipoDocCobro = mProps.CodTipoDocCobro
            !Descripcion = mProps.Descripcion
            !IdForma = mProps.IdForma
            !IdBanco = mProps.IdBanco
            !BandValida = mProps.BandValida
            !FechaGrabado = Now
            .Update
    #If DAOLIB Then
            rs.Bookmark = rs.LastModified
    #End If
            .Move 0             'Para actualizar
            mProps.FechaGrabado = !FechaGrabado
            mProps.IDTipoDocCobro = !IDTipoDocCobro
            .Close
        End With
        
        'Graba un registro para auditoria   *** ANGEL 20/mar/02
        If mbooEsNuevo Then
            mProps.objEmpresa.GrabaGNLogAccion "NUE_TipoDocCobro", "TipoDocCobro Nuevo: " & _
                                                             mProps.CodTipoDocCobro, "DT"
        Else
            mProps.objEmpresa.GrabaGNLogAccion "MOD_TipoDocCobro", "TipoDocCobro Modificado: " & _
                                                             mProps.CodTipoDocCobro, "DT"
        End If
        
        Set rs = Nothing
        mbooEsNuevo = False
        mbooModificado = False
        
        'Genera evento Refresh del objeto Empresa
        mProps.objEmpresa.RaiseRefreshEvent REFRESH_TIPODOCCOBRO
    End Sub
    
    Friend Sub Recuperar(key As Variant)
        Dim sql As String, rs As Recordset
        
        If VarType(key) = vbString Then
            ''sql = "SELECT * FROM IVTipoDocCobro WHERE CodTipoDocCobro='" & key & "'"
            sql = "SELECT IVTipoDocCobro.*, CODFORMA"
            sql = sql & " FROM IVTipoDocCobro INNER JOIN TSFORMACOBROPAGO"
            sql = sql & "  ON IVTipoDocCobro.IDFORMA = TSFORMACOBROPAGO.IDFORMA"
            sql = sql & "  WHERE CodTipoDocCobro='" & key & "'"
        Else
            'sql = "SELECT * FROM IVTipoDocCobro WHERE IdTipoDocCobro=" & key
            sql = "SELECT IVTipoDocCobro.*, CODFORMA"
            sql = sql & " FROM IVTipoDocCobro INNER JOIN TSFORMACOBROPAGO"
            sql = sql & "  ON IVTipoDocCobro.IDFORMA = TSFORMACOBROPAGO.IDFORMA"
            sql = sql & "  WHERE IdTipoDocCobro=" & key
            
        End If
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        
        'Recupera los datos de FormaCopro
        With rs
            mProps.IDTipoDocCobro = !IDTipoDocCobro
            mProps.CodTipoDocCobro = !CodTipoDocCobro
            If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
            If Not IsNull(!IdForma) Then mProps.IdForma = !IdForma
            If Not IsNull(!CodForma) Then mProps.CodForma = !CodForma
            If Not IsNull(!IdBanco) Then mProps.IdBanco = !IdBanco
            If Not IsNull(!CodBanco) Then mProps.CodBanco = !CodBanco
            
            If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
           
            mProps.FechaGrabado = !FechaGrabado
            .Close
        End With
        
        Set rs = Nothing
        mbooEsNuevo = False
        mbooModificado = False
    End Sub
    
    Public Sub Eliminar()
        Dim sql As String, rs As Recordset
    
        'Verifica si existe Kardex del TipoDocCobro
        sql = "SELECT IdTipoDocCobro FROM TSKardex WHERE IdTipoDocCobro=" & mProps.IDTipoDocCobro
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If Not rs.EOF Then
            Err.Raise ERR_NOELIMINABLE, "IVTipoDocCobro.Eliminar", MSGERR_NOELIMINABLE
        End If
        Set rs = Nothing
    
        'Elimina el TipoDocCobro
        sql = "DELETE FROM IVTipoDocCobro WHERE CodTipoDocCobro='" & mProps.CodTipoDocCobro & "'"
        mProps.objEmpresa.Execute sql, True
    
        'Graba un registro para auditoria   *** ANGEL 20/mar/02
        mProps.objEmpresa.GrabaGNLogAccion "ELI_BAN", "TipoDocCobro Eliminado: " & _
                                                             mProps.CodTipoDocCobro, "DT"
                                                             
        mbooEsNuevo = True
        mbooModificado = True
        
        'Genera evento Refresh del objeto Empresa
        mProps.objEmpresa.RaiseRefreshEvent REFRESH_TIPODOCCOBRO
    End Sub
    
    Public Function Clone() As IVTipoDocCobro
        Dim nuevo As IVTipoDocCobro
        
        Set nuevo = New IVTipoDocCobro
        'Copia las propiedades al nuevo objeto
        nuevo.Props = mProps
        
        Set Clone = nuevo
        Set nuevo = Nothing
    End Function
    
    Friend Property Let Props(value As T_PROPS)         'Para el método Clone
        mProps = value
        mProps.IDTipoDocCobro = 0
        mbooModificado = True
    End Property
    
    Friend Property Set Empresa(objEmpresa As Empresa)
        Set mProps.objEmpresa = objEmpresa
    End Property
    
    Public Property Get Empresa() As Empresa
        Set Empresa = mProps.objEmpresa
    End Property
    
    
    Public Property Let CodTipoDocCobro(ByVal value As String)
        If Len(value) > 20 Then
            Err.Raise ERR_INVALIDO, "IVTipoDocCobro_CodTipoDocCobro", MSGERR_INVALIDO
        End If
        ValidaCodigo value      'Valida si contiene caracteres no validos
        
        mProps.CodTipoDocCobro = value
        mbooModificado = True
    End Property
    
    Public Property Get CodTipoDocCobro() As String
        CodTipoDocCobro = Trim$(mProps.CodTipoDocCobro)
    End Property
    
    Public Property Let Descripcion(ByVal value As String)
        If Len(value) > 50 Then
            Err.Raise ERR_INVALIDO, "IVTipoDocCobro_Descripcion", MSGERR_INVALIDO
        End If
        
        mProps.Descripcion = value
        mbooModificado = True
    End Property
    
    Public Property Get Descripcion() As String
        Descripcion = Trim$(mProps.Descripcion)
    End Property
    
    Public Property Let IdForma(ByVal value As Long)
        mProps.IdForma = value
        mbooModificado = True
    End Property
    
    Public Property Get IdForma() As Long
        IdForma = mProps.IdForma
    End Property
    
    
    Public Property Let CodForma(ByVal value As String)
        Dim sql As String, rs As Recordset
        
        'Cuando cambia el valor
        If value <> mProps.CodForma Then
            'Actualiza IdForma también
            sql = "SELECT IdForma FROM TsFormaCobroPago WHERE CodForma='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVTipoDocCobro_CodForma", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdForma = rs!IdForma
            End If
            rs.Close
        End If
    
        Set rs = Nothing
        mProps.CodForma = value
        mbooModificado = True
    End Property
    
    Public Property Get CodForma() As String
        CodForma = Trim$(mProps.CodForma)
    End Property
    
    
    Public Property Let BandValida(ByVal value As Boolean)
        mProps.BandValida = value
        mbooModificado = True
    End Property
    
    Public Property Get BandValida() As Boolean
        BandValida = mProps.BandValida
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
        With mProps
            .BandValida = True
        End With
        
        mbooEsNuevo = True
        mbooModificado = False
    End Sub
    
    Private Sub Class_Terminate()
        Set mProps.objEmpresa = Nothing
    End Sub
    
    Public Property Let IdBanco(ByVal value As Long)
        mProps.IdBanco = value
        mbooModificado = True
    End Property
    
    Public Property Get IdBanco() As Long
        IdBanco = mProps.IdBanco
    End Property
    
    
    Public Property Let CodBanco(ByVal value As String)
        Dim sql As String, rs As Recordset
        
        'Cuando cambia el valor
        If value <> mProps.CodBanco Then
            'Actualiza IdBanco también
            sql = "SELECT IdBanco FROM TsBancoCobroPago WHERE CodBanco='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVTipoDocCobro_CodBanco", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdBanco = rs!IdBanco
            End If
            rs.Close
        End If
    
        Set rs = Nothing
        mProps.CodBanco = value
        mbooModificado = True
    End Property
    
    Public Property Get CodBanco() As String
        CodBanco = Trim$(mProps.CodBanco)
    End Property
    
