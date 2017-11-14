VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVBodega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Empresa"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit




Private Type T_PROPS
    IdBodega As Long
    CodBodega As String
    Descripcion As String
    BandValida As Boolean
    IDSucursal As Long
    CodSucursal As String

    IDBodegaHijo As Long
    CodBodegaHijo As String
    BandBodegaHijo As Boolean
        
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
    sql = "SELECT CodBodega FROM IVBodega " & _
          "WHERE (CodBodega='" & mProps.CodBodega & "') AND " & _
                "(IdBodega<>" & mProps.IdBodega & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "IVBodega", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM IVBodega WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM IVBodega WHERE IdBodega=" & mProps.IdBodega
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "IVBodega_Grabar", MSGERR_MODIFICADO
        End If
    End If

    
    If mProps.CodBodega = mProps.CodBodegaHijo Then
        Err.Raise ERR_REPITECODIGO, "IVBodega", MSGERR_REPITECODIGO
    End If
    'Graba el registro en Pos
    With rs
        !CodBodega = mProps.CodBodega
        !Descripcion = mProps.Descripcion
        !BandValida = mProps.BandValida
        !IDSucursal = mProps.IDSucursal
        !IDBodegaHijo = mProps.IDBodegaHijo
        !BandBodegaHijo = mProps.BandBodegaHijo
        
        
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IdBodega = !IdBodega
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_BOD", "Bodega Nueva: " & _
                                                         mProps.CodBodega, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_BOD", "Bodega Modificada: " & _
                                                         mProps.CodBodega, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_BODEGA
    
    'CREA VISTA vwExistencia01

    VerificaExistenciaVista "vwExistencia" & mProps.IdBodega
    
    sql = " CREATE view vwExistencia" & mProps.IdBodega
    sql = sql & " as"
    sql = sql & " select idinventario, ive.idbodega, exist from ivexist ive"
    sql = sql & " inner join ivbodega ivb"
    sql = sql & " on ive.idbodega=ivb.idbodega"
    sql = sql & " where codbodega='" & mProps.CodBodega & "'"

    Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
    Set rs = Nothing
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT IVBodega.*, codsucursal, ivbPadre.codbodega as CodBodegaHijo "
        sql = sql & " FROM IVBodega "
        sql = sql & " left join gnsucursal on ivbodega.idsucursal=gnsucursal.idsucursal"
        sql = sql & " left join IVBodega ivbPadre on ivbodega.idBodegaHijo=ivbPadre.idBodega"
        sql = sql & " WHERE IVBodega.CodBodega='" & key & "'"
    Else
        sql = "SELECT IVBodega.*, codsucursal, ivbPadre.codbodega as CodBodegaHijo "
        sql = sql & " FROM IVBodega "
        sql = sql & " left join gnsucursal on ivbodega.idsucursal=gnsucursal.idsucursal  "
        sql = sql & " left join IVBodega ivbPadre on ivbodega.idBodegaHijo=ivbPadre.idBodega"
        sql = sql & " WHERE IVBodega.IdBodega=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IdBodega = !IdBodega
        mProps.CodBodega = !CodBodega
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!IDSucursal) Then mProps.IDSucursal = !IDSucursal
        If Not IsNull(!CodSucursal) Then mProps.CodSucursal = !CodSucursal
        
        If Not IsNull(!IDBodegaHijo) Then mProps.IDBodegaHijo = !IDBodegaHijo
        If Not IsNull(!CodBodegaHijo) Then mProps.CodBodegaHijo = !CodBodegaHijo
        If Not IsNull(!BandBodegaHijo) Then mProps.BandBodegaHijo = !BandBodegaHijo
        
        
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset

    'Verifica si existe items en esta bodega
    sql = "SELECT IdBodega FROM IVExist WHERE IdBodega=" & mProps.IdBodega
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "IVBodega.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing

    'Elimina la empresa
    sql = "DELETE FROM IVBodega WHERE IdBodega=" & mProps.IdBodega
    mProps.objEmpresa.Execute sql, True

    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_BOD", "Bodega Eliminada: " & _
                                                         mProps.CodBodega, "DT"
                                                         
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_BODEGA
End Sub

Friend Function Clone() As IVBodega
    Dim nuevo As IVBodega
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New IVBodega
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdBodega = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property







Public Property Let CodBodega(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "IVBodega_CodBodega", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodBodega = value
    mbooModificado = True
End Property

Public Property Get CodBodega() As String
    CodBodega = Trim$(mProps.CodBodega)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 120 Then
        Err.Raise ERR_INVALIDO, "IVBodega_Descripcion", MSGERR_INVALIDO
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
    
    mProps.BandValida = True
    mProps.IDBodegaHijo = 0
    mProps.BandBodegaHijo = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub


Public Property Let IDSucursal(ByVal value As Long)
    mProps.IDSucursal = value
    mbooModificado = True
End Property

Public Property Get IDSucursal() As Long
    IDSucursal = mProps.IDSucursal
End Property


Public Property Let CodSucursal(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    If value <> mProps.CodSucursal Then
        If Len(value) > 0 Then
            sql = "SELECT IdSucursal FROM GNSucursal WHERE CodSucursal='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVBodega_CodSucursal", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IDSucursal = rs!IDSucursal
            End If
            rs.Close
        Else
            IDSucursal = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodSucursal = value
    mbooModificado = True
End Property

Public Property Get CodSucursal() As String
    Dim sql As String
    Dim rs As Recordset

    'Cuando pide por primera vez
    If (Len(mProps.CodSucursal) = 0) And (mProps.IDSucursal <> 0) Then
        'Obtiene CodResponsable
        sql = "SELECT CodSucursal FROM GNSucursal WHERE IdSucursal=" & mProps.IDSucursal
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
        Else
            mProps.CodSucursal = rs!CodSucursal
        End If
        rs.Close
    End If

    Set rs = Nothing
    CodSucursal = Trim$(mProps.CodSucursal)
End Property


Public Property Let IDBodegaHijo(ByVal value As Long)
    mProps.IDBodegaHijo = value
    mbooModificado = True
End Property

Public Property Get IDBodegaHijo() As Long
    IDBodegaHijo = mProps.IDBodegaHijo
End Property


Public Property Let CodBodegaHijo(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    If value <> mProps.CodBodegaHijo Then
        If Len(value) > 0 Then
            sql = "SELECT IdBodega FROM IVBodega WHERE CodBodega='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVBodega_CodBodegaHijo", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IDBodegaHijo = rs!IdBodega
            End If
            rs.Close
        Else
            IDBodegaHijo = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodBodegaHijo = value
    mbooModificado = True
End Property

Public Property Get CodBodegaHijo() As String
    Dim sql As String
    Dim rs As Recordset

    'Cuando pide por primera vez
    If (Len(mProps.CodBodegaHijo) = 0) And (mProps.IDBodegaHijo <> 0) Then
        'Obtiene CodResponsable
        sql = "SELECT CodBodega FROM IVBodega WHERE IdBodega=" & mProps.IDBodegaHijo
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
        Else
            mProps.CodBodegaHijo = rs!CodBodega
        End If
        rs.Close
    End If

    Set rs = Nothing
    CodBodegaHijo = Trim$(mProps.CodBodegaHijo)
End Property


Public Property Let BandBodegaHijo(ByVal value As Boolean)
    mProps.BandBodegaHijo = value
    mbooModificado = True
End Property

Public Property Get BandBodegaHijo() As Boolean
    BandBodegaHijo = mProps.BandBodegaHijo
End Property


