VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCProvincia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
    Option Explicit

Private Type T_PROPS
    IdProvincia As Long
    CodProvincia As String
    Descripcion As String
    BandValida As Boolean
    fechagrabado As Date
    CodProvinciaSC As String
    IdPais As Long
    CodPais As String
    objEmpresa As Empresa       'Referencia al objeto padre
    
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Public Sub Grabar()
    Dim sql As String, rs As Recordset
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    'Verifica si no repite el código                '*** MAKOTO 17/oct/00
    sql = "SELECT CodProvincia FROM PCProvincia " & _
          "WHERE (CodProvincia='" & mProps.CodProvincia & "') AND " & _
                "(IdProvincia<>" & mProps.IdProvincia & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "PCProvincia", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM PCProvincia WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM PCProvincia WHERE IDProvincia=" & mProps.IdProvincia
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "PCProvincia.Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro
    With rs
        If Not !CodProvincia = mProps.CodProvincia Then
            mProps.objEmpresa.GrabaAccionModulo !CodProvincia, mProps.CodProvincia, "Codigo Provincia", "MOD_SUC", "GN"
        End If
        !CodProvincia = mProps.CodProvincia
        If Not !Descripcion = mProps.Descripcion Then
            mProps.objEmpresa.GrabaAccionModulo !Descripcion, mProps.Descripcion, "Descripcion Provincia", "MOD_SUC", "GN"
        End If
        !Descripcion = mProps.Descripcion
        !BandValida = mProps.BandValida
        
        !CodProvinciaSC = mProps.CodProvinciaSC
        !IdPais = mProps.IdPais
        
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IdProvincia = !IdProvincia
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_PCPROV", "Provincia Nuevo: " & _
                                                         mProps.CodProvincia, "GN"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_PCPROV", "Provincia Modificado: " & _
                                                         mProps.CodProvincia, "GN"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PCPROVINCIA
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT pcp.*,pcpais.codpais FROM PCProvincia pcp Left Join pcpais on pcpais.idpais =pcp.idpais  WHERE CodProvincia='" & key & "'"
    Else
        sql = "SELECT pcp.*,pcpais.codpais FROM PCProvincia pcp Left Join pcpais on pcpais.idpais =pcp.idpais WHERE IdProvincia=" & key
    End If

    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    rs.MoveFirst   'para que se coloque al inicio jeaa
    'Recupera los datos
    With rs
        mProps.IdProvincia = !IdProvincia
        mProps.CodProvincia = !CodProvincia
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!CodProvinciaSC) Then mProps.CodProvinciaSC = !CodProvinciaSC
        If Not IsNull(!IdPais) Then mProps.IdPais = !IdPais
        If Not IsNull(!CodPais) Then mProps.CodPais = !CodPais
        
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
    sql = "SELECT IDPROVINCIA FROM PCCANTON WHERE IdProvincia=" & mProps.IdProvincia
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "PCProvincia.Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close

    'Elimina el registro
    sql = "DELETE FROM PCProvincia WHERE IdProvincia=" & mProps.IdProvincia
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_PCPROV", "Provincia Eliminado: " & _
                                                         mProps.CodProvincia, "GN"

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PCPROVINCIA
End Sub

Friend Function Clone() As PCProvincia
    Dim nuevo As PCProvincia

    'Copia las propiedades al nuevo objeto
    Set nuevo = New PCProvincia
    nuevo.Props = mProps

    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdProvincia = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property





Public Property Let CodProvincia(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "PCProvincia_CodProvincia", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodProvincia = value
    mbooModificado = True
End Property

Public Property Get CodProvincia() As String
    CodProvincia = Trim$(mProps.CodProvincia)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 40 Then
        Err.Raise ERR_INVALIDO, "CodProvincia_Descripcion", MSGERR_INVALIDO
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

Public Property Let CodProvinciaSC(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "PCProvincia_CodProvinciaSC", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodProvinciaSC = value
    mbooModificado = True
End Property

Public Property Get CodProvinciaSC() As String
    CodProvinciaSC = Trim$(mProps.CodProvinciaSC)
End Property

Public Property Let IdPais(ByVal value As Long)
    mProps.IdPais = value
    mbooModificado = True
End Property

Public Property Get IdPais() As Long
    IdPais = mProps.IdPais
End Property


Public Property Let CodPais(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodPais Then
        'Actualiza IdCanton también
        sql = "SELECT Idpais FROM PCPais WHERE CodPais='" & value & "'"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "PCProvincia_CodPais", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            IdPais = rs!IdPais
        End If
        rs.Close
    End If

    Set rs = Nothing
    mProps.CodPais = value
    mbooModificado = True
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "PCProvincia_CodPais", MSGERR_INVALIDO
    End If
    mProps.CodPais = value
    mbooModificado = True
End Property

Public Property Get CodPais() As String
    CodPais = Trim$(mProps.CodPais)
End Property


