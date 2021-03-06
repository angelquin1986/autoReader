VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FCVTablaComision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdTablaComision As Long
    CodTablaComision As String
    Nombre As String
    BandValida As Boolean
        
    FechaGrabado As Date
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean


Public Sub Grabar()
    Dim sql As String, rs As Recordset
    
    'Si no est� modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el c�digo                '*** MAKOTO 17/oct/00
    sql = "SELECT CodTablaComision FROM FCVTablaComision " & _
          "WHERE (CodTablaComision='" & mProps.CodTablaComision & "') AND " & _
                "(IdTablaComision<>" & mProps.IdTablaComision & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "FCVTablaComision", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM FCVTablaComision WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM FCVTablaComision WHERE IDTablaComision=" & mProps.IdTablaComision
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "FCVTablaComision.Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro
    With rs
        !CodTablaComision = mProps.CodTablaComision
        !Nombre = mProps.Nombre
        !BandValida = mProps.BandValida
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.FechaGrabado = !FechaGrabado
        mProps.IdTablaComision = !IdTablaComision
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_CTLOC", "TablaComision Nuevo: " & _
                                                         mProps.CodTablaComision, "CT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_CTLOC", "TablaComision Modificado: " & _
                                                         mProps.CodTablaComision, "CT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_FCVTABLACOMISION
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM FCVTablaComision  WHERE CodTablaComision='" & key & "'"
    Else
        sql = "SELECT * FROM FCVTablaComision WHERE IdTablaComision=" & key
    End If

    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    rs.MoveFirst   'para que se coloque al inicio jeaa
    'Recupera los datos
    With rs
        mProps.IdTablaComision = !IdTablaComision
        mProps.CodTablaComision = !CodTablaComision
        If Not IsNull(!Nombre) Then mProps.Nombre = !Nombre
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
    
    'Verifica si ya existe una transacci�n usando el mismo
    sql = "SELECT CodVendedor FROM FcVendedor  WHERE IdTablaComision=" & mProps.IdTablaComision
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "FCVTablaComision.Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close

    'Elimina el registro
    sql = "DELETE FROM FCVTablaComision WHERE IdTablaComision=" & mProps.IdTablaComision
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_TablaComision", "TablaComision Eliminado: " & _
                                                         mProps.CodTablaComision, "DT"

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_FCVTABLACOMISION
End Sub

Friend Function Clone() As FCVTablaComision
    Dim nuevo As FCVTablaComision

    'Copia las propiedades al nuevo objeto
    Set nuevo = New FCVTablaComision
    nuevo.Props = mProps

    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el m�todo Clone
    mProps = value
    mProps.IdTablaComision = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Get IdTablaComision() As Long
    IdTablaComision = mProps.IdTablaComision
End Property


Public Property Let CodTablaComision(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "FCVTablaComision_CodTablaComision", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodTablaComision = value
    mbooModificado = True
End Property

Public Property Get CodTablaComision() As String
    CodTablaComision = Trim$(mProps.CodTablaComision)
End Property

Public Property Let Nombre(ByVal value As String)
    If Len(value) > 40 Then
        Err.Raise ERR_INVALIDO, "CodTablaComision_Nombre", MSGERR_INVALIDO
    End If
    
    mProps.Nombre = value
    mbooModificado = True
End Property

Public Property Get Nombre() As String
    Nombre = Trim$(mProps.Nombre)
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
    'Valores iniciales
    mProps.BandValida = True
    
    mbooEsNuevo = True
    mbooModificado = False
    
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub







