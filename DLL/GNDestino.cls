VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GNDestino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IDDestino As Long
    CodDestino As String
    Descripcion As String
    BandValida As Boolean
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
    
    'Verifica si no repite el código                '*** MAKOTO 17/oct/00
    sql = "SELECT CodDestino FROM GNDestino " & _
          "WHERE (CodDestino='" & mProps.CodDestino & "') AND " & _
                "(IdDestino<>" & mProps.IDDestino & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "GNDestino", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM GNDestino WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM GNDestino WHERE IDDestino=" & mProps.IDDestino
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "GNDestino.Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro
    With rs
        If Not !CodDestino = mProps.CodDestino Then
            mProps.objEmpresa.GrabaAccionModulo !CodDestino, mProps.CodDestino, "Codigo Destino", "MOD_SUC", "GN"
        End If
        !CodDestino = mProps.CodDestino
        If Not !Descripcion = mProps.Descripcion Then
            mProps.objEmpresa.GrabaAccionModulo !Descripcion, mProps.Descripcion, "Descripcion Destino", "MOD_SUC", "GN"
        End If
        !Descripcion = mProps.Descripcion
        !BandValida = mProps.BandValida
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.FechaGrabado = !FechaGrabado
        mProps.IDDestino = !IDDestino
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_DTN", "Destino Nuevo: " & _
                                                         mProps.CodDestino, "GN"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_DTN", "Destino Modificado: " & _
                                                         mProps.CodDestino, "GN"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_GNDESTINO
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM GNDestino  WHERE CodDestino='" & key & "'"
    Else
        sql = "SELECT * FROM GNDestino WHERE IdDestino=" & key
    End If

    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    rs.MoveFirst   'para que se coloque al inicio jeaa
    'Recupera los datos
    With rs
        mProps.IDDestino = !IDDestino
        mProps.CodDestino = !CodDestino
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
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
    
    'Verifica si ya existe una transacción usando el mismo
    sql = "SELECT codmotivo FROM Motivo WHERE IdDestino=" & mProps.IDDestino
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "GNDestino.Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close

    'Elimina el registro
    sql = "DELETE FROM GNDestino WHERE IdDestino=" & mProps.IDDestino
    mProps.objEmpresa.Execute sql, True
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_DTN", "Destino Eliminado: " & _
                                                         mProps.CodDestino, "GN"

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_GNDESTINO
End Sub

Friend Function Clone() As GNDestino
    Dim nuevo As GNDestino

    'Copia las propiedades al nuevo objeto
    Set nuevo = New GNDestino
    nuevo.Props = mProps

    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IDDestino = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property





Public Property Let CodDestino(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "GNDestino_CodDestino", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodDestino = value
    mbooModificado = True
End Property

Public Property Get CodDestino() As String
    CodDestino = Trim$(mProps.CodDestino)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 40 Then
        Err.Raise ERR_INVALIDO, "CodDestino_Descripcion", MSGERR_INVALIDO
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






