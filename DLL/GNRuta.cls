VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GNRuta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    Id As Long
    Codruta As String
    Descripcion As String
    BandValida As Boolean
    fechagrabado As Date
    diasemana As String
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
    sql = "SELECT CodRuta FROM GNRuta " & _
          "WHERE (CodRuta='" & mProps.Codruta & "') AND " & _
                "(Id<>" & mProps.Id & ")"
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "GNRuta", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM GNRuta WHERE 1=0"
        Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
        rs.AddNew
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM GNRuta WHERE Id=" & mProps.Id
        Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "GNruta_Grabar", MSGERR_MODIFICADO
        End If
    End If
    'Graba el registro en Pos
    With rs
        !Codruta = mProps.Codruta
        !Descripcion = mProps.Descripcion
        !BandValida = mProps.BandValida
        !diasemana = mProps.diasemana
        !fechagrabado = Now
        
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.Id = !Id
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        gobjMain.EmpresaActual.GrabaGNLogAccion "NUE_RUTA", "Seguro Nuevo: " & _
                                                         mProps.Codruta, "DT"
    Else
       gobjMain.EmpresaActual.GrabaGNLogAccion "MOD_RUTA", "Seguro Modificado: " & _
                                                         mProps.Codruta, "DT"
    End If
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    'Genera evento Refresh del objeto Empresa
    gobjMain.EmpresaActual.RaiseRefreshEvent REFRESH_GNRUTA
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    If VarType(key) = vbString Then
        sql = "SELECT * FROM GNRuta"
        sql = sql & " WHERE Codruta='" & key & "'"
    Else
        sql = "SELECT * FROM GNruta "
        sql = sql & " Where Id = " & key
    End If
    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    'Recupera los datos de Usuario
    With rs
        mProps.Id = !Id
        mProps.Codruta = !Codruta
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!diasemana) Then mProps.diasemana = !diasemana
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
'    sql = "SELECT IdSeguro FROM GNoferta WHERE IdSeguro=" & mProps.Id
'    Set rs = mProps.objEmpresa.OpenRecordset(sql)
'    If Not rs.EOF Then
'        Err.Raise ERR_NOELIMINABLE, "GNseguro.Eliminar", MSGERR_NOELIMINABLE
'    End If
'    rs.Close

    'Elimina el registro
    sql = "DELETE FROM GNRuta WHERE Id=" & mProps.Id
    gobjMain.EmpresaActual.Execute sql, True
    
    gobjMain.EmpresaActual.GrabaGNLogAccion "ELI_RUTA", "Seguro Eliminado: " & _
                                                         mProps.Codruta, "DT"
    mbooEsNuevo = True
    mbooModificado = True
    'Genera evento Refresh del objeto Empresa
    gobjMain.EmpresaActual.RaiseRefreshEvent REFRESH_GNRUTA
End Sub

Friend Function Clone() As GNRuta
    Dim nuevo As GNRuta
    'Copia las propiedades al nuevo objeto
    Set nuevo = New GNRuta
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

Public Property Let Codruta(ByVal value As String)
    If Len(value) > 10 Then
        Err.Raise ERR_INVALIDO, "GNRuta_CodRuta", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    mProps.Codruta = value
    mbooModificado = True
End Property

Public Property Get Codruta() As String
    Codruta = Trim$(mProps.Codruta)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 40 Then
        Err.Raise ERR_INVALIDO, "GNRuta_Descripcion", MSGERR_INVALIDO
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
Public Property Get diasemana() As String
    diasemana = Trim$(mProps.diasemana)
End Property

Public Property Let diasemana(ByVal value As String)
    If Len(value) > 40 Then
        Err.Raise ERR_INVALIDO, "GNRuta_diasemana", MSGERR_INVALIDO
    End If
    mProps.diasemana = value
    mbooModificado = True
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
    Set mProps.objEmpresa2 = Nothing
End Sub

