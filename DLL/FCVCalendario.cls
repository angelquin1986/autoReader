VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FCVCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdCalendario As Long
    IdVendedor As Long
    CodVendedor As String
    fecha As Date
    Calendario As String
    fechagrabado As Date
    objEmpresa As Empresa       'Referencia al objeto padre
     objEmpresa2 As Empresa2
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean


Public Sub Grabar()
    Dim sql As String, rs As Recordset
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código                '*** MAKOTO 17/oct/00
'    sql = "SELECT Calendario FROM FCVCalendario " & _
'          "WHERE (fecha='" & mProps.Fecha & "') AND " & _
'                "(Idvendedor=" & mProps.IdVendedor & ")"
'
'    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
'    If Not rs.EOF Then      'Si ya existe el codigo
'        Err.Raise ERR_REPITECODIGO, "FCVCalendario", MSGERR_REPITECODIGO
'    End If
'    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM FCVCalendario WHERE 1=0"
        Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM FCVCalendario WHERE Id=" & mProps.IdCalendario
        Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
'        If rs!fechagrabado <> mProps.fechagrabado Then
'            Err.Raise ERR_MODIFICADO, "FCVCalendario.Grabar", MSGERR_MODIFICADO
'        End If
    End If

    'Graba el registro
    With rs
        !IdVendedor = mProps.IdVendedor
        !fecha = mProps.fecha
        !Calendario = mProps.Calendario
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IdCalendario = !Id
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
'    If mbooEsNuevo Then
'        mProps.objEmpresa.GrabaGNLogAccion "NUE_CAL", "TablaComision Nuevo: " & _
'                                                         mProps.Calendario, "FCV"
'    Else
'        mProps.objEmpresa.GrabaGNLogAccion "MOD_CAL", "TablaComision Modificado: " & _
'                                                         mProps.Calendario, "FCV"
'    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    'Genera evento Refresh del objeto Empresa
'    mProps.objEmpresa.RaiseRefreshEvent REFRESH_FCVCALENDARIO
End Sub

Friend Sub Recuperar(key As Variant, key2 As Variant)
    Dim sql As String, rs As Recordset
    
'    If VarType(key) = vbString Then
'        sql = "SELECT fcvc.*, fcv.Codvendedor FROM FCVCalendario FCVC INNER JOIN FCVENDEDOR FCV ON FCVC.Idvendedor= fcv.idvendedor WHERE Calendario='" & key & "'"
'    Else
        sql = "SELECT * FROM FCVCalendario WHERE Idvendedor=" & key & " and fecha='" & key2 & "'"
'    End If

    Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
    
    'Recupera los datos
    With rs
        mProps.IdCalendario = !Id
        mProps.IdVendedor = !IdVendedor
'        mProps.CodVendedor = !CodVendedor
        If Not IsNull(!Calendario) Then mProps.Calendario = !Calendario
        If Not IsNull(!fecha) Then mProps.fecha = !fecha
'        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    
    'Verifica si ya existe una transacción usando el mismo
'    sql = "SELECT CodVendedor FROM FcVendedor  WHERE IdCalendario=" & mProps.IdCalendario
'    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "FCVCalendario.Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close

    'Elimina el registro
'    sql = "DELETE FROM FCVCalendario WHERE IdCalendario=" & mProps.IdCalendario
'    mProps.objEmpresa.Execute sql, True
'
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_TablaComision", "TablaComision Eliminado: " & _
                                                         mProps.Calendario, "DT"

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_FCVCALENDARIO
End Sub

Friend Function Clone() As FCVCalendario
    Dim nuevo As FCVCalendario

    'Copia las propiedades al nuevo objeto
    Set nuevo = New FCVCalendario
    nuevo.Props = mProps

    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdCalendario = 0
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


Public Property Get IdCalendario() As Long
    IdCalendario = mProps.IdCalendario
End Property


Public Property Let Calendario(ByVal value As String)
    If Len(value) > 96 Then
        Err.Raise ERR_INVALIDO, "FCVCalendario_Calendario", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.Calendario = value
    mbooModificado = True
End Property

Public Property Get Calendario() As String
    Calendario = Trim$(mProps.Calendario)
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
    mProps.Calendario = "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000"
    mbooEsNuevo = True
    mbooModificado = False
    
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Property Let IdVendedor(ByVal value As Long)
    mProps.IdVendedor = value
End Property

Public Property Get IdVendedor() As Long
    IdVendedor = mProps.IdVendedor
End Property


Public Property Let CodVendedor(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodVendedor Then
        If Len(value) > 0 Then
            'Actualiza IdVendedor también
            sql = "SELECT IdVendedor FROM FCVendedor WHERE CodVendedor='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "FCVCalendario.CodVendedor", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdVendedor = rs!IdVendedor
            End If
            rs.Close
            Set rs = Nothing
        Else
            IdVendedor = 0
        End If
        
        mProps.CodVendedor = value
        
    End If
End Property

Public Property Get CodVendedor() As String
    CodVendedor = Trim$(mProps.CodVendedor)
End Property


Public Property Let fecha(value As Date)
    'Solo cuando cambia el valor
    If mProps.fecha <> value Then
        mProps.fecha = value
        mbooModificado = True
    End If
End Property

Public Property Get fecha() As Date
    fecha = mProps.fecha
End Property







