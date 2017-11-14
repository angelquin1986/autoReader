VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdPlan  As Long
    CodPlan As String
    CodPlanEnlaza As String
    Descripcion As String 'ESTE NO VOY A UTILIZAR BORRAR
    Nivel As Byte
    IdPlanEnlaza As Long
    BandValida As Boolean
    FechaGrabado As Date
    IdInventario As Long
    CodInventario As String
    BandTotal As Boolean
    Dias As Integer
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean
Private mcolIVPlanItem As Collection

Public Sub Grabar()
    Dim sql As String, rs As Recordset
    
'    If TipoCuenta = 0 Then
'        Err.Raise ERR_INVALIDO, "CTCuenta_Grabar", _
'            "Tipo de cuenta no está especificado."
'    End If
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código                '*** MAKOTO 17/oct/00 Modificado
    sql = "SELECT CodPlan FROM IVPlan " & _
          "WHERE (CodPlan='" & mProps.CodPlan & "') AND " & _
                "(IdPlan<>" & mProps.IdPlan & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "IVPlan", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM IVPlan WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM IVPlan WHERE IdPlan=" & mProps.IdPlan
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
       
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "IVPlan_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro
    With rs
        !CodPlan = mProps.CodPlan
'        !Descripcion = mProps.Descripcion
        !Nivel = mProps.Nivel
        !IdPlanEnlaza = mProps.IdPlanEnlaza
        !BandValida = mProps.BandValida
        !IdInventario = mProps.IdInventario
        !Dias = mProps.Dias
        !BandTotal = mProps.BandTotal
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified       'Para actualizar
#End If
        .Move 0             'Para actualizar
        mProps.FechaGrabado = !FechaGrabado
        mProps.IdPlan = !IdPlan

        
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 18/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_PLAN", "Plan Nueva: " & _
                                                         mProps.CodPlan, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_PLAN", "Plan Modificada: " & _
                                                         mProps.CodPlan, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVPLAN
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT A.*,B.CodPlan AS CodPlanEnlaza,iv.idinventario,iv.codinventario " & _
            " FROM IVPlan A Inner Join ivinventario iv on iv.idinventario = a.idinventario " & _
          " LEFT JOIN IVPlan B " & _
          " ON A.IdPlanEnlaza=B.IdPlan  " & _
          " "

    'sql = "Select * from ivplan a "
    
    If VarType(key) = vbString Then
        sql = sql & "WHERE a.CodPlan='" & key & "'"
    Else
        sql = sql & "WHERE a.IdPlan=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IdPlan = !IdPlan
        mProps.CodPlan = !CodPlan
        'If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!Nivel) Then mProps.Nivel = !Nivel
        If Not IsNull(!IdPlanEnlaza) Then mProps.IdPlanEnlaza = !IdPlanEnlaza
        If Not IsNull(!CodPlanEnlaza) Then mProps.CodPlanEnlaza = !CodPlanEnlaza
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!IdInventario) Then mProps.IdInventario = !IdInventario
        If Not IsNull(!CodInventario) Then mProps.CodInventario = !CodInventario
        If Not IsNull(!BandTotal) Then mProps.BandTotal = !BandTotal
        If Not IsNull(!Dias) Then mProps.Dias = !Dias
        
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
   ' RecuperaIVPlanDetalle
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    EliminaPlan mProps.IdPlan

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVPLAN
End Sub

Private Sub EliminaPlan(Id As Long)
    Dim sql As String, rs As Recordset
    
    'Verifica si ya tiene movimientos
    'sql = "SELECT IdPlan FROM CTLibroDetalle WHERE IdCuenta=" & Id
    'Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'If Not rs.EOF Then
        'Si ya tiene movimientos
    '    Err.Raise ERR_NOELIMINABLE, "CTCuenta.Eliminar", _
                  "La cuenta '" & mProps.CodCuenta & _
                  "' ya tiene movimientos, " & _
                  "por lo que no puede ser eliminada."
    'End If
    'rs.Close
    
    'Verifica si existe en PCProvCli registros que hace referencia a la cuenta
'    sql = "SELECT IdCuentaContable FROM PCProvCli WHERE IdCuentaContable=" & Id
'    Set rs = mProps.objEmpresa.OpenRecordset(sql)
'    If Not rs.EOF Then
'        'Si ya existen registros que hacen referencia
'        Err.Raise ERR_NOELIMINABLE, "CTCuenta.Eliminar", _
'                  "La cuenta '" & mProps.CodCuenta & _
'                  "' está siendo usada por uno o más proveedores/clientes, " & _
'                  "por lo que no puede ser eliminada."
'    End If
'    rs.Close
    
      
       
    'Obtiene los hijos
    sql = "SELECT IdPlan FROM IVPlan WHERE IdplanEnlaza=" & Id
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        'Elimina los hijos primero
        Do Until .EOF
            EliminaPlan rs!IdPlan
            .MoveNext
        Loop
        .Close
    End With

    'Elimina si misma
    sql = "DELETE FROM IVPlan WHERE IdPlan=" & Id
    mProps.objEmpresa.Execute sql, True
    
    Set rs = Nothing
    
    'Graba un registro para auditoria   *** ANGEL 18/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_IVPLAN", "Plan Eliminado: " & _
                                                         mProps.CodPlan, "DT"
End Sub

Public Function Clone() As IVPlan
    Dim nuevo As IVPlan
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New IVPlan
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdPlan = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Let CodPlan(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "IVPlan_CodPlan", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodPlan = value
    mbooModificado = True
End Property

Public Property Get CodPlan() As String
    CodPlan = Trim$(mProps.CodPlan)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "IVPlan_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let Nivel(ByVal value As Byte)
    mProps.Nivel = value
    mbooModificado = True
End Property

Public Property Get Nivel() As Byte
    Nivel = mProps.Nivel
End Property

Public Property Let IdPlanEnlaza(ByVal value As Long)
    mProps.IdPlanEnlaza = value
    mbooModificado = True
End Property

Public Property Get IdPlanEnlaza() As Long
    IdPlanEnlaza = mProps.IdPlanEnlaza
End Property

Public Property Let CodPlanEnlaza(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodPlanEnlaza Then
        'Actualiza IdCuentaSuma también
        sql = "SELECT IdPlan FROM IVPlan WHERE CodPlan='" & value & "'"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "IVPlan_CodPlan", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            IdPlanEnlaza = rs!IdPlan
        End If
        rs.Close
    End If
    
    Set rs = Nothing
    mProps.CodPlanEnlaza = value
    mbooModificado = True
End Property

Public Property Get CodPlanEnlaza() As String
    CodPlanEnlaza = Trim$(mProps.CodPlanEnlaza)
End Property


'Public Property Let BandTotal(ByVal value As Boolean)
'    Dim sql As String, rs As Recordset
'
'    'Si es modificación
'    If Not EsNuevo Then
'        'Cuando intenta cambiar al False, verifica si tiene hijos
'        If Not value Then
'            'Verifica si tiene hijos
'            sql = "SELECT CodCuenta FROM CTCuenta WHERE IdCuentaSuma=" & mProps.IdCuenta
'            Set rs = mProps.objEmpresa.OpenRecordset(sql)
'            'Si ya tiene hijos
'            If Not rs.EOF Then
'                Err.Raise ERR_INVALIDO, "CTCuenta_BandTotal", _
'                          "La cuenta '" & mProps.CodCuenta & "' tiene otras cuentas que pertenece a la misma, " & _
'                          "por lo que no se puede ser como cuenta de movimiento."
'            End If
'
'        'Cuando intenta cambiar a True, verifica si ya tiene movimientos
'        Else
'            'Verifica si tiene movimientos
'            sql = "SELECT IdCuenta FROM CTLibroDetalle WHERE IdCuenta=" & mProps.IdCuenta
'            Set rs = mProps.objEmpresa.OpenRecordset(sql)
'            'Si ya tiene movimientos
'            If Not rs.EOF Then
'                Err.Raise ERR_INVALIDO, "CTCuenta.BandTotal", _
'                          "La cuenta '" & mProps.CodCuenta & "' ya tiene movimientos, " & _
'                          "por lo que no puede ser como cuenta de total."
'            End If
'        End If
'    End If
'
'    Set rs = Nothing
'    mProps.BandTotal = value
'    mbooModificado = True
'End Property

'Public Property Get BandTotal() As Boolean
'    BandTotal = mProps.BandTotal
'End Property



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
    'Valores predeterminados
    With mProps
        .BandValida = True
    End With

    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub


Public Function CreaIVPlanItem() As IVPlanItem
    Dim obj As IVPlanItem
   Set obj = New IVPlanItem
    Set obj.objPlan = Me
    Set CreaIVPlanItem = obj
    Set obj = Nothing
End Function

Public Function RecuperaIVPlanItem(key As Variant) As IVPlanItem
    Dim obj As IVPlanItem
    Set obj = New IVPlanItem
    Set obj.objPlan = Me
    On Error Resume Next
    obj.Recuperar key
    If Err.Number = 0 Then Set RecuperaIVPlanItem = obj
    Err.Clear
    On Error GoTo 0
    Set obj = Nothing
End Function


Private Sub RecuperaIVPlanDetalle()
    Dim sql As String, rs As Recordset, i As Long, mod_antes As Boolean
        
    mod_antes = mbooModificado      'Guarda el estado de Modificado
    If mcolIVPlanItem Is Nothing Then Set mcolIVPlanItem = New Collection
    
'    BorrarIVKardex
    
    
    'Recupera los IVKardex
    'sql = "SELECT * FROM vwIVKardex WHERE TransId=" & mProps.transid & " ORDER BY Orden"
'    VerificaExistenciaTabla 99
'    sql = "select ivplan.codplan as codplanNivel,ivplan.descripcion as DescNivel, " & _
'          " iv.idinventario as idinventario," & _
'          " ivplan1.codplan as codplanNivel1,ivplan1.descripcion as DescNivel1," & _
'          " iv1.idinventario as idinventario1," & _
'          " ivplan2.codplan as codplanNivel2,ivplan2.descripcion as DescNivel2,iv2.idinventario as idinventario2," & _
'          " ivplan3.codplan as codplanNivel3,ivplan3.descripcion as DescNivel3,iv3.idinventario as idinventario3," & _
'          " ivplan4.codplan as codplanNivel4,ivplan4.descripcion as DescNivel4,iv4.idinventario as idinventario4," & _
'          " ivplan5.codplan as codplanNivel5,ivplan5.descripcion as DescNivel5,iv5.idinventario as idinventario5" & _
'          " Into tmp99" & _
'          " from ivplan  " & _
'          " left join ivplanItem ivpI left join ivinventario iv on iv.idinventario = ivpi.idinventario on ivpI.idplan = ivplan.idplan  left join ivplan  ivplan1 " & _
'          " left join ivplanItem ivpI1 left join ivinventario iv1 on iv1.idinventario = ivpi1.idinventario on ivpI1.idplan = ivplan1.idplan left join ivplan ivplan2 " & _
'          " left join ivplanItem ivpI2 left join ivinventario iv2 on iv2.idinventario = ivpi2.idinventario on ivpI2.idplan = ivplan2.idplan" & _
'          " left join ivplan ivplan3 left join ivplanItem ivpI3 left join ivinventario iv3 on iv3.idinventario = ivpi3.idinventario on ivpI3.idplan = ivplan3.idplan" & _
'          " left join ivplan ivplan4 left join ivplanItem ivpI4 left join ivinventario iv4 on iv4.idinventario = ivpi4.idinventario on ivpI4.idplan = ivplan4.idplan" & _
'          " left join ivplan ivplan5 left join ivplanItem ivpI5 left join ivinventario iv5 on iv5.idinventario = ivpi5.idinventario on ivpI5.idplan = ivplan5.idplan" & _
'          " on ivplan5.idplanenlaza = ivplan4.idplan" & _
'          " on ivplan4.idplanenlaza = ivplan3.idplan" & _
'          " on ivplan3.idplanenlaza = ivplan2.idplan" & _
'          " on ivplan2.idplanenlaza = ivplan1.idplan" & _
'          " on ivplan1.idplanenlaza = ivplan.idplan" & _
'          " where ivplan.codplan = '" & mProps.CodPlan & "'"
'          mProps.objEmpresa.EjecutarSQL sql, 1
'
'          VerificaExistenciaTabla 1
'          sql = "select iv.idinventario,x.descNivel "
'          sql = sql & " Into tmp1"
'          sql = sql & " from tmp99 x"
'          sql = sql & " inner join ivinventario iv on iv.idinventario = x.idinventario"
'          mProps.objEmpresa.EjecutarSQL sql, 1
'
'          VerificaExistenciaTabla 2
'          sql = " select iv.idinventario,x.descNivel1"
'          sql = sql & " into tmp2"
'          sql = sql & " from tmp99 x"
'          sql = sql & " inner join ivinventario iv on iv.idinventario = x.idinventario1"
'          mProps.objEmpresa.EjecutarSQL sql, 1
'
'          VerificaExistenciaTabla 3
'          sql = "select iv.idinventario,x.descNivel2 "
'          sql = sql & " into tmp3 "
'          sql = sql & " from tmp99 x "
'          sql = sql & " inner join ivinventario iv on iv.idinventario = x.idinventario2"
'          mProps.objEmpresa.EjecutarSQL sql, 1
'
'          VerificaExistenciaTabla 4
'          sql = "select iv.idinventario,x.descNivel3 "
'          sql = sql & " into tmp4 "
'          sql = sql & " from tmp99 x"
'          sql = sql & " inner join ivinventario iv on iv.idinventario = x.idinventario3"
'          mProps.objEmpresa.EjecutarSQL sql, 1
'
'          VerificaExistenciaTabla 5
'          sql = "select iv.idinventario,x.descNivel4"
'          sql = sql & " into tmp5 "
'          sql = sql & " from tmp99 x"
'          sql = sql & " inner join ivinventario iv on iv.idinventario = x.idinventario4"
'          mProps.objEmpresa.EjecutarSQL sql, 1
'
'          VerificaExistenciaTabla 6
'          sql = " select iv.idinventario,x.descNivel5"
'          sql = sql & " into tmp6"
'          sql = sql & " from tmp99 x"
'          sql = sql & " inner join ivinventario iv on iv.idinventario = x.idinventario5"
'          mProps.objEmpresa.EjecutarSQL sql, 1
'
'          VerificaExistenciaTabla 7
'            sql = "select * into tmp7 from tmp1 "
'            sql = sql & " Union All"
'            sql = sql & " select * from tmp2"
'            sql = sql & " Union All"
'            sql = sql & " select * from tmp3"
'            sql = sql & " Union All"
'            sql = sql & " select * from tmp4"
'            sql = sql & " Union All"
'            sql = sql & " select * from tmp5"
'            sql = sql & " Union All"
'            sql = sql & " select * from tmp6"
'        mProps.objEmpresa.EjecutarSQL sql, 1
'
'        sql = " Select iv.idinventario,iv.codinventario,iv.descripcion,tmp7.DescNivel from tmp7 inner join ivinventario iv on iv.idinventario = tmp7.idinventario"
        sql = "Select iv.idinventario,iv.codinventario,iv.descripcion,ivori.codinventario as descNivel,ivp.dias as frecuencia "
        sql = sql & " from ivplan ivp inner join ivplan ivpOri inner join ivinventario ivOri on ivori.idinventario = ivpori.idinventario "
        sql = sql & " on ivpori.idplan = ivp.idplanenlaza inner join ivinventario iv on iv.idinventario = ivp.idinventario"
        sql = sql & " Where ivp.idplanenlaza= " & mProps.IdPlan
        
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            If !IdInventario > 0 Then                 'Para asegurar el Bug de ADO
                i = AddIVPlanItem
                IVPlanItem(i).RecuperarDetalle rs
            End If
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    mbooModificado = mod_antes      'Para que no cambie el estado de Modificado
End Sub
Public Function AddIVPlanItem() As Long
    Dim nuevo As IVPlanItem
    
    'Cuando SoloVer está establecido a True, genera error
'    If mbooSoloVer Then Err.Raise ERR_SOLOVER, "GNComprobante.AddIVKardex", MSGERR_SOLOVER
    If mcolIVPlanItem Is Nothing Then RecuperaIVPlanDetalle
  
    
    Set nuevo = New IVPlanItem
    Set nuevo.objPlan = Me
        
    mcolIVPlanItem.Add nuevo
    AddIVPlanItem = mcolIVPlanItem.Count
    Set nuevo = Nothing
    mbooModificado = True
End Function

Public Property Get CountIVPlanItem() As Long
    If mcolIVPlanItem Is Nothing Then RecuperaIVPlanDetalle
    CountIVPlanItem = mcolIVPlanItem.Count
End Property

Public Property Get IVPlanItem(ix As Long) As IVPlanItem
    If mcolIVPlanItem Is Nothing Then RecuperaIVPlanDetalle
    Set IVPlanItem = mcolIVPlanItem.item(ix)
End Property

Public Property Let IdInventario(ByVal value As Long)
    mProps.IdInventario = value
    mbooModificado = True
End Property

Public Property Get IdInventario() As Long
    IdInventario = mProps.IdInventario
End Property

Public Property Get CodInventario() As String
    CodInventario = mProps.CodInventario
End Property

Public Property Let CodInventario(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodInventario Then
        If Len(value) > 0 Then
            
            sql = "SELECT IdInventario FROM IVInventario WHERE CodInventario='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVPlan_CodInventario", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdInventario = rs!IdInventario
            End If
            rs.Close
        Else
            IdInventario = 0
        End If
        
        mProps.CodInventario = value
        mbooModificado = True
    End If

    Set rs = Nothing
End Property

Public Property Get Dias() As Integer
    Dias = mProps.Dias
End Property

Public Property Let Dias(ByVal value As Integer)
    mProps.Dias = value
    mbooModificado = True
End Property

Public Property Let BandTotal(ByVal value As Boolean)
    mProps.BandTotal = value
    mbooModificado = True
End Property

Public Property Get BandTotal() As Boolean
    BandTotal = mProps.BandTotal
End Property

Public Property Get IdPlan() As Long
    IdPlan = mProps.IdPlan
End Property

