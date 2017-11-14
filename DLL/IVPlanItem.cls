VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVPlanItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdPlanItem  As Long
    IdPlan As Long
    CodPlan As String
    IdInventario As Long
    CodInventario As String
    IdGrupo1 As Long
    CodGrupo1 As String
    Prioridad As Integer
    Frecuencia As Integer
    DescPlan As String
    FechaGrabado As Date
    DescItem As String
    objPlan As IVPlan
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Public Sub Grabar()
    Dim sql As String, rs As Recordset
'    If TipoCuenta = 0 Then
'        Err.Raise ERR_INVALIDO, "CTCuenta_Grabar", _
'            "Tipo de cuenta no está especificado."
'    End If
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    'Verifica si no repite el código                '*** MAKOTO 17/oct/00 Modificado
'    sql = "SELECT CodPlan FROM IVPlan " & _
'          "WHERE (CodPlan='" & mProps.CodPlan & "') AND " & _
'                "(IdPlan<>" & mProps.IdPlan & ")"
'    Set rs = mProps.objEmpresa.OpenRecordset(sql)
'    If Not rs.EOF Then      'Si ya existe el codigo
'        Err.Raise ERR_REPITECODIGO, "IVPlan", MSGERR_REPITECODIGO
'    End If
'    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM IVPlanItem WHERE 1=0"
        Set rs = mProps.objPlan.Empresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM IVPlanItem WHERE IdPlanitem=" & mProps.IdPlanItem
        Set rs = mProps.objPlan.Empresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
       
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "IVPlanItem_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro
    With rs
        '!IdPlanItem = mProps.IdPlanItem
        !IdPlan = mProps.IdPlan
        !IdInventario = mProps.IdInventario
        !IdGrupo1 = mProps.IdGrupo1
        !Prioridad = mProps.Prioridad
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified       'Para actualizar
#End If
        .Move 0             'Para actualizar
        mProps.FechaGrabado = !FechaGrabado
        mProps.IdPlanItem = !IdPlanItem

        
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 18/mar/02
    If mbooEsNuevo Then
        mProps.objPlan.Empresa.GrabaGNLogAccion "NUE_PLANITEM", "Plan Detalle Nueva: " & _
                                                         mProps.CodPlan, "DT"
    Else
        mProps.objPlan.Empresa.GrabaGNLogAccion "MOD_PLANITEM", "Plan Detalle Modificada: " & _
                                                         mProps.CodPlan, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objPlan.Empresa.RaiseRefreshEvent REFRESH_IVPLANITEM
End Sub

Public Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT ivpi.*,iv.Codinventario,ivp.codplan,ivg1.codgrupo1 " & _
            "  FROM IVPlanItem ivpi " & _
          " LEFT JOIN IVPlan ivp ON ivp.IdPlan=ivpi.IdPlan " & _
          " LEFT JOIN IVInventario iv on iv.idinventario = ivpi.idinventario" & _
          " LEFT JOIN IVGrupo1 ivg1 on ivg1.idGrupo1 = ivpi.idGrupo1 "
    
           sql = sql & "WHERE ivpi.IdPlanItem=" & key
    
    Set rs = mProps.objPlan.Empresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IdPlanItem = !IdPlanItem
        mProps.IdPlan = !IdPlan
        If Not IsNull(!CodPlan) Then mProps.CodPlan = !CodPlan
        mProps.IdInventario = !IdInventario
        If Not IsNull(!CodInventario) Then mProps.CodInventario = !CodInventario
        mProps.IdGrupo1 = !IdGrupo1
        If Not IsNull(!CodGrupo1) Then mProps.CodGrupo1 = !CodGrupo1
        If Not IsNull(!Prioridad) Then mProps.Prioridad = !Prioridad
        If Not IsNull(!Frecuencia) Then mProps.Frecuencia = !Frecuencia
        
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    EliminaPlan mProps.IdPlanItem

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objPlan.Empresa.RaiseRefreshEvent REFRESH_IVPLANITEM
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
    sql = "SELECT IdPlanItem FROM IVPlanItem WHERE IdplanItem=" & Id
    Set rs = mProps.objPlan.Empresa.OpenRecordset(sql)
    With rs
        'Elimina los hijos primero
        Do Until .EOF
            EliminaPlan rs!IdPlanItem
            .MoveNext
        Loop
        .Close
    End With

    'Elimina si misma
    sql = "DELETE FROM IVPlanitem WHERE IdPlanitem=" & Id
    mProps.objPlan.Empresa.Execute sql, True
    
    Set rs = Nothing
    
    'Graba un registro para auditoria   *** ANGEL 18/mar/02
    mProps.objPlan.Empresa.GrabaGNLogAccion "ELI_IVPLANITEM", "PlanItem Eliminado: " & _
                                                         mProps.CodPlan, "DT"
End Sub

Public Function Clone() As IVPlanItem
    Dim nuevo As IVPlanItem
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New IVPlanItem
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdPlanItem = 0
    mbooModificado = True
End Property

Friend Property Set objPlan(objPlan As IVPlan)
    Set mProps.objPlan = objPlan
End Property

Public Property Get objPlan() As IVPlan
    Set objPlan = mProps.objPlan
End Property

'Public Property Let CodPlan(ByVal value As String)
'    mProps.CodPlan = value
'    mbooModificado = True
'End Property

Public Property Get CodPlan() As String
    CodPlan = Trim$(mProps.CodPlan)
End Property
Public Property Get DescPlan() As String
    DescPlan = Trim$(mProps.DescPlan)
End Property
Public Property Let DescPlan(ByVal value As String)
    mProps.DescPlan = value
    mbooModificado = True
End Property


Public Property Let Prioridad(ByVal value As Integer)
    mProps.Prioridad = value
    mbooModificado = True
End Property

Public Property Get Prioridad() As Integer
    Prioridad = mProps.Prioridad
End Property

'
'
'
'Public Property Let BandValida(ByVal value As Boolean)
'    mProps.BandValida = value
'    mbooModificado = True
'End Property
'
'Public Property Get BandValida() As Boolean
'    BandValida = mProps.BandValida
'End Property



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
    
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objPlan = Nothing
End Sub

Public Property Let CodPlan(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodPlan Then
        If Len(value) > 0 Then
            'Actualiza IdProveedor también
            sql = "SELECT IdPlan FROM IVPlan WHERE CodPlan='" & value & "'"
            Set rs = mProps.objPlan.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVPlan_CodPlan", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdPlan = rs!IdPlan
            End If
            rs.Close
        Else
            mProps.IdPlan = 0
        End If
    End If
    Set rs = Nothing
    mProps.CodPlan = value
    mbooModificado = True
End Property

Public Property Let CodInventario(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodInventario Then
        If Len(value) > 0 Then
            'Actualiza IdProveedor también
            sql = "SELECT IdInventario FROM IVinventario WHERE CodInventario='" & value & "'"
            Set rs = mProps.objPlan.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVInventario_CodInventario", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdInventario = rs!IdInventario
            End If
            rs.Close
        Else
            mProps.IdInventario = 0
        End If
    End If
    Set rs = Nothing
    mProps.CodInventario = value
    mbooModificado = True
End Property

Public Property Get CodInventario() As String
    CodInventario = Trim$(mProps.CodInventario)
End Property

Public Property Let CodGrupo1(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodGrupo1 Then
        If Len(value) > 0 Then
            'Actualiza IdProveedor también
            sql = "SELECT IdGrupo1 FROM IVGrupo1 WHERE CodGrupo1='" & value & "'"
            Set rs = mProps.objPlan.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVGrupo1_CodGrupo1", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                mProps.IdGrupo1 = rs!IdGrupo1
            End If
            rs.Close
        Else
            mProps.IdGrupo1 = 0
        End If
    End If
    Set rs = Nothing
    mProps.CodGrupo1 = value
    mbooModificado = True
End Property

Public Property Get CodGrupo1() As String
    CodGrupo1 = Trim$(mProps.CodGrupo1)
End Property
Public Property Get DescItem() As String
    DescItem = Trim$(mProps.DescItem)
End Property
Public Property Let DescItem(ByVal value As String)
    mProps.DescItem = value
    mbooModificado = True
End Property

Public Property Let Frecuencia(ByVal value As Integer)
    mProps.Frecuencia = value
    mbooModificado = True
End Property

Public Property Get Frecuencia() As Integer
    Frecuencia = mProps.Frecuencia
End Property
Friend Sub RecuperarDetalle(rs As Recordset)
Dim i As Integer
    With rs
        If Not .EOF Then
            If Not IsNull(!IdInventario) Then mProps.IdInventario = !IdInventario
            If Not IsNull(!CodInventario) Then mProps.CodInventario = !CodInventario
            If Not IsNull(!Descripcion) Then mProps.DescItem = !Descripcion
            If Not IsNull(!Frecuencia) Then mProps.Frecuencia = !Frecuencia
        End If
    End With
End Sub

