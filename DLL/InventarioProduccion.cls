VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InventarioProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Type T_PROPS
    Id As Long
    IdInventario As Long
    CodInventario As String
    Descripcion As String
    IdGrupo(1 To 5) As Long
    CodGrupo(1 To 5) As String
    
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Public Sub Grabar()
    Dim sql As String, i As Integer
    Dim rs As Recordset, rsDetalle As Recordset
    Dim rsSuc As Recordset, gnSuc As GNSucursal
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
        
        
    'Verifica si no repite el código
    sql = "SELECT idInventario FROM IVInventarioProduccion " & _
          "WHERE (idInventario='" & mProps.IdInventario & "')"
          
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "IVInventarioDetalle", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Verifica si no repite el código Alterno 1
    
    
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM IVInventarioProduccion WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM IVInventarioProduccion WHERE Id=" & mProps.Id
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        
    End If
    
    'Comienza la transaccion ----------------------------jeaa 16-03-04
#If DAOLIB Then
    gobjWS.BeginTrans
#Else
    mProps.objEmpresa.Coneccion.BeginTrans
#End If
    On Error GoTo ErrTrap
    'Graba el registro
    With rs
        !IdInventario = mProps.IdInventario
        !Descripcion = mProps.Descripcion
        
        
                
        For i = 1 To IVGRUPO_MAX
            .Fields("IdGrupo" & i) = mProps.IdGrupo(i)
        Next i

        .Update
        
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        rs.Move 0           'Para actualizar
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    
    Set rs = Nothing
    
    
    
    mbooEsNuevo = False
    mbooModificado = False
    
    'Completa la transacción ---------------------------- jeaa 16-03-04
#If DAOLIB Then
    gobjWS.CommitTrans
#Else
    mProps.objEmpresa.Coneccion.CommitTrans
#End If
    Exit Sub
ErrTrap:
    'Anula la transacción ---------------------------- jeaa 16-03-04
#If DAOLIB Then
    gobjWS.Rollback
#Else
    mProps.objEmpresa.Coneccion.RollbackTrans
#End If

End Sub


Friend Sub Recuperar(key As Variant)
    Dim sql As String, i As Integer
    Dim rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM IVInventarioProduccion WHERE id='" & key & "'"
    Else
        sql = "SELECT * FROM IVInventarioProduccion WHERE Id=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.Id = !Id
        mProps.IdInventario = !IdInventario
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        
        For i = 1 To IVGRUPO_MAX
            If Not IsNull(.Fields("IdGrupo" & i)) Then mProps.IdGrupo(i) = .Fields("IdGrupo" & i)
            If Not IsNull(.Fields("CodGrupo" & i)) Then mProps.CodGrupo(i) = .Fields("CodGrupo" & i)
        Next i
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub



Public Sub Eliminar()
    Dim sql As String, rs As Recordset
    
    
    
    'Elimina el registro
    sql = "DELETE FROM IVInventarioProduccion WHERE Id=" & mProps.Id
    mProps.objEmpresa.Execute sql, True

    
                                                         
    mbooEsNuevo = True
    mbooModificado = True
    
    
End Sub


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

Friend Property Get Id() As Long
    Id = mProps.Id
End Property
Friend Property Get IdInventario() As Long
    IdInventario = mProps.IdInventario
End Property



Public Property Get IdInventario() As Long
    IdInventario = Trim$(mProps.IdInventario)
End Property


Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 120 Then
        Err.Raise ERR_INVALIDO, "IVInventario_Descripcion", MSGERR_INVALIDO
    End If
    
    
    If InStr(value, ";") > 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario_Descripcion", _
            "El caracter ';' no es válido para la descripción."
    End If
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, "|") > 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario_Descripcion", _
            "El caracter '|' no es válido para la descripción."
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property



Public Property Let IdGrupo(NumGrupo As Integer, ByVal value As Long)
    mProps.IdGrupo(NumGrupo) = value
    mbooModificado = True
End Property

Public Property Get IdGrupo(NumGrupo As Integer) As Long
    IdGrupo = mProps.IdGrupo(NumGrupo)
End Property


Public Property Let CodGrupo(NumGrupo As Integer, ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodGrupo(NumGrupo) Then
        If Len(value) > 0 Then
            'Actualiza IdGrupo1 también
            sql = "SELECT IdGrupo" & NumGrupo & " FROM IVGrupo" & NumGrupo & " WHERE CodGrupo" & NumGrupo & "='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "IVInventario_CodGrupo" & NumGrupo, MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdGrupo(NumGrupo) = rs.Fields("IdGrupo" & NumGrupo)
            End If
            rs.Close
        Else
            IdGrupo(NumGrupo) = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodGrupo(NumGrupo) = value
    mbooModificado = True
End Property

Public Property Get CodGrupo(NumGrupo As Integer) As String
    CodGrupo = Trim$(mProps.CodGrupo(NumGrupo))
End Property


Public Property Get EsNuevo() As Boolean
    EsNuevo = mbooEsNuevo
End Property

Public Property Get Modificado() As Boolean
    Modificado = mbooModificado
End Property

Friend Property Let Modificado(ByVal value As Boolean)
    mbooModificado = value
End Property



Private Sub Class_Initialize()
    
    
    
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    
    Set mProps.objEmpresa = Nothing
End Sub

