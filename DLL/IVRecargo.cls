VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVRecargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Empresa"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type T_PROPS
    IdRecargo As Long
    CodRecargo As String
    Descripcion As String
    Signo As Integer
    BandAfectaBaseAnexo As Boolean
    BandBaseCero As Boolean
        
    fechagrabado As Date
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mcolRecargoDetalle As Collection  '***Diego

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean



Public Sub Grabar()
    Dim sql As String, rs As Recordset
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    'Verifica si no repite el código
    sql = "SELECT CodRecargo FROM IVRecargo " & _
          "WHERE (CodRecargo='" & mProps.CodRecargo & "') AND " & _
                "(IdRecargo<>" & mProps.IdRecargo & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "IVRecargo", MSGERR_REPITECODIGO
    End If
    rs.Close
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM IVRecargo WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM IVRecargo WHERE IdRecargo=" & mProps.IdRecargo
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)

        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "IVRecargo_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro
    With rs
        !CodRecargo = mProps.CodRecargo
        !Descripcion = mProps.Descripcion
        !Signo = mProps.Signo
        !BandAfectaBaseAnexo = mProps.BandAfectaBaseAnexo
        !BandBaseCero = mProps.BandBaseCero
        !fechagrabado = Now
        .Update
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IdRecargo = !IdRecargo
        .Close
    End With
    Set rs = Nothing
    GrabaRecargoDetalle
    mbooEsNuevo = False
    mbooModificado = False
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_RECARGO
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String
    Dim rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM IVRecargo WHERE CodRecargo='" & key & "'"
    Else
        sql = "SELECT * FROM IVRecargo WHERE IdRecargo=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'Recupera los datos de Usuario
    With rs
        mProps.IdRecargo = !IdRecargo
        mProps.CodRecargo = !CodRecargo
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!Signo) Then mProps.Signo = !Signo
        If Not IsNull(!BandAfectaBaseAnexo) Then mProps.BandAfectaBaseAnexo = !BandAfectaBaseAnexo
        If Not IsNull(!BandBaseCero) Then mProps.BandBaseCero = !BandBaseCero
        
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    'Recupera  el detalle
    RecuperaDetalle
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset

    'Verifica si existe en IVKardexRecargo
    sql = "SELECT IdRecargo FROM IVKardexRecargo WHERE IdRecargo=" & mProps.IdRecargo
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "IVRecargo.Eliminar", MSGERR_NOELIMINABLE
    End If
    'Verifica si existe en GNTransRecargo
    sql = "SELECT IdRecargo FROM GNTransRecargo WHERE IdRecargo=" & mProps.IdRecargo
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "IVRecargo.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing
    
    'Elimina el registro
    sql = "DELETE FROM IVRecargo WHERE CodRecargo='" & mProps.CodRecargo & "'"
    mProps.objEmpresa.Execute sql, True

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_RECARGO
End Sub

Public Function Clone() As IVRecargo
    Dim nuevo As IVRecargo
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New IVRecargo
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdRecargo = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Get IdRecargo() As Long
    IdRecargo = mProps.IdRecargo
End Property


Public Property Let CodRecargo(ByVal value As String)
    If Len(value) > 5 Then
        Err.Raise ERR_INVALIDO, "IVRecargo_CodRecargo", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos

    mProps.CodRecargo = value
    mbooModificado = True
End Property

Public Property Get CodRecargo() As String
    CodRecargo = Trim$(mProps.CodRecargo)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "IVRecargo_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let Signo(ByVal value As Integer)
    If value < -1 Or value > 1 Then
        Err.Raise ERR_INVALIDO, "IVRecargo_Signo", MSGERR_INVALIDO
    End If

    mProps.Signo = value
    mbooModificado = True
End Property

Public Property Get Signo() As Integer
    Signo = mProps.Signo
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

Friend Property Let Modificado(ByVal value As Boolean)
    mbooModificado = value
End Property


Private Sub Class_Initialize()
    'Valores predeterminados
    With mProps
        .Signo = 1
    End With
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Private Sub BorrarDetalle()
    'Borra  toda  la informacion de la coleccion
     If mcolRecargoDetalle Is Nothing Then Set mcolRecargoDetalle = New Collection
    With mcolRecargoDetalle
        Do While .Count > 0
            .Remove .Count
        Loop
    End With
    mbooModificado = True
End Sub



Private Sub RecuperaDetalle()
    Dim sql As String, rs As Recordset, i As Long, mod_antes As Boolean
    Dim nuevo As IVRecargoDetalle
    mod_antes = mbooModificado      'Guarda el estado de Modificado
    If mcolRecargoDetalle Is Nothing Then Set mcolRecargoDetalle = New Collection
    
    BorrarDetalle
    
    'Recupera los IVKardex
    sql = "SELECT * FROM IVRecargoDetalle WHERE IdRecargo=" & mProps.IdRecargo & " ORDER BY Orden"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            'If !IdRecargo > 0 Then                 'Para asegurar el Bug de ADO
                Set nuevo = New IVRecargoDetalle
                mcolRecargoDetalle.Add nuevo
                'mcolIVKardex.Add nuevo
                i = mcolRecargoDetalle.Count
                Set IVRecargoDetalle(i).IVRecargo = Me
                IVRecargoDetalle(i).Recuperar rs
            'End If
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    mbooModificado = mod_antes      'Para que no cambie el estado de Modificado
End Sub

Public Property Get IVRecargoDetalle(ix As Long) As IVRecargoDetalle
    If mcolRecargoDetalle Is Nothing Then RecuperaDetalle   'posilbe recursividad
    Set IVRecargoDetalle = mcolRecargoDetalle.item(ix)
End Property

Public Property Get NumRecargoDetalle() As Long
    If mcolRecargoDetalle Is Nothing Then
        NumRecargoDetalle = 0
    Else
        NumRecargoDetalle = mcolRecargoDetalle.Count
    End If
    'If ObtieneUltimo(False) Then PrecioUltimoEgreso = mUEgreso.Precio
End Property


'--------------- Manejo de coleccion de RecargoDetalle
Public Function AddDetalleRecargo() As Long
    Dim nuevo As IVRecargoDetalle
    
    If mcolRecargoDetalle Is Nothing Then RecuperaDetalle
    Set nuevo = New IVRecargoDetalle
    Set nuevo.IVRecargo = Me    'Me  olvide  del padre
    'Proporciona valores predeterminados
    'nuevo.CodRecargo = Me.CodRecargo
    nuevo.IdRecargo = Me.IdRecargo
    mcolRecargoDetalle.Add nuevo
    AddDetalleRecargo = mcolRecargoDetalle.Count   'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function

Private Sub GrabaRecargoDetalle()
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As IVRecargoDetalle, sql As String
    'Elimina  todos  los registros
    sql = "Delete IVRD " & _
          "From IVRecargo IVR " & _
          "INNER JOIN IVRecargoDetalle IVRD ON IVR.IdRecargo = IVRD.IdRecargo " & _
          "Where IVRD.IdRecargo = " & mProps.IdRecargo
    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    If Not mcolRecargoDetalle Is Nothing Then
        sql = "SELECT * FROM IVRecargoDetalle WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        For i = 1 To mcolRecargoDetalle.Count
            rs1.AddNew
            Set obj = mcolRecargoDetalle.item(i)
            rs1!IdRecargo = obj.IdRecargo
            rs1!Valor = obj.Valor
            rs1!NumPCGrupo = obj.NumPCGrupo
            rs1!idPcGrupo = obj.idPcGrupo
            rs1!Orden = obj.Orden
            rs1.Update
        Next i
    End If
    Set rs1 = Nothing
    Set obj = Nothing
End Sub

Public Sub RemoveRecargoDetalle(ByVal ix As Long, Optional obj As IVRecargoDetalle)
    If mcolRecargoDetalle Is Nothing Then RecuperaDetalle
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolRecargoDetalle.Count
            If mcolRecargoDetalle.item(ix) Is obj Then
                mcolRecargoDetalle.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolRecargoDetalle.Remove ix
    End If
    
    'mcolDetalleFamilia.Remove ix
    mbooModificado = True
End Sub


Public Property Let BandAfectaBaseAnexo(ByVal value As Boolean)
    mProps.BandAfectaBaseAnexo = value
    mbooModificado = True
End Property

Public Property Get BandAfectaBaseAnexo() As Boolean
    BandAfectaBaseAnexo = mProps.BandAfectaBaseAnexo
End Property

Public Property Let BandBaseCero(ByVal value As Boolean)
    mProps.BandBaseCero = value
    mbooModificado = True
End Property

Public Property Get BandBaseCero() As Boolean
    BandBaseCero = mProps.BandBaseCero
End Property

