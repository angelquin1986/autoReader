VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVDescuento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***jeaa  30/12/2004
Private Type T_PROPS
    IdDescuento As Long
    CodDescuento As String
    Descripcion As String
    BandValida As Boolean
    BandxCliente As Boolean
    BandxItem As Boolean
    BandCaducidad As Boolean
    FechaDesde As Date
    FechaHasta As Date
    HoraDesde As Date
    HoraHasta As Date
    NumPCGrupo  As Byte
    NumIVGrupo  As Byte
    fechagrabado As Date
    BandAcumula As Boolean
    BandTodoCliente As Boolean
    IsCopy As Boolean
    
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mcolDescuentoDetallePC As Collection  '*** jeaa
Private mcolDescuentoDetalleIV As Collection  '*** jeaa
Private mcolDescuentoDetalleFC As Collection  '*** jeaa

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Public Sub Grabar()
    Dim sql As String, rs As Recordset, rs1  As Recordset, i As Long, obj As Object, rspc  As Recordset, pc As PCProvCli
''    Dim rs1 As Recordset, i As Long, obj As IVDescuentoDetallePC, rspc  As Recordset, pc As PCProvCli
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    
    'Verifica si no repite el código
    sql = "SELECT codDescuento FROM IVDescuento "
    sql = sql & " WHERE codDescuento='" & mProps.CodDescuento & "'"
    sql = sql & " and iddescuento<> " & mProps.IdDescuento

    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "IVDescuento", MSGERR_REPITECODIGO
    End If
    rs.Close
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM IVDescuento WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM IVDescuento WHERE IdDescuento=" & mProps.IdDescuento
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)

        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "IVDescuento_Grabar", MSGERR_MODIFICADO
        End If
    End If
    
    
    ''Verifica que no exista mas de 1 descuentos aplicados a un mismo cliente

        If Not mcolDescuentoDetallePC Is Nothing Then
        sql = "SELECT * FROM IVDescuentoDetallePC WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        If mProps.BandxCliente Then
            For i = 1 To mcolDescuentoDetallePC.Count
                Set obj = mcolDescuentoDetallePC.item(i)
                sql = " select coddescuento, descripcion ,idprovcli from IVDescuento IVR inner join IVDescuentoDetallePC p on IVR.IdDescuento = p.iddescuento where idprovcli=" & obj.IdProvcli
                'sql = sql & " and ivr.bandvalida=1"
                sql = sql & " and ivr.iddescuento<>" & Me.IdDescuento
                Set rspc = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
                If rspc.RecordCount > 0 Then
                    If rspc.Fields(2) <> 0 Then
                        Set pc = mProps.objEmpresa.RecuperaPCProvCliQuick(rspc.Fields(2))
                        MsgBox " Para el cliente " & pc.RUC & " - " & pc.Nombre & Chr(13) & " Ya existen asignados descuentos en el código : " & rspc.Fields(0) & " - " & rspc.Fields(1)
                        Set pc = Nothing
                        Exit Sub
                    End If
                End If
            Next i
        End If
    End If
    Set rs1 = Nothing
    Set obj = Nothing
    Set pc = Nothing
    

    'Graba el registro
    With rs
        'jeaa 05/10/2010
        !CodDescuento = mProps.CodDescuento
        !Descripcion = mProps.Descripcion
        !BandValida = mProps.BandValida
        !BandxCliente = mProps.BandxCliente
        !BandTodoCliente = mProps.BandTodoCliente
        !BandxItem = mProps.BandxItem
        !BandCaducidad = mProps.BandCaducidad
        !FechaDesde = mProps.FechaDesde
        !FechaHasta = mProps.FechaHasta
        !HoraDesde = mProps.HoraDesde
        !HoraHasta = mProps.HoraHasta
        !NumPCGrupo = mProps.NumPCGrupo
        !NumIVGrupo = mProps.NumIVGrupo
        !BandAcumula = mProps.BandAcumula
        !BandTodoCliente = mProps.BandTodoCliente
    
        
        !fechagrabado = Now
        .Update
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IdDescuento = !IdDescuento
        .Close
    End With
    Set rs = Nothing
    GrabaDescuentoDetallePC
    GrabaDescuentoDetalleIV
    GrabaDescuentoDetalleFC
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_DES", "Descto Nuevo: " & _
                                                         mProps.CodDescuento, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_DES", "Descto Modificado: " & _
                                                         mProps.CodDescuento, "DT"
    End If
    
    mbooEsNuevo = False
    mbooModificado = False
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVDESCUENTO
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String
    Dim rs As Recordset
    Dim v As Variant

    v = Split(key, ",")
        If VarType(key) = vbString Then
            sql = "SELECT *"
            sql = sql & " FROM IVDescuento "
            sql = sql & " WHERE CodDescuento='" & key & "'"
        Else
            sql = "SELECT *"
            sql = sql & " FROM IVDescuento "
            sql = sql & " WHERE Iddescuento=" & key
        End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'Recupera los datos de Usuario
    With rs
        mProps.IdDescuento = !IdDescuento
        If Not IsNull(!CodDescuento) Then mProps.CodDescuento = !CodDescuento
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!BandxCliente) Then mProps.BandxCliente = !BandxCliente
        If Not IsNull(!BandxItem) Then mProps.BandxItem = !BandxItem
        If Not IsNull(!BandCaducidad) Then mProps.BandCaducidad = !BandCaducidad
        If Not IsNull(!NumPCGrupo) Then mProps.NumPCGrupo = !NumPCGrupo
        If Not IsNull(!NumIVGrupo) Then mProps.NumIVGrupo = !NumIVGrupo
        mProps.FechaDesde = !FechaDesde
        mProps.FechaHasta = !FechaHasta
        mProps.HoraDesde = !HoraDesde
        mProps.HoraHasta = !HoraHasta
        If Not IsNull(!BandAcumula) Then mProps.BandAcumula = !BandAcumula
        If Not IsNull(!BandTodoCliente) Then mProps.BandTodoCliente = !BandTodoCliente
        mProps.fechagrabado = !fechagrabado
        
        .Close
    End With
    'Recupera  el detalle
    RecuperaDetallePC
    RecuperaDetalleIV
    RecuperaDetalleFC
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
  
    'Verifica si existe items registrados
    sql = "select * FROM gncomprobante  WHERE Iddescuento=" & mProps.IdDescuento
'    Set rs = mProps.objEmpresa.OpenRecordset(sql)
'    If Not rs.EOF Then
'        Err.Raise ERR_NOELIMINABLE, "Tiene registro relacionados ", MSGERR_NOELIMINABLE
'    End If
'    Set rs = Nothing

    sql = "DELETE FROM IVDescuentoDetallepc WHERE Iddescuento=" & mProps.IdDescuento
    mProps.objEmpresa.Execute sql, True
  
    sql = "DELETE FROM IVDescuentoDetallefc WHERE Iddescuento=" & mProps.IdDescuento
    mProps.objEmpresa.Execute sql, True
  
    sql = "DELETE FROM IVDescuentoDetalleiv WHERE Iddescuento=" & mProps.IdDescuento
    mProps.objEmpresa.Execute sql, True
  
  
  
    'Elimina el registro
    sql = "DELETE FROM IVDescuento WHERE Iddescuento=" & mProps.IdDescuento
    mProps.objEmpresa.Execute sql, True

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVDESCUENTO
End Sub

Public Function Clone() As IVDescuento
    Dim nuevo As IVDescuento, ivdescPC As IVDescuentoDetallePC, i As Long
    Dim ivdescIV As IVDescuentoDetalleIV, ivdescFC As IVDescuentoDetalleFC
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New IVDescuento
    nuevo.Props = mProps
    
    'Copia los contactos al nuevo objeto
    For i = 1 To mcolDescuentoDetallePC.Count
        Set ivdescPC = IVDescuentoDetallePC(i).Clone
        nuevo.AddDetalleDescuentoPC
        
        nuevo.IVDescuentoDetallePC(i).CodProvcli = IVDescuentoDetallePC(i).CodProvcli
        nuevo.IVDescuentoDetallePC(i).Nombre = IVDescuentoDetallePC(i).Nombre
        nuevo.IVDescuentoDetallePC(i).Orden = IVDescuentoDetallePC(i).Orden
        
    Next i
    
    For i = 1 To mcolDescuentoDetalleIV.Count
        Set ivdescIV = IVDescuentoDetalleIV(i).Clone
        nuevo.AddDetalleDescuentoIV
        nuevo.IVDescuentoDetalleIV(i).CodInventario = IVDescuentoDetalleIV(i).CodInventario
        nuevo.IVDescuentoDetalleIV(i).Orden = IVDescuentoDetalleIV(i).Orden
        nuevo.IVDescuentoDetalleIV(i).Descuento = IVDescuentoDetalleIV(i).Descuento
        nuevo.IVDescuentoDetalleIV(i).Descripcion = IVDescuentoDetalleIV(i).Descripcion
        nuevo.IVDescuentoDetalleIV(i).Unidad = IVDescuentoDetalleIV(i).Unidad
        
        
    Next i
    
    For i = 1 To mcolDescuentoDetalleFC.Count
        Set ivdescFC = IVDescuentoDetalleFC(i).Clone
        nuevo.AddDetalleDescuentoFC
        
        nuevo.IVDescuentoDetalleFC(i).CodForma = IVDescuentoDetalleFC(i).CodForma
        nuevo.IVDescuentoDetalleFC(i).Orden = IVDescuentoDetalleFC(i).Orden
        nuevo.IVDescuentoDetalleFC(i).Penal = IVDescuentoDetalleFC(i).Penal
        
    Next i
    
    
    
    
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdDescuento = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Get IdDescuento() As Long
    IdDescuento = mProps.IdDescuento
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

Public Property Let Modificado(ByVal value As Boolean)
    mbooModificado = value
    
End Property


Private Sub Class_Initialize()
    'Valores predeterminados
    mbooEsNuevo = True
    mbooModificado = False
    mProps.BandValida = True
    mProps.BandCaducidad = False
    mProps.BandxCliente = False
    mProps.BandxItem = False
    mProps.BandAcumula = False
    mProps.BandTodoCliente = False

End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Property Let CodDescuento(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "IVDescuento_CodDescuento", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodDescuento = value
    mbooModificado = True
End Property

Public Property Get CodDescuento() As String
    CodDescuento = Trim$(mProps.CodDescuento)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 254 Then
        Err.Raise ERR_INVALIDO, "IVDescuento_Descripcion", MSGERR_INVALIDO
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

Public Property Let BandCaducidad(ByVal value As Boolean)
    mProps.BandCaducidad = value
    mbooModificado = True
End Property

Public Property Get BandCaducidad() As Boolean
    BandCaducidad = mProps.BandCaducidad
End Property

Public Property Let FechaDesde(value As Date)
        mProps.FechaDesde = value
        mbooModificado = True
End Property

Public Property Get FechaDesde() As Date
    FechaDesde = mProps.FechaDesde
End Property

Public Property Let FechaHasta(value As Date)
        mProps.FechaHasta = value
        mbooModificado = True
End Property

Public Property Get FechaHasta() As Date
    FechaHasta = mProps.FechaHasta
End Property


Public Property Let HoraDesde(value As Date)
        mProps.HoraDesde = value
        mbooModificado = True
End Property

Public Property Get HoraDesde() As Date
    HoraDesde = mProps.HoraDesde
End Property

Public Property Let HoraHasta(value As Date)
        mProps.HoraHasta = value
        mbooModificado = True
End Property

Public Property Get HoraHasta() As Date
    HoraHasta = mProps.HoraHasta
End Property


Public Property Let NumPCGrupo(ByVal value As Byte)
    mProps.NumPCGrupo = value
    mbooModificado = True
End Property

Public Property Get NumPCGrupo() As Byte
    NumPCGrupo = mProps.NumPCGrupo
End Property

Public Property Let NumIVGrupo(ByVal value As Byte)
    mProps.NumIVGrupo = value
    mbooModificado = True
End Property

Public Property Get NumIVGrupo() As Byte
    NumIVGrupo = mProps.NumIVGrupo
End Property

'*********** coleccion de GrupoPC
Private Sub GrabaDescuentoDetallePC()
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As IVDescuentoDetallePC, sql As String
    
    Dim rspc As Recordset, pc As PCProvCli
    ' primero revisa si hay otro descuento asignado al mismo cliente
    
    'Elimina  todos  los registros
    sql = "Delete IVRD " & _
          "From IVDescuento IVR " & _
          "INNER JOIN IVDescuentoDetallePC IVRD ON IVR.IdDescuento = IVRD.IdDescuento " & _
          "Where IVRD.IdDescuento = " & mProps.IdDescuento
    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    
''            sql = "SELECT  IdDescuento ,NumPCGrupo  ,IdPCGrupo   ,IdPRovCli   , 0 as Orden FROM IVDescuentoDetallePC WHERE 1=0"
''            Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
    
    
    If Not mcolDescuentoDetallePC Is Nothing Then
    
        For i = 1 To mcolDescuentoDetallePC.Count
        
            'sql = "SELECT  IdDescuento ,NumPCGrupo  ,IdPCGrupo   ,IdPRovCli   , 0 as Orden "
            sql = "SELECT  * "
            sql = sql & " FROM IVDescuentoDetallePC "
            sql = sql & " WHERE 1=0"
            Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        
                
            rs1.AddNew
            Set obj = mcolDescuentoDetallePC.item(i)
            If obj.IdProvcli <> 0 Then
                If mProps.BandxCliente Then
                    sql = " select coddescuento, descripcion ,idprovcli "
                    sql = sql & " from IVDescuento IVR inner join IVDescuentoDetallePC p "
                    sql = sql & " on IVR.IdDescuento = p.iddescuento "
                    sql = sql & " where idprovcli=" & obj.IdProvcli
                    sql = sql & " and ivr.iddescuento<>" & Me.IdDescuento
                    Set rspc = mProps.objEmpresa.OpenRecordset(sql)
                        If rspc.RecordCount > 0 Then
                                If rspc.Fields(2) <> 0 Then
                                    Set pc = mProps.objEmpresa.RecuperaPCProvCliQuick(rspc.Fields(2))
                                    MsgBox " Para el cliente " & pc.RUC & " - " & pc.Nombre & Chr(13) & " Ya existen descuentos asignados en el descuento: " & rspc.Fields(0) & " - " & rspc.Fields(1)
                                    Set pc = Nothing
                                    Exit Sub
                                End If
                        End If
                End If
            End If
            rs1!IdDescuento = mProps.IdDescuento
            rs1!NumPCGrupo = obj.NumPCGrupo
            rs1!idPcGrupo = obj.idPcGrupo
            rs1!IdProvcli = obj.IdProvcli
            rs1!Orden = obj.Orden
            rs1.Update
            'End If
        Next i
    End If
    Set rs1 = Nothing
    Set obj = Nothing
End Sub



Private Sub BorrarDetallePC()
    'Borra  toda  la informacion de la coleccion
     If mcolDescuentoDetallePC Is Nothing Then Set mcolDescuentoDetallePC = New Collection
    With mcolDescuentoDetallePC
        Do While .Count > 0
            .Remove .Count
        Loop
    End With
    mbooModificado = True
End Sub



Private Sub RecuperaDetallePC()
    Dim sql As String, rs As Recordset, i As Long, mod_antes As Boolean
    Dim nuevo As IVDescuentoDetallePC
    mod_antes = mbooModificado      'Guarda el estado de Modificado
    If mcolDescuentoDetallePC Is Nothing Then Set mcolDescuentoDetallePC = New Collection
    
    BorrarDetallePC
    
    'Recupera los IVKardex
    sql = "SELECT * FROM IVDescuentoDetallePC i left join pcprovcli p on i.idprovcli=p.idprovcli WHERE IdDescuento=" & mProps.IdDescuento & " ORDER BY p.nombre "
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            Set nuevo = New IVDescuentoDetallePC
            mcolDescuentoDetallePC.Add nuevo
            i = mcolDescuentoDetallePC.Count
            Set IVDescuentoDetallePC(i).IVDescuento = Me
            IVDescuentoDetallePC(i).Recuperar rs
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    mbooModificado = mod_antes      'Para que no cambie el estado de Modificado
End Sub

Public Property Get IVDescuentoDetallePC(ix As Long) As IVDescuentoDetallePC
    If mcolDescuentoDetallePC Is Nothing Then RecuperaDetallePC   'posilbe recursividad
    Set IVDescuentoDetallePC = mcolDescuentoDetallePC.item(ix)
End Property

Public Property Get NumDescuentoDetallePC() As Long
    If mcolDescuentoDetallePC Is Nothing Then
        NumDescuentoDetallePC = 0
    Else
        NumDescuentoDetallePC = mcolDescuentoDetallePC.Count
    End If
End Property


'--------------- Manejo de coleccion de DescuentoDetallePC
Public Function AddDetalleDescuentoPC() As Long
    Dim nuevo As IVDescuentoDetallePC
    
    If mcolDescuentoDetallePC Is Nothing Then RecuperaDetallePC
    Set nuevo = New IVDescuentoDetallePC
    Set nuevo.IVDescuento = Me    'Me  olvide  del padre
    'Proporciona valores predeterminados
    nuevo.IdDescuento = Me.IdDescuento
    mcolDescuentoDetallePC.Add nuevo
    AddDetalleDescuentoPC = mcolDescuentoDetallePC.Count   'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function

Public Sub RemoveDescuentoDetallePC(ByVal ix As Long, Optional obj As IVDescuentoDetallePC)
    If mcolDescuentoDetallePC Is Nothing Then RecuperaDetallePC
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolDescuentoDetallePC.Count
            If mcolDescuentoDetallePC.item(ix) Is obj Then
                mcolDescuentoDetallePC.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolDescuentoDetallePC.Remove ix
    End If
    
    mbooModificado = True
End Sub

'*********** coleccion de GrupoIV
Private Sub GrabaDescuentoDetalleIV()
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As IVDescuentoDetalleIV, sql As String
    'Elimina  todos  los registros
    sql = "Delete IVRD " & _
          "From IVDescuento IVR " & _
          "INNER JOIN IVDescuentoDetalleIV IVRD ON IVR.IdDescuento = IVRD.IdDescuento " & _
          "Where IVRD.IdDescuento = " & mProps.IdDescuento
    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    If Not mcolDescuentoDetalleIV Is Nothing Then
        sql = "SELECT * FROM IVDescuentoDetalleIV WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        For i = 1 To mcolDescuentoDetalleIV.Count
            rs1.AddNew
            Set obj = mcolDescuentoDetalleIV.item(i)
            rs1!IdDescuento = mProps.IdDescuento
            rs1!NumIVGrupo = obj.NumIVGrupo
            rs1!IDIVGrupo = obj.IDIVGrupo
            rs1!IdInventario = obj.IdInventario
            rs1!Descuento = obj.Descuento
            rs1!CantMin = obj.CantMin
            
            rs1!Orden = obj.Orden
            
            rs1.Update
        Next i
    End If
    Set rs1 = Nothing
    Set obj = Nothing
End Sub

Public Sub BorrarDetalleIV()
    'Borra  toda  la informacion de la coleccion
     If mcolDescuentoDetalleIV Is Nothing Then Set mcolDescuentoDetalleIV = New Collection
    With mcolDescuentoDetalleIV
        Do While .Count > 0
            .Remove .Count
        Loop
    End With
    mbooModificado = True
End Sub



Private Sub RecuperaDetalleIV()
    Dim sql As String, rs As Recordset, i As Long, mod_antes As Boolean
    Dim nuevo As IVDescuentoDetalleIV
    mod_antes = mbooModificado      'Guarda el estado de Modificado
    If mcolDescuentoDetalleIV Is Nothing Then Set mcolDescuentoDetalleIV = New Collection
    
    BorrarDetalleIV
    
    'Recupera los IVKardex
    sql = "SELECT * FROM IVDescuentoDetalleIV i left join ivinventario iv on i.idinventario=iv.idinventario WHERE IdDescuento=" & mProps.IdDescuento & " ORDER BY iv.descripcion "
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            Set nuevo = New IVDescuentoDetalleIV
            mcolDescuentoDetalleIV.Add nuevo
            i = mcolDescuentoDetalleIV.Count
            Set IVDescuentoDetalleIV(i).IVDescuento = Me
            IVDescuentoDetalleIV(i).Recuperar rs
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    mbooModificado = mod_antes      'Para que no cambie el estado de Modificado
End Sub

Public Property Get IVDescuentoDetalleIV(ix As Long) As IVDescuentoDetalleIV
    If mcolDescuentoDetalleIV Is Nothing Then RecuperaDetalleIV   'posilbe recursividad
    Set IVDescuentoDetalleIV = mcolDescuentoDetalleIV.item(ix)
End Property

Public Property Get NumDescuentoDetalleIV() As Long
    If mcolDescuentoDetalleIV Is Nothing Then
        NumDescuentoDetalleIV = 0
    Else
        NumDescuentoDetalleIV = mcolDescuentoDetalleIV.Count
    End If
End Property


'--------------- Manejo de coleccion de DescuentoDetalleIV
Public Function AddDetalleDescuentoIV() As Long
    Dim nuevo As IVDescuentoDetalleIV
    
    If mcolDescuentoDetalleIV Is Nothing Then RecuperaDetalleIV
    Set nuevo = New IVDescuentoDetalleIV
    Set nuevo.IVDescuento = Me    'Me  olvide  del padre
    'Proporciona valores predeterminados
    nuevo.IdDescuento = Me.IdDescuento
    mcolDescuentoDetalleIV.Add nuevo
    AddDetalleDescuentoIV = mcolDescuentoDetalleIV.Count   'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function

Public Sub RemoveDescuentoDetalleIV(ByVal ix As Long, Optional obj As IVDescuentoDetalleIV)
    If mcolDescuentoDetalleIV Is Nothing Then RecuperaDetalleIV
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolDescuentoDetalleIV.Count
            If mcolDescuentoDetalleIV.item(ix) Is obj Then
                mcolDescuentoDetalleIV.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolDescuentoDetalleIV.Remove ix
    End If
    
    mbooModificado = True
End Sub



'*********** coleccion de GrupoFC
Private Sub GrabaDescuentoDetalleFC()
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As IVDescuentoDetalleFC, sql As String
    'Elimina  todos  los registros
    sql = "Delete IVRD " & _
          "From IVDescuento IVR " & _
          "INNER JOIN IVDescuentoDetalleFC IVRD ON IVR.IdDescuento = IVRD.IdDescuento " & _
          "Where IVRD.IdDescuento = " & mProps.IdDescuento
    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    If Not mcolDescuentoDetalleFC Is Nothing Then
        sql = "SELECT * FROM IVDescuentoDetalleFC WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        For i = 1 To mcolDescuentoDetalleFC.Count
            rs1.AddNew
            Set obj = mcolDescuentoDetalleFC.item(i)
            rs1!IdDescuento = mProps.IdDescuento
            rs1!Descuento = obj.Descuento
            rs1!Penal = obj.Penal
            rs1!IDForma = obj.IDForma
            rs1!Orden = obj.Orden
            rs1.Update
        Next i
    End If
    Set rs1 = Nothing
    Set obj = Nothing
End Sub



Private Sub BorrarDetalleFC()
    'Borra  toda  la informacion de la coleccion
     If mcolDescuentoDetalleFC Is Nothing Then Set mcolDescuentoDetalleFC = New Collection
    With mcolDescuentoDetalleFC
        Do While .Count > 0
            .Remove .Count
        Loop
    End With
    mbooModificado = True
End Sub



Private Sub RecuperaDetalleFC()
    Dim sql As String, rs As Recordset, i As Long, mod_antes As Boolean
    Dim nuevo As IVDescuentoDetalleFC
    mod_antes = mbooModificado      'Guarda el estado de Modificado
    If mcolDescuentoDetalleFC Is Nothing Then Set mcolDescuentoDetalleFC = New Collection
    
    BorrarDetalleFC
    
    'Recupera los IVKardex
    sql = "SELECT * FROM IVDescuentoDetalleFC WHERE IdDescuento=" & mProps.IdDescuento & " ORDER BY Orden"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            Set nuevo = New IVDescuentoDetalleFC
            mcolDescuentoDetalleFC.Add nuevo
            i = mcolDescuentoDetalleFC.Count
            Set IVDescuentoDetalleFC(i).IVDescuento = Me
            IVDescuentoDetalleFC(i).Recuperar rs
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    mbooModificado = mod_antes      'Para que no cambie el estado de Modificado
End Sub

Public Property Get IVDescuentoDetalleFC(ix As Long) As IVDescuentoDetalleFC
    If mcolDescuentoDetalleFC Is Nothing Then RecuperaDetalleFC   'posilbe recursividad
    Set IVDescuentoDetalleFC = mcolDescuentoDetalleFC.item(ix)
End Property

Public Property Get NumDescuentoDetalleFC() As Long
    If mcolDescuentoDetalleFC Is Nothing Then
        NumDescuentoDetalleFC = 0
    Else
        NumDescuentoDetalleFC = mcolDescuentoDetalleFC.Count
    End If
End Property


'--------------- Manejo de coleccion de DescuentoDetalleFC
Public Function AddDetalleDescuentoFC() As Long
    Dim nuevo As IVDescuentoDetalleFC
    
    If mcolDescuentoDetalleFC Is Nothing Then RecuperaDetalleFC
    Set nuevo = New IVDescuentoDetalleFC
    Set nuevo.IVDescuento = Me    'Me  olvide  del padre
    nuevo.IdDescuento = Me.IdDescuento
    mcolDescuentoDetalleFC.Add nuevo
    AddDetalleDescuentoFC = mcolDescuentoDetalleFC.Count   'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function

Public Sub RemoveDescuentoDetalleFC(ByVal ix As Long, Optional obj As IVDescuentoDetalleFC)
    If mcolDescuentoDetalleFC Is Nothing Then RecuperaDetalleFC
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolDescuentoDetalleFC.Count
            If mcolDescuentoDetalleFC.item(ix) Is obj Then
                mcolDescuentoDetalleFC.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolDescuentoDetalleFC.Remove ix
    End If
    
    mbooModificado = True
End Sub


Public Property Let BandxCliente(ByVal value As Boolean)
    mProps.BandxCliente = value
    mbooModificado = True
End Property

Public Property Get BandxCliente() As Boolean
    BandxCliente = mProps.BandxCliente
End Property

Public Property Let BandTodoCliente(ByVal value As Boolean)
    mProps.BandTodoCliente = value
    mbooModificado = True
End Property

Public Property Get BandTodoCliente() As Boolean
    BandTodoCliente = mProps.BandTodoCliente
End Property


Public Property Let BandxItem(ByVal value As Boolean)
    mProps.BandxItem = value
    mbooModificado = True
End Property

Public Property Get BandxItem() As Boolean
    BandxItem = mProps.BandxItem
End Property

Public Function RecuperaDescuentoDetallePC(ix As Long) As IVDescuentoDetallePC
    Set RecuperaDescuentoDetallePC = mcolDescuentoDetallePC.item(ix)
End Function

Public Function RecuperaDescuentoDetalleIV(ix As Long) As IVDescuentoDetalleIV
    Set RecuperaDescuentoDetalleIV = mcolDescuentoDetalleIV.item(ix)
End Function

Public Property Let BandAcumula(ByVal value As Boolean)
    mProps.BandAcumula = value
    mbooModificado = True
End Property

Public Property Get BandAcumula() As Boolean
    BandAcumula = mProps.BandAcumula
End Property

Public Function RegTodoXGrupo(ByVal CodsGrupo As String) As Recordset
Dim sql As String
Dim rs As Recordset
    sql = "select ivg.Idgrupo" & mProps.NumIVGrupo & " as idIvg,iv.codinventario,iv.Descripcion,0 as dsct "
    sql = sql & "from ivinventario iv"
    sql = sql & " inner join ivgrupo" & mProps.NumIVGrupo & " ivg ON ivg.idgrupo" & mProps.NumIVGrupo & " = iv.idgrupo" & mProps.NumIVGrupo
    sql = sql & " WHERE iv.bandvalida=1 and  ivg.codgrupo" & mProps.NumIVGrupo & " IN (" & CodsGrupo & ")"
    sql = sql & " order by ivg.codgrupo" & mProps.NumIVGrupo & ",iv.descripcion"
    Set rs = Me.Empresa.OpenRecordset(sql)
    If Not rs Is Nothing Then
        Set RegTodoXGrupo = rs
    End If
    Exit Function
End Function

'AUC agrega unidad Megalimpio
Public Function RegTodoXGrupoUnidad(ByVal CodsGrupo As String) As Recordset
Dim sql As String
Dim rs As Recordset
    sql = "select ivg.Idgrupo" & mProps.NumIVGrupo & " as idIvg,iv.codinventario,iv.Descripcion,ivu.codunidad,0 as dsct "
    sql = sql & "from ivinventario iv left join ivunidad  ivu on ivu.idunidad = iv.idunidad"
   sql = sql & " inner join ivgrupo" & mProps.NumIVGrupo & " ivg ON ivg.idgrupo" & mProps.NumIVGrupo & " = iv.idgrupo" & mProps.NumIVGrupo
    sql = sql & " WHERE iv.bandvalida=1 and  ivg.codgrupo" & mProps.NumIVGrupo & " IN (" & CodsGrupo & ")"
    sql = sql & " order by ivg.codgrupo" & mProps.NumIVGrupo & ",iv.descripcion"
    Set rs = Me.Empresa.OpenRecordset(sql)
    If Not rs Is Nothing Then
        Set RegTodoXGrupoUnidad = rs
    End If
    Exit Function
End Function

Public Property Let IsCopy(ByVal value As Boolean)
    mProps.IsCopy = value
End Property

Public Property Get IsCopy() As Boolean
    IsCopy = Trim$(mProps.IsCopy)
End Property

Public Function ObtieneIdUltimoDescuento() As Long
Dim sql As String
Dim rs As Recordset
    sql = "select top 1 iddescuento from ivdescuento order by iddescuento desc"

    Set rs = Me.Empresa.OpenRecordset(sql)
    If Not rs Is Nothing Then
        ObtieneIdUltimoDescuento = rs.Fields(0)
    End If
    Exit Function

End Function
