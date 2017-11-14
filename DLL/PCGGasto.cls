VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCGGasto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdPCGGasto As Long
    Codigo As String
    FechaGrabado As Date
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mcolPCGGastoDetallePC As Collection

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Public Sub Grabar()
    Dim sql As String, rs As Recordset
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
'    sql = "SELECT idpcggasto,idpcgrupo FROM pcggasto"
'    sql = sql & " WHERE idpcggasto<> " & mProps.IdPCGGasto
'    Set rs = mProps.objEmpresa.OpenRecordset(sql)
'    If Not rs.EOF Then      'Si ya existe el codigo
'        Err.Raise ERR_REPITECODIGO, "PCGGasto", MSGERR_REPITECODIGO
'    End If
'    rs.Close
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM pcggasto WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM pcggasto WHERE Idpcggasto=" & mProps.IdPCGGasto
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)

        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "IVDescuento_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro
    With rs
        !Codigo = mProps.Codigo
        !FechaGrabado = Now
        .Update
        .Move 0             'Para actualizar
        mProps.FechaGrabado = !FechaGrabado
        mProps.IdPCGGasto = !IdPCGGasto
        .Close
    End With
    Set rs = Nothing
    GrabaPCGastoDetallePC
   ' GrabaDescuentoDetalleIV
    'GrabaDescuentoDetalleFC
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_DES", "PCGGasto Nuevo: " & _
                                                         mProps.Codigo, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_DES", "PCGGasto Modificado: " & _
                                                         mProps.Codigo, "DT"
    End If
    mbooEsNuevo = False
    mbooModificado = False
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PCGGASTO
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String
    Dim rs As Recordset
    Dim v As Variant

    'v = Split(key, ",")
        If VarType(key) = vbString Then
            sql = "SELECT *"
            sql = sql & " FROM PCGGasto "
            sql = sql & " WHERE idpcgGasto='" & key & "'"
        Else
            sql = "SELECT *"
            sql = sql & " FROM PCGGasto "
            sql = sql & " WHERE idpcgGasto=" & key
        End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    'Recupera los datos de Usuario
    With rs
        mProps.IdPCGGasto = !IdPCGGasto
        mProps.Codigo = !Codigo
        
        mProps.FechaGrabado = !FechaGrabado
        
        .Close
    End With
    'Recupera  el detalle
    RecuperaDetallePC
   
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset
  
    sql = "DELETE FROM pcggastoDetallepc WHERE Idpcggasto=" & mProps.IdPCGGasto
    mProps.objEmpresa.Execute sql, True
  
    'Elimina el registro
    sql = "DELETE FROM pcggasto WHERE Idpcggasto=" & mProps.IdPCGGasto
    mProps.objEmpresa.Execute sql, True
    
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PCGGASTO
End Sub

Public Function Clone() As PCGGasto
    Dim nuevo As PCGGasto
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New PCGGasto
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdPCGGasto = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Get IdPCGGasto() As Long
    IdPCGGasto = mProps.IdPCGGasto
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

Friend Property Let Modificado(ByVal value As Boolean)
    mbooModificado = value
End Property


Private Sub Class_Initialize()
    'Valores predeterminados
    mbooEsNuevo = True
    mbooModificado = False
    'mProps.BandValida = True
    'mProps.BandCaducidad = False
    'mProps.BandxEmp = False
    'mProps.BandxItem = False
    'mProps.BandAcumula = False

End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub


'
Public Property Let Codigo(ByVal value As String)
    mProps.Codigo = value
    mbooModificado = True
End Property

Public Property Get Codigo() As String
    Codigo = mProps.Codigo
End Property
'
Private Sub BorrarDetallePC()
    'Borra  toda  la informacion de la coleccion
     If mcolPCGGastoDetallePC Is Nothing Then Set mcolPCGGastoDetallePC = New Collection
    With mcolPCGGastoDetallePC
        Do While .Count > 0
            .Remove .Count
        Loop
    End With
    mbooModificado = True
End Sub

Private Sub RecuperaDetallePC()
    Dim sql As String, rs As Recordset, i As Long, mod_antes As Boolean
    Dim nuevo As PCGGastoDetallePC
    mod_antes = mbooModificado      'Guarda el estado de Modificado
    If mcolPCGGastoDetallePC Is Nothing Then Set mcolPCGGastoDetallePC = New Collection
    
    BorrarDetallePC
    
    
    sql = "SELECT * FROM PCGGastoDetallePC WHERE Idpcggasto=" & mProps.IdPCGGasto & " ORDER BY Orden"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        Do Until .EOF
            Set nuevo = New PCGGastoDetallePC
            mcolPCGGastoDetallePC.Add nuevo
            i = mcolPCGGastoDetallePC.Count
            Set PCGGastoDetallePC(i).PCGGasto = Me
            PCGGastoDetallePC(i).Recuperar rs
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    mbooModificado = mod_antes      'Para que no cambie el estado de Modificado
End Sub

Public Property Get PCGGastoDetallePC(ix As Long) As PCGGastoDetallePC
    If mcolPCGGastoDetallePC Is Nothing Then RecuperaDetallePC   'posilbe recursividad
    Set PCGGastoDetallePC = mcolPCGGastoDetallePC.item(ix)
End Property

Public Property Get NumPCGastoDetallePC() As Long
    If mcolPCGGastoDetallePC Is Nothing Then
        NumPCGastoDetallePC = 0
    Else
        NumPCGastoDetallePC = mcolPCGGastoDetallePC.Count
    End If
End Property

Public Function AddDetallePCGGastoPC() As Long
    Dim nuevo As PCGGastoDetallePC
    
    If mcolPCGGastoDetallePC Is Nothing Then RecuperaDetallePC
    Set nuevo = New PCGGastoDetallePC
    Set nuevo.PCGGasto = Me     'Me  olvide  del padre
    'Proporciona valores predeterminados
    nuevo.IdPCGGasto = Me.IdPCGGasto
    mcolPCGGastoDetallePC.Add nuevo
    AddDetallePCGGastoPC = mcolPCGGastoDetallePC.Count   'Devuelve el indice de nuevo elemento
    Set nuevo = Nothing
    mbooModificado = True
End Function

Public Sub RemoveDescuentoDetallePC(ByVal ix As Long, Optional obj As IVDescuentoDetallePC)
    If mcolPCGGastoDetallePC Is Nothing Then RecuperaDetallePC
    
    'Si está especificado un objeto
    If Not (obj Is Nothing) Then
        'Busca el indice del objeto
        For ix = 1 To mcolPCGGastoDetallePC.Count
            If mcolPCGGastoDetallePC.item(ix) Is obj Then
                mcolPCGGastoDetallePC.Remove ix
                Exit For
            End If
        Next ix
        Debug.Assert True     'No se puede dar
    Else
        mcolPCGGastoDetallePC.Remove ix
    End If
    
    mbooModificado = True
End Sub

Public Function RecuperaPCGGastoDetallePC(ix As Long) As PCGGastoDetallePC
    Set RecuperaPCGGastoDetallePC = mcolPCGGastoDetallePC.item(ix)
End Function

Public Function RecuperaEmpxPcGrupo(ByVal CodsGrupo As String) As Recordset
Dim sql As String
Dim rs As Recordset
Dim PcGrupo As Integer
    PcGrupo = Me.Empresa.GNOpcion.ObtenerValor("PCGrupoGasto")
    sql = "select pcg.Idgrupo" & PcGrupo & " ,pc.codProvCli,pc.nombre,'' as gasto "
    sql = sql & "from PCProvCli pc "
    sql = sql & " inner join pcGrupo" & PcGrupo & " pcg ON pc.idgrupo" & PcGrupo & " = pcg.idgrupo" & PcGrupo
    sql = sql & " WHERE bandEmpleado = 1 AND pcg.codgrupo" & PcGrupo & " IN (" & CodsGrupo & ")"
    sql = sql & " order by pcg.codgrupo" & PcGrupo
    Set rs = Me.Empresa.OpenRecordset(sql)
    If Not rs Is Nothing Then
        Set RecuperaEmpxPcGrupo = rs
    End If
    Exit Function
End Function

Private Sub GrabaPCGastoDetallePC()
    Dim rs1 As Recordset, numReg As Long, i As Long
    Dim obj As PCGGastoDetallePC, sql As String
    'Elimina  todos  los registros
    sql = "Delete pcgg " & _
          "From pcggasto pcgg " & _
          "INNER JOIN pcggastoDetallePC pcggD ON pcgg.Idpcggasto = pcggD.Idpcggasto " & _
          "Where pcggD.Idpcggasto = " & mProps.IdPCGGasto
    mProps.objEmpresa.EjecutarSQL sql, numReg
    'guarda  todos  los registros
    If Not mcolPCGGastoDetallePC Is Nothing Then
        sql = "SELECT * FROM pcggastoDetallePC WHERE 1=0"
        Set rs1 = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        For i = 1 To mcolPCGGastoDetallePC.Count
            rs1.AddNew
            Set obj = mcolPCGGastoDetallePC.item(i)
            rs1!IdPCGGasto = mProps.IdPCGGasto
            rs1!idPcGrupo = obj.idPcGrupo
            rs1!IdProvCli = obj.IdProvCli
            rs1!idGasto = obj.idGasto
            rs1!Orden = obj.Orden
            rs1.Update
        Next i
    End If
    Set rs1 = Nothing
    Set obj = Nothing
End Sub
