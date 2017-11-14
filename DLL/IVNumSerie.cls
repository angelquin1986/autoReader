VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVNumSerie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdSerie As Long
    Campo1 As String
    campo2 As String
    Campo3 As String
    Campo4 As String
    Campo5 As String
    Bandera As Boolean
    FechaCreacion As Date
    fechagrabado As Date
    transid As Long
    IdInventario As Long
    IdIvkardexI As Long
    IdIvkardexE As Long
    IdBodega As Long
    Orden As Integer
    Cantidad As Currency
    objEmpresa As Empresa
    objInventario As IVInventario
    objIVKardex As IVKardex
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Friend Sub Recuperar(key As Variant)
    Dim sql As String, i As Integer
    Dim rs As Recordset

    If VarType(key) = vbString Then
        sql = "SELECT * FROM IVSerie WHERE idSerie='" & key & "'"
    Else
        sql = "SELECT * FROM IVSerie WHERE idSerie=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If rs.EOF Then Exit Sub
    'Recupera los datos de Usuario
    With rs
        mProps.IdSerie = !IdSerie
        If Not IsNull(!Campo1) Then mProps.Campo1 = !Campo1
        If Not IsNull(!campo2) Then mProps.campo2 = !campo2
        If Not IsNull(!Campo3) Then mProps.Campo3 = !Campo3
        If Not IsNull(!Campo4) Then mProps.Campo4 = !Campo4
        If Not IsNull(!Campo5) Then mProps.Campo5 = !Campo5
        If Not IsNull(!FechaCreacion) Then mProps.FechaCreacion = !FechaCreacion
        If Not IsNull(!IdInventario) Then mProps.IdInventario = !IdInventario

        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
   Dim sql As String, rs As Recordset
    
    'Verifica si ya existe un Kardex usando el mismo IdInventario
    sql = "SELECT Id FROM ivkardexserie WHERE IdSerie=" & mProps.IdSerie
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "IVSerie_Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close
    
    
    
    'Elimina el registro
    sql = "DELETE FROM IVSerie WHERE IdSerie=" & mProps.IdSerie
    mProps.objEmpresa.Execute sql, True

    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_ITEMSerie", "Serie Eliminada: " & _
                                                         mProps.Campo1, "DT"
                                                         
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVSERIE

End Sub

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdSerie = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(value As Empresa)
    Set mProps.objEmpresa = value
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Friend Property Set IVInventario(value As IVInventario)
    Set mProps.objInventario = value
End Property

Public Property Get IVInventario() As IVInventario
    Set IVInventario = mProps.objInventario
End Property

Friend Property Set objIVKardex(value As IVKardex)  'AUC estos estan aqui por la desintegridad ya no sirven en este modulo
    Set mProps.objIVKardex = value
End Property

Public Property Get objIVKardex() As IVKardex
    Set objIVKardex = mProps.objIVKardex
End Property
Public Property Get IdSerie() As Long
    IdSerie = mProps.IdSerie
End Property

Public Property Get IdInventario() As Long
    IdInventario = Trim$(mProps.IdInventario)
End Property

Public Property Let IdInventario(ByVal value As Long)
    mProps.IdInventario = value
    mbooModificado = True
End Property

Public Property Let Campo1(ByVal value As String)
    If Len(value) > 40 Then
        Err.Raise ERR_INVALIDO, "IVSerie_Campo1", MSGERR_INVALIDO
    End If
    If InStr(value, ";") > 0 Then
        Err.Raise ERR_INVALIDO, "IVSerie_Campo1", _
            "El caracter ';' no es válido para la Campo1"
    End If
    If InStr(value, "|") > 0 Then
        Err.Raise ERR_INVALIDO, "IVSerie_Campo1", _
            "El caracter '|' no es válido para la Campo1"
    End If
    mProps.Campo1 = value
    mbooModificado = True
End Property

Public Property Get Campo1() As String
    Campo1 = Trim$(mProps.Campo1)
End Property

Public Property Let campo2(ByVal value As String)
    If Len(value) > 40 Then
        Err.Raise ERR_INVALIDO, "IVSerie_Campo2", MSGERR_INVALIDO
    End If
    If InStr(value, ";") > 0 Then
        Err.Raise ERR_INVALIDO, "IVSerie_Campo2", _
            "El caracter ';' no es válido para la Campo2"
    End If
    If InStr(value, "|") > 0 Then
        Err.Raise ERR_INVALIDO, "IVSerie_Campo2", _
            "El caracter '|' no es válido para la Campo2"
    End If
    mProps.campo2 = value
    mbooModificado = True
End Property

Public Property Get campo2() As String
    campo2 = Trim$(mProps.campo2)
End Property

Public Property Let Campo3(ByVal value As String)
    If Len(value) > 40 Then
        Err.Raise ERR_INVALIDO, "IVSerie_Campo3", MSGERR_INVALIDO
    End If
    If InStr(value, ";") > 0 Then
        Err.Raise ERR_INVALIDO, "IVSerie_Campo3", _
            "El caracter ';' no es válido para la Campo3"
    End If
    If InStr(value, "|") > 0 Then
        Err.Raise ERR_INVALIDO, "IVSerie_Campo3", _
            "El caracter '|' no es válido para la Campo3"
    End If
    mProps.Campo3 = value
    mbooModificado = True
End Property

Public Property Get Campo3() As String
    Campo3 = Trim$(mProps.Campo3)
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

Public Property Let Campo4(ByVal value As String)
    If Len(value) > 40 Then
        Err.Raise ERR_INVALIDO, "IVSerie_Campo4", MSGERR_INVALIDO
    End If
    If InStr(value, ";") > 0 Then
        Err.Raise ERR_INVALIDO, "IVSerie_Campo4", _
            "El caracter ';' no es válido para la Campo4"
    End If
    If InStr(value, "|") > 0 Then
        Err.Raise ERR_INVALIDO, "IVSerie_Campo4", _
            "El caracter '|' no es válido para la Campo4"
    End If
    mProps.Campo4 = Trim(value)
    mbooModificado = True
End Property

Public Property Get Campo4() As String
    Campo4 = mProps.Campo4
End Property

Public Property Let Campo5(ByVal value As String)
    If Len(value) > 40 Then
        Err.Raise ERR_INVALIDO, "IVSerie_Campo5", MSGERR_INVALIDO
    End If
    If InStr(value, ";") > 0 Then
        Err.Raise ERR_INVALIDO, "IVSerie_Campo5", _
            "El caracter ';' no es válido para la Campo5"
    End If
    If InStr(value, "|") > 0 Then
        Err.Raise ERR_INVALIDO, "IVSerie_Campo3", _
            "El caracter '|' no es válido para la Campo5"
    End If
    mProps.Campo5 = value
    mbooModificado = True
End Property

Public Property Get Campo5() As String
    Campo5 = Trim$(mProps.Campo5)
End Property
Public Property Let FechaCreacion(ByVal value As Date)
    mProps.FechaCreacion = value
    mbooModificado = True
End Property

Public Property Get FechaCreacion() As Date
    FechaCreacion = mProps.FechaCreacion
End Property

Private Sub Class_Initialize()
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Property Get IdIvkardexE() As Long
    IdIvkardexE = mProps.IdIvkardexE
End Property
Public Property Let IdIvkardexE(ByVal value As Long)
    mProps.IdIvkardexE = value
    mbooModificado = True
End Property

Public Property Get IdIvkardexI() As Long
    IdIvkardexI = mProps.IdIvkardexI
End Property
Public Property Let IdIvkardexI(ByVal value As Long)
    mProps.IdIvkardexI = value
    mbooModificado = True
End Property
Public Property Get Bandera() As Boolean
    Bandera = mProps.Bandera
End Property
Public Property Let Bandera(ByVal value As Boolean)
    mProps.Bandera = value
    mbooModificado = True
End Property


Public Property Get transid() As Long
    transid = mProps.transid
End Property
Public Property Let transid(ByVal value As Long)
    mProps.transid = value
    mbooModificado = True
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property
Public Property Let Orden(ByVal value As Integer)
    mProps.Orden = value
    mbooModificado = True
End Property

Public Sub GrabarIVSerie(ByVal IdSerie As Long)
End Sub
 Public Sub GrabarIVSerieNew(ByRef IdSerie As Long)
    Dim sql As String, i As Integer
    Dim rs As Recordset, rsDetalle As Recordset

    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    On Error GoTo ErrTrap
        
    If IdSerie <> 0 Then
        sql = "SELECT * FROM IVSerie WHERE IDSERIE = " & IdSerie
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
    Else
        sql = "SELECT * FROM IVSerie WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    End If
    'Si es modificada
 '   Else
        'Recupera el registro correspondiente para modificar
  '      sql = "SELECT * FROM IVSerie WHERE Idserie=" & IdSerie
   '     Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
   ' End If
    
    'Graba el registro
    With rs
'        !IdIvkardexI = mProps.IdIvkardexI
 '       !IdIvkardexE = mProps.IdIvkardexE
        !Campo1 = mProps.Campo1
        !campo2 = mProps.campo2
        !Campo3 = mProps.Campo3
        !Campo4 = mProps.Campo4
        !Campo5 = mProps.Campo5
'        !transid = mProps.transid
        '!IdBodega = mProps.IdBodega
        !IdInventario = mProps.IdInventario
 '       !Orden = mProps.Orden
        !FechaCreacion = mProps.FechaCreacion
        '!Cantidad = mProps.Cantidad
        !fechagrabado = Date
        .Update
        rs.Move 0           'Para actualizar
        IdSerie = !IdSerie
        .Close
    End With
    Set rs = Nothing
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_ITEMSERIE", "ItemSerie Nuevo: " & _
                                                         mProps.Campo1, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_ITEMSERIE", "ItemSerie Modificado: " & _
                                                         mProps.Campo1, "DT"
    End If
    mbooEsNuevo = False
    mbooModificado = False
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVSERIE
    Exit Sub
ErrTrap:
 MsgBox Err.Description
End Sub


Public Property Get IdBodega() As Long
    IdBodega = mProps.IdBodega
End Property
Public Property Let IdBodega(ByVal value As Long)
    mProps.IdBodega = value
    mbooModificado = True
End Property

Public Property Get Cantidad() As Currency
    Cantidad = mProps.Cantidad
End Property
Public Property Let Cantidad(ByVal value As Currency)
    mProps.Cantidad = value
    mbooModificado = True
End Property

Public Sub VerificaDatos()
    Dim cant As Currency
    Dim item As IVInventario
    Dim pc As PCProvCli, nunPCGrupo As Integer, cadPrecios As String, idPcGrupo As Long
    Dim nprecio As Integer
    Dim cad As String
    
    If mProps.Cantidad = 0 Then
        Err.Raise ERR_INVALIDO, "IVKNumSerie_VerificaDatos", _
            "La cantidad no puede ser 0. (" & mProps.Campo1 & ")"
    End If
    If mProps.IdBodega = 0 Then
        Err.Raise ERR_INVALIDO, "IVKNumSerie_VerificaDatos", _
            "Debe seleccionar una bodega. (" & mProps.IdBodega & ")"
    End If
    If mProps.IdInventario = 0 Then
        Err.Raise ERR_INVALIDO, "IVKNumSerie_VerificaDatos", _
            "Debe seleccionar un item."
    End If
    cad = mProps.Campo1
    If VerificaExiste(cad, 1) Then
        Err.Raise ERR_INVALIDO, "IVKNumSerie_VerificaDatos", _
           cad & " El valor ya existe..."
    End If
    cad = mProps.campo2
    If VerificaExiste(cad, 2) Then
        Err.Raise ERR_INVALIDO, "IVKNumSerie_VerificaDatos", _
           cad & " El valor ya existe..."
    End If
    cad = mProps.Campo3
    If VerificaExiste(cad, 3) Then
        Err.Raise ERR_INVALIDO, "IVKNumSerie_VerificaDatos", _
           cad & " El valor ya existe..."
    End If
    cad = mProps.Campo4
    If VerificaExiste(cad, 4) Then
        Err.Raise ERR_INVALIDO, "IVKNumSerie_VerificaDatos", _
             cad & " El valor ya existe..."
    End If
End Sub

Private Function VerificaExiste(ByRef campo As String, NumCampo As Integer) As Boolean
Dim sql As String
Dim rs As Recordset
Select Case NumCampo
    Case 1
        sql = "Select * from IVSerie where campo1 = '" & campo & "'"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.RecordCount > 0 Then
            campo = rs!Campo1
            VerificaExiste = True
        End If
    
    Case 2
        sql = "Select * from IVSerie where campo2 = '" & campo & "'"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.RecordCount > 0 Then
            campo = rs!campo2
            VerificaExiste = True
        End If
    Case 3
        sql = "Select * from IVSerie where campo3 = '" & campo & "'"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.RecordCount > 0 Then
            campo = rs!Campo3
            VerificaExiste = True
        End If
    Case 4
        sql = "Select * from IVSerie where campo4 = '" & campo & "'"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.RecordCount > 0 Then
            campo = rs!Campo4
            VerificaExiste = True
        End If
End Select
    Exit Function
End Function

Friend Sub RecuperarXItem(key As Variant)
    Dim sql As String, i As Integer
    Dim rs As Recordset

    If VarType(key) = vbString Then
        sql = "SELECT ivs.* FROM IVSerie ivs Inner Join Ivinventario iv on iv.idinventario = ivs.idinventario WHERE iv.idInventario='" & key & "'"
    Else
        sql = "SELECT ivs.* FROM IVSerie ivs Inner Join Ivinventario iv on iv.idinventario = ivs.idinventario WHERE iv.idInventario=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If rs.EOF Then Exit Sub
    'Recupera los datos de Usuario
    With rs
        mProps.IdSerie = !IdSerie
        If Not IsNull(!Campo1) Then mProps.Campo1 = !Campo1
        If Not IsNull(!campo2) Then mProps.campo2 = !campo2
        If Not IsNull(!Campo3) Then mProps.Campo3 = !Campo3
        If Not IsNull(!Campo4) Then mProps.Campo4 = !Campo4
        If Not IsNull(!Campo5) Then mProps.Campo5 = !Campo5
        If Not IsNull(!FechaCreacion) Then mProps.FechaCreacion = !FechaCreacion
        If Not IsNull(!IdInventario) Then mProps.IdInventario = !IdInventario

        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Friend Sub RecuperarxidIvKardex(key As Long)
    Dim sql As String, i As Integer
    Dim rs As Recordset

    
    sql = "SELECT ivs.* FROM IVSerie ivs"
    sql = sql & " Inner Join IvkardexSerie ivks on ivks.idserie = ivs.idserie"
    sql = sql & " WHERE ivks.idivkardex = " & key
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If rs.EOF Then Exit Sub
    'Recupera los datos de Usuario
    With rs
        Do While Not rs.EOF
            mProps.IdSerie = !IdSerie                                                 'este campo como me barro queda el ultimo
            If Not IsNull(!Campo1) Then mProps.Campo1 = mProps.Campo1 & " " & !Campo1 'anido los campos para
            If Not IsNull(!campo2) Then mProps.campo2 = mProps.campo2 & " " & !campo2 'cuando sean mas de uno
            If Not IsNull(!Campo3) Then mProps.Campo3 = mProps.Campo3 & " " & !Campo3 'en ivkardex ejm vendo el mismo
            If Not IsNull(!Campo4) Then mProps.Campo4 = mProps.Campo4 & " " & !Campo4 ' item en cant 2
            If Not IsNull(!Campo5) Then mProps.Campo5 = mProps.Campo5 & " " & !Campo5 '
            If Not IsNull(!FechaCreacion) Then mProps.FechaCreacion = !FechaCreacion
            If Not IsNull(!IdInventario) Then mProps.IdInventario = !IdInventario
            rs.MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

