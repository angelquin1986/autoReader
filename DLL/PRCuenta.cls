VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PRCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdCuenta As Long
    CodCuenta As String
    NombreCuenta As String
    Nivel As Byte
    IdCuentaSuma As Long
    CodCuentaSuma As String
    TipoCuenta As Byte
    BandDeudor As Boolean
    BandTotal As Boolean
    BandValida As Boolean
    IDLocal  As Long
    CodLocal As String
    
    IDCuentaCont2  As Long
    CodCuentaCont2 As String
    
    IDCuentaCont3  As Long
    CodCuentaCont3 As String
    
    IDCuentaCont4  As Long
    CodCuentaCont4 As String
    
    IDProvCli  As Long
    CodProvCli As String
    
    
    BandGastoDistribuido As Boolean             'jeaa 17/01/2008 para gastos distribuidos
    ValPresupuesto As Currency                      'jeaa 10/01/2009 para presupuestos
    FechaGrabado As Date
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean


Public Sub Grabar()
    Dim sql As String, rs As Recordset
    
    If TipoCuenta = 0 Then
        Err.Raise ERR_INVALIDO, "PRCuenta_Grabar", _
            "Tipo de cuenta no está especificado."
    End If
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código                '*** MAKOTO 17/oct/00 Modificado
    sql = "SELECT CodCuenta FROM PRCuenta " & _
          "WHERE (CodCuenta='" & mProps.CodCuenta & "') AND " & _
                "(IdCuenta<>" & mProps.IdCuenta & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "PRCuenta", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM PRCuenta WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM PRCuenta WHERE IdCuenta=" & mProps.IdCuenta
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
       
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "PRCuenta_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro
    With rs
        !CodCuenta = mProps.CodCuenta
        !NombreCuenta = mProps.NombreCuenta
        !Nivel = mProps.Nivel
        !IdCuentaSuma = mProps.IdCuentaSuma
        !TipoCuenta = mProps.TipoCuenta
        !BandDeudor = mProps.BandDeudor
        !BandTotal = mProps.BandTotal
        !BandValida = mProps.BandValida
        !IDLocal = mProps.IDLocal
'        !BandGastoDistribuido = mProps.BandGastoDistribuido  'jeaa 17/01/2008
        !ValPresupuesto = mProps.ValPresupuesto  'jeaa 10/01/2009
        
        !IDCuentaCont2 = mProps.IDCuentaCont2
        !IDCuentaCont3 = mProps.IDCuentaCont3
        !IDCuentaCont4 = mProps.IDCuentaCont4
        !IDProvCli = mProps.IDProvCli
        
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified       'Para actualizar
#End If
        .Move 0             'Para actualizar
        mProps.FechaGrabado = !FechaGrabado
        mProps.IdCuenta = !IdCuenta
        
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 18/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_PRCTA", "Cuenta Presupuesto Nueva: " & _
                                                         mProps.CodCuenta, "PR"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_PRCTA", "Cuenta Presupuesto Modificada: " & _
                                                         mProps.CodCuenta, "PR"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PRCUENTA
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT A.*,B.CodCuenta AS CodCuentaSuma, ct.codcuenta as codlocal, "
    sql = sql & " ct2.codcuenta as CodCuentaCont2,  "
    sql = sql & " ct3.codcuenta as CodCuentaCont3,  "
    sql = sql & " ct4.codcuenta as CodCuentaCont4, pc.codprovcli  "
    sql = sql & " FROM PRCuenta A "
    sql = sql & " left join ctcuenta ct on a.idlocal=ct.idcuenta"
    sql = sql & " left join ctcuenta ct2 on a.IDCuentaCont2 = ct2.idcuenta"
    sql = sql & " left join ctcuenta ct3 on a.IDCuentaCont3 = ct3.idcuenta"
    sql = sql & " left join ctcuenta ct4 on a.IDCuentaCont4  = ct4.idcuenta"
    sql = sql & " LEFT JOIN PRCuenta B ON A.IdCuentaSuma=B.IdCuenta "
    sql = sql & " LEFT JOIN Pcprovcli pc ON A.Idprovcli=pc.Idprovcli "
        
    sql = sql & ""
        
    If VarType(key) = vbString Then
        sql = sql & "WHERE A.CodCuenta='" & key & "'"
    Else
        sql = sql & "WHERE A.IdCuenta=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IdCuenta = !IdCuenta
        mProps.CodCuenta = !CodCuenta
        If Not IsNull(!NombreCuenta) Then mProps.NombreCuenta = !NombreCuenta
        If Not IsNull(!Nivel) Then mProps.Nivel = !Nivel
        If Not IsNull(!IdCuentaSuma) Then mProps.IdCuentaSuma = !IdCuentaSuma
        If Not IsNull(!CodCuentaSuma) Then mProps.CodCuentaSuma = !CodCuentaSuma
        If Not IsNull(!TipoCuenta) Then mProps.TipoCuenta = !TipoCuenta
        If Not IsNull(!BandDeudor) Then mProps.BandDeudor = !BandDeudor
        If Not IsNull(!BandTotal) Then mProps.BandTotal = !BandTotal
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!IDLocal) Then mProps.IDLocal = !IDLocal
        If Not IsNull(!CodLocal) Then mProps.CodLocal = !CodLocal
'        If Not IsNull(!BandGastoDistribuido) Then mProps.BandGastoDistribuido = !BandGastoDistribuido
        If Not IsNull(!ValPresupuesto) Then mProps.ValPresupuesto = !ValPresupuesto
        
        
        If Not IsNull(!IDCuentaCont2) Then mProps.IDCuentaCont2 = !IDCuentaCont2
        If Not IsNull(!CodCuentaCont2) Then mProps.CodCuentaCont2 = !CodCuentaCont2
        If Not IsNull(!IDCuentaCont3) Then mProps.IDCuentaCont3 = !IDCuentaCont3
        If Not IsNull(!CodCuentaCont3) Then mProps.CodCuentaCont3 = !CodCuentaCont3
        If Not IsNull(!IDCuentaCont4) Then mProps.IDCuentaCont4 = !IDCuentaCont4
        If Not IsNull(!CodCuentaCont4) Then mProps.CodCuentaCont4 = !CodCuentaCont4
        
        If Not IsNull(!IDProvCli) Then mProps.IDProvCli = !IDProvCli
        If Not IsNull(!CodProvCli) Then mProps.CodProvCli = !CodProvCli
        
        
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    EliminaCuenta mProps.IdCuenta

    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_PRCUENTA
End Sub

Private Sub EliminaCuenta(id As Long)
    Dim sql As String, rs As Recordset
    
    'Verifica si ya tiene movimientos
    sql = "SELECT IdCuenta FROM PRLibroDetalle WHERE IdCuenta=" & id
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        'Si ya tiene movimientos
        Err.Raise ERR_NOELIMINABLE, "PRCuenta.Eliminar", _
                  "La cuenta '" & mProps.CodCuenta & _
                  "' ya tiene movimientos, " & _
                  "por lo que no puede ser eliminada."
    End If
    rs.Close
    

    'Obtiene los hijos
    sql = "SELECT IdCuenta FROM PRCuenta WHERE IdCuentaSuma=" & id
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    With rs
        'Elimina los hijos primero
        Do Until .EOF
            EliminaCuenta rs!IdCuenta
            .MoveNext
        Loop
        .Close
    End With

    'Elimina si misma
    sql = "DELETE FROM PRCuenta WHERE IdCuenta=" & id
    mProps.objEmpresa.Execute sql, True
    
    Set rs = Nothing
    
    'Graba un registro para auditoria   *** ANGEL 18/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_PRCTA", "Cuenta Presupuesto Eliminada: " & _
                                                         mProps.CodCuenta, "PR"
End Sub

Public Function Clone() As PRCuenta
    Dim nuevo As PRCuenta
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New PRCuenta
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdCuenta = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Property Let CodCuenta(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PRCuenta_CodCuenta", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodCuenta = value
    mbooModificado = True
End Property

Public Property Get CodCuenta() As String
    CodCuenta = Trim$(mProps.CodCuenta)
End Property

Public Property Let NombreCuenta(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "PRCuenta_NombreCuenta", MSGERR_INVALIDO
    End If
    
    mProps.NombreCuenta = value
    mbooModificado = True
End Property

Public Property Get NombreCuenta() As String
    NombreCuenta = Trim$(mProps.NombreCuenta)
End Property

Public Property Let Nivel(ByVal value As Byte)
    mProps.Nivel = value
    mbooModificado = True
End Property

Public Property Get Nivel() As Byte
    Nivel = mProps.Nivel
End Property

Public Property Let IdCuentaSuma(ByVal value As Long)
    mProps.IdCuentaSuma = value
    mbooModificado = True
End Property

Public Property Get IdCuentaSuma() As Long
    IdCuentaSuma = mProps.IdCuentaSuma
End Property

Public Property Let CodCuentaSuma(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCuentaSuma Then
        'Actualiza IdCuentaSuma también
        sql = "SELECT IdCuenta FROM PRCuenta WHERE CodCuenta='" & value & "'"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "PRCuenta_CodCuentaSuma", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            IdCuentaSuma = rs!IdCuenta
        End If
        rs.Close
    End If
    
    Set rs = Nothing
    mProps.CodCuentaSuma = value
    mbooModificado = True
End Property

Public Property Get CodCuentaSuma() As String
    CodCuentaSuma = Trim$(mProps.CodCuentaSuma)
End Property



Public Property Let TipoCuenta(ByVal value As Byte)
    If value < 1 Or value > 5 Then
        Err.Raise ERR_INVALIDO, "PRCuenta_TipoCuenta", MSGERR_INVALIDO
    End If
    
    'Cambia el valor de BandDeudor
    Select Case value
    Case 1, 5
        mProps.BandDeudor = True
    Case 2, 3, 4
        mProps.BandDeudor = False
    End Select
    
    mProps.TipoCuenta = value
    mbooModificado = True
End Property

Public Property Get TipoCuenta() As Byte
    TipoCuenta = mProps.TipoCuenta
End Property

Public Property Get BandDeudor() As Boolean
    BandDeudor = mProps.BandDeudor
End Property

Public Property Let BandTotal(ByVal value As Boolean)
    Dim sql As String, rs As Recordset
    
    'Si es modificación
    If Not EsNuevo Then
        'Cuando intenta cambiar al False, verifica si tiene hijos
        If Not value Then
            'Verifica si tiene hijos
            sql = "SELECT CodCuenta FROM PRCuenta WHERE IdCuentaSuma=" & mProps.IdCuenta
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            'Si ya tiene hijos
            If Not rs.EOF Then
                Err.Raise ERR_INVALIDO, "PRCuenta_BandTotal", _
                          "La cuenta '" & mProps.CodCuenta & "' tiene otras cuentas que pertenece a la misma, " & _
                          "por lo que no se puede ser como cuenta de movimiento."
            End If
            
        'Cuando intenta cambiar a True, verifica si ya tiene movimientos
        Else
            'Verifica si tiene movimientos
            sql = "SELECT IdCuenta FROM CTLibroDetalle WHERE IdCuenta=" & mProps.IdCuenta
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            'Si ya tiene movimientos
            If Not rs.EOF Then
                Err.Raise ERR_INVALIDO, "PRCuenta.BandTotal", _
                          "La cuenta '" & mProps.CodCuenta & "' ya tiene movimientos, " & _
                          "por lo que no puede ser como cuenta de total."
            End If
        End If
    End If
    
    Set rs = Nothing
    mProps.BandTotal = value
    mbooModificado = True
End Property

Public Property Get BandTotal() As Boolean
    BandTotal = mProps.BandTotal
End Property



Public Property Let BandValida(ByVal value As Boolean)
    mProps.BandValida = value
    mbooModificado = True
End Property

Public Property Get BandValida() As Boolean
    BandValida = mProps.BandValida
End Property

Public Property Let IDLocal(ByVal value As Long)
'    If Len(value) > 2 Then
'        Err.Raise ERR_INVALIDO, "PRCuenta_Local", MSGERR_INVALIDO
'    End If
    mProps.IDLocal = value
    mbooModificado = True
End Property

Public Property Get IDLocal() As Long
    IDLocal = mProps.IDLocal
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
        .TipoCuenta = 0             'Tipo indefinido
        .BandDeudor = True
        .BandTotal = False
        .BandValida = True
        .BandGastoDistribuido = False
    End With

    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Property Let CodLocal(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodLocal Then
        If Len(value) > 0 Then
            'Actualiza IdProveedor también
            sql = "SELECT IdCuenta FROM CtCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "CtCuenta_CodCuenta", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IDLocal = rs!IdCuenta
            End If
            rs.Close
        Else
            IDLocal = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodLocal = value
    mbooModificado = True
End Property

Public Property Get CodLocal() As String
    CodLocal = Trim$(mProps.CodLocal)
End Property

Public Property Let BandGastoDistribuido(ByVal value As Boolean)
    mProps.BandGastoDistribuido = value
    mbooModificado = True
End Property

Public Property Get BandGastoDistribuido() As Boolean
    BandGastoDistribuido = mProps.BandGastoDistribuido
End Property

Public Property Let ValPresupuesto(ByVal value As Currency)
    mProps.ValPresupuesto = value
    mbooModificado = True
End Property

Public Property Get ValPresupuesto() As Currency
    ValPresupuesto = mProps.ValPresupuesto
End Property

'------------2

Public Property Let IDCuentaCont2(ByVal value As Long)
'    If Len(value) > 2 Then
'        Err.Raise ERR_INVALIDO, "PRCuenta_CuentaCont", MSGERR_INVALIDO
'    End If
    mProps.IDCuentaCont2 = value
    mbooModificado = True
End Property

Public Property Get IDCuentaCont2() As Long
    IDCuentaCont2 = mProps.IDCuentaCont2
End Property

Public Property Let CodCuentaCont2(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCuentaCont2 Then
        If Len(value) > 0 Then
            'Actualiza IdProveedor también
            sql = "SELECT IdCuenta FROM CtCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "CtCuenta_CodCuenta2", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IDCuentaCont2 = rs!IdCuenta
            End If
            rs.Close
        Else
            IDCuentaCont2 = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodCuentaCont2 = value
    mbooModificado = True
End Property

Public Property Get CodCuentaCont2() As String
    CodCuentaCont2 = Trim$(mProps.CodCuentaCont2)
End Property


'------------3

Public Property Let IDCuentaCont3(ByVal value As Long)
'    If Len(value) > 2 Then
'        Err.Raise ERR_INVALIDO, "PRCuenta_CuentaCont", MSGERR_INVALIDO
'    End If
    mProps.IDCuentaCont3 = value
    mbooModificado = True
End Property

Public Property Get IDCuentaCont3() As Long
    IDCuentaCont3 = mProps.IDCuentaCont3
End Property

Public Property Let CodCuentaCont3(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCuentaCont3 Then
        If Len(value) > 0 Then
            'Actualiza IdProveedor también
            sql = "SELECT IdCuenta FROM CtCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "CtCuenta_CodCuenta", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IDCuentaCont3 = rs!IdCuenta
            End If
            rs.Close
        Else
            IDCuentaCont3 = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodCuentaCont3 = value
    mbooModificado = True
End Property

Public Property Get CodCuentaCont3() As String
    CodCuentaCont3 = Trim$(mProps.CodCuentaCont3)
End Property


'------------4

Public Property Let IDCuentaCont4(ByVal value As Long)
'    If Len(value) > 2 Then
'        Err.Raise ERR_INVALIDO, "PRCuenta_CuentaCont", MSGERR_INVALIDO
'    End If
    mProps.IDCuentaCont4 = value
    mbooModificado = True
End Property

Public Property Get IDCuentaCont4() As Long
    IDCuentaCont4 = mProps.IDCuentaCont4
End Property

Public Property Let CodCuentaCont4(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCuentaCont4 Then
        If Len(value) > 0 Then
            'Actualiza IdProveedor también
            sql = "SELECT IdCuenta FROM CtCuenta WHERE CodCuenta='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "CtCuenta_CodCuenta", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IDCuentaCont4 = rs!IdCuenta
            End If
            rs.Close
        Else
            IDCuentaCont4 = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodCuentaCont4 = value
    mbooModificado = True
End Property

Public Property Get CodCuentaCont4() As String
    CodCuentaCont4 = Trim$(mProps.CodCuentaCont4)
End Property

Public Property Let IDProvCli(ByVal value As Long)
    mProps.IDProvCli = value
    mbooModificado = True
End Property

Public Property Get IDProvCli() As Long
    IDProvCli = mProps.IDProvCli
End Property

Public Property Let CodProvCli(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodProvCli Then
        If Len(value) > 0 Then
            'Actualiza IdProveedor también
            sql = "SELECT IdProvCli FROM PCProvCli WHERE CodProvCli='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "CtCuenta_CodCuenta", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IDProvCli = rs!IDProvCli
            End If
            rs.Close
        Else
            IDProvCli = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodProvCli = value
    mbooModificado = True
End Property

Public Property Get CodProvCli() As String
    CodProvCli = Trim$(mProps.CodProvCli)
End Property

