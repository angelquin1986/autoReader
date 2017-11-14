VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Empresa"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
    bandValida As Boolean
    IDLocal  As Long
    CodLocal As String
    BandGastoDistribuido As Boolean             'jeaa 17/01/2008 para gastos distribuidos
    ValPresupuesto As Currency                      'jeaa 10/01/2009 para presupuestos
    FechaGrabado As Date
    IdCuentaSC As Long
    CodCuentaSC As String
    Campo101 As Integer
    IdCuentaFE As Long
    CodCuentaFE As String
    TipoCuentaFE As Byte
    BandDeudorFE As Boolean
    FormulaCuentaFE As String
    TipoFormulaFE As Byte
    CampoF101 As Integer
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean


Public Sub Grabar()
    Dim sql As String, rs As Recordset
    
    If TipoCuenta = 0 Then
        Err.Raise ERR_INVALIDO, "CTCuenta_Grabar", _
            "Tipo de cuenta no está especificado."
    End If
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código                '*** MAKOTO 17/oct/00 Modificado
    sql = "SELECT CodCuenta FROM CTCuenta " & _
          "WHERE (CodCuenta='" & mProps.CodCuenta & "') AND " & _
                "(IdCuenta<>" & mProps.IdCuenta & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "CTCuenta", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM CTCuenta WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM CTCuenta WHERE IdCuenta=" & mProps.IdCuenta
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
       
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "CTCuenta_Grabar", MSGERR_MODIFICADO
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
        !bandValida = mProps.bandValida
        !IDLocal = mProps.IDLocal
        !BandGastoDistribuido = mProps.BandGastoDistribuido  'jeaa 17/01/2008
        !ValPresupuesto = mProps.ValPresupuesto  'jeaa 10/01/2009
        !IdCuentaSC = mProps.IdCuentaSC
        !Campo101 = mProps.Campo101
        !IdCuentaFE = mProps.IdCuentaFE
        !TipoCuentaFE = mProps.TipoCuentaFE
        !BandDeudorFE = mProps.BandDeudorFE
        !FormulaCuentaFE = mProps.FormulaCuentaFE
        !TipoFormulaFE = mProps.TipoFormulaFE
        !CampoF101 = mProps.CampoF101
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
        mProps.objEmpresa.GrabaGNLogAccion "NUE_CUENTA", "Cuenta Contable Nueva: " & _
                                                         mProps.CodCuenta, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_CUENTA", "Cuenta Contable Modificada: " & _
                                                         mProps.CodCuenta, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_CUENTA
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    sql = "SELECT A.*,B.CodCuenta AS CodCuentaSuma, SC.Codcuenta As CodCuentaSC,  " & _
            " FE.Codcuenta As CodCuentaFE FROM CTCuenta A " & _
          " LEFT JOIN CTCuenta B ON A.IdCuentaSuma=B.IdCuenta " & _
          " LEFT JOIN CTCuentaSC SC ON A.IdCuentaSC = SC.IdCuenta " & _
          " LEFT JOIN CTCuentaFE FE ON A.IdCuentaFE = FE.IdCuenta "
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
        If Not IsNull(!bandValida) Then mProps.bandValida = !bandValida
        If Not IsNull(!IDLocal) Then mProps.IDLocal = !IDLocal
        If Not IsNull(!BandGastoDistribuido) Then mProps.BandGastoDistribuido = !BandGastoDistribuido
        If Not IsNull(!ValPresupuesto) Then mProps.ValPresupuesto = !ValPresupuesto
        If Not IsNull(!IdCuentaSC) Then mProps.IdCuentaSC = !IdCuentaSC
        If Not IsNull(!CodCuentaSC) Then mProps.CodCuentaSC = !CodCuentaSC
        If Not IsNull(!Campo101) Then mProps.Campo101 = !Campo101
        If Not IsNull(!IdCuentaFE) Then mProps.IdCuentaFE = !IdCuentaFE
        If Not IsNull(!CodCuentaFE) Then mProps.CodCuentaFE = !CodCuentaFE
        If Not IsNull(!TipoCuentaFE) Then mProps.TipoCuentaFE = !TipoCuentaFE
        If Not IsNull(!BandDeudorFE) Then mProps.BandDeudorFE = !BandDeudorFE
        If Not IsNull(!FormulaCuentaFE) Then mProps.FormulaCuentaFE = !FormulaCuentaFE
        If Not IsNull(!TipoFormulaFE) Then mProps.TipoFormulaFE = !TipoFormulaFE
        If Not IsNull(!CampoF101) Then mProps.CampoF101 = !CampoF101
        
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
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_CUENTA
End Sub

Private Sub EliminaCuenta(Id As Long)
    Dim sql As String, rs As Recordset
    
    'Verifica si ya tiene movimientos
    sql = "SELECT IdCuenta FROM CTLibroDetalle WHERE IdCuenta=" & Id
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        'Si ya tiene movimientos
        Err.Raise ERR_NOELIMINABLE, "CTCuenta.Eliminar", _
                  "La cuenta '" & mProps.CodCuenta & _
                  "' ya tiene movimientos, " & _
                  "por lo que no puede ser eliminada."
    End If
    rs.Close
    
    'Verifica si existe en PCProvCli registros que hace referencia a la cuenta
    sql = "SELECT IdCuentaContable FROM PCProvCli WHERE IdCuentaContable=" & Id
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        'Si ya existen registros que hacen referencia
        Err.Raise ERR_NOELIMINABLE, "CTCuenta.Eliminar", _
                  "La cuenta '" & mProps.CodCuenta & _
                  "' está siendo usada por uno o más proveedores/clientes, " & _
                  "por lo que no puede ser eliminada."
    End If
    rs.Close
    
    'Verifica si existe en TSBanco registros que hace referencia a la cuenta
    sql = "SELECT IdCuentaContable FROM TSBanco WHERE IdCuentaContable=" & Id
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        'Si ya existen registros que hacen referencia
        Err.Raise ERR_NOELIMINABLE, "CTCuenta.Eliminar", _
                  "La cuenta '" & mProps.CodCuenta & _
                  "' está siendo usada por uno o más bancos, " & _
                  "por lo que no puede ser eliminada."
    End If
    rs.Close
    
    'Verifica si existe en IVInventario registros que hacen referencia a la cuenta
    sql = "SELECT IdInventario FROM IVInventario " & _
          "WHERE (IdCuentaActivo=" & Id & ") OR " & _
                "(IdCuentaCosto=" & Id & ") OR " & _
                "(IdCuentaVenta=" & Id & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        'Si ya existen registros que hacen referencia
        Err.Raise ERR_NOELIMINABLE, "CTCuenta.Eliminar", _
                  "La cuenta '" & mProps.CodCuenta & _
                  "' está siendo usada por uno o más items de inventario, " & _
                  "por lo que no puede ser eliminada."
    End If
    rs.Close
    
    'Obtiene los hijos
    sql = "SELECT IdCuenta FROM CTCuenta WHERE IdCuentaSuma=" & Id
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
    sql = "DELETE FROM CTCuenta WHERE IdCuenta=" & Id
    mProps.objEmpresa.Execute sql, True
    
    Set rs = Nothing
    
    'Graba un registro para auditoria   *** ANGEL 18/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_CUENTA", "Cuenta Contable Eliminada: " & _
                                                         mProps.CodCuenta, "DT"
End Sub

Public Function Clone() As CTCuenta
    Dim nuevo As CTCuenta
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New CTCuenta
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
        Err.Raise ERR_INVALIDO, "CTCuenta_CodCuenta", MSGERR_INVALIDO
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
        Err.Raise ERR_INVALIDO, "CTCuenta_NombreCuenta", MSGERR_INVALIDO
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
        sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "CTCuenta_CodCuentaSuma", MSGERR_NOHAYCODIGO
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
        Err.Raise ERR_INVALIDO, "CTCuenta_TipoCuenta", MSGERR_INVALIDO
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
            sql = "SELECT CodCuenta FROM CTCuenta WHERE IdCuentaSuma=" & mProps.IdCuenta
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            'Si ya tiene hijos
            If Not rs.EOF Then
                Err.Raise ERR_INVALIDO, "CTCuenta_BandTotal", _
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
                Err.Raise ERR_INVALIDO, "CTCuenta.BandTotal", _
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



Public Property Let bandValida(ByVal value As Boolean)
    mProps.bandValida = value
    mbooModificado = True
End Property

Public Property Get bandValida() As Boolean
    bandValida = mProps.bandValida
End Property

Public Property Let IDLocal(ByVal value As Long)
'    If Len(value) > 2 Then
'        Err.Raise ERR_INVALIDO, "CTCuenta_Local", MSGERR_INVALIDO
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
        .bandValida = True
        .BandGastoDistribuido = False
        .TipoCuentaFE = 0
        .BandDeudorFE = True
        .TipoFormulaFE = 0
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
            sql = "SELECT IdLocal FROM CtLocal WHERE CodLocal='" & value & "'"
            Set rs = mProps.objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "CtLocal_CodLocal", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IDLocal = rs!IDLocal
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

Public Property Let IdCuentaSC(ByVal value As Long)
    mProps.IdCuentaSC = value
    mbooModificado = True
End Property

Public Property Get IdCuentaSC() As Long
    IdCuentaSC = mProps.IdCuentaSC
End Property

Public Property Let CodCuentaSC(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCuentaSC Then
        'Actualiza IdCuentaSC también
        sql = "SELECT IdCuenta FROM CTCuentaSC WHERE CodCuenta='" & value & "'"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "CTCuenta_CodCuentaSC", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            IdCuentaSC = rs!IdCuenta
        End If
        rs.Close
    End If
    
    Set rs = Nothing
    mProps.CodCuentaSC = value
    mbooModificado = True
End Property

Public Property Get CodCuentaSC() As String
    CodCuentaSC = Trim$(mProps.CodCuentaSC)
End Property

Public Property Let Campo101(ByVal value As Currency)
    mProps.Campo101 = value
    mbooModificado = True
End Property

Public Property Get Campo101() As Currency
    Campo101 = mProps.Campo101
End Property

Public Property Let IdCuentaFE(ByVal value As Long)
    mProps.IdCuentaFE = value
    mbooModificado = True
End Property

Public Property Get IdCuentaFE() As Long
    IdCuentaFE = mProps.IdCuentaFE
End Property

Public Property Let CodCuentaFE(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCuentaFE Then
        If value = "" Then
            IdCuentaFE = 0
            Exit Property
        End If
        'Actualiza IdCuentaFE también
        sql = "SELECT IdCuenta FROM CTCuentaFE WHERE CodCuenta='" & value & "'"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "CTCuenta_CodCuentaFE", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            IdCuentaFE = rs!IdCuenta
        End If
        rs.Close
    End If
    
    Set rs = Nothing
    mProps.CodCuentaFE = value
    mbooModificado = True
End Property

Public Property Get CodCuentaFE() As String
    CodCuentaFE = Trim$(mProps.CodCuentaFE)
End Property

Public Property Let TipoCuentaFE(ByVal value As Byte)
    If value < 0 Or value > 3 Then
        Err.Raise ERR_INVALIDO, "CTCuenta_TipoCuentaFE", MSGERR_INVALIDO
    End If
    
    'Cambia el valor de BandDeudorfe
    Select Case value
    Case 0, 1, 5
        mProps.BandDeudorFE = True
    Case 2, 3, 4
        mProps.BandDeudorFE = False
    End Select
    
    mProps.TipoCuentaFE = value
    mbooModificado = True
End Property

Public Property Get TipoCuentaFE() As Byte
    TipoCuentaFE = mProps.TipoCuentaFE
End Property

Public Property Get BandDeudorFE() As Boolean
    BandDeudorFE = mProps.BandDeudorFE
End Property


Public Property Let FormulaCuentaFE(ByVal value As String)
    If Len(value) > 256 Then
        Err.Raise ERR_INVALIDO, "CTCuenta_FormulaCuentaFE", MSGERR_INVALIDO
    End If
    
    mProps.FormulaCuentaFE = value
    mbooModificado = True
End Property

Public Property Get FormulaCuentaFE() As String
    FormulaCuentaFE = Trim$(mProps.FormulaCuentaFE)
End Property


Public Property Let TipoFormulaFE(ByVal value As Byte)
    If value < 0 Or value > 3 Then
        Err.Raise ERR_INVALIDO, "CTCuenta_TipoCuentaFE", MSGERR_INVALIDO
    End If
    
    mProps.TipoFormulaFE = value
    mbooModificado = True
End Property

Public Property Get TipoFormulaFE() As Byte
    TipoFormulaFE = mProps.TipoFormulaFE
End Property


Public Property Let CampoF101(ByVal value As Currency)
    mProps.CampoF101 = value
    mbooModificado = True
End Property

Public Property Get CampoF101() As Currency
    CampoF101 = mProps.CampoF101
End Property

