VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TSBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Empresa"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Private Type T_PROPS
    IDBanco As Long
    CodBanco As String
    Descripcion As String
    IdCuentaContable As Long
    CodCuentaContable As String
    NumCuenta As String
    Nombre As String
    BandValida As Boolean
    NumChequeSiguiente As Long  'jeaa secuencial de cheque
    FechaGrabado As Date
    objEmpresa As Empresa       'Referencia al objeto padre
End Type

Private mProps As T_PROPS
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean


Public Sub Grabar()
    Dim sql As String, rs As Recordset
    
    'Si no está modificado no hace nada
    If Not mbooModificado Then Exit Sub
    
    'Verifica si no repite el código
    sql = "SELECT CodBanco FROM TSBanco " & _
          "WHERE (CodBanco='" & mProps.CodBanco & "') AND " & _
                "(IdBanco<>" & mProps.IDBanco & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "TSBanco", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM TSBanco WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM TSBanco WHERE IdBanco=" & mProps.IDBanco
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!FechaGrabado <> mProps.FechaGrabado Then
            Err.Raise ERR_MODIFICADO, "TSBanco_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en Pos
    With rs
        !CodBanco = mProps.CodBanco
        !Descripcion = mProps.Descripcion
        !IdCuentaContable = mProps.IdCuentaContable
        !NumCuenta = mProps.NumCuenta
        !Nombre = mProps.Nombre
        !BandValida = mProps.BandValida
        !NumChequeSiguiente = mProps.NumChequeSiguiente
        !NumChequeSiguiente = mProps.NumChequeSiguiente
        
        !FechaGrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.FechaGrabado = !FechaGrabado
        mProps.IDBanco = !IDBanco
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_BANCO", "Banco Nuevo: " & _
                                                         mProps.CodBanco, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_BANCO", "Banco Modificado: " & _
                                                         mProps.CodBanco, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_BANCO
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT * FROM vwTSBanco WHERE CodBanco='" & key & "'"
    Else
        sql = "SELECT * FROM vwTSBanco WHERE IdBanco=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de Usuario
    With rs
        mProps.IDBanco = !IDBanco
        mProps.CodBanco = !CodBanco
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!IdCuentaContable) Then mProps.IdCuentaContable = !IdCuentaContable
        If Not IsNull(!CodCuentaContable) Then mProps.CodCuentaContable = !CodCuentaContable
        If Not IsNull(!NumCuenta) Then mProps.NumCuenta = !NumCuenta
        If Not IsNull(!Nombre) Then mProps.Nombre = !Nombre
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!NumChequeSiguiente) Then mProps.NumChequeSiguiente = !NumChequeSiguiente
        
        mProps.FechaGrabado = !FechaGrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset

    'Verifica si existe Kardex del banco
    sql = "SELECT IdBanco FROM TSKardex WHERE IdBanco=" & mProps.IDBanco
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "TSBanco.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing

    'Elimina el banco
    sql = "DELETE FROM TSBanco WHERE CodBanco='" & mProps.CodBanco & "'"
    mProps.objEmpresa.Execute sql, True

    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_BAN", "Banco Eliminado: " & _
                                                         mProps.CodBanco, "DT"
                                                         
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_BANCO
End Sub

Public Function Clone() As TSBanco
    Dim nuevo As TSBanco
    
    Set nuevo = New TSBanco
    'Copia las propiedades al nuevo objeto
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IDBanco = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property

Public Function CalculaSaldo(FechaHasta As Date) As Currency
    Dim sql As String, rs As Recordset
    
    sql = "SELECT SUM(Debe)-SUM(Haber) AS Saldo " & _
          "FROM TSKardex INNER JOIN GNComprobante ON TSKardex.TransID=GNComprobante.TransID " & _
          "WHERE BandAnulado=" & CadenaBool(False, mProps.objEmpresa.TipoDB) & _
          " AND TSKardex.FechaVenci<=" & FechaYMD(FechaHasta, mProps.objEmpresa.TipoDB)
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not IsNull(rs!Saldo) Then
        CalculaSaldo = rs!Saldo
    Else
        CalculaSaldo = 0
    End If
    Set rs = Nothing
End Function


Public Property Let CodBanco(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "TSBanco_CodBanco", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodBanco = value
    mbooModificado = True
End Property

Public Property Get CodBanco() As String
    CodBanco = Trim$(mProps.CodBanco)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "TSBanco_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let IdCuentaContable(ByVal value As Long)
    mProps.IdCuentaContable = value
    mbooModificado = True
End Property

Public Property Get IdCuentaContable() As Long
    IdCuentaContable = mProps.IdCuentaContable
End Property


Public Property Let CodCuentaContable(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCuentaContable Then
        'Actualiza IdCuentaContable también
        sql = "SELECT IdCuenta FROM CTCuenta WHERE CodCuenta='" & value & "'"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "TSBanco_CodCuentaContable", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            IdCuentaContable = rs!IdCuenta
        End If
        rs.Close
    End If

    Set rs = Nothing
    mProps.CodCuentaContable = value
    mbooModificado = True
End Property

Public Property Get CodCuentaContable() As String
    CodCuentaContable = Trim$(mProps.CodCuentaContable)
End Property

Public Property Let NumCuenta(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "TSBanco_NumCuenta", MSGERR_INVALIDO
    End If
    
    mProps.NumCuenta = value
    mbooModificado = True
End Property

Public Property Get NumCuenta() As String
    NumCuenta = Trim$(mProps.NumCuenta)
End Property

Public Property Let Nombre(ByVal value As String)
    If Len(value) > 40 Then
        Err.Raise ERR_INVALIDO, "TSBanco_Nombre", MSGERR_INVALIDO
    End If
    
    mProps.Nombre = value
    mbooModificado = True
End Property

Public Property Get Nombre() As String
    Nombre = Trim$(mProps.Nombre)
End Property

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
    With mProps
        .BandValida = True
    End With
    
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Property Let NumChequeSiguiente(ByVal value As Long)
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "TSBANCO_NumChequeSiguiente", MSGERR_INVALIDO
    End If
    
    mProps.NumChequeSiguiente = value
    mbooModificado = True
End Property

Public Property Get NumChequeSiguiente() As Long
    NumChequeSiguiente = mProps.NumChequeSiguiente
End Property

