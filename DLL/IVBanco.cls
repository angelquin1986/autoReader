VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private Type T_PROPS
    IdBanco As Long
    CodBanco As String
    Descripcion As String
    IDForma As Long
    CodForma As String
    IdCliente As Long
    CodCliente As String
    BandValida As Boolean
    fechagrabado As Date
    Factor As Integer
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
    sql = "SELECT CodBanco FROM IVBanco " & _
          "WHERE (CodBanco='" & mProps.CodBanco & "') AND " & _
                "(IdBanco<>" & mProps.IdBanco & ")"
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then      'Si ya existe el codigo
        Err.Raise ERR_REPITECODIGO, "IVBanco", MSGERR_REPITECODIGO
    End If
    rs.Close
    
    'Si es nuevo
    If mbooEsNuevo Then
        sql = "SELECT * FROM IVBanco WHERE 1=0"
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
        rs.AddNew
    
    'Si es modificada
    Else
        'Recupera el registro correspondiente para modificar
        sql = "SELECT * FROM IVBanco WHERE IdBanco=" & mProps.IdBanco
        Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql)
#If DAOLIB Then
        rs.Edit
#End If
        
        'Verifica la FechaGrabado para asegurar que otros usuarios no ha modificado el mismo registro
        If rs!fechagrabado <> mProps.fechagrabado Then
            Err.Raise ERR_MODIFICADO, "IVBanco_Grabar", MSGERR_MODIFICADO
        End If
    End If

    'Graba el registro en Pos
    With rs
        !CodBanco = mProps.CodBanco
        !Descripcion = mProps.Descripcion
'        !IdForma = mProps.IdForma
'        !CodForma = mProps.CodForma
        !BandValida = mProps.BandValida
        !IdCliente = mProps.IdCliente
        !Factor = mProps.Factor
        !fechagrabado = Now
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.fechagrabado = !fechagrabado
        mProps.IdBanco = !IdBanco
        .Close
    End With
    
    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    If mbooEsNuevo Then
        mProps.objEmpresa.GrabaGNLogAccion "NUE_BCO", "Banco Nuevo: " & _
                                                         mProps.CodBanco, "DT"
    Else
        mProps.objEmpresa.GrabaGNLogAccion "MOD_BCO", "Banco Modificado: " & _
                                                         mProps.CodBanco, "DT"
    End If
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVBANCO
End Sub

Friend Sub Recuperar(key As Variant)
    Dim sql As String, rs As Recordset
    
    If VarType(key) = vbString Then
        sql = "SELECT IVBanco.*, CodProvcli as CodCliente "
        sql = sql & " FROM IVBanco "
        sql = sql & " LEFT JOIN PCProvcli "
        sql = sql & "  ON IVBanco.IDcliente = Pcprovcli.IDProvcli"
        sql = sql & "  WHERE CodBanco='" & key & "'"
    Else
        sql = "SELECT IVBanco.*, CodProvcli as CodCliente "
        sql = sql & " FROM IVBanco "
        sql = sql & " LEFT JOIN PCProvcli "
        sql = sql & "  ON IVBanco.IDcliente = Pcprovcli.IDProvcli"
        sql = sql & "  WHERE IdBanco=" & key
    End If
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    
    'Recupera los datos de FormaCopro
    With rs
        mProps.IdBanco = !IdBanco
        mProps.CodBanco = !CodBanco
        If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
        If Not IsNull(!CodForma) Then mProps.CodForma = !CodForma
        If Not IsNull(!IdCliente) Then mProps.IdCliente = !IdCliente
        If Not IsNull(!CodCliente) Then mProps.CodCliente = !CodCliente
       
        If Not IsNull(!BandValida) Then mProps.BandValida = !BandValida
        If Not IsNull(!Factor) Then mProps.Factor = !Factor
       
        mProps.fechagrabado = !fechagrabado
        .Close
    End With
    
    Set rs = Nothing
    mbooEsNuevo = False
    mbooModificado = False
End Sub

Public Sub Eliminar()
    Dim sql As String, rs As Recordset

    'Verifica si existe Kardex del Banco
    sql = "SELECT IdBanco FROM pcKardex WHERE IdBanco=" & mProps.IdBanco
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "IVBanco.Eliminar", MSGERR_NOELIMINABLE
    End If
    Set rs = Nothing

    'Elimina el Banco
    sql = "DELETE FROM IVBanco WHERE CodBanco='" & mProps.CodBanco & "'"
    mProps.objEmpresa.Execute sql, True

    'Graba un registro para auditoria   *** ANGEL 20/mar/02
    mProps.objEmpresa.GrabaGNLogAccion "ELI_BAN", "Banco Eliminado: " & _
                                                         mProps.CodBanco, "DT"
                                                         
    mbooEsNuevo = True
    mbooModificado = True
    
    'Genera evento Refresh del objeto Empresa
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_IVBANCO
End Sub

Public Function Clone() As IVBanco
    Dim nuevo As IVBanco
    
    Set nuevo = New IVBanco
    'Copia las propiedades al nuevo objeto
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdBanco = 0
    mbooModificado = True
End Property

Friend Property Set Empresa(objEmpresa As Empresa)
    Set mProps.objEmpresa = objEmpresa
End Property

Public Property Get Empresa() As Empresa
    Set Empresa = mProps.objEmpresa
End Property


Public Property Let CodBanco(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "IVBanco_CodBanco", MSGERR_INVALIDO
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
        Err.Raise ERR_INVALIDO, "IVBanco_Descripcion", MSGERR_INVALIDO
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let IDForma(ByVal value As Long)
    mProps.IDForma = value
    mbooModificado = True
End Property

Public Property Get IDForma() As Long
    IDForma = mProps.IDForma
End Property


Public Property Let CodForma(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    
    If Len(value) > 50 Then
        Err.Raise ERR_INVALIDO, "IVBanco_CodForma", MSGERR_INVALIDO
    End If
    
    mProps.CodForma = value
    mbooModificado = True
   
    
''''    'Cuando cambia el valor
''''    If value <> mProps.CodForma Then
''''        'Actualiza IdForma también
''''        sql = "SELECT IdForma FROM TsFormaCobroPago WHERE CodForma='" & value & "'"
''''        Set rs = mProps.objEmpresa.OpenRecordset(sql)
''''        If rs.EOF Then
''''            Err.Raise ERR_NOHAYCODIGO, "IVBanco_CodForma", MSGERR_NOHAYCODIGO
''''            Exit Property
''''        Else
''''            IdForma = rs!IdForma
''''        End If
''''        rs.Close
''''    End If

    Set rs = Nothing
    mProps.CodForma = value
    mbooModificado = True
End Property

Public Property Get CodForma() As String
    CodForma = Trim$(mProps.CodForma)
End Property


Public Property Let BandValida(ByVal value As Boolean)
    mProps.BandValida = value
    mbooModificado = True
End Property

Public Property Get BandValida() As Boolean
    BandValida = mProps.BandValida
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


Private Sub Class_Initialize()
    With mProps
        .BandValida = True
        .Factor = 1
    End With
    
    mbooEsNuevo = True
    mbooModificado = False
End Sub

Private Sub Class_Terminate()
    Set mProps.objEmpresa = Nothing
End Sub

Public Property Let IdCliente(ByVal value As Long)
    mProps.IdCliente = value
    mbooModificado = True
End Property

Public Property Get IdCliente() As Long
    IdCliente = mProps.IdCliente
End Property


Public Property Let CodCliente(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodCliente Then
        'Actualiza IdCliente también
        sql = "SELECT IdProvCli FROM Pcprovcli WHERE CodProvCli='" & value & "'"
        Set rs = mProps.objEmpresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "IVBanco_CodCliente", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            IdCliente = rs!IdProvCli
        End If
        rs.Close
    End If

    Set rs = Nothing
    mProps.CodCliente = value
    mbooModificado = True
End Property

Public Property Get CodCliente() As String
    CodCliente = Trim$(mProps.CodCliente)
End Property

Public Property Let Factor(ByVal value As Long)
    mProps.Factor = value
    mbooModificado = True
End Property

Public Property Get Factor() As Long
    Factor = mProps.Factor
End Property

