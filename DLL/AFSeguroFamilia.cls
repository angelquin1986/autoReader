VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AFSeguroFamilia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdInventario As Long
    CodInventario As String
    IdTipoSeguro As Long
    CodTipoSeguro As String
    DescTipoSeguro As String
    IdProveedor As Long
    CodProveedor As String
    NombreProveedor As String
    Referencia As String
    Cantidad As Currency
    ValorAsegurado As Currency 'AUC 28/09/07 para promociones
    FechaInicio As Date
    FechaFin As Date
    objAFInventario As AFInventario          'Referencia al objeto padre
    objGNtrans As GNTrans
    
End Type
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Private mProps As T_PROPS

Friend Property Set AFInventario(obj As AFInventario)
    Set mProps.objAFInventario = obj
End Property

Public Property Get AFInventario() As AFInventario
    Set AFInventario = mProps.objAFInventario
End Property

Friend Property Let IdInventario(ByVal value As Long)
    mProps.IdInventario = value
End Property

Friend Property Get IdInventario() As Long
    IdInventario = mProps.IdInventario
End Property

Public Property Let CodInventario(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "AFInventario_CodInventario", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodInventario = value
    mProps.objAFInventario.Modificado = True
    'mbooModificado = True
End Property

Public Property Get CodInventario() As String
    CodInventario = Trim$(mProps.CodInventario)
End Property


Friend Property Let IdTipoSeguro(ByVal value As Long)
    mProps.IdTipoSeguro = value
End Property

Friend Property Get IdTipoSeguro() As Long
    IdTipoSeguro = mProps.IdTipoSeguro
End Property


Public Property Let CodTipoSeguro(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodTipoSeguro Then
        'Actualiza IdCuenta también
        sql = "SELECT IdTipoSeguro, Descripcion FROM AFTipoSeguro WHERE CodTipoSeguro='" & value & "'"
        'Set rs = mProps.Empresa.OpenRecordset(sql)
        Set rs = Me.AFInventario.Empresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "AFTipoSeguro_CodTipoSeguro", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            mProps.IdTipoSeguro = rs!IdTipoSeguro
            mProps.DescTipoSeguro = rs!Descripcion
        End If
        rs.Close
        Set rs = Nothing
        
        mProps.CodTipoSeguro = value
        mProps.objAFInventario.Modificado = True
    End If
End Property

Public Property Get CodTipoSeguro() As String
    CodTipoSeguro = Trim$(mProps.CodTipoSeguro)
End Property


Friend Property Let IdProveedor(ByVal value As Long)
    mProps.IdProveedor = value
End Property

Friend Property Get IdProveedor() As Long
    IdProveedor = mProps.IdProveedor
End Property


Public Property Let CodProveedor(ByVal value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If value <> mProps.CodProveedor Then
        'Actualiza IdCuenta también
        sql = "SELECT IdProvCli, Nombre FROM PCProvCli WHERE CodProvCli='" & value & "'"
        'Set rs = mProps.Empresa.OpenRecordset(sql)
        Set rs = Me.AFInventario.Empresa.OpenRecordset(sql)
        If rs.EOF Then
            Err.Raise ERR_NOHAYCODIGO, "PCProvCli_CodProvCli", MSGERR_NOHAYCODIGO
            Exit Property
        Else
            mProps.IdProveedor = rs!IdProvCli
            mProps.NombreProveedor = rs!Nombre
        End If
        rs.Close
        Set rs = Nothing
        
        mProps.CodProveedor = value
        mProps.objAFInventario.Modificado = True
    End If
End Property

Public Property Get CodProveedor() As String
    CodProveedor = Trim$(mProps.CodProveedor)
End Property


Public Property Let FechaInicio(value As Date)
    If Not (IsDate(value)) Then
        Err.Raise ERR_INVALIDO, "objAFInventario_FechaInicio", MSGERR_INVALIDO
    End If
    mProps.FechaInicio = value
    mProps.objAFInventario.Modificado = True
End Property

Public Property Get FechaInicio() As Date
    FechaInicio = mProps.FechaInicio
End Property

Public Property Let FechaFin(value As Date)
    If Not (IsDate(value)) Then
        Err.Raise ERR_INVALIDO, "objAFInventario_FechaFin", MSGERR_INVALIDO
    End If
    
    mProps.FechaFin = value
    mProps.objAFInventario.Modificado = True
End Property

Public Property Get FechaFin() As Date
    FechaFin = mProps.FechaFin
End Property

Private Sub Class_Terminate()
    Set mProps.objAFInventario = Nothing
End Sub


 'key puede ser IdInventario  o CodInventario
Public Sub Recuperar(key As Variant)
   Dim sql As String, i As Integer
    Dim rs As Recordset
    Dim cond As String
    cond = " AF.CodInventario='" & key & "'"
        If VarType(key) = vbString Then
            sql = "Select AF.IdInventario,  AF.CodInventario, AF.Descripcion, U.CodUnidad " & _
            "FROM AFInventario AF WHERE" & cond
        Else
            sql = "Select AF.IdInventario,  AF.CodInventario, AF.Descripcion  " & _
            "FROM AFInventario AF WHERE " & cond
        End If
    Set rs = Me.AFInventario.Empresa.OpenRecordset(sql)
    If rs.EOF Then 'AUC cambiado antes estuvo NOT (RS.EOF)
       Err.Raise ERR_NOHAYCODIGO, "IvFamiliaDetalle", MSGERR_NOHAYCODIGO
    Else
        mProps.IdTipoSeguro = rs!IdTipoSeguro
        mProps.CodTipoSeguro = rs!CodTipoSeguro
        mProps.IdProveedor = rs!IdProveedor
        mProps.CodProveedor = rs!CodProveedor
        If Not IsNull(rs!Cantidad) Then mProps.Cantidad = rs!Cantidad
        If Not IsNull(rs!ValorAsegurado) Then mProps.ValorAsegurado = rs!ValorAsegurado
        If Not IsNull(rs!Referencia) Then mProps.Referencia = rs!Referencia
        mProps.FechaInicio = rs!FechaInicio
        mProps.FechaFin = rs!FechaFin
  End If
End Sub


Friend Property Set GNTrans(obj As GNTrans)
    Set mProps.objGNtrans = obj
End Property

Public Property Get GNTrans() As GNTrans
    Set GNTrans = mProps.objGNtrans
End Property

'key puede ser CodTipoSeguro o Descripcion
Public Sub RecuperarPreparacionDescripcion(key As Variant)
   Dim sql As String, i As Integer
    Dim rs As Recordset
    Dim cond As String
        cond = " IV.descripcion='" & key & "'"
       If VarType(key) = vbString Then
            sql = "Select IV.IdTipoSeguro,  IV.CodTipoSeguro, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
            "FROM AFInventario IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad WHERE" & cond
        Else
            sql = "Select IV.IdTipoSeguro,  IV.CodTipoSeguro, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
            "FROM AFInventario IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad  WHERE " & cond
        End If
   Set rs = Me.AFInventario.Empresa.OpenRecordset(sql)
    If rs.EOF Then 'AUC cambiado antes estuvo NOT (RS.EOF)
       Err.Raise ERR_NOHAYCODIGO, "IvFamiliaDetalle", MSGERR_NOHAYCODIGO
    Else
        mProps.IdTipoSeguro = rs!IdTipoSeguro
        mProps.CodTipoSeguro = rs!CodTipoSeguro
        mProps.IdProveedor = rs!IdProveedor
        mProps.CodProveedor = rs!CodProveedor
        If Not IsNull(rs!Cantidad) Then mProps.Cantidad = rs!Cantidad
        If Not IsNull(rs!ValorAsegurado) Then mProps.ValorAsegurado = rs!ValorAsegurado
  End If
End Sub


'AUC 26/09/07 para promociones
Public Property Let ValorAsegurado(ByVal value As Currency)
    mProps.ValorAsegurado = value
    mProps.objAFInventario.Modificado = True   'trabaja con  la bandera modificado del papa
End Property

Public Property Get ValorAsegurado() As Currency
    ValorAsegurado = mProps.ValorAsegurado
End Property

Public Property Let Cantidad(ByVal value As Currency)
    mProps.Cantidad = value
    mProps.objAFInventario.Modificado = True   'trabaja con  la bandera modificado del papa
End Property

Public Property Get Cantidad() As Currency
    Cantidad = mProps.Cantidad
End Property

Public Property Let Referencia(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "AFInventario_Referencia", MSGERR_INVALIDO
    End If
    mProps.Referencia = value
    mProps.objAFInventario.Modificado = True   'trabaja con  la bandera modificado del papa
End Property

Public Property Get Referencia() As String
    Referencia = mProps.Referencia
End Property


Public Property Get DescTipoSeguro() As String 'Propiedad  solo de lectura
    DescTipoSeguro = Trim$(mProps.DescTipoSeguro)
End Property

Public Property Let DescTipoSeguro(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "IVInventario_DescTipoSeguro", MSGERR_INVALIDO
    End If
    
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, ";") > 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario_DescTipoSeguro", _
            "El caracter ';' no es válido para la descripción."
    End If
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, "|") > 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario_DescTipoSeguro", _
            "El caracter '|' no es válido para la descripción."
    End If
    
    mProps.DescTipoSeguro = value
    mbooModificado = True
End Property


Public Property Get NombreProveedor() As String 'Propiedad  solo de lectura
    NombreProveedor = Trim$(mProps.NombreProveedor)
End Property

Public Property Let NombreProveedor(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "IVInventario_NombreProveedor", MSGERR_INVALIDO
    End If
    
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, ";") > 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario_NombreProveedor", _
            "El caracter ';' no es válido para la descripción."
    End If
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, "|") > 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario_NombreProveedor", _
            "El caracter '|' no es válido para la descripción."
    End If
    
    mProps.NombreProveedor = value
    mbooModificado = True
End Property




