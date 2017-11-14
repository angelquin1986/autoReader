VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GnContratoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdVehiculo As Long
    CodVehiculo As String
    Descripcion As String
    Cantidad As Integer
    Orden As Integer
    objGNContrato As GNContrato          'Referencia al objeto padre
    objGNtrans As GNTrans
    
    
    
End Type
Private mbooEsNuevoDet As Boolean
Private mbooModificado As Boolean

Private mProps As T_PROPS

Friend Property Set GNContrato(obj As GNContrato)
    Set mProps.objGNContrato = obj
End Property

Public Property Get GNContrato() As GNContrato
    Set GNContrato = mProps.objGNContrato
End Property

Friend Property Let IdVehiculo(ByVal value As Long)
    mProps.IdVehiculo = value
End Property

Friend Property Get IdVehiculo() As Long
    IdVehiculo = mProps.IdVehiculo
End Property


Public Property Let CodVehiculo(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNContrato_CodVehiculo ", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodVehiculo = value
    mbooModificado = True
    'mbooModificado = True
End Property

Public Property Get CodVehiculo() As String
    CodVehiculo = Trim$(mProps.CodVehiculo)
End Property


Public Property Let Cantidad(ByVal value As Currency)
    mProps.Cantidad = value
    mbooModificado = True    'trabaja con  la bandera modificado del papa
End Property

Public Property Get Cantidad() As Currency
    Cantidad = mProps.Cantidad
End Property

Private Sub Class_Terminate()
    Set mProps.objGNContrato = Nothing
End Sub


Public Property Get Descripcion() As String 'Propiedad  solo de lectura
    
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "GNContrato_Descripcion", MSGERR_INVALIDO
    End If
    
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, ";") > 0 Then
        Err.Raise ERR_INVALIDO, "GNContrato_Descripcion", _
            "El caracter ';' no es válido para la descripción."
    End If
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, "|") > 0 Then
        Err.Raise ERR_INVALIDO, "GNContrato_Descripcion", _
            "El caracter '|' no es válido para la descripción."
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Sub Recuperar(key As Variant)
   Dim sql As String, i As Integer
    Dim rs As Recordset
    Dim cond As String
    cond = " IV.CodVehiculo ='" & key & "'"
    If VarType(key) = vbString Then
        sql = "Select IV.IdVehiculo ,  IV.CodVehiculo , IV.Descripcion  " & _
        "FROM GNVehiculo IV  WHERE" & cond
    Else
        sql = "Select IV.IdVehiculo ,  IV.CodVehiculo , IV.Descripcion  " & _
        "FROM IvVehiculo IV   WHERE " & cond
    End If
    Set rs = Me.GNContrato.Empresa.OpenRecordset(sql)
    If rs.EOF Then 'AUC cambiado antes estuvo NOT (RS.EOF)
       Err.Raise ERR_NOHAYCODIGO, "GNContrtoDetalle", MSGERR_NOHAYCODIGO
    Else
        mProps.IdVehiculo = rs!IdVehiculo
        mProps.CodVehiculo = rs!CodVehiculo
        mProps.Descripcion = rs!Descripcion

  End If
End Sub


Friend Property Set GNTrans(obj As GNTrans)
    Set mProps.objGNtrans = obj
End Property

Public Property Get GNTrans() As GNTrans
    Set GNTrans = mProps.objGNtrans
End Property


Public Property Let Orden(ByVal value As Integer)
    mProps.Orden = value
    mbooModificado = True
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property



