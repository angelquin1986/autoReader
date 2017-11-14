VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GnProyectoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdInventario As Long
    CodInventario As String
    Descripcion As String
    Cantidad As Currency
    Orden As Integer
    objGNProyecto As GNProyecto          'Referencia al objeto padre
    objGNtrans As GNTrans
    
    
    
End Type
Private mbooEsNuevoDet As Boolean
Private mbooModificado As Boolean

Private mProps As T_PROPS

Friend Property Set GNProyecto(obj As GNProyecto)
    Set mProps.objGNProyecto = obj
End Property

Public Property Get GNProyecto() As GNProyecto
    Set GNProyecto = mProps.objGNProyecto
End Property

Friend Property Let IdInventario(ByVal value As Long)
    mProps.IdInventario = value
End Property

Friend Property Get IdInventario() As Long
    IdInventario = mProps.IdInventario
End Property


Public Property Let CodInventario(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNProyecto_CodInventario", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodInventario = value
    mbooModificado = True
    'mbooModificado = True
End Property

Public Property Get CodInventario() As String
    CodInventario = Trim$(mProps.CodInventario)
End Property


Public Property Let Cantidad(ByVal value As Currency)
    mProps.Cantidad = value
    mbooModificado = True    'trabaja con  la bandera modificado del papa
End Property

Public Property Get Cantidad() As Currency
    Cantidad = mProps.Cantidad
End Property

Private Sub Class_Terminate()
    Set mProps.objGNProyecto = Nothing
End Sub


Public Property Get Descripcion() As String 'Propiedad  solo de lectura
    
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "GNProyecto_Descripcion", MSGERR_INVALIDO
    End If
    
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, ";") > 0 Then
        Err.Raise ERR_INVALIDO, "GNProyecto_Descripcion", _
            "El caracter ';' no es válido para la descripción."
    End If
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, "|") > 0 Then
        Err.Raise ERR_INVALIDO, "GNProyecto_Descripcion", _
            "El caracter '|' no es válido para la descripción."
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

Public Sub Recuperar(key As Variant)
   Dim sql As String, i As Integer
    Dim rs As Recordset
    Dim cond As String
    cond = " IV.CodInventario='" & key & "'"
    If VarType(key) = vbString Then
        sql = "Select IV.IdInventario,  IV.CodInventario, IV.Descripcion  " & _
        "FROM IvInventario IV  WHERE" & cond
    Else
        sql = "Select IV.IdInventario,  IV.CodInventario, IV.Descripcion  " & _
        "FROM IvInventario IV   WHERE " & cond
    End If
    Set rs = Me.GNProyecto.Empresa.OpenRecordset(sql)
    If rs.EOF Then 'AUC cambiado antes estuvo NOT (RS.EOF)
       Err.Raise ERR_NOHAYCODIGO, "IvFamiliaDetalle", MSGERR_NOHAYCODIGO
    Else
        mProps.IdInventario = rs!IdInventario
        mProps.CodInventario = rs!CodInventario
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

