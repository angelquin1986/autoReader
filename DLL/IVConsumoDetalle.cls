VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVConsumoDetalle"
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
    Precio As Currency
    Costo As Currency
    Orden As Integer
    Ticket As Long
    objIVKardex As IVKardex 'Referencia al objeto padre
    objGNtrans As GNTrans
    
    
    
End Type
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Private mProps As T_PROPS

Friend Property Set IVKardex(obj As IVKardex)
    Set mProps.objIVKardex = obj
End Property

Public Property Get IVKardex() As IVKardex
    Set IVKardex = mProps.objIVKardex
End Property

Friend Property Let IdInventario(ByVal value As Long)
    mProps.IdInventario = value
End Property

Friend Property Get IdInventario() As Long
    IdInventario = mProps.IdInventario
End Property


Public Property Let CodInventario(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "IVInventario_CodInventario", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    mProps.CodInventario = value
End Property

Public Property Get CodInventario() As String
    CodInventario = Trim$(mProps.CodInventario)
End Property


Public Property Let Cantidad(ByVal value As Currency)
    mProps.Cantidad = value
End Property

Public Property Get Cantidad() As Currency
    Cantidad = mProps.Cantidad
End Property

Private Sub Class_Terminate()
    'Set mProps.objIVIKardex = Nothing
End Sub


Public Property Get Descripcion() As String 'Propiedad  solo de lectura
    
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "IVInventario_Descripcion", MSGERR_INVALIDO
    End If
    
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, ";") > 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario_Descripcion", _
            "El caracter ';' no es válido para la descripción."
    End If
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, "|") > 0 Then
        Err.Raise ERR_INVALIDO, "IVInventario_Descripcion", _
            "El caracter '|' no es válido para la descripción."
    End If
    
    mProps.Descripcion = value
    mbooModificado = True
End Property

 'key puede ser IdInventario  o CodInventario
Public Sub Recuperar(key As Variant)
   Dim sql As String, i As Integer
    Dim rs As Recordset
    Dim cond As String
    cond = " IV.CodInventario='" & key & "'"
    If VarType(key) = vbString Then
        sql = "Select IV.IdInventario,  IV.CodInventario, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
        "FROM IVInventario IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad WHERE" & cond
    Else
        sql = "Select IV.IdInventario,  IV.CodInventario, IV.Descripcion, U.CodUnidad,IV.costoultimoIngreso " & _
        "FROM IVInventario IV LEFT JOIN IVUnidad U ON U.idUnidad = IV.IdUnidad  WHERE " & cond
    End If
End Sub

Friend Property Set GNTrans(obj As GNTrans)
    Set mProps.objGNtrans = obj
End Property

Public Property Get GNTrans() As GNTrans
    Set GNTrans = mProps.objGNtrans
End Property

'AUC 23/03/06
Public Property Let Precio(ByVal value As Currency)
    mProps.Precio = value
End Property

Public Property Get Precio() As Currency
    Precio = mProps.Precio
End Property

Public Property Let Costo(ByVal value As Currency)
    mProps.Costo = value
End Property

Public Property Get Costo() As Currency
    Costo = mProps.Costo
End Property


Public Property Let Orden(ByVal value As Integer)
    mProps.Orden = value
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property



