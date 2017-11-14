VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GNCentroCostoDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdCentroCostoDetalle As Long
    CodCentroDet As String
    DescripcionDet As String
    Cantidad As Integer
    Orden As Integer
    objGNCentro As GNCentroCosto             'Referencia al objeto padre
End Type
Private mbooEsNuevoDet As Boolean
Private mbooModificado As Boolean
Private mProps As T_PROPS

Friend Property Set GNCentroCosto(obj As GNCentroCosto)
    Set mProps.objGNCentro = obj
End Property

Public Property Get GNCentroCosto() As GNCentroCosto
    Set GNCentroCosto = mProps.objGNCentro
End Property

Friend Property Get IdCentroCostoDetalle() As Long
    IdCentroCostoDetalle = mProps.IdCentroCostoDetalle
End Property

Public Property Let CodCentroDet(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "GNContrato_CodCentroDet ", MSGERR_INVALIDO
    End If
    ValidaCodigo value      'Valida si contiene caracteres no validos
    
    mProps.CodCentroDet = value
    mbooModificado = True
    'mbooModificado = True
End Property

Public Property Get CodCentroDet() As String
    CodCentroDet = Trim$(mProps.CodCentroDet)
End Property


Private Sub Class_Terminate()
    Set mProps.objGNCentro = Nothing
End Sub


Public Property Get DescripcionDet() As String 'Propiedad  solo de lectura
    DescripcionDet = Trim$(mProps.DescripcionDet)
End Property

Public Property Let DescripcionDet(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "GNContrato_DescripcionDet", MSGERR_INVALIDO
    End If
    
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, ";") > 0 Then
        Err.Raise ERR_INVALIDO, "GNContrato_DescripcionDet", _
            "El caracter ';' no es válido para la descripción."
    End If
    '*** MAKOTO 23/jul/00 Agregado
    If InStr(value, "|") > 0 Then
        Err.Raise ERR_INVALIDO, "GNContrato_DescripcionDet", _
            "El caracter '|' no es válido para la descripción."
    End If
    
    mProps.DescripcionDet = value
    mbooModificado = True
End Property

Public Sub Recuperar(key As Variant)
   Dim sql As String, i As Integer
    Dim rs As Recordset
    Dim cond As String
    cond = " gnc.CodCentroDet ='" & key & "'"
    If VarType(key) = vbString Then
        sql = "Select gnc.IdCentro,  gnc CodCentroDet , gnc.DescripcionDet  " & _
        "FROM GNCentroCostoDetalle gnc  WHERE" & cond
    Else
        sql = "Select IV.IdCentro, gnc.CodCentroDet , gnc.DescripcionDet  " & _
        "FROM GNCentroCostoDetalle gnc   WHERE " & cond
    End If
    Set rs = Me.GNCentroCosto.Empresa.OpenRecordset(sql)
    If rs.EOF Then 'AUC cambiado antes estuvo NOT (RS.EOF)
       Err.Raise ERR_NOHAYCODIGO, "GNCentroDetalle", MSGERR_NOHAYCODIGO
    Else
        mProps.IdCentroCostoDetalle = rs!IdCentroCostoDetalle
        mProps.CodCentroDet = rs!CodCentroDet
        mProps.DescripcionDet = rs!DescripcionDet

  End If
End Sub


Public Property Let Orden(ByVal value As Integer)
    mProps.Orden = value
    mbooModificado = True
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property

