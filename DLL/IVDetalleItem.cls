VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVDetalleItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Type T_PROPS
    IdProvCli As Long
    IdInventario As Long
    CodInventario As String
    Descripcion As String
    FechaGrabado As Date
    Cantidad As Currency
    PU As Currency
    fechaIni As Date
    FechaFin As Date
    BandProntoPago As Boolean
    Frecuencia As String
    Referencia As String
    BandPublicidad As Boolean 'AUC para diferencia y clasificar en una sola factura
    Contrato As String
    Plazo As Integer
    objPcProvCli As PCProvCli
End Type
Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean
Private mProps As T_PROPS

Friend Property Set PCProvCli(obj As PCProvCli)
    Set mProps.objPcProvCli = obj
End Property

Public Property Get PCProvCli() As PCProvCli
    Set PCProvCli = mProps.objPcProvCli
End Property

Public Property Let IdInventario(ByVal value As Long)
    mProps.IdInventario = value
End Property

Public Property Get IdInventario() As Long
    IdInventario = mProps.IdInventario
End Property


Public Property Get CodInventario() As String
    CodInventario = Trim$(mProps.CodInventario)
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Private Sub Class_Terminate()
    Set mProps.objPcProvCli = Nothing
End Sub

 'key puede ser IdInventario  o CodInventario
Public Sub Recuperar(key As Variant, key1 As Variant)
    Dim sql As String, i As Integer
    Dim rs As Recordset
    
    sql = "Select iv.codinventario,iv.descripcion FROM vwivinventariorecuperar  iv  WHERE "
    
    If Len(key) > 0 Then
        sql = sql & " codinventario='" & key & "'"
    ElseIf Len(key1) > 0 Then
        sql = sql & " Descripcion='" & key1 & "'"
    End If
    Set rs = mProps.objPcProvCli.Empresa.OpenRecordset(sql)
    If rs.EOF Then
        Err.Raise ERR_NOHAYCODIGO, "PCProvCliInventario", MSGERR_NOHAYCODIGO
    Else
        mProps.CodInventario = rs!CodInventario
        mProps.Descripcion = rs!Descripcion
        
    End If
End Sub

Public Property Get FechaGrabado() As Date
    FechaGrabado = mProps.FechaGrabado
End Property

Public Property Let FechaGrabado(ByVal value As Date)
    mProps.FechaGrabado = value
    'mProps.objPCGrupo.Modificado = True
End Property

Public Property Let CodInventario(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "IVDetallePcProvCli_CodInventario", MSGERR_INVALIDO
    End If
    
    mProps.CodInventario = value
    
End Property

Public Property Let Descripcion(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "IVDetallePcProvCli_Descripcion", MSGERR_INVALIDO
    End If
    mProps.Descripcion = value
    'mbooModificado = True
End Property
Public Property Let Cantidad(ByVal value As Currency)
    mProps.Cantidad = value
     mProps.objPcProvCli.SetModificado
End Property

Public Property Get Cantidad() As Currency
    Cantidad = mProps.Cantidad
End Property

Public Property Let PU(ByVal value As Currency)
    mProps.PU = value
    mProps.objPcProvCli.SetModificado
End Property

Public Property Get PU() As Currency
    PU = mProps.PU
End Property
Public Property Let IdProvCli(ByVal value As Long)
    mProps.IdProvCli = value
End Property

Public Property Get IdProvCli() As Long
    IdProvCli = mProps.IdProvCli
End Property

Public Property Let Frecuencia(ByVal value As String)
    If Len(value) > 30 Then
        Err.Raise ERR_INVALIDO, "IVDetallePcProvCli_Frecuencia", MSGERR_INVALIDO
    End If
    mProps.Frecuencia = value
    mProps.objPcProvCli.SetModificado
End Property

Public Property Get Frecuencia() As String
    Frecuencia = mProps.Frecuencia
End Property

Public Property Let fechaIni(ByVal value As Date)
    mProps.fechaIni = value
    mProps.objPcProvCli.SetModificado
End Property

Public Property Get fechaIni() As Date
    fechaIni = mProps.fechaIni
End Property

Public Property Let FechaFin(ByVal value As Date)
    mProps.FechaFin = value
    mProps.objPcProvCli.SetModificado
End Property

Public Property Get FechaFin() As Date
    FechaFin = mProps.FechaFin
End Property

Public Property Let BandProntoPago(ByVal value As Boolean)
    mProps.BandProntoPago = value
    mProps.objPcProvCli.SetModificado
End Property

Public Property Get BandProntoPago() As Boolean
    BandProntoPago = mProps.BandProntoPago
End Property

Public Property Let Referencia(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "IVDetallePcProvCli_Referencia", MSGERR_INVALIDO
    End If
    mProps.Referencia = value
    mProps.objPcProvCli.SetModificado
End Property

Public Property Get Referencia() As String
    Referencia = Trim$(mProps.Referencia)
End Property

Public Property Let BandPublicidad(ByVal value As Boolean)
    mProps.BandPublicidad = value
    mProps.objPcProvCli.SetModificado
End Property

Public Property Get BandPublicidad() As Boolean
    BandPublicidad = mProps.BandPublicidad
End Property

Public Property Let Contrato(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "IVDetallePcProvCli_Referencia", MSGERR_INVALIDO
    End If
    mProps.Contrato = value
    mProps.objPcProvCli.SetModificado
End Property

Public Property Get Contrato() As String
    Contrato = Trim$(mProps.Contrato)
End Property

Public Property Let Plazo(ByVal value As Integer)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "IVDetallePcProvCli_Plazo", MSGERR_INVALIDO
    End If
    mProps.Plazo = value
    mProps.objPcProvCli.SetModificado
End Property

Public Property Get Plazo() As Integer
    Plazo = Trim$(mProps.Plazo)
End Property


