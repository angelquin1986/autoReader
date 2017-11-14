VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PRDetalleGasto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'---- OJO
'--------- jeaaa
'-------- no se utila para nada ESTA CLASE
Option Explicit

Private Type T_PROP
    Id As Long
    TransID As Long
    IdCuenta As Long
    CodCuenta As String
    Mes As Date
    Valor As Currency
    BandUtilizado As Byte
    idTransUtiliza As Long
    Orden As Integer
    FechaGrabado As Date
    objPRLibroDetalle As PRLibroDetalle 'Referencia al objeto padre

End Type

Private mbooEsNuevo As Boolean
Private mbooModificado As Boolean

Private mProps As T_PROP

Friend Property Set PRLibroDetalle(obj As PRLibroDetalle)
    Set mProps.objPRLibroDetalle = obj
End Property

Public Property Get PRLibroDetalle() As PRLibroDetalle
    Set PRLibroDetalle = mProps.objPRLibroDetalle
End Property

Public Property Let IdCuenta(ByVal value As Long)
    mProps.IdCuenta = value
End Property

Public Property Get IdCuenta() As Long
    IdCuenta = mProps.IdCuenta
End Property

Public Property Let CodCuenta(ByVal value As String)
    If Len(value) > 20 Then
        Err.Raise ERR_INVALIDO, "PRLibroDetalle_CodCuenta", MSGERR_INVALIDO
    End If
    mProps.CodCuenta = value
End Property

Public Property Get CodCuenta() As String
    CodCuenta = Trim$(mProps.CodCuenta)
End Property



Public Sub VerificaDatos()
    'Verifica datos
    If mProps.Valor = 0 And mProps.Mes = "0" Then
        Err.Raise ERR_INVALIDO, "PRDetalleGasto_VerificaDatos", _
            "Se intentó grabar sin valor de Gasto o mes de Gasto."
    End If
    
End Sub

Friend Sub Grabar()
''''    Dim sql As String, rs As Recordset
''''
''''    VerificaDatos
''''
''''    sql = "SELECT * FROM PRDetalleGasto WHERE 1=0"
''''    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
''''    With rs
''''        .AddNew
''''        !TransID = mProps.objGNComprobante.TransID
''''        !Valor = mProps.Valor
''''        !Mes = mProps.Mes
''''        !Orden = mProps.Orden
''''        !FechaGrabado = mProps.objGNComprobante.FechaTrans
''''        !IdCuenta = mProps.IdCuenta
''''        !BandUtilizado = mProps.BandUtilizado
''''        !idTransUtilza = mProps.idTransUtiliza
''''        !IdPRLibroDetalleRef = 0
''''        .Update
''''
''''#If DAOLIB Then
''''        rs.Bookmark = rs.LastModified
''''#End If
''''        .Move 0             'Para actualizar
''''        mProps.Id = !Id             'Guarda el Id(Autonumerico)
''''        .Close
''''    End With
''''
''''    Set rs = Nothing
End Sub

Friend Sub Recuperar(rs As Recordset)
'''    With rs
'''        If Not .EOF Then
'''            mProps.Id = !Id
'''            If Not IsNull(!Valor) Then mProps.Valor = !Valor
'''            If Not IsNull(!Cantidad) Then mProps.Mes = !Mes
'''            If Not IsNull(!FechaGrabado) Then mProps.FechaGrabado = !FechaGrabado
'''            If Not IsNull(!Orden) Then mProps.Orden = !Orden
'''            If Not IsNull(!IdCuenta) Then !IdCuenta = mProps.IdCuenta
'''            If Not IsNull(!BandUtilizado) Then !BandUtilizado = mProps.BandUtilizado
'''            If Not IsNull(!idTransUtiliza) Then !idTransUtiliza = mProps.idTransUtiliza
'''
'''
'''        End If
'''    End With
End Sub

Friend Function Clone() As PRDetalleGasto
'''    Dim nuevo As PRDetalleGasto
'''
'''    Set nuevo = New PRDetalleGasto
'''    nuevo.Props = mProps
'''    Set Clone = nuevo
'''    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROP)          'Para el método Clone
'    mProps = value
'    mProps.Id = 0

End Property

Friend Property Set gnComprobante(value As gnComprobante)
'    Set mProps.objGNComprobante = value
End Property

Public Property Get gnComprobante() As gnComprobante
'    Set gnComprobante = mProps.objGNComprobante
End Property


Public Property Get Id() As Long
    Id = mProps.Id
End Property

Public Property Let FechaGrabado(ByVal value As Date)
    mProps.FechaGrabado = value
End Property

Public Property Get FechaGrabado() As Date
    FechaGrabado = mProps.FechaGrabado
End Property


Public Property Let Orden(ByVal value As Integer)
    mProps.Orden = value
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property

Public Property Let Valor(ByVal value As Currency)
    'Convierte en sucres
'    value = value * mProps.objGNComprobante.Cotizacion("")
    
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "PRDetalleGasto_Valor", MSGERR_INVALIDO
    ElseIf value > 0 Then
        mProps.Valor = value
    Else
        mProps.Valor = 0
    End If
    

End Property

Public Property Get Valor() As Currency
    'Devuelve en Moneda ex.
    Valor = mProps.Valor '/ mProps.objGNComprobante.Cotizacion("")
End Property

Public Property Let Mes(ByVal value As Date)
    mProps.Mes = value

End Property

Public Property Get Mes() As Date
    Mes = mProps.Mes
End Property

Public Property Let BandUtilizado(ByVal value As Byte)
    mProps.BandUtilizado = value

End Property

Public Property Get BandUtilizado() As Byte
    BandUtilizado = mProps.BandUtilizado
End Property

Public Property Let idTransUtiliza(ByVal value As Long)
    mProps.idTransUtiliza = value

End Property

Public Property Get idTransUtiliza() As Long
    idTransUtiliza = mProps.idTransUtiliza
End Property

