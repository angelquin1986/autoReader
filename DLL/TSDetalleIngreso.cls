VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TSDetalleIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Private Type T_PROP
    Id As Long
    transid As Long
    IdTransRef As Long
    Valor As Currency
    Orden As Integer
    FechaGrabado As Date
    objGNComprobante As GNComprobante           'Referencia al objeto padre
End Type

Private mProps As T_PROP


Public Sub VerificaDatos()
    'Verifica datos
    If mProps.Valor = 0 And mProps.IdTransRef = 0 Then
        Err.Raise ERR_INVALIDO, "TSDetalleIngreso_VerificaDatos", _
            "Se intentó grabar sin valor de Denomicacion ni de IdTransRef en detalle de Efectivo."
    End If
    
End Sub

Friend Sub Grabar()
    Dim sql As String, rs As Recordset
    
    VerificaDatos
   
    sql = "SELECT * FROM TSDetalleIngreso WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !transid = mProps.objGNComprobante.transid
        !Valor = mProps.Valor
        !IdTransRef = mProps.IdTransRef
        !Orden = mProps.Orden
        !FechaGrabado = mProps.objGNComprobante.FechaTrans
        .Update
        
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.Id = !Id             'Guarda el Id(Autonumerico)
        .Close
    End With

    Set rs = Nothing
End Sub

Friend Sub Recuperar(rs As Recordset)
    With rs
        If Not .EOF Then
            mProps.Id = !Id
            If Not IsNull(!Valor) Then mProps.Valor = !Valor
            If Not IsNull(!IdTransRef) Then mProps.IdTransRef = !IdTransRef
            If Not IsNull(!FechaGrabado) Then mProps.FechaGrabado = !FechaGrabado
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            If Not IsNull(!IdTransRef) Then mProps.IdTransRef = !IdTransRef
        End If
    End With
End Sub

Friend Function Clone() As TSDetalleIngreso
    Dim nuevo As TSDetalleIngreso
    
    Set nuevo = New TSDetalleIngreso
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROP)          'Para el método Clone
    mProps = value
    mProps.Id = 0
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Set GNComprobante(value As GNComprobante)
    Set mProps.objGNComprobante = value
End Property

Public Property Get GNComprobante() As GNComprobante
    Set GNComprobante = mProps.objGNComprobante
End Property

Public Property Get Id() As Long
    Id = mProps.Id
End Property

Public Property Let FechaGrabado(ByVal value As Date)
    mProps.FechaGrabado = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get FechaGrabado() As Date
    FechaGrabado = mProps.FechaGrabado
End Property

Public Property Let Orden(ByVal value As Integer)
    mProps.Orden = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property

Private Sub Class_Terminate()
    Set mProps.objGNComprobante = Nothing
End Sub

Public Property Let Valor(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "TSDetalleIngreso_Valor", MSGERR_INVALIDO
    ElseIf value > 0 Then
        mProps.Valor = value
    Else
        mProps.Valor = 0
    End If
    
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Valor() As Currency
    'Devuelve en Moneda ex.
    Valor = mProps.Valor / mProps.objGNComprobante.Cotizacion("")
End Property

Public Property Let IdTransRef(ByVal value As Long)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
   
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "TSDetalleIngreso_IdTransRef", MSGERR_INVALIDO
    ElseIf value > 0 Then
        mProps.IdTransRef = value
    Else
        mProps.IdTransRef = 0
    End If
    
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get IdTransRef() As Long
    'Devuelve en Moneda ex.
    IdTransRef = mProps.IdTransRef / mProps.objGNComprobante.Cotizacion("")
End Property



