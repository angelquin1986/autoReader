VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TSDetalleTarjeta"
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
    Cantidad As Currency
    Orden As Integer
    Tipo As Integer
    IdTarjeta As Long
    CodTarjeta As String
    Nombre As String
    NumDoc As String
    FechaGrabado As Date
    objGNComprobante As GNComprobante           'Referencia al objeto padre
End Type

Private mProps As T_PROP


Public Sub VerificaDatos()
    'Verifica datos
    If mProps.Valor = 0 And mProps.Cantidad = 0 Then
        Err.Raise ERR_INVALIDO, "TSDetalleTarjeta_VerificaDatos", _
            "Se intentó grabar sin valor de Denomicacion ni de Cantidad en detalle de Efectivo."
    End If
    
End Sub

Friend Sub Grabar()
    Dim sql As String, rs As Recordset
    
    VerificaDatos
   
    sql = "SELECT * FROM TSDetalleTarjeta WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !transid = mProps.objGNComprobante.transid
        !Valor = mProps.Valor
        !Cantidad = mProps.Cantidad
        !Orden = mProps.Orden
        !Tipo = mProps.Tipo
        !IdTarjeta = mProps.IdTarjeta
        !Nombre = mProps.Nombre
        !NumDoc = mProps.NumDoc
        !FechaGrabado = mProps.objGNComprobante.FechaTrans
        !IdTransRef = mProps.IdTransRef
        
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
            If Not IsNull(!Cantidad) Then mProps.Cantidad = !Cantidad
            If Not IsNull(!FechaGrabado) Then mProps.FechaGrabado = !FechaGrabado
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            If Not IsNull(!Tipo) Then mProps.Tipo = !Tipo
            If Not IsNull(!IdTarjeta) Then mProps.IdTarjeta = !IdTarjeta
            If Not IsNull(!CodTarjeta) Then mProps.CodTarjeta = !CodTarjeta
            If Not IsNull(!Nombre) Then mProps.Nombre = !Nombre
            If Not IsNull(!NumDoc) Then mProps.NumDoc = !NumDoc
            If Not IsNull(!IdTransRef) Then mProps.IdTransRef = !IdTransRef
        End If
    End With
End Sub

Friend Function Clone() As TSDetalleTarjeta
    Dim nuevo As TSDetalleTarjeta
    
    Set nuevo = New TSDetalleTarjeta
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

Public Property Let Tipo(ByVal value As Integer)
    mProps.Tipo = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Tipo() As Integer
    Tipo = mProps.Tipo
End Property

Private Sub Class_Terminate()
    Set mProps.objGNComprobante = Nothing
End Sub



Public Property Let Valor(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "TSDetalleTarjeta_Valor", MSGERR_INVALIDO
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

Public Property Let Cantidad(ByVal value As Currency)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
    
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "TSDetalleTarjeta_Cantidad", MSGERR_INVALIDO
    ElseIf value > 0 Then
        mProps.Cantidad = value
    Else
        mProps.Cantidad = 0
    End If
    
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Cantidad() As Currency
    'Devuelve en Moneda ex.
    Cantidad = mProps.Cantidad / mProps.objGNComprobante.Cotizacion("")
End Property

Public Property Let IdTarjeta(ByVal value As Long)
    mProps.IdTarjeta = value
    mProps.objGNComprobante.SetModificado
End Property


Friend Property Get IdTarjeta() As Long
    IdTarjeta = mProps.IdTarjeta
End Property


Public Property Let CodTarjeta(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodTarjeta Then
        If Len(value) > 0 Then
            'Actualiza IdTarjeta también
            sql = "SELECT IdTarjeta, dESCRIPCION FROM IVTarjeta WHERE CodTarjeta='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "TSKardex_CodTarjeta", MSGERR_NOHAYCODIGO
            Else
                mProps.IdTarjeta = rs!IdTarjeta
            End If
            rs.Close
        Else
            mProps.IdTarjeta = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodTarjeta = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodTarjeta() As String
    CodTarjeta = Trim$(mProps.CodTarjeta)
End Property

'----
Public Property Let Nombre(ByVal value As String)
    mProps.Nombre = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Nombre() As String
    Nombre = mProps.Nombre
End Property

Public Property Let NumDoc(ByVal value As String)
    mProps.NumDoc = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get NumDoc() As String
    NumDoc = mProps.NumDoc
End Property






Public Property Let IdTransRef(ByVal value As Long)
    'Convierte en sucres
    value = value * mProps.objGNComprobante.Cotizacion("")
   
    If value < 0 Then
        Err.Raise ERR_INVALIDO, "TSDetalleEgreso_IdTransRef", MSGERR_INVALIDO
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


