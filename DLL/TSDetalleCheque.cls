VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TSDetalleCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Private Type T_PROP
    Id As Long
    TransID As Long
    Valor As Currency
    Cantidad As Currency
    Orden As Integer
    Tipo As Integer
    IdBanco As Long
    CodBanco As String
    Nombre As String
    NumDoc As String
    FechaGrabado As Date
    IdTransRef As Long
    objGNComprobante As gnComprobante           'Referencia al objeto padre
End Type

Private mProps As T_PROP


Public Sub VerificaDatos()
    'Verifica datos
    If mProps.Valor = 0 And mProps.Cantidad = 0 Then
        Err.Raise ERR_INVALIDO, "TSDetalleCheque_VerificaDatos", _
            "Se intentó grabar sin valor de Denomicacion ni de Cantidad en detalle de Efectivo."
    End If
    
End Sub

Friend Sub Grabar()
    Dim sql As String, rs As Recordset
    
    VerificaDatos
   
    sql = "SELECT * FROM TSDetalleCheque WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !TransID = mProps.objGNComprobante.TransID
        !Valor = mProps.Valor
        !Cantidad = mProps.Cantidad
        !Orden = mProps.Orden
        !Tipo = mProps.Tipo
        !IdBanco = mProps.IdBanco
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
            If Not IsNull(!IdBanco) Then mProps.IdBanco = !IdBanco
            If Not IsNull(!CodBanco) Then mProps.CodBanco = !CodBanco
            If Not IsNull(!Nombre) Then mProps.Nombre = !Nombre
            If Not IsNull(!NumDoc) Then mProps.NumDoc = !NumDoc
            If Not IsNull(!IdTransRef) Then mProps.IdTransRef = !IdTransRef
        End If
    End With
End Sub

Friend Function Clone() As TSDetalleCheque
    Dim nuevo As TSDetalleCheque
    
    Set nuevo = New TSDetalleCheque
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROP)          'Para el método Clone
    mProps = value
    mProps.Id = 0
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Set gnComprobante(value As gnComprobante)
    Set mProps.objGNComprobante = value
End Property

Public Property Get gnComprobante() As gnComprobante
    Set gnComprobante = mProps.objGNComprobante
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
        Err.Raise ERR_INVALIDO, "TSDetalleCheque_Valor", MSGERR_INVALIDO
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
        Err.Raise ERR_INVALIDO, "TSDetalleCheque_Cantidad", MSGERR_INVALIDO
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

Public Property Let IdBanco(ByVal value As Long)
    mProps.IdBanco = value
    mProps.objGNComprobante.SetModificado
End Property


Friend Property Get IdBanco() As Long
    IdBanco = mProps.IdBanco
End Property


Public Property Let CodBanco(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodBanco Then
        If Len(value) > 0 Then
            'Actualiza IdBanco también
            sql = "SELECT IdBanco, dESCRIPCION FROM IVBanco WHERE CodBanco='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "TSKardex_CodBanco", MSGERR_NOHAYCODIGO
            Else
                mProps.IdBanco = rs!IdBanco
            End If
            rs.Close
        Else
            mProps.IdBanco = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodBanco = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodBanco() As String
    CodBanco = Trim$(mProps.CodBanco)
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




