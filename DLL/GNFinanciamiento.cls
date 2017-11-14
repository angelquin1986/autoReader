VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GNFinanciamiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROP
    Id As Long
    TasaMensual As Currency
    ValorEntrada As Currency
    MesesGracia As Integer
    DiaPago As Integer
    NumeroPagos As Integer
    FechaPrimerPago As Date
    Capital As Currency
    Interes As Currency
    NuevaTasa As Currency
    Mora As Currency
    Descuento As Currency
    Adicional1 As Currency
    FechaAdicional1 As Date
    Adicional2 As Currency
    FechaAdicional2 As Date
    Adicional3 As Currency
    FechaAdicional3 As Date
    TipoCuota As Integer ' si son cuotas separadas o unificadas vehiculo/seguro/dispositivo
    TipoRef As Integer  'para saber si es refinanciamiento o pronto pago
    objGNComprobante As GNComprobante           'Referencia al objeto padre
End Type
Private mProps As T_PROP

Friend Sub Grabar()
    Dim sql As String, rs As Recordset
    sql = "SELECT * FROM GnFinanciamiento WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !TransID = mProps.objGNComprobante.TransID
        !TasaMensual = mProps.TasaMensual
        !MesesGracia = mProps.MesesGracia
        !ValorEntrada = mProps.ValorEntrada
        !FechaPrimerPago = mProps.FechaPrimerPago
        !DiaPago = mProps.DiaPago
        !NumeroPagos = mProps.NumeroPagos
        !Capital = mProps.Capital
        !Interes = mProps.Interes
        !NuevaTasa = mProps.NuevaTasa
        !Mora = mProps.Mora
        !Descuento = mProps.Descuento
        !Adicional1 = mProps.Adicional1
        !FechaAdicional1 = mProps.FechaAdicional1
        !Adicional2 = mProps.Adicional2
        !FechaAdicional2 = mProps.FechaAdicional2
        !Adicional3 = mProps.Adicional3
        !TipoCuota = mProps.TipoCuota
        !TipoRef = mProps.TipoRef
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.Id = !Id             'Guarda el Id(Autonumerico)
        .Close
    End With
End Sub

Friend Sub Recuperar(rs As Recordset)
    With rs
        If Not .EOF Then
            mProps.Id = !Id
            If Not IsNull(!TasaMensual) Then mProps.TasaMensual = !TasaMensual
            If Not IsNull(!MesesGracia) Then mProps.MesesGracia = !MesesGracia
            If Not IsNull(!DiaPago) Then mProps.DiaPago = !DiaPago
            If Not IsNull(!NumeroPagos) Then mProps.NumeroPagos = !NumeroPagos
            If Not IsNull(!FechaPrimerPago) Then mProps.FechaPrimerPago = !FechaPrimerPago
            If Not IsNull(!Capital) Then mProps.Capital = !Capital
            If Not IsNull(!Interes) Then mProps.Interes = !Interes
            If Not IsNull(!NuevaTasa) Then mProps.NuevaTasa = !NuevaTasa
            If Not IsNull(!Mora) Then mProps.Mora = !Mora
            If Not IsNull(!Descuento) Then mProps.Descuento = !Descuento
            If Not IsNull(!Adicional1) Then mProps.Adicional1 = !Adicional1
            If Not IsNull(!FechaAdicional1) Then mProps.FechaAdicional1 = !FechaAdicional1
            If Not IsNull(!Adicional2) Then mProps.Adicional2 = !Adicional2
            If Not IsNull(!FechaAdicional2) Then mProps.FechaAdicional2 = !FechaAdicional2
            If Not IsNull(!Adicional3) Then mProps.Adicional3 = !Adicional3
            If Not IsNull(!FechaAdicional3) Then mProps.FechaAdicional3 = !FechaAdicional3
            If Not IsNull(!TipoRef) Then mProps.TipoRef = !TipoRef
            If Not IsNull(!TipoCuota) Then mProps.TipoCuota = !TipoCuota
        End If
    End With
End Sub

Friend Function Clone() As GNFinanciamiento
    Dim nuevo As GNFinanciamiento
    
    Set nuevo = New GNFinanciamiento
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

'AUC cambiado para el reproceso de recetas 19/10/06
Public Property Get Id() As Long
    Id = mProps.Id
End Property


Public Property Let Capital(ByVal value As Currency)
    mProps.Capital = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Capital() As Currency
    Capital = mProps.Capital
End Property
Public Property Let Interes(ByVal value As Currency)
    mProps.Interes = value
        mProps.objGNComprobante.SetModificado
End Property

Public Property Get Interes() As Currency
    Interes = mProps.Interes
End Property
Public Property Let NuevaTasa(ByVal value As Currency)
    mProps.NuevaTasa = value
        mProps.objGNComprobante.SetModificado
End Property

Public Property Get NuevaTasa() As Currency
    NuevaTasa = mProps.NuevaTasa
End Property

Public Property Let Mora(ByVal value As Currency)
    mProps.Mora = value
        mProps.objGNComprobante.SetModificado
End Property

Public Property Get Mora() As Currency
   Mora = mProps.Mora
End Property

Public Property Let Descuento(ByVal value As Currency)
    mProps.Descuento = value
        mProps.objGNComprobante.SetModificado
End Property

Public Property Get Descuento() As Currency
   Descuento = mProps.Descuento
End Property


Public Property Let Adicional1(ByVal value As Currency)
    mProps.Adicional1 = value
        mProps.objGNComprobante.SetModificado
End Property

Public Property Get Adicional1() As Currency
   Adicional1 = mProps.Adicional1
End Property

Public Property Let FechaAdicional1(value As Date)
    
    If mProps.FechaAdicional1 <> value Then
        mProps.FechaAdicional1 = value
            mProps.objGNComprobante.SetModificado
    End If
End Property

Public Property Get FechaAdicional1() As Date
    FechaAdicional1 = mProps.FechaAdicional1
End Property
Public Property Let Adicional2(ByVal value As Currency)
    mProps.Adicional2 = value
        mProps.objGNComprobante.SetModificado
End Property

Public Property Get Adicional2() As Currency
   Adicional2 = mProps.Adicional2
End Property

Public Property Let FechaAdicional2(value As Date)
    
    
    If mProps.FechaAdicional2 <> value Then
        mProps.FechaAdicional2 = value
            mProps.objGNComprobante.SetModificado
    End If
End Property
Public Property Get FechaAdicional2() As Date
    FechaAdicional2 = mProps.FechaAdicional2
End Property
Public Property Let FechaAdicional3(value As Date)
    If mProps.FechaAdicional3 <> value Then
        mProps.FechaAdicional3 = value
            mProps.objGNComprobante.SetModificado
    End If
End Property
Public Property Get FechaAdicional3() As Date
    FechaAdicional3 = mProps.FechaAdicional3
End Property
Public Property Let NumeroPagos(ByVal value As Integer)
    mProps.NumeroPagos = value
        mProps.objGNComprobante.SetModificado
End Property
Public Property Get NumeroPagos() As Integer
    NumeroPagos = mProps.NumeroPagos
End Property
Public Property Let FechaPrimerPago(ByVal value As Date)
    mProps.FechaPrimerPago = value
        mProps.objGNComprobante.SetModificado
End Property
Public Property Get FechaPrimerPago() As Date
    FechaPrimerPago = mProps.FechaPrimerPago
End Property
Public Property Let TasaMensual(ByVal value As Currency)
    mProps.TasaMensual = value
        mProps.objGNComprobante.SetModificado
End Property
Public Property Get TasaMensual() As Currency
    TasaMensual = mProps.TasaMensual
End Property
Public Property Let Adicional3(ByVal value As Currency)
    mProps.Adicional3 = value
        mProps.objGNComprobante.SetModificado
End Property
Public Property Get Adicional3() As Currency
   Adicional3 = mProps.Adicional3
End Property
Public Property Let TipoCuota(ByVal value As Integer)
        mProps.TipoCuota = value
        mProps.objGNComprobante.SetModificado
End Property
Public Property Get TipoCuota() As Integer
   TipoCuota = mProps.TipoCuota
End Property
Public Property Let TipoRef(ByVal value As Integer)
        mProps.TipoRef = value
        mProps.objGNComprobante.SetModificado
End Property
Public Property Get TipoRef() As Integer
   TipoRef = mProps.TipoRef
End Property

