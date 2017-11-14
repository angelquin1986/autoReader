VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CcDocAsignado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROP
    Id As Long
    CodCentro As String
    Descripcion As String
    TransID As Long
    Trans As String
    Valor As Currency
    VCancelado As Currency
    CodMoneda As String
    Cotizacion As Currency
    FechaEmision As Date
    
    '*** MAKOTO 12/ene/01 Agregados
    CodTrans As String
    NumTrans As Long
    NumDocRef As String
    Orden As Long
    
    Guid As String          '*** MAKOTO 16/mar/01 Agregado
    
    objCcKardex As CcKardex
End Type

Private mProps As T_PROP


Friend Sub Recuperar(Id As Long)
    Dim sql As String, campo As String
    Dim rs As Recordset
  
  
'    sql = "SELECT * FROM vwConsPCPorCobrarPagar WHERE Id=" & Id

#If DAOLIB Then
    Dim qdef As QueryDef
    Set qdef = mProps.objCcKardex.GNComprobante.Empresa.Database.QueryDefs("spConsPCDocAsignado")
    qdef.Parameters("ParamId") = Id
    Set rs = qdef.OpenRecordset(dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    sql = "spConsPCDocAsignado " & Id
    rs.Open sql, mProps.objCcKardex.GNComprobante.Empresa.Coneccion, adOpenStatic, adLockReadOnly
    UltimoRecordset rs
#End If
    With rs
        If Not .EOF Then
            mProps.Id = !Id
            If Not IsNull(!CodCentro) Then mProps.CodCentro = !CodCentro
            If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
            If Not IsNull(!TransID) Then mProps.TransID = !TransID
            If Not IsNull(!Trans) Then mProps.Trans = !Trans
            If Not IsNull(!CodMoneda) Then mProps.CodMoneda = !CodMoneda
            Select Case mProps.objCcKardex.GNComprobante.Empresa.GNOpcion.IndiceMoneda(mProps.CodMoneda)
            Case 1
                mProps.Cotizacion = 1
                If Not IsNull(!Valor1) Then mProps.Valor = !Valor1
                If Not IsNull(!VCancelado1) Then mProps.VCancelado = !VCancelado1
            Case 2
                If Not IsNull(!Cotizacion2) Then mProps.Cotizacion = !Cotizacion2
                If Not IsNull(!Valor2) Then mProps.Valor = !Valor2
                If Not IsNull(!VCancelado2) Then mProps.VCancelado = !VCancelado2
            Case 3
                If Not IsNull(!Cotizacion3) Then mProps.Cotizacion = !Cotizacion3
                If Not IsNull(!Valor3) Then mProps.Valor = !Valor3
                If Not IsNull(!VCancelado3) Then mProps.VCancelado = !VCancelado3
            Case 4
                If Not IsNull(!Cotizacion4) Then mProps.Cotizacion = !Cotizacion4
                If Not IsNull(!Valor4) Then mProps.Valor = !Valor4
                If Not IsNull(!VCancelado4) Then mProps.VCancelado = !VCancelado4
            End Select
            
            If Not IsNull(!FechaEmision) Then mProps.FechaEmision = !FechaEmision
            
            '*** MAKOTO 12/ene/01 Agregados
            If Not IsNull(!CodTrans) Then mProps.CodTrans = !CodTrans
            If Not IsNull(!NumTrans) Then mProps.NumTrans = !NumTrans
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            
            '*** OLIVER 28/jun/2001 Agregado
            If Not IsNull(!NumDocRef) Then mProps.NumDocRef = !NumDocRef
            
            If Not IsNull(!Guid) Then mProps.Guid = !Guid   '*** MAKOTO 16/mar/01
        End If
        .Close
    End With
    Set rs = Nothing
End Sub

Friend Property Set CcKardex(value As CcKardex)
    Set mProps.objCcKardex = value
End Property

Public Property Get CcKardex() As CcKardex
    Set CcKardex = mProps.objCcKardex
End Property


Public Property Get Id() As Long
    Id = mProps.Id
End Property

Public Property Get CodCentro() As String
    CodCentro = Trim$(mProps.CodCentro)
End Property

Public Property Get Descripcion() As String
    Descripcion = Trim$(mProps.Descripcion)
End Property

Public Property Get TransID() As Long
    TransID = mProps.TransID
End Property

Public Property Get Trans() As String
    Trans = Trim$(mProps.Trans)
End Property

Public Property Get Valor() As Currency
    Valor = mProps.Valor
End Property

Public Property Get VCancelado() As Currency
    VCancelado = mProps.VCancelado
End Property

Public Property Get Saldo() As Currency
    Saldo = Valor - VCancelado
End Property


Public Property Get CodMoneda() As String
    CodMoneda = Trim$(mProps.CodMoneda)
End Property

Public Property Get Cotizacion() As Currency
    Cotizacion = mProps.Cotizacion
End Property


Public Property Get FechaEmision() As Date
    FechaEmision = mProps.FechaEmision
End Property


'*** MAKOTO 12/ene/01 Agregado
Public Property Get CodTrans() As String
    CodTrans = mProps.CodTrans
End Property

'*** MAKOTO 12/ene/01 Agregado
Public Property Get NumTrans() As Long
    NumTrans = mProps.NumTrans
End Property

'*** MAKOTO 12/ene/01 Agregado
Public Property Get Orden() As Long
    Orden = mProps.Orden
End Property

'*** MAKOTO 16/mar/01 Agregado
Public Property Get Guid() As String
    Guid = mProps.Guid
End Property
'*** Oliver 28/jun/2001 Agregado
Public Property Get NumDocRef() As String
    NumDocRef = Trim$(mProps.NumDocRef)
End Property

''*** Oliver 28/jun/2001 Agregado
'Public Property Let NumDocRef(ByVal Value As String)
'    mProps.NumDocRef = Value
'End Property




Private Sub Class_Initialize()
    mProps.Cotizacion = 1
End Sub

Private Sub Class_Terminate()
    Set mProps.objCcKardex = Nothing
End Sub








