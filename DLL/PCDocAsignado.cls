VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCDocAsignado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROP
    Id As Long
    CodProvcli As String
    NombreProvCli As String
    transid As Long
    Trans As String
    Doc As String
    CodForma As String
    Valor As Currency
    VCancelado As Currency
    CodMoneda As String
    Cotizacion As Currency
    FechaEmision As Date
    Plazo As Long
    FechaVenci As Date
    
    '*** MAKOTO 12/ene/01 Agregados
    CodTrans As String
    NumTrans As Long
    NumDocRef As String
    Orden As Long
    
    
    ValorCapital1 As Currency
    ValorInteres1 As Currency
    TasaInteres1 As Currency
    
    ValorCapital2 As Currency
    ValorInteres2 As Currency
    TasaInteres2 As Currency
    
    ValorCapital3 As Currency
    ValorInteres3 As Currency
    TasaInteres3 As Currency
    
    
'    CodVendedor As String       'jeaa 27/05/2007
'    NombreVendedor As String
'    CodCobrador As String
'    NombreCobrador As String
    
    
    Guid As String          '*** MAKOTO 16/mar/01 Agregado
    
    objPCKardex As PCKardex
End Type

Private mProps As T_PROP


Friend Sub Recuperar(Id As Long)
    Dim sql As String, campo As String
    Dim rs As Recordset
  
  
'    sql = "SELECT * FROM vwConsPCPorCobrarPagar WHERE Id=" & Id

#If DAOLIB Then
    Dim qdef As QueryDef
    Set qdef = mProps.objPCKardex.GNComprobante.Empresa.Database.QueryDefs("spConsPCDocAsignado")
    qdef.Parameters("ParamId") = Id
    Set rs = qdef.OpenRecordset(dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    sql = "spConsPCDocAsignado " & Id
    rs.Open sql, mProps.objPCKardex.GNComprobante.Empresa.Coneccion, adOpenStatic, adLockReadOnly
    UltimoRecordset rs
#End If

    With rs
        If Not .EOF Then
            mProps.Id = !Id
            If Not IsNull(!CodProvcli) Then mProps.CodProvcli = !CodProvcli
            If Not IsNull(!Nombre) Then mProps.NombreProvCli = !Nombre
            If Not IsNull(!transid) Then mProps.transid = !transid
            If Not IsNull(!Trans) Then mProps.Trans = !Trans
            If Not IsNull(!Doc) Then mProps.Doc = !Doc
            If Not IsNull(!CodForma) Then mProps.CodForma = !CodForma
            If Not IsNull(!CodMoneda) Then mProps.CodMoneda = !CodMoneda
            Select Case mProps.objPCKardex.GNComprobante.Empresa.GNOpcion.IndiceMoneda(mProps.CodMoneda)
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
            If Not IsNull(!FechaVenci) Then mProps.FechaVenci = !FechaVenci
            
            '*** MAKOTO 12/ene/01 Agregados
            If Not IsNull(!CodTrans) Then mProps.CodTrans = !CodTrans
            If Not IsNull(!NumTrans) Then mProps.NumTrans = !NumTrans
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            
            '*** OLIVER 28/jun/2001 Agregado
            If Not IsNull(!NumDocRef) Then mProps.NumDocRef = !NumDocRef
            
            If Not IsNull(!Guid) Then mProps.Guid = !Guid   '*** MAKOTO 16/mar/01
            
'            If Not IsNull(!CodVendedor) Then mProps.CodVendedor = !CodVendedor
'            If Not IsNull(!NombreVendedor) Then mProps.NombreVendedor = !NombreVendedor
'            If Not IsNull(!CodCobrador) Then mProps.CodCobrador = !CodCobrador
'            If Not IsNull(!NombreCobrador) Then mProps.NombreCobrador = !NombreCobrador
            
            
            If Not IsNull(!ValorCapital1) Then mProps.ValorCapital1 = !ValorCapital1
            If Not IsNull(!ValorInteres1) Then mProps.ValorInteres1 = !ValorInteres1
            If Not IsNull(!TasaInteres1) Then mProps.TasaInteres1 = !TasaInteres1
            If Not IsNull(!ValorCapital2) Then mProps.ValorCapital2 = !ValorCapital2
            If Not IsNull(!ValorInteres2) Then mProps.ValorInteres2 = !ValorInteres2
            If Not IsNull(!TasaInteres2) Then mProps.TasaInteres2 = !TasaInteres2
            If Not IsNull(!ValorCapital3) Then mProps.ValorCapital3 = !ValorCapital3
            If Not IsNull(!ValorInteres3) Then mProps.ValorInteres3 = !ValorInteres3
            If Not IsNull(!TasaInteres3) Then mProps.TasaInteres3 = !TasaInteres3
            
            
        End If
        .Close
    End With
    Set rs = Nothing
End Sub

Friend Property Set PCKardex(value As PCKardex)
    Set mProps.objPCKardex = value
End Property

Public Property Get PCKardex() As PCKardex
    Set PCKardex = mProps.objPCKardex
End Property


Public Property Get Id() As Long
    Id = mProps.Id
End Property

Public Property Get CodProvcli() As String
    CodProvcli = Trim$(mProps.CodProvcli)
End Property

Public Property Get NombreProvCli() As String
    NombreProvCli = Trim$(mProps.NombreProvCli)
End Property

Public Property Get transid() As Long
    transid = mProps.transid
End Property

Public Property Get Trans() As String
    Trans = Trim$(mProps.Trans)
End Property

Public Property Get Doc() As String
    Doc = Trim$(mProps.Doc)
End Property

Public Property Get CodForma() As String
    CodForma = Trim$(mProps.CodForma)
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

Public Property Get Plazo() As Long
    Plazo = mProps.FechaVenci - mProps.FechaEmision
End Property

Public Property Get FechaVenci() As Date
    FechaVenci = mProps.FechaVenci
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
    Set mProps.objPCKardex = Nothing
End Sub







'Public Property Get CodVendedor() As String
'    CodVendedor = Trim$(mProps.CodVendedor)
'End Property
'
'Public Property Get NombreVendedor() As String
'    NombreVendedor = Trim$(mProps.NombreVendedor)
'End Property
'
'
'Public Property Get CodCobrador() As String
'    CodCobrador = Trim$(mProps.CodCobrador)
'End Property
'
'Public Property Get NombreCobrador() As String
'    NombreCobrador = Trim$(mProps.NombreCobrador)
'End Property
'

Friend Sub RecuperarT(Id As Long)
    Dim sql As String, campo As String
    Dim rs As Recordset
  
  
'    sql = "SELECT * FROM vwConsPCPorCobrarPagar WHERE Id=" & Id

#If DAOLIB Then
    Dim qdef As QueryDef
    Set qdef = mProps.objPCKardex.GNComprobante.Empresa.Database.QueryDefs("spConsPCDocAsignado")
    qdef.Parameters("ParamId") = Id
    Set rs = qdef.OpenRecordset(dbOpenSnapshot, dbReadOnly)
#Else
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    sql = "spConsPCDocAsignado " & Id
    rs.Open sql, mProps.objPCKardex.GNComprobante.Empresa.Coneccion, adOpenStatic, adLockReadOnly
    UltimoRecordset rs
#End If
    With rs
        If Not .EOF Then
            mProps.Id = !Id
            If Not IsNull(!CodProvcli) Then mProps.CodProvcli = !CodProvcli
            If Not IsNull(!Nombre) Then mProps.NombreProvCli = !Nombre
            If Not IsNull(!transid) Then mProps.transid = !transid
            If Not IsNull(!Trans) Then mProps.Trans = !Trans
            If Not IsNull(!Doc) Then mProps.Doc = !Doc
            If Not IsNull(!CodForma) Then mProps.CodForma = !CodForma
            If Not IsNull(!CodMoneda) Then mProps.CodMoneda = !CodMoneda
            Select Case mProps.objPCKardex.GNComprobante.Empresa.GNOpcion.IndiceMoneda(mProps.CodMoneda)
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
            If Not IsNull(!FechaVenci) Then mProps.FechaVenci = !FechaVenci
            
            '*** MAKOTO 12/ene/01 Agregados
            If Not IsNull(!CodTrans) Then mProps.CodTrans = !CodTrans
            If Not IsNull(!NumTrans) Then mProps.NumTrans = !NumTrans
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            
            '*** OLIVER 28/jun/2001 Agregado
            If Not IsNull(!NumDocRef) Then mProps.NumDocRef = !NumDocRef
            
            If Not IsNull(!Guid) Then mProps.Guid = !Guid   '*** MAKOTO 16/mar/01
            
'            If Not IsNull(!CodVendedor) Then mProps.CodVendedor = !CodVendedor
'            If Not IsNull(!NombreVendedor) Then mProps.NombreVendedor = !NombreVendedor
'            If Not IsNull(!CodCobrador) Then mProps.CodCobrador = !CodCobrador
'            If Not IsNull(!NombreCobrador) Then mProps.NombreCobrador = !NombreCobrador
            
        End If
        .Close
    End With
    Set rs = Nothing
End Sub

Public Property Let ValorCapital1(ByVal value As Currency)
    mProps.ValorCapital1 = value
End Property

Public Property Get ValorCapital1() As Currency
    ValorCapital1 = mProps.ValorCapital1
End Property

Public Property Let ValorCapital2(ByVal value As Currency)
    mProps.ValorCapital2 = value
End Property

Public Property Get ValorCapital2() As Currency
    ValorCapital2 = mProps.ValorCapital2
End Property

Public Property Let ValorCapital3(ByVal value As Currency)
    mProps.ValorCapital3 = value
End Property

Public Property Get ValorCapital3() As Currency
    ValorCapital3 = mProps.ValorCapital3
End Property

Public Property Let ValorInteres1(ByVal value As Currency)
    mProps.ValorInteres1 = value
End Property

Public Property Get ValorInteres1() As Currency
    ValorInteres1 = mProps.ValorInteres1
End Property

Public Property Let ValorInteres2(ByVal value As Currency)
    mProps.ValorInteres2 = value
End Property

Public Property Get ValorInteres2() As Currency
    ValorInteres2 = mProps.ValorInteres2
End Property

Public Property Let ValorInteres3(ByVal value As Currency)
    mProps.ValorInteres3 = value
End Property

Public Property Get ValorInteres3() As Currency
    ValorInteres3 = mProps.ValorInteres3
End Property
Public Function SaldoxTrans(ByVal IdAsignado As Long, ByVal CodMoneda As String) As Currency
Dim sql As String
Dim numReg As Long
Dim X As Integer
Dim rs As Recordset, rs1 As Recordset
        sql = " Select TOP 1 pck2.transid from pckardex pck Inner join pckardex pck2 on pck2.id = pck.idasignado Where pck.idasignado = " & IdAsignado
        Set rs = Me.PCKardex.GNComprobante.Empresa.OpenRecordset(sql)
        If rs.RecordCount > 0 Then
                X = Me.PCKardex.GNComprobante.Empresa.GNOpcion.IndiceMoneda(CodMoneda)
                sql = "SELECT  sum ( isnull(Valor" & X & ",0) - isnull(VCancelado" & X & ",0)  ) AS SaldoTrans "
                sql = sql & " From "
                sql = sql & " VWCONSPCPORCOBRARPAGAR "
                sql = sql & " Where Transid= " & rs!transid
                Set rs1 = Me.PCKardex.GNComprobante.Empresa.OpenRecordset(sql)
                If rs1.RecordCount > 0 Then
                    If Not IsNull(rs1!saldotrans) Then
                        SaldoxTrans = rs1!saldotrans
                    End If
                End If
                Set rs = Nothing
                Set rs1 = Nothing
        End If
End Function

