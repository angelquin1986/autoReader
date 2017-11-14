VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVKardexConciliaCarcasa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Private Type T_PROP
    id As Long
    IdTrans As Long
    Ticket As Long
    FechaGrabado As Date
    BandConciliado As Boolean
    Orden As Integer
    objGNComprobante As gnComprobante           'Referencia al objeto padre
End Type

Private mProps As T_PROP


Public Sub VerificaDatos()
    'Verifica datos

End Sub

Friend Sub grabar()
    Dim sql As String, rs As Recordset

    VerificaDatos

    sql = "SELECT * FROM IVKardexConciliaCarcasa WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !transid = mProps.objGNComprobante.transid
        !Ticket = mProps.Ticket
        !FechaGrabado = mProps.FechaGrabado
        !BandConciliado = mProps.BandConciliado
        !Orden = mProps.Orden
        .Update

#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.id = !id             'Guarda el Id(Autonumerico)
        .Close
    End With

    'If mProps.IdTrans= 507 Then MsgBox "Hola"
'    If mProps.BandConciliado Then
'        sql = "update VTKardex set BandConciliado=" & IIf(mProps.BandConciliado, 1, 0)
'        sql = sql & " where id=" & mProps.Ticket
'        Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
'    End If
    'cambia estado en VTKardex
    Set rs = Nothing
End Sub

Friend Sub recuperar(rs As Recordset)
    With rs
        If Not .EOF Then
            mProps.id = !id
            If Not IsNull(!Ticket) Then mProps.Ticket = !Ticket
            If Not IsNull(!FechaGrabado) Then mProps.FechaGrabado = !FechaGrabado
            If Not IsNull(!BandConciliado) Then mProps.BandConciliado = !BandConciliado
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
        End If
    End With
End Sub

Friend Function Clone() As IVKardexConciliaCarcasa
    Dim nuevo As IVKardexConciliaCarcasa

    Set nuevo = New IVKardexConciliaCarcasa
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(Value As T_PROP)          'Para el método Clone
    mProps = Value
    mProps.id = 0
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Set gnComprobante(Value As gnComprobante)
    Set mProps.objGNComprobante = Value
End Property

Public Property Get gnComprobante() As gnComprobante
    Set gnComprobante = mProps.objGNComprobante
End Property



Public Property Get id() As Long
    id = mProps.id
End Property


Public Property Let Ticket(ByVal Value As Long)
    mProps.Ticket = Value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Ticket() As Long
    Ticket = mProps.Ticket
End Property



Public Property Let FechaGrabado(ByVal Value As Date)
    mProps.FechaGrabado = Value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get FechaGrabado() As Date
    FechaGrabado = mProps.FechaGrabado
End Property

Public Property Let BandConciliado(ByVal Value As Boolean)
    mProps.BandConciliado = Value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get BandConciliado() As Boolean
    BandConciliado = mProps.BandConciliado
End Property


Public Property Let Orden(ByVal Value As Integer)
    mProps.Orden = Value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property

Private Sub Class_Terminate()
    Set mProps.objGNComprobante = Nothing
End Sub


'*** oliver 29/may/2001


Public Function Verificar_ExisteDocumento(CodVendedor As String, CodTipoDoc As String, NumDoc As String) As Boolean
    Dim rs As Recordset, sql As String, numReg As Long
    On Error GoTo ErrTrap
    sql = "SELECT Count(TransID) as Cuantos FROM " & _
           "(VTKardex INNER JOIN TSBanco ON VTKardex.IdVendedor = TSBanco.IdVendedor) " & _
           "INNER JOIN TSTipoDocBanco ON VTKardex.IdTrans= TSTipoDocBanco.Ticket" & _
           "WHERE TSBanco.CodVendedor = '" & CodVendedor & "' AND " & _
           "TSTipoDocBanco.CodTipoDoc = '" & CodTipoDoc & "' AND " & _
           "VTKardex.NumDoc='" & NumDoc & "'"
    'Debug.Print sql
    Set rs = Me.gnComprobante.Empresa.OpenRecordset(sql)
        Verificar_ExisteDocumento = IIf(rs!Cuantos = 0, False, True)
    Exit Function
ErrTrap:
        MsgBox Err.Number & " " & Err.Description, vbInformation
        Verificar_ExisteDocumento = False
End Function






