VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VTKardexConcilia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Private Type T_PROP
    Id As Long
    IdVendedor As Long
    CodVendedor As String
    IdTrans As Long
    FechaGrabado As Date
    BandConciliado As Boolean
    Orden As Integer
    objGNComprobante As gnComprobante           'Referencia al objeto padre
End Type

Private mProps As T_PROP


Public Sub VerificaDatos()
    'Verifica datos

    If mProps.IdVendedor = 0 Then
        Err.Raise ERR_INVALIDO, "VTKardexConcilia_VerificaDatos", _
            "Debe seleccionar un banco."
    End If
'    If mProps.IdTrans= 0 Then
'        Err.Raise ERR_INVALIDO, "VTKardexConcilia_VerificaDatos", _
'            "Debe seleccionar un Detalle del documento de Bancos."
'    End If
End Sub

Friend Sub Grabar()
    Dim sql As String, rs As Recordset

    VerificaDatos

    sql = "SELECT * FROM VTKardexConcilia WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !TransID = mProps.objGNComprobante.TransID
        !IdVendedor = mProps.IdVendedor
        !IdTrans = mProps.IdTrans
        !FechaGrabado = mProps.FechaGrabado
        !BandConciliado = mProps.BandConciliado
        !Orden = mProps.Orden
        .Update

#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.Id = !Id             'Guarda el Id(Autonumerico)
        .Close
    End With

    'If mProps.IdTrans= 507 Then MsgBox "Hola"
'    If mProps.BandConciliado Then
'        sql = "update VTKardex set BandConciliado=" & IIf(mProps.BandConciliado, 1, 0)
'        sql = sql & " where id=" & mProps.IdTrans
'        Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
'    End If
    'cambia estado en VTKardex
    Set rs = Nothing
End Sub

Friend Sub Recuperar(rs As Recordset)
    With rs
        If Not .EOF Then
            mProps.Id = !Id
            If Not IsNull(!IdVendedor) Then mProps.IdVendedor = !IdVendedor
            If Not IsNull(!IdTrans) Then mProps.IdTrans = !IdTrans
            If Not IsNull(!FechaGrabado) Then mProps.FechaGrabado = !FechaGrabado
            If Not IsNull(!BandConciliado) Then mProps.BandConciliado = !BandConciliado
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
        End If
    End With
End Sub

Friend Function Clone() As VTKardexConcilia
    Dim nuevo As VTKardexConcilia

    Set nuevo = New VTKardexConcilia
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

Friend Property Get IdVendedor() As Long
    IdVendedor = mProps.IdVendedor
End Property


Public Property Let CodVendedor(ByVal value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If value <> mProps.CodVendedor Then
        If Len(value) > 0 Then
            'Actualiza IdVendedor también
            sql = "SELECT IdVendedor FROM FcVendedor WHERE CodVendedor='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "VTKardex_CodVendedor", MSGERR_NOHAYCODIGO
            Else
                mProps.IdVendedor = rs!IdVendedor
            End If
            rs.Close
        Else
            mProps.IdVendedor = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodVendedor = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodVendedor() As String
    CodVendedor = Trim$(mProps.CodVendedor)
End Property

Public Property Let IdTrans(ByVal value As Long)
    mProps.IdTrans = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get IdTrans() As Long
    IdTrans = mProps.IdTrans
End Property



Public Property Let FechaGrabado(ByVal value As Date)
    mProps.FechaGrabado = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get FechaGrabado() As Date
    FechaGrabado = mProps.FechaGrabado
End Property

Public Property Let BandConciliado(ByVal value As Boolean)
    mProps.BandConciliado = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get BandConciliado() As Boolean
    BandConciliado = mProps.BandConciliado
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


'*** oliver 29/may/2001


Public Function Verificar_ExisteDocumento(CodVendedor As String, CodTipoDoc As String, NumDoc As String) As Boolean
    Dim rs As Recordset, sql As String, numReg As Long
    On Error GoTo ErrTrap
    sql = "SELECT Count(TransID) as Cuantos FROM " & _
           "(VTKardex INNER JOIN TSBanco ON VTKardex.IdVendedor = TSBanco.IdVendedor) " & _
           "INNER JOIN TSTipoDocBanco ON VTKardex.IdTrans= TSTipoDocBanco.IdTrans" & _
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




