VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VTKardexConciliaTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Private Type T_PROP
    id As Long
    IdVendedor As Long
    CodVendedor As String
    IdTrans As Long
    FechaGrabado As Date
    BandConciliado As Boolean
    Orden As Integer
    
    IdProveedor As Long
    CodProveedor As String
    
    IdEmpleado As Long
    CodEmpleado As String
    
    IdVehiculo As Long
    CodVehiculo As String
    
    
    objGNComprobante As gnComprobante           'Referencia al objeto padre
End Type

Private mProps As T_PROP


Public Sub VerificaDatos()
    'Verifica datos

'    If mProps.IdEmpleado = 0 And mProps.IdProveedor = 0 Then
'        Err.Raise ERR_INVALIDO, "VTKardexConciliaTrans_VerificaDatos", _
'            "Debe seleccionar un Empleado o un Proveedor."
'    End If
    
    If mProps.IdVehiculo = 0 Then
        Err.Raise ERR_INVALIDO, "VTKardexConciliaTrans_VerificaDatos", _
            "Debe seleccionar un Empleado o un Proveedor."
    End If

    
'    If mProps.IdTrans= 0 Then
'        Err.Raise ERR_INVALIDO, "VTKardexConciliaTrans_VerificaDatos", _
'            "Debe seleccionar un Detalle del documento de Bancos."
'    End If
End Sub

Friend Sub grabar()
    Dim sql As String, rs As Recordset

    VerificaDatos

    sql = "SELECT * FROM VTKardexConciliaTrans WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !transid = mProps.objGNComprobante.transid
        !IdVendedor = mProps.IdVendedor
        !IdTrans = mProps.IdTrans
        !FechaGrabado = mProps.FechaGrabado
        !BandConciliado = mProps.BandConciliado
        !IdEmpleado = mProps.IdEmpleado
        !IdProveedor = mProps.IdProveedor
        !IdVehiculo = mProps.IdVehiculo
        !Orden = mProps.Orden
        .Update

#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.id = !id             'Guarda el Id(Autonumerico)
        .Close
    End With

    Set rs = Nothing
End Sub

Friend Sub recuperar(rs As Recordset)
    With rs
        If Not .EOF Then
            mProps.id = !id
            If Not IsNull(!IdVendedor) Then mProps.IdVendedor = !IdVendedor
            If Not IsNull(!IdTrans) Then mProps.IdTrans = !IdTrans
            If Not IsNull(!FechaGrabado) Then mProps.FechaGrabado = !FechaGrabado
            If Not IsNull(!BandConciliado) Then mProps.BandConciliado = !BandConciliado
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            If Not IsNull(!IdEmpleado) Then mProps.IdEmpleado = !IdEmpleado
            If Not IsNull(!IdProveedor) Then mProps.IdProveedor = !IdProveedor
            If Not IsNull(!IdVehiculo) Then mProps.IdVehiculo = !IdVehiculo
        End If
    End With
End Sub

Friend Function Clone() As VTKardexConciliaTrans
    Dim nuevo As VTKardexConciliaTrans

    Set nuevo = New VTKardexConciliaTrans
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

Friend Property Get IdVendedor() As Long
    IdVendedor = mProps.IdVendedor
End Property


Public Property Let CodVendedor(ByVal Value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If Value <> mProps.CodVendedor Then
        If Len(Value) > 0 Then
            'Actualiza IdVendedor también
            sql = "SELECT IdVendedor FROM FcVendedor WHERE CodVendedor='" & Value & "'"
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
    mProps.CodVendedor = Value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodVendedor() As String
    CodVendedor = Trim$(mProps.CodVendedor)
End Property

Public Property Let IdTrans(ByVal Value As Long)
    mProps.IdTrans = Value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get IdTrans() As Long
    IdTrans = mProps.IdTrans
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






Friend Property Get IdEmpleado() As Long
    IdEmpleado = mProps.IdEmpleado
End Property


Public Property Let CodEmpleado(ByVal Value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If Value <> mProps.CodEmpleado Then
        If Len(Value) > 0 Then
            'Actualiza IdEmpleado también
            sql = "SELECT IdProvcli FROM Pcprovcli WHERE CodProvCli='" & Value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "VTKardex_CodEmpleado", MSGERR_NOHAYCODIGO
            Else
                mProps.IdEmpleado = rs!IdProvCli
            End If
            rs.Close
        Else
            mProps.IdEmpleado = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodEmpleado = Value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodEmpleado() As String
    CodEmpleado = Trim$(mProps.CodEmpleado)
End Property


Friend Property Get IdProveedor() As Long
    IdProveedor = mProps.IdProveedor
End Property


Public Property Let CodProveedor(ByVal Value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If Value <> mProps.CodProveedor Then
        If Len(Value) > 0 Then
            'Actualiza IdProveedor también
            sql = "SELECT IdProvcli FROM Pcprovcli WHERE CodProvCli='" & Value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "VTKardex_CodProveedor", MSGERR_NOHAYCODIGO
            Else
                mProps.IdProveedor = rs!IdProvCli
            End If
            rs.Close
        Else
            mProps.IdProveedor = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodProveedor = Value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodProveedor() As String
    CodProveedor = Trim$(mProps.CodProveedor)
End Property


Friend Property Get IdVehiculo() As Long
    IdVehiculo = mProps.IdVehiculo
End Property


Public Property Let CodVehiculo(ByVal Value As String)
    Dim sql As String, rs As Recordset

    'Cuando cambia el valor
    If Value <> mProps.CodVehiculo Then
        If Len(Value) > 0 Then
            'Actualiza IdVehiculo también
            sql = "SELECT Idvehiculo FROM gnvehiculo WHERE Codvehiculo='" & Value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "VTKardex_CodVehiculo", MSGERR_NOHAYCODIGO
            Else
                mProps.IdVehiculo = rs!IdVehiculo
            End If
            rs.Close
        Else
            mProps.IdVehiculo = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodVehiculo = Value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodVehiculo() As String
    CodVehiculo = Trim$(mProps.CodVehiculo)
End Property

