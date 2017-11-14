VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TSKardexConcilia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Private Type T_PROP
    id As Long
    IdBanco As Long
    CodBanco As String
    IdTsKardex As Long
    FechaGrabado As Date
    BandConciliado As Boolean
    Orden As Integer
    objGNComprobante As GNComprobante           'Referencia al objeto padre
End Type

Private mProps As T_PROP


Public Sub VerificaDatos()
    'Verifica datos
    If mProps.IdBanco = 0 Then
        Err.Raise ERR_INVALIDO, "TSKardex_VerificaDatos", _
            "Debe seleccionar un banco."
    End If
    If mProps.IdTsKardex = 0 Then
        Err.Raise ERR_INVALIDO, "TSKardex_VerificaDatos", _
            "Debe seleccionar un tipo de documento en detalle de Bancos."
    End If
End Sub

Friend Sub Grabar()
    Dim sql As String, rs As Recordset
    
    VerificaDatos
   
    sql = "SELECT * FROM TSConciliaBanco WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !TransID = mProps.objGNComprobante.TransID
        !IdBanco = mProps.IdBanco
        !IdTsKardex = mProps.IdTsKardex
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
    Set rs = Nothing
End Sub

Friend Sub Recuperar(rs As Recordset)
    With rs
        If Not .EOF Then
            mProps.id = !id
            If Not IsNull(!IdBanco) Then mProps.IdBanco = !IdBanco
            If Not IsNull(!CodBanco) Then mProps.CodBanco = !CodBanco
            If Not IsNull(!IdTsKardex) Then mProps.IdTsKardex = !IdTsKardex
            If Not IsNull(!FechaGrabado) Then mProps.FechaGrabado = !FechaGrabado
            If Not IsNull(!BandConciliado) Then mProps.BandConciliado = !BandConciliado
            If Not IsNull(!Orden) Then mProps.Orden = !Orden
        End If
    End With
End Sub

Friend Function Clone() As TSConciliaBanco
    Dim nuevo As TSConciliaBanco
    
    Set nuevo = New TSConciliaBanco
    nuevo.Props = mProps
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROP)          'Para el método Clone
    mProps = value
    mProps.id = 0
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Set GNComprobante(value As GNComprobante)
    Set mProps.objGNComprobante = value
End Property

Public Property Get GNComprobante() As GNComprobante
    Set GNComprobante = mProps.objGNComprobante
End Property



Public Property Get id() As Long
    id = mProps.id
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
            sql = "SELECT IdBanco,IdCuentaContable FROM TSBanco WHERE CodBanco='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "TSKardex_CodBanco", MSGERR_NOHAYCODIGO
            Else
                mProps.IdBanco = rs!IdBanco
                mProps.auxIdCuenta = rs!IdCuentaContable
            End If
            rs.Close
        Else
            mProps.IdBanco = 0
            mProps.auxIdCuenta = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodBanco = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodBanco() As String
    CodBanco = Trim$(mProps.CodBanco)
End Property

Friend Property Let IdTsKardex(ByVal value As Long)
    mProps.IdTsKardex = value
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Get IdTsKardex() As Long
    IdTsKardex = mProps.IdTsKardex
End Property



Public Property Let FechaGrabado(ByVal value As Date)
    mProps.FechaGrabado = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get FechaGrabado() As Date
    FechaGrabado = mProps.FechaGrabado
End Property

Public Property Let BandConciliado(ByVal value As Date)
    mProps.BandConciliado = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get BandConciliado() As Date
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


Public Function Verificar_ExisteDocumento(CodBanco As String, CodTipoDoc As String, NumDoc As String) As Boolean
    Dim rs As Recordset, sql As String, numReg As Long
    On Error GoTo ErrTrap
    sql = "SELECT Count(TransID) as Cuantos FROM " & _
           "(TSKardex INNER JOIN TSBanco ON TSKardex.IdBanco = TSBanco.IdBanco) " & _
           "INNER JOIN TSTipoDocBanco ON TSKardex.IdTsKardex = TSTipoDocBanco.IdTsKardex " & _
           "WHERE TSBanco.CodBanco = '" & CodBanco & "' AND " & _
           "TSTipoDocBanco.CodTipoDoc = '" & CodTipoDoc & "' AND " & _
           "TSKARDEX.NumDoc='" & NumDoc & "'"
    'Debug.Print sql
    Set rs = Me.GNComprobante.Empresa.OpenRecordset(sql)
        Verificar_ExisteDocumento = IIf(rs!Cuantos = 0, False, True)
    Exit Function
ErrTrap:
        MsgBox Err.Number & " " & Err.Description, vbInformation
        Verificar_ExisteDocumento = False
End Function

