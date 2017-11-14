VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GNKardexDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROP
    Id As Long
    transid As Long
    IdProvcli As Long
    CodEmpleado As String
    Cantidad As Integer
    Saldo As Currency
    Orden As Integer
    BandSelec As Boolean
    TransidDoc As Long
    OrdenVisita As Long
    IdProvCliOri As Long
    CodEmpleadoOri As String
    IdRuta As Long
    Codruta As String
    IdRutaOri As Long
    CodrutaOri As String
    objGNComprobante As GNComprobante           'Referencia al objeto padre
End Type

Private mProps As T_PROP

Public Sub VerificaDatos()
    Dim cant As Currency
    Dim item As AFInventario
    Dim pc As PCProvCli, nunPCGrupo As Integer, cadPrecios As String, idPcGrupo As Long
    Dim nprecio As Integer
    
'       La cantidad no puede ser 0.
    If mProps.Cantidad = 0 Then
       Err.Raise ERR_INVALIDO, "GNKardexDocumento_VerificaDatos", _
           "La cantidad no puede ser 0. (" & mProps.Cantidad & ")"
    End If
    
    If mProps.IdRuta = 0 Then
        Err.Raise ERR_INVALIDO, "GNKardexDocumento_VerificaDatos", _
            "Debe seleccionar una Ruta. (" & mProps.CodEmpleado & ")"
            
    End If
    If mProps.objGNComprobante.GNTrans.CodPantalla <> "VTCONCITCO" Then
        If mProps.IdProvcli = 0 Then
            Err.Raise ERR_INVALIDO, "GNKardexDocumento_VerificaDatos", _
            "Debe seleccionar un Empleado. (" & mProps.CodEmpleado & ")"
        End If
    End If

    If mProps.TransidDoc = 0 Then
        Err.Raise ERR_INVALIDO, "GNKardexDocumento_VerificaDatos", _
            "Debe seleccionar un Documento."
    End If
End Sub



''''Private Function SumaItemRepetido() As Long
''''    Dim i As Long, ivk As GNKardexDocumento, cantOtras As Currency
''''
''''    For i = 1 To mProps.objGNComprobante.CountGNKardexDocumento
''''        Set ivk = mProps.objGNComprobante.GNKardexDocumento(i)
''''        If (Not (ivk Is Me)) And (ivk.CodInventario = Me.CodInventario) Then
''''            cantOtras = cantOtras + ivk.Cantidad
''''        End If
''''    Next i
''''    Set ivk = Nothing
''''
''''    SumaItemRepetido = cantOtras
''''End Function

Friend Sub Grabar()
    Dim sql As String, rs As Recordset
    Dim item As AFInventario
    Dim pc As PCProvCli
    Dim SecuencialItem As Long
    VerificaDatos
    If mProps.objGNComprobante.GNTrans.CodPantalla = "VTCONCITCO" Then
        AseguraBodegaRuta
    Else
        AseguraBodega
    End If
    sql = "SELECT * FROM GNKardexDocumento WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !transid = mProps.objGNComprobante.transid
        !IdProvcli = mProps.IdProvcli
        !Cantidad = mProps.Cantidad
        !Orden = mProps.Orden
        !TransidDoc = mProps.TransidDoc
        !Saldo = mProps.Saldo
        !OrdenVisita = mProps.OrdenVisita
        !IdRuta = mProps.IdRuta
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.Id = !Id             'Guarda el Id(Autonumerico)
        .Close
    End With
    
    Set item = Nothing
    Set rs = Nothing
    Set pc = Nothing
    Set rs = Nothing
End Sub

Friend Sub Recuperar(rs As Recordset)
    With rs
        If Not .EOF Then
            mProps.Id = !Id

            If Not IsNull(!IdProvcli) Then mProps.IdProvcli = !IdProvcli
            If Not IsNull(!CodEmpleado) Then mProps.CodEmpleado = !CodEmpleado
            If Not IsNull(!Cantidad) Then mProps.Cantidad = !Cantidad
            If Not IsNull(!TransidDoc) Then mProps.TransidDoc = !TransidDoc
            If Not IsNull(!Saldo) Then mProps.Saldo = !Saldo
            If Not IsNull(!OrdenVisita) Then mProps.OrdenVisita = !OrdenVisita
            If Not IsNull(!IdRuta) Then mProps.IdRuta = !IdRuta
            If Not IsNull(!Codruta) Then mProps.Codruta = !Codruta
        End If
    End With
End Sub

Friend Function Clone() As GNKardexDocumento
    Dim nuevo As GNKardexDocumento

    Set nuevo = New GNKardexDocumento
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



Public Property Let TransidDoc(ByVal value As Long)
    mProps.TransidDoc = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get TransidDoc() As Long
    TransidDoc = mProps.TransidDoc
End Property



Public Property Let IdProvcli(ByVal value As Long)
    mProps.IdProvcli = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get IdProvcli() As Long
    IdProvcli = mProps.IdProvcli
End Property

Public Property Let CodEmpleado(ByVal value As String)
    Dim sql As String, rs As Recordset

  '  Cuando cambia el valor
    If value <> mProps.CodEmpleado Then
        If Len(value) > 0 Then
   '         Actualiza IdProvCli también
            sql = "SELECT IdProvCli FROM Empleado WHERE CodProvCli='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNKardexDocumento_CodEmpleado", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                Me.IdProvcli = rs!IdProvcli
            End If
            rs.Close
        Else
            Me.IdProvcli = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodEmpleado = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodEmpleado() As String
    CodEmpleado = Trim$(mProps.CodEmpleado)
End Property

Public Property Let Cantidad(ByVal value As Currency)
    mProps.Cantidad = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Cantidad() As Currency
    Cantidad = mProps.Cantidad
End Property



Public Property Let Orden(ByVal value As Integer)
    mProps.Orden = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Orden() As Integer
    Orden = mProps.Orden
End Property



Private Sub Class_Initialize()
    mProps.Cantidad = 1
    mProps.BandSelec = 0
    mProps.Saldo = 0
End Sub

Private Sub Class_Terminate()
    Set mProps.objGNComprobante = Nothing
End Sub


Public Property Let BandSelec(ByVal value As Boolean)
    mProps.BandSelec = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get BandSelec() As Boolean
    BandSelec = mProps.BandSelec
End Property


Private Sub AseguraBodega()
    Dim sql As String, rs As Recordset
    
    sql = "SELECT IdprovCli FROM GNExistDocumento WHERE transid=" & _
                mProps.TransidDoc & " AND IdProvCli=" & mProps.IdProvcli
    
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
    'Si no existe el registro correspondiente, lo inserta
    If rs.EOF Then
        sql = "INSERT INTO GNExistDocumento (transid,Idprovcli,Exist,orden) " & _
              "VALUES(" & mProps.TransidDoc & "," & mProps.IdProvcli & "," & 0 & "," & mProps.OrdenVisita & ")"
        mProps.objGNComprobante.Empresa.Execute sql, True
    End If
    rs.Close
    Set rs = Nothing
End Sub


Public Property Let Saldo(ByVal value As Currency)
    mProps.Saldo = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Saldo() As Currency
    Saldo = mProps.Saldo
End Property

Public Property Let OrdenVisita(ByVal value As Integer)
    mProps.OrdenVisita = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get OrdenVisita() As Integer
    OrdenVisita = mProps.OrdenVisita
End Property

Public Property Let CodEmpleadoOri(ByVal value As String)
    Dim sql As String, rs As Recordset
  '  Cuando cambia el valor
    If value <> mProps.CodEmpleadoOri Then
        If Len(value) > 0 Then
   '         Actualiza IdProvCli también
            sql = "SELECT IdProvCli FROM Empleado WHERE CodProvCli='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNKardexDocumento_CodEmpleadoOri", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                Me.IdProvCliOri = rs!IdProvcli
            End If
            rs.Close
        Else
            Me.IdProvCliOri = 0
        End If
    End If
    Set rs = Nothing
    mProps.CodEmpleadoOri = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodEmpleadoOri() As String
    CodEmpleadoOri = Trim$(mProps.CodEmpleadoOri)
End Property
Public Property Let IdProvCliOri(ByVal value As Long)
    mProps.IdProvCliOri = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get IdProvCliOri() As Long
    IdProvCliOri = mProps.IdProvCliOri
End Property


Public Property Let Codruta(ByVal value As String)
    Dim sql As String, rs As Recordset
  '  Cuando cambia el valor
    If value <> mProps.Codruta Then
        If Len(value) > 0 Then
   '         Actualiza IdProvCli también
            sql = "SELECT Id FROM GNRuta WHERE CodRuta='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNKardexDocumento_CodRuta", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                Me.IdRuta = rs!Id
            End If
            rs.Close
        Else
            Me.IdRuta = 0
        End If
    End If
    Set rs = Nothing
    mProps.Codruta = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Codruta() As String
    Codruta = Trim$(mProps.Codruta)
End Property
Public Property Let IdRuta(ByVal value As Long)
    mProps.IdRuta = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get IdRuta() As Long
    IdRuta = mProps.IdRuta
End Property

Public Property Let CodrutaOri(ByVal value As String)
    Dim sql As String, rs As Recordset
  '  Cuando cambia el valor
    If value <> mProps.CodrutaOri Then
        If Len(value) > 0 Then
   '         Actualiza IdProvCli también
            sql = "SELECT Id FROM GNRuta WHERE CodRuta='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "GNKardexDocumento_CodRuta", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                Me.IdRutaOri = rs!Id
            End If
            rs.Close
        Else
            Me.IdRutaOri = 0
        End If
    End If
    Set rs = Nothing
    mProps.CodrutaOri = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodrutaOri() As String
    CodrutaOri = Trim$(mProps.CodrutaOri)
End Property
Public Property Let IdRutaOri(ByVal value As Long)
    mProps.IdRutaOri = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get IdRutaOri() As Long
    IdRutaOri = mProps.IdRutaOri
End Property

Private Sub AseguraBodegaRuta()
    Dim sql As String, rs As Recordset
    
    sql = "SELECT IdRuta FROM GNExistDocumento WHERE transid=" & _
                mProps.TransidDoc & " AND Idruta=" & mProps.IdRuta
    
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
    'Si no existe el registro correspondiente, lo inserta
    If rs.EOF Then
        sql = "INSERT INTO GNExistDocumento (transid,Idprovcli,Exist,orden,IdRuta) " & _
              "VALUES(" & mProps.TransidDoc & "," & mProps.IdRuta & "," & 0 & "," & mProps.OrdenVisita & "," & mProps.IdRuta & ")"
        mProps.objGNComprobante.Empresa.Execute sql, True
    End If
    rs.Close
    Set rs = Nothing
End Sub

