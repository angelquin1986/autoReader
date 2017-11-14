VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AFKardexCustodio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROP
    Id As Long
    IdInventario As Long
    CodInventario As String
    IdProvCli As Long
    CodEmpleado As String
    Cantidad As Currency

    Nota As String
    Orden As Integer
    Nota1 As String

    FechaEntrega As Date
    FechaDevol As Date
    objGNComprobante As GNComprobante           'Referencia al objeto padre
End Type


Private mProps As T_PROP

Public Sub VerificaDatos()
    Dim cant As Currency
    Dim item As AFInventario
    Dim nunPCGrupo As Integer, cadPrecios As String, idPcGrupo As Long
    Dim nprecio As Integer


   Set item = mProps.objGNComprobante.Empresa.RecuperaAFInventario(mProps.CodInventario)


'       La cantidad no puede ser 0.
    If mProps.Cantidad = 0 Then
       Err.Raise ERR_INVALIDO, "AFKardexCustodio_VerificaDatos", _
           "La cantidad no puede ser 0. (" & mProps.CodInventario & ")"
    End If
    

    If mProps.IdProvCli = 0 Then
        Err.Raise ERR_INVALIDO, "AFKardexCustodio_VerificaDatos", _
            "Debe seleccionar un Empleado. (" & mProps.CodInventario & ")"
    End If

    If mProps.IdInventario = 0 Then
        Err.Raise ERR_INVALIDO, "AFKardexCustodio_VerificaDatos", _
            "Debe seleccionar un Activo."
    End If


End Sub



Private Function SumaItemRepetido() As Long
    Dim i As Long, ivk As AFKardexCustodio, cantOtras As Currency

    For i = 1 To mProps.objGNComprobante.CountAFKardexCustodio
        Set ivk = mProps.objGNComprobante.AFKardexCustodio(i)
        If (Not (ivk Is Me)) And (ivk.CodInventario = Me.CodInventario) Then
            cantOtras = cantOtras + ivk.Cantidad
        End If
    Next i
    Set ivk = Nothing

    SumaItemRepetido = cantOtras
End Function

Friend Sub Grabar()
    Dim sql As String, rs As Recordset
    Dim item As AFInventario
    
    Dim SecuencialItem As Long
    VerificaDatos
    AseguraBodega
    sql = "SELECT * FROM AFKardexCustodio WHERE 1=0"
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordsetParaEdit(sql)
    With rs
        .AddNew
        !transid = mProps.objGNComprobante.transid
        !IdInventario = mProps.IdInventario
        !IdProvCli = mProps.IdProvCli
        !Cantidad = mProps.Cantidad
        !Orden = mProps.Orden
        !Nota = mProps.Nota
        !Nota1 = mProps.Nota1
'''        !FechaDevol = mProps.FechaDevol   'AUC para demostracion de los andamios
'''        !FechaEntrega = mProps.FechaEntrega    'AUC para demostracion de los andamios
        .Update
#If DAOLIB Then
        rs.Bookmark = rs.LastModified
#End If
        .Move 0             'Para actualizar
        mProps.Id = !Id             'Guarda el Id(Autonumerico)
        .Close
        
       
    End With
    
    ' ACTUALIZA EN LA TABLA CUSTODIO
    
    If mProps.Cantidad > 0 Then
        sql = "UPDATE AFInventario set idempleado= " & mProps.IdProvCli & "where idinventario=" & mProps.IdInventario
        Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
    End If
    Set item = Nothing
    Set rs = Nothing

    Set rs = Nothing
End Sub

Friend Sub Recuperar(rs As Recordset)
    With rs
        If Not .EOF Then
            mProps.Id = !Id
            If Not IsNull(!IdInventario) Then mProps.IdInventario = !IdInventario
            If Not IsNull(!IdProvCli) Then mProps.IdProvCli = !IdProvCli
            If Not IsNull(!CodInventario) Then mProps.CodInventario = !CodInventario
            If Not IsNull(!CodEmpleado) Then mProps.CodEmpleado = !CodEmpleado
            If Not IsNull(!Cantidad) Then mProps.Cantidad = !Cantidad
            If Not IsNull(!Nota) Then mProps.Nota = !Nota
            If Not IsNull(!Nota1) Then mProps.Nota1 = !Nota1
'            If Not IsNull(!FechaDevol) Then mProps.FechaDevol = !FechaDevol      ' *** AUC para demostracion de los andamios
'            If Not IsNull(!FechaEntrega) Then mProps.FechaEntrega = !FechaEntrega

        End If
    End With
End Sub

Friend Function Clone() As AFKardexCustodio
    Dim nuevo As AFKardexCustodio

    Set nuevo = New AFKardexCustodio
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


Public Property Let CodInventario(ByVal value As String)
    Dim sql As String, rs As Recordset

'    Cuando cambia el valor
    If value <> mProps.CodInventario Then
        If Len(value) > 0 Then
 '           Actualiza IdInventario también
            sql = "SELECT IdInventario " & _
                  "FROM AFInventario WHERE CodInventario='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "AFKardexCustodio_CodInventario", MSGERR_NOHAYCODIGO
            Else
                mProps.IdInventario = rs!IdInventario
            End If
            rs.Close
        Else
            mProps.IdInventario = 0
        End If
    End If

    Set rs = Nothing
    mProps.CodInventario = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get CodInventario() As String
    CodInventario = Trim$(mProps.CodInventario)
End Property

Friend Property Let IdProvCli(ByVal value As Long)
    mProps.IdProvCli = value
    mProps.objGNComprobante.SetModificado
End Property

Friend Property Get IdProvCli() As Long
    IdProvCli = mProps.IdProvCli
End Property

Public Property Let CodEmpleado(ByVal value As String)
    Dim sql As String, rs As Recordset

  '  Cuando cambia el valor
    If value <> mProps.CodEmpleado Then
        If Len(value) > 0 Then
   '         Actualiza IdProvCli también
            sql = "SELECT IdProvCli FROM EMPLEADO WHERE CodProvCli='" & value & "'"
            Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "AFKardexCustodio_CodEmpleado", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                Me.IdProvCli = rs!IdProvCli
            End If
            rs.Close
        Else
            Me.IdProvCli = 0
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

Public Property Let Nota(ByVal value As String)
    value = RTrim$(value)
    If Len(value) > 80 Then                 '*** MAKOTO 16/oct/00 Modifcado 40-->80
        Err.Raise ERR_INVALIDO, "AFKardexCustodio_Nota", MSGERR_INVALIDO
    End If

    mProps.Nota = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Nota() As String
    Nota = RTrim$(mProps.Nota)
End Property


Private Sub Class_Terminate()
    Set mProps.objGNComprobante = Nothing
End Sub



Public Property Let IdInventario(ByVal value As Long)
    mProps.IdInventario = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Let FechaDevol(value As Date)
    If mProps.FechaDevol <> value Then
        mProps.FechaDevol = value
    End If
End Property

Public Property Get FechaDevol() As Date
    FechaDevol = mProps.FechaDevol
End Property

Public Property Let FechaEntrega(value As Date)
    If mProps.FechaEntrega <> value Then
        mProps.FechaEntrega = value
    End If
End Property

Public Property Get FechaEntrega() As Date
    FechaEntrega = mProps.FechaEntrega
End Property

Public Property Let Nota1(ByVal value As String)
    value = RTrim$(value)
    If Len(value) > 80 Then                 '*** MAKOTO 16/oct/00 Modifcado 40-->80
        Err.Raise ERR_INVALIDO, "AFKardexCustodio_Nota1", MSGERR_INVALIDO
    End If

    mProps.Nota1 = value
    mProps.objGNComprobante.SetModificado
End Property

Public Property Get Nota1() As String
    Nota1 = RTrim$(mProps.Nota1)
End Property


Private Sub AseguraBodega()
    Dim sql As String, rs As Recordset
    
    sql = "SELECT Idprovcli FROM afExistCustodio WHERE IdInventario=" & _
                mProps.IdInventario & " AND IdProvcli=" & mProps.IdProvCli
    
    Set rs = mProps.objGNComprobante.Empresa.OpenRecordset(sql)
    'Si no existe el registro correspondiente, lo inserta
    If rs.EOF Then
        sql = "INSERT INTO AFExistCustodio (IdInventario,IdProvcli,Exist) " & _
              "VALUES(" & mProps.IdInventario & "," & mProps.IdProvCli & "," & 0 & ")"
        mProps.objGNComprobante.Empresa.Execute sql, True
    End If
    rs.Close
    Set rs = Nothing
End Sub


