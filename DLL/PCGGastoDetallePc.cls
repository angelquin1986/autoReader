VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PCGGastoDetallePc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type T_PROPS
    IdGastoDetallePC As Long
    IdPCGGasto As Long
    NumPCGrupo  As Byte
    idPcGrupo  As Integer
    CodPCGrupo As String      'revisar  si nos sirve
    IdProvCli  As Long
    CodProvCli As String      'revisar  si nos sirve
    idGasto As Long
    CodGasto As String
    Nombre As String
    
    Orden As Long
    objPCGGasto As PCGGasto
End Type

Private mProps As T_PROPS

'Establece relacion  con el PADRE
Friend Property Set PCGGasto(objPCGGasto As PCGGasto)
    Set mProps.objPCGGasto = objPCGGasto
End Property

Public Property Get PCGGasto() As PCGGasto
    Set PCGGasto = mProps.objPCGGasto
End Property

Friend Sub Recuperar(rs As Recordset)
    Dim sql As String, rs1  As Recordset
    With rs
        If Not .EOF Then
            mProps.IdPCGGasto = !IdPCGGasto
            mProps.IdGastoDetallePC = !IdGastoDetallePC
           ' If Not IsNull(!NumPCGrupo) Then mProps.NumPCGrupo = !NumPCGrupo
            If Not IsNull(!idPcGrupo) Then mProps.idPcGrupo = !idPcGrupo
            If Not IsNull(!IdProvCli) Then mProps.IdProvCli = !IdProvCli
            If Not IsNull(!idGasto) Then mProps.idGasto = !idGasto

            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            'recupera  el codigo  de PCGrupo
            On Error Resume Next
            If mProps.NumPCGrupo > 0 And mProps.NumPCGrupo < PCGRUPO_MAX Then  'Abre todo
                sql = "Select CodGrupo" & mProps.NumPCGrupo & " From PCGrupo" & mProps.NumPCGrupo & _
                      " Where IDGrupo" & mProps.NumPCGrupo & " = " & mProps.idPcGrupo
                Set rs1 = mProps.objPCGGasto.Empresa.OpenRecordset(sql)
                If Not rs1.EOF Then
                    mProps.CodPCGrupo = rs1.Fields(0).value
                End If
                rs1.Close
                Set rs1 = Nothing
            Else
                sql = "Select Codprovcli,nombre From PCProvcli"
                sql = sql & " Where IDprovcli = " & mProps.IdProvCli
                Set rs1 = mProps.objPCGGasto.Empresa.OpenRecordset(sql)
                If Not rs1.EOF Then
                    mProps.CodProvCli = rs1.Fields(0).value
                    mProps.Nombre = rs1.Fields(1).value
                End If
                rs1.Close
                Set rs1 = Nothing
                'recupera codgasto
                sql = "Select CodGasto From gngasto"
                sql = sql & " Where idgasto = " & mProps.idGasto
                Set rs1 = mProps.objPCGGasto.Empresa.OpenRecordset(sql)
                If Not rs1.EOF Then
                    mProps.CodGasto = rs1.Fields(0).value
                    
                End If
                rs1.Close
                Set rs1 = Nothing
            
            End If
        
        End If
    End With
End Sub

Public Function Clone() As PCGGastoDetallePC
    Dim nuevo As PCGGastoDetallePC
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New PCGGastoDetallePC
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    'mProps.IdDescuento = 0
    'mbooModificado = True
End Property


Public Property Get IdPCGGasto() As Long
    IdPCGGasto = mProps.IdPCGGasto
End Property

Public Property Let IdPCGGasto(ByVal value As Long)
    mProps.IdPCGGasto = value
End Property

Public Property Get Orden() As Long
    Orden = mProps.Orden
End Property

Public Property Let Orden(ByVal value As Long)
    mProps.Orden = value
End Property

Public Property Let NumPCGrupo(ByVal value As Byte)
    mProps.NumPCGrupo = value
    mProps.objPCGGasto.Modificado = True
End Property

Public Property Get NumPCGrupo() As Byte
    NumPCGrupo = mProps.NumPCGrupo
End Property

Public Property Get CodPCGrupo() As String
    CodPCGrupo = mProps.CodPCGrupo
End Property

Public Property Let CodPCGrupo(value As String)
    Dim rs As Recordset, sql As String
    mProps.CodPCGrupo = value
    If value = "" Then
        mProps.idPcGrupo = 0
    Else
        'actualiza  el IDGrupo
        sql = "Select IdGrupo" & mProps.NumPCGrupo & " From PCGrupo" & mProps.NumPCGrupo & _
              " Where CodGrupo" & mProps.NumPCGrupo & " = '" & value & "'"
        Set rs = mProps.objPCGGasto.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then
            mProps.idPcGrupo = rs.Fields(0).value
        End If
        rs.Close
        Set rs = Nothing
    End If
    mProps.objPCGGasto.Modificado = True
End Property


Public Property Get idPcGrupo() As Long
    idPcGrupo = mProps.idPcGrupo
End Property



Private Sub Class_Terminate()
    Set mProps.objPCGGasto = Nothing
End Sub

Public Property Get CodProvCli() As String
    CodProvCli = mProps.CodProvCli
End Property

Public Property Let CodProvCli(value As String)
    Dim rs As Recordset, sql As String
    mProps.CodProvCli = value
    If value = "" Then
        mProps.IdProvCli = 0
    Else
        'actualiza  el IDGrupo
        sql = "Select IdProvCli From PcProvCli"
        sql = sql & " Where CodProvCli = '" & value & "'"
        Set rs = mProps.objPCGGasto.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then
            mProps.IdProvCli = rs.Fields(0).value
        End If
        rs.Close
        Set rs = Nothing
    End If
    mProps.objPCGGasto.Modificado = True
End Property

Public Property Get IdProvCli() As Long
    IdProvCli = mProps.IdProvCli
End Property




Public Property Let Nombre(ByVal value As String)
    If Len(value) > 80 Then
        Err.Raise ERR_INVALIDO, "IVDescuentoDetallePC_Nombre", MSGERR_INVALIDO
    End If
    mProps.Nombre = value
    'mbooModificado = True
End Property

Public Property Get Nombre() As String
    Nombre = Trim$(mProps.Nombre)
End Property




'------------------
Public Property Get CodGasto() As String
    CodGasto = mProps.CodGasto
End Property

Public Property Let CodGasto(value As String)
    Dim rs As Recordset, sql As String
    mProps.CodGasto = value
    If value = "" Then
        mProps.idGasto = 0
    Else
        'actualiza  el IDGrupo
        sql = "Select IdGasto From gnGasto"
        sql = sql & " Where codGasto = '" & value & "'"
        Set rs = mProps.objPCGGasto.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then
            mProps.idGasto = rs.Fields(0).value
        End If
        rs.Close
        Set rs = Nothing
    End If
    mProps.objPCGGasto.Modificado = True
End Property

Public Property Get idGasto() As Long
    idGasto = mProps.idGasto
End Property

Public Property Let idPcGrupo(ByVal value As Long)
    mProps.idPcGrupo = value
End Property

