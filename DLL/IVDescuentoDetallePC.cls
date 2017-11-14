VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVDescuentoDetallePC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***Diego 15/09/2003
Private Type T_PROPS
    IdDescuento As Long
    IdDescuentoDetallePC As Long
    NumPCGrupo  As Byte
    idPcGrupo  As Integer
    CodPCGrupo As String      'revisar  si nos sirve
    IdProvcli  As Long
    CodProvcli As String      'revisar  si nos sirve
    Nombre As String
    
    Orden As Long
    objIVDescuento As IVDescuento   'Referencia al Objeto Padre
End Type

Private mProps As T_PROPS
'Private mbooEsNuevo As Boolean
'Private mbooModificado As Boolean


'Establece relacion  con el PADRE
Friend Property Set IVDescuento(objIVDescuento As IVDescuento)
    Set mProps.objIVDescuento = objIVDescuento
End Property

Public Property Get IVDescuento() As IVDescuento
    Set IVDescuento = mProps.objIVDescuento
End Property

Friend Sub Recuperar(rs As Recordset)
    Dim sql As String, rs1  As Recordset
    With rs
        If Not .EOF Then
            mProps.IdDescuento = !IdDescuento
            mProps.IdDescuentoDetallePC = !IdDescuentoDetallePC
            If Not IsNull(!NumPCGrupo) Then mProps.NumPCGrupo = !NumPCGrupo
            If Not IsNull(!idPcGrupo) Then mProps.idPcGrupo = !idPcGrupo
            If Not IsNull(!IdProvcli) Then mProps.IdProvcli = !IdProvcli

            If Not IsNull(!Orden) Then mProps.Orden = !Orden
            'recupera  el codigo  de PCGrupo
            On Error Resume Next
            If mProps.NumPCGrupo > 0 And mProps.NumPCGrupo < PCGRUPO_MAX Then  'Abre todo
                sql = "Select CodGrupo" & mProps.NumPCGrupo & " From PCGrupo" & mProps.NumPCGrupo & _
                      " Where IDGrupo" & mProps.NumPCGrupo & " = " & mProps.idPcGrupo
                Set rs1 = mProps.objIVDescuento.Empresa.OpenRecordset(sql)
                If Not rs1.EOF Then
                    mProps.CodPCGrupo = rs1.Fields(0).value
                End If
                rs1.Close
                Set rs1 = Nothing
            Else
                sql = "Select Codprovcli,nombre From PCProvcli"
                sql = sql & " Where IDprovcli = " & mProps.IdProvcli
                Set rs1 = mProps.objIVDescuento.Empresa.OpenRecordset(sql)
                If Not rs1.EOF Then
                    mProps.CodProvcli = rs1.Fields(0).value
                    mProps.Nombre = rs1.Fields(1).value
                End If
                rs1.Close
                Set rs1 = Nothing
            
            End If
        
        End If
    End With
End Sub

Public Function Clone() As IVDescuentoDetallePC
    Dim nuevo As IVDescuentoDetallePC
    
    'Copia las propiedades al nuevo objeto
    Set nuevo = New IVDescuentoDetallePC
    nuevo.Props = mProps
    
    Set Clone = nuevo
    Set nuevo = Nothing
End Function

Friend Property Let Props(value As T_PROPS)         'Para el método Clone
    mProps = value
    mProps.IdDescuento = 0
    'mbooModificado = True
End Property


Public Property Get IdDescuento() As Long
    IdDescuento = mProps.IdDescuento
End Property

Public Property Let IdDescuento(ByVal value As Long)
    mProps.IdDescuento = value
End Property

Public Property Get Orden() As Long
    Orden = mProps.Orden
End Property

Public Property Let Orden(ByVal value As Long)
    mProps.Orden = value
End Property

Public Property Let NumPCGrupo(ByVal value As Byte)
    mProps.NumPCGrupo = value
    mProps.objIVDescuento.Modificado = True
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
        Set rs = mProps.objIVDescuento.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then
            mProps.idPcGrupo = rs.Fields(0).value
        End If
        rs.Close
        Set rs = Nothing
    End If
    mProps.objIVDescuento.Modificado = True
End Property


Public Property Get idPcGrupo() As Long
    idPcGrupo = mProps.idPcGrupo
End Property



Private Sub Class_Initialize()
    mProps.Orden = 0
End Sub

Private Sub Class_Terminate()
    Set mProps.objIVDescuento = Nothing
End Sub

Public Property Get CodProvcli() As String
    CodProvcli = mProps.CodProvcli
End Property

Public Property Let CodProvcli(value As String)
    Dim rs As Recordset, sql As String
    mProps.CodProvcli = value
    If value = "" Then
        mProps.IdProvcli = 0
    Else
        'actualiza  el IDGrupo
        sql = "Select IdProvCli From PcProvCli"
        sql = sql & " Where CodProvCli = '" & value & "'"
        Set rs = mProps.objIVDescuento.Empresa.OpenRecordset(sql)
        If Not rs.EOF Then
            mProps.IdProvcli = rs.Fields(0).value
        End If
        rs.Close
        Set rs = Nothing
    End If
    mProps.objIVDescuento.Modificado = True
End Property

Public Property Get IdProvcli() As Long
    IdProvcli = mProps.IdProvcli
End Property



Public Sub recuperarNEW(key As Variant)
   Dim sql As String, i As Integer
    Dim rs As Recordset
    Dim cond As String
    
        cond = " CodpROVCLI='" & key & "'"
        If VarType(key) = vbString Then
            sql = "Select idprovcli, CodProvcli, nombre"
            sql = sql & " FROM pcprovcli WHERE" & cond
        Else
            sql = "Select CodProvcli, nombre"
            sql = sql & " FROM pcprovcli WHERE" & cond
        End If
    Set rs = Me.IVDescuento.Empresa.OpenRecordset(sql)
    If rs.EOF Then 'AUC cambiado antes estuvo NOT (RS.EOF)
       Err.Raise ERR_NOHAYCODIGO, "IVDescuentoDetallePC", MSGERR_NOHAYCODIGO
    Else
                mProps.IdProvcli = rs!IdProvcli
                mProps.CodProvcli = rs!CodProvcli
                mProps.Nombre = rs!Nombre


  End If
End Sub

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

Public Sub RecuperarNombre(key As Variant)
   Dim sql As String, i As Integer
    Dim rs As Recordset
    Dim cond As String
        cond = " Nombre ='" & key & "'"
       If VarType(key) = vbString Then
            sql = "Select idprovcli,codprovcli,nombre " & _
            "FROM pcprovcli WHERE bandcliente=1 and " & cond
        Else
            sql = "Select idprovcli,codprovcli,nombre " & _
            "FROM pcprovcli WHERE bandcliente=1 and " & cond
        End If
   Set rs = Me.IVDescuento.Empresa.OpenRecordset(sql)
    If rs.EOF Then 'AUC cambiado antes estuvo NOT (RS.EOF)
       Err.Raise ERR_NOHAYCODIGO, "IvFamiliaDetalle", MSGERR_NOHAYCODIGO
    Else
            mProps.IdProvcli = rs!IdProvcli
            mProps.CodProvcli = rs!CodProvcli
            mProps.Nombre = rs!Nombre
  End If
End Sub

Public Sub RecuperarNombrexCodigo(key As Variant)
   Dim sql As String, i As Integer
    Dim rs As Recordset
    Dim cond As String
        cond = " CodProvCli ='" & key & "'"
       If VarType(key) = vbString Then
            sql = "Select idprovcli,codprovcli,nombre " & _
            "FROM pcprovcli WHERE bandcliente=1 and " & cond
        Else
            sql = "Select idprovcli,codprovcli,nombre " & _
            "FROM pcprovcli WHERE bandcliente=1 and " & cond
        End If
   Set rs = Me.IVDescuento.Empresa.OpenRecordset(sql)
    If rs.EOF Then 'AUC cambiado antes estuvo NOT (RS.EOF)
       Err.Raise ERR_NOHAYCODIGO, "IvFamiliaDetalle", MSGERR_NOHAYCODIGO
    Else
            mProps.IdProvcli = rs!IdProvcli
            mProps.CodProvcli = rs!CodProvcli
            mProps.Nombre = rs!Nombre
  End If
End Sub


