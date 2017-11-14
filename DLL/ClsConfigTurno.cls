VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsConfigTurno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public mbooEsNuevo As Boolean
Public mbooModificado As Boolean

Private Type T_PROPS
      IdTurno As Long
      CodigoTurno As String
      tolentrada As Long
      tolentradarec As Long
      exceso As Long
      insuficiente As Long
      marcas As Long
      tolRecesoHE As Long
      objEmpresa As Empresa       'Referencia al objeto padre
End Type
Private mProps As T_PROPS

Friend Property Set objEmpresa(obj As Empresa)
      Set mProps.objEmpresa = obj
End Property

Public Property Get objEmpresa() As Empresa
      Set objEmpresa = mProps.objEmpresa
End Property

Public Property Get IdTurno() As Long
      IdTurno = mProps.IdTurno
End Property

Public Property Let IdTurno(ByVal value As Long)
      mProps.IdTurno = value
      mbooModificado = True
End Property

Public Property Get tolentrada() As Long
      tolentrada = mProps.tolentrada
End Property

Public Property Let tolentrada(ByVal value As Long)
      mProps.tolentrada = value
      mbooModificado = True
End Property

Public Property Get tolentradarec() As Long
      tolentradarec = mProps.tolentradarec
End Property

Public Property Let tolentradarec(ByVal value As Long)
      mProps.tolentradarec = value
      mbooModificado = True
End Property

Public Property Get exceso() As Long
      exceso = mProps.exceso
End Property

Public Property Let exceso(ByVal value As Long)
      mProps.exceso = value
      mbooModificado = True
End Property

Public Property Get insuficiente() As Long
      insuficiente = mProps.insuficiente
End Property

Public Property Let insuficiente(ByVal value As Long)
      mProps.insuficiente = value
      mbooModificado = True
End Property

Public Property Get marcas() As Long
      marcas = mProps.marcas
End Property

Public Property Let marcas(ByVal value As Long)
      mProps.marcas = value
      mbooModificado = True
End Property

Public Property Get tolRecesoHE() As Long
      tolRecesoHE = mProps.tolRecesoHE
End Property

Public Property Let tolRecesoHE(ByVal value As Long)
      mProps.tolRecesoHE = value
      mbooModificado = True
End Property


Public Sub Grabar(ByVal cod As Long)
      Dim rs  As Recordset
      Dim sql, Fecha As String
      Dim cond As String
      sql = "Select *  from RConfigTurno"
      With rs
            If mbooModificado = False Then Exit Sub
            If mbooEsNuevo = True Then
                  cond = ""
            Else
                  cond = " where IdTurno = " & cod
            End If
            Set rs = objEmpresa.OpenRecordsetParaEdit(sql & cond)
            If mbooEsNuevo = True Then rs.AddNew
            rs!IdTurno = cod
            rs!tolentrada = mProps.tolentrada
            rs!tolentradarec = mProps.tolentradarec
            rs!exceso = mProps.exceso
            rs!insuficiente = mProps.insuficiente
            rs!marcas = mProps.marcas
            rs!tolRecesoHE = mProps.tolRecesoHE
            rs.Update
      End With
      Set rs = Nothing
End Sub

Public Sub BorradoFisico(cod As Long)
      Dim sql As String
      sql = "delete from RConfigTurno where IdTurno = " & cod
      objEmpresa.EjecutarSQL sql, 1
End Sub

Public Sub Recuperar(ByVal key As Long)

    Dim rs As Recordset
    Dim sql As String
       
    sql = "SELECT T.IDTURNO, T.CodTurno, ct.TolEntrada, ct.TolEntradarec, ct.Insuficiente, ct.exceso,"
    sql = sql & " ct.marcas, ct.marcas,  ct.TolRecesoHE  "
    sql = sql & " FROM RTurno T INNER JOIN RConfigTurno ct ON t.IdTurno = ct.IdTurno"
    sql = sql & "  WHERE t.idturno=" & key

      Set rs = objEmpresa.OpenRecordsetParaEdit(sql)
      If rs.RecordCount > 0 Then
            With rs
                  mProps.IdTurno = !IdTurno
                  mProps.tolentrada = !tolentrada
                  mProps.tolentradarec = !tolentradarec
                  mProps.exceso = !exceso
                  mProps.insuficiente = !insuficiente
                  mProps.marcas = !marcas
                  If Not IsNull(rs!tolRecesoHE) Then
                        mProps.tolRecesoHE = rs!tolRecesoHE
                   Else
                        mProps.tolRecesoHE = 0
                    End If
            End With
            mbooEsNuevo = False
      Else
            'Err.Raise ERR_NOHAYCODIGO, "clsTurno", MSGERR_NOHAYCODIGO
            mbooEsNuevo = True
      End If
End Sub

Public Sub Class_Initialize()
      mbooModificado = False
      mbooEsNuevo = True
End Sub

Public Property Get CodigoTurno() As String
    CodigoTurno = mProps.CodigoTurno
End Property

Public Property Let CodigoTurno(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodigoTurno Then
        If Len(value) > 0 Then
            'Actualiza IdTurno también
            sql = "SELECT IdTurno FROM RTurno WHERE codigo='" & value & "'"
            Set rs = objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "ConfigTurno_Codturno", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdTurno = rs!IdTurno
            End If
            rs.Close
        Else
            IdTurno = 0
        End If
    End If
    Set rs = Nothing
        mProps.CodigoTurno = value
        mbooModificado = True
End Property

