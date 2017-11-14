VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RTurnoRotativo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public mbooEsNuevo As Boolean
Public mbooModificado As Boolean

Private Type T_Props
      IdTurno As Long
      CodigoTurno As String
      entrada As Date
      salida As Date
      objEmpresa As Empresa
End Type
Private mProps As T_Props

Friend Property Set objEmpresa(obj As Empresa)
    Set mProps.objEmpresa = obj
End Property
'
Public Property Get objEmpresa() As Empresa
    Set objEmpresa = mProps.objEmpresa
End Property

Public Property Get IdTurno() As Long
    IdTurno = mProps.IdTurno
End Property

Public Property Let IdTurno(ByVal Value As Long)
      mProps.IdTurno = Value
      mbooModificado = True
End Property

Public Property Get dia() As Long
    dia = mProps.dia
End Property

Public Property Let dia(ByVal Value As Long)
    mProps.dia = Value
    mbooModificado = True
End Property

Public Property Get entrada() As Date
    entrada = mProps.entrada
End Property

Public Property Let entrada(ByVal Value As Date)
        mProps.entrada = Value
        mbooModificado = True
End Property

Public Property Get salida() As Date
    salida = mProps.salida
End Property

Public Property Let salida(ByVal Value As Date)
        mProps.salida = Value
        mbooModificado = True
End Property


Public Sub grabar()
    Dim rs As Recordset
    Dim sql, fecha As String
    Dim cond As String 'h25,h100,
      sql = "Select idTurno,entrada,salida from RTurnoRotativo"
     If mbooModificado = False Then Exit Sub
      If mbooEsNuevo = True Then
          cond = ""
       Else
          cond = " where idturno = " & mProps.IdTurno
      End If
      Set rs = objEmpresa.OpenRecordsetParaEdit(sql & cond)
      If mbooEsNuevo = True Then rs.AddNew
      With rs
            rs!IdTurno = mProps.IdTurno
            fecha = HoraHMS(mProps.entrada, True, True)
            rs!entrada = CDate(fecha)
            fecha = HoraHMS(mProps.salida, True, True)
            rs!salida = CDate(fecha)
            rs.Update
        End With
        Set rs = Nothing
End Sub



Public Sub BorradoFisico(Id As Long)
    Dim sql As String
    sql = "delete from RTurnoRotativo where IdTurno = " & Id
    objEmpresa.EjecutarSQL sql, 1
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_TURNOS
End Sub

Public Function recuperar(ByVal key As Long) As Recordset
    Dim rs As Recordset
    Dim sql As String
    Dim i As Integer
    
    sql = "SELECT tr.IdTurno,tr.Entrada," & _
            " tr.salida " & _
            " FROM RTurno t INNER JOIN RTurnoRotativo tr  ON T.IdTurno = tr.IdTurno  " & _
            " WHERE t.idTurno=" & key
            
    Set rs = objEmpresa.OpenRecordsetParaEdit(sql)
    If rs.RecordCount > 0 Then
          Set recuperar = rs
            mbooEsNuevo = False
    Else
'        Err.Raise ERR_NOHAYCODIGO, "clsTurno", MSGERR_NOHAYCODIGO
        mbooEsNuevo = True
    End If
End Function


Public Sub Class_Initialize()
    mbooModificado = False
    mbooEsNuevo = True
End Sub

Public Property Get CodigoTurno() As String
    CodigoTurno = mProps.CodigoTurno
End Property

Public Property Let CodigoTurno(ByVal Value As String)
    Dim sql As String, rs As Recordset
    'Cuando cambia el valor
    If Value <> mProps.CodigoTurno Then
        If Len(Value) > 0 Then
            'Actualiza IdTurno también
            sql = "SELECT IdTurno FROM  RTurno WHERE codigo='" & Value & "'"
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
        mProps.CodigoTurno = Value
        mbooModificado = True
End Property


