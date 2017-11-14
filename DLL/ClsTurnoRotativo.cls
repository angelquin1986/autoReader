VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTurnoRotativo"
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
      InicioReceso As Date
      FinReceso As Date
      DesdeH25 As Date
      HastaH25 As Date
      Orden As Integer
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

Public Property Get Orden() As Long
    Orden = mProps.Orden
End Property

Public Property Let Orden(ByVal Value As Long)
    mProps.Orden = Value
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
      sql = "Select idTurno,entrada,salida,orden,InicioReceso,FinReceso,DesdeH25,HastaH25 from RTurnoRotativo"
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
            rs!Orden = mProps.Orden
            rs!InicioReceso = mProps.InicioReceso
            rs!FinReceso = mProps.FinReceso
            rs!DesdeH25 = mProps.DesdeH25
            rs!HastaH25 = mProps.HastaH25
            rs.Update
        End With
        Set rs = Nothing
End Sub



Public Sub BorradoFisico(id As Long)
    Dim sql As String
    sql = "delete from RTurnoRotativo where IdTurno = " & id
    objEmpresa.EjecutarSQL sql, 1
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_TURNOS
End Sub

Public Function recuperar(ByVal key As Long) As Recordset
    Dim rs As Recordset
    Dim sql As String
    Dim i As Integer
    
    sql = "SELECT tr.IdTurno,tr.Entrada," & _
            " tr.salida,tr.InicioReceso,tr.finreceso  " & _
            " tr.DesdeH25,tr.HastaH25 " & _
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

Public Property Get InicioReceso() As Date
    InicioReceso = mProps.InicioReceso
End Property

Public Property Let InicioReceso(ByVal Value As Date)
        mProps.InicioReceso = Value
        mbooModificado = True
End Property


Public Property Get FinReceso() As Date
    FinReceso = mProps.FinReceso
End Property

Public Property Let FinReceso(ByVal Value As Date)
        mProps.FinReceso = Value
        mbooModificado = True
End Property

Public Property Get DesdeH25() As Date
    DesdeH25 = mProps.DesdeH25
End Property

Public Property Let DesdeH25(ByVal Value As Date)
        mProps.DesdeH25 = Value
        mbooModificado = True
End Property

Public Property Get HastaH25() As Date
    HastaH25 = mProps.HastaH25
End Property

Public Property Let HastaH25(ByVal Value As Date)
        mProps.HastaH25 = Value
        mbooModificado = True
End Property
