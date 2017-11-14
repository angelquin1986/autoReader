VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTurnoMensualDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public mbooEsNuevo As Boolean
Public mbooModificado As Boolean
Private Type T_Props
      id As Long
      IdTurnoM As Long
      fecha As Date
      IdEmpleado As Long
      CodEmpleado As String
      IdTurno As Long
      CodigoTurno As String
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

Public Property Get IdTurnoM() As Long
    IdTurnoM = mProps.IdTurnoM
End Property

Public Property Let IdTurnoM(ByVal Value As Long)
      mProps.IdTurnoM = Value
      mbooModificado = True
End Property

Public Property Get fecha() As Date
    fecha = mProps.fecha
End Property

Public Property Let fecha(ByVal Value As Date)
    mProps.fecha = Value
    mbooModificado = True
End Property


Public Sub grabar()
    Dim rs As Recordset
    Dim sql, fecha As String
    Dim cond As String
      sql = "Select IdTurnom,fecha,idEmpleado,idturno from RTurnoMensualDetalle"
     If mbooModificado = False Then Exit Sub
      If mbooEsNuevo = True Then
          cond = ""
       Else
          cond = " where IdTurnoM = " & mProps.IdTurnoM & " and fecha= " & mProps.fecha
          cond = cond & " and idEmpleado = " & mProps.IdEmpleado
      End If
      Set rs = objEmpresa.OpenRecordsetParaEdit(sql & cond)
      If mbooEsNuevo = True Then rs.AddNew
      With rs
            rs!IdTurnoM = mProps.IdTurnoM
            rs!IdEmpleado = mProps.IdEmpleado
            rs!IdTurno = mProps.IdTurno
            rs!fecha = mProps.fecha
            rs.Update
        End With
        Set rs = Nothing
End Sub



Public Sub BorradoFisico(cod As Long)
    Dim sql As String
    sql = "delete from RTurnoMensualDetalle where IdTurnoM = " & cod
    objEmpresa.EjecutarSQL sql, 1
    
End Sub



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
            sql = "SELECT IdTurno FROM RTurno WHERE codigo='" & Value & "'"
            Set rs = objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "TurnoMensualDetalle_CodTurno", MSGERR_NOHAYCODIGO
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

Public Property Let IdTurno(ByVal Value As Long)
    mProps.IdTurno = Value
    mbooModificado = True
End Property

Public Property Get IdTurno() As Long
    IdTurno = mProps.IdTurno
End Property

Public Property Get CodEmpleado() As String
    CodEmpleado = mProps.CodEmpleado
End Property

Public Property Let CodEmpleado(ByVal Value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If Value <> mProps.CodEmpleado Then
        If Len(Value) > 0 Then
            'Actualiza idEmpleado también
            sql = "SELECT idEmpleado FROM Persona WHERE codigo='" & Value & "'"
            Set rs = objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "grupo_Codsecciom", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdEmpleado = rs!IdEmpleado
            End If
            rs.Close
        Else
            IdEmpleado = 0
        End If
    End If

    Set rs = Nothing


        mProps.CodEmpleado = Value
        mbooModificado = True
End Property

Public Property Let IdEmpleado(ByVal Value As Long)
    mProps.IdEmpleado = Value
    mbooModificado = True
End Property

Public Property Get IdEmpleado() As Long
    IdEmpleado = mProps.IdEmpleado
End Property


Public Function recuperar(ByVal key As Variant)
    Dim rs As Recordset
    Dim sql As String
    Dim i As Integer
    sql = "Select IdTurnoM,fecha,idEmpleado,idturno from RTurnoMensualDetalle"
    sql = sql & " WHERE IdTurnoD=" & key
    Set rs = objEmpresa.OpenRecordsetParaEdit(sql)
    If rs.RecordCount > 0 Then
        With rs
                    If Not IsNull(!IdTurnoM) Then mProps.IdTurnoM = !IdTurnoM
                    If Not IsNull(!fecha) Then mProps.fecha = !fecha
                    If Not IsNull(!IdEmpleado) Then mProps.IdEmpleado = !IdEmpleado
                    If Not IsNull(!IdTurno) Then mProps.IdTurno = !IdTurno
        End With
            mbooEsNuevo = False
      Set rs = Nothing
            mbooEsNuevo = False
    Else
'        Err.Raise ERR_NOHAYCODIGO, "clsTurno", MSGERR_NOHAYCODIGO
        mbooEsNuevo = True
    End If
    
End Function

