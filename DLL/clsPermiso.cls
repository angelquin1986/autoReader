VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPermiso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public mbooEsNuevo As Boolean
Public mbooModificado As Boolean

Private Type T_PROPS
      IdPermiso As Long
      IdEmpleado As Long
      CodigoEmpleado As String
      IdMotivo As Long
      CodigoMotivo As String
      HoraInicio As Date
      HoraFinal As Date
      FechaInicio As Date
      FechaFinal As Date
      Observa As String
      Autoriza As String
      objEmpresa As Empresa       'Referencia al objeto padre
End Type
Private mProps As T_PROPS

Friend Property Set objEmpresa(obj As Empresa)
      Set mProps.objEmpresa = obj
End Property

Public Property Get objEmpresa() As Empresa
      Set objEmpresa = mProps.objEmpresa
End Property

Public Property Get IdEmpleado() As Long
      IdEmpleado = mProps.IdEmpleado
End Property

Public Property Let IdEmpleado(ByVal value As Long)
      mProps.IdEmpleado = value
      mbooModificado = True
End Property

Public Property Get IdMotivo() As Long
      IdMotivo = mProps.IdMotivo
End Property

Public Property Let IdMotivo(ByVal value As Long)
      mProps.IdMotivo = value
      mbooModificado = True
End Property

Public Property Get HoraInicio() As Date
      HoraInicio = mProps.HoraInicio
End Property

Public Property Let HoraInicio(ByVal value As Date)
      mProps.HoraInicio = value
      mbooModificado = True
End Property

Public Property Get HoraFinal() As Date
      HoraFinal = mProps.HoraFinal
End Property

Public Property Let HoraFinal(ByVal value As Date)
      mProps.HoraFinal = value
      mbooModificado = True
End Property
Public Property Get FechaInicio() As Date
      FechaInicio = mProps.FechaInicio
End Property

Public Property Let FechaInicio(ByVal value As Date)
      mProps.FechaInicio = value
      mbooModificado = True
End Property

Public Property Get FechaFinal() As Date
      FechaFinal = mProps.FechaFinal
End Property

Public Property Let FechaFinal(ByVal value As Date)
      mProps.FechaFinal = value
      mbooModificado = True
End Property

Public Sub Grabar()
      Dim rs  As Recordset
      Dim sql, fecha As String
      Dim cond As String
      sql = "Select *  from RPermiso"
      With rs
            If mbooModificado = False Then Exit Sub
            If mbooEsNuevo = True Then
                cond = ""
             Else
                cond = " where Idpermiso = " & mProps.IdPermiso
            End If
            Set rs = objEmpresa.OpenRecordsetParaEdit(sql)
            If mbooEsNuevo = True Then rs.AddNew
            rs!IdEmpleado = mProps.IdEmpleado
            rs!IdMotivo = mProps.IdMotivo
            rs!HoraInicio = mProps.HoraInicio
            rs!HoraFinal = mProps.HoraFinal
            rs!FechaInicio = mProps.FechaInicio
            rs!FechaFinal = mProps.FechaFinal
            rs!Observa = mProps.Observa
            rs!Autoriza = mProps.Autoriza
            rs.Update
      End With
      Set rs = Nothing
      mProps.objEmpresa.RaiseRefreshEvent REFRESH_RPERMISO
      mbooEsNuevo = False
     mbooModificado = False
End Sub

Public Sub BorradoFisico(Id As Long)
      Dim sql As String
      sql = "delete from RPermiso where IdPermiso = " & Id
      objEmpresa.EjecutarSQL sql, 1
      mProps.objEmpresa.RaiseRefreshEvent REFRESH_RPERMISO
End Sub

Public Sub Recuperar(ByVal key As Long)

    Dim rs As Recordset
    Dim sql As String
    
    sql = "SELECT rp.idpermiso,rp.idempleado,pc.codProvCli,rp.idmotivo,m.codigo as codmotivo,"
    sql = sql & " horainicio,horafinal,fechainicio,fechafinal,observa,autoriza "
    sql = sql & " FROM RPermiso rp Left Join Empleado pc ON pc.IdProvcli = rp.idEmpleado"
    sql = sql & " Left Join rmotivo m ON m.Idmotivo = rp.idmotivo"
    sql = sql & "  WHERE rp.idPermiso=" & key

      Set rs = objEmpresa.OpenRecordsetParaEdit(sql)
      If rs.RecordCount > 0 Then
            With rs
                  mProps.IdPermiso = !IdPermiso
                  mProps.IdEmpleado = !IdEmpleado
                  mProps.CodigoEmpleado = !CodProvCli
                  mProps.IdMotivo = !IdMotivo
                  mProps.CodigoMotivo = !CodMotivo
                  mProps.HoraInicio = !HoraInicio
                  mProps.HoraFinal = !HoraFinal
                  mProps.FechaInicio = !FechaInicio
                  mProps.FechaFinal = !FechaFinal
                  mProps.Observa = !Observa
                  mProps.Autoriza = !Autoriza
                  
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

Public Property Get CodigoEmpleado() As String
    CodigoEmpleado = mProps.CodigoEmpleado
End Property

Public Property Let CodigoEmpleado(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodigoEmpleado Then
        If Len(value) > 0 Then
            'Actualiza IdTurno también
            sql = "SELECT IdProvcli FROM Empleado WHERE codprovcli='" & value & "'"
            Set rs = objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "RPermiso_CodEmpleado", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdEmpleado = rs!IdProvCli
            End If
            rs.Close
        Else
            IdEmpleado = 0
        End If
    End If
    Set rs = Nothing
        mProps.CodigoEmpleado = value
        mbooModificado = True
End Property


Public Property Get CodigoMotivo() As String
    CodigoMotivo = mProps.CodigoMotivo
End Property

Public Property Let CodigoMotivo(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodigoMotivo Then
        If Len(value) > 0 Then
            'Actualiza IdTurno también
            sql = "SELECT IdMotivo FROM RMotivo WHERE codigo='" & value & "'"
            Set rs = objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "RPermiso_CodMotivo", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                IdMotivo = rs!IdMotivo
            End If
            rs.Close
        Else
            IdMotivo = 0
        End If
    End If
    Set rs = Nothing
        mProps.CodigoMotivo = value
        mbooModificado = True
End Property

Public Property Get Observa() As String
      Observa = mProps.Observa
End Property

Public Property Let Observa(ByVal value As String)
      mProps.Observa = value
      mbooModificado = True
End Property

Public Property Get Autoriza() As String
      Autoriza = mProps.Autoriza
End Property

Public Property Let Autoriza(ByVal value As String)
      mProps.Autoriza = value
      mbooModificado = True
End Property
