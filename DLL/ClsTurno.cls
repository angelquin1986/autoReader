VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTurno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public mbooEsNuevo As Boolean
Public mbooModificado As Boolean

Private Type T_Props
    IdTurno As Long
    Codturno As String
    Descripcion As String
    diaslaborables As Integer
    diasdescanso As Integer
    horaentrada As Date
    Horasalida As Date
    h100entrada As Date
    h25entrada As Date
    bandReceso As Boolean
    horiniciores As Date
    horfinres As Date
    Tipo As Integer
    CodigoAutoriza As String
    TurnoNocturno As Boolean
    EventoCalculado As Boolean
    HorasTrabajadas As Boolean
    NumHoras As Integer
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

Public Property Get Codturno() As String
    Codturno = Trim$(mProps.Codturno)
End Property

Public Property Let Codturno(ByVal Value As String)
    If Len(Value) > 20 Then Err.Raise 9999, "clsTurno", "longitud maxima 20 caracteres"
    mProps.Codturno = Value
End Property
Public Property Get Descripcion() As String
    Descripcion = mProps.Descripcion
End Property
Public Property Let Tipo(ByVal Value As Integer)
    mProps.Tipo = Value
End Property
Public Property Get Tipo() As Integer
    Tipo = mProps.Tipo
End Property
Public Property Let Descripcion(ByVal Value As String)
    mProps.Descripcion = Value
    mbooModificado = True
End Property
Public Property Get diaslaborables() As Integer
    diaslaborables = mProps.diaslaborables
End Property
Public Property Let diaslaborables(ByVal Value As Integer)
        mProps.diaslaborables = Value
        mbooModificado = True
End Property
Public Property Get diasdescanso() As Integer
    diasdescanso = mProps.diasdescanso
End Property
Public Property Let diasdescanso(ByVal Value As Integer)
        mProps.diasdescanso = Value
        mbooModificado = True
End Property
Public Property Get horaentrada() As Date
    horaentrada = mProps.horaentrada
End Property
Public Property Let horaentrada(ByVal Value As Date)
        mProps.horaentrada = Value
        mbooModificado = True
End Property

Public Property Get h25entrada() As Date
    h25entrada = mProps.h25entrada
End Property

Public Property Let h25entrada(ByVal Value As Date)
        mProps.h25entrada = Value
        mbooModificado = True
End Property

Public Property Get Horasalida() As Date
    Horasalida = mProps.Horasalida
End Property

Public Property Let Horasalida(ByVal Value As Date)
        mProps.Horasalida = Value
        mbooModificado = True
End Property

Public Property Get h100entrada() As Date
    h100entrada = mProps.h100entrada
End Property

Public Property Let h100entrada(ByVal Value As Date)
        mProps.h100entrada = Value
        mbooModificado = True
End Property

Public Property Get receso() As Boolean
    receso = mProps.bandReceso
End Property

Public Property Let receso(ByVal Value As Boolean)
        mProps.bandReceso = Value
        mbooModificado = True
End Property

Public Property Get TurnoNocturno() As Boolean
    TurnoNocturno = mProps.TurnoNocturno
End Property

Public Property Let TurnoNocturno(ByVal Value As Boolean)
        mProps.TurnoNocturno = Value
        mbooModificado = True
End Property

Public Property Get EventoCalculado() As Boolean
    EventoCalculado = mProps.EventoCalculado
End Property

Public Property Let EventoCalculado(ByVal Value As Boolean)
        mProps.EventoCalculado = Value
        mbooModificado = True
End Property

Public Property Get horiniciores() As Date
    horiniciores = mProps.horiniciores
End Property

Public Property Let horiniciores(ByVal Value As Date)
        mProps.horiniciores = Value
        mbooModificado = True
End Property

Public Property Get horfinres() As Date
    horfinres = mProps.horfinres
End Property

Public Property Let horfinres(ByVal Value As Date)
        mProps.horfinres = Value
        mbooModificado = True
End Property

Public Function grabar()
    Dim rs As Recordset
    Dim sql, fecha As String
    Dim cond As String
    
     sql = "Select codTurno,descripcion,diaslaborables, "
     sql = sql & " diasdescanso,horaentrada,horasalida,"
     sql = sql & " bandreceso,Horiniciores,horfinres,tipo, "
     sql = sql & " TurnoNocturno, EventoCalculado,HorasTrabajadas,NumHoras "
     sql = sql & " from RTurno"
        With rs
            If mbooModificado = False Then Exit Function
            If mbooEsNuevo = True Then
                cond = ""
             Else
                cond = " where IdTurno = " & mProps.IdTurno
            End If
            Set rs = objEmpresa.OpenRecordsetParaEdit(sql & cond)
            If mbooEsNuevo = True Then rs.AddNew
            rs!Codturno = mProps.Codturno
            rs!Descripcion = mProps.Descripcion
            rs!diaslaborables = mProps.diaslaborables
            rs!diasdescanso = mProps.diasdescanso
            fecha = HoraHMS(mProps.horaentrada, True, True)
            rs!horaentrada = CDate(fecha)
           ' Fecha = HoraHMS(mProps.h25entrada, True, True)
           ' rs!h25entrada = CDate(Fecha)
            fecha = HoraHMS(mProps.Horasalida, True, True)
            rs!Horasalida = CDate(fecha)
'            Fecha = HoraHMS(mProps.h100entrada, True, True)
 '           rs!h100entrada = CDate(Fecha)
            rs!bandReceso = mProps.bandReceso
            rs!TurnoNocturno = mProps.TurnoNocturno
            rs!EventoCalculado = mProps.EventoCalculado
            If mProps.bandReceso = True Then
                fecha = HoraHMS(mProps.horiniciores, True, True)
                rs!horiniciores = CDate(fecha)
                fecha = HoraHMS(mProps.horfinres, True, True)
                rs!horfinres = CDate(fecha)
            Else
                fecha = HoraHMS(0, True, True)
                rs!horiniciores = CDate(fecha)
                fecha = HoraHMS(0, True, True)
                rs!horfinres = CDate(fecha)
            End If
            rs!Tipo = mProps.Tipo
            rs!HorasTrabajadas = mProps.HorasTrabajadas
            rs!NumHoras = mProps.NumHoras
            
            rs.Update
        End With
      If mbooEsNuevo = True Then
            sql = "Select IdTurno from RTurno"
            Set rs = objEmpresa.OpenRecordsetParaEdit(sql)
            rs.MoveLast
            grabar = rs.Fields(0)
      End If
      mProps.objEmpresa.RaiseRefreshEvent REFRESH_TURNOS
End Function

Public Sub BorradoFisico(cod As Long)
    Dim sql As String
    Dim rs As Recordset
        'Verifica si ya existe un item usando el mismo CodGrupo
    sql = "SELECT CodProvCli FROM PCProvCli " & _
          " WHERE Idturno=" & cod
    
    Set rs = mProps.objEmpresa.OpenRecordset(sql)
    If Not rs.EOF Then
        Err.Raise ERR_NOELIMINABLE, "Turno_Eliminar", MSGERR_NOELIMINABLE
    End If
    rs.Close
    sql = "delete from RTurno where IdTurno = " & cod
    objEmpresa.EjecutarSQL sql, 1
    
    sql = "delete from RTurnoRotativo where IdTurno = " & cod
    objEmpresa.EjecutarSQL sql, 1
    
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_TURNOS
End Sub

Public Sub recuperar(ByVal key As Long)
    Dim rs As Recordset
    Dim sql As String
        sql = "SELECT * FROM RTurno WHERE IdTurno=" & key
    Set rs = objEmpresa.OpenRecordsetParaEdit(sql)
    If rs.RecordCount > 0 Then
        With rs
            mProps.IdTurno = !IdTurno
            If Not IsNull(!Codturno) Then mProps.Codturno = !Codturno
            If Not IsNull(!Descripcion) Then mProps.Descripcion = !Descripcion
            If Not IsNull(!diaslaborables) Then mProps.diaslaborables = !diaslaborables
            If Not IsNull(!diasdescanso) Then mProps.diasdescanso = !diasdescanso
            If Not IsNull(!horaentrada) Then mProps.horaentrada = !horaentrada
            If Not IsNull(!Horasalida) Then mProps.Horasalida = !Horasalida
            'If Not IsNull(!h100entrada) Then mProps.h100entrada = !h100entrada
            'If Not IsNull(!h25entrada) Then mProps.h25entrada = !h25entrada
            mProps.bandReceso = !bandReceso
            mProps.TurnoNocturno = !TurnoNocturno
            mProps.EventoCalculado = !EventoCalculado
            If Not IsNull(!horiniciores) Then mProps.horiniciores = !horiniciores
            If Not IsNull(!horfinres) Then mProps.horfinres = !horfinres
            If Not IsNull(!Tipo) Then mProps.Tipo = !Tipo
            If Not IsNull(!HorasTrabajadas) Then mProps.HorasTrabajadas = !HorasTrabajadas
            If Not IsNull(!NumHoras) Then mProps.NumHoras = !NumHoras
        End With
            mbooEsNuevo = False
    Else
        mbooEsNuevo = True
    End If
End Sub
Public Sub Class_Initialize()
            mbooModificado = False
            mbooEsNuevo = True
End Sub
Public Property Get HorasTrabajadas() As Boolean
    HorasTrabajadas = mProps.HorasTrabajadas
End Property
Public Property Let HorasTrabajadas(ByVal Value As Boolean)
        mProps.HorasTrabajadas = Value
        mbooModificado = True
End Property
Public Property Get NumHoras() As Integer
    NumHoras = mProps.NumHoras
End Property
Public Property Let NumHoras(ByVal Value As Integer)
        mProps.NumHoras = Value
        mbooModificado = True
End Property

Public Sub EliminaTurnoRotativo(id As Variant)
    Dim sql As String
    Dim rs As Recordset
    sql = "delete from RTurnoRotativo where IdTurno = " & id
    objEmpresa.EjecutarSQL sql, 1
End Sub
