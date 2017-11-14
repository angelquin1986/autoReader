VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsJornada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public mbooEsNuevo As Boolean
Public mbooModificado As Boolean

Private Type T_PROPS
    IdJornada As Long
    IdEmpleado As Long
    CodigoPersona As String
    fecha As Date
    Hora As Date
    secuencia As Long
    HoraTurno As Date
    Marca1 As Date
    Marca2 As Date
    Marca3 As Date
    Marca4 As Date
    Marca5 As Date
    Marca6 As Date
    Marca7 As Date
    Marca8 As Date
    Marca9 As Date
    Marca10 As Date
    Marca11 As Date
    Marca12 As Date
    Marca13 As Date
    Marca14 As Date
    Marca15 As Date
    Marca16 As Date
    Marca17 As Date
    Marca18 As Date
    Marca19 As Date
    Marca20 As Date
    CodigoReloj As String
    idProveedor As Long
    objEmpresa As Empresa
End Type
'Private objBD As SiiRelojdll.ClasMain
Private mProps As T_PROPS

Friend Property Set objEmpresa(obj As Empresa)
    Set mProps.objEmpresa = obj
End Property

Public Property Get objEmpresa() As Empresa
    Set objEmpresa = mProps.objEmpresa
End Property

Public Property Get IdJornada() As Long
    IdJornada = mProps.IdJornada
End Property

Public Property Let IdJornada(ByVal value As Long)
    mProps.IdJornada = value
End Property


Public Property Get IdEmpleado() As Long
    IdEmpleado = mProps.IdEmpleado
End Property

Public Property Let IdEmpleado(ByVal value As Long)
        mProps.IdEmpleado = value
      mbooModificado = True
End Property


Public Property Get fecha() As Date
    fecha = mProps.fecha
End Property

Public Property Let fecha(ByVal value As Date)
        mProps.fecha = value
        mbooModificado = True
End Property

Public Property Get Hora() As Date
    Hora = mProps.Hora
End Property

Public Property Let Hora(ByVal value As Date)
        mProps.Hora = value
        mbooModificado = True
End Property

Public Property Get secuencia() As Long
    secuencia = mProps.secuencia
End Property

Public Property Let secuencia(ByVal value As Long)
        mProps.secuencia = value
        mbooModificado = True
End Property


'Public Property Set BaseDatos(ByRef obj As ClasMain)
'    Set objBD = obj
'End Property

Public Sub Grabar()
    Dim rs As Recordset
    Dim sql As String
    Dim s As String
    Dim cond As String
    Dim fec As String
    Dim evento As String
     sql = "Select * from RAsistencia"
'     sql = sql & " Where idempleado  =" & mProps.IdEmpleado
'     sql = sql & " And fecha  = '" & mProps.fecha & "'"
'     sql = sql & " And marca1  =  '" & mProps.Marca1 & "'"
'     sql = sql & " And marca2  = '" & mProps.Marca2 & "'"
'     sql = sql & " And marca3  = '" & mProps.Marca3 & "'"
'     sql = sql & " And marca4  = '" & mProps.Marca4 & "'"
'     sql = sql & " And marca5  = '" & mProps.Marca5 & "'"
'     sql = sql & " And marca6  = '" & mProps.Marca6 & "'"
'     sql = sql & " And marca7  =  '" & mProps.Marca7 & "'"
'     sql = sql & " And marca8  = '" & mProps.Marca8 & "'"
'     sql = sql & " And marca9  = '" & mProps.Marca9 & "'"
'     sql = sql & " And marca10 = '" & mProps.Marca10 & "'"
'     Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
     
     
    If mbooModificado = False Then Exit Sub
    If mbooEsNuevo = True Then
        cond = ""
     Else
        cond = " where Idjornada = " & mProps.IdJornada
    End If
    sql = sql & cond
        Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
    With rs
        If mbooEsNuevo = True Then rs.AddNew
        rs!IdEmpleado = mProps.IdEmpleado
        rs!fecha = mProps.fecha
        s = HoraHMS(mProps.Hora, Me.objEmpresa.TipoDB, True, True)
        rs!Hora = s
        'rs!Hora = mProps.Hora' AUC cambiando no deja hacer sql
        rs!secuencia = mProps.secuencia
        rs!HoraTurno = mProps.HoraTurno
        rs!Marca1 = mProps.Marca1
        rs!Marca2 = mProps.Marca2
        rs!Marca3 = mProps.Marca3
        rs!Marca4 = mProps.Marca4
        rs!Marca5 = mProps.Marca5
        rs!Marca6 = mProps.Marca6
        rs!Marca7 = mProps.Marca7
        rs!Marca8 = mProps.Marca8
        rs!Marca9 = mProps.Marca9
        rs!Marca10 = mProps.Marca10
        rs!Marca11 = mProps.Marca11
        rs!Marca12 = mProps.Marca12
        rs!Marca13 = mProps.Marca13
        rs!Marca14 = mProps.Marca14
        rs!Marca15 = mProps.Marca15
        rs!Marca16 = mProps.Marca16
        rs!Marca17 = mProps.Marca17
        rs!Marca18 = mProps.Marca18
        rs!Marca19 = mProps.Marca19
        rs!Marca20 = mProps.Marca20
        rs!idProveedor = mProps.idProveedor
        rs.Update
    End With
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_JORNADA

End Sub

Public Sub BorradoFisico(cod As String)
    Dim sql As String
    
    sql = "delete from RAsistencia where idJornada = " & cod
    gobjMain.EmpresaActual.EjecutarSQL sql, 1
    'mProps.objmain.RaiseRefreshEvent REFRESH_JORNADA
End Sub

Public Sub Recuperar(ByVal key As Variant)
    Dim rs As Recordset
    Dim sql As String
     sql = "Select IdJornada,IdEmpleado,fecha,hora,secuencia,horaturno from RAsistencia WHERE IdJornada=" & key
    Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
    If rs.RecordCount > 0 Then
        With rs
            If Not IsNull(!IdJornada) Then mProps.IdJornada = !IdJornada
            If Not IsNull(!IdEmpleado) Then mProps.IdEmpleado = !IdEmpleado
            If Not IsNull(!fecha) Then mProps.fecha = !fecha
            If Not IsNull(!Hora) Then mProps.Hora = !Hora
            If Not IsNull(!secuencia) Then mProps.secuencia = !secuencia
            If Not IsNull(!HoraTurno) Then mProps.HoraTurno = !HoraTurno
            If Not IsNull(!Marca1) Then mProps.Marca1 = !Marca1
            If Not IsNull(!Marca2) Then mProps.Marca2 = !Marca2
            If Not IsNull(!Marca3) Then mProps.Marca3 = !Marca3
            If Not IsNull(!Marca4) Then mProps.Marca4 = !Marca4
            If Not IsNull(!Marca5) Then mProps.Marca5 = !Marca5
            If Not IsNull(!Marca6) Then mProps.Marca6 = !Marca6
            If Not IsNull(!Marca7) Then mProps.Marca7 = !Marca7
            If Not IsNull(!Marca8) Then mProps.Marca8 = !Marca8
            If Not IsNull(!Marca9) Then mProps.Marca9 = !Marca9
            If Not IsNull(!Marca10) Then mProps.Marca10 = !Marca10
            If Not IsNull(!Marca11) Then mProps.Marca11 = !Marca11
            If Not IsNull(!Marca12) Then mProps.Marca12 = !Marca12
            If Not IsNull(!Marca13) Then mProps.Marca13 = !Marca13
            If Not IsNull(!Marca14) Then mProps.Marca14 = !Marca14
            If Not IsNull(!Marca15) Then mProps.Marca15 = !Marca15
            If Not IsNull(!Marca16) Then mProps.Marca16 = !Marca16
            If Not IsNull(!Marca17) Then mProps.Marca17 = !Marca17
            If Not IsNull(!Marca18) Then mProps.Marca18 = !Marca18
            If Not IsNull(!Marca19) Then mProps.Marca19 = !Marca19
            If Not IsNull(!Marca20) Then mProps.Marca20 = !Marca20
             mProps.idProveedor = !idProveedor
        End With
            mbooEsNuevo = False
    Else
'        Err.Raise ERR_NOHAYCODIGO, "clsSeccion", MSGERR_NOHAYCODIGO
    End If
    Set rs = Nothing
End Sub

Public Sub Class_Initialize()
            mbooModificado = False
            mbooEsNuevo = True
End Sub

Public Sub recuperarParaTransportar(ByRef key As String, ByRef key1 As Date, ByRef key2 As Date)
'    Dim rs As Recordset
'    Dim sql As String
'      ''sql = "Select * from Asistencia WHERE codigoPersona=" & "'" & key & "'" & " AND  fecha=DateValue('" & key1 & "') AND  Hora= TimeValue('" & key2 & "')"    ' con esto funciona AUC
'        sql = "SELECT idJornada,persona.idpersona,persona.Codigo,Fecha,Hora,Secuencia, IdCliente, horaTurno FROM asistencia inner join persona on asistencia.Idpersona=persona.idpersona " & _
'        "WHERE persona.codigo=" & "'" & key & "'" & " AND  asistencia.fecha=DateValue('" & key1 & "') AND asistencia.Hora= timeValue('" & key2 & "')"
'
'    Set rs = mProps.objmain.OpenRecordsetParaEdit(sql)
'    If rs.RecordCount > 0 Then
'        With rs
'            If Not IsNull(!IdJornada) Then mProps.IdJornada = !IdJornada
'            If Not IsNull(!IdEmpleado) Then mProps.IdEmpleado = !IdEmpleado
'            If Not IsNull(!fecha) Then mProps.fecha = !fecha
'            If Not IsNull(!Hora) Then mProps.Hora = !Hora
'            If Not IsNull(!secuencia) Then mProps.secuencia = !secuencia
'            If Not IsNull(!HoraTurno) Then mProps.HoraTurno = !HoraTurno
'        End With
'            mbooEsNuevo = False
''    Else
''        Err.Raise ERR_NOHAYCODIGO, "clsjornada", MSGERR_NOHAYCODIGO
''      mbooEsNuevo = True
''        Exit Sub
'    End If
'    Set rs = Nothing
End Sub


Public Property Get CodigoPersona() As String
    CodigoPersona = mProps.CodigoPersona
End Property

Public Property Let CodigoPersona(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodigoPersona Then
        If Len(value) > 0 Then
            'Actualiza IdPersona también
            sql = "SELECT IdProvCli FROM Pcprovcli WHERE codprovcli='" & value & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Pcprovcli_CodEmpleado", MSGERR_NOHAYCODIGO
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
        mProps.CodigoPersona = value
        mbooModificado = True
End Property


Public Property Get CodigoReloj() As String
    CodigoReloj = mProps.CodigoReloj
End Property

Public Property Let CodigoReloj(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodigoReloj Then
        If Len(value) > 0 Then
            'Actualiza IdPersona también
            sql = "SELECT IdProvCli FROM Pcprovcli pc Inner Join Personal p on p.idEmpleado = pc.idprovcli WHERE  p.CodigoReloj='" & value & "'"
            Set rs = gobjMain.EmpresaActual.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "Pcprovcli_CodEmpleado", MSGERR_NOHAYCODIGO
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
        mProps.CodigoReloj = value
        mbooModificado = True
End Property

Public Property Get HoraTurno() As Date
    HoraTurno = mProps.HoraTurno
End Property

Public Property Let HoraTurno(ByVal value As Date)
        mProps.HoraTurno = value
        mbooModificado = True
End Property

Public Property Get Marca1() As Date
    Marca1 = mProps.Marca1
End Property

Public Property Let Marca1(ByVal value As Date)
        mProps.Marca1 = value
        mbooModificado = True
End Property

Public Property Get Marca2() As Date
    Marca2 = mProps.Marca2
End Property

Public Property Let Marca2(ByVal value As Date)
        mProps.Marca2 = value
        mbooModificado = True
End Property
Public Property Get Marca3() As Date
    Marca3 = mProps.Marca3
End Property

Public Property Let Marca3(ByVal value As Date)
        mProps.Marca3 = value
        mbooModificado = True
End Property
Public Property Get Marca4() As Date
    Marca4 = mProps.Marca4
End Property

Public Property Let Marca4(ByVal value As Date)
        mProps.Marca4 = value
        mbooModificado = True
End Property
Public Property Get Marca5() As Date
    Marca5 = mProps.Marca5
End Property

Public Property Let Marca5(ByVal value As Date)
        mProps.Marca5 = value
        mbooModificado = True
End Property
Public Property Get Marca6() As Date
    Marca6 = mProps.Marca6
End Property

Public Property Let Marca6(ByVal value As Date)
        mProps.Marca6 = value
        mbooModificado = True
End Property
Public Property Get Marca7() As Date
    Marca7 = mProps.Marca7
End Property

Public Property Let Marca7(ByVal value As Date)
        mProps.Marca7 = value
        mbooModificado = True
End Property
Public Property Get Marca8() As Date
    Marca8 = mProps.Marca8
End Property

Public Property Let Marca8(ByVal value As Date)
        mProps.Marca8 = value
        mbooModificado = True
End Property
Public Property Get Marca9() As Date
    Marca9 = mProps.Marca9
End Property

Public Property Let Marca9(ByVal value As Date)
        mProps.Marca9 = value
        mbooModificado = True
End Property
Public Property Get Marca10() As Date
    Marca10 = mProps.Marca10
End Property

Public Property Let Marca10(ByVal value As Date)
        mProps.Marca10 = value
        mbooModificado = True
End Property

Public Property Get Marca11() As Date
    Marca11 = mProps.Marca11
End Property

Public Property Let Marca11(ByVal value As Date)
        mProps.Marca11 = value
        mbooModificado = True
End Property
Public Property Get Marca12() As Date
    Marca12 = mProps.Marca12
End Property

Public Property Let Marca12(ByVal value As Date)
        mProps.Marca12 = value
        mbooModificado = True
End Property
Public Property Get Marca13() As Date
    Marca13 = mProps.Marca13
End Property

Public Property Let Marca13(ByVal value As Date)
        mProps.Marca13 = value
        mbooModificado = True
End Property
Public Property Get Marca14() As Date
    Marca14 = mProps.Marca14
End Property

Public Property Let Marca14(ByVal value As Date)
        mProps.Marca14 = value
        mbooModificado = True
End Property
Public Property Get Marca15() As Date
    Marca15 = mProps.Marca15
End Property

Public Property Let Marca15(ByVal value As Date)
        mProps.Marca15 = value
        mbooModificado = True
End Property
Public Property Get Marca16() As Date
    Marca16 = mProps.Marca16
End Property

Public Property Let Marca16(ByVal value As Date)
        mProps.Marca16 = value
        mbooModificado = True
End Property
Public Property Get Marca17() As Date
    Marca17 = mProps.Marca17
End Property

Public Property Let Marca17(ByVal value As Date)
        mProps.Marca17 = value
        mbooModificado = True
End Property
Public Property Get Marca18() As Date
    Marca18 = mProps.Marca18
End Property

Public Property Let Marca18(ByVal value As Date)
        mProps.Marca18 = value
        mbooModificado = True
End Property

Public Property Get Marca19() As Date
    Marca19 = mProps.Marca19
End Property

Public Property Let Marca19(ByVal value As Date)
        mProps.Marca19 = value
        mbooModificado = True
End Property
Public Property Get Marca20() As Date
    Marca20 = mProps.Marca20
End Property

Public Property Let Marca20(ByVal value As Date)
        mProps.Marca20 = value
        mbooModificado = True
End Property




Public Sub recuperarResumido(ByVal key As Variant)
    Dim rs As Recordset
    Dim sql As String
     sql = "Select IdJornada,IdEmpleado,fecha,hora,secuencia,horaturno from RAsistencia WHERE IdJornada=" & key
    Set rs = gobjMain.EmpresaActual.OpenRecordsetParaEdit(sql)
    If rs.RecordCount > 0 Then
        With rs
            If Not IsNull(!IdJornada) Then mProps.IdJornada = !IdJornada
            If Not IsNull(!IdEmpleado) Then mProps.IdEmpleado = !IdEmpleado
            If Not IsNull(!fecha) Then mProps.fecha = !fecha
            If Not IsNull(!Hora) Then mProps.Hora = !Hora
            If Not IsNull(!secuencia) Then mProps.secuencia = !secuencia
            If Not IsNull(!HoraTurno) Then mProps.HoraTurno = !HoraTurno
        End With
            mbooEsNuevo = False
    Else
'        Err.Raise ERR_NOHAYCODIGO, "clsSeccion", MSGERR_NOHAYCODIGO
    End If
    Set rs = Nothing
End Sub






Public Property Get idProveedor() As Long
    idProveedor = mProps.idProveedor
End Property

Public Property Let idProveedor(ByVal value As Long)
        mProps.idProveedor = value
      mbooModificado = True
End Property

