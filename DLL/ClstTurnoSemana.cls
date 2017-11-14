VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTurnoSemana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public mbooEsNuevo As Boolean
Public mbooModificado As Boolean

Private Type T_PROPS
'      IdTurno As Long                 'tabla turno
'      Codigo As String                 'tabla turno
'      Descripcion As String        'tabla turno
'      Tipo As String                     'tabla turno
      IdTurno As Long
      CodigoTurno As String
      dia As Long
      entrada As Date
      salida As Date
      receso As Boolean
      InicioReceso As Date
      FinReceso As Date
      h25 As Date
      h100 As Date
      Treceso As Date ' AUC tiempo del receso
      BandValida As Boolean
      objEmpresa As Empresa
End Type
Private mProps As T_PROPS

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

Public Property Let IdTurno(ByVal value As Long)
      mProps.IdTurno = value
      mbooModificado = True
End Property

Public Property Get dia() As Long
    dia = mProps.dia
End Property

Public Property Let dia(ByVal value As Long)
    mProps.dia = value
    mbooModificado = True
End Property

Public Property Get entrada() As Date
    entrada = mProps.entrada
End Property

Public Property Let entrada(ByVal value As Date)
        mProps.entrada = value
        mbooModificado = True
End Property

Public Property Get salida() As Date
    salida = mProps.salida
End Property

Public Property Let salida(ByVal value As Date)
        mProps.salida = value
        mbooModificado = True
End Property

Public Property Get h25() As Date
    h25 = mProps.h25
End Property

Public Property Let h25(ByVal value As Date)
        mProps.h25 = value
        mbooModificado = True
End Property

Public Property Get h100() As Date
    h100 = mProps.h100
End Property

Public Property Let h100(ByVal value As Date)
        mProps.h100 = value
        mbooModificado = True
End Property

Public Property Get receso() As Boolean
    receso = mProps.receso
End Property

Public Property Let receso(ByVal value As Boolean)
        mProps.receso = value
        mbooModificado = True
End Property

Public Property Get InicioReceso() As Date
    InicioReceso = mProps.InicioReceso
End Property

Public Property Let InicioReceso(ByVal value As Date)
        mProps.InicioReceso = value
        mbooModificado = True
End Property

Public Property Get FinReceso() As Date
    FinReceso = mProps.FinReceso
End Property

Public Property Let FinReceso(ByVal value As Date)
        mProps.FinReceso = value
        mbooModificado = True
End Property

Public Sub Grabar()
    Dim rs As Recordset
    Dim sql, fecha As String
    Dim cond As String 'h25,h100,
      'sql = "Select idTurno,dia,entrada,salida,Treceso,inicioreceso,finreceso,treceso from RTurnosemanal"
      sql = "Select * from RTurnosemanal"
     If mbooModificado = False Then Exit Sub
      If mbooEsNuevo = True Then
          cond = ""
       Else
          cond = " where idturno = " & mProps.IdTurno & "and dia= " & mProps.dia
      End If
      Set rs = objEmpresa.OpenRecordsetParaEdit(sql & cond)
      If mbooEsNuevo = True Then rs.AddNew
      With rs
            rs!IdTurno = mProps.IdTurno
            rs!dia = mProps.dia
            fecha = HoraHMS(mProps.entrada, True, True)
            rs!entrada = CDate(fecha)
            fecha = HoraHMS(mProps.salida, True, True)
            rs!salida = CDate(fecha)
            'Fecha = HoraHMS(mProps.h25, True, True)
            'rs!h25 = CDate(Fecha)
'            fecha = HoraHMS(mProps.h50, True, True)
'            rs!h50 = CDate(fecha)
            'Fecha = HoraHMS(mProps.h100, True, True)
            'rs!h100 = CDate(Fecha)
            fecha = HoraHMS(mProps.receso, True, True)
            rs!Treceso = mProps.receso
            rs!BandValida = mProps.BandValida
            If mProps.receso = True Then
                  fecha = HoraHMS(mProps.InicioReceso, True, True)
                  rs!InicioReceso = CDate(fecha)
                  fecha = HoraHMS(mProps.FinReceso, True, True)
                  rs!FinReceso = CDate(fecha)
            Else
                  fecha = HoraHMS(0, True, True)
                  rs!InicioReceso = CDate(fecha)
                  fecha = HoraHMS(0, True, True)
                  rs!FinReceso = CDate(fecha)
            End If
            rs!Treceso = mProps.Treceso 'AUC para colocar el tiempo de receso
            rs.Update
        End With
        Set rs = Nothing
End Sub



Public Sub BorradoFisico(cod As Long)
    Dim sql As String
    sql = "delete from RTurnosemanal where IdTurno = " & cod
    objEmpresa.EjecutarSQL sql, 1
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_TURNOS
End Sub

Public Function Recuperar(ByVal key As Long) As Recordset
    Dim rs As Recordset
    Dim sql As String
    Dim i As Integer
    
    sql = "SELECT ts.IdTurno, ts.dia, ts.Entrada," & _
            " ts.salida,''as h25, '' as h100,ts.InicioReceso,ts.FinReceso, ts.treceso " & _
            " FROM RTurno t INNER JOIN RTurnoSemanal ts  ON T.IdTurno = ts.IdTurno  " & _
            " WHERE t.idTurno=" & key
            
    Set rs = objEmpresa.OpenRecordsetParaEdit(sql)
    If rs.RecordCount > 0 Then
          Set Recuperar = rs
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

Public Function recuperarImp(ByVal key As Variant, ByVal key1 As Variant) As Recordset
    Dim rs As Recordset
    Dim sql As String
    Dim i As Integer
    sql = "SELECT IdTurno,dia,entrada,salida, receso,h100,h25,inicioreceso,finreceso,treceso FROM TurnoSemanal WHERE IdTurno = " & key & "and dia = " & key1 & " order by Dia"
    Set rs = objEmpresa.OpenRecordsetParaEdit(sql)
    If rs.RecordCount > 0 Then
        With rs
            If Not IsNull(!IdTurno) Then mProps.IdTurno = !IdTurno
            If Not IsNull(!dia) Then mProps.dia = !dia
            mProps.entrada = !entrada
            If Not IsNull(!salida) Then mProps.salida = !salida
            If Not IsNull(!receso) Then mProps.receso = !receso
            If Not IsNull(!InicioReceso) Then mProps.InicioReceso = !InicioReceso
            If Not IsNull(!FinReceso) Then mProps.FinReceso = !FinReceso
            If Not IsNull(!h100) Then mProps.h100 = !h100
            If Not IsNull(!h25) Then mProps.h25 = !h25
            If Not IsNull(!Treceso) Then mProps.Treceso = !Treceso
        End With
            mbooEsNuevo = False
    Else
        'Err.Raise ERR_NOHAYCODIGO, "clsSeccion", MSGERR_NOHAYCODIGO
        mbooEsNuevo = True
    End If
End Function

Public Property Get CodigoTurno() As String
    CodigoTurno = mProps.CodigoTurno
End Property

Public Property Let CodigoTurno(ByVal value As String)
    Dim sql As String, rs As Recordset
    
    'Cuando cambia el valor
    If value <> mProps.CodigoTurno Then
        If Len(value) > 0 Then
            'Actualiza IdTurno también
            sql = "SELECT IdTurno FROM Turno WHERE codigo='" & value & "'"
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

Public Function recuperarImportar(ByVal key As Variant, ByVal key1 As Integer) As Recordset
    Dim rs As Recordset
    Dim sql As String
    Dim i As Integer
'    sql = "SELECT idturno,dia,iif(entrada=0,'',entrada),iif(salida=0,'',salida),iif(h100=0,'',h100),iif(h25=0,'',h25),iif(inicioreceso=0,'',inicioreceso),iif(finreceso=0,'',finreceso) FROM TurnoSemanal WHERE idTurno = " & key & " order by Dia"
    sql = "SELECT TurnoSemanal.IdTurno, TurnoSemanal.dia, " & _
            " IIF(TurnoSemanal.Entrada=0,'',TurnoSemanal.Entrada), " & _
            " IIF(TurnoSemanal.salida=0,'',TurnoSemanal.salida), " & _
            " IIF(TurnoSemanal.h100=0,'',TurnoSemanal.h100), " & _
            " IIF(TurnoSemanal.h25=0,'',TurnoSemanal.h25), " & _
            " IIF(TurnoSemanal.InicioReceso=0,'',TurnoSemanal.InicioReceso)," & _
            " IIF(TurnoSemanal.FinReceso=0,'',TurnoSemanal.FinReceso), " & _
            " IIF(TurnoSemanal.treceso=0,'',TurnoSemanal.treceso) " & _
           " FROM Turno INNER JOIN TurnoSemanal " & _
            " ON Turno.IdTurno = TurnoSemanal.IdTurno " & _
            " WHERE (((Turno.Codigo)='" & key & "') AND TurnoSemanal.dia=" & key1 & ")"

    Set rs = objEmpresa.OpenRecordsetParaEdit(sql)
    If rs.RecordCount > 0 Then
          Set recuperarImportar = rs
            mbooEsNuevo = False
    Else
'        Err.Raise ERR_NOHAYCODIGO, "clsTurno", MSGERR_NOHAYCODIGO
        mbooEsNuevo = True
    End If
End Function
'AUC agregado 11/2011
Public Property Let Treceso(ByVal value As Date)
        mProps.Treceso = value
        mbooModificado = True
End Property

Public Property Get Treceso() As Date
    Treceso = mProps.Treceso
End Property

Public Property Get BandValida() As Boolean
    BandValida = mProps.BandValida
End Property

Public Property Let BandValida(ByVal value As Boolean)
        mProps.BandValida = value
        mbooModificado = True
End Property

