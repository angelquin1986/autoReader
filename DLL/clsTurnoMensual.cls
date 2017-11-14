VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTurnoMensual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public mbooEsNuevo As Boolean
Public mbooModificado As Boolean

Private Type T_Props
        IdTurnoM As Long
        mes As String
        Anio As String
        IdSeccion As Long
        CodSeccion As String
        bandValida As Boolean
        FechaDesde As Date
        FechaHasta As Date
        
        FechaGrabado As Date
        BandPorHora As Boolean
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

Public Function grabar()
    Dim rs As Recordset
    Dim sql, fecha As String
    Dim idG As String
    Dim cond As String
    idG = mProps.objEmpresa.GNOpcion.ObtenerValor("Horario") + 1
      sql = "Select * from RTurnoMensual "
     If mbooModificado = False Then Exit Function
      If mbooEsNuevo = True Then
          cond = ""
       Else
          cond = " where idturnoM = " & mProps.IdTurnoM
      End If
      Set rs = mProps.objEmpresa.OpenRecordsetParaEdit(sql & cond)
      If mbooEsNuevo = True Then rs.AddNew
      With rs
          '  rs!idturno = mProps.idturno
            rs!mes = mProps.mes
            rs!Anio = mProps.Anio
            Select Case idG
                Case 1: rs!IdGrupo1 = mProps.IdSeccion
                Case 2: rs!IdGrupo2 = mProps.IdSeccion
                Case 3: rs!IdGrupo3 = mProps.IdSeccion
                Case 4: rs!IdGrupo4 = mProps.IdSeccion
           End Select
            rs!BandPorHora = mProps.BandPorHora
            rs!bandValida = mProps.bandValida
            rs!FechaDesde = mProps.FechaDesde
            rs!FechaHasta = mProps.FechaHasta
            rs!FechaGrabado = mProps.FechaGrabado
            rs.Update
        End With
        
      If mbooEsNuevo = True Then
            sql = "Select IdTurnoM from RTurnoMensual order by IdTurnoM "
            Set rs = objEmpresa.OpenRecordsetParaEdit(sql)
            rs.MoveLast
            grabar = rs.Fields(0)
      Else
            grabar = mProps.IdTurnoM
      End If
        Set rs = Nothing
        mProps.objEmpresa.RaiseRefreshEvent REFRESH_TURNOS
End Function


Public Function recuperar(ByVal key As Variant) As Recordset
Dim idPG As Long

    Dim rs As Recordset
    Dim sql As String
    Dim i As Integer
    idPG = objEmpresa.GNOpcion.ObtenerValor("Horario") + 1
    
        'sql = "SELECT IdTurnom, mes, anio,"
        
        sql = "SELECT IdTurnom, FechaDesde, FechaHasta,"
        Select Case idPG
            Case 1:  sql = sql & " RTurnoMensual.IdGrupo1 , pcg.CodGrupo1 as codseccion"
            Case 2:  sql = sql & " RTurnoMensual.IdGrupo2 , pcg.CodGrupo2 as codseccion"
            Case 3:  sql = sql & " RTurnoMensual.IdGrupo3 , pcg.CodGrupo3 as codseccion"
            Case 4:  sql = sql & " RTurnoMensual.IdGrupo4 , pcg.CodGrupo4 as codseccion"
        End Select
        sql = sql & " ,bandPorHora "
        sql = sql & " FROM RTurnoMensual INNER JOIN pcgrupo" & idPG & " pcg ON RTurnoMensual.Idgrupo" & idPG & " = pcg.idgrupo" & idPG
        sql = sql & " WHERE RTurnoMensual.IdTurnom=" & key
    Set rs = objEmpresa.OpenRecordsetParaEdit(sql)
    If rs.RecordCount > 0 Then
        With rs
                    If Not IsNull(!IdTurnoM) Then mProps.IdTurnoM = !IdTurnoM
'                    If Not IsNull(!Anio) Then mProps.Anio = !Anio
 '                   If Not IsNull(!mes) Then mProps.mes = !mes
                    Select Case idPG
                        Case 1: If Not IsNull(!IdGrupo1) Then mProps.IdSeccion = !IdGrupo1
                        Case 2: If Not IsNull(!IdGrupo2) Then mProps.IdSeccion = !IdGrupo2
                        Case 3: If Not IsNull(!IdGrupo3) Then mProps.IdSeccion = !IdGrupo3
                        Case 4: If Not IsNull(!IdGrupo4) Then mProps.IdSeccion = !IdGrupo4
                    End Select
                    If Not IsNull(!CodSeccion) Then mProps.CodSeccion = !CodSeccion
                    If Not IsNull(!FechaDesde) Then mProps.FechaDesde = !FechaDesde
                    If Not IsNull(!FechaHasta) Then mProps.FechaHasta = !FechaHasta
                    If Not IsNull(!BandPorHora) Then mProps.BandPorHora = !BandPorHora
        End With
            mbooEsNuevo = False
      Set rs = Nothing
    Else
'        Err.Raise ERR_NOHAYCODIGO, "clsTurno", MSGERR_NOHAYCODIGO
        mbooEsNuevo = True
    End If
End Function


Public Sub Class_Initialize()
            mbooModificado = False
            mbooEsNuevo = True
End Sub

Public Property Get mes() As String
    mes = Trim$(mProps.mes)
End Property

Public Property Let mes(ByVal Value As String)
      mProps.mes = Value
      mbooModificado = True
End Property

Public Property Get Anio() As String
    Anio = Trim$(mProps.Anio)
End Property

Public Property Let Anio(ByVal Value As String)
      mProps.Anio = Value
      mbooModificado = True
End Property

Public Property Get CodSeccion() As String
    CodSeccion = mProps.CodSeccion
End Property

Public Property Let CodigoSeccion(ByVal Value As String)
    Dim sql As String, rs As Recordset
    Dim idPG As Integer
    idPG = objEmpresa.GNOpcion.ObtenerValor("Horario+1")
    'Cuando cambia el valor
    If Value <> mProps.CodSeccion Then
        If Len(Value) > 0 Then
            'Actualiza Idseccion también
            Select Case idPG
                Case 1:  sql = "SELECT IdGrupo1 FROM pcgrupo1 WHERE codgrupo1='" & Value & "'"
                Case 2:  sql = "SELECT IdGrupo2 FROM pcgrupo2 WHERE codgrupo2='" & Value & "'"
                Case 3:  sql = "SELECT IdGrupo3 FROM pcgrupo3 WHERE codgrupo3='" & Value & "'"
                Case 4:  sql = "SELECT IdGrupo4 FROM pcgrupo4 WHERE codgrupo4='" & Value & "'"
            End Select
            
            Set rs = objEmpresa.OpenRecordset(sql)
            If rs.EOF Then
                Err.Raise ERR_NOHAYCODIGO, "grupo_Codsecciom", MSGERR_NOHAYCODIGO
                Exit Property
            Else
                Select Case idPG
                    Case 1: IdSeccion = rs!IdGrupo1
                    Case 2: IdSeccion = rs!IdGrupo2
                    Case 3: IdSeccion = rs!IdGrupo3
                    Case 4: IdSeccion = rs!IdGrupo4
                End Select
            End If
            rs.Close
        Else
            IdSeccion = 0
        End If
    End If

    Set rs = Nothing


        mProps.CodSeccion = Value
        mbooModificado = True
End Property

Public Property Let IdSeccion(ByVal Value As Long)
    mProps.IdSeccion = Value
    mbooModificado = True
End Property

Public Property Get IdSeccion() As Long
    IdSeccion = mProps.IdSeccion
End Property

Public Sub BorradoFisico(cod As Long)
    Dim sql As String
    
    sql = "delete from RTurnoMensual where idTurnom = " & cod
    objEmpresa.EjecutarSQL sql, 1
    mProps.objEmpresa.RaiseRefreshEvent REFRESH_TURNOS
End Sub
'AUC por hora
Public Property Get BandPorHora() As Boolean
    BandPorHora = mProps.BandPorHora
End Property

Public Property Let BandPorHora(ByVal Value As Boolean)
      mProps.BandPorHora = Value
      mbooModificado = True
End Property

Public Property Get FechaDesde() As Date
    FechaDesde = mProps.FechaDesde
End Property

Public Property Let FechaDesde(ByVal Value As Date)
      mProps.FechaDesde = Value
      mbooModificado = True
End Property
Public Property Get FechaHasta() As Date
    FechaHasta = mProps.FechaHasta
End Property

Public Property Let FechaHasta(ByVal Value As Date)
      mProps.FechaHasta = Value
      mbooModificado = True
End Property

